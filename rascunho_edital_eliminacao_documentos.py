# importações de bibliotecas
import io
import os
import re
import sys
from datetime import datetime

import pandas as pd
from docxtpl import DocxTemplate
from openpyxl import load_workbook

# ============================================================
# CONSTANTES DE CAMINHO E MEDIDAS
# ============================================================
# pasta do SharePoint (DETRAN - Divisão de Gestão Documental) sincronizada localmente
# na pasta do usuário logado na máquina (ex.: C:\Users\<usuario>\...)
PASTA = os.path.join(
    os.path.expanduser("~"),
    "PRODESP",
    "DETRAN - DIVISÃO DE GESTÃO DOCUMENTAL - Documentos",
    "Editais de Eliminação de Documentos"
)
ARQUIVO = os.path.join(PASTA, "Relacao de Expurgo para Rascunho.xlsx")
EDITAIS = os.path.join(PASTA, "Editais Elaborados")
MODELO = os.path.join(PASTA, "Modelos")

MODELO_EDITAL = os.path.join(MODELO, "modelo_edital.docx")  # CAIXA
MODELO_EDITAL_MASSA = os.path.join(MODELO, "modelo_edital_massa.docx")  # MASSA

METRAGEM_MEDIDA = 0.14  # CAIXA: metros lineares por caixa
METROS_LINEARES_POR_METRO_CUBICO = 12  # MASSA: conversão de m³ para metros lineares


# ============================================================
# FUNÇÕES AUXILIARES (comuns a CAIXA e MASSA)
# ============================================================
_UNIDADES = [
    '', 'uma', 'duas', 'três', 'quatro',
    'cinco', 'seis', 'sete', 'oito', 'nove'
]

_DEZ_A_DEZENOVE = [
    'dez', 'onze', 'doze', 'treze', 'catorze',
    'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove'
]

_DEZENAS = [
    '', '', 'vinte', 'trinta', 'quarenta',
    'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa'
]

_CENTENAS = [
    '', 'cento', 'duzentas', 'trezentas', 'quatrocentas',
    'quinhentas', 'seiscentas', 'setecentas', 'oitocentas', 'novecentas'
]

_MESES = [
    'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
    'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
]


def _grupo_ate_999_extenso(numero):
    if numero == 0:
        return ''

    if numero == 100:
        return 'cem'

    centena, resto = divmod(numero, 100)
    dezena, unidade = divmod(resto, 10)

    partes = []

    if centena:
        partes.append(_CENTENAS[centena])

    if 10 <= resto <= 19:
        partes.append(_DEZ_A_DEZENOVE[resto - 10])
    else:
        if dezena:
            partes.append(_DEZENAS[dezena])
        if unidade:
            partes.append(_UNIDADES[unidade])

    return ' e '.join(partes)


def numero_por_extenso(numero):
    # CAIXA: sempre no feminino, usado só para quantidade de caixas
    numero = int(numero)

    if not 0 <= numero <= 999_999:
        raise ValueError(
            f"Número fora do intervalo suportado (0 a 999.999): {numero}"
        )

    if numero == 0:
        return 'zero'

    milhar, resto = divmod(numero, 1000)

    if not milhar:
        return _grupo_ate_999_extenso(resto)

    texto_milhar = (
        'mil' if milhar == 1
        else f"{_grupo_ate_999_extenso(milhar)} mil"
    )

    if not resto:
        return texto_milhar

    conector = ' e ' if (resto < 100 or resto % 100 == 0) else ', '

    return f"{texto_milhar}{conector}{_grupo_ate_999_extenso(resto)}"


def limpar_nome(nome):
    # remove caracteres que o Windows não aceita em nome de arquivo
    return re.sub(r'[\\/*?:"<>|]', "_", str(nome))


def limpar_texto(valor):
    """
    Converte a célula para texto e remove o artefato '_x000D_' que o Excel
    grava quando uma célula tem quebra de linha manual (Alt+Enter) e o
    openpyxl não converte de volta para uma quebra de linha de verdade.
    Usado em toda coluna de texto lida da planilha (CAIXA e MASSA), então
    o restante do código nunca precisa se preocupar com esse artefato.
    """
    if pd.isna(valor):
        return ''
    return str(valor).replace('_x000D_', ' ').strip()


def capitalizar_personalizado(texto):
    # CAIXA: capitaliza cada palavra, exceto preposições/artigos no meio da frase
    conectores = {
        'de', 'do', 'da', 'dos', 'das', 'e',
        'em', 'com', 'no', 'na', 'nos', 'nas',
        'a', 'o', 'as', 'os'
    }

    palavras = texto.lower().split()
    resultado = []

    for i, palavra in enumerate(palavras):
        if i == 0 or palavra not in conectores:
            resultado.append(palavra.capitalize())
        else:
            resultado.append(palavra)

    return ' '.join(resultado)


def extrair_chave_ordenacao(codigo):
    # CAIXA: transforma "1.10.2" em [1, 10, 2] para ordenar a Série documental numericamente
    return [int(p) for p in str(codigo).split('.') if p.isdigit()]


def data_por_extenso(data):
    # não depende do locale do sistema operacional (usado por CAIXA e MASSA)
    return f"{data.day:02d} de {_MESES[data.month - 1]} de {data.year}"


def formatar_numero_br(valor):
    # MASSA: formata no padrão brasileiro (vírgula decimal, sem zeros à direita)
    texto = f"{float(valor):.2f}"
    if '.' in texto:
        texto = texto.rstrip('0').rstrip('.')
    return texto.replace('.', ',')


def buscar_regiao_administrativa(municipio, mapa_regiao):
    # usado por CAIXA e MASSA: região administrativa do cabeçalho do edital
    regiao = mapa_regiao.get(str(municipio).strip().lower())

    if not regiao:
        raise ValueError(
            f"Município '{municipio}' não encontrado na aba 'Municípios'."
        )

    return regiao


# ============================================================
# LEITURA DA PLANILHA (comum: aba "Municípios" e aba "Membros CADA")
# ============================================================
def carregar_mapa_regiao(arquivo):
    """Lê a aba 'Municípios' e monta o dicionário {município: região administrativa}, usado tanto por CAIXA quanto por MASSA."""
    municipios_df = pd.read_excel(
        arquivo,
        sheet_name="Municípios",
        engine='openpyxl'
    )

    municipios_df['Município'] = municipios_df['Município'].apply(limpar_texto)
    municipios_df['Região Administrativa'] = (
        municipios_df['Região Administrativa'].apply(limpar_texto)
    )

    return {
        municipio.lower(): regiao
        for municipio, regiao in zip(
            municipios_df['Município'],
            municipios_df['Região Administrativa']
        )
    }


def carregar_membro_assinante(arquivo):
    """
    Lê a aba 'Membros CADA' e devolve (nome, cargo) de quem assina o
    edital: o membro com STATUS 'Ativo'. Se ninguém estiver ativo, assina
    a coordenadora padrão. Usado tanto por CAIXA quanto por MASSA.
    """
    membros = pd.read_excel(
        arquivo,
        sheet_name="Membros CADA",
        engine='openpyxl'
    )

    membros['STATUS'] = membros['STATUS'].apply(limpar_texto)

    membros_ativos = membros[membros['STATUS'].str.lower() == 'ativo']

    if len(membros_ativos) > 1:
        raise ValueError(
            "Esperado no máximo 1 membro com STATUS 'Ativo' na aba "
            f"'Membros CADA', encontrado(s): {len(membros_ativos)}"
        )

    if membros_ativos.empty:
        # nenhum membro ativo: assina a coordenadora padrão
        return "IARA LOPES DA SILVA", "Coordenadora"

    nome_membro = str(membros_ativos['NOME'].iloc[0]).strip().upper()
    cargo_membro = str(membros_ativos['CARGO'].iloc[0]).strip()

    return nome_membro, cargo_membro


# ============================================================
# EDITAL DE CAIXA
# ============================================================
def carregar_editais_caixa(arquivo):
    """
    Lê a aba 'Edital de Caixa', limpa os campos de texto (removendo o
    artefato _x000D_ e espaços nas pontas) e devolve só as linhas
    marcadas com Status Edital = 'Criar edital'.
    """
    dataframe = pd.read_excel(
        arquivo,
        sheet_name="Edital de Caixa",
        engine='openpyxl'
    )

    colunas_texto = [
        'Função',
        'Subfunção',
        'Atividade',
        'Série documental',
        'Descrição documental',
        'Observações complementares',
        'Município'
    ]

    for campo in colunas_texto:
        dataframe[campo] = dataframe[campo].apply(limpar_texto)

    return dataframe[
        dataframe['Status Edital'].str.contains(
            "Criar edital",
            case=False,
            na=False
        )
    ]


def montar_itens_detalhamento_caixa(grupo):
    """
    A partir das linhas de um grupo (mesmo N° Processo SEI + Município),
    ordena os itens pela Série documental e monta a lista de itens do
    detalhamento (um edital de caixa lista vários itens/documentos).
    """
    grupo = grupo.copy()
    grupo['chave_ordenacao'] = grupo['Série documental'].apply(extrair_chave_ordenacao)

    # agrupa linhas duplicadas do mesmo item (mesma classificação completa)
    detalhamento = grupo.groupby([
        'Função',
        'Subfunção',
        'Atividade',
        'Série documental',
        'Descrição documental',
        'Data Limite',
        'Quantidade',
        'Observações complementares'
    ], dropna=False).first().reset_index()

    detalhamento['chave_ordenacao'] = (
        detalhamento['Série documental'].apply(extrair_chave_ordenacao)
    )
    detalhamento = detalhamento.sort_values(by='chave_ordenacao')

    itens_detalhamento = []

    for _, row in detalhamento.iterrows():
        qtde_caixas = int(row['Quantidade'])
        caixas_extenso = numero_por_extenso(qtde_caixas)
        caixas_extenso = (
            f"({caixas_extenso}) Caixa" if qtde_caixas == 1
            else f"({caixas_extenso}) Caixas"
        )

        itens_detalhamento.append({
            "funcao": capitalizar_personalizado(row['Função']),
            "subfuncao": row['Subfunção'].capitalize(),
            "atividade": row['Atividade'],
            "serie_documental": row['Série documental'],
            "descricao_documental": row['Descrição documental'],
            "data_limite": limpar_texto(row['Data Limite']),
            "qtde_caixas": str(qtde_caixas).zfill(2),
            "caixas_extenso": caixas_extenso,
            "observacoes_complementares": row['Observações complementares']
        })

    return itens_detalhamento


def gerar_edital_caixa(
    processo, municipio, grupo, mapa_regiao,
    nome_membro, cargo_membro, modelo_edital_bytes, aba_caixa
):
    """
    Gera o .docx de um único edital de caixa (um grupo de N° Processo SEI
    + Município, com um item para cada linha da planilha) e marca as
    linhas correspondentes como 'Edital criado' na planilha, em memória
    (o arquivo Excel é salvo uma única vez, no final do script).
    """
    data_edital = data_por_extenso(datetime.now()).upper()
    regiao = buscar_regiao_administrativa(municipio, mapa_regiao)
    itens_detalhamento = montar_itens_detalhamento_caixa(grupo)

    # rodapé: total de caixas e conversão para metros lineares
    total_caixas = int(grupo['Quantidade'].sum())
    total_caixas_extenso = numero_por_extenso(total_caixas)
    total_caixas_extenso = (
        f"({total_caixas_extenso}) Caixa" if total_caixas == 1
        else f"({total_caixas_extenso}) Caixas"
    )
    total_metros_lineares = total_caixas * METRAGEM_MEDIDA

    contexto_edital = {
        "data_edital": data_edital,
        "regiao": regiao,
        "municipio": municipio,
        "itens": itens_detalhamento,
        "total_caixas": str(total_caixas),
        "total_caixas_extenso": total_caixas_extenso,
        "total_metros_lineares": f"{total_metros_lineares:.2f}",
        "nome_membro": nome_membro,
        "cargo_membro": cargo_membro
    }

    nome_arquivo = (
        f"Edital_{limpar_nome(municipio)}_"
        f"{limpar_nome(processo)}.docx"
    )
    caminho_arquivo = os.path.join(EDITAIS, nome_arquivo)

    documento = DocxTemplate(io.BytesIO(modelo_edital_bytes))
    documento.render(contexto_edital)
    documento.save(caminho_arquivo)

    for idx in grupo.index:
        excel_row = idx + 2
        aba_caixa[f"O{excel_row}"] = "Edital criado"


def gerar_editais_caixa(
    df_criar_edital, mapa_regiao,
    nome_membro, cargo_membro, modelo_edital_bytes, aba_caixa
):
    """Percorre todos os grupos (N° Processo SEI + Município) da aba de caixa e gera um edital para cada um."""
    solicitacao = df_criar_edital.groupby(
        ['N° Processo SEI', 'Município'],
        dropna=False
    )

    for (processo, municipio), grupo in solicitacao:
        gerar_edital_caixa(
            processo, municipio, grupo, mapa_regiao,
            nome_membro, cargo_membro, modelo_edital_bytes, aba_caixa
        )


# ============================================================
# EDITAL DE MASSA
# ============================================================
def carregar_editais_massa(arquivo):
    """
    Lê a aba 'Edital de Massa', limpa os campos de texto e devolve
    (linhas marcadas com Status Edital = 'Criar edital', nome da coluna
    do Processo SEI usada nessa aba — aceita tanto 'N°' quanto 'Nº').
    """
    dataframe_massa = pd.read_excel(
        arquivo,
        sheet_name="Edital de Massa",
        engine='openpyxl'
    )
    dataframe_massa.columns = [
        str(coluna).strip() for coluna in dataframe_massa.columns
    ]

    coluna_processo_massa = (
        "N° Processo SEI" if "N° Processo SEI" in dataframe_massa.columns
        else "Nº Processo SEI"
    )

    for campo in ['Município', 'Observações complementares']:
        dataframe_massa[campo] = dataframe_massa[campo].apply(limpar_texto)

    df_criar_edital_massa = dataframe_massa[
        dataframe_massa['Status Edital'].str.contains(
            "Criar edital",
            case=False,
            na=False
        )
    ]

    return df_criar_edital_massa, coluna_processo_massa


def gerar_edital_massa(
    idx, row, coluna_processo_massa, mapa_regiao,
    nome_membro, cargo_membro, modelo_edital_massa_bytes, aba_massa
):
    """
    Gera o .docx de um único edital de massa (cada linha da planilha é
    um edital, sem agrupamento) e marca a linha como 'Edital criado' na
    planilha, em memória (o arquivo Excel é salvo uma única vez, no
    final do script).
    """
    data_edital = data_por_extenso(datetime.now()).upper()

    processo = row[coluna_processo_massa]
    municipio = str(row['Município']).strip()
    regiao = buscar_regiao_administrativa(municipio, mapa_regiao)

    # volume em metros cúbicos e conversão para metros lineares
    comprimento = float(row['Comprimento'])
    largura = float(row['Largura'])
    altura = float(row['Altura'])
    metros_cubicos = comprimento * largura * altura
    total_metros_lineares = metros_cubicos * METROS_LINEARES_POR_METRO_CUBICO

    contexto_edital = {
        "data_edital": data_edital,
        "regiao": regiao,
        "municipio": municipio,
        "altura": formatar_numero_br(altura),
        "comprimento": formatar_numero_br(comprimento),
        "largura": formatar_numero_br(largura),
        "metros_cubicos": formatar_numero_br(metros_cubicos),
        "total_metros_lineares": formatar_numero_br(total_metros_lineares),
        "observacoes_complementares": row['Observações complementares'],
        "nome_membro": nome_membro,
        "cargo_membro": cargo_membro
    }

    excel_row = idx + 2
    nome_arquivo = (
        f"Edital_Massa_{limpar_nome(municipio)}_"
        f"{limpar_nome(processo)}_L{excel_row}.docx"
    )
    caminho_arquivo = os.path.join(EDITAIS, nome_arquivo)

    documento = DocxTemplate(io.BytesIO(modelo_edital_massa_bytes))
    documento.render(contexto_edital)
    documento.save(caminho_arquivo)

    aba_massa[f"H{excel_row}"] = "Edital criado"


def gerar_editais_massa(
    df_criar_edital_massa, coluna_processo_massa, mapa_regiao,
    nome_membro, cargo_membro, modelo_edital_massa_bytes, aba_massa
):
    """Percorre cada linha da aba de massa e gera um edital por linha (sem agrupamento)."""
    for idx, row in df_criar_edital_massa.iterrows():
        gerar_edital_massa(
            idx, row, coluna_processo_massa, mapa_regiao,
            nome_membro, cargo_membro, modelo_edital_massa_bytes, aba_massa
        )


# ============================================================
# ORQUESTRAÇÃO (CAIXA + MASSA)
# ============================================================
def main():
    try:
        # CAIXA: linhas marcadas para gerar edital de documentos catalogados
        df_criar_edital = carregar_editais_caixa(ARQUIVO)

        # MASSA: linhas marcadas para gerar edital em bloco (ex.: sinistro)
        df_criar_edital_massa, coluna_processo_massa = carregar_editais_massa(ARQUIVO)

        if df_criar_edital.empty and df_criar_edital_massa.empty:
            print("Nenhum edital a ser criado.")
            sys.exit()

        # comum a CAIXA e MASSA: região administrativa e assinante do edital
        mapa_regiao = carregar_mapa_regiao(ARQUIVO)
        nome_membro, cargo_membro = carregar_membro_assinante(ARQUIVO)

        # Excel e templates carregados uma única vez, fora dos loops de geração
        wb = load_workbook(ARQUIVO)
        aba_caixa = wb["Edital de Caixa"]
        aba_massa = wb["Edital de Massa"]

        # CAIXA: gera um edital por grupo (N° Processo SEI + Município)
        if not df_criar_edital.empty:
            with open(MODELO_EDITAL, 'rb') as arquivo_modelo:
                modelo_edital_bytes = arquivo_modelo.read()

            gerar_editais_caixa(
                df_criar_edital, mapa_regiao,
                nome_membro, cargo_membro, modelo_edital_bytes, aba_caixa
            )

        # MASSA: gera um edital por linha da planilha
        if not df_criar_edital_massa.empty:
            with open(MODELO_EDITAL_MASSA, 'rb') as arquivo_modelo:
                modelo_edital_massa_bytes = arquivo_modelo.read()

            gerar_editais_massa(
                df_criar_edital_massa, coluna_processo_massa, mapa_regiao,
                nome_membro, cargo_membro, modelo_edital_massa_bytes, aba_massa
            )

        # planilha salva uma única vez, com todas as marcações de status (CAIXA e MASSA)
        wb.save(ARQUIVO)

        print("Editais criados com sucesso!")

    except FileNotFoundError as erro:
        sys.exit(
            f"Arquivo não encontrado: {erro.filename}\n"
            "Verifique se a pasta do SharePoint está sincronizada e se a "
            "planilha e os templates estão no lugar esperado."
        )
    except PermissionError as erro:
        sys.exit(
            f"Não foi possível acessar '{erro.filename}'.\n"
            "Verifique se o arquivo não está aberto no Excel ou no Word "
            "e execute o script novamente."
        )
    except KeyError as erro:
        sys.exit(
            f"Coluna não encontrada na planilha: {erro}\n"
            "Verifique se os cabeçalhos das abas não foram alterados."
        )
    except ValueError as erro:
        sys.exit(f"Erro nos dados da planilha: {erro}")


if __name__ == "__main__":
    main()
