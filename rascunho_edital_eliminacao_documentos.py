# importações de bibliotecas
import io
import os
import re
import sys
from datetime import datetime

import pandas as pd
from docxtpl import DocxTemplate
from openpyxl import load_workbook

# constantes de caminho e medidas
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

MODELO_EDITAL = os.path.join(MODELO, "modelo_edital_caixa.docx")
MODELO_EDITAL_MASSA = os.path.join(MODELO, "modelo_edital_massa.docx")

METRAGEM_MEDIDA = 0.14  # metros lineares por caixa
METROS_LINEARES_POR_METRO_CUBICO = 12  # conversão de m³ para metros lineares

# funções auxiliares
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
    # sempre no feminino: usado só para quantidade de caixas
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
    return re.sub(r'[\\/*?:"<>|]', "_", str(nome))


def capitalizar_personalizado(texto):
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
    return [int(p) for p in str(codigo).split('.') if p.isdigit()]


def data_por_extenso(data):
    # não depende do locale do sistema operacional
    return f"{data.day:02d} de {_MESES[data.month - 1]} de {data.year}"


def formatar_numero_br(valor):
    texto = f"{float(valor):.2f}"
    if '.' in texto:
        texto = texto.rstrip('0').rstrip('.')
    return texto.replace('.', ',')


def buscar_regiao_administrativa(municipio, mapa_regiao):
    regiao = mapa_regiao.get(str(municipio).strip().lower())

    if not regiao:
        raise ValueError(
            f"Município '{municipio}' não encontrado na aba 'Municípios'."
        )

    return regiao


# código principal
try:
    dataframe = pd.read_excel(
        ARQUIVO,
        sheet_name="Edital de Caixa",
        engine='openpyxl'
    )

    limpar_colunas = [
        'Função',
        'Subfunção',
        'Atividade',
        'Série documental',
        'Descrição documental',
        'Observações complementares'
    ]

    for campo in limpar_colunas + ['Município']:
        dataframe[campo] = (
            dataframe[campo].fillna('').astype(str).str.strip()
        )

    # filtra apenas as linhas com "Criar edital"
    df_criar_edital = dataframe[
        dataframe['Status Edital'].str.contains(
            "Criar edital",
            case=False,
            na=False
        )
    ]

    # planilha "Edital de Massa"
    dataframe_massa = pd.read_excel(
        ARQUIVO,
        sheet_name="Edital de Massa",
        engine='openpyxl'
    )
    dataframe_massa.columns = [
        str(coluna).strip() for coluna in dataframe_massa.columns
    ]

    COLUNA_PROCESSO_MASSA = (
        "N° Processo SEI" if "N° Processo SEI" in dataframe_massa.columns
        else "Nº Processo SEI"
    )

    for campo in ['Município', 'Observações complementares']:
        dataframe_massa[campo] = (
            dataframe_massa[campo].fillna('').astype(str).str.strip()
        )

    df_criar_edital_massa = dataframe_massa[
        dataframe_massa['Status Edital'].str.contains(
            "Criar edital",
            case=False,
            na=False
        )
    ]

    if df_criar_edital.empty and df_criar_edital_massa.empty:
        print("Nenhum edital a ser criado.")
        sys.exit()

    # município -> região administrativa
    municipios_df = pd.read_excel(
        ARQUIVO,
        sheet_name="Municípios",
        engine='openpyxl'
    )

    municipios_df['Município'] = (
        municipios_df['Município'].fillna('').astype(str).str.strip()
    )
    municipios_df['Região Administrativa'] = (
        municipios_df['Região Administrativa'].fillna('').astype(str).str.strip()
    )

    mapa_regiao = {
        municipio.lower(): regiao
        for municipio, regiao in zip(
            municipios_df['Município'],
            municipios_df['Região Administrativa']
        )
    }

    # membro da Comissão de Avaliação de Documentos e Acesso (CADA) que assina o edital
    membros = pd.read_excel(
        ARQUIVO,
        sheet_name="Membros CADA",
        engine='openpyxl'
    )

    membros['STATUS'] = membros['STATUS'].fillna('').astype(str).str.strip()

    membros_ativos = membros[membros['STATUS'].str.lower() == 'ativo']

    if len(membros_ativos) > 1:
        raise ValueError(
            "Esperado no máximo 1 membro com STATUS 'Ativo' na aba "
            f"'Membros CADA', encontrado(s): {len(membros_ativos)}"
        )

    if membros_ativos.empty:
        # nenhum membro ativo: assina a coordenadora padrão
        nome_membro = "IARA LOPES DA SILVA"
        cargo_membro = "Coordenadora"
    else:
        nome_membro = str(membros_ativos['NOME'].iloc[0]).strip().upper()
        cargo_membro = str(membros_ativos['CARGO'].iloc[0]).strip()

    # templates e planilha carregados uma única vez, fora dos loops
    wb = load_workbook(ARQUIVO)
    aba_caixa = wb["Edital de Caixa"]
    aba_massa = wb["Edital de Massa"]

    if not df_criar_edital.empty:
        with open(MODELO_EDITAL, 'rb') as arquivo_modelo:
            modelo_edital_bytes = arquivo_modelo.read()

    if not df_criar_edital_massa.empty:
        with open(MODELO_EDITAL_MASSA, 'rb') as arquivo_modelo:
            modelo_edital_massa_bytes = arquivo_modelo.read()

    solicitacao = df_criar_edital.groupby(
        ['N° Processo SEI', 'Município'],
        dropna=False
    )

    for (processo, municipio), grupo in solicitacao:

        # cabeçalho
        data_edital = data_por_extenso(datetime.now()).upper()

        regiao = buscar_regiao_administrativa(municipio, mapa_regiao)

        # limpeza das colunas
        for campo in limpar_colunas:
            grupo.loc[:, campo] = (
                grupo[campo]
                .fillna('')
                .astype(str)
                .str.strip()
            )

        # ordenação
        grupo.loc[:, 'chave_ordenacao'] = (
            grupo['Série documental']
            .apply(extrair_chave_ordenacao)
        )

        # agrupamento
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
            detalhamento['Série documental']
            .apply(extrair_chave_ordenacao)
        )

        detalhamento = detalhamento.sort_values(
            by='chave_ordenacao'
        )

        # detalhamento
        itens_detalhamento = []

        for _, row in detalhamento.iterrows():

            qtde_caixas = int(row['Quantidade'])

            caixas_extenso = numero_por_extenso(qtde_caixas)

            if qtde_caixas == 1:
                caixas_extenso = f"({caixas_extenso}) Caixa"
            else:
                caixas_extenso = f"({caixas_extenso}) Caixas"

            itens_detalhamento.append({
                "funcao": capitalizar_personalizado(
                    row['Função']
                ).replace('_x000D_', ''),

                "subfuncao": row['Subfunção']
                .replace('_x000D_', ' ')
                .capitalize(),

                "atividade": row['Atividade']
                .replace('_x000D_', ' '),

                "serie_documental": row['Série documental']
                .replace('_x000D_', ' '),

                "descricao_documental": row['Descrição documental']
                .replace('_x000D_', ''),

                "data_limite": str(row['Data Limite'])
                .replace('_x000D_', ' '),

                "qtde_caixas": str(qtde_caixas).zfill(2),

                "caixas_extenso": caixas_extenso,

                "observacoes_complementares": (
                    row['Observações complementares']
                    .replace('_x000D_', ' ')
                )
            })

        # rodapé
        total_caixas = int(grupo['Quantidade'].sum())

        total_caixas_extenso = numero_por_extenso(total_caixas)

        total_metros_lineares = (
            total_caixas * METRAGEM_MEDIDA
        )

        if total_caixas == 1:
            total_caixas_extenso = (
                f"({total_caixas_extenso}) Caixa"
            )
        else:
            total_caixas_extenso = (
                f"({total_caixas_extenso}) Caixas"
            )

        # conteúdo final
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

        # salva documento
        nome_arquivo = (
            f"Edital_{limpar_nome(municipio)}_"
            f"{limpar_nome(processo)}.docx"
        )

        caminho_arquivo = os.path.join(
            EDITAIS,
            nome_arquivo
        )

        documento = DocxTemplate(io.BytesIO(modelo_edital_bytes))
        documento.render(contexto_edital)
        documento.save(caminho_arquivo)

        # marca status (planilha é salva uma única vez, ao final do script)
        for idx in grupo.index:
            excel_row = idx + 2
            aba_caixa[f"O{excel_row}"] = "Edital criado"

    # edital de massa: cada linha da planilha gera o seu próprio edital
    for idx, row in df_criar_edital_massa.iterrows():

        data_edital = data_por_extenso(datetime.now()).upper()

        processo = row[COLUNA_PROCESSO_MASSA]
        municipio = str(row['Município']).strip()

        regiao = buscar_regiao_administrativa(municipio, mapa_regiao)

        comprimento = float(row['Comprimento'])
        largura = float(row['Largura'])
        altura = float(row['Altura'])

        metros_cubicos = comprimento * largura * altura

        total_metros_lineares = (
            metros_cubicos * METROS_LINEARES_POR_METRO_CUBICO
        )

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

        # salva documento
        excel_row = idx + 2

        nome_arquivo = (
            f"Edital_Massa_{limpar_nome(municipio)}_"
            f"{limpar_nome(processo)}_L{excel_row}.docx"
        )

        caminho_arquivo = os.path.join(
            EDITAIS,
            nome_arquivo
        )

        documento = DocxTemplate(io.BytesIO(modelo_edital_massa_bytes))
        documento.render(contexto_edital)
        documento.save(caminho_arquivo)

        # marca status (planilha é salva uma única vez, ao final do script)
        aba_massa[f"H{excel_row}"] = "Edital criado"

    # planilha salva uma única vez, com todas as marcações de status
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
