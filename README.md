# Editais de Eliminação (Expurgo)

Script em Python para automatizar a elaboração de rascunhos de **editais de eliminação de documentos** (expurgo) a partir de uma planilha de controle.

## Para que serve

A partir de uma planilha Excel com a relação de documentos a serem eliminados, o script gera dois tipos de edital — **Edital de Caixa** (documentos catalogados, eliminados conforme a Tabela de Temporalidade) e **Edital de Massa** (documentos eliminados em bloco, ex.: por sinistro) — cada um com seu próprio template Word:

- Filtra os registros marcados como **"Criar edital"** em cada aba;
- Busca a **Região Administrativa** de cada município na aba "Municípios" (usada tanto pelo edital de caixa quanto pelo de massa);
- Busca na aba "Membros CADA" o membro com **STATUS "Ativo"** e usa o nome e cargo dele na assinatura do edital;
- **Edital de Caixa**: agrupa os itens por **Número de Processo SEI** e **Município**, ordena pela **Série documental** e preenche o template com o detalhamento de cada item, incluindo a conversão de quantidades por extenso;
- **Edital de Massa**: cada linha da planilha gera o seu próprio edital — calcula o **Volume em metros cúbicos** (Comprimento × Largura × Altura) e o converte em **Total de Metros Lineares**;
- Gera um documento **Word (.docx)** para cada edital, na pasta de editais elaborados;
- Atualiza a planilha original, marcando os registros processados como **"Edital criado"**.

## Tecnologias utilizadas

- **Python 3**
- [pandas](https://pandas.pydata.org/) — leitura, filtragem, agrupamento e ordenação dos dados da planilha
- [openpyxl](https://openpyxl.readthedocs.io/) — leitura/escrita do arquivo Excel (engine do pandas e atualização de status)
- [docxtpl](https://docxtpl.readthedocs.io/) — preenchimento do template `.docx` do edital com placeholders no estilo Jinja2 (`{{ variavel }}`), preservando a formatação do Word
- [python-docx](https://python-docx.readthedocs.io/) — biblioteca usada internamente pelo docxtpl para manipular o `.docx`
- Bibliotecas padrão do Python: `os`, `re`, `datetime`

A conversão de quantidades numéricas por extenso (edital de caixa) é feita por uma função própria (`numero_por_extenso`), sem depender de biblioteca externa. Isso porque o [num2words](https://github.com/savoirfairelinux/num2words) não flexiona em gênero para português (`pt_BR`) — ele nem aceita o parâmetro `gender`, e sempre retorna a forma masculina (ex.: "dois", nunca "duas"). Como o script sempre se refere a "caixas" (substantivo feminino), foi implementada uma conversão própria (0 a 999.999) sempre no feminino, validada contra a saída do `num2words` (masculina, com flexão manual) em quase 30 mil números.

A data por extenso do cabeçalho também é montada por uma função própria (`data_por_extenso`), com os nomes dos meses em português fixos no código — em vez de depender de `locale.setlocale`, que usa nomes de locale (`pt_BR.UTF-8`) específicos de Linux/macOS e não funciona nas máquinas Windows do DETRAN.

Já o edital de massa não lida com quantidades por extenso — ele converte volume em metros cúbicos para metros lineares multiplicando por um fator fixo (`METROS_LINEARES_POR_METRO_CUBICO = 12`), e formata os números no padrão brasileiro (vírgula decimal, sem zeros à direita) através da função `formatar_numero_br`.

## Estrutura esperada

O script precisa ser executado com a pasta do SharePoint [Editais de Eliminação de Documentos](https://governosp.sharepoint.com/:f:/r/teams/DETRAN-DIVISODEGESTODOCUMENTAL/Documentos%20Compartilhados/Editais%20de%20Elimina%C3%A7%C3%A3o%20de%20Documentos?d=w159ef8b07b224315b0a4aff6c6058d69&csf=1&web=1&e=jMhxke) (DETRAN - Divisão de Gestão Documental) sincronizada localmente pelo OneDrive. O caminho é montado dinamicamente a partir da pasta do usuário logado na máquina (`C:\Users\<usuario>\PRODESP\DETRAN - DIVISÃO DE GESTÃO DOCUMENTAL - Documentos\Editais de Eliminação de Documentos`), então o script funciona em qualquer computador sem precisar alterar o código. Dentro dela:

- `Relacao de Expurgo para Rascunho.xlsx` — planilha de controle, com quatro abas:
  - `Edital de Caixa` — relação de documentos catalogados a eliminar (Nº Processo SEI, Município, Função, Subfunção, Atividade, Série documental, Descrição documental, Data Limite, Quantidade, Observações complementares, Status Edital). Várias linhas com o mesmo **Nº Processo SEI + Município** viram um único edital, com um item para cada linha
  - `Edital de Massa` — documentos eliminados em bloco (Data do pedido, Nº Processo SEI, Município, Comprimento, Largura, Altura, Observações complementares, Status Edital). **Cada linha gera o seu próprio edital** (sem agrupamento). O volume em m³ é Comprimento × Largura × Altura, convertido em metros lineares
  - `Municípios` — colunas `Município` e `Região Administrativa`, usada como referência para preencher a região administrativa no cabeçalho de ambos os tipos de edital. Município não cadastrado nesta aba interrompe o script com erro
  - `Membros CADA` — colunas `NOME`, `CARGO` e `STATUS`; o membro com `STATUS` "Ativo" é quem assina o edital. Se nenhum membro estiver "Ativo", assina fixo "IARA LOPES DA SILVA" / "Coordenadora". O script para com erro apenas se houver **mais de um** membro "Ativo" ao mesmo tempo (assinatura ambígua)
- `Editais Elaborados/` — pasta de saída dos editais gerados em `.docx`
- `Modelos/modelo_edital.docx` — template do edital de caixa, com placeholders Jinja2 (`{{ data_edital }}`, `{{ regiao }}`, `{{ municipio }}`, `{{ total_caixas }}`, `{{ nome_membro }}`, `{{ cargo_membro }}` etc.) e um trecho repetido para cada item do detalhamento (`{% for item in itens %} ... {% endfor %}`)
- `Modelos/modelo_edital_massa.docx` — template do edital de massa, com placeholders (`{{ altura }}`, `{{ comprimento }}`, `{{ largura }}`, `{{ metros_cubicos }}`, `{{ total_metros_lineares }}`, `{{ observacoes_complementares }}` etc.)

Por serem `.docx` reais, a formatação (negrito, títulos, espaçamento) de ambos os templates pode ser ajustada diretamente no Word, sem alterar o código.

## O que precisa ser melhorado

### Bugs conhecidos

- **Formatação de data**: `str(row['Data Limite'])` imprime timestamps do pandas em formato bruto (ex.: `2026-01-01 00:00:00`) em vez de um formato de data legível (`01/01/2026`).
- **Coluna fixa `"O{linha}"` (caixa) e `"H{linha}"` (massa) para marcar status no Excel**: se a planilha for reorganizada, a marcação "Edital criado" passa a ser escrita na coluna errada sem qualquer erro.
- **Coluna "Data do pedido" da aba "Edital de Massa" não é usada**: o script lê a coluna, mas o cabeçalho do edital de massa usa a data de hoje (`datetime.now()`), igual ao edital de caixa — não a data do pedido registrada na planilha.

### Melhorias estruturais

- **Template `.docx` aberto a cada grupo processado**: o `DocxTemplate(MODELO_EDITAL)` é instanciado uma vez por grupo (já bem mais leve do que reabrir arquivo por linha, como antes). Ainda pode ser otimizado para carregar o template uma única vez fora do loop.
- **Excel salvo a cada grupo processado**: `load_workbook`/`wb.save()` roda uma vez por grupo, reescrevendo o arquivo inteiro repetidamente. O ideal é abrir uma vez, atualizar todas as linhas e salvar uma única vez ao final.
- **Ausência de tratamento de erros** para situações comuns: arquivo aberto no Excel/Word, planilha ou colunas ausentes, valores não numéricos em "Quantidade" etc.
- **Uso de `exit()`** em vez de `sys.exit()` — funciona em REPL, mas não é a prática recomendada em scripts.
- **Todo o código em nível de módulo**, sem funções (`main()`) nem `if __name__ == "__main__":` — dificulta testes, reuso e uma futura integração com uma interface gráfica.
- **Redundância**: `municipio` já vem da chave do `groupby` (edital de caixa), mas é reatribuído logo em seguida a partir do próprio grupo.
- **`wb = load_workbook(ARQUIVO)` dentro do loop de massa**, assim como no de caixa: o Excel é reaberto e resalvo a cada linha processada, em vez de uma única vez ao final.
