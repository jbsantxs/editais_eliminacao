# Editais de Eliminação (Expurgo)

Script em Python para automatizar a elaboração de rascunhos de **editais de eliminação de documentos** (expurgo) a partir de uma planilha de controle.

## Para que serve

A partir de uma planilha Excel com a relação de documentos a serem eliminados, o script:

- Filtra os registros marcados como **"Criar edital"**;
- Agrupa os itens por **Número de Processo SEI**, **Região Administrativa** e **Município**;
- Ordena os itens de cada edital pela **Série documental**;
- Preenche modelos de texto (cabeçalho, detalhamento e rodapé) com os dados de cada grupo, incluindo a conversão de quantidades numéricas por extenso;
- Gera um documento **Word (.docx)** para cada edital, na pasta de editais elaborados;
- Atualiza a planilha original, marcando os registros processados como **"Edital criado"**.

## Tecnologias utilizadas

- **Python 3**
- [pandas](https://pandas.pydata.org/) — leitura, filtragem, agrupamento e ordenação dos dados da planilha
- [openpyxl](https://openpyxl.readthedocs.io/) — leitura/escrita do arquivo Excel (engine do pandas e atualização de status)
- [python-docx](https://python-docx.readthedocs.io/) — geração dos documentos Word (.docx) dos editais
- [num2words](https://github.com/savoirfairelinux/num2words) — conversão de quantidades numéricas por extenso (em português)
- Bibliotecas padrão do Python: `locale`, `os`, `re`, `datetime`

## Estrutura esperada

O script precisa ser executado com a pasta do SharePoint [Editais de Eliminação de Documentos](https://governosp.sharepoint.com/:f:/r/teams/DETRAN-DIVISODEGESTODOCUMENTAL/Documentos%20Compartilhados/Editais%20de%20Elimina%C3%A7%C3%A3o%20de%20Documentos?d=w159ef8b07b224315b0a4aff6c6058d69&csf=1&web=1&e=jMhxke) (DETRAN - Divisão de Gestão Documental) sincronizada localmente pelo OneDrive. O caminho é montado dinamicamente a partir da pasta do usuário logado na máquina (`C:\Users\<usuario>\PRODESP\DETRAN - DIVISÃO DE GESTÃO DOCUMENTAL - Documentos\Editais de Eliminação de Documentos`), então o script funciona em qualquer computador sem precisar alterar o código. Dentro dela:

- `Relacao de Expurgo.xlsx` — planilha de controle (aba "Edital de Caixa")
- `Editais Elaborados/` — pasta de saída dos editais gerados em `.docx`
- `Modelos/` — pasta com os modelos de texto usados no preenchimento:
  - `modelo_cabeçalho.txt`
  - `modelo_detalhamento.txt`
  - `modelo_rodapé.txt`

## O que precisa ser melhorado

### Bugs conhecidos

- **Locale incompatível com Windows** (`locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')`): o nome de locale usado é de Linux/macOS, mas os caminhos do projeto são de Windows. No Windows isso costuma lançar `locale.Error: unsupported locale setting` — provável causa do erro relatado.
- **Coluna "Descrição documental" não sanitizada**: diferente das demais colunas de texto, ela não passa por `astype(str)` antes do `.replace()`. Se vier vazia (`NaN`) na planilha, quebra com `AttributeError: 'float' object has no attribute 'replace'`.
- **Perda silenciosa de linhas no agrupamento do detalhamento**: o `groupby(...)` usado para montar o detalhamento descarta por padrão qualquer linha com `NaN` em uma das colunas-chave (ex.: `Data Limite` vazia), removendo o item do edital sem nenhum aviso.
- **Formatação de data**: `str(row['Data Limite'])` imprime timestamps do pandas em formato bruto (ex.: `2026-01-01 00:00:00`) em vez de um formato de data legível (`01/01/2026`).
- **Coluna fixa `"O{linha}"` para marcar status no Excel**: se a planilha for reorganizada, a marcação "Edital criado" passa a ser escrita na coluna errada sem qualquer erro.

### Melhorias estruturais

- **Leitura repetida dos templates dentro do loop**: os arquivos de cabeçalho, detalhamento e rodapé são reabertos a cada grupo (o de detalhamento, a cada linha). Deveriam ser lidos uma única vez.
- **Excel salvo a cada grupo processado**: `load_workbook`/`wb.save()` roda uma vez por grupo, reescrevendo o arquivo inteiro repetidamente. O ideal é abrir uma vez, atualizar todas as linhas e salvar uma única vez ao final.
- **Dependência de locale do sistema operacional** para nome do mês por extenso: frágil entre máquinas diferentes; um dicionário de meses em português (ou uso de `babel`) elimina essa dependência.
- **Documento Word gerado sem formatação real**: o `.docx` final é montado concatenando texto e criando um parágrafo por linha, o que impede negrito, títulos, tabelas etc. Vale considerar um template `.docx` real em vez de arquivos `.txt`.
- **Caminhos fixos (hardcoded)** apontando para uma pasta específica do OneDrive de um usuário — torna o script não portátil entre máquinas/usuários.
- **Ausência de tratamento de erros** para situações comuns: arquivo aberto no Excel/Word, planilha ou colunas ausentes, valores não numéricos em "Quantidade" etc.
- **Uso de `exit()`** em vez de `sys.exit()` — funciona em REPL, mas não é a prática recomendada em scripts.
- **Todo o código em nível de módulo**, sem funções (`main()`) nem `if __name__ == "__main__":` — dificulta testes, reuso e uma futura integração com uma interface gráfica.
- **Redundância**: `regiao` e `municipio` já vêm da chave do `groupby`, mas são reatribuídos logo em seguida a partir do próprio grupo.
- **Lógica de substituição de placeholders duplicada** em cabeçalho, detalhamento e rodapé — poderia ser extraída em uma função única reutilizável.
