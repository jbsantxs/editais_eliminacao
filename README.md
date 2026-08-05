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

- `Relacao de Expurgo.xlsx` — planilha de controle (aba "Edital de Caixa")
- `Editais Elaborados/` — pasta de saída dos editais gerados em `.docx`
- `Modelos/` — pasta com os modelos de texto usados no preenchimento:
  - `modelo_cabeçalho.txt`
  - `modelo_detalhamento.txt`
  - `modelo_rodapé.txt`
