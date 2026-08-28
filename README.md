# Editais de Eliminação (Expurgo)

Script em Python para automatizar a elaboração de rascunhos de **editais de eliminação de documentos** (expurgo) a partir de uma planilha de controle.

> 📄 **Dossiê técnico para a Diretoria**: [`Documentacao/Dossie_Tecnico_Editais_Eliminacao.docx`](Documentacao/Dossie_Tecnico_Editais_Eliminacao.docx) — explica o sistema (arquitetura, fluxo de execução, regras de cálculo, validações) e traz indicadores de desempenho antes/depois da automação (outubro/2025), a partir do histórico real de editais publicados.

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
- Bibliotecas padrão do Python: `io`, `math`, `os`, `re`, `sys`, `datetime`

Instale as dependências com:

```
pip install -r requirements.txt
```

A conversão de quantidades numéricas por extenso (edital de caixa) é feita por uma função própria (`numero_por_extenso`), sem depender de biblioteca externa. Isso porque o [num2words](https://github.com/savoirfairelinux/num2words) não flexiona em gênero para português (`pt_BR`) — ele nem aceita o parâmetro `gender`, e sempre retorna a forma masculina (ex.: "dois", nunca "duas"). Como o script sempre se refere a "caixas" (substantivo feminino), foi implementada uma conversão própria (0 a 999.999) sempre no feminino, validada contra a saída do `num2words` (masculina, com flexão manual) em quase 30 mil números.

A data por extenso do cabeçalho também é montada por uma função própria (`data_por_extenso`), com os nomes dos meses em português fixos no código — em vez de depender de `locale.setlocale`, que usa nomes de locale (`pt_BR.UTF-8`) específicos de Linux/macOS e não funciona nas máquinas Windows do DETRAN.

O arquivo Excel e os templates `.docx` são abertos **uma única vez**, fora dos loops de geração: `load_workbook(ARQUIVO)` roda uma vez no início e `wb.save(ARQUIVO)` uma vez no final (marcando "Edital criado" em memória a cada edital gerado, sem resalvar a cada iteração); os bytes de cada template são lidos do disco uma vez e reutilizados via `io.BytesIO` para instanciar um `DocxTemplate` novo a cada edital (o docxtpl exige uma instância nova por render, mas isso evita reabrir o arquivo do disco a cada vez).

Já o edital de massa não lida com quantidades por extenso — ele converte volume em metros cúbicos para metros lineares multiplicando por um fator fixo (`METROS_LINEARES_POR_METRO_CUBICO = 12`), e formata os números no padrão brasileiro (vírgula decimal, sem zeros à direita) através da função `formatar_numero_br`.

A quebra de linha da "Data Limite" (`quebrar_data_limite`) usa `docxtpl.Listing` em vez de `docxtpl.RichText`: `Listing` insere o texto (com `\n`) diretamente no nó de texto existente do Word, e o docxtpl converte cada `\n` em uma quebra de linha real (`<w:br/>`) após a renderização, preservando a formatação do parágrafo. `RichText` foi testado primeiro e descartado — ele gera uma tag `<w:r>` própria, e como o placeholder `{{ item.data_limite }}` está no meio de um texto fixo ("Data limite:  {{ ... }}") dentro do mesmo run, a substituição resultava em XML inválido (um `<w:r>` aninhado dentro de um `<w:t>`).

Toda célula de texto lida da planilha passa pela função `limpar_texto`, que corrige na origem o artefato `_x000D_` (como o Excel grava uma quebra de linha manual — Alt+Enter — dentro de uma célula, quando o openpyxl não converte esse código de volta para uma quebra de linha de verdade). Como a limpeza acontece uma única vez, no carregamento de cada planilha, o restante do código nunca precisa lidar com esse artefato item por item.

## Organização do código

O código é dividido em funções, agrupadas em três blocos (com comentários no próprio arquivo indicando cada um):

- **Comuns a CAIXA e MASSA**: funções auxiliares (`numero_por_extenso`, `limpar_texto`, `data_por_extenso` etc.) e leitura das abas compartilhadas (`carregar_mapa_regiao`, `carregar_membro_assinante`);
- **Edital de Caixa**: `carregar_editais_caixa`, `montar_itens_detalhamento_caixa`, `gerar_edital_caixa` (um edital) e `gerar_editais_caixa` (todos os editais);
- **Edital de Massa**: `carregar_editais_massa`, `gerar_edital_massa` (um edital) e `gerar_editais_massa` (todos os editais);
- **`main()`**: orquestra as chamadas acima dentro do `try/except` e é o ponto de entrada do script (`if __name__ == "__main__":`).

## Tratamento de erros

O código principal roda dentro de um `try/except` que traduz as falhas mais comuns em mensagens diretas (sem traceback do Python) e encerra o script com `sys.exit()` (código de saída 1):

- **`FileNotFoundError`** — planilha ou template não encontrados (ex.: pasta do SharePoint não sincronizada);
- **`PermissionError`** — planilha aberta no Excel ou template aberto no Word, bloqueando a leitura/escrita;
- **`KeyError`** — coluna esperada não existe na aba (cabeçalho renomeado ou removido);
- **`ValueError`** — inclui aba inexistente na planilha (ex.: "Municípios" renomeada), valor ausente ou não numérico em "Quantidade" (caixa) ou em "Comprimento"/"Largura"/"Altura" (massa), coluna "Status Edital" não encontrada, município ausente na aba "Municípios" e mais de um membro "Ativo" na aba "Membros CADA".

O caso de "Nenhum edital a ser criado" não é um erro — usa `sys.exit()` sem mensagem de erro (código de saída 0), e não é capturado pelos `except` acima (`SystemExit` não herda de `Exception`).

## Estrutura esperada

O script precisa ser executado com a pasta do SharePoint [Editais de Eliminação de Documentos](https://governosp.sharepoint.com/:f:/r/teams/DETRAN-DIVISODEGESTODOCUMENTAL/Documentos%20Compartilhados/Editais%20de%20Elimina%C3%A7%C3%A3o%20de%20Documentos?d=w159ef8b07b224315b0a4aff6c6058d69&csf=1&web=1&e=jMhxke) (DETRAN - Divisão de Gestão Documental) sincronizada localmente pelo OneDrive. O caminho é montado dinamicamente a partir da pasta do usuário logado na máquina (`C:\Users\<usuario>\PRODESP\DETRAN - DIVISÃO DE GESTÃO DOCUMENTAL - Documentos\Editais de Eliminação de Documentos`), então o script funciona em qualquer computador sem precisar alterar o código. Dentro dela:

- `Relacao de Expurgo para Rascunho.xlsx` — planilha de controle, com quatro abas:
  - `Edital de Caixa` — relação de documentos catalogados a eliminar (Nº Processo SEI, Município, Função, Subfunção, Atividade, Série documental, Descrição documental, Data Limite, Quantidade, Observações complementares, Status Edital). Várias linhas com o mesmo **Nº Processo SEI + Município** viram um único edital, com um item para cada linha. Quando "Data Limite" lista mais de 6 anos separados por "/" (ex.: `2001/2002/.../2013`), o script quebra a linha automaticamente — até 6 anos na primeira linha e até 8 por linha nas seguintes — reproduzindo a quebra que a equipe de publicação do Diário Oficial já fazia manualmente por questão de espaço
  - `Edital de Massa` — documentos eliminados em bloco (Data do pedido, Nº Processo SEI, Município, Comprimento, Largura, Altura, Observações complementares, Status Edital). **Cada linha gera o seu próprio edital** (sem agrupamento). O volume em m³ é Comprimento × Largura × Altura, convertido em metros lineares. A coluna "Data do pedido" é apenas de referência manual dos digitadores — não é lida pelo script (o cabeçalho do edital usa a data de hoje)
  - `Municípios` — colunas `Município` e `Região Administrativa`, usada como referência para preencher a região administrativa no cabeçalho de ambos os tipos de edital. Município não cadastrado nesta aba interrompe o script com erro
  - `Membros CADA` — colunas `NOME`, `CARGO` e `STATUS`; o membro com `STATUS` "Ativo" é quem assina o edital. Se nenhum membro estiver "Ativo", assina fixo "IARA LOPES DA SILVA" / "Coordenadora". O script para com erro apenas se houver **mais de um** membro "Ativo" ao mesmo tempo (assinatura ambígua)
- `Editais Elaborados/` — pasta de saída dos editais gerados em `.docx`
- `Modelos/modelo_edital_caixa.docx` — template do edital de caixa, com placeholders Jinja2 (`{{ data_edital }}`, `{{ regiao }}`, `{{ municipio }}`, `{{ total_caixas }}`, `{{ nome_membro }}`, `{{ cargo_membro }}` etc.) e um trecho repetido para cada item do detalhamento (`{% for item in itens %} ... {% endfor %}`)
- `Modelos/modelo_edital_massa.docx` — template do edital de massa, com placeholders (`{{ altura }}`, `{{ comprimento }}`, `{{ largura }}`, `{{ metros_cubicos }}`, `{{ total_metros_lineares }}`, `{{ observacoes_complementares }}` etc.)

Por serem `.docx` reais, a formatação (negrito, títulos, espaçamento) de ambos os templates pode ser ajustada diretamente no Word, sem alterar o código.

