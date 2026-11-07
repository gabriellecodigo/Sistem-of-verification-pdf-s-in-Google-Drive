Verificador genérico de arquivos do Google Drive

Script reutilizável para comparar valores de uma ou mais colunas do Google Sheets com nomes de arquivos armazenados em uma pasta do Google Drive e em suas subpastas. O sistema usa o serviço integrado `DriveApp`; não exige ativar a API avançada do Drive e não lê o conteúdo dos arquivos.

## Descrição curta para o GitHub

Verificador configurável de arquivos do Google Drive para Google Sheets, feito em Google Apps Script. Permite selecionar abas e colunas, aplicar regras de comparação por nome, filtrar extensões e ignorar registros por condição. Gera um relatório separado e permite revisar duplicidades sem alterar os arquivos de origem.

## Arquivos

- `DriveFileVerifier.gs`: menu, leitura da planilha, busca no Drive, comparação e geração do relatório.
- `DriveVerifierSidebar.html`: painel lateral de configuração e revisão de duplicidades.
- `appsscript.json`: exemplo de manifesto com os escopos mínimos usados pelo módulo.

## Instalação

1. Abra a planilha Google Sheets e acesse **Extensões → Apps Script**.
2. Crie um arquivo de script chamado `DriveFileVerifier` e cole o conteúdo de `DriveFileVerifier.gs`.
3. Crie um arquivo HTML chamado `DriveVerifierSidebar` e cole o conteúdo de `DriveVerifierSidebar.html`.
4. Se o projeto usar escopos definidos manualmente, copie os escopos de `appsscript.json` para o manifesto do projeto. Se já houver escopos, some os necessários em vez de apagar os existentes.
5. Salve e recarregue a planilha. Use o menu **Verificador Drive → Abrir painel** e autorize o acesso solicitado.

O código define `onOpen()` para criar seu menu. Se for instalado junto com outro sistema que também define `onOpen()`, mantenha apenas uma função `onOpen()` e chame `DV_addMenu_()` dentro dela.

## Como usar

1. Informe o link ou ID da pasta do Drive que será pesquisada. A busca inclui subpastas.
2. Escolha a aba e marque uma ou mais colunas de origem. Os títulos são lidos da linha 1.
3. Informe a primeira linha de dados, normalmente `2`.
4. Escolha o modo de comparação:
   - **Igual ao nome extraído**: o valor da célula deve ser igual à chave obtida do arquivo.
   - **Nome começa com o valor**: a chave do arquivo começa pelo valor da célula.
   - **Nome contém o valor**: a chave do arquivo contém o valor da célula.
5. Opcionalmente, informe um padrão de expressão regular para extrair a chave. Use sintaxe JavaScript, sem barras delimitadoras, e coloque o valor comparável no primeiro grupo entre parênteses. Por exemplo, para um arquivo `123456789012_manifesto.pdf`, use `^(.+?)_manifesto.*$`; a chave extraída será `123456789012`.
6. Filtre extensões com valores separados por vírgula ou ponto e vírgula, como `pdf, xml`. Deixe vazio para aceitar qualquer extensão.
7. Se necessário, selecione uma coluna condicional e valores que devem ser ignorados, como `Cancelado; Inativo`. Esses registros aparecem no relatório com status **Ignorado pela condição**.
8. Clique em **Verificar Drive**. O sistema cria uma nova aba `DV_RESULTADO_...` e deixa a aba de origem intacta.
9. Se houver duplicidades, o painel permite manter todos os arquivos no relatório ou escolher qual arquivo considerar. A decisão vale para o relatório atual e é anotada nele; nenhum arquivo é enviado à lixeira ou modificado.

## Relatório

Cada linha corresponde a um valor lido e a um arquivo correspondente. Os resultados incluem status, aba/linha/coluna de origem, valor comparado, nome e links do arquivo e da pasta, quantidade de correspondências e tratamento de duplicidade. Os dados internos usados para aplicar as escolhas ficam em colunas ocultas no relatório.

Possíveis status:

- **Encontrado**: uma correspondência.
- **Não encontrado**: nenhuma correspondência.
- **Duplicado**: mais de uma correspondência para a mesma coluna e valor.
- **Ignorado pela condição**: o valor da coluna condicional consta na lista de exclusão.

## Regras de comparação

- Sem expressão regular, a chave do arquivo é o nome sem a última extensão. Por exemplo, `12345.pdf` vira `12345`.
- Com expressão regular, o primeiro grupo de captura vira a chave. Nomes que não corresponderem ao padrão não entram na comparação.
- Por padrão, letras maiúsculas/minúsculas são ignoradas. Marque **Diferenciar maiúsculas de minúsculas** para exigir a mesma caixa.
- Espaços externos são removidos e sequências de espaços são reduzidas a um espaço.
- A verificação usa somente os nomes dos arquivos; não executa OCR e não abre PDFs ou outros documentos.
- Para pastas grandes, prefira comparação exata. Os modos “começa com” e “contém” comparam cada nome de arquivo com os valores distintos selecionados e podem levar mais tempo.

## Permissões e limites

O usuário que executa o script precisa ter acesso de leitura à pasta e autorização para o Drive e a planilha. O serviço `DriveApp` é integrado ao Apps Script; não é necessário habilitar o serviço avançado Drive API. Em ambientes Google Workspace, a política do administrador ainda pode restringir o acesso.

A busca percorre recursivamente a pasta configurada. Pastas muito grandes podem exceder o tempo ou as cotas do Apps Script; nesses casos, reduza a pasta ou aplique um filtro de extensão e comparação exata. A busca é somente leitura no Drive e cria uma aba de relatório na planilha.
