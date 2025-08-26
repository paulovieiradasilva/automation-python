# Documentação Funcional

## Visão Geral do Sistema

Este sistema Python é projetado para automatizar o processamento de planilhas Excel, especificamente para gerar relatórios de garantias e de Project Room. Ele lida com a conversão de arquivos `.xls` para `.xlsx`, limpeza de dados e organização de informações em diferentes abas de relatórios.

## Estrutura do Projeto

O projeto é composto pelos seguintes módulos:

- `main.py`: Ponto de entrada da aplicação, orquestra a execução dos processos.
- `processar_xls.py`: Responsável por processar e converter arquivos Excel.
- `relatorio_garantias.py`: Gera o relatório de garantias, filtrando e organizando dados de incidentes.
- `relatorio_project_room.py`: Gera o relatório de Project Room, com lógica similar ao de garantias.
- `utils.py`: Contém funções utilitárias para logging, manipulação de arquivos e operações em planilhas Excel.
- `config.py`: Define mapeamentos de colunas e outras configurações globais.

## Detalhamento dos Módulos

### `main.py`

**Propósito:** Orquestrar a execução dos processos de automação.

**Funcionalidades:**
- Executa a função principal de `processar_xls` para converter e tratar os arquivos Excel.
- Executa a função principal de `relatorio_garantias` para gerar o relatório de garantias.
- Condicionalmente, executa a função principal de `relatorio_project_room` apenas às segundas-feiras.
- Utiliza o `log_tempo` para registrar o tempo de execução de cada etapa principal.

### `processar_xls.py`

**Propósito:** Converter arquivos `.xls` para `.xlsx` e realizar pré-processamento.

**Funcionalidades:**
- `limpar_uploads(pasta_base: Path)`: Remove todos os arquivos `.xlsx` da pasta `uploads` dentro do diretório `data`.
- `processar_arquivo_xlsx(caminho_origem: Path, caminho_destino: Path, excel)`: Abre um arquivo `.xls` usando `win32com.client`, remove linhas específicas (cabeçalhos, 


linhas com 'SKYIT-182', e a linha 'Gerado em'), remove imagens do topo e ajusta a formatação antes de salvar como `.xlsx`.
- `processar_arquivos_xls(folder_data: Path, arquivos_info: list[dict], del_xls: bool)`: Itera sobre uma lista de arquivos `.xls` (definidos por regex e novo nome), processa cada um usando `processar_arquivo_xlsx`, e opcionalmente deleta o arquivo `.xls` original.
- `main()`: Orquestra o processo de limpeza da pasta `uploads`, e a conversão e tratamento dos arquivos `.xls` listados no mapeamento.

### `relatorio_garantias.py`

**Propósito:** Gerar o relatório de garantias, consolidando dados de incidentes.

**Funcionalidades:**
- `abrir_planilhas()`: Abre as planilhas de origem (`Filtro Incidentes (Jira).xlsx`, `Projetos (Jira).xlsx`) e as planilhas de destino (`Relatorio Incidentes_Garantia_Projetos_v5.xlsx`) para processamento.
- `preparar_mapeamento(ws_origem, ws_destino, mapa_colunas)`: Prepara os índices das colunas de origem e destino com base em um mapeamento customizado.
- `preparar_mapeamento_simples(ws_origem, ws_destino)`: Prepara os índices das colunas de origem e destino para um mapeamento direto (coluna com o mesmo nome).
- `copiar_para_aba(...)`: Copia linhas de uma planilha de origem para uma aba de destino, mantendo a formatação e ajustando fórmulas se necessário.
- `obter_ultima_linha(ws, coluna_chave)`: Retorna o índice da última linha com dados em uma coluna específica.
- `processar_rf(ws_origem, ws_destino)`: Filtra e copia linhas com status 'Resolvido' e 'Finalizado' da planilha de origem para a aba 'Resolvidos-Fechados' do relatório de garantias.
- `processar_ri(ws_origem, ws_destino)`: Filtra e copia linhas com status diferente de 'Resolvido' e 'Finalizado' da planilha de origem para a aba 'RI' do relatório de garantias.
- `processar_projetos(ws_origem, ws_destino)`: Copia todas as linhas da planilha de origem de projetos para a aba 'Projetos' do relatório de garantias.
- `main()`: Orquestra a abertura das planilhas, o processamento das abas de projetos, RI e RF, e salva o relatório final.

### `relatorio_project_room.py`

**Propósito:** Gerar o relatório de Project Room, consolidando dados de incidentes.

**Funcionalidades:**
- `abrir_planilhas()`: Abre a planilha de origem (`Project Room (Jira).xlsx`) e as planilhas de destino (`Relatorio de Incidentes_Project Room_v1.xlsx`) para processamento.
- `preparar_mapeamento(ws_origem, ws_destino, mapa_colunas)`: Prepara os índices das colunas de origem e destino com base em um mapeamento customizado.
- `copiar_para_aba(...)`: Copia linhas de uma planilha de origem para uma aba de destino, mantendo a formatação e ajustando fórmulas se necessário.
- `obter_ultima_linha(ws, coluna_chave)`: Retorna o índice da última linha com dados em uma coluna específica.
- `processar_rf(ws_origem, ws_destino)`: Filtra e copia linhas com status 'Resolvido' e 'Finalizado' da planilha de origem para a aba 'Resolvidos-Fechados' do relatório de Project Room.
- `processar_ri(ws_origem, ws_destino)`: Filtra e copia linhas com status diferente de 'Resolvido', 'Finalizado' e 'Cancelado' da planilha de origem para a aba 'Relatório de Incidentes' do relatório de Project Room.
- `main()`: Orquestra a abertura das planilhas, o processamento das abas RI e RF, e salva o relatório final.

### `utils.py`

**Propósito:** Fornecer funções utilitárias para logging, manipulação de arquivos e operações em planilhas Excel.

**Funcionalidades:**
- `log(message)`: Função simples para registrar mensagens no console.
- `log_tempo(mensagem)`: Context manager que mede e loga o tempo de execução de um bloco de código.
- `preparar_destino(ws_destino, linha_modelo)`: Limpa uma planilha de destino, mantendo uma linha modelo.
- `preparar_pasta(subpasta)`: Garante a existência de uma estrutura de pastas (`data` e subpastas) e retorna o caminho absoluto.
- `localizar_arquivo(pasta, nome_arquivo)`: Encontra um arquivo `.xls` em uma pasta que corresponde a um padrão de nome.
- `ajustar_formula_linha(formula, linha_origem, linha_destino)`: Ajusta referências de linha em fórmulas Excel ao copiar células.
- `copiar_linha_com_formula(...)`: Copia uma linha para outra na mesma planilha, com opções para copiar colunas específicas, colunas extras e ajustar fórmulas.
- `filtrar_linhas(...)`: Filtra linhas de uma planilha com base em valores de uma coluna de status (incluir/excluir) e retorna os índices das linhas filtradas.
- `obter_ultima_linha_com_dados(ws, coluna_chave)`: Retorna o número da última linha com dados em uma coluna específica.
- `deletar_linhas(sheet, linhas, log_prefix)`: Deleta múltiplas linhas em blocos consecutivos de uma planilha Excel usando COM.
- `salvar_excel(workbook, caminho)`: Salva um workbook do Excel em um caminho especificado.

### `config.py`

**Propósito:** Armazenar configurações globais do sistema.

**Conteúdo:**
- `MAPEAMENTO_COLUNAS`: Um dicionário que mapeia nomes de colunas de origem para nomes de colunas de destino, usado para padronizar os cabeçalhos nos relatórios.
- `COLUNAS_RELATORIO`: Uma lista de letras de colunas que são consideradas 


colunas relevantes para os relatórios.

## Fluxo de Execução

1.  **Início (`main.py`):**
    *   A automação é iniciada.
    *   O tempo de execução é logado.

2.  **Processamento de Arquivos Excel (`processar_xls.py`):**
    *   A pasta `uploads` é limpa de arquivos `.xlsx` antigos.
    *   Arquivos `.xls` são localizados na pasta `data` (ou subpastas).
    *   Cada arquivo `.xls` é aberto, tem linhas específicas removidas (cabeçalhos, linhas com "SKYIT-182", "Gerado em"), imagens são removidas, e a formatação é ajustada.
    *   Os arquivos são salvos como `.xlsx` na pasta `uploads` com nomes padronizados.
    *   Opcionalmente, os arquivos `.xls` originais são deletados.

3.  **Geração do Relatório de Garantias (`relatorio_garantias.py`):**
    *   As planilhas de origem (`Filtro Incidentes (Jira).xlsx`, `Projetos (Jira).xlsx`) e a planilha de destino (`Relatorio Incidentes_Garantia_Projetos_v5.xlsx`) são abertas.
    *   **Processamento de Projetos:** Todas as linhas da planilha de origem de projetos são copiadas para a aba "Projetos" do relatório de garantias.
    *   **Processamento de RI (Relatório de Incidentes):** Linhas da planilha de origem de filtros com status diferente de 'Resolvido' e 'Finalizado' são copiadas para a aba "RI".
    *   **Processamento de RF (Resolvidos e Fechados):** Linhas da planilha de origem de filtros com status 'Resolvido' e 'Finalizado' são copiadas para a aba "Resolvidos-Fechados".
    *   O relatório de garantias é salvo.
    *   Todas as planilhas são fechadas.

4.  **Geração do Relatório de Project Room (`relatorio_project_room.py` - Condicional):**
    *   Este passo é executado apenas se o dia atual for segunda-feira.
    *   As planilhas de origem (`Project Room (Jira).xlsx`) e a planilha de destino (`Relatorio de Incidentes_Project Room_v1.xlsx`) são abertas.
    *   **Processamento de RI (Relatório de Incidentes):** Linhas da planilha de origem de filtros com status diferente de 'Resolvido', 'Finalizado' e 'Cancelado' são copiadas para a aba "Relatório de Incidentes".
    *   **Processamento de RF (Resolvidos e Fechados):** Linhas da planilha de origem de filtros com status 'Resolvido' e 'Finalizado' são copiadas para a aba "Resolvidos-Fechados".
    *   O relatório de Project Room é salvo.
    *   Todas as planilhas são fechadas.

5.  **Fim:** A automação é concluída e o tempo total de execução é logado.

## Dependências

- `python` (versão 3.x)
- `openpyxl`: Para manipulação de arquivos `.xlsx`.
- `pywin32`: Para interação com o Excel via COM (necessário para `.xls` e manipulação de objetos Excel como imagens e deleção de linhas).

## Como Executar

1.  **Preparação do Ambiente:**
    *   Certifique-se de ter o Python instalado.
    *   Instale as dependências:
        ```bash
        pip install openpyxl pywin32
        ```
2.  **Organização dos Arquivos:**
    *   Coloque os arquivos `.xls` de origem na pasta `data` (ou na subpasta `data/uploads` se for o caso).
    *   Certifique-se de que os modelos de relatório (`Relatorio Incidentes_Garantia_Projetos_v5.xlsx` e `Relatorio de Incidentes_Project Room_v1.xlsx`) estejam na pasta `data`.
3.  **Execução:**
    *   Execute o script principal:
        ```bash
        python main.py
        ```

## Considerações Importantes

- O script depende da instalação do Microsoft Excel no ambiente de execução para o processamento de arquivos `.xls` e manipulação de objetos COM.
- Os nomes dos arquivos de origem (`Filtro Incidentes (Jira).xlsx`, `Projetos (Jira).xlsx`, `Project Room (Jira).xlsx`) e os modelos de relatório são esperados em locais específicos (`data` e `data/uploads`). Qualquer alteração nesses nomes ou locais pode exigir ajustes no código.
- A lógica de filtragem e cópia de dados é baseada em nomes de colunas e status específicos (e.g., "Resolvido", "Finalizado"). Alterações nos dados de origem podem impactar a funcionalidade.
- O relatório de Project Room é gerado apenas às segundas-feiras, conforme a lógica definida em `main.py`.


