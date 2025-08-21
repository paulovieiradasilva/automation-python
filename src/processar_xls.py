from pathlib import Path
import win32com.client as win32
from utils import log, log_tempo, localizar_arquivo, preparar_pasta, salvar_excel


def limpar_uploads(pasta_base: Path) -> int:
    """Remove todos os .xlsx da pasta uploads com logging"""
    pasta = pasta_base / "uploads"
    if not pasta.exists():
        log("[DIRETORIO] Pasta uploads não encontrada")
        return 0

    removidos = 0
    for arq in pasta.glob("*.xlsx"):
        try:
            arq.unlink()
            log(f"[DIRETORIO] Removido: {arq.name}")
            removidos += 1
        except Exception as e:
            log(f"[DIRETORIO] Erro ao remover {arq.name}: {str(e)}")

    log(f"[DIRETORIO] Total removido: {removidos}")
    return removidos


def processar_arquivo_xlsx(caminho_origem: Path, caminho_destino: Path):
    """
    Processa o arquivo Excel e salva na pasta de destino especificada
    """
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False  # Desativa alertas para substituição
    workbook = None

    # OTIMIZAÇÕES CRÍTICAS
    excel.ScreenUpdating = False
    excel.EnableEvents = False  # Adicionado: desativa eventos

    try:
        # Abrir arquivo original
        workbook = excel.Workbooks.Open(str(caminho_origem))
        sheet = workbook.Sheets(1)

        # Etapa 1: Remover imagens PRIMEIRO (antes de outras modificações)
        total_imagens = 0
        for shape in list(sheet.Shapes):  # Usamos list() para criar uma cópia
            if shape.TopLeftCell.Row == 1:  # Verifica se a imagem está na linha 1
                shape.Delete()  # Remove a imagem
                total_imagens += 1
                log(f"[TRATAMENTO] Total de imagens removidas: {total_imagens}")

        # Etapa 2: Remover linhas 1-3
        sheet.Rows("1:3").Delete()
        log("[TRATAMENTO] Linhas 1, 2 e 3 removidas.")

        # Etapa 3: Verificar e remover última linha se necessário
        ultima_linha = sheet.UsedRange.Rows.Count
        valores_ultima_linha = [
            sheet.Cells(ultima_linha, col).Value for col in range(1, 6)
        ]
        if any(valor and "Gerado em" in str(valor) for valor in valores_ultima_linha):
            sheet.Rows(ultima_linha).Delete()
            log(
                f"[TRATAMENTO] Última linha ({ultima_linha}) contendo 'Gerado em' removida."
            )

        # Etapa 4: Ajustar formatação
        sheet.UsedRange.WrapText = False
        log("[TRATAMENTO] Quebra de texto desativada.")

        # Garantir que a pasta de destino existe
        caminho_destino.parent.mkdir(parents=True, exist_ok=True)

        # Salvar como XLSX na pasta uploads
        salvar_excel(workbook, caminho_destino)

        return caminho_destino

    except Exception as e:
        log(f"[TRATAMENTO] Erro ao processar arquivo: {e}")
        return None

    finally:
        if workbook:
            workbook.Close(SaveChanges=False)
        excel.Quit()


def processar_arquivos_xls(folder_data: Path, arquivos_info: list[dict], del_xls: bool):
    for regex, novo_nome in arquivos_info:
        arquivo = localizar_arquivo(folder_data, regex)
        if not arquivo:
            log(f"[ARQUIVO] Nenhum arquivo encontrado para padrão: {regex}")
            continue

        # Definir caminho de destino na pasta uploads
        caminho_destino = folder_data / "uploads" / novo_nome

        # Processar o arquivo
        with log_tempo(f"[ARQUIVO] Processando {arquivo.name}"):
            resultado = processar_arquivo_xlsx(arquivo, caminho_destino)

            if resultado and del_xls and arquivo.suffix.lower() == ".xls":
                arquivo.unlink()
                log(f"[ARQUIVO] .xls original removido: {arquivo}")


def main():
    with log_tempo("[PROCESSAMENTO] .xls para .xlsx"):
        # Diretório onde os arquivos estao
        data = preparar_pasta()

        # Limpeza inicial
        with log_tempo("[DIRETORIO] Limpeza da pasta uploads"):
            limpar_uploads(data)

        mapeamento = [
            (
                r"Filtro Incidentes - Garantia de Projetos \(Jira\).*\.xls",
                "Filtro Incidentes (Jira).xlsx",
            ),
            (r"Projetos \(Jira\).*\.xls", "Projetos (Jira).xlsx"),
            (r"Defeitos SKY AD \(Jira\).*\.xls", "Defeitos SKY AD (Jira).xlsx"),
            (r"Project Room \(Jira\).*\.xls", "Project Room (Jira).xlsx"),
            (r"Relatório RM \(Jira\).*\.xls", "Relatório RM (Jira).xlsx"),
        ]

        with log_tempo("[ARQUIVOS] Conversão e tratamento dos .xls"):
            processar_arquivos_xls(data, mapeamento, del_xls=False)


if __name__ == "__main__":
    main()
