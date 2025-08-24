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
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    # excel.Calculation = win32.constants.xlCalculationManual

    workbook = None

    try:
        workbook = excel.Workbooks.Open(str(caminho_origem))
        sheet = workbook.Sheets(1)
        used_range = sheet.UsedRange
        valores = used_range.Value  # tupla 2D de toda a planilha

        # --- Determinar linhas a remover em memória ---
        linhas_para_remover = set()

        # Remover linhas 1-3
        linhas_para_remover.update([1, 2, 3])

        # Verificar última linha
        ultima_linha = len(valores)
        ultima_linha_valores = valores[-1][:5]  # primeiras 5 colunas
        if any(valor and "Gerado em" in str(valor) for valor in ultima_linha_valores):
            linhas_para_remover.add(ultima_linha)

        # --- Remover todas as imagens da linha 1 em batch ---
        shapes_para_remover = [s for s in sheet.Shapes if s.TopLeftCell.Row == 1]
        if shapes_para_remover:
            sheet.Shapes.Range([s.Name for s in shapes_para_remover]).Delete()
        log(f"[TRATAMENTO] Total de imagens removidas: {len(shapes_para_remover)}")

        # --- Deletar todas as linhas de uma vez em batch ---
        if linhas_para_remover:
            # Excel aceita intervalos consecutivos, então podemos ordenar e agrupar
            for row in sorted(linhas_para_remover, reverse=True):
                sheet.Rows(row).Delete()
            log(f"[TRATAMENTO] Linhas removidas: {sorted(linhas_para_remover)}")

        # --- Ajustar formatação de toda a UsedRange ---
        sheet.UsedRange.WrapText = False
        log("[TRATAMENTO] Quebra de texto desativada.")

        # Garantir pasta destino
        caminho_destino.parent.mkdir(parents=True, exist_ok=True)

        # Salvar
        salvar_excel(workbook, caminho_destino)
        return caminho_destino

    except Exception as e:
        log(f"[TRATAMENTO] Erro ao processar arquivo: {e}")
        return None

    finally:
        if workbook:
            workbook.Close(SaveChanges=False)
        # excel.Calculation = win32.constants.xlCalculationAutomatic
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
            (r"Relatório RM \(Jira\).*\.xls", "Relatório RM (Jira).xlsx"),
            (r"Project Room \(Jira\).*\.xls", "Project Room (Jira).xlsx"),
            (
                r"Filtro Incidentes - Garantia de Projetos \(Jira\).*\.xls",
                "Filtro Incidentes (Jira).xlsx",
            ),
            (r"Projetos \(Jira\).*\.xls", "Projetos (Jira).xlsx"),
            (r"Defeitos SKY AD \(Jira\).*\.xls", "Defeitos SKY AD (Jira).xlsx"),
        ]

        with log_tempo("[ARQUIVOS] Conversão e tratamento dos .xls"):
            processar_arquivos_xls(data, mapeamento, del_xls=False)


if __name__ == "__main__":
    main()
