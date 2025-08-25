import time
from pathlib import Path
import win32com.client as win32
from utils import (
    log,
    log_tempo,
    salvar_excel,
    deletar_linhas,
    preparar_pasta,
    localizar_arquivo,
)

APAGAR_XLS = True

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


def processar_arquivo_xlsx_(caminho_origem: Path, caminho_destino: Path, excel):
    xlManual = -4135
    xlAutomatic = -4105

    workbook = None
    try:
        with log_tempo("[TRATAMENTO] Abrir Excel (.xls)"):
            workbook = excel.Workbooks.Open(str(caminho_origem))
            sheet = workbook.Sheets(1)
            sheet.DisplayPageBreaks = False
            valores = sheet.UsedRange.Value

        # Determinar linhas para remover
        linhas_para_remover = {1, 2, 3}
        with log_tempo("[TRATAMENTO] Identificar última linha"):
            ultima_linha = len(valores)
            ultima_linha_valores = valores[-1][:5]
            if any(
                valor and "Gerado em" in str(valor) for valor in ultima_linha_valores
            ):
                linhas_para_remover.add(ultima_linha)

        # Percorrer todas as linhas para remover conteúdo exato "SKYIT-182"
        for idx, linha in enumerate(valores, start=1):
            if any(valor == "SKYIT-182" for valor in linha):
                linhas_para_remover.add(idx)

        # Remover imagem do topo (sempre a 1ª)
        with log_tempo("[TRATAMENTO] Remover imagem topo"):
            if sheet.Shapes.Count > 0:
                sheet.Shapes(1).Delete()
                log("[TRATAMENTO] Imagem do topo removida")

        # Remover linhas
        with log_tempo("[TRATAMENTO] Remover linhas"):
            excel.Calculation = xlManual
            deletar_linhas(sheet, linhas_para_remover)
            excel.CutCopyMode = False
            excel.Calculation = xlAutomatic
        log(f"[TRATAMENTO] Linhas removidas: {sorted(linhas_para_remover)}")

        # Ajustar formatação
        with log_tempo("[TRATAMENTO] Ajustar formatação"):
            sheet.UsedRange.WrapText = False
        log("[TRATAMENTO] Quebra de texto desativada.")

        # Salvar
        caminho_destino.parent.mkdir(parents=True, exist_ok=True)
        with log_tempo("[TRATAMENTO] Salvar arquivo"):
            salvar_excel(workbook, caminho_destino)
        return caminho_destino

    except Exception as e:
        log(f"[TRATAMENTO] Erro ao processar arquivo: {e}")
        return None

    finally:
        if workbook:
            workbook.Close(SaveChanges=False)


def processar_arquivo_xlsx(caminho_origem: Path, caminho_destino: Path, excel):
    xlManual = -4135
    xlAutomatic = -4105

    workbook = None
    try:
        with log_tempo("[TRATAMENTO] Abrir arquivo (.xls)"):
            workbook = excel.Workbooks.Open(str(caminho_origem))
            sheet = workbook.Sheets(1)
            sheet.DisplayPageBreaks = False
            valores = sheet.UsedRange.Value

        # Determinar linhas para remover
        linhas_para_remover = {1, 2, 3}

        with log_tempo("[TRATAMENTO] Identificar última linha e SKYIT-182"):
            ultima_linha = len(valores)
            ultima_linha_valores = valores[-1][:5]

            # Verificar última linha "Gerado em"
            if any(
                valor and "Gerado em" in str(valor) for valor in ultima_linha_valores
            ):
                linhas_para_remover.add(ultima_linha)

            # Identificar linhas com conteúdo exato "SKYIT-182" (mais performático)
            linhas_skyit = [
                idx
                for idx, linha in enumerate(valores, start=1)
                if "SKYIT-182" in linha
            ]
            linhas_para_remover.update(linhas_skyit)

        # Remover imagem do topo (sempre a 1ª)
        with log_tempo("[TRATAMENTO] Remover imagem topo"):
            if sheet.Shapes.Count > 0:
                sheet.Shapes(1).Delete()
                log("[TRATAMENTO] Imagem do topo removida")

        # Remover linhas
        with log_tempo("[TRATAMENTO] Remover linhas"):
            excel.Calculation = xlManual
            deletar_linhas(sheet, linhas_para_remover)
            excel.CutCopyMode = False
            excel.Calculation = xlAutomatic
        log(f"[TRATAMENTO] Linhas removidas: {sorted(linhas_para_remover)}")

        # Ajustar formatação
        with log_tempo("[TRATAMENTO] Ajustar formatação (Quebrar texto)"):
            sheet.UsedRange.WrapText = False

        # Salvar
        caminho_destino.parent.mkdir(parents=True, exist_ok=True)
        salvar_excel(workbook, caminho_destino)

        return caminho_destino

    except Exception as e:
        log(f"[TRATAMENTO] Erro ao processar arquivo: {e}")
        return None

    finally:
        if workbook:
            workbook.Close(SaveChanges=False)


def processar_arquivos_xls(folder_data: Path, arquivos_info: list[dict], del_xls: bool):

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False

    try:
        for regex, novo_nome in arquivos_info:
            arquivo = localizar_arquivo(folder_data, regex)
            if not arquivo:
                log(f"[ARQUIVO] Nenhum arquivo encontrado para padrão: {regex}")
                continue

            # Definir caminho de destino na pasta uploads
            caminho_destino = folder_data / "uploads" / novo_nome

            # Processar o arquivo
            with log_tempo(f"[ARQUIVO] Processar {arquivo.name}"):
                resultado = processar_arquivo_xlsx(arquivo, caminho_destino, excel)

                if resultado and del_xls and arquivo.suffix.lower() == ".xls":
                    arquivo.unlink()
                    log(f"[ARQUIVO] .xls original removido: {arquivo}")
    finally:
        excel.Quit()


def main():
    with log_tempo("[PROCESSAMENTO] .xls para .xlsx"):
        # Diretório onde os arquivos estao
        data = preparar_pasta()

        # Limpeza inicial
        with log_tempo("[DIRETORIO] Limpar pasta uploads"):
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
            processar_arquivos_xls(data, mapeamento, del_xls=APAGAR_XLS)


if __name__ == "__main__":
    main()
