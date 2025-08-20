import processar_xls
import relatorio_garantias
import relatorio_project_room

from utils import log_tempo


def main():
    with log_tempo("[MASTER] execução"):
        processar_xls.main()
        relatorio_garantias.main()
        relatorio_project_room.main()


if __name__ == "__main__":
    main()
