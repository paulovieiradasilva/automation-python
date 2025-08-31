import datetime

import processar_xls
import relatorio_garantias
import relatorio_project_room
from utils import log_tempo

# Verifica se hoje é segunda-feira (0)
hoje = datetime.datetime.today().weekday()
dias_da_semana = [0]  # 0 = segunda-feira


def main():
    with log_tempo("[MASTER] Automação"):
        relatorio_garantias.main()

        if hoje in dias_da_semana:
            relatorio_project_room.main()


if __name__ == "__main__":
    main()
