# 🚀 Projeto AUTOMAÇÃO_V3

Este projeto contém scripts de automação em **Python** para processar relatórios em Excel (.xls).

---

## 📂 Estrutura do Projeto
```
AUTOMAÇÃO_V3/
│── src/                
│   ├── config.py                 # Configurações gerais
│   ├── main.py                   # Script principal
│   ├── processar_xls.py          # Funções para processar planilhas
│   ├── relatorio_garantias.py    # Relatório de garantias
│   ├── relatorio_project_room.py # Relatório Project Room
│   ├── utils.py                  # Funções auxiliares
│   └── data/                     # Pasta para colocar as extrações .xls (IGNORADA pelo Git)
│
│── venv/                # Ambiente virtual (IGNORADO pelo Git)
│── requirements.txt     # Dependências do projeto
│── README.md
└── .gitignore
```

---

## ⚙️ Configuração do Ambiente

1. **Criar ambiente virtual**
   ```bash
   python -m venv venv
   ```

2. **Ativar ambiente virtual**
   - Windows:
     ```bash
     venv\Scripts\activate
     ```
   - Linux/Mac:
     ```bash
     source venv/bin/activate
     ```

3. **Instalar dependências**
   ```bash
   pip install -r requirements.txt
   ```

---

## 📥 Entrada de Dados

A pasta `src/data/` deve ser utilizada para armazenar os arquivos de extração **.xls** que serão processados pelo sistema.

⚠️ **Importante:** Essa pasta está no `.gitignore`, portanto os arquivos `.xls` **não serão versionados no repositório**.

Exemplo de uso:
```
src/data/
│── extracao_clientes.xls
│── extracao_projetos.xls
```

---

## ▶️ Executando o Projeto

Após configurar o ambiente e colocar os arquivos `.xls` na pasta `src/data/`, execute o script principal:

```bash
python src/main.py
```

---

## 📝 Uso do `.gitignore`

O arquivo `.gitignore` já está configurado para ignorar arquivos temporários e pastas que não devem ser versionadas:

```gitignore
# Ignorar cache do Python
__pycache__/
*.pyc
*.pyo

# Ignorar venv
venv/

# Ignorar pasta data dentro de src
/src/data/
```

---

## ✅ Boas Práticas

- Coloque sempre os arquivos de entrada `.xls` na pasta `src/data/`.
- Nunca faça commit de arquivos temporários ou da pasta `data/`.
- Use `git status` antes de commitar para garantir que apenas arquivos de código/configuração estão sendo versionados.
- Atualize o `requirements.txt` sempre que adicionar uma nova biblioteca:
  ```bash
  pip freeze > requirements.txt
  ```

---