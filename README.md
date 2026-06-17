# SeuMed — Sistema de Repasses Médicos

Plataforma para apurar honorários de repasses médicos, gerar os arquivos de
importação da **OMIE** (contas a pagar e a receber) e emitir o documento de
**repasse** enviado ao médico para conferência — a partir de uma fonte única
da verdade, com histórico e auditoria.

## Fluxo do negócio

1. **MedPlus** exporta o relatório da agenda de procedimentos realizados.
2. O sistema aplica as **regras de honorário** e apura o repasse de cada médico.
3. O sistema gera os arquivos de **importação da OMIE** (contas a pagar / a receber).
4. O sistema emite o **repasse médico** (PDF) para conferência do profissional.

## Tecnologia

- **Python 3.13** + **Django 6.0** (admin, ORM, autenticação, templates)
- **pandas** / **openpyxl** — leitura e geração de planilhas
- **reportlab** — geração de PDFs
- Banco de dados: **SQLite** (desenvolvimento) — migrável para PostgreSQL

## Como rodar localmente (Windows)

```powershell
# 1. Instalar dependências (uma vez)
.venv\Scripts\python.exe -m pip install -r requirements.txt

# 2. Aplicar migrações do banco
.venv\Scripts\python.exe manage.py migrate

# 3. (uma vez) criar usuário administrador
.venv\Scripts\python.exe manage.py createsuperuser

# 4. Subir o servidor
.venv\Scripts\python.exe manage.py runserver
```

Acesse: <http://localhost:8000/> (sistema) e <http://localhost:8000/admin/> (administração).

## Estrutura

```
config/      Configurações do projeto (settings, urls)
repasses/    Módulo de negócio (modelos, views, admin)
templates/   Páginas HTML
static/      CSS e arquivos estáticos
```

## Situação atual

Protótipo inicial (v0.1): estrutura, tela inicial e cadastro de médicos.
Próximos módulos: importação do MedPlus, regras de honorário, exportação OMIE
e emissão do repasse em PDF.
