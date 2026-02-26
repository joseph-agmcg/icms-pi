# ICMS-PI — Automação SEFAZ-PI

Automação para processos de **ICMS** no portal DAR Web da SEFAZ-PI: **ATC** (Antecipado), **Normal** e **Difal**.

---

## Visão geral

- **Interface gráfica** (CustomTkinter): carregar planilha Excel, visualizar I.E.s e valores, selecionar quais executar e rodar em lote.
- **Planilha**: deve conter colunas de Inscrição Estadual e valores de **ATC**, **Normal** e **Difal**; o sistema extrai automaticamente o período e os dados para preenchimento da DAR.
- **Processos**: ATC (113011); Normal (113000, em construção); DIFAL (113001 – Imposto, Juros e Multa).

---

## Estrutura do projeto

| Pasta/Arquivo | Função |
|---------------|--------|
| **`pyproject.toml`** | Metadados do projeto (nome, versão, entrypoint `icms_pi`). |
| **`requirements.txt`** | Dependências para `pip install -r requirements.txt`. |
| **`.env`** | Variáveis sensíveis. **Não commitar.** |
| **`.env.example`** | Exemplo do `.env` sem valores reais. |
| **`logs/`** | Criada automaticamente; arquivos `.log` com timestamp. |
| **`src/icms_pi/`** | Comando central: GUI, extração Excel e logger; orquestra ATC, Normal e Difal. |
| **`src/atc/`** | Automação do ICMS Antecipado (código 113011). |
| **`src/normal/`** | ICMS Normal — em construção. |
| **`src/difal/`** | Automação do ICMS DIFAL (código 113001, valor DIF. ALIQUOTA). |

---

## Documentação por módulo

| Módulo | Descrição | README |
|--------|-----------|--------|
| **ATC** (113011) | Fluxo 113011 no DAR Web, passo a passo e automação. | [→ `src/atc/README.md`](src/atc/README.md) |
| **Normal** (113000) | Fluxo 113000 no DAR Web, passo a passo e automação (em construção). | [→ `src/normal/README.md`](src/normal/README.md) |
| **Difal** (113001) | Fluxo 113001 no DAR Web, passo a passo e automação (Imposto, Juros e Multa; valor DIF. ALIQUOTA). | [→ `src/difal/README.md`](src/difal/README.md) |

---

## Como rodar

1. Configurar `.env` a partir de `.env.example`.
2. Instalar dependências: `pip install -r requirements.txt` (ou `pip install -e .`).
3. Abrir a GUI:
   ```bash
   python -m icms_pi.gui_app
   ```
   Com o projeto instalado em modo editável (`pip install -e .`), também: `icms_pi`.
4. Na interface: selecionar a planilha Excel, revisar I.E.s e executar o processo desejado (ATC, Normal ou Difal).
