# ICMS-PI — Automação SEFAZ-PI

Automação para processos de **ICMS** no portal DAR Web da SEFAZ-PI: **ATC** (Antecipação Parcial), **Normal** (Imposto, Juros e Multa) e **DIFAL** (Diferencial de Alíquota — Imposto, Juros e Multa).

---

## Visão geral

- **Interface gráfica** (CustomTkinter): carregar planilha Excel, visualizar I.E. e valores, selecionar quais executar e rodar em lote.
- **Planilha**: deve conter colunas de Inscrição Estadual e valores de **ATC**, **NORMAL** e **DIF. ALIQUOTA**; o sistema extrai automaticamente o período e os dados para preenchimento da DAR.
- **Processos**: **ATC** (113011 – Antecipação Parcial); **Normal** (113000 – Imposto, Juros e Multa); **DIFAL** (113001 – Imposto, Juros e Multa, valor DIF. ALIQUOTA).

---

## Estrutura do projeto

| Pasta/Arquivo | Função |
|---------------|--------|
| **`pyproject.toml`** | Metadados do projeto (nome, versão, entrypoint `icms_pi`). |
| **`requirements.txt`** | Dependências para `pip install -r requirements.txt`. |
| **`.env`** | Variáveis sensíveis. **Não commitar.** |
| **`.env.example`** | Exemplo do `.env` sem valores reais. |
| **`logs/`** | Criada automaticamente; arquivos `.log` com timestamp. |
| **`src/icms_pi/`** | Comando central: GUI, extração Excel e logger; orquestra ATC, Normal e DIFAL. |
| **`src/atc/`** | Automação do ICMS Antecipado (código 113011, coluna ATC). |
| **`src/normal/`** | Automação do ICMS Normal (código 113000, coluna NORMAL). |
| **`src/difal/`** | Automação do ICMS DIFAL (código 113001, coluna DIF. ALIQUOTA). |

---

## Documentação por módulo

| Módulo | Descrição | README |
|--------|-----------|--------|
| **ATC** (113011) | ICMS Antecipação Parcial — fluxo no DAR Web, passo a passo e automação. | [→ `src/atc/README.md`](src/atc/README.md) |
| **Normal** (113000) | ICMS Normal — Imposto, Juros e Multa; fluxo no DAR Web, passo a passo e automação. | [→ `src/normal/README.md`](src/normal/README.md) |
| **DIFAL** (113001) | ICMS DIFAL — Imposto, Juros e Multa (valor DIF. ALIQUOTA); fluxo no DAR Web, passo a passo e automação. | [→ `src/difal/README.md`](src/difal/README.md) |

---

## Como rodar

1. Configurar `.env` a partir de `.env.example`.
2. Instalar dependências: `pip install -r requirements.txt` (ou `pip install -e .`).
3. Abrir a GUI:
   ```bash
   python -m icms_pi.gui_app
   ```
   Com o projeto instalado em modo editável (`pip install -e .`), também: `icms_pi`.
4. Na interface: selecionar a planilha Excel, revisar I.E. e executar o processo desejado (ATC, Normal ou DIFAL).
