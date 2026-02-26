# ICMS Normal – Imposto, Juros e Multa (código 113000)

Este documento descreve o **processo Normal** no portal DAR Web da **SEFAZ-PI**: fluxo manual, regras de negócio e como executar a automação.

> A estrutura de pastas e arquivos do projeto está no [README principal](../../README.md).

---

## Passo a passo do processo manual que a automação substitui

1. Acessar no navegador o portal DAR Web da SEFAZ-PI (`https://webas.sefaz.pi.gov.br/darweb/faces/views/index.xhtml`).
2. Clicar no **Menu ICMS** e, na tela de seleção, escolher **113000 - ICMS - IMPOSTO, JUROS E MULTA** e clicar em **Avançar**.
3. Informar a **Inscrição Estadual** no campo `#fieldInscricaoEstadual`, selecionar **Substituição Tributária: Não** (`cmbSubstituicao`) e clicar em **Avançar**.
4. No formulário "Cálculos do Imposto", preencher **período de referência** (mês/ano), **data de vencimento**, **data de pagamento** (ex.: dia 15 do mês de referência) e **valor principal** (coluna **NORMAL** da planilha).
5. Clicar em **Calcular Imposto**.
6. Clicar em **Avançar** para seguir ao pagamento e concluir o fluxo.
7. Fazer o **download do comprovante** da DAR gerada (quando aplicável) e **renomear o arquivo** com o padrão desejado (por I.E. e competência).
8. Repetir os passos 1 a 7 para cada I.E./filial (em geral a partir de uma planilha Excel com I.E., coluna **NORMAL** e período).

A automação replica esse fluxo por I.E., preenchendo os campos a partir do Excel. Envio final do formulário, download do comprovante e renomeação automática estão nos planos futuros.

---

## Vídeo do processo manual

Quando aplicável, vídeo do processo manual para referência futura:

- **[Pasta com vídeos do processo manual](https://drive.google.com/drive/folders/1lXeoZHH2d5bk2gXDSO87DxaoDQdNA748)**

---

## Vídeo da automação funcionando

Vídeo mostrando o bot em execução (carregando Excel, preenchendo o portal e rodando em lote):

- **A adicionar** (será incluído quando a automação completa estiver pronta).

---

## O que a automação faz

- Acessa o portal DAR Web da SEFAZ-PI (mesma URL do ATC e DIFAL).
- Para cada I.E. da planilha com valor na coluna **NORMAL**: clica no **Menu ICMS**, seleciona **113000 - ICMS - IMPOSTO, JUROS E MULTA**, clica em **Avançar**, preenche a **I.E.** em `#fieldInscricaoEstadual`, seleciona **Substituição Tributária: Não**, clica em **Avançar**, preenche **período de referência**, **data de vencimento**, **data de pagamento** (dia 15 do mês de referência) e **valor principal** (coluna NORMAL) e clica em **Calcular Imposto**.
- Permite carregar um Excel de filiais (coluna I.E. e coluna **NORMAL**), visualizar os dados extraídos (incluindo Valor ATC, NORMAL e DIF. ALIQUOTA), escolher quais I.E. executar e rodar o fluxo em lote (junto ou separado do ATC e DIFAL).
- Trata **datas passadas**: se a data de vencimento (dia 15 do mês de referência) já passou, a I.E. é ignorada e registrada com motivo do descarte.
- Em erro: registra log e salva screenshot na pasta de erros configurada.

**Planos futuros (a implementar):**

- Enviar o formulário final (clique em **Avançar** após "Calcular Imposto").
- Download automático do comprovante/guia DAR gerada.
- Renomeação automática do arquivo no padrão desejado (ex.: por I.E. e competência) e organização em pastas.
- Empacotar toda a automação em um executável (ex.: PyInstaller/Nuitka) para distribuição sem exigir instalação de Python.

---

## Regras de negócio

- **Normal** = recolhimento de ICMS no regime normal. O valor principal vem da coluna **NORMAL** da planilha de apuração.
- O processo usa o mesmo site e a mesma estrutura geral do DIFAL (código 113001), com as diferenças: código **113000** e valor principal vindo da coluna **NORMAL**.

---

## Dependências

- **Python** 3.13+
- **Pacotes Python** (na raiz do projeto):
  ```bash
  pip install -r requirements.txt
  ```
  Inclui: `playwright`, `openpyxl`, `customtkinter`, `python-dotenv`.
- **Navegador Playwright** (Chromium):
  ```bash
  python -m playwright install chromium
  ```

O projeto pode ser instalado em modo editável para expor o comando `icms_pi` e facilitar o uso.

---

## Variáveis de ambiente necessárias

Copie `.env.example` para `.env` e preencha conforme necessário:

| Variável | Obrigatória | Descrição |
|----------|-------------|-----------|
| `PASTA_SAIDA_RESULTADOS` | Não | Pasta para resultados (padrão: `resultados`). |
| `PASTA_CAPTURAS_DE_TELA_ERROS` | Não | Pasta para screenshots em caso de erro (padrão: `capturas_erros`). |

**Não commitar o arquivo `.env`.**

---

## Como executar

A aplicação é executada pela **interface gráfica** (comando central do projeto). Veja no [README principal](../../README.md) a seção **Como rodar**.

1. Na raiz do repositório: `pip install -e .` (ou `pip install -r requirements.txt`).
2. Abrir a GUI: `python -m icms_pi.gui_app` (ou `icms_pi`, se instalado em modo editável).
3. Selecionar um arquivo `.xlsx` com coluna de Inscrição Estadual (ex.: "INSC.ESTADUAL") e coluna **NORMAL**.
4. Usar **Ver dados extraídos** para conferir a extração.
5. Na interface: escolher o processo **ICMS Normal PI**, revisar I.E.s e executar o fluxo em lote (todas ou algumas I.E.).

**Planilha Excel:** Cabeçalho com coluna de I.E. (ex.: "INSC.ESTADUAL") e coluna **NORMAL** (nome exato da coluna); dados lidos até a linha de TOTAL ou até a primeira I.E. vazia. Período de referência extraído da área do título (mês/ano). I.E. aceita formato com pontos e traço; é normalizada para 9 dígitos. Suporte a planilhas no formato **APURAÇÃO DE ICMS PIAUÍ**.

**Logs:** Terminal em INFO+; arquivo em `logs/` em DEBUG+.

[← Voltar ao README principal](../../README.md)
