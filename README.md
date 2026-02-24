# ICMS Antecipado - Piauí

Automação com **Playwright (Python assíncrono)** para o portal DAR Web da SEFAZ-PI. Interface gráfica para carregar Excel de filiais e executar o fluxo em lote por I.E. (ICMS Antecipado – código 113011).

---

## Passo a passo do processo manual que a automação substitui

1. Acessar no navegador o portal DAR Web da SEFAZ-PI (`https://webas.sefaz.pi.gov.br/darweb/faces/views/index.xhtml`).
2. Clicar no **Menu ICMS** e, na tela de seleção, escolher **113011 - ICMS – ANTECIPAÇÃO PARCIAL** e clicar em **Avançar**.
3. Informar a Inscrição Estadual (I.E.) no campo e clicar em **Avançar**.
4. No formulário "Cálculos do Imposto", preencher **período de referência** (mês/ano), **data de vencimento**, **data de pagamento** (ex.: dia 15 do mês de referência) e **valor principal** (coluna ATC da planilha).
5. Clicar em **Calcular Imposto**.
6. (Opcional/futuro) Clicar em **Avançar** para seguir ao pagamento e concluir o fluxo.
7. Fazer o **download do comprovante** da DAR gerada (quando aplicável) e **renomear o arquivo** com o padrão desejado (por I.E. e competência).
8. Repetir os passos 1 a 7 para cada I.E./filial (em geral a partir de uma planilha Excel com as I.E.s, coluna **ATC** e período).

A automação replica esse fluxo por I.E., preenchendo os campos a partir do Excel. Envio final do formulário, download do comprovante e renomeação automática estão nos planos futuros.

---

## Vídeo do processo manual

Quando aplicável, vídeo do processo manual para referência futura:

- **[Pasta com vídeos do processo manual (ICMS ATC PI, etc.)](https://drive.google.com/drive/folders/1lXeoZHH2d5bk2gXDSO87DxaoDQdNA748)**

---

## Vídeo da automação funcionando

Vídeo mostrando o bot em execução (carregando Excel, preenchendo o portal e rodando em lote):

- **A adicionar** (será incluído quando a automação completa estiver pronta).

---

## O que a automação faz

- Acessa o portal DAR Web da SEFAZ-PI (página inicial `index.xhtml`).
- Para cada I.E. da planilha: clica no **Menu ICMS**, seleciona **113011 - ICMS – ANTECIPAÇÃO PARCIAL**, clica em **Avançar**, preenche a **I.E.** e avança, preenche **período de referência**, **data de vencimento**, **data de pagamento** (dia 15 do mês de referência) e **valor principal** (coluna ATC) e clica em **Calcular Imposto**.
- Permite carregar um Excel de filiais (coluna I.E. = INSC.ESTADUAL e coluna **ATC**), visualizar os dados extraídos, escolher quais I.E. executar e rodar o fluxo em lote.
- Trata **datas passadas**: se a data de vencimento (dia 15 do mês de referência) já passou, a I.E. é ignorada e registrada com motivo do descarte.
- Em erro: registra log e salva screenshot na pasta de erros configurada.

**Planos futuros (a implementar):**

- Enviar o formulário final (clique em **Avançar** após "Calcular Imposto").
- Download automático do comprovante/guia DAR gerada.
- Renomeação automática do arquivo no padrão desejado (ex.: por I.E. e competência) e organização em pastas.
- Empacotar toda a automação em um executável (ex.: PyInstaller/Nuitka) para distribuição sem exigir instalação de Python.

---

## Regras de negócio

### Contexto

- **DAR** = Documento de Arrecadação: guia de pagamento usada para recolher tributos no Piauí (como o ICMS).
- **ICMS Antecipado (113011)** = receita de antecipação parcial do ICMS que empresas precisam declarar e pagar à SEFAZ-PI.
- Na prática, quem tem várias **filiais** (cada uma com uma Inscrição Estadual — I.E.) precisa gerar **uma DAR por filial**, por período, no portal da SEFAZ-PI. O valor a recolher vem da coluna **ATC** da planilha de apuração.

### O que o bot faz na prática

- Você já tem uma **planilha Excel** de apuração de ICMS (típica de "APURAÇÃO DE ICMS PIAUÍ") com as filiais, o período (mês/ano) e o **valor ATC** (valor principal) de cada uma.
- O bot **lê essa planilha** e, para cada filial que tem valor ATC a pagar:
  - Acessa o portal DAR Web da SEFAZ-PI.
  - Seleciona o código **113011 - ICMS – ANTECIPAÇÃO PARCIAL** e avança.
  - Informa a I.E. da filial e avança.
  - Preenche **período de referência** (mês/ano da planilha), **data de vencimento**, **data de pagamento** (dia 15 do mês de referência) e **valor principal** (coluna ATC).
  - Clica em **Calcular Imposto**.
- I.E.s cuja data de vencimento já está no passado são **puladas** e listadas com motivo claro (portal não aceita datas passadas ou regra de negócio).
- Ou seja: o bot **replica no site** o que está na planilha, filial a filial, sem você precisar digitar cada I.E. e cada valor manualmente.

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

O projeto pode ser instalado em modo editável para expor um comando único (`icms-atc-pi`) e facilitar o uso.

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

### Interface gráfica (Excel + lote)

1. Instale o projeto em modo editável (na raiz do repositório):
   ```powershell
   pip install -e .
   ```
2. Depois, execute a GUI com o comando:
   ```powershell
   icms-atc-pi
   ```
3. Selecione um arquivo `.xlsx` com coluna de Inscrição Estadual (ex.: "INSC.ESTADUAL") e coluna **ATC**.
4. Use **Ver dados extraídos** para conferir a extração.
5. Use as opções da interface para executar o fluxo em lote (todas ou algumas I.E.).

### Só automação (linha de comando)

Se preferir rodar sem a GUI ou sem instalar o comando:

```powershell
$env:PYTHONPATH="src"
python -m sefaz_pi.main --excel "C:\caminho\para\icms.xlsx"
```

Ou em modo headless (sem abrir janela do navegador):

```powershell
$env:PYTHONPATH="src"
python -m sefaz_pi.main --excel "C:\caminho\para\icms.xlsx" --headless
```

**Planilha Excel:** Cabeçalho com coluna de I.E. (ex.: "INSC.ESTADUAL") e coluna **ATC** (valor principal); dados lidos até a linha de TOTAL ou até a primeira I.E. vazia. Período de referência extraído da área do título (mês/ano). I.E. aceita formato com pontos e traço; é normalizada para 9 dígitos. Suporte a planilhas no formato **APURAÇÃO DE ICMS PIAUÍ**.

**Logs:** Terminal em INFO+; arquivo em `logs/` em DEBUG+.

---

## Estrutura do projeto

| Pasta/Arquivo | Função |
|---------------|--------|
| **`pyproject.toml`** | Metadados do projeto (nome, versão, entrypoint `icms-atc-pi`). |
| **`requirements.txt`** | Dependências para `pip install -r requirements.txt`. |
| **`.env`** | Variáveis sensíveis. **Não commitar.** |
| **`.env.example`** | Exemplo do `.env` sem valores reais. |
| **`logs/`** | Criada automaticamente; arquivos `.log` com timestamp. |
| **`src/sefaz_pi/main.py`** | Ponto de entrada CLI; modo GUI/CLI com `--excel`; suporta `--headless`. |
| **`src/sefaz_pi/gui_app.py`** | Interface gráfica (CustomTkinter): upload Excel, extração I.E./ATC, execução em lote. |
| **`src/sefaz_pi/automacao_sefaz_pi.py`** | Classe principal `AutomacaoAntecipacaoParcialPI`; orquestra o fluxo por I.E. |
| **`src/sefaz_pi/configuracoes.py`** | Carrega `.env`; URLs, seletores, timeouts, pastas. |
| **`src/sefaz_pi/excel_filiais.py`** | Extração de dados do Excel (formato ICMS/filiais, coluna ATC). |
| **`src/sefaz_pi/logger.py`** | Logging: terminal INFO+, arquivo DEBUG+ em `logs/`. |
| **`src/sefaz_pi/navegacao/`** | Ações atômicas na página (aguardar, preencher campos, screenshot em erro). |

