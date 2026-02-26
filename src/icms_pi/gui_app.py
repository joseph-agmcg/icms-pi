"""Interface desktop com CustomTkinter para o sistema ICMS-PI (ATC, Normal, DIFAL). Exibe Valor ATC e DIF. ALIQUOTA."""

import asyncio
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from icms_pi import configuracoes
from icms_pi.excel_filiais import (
    extrair_todos_os_dados,
    obter_dados_para_dae,
    obter_ies_dos_dados,
    _obter_chave_ie,
)
from icms_pi.logger import configurar_logger_da_aplicacao
from atc.automacao_sefaz_pi import AutomacaoAntecipacaoParcialPI, _valor_atc_invalido
from difal.automacao_sefaz_pi import AutomacaoDifalPI, _valor_difal_invalido
from normal.automacao_sefaz_pi import AutomacaoNormalPI, _valor_normal_invalido

logger = configurar_logger_da_aplicacao(__name__)


PROCESSOS_ICMS_PI: list[tuple[str, str]] = [
    ("antecipado", "ICMS Antecipado PI"),
    ("normal", "ICMS Normal PI"),
    ("difal", "ICMS Difal PI"),
]

_PROCESSO_POR_ID: dict[str, str] = {pid: nome for pid, nome in PROCESSOS_ICMS_PI}

# Rótulos curtos para o seletor de lista (mesmo tamanho visual)
_PID_PARA_LABEL_CURTO: dict[str, str] = {
    "antecipado": "ATC",
    "normal": "Normal",
    "difal": "DIFAL",
}
_LABEL_CURTO_PARA_PID: dict[str, str] = {v: k for k, v in _PID_PARA_LABEL_CURTO.items()}


# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------

def _nome_processo_legivel(processo_id: str) -> str:
    return _PROCESSO_POR_ID.get(processo_id, processo_id)


def _ie_para_exibicao(ie: str | None) -> str:
    if ie is None:
        return ""
    s = str(ie).strip()
    return s.replace(".", "").replace("-", "").replace("/", "")


def _formato_intervalo_ms(ms: int) -> str:
    if ms >= 60_000:
        return f"{ms // 60_000} min"
    if ms >= 1_000:
        return f"{ms // 1_000} s"
    return f"{ms} ms"


def _item_executavel_para_processo(item: dict[str, object], processo_id: str) -> bool:
    """Retorna True se o item tem valor válido para o processo (antecipado=ATC, normal=NORMAL, difal=DIF. ALIQUOTA)."""
    if processo_id == "antecipado":
        return not _valor_atc_invalido(item.get("valor_atc"))
    if processo_id == "normal":
        return not _valor_normal_invalido(item.get("valor_normal"))
    if processo_id == "difal":
        return not _valor_difal_invalido(item.get("valor_difal"))
    return False


def _contar_executaveis_ignoradas(
    lista_dados: list[dict[str, object]],
    processos_ids: list[str] | None = None,
) -> tuple[int, int]:
    """Conta executáveis (com valor válido para pelo menos um dos processos) e ignoradas."""
    if not processos_ids:
        processos_ids = ["antecipado"]
    executaveis = sum(
        1 for item in lista_dados
        if any(_item_executavel_para_processo(item, pid) for pid in processos_ids)
    )
    return executaveis, len(lista_dados) - executaveis


# ---------------------------------------------------------------------------
# Execução em background (thread separada)
# ---------------------------------------------------------------------------

def _executar_lote_em_background(
    lista_dados: list[dict[str, object]],
    processos_ids: list[str],
    headless: bool,
    result_callback=None,
    lista_por_processo: dict[str, list[dict[str, object]]] | None = None,
) -> None:
    if not processos_ids:
        return

    # Se lista_por_processo foi passada (modo 3 listas), usa ela; senão filtra por valor
    if lista_por_processo is None:
        lista_por_processo = {}
        for pid in processos_ids:
            if pid == "antecipado":
                lista_por_processo[pid] = [
                    item for item in lista_dados
                    if not _valor_atc_invalido(item.get("valor_atc"))
                ]
            elif pid == "normal":
                lista_por_processo[pid] = [
                    item for item in lista_dados
                    if not _valor_normal_invalido(item.get("valor_normal"))
                ]
            elif pid == "difal":
                lista_por_processo[pid] = [
                    item for item in lista_dados
                    if not _valor_difal_invalido(item.get("valor_difal"))
                ]

    total = sum(len(lista_por_processo.get(pid, [])) for pid in processos_ids)
    intervalo_txt = _formato_intervalo_ms(configuracoes.INTERVALO_ENTRE_EXECUCOES_MS)
    logger.info(
        "Iniciando lote: processos=%s, total itens=%s, intervalo=%s, headless=%s",
        processos_ids, total, intervalo_txt, headless,
    )

    def _worker() -> None:
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)

            ies_ok: list[str] = []
            ies_erro: list[tuple[str, str]] = []

            if "antecipado" in processos_ids:
                lista_atc = lista_por_processo.get("antecipado", [])
                if lista_atc:
                    async def _rodar_antecipado() -> None:
                        automacao = AutomacaoAntecipacaoParcialPI(headless=headless)
                        ok, erro = await automacao.executar_fluxo_por_ie_pi(lista_atc)
                        ies_ok.extend(ok)
                        ies_erro.extend(erro)
                    loop.run_until_complete(_rodar_antecipado())

            if "normal" in processos_ids:
                lista_normal = lista_por_processo.get("normal", [])
                if lista_normal:
                    async def _rodar_normal() -> None:
                        automacao = AutomacaoNormalPI(headless=headless)
                        ok, erro = await automacao.executar_fluxo_por_ie_pi(lista_normal)
                        ies_ok.extend(ok)
                        ies_erro.extend(erro)
                    loop.run_until_complete(_rodar_normal())

            if "difal" in processos_ids:
                lista_difal = lista_por_processo.get("difal", [])
                if lista_difal:
                    async def _rodar_difal() -> None:
                        automacao = AutomacaoDifalPI(headless=headless)
                        ok, erro = await automacao.executar_fluxo_por_ie_pi(lista_difal)
                        ies_ok.extend(ok)
                        ies_erro.extend(erro)
                    loop.run_until_complete(_rodar_difal())

            if result_callback is not None:
                result_callback(ies_ok, ies_erro)
        finally:
            try:
                loop.close()
            except Exception:
                logger.exception("Falha ao fechar event loop da GUI.")

    threading.Thread(target=_worker, daemon=True).start()


# ---------------------------------------------------------------------------
# Janela de visualização dos dados extraídos
# ---------------------------------------------------------------------------

def _mostrar_janela_dados_extraidos(
    parent: ctk.CTk,
    caminho: Path,
    dados_extraidos: list[dict[str, object]],
    nomes_colunas: list[str],
    nome_para_indice: dict[str, int],
) -> None:
    if not dados_extraidos or not nomes_colunas:
        messagebox.showinfo("Dados", "Nenhum dado extraído para exibir.")
        return

    janela = ctk.CTkToplevel(parent)
    janela.title("Dados extraídos do Excel")
    janela.geometry("920x520")
    janela.minsize(500, 340)

    ctk.CTkLabel(
        janela,
        text=(
            f"Arquivo: {caminho.name}  |  "
            f"{len(dados_extraidos)} linhas  |  "
            f"{len(nomes_colunas)} colunas"
        ),
        font=ctk.CTkFont(size=12),
    ).pack(pady=8)

    texto = ctk.CTkTextbox(
        janela, font=ctk.CTkFont(family="Consolas", size=12), wrap="none",
    )
    texto.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    colunas = nomes_colunas
    larguras = [max(len(str(col)), 4) for col in colunas]
    for i, col in enumerate(colunas):
        larguras[i] = max(
            larguras[i],
            max(
                (len(str(linha.get(col, "") or "")) for linha in dados_extraidos[:100]),
                default=0,
            ),
        )
        larguras[i] = min(larguras[i], 30)

    def _cel(s: object, w: int) -> str:
        ss = "" if s is None else str(s).strip()
        return (ss[: w - 2] + "..") if len(ss) > w else ss.ljust(w)

    ie_key = _obter_chave_ie(nome_para_indice)

    linha_cab = " | ".join(_cel(col, larguras[i]) for i, col in enumerate(colunas))
    sep = "-+-".join("-" * w for w in larguras)
    linhas_txt = [linha_cab, sep]
    for linha in dados_extraidos:
        def _valor_celula(col: str, idx: int, _linha=linha) -> str:
            v = _linha.get(col)
            if ie_key and col == ie_key and v is not None:
                return _cel(_ie_para_exibicao(str(v)), larguras[idx])
            return _cel(v, larguras[idx])
        linhas_txt.append(
            " | ".join(_valor_celula(col, i) for i, col in enumerate(colunas))
        )

    texto.insert("1.0", "\n".join(linhas_txt))
    texto.configure(state="disabled")


# ---------------------------------------------------------------------------
# Janela principal
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    """Janela principal da aplicação ICMS-PI (GUI unificada)."""

    _MODO_TABELA = "tabela"
    _MODO_SELECAO = "selecao"

    def __init__(self) -> None:
        super().__init__()
        self.title("ICMS-PI — Automação SEFAZ-PI")
        self.geometry("1060x700")
        self.minsize(820, 580)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self._caminho_excel: Path | None = None
        self._dados_extraidos: list[dict[str, object]] = []
        self._lista_dados: list[dict[str, object]] = []
        self._nomes_colunas: list[str] = []
        self._nome_para_indice: dict[str, int] = {}
        self._mes_ref: int = 0
        self._ano_ref: int = 0
        self._executando = False

        self._modo_ies = self._MODO_TABELA
        self._vars_selecao: list[tuple[dict[str, object], ctk.BooleanVar]] = []
        # Modo 3 listas: qual processo está em edição e seleção por processo
        self._processo_selecao_visivel: str = "antecipado"
        self._selecao_por_processo: dict[str, list[tuple[dict[str, object], ctk.BooleanVar]]] = {}

        self._construir_layout()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
    def _construir_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self._criar_cabecalho()
        self._criar_painel_arquivo()
        self._criar_area_central()
        self._criar_barra_status()

    def _criar_cabecalho(self) -> None:
        frame = ctk.CTkFrame(self, corner_radius=0, height=52)
        frame.grid(row=0, column=0, sticky="ew")
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            frame, text="ICMS-PI",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).grid(row=0, column=0, padx=(20, 8), pady=10)

        ctk.CTkLabel(
            frame,
            text="Automação SEFAZ-PI  •  ATC / Normal / Difal",
            font=ctk.CTkFont(size=12), text_color="gray",
        ).grid(row=0, column=1, sticky="w", padx=4)

    def _criar_painel_arquivo(self) -> None:
        frame = ctk.CTkFrame(self, corner_radius=8)
        frame.grid(row=1, column=0, sticky="ew", padx=14, pady=(10, 0))
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            frame, text="Planilha Excel:",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, padx=(14, 8), pady=12)

        self._label_arquivo = ctk.CTkLabel(
            frame, text="Nenhum arquivo selecionado.",
            font=ctk.CTkFont(size=12), text_color="gray", anchor="w",
        )
        self._label_arquivo.grid(row=0, column=1, sticky="ew", padx=4)

        self._btn_ver_dados = ctk.CTkButton(
            frame, text="Ver dados extraídos", width=160,
            command=self._ao_ver_dados, state="disabled",
        )
        self._btn_ver_dados.grid(row=0, column=2, padx=(4, 4), pady=12)
        self._btn_ver_dados.grid_remove()

        self._btn_abrir = ctk.CTkButton(
            frame, text="Selecionar arquivo…", width=140,
            command=self._selecionar_arquivo,
        )
        self._btn_abrir.grid(row=0, column=3, padx=(4, 14), pady=12)

    def _criar_area_central(self) -> None:
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.grid(row=2, column=0, sticky="nsew", padx=14, pady=10)
        container.grid_columnconfigure(0, weight=3, uniform="col")
        container.grid_columnconfigure(1, weight=2, uniform="col")
        container.grid_rowconfigure(0, weight=1)

        col_esq = ctk.CTkFrame(container, fg_color="transparent")
        col_esq.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        col_esq.grid_rowconfigure(1, weight=1)
        col_esq.grid_columnconfigure(0, weight=1)

        self._criar_resumo(col_esq)
        self._criar_painel_ies(col_esq)

        col_dir = ctk.CTkFrame(container, fg_color="transparent")
        col_dir.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        col_dir.grid_rowconfigure(1, weight=1)
        col_dir.grid_columnconfigure(0, weight=1)

        self._criar_painel_processos(col_dir)
        self._criar_painel_log(col_dir)

    # --- Cards de resumo ---
    def _criar_resumo(self, parent: ctk.CTkFrame) -> None:
        frame = ctk.CTkFrame(parent, corner_radius=8)
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self._lbl_periodo = self._card_info(frame, "Período", "—", 0)
        self._lbl_total = self._card_info(frame, "Total IEs", "0", 1)
        self._lbl_exec = self._card_info(frame, "Executáveis", "0", 2)
        self._lbl_ignor = self._card_info(frame, "Ignoradas", "0", 3)

    def _card_info(
        self, parent: ctk.CTkFrame, titulo: str, valor: str, col: int,
    ) -> ctk.CTkLabel:
        wrapper = ctk.CTkFrame(parent, corner_radius=6)
        wrapper.grid(row=0, column=col, padx=6, pady=8, sticky="ew")
        wrapper.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            wrapper, text=titulo, font=ctk.CTkFont(size=11), text_color="gray",
        ).grid(row=0, column=0, pady=(6, 0))

        lbl = ctk.CTkLabel(
            wrapper, text=valor, font=ctk.CTkFont(size=17, weight="bold"),
        )
        lbl.grid(row=1, column=0, pady=(0, 6))
        return lbl

    # --- Painel de IEs (tabela + seleção alternável) ---
    def _criar_painel_ies(self, parent: ctk.CTkFrame) -> None:
        self._frame_ies = ctk.CTkFrame(parent, corner_radius=8)
        self._frame_ies.grid(row=1, column=0, sticky="nsew")
        self._frame_ies.grid_rowconfigure(1, weight=1)
        self._frame_ies.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self._frame_ies, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 2))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header, text="Inscrições Estaduais",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        self._btn_alternar_modo = ctk.CTkButton(
            header, text="Selecionar IEs", width=130, height=28,
            font=ctk.CTkFont(size=11),
            command=self._alternar_modo_ies, state="disabled",
        )
        self._btn_alternar_modo.grid(row=0, column=1, sticky="e")

        _GRID_CONTEUDO = dict(row=1, column=0, sticky="nsew", padx=8, pady=(4, 8))

        # Modo tabela (textbox)
        self._textbox_ies = ctk.CTkTextbox(
            self._frame_ies, font=ctk.CTkFont(family="Consolas", size=12),
            state="disabled", wrap="none",
        )
        self._textbox_ies.grid(**_GRID_CONTEUDO)

        # Modo seleção: seletor de processo + uma lista por vez (3 listas)
        self._container_selecao = ctk.CTkFrame(self._frame_ies, fg_color="transparent")
        self._container_selecao.grid_columnconfigure(0, weight=1)
        self._container_selecao.grid_rowconfigure(1, weight=1)
        self._container_selecao.grid(**_GRID_CONTEUDO)
        self._container_selecao.grid_remove()

        self._grid_conteudo = _GRID_CONTEUDO

        # Linha 0: seletor "Lista do processo:" — só processos marcados no painel Processos; rótulos curtos
        frame_selector = ctk.CTkFrame(self._container_selecao, fg_color="transparent")
        frame_selector.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
        frame_selector.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            frame_selector, text="Lista do processo:",
            font=ctk.CTkFont(size=11), text_color="gray",
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._var_processo_selecao = ctk.StringVar(value="ATC")
        self._frame_seletor_botoes = ctk.CTkFrame(frame_selector, fg_color="transparent")
        self._frame_seletor_botoes.grid(row=0, column=1, sticky="w")
        self._segmented_processos: ctk.CTkSegmentedButton | None = None

        # Linha 1: container das 3 listas (só uma visível por vez)
        self._container_listas = ctk.CTkFrame(self._container_selecao, fg_color="transparent")
        self._container_listas.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        self._container_listas.grid_columnconfigure(0, weight=1)
        self._container_listas.grid_rowconfigure(0, weight=1)
        # Frames por processo (preenchidos ao entrar no modo seleção)
        self._frame_lista_por_processo: dict[str, ctk.CTkFrame] = {}
        self._frame_scroll_por_processo: dict[str, ctk.CTkScrollableFrame] = {}
        self._barra_por_processo: dict[str, ctk.CTkFrame] = {}
        self._lbl_contador_por_processo: dict[str, ctk.CTkLabel] = {}

    # --- Painel de processos + executar ---
    def _criar_painel_processos(self, parent: ctk.CTkFrame) -> None:
        frame = ctk.CTkFrame(parent, corner_radius=8)
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        frame.grid_columnconfigure(0, weight=1)

        row_idx = 0

        ctk.CTkLabel(
            frame, text="Processos",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=row_idx, column=0, sticky="w", padx=12, pady=(8, 2))
        row_idx += 1

        self._vars_processos: dict[str, ctk.BooleanVar] = {}
        for pid, nome in PROCESSOS_ICMS_PI:
            var = ctk.BooleanVar(value=(pid == "antecipado"))
            self._vars_processos[pid] = var
            var.trace_add("write", lambda *_: self._atualizar_resumo_e_tabela_processos())
            cb = ctk.CTkCheckBox(frame, text=nome, variable=var, font=ctk.CTkFont(size=12))
            cb.grid(row=row_idx, column=0, sticky="w", padx=22, pady=2)
            row_idx += 1

        ctk.CTkFrame(frame, height=1).grid(
            row=row_idx, column=0, sticky="ew", padx=12, pady=6,
        )
        row_idx += 1

        self._var_headless = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            frame, text="Headless (navegador oculto)",
            variable=self._var_headless, font=ctk.CTkFont(size=12),
        ).grid(row=row_idx, column=0, sticky="w", padx=22, pady=(0, 4))
        row_idx += 1

        ctk.CTkFrame(frame, height=1).grid(
            row=row_idx, column=0, sticky="ew", padx=12, pady=6,
        )
        row_idx += 1

        self._btn_executar = ctk.CTkButton(
            frame, text="▶  Executar", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="green", hover_color="darkgreen",
            command=self._ao_executar, state="disabled",
        )
        self._btn_executar.grid(row=row_idx, column=0, padx=12, pady=(4, 12), sticky="ew")

    # --- Painel de log ---
    def _criar_painel_log(self, parent: ctk.CTkFrame) -> None:
        frame = ctk.CTkFrame(parent, corner_radius=8)
        frame.grid(row=1, column=0, sticky="nsew")
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            frame, text="Log de execução",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 2))

        self._textbox_log = ctk.CTkTextbox(
            frame, font=ctk.CTkFont(family="Consolas", size=11),
            state="disabled", wrap="word",
        )
        self._textbox_log.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))

    # --- Barra de status ---
    def _criar_barra_status(self) -> None:
        frame = ctk.CTkFrame(self, corner_radius=0, height=28)
        frame.grid(row=3, column=0, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)

        self._lbl_status = ctk.CTkLabel(
            frame, text="Pronto", font=ctk.CTkFont(size=11),
            text_color="gray", anchor="w",
        )
        self._lbl_status.grid(row=0, column=0, sticky="ew", padx=14, pady=3)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _log(self, msg: str) -> None:
        self._textbox_log.configure(state="normal")
        self._textbox_log.insert("end", msg + "\n")
        self._textbox_log.see("end")
        self._textbox_log.configure(state="disabled")

    def _status(self, msg: str) -> None:
        self._lbl_status.configure(text=msg)

    def _habilitar_botoes(self, habilitado: bool = True) -> None:
        estado = "normal" if habilitado else "disabled"
        self._btn_abrir.configure(state=estado)
        self._btn_executar.configure(state=estado)
        self._btn_alternar_modo.configure(state=estado)
        if habilitado and self._dados_extraidos:
            self._btn_ver_dados.configure(state="normal")
        else:
            self._btn_ver_dados.configure(state=estado)

    # ------------------------------------------------------------------
    # Alternar modo tabela / seleção
    # ------------------------------------------------------------------
    def _alternar_modo_ies(self) -> None:
        if self._modo_ies == self._MODO_TABELA:
            self._mostrar_modo_selecao()
        else:
            self._mostrar_modo_tabela()

    def _mostrar_modo_tabela(self) -> None:
        self._modo_ies = self._MODO_TABELA
        self._container_selecao.grid_remove()
        self._textbox_ies.grid(**self._grid_conteudo)
        self._btn_alternar_modo.configure(text="Selecionar IEs")

    def _mostrar_modo_selecao(self) -> None:
        self._modo_ies = self._MODO_SELECAO
        self._textbox_ies.grid_remove()
        self._container_selecao.grid(**self._grid_conteudo)
        self._btn_alternar_modo.configure(text="Ver tabela")
        self._atualizar_seletor_processos_selecao()
        self._construir_listas_por_processo()
        self._mostrar_lista_do_processo(self._processo_selecao_visivel)

    def _atualizar_seletor_processos_selecao(self) -> None:
        """Mostra no seletor só os processos marcados no painel Processos; rótulos curtos (ATC, Normal, DIFAL)."""
        processos_ativos = [p for p, var in self._vars_processos.items() if var.get()]
        if not processos_ativos:
            processos_ativos = ["antecipado", "normal", "difal"]
        valores = [_PID_PARA_LABEL_CURTO[pid] for pid in ("antecipado", "normal", "difal") if pid in processos_ativos]
        if not valores:
            valores = ["ATC", "Normal", "DIFAL"]
        for w in self._frame_seletor_botoes.winfo_children():
            w.destroy()
        self._segmented_processos = None
        if valores:
            atual_label = _PID_PARA_LABEL_CURTO.get(self._processo_selecao_visivel, "ATC")
            if atual_label not in valores:
                atual_label = valores[0]
                self._processo_selecao_visivel = _LABEL_CURTO_PARA_PID.get(atual_label, "antecipado")
            self._var_processo_selecao.set(atual_label)
            self._segmented_processos = ctk.CTkSegmentedButton(
                self._frame_seletor_botoes,
                values=valores,
                variable=self._var_processo_selecao,
                command=self._ao_trocar_processo_selecao,
            )
            self._segmented_processos.pack(side="left")

    def _ao_trocar_processo_selecao(self, valor_segmented: str) -> None:
        pid = _LABEL_CURTO_PARA_PID.get(valor_segmented, "antecipado")
        self._processo_selecao_visivel = pid
        self._mostrar_lista_do_processo(pid)

    def _mostrar_lista_do_processo(self, pid: str) -> None:
        for p, frame in self._frame_lista_por_processo.items():
            frame.grid_remove()
        if pid in self._frame_lista_por_processo:
            self._frame_lista_por_processo[pid].grid(row=0, column=0, sticky="nsew")
        self._atualizar_contador_selecao()

    def _construir_listas_por_processo(self) -> None:
        """Cria as 3 listas (ATC, Normal, DIFAL) e preenche cada uma com IEs do processo."""
        self._selecao_por_processo.clear()
        for p, frame in self._frame_lista_por_processo.items():
            frame.destroy()
        self._frame_lista_por_processo.clear()
        self._frame_scroll_por_processo.clear()
        self._barra_por_processo.clear()
        self._lbl_contador_por_processo.clear()

        for pid in ("antecipado", "normal", "difal"):
            itens_processo = self._itens_executaveis_para_processo(pid)
            frame_lista = ctk.CTkFrame(self._container_listas, fg_color="transparent")
            frame_lista.grid(row=0, column=0, sticky="nsew")
            frame_lista.grid_columnconfigure(0, weight=1)
            frame_lista.grid_rowconfigure(0, weight=1)
            self._frame_lista_por_processo[pid] = frame_lista

            scroll = ctk.CTkScrollableFrame(frame_lista)
            scroll.grid(row=0, column=0, sticky="nsew", padx=4, pady=(4, 0))
            self._frame_scroll_por_processo[pid] = scroll

            barra = ctk.CTkFrame(frame_lista, fg_color="transparent")
            barra.grid(row=1, column=0, sticky="ew", padx=4, pady=(4, 6))
            self._barra_por_processo[pid] = barra

            ctk.CTkButton(
                barra, text="Selecionar todas", width=140, height=28,
                font=ctk.CTkFont(size=11),
                command=lambda p=pid: self._marcar_todas_ies_processo(p),
            ).pack(side="left", padx=(0, 6))
            ctk.CTkButton(
                barra, text="Desmarcar todas", width=130, height=28,
                font=ctk.CTkFont(size=11),
                command=lambda p=pid: self._desmarcar_todas_ies_processo(p),
            ).pack(side="left", padx=(0, 6))
            lbl_cont = ctk.CTkLabel(
                barra, text="Selecionadas: 0",
                font=ctk.CTkFont(size=11), text_color="gray",
            )
            lbl_cont.pack(side="left", padx=8)
            self._lbl_contador_por_processo[pid] = lbl_cont

            # Valor a exibir: só o do processo atual (ATC, NORMAL ou DIF. ALIQUOTA)
            chave_valor = {"antecipado": "valor_atc", "normal": "valor_normal", "difal": "valor_difal"}
            valor_key = chave_valor.get(pid, "valor_atc")

            vars_list: list[tuple[dict[str, object], ctk.BooleanVar]] = []
            for item in itens_processo:
                ie = _ie_para_exibicao(item.get("ie") or item.get("ie_digitos") or "")
                val = item.get(valor_key)
                valor_txt = (
                    f"{val:.2f}"
                    if isinstance(val, (int, float))
                    else (str(val) if val is not None else "—")
                )
                var = ctk.BooleanVar(value=True)
                vars_list.append((item, var))
                row = ctk.CTkFrame(scroll, fg_color="transparent")
                row.pack(fill="x", pady=1)
                cb = ctk.CTkCheckBox(row, text="", variable=var, width=28)
                cb.pack(side="left", padx=(0, 6), pady=3)
                var.trace_add("write", lambda *_, proc=pid: self._atualizar_contador_selecao())
                ctk.CTkLabel(
                    row, text=ie,
                    font=ctk.CTkFont(family="Consolas", size=12), width=110,
                ).pack(side="left", padx=4, pady=3)
                ctk.CTkLabel(
                    row, text=valor_txt,
                    font=ctk.CTkFont(family="Consolas", size=12), width=80,
                ).pack(side="left", padx=4, pady=3)
            self._selecao_por_processo[pid] = vars_list
            n = sum(1 for _, v in vars_list if v.get())
            self._lbl_contador_por_processo[pid].configure(text=f"Selecionadas: {n}")

        for f in self._frame_lista_por_processo.values():
            f.grid_remove()

    def _itens_executaveis_para_processo(self, pid: str) -> list[dict[str, object]]:
        """Retorna os itens que têm valor/critério para o processo (para montar a lista)."""
        if pid == "antecipado":
            return [
                item for item in self._lista_dados
                if not _valor_atc_invalido(item.get("valor_atc"))
            ]
        if pid == "normal":
            return [
                item for item in self._lista_dados
                if not _valor_normal_invalido(item.get("valor_normal"))
            ]
        if pid == "difal":
            return [
                item for item in self._lista_dados
                if not _valor_difal_invalido(item.get("valor_difal"))
            ]
        return []

    def _marcar_todas_ies_processo(self, pid: str) -> None:
        for _, var in self._selecao_por_processo.get(pid, []):
            var.set(True)

    def _desmarcar_todas_ies_processo(self, pid: str) -> None:
        for _, var in self._selecao_por_processo.get(pid, []):
            var.set(False)

    def _atualizar_contador_selecao(self) -> None:
        pid = self._processo_selecao_visivel
        n = sum(
            1 for _, var in self._selecao_por_processo.get(pid, [])
            if var.get()
        )
        if pid in self._lbl_contador_por_processo:
            self._lbl_contador_por_processo[pid].configure(text=f"Selecionadas: {n}")

    def _atualizar_resumo_e_tabela_processos(self) -> None:
        """Atualiza contagem Executáveis/Ignoradas e tabela de IEs quando a seleção de processos muda."""
        if not self._lista_dados:
            return
        processos_ativos = [p for p, var in self._vars_processos.items() if var.get()]
        processos_ativos = processos_ativos or ["antecipado", "normal", "difal"]
        qtd_exec, qtd_ign = _contar_executaveis_ignoradas(
            self._lista_dados, processos_ativos
        )
        self._lbl_exec.configure(text=str(qtd_exec))
        self._lbl_ignor.configure(text=str(qtd_ign))
        self._btn_executar.configure(state="normal" if qtd_exec > 0 else "disabled")
        self._preencher_tabela_ies()
        if self._modo_ies == self._MODO_SELECAO:
            self._atualizar_seletor_processos_selecao()
            self._mostrar_lista_do_processo(self._processo_selecao_visivel)

    # ------------------------------------------------------------------
    # Ações
    # ------------------------------------------------------------------
    def _selecionar_arquivo(self) -> None:
        if self._caminho_excel is not None:
            trocar = messagebox.askyesno(
                "Trocar planilha",
                "Já existe uma planilha carregada.\n"
                "Deseja selecionar outro arquivo e substituir os dados atuais?",
            )
            if not trocar:
                return

        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not caminho:
            return

        self._caminho_excel = Path(caminho)
        logger.info("Arquivo selecionado: %s", self._caminho_excel.resolve())
        self._label_arquivo.configure(text=self._caminho_excel.name, text_color="white")
        self._carregar_planilha()
        self._btn_abrir.configure(text="Trocar arquivo…")

    def _carregar_planilha(self) -> None:
        if self._caminho_excel is None:
            return
        self._status("Carregando planilha…")
        self._log(f"Abrindo: {self._caminho_excel.name}")

        try:
            self._dados_extraidos, self._nome_para_indice, self._mes_ref, self._ano_ref = (
                extrair_todos_os_dados(self._caminho_excel)
            )
            self._nomes_colunas = [
                k for k, _ in sorted(self._nome_para_indice.items(), key=lambda x: x[1])
            ]
            self._lista_dados = obter_dados_para_dae(
                self._dados_extraidos, self._nome_para_indice,
                self._mes_ref, self._ano_ref,
            )
        except Exception as e:
            messagebox.showerror("Erro ao carregar planilha", str(e))
            self._log(f"ERRO: {e}")
            self._status("Falha ao carregar planilha")
            return

        processos_ativos = [pid for pid, var in self._vars_processos.items() if var.get()]
        qtd_exec, qtd_ign = _contar_executaveis_ignoradas(
            self._lista_dados,
            processos_ativos if processos_ativos else ["antecipado", "normal", "difal"],
        )
        total = len(self._lista_dados)

        self._lbl_periodo.configure(text=f"{self._mes_ref:02d}/{self._ano_ref}")
        self._lbl_total.configure(text=str(total))
        self._lbl_exec.configure(text=str(qtd_exec))
        self._lbl_ignor.configure(text=str(qtd_ign))

        self._preencher_tabela_ies()
        if self._modo_ies == self._MODO_SELECAO:
            self._mostrar_modo_tabela()

        self._btn_ver_dados.configure(state="normal")
        self._btn_ver_dados.grid()
        self._btn_alternar_modo.configure(state="normal" if total > 0 else "disabled")
        self._btn_executar.configure(state="normal" if qtd_exec > 0 else "disabled")

        self._log(
            f"Planilha carregada: {total} registros, "
            f"{qtd_exec} executáveis, período {self._mes_ref:02d}/{self._ano_ref}"
        )
        self._status("Planilha carregada — pronto para executar")
        logger.info(
            "Extração: %d linhas, %d colunas, %d IEs, período %02d/%04d",
            len(self._dados_extraidos), len(self._nomes_colunas),
            total, self._mes_ref, self._ano_ref,
        )

    def _preencher_tabela_ies(self) -> None:
        self._textbox_ies.configure(state="normal")
        self._textbox_ies.delete("1.0", "end")

        # Status por processo: colunas ATC, NORMAL e DIFAL (pendente / —)
        header = (
            f"{'#':>4}  {'I.E.':>12}  {'Valor ATC':>14}  {'NORMAL':>14}  {'DIF. ALIQUOTA':>14}  {'ATC':>8}  {'Normal':>8}  {'DIFAL':>8}\n"
        )
        sep = "─" * 96 + "\n"
        self._textbox_ies.insert("end", header)
        self._textbox_ies.insert("end", sep)

        for idx, item in enumerate(self._lista_dados, 1):
            ie = _ie_para_exibicao(str(item.get("ie", "")))
            valor_atc = item.get("valor_atc")
            valor_normal = item.get("valor_normal")
            valor_difal = item.get("valor_difal")
            if _valor_atc_invalido(valor_atc):
                atc_str = "—"
            else:
                atc_str = f"R$ {float(valor_atc):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if _valor_normal_invalido(valor_normal):
                normal_str = "—"
            else:
                normal_str = f"R$ {float(valor_normal):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if _valor_difal_invalido(valor_difal):
                difal_str = "—"
            else:
                difal_str = f"R$ {float(valor_difal):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            status_atc = "pendente" if not _valor_atc_invalido(valor_atc) else "ignorada"
            status_normal = "pendente" if not _valor_normal_invalido(valor_normal) else "ignorada"
            status_difal = "pendente" if not _valor_difal_invalido(valor_difal) else "ignorada"
            self._textbox_ies.insert(
                "end",
                f"{idx:>4}  {ie:>12}  {atc_str:>14}  {normal_str:>14}  {difal_str:>14}  {status_atc:>8}  {status_normal:>8}  {status_difal:>8}\n",
            )

        self._textbox_ies.configure(state="disabled")

    def _ao_ver_dados(self) -> None:
        if not self._dados_extraidos or not self._caminho_excel:
            messagebox.showinfo("Dados", "Nenhum dado extraído para exibir.")
            return
        _mostrar_janela_dados_extraidos(
            self, self._caminho_excel, self._dados_extraidos,
            self._nomes_colunas, self._nome_para_indice,
        )

    # --- Executar ---
    def _ao_executar(self) -> None:
        if self._executando or not self._lista_dados:
            return

        processos = [pid for pid, var in self._vars_processos.items() if var.get()]
        if not processos:
            messagebox.showwarning("Nenhum processo", "Selecione ao menos um processo.")
            return

        if self._modo_ies == self._MODO_SELECAO:
            lista_por_processo = {
                pid: [item for (item, var) in self._selecao_por_processo.get(pid, []) if var.get()]
                for pid in processos
            }
            total_selecionadas = sum(len(lst) for lst in lista_por_processo.values())
            if total_selecionadas == 0:
                messagebox.showwarning(
                    "Aviso",
                    "Nenhuma I.E. selecionada nas listas dos processos. "
                    "Escolha o processo (ATC / Normal / DIFAL) e marque as IEs desejadas.",
                )
                return
            descricao_qtd = f"{total_selecionadas} IE(s) selecionada(s) nos processos"
        else:
            lista_por_processo = None
            lista = [
                item for item in self._lista_dados
                if any(_item_executavel_para_processo(item, p) for p in processos)
            ]
            if not lista:
                messagebox.showwarning(
                    "Aviso",
                    "Nenhuma I.E. com valor para o(s) processo(s) selecionado(s). "
                    "Alterne para 'Selecionar IEs' para ver o detalhamento.",
                )
                return
            descricao_qtd = f"{len(lista)} IE(s) executável(is)"

        nomes = ", ".join(_nome_processo_legivel(p) for p in processos)

        confirmacao = messagebox.askyesno(
            "Confirmar execução",
            f"Executar {nomes} para {descricao_qtd}?\n"
            f"Período: {self._mes_ref:02d}/{self._ano_ref}\n"
            f"Headless: {'Sim' if self._var_headless.get() else 'Não'}",
        )
        if not confirmacao:
            return

        self._executando = True
        self._habilitar_botoes(False)
        self._btn_executar.configure(text="⏳  Executando…")
        self._status(f"Executando {descricao_qtd}…")
        self._log(f"\n{'═' * 40}")
        self._log(f"Executando: {nomes}  |  {descricao_qtd}  |  headless={self._var_headless.get()}")
        self._log(f"{'═' * 40}")

        def _ao_finalizar(ies_ok: list[str], ies_erro: list[tuple[str, str]]) -> None:
            self.after(0, self._finalizar_execucao, ies_ok, ies_erro)

        if self._modo_ies == self._MODO_SELECAO and lista_por_processo is not None:
            _executar_lote_em_background(
                [], processos,
                self._var_headless.get(), result_callback=_ao_finalizar,
                lista_por_processo=lista_por_processo,
            )
        else:
            _executar_lote_em_background(
                lista, processos,
                self._var_headless.get(), result_callback=_ao_finalizar,
            )

    def _finalizar_execucao(
        self, ies_ok: list[str], ies_erro: list[tuple[str, str]],
    ) -> None:
        self._executando = False
        self._habilitar_botoes(True)
        self._btn_executar.configure(text="▶  Executar")

        self._log(f"\n{'─' * 40}")
        self._log(f"Concluído: {len(ies_ok)} sucesso, {len(ies_erro)} erro(s)")
        if ies_ok:
            self._log(f"  Sucesso: {', '.join(ies_ok)}")
        if ies_erro:
            for ie, motivo in ies_erro:
                self._log(f"  Erro IE {ie}: {motivo}")
        self._log(f"{'─' * 40}\n")

        total = len(ies_ok) + len(ies_erro)
        if ies_erro:
            self._status(f"Finalizado: {len(ies_ok)}/{total} sucesso, {len(ies_erro)} erro(s)")
        else:
            self._status("Execução concluída com sucesso!")


def main() -> None:
    """Ponto de entrada da GUI ICMS-PI."""
    logger.info("Iniciando interface ICMS-PI.")
    app = App()
    app.mainloop()
    logger.info("Interface encerrada.")


if __name__ == "__main__":
    main()
    sys.exit(0)
