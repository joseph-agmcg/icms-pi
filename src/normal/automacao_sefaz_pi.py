"""
Automação ICMS Normal / Imposto, Juros e Multa (SEFAZ-PI, DAR Web).
Orquestra o fluxo por I.E.: portal PI → Menu ICMS → código 113000 → IE, substituição NÃO,
Avançar → preenche período, datas e valor principal (coluna NORMAL).
"""

import asyncio
from datetime import date, datetime

from playwright.async_api import async_playwright, Browser, Page, BrowserContext

from icms_pi import configuracoes
from . import configuracoes as configuracoes_normal
from icms_pi.logger import configurar_logger_da_aplicacao
from atc.navegacao.acoes_pagina import (
    aguardar_pagina_carregar,
    preencher_campo_data_mascarado,
    preencher_campo_valor_mascarado,
    tirar_captura_de_tela_em_erro,
)

logger = configurar_logger_da_aplicacao(__name__)


def _valor_normal_invalido(valor: object) -> bool:
    """Retorna True se o valor Normal (coluna NORMAL) não deve ser processado."""
    if valor is None:
        return True
    if isinstance(valor, (int, float)) and valor == 0:
        return True
    if isinstance(valor, str):
        s = valor.strip().lower()
        if not s or s == "null":
            return True
    return False


def _data_dia_15_mes_referencia(mes_ref: int, ano_ref: int) -> tuple[int, int, int]:
    """Retorna (dia, mês, ano) para o dia 15 do mês de referência."""
    return 15, mes_ref, ano_ref


def _data_vencimento_no_passado(mes_ref: int, ano_ref: int) -> bool:
    """Retorna True se a data de vencimento (dia 15 do mês de referência) já passou."""
    dia, mes, ano = _data_dia_15_mes_referencia(mes_ref, ano_ref)
    try:
        data_venc = date(ano, mes, dia)
        return data_venc < date.today()
    except ValueError:
        return True


class AutomacaoNormalPI:
    """
    Automação para ICMS Normal / Imposto, Juros e Multa (SEFAZ-PI, DAR Web).
    Fluxo: portal PI → Menu ICMS → 113000 → IE, substituição tributária NÃO →
    Avançar → período, datas (dia 15), valor principal (coluna NORMAL) → Calcular Imposto.
    """

    def __init__(self, headless: bool = False) -> None:
        self._headless = headless
        self._playwright = None
        self._browser: Browser | None = None
        self._context: BrowserContext | None = None
        self._pagina: Page | None = None

    async def _iniciar_browser(self) -> None:
        logger.info("Iniciando navegador (ICMS Normal PI).")
        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(headless=self._headless)
        self._context = await self._browser.new_context(
            viewport={"width": 1280, "height": 720},
            ignore_https_errors=True,
        )
        self._pagina = await self._context.new_page()
        self._pagina.set_default_timeout(configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS)
        logger.debug("Navegador e página prontos.")

    async def _encerrar_browser(self) -> None:
        if self._context:
            await self._context.close()
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        logger.info("Navegador encerrado.")

    def _nome_captura_erro(self, etapa: str, ie: str = "") -> str:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sufixo = f"_{ie}" if ie else ""
        return f"erro_normal_{etapa}{sufixo}_{timestamp}.png"

    async def _acessar_pagina_inicial_pi(self) -> None:
        """Acessa a URL do DAR Web (SEFAZ-PI)."""
        logger.info("Acessando %s", configuracoes.URL_PORTAL_DARWEB_SEFAZ_PI)
        await self._pagina.goto(configuracoes.URL_PORTAL_DARWEB_SEFAZ_PI)
        await aguardar_pagina_carregar(self._pagina)

    async def _clicar_menu_icms_pi(self) -> None:
        """Clica no link do menu ICMS."""
        locator = self._pagina.locator(configuracoes_normal.SELETOR_PI_MENU_ICMS).filter(
            has_text="ICMS"
        )
        await locator.first.wait_for(
            state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS
        )
        await locator.first.click()
        await aguardar_pagina_carregar(self._pagina)
        logger.debug("Clicado no menu ICMS.")

    async def _selecionar_imposto_juros_multa_pi(self) -> None:
        """Seleciona a opção 113000 - ICMS - IMPOSTO, JUROS E MULTA no select."""
        select = self._pagina.locator(configuracoes_normal.SELETOR_PI_SELECT_CODIGO)
        await select.wait_for(
            state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS
        )
        await select.select_option(
            label=configuracoes_normal.VALOR_OPCAO_PI_IMPOSTO_JUROS_MULTA
        )
        await aguardar_pagina_carregar(self._pagina)
        logger.debug("Selecionado: 113000 - ICMS Imposto, Juros e Multa.")

    async def _clicar_botao_avancar_pi(self) -> None:
        """Clica no botão Avançar (mesmo seletor do ATC)."""
        locator = self._pagina.locator(configuracoes_normal.SELETOR_PI_BOTAO_AVANCAR).filter(
            has_text="Avançar"
        )
        await locator.first.wait_for(
            state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS
        )
        await locator.first.click()
        await aguardar_pagina_carregar(self._pagina)
        logger.debug("Clicado no botão Avançar.")

    async def _preencher_ie_pi(self, ie_digitos: str) -> None:
        """Preenche o campo de Inscrição Estadual (#fieldInscricaoEstadual)."""
        campo = self._pagina.locator(configuracoes_normal.SELETOR_PI_CAMPO_IE)
        await campo.wait_for(
            state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS
        )
        await campo.fill("")
        await campo.fill(ie_digitos)
        logger.debug("Campo IE preenchido com %s", ie_digitos)

    async def _selecionar_substituicao_nao_pi(self) -> None:
        """Seleciona 'Não' no campo Substituição Tributária (cmbSubstituicao)."""
        select = self._pagina.locator(configuracoes_normal.SELETOR_PI_SUBSTITUICAO)
        await select.wait_for(
            state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS
        )
        await select.select_option(value=configuracoes_normal.VALOR_SUBSTITUICAO_NAO)
        logger.debug("Substituição tributária: NÃO.")

    async def _preencher_periodo_pi(self, mes_ref: int, ano_ref: int) -> None:
        """Preenche o período de referência (formato MM/AAAA)."""
        periodo_str = f"{mes_ref:02d}/{ano_ref}"
        locator = self._pagina.locator(configuracoes_normal.SELETOR_PI_PERIODO_REFERENCIA)
        await locator.wait_for(
            state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS
        )
        await locator.fill("")
        await locator.fill(periodo_str)
        logger.debug("Período de referência preenchido: %s", periodo_str)

    async def _preencher_datas_vencimento_pagamento_pi(
        self, mes_ref: int, ano_ref: int
    ) -> None:
        """Preenche Vencimento e Pagamento com o dia 15 do mês de referência."""
        dia, mes, ano = _data_dia_15_mes_referencia(mes_ref, ano_ref)
        data_str = f"{dia:02d}/{mes:02d}/{ano}"
        await preencher_campo_data_mascarado(
            self._pagina,
            configuracoes_normal.SELETOR_PI_DATA_VENCIMENTO,
            data_str,
        )
        await preencher_campo_data_mascarado(
            self._pagina,
            configuracoes_normal.SELETOR_PI_DATA_PAGAMENTO,
            data_str,
        )
        logger.debug("Datas Vencimento e Pagamento preenchidas: %s", data_str)

    async def _preencher_valor_principal_pi(self, valor_principal: float) -> None:
        """Preenche o valor principal (coluna NORMAL)."""
        await preencher_campo_valor_mascarado(
            self._pagina,
            configuracoes_normal.SELETOR_PI_VALOR_PRINCIPAL,
            valor_principal,
        )
        logger.debug("Valor principal (Normal) preenchido: %s", valor_principal)

    async def _clicar_botao_calcular_imposto_pi(self) -> None:
        """Clica no botão Calcular Imposto."""
        locator = self._pagina.locator(
            configuracoes_normal.SELETOR_PI_BOTAO_CALCULAR_IMPOSTO
        ).filter(has_text="Calcular Imposto")
        await locator.first.wait_for(
            state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS
        )
        await locator.first.click()
        await aguardar_pagina_carregar(self._pagina)
        logger.debug("Clicado no botão Calcular Imposto.")

    async def executar_fluxo_por_ie_pi(
        self,
        lista_dados: list[dict[str, object]],
    ) -> tuple[list[str], list[tuple[str, str]]]:
        """
        Para cada item em lista_dados (ie, ie_digitos, valor_normal, mes_ref, ano_ref):
        acessa o portal PI, menu ICMS, 113000, preenche IE, substituição NÃO, Avançar,
        período, datas (dia 15) e valor principal (NORMAL).
        Retorna (IEs com sucesso, lista de (IE, motivo) com erro).
        """
        ies_sucesso: list[str] = []
        ies_erro: list[tuple[str, str]] = []
        total = len(lista_dados)

        try:
            await self._iniciar_browser()
            await self._acessar_pagina_inicial_pi()

            for indice, item in enumerate(lista_dados):
                ie = str(item.get("ie", ""))
                ie_digitos = str(item.get("ie_digitos", "") or ie)
                valor_normal = item.get("valor_normal")
                try:
                    mes_ref = int(item.get("mes_ref"))
                    ano_ref = int(item.get("ano_ref"))
                except (TypeError, ValueError):
                    ies_erro.append(
                        (ie or "(vazio)", "Período (mês/ano) ausente nos dados da planilha")
                    )
                    continue

                if not ie or not ie_digitos:
                    ies_erro.append((ie or "(vazio)", "IE inválida ou vazia"))
                    continue

                if _valor_normal_invalido(valor_normal):
                    logger.info(
                        "IE %s pulada: valor NORMAL ausente, zero ou vazio.",
                        ie,
                    )
                    continue

                if _data_vencimento_no_passado(mes_ref, ano_ref):
                    motivo = "Data de vencimento no passado — portal não permite datas passadas"
                    ies_erro.append((ie, motivo))
                    logger.info("IE %s pulada: %s", ie, motivo)
                    continue

                logger.info("Processando IE %s Normal (%d/%d).", ie, indice + 1, total)

                if indice > 0:
                    ms = configuracoes.INTERVALO_ENTRE_EXECUCOES_MS
                    logger.info(
                        "Aguardando %d ms (%.1f s) antes da próxima IE.",
                        ms, ms / 1000.0,
                    )
                    await asyncio.sleep(ms / 1000.0)
                    await self._acessar_pagina_inicial_pi()

                try:
                    await self._clicar_menu_icms_pi()
                    await self._selecionar_imposto_juros_multa_pi()
                    await self._clicar_botao_avancar_pi()
                    await self._preencher_ie_pi(ie_digitos)
                    await self._selecionar_substituicao_nao_pi()
                    await self._clicar_botao_avancar_pi()
                    await self._preencher_periodo_pi(mes_ref, ano_ref)
                    await self._preencher_datas_vencimento_pagamento_pi(mes_ref, ano_ref)
                    await self._preencher_valor_principal_pi(float(valor_normal))
                    await self._clicar_botao_calcular_imposto_pi()
                except Exception as e:
                    logger.exception(
                        "Erro ao preencher formulário Normal PI para IE %s: %s", ie, e
                    )
                    await tirar_captura_de_tela_em_erro(
                        self._pagina,
                        self._nome_captura_erro("formulario_normal", ie),
                    )
                    motivo = (
                        str(e).split("\n")[0].strip() if str(e)
                        else "Falha ao preencher formulário ICMS Normal"
                    )
                    if len(motivo) > 80:
                        motivo = motivo[:77] + "..."
                    ies_erro.append((ie, motivo))
                    continue

                ies_sucesso.append(ie)
                logger.info("IE %s concluída (formulário Normal preenchido).", ie)

        finally:
            await self._encerrar_browser()

        logger.info(
            "Fluxo Normal PI finalizado: %d sucesso, %d erro.",
            len(ies_sucesso), len(ies_erro),
        )
        return ies_sucesso, ies_erro
