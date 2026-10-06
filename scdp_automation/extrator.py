"""Consulta o relatório anual CCH e salva viagens validadas em JSON."""

from __future__ import annotations

import argparse
import asyncio
import re
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urljoin, urlsplit
from zoneinfo import ZoneInfo

from loguru import logger
from playwright.async_api import Error as PlaywrightError
from playwright.async_api import Locator, Page, async_playwright
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from scdp_automation.autenticacao import authenticate_gov_br, load_credentials
from scdp_automation.chrome_profile_setup import selected_profile_directory
from scdp_automation.config import REPO_ROOT, current_year, sync_extraction_config
from scdp_automation.logging_config import configure_logging
from scdp_automation.navegador_chrome import connect_visible_chrome
from scdp_automation.output_history import OutputHistory
from scdp_automation.relatorio import Viagem, load_trips, parse_report_rows, save_json
from scdp_automation.xlsx_output import DEFAULT_WORKBOOK, publish_workbook

SCDP_URL = "https://www2.scdp.gov.br/"
DEFAULT_OUTPUT = Path(f"output/viagens_scdp_{current_year()}.json")


def resolve_report_url(current_url: str, report_href: str) -> str:
    """Resolve o link do relatório encontrado no menu autenticado do SCDP."""
    return urljoin(current_url, report_href)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera JSON validado das viagens do relatório Viagem do SCDP."
    )
    parser.add_argument(
        "--limite",
        type=int,
        default=0,
        help="Limita as novas PCDPs desta execução (0 = todas as pendentes).",
    )
    args = parser.parse_args(argv)
    if args.limite < 0:
        parser.error("--limite não pode ser negativo.")
    config = sync_extraction_config()
    args.ano = config.ano
    args.output = Path("output") / config.checkpoint_name
    args.workbook = REPO_ROOT / "output" / config.nome_planilha
    return args


async def find_report_table(page: Page) -> Locator:
    """Localiza a tabela do relatório pela coluna PCDP, sem depender de IDs JSF."""
    tables = page.locator("table")
    for index in range(await tables.count()):
        table = tables.nth(index)
        text = " ".join((await table.inner_text()).split())
        if "Número da Solicitação" in text and (
            "Quantidade Diárias" in text
            or "Quantidade de Diárias" in text
            or "Diárias" in text
        ):
            return table
    raise RuntimeError(
        "Não encontrei a tabela do relatório Viagem. Confirme que a pesquisa foi "
        "executada e que a página atual contém os resultados."
    )


async def wait_for_login(page: Page) -> None:
    logger.info("Aguardando retorno ao SCDP e abertura do menu RELATÓRIOS.")
    await page.wait_for_url(
        re.compile(r"^https://www2\.scdp\.gov\.br/novoscdp/"), timeout=0
    )
    await page.get_by_role("menuitem", name="RELATÓRIOS", exact=True).wait_for(
        state="visible", timeout=0
    )


async def select_cch(page: Page) -> None:
    """Confirma ou seleciona CCH no autocomplete do órgão solicitante."""
    field = page.locator('[id="selOrgao:autoComplete_input"]')
    await field.wait_for(state="visible")
    current_value = await field.input_value()
    if "121766" in current_value and "Chapecó" in current_value:
        return

    await field.fill("Campus Chapecó")
    suggestion = page.get_by_role(
        "option", name=re.compile(r"Campus Chapecó/SC.*121766", re.IGNORECASE)
    ).first
    try:
        await suggestion.wait_for(state="visible", timeout=5_000)
        await suggestion.click()
    except PlaywrightTimeoutError:
        # O autocomplete PrimeFaces pode expor a sugestão como item de lista,
        # sem role=option. Aceita somente a opção que contém o código do CCH.
        suggestion_text = page.get_by_text(
            re.compile(r"Campus Chapecó/SC.*121766", re.IGNORECASE)
        ).last
        await suggestion_text.wait_for(state="visible", timeout=5_000)
        await suggestion_text.click()

    selected_value = await field.input_value()
    if "121766" not in selected_value or "Chapecó" not in selected_value:
        raise RuntimeError(
            "Não consegui confirmar CCH — Campus Chapecó/SC no filtro de órgão. "
            "A consulta foi interrompida para evitar extrair outro órgão."
        )


async def open_annual_cch_report(page: Page, year: int | None = None) -> None:
    """Navega aos filtros, escolhe o ano de exercício completo e pesquisa CCH."""
    # O link oficial já existe no DOM; hover depende de frames de animação
    # que o Chrome pode suspender quando a janela fica em segundo plano.
    report_link = (
        page.locator('a[href*="/relatorio/relatorio_viagem.xhtml"]')
        .filter(has_text="Viagem")
        .last
    )
    await report_link.wait_for(state="attached")
    report_href = await report_link.get_attribute("href")
    if not report_href:
        raise RuntimeError("O link do relatório Viagem não possui href navegável.")
    await page.goto(
        resolve_report_url(page.url, report_href), wait_until="domcontentloaded"
    )

    await page.locator("#chkAnoExercicio").wait_for(state="visible")
    await select_cch(page)

    await ensure_annual_period(page, year)
    async with page.expect_navigation(wait_until="domcontentloaded"):
        await page.get_by_role("button", name="Pesquisar", exact=True).click(force=True)
    await find_report_table(page)
    # Mantém o tamanho atual para não disparar outra submissão JSF concorrente.


async def ensure_annual_period(page: Page, year: int | None = None) -> None:
    """Espera as datas anuais; remarcar recupera um checkbox com estado antigo."""
    annual = page.locator("#chkAnoExercicio")
    year = year or sync_extraction_config().ano
    condition = r"""year => {
      const start = document.getElementById('dataInicioRelatorio:inputCalendario_input');
      const end = document.getElementById('dataFimRelatorio:inputCalendario_input');
      return document.getElementById('chkAnoExercicio')?.checked
        && start?.value === '01/01/' + year
        && end?.value === '31/12/' + year;
    }"""
    if await annual.is_checked():
        if await page.evaluate(condition, arg=year):
            return
        old_input = await page.locator(
            '[id="dataInicioRelatorio:inputCalendario_input"]'
        ).element_handle()
        await annual.uncheck(force=True)
        # O AJAX Mojarra substitui o campo ao atualizar divPeriodo.
        await page.wait_for_function(
            "element => !element.isConnected", arg=old_input, polling=200
        )
    await annual.check(force=True)
    # Polling por tempo também funciona com o Chrome sem frames de animação.
    await page.wait_for_function(condition, arg=year, polling=200, timeout=30_000)


async def listing_numbers(page: Page) -> list[str]:
    table = await find_report_table(page)
    return await table.locator("a").evaluate_all(
        r"""links => links.map(a => a.innerText.trim()).filter(t => /^\d{6}\/\d{2}(?:-\d+[A-Z]+)?$/.test(t))"""
    )


async def wait_listing_change(page: Page, previous_numbers: list[str]) -> None:
    await page.wait_for_function(
        r"""old => {
          const current = [...document.querySelectorAll('table a')]
            .map(a => a.innerText.trim()).filter(t => /^\d{6}\/\d{2}(?:-\d+[A-Z]+)?$/.test(t));
          return current.length > 0 && JSON.stringify(current) !== JSON.stringify(old);
        }""",
        arg=previous_numbers,
        polling=200,
        timeout=30_000,
    )


async def next_page(page: Page) -> bool:
    """Avança pelo controle real do SCDP e espera a nova lista completa."""
    control = page.locator(".ui-paginator-next").first
    await control.wait_for(state="attached")
    if "ui-state-disabled" in (await control.get_attribute("class") or ""):
        return False
    before = await listing_numbers(page)
    await control.click(force=True)
    await wait_listing_change(page, before)
    return True


async def collect_listing(page: Page) -> list[Viagem]:
    """Coleta os trechos e totais de todas as páginas antes das consultas."""
    trips: dict[str, Viagem] = {}
    seen_pages: set[tuple[str, ...]] = set()
    while True:
        table = await find_report_table(page)
        rows = await table.evaluate("""table => [...table.rows].map(row =>
          [...row.cells].map(cell => ({text: cell.innerText.trim(),
            rowspan: cell.rowSpan, colspan: cell.colSpan})))""")
        current = parse_report_rows(rows)
        signature = tuple(v.numero_da_solicitacao for v in current)
        if signature in seen_pages:
            raise RuntimeError("A paginação repetiu uma página; coleta interrompida.")
        seen_pages.add(signature)
        for trip in current:
            if trip.numero_da_solicitacao in trips:
                raise RuntimeError(
                    f"PCDP repetida na paginação: {trip.numero_da_solicitacao}"
                )
            trips[trip.numero_da_solicitacao] = trip
        logger.info(
            "Página {} coletada: {} solicitações nesta página, {} no total.",
            len(seen_pages),
            len(current),
            len(trips),
        )
        if not await next_page(page):
            return list(trips.values())


async def consult_trip_reason(page: Page, number: str) -> str:
    """Repete uma vez a consulta somente leitura em caso de timeout."""
    try:
        return await _consult_trip_reason(page, number)
    except PlaywrightTimeoutError:
        logger.warning("Timeout ao consultar descrição; repetindo a consulta uma vez.")
        return await _consult_trip_reason(page, number)


async def _consult_trip_reason(page: Page, number: str) -> str:
    """Consulta a descrição do motivo na tela oficial de situação da PCDP."""
    await page.goto(
        "https://www2.scdp.gov.br/novoscdp/pages/consultar_solicitacao/"
        "consultar_situacao_solicitacao_filter.xhtml",
        wait_until="domcontentloaded",
    )
    field = page.locator('[id="tabViewSituacao:consultaSimples:numeroPcdpTxt"]')
    await field.wait_for(state="visible", timeout=30_000)
    # O formulário limita a entrada a nove caracteres; o detalhe devolve
    # o sufixo da complementação, que é conferido antes de aceitar a descrição.
    await field.fill(number.partition("-")[0])
    await field.blur()
    async with page.expect_navigation(
        url=re.compile(r"consultar_solicitacao_detalhes\.xhtml"),
        wait_until="domcontentloaded",
        timeout=30_000,
    ):
        await page.locator(
            '[id="tabViewSituacao:consultaSimples:consultarPCDP"]'
        ).click(force=True)
    description = page.locator(
        '[id="cabecalhoViagem:descricaoMotivoViagemText:outputText"]'
    )
    await description.wait_for(state="visible", timeout=30_000)
    # Confirma que o redirect pertence à PCDP consultada, não a outro detalhe.
    actual_number = (
        await page.locator(
            '[id="cabecalhoViagem:numSolicitacao:numSolicitacao_text:outputText"]'
        ).inner_text()
    ).strip()
    if actual_number != number:
        raise RuntimeError(f"Consulta de {number} retornou a PCDP {actual_number}.")
    return (await description.inner_text()).strip()


def extract_description(detail_text: str) -> str:
    """Extrai o conteúdo visível logo após o rótulo da descrição, se presente."""
    lines = [line.strip() for line in detail_text.splitlines() if line.strip()]
    labels = (
        "Descrição do Motivo da Viagem",
        "Descrição do motivo da viagem",
        "Descrição/Justificativa",
    )
    for index, line in enumerate(lines):
        normalized = re.sub(r"\s+", " ", line).strip()
        for label in labels:
            if normalized.casefold().startswith(label.casefold()):
                remainder = normalized[len(label) :].lstrip(" :")
                if remainder:
                    return remainder
                if index + 1 < len(lines):
                    return lines[index + 1]
    return ""


async def start_login_if_needed(page: Page) -> None:
    """Inicia o login gov.br se necessário e preenche credenciais locais."""
    if urlsplit(page.url).hostname in {"sso.acesso.gov.br", "acesso.gov.br"}:
        await authenticate_gov_br(page, load_credentials())
        return
    reports_menu = page.get_by_role("menuitem", name="RELATÓRIOS", exact=True)
    try:
        await reports_menu.wait_for(state="visible", timeout=2_000)
        return
    except PlaywrightTimeoutError:
        pass
    login_link = page.locator('[id$=":linkLogarGovBR"]')
    await login_link.wait_for(state="visible", timeout=30_000)
    if await reports_menu.is_visible():
        return
    await login_link.click(force=True)
    await page.wait_for_url(
        re.compile(
            r"^https://(?:sso\.acesso\.gov\.br/login|"
            r"www2\.scdp\.gov\.br/novoscdp/pages/main\.xhtml)"
        ),
        timeout=30_000,
    )
    if urlsplit(page.url).hostname == "www2.scdp.gov.br":
        await reports_menu.wait_for(state="visible", timeout=30_000)
        return
    logger.info("Abrindo autenticação gov.br na janela visível.")
    await authenticate_gov_br(page, load_credentials())


async def _consult_with_safe_failure(
    page: Page, number: str, index: int, total: int
) -> str:
    """Consult the reason without exposing request data in an exception chain."""
    try:
        return await consult_trip_reason(page, number)
    except PlaywrightError, RuntimeError:
        logger.error(
            "Falha ao consultar descrição {} de {}; checkpoint preservado.",
            index,
            total,
        )
        raise RuntimeError(
            "Falha ao consultar descrição. O progresso foi preservado; "
            "execute novamente para retomar."
        ) from None


def needs_verification(trip: Viagem, execution_date: date) -> bool:
    """Future trips are consulted every run, regardless of their checkpoint."""
    return (
        trip.data_inicio > execution_date
        or trip.descricao_do_motivo_da_viagem is None
        or trip.data_da_ultima_verificacao is None
    )


async def collect_pending_descriptions(
    page: Page,
    trips: list[Viagem],
    pending: list[Viagem],
    output: Path,
) -> None:
    """Consulta descrições em série e salva cada checkpoint."""
    query_lock = asyncio.Lock()
    for index, trip in enumerate(pending, start=1):
        async with query_lock:
            await page.bring_to_front()
            trip.descricao_do_motivo_da_viagem = await _consult_with_safe_failure(
                page, trip.numero_da_solicitacao, index, len(pending)
            )
            trip.data_da_ultima_verificacao = datetime.now(
                ZoneInfo("America/Sao_Paulo")
            ).date()
            save_json(output, trips)
            logger.info("Descrição {} de {} gravada.", index, len(pending))


def save_checkpoint_and_publish(
    trips: list[Viagem],
    checkpoint_path: Path,
    workbook_path: Path = DEFAULT_WORKBOOK,
    *,
    history: OutputHistory | None = None,
) -> Path | None:
    """Persist the complete JSON checkpoint before refreshing the workbook."""
    backup = history.archive_previous() if history is not None else None
    save_json(checkpoint_path, trips)
    if history is None:
        return publish_workbook(trips, workbook_path)
    publish_workbook(trips, workbook_path, create_backup=False)
    history.prune()
    return backup


async def run(argv: list[str] | None = None) -> None:
    configure_logging()
    args = parse_args(argv)
    output = args.output
    workbook_path = args.workbook
    history = OutputHistory(output, workbook_path)
    previous = {v.numero_da_solicitacao: v for v in load_trips(output)}

    async with async_playwright() as playwright:
        # Todas as abas usam o mesmo perfil Chrome e a mesma sessão autenticada.
        user_data_dir = Path(__file__).resolve().parents[1] / ".scdp-browser"
        profile_directory = selected_profile_directory(user_data_dir)
        if (
            not (user_data_dir / profile_directory).is_dir()
            or not (user_data_dir / "Local State").is_file()
        ):
            raise RuntimeError(
                "A cópia local do perfil Chrome selecionado está incompleta em "
                f"{user_data_dir}. Confirme que o perfil selecionado e Local State existem."
            )
        logger.info("Abrindo ou conectando ao Chrome visível.")
        browser = await connect_visible_chrome(playwright, user_data_dir)
        if not browser.contexts:
            raise RuntimeError("O Chrome conectado não expôs um contexto padrão.")
        context = browser.contexts[0]
        pages = context.pages
        page = next(
            (
                candidate
                for candidate in pages
                if urlsplit(candidate.url).hostname
                in {"www2.scdp.gov.br", "sso.acesso.gov.br", "acesso.gov.br"}
            ),
            pages[0] if pages else await context.new_page(),
        )
        await page.bring_to_front()
        if page.url == "about:blank" or urlsplit(page.url).hostname not in {
            "www2.scdp.gov.br",
            "sso.acesso.gov.br",
            "acesso.gov.br",
        }:
            logger.info("Abrindo o SCDP na aba inicial do Chrome.")
            await page.goto(SCDP_URL, wait_until="domcontentloaded")
        logger.info("Página SCDP carregada no navegador visível.")
        await start_login_if_needed(page)
        await wait_for_login(page)
        await open_annual_cch_report(page, args.ano)

        logger.info("Coletando todas as páginas do relatório Viagem.")
        trips = await collect_listing(page)
        for trip in trips:
            old = previous.get(trip.numero_da_solicitacao)
            if old is not None:
                trip.descricao_do_motivo_da_viagem = old.descricao_do_motivo_da_viagem
                trip.data_da_ultima_verificacao = old.data_da_ultima_verificacao
        backup = save_checkpoint_and_publish(
            trips, output, workbook_path, history=history
        )
        logger.info("Workbook de gastos atualizado em {}.", workbook_path.resolve())
        if backup is None:
            logger.info("Nenhum backup anterior existia para o workbook.")
        else:
            logger.info("Backup anterior do workbook em {}.", backup.resolve())
        logger.info(
            "Classifique PCDPs sem código na última coluna de BASE VIAGENS e "
            "preencha as alocações anuais em APOIO."
        )
        pending = [
            v
            for v in trips
            if needs_verification(v, datetime.now(ZoneInfo("America/Sao_Paulo")).date())
        ]
        if args.limite:
            pending = pending[: args.limite]
        logger.info(
            "{} solicitações no relatório; {} descrições pendentes.",
            len(trips),
            len(pending),
        )
        await collect_pending_descriptions(page, trips, pending, output)
        save_checkpoint_and_publish(trips, output, workbook_path, history=history)
        remaining = sum(v.descricao_do_motivo_da_viagem is None for v in trips)
        logger.info("JSON atualizado em {}.", output.resolve())
        logger.info("Descrições pendentes: {}.", remaining)
        # Desconecta o Playwright sem fechar o Chrome visível.
        logger.info("Desconectando o Playwright; o Chrome permanecerá aberto.")
        try:
            await asyncio.wait_for(browser.close(), timeout=5)
        except TimeoutError:
            logger.warning(
                "A desconexão excedeu 5 segundos; os arquivos foram gravados."
            )
