import argparse
import json
import logging
import os
from pathlib import Path

from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openpyxl import Workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

URL = "http://guri.unipampa.edu.br"
DEFAULT_OUTPUT = "cadeiras.xlsx"
DEFAULT_TIMEOUT = 20


def setup_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extrai disciplinas e faltas do GURI.")
    parser.add_argument("--ano", default=os.getenv("GURI_ANO"), help="Ano a filtrar (ex.: 2026).")
    parser.add_argument(
        "--semestre",
        default=os.getenv("GURI_SEMESTRE"),
        help='Semestre/período a filtrar (ex.: "2. Semestre").',
    )
    parser.add_argument("--all", action="store_true", help="Ignora filtros e exporta todas as disciplinas.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Arquivo Excel de saída.")
    parser.add_argument("--config", default="config.json", help="Caminho para config.json.")
    parser.add_argument("--headless", action="store_true", help="Executa o Chrome em modo headless.")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help="Timeout padrão em segundos.")
    parser.add_argument("--verbose", action="store_true", help="Ativa logs detalhados.")
    return parser.parse_args()


def load_credentials(config_path: str) -> tuple[str, str]:
    load_dotenv()
    login = os.getenv("GURI_LOGIN")
    senha = os.getenv("GURI_SENHA")
    if login and senha:
        return login, senha

    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(
            f"Credenciais não encontradas. Defina GURI_LOGIN/GURI_SENHA ou crie {config_path}."
        )

    with path.open(encoding="utf-8") as config_file:
        config = json.load(config_file)

    login = config.get("login")
    senha = config.get("senha")
    if not login or not senha:
        raise ValueError(f"Arquivo {config_path} deve conter as chaves 'login' e 'senha'.")
    return login, senha


def build_driver(headless: bool) -> webdriver.Chrome:
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=options)


def wait_for(driver: webdriver.Chrome, timeout: int, condition):
    return WebDriverWait(driver, timeout).until(condition)


def js_click(driver: webdriver.Chrome, element) -> None:
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    driver.execute_script("arguments[0].click();", element)


def login_guri(driver: webdriver.Chrome, timeout: int, login: str, senha: str) -> None:
    logging.info("Acessando GURI")
    driver.get(URL)

    campo_login = wait_for(driver, timeout, EC.presence_of_element_located((By.ID, "login")))
    campo_senha = wait_for(driver, timeout, EC.presence_of_element_located((By.ID, "senha")))
    campo_login.clear()
    campo_login.send_keys(login)
    campo_senha.clear()
    campo_senha.send_keys(senha)

    botao_entrar = wait_for(driver, timeout, EC.presence_of_element_located((By.ID, "entrar")))
    js_click(driver, botao_entrar)
    wait_for(driver, timeout, EC.presence_of_element_located((By.ID, "novo_conteudo-login")))
    logging.info("Login realizado")


def abrir_diario_de_classe(driver: webdriver.Chrome, timeout: int) -> None:
    logging.info("Abrindo Diário de Classe")
    try:
        botao = wait_for(
            driver,
            timeout,
            EC.presence_of_element_located((By.XPATH, "//button[contains(normalize-space(.), 'Diário de Classe')]")),
        )
    except TimeoutException:
        botao = wait_for(
            driver,
            timeout,
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="novo_conteudo-login"]/div/div/div/div/div/div/div[2]/button[8]')
            ),
        )
    js_click(driver, botao)
    wait_for(driver, timeout, EC.presence_of_element_located((By.TAG_NAME, "tbody")))


def listar_disciplinas(driver: webdriver.Chrome) -> list[dict]:
    soup = BeautifulSoup(driver.page_source, "html.parser")
    tbody = soup.find("tbody")
    if not tbody:
        return []

    disciplinas = []
    for linha in tbody.find_all("tr"):
        celulas = linha.find_all("td")
        if len(celulas) < 5:
            continue
        disciplinas.append(
            {
                "id": celulas[0].text.strip(),
                "nome": celulas[1].text.strip(),
                "semestre": celulas[2].text.strip(),
                "carga_horaria": celulas[3].text.strip(),
                "ano": celulas[4].text.strip(),
            }
        )
    return disciplinas


def disciplina_corresponde(disciplina: dict, ano: str | None, semestre: str | None) -> bool:
    if ano and disciplina["ano"] != ano:
        return False
    if semestre and disciplina["semestre"] != semestre:
        return False
    return True


def extrair_faltas(driver: webdriver.Chrome, timeout: int, disciplina_id: str) -> int | str:
    driver.get(f"https://guri.unipampa.edu.br/pta/diario/visualizar/{disciplina_id}")
    wait_for(driver, timeout, EC.presence_of_element_located((By.TAG_NAME, "body")))
    soup = BeautifulSoup(driver.page_source, "html.parser")
    faltas_tags = soup.find_all(title="Falta")
    contador = len(faltas_tags)
    if contador == 0:
        return "aulas não lançadas"
    return contador - 1


def exportar_xlsx(disciplinas: list[dict], output: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Disciplinas"
    sheet.append(["ID", "Nome", "Carga Horária", "Faltas", "Faltas Possíveis"])

    for d in disciplinas:
        row_index = sheet.max_row + 1
        sheet.append(
            [
                d["id"],
                d["nome"],
                d["carga_horaria"],
                d["faltas"],
                f"=(C{row_index} * 0.25 * 60) / 50",
            ]
        )

    workbook.save(output)
    logging.info("Arquivo salvo: %s (%d disciplinas)", output, len(disciplinas))


def main() -> int:
    args = parse_args()
    setup_logging(args.verbose)

    ano = None if args.all else args.ano
    semestre = None if args.all else args.semestre

    try:
        login, senha = load_credentials(args.config)
    except Exception as exc:
        logging.error("%s", exc)
        return 1

    driver = None
    try:
        driver = build_driver(args.headless)
        login_guri(driver, args.timeout, login, senha)
        abrir_diario_de_classe(driver, args.timeout)

        disciplinas = listar_disciplinas(driver)
        logging.info("Disciplinas encontradas na tabela: %d", len(disciplinas))

        selecionadas = [d for d in disciplinas if disciplina_corresponde(d, ano, semestre)]
        if not selecionadas:
            logging.warning("Nenhuma disciplina encontrada para o filtro: ano=%s semestre=%s", ano, semestre)
            for d in disciplinas:
                logging.info("Disponível: %s | %s | %s", d["ano"], d["semestre"], d["nome"])
            exportar_xlsx([], args.output)
            return 0

        resultados = []
        for disciplina in selecionadas:
            logging.info("Extraindo: %s - %s", disciplina["id"], disciplina["nome"])
            try:
                disciplina["faltas"] = extrair_faltas(driver, args.timeout, disciplina["id"])
            except Exception as exc:
                logging.exception("Falha ao extrair disciplina %s: %s", disciplina["id"], exc)
                disciplina["faltas"] = "erro ao extrair"
            resultados.append(disciplina)

        exportar_xlsx(resultados, args.output)
        return 0

    except (WebDriverException, TimeoutException, FileNotFoundError, ValueError) as exc:
        logging.error("Erro na execução: %s", exc)
        return 1
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                logging.debug("Falha ao encerrar o navegador", exc_info=True)


if __name__ == "__main__":
    raise SystemExit(main())
