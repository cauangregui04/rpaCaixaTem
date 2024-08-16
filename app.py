from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
import time
import pandas as pd
from io import StringIO

global carga

carga = pd.read_excel("cargaCaixa.xlsx", dtype=str)

i = 0

servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
prefs = {
    "download.prompt_for_download": True,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": False,
}
options.add_experimental_option("prefs", prefs)
navegador = webdriver.Chrome(service=servico, options=options)

navegador.get("https://caixaaqui.caixa.gov.br/caixaaqui/CaixaAquiController")
navegador.maximize_window()
time.sleep(3)
navegador.find_element(
    "xpath",
    "/html/body/table/tbody/tr/td/center/form/div/div[2]/table/tbody/tr[1]/td[3]/input",
).send_keys("numero")
navegador.find_element(
    "xpath",
    "/html/body/table/tbody/tr/td/center/form/div/div[2]/table/tbody/tr[2]/td[3]/input",
).send_keys("usu")
navegador.find_element(
    "xpath",
    "/html/body/table/tbody/tr/td/center/form/div/div[2]/table/tbody/tr[3]/td[3]/input",
).send_keys("senha")
navegador.find_element(
    "xpath",
    "/html/body/table/tbody/tr/td/center/form/div/div[2]/table/tbody/tr[4]/td[2]/input",
).click()
time.sleep(3)

try:
    navegador.find_element(
        "xpath", "/html/body/center/div/table/tbody/tr[2]/td/div/a[1]"
    ).click()
    wait = WebDriverWait(navegador, timeout=2)
    alert = wait.until(lambda d: d.switch_to.alert)
    text = alert.text
    alert.accept()
    time.sleep(2)
    navegador.find_element(
        "xpath",
        "/html/body/table/tbody/tr/td/center/form/div/div[2]/table/tbody/tr[1]/td[3]/input",
    ).send_keys("numero")
    navegador.find_element(
        "xpath",
        "/html/body/table/tbody/tr/td/center/form/div/div[2]/table/tbody/tr[2]/td[3]/input",
    ).send_keys("usu")
    navegador.find_element(
        "xpath",
        "/html/body/table/tbody/tr/td/center/form/div/div[2]/table/tbody/tr[3]/td[3]/input",
    ).send_keys("senha")
    navegador.find_element(
        "xpath",
        "/html/body/table/tbody/tr/td/center/form/div/div[2]/table/tbody/tr[4]/td[2]/input",
    ).click()
    time.sleep(5)
    navegador.find_element(
        "xpath", "/html/body/div/div[1]/div[2]/form/center/table[2]/tbody/tr[1]/td/a"
    ).click()

except:
    navegador.find_element(
        "xpath", "/html/body/div/div[1]/div[2]/form/center/table[2]/tbody/tr[1]/td/a"
    ).click()  # SERVIÇO AO CLIENTE

time.sleep(2)
navegador.find_element(
    "xpath", "/html/body/center/table[2]/tbody/tr[1]/td/a"
).click()  # NEGÓCIOS
time.sleep(3)
navegador.find_element(
    "xpath", "/html/body/center/form/table[1]/tbody/tr[1]/td/a"
).click()  # PESQUISAR CLIENTES
time.sleep(3)


def caixaAqui():
    global dados
    global df1

    cpf = carga.iloc[i, 0]

    navegador.find_element(
        "xpath",
        "/html/body/form/center/div[2]/div[2]/table/tbody/tr[3]/td/div[1]/input",
    ).send_keys(
        cpf
    )  # CPF
    time.sleep(2)
    navegador.find_element(
        "xpath", "/html/body/form/center/div[2]/div[2]/table/tbody/tr[3]/td/div[1]/a"
    ).click()  # PESQUISAR BOTÃO
    time.sleep(2)
    navegador.find_element("xpath", "/html/body/center/div[2]").click()
    tbl = navegador.find_element(
        "xpath", "/html/body/center/div[2]/div[2]"
    ).get_attribute("outerHTML")

    dados = pd.read_html(StringIO(tbl))
    cons = pd.read_html(StringIO(tbl))[0]

    a = 1

    for a in range(len(dados)):
        temp = pd.read_html(StringIO(tbl))[a]
        cons = pd.concat([cons, temp])
        a += 1

    navegador.find_element("xpath", "/html/body/center/div[2]/div[1]/div[2]/a").click()

    if i == 0:
        df1 = cons
    else:
        df1 = pd.concat([df1, cons])


def exportExcel(writer):
    writer = pd.ExcelWriter("CaixaAqui.xlsx")
    df1.to_excel(writer, index=False)
    writer.close()


i = 0

try:
    while i < len(carga):
        print(f"PROCESSO {i + 1} DE {len(carga)} DA CARGA EM ANDAMENTO")
        caixaAqui()
        i += 1
except:
    print(f"PROCESSO PAROU NA LINHA {i + 1} DA CARGA")

exportExcel(df1)
input("===========PROCESSO FINALIZADO===========")
