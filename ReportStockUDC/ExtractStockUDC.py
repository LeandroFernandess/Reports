"""

Filename: ExtractStockUDC.py
Developer: Leandro Fernandes
Date: 20/08/2024
Description: O código extrai o arquivo contendo o estoque atual do armazém.

"""

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import os
import shutil

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


def ExtractFile():

    print(
        "A extração do arquivo do Estoque UDC será iniciada!",
        "\n" * 2,
    )
    service = Service(
        r"C:\Users\F89074d\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.12\Driver EDGE\msedgedriver.exe"
    )
    Browser = webdriver.Edge(service=service)

    Browser.maximize_window()

    Browser.get(
        "https://contagem.click.cnh.com/ck/bs/browser/main.jsp?_PAGE_URL=%2Fck%2Fbs%2Fpages%2FstartItem.jsp%3FCD_ITEM%3DStock.Accounting.stock.Estoque%2520UdC&_ACTION_TYPE=5&OBJECT_ID=VIEW_GIACENZA_UDC"
    )
    username_field = WebDriverWait(Browser, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div/div[2]/div[1]/form/table/tbody/tr[1]/td/input")
        )
    )
    username_field.send_keys("F89074D")

    Browser.find_element(By.NAME, "_PASSWORD").send_keys("Cnhindustrial.2024")

    Browser.find_element(By.NAME, "login").click()

    Browser.find_element(By.ID, "QUERY_BUTTON").click()

    iframe_index = 1
    total_iframes = len(Browser.find_elements(By.TAG_NAME, "iframe"))

    if iframe_index < total_iframes:
        Browser.switch_to.frame(iframe_index)
    else:
        print(
            f"O index do iframe {iframe_index} não foi localizado. Tem um total de {total_iframes} iframes na página."
        )
        return

    # Executa o JavaScript da página para extrair o arquivo em formato CSV
    export_csv_script = """
    var item = document.querySelector("a[onclick*='showExportCsv']");
    if (item) {
        item.click();
        return true;
    } else {
        return false;
    }
    """

    try:
        Browser.execute_script(export_csv_script)
    except Exception as e:
        print(f"Ocorreu um erro ao tentar extrair o arquivo CSV: {e}")
        return

    handles_before = Browser.window_handles
    Browser.switch_to.window(handles_before[1])

    Browser.maximize_window()
    Browser.find_element(By.ID, "ok_btn").click()

    file_path = r"C:\Users\F89074d\Downloads\Estoque+udc.csv"
    while not os.path.exists(file_path):
        sleep(1)
    origin_path = r"C:\Users\F89074d\Downloads\Estoque+udc.csv"
    destination_path = r"C:\Users\F89074d\Documents\Python - Arquivos Base"

    try:
        shutil.copy(origin_path, destination_path)
    except Exception as e:
        print(f"Ocorreu um erro para transferir o arquivo: {e}")

    os.remove(origin_path)
    print(
        "\n",
        "A extração do arquivo do Estoque UDC foi finalizada com sucesso!",
        "\n" * 2,
    )
