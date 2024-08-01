import smtplib
import pandas as pd
from datetime import datetime
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from bs4 import BeautifulSoup as bs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


def get_dados_produtos(html):

    soup = bs(html, "html.parser")
    div_produtos = soup.find(
        class_="sc-fBWQRz cULVBz sc-fulCBj fxxByy sc-hsUFQk ceNiJh"
    )
    nomes_produtos = [
        tags.contents[0] for tags in div_produtos.find_all(class_="sc-doohEh dHamKz")
    ]
    qtd_avaliacoes = [
        tags.contents[0] for tags in div_produtos.find_all(class_="sc-epqpcT jdMYPv")
    ]
    urls_produtos = [
        f"https://www.magazineluiza.com.br{link['href']}"
        for link in div_produtos.find_all(
            class_="sc-eBMEME uPWog sc-dxUMQK jeUYOh sc-dxUMQK jeUYOh"
        )
    ]

    dados_produtos = []
    for nome, avaliacao, url in zip(nomes_produtos, qtd_avaliacoes, urls_produtos):
        produto = {
            "PRODUTO": nome.text.strip(),
            "QTD_AVAL": int(avaliacao.text.strip().split()[1][1:-1]),
            "URL": url,
        }
        dados_produtos.append(produto)

    df = pd.DataFrame(dados_produtos)
    df = df.loc[~df["QTD_AVAL"].eq(0)]
    df_piores = df.loc[df["QTD_AVAL"].between(0, 100)].sort_values(
        "QTD_AVAL", ascending=False
    )
    df_melhores = df.loc[~df["QTD_AVAL"].between(0, 100)].sort_values(
        "QTD_AVAL", ascending=False
    )
    with pd.ExcelWriter("output\\Notebook.xlsx") as writer:
        df_melhores.to_excel(writer, sheet_name="Melhores", index=False)
        df_piores.to_excel(writer, sheet_name="Piores", index=False)

    print("\nArquivo de output gerado.")
    return


def send_mail():

    email_address = "testerobo55@gmail.com"
    email_password = "ozpi owip wonz luhv"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    msg = MIMEMultipart()
    msg["From"] = email_address
    msg["To"] = email_address
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = "Relatório Notebooks"

    content = "Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.\n\nAtenciosamente,\nRobô"
    msg.attach(MIMEText(content))

    part = MIMEBase("application", "octet-stream")
    part.set_payload(open("output\\Notebook.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", 'attachment; filename="Notebook.xlsx"')
    msg.attach(part)

    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(email_address, email_password)
    smtp.sendmail(email_address, email_address, msg.as_string())
    smtp.quit()


def main():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    for _ in range(3):
        try:
            driver.get("https://www.magazineluiza.com.br/")
            WebDriverWait(driver, 5).until(
                lambda driver: driver.execute_script("return document.readyState")
                == "complete"
            )
            break
        except:
            print("Timeout ao carregar o site. Tentando novamente.")
    else:
        print("Site fora do ar. Gerando log...")
        f = open("logfile.txt", "a")
        f.write(
            "{0} -- {1}\n".format(
                datetime.now().strftime("%Y-%m-%d %H:%M"), "Site fora do ar."
            )
        )
        f.close()
        return

    search = driver.find_element(By.XPATH, '//*[@id="input-search"]')
    search.send_keys("notebooks")
    search.submit()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="__next"]/div/main/section[4]/div[4]')
        )
    )

    get_dados_produtos(driver.page_source)

    driver.quit()
    send_mail()


if __name__ == "__main__":
    main()
