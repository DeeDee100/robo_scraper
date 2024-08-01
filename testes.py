import smtplib
import requests
import pandas as pd
from bs4 import BeautifulSoup as bs
from email.mime.text import MIMEText
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

def capturar_dados_produtos(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0'}

    response = requests.get(url, headers=headers)
    soup = bs(response.content, 'html.parser')
    div_produtos = soup.find(class_="sc-fBWQRz cULVBz sc-fulCBj fxxByy sc-hsUFQk ceNiJh")
    nomes_produtos = [tags.contents[0] for tags in div_produtos.find_all(class_="sc-doohEh dHamKz")]
    qtd_avaliacoes = [tags.contents[0] for tags in div_produtos.find_all(class_="sc-epqpcT jdMYPv")]
    urls_produtos = [f"https://www.magazineluiza.com.br{link['href']}" for link in div_produtos.find_all(class_ = 'sc-eBMEME uPWog sc-dxUMQK jeUYOh sc-dxUMQK jeUYOh')]

    # breakpoint()
    dados_produtos = []
    for nome, avaliacao, url in zip(nomes_produtos, qtd_avaliacoes, urls_produtos):
        produto = {
            "PRODUTO": nome.text.strip(),
            "QTD_AVAL": int(avaliacao.text.strip().split()[1][1:-1]),
            "URL": url
        }
        dados_produtos.append(produto)

    # breakpoint()
    df = pd.DataFrame(dados_produtos)
    df = df.loc[~df['QTD_AVAL'].eq(0)]
    df_piores = df.loc[df['QTD_AVAL'].between(0, 100)].sort_values('QTD_AVAL', ascending=False)
    df_melhores = df.loc[~df['QTD_AVAL'].between(0, 100)].sort_values('QTD_AVAL', ascending=False)
    with pd.ExcelWriter("output\output.xlsx") as writer:
        df_melhores.to_excel(writer, sheet_name="Melhores", index=False)
        df_piores.to_excel(writer, sheet_name="Piores", index=False)

    return dados_produtos

# Exemplo de uso:
# url_magazine_luiza = "https://www.magazineluiza.com.br/busca/notebooks/?from=submit"
# dados = capturar_dados_produtos(url_magazine_luiza)
# for produto in dados:
#     print(produto)
    


# send_email()

# Set up email account information
# email_address = "testerobo55@gmail.com"
# email_password = "ozpi owip wonz luhv"
# smtp_server = "smtp.gmail.com"
# smtp_port = 587


# msg = EmailMessage()
# msg['Subject'] = 'Relatório Notebooks!'
# msg['From'] = email_address
# msg['To'] = 'testerobo55@gmail.com'
# msg.set_content('Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.\n\nAtenciosamente,\nRobô')
# with open('output\\Notebook.xlsx', 'rb') as f:
#     file_data = f.read()
# msg.add_attachment(file_data, ntype="application", subtype="octet-stream", filename='Notebook.xlsx')


# with smtplib.SMTP(smtp_server, smtp_port) as server:
#     server.starttls()
#     server.login(email_address, email_password)
#     server.send_message(msg)




def send_mail():
    
    email_address = "testerobo55@gmail.com"
    email_password = "ozpi owip wonz luhv"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = email_address
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = 'Relatório Notebooks'
    
    content = 'Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.\n\nAtenciosamente,\nRobô'
    msg.attach(MIMEText(content))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("output\\Notebook.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="Notebook.xlsx"')
    msg.attach(part)

    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(email_address,email_password)
    smtp.sendmail(email_address, email_address, msg.as_string())
    smtp.quit()
    
send_mail()