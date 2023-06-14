from smtplib import SMTP_SSL, SMTP_SSL_PORT
from email.mime.multipart import MIMEMultipart, MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pandas as pd
import PySimpleGUI as sg
import os
from os.path import exists

sg.theme('LightBrown13')

element0 = [
    [sg.Text("Selecione a planilha:", font=('Ubuntu', 12))],
    [sg.InputText(key="-FILE_PATH-", font=('Ubuntu', 10)),
     sg.FileBrowse("Procurar", font=('Ubuntu', 10), initial_folder=os.getcwd())],
]

element1 = [
    [sg.Text("Email"), sg.Input(key='email', size=(20, 1))],
    [sg.Text("Senha"), sg.InputText(
        key='senha', password_char='*', size=(20, 1))],
    [sg.Button("Enviar")]
]

element2 = [
    [sg.Radio('Renovação', 'group 1',
              key='Renovacao', enable_events=True)],
    [sg.Radio('Recuperacao', 'group 1',
              key='Recuperacao', enable_events=True)],
    [sg.Radio('Sinercon', 'group 1',
              key='Sinercon', enable_events=True)],
    [sg.Radio('Passados', 'group 1',
              key='Passados', enable_events=True)],
    [sg.Radio('Novos', 'group 1',
              key='Novos', enable_events=True)],
]

element3 = [
    [sg.Text("Titulo do email", font=('Ubuntu', 12)),
     sg.InputText(key="titulo")],
]

layout = [
    [sg.Column(element0, vertical_alignment="left",
               element_justification='left')],
    [sg.HorizontalSeparator()],
    [sg.Column(element1, element_justification='left'), sg.VerticalSeparator(
    ), sg.Column(element2, vertical_alignment="left", element_justification='left')],
    [sg.HorizontalSeparator()],
    [sg.Column(element3, element_justification='left')],
]

# Definir o caminho para o arquivo de ícone
icon_path = './logo-disparador.ico'

janela = sg.Window('Disparador de e-mail', layout, icon=icon_path)

def isNaN(value):
    try:
        import math
        return math.isnan(float(value))
    except:
        return False


def Novos(email, senha, contatos, titulo):
    for i, message in enumerate(contatos["EMAIL"]):
        nome = contatos.loc[i, "NOME"]  # produto aqui pedrão
        protocolo = contatos.loc[i, "PROTOCOLO"]  # protocolo aqui pedrão
        produto = contatos.loc[i, "PRODUTO"]  # produto aqui pedrovski
        formattedDatetime = contatos.loc[i, "VENCIMENTO"]  # data e hora tbm po

        with open("LogEnviados.txt", "a") as arquivo:
            email_enviar = contatos.loc[i, "EMAIL"]
            if not isNaN(email_enviar):
                email_message = MIMEMultipart()
                email_message.add_header('To', email_enviar)
                email_message.add_header('From', email)
                email_message.add_header('Subject', titulo)
                email_message.add_header('X-Priority', '1')
                mensage = f"""<center>
                <html>
  <head>
    <base target="_top" />
    <style>
      @import url("https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;500;600;700;800&display=swap");
    </style>
  </head>
  <center>
    <header>
      <div
        style="
          margin-bottom: 20px;
          background-color: #006c98;
          width: 620px;
          border-radius: 20px;
        "
      >
        <a href="https://webcertificados.com.br/"
          ><img
            id="logo-header"
            src="https://arquivos.webcertificados.com/wl/?id=Tm9ZFTQdn8BfQTuVhhswApdHGgRx5xIY&fmode=open"
            alt="Logo WebCertificados"
            width="300"
        /></a>
      </div>
    </header>
    <body>
      <div
        style="
          border-radius: 20px;
          background-color: #ecf8ff;
          height: 540px;
          width: 600px;
          line-height: 24px;
          padding: 10px;
          font-family: 'Open Sans', sans-serif;
          font-size: 18px;
          -ms-text-size-adjust: 100%;
          -webkit-text-size-adjust: 100%;
          word-break: break-word;
        "
      >
        <div
          style="
            font-family: 'Open Sans', sans-serif;
            font-size: 18px;
            line-height: 0.6;
            color: rgb(24, 24, 24);
          "
        >
          <h3>Prezado(a), {nome}</h3>
          <br />
          <h3>A WebCertificados agradece a preferência e</h3>
          <h3>nos colocamos à disposição para sanar suas dúvidas.</h3>
          <h3>Por ser nosso cliente, você tem outros benefícios.</h3>
          <h3>Clique no botão abaixo e descubra!</h3>
        </div>
        <div
          style="
            font-family: 'Open Sans', sans-serif;
            font-size: 18px;
            line-height: 0.6;
            margin-top: 50px;
            color: rgb(24, 24, 24);
          "
        >
          <h3>Detalhes do seu pedido</h3>
          <h3>Protocolo: {protocolo}</h3>
          <h3>Produto: {produto}</h3>
          <h3>Emissão: {formattedDatetime}</h3>
          <div style="margin-top: 50px">
            <a
              href="https://api.whatsapp.com/send?phone=553232175259&text=Ol%C3%A1,%20quero%20benefícios!"
              ><button
                style="
                  font-size: 18px;
                  border: none;
                  color: #fff;
                  width: 20vw;
                  height: 5vh;
                  background: linear-gradient(120deg, #006c98, #00445f);
                  border-radius: 20px;
                "
              >
                CLIQUE AQUI!
              </button></a
            >
            <h3 style="margin-top: 15px">Renove e ganhe benefícios!</h3>
            <h3 style="margin-top: -5px; font-size: 12px;"> Dúvidas? Nosso Telefone: <a href="tel:+553232175259" style="color: rgb(24, 24, 24);">(32) 3217-5259</a></h5>
          </div>
        </div>
      </div>
    </body>
  </center>
</html>
</center>"""
                html_part = MIMEText(mensage, 'html')

                email_message.attach(html_part)

                smtp_server = SMTP_SSL(
                    'sh-pro46.hostgator.com.br', port=SMTP_SSL_PORT)
                smtp_server.set_debuglevel(1)
                smtp_server.login(email, senha)
                smtp_server.sendmail(email, email_enviar,
                                     email_message.as_bytes())
                arquivo.write(
                    f"Email enviado com sucesso para: {email_enviar}\n")
                smtp_server.quit()


def Passados(email, senha, contatos, titulo):
    for i, message in enumerate(contatos["EMAIL"]):
        nome = contatos.loc[i, "NOME"]  # produto aqui pedrão
        protocolo = contatos.loc[i, "PROTOCOLO"]  # protocolo aqui pedrão
        produto = contatos.loc[i, "PRODUTO"]  # produto aqui pedrovski
        formattedDatetime = contatos.loc[i, "VENCIMENTO"]  # data e hora tbm po

        with open("LogEnviados.txt", "a") as arquivo:
            email_enviar = contatos.loc[i, "EMAIL"]
            if not isNaN(email_enviar):
                email_message = MIMEMultipart()
                email_message.add_header('To', email_enviar)
                email_message.add_header('From', email)
                email_message.add_header('Subject', titulo)
                email_message.add_header('X-Priority', '1')
                mensage = f"""<center>
                <html>
  <head>
    <base target="_top" />
    <style>
      @import url("https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;500;600;700;800&display=swap");
    </style>
  </head>
  <center>
    <header>
      <div
        style="
          margin-bottom: 20px;
          background-color: #006c98;
          width: 620px;
          border-radius: 20px;
        "
      >
        <a href="https://webcertificados.com.br/"
          ><img
            id="logo-header"
            src="https://arquivos.webcertificados.com/wl/?id=Tm9ZFTQdn8BfQTuVhhswApdHGgRx5xIY&fmode=open"
            alt="Logo WebCertificados"
            width="300"
        /></a>
      </div>
    </header>
    <body>
      <div
        style="
          border-radius: 20px;
          background-color: #ecf8ff;
          height: 540px;
          width: 600px;
          line-height: 24px;
          padding: 10px;
          font-family: 'Open Sans', sans-serif;
          font-size: 18px;
          -ms-text-size-adjust: 100%;
          -webkit-text-size-adjust: 100%;
          word-break: break-word;
        "
      >
        <div
          style="
            font-family: 'Open Sans', sans-serif;
            font-size: 18px;
            line-height: 0.6;
            color: rgb(24, 24, 24);
          "
        >
          <h3>Prezado(a), {nome}</h3>
          <br />
          <h3>Seu certificado digital expirou e</h3>
          <h3>você pode ficar sem faturar.</h3>
          <h3>Faça sua renovação hoje</h3>
          <h3>de forma rápida e totalmente online!,</h3>
          <h3>Clique no botão abaixo e renove!</h3>
        </div>
        <div
          style="
            font-family: 'Open Sans', sans-serif;
            font-size: 18px;
            line-height: 0.6;
            margin-top: 50px;
            color: rgb(24, 24, 24);
          "
        >
          <h3>Detalhes do seu pedido</h3>
          <h3>Protocolo: {protocolo}</h3>
          <h3>Produto: {produto}</h3>
          <h3>Vencimento: {formattedDatetime}</h3>
          <div style="margin-top: 50px">
            <a
              href="https://api.whatsapp.com/send?phone=553232175259&text=Ol%C3%A1,%20quero%20voltar%20a%20ser%20WebCertificados!"
              ><button
                style="
                  font-size: 18px;
                  border: none;
                  color: #fff;
                  width: 20vw;
                  height: 5vh;
                  background: linear-gradient(120deg, #006c98, #00445f);
                  border-radius: 20px;
                "
              >
                CLIQUE AQUI!
              </button></a
            >
            <h3 style="margin-top: 15px">Renove e ganhe benefícios!</h3>
            <h3 style="margin-top: -5px; font-size: 12px;"> Dúvidas? Nosso Telefone: <a href="tel:+553232175259" style="color: rgb(24, 24, 24);">(32) 3217-5259</a></h5>
          </div>
        </div>
      </div>
    </body>
  </center>
</html>
</center>"""
                html_part = MIMEText(mensage, 'html')

                email_message.attach(html_part)

                smtp_server = SMTP_SSL(
                    'sh-pro46.hostgator.com.br', port=SMTP_SSL_PORT)
                smtp_server.set_debuglevel(1)
                smtp_server.login(email, senha)
                smtp_server.sendmail(email, email_enviar,
                                     email_message.as_bytes())
                arquivo.write(
                    f"Email enviado com sucesso para: {email_enviar}\n")
                smtp_server.quit()


def Recuperacao(email, senha, contatos, titulo):
    for i, message in enumerate(contatos["EMAIL"]):
        nome = contatos.loc[i, "NOME"]  # produto aqui pedrão

        with open("LogEnviados.txt", "a") as arquivo:
            email_enviar = contatos.loc[i, "EMAIL"]
            if not isNaN(email_enviar):
                email_message = MIMEMultipart()
                email_message.add_header('To', email_enviar)
                email_message.add_header('From', email)
                email_message.add_header('Subject', titulo)
                email_message.add_header('X-Priority', '1')
                mensage = f"""<center>
                <html>
  <head>
    <base target="_top" />
    <style>
      @import url("https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;500;600;700;800&display=swap");
    </style>
  </head>
  <center>
    <header>
      <div
        style="
          margin-bottom: 20px;
          background-color: #006c98;
          width: 620px;
          border-radius: 20px;
        "
      >
        <a href="https://webcertificados.com.br/"
          ><img
            id="logo-header"
            src="https://arquivos.webcertificados.com/wl/?id=Tm9ZFTQdn8BfQTuVhhswApdHGgRx5xIY&fmode=open"
            href="https://webcertificados.com.br/"
            alt="Logo WebCertificados"
            width="300"
        /></a>
      </div>
    </header>
    <body>
      <div
        style="
          border-radius: 20px;
          background-color: #ecf8ff;
          height: 530px;
          width: 600px;
          padding: 10px;
          -ms-text-size-adjust: 100%;
          -webkit-text-size-adjust: 100%;
          word-break: break-word;
        "
      >
        <div
          style="
            font-family: 'Open Sans', sans-serif;
            font-size: 18px;
            line-height: 0.6;
            color: rgb(24, 24, 24);
          "
        >
          <h3>Olá, {nome}</h3>
          <br />
          <h3>Há algum tempo você fez a emissão do seu certificado</h3>
          <h3>digital e não retornou para renová-lo. <b>Temos preços</b></h3>
          <h3><b>promocionais</b> a partir de R$99,99 para pessoa física</h3>
          <h3>e de R$149,99 para pessoa jurídica.</h3>
        </div>
        <div
          style="
            font-family: 'Open Sans', sans-serif;
            font-size: 18px;
            line-height: 0.6;
            margin-top: 50px;
            color: rgb(24, 24, 24);
          "
        >
          <h3>Além disso, temos outros benefícios para você,</h3>
          <h3>como plano corporativo de consulta de crédito</h3>
          <h3>e portal assinador.</h3>
          <br />
          <h3>Clique no botão abaixo e saiba mais.</h3>
          <div style="margin-top: 20px">
            <a
              href="https://api.whatsapp.com/send?phone=553232175259&text=Ol%C3%A1,%20quero%20voltar%20a%20ser%20WebCertificados!"
              ><button
                style="
                  font-size: 18px;
                  border: none;
                  color: #fff;
                  width: 16vw;
                  height: 5vh;
                  background: linear-gradient(
                    120deg,
                    #006c98,
                    #00212e
                  );
                  border-radius: 20px;
                  z-index: 1;
                "
              >
                Clique para saber mais!
              </button></a
            >
            <h3 style="font-size: 12px;"> Dúvidas? Nosso Telefone: <a href="tel:+553232175259" style="color: rgb(24, 24, 24);">(32) 3217-5259</a></h5>
          </div>
        </div>
      </div>
    </body>
  </center>
</html>
</center>"""
                html_part = MIMEText(mensage, 'html')

                email_message.attach(html_part)

                smtp_server = SMTP_SSL(
                    'sh-pro46.hostgator.com.br', port=SMTP_SSL_PORT)
                smtp_server.set_debuglevel(1)
                smtp_server.login(email, senha)
                smtp_server.sendmail(email, email_enviar,
                                     email_message.as_bytes())
                arquivo.write(
                    f"Email enviado com sucesso para: {email_enviar}\n")
                smtp_server.quit()


def Renovacao(email, senha, contatos, titulo):
    for i, message in enumerate(contatos["EMAIL"]):
        nome = contatos.loc[i, "NOME"]  # produto aqui pedrão
        protocolo = contatos.loc[i, "PROTOCOLO"]  # protocolo aqui pedrão
        produto = contatos.loc[i, "PRODUTO"]  # produto aqui pedrovski
        vencimento = contatos.loc[i, "VENCIMENTO"]  # vencimento tbm claro

        with open("LogEnviados.txt", "a") as arquivo:
            email_enviar = contatos.loc[i, "EMAIL"]
            if not isNaN(email_enviar):
                email_message = MIMEMultipart()
                email_message.add_header('To', email_enviar)
                email_message.add_header('From', email)
                email_message.add_header('Subject', titulo)
                email_message.add_header('X-Priority', '1')
                mensage = f"""<center>
                <html>
          <head>
            <base target="_top" />
            <style>
              @import url("https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;500;600;700;800&display=swap");
            </style>
          </head>
          <center>
            <header>
              <div
                style="
                  margin-bottom: 20px;
                  background-color: #006c98;
                  width: 620px;
                  border-radius: 20px;
                "
              >
                <a href="https://webcertificados.com.br/"
                  ><img
                    id="logo-header"
                    src="https://arquivos.webcertificados.com/wl/?id=Tm9ZFTQdn8BfQTuVhhswApdHGgRx5xIY&fmode=open"
                    alt="Logo WebCertificados"
                    width="300"
                /></a>
              </div>
            </header>
            <body>
              <div
                style="
                  border-radius: 20px;
                  background-color: #ecf8ff;
                  height: 540px;
                  width: 600px;
                  line-height: 24px;
                  padding: 10px;
                  font-family: 'Open Sans', sans-serif;
                  font-size: 18px;
                  -ms-text-size-adjust: 100%;
                  -webkit-text-size-adjust: 100%;
                  word-break: break-word;
                "
              >
                <div
                  style="
                    font-family: 'Open Sans', sans-serif;
                    font-size: 18px;
                    line-height: 0.6;
                    color: rgb(24, 24, 24);
                  "
                >
                  <h3>Prezado(a), {nome}</h3>
                  <br />
                  <h3>Seu certificado digital irá vencer em breve e</h3>
                  <h3>para evitar dor de cabeça renove-o agora</h3>
                  <h3>de forma rápida e totalmente online</h3>
                  <h3>Clique no botão abaixo!</h3>
                </div>
                <div
                  style="
                    font-family: 'Open Sans', sans-serif;
                    font-size: 18px;
                    line-height: 0.6;
                    margin-top: 50px;
                    color: rgb(24, 24, 24);
                  "
                >
                  <h3>Detalhes do seu pedido</h3>
                  <h3>Protocolo: {protocolo}</h3>
                  <h3>Produto: {produto}</h3>
                  <h3>Vencimento: {vencimento}</h3>
                  <div style="margin-top: 50px">
                    <a
                      href="https://api.whatsapp.com/send?phone=553232175259&text=Ol%C3%A1,%20quero%20renovar%20agora!"
                      ><button
                        style="
                          font-size: 18px;
                          border: none;
                          color: #fff;
                          width: 20vw;
                          height: 5vh;
                          background: linear-gradient(120deg, #006c98, #00445f);
                          border-radius: 20px;
                        "
                      >
                        CLIQUE AQUI!
                      </button></a
                    >
                    <h3 style="margin-top: 15px">Renove e ganhe benefícios!</h3>
                    <h3 style="margin-top: -5px; font-size: 12px;"> Dúvidas? Nosso Telefone: <a href="tel:+553232175259" style="color: rgb(24, 24, 24);">(32) 3217-5259</a></h5>
                  </div>
                </div>
              </div>
            </body>
          </center>
        </html>
        </center>"""
                html_part = MIMEText(mensage, 'html')

                email_message.attach(html_part)

                smtp_server = SMTP_SSL(
                    'sh-pro46.hostgator.com.br', port=SMTP_SSL_PORT)
                smtp_server.set_debuglevel(1)
                smtp_server.login(email, senha)
                smtp_server.sendmail(email, email_enviar,
                                     email_message.as_bytes())
                arquivo.write(
                    f"Email enviado com sucesso para: {email_enviar}\n")
                smtp_server.quit()


def Sinercon(email, senha, contatos, titulo):
    for i, message in enumerate(contatos["EMAIL"]):
        nome = contatos.loc[i, "NOME"]  # produto aqui pedrão
        protocolo = contatos.loc[i, "PROTOCOLO"]  # protocolo aqui pedrão
        produto = contatos.loc[i, "PRODUTO"]  # produto aqui pedrovski
        formattedDatetime = contatos.loc[i, "VENCIMENTO"]  # data e hora tbm po

        with open("LogEnviados.txt", "a") as arquivo:
            email_enviar = contatos.loc[i, "EMAIL"]
            if not isNaN(email_enviar):
                email_message = MIMEMultipart()
                email_message.add_header('To', email_enviar)
                email_message.add_header('From', email)
                email_message.add_header('Subject', titulo)
                email_message.add_header('X-Priority', '1')
                mensage = f"""<center>
                <html>
  <head>
    <base target="_top" />
    <style>
      @import url("https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;500;600;700;800&display=swap");
    </style>
  </head>
  <center>
    <header>
      <div
        style="
          margin-bottom: 20px;
          background-color: #006c98;
          width: 620px;
          border-radius: 20px;
        "
      >
        <a href="https://webcertificados.com.br/"
          ><img
            id="logo-header"
            src="https://arquivos.webcertificados.com/wl/?id=Tm9ZFTQdn8BfQTuVhhswApdHGgRx5xIY&fmode=open"
            alt="Logo WebCertificados"
            width="300"
        /></a>
      </div>
    </header>
    <body>
      <div
        style="
          border-radius: 20px;
          background-color: #ecf8ff;
          height: 550px;
          width: 600px;
          padding: 10px;
          font-family: 'Open Sans', sans-serif;
          font-size: 18px;
          -ms-text-size-adjust: 100%;
          -webkit-text-size-adjust: 100%;
          word-break: break-word;
        "
      >
        <div
          style="
            font-family: 'Open Sans', sans-serif;
            font-size: 16px;
            line-height: 0.6;
            color: rgb(24, 24, 24);
          "
        >
          <h3>Olá, {nome}</h3>
          <br />
          <h3>O seu certificado digital vencerá em poucos dias</h3>
          <h3>e temos uma oferta especial para você.</h3>
          <h3>Se você já fez a renovação, entre em contato conosco</h3>
          <h3>para saber outros benefícios disponíveis.</h3>
          <br />
        </div>
        <div
          style="
            font-family: 'Open Sans', sans-serif;
            font-size: 18px;
            line-height: 0.6;
            margin-top: 20px;
            color: rgb(24, 24, 24);
          "
        >
          <h3>Detalhes do seu pedido</h3>
          <h3>Protocolo: {protocolo}</h3>
          <h3>Produto: {produto}</h3>
          <h3>Vencimento: {formattedDatetime}</h3>
          <br />
          <h3>Clique no botão abaixo e saiba mais.</h3>
          <div style="margin-top: 20px">
            <a
              href="https://api.whatsapp.com/send?phone=553288067065&text=Ol%C3%A1,%20quero%20renovar%20com%20desconto!"
              ><button
                style="
                  font-size: 18px;
                  border: none;
                  color: #fff;
                  width: 22vw;
                  height: 5vh;
                  background: linear-gradient(
                    120deg,
                    #006c98,
                    #00212e
                  );
                  border-radius: 30px;
                  z-index: 1;
                "
              >
                RENOVE AGORA COM DESCONTO
              </button></a
            >
            <h3 style="font-size: 12px;"> Dúvidas? Nosso Telefone: <a href="tel:+553232175259" style="color: rgb(24, 24, 24);">(32) 3217-5259</a></h5>
          </div>
        </div>
      </div>
    </body>
  </center>
</html>
</center>"""
                html_part = MIMEText(mensage, 'html')

                email_message.attach(html_part)

                smtp_server = SMTP_SSL(
                    'sh-pro46.hostgator.com.br', port=SMTP_SSL_PORT)
                smtp_server.set_debuglevel(1)
                smtp_server.login(email, senha)
                smtp_server.sendmail(email, email_enviar,
                                     email_message.as_bytes())
                arquivo.write(
                    f"Email enviado com sucesso para: {email_enviar}\n")
                smtp_server.quit()


while True:
    eventos, valores = janela.read()
    if eventos == sg.WINDOW_CLOSED:
        break
    if eventos == 'Enviar':
        if len(valores["-FILE_PATH-"]) > 1:
            if len(valores["titulo"]) > 1:

                if valores["Novos"]:
                    email = valores["email"]
                    senha = valores["senha"]
                    titulo = valores["titulo"]
                    contatos = pd.read_excel(valores["-FILE_PATH-"])

                    Novos(email, senha, contatos, titulo)
                    sg.popup("Emails enviados com sucesso!",
                             "Dexei um arquivo chamado LogEnviados junto comigo!", "Da uma verificadinha la!")

                elif valores["Passados"]:
                    email = valores["email"]
                    senha = valores["senha"]
                    titulo = valores["titulo"]
                    contatos = pd.read_excel(valores["-FILE_PATH-"])

                    Passados(email, senha, contatos, titulo)
                    sg.popup("Emails enviados com sucesso!",
                             "Deixei um arquivo chamado LogEnviados junto comigo!", "Dá uma verificadinha lá!")

                elif valores["Recuperacao"]:
                    email = valores["email"]
                    senha = valores["senha"]
                    titulo = valores["titulo"]
                    contatos = pd.read_excel(valores["-FILE_PATH-"])

                    Recuperacao(email, senha, contatos, titulo)
                    sg.popup("Emails enviados com sucesso!",
                             "Deixei um arquivo chamado LogEnviados junto comigo!", "Dá uma verificadinha lá!")

                elif valores["Renovacao"]:
                    email = valores["email"]
                    senha = valores["senha"]
                    titulo = valores["titulo"]
                    contatos = pd.read_excel(valores["-FILE_PATH-"])

                    Renovacao(email, senha, contatos, titulo)
                    sg.popup("Emails enviados com sucesso!",
                             "Deixei um arquivo chamado LogEnviados junto comigo!", "Dá uma verificadinha lá!")

                elif valores["Sinercon"]:
                    email = valores["email"]
                    senha = valores["senha"]
                    titulo = valores["titulo"]
                    contatos = pd.read_excel(valores["-FILE_PATH-"])

                    Sinercon(email, senha, contatos, titulo)
                    sg.popup("Emails enviados com sucesso!",
                             "Deixei um arquivo chamado LogEnviados junto comigo!", "Dá uma verificadinha lá!")

            else:
                sg.popup("Coloque um titulo para enviar o email")
        else:
            sg.popup("Preciso saber qual planilha tenho que ler os dados!")
