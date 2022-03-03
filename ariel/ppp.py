from __future__ import print_function

from google.oauth2.credentials import Credentials
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
from tkinter.messagebox import showinfo
from os.path import dirname, basename
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from datetime import timedelta as td
from datetime import datetime as dt
import win32com.client as win32
from email import encoders
from textwrap import wrap
import logging as log
from os import getenv
import pandas as pd
import pywintypes
import smtplib

de_para_colunas = {
    "Data hora ": r"{{ DATA-HORA }}",
    "Endereço de e-mail": r"{{ DESTINATARIO }}",
    "Escolha uma das unidades da ATIVA MEDICINA:": r"{{ UNIDADE }}",
    "Informe a razão social da EMPRESA:": r"{{ NOME-EMPRESA }}",
    "Informe o nome da empresa:": r"{{ EMPRESA }}",
    "Informe o nome completo do TRABALHADOR:": r"{{ NOME }}",
    "Informe o CARGO do TRABALHADOR:": r"{{ CARGO }}",
    "Informe o CBO do CARGO do TRABALHADOR:": r"{{ CBO }}",
    "Selecione uma das alternativas abaixo relacionadas ao TRABALHADOR:": r"{{ BR-PDH }}",
    "Informe o NIT/PIS do TRABALHADOR:": r"{{ NIT-PIS }}",
    "Informe a DATA DE NASCIMENTO do TRABALHADOR:": r"{{ DT-NASCIMENTO }}",
    "Informe o sexo do TRABALHADOR:": r"{{ SEXO }}",
    "Informe o número da CTPS do TRABALHADOR  (Número, Série e UF) : ": r"{{ CTPS }}",
    "Informe a data de ADMISSÃO do TRABALHADOR:": r"{{ DT-ADMISSAO }}",
    "Informe a data da DEMISSÃO do trabalhador:": r"{{ DT-DEMISSAO }}",
    "Informe o regime de revezamento do trabalhador:": r"{{ REGIME-REVEZAMENTO }}",
    "Caso tenha emitido CAT - Comunicação de Acidente de Trabalho para este trabalhador, informe a DATA da CAT: ": r"{{ DT-CAT }}",
    "Caso tenha emitido CAT - Comunicação de Acidente de Trabalho para este trabalhador, informe o número da CAT: ": r"{{ N-CAT }}",
    "Selecione o código da ocorrência GFIP:": r"{{ GFPI }}",
    "Foi tentada a implementação de medidas de proteção coletiva, de caráter administrativo ou de organização do trabalho, optando-se pelo EPI por inviabilidade técnica, insuficiência ou interinidade, ou ainda em caráter complementar ou emergencial?": r"{{ SN1 }}",
    "Foram observadas as condições de funcionamento e do uso ininterrupto do EPI ao longo do tempo, conforme especificação técnica do fabricante, ajustada às condições de campo?": r"{{ SN2 }}",
    "Foi observado o prazo de validade, conforme Certificado de Aprovação-CA do MTE?": r"{{ SN3 }}",
    "Foi observada a periodicidade de troca definida pelos programas ambientais, comprovada mediante recibo assinado pelo usuário em época própria?": r"{{ SN4 }}",
    "Foi observada a higienização do EPI?": r"{{ SN5 }}",
    "Informe o nome completo do REPRESENTANTE LEGAL da empresa:": r"{{ NOME-REPRESENTANTE }}",
    "Informe o NIT/PIS do REPRESENTANTE LEGAL da empresa:": r"{{ NIT-PIS-REPRESENTANTE }}",
    "Informe o CARGO do REPRESENTANTE LEGAL da empresa:": r"{{ CARGO-REPRESENTANTE }}",
    "Informe o CNPJ da empresa:": r"{{ CNPJ }}",
    "SETOR": r"{{ SETOR }}",
    "Descrição das atividades": r"{{ DESCRICAO-ATIVIDADES }}",
    **{f"{kv}° Tipo": r"{{ TIPO-RISCO }}-PT-" + f"{kv:02d}" for kv in range(1, 13)},
    **{f"{kv}° Fator de Risco": r"{{ FATOR-RISCO }}-PT-" + f"{kv:02d}" for kv in range(1, 13)},
    **{f"{kv}° Intens./Conc.": r"{{ TIPO-EXPOSICAO }}-PT-" + f"{kv:02d}" for kv in range(1, 13)},
    "Técnica Utilizada": r"{{ TECNICA }}",
    "EPC Eficaz (S/N)": r"{{ EPC }}",
    "EPI Eficaz (S/N)": r"{{ EPI }}",
    "CNAE": r"{{ CNAE }}",
    "16.2 NIT TST": r"{{ NIT-PIS-RA }}",
    "18.2 NIT MEDICO": r"{{ NIT-PIS-MB }}",
    "Registro Conselho de Classe (REGISTROS AMBIENTAIS)": r"{{ REGISTRO-RA }}",
    "Nome do Profissional Legalmente Habilitado (REGISTROS AMBIENTAIS)": r"{{ NOME-RESPONSAVEL-RA }}",
    "Registro Conselho de Classe (MONITORAÇÃO BIOLÓGICA)": r"{{ REGISTRO-MB }}",
    "Nome do Profissional Legalmente Habilitado (MONITORAÇÃO BIOLÓGICA)": r"{{ NOME-RESPONSAVEL-MB }}",
    "Data de emissão (geração do PDF) =hoje()": r"{{ REALIZADO }}"
}
log.basicConfig(level = log.INFO, format = "[%(asctime)s] %(levelname)s: %(message)s", datefmt = "%d-%m-%Y %H:%M", handlers = [
    log.FileHandler("deskrobotlog.log"),
    log.StreamHandler()
])

class DeskRobot:
    def __init__(self, email, senha) -> None:
        log.info("Iniciando BOT.")
        self.__email = email
        self.__senha = senha
        try:
            self.__word_app = win32.Dispatch("Word.Application")
            self.__excel_app = win32.Dispatch("Excel.Application")
            self.__word_app.Visible = False
            self.__sheets_service = build('sheets', 'v4', credentials = self.__connect())
            self.__email_server = smtplib.SMTP("smtp.gmail.com: 587")
            self.__email_server.starttls()
            self.__email_server.login(self.__email, self.__senha)
        except pywintypes.com_error as word_error:
            log.error("Problemas ao tentar abrir o word.")
            log.error(str(word_error))
        except smtplib.SMTPAuthenticationError as email_error:
            log.error("Problemas ao tentar logar no email.")
            log.error(str(email_error))

        self.__id_worksheet = "1yHq1t_ZiePFEJdZmADL6GI4Zt3w3ZpGUdcV8o7AmEAU"
        self.__doc_template = f"{dirname(__file__)}/ppp-template.docx"
        self.__address = "tratamento de dados!A:BY"
        self.__df = None

    def __del__(self):
        try:
            self.__word_app.Quit()
            self.__excel_app.Quit()
            self.__email_server.quit()
        except:
            pass

    def __connect(self):
        SCOPES = [ 'https://www.googleapis.com/auth/spreadsheets.readonly' ]
        open("token.json", "w").write(str({
                    "token": f'{getenv("G_TOKEN")}',
                    "refresh_token": f'{getenv("G_REFRESH_TOKEN")}',
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "client_id": f'{getenv("G_CLIENT_ID")}',
                    "client_secret": f'{getenv("G_CLIENT_SECRET")}',
                    "scopes": SCOPES,
                    "expiry": f"{(dt.now() + td(hours = 5)).isoformat()}Z"
                }).replace("\'", "\""))
        return Credentials.from_authorized_user_file("token.json", SCOPES)

    def __get_worksheet_info(self):
        log.info("Carregando informações da planilha google.")
        self.__df = pd.DataFrame(self.__sheets_service.spreadsheets().values().get(
                spreadsheetId = self.__id_worksheet, range = self.__address).execute()["values"])
        self.__df.columns = self.__df.iloc[0]
        self.__df = self.__df.fillna("").rename(columns = de_para_colunas).drop(self.__df.index[0])
        self.__df = self.__df[self.__df[r"{{ REALIZADO }}"] == ""]

    def __get_attach(self, attach_path):
        log.info(f"Anexando arquivo '{basename(attach_path)}'.")
        with open(attach_path, "rb") as arquivo_anexo:
            att = MIMEBase("application", "octet-stream")
            att.set_payload(arquivo_anexo.read())
        encoders.encode_base64(att)
        att.add_header("Content-Disposition", f"attachment; filename= {basename(attach_path)}")
        return att

    def execute(self):
        self.__get_worksheet_info()
        log.info(f"Pendentes de envio: {self.__df.shape[0]}")
        print(self.__df)
        for _, row in self.__df.iterrows():
            try:
                log.info("Abrindo template.")
                word_doc = self.__word_app.Documents.Open(self.__doc_template)
                excel_doc = self.__excel_app.Workbooks.Open(Filename=f"{dirname(__file__)}\ppp.xlsm", ReadOnly=1)
                self.__excel_app.Visible = False

                template = word_doc.Content.Find
                for column in self.__df.columns:
                    if column in [ r"{{ DESCRICAO-ATIVIDADES }}" ]:
                        info_desc = row[column] if len(row[column]) > 10 else "-" * 100
                        for p, parte in enumerate(wrap(info_desc, int(len(info_desc) / 9))):
                            log.info(f"Substituindo {column}-PT-{p} <--> {parte}")
                            self.__excel_app.Application.Run("ppp.xlsm!ppp.replace_info", f"{column}-PT-{p}", parte, template)
                        self.__excel_app.Application.Run("ppp.xlsm!ppp.replace_info", r"{{ DESCRICAO-ATIVIDADES }}-PT-5", "", template)
                    else:
                        self.__excel_app.Application.Run("ppp.xlsm!ppp.replace_info", column, row[column], template)
                        log.info(f"Substituindo {column} <--> {row[column]}")

                self.__excel_app.Application.Run("ppp.xlsm!ppp.replace_info", r"{{ DT-HOJE }}", dt.now().strftime("%d/%m/%Y"), template)
                log.info(f"Inserindo data de hoje: {dt.now().strftime('%d/%m/%Y')}")
                pathfilename = f"{dirname(__file__)}\documents\{row[r'{{ NOME }}'].lower()}"
                log.info(f"Salvando arquivos.")
                word_doc.SaveAs(f"{pathfilename}.docx")
                word_doc.SaveAs(f"{pathfilename}.pdf", FileFormat = 17)
                word_doc.Close(False)
                excel_doc.Close(False)

            except Exception as problema:
                log.error("Problemas ao tentar substituir as informações.")
                log.error(str(problema))
                log.error(f"Verificar -> {row[r'{{ NOME }}']}.")
                continue
            try:
                log.info("Configurando email.")
                email = MIMEMultipart()
                email["Subject"] = f"PPP ({row[r'{{ NOME }}']})"
                email["From"] = self.__email
                email["To"] = row[r"{{ DESTINATARIO }}"]
                email.attach(MIMEText("Prezado(a) Cliente,<br>Conforme solicitado segue PPP em anexo.<br>Atenciosamente.", "html"))
                email.attach(self.__get_attach(f"{pathfilename}.pdf"))
                log.info("Enviando email.")
                self.__email_server.sendmail(email["From"], [ email["To"] ], email.as_string())

            except Exception as problema:
                log.error("Problemas ao tentar enviar email.")
                log.error(str(problema))
                log.error(f"Verificar -> {row[r'{{ NOME }}']}.")

        showinfo("Sucesso", "Processo concluído!")

if "__main__" == __name__:
    bot = DeskRobot(str(input("Email: ")), str(input("Senha: ")))
    bot.execute()
    log.info("Encerrando BOT.")