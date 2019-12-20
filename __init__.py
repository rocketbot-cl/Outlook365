# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"
    
    pip install <package> -t .

"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import imaplib
import email
import os
import sys
from email.utils import make_msgid
from textwrap import dedent

from bs4 import BeautifulSoup
import base64
import re

base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'Outlook365' + os.sep + 'libs' + os.sep
# print(cur_path)
sys.path.append(cur_path)
# print(cur_path )

from mailparser import mailparser

global fromaddr
global server
global password
global mail
global id_

"""
    Obtengo el modulo que fue invocado
"""
module = GetParams("module")


def parse_uid(tmp):
    pattern_uid = re.compile('\d+ \(UID (?P<uid>\d+)\)')
    print('tmp', tmp)
    try:
        tmp = tmp.decode()
    except:
        pass
    match = pattern_uid.match(tmp)
    return match.group('uid')


if module == "conf_mail":

    conx = ""

    try:
        fromaddr = GetParams('from')
        password = GetParams('password')
        var_ = GetParams('var_')

        mail = imaplib.IMAP4_SSL('outlook.office365.com')
        mail.login(fromaddr, password)
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(fromaddr, password)
        conx = True
    except:
        PrintException()
        conx = False

    SetVar(var_, conx)

if module == "send_mail":
    to = GetParams('to')
    subject = GetParams('subject')
    body_ = GetParams('body')
    attached_file = GetParams('attached_file')
    print(to, subject, body_, attached_file)

    try:
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = to
        msg['Subject'] = subject
        body = body_
        msg.attach(MIMEText(body, 'html'))

        if attached_file:
            if os.path.exists(attached_file):
                filename = os.path.basename(attached_file)
                attachment = open(attached_file, "rb")
                part = MIMEBase('application', 'octet-stream')
                part.set_payload((attachment).read())
                attachment.close()
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                msg.attach(part)
        text = msg.as_string()
        server.sendmail(fromaddr, to.split(","), text)
        server.close()


    except Exception as e:
        PrintException()
        raise e

if module == "get_mail":
    filtro = GetParams('filtro')
    var_ = GetParams('var_')

    mail = imaplib.IMAP4_SSL('outlook.office365.com')
    mail.login(fromaddr, password)
    mail.list()
    # Out: list of "folders" aka labels in gmail.
    mail.select("inbox")  # connect to inbox.

    if filtro and len(filtro) > 0:
        result, data = mail.search(None, filtro, "ALL")
    else:
        result, data = mail.search(None, "ALL")

    ids = data[0]  # data is a list.
    id_list = ids.split()  # ids is a space separated string
    lista = [b.decode() for b in id_list]

    # print('ID',id_list)
    SetVar(var_, lista)

if module == "get_unread":
    filtro = GetParams('filtro')
    var_ = GetParams('var_')

    mail = imaplib.IMAP4_SSL('outlook.office365.com')
    mail.login(fromaddr, password)
    mail.list()
    # Out: list of "folders" aka labels in gmail.
    mail.select("inbox")  # connect to inbox.

    if filtro and len(filtro) > 0:
        result, data = mail.search(None, filtro, "UNSEEN")
    else:
        result, data = mail.search(None, "UNSEEN")

    ids = data[0]  # data is a list.
    id_list = ids.split()  # ids is a space separated string

    lista = [b.decode() for b in id_list]

    # print('ID',id_list)
    SetVar(var_, lista)

if module == "read_mail":
    id_ = GetParams('id_')
    var_ = GetParams('var_')
    att_folder = GetParams('att_folder')

    mail = imaplib.IMAP4_SSL('outlook.office365.com')
    mail.login(fromaddr, password)
    mail.select("inbox")

    # mail.select()
    typ, data = mail.fetch(id_, '(RFC822)')
    raw_email = data[0][1]
    # converts byte literal to string removing b''
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)

    mail_ = mailparser.parse_from_string(raw_email_string)

    try:

        bs = BeautifulSoup(mail_.body, 'html.parser').body.get_text()
    except:
        bs = mail_.body

    bs = bs.split('--- mail_boundary ---')[0]
    print(bs)
    nameFile = []

    for att in mail_.attachments:
        name_ = att['filename']
        nameFile.append(name_)

        fileb = att['payload']
        cont = base64.b64decode(fileb)
        if att_folder:
            with open(os.path.join(att_folder, name_), 'wb') as file_:
                file_.write(cont)
                file_.close()

    final = {"date": mail_.date.__str__(), 'subject': mail_.subject, 'from': ", ".join([b for (a, b) in mail_.from_]),
             'to': ", ".join([b for (a, b) in mail_.to]), 'body': bs, 'files': nameFile}

    SetVar(var_, final)

if module == "reply_email":

    try:
        id_ = GetParams('id_')
        body_ = GetParams('body')
        attached_file = GetParams('attached_file')
        print(body_, attached_file)

        mail = imaplib.IMAP4_SSL('outlook.office365.com')
        mail.login(fromaddr, password)
        mail.select("inbox")
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(fromaddr, password)

        # mail.select()
        typ, data = mail.fetch(id_, '(RFC822)')
        raw_email = data[0][1]
        mm = email.message_from_bytes(raw_email)

        # msg = MIMEMultipart()
        # msg.attach(MIMEText(body_, 'plain'))

        #    m_ = create_auto_reply(mm, body_)
        mail__ = MIMEMultipart('alternative')
        mail__['Message-ID'] = make_msgid()
        mail__['References'] = mail__['In-Reply-To'] = mm['Message-ID']
        mail__['Subject'] = 'Re: ' + mm['Subject']
        mail__['From'] = mm['To'] = mm['Reply-To'] or mm['From']
        mail__.attach(MIMEText(dedent(body_), 'plain'))

        if attached_file:
            if os.path.exists(attached_file):
                filename = os.path.basename(attached_file)
                attachment = open(attached_file, "rb")
                part = MIMEBase('application', 'octet-stream')
                part.set_payload((attachment).read())
                attachment.close()
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                mail__.attach(part)

        to_ = mm['From'].split("<")[1].replace(">", "")
        print(fromaddr, to_)
        server.sendmail(fromaddr, to_, mail__.as_bytes())
        server.close()
        mail.logout()
    except:
        PrintException()

if module == "create_folder":
    try:
        folder_name = GetParams('folder_name')
        host = "outlook.office365.com"
        mail = imaplib.IMAP4_SSL(host)
        mail.login(fromaddr, password)
        mail.create(folder_name)
    except:
        PrintException()

if module == "move_mail":
    # imap = GetGlobals('email')
    id_ = GetParams("id_")
    label_ = GetParams("label_")
    var = GetParams("var")

    if not id_:
        raise Exception("No ha ingresado ID de email a mover")
    if not label_:
        raise Exception("No ha ingresado carpeta de destino")
    try:
        # login on IMAP server
        # if imap.IMAP_SSL:
        #     mail = imaplib.IMAP4_SSL('outlook.office365.com')
        # else:
        #     mail = imaplib.IMAP4('outlook.office365.com')

        mail = imaplib.IMAP4_SSL('outlook.office365.com')
        mail.login(fromaddr, password)
        mail.select('inbox', readonly=False)
        resp, data = mail.fetch(id_, "(UID)")
        msg_uid = parse_uid(data[0])

        result = mail.uid('COPY', msg_uid, label_)

        if result[0] == 'OK':
            mov, data = mail.uid('STORE', msg_uid, '+FLAGS', '(\Deleted)')
            res = mail.expunge()
            if var:
                ret = True if res[0] == 'OK' else False
                SetVar(var, ret)
        else:
            raise Exception(result)
    except Exception as e:
        raise e
