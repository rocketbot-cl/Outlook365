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

# to ignore reportUndefinedVariable on vscode
SetVar = SetVar  # type: ignore
GetParams = GetParams  # type: ignore
PrintException = PrintException  # type: ignore 
tmp_global_obj = tmp_global_obj  # type: ignore

import os
import sys
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import make_msgid, parsedate_to_datetime
from textwrap import dedent
from typing import Union
base_path = tmp_global_obj["basepath"] #get rocketbot directory
cur_path = os.path.join(base_path, 'modules', 'Outlook365', 'libs')

if cur_path not in sys.path:
    sys.path.append(cur_path)

from mail_common import Mail



"""
    Obtengo el modulo que fue invocado
"""
module = GetParams("module")

global outlook_365

try:
    if module == "conf_mail":

        class Outlook365(Mail):
            def __init__(self, user, pwd, timeout):
                super().__init__(user, pwd, timeout,
                                        smtp_host='smtp.office365.com', smtp_port=587,
                                        imap_host='outlook.office365.com', imap_port=993,
                                        pop_host="outlook.office365.com", pop_port=995)


        fromaddr = GetParams('from')
        password = GetParams('password')
        timeout = GetParams('timeout')
        var_ = GetParams('var_')
        not_imap = GetParams('not_imap')
        
        if timeout is None:
            timeout = 99
        if isinstance(timeout, str) and not timeout.isdigit():
            raise Exception("Timeout must be a number")

        timeout = int(timeout)
        conx = False
        try:
            outlook_365 = Outlook365(fromaddr, password, timeout)
            server = outlook_365.connect_smtp()
            
            if not_imap:
                print("IMAP connection avoided.")
            else:
                outlook_365.connect_imap()
                
            conx = True
        except:
            PrintException()

        if var_:
            SetVar(var_, conx)

    if module == "send_mail":
        to = GetParams('to')
        subject = GetParams('subject')
        body_ = GetParams('body')
        cc = GetParams('cc')
        bcc = GetParams('bcc')
        attached_file = GetParams('attached_file')
        files = GetParams('attached_folder')
        
        type_ = 'multipart'

        if cc is None:
            cc = ""
        if bcc is None:
            bcc = ""
        if attached_file is None:
            attached_file = ""
        if files is None:
            files = ""

        outlook_365.send_mail(
            to,
            subject,
            cc=cc,
            bcc=bcc,
            attachments_path=[attached_file, files],
            type_=type_,
            body=body_
        )

    if module == "get_mail":
        filtro = GetParams('filtro')
        var_ = GetParams('var_')
        folder = GetParams("folder")

        if not folder:
            folder = "inbox"

        if not filtro:
            filtro = "All"

        # lista = outlook_365.get_mail(filtro, folder)
        mail = outlook_365.imap
        filter_ = filtro
        mail.list()
        mail.select(folder)

        result, data = mail.search(None, filter_)
        ids = data[0]  # data is a list.
        id_list = ids.split()  # ids is a space separated string
        mail.logout()
        lista = [b.decode() for b in id_list]
        if var_:
            SetVar(var_, lista)

    if module == "get_unread":
        filtro = GetParams('filtro')
        var_ = GetParams('var_')
        folder = GetParams("folder")

        if not folder:
            folder = "inbox"

        if filtro is None:
            filtro = ""

        filtro = " ".join(["UNSEEN", filtro]).strip()
        lista = outlook_365.get_mail(filtro, folder)

        if var_:
            SetVar(var_, lista)

    if module == "read_mail":
        id_ = GetParams('id_')
        var_ = GetParams('var_')
        att_folder = GetParams('att_folder')
        folder = GetParams("folder")
        not_parsed = GetParams("not_parsed")

        if not folder:
            folder = "inbox"

        final = outlook_365.read_mail(id_, folder, att_folder)
        
        if var_:
            del final["mail"]
            if len(final['body']) > 1:
                if not not_parsed or eval(not_parsed) == False:
                    del final['body'][1]
                else:
                    del final['body'][0]

            SetVar(var_, final)

    if module == "reply_email":

        id_ = GetParams('id_')
        body_ = GetParams('body')
        folder = GetParams("folder")
        cc = GetParams('cc')
        bcc = GetParams('bcc')
        attached_file = GetParams('attached_file')
        files = GetParams('attached_folder')

        if cc is None:
            cc = ""
        if bcc is None:
            bcc = ""
        
        if not folder:
            folder = "inbox"
        if attached_file is None:
            attached_file = ""
        if files is None:
            files = ""
        
        attached_file_list = [attached_file, files]
        
        outlook_365.reply_mail(id_, folder, body_, attached_file_list, cc, bcc)

    if module == "create_folder":

        folder_name = GetParams('folder_name')
        mail = outlook_365.connect_imap()
        mail.create(folder_name)
        mail.logout()

    if module == "move_mail":
        id_ = GetParams("id_")
        label_ = GetParams("label_")
        from_ = GetParams("from")
        var = GetParams("var")

        if not id_:
            raise Exception("No ha ingresado ID de email a mover")
        if not label_:
            raise Exception("No ha ingresado carpeta de destino")
        if not from_ or from_ == "Inbox":
            from_ = "inbox"

        res = False
        res = outlook_365.move_mail(id_, from_, label_)
        if var:
            ret = True if res[0] == 'OK' else False
            SetVar(var, ret)

    if module == "forward":
        id_ = GetParams('id_')
        to_ = GetParams('email')
        cc = GetParams('cc')
        bcc = GetParams('bcc')

        if cc is None:
            cc = ""
        if bcc is None:
            bcc = ""

        outlook_365.forward_email(id_, "inbox", to_, cc, bcc)

    if module == "list_folders":
        result = GetParams('var')

        mail = outlook_365.connect_imap()
        folders = [folder.decode().split('"/"')[1]
                   for folder in mail.list()[1]]
        if result:
            SetVar(result, folders)

    if module == 'markAsUnread':
        id_ = GetParams("id_")
        folder = GetParams("folder")
        var = GetParams("var")

        if not folder:
            folder = "inbox"

        outlook_365.mark_as_unread(id_, folder)
    
    
    if module == "get_attachments":
        id_ = GetParams('id_')
        var_ = GetParams('var_')
        att_folder = GetParams('att_folder')
        folder = GetParams("folder")

        if not folder:
            folder = "inbox"

        final = outlook_365.get_attachments(id_, folder, att_folder)

    if module == "close":
        outlook_365 = None

except Exception as e:
    import traceback
    traceback.print_exc()
    PrintException()
    raise e
