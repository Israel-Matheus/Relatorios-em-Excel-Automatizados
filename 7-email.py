import smtplib
import ssl
import mimetypes
from email.message import EmailMessage

# dados do e-mail
password = open("senha", "r").read()
from_email = "" # coloque seu e-mail aqui
to_email = "" # coloque aqui o e-mail pra quem vai enviar
subject = "Automação na Planilha"
body = """
Olá. Segue me anexo a automação da planilha
para a empresa XYZ Automação.

Qualquer dúvida estou a disposição!
"""

# montando a estrutura do e-mail
message = EmailMessage()
message["From"] = from_email
message["To"] = to_email
message["Subject"] = subject

message.set_content(body)
safe = ssl.create_default_context()

# adicionar anexo
attachment = "test.xlsx"
mime_type, mime_subtype = mimetypes.guess_type(attachment)[0].split("/")
with open(attachment, "rb") as a:
    message.add_attachment(
        a.read(),
        maintype=mime_type,
        subtype=mime_subtype,
        filename=attachment
    )

# envio do e-mail
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=safe) as smtp:
    smtp.login(from_email, password)
    smtp.sendmail(
        from_email,
        to_email,
        message.as_string()
    )