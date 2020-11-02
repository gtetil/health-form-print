# install:  python-docx, pywin32

from docx import Document
import win32api
import win32print
from shutil import copyfile
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

doc_name = 'Employee Health Check Form.docx'
doc_name_temp = 'Employee Health Check Form temp.docx'
copyfile(doc_name, doc_name_temp)
f = open(doc_name_temp, 'rb')
document = Document(f)

def print_doc(path):
    win32api.ShellExecute (
      0,
      "print",
      path,
      #
      # If this is None, the default printer will
      # be used anyway.
      #
      '/d:"%s"' % win32print.GetDefaultPrinter (),
      ".",
      0
    )

for table in document.tables:
    for row in table.rows:
        for cell in row.cells:
            if "date_tag" in cell.text:
                date = datetime.today().strftime('%#m/%#d/%Y')
                cell.text = date

document.save(doc_name_temp)

f.close()
#print_doc(doc_name_temp)

gmail_user = "garrett.tetil@gmail.com"
gmail_pwd = "Waterside0!"
FROM = "garrett.tetil@gmail.com"
TO = 'jkozan@tifs.com'
SUBJECT = 'Health Form'

msg = MIMEMultipart()
msg['From'] = FROM
msg['To'] = TO
msg['Subject'] = SUBJECT

with open(doc_name_temp, "rb") as fil:
    part = MIMEApplication(
        fil.read(),
        Name=doc_name_temp
    )
# After the file is closed
part['Content-Disposition'] = 'attachment; filename="%s"' % doc_name
msg.attach(part)

# SMTP_SSL Example
server_ssl = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server_ssl.ehlo() # optional, called by login()
server_ssl.login(gmail_user, gmail_pwd)
# ssl server doesn't support or need tls, so don't call server_ssl.starttls()
server_ssl.sendmail(FROM, TO, msg.as_string())
#server_ssl.quit()
server_ssl.close()
print('successfully sent the mail')