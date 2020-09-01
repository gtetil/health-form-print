# install:  python-docx, pywin32

from docx import Document
import win32api
import win32print
from shutil import copyfile
from datetime import datetime

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
print_doc(doc_name_temp)


