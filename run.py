from docxtpl import DocxTemplate
from datetime import datetime
import subprocess
from uuid import uuid4

# from comtypes import client

# def convert_to_pdf(doc):
#     try:
#         word = client.DispatchEx("Word.Application")
#         new_name = doc.replace(".docx", r".pdf")
#         worddoc = word.Documents.Open(doc)
#         worddoc.SaveAs(new_name, FileFormat = 17)
#         worddoc.Close()
#     except Exception as e:
#         return e
#     finally:
#         word.Quit()

DATA_SET = {
    'user': {
        'name': 'Joseavi',
        'age': 30,
    },
    'institution': 'Edoo',
    'date': datetime.utcnow(),
    'country': 'Estados Unidos',
    'children': [
        {
            'name': 'juanito',
            'level': 'Primero Primaria',
        },
        {
            'name': 'marito',
            'level': 'Segundo Primaria',
        },
    ]
}

doc = DocxTemplate('./docs/contract.docx')
doc.render(DATA_SET)

destiny = './docs/output.docx'.format(uuid4())
doc.save(destiny)

p = subprocess.Popen(['libreoffice', '--convert-to', 'pdf', destiny, '--outdir', './docs'], stdin=subprocess.PIPE, stdout=subprocess.PIPE)
outs, errs = p.communicate()

# with open('./docs/output-{}.pdf'.format(uuid4()), 'wb') as f:
#     f.write(outs)
#
# print(type(outs), 'error ->', errs)


# word = client.CreateObject('Word.Application')
# word.Do
