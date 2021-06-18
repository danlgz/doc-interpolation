from docxtpl import DocxTemplate
from datetime import datetime
import subprocess
from time import time


starts = time()

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

destiny = './output/output.docx'
doc.save(destiny)

p = subprocess.Popen(['libreoffice', '--convert-to', 'pdf', destiny, '--outdir', './output'], stdin=subprocess.PIPE, stdout=subprocess.PIPE)
outs, errs = p.communicate()

delta = round((time() - starts) * 1000, 2)
print('Time: {}ms'.format(delta))
