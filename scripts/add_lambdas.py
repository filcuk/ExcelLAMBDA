# add_lambdas.py
from pathlib import Path
import xlwings as xw

# TODO: adjust the DIRECTORY path and the LAMBDAS you want to add
DIRECTORY = r'C:\Users\felix\Desktop\lambda'

LAMBDAS = (
    ('Hypotenuse', '=LAMBDA(a, b, SQRT((a^2+b^2)))'),
    ('CountWords', '=LAMBDA(text, LEN(TRIM(text)) - LEN(SUBSTITUTE(TRIM(text), " ", "")) + 1)')
)

app = xw.App(visible=False)
for path in Path(DIRECTORY).rglob('[!~$]*.xls*'):
    print(f'Adding Lambda Functions to: {path}')
    book = app.books.open(path)
    for func in LAMBDAS:
        name, refers_to = func[0], func[1]
        if name in [n.name for n in book.names]:
            book.names[func[0]].delete()
        book.names.add(name=name, refers_to=refers_to)
    book.save()
    book.close()
app.quit()
