# lambda_report.py
from pathlib import Path
import xlwings as xw

# TODO: adjust the DIRECTORY path
DIRECTORY = r'C:\Users\felix\Desktop\lambda'

app = xw.App(visible=False)
for path in Path(DIRECTORY).rglob('[!~$]*.xls*'):
    book = app.books.open(path)
    print(f'------ {path} ------')
    for name in book.names:
        refers_to = name.refers_to.replace('_xlfn.', '').replace('_xlpm.', '')
        lambda_functions = []
        if refers_to.lower().startswith('=lambda'):
            lambda_functions.append(f'{name.name}: {refers_to}')
        if lambda_functions:
            for func in lambda_functions:
                print(func)
    print()
    book.close()
app.quit()
