import docx
import subprocess
lpr = subprocess.Popen("/usr/bin/lpr", stdin=subprocess.PIPE)

doc = docx.Document()
doc.add_heading('Вывод документа для печати', 0)
doc.add_heading('Таблица', level=1)
print("Введите три значения для таблицы ")
data = (
    (1, input('Ячейка \n')),
    (2, input()),
    (3, input())
)

table = doc.add_table(rows=1, cols=2)
row = table.rows[0].cells
row[0].text = 'Id'
row[1].text = 'Name'

for id, name in data:
    row = table.add_row().cells
    row[0].text = str(id)
    row[1].text = name

doc.add_paragraph()
doc.add_paragraph(input("Введите текст\n"))
doc.save('file.docx')
lpr.stdin.write(doc)

