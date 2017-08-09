from docx import Document
document = Document()
rows=10
table = document.add_table(rows=rows, cols=3)
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Qty'
hdr_cells[1].text = 'Id'
hdr_cells[2].text = 'Desc'
for i in range(rows):
    row_cells = table.add_row().cells
    row_cells[0].text ='Qty'
    row_cells[1].text ='Qty'
    row_cells[2].text ='Qty'

