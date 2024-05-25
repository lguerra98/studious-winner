import pandas as pd
import numpy as np
from docx import Document
from datetime import datetime as dt
import random
import os



df = pd.read_excel(os.path.abspath("invoice.xlsx"), sheet_name="Sheet1").ffill().set_index("Cliente")
clients = pd.read_excel(os.path.abspath("inventario.xlsm"), sheet_name="clientes")
inventario = pd.read_excel(os.path.abspath("inventario.xlsm"), sheet_name="inventario")

# os.remove("invoice.xlsx")

client = df.index.values[0]

address = clients.loc[clients["tienda"] == client, "direccion"].iloc[0]

email = clients.loc[clients["tienda"] == client, "email"].iloc[0]

NIT = "123456789-0"
invoice = Document()

invoice.add_heading("Factura", 0)

p_1 = invoice.add_paragraph()
p_1.add_run("NIT: ").bold = True
p_1.add_run(NIT)
p_2 = invoice.add_paragraph()
p_2.add_run("Direccion: ").bold = True
p_2.add_run(address)
p_3 = invoice.add_paragraph()
p_3.add_run("email: ").bold = True
p_3.add_run(email)
invoice.add_paragraph("")
p_4 = invoice.add_paragraph()
p_4.add_run("Fecha: ").bold = True
p_4.add_run(dt.today().strftime("%d/%m/%Y"))
p_5 = invoice.add_paragraph()
p_5.add_run("No de Factura: ").bold = True
p_5.add_run(f"{random.randint(0, 1000000)}")
invoice.add_paragraph("")
invoice.add_heading("Detalle de la Compra", level=1)
invoice.add_paragraph("")

table = invoice.add_table(rows=1, cols=4)

for i,j in zip(table.rows[0].cells, df.columns):
    i.text = j
    
for i in df.values:
    new_row = table.add_row().cells
    for j,k in zip(new_row, i):
        j.text = f"{k}"

last_row = table.add_row().cells

last_row[0].text = "Total"
last_row[-1].text = f"{df['Valor'].sum()}"

invoice.save("invoice.docx")







