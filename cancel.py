import os
import pandas as pd

n_fac = input("Ingresar numero de factura a cancelar: ")

for yr in os.listdir("facturas"):
    if os.path.isdir(f"facturas/{yr}"):
        for mt in os.listdir(f"facturas/{yr}"):
            for day in os.listdir(f"facturas/{yr}/{mt}"):
                if os.path.isdir(f"facturas/{yr}/{mt}/{day}"):
                    for z in os.listdir(f"facturas/{yr}/{mt}/{day}"):
                        if z.endswith(f"{n_fac}.xlsx"):
                            path_cancel = f"facturas/{yr}/{mt}/{day}/{z}"
                            to_cancel = pd.read_excel(path_cancel)

to_cancel["DIRECCION CLIENTE"] = "CANCELADO"

to_cancel.to_excel(r"C:\Users\Lukas\Desktop\studious-winner\copy.xlsx", index=False)

with pd.ExcelWriter(path_cancel) as writer:
    to_cancel.to_excel(writer, sheet_name="factura", index=False)

os.remove(path_cancel.split(".")[0] + ".docx")