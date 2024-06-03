from plotly.subplots import make_subplots
import plotly.graph_objs as go
import plotly.express as px
import pandas as pd
import webbrowser
import os
import urllib.parse
from datetime import datetime

ventas = pd.DataFrame(columns=["DESCRIPCIÓN", "CANT", "PRECIO UNIDAD", "TOTAL", "METODO DE PAGO"])

mes = datetime.now().strftime("%B")
year = datetime.now().strftime("%Y")


for yr in os.listdir("facturas"):
    if os.path.isdir(f"facturas/{yr}"):
        for mt in os.listdir(f"facturas/{yr}"):
            if mt == mes:
                days = os.listdir(f"facturas/{yr}/{mt}")
                print(days)
                for day in days:
                    if os.path.isdir(f"facturas/{yr}/{mt}/{day}"):
                        facturas = os.listdir(f"facturas/{yr}/{mt}/{day}")
                        for factura in facturas:
                            
                            if factura.endswith(".xlsx"):
                                path = f"facturas/{yr}/{mt}/{day}/{factura}"
                                data = pd.read_excel(path)
                                if data["DIRECCION CLIENTE"].iloc[0] == "CANCELADO":
                                    continue
                                else:
                                    ventas = pd.concat([ventas, data[["DESCRIPCIÓN", "CANT", "PRECIO UNIDAD", "TOTAL", "METODO DE PAGO"]]])
                    else:
                        continue
                      
                      
ventas.ffill(inplace=True)

datos_ventas = ventas.groupby("DESCRIPCIÓN", as_index=False)[["CANT", "PRECIO UNIDAD", "TOTAL"]].sum()
datos_pagos = ventas.groupby("METODO DE PAGO", as_index=False)["TOTAL"].sum()

fig = make_subplots(rows=1, cols=2, subplot_titles=("<b>Productos Vendidos</b>", "<b>Analisis Metodos de pago</b>"), specs=[[{'type': 'xy'}, {'type': 'domain'}]])

pie = go.Pie(labels=datos_pagos["METODO DE PAGO"], values=datos_pagos["TOTAL"], name="Bar chart", hovertext=datos_pagos["METODO DE PAGO"], 
            hovertemplate='<b>%{hovertext}</b><br>Total vendido:$%{value:,.2f}<extra></extra>', hole=0.55, marker={'line': {'color': 'white', 'width': 1.9}}, 
            legendgrouptitle={'text': '<b>Metodo de pago</b><br>'})

fig.add_trace(pie, row=1, col=2)

fig.add_annotation(text=f'<b>Total vendido<br>{int(datos_pagos["TOTAL"].sum()):,.0f}</b>', showarrow=False, xref="paper", yref="paper", x=0.837, y=0.5, 
                  font={'size':20})
fig.update_traces(marker={"colors":px.colors.qualitative.Vivid}, row=1, col=2)

datos_ventas = datos_ventas.sort_values(by="CANT")

bar = go.Bar(x=datos_ventas["DESCRIPCIÓN"], y=datos_ventas["CANT"], 
             hovertemplate="<b>Producto:</b> %{x}<><br><b>Cantidad vendidad:</b> %{y}<br><b>Total ventas:</b> %{text}<extra></extra>", 
             text="$" + datos_ventas["TOTAL"].round(1).astype(str), marker={'color': 'rgba(53, 174, 230, 0.63)', 'pattern': {'shape': ''}}, 
             textposition="outside", showlegend=False)


fig.add_trace(bar, row=1, col=1)

fig.update_layout(xaxis={'title':"<b>Productos Vendidos</b>", 'anchor': 'y', 'domain': [0.0, 0.45], 'title': {'text': '<b>Producto</b>'}}, 
                  yaxis={'anchor': 'x', 'domain': [0.0, 1.0], 'title': {'text': '<b>Cantidad vendidad</b>'}}, 
                  plot_bgcolor="white")

path = f"facturas/{year}/{mes}"



fig.write_html(f"{path}/analis_ventas.html")

full_html = os.path.abspath(f"{path}/analis_ventas.html")

file_url = urllib.parse.urljoin('file:', urllib.request.pathname2url(full_html))


chrome_path = "C:/Program Files/Google/Chrome/Application/chrome_proxy.exe"

webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))



webbrowser.get('chrome').open(file_url)



