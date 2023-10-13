import dash
from dash import Dash, dcc, html, dash_table, Input, Output, State, callback
from dash.dependencies import Input, Output
import os
import base64
import pandas as pd
from dash.exceptions import PreventUpdate
from flask import send_from_directory
import dash_bootstrap_components as dbc
import pdf_Scraper as pdScrap
import binascii
external_stylesheets = [
    {
        'href': 'styles.css',
        'rel': 'stylesheet'
    }
]

app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, {
        'href': 'styles.css',
        'rel': 'stylesheet'
}, dbc.icons.BOOTSTRAP], title='HorasExtras')
server = app.server

app.layout = html.Div([
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            'Drag and Drop or ',
            html.A('Select Files')
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        # Allow multiple files to be uploaded
        multiple=True
    ),
    html.Div(id='output-data-upload'),
    html.Button('Generar reportes', id='procesarPDF-button'),
    dcc.Download(id="download-csv"), 
    html.Button('Descargar Excel', id='download-excel-button'),
    html.Div(id='output-excel')
])
outFiles = []

@app.callback(
    Output('output-excel', 'children'),
    Output('download-excel-button', 'style'),
    Input('upload-data', 'contents'),
    Input('procesarPDF-button', 'n_clicks')
)
def process_and_generate_excel(contents, n_clicks):
    if not contents:
        raise PreventUpdate

    if n_clicks is None:
        raise PreventUpdate

    files = []
    i = 0
    for pdf_content in contents:
        try:
            decoded_data = base64.b64decode(pdf_content.split(",")[1])
            with open("archivo"+str(i)+".pdf", "wb") as pdf_file:
                pdf_file.write(decoded_data)
            print("PDF exportado con éxito.")
        except binascii.Error as e:
            print("Error al decodificar la cadena base64:", str(e))
        except Exception as e:
            print("Error inesperado:", str(e))
        files.append("archivo"+str(i)+".pdf")
        i+=1
    out = pdScrap.crearReporte(files, "")

    for k in range (0, len(out)):
        outFiles.append(out[k])


    for j in range(i, 0):
        print("remove")
        os.remove("archivo"+str(j)+".pdf")


    return 'Archivo Excel generado con éxito.', {'display': 'block'}


@callback(
    Output("download-csv", "data"),
    Input("download-excel-button", "n_clicks"),
    prevent_initial_call=True,
)
def func(n_clicks):
    pd = outFiles[0]
    return dcc.send_data_frame(pd.to_excel, "mydf.xlsx", sheet_name="Sheet_name_1")





if __name__ == '__main__':
    app.run_server(debug=True)
