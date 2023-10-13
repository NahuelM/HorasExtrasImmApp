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
    html.Div(id="imgCointainer", className="imgContainer"),
    html.Div([
        dbc.Button('Generar reportes', color="success", id='procesarPDF-button', disabled=True),
        dcc.Loading(
                id="loading-2",
                children=[html.Div([html.Div(id="loading-output-2")])],
                type="circle",
            ),
        dbc.Button('Descargar Excel', color="success", id='download-excel-button', disabled=True),
    ],className="d-grid gap-2 col-6 mx-auto"),
    
    dcc.Download(id="download-csv"), 
    
    html.Div(id='output-excel')
])
outFiles = []

files = []
k = 0


@app.callback(
    Output('imgCointainer', 'children'),
    Output('procesarPDF-button', 'disabled'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename'),
)
def uploadFiles(contents, filename):
    if not contents:
        raise PreventUpdate
    divs = []
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
        i +=1
    for imgName in filename:
        div = html.Div([
                html.P(imgName),
                html.Img(src=app.get_asset_url('Icon_pdf_file.pdf')),  
        ])

        divs.append(div)
    k = i
    return divs, False

@app.callback(
    Output('output-excel', 'children'),
    Output('download-excel-button', 'style'),
    Output('download-excel-button', 'disabled'),
    Output("loading-output-2", "children"),
    Input('procesarPDF-button', 'n_clicks')
)
def generate_excel(n_clicks):
    if not n_clicks:
        raise PreventUpdate
    out = pdScrap.crearReporte(files, "")

    outFiles.append(out)

    for j in range(k-1, -1, -1):
        os.remove("archivo"+str(j)+".pdf")

    return 'Archivo Excel generado con éxito.', {'display': 'block'}, False, None


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
