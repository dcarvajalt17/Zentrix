import pandas as pd
from dash import Dash, html, dcc, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import plotly.express as px
import base64
import os

from Calculo_de_niveles_de_consumo import ejecutar_parametros
from Buffer_Mejorado import exportar_resumen

# === Inicializar App con estilo corporativo ===
app = Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.LUX,
        "https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap"
    ]
)
server = app.server

# === Codificar logo ===
logo_path = "assets/logo_zentrix.png"
encoded_image = ""
logo_img = html.Div("Logo no disponible", style={'color': 'red'})

if os.path.exists(logo_path):
    with open(logo_path, 'rb') as image_file:
        encoded_image = base64.b64encode(image_file.read()).decode()
    logo_img = html.Img(src='data:image/png;base64,' + encoded_image, className='dash-logo')

# === Variables Globales ===
df_buffer = pd.DataFrame()  
df_nobuffer = pd.DataFrame()

# === Layout ===
app.layout = dbc.Container([
    html.Div([
    html.Div([
        html.Img(src='/assets/logo_zentrix.png', className='dash-logo'),
        html.Div([
            html.H1("Zentrix Material Planning", className='main-title'),
            html.P("by Daniel Carvajal/Jesus Cabeza, 2025", className='subtitle')
        ])
    ], className='logo-title-group')
], className='header-container'),

    dbc.Card([
        dbc.CardHeader(html.H4("Ejecución del Plan DDMRP", className="text-primary")),
        dbc.CardBody([
            dbc.Row([
                dbc.Col(
                    dbc.Input(
                        id='input-fecha',
                        type='text',
                        placeholder='DD/MM/YYYY',
                        debounce=True,
                        style={
                            'color': 'white',
                            'background-color': '#292951',
                            'border-color': '#4f93ce'
                        },
                        className="custom-placeholder"
                    ),
                    width=3
                ),
                dbc.Col(
                    dbc.Button("Ejecutar flujo completo", id="btn-ejecutar", color="dark"),
                    width="auto"
                )
            ], className="mb-3"),

            dcc.Loading(
                dbc.Textarea(
                    id='log-output',
                    value='',
                    rows=6,
                    className="log-box",
                    style={'color': 'black'}
                ),
                type='default'
            )
        ])
    ], className="mb-4 shadow-sm"),

    dbc.Card([
        dbc.CardHeader(html.H4("Plan de Abastecimiento")),
        dbc.CardBody([
            dcc.Tabs(id='tabs-selector', value='Buffer', children=[
                dcc.Tab(label='Buffer', value='Buffer'),
                dcc.Tab(label='No Buffer', value='No Buffer')
            ]),
            dbc.Row([
                dbc.Col(dcc.Dropdown(id='filtro-familia', placeholder="Familia", clearable=True, style={'color': 'white'}), width=3),
                dbc.Col(dcc.Dropdown(id='filtro-subfamilia', placeholder="Subfamilia", clearable=True, style={'color': 'white'}), width=3),
                dbc.Col(dcc.Dropdown(id='filtro-referencia', placeholder="Descripción", clearable=True, style={'color': 'white'}), width=3)
            ], className="mt-3 mb-3"),
            dcc.Graph(id='grafico-inventario')
        ])
    ], className="mb-4 shadow-sm"),

    dbc.Card([
        dbc.CardHeader(html.H4("Costos de Aprovisionamiento")),
        dbc.CardBody([
            dbc.Row([
                dbc.Col(dcc.Dropdown(id='filtro-costos-familia', placeholder="Familia", clearable=True, style={'color': 'white'}), width=3),
                dbc.Col(dcc.Dropdown(id='filtro-costos-subfamilia', placeholder="Subfamilia", clearable=True, style={'color': 'white'}), width=3),
                dbc.Col(dcc.Dropdown(id='filtro-costos-referencia', placeholder="Descripción", clearable=True, style={'color': 'white'}), width=3)
            ], className="mb-3"),
            dcc.Tabs(id='tabs-costos', value='Buffer', children=[
                dcc.Tab(label='Buffer', value='Buffer'),
                dcc.Tab(label='No Buffer', value='No Buffer'),
                dcc.Tab(label='Totalizado', value='Totalizado')
            ]),
            dcc.Graph(id='grafico-costos')
        ])
    ])
], fluid=True)

# === Callback ejecución flujo ===
@app.callback(
    Output('log-output', 'value'),
    Input('btn-ejecutar', 'n_clicks'),
    State('input-fecha', 'value'),
    prevent_initial_call=True
)
def ejecutar_flujo(n, fecha_usuario):
    global df_buffer, df_nobuffer
    logs = []
    def log(msg):
        print(msg)
        logs.append(msg)
    try:
        log("\U0001F449 Ejecutando cálculo de parámetros...")
        ejecutar_parametros('data-consumo1.xlsx', 'Referencia V2.xlsx', fecha_usuario)
        log("✅ Paso 1 completado.")

        log("📦 Ejecutando planificación y exportación...")
        exportar_resumen('Referencia V2.xlsx', 'data-consumo1.xlsx', 'Resumen_Buffer_NoBuffer_Semanal.xlsx')
        log("✅ Paso 2 completado.")

        log("\U0001F4CA Cargando datos...")
        df_buffer = pd.read_excel("Resumen_Buffer_NoBuffer_Semanal.xlsx", sheet_name="Buffer")
        df_nobuffer = pd.read_excel("Resumen_Buffer_NoBuffer_Semanal.xlsx", sheet_name="No Buffer")

        for df in [df_buffer, df_nobuffer]:
            df['Fecha_Evaluacion'] = pd.to_datetime(df['Fecha_Evaluacion'], errors='coerce')
            df['Mes'] = df['Fecha_Evaluacion'].dt.to_period('M').dt.to_timestamp()
            df['Familia'] = df['Familia'].fillna('SinDato')
            df['Subfamilia'] = df['Subfamilia'].fillna('SinDato')
            df['Descripcion_F'] = df['Descripcion_F'].fillna('SinDato')

        log("✅ Datos cargados correctamente.")
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print("❌ ERROR:", error_msg)
        log(f"❌ Error durante la ejecución:\n{error_msg}")
    return "\n".join(logs)

# === Callbacks de filtros ===
@app.callback(
    Output('filtro-familia', 'options'),
    Input('tabs-selector', 'value')
)
def cargar_familias(tab):
    df = df_buffer if tab == 'Buffer' else df_nobuffer
    return [{'label': i, 'value': i} for i in sorted(df['Familia'].dropna().unique())]

@app.callback(
    Output('filtro-subfamilia', 'options'),
    Input('filtro-familia', 'value'),
    State('tabs-selector', 'value')
)
def actualizar_subfamilia(familia, tab):
    df = df_buffer if tab == 'Buffer' else df_nobuffer
    if familia:
        df = df[df['Familia'] == familia]
    return [{'label': i, 'value': i} for i in sorted(df['Subfamilia'].dropna().unique())]

@app.callback(
    Output('filtro-referencia', 'options'),
    Input('filtro-subfamilia', 'value'),
    State('filtro-familia', 'value'),
    State('tabs-selector', 'value')
)
def actualizar_referencia(subfamilia, familia, tab):
    df = df_buffer if tab == 'Buffer' else df_nobuffer
    if familia:
        df = df[df['Familia'] == familia]
    if subfamilia:
        df = df[df['Subfamilia'] == subfamilia]
    return [{'label': i, 'value': i} for i in sorted(df['Descripcion_F'].dropna().unique())]

@app.callback(
    Output('grafico-inventario', 'figure'),
    Input('filtro-referencia', 'value'),
    State('tabs-selector', 'value'),
    State('filtro-familia', 'value'),
    State('filtro-subfamilia', 'value')
)
def graficar_inventario(desc, tab, fam, subfam):
    if not desc:
        return go.Figure().update_layout(title="Seleccione una referencia para ver el gráfico")

    df = df_buffer if tab == 'Buffer' else df_nobuffer
    if fam:
        df = df[df['Familia'] == fam]
    if subfam:
        df = df[df['Subfamilia'] == subfam]
    if desc:
        df = df[df['Descripcion_F'] == desc]

    if df.empty:
        return go.Figure().update_layout(title="No hay datos para graficar")

    df = df.sort_values('Fecha_Evaluacion')
    fig = go.Figure()

    if tab == 'Buffer':
        fig.add_trace(go.Scatter(x=df['Fecha_Evaluacion'], y=df['Red_Total'], name='Zona Roja',
                                 fill='tozeroy', line=dict(color='red'), opacity=0.2))
        fig.add_trace(go.Scatter(x=df['Fecha_Evaluacion'], y=df['TO_Yellow'], name='Zona Amarilla',
                                 fill='tonexty', line=dict(color='yellow'), opacity=0.2))
        fig.add_trace(go.Scatter(x=df['Fecha_Evaluacion'], y=df['TO_Green'], name='Zona Verde',
                                 fill='tonexty', line=dict(color='green'), opacity=0.2))
    else:
        fig.add_trace(go.Scatter(x=df['Fecha_Evaluacion'], y=df['Inventario_Minimo'], name='Min',
                                 fill='tozeroy', line=dict(color='red'), opacity=0.2))
        fig.add_trace(go.Scatter(x=df['Fecha_Evaluacion'], y=df['Inventario_Objetivo'], name='Objetivo',
                                 fill='tonexty', line=dict(color='yellow'), opacity=0.2))
        fig.add_trace(go.Scatter(x=df['Fecha_Evaluacion'], y=df['Inventario_Maximo'], name='Max',
                                 fill='tonexty', line=dict(color='green'), opacity=0.2))

    # Línea: posición inventario
    fig.add_trace(go.Scatter(
        x=df['Fecha_Evaluacion'],
        y=df['Posicion'],
        mode='lines',
        name='Posición Inventario',
        line=dict(color='blue', dash='dot'),
        hovertemplate="Semana: %{x|%d-%m-%Y}<br>Posición: %{y}<extra></extra>"
    ))

    # Línea: Stock Inicial (Stock OH)
    fig.add_trace(go.Scatter(
        x=df['Fecha_Evaluacion'],
        y=df['Stock_Inicial'],
        mode='lines+markers',
        name='Stock OH',
        line=dict(color='black'),
        hovertemplate="Semana: %{x|%d-%m-%Y}<br>Stock OH: %{y}<extra></extra>"
    ))

    # Línea: Consumo con Cant. Pedir en tooltip
    fig.add_trace(go.Bar(
        x=df['Fecha_Evaluacion'],
        y=df['Consumo_Proy'],
        name='Consumo',
        marker_color='grey',
        customdata=df[['Cantidad_Pedir']],
        hovertemplate="Semana: %{x|%d-%m-%Y}<br>Consumo: %{y}<br>Cant. Pedir: %{customdata[0]}<extra></extra>"
))
 

    # Línea punteada: Cantidad a Pedir (opcional)
    fig.add_trace(go.Scatter(
        x=df['Fecha_Evaluacion'],
        y=df['Cantidad_Pedir'],
        mode='lines+markers',
        name='Cantidad a Pedir',
        line=dict(color='purple', dash='dot'),
        visible='legendonly',
        hovertemplate="Semana: %{x|%d-%m-%Y}<br>Cant. Pedir: %{y}<extra></extra>"
    ))


    fig.update_layout(title=f"Comportamiento de Inventario - {desc}", template="plotly_white")
    return fig

@app.callback(
    Output('filtro-costos-familia', 'options'),
    Input('tabs-costos', 'value')
)
def cargar_costos_familias(tab):
    df = df_buffer if tab == 'Buffer' else df_nobuffer if tab == 'No Buffer' else pd.concat([df_buffer, df_nobuffer])
    return [{'label': i, 'value': i} for i in sorted(df['Familia'].dropna().unique())]

@app.callback(
    Output('filtro-costos-subfamilia', 'options'),
    Input('filtro-costos-familia', 'value'),
    State('tabs-costos', 'value')
)
def actualizar_costos_subfamilia(familia, tab):
    df = df_buffer if tab == 'Buffer' else df_nobuffer if tab == 'No Buffer' else pd.concat([df_buffer, df_nobuffer])
    if familia:
        df = df[df['Familia'] == familia]
    return [{'label': i, 'value': i} for i in sorted(df['Subfamilia'].dropna().unique())]

@app.callback(
    Output('filtro-costos-referencia', 'options'),
    Input('filtro-costos-subfamilia', 'value'),
    State('filtro-costos-familia', 'value'),
    State('tabs-costos', 'value')
)
def actualizar_costos_referencia(subfamilia, familia, tab):
    df = df_buffer if tab == 'Buffer' else df_nobuffer if tab == 'No Buffer' else pd.concat([df_buffer, df_nobuffer])
    if familia:
        df = df[df['Familia'] == familia]
    if subfamilia:
        df = df[df['Subfamilia'] == subfamilia]
    return [{'label': i, 'value': i} for i in sorted(df['Descripcion_F'].dropna().unique())]

@app.callback(
    Output('grafico-costos', 'figure'),
    Input('tabs-costos', 'value'),
    Input('filtro-costos-familia', 'value'),
    Input('filtro-costos-subfamilia', 'value'),
    Input('filtro-costos-referencia', 'value')
)
def graficar_costos(tab, fam, subfam, desc):
    df = df_buffer if tab == 'Buffer' else df_nobuffer if tab == 'No Buffer' else pd.concat([df_buffer, df_nobuffer])
    if fam:
        df = df[df['Familia'] == fam]
    if subfam:
        df = df[df['Subfamilia'] == subfam]
    if desc:
        df = df[df['Descripcion_F'] == desc]
    df = df.groupby('Mes')['vr_compra'].sum().reset_index()
    fig = px.bar(df, x='Mes', y='vr_compra', text_auto='.2s', labels={'vr_compra': 'Costo de Compra'})
    fig.update_traces(marker_color='indigo', hovertemplate='Costo: %{y:$,.0f}')
    fig.update_layout(title="Costo de Aprovisionamiento por Mes", yaxis_tickprefix="$", yaxis_tickformat=".2s")
    return fig

if __name__ == '__main__':
    print("🌐 Iniciando app en puerto 8060...")
    app.run(debug=True, port=8060)
