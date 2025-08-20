import pandas as pd
from dash import Dash, html, dcc, Input, Output
import plotly.graph_objects as go

# === Variables globales ===
df_buffer = pd.DataFrame()
df_nobuffer = pd.DataFrame()

def preparar_df(df):
    df.rename(columns={'to_green': 'TO_Green', 'to_yellow': 'TO_Yellow'}, inplace=True)
    df['Fecha_Pedido'] = pd.to_datetime(df['Fecha_Pedido'], errors='coerce')
    df['Referencia'] = df['Referencia'].astype(str).str.strip().str.upper()
    df['Familia'] = df['Familia'].fillna("SinDato")
    df['Subfamilia'] = df['Subfamilia'].fillna("SinDato")
    df['Descripcion_F'] = df['Descripcion_F'].fillna("SinDato")
    df['Inventario_Min'] = df.get('Inventario_Minimo', 0)
    df['Inventario_Max'] = df.get('Inventario_Maximo', 0)
    df['Inventario_Objetivo'] = df.get('Inventario_Objetivo', 0)
    return df

def crear_app(ruta_archivo="Resumen_Buffer_NoBuffer_Semanal.xlsx"):
    global df_buffer, df_nobuffer

    df_buffer = preparar_df(pd.read_excel(ruta_archivo, sheet_name="Buffer"))
    df_nobuffer = preparar_df(pd.read_excel(ruta_archivo, sheet_name="No Buffer"))

    app = Dash(__name__)
    server = app.server

    app.layout = html.Div([
        html.H2("Visualizador Semanal de Inventario DDMRP"),

        dcc.Tabs(id='tabs-selector', value='Buffer', children=[
            dcc.Tab(label='Buffer', value='Buffer'),
            dcc.Tab(label='No Buffer', value='No Buffer'),
        ]),

        html.Div([
            html.Label("Filtrar por Subfamilia:"),
            dcc.Dropdown(id='subfamilia-dropdown', placeholder="Selecciona una Subfamilia"),
            html.Label("Filtrar por Descripción:"),
            dcc.Dropdown(id='descripcion-dropdown', placeholder="Selecciona una Descripción"),
            html.Label("Selecciona una Referencia:"),
            dcc.Dropdown(id='ref-dropdown', placeholder="Selecciona una Referencia")
        ], id='contenedor-filtros'),

        dcc.Graph(id='grafico-inventario'),
    ])

    @app.callback(
        Output('subfamilia-dropdown', 'options'),
        Input('tabs-selector', 'value')
    )
    def actualizar_subfamilias(tab):
        df = df_buffer if tab == 'Buffer' else df_nobuffer
        return [{'label': sf, 'value': sf} for sf in sorted(df['Subfamilia'].dropna().unique())]

    @app.callback(
        Output('descripcion-dropdown', 'options'),
        [Input('tabs-selector', 'value'), Input('subfamilia-dropdown', 'value')]
    )
    def actualizar_descripciones(tab, subfamilia):
        df = df_buffer if tab == 'Buffer' else df_nobuffer
        if not subfamilia:
            return []
        return [{'label': d, 'value': d} for d in sorted(df[df['Subfamilia'] == subfamilia]['Descripcion_F'].dropna().unique())]

    @app.callback(
        Output('ref-dropdown', 'options'),
        [Input('tabs-selector', 'value'), Input('subfamilia-dropdown', 'value'), Input('descripcion-dropdown', 'value')]
    )
    def actualizar_referencias(tab, subfamilia, descripcion):
        df = df_buffer if tab == 'Buffer' else df_nobuffer
        dff = df.copy()
        if subfamilia:
            dff = dff[dff['Subfamilia'] == subfamilia]
        if descripcion:
            dff = dff[dff['Descripcion_F'] == descripcion]
        return [{'label': r, 'value': r} for r in sorted(dff['Referencia'].dropna().unique())]

    @app.callback(
        Output('grafico-inventario', 'figure'),
        [Input('tabs-selector', 'value'), Input('ref-dropdown', 'value')]
    )
    def actualizar_grafico(tab, referencia):
        df = df_buffer if tab == 'Buffer' else df_nobuffer
        dff = df[df['Referencia'] == referencia].copy()

        if dff.empty:
            return go.Figure().update_layout(title="No hay datos", template="plotly_white")

        fig = go.Figure()
        dff = dff.fillna(0)

        if tab == 'Buffer':
            fig.add_trace(go.Scatter(x=list(dff['Fecha_Pedido']) + list(dff['Fecha_Pedido'])[::-1],
                                     y=[0]*len(dff) + list(dff['Red_Total'])[::-1], fill='toself',
                                     fillcolor='rgba(255,0,0,0.2)', line=dict(width=0), name='Zona Roja'))
            fig.add_trace(go.Scatter(x=list(dff['Fecha_Pedido']) + list(dff['Fecha_Pedido'])[::-1],
                                     y=list(dff['Red_Total']) + list(dff['TO_Yellow'])[::-1], fill='toself',
                                     fillcolor='rgba(255,255,0,0.2)', line=dict(width=0), name='Zona Amarilla'))
            fig.add_trace(go.Scatter(x=list(dff['Fecha_Pedido']) + list(dff['Fecha_Pedido'])[::-1],
                                     y=list(dff['TO_Yellow']) + list(dff['TO_Green'])[::-1], fill='toself',
                                     fillcolor='rgba(0,255,0,0.2)', line=dict(width=0), name='Zona Verde'))
        else:
            objetivo_ajustado = dff['Inventario_Objetivo'] * 1.2
            sobrestock_limite = dff[['Inventario_Proy', 'Inventario_Max']].max(axis=1)
            fig.add_trace(go.Scatter(x=list(dff['Fecha_Pedido']) + list(dff['Fecha_Pedido'])[::-1],
                                     y=[0]*len(dff) + list(dff['Inventario_Min'])[::-1], fill='toself',
                                     fillcolor='rgba(255,0,0,0.2)', line=dict(width=0), name='Zona Crítica'))
            fig.add_trace(go.Scatter(x=list(dff['Fecha_Pedido']) + list(dff['Fecha_Pedido'])[::-1],
                                     y=list(dff['Inventario_Min']) + list(objetivo_ajustado)[::-1], fill='toself',
                                     fillcolor='rgba(255,255,0,0.2)', line=dict(width=0), name='Zona Objetivo'))
            fig.add_trace(go.Scatter(x=list(dff['Fecha_Pedido']) + list(dff['Fecha_Pedido'])[::-1],
                                     y=list(objetivo_ajustado) + list(dff['Inventario_Max'])[::-1], fill='toself',
                                     fillcolor='rgba(0,255,0,0.2)', line=dict(width=0), name='Zona Máxima'))
            fig.add_trace(go.Scatter(x=list(dff['Fecha_Pedido']) + list(dff['Fecha_Pedido'])[::-1],
                                     y=list(dff['Inventario_Max']) + list(sobrestock_limite)[::-1], fill='toself',
                                     fillcolor='rgba(0,0,255,0.2)', line=dict(width=0), name='Zona Sobrestock'))

        fig.add_trace(go.Scatter(x=dff['Fecha_Pedido'], y=dff['Inventario_Proy'], mode='markers+lines',
                                 name='Posición', marker=dict(size=6, color='blue'),
                                 line=dict(color='black', width=1.5)))

        fig.add_trace(go.Bar(x=dff['Fecha_Pedido'], y=dff['Consumo_Proy'], name='Consumo Proy',
                             marker_color='gray', opacity=0.6))

        fig.update_layout(
            title=f"Comportamiento de Inventario - {referencia} ({tab})",
            xaxis_title="Fecha",
            yaxis_title="Inventario y Consumo",
            hovermode="x unified",
            template="plotly_white",
            margin=dict(l=60, r=120, t=60, b=40),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0)
        )

        return fig

    return app

def lanzar_visualizador():
    app = crear_app()
    app.run_server(debug=True, port=8050, use_reloader=False)

if __name__ == "__main__":
    lanzar_visualizador()
