import pandas as pd
from dash import Dash, dcc, html, dash_table, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import base64

# ================================
#  FUNCIONES AUXILIARES
# ================================
def normalizar_columnas(df):
    df.columns = (
        df.columns.str.replace("\n", " ")
        .str.replace("  ", " ")
        .str.upper()
        .str.strip()
    )
    return df

def limpiar_riesgo(valor):
    if pd.isna(valor):
        return "NO ENCONTRADO"
    v = str(valor).upper().strip()
    if "BAJO" in v:
        return "BAJO"
    elif "MEDIO" in v:
        return "MEDIO"
    elif "ALTO" in v:
        return "ALTO"
    else:
        return "NO ENCONTRADO"

def comparar_riesgo(row):
    r1, r2 = row["RIESGO_2025_1"], row["RIESGO_2025_2"]
    if r1 in ["BAJO", "MEDIO", "ALTO"] and r2 == "NO ENCONTRADO":
        return "SOLO EN 2025-1"
    elif r1 == "NO ENCONTRADO" and r2 in ["BAJO", "MEDIO", "ALTO"]:
        return "SOLO EN 2025-2"
    elif r1 == r2:
        return "SE MANTUVO"
    elif (r1 == "ALTO" and r2 in ["MEDIO", "BAJO"]) or (r1 == "MEDIO" and r2 == "BAJO"):
        return "MEJORO"
    elif (r1 == "BAJO" and r2 in ["MEDIO", "ALTO"]) or (r1 == "MEDIO" and r2 == "ALTO"):
        return "EMPEORO"
    else:
        return "OTRO"

# ================================
#  LEER BASE BECARIOS
# ================================
try:
    file_id = "1OOiHkMC4XOXgFBwId1hjKMfBt7QuY2lj"  # ID de BECARIOS
    url = f"https://drive.google.com/uc?export=download&id={file_id}"

    df_becarios = pd.read_excel(url, sheet_name="BECARIOS")
    df_becarios = normalizar_columnas(df_becarios)

    # Normalizar columnas de riesgo
    col_riesgo1 = [c for c in df_becarios.columns if "2025-1" in c][0]
    col_riesgo2 = [c for c in df_becarios.columns if "2025-2" in c][0]

    df_becarios["RIESGO_2025_1"] = df_becarios[col_riesgo1].apply(limpiar_riesgo)
    df_becarios["RIESGO_2025_2"] = df_becarios[col_riesgo2].apply(limpiar_riesgo)

    # Calcular evolución
    df_becarios["EVOLUCION"] = df_becarios.apply(comparar_riesgo, axis=1)

except Exception as e:
    print(f"Error al cargar datos: {e}")
    # Crear datos de ejemplo en caso de error
    df_becarios = pd.DataFrame({
        'APELLIDOS Y NOMBRES': ['Estudiante 1', 'Estudiante 2', 'Estudiante 3'],
        'TIPO DE BENEFICIO': ['BECA', 'CREDITO', 'BECA'],
        'RIESGO_2025_1': ['ALTO', 'MEDIO', 'BAJO'],
        'RIESGO_2025_2': ['MEDIO', 'MEDIO', 'BAJO'],
        'EVOLUCION': ['MEJORO', 'SE MANTUVO', 'SE MANTUVO']
    })

# ================================
#  TABLAS RESUMEN
# ================================
riesgo_count_1 = df_becarios["RIESGO_2025_1"].value_counts().reset_index()
riesgo_count_1.columns = ["NIVEL", "TOTAL"]
riesgo_count_1["MOMENTO"] = "2025-1"

riesgo_count_2 = df_becarios["RIESGO_2025_2"].value_counts().reset_index()
riesgo_count_2.columns = ["NIVEL", "TOTAL"]
riesgo_count_2["MOMENTO"] = "2025-2"

riesgo_resumen = pd.concat([riesgo_count_1, riesgo_count_2], ignore_index=True)

# ================================
#  TABLA EN FORMATO ANCHO (NIVEL, 2025-1, 2025-2)
# ================================
orden_niveles = ["ALTO", "MEDIO", "BAJO"]

tabla_resumen = riesgo_resumen.pivot_table(
    index="NIVEL", columns="MOMENTO", values="TOTAL", fill_value=0
).reset_index()

# Quitar NO ENCONTRADO y ordenar
tabla_resumen = tabla_resumen[tabla_resumen["NIVEL"].isin(orden_niveles)]
tabla_resumen["NIVEL"] = pd.Categorical(tabla_resumen["NIVEL"], categories=orden_niveles, ordered=True)
tabla_resumen = tabla_resumen.sort_values("NIVEL")

# ================================
#  SEPARAR NO ENCONTRADOS
# ================================
riesgo_validos = riesgo_resumen[riesgo_resumen["NIVEL"].isin(orden_niveles)]
riesgo_no_encontrado = riesgo_resumen[riesgo_resumen["NIVEL"] == "NO ENCONTRADO"]

# ================================
#  INDICADORES (KPIs)
# ================================
total_becarios = len(df_becarios)
mejoraron = (df_becarios["EVOLUCION"] == "MEJORO").sum()
empeoraron = (df_becarios["EVOLUCION"] == "EMPEORO").sum()

se_mantuvieron = df_becarios[
    (df_becarios["EVOLUCION"] == "SE MANTUVO") &
    (df_becarios["RIESGO_2025_1"] != "NO ENCONTRADO") &
    (df_becarios["RIESGO_2025_2"] != "NO ENCONTRADO")
].shape[0]

# ================================
#  DASHBOARD
# ================================
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "Dashboard Becarios 2025"

# Paleta de colores profesional
COLORS = {
    'primary': '#2E86AB',      # Azul elegante
    'success': '#A23B72',      # Rosa profesional
    'warning': '#F18F01',      # Naranja vibrante
    'danger': '#C73E1D',       # Rojo elegante
    'info': '#5C4B51',         # Gris violáceo
    'secondary': '#81B29A',    # Verde menta
    'dark': '#3D5467',         # Azul grisáceo
    'light': '#F3E37C'         # Amarillo suave
}

def tarjeta_moderna(titulo, valor, color_hex, icono="", descripcion=""):
    return dbc.Card([
        dbc.CardBody([
            html.Div([
                html.Div([
                    html.I(className=f"fas fa-{icono}", style={
                        'fontSize': '2.5rem', 
                        'color': color_hex,
                        'marginBottom': '10px'
                    }),
                ], className="text-center"),
                html.H3(valor, className="card-title text-center", style={
                    'color': color_hex, 
                    'fontWeight': 'bold',
                    'fontSize': '2.2rem',
                    'marginBottom': '5px'
                }),
                html.H6(titulo, className="card-subtitle text-center", style={
                    'color': '#666',
                    'fontWeight': '500',
                    'fontSize': '0.9rem'
                }),
                html.P(descripcion, className="card-text text-center", style={
                    'color': '#888',
                    'fontSize': '0.8rem',
                    'marginTop': '8px',
                    'marginBottom': '0'
                })
            ])
        ], style={'padding': '20px'})
    ], style={
        'border': 'none',
        'borderRadius': '15px',
        'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
        'background': 'linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%)',
        'transition': 'transform 0.3s ease',
        'margin': '10px 5px',
        'height': '180px'
    }, className="h-100 hover-card")

def grafico_riesgo():
    fig = px.bar(
        riesgo_validos,
        x="NIVEL",
        y="TOTAL",
        color="MOMENTO",
        barmode="group",
        text="TOTAL",
        category_orders={"NIVEL": orden_niveles},
        color_discrete_map={
            "2025-1": COLORS['info'],
            "2025-2": COLORS['primary']
        }
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(
        title={
            'text': '<b>📊 Distribución de Becarios por Nivel de Riesgo</b>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': COLORS['primary']}
        },
        xaxis_title="<b>Nivel de Riesgo</b>",
        yaxis_title="<b>Número de Becarios</b>",
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Arial, sans-serif", size=12),
        height=450,
        margin=dict(t=90, b=50, l=50, r=50),
        legend=dict(
            title="<b>Período</b>",
            orientation="h",
            yanchor="bottom",
            y=1.05,
            xanchor="center",
            x=0.5
        )
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor='rgba(128,128,128,0.1)')
    return fig

def grafico_no_encontrado():
    if riesgo_no_encontrado.empty:
        # Crear un gráfico vacío si no hay datos
        fig = go.Figure()
        fig.add_annotation(
            text="No hay datos de 'NO ENCONTRADO'",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color="gray")
        )
        fig.update_layout(
            title="📉 Evolución de NO ENCONTRADOS (2025-1 vs 2025-2)",
            height=400
        )
        return fig
    
    fig = px.bar(
        riesgo_no_encontrado,
        x="TOTAL",
        y="MOMENTO",
        orientation="h",
        text="TOTAL",
        color="MOMENTO",
        color_discrete_map={
            "2025-1": COLORS['secondary'],
            "2025-2": COLORS['dark']
        }
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(
        title={
            'text': '<b>📉 Evolución de NO ENCONTRADOS (2025-1 vs 2025-2)</b>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': COLORS['primary']}
        },
        xaxis_title="<b>Cantidad</b>",
        yaxis_title="",
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Arial, sans-serif", size=12),
        height=400,
        margin=dict(t=60, b=50, l=50, r=50),
        yaxis=dict(
            showticklabels=True, 
            categoryorder="total descending",
            type='category'
        ),
        showlegend=False
    )
    fig.update_xaxes(showgrid=True, gridcolor='rgba(128,128,128,0.1)')
    fig.update_yaxes(showgrid=False)
    return fig

# ================================
#  GRÁFICO: EMPEORARON POR BENEFICIO (AJUSTADO)
# ================================
def grafico_empeoraron_por_beneficio():
    df_empeoraron = df_becarios[df_becarios["EVOLUCION"] == "EMPEORO"]

    if df_empeoraron.empty:
        fig = go.Figure()
        fig.add_annotation(
            text="No hay estudiantes que hayan empeorado",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color="gray")
        )
        fig.update_layout(
            title="📊 Empeoraron por Tipo de Beneficio",
            height=450
        )
        return fig

    empeoraron_por_beneficio = (
        df_empeoraron.groupby("TIPO DE BENEFICIO")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )

    fig = px.bar(
        empeoraron_por_beneficio,
        x="TIPO DE BENEFICIO",
        y="TOTAL",
        text="TOTAL",
        color="TIPO DE BENEFICIO",
        color_discrete_sequence=[COLORS['danger'], COLORS['warning'], COLORS['info'], COLORS['secondary'], COLORS['dark']]
    )
    
    fig.update_traces(
        textposition="auto",
        textfont=dict(size=14, weight='bold')
    )
    
    fig.update_layout(
        title={
            'text': '<b>📊 Empeoraron por Tipo de Beneficio</b>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': COLORS['primary']}
        },
        xaxis_title="<b>Tipo de Beneficio</b>",
        yaxis_title="<b>Cantidad de Estudiantes</b>",
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Arial, sans-serif", size=12),
        height=450,
        margin=dict(t=60, b=50, l=50, r=50),
        showlegend=False
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor='rgba(128,128,128,0.1)')
    return fig

# ================================
#  TABLA: BENEFICIO vs NIVELES (2025-2)
# ================================
def tabla_beneficio_niveles_2025_2():
    df_2025_2 = df_becarios[df_becarios["RIESGO_2025_2"].isin(orden_niveles)]

    if df_2025_2.empty:
        return html.Div("No hay datos disponibles para mostrar", 
                       style={'textAlign': 'center', 'color': 'gray', 'padding': '20px'})

    tabla = (
        df_2025_2.groupby(["TIPO DE BENEFICIO", "RIESGO_2025_2"])
        .size()
        .reset_index(name="CANTIDAD")
    )

    tabla_pivot = tabla.pivot(index="TIPO DE BENEFICIO", columns="RIESGO_2025_2", values="CANTIDAD").fillna(0).reset_index()

    columnas_orden = ["TIPO DE BENEFICIO"] + [col for col in ["ALTO", "MEDIO", "BAJO"] if col in tabla_pivot.columns]
    tabla_pivot = tabla_pivot.reindex(columns=columnas_orden, fill_value=0)

    return dash_table.DataTable(
        columns=[{"name": col, "id": col} for col in tabla_pivot.columns],
        data=tabla_pivot.to_dict("records"),
        style_table={
            "overflowX": "auto",
            "borderRadius": "15px",
            "overflow": "hidden",
            "boxShadow": "0 4px 15px rgba(0,0,0,0.1)"
        },
        style_header={
            "backgroundColor": COLORS['warning'],
            "color": "white",
            "fontWeight": "bold",
            "textAlign": "center",
            "border": "none",
            "fontSize": "14px"
        },
        style_cell={
            "textAlign": "center",
            "backgroundColor": "#ffffff",
            "color": "#333",
            "padding": "12px",
            "border": "1px solid #e0e0e0",
            "fontSize": "13px"
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f8f9fa'
            }
        ]
    )

# ================================
#  TABLA DE RIESGO ESTILIZADA
# ================================
tabla_riesgo = dash_table.DataTable(
    columns=[{"name": i, "id": i} for i in tabla_resumen.columns],
    data=tabla_resumen.to_dict("records"),
    style_table={
        "overflowX": "auto",
        "borderRadius": "15px",
        "overflow": "hidden",
        "boxShadow": "0 4px 15px rgba(0,0,0,0.1)"
    },
    style_header={
        "backgroundColor": COLORS['primary'],
        "color": "white",
        "fontWeight": "bold",
        "textAlign": "center",
        "border": "none",
        "fontSize": "14px"
    },
    style_cell={
        "textAlign": "center",
        "backgroundColor": "#ffffff",
        "color": "#333",
        "padding": "12px",
        "border": "1px solid #e0e0e0",
        "fontSize": "13px"
    },
    style_data_conditional=[
        {
            'if': {'row_index': 'odd'},
            'backgroundColor': '#f8f9fa'
        }
    ]
)

# ================================
#  LAYOUT PRINCIPAL
# ================================
app.layout = html.Div([
    # Header con gradiente
    html.Div([
        html.Div([
            html.H1([
                html.I(className="fas fa-graduation-cap", style={'marginRight': '15px'}),
                "Dashboard de Becarios 2025"
            ], className="text-center", style={
                'color': 'white', 
                'fontWeight': 'bold',
                'fontSize': '2.5rem',
                'textShadow': '2px 2px 4px rgba(0,0,0,0.3)',
                'margin': '0'
            }),
            html.P("Análisis del Riesgo Académico", 
                   className="text-center", style={
                'color': 'rgba(255,255,255,0.9)', 
                'fontSize': '1.1rem',
                'marginTop': '10px',
                'marginBottom': '0'
            })
        ], style={'padding': '40px 0'})
    ], style={
        'background': f'linear-gradient(135deg, {COLORS["primary"]} 0%, {COLORS["info"]} 100%)',
        'marginBottom': '30px'
    }),

    dbc.Container([
        # KPIs Principales
        html.Div([
            html.H3("📈 Indicadores Clave", style={
                'color': COLORS['primary'], 
                'fontWeight': 'bold',
                'marginBottom': '25px',
                'textAlign': 'center'
            })
        ]),
        
        dbc.Row([
            dbc.Col(tarjeta_moderna(
                "Total de Becarios", 
                total_becarios, 
                COLORS['primary'], 
                "users",
                "Población total"
            ), lg=3, md=6, sm=12),
            dbc.Col(tarjeta_moderna(
                "Mejoraron", 
                mejoraron, 
                COLORS['success'], 
                "arrow-up",
                "Evolución positiva"
            ), lg=3, md=6, sm=12),
            dbc.Col(tarjeta_moderna(
                "Empeoraron", 
                empeoraron, 
                COLORS['danger'], 
                "arrow-down",
                "Evolución negativa"
            ), lg=3, md=6, sm=12),
            dbc.Col(tarjeta_moderna(
                "Se Mantuvieron", 
                se_mantuvieron, 
                COLORS['secondary'], 
                "minus",
                "Sin cambios"
            ), lg=3, md=6, sm=12),
        ], className="mb-4"),

        html.Hr(style={'border': f'1px solid {COLORS["primary"]}', 'margin': '40px 0'}),

        # Gráficos principales
        html.Div([
            html.H3("📊 Análisis Visual", style={
                'color': COLORS['primary'], 
                'fontWeight': 'bold',
                'marginBottom': '25px',
                'textAlign': 'center'
            })
        ]),

        dbc.Row([
            dbc.Col([
                html.Div([
                    dcc.Graph(figure=grafico_riesgo())
                ], style={
                    'backgroundColor': 'white',
                    'borderRadius': '15px',
                    'padding': '20px',
                    'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
                    'margin': '10px'
                })
            ], lg=8, md=12),
            dbc.Col([
                html.Div([
                    html.H5("📋 Resumen por Período", style={
                        'color': COLORS['primary'],
                        'textAlign': 'center',
                        'marginBottom': '20px'
                    }),
                    tabla_riesgo
                ], style={
                    'backgroundColor': 'white',
                    'borderRadius': '15px',
                    'padding': '20px',
                    'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
                    'margin': '10px',
                    'height': '450px',
                    'display': 'flex',
                    'flexDirection': 'column',
                    'justifyContent': 'center'
                })
            ], lg=4, md=12)
        ], className="mb-5"),

        dbc.Row([
            dbc.Col([
                html.Div([
                    dcc.Graph(figure=grafico_no_encontrado())
                ], style={
                    'backgroundColor': 'white',
                    'borderRadius': '15px',
                    'padding': '20px',
                    'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
                    'margin': '10px'
                })
            ], lg=12)
        ], className="mb-5"),

        dbc.Row([
            dbc.Col([
                html.Div([
                    dcc.Graph(figure=grafico_empeoraron_por_beneficio())
                ], style={
                    'backgroundColor': 'white',
                    'borderRadius': '15px',
                    'padding': '20px',
                    'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
                    'margin': '10px'
                })
            ], lg=12)
        ], className="mb-5"),

        html.Div([
            html.H5([
                html.I(className="fas fa-table", style={'marginRight': '10px'}),
                "Distribución de Niveles por Tipo de Beneficio (2025-2)"
            ], style={
                'color': COLORS['warning'],
                'textAlign': 'center',
                'fontWeight': 'bold',
                'marginBottom': '25px'
            })
        ]),
        
        dbc.Row([
            dbc.Col([
                html.Div([
                    tabla_beneficio_niveles_2025_2()
                ], style={
                    'backgroundColor': 'white',
                    'borderRadius': '15px',
                    'padding': '20px',
                    'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
                    'margin': '10px'
                })
            ], lg=12)
        ], className="mb-5"),

        html.Hr(style={'border': f'1px solid {COLORS["primary"]}', 'margin': '40px 0'}),

        # Sección de descarga mejorada
        dbc.Row([
            dbc.Col([
                html.Div([
                    html.H4([
                        html.I(className="fas fa-download", style={'marginRight': '10px'}),
                        "Exportar Resultados"
                    ], style={'color': COLORS['primary'], 'textAlign': 'center'}),
                    html.P("Descarga un archivo Excel con el análisis detallado por categoría de evolución", 
                           style={'textAlign': 'center', 'color': '#666', 'marginBottom': '25px'}),
                    html.Div([
                        dbc.Button([
                            html.I(className="fas fa-file-excel", style={'marginRight': '8px'}),
                            "Descargar Excel Completo"
                        ], 
                        id="btn_excel", 
                        n_clicks=0, 
                        color="success", 
                        size="lg",
                        style={
                            'borderRadius': '25px',
                            'padding': '12px 30px',
                            'fontWeight': 'bold',
                            'boxShadow': '0 4px 15px rgba(0,0,0,0.2)'
                        })
                    ], className="text-center"),
                    # Mensaje de estado
                    html.Div(id="download-status", style={
                        'textAlign': 'center', 
                        'marginTop': '15px',
                        'fontSize': '0.9rem'
                    }),
                    dcc.Download(id="download_excel")
                ], style={
                    'backgroundColor': 'white',
                    'borderRadius': '15px',
                    'padding': '30px',
                    'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
                    'margin': '10px'
                })
            ], lg=12)
        ]),

        html.Hr(style={'margin': '40px 0'}),

        # Footer
        html.Div([
            html.P([
                html.I(className="fas fa-chart-line", style={'marginRight': '8px'}),
                "Dashboard de Análisis Académico | ",
                html.Strong("Población: Becarios"),
                " | Comparativa Riesgo Académico 2025-1 vs 2025-2"
            ], style={
                'textAlign': 'center', 
                'color': '#888', 
                'fontSize': '0.9rem',
                'margin': '0'
            })
        ], style={'padding': '20px 0'})

    ], fluid=True)
], style={
    'backgroundColor': '#f8f9fa',
    'minHeight': '100vh',
    'fontFamily': 'Arial, sans-serif'
})

# ================================
#  CALLBACK CORREGIDO PARA DESCARGA
# ================================
@app.callback(
    [Output("download_excel", "data"),
     Output("download-status", "children")],
    Input("btn_excel", "n_clicks"),
    prevent_initial_call=True
)
def descargar_excel(n_clicks):
    if n_clicks == 0:
        return None, ""
    
    try:
        # Crear DataFrames para cada categoría
        df_mejoraron = df_becarios[df_becarios["EVOLUCION"] == "MEJORO"].copy()
        df_empeoraron = df_becarios[df_becarios["EVOLUCION"] == "EMPEORO"].copy()
        df_se_mantuvieron = df_becarios[
            (df_becarios["EVOLUCION"] == "SE MANTUVO") &
            (df_becarios["RIESGO_2025_1"] != "NO ENCONTRADO") &
            (df_becarios["RIESGO_2025_2"] != "NO ENCONTRADO")
        ].copy()
        
        # Crear buffer en memoria
        output = io.BytesIO()
        
        # Usar openpyxl (más universal, viene con pandas por defecto)
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Escribir hojas principales
            if not df_mejoraron.empty:
                df_mejoraron.to_excel(writer, sheet_name="Mejoraron", index=False)
            
            if not df_empeoraron.empty:
                df_empeoraron.to_excel(writer, sheet_name="Empeoraron", index=False)
            
            if not df_se_mantuvieron.empty:
                df_se_mantuvieron.to_excel(writer, sheet_name="Se Mantuvieron", index=False)
            
            # Crear hoja de resumen
            df_resumen = pd.DataFrame({
                'CATEGORIA': ['MEJORARON', 'EMPEORARON', 'SE MANTUVIERON', 'TOTAL'],
                'CANTIDAD': [
                    len(df_mejoraron), 
                    len(df_empeoraron), 
                    len(df_se_mantuvieron), 
                    len(df_becarios)
                ],
                'PORCENTAJE': [
                    f"{len(df_mejoraron)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%",
                    f"{len(df_empeoraron)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%", 
                    f"{len(df_se_mantuvieron)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%",
                    "100.0%"
                ]
            })
            
            df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
            
            # Formateo básico con openpyxl
            try:
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                
                # Colores para cada hoja
                colores_hojas = {
                    "Mejoraron": "28A745",
                    "Empeoraron": "DC3545", 
                    "Se Mantuvieron": "6C757D",
                    "Resumen": "2E86AB"
                }
                
                # Aplicar formato a cada hoja
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    color_hex = colores_hojas.get(sheet_name, "2E86AB")
                    
                    # Formatear encabezados
                    header_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF", size=12)
                    header_alignment = Alignment(horizontal="center", vertical="center")
                    
                    # Aplicar formato a la primera fila (encabezados)
                    for cell in worksheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = header_alignment
                    
                    # Ajustar anchos de columnas automáticamente
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = max(adjusted_width, 12)
                        
            except ImportError:
                # Si no se puede importar openpyxl.styles, continuar sin formato
                pass
        
        # Preparar archivo para descarga
        output.seek(0)
        
        return dcc.send_bytes(
            output.getvalue(), 
            filename="evolucion_becarios_dashboard.xlsx"
        ), html.Div([
            html.I(className="fas fa-check-circle", style={'color': 'green', 'marginRight': '5px'}),
            "¡Excel generado exitosamente!"
        ])
        
    except Exception as e:
        print(f"Error en descarga: {e}")
        return None, html.Div([
            html.I(className="fas fa-exclamation-triangle", style={'color': 'red', 'marginRight': '5px'}),
            f"Error al generar el archivo: {str(e)}"
        ])

# ================================
#  CONFIGURACIÓN PARA RENDER
# ================================
server = app.server

# Configuración adicional para Render
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    debug_mode = os.environ.get("DEBUG", "False").lower() == "true"
    app.run(
        host="0.0.0.0", 
        port=port, 
        debug=debug_mode
    )