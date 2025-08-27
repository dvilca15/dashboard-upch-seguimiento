import pandas as pd
from dash import Dash, dcc, html, dash_table, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import io

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
        margin=dict(t=90, b=50, l=50, r=50),  # Aumentar margen superior
        legend=dict(
            title="<b>Período</b>",
            orientation="h",
            yanchor="bottom",
            y=1.05,  # Separar más la leyenda del título
            xanchor="center",
            x=0.5
        )
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor='rgba(128,128,128,0.1)')
    return fig

def grafico_no_encontrado():
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
        yaxis_title="",  # Quitar título del eje Y
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Arial, sans-serif", size=12),
        height=400,
        margin=dict(t=60, b=50, l=50, r=50),
        yaxis=dict(
            showticklabels=True, 
            categoryorder="total descending",
            type='category'  # Forzar tipo categórico para evitar interpretación de fechas
        ),
        showlegend=False  # Quitar leyenda para limpiar el gráfico
    )
    fig.update_xaxes(showgrid=True, gridcolor='rgba(128,128,128,0.1)')
    fig.update_yaxes(showgrid=False)
    return fig

# ================================
#  GRÁFICO: EMPEORARON POR BENEFICIO (AJUSTADO)
# ================================
def grafico_empeoraron_por_beneficio():
    df_empeoraron = df_becarios[df_becarios["EVOLUCION"] == "EMPEORO"]

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
    
    # Ajuste para posicionar mejor el texto en las barras
    fig.update_traces(
        textposition="auto",  # Auto selecciona la mejor posición para cada barra
        textfont=dict(size=14, weight='bold')  # Texto más visible, sin color fijo
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

    tabla = (
        df_2025_2.groupby(["TIPO DE BENEFICIO", "RIESGO_2025_2"])
        .size()
        .reset_index(name="CANTIDAD")
    )

    tabla_pivot = tabla.pivot(index="TIPO DE BENEFICIO", columns="RIESGO_2025_2", values="CANTIDAD").fillna(0).reset_index()

    columnas_orden = ["TIPO DE BENEFICIO", "ALTO", "MEDIO", "BAJO"]
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
                    'height': '450px',  # Mismo alto que el gráfico
                    'display': 'flex',
                    'flexDirection': 'column',
                    'justifyContent': 'center'  # Centrar verticalmente el contenido
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

        # Sección de descarga
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
#  CALLBACK PARA DESCARGA CON FORMATO SIMPLIFICADO
# ================================
@app.callback(
    Output("download_excel", "data"),
    Input("btn_excel", "n_clicks"),
    prevent_initial_call=True
)
def descargar_excel(n_clicks):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    output = io.BytesIO()
    
    # Crear DataFrames para cada hoja
    df_mejoraron = df_becarios[df_becarios["EVOLUCION"] == "MEJORO"]
    df_empeoraron = df_becarios[df_becarios["EVOLUCION"] == "EMPEORO"]
    df_se_mantuvieron = df_becarios[
        (df_becarios["EVOLUCION"] == "SE MANTUVO") &
        (df_becarios["RIESGO_2025_1"] != "NO ENCONTRADO") &
        (df_becarios["RIESGO_2025_2"] != "NO ENCONTRADO")
    ]
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Escribir cada DataFrame
        df_mejoraron.to_excel(writer, sheet_name="Mejoraron", index=False)
        df_empeoraron.to_excel(writer, sheet_name="Empeoraron", index=False)
        df_se_mantuvieron.to_excel(writer, sheet_name="Se Mantuvieron", index=False)
        
        # Formatear cada hoja SIN TABLAS
        hojas_info = [
            ("Mejoraron", df_mejoraron, "28A745"),      # Verde
            ("Empeoraron", df_empeoraron, "DC3545"),     # Rojo
            ("Se Mantuvieron", df_se_mantuvieron, "6C757D")  # Gris
        ]
        
        # Definir estilos
        border_thin = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        for nombre_hoja, df, color_hex in hojas_info:
            if not df.empty:
                ws = writer.sheets[nombre_hoja]
                
                max_row = len(df) + 1
                max_col = len(df.columns)
                
                # Estilos para encabezados
                header_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF", size=12)
                header_alignment = Alignment(horizontal="center", vertical="center")
                
                # Estilos para datos
                data_alignment = Alignment(horizontal="center", vertical="center")
                data_font = Font(size=10)
                
                # Aplicar formato a encabezados
                for col in range(1, max_col + 1):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment
                    cell.border = border_thin
                
                # Aplicar formato a datos con filas alternadas
                light_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
                
                for row in range(2, max_row + 1):
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.font = data_font
                        cell.border = border_thin
                        
                        # Alternar colores de filas
                        if row % 2 == 0:
                            cell.fill = light_fill
                        
                        # Centrar solo columnas numéricas/cortas
                        col_name = df.columns[col-1]
                        if col_name not in ["APELLIDOS Y NOMBRES", "DESCRIPCION"]:
                            cell.alignment = data_alignment
                
                # Ajustar anchos de columnas
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    
                    # Establecer ancho óptimo
                    adjusted_width = min(max(max_length + 2, 12), 50)
                    ws.column_dimensions[column].width = adjusted_width
                
                # Ajustar altura de filas
                ws.row_dimensions[1].height = 25  # Encabezado más alto
                for row in range(2, max_row + 1):
                    ws.row_dimensions[row].height = 20
        
        # Agregar hoja de resumen
        df_resumen = pd.DataFrame({
            'CATEGORIA': ['MEJORARON', 'EMPEORARON', 'SE MANTUVIERON', 'TOTAL'],
            'CANTIDAD': [len(df_mejoraron), len(df_empeoraron), len(df_se_mantuvieron), len(df_becarios)],
            'PORCENTAJE': [
                f"{len(df_mejoraron)/len(df_becarios)*100:.1f}%",
                f"{len(df_empeoraron)/len(df_becarios)*100:.1f}%", 
                f"{len(df_se_mantuvieron)/len(df_becarios)*100:.1f}%",
                "100.0%"
            ]
        })
        
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
        ws_resumen = writer.sheets["Resumen"]
        
        # Formatear hoja de resumen SIN TABLA
        header_fill_resumen = PatternFill(start_color="2E86AB", end_color="2E86AB", fill_type="solid")
        
        # Encabezados del resumen
        for col in range(1, 4):
            cell = ws_resumen.cell(row=1, column=col)
            cell.fill = header_fill_resumen
            cell.font = Font(bold=True, color="FFFFFF", size=12)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_thin
        
        # Datos del resumen con colores especiales
        colores_resumen = ["28A745", "DC3545", "6C757D", "FFC107"]  # Verde, Rojo, Gris, Amarillo
        
        for row in range(2, 6):
            for col in range(1, 4):
                cell = ws_resumen.cell(row=row, column=col)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border_thin
                cell.font = Font(size=11, bold=True)
                
                # Color de fila según categoría
                if row <= 5:
                    fill_color = colores_resumen[row-2] if row < 5 else "2E86AB"
                    cell.fill = PatternFill(start_color=f"{fill_color}20", end_color=f"{fill_color}20", fill_type="solid")
        
        # Ajustar anchos del resumen
        ws_resumen.column_dimensions['A'].width = 20
        ws_resumen.column_dimensions['B'].width = 15  
        ws_resumen.column_dimensions['C'].width = 15
        
        # Altura de filas del resumen
        ws_resumen.row_dimensions[1].height = 25
        for row in range(2, 6):
            ws_resumen.row_dimensions[row].height = 22
    
    output.seek(0)
    return dcc.send_bytes(output.getvalue(), "evolucion_becarios_profesional.xlsx")

# ================================
if __name__ == "__main__":
    app.run(debug=True)