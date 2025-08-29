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
app.title = "Dashboard Riesgo Academico 2025"

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
    """
    Gráfico de barras agrupadas mostrando becarios con/sin datos de riesgo
    """
    
    # Obtener datos
    no_encontrados_2025_1 = (df_becarios["RIESGO_2025_1"] == "NO ENCONTRADO").sum()
    no_encontrados_2025_2 = (df_becarios["RIESGO_2025_2"] == "NO ENCONTRADO").sum()
    
    # Total de becarios por período
    total_2025_1 = 3169  # Dato proporcionado
    total_2025_2 = len(df_becarios)  # Conteo actual del DataFrame
    
    # Calcular "ENCONTRADOS" (becarios con datos de riesgo)
    encontrados_2025_1 = total_2025_1 - no_encontrados_2025_1
    encontrados_2025_2 = total_2025_2 - no_encontrados_2025_2
    
    # Crear DataFrame para barras agrupadas
    datos_agrupados = pd.DataFrame({
        'PERIODO': ['2025-1', '2025-2'],
        'BECARIOS': [total_2025_1, total_2025_2],  # Mostrar el total completo
        'NO ENCONTRADOS': [no_encontrados_2025_1, no_encontrados_2025_2],
        'TOTAL': [total_2025_1, total_2025_2]
    })
    
    # Crear gráfico de barras agrupadas
    fig = px.bar(
        datos_agrupados,
        x="PERIODO",
        y=["BECARIOS", "NO ENCONTRADOS"],
        barmode="group",
        text_auto=True,
        color_discrete_map={
            "BECARIOS": "#4A90A4",    # Azul grisáceo suave
            "NO ENCONTRADOS": "#E85A4F"          # Rojo coral suave
        }
    )
    
    # Configurar texto en las barras
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=12, weight='bold')
    )
    
    # Configurar layout
    fig.update_layout(
        title={
            'text': '<b>Becarios sin Datos de Riesgo</b>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': COLORS['primary']}
        },
        xaxis_title="<b>Período Académico</b>",
        yaxis_title="<b>Número de Becarios</b>",
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Arial, sans-serif", size=12),
        height=450,
        margin=dict(t=90, b=50, l=50, r=50),
        legend=dict(
            title="<b>Composición</b>",
            orientation="h",
            yanchor="bottom",
            y=1.05,
            xanchor="center",
            x=0.5
        )
    )
    
    # Agregar totales y porcentajes encima de cada grupo de barras
    for i, periodo in enumerate(['2025-1', '2025-2']):
        total = datos_agrupados.iloc[i]['TOTAL']
        no_encontrados = datos_agrupados.iloc[i]['NO ENCONTRADOS']
        porcentaje = (no_encontrados / total * 100) if total > 0 else 0
        
        # Porcentaje sin datos (ya no necesitamos mostrar el total porque está en la barra azul)
        fig.add_annotation(
            x=periodo,
            y=max(total_2025_1, total_2025_2, no_encontrados_2025_1, no_encontrados_2025_2) + 150,
            text=f"<b>{porcentaje:.1f}%</b> sin datos",
            showarrow=False,
            font=dict(size=12, color=COLORS['danger']),
            xanchor='center'
        )
    
    # Configurar ejes
    fig.update_xaxes(
        showgrid=False,
        tickmode='array',
        tickvals=['2025-1', '2025-2'],
        ticktext=['2025-1', '2025-2']
    )
    fig.update_yaxes(showgrid=True, gridcolor='rgba(128,128,128,0.1)')
    
    return fig

# ================================
#  GRÁFICO: EMPEORARON POR MODALIDAD (MEJORADO)
# ================================
def grafico_empeoraron_por_modalidad():
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
            title="📊 Empeoraron por Modalidad",
            height=450
        )
        return fig

    # Verificar si existe la columna MODALIDAD
    if "MODALIDAD" not in df_empeoraron.columns:
        fig = go.Figure()
        fig.add_annotation(
            text="Columna 'MODALIDAD' no encontrada en los datos",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color="orange")
        )
        fig.update_layout(
            title="📊 Empeoraron por Modalidad",
            height=450
        )
        return fig

    # Crear una copia del DataFrame para modificar
    df_temp = df_empeoraron.copy()
    
    # UNIFICAR MODALIDADES CNA
    df_temp["MODALIDAD"] = df_temp["MODALIDAD"].replace({
        "Beca CNA y PA": "CNA Y PA",
        "CNA": "CNA Y PA"
    })
    
    # También podemos unificar otras modalidades similares si existen
    # Ejemplo: unificar variaciones de nombres similares
    df_temp["MODALIDAD"] = df_temp["MODALIDAD"].str.strip()  # Eliminar espacios extras

    # Agrupar por modalidad
    empeoraron_por_modalidad = (
        df_temp.groupby("MODALIDAD")
        .size()
        .reset_index(name="TOTAL")
        .sort_values("TOTAL", ascending=False)
    )

    # Agrupar modalidades con 3 o menos estudiantes en "OTROS"
    modalidades_principales = empeoraron_por_modalidad[empeoraron_por_modalidad["TOTAL"] > 3].copy()
    modalidades_menores = empeoraron_por_modalidad[empeoraron_por_modalidad["TOTAL"] <= 3]
    
    # Si hay modalidades menores, crear categoría "OTROS"
    if not modalidades_menores.empty:
        otros_total = modalidades_menores["TOTAL"].sum()
        otros_row = pd.DataFrame({
            "MODALIDAD": ["OTROS"],
            "TOTAL": [otros_total]
        })
        empeoraron_final = pd.concat([modalidades_principales, otros_row], ignore_index=True)
    else:
        empeoraron_final = modalidades_principales

    # Ordenar de mayor a menor
    empeoraron_final = empeoraron_final.sort_values("TOTAL", ascending=False)

    # FUNCIÓN MEJORADA PARA FORMATEAR ETIQUETAS
    def formatear_etiqueta(texto, max_chars_por_linea=12):
        """
        Formatea las etiquetas agregando saltos de línea basado en el ancho de la columna
        """
        if len(texto) <= max_chars_por_linea:
            return texto
        
        # Dividir por espacios
        palabras = texto.split()
        lineas = []
        linea_actual = ""
        
        for palabra in palabras:
            # Si agregar la palabra excede el límite, crear nueva línea
            test_linea = (linea_actual + " " + palabra).strip() if linea_actual else palabra
            
            if len(test_linea) > max_chars_por_linea and linea_actual:
                lineas.append(linea_actual.strip())
                linea_actual = palabra
            else:
                linea_actual = test_linea
        
        # Agregar la última línea
        if linea_actual:
            lineas.append(linea_actual.strip())
        
        # Si una palabra individual es muy larga, dividirla
        lineas_finales = []
        for linea in lineas:
            if len(linea) > max_chars_por_linea * 1.5:  # Si es muy larga
                # Dividir la línea larga
                while len(linea) > max_chars_por_linea:
                    # Buscar el mejor punto de corte
                    punto_corte = max_chars_por_linea
                    if ' ' in linea[:punto_corte]:
                        # Encontrar el último espacio antes del límite
                        punto_corte = linea.rfind(' ', 0, max_chars_por_linea)
                    
                    lineas_finales.append(linea[:punto_corte].strip())
                    linea = linea[punto_corte:].strip()
                
                if linea:  # Agregar el resto si queda algo
                    lineas_finales.append(linea)
            else:
                lineas_finales.append(linea)
        
        # Unir con salto de línea HTML
        return "<br>".join(lineas_finales)

    # Calcular el ancho disponible por columna dinámicamente
    num_columnas = len(empeoraron_final)
    if num_columnas <= 5:
        max_chars = 15
    elif num_columnas <= 8:
        max_chars = 12
    else:
        max_chars = 10

    # Calcular porcentajes
    total_empeoraron = empeoraron_final["TOTAL"].sum()
    empeoraron_final["PORCENTAJE"] = (empeoraron_final["TOTAL"] / total_empeoraron * 100).round(1)
    
    # Crear etiquetas combinadas (número + porcentaje)
    empeoraron_final["ETIQUETA_COMPLETA"] = empeoraron_final.apply(
        lambda row: f"{row['PORCENTAJE']}%<br>{row['TOTAL']}", axis=1
    )
    
    # Aplicar formateo a las modalidades
    empeoraron_final["MODALIDAD_FORMATTED"] = empeoraron_final["MODALIDAD"].apply(
        lambda x: formatear_etiqueta(x, max_chars_por_linea=max_chars)
    )

    # Crear el gráfico
    fig = px.bar(
        empeoraron_final,
        x="MODALIDAD_FORMATTED",
        y="TOTAL",
        text="ETIQUETA_COMPLETA",
        color="MODALIDAD_FORMATTED",
        color_discrete_sequence=[
            COLORS['danger'], COLORS['warning'], COLORS['info'], 
            COLORS['secondary'], COLORS['dark'], COLORS['primary'], 
            COLORS['success'], COLORS['light']
        ]
    )
    
    # AQUÍ ESTÁ LA CORRECCIÓN PRINCIPAL - Configurar el texto más específicamente
    fig.update_traces(
        textposition="outside",
        textfont=dict(size=11, weight='bold'),  # Reducir ligeramente el tamaño
        # Configurar la posición del texto con más control
        texttemplate='%{text}',
        # Agregar hover con información completa
        hovertemplate="<b>%{customdata}</b><br>" +
                      "Cantidad: %{y}<br>" +
                      "Porcentaje: %{text}<br>" +
                      "<extra></extra>",
        customdata=empeoraron_final["MODALIDAD"]  # Mostrar nombre completo en hover
    )
    
    # Calcular altura dinámica basada en las etiquetas
    max_lineas = max(empeoraron_final["MODALIDAD_FORMATTED"].apply(lambda x: x.count('<br>') + 1))
    
    # AUMENTAR LA ALTURA PARA DAR MÁS ESPACIO AL TEXTO
    valor_maximo = empeoraron_final["TOTAL"].max()
    altura_base = 600  # Aumentar altura base significativamente
    altura_extra = max(0, (max_lineas - 1) * 25)  # Más espacio por línea adicional
    # Agregar espacio extra basado en el valor máximo para el texto superior
    altura_texto_superior = max(80, valor_maximo * 0.3)  # Espacio proporcional al valor máximo
    altura_total = altura_base + altura_extra + altura_texto_superior
    
    fig.update_layout(
        title={
            'text': '<b>📊 Empeoraron por Modalidad</b>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': COLORS['primary']}
        },
        xaxis_title="<b>Modalidad</b>",
        yaxis_title="<b>Cantidad de Estudiantes</b>",
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Arial, sans-serif", size=12),
        height=altura_total,
        # AJUSTAR MÁRGENES PARA DAR MÁS ESPACIO AL TEXTO SUPERIOR
        margin=dict(t=100, b=120 + altura_extra, l=50, r=50),  # Más margen superior
        showlegend=False,
        # Configuración del eje X optimizada
        xaxis=dict(
            tickangle=0,  # Etiquetas horizontales
            tickmode='array',
            tickvals=list(range(len(empeoraron_final))),
            ticktext=empeoraron_final["MODALIDAD_FORMATTED"].tolist(),
            automargin=True,
            tickfont=dict(size=10)  # Tamaño de fuente ajustado
        ),
        # AJUSTAR EL EJE Y PARA DAR MÁS ESPACIO ARRIBA
        yaxis=dict(
            range=[0, valor_maximo * 1.25]  # Aumentar el rango superior
        )
    )
    
    fig.update_xaxes(
        showgrid=False,
        # Ajustar espaciado entre etiquetas
        categoryorder='array',
        categoryarray=empeoraron_final["MODALIDAD_FORMATTED"].tolist()
    )
    
    fig.update_yaxes(
        showgrid=True, 
        gridcolor='rgba(128,128,128,0.1)'
    )
    
    return fig

def grafico_torta_riesgo_psicologico():
    """
    Gráfico de torta mostrando la distribución del riesgo psicológico
    en estudiantes que empeoraron
    """
    # Filtrar solo estudiantes que empeoraron
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
            title="🧠 Riesgo Psicológico - Estudiantes que Empeoraron",
            height=770
        )
        return fig
    
    # Buscar la columna de riesgo psicológico
    col_psicologico = None
    for col in df_empeoraron.columns:
        if "RIESGO PSICOLÓGICO" in col.upper() and "2025" in col and "2" in col:
            col_psicologico = col
            break
    
    if not col_psicologico:
        fig = go.Figure()
        fig.add_annotation(
            text="Columna 'RIESGO PSICOLÓGICO INICIAL 2025-2' no encontrada",
            xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color="orange")
        )
        fig.update_layout(
            title="🧠 Riesgo Psicológico - Estudiantes que Empeoraron",
            height=770
        )
        return fig
    
    # Función para limpiar riesgo psicológico
    def limpiar_riesgo_psicologico(valor):
        if pd.isna(valor):
            return "SIN INFORMACIÓN"
        v = str(valor).upper().strip()
        if "BAJO" in v:
            return "BAJO"
        elif "MEDIO" in v:
            return "MEDIO"  
        elif "ALTO" in v:
            return "ALTO"
        else:
            return "SIN INFORMACIÓN"
    
    # Aplicar limpieza
    df_empeoraron_psico = df_empeoraron.copy()
    df_empeoraron_psico["RIESGO_PSICO_LIMPIO"] = df_empeoraron_psico[col_psicologico].apply(limpiar_riesgo_psicologico)
    
    # Contar distribución
    conteo_psico = df_empeoraron_psico["RIESGO_PSICO_LIMPIO"].value_counts().reset_index()
    conteo_psico.columns = ["NIVEL", "CANTIDAD"]
    
    # Definir colores para cada nivel
    colores_psico = {
        "ALTO": "#C73E1D",           # Rojo intenso
        "MEDIO": "#F18F01",          # Naranja
        "BAJO": "#81B29A",           # Verde menta
        "SIN INFORMACIÓN": "#5C4B51"  # Gris
    }
    
    # Crear lista de colores en el orden correcto
    colores_grafico = [colores_psico.get(nivel, "#5C4B51") for nivel in conteo_psico["NIVEL"]]
    
    # Calcular porcentajes
    total = conteo_psico["CANTIDAD"].sum()
    conteo_psico["PORCENTAJE"] = (conteo_psico["CANTIDAD"] / total * 100).round(1)
    
    # Crear gráfico de torta
    fig = go.Figure(data=[go.Pie(
        labels=conteo_psico["NIVEL"],
        values=conteo_psico["CANTIDAD"],
        hole=0.4,  # Dona
        marker_colors=colores_grafico,
        textinfo="label+percent+value",
        texttemplate="<b>%{label}</b><br>%{value} estudiantes<br>(%{percent})",
        textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>" +
                      "Cantidad: %{value}<br>" +
                      "Porcentaje: %{percent}<br>" +
                      "<extra></extra>",
        # Separar un poco las secciones
        pull=[0.05 if nivel == "ALTO" else 0.02 for nivel in conteo_psico["NIVEL"]]
    )])
    
    fig.update_layout(
        title={
            'text': '<b>🧠 Riesgo Psicológico en Estudiantes que Empeoraron</b><br>' +
                    f'<sub>Total: {total} estudiantes</sub>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': COLORS['primary']}
        },
        font=dict(family="Arial, sans-serif", size=12),
        height=770,
        margin=dict(t=100, b=50, l=50, r=50),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        # Configurar leyenda
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.05,
            font=dict(size=11)
        )
    )
    
    # Agregar anotación en el centro de la dona
    fig.add_annotation(
        text=f"<b>{total}</b><br>Estudiantes<br>Empeoraron",
        x=0.5, y=0.5,
        font=dict(size=14, color=COLORS['primary']),
        showarrow=False
    )
    
    return fig
# ================================
#  TABLA: MODALIDAD vs NIVELES (2025-2) - ORDENADA POR TOTAL
# ================================
def tabla_modalidad_niveles_2025_2():
    df_2025_2 = df_becarios[df_becarios["RIESGO_2025_2"].isin(orden_niveles)]

    if df_2025_2.empty:
        return html.Div("No hay datos disponibles para mostrar", 
                       style={'textAlign': 'center', 'color': 'gray', 'padding': '20px'})

    # Verificar si existe la columna MODALIDAD
    if "MODALIDAD" not in df_2025_2.columns:
        return html.Div("Columna 'MODALIDAD' no encontrada en los datos", 
                       style={'textAlign': 'center', 'color': 'orange', 'padding': '20px'})

    # TABLA ORIGINAL: Modalidad vs Niveles de Riesgo (2025-2)
    tabla_niveles = (
        df_2025_2.groupby(["MODALIDAD", "RIESGO_2025_2"])
        .size()
        .reset_index(name="CANTIDAD")
    )
    
    tabla_pivot_niveles = tabla_niveles.pivot(
        index="MODALIDAD", 
        columns="RIESGO_2025_2", 
        values="CANTIDAD"
    ).fillna(0).reset_index()

    # Asegurar que todas las columnas de niveles existan
    for nivel in ["ALTO", "MEDIO", "BAJO"]:
        if nivel not in tabla_pivot_niveles.columns:
            tabla_pivot_niveles[nivel] = 0
    
    # Reordenar columnas de niveles
    columnas_niveles = ["MODALIDAD"] + [col for col in ["ALTO", "MEDIO", "BAJO"] if col in tabla_pivot_niveles.columns]
    tabla_pivot_niveles = tabla_pivot_niveles.reindex(columns=columnas_niveles, fill_value=0)

    # NUEVA FUNCIONALIDAD: Agregar columnas de evolución por modalidad
    # Filtrar solo estudiantes que tienen datos válidos en ambos períodos (ALTO, MEDIO, BAJO)
    df_evolucion_valida = df_becarios[
        (df_becarios["RIESGO_2025_1"].isin(["ALTO", "MEDIO", "BAJO"])) & 
        (df_becarios["RIESGO_2025_2"].isin(["ALTO", "MEDIO", "BAJO"])) &
        (df_becarios["EVOLUCION"].isin(["MEJORO", "EMPEORO", "SE MANTUVO"])) &
        (df_becarios["MODALIDAD"].notna())
    ]
    
    tabla_evolucion = (
        df_evolucion_valida.groupby(["MODALIDAD", "EVOLUCION"])
        .size()
        .reset_index(name="CANTIDAD")
    )
    
    tabla_pivot_evolucion = tabla_evolucion.pivot(
        index="MODALIDAD", 
        columns="EVOLUCION", 
        values="CANTIDAD"
    ).fillna(0).reset_index()
    
    # Asegurar que existan todas las columnas de evolución
    for evolucion in ["MEJORO", "EMPEORO", "SE MANTUVO"]:
        if evolucion not in tabla_pivot_evolucion.columns:
            tabla_pivot_evolucion[evolucion] = 0
    
    # Seleccionar solo las columnas de interés para evolución
    columnas_evolucion = ["MODALIDAD", "MEJORO", "EMPEORO", "SE MANTUVO"]
    tabla_pivot_evolucion = tabla_pivot_evolucion[
        [col for col in columnas_evolucion if col in tabla_pivot_evolucion.columns]
    ]

    # COMBINAR AMBAS TABLAS
    tabla_final = pd.merge(tabla_pivot_niveles, tabla_pivot_evolucion, on="MODALIDAD", how="outer").fillna(0)
    
    # Convertir a enteros las columnas numéricas
    for col in tabla_final.columns:
        if col != "MODALIDAD":
            tabla_final[col] = tabla_final[col].astype(int)

    # Calcular TOTAL (solo de los niveles de riesgo)
    columnas_numericas_niveles = [col for col in ["ALTO", "MEDIO", "BAJO"] if col in tabla_final.columns]
    tabla_final["TOTAL_RIESGO"] = tabla_final[columnas_numericas_niveles].sum(axis=1)
    
    # Ordenar por TOTAL_RIESGO de mayor a menor
    tabla_final = tabla_final.sort_values("TOTAL_RIESGO", ascending=False)
    
    # Reorganizar columnas: MODALIDAD + NIVELES + TOTAL_RIESGO + EVOLUCIONES
    columnas_finales = ["MODALIDAD"]
    
    # Agregar columnas de niveles
    columnas_finales.extend([col for col in ["ALTO", "MEDIO", "BAJO"] if col in tabla_final.columns])
    
    # Agregar total de riesgo
    columnas_finales.append("TOTAL_RIESGO")
    
    # Agregar columnas de evolución
    columnas_finales.extend([col for col in ["MEJORO", "EMPEORO", "SE MANTUVO"] if col in tabla_final.columns])
    
    # Aplicar el orden final
    tabla_final = tabla_final[columnas_finales]

    # Crear nombres más descriptivos para las columnas
    nombres_columnas = []
    for col in tabla_final.columns:
        if col == "MODALIDAD":
            nombres_columnas.append({"name": "MODALIDAD", "id": col})
        elif col in ["ALTO", "MEDIO", "BAJO"]:
            nombres_columnas.append({"name": f"RIESGO {col}", "id": col})
        elif col == "TOTAL_RIESGO":
            nombres_columnas.append({"name": "TOTAL", "id": col})
        elif col == "MEJORO":
            nombres_columnas.append({"name": "MEJORARON", "id": col})
        elif col == "EMPEORO":
            nombres_columnas.append({"name": "EMPEORARON", "id": col})
        elif col == "SE MANTUVO":
            nombres_columnas.append({"name": "SE MANTUVIERON", "id": col})
        else:
            nombres_columnas.append({"name": col, "id": col})

    return dash_table.DataTable(
        columns=nombres_columnas,
        data=tabla_final.to_dict("records"),
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
            "fontSize": "13px",
            "whiteSpace": "normal",
            "height": "auto",
            "lineHeight": "1.2"
        },
        style_cell={
            "textAlign": "center",
            "backgroundColor": "#ffffff",
            "color": "#333",
            "padding": "10px 8px",
            "border": "1px solid #e0e0e0",
            "fontSize": "12px",
            "whiteSpace": "normal",
            "height": "auto"
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f8f9fa'
            },
            # Resaltar columna TOTAL
            {
                'if': {'column_id': 'TOTAL_RIESGO'},
                'backgroundColor': '#e8f5e8',
                'fontWeight': 'bold'
            },
            # Resaltar columnas de evolución con colores diferentes
            {
                'if': {'column_id': 'MEJORO'},
                'backgroundColor': '#e8f8e8',
                'color': '#2d5a2d'
            },
            {
                'if': {'column_id': 'EMPEORO'},
                'backgroundColor': '#fee8e8',
                'color': '#8b2635'
            },
            {
                'if': {'column_id': 'SE MANTUVO'},
                'backgroundColor': '#f0f8ff',
                'color': '#1e4d72'
            },
            # Resaltar modalidad
            {
                'if': {'column_id': 'MODALIDAD'},
                'textAlign': 'left',
                'fontWeight': 'bold',
                'backgroundColor': '#f9f9f9'
            }
        ],
        # Permitir ajuste automático del ancho de columnas
        style_cell_conditional=[
            {
                'if': {'column_id': 'MODALIDAD'},
                'width': '25%',
                'minWidth': '200px'
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

        # Análisis de estudiantes que empeoraron
        html.Div([
            html.H3("🔍 Análisis Detallado - Estudiantes que Empeoraron", style={
                'color': COLORS['danger'], 
                'fontWeight': 'bold',
                'marginBottom': '25px',
                'textAlign': 'center'
            })
        ]),

        dbc.Row([
            dbc.Col([
                html.Div([
                    dcc.Graph(figure=grafico_empeoraron_por_modalidad())
                ], style={
                    'backgroundColor': 'white',
                    'borderRadius': '15px',
                    'padding': '20px',
                    'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
                    'margin': '10px'
                })
            ], lg=7, md=12),
            dbc.Col([
                html.Div([
                    dcc.Graph(figure=grafico_torta_riesgo_psicologico())
                ], style={
                    'backgroundColor': 'white',
                    'borderRadius': '15px',
                    'padding': '20px',
                    'boxShadow': '0 8px 25px rgba(0,0,0,0.1)',
                    'margin': '10px'
                })
            ], lg=5, md=12)
        ], className="mb-5"),

        html.Div([
            html.H5([
                html.I(className="fas fa-table", style={'marginRight': '10px'}),
                "Distribución de Niveles por Modalidad (2025-2) - Ordenado por Total"
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
                    tabla_modalidad_niveles_2025_2()
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
        # HOJAS EXISTENTES - Crear DataFrames para cada categoría de evolución
        df_mejoraron = df_becarios[df_becarios["EVOLUCION"] == "MEJORO"].copy()
        df_empeoraron = df_becarios[df_becarios["EVOLUCION"] == "EMPEORO"].copy()
        df_se_mantuvieron = df_becarios[
            (df_becarios["EVOLUCION"] == "SE MANTUVO") &
            (df_becarios["RIESGO_2025_1"] != "NO ENCONTRADO") &
            (df_becarios["RIESGO_2025_2"] != "NO ENCONTRADO")
        ].copy()
        
        # NUEVAS HOJAS - Crear DataFrames por disponibilidad de datos de riesgo
        
        # 1. Estudiantes con riesgo SOLO en 2025-1 (tienen datos en 2025-1, NO en 2025-2)
        df_solo_2025_1 = df_becarios[
            (df_becarios["RIESGO_2025_1"].isin(["ALTO", "MEDIO", "BAJO"])) &
            (df_becarios["RIESGO_2025_2"] == "NO ENCONTRADO")
        ].copy()
        
        # 2. Estudiantes con riesgo SOLO en 2025-2 (tienen datos en 2025-2, NO en 2025-1)  
        df_solo_2025_2 = df_becarios[
            (df_becarios["RIESGO_2025_1"] == "NO ENCONTRADO") &
            (df_becarios["RIESGO_2025_2"].isin(["ALTO", "MEDIO", "BAJO"]))
        ].copy()
        
        # 3. Estudiantes SIN información de riesgo en NINGÚN período
        df_sin_informacion = df_becarios[
            (df_becarios["RIESGO_2025_1"] == "NO ENCONTRADO") &
            (df_becarios["RIESGO_2025_2"] == "NO ENCONTRADO")
        ].copy()
        
        # Crear buffer en memoria
        output = io.BytesIO()
        
        # Usar openpyxl para crear el archivo Excel
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # HOJAS PRINCIPALES (por evolución)
            if not df_mejoraron.empty:
                df_mejoraron.to_excel(writer, sheet_name="Mejoraron", index=False)
            
            if not df_empeoraron.empty:
                df_empeoraron.to_excel(writer, sheet_name="Empeoraron", index=False)
            
            if not df_se_mantuvieron.empty:
                df_se_mantuvieron.to_excel(writer, sheet_name="Se Mantuvieron", index=False)
            
            # NUEVAS HOJAS (por disponibilidad de datos)
            if not df_solo_2025_1.empty:
                df_solo_2025_1.to_excel(writer, sheet_name="Solo Riesgo 2025-1", index=False)
            
            if not df_solo_2025_2.empty:
                df_solo_2025_2.to_excel(writer, sheet_name="Solo Riesgo 2025-2", index=False)
            
            if not df_sin_informacion.empty:
                df_sin_informacion.to_excel(writer, sheet_name="Sin Informacion", index=False)
            
            # HOJA DE RESUMEN AMPLIADA
            df_resumen = pd.DataFrame({
                'CATEGORIA': [
                    'MEJORARON', 
                    'EMPEORARON', 
                    'SE MANTUVIERON',
                    'SOLO TIENEN RIESGO 2025-1',
                    'SOLO TIENEN RIESGO 2025-2', 
                    'SIN INFORMACIÓN AMBOS PERÍODOS',
                    'TOTAL BECARIOS'
                ],
                'CANTIDAD': [
                    len(df_mejoraron), 
                    len(df_empeoraron), 
                    len(df_se_mantuvieron),
                    len(df_solo_2025_1),
                    len(df_solo_2025_2),
                    len(df_sin_informacion),
                    len(df_becarios)
                ],
                'PORCENTAJE': [
                    f"{len(df_mejoraron)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%",
                    f"{len(df_empeoraron)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%", 
                    f"{len(df_se_mantuvieron)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%",
                    f"{len(df_solo_2025_1)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%",
                    f"{len(df_solo_2025_2)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%",
                    f"{len(df_sin_informacion)/len(df_becarios)*100:.1f}%" if len(df_becarios) > 0 else "0.0%",
                    "100.0%"
                ],
                'DESCRIPCION': [
                    'Estudiantes que redujeron su nivel de riesgo',
                    'Estudiantes que aumentaron su nivel de riesgo',
                    'Estudiantes que mantuvieron el mismo nivel',
                    'Solo aparecen en registro 2025-1',
                    'Solo aparecen en registro 2025-2', 
                    'No tienen datos de riesgo en ningún período',
                    'Total de becarios en la base de datos'
                ]
            })
            
            df_resumen.to_excel(writer, sheet_name="Resumen General", index=False)
            
            # Formateo con openpyxl (si está disponible)
            try:
                from openpyxl.styles import Font, PatternFill, Alignment
                
                # Colores para cada tipo de hoja
                colores_hojas = {
                    "Mejoraron": "28A745",           # Verde
                    "Empeoraron": "DC3545",          # Rojo
                    "Se Mantuvieron": "6C757D",      # Gris
                    "Solo Riesgo 2025-1": "FFC107", # Amarillo
                    "Solo Riesgo 2025-2": "17A2B8", # Azul claro
                    "Sin Informacion": "6F42C1",     # Púrpura
                    "Resumen General": "2E86AB"      # Azul principal
                }
                
                # Aplicar formato a cada hoja
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    color_hex = colores_hojas.get(sheet_name, "2E86AB")
                    
                    # Formatear encabezados
                    header_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF", size=12)
                    header_alignment = Alignment(horizontal="center", vertical="center")
                    
                    # Aplicar formato a la primera fila
                    for cell in worksheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = header_alignment
                    
                    # Ajustar anchos de columnas
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
                # Continuar sin formato si openpyxl.styles no está disponible
                pass
        
        # Preparar archivo para descarga
        output.seek(0)
        
        return dcc.send_bytes(
            output.getvalue(), 
            filename="analisis_completo_becarios_2025.xlsx"
        ), html.Div([
            html.I(className="fas fa-check-circle", style={'color': 'green', 'marginRight': '5px'}),
            f"¡Excel generado con {len(writer.sheets)} hojas!"
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