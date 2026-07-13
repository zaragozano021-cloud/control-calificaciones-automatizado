import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Configuración inicial de la página web
st.set_page_config(page_title="Dashboard Académico", layout="wide")

st.title("📊 Dashboard de Rendimiento Académico")
st.markdown("---")

# Cargar los datos del archivo Excel generado
@st.cache_data
def cargar_datos():
    # Reemplaza con el nombre exacto de tu archivo de salida
    file_path = "Control_Calificaciones_GitHub.xlsx" 
    base_notas = pd.read_excel(file_path, sheet_name="Base_Notas")
    asistencia = pd.read_excel(file_path, sheet_name="Control_Asistencia")
    return base_notas, asistencia

try:
    df_notas, df_asist = cargar_datos()

    # ==========================================
    # 1. SECCIÓN DE TARJETAS (KPIs GLOBALES)
    # ==========================================
    col1, col2, col3, col4 = st.columns(4)
    
    total_estudiantes = len(df_notas)
    promedio_general = df_notas["Definitiva Semestre"].mean()
    aprobados = len(df_notas[df_notas["Objetivo Aprendizaje"] == "Cumple Objetivos"])
    reprobados = len(df_notas[df_notas["Objetivo Aprendizaje"] == "No Cumple"])

    col1.metric("TOTAL ESTUDIANTES", total_estudiantes)
    col2.metric("PROMEDIO GENERAL", f"{promedio_general:.2f}")
    col3.metric("APROBADOS", aprobados)
    col4.metric("REPROBADOS", reprobados, delta_color="inverse")

    st.markdown("---")

    # ==========================================
    # 2. FILTRO SELECTOR INDIVIDUAL
    # ==========================================
    st.subheader("👤 Análisis Individual por Estudiante")
    estudiante_seleccionado = st.selectbox("Seleccione un Alumno de la Lista:", df_notas["Estudiante"].unique())

    # Filtrar datos del alumno seleccionado
    datos_alumno = df_notas[df_notas["Estudiante"] == estudiante_seleccionado].iloc[0]
   # Filtra los datos del estudiante
filtro_estudiante = df_asist[df_asist["Estudiante"] == estudiante_seleccionado]

# Verifica si el filtro encontró algo antes de intentar extraer el índice
if not filtro_estudiante.empty:
    datos_asist_alumno = filtro_estudiante.iloc[0]
else:
    st.warning("No se encontraron datos para el estudiante seleccionado.")
    datos_asist_alumno = None # O maneja el caso vacío según tu diseño

    # Mostrar información resumida del estudiante en columnas
    c1, c2, c3 = st.columns(3)
    c1.write(f"**Código:** {datos_alumno['Código']}")
    c2.write(f"**Nota Definitiva:** {datos_alumno['Definitiva Semestre']:.2f}")
    c3.write(f"**Total Faltas:** {datos_asist_alumno['Total Faltas']}")

    # ==========================================
    # 3. GRÁFICOS INTERACTIVOS (ANÁLISIS DE RAPs)
    # ==========================================
    st.markdown("### 📈 Análisis de Resultados de Aprendizaje (RAP)")
    
    # Calcular promedios grupales por corte (RAP)
    prom_grupo_c1 = df_notas["Nota C1"].mean()
    prom_grupo_c2 = df_notas["Nota C2"].mean()
    prom_grupo_c3 = df_notas["Nota C3"].mean()

    # Datos para el gráfico comparativo
    categorias = ["RAP 1: Fundamentos (C1)", "RAP 2: Aplicación (C2)", "RAP 3: Integración (C3)"]
    notas_estudiante = [datos_alumno["Nota C1"], datos_alumno["Nota C2"], datos_alumno["Nota C3"]]
    notas_grupo = [prom_grupo_c1, prom_grupo_c2, prom_grupo_c3]

    # Renderizar gráfico usando Matplotlib / Seaborn
    fig, ax = plt.subplots(figsize=(10, 4))
    sns.lineplot(x=categorias, y=notas_grupo, marker="o", label="Promedio Grupo", color="#64748B", linewidth=2.5, ax=ax)
    sns.lineplot(x=categorias, y=notas_estudiante, marker="s", label=estudiante_seleccionado, color="#0F766E", linewidth=3, ax=ax)
    
    ax.set_ylim(0, 5.2)
    ax.set_ylabel("Calificación Académica")
    ax.set_title("Comparativa: Estudiante vs Promedio Grupal")
    ax.grid(True, linestyle="--", alpha=0.5)
    
    st.pyplot(fig)

    # ==========================================
    # 4. TABLA DE ASISTENCIA DIDÁCTICA
    # ==========================================
    st.markdown("---")
    st.subheader("📅 Control de Asistencia de la Clase")
    
    # Estilizar la tabla de asistencia simulando el formato didáctico (Colores pastel)
    def colorear_asistencia(val):
        if val == 1:
            return 'background-color: #FCE8E6; color: #C5221F; font-weight: bold;' # Alerta Inasistencia
        elif str(val).upper() == 'J':
            return 'background-color: #FEF7E0; color: #B06000; font-weight: bold;' # Justificado
        elif val == 0:
            return 'background-color: #E6F4EA; color: #137333;' # Asistió
        return ''

    # Mostrar las columnas de asistencia estilizadas
    columnas_asistencia = ["Código", "Estudiante"] + [f"S{s} C{c}" for s in range(1, 17) for c in range(1, 3)] + ["Total Faltas", "Estado Asistencia"]
    df_asist_filtrado = df_asist[columnas_asistencia].dropna(subset=["Estudiante"])
    
    st.dataframe(df_asist_filtrado.style.applymap(colorear_asistencia, subset=df_asist_filtrado.columns[2:-2]), height=400)

except FileNotFoundError:
    st.error("⚠️ No se encontró el archivo de Excel. Asegúrate de ejecutar primero el script de automatización para generar 'Control_Calificaciones_GitHub.xlsx'.")
