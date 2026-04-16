import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import locale
from dotenv import load_dotenv
from dateutil.relativedelta import relativedelta


# Importa las funciones encapsuladas de tu otro archivo
from logica_informes_old import ejecutar_fase_1, ejecutar_fase_2

# --- Configuración de la página de Streamlit ---
st.set_page_config(
    page_title="Automatización Informes COAP",
    page_icon="🤖",
    layout="wide"
)

# Configurar locale para español para los nombres de los meses
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
    except locale.Error:
        st.warning("No se pudo configurar el idioma a español. Los nombres de los meses podrían aparecer en inglés.")

# --- Título y Descripción ---
st.title("Automatización de Informes COAP 📊")
st.markdown("Esta aplicación automatiza la generación de informes COAP, desde la extracción de datos de Athena (Fase 1) hasta la generación de comentarios con IA (Fase 2).")
st.markdown("---")


# Para evitar el problema del websocket de streamlit
@st.cache_data(show_spinner=False)
def run_fase1_cached(load_ids, fecha_cierre_dt, plantillas_bytes, credenciales):
    # Aquí simplemente llamas a tu lógica real
    return ejecutar_fase_1(load_ids, fecha_cierre_dt, plantillas_bytes, credenciales)


# --- Barra Lateral (Sidebar) para Configuración y Carga de Archivos ---
with st.sidebar:
    # st.header("⚙️ Configuración General")
    
    # Cargar credenciales desde .env para facilitar el desarrollo local
    load_dotenv()

    # st.subheader("🔑 Credenciales")
    # aws_access_key_id = st.text_input("AWS Access Key ID", value=os.getenv("AWS_API_ID", ""), type="password")
    # aws_secret_access_key = st.text_input("AWS Secret Access Key", value=os.getenv("AWS_API_KEY", ""), type="password")
    # aws_s3_staging_dir = st.text_input("AWS S3 Staging Directory", value=os.getenv("AWS_S3", ""))
    # aws_region = st.text_input("AWS Region", value=os.getenv("AWS_REGION", "eu-west-1"))
    # gemini_api_key = st.text_input("Google Gemini API Key", value=os.getenv("GEMINI_API_KEY", ""), type="password")
    
    aws_access_key_id = os.getenv("AWS_API_ID", "")
    aws_secret_access_key = os.getenv("AWS_API_KEY", "")
    aws_s3_staging_dir = ue=os.getenv("AWS_S3", "")
    aws_region = os.getenv("AWS_REGION", "eu-west-1")
    gemini_api_key = os.getenv("GEMINI_API_KEY", "")
    
    # st.info("Las credenciales son necesarias para la sesión actual y no se almacenan.")

    st.subheader("📁 Plantillas Base")
    st.info("Carga todos los archivos de plantilla necesarios para ambos procesos.")

    # Diccionario para almacenar los archivos subidos
    if 'archivos_subidos' not in st.session_state:
        st.session_state.archivos_subidos = {}

    # Widgets para cargar archivos
    st.session_state.archivos_subidos['plantilla_coap_xlsx'] = st.file_uploader("1. Alco Mes Anterior.xlsx", type="xlsx")
    st.session_state.archivos_subidos['plantilla_efectos'] = st.file_uploader("2. Plantilla_Efecto_Balance_Curva.xlsx", type="xlsx")
    st.session_state.archivos_subidos['plantilla_datos_medios'] = st.file_uploader("3. Plantilla_Datos_Medios.xlsx", type="xlsx")
    st.session_state.archivos_subidos['plantilla_coap_pptx'] = st.file_uploader("4. Plantilla COAP.pptx", type="pptx")
    st.session_state.archivos_subidos['prompt_main'] = "./prompt.txt"
    st.session_state.archivos_subidos['prompt_podcast'] = "./prompt_podcast.txt"


# --- Pestañas para cada fase ---
tab1, tab2 = st.tabs(["Fase 1: Actualizar Informes Excel", "Fase 2: Generar Comentarios con IA"])

# ============================ FASE 1 ============================
with tab1:
    st.header("Fase 1: Extracción de datos y actualización de Excel")
    st.markdown("Configura los parámetros mensuales y ejecuta la extracción desde **Amazon Athena** para rellenar los informes de Excel.")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Parámetros del Mes")
        # Widget de fecha para seleccionar el cierre
        fecha_cierre_dt = st.date_input(
            "Selecciona la fecha de cierre actual",
            value=datetime(2025, 4, 1),
            format="DD/MM/YYYY"
        )

        # Calculate the previous month
        mes_anterior_dt = fecha_cierre_dt - relativedelta(months=1)

        st.info(f"El mes de datos a procesar será: **{mes_anterior_dt.strftime('%B %Y').capitalize()}**")
        # st.info(f"El mes de datos a procesar será: **{fecha_cierre_dt.strftime('%B %Y').capitalize()}**")

    with col2:
        st.subheader("IDs de Carga (Load IDs)")
        # Expander para no ocupar mucho espacio
        with st.expander("Editar Load IDs", expanded=False):
            load_ids = {
                'cierre_base': st.text_input('cierre_base', '691450ebe0e4bd2690be54f0'),
                'cierre_up': st.text_input('cierre_up', '69145da2e0e4bd2690c64b39'),
                'cierre_dwn': st.text_input('cierre_dwn', '691466b7e0e4bd2690ca465e'),
                'cierre_base_efecto_curva': st.text_input('cierre_base_efecto_curva', '69171061d14aa54ed14d7305'),
                'cierre_up_efecto_curva': st.text_input('cierre_up_efecto_curva', '69171061d14aa54ed14d7305'),
                'cierre_base_efecto_balance': st.text_input('cierre_base_efecto_balance', '69171986d14aa54ed155694f'),
                'cierre_up_efecto_balance': st.text_input('cierre_up_efecto_balance', '69171c32d14aa54ed1597b09'),
            }
            # load_ids = {
            #     'cierre_base': st.text_input('cierre_base', '682329a922560b72fbd4d73b'),
            #     'cierre_up': st.text_input('cierre_up', '68232be222560b72fbd944ff'),
            #     'cierre_dwn': st.text_input('cierre_dwn', '6823290022560b72fbd4d73a'),
            #     'cierre_base_efecto_curva': st.text_input('cierre_base_efecto_curva', '67fe2fb4b327792994f7c9f2'),
            #     'cierre_up_efecto_curva': st.text_input('cierre_up_efecto_curva', '67fe37dbb327792994fc2931'),
            #     'cierre_base_efecto_balance': st.text_input('cierre_base_efecto_balance', '67fe3ae3b327792994008870'),
            #     'cierre_up_efecto_balance': st.text_input('cierre_up_efecto_balance', '67fe3fc0b327792994048454'),
            # }
    
    st.divider()

    if st.button("▶️ Ejecutar Fase 1", type="primary", use_container_width=True):
        # Validaciones de entradas
        credenciales_ok = all([aws_access_key_id, aws_secret_access_key, aws_s3_staging_dir, aws_region])
        plantillas_fase1_ok = all([
            st.session_state.archivos_subidos['plantilla_coap_xlsx'],
            st.session_state.archivos_subidos['plantilla_efectos'],
            st.session_state.archivos_subidos['plantilla_datos_medios']
        ])
        
        if not credenciales_ok:
            st.error("❌ Faltan credenciales de AWS en la barra lateral.")
        elif not plantillas_fase1_ok:
            st.error("❌ Faltan una o más plantillas de Excel para la Fase 1 en la barra lateral.")
        else:
            with st.spinner("Conectando a Athena y procesando datos... Este proceso puede tardar varios minutos."):
                try:
                    # Preparar argumentos para la función
                    plantillas_bytes = {
                        "coap": st.session_state.archivos_subidos['plantilla_coap_xlsx'].getvalue(),
                        "efectos": st.session_state.archivos_subidos['plantilla_efectos'].getvalue(),
                        "datos_medios": st.session_state.archivos_subidos['plantilla_datos_medios'].getvalue()
                    }
                    credenciales = {
                        "aws_id": aws_access_key_id, "aws_key": aws_secret_access_key,
                        "s3_dir": aws_s3_staging_dir, "region": aws_region
                    }
                    
                    # Llamada a la función de lógica real
                    resultados = run_fase1_cached(load_ids, fecha_cierre_dt, plantillas_bytes, credenciales)
                    st.session_state.resultados_fase_1 = resultados
                    st.session_state.fecha_cierre_fase1 = fecha_cierre_dt

                except Exception as e:
                    print("Error durante la fase 1: ", e)
                    st.error(f"Ha ocurrido un error durante la ejecución de la Fase 1: {e}")
    
    # Mostrar resultados y botones de descarga si la Fase 1 fue exitosa
    if 'resultados_fase_1' in st.session_state and st.session_state.resultados_fase_1:
        st.success("✅ ¡Fase 1 completada! Puedes descargar los archivos generados.")
        resultados = st.session_state.resultados_fase_1
        
        dl_col1, dl_col2, dl_col3 = st.columns(3)
        with dl_col1:
            st.download_button(
                label=f"Descargar {resultados['principal'][0]}", data=resultados['principal'][1],
                file_name=resultados['principal'][0], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with dl_col2:
            st.download_button(
                label=f"Descargar {resultados['efectos'][0]}", data=resultados['efectos'][1],
                file_name=resultados['efectos'][0], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with dl_col3:
            st.download_button(
                label=f"Descargar {resultados['datos_medios'][0]}", data=resultados['datos_medios'][1],
                file_name=resultados['datos_medios'][0], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        # Guardar el archivo principal para usarlo en la Fase 2
        st.session_state.excel_generado_fase1 = (resultados['principal'][0], resultados['principal'][1])


# ============================ FASE 2 ============================
with tab2:
    st.header("Fase 2: Generación de Comentarios con IA")
    st.markdown("Sube los archivos necesarios y la IA (**Gemini**) generará los comentarios para las diapositivas y un guion de podcast.")

    st.subheader("Archivos de Datos")
    
    # Usar archivo de Fase 1 o pedir que se suba
    # if 'excel_generado_fase1' in st.session_state:
    #     st.success(f"Se usará el archivo `{st.session_state.excel_generado_fase1[0]}` generado en la Fase 1.")
    #     st.session_state.archivos_subidos['alco_actual_bytes'] = st.session_state.excel_generado_fase1[1]
    # else:
    #     st.warning("No se ha ejecutado la Fase 1. Por favor, sube el archivo ALCO del mes actual.")
    
    alco_actual_upload = st.file_uploader("Archivo ALCO del mes actual", type=["xlsx", "xls"])
    if alco_actual_upload:
        st.session_state.archivos_subidos['alco_actual_bytes'] = alco_actual_upload.getvalue()

    # Pedir la presentación del mes anterior
    st.session_state.archivos_subidos['pptx_anterior'] = st.file_uploader("Presentación COAP del Mes Anterior (referencia)", type="pptx")
    
    st.divider()
    
    st.subheader("Opciones de Generación")
    # Dar al usuario el formato para el mes de cierre
    mes_cierre_fase2_str = st.text_input(
        "Confirma el Mes de Cierre del Reporte (ej: Mayo 2025)",
        value=st.session_state.get('fecha_cierre_fase1', datetime.now()).strftime("%B %Y").capitalize()
    )

    if st.button("▶️ Ejecutar Fase 2", type="primary", use_container_width=True):
        # Validaciones
        credenciales_ok = bool(gemini_api_key)
        plantillas_fase2_ok = all([
            st.session_state.archivos_subidos.get('plantilla_coap_pptx'),
            st.session_state.archivos_subidos.get('prompt_main'),
            st.session_state.archivos_subidos.get('prompt_podcast')
        ])
        archivos_datos_ok = all([
            st.session_state.archivos_subidos.get('alco_actual_bytes'),
            st.session_state.archivos_subidos.get('pptx_anterior')
        ])

        if not credenciales_ok:
            st.error("❌ Falta la API Key de Gemini en la barra lateral.")
        elif not plantillas_fase2_ok:
            st.error("❌ Faltan una o más plantillas (PPTX o TXT de prompts) en la barra lateral.")
        elif not archivos_datos_ok:
            st.error("❌ Falta el archivo ALCO del mes actual o la presentación del mes anterior.")
        else:
            with st.spinner("🤖 Generando comentarios con IA... Este proceso puede ser lento."):
                try:
                     # Lee el contenido de los archivos de prompt desde sus rutas
                    with open(st.session_state.archivos_subidos['prompt_main'], 'r', encoding='utf-8') as f:
                        prompt_main_content = f.read()
                    with open(st.session_state.archivos_subidos['prompt_podcast'], 'r', encoding='utf-8') as f:
                        prompt_podcast_content = f.read()

                    archivos_bytes = {
                        "plantilla_pptx": st.session_state.archivos_subidos['plantilla_coap_pptx'].getvalue(),
                        "prompt_main": prompt_main_content,
                        "prompt_podcast": prompt_podcast_content,
                        "alco_excel": st.session_state.archivos_subidos['alco_actual_bytes'],
                        "pptx_anterior": st.session_state.archivos_subidos['pptx_anterior'].getvalue()
                    }
                    # archivos_bytes = {
                    #     "plantilla_pptx": st.session_state.archivos_subidos['plantilla_coap_pptx'].getvalue(),
                    #     "prompt_main": st.session_state.archivos_subidos['prompt_main'].getvalue().decode('utf-8'),
                    #     "prompt_podcast": st.session_state.archivos_subidos['prompt_podcast'].getvalue().decode('utf-8'),
                    #     "alco_excel": st.session_state.archivos_subidos['alco_actual_bytes'],
                    #     "pptx_anterior": st.session_state.archivos_subidos['pptx_anterior'].getvalue()
                    # }
                    
                    # El script ya tiene una configuración por defecto, la pasamos por si se quiere cambiar
                    config_ppt = {"capture_images": False} # Desactivado por defecto para compatibilidad web
                    
                    st.session_state.resultados_fase_2 = ejecutar_fase_2(
                        mes_cierre_fase2_str, archivos_bytes, gemini_api_key, config_ppt
                    )

                except Exception as e:
                    st.error(f"Ha ocurrido un error durante la ejecución de la Fase 2: {e}")

    # Mostrar resultados de la Fase 2
    if 'resultados_fase_2' in st.session_state and st.session_state.resultados_fase_2:
        st.success("✅ ¡Fase 2 completada! Puedes descargar los archivos generados.")
        resultados = st.session_state.resultados_fase_2
        
        dl2_col1, dl2_col2, dl2_col3 = st.columns(3)
        with dl2_col1:
            st.download_button(
                label=f"Descargar {resultados['pptx'][0]}", data=resultados['pptx'][1],
                file_name=resultados['pptx'][0], mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        with dl2_col2:
            st.download_button(
                label=f"Descargar {resultados['comentarios'][0]}", data=resultados['comentarios'][1],
                file_name=resultados['comentarios'][0], mime="text/plain"
            )
        with dl2_col3:
            st.download_button(
                label=f"Descargar {resultados['podcast'][0]}", data=resultados['podcast'][1],
                file_name=resultados['podcast'][0], mime="text/plain"
            )

        with st.expander("Ver Comentarios Generados"):
            st.text(resultados['comentarios'][1].decode('utf-8'))
        with st.expander("Ver Guion de Podcast Generado"):
            st.markdown(resultados['podcast'][1].decode('utf-8'))