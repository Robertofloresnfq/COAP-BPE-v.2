import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import locale
from dotenv import load_dotenv
from dateutil.relativedelta import relativedelta
import requests

from logica_informes import ejecutar_fase_1, ejecutar_fase_2

from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
import pickle
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

CLIENT_SECRETS_FILE = "client_secret.json"
SCOPES = ['https://www.googleapis.com/auth/drive'] 


# --- Funciones Auxiliares para Google Drive ---
def get_google_drive_credentials():
    creds = None
    # Intenta cargar credenciales previamente guardadas
    if os.path.exists('token.pickle'):
        try:
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        except Exception as e:
            st.warning(f"Error al cargar credenciales guardadas: {e}. Se intentará reautenticar.")
            creds = None # Forzar reautenticación si el archivo está corrupto

    # Si no hay credenciales válidas disponibles, o si expiraron, inicia el flujo de autenticación
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)
                st.success("Credenciales de Google Drive actualizadas.")
            except Exception as e:
                st.error(f"Error al refrescar el token: {e}. Por favor, vuelve a autenticarte.")
                if os.path.exists('token.pickle'):
                    os.remove('token.pickle') # Eliminar token inválido
                creds = None
        else:
            try:
                flow = Flow.from_client_secrets_file(
                    CLIENT_SECRETS_FILE,
                    scopes=SCOPES,
                    # Usa st.experimental_get_query_params para obtener la URL base de la app.
                    # Esto es más robusto para local y Cloud Run.
                    # redirect_uri=st.experimental_get_query_params().get("redirect_uri", ["http://localhost:8501"])[0] 
                    redirect_uri=st.query_params.get("redirect_uri", ["http://localhost:8501"])[0] 
                )
            except FileNotFoundError:
                st.error(f"❌ Error: Archivo '{CLIENT_SECRETS_FILE}' no encontrado. Necesitas este archivo para la autenticación de Google Drive.")
                st.info("Por favor, descarga tu `client_secret.json` desde Google Cloud Console y colócalo en la misma carpeta que tu script.")
                st.stop()
            except Exception as e:
                st.error(f"❌ Error al inicializar el flujo OAuth desde '{CLIENT_SECRETS_FILE}': {e}")
                st.stop()

            auth_url, _ = flow.authorization_url(prompt='consent')
            st.write(f"Por favor, autoriza la aplicación haciendo clic en este enlace: [Autorizar]({auth_url})")

            # query_params = st.experimental_get_query_params()
            query_params = st.query_params
            if "code" in query_params:
                auth_code = query_params["code"][0] # Obtener el primer elemento de la lista
                try:
                    flow.fetch_token(code=auth_code)
                    creds = flow.credentials
                    with open('token.pickle', 'wb') as token:
                        pickle.dump(creds, token)
                    st.success("¡Autenticado con Google Drive!")
                    # Limpiar los parámetros de la URL para evitar re-autenticación en cada recarga
                    # st.experimental_set_query_params()
                    st.query_params = {}
                    st.rerun() # Recargar la app para limpiar la URL y mostrar el estado autenticado
                except Exception as e:
                    st.error(f"Error al obtener el token de acceso: {e}. Por favor, intenta de nuevo.")
                    if os.path.exists('token.pickle'):
                        os.remove('token.pickle')
                    # st.experimental_set_query_params()
                    st.query_params = {}
                    st.stop()
            else:
                st.info("Esperando autorización de Google Drive...")
                st.stop() # Detén la ejecución hasta que el código sea recibido

    return creds

def download_file_from_drive(service, file_id: str) -> bytes:
    """Descarga el contenido de un archivo de Drive por su ID."""
    try:
        # Para archivos de Google Sheets, usa export_media para obtener XLSX
        if service.files().get(fileId=file_id, fields='mimeType').execute()['mimeType'] == 'application/vnd.google-apps.spreadsheet':
            request = service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            request = service.files().get_media(fileId=file_id)
        
        file_content = io.BytesIO()
        downloader = MediaIoBaseDownload(file_content, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        return file_content.getvalue()
    except HttpError as e:
        st.error(f"Error al descargar archivo de Drive (ID: {file_id}): {e.content.decode('utf-8')}")
        raise
    except Exception as e:
        st.error(f"Error inesperado al descargar archivo de Drive (ID: {file_id}): {e}")
        raise

def upload_file_to_drive(service, filename: str, content_bytes: bytes, mime_type: str, parent_folder_id = None) -> str:
    """Sube un archivo a Google Drive, actualizando si ya existe por nombre en la carpeta."""
    file_metadata = {'name': filename}
    if parent_folder_id:
        file_metadata['parents'] = [parent_folder_id]

    media_body = io.BytesIO(content_bytes)
    
    existing_file_id = None
    try:
        # Buscar archivo existente por nombre en la carpeta de destino
        query = f"name = '{filename}' and trashed = false"
        if parent_folder_id:
            query += f" and '{parent_folder_id}' in parents"
        response = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])
        if files:
            existing_file_id = files[0]['id']
            st.info(f"Archivo existente '{filename}' encontrado (ID: {existing_file_id}). Se actualizará.")
    except Exception as e:
        st.warning(f"No se pudo verificar la existencia del archivo: {e}")

    try:
        if existing_file_id:
            # Actualizar archivo existente
            request = service.files().update(fileId=existing_file_id,
                                             media_body=media_body,
                                             body=file_metadata,
                                             mimeType=mime_type,
                                             fields='id, name')
        else:
            # Crear nuevo archivo
            request = service.files().create(media_body=media_body,
                                             body=file_metadata,
                                             mimeType=mime_type,
                                             fields='id, name')

        response = request.execute()
        st.success(f"Archivo '{filename}' {'actualizado' if existing_file_id else 'subido'} a Google Drive con ID: {response.get('id')}")
        return response.get('id')
    except Exception as e:
        st.error(f"Error al subir archivo a Drive ('{filename}'): {e}")
        raise

def find_or_create_folder(service, folder_name: str, parent_folder_id = None) -> str:
    """Busca una carpeta por nombre dentro de una carpeta padre, o la crea si no existe."""
    query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    if parent_folder_id:
        query += f" and '{parent_folder_id}' in parents"

    try:
        response = service.files().list(q=query, spaces='drive', fields='files(id)').execute()
        folders = response.get('files', [])
        if folders:
            return folders[0]['id']
        else:
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            if parent_folder_id:
                file_metadata['parents'] = [parent_folder_id]
            
            folder = service.files().create(body=file_metadata, fields='id').execute()
            st.info(f"Carpeta '{folder_name}' creada con ID: {folder.get('id')}")
            return folder.get('id')
    except Exception as e:
        st.error(f"Error al buscar o crear carpeta '{folder_name}': {e}")
        raise

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



# --- Autenticación de Google Drive al inicio ---
google_creds = get_google_drive_credentials()
drive_service = None
if google_creds:
    try:
        drive_service = build('drive', 'v3', credentials=google_creds)
        st.session_state.drive_service = drive_service
    except Exception as e:
        st.error(f"Error al construir el servicio de Drive: {e}")
        st.session_state.drive_service = None


# --- Barra Lateral (Sidebar) para Configuración y Carga de Archivos ---
with st.sidebar:
    # Cargar credenciales desde .env para facilitar el desarrollo local
    load_dotenv()

    aws_access_key_id = os.getenv("AWS_API_ID", "")
    aws_secret_access_key = os.getenv("AWS_API_KEY", "")
    # Corregido el typo aquí:
    aws_s3_staging_dir = os.getenv("AWS_S3", "") 
    aws_region = os.getenv("AWS_REGION", "eu-west-1")
    gemini_api_key = os.getenv("GEMINI_API_KEY", "")
    
    if drive_service:
        st.success("✅ Conectado a Google Drive.")
    else:
        st.warning("Por favor, autentícate con Google Drive para usar las funciones de Drive.")

    st.subheader("📁 IDs de la plantilla ALCO de Google Drive")
    # st.info("Introduce los IDs de los archivos de Google Drive para las plantillas.")

    # Diccionario para almacenar los IDs de los archivos de Drive
    if 'drive_template_ids' not in st.session_state:
        st.session_state.drive_template_ids = {}

    st.subheader("📁 Plantillas Base")
    st.info("Carga todos los archivos de plantilla necesarios para ambos procesos.")

    # Diccionario para almacenar los archivos subidos
    if 'archivos_subidos' not in st.session_state:
        st.session_state.archivos_subidos = {}

    # Widgets para cargar archivos
    st.session_state.drive_template_ids['plantilla_coap_xlsx_id'] = st.text_input("ID de Drive: 1. Alco Mes Anterior.xlsx", key="id_coap_xlsx")
    st.session_state.archivos_subidos['plantilla_efectos'] = st.file_uploader("2. Plantilla_Efecto_Balance_Curva.xlsx", type="xlsx")
    st.session_state.archivos_subidos['plantilla_datos_medios'] = st.file_uploader("3. Plantilla_Datos_Medios.xlsx", type="xlsx")
    st.session_state.archivos_subidos['plantilla_coap_pptx'] = st.file_uploader("4. Plantilla COAP.pptx", type="pptx")
    
    # Asegurarse de que los prompts se cargan desde el sistema de archivos si existen
    prompt_main_path = "./prompt.txt"
    prompt_podcast_path = "./prompt_podcast.txt"
    
    if os.path.exists(prompt_main_path):
        st.session_state.archivos_subidos['prompt_main'] = prompt_main_path
    else:
        st.warning(f"Archivo de prompt principal no encontrado: {prompt_main_path}")
        st.session_state.archivos_subidos['prompt_main'] = None 
        
    if os.path.exists(prompt_podcast_path):
        st.session_state.archivos_subidos['prompt_podcast'] = prompt_podcast_path
    else:
        st.warning(f"Archivo de prompt podcast no encontrado: {prompt_podcast_path}")
        st.session_state.archivos_subidos['prompt_podcast'] = None 


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

    with col2:
        st.subheader("IDs de Carga (Load IDs)")
        # Expander para no ocupar mucho espacio
        with st.expander("Editar Load IDs", expanded=False):
            load_ids = {
                'cierre_base': st.text_input('cierre_base', '684833b63f68683f4f47c7c3'),
                'cierre_up': st.text_input('cierre_up', '6847dafd3f68683f4f47c7c1'),
                'cierre_dwn': st.text_input('cierre_dwn', '684862413f68683f4f4e7688'),
                'cierre_base_efecto_curva': st.text_input('cierre_base_efecto_curva', '684fdb65fb9753150202f140'),
                'cierre_up_efecto_curva': st.text_input('cierre_up_efecto_curva', '684fdd34fb9753150207669c'),
                'cierre_base_efecto_balance': st.text_input('cierre_base_efecto_balance', '684fe17afb975315020bdbf8'),
                'cierre_up_efecto_balance': st.text_input('cierre_up_efecto_balance', '684fe3b7fb975315021049bd'),
            }
    
    st.divider()

    if st.button("▶️ Ejecutar Fase 1", type="primary", use_container_width=True):
        # Validaciones de entradas
        credenciales_ok = all([aws_access_key_id, aws_secret_access_key, aws_s3_staging_dir, aws_region])
        plantillas_fase1_ok = all([
            st.session_state.archivos_subidos.get('plantilla_coap_xlsx_id'), # Usar .get() por si no se ha cargado aún
            st.session_state.archivos_subidos.get('plantilla_efectos'),
            st.session_state.archivos_subidos.get('plantilla_datos_medios')
        ])
        
        if not credenciales_ok:
            st.error("❌ Faltan credenciales de AWS en la barra lateral.")
        elif not plantillas_fase1_ok:
            st.error("❌ Faltan una o más plantillas de Excel para la Fase 1 en la barra lateral. Asegúrate de subir o cargar el archivo 'Alco Mes Anterior.xlsx'.")
        else:
            with st.spinner("Conectando a Athena y procesando datos... Este proceso puede tardar varios minutos."):
                try:
                    # Preparar argumentos para la función
                    plantillas_bytes = {
                        "coap_file_id": st.session_state.archivos_subidos['plantilla_coap_xlsx_id'],
                        "efectos": st.session_state.archivos_subidos['plantilla_efectos'].getvalue(),
                        "datos_medios": st.session_state.archivos_subidos['plantilla_datos_medios'].getvalue()
                    }
                    credenciales = {
                        "aws_id": aws_access_key_id, "aws_key": aws_secret_access_key,
                        "s3_dir": aws_s3_staging_dir, "region": aws_region
                    }

                    # Llamada a la función de lógica real
                    st.session_state.resultados_fase_1 = ejecutar_fase_1(load_ids, fecha_cierre_dt, plantillas_bytes, credenciales, drive_service)
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
            # st.download_button(
            #     label=f"Descargar {resultados['principal'][0]}", data=resultados['principal'][1],
            #     file_name=resultados['principal'][0], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            # )
            st.write(f"- **{resultados['principal'][0]}**: (Ver en Drive)(https://drive.google.com/file/d/{resultados['principal'][1]}/view?usp=sharing)")
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
            st.error("❌ Faltan una o más plantillas (PPTX o TXT de prompts) en la barra lateral. Asegúrate de que los archivos 'prompt.txt' y 'prompt_podcast.txt' existan en la misma carpeta que tu script.")
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