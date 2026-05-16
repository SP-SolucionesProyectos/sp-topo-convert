# =========================================================
# SP Topo-Convert V7 | PARTE 1
# Base limpia: imports, config, sesión, logo y helpers
# =========================================================

import os
import io
import re
import json
import time
import uuid
import string
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st
import pyproj
import folium
import ezdxf

from PIL import Image
from streamlit_folium import folium_static
from streamlit_cookies_manager import EncryptedCookieManager
from pathlib import Path
from contextlib import contextmanager
from datetime import datetime, timedelta, date
from google.oauth2.service_account import Credentials
import gspread
# =========================================================
# CONFIGURACIÓN DE PÁGINA
# =========================================================

st.set_page_config(
    page_title="SP Topo-Convert",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# ESTILO VISUAL
# =========================================================

st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {
        background-color: #0E1117;
    }
    .stApp {
        background: #F5F7FA;
    }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0B1F3A 0%, #123C73 100%);
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }
    h1, h2, h3, h4, h5, h6 {
        color: #0A3D62;
    }
    .small-muted {
        color: #607080;
        font-size: 0.92rem;
    }
    .card-res {
        background: white;
        border-left: 8px solid #0A3D62;
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.08);
        margin-bottom: 12px;
    }
    .stButton > button {
        background: #0A3D62;
        color: white;
        border-radius: 12px;
        border: none;
        font-weight: 700;
        height: 3rem;
    }
    .stDownloadButton > button {
        background: #F6B93B;
        color: #1f1f1f;
        border-radius: 12px;
        border: none;
        font-weight: 700;
        height: 3rem;
    }
    .footer-box {
        text-align: center;
        color: #607080;
        padding: 18px 0 6px 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# CONSTANTES GLOBALES
# =========================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SHEET_ID = "1olSmRjPHBYV-NTc-GtrVWU2DBB13haPIixZfz6QzUOw"

APP_NAME = "SP Topo-Convert"
APP_VERSION = "7.0.0"
FREE_DAILY_CREDITS = 5
FREE_MAX_ROWS = 20
FREE_EXPORTS_PER_DAY = 5
MAX_MASSIVE_ROWS = 50000
MAX_PREVIEW_ROWS = 100
VALID_ZONES = ["17S", "18S", "19S"]
VALID_INPUT_TYPES = ["UTM", "GEO"]
VALID_ACTIONS = ["convertir", "ubicar"]
VALID_CONVERSIONS = ["PSAD56_to_WGS84", "WGS84_to_PSAD56"]
BASE_COLUMNS = ["PUNTO", "ESTE", "NORTE", "LAT", "LON", "Z", "DESCRIPCION", "ERROR"]
PLANES = ["FREE", "DIARIO", "SEMANAL", "MENSUAL", "ANUAL", "PRO", "ADMIN"]

LOGO_PATH = Path("logo.png")
PERU_CENTER = (-9.189967, -75.015152)

VALID_DATUMS = ["PSAD56", "WGS84"]

COLUMN_ALIASES = {
    "PUNTO": ["PUNTO", "ID", "PTO", "POINT", "NOMBRE"],
    "ESTE": ["ESTE", "EAST", "X"],
    "NORTE": ["NORTE", "NORTH", "Y"],
    "LAT": ["LAT", "LATITUD", "LATITUDE"],
    "LON": ["LON", "LONGITUD", "LONGITUDE"],
    "Z": ["Z", "COTA", "ELEVACION", "ELEV", "ALTURA"],
    "DESCRIPCION": ["DESCRIPCION", "DESC", "OBS", "DETALLE"],
}

# =========================================================
# SESSION STATE
# =========================================================

def init_session_state():
    defaults = {
        "fecha_expiracion": "",
        "ultima_fecha_free": "",
        "user_id": "anon",
        "local_user_id": "",
        "plan": "FREE",
        "tipo_licencia": "NINGUNA",
        "es_pro": False,
        "es_admin": False,
        "credits_free": FREE_DAILY_CREDITS,
        "credits_used_today": 0,
        "free_massive_used": 0,
        "free_exports_used": 0,
        "ultima_fecha_free": "",
        "df_original": None,
        "df_resultado": None,
        "df_export": None,
        "errores": [],
        "config": {},
        "menu_actual": "CONVERTIR",
        "submodulo_actual": "MANUAL",
        "onboarding_visto": False,
        "last_message": "",
        "procesando": False,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# =========================================================
# COOKIES
# =========================================================

cookies = EncryptedCookieManager(
    password=os.environ.get("COOKIE_PASSWORD", "SP_Topo_Secure_Key_2026_JGZ"),
    prefix="sp_topo/",
)

if not cookies.ready():
    st.stop()


# =========================================================
# HELPERS BÁSICOS
# =========================================================

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_str():
    return datetime.now().strftime("%Y-%m-%d")


def generar_uuid():
    return str(uuid.uuid4())


def get_or_create_user_id():
    if not st.session_state.local_user_id:
        cookie_user = cookies.get("local_user_id", "")
        if cookie_user:
            st.session_state.local_user_id = cookie_user
        else:
            st.session_state.local_user_id = generar_uuid()
            cookies["local_user_id"] = st.session_state.local_user_id
            cookies.save()
    return st.session_state.local_user_id


def normalizar_texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and np.isnan(valor):
        return ""
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def sanitizar_texto(valor):
    texto = normalizar_texto(valor)
    texto = texto.replace("\n", " ").replace("\r", " ")
    texto = re.sub(r"[<>]", "", texto)
    return texto[:200]


def limpiar_numero(valor):
    if valor is None:
        return np.nan
    if pd.isna(valor):
        return np.nan
    if isinstance(valor, str):
        valor = valor.strip().replace(",", ".")
        if valor == "":
            return np.nan
    try:
        return float(valor)
    except Exception:
        return np.nan


def letras_excel(indice):
    resultado = ""
    n = int(indice)
    while n >= 0:
        resultado = string.ascii_uppercase[n % 26] + resultado
        n = n // 26 - 1
    return resultado


def generar_nombre_archivo(base, ext):
    fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base}_{fecha}.{ext}"


def mostrar_mensaje(msg, tipo="info"):
    st.session_state.last_message = msg
    if tipo == "success":
        st.success(msg)
    elif tipo == "warning":
        st.warning(msg)
    elif tipo == "error":
        st.error(msg)
    else:
        st.info(msg)


# =========================================================
# LOGO Y HEADER BASE
# =========================================================

def cargar_logo(path="logo.png"):
    try:
        if os.path.exists(path):
            return Image.open(path)
    except Exception:
        return None
    return None


def render_logo_sidebar():
    logo = cargar_logo()
    with st.sidebar:
        if logo is not None:
            st.image(logo, use_container_width=True)
        st.markdown(f"### {APP_NAME}")
        st.caption(f"Versión {APP_VERSION}")


def render_header_base(titulo, subtitulo=""):
    st.markdown(f"# {titulo}")
    if subtitulo:
        st.markdown(f"<div class='small-muted'>{subtitulo}</div>", unsafe_allow_html=True)


# =========================================================
# HELPERS DE DATOS
# =========================================================

def crear_df_base_vacio():
    return pd.DataFrame(columns=BASE_COLUMNS)


def crear_df_manual_utm(punto, este, norte, z=None, descripcion=None):
    return pd.DataFrame([
        {
            "PUNTO": punto,
            "ESTE": este,
            "NORTE": norte,
            "LAT": np.nan,
            "LON": np.nan,
            "Z": z if z is not None else np.nan,
            "DESCRIPCION": descripcion if descripcion is not None else "",
            "ERROR": "",
        }
    ])


def crear_df_manual_geo(punto, lat, lon, z=None, descripcion=None):
    return pd.DataFrame([
        {
            "PUNTO": punto,
            "ESTE": np.nan,
            "NORTE": np.nan,
            "LAT": lat,
            "LON": lon,
            "Z": z if z is not None else np.nan,
            "DESCRIPCION": descripcion if descripcion is not None else "",
            "ERROR": "",
        }
    ])


def normalizar_base(df):
    if df is None or df.empty:
        return crear_df_base_vacio()

    df = df.copy()
    for col in BASE_COLUMNS:
        if col not in df.columns:
            df[col] = np.nan if col in ["ESTE", "NORTE", "LAT", "LON", "Z"] else ""
    return df[BASE_COLUMNS].copy()


# =========================================================
# UI BASE
# =========================================================

def render_footer_base():
    st.markdown("---")
    st.markdown(
        "<div class='footer-box'>SP Topo-Convert | Procesamiento geoespacial profesional</div>",
        unsafe_allow_html=True,
    )


def render_sidebar_base():
    with st.sidebar:
        st.divider()
        st.markdown("## Navegación")
        if st.button("Convertir", use_container_width=True):
            st.session_state.menu_actual = "CONVERTIR"
        if st.button("Ubicar", use_container_width=True):
            st.session_state.menu_actual = "UBICAR"
        if st.button("Planes", use_container_width=True):
            st.session_state.menu_actual = "PLANES"
        if st.button("Soporte", use_container_width=True):
            st.session_state.menu_actual = "SOPORTE"
        st.divider()
        st.markdown("## Estado")
        st.caption(f"Plan: {st.session_state.plan}")
        st.caption(f"Créditos free: {st.session_state.credits_free}")
        st.caption(f"Usuario: {st.session_state.local_user_id[:8] if st.session_state.local_user_id else 'anon'}")


# =========================================================
# INICIALIZACIÓN
# =========================================================

init_session_state()
get_or_create_user_id()
render_logo_sidebar()
# =========================================================
# SP Topo-Convert V7 | PARTE 2
# Google Sheets, usuarios, licencias, logs y tickets
# =========================================================


# =========================================================
# CONEXIÓN GOOGLE SHEETS
# =========================================================

@st.cache_resource
def connect_gsheet():

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        st.secrets["SHEET_ID"]
    )

    return spreadsheet


def load_sheet(sheet_name):
    try:
        ss = connect_gsheet()
        if ss is None:
            return pd.DataFrame()

        ws = ss.worksheet(sheet_name)
        records = ws.get_all_records()
        return pd.DataFrame(records)

    except Exception as e:
        mostrar_mensaje(f"Error leyendo hoja {sheet_name}: {e}", "error")
        return pd.DataFrame()


def save_sheet(sheet_name, df):
    try:
        ss = connect_gsheet()
        if ss is None:
            return False

        ws = ss.worksheet(sheet_name)
        df2 = df.copy().fillna("")
        data = [df2.columns.tolist()] + df2.astype(str).values.tolist()

        ws.clear()
        ws.update(data)
        return True

    except Exception as e:
        mostrar_mensaje(f"Error guardando hoja {sheet_name}: {e}", "error")
        return False

# =========================================================
# CARGAR TABLAS
# =========================================================

def get_usuarios_df():
    df = load_sheet("Usuarios")

    if df.empty:
        columns = [
            "user_id",
            "plan",
            "tipo_licencia",
            "fecha_activacion",
            "fecha_expiracion",
            "creditos_disponibles",
            "ultima_fecha_free",
            "ultima_ip",
        ]
        df = pd.DataFrame(columns=columns)

    return df


def get_licencias_df():
    df = load_sheet("Licencias")

    if df.empty:
        columns = [
            "codigo",
            "plan_tipo",
            "estado",
            "usado_por",
            "fecha_uso",
        ]
        df = pd.DataFrame(columns=columns)

    return df


def get_logs_df():
    df = load_sheet("Logs")

    if df.empty:
        columns = [
            "fecha",
            "user_id",
            "licencia_activa",
            "accion",
            "nombre_archivo",
            "puntos_ok",
            "errores_filas",
            "tiempo_ejecucion",
        ]
        df = pd.DataFrame(columns=columns)

    return df


def get_tickets_df():
    df = load_sheet("Tickets")

    if df.empty:
        columns = [
            "ticket_id",
            "user_id",
            "mensaje",
            "estado",
            "respuesta",
            "fecha",
        ]
        df = pd.DataFrame(columns=columns)

    return df


# =========================================================
# USUARIOS
# =========================================================

def buscar_usuario(user_id):
    df = get_usuarios_df()

    if df.empty:
        return None

    result = df[df["user_id"].astype(str) == str(user_id)]

    if result.empty:
        return None

    return result.iloc[0].to_dict()


def crear_usuario_free(user_id):
    df = get_usuarios_df()

    nuevo = {
        "user_id": user_id,
        "plan": "FREE",
        "tipo_licencia": "NINGUNA",
        "fecha_activacion": "",
        "fecha_expiracion": "",
        "creditos_disponibles": FREE_DAILY_CREDITS,
        "ultima_fecha_free": today_str(),
        "ultima_ip": "",
    }

    df = pd.concat([df, pd.DataFrame([nuevo])], ignore_index=True)

    save_sheet("Usuarios", df)

    return nuevo


def actualizar_usuario(user_id, updates):
    df = get_usuarios_df()

    if df.empty:
        return False

    idx = df[df["user_id"].astype(str) == str(user_id)].index

    if len(idx) == 0:
        return False

    idx = idx[0]

    for key, value in updates.items():
        if key in df.columns:
            df.at[idx, key] = value

    return save_sheet("Usuarios", df)


def cargar_usuario_session(user_data):
    st.session_state.user_id = user_data.get("user_id", "")
    st.session_state.plan = user_data.get("plan", "FREE")
    st.session_state.tipo_licencia = user_data.get("tipo_licencia", "NINGUNA")

    try:
        st.session_state.credits_free = int(
            user_data.get("creditos_disponibles", FREE_DAILY_CREDITS)
        )
    except:
        st.session_state.credits_free = FREE_DAILY_CREDITS

    st.session_state.es_admin = (
        str(user_data.get("tipo_licencia", "")).upper() == "ADMIN"
    )

    st.session_state.es_pro = (
        st.session_state.plan != "FREE"
    )


def iniciar_usuario():
    user_id = st.session_state.local_user_id

    user_data = buscar_usuario(user_id)

    if user_data is None:
        user_data = crear_usuario_free(user_id)

    cargar_usuario_session(user_data)


# =========================================================
# RESET CRÉDITOS FREE
# =========================================================

def reset_creditos_diarios():
    if st.session_state.plan != "FREE":
        return

    user_data = buscar_usuario(st.session_state.user_id)

    if user_data is None:
        return

    ultima_fecha = str(user_data.get("ultima_fecha_free", ""))

    if ultima_fecha != today_str():

        actualizar_usuario(
            st.session_state.user_id,
            {
                "creditos_disponibles": FREE_DAILY_CREDITS,
                "ultima_fecha_free": today_str(),
            },
        )

        st.session_state.credits_free = FREE_DAILY_CREDITS


# =========================================================
# LICENCIAS
# =========================================================

def buscar_licencia(codigo):
    df = get_licencias_df()

    if df.empty:
        return None

    result = df[df["codigo"].astype(str) == str(codigo)]

    if result.empty:
        return None

    return result.iloc[0].to_dict()


def activar_licencia(codigo):
    licencia = buscar_licencia(codigo)

    if licencia is None:
        return False, "Licencia no encontrada"

    estado = str(licencia.get("estado", "")).upper()

    if estado == "USADO":
        return False, "Licencia ya utilizada"

    plan_tipo = str(licencia.get("plan_tipo", "")).upper()

    ahora = datetime.now()

    expiracion = ""

    if plan_tipo == "DIARIO":
        expiracion = ahora + timedelta(days=1)

    elif plan_tipo == "SEMANAL":
        expiracion = ahora + timedelta(days=7)

    elif plan_tipo == "MENSUAL":
        expiracion = ahora + timedelta(days=30)

    elif plan_tipo == "ANUAL":
        expiracion = ahora + timedelta(days=365)

    elif plan_tipo == "ADMIN":
        expiracion = ahora + timedelta(days=3650)

    actualizar_usuario(
        st.session_state.user_id,
        {
            "plan": "PRO" if plan_tipo != "ADMIN" else "ADMIN",
            "tipo_licencia": plan_tipo,
            "fecha_activacion": now_str(),
            "fecha_expiracion": expiracion.strftime("%Y-%m-%d %H:%M:%S"),
            "creditos_disponibles": 999999,
        },
    )

    df = get_licencias_df()

    idx = df[df["codigo"].astype(str) == str(codigo)].index

    if len(idx) > 0:

        idx = idx[0]

        df.at[idx, "estado"] = "USADO"
        df.at[idx, "usado_por"] = st.session_state.user_id
        df.at[idx, "fecha_uso"] = now_str()

        save_sheet("Licencias", df)

    iniciar_usuario()

    return True, f"Licencia {plan_tipo} activada"


# =========================================================
# VALIDAR EXPIRACIÓN
# =========================================================

def validar_expiracion_pro():
    if st.session_state.plan == "FREE":
        return

    user_data = buscar_usuario(st.session_state.user_id)

    if user_data is None:
        return

    fecha_exp = str(user_data.get("fecha_expiracion", "")).strip()

    if fecha_exp == "":
        return

    try:

        fecha_exp = datetime.strptime(
            fecha_exp,
            "%Y-%m-%d %H:%M:%S",
        )

        if datetime.now() > fecha_exp:

            actualizar_usuario(
                st.session_state.user_id,
                {
                    "plan": "FREE",
                    "tipo_licencia": "NINGUNA",
                    "creditos_disponibles": FREE_DAILY_CREDITS,
                },
            )

            iniciar_usuario()

            mostrar_mensaje(
                "Tu licencia expiró, regresaste al plan FREE",
                "warning",
            )

    except:
        pass


# =========================================================
# LOGS
# =========================================================

def registrar_log(
    accion="",
    nombre_archivo="",
    puntos_ok=0,
    errores_filas=0,
    tiempo_ejecucion=0,
):

    df = get_logs_df()

    nuevo = {
        "fecha": now_str(),
        "user_id": st.session_state.user_id,
        "licencia_activa": st.session_state.tipo_licencia,
        "accion": accion,
        "nombre_archivo": nombre_archivo,
        "puntos_ok": puntos_ok,
        "errores_filas": errores_filas,
        "tiempo_ejecucion": tiempo_ejecucion,
    }

    df = pd.concat([df, pd.DataFrame([nuevo])], ignore_index=True)

    save_sheet("Logs", df)


# =========================================================
# TICKETS
# =========================================================

def crear_ticket(mensaje):
    df = get_tickets_df()

    nuevo = {
        "ticket_id": generar_uuid(),
        "user_id": st.session_state.user_id,
        "mensaje": sanitizar_texto(mensaje),
        "estado": "ABIERTO",
        "respuesta": "",
        "fecha": now_str(),
    }

    df = pd.concat([df, pd.DataFrame([nuevo])], ignore_index=True)

    save_sheet("Tickets", df)

    return True


def listar_mis_tickets():
    df = get_tickets_df()

    if df.empty:
        return pd.DataFrame()

    return df[
        df["user_id"].astype(str)
        == str(st.session_state.user_id)
    ]


# =========================================================
# INICIALIZAR USUARIO REAL
# =========================================================

iniciar_usuario()
reset_creditos_diarios()
validar_expiracion_pro()

# =========================================================
# SP Topo-Convert V7 | PARTE 3
# Validadores, archivos, columnas y normalización
# =========================================================


# =========================================================
# VALIDAR EXTENSIONES
# =========================================================

VALID_FILE_EXTENSIONS = [
    "csv",
    "xlsx",
    "xls",
    "txt",
]


def obtener_extension_archivo(filename):
    if "." not in filename:
        return ""
    return filename.split(".")[-1].lower()


def validar_archivo_subido(uploaded_file):

    if uploaded_file is None:
        return False, "No se subió archivo"

    ext = obtener_extension_archivo(uploaded_file.name)

    if ext not in VALID_FILE_EXTENSIONS:
        return (
            False,
            "Formato no válido. Usa CSV, XLSX, XLS o TXT",
        )

    return True, "OK"


# =========================================================
# LEER ARCHIVOS
# =========================================================

def leer_csv(uploaded_file):
    try:
        return pd.read_csv(uploaded_file)
    except:
        return pd.read_csv(
            uploaded_file,
            sep=";",
            encoding="latin-1",
        )


def leer_excel(uploaded_file):
    return pd.read_excel(uploaded_file)


def leer_txt(uploaded_file):
    try:
        return pd.read_csv(
            uploaded_file,
            delim_whitespace=True,
            header=None,
        )
    except:
        return pd.read_csv(
            uploaded_file,
            sep=";",
            header=None,
        )


def leer_archivo(uploaded_file):

    valido, mensaje = validar_archivo_subido(uploaded_file)

    if not valido:
        raise Exception(mensaje)

    ext = obtener_extension_archivo(uploaded_file.name)

    if ext == "csv":
        return leer_csv(uploaded_file)

    if ext in ["xlsx", "xls"]:
        return leer_excel(uploaded_file)

    if ext == "txt":
        return leer_txt(uploaded_file)

    raise Exception("Formato no soportado")


# =========================================================
# VALIDAR DATAFRAME
# =========================================================

def validar_dataframe(df):

    if df is None:
        return False, "DataFrame vacío"

    if df.empty:
        return False, "El archivo no contiene datos"

    if len(df.columns) == 0:
        return False, "No se detectaron columnas"

    return True, "OK"


# =========================================================
# NORMALIZAR COLUMNAS
# =========================================================

def normalize_column_name(name):

    name = normalizar_texto(name)

    name = name.upper()

    name = (
        name.replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )

    name = re.sub(r"[^A-Z0-9]", "", name)

    return name


def obtener_columnas_normalizadas(df):

    mapping = {}

    for col in df.columns:
        mapping[col] = normalize_column_name(col)

    return mapping


# =========================================================
# DETECCIÓN AUTOMÁTICA
# =========================================================

AUTO_COLUMN_PATTERNS = {
    "PUNTO": [
        "PUNTO",
        "ID",
        "PTO",
        "POINT",
        "NOMBRE",
    ],
    "ESTE": [
        "ESTE",
        "EAST",
        "X",
    ],
    "NORTE": [
        "NORTE",
        "NORTH",
        "Y",
    ],
    "LAT": [
        "LAT",
        "LATITUD",
        "LATITUDE",
    ],
    "LON": [
        "LON",
        "LONGITUD",
        "LONGITUDE",
    ],
    "Z": [
        "Z",
        "COTA",
        "ELEVACION",
        "ELEV",
        "ALTURA",
    ],
    "DESCRIPCION": [
        "DESCRIPCION",
        "DESC",
        "OBS",
        "DETALLE",
    ],
}


def detectar_columnas_automaticamente(df):

    detected = {}

    normalized = obtener_columnas_normalizadas(df)

    for original_col, normalized_col in normalized.items():

        for target, patterns in AUTO_COLUMN_PATTERNS.items():

            for pattern in patterns:

                if pattern in normalized_col:

                    if target not in detected:
                        detected[target] = original_col

    return detected


# =========================================================
# COLUMNAS EXCEL A/B/C
# =========================================================

def convertir_columnas_excel(df):

    nuevas = []

    for i in range(len(df.columns)):
        nuevas.append(letras_excel(i))

    df.columns = nuevas

    return df


# =========================================================
# MAPEO MANUAL
# =========================================================

def obtener_mapeo_columnas(
    df,
    tiene_encabezado=True,
    incluir_z=False,
    incluir_desc=False,
):

    columnas = list(df.columns)

    detected = {}

    if tiene_encabezado:
        detected = detectar_columnas_automaticamente(df)

    st.markdown("### Configuración de columnas")

    col1, col2, col3 = st.columns(3)

    with col1:
        punto_col = st.selectbox(
            "PUNTO",
            columnas,
            index=(
                columnas.index(detected["PUNTO"])
                if "PUNTO" in detected
                else 0
            ),
        )

    with col2:
        este_col = st.selectbox(
            "ESTE / LAT",
            columnas,
            index=(
                columnas.index(detected["ESTE"])
                if "ESTE" in detected
                else 0
            ),
        )

    with col3:
        norte_col = st.selectbox(
            "NORTE / LON",
            columnas,
            index=(
                columnas.index(detected["NORTE"])
                if "NORTE" in detected
                else 0
            ),
        )

    z_col = None
    desc_col = None

    if incluir_z:

        z_col = st.selectbox(
            "Z / Elevación",
            columnas,
            index=(
                columnas.index(detected["Z"])
                if "Z" in detected
                else 0
            ),
        )

    if incluir_desc:

        desc_col = st.selectbox(
            "Descripción",
            columnas,
            index=(
                columnas.index(detected["DESCRIPCION"])
                if "DESCRIPCION" in detected
                else 0
            ),
        )

    return {
        "PUNTO": punto_col,
        "ESTE": este_col,
        "NORTE": norte_col,
        "Z": z_col,
        "DESCRIPCION": desc_col,
    }


# =========================================================
# VALIDAR COORDENADAS
# =========================================================

def validar_utm(este, norte):

    este = limpiar_numero(este)
    norte = limpiar_numero(norte)

    if np.isnan(este):
        return False

    if np.isnan(norte):
        return False

    if este < 100000 or este > 900000:
        return False

    if norte < 1000000 or norte > 10000000:
        return False

    return True


def validar_latlon(lat, lon):

    lat = limpiar_numero(lat)
    lon = limpiar_numero(lon)

    if np.isnan(lat):
        return False

    if np.isnan(lon):
        return False

    if lat < -90 or lat > 90:
        return False

    if lon < -180 or lon > 180:
        return False

    return True


# =========================================================
# VALIDAR DATUM Y ZONA
# =========================================================

def validar_datum(datum):

    valid = [
        "PSAD56",
        "WGS84",
    ]

    return datum in valid


def validar_zona(zona):
    return zona in VALID_ZONES


# =========================================================
# CREAR DATAFRAME NORMALIZADO
# =========================================================

def construir_dataframe_base(
    df_original,
    mapping,
    input_type="UTM",
):

    rows = []

    for idx, row in df_original.iterrows():

        try:

            punto = sanitizar_texto(
                row[mapping["PUNTO"]]
            )

            valor1 = row[mapping["ESTE"]]
            valor2 = row[mapping["NORTE"]]

            z = np.nan
            desc = ""

            if mapping.get("Z"):
                z = limpiar_numero(
                    row[mapping["Z"]]
                )

            if mapping.get("DESCRIPCION"):
                desc = sanitizar_texto(
                    row[mapping["DESCRIPCION"]]
                )

            if input_type == "UTM":

                rows.append(
                    {
                        "PUNTO": punto,
                        "ESTE": limpiar_numero(valor1),
                        "NORTE": limpiar_numero(valor2),
                        "LAT": np.nan,
                        "LON": np.nan,
                        "Z": z,
                        "DESCRIPCION": desc,
                        "ERROR": "",
                    }
                )

            else:

                rows.append(
                    {
                        "PUNTO": punto,
                        "ESTE": np.nan,
                        "NORTE": np.nan,
                        "LAT": limpiar_numero(valor1),
                        "LON": limpiar_numero(valor2),
                        "Z": z,
                        "DESCRIPCION": desc,
                        "ERROR": "",
                    }
                )

        except Exception as e:

            rows.append(
                {
                    "PUNTO": "",
                    "ESTE": np.nan,
                    "NORTE": np.nan,
                    "LAT": np.nan,
                    "LON": np.nan,
                    "Z": np.nan,
                    "DESCRIPCION": "",
                    "ERROR": str(e),
                }
            )

    return pd.DataFrame(rows)


# =========================================================
# PREVIEW
# =========================================================

def render_preview_dataframe(df, titulo="Preview"):

    st.markdown(f"### {titulo}")

    if df is None or df.empty:
        st.warning("No hay datos")
        return

    st.dataframe(
        df.head(MAX_PREVIEW_ROWS),
        use_container_width=True,
        height=400,
    )

    if len(df) > MAX_PREVIEW_ROWS:
        st.caption(
            f"Mostrando {MAX_PREVIEW_ROWS} de {len(df)} filas"
        )


# =========================================================
# ERRORES
# =========================================================

def render_errores(df):

    if df is None:
        return

    if "ERROR" not in df.columns:
        return

    errores = df[
        df["ERROR"].astype(str).str.strip() != ""
    ]

    if errores.empty:
        st.success("Sin errores detectados")
        return

    st.warning(
        f"{len(errores)} filas tuvieron errores"
    )

    with st.expander("Ver errores"):

        st.dataframe(
            errores,
            use_container_width=True,
        )

# =========================================================
# SP Topo-Convert V7 | PARTE 4
# Motor geoespacial y conversiones
# =========================================================


# =========================================================
# EPSG PERÚ
# =========================================================

EPSG_CONFIG = {

    "PSAD56": {
        "17S": 24877,
        "18S": 24878,
        "19S": 24879,
    },

    "WGS84": {
        "17S": 32717,
        "18S": 32718,
        "19S": 32719,
    },
}


# =========================================================
# OBTENER TRANSFORMADOR
# =========================================================

@st.cache_resource
def get_transformer(
    source_datum,
    source_zone,
    target_datum,
    target_zone,
):

    source_epsg = EPSG_CONFIG[source_datum][source_zone]
    target_epsg = EPSG_CONFIG[target_datum][target_zone]

    transformer = pyproj.Transformer.from_crs(
        f"EPSG:{source_epsg}",
        f"EPSG:{target_epsg}",
        always_xy=True,
    )

    return transformer


@st.cache_resource
def get_transformer_to_geo(
    source_datum,
    source_zone,
):

    source_epsg = EPSG_CONFIG[source_datum][source_zone]

    transformer = pyproj.Transformer.from_crs(
        f"EPSG:{source_epsg}",
        "EPSG:4326",
        always_xy=True,
    )

    return transformer


@st.cache_resource
def get_transformer_from_geo(
    target_datum,
    target_zone,
):

    target_epsg = EPSG_CONFIG[target_datum][target_zone]

    transformer = pyproj.Transformer.from_crs(
        "EPSG:4326",
        f"EPSG:{target_epsg}",
        always_xy=True,
    )

    return transformer


# =========================================================
# CONVERSIONES
# =========================================================

def convertir_utm_a_utm(
    este,
    norte,
    source_datum,
    source_zone,
    target_datum,
    target_zone,
):

    transformer = get_transformer(
        source_datum,
        source_zone,
        target_datum,
        target_zone,
    )

    x, y = transformer.transform(
        float(este),
        float(norte),
    )

    return x, y


def convertir_utm_a_geo(
    este,
    norte,
    source_datum,
    source_zone,
):

    transformer = get_transformer_to_geo(
        source_datum,
        source_zone,
    )

    lon, lat = transformer.transform(
        float(este),
        float(norte),
    )

    return lat, lon


def convertir_geo_a_utm(
    lat,
    lon,
    target_datum,
    target_zone,
):

    transformer = get_transformer_from_geo(
        target_datum,
        target_zone,
    )

    x, y = transformer.transform(
        float(lon),
        float(lat),
    )

    return x, y


# =========================================================
# FORMATEADORES
# =========================================================

def fmt_coord(value, decimals=3):

    try:
        if pd.isna(value):
            return ""
        return f"{float(value):,.{decimals}f}"

    except:
        return ""


def fmt_latlon(value, decimals=8):

    try:
        if pd.isna(value):
            return ""
        return f"{float(value):.{decimals}f}"

    except:
        return ""


# =========================================================
# PROCESAR FILA UTM
# =========================================================

def procesar_fila_utm_conversion(
    row,
    source_datum,
    source_zone,
    target_datum,
    target_zone,
):

    try:

        este = limpiar_numero(row["ESTE"])
        norte = limpiar_numero(row["NORTE"])

        if not validar_utm(este, norte):

            row["ERROR"] = "UTM inválido"
            return row

        nuevo_este, nuevo_norte = convertir_utm_a_utm(
            este,
            norte,
            source_datum,
            source_zone,
            target_datum,
            target_zone,
        )

        lat, lon = convertir_utm_a_geo(
            nuevo_este,
            nuevo_norte,
            target_datum,
            target_zone,
        )

        row["ESTE"] = nuevo_este
        row["NORTE"] = nuevo_norte
        row["LAT"] = lat
        row["LON"] = lon
        row["ERROR"] = ""

        return row

    except Exception as e:

        row["ERROR"] = str(e)
        return row


def procesar_fila_utm_ubicacion(
    row,
    datum,
    zone,
):

    try:

        este = limpiar_numero(row["ESTE"])
        norte = limpiar_numero(row["NORTE"])

        if not validar_utm(este, norte):

            row["ERROR"] = "UTM inválido"
            return row

        lat, lon = convertir_utm_a_geo(
            este,
            norte,
            datum,
            zone,
        )

        row["LAT"] = lat
        row["LON"] = lon
        row["ERROR"] = ""

        return row

    except Exception as e:

        row["ERROR"] = str(e)
        return row


# =========================================================
# PROCESAR FILA GEO
# =========================================================

def procesar_fila_geo_conversion(
    row,
    target_datum,
    target_zone,
):

    try:

        lat = limpiar_numero(row["LAT"])
        lon = limpiar_numero(row["LON"])

        if not validar_latlon(lat, lon):

            row["ERROR"] = "Lat/Lon inválidos"
            return row

        este, norte = convertir_geo_a_utm(
            lat,
            lon,
            target_datum,
            target_zone,
        )

        row["ESTE"] = este
        row["NORTE"] = norte
        row["ERROR"] = ""

        return row

    except Exception as e:

        row["ERROR"] = str(e)
        return row


def procesar_fila_geo_ubicacion(
    row,
):

    try:

        lat = limpiar_numero(row["LAT"])
        lon = limpiar_numero(row["LON"])

        if not validar_latlon(lat, lon):

            row["ERROR"] = "Lat/Lon inválidos"
            return row

        row["ERROR"] = ""

        return row

    except Exception as e:

        row["ERROR"] = str(e)
        return row


# =========================================================
# MOTOR PRINCIPAL
# =========================================================

def ejecutar_conversion_masiva(
    df,
    input_type,
    source_datum,
    source_zone,
    target_datum,
    target_zone,
):

    resultados = []

    for _, row in df.iterrows():

        row = row.copy()

        if input_type == "UTM":

            row = procesar_fila_utm_conversion(
                row,
                source_datum,
                source_zone,
                target_datum,
                target_zone,
            )

        else:

            row = procesar_fila_geo_conversion(
                row,
                target_datum,
                target_zone,
            )

        resultados.append(row)

    return pd.DataFrame(resultados)


def ejecutar_ubicacion_masiva(
    df,
    input_type,
    datum,
    zone,
):

    resultados = []

    for _, row in df.iterrows():

        row = row.copy()

        if input_type == "UTM":

            row = procesar_fila_utm_ubicacion(
                row,
                datum,
                zone,
            )

        else:

            row = procesar_fila_geo_ubicacion(
                row,
            )

        resultados.append(row)

    return pd.DataFrame(resultados)


# =========================================================
# ESTADÍSTICAS
# =========================================================

def obtener_estadisticas(df):

    if df is None or df.empty:

        return {
            "total": 0,
            "ok": 0,
            "errores": 0,
        }

    total = len(df)

    errores = len(
        df[
            df["ERROR"].astype(str).str.strip() != ""
        ]
    )

    ok = total - errores

    return {
        "total": total,
        "ok": ok,
        "errores": errores,
    }


def render_estadisticas(df):

    stats = obtener_estadisticas(df)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(
            "Total",
            stats["total"],
        )

    with col2:
        st.metric(
            "Correctos",
            stats["ok"],
        )

    with col3:
        st.metric(
            "Errores",
            stats["errores"],
        )


# =========================================================
# RESULTADO MANUAL
# =========================================================

def render_resultado_manual(df):

    if df is None or df.empty:
        return

    row = df.iloc[0]

    st.markdown(
        """
        <div class='card-res'>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:

        st.markdown("### Coordenadas")

        st.write(
            f"**PUNTO:** {row['PUNTO']}"
        )

        st.write(
            f"**ESTE:** {fmt_coord(row['ESTE'])}"
        )

        st.write(
            f"**NORTE:** {fmt_coord(row['NORTE'])}"
        )

        if not pd.isna(row["Z"]):

            st.write(
                f"**Z:** {fmt_coord(row['Z'])}"
            )

    with col2:

        st.markdown("### Geográficas")

        st.write(
            f"**LAT:** {fmt_latlon(row['LAT'])}"
        )

        st.write(
            f"**LON:** {fmt_latlon(row['LON'])}"
        )

        if row["DESCRIPCION"]:

            st.write(
                f"**DESC:** {row['DESCRIPCION']}"
            )

    st.markdown(
        "</div>",
        unsafe_allow_html=True,
    )


# =========================================================
# RESUMEN
# =========================================================

def render_resumen_proceso(df):

    stats = obtener_estadisticas(df)

    st.success(
        f"Procesados: {stats['ok']} | "
        f"Errores: {stats['errores']}"
    )

# =========================================================
# SP Topo-Convert V7 | PARTE 5
# Mapas, visualización y preview geográfico
# =========================================================


# =========================================================
# ICONOS MAPA
# =========================================================

MAP_COLORS = {
    "ok": "blue",
    "error": "red",
    "single": "green",
}


# =========================================================
# CENTRO MAPA
# =========================================================

def calcular_centro_mapa(df):

    if df is None or df.empty:
        return -9.189967, -75.015152

    validos = df[
        (df["LAT"].notna()) &
        (df["LON"].notna())
    ]

    if validos.empty:
        return -9.189967, -75.015152

    lat = validos["LAT"].mean()
    lon = validos["LON"].mean()

    return lat, lon


# =========================================================
# TOOLTIP
# =========================================================

def construir_tooltip(row):

    texto = f"""
    <b>PUNTO:</b> {row['PUNTO']}<br>
    <b>ESTE:</b> {fmt_coord(row['ESTE'])}<br>
    <b>NORTE:</b> {fmt_coord(row['NORTE'])}<br>
    <b>LAT:</b> {fmt_latlon(row['LAT'])}<br>
    <b>LON:</b> {fmt_latlon(row['LON'])}<br>
    """

    if not pd.isna(row["Z"]):

        texto += f"""
        <b>Z:</b> {fmt_coord(row['Z'])}<br>
        """

    if row["DESCRIPCION"]:

        texto += f"""
        <b>DESC:</b> {row['DESCRIPCION']}<br>
        """

    if row["ERROR"]:

        texto += f"""
        <b>ERROR:</b> {row['ERROR']}<br>
        """

    return texto


# =========================================================
# MAPA BASE
# =========================================================

def crear_mapa_base(df):

    lat, lon = calcular_centro_mapa(df)

    mapa = folium.Map(
        location=[lat, lon],
        zoom_start=15,
        control_scale=True,
        tiles="OpenStreetMap",
    )

    return mapa


# =========================================================
# AGREGAR PUNTO
# =========================================================

def agregar_marker(
    mapa,
    row,
    color="blue",
):

    if pd.isna(row["LAT"]):
        return

    if pd.isna(row["LON"]):
        return

    popup = construir_tooltip(row)

    folium.Marker(
        location=[
            row["LAT"],
            row["LON"],
        ],
        popup=popup,
        tooltip=row["PUNTO"],
        icon=folium.Icon(
            color=color,
            icon="map-marker",
            prefix="fa",
        ),
    ).add_to(mapa)


# =========================================================
# MAPA MANUAL
# =========================================================

def render_mapa_manual(df):

    if df is None or df.empty:
        return

    mapa = crear_mapa_base(df)

    row = df.iloc[0]

    color = (
        MAP_COLORS["single"]
        if row["ERROR"] == ""
        else MAP_COLORS["error"]
    )

    agregar_marker(
        mapa,
        row,
        color=color,
    )

    st.markdown("### Ubicación en mapa")

    folium_static(
        mapa,
        width=1400,
        height=500,
    )


# =========================================================
# MAPA MASIVO
# =========================================================

def render_mapa_masivo(df):

    if df is None or df.empty:
        return

    validos = df[
        (df["LAT"].notna()) &
        (df["LON"].notna())
    ]

    if validos.empty:

        st.warning(
            "No hay coordenadas válidas para mostrar"
        )

        return

    mapa = crear_mapa_base(validos)

    for _, row in validos.iterrows():

        color = (
            MAP_COLORS["ok"]
            if row["ERROR"] == ""
            else MAP_COLORS["error"]
        )

        agregar_marker(
            mapa,
            row,
            color=color,
        )

    st.markdown("### Mapa de puntos")

    folium_static(
        mapa,
        width=1400,
        height=600,
    )


# =========================================================
# PREVIEW RESULTADO
# =========================================================

def render_preview_resultado(df):

    if df is None or df.empty:
        return

    columnas_preview = [
        "PUNTO",
        "ESTE",
        "NORTE",
        "LAT",
        "LON",
        "Z",
        "DESCRIPCION",
        "ERROR",
    ]

    preview = df.copy()

    for col in ["ESTE", "NORTE", "Z"]:

        if col in preview.columns:

            preview[col] = preview[col].apply(
                lambda x: fmt_coord(x)
            )

    for col in ["LAT", "LON"]:

        if col in preview.columns:

            preview[col] = preview[col].apply(
                lambda x: fmt_latlon(x)
            )

    st.markdown("### Resultado procesado")

    st.dataframe(
        preview[columnas_preview],
        use_container_width=True,
        height=500,
    )


# =========================================================
# PANEL RESUMEN
# =========================================================

def render_panel_resumen(df):

    stats = obtener_estadisticas(df)

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            """
            <div class='card-res'>
            <h3>Total</h3>
            """,
            unsafe_allow_html=True,
        )

        st.metric(
            "",
            stats["total"],
        )

        st.markdown(
            "</div>",
            unsafe_allow_html=True,
        )

    with col2:

        st.markdown(
            """
            <div class='card-res'>
            <h3>Correctos</h3>
            """,
            unsafe_allow_html=True,
        )

        st.metric(
            "",
            stats["ok"],
        )

        st.markdown(
            "</div>",
            unsafe_allow_html=True,
        )

    with col3:

        st.markdown(
            """
            <div class='card-res'>
            <h3>Errores</h3>
            """,
            unsafe_allow_html=True,
        )

        st.metric(
            "",
            stats["errores"],
        )

        st.markdown(
            "</div>",
            unsafe_allow_html=True,
        )


# =========================================================
# TABS RESULTADOS
# =========================================================

def render_tabs_resultados(df, modo="manual"):

    if df is None or df.empty:
        return

    tab1, tab2, tab3 = st.tabs([
        "Preview",
        "Mapa",
        "Errores",
    ])

    with tab1:

        render_panel_resumen(df)

        render_preview_resultado(df)

    with tab2:

        if modo == "manual":
            render_mapa_manual(df)
        else:
            render_mapa_masivo(df)

    with tab3:

        render_errores(df)


# =========================================================
# FILTRAR VÁLIDOS
# =========================================================

def obtener_df_validos(df):

    if df is None or df.empty:
        return pd.DataFrame()

    return df[
        df["ERROR"].astype(str).str.strip() == ""
    ].copy()


# =========================================================
# FILTRAR ERRORES
# =========================================================

def obtener_df_errores(df):

    if df is None or df.empty:
        return pd.DataFrame()

    return df[
        df["ERROR"].astype(str).str.strip() != ""
    ].copy()


# =========================================================
# CONTADORES
# =========================================================

def total_validos(df):

    return len(
        obtener_df_validos(df)
    )


def total_errores(df):

    return len(
        obtener_df_errores(df)
    )


# =========================================================
# ALERTAS UX
# =========================================================

def render_alertas_proceso(df):

    ok = total_validos(df)
    err = total_errores(df)

    if ok > 0:

        st.success(
            f"Se procesaron correctamente {ok} puntos"
        )

    if err > 0:

        st.warning(
            f"{err} filas tuvieron errores"
        )

# =========================================================
# SP Topo-Convert V7 | PARTE 6
# Exportaciones CSV, Excel, KML y DXF
# =========================================================


# =========================================================
# DATAFRAME EXPORTABLE
# =========================================================

def preparar_dataframe_exportacion(df):

    if df is None or df.empty:
        return pd.DataFrame()

    export_df = obtener_df_validos(df)

    if export_df.empty:
        return pd.DataFrame()

    export_df = export_df.copy()

    columnas = [
        "PUNTO",
        "ESTE",
        "NORTE",
        "LAT",
        "LON",
        "Z",
        "DESCRIPCION",
    ]

    for col in columnas:

        if col not in export_df.columns:

            export_df[col] = ""

    return export_df[columnas]


# =========================================================
# CSV
# =========================================================

def generar_csv_bytes(df):

    export_df = preparar_dataframe_exportacion(df)

    if export_df.empty:
        return None

    csv_bytes = export_df.to_csv(
        index=False,
        encoding="utf-8-sig",
    ).encode("utf-8-sig")

    return csv_bytes


# =========================================================
# EXCEL
# =========================================================

def generar_excel_bytes(df):

    export_df = preparar_dataframe_exportacion(df)

    if export_df.empty:
        return None

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
    ) as writer:

        export_df.to_excel(
            writer,
            index=False,
            sheet_name="RESULTADOS",
        )

        workbook = writer.book
        worksheet = writer.sheets["RESULTADOS"]

        format_header = workbook.add_format({
            "bold": True,
            "bg_color": "#0A3D62",
            "font_color": "white",
            "border": 1,
        })

        for col_num, value in enumerate(export_df.columns.values):

            worksheet.write(
                0,
                col_num,
                value,
                format_header,
            )

            worksheet.set_column(
                col_num,
                col_num,
                20,
            )

    output.seek(0)

    return output.getvalue()


# =========================================================
# KML
# =========================================================

def generar_kml(df):

    export_df = preparar_dataframe_exportacion(df)

    if export_df.empty:
        return None

    kml = []

    kml.append(
        '<?xml version="1.0" encoding="UTF-8"?>'
    )

    kml.append(
        '<kml xmlns="http://www.opengis.net/kml/2.2">'
    )

    kml.append("<Document>")

    kml.append(
        f"<name>{APP_NAME}</name>"
    )

    for _, row in export_df.iterrows():

        lat = row["LAT"]
        lon = row["LON"]

        if pd.isna(lat):
            continue

        if pd.isna(lon):
            continue

        nombre = sanitizar_texto(
            row["PUNTO"]
        )

        descripcion = sanitizar_texto(
            row["DESCRIPCION"]
        )

        z = row["Z"]

        kml.append("<Placemark>")

        kml.append(
            f"<name>{nombre}</name>"
        )

        desc_text = f"""
        ESTE: {fmt_coord(row['ESTE'])}
        NORTE: {fmt_coord(row['NORTE'])}
        """

        if not pd.isna(z):

            desc_text += f"""
            Z: {fmt_coord(z)}
            """

        if descripcion:

            desc_text += f"""
            DESC: {descripcion}
            """

        kml.append(
            f"<description><![CDATA[{desc_text}]]></description>"
        )

        kml.append("<Point>")

        kml.append(
            f"<coordinates>{lon},{lat},0</coordinates>"
        )

        kml.append("</Point>")

        kml.append("</Placemark>")

    kml.append("</Document>")
    kml.append("</kml>")

    return "\n".join(kml).encode("utf-8")


# =========================================================
# DXF
# =========================================================

def generar_dxf(df):

    export_df = preparar_dataframe_exportacion(df)

    if export_df.empty:
        return None

    doc = ezdxf.new("R2010")

    msp = doc.modelspace()

    if "PUNTOS" not in doc.layers:
        doc.layers.new(
            name="PUNTOS",
            dxfattribs={"color": 1},
        )

    if "TEXTOS" not in doc.layers:
        doc.layers.new(
            name="TEXTOS",
            dxfattribs={"color": 5},
        )

    for _, row in export_df.iterrows():

        este = limpiar_numero(
            row["ESTE"]
        )

        norte = limpiar_numero(
            row["NORTE"]
        )

        z = limpiar_numero(
            row["Z"]
        )

        if pd.isna(z):
            z = 0

        nombre = sanitizar_texto(
            row["PUNTO"]
        )

        descripcion = sanitizar_texto(
            row["DESCRIPCION"]
        )

        texto = nombre

        if descripcion:

            texto += f" - {descripcion}"

        punto = (
            este,
            norte,
            z,
        )

        msp.add_point(
            punto,
            dxfattribs={
                "layer": "PUNTOS",
            },
        )

        msp.add_text(
            texto,
            dxfattribs={
                "height": 1.8,
                "layer": "TEXTOS",
            },
        ).set_placement(
            (
                este + 1,
                norte + 1,
                z,
            )
        )

    output = io.BytesIO()

    doc.write(output)

    output.seek(0)

    return output.getvalue()


# =========================================================
# VALIDAR EXPORTACIÓN
# =========================================================

def validar_exportacion(df):

    if df is None:
        return False

    if df.empty:
        return False

    validos = obtener_df_validos(df)

    if validos.empty:
        return False

    return True


# =========================================================
# GASTAR CRÉDITO
# =========================================================

def consumir_credito_exportacion():

    if st.session_state.plan != "FREE":
        return True

    creditos = int(
        st.session_state.credits_free
    )

    if creditos <= 0:

        mostrar_mensaje(
            "No tienes créditos disponibles",
            "warning",
        )

        return False

    nuevos = creditos - 1

    actualizar_usuario(
        st.session_state.user_id,
        {
            "creditos_disponibles": nuevos,
        },
    )

    st.session_state.credits_free = nuevos

    return True


# =========================================================
# BOTONES EXPORTACIÓN
# =========================================================

def render_exportaciones(df):

    if not validar_exportacion(df):

        st.warning(
            "No hay datos válidos para exportar"
        )

        return

    st.markdown("## Exportar resultados")

    col1, col2, col3, col4 = st.columns(4)

    nombre_base = "SP_Topo_Convert"

    # CSV
    with col1:

        csv_data = generar_csv_bytes(df)

        if csv_data is not None:

            st.download_button(
                label="📄 Descargar CSV",
                data=csv_data,
                file_name=generar_nombre_archivo(
                    nombre_base,
                    "csv",
                ),
                mime="text/csv",
                use_container_width=True,
                on_click=consumir_credito_exportacion,
            )

    # EXCEL
    with col2:

        excel_data = generar_excel_bytes(df)

        if excel_data is not None:

            st.download_button(
                label="📊 Descargar Excel",
                data=excel_data,
                file_name=generar_nombre_archivo(
                    nombre_base,
                    "xlsx",
                ),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
                on_click=consumir_credito_exportacion,
            )

    # KML
    with col3:

        kml_data = generar_kml(df)

        if kml_data is not None:

            st.download_button(
                label="🌍 Descargar KML",
                data=kml_data,
                file_name=generar_nombre_archivo(
                    nombre_base,
                    "kml",
                ),
                mime="application/vnd.google-earth.kml+xml",
                use_container_width=True,
                on_click=consumir_credito_exportacion,
            )

    # DXF
    with col4:

        dxf_data = generar_dxf(df)

        if dxf_data is not None:

            st.download_button(
                label="📐 Descargar DXF",
                data=dxf_data,
                file_name=generar_nombre_archivo(
                    nombre_base,
                    "dxf",
                ),
                mime="application/dxf",
                use_container_width=True,
                on_click=consumir_credito_exportacion,
            )


# =========================================================
# LOG EXPORTACIÓN
# =========================================================

def registrar_exportacion(
    tipo,
    cantidad,
):

    registrar_log(
        accion=f"EXPORT_{tipo}",
        nombre_archivo=tipo,
        puntos_ok=cantidad,
        errores_filas=0,
        tiempo_ejecucion=0,
    )


# =========================================================
# PANEL EXPORTACIÓN
# =========================================================

def render_panel_exportacion(df):

    st.markdown("---")

    st.markdown(
        """
        <div class='card-res'>
        <h3>Exportación profesional</h3>
        <p>
        Descarga tus resultados en formatos reales
        de trabajo topográfico.
        </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_exportaciones(df)

# =========================================================
# SP Topo-Convert V7 | PARTE 7
# Google Sheets, usuarios, créditos y licencias
# =========================================================


# =========================================================
# CONEXIÓN GOOGLE SHEETS
# =========================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def read_sheet(sheet_name):
    return load_sheet(sheet_name)


def update_sheet(sheet_name, df):
    return save_sheet(sheet_name, df)
# =========================================================
# OBTENER USUARIOS
# =========================================================

def get_users_df():
    return read_sheet("Usuarios")


def get_licenses_df():
    return read_sheet("Licencias")


def get_logs_df():
    return read_sheet("Logs")


# =========================================================
# BUSCAR USUARIO
# =========================================================

def get_user(user_id):

    df = get_users_df()

    if df.empty:
        return None

    result = df[
        df["user_id"].astype(str) == str(user_id)
    ]

    if result.empty:
        return None

    return result.iloc[0].to_dict()


# =========================================================
# CREAR USUARIO
# =========================================================

def create_user(user_id):

    df = get_users_df()

    nuevo = pd.DataFrame([
        {
            "user_id": user_id,
            "plan": "FREE",
            "tipo_licencia": "NINGUNA",
            "fecha_activacion": "",
            "fecha_expiracion": "",
            "creditos_disponibles": FREE_DAILY_CREDITS,
            "ultima_fecha_free": str(date.today()),
        }
    ])

    df = pd.concat(
        [df, nuevo],
        ignore_index=True,
    )

    update_sheet(
        "Usuarios",
        df,
    )

    return nuevo.iloc[0].to_dict()


# =========================================================
# ACTUALIZAR USUARIO
# =========================================================

def actualizar_usuario(
    user_id,
    updates,
):

    df = get_users_df()

    if df.empty:
        return False

    mask = (
        df["user_id"].astype(str)
        == str(user_id)
    )

    if not mask.any():
        return False

    for key, value in updates.items():

        if key in df.columns:

            df.loc[mask, key] = value

    update_sheet(
        "Usuarios",
        df,
    )

    return True


# =========================================================
# RESET CRÉDITOS FREE
# =========================================================

def reset_creditos_si_corresponde():

    user_id = st.session_state.user_id

    user = get_user(user_id)

    if user is None:
        return

    hoy = str(date.today())

    ultima = str(
        user.get(
            "ultima_fecha_free",
            ""
        )
    )

    if hoy != ultima:

        actualizar_usuario(
            user_id,
            {
                "creditos_disponibles": FREE_DAILY_CREDITS,
                "ultima_fecha_free": hoy,
            },
        )

        st.session_state.credits_free = (
            FREE_DAILY_CREDITS
        )


# =========================================================
# CARGAR SESSION USER
# =========================================================

def load_user_session():

    user_id = st.session_state.user_id

    user = get_user(user_id)

    if user is None:
        user = create_user(user_id)

    reset_creditos_si_corresponde()

    st.session_state.plan = (
        user.get("plan", "FREE")
    )

    st.session_state.tipo_licencia = user.get("tipo_licencia", "NINGUNA")
    st.session_state.fecha_expiracion = user.get("fecha_expiracion", "")
    st.session_state.ultima_fecha_free = user.get("ultima_fecha_free", "")

    st.session_state.credits_free = int(
        float(
            user.get(
                "creditos_disponibles",
                FREE_DAILY_CREDITS,
            )
        )
    )


# =========================================================
# VALIDAR LICENCIA
# =========================================================

def validar_licencia(codigo):

    licencias = get_licenses_df()

    if licencias.empty:
        return False, "Sin licencias"

    result = licencias[
        licencias["codigo"].astype(str)
        == str(codigo)
    ]

    if result.empty:
        return False, "Código inválido"

    row = result.iloc[0]

    estado = str(
        row.get("estado", "")
    ).upper()

    if estado == "USADO":
        return False, "Código ya usado"

    return True, row.to_dict()


# =========================================================
# DURACIÓN PLAN
# =========================================================

def get_plan_duration_days(plan):

    plan = str(plan).upper()

    mapping = {
        "DIARIO": 1,
        "SEMANAL": 7,
        "MENSUAL": 30,
        "ANUAL": 365,
        "ADMIN": 9999,
    }

    return mapping.get(plan, 0)


# =========================================================
# ACTIVAR LICENCIA
# =========================================================

def activar_licencia(codigo):

    valido, result = validar_licencia(codigo)

    if not valido:

        mostrar_mensaje(
            result,
            "error",
        )

        return False

    licencia = result

    tipo = licencia["plan_tipo"]

    dias = get_plan_duration_days(tipo)

    fecha_inicio = datetime.now()

    fecha_fin = (
        fecha_inicio + timedelta(days=dias)
    )

    plan = (
        "ADMIN"
        if tipo == "ADMIN"
        else "PRO"
    )

    actualizar_usuario(
        st.session_state.user_id,
        {
            "plan": plan,
            "tipo_licencia": tipo,
            "fecha_activacion": str(fecha_inicio),
            "fecha_expiracion": str(fecha_fin),
        },
    )

    licencias = get_licenses_df()

    mask = (
        licencias["codigo"].astype(str)
        == str(codigo)
    )

    licencias.loc[
        mask,
        "estado"
    ] = "USADO"

    licencias.loc[
        mask,
        "usado_por"
    ] = st.session_state.user_id

    licencias.loc[
        mask,
        "fecha_uso"
    ] = str(datetime.now())

    update_sheet(
        "Licencias",
        licencias,
    )

    st.session_state.plan = plan

    st.session_state.tipo_licencia = tipo
    st.session_state.fecha_expiracion = fecha_fin.strftime("%Y-%m-%d %H:%M:%S")

    mostrar_mensaje(
        f"Plan {tipo} activado",
        "success",
    )

    return True


# =========================================================
# VALIDAR EXPIRACIÓN
# =========================================================

def validar_expiracion_plan():

    if st.session_state.plan == "FREE":
        return

    user = get_user(
        st.session_state.user_id
    )

    if user is None:
        return

    fecha_exp = str(
        user.get(
            "fecha_expiracion",
            ""
        )
    )

    if not fecha_exp:
        return

    try:

        fecha_exp = pd.to_datetime(
            fecha_exp
        )

        if datetime.now() > fecha_exp:

            actualizar_usuario(
                st.session_state.user_id,
                {
                    "plan": "FREE",
                    "tipo_licencia": "NINGUNA",
                },
            )

            st.session_state.plan = "FREE"

            st.session_state.tipo_licencia = (
                "NINGUNA"
            )

    except:
        pass


# =========================================================
# CONSUMIR CRÉDITOS
# =========================================================

def consumir_creditos(cantidad=1):

    if st.session_state.plan != "FREE":
        return True

    disponibles = int(
        st.session_state.credits_free
    )

    if disponibles < cantidad:

        mostrar_mensaje(
            "Créditos insuficientes",
            "warning",
        )

        return False

    nuevos = disponibles - cantidad

    actualizar_usuario(
        st.session_state.user_id,
        {
            "creditos_disponibles": nuevos,
        },
    )

    st.session_state.credits_free = nuevos

    return True


# =========================================================
# VALIDAR LÍMITE MASIVO FREE
# =========================================================

def validar_limite_free(df):

    if st.session_state.plan != "FREE":
        return True

    if len(df) <= FREE_MAX_ROWS:
        return True

    st.warning(
        f"""
        Plan FREE permite máximo
        {FREE_MAX_ROWS} filas masivas.
        """
    )

    return False


# =========================================================
# LOGS
# =========================================================

def registrar_log(
    accion,
    nombre_archivo="",
    puntos_ok=0,
    errores_filas=0,
    tiempo_ejecucion=0,
):

    logs = get_logs_df()

    nuevo = pd.DataFrame([
        {
            "fecha": str(datetime.now()),
            "user_id": st.session_state.user_id,
            "licencia_activa": st.session_state.tipo_licencia,
            "accion": accion,
            "nombre_archivo": nombre_archivo,
            "puntos_ok": puntos_ok,
            "errores_filas": errores_filas,
            "tiempo_ejecucion": tiempo_ejecucion,
        }
    ])

    logs = pd.concat(
        [logs, nuevo],
        ignore_index=True,
    )

    update_sheet(
        "Logs",
        logs,
    )


# =========================================================
# PANEL PLAN
# =========================================================

def render_user_plan():

    st.markdown(
        """
        <div class='sidebar-card'>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### Tu plan")

    st.write(
        f"Plan: {st.session_state.plan}"
    )

    st.write(
        f"Licencia: {st.session_state.tipo_licencia}"
    )

    if st.session_state.plan == "FREE":

        st.write(
            f"""
            Créditos:
            {st.session_state.credits_free}
            """
        )

    st.markdown(
        "</div>",
        unsafe_allow_html=True,
    )

# =========================================================
# SP Topo-Convert V7 | PARTE 8
# Módulo Manual | Convertir y Ubicar
# =========================================================


# =========================================================
# FORM MANUAL
# =========================================================

def render_manual_form():

    st.markdown(
        """
        <div class='section-title'>
        Conversión y ubicación manual
        </div>
        """,
        unsafe_allow_html=True,
    )

    tab1, tab2 = st.tabs([
        "Convertir",
        "Ubicar",
    ])

    with tab1:
        render_manual_convertir()

    with tab2:
        render_manual_ubicar()


# =========================================================
# INPUT TYPE
# =========================================================

def render_input_type_manual(key="manual"):

    tipo = st.radio(
        "Tipo de coordenadas",
        [
            "UTM",
            "LAT/LON",
        ],
        horizontal=True,
        key=f"tipo_{key}",
    )

    return tipo


# =========================================================
# SELECTORES
# =========================================================

def render_selectores_conversion(key="manual"):

    col1, col2 = st.columns(2)

    with col1:

        source_datum = st.selectbox(
            "Datum origen",
            VALID_DATUMS,
            key=f"src_datum_{key}",
        )

        source_zone = st.selectbox(
            "Zona origen",
            VALID_ZONES,
            key=f"src_zone_{key}",
        )

    with col2:

        target_datum = st.selectbox(
            "Datum destino",
            VALID_DATUMS,
            index=1,
            key=f"tgt_datum_{key}",
        )

        target_zone = st.selectbox(
            "Zona destino",
            VALID_ZONES,
            index=1,
            key=f"tgt_zone_{key}",
        )

    return (
        source_datum,
        source_zone,
        target_datum,
        target_zone,
    )


def render_selectores_ubicacion(key="manual"):

    datum = st.selectbox(
        "Datum",
        VALID_DATUMS,
        key=f"datum_{key}",
    )

    zone = st.selectbox(
        "Zona",
        VALID_ZONES,
        index=1,
        key=f"zone_{key}",
    )

    return datum, zone


# =========================================================
# INPUTS MANUALES
# =========================================================

def render_inputs_utm_manual(prefix):
    # Crear columnas para que los campos queden alineados uno al lado del otro
    col1, col2, col3 = st.columns(3)
    
    with col1:
        este = st.number_input("ESTE", format="%.3f", key=f"{prefix}_este")
    with col2:
        norte = st.number_input("NORTE", format="%.3f", key=f"{prefix}_norte")
    with col3:
        elevacion = st.number_input("Z / Elevación", format="%.3f", key=f"{prefix}_z")
        
    descripcion = st.text_input("Descripción", key=f"{prefix}_desc")
    
    return {"ESTE": este, "NORTE": norte, "Z": elevacion, "DESCRIPCION": descripcion}


def render_inputs_geo_manual(prefix):
    col1, col2 = st.columns(2)

    with col1:
        punto = st.text_input("PUNTO", placeholder="P1", key=f"{prefix}_punto")
        lat = st.text_input("LATITUD", placeholder="-12.046374", key=f"{prefix}_lat")
        z = st.text_input("Z / Elevación", placeholder="Opcional", key=f"{prefix}_z")

    with col2:
        lon = st.text_input("LONGITUD", placeholder="-77.042793", key=f"{prefix}_lon")
        descripcion = st.text_input("Descripción", placeholder="Opcional", key=f"{prefix}_desc")

    return {
        "PUNTO": punto,
        "LAT": lat,
        "LON": lon,
        "Z": z,
        "DESCRIPCION": descripcion,
    }

# =========================================================
# DATAFRAME MANUAL
# =========================================================

def construir_df_manual(
    data,
    input_type="UTM",
):

    if input_type == "UTM":

        row = {
            "PUNTO": sanitizar_texto(
                data["PUNTO"]
            ),
            "ESTE": limpiar_numero(
                data["ESTE"]
            ),
            "NORTE": limpiar_numero(
                data["NORTE"]
            ),
            "LAT": np.nan,
            "LON": np.nan,
            "Z": limpiar_numero(
                data["Z"]
            ),
            "DESCRIPCION": sanitizar_texto(
                data["DESCRIPCION"]
            ),
            "ERROR": "",
        }

    else:

        row = {
            "PUNTO": sanitizar_texto(
                data["PUNTO"]
            ),
            "ESTE": np.nan,
            "NORTE": np.nan,
            "LAT": limpiar_numero(
                data["LAT"]
            ),
            "LON": limpiar_numero(
                data["LON"]
            ),
            "Z": limpiar_numero(
                data["Z"]
            ),
            "DESCRIPCION": sanitizar_texto(
                data["DESCRIPCION"]
            ),
            "ERROR": "",
        }

    return pd.DataFrame([row])


# =========================================================
# MANUAL CONVERTIR
# =========================================================

def render_manual_convertir():

    input_type = render_input_type_manual(
        "manual_convert"
    )

    (
        source_datum,
        source_zone,
        target_datum,
        target_zone,
    ) = render_selectores_conversion(
        "manual_convert"
    )

    st.markdown("---")

    if input_type == "UTM":

        data = render_inputs_utm_manual("manual_convert")
    else:
        data = render_inputs_geo_manual("manual_convert")

    ejecutar = st.button(
        "Convertir punto",
        use_container_width=True,
        type="primary",
    )

    if not ejecutar:
        return

    try:

        df = construir_df_manual(
            data,
            input_type,
        )

        if input_type == "UTM":

            resultado = ejecutar_conversion_masiva(
                df=df,
                input_type="UTM",
                source_datum=source_datum,
                source_zone=source_zone,
                target_datum=target_datum,
                target_zone=target_zone,
            )

        else:

            resultado = ejecutar_conversion_masiva(
                df=df,
                input_type="LAT/LON",
                source_datum="",
                source_zone="",
                target_datum=target_datum,
                target_zone=target_zone,
            )

        render_resultado_manual(
            resultado
        )

        render_mapa_manual(
            resultado
        )

        render_panel_exportacion(
            resultado
        )

        registrar_log(
            accion="MANUAL_CONVERSION",
            puntos_ok=1,
        )

    except Exception as e:

        st.error(str(e))


# =========================================================
# MANUAL UBICAR
# =========================================================

def render_manual_ubicar():

    input_type = render_input_type_manual(
        "manual_ubicar"
    )

    if input_type == "UTM":

        datum, zone = (
            render_selectores_ubicacion(
                "manual_ubicar"
            )
        )

    st.markdown("---")

    if input_type == "UTM":

        data = render_inputs_utm_manual("manual_ubicar")
    else:
        data = render_inputs_geo_manual("manual_ubicar")

    ejecutar = st.button(
        "Ubicar punto",
        use_container_width=True,
        type="primary",
    )

    if not ejecutar:
        return

    try:

        df = construir_df_manual(
            data,
            input_type,
        )

        if input_type == "UTM":

            resultado = ejecutar_ubicacion_masiva(
                df=df,
                input_type="UTM",
                datum=datum,
                zone=zone,
            )

        else:

            resultado = ejecutar_ubicacion_masiva(
                df=df,
                input_type="LAT/LON",
                datum="",
                zone="",
            )

        render_resultado_manual(
            resultado
        )

        render_mapa_manual(
            resultado
        )

        render_panel_exportacion(
            resultado
        )

        registrar_log(
            accion="MANUAL_UBICACION",
            puntos_ok=1,
        )

    except Exception as e:

        st.error(str(e))
# =========================================================
# SP Topo-Convert V7 | PARTE 9
# Módulo Masivo | Convertir y Ubicar
# =========================================================


# =========================================================
# CONFIG ARCHIVO
# =========================================================

def render_configuracion_archivo(prefix="masivo"):

    st.markdown(
        """
        <div class='section-title'>
            Configuración del archivo
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        tiene_encabezado = st.checkbox(
            "El archivo tiene encabezado",
            value=True,
            key=f"{prefix}_check_encabezado",
        )

    with col2:

        incluir_z = st.checkbox(
            "Contiene elevación (Z)",
            value=False,
            key=f"{prefix}_incluye_z",
        )

    with col3:

        incluir_desc = st.checkbox(
            "Contiene descripción",
            value=False,
            key=f"{prefix}_incluye_desc",
        )

    return tiene_encabezado, incluir_z, incluir_desc

# =========================================================
# SUBIR ARCHIVO
# =========================================================

def render_uploader_masivo():

    uploaded_file = st.file_uploader(
        "Sube tu archivo",
        type=VALID_FILE_EXTENSIONS,
        help=(
            "Formatos: CSV, XLSX, XLS y TXT"
        ),
    )

    return uploaded_file


# =========================================================
# LEER ARCHIVO MASIVO
# =========================================================

def procesar_archivo_subido(
    uploaded_file,
    tiene_encabezado,
):

    if uploaded_file is None:
        return None

    df = leer_archivo(uploaded_file)

    valido, mensaje = validar_dataframe(df)

    if not valido:

        st.error(mensaje)

        return None

    if not tiene_encabezado:

        df = convertir_columnas_excel(df)

    return df


# =========================================================
# INPUT TYPE MASIVO
# =========================================================

def render_input_type_masivo(key):

    tipo = st.radio(
        "Tipo de coordenadas",
        [
            "UTM",
            "LAT/LON",
        ],
        horizontal=True,
        key=key,
    )

    return tipo


# =========================================================
# SELECTORES MASIVOS
# =========================================================

def render_selectores_conversion_masivo():

    col1, col2 = st.columns(2)

    with col1:

        source_datum = st.selectbox(
            "Datum origen",
            VALID_DATUMS,
            key="masivo_src_datum",
        )

        source_zone = st.selectbox(
            "Zona origen",
            VALID_ZONES,
            index=1,
            key="masivo_src_zone",
        )

    with col2:

        target_datum = st.selectbox(
            "Datum destino",
            VALID_DATUMS,
            index=1,
            key="masivo_tgt_datum",
        )

        target_zone = st.selectbox(
            "Zona destino",
            VALID_ZONES,
            index=1,
            key="masivo_tgt_zone",
        )

    return (
        source_datum,
        source_zone,
        target_datum,
        target_zone,
    )


def render_selectores_ubicacion_masivo():

    col1, col2 = st.columns(2)

    with col1:

        datum = st.selectbox(
            "Datum",
            VALID_DATUMS,
            key="masivo_ub_datum",
        )

    with col2:

        zone = st.selectbox(
            "Zona",
            VALID_ZONES,
            index=1,
            key="masivo_ub_zone",
        )

    return datum, zone


# =========================================================
# PREVIEW ORIGINAL
# =========================================================

def render_preview_original(df):

    st.markdown("### Archivo original")

    st.dataframe(
        df.head(MAX_PREVIEW_ROWS),
        use_container_width=True,
        height=350,
    )


# =========================================================
# VALIDAR CRÉDITOS
# =========================================================

def validar_creditos_masivo(df):

    if st.session_state.plan != "FREE":
        return True

    filas = len(df)

    creditos = int(
        st.session_state.credits_free
    )

    if filas > creditos:

        st.warning(
            f"""
            Necesitas {filas} créditos.
            Créditos disponibles:
            {creditos}
            """
        )

        return False

    return True


# =========================================================
# CONSUMIR CRÉDITOS MASIVO
# =========================================================

def consumir_creditos_masivo(df):

    if st.session_state.plan != "FREE":
        return True

    filas = len(df)

    return consumir_creditos(filas)


# =========================================================
# MASIVO CONVERTIR
# =========================================================

def render_masivo_convertir():

    st.markdown(
        """
        <div class='section-title'>
        Conversión masiva
        </div>
        """,
        unsafe_allow_html=True,
    )

    input_type = render_input_type_masivo(
        "masivo_convert_tipo"
    )

    (
        source_datum,
        source_zone,
        target_datum,
        target_zone,
    ) = render_selectores_conversion_masivo()

    (
        tiene_encabezado,
        incluir_z,
        incluir_desc,
    ) = render_configuracion_archivo()

    uploaded_file = render_uploader_masivo()

    if uploaded_file is None:
        return

    df_original = procesar_archivo_subido(
        uploaded_file,
        tiene_encabezado,
    )

    if df_original is None:
        return

    if not validar_limite_free(df_original):
        return

    render_preview_original(df_original)

    mapping = obtener_mapeo_columnas(
        df_original,
        tiene_encabezado,
        incluir_z,
        incluir_desc,
    )

    ejecutar = st.button(
        "Convertir y previsualizar",
        type="primary",
        use_container_width=True,
    )

    if not ejecutar:
        return

    if not validar_creditos_masivo(
        df_original
    ):
        return

    try:

        with st.spinner(
            "Procesando coordenadas..."
        ):

            df_base = construir_dataframe_base(
                df_original=df_original,
                mapping=mapping,
                input_type=input_type,
            )

            resultado = ejecutar_conversion_masiva(
                df=df_base,
                input_type=input_type,
                source_datum=source_datum,
                source_zone=source_zone,
                target_datum=target_datum,
                target_zone=target_zone,
            )

            consumir_creditos_masivo(
                resultado
            )

            render_alertas_proceso(
                resultado
            )

            render_tabs_resultados(
                resultado,
                modo="masivo",
            )

            render_panel_exportacion(
                resultado
            )

            registrar_log(
                accion="MASIVO_CONVERSION",
                nombre_archivo=uploaded_file.name,
                puntos_ok=total_validos(
                    resultado
                ),
                errores_filas=total_errores(
                    resultado
                ),
            )

    except Exception as e:

        st.error(str(e))


# =========================================================
# MASIVO UBICAR
# =========================================================

def render_masivo_ubicar():

    st.markdown(
        """
        <div class='section-title'>
        Ubicación masiva
        </div>
        """,
        unsafe_allow_html=True,
    )

    input_type = render_input_type_masivo(
        "masivo_ubicar_tipo"
    )

    if input_type == "UTM":

        datum, zone = (
            render_selectores_ubicacion_masivo()
        )

    (
        tiene_encabezado,
        incluir_z,
        incluir_desc,
    ) = render_configuracion_archivo()

    uploaded_file = render_uploader_masivo()

    if uploaded_file is None:
        return

    df_original = procesar_archivo_subido(
        uploaded_file,
        tiene_encabezado,
    )

    if df_original is None:
        return

    if not validar_limite_free(df_original):
        return

    render_preview_original(df_original)

    mapping = obtener_mapeo_columnas(
        df_original,
        tiene_encabezado,
        incluir_z,
        incluir_desc,
    )

    ejecutar = st.button(
        "Ubicar y previsualizar",
        type="primary",
        use_container_width=True,
    )

    if not ejecutar:
        return

    if not validar_creditos_masivo(
        df_original
    ):
        return

    try:

        with st.spinner(
            "Ubicando puntos..."
        ):

            df_base = construir_dataframe_base(
                df_original=df_original,
                mapping=mapping,
                input_type=input_type,
            )

            resultado = ejecutar_ubicacion_masiva(
                df=df_base,
                input_type=input_type,
                datum=(
                    datum
                    if input_type == "UTM"
                    else ""
                ),
                zone=(
                    zone
                    if input_type == "UTM"
                    else ""
                ),
            )

            consumir_creditos_masivo(
                resultado
            )

            render_alertas_proceso(
                resultado
            )

            render_tabs_resultados(
                resultado,
                modo="masivo",
            )

            render_panel_exportacion(
                resultado
            )

            registrar_log(
                accion="MASIVO_UBICACION",
                nombre_archivo=uploaded_file.name,
                puntos_ok=total_validos(
                    resultado
                ),
                errores_filas=total_errores(
                    resultado
                ),
            )

    except Exception as e:

        st.error(str(e))


# =========================================================
# PANEL MASIVO
# =========================================================

def render_masivo_panel():

    tab1, tab2 = st.tabs([
        "Convertir",
        "Ubicar",
    ])

    with tab1:
        render_masivo_convertir()

    with tab2:
        render_masivo_ubicar()

# =========================================================
# SP Topo-Convert V7 | PARTE 10
# Admin Panel, Dashboard y métricas
# =========================================================


# =========================================================
# VALIDAR ADMIN
# =========================================================

def is_admin():

    return (
        st.session_state.plan == "ADMIN"
    )


# =========================================================
# MÉTRICAS GENERALES
# =========================================================

def obtener_metricas_generales():

    users = get_users_df()
    logs = get_logs_df()
    licencias = get_licenses_df()

    total_users = (
        0 if users.empty else len(users)
    )

    total_logs = (
        0 if logs.empty else len(logs)
    )

    total_licencias = (
        0 if licencias.empty else len(licencias)
    )

    licencias_usadas = 0

    if not licencias.empty:

        licencias_usadas = len(
            licencias[
                licencias["estado"]
                .astype(str)
                .str.upper() == "USADO"
            ]
        )

    return {
        "usuarios": total_users,
        "logs": total_logs,
        "licencias": total_licencias,
        "licencias_usadas": licencias_usadas,
    }


# =========================================================
# MÉTRICAS PLANES
# =========================================================

def obtener_metricas_planes():

    users = get_users_df()

    if users.empty:

        return {
            "FREE": 0,
            "PRO": 0,
            "ADMIN": 0,
        }

    planes = (
        users["plan"]
        .astype(str)
        .value_counts()
        .to_dict()
    )

    return {
        "FREE": planes.get("FREE", 0),
        "PRO": planes.get("PRO", 0),
        "ADMIN": planes.get("ADMIN", 0),
    }


# =========================================================
# MÉTRICAS ACCIONES
# =========================================================

def obtener_metricas_acciones():

    logs = get_logs_df()

    if logs.empty:
        return {}

    acciones = (
        logs["accion"]
        .astype(str)
        .value_counts()
        .to_dict()
    )

    return acciones


# =========================================================
# TOP ACCIONES
# =========================================================

def render_top_acciones():

    acciones = obtener_metricas_acciones()

    if not acciones:

        st.info(
            "Aún no existen registros"
        )

        return

    data = pd.DataFrame({
        "Acción": list(
            acciones.keys()
        ),
        "Cantidad": list(
            acciones.values()
        ),
    })

    st.markdown(
        "### Acciones más usadas"
    )

    st.dataframe(
        data,
        use_container_width=True,
        height=300,
    )


# =========================================================
# DASHBOARD RESUMEN
# =========================================================

def render_dashboard_cards():

    metricas = (
        obtener_metricas_generales()
    )

    planes = (
        obtener_metricas_planes()
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        st.metric(
            "Usuarios",
            metricas["usuarios"],
        )

    with col2:

        st.metric(
            "Logs",
            metricas["logs"],
        )

    with col3:

        st.metric(
            "Licencias",
            metricas["licencias"],
        )

    with col4:

        st.metric(
            "Usadas",
            metricas["licencias_usadas"],
        )

    st.markdown("---")

    col5, col6, col7 = st.columns(3)

    with col5:

        st.metric(
            "FREE",
            planes["FREE"],
        )

    with col6:

        st.metric(
            "PRO",
            planes["PRO"],
        )

    with col7:

        st.metric(
            "ADMIN",
            planes["ADMIN"],
        )


# =========================================================
# TABLA USUARIOS
# =========================================================

def render_admin_users():

    users = get_users_df()

    st.markdown(
        "### Usuarios registrados"
    )

    if users.empty:

        st.info(
            "No existen usuarios"
        )

        return

    st.dataframe(
        users,
        use_container_width=True,
        height=400,
    )


# =========================================================
# TABLA LICENCIAS
# =========================================================

def render_admin_licenses():

    licencias = get_licenses_df()

    st.markdown(
        "### Licencias"
    )

    if licencias.empty:

        st.info(
            "No existen licencias"
        )

        return

    st.dataframe(
        licencias,
        use_container_width=True,
        height=400,
    )


# =========================================================
# TABLA LOGS
# =========================================================

def render_admin_logs():

    logs = get_logs_df()

    st.markdown(
        "### Logs del sistema"
    )

    if logs.empty:

        st.info(
            "No existen logs"
        )

        return

    st.dataframe(
        logs.sort_values(
            by="fecha",
            ascending=False,
        ),
        use_container_width=True,
        height=450,
    )


# =========================================================
# CREAR LICENCIA
# =========================================================

def generar_codigo_licencia(
    prefijo="SP"
):

    random_str = uuid.uuid4().hex[:8]

    return (
        f"{prefijo}-"
        f"{random_str}"
    ).upper()


# =========================================================
# FORM LICENCIAS
# =========================================================

def render_admin_generate_license():

    st.markdown(
        "### Crear nueva licencia"
    )

    col1, col2 = st.columns(2)

    with col1:

        tipo = st.selectbox(
            "Tipo licencia",
            [
                "DIARIO",
                "SEMANAL",
                "MENSUAL",
                "ANUAL",
                "ADMIN",
            ],
        )

    with col2:

        cantidad = st.number_input(
            "Cantidad",
            min_value=1,
            max_value=100,
            value=1,
        )

    generar = st.button(
        "Generar licencias",
        use_container_width=True,
        type="primary",
    )

    if not generar:
        return

    licencias = get_licenses_df()

    nuevas = []

    for _ in range(cantidad):

        codigo = generar_codigo_licencia()

        nuevas.append({
            "codigo": codigo,
            "plan_tipo": tipo,
            "estado": "DISPONIBLE",
            "usado_por": "",
            "fecha_uso": "",
        })

    nuevas_df = pd.DataFrame(
        nuevas
    )

    licencias = pd.concat(
        [
            licencias,
            nuevas_df,
        ],
        ignore_index=True,
    )

    update_sheet(
        "Licencias",
        licencias,
    )

    st.success(
        f"""
        {cantidad} licencias creadas
        correctamente
        """
    )

    st.dataframe(
        nuevas_df,
        use_container_width=True,
    )


# =========================================================
# ADMIN PANEL
# =========================================================

def render_admin_panel():

    if not is_admin():
        return

    st.markdown("---")

    st.markdown(
        """
        <div class='section-title'>
        Panel Administrador
        </div>
        """,
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "Dashboard",
        "Usuarios",
        "Licencias",
        "Logs",
    ])

    with tab1:

        render_dashboard_cards()

        render_top_acciones()

        render_admin_generate_license()

    with tab2:

        render_admin_users()

    with tab3:

        render_admin_licenses()

    with tab4:

        render_admin_logs()

# =========================================================
# SP Topo-Convert V7 | PARTE 11
# Home, Sidebar, Licencias y App Principal
# =========================================================


# =========================================================
# HERO
# =========================================================

def render_hero():

    st.markdown(
        f"""
        <div class='hero-box'>

            <div class='hero-title'>
                {APP_NAME}
            </div>

            <div class='hero-subtitle'>
                Conversión, ubicación y exportación
                profesional de coordenadas topográficas
            </div>

        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# SIDEBAR
# =========================================================

def render_sidebar():

    with st.sidebar:

        if LOGO_PATH.exists():

            st.image(
                str(LOGO_PATH),
                use_container_width=True,
            )

        st.markdown("---")

        render_user_plan()

        st.markdown("---")

        render_license_box()

        st.markdown("---")

        render_pricing()

        st.markdown("---")

        render_free_info()


# =========================================================
# LICENCIAS
# =========================================================

def render_license_box():

    st.markdown(
        "### Activar licencia"
    )

    codigo = st.text_input(
        "Ingresa tu código",
        placeholder="SP-XXXXXX",
    )

    activar = st.button(
        "Activar plan",
        use_container_width=True,
    )

    if activar:

        if not codigo:

            st.warning(
                "Ingresa un código"
            )

        else:

            activar_licencia(codigo)


# =========================================================
# PLANES
# =========================================================

def render_pricing():

    st.markdown(
        "### Planes"
    )

    st.markdown(
        """
        <div class='pricing-card'>
            <h4>🟢 Gratis</h4>
            <p>5 créditos diarios</p>
            <p>Manual ilimitado</p>
            <p>Masivo hasta 20 filas</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class='pricing-card'>
            <h4>🔵 Diario</h4>
            <p>S/ 5</p>
            <p>Uso ilimitado</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class='pricing-card'>
            <h4>🟡 Semanal</h4>
            <p>S/ 10</p>
            <p>Uso ilimitado</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class='pricing-card'>
            <h4>🟠 Mensual</h4>
            <p>S/ 25</p>
            <p>Uso ilimitado</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class='pricing-card'>
            <h4>🔴 Anual</h4>
            <p>S/ 250</p>
            <p>Uso ilimitado</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# INFO FREE
# =========================================================

def render_free_info():

    st.markdown(
        "### Información FREE"
    )

    st.info(
        f"""
        • Manual ilimitado

        • 5 créditos diarios

        • Exportar consume 1 crédito

        • Masivo consume créditos
        por fila procesada

        • Máximo {FREE_MAX_ROWS}
        filas por archivo
        """
    )


# =========================================================
# HOME
# =========================================================

def render_home():

    render_hero()

    st.markdown("")

    tab1, tab2 = st.tabs([
        "Manual",
        "Masivo",
    ])

    with tab1:

        render_manual_form()

    with tab2:

        render_masivo_panel()

    render_admin_panel()


# =========================================================
# FOOTER
# =========================================================

def render_footer():

    st.markdown("---")

    st.markdown(
        f"""
        <div style='text-align:center;'>

        <strong>{APP_NAME}</strong>

        <br>

        Plataforma profesional para
        topografía y georreferenciación

        <br><br>

        © 2026 Todos los derechos reservados

        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# INIT SESSION
# =========================================================

def initialize_app():

    init_session_state()

    load_user_session()

    validar_expiracion_plan()


# =========================================================
# APP MAIN
# =========================================================

def main():

    initialize_app()

    render_sidebar()

    render_home()

    render_footer()


# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    main()

# =========================================================
# SP Topo-Convert V7 | PARTE 12
# Funciones faltantes críticas y correcciones finales
# =========================================================


# =========================================================
# CONVERTIR COLUMNAS EXCEL
# =========================================================

def convertir_columnas_excel(df):

    nuevas = []

    for i in range(len(df.columns)):

        nuevas.append(
            get_excel_column_name(i)
        )

    df.columns = nuevas

    return df


# =========================================================
# EXCEL COLUMN NAME
# =========================================================

def get_excel_column_name(n):

    result = ""

    while True:

        n, remainder = divmod(n, 26)

        result = (
            chr(65 + remainder)
            + result
        )

        if n == 0:
            break

        n -= 1

    return result


# =========================================================
# LIMPIAR DATAFRAME
# =========================================================

def limpiar_dataframe(df):

    if df is None:
        return pd.DataFrame()

    df = df.copy()

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    df = df.fillna("")

    return df


# =========================================================
# VALIDAR DATAFRAME
# =========================================================

def validar_dataframe(df):

    if df is None:
        return False, "Archivo inválido"

    if df.empty:
        return False, "Archivo vacío"

    if len(df.columns) < 2:

        return (
            False,
            "Archivo con pocas columnas",
        )

    return True, "OK"


# =========================================================
# LEER TXT
# =========================================================

def leer_txt(uploaded_file):

    try:

        df = pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python",
            encoding="utf-8",
        )

        return limpiar_dataframe(df)

    except:

        try:

            uploaded_file.seek(0)

            df = pd.read_csv(
                uploaded_file,
                sep=r"\s+",
                engine="python",
                encoding="latin1",
            )

            return limpiar_dataframe(df)

        except Exception as e:

            raise Exception(
                f"TXT inválido: {e}"
            )


# =========================================================
# LEER CSV
# =========================================================

def leer_csv(uploaded_file):

    try:

        df = pd.read_csv(
            uploaded_file,
            encoding="utf-8",
        )

        return limpiar_dataframe(df)

    except:

        try:

            uploaded_file.seek(0)

            df = pd.read_csv(
                uploaded_file,
                encoding="latin1",
            )

            return limpiar_dataframe(df)

        except Exception as e:

            raise Exception(
                f"CSV inválido: {e}"
            )


# =========================================================
# LEER EXCEL
# =========================================================

def leer_excel(uploaded_file):

    try:

        df = pd.read_excel(
            uploaded_file,
            engine="openpyxl",
        )

        return limpiar_dataframe(df)

    except Exception as e:

        raise Exception(
            f"Excel inválido: {e}"
        )


# =========================================================
# LEER ARCHIVO
# =========================================================

def leer_archivo(uploaded_file):

    extension = (
        Path(uploaded_file.name)
        .suffix
        .lower()
    )

    if extension == ".csv":

        return leer_csv(
            uploaded_file
        )

    elif extension in [
        ".xlsx",
        ".xls",
    ]:

        return leer_excel(
            uploaded_file
        )

    elif extension == ".txt":

        return leer_txt(
            uploaded_file
        )

    else:

        raise Exception(
            "Formato no soportado"
        )


# =========================================================
# DETECTAR COLUMNAS
# =========================================================

def detectar_columna(
    columnas,
    aliases,
):

    columnas_lower = [
        str(col).strip().lower()
        for col in columnas
    ]

    for alias in aliases:

        alias = alias.lower()

        for i, col in enumerate(columnas_lower):

            if alias == col:
                return columnas[i]

    for alias in aliases:

        alias = alias.lower()

        for i, col in enumerate(columnas_lower):

            if alias in col:
                return columnas[i]

    return None


# =========================================================
# MAPEADOR COLUMNAS
# =========================================================

def obtener_mapeo_columnas(
    df,
    tiene_encabezado=True,
    incluir_z=False,
    incluir_desc=False,
):

    columnas = list(df.columns)

    st.markdown("## Selección de columnas")

    punto_default = detectar_columna(
        columnas,
        COLUMN_ALIASES["PUNTO"],
    )

    este_default = detectar_columna(
        columnas,
        COLUMN_ALIASES["ESTE"],
    )

    norte_default = detectar_columna(
        columnas,
        COLUMN_ALIASES["NORTE"],
    )

    lat_default = detectar_columna(
        columnas,
        COLUMN_ALIASES["LAT"],
    )

    lon_default = detectar_columna(
        columnas,
        COLUMN_ALIASES["LON"],
    )

    z_default = detectar_columna(
        columnas,
        COLUMN_ALIASES["Z"],
    )

    desc_default = detectar_columna(
        columnas,
        COLUMN_ALIASES["DESCRIPCION"],
    )

    col1, col2 = st.columns(2)

    with col1:

        punto = st.selectbox(
            "Columna PUNTO",
            columnas,
            index=(
                columnas.index(punto_default)
                if punto_default in columnas
                else 0
            ),
        )

        este = st.selectbox(
            "Columna ESTE",
            columnas,
            index=(
                columnas.index(este_default)
                if este_default in columnas
                else 0
            ),
        )

        lat = st.selectbox(
            "Columna LAT",
            columnas,
            index=(
                columnas.index(lat_default)
                if lat_default in columnas
                else 0
            ),
        )

    with col2:

        norte = st.selectbox(
            "Columna NORTE",
            columnas,
            index=(
                columnas.index(norte_default)
                if norte_default in columnas
                else 0
            ),
        )

        lon = st.selectbox(
            "Columna LON",
            columnas,
            index=(
                columnas.index(lon_default)
                if lon_default in columnas
                else 0
            ),
        )

    z = None
    descripcion = None

    if incluir_z:

        z = st.selectbox(
            "Columna Z",
            columnas,
            index=(
                columnas.index(z_default)
                if z_default in columnas
                else 0
            ),
        )

    if incluir_desc:

        descripcion = st.selectbox(
            "Columna DESCRIPCIÓN",
            columnas,
            index=(
                columnas.index(desc_default)
                if desc_default in columnas
                else 0
            ),
        )

    return {
        "PUNTO": punto,
        "ESTE": este,
        "NORTE": norte,
        "LAT": lat,
        "LON": lon,
        "Z": z,
        "DESCRIPCION": descripcion,
    }


# =========================================================
# DATAFRAME BASE
# =========================================================

def construir_dataframe_base(
    df_original,
    mapping,
    input_type="UTM",
):

    filas = []

    for _, row in df_original.iterrows():

        try:

            nueva = {
                "PUNTO": sanitizar_texto(
                    row[
                        mapping["PUNTO"]
                    ]
                ),
                "ESTE": np.nan,
                "NORTE": np.nan,
                "LAT": np.nan,
                "LON": np.nan,
                "Z": np.nan,
                "DESCRIPCION": "",
                "ERROR": "",
            }

            if input_type == "UTM":

                nueva["ESTE"] = limpiar_numero(
                    row[
                        mapping["ESTE"]
                    ]
                )

                nueva["NORTE"] = limpiar_numero(
                    row[
                        mapping["NORTE"]
                    ]
                )

            else:

                nueva["LAT"] = limpiar_numero(
                    row[
                        mapping["LAT"]
                    ]
                )

                nueva["LON"] = limpiar_numero(
                    row[
                        mapping["LON"]
                    ]
                )

            if mapping["Z"]:

                nueva["Z"] = limpiar_numero(
                    row[
                        mapping["Z"]
                    ]
                )

            if mapping["DESCRIPCION"]:

                nueva["DESCRIPCION"] = (
                    sanitizar_texto(
                        row[
                            mapping[
                                "DESCRIPCION"
                            ]
                        ]
                    )
                )

            filas.append(nueva)

        except Exception as e:

            filas.append({
                "PUNTO": "",
                "ESTE": np.nan,
                "NORTE": np.nan,
                "LAT": np.nan,
                "LON": np.nan,
                "Z": np.nan,
                "DESCRIPCION": "",
                "ERROR": str(e),
            })

    return pd.DataFrame(filas)


# =========================================================
# RESULTADO MANUAL
# =========================================================

def render_resultado_manual(df):

    if df is None:
        return

    if df.empty:
        return

    row = df.iloc[0]

    st.markdown(
        """
        <div class='section-title'>
        Resultado
        </div>
        """,
        unsafe_allow_html=True,
    )

    if row["ERROR"]:

        st.error(
            row["ERROR"]
        )

        return

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "ESTE",
            fmt_coord(
                row["ESTE"]
            ),
        )

    with col2:

        st.metric(
            "NORTE",
            fmt_coord(
                row["NORTE"]
            ),
        )

    with col3:

        st.metric(
            "LAT / LON",
            (
                f"""
                {fmt_latlon(row['LAT'])}
                ,
                {fmt_latlon(row['LON'])}
                """
            ),
        )

    if not pd.isna(row["Z"]):

        st.info(
            f"""
            Elevación:
            {fmt_coord(row['Z'])}
            """
        )

    if row["DESCRIPCION"]:

        st.info(
            f"""
            Descripción:
            {row['DESCRIPCION']}
            """
        )

# =========================================================
# SP Topo-Convert V7 | PARTE 13
# Mapas, previews y resultados finales
# =========================================================


# =========================================================
# MAPA BASE
# =========================================================

def crear_mapa_base():

    mapa = folium.Map(
        location=PERU_CENTER,
        zoom_start=6,
        control_scale=True,
        tiles="OpenStreetMap",
    )

    folium.TileLayer(
        "CartoDB positron"
    ).add_to(mapa)

    folium.LayerControl().add_to(mapa)

    return mapa


# =========================================================
# POPUP HTML
# =========================================================

def generar_popup_html(row):

    html = f"""
    <div style='width:220px;'>

    <h4>{row['PUNTO']}</h4>

    <hr>

    <b>ESTE:</b>
    {fmt_coord(row['ESTE'])}

    <br>

    <b>NORTE:</b>
    {fmt_coord(row['NORTE'])}

    <br>

    <b>LAT:</b>
    {fmt_latlon(row['LAT'])}

    <br>

    <b>LON:</b>
    {fmt_latlon(row['LON'])}

    """

    if not pd.isna(row["Z"]):

        html += f"""
        <br>

        <b>Z:</b>
        {fmt_coord(row['Z'])}
        """

    if row["DESCRIPCION"]:

        html += f"""
        <br>

        <b>DESC:</b>
        {row['DESCRIPCION']}
        """

    html += "</div>"

    return html


# =========================================================
# AGREGAR PUNTO MAPA
# =========================================================

def agregar_punto_mapa(
    mapa,
    row,
):

    if pd.isna(row["LAT"]):
        return

    if pd.isna(row["LON"]):
        return

    popup = generar_popup_html(row)

    tooltip = str(
        row["PUNTO"]
    )

    folium.Marker(
        location=[
            row["LAT"],
            row["LON"],
        ],
        popup=popup,
        tooltip=tooltip,
        icon=folium.Icon(
            color="blue",
            icon="map-marker",
            prefix="fa",
        ),
    ).add_to(mapa)


# =========================================================
# AUTO FIT MAP
# =========================================================

def auto_fit_bounds(
    mapa,
    df,
):

    validos = obtener_df_validos(df)

    if validos.empty:
        return mapa

    bounds = []

    for _, row in validos.iterrows():

        if pd.isna(row["LAT"]):
            continue

        if pd.isna(row["LON"]):
            continue

        bounds.append([
            row["LAT"],
            row["LON"],
        ])

    if bounds:

        mapa.fit_bounds(bounds)

    return mapa


# =========================================================
# MAPA MANUAL
# =========================================================

def render_mapa_manual(df):

    if df is None:
        return

    validos = obtener_df_validos(df)

    if validos.empty:
        return

    st.markdown(
        """
        <div class='section-title'>
        Ubicación en mapa
        </div>
        """,
        unsafe_allow_html=True,
    )

    mapa = crear_mapa_base()

    for _, row in validos.iterrows():

        agregar_punto_mapa(
            mapa,
            row,
        )

    mapa = auto_fit_bounds(
        mapa,
        validos,
    )

    folium_static(mapa, width=1400, height=500)


# =========================================================
# MAPA MASIVO
# =========================================================

def render_mapa_masivo(df):

    if df is None:
        return

    validos = obtener_df_validos(df)

    if validos.empty:

        st.warning(
            "No existen puntos válidos"
        )

        return

    st.markdown(
        """
        <div class='section-title'>
        Vista geográfica masiva
        </div>
        """,
        unsafe_allow_html=True,
    )

    mapa = crear_mapa_base()

    for _, row in validos.iterrows():

        agregar_punto_mapa(
            mapa,
            row,
        )

    mapa = auto_fit_bounds(
        mapa,
        validos,
    )

    folium_static(mapa, width=1400, height=650)


# =========================================================
# PREVIEW RESULTADOS
# =========================================================

def render_preview_resultados(df):

    if df is None:
        return

    if df.empty:
        return

    st.markdown(
        """
        <div class='section-title'>
        Resultado procesado
        </div>
        """,
        unsafe_allow_html=True,
    )

    preview = df.copy()

    st.dataframe(
        preview,
        use_container_width=True,
        height=450,
    )


# =========================================================
# ALERTAS RESULTADO
# =========================================================

def render_alertas_proceso(df):

    validos = total_validos(df)

    errores = total_errores(df)

    col1, col2 = st.columns(2)

    with col1:

        st.success(
            f"""
            Procesados correctamente:
            {validos}
            """
        )

    with col2:

        if errores > 0:

            st.warning(
                f"""
                Filas con error:
                {errores}
                """
            )

        else:

            st.info(
                "Sin errores detectados"
            )


# =========================================================
# TAB ERRORES
# =========================================================

def render_tab_errores(df):

    errores = obtener_df_errores(df)

    if errores.empty:

        st.success(
            "No existen errores"
        )

        return

    st.dataframe(
        errores,
        use_container_width=True,
        height=300,
    )


# =========================================================
# TAB VÁLIDOS
# =========================================================

def render_tab_validos(df):

    validos = obtener_df_validos(df)

    if validos.empty:

        st.warning(
            "No existen resultados válidos"
        )

        return

    st.dataframe(
        validos,
        use_container_width=True,
        height=400,
    )


# =========================================================
# TABS RESULTADOS
# =========================================================

def render_tabs_resultados(
    df,
    modo="masivo",
):

    tab1, tab2, tab3 = st.tabs([
        "Resultados",
        "Mapa",
        "Errores",
    ])

    with tab1:

        render_preview_resultados(df)

    with tab2:

        if modo == "manual":

            render_mapa_manual(df)

        else:

            render_mapa_masivo(df)

    with tab3:

        render_tab_errores(df)


# =========================================================
# RESUMEN FINAL
# =========================================================

def render_resumen_final(df):

    validos = total_validos(df)

    errores = total_errores(df)

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:

        st.metric(
            "Correctos",
            validos,
        )

    with col2:

        st.metric(
            "Errores",
            errores,
        )


# =========================================================
# DESCARGA ERRORES CSV
# =========================================================

def render_descarga_errores(df):

    errores = obtener_df_errores(df)

    if errores.empty:
        return

    csv_bytes = errores.to_csv(
        index=False,
        encoding="utf-8-sig",
    ).encode("utf-8-sig")

    st.download_button(
        label="Descargar errores CSV",
        data=csv_bytes,
        file_name=generar_nombre_archivo(
            "errores",
            "csv",
        ),
        mime="text/csv",
        use_container_width=True,
    )

# =========================================================
# SP Topo-Convert V7 | PARTE 14
# Exportaciones CSV, Excel, KML y DXF
# =========================================================


# =========================================================
# CSV EXPORT
# =========================================================

def exportar_csv(df):

    return df.to_csv(
        index=False,
        encoding="utf-8-sig",
    ).encode("utf-8-sig")


# =========================================================
# EXCEL EXPORT
# =========================================================

def exportar_excel(df):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="RESULTADOS",
            index=False,
        )

        workbook = writer.book
        worksheet = (
            writer.sheets["RESULTADOS"]
        )

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#0F172A",
            "font_color": "white",
            "border": 1,
        })

        for col_num, value in enumerate(df.columns.values):

            worksheet.write(
                0,
                col_num,
                value,
                header_format,
            )

            worksheet.set_column(
                col_num,
                col_num,
                22,
            )

    output.seek(0)

    return output.getvalue()


# =========================================================
# KML POINT
# =========================================================

def generar_kml_point(
    row,
):

    punto = (
        row["PUNTO"]
        if row["PUNTO"]
        else "PUNTO"
    )

    descripcion = (
        row["DESCRIPCION"]
        if row["DESCRIPCION"]
        else ""
    )

    return f"""
    <Placemark>

        <name>{punto}</name>

        <description>
        {descripcion}
        </description>

        <Point>

            <coordinates>
            {row['LON']},
            {row['LAT']},
            0
            </coordinates>

        </Point>

    </Placemark>
    """


# =========================================================
# EXPORTAR KML
# =========================================================

def exportar_kml(df):

    validos = obtener_df_validos(df)

    kml_content = """
    <?xml version="1.0" encoding="UTF-8"?>

    <kml xmlns="http://www.opengis.net/kml/2.2">

    <Document>
    """

    for _, row in validos.iterrows():

        if pd.isna(row["LAT"]):
            continue

        if pd.isna(row["LON"]):
            continue

        kml_content += generar_kml_point(
            row
        )

    kml_content += """
    </Document>
    </kml>
    """

    return kml_content.encode("utf-8")


# =========================================================
# CREAR DXF
# =========================================================

def crear_dxf_document():

    doc = ezdxf.new()

    msp = doc.modelspace()

    return doc, msp


# =========================================================
# DXF POINT
# =========================================================

def agregar_punto_dxf(
    msp,
    row,
):

    if pd.isna(row["ESTE"]):
        return

    if pd.isna(row["NORTE"]):
        return

    x = float(row["ESTE"])
    y = float(row["NORTE"])

    punto = (
        row["PUNTO"]
        if row["PUNTO"]
        else "P"
    )

    msp.add_point(
        (
            x,
            y,
        )
    )

    msp.add_text(
        punto,
        dxfattribs={
            "height": 1.8,
        },
    ).set_placement(
        (
            x + 1,
            y + 1,
        )
    )


# =========================================================
# EXPORTAR DXF
# =========================================================

def exportar_dxf(df):

    validos = obtener_df_validos(df)

    doc, msp = crear_dxf_document()

    for _, row in validos.iterrows():

        agregar_punto_dxf(
            msp,
            row,
        )

    output = io.BytesIO()

    doc.write(output)

    output.seek(0)

    return output.getvalue()


# =========================================================
# NOMBRE ARCHIVO
# =========================================================

def generar_nombre_archivo(
    prefix="resultado",
    extension="csv",
):

    fecha = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    return (
        f"{prefix}_{fecha}.{extension}"
    )


# =========================================================
# VALIDAR EXPORTACIÓN
# =========================================================

def validar_exportacion():

    if st.session_state.plan != "FREE":
        return True

    if st.session_state.credits_free <= 0:

        st.warning(
            """
            No tienes créditos disponibles
            para exportar
            """
        )

        return False

    return True


# =========================================================
# CONSUMIR EXPORTACIÓN
# =========================================================

def consumir_credito_exportacion():

    if st.session_state.plan != "FREE":
        return True

    return consumir_creditos(1)


# =========================================================
# EXPORT BUTTON
# =========================================================

def render_export_button(
    label,
    data,
    file_name,
    mime,
    key,
):

    st.download_button(
        label=label,
        data=data,
        file_name=file_name,
        mime=mime,
        use_container_width=True,
        key=key,
    )


# =========================================================
# PANEL EXPORTACIÓN
# =========================================================

def render_panel_exportacion(df):

    if df is None:
        return

    if df.empty:
        return

    st.markdown(
        """
        <div class='section-title'>
        Exportaciones
        </div>
        """,
        unsafe_allow_html=True,
    )

    validos = obtener_df_validos(df)

    if validos.empty:

        st.warning(
            "No existen datos válidos"
        )

        return

    if not validar_exportacion():
        return

    col1, col2 = st.columns(2)

    with col1:

        csv_data = exportar_csv(
            validos
        )

        render_export_button(
            label="Descargar CSV",
            data=csv_data,
            file_name=generar_nombre_archivo(
                "coordenadas",
                "csv",
            ),
            mime="text/csv",
            key="csv_export",
        )

        excel_data = exportar_excel(
            validos
        )

        render_export_button(
            label="Descargar Excel",
            data=excel_data,
            file_name=generar_nombre_archivo(
                "coordenadas",
                "xlsx",
            ),
            mime=(
                "application/"
                "vnd.openxmlformats"
                "-officedocument."
                "spreadsheetml.sheet"
            ),
            key="excel_export",
        )

    with col2:

        kml_data = exportar_kml(
            validos
        )

        render_export_button(
            label="Descargar KML",
            data=kml_data,
            file_name=generar_nombre_archivo(
                "coordenadas",
                "kml",
            ),
            mime=(
                "application/"
                "vnd.google-earth.kml+xml"
            ),
            key="kml_export",
        )

        dxf_data = exportar_dxf(
            validos
        )

        render_export_button(
            label="Descargar DXF",
            data=dxf_data,
            file_name=generar_nombre_archivo(
                "coordenadas",
                "dxf",
            ),
            mime="application/dxf",
            key="dxf_export",
        )

    st.info(
        """
        Usuarios FREE:
        cada exportación consume 1 crédito
        """
    )

    render_descarga_errores(df)

# =========================================================
# SP Topo-Convert V7 | PARTE 15
# Seguridad, expiraciones, créditos y estabilidad
# =========================================================


# =========================================================
# FECHA ACTUAL
# =========================================================

def now_peru():

    return datetime.now()


# =========================================================
# PARSE FECHA
# =========================================================

def parse_datetime(value):

    if not value:
        return None

    try:

        return datetime.strptime(
            str(value),
            "%Y-%m-%d %H:%M:%S",
        )

    except:

        return None


# =========================================================
# VALIDAR EXPIRACIÓN
# =========================================================

def validar_expiracion_plan():

    if st.session_state.plan in [
        "FREE",
        "ADMIN",
    ]:
        return

    expiracion = parse_datetime(
        st.session_state.fecha_expiracion
    )

    if expiracion is None:
        return

    if now_peru() <= expiracion:
        return

    users = get_users_df()

    mask = (
        users["user_id"]
        == st.session_state.user_id
    )

    users.loc[
        mask,
        "plan"
    ] = "FREE"

    users.loc[
        mask,
        "tipo_licencia"
    ] = "NINGUNA"

    users.loc[
        mask,
        "fecha_activacion"
    ] = ""

    users.loc[
        mask,
        "fecha_expiracion"
    ] = ""

    users.loc[
        mask,
        "creditos_disponibles"
    ] = FREE_DAILY_CREDITS

    update_sheet(
        "Usuarios",
        users,
    )

    st.session_state.plan = "FREE"

    st.session_state.tipo_licencia = (
        "NINGUNA"
    )

    st.session_state.fecha_expiracion = ""

    st.success(
        """
        Tu licencia expiró.
        Volviste al plan FREE.
        """
    )


# =========================================================
# RESET FREE DIARIO
# =========================================================

def reset_free_daily_credits():

    if st.session_state.plan != "FREE":
        return

    users = get_users_df()

    if users.empty:
        return

    hoy = (
        datetime.now()
        .strftime("%Y-%m-%d")
    )

    ultima_fecha = str(
        st.session_state.ultima_fecha_free
    )

    if ultima_fecha == hoy:
        return

    mask = (
        users["user_id"]
        == st.session_state.user_id
    )

    users.loc[
        mask,
        "creditos_disponibles"
    ] = FREE_DAILY_CREDITS

    users.loc[
        mask,
        "ultima_fecha_free"
    ] = hoy

    update_sheet(
        "Usuarios",
        users,
    )

    st.session_state.credits_free = (
        FREE_DAILY_CREDITS
    )

    st.session_state.ultima_fecha_free = (
        hoy
    )


# =========================================================
# VALIDAR FREE FILAS
# =========================================================

def validar_limite_free(df):

    if st.session_state.plan != "FREE":
        return True

    if len(df) <= FREE_MAX_ROWS:
        return True

    st.warning(
        f"""
        El plan FREE solo permite
        hasta {FREE_MAX_ROWS} filas
        por archivo
        """
    )

    return False


# =========================================================
# VALIDAR CRÉDITOS
# =========================================================

def validar_creditos(
    cantidad=1,
):

    if st.session_state.plan != "FREE":
        return True

    disponibles = int(
        st.session_state.credits_free
    )

    if disponibles >= cantidad:
        return True

    st.error(
        """
        Créditos insuficientes
        """
    )

    return False


# =========================================================
# CONSUMIR CRÉDITOS
# =========================================================

def consumir_creditos(
    cantidad=1,
):

    if st.session_state.plan != "FREE":
        return True

    if not validar_creditos(
        cantidad
    ):
        return False

    users = get_users_df()

    mask = (
        users["user_id"]
        == st.session_state.user_id
    )

    actual = int(
        st.session_state.credits_free
    )

    nuevo = max(
        actual - cantidad,
        0,
    )

    users.loc[
        mask,
        "creditos_disponibles"
    ] = nuevo

    update_sheet(
        "Usuarios",
        users,
    )

    st.session_state.credits_free = (
        nuevo
    )

    return True


# =========================================================
# INFO SESIÓN
# =========================================================

def render_user_plan():

    plan = (
        st.session_state.plan
    )

    creditos = (
        st.session_state.credits_free
    )

    licencia = (
        st.session_state.tipo_licencia
    )

    st.markdown(
        "### Estado de cuenta"
    )

    if plan == "ADMIN":

        st.success(
            """
            ADMINISTRADOR
            Acceso total habilitado
            """
        )

        return

    if plan == "PRO":

        st.success(
            f"""
            PLAN PRO
            Licencia:
            {licencia}
            """
        )

        if st.session_state.fecha_expiracion:

            st.info(
                f"""
                Expira:
                {st.session_state.fecha_expiracion}
                """
            )

        return

    st.warning(
        f"""
        PLAN FREE

        Créditos:
        {creditos}
        """
    )


# =========================================================
# TIEMPO EJECUCIÓN
# =========================================================

def start_timer():

    st.session_state.process_start = (
        time.time()
    )


def end_timer():

    if (
        "process_start"
        not in st.session_state
    ):
        return 0

    return round(
        time.time()
        - st.session_state.process_start,
        2,
    )


# =========================================================
# LOG DETALLADO
# =========================================================

def registrar_log_detallado(
    accion="",
    nombre_archivo="",
    puntos_ok=0,
    errores_filas=0,
):

    tiempo = end_timer()

    registrar_log(
        accion=accion,
        nombre_archivo=nombre_archivo,
        puntos_ok=puntos_ok,
        errores_filas=errores_filas,
        tiempo_ejecucion=tiempo,
    )


# =========================================================
# ERROR UI
# =========================================================

def render_error_box(
    mensaje,
):

    st.markdown(
        f"""
        <div class='error-box'>
        {mensaje}
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# SUCCESS UI
# =========================================================

def render_success_box(
    mensaje,
):

    st.markdown(
        f"""
        <div class='success-box'>
        {mensaje}
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# LOADING UI
# =========================================================

@contextmanager
def process_loading(
    mensaje="Procesando..."
):

    with st.spinner(mensaje):

        start_timer()

        yield


# =========================================================
# PROTEGER APP
# =========================================================

def protect_app():

    reset_free_daily_credits()

    validar_expiracion_plan()


# =========================================================
# REEMPLAZAR initialize_app
# =========================================================

def initialize_app():

    init_session_state()

    load_user_session()

    protect_app()

# =========================================================
# SP Topo-Convert V7 | PARTE 16
# Mejoras UX, plantillas, helpers y estabilidad visual
# =========================================================


# =========================================================
# HELP CARD
# =========================================================

def render_help_card():

    st.markdown(
        """
        <div class='help-card'>

        <h3>¿Cómo usar la plataforma?</h3>

        <ul>

        <li>
        Manual:
        convierte o ubica un punto rápidamente
        </li>

        <li>
        Masivo:
        procesa múltiples puntos desde
        Excel o CSV
        </li>

        <li>
        Exporta resultados en:
        CSV, Excel, KML o DXF
        </li>

        <li>
        Visualiza tus puntos directamente
        en el mapa
        </li>

        </ul>

        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# PLANTILLA UTM
# =========================================================

def generar_template_utm():

    df = pd.DataFrame({
        "PUNTO": [
            "P1",
            "P2",
            "P3",
        ],
        "ESTE": [
            500000,
            500120,
            500250,
        ],
        "NORTE": [
            8500000,
            8500150,
            8500320,
        ],
        "Z": [
            120,
            122,
            125,
        ],
        "DESCRIPCION": [
            "PUNTO A",
            "PUNTO B",
            "PUNTO C",
        ],
    })

    return df


# =========================================================
# PLANTILLA GEO
# =========================================================

def generar_template_geo():

    df = pd.DataFrame({
        "PUNTO": [
            "P1",
            "P2",
            "P3",
        ],
        "LAT": [
            -12.046374,
            -12.046800,
            -12.047200,
        ],
        "LON": [
            -77.042793,
            -77.043200,
            -77.044000,
        ],
        "Z": [
            120,
            121,
            124,
        ],
        "DESCRIPCION": [
            "PUNTO A",
            "PUNTO B",
            "PUNTO C",
        ],
    })

    return df


# =========================================================
# DESCARGAR TEMPLATE
# =========================================================

def render_templates_download():

    st.markdown(
        """
        <div class='section-title'>
        Plantillas de ejemplo
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:

        utm_excel = exportar_excel(
            generar_template_utm()
        )

        st.download_button(
            label="Plantilla UTM",
            data=utm_excel,
            file_name="template_utm.xlsx",
            mime=(
                "application/"
                "vnd.openxmlformats"
                "-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    with col2:

        geo_excel = exportar_excel(
            generar_template_geo()
        )

        st.download_button(
            label="Plantilla LAT/LON",
            data=geo_excel,
            file_name="template_geo.xlsx",
            mime=(
                "application/"
                "vnd.openxmlformats"
                "-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )


# =========================================================
# ESTADÍSTICAS RÁPIDAS
# =========================================================

def render_stats_quick(df):

    if df is None:
        return

    validos = total_validos(df)

    errores = total_errores(df)

    total = len(df)

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Total",
            total,
        )

    with col2:

        st.metric(
            "Correctos",
            validos,
        )

    with col3:

        st.metric(
            "Errores",
            errores,
        )


# =========================================================
# MAPA SATELITAL
# =========================================================

def agregar_capa_satelital(
    mapa,
):

    folium.TileLayer(
        tiles=(
            "https://server.arcgisonline.com/"
            "ArcGIS/rest/services/"
            "World_Imagery/MapServer/"
            "tile/{z}/{y}/{x}"
        ),
        attr="Esri",
        name="Satélite",
        overlay=False,
        control=True,
    ).add_to(mapa)

    return mapa


# =========================================================
# REEMPLAZAR CREAR MAPA
# =========================================================

def crear_mapa_base():

    mapa = folium.Map(
        location=PERU_CENTER,
        zoom_start=6,
        control_scale=True,
        tiles="OpenStreetMap",
    )

    folium.TileLayer(
        "CartoDB positron",
        name="Claro",
    ).add_to(mapa)

    agregar_capa_satelital(
        mapa
    )

    folium.LayerControl(
        collapsed=False
    ).add_to(mapa)

    return mapa


# =========================================================
# BADGE PLAN
# =========================================================

def render_plan_badge():

    plan = (
        st.session_state.plan
    )

    if plan == "ADMIN":

        color = "#dc2626"

    elif plan == "PRO":

        color = "#2563eb"

    else:

        color = "#f59e0b"

    st.markdown(
        f"""
        <div style='
            background:{color};
            color:white;
            padding:10px;
            border-radius:12px;
            text-align:center;
            font-weight:700;
            margin-bottom:15px;
        '>

        PLAN {plan}

        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# REEMPLAZAR SIDEBAR
# =========================================================

def render_sidebar():

    with st.sidebar:

        logo = cargar_logo()
        if logo is not None:
            st.image(logo, use_container_width=True)

        render_plan_badge()

        render_user_plan()

        st.markdown("---")

        render_license_box()

        st.markdown("---")

        render_templates_download()

        st.markdown("---")

        render_pricing()

        st.markdown("---")

        render_free_info()

        st.markdown("---")

        render_help_card()


# =========================================================
# EMPTY RESULT
# =========================================================

def render_empty_result():

    st.info(
        """
        Aquí aparecerán los resultados
        procesados
        """
    )


# =========================================================
# REEMPLAZAR HOME
# =========================================================

def render_home():

    render_hero()

    st.markdown("")

    tab1, tab2 = st.tabs([
        "Manual",
        "Masivo",
    ])

    with tab1:

        render_manual_form()

    with tab2:

        render_masivo_panel()

    render_admin_panel()

    render_footer()


# =========================================================
# WARNING LICENSE
# =========================================================

def render_license_warning():

    if st.session_state.plan != "FREE":
        return

    creditos = int(
        st.session_state.credits_free
    )

    if creditos > 1:
        return

    st.warning(
        """
        Te queda 1 crédito disponible.
        Activa un plan PRO para
        uso ilimitado.
        """
    )


# =========================================================
# REEMPLAZAR MAIN
# =========================================================

def main():

    initialize_app()

    render_sidebar()

    render_license_warning()

    render_home()

# =========================================================
# SP Topo-Convert V7 | PARTE 17
# Correcciones integrales y mejoras de estabilidad
# =========================================================


# =========================================================
# REEMPLAZAR EXPORT BUTTON
# =========================================================

def render_export_button(
    label,
    data,
    file_name,
    mime,
    key,
):

    clicked = st.download_button(
        label=label,
        data=data,
        file_name=file_name,
        mime=mime,
        use_container_width=True,
        key=key,
    )

    if clicked:

        consumir_credito_exportacion()

        registrar_log_detallado(
            accion=f"EXPORT_{file_name}",
        )


# =========================================================
# VALIDAR COLUMNAS UTM
# =========================================================

def validar_columnas_utm(mapping):

    if not mapping["ESTE"]:

        st.error(
            "Debes seleccionar ESTE"
        )

        return False

    if not mapping["NORTE"]:

        st.error(
            "Debes seleccionar NORTE"
        )

        return False

    return True


# =========================================================
# VALIDAR COLUMNAS GEO
# =========================================================

def validar_columnas_geo(mapping):

    if not mapping["LAT"]:

        st.error(
            "Debes seleccionar LAT"
        )

        return False

    if not mapping["LON"]:

        st.error(
            "Debes seleccionar LON"
        )

        return False

    return True


# =========================================================
# VALIDAR MAPPING
# =========================================================

def validar_mapping(
    mapping,
    input_type="UTM",
):

    if not mapping["PUNTO"]:

        st.error(
            "Debes seleccionar PUNTO"
        )

        return False

    if input_type == "UTM":

        return validar_columnas_utm(
            mapping
        )

    return validar_columnas_geo(
        mapping
    )


# =========================================================
# VALIDAR FILA UTM
# =========================================================

def validar_fila_utm(
    este,
    norte,
):

    if pd.isna(este):
        return False

    if pd.isna(norte):
        return False

    if este <= 0:
        return False

    if norte <= 0:
        return False

    return True


# =========================================================
# VALIDAR FILA GEO
# =========================================================

def validar_fila_geo(
    lat,
    lon,
):

    if pd.isna(lat):
        return False

    if pd.isna(lon):
        return False

    if lat < -90 or lat > 90:
        return False

    if lon < -180 or lon > 180:
        return False

    return True


# =========================================================
# REEMPLAZAR CONSTRUIR BASE
# =========================================================

def construir_dataframe_base(
    df_original,
    mapping,
    input_type="UTM",
):

    filas = []

    for _, row in df_original.iterrows():

        try:

            nueva = {
                "PUNTO": "",
                "ESTE": np.nan,
                "NORTE": np.nan,
                "LAT": np.nan,
                "LON": np.nan,
                "Z": np.nan,
                "DESCRIPCION": "",
                "ERROR": "",
            }

            nueva["PUNTO"] = (
                sanitizar_texto(
                    row[
                        mapping["PUNTO"]
                    ]
                )
            )

            if input_type == "UTM":

                nueva["ESTE"] = (
                    limpiar_numero(
                        row[
                            mapping["ESTE"]
                        ]
                    )
                )

                nueva["NORTE"] = (
                    limpiar_numero(
                        row[
                            mapping["NORTE"]
                        ]
                    )
                )

                if not validar_fila_utm(
                    nueva["ESTE"],
                    nueva["NORTE"],
                ):

                    nueva["ERROR"] = (
                        "UTM inválido"
                    )

            else:

                nueva["LAT"] = (
                    limpiar_numero(
                        row[
                            mapping["LAT"]
                        ]
                    )
                )

                nueva["LON"] = (
                    limpiar_numero(
                        row[
                            mapping["LON"]
                        ]
                    )
                )

                if not validar_fila_geo(
                    nueva["LAT"],
                    nueva["LON"],
                ):

                    nueva["ERROR"] = (
                        "LAT/LON inválido"
                    )

            if mapping["Z"]:

                nueva["Z"] = (
                    limpiar_numero(
                        row[
                            mapping["Z"]
                        ]
                    )
                )

            if mapping["DESCRIPCION"]:

                nueva["DESCRIPCION"] = (
                    sanitizar_texto(
                        row[
                            mapping[
                                "DESCRIPCION"
                            ]
                        ]
                    )
                )

            filas.append(nueva)

        except Exception as e:

            filas.append({
                "PUNTO": "",
                "ESTE": np.nan,
                "NORTE": np.nan,
                "LAT": np.nan,
                "LON": np.nan,
                "Z": np.nan,
                "DESCRIPCION": "",
                "ERROR": str(e),
            })

    return pd.DataFrame(filas)


# =========================================================
# REEMPLAZAR MASIVO CONVERTIR
# =========================================================

def render_masivo_convertir():

    st.markdown(
        """
        <div class='section-title'>
        Conversión masiva
        </div>
        """,
        unsafe_allow_html=True,
    )

    input_type = render_input_type_masivo(
        "masivo_convert_tipo"
    )

    (
        source_datum,
        source_zone,
        target_datum,
        target_zone,
    ) = render_selectores_conversion_masivo()

    (
        tiene_encabezado,
        incluir_z,
        incluir_desc,
    ) = render_configuracion_archivo()

    uploaded_file = render_uploader_masivo()

    if uploaded_file is None:
        return

    df_original = procesar_archivo_subido(
        uploaded_file,
        tiene_encabezado,
    )

    if df_original is None:
        return

    if not validar_limite_free(
        df_original
    ):
        return

    render_preview_original(
        df_original
    )

    mapping = obtener_mapeo_columnas(
        df_original,
        tiene_encabezado,
        incluir_z,
        incluir_desc,
    )

    if not validar_mapping(
        mapping,
        input_type,
    ):
        return

    ejecutar = st.button(
        "Convertir y previsualizar",
        type="primary",
        use_container_width=True,
    )

    if not ejecutar:
        return

    if not validar_creditos_masivo(
        df_original
    ):
        return

    try:

        with process_loading(
            "Convirtiendo coordenadas..."
        ):

            df_base = construir_dataframe_base(
                df_original=df_original,
                mapping=mapping,
                input_type=input_type,
            )

            resultado = ejecutar_conversion_masiva(
                df=df_base,
                input_type=input_type,
                source_datum=source_datum,
                source_zone=source_zone,
                target_datum=target_datum,
                target_zone=target_zone,
            )

            consumir_creditos_masivo(
                resultado
            )

            render_stats_quick(
                resultado
            )

            render_alertas_proceso(
                resultado
            )

            render_tabs_resultados(
                resultado,
                modo="masivo",
            )

            render_panel_exportacion(
                resultado
            )

            registrar_log_detallado(
                accion="MASIVO_CONVERSION",
                nombre_archivo=uploaded_file.name,
                puntos_ok=total_validos(
                    resultado
                ),
                errores_filas=total_errores(
                    resultado
                ),
            )

    except Exception as e:

        render_error_box(
            str(e)
        )