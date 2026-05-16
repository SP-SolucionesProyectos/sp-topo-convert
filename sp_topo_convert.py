# =============================================================================
# SP TOPO-CONVERT | BLUEPRINT DE DIAGNÓSTICO Y CORE ESTRUCTURAL
# =============================================================================

import os
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pyproj
import folium
from streamlit_folium import folium_static

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN OBLIGATORIA DE LA PÁGINA (Primera línea ejecutable de ST)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="SP Topo-Convert | Panel de Diagnóstico",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título Corporativo en HTML Limpio
st.markdown("<h1 style='text-align: center; color: #0A3D62;'>SP Topo-Convert</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #7F8C8D;'>Monitoreo de Infraestructura, Conectividad y Formularios Geoespaciales</p>", unsafe_allow_html=True)
st.write("---")

# -----------------------------------------------------------------------------
# 2. MOTOR DE VERIFICACIÓN DE CREDENCIALES GOOGLE SHEETS (Módulo de Control)
# -----------------------------------------------------------------------------
st.subheader("🔍 Estado de Conectividad con la Base de Datos")

@st.cache_resource
def probar_conexion_gcp():
    """Valida la integridad criptográfica de la llave PEM y la conexión con gspread."""
    if "gcp_service_account" not in st.secrets:
        return False, "❌ Error: No se encontró la sección '[gcp_service_account]' en los Secrets de Streamlit."
    
    try:
        # Extraer credenciales desde los secrets del contenedor web
        creds_dict = dict(st.secrets["gcp_service_account"])
        
        # Corrector dinámico de padding y saltos de línea físicos para entornos Unix (Streamlit Cloud)
        if "private_key" in creds_dict and "\\n" in creds_dict["private_key"]:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        # Autenticación con Google Cloud API
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(credentials)
        
        # Intentar obtener el ID de la hoja configurada como prueba de fuego
        if "SHEET_ID" in st.secrets:
            sheet_id = st.secrets["SHEET_ID"]
            gc.open_by_key(sheet_id)
            return True, f"✅ Conexión Exitosa. Llave PEM procesada correctamente. Google Sheet (ID: {sheet_id[:8]}...) vinculada de forma óptima."
        
        return True, "✅ Conexión API Exitosa. (Nota: Variable 'SHEET_ID' global no definida en la raíz de los Secrets)."
        
    except Exception as e:
        error_msg = str(e)
        if "InvalidPadding" in error_msg or "PEM" in error_msg:
            return False, (
                "❌ Error Criptográfico (InvalidPadding): La llave PEM está mal estructurada.\n\n"
                "Asegúrate de haber ingresado la clave en el panel de Streamlit usando comillas triples "
                "y saltos de línea reales como se indicó previamente."
            )
        return False, f"❌ Error Inesperado al inicializar la Service Account: {error_msg}"

# Ejecutar el test e informar en pantalla
conexion_ok, mensaje_conexion = probar_conexion_gcp()
if conexion_ok:
    st.success(mensaje_conexion)
else:
    st.error(mensaje_conexion)

st.write("---")

# -----------------------------------------------------------------------------
# 3. FUNCIONES DE CAPTURA CON LLAVES ÚNICAS Y DISEÑO SIMÉTRICO (Solución Visual)
# -----------------------------------------------------------------------------

def render_inputs_utm_manual(prefix):
    """Renderiza campos UTM distribuidos horizontalmente en 3 columnas."""
    st.markdown("**Campos de Entrada UTM**")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        este = st.number_input("ESTE (X)", format="%.3f", key=f"{prefix}_utm_este_input", value=0.0)
    with col2:
        norte = st.number_input("NORTE (Y)", format="%.3f", key=f"{prefix}_utm_norte_input", value=0.0)
    with col3:
        elevacion = st.number_input("Z / Elevación", format="%.3f", key=f"{prefix}_utm_z_input", value=0.0)
        
    descripcion = st.text_input("Descripción del punto", key=f"{prefix}_utm_desc_input", placeholder="Ej: Vértice PR-01")
    
    return {"TIPO": "UTM", "ESTE": este, "NORTE": norte, "Z": elevacion, "DESCRIPCION": descripcion}


def render_inputs_geo_manual(prefix):
    """Renderiza campos Geográficos distribuidos horizontalmente en 3 columnas."""
    st.markdown("**Campos de Entrada Geográficos (WGS84)**")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        lat = st.text_input("Latitud", placeholder="Ej: -12.046374", key=f"{prefix}_geo_lat_input")
    with col2:
        lon = st.text_input("Longitud", placeholder="Ej: -77.042793", key=f"{prefix}_geo_lon_input")
    with col3:
        z = st.text_input("Z / Elevación (Opcional)", placeholder="Ej: 120.50", key=f"{prefix}_geo_z_input")
        
    descripcion = st.text_input("Descripción del punto", key=f"{prefix}_geo_desc_input", placeholder="Ej: Hito GPS")
    
    return {"TIPO": "GEO", "LATITUD": lat, "LONGITUD": lon, "Z": z, "DESCRIPCION": descripcion}

# -----------------------------------------------------------------------------
# 4. CONTROLADORES DE RENDERIZADO DE PESTAÑAS (Solución a TypeError y DuplicateKey)
# -----------------------------------------------------------------------------

def render_manual_convertir():
    """Pestaña de Conversión."""
    st.markdown("### 🔄 Módulo de Conversión Manual")
    
    # Selector de tipo de coordenada para control condicional interno
    input_type = st.radio(
        "Seleccione el tipo de coordenada de origen:",
        ["UTM", "LAT/LON"],
        key="convertir_selector_origen"
    )
    
    # Bifurcación estricta pasando el prefijo de ambiente "manual_convert"
    if input_type == "UTM":
        data = render_inputs_utm_manual("manual_convert")
    else:
        data = render_inputs_geo_manual("manual_convert")
        
    st.info(f"Datos listos para procesar en Conversión: {data}")


def render_manual_ubicar():
    """Pestaña de Ubicación."""
    st.markdown("### 📍 Módulo de Ubicación en Mapa")
    
    input_type = st.radio(
        "Seleccione el tipo de coordenada de origen:",
        ["UTM", "LAT/LON"],
        key="ubicar_selector_origen"
    )
    
    # Bifurcación estricta pasando el prefijo de ambiente "manual_ubicar"
    if input_type == "UTM":
        data = render_inputs_utm_manual("manual_ubicar")
    else:
        data = render_inputs_geo_manual("manual_ubicar")
        
    st.info(f"Datos listos para procesar en Ubicación: {data}")

# -----------------------------------------------------------------------------
# 5. ORQUESTADOR CENTRAL / MAIN
# -----------------------------------------------------------------------------

def main():
    st.subheader("🛠️ Panel de Pruebas de Interfaz")
    
    # Simulación del contenedor de pestañas principales (Manual / Masivo)
    tab_manual, tab_masivo = st.tabs(["Manual", "Masivo"])
    
    with tab_manual:
        st.write("---")
        # Sub-pestañas operativas de la sección manual
        subtab_convertir, subtab_ubicar = st.tabs(["Convertir", "Ubicar"])
        
        with subtab_convertir:
            render_manual_convertir()
            
        with subtab_ubicar:
            render_manual_ubicar()
            
    with tab_masivo:
        st.write("---")
        st.markdown("### 📦 Procesamiento Masivo")
        st.info("El sistema masivo se encuentra a la espera del estado de las pruebas manuales.")

if __name__ == "__main__":
    main()