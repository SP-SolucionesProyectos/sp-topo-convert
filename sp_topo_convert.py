# =========================================================
# BLOQUE 1: IMPORTACIONES Y CONFIGURACIÓN DE PÁGINA
# =========================================================
import streamlit as st
import pandas as pd
import pyproj
import folium
import os
import io
from datetime import datetime, timedelta
from streamlit_folium import folium_static
from streamlit_gsheets import GSheetsConnection
import ezdxf
import time
import string



# Configuración de pestaña (Debe ser lo primero siempre)
st.set_page_config(page_title="SP Topo-Convert", layout="wide", page_icon="🌍")

#1.1 INYECCIÓN DE ESTILOS PERSONALIZADOS (Versión Optimizada sin perder estilo)
st.markdown("""
    <style>
    /* Ajuste de espacio superior para que no se vea vacío arriba */
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    
    /* Colores corporativos de SP Soluciones y Proyectos */
    .stApp { background-color: #fffdf0; } /* Fondo crema claro */
    
    [data-testid="stSidebar"] { 
        background-color: #008080; 
        border-right: 3px solid #006666; 
    }

    /* Botones de navegación grandes y llamativos */
    .stButton > button {
        height: 3.5rem !important;
        width: 100% !important;
        font-size: 1.1rem !important;
        font-weight: bold !important;
        border-radius: 12px !important;
    }

    /* Tarjetas de resultados con sombra profesional */
    .res-card {
        background: white;
        padding: 20px;
        border-radius: 15px;
        border-left: 10px solid #008080;
        box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1);
        margin-bottom: 10px;
    }

    /* Eliminación de espacios innecesarios entre elementos */
    [data-testid="stVerticalBlock"] > div {
        padding-top: 0.1rem !important;
        padding-bottom: 0.1rem !important;
    }
    </style>
""", unsafe_allow_html=True)


# =========================================================
# BLOQUE 2: CONEXIÓN A BASE DE DATOS Y LOGS (CEREBRO)
# =========================================================
try:
    # Tu conexión establecida en .streamlit/secrets.toml
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"Error de conexión a la base de datos: {e}")

def registrar_actividad(tipo, detalle,zona="Desconocida"):
    """Guarda cada uso en tu hoja de Google Sheets para control administrativo"""
    try:
        # Leemos la hoja 'Logs' sin caché para tener datos frescos
        df_actual = conn.read(worksheet="Logs", ttl=0)
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        nueva_fila = pd.DataFrame([{
            "Fecha": ahora, 
            "Tipo_Uso": tipo, 
            "Ubicacion": zona, 
            "Detalle": detalle
        }])
        df_final = pd.concat([df_actual, nueva_fila], ignore_index=True)
        conn.update(worksheet="Logs", data=df_final)
    except:
        # Si falla el internet o el log, el programa sigue funcionando para el usuario
        pass

# =========================================================
# BLOQUE 3: GESTIÓN DE SESIÓN Y LÓGICA DE TIERS
# =========================================================
# Inicialización de variables de estado
if 'es_pro' not in st.session_state: st.session_state.es_pro = False
if 'es_admin' not in st.session_state: st.session_state.es_admin = False
if 'es_pase_diario' not in st.session_state: st.session_state.es_pase_diario = False
if 'menu_actual' not in st.session_state: st.session_state.menu_actual = "CONVERTIDOR"
if 'consultas' not in st.session_state: st.session_state.consultas = 0
if 'creditos_consumidos' not in st.session_state: st.session_state.creditos_consumidos = 0
if 'resultado' not in st.session_state: st.session_state.resultado = None
if 'descargas_kml' not in st.session_state: st.session_state.descargas_kml = 0
if 'df_temporal' not in st.session_state: st.session_state.df_temporal = None
if 'df_para_kml' not in st.session_state: st.session_state.df_para_kml = None

# Configuración de Límites
LIMITE_GRATIS_DIARIO = 10
LIMITE_FILAS_PASE_DIARIO = 500
LIMITE_KML_PASE_DIARIO = 100

# Reinicio de 24 horas
if 'inicio_sesion_gratis' not in st.session_state:
    st.session_state.inicio_sesion_gratis = datetime.now()

if (datetime.now() - st.session_state.inicio_sesion_gratis) > timedelta(days=1):
    st.session_state.consultas = 0
    st.session_state.inicio_sesion_gratis = datetime.now()
    st.toast("🎁 Tus créditos gratuitos se han renovado.")

# =========================================================
# BLOQUE 4: MOTOR DE CÁLCULO Y TRANSFORMACIONES CORE
# =========================================================
def validar_coordenadas_utm(e, n):
    """Verifica si las coordenadas están dentro de rangos técnicos razonables"""
    # Rango típico UTM: Este entre 100k-900k | Norte entre 0 y 10M
    if not (100000 <= e <= 900000):
        return False, f"Este ({e}) fuera de rango UTM."
    if not (0 <= n <= 10000000):
        return False, f"Norte ({n}) fuera de rango UTM."
    return True, ""

def realizar_conversion(e, n, zona_str, sentido):
    try:
        z_utm = zona_str.replace("S", "")
        
        # CASO A: SOLO UBICAR (No hay transformación Helmert)
        if "UBICAR" in sentido:
            # Determinamos el datum de entrada para el mapa
            datum_entrada = "epsg:327" if "WGS84" in sentido else "epsg:248" # Simplificado para el ejemplo
            # Usamos WGS84 (327 + zona) por defecto para el mapa de Google Earth
            geo_trans = pyproj.Transformer.from_crs(f"epsg:327{z_utm}", "epsg:4326", always_xy=True)
            lon, lat = geo_trans.transform(e, n)
            return e, n, lat, lon # Devuelve los mismos E, N originales

        # CASO B: CONVERSIÓN REAL (Tu código original de Helmert)
        if "PSAD56 a WGS84" in sentido:
            p = (f"+proj=pipeline +step +inv +proj=utm +zone={z_utm} +south +ellps=intl "
                 f"+step +proj=cart +ellps=intl "
                 f"+step +proj=helmert +x=-288 +y=175 +z=-376 "
                 f"+step +inv +proj=cart +ellps=WGS84 "
                 f"+step +proj=utm +zone={z_utm} +south +ellps=WGS84")
        else:
            p = (f"+proj=pipeline +step +inv +proj=utm +zone={z_utm} +south +ellps=WGS84 "
                 f"+step +proj=cart +ellps=WGS84 "
                 f"+step +inv +proj=helmert +x=-288 +y=175 +z=-376 "
                 f"+step +inv +proj=cart +ellps=intl "
                 f"+step +proj=utm +zone={z_utm} +south +ellps=intl")
        
        trans = pyproj.Transformer.from_pipeline(p)
        re, rn = trans.transform(e, n)
        geo_trans = pyproj.Transformer.from_crs(f"epsg:327{z_utm}", "epsg:4326", always_xy=True)
        # Si el destino es WGS84 usamos los resultados, si no, los originales para el mapa
        lon, lat = geo_trans.transform(re if "WGS84" in sentido else e, rn if "WGS84" in sentido else n)
        
        return re, rn, lat, lon
    except Exception as e:
        return None

# =========================================================
# BLOQUE 5: GENERACIÓN DE ARCHIVOS DE EXPORTACIÓN (KML Y PDF)
# =========================================================

def generar_kml(lat, lon, nombre="Punto_SP"):
    """Crea estructura KML estándar para un solo punto"""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Placemark>
    <name>{nombre}</name>
    <description>Convertido por SP Topo-Convert</description>
    <Point><coordinates>{lon},{lat},0</coordinates></Point>
  </Placemark>
</kml>"""

def generar_kml_masivo(df):
    """Crea un archivo KML optimizado para Google Earth (Sujeto al suelo)"""
    placemarks = ""
    for i, row in df.iterrows():
        p_lat, p_lon = row.get('LAT', 0), row.get('LON', 0)
        nombre_punto = row.get('PUNTO', f"P-{(i+1)}")
        
        # IMPORTANTE: El '0' al final de coordinates fuerza la elevación a cero.
        # <altitudeMode>clampToGround</altitudeMode> asegura que el punto toque el relieve.
        placemarks += f"""
  <Placemark>
    <name>{nombre_punto}</name>
    <Style>
      <IconStyle><color>ff00ffff</color><scale>1.1</scale></IconStyle>
    </Style>
    <Point>
      <altitudeMode>clampToGround</altitudeMode>
      <coordinates>{p_lon},{p_lat},0</coordinates>
    </Point>
  </Placemark>"""
    
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Levantamiento SP - Nuevo Chimbote</name>
    {placemarks}
  </Document>
</kml>"""

def exportar_cad(df, col_p, col_e, col_n, col_z=None, col_c=None):
    """Genera un archivo DXF con puntos 3D y etiquetas de ingeniería"""
    try:
        doc = ezdxf.new('R2010')
        doc.header['$PDMODE'] = 3 
        doc.header['$PDSIZE'] = 0.5
        
        msp = doc.modelspace()
        
        # Capas con colores estándar de topografía
        doc.layers.new(name='SP_PUNTOS', dxfattribs={'color': 1})     # Rojo
        doc.layers.new(name='SP_NRO_PUNTO', dxfattribs={'color': 2}) # Amarillo
        doc.layers.new(name='SP_ELEVACION', dxfattribs={'color': 4}) # Cian (estándar para Z)
        doc.layers.new(name='SP_CODIGO', dxfattribs={'color': 3})    # Verde

        for _, row in df.iterrows():
            try:
                def limpiar(v): return float(str(v).replace(',', '.'))
                
                p_id = str(row[col_p])
                e = limpiar(row[col_e])
                n = limpiar(row[col_n])
                # Elevación opcional (si no hay, se asume 0)
                z = limpiar(row[col_z]) if col_z and col_z in row else 0.0
                txt_cod = str(row[col_c]) if col_c and col_c in row else ""

                # 1. El Punto en 3D (X, Y, Z)
                msp.add_point((e, n, z), dxfattribs={'layer': 'SP_PUNTOS'})

                # 2. Etiquetas (Ajustadas para no solaparse)
                # Número de Punto (Arriba derecha)
                msp.add_text(p_id, dxfattribs={'layer': 'SP_NRO_PUNTO', 'height': 0.3}).set_placement((e + 0.4, n + 0.4))
                
                # Elevación (Centro derecha)
                if col_z:
                    msp.add_text(f"{z:.3f}", dxfattribs={'layer': 'SP_ELEVACION', 'height': 0.3}).set_placement((e + 0.4, n))

                # Código/Descripción (Abajo derecha)
                if txt_cod and txt_cod != "nan":
                    msp.add_text(txt_cod, dxfattribs={'layer': 'SP_CODIGO', 'height': 0.3}).set_placement((e + 0.4, n - 0.4))
            except:
                continue 

        out = io.StringIO()
        doc.write(out)
        return out.getvalue()
    except Exception:
        return None

# =========================================================
# BLOQUE 6: SISTEMA DE EXPORTACIÓN PROFESIONAL (KML & CAD) - ACCESO ABIERTO
# =========================================================
def mostrar_botones_exportacion(datos, es_masivo=False, col_z=None, col_c=None):

    """
    Gestiona la descarga de archivos con lógica de "Endulce":
    - Gratis: Descarga KML y DXF limitados a los primeros 10 puntos.
    - Pase Diario: KML y DXF limitados a 100 y 500 pts respectivamente.
    - Pro/Admin: Sin límites.
    """
    
    # 1. Definición de Permisos y Límites
    es_premium = (st.session_state.get('es_pro', False) or st.session_state.get('es_admin', False))
    es_pase = st.session_state.get('es_pase_diario', False)
    
    # Determinamos el límite de puntos según el plan para el "recorte"
    if es_premium:
        limite_puntos = 999999 
    elif es_pase:
        limite_puntos = LIMITE_FILAS_PASE_DIARIO # 500 puntos
    else:
        limite_puntos = LIMITE_GRATIS_DIARIO # 10 puntos

    # 2. Preparación de datos (Recorte)
    if es_masivo:
        datos_a_exportar = datos.head(limite_puntos)
        puntos_finales = len(datos_a_exportar)
        
        if es_premium:
            # Si es PRO o ADMIN, no mostramos mensaje de "recorte"
            pass
        elif es_pase:
            # Si es PASE DIARIO, solo avisamos si el archivo original era más grande que 500
            if len(datos) > LIMITE_FILAS_PASE_DIARIO:
                st.info(f"🎫 **Pase Diario:** Tu archivo excedió el límite y se procesaron los primeros {puntos_finales} puntos.")
        else:
            # Si es GRATUITO, siempre mostramos el mensaje informativo
            st.info(f"💡 **Modo Gratuito:** Tu archivo pudo procesar los primeros {puntos_finales} puntos.")
    else:
        datos_a_exportar = datos

    st.write("---")
    st.markdown("##### 📥 Generar Archivos de Ingeniería")
    
    col_kml, col_cad = st.columns(2)
    
    # --- SUB-BLOQUE: KML (Google Earth) ---
    with col_kml:
        try:
            if es_masivo:
                kml_str = generar_kml_masivo(datos_a_exportar)
                nombre_kml = f"SP_Masivo_{datetime.now().strftime('%d%m_%H%M')}.kml"
            else:
                kml_str = generar_kml(datos[2], datos[3]) 
                nombre_kml = f"SP_Punto_{datetime.now().strftime('%H%M')}.kml"
            
            st.download_button(
                label="🌍 Descargar KML",
                data=kml_str,
                file_name=nombre_kml,
                mime="application/vnd.google-earth.kml+xml",
                use_container_width=True,
                key=f"btn_kml_{'masivo' if es_masivo else 'ind'}"
            )
        except Exception as e:
            st.error(f"Error KML: {e}")

    # --- SUB-BLOQUE: DXF (AutoCAD) - AHORA ABIERTO ---
    with col_cad:
        try:
            if es_masivo:
                dxf_data = exportar_cad(
                    datos_a_exportar, 'PUNTO', 'ESTE', 'NORTE',
                    col_z='Z_CAD' if 'Z_CAD' in datos_a_exportar else None, 
                    col_c='DESC_CAD' if 'DESC_CAD' in datos_a_exportar else None
                )   
            else:
                df_ind = pd.DataFrame([{'PUNTO': 'P-01', 'ESTE': datos[0], 'NORTE': datos[1]}])
                dxf_data = exportar_cad(df_ind, 'PUNTO', 'ESTE', 'NORTE')

            if dxf_data:
                st.download_button(
                    label="📐 Descargar DXF (AutoCAD)",
                    data=dxf_data,
                    file_name=f"SP_Plano_{datetime.now().strftime('%d%m_%H%M')}.dxf",
                    mime="application/dxf",
                    use_container_width=True,
                    key=f"btn_dxf_{'masivo' if es_masivo else 'ind'}",
                    type="primary" if not es_premium else "secondary"
                )
        except Exception as e:
            st.error(f"Error CAD: {e}")

    # Mensaje de marketing sutil corregido
    if not es_premium:
        msg_limite = "500" if es_pase else "10"
        st.markdown(f"""
            <div style="text-align: center; margin-top: 10px;">
                <small>¿Necesitas procesar más de {LIMITE_GRATIS_DIARIO} puntos?
                <a href="#" onclick="window.location.reload();"> VER PLANES PRO</a></small>
            </div>
        """, unsafe_allow_html=True)

# =========================================================
# BLOQUE 7: SIDEBAR (IDENTIDAD CORPORATIVA)
# =========================================================
with st.sidebar:
    # 1. LOGO Y TÍTULO
    ruta_logo = "logo.png"
    
    if os.path.exists(ruta_logo):
        # El margin-top: -50px sube el logo al tope del sidebar
        st.markdown('<div style="margin-top: -50px;">', unsafe_allow_html=True)
        st.image(ruta_logo, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.error(f"No se encontró el logo.")
    
    st.markdown(f"""
        <div style="text-align: center; margin-top: -10px;">
            <p style="font-size: 0.85em; color: #facc15; opacity: 0.9;">Innovación digital a tu medida</p>
        </div>
        <hr style="border-color: rgba(250, 204, 21, 0.3); margin: 15px 0;">
    """, unsafe_allow_html=True)

    # 2. GUÍA RÁPIDA (En lugar de la publicidad repetida)
    # Esto ayuda a que el usuario no se pierda y da un aspecto de "Herramienta Profesional"
    st.markdown("""
        <div style="padding: 10px; border-radius: 10px; background-color: rgba(255,255,255,0.05); margin-bottom: 20px;">
        <p style="color: #facc15; font-weight: bold; font-size: 0.9em; margin-bottom: 5px;">📍 FLUJO DE TRABAJO</p>
        <ol style="color: #e2e8f0; font-size: 0.75em; padding-left: 15px;">
            <li><b>Configura:</b> Elige el sistema (PSAD56/WGS84) y tu zona UTM.</li>
            <li><b>Procesa:</b> Ingresa datos manuales o sube un archivo Excel/CSV.</li>
            <li><b>Visualiza:</b> Revisa la ubicación exacta en el mapa satelital.</li>
            <li><b>Ingeniería:</b> Exporta en Excel, KML y/o DXF para AutoCAD y Google Earth.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

    # 3. PUBLICIDAD ÚNICA Y MEJORADA
    st.markdown("""
        <div style="background: linear-gradient(145deg, #004d4d, #006666); 
                    padding: 15px; border-radius: 12px; border: 2px solid #facc15;
                    box-shadow: 0px 4px 15px rgba(0,0,0,0.3);">
            <p style="color: #facc15; font-size: 0.75em; font-weight: bold; margin-bottom: 5px; letter-spacing: 1px;">ANUNCIO</p>
            <p style="color: white; font-size: 0.95em; font-weight: bold; margin-bottom: 5px;">🚀 ¿Buscas una App o Web?</p>
            <p style="color: #cbd5e1; font-size: 0.75em; margin-bottom: 12px; line-height: 1.3;">
                Desarrollamos software profesional en <b>Nuevo Chimbote</b> para potenciar tu negocio.
            </p>
            <a href="https://wa.me/51924886915" target="_blank" 
               style="display: block; background-color: #facc15; color: #000; 
                      text-align: center; padding: 10px; border-radius: 8px; 
                      text-decoration: none; font-weight: bold; font-size: 0.85em;
                      transition: 0.3s;">
               💬 ¡COTIZA AQUÍ!
            </a>
        </div>
    """, unsafe_allow_html=True)

    # 4. FOOTER (Solo una vez)
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(f"""
        <div style="text-align: center; opacity: 0.6;">
            <p style="font-size: 0.7em; color: white;">v1.2.5 | © 2026 JGZ<br>Ingeniería de Sistemas - UTP</p>
        </div>
    """, unsafe_allow_html=True)

# =========================================================
# BLOQUE 8: VISTA PRINCIPAL - CONVERTIDOR (PARTE 1)
# =========================================================

# 1. NAVEGACIÓN SUPERIOR PROFESIONAL
_, b1, b2, b3 = st.columns([1.5, 1.2, 1.2, 1.2])

with b1:
    if st.button("🌍 CONVERTIDOR", use_container_width=True, type="primary" if st.session_state.menu_actual == "CONVERTIDOR" else "secondary"):
        st.session_state.menu_actual = "CONVERTIDOR"
        st.rerun()
with b2:
    if st.button("💎 PLAN PRO", use_container_width=True, type="primary" if st.session_state.menu_actual == "PRO" else "secondary"):
        st.session_state.menu_actual = "PRO"
        st.rerun()
with b3:
    if st.button("📞 SOPORTE", use_container_width=True, type="primary" if st.session_state.menu_actual == "CONTACTO" else "secondary"):
        st.session_state.menu_actual = "CONTACTO"
        st.rerun()
# --- NUEVO: BARRA DE ESTADO DE CRÉDITOS DINÁMICA ---
es_pro_admin = st.session_state.get('es_pro', False) or st.session_state.get('es_admin', False)
es_pase = st.session_state.get('es_pase_diario', False)

st.write("") # Espaciador sutil

if es_pro_admin:
    st.success("🚀 **Plan PRO Activo:** Tienes acceso ilimitado a todas las herramientas.")
elif es_pase:
    creditos_restantes = max(0, LIMITE_FILAS_PASE_DIARIO - st.session_state.consultas)
    st.info(f"🎫 **Pase Diario Activo:** Puedes convertir puntos de manera ilimitada (Manual) y masiva (Hasta {creditos_restantes} filas restantes).")
else:
    puntos_restantes = max(0, LIMITE_GRATIS_DIARIO - st.session_state.consultas)
    if puntos_restantes > 0:
        st.warning(f"💎 **Versión Gratuita:** Te quedan **{puntos_restantes}/{LIMITE_GRATIS_DIARIO}** créditos hoy.")
    else:
        st.error(f"❌ **Créditos agotados:** Has usado tus {LIMITE_GRATIS_DIARIO} créditos. Activa un pase para continuar hoy.")

# 2. LÓGICA DEL CONVERTIDOR
if st.session_state.menu_actual == "CONVERTIDOR":
    st.markdown('<h1 style="margin-bottom: -10px;">SP TOPO-CONVERT 🌍</h1>', unsafe_allow_html=True)
    st.markdown("""
        <p style="opacity: 0.9; margin-bottom: 20px; font-size: 1.1em; line-height: 1.4;">
            Esta plataforma ha sido diseñada para <b>convertir y ubicar puntos geográficos en todo el Perú</b> 
            con precisión profesional. Permite procesar datos individuales o masivos y exportarlos directamente 
            a formatos <b>KML (Google Earth)</b> e incluso <b>DXF para AutoCAD</b>.
        </p>
    """, unsafe_allow_html=True)

    # Configuración de modo y zona UTM
    opciones_modo = ["PSAD56 a WGS84 (Convertir)", "WGS84 a PSAD56 (Convertir)", "UBICAR PUNTO WGS84", "UBICAR PUNTO PSAD56"]
    modo = st.selectbox("Acción a realizar:", opciones_modo)
    zona_global = st.selectbox("Zona UTM:", ["17S", "18S", "19S"])
    st.session_state.modo_seleccionado = modo 

    # --- CONTROL DE CRÉDITOS Y BLOQUEOS ---
    # Verificamos si el usuario tiene algún plan activo
    es_limitado = not (st.session_state.get('es_pro', False) or 
                   st.session_state.get('es_admin', False) or 
                   st.session_state.get('es_pase_diario', False))
    puntos_restantes = max(0, LIMITE_GRATIS_DIARIO - st.session_state.consultas)
    bloqueo_total = es_limitado and puntos_restantes <= 0

    # Creación de Pestañas
    tab1, tab2 = st.tabs(["🎯 Individual", "📂 Masivo (Excel/CSV)"])

    # --- TAB 1: CONVERSIÓN ÚNICA ---
    with tab1:
        c1, c2 = st.columns([1, 1.2], gap="large")
    
        with c1:
            # Caso A: Usuario agotó sus créditos gratuitos
            if bloqueo_total:
                st.error(f"❌ Has agotado tus {LIMITE_GRATIS_DIARIO} créditos gratuitos de hoy.")
                st.info("Para seguir procesando coordenadas, activa el Plan PRO o un Pase Diario.")
                if st.button("🚀 ADQUIRIR ACCESO", key="btn_pro_tab1"):
                    st.session_state.menu_actual = "PRO"
                    st.rerun()
            
            # Caso B: Usuario con acceso (Gratis con puntos o Premium)
            else:
                st.subheader("Entrada de Datos")

                # Entradas numéricas (Persistentes)
                e_u = st.number_input("Coordenada ESTE (X):", value=None, placeholder="Ej: 771218.100", format="%.3f")
                n_u = st.number_input("Coordenada NORTE (Y):", value=None, placeholder="Ej: 8997417.000", format="%.3f")
                
                label_boton = "📍 UBICAR EN MAPA" if "UBICAR" in modo else "🔄 CONVERTIR COORDENADAS"
                
                if st.button(label_boton, use_container_width=True, type="primary"):
                    with st.spinner("Procesando..."):
                        
                        # Ejecución según el modo seleccionado
                        if "UBICAR" in modo:
                            z_utm = zona_global.replace("S", "")
                            epsg_entrada = "epsg:248" + z_utm if "PSAD56" in modo else "epsg:327" + z_utm
                            geo_trans = pyproj.Transformer.from_crs(epsg_entrada, "epsg:4326", always_xy=True)
                            lon, lat = geo_trans.transform(e_u, n_u)
                            # Guardamos en sesión: E_original, N_original, Lat, Lon
                            st.session_state.resultado = (e_u, n_u, lat, lon)
                        else:
                            res = realizar_conversion(e_u, n_u, zona_global, modo)
                            if res:
                                st.session_state.resultado = res
                        
                        # Consumo de crédito y registro
                        st.session_state.consultas += 1
                        registrar_actividad("Individual", f"{modo}", zona=zona_global)               
        with c2:
            if st.session_state.resultado:
                re, rn, lat, lon = st.session_state.resultado
                st.subheader("Resultado de Conversión")
                
                # Tarjetas de coordenadas con estilo SP
                st.markdown(f"""
                    <div class="res-card">
                        <div style="display: flex; justify-content: space-between;">
                            <div>
                                <small style="color: #666;">COORDENADA ESTE (X)</small><br>
                                <b style="font-size: 1.4em; color: #008080;">{re:,.3f}</b>
                            </div>
                            <div style="text-align: right;">
                                <small style="color: #666;">NORTE (Y)</small><br>
                                <b style="font-size: 1.4em; color: #008080;">{rn:,.3f}</b>
                            </div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                # Vista Geográfica (Lat/Lon)
                st.info(f"📍 **Geográficas (WGS84):** Lat: {lat:.8f} | Lon: {lon:.8f}")

                # Renderizado del Mapa
                with st.expander("🗺️ Ver Ubicación en Mapa de Calles", expanded=True):
                    try:
                        # Cambiamos a zoom 17 para ver mejor los nombres
                        m = folium.Map(location=[lat, lon], zoom_start=17, control_scale=True)

                        folium.Marker(
                            [lat, lon], 
                            popup=f"Punto SP\nE:{re:,.2f} N:{rn:,.2f}",
                            icon=folium.Icon(color='red', icon='info-sign')
                        ).add_to(m)
                        
                        folium_static(m, height=400) # Subí a 400 para que se vea más grande
                    except Exception as e:
                        st.error(f"Error al cargar el mapa: {e}")

                # --- INTEGRACIÓN DE DESCARGAS (BLOQUE 6) ---
                # Pasamos los datos como tupla para que el Bloque 6 genere el PDF individual
                mostrar_botones_exportacion(st.session_state.resultado, es_masivo=False)
                
                # Botón para limpiar resultado actual
                if st.button("🧹 Limpiar Resultado"):
                    st.session_state.resultado = None
                    st.rerun()
            else:
                # Estado de espera profesional
                st.markdown("""
                    <div style="text-align: center; padding: 50px; border: 2px dashed #ccc; border-radius: 15px; opacity: 0.5;">
                        <p style="font-size: 4em;">📍</p>
                        <p>Los resultados y el mapa aparecerán aquí <br> después de procesar tus coordenadas.</p>
                    </div>
                """, unsafe_allow_html=True)

    # --- TAB 2: PROCESAMIENTO MASIVO (MEJORADO) ---
    with tab2:
        if 'df_temporal' in st.session_state and st.session_state.df_temporal is not None:
            st.success(f"✅ Procesamiento completado: {len(st.session_state.df_temporal)} puntos.")
            
            # Vista previa con la columna vacía incluida
            with st.expander("📄 Ver Previsualización de Datos", expanded=True):
                st.dataframe(st.session_state.df_temporal.head(100), use_container_width=True)

            # Preparar descarga Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                st.session_state.df_temporal.to_excel(writer, index=False, sheet_name='Conversión_SP')
            
            st.download_button(
                label="📊 DESCARGAR EXCEL CON RESULTADOS", 
                data=output.getvalue(), 
                file_name=f"SP_Levantamiento_{datetime.now().strftime('%d%m_%H%M')}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True, type="primary"
            )
            st.write("---")

            # 2. DISTRIBUCIÓN: MAPA (Izquierda) | EXPORTACIÓN Y OTROS (Derecha)
            col_mapa, col_controles = st.columns([1.8, 1], gap="medium")

            with col_mapa:
                with st.expander("🗺️ Ver Vista Previa de Calles", expanded=True):
                    try:
                        df_mapa = st.session_state.df_para_kml.head(20)
                        if not df_mapa.empty:
                            centro_lat = df_mapa['LAT'].mean()
                            centro_lon = df_mapa['LON'].mean()
                            
                            m_masivo = folium.Map(location=[centro_lat, centro_lon], zoom_start=16)
                            
                            for _, r in df_mapa.iterrows():
                                # SOLUCIÓN: Si el nombre es nulo o NaN, le ponemos un nombre genérico
                                nombre_limpio = str(r['PUNTO']) if pd.notnull(r['PUNTO']) else "Sin Nombre"
                                # 1. Dibujamos un círculo pequeño y estético
                                folium.CircleMarker(
                                    location=[r['LAT'], r['LON']],
                                    radius=5,
                                    color='#1E88E5', # Azul profesional
                                    fill=True,
                                    fill_color='#1E88E5',
                                    fill_opacity=0.7,
                                    popup=f"Punto: {r['PUNTO']}"
                                ).add_to(m_masivo)

                                # 2. Agregamos el número del punto como texto flotante
                                folium.map.Marker(
                                    [r['LAT'], r['LON']],
                                    icon=folium.DivIcon(
                                        icon_size=(150,36),
                                        icon_anchor=(0,0),
                                        html=f'<div style="font-size: 10pt; color: #d35400; font-weight: bold; text-shadow: 1px 1px white;">{r["PUNTO"]}</div>',
                                    )
                                ).add_to(m_masivo)
                            
                            folium_static(m_masivo, height=400)
                            st.caption("📍 Vista previa: Puntos numerados para mayor precisión.")
                    except Exception as e:
                        st.warning(f"El mapa no pudo cargarse: {e}")


            with col_controles:
                # Botones de Exportación (KML/CAD) integrados aquí
                mostrar_botones_exportacion(st.session_state.df_para_kml, es_masivo=True)
                
                st.write("") # Espaciador
                # Botón Cargar Otro al final del costado
                if st.button("🗑️ Cargar otro archivo", use_container_width=True):
                    del st.session_state.df_temporal
                    st.rerun()
            
        elif bloqueo_total:
            st.warning("📍 **Límite de créditos alcanzado**")
            st.info("Activa el Plan PRO para procesar más archivos.")
            if st.button("🚀 VER PLANES", key="btn_bloqueo_masivo"):
                st.session_state.menu_actual = "PRO"; st.rerun()
        else:
            st.subheader("Carga de Levantamientos Masivos")
            file = st.file_uploader("Seleccionar archivo (Excel o CSV):", type=['xlsx', 'csv'])
            
            # --- CONTROL DE ENCABEZADOS ---
            tiene_header = st.checkbox("¿El archivo tiene fila de encabezados? (Nombres de columnas)", value=True)
            
            if file:
                try:
                    # Lectura inicial del archivo
                    if file.name.endswith('.csv'):
            # El motor 'python' con sep=None detecta si es coma o punto y coma automáticamente
                        df_to_proc = pd.read_csv(file, sep=None, engine='python', header=0 if tiene_header else None)
                    else:
                        df_to_proc = pd.read_excel(file, header=0 if tiene_header else None)

                    # --- LÓGICA DE LETRAS TIPO EXCEL (A, B, C...) ---
                    def index_to_letter(n):
                        res = ""
                        while n >= 0:
                            res = string.ascii_uppercase[n % 26] + res
                            n = n // 26 - 1
                        return res

                    # Definición de opciones para los selectores
                    if tiene_header:
                        opciones_cols = list(df_to_proc.columns)
                    else:
                        # AQUÍ EL CAMBIO: Generamos "Columna A", "Columna B", etc.
                        opciones_cols = [f"Columna {index_to_letter(i)}" for i in range(len(df_to_proc.columns))]
                    
                    # Diccionario para mapear la opción seleccionada al índice real (0, 1, 2...)
                    dict_indices = {op: i for i, op in enumerate(opciones_cols)}

                    # Función de ayuda para detección automática (se mantiene igual)
                    def auto_detect(lista, palabras_clave):
                        for p in palabras_clave:
                            for col in lista:
                                if p.lower() in str(col).lower():
                                    return col
                        return lista[0]

                    st.write("---")
                    st.markdown("##### ⚙️ Configuración del Levantamiento")
                    
                    c_opt1, c_opt2 = st.columns(2)
                    tiene_z = c_opt1.checkbox("¿El archivo tiene Elevación (Z)?", key="chk_z")
                    tiene_desc = c_opt2.checkbox("¿El archivo tiene Código/Descripción?", key="chk_desc")

                    # Selectores principales (Ahora con letras si no hay header)
                    c1, c2, c3 = st.columns(3)
                    sel_p = c1.selectbox("ID / Punto:", opciones_cols, index=opciones_cols.index(auto_detect(opciones_cols, ['PUNTO', 'ID', 'PNT', 'NOMBRE'])))
                    sel_e = c2.selectbox("Este (X):", opciones_cols, index=opciones_cols.index(auto_detect(opciones_cols, ['ESTE', 'X', 'EAST'])))
                    sel_n = c3.selectbox("Norte (Y):", opciones_cols, index=opciones_cols.index(auto_detect(opciones_cols, ['NORTE', 'Y', 'NORTH'])))

                    sel_z = None
                    sel_desc = None
                    if tiene_z or tiene_desc:
                        c4, c5 = st.columns(2)
                        if tiene_z:
                            sel_z = c4.selectbox("Elevación (Z):", opciones_cols, index=opciones_cols.index(auto_detect(opciones_cols, ['Z', 'ELEV', 'ALT', 'COTA'])))
                        if tiene_desc:
                            sel_desc = c5.selectbox("Descripción:", opciones_cols, index=opciones_cols.index(auto_detect(opciones_cols, ['DESC', 'COD', 'COMENT', 'TIPO'])))

                    # Mapeo final de columnas (Nombres o Índices reales)
                    # Usamos dict_indices para que el programa sepa que "Columna B" es el índice 1
                    col_p = sel_p if tiene_header else dict_indices[sel_p]
                    col_e = sel_e if tiene_header else dict_indices[sel_e]
                    col_n = sel_n if tiene_header else dict_indices[sel_n]
                    col_z = (sel_z if tiene_header else dict_indices[sel_z]) if sel_z else None
                    col_c = (sel_desc if tiene_header else dict_indices[sel_desc]) if sel_desc else None

                    # --- EL RESTO DEL BOTÓN DE CONVERSIÓN SIGUE IGUAL ---
                    # --- BOTÓN DE CONVERSIÓN MASIVA (PARTE 1) ---
                    if st.button("🚀 INICIAR CONVERSIÓN MASIVA", use_container_width=True, type="primary"):
                        # 1. SEGURIDAD Y RECORTE DE DATOS: Determinamos el límite REAL según el plan
                        if not (st.session_state.es_pro or st.session_state.es_admin or st.session_state.es_pase_diario):
                            puntos_permitidos = max(0, LIMITE_GRATIS_DIARIO - st.session_state.consultas)
                            if puntos_permitidos <= 0:
                                st.error("❌ Has agotado tus créditos gratuitos de hoy.")
                                st.stop()
                            
                            df_a_procesar = df_to_proc.head(puntos_permitidos).copy()
                            st.warning(f"⚠️ Modo Gratuito: Procesando solo los primeros {len(df_a_procesar)} puntos.")
                        else:
                            # Usuarios con Pase Diario o PRO procesan todo el archivo (o hasta su límite de pase)
                            if st.session_state.es_pase_diario:
                                puntos_restantes_pase = max(0, LIMITE_FILAS_PASE_DIARIO - st.session_state.consultas)
                                if puntos_restantes_pase <= 0:
                                    st.error("❌ Tu Pase Diario ha llegado al límite de 500 puntos.")
                                    st.stop()
                                df_a_procesar = df_to_proc.head(puntos_restantes_pase).copy()
                            else:
                                df_a_procesar = df_to_proc.copy()

                        # 2. INICIO DEL PROCESAMIENTO
                        with st.spinner("Procesando coordenadas..."):
                            res_list = []
                            errores = []
                            
                            # UN SOLO BUCLE: Limpio y eficiente
                            for idx, row in df_a_procesar.iterrows():
                                try:
                                    # Limpieza de datos
                                    raw_e = str(row[col_e]).strip() if pd.notnull(row[col_e]) else ""
                                    raw_n = str(row[col_n]).strip() if pd.notnull(row[col_n]) else ""
                                    
                                    if not raw_e or not raw_n or raw_e.lower() == 'nan':
                                        raise ValueError("Celda vacía o nula")

                                    val_e = float(raw_e.replace(',', '.'))
                                    val_n = float(raw_n.replace(',', '.'))
                                    
                                    # Validación técnica de rangos UTM
                                    es_valido, msg_error = validar_coordenadas_utm(val_e, val_n)
                                    
                                    if es_valido:
                                        conv = realizar_conversion(val_e, val_n, zona_global, modo)
                                        if conv:
                                            res_list.append(conv)
                                        else:
                                            res_list.append((0, 0, 0, 0))
                                            errores.append(f"Fila {idx+1}: Error en motor de cálculo.")
                                    else:
                                        res_list.append((0, 0, 0, 0))
                                        errores.append(f"Fila {idx+1}: {msg_error}")

                                except (ValueError, TypeError):
                                    res_list.append((0, 0, 0, 0))
                                    errores.append(f"Fila {idx+1}: Formato numérico inválido.")
                                except Exception as e:
                                    res_list.append((0, 0, 0, 0))
                                    errores.append(f"Fila {idx+1}: {str(e)}")
                            # 1. Mostrar reporte de errores si existen
                            if errores:
                                with st.expander("⚠️ Reporte de Calidad: Filas con observaciones"):
                                    for err in errores:
                                        st.write(f"• {err}")
                                    st.info("Nota: Las filas con error se marcaron con 0 para no detener el proceso.")

                            # 2. Creación del DataFrame de resultados
                            # res_list tiene exactamente el mismo largo que df_a_procesar
                            df_res = pd.DataFrame(res_list, columns=['ESTE_CONV', 'NORTE_CONV', 'LATITUD', 'LONGITUD'])
                            
                            # Concatenamos reseteando índices para evitar desplazamientos
                            df_final = pd.concat([df_a_procesar.reset_index(drop=True), df_res], axis=1)
                            
                            # 3. Guardar en session_state para persistencia
                            st.session_state.df_temporal = df_final
                            st.session_state.df_para_kml = df_final.copy()
                            
                            # Mapeo de columnas necesarias para el Bloque 6 (KML/CAD)
                            st.session_state.df_para_kml['PUNTO'] = df_final[col_p]
                            st.session_state.df_para_kml['LAT'] = df_final['LATITUD']
                            st.session_state.df_para_kml['LON'] = df_final['LONGITUD']
                            st.session_state.df_para_kml['ESTE'] = df_final['ESTE_CONV']
                            st.session_state.df_para_kml['NORTE'] = df_final['NORTE_CONV']
                            
                            if col_z: st.session_state.df_para_kml['Z_CAD'] = df_final[col_z]
                            if col_c: st.session_state.df_para_kml['DESC_CAD'] = df_final[col_c]

                            # 4. Actualización de Créditos y Sincronización GSheets
                            puntos_procesados = len(df_a_procesar)

                            if st.session_state.get('es_pase_diario') and 'codigo_activo' in st.session_state:
                                try:
                                    df_db = conn.read(worksheet="Usuarios", ttl=0)
                                    cod_actual = st.session_state.codigo_activo
                                    fila_mask = df_db['Codigo'].astype(str) == str(cod_actual)
                                    
                                    if fila_mask.any():
                                        # 1. Obtener lo que ya había consumido en la nube
                                        val_previo = pd.to_numeric(df_db.loc[fila_mask, 'Creditos_Usados'], errors='coerce').fillna(0).iloc[0]
                                        nuevo_total = int(val_previo + puntos_procesados)
                                        
                                        # 2. Actualizar la base de datos (GSheets)
                                        df_db.loc[fila_mask, 'Creditos_Usados'] = nuevo_total
                                        conn.update(worksheet="Usuarios", data=df_db)
                                        
                                        # 3. Sincronizar el estado local para que la barra de UI se actualice
                                        st.session_state.consultas = nuevo_total
                                except Exception as e:
                                    st.warning(f"Aviso: Se procesó localmente pero falló la sincronización con la nube: {e}")
                            else:
                                st.session_state.consultas += puntos_procesados

                            registrar_actividad("Masivo", f"Archivo de {puntos_procesados} puntos", zona=zona_global)
                            # 5. Finalización con éxito
                            st.balloons()
                            st.success(f"✅ ¡Conversión completada! Se procesaron {puntos_procesados} puntos.")
                            time.sleep(1.5) 


                except Exception as e:
                    st.error(f"Error al leer el archivo: {e}")

# =========================================================
# BLOQUE 9: VISTA PROFESIONAL - PLANES Y ACTIVACIÓN
# =========================================================
if st.session_state.menu_actual == "PRO":
    st.title("💎 Membresía SP Profesional")
    st.markdown("##### Potencia tu flujo de trabajo con acceso ilimitado")

    # --- Lógica dinámica de estados de suscripción ---
    if st.session_state.get('es_admin'):
        st.success("👑 **MODO MAESTRO ACTIVO:** Acceso Total JGZ - Disfruta del control total.")
    elif st.session_state.get('es_pro'):
        st.success("💎 **PLAN PRO ACTIVO:** Tienes acceso ilimitado a todas las herramientas.")
    elif st.session_state.get('es_pase_diario'):
        puntos_restantes = 500 - st.session_state.consultas
        st.success(f"🎫 **PASE DIARIO ACTIVO:** Te quedan {puntos_restantes} créditos para procesamiento masivo.")
    else:
        st.info("💡 **Versión Gratuita:** Actualmente tienes un límite de 10 puntos por día.")

    # 2. Diseño de Tarjetas de Planes (CSS Inyectado)
    st.markdown("""
        <style>
        .pro-card {
            padding: 25px;
            border-radius: 15px;
            text-align: center;
            margin-bottom: 20px;
            transition: transform 0.3s;
            border: 1px solid #e2e8f0;
            height: 100%;
        }
        .pro-card:hover { transform: translateY(-5px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
        .price { font-size: 2em; font-weight: 800; margin: 10px 0; color: #008080; }
        .feature-list { text-align: left; font-size: 0.85em; list-style: none; padding: 0; min-height: 120px; }
        .feature-list li { margin-bottom: 8px; color: #475569; }
        .card-diario { background-color: #ffffff; border-top: 5px solid #64748b; }
        .card-mensual { background-color: #f0f7ff; border-top: 5px solid #004d99; border-bottom: 2px solid #004d99; }
        .card-anual { background-color: #fffbeb; border-top: 5px solid #fbbf24; }
        </style>
    """, unsafe_allow_html=True)

    cp1, cp2, cp3 = st.columns(3)
    
    with cp1:
        st.markdown(f'''
            <div class="pro-card card-diario">
                <h3>Pase Diario</h3>
                <p class="price">S/ 5.00</p>
                <ul class="feature-list">
                    <li>✅ Acceso total por 24 horas</li>
                    <li>✅ Carga masiva (Hasta {LIMITE_FILAS_PASE_DIARIO} filas)</li>
                    <li>✅ Exportación KML e ingeniería</li>
                    <li>✅ Soporte vía WhatsApp</li>
                </ul>
            </div>
        ''', unsafe_allow_html=True)
    
    with cp2:
        st.markdown('''
            <div class="pro-card card-mensual">
                <h3>Plan Mensual</h3>
                <p class="price">S/ 20.00</p>
                <ul class="feature-list">
                    <li>✅ <b>Sin límites de procesamiento</b></li>
                    <li>✅ KML y DXF Masivo ilimitado</li>
                    <li>✅ Conversión de proyectos grandes</li>
                    <li>✅ Acceso Premium por 30 días</li>
                </ul>
            </div>
        ''', unsafe_allow_html=True)
    
    with cp3:
        st.markdown('''
            <div class="pro-card card-anual">
                <h3>Plan Anual</h3>
                <p class="price">S/ 180.00</p>
                <ul class="feature-list">
                    <li>✅ <b>Todo lo del Plan Mensual</b></li>
                    <li>✅ Ahorro real de S/ 60.00 al año</li>
                    <li>✅ Prioridad en nuevas funciones</li>
                    <li>✅ Atención VIP personalizada</li>
                </ul>
            </div>
        ''', unsafe_allow_html=True)

    st.write("---")

    # 3. Sección de Pago y Validación
    col_pago, col_cod = st.columns([1.2, 1], gap="large")
    
    with col_pago:
        st.subheader("📱 1. Realiza el Pago (Yape/Plin)")
        st.write("Envía el monto del plan elegido y recibe tu acceso inmediato.")
        
        # Caja de información de pago destacada con el mensaje corregido
        st.markdown("""
            <div style="background-color: #f8fafc; padding: 15px; border-radius: 10px; border-left: 5px solid #008080; margin-bottom: 20px;">
                <p style="margin-bottom: 5px;"><b>Titular:</b> Joan Gue.</p>
                <p style="margin-bottom: 5px;"><b>Número:</b> 924 886 915</p>
                <p style="margin-bottom: 0px; color: #008080; font-weight: bold; font-size: 1.1em;">✅ ¡Código instantáneo por WhatsApp!</p>
            </div>
        """, unsafe_allow_html=True)
        
        qr_path = "mi_qr.png"
        if os.path.exists(qr_path):
            # Contenedor para centrar la imagen
            _, col_qr_center, _ = st.columns([0.2, 1, 0.2])
            with col_qr_center:
                st.image(qr_path, caption="Escanea para pagar", use_container_width=True)
        else:
            st.warning("📸 Imagen 'mi_qr.png' no detectada.")

    with col_cod:
        st.subheader("🔑 2. Activar Acceso")
        st.write("Introduce el código de activación que te enviamos:")
        
        codigo_ingresado = st.text_input("Código de activación:", placeholder="Ej: SP-XXXXXX", label_visibility="collapsed")
        
        if st.button("🚀 VALIDAR Y ACTIVAR", use_container_width=True, type="primary"):
            if not codigo_ingresado:
                st.warning("Por favor, ingresa un código.")
            else:
                with st.status("🚀 Verificando credenciales...", expanded=True) as status:
                    try:
                        # 1. Lectura de base de datos
                        df_codigos = conn.read(worksheet="Usuarios", ttl=0)
                        
                        # Limpieza de datos para comparación
                        df_codigos['Codigo'] = df_codigos['Codigo'].astype(str).str.strip()
                        cod_limpio = codigo_ingresado.strip()
                        ahora = datetime.now()

                        if cod_limpio in df_codigos['Codigo'].values:
                            idx = df_codigos[df_codigos['Codigo'] == cod_limpio].index[0]
                            datos_cod = df_codigos.iloc[idx]
                            st.session_state.codigo_activo = cod_limpio
                            
                            # --- Caso PASE DIARIO ---
                            if datos_cod['Tipo'] == 'DIARIO':
                                if str(datos_cod['Usado']).upper() == 'NO':
                                    f_exp = ahora + timedelta(days=1)
                                    df_codigos.at[idx, 'Usado'] = 'SI'
                                    df_codigos.at[idx, 'Fecha_Activacion'] = ahora.strftime("%Y-%m-%d %H:%M:%S")
                                    df_codigos.at[idx, 'Fecha_Expiracion'] = f_exp.strftime("%Y-%m-%d %H:%M:%S")
                                    df_codigos.at[idx, 'Creditos_Usados'] = 0
                                    
                                    conn.update(worksheet="Usuarios", data=df_codigos)
                                    st.session_state.es_pase_diario = True
                                    st.session_state.consultas = 0
                                    status.update(label="✅ ¡Pase Diario Activado!", state="complete")
                                    st.balloons()
                                    time.sleep(2)
                                    st.rerun()
                                else:
                                    f_exp_dt = pd.to_datetime(datos_cod['Fecha_Expiracion'])
                                    if ahora < f_exp_dt:
                                        st.session_state.es_pase_diario = True
                                        st.session_state.consultas = int(pd.to_numeric(datos_cod['Creditos_Usados'], errors='coerce') or 0)
                                        status.update(label="✅ Suscripción recuperada.", state="complete")
                                        st.rerun()
                                    else:
                                        st.error("❌ Este código de Pase Diario ha expirado.")
                            
                            # --- Caso MENSUAL/ANUAL ---
                            elif datos_cod['Tipo'] in ['MENSUAL', 'ANUAL']:
                                st.session_state.es_pro = True
                                status.update(label=f"✅ Plan {datos_cod['Tipo']} Activado!", state="complete")
                                st.balloons()
                                time.sleep(2)
                                st.rerun()

                        elif cod_limpio == "ADMIN-JGZ-2026":
                            st.session_state.es_admin = True
                            status.update(label="👑 Modo Administrador Iniciado", state="complete")
                            st.rerun()
                        else:
                            st.error("❌ Código no válido o inexistente.")
                    except Exception as e:
                        st.error(f"Error de conexión: {e}")



# =========================================================
# BLOQUE 10: VISTA DE CONTACTO
# =========================================================
elif st.session_state.menu_actual == "CONTACTO":
    st.markdown("## 📞 Centro de Soporte Técnico")
    
    col_inf, col_frm = st.columns([1, 1], gap="large")
    
    with col_inf:
        st.write("") # Espaciador
        st.write("")
        st.link_button("💬 CHATEAR POR WHATSAPP", "https://wa.me/51924886915", use_container_width=True, type="primary")
        st.info("Atención inmediata para problemas técnicos o dudas sobre licencias.")

    with col_frm:
        with st.form("soporte_form"):
            st.markdown("##### Enviar Ticket de Consulta")
            u_nombre = st.text_input("Nombre completo:")
            u_correo = st.text_input("Correo o Celular:")
            u_msj = st.text_area("Cuéntanos tu duda o requerimiento:")
            
            if st.form_submit_button("ENVIAR TICKET"):
                if u_nombre and u_msj:
                    registrar_actividad("Consulta Soporte", f"De: {u_nombre} - {u_msj[:50]}...")
                    st.success("✅ Tu mensaje ha sido enviado. Te responderemos a la brevedad.")
                else:
                    st.warning("Por favor, completa los campos básicos.")

# =========================================================
# BLOQUE 11: PANEL ADMINISTRATIVO (JGZ)
# =========================================================
elif st.session_state.menu_actual == "ADMIN":
    if not st.session_state.es_admin:
        st.error("No tienes permisos para ver esto.")
        st.session_state.menu_actual = "CONVERTIDOR"
        st.rerun()

    st.title("🔐 Panel Maestro de Gestión")
    st.write("Control total de **SP Topo-Convert**")

    # Métricas en tiempo real
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Puntos Proc. Hoy", st.session_state.consultas)
    m2.metric("Descargas KML", st.session_state.descargas_kml)
    m3.metric("Versión", "1.2.5")
    m4.metric("Estado DB", "Conectado")

    st.divider()
    
    tab_logs, tab_cods = st.tabs(["📊 Historial de Logs", "🎫 Generador de Códigos"])
    
    with tab_logs:
        st.subheader("Actividad Global de Usuarios")
        try:
            df_logs = conn.read(worksheet="Logs", ttl=0)
            # Mostramos los últimos logs primero
            st.dataframe(df_logs.sort_index(ascending=False), use_container_width=True)
            
            if st.button("🔄 Refrescar Historial"):
                st.rerun()
        except:
            st.error("No se pudo cargar la hoja de Logs.")

    with tab_cods:
        st.subheader("Gestión de Licencias")
        st.write("Aquí puedes visualizar los códigos activos en la base de datos.")
        try:
            df_c = conn.read(worksheet="Usuarios", ttl=0)
            st.table(df_c)
        except:
            st.info("Configura una hoja llamada 'Codigos' en tu GSheets para gestionar ventas.")

    if st.button("🚪 CERRAR SESIÓN ADMIN"):
        st.session_state.es_admin = False
        st.session_state.menu_actual = "CONVERTIDOR"
        st.rerun()