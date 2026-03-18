import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import re
import calendar
import io
import os
import glob
from pathlib import Path
import html as html_lib
import threading

# ----------------------------------------------------------------------------- 
# IMPORTAR MÓDULOS LOCALES
# -----------------------------------------------------------------------------
import styles
import charts
import utils
import importlib
importlib.reload(utils)
from create_coord_report import create_coordinadora_pdf
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import tempfile

def render_excel_download(label, df, file_name):
    """
    Genera un archivo Excel desde un DataFrame y utiliza st.download_button con estilo centralizado.
    """
    import pandas as pd
    import streamlit as st
    import io

    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="DATOS")
        buffer.seek(0)

        st.download_button(
            label=str(label).upper(),
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"ERROR GENERANDO EXCEL: {e}")

def generar_pdf_propuesta(rows_prop: list, comp_df: pd.DataFrame, fecha_str: str) -> bytes:
    """
    Genera un PDF con la propuesta de reasignación.
    rows_prop : lista de dicts {Programa, Coordinadora Actual, Coordinadora Propuesta}
    comp_df   : DataFrame con columnas [COORDINADORA RESPONSABLE, Puntaje_Actual, Puntaje_Propuesta]
    fecha_str : string de fecha para el encabezado
    """
    from fpdf import FPDF

    # ── 1. Gráfico Antes vs Después con matplotlib ─────────────────────────────
    fig, ax = plt.subplots(figsize=(10, 4))
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#1e293b")

    coords  = [" ".join(str(c).split()[:2]) for c in comp_df["COORDINADORA RESPONSABLE"]]
    x       = range(len(coords))
    width   = 0.35
    bars_a  = ax.bar([i - width/2 for i in x], comp_df["Puntaje_Actual"],    width, label="Actual",    color="#ef4444", alpha=0.85)
    bars_p  = ax.bar([i + width/2 for i in x], comp_df["Puntaje_Propuesta"], width, label="Propuesta", color="#10b981", alpha=0.85)

    if not comp_df.empty and comp_df["Puntaje_Actual"].sum() > 0:
        avg = comp_df["Puntaje_Actual"].sum() / max(len(comp_df), 1)
        ax.axhline(avg, color="#f59e0b", linestyle="--", linewidth=1.2, label=f"Ideal: {avg:.1f}")

    ax.set_xticks(list(x))
    ax.set_xticklabels(coords, rotation=20, ha="right", color="#e2e8f0", fontsize=9)
    ax.tick_params(axis="y", colors="#e2e8f0")
    ax.set_ylabel("Carga (Puntaje)", color="#94a3b8")
    ax.set_title("Carga Antes vs Propuesta", color="#f1f5f9", fontsize=13, pad=10)
    ax.legend(facecolor="#1e293b", labelcolor="#e2e8f0", fontsize=9)
    ax.spines[:].set_color("#334155")
    ax.yaxis.grid(True, color="#334155", linewidth=0.5)
    plt.tight_layout()

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
        chart_path = tmp_img.name
        plt.savefig(chart_path, dpi=130, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)

    # ── 2. Construir PDF con fpdf2 ─────────────────────────────────────────────
    class PDF(FPDF):
        def header(self):
            self.set_fill_color(15, 23, 42)
            self.rect(0, 0, 210, 22, "F")
            self.set_font("Helvetica", "B", 13)
            self.set_text_color(241, 245, 249)
            self.set_y(6)
            self.cell(0, 10, "Propuesta de Reasignación de Carga", align="C")
            self.ln(2)
            self.set_font("Helvetica", "", 8)
            self.set_text_color(100, 116, 139)
            self.cell(0, 5, f"Generado: {fecha_str}", align="C")
            self.ln(6)

        def footer(self):
            self.set_y(-12)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(100, 116, 139)
            self.cell(0, 10, f"Página {self.page_no()}", align="C")

    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_margins(14, 26, 14)

    # Subtítulo tabla
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(14, 165, 233)
    pdf.cell(0, 7, "Detalle de Reasignaciones", ln=True)
    pdf.ln(1)

    # Encabezados tabla
    col_widths = [90, 45, 45]
    headers    = ["Programa", "Coordinadora Actual", "Coordinadora Propuesta"]
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(30, 41, 59)
    pdf.set_text_color(226, 232, 240)
    for w, h in zip(col_widths, headers):
        pdf.cell(w, 7, h, border=1, fill=True, align="C")
    pdf.ln()

    # Filas
    pdf.set_font("Helvetica", "", 7.5)
    for i, row in enumerate(rows_prop):
        fill = i % 2 == 0
        if fill:
            pdf.set_fill_color(30, 41, 59)
        else:
            pdf.set_fill_color(15, 23, 42)
        pdf.set_text_color(226, 232, 240)
        programa  = str(row.get("Programa", ""))[:60]
        coord_act = str(row.get("Coordinadora Actual", ""))[:30]
        coord_pro = str(list(row.values())[-1])[:30]
        pdf.cell(col_widths[0], 6, programa,  border="LR", fill=fill)
        pdf.cell(col_widths[1], 6, coord_act, border="LR", fill=fill, align="C")
        pdf.set_text_color(16, 185, 129)
        pdf.cell(col_widths[2], 6, coord_pro, border="LR", fill=fill, align="C")
        pdf.set_text_color(226, 232, 240)
        pdf.ln()
    # Línea cierre tabla
    pdf.cell(sum(col_widths), 0, "", border="T")
    pdf.ln(6)

    # Gráfico
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(14, 165, 233)
    pdf.cell(0, 7, "Impacto Visual: Carga Antes vs Propuesta", ln=True)
    pdf.ln(1)
    pdf.image(chart_path, x=14, w=182)
    pdf.ln(4)

    # Métricas resumen
    if not comp_df.empty:
        before_std = comp_df["Puntaje_Actual"].std()
        after_std  = comp_df["Puntaje_Propuesta"].std()
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(148, 163, 184)
        pdf.cell(0, 6, f"Desv. estándar antes: {before_std:.2f}    "
                       f"Desv. estándar propuesta: {after_std:.2f}    "
                       f"Programas a mover: {len(rows_prop)}", ln=True)

    # Limpiar imagen temporal
    try:
        os.remove(chart_path)
    except Exception:
        pass

    return bytes(pdf.output())


# ----------------------------------------------------------------------------- 
# CONFIGURACIÓN DE PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Calendario Postgrado Ejecutivos",
    layout="wide",
    page_icon="🎓",
)

# Inyectar tema oscuro unificado y meta tags de protección
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)
st.markdown(styles.APP_STYLE, unsafe_allow_html=True)

# ----------------------------------------------------------------------------- 
# ESTADO DE LA APLICACIÓN (Persistencia de Datos)
# -----------------------------------------------------------------------------
if "df_base" not in st.session_state:
    st.session_state.df_base = pd.DataFrame()
if "df_reservas" not in st.session_state:
    st.session_state.df_reservas = pd.DataFrame()
if "home_kpi_filter" not in st.session_state:
    st.session_state.home_kpi_filter = "total"
if "nav_categoria" not in st.session_state:
    st.session_state.nav_categoria = "📊 Dashboard"

# Asignar variables globales desde el estado
df_base = st.session_state.df_base
df_reservas = st.session_state.df_reservas


# ----------------------------------------------------------------------------- 
# FUNCIONES AUXILIARES
# -----------------------------------------------------------------------------
def _normalizar_columnas_reservas(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza columnas de un archivo de reservas al esquema estándar."""
    df.columns = df.columns.astype(str).str.strip().str.upper()
    res_map = {}
    for c in df.columns:
        if any(x in c for x in ["FECHA", "DATE", "DIA"]):   res_map[c] = "FECHA"
        elif any(x in c for x in ["INICIO", "START", "DESDE", "HORA I"]): res_map[c] = "HORA_INICIO"
        elif any(x in c for x in ["FIN", "END", "HASTA", "HORA F"]):     res_map[c] = "HORA_FIN"
        elif any(x in c for x in ["SALA", "ROOM", "LUGAR", "AULA"]):     res_map[c] = "SALA"
        elif any(x in c for x in ["EVENTO", "ACTIVITY", "MOTIVO", "ASIGNATURA", "NOMBRE"]): res_map[c] = "EVENTO"
        elif any(x in c for x in ["SOLICITANTE", "PROFESOR", "RESPONSABLE"]): res_map[c] = "SOLICITANTE"
    df = df.rename(columns=res_map)
    if "FECHA" in df.columns:
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors='coerce')
    return df

def calc_mod(sede):
    """Determina la modalidad según el nombre de sede."""
    s = str(sede).upper()
    if "ONLINE" in s or "ZOOM" in s or "VIRTUAL" in s: return "Online"
    if "HYBRID" in s or "HIBRID" in s: return "Híbrido"
    return "Presencial"

def sorted_clean(series_or_list) -> list:
    """Ordena una serie/lista eliminando NaN y asegurando que todo sea str.
    Evita TypeError: '<' not supported between instances of 'float' and 'str'.
    """
    if hasattr(series_or_list, 'dropna'):
        values = series_or_list.dropna().unique().tolist()
    else:
        values = [v for v in series_or_list if v == v and v is not None]
    
    cleaned = set()
    for v in values:
        s = str(v).strip()
        if s.upper() not in ("", "NAN", "NONE", "N/A", "NAT", "NA", "POR DEFINIR"):
            cleaned.add(s)
            
    return sorted(list(cleaned))

# ----------------------------------------------------------------------------- 
# FUNCIONES DE CÁLCULO
# -----------------------------------------------------------------------------
def resumen_coordinadoras_semana(df_filtrado: pd.DataFrame) -> pd.DataFrame:
    if df_filtrado.empty:
        return pd.DataFrame()
    cols_req = ["Dia_Semana", "Modalidad_Calc", "PROGRAMA", "COORDINADORA RESPONSABLE"]
    if not all(c in df_filtrado.columns for c in cols_req):
        return pd.DataFrame()
    
    # Agrupación robusta que evita desalineación de índices
    resumen = df_filtrado.groupby("COORDINADORA RESPONSABLE").agg(
        dias_clase_semana=("Dia_Semana", "nunique"),
        Modalidades=("Modalidad_Calc", lambda x: ", ".join(sorted(x.dropna().unique()))),
        Programas=("PROGRAMA", lambda x: ", ".join(sorted(x.dropna().unique())))
    ).reset_index()
    
    return resumen.rename(columns={
        "COORDINADORA RESPONSABLE": "Coordinadora",
        "dias_clase_semana": "Días Activos (Semana)"
    })


def resumen_modalidad(df_f):
    if df_f.empty or "Modalidad_Calc" not in df_f.columns:
        return pd.DataFrame()
    return df_f.groupby("Modalidad_Calc", as_index=False).agg(
        Sesiones=("PROGRAMA", "size"),
        Programas=("PROGRAMA", "nunique")
    ).sort_values("Sesiones", ascending=False)


def resumen_sede(df_f):
    if df_f.empty or "SEDE" not in df_f.columns:
        return pd.DataFrame()
    return df_f.groupby("SEDE", as_index=False).agg(
        Sesiones=("PROGRAMA", "size"),
        Programas=("PROGRAMA", "nunique")
    ).sort_values("Sesiones", ascending=False)


def resumen_calidad_datos(df_all):
    campos = ["DIAS/FECHAS", "PROGRAMA", "COORDINADORA RESPONSABLE", "Modalidad_Calc", "SEDE", "HORARIO"]
    data   = []
    total  = len(df_all)
    if total == 0:
        return pd.DataFrame()
    for col in campos:
        if col in df_all.columns:
            faltantes = df_all[col].isna().sum()
            data.append({"Campo": col, "Faltantes": faltantes, "%": round((faltantes/total)*100, 1)})
    return pd.DataFrame(data)


# -----------------------------------------------------------------------------
# PREPROCESAMIENTO DE DATOS (Columnas Calculadas)
# -----------------------------------------------------------------------------
if not df_base.empty:
    # Asegurar columnas base
    cols_base = ["PROGRAMA", "COORDINADORA RESPONSABLE", "SEDE", "HORARIO", "PROFESOR", "ASIGNATURA", "SALA"]
    for col in cols_base:
        if col not in df_base.columns:
            df_base[col] = "N/A"

    if "DIAS/FECHAS" in df_base.columns:
        # Día de la Semana y Mes
        dias_es = {0:"Lunes", 1:"Martes", 2:"Miércoles", 3:"Jueves", 4:"Viernes", 5:"Sábado", 6:"Domingo"}
        meses_es = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 
                    7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
        
        if "Dia_Semana" not in df_base.columns:
            df_base["Dia_Semana"] = df_base["DIAS/FECHAS"].dt.dayofweek.map(dias_es)
        if "Mes" not in df_base.columns:
            df_base["Mes"] = df_base["DIAS/FECHAS"].dt.month.map(meses_es)
        if "Modalidad_Calc" not in df_base.columns:
            df_base["Modalidad_Calc"] = df_base["SEDE"].apply(calc_mod)
        if "Programa_Abrev" not in df_base.columns:
            df_base["Programa_Abrev"] = df_base["PROGRAMA"].apply(utils.abbreviate_program_name)

# Variable global hoy
hoy = pd.Timestamp.now()


with st.sidebar:
    st.markdown("### 🎓 Calendario UAI")
    st.markdown('<p style="font-size:12px; color:#94a3b8; margin-bottom: 15px;">Arrastra o selecciona el Calendario Académico (Excel / CSV)</p>', unsafe_allow_html=True)
    
    if "df_hash" not in st.session_state:
        st.session_state.df_hash = None
    
    # Puente de navegación para accesos rápidos
    if "goto_nav" in st.session_state:
        target = st.session_state.pop("goto_nav")
        st.session_state["nav_categoria"] = target["cat"]
        st.session_state["sub_tab_directo"] = target.get("sub")

    # Si hay datos cargados, mostramos los uploaders en el panel lateral para permitir cambios
    if not df_base.empty:
        st.markdown('<p style="font-size:12px; color:#94a3b8; margin-bottom: 15px;">📥 Cambiar Archivo Académico</p>', unsafe_allow_html=True)
        uploaded_main_side = st.file_uploader(
            "Cargar Archivo Académico",
            type=["xlsx", "xls", "csv"],
            key="uploader_principal_side",
            label_visibility="collapsed"
        )
        
        if uploaded_main_side:
            file_stats_up = f"upload_{uploaded_main_side.name}_{uploaded_main_side.size}"
            if st.session_state.df_hash != file_stats_up:
                with st.spinner("📊 Procesando..."):
                    try:
                        loaded_df = utils.load_data(io.BytesIO(uploaded_main_side.getvalue()))
                        if loaded_df is not None and not loaded_df.empty:
                            st.session_state.df_base = loaded_df
                            st.session_state.df_hash = file_stats_up
                            st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
            
            df_base = st.session_state.df_base
            if not df_base.empty:
                st.success(f"✓ {len(df_base):,} registros")

        st.markdown("---")
        st.markdown("### 📅 Reservas (Opcional)")
        st.markdown('<p style="font-size:12px; color:#94a3b8; margin-bottom: 15px;">Sube o actualiza tus reservas (Excel/CSV)</p>', unsafe_allow_html=True)
        
        uploaded_reservas_side = st.file_uploader(
            "Cargar Reservas (Opcional)",
            type=["xlsx", "xls", "csv"],
            key="uploader_reservas_side",
            label_visibility="collapsed"
        )
        
        if uploaded_reservas_side:
            try:
                res_bytes = uploaded_reservas_side.getvalue()
                if uploaded_reservas_side.name.endswith('.csv'):
                    temp_res = pd.read_csv(io.BytesIO(res_bytes))
                else:
                    temp_res = pd.read_excel(io.BytesIO(res_bytes))
                st.session_state.df_reservas = _normalizar_columnas_reservas(temp_res)
                st.success(f"✅ {len(st.session_state.df_reservas)} reservas")
            except Exception as e:
                st.error(f"Error: {e}")

        st.markdown("---")
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("🧹 Reiniciar", use_container_width=True):
                utils.reset_filters()
                st.rerun()
        with col_btn2:
            if st.button("🔄 Caché", use_container_width=True):
                st.cache_data.clear()
                st.rerun()

    # --- MEJORA: Sistema de descargas Base64 Universal ---
    def render_static_download(label, data_bytes, file_name):
        import streamlit as st
        ext = file_name.split('.')[-1].lower()
        mime_types = {
            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'xls': 'application/vnd.ms-excel', 
            'pdf': 'application/pdf',
            'ics': 'text/calendar', 
            'csv': 'text/csv'
        }
        mime = mime_types.get(ext, 'application/octet-stream')
        
        st.download_button(
            label=str(label).upper(),
            data=data_bytes,
            file_name=file_name,
            mime=mime,
            use_container_width=True
        )

    if not df_base.empty:
        st.markdown("---")
        st.markdown("### 📤 Centro de Descargas")
        
        # 1. BASE DE DATOS
        st.markdown("#### 📊 Base Completa")
        render_excel_download(
            "💾 Descargar Excel (.xlsx)",
            df_base,
            f"Gestion_Postgrado_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        # 2. INFORME PERSONALIZADO (Compacto)
        st.markdown("#### 🎯 Informe Personalizado")
        st.caption("Filtra datos específicos para exportar.")
        
        tipo_filtro = st.radio("Filtrar por:", ["Coordinadora", "Programa", "Fechas"], horizontal=True, label_visibility="collapsed")
        
        df_rep = df_base.copy()
        if tipo_filtro == "Coordinadora" and "COORDINADORA RESPONSABLE" in df_base.columns:
            coords = sorted_clean(df_base["COORDINADORA RESPONSABLE"])
            coord_sel = st.selectbox("Seleccionar:", coords, label_visibility="collapsed")
            df_rep = df_base[df_base["COORDINADORA RESPONSABLE"] == coord_sel]
            
            if not df_rep.empty:
                 try:
                    pdf_data = create_coordinadora_pdf(coord_sel, df_rep, f"Reporte General")
                    render_static_download(
                        f"📄 Descargar Informe PDF",
                        pdf_data,
                        f"Informe_{coord_sel.replace(' ','_')[:10]}.pdf"
                    )
                 except Exception as e:
                    st.error(f"Error al generar PDF: {e}")
            
        elif tipo_filtro == "Programa" and "PROGRAMA" in df_base.columns:
            progs = sorted_clean(df_base["PROGRAMA"])
            prog_sel = st.selectbox("Seleccionar:", progs, label_visibility="collapsed")
            df_rep = df_base[df_base["PROGRAMA"] == prog_sel]
            
            if not df_rep.empty:
                 render_excel_download(
                    f"⬇️ Descargar Excel ({len(df_rep)})",
                    df_rep,
                    f"Reporte_Programa_{datetime.now().strftime('%Y%m%d')}.xlsx"
                 )
            
        elif tipo_filtro == "Fechas" and "DIAS/FECHAS" in df_base.columns:
            c_f1, c_f2 = st.columns(2)
            f_min = df_base["DIAS/FECHAS"].min()
            f_max = df_base["DIAS/FECHAS"].max()
            f_ini = c_f1.date_input("Desde", value=f_min, label_visibility="collapsed")
            f_fin = c_f2.date_input("Hasta", value=f_max, label_visibility="collapsed")
            df_rep = df_base[(df_base["DIAS/FECHAS"].dt.date >= f_ini) & (df_base["DIAS/FECHAS"].dt.date <= f_fin)]
            
            if not df_rep.empty:
                 render_excel_download(
                    f"⬇️ Descargar Excel ({len(df_rep)})",
                    df_rep,
                    f"Reporte_Fechas_{datetime.now().strftime('%Y%m%d')}.xlsx"
                 )

        # 3. DOCUMENTACIÓN
        st.markdown("#### 📘 Ayuda")
        try:
            manual_path = Path(__file__).parent / "Manual_de_Usuario.pdf"
            if manual_path.exists():
                with open(manual_path, "rb") as f:
                    pdf_data = f.read()
                
                st.download_button(
                    label="📄 MANUAL DE USUARIO (PDF)",
                    data=pdf_data,
                    file_name="Manual_Usuario_Calendario.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.caption("Manual no disponible.")
        except Exception:
            pass
    
# ----------------------------------------------------------------------------- 
# STOP si no hay datos - ORIGINAL LOGO EMPTY STATE
# -----------------------------------------------------------------------------
if df_base is None or df_base.empty:
    # Contenido Central - Portada
    st.markdown("""
    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; margin-top: 5vh; text-align: center;">
        <div style="font-size: 80px; margin-bottom: 10px; filter: drop-shadow(0 0 20px rgba(14,165,233,0.3));">🎓</div>
        <h1 style="font-family: 'Outfit', sans-serif; font-weight: 800; font-size: 36px; color: #f1f5f9; margin-bottom: 10px;">
            Calendario Postgrado Ejecutivos
        </h1>
        <p style="color: #94a3b8; font-size: 18px; margin-bottom: 30px; max-width: 500px;">
            Carga tus archivos a continuación para iniciar el sistema.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Área de carga centralizada
    col_esp_1, col_center_upload, col_esp_2 = st.columns([1, 2, 1])
    with col_center_upload:
        st.markdown("### 📥 1. Archivo Académico Base")
        uploaded_main_center = st.file_uploader(
            "Sube tu Calendario Académico aquí (Excel o CSV)",
            type=["xlsx", "xls", "csv"],
            key="uploader_principal_center"
        )
        
        if uploaded_main_center:
            file_stats_up = f"upload_{uploaded_main_center.name}_{uploaded_main_center.size}"
            if st.session_state.df_hash != file_stats_up:
                with st.spinner("📊 Procesando..."):
                    try:
                        loaded_df = utils.load_data(io.BytesIO(uploaded_main_center.getvalue()))
                        if loaded_df is not None and not loaded_df.empty:
                            st.session_state.df_base = loaded_df
                            st.session_state.df_hash = file_stats_up
                            st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📅 2. Reservas (Opcional)")
        uploaded_reservas_center = st.file_uploader(
            "Sube las reservas de salas aquí (Opcional)",
            type=["xlsx", "xls", "csv"],
            key="uploader_reservas_center"
        )
        
        if uploaded_reservas_center:
            try:
                res_bytes = uploaded_reservas_center.getvalue()
                if uploaded_reservas_center.name.endswith('.csv'):
                    temp_res = pd.read_csv(io.BytesIO(res_bytes))
                else:
                    temp_res = pd.read_excel(io.BytesIO(res_bytes))
                st.session_state.df_reservas = _normalizar_columnas_reservas(temp_res)
                st.success(f"✅ {len(st.session_state.df_reservas)} reservas")
            except Exception as e:
                st.error(f"Error: {e}")
    
    st.stop()
# EXPLICACIÓN: Permite jump rápido a las secciones más importantes
# Los iconos ayudan a identificar visualmente cada categoría
# Estilos manejados en styles.py

# Categorías de navegación
categorias = {
    "📊 Dashboard": ["🏠 Home", "🌐 Global", "👩‍💼 Coordinadoras", "📊 Comparativa", "📈 Analytics"],
    "📚 Programas": ["🧾 Resumen", "🏫 Calidad"],
    "📅 Calendarios": ["📅 Calendario", "📅 Reservas"],
    "⚙️ Administración": ["🔒 Gestión", "🧩 Turnos", "✅ Validaciones"],
    "📘 Ayuda": ["📖 Manual de Usuario", "💡 FAQs"]
}

# --- ENCABEZADO NAVEGACIÓN ÚNICA ---
# EXPLICACIÓN: Usamos columnas con botones para una navegación interactiva y visual
col_nav_title, col_nav1, col_nav2, col_nav3, col_nav4 = st.columns([1.2, 2, 2, 2, 2])

with col_nav_title:
    st.markdown('<div style="padding-top:10px; font-size:11px; font-weight:800; color:#38bdf8; letter-spacing:0.2em;">DASHBOARD V2</div>', unsafe_allow_html=True)

with col_nav1:
    if st.button("📊 Dashboard", key="pill_dash", use_container_width=True):
        st.session_state.nav_categoria = "📊 Dashboard"
        st.rerun()
with col_nav2:
    if st.button("📚 Programas", key="pill_prog", use_container_width=True):
        st.session_state.nav_categoria = "📚 Programas"
        st.rerun()
with col_nav3:
    if st.button("📅 Calendarios", key="pill_cal", use_container_width=True):
        st.session_state.nav_categoria = "📅 Calendarios"
        st.rerun()
with col_nav4:
    if st.button("⚙️ Admin", key="pill_admin", use_container_width=True):
        st.session_state.nav_categoria = "⚙️ Administración"
        st.rerun()

# Recuperar la selección
categoria_sel = st.session_state.nav_categoria


# Filtro global por nombre abreviado — en sidebar (siempre visible)
with st.sidebar:
    st.markdown("---")
    st.markdown("### 🔎 Filtro Global")
    abrevs_disponibles = sorted_clean(df_base["Programa_Abrev"]) if "Programa_Abrev" in df_base.columns else []
    sel_abrev_global = st.multiselect(
        "Programa (nombre abreviado)",
        abrevs_disponibles,
        key="filtro_abrev_global",
        placeholder="Todos los programas..."
    )
    if sel_abrev_global:
        n_sel = df_base["Programa_Abrev"].isin(sel_abrev_global).sum()
        st.caption(f"✅ {len(sel_abrev_global)} programa(s) · {n_sel:,} registros")
    else:
        st.caption(f"📊 Mostrando todos: {len(df_base):,} registros")

# Aplicar filtro global: filtra df_base para TODOS los tabs
if sel_abrev_global and "Programa_Abrev" in df_base.columns:
    df_base = df_base[df_base["Programa_Abrev"].isin(sel_abrev_global)].copy()

# --- INICIALIZACIÓN DE VARIABLES DE TABS ---
# Prevenir errores de referencia antes de la asignación
tab_home = tab_global = tab1 = tab2 = tab_analytics = None
tab4 = tab5 = tab_calendario = tab_reservas = None
tab_gestion = tab_turnos = tab_validaciones = tab_reasignacion = None
tab_manual = tab_faqs = None

# --- TABS SEGÚN CATEGORÍA SELECCIONADA ---
# EXPLICACIÓN: Implementamos Reordenamiento Dinámico para "Navegación Directa"
if categoria_sel == "📊 Dashboard":
    tabs_list = ["🏠 Home", "🌐 Global", "👩‍💼 Coordinadoras", "📊 Comparativa", "📈 Analytics"]
    # Si venimos de un acceso rápido, movemos el deseado al principio
    if st.session_state.get("sub_tab_directo") in tabs_list:
        sub = st.session_state.pop("sub_tab_directo")
        tabs_list.remove(sub)
        tabs_list.insert(0, sub)
    
    # Desempacar dinámicamente según lo que haya primero
    created_tabs = st.tabs(tabs_list)
    tab_home = created_tabs[tabs_list.index("🏠 Home")]
    tab_global = created_tabs[tabs_list.index("🌐 Global")]
    tab1 = created_tabs[tabs_list.index("👩‍💼 Coordinadoras")]
    tab2 = created_tabs[tabs_list.index("📊 Comparativa")]
    tab_analytics = created_tabs[tabs_list.index("📈 Analytics")]

elif categoria_sel == "📚 Programas":
    tabs_list = ["📋 Resumen", "✔️ Calidad & Sede"]
    if st.session_state.get("sub_tab_directo") == "Programas":
        st.session_state.pop("sub_tab_directo") # Ya estamos aquí
    created_tabs = st.tabs(tabs_list)
    tab4 = created_tabs[tabs_list.index("📋 Resumen")]
    tab5 = created_tabs[tabs_list.index("✔️ Calidad & Sede")]

elif categoria_sel == "📅 Calendarios":
    tabs_list = ["📅 Calendario Académico", "🏢 Reservas"]
    if st.session_state.get("sub_tab_directo") in ["Calendario", "Salas"]:
        sub_name = "📅 Calendario Académico" if st.session_state["sub_tab_directo"] == "Calendario" else "🏢 Reservas"
        st.session_state.pop("sub_tab_directo")
        tabs_list.remove(sub_name)
        tabs_list.insert(0, sub_name)
    
    created_tabs = st.tabs(tabs_list)
    tab_calendario = created_tabs[tabs_list.index("📅 Calendario Académico")]
    tab_reservas = created_tabs[tabs_list.index("🏢 Reservas")]

elif categoria_sel == "⚙️ Administración":
    tab_gestion, tab_turnos, tab_validaciones = st.tabs([
        "🔒 Gestión", "🧩 Turnos", "✅ Validaciones"
    ])

elif categoria_sel == "📘 Ayuda":
    tab_manual, tab_faqs = st.tabs(["📖 Manual de Usuario", "💡 FAQs"])
# =============================================================================
# TAB HOME: DASHBOARD EJECUTIVO
# =============================================================================
if tab_home:
  with tab_home:
    st.markdown("## 🎯 DASHBOARD EJECUTIVO")
    
    # --- CONFIGURACIÓN DE FILTROS KPI ---
    if "home_kpi_filter" not in st.session_state:
        st.session_state.home_kpi_filter = "total"

    today = datetime.now()
    df_hoy_base = df_base[df_base["DIAS/FECHAS"].dt.date == today.date()] if "DIAS/FECHAS" in df_base.columns else pd.DataFrame()
    df_semana_base = df_base[(df_base["DIAS/FECHAS"] >= today) & (df_base["DIAS/FECHAS"] <= today + pd.Timedelta(days=7))] if "DIAS/FECHAS" in df_base.columns else pd.DataFrame()

    # Filtrar df_home según el KPI seleccionado para las tablas de abajo
    df_home = df_base.copy()
    
    # Asegurar que esté ordenado cronológicamente para el dashboard
    if "DIAS/FECHAS" in df_home.columns:
        df_home = df_home.sort_values("DIAS/FECHAS")
        
    if st.session_state.home_kpi_filter == "hoy":
        df_home = df_hoy_base
    elif st.session_state.home_kpi_filter == "semana":
        df_home = df_semana_base
    elif st.session_state.home_kpi_filter == "total":
        # Mostrar las próximas 10 sesiones desde hoy en adelante si es posible
        if "DIAS/FECHAS" in df_home.columns:
            df_home = df_home[df_home["DIAS/FECHAS"] >= today.replace(hour=0, minute=0, second=0)]
    
    # Los estilos ahora se manejan centralizadamente en styles.py

    k1, k2, k3, k4 = st.columns(4)
    
    with k1:
        if st.button(f"📚 {len(df_base):,}\nTOTAL SESIONES", key="kpi_tot", help="VER TODAS LAS SESIONES", use_container_width=True):
            st.session_state.home_kpi_filter = "total"
            st.rerun()
    with k2:
        num_progs = df_base["PROGRAMA"].nunique() if "PROGRAMA" in df_base.columns else 0
        if st.button(f"🎓 {num_progs}\nPROGRAMAS", key="kpi_prog", help="VISTA DE PROGRAMAS", use_container_width=True):
            st.session_state.home_kpi_filter = "programas"
            st.rerun()
    with k3:
        if st.button(f"🗓️ {len(df_hoy_base)}\nCLASES HOY", key="kpi_hoy", help="VER CLASES DE HOY", use_container_width=True):
            st.session_state.home_kpi_filter = "hoy"
            st.rerun()
    with k4:
        if st.button(f"📅 {len(df_semana_base)}\nPRÓXIMOS 7 DÍAS", key="kpi_sem", help="VER CLASES DE LA SEMANA", use_container_width=True):
            st.session_state.home_kpi_filter = "semana"
            st.rerun()
    
    # Indicador de Filtro
    filter_labels = {"total": "GENERAL", "hoy": "HOY", "semana": "PRÓXIMOS 7 DÍAS", "programas": "PROGRAMAS"}
    st.caption(f"🔍 VISTA ACTIVA: **{filter_labels.get(st.session_state.home_kpi_filter)}**")
    
    st.markdown("---")
    
    # --- SECCIÓN DE MONITOREO ---
    # Alertas y Próximas Clases ahora se apilan para una mejor lectura de datos expandidos
    
    st.markdown("### 🚨 ALERTAS ACTIVAS")
    alertas = []
        
    # Cursos que terminan esta semana
    if "DIAS/FECHAS" in df_home.columns and "PROGRAMA" in df_home.columns:
        fin_semana = today + pd.Timedelta(days=7)
        ultimas_fechas = df_home.groupby("PROGRAMA")["DIAS/FECHAS"].max().reset_index()
        cursos_terminan = ultimas_fechas[(ultimas_fechas["DIAS/FECHAS"] >= today) & (ultimas_fechas["DIAS/FECHAS"] <= fin_semana)]
        if len(cursos_terminan) > 0:
            alertas.append({"Tipo": "⚠️ Cursos finalizan", "Detalle": f"{len(cursos_terminan)} programas terminan esta semana", "Prioridad": "Alta"})
    
    # Profesores sin asignar
    if "PROFESOR" in df_home.columns:
        sin_prof = df_home[df_home["PROFESOR"].isna() | (df_home["PROFESOR"].astype(str).str.upper().isin(["NAN","","POR DEFINIR","SIN PROFESOR"]))]
        if len(sin_prof) > 0:
            alertas.append({"Tipo": "👨‍🏫 Sin Profesor", "Detalle": f"{len(sin_prof)} sesiones sin docente", "Prioridad": "Media"})
    
    # Salas sin asignar
    if "SALA" in df_home.columns:
        sin_sala = df_home[df_home["SALA"].isna() | (df_home["SALA"].astype(str).str.upper().isin(["NAN","","POR ASIGNAR","SIN SALA"]))]
        if len(sin_sala) > 0:
            alertas.append({"Tipo": "🏢 Sin Sala", "Detalle": f"{len(sin_sala)} sesiones sin sala", "Prioridad": "Media"})
    
    # Días críticos
    if "COORDINADORA RESPONSABLE" in df_home.columns:
        carga = df_home.groupby(["COORDINADORA RESPONSABLE","DIAS/FECHAS"])["PROGRAMA"].nunique().reset_index()
        criticos = carga[carga["PROGRAMA"] > 2]
        if len(criticos) > 0:
            alertas.append({"Tipo": "🔥 Días Críticos", "Detalle": f"{len(criticos)} días con >2 programas", "Prioridad": "Alta"})
    
    if alertas:
        # Presentar alertas como tarjetas de estado
        for al in alertas:
            color_class = "pill-orange" if al["Prioridad"] == "Alta" else "pill-blue"
            st.markdown(f"""
            <div style="background:rgba(30,41,59,0.5); border-left:4px solid {'#f59e0b' if al['Prioridad']=='Alta' else '#3b82f6'}; 
                        border-radius:8px; padding:12px; margin-bottom:8px;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-weight:700; font-size:12px; color:#f1f5f9;">{al['Tipo']}</span>
                    <span class="kpi-pill {color_class}" style="font-size:10px;">{al['Prioridad']}</span>
                </div>
                <div style="font-size:11px; color:#94a3b8; margin-top:4px;">{al['Detalle']}</div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.success("✅ No hay alertas pendientes")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📅 PRÓXIMAS CLASES")
    
    # --- PRÓXIMAS CLASES (CON MÁS DETALLE) ---
    df_proximas = df_home.copy()
    if not df_proximas.empty:
        # Renderizado de Tarjeta Premium de Clases
        rows_html = ""
        # Limitar a 10 para no saturar
        for _, row in df_proximas.head(10).iterrows():
            horario = row.get("HORARIO", "00:00")
            programa = str(row.get("PROGRAMA", "Sin Programa"))[:50] + ("..." if len(str(row.get("PROGRAMA", ""))) > 50 else "")
            sede = str(row.get("SEDE", "Presencial")).upper()
            modalidad_class = "online" if "ONLINE" in sede else "presencial"
            modalidad_label = "Online" if "ONLINE" in sede else "Presencial"
            
            # Nuevos datos solicitados
            fecha_dt = row.get("DIAS/FECHAS")
            fecha_str = fecha_dt.strftime("%d/%m") if pd.notnull(fecha_dt) else ""
            coord = str(row.get("COORDINADORA RESPONSABLE", "Sin Coord."))[:25]
            
            rows_html += (
                f'<div class="table-row">'
                f'<span>{fecha_str}</span>'
                f'<span style="color:#38bdf8">{horario}</span>'
                f'<span>{programa}</span>'
                f'<span>{coord}</span>'
                f'<span><span class="pill {modalidad_class}">{modalidad_label}</span></span>'
                f'</div>'
            )
        
        html_card = (
            f'<div class="card">'
            f'<div class="card-title">📋 Próximas Clases</div>'
            f'<div class="table-row header">'
            f'<span>Día</span><span>Horario</span><span>Programa</span><span>Coordinadora</span><span>Modalidad</span>'
            f'</div>'
            f'{rows_html}'
            f'</div>'
        )
        st.markdown(html_card, unsafe_allow_html=True)
    else:
        st.info("📭 No hay clases en esta selección")
    
    st.markdown("---")
    
    # --- NUEVOS CUADROS DESPLEGABLES SOLICITADOS ---
    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        with st.expander("📅 Sesiones de Hoy", expanded=True):
            if not df_hoy_base.empty:
                st.dataframe(df_hoy_base[["HORARIO", "PROGRAMA", "SALA"]].sort_values("HORARIO"), hide_index=True, use_container_width=True)
            else:
                st.write("No hay sesiones programadas para hoy.")
                
    with col_exp2:
        with st.expander("🗓️ Sesiones del Mes", expanded=False):
            if not df_base.empty:
                current_month = datetime.now().month
                df_mes = df_base[df_base["DIAS/FECHAS"].dt.month == current_month]
                if not df_mes.empty:
                    st.write(f"Total sesiones en el mes: **{len(df_mes)}**")
                    st.dataframe(df_mes[["DIAS/FECHAS", "PROGRAMA", "COORDINADORA RESPONSABLE"]].head(50), hide_index=True, use_container_width=True)
                else:
                    st.write("No hay sesiones registradas este mes.")
            else:
                st.write("Carga datos para ver el resumen mensual.")

    st.markdown("---")
    
    # --- GRÁFICO TENDENCIA ---
    st.markdown("### 📈 Tendencia Mensual")
    if "DIAS/FECHAS" in df_base.columns and not df_base.empty:
        df_trend = df_base.copy()
        df_trend["Mes_Str"] = df_trend["DIAS/FECHAS"].dt.to_period("M").astype(str)
        tendencia = df_trend.groupby("Mes_Str").size().reset_index(name="Sesiones")
        fig_trend = px.area(tendencia, x="Mes_Str", y="Sesiones")
        fig_trend.update_layout(
            height=250, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#e2e8f0", size=10), margin=dict(l=0, r=0, t=10, b=0)
        )
        st.plotly_chart(fig_trend, use_container_width=True)
    # --- QUICK NAVIGATION CHIPS ---
    st.markdown("### ⚡ Acceso Rápido")
    
    # Contenedor para botones horizontales
    # Contenedor para botones en cuadrícula (2x3)
    q_col1, q_col2, q_col3 = st.columns(3)
    with q_col1:
        if st.button("📅 Calendario Directo", key="btn_q_cal"):
            st.session_state.goto_nav = {"cat": "📅 Calendarios", "sub": "Calendario"}
            st.rerun()
    with q_col2:
        if st.button("👩‍💼 Coordinadoras", key="btn_q_coord"):
            st.session_state.goto_nav = {"cat": "📊 Dashboard", "sub": "👩‍💼 Coordinadoras"}
            st.rerun()
    with q_col3:
        if st.button("🏛️ Revisar Salas", key="btn_q_salas"):
            st.session_state.goto_nav = {"cat": "📅 Calendarios", "sub": "Salas"}
            st.rerun()

    q_col4, q_col5, q_col6 = st.columns(3)
    with q_col4:
        if st.button("📈 Analytics", key="btn_q_ana"):
            st.session_state.goto_nav = {"cat": "📊 Dashboard", "sub": "📈 Analytics"}
            st.rerun()
    with q_col5:
        if st.button("📚 Ver Programas", key="btn_q_progs"):
            st.session_state.goto_nav = {"cat": "📚 Programas", "sub": "Programas"}
            st.rerun()
    with q_col6:
        # Botón vacío o extra (Global por ejemplo)
        if st.button("🌐 Vista Global", key="btn_q_global"):
            st.session_state.goto_nav = {"cat": "📊 Dashboard", "sub": "🌐 Global"}
            st.rerun()

# =============================================================================
# TAB KANBAN: PANEL DE ESTADOS DE PROGRAMAS (REDISEÑADO)
# =============================================================================

# =============================================================================
if tab_analytics:
  with tab_analytics:
    if df_base.empty or "DIAS/FECHAS" not in df_base.columns:
        st.warning("⚠️ Primero sube un archivo en el panel lateral.")
    else:
        st.markdown("## 📈 Analytics Avanzado")
        mapa_meses = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
                      7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
    
        # --- FILTROS PRINCIPALES ---
        # EXPLICACIÓN: Estos filtros afectan todas las métricas del tab Analytics.
        # Permiten analizar datos por coordinadora específica o programa.
        st.markdown("### 🎛️ Filtros")
        
        col_filt1, col_filt2, col_filt3 = st.columns(3)
        
        with col_filt1:
            # Filtro de Coordinadora
            coords_analytics = ["Todas"] + sorted_clean(df_base["COORDINADORA RESPONSABLE"]) if "COORDINADORA RESPONSABLE" in df_base.columns else ["Todas"]
            sel_coord_analytics = st.selectbox("👩‍💼 Coordinadora", coords_analytics, key="analytics_coord")
        
        with col_filt2:
            # Filtro de Programa (en cascada, depende de coordinadora seleccionada)
            if sel_coord_analytics == "Todas":
                df_filt_coord = df_base
            else:
                df_filt_coord = df_base[df_base["COORDINADORA RESPONSABLE"] == sel_coord_analytics]
            
            progs_analytics = ["Todos"] + sorted_clean(df_filt_coord["PROGRAMA"]) if "PROGRAMA" in df_filt_coord.columns else ["Todos"]
            sel_prog_analytics = st.selectbox("🎓 Programa", progs_analytics, key="analytics_prog")
        
        with col_filt3:
            # Mostrar conteo de registros filtrados
            st.markdown("<br>", unsafe_allow_html=True)
            # Aplicar filtros
            df_analytics = df_base.copy()
            if sel_coord_analytics != "Todas":
                df_analytics = df_analytics[df_analytics["COORDINADORA RESPONSABLE"] == sel_coord_analytics]
            if sel_prog_analytics != "Todos":
                df_analytics = df_analytics[df_analytics["PROGRAMA"] == sel_prog_analytics]
            
            st.info(f"📊 {len(df_analytics):,} registros")
    
        st.markdown("---")
        
        # --- SECCIÓN 1: BÚSQUEDA GLOBAL ---
        # EXPLICACIÓN: Permite buscar cualquier texto en todo el dataset
        # y mostrar resultados filtrados instantáneamente.
        st.markdown("### 🔍 Búsqueda Global")
        st.caption("Busca programas, profesores, salas, coordinadoras o cualquier texto.")
        
        busqueda = st.text_input("", placeholder="Escribe para buscar...", key="busqueda_global", label_visibility="collapsed")
        
        if busqueda and len(busqueda) >= 2:
            # Buscar en todas las columnas de texto
            mask = pd.Series(False, index=df_analytics.index)
            for col in df_analytics.columns:
                if df_analytics[col].dtype == "object":
                    mask |= df_analytics[col].astype(str).str.contains(busqueda, case=False, na=False)
            
            resultados = df_analytics[mask]
            
            if not resultados.empty:
                st.success(f"✅ {len(resultados)} resultados encontrados")
                # Mostrar columnas relevantes
                cols_busq = ["PROGRAMA", "ASIGNATURA", "PROFESOR", "COORDINADORA RESPONSABLE", "DIAS/FECHAS", "SEDE"]
                cols_exist = [c for c in cols_busq if c in resultados.columns]
                df_busq = resultados[cols_exist].copy()
                if "DIAS/FECHAS" in df_busq.columns:
                    df_busq["DIAS/FECHAS"] = df_busq["DIAS/FECHAS"].dt.strftime("%d-%m-%Y")
                st.dataframe(df_busq.head(20), hide_index=True, use_container_width=True)
            else:
                st.info(f"No se encontraron resultados for '{busqueda}'")
        
        st.markdown("---")
        
        # --- SECCIÓN 2: TASA DE CUMPLIMIENTO ---
        # EXPLICACIÓN: Calcula el porcentaje de clases que ya se realizaron
        # comparado con el total planificado. Útil para medir avance.
        st.markdown("### 📊 Tasa de Cumplimiento")
        st.caption("Porcentaje de sesiones realizadas respecto al total planificado.")
    
        if "DIAS/FECHAS" in df_analytics.columns:
            # Filtro por año
            year_cumpl = st.selectbox("Año", sorted(df_analytics["DIAS/FECHAS"].dt.year.unique()), key="cumpl_year")
            df_cumpl = df_analytics[df_analytics["DIAS/FECHAS"].dt.year == year_cumpl].copy()
            
            # Calcular sesiones pasadas vs futuras (hoy es pd.Timestamp global)
            sesiones_pasadas = len(df_cumpl[df_cumpl["DIAS/FECHAS"] <= hoy])
            sesiones_futuras = len(df_cumpl[df_cumpl["DIAS/FECHAS"] > hoy])
            total_sesiones = len(df_cumpl)
            
            # Tasa de cumplimiento (% de sesiones ya realizadas)
            tasa = (sesiones_pasadas / total_sesiones * 100) if total_sesiones > 0 else 0
            
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("📅 Sesiones Planificadas", f"{total_sesiones:,}")
            col_m2.metric("✅ Sesiones Realizadas", f"{sesiones_pasadas:,}")
            col_m3.metric("⏳ Sesiones Pendientes", f"{sesiones_futuras:,}")
            col_m4.metric("📈 Tasa Cumplimiento", f"{tasa:.1f}%")
            
            # Gráfico de progreso por coordinadora
            if "COORDINADORA RESPONSABLE" in df_cumpl.columns:
                _hoy = hoy  # captura para lambda
                cumpl_coord = df_cumpl.groupby("COORDINADORA RESPONSABLE", as_index=False).agg(
                    Realizadas=("DIAS/FECHAS", lambda x: (x <= _hoy).sum()),
                    Pendientes=("DIAS/FECHAS", lambda x: (x > _hoy).sum()),
                    Total=("DIAS/FECHAS", "count")
                )
                cumpl_coord["% Cumplido"] = (cumpl_coord["Realizadas"] / cumpl_coord["Total"] * 100).round(1)
                
                # Altura dinámica basada en el número de coordinadoras (mínimo 300px)
                altura_grafico = max(300, len(cumpl_coord) * 40 + 100)
                
                fig_cumpl = px.bar(
                    cumpl_coord.sort_values("% Cumplido", ascending=True),
                    x="% Cumplido", y="COORDINADORA RESPONSABLE",
                    orientation="h", text="% Cumplido",
                    title="Tasa de Cumplimiento por Coordinadora",
                    color="% Cumplido",
                    color_continuous_scale=["#ef4444", "#fbbf24", "#22c55e"]
                )
                fig_cumpl.update_layout(
                    height=altura_grafico, 
                    showlegend=False,
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#e2e8f0"),
                    margin=dict(l=10, r=10, t=40, b=10), # Márgenes ajustados
                    yaxis=dict(title=None), # Quitar título del eje Y para ganar espacio
                    xaxis=dict(title="% Cumplido")
                )
                fig_cumpl.update_traces(
                    texttemplate="%{text:.1f}%", 
                    textposition="outside",
                    cliponaxis=False # Permitir que el texto sobresalga si es necesario
                )
                st.plotly_chart(fig_cumpl, use_container_width=True)
    
        st.markdown("---")
        
        # --- SECCIÓN 3: COMPARATIVA HISTÓRICA ---
        # EXPLICACIÓN: Compara las métricas del año actual con el año anterior
        # para identificar tendencias y cambios significativos.
        st.markdown("### 📉 Comparativa Histórica (Año vs Año)")
        st.caption("Compara el rendimiento del año actual con el anterior.")
        
        if "DIAS/FECHAS" in df_analytics.columns:
            años_disp = sorted(df_analytics["DIAS/FECHAS"].dt.year.unique(), reverse=True)
            
            if len(años_disp) >= 2:
                col_y1, col_y2 = st.columns(2)
                with col_y1:
                    año_actual = st.selectbox("Año Actual", años_disp, key="hist_actual")
                with col_y2:
                    años_ant = [a for a in años_disp if a < año_actual]
                    año_anterior = st.selectbox("Año Anterior", años_ant if años_ant else años_disp, key="hist_anterior")
                
                df_act = df_analytics[df_analytics["DIAS/FECHAS"].dt.year == año_actual].copy()
                df_ant = df_analytics[df_analytics["DIAS/FECHAS"].dt.year == año_anterior].copy()
                
                # Métricas comparativas
                col_h1, col_h2, col_h3, col_h4 = st.columns(4)
                
                # Sesiones
                ses_act, ses_ant = len(df_act), len(df_ant)
                delta_ses = ((ses_act - ses_ant) / ses_ant * 100) if ses_ant > 0 else 0
                col_h1.metric("📚 Sesiones", f"{ses_act:,}", f"{delta_ses:+.1f}%")
                
                # Programas
                prog_act = df_act["PROGRAMA"].nunique() if "PROGRAMA" in df_act.columns else 0
                prog_ant = df_ant["PROGRAMA"].nunique() if "PROGRAMA" in df_ant.columns else 0
                delta_prog = ((prog_act - prog_ant) / prog_ant * 100) if prog_ant > 0 else 0
                col_h2.metric("🎓 Programas", prog_act, f"{delta_prog:+.1f}%")
                
                # Profesores
                prof_act = df_act["PROFESOR"].nunique() if "PROFESOR" in df_act.columns else 0
                prof_ant = df_ant["PROFESOR"].nunique() if "PROFESOR" in df_ant.columns else 0
                delta_prof = ((prof_act - prof_ant) / prof_ant * 100) if prof_ant > 0 else 0
                col_h3.metric("👨‍🏫 Profesores", prof_act, f"{delta_prof:+.1f}%")
                
                # Días activos
                dias_act = df_act["DIAS/FECHAS"].dt.date.nunique()
                dias_ant = df_ant["DIAS/FECHAS"].dt.date.nunique()
                delta_dias = ((dias_act - dias_ant) / dias_ant * 100) if dias_ant > 0 else 0
                col_h4.metric("📅 Días Activos", dias_act, f"{delta_dias:+.1f}%")
                
                # Gráfico comparativo mensual
                df_act["Mes"] = df_act["DIAS/FECHAS"].dt.month
                df_ant["Mes"] = df_ant["DIAS/FECHAS"].dt.month
                
                comp_act = df_act.groupby("Mes").size().reset_index(name="Sesiones")
                comp_act["Año"] = str(año_actual)
                comp_ant = df_ant.groupby("Mes").size().reset_index(name="Sesiones")
                comp_ant["Año"] = str(año_anterior)
                
                comp_total = pd.concat([comp_act, comp_ant])
                comp_total["Mes_Nombre"] = comp_total["Mes"].map(mapa_meses)
                comp_total = comp_total.sort_values("Mes")
                
                fig_comp = px.line(
                    comp_total, x="Mes_Nombre", y="Sesiones", color="Año",
                    markers=True, title=f"Evolución Mensual: {año_actual} vs {año_anterior}"
                )
                fig_comp.update_layout(
                    height=350,
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#e2e8f0"),
                    xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor="#334155")
                )
                st.plotly_chart(fig_comp, use_container_width=True)
        else:
            st.info("Se necesitan al menos 2 años de datos para la comparativa.")

        # --- SECCIÓN 4: CARGA MENSUAL DETALLADA (PUNTAJE) ---
        st.markdown("---")
        st.markdown("### 🌡️ Análisis de Carga Laboral Mensual")
        st.caption("Cálculo basado en Puntaje: (Sesiones/4) × Factor Alumnos.")
        
        if not df_analytics.empty and "DIAS/FECHAS" in df_analytics.columns:
            df_points = df_analytics.copy()
            df_points["Mes_N"] = df_points["DIAS/FECHAS"].dt.month
            df_points["Año_N"] = df_points["DIAS/FECHAS"].dt.year
            
            # Buscar columna de alumnos
            col_alum_p = "Nº ALUMNOS"
            if col_alum_p not in df_points.columns:
                col_alum_p = next((c for c in df_points.columns if "ALUMNO" in c.upper()), None)
            
            if col_alum_p:
                df_points[col_alum_p] = pd.to_numeric(df_points[col_alum_p], errors="coerce").fillna(0)
            else:
                df_points["Nº ALUMNOS"] = 0
                col_alum_p = "Nº ALUMNOS"
            
            # Puntaje por Programa/Mes
            carga_mes = df_points.groupby(["Año_N", "Mes_N", "COORDINADORA RESPONSABLE", "PROGRAMA"]).agg(
                Sesiones=("DIAS/FECHAS", "count"),
                Alumnos=(col_alum_p, "max")
            ).reset_index()
            
            def _get_f(n):
                if n == 0: return 1.0
                if n < 20: return 1.0
                if n < 30: return 1.2
                if n < 40: return 1.4
                if n < 49: return 1.7
                return 2.0
            
            carga_mes["Puntaje"] = (carga_mes["Sesiones"]/4) * carga_mes["Alumnos"].apply(_get_f)
            
            # Resumen por Mes
            res_mensual = carga_mes.groupby(["Año_N", "Mes_N", "COORDINADORA RESPONSABLE"])["Puntaje"].sum().reset_index()
            res_mensual["Mes_Nombre"] = res_mensual["Mes_N"].map(mapa_meses)
            
            años_p = sorted(res_mensual["Año_N"].unique(), reverse=True)
            sel_y_p = st.selectbox("Año para Análisis de Carga", años_p, key="analytics_p_year")
            
            df_p = res_mensual[res_mensual["Año_N"] == sel_y_p].copy()
            
            if not df_p.empty:
                # Gráfico Barras
                fig_p = px.bar(
                    df_p, x="Mes_Nombre", y="Puntaje", color="COORDINADORA RESPONSABLE",
                    category_orders={"Mes_Nombre": list(mapa_meses.values())},
                    title=f"Evolución Mensual de Carga (Puntos) - Año {sel_y_p}",
                    barmode="group", text_auto=".1f"
                )
                fig_p.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#e2e8f0"), height=380,
                    margin=dict(l=10, r=10, t=50, b=10),
                    legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center")
                )
                st.plotly_chart(fig_p, use_container_width=True)
                
                # Heatmap Tabla
                st.markdown("#### 🧶 Mapa de Calor: Intensidad de Carga")
                pivot_p = df_p.pivot_table(
                    index="COORDINADORA RESPONSABLE",
                    columns="Mes_Nombre",
                    values="Puntaje",
                    aggfunc="sum",
                    fill_value=0
                )
                ord_cols = [m for m in mapa_meses.values() if m in pivot_p.columns]
                pivot_p = pivot_p[ord_cols]
                
                st.dataframe(
                    pivot_p.style.background_gradient(cmap="OrRd", axis=None).format("{:.2f}"),
                    use_container_width=True
                )
            else:
                st.info("No se hallaron datos para este periodo.")

# =============================================================================
# TAB 1: COORDINADORAS
# =============================================================================
if tab1:
  with tab1:
    if df_base.empty or "DIAS/FECHAS" not in df_base.columns:
        st.warning("⚠️ Primero sube un archivo en el panel lateral.")
    else:
        # Filtros en cascada
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            years_disp = sorted(df_base["DIAS/FECHAS"].dt.year.unique())
            sel_year = st.multiselect("1. Año", years_disp, key="t1_year", placeholder="Todos")
        df_1 = df_base[df_base["DIAS/FECHAS"].dt.year.isin(sel_year)] if sel_year else df_base

        with c2:
            if "Mes" in df_1.columns:
                meses_disp = sorted(df_1["Mes"].unique(),
                    key=lambda x: list(utils.MESES_NOMBRE.values()).index(x) if x in utils.MESES_NOMBRE.values() else 99)
                sel_mes = st.multiselect("2. Mes", meses_disp, key="t1_mes", placeholder="Todos")
                df_2 = df_1[df_1["Mes"].isin(sel_mes)] if sel_mes else df_1
            else:
                df_2 = df_1

        with c3:
            sedes_disp = sorted(df_2["SEDE"].dropna().astype(str).unique()) if "SEDE" in df_2.columns else []
            sel_sede = st.multiselect("3. Sede", sedes_disp, key="t1_sede", placeholder="Todas")
        df_3 = df_2[df_2["SEDE"].isin(sel_sede)] if sel_sede and "SEDE" in df_2.columns else df_2

        with c4:
            mods_disp = sorted(df_3["Modalidad_Calc"].dropna().astype(str).unique()) if "Modalidad_Calc" in df_3.columns else []
            sel_mod = st.multiselect("4. Modalidad", mods_disp, key="t1_mod", placeholder="Todas")
        df_4 = df_3[df_3["Modalidad_Calc"].isin(sel_mod)] if sel_mod and "Modalidad_Calc" in df_3.columns else df_3

        c5, c6, c7, c8 = st.columns(4)
        with c5:
            coords_disp = sorted(df_4["COORDINADORA RESPONSABLE"].dropna().astype(str).unique()) if "COORDINADORA RESPONSABLE" in df_4.columns else []
            sel_coord = st.multiselect("5. Coordinadora", coords_disp, key="t1_coord", placeholder="Todas")
        df_5 = df_4[df_4["COORDINADORA RESPONSABLE"].isin(sel_coord)] if sel_coord and "COORDINADORA RESPONSABLE" in df_4.columns else df_4

        with c6:
            progs_disp = sorted_clean(df_5["PROGRAMA"]) if "PROGRAMA" in df_5.columns else []
            sel_prog = st.multiselect("6. Programa", progs_disp, key="t1_prog", placeholder="Todos")
        df_6 = df_5[df_5["PROGRAMA"].isin(sel_prog)] if sel_prog and "PROGRAMA" in df_5.columns else df_5

        with c7:
            if "PROFESOR" in df_6.columns:
                profs_disp = sorted(df_6["PROFESOR"].dropna().astype(str).unique())
                sel_prof = st.multiselect("7. Profesor", profs_disp, key="t1_prof", placeholder="Todos")
                df_7 = df_6[df_6["PROFESOR"].isin(sel_prof)] if sel_prof else df_6
            else:
                df_7 = df_6

        with c8:
            dias_orden = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
            dias_disp = sorted(df_7["Dia_Semana"].unique(), key=lambda x: dias_orden.index(x) if x in dias_orden else 99) if "Dia_Semana" in df_7.columns else []
            sel_dia = st.multiselect("8. Día Semana", dias_disp, key="t1_dia", placeholder="Todos")
        df_final_t1 = df_7[df_7["Dia_Semana"].isin(sel_dia)] if sel_dia and "Dia_Semana" in df_7.columns else df_7

        st.markdown("---")

        if df_final_t1.empty:
            st.warning("⚠️ No se encontraron clases con esta combinación de filtros.")
        else:
            # KPIs
            total_sesiones = len(df_final_t1)
            total_progs = df_final_t1["PROGRAMA"].nunique() if "PROGRAMA" in df_final_t1.columns else 0

            carga_diaria = df_final_t1.groupby(["COORDINADORA RESPONSABLE","DIAS/FECHAS"]).agg(
                N_Progs=("PROGRAMA","nunique"),
                Programas=("PROGRAMA", lambda x: ", ".join(sorted_clean(x))),
                Dia=("Dia_Semana","first")
            ).reset_index()
            dias_criticos = carga_diaria[carga_diaria["N_Progs"] > 2]

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Sesiones", total_sesiones)
            k2.metric("Programas", total_progs)
            k3.metric("Días Activos", df_final_t1["DIAS/FECHAS"].nunique())
            k4.metric("Días Críticos (>2 Prog)", len(dias_criticos), delta_color="inverse")

            # Gráfico: Top 20 Días más intensos para evitar solapamiento ("información montada")
            df_plot = carga_diaria.copy()
            df_plot = df_plot.sort_values("N_Progs", ascending=False).head(20).sort_values("DIAS/FECHAS")
            df_plot["Fecha"] = df_plot["DIAS/FECHAS"].dt.strftime("%d-%m-%Y")
            color_by = "COORDINADORA RESPONSABLE" if df_plot["COORDINADORA RESPONSABLE"].nunique() > 1 else "N_Progs"

            fig_d = px.bar(df_plot, x="Fecha", y="N_Progs", color=color_by,
                           title="⚡ Intensidad de Programas por Día (Top 20 Días con más carga)",
                           labels={"N_Progs":"Cant. Programas"},
                           text="N_Progs")
            fig_d.add_hline(y=2, line_dash="dot", line_color="#ef4444", annotation_text="Límite Sostenible (2)")
            fig_d.update_traces(textposition='outside')
            st.plotly_chart(charts.update_chart_layout(fig_d), use_container_width=True)

            # Tabla Días Críticos
            st.markdown("##### 🚨 Detalle Días Críticos")
            if not dias_criticos.empty:
                dias_criticos = dias_criticos.copy()
                dias_criticos["Fecha"] = dias_criticos["DIAS/FECHAS"].dt.strftime("%d-%m-%Y")
                st.dataframe(
                    dias_criticos[["Fecha","Dia","COORDINADORA RESPONSABLE","N_Progs","Programas"]],
                    hide_index=True, use_container_width=True,
                    column_config={
                        "COORDINADORA RESPONSABLE": "Coordinadora",
                        "N_Progs": st.column_config.NumberColumn("Nº"),
                        "Programas": st.column_config.TextColumn("Programas", width="large"),
                    }
                )
            else:
                st.success("¡Excelente! No hay días con sobrecarga (>2 programas).")

            # Tabla Detalle
            st.markdown("---")
            st.subheader("📅 Calendario Detallado")
            cols_ver = ["DIAS/FECHAS","Dia_Semana","HORARIO","PROGRAMA","COORDINADORA RESPONSABLE","SEDE","Modalidad_Calc","ASIGNATURA"]
            cols_existentes = [c for c in cols_ver if c in df_final_t1.columns]
            df_show = df_final_t1[cols_existentes].copy()
            df_show["DIAS/FECHAS"] = df_show["DIAS/FECHAS"].dt.strftime("%d-%m-%Y")
            st.dataframe(df_show, hide_index=True, use_container_width=True)

# =============================================================================
# TAB 2: COMPARATIVA
# =============================================================================
if tab2:
  with tab2:
    df_t2 = None
    if df_base.empty or "DIAS/FECHAS" not in df_base.columns:
        st.warning("⚠️ Primero sube un archivo en el panel lateral.")
    else:
        st.markdown("### 📊 Comparativa de Carga y Gestión")
    
        # Filtros limpia visualmente
        with st.container():
            c2_1, c2_2, c2_3 = st.columns(3)
            sel_y2 = c2_1.multiselect("Año", sorted(df_base["DIAS/FECHAS"].dt.year.unique()), key="t2_y")
            sel_c2 = c2_2.multiselect("Coordinadoras", sorted_clean(df_base["COORDINADORA RESPONSABLE"]) if "COORDINADORA RESPONSABLE" in df_base.columns else [], key="t2_c")
            sel_m2 = c2_3.multiselect("Modalidad", sorted_clean(df_base["Modalidad_Calc"]) if "Modalidad_Calc" in df_base.columns else [], key="t2_m")

        mask2 = pd.Series(True, index=df_base.index)
        if sel_y2: mask2 &= df_base["DIAS/FECHAS"].dt.year.isin(sel_y2)
        if sel_c2: mask2 &= df_base["COORDINADORA RESPONSABLE"].isin(sel_c2)
        if sel_m2: mask2 &= df_base["Modalidad_Calc"].isin(sel_m2)
        df_t2 = df_base[mask2].copy()

    if df_t2 is None:
        pass
    elif df_t2.empty:
        st.warning("⚠️ No hay datos con los filtros seleccionados.")
    else:
        st.markdown("---")
        
        # Usamos Tabs internos para limpiar la interfaz y dar espacio a cada gráfico
        tab_ranking, tab_calor = st.tabs(["📊 Ranking de Carga", "🔥 Mapa de Calor Semanal"])
        
        # 1. TAB RANKING (Barra Horizontal)
        with tab_ranking:
            carga = df_t2["COORDINADORA RESPONSABLE"].value_counts().reset_index()
            carga.columns = ["Coordinadora","Sesiones"]
            carga = carga.sort_values("Sesiones", ascending=True)
            
            fig_c = px.bar(
                carga, 
                x="Sesiones", y="Coordinadora", 
                orientation='h',
                color="Sesiones",
                text="Sesiones",
                title="🏆 Ranking de Coordinadoras por Total de Sesiones",
                color_continuous_scale="Blues"
            )
            fig_c.update_layout(
                height=500, # Más alto para ver bien
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#e2e8f0"),
                margin=dict(l=0, r=0, t=40, b=0),
                coloraxis_showscale=False, # OCULTAR BARRA DE COLOR (reduce ruido)
                xaxis=dict(showgrid=False)
            )
            fig_c.update_traces(textposition='outside')
            st.plotly_chart(fig_c, use_container_width=True)

        # 2. TAB MAPA DE CALOR (Heatmap)
        with tab_calor:
            if "Dia_Semana" in df_t2.columns:
                df_heat_c = df_t2[df_t2["Dia_Semana"] != "Domingo"].copy()
                
                if not df_heat_c.empty:
                    # Abreviaturas para limpiar visualmente
                    dias_map = {
                        "Lunes": "Lun", "Martes": "Mar", "Miércoles": "Mié", 
                        "Jueves": "Jue", "Viernes": "Vie", "Sábado": "Sáb"
                    }
                    df_heat_c["Dia_Abrev"] = df_heat_c["Dia_Semana"].map(dias_map)
                    dias_ord = ["Lun","Mar","Mié","Jue","Vie","Sáb"]
                    
                    df_heat_c["Dia_Abrev"] = pd.Categorical(df_heat_c["Dia_Abrev"], categories=dias_ord, ordered=True)
                    
                    # Agrupar
                    heatmap_data = df_heat_c.groupby(["COORDINADORA RESPONSABLE", "Dia_Abrev"]).size().reset_index(name="Sesiones")
                    
                    # Pivotar
                    pivot = heatmap_data.pivot(index="COORDINADORA RESPONSABLE", columns="Dia_Abrev", values="Sesiones").fillna(0)
                    
                    # Graficar
                    fig_h = px.imshow(
                        pivot,
                        labels=dict(x="Día", y="Coordinadora", color="Sesiones"),
                        x=pivot.columns,
                        y=pivot.index,
                        color_continuous_scale="Viridis",
                        aspect="auto",
                        text_auto=True # Mostrar números dentro de los cuadros
                    )
                    
                    fig_h.update_layout(
                        height=500,
                        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                        font=dict(color="#e2e8f0"),
                        margin=dict(t=30, b=10, l=10, r=10),
                        xaxis=dict(side="top", title=None),
                        yaxis=dict(title=None)
                    )
                    st.plotly_chart(fig_h, use_container_width=True)
                else:
                    st.info("Sin datos para generar Heatmap.")

        st.markdown("### 📋 Resumen Detallado")
        df_res = resumen_coordinadoras_semana(df_t2)
        st.dataframe(
            df_res, 
            hide_index=True, 
            use_container_width=True,
            column_config={
                "Coordinadora": st.column_config.TextColumn("Coordinadora", width="medium"),
                "Días Activos (Semana)": st.column_config.ProgressColumn(
                    "Días Activos", 
                    format="%d", 
                    min_value=0, 
                )
            }
        )

# =============================================================================
# TAB GLOBAL: DASHBOARD (NUEVO)
# =============================================================================
if tab_global:
  with tab_global:
    if df_base.empty or "DIAS/FECHAS" not in df_base.columns:
        st.warning("⚠️ Primero sube un archivo en el panel lateral.")
    else:
        st.markdown("### 🌐 Visión Global de Programación")
        
        # Filtros Globales Limpios
        with st.container():
            col3_1, col3_2, col3_3 = st.columns([1, 1, 2])
            years_g = sorted(df_base["DIAS/FECHAS"].dt.year.dropna().unique())
            curr_year = pd.Timestamp.now().year
            def_year = [curr_year] if curr_year in years_g else []
            sel_y3 = col3_1.multiselect("Año", years_g, default=def_year, key="tg_y")
            
            meses_dict = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 
                          7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}
            curr_month_num = pd.Timestamp.now().month
            curr_month_name = meses_dict.get(curr_month_num)
            def_month = [curr_month_name] if curr_month_name else []
            sel_m3 = col3_2.multiselect("Mes", list(meses_dict.values()), default=def_month, key="tg_m", placeholder="Todos los meses")

            # Filtrar opciones de programa según año y mes seleccionados
            df_temp = df_base.copy()
            if sel_y3:
                df_temp = df_temp[df_temp["DIAS/FECHAS"].dt.year.isin(sel_y3)]
            if sel_m3:
                meses_inv = {v: k for k, v in meses_dict.items()}
                sel_m3_num = [meses_inv[m] for m in sel_m3]
                df_temp = df_temp[df_temp["DIAS/FECHAS"].dt.month.isin(sel_m3_num)]
                
            progs_active = sorted_clean(df_temp["PROGRAMA"]) if "PROGRAMA" in df_temp.columns else []
            sel_prog3 = col3_3.multiselect("Programa(s)", progs_active, key="tg_p", placeholder="Todos los programas activos")

        df_t3 = df_temp.copy()
        if sel_prog3 and "PROGRAMA" in df_t3.columns:
            df_t3 = df_t3[df_t3["PROGRAMA"].isin(sel_prog3)]

        # Preparar datos
        df_t3["Mes"] = df_t3["DIAS/FECHAS"].dt.to_period("M").astype(str)

        # 1. Gráfico de Evolución / Distribución Mensual
        st.markdown("#### 📅 Distribución Mensual por Programa")
        if "PROGRAMA" in df_t3.columns:
            df_t3["Programa_Abrev"] = df_t3["PROGRAMA"].apply(lambda x: utils.abbreviate_program_name(x, max_len=35))
            
            data_g = df_t3.groupby(["Mes", "Programa_Abrev"]).size().reset_index(name="Sesiones")

        # ── CORRECCIÓN: forzar TODOS los meses del año seleccionado ──────
        years_en_datos = sorted(df_t3["DIAS/FECHAS"].dt.year.dropna().unique().astype(int))
        all_months = []
        for yr in years_en_datos:
            for mo in range(1, 13):
                all_months.append(f"{yr}-{mo:02d}")
        meses_en_datos = sorted(df_t3["Mes"].unique())
        all_months_ordered = [m for m in all_months if m in meses_en_datos]

        data_g["Mes"] = pd.Categorical(data_g["Mes"], categories=all_months_ordered, ordered=True)
        data_g = data_g.sort_values("Mes")
        # ── FIN CORRECCIÓN ─────────────────────────────────────────────────

        fig_g = px.bar(
            data_g, x="Mes", y="Sesiones", color="Programa_Abrev",
            title="Distribución Mensual" + (" (Filtrada)" if sel_prog3 else " (Todos los Programas)"),
            category_orders={"Mes": all_months_ordered}
        )
        fig_g.update_layout(
            height=600,
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#e2e8f0"),
            legend=dict(
                orientation="h",
                yanchor="top", y=-0.2,
                xanchor="center", x=0.5,
                title=None
            ),
            margin=dict(b=150),
            xaxis=dict(
                tickangle=-45,
                type="category"   # fuerza eje categórico → muestra todos los meses
            )
        )
        st.plotly_chart(fig_g, use_container_width=True)
    
        # 2. Heatmap Global Mejorado: Carga por Coordinadora
        st.markdown("---")
        st.markdown("#### 🔥 Mapa de Calor: Carga Semanal por Coordinadora")
    
        if "COORDINADORA RESPONSABLE" in df_t3.columns and "Dia_Semana" in df_t3.columns:
            df_heat = df_t3.copy()
            
            # Filtrar Domingo
            df_heat = df_heat[df_heat["Dia_Semana"] != "Domingo"]

            if not df_heat.empty:
                # Abreviar días
                dias_map = {"Lunes":"Lun","Martes":"Mar","Miércoles":"Mié","Jueves":"Jue","Viernes":"Vie","Sábado":"Sáb"}
                df_heat["Dia_Abrev"] = df_heat["Dia_Semana"].map(dias_map)
                dias_ord = ["Lun","Mar","Mié","Jue","Vie","Sáb"]
                df_heat["Dia_Abrev"] = pd.Categorical(df_heat["Dia_Abrev"], categories=dias_ord, ordered=True)

                # Agrupar por Coordinadora y Día
                heatmap_data = df_heat.groupby(["COORDINADORA RESPONSABLE", "Dia_Abrev"]).size().reset_index(name="Sesiones")
                pivot = heatmap_data.pivot(index="COORDINADORA RESPONSABLE", columns="Dia_Abrev", values="Sesiones").fillna(0)
                
                # Graficar
                fig_heat = px.imshow(
                    pivot,
                    labels=dict(x="Día", y="Coordinadora", color="Sesiones"),
                    x=pivot.columns,
                    y=pivot.index,
                    color_continuous_scale="Viridis", # Escala clara y profesional
                    aspect="auto",
                    text_auto=True # Mostrar números
                )
                fig_heat.update_layout(
                    height=500, # Altura suficiente para todas las coordinadoras
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#e2e8f0"),
                    title="Intensidad de Trabajo por Coordinadora y Día",
                    xaxis=dict(side="top") # Días arriba
                )
                st.plotly_chart(fig_heat, use_container_width=True)
                
                # KPI Resumen
                top_coord = heatmap_data.groupby("COORDINADORA RESPONSABLE")["Sesiones"].sum().idxmax()
                top_day = heatmap_data.groupby("Dia_Abrev")["Sesiones"].sum().idxmax()
                st.caption(f"ℹ️ Coordinadora con más carga visualizada: **{top_coord}**. Día más ocupado: **{top_day}**.")

            else:
                st.info("No hay datos para generar el mapa de calor (excluyendo domingos).")
        else:
            st.info("Se requiere información de Coordinadora y Día de la Semana.")

# =============================================================================
# TAB 4: RESUMEN PROGRAMAS
# =============================================================================
if tab4:
  with tab4:
    with st.expander("🔍 Filtros", expanded=True):
        f1, f2, f3 = st.columns(3)
        s_y4 = f1.multiselect("Año", sorted(df_base["DIAS/FECHAS"].dt.year.unique()), key="t4_y")
        d4_1 = df_base[df_base["DIAS/FECHAS"].dt.year.isin(s_y4)] if s_y4 else df_base
        s_c4 = f2.multiselect("Coordinadora", sorted_clean(d4_1["COORDINADORA RESPONSABLE"]) if "COORDINADORA RESPONSABLE" in d4_1.columns else [], key="t4_c")
        d4_2 = d4_1[d4_1["COORDINADORA RESPONSABLE"].isin(s_c4)] if s_c4 else d4_1
        s_p4 = f3.multiselect("Programa", sorted_clean(d4_2["PROGRAMA"]) if "PROGRAMA" in d4_2.columns else [], key="t4_p")
    df_final_t4 = d4_2[d4_2["PROGRAMA"].isin(s_p4)] if s_p4 else d4_2

    if df_final_t4.empty:
        st.warning("No hay datos.")
    else:
        # hoy ya declarado globalmente
        stats = df_final_t4.groupby("PROGRAMA").agg(
            Inicio=("DIAS/FECHAS","min"),
            Fin=("DIAS/FECHAS","max"),
            Sesiones=("DIAS/FECHAS","count"),
            Coords=("COORDINADORA RESPONSABLE", lambda x: ", ".join(sorted_clean(x)))
        ).reset_index()

        def get_avance(r):
            if pd.isna(r["Inicio"]) or pd.isna(r["Fin"]): return 0
            total = (r["Fin"] - r["Inicio"]).days
            if total <= 0: return 100
            elapsed = (hoy - r["Inicio"]).days
            return 100 if elapsed >= total else max(0, int((elapsed/total)*100))

        stats["% Avance"] = stats.apply(get_avance, axis=1)

        st.dataframe(
            stats.sort_values("% Avance", ascending=False),
            column_config={
                "% Avance": st.column_config.ProgressColumn("Progreso", format="%d%%", min_value=0, max_value=100),
                "Inicio": st.column_config.DateColumn("Inicio", format="DD/MM/YYYY"),
                "Fin": st.column_config.DateColumn("Fin", format="DD/MM/YYYY"),
            },
            hide_index=True, use_container_width=True
        )

        # ── Botones de descarga ───────────────────────────────────────────────
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        # Preparar datos para exportar (convertir fechas a texto legible)
        _stats_exp = stats.sort_values("% Avance", ascending=False).copy()
        for _col in ["Inicio", "Fin"]:
            if _col in _stats_exp.columns:
                _stats_exp[_col] = pd.to_datetime(
                    _stats_exp[_col], errors="coerce"
                ).dt.strftime("%d/%m/%Y")

        _detalle_exp = df_final_t4.copy()
        for _col in _detalle_exp.select_dtypes(
            include=["datetime64[ns]", "datetimetz"]
        ).columns:
            _detalle_exp[_col] = _detalle_exp[_col].dt.strftime("%d/%m/%Y")

        dl1, dl2, dl3 = st.columns(3)

        # 1️⃣ EXCEL — guardar en ~/Downloads y abrir directamente en Excel
        # (la app corre en local, por eso podemos escribir al disco del usuario)
        with dl1:
            if st.button("📊 Guardar Excel y Abrir", use_container_width=True,
                         type="primary", key="btn_excel_abrir_t4"):
                try:
                    _downloads = Path.home() / "Downloads"
                    _downloads.mkdir(exist_ok=True)
                    _fname_xl = f"Resumen_Programas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    _ruta_xl  = _downloads / _fname_xl

                    with pd.ExcelWriter(str(_ruta_xl), engine="openpyxl") as _writer:
                        _stats_exp.to_excel(_writer, index=False, sheet_name="Resumen")
                        _detalle_exp.to_excel(_writer, index=False, sheet_name="Detalle")

                    os.system(f'open "{_ruta_xl}"')
                    st.success(f"✅ Guardado en:\n`~/Downloads/{_fname_xl}`\n\nAbriendo en Excel...")
                except Exception as _e_xl:
                    st.error(f"❌ Error: {_e_xl}")

        # 2️⃣ CSV Resumen (respaldo de descarga por navegador)
        with dl2:
            render_static_download(
                "📄 CSV Resumen",
                _stats_exp.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                f"Resumen_Programas_{datetime.now().strftime('%Y%m%d')}.csv"
            )

        # 3️⃣ Descarga Calendario ICS
        with dl3:
            def _generar_ics(df_eventos: pd.DataFrame) -> str:
                """Genera un string en formato iCalendar (.ics) a partir del dataframe."""
                lines = [
                    "BEGIN:VCALENDAR",
                    "VERSION:2.0",
                    "PRODID:-//Calendario Postgrado UAI//ES",
                    "CALSCALE:GREGORIAN",
                    "METHOD:PUBLISH",
                ]
                for _, row in df_eventos.iterrows():
                    try:
                        fecha = pd.to_datetime(row["DIAS/FECHAS"])
                        fecha_str = fecha.strftime("%Y%m%d")
                        programa  = str(row.get("PROGRAMA",  "Sesión"))[:60]
                        coord     = str(row.get("COORDINADORA RESPONSABLE", ""))[:40]
                        sede      = str(row.get("SEDE", ""))[:40]
                        horario   = str(row.get("HORARIO", ""))[:20]
                        uid       = f"{fecha_str}-{abs(hash(programa+coord))}@uai.cl"
                        lines += [
                            "BEGIN:VEVENT",
                            f"UID:{uid}",
                            f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
                            f"DTSTART;VALUE=DATE:{fecha_str}",
                            f"DTEND;VALUE=DATE:{fecha_str}",
                            f"SUMMARY:{programa}",
                            f"DESCRIPTION:Coordinadora: {coord}\\nHorario: {horario}\\nSede: {sede}",
                            f"LOCATION:{sede}",
                            "END:VEVENT",
                        ]
                    except Exception:
                        continue
                lines.append("END:VCALENDAR")
                return "\r\n".join(lines)

            _ics_str = _generar_ics(df_final_t4)
            render_static_download(
                "📅 Calendario (.ics)",
                _ics_str.encode("utf-8"),
                f"Calendario_Programas_{datetime.now().strftime('%Y%m%d')}.ics"
            )

# =============================================================================
# TAB 5: CALIDAD & SEDE
# =============================================================================
if tab5:
  with tab5:
    f5_1, f5_2 = st.columns(2)
    sy5 = f5_1.multiselect("Año", sorted(df_base["DIAS/FECHAS"].dt.year.unique()), key="t5_y")
    df_pre_year = df_base[df_base["DIAS/FECHAS"].dt.year.isin(sy5)] if sy5 else df_base
    sc5 = f5_2.multiselect("Coordinadora", sorted(df_pre_year["COORDINADORA RESPONSABLE"].dropna().unique()) if "COORDINADORA RESPONSABLE" in df_pre_year.columns else [], key="t5_c")
    df_t5 = df_pre_year[df_pre_year["COORDINADORA RESPONSABLE"].isin(sc5)] if sc5 else df_pre_year

    if df_t5.empty:
        st.warning("No hay datos.")
    else:
        cm, cs = st.columns(2)
        with cm:
            st.markdown("### Modalidad")
            df_m = resumen_modalidad(df_t5)
            if not df_m.empty:
                fig_m = px.pie(df_m, names="Modalidad_Calc", values="Sesiones", hole=0.4)
                st.plotly_chart(charts.update_chart_layout(fig_m), use_container_width=True)
                st.dataframe(df_m, hide_index=True, use_container_width=True)

        with cs:
            st.markdown("### Sede")
            df_s = resumen_sede(df_t5)
            if not df_s.empty:
                fig_s = px.bar(df_s, x="SEDE", y="Sesiones", color="Sesiones")
                st.plotly_chart(charts.update_chart_layout(fig_s), use_container_width=True)
                st.dataframe(df_s, hide_index=True, use_container_width=True)

        st.markdown("---")
        st.markdown("### 🕸️ Análisis de Tendencias (Radar)")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            top_n_sedes = df_t5["SEDE"].value_counts().head(8).index.tolist()
            df_radar_sede = df_t5[df_t5["SEDE"].isin(top_n_sedes)].groupby("SEDE").size().reset_index(name="Sesiones")
            if not df_radar_sede.empty:
                fig_r1 = px.line_polar(df_radar_sede, r="Sesiones", theta="SEDE",
                                       line_close=True, markers=True, title="Top Sedes (Volumen)")
                fig_r1.update_traces(fill='toself', line_color='#2dd4bf')
                st.plotly_chart(charts.update_chart_layout(fig_r1), use_container_width=True)
        with col_r2:
            df_radar_mod = df_t5.groupby("Modalidad_Calc").size().reset_index(name="Sesiones")
            if not df_radar_mod.empty:
                fig_r2 = px.line_polar(df_radar_mod, r="Sesiones", theta="Modalidad_Calc",
                                       line_close=True, markers=True, title="Distribución por Modalidad")
                fig_r2.update_traces(fill='toself', line_color='#38bdf8')
                st.plotly_chart(charts.update_chart_layout(fig_r2), use_container_width=True)

        st.markdown("---")
        st.markdown("### 🧹 Auditoría de Datos (Valores Faltantes)")
        df_q = resumen_calidad_datos(df_base)
        st.dataframe(df_q, hide_index=True, use_container_width=True)

# =============================================================================
# TAB CALENDARIO
# =============================================================================
if tab_calendario:
  with tab_calendario:
    col_c1, col_c2, col_c3, col_c4 = st.columns(4)

    current_year = datetime.now().year
    current_month = datetime.now().month
    years_avail = sorted(df_base["DIAS/FECHAS"].dt.year.dropna().unique().astype(int).tolist())
    if not years_avail:
        years_avail = [current_year]

    year_options = ["Todos"] + [str(y) for y in years_avail]
    sel_cal_year_txt = col_c1.selectbox(
        "Año",
        year_options,
        index=year_options.index(str(current_year)) if str(current_year) in year_options else 0,
        key="cal_year"
    )
    sel_cal_year = None if sel_cal_year_txt == "Todos" else int(sel_cal_year_txt)

    meses_nombres = ["Todos","Enero","Febrero","Marzo","Abril","Mayo","Junio",
                     "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
    sel_cal_month_txt = col_c2.selectbox(
        "Mes", meses_nombres,
        index=current_month,  # Enero=1 → índice 1 en lista con "Todos" al inicio
        key="cal_month"
    )
    sel_cal_month = None if sel_cal_month_txt == "Todos" else meses_nombres.index(sel_cal_month_txt)

    # Máscara de tiempo base (tolerante a "Todos")
    mask_time = pd.Series(True, index=df_base.index)
    if sel_cal_year is not None:
        mask_time &= df_base["DIAS/FECHAS"].dt.year == sel_cal_year
    if sel_cal_month is not None:
        mask_time &= df_base["DIAS/FECHAS"].dt.month == sel_cal_month

    coords_cal = sorted_clean(df_base[mask_time]["COORDINADORA RESPONSABLE"]) if "COORDINADORA RESPONSABLE" in df_base.columns else []
    sel_cal_coord = col_c3.multiselect("Coordinadora", coords_cal, key="cal_coord")

    progs_avail = sorted(
        df_base[df_base["COORDINADORA RESPONSABLE"].isin(sel_cal_coord)]["PROGRAMA"].dropna().unique().tolist()
    ) if sel_cal_coord else (sorted_clean(df_base["PROGRAMA"]) if "PROGRAMA" in df_base.columns else [])
    sel_cal_prog = col_c4.multiselect("Programa", progs_avail, key="cal_prog")

    # Filtrar datos del calendario
    mask_cal = mask_time.copy()
    if sel_cal_coord:
        mask_cal &= df_base["COORDINADORA RESPONSABLE"].isin(sel_cal_coord)
    if sel_cal_prog:
        mask_cal &= df_base["PROGRAMA"].isin(sel_cal_prog)
    df_cal = df_base[mask_cal].copy()

    # Si "Todos" meses/años → usar el mes/año actual para renderizar la grilla
    render_year  = sel_cal_year  if sel_cal_year  is not None else current_year
    render_month = sel_cal_month if sel_cal_month is not None else current_month

    # Eventos por día
    events_map = {}
    if not df_cal.empty:
        df_cal["day_temp"] = df_cal["DIAS/FECHAS"].dt.day
        cols_to_map = ["PROGRAMA","COORDINADORA RESPONSABLE"]
        for opt in ["ASIGNATURA","HORARIO","PROFESOR","SEDE","Modalidad_Calc","SALA"]:
            if opt in df_cal.columns: cols_to_map.append(opt)
        for d, group in df_cal.groupby("day_temp"):
            events_map[d] = group[cols_to_map].to_dict("records")

    # Abreviatura inteligente (using utility)
    def smart_abbr(txt):
        return utils.abbreviate_program_name(txt, max_len=40)

    # Los estilos ahora se manejan centralizadamente en styles.py

    cal_obj = calendar.Calendar(firstweekday=0)
    month_days = cal_obj.monthdayscalendar(render_year, render_month)
    now = datetime.now()

    _NULO = ("nan", "none", "s/d", "-", "por definir", "")

    def _clean(val):
        s = str(val).strip()
        return "" if s.lower() in _NULO else s

    html = '<div class="apple-cal">'
    html += '<div class="apple-grid-header">'
    for d in ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]:
        html += f'<div class="apple-weekday">{d}</div>'
    html += '</div>'
    html += '<div class="apple-grid-body">'

    for week in month_days:
        for day in week:
            if day == 0:
                html += '<div class="apple-cell" style="background-color:#0a0f1e;"></div>'
                continue

            is_today = (render_year == now.year and render_month == now.month and day == now.day)
            cell_cls = "apple-cell is-today" if is_today else "apple-cell"

            html += f'<div class="{cell_cls}">'
            if is_today:
                html += f'<div class="apple-day-num"><span>{day}</span></div>'
            else:
                html += f'<div class="apple-day-num">{day}</div>'

            day_events = events_map.get(day, [])
            html += '<div class="apple-events">'

            for i, evt in enumerate(day_events):
                if i >= 4:
                    rem = len(day_events) - i
                    html += f'<div style="font-size:9px;text-align:center;color:#8e8e93;">+{rem} más...</div>'
                    break

                prog    = smart_abbr(_clean(evt.get("PROGRAMA", "")))
                asig    = smart_abbr(_clean(evt.get("ASIGNATURA", "")))
                horario = _clean(evt.get("HORARIO", ""))
                prof    = _clean(evt.get("PROFESOR", ""))
                prof    = prof.title()[:22] if prof else ""
                sala    = _clean(evt.get("SALA", ""))
                sede    = _clean(evt.get("SEDE", ""))

                # Construir línea de metadatos solo con valores reales
                meta_parts = []
                if prof: meta_parts.append(f"👤 {prof}")
                if sala: meta_parts.append(f"📍 {sala}")
                if sede: meta_parts.append(f"🏢 {sede}")
                meta_txt = "  ".join(meta_parts)

                html += f'''<div class="event-card">
                    <div class="evt-time">{horario}</div>
                    <div class="evt-prog">{prog}</div>
                    <div class="evt-subj">{asig}</div>
                    <div class="evt-meta">{meta_txt}</div>
                </div>'''

            html += '</div></div>'

    html += '</div></div>'
    st.markdown(html, unsafe_allow_html=True)

    # Tabla detalle
    st.markdown("---")
    st.subheader(f"Agenda: {sel_cal_month_txt} {sel_cal_year_txt}")
    if df_cal.empty:
        st.info("No hay clases programadas.")
    else:
        df_view = df_cal.copy()
        df_view["Fecha"] = df_view["DIAS/FECHAS"].dt.strftime("%d-%m-%Y")
        cols_base = ["Fecha","HORARIO","PROGRAMA","ASIGNATURA","PROFESOR","COORDINADORA RESPONSABLE","SEDE","SALA"]
        cols_show = [c for c in cols_base if c in df_view.columns]
        st.dataframe(df_view[cols_show].sort_values(["Fecha","HORARIO"]), hide_index=True, use_container_width=True)
        
        # --- EXPORTACIÓN iCAL ---
        st.markdown("---")
        st.markdown("### 📤 Exportar a Calendario")
        st.caption("Descarga el archivo .ics para importar a Outlook, Google Calendar o Apple Calendar.")
        
        def generate_ical(df_export):
            """Genera archivo iCal (.ics) desde el DataFrame"""
            ical = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Calendario Postgrado//ES", "CALSCALE:GREGORIAN"]
            
            for _, row in df_export.iterrows():
                try:
                    fecha = row["DIAS/FECHAS"]
                    horario = str(row.get("HORARIO", "09:00-10:00"))
                    programa = str(row.get("PROGRAMA", "Clase"))
                    asignatura = str(row.get("ASIGNATURA", ""))
                    profesor = str(row.get("PROFESOR", ""))
                    sede = str(row.get("SEDE", ""))
                    sala = str(row.get("SALA", ""))
                    
                    # Parsear horario
                    if "-" in horario:
                        h_ini, h_fin = horario.split("-")
                        h_ini = h_ini.strip().replace(":", "")
                        h_fin = h_fin.strip().replace(":", "")
                    else:
                        h_ini, h_fin = "0900", "1000"
                    
                    # Formatear fecha
                    fecha_str = fecha.strftime("%Y%m%d")
                    
                    uid = f"{fecha_str}-{h_ini}-{programa[:10]}@postgrado"
                    
                    ical.append("BEGIN:VEVENT")
                    ical.append(f"UID:{uid}")
                    ical.append(f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}")
                    ical.append(f"DTSTART:{fecha_str}T{h_ini}00")
                    ical.append(f"DTEND:{fecha_str}T{h_fin}00")
                    ical.append(f"SUMMARY:{programa} - {asignatura}")
                    ical.append(f"DESCRIPTION:Profesor: {profesor}\\nCoordinadora: {row.get('COORDINADORA RESPONSABLE', '')}")
                    ical.append(f"LOCATION:{sede} {sala}")
                    ical.append("END:VEVENT")
                except Exception:
                    continue
            
            ical.append("END:VCALENDAR")
            return "\n".join(ical)
        
        col_ical1, col_ical2 = st.columns(2)
        with col_ical1:
            ical_content = generate_ical(df_cal)
            render_static_download(
                "📅 Descargar iCal (.ics)",
                ical_content.encode('utf-8'),
                f"calendario_{sel_cal_month_txt}_{sel_cal_year}.ics"
            )
        with col_ical2:
            render_excel_download(
                "📊 Descargar Excel (.xlsx)",
                df_view[cols_show],
                f"agenda_{sel_cal_month_txt}_{sel_cal_year}.xlsx"
            )

# =============================================================================
# TAB RESERVAS
# =============================================================================
if tab_reservas:
  with tab_reservas:
    if df_reservas.empty:
        st.info("ℹ️ Carga un archivo de reservas en el panel lateral para visualizar este calendario.")
    else:
        # ── Filtros ──────────────────────────────────────────────────────────
        res_col1, res_col2, res_col3 = st.columns(3)

        current_year  = datetime.now().year
        current_month = datetime.now().month

        if "FECHA" in df_reservas.columns:
            years_res = sorted(df_reservas["FECHA"].dropna().dt.year.unique())
        else:
            years_res = [current_year]

        sel_res_year = res_col1.selectbox(
            "Año", years_res,
            index=years_res.index(current_year) if current_year in years_res else 0,
            key="res_year"
        )

        meses_nombres_res = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                             "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
        sel_res_month_txt = res_col2.selectbox("Mes", meses_nombres_res, index=current_month - 1, key="res_month")
        sel_res_month = meses_nombres_res.index(sel_res_month_txt) + 1

        salas_res = []
        if "SALA" in df_reservas.columns:
            salas_res = sorted(df_reservas["SALA"].dropna().astype(str).unique())
        sel_res_sala = res_col3.multiselect("Filtrar Sala", salas_res, key="res_sala", placeholder="Todas")

        # ── Filtrar datos ─────────────────────────────────────────────────────
        if "FECHA" in df_reservas.columns:
            mask_res = (
                (df_reservas["FECHA"].dt.year  == sel_res_year) &
                (df_reservas["FECHA"].dt.month == sel_res_month)
            )
            if sel_res_sala:
                mask_res &= df_reservas["SALA"].astype(str).isin(sel_res_sala)
            df_res_cal = df_reservas[mask_res].copy()
        else:
            df_res_cal = df_reservas.copy()

        # ── Mapa de eventos por día ───────────────────────────────────────────
        res_events_map = {}
        if not df_res_cal.empty and "FECHA" in df_res_cal.columns:
            df_res_cal["day_temp"] = df_res_cal["FECHA"].dt.day
            res_cols_map = []
            for col in ["EVENTO","HORA_INICIO","HORA_FIN","SALA","SOLICITANTE"]:
                if col in df_res_cal.columns:
                    res_cols_map.append(col)
            for d, grp in df_res_cal.groupby("day_temp"):
                res_events_map[d] = grp[res_cols_map].to_dict("records")

        # Los estilos ahora se manejan en styles.py

        # ── Build HTML calendario ─────────────────────────────────────────────
        cal_obj_res = calendar.Calendar(firstweekday=0)
        month_days_res = cal_obj_res.monthdayscalendar(sel_res_year, sel_res_month)
        now_res = datetime.now()

        html_res = '<div class="res-cal">'
        # Cabecera mes
        html_res += f'''
        <div style="padding:16px 24px; background-color:#0f172a; border-bottom:1px solid #334155; display:flex; justify-content:space-between; align-items:baseline;">
            <div style="font-size:26px; font-weight:600; color:white;">{sel_res_month_txt} {sel_res_year}</div>
            <div style="font-size:13px; color:#94a3b8;">🏢 Reservas de Salas</div>
        </div>'''
        # Encabezado días
        html_res += '<div class="res-grid-header">'
        for d in ["LUN","MAR","MIÉ","JUE","VIE","SÁB","DOM"]:
            html_res += f'<div class="res-weekday">{d}</div>'
        html_res += '</div>'
        # Grid
        html_res += '<div class="res-grid-body">'

        for week in month_days_res:
            for day in week:
                if day == 0:
                    html_res += '<div class="res-cell" style="background-color:#0f172a;"></div>'
                    continue

                is_today_res = (sel_res_year == now_res.year and sel_res_month == now_res.month and day == now_res.day)
                cell_cls_res = "res-cell res-today" if is_today_res else "res-cell"

                html_res += f'<div class="{cell_cls_res}">'
                day_num = f'<span>{day}</span>' if is_today_res else str(day)
                html_res += f'<div class="res-day-num">{day_num}</div>'

                day_evts = res_events_map.get(day, [])
                html_res += '<div class="apple-events">'

                for i, evt in enumerate(day_evts):
                    if i >= 4:
                        rem = len(day_evts) - i
                        html_res += f'<div style="font-size:9px;text-align:center;color:#8e8e93;">+{rem} más...</div>'
                        break

                    hora_ini = str(evt.get("HORA_INICIO", ""))
                    hora_fin = str(evt.get("HORA_FIN", ""))
                    horario_str = f"{hora_ini} - {hora_fin}" if hora_ini and hora_fin and hora_ini not in ["nan",""] else hora_ini

                    evento_nm = str(evt.get("EVENTO", "Reserva"))
                    if evento_nm.lower() in ["nan","none",""]: evento_nm = "Reserva"

                    sala_nm = str(evt.get("SALA", ""))
                    if sala_nm.lower() in ["nan","none",""]: sala_nm = "Por definir"

                    solicit = str(evt.get("SOLICITANTE", ""))
                    if solicit.lower() in ["nan","none",""]: solicit = ""
                    # Acortar nombre solicitante
                    parts_s = solicit.split(" ")
                    solicit_nm = " ".join(parts_s[:2]).title() if len(parts_s) >= 2 else solicit.title()

                    html_res += f'''<div class="res-card">
                        <div class="res-time">{horario_str}</div>
                        <div class="res-evento">{evento_nm[:40]}</div>
                        <div class="res-sala">📍 {sala_nm}</div>
                        <div class="res-meta">👤 {solicit_nm}</div>
                    </div>'''

                html_res += '</div></div>'

        html_res += '</div></div>'
        st.markdown(html_res, unsafe_allow_html=True)

        # ── Tabla detalle ─────────────────────────────────────────────────────
        st.markdown("---")
        st.subheader(f"📝 Reservas: {sel_res_month_txt} {sel_res_year}")
        if df_res_cal.empty:
            st.info("No hay reservas en este período.")
        else:
            df_res_view = df_res_cal.copy()
            if "FECHA" in df_res_view.columns:
                df_res_view["Fecha"] = df_res_view["FECHA"].dt.strftime("%d-%m-%Y")
            cols_res_base = ["Fecha","HORA_INICIO","HORA_FIN","SALA","EVENTO","SOLICITANTE"]
            cols_res_show = [c for c in cols_res_base if c in df_res_view.columns]
            sort_cols = [c for c in ["Fecha","HORA_INICIO"] if c in df_res_view.columns]
            st.dataframe(
                df_res_view[cols_res_show].sort_values(sort_cols) if sort_cols else df_res_view[cols_res_show],
                hide_index=True, use_container_width=True
            )
            render_static_download(
                "📥 Descargar Reservas (CSV)",
                df_res_view[cols_res_show].to_csv(index=False).encode("utf-8"),
                f"reservas_{sel_res_month_txt}_{sel_res_year}.csv"
            )

# =============================================================================
# TAB GESTIÓN (COMPLETO)
# =============================================================================
if tab_gestion:
  with tab_gestion:
    password = st.text_input("Contraseña Administrador", type="password", key="gestion_pwd")

    # Contraseña desde secrets o variable de entorno (con fallback a 'admin' solo para desarrollo)
    try:
        _admin_pwd = st.secrets.get("admin_password", "admin")
    except Exception:
        _admin_pwd = "admin"
    if password == _admin_pwd:
        st.success("Acceso Concedido")

        # ── Alertas pendientes ────────────────────────────────────────────
        try:
            curr_y = datetime.now().year
            df_alert = df_base[df_base["DIAS/FECHAS"].dt.year == curr_y].copy()

            def get_full_end_dt_alert_g(row):
                d = row["DIAS/FECHAS"]
                t_str = str(row.get("HORARIO",""))
                t_obj = None
                if "-" in t_str:
                    try:
                        end_t = t_str.split("-")[1].strip()
                        for fmt in ("%H:%M","%H:%M:%S"):
                            try:
                                t_obj = datetime.strptime(end_t, fmt).time()
                                break
                            except: continue
                    except: pass
                if t_obj:
                    return datetime(d.year, d.month, d.day, t_obj.hour, t_obj.minute, t_obj.second)
                return datetime(d.year, d.month, d.day, 23, 59, 59)

            if not df_alert.empty and "ASIGNATURA" in df_alert.columns:
                df_alert["Full_End_DT"] = df_alert.apply(get_full_end_dt_alert_g, axis=1)
                grp_alert = df_alert.groupby(["PROGRAMA","ASIGNATURA"])["Full_End_DT"].max()
                map_alert = grp_alert.to_dict()
                pend_decls = []
                now_t = datetime.now()
                for (prog, cur), end_dt in map_alert.items():
                    if end_dt < now_t:
                        pend_decls.append({
                            "PROGRAMA": prog, "ASIGNATURA": cur,
                            "FECHA TÉRMINO": end_dt.strftime("%d-%m-%Y %H:%M")
                        })
                if pend_decls:
                    st.warning(f"⚠️ **Tienes {len(pend_decls)} cursos finalizados que requieren declaración.**")
                    with st.expander("Ver Listado de Pendientes", expanded=True):
                        st.dataframe(pd.DataFrame(pend_decls), hide_index=True, use_container_width=True)
        except Exception as e:
            st.error(f"Error calculando alertas: {e}")

        st.markdown("---")

        # ── Filtro año ────────────────────────────────────────────────────
        years_gestion = sorted(df_base["DIAS/FECHAS"].dt.year.unique())
        sel_year_g = st.multiselect("Filtrar por Año", years_gestion, key="tg_year", default=years_gestion)
        
        if sel_year_g:
            df_g = df_base[df_base["DIAS/FECHAS"].dt.year.isin(sel_year_g)].copy()

            # ── Matriz Sesiones Mensuales ───────────────────────────────────
            st.markdown("### 🗓️ Matriz de Sesiones Mensuales")
            st.caption("Visualización de la carga de sesiones por mes y coordinadora.")

            mapa_meses = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                          7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

            if not df_g.empty:
                df_matrix = df_g.copy()
                df_matrix["Mes_Num"] = df_matrix["DIAS/FECHAS"].dt.month
                df_matrix["Mes"] = df_matrix["Mes_Num"].map(mapa_meses)

                df_matrix["PROGRAMA"] = df_matrix["PROGRAMA"].apply(lambda x: utils.abbreviate_program_name(x, max_len=40))

                matrix = df_matrix.pivot_table(
                    index=["COORDINADORA RESPONSABLE","PROGRAMA"],
                    columns="Mes", values="DIAS/FECHAS", aggfunc="count", fill_value=0
                )
                meses_orden = list(mapa_meses.values())
                cols_exist = [m for m in meses_orden if m in matrix.columns]
                matrix = matrix[cols_exist]

                st.dataframe(
                    matrix.style.background_gradient(cmap="RdYlGn_r", axis=None, vmin=0, vmax=20),
                    use_container_width=True
                )
            else:
                st.info("No hay datos disponibles.")

            # ── Carga Laboral ───────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### ⚖️ Carga Laboral (Puntaje)")
            st.caption("Cálculo: Factor Sesiones (Sesiones/4) × Factor Alumnos.")

            df_g["Mes_Nombre"] = df_g["DIAS/FECHAS"].dt.month.map(mapa_meses)
            meses_disp_num = sorted(df_g["DIAS/FECHAS"].dt.month.unique())
            meses_carga = ["Todos los meses"] + [mapa_meses[m] for m in meses_disp_num]
            sel_mes_carga = st.selectbox("Mes para Cálculo", meses_carga, key="tg_carga_mes")

            if sel_mes_carga:
                df_carga = df_g.copy() if sel_mes_carga == "Todos los meses" else df_g[df_g["Mes_Nombre"] == sel_mes_carga].copy()

                if not df_carga.empty:
                    col_alumnos = "Nº ALUMNOS"
                    if col_alumnos not in df_carga.columns:
                        col_alumnos = next((c for c in df_carga.columns if "ALUMNO" in c.upper()), None)
                    if not col_alumnos:
                        df_carga["Nº ALUMNOS"] = 0
                        col_alumnos = "Nº ALUMNOS"
                        st.warning("⚠️ No se encontró columna 'Nº ALUMNOS'. Se asume 0.")

                    df_carga[col_alumnos] = pd.to_numeric(df_carga[col_alumnos], errors="coerce").fillna(0)
                    carga_prog = df_carga.groupby(["COORDINADORA RESPONSABLE","PROGRAMA"]).agg(
                        Sesiones=("DIAS/FECHAS", "count"),
                        Alumnos=(col_alumnos, "max")
                    ).reset_index()

                    def get_factor_alumnos(n):
                        if n == 0: return 1.0
                        if n < 20: return 1.0
                        if n < 30: return 1.2
                        if n < 40: return 1.4
                        if n < 49: return 1.7
                        return 2.0

                    carga_prog["Factor_Sesiones"] = carga_prog["Sesiones"] / 4
                    carga_prog["Factor_Alumnos"] = carga_prog["Alumnos"].apply(get_factor_alumnos)
                    carga_prog["Puntaje"] = carga_prog["Factor_Sesiones"] * carga_prog["Factor_Alumnos"]
                    resumen_carga = carga_prog.groupby("COORDINADORA RESPONSABLE")["Puntaje"].sum().reset_index()
                    resumen_carga = resumen_carga.sort_values("Puntaje", ascending=False)

                    c1_g, c2_g = st.columns([2, 1])
                    with c1_g:
                        st.dataframe(resumen_carga, hide_index=True, use_container_width=True,
                                     column_config={
                                         "Puntaje": st.column_config.NumberColumn("Puntaje Total", format="%.2f"),
                                         "COORDINADORA RESPONSABLE": "Coordinadora"
                                     })
                    with c2_g:
                        st.metric("Promedio Gestión", f"{resumen_carga['Puntaje'].mean():.2f}")

                    with st.expander("Ver Detalle por Programa"):
                        carga_prog["Estado Alumnos"] = carga_prog["Alumnos"].apply(lambda x: "Por definir" if x == 0 else "Ok")
                        st.dataframe(carga_prog, hide_index=True, use_container_width=True,
                                     column_config={
                                         "Factor_Sesiones": st.column_config.NumberColumn("Fac. Sesiones", format="%.2f"),
                                         "Factor_Alumnos": st.column_config.NumberColumn("Fac. Alumnos", format="%.1f"),
                                         "Puntaje": st.column_config.NumberColumn("Puntaje", format="%.2f"),
                                     })

                    # ── Reasignación y Balanceo (embebido en Gestión) ──────────────
                    st.markdown("---")
                    st.markdown("""
                    <div style="margin-bottom:8px;">
                      <h3 style="color:#f1f5f9;font-size:17px;font-weight:800;margin:0;">
                        🤖 Reasignación y Balanceo de Carga
                      </h3>
                      <p style="color:#64748b;font-size:12px;margin:4px 0 0;">
                        Simula cambios de coordinadora y genera propuestas automáticas de balanceo equitativo.
                      </p>
                    </div>
                    """, unsafe_allow_html=True)

                    coords_disp_r = sorted_clean(df_base["COORDINADORA RESPONSABLE"])

                    # KPIs de carga
                    pmax_g = resumen_carga["Puntaje"].max() if not resumen_carga.empty else 0
                    pmin_g = resumen_carga["Puntaje"].min() if not resumen_carga.empty else 0
                    kr1, kr2, kr3, kr4 = st.columns(4)
                    kr1.metric("Programas", carga_prog["PROGRAMA"].nunique())
                    kr2.metric("Coordinadoras", resumen_carga["COORDINADORA RESPONSABLE"].nunique())
                    kr3.metric("↑ Máx Carga", f"{pmax_g:.1f}")
                    kr4.metric("⇕ Desbalance", f"{pmax_g - pmin_g:.1f}", delta_color="inverse")

                    # Gráfico carga actual
                    fig_cr = px.bar(
                        resumen_carga, x="COORDINADORA RESPONSABLE", y="Puntaje",
                        color="Puntaje", color_continuous_scale="RdYlGn_r",
                        labels={"COORDINADORA RESPONSABLE": "Coordinadora", "Puntaje": "Puntaje de Carga"}
                    )
                    fig_cr.update_layout(
                        height=300, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                        font=dict(color="#e2e8f0"), showlegend=False,
                        xaxis=dict(tickangle=-30), yaxis=dict(gridcolor="#334155")
                    )
                    st.plotly_chart(fig_cr, use_container_width=True)

                    # Simulador
                    if "sim_cambios" not in st.session_state:
                        st.session_state["sim_cambios"] = {}
                    if "propuesta_balanceo" not in st.session_state:
                        st.session_state["propuesta_balanceo"] = None

                    progs_sim = sorted_clean(carga_prog["PROGRAMA"])

                    # ── NUEVO: Configuración de Disponibilidad ────────────────
                    especiales = ["Sin coord", "POR ASIGNAR", "Por definir", "N/A", "Sin programa", ""]
                    especiales_upper = {e.strip().upper() for e in especiales}
                    
                    c_conf1, c_conf2 = st.columns(2)
                    with c_conf1:
                        # Preseleccionar 'especiales' por defecto si es la primera vez
                        if "coords_fuera_sim" not in st.session_state:
                            st.session_state["coords_fuera_sim"] = [c for c in coords_disp_r if c.strip().upper() in especiales_upper]

                        coords_fuera = st.multiselect(
                            "🚪 OMITIR (Fuera del Simulador - Se ocultan con sus programas)",
                            coords_disp_r,
                            key="coords_fuera_sim",
                            help="Las personas seleccionadas y sus programas NO se mostrarán ni se considerarán para el balanceo. Ideal para trabajar con un grupo reducido.",
                            placeholder="Selecciona para ocultar del análisis..."
                        )
                    
                    # Normalización para usos posteriores
                    fuera_set_norm = {str(c).strip().upper() for c in coords_fuera}

                    with c_conf2:
                        # Coordinadoras fijas (no reciben ni pierden, opcional para mayor control)
                        coords_fijas = st.multiselect(
                            "📌 Mantener Fijas (Congeladas)",
                            [c for c in coords_disp_r if str(c).strip().upper() not in fuera_set_norm],
                            key="coords_fijas_sim",
                            help="Mantienen su carga actual. El balanceador automático no les quitará ni les pondrá nuevos programas.",
                            placeholder="Carga fija..."
                        )

                    # --- FILTRADO TOTAL ---
                    # 1. Los programas de las personas de fuera se ocultan del selectbox y de la tabla
                    progs_excluir = set(carga_prog[carga_prog["COORDINADORA RESPONSABLE"].str.strip().str.upper().isin(fuera_set_norm)]["PROGRAMA"])
                    progs_sim_clean = [p for p in progs_sim if p not in progs_excluir]

                    # 2. El pool de receptoras activas (Solo gente REAL y NO excluida)
                    coords_pool = [
                        c for c in coords_disp_r 
                        if str(c).strip().upper() not in fuera_set_norm 
                        and str(c).strip().upper() not in especiales_upper
                    ]
                    
                    if not coords_pool:
                        coords_pool = [c for c in coords_disp_r if str(c).strip().upper() not in fuera_set_norm]

                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # --- BOTÓN DE BALANCEO REPOSICIONADO ARRIBA ---
                    if st.button("🤖 GENERAR BALANCEO AUTOMÁTICO", use_container_width=True, type="primary", help="Calcula una distribución equitativa basada en promedios y carga actual."):
                        # Filtrar coords válidas
                        def _coords_validas(lista):
                            return [c for c in lista if c is not None and str(c).strip() not in ("", "nan", "None", "N/A")]

                        fijas_set = {str(c).strip().upper() for c in coords_fijas}
                        fuera_set = fuera_set_norm
                        esp_set = especiales_upper
                        
                        coords_disp_clean = _coords_validas(coords_disp_r)
                        coords_activas = [
                            c for c in coords_disp_clean 
                            if str(c).strip().upper() not in fuera_set 
                            and str(c).strip().upper() not in fijas_set
                            and str(c).strip().upper() not in esp_set
                        ]
                        
                        if not coords_activas:
                            coords_activas = [c for c in coords_disp_clean if str(c).strip().upper() not in fuera_set]

                        if not coords_activas and fuera_set:
                            st.warning("⚠️ Debes dejar al menos una coordinadora activa para recibir carga.")
                        else:
                            scores_g = carga_prog.groupby("PROGRAMA")["Puntaje"].sum().sort_values(ascending=False).to_dict()
                            
                            # --- LÓGICA DE BALANCEO PROACTIVA INTEGRAL ---
                            loads_g = {str(c).strip().upper(): 0.0 for c in coords_disp_clean}
                            prop_g = {}
                            all_movable_progs = []

                            total_points = sum(scores_g.values())
                            receptores_norm = {str(c).strip().upper() for c in coords_activas}
                            
                            for pn, ps in scores_g.items():
                                match_orig = carga_prog[carga_prog["PROGRAMA"] == pn]["COORDINADORA RESPONSABLE"]
                                coord_orig = str(match_orig.iloc[0]).strip() if not match_orig.empty else "Sin asignar"
                                current_res = st.session_state["sim_cambios"].get(pn, coord_orig)
                                current_norm = str(current_res).strip().upper()
                                
                                if current_norm in fijas_set or current_norm in fuera_set:
                                    if current_norm in loads_g: loads_g[current_norm] += ps
                                else:
                                    all_movable_progs.append((pn, ps, current_norm))

                            all_movable_progs.sort(key=lambda x: x[1], reverse=True)
                            for pn, ps, old_norm in all_movable_progs:
                                best_norm = min(receptores_norm, key=lambda c_norm: loads_g.get(c_norm, 0))
                                if old_norm in esp_set or loads_g[best_norm] + ps < loads_g.get(old_norm, 0):
                                    prop_g[pn] = next(c for c in coords_disp_clean if str(c).strip().upper() == best_norm)
                                    loads_g[best_norm] += ps
                                else:
                                    if old_norm in loads_g: loads_g[old_norm] += ps

                            prop_g_final = {}
                            for p, n_c in prop_g.items():
                                orig = str(carga_prog[carga_prog["PROGRAMA"] == p]["COORDINADORA RESPONSABLE"].iloc[0]).strip()
                                if str(n_c).strip() != orig: prop_g_final[p] = n_c

                            st.session_state["propuesta_balanceo"] = prop_g_final
                            st.rerun()

                    st.markdown("---")

                    # --- RESULTADOS DEL BALANCEO AUTOMÁTICO (Reposicionados) ---
                    if st.session_state["propuesta_balanceo"] is not None:
                        st.markdown("#### ✨ PROPUESTA DE REASIGNACIÓN")

                        _prop_limpio = {
                            k: v for k, v in st.session_state["propuesta_balanceo"].items()
                            if k is not None
                            and v is not None
                            and str(v).strip() not in ("", "nan", "None", "N/A")
                        }

                        if len(_prop_limpio) > 0:
                            st.info(f"💡 SE HAN IDENTIFICADO {len(_prop_limpio)} PROGRAMAS PARA REASIGNAR (HUÉRFANOS O POR EXCESO DE CARGA).")
                            
                            rows_p = []
                            for pv, cv in _prop_limpio.items():
                                match_p = carga_prog[carga_prog["PROGRAMA"] == pv]["COORDINADORA RESPONSABLE"]
                                raw_old = match_p.iloc[0] if not match_p.empty else "N/A"
                                old2 = str(raw_old).strip() if raw_old is not None else "N/A"
                                if old2 in ("nan", "None", "", "N/A"): old2 = "SIN ASIGNAR"
                                rows_p.append({"Programa": pv, "Coordinadora Actual": old2, "Coordinadora Propuesta 🔄": str(cv)})
                            
                            st.dataframe(pd.DataFrame(rows_p), hide_index=True, use_container_width=True)

                            # 📊 Análisis de impacto
                            st.markdown("##### 📊 IMPACTO VISUAL: ANTES VS DESPUÉS")
                            
                            prop_map = _prop_limpio
                            df_prop_sim = carga_prog.copy()
                            for pv_g, cv_g in prop_map.items():
                                df_prop_sim.loc[df_prop_sim["PROGRAMA"] == pv_g, "COORDINADORA RESPONSABLE"] = str(cv_g)
                            
                            res_prop = (
                                df_prop_sim.groupby("COORDINADORA RESPONSABLE", dropna=True)["Puntaje"].sum().reset_index()
                                .rename(columns={"Puntaje": "Puntaje_Propuesta"})
                            )

                            # Resumen tabla comparativa
                            st.markdown("##### 📝 RESUMEN DE REDISTRIBUCIÓN (CARGA EN PUNTOS)")
                            comp_balanceo = pd.merge(
                                resumen_carga[["COORDINADORA RESPONSABLE", "Puntaje"]].rename(columns={"Puntaje": "Actual"}),
                                res_prop.rename(columns={"Puntaje_Propuesta": "Propuesto"}),
                                on="COORDINADORA RESPONSABLE", how="outer"
                            ).fillna(0)
                            
                            comp_balanceo = comp_balanceo[comp_balanceo["COORDINADORA RESPONSABLE"].str.upper().isin({str(c).upper() for c in coords_pool})]
                            comp_balanceo["Dif."] = comp_balanceo["Propuesto"] - comp_balanceo["Actual"]
                            
                            def _style_diff(val):
                                color = '#f87171' if val > 0 else '#4ade80' if val < 0 else '#94a3b8'
                                return f'color: {color}; font-weight: bold'

                            st.dataframe(
                                comp_balanceo.style.applymap(_style_diff, subset=['Dif.'])
                                .format({"Actual": "{:.1f}", "Propuesto": "{:.1f}", "Dif.": "{:+.1f}"}),
                                hide_index=True, use_container_width=True
                            )

                            # Unir con carga actual para gráfico
                            resumen_carga_clean = resumen_carga[
                                resumen_carga["COORDINADORA RESPONSABLE"].notna() &
                                resumen_carga["COORDINADORA RESPONSABLE"].astype(str).str.strip().isin(
                                    [c for c in resumen_carga["COORDINADORA RESPONSABLE"].astype(str) if c.strip() not in ("nan", "None", "N/A", "")]
                                )
                            ]
                            comp_prop = pd.merge(
                                resumen_carga_clean[["COORDINADORA RESPONSABLE", "Puntaje"]].rename(columns={"Puntaje": "Puntaje_Actual"}),
                                res_prop, on="COORDINADORA RESPONSABLE", how="outer"
                            ).fillna(0)
                            
                            _valid_coords = [c for c in comp_prop["COORDINADORA RESPONSABLE"].astype(str) if c.strip() not in ("nan", "None", "N/A", "")]
                            comp_prop = comp_prop[comp_prop["COORDINADORA RESPONSABLE"].notna() & comp_prop["COORDINADORA RESPONSABLE"].astype(str).str.strip().isin(_valid_coords)].copy()
                            comp_prop = comp_prop[~comp_prop["COORDINADORA RESPONSABLE"].str.strip().str.upper().isin(fuera_set_norm)].copy()

                            df_chart = pd.melt(comp_prop, id_vars=["COORDINADORA RESPONSABLE"], value_vars=["Puntaje_Actual", "Puntaje_Propuesta"], var_name="Estado", value_name="Puntaje")
                            df_chart["Estado"] = df_chart["Estado"].map({"Puntaje_Actual": "🔴 Actual", "Puntaje_Propuesta": "🟢 Propuesta"})
                            df_chart["Coord"] = df_chart["COORDINADORA RESPONSABLE"].apply(lambda x: " ".join(str(x).split()[:2]))
                            
                            fig_prop = px.bar(df_chart, x="Coord", y="Puntaje", color="Estado", barmode="group",
                                             color_discrete_map={"🔴 Actual": "#ef4444", "🟢 Propuesta": "#10b981"},
                                             labels={"Coord": "Coordinadora", "Puntaje": "Carga (Puntaje)", "Estado": ""})
                            fig_prop.update_layout(height=340, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                                                 font=dict(color="#e2e8f0", size=11), legend=dict(orientation="h", y=1.08, x=0),
                                                 xaxis=dict(tickangle=-20, gridcolor="rgba(0,0,0,0)"), yaxis=dict(gridcolor="#334155"), bargap=0.25)
                            
                            if not comp_prop.empty and comp_prop["Puntaje_Actual"].sum() > 0:
                                avg_ideal = comp_prop["Puntaje_Actual"].sum() / max(len(comp_prop), 1)
                                fig_prop.add_hline(y=avg_ideal, line_dash="dot", line_color="#f59e0b",
                                                  annotation_text=f"⚖️ Ideal: {avg_ideal:.1f}", annotation_position="top right", annotation_font_color="#f59e0b")
                            st.plotly_chart(fig_prop, use_container_width=True)

                            # Métricas 
                            before_std = comp_prop["Puntaje_Actual"].std()
                            after_std  = comp_prop["Puntaje_Propuesta"].std()
                            smz1, smz2, smz3 = st.columns(3)
                            smz1.metric("📉 Desv. Estándar Antes",    f"{before_std:.2f}")
                            smz2.metric("📉 Desv. Estándar Después",  f"{after_std:.2f}",
                                       delta=f"{after_std - before_std:.2f}", delta_color="inverse")
                            smz3.metric("✅ Programas a mover", len(prop_map))

                            # Acción 
                            _pdf_data = None
                            try:
                                _pdf_data = generar_pdf_propuesta(rows_prop=rows_p, comp_df=comp_prop, fecha_str=datetime.now().strftime("%d/%m/%Y %H:%M"))
                            except: pass

                            bpz1, bpz2, bpz3 = st.columns(3)
                            with bpz1:
                                if st.button("✅ Aceptar Propuesta", type="primary", use_container_width=True, key="acc_prop_top"):
                                    for pvk, cvk in st.session_state["propuesta_balanceo"].items(): st.session_state["sim_cambios"][pvk] = cvk
                                    st.session_state["propuesta_balanceo"] = None
                                    st.rerun()
                            with bpz2:
                                if _pdf_data:
                                    render_static_download("📄 Descargar PDF", _pdf_data, f"propuesta_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
                                else: st.button("📄 Descargar PDF", disabled=True, use_container_width=True, key="pdf_no_top")
                            with bpz3:
                                if st.button("❌ Descartar", use_container_width=True, key="disc_prop_top"):
                                    st.session_state["propuesta_balanceo"] = None
                                    st.rerun()
                        else:
                            st.success("👍 La distribución ya está óptimamente balanceada.")
                            if st.button("Cerrar", key="gest_cerrar_top"):
                                st.session_state["propuesta_balanceo"] = None
                                st.rerun()

                    st.markdown("---")
                    st.markdown("##### ✍️ REASIGNACIÓN MANUAL")
                    col_s1, col_s2, col_s3 = st.columns(3)
                    with col_s1:
                        prog_sel_g = st.selectbox("PROGRAMA A REASIGNAR", progs_sim_clean, key="sim_prog")
                    coord_real_g = carga_prog[carga_prog["PROGRAMA"]==prog_sel_g]["COORDINADORA RESPONSABLE"].iloc[0] \
                        if not carga_prog[carga_prog["PROGRAMA"]==prog_sel_g].empty else "N/A"
                    coord_sim_g = st.session_state["sim_cambios"].get(prog_sel_g, coord_real_g)
                    with col_s2:
                        st.text_input("→ COORDINADORA ACTUAL", value=coord_sim_g, disabled=True, key="gest_coord_actual")
                    with col_s3:
                        # Usamos la lista filtrada coords_pool
                        nueva_coord_g = st.selectbox("NUEVA COORDINADORA", coords_pool, key="sim_new_coord")

                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        if st.button("🔄 APLICAR CAMBIO MANUAL", type="primary", use_container_width=True):
                            st.session_state["sim_cambios"][prog_sel_g] = nueva_coord_g
                            st.session_state["propuesta_balanceo"] = None
                            st.rerun()
                    with col_b2:
                        if (st.session_state["sim_cambios"] or st.session_state["propuesta_balanceo"]) \
                           and st.button("🗑️ BORRAR SIMULADOR", use_container_width=True):
                            st.session_state["sim_cambios"] = {}
                            st.session_state["propuesta_balanceo"] = None
                            st.rerun()


                    # Impacto simulación
                    if st.session_state["sim_cambios"]:
                        st.markdown("---")
                        st.markdown("#### 📊 Impacto de la Simulación")
                        df_sim_g = carga_prog.copy()
                        for ps2, cs2 in st.session_state["sim_cambios"].items():
                            df_sim_g.loc[df_sim_g["PROGRAMA"]==ps2, "COORDINADORA RESPONSABLE"] = cs2
                        sim_g = df_sim_g.groupby(["COORDINADORA RESPONSABLE","PROGRAMA"]).agg(
                            Sesiones=("Sesiones","sum"), Alumnos=("Alumnos","max")
                        ).reset_index()
                        sim_g["Puntaje"] = (sim_g["Sesiones"]/4) * sim_g["Alumnos"].apply(get_factor_alumnos)
                        res_sim_g = sim_g.groupby("COORDINADORA RESPONSABLE")["Puntaje"].sum().reset_index()
                        comp_g = pd.merge(
                            resumen_carga[["COORDINADORA RESPONSABLE","Puntaje"]],
                            res_sim_g[["COORDINADORA RESPONSABLE","Puntaje"]],
                            on="COORDINADORA RESPONSABLE", how="outer", suffixes=("_Actual","_Simulado")
                        ).fillna(0)
                        comp_g["Variación"] = comp_g["Puntaje_Simulado"] - comp_g["Puntaje_Actual"]
                        cg1, cg2 = st.columns([2, 1])
                        with cg1:
                            st.dataframe(comp_g.sort_values("Puntaje_Simulado", ascending=False),
                                         hide_index=True, use_container_width=True,
                                         column_config={
                                             "Puntaje_Actual": st.column_config.NumberColumn("Carga Actual", format="%.2f"),
                                             "Puntaje_Simulado": st.column_config.NumberColumn("Carga Simulada", format="%.2f"),
                                             "Variación": st.column_config.NumberColumn("Δ", format="%+.2f"),
                                         })
                        with cg2:
                            pmax2 = res_sim_g["Puntaje"].max() if not res_sim_g.empty else 0
                            pmin2 = res_sim_g["Puntaje"].min() if not res_sim_g.empty else 0
                            st.metric("↕ Desbalance Antes", f"{pmax_g - pmin_g:.2f}")
                            st.metric("↕ Desbalance Después", f"{pmax2 - pmin2:.2f}",
                                      delta=f"{(pmax2 - pmin2) - (pmax_g - pmin_g):.2f}", delta_color="inverse")
                        st.caption(f"📌 {len(st.session_state['sim_cambios'])} cambio(s) en simulación")
                else:
                    st.info(f"No hay datos para {sel_mes_carga}.")
        else:
            st.info("Selecciona un año.")
    elif password:
        st.error("Contraseña incorrecta")
    else:
        st.info("🔒 Ingrese contraseña para continuar.")



# =============================================================================
# TAB VALIDACIONES
# =============================================================================
# =============================================================================
# TAB TURNOS: GESTIÓN Y VALIDACIÓN DE TURNOS
# =============================================================================
if tab_turnos:
  with tab_turnos:
    st.markdown("## 🧩 Gestión de Turnos y Cargas")
    st.markdown("Auditoría de cumplimiento de reglas de negocio para asignación de coordinadoras.")

    # Filtros de Análisis
    with st.container():
        c_t1, c_t2, c_t3 = st.columns([1, 1, 2])
        
        # Obtener años disponibles
        if "DIAS/FECHAS" in df_base.columns:
            anos_disp = sorted(df_base["DIAS/FECHAS"].dt.year.dropna().unique().astype(int))
        else:
            anos_disp = [datetime.now().year]
            
        ano_sel = c_t1.selectbox("Año", anos_disp, index=len(anos_disp)-1 if anos_disp else 0)
        
        # Selector de Mes
        meses_nombres = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
        }
        mes_actual = datetime.now().month
        mes_sel_nombre = c_t2.selectbox("Mes", list(meses_nombres.values()), index=mes_actual-1)
        
        # Calcular fechas inicio y fin del mes seleccionado
        mes_idx = [k for k,v in meses_nombres.items() if v == mes_sel_nombre][0]
        # calendar ya importado al inicio del archivo
        _, last_day = calendar.monthrange(ano_sel, mes_idx)
        
        f_inicio = datetime(ano_sel, mes_idx, 1).date()
        f_fin = datetime(ano_sel, mes_idx, last_day).date()
    
    # Filtrar DF principal
    if "DIAS/FECHAS" in df_base.columns:
        mask_t = (df_base["DIAS/FECHAS"].dt.date >= f_inicio) & (df_base["DIAS/FECHAS"].dt.date <= f_fin)
        df_turnos = df_base[mask_t].copy()
        
        if not df_turnos.empty:
            st.info(f"📅 Analizando planificación para: **{mes_sel_nombre} {ano_sel}**")
            
            # --- INTRODUCCIÓN ---
            st.markdown("""
            Esta herramienta audita la programación para asegurar una **operación eficiente y balanceada**. 
            Revisamos tres aspectos clave:
            1. **📍 Cobertura de Sedes**: Evitar aglomeraciones innecesarias.
            2. **⚡ Carga Diaria**: Proteger a las coordinadoras de sobrecarga.
            3. **🏃‍♀️ Logística**: Evitar desplazamientos imposibles entre sedes.
            """)
            st.markdown("---")
            
            # --- 1. CONTROL DE SEDES ---
            st.subheader("1. 📍 Control de Ocupación en Sedes")
            st.caption("Verificamos que no haya exceso de personal en una misma sede (Máximo sugerido: 2 coordinadoras).")
            
            if "SEDE" in df_turnos.columns and "COORDINADORA RESPONSABLE" in df_turnos.columns:
                sat_sede = df_turnos.groupby(["DIAS/FECHAS", "SEDE"])["COORDINADORA RESPONSABLE"].nunique().reset_index(name="Cant_Coords")
                
                # --- HEATMAP ---
                st.markdown("**Mapa de Calor de Ocupación**")
                if not sat_sede.empty:
                    # (Código del Heatmap se mantiene igual, solo cambio el título visual)
                    sat_sede["Fecha_Str"] = sat_sede["DIAS/FECHAS"].dt.strftime("%d-%m") # Fecha más corta
                    pivot_sede = sat_sede.pivot(index="SEDE", columns="Fecha_Str", values="Cant_Coords").fillna(0)
                    pivot_sede = pivot_sede.reindex(sorted(pivot_sede.columns, key=lambda x: datetime.strptime(str(x) + f"-{ano_sel}", "%d-%m-%Y")), axis=1) # Sort hack
                    
                    fig_occ = px.imshow(
                        pivot_sede,
                        labels=dict(x="Día", y="Sede", color="Personal"),
                        x=pivot_sede.columns,
                        y=pivot_sede.index,
                        color_continuous_scale=[[0,"#dcfce7"],[0.5,"#fde047"],[1,"#ef4444"]], # Verde suave -> Amarillo -> Rojo
                        aspect="auto", text_auto=True
                    )
                    fig_occ.update_layout(height=300, margin=dict(t=20, b=20), xaxis=dict(side="top", type="category"), font=dict(color="#e2e8f0"), paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_occ, use_container_width=True)

                # Alertas
                conflictos_sede = sat_sede[sat_sede["Cant_Coords"] > 2].sort_values("DIAS/FECHAS")
                if not conflictos_sede.empty:
                    st.warning(f"⚠️ Atención: Hay **{len(conflictos_sede)} días** con más de 2 coordinadoras en una sede.")
                    conflictos_sede["Fecha"] = conflictos_sede["DIAS/FECHAS"].dt.strftime("%A %d")
                    st.dataframe(conflictos_sede[["Fecha", "SEDE", "Cant_Coords"]], hide_index=True, use_container_width=True)
                else:
                    st.success("✅ **Sedes Balanceadas**: Ninguna sede supera el límite de 2 personas.")
            else:
                st.warning("Faltan datos para analizar sedes.")

            st.markdown("---")

            # --- 2. CARGA INDIVIDUAL ---
            st.subheader("2. ⚡ Carga de Trabajo Diaria")
            st.caption("Monitoreamos que ninguna coordinadora tenga que atender más de **2 programas distintos** el mismo día.")
            
            if "COORDINADORA RESPONSABLE" in df_turnos.columns and "PROGRAMA" in df_turnos.columns:
                sobrecarga = df_turnos.groupby(["DIAS/FECHAS", "COORDINADORA RESPONSABLE"])["PROGRAMA"].nunique().reset_index(name="Cant_Progs")
                conflictos_carga = sobrecarga[sobrecarga["Cant_Progs"] > 2].sort_values("DIAS/FECHAS")
                
                if not conflictos_carga.empty:
                    st.error(f"🚫 **Sobrecarga Detectada**: Se encontraron {len(conflictos_carga)} casos de alta exigencia.")
                    conflictos_carga["Fecha"] = conflictos_carga["DIAS/FECHAS"].dt.strftime("%A %d")
                    st.dataframe(conflictos_carga[["Fecha", "COORDINADORA RESPONSABLE", "Cant_Progs"]], hide_index=True, use_container_width=True)
                else:
                    st.success("✅ **Carga Saludable**: Nadie tiene más de 2 programas diarios.")
            
            st.markdown("---")

            # --- 3. LOGÍSTICA ---
            st.subheader("3. 🏃‍♀️ Movilidad y Desplazamientos")
            st.caption("Verificamos que nadie esté asignado a **dos sedes diferentes** el mismo día (imposibilidad logística).")
            
            if "SEDE" in df_turnos.columns:
                sedes_cross = df_turnos.groupby(["DIAS/FECHAS", "COORDINADORA RESPONSABLE"])["SEDE"].nunique().reset_index(name="Cant_Sedes")
                conflictos_sedes = sedes_cross[sedes_cross["Cant_Sedes"] > 1]
                
                if not conflictos_sedes.empty:
                    st.error(f"🚨 **Conflicto Logístico**: {len(conflictos_sedes)} casos de asignación en múltiples sedes.")
                    conflictos_sedes["Fecha"] = conflictos_sedes["DIAS/FECHAS"].dt.strftime("%A %d")
                    st.dataframe(conflictos_sedes[["Fecha", "COORDINADORA RESPONSABLE", "Cant_Sedes"]], hide_index=True, use_container_width=True)
                else:
                    st.success("✅ **Logística OK**: No hay cruces de sedes.")

            # --- NUEVA SECCIÓN: CALENDARIO Y EQUIDAD ---
            st.markdown("---")
            st.markdown("### ⚖️ Balance y Equidad")
            
            # Métricas de Equidad
            col_eq1, col_eq2 = st.columns([2, 1])
            with col_eq1:
                st.caption("Resumen de carga laboral en el periodo seleccionado.")
                
                df_turnos["Es_Sabado"] = df_turnos["DIAS/FECHAS"].dt.weekday == 5
                
                # Agregación por coordinadora
                equidad = df_turnos.groupby("COORDINADORA RESPONSABLE").agg(
                    Dias_Totales=("DIAS/FECHAS", "nunique"),
                    Sabados_Asignados=("Es_Sabado", lambda x: x[x].nunique() if x.any() else 0),
                    Sedes_Visitadas=("SEDE", "nunique")
                ).reset_index()
                
                # Ordenar por quine trabaja más
                equidad = equidad.sort_values("Dias_Totales", ascending=False)
                
                st.dataframe(
                    equidad, 
                    hide_index=True, 
                    use_container_width=True,
                    column_config={
                        "Dias_Totales": st.column_config.ProgressColumn("Días Activos", max_value=30, format="%d días"),
                        "Sabados_Asignados": st.column_config.NumberColumn("Sábados", format="%d 🛑")
                    }
                )
            
            with col_eq2:
                st.info("💡 **Tip de Gestión**: Intenta rotar los sábados equitativamente. Una carga de sábados desbalanceada es la principal causa de fatiga.")

            st.markdown("---")
            st.markdown("### 📅 Matriz de Operaciones (Roster)")

            if not df_turnos.empty:
                # Selector de vista
                vista_cal = st.radio("Modo:", ["🧩 Roster con Sedes", "👤 Detalle Individual"], horizontal=True)
                
                if vista_cal == "🧩 Roster con Sedes":
                    df_asistencia = df_turnos.copy()
                    df_asistencia["Dia_Mes"] = df_asistencia["DIAS/FECHAS"].dt.day
                    
                    # Función inteligente: Obtener inicial de la sede
                    def get_sede_initials(series):
                        sedes = series.unique()
                        # Tomar las primeras 3 letras de la primera sede
                        if len(sedes) > 0:
                            s = str(sedes[0]).upper()
                            # Mapeo común
                            if "VITACURA" in s: return "VIT"
                            if "PEÑALOLEN" in s or "PENALOLEN" in s: return "PEÑ"
                            if "BELLAVISTA" in s: return "BEL"
                            if "ONLINE" in s or "ZOOM" in s: return "ONL"
                            return s[:3]
                        return ""

                    roster = df_asistencia.pivot_table(
                        index="COORDINADORA RESPONSABLE", 
                        columns="Dia_Mes", 
                        values="SEDE", 
                        aggfunc=get_sede_initials
                    ).fillna("")
                    
                    dias_rango = sorted(df_asistencia["Dia_Mes"].unique())
                    roster = roster.reindex(columns=dias_rango, fill_value="")
                    
                    st.dataframe(roster, use_container_width=True)
                    st.caption("Claves: **VIT**: Vitacura, **PEÑ**: Peñalolén, **BEL**: Bellavista, **ONL**: Online.")

                else: # Detalle Individual
                    coords_disp = sorted_clean(df_turnos["COORDINADORA RESPONSABLE"])
                    if coords_disp:
                        c_sel1, c_sel2 = st.columns([3, 1])
                        coord_sel = c_sel1.selectbox("Coordinadora:", coords_disp)
                        
                        df_indiv = df_turnos[df_turnos["COORDINADORA RESPONSABLE"] == coord_sel].copy()
                        
                        if not df_indiv.empty:
                            df_indiv["Dia_Semana_Num"] = df_indiv["DIAS/FECHAS"].dt.weekday
                            dias_map = {0:"Lun", 1:"Mar", 2:"Mié", 3:"Jue", 4:"Vie", 5:"Sáb", 6:"Dom"}
                            df_indiv["Dia_Nom"] = df_indiv["Dia_Semana_Num"].map(dias_map)
                            df_indiv["Dia_Num"] = df_indiv["DIAS/FECHAS"].dt.day.astype(str)
                            df_indiv["Week"] = df_indiv["DIAS/FECHAS"].dt.isocalendar().week # Re-calcular semana
                            
                            # Hack para que el calendario se vea bien (semana vs dia)
                            cal_data = df_indiv.groupby(["DIAS/FECHAS","Dia_Semana_Num","Week","Dia_Num"]).size().reset_index(name="Carga")
                            
                            fig_cal = px.scatter(
                                cal_data, x="Dia_Semana_Num", y="Week", text="Dia_Num", size=[35]*len(cal_data),
                                color="Carga", color_continuous_scale="Blugrn",
                                title=f"Calendario: {coord_sel}"
                            )
                            fig_cal.update_traces(marker=dict(symbol="square", line=dict(width=1, color="white")), textfont=dict(color="white", size=14))
                            fig_cal.update_layout(
                                height=350,
                                xaxis=dict(tickmode="array", tickvals=[0,1,2,3,4,5,6], ticktext=["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Dom"], side="top"),
                                yaxis=dict(showticklabels=False, autorange="reversed", title=None),
                                coloraxis_showscale=False
                            )
                            st.plotly_chart(fig_cal, use_container_width=True)
                            st.dataframe(df_indiv[["DIAS/FECHAS", "HORARIO", "SEDE", "PROGRAMA"]], hide_index=True, use_container_width=True)

        else:
            st.info("No hay clases programadas en este rango de fechas.")
    else:
        st.error("No se puede validar sin columna de Fechas.")

# =============================================================================
# TAB VALIDACIONES
# =============================================================================
if tab_validaciones:
  with tab_validaciones:
    st.info("💡 Análisis de consistencia de la programación.")

    cols_horas = ["HORA_INICIO","HORA_FIN"]
    columnas_ok = all(c in df_base.columns for c in cols_horas)

    if not columnas_ok:
        st.warning("⚠️ Falta información de Horarios (HORA_INICIO, HORA_FIN).")
    else:
        df_val = df_base.dropna(subset=cols_horas).copy()

        def combine_dt(row, col_hora):
            t = row[col_hora]
            if isinstance(t, str):
                t = t.strip()
                try: t_obj = datetime.strptime(t, "%H:%M:%S").time()
                except:
                    try: t_obj = datetime.strptime(t, "%H:%M").time()
                    except: return pd.NaT
                return datetime.combine(row["DIAS/FECHAS"].date(), t_obj)
            return pd.NaT

        df_val["start_dt"] = df_val.apply(lambda x: combine_dt(x, "HORA_INICIO"), axis=1)
        df_val["end_dt"] = df_val.apply(lambda x: combine_dt(x, "HORA_FIN"), axis=1)
        df_val = df_val.dropna(subset=["start_dt","end_dt"])

        # Validación Profesores
        with st.expander("👨‍🏫 Choques de Horario (Profesores)", expanded=True):
            choques_prof = []
            ignore_prof = ["SIN PROFESOR","POR DEFINIR","NAN"]
            df_prof = df_val[~df_val["PROFESOR"].astype(str).str.upper().isin(ignore_prof)] if "PROFESOR" in df_val.columns else pd.DataFrame()

            if not df_prof.empty:
                for prof, sub in df_prof.groupby("PROFESOR"):
                    sub = sub.sort_values("start_dt")
                    for i in range(len(sub)-1):
                        curr = sub.iloc[i]
                        nxt = sub.iloc[i+1]
                        if nxt["start_dt"] < curr["end_dt"]:
                            choques_prof.append({
                                "Profesor": prof,
                                "Fecha": curr["DIAS/FECHAS"].strftime("%d-%m-%Y"),
                                "Conflicto 1": f"{curr['PROGRAMA']} ({curr['HORA_INICIO']}-{curr['HORA_FIN']})",
                                "Conflicto 2": f"{nxt['PROGRAMA']} ({nxt['HORA_INICIO']}-{nxt['HORA_FIN']})"
                            })

            if choques_prof:
                st.error(f"⚠️ {len(choques_prof)} conflictos de profesor.")
                st.dataframe(pd.DataFrame(choques_prof), use_container_width=True)
            else:
                st.success("✅ Sin conflictos de profesor.")

        # Validación Salas
        with st.expander("🏢 Choques de Sala", expanded=True):
            if "SALA" not in df_val.columns:
                st.info("No hay columna 'SALA'.")
            else:
                choques_sala = []
                salas_ignore = ["SIN SALA","ONLINE","ZOOM","NAN"]
                mask_salas = df_val["SALA"].astype(str).str.upper().apply(lambda x: not any(ign in x for ign in salas_ignore))
                df_salas = df_val[mask_salas]

                if not df_salas.empty:
                    for sala, sub in df_salas.groupby("SALA"):
                        sub = sub.sort_values("start_dt")
                        for i in range(len(sub)-1):
                            curr = sub.iloc[i]
                            nxt = sub.iloc[i+1]
                            if nxt["start_dt"] < curr["end_dt"]:
                                choques_sala.append({
                                    "Sala": sala,
                                    "Fecha": curr["DIAS/FECHAS"].strftime("%d-%m-%Y"),
                                    "Evento 1": f"{curr['PROGRAMA']}",
                                    "Evento 2": f"{nxt['PROGRAMA']}"
                                })

                if choques_sala:
                    st.warning(f"⚠️ {len(choques_sala)} conflictos de sala.")
                    st.dataframe(pd.DataFrame(choques_sala), use_container_width=True)
                else:
                    st.success("✅ Sin conflictos de sala.")

# =============================================================================
# TAB AYUDA & MANUAL: Módulo de instrucciones
# =============================================================================
if tab_manual:
    with tab_manual:
        st.markdown("## Manual de Usuario")
        st.markdown("Selecciona una seccion para ver las instrucciones:")
        st.markdown("---")

        PASOS = [
            ("📂 Carga de Datos",
             "Usa la barra lateral para conectarte a la carpeta local o de OneDrive. Tambien puedes subir tu archivo directamente haciendo drag and drop.",
             "manual_carga.png",
             "https://dummyimage.com/900x350/1e293b/94a3b8&text=1.+Carga+de+Datos"),
            ("📊 Dashboard",
             "El Dashboard principal muestra KPIs clave: sesiones de hoy, proximos 7 dias, alertas de cursos por terminar y sesiones sin profesor asignado.",
             "manual_dashboard.png",
             "https://dummyimage.com/900x350/1e293b/94a3b8&text=2.+Dashboard+y+KPIs"),
            ("Coordinadoras",
             "Navega al tab Coordinadoras para filtrar las clases por responsable. Puedes ver su carga semanal, programas y sesiones asignadas.",
             "manual_coordinadoras.png",
             "https://dummyimage.com/900x350/1e293b/94a3b8&text=3.+Filtros+Coordinadoras"),
            ("📋 Kanban",
             "El Kanban agrupa los programas por estado: Planificado, En Curso, Finalizado o Declarado. Puedes cambiar el estado manualmente desde el mismo panel.",
             "manual_kanban.png",
             "https://dummyimage.com/900x350/1e293b/94a3b8&text=4.+Panel+Kanban"),
            ("📅 Calendarios",
             "Visualiza el calendario academico semanal o mensual. Si cargas un Excel de reservas podras ver que salas estan bloqueadas o confirmadas.",
             "manual_calendario.png",
             "https://dummyimage.com/900x350/1e293b/94a3b8&text=5.+Calendario+y+Reservas"),
        ]

        tabs_manual = st.tabs([p[0] for p in PASOS])
        for i, tab_paso in enumerate(tabs_manual):
            titulo, desc, img_name, placeholder = PASOS[i]
            with tab_paso:
                st.markdown(f"**{desc}**")
                st.markdown("")
                ruta_img = Path(__file__).parent / "img" / img_name
                if ruta_img.exists():
                    st.image(str(ruta_img), use_container_width=True)
                else:
                    st.image(placeholder, use_container_width=True)

if tab_faqs:
    with tab_faqs:
        st.markdown("## Preguntas Frecuentes")
        st.markdown("---")

        FAQS = [
            ("Error: columna DIAS/FECHAS no encontrada",
             "Verifica que tu Excel tenga una columna llamada exactamente **DIAS/FECHAS** en la primera hoja. Las mayusculas y la barra inclinada son obligatorias."),
            ("Como descargo un informe PDF de coordinadora",
             "En la barra lateral, baja a **Centro de Descargas - Informe Personalizado**, filtra por Coordinadora y haz clic en Descargar Informe PDF."),
            ("El Kanban no muestra programas",
             "El Kanban requiere las columnas **PROGRAMA** y **DIAS/FECHAS** para calcular fechas de inicio y fin de cada programa."),
            ("Como cargo archivos desde OneDrive",
             "En la barra lateral selecciona **Carpeta personalizada (OneDrive, etc.)** e ingresa la ruta completa de tu carpeta de OneDrive."),
            ("Puedo filtrar por mas de una coordinadora",
             "Por ahora el filtro es individual. Selecciona la coordinadora deseada en el selector de la pestana correspondiente."),
        ]

        for pregunta, respuesta in FAQS:
            st.markdown(f"""
<div style="background:#1e293b;border:1px solid #334155;border-radius:8px;
            padding:14px 18px;margin-bottom:10px;">
  <p style="color:#38bdf8;font-weight:700;margin:0 0 6px 0;font-size:14px;">❓ {pregunta}</p>
  <p style="color:#cbd5e1;margin:0;font-size:13px;line-height:1.6;">{respuesta}</p>
</div>
""", unsafe_allow_html=True)

st.caption("Dashboard v2.0 - Carga local sin Axios ✅")
