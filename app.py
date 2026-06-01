import re
import unicodedata
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Reporte WhatsApp Arqueología", page_icon="📋", layout="wide")

st.title("📋 Generador de reporte desde WhatsApp")
st.write(
    "Pega reportes de WhatsApp, revisa/edita los datos extraídos y descarga "
    "un Excel listo para copiar a la planilla maestra."
)

# ==========================================================
# CONFIGURACIÓN
# ==========================================================

MATERIAL_MAP = {
    "loza": "loza", "lozas": "loza",
    "ceramica": "cerámica", "ceramicas": "cerámica",
    "cerámica": "cerámica", "cerámicas": "cerámica",
    "vidrio": "vidrio", "vidrios": "vidrio",
    "metal": "metal", "metales": "metal",
    "osteofauna": "osteofauna",
    "plastico": "plástico", "plasticos": "plástico",
    "plástico": "plástico", "plásticos": "plástico",
    "malacologico": "malacológico", "malacologicos": "malacológico",
    "malacológico": "malacológico", "malacológicos": "malacológico",
    "miscelaneo": "misceláneo", "miscelaneos": "misceláneo",
    "misceláneo": "misceláneo", "misceláneos": "misceláneo",
    "baldosa": "baldosa", "baldosas": "baldosa",
}

ORDEN_DESCRIPCION = [
    "vidrio", "loza", "metal", "osteofauna", "cerámica", "plástico", "baldosa"
]

ORDEN_FRECUENCIA = [
    "loza", "metal", "baldosa", "cerámica", "vidrio", "osteofauna", "misceláneo"
]

MATERIAL_SIN_MATERIAL = "sin material"
MATERIALES_EXCLUIDOS_DESCRIPCION = {"malacológico", "misceláneo", MATERIAL_SIN_MATERIAL}


@dataclass
class Registro:
    fecha: Optional[str]
    unidad: Optional[str]
    nivel: int
    estrato: Optional[str]
    cronologia: str
    material: str
    cantidad: Optional[int]
    material_nuevo: bool = False
    linea_original: str = ""


# ==========================================================
# UTILIDADES
# ==========================================================

def strip_accents_lower(text):
    text = str(text).strip().lower()
    return "".join(
        c for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )


def limpiar_texto_material(raw):
    key = strip_accents_lower(raw)
    key = re.sub(r"[^a-zñáéíóúü ]", "", key).strip()
    key = re.sub(r"\s+", " ", key)
    return key


def normalizar_material(raw):
    key = limpiar_texto_material(raw)
    return MATERIAL_MAP.get(key)


def normalizar_material_desconocido(raw):
    """Conserva materiales no catastrados para que no se pierdan."""
    key = limpiar_texto_material(raw)
    if not key:
        return None
    # Evita capturar frases largas como si fueran materiales.
    if len(key.split()) > 3:
        return None
    return key


def material_con_alerta(raw):
    mat = normalizar_material(raw)
    if mat:
        return mat, False
    mat_nuevo = normalizar_material_desconocido(raw)
    if mat_nuevo:
        return mat_nuevo, True
    return None, False


def normalizar_fecha(d, m, y=None):
    if y is None:
        y = "2026"
    elif len(str(y)) == 2:
        y = "20" + str(y)
    return f"{int(d):02d}-{int(m):02d}-{y}"


def nivel_a_profundidad_inicio(nivel):
    try:
        nivel = int(nivel)
    except Exception:
        return "No informado"
    if nivel == 1:
        return "Superficie"
    return f"{nivel} ({(nivel - 1) * 10}-{nivel * 10} cm)"


def nivel_a_profundidad_termino(nivel):
    try:
        nivel = int(nivel)
    except Exception:
        return "No informado"
    return f"{nivel} ({(nivel - 1) * 10}-{nivel * 10} cm)"


def titulo_material(material):
    if pd.isna(material) or str(material).strip() == "":
        return "No informado"
    return str(material).strip().capitalize()


# ==========================================================
# NORMALIZACIÓN DEL TEXTO
# ==========================================================

def preparar_texto(texto):
    texto = texto.replace("•", "-").replace("·", "-")
    texto = texto.replace("–", "-").replace("—", "-")
    texto = texto.replace("\u2060", "")
    texto = texto.replace("\u202f", " ")

    # Encabezado WhatsApp: [18/5, 4:29 p.m.] Nombre:
    texto = re.sub(
        r"\[(\d{1,2})/(\d{1,2}),[^\]]+\]\s*([^:]+):",
        lambda m: f"\nFECHA_WHATSAPP {normalizar_fecha(m.group(1), m.group(2))}\n",
        texto
    )

    # Fecha explícita: 18-05-26 / 18/05 / 18/05/2026
    texto = re.sub(
        r"(?<!\d)(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?(?!\d)",
        lambda m: f"\nFECHA_REPORTE {normalizar_fecha(m.group(1), m.group(2), m.group(3))}\n",
        texto
    )

    # Unidad explícita
    texto = re.sub(
        r"\bUnidad\s+([0-9]+-[a-z]+[0-9]*|[0-9]+[a-z]?)\b",
        r"\nUNIDAD \1\n",
        texto,
        flags=re.I
    )

    # Unidad sin palabra Unidad: 9-r41 / 8-re / 7a / 7b
    texto = re.sub(
        r"(?<!NIVEL\s)(?<!Nivel\s)(?<!N)(?<!N\s)\b([0-9]+-[a-z]+[0-9]*|[0-9]+[a-z])\b",
        r"\nUNIDAD \1\n",
        texto,
        flags=re.I
    )

    # Nivel con capa explícita: Nivel 31 - Capa G / Nivel 45, Capa G
    texto = re.sub(
        r"\bNivel\s+(\d{1,2})\s*[-:.,]?\s*Capa\s*([A-Za-z]{1,4})\b",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I
    )

    # N31, G
    texto = re.sub(
        r"\bN\s*(\d{1,2})\s*,\s*([A-Za-z]{1,4})\b",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I
    )

    # 1, RC / 4, A / 11, RD
    texto = re.sub(
        r"(?<!\d)\b(\d{1,2})\s*,\s*([A-Za-z]{1,4})(?![A-Za-z])",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I
    )

    # Nivel sin capa: Nivel 37, queda picado
    texto = re.sub(
        r"\bNivel\s+(\d{1,2})\b(?!\s*,)(?!\s*[-:.,]?\s*Capa)",
        r"\nNIVEL \1\n",
        texto,
        flags=re.I
    )

    # Cronologías. Reconoce también Histórico: 0 / Subactual: 0
    texto = re.sub(
        r"(?:^|\s)-?\s*(Material arqueológico|Material arqueologico|Histórico|Historico)\s*(?:\(?n\s*=\s*\d+\)?|=\s*\d+|:\s*\d+|\d+)?\s*:??",
        r"\nCRONO histórico\n",
        texto,
        flags=re.I
    )

    texto = re.sub(
        r"(?:^|\s)-?\s*(Material subactual|Subactual)\s*(?:\(?n\s*=\s*\d+\)?|=\s*\d+|:\s*\d+)?\s*:??",
        r"\nCRONO subactual\n",
        texto,
        flags=re.I
    )

    texto = re.sub(r"\s+-\s*", r"\n- ", texto)

    lineas = []
    for linea in texto.splitlines():
        linea = linea.strip()
        if linea:
            lineas.append(re.sub(r"\s+", " ", linea))
    return lineas


# ==========================================================
# EXTRACCIÓN DE MATERIALES
# ==========================================================

def extraer_items_materiales(linea):
    original = linea
    linea = linea.strip().lstrip("- ").strip()
    resultados = []

    # material: número / material:
    m = re.match(r"^([a-záéíóúñü ]+?)\s*:\s*(\d+)?\s*$", linea, flags=re.I)
    if m:
        mat, nuevo = material_con_alerta(m.group(1))
        if mat:
            cant = int(m.group(2)) if m.group(2) else None
            return [(mat, cant, original, nuevo)]

    # material (n)
    m = re.match(r"^([a-záéíóúñü ]+?)\s*\((\d+)\)", linea, flags=re.I)
    if m:
        mat, nuevo = material_con_alerta(m.group(1))
        if mat:
            return [(mat, int(m.group(2)), original, nuevo)]

    # material número: loza 7
    m = re.match(r"^([a-záéíóúñü ]+?)\s+(\d+)\b", linea, flags=re.I)
    if m:
        mat, nuevo = material_con_alerta(m.group(1))
        if mat:
            return [(mat, int(m.group(2)), original, nuevo)]

    # número material / varios en una línea: 1 loza, 3 vidrios
    for cant, mat_raw in re.findall(r"(\d+)\s+([a-záéíóúñü]+)", linea, flags=re.I):
        mat, nuevo = material_con_alerta(mat_raw)
        if mat:
            resultados.append((mat, int(cant), original, nuevo))

    return resultados


# ==========================================================
# PARSER GENERAL
# ==========================================================

def parsear_todo(texto):
    lineas = preparar_texto(texto)
    registros = []
    alertas = []
    cronologias_declaradas = []

    fecha_actual = None
    fecha_whatsapp = None
    unidad_actual = None
    nivel_actual = None
    estrato_actual = None
    cronologia_actual = None

    for linea in lineas:
        if linea.startswith("FECHA_WHATSAPP"):
            fecha_whatsapp = linea.replace("FECHA_WHATSAPP", "").strip()
            if fecha_actual is None:
                fecha_actual = fecha_whatsapp
            continue

        if linea.startswith("FECHA_REPORTE"):
            fecha_actual = linea.replace("FECHA_REPORTE", "").strip()
            continue

        if linea.startswith("UNIDAD"):
            unidad_actual = linea.replace("UNIDAD", "").strip().lower()
            nivel_actual = None
            estrato_actual = None
            cronologia_actual = None
            if fecha_actual is None:
                fecha_actual = fecha_whatsapp
            continue

        if linea.startswith("NIVEL"):
            resto = linea.replace("NIVEL", "").strip()
            m = re.match(r"(\d{1,2})\s*,\s*([A-Za-z]{1,4})", resto)
            if m:
                nivel_actual = int(m.group(1))
                estrato_actual = m.group(2).upper()
                cronologia_actual = None
            else:
                m = re.match(r"(\d{1,2})", resto)
                if m:
                    nivel_actual = int(m.group(1))
                    estrato_actual = None
                    cronologia_actual = None
                    alertas.append(f"Estrato pendiente: unidad {unidad_actual or 'No informado'}, nivel {nivel_actual}")
            continue

        if linea.startswith("CRONO"):
            cronologia_actual = linea.replace("CRONO", "").strip().lower()
            if unidad_actual is None or nivel_actual is None:
                alertas.append(f"Cronología sin unidad/nivel: {linea}")
            else:
                cronologias_declaradas.append({
                    "fecha": fecha_actual or fecha_whatsapp or "No informado",
                    "unidad": unidad_actual or "No informado",
                    "nivel": nivel_actual,
                    "estrato": estrato_actual,
                    "cronologia": cronologia_actual,
                })
            continue

        items = extraer_items_materiales(linea)
        if items:
            for material, cantidad, original, material_nuevo in items:
                unidad_para_registro = unidad_actual or "No informado"
                if unidad_actual is None:
                    alertas.append(f"Material sin unidad: {original}")

                if nivel_actual is None:
                    alertas.append(f"Material sin nivel: unidad {unidad_para_registro}, {original}")
                    continue

                cronologia_para_registro = cronologia_actual or "No informado"
                if cronologia_actual is None:
                    alertas.append(f"Material sin cronología: unidad {unidad_para_registro}, nivel {nivel_actual}, {original}")

                if cantidad is None:
                    alertas.append(f"Cantidad pendiente: unidad {unidad_para_registro}, nivel {nivel_actual}, {material}")

                if material_nuevo:
                    alertas.append(f"Material nuevo no catastrado: '{material}' en unidad {unidad_para_registro}, nivel {nivel_actual}")

                registros.append(Registro(
                    fecha=fecha_actual or fecha_whatsapp or "No informado",
                    unidad=unidad_para_registro,
                    nivel=nivel_actual,
                    estrato=estrato_actual,
                    cronologia=cronologia_para_registro,
                    material=material,
                    cantidad=cantidad,
                    material_nuevo=material_nuevo,
                    linea_original=original,
                ))

        if "conteo pendiente" in strip_accents_lower(linea):
            alertas.append(f"Conteo pendiente: unidad {unidad_actual or 'No informado'}, nivel {nivel_actual or 'No informado'}")

    # Crea filas para cronologías/niveles declarados sin materiales, por ejemplo Histórico: 0.
    claves_con_material = {
        (r.fecha, r.unidad, int(r.nivel), r.estrato, r.cronologia)
        for r in registros
        if r.nivel is not None
    }

    for c in cronologias_declaradas:
        clave = (c["fecha"], c["unidad"], int(c["nivel"]), c["estrato"], c["cronologia"])
        if clave not in claves_con_material:
            registros.append(Registro(
                fecha=clave[0],
                unidad=clave[1],
                nivel=clave[2],
                estrato=clave[3],
                cronologia=clave[4],
                material=MATERIAL_SIN_MATERIAL,
                cantidad=0,
                material_nuevo=False,
                linea_original="Cronología declarada sin materiales",
            ))

    return registros, alertas


# ==========================================================
# RESUMEN Y VALIDACIONES
# ==========================================================

def registros_a_df(registros):
    columnas = ["fecha", "unidad", "nivel", "estrato", "cronologia", "material", "cantidad", "material_nuevo", "linea_original"]
    if not registros:
        return pd.DataFrame(columns=columnas)
    return pd.DataFrame([vars(r) for r in registros])[columnas]


def limpiar_df_editado(df):
    columnas = ["fecha", "unidad", "nivel", "estrato", "cronologia", "material", "cantidad", "material_nuevo", "linea_original"]
    if df is None or df.empty:
        return pd.DataFrame(columns=columnas)

    df = df.copy()
    for col in columnas:
        if col not in df.columns:
            df[col] = None
    df = df[columnas]

    df["fecha"] = df["fecha"].fillna("No informado").astype(str).str.strip()
    df["unidad"] = df["unidad"].fillna("No informado").astype(str).str.lower().str.strip()
    df["estrato"] = df["estrato"].fillna("No informado").astype(str).str.upper().str.strip()
    df.loc[df["estrato"].isin(["", "NONE", "NAN"]), "estrato"] = "No informado"
    df["cronologia"] = df["cronologia"].fillna("No informado").astype(str).str.lower().str.strip()
    df["material"] = df["material"].fillna(MATERIAL_SIN_MATERIAL).astype(str).str.lower().str.strip()
    df["nivel"] = pd.to_numeric(df["nivel"], errors="coerce")
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce")
    df["material_nuevo"] = df["material_nuevo"].fillna(False).astype(bool)
    df["linea_original"] = df["linea_original"].fillna("").astype(str)

    df = df.dropna(subset=["nivel"])
    if not df.empty:
        df["nivel"] = df["nivel"].astype(int)
    return df


def nombre_material_descripcion(material):
    if material == "baldosa":
        return "Material constructivo"
    return titulo_material(material)


def nombre_material_frecuencia(material):
    nombres = {
        "loza": "Loza",
        "metal": "Metal",
        "baldosa": "baldosa",
        "cerámica": "Cerámica",
        "vidrio": "Vidrio",
        "osteofauna": "Osteofauna",
        "misceláneo": "Misceláneo",
    }
    return nombres.get(str(material), titulo_material(material))


def resumen_por_grupo_df(df):
    df = limpiar_df_editado(df)
    if df.empty:
        return pd.DataFrame()

    resumenes = []

    for (fecha, unidad), grupo in df.groupby(["fecha", "unidad"], dropna=False):
        niveles = sorted(grupo["nivel"].dropna().astype(int).unique())
        if not niveles:
            continue

        df_con_material = grupo[
            (grupo["material"] != MATERIAL_SIN_MATERIAL)
            & (grupo["cantidad"].notna())
            & (grupo["cantidad"] > 0)
        ]

        historicos = grupo[
            (grupo["cronologia"] == "histórico")
            & (grupo["material"] != MATERIAL_SIN_MATERIAL)
            & (grupo["cantidad"].notna())
            & (grupo["cantidad"] > 0)
        ]

        # Estratigrafía: muestra estratos declarados aunque tengan N=0.
        por_estrato = defaultdict(int)
        estratos_declarados = [
            e for e in grupo["estrato"].dropna().unique()
            if str(e).strip() not in {"", "No informado", "NONE", "NAN"}
        ]
        for e in estratos_declarados:
            por_estrato[e] += 0
        for _, r in historicos.iterrows():
            estrato = r["estrato"] if r["estrato"] != "No informado" else "SIN ESTRATO"
            por_estrato[estrato] += int(r["cantidad"])

        col_j = " ".join(f"{estrato} (N={total})" for estrato, total in por_estrato.items()) or "No informado"

        materiales_presentes = {
            m for m in df_con_material["material"].dropna().unique()
            if m not in MATERIALES_EXCLUIDOS_DESCRIPCION
        }
        orden_desc = ORDEN_DESCRIPCION + sorted([m for m in materiales_presentes if m not in ORDEN_DESCRIPCION])
        col_l = ", ".join(nombre_material_descripcion(m) for m in orden_desc if m in materiales_presentes) or "No informado"

        cronologias = {c for c in grupo["cronologia"].dropna().unique() if c not in {"", "No informado"}}
        if cronologias == {"histórico", "subactual"}:
            col_m = "Histórico / Subactual"
        elif cronologias == {"histórico"}:
            col_m = "Histórico"
        elif cronologias == {"subactual"}:
            col_m = "Subactual"
        else:
            col_m = "No informado"

        # Frecuencia históricos: todo lo histórico separado, incluyendo misceláneo.
        por_material_hist = defaultdict(int)
        for _, r in historicos.iterrows():
            por_material_hist[r["material"]] += int(r["cantidad"])
        orden_freq = ORDEN_FRECUENCIA + sorted([m for m in por_material_hist if m not in ORDEN_FRECUENCIA])
        col_n = ", ".join(
            f"{nombre_material_frecuencia(m)} (N={por_material_hist[m]})"
            for m in orden_freq
            if m in por_material_hist
        ) or "No informado"

        resumenes.append({
            "Fecha": fecha,
            "Unidad": unidad,
            "Nivel de inicio": nivel_a_profundidad_inicio(min(niveles)),
            "Nivel de término": nivel_a_profundidad_termino(max(niveles)),
            "Total niveles excavados": len(niveles),
            "Nivel de cierre": "",
            "Estratigrafía": col_j,
            "Material cultural": "si" if not df_con_material.empty else "no",
            "Descripción general": col_l,
            "Cronología": col_m,
            "Frecuencia históricos": col_n,
        })

    return pd.DataFrame(resumenes)


def validar_df(df, alertas_parser=None):
    df = limpiar_df_editado(df)
    alertas = list(alertas_parser or [])

    if df.empty:
        return pd.DataFrame({"Alerta": ["No hay datos extraídos ni ingresados manualmente."]})

    for idx, r in df.iterrows():
        fila = idx + 1
        if r["fecha"] == "No informado":
            alertas.append(f"Fila editable {fila}: fecha no informada.")
        if r["unidad"] in {"", "No informado"}:
            alertas.append(f"Fila editable {fila}: unidad no informada.")
        if r["estrato"] == "No informado":
            alertas.append(f"Fila editable {fila}: estrato/capa no informado.")
        if r["cronologia"] == "No informado":
            alertas.append(f"Fila editable {fila}: cronología no informada.")
        if r["material"] != MATERIAL_SIN_MATERIAL and pd.isna(r["cantidad"]):
            alertas.append(f"Fila editable {fila}: cantidad no informada para material '{r['material']}'.")
        if bool(r["material_nuevo"]):
            alertas.append(f"Fila editable {fila}: material nuevo/no catastrado '{r['material']}'.")

    if not alertas:
        return pd.DataFrame({"Estado": ["Sin alertas"]})
    return pd.DataFrame({"Alerta": alertas})


# ==========================================================
# EXPORTACIÓN
# ==========================================================

def crear_excel_en_memoria(df_resumen, df_detalle, df_alertas):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resumen.to_excel(writer, sheet_name="Para copiar", index=False)
        df_detalle.to_excel(writer, sheet_name="Detalle editable", index=False)
        df_alertas.to_excel(writer, sheet_name="Alertas", index=False)
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = min(max_len + 2, 70)
    output.seek(0)
    return output.getvalue()


def crear_respaldo_zip(texto_original, df_resumen, df_detalle, df_alertas):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output = BytesIO()
    excel_bytes = crear_excel_en_memoria(df_resumen, df_detalle, df_alertas)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"{timestamp}_reporte_para_copiar.xlsx", excel_bytes)
        zip_file.writestr(f"{timestamp}_texto_original.txt", texto_original)
        zip_file.writestr(f"{timestamp}_resumen_para_copiar.csv", df_resumen.to_csv(index=False).encode("utf-8-sig"))
        zip_file.writestr(f"{timestamp}_detalle_editable.csv", df_detalle.to_csv(index=False).encode("utf-8-sig"))
        zip_file.writestr(f"{timestamp}_alertas.csv", df_alertas.to_csv(index=False).encode("utf-8-sig"))
    output.seek(0)
    return output.getvalue(), timestamp


# ==========================================================
# INTERFAZ
# ==========================================================

if "texto_reportes" not in st.session_state:
    st.session_state["texto_reportes"] = ""
if "resultado_generado" not in st.session_state:
    st.session_state["resultado_generado"] = False
if "df_extraido" not in st.session_state:
    st.session_state["df_extraido"] = pd.DataFrame()
if "alertas_parser" not in st.session_state:
    st.session_state["alertas_parser"] = []


def borrar_todo():
    st.session_state["texto_reportes"] = ""
    st.session_state["resultado_generado"] = False
    st.session_state["df_extraido"] = pd.DataFrame()
    st.session_state["alertas_parser"] = []


st.text_area(
    "Pega aquí uno o varios reportes de WhatsApp",
    height=380,
    key="texto_reportes",
    placeholder="""Ejemplo:
[28/5, 16:10] Persona: Unidad 7c
Nivel 45, Capa G
Subactual: 0
Histórico:
- loza: 2

Nivel 46, Capa G
Subactual: 0
Histórico: 0"""
)

col1, col2 = st.columns([1, 1])
with col1:
    procesar = st.button("Procesar reportes", type="primary")
with col2:
    st.button("Borrar resultados", on_click=borrar_todo)

if procesar:
    texto = st.session_state["texto_reportes"]
    if not texto.strip():
        st.warning("Debes pegar al menos un reporte.")
    else:
        registros, alertas_parser = parsear_todo(texto)
        st.session_state["df_extraido"] = registros_a_df(registros)
        st.session_state["alertas_parser"] = alertas_parser
        st.session_state["resultado_generado"] = True

if st.session_state["resultado_generado"]:
    st.subheader("1. Revisar y editar datos extraídos")
    st.info(
        "Puedes corregir campos, agregar filas nuevas o eliminar filas antes de descargar. "
        "Si falta información, puedes escribirla manualmente o dejarla como 'No informado'."
    )

    df_base = st.session_state["df_extraido"].copy()
    columnas_editor = ["fecha", "unidad", "nivel", "estrato", "cronologia", "material", "cantidad", "material_nuevo", "linea_original"]
    for col in columnas_editor:
        if col not in df_base.columns:
            df_base[col] = None
    df_base = df_base[columnas_editor]

    df_editado = st.data_editor(
        df_base,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "fecha": st.column_config.TextColumn("Fecha"),
            "unidad": st.column_config.TextColumn("Unidad"),
            "nivel": st.column_config.NumberColumn("Nivel", step=1),
            "estrato": st.column_config.TextColumn("Estrato/Capa"),
            "cronologia": st.column_config.SelectboxColumn("Cronología", options=["histórico", "subactual", "No informado"]),
            "material": st.column_config.TextColumn("Material"),
            "cantidad": st.column_config.NumberColumn("Cantidad", step=1),
            "material_nuevo": st.column_config.CheckboxColumn("Material nuevo"),
            "linea_original": st.column_config.TextColumn("Línea original"),
        },
        key="editor_detalle"
    )

    df_editado_limpio = limpiar_df_editado(df_editado)
    df_resumen = resumen_por_grupo_df(df_editado_limpio)
    df_alertas = validar_df(df_editado_limpio, st.session_state.get("alertas_parser", []))

    st.subheader("2. Reporte para copiar")
    if df_resumen.empty:
        st.error("No se pudo generar resumen. Revisa o completa la tabla editable.")
    else:
        st.dataframe(df_resumen, use_container_width=True)

        excel_bytes = crear_excel_en_memoria(df_resumen, df_editado_limpio, df_alertas)
        col_descarga1, col_descarga2 = st.columns([1, 1])

        with col_descarga1:
            st.download_button(
                label="Descargar Excel",
                data=excel_bytes,
                file_name="reporte_para_copiar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_descarga2:
            respaldo_bytes, timestamp = crear_respaldo_zip(
                texto_original=st.session_state["texto_reportes"],
                df_resumen=df_resumen,
                df_detalle=df_editado_limpio,
                df_alertas=df_alertas
            )
            st.download_button(
                label="Descargar respaldo completo ZIP",
                data=respaldo_bytes,
                file_name=f"respaldo_reporte_{timestamp}.zip",
                mime="application/zip"
            )

    with st.expander("Ver alertas y validaciones"):
        st.dataframe(df_alertas, use_container_width=True)
