import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Reporte WhatsApp Arqueología",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Generador de reporte desde WhatsApp")
st.write(
    "Pega los reportes de WhatsApp, procesa la información y descarga un Excel "
    "con una hoja lista para copiar a la planilla maestra."
)


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
    "vidrio", "loza", "metal", "osteofauna",
    "cerámica", "plástico", "baldosa"
]

ORDEN_FRECUENCIA = [
    "loza", "metal", "baldosa", "cerámica",
    "vidrio", "osteofauna"
]


@dataclass
class Registro:
    fecha: Optional[str]
    unidad: Optional[str]
    nivel: int
    estrato: Optional[str]
    cronologia: str
    material: str
    cantidad: Optional[int]
    linea_original: str = ""


def strip_accents_lower(text):
    text = str(text).strip().lower()
    return "".join(
        c for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )


def normalizar_material(raw):
    key = strip_accents_lower(raw)
    key = re.sub(r"[^a-zñ ]", "", key).strip()
    return MATERIAL_MAP.get(key)


def normalizar_fecha(d, m, y=None):
    if y is None:
        y = "2026"
    elif len(str(y)) == 2:
        y = "20" + str(y)
    return f"{int(d):02d}-{int(m):02d}-{y}"


def nivel_a_profundidad_inicio(nivel):
    if nivel == 1:
        return "Superficie"
    return f"{nivel} ({(nivel - 1) * 10}-{nivel * 10} cm)"


def nivel_a_profundidad_termino(nivel):
    return f"{nivel} ({(nivel - 1) * 10}-{nivel * 10} cm)"


def preparar_texto(texto):
    texto = texto.replace("•", "-").replace("·", "-")
    texto = texto.replace("–", "-").replace("—", "-")
    texto = texto.replace("\u2060", "")
    texto = texto.replace("\u202f", " ")

    texto = re.sub(
        r"\[(\d{1,2})/(\d{1,2}),[^\]]+\]\s*([^:]+):",
        lambda m: f"\nFECHA_WHATSAPP {normalizar_fecha(m.group(1), m.group(2))}\n",
        texto
    )

    texto = re.sub(
        r"(?<!\d)(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?(?!\d)",
        lambda m: f"\nFECHA_REPORTE {normalizar_fecha(m.group(1), m.group(2), m.group(3))}\n",
        texto
    )

    texto = re.sub(
        r"\bUnidad\s+([0-9]+-[a-z]+[0-9]*|[0-9]+[a-z]?)\b",
        r"\nUNIDAD \1\n",
        texto,
        flags=re.I
    )

    texto = re.sub(
        r"(?<!NIVEL\s)(?<!Nivel\s)(?<!N)(?<!N\s)\b([0-9]+-[a-z]+[0-9]*|[0-9]+[a-z])\b",
        r"\nUNIDAD \1\n",
        texto,
        flags=re.I
    )

    texto = re.sub(
        r"\bNivel\s+(\d{1,2})\s*[-:.]?\s*Capa\s*([A-Za-z]{1,4})\b",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I
    )

    texto = re.sub(
        r"\bN\s*(\d{1,2})\s*,\s*([A-Za-z]{1,4})\b",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I
    )

    texto = re.sub(
        r"(?<!\d)\b(\d{1,2})\s*,\s*([A-Za-z]{1,4})(?![A-Za-z])",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I
    )

    texto = re.sub(
        r"\bNivel\s+(\d{1,2})\b(?!\s*,)(?!\s*[-:.]?\s*Capa)",
        r"\nNIVEL \1\n",
        texto,
        flags=re.I
    )

    texto = re.sub(
        r"(?:^|\s)-?\s*(Material arqueológico|Material arqueologico|Histórico|Historico)\s*(?:\(?n\s*=\s*\d+\)?|=\s*\d+|\d+)?\s*:?",
        r"\nCRONO histórico\n",
        texto,
        flags=re.I
    )

    texto = re.sub(
        r"(?:^|\s)-?\s*(Material subactual|Subactual)\s*(?:\(?n\s*=\s*\d+\)?|=\s*\d+)?\s*:?",
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


def extraer_items_materiales(linea):
    original = linea
    linea = linea.strip().lstrip("- ").strip()
    resultados = []

    m = re.match(r"^([a-záéíóúñü ]+?)\s*:\s*(\d+)?", linea, flags=re.I)
    if m:
        mat = normalizar_material(m.group(1))
        if mat:
            cant = int(m.group(2)) if m.group(2) else None
            return [(mat, cant, original)]

    m = re.match(r"^([a-záéíóúñü ]+?)\s*\((\d+)\)", linea, flags=re.I)
    if m:
        mat = normalizar_material(m.group(1))
        if mat:
            return [(mat, int(m.group(2)), original)]

    m = re.match(r"^([a-záéíóúñü ]+?)\s+(\d+)\b", linea, flags=re.I)
    if m:
        mat = normalizar_material(m.group(1))
        if mat:
            return [(mat, int(m.group(2)), original)]

    for cant, mat_raw in re.findall(r"(\d+)\s+([a-záéíóúñü]+)", linea, flags=re.I):
        mat = normalizar_material(mat_raw)
        if mat:
            resultados.append((mat, int(cant), original))

    return resultados


def parsear_todo(texto):
    lineas = preparar_texto(texto)

    registros = []
    alertas = []

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
                    alertas.append(
                        f"Estrato pendiente: unidad {unidad_actual}, nivel {nivel_actual}"
                    )
            continue

        if linea.startswith("CRONO"):
            cronologia_actual = linea.replace("CRONO", "").strip().lower()
            continue

        items = extraer_items_materiales(linea)

        if items:
            for material, cantidad, original in items:
                if unidad_actual is None:
                    alertas.append(f"Material sin unidad: {original}")
                    continue

                if nivel_actual is None:
                    alertas.append(f"Material sin nivel: unidad {unidad_actual}, {original}")
                    continue

                if cronologia_actual is None:
                    alertas.append(
                        f"Material sin cronología: unidad {unidad_actual}, nivel {nivel_actual}, {original}"
                    )
                    continue

                if cantidad is None:
                    alertas.append(
                        f"Cantidad pendiente: unidad {unidad_actual}, nivel {nivel_actual}, {material}"
                    )

                registros.append(
                    Registro(
                        fecha=fecha_actual or fecha_whatsapp,
                        unidad=unidad_actual,
                        nivel=nivel_actual,
                        estrato=estrato_actual,
                        cronologia=cronologia_actual,
                        material=material,
                        cantidad=cantidad,
                        linea_original=original
                    )
                )

        if "conteo pendiente" in strip_accents_lower(linea):
            alertas.append(f"Conteo pendiente: unidad {unidad_actual}, nivel {nivel_actual}")

    return registros, alertas


def nombre_material_descripcion(material):
    if material == "baldosa":
        return "Material constructivo"
    return material.capitalize()


def normalizar_material_frecuencia(material):
    if material == "misceláneo":
        return "loza"
    return material


def nombre_material_frecuencia(material):
    nombres = {
        "loza": "Loza",
        "metal": "Metal",
        "baldosa": "baldosa",
        "cerámica": "Cerámica",
        "vidrio": "Vidrio",
        "osteofauna": "Osteofauna"
    }
    return nombres.get(material, material.capitalize())


def resumen_por_grupo(registros):
    grupos = defaultdict(list)

    for r in registros:
        grupos[(r.fecha, r.unidad)].append(r)

    resumenes = []

    for (fecha, unidad), regs in grupos.items():
        validos = [r for r in regs if r.cantidad is not None]

        if not validos:
            continue

        niveles = sorted({r.nivel for r in validos})
        historicos = [r for r in validos if r.cronologia == "histórico"]

        por_estrato = defaultdict(int)
        for r in historicos:
            estrato = r.estrato if r.estrato else "SIN ESTRATO"
            por_estrato[estrato] += r.cantidad

        col_j = " ".join(
            f"{estrato} (N={total})"
            for estrato, total in por_estrato.items()
        )

        materiales_presentes = {
            r.material for r in validos
            if r.material not in {"malacológico", "misceláneo"}
        }

        col_l = ", ".join(
            nombre_material_descripcion(m)
            for m in ORDEN_DESCRIPCION
            if m in materiales_presentes
        )

        cronologias = {r.cronologia for r in validos}

        if cronologias == {"histórico", "subactual"}:
            col_m = "Histórico / Subactual"
        elif cronologias == {"histórico"}:
            col_m = "Histórico"
        elif cronologias == {"subactual"}:
            col_m = "Subactual"
        else:
            col_m = ""

        por_material_hist = defaultdict(int)

        for r in historicos:
            mat = normalizar_material_frecuencia(r.material)
            por_material_hist[mat] += r.cantidad

        col_n = ", ".join(
            f"{nombre_material_frecuencia(m)} (N={por_material_hist[m]})"
            for m in ORDEN_FRECUENCIA
            if m in por_material_hist
        )

        resumenes.append({
            "Fecha": fecha,
            "Unidad": unidad,
            "Nivel de inicio": nivel_a_profundidad_inicio(min(niveles)),
            "Nivel de término": nivel_a_profundidad_termino(max(niveles)),
            "Total niveles excavados": len(niveles),
            "Nivel de cierre": "",
            "Estratigrafía": col_j,
            "Material cultural": "si",
            "Descripción general": col_l,
            "Cronología": col_m,
            "Frecuencia históricos": col_n,
        })

    return pd.DataFrame(resumenes)


def registros_a_df(registros):
    if not registros:
        return pd.DataFrame()
    return pd.DataFrame([vars(r) for r in registros])


def crear_excel_en_memoria(df_resumen, df_detalle, df_alertas):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resumen.to_excel(writer, sheet_name="Para copiar", index=False)
        df_detalle.to_excel(writer, sheet_name="Detalle normalizado", index=False)
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


# ==========================================================
# INTERFAZ
# ==========================================================

if "texto_reportes" not in st.session_state:
    st.session_state["texto_reportes"] = ""

if "resultado_generado" not in st.session_state:
    st.session_state["resultado_generado"] = False

if "df_resumen" not in st.session_state:
    st.session_state["df_resumen"] = pd.DataFrame()

if "df_detalle" not in st.session_state:
    st.session_state["df_detalle"] = pd.DataFrame()

if "df_alertas" not in st.session_state:
    st.session_state["df_alertas"] = pd.DataFrame()


def borrar_todo():
    st.session_state["texto_reportes"] = ""
    st.session_state["resultado_generado"] = False
    st.session_state["df_resumen"] = pd.DataFrame()
    st.session_state["df_detalle"] = pd.DataFrame()
    st.session_state["df_alertas"] = pd.DataFrame()


st.text_area(
    "Pega aquí uno o varios reportes de WhatsApp",
    height=380,
    key="texto_reportes",
    placeholder="""Ejemplo:
[18/5, 4:37 p.m.] Valeria: 9-r41
1, RC
Subactual
- plástico: 6
Histórico
- cerámica: 4"""
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
        registros, alertas = parsear_todo(texto)

        df_resumen = resumen_por_grupo(registros)
        df_detalle = registros_a_df(registros)

        if alertas:
            df_alertas = pd.DataFrame({"Alerta": alertas})
        else:
            df_alertas = pd.DataFrame({"Estado": ["Sin alertas"]})

        st.session_state["df_resumen"] = df_resumen
        st.session_state["df_detalle"] = df_detalle
        st.session_state["df_alertas"] = df_alertas
        st.session_state["resultado_generado"] = True


if st.session_state["resultado_generado"]:
    df_resumen = st.session_state["df_resumen"]
    df_detalle = st.session_state["df_detalle"]
    df_alertas = st.session_state["df_alertas"]

    st.subheader("Reporte para copiar")

    if df_resumen.empty:
        st.error("No se pudo generar el resumen. Revisa si el texto contiene unidad, nivel, cronología y materiales.")
    else:
        st.dataframe(df_resumen, use_container_width=True)

        excel_bytes = crear_excel_en_memoria(df_resumen, df_detalle, df_alertas)

        st.download_button(
            label="Descargar Excel",
            data=excel_bytes,
            file_name="reporte_para_copiar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with st.expander("Ver detalle normalizado"):
        st.dataframe(df_detalle, use_container_width=True)

    with st.expander("Ver alertas"):
        st.dataframe(df_alertas, use_container_width=True)
