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
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Reporte WhatsApp Arqueología",
    page_icon="📋",
    layout="wide",
)


MATERIAL_MAP = {
    "loza": "loza",
    "lozas": "loza",
    "ceramica": "cerámica",
    "ceramicas": "cerámica",
    "cerámica": "cerámica",
    "cerámicas": "cerámica",
    "vidrio": "vidrio",
    "vidrios": "vidrio",
    "metal": "metal",
    "metales": "metal",
    "osteofauna": "osteofauna",
    "osteofaunas": "osteofauna",
    "plastico": "plástico",
    "plasticos": "plástico",
    "plástico": "plástico",
    "plásticos": "plástico",
    "malacologico": "malacológico",
    "malacologicos": "malacológico",
    "malacológico": "malacológico",
    "malacológicos": "malacológico",
    "miscelaneo": "misceláneo",
    "miscelaneos": "misceláneo",
    "misceláneo": "misceláneo",
    "misceláneos": "misceláneo",
    "baldosa": "baldosa",
    "baldosas": "baldosa",
}

ORDEN_DESCRIPCION = [
    "vidrio",
    "loza",
    "metal",
    "osteofauna",
    "cerámica",
    "plástico",
    "baldosa",
]

ORDEN_FRECUENCIA = [
    "loza",
    "metal",
    "baldosa",
    "cerámica",
    "vidrio",
    "osteofauna",
    "misceláneo",
]

ORDEN_CRONOLOGIAS = [
    "histórico",
    "prehispánico",
    "subactual",
]

MATERIAL_SIN_MATERIAL = "sin material"

MATERIALES_EXCLUIDOS_DESCRIPCION = {
    "malacológico",
    "misceláneo",
    MATERIAL_SIN_MATERIAL,
}

COLUMNAS_EDITOR = [
    "fecha",
    "unidad",
    "nivel",
    "estrato",
    "cronologia",
    "material",
    "cantidad",
    "material_nuevo",
    "linea_original",
]

COLUMNAS_RESUMEN = [
    "Fecha",
    "Unidad",
    "Nivel de inicio",
    "Nivel de término",
    "Total niveles excavados",
    "Nivel de cierre",
    "Estratigrafía",
    "Material cultural",
    "Descripción general",
    "Cronología",
    "Frecuencia históricos",
    "Frecuencia prehispánicos",
]

REPORTE_TIPO = """[dd/mm, hh:mm] Nombre:
dd/mm/aa

Unidad:

N__, capa __

Subactual: 0

Histórico:

Prehispánico:
"""

REPORTE_EJEMPLO = """[10/7, 6:43 p.m.] Valentina: 10/07/26

Unidad 18-4

N38, capa C

Subactual: 0

Histórico:
-1 osteofauna
-2 lozas
-1 metal

Prehispánico:
-1 cerámica

N39, capa C

Subactual: 0

Histórico:
-11 lozas
-15 vidrios
-10 osteofaunas

Pre-hispanico:
-4 cerámicas
"""


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


def strip_accents_lower(text):
    text = str(text).strip().lower()
    return "".join(
        char
        for char in unicodedata.normalize("NFD", text)
        if unicodedata.category(char) != "Mn"
    )


def limpiar_texto_material(raw):
    key = strip_accents_lower(raw)
    key = re.sub(r"[^a-zñáéíóúü ]", "", key).strip()
    return re.sub(r"\s+", " ", key)


def normalizar_material(raw):
    return MATERIAL_MAP.get(limpiar_texto_material(raw))


def normalizar_material_desconocido(raw):
    key = limpiar_texto_material(raw)

    if not key:
        return None

    # Evita convertir frases narrativas largas en materiales.
    if len(key.split()) > 3:
        return None

    return key


def material_con_alerta(raw):
    material = normalizar_material(raw)

    if material:
        return material, False

    material_nuevo = normalizar_material_desconocido(raw)

    if material_nuevo:
        return material_nuevo, True

    return None, False


def normalizar_fecha(dia, mes, anio=None):
    if anio is None:
        anio = "2026"
    elif len(str(anio)) == 2:
        anio = "20" + str(anio)

    return f"{int(dia):02d}-{int(mes):02d}-{anio}"


def nivel_a_profundidad_inicio(nivel):
    try:
        nivel = int(nivel)
    except (TypeError, ValueError):
        return "No informado"

    if nivel == 1:
        return "Superficie"

    return f"{nivel} ({(nivel - 1) * 10}-{nivel * 10} cm)"


def nivel_a_profundidad_termino(nivel):
    try:
        nivel = int(nivel)
    except (TypeError, ValueError):
        return "No informado"

    return f"{nivel} ({(nivel - 1) * 10}-{nivel * 10} cm)"


def titulo_material(material):
    if pd.isna(material) or str(material).strip() == "":
        return "No informado"

    return str(material).strip().capitalize()


def nombre_cronologia(cronologia):
    nombres = {
        "histórico": "Histórico",
        "prehispánico": "Prehispánico",
        "subactual": "Subactual",
    }
    return nombres.get(cronologia, str(cronologia).capitalize())


def preparar_texto(texto):
    texto = texto.replace("•", "-").replace("·", "-")
    texto = texto.replace("–", "-").replace("—", "-")
    texto = texto.replace("\u2060", "")
    texto = texto.replace("\u202f", " ")

    # Se usa un token temporal para evitar que la fecha del encabezado
    # de WhatsApp sea procesada otra vez por la regla de fecha explícita.
    def reemplazar_encabezado_whatsapp(match):
        dia = int(match.group(1))
        mes = int(match.group(2))
        return f"\n__FWH_2026_{mes:02d}_{dia:02d}__\n"

    texto = re.sub(
        r"\[(\d{1,2})/(\d{1,2}),[^\]]+\]\s*([^:]+):",
        reemplazar_encabezado_whatsapp,
        texto,
    )

    # Fecha explícita: 18-05-26 / 18/05 / 18/05/2026.
    texto = re.sub(
        r"(?<!\d)(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?(?!\d)",
        lambda match: (
            "\nFECHA_REPORTE "
            + normalizar_fecha(
                match.group(1),
                match.group(2),
                match.group(3),
            )
            + "\n"
        ),
        texto,
    )

    texto = re.sub(
        r"__FWH_(\d{4})_(\d{2})_(\d{2})__",
        lambda match: (
            f"FECHA_WHATSAPP "
            f"{match.group(3)}-{match.group(2)}-{match.group(1)}"
        ),
        texto,
    )

    # Unidad explícita.
    texto = re.sub(
        r"\bUnidad\s+([0-9]+-[a-z]+[0-9]*|[0-9]+[a-z]?)\b",
        r"\nUNIDAD \1\n",
        texto,
        flags=re.I,
    )

    # Unidad sin la palabra "Unidad": 9-r41 / 8-re / 7a.
    texto = re.sub(
        r"(?<!NIVEL\s)(?<!Nivel\s)(?<!N)(?<!N\s)"
        r"\b([0-9]+-[a-z]+[0-9]*|[0-9]+[a-z])\b",
        r"\nUNIDAD \1\n",
        texto,
        flags=re.I,
    )

    # Nivel 31 - Capa G / Nivel 45, Capa G.
    texto = re.sub(
        r"\bNivel\s+(\d{1,2})\s*[-:.,]?\s*Capa\s*([A-Za-z]{1,4})\b",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I,
    )

    # N37, capa C.
    texto = re.sub(
        r"\bN\s*(\d{1,2})\s*,\s*Capa\s*([A-Za-z]{1,4})\b",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I,
    )

    # N31, G.
    texto = re.sub(
        r"\bN\s*(\d{1,2})\s*,\s*([A-Za-z]{1,4})\b",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I,
    )

    # 1, RC / 4, A / 11, RD.
    texto = re.sub(
        r"(?<!\d)\b(\d{1,2})\s*,\s*([A-Za-z]{1,4})(?![A-Za-z])",
        r"\nNIVEL \1, \2\n",
        texto,
        flags=re.I,
    )

    # Nivel sin capa: Nivel 37, queda picado.
    texto = re.sub(
        r"\bNivel\s+(\d{1,2})\b"
        r"(?!\s*,)(?!\s*[-:.,]?\s*Capa)",
        r"\nNIVEL \1\n",
        texto,
        flags=re.I,
    )

    # Histórico y Material arqueológico.
    texto = re.sub(
        r"(?:^|\s)-?\s*"
        r"(Material arqueológico|Material arqueologico|Histórico|Historico)"
        r"\s*(?:\(?n\s*=\s*\d+\)?|=\s*\d+|:\s*\d+|\d+)?\s*:?",
        r"\nCRONO histórico\n",
        texto,
        flags=re.I,
    )

    # Prehispánico y variantes.
    texto = re.sub(
        r"(?:^|\s)-?\s*"
        r"(Pre[\s-]*hisp[aá]nico)"
        r"\s*(?:\(?n\s*=\s*\d+\)?|=\s*\d+|:\s*\d+|\d+)?\s*:?",
        r"\nCRONO prehispánico\n",
        texto,
        flags=re.I,
    )

    # Subactual y Material subactual.
    texto = re.sub(
        r"(?:^|\s)-?\s*"
        r"(Material subactual|Subactual)"
        r"\s*(?:\(?n\s*=\s*\d+\)?|=\s*\d+|:\s*\d+|\d+)?\s*:?",
        r"\nCRONO subactual\n",
        texto,
        flags=re.I,
    )

    # Separa ítems con guion incluso cuando vienen como "-1 loza".
    texto = re.sub(r"\s*-\s*(?=\d|[A-Za-zÁÉÍÓÚÜÑáéíóúüñ])", r"\n- ", texto)

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

    # material: número / material:
    match = re.match(
        r"^([a-záéíóúñü ]+?)\s*:\s*(\d+)?\s*$",
        linea,
        flags=re.I,
    )

    if match:
        material, nuevo = material_con_alerta(match.group(1))

        if material:
            cantidad = int(match.group(2)) if match.group(2) else None
            return [(material, cantidad, original, nuevo)]

    # material (n)
    match = re.match(
        r"^([a-záéíóúñü ]+?)\s*\((\d+)\)",
        linea,
        flags=re.I,
    )

    if match:
        material, nuevo = material_con_alerta(match.group(1))

        if material:
            return [(material, int(match.group(2)), original, nuevo)]

    # material número: loza 7.
    match = re.match(
        r"^([a-záéíóúñü ]+?)\s+(\d+)\b",
        linea,
        flags=re.I,
    )

    if match:
        material, nuevo = material_con_alerta(match.group(1))

        if material:
            return [(material, int(match.group(2)), original, nuevo)]

    # número material / varios en una línea: 1 loza, 3 vidrios.
    for cantidad, material_raw in re.findall(
        r"(\d+)\s+([a-záéíóúñü]+)",
        linea,
        flags=re.I,
    ):
        material, nuevo = material_con_alerta(material_raw)

        if material:
            resultados.append(
                (material, int(cantidad), original, nuevo)
            )

    return resultados


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
            fecha_whatsapp = linea.replace(
                "FECHA_WHATSAPP",
                "",
            ).strip()

            # Cada mensaje nuevo reinicia la fecha de contexto.
            fecha_actual = fecha_whatsapp
            continue

        if linea.startswith("FECHA_REPORTE"):
            fecha_actual = linea.replace(
                "FECHA_REPORTE",
                "",
            ).strip()
            continue

        if linea.startswith("UNIDAD"):
            unidad_actual = linea.replace(
                "UNIDAD",
                "",
            ).strip().lower()

            nivel_actual = None
            estrato_actual = None
            cronologia_actual = None
            continue

        if linea.startswith("NIVEL"):
            resto = linea.replace("NIVEL", "").strip()

            match = re.match(
                r"(\d{1,2})\s*,\s*([A-Za-z]{1,4})",
                resto,
            )

            if match:
                nivel_actual = int(match.group(1))
                estrato_actual = match.group(2).upper()
                cronologia_actual = None
            else:
                match = re.match(r"(\d{1,2})", resto)

                if match:
                    nivel_actual = int(match.group(1))
                    estrato_actual = None
                    cronologia_actual = None

                    alertas.append(
                        f"Estrato pendiente: unidad "
                        f"{unidad_actual or 'No informado'}, "
                        f"nivel {nivel_actual}"
                    )

            continue

        if linea.startswith("CRONO"):
            cronologia_actual = linea.replace(
                "CRONO",
                "",
            ).strip().lower()

            if unidad_actual is None or nivel_actual is None:
                alertas.append(
                    f"Cronología sin unidad/nivel: {linea}"
                )
            else:
                cronologias_declaradas.append(
                    {
                        "fecha": (
                            fecha_actual
                            or fecha_whatsapp
                            or "No informado"
                        ),
                        "unidad": unidad_actual or "No informado",
                        "nivel": nivel_actual,
                        "estrato": estrato_actual,
                        "cronologia": cronologia_actual,
                    }
                )

            continue

        items = extraer_items_materiales(linea)

        if items:
            for (
                material,
                cantidad,
                original,
                material_nuevo,
            ) in items:
                unidad_registro = unidad_actual or "No informado"

                if unidad_actual is None:
                    alertas.append(
                        f"Material sin unidad: {original}"
                    )

                if nivel_actual is None:
                    alertas.append(
                        f"Material sin nivel: unidad "
                        f"{unidad_registro}, {original}"
                    )
                    continue

                cronologia_registro = cronologia_actual or "No informado"

                if cronologia_actual is None:
                    alertas.append(
                        f"Material sin cronología: unidad "
                        f"{unidad_registro}, nivel "
                        f"{nivel_actual}, {original}"
                    )

                if cantidad is None:
                    alertas.append(
                        f"Cantidad pendiente: unidad "
                        f"{unidad_registro}, nivel "
                        f"{nivel_actual}, {material}"
                    )

                if material_nuevo:
                    alertas.append(
                        f"Material nuevo/no catastrado: "
                        f"'{material}' en unidad "
                        f"{unidad_registro}, nivel "
                        f"{nivel_actual}"
                    )

                registros.append(
                    Registro(
                        fecha=(
                            fecha_actual
                            or fecha_whatsapp
                            or "No informado"
                        ),
                        unidad=unidad_registro,
                        nivel=nivel_actual,
                        estrato=estrato_actual,
                        cronologia=cronologia_registro,
                        material=material,
                        cantidad=cantidad,
                        material_nuevo=material_nuevo,
                        linea_original=original,
                    )
                )

        if "conteo pendiente" in strip_accents_lower(linea):
            alertas.append(
                f"Conteo pendiente: unidad "
                f"{unidad_actual or 'No informado'}, "
                f"nivel {nivel_actual or 'No informado'}"
            )

    # Crea una fila placeholder para cronologías declaradas sin materiales:
    # Histórico: 0, Prehispánico: 0 o Subactual: 0.
    claves_con_material = {
        (
            registro.fecha,
            registro.unidad,
            int(registro.nivel),
            registro.estrato,
            registro.cronologia,
        )
        for registro in registros
        if registro.nivel is not None
    }

    for cronologia in cronologias_declaradas:
        clave = (
            cronologia["fecha"],
            cronologia["unidad"],
            int(cronologia["nivel"]),
            cronologia["estrato"],
            cronologia["cronologia"],
        )

        if clave not in claves_con_material:
            registros.append(
                Registro(
                    fecha=clave[0],
                    unidad=clave[1],
                    nivel=clave[2],
                    estrato=clave[3],
                    cronologia=clave[4],
                    material=MATERIAL_SIN_MATERIAL,
                    cantidad=0,
                    material_nuevo=False,
                    linea_original="Cronología declarada sin materiales",
                )
            )

    return registros, alertas


def registros_a_df(registros):
    if not registros:
        return pd.DataFrame(columns=COLUMNAS_EDITOR)

    return pd.DataFrame(
        [vars(registro) for registro in registros]
    )[COLUMNAS_EDITOR]


def limpiar_df_editado(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNAS_EDITOR)

    df = df.copy()

    for columna in COLUMNAS_EDITOR:
        if columna not in df.columns:
            df[columna] = None

    df = df[COLUMNAS_EDITOR]

    df["fecha"] = (
        df["fecha"]
        .fillna("No informado")
        .astype(str)
        .str.strip()
    )

    df["unidad"] = (
        df["unidad"]
        .fillna("No informado")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    df["estrato"] = (
        df["estrato"]
        .fillna("No informado")
        .astype(str)
        .str.upper()
        .str.strip()
    )

    df.loc[
        df["estrato"].isin(["", "NONE", "NAN"]),
        "estrato",
    ] = "No informado"

    df["cronologia"] = (
        df["cronologia"]
        .fillna("No informado")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    # Normaliza variantes editadas manualmente.
    df["cronologia"] = df["cronologia"].replace(
        {
            "historico": "histórico",
            "pre-hispanico": "prehispánico",
            "pre-hispánico": "prehispánico",
            "prehispanico": "prehispánico",
            "pre hispanico": "prehispánico",
            "pre hispánico": "prehispánico",
        }
    )

    df["material"] = (
        df["material"]
        .fillna(MATERIAL_SIN_MATERIAL)
        .astype(str)
        .str.lower()
        .str.strip()
    )

    df["nivel"] = pd.to_numeric(
        df["nivel"],
        errors="coerce",
    )

    df["cantidad"] = pd.to_numeric(
        df["cantidad"],
        errors="coerce",
    )

    df["material_nuevo"] = (
        df["material_nuevo"]
        .fillna(False)
        .astype(bool)
    )

    df["linea_original"] = (
        df["linea_original"]
        .fillna("")
        .astype(str)
    )

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

    return nombres.get(
        str(material),
        titulo_material(material),
    )


def formatear_frecuencia(df_cronologia):
    por_material = defaultdict(int)

    for _, fila in df_cronologia.iterrows():
        por_material[fila["material"]] += int(
            fila["cantidad"]
        )

    orden = ORDEN_FRECUENCIA + sorted(
        material
        for material in por_material
        if material not in ORDEN_FRECUENCIA
    )

    return ", ".join(
        f"{nombre_material_frecuencia(material)} "
        f"(N={por_material[material]})"
        for material in orden
        if material in por_material
    ) or "No informado"


def resumen_por_grupo_df(df):
    df = limpiar_df_editado(df)

    if df.empty:
        return pd.DataFrame(columns=COLUMNAS_RESUMEN)

    resumenes = []

    for (fecha, unidad), grupo in df.groupby(
        ["fecha", "unidad"],
        dropna=False,
    ):
        niveles = sorted(
            grupo["nivel"]
            .dropna()
            .astype(int)
            .unique()
        )

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

        prehispanicos = grupo[
            (grupo["cronologia"] == "prehispánico")
            & (grupo["material"] != MATERIAL_SIN_MATERIAL)
            & (grupo["cantidad"].notna())
            & (grupo["cantidad"] > 0)
        ]

        # Se mantiene la regla previa: Estratigrafía suma los históricos.
        por_estrato = defaultdict(int)

        estratos_declarados = [
            estrato
            for estrato in grupo["estrato"].dropna().unique()
            if str(estrato).strip()
            not in {
                "",
                "No informado",
                "NONE",
                "NAN",
            }
        ]

        for estrato in estratos_declarados:
            por_estrato[estrato] += 0

        for _, fila in historicos.iterrows():
            estrato = (
                fila["estrato"]
                if fila["estrato"] != "No informado"
                else "SIN ESTRATO"
            )

            por_estrato[estrato] += int(fila["cantidad"])

        col_estratigrafia = " ".join(
            f"{estrato} (N={total})"
            for estrato, total in por_estrato.items()
        ) or "No informado"

        materiales_presentes = {
            material
            for material in df_con_material["material"]
            .dropna()
            .unique()
            if material not in MATERIALES_EXCLUIDOS_DESCRIPCION
        }

        orden_descripcion = (
            ORDEN_DESCRIPCION
            + sorted(
                material
                for material in materiales_presentes
                if material not in ORDEN_DESCRIPCION
            )
        )

        col_descripcion = ", ".join(
            nombre_material_descripcion(material)
            for material in orden_descripcion
            if material in materiales_presentes
        ) or "No informado"

        cronologias_presentes = {
            cronologia
            for cronologia in grupo["cronologia"]
            .dropna()
            .unique()
            if cronologia not in {"", "No informado"}
        }

        col_cronologia = " / ".join(
            nombre_cronologia(cronologia)
            for cronologia in ORDEN_CRONOLOGIAS
            if cronologia in cronologias_presentes
        ) or "No informado"

        resumenes.append(
            {
                "Fecha": fecha,
                "Unidad": unidad,
                "Nivel de inicio": nivel_a_profundidad_inicio(
                    min(niveles)
                ),
                "Nivel de término": nivel_a_profundidad_termino(
                    max(niveles)
                ),
                "Total niveles excavados": len(niveles),
                "Nivel de cierre": "",
                "Estratigrafía": col_estratigrafia,
                "Material cultural": (
                    "si" if not df_con_material.empty else "no"
                ),
                "Descripción general": col_descripcion,
                "Cronología": col_cronologia,
                "Frecuencia históricos": formatear_frecuencia(
                    historicos
                ),
                "Frecuencia prehispánicos": formatear_frecuencia(
                    prehispanicos
                ),
            }
        )

    return pd.DataFrame(
        resumenes,
        columns=COLUMNAS_RESUMEN,
    )


def validar_df(df, alertas_parser=None):
    df = limpiar_df_editado(df)
    alertas = list(alertas_parser or [])

    if df.empty:
        return pd.DataFrame(
            {
                "Alerta": [
                    "No hay datos extraídos ni ingresados manualmente."
                ]
            }
        )

    for indice, fila in df.iterrows():
        numero_fila = indice + 1

        if fila["fecha"] == "No informado":
            alertas.append(
                f"Fila editable {numero_fila}: fecha no informada."
            )

        if fila["unidad"] in {"", "No informado"}:
            alertas.append(
                f"Fila editable {numero_fila}: unidad no informada."
            )

        if fila["estrato"] == "No informado":
            alertas.append(
                f"Fila editable {numero_fila}: "
                f"estrato/capa no informado."
            )

        if fila["cronologia"] == "No informado":
            alertas.append(
                f"Fila editable {numero_fila}: "
                f"cronología no informada."
            )

        if (
            fila["material"] != MATERIAL_SIN_MATERIAL
            and pd.isna(fila["cantidad"])
        ):
            alertas.append(
                f"Fila editable {numero_fila}: "
                f"cantidad no informada para material "
                f"'{fila['material']}'."
            )

        if bool(fila["material_nuevo"]):
            alertas.append(
                f"Fila editable {numero_fila}: "
                f"material nuevo/no catastrado "
                f"'{fila['material']}'."
            )

    if not alertas:
        return pd.DataFrame({"Estado": ["Sin alertas"]})

    return pd.DataFrame({"Alerta": alertas})


def aplicar_formato_excel(writer, nombre_hoja):
    worksheet = writer.book[nombre_hoja]

    relleno_encabezado = PatternFill(
        "solid",
        fgColor="D9EAF7",
    )

    fuente_encabezado = Font(bold=True)

    alineacion_encabezado = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )

    # Cuerpo: alineación horizontal izquierda y vertical centrada.
    alineacion_cuerpo = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True,
    )

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for celda in worksheet[1]:
        celda.fill = relleno_encabezado
        celda.font = fuente_encabezado
        celda.alignment = alineacion_encabezado

    for fila in worksheet.iter_rows(min_row=2):
        for celda in fila:
            celda.alignment = alineacion_cuerpo

    for celdas_columna in worksheet.columns:
        letra_columna = get_column_letter(
            celdas_columna[0].column
        )

        largo_maximo = 0

        for celda in celdas_columna:
            if celda.value is not None:
                largo_maximo = max(
                    largo_maximo,
                    len(str(celda.value)),
                )

        worksheet.column_dimensions[
            letra_columna
        ].width = min(
            max(largo_maximo + 2, 14),
            70,
        )

    for fila in worksheet.iter_rows():
        worksheet.row_dimensions[fila[0].row].height = (
            35 if fila[0].row == 1 else 45
        )


def crear_excel_en_memoria(
    df_resumen,
    df_detalle,
    df_alertas,
):
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        df_resumen.to_excel(
            writer,
            sheet_name="Para copiar",
            index=False,
        )

        df_detalle.to_excel(
            writer,
            sheet_name="Detalle editable",
            index=False,
        )

        df_alertas.to_excel(
            writer,
            sheet_name="Alertas",
            index=False,
        )

        for nombre_hoja in [
            "Para copiar",
            "Detalle editable",
            "Alertas",
        ]:
            aplicar_formato_excel(
                writer,
                nombre_hoja,
            )

    output.seek(0)
    return output.getvalue()


def crear_respaldo_zip(
    texto_original,
    df_resumen,
    df_detalle,
    df_alertas,
):
    timestamp = datetime.now().strftime(
        "%Y-%m-%d_%H-%M-%S"
    )

    output = BytesIO()

    excel_bytes = crear_excel_en_memoria(
        df_resumen,
        df_detalle,
        df_alertas,
    )

    with zipfile.ZipFile(
        output,
        "w",
        zipfile.ZIP_DEFLATED,
    ) as archivo_zip:
        archivo_zip.writestr(
            f"{timestamp}_reporte_para_copiar.xlsx",
            excel_bytes,
        )

        archivo_zip.writestr(
            f"{timestamp}_texto_original.txt",
            texto_original,
        )

        archivo_zip.writestr(
            f"{timestamp}_resumen_para_copiar.csv",
            df_resumen.to_csv(index=False).encode("utf-8-sig"),
        )

        archivo_zip.writestr(
            f"{timestamp}_detalle_editable.csv",
            df_detalle.to_csv(index=False).encode("utf-8-sig"),
        )

        archivo_zip.writestr(
            f"{timestamp}_alertas.csv",
            df_alertas.to_csv(index=False).encode("utf-8-sig"),
        )

    output.seek(0)
    return output.getvalue(), timestamp


if "texto_reportes" not in st.session_state:
    st.session_state["texto_reportes"] = ""

if "resultado_generado" not in st.session_state:
    st.session_state["resultado_generado"] = False

if "df_extraido" not in st.session_state:
    st.session_state["df_extraido"] = pd.DataFrame(
        columns=COLUMNAS_EDITOR
    )

if "alertas_parser" not in st.session_state:
    st.session_state["alertas_parser"] = []


def limpiar_caja_texto():
    st.session_state["texto_reportes"] = ""


def limpiar_todo():
    st.session_state["texto_reportes"] = ""
    st.session_state["resultado_generado"] = False
    st.session_state["df_extraido"] = pd.DataFrame(
        columns=COLUMNAS_EDITOR
    )
    st.session_state["alertas_parser"] = []


def agregar_reporte_tipo():
    texto_actual = st.session_state.get(
        "texto_reportes",
        "",
    )

    if texto_actual.strip():
        st.session_state["texto_reportes"] = (
            texto_actual.strip()
            + "\n\n"
            + REPORTE_TIPO
        )
    else:
        st.session_state["texto_reportes"] = REPORTE_TIPO


def cargar_reporte_ejemplo():
    texto_actual = st.session_state.get(
        "texto_reportes",
        "",
    )

    if texto_actual.strip():
        st.session_state["texto_reportes"] = (
            texto_actual.strip()
            + "\n\n"
            + REPORTE_EJEMPLO
        )
    else:
        st.session_state["texto_reportes"] = REPORTE_EJEMPLO


st.title("📋 Generador de reportes arqueológicos desde WhatsApp")

st.write(
    "Convierte mensajes de WhatsApp en un reporte "
    "estructurado, editable y exportable a Excel."
)

with st.expander(
    "ℹ️ Instrucciones de uso",
    expanded=True,
):
    st.markdown(
        """
        **La aplicación permite:**

        - Extraer fecha, unidad, nivel, estrato/capa,
          cronología, materiales y cantidades.
        - Reconocer **Histórico**, **Prehispánico** y **Subactual**.
        - Reconocer variantes como `prehispánico`,
          `prehispanico`, `pre-hispánico`,
          `pre-hispanico` y `pre hispánico`.
        - Detectar niveles con conteos en cero.
        - Conservar y alertar materiales nuevos no catastrados.
        - Corregir, agregar o eliminar filas manualmente.
        - Descargar un Excel y un respaldo completo ZIP.

        **Flujo recomendado:**

        1. Pega uno o varios reportes.
        2. Presiona **Procesar reportes**.
        3. Revisa y corrige la tabla editable.
        4. Revisa las alertas.
        5. Descarga el resultado.
        """
    )

st.subheader("1. Ingreso de reportes")

st.text_area(
    "Pega aquí uno o varios reportes de WhatsApp",
    height=360,
    key="texto_reportes",
    placeholder=(
        "[10/7, 18:43] Persona: 10/07/26\n"
        "Unidad 18-4\n"
        "N38, capa C\n"
        "Subactual: 0\n"
        "Histórico:\n"
        "-2 lozas\n"
        "Prehispánico:\n"
        "-1 cerámica"
    ),
)

columnas_botones = st.columns([1.15, 1, 1, 1, 1])

with columnas_botones[0]:
    procesar = st.button(
        "Procesar reportes",
        type="primary",
    )

with columnas_botones[1]:
    st.button(
        "Limpiar caja",
        on_click=limpiar_caja_texto,
    )

with columnas_botones[2]:
    st.button(
        "Agregar reporte tipo",
        on_click=agregar_reporte_tipo,
    )

with columnas_botones[3]:
    st.button(
        "Cargar ejemplo",
        on_click=cargar_reporte_ejemplo,
    )

with columnas_botones[4]:
    st.button(
        "Limpiar todo",
        on_click=limpiar_todo,
    )

if procesar:
    texto = st.session_state["texto_reportes"]

    if not texto.strip():
        st.warning("Debes pegar al menos un reporte.")
    else:
        registros, alertas_parser = parsear_todo(texto)

        st.session_state["df_extraido"] = registros_a_df(
            registros
        )

        st.session_state["alertas_parser"] = alertas_parser
        st.session_state["resultado_generado"] = True

if st.session_state["resultado_generado"]:
    st.subheader("2. Revisar y editar datos extraídos")

    st.info(
        "Puedes corregir campos, agregar filas nuevas o eliminar filas. "
        "Cuando falte información, puedes completarla o dejarla como "
        "'No informado'."
    )

    df_base = st.session_state["df_extraido"].copy()

    for columna in COLUMNAS_EDITOR:
        if columna not in df_base.columns:
            df_base[columna] = None

    df_base = df_base[COLUMNAS_EDITOR]

    df_editado = st.data_editor(
        df_base,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "fecha": st.column_config.TextColumn("Fecha"),
            "unidad": st.column_config.TextColumn("Unidad"),
            "nivel": st.column_config.NumberColumn(
                "Nivel",
                step=1,
            ),
            "estrato": st.column_config.TextColumn(
                "Estrato/Capa"
            ),
            "cronologia": st.column_config.SelectboxColumn(
                "Cronología",
                options=[
                    "histórico",
                    "prehispánico",
                    "subactual",
                    "No informado",
                ],
            ),
            "material": st.column_config.TextColumn("Material"),
            "cantidad": st.column_config.NumberColumn(
                "Cantidad",
                step=1,
            ),
            "material_nuevo": st.column_config.CheckboxColumn(
                "Material nuevo"
            ),
            "linea_original": st.column_config.TextColumn(
                "Línea original"
            ),
        },
        key="editor_detalle",
    )

    df_editado_limpio = limpiar_df_editado(df_editado)

    df_resumen = resumen_por_grupo_df(
        df_editado_limpio
    )

    df_alertas = validar_df(
        df_editado_limpio,
        st.session_state.get(
            "alertas_parser",
            [],
        ),
    )

    st.subheader("3. Alertas y validaciones")

    if "Estado" in df_alertas.columns:
        st.success("Sin alertas.")
    else:
        st.warning(
            "Existen alertas o campos que requieren revisión."
        )

    with st.expander(
        "Ver alertas y validaciones",
        expanded=("Alerta" in df_alertas.columns),
    ):
        st.dataframe(
            df_alertas,
            use_container_width=True,
        )

    materiales_nuevos = sorted(
        set(
            df_editado_limpio.loc[
                df_editado_limpio["material_nuevo"] == True,
                "material",
            ].dropna()
        )
    )

    if materiales_nuevos:
        st.subheader("Materiales no reconocidos detectados")
        st.warning(
            "Estos materiales se conservarán en el reporte, "
            "pero conviene revisar su nombre."
        )
        st.write(
            ", ".join(
                titulo_material(material)
                for material in materiales_nuevos
            )
        )

    st.subheader("4. Reporte para copiar")

    if df_resumen.empty:
        st.error(
            "No se pudo generar el resumen. "
            "Revisa o completa la tabla editable."
        )
    else:
        st.dataframe(
            df_resumen,
            use_container_width=True,
        )

        excel_bytes = crear_excel_en_memoria(
            df_resumen,
            df_editado_limpio,
            df_alertas,
        )

        columnas_descarga = st.columns([1, 1])

        with columnas_descarga[0]:
            st.download_button(
                label="Descargar Excel",
                data=excel_bytes,
                file_name="reporte_para_copiar.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )

        with columnas_descarga[1]:
            respaldo_bytes, timestamp = crear_respaldo_zip(
                texto_original=st.session_state["texto_reportes"],
                df_resumen=df_resumen,
                df_detalle=df_editado_limpio,
                df_alertas=df_alertas,
            )

            st.download_button(
                label="Descargar respaldo completo ZIP",
                data=respaldo_bytes,
                file_name=f"respaldo_reporte_{timestamp}.zip",
                mime="application/zip",
            )
