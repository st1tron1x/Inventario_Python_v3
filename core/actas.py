# core/actas.py
"""
Generación del Acta de Entrega de Elementos y Equipos Tecnológicos
(formato FR-GTI-003 V1), a partir de assets/plantilla_acta.xlsx.

Sigue el mismo patrón de rutas que database.py (compatible con
ejecución normal y con el .exe empaquetado con PyInstaller).
"""

import os
import sys
import json
import zipfile
import re
from datetime import datetime

from openpyxl import load_workbook

from core.database import conectar


# =========================================================
# Rutas base (mismo criterio que database.py)
# =========================================================

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

RUTA_PLANTILLA = os.path.join(BASE_DIR, "assets", "plantilla_acta.xlsx")
CARPETA_ACTAS = os.path.join(BASE_DIR, "actas")


# =========================================================
# Datos de quien ENTREGA el equipo (fijo por ahora).
# Ajusta esto si distintas personas hacen las entregas.
# =========================================================

ENTREGADO_POR = {
    "nombre": "William de la Cruz Méndez",
    "cargo": "Auxiliar de Sistemas",
    "documento": "1111812084",  # <-- pon aquí la cédula real de quien entrega
}

# Tipo de equipo que se reporta en "5. Detalle del Equipo Asignado".
# El inventario no distingue portátil/escritorio, así que queda fijo;
# cámbialo si empiezan a manejar equipos de escritorio.
TIPO_EQUIPO = "Computador Portátil"


# =========================================================
# Mapeo de celdas de la plantilla FR-GTI-003
# =========================================================

CELDAS = {
    # 3. Datos del responsable (colaborador que recibe)
    "responsable_nombre":     "C9",
    "responsable_documento":  "C10",
    "responsable_cargo":      "C11",
    "responsable_area":       "C12",
    "responsable_sede":       "C13",
    "responsable_correo":     "C14",
    # C15 "Teléfono" no existe en la base de datos -> se deja tal cual venga en la plantilla.

    # 5. Detalle del equipo asignado
    "equipo_tipo":         "C19",
    "equipo_hostname":     "C20",
    "equipo_fabricante":   "C21",
    "equipo_modelo":       "C22",
    "equipo_serial":       "C23",
    "equipo_procesador":   "C24",
    "equipo_memoria":      "C25",
    "equipo_disco1":       "C26",
    "equipo_disco2":       "C27",
}

# 4. Declaración de responsabilidad: es un párrafo con los espacios en
# blanco incrustados en el mismo texto (no hay celdas separadas), así
# que se reconstruye completo reemplazando cada blanco.
CELDA_DECLARACION = "B17"

_PLANTILLA_DECLARACION = (
    "Yo, {nombre}, identificado(a) con CC N.° {documento}, quien se "
    "desempeña como ({cargo}) en el área de {area} declaro haber "
    "recibido en perfectas condiciones el equipo tecnológico descrito "
    "a continuación y me comprometo a hacer uso responsable del mismo, "
    "a reportar cualquier daño, pérdida o anomalía y a devolverlo en "
    "las mismas condiciones al momento de mi retiro o cuando la "
    "empresa lo solicite."
)

# 7. Firmas
CELDA_FIRMA_NOMBRE_RESPONSABLE = "B43"
CELDA_FIRMA_CARGO_RESPONSABLE = "B44"
CELDA_FIRMA_CC_RESPONSABLE = "B45"
CELDA_FIRMA_FECHA_RESPONSABLE = "B46"

CELDA_FIRMA_NOMBRE_ENTREGA = "E43"
CELDA_FIRMA_CARGO_ENTREGA = "E44"
CELDA_FIRMA_CC_ENTREGA = "E45"
CELDA_FIRMA_FECHA_ENTREGA = "E46"

# NOTA: B45 ("CC: XXXXXXXX") y E45 ("Cargo: XXXXXXXXXX") están duplicadas
# respecto a B44/E44 en la plantilla actual — parece un error de copiado
# al armar el FR-GTI-003. Se dejan intactas a propósito hasta confirmar
# con quien mantiene el formato qué dato va realmente ahí.


# =========================================================
# Utilidades de formato
# =========================================================

def _formatear_memoria(memoria_raw):
    """Combina todos los módulos de RAM en un solo texto (campo único en la plantilla)."""

    try:
        memorias = json.loads(memoria_raw) if memoria_raw else []
    except (TypeError, json.JSONDecodeError):
        return memoria_raw or "No detectado"

    if not memorias:
        return "No detectado"

    partes = [
        f"{m.get('capacidad_gb', '?')} GB {m.get('tipo', 'Desconocido')} "
        f"{m.get('velocidad_mhz', '?')}MHz ({m.get('fabricante', 'Unknown')})"
        for m in memorias
    ]

    return " + ".join(partes)


def _formatear_disco(disco):
    return (
        f"{disco.get('tipo', 'Desconocido')} "
        f"{disco.get('capacidad_gb', '?')} GB "
        f"({disco.get('modelo', 'Unknown')})"
    )


def _lista_discos(discos_raw):
    """Devuelve hasta 2 discos formateados como texto: [disco1, disco2]."""

    try:
        discos = json.loads(discos_raw) if discos_raw else []
    except (TypeError, json.JSONDecodeError):
        discos = []

    textos = [_formatear_disco(d) for d in discos]

    disco1 = textos[0] if len(textos) > 0 else "No detectado"
    disco2 = textos[1] if len(textos) > 1 else ""

    return disco1, disco2


def _formatear_sistema(sistema_raw):

    try:
        datos = json.loads(sistema_raw.replace("'", '"')) if sistema_raw else {}
        nombre = datos.get("nombre", "")
        version = datos.get("version", "")
        return f"{nombre} ({version})".strip()
    except Exception:
        return sistema_raw or "No detectado"


def _nombre_archivo(nombre_colaborador):

    limpio = "".join(
        c for c in nombre_colaborador.upper().replace(" ", "_")
        if c.isalnum() or c == "_"
    )

    #fecha = datetime.now().strftime("%Y%m%d_%H%M")

    return f"ACTA_ENTREGA_{limpio}_.xlsx"


def _parsear_fecha_registro(fecha_registro):
    """Convierte el texto de equipos.fecha_registro a datetime."""

    if isinstance(fecha_registro, datetime):
        return fecha_registro

    for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(fecha_registro, formato)
        except (TypeError, ValueError):
            continue

    return datetime.now()


# =========================================================
# Acceso a datos
# =========================================================

def obtener_entregas(colaborador_id=None):
    """
    Devuelve la(s) fila(s) unidas de colaboradores + su último equipo
    asignado. Si colaborador_id es None, trae todos los colaboradores
    que tengan al menos un equipo registrado.
    """

    conexion = conectar()
    cursor = conexion.cursor()

    consulta = """
        SELECT
            c.id,
            c.tipo_documento,
            c.documento,
            c.nombre,
            c.cargo,
            c.area,
            c.sede,
            c.correo,

            e.hostname,
            e.usuario_windows,
            e.serial,
            e.fabricante,
            e.modelo,
            e.procesador,
            e.memoria,
            e.discos,
            e.ip,
            e.mac,
            e.sistema,
            e.fecha_registro

        FROM colaboradores c
        INNER JOIN equipos e ON e.colaborador_id = c.id
    """

    parametros = ()

    if colaborador_id is not None:
        consulta += " WHERE c.id = ?"
        parametros = (colaborador_id,)

    consulta += " ORDER BY c.id, e.fecha_registro DESC"

    cursor.execute(consulta, parametros)
    filas = cursor.fetchall()

    conexion.close()

    vistos = set()
    resultado = []

    for fila in filas:
        colaborador_id_fila = fila[0]

        if colaborador_id_fila in vistos:
            continue

        vistos.add(colaborador_id_fila)
        resultado.append(fila)

    return resultado


# =========================================================
# Generación del acta
# =========================================================

# =========================================================
# Checkboxes nativos de Excel (featurePropertyBag)
# =========================================================
# openpyxl no soporta la parte "xl/featurePropertyBag/featurePropertyBag.xml"
# que usa Excel (2024+) para renderizar celdas booleanas como casillas de
# verificación reales, y la elimina al guardar. Sin ella, esas celdas se
# ven como texto "VERDADERO"/"FALSO" en vez de un cuadro marcable.
# Por eso, después de guardar con openpyxl, se reinyecta esa parte
# (y sus referencias) copiándola tal cual de la plantilla original.

def _leer_feature_property_bag():
    """Lee el featurePropertyBag.xml de la plantilla, si existe."""

    try:
        with zipfile.ZipFile(RUTA_PLANTILLA) as z:
            return z.read("xl/featurePropertyBag/featurePropertyBag.xml")
    except (FileNotFoundError, KeyError):
        return None


def _restaurar_checkboxes(ruta_xlsx):
    """
    Reinyecta xl/featurePropertyBag/featurePropertyBag.xml (y sus
    referencias en [Content_Types].xml y workbook.xml.rels) en el
    archivo ya guardado por openpyxl, para que las casillas de
    verificación se sigan viendo como checkbox real y no como texto.
    No hace nada si la plantilla no tenía esa parte.
    """

    contenido_bag = _leer_feature_property_bag()

    if contenido_bag is None:
        return

    with zipfile.ZipFile(ruta_xlsx, "r") as zin:
        datos = {n: zin.read(n) for n in zin.namelist()}

    # 1) Declarar el content-type de la parte
    ct = datos["[Content_Types].xml"].decode("utf-8")

    if "featurePropertyBag" not in ct:
        override = (
            '<Override PartName="/xl/featurePropertyBag/featurePropertyBag.xml" '
            'ContentType="application/vnd.ms-excel.featurepropertybag+xml"/>'
        )
        ct = ct.replace("</Types>", override + "</Types>")
        datos["[Content_Types].xml"] = ct.encode("utf-8")

    # 2) Relacionar la parte desde el workbook
    ruta_rels = "xl/_rels/workbook.xml.rels"
    rels = datos[ruta_rels].decode("utf-8")

    if "FeaturePropertyBag" not in rels:
        ids_usados = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels)]
        nuevo_id = max(ids_usados) + 1 if ids_usados else 1

        relacion = (
            f'<Relationship Id="rId{nuevo_id}" '
            'Type="http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag" '
            'Target="featurePropertyBag/featurePropertyBag.xml"/>'
        )
        rels = rels.replace("</Relationships>", relacion + "</Relationships>")
        datos[ruta_rels] = rels.encode("utf-8")

    # 3) Agregar la parte en sí
    datos["xl/featurePropertyBag/featurePropertyBag.xml"] = contenido_bag

    ruta_tmp = ruta_xlsx + ".tmp"

    with zipfile.ZipFile(ruta_tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for nombre, contenido in datos.items():
            zout.writestr(nombre, contenido)

    os.replace(ruta_tmp, ruta_xlsx)


def _rellenar_plantilla(fila, ruta_salida):

    (
        _id, tipo_documento, documento, nombre, cargo, area, sede, correo,
        hostname, usuario_windows, serial, fabricante, modelo, procesador,
        memoria, discos, ip, mac, sistema, fecha_registro
    ) = fila

    if not os.path.exists(RUTA_PLANTILLA):
        raise FileNotFoundError(
            f"No se encontró la plantilla en: {RUTA_PLANTILLA}\n"
            "Copia el FR-GTI-003 a assets/plantilla_acta.xlsx"
        )

    wb = load_workbook(RUTA_PLANTILLA)
    ws = wb.active

    fecha_entrega = _parsear_fecha_registro(fecha_registro)
    fecha_texto = fecha_entrega.strftime("%d/%m/%Y")

    disco1, disco2 = _lista_discos(discos)

    valores = {
        "responsable_nombre": nombre,
        "responsable_documento": f"{tipo_documento} {documento}",
        "responsable_cargo": cargo,
        "responsable_area": area,
        "responsable_sede": sede,
        "responsable_correo": correo,

        "equipo_tipo": TIPO_EQUIPO,
        "equipo_hostname": hostname,
        "equipo_fabricante": fabricante,
        "equipo_modelo": modelo,
        "equipo_serial": serial,
        "equipo_procesador": procesador,
        "equipo_memoria": _formatear_memoria(memoria),
        "equipo_disco1": disco1,
        "equipo_disco2": disco2,
        "equipo_sistema": _formatear_sistema(sistema),
    }

    for campo, celda in CELDAS.items():
        ws[celda] = valores.get(campo, "")

    # 1. Fecha de entrega (B6 es solo la etiqueta; la respuesta va en B7).
    ws["B7"] = fecha_texto

    # 2. Tipo de novedad y accesorios: no se tocan. Se dejan con su
    # valor original (False) — con el featurePropertyBag restaurado
    # más abajo, Excel las sigue mostrando como checkbox real, no
    # como texto "FALSO". Se marcan a mano con clic en Excel.

    # 4. Declaración de responsabilidad
    ws[CELDA_DECLARACION] = _PLANTILLA_DECLARACION.format(
        nombre=nombre,
        documento=documento,
        cargo=cargo,
        area=area,
    )

    # 7. Firmas
    ws[CELDA_FIRMA_NOMBRE_RESPONSABLE] = f"Nombre: {nombre}"
    ws[CELDA_FIRMA_CC_RESPONSABLE] = f"CC: {documento}"
    ws[CELDA_FIRMA_CARGO_RESPONSABLE] = f"Cargo: {cargo}"
    ws[CELDA_FIRMA_FECHA_RESPONSABLE] = f"Fecha: {fecha_texto}"

    ws[CELDA_FIRMA_NOMBRE_ENTREGA] = f"Nombre: {ENTREGADO_POR['nombre']}"
    ws[CELDA_FIRMA_CARGO_ENTREGA] = f"Cargo: {ENTREGADO_POR['cargo']}"
    ws[CELDA_FIRMA_CC_ENTREGA] = f"CC: {ENTREGADO_POR['documento']}"
    ws[CELDA_FIRMA_FECHA_ENTREGA] = f"Fecha: {fecha_texto}"

    os.makedirs(CARPETA_ACTAS, exist_ok=True)
    wb.save(ruta_salida)

    _restaurar_checkboxes(ruta_salida)


def generar_acta(colaborador_id):
    """Genera el acta de entrega de UN colaborador. Devuelve la ruta del archivo."""

    filas = obtener_entregas(colaborador_id)

    if not filas:
        raise ValueError(
            "El colaborador no tiene ningún equipo registrado todavía."
        )

    fila = filas[0]
    nombre = fila[3]

    ruta_salida = os.path.join(CARPETA_ACTAS, _nombre_archivo(nombre))

    _rellenar_plantilla(fila, ruta_salida)

    return ruta_salida


def generar_todas_las_actas():
    """Genera un acta por cada colaborador con equipo asignado. Devuelve lista de rutas."""

    filas = obtener_entregas()

    if not filas:
        raise ValueError("No hay equipos registrados en la base de datos.")

    rutas = []

    for fila in filas:
        nombre = fila[3]
        ruta_salida = os.path.join(CARPETA_ACTAS, _nombre_archivo(nombre))
        _rellenar_plantilla(fila, ruta_salida)
        rutas.append(ruta_salida)

    return rutas