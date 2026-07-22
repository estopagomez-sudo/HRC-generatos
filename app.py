import streamlit as st
import io, zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# Namespaces
# ─────────────────────────────────────────────────────────────────────────────
NS_SS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

ET.register_namespace("xdr", NS_XDR)
ET.register_namespace("a",   NS_A)
ET.register_namespace("",    NS_SS)

# ─────────────────────────────────────────────────────────────────────────────
# Leer datos_entrada.xlsx
# ─────────────────────────────────────────────────────────────────────────────
def leer_datos(f):
    wb = load_workbook(f, data_only=True)
    ws = wb["DATOS_PROYECTO"]

    def v(ref):
        val = ws[ref].value
        return str(val).strip() if val not in (None, "") else ""

    return {
        "num_rev":         v("B2"),   # 00, 01, 02…
        "titulo_doc":      v("B4"),   # Título del documento
        "cod_rev_general": v("B5"),   # Código de Revisión General
        "clave_ref":       v("B6"),   # Clave de referencia
        "cod_documento":   v("B7"),   # Código del documento
        "originador":      v("B8"),   # Originador
        "entidad_rev":     v("B9"),   # Entidad Revisora
        "fecha_ciclo":     v("B12"),  # Fecha revisor respondió
    }

# ─────────────────────────────────────────────────────────────────────────────
# Helpers XML
# ─────────────────────────────────────────────────────────────────────────────
def get_sheet_map(files):
    """Devuelve dict {nombre_hoja: número} leyendo workbook.xml."""
    wb_xml  = files["xl/workbook.xml"].decode("utf-8")
    wb_root = ET.fromstring(wb_xml)
    ns      = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    result  = {}
    for i, sh in enumerate(wb_root.findall(f".//{{{ns}}}sheet"), start=1):
        result[sh.get("name")] = i
    return result

def get_drawing_path(files, sheet_num):
    """Devuelve la ruta interna del drawing asociado a una hoja."""
    rels_path = f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"
    if rels_path not in files:
        return None
    rels_root = ET.fromstring(files[rels_path].decode("utf-8"))
    ns_r  = "http://schemas.openxmlformats.org/package/2006/relationships"
    tipo  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    for rel in rels_root.findall(f"{{{ns_r}}}Relationship"):
        if rel.get("Type") == tipo:
            return rel.get("Target").replace("../", "xl/")
    return None

def actualizar_textbox(drawing_xml_str, shape_name, nuevos_runs):
    """
    Reemplaza los runs de texto de un shape manteniendo su estilo (rPr).

    nuevos_runs: lista de strings, uno por run.
                 Si hay más runs nuevos que originales, se repite el estilo
                 del último run original.
                 Si se pasa un solo string se pone en el primer run.
    """
    root  = ET.fromstring(drawing_xml_str)
    found = False

    for anchor in root.findall(f"{{{NS_XDR}}}twoCellAnchor"):
        sp = anchor.find(f"{{{NS_XDR}}}sp")
        if sp is None:
            continue
        cNvPr = sp.find(f".//{{{NS_XDR}}}cNvPr")
        if cNvPr is None or cNvPr.get("name") != shape_name:
            continue

        txBody = sp.find(f"{{{NS_XDR}}}txBody")
        if txBody is None:
            continue

        # Tomar el primer párrafo existente y sus runs para copiar estilos
        parrafos = txBody.findall(f"{{{NS_A}}}p")
        if not parrafos:
            continue

        primer_p = parrafos[0]
        runs_orig = primer_p.findall(f"{{{NS_A}}}r")

        # Guardar rPr de cada run original (para preservar fuente/tamaño/color)
        rprs = []
        for r in runs_orig:
            rPr = r.find(f"{{{NS_A}}}rPr")
            rprs.append(rPr)

        # Borrar todos los párrafos existentes
        for p in parrafos:
            txBody.remove(p)

        # Crear nuevo párrafo único con los runs nuevos
        p_new = ET.SubElement(txBody, f"{{{NS_A}}}p")

        for i, texto in enumerate(nuevos_runs):
            r_new = ET.SubElement(p_new, f"{{{NS_A}}}r")

            # Copiar rPr del run original correspondiente (o el último disponible)
            rpr_orig = rprs[i] if i < len(rprs) else (rprs[-1] if rprs else None)
            if rpr_orig is not None:
                import copy
                r_new.append(copy.deepcopy(rpr_orig))

            t_el = ET.SubElement(r_new, f"{{{NS_A}}}t")
            t_el.text = texto
            # Preservar espacios si los hay
            if texto != texto.strip():
                t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        # Añadir endParaRPr del párrafo original si existía
        endPr = primer_p.find(f"{{{NS_A}}}endParaRPr") if parrafos else None
        if endPr is not None:
            import copy
            p_new.append(copy.deepcopy(endPr))

        found = True
        break

    return ET.tostring(root, encoding="unicode", xml_declaration=False), found


# ─────────────────────────────────────────────────────────────────────────────
# UI Streamlit
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador HRC", page_icon="📋", layout="centered")
st.title("📋 Generador HRC — Paso 1: Portada")
st.caption("Modifica únicamente la portada. Verificamos antes de continuar.")

st.divider()
col1, col2 = st.columns(2)
with col1:
    st.markdown("**1 · Excel de datos**")
    f_datos = st.file_uploader("datos_entrada.xlsx", type=["xlsx"], key="datos")
with col2:
    st.markdown("**2 · Plantilla HRC**")
    f_plantilla = st.file_uploader("plantilla_HRC.xlsx", type=["xlsx"], key="plantilla")

# Vista previa de datos leídos
datos = None
if f_datos:
    try:
        datos = leer_datos(f_datos)
        st.divider()
        st.subheader("Datos leídos")

        col_a, col_b = st.columns([1, 2])
        with col_a:
            st.markdown("Revisión")
            st.markdown("Título del documento")
            st.markdown("Código revisión general")
            st.markdown("Clave de referencia")
            st.markdown("Código documento")
            st.markdown("Originador")
            st.markdown("Entidad revisora")
            st.markdown("Fecha ciclo")
        with col_b:
            st.code(datos["num_rev"]         or "(vacío)")
            st.code(datos["titulo_doc"]      or "(vacío)")
            st.code(datos["cod_rev_general"] or "(vacío)")
            st.code(datos["clave_ref"]       or "(vacío)")
            st.code(datos["cod_documento"]   or "(vacío)")
            st.code(datos["originador"]      or "(vacío)")
            st.code(datos["entidad_rev"]     or "(vacío)")
            st.code(datos["fecha_ciclo"]     or "(vacío)")

        if not datos["titulo_doc"]:
            st.warning("⚠️  Celda B4 vacía — el título del documento no se escribirá en la portada.")
    except Exception as e:
        st.error(f"Error leyendo datos_entrada.xlsx: {e}")

st.divider()

# ── Lo que se va a modificar ──────────────────────────────────────────────────
with st.expander("📌 ¿Qué celdas/shapes se van a modificar?", expanded=True):
    st.markdown("""
**Hoja: `1. PORTADA (3)`** — drawing1.xml

| Shape | Contenido | Acción |
|-------|-----------|--------|
| `CuadroTexto 3` | Header superior (descripción del proyecto) | **No se toca** — texto fijo del proyecto |
| `CuadroTexto 4` | Título del documento (parte inferior) | ✅ Se actualiza con `B4` del Excel de datos |

> El run 3 (último) de `CuadroTexto 4` es donde vive el título.  
> Los 3 runs anteriores (proyecto, tren, tramo) **no se tocan**.
""")

# ── Botón generar ─────────────────────────────────────────────────────────────
generar = st.button(
    "⚡ Generar — solo portada",
    disabled=(not f_datos or not f_plantilla or datos is None),
    type="primary"
)

if generar and datos and f_plantilla:
    try:
        f_plantilla.seek(0)
        plantilla_bytes = f_plantilla.read()

        # Cargar todos los archivos del xlsx en memoria
        with zipfile.ZipFile(io.BytesIO(plantilla_bytes), "r") as zin:
            files = {name: zin.read(name) for name in zin.namelist()}

        sheet_map = get_sheet_map(files)
        log = []

        # ── Buscar hoja PORTADA ────────────────────────────────────────────────
        portada_sheet = next((s for s in sheet_map if "PORTADA" in s.upper()), None)
        if not portada_sheet:
            st.error("No se encontró ninguna hoja con 'PORTADA' en el nombre.")
            st.stop()

        snum         = sheet_map[portada_sheet]
        drawing_path = get_drawing_path(files, snum)
        log.append(f"Hoja portada: '{portada_sheet}' → sheet{snum}.xml → {drawing_path}")

        if not drawing_path or drawing_path not in files:
            st.error(f"Drawing no encontrado: {drawing_path}")
            st.stop()

        drawing_xml = files[drawing_path].decode("utf-8")

        # ── CuadroTexto 4: actualizar solo el run del título (run 3) ───────────
        # Estructura original:
        #   run0: 'Supervisión de Sistemas de Señalizacion Ferroviaria…'
        #   run1: ':: Tren'
        #   run2: ' México-Querétaro-Irapuato  '
        #   run3: 'Ingeniería Conceptual Casetas Técnicas'  ← este cambia
        #
        # Pasamos los 4 runs: los 3 primeros fijos, el último con el nuevo título.
        runs_nuevos = [
            "Supervisión de Sistemas de Señalizacion Ferroviaria, Telecomunicaciones y Control",
            ":: Tren",
            " México-Querétaro-Irapuato  ",
            datos["titulo_doc"] or "Ingeniería Conceptual Casetas Técnicas",
        ]

        drawing_xml, found = actualizar_textbox(drawing_xml, "CuadroTexto 4", runs_nuevos)

        if found:
            log.append(f"✅ CuadroTexto 4 → run 3 actualizado con: '{datos['titulo_doc']}'")
        else:
            log.append("⚠️  CuadroTexto 4 no encontrado en el drawing")

        files[drawing_path] = drawing_xml.encode("utf-8")

        # ── Reconstruir xlsx ────────────────────────────────────────────────────
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for fname, content in files.items():
                zout.writestr(fname, content)
        out_buf.seek(0)

        # ── Resultado ──────────────────────────────────────────────────────────
        st.success("✅ Portada generada correctamente")
        with st.expander("Log detallado"):
            for line in log:
                st.text(line)

        num_rev = datos["num_rev"].zfill(2)
        st.download_button(
            label="⬇️ Descargar Excel con portada actualizada",
            data=out_buf.getvalue(),
            file_name=f"HRC_S{num_rev}_portada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.info("📌 Abre el Excel descargado y verifica la portada. Cuando esté correcto, continuamos con el Control de Firmas.")

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)
