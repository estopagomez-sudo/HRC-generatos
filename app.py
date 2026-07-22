import streamlit as st
import io, zipfile, copy, re
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
        "cod_rev_general": v("B5"),   # Código de Revisión General  ej: MI-TTT-SSTA-02-YYY-HRC-GE_YYY-0014-S00
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
    wb_root = ET.fromstring(files["xl/workbook.xml"].decode("utf-8"))
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    return {sh.get("name"): i for i, sh in
            enumerate(wb_root.findall(f".//{{{ns}}}sheet"), start=1)}

def get_drawing_path(files, sheet_num):
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

# ── Drawing: actualizar runs de un textbox manteniendo estilos ────────────────
def actualizar_textbox(drawing_xml_str, shape_name, nuevos_runs):
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

        parrafos  = txBody.findall(f"{{{NS_A}}}p")
        primer_p  = parrafos[0] if parrafos else None
        runs_orig = primer_p.findall(f"{{{NS_A}}}r") if primer_p else []
        rprs      = [r.find(f"{{{NS_A}}}rPr") for r in runs_orig]
        endPr     = primer_p.find(f"{{{NS_A}}}endParaRPr") if primer_p else None

        for p in parrafos:
            txBody.remove(p)

        p_new = ET.SubElement(txBody, f"{{{NS_A}}}p")
        for i, texto in enumerate(nuevos_runs):
            r_new   = ET.SubElement(p_new, f"{{{NS_A}}}r")
            rpr_src = rprs[i] if i < len(rprs) else (rprs[-1] if rprs else None)
            if rpr_src is not None:
                r_new.append(copy.deepcopy(rpr_src))
            t_el = ET.SubElement(r_new, f"{{{NS_A}}}t")
            t_el.text = texto
            if texto != texto.strip():
                t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        if endPr is not None:
            p_new.append(copy.deepcopy(endPr))

        found = True
        break

    return ET.tostring(root, encoding="unicode", xml_declaration=False), found

# ── Sheet XML: actualizar oddHeader ──────────────────────────────────────────
def actualizar_header(sheet_xml_str, cod_rev_general, num_rev):
    """
    Reemplaza en oddHeader:
      - El código de documento completo (todo antes de \nRev.)
      - La revisión (S00, S01…)
    Formato original:
      &L&G&R&9No. Doc. MI-TTT-SSTA-02-YYY-HRC-GE_YYY-0014 \nRev. S00\n...
    """
    root = ET.fromstring(sheet_xml_str)
    hf   = root.find(f"{{{NS_SS}}}headerFooter")
    if hf is None:
        return ET.tostring(root, encoding="unicode", xml_declaration=False), False

    oh = hf.find(f"{{{NS_SS}}}oddHeader")
    if oh is None or oh.text is None:
        return ET.tostring(root, encoding="unicode", xml_declaration=False), False

    original = oh.text

    # Extraer la parte del código: todo lo que hay entre "No. Doc. " y "\n"
    # y reemplazar con cod_rev_general (sin el sufijo -S00 ya que va separado)
    # El código de revisión general puede venir como "MI-...-S00" o sin sufijo
    # En el header aparece sin el sufijo de revisión: "MI-TTT-SSTA-02-YYY-HRC-GE_YYY-0014"
    # Separamos: código base = todo hasta el último "-S" o directamente el campo
    cod_base = re.sub(r'-S\d+$', '', cod_rev_general)  # quitar -S00 si viene incluido
    rev_str  = f"S{num_rev.zfill(2)}"

    # Patrón: "No. Doc. XXXX \nRev. SYY"
    nuevo = re.sub(
        r'(No\. Doc\. )([^\n]+)(\nRev\. )(S\d+)',
        lambda m: f"{m.group(1)}{cod_base} {m.group(3)}{rev_str}",
        original
    )

    oh.text = nuevo
    return ET.tostring(root, encoding="unicode", xml_declaration=False), (nuevo != original)


# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador HRC", page_icon="📋", layout="centered")
st.title("📋 Generador HRC — Paso 1: Portada")
st.caption("Modifica la portada: título en el cuadro de texto y código en el header de página.")

st.divider()
col1, col2 = st.columns(2)
with col1:
    st.markdown("**1 · Excel de datos**")
    f_datos = st.file_uploader("datos_entrada.xlsx", type=["xlsx"], key="datos")
with col2:
    st.markdown("**2 · Plantilla HRC**")
    f_plantilla = st.file_uploader("plantilla_HRC.xlsx", type=["xlsx"], key="plantilla")

datos = None
if f_datos:
    try:
        datos = leer_datos(f_datos)
        st.divider()
        st.subheader("Datos leídos")
        col_a, col_b = st.columns([1, 2])
        with col_a:
            for lbl in ["Revisión", "Título doc", "Cód. revisión gral", "Clave ref",
                        "Cód. documento", "Originador", "Entidad revisora", "Fecha ciclo"]:
                st.markdown(lbl)
        with col_b:
            for key in ["num_rev","titulo_doc","cod_rev_general","clave_ref",
                        "cod_documento","originador","entidad_rev","fecha_ciclo"]:
                st.code(datos[key] or "(vacío)")

        vacios = [k for k in ["titulo_doc","cod_rev_general","num_rev"] if not datos[k]]
        if vacios:
            st.warning(f"⚠️  Campos vacíos en datos_entrada: {', '.join(vacios)}")
    except Exception as e:
        st.error(f"Error leyendo datos_entrada.xlsx: {e}")

st.divider()

with st.expander("📌 ¿Qué se va a modificar?", expanded=True):
    st.markdown("""
**Hoja `1. PORTADA (3)`**

| Elemento | Ubicación | Campo origen | Acción |
|----------|-----------|-------------|--------|
| Header de página (arriba) | `oddHeader` en sheet1.xml | `B5` Cód. revisión gral + `B2` Revisión | ✅ Se actualiza |
| Título del documento | `CuadroTexto 4` run 3 — drawing1.xml | `B4` Título | ✅ Se actualiza |
| Header descriptivo largo | `CuadroTexto 3` — drawing1.xml | — | ❌ No se toca |

**Hoja `2. CONTROL DE FIRMAS (2)`**

| Elemento | Ubicación | Campo origen | Acción |
|----------|-----------|-------------|--------|
| Header de página (arriba) | `oddHeader` en sheet2.xml | `B5` + `B2` | ✅ Se actualiza |
""")

generar = st.button(
    "⚡ Generar — portada + headers",
    disabled=(not f_datos or not f_plantilla or datos is None),
    type="primary"
)

if generar and datos and f_plantilla:
    try:
        f_plantilla.seek(0)
        with zipfile.ZipFile(io.BytesIO(f_plantilla.read()), "r") as zin:
            files = {name: zin.read(name) for name in zin.namelist()}

        sheet_map = get_sheet_map(files)
        log = []
        num_rev = datos["num_rev"].zfill(2)

        # ── 1. CuadroTexto 4 en drawing de PORTADA ───────────────────────────
        portada_sheet = next((s for s in sheet_map if "PORTADA" in s.upper()), None)
        if portada_sheet:
            snum         = sheet_map[portada_sheet]
            drawing_path = get_drawing_path(files, snum)
            if drawing_path and drawing_path in files:
                drawing_xml = files[drawing_path].decode("utf-8")
                runs_nuevos = [
                    "Supervisión de Sistemas de Señalizacion Ferroviaria, Telecomunicaciones y Control",
                    ":: Tren",
                    " México-Querétaro-Irapuato  ",
                    datos["titulo_doc"] or "Ingeniería Conceptual Casetas Técnicas",
                ]
                drawing_xml, found = actualizar_textbox(drawing_xml, "CuadroTexto 4", runs_nuevos)
                files[drawing_path] = drawing_xml.encode("utf-8")
                log.append(f"{'✅' if found else '⚠️ '} CuadroTexto 4 → '{datos['titulo_doc']}'")

        # ── 2. oddHeader en PORTADA y CONTROL DE FIRMAS ───────────────────────
        hojas_con_header = [s for s in sheet_map
                            if "PORTADA" in s.upper() or "CONTROL" in s.upper() or "FIRMA" in s.upper()]

        for sname in hojas_con_header:
            snum       = sheet_map[sname]
            sheet_path = f"xl/worksheets/sheet{snum}.xml"
            if sheet_path not in files:
                continue
            sheet_xml = files[sheet_path].decode("utf-8")
            sheet_xml, changed = actualizar_header(
                sheet_xml, datos["cod_rev_general"], num_rev)
            files[sheet_path] = sheet_xml.encode("utf-8")
            log.append(f"{'✅' if changed else '⚠️ (sin cambio)'} oddHeader '{sname}' → "
                       f"No. Doc. {re.sub(r'-S\\d+$','',datos['cod_rev_general'])}  Rev. S{num_rev}")

        # ── Reconstruir xlsx ──────────────────────────────────────────────────
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for fname, content in files.items():
                zout.writestr(fname, content)
        out_buf.seek(0)

        st.success("✅ Portada y headers generados")
        with st.expander("Log detallado"):
            for line in log:
                st.text(line)

        st.download_button(
            label="⬇️ Descargar Excel con portada actualizada",
            data=out_buf.getvalue(),
            file_name=f"HRC_S{num_rev}_portada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.info("📌 Verifica el header de página (vista Diseño de página en Excel) y el título en la portada. Cuando esté correcto, continuamos.")

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)
