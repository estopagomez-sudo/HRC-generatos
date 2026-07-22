import streamlit as st
import io, zipfile, re
from openpyxl import load_workbook

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
        "num_rev":         v("B2"),
        "titulo_doc":      v("B4"),
        "cod_rev_general": v("B5"),
        "clave_ref":       v("B6"),
        "cod_documento":   v("B7"),
        "originador":      v("B8"),
        "entidad_rev":     v("B9"),
        "fecha_ciclo":     v("B12"),
    }

# ─────────────────────────────────────────────────────────────────────────────
# ESTRATEGIA: todas las modificaciones son STRING REPLACE sobre XML crudo.
# NO usamos ET para serializar — ET renombra namespaces (xdr→ns0, a→ns1)
# y Excel no reconoce el archivo resultante.
# ─────────────────────────────────────────────────────────────────────────────

def get_sheet_map(files):
    """Lee workbook.xml con regex para no depender de ET."""
    wb_xml = files["xl/workbook.xml"].decode("utf-8")
    sheets = re.findall(r'<sheet\s+name="([^"]+)"[^/]*/>', wb_xml)
    return {name: i+1 for i, name in enumerate(sheets)}

def get_drawing_path(files, sheet_num):
    rpath = f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"
    if rpath not in files:
        return None
    rels  = files[rpath].decode("utf-8")
    match = re.search(
        r'Type="[^"]*relationships/drawing"[^>]*Target="([^"]+)"', rels)
    if match:
        return match.group(1).replace("../", "xl/")
    return None

# ── PORTADA: modificar el último <a:t> dentro de CuadroTexto 4 ───────────────
def actualizar_titulo_portada(drawing_bytes, nuevo_titulo):
    """
    Reemplaza el último run de texto dentro de CuadroTexto 4.
    Opera sobre el XML crudo sin parsear con ET.
    """
    xml = drawing_bytes.decode("utf-8")

    # Extraer el bloque del shape CuadroTexto 4
    match = re.search(r'name="CuadroTexto 4".*?</xdr:sp>', xml, re.DOTALL)
    if not match:
        return drawing_bytes, False

    shape_original = match.group(0)

    # Encontrar todos los textos <a:t>...</a:t> en el shape
    textos = re.findall(r'<a:t>([^<]*)</a:t>', shape_original)
    if not textos:
        return drawing_bytes, False

    # El último texto es el título del documento
    titulo_actual = textos[-1]

    # Escapar el título nuevo para XML
    titulo_escapado = (nuevo_titulo
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;"))

    # Reemplazar solo dentro del shape (última ocurrencia del último texto)
    shape_nuevo = shape_original.replace(
        f"<a:t>{titulo_actual}</a:t>",
        f"<a:t>{titulo_escapado}</a:t>",
        1
    )

    xml_nuevo = xml[:match.start()] + shape_nuevo + xml[match.end():]
    return xml_nuevo.encode("utf-8"), True

# ── HEADER DE PÁGINA: CR+LF reales en el oddHeader ───────────────────────────
def actualizar_odd_header(sheet_bytes, cod_rev_general, num_rev):
    """
    Formato del oddHeader:
      No. Doc. MI-TTT-SSTA-02-YYY-HRC-GE_YYY-0014 \r\nRev. S00
    donde \r\n son CR+LF reales.
    """
    xml_str  = sheet_bytes.decode("utf-8")
    cod_base = re.sub(r"-S\d{2}$", "", cod_rev_general).strip()
    rev_str  = "S" + num_rev.zfill(2)

    patron = re.compile(
        r"(No\.[ ]Doc\.[ ])([^ \r\n]+)([ ]+)(\r\n)(Rev\.[ ])(S\d{2})"
    )
    nueva, n = patron.subn(
        lambda m: (m.group(1) + cod_base + m.group(3) +
                   m.group(4) + m.group(5) + rev_str),
        xml_str
    )
    return nueva.encode("utf-8"), n > 0


# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador HRC", page_icon="📋", layout="centered")
st.title("📋 Generador HRC — Paso 1: Portada")
st.caption("Modifica el título (CuadroTexto 4) y el header de página en Portada y Control de Firmas.")

st.divider()
c1, c2 = st.columns(2)
with c1:
    st.markdown("**1 · Excel de datos**")
    f_datos = st.file_uploader("datos_entrada.xlsx", type=["xlsx"], key="datos")
with c2:
    st.markdown("**2 · Plantilla HRC**")
    f_plantilla = st.file_uploader("plantilla_HRC.xlsx", type=["xlsx"], key="plantilla")

datos = None
if f_datos:
    try:
        datos = leer_datos(f_datos)
        st.divider()
        st.subheader("Datos leídos")
        ca, cb = st.columns([1, 2])
        labels = ["Revisión","Título doc","Cód. revisión gral","Clave ref",
                  "Cód. documento","Originador","Entidad revisora","Fecha ciclo"]
        keys   = ["num_rev","titulo_doc","cod_rev_general","clave_ref",
                  "cod_documento","originador","entidad_rev","fecha_ciclo"]
        with ca:
            for l in labels: st.markdown(l)
        with cb:
            for k in keys:   st.code(datos[k] or "(vacío)")
        vacios = [k for k in ["num_rev","titulo_doc","cod_rev_general"] if not datos[k]]
        if vacios:
            st.warning(f"⚠️  Campos vacíos: {', '.join(vacios)}")
    except Exception as e:
        st.error(f"Error leyendo datos_entrada.xlsx: {e}")

st.divider()
with st.expander("📌 Qué se modifica en este paso", expanded=True):
    st.markdown("""
| Hoja | Elemento | Campo origen | Acción |
|------|----------|-------------|--------|
| `1. PORTADA` | Header de página (`No. Doc. / Rev.`) | `B5` + `B2` | ✅ Se actualiza |
| `1. PORTADA` | `CuadroTexto 4` — último run (título) | `B4` | ✅ Se actualiza |
| `2. CONTROL DE FIRMAS` | Header de página (`No. Doc. / Rev.`) | `B5` + `B2` | ✅ Se actualiza |
| `1. PORTADA` | `CuadroTexto 3` (descripción larga) | — | ❌ No se toca |
""")

generar = st.button("⚡ Generar — portada + headers",
                    disabled=(not f_datos or not f_plantilla or datos is None),
                    type="primary")

if generar and datos and f_plantilla:
    try:
        f_plantilla.seek(0)
        raw = f_plantilla.read()

        with zipfile.ZipFile(io.BytesIO(raw), "r") as zin:
            infos = {item.filename: item for item in zin.infolist()}
            files = {name: zin.read(name) for name in zin.namelist()}

        sheet_map = get_sheet_map(files)
        num_rev   = datos["num_rev"].zfill(2)
        log       = []

        # ── 1. Título en CuadroTexto 4 (drawing de la portada) ───────────────
        portada = next((s for s in sheet_map if "PORTADA" in s.upper()), None)
        if portada:
            snum  = sheet_map[portada]
            dpath = get_drawing_path(files, snum)
            if dpath and dpath in files:
                titulo = datos["titulo_doc"] or "Ingeniería Conceptual Casetas Técnicas"
                new_bytes, found = actualizar_titulo_portada(files[dpath], titulo)
                files[dpath] = new_bytes
                log.append(f"{'✅' if found else '⚠️ shape no encontrado'} "
                           f"CuadroTexto 4 → '{titulo}'")
            else:
                log.append("⚠️  Drawing de portada no encontrado")
        else:
            log.append("⚠️  Hoja PORTADA no encontrada")

        # ── 2. oddHeader en PORTADA y CONTROL DE FIRMAS ───────────────────────
        hojas_hdr = [s for s in sheet_map
                     if "PORTADA" in s.upper() or "CONTROL" in s.upper() or "FIRMA" in s.upper()]
        for sname in hojas_hdr:
            snum  = sheet_map[sname]
            spath = f"xl/worksheets/sheet{snum}.xml"
            if spath not in files:
                continue
            new_bytes, changed = actualizar_odd_header(
                files[spath], datos["cod_rev_general"], num_rev)
            files[spath] = new_bytes
            cod_base = re.sub(r"-S\d{2}$", "", datos["cod_rev_general"]).strip()
            log.append(f"{'✅' if changed else '⚠️ sin cambio'} "
                       f"Header '{sname}' → No. Doc. {cod_base}  Rev. S{num_rev}")

        # ── Reconstruir ZIP preservando tipos de compresión ───────────────────
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w") as zout:
            for fname, content in files.items():
                ct = infos[fname].compress_type if fname in infos else zipfile.ZIP_DEFLATED
                zout.writestr(fname, content, compress_type=ct)
        out_buf.seek(0)

        st.success("✅ Portada generada")
        with st.expander("Log detallado"):
            for line in log: st.text(line)

        st.download_button(
            label="⬇️ Descargar Excel con portada actualizada",
            data=out_buf.getvalue(),
            file_name=f"HRC_S{num_rev}_portada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.info("📌 Verifica: título en portada y header en Vista → Diseño de página. Cuando esté correcto continuamos.")

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)
