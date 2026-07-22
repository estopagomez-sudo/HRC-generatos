import streamlit as st
import io, zipfile, re
from datetime import datetime
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# Leer datos_entrada.xlsx
# ─────────────────────────────────────────────────────────────────────────────
def leer_datos(f):
    wb = load_workbook(f, data_only=True)
    ws = wb["DATOS_PROYECTO"]

    def v(ref):
        val = ws[ref].value
        if val is None:
            return ""
        if isinstance(val, datetime):
            return val.strftime("%d/%m/%Y")
        s = str(val).strip()
        # Limpiar prefijos de ejemplo que el usuario haya dejado
        s = re.sub(r"^ej:\s*", "", s, flags=re.IGNORECASE)
        return s

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
# Helpers — todo string replace, nunca ET.tostring()
# ─────────────────────────────────────────────────────────────────────────────
def get_sheet_map(files):
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
    return match.group(1).replace("../", "xl/") if match else None

# ── Portada: reemplazar último <a:t> de CuadroTexto 4 ────────────────────────
def actualizar_titulo_portada(drawing_bytes, nuevo_titulo):
    xml   = drawing_bytes.decode("utf-8")
    match = re.search(r'name="CuadroTexto 4".*?</xdr:sp>', xml, re.DOTALL)
    if not match:
        return drawing_bytes, False, "shape 'CuadroTexto 4' no encontrado"

    shape  = match.group(0)
    textos = re.findall(r'<a:t>([^<]*)</a:t>', shape)
    if not textos:
        return drawing_bytes, False, "no se encontraron runs de texto"

    titulo_esc  = (nuevo_titulo
                   .replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;"))
    shape_nuevo = shape.replace(
        f"<a:t>{textos[-1]}</a:t>",
        f"<a:t>{titulo_esc}</a:t>",
        1
    )
    xml_nuevo = xml[:match.start()] + shape_nuevo + xml[match.end():]
    return xml_nuevo.encode("utf-8"), True, f"'{textos[-1]}' → '{nuevo_titulo}'"

# ── Header de página: reemplazar TODO el bloque No. Doc. + Rev. ───────────────
def actualizar_odd_header(sheet_bytes, cod_rev_general, num_rev):
    """
    Reemplaza toda la línea 'No. Doc. XXXX' y 'Rev. SXX'
    sin importar cuál sea el valor actual.
    """
    xml_str  = sheet_bytes.decode("utf-8")
    # Quitar sufijo -S00/-S01 del código si viene incluido (va en Rev. por separado)
    cod_base = re.sub(r"-S\d{2}$", "", cod_rev_general).strip()
    rev_str  = "S" + num_rev.zfill(2)

    # Reemplaza TODO lo que haya entre "No. Doc. " y el primer \r\n real
    nueva, n = re.subn(
        r"(No\.[ ]Doc\.[ ])[^\r\n]+(\r\n)(Rev\.[ ])S\d{2}",
        lambda m: f"{m.group(1)}{cod_base} {m.group(2)}{m.group(3)}{rev_str}",
        xml_str
    )
    return nueva.encode("utf-8"), n > 0


# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador HRC", page_icon="📋", layout="centered")
st.title("📋 Generador HRC — Paso 1: Portada")
st.caption("Modifica el título y el header de página en Portada y Control de Firmas.")

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
| `1. PORTADA` | Header (`No. Doc. / Rev.`) | `B5` + `B2` | ✅ Se actualiza |
| `1. PORTADA` | `CuadroTexto 4` — último run (título) | `B4` | ✅ Se actualiza |
| `2. CONTROL DE FIRMAS` | Header (`No. Doc. / Rev.`) | `B5` + `B2` | ✅ Se actualiza |
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

        # ── 1. Título en CuadroTexto 4 ────────────────────────────────────────
        portada = next((s for s in sheet_map if "PORTADA" in s.upper()), None)
        if portada:
            snum  = sheet_map[portada]
            dpath = get_drawing_path(files, snum)
            if dpath and dpath in files:
                titulo    = datos["titulo_doc"] or "Ingeniería Conceptual Casetas Técnicas"
                new_bytes, ok, msg = actualizar_titulo_portada(files[dpath], titulo)
                files[dpath] = new_bytes
                log.append(f"{'✅' if ok else '⚠️'} CuadroTexto 4 → {msg}")
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

        # ── Reconstruir ZIP ───────────────────────────────────────────────────
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w") as zout:
            for fname, content in files.items():
                ct = infos[fname].compress_type if fname in infos else zipfile.ZIP_DEFLATED
                zout.writestr(fname, content, compress_type=ct)
        out_buf.seek(0)

        st.success("✅ Portada generada")
        with st.expander("Log detallado"):
            for line in log: st.text(line)

        # Nombre del archivo basado en el código de revisión general
        nombre_salida = f"{datos['cod_rev_general'] or f'HRC_S{num_rev}'}.xlsx"
        st.download_button(
            label=f"⬇️ Descargar {nombre_salida}",
            data=out_buf.getvalue(),
            file_name=nombre_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.info("📌 Verifica: título en portada y header en Vista → Diseño de página. Cuando esté correcto continuamos.")

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)
