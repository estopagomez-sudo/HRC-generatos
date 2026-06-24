import streamlit as st
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import io
import copy

st.set_page_config(page_title="Generador HRC", page_icon="📋", layout="centered")

st.title("📋 Generador de HRC")
st.caption("Rellena automáticamente las cabeceras de todas las hojas HRC a partir del Excel de datos.")

# ─────────────────────────────────────────────────────────────────────────────
# SUBIDA DE ARCHIVOS
# ─────────────────────────────────────────────────────────────────────────────
st.divider()
col1, col2 = st.columns(2)
with col1:
    st.markdown("**1 · Excel de datos**")
    f_datos = st.file_uploader("datos_entrada.xlsx", type=["xlsx"], key="datos")
with col2:
    st.markdown("**2 · Plantilla HRC**")
    f_plantilla = st.file_uploader("plantilla_HRC.xlsx", type=["xlsx"], key="plantilla")

# ─────────────────────────────────────────────────────────────────────────────
# LECTURA DEL EXCEL DE DATOS (preview)
# ─────────────────────────────────────────────────────────────────────────────
def leer_datos(f):
    wb = load_workbook(f, data_only=True)
    ws = wb["DATOS_PROYECTO"]

    def v(ref):
        val = ws[ref].value
        return str(val).strip() if val else ""

    def fmt_fecha(val):
        if isinstance(val, datetime):
            return val.strftime("%d/%m/%Y")
        return str(val).strip() if val else ""

    num_rev = v("B2").zfill(2)  # "00", "01", "02"...

    return {
        "num_rev":          num_rev,
        "titulo_doc":       v("B6"),
        "nombre_realizado": v("B9"),
        "nombre_revisado":  v("B10"),
        "nombre_aprobado":  v("B11"),
        "fecha_firmas":     v("B12"),
        "autor_cambio":     v("B15"),
        "seccion_afectada": v("B16"),
        "desc_cambio":      v("B17"),
        "cod_rev_general":  v("B21"),
        "clave_ref":        v("B22"),
        "titulo_cab":       v("B23"),
        "cod_documento":    v("B24"),
        "originador":       v("B25"),
        "revisor_ppal":     v("B26"),
        "entidad_rev":      v("B27"),
        "fecha_ciclo":      fmt_fecha(ws["B30"].value),
    }

if f_datos:
    try:
        datos = leer_datos(f_datos)
        num_rev = datos["num_rev"]
        es_primera = (num_rev == "00")

        st.divider()
        st.subheader("Vista previa de datos")

        st.markdown(f"**Revisión:** `S{num_rev}`  →  "
                    + ("✅ Primera revisión — se llenará portada, firmas y HRC"
                       if es_primera
                       else "🔄 Revisión posterior — solo cabeceras HRC y registro de cambios"))

        if es_primera:
            with st.expander("Portada y Control de Firmas", expanded=False):
                st.write(f"Título: `{datos['titulo_doc']}`")
                st.write(f"Realizado por: `{datos['nombre_realizado']}`")
                st.write(f"Revisado por: `{datos['nombre_revisado']}`")
                st.write(f"Aprobado por: `{datos['nombre_aprobado']}`")
                st.write(f"Fecha: `{datos['fecha_firmas']}`")

        with st.expander("Registro de cambios", expanded=False):
            st.write(f"Autor: `{datos['autor_cambio']}`")
            st.write(f"Sección afectada: `{datos['seccion_afectada']}`")
            st.write(f"Cambio: `{datos['desc_cambio']}`")

        with st.expander("Cabecera HRC", expanded=False):
            st.write(f"Código revisión general: `{datos['cod_rev_general']}`")
            st.write(f"Clave referencia: `{datos['clave_ref']}`")
            st.write(f"Título: `{datos['titulo_cab']}`")
            st.write(f"Código documento: `{datos['cod_documento']}`")
            st.write(f"Originador: `{datos['originador']}`")
            st.write(f"Revisor principal: `{datos['revisor_ppal']}`")
            st.write(f"Entidad revisora: `{datos['entidad_rev']}`")
            st.write(f"Fecha del ciclo: `{datos['fecha_ciclo']}`")

    except Exception as e:
        st.error(f"Error leyendo datos_entrada.xlsx: {e}")
        datos = None
else:
    datos = None

# ─────────────────────────────────────────────────────────────────────────────
# BOTÓN GENERAR
# ─────────────────────────────────────────────────────────────────────────────
st.divider()
generar = st.button("⚡ Generar Excel llenado", disabled=(not f_datos or not f_plantilla), type="primary")

if generar and f_datos and f_plantilla:
    try:
        f_datos.seek(0)
        datos = leer_datos(f_datos)
        num_rev = datos["num_rev"]
        es_primera = (num_rev == "00")

        # Ciclo: 00→col N (1er), 01→col O (2do), 02→col P (3er)
        # Para revisiones 03+ se queda en P (3er ciclo)
        ciclo_col = {"00": "N", "01": "O", "02": "P"}.get(num_rev, "P")

        f_plantilla.seek(0)
        wb = load_workbook(f_plantilla)
        sheet_names = wb.sheetnames

        log = []

        # ── PORTADA (solo rev 00) ─────────────────────────────────────────────
        if es_primera:
            portadas = [s for s in sheet_names if "PORTADA" in s.upper()]
            if portadas:
                wp = wb[portadas[0]]
                for row in wp.iter_rows():
                    for cell in row:
                        if cell.value and len(str(cell.value).strip()) > 5:
                            cell.value = datos["titulo_doc"]
                            log.append(f"✅ Portada → título escrito en {cell.coordinate}")
                            break
                    else:
                        continue
                    break

        # ── CONTROL DE FIRMAS (solo rev 00) ──────────────────────────────────
        if es_primera:
            ctrl = [s for s in sheet_names if "CONTROL" in s.upper() or "FIRMA" in s.upper()]
            if ctrl:
                wf = wb[ctrl[0]]
                wf["B5"].value = "\n" + datos["nombre_realizado"]
                wf["D5"].value = " "  + datos["nombre_revisado"]
                wf["F5"].value = datos["nombre_aprobado"]
                wf["B9"].value = f"Fecha\n{datos['fecha_firmas']}"
                wf["D9"].value = f"Fecha\n{datos['fecha_firmas']}"
                wf["F9"].value = f"Fecha\n{datos['fecha_firmas']}"
                log.append("✅ Control de Firmas → nombres y fechas escritos")

        # ── REGISTRO DE CAMBIOS (todas las revisiones) ────────────────────────
        ctrl = [s for s in sheet_names if "CONTROL" in s.upper() or "FIRMA" in s.upper()]
        if ctrl:
            wf = wb[ctrl[0]]
            # La fila de datos empieza en 14; buscar la primera fila vacía desde ahí
            fila_cambio = 14
            for r in range(14, 30):
                if wf.cell(r, 2).value is None:
                    fila_cambio = r
                    break
            wf.cell(fila_cambio, 2).value = num_rev
            wf.cell(fila_cambio, 3).value = datos["fecha_ciclo"]
            wf.cell(fila_cambio, 4).value = " " + datos["autor_cambio"]
            wf.cell(fila_cambio, 5).value = datos["seccion_afectada"]
            wf.cell(fila_cambio, 6).value = datos["desc_cambio"]
            log.append(f"✅ Registro de cambios → fila {fila_cambio} (rev {num_rev})")

        # ── CABECERAS HRC (todas las revisiones) ──────────────────────────────
        hrc_sheets = [s for s in sheet_names if "HRC" in s.upper()]
        for sname in hrc_sheets:
            ws = wb[sname]

            # Campos fijos del header (siempre se actualizan)
            ws["F7"].value  = datos["cod_rev_general"]
            ws["F9"].value  = datos["clave_ref"]
            ws["F10"].value = datos["titulo_cab"]
            ws["F11"].value = datos["cod_documento"]
            ws["F12"].value = datos["originador"]
            ws["J11"].value = datos["revisor_ppal"]
            ws["J12"].value = datos["entidad_rev"]

            # Fecha del ciclo correcto — no toca los otros ciclos
            ws[f"{ciclo_col}11"].value = datos["fecha_ciclo"]

            # Primer ciclo: marcar originador como NA si es 00
            if es_primera:
                ws["N10"].value = "NA"

            log.append(f"✅ {sname} → cabecera actualizada (ciclo col {ciclo_col})")

        # ── GUARDAR ───────────────────────────────────────────────────────────
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success(f"✅ Procesadas {len(hrc_sheets)} hojas HRC  |  Revisión S{num_rev}")
        with st.expander("Log detallado"):
            for line in log:
                st.text(line)

        nombre_salida = f"HRC_S{num_rev}_llenado.xlsx"
        st.download_button(
            label=f"⬇️ Descargar {nombre_salida}",
            data=output,
            file_name=nombre_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error procesando la plantilla: {e}")
        st.exception(e)
