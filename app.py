import re
import io
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# =========================
# Configuración de la app
# =========================
st.set_page_config(page_title="IRIZAR – Fabricación", layout="wide")
st.title("IRIZAR – Resumen de Fabricación")
st.caption("Sube los ficheros, pulsa **Procesar** y descarga el Excel consolidado.")

# ==============
# Utilidades
# ==============
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(r"\s+", "_", regex=True)
        .str.replace(r"[^\w\.]", "", regex=True)
        .str.lower()
    )
    return df

def extract_ref_base(val) -> str:
    """Devuelve 7 dígitos consecutivos como código base (p.ej. 8004778 de 'CER8004778-23')."""
    if pd.isna(val):
        return np.nan
    s = str(val).replace(".", "").replace(" ", "")
    m = re.search(r"(\d{7})", s)
    return m.group(1) if m else np.nan

def to_number(x):
    """Convierte '1.234,56' / '1234,00' / '1234' → float, NaN seguro."""
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace(" ", "")
    if "." in s and "," in s and s.index(".") < s.index(","):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return np.nan

def read_excel_any(file_like) -> pd.DataFrame:
    """Lee la primera hoja no vacía y normaliza cabeceras."""
    xl = pd.ExcelFile(file_like, engine="openpyxl")
    for sheet in xl.sheet_names:
        df = pd.read_excel(file_like, sheet_name=sheet, engine="openpyxl")
        df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
        if df.shape[0] and df.shape[1]:
            return normalize_columns(df)
    return pd.DataFrame()

def parse_date_from_filename(name: str):
    """
    Detecta fecha (y hora) en nombres tipo:
    '16 09 2025', '23-09-25', '23_09_2025 11 00H', '23-09-2025 11-00', etc.
    """
    s = name.lower().replace("_", " ").replace(".", " ").replace(",", " ")
    m = re.search(r"(\d{1,2})[-\s](\d{1,2})[-\s](\d{2,4})(?:[\s\-](\d{1,2})[\s:](\d{1,2}))?", s)
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    hh = int(m.group(4)) if m.group(4) else 0
    mm = int(m.group(5)) if m.group(5) else 0
    if y < 100:
        y = 2000 + y
    try:
        return datetime(y, mo, d, hh, mm)
    except:
        return None

def format_iso(d: datetime) -> str:
    return d.strftime("%Y-%m-%d") if d else datetime.now().strftime("%Y-%m-%d")

def find_col(cands, df_cols):
    for c in cands:
        if c in df_cols:
            return c
    return None

def suggest_numeric_neighbors(ref_str, maes_df, topn=3):
    """Sugerencias simples por cercanía numérica del código."""
    try:
        r = int(ref_str)
    except:
        return ""
    maes_df = maes_df.dropna(subset=["ref_base"]).copy()
    maes_df["code_int"] = maes_df["ref_base"].astype(int)
    maes_df["delta"] = (maes_df["code_int"] - r).abs()
    top = maes_df.nsmallest(topn, "delta")
    tips = [f"{row.ref_base} - {row.descripcion_maestro}" for _, row in top.iterrows()]
    return "; ".join(tips)

# =========================
# UI – Subida de ficheros
# =========================
with st.sidebar:
    st.header("1) Sube los Excel")
    up_log_prev = st.file_uploader("Stock Logeten – Semana anterior (.xlsx)", type=["xlsx"], key="log_prev")
    up_log_curr = st.file_uploader("Stock Logeten – Semana actual (.xlsx)", type=["xlsx"], key="log_curr")
    up_fabrica  = st.file_uploader("Stock en fábrica (.xlsx)", type=["xlsx"], key="fab")
    up_pend     = st.file_uploader("Pendientes de servir (.xlsx)", type=["xlsx"], key="pen")
    up_ordenes  = st.file_uploader("Órdenes de fabricación (.xlsx)", type=["xlsx"], key="ord")
    up_maestro  = st.file_uploader("Maestro de referencias (.xlsx)", type=["xlsx"], key="mae")

    st.subheader("Opcional")
    up_resumen_prev = st.file_uploader("Resumen anterior (para reusar Comentarios) (.xlsx)", type=["xlsx"], key="prev")
    up_mapeos       = st.file_uploader("Mapeos_usuario (.xlsx con hoja 'Mapeos_usuario')", type=["xlsx"], key="map")

    st.header("2) Opciones")
    add_suggestions = st.checkbox("Añadir sugerencias para referencias no encontradas en el maestro", value=True)
    st.header("3) Procesar")
    go = st.button("Procesar y generar Excel", type="primary", use_container_width=True)

# =======================================
# LÓGICA PRINCIPAL (cuando pulsas botón)
# =======================================
if go:
    # Validaciones mínimas
    missing = [name for name, f in {
        "Logeten anterior": up_log_prev,
        "Logeten actual": up_log_curr,
        "Fábrica": up_fabrica,
        "Pendientes": up_pend,
        "Órdenes": up_ordenes,
        "Maestro": up_maestro,
    }.items() if f is None]

    if missing:
        st.error("Faltan ficheros obligatorios: " + ", ".join(missing))
        st.stop()

    with st.spinner("Leyendo y normalizando datos..."):
        # --- Logeten anterior
        log_prev = read_excel_any(up_log_prev)
        cols_prev = log_prev.columns.tolist()
        code_col_prev = find_col(["codigo","cod","ref","referencia"], cols_prev) or (cols_prev[0] if cols_prev else None)
        qty_col_prev  = find_col(["00cantidad","cantidad","unidades"], cols_prev) or (cols_prev[-1] if cols_prev else None)
        desc_col_prev = next((c for c in cols_prev if "desc" in c), None)
        if not log_prev.empty:
            log_prev["ref_base"] = log_prev[code_col_prev].apply(extract_ref_base)
            log_prev["stock_logeten_prev"] = log_prev[qty_col_prev].apply(to_number)
            if desc_col_prev: log_prev["desc_fuente"] = log_prev[desc_col_prev]
            prev_agg = log_prev.groupby("ref_base", as_index=False).agg({
                "stock_logeten_prev":"sum",
                **({"desc_fuente":"first"} if "desc_fuente" in log_prev.columns else {})
            })
        else:
            prev_agg = pd.DataFrame(columns=["ref_base","stock_logeten_prev"])

        # --- Logeten actual
        log_curr = read_excel_any(up_log_curr)
        cols_curr = log_curr.columns.tolist()
        code_col_curr = find_col(["codigo","cod","ref","referencia"], cols_curr) or (cols_curr[0] if cols_curr else None)
        qty_col_curr  = find_col(["00cantidad","cantidad","unidades"], cols_curr) or (cols_curr[-1] if cols_curr else None)
        desc_col_curr = next((c for c in cols_curr if "desc" in c), None)
        if not log_curr.empty:
            log_curr["ref_base"] = log_curr[code_col_curr].apply(extract_ref_base)
            log_curr["stock_logeten_curr"] = log_curr[qty_col_curr].apply(to_number)
            if desc_col_curr: log_curr["desc_fuente"] = log_curr[desc_col_curr]
            curr_agg = log_curr.groupby("ref_base", as_index=False).agg({
                "stock_logeten_curr":"sum",
                **({"desc_fuente":"first"} if "desc_fuente" in log_curr.columns else {})
            })
        else:
            curr_agg = pd.DataFrame(columns=["ref_base","stock_logeten_curr"])

        # --- Fábrica
        fab = read_excel_any(up_fabrica)
        cols_fab = fab.columns.tolist()
        code_col_fab = find_col(["ref.","ref","referencia","codigo","código"], cols_fab) or (cols_fab[0] if cols_fab else None)
        qty_col_fab  = find_col(["cantidad","cant","qty","unidades"], cols_fab) or (cols_fab[-1] if cols_fab else None)
        desc_col_fab = next((c for c in cols_fab if "desc" in c), None)
        if not fab.empty:
            fab["ref_base"] = fab[code_col_fab].apply(extract_ref_base)
            fab["stock_fabrica"] = fab[qty_col_fab].apply(to_number)
            if desc_col_fab: fab["desc_fuente"] = fab[desc_col_fab]
            fab_agg = fab.groupby("ref_base", as_index=False).agg({
                "stock_fabrica":"sum",
                **({"desc_fuente":"first"} if "desc_fuente" in fab.columns else {})
            })
        else:
            fab_agg = pd.DataFrame(columns=["ref_base","stock_fabrica"])

        # --- Pendientes de servir
        pen = read_excel_any(up_pend)
        cols_pen = pen.columns.tolist()
        col_art = find_col(["articulo","artículo","articulo.","art","ref","referencia","codigo"], cols_pen) or (cols_pen[0] if cols_pen else None)
        col_pend = find_col(["pendiente","pendientes"], cols_pen)
        if not col_pend:
            col_cant = find_col(["cantidad","cant"], cols_pen)
            col_serv = find_col(["servido","servidos"], cols_pen)
            if col_cant and col_serv:
                pen["__cant__"] = pen[col_cant].apply(to_number)
                pen["__serv__"] = pen[col_serv].apply(to_number)
                pen["pendiente"] = pen["__cant__"] - pen["__serv__"]
                col_pend = "pendiente"
            else:
                col_pend = cols_pen[-1] if cols_pen else None
        desc_col_pen = next((c for c in cols_pen if "desc" in c), None)
        if not pen.empty:
            pen["ref_base"] = pen[col_art].apply(extract_ref_base)
            pen["pendiente_u"] = pen[col_pend].apply(to_number)
            if desc_col_pen: pen["desc_fuente"] = pen[desc_col_pen]
            pen_valid = pen[~pen["ref_base"].isna()].copy()
            pend_agg = pen_valid.groupby("ref_base", as_index=False).agg({
                "pendiente_u":"sum",
                **({"desc_fuente":"first"} if "desc_fuente" in pen_valid.columns else {})
            })
        else:
            pend_agg = pd.DataFrame(columns=["ref_base","pendiente_u"])

        # --- Órdenes
        ordf = read_excel_any(up_ordenes)
        cols_ord = ordf.columns.tolist()
        col_ref_o = find_col(["referencia","ref","articulo","codigo","código"], cols_ord) or (cols_ord[0] if cols_ord else None)
        col_qty_o = find_col(["cantidad","cant","cantidad_pendiente","u_pendientes"], cols_ord)
        if not col_qty_o:
            num_cols = [c for c in cols_ord if pd.api.types.is_numeric_dtype(ordf[c])]
            col_qty_o = num_cols[0] if num_cols else (cols_ord[-1] if len(cols_ord) else None)
        col_fent  = next((c for c in cols_ord if "entrega" in c or "f_entrega" in c or "fechaentrega" in c), None)
        desc_col_ord = next((c for c in cols_ord if "desc" in c), None)
        if not ordf.empty:
            ordf["ref_base"] = ordf[col_ref_o].apply(extract_ref_base)
            ordf["ordenes_u"] = ordf[col_qty_o].apply(to_number)
            if col_fent:
                def _to_date(x):
                    if pd.isna(x): return pd.NaT
                    if isinstance(x, (pd.Timestamp, datetime)): return pd.to_datetime(x)
                    try: return pd.to_datetime(x, dayfirst=True, errors="coerce")
                    except: return pd.NaT
                ordf["f_entrega"] = ordf[col_fent].apply(_to_date)
            if desc_col_ord: ordf["desc_fuente"] = ordf[desc_col_ord]
            ordf_valid = ordf[~ordf["ref_base"].isna()].copy()
            ordenes_agg = ordf_valid.groupby("ref_base", as_index=False).agg({
                "ordenes_u":"sum",
                **({"f_entrega":"min"} if "f_entrega" in ordf_valid.columns else {}),
                **({"desc_fuente":"first"} if "desc_fuente" in ordf_valid.columns else {})
            })
            ordenes_agg["n_ordenes"] = ordf_valid.groupby("ref_base").size().values
        else:
            ordenes_agg = pd.DataFrame(columns=["ref_base","ordenes_u"])

        # --- Maestro
        maestro = read_excel_any(up_maestro)
        cols_mae = maestro.columns.tolist()
        col_ref_m = find_col(["codigo","código","ref","referencia"], cols_mae) or (cols_mae[0] if len(cols_mae) else None)
        col_desc_m = next((c for c in cols_mae if "desc" in c), None)
        if maestro.empty:
            maestro_agg = pd.DataFrame(columns=["ref_base","descripcion_maestro"])
        else:
            maestro["ref_base"] = maestro[col_ref_m].apply(extract_ref_base)
            if not col_desc_m and len(cols_mae) > 1:
                col_desc_m = cols_mae[1]
            maestro_agg = maestro[["ref_base", col_desc_m]].drop_duplicates()
            maestro_agg = maestro_agg.rename(columns={col_desc_m: "descripcion_maestro"})

        # --- Comentarios previos
        comentarios_prev = pd.DataFrame(columns=["Referencia","Comentarios"])
        if up_resumen_prev is not None:
            try:
                xl_prev = pd.ExcelFile(up_resumen_prev, engine="openpyxl")
                if "Resumen" in xl_prev.sheet_names:
                    tmp = pd.read_excel(up_resumen_prev, sheet_name="Resumen", engine="openpyxl")
                    if "Referencia" in tmp.columns and "Comentarios" in tmp.columns:
                        comentarios_prev = tmp[["Referencia","Comentarios"]].dropna()
            except:
                pass

        # --- Mapeos_usuario
        mapeos_usuario = pd.DataFrame(columns=["origen","destino"])
        if up_mapeos is not None:
            try:
                xl_map = pd.ExcelFile(up_mapeos, engine="openpyxl")
                if "Mapeos_usuario" in xl_map.sheet_names:
                    mp = pd.read_excel(up_mapeos, sheet_name="Mapeos_usuario", engine="openpyxl").dropna(how="all")
                    mp = mp.rename(columns={c:c.lower() for c in mp.columns})
                    if "origen" in mp.columns and "destino" in mp.columns:
                        mapeos_usuario = mp[["origen","destino"]]
            except:
                pass

        # Aplica mapeos (si existen): origen_base -> destino_base
        if not mapeos_usuario.empty:
            mapeos_usuario["origen_base"]  = mapeos_usuario["origen"].apply(extract_ref_base)
            mapeos_usuario["destino_base"] = mapeos_usuario["destino"].apply(extract_ref_base)
            mapeos_usuario = mapeos_usuario.dropna(subset=["origen_base","destino_base"]).drop_duplicates()
            mp = dict(zip(mapeos_usuario["origen_base"], mapeos_usuario["destino_base"]))
            for df in ["prev_agg","curr_agg","fab_agg","pend_agg","ordenes_agg"]:
                if df in locals() and not locals()[df].empty:
                    locals()[df]["ref_base"] = locals()[df]["ref_base"].replace(mp)

    with st.spinner("Cruzando, calculando y preparando salida..."):
        # Universo + merge
        refs = pd.concat([
            prev_agg[["ref_base"]],
            curr_agg[["ref_base"]],
            fab_agg[["ref_base"]],
            pend_agg[["ref_base"]],
            ordenes_agg[["ref_base"]],
            maestro_agg[["ref_base"]],
        ], ignore_index=True).dropna().drop_duplicates()

        resumen = (refs
            .merge(prev_agg, on="ref_base", how="left")
            .merge(curr_agg, on="ref_base", how="left")
            .merge(fab_agg,  on="ref_base", how="left")
            .merge(pend_agg, on="ref_base", how="left")
            .merge(ordenes_agg, on="ref_base", how="left")
            .merge(maestro_agg, on="ref_base", how="left")
        )

        for c in ["stock_logeten_prev","stock_logeten_curr","stock_fabrica","pendiente_u","ordenes_u"]:
            if c in resumen.columns:
                resumen[c] = resumen[c].fillna(0.0)

        # Descripción priorizando maestro
        desc_candidates = [c for c in ["desc_fuente_x","desc_fuente_y","desc_fuente"] if c in resumen.columns]
        def pick_desc(row):
            if pd.notna(row.get("descripcion_maestro", np.nan)):
                return row["descripcion_maestro"]
            for d in desc_candidates:
                v = row.get(d, np.nan)
                if pd.notna(v) and str(v).strip():
                    return v
            return ""
        resumen["Descripcion"] = resumen.apply(pick_desc, axis=1)

        # Cálculos
        resumen["diferencia_stock"] = resumen["stock_logeten_prev"] - resumen["stock_logeten_curr"]
        resumen["situacion_actual"] = resumen["pendiente_u"] - (resumen["stock_fabrica"] + resumen["ordenes_u"])
        resumen["En maestro"] = np.where(resumen["descripcion_maestro"].notna(), "Sí", "No")
        resumen["Nº órdenes"] = resumen.get("n_ordenes", pd.Series([np.nan]*len(resumen)))
        resumen["Próxima entrega"] = resumen.get("f_entrega", pd.Series([pd.NaT]*len(resumen)))

        # Comentarios previos
        resumen["Referencia"] = resumen["ref_base"]
        resumen["Comentarios"] = ""
        if not comentarios_prev.empty:
            comentarios_prev = comentarios_prev.dropna().drop_duplicates(subset=["Referencia"], keep="last")
            resumen = resumen.merge(comentarios_prev, on="Referencia", how="left", suffixes=("","_prev"))
            resumen["Comentarios"] = resumen["Comentarios_prev"].combine_first(resumen["Comentarios"])
            resumen = resumen.drop(columns=[c for c in resumen.columns if c.endswith("_prev")], errors="ignore")

        # Fechas en cabecera
        date_prev = parse_date_from_filename(getattr(up_log_prev, "name", "")) if up_log_prev else None
        date_curr = parse_date_from_filename(getattr(up_log_curr, "name", "")) if up_log_curr else None
        date_fab  = parse_date_from_filename(getattr(up_fabrica,  "name", "")) if up_fabrica  else None

        col_prev = f"Stock semana anterior (Logeten) [{date_prev.strftime('%d/%m/%Y')}]" if date_prev else "Stock semana anterior (Logeten)"
        col_curr = f"Stock semana actual (Logeten) [{date_curr.strftime('%d/%m/%Y')}]" if date_curr else "Stock semana actual (Logeten)"
        col_fab  = f"Stock en fábrica (u) [{date_fab.strftime('%d/%m/%Y %H:%M') if date_fab and (date_fab.hour or date_fab.minute) else (date_fab.strftime('%d/%m/%Y') if date_fab else '')}]".strip()
        if col_fab.endswith("[]"):
            col_fab = "Stock en fábrica (u)"

        rename_map = {
            "stock_logeten_prev": col_prev,
            "stock_logeten_curr": col_curr,
            "diferencia_stock": "Diferencia de stock (ant - act)",
            "pendiente_u": "Referencias acumuladas pendientes de servir (u)",
            "stock_fabrica": col_fab,
            "ordenes_u": "Órdenes de fabricación en curso (u)",
            "situacion_actual": "Situación actual = Pendiente − (Fábrica + Órdenes)",
        }
        resumen = resumen.rename(columns=rename_map)

        final_cols = [
            "Referencia",
            "Descripcion",
            col_prev,
            col_curr,
            "Diferencia de stock (ant - act)",
            "Referencias acumuladas pendientes de servir (u)",
            col_fab,
            "Órdenes de fabricación en curso (u)",
            "Situación actual = Pendiente − (Fábrica + Órdenes)",
            "Comentarios",
            "En maestro",
            "Nº órdenes",
            "Próxima entrega",
        ]
        final_cols = [c for c in final_cols if c in resumen.columns] + [c for c in resumen.columns if c not in final_cols]
        resumen = resumen[final_cols].sort_values(by="Situación actual = Pendiente − (Fábrica + Órdenes)", ascending=False)

        # No_en_maestro + sugerencias
        maes_df = maestro_agg.copy()
        no_mae = resumen[resumen["En maestro"]=="No"][["Referencia","Descripcion"]].copy()
        if add_suggestions and not no_mae.empty and not maes_df.empty:
            no_mae["Sugerencias"] = no_mae["Referencia"].apply(lambda x: suggest_numeric_neighbors(x, maes_df, 3))
        else:
            no_mae["Sugerencias"] = ""

        # Auditoría
        auditoria_df = pd.DataFrame({
            "archivo": [
                getattr(up_log_prev,"name",""), getattr(up_log_curr,"name",""),
                getattr(up_fabrica,"name",""), getattr(up_pend,"name",""),
                getattr(up_ordenes,"name",""), getattr(up_maestro,"name","")
            ],
            "fecha_detectada": [
                str(date_prev), str(date_curr), str(date_fab), "", "", ""
            ]
        })

        # Excel en memoria con formato
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            resumen.to_excel(writer, sheet_name="Resumen", index=False)
            auditoria_df.to_excel(writer, sheet_name="Auditoria", index=False)
            if not no_mae.empty:
                no_mae.to_excel(writer, sheet_name="No_en_maestro", index=False)
            # Plantilla Mapeos_usuario
            if up_mapeos is None:
                pd.DataFrame({"origen": [], "destino": []}).to_excel(writer, sheet_name="Mapeos_usuario", index=False)

            # Formatos
            wb  = writer.book
            ws  = writer.sheets["Resumen"]
            fmt_int = wb.add_format({'num_format': '#.##0', 'align':'right'})
            fmt_head= wb.add_format({'bold': True})
            fmt_red = wb.add_format({'bg_color': '#FFC7CE'})
            fmt_green = wb.add_format({'bg_color': '#C6EFCE'})

            ws.freeze_panes(1, 2)
            for i, col in enumerate(resumen.columns):
                max_len = max(10, resumen[col].astype(str).str.len().clip(upper=60).max() if len(resumen) else 10)
                ws.set_column(i, i, min(45, max_len + 2))
            ws.set_row(0, None, fmt_head)

            int_cols = [
                col_prev, col_curr, "Diferencia de stock (ant - act)",
                "Referencias acumuladas pendientes de servir (u)",
                col_fab, "Órdenes de fabricación en curso (u)",
                "Situación actual = Pendiente − (Fábrica + Órdenes)", "Nº órdenes"
            ]
            for c in int_cols:
                if c in resumen.columns:
                    j = resumen.columns.get_loc(c)
                    ws.set_column(j, j, None, fmt_int)

            if "Situación actual = Pendiente − (Fábrica + Órdenes)" in resumen.columns:
                j = resumen.columns.get_loc("Situación actual = Pendiente − (Fábrica + Órdenes)")
                nrows = len(resumen) + 1
                ws.conditional_format(1, j, nrows, j, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red})
                ws.conditional_format(1, j, nrows, j, {'type': 'cell', 'criteria': '<=', 'value': 0, 'format': fmt_green})

        # Nombre de salida: "Fecha stock ultimo fichero_Fabricacion_Irizar.xlsx"
        dates_for_naming = [d for d in [date_prev, date_curr, date_fab] if d]
        out_date = max(dates_for_naming) if dates_for_naming else datetime.now()
        out_name = f"{format_iso(out_date)}_Fabricacion_Irizar.xlsx"

    st.success("¡Listo! Descarga el Excel consolidado:")
    st.download_button(
        label="⬇️ Descargar Excel",
        data=buffer.getvalue(),
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()
    st.subheader("Vista previa rápida (primeras filas)")
    st.dataframe(resumen.head(25), use_container_width=True)
