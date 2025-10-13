# app.py
import io
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------
# Helpers
# -----------------------------
def normalize_dni_value(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    s = re.sub(r'\D', '', s)
    return s if s else pd.NA

def normalize_email(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip().lower()
    s = s.replace("mailto:", "").strip(" ,;.")
    s = re.sub(r"\s+", "", s)
    return s if ("@" in s and "." in s) else pd.NA

def extract_year(value):
    if pd.isna(value):
        return pd.NA
    s = str(value)
    m = re.search(r'(20\d{2})', s)
    return int(m.group(1)) if m else pd.NA

def strip_accents(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    return "".join([c for c in s if not unicodedata.combining(c)]).lower()

def find_col(df: pd.DataFrame, keywords, prefer_exact=None):
    cols = [c for c in df.columns if isinstance(c, str)]
    norm = {strip_accents(c): c for c in cols}
    if prefer_exact:
        for pe in prefer_exact:
            key = strip_accents(pe)
            if key in norm:
                return norm[key]
    for c in cols:
        cu = strip_accents(c)
        if any(k in cu for k in keywords):
            return c
    return None

def split_fullname(fullname):
    if pd.isna(fullname):
        return pd.NA, pd.NA
    s = str(fullname).strip()
    if not s:
        return pd.NA, pd.NA
    parts = s.split()
    if len(parts) == 1:
        return parts[0], pd.NA
    return parts[0], " ".join(parts[1:])

def drop_emailish_cols(df: pd.DataFrame, keep_email=True) -> pd.DataFrame:
    """
    Remove all columns that look like email headers except the canonical 'Email'.
    Prevents duplicate 'Dirección de correo_x' columns during merges.
    """
    to_drop = []
    for c in list(df.columns):
        cu = strip_accents(c)
        looks_like_email = (
            "correo" in cu or cu in {"email", "e-mail", "mail", "direcciondecorreo"}
        )
        if looks_like_email and not (keep_email and c == "Email"):
            to_drop.append(c)
    if to_drop:
        df = df.drop(columns=to_drop, errors="ignore")
    return df

def load_teacher_email_roster(roster_file):
    wb = pd.read_excel(roster_file, sheet_name=0, dtype=str, engine="openpyxl")
    df = list(wb.values())[0] if isinstance(wb, dict) else wb

    email_col = find_col(
        df, ["correo", "email", "e-mail", "mail"],
        prefer_exact=["Dirección de correo", "Correo", "Email", "E-mail"]
    )
    if email_col is None:
        for c in df.columns:
            frac = df[c].astype(str).str.contains("@", na=False).mean()
            if frac > 0.3:
                email_col = c
                break
    if email_col is None:
        return pd.DataFrame(columns=["Email", "Nombre", "Apellido(s)"])

    nombres_col = find_col(df, ["nombres", "nombre"], prefer_exact=["NOMBRES", "Nombre"])
    a_pat_col   = find_col(df, ["apellido paterno", "a. paterno"])
    a_mat_col   = find_col(df, ["apellido materno", "a. materno"])
    full_col    = find_col(df, ["docente", "profesor", "nombre completo",
                                "nombres y apellidos", "apellidos y nombres"])

    out = pd.DataFrame()
    out["Email"] = df[email_col].apply(normalize_email)

    if nombres_col and a_pat_col and a_mat_col:
        out["Nombre"] = df[nombres_col].astype(str).str.strip()
        out["Apellido(s)"] = (
            df[a_pat_col].astype(str).str.strip() + " " +
            df[a_mat_col].astype(str).str.strip()
        ).str.replace(r"\s+", " ", regex=True).str.strip()
    elif full_col:
        tmp = df[full_col].apply(split_fullname)
        out["Nombre"] = tmp.apply(lambda t: t[0])
        out["Apellido(s)"] = tmp.apply(lambda t: t[1])
    else:
        out["Nombre"] = pd.NA
        out["Apellido(s)"] = pd.NA

    out = out.dropna(subset=["Email"]).drop_duplicates(subset=["Email"])
    out = drop_emailish_cols(out)  # keep only canonical 'Email'
    return out[["Email", "Nombre", "Apellido(s)"]]

def read_workbook(file):
    raw = pd.read_excel(file, sheet_name=None, dtype=str, engine="openpyxl")
    book = {}
    for name, df in raw.items():
        key = strip_accents(name).replace(" ", "")
        book[key] = df.copy()
    return book

def get_sheet(book, possible_names):
    keys = list(book.keys())
    for candidate in possible_names:
        cand = strip_accents(candidate).replace(" ", "")
        for k in keys:
            if cand == k or cand in k or k in cand:
                return book[k].copy()
    return None

def coalesce_keys(df: pd.DataFrame) -> pd.DataFrame:
    """Collapse *_x/*_y to clean keys and remove leftover email-ish headers."""
    for key in ["Email", "DNI", "Nombre", "Apellido(s)"]:
        kx, ky = f"{key}_x", f"{key}_y"
        if kx in df.columns or ky in df.columns:
            if key not in df.columns:
                df[key] = pd.NA
            if kx in df.columns:
                df[key] = df[key].where(df[key].notna(), df[kx])
            if ky in df.columns:
                df[key] = df[key].where(df[key].notna(), df[ky])
            df = df.drop(columns=[c for c in [kx, ky] if c in df.columns])
    return drop_emailish_cols(df)  # also purge any new email-ish headers

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(master_file, roster_file=None):
    # Read master workbook once
    book = read_workbook(master_file)
    df_induccion      = get_sheet(book, ["inducción", "induccion"])
    df_nota_ind       = get_sheet(book, ["nota inducción", "nota induccion"])
    df_biblio         = get_sheet(book, ["bus. biblioteca", "biblioteca", "bus biblioteca"])
    df_diseno         = get_sheet(book, ["diseño de sesión", "diseno de sesion", "diseño sesion"])
    df_comptec        = get_sheet(book, ["comp. tec", "competencias tecn", "comp tec"])
    df_integracion    = get_sheet(book, ["integración", "integracion"])
    df_rsu            = get_sheet(book, ["rsu"])
    df_estress        = get_sheet(book, ["estress", "estres", "estrés"])
    df_habcom         = get_sheet(book, ["hab. comunicación", "hab comunicacion", "habilidades comunicacion"])

    if df_induccion is None and df_nota_ind is None:
        raise ValueError("Neither 'Inducción' nor 'nota Inducción' sheet found in Master file.")

    # ---- Base union
    frames = []
    if df_nota_ind is not None:
        cols_keep = ["PERIODO", "DNI", "Nombre", "Apellido(s)", "Dirección de correo", "Total del curso (Real)"]
        exist = [c for c in cols_keep if c in df_nota_ind.columns]
        tmp = df_nota_ind[exist].copy()
        if "PERIODO" in tmp.columns:
            tmp = tmp.rename(columns={"PERIODO": "Periodo"})
        if "Total del curso (Real)" in tmp.columns:
            tmp = tmp.rename(columns={"Total del curso (Real)": "induccion"})
        frames.append(tmp)

    if df_induccion is not None:
        cols_keep = ["Periodo", "DNI", "Nombre", "Apellido(s)", "Dirección de correo", "Calificación"]
        exist = [c for c in cols_keep if c in df_induccion.columns]
        tmp = df_induccion[exist].copy()
        if "Calificación" in tmp.columns:
            tmp = tmp.rename(columns={"Calificación": "induccion"})
        frames.append(tmp)

    all_data = pd.concat(frames, ignore_index=True)

    # Canonical IDs
    all_data["DNI"] = all_data["DNI"].apply(normalize_dni_value) if "DNI" in all_data.columns else pd.NA
    if "Dirección de correo" in all_data.columns:
        all_data["Email"] = all_data["Dirección de correo"].apply(normalize_email)
    else:
        all_data["Email"] = pd.NA
    all_data["Year"] = all_data["Periodo"].apply(extract_year) if "Periodo" in all_data.columns else pd.NA

    # Drop all non-canonical email headers from base
    all_data = drop_emailish_cols(all_data)

    # ---- Merge utilities
    def smart_merge(left: pd.DataFrame, right: pd.DataFrame, on, how="left", restrict_cols=None):
        """
        Merge using 'on' (string or list). Before merging, drop any email-ish columns
        from 'right' and optionally restrict to needed columns only.
        """
        right = drop_emailish_cols(right)
        if restrict_cols is not None:
            keep = [c for c in restrict_cols if c in right.columns]
            if keep:
                right = right[keep].copy()
        out = pd.merge(left, right, on=on, how=how)  # no explicit suffixes; keep clean
        return coalesce_keys(out)

    def add_by_key(src_df, score_cols_map, join_order):
        """
        Try merges in order:
          join_order: list of dicts like {"on": "Email"} or {"on": ["Nombre","Apellido(s)"]} or {"on": "DNI"}
        score_cols_map: dict {existing_col_in_src: new_col_name}
        Only bring in required columns: join keys + score columns.
        """
        nonlocal all_data
        if src_df is None:
            return

        src = src_df.copy()
        # rename score columns we know
        for old, new in list(score_cols_map.items()):
            if old in src.columns:
                src = src.rename(columns={old: new})

        # Prepare canonical Email/DNI in source if present
        email_in_src = find_col(src, ["correo", "email", "e-mail", "mail"])
        if email_in_src:
            src["Email"] = src[email_in_src].apply(normalize_email)
        if "DNI" in src.columns:
            src["DNI"] = src["DNI"].apply(normalize_dni_value)

        # Build the minimal column set to carry into the merge
        score_cols = list(score_cols_map.values())
        minimal_cols = set(score_cols)
        # names may be needed for name-join
        if "Nombre" in src.columns: minimal_cols.add("Nombre")
        if "Apellido(s)" in src.columns: minimal_cols.add("Apellido(s)")
        if "Email" in src.columns: minimal_cols.add("Email")
        if "DNI" in src.columns: minimal_cols.add("DNI")

        # try in order
        for opt in join_order:
            key = opt["on"]
            # ensure key(s) exist on both sides
            if isinstance(key, list):
                if not all(col in all_data.columns for col in key): 
                    continue
                if not all(col in src.columns for col in key):
                    continue
                restrict = list(minimal_cols.union(key))
                all_data = smart_merge(all_data, src, on=key, restrict_cols=restrict)
                return
            else:
                if key not in all_data.columns or key not in src.columns:
                    continue
                restrict = list(minimal_cols.union([key]))
                all_data = smart_merge(all_data, src, on=key, restrict_cols=restrict)
                return
        # if no join matched, skip silently

    # ---- Bring in each sheet
    # Bus. biblioteca
    if df_biblio is not None:
        prom = find_col(df_biblio, ["promedio"]) or "Promedio"
        add_by_key(
            df_biblio,
            score_cols_map={prom: "bus_biblioteca"} if prom in df_biblio.columns else {},
            join_order=[{"on": "Email"}, {"on": "DNI"}, {"on": ["Nombre", "Apellido(s)"]}]
        )

    # Diseño de sesión
    if df_diseno is not None:
        prom_col = find_col(df_diseno, ["promedio"]) or "Promedio"
        add_by_key(
            df_diseno,
            score_cols_map={prom_col: "diseno_sesion"} if prom_col in df_diseno.columns else {},
            join_order=[{"on": ["Nombre", "Apellido(s)"]}, {"on": "Email"}, {"on": "DNI"}]
        )

    # Comp. Tec
    if df_comptec is not None:
        ren = {
            "Cuestionario:Reto: Zoom básico": "Zoom_basico",
            "Cuestionario:Reto: Zoom Avanzado": "Zoom_Avanzado",
            "Cuestionario:Reto: Grupos Moodle": "Grupos_Moodle",
            "Cuestionario:Reto: Rúbrica": "Rubrica",
            "Cuestionario:Reto: Padlet": "Padlet",
            "Cuestionario:Reto: Nearpod": "Nearpod",
            "Cuestionario:Reto: Tareas y foros": "Tareas_y_foros",
        }
        add_by_key(
            df_comptec,
            score_cols_map={k: v for k, v in ren.items() if k in df_comptec.columns},
            join_order=[{"on": ["Nombre", "Apellido(s)"]}, {"on": "Email"}, {"on": "DNI"}]
        )

    # Integración
    if df_integracion is not None:
        integ_col = None
        for cand in [
            "Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)",
            "Tarea:Producto final", "Producto final", "Integración", "Integracion"
        ]:
            if cand in df_integracion.columns:
                integ_col = cand
                break
        add_by_key(
            df_integracion,
            score_cols_map={integ_col: "integracion"} if integ_col else {},
            join_order=[{"on": ["Nombre", "Apellido(s)"]}, {"on": "Email"}, {"on": "DNI"}]
        )

    # RSU / Estress / Hab. comunicación
    def pick_score_col(df_src, default_old):
        # look for tarea/promedio-like column
        for c in df_src.columns:
            cu = strip_accents(c)
            if cu.startswith("tarea") or "promedio" in cu:
                return c
        return default_old if default_old in df_src.columns else None

    for sub_df, out_col, default_old in [
        (df_rsu, "rsu", "Tarea: Producto final"),
        (df_estress, "estress", "Tarea:Producto final"),
        (df_habcom, "hab_comunicacion", "Tarea:Producto final"),
    ]:
        if sub_df is None:
            continue
        score_col = pick_score_col(sub_df, default_old)
        add_by_key(
            sub_df,
            score_cols_map={score_col: out_col} if score_col else {},
            join_order=[{"on": "Email"}, {"on": "DNI"}, {"on": ["Nombre", "Apellido(s)"]}]
        )

    # ---- Filter/fill by Teacher Roster
    if roster_file is not None:
        roster = load_teacher_email_roster(roster_file)
        roster["Email"] = roster["Email"].apply(normalize_email)
        roster = roster.dropna(subset=["Email"]).drop_duplicates(subset=["Email"])
        if len(roster) > 0:
            valid_emails = set(roster["Email"].tolist())
            all_data = all_data[all_data["Email"].isin(valid_emails)]
            # only bring Email + names from roster; there are no email-ish extras here
            all_data = pd.merge(all_data, roster, on="Email", how="left")
            # Fill names from roster when missing
            def blank(x):
                return (pd.isna(x)) or (str(x).strip() == "")
            if "Nombre_roster" in all_data.columns:  # never created (we merged on distinct cols), but safe
                all_data["Nombre"] = np.where(all_data["Nombre"].apply(blank),
                                              all_data["Nombre_roster"], all_data["Nombre"])
                all_data["Apellido(s)"] = np.where(all_data["Apellido(s)"].apply(blank),
                                                   all_data["Apellido(s)_roster"], all_data["Apellido(s)"])
                all_data = all_data.drop(columns=[c for c in ["Nombre_roster", "Apellido(s)_roster"] if c in all_data.columns])

    # ---- Numeric components
    numeric_columns = [
        "induccion", "bus_biblioteca", "diseno_sesion",
        "Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica",
        "Padlet", "Nearpod", "Tareas_y_foros",
        "integracion", "rsu", "estress", "hab_comunicacion"
    ]
    for col in numeric_columns:
        if col not in all_data.columns:
            all_data[col] = 0
        all_data[col] = pd.to_numeric(all_data[col], errors="coerce").fillna(0)

    # ---- Metrics
    all_data["Average"] = all_data[numeric_columns].mean(axis=1).round(2)

    def calculate_percentage(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round(available / len(numeric_columns) * 100, 2) if len(numeric_columns) else 0.0

    all_data["Percentage"] = all_data.apply(calculate_percentage, axis=1)
    all_data["Marks_Out_Of_20"] = (all_data["Percentage"] / 5).round(2)

    # ---- Year filtering & de-dup by Email
    if "Year" not in all_data.columns:
        all_data["Year"] = pd.NA
    filtered = all_data[all_data["Year"].isin([2024, 2025])].copy()
    filtered = filtered[filtered[numeric_columns].sum(axis=1) > 0]
    if filtered.empty:
        return pd.DataFrame()

    filtered["YearPref"] = filtered["Year"].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=["Marks_Out_Of_20", "Average", "YearPref"],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=["Email"], keep="first").copy()
    highest["Highest_Score_Year"] = highest["Year"]

    # ---- Final columns
    final_columns = [
        "Periodo", "Highest_Score_Year", "Email", "DNI", "Nombre", "Apellido(s)",
        "induccion", "bus_biblioteca", "diseno_sesion",
        "Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica",
        "Padlet", "Nearpod", "Tareas_y_foros",
        "integracion", "rsu", "estress", "hab_comunicacion",
        "Average", "Marks_Out_Of_20", "Percentage"
    ]
    for c in final_columns:
        if c not in highest.columns:
            highest[c] = pd.NA
    return highest[final_columns]

# -----------------------------
# Streamlit App
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload **Master** Excel and the **Teacher Email Roster** (e.g., *C9_25-II_121025.xlsx*). "
        "Teachers are identified by **Email** and we keep the **highest Marks Out Of 20** across 2024 & 2025."
    )

    up_master = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    up_roster = st.file_uploader("Choose the Teacher Email Roster Excel file", type=["xlsx", "xls"], key="roster")

    if up_master is not None and up_roster is not None:
        try:
            with st.spinner("Processing your Excel files and comparing 2024 vs 2025 (by Email)…"):
                final_data = extract_data_from_excel(up_master, roster_file=up_roster)

            if final_data.empty:
                st.warning("No records with scores found for 2024 or 2025 (after filtering by the email roster).")
                return

            st.success(f"Done! Showing highest marks per teacher (unique by Email). Total: {len(final_data)}")

            st.subheader("Preview")
            st.dataframe(final_data.head(30))

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Teachers", len(final_data))
            with col2:
                st.metric("Avg Marks (Out of 20)", f"{final_data['Marks_Out_Of_20'].mean():.2f}")
            with col3:
                st.metric("Avg Percentage", f"{final_data['Percentage'].mean():.2f}%")

            st.subheader("Distribution by Highest Score Year")
            year_counts = final_data["Highest_Score_Year"].value_counts().sort_index()
            st.bar_chart(year_counts)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"Highest_Marks_2024_vs_2025_by_Email_{timestamp}.xlsx"
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_data.to_excel(writer, index=False, sheet_name="Highest Marks (Unique by Email)")
            output.seek(0)
            st.download_button(
                label="📥 Download (Unique by Email, Highest of 2024/2025)",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error while processing: {e}")
            st.info("If this repeats, ensure requirements include `openpyxl` and check Render logs.")
    else:
        st.info("👆 Please upload both the Master Excel and the Teacher Email Roster Excel to get started.")
        st.subheader("Expected Columns (quick reference)")
        st.markdown("""
**Master Excel** (fuzzy sheet detection works):
- **Inducción / nota Inducción**: `Periodo/PERIODO`, `DNI`, `Nombre`, `Apellido(s)`, `Dirección de correo` (to build Email), `Calificación` or `Total del curso (Real)`
- **Bus. biblioteca**: `DNI` and/or `Email`, `Promedio`
- **Diseño de sesión**: `Nombre`, `Apellido(s)`, `Promedio`
- **Comp. Tec**: Zoom básico/Avanzado, Grupos Moodle, Rúbrica, Padlet, Nearpod, Tareas y foros
- **Integración**: final score column (title varies)
- **RSU / estress / Hab. comunicación**: score column (title varies)

**Teacher Email Roster (e.g., C9_25-II_121025.xlsx)**:
- One email column (e.g., `Dirección de correo`, `Correo`, `Email`).
- Optional: names (`Nombres`, `Apellido Paterno`, `Apellido Materno`) or a single full-name column.
        """)

if __name__ == "__main__":
    main()
