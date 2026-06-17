from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import streamlit as st

import combinar_libro_arqueos as combinador


st.set_page_config(
    page_title="FBP Caja y Arqueos",
    page_icon="FBP",
    layout="wide",
)

st.markdown(
    """
    <style>
    :root {
        --teal: #087887;
        --teal-dark: #056170;
        --teal-soft: #DDF6F7;
        --teal-soft-2: #ECFAFA;
        --ink: #071833;
        --muted: #54657A;
        --line: #D8E3EA;
        --panel: #FFFFFF;
        --bg: #F4F7FA;
    }
    .stApp {
        background: var(--bg);
        color: var(--ink);
    }
    .block-container {
        padding-top: 0.6rem;
        padding-bottom: 1.6rem;
        max-width: 1260px;
    }
    .topbar {
        height: 48px;
        margin: -0.6rem calc(50% - 50vw) 1rem calc(50% - 50vw);
        padding: 0 24px;
        background: rgba(255,255,255,0.96);
        border-bottom: 1px solid rgba(8, 120, 135, 0.08);
        display: flex;
        align-items: center;
        gap: 18px;
        box-shadow: none;
    }
    .brand-mark {
        width: 32px;
        height: 32px;
        border-radius: 10px;
        background: var(--teal-soft);
        color: var(--teal);
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-weight: 900;
        border: 1px solid #C7EEF1;
    }
    .brand-name {
        font-size: 1.12rem;
        font-weight: 800;
        color: var(--teal-dark);
        margin-right: 16px;
    }
    .nav-chip {
        display: inline-flex;
        align-items: center;
        height: 34px;
        padding: 0 14px;
        border-radius: 10px;
        color: #56677C;
        font-size: 0.92rem;
        font-weight: 650;
    }
    .nav-chip.active {
        color: var(--teal-dark);
        background: #CFF2F5;
    }
    .hero-wrap {
        display: grid;
        grid-template-columns: minmax(0, 760px);
        gap: 0;
        align-items: stretch;
        margin-bottom: 0.55rem;
    }
    .hero-copy {
        padding: 0.35rem 0.1rem 0.15rem 0.1rem;
    }
    .result-pill {
        display: inline-flex;
        align-items: center;
        gap: 9px;
        padding: 9px 14px;
        border-radius: 999px;
        color: var(--teal-dark);
        background: #CFF6EC;
        font-size: 0.94rem;
        font-weight: 760;
        margin-bottom: 0.45rem;
    }
    .pill-dot {
        width: 22px;
        height: 22px;
        border-radius: 999px;
        background: #0AA88F;
        color: white;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 0.84rem;
        font-weight: 900;
    }
    .app-title {
        font-size: 2.15rem;
        line-height: 1.05;
        font-weight: 850;
        color: var(--ink);
        letter-spacing: 0;
        margin-bottom: 0.45rem;
    }
    .app-subtitle {
        color: var(--muted);
        font-size: 1.05rem;
        max-width: 690px;
        margin-bottom: 0.65rem;
    }
    .summary-panel {
        border: 1px solid var(--line);
        background: var(--panel);
        border-radius: 18px;
        padding: 1.05rem;
        box-shadow: 0 18px 36px rgba(20, 33, 52, 0.08);
    }
    .summary-inner {
        border: 1px solid #D8E3EA;
        border-radius: 14px;
        padding: 1.2rem 1.35rem;
        background: linear-gradient(180deg, #FFFFFF 0%, #FBFDFE 100%);
    }
    .summary-label {
        color: #718198;
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-weight: 800;
        margin-bottom: 0.45rem;
    }
    .summary-number {
        color: var(--teal-dark);
        font-size: 2.35rem;
        line-height: 1;
        font-weight: 850;
    }
    .upload-intro {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 18px;
        margin: 0.5rem 0 0.9rem 0;
    }
    .upload-title {
        color: var(--ink);
        font-size: 1.05rem;
        font-weight: 800;
        margin: 0;
    }
    .upload-hint {
        color: var(--muted);
        font-size: 0.9rem;
        margin: 0;
    }
    .upload-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        margin: 0 0 0.55rem 0;
    }
    .upload-label {
        color: var(--ink);
        font-size: 1rem;
        font-weight: 830;
    }
    .mini-pill {
        display: inline-flex;
        align-items: center;
        height: 25px;
        padding: 0 10px;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 800;
        border: 1px solid #D7E1E8;
        background: #F3F6F9;
        color: #6B7788;
    }
    .mini-pill.ready {
        background: #D6F6EE;
        color: var(--teal-dark);
        border-color: #B8ECE4;
    }
    .file-card {
        height: 176px;
        min-height: 176px;
        box-sizing: border-box;
        border: 1px solid var(--line);
        border-radius: 18px;
        background: linear-gradient(180deg, #FFFFFF 0%, #F7FBFC 100%);
        box-shadow: 0 18px 34px rgba(20, 33, 52, 0.07);
        padding: 1.15rem 1.25rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        position: relative;
        transition: border-color 150ms ease, box-shadow 150ms ease, transform 150ms ease;
    }
    .file-card:hover {
        border-color: #0F8B9A;
        box-shadow: 0 22px 40px rgba(8, 120, 135, 0.12);
        transform: translateY(-1px);
    }
    .file-icon {
        width: 42px;
        height: 42px;
        border-radius: 13px;
        background: var(--teal-soft);
        color: var(--teal-dark);
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-weight: 900;
        margin-right: 12px;
    }
    .file-main {
        display: flex;
        align-items: center;
        min-width: 0;
    }
    .file-name {
        color: var(--ink);
        font-weight: 830;
        overflow-wrap: anywhere;
    }
    .file-meta {
        color: var(--muted);
        font-size: 0.86rem;
        margin-top: 0.15rem;
    }
    .remove-cross {
        width: 34px;
        height: 34px;
        border-radius: 999px;
        background: #FDE8E8;
        color: #C81E1E;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-weight: 900;
        font-size: 1.12rem;
        border: 1px solid #F8B4B4;
        text-decoration: none !important;
        opacity: 0;
        transform: translateX(4px);
        transition: opacity 150ms ease, transform 150ms ease, background 150ms ease;
        flex: 0 0 auto;
    }
    .file-card:hover .remove-cross {
        opacity: 1;
        transform: translateX(0);
    }
    .remove-cross:hover {
        background: #FBD5D5;
        border-color: #F05252;
        color: #9B1C1C;
    }
    div[data-testid="stFileUploader"] {
        height: 176px;
        box-sizing: border-box;
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 0.8rem;
        box-shadow: 0 18px 34px rgba(20, 33, 52, 0.07);
    }
    div[data-testid="stFileUploader"] label p {
        color: var(--ink);
        font-weight: 800;
        font-size: 1rem;
    }
    div[data-testid="stFileUploader"] section {
        background: linear-gradient(180deg, #FFFFFF 0%, #F7FBFC 100%);
        border: 1.5px dashed #91C9D1;
        border-radius: 15px;
        height: 100%;
        min-height: 0;
        box-sizing: border-box;
        padding: 0.9rem;
    }
    div[data-testid="stFileUploaderDropzone"] {
        height: 100%;
        min-height: 0;
        align-items: center;
    }
    div[data-testid="stFileUploaderDropzone"] > div {
        min-height: 0;
        justify-content: center;
    }
    div[data-testid="stFileUploaderDropzone"] button {
        background: var(--teal-soft) !important;
        color: var(--teal-dark) !important;
        border: 0 !important;
        border-radius: 10px !important;
        font-weight: 800 !important;
    }
    div[data-testid="stDownloadButton"] button,
    div[data-testid="stButton"] button {
        border-radius: 10px;
        font-weight: 800;
        min-height: 2.9rem;
        border-color: #C7D7E0;
    }
    div[data-testid="stButton"]:has(button:not([kind="primary"])) {
        width: 34px;
        height: 0;
        margin-top: -106px;
        margin-left: calc(100% - 54px);
        position: relative;
        z-index: 20;
    }
    div[data-testid="stButton"] button:not([kind="primary"]) {
        width: 34px !important;
        min-width: 34px !important;
        height: 34px !important;
        min-height: 34px !important;
        padding: 0 !important;
        border-radius: 999px !important;
        opacity: 0 !important;
        cursor: pointer !important;
    }
    div[data-testid="stButton"] button[kind="primary"],
    div[data-testid="stDownloadButton"] button[kind="primary"] {
        background: var(--teal) !important;
        border-color: var(--teal) !important;
        color: #FFFFFF !important;
    }
    div[data-testid="stAlert"] {
        border-radius: 14px;
        border-color: #CFE4EA;
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid var(--line);
        border-radius: 18px;
        overflow: hidden;
        box-shadow: 0 18px 34px rgba(20, 33, 52, 0.07);
    }
    hr {
        margin: 1.5rem 0;
    }
    @media (max-width: 900px) {
        .hero-wrap { grid-template-columns: 1fr; }
        .app-title { font-size: 2rem; }
        .topbar { gap: 10px; padding: 0 14px; }
        .nav-chip { display: none; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)
MONEY_COLUMNS = [
    "Ingresos Tarjeta",
    "Ingresos Efectivo",
    "Salidas",
    "Retiradas Efectivo",
    "Saldo Real",
    "Saldo Teorico",
    "Descuadre",
]


def clear_combined_result() -> None:
    st.session_state.combined_df = None
    st.session_state.combined_xlsx = None


def reset_upload(kind: str) -> None:
    st.session_state[f"{kind}_upload_nonce"] += 1
    st.session_state[f"{kind}_file"] = None
    clear_combined_result()




def store_uploaded_file(kind: str, uploaded_file: Any) -> None:
    if uploaded_file is None:
        return
    st.session_state[f"{kind}_file"] = {
        "name": uploaded_file.name,
        "bytes": uploaded_file.getvalue(),
    }
    clear_combined_result()
    st.rerun()


def upload_header(label: str, stored_file: Optional[dict[str, Any]]) -> None:
    pill_class = "mini-pill ready" if stored_file else "mini-pill"
    pill_text = "Listo" if stored_file else "Pendiente"
    st.markdown(
        f"""
        <div class="upload-head">
            <div class="upload-label">{label}</div>
            <div class="{pill_class}">{pill_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def file_card(kind: str, stored_file: dict[str, Any]) -> None:
    size_kb = len(stored_file["bytes"]) / 1024
    extension = Path(stored_file["name"]).suffix.replace(".", "").upper() or "XLS"
    st.markdown(
        f"""
        <div class="file-card">
            <div class="file-main">
                <div class="file-icon">{extension}</div>
                <div>
                    <div class="file-name">{stored_file['name']}</div>
                    <div class="file-meta">{size_kb:.1f} KB · archivo cargado</div>
                </div>
            </div>
            <span class="remove-cross" title="Quitar archivo">×</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.button(
        "×",
        key=f"remove_{kind}",
        help="Quitar archivo",
        on_click=reset_upload,
        args=(kind,),
    )


def save_uploaded_file(uploaded_file: dict[str, Any], folder: Path, filename: str) -> Path:
    suffix = Path(uploaded_file["name"]).suffix.lower()
    path = folder / f"{filename}{suffix}"
    path.write_bytes(uploaded_file["bytes"])
    return path


def build_combined_workbook(libro_file: Any, arqueos_file: Any) -> tuple[pd.DataFrame, bytes]:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        libro_path = save_uploaded_file(libro_file, tmp_path, "libro_caja")
        arqueos_path = save_uploaded_file(arqueos_file, tmp_path, "arqueos")

        libro_rows = combinador.extract_libro(libro_path)
        arqueo_rows = combinador.extract_arqueos(arqueos_path)
        combinador.validate_same_dates(libro_rows, arqueo_rows)
        combined_rows = combinador.combine_rows(libro_rows, arqueo_rows)

        output_path = tmp_path / "resumen_final_caja_arqueos.xlsx"
        combinador.write_xlsx(combined_rows, output_path)

        df = pd.DataFrame(combined_rows, columns=combinador.OUTPUT_COLUMNS)
        return df, output_path.read_bytes()


def format_preview(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    styles = []
    for column in df.columns:
        if column == "Descuadre":
            styles.append(
                {
                    "selector": f"td.col{df.columns.get_loc(column)}",
                    "props": [("font-weight", "650")],
                }
            )
    return (
        df.style.format({column: "{:,.2f} €" for column in MONEY_COLUMNS})
        .map(lambda value: "color: #b42318; font-weight: 700" if value < 0 else "color: #027a48; font-weight: 700" if value > 0 else "", subset=["Descuadre"])
        .set_table_styles(styles)
    )


st.markdown(
    """
    <div class="topbar">
        <div class="brand-mark">FBP</div>
        <div class="brand-name">Farmacia Bafalluy</div>
        <div class="nav-chip">Caja</div>
        <div class="nav-chip active">Combinar</div>
    </div>
    <div class="hero-wrap">
        <div class="hero-copy">
            <div class="result-pill"><span class="pill-dot">✓</span> Preparado para combinar</div>
            <div class="app-title">Libro de caja y arqueos</div>
            <div class="app-subtitle">Sube los dos archivos del mismo periodo, valida que las fechas encajan y descarga un Excel final con formato listo para revisar.</div>
        </div>

    </div>

    """,
    unsafe_allow_html=True,
)

if "libro_upload_nonce" not in st.session_state:
    st.session_state.libro_upload_nonce = 0
if "arqueos_upload_nonce" not in st.session_state:
    st.session_state.arqueos_upload_nonce = 0
if "libro_file" not in st.session_state:
    st.session_state.libro_file = None
if "arqueos_file" not in st.session_state:
    st.session_state.arqueos_file = None
if "combined_df" not in st.session_state:
    st.session_state.combined_df = None
if "combined_xlsx" not in st.session_state:
    st.session_state.combined_xlsx = None


left, right = st.columns(2, gap="large")
with left:
    upload_header("Libro de caja", st.session_state.libro_file)
    if st.session_state.libro_file:
        file_card("libro", st.session_state.libro_file)
    else:
        uploaded_libro = st.file_uploader(
            "Libro de caja",
            type=["xlsx"],
            key=f"libro_file_uploader_{st.session_state.libro_upload_nonce}",
            help="Archivo exportado como libro_de_caja*.xlsx",
            label_visibility="collapsed",
        )
        store_uploaded_file("libro", uploaded_libro)
with right:
    upload_header("Arqueos", st.session_state.arqueos_file)
    if st.session_state.arqueos_file:
        file_card("arqueos", st.session_state.arqueos_file)
    else:
        uploaded_arqueos = st.file_uploader(
            "Arqueos",
            type=["xlsx", "csv"],
            key=f"arqueos_file_uploader_{st.session_state.arqueos_upload_nonce}",
            help="Archivo arqueos*.xlsx o arqueos*.csv",
            label_visibility="collapsed",
        )
        store_uploaded_file("arqueos", uploaded_arqueos)

st.divider()

libro_file = st.session_state.libro_file
arqueos_file = st.session_state.arqueos_file
combine_disabled = libro_file is None or arqueos_file is None
if st.button("Combinar", type="primary", disabled=combine_disabled, use_container_width=True):
    try:
        with st.spinner("Validando fechas y creando Excel final..."):
            df, xlsx_bytes = build_combined_workbook(libro_file, arqueos_file)
            st.session_state.combined_df = df
            st.session_state.combined_xlsx = xlsx_bytes
        st.success(f"Combinado correctamente: {len(df)} dias.")
    except Exception as exc:
        st.session_state.combined_df = None
        st.session_state.combined_xlsx = None
        st.error("No he podido combinar los archivos.")
        st.code(str(exc), language="text")

if combine_disabled:
    st.info("Sube los dos archivos para activar el boton de combinar.")

if st.session_state.combined_df is not None and st.session_state.combined_xlsx is not None:
    st.subheader("Vista previa")
    preview_df = st.session_state.combined_df.copy()
    st.dataframe(
        format_preview(preview_df),
        use_container_width=True,
        hide_index=True,
        height=420,
    )
    st.download_button(
        "Descargar Excel final",
        data=st.session_state.combined_xlsx,
        file_name="resumen_final_caja_arqueos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )












