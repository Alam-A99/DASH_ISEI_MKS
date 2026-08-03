"""
Dashboard Ketimpangan Kabupaten/Kota Sulawesi Selatan
================================================================
Jalankan dengan:
    streamlit run app.py

Dataset yang didukung: file .xlsx / .csv dengan struktur kolom seperti
OK_Panel_PDRB_KAB-Kota_Sulsel_14-25_2Inclusive_Growth.xlsx
(kabupaten, tahun, 9 kolom fungsi belanja pemerintah, Total,
17 kolom sektor lapangan usaha A-U, PDRB, ipm, miskin,
pertumbuhan, gini, unemploy, pop_ribujiwa, pdrb_perkap, pct_formal, pend_miskin)
"""

import io
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import statsmodels.api as sm
import streamlit as st
# ----------------------------------------------------------------------------
# KONFIGURASI HALAMAN
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Ketimpangan Sulsel",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# KONSTAN & MAPPING KOLOM
# ----------------------------------------------------------------------------
SECTOR_COLS = {
    "A_Pertanian_KP": "A. Pertanian, Kehutanan & Perikanan",
    "B_Pertambangan": "B. Pertambangan & Penggalian",
    "C_Industri_Pengolahan": "C. Industri Pengolahan",
    "D_Pengadaan_LG": "D. Pengadaan Listrik & Gas",
    "E_Pengadaan_Air_dll": "E. Pengadaan Air, Pengelolaan Sampah",
    "F_Konstruksi": "F. Konstruksi",
    "G_Perdagangan_Besar_dll": "G. Perdagangan Besar & Eceran",
    "H_Transportasi": "H. Transportasi & Pergudangan",
    "I_Penyediaan_AkMamin": "I. Penyediaan Akomodasi & Makan Minum",
    "J_Infokom": "J. Informasi & Komunikasi",
    "K_Jasa_Keuangan": "K. Jasa Keuangan & Asuransi",
    "L_Real_Estate": "L. Real Estate",
    "M_N_Jasa_Perusahaan": "M,N. Jasa Perusahaan",
    "O_Administrasi_Pem": "O. Administrasi Pemerintahan",
    "P_Jasa_Pendidikan": "P. Jasa Pendidikan",
    "Q_Jasa_KesSos": "Q. Jasa Kesehatan & Sosial",
    "R_S_T_U_Jasa_Lain": "R,S,T,U. Jasa Lainnya",
}

GOV_FUNC_COLS = {
    "Pelayanan_Umum": "Pelayanan Umum",
    "Ketertiban_aman": "Ketertiban & Keamanan",
    "Ekonomi": "Ekonomi",
    "Lingk_hidup": "Lingkungan Hidup",
    "Perum_Fasum": "Perumahan & Fasilitas Umum",
    "Kesehatan": "Kesehatan",
    "Pariwisata": "Pariwisata",
    "Pendidikan": "Pendidikan",
    "Perlindungan_Sosial": "Perlindungan Sosial",
}

SOCIAL_COLS = {
    "ipm": "Indeks Pembangunan Manusia (IPM)",
    "miskin": "Tingkat Kemiskinan (%)",
    "pertumbuhan": "Pertumbuhan Ekonomi (%)",
    "gini": "Rasio Gini (Ketimpangan)",
    "unemploy": "Tingkat Pengangguran Terbuka (%)",
    "pop_ribujiwa": "Jumlah Penduduk (ribu jiwa)",
    "pdrb_perkap": "PDRB per Kapita",
    "pct_formal": "% Pekerja Sektor Formal",
    "pend_miskin": "Rata-rata Lama Pendidikan Penduduk Miskin",
}

REQUIRED_BASE_COLS = ["kabupaten", "tahun", "PDRB", "gini"]

# Variabel yang boleh dipilih sebagai Variabel Dependen (Y) pada tab Model Regresi Panel.
DEPENDENT_VAR_OPTIONS = {
    "pertumbuhan": "Pertumbuhan Ekonomi (%)",
    "gini": "Rasio Gini (Ketimpangan)",
    "unemploy": "Tingkat Pengangguran Terbuka (%)",
    "pend_miskin": "Rata-rata Lama Pendidikan Penduduk Miskin",
    "miskin": "Tingkat Kemiskinan (%)",
}

DEFAULT_GITHUB_URL = "https://github.com/Alam-A99/DASH_ISEI_MKS/blob/main/OK_Panel_PDRB_KAB-Kota_Sulsel_14-25.xlsx"

# ----------------------------------------------------------------------------
# SAMPLE DATA (fallback)
# ----------------------------------------------------------------------------
SAMPLE_CSV = """kabupaten;tahun;Pelayanan_Umum;Ketertiban_aman;Ekonomi;Lingk_hidup;Perum_Fasum;Kesehatan;Pariwisata;Pendidikan;Perlindungan_Sosial;Total;A_Pertanian_KP;B_Pertambangan;C_Industri_Pengolahan;D_Pengadaan_LG;E_Pengadaan_Air_dll;F_Konstruksi;G_Perdagangan_Besar_dll;H_Transportasi;I_Penyediaan_AkMamin;J_Infokom;K_Jasa_Keuangan;L_Real_Estate;M_N_Jasa_Perusahaan;O_Administrasi_Pem;P_Jasa_Pendidikan;Q_Jasa_KesSos;R_S_T_U_Jasa_Lain;PDRB;ipm;miskin;pertumbuhan;gini;unemploy;pop_ribujiwa;pdrb_perkap;pct_formal;pend_miskin
Bantaeng;2014;142,50;38,20;215,30;45,60;67,80;98,40;22,10;185,40;54,70;870,00;1302,07;90,38;182,07;4,90;3,33;576,72;564,15;45,59;29,75;103,32;88,11;208,92;5,02;248,55;222,02;90,52;53,87;3819,28;65,77;9,68;8,33;0,46;2,40;182,28;27,23;45,44;17,70
Bantaeng;2015;155,80;41,20;238,60;49,10;72,30;108,50;25,80;198,70;60,20;950,20;1307,02;116,97;196,86;5,40;3,37;600,44;657,08;53,95;33,15;113,31;93,16;231,45;5,34;258,17;232,02;105,95;59,44;4073,06;66,20;9,53;6,64;0,44;4,07;183,24;30,42;40,83;17,55
Bantaeng;2016;168,30;45,10;258,90;52,40;78,60;118,20;29,40;215,30;65,80;1032,00;1411,06;130,11;205,94;5,97;3,79;643,74;693,87;54,91;34,43;126,99;105,68;240,25;5,64;290,87;253,00;106,51;61,46;4374,21;66,59;9,51;7,39;0,38;4,07;184,46;34,13;36,23;17,53
Bantaeng;2017;178,90;48,30;278,50;56,80;85,20;128,40;33,60;230,50;71,90;1112,10;1489,99;143,98;218,79;6,30;4,03;713,15;761,09;56,19;38,36;137,29;109,36;247,45;6,00;306,93;278,46;110,11;66,69;4694,16;67,27;9,66;7,31;0,42;5,23;185,52;37,41;34,56;17,91
Barru;2014;138,40;36,50;205,80;43,90;65,20;95,60;21,50;180,20;52,90;840,00;1004,39;60,65;145,10;3,13;2,96;367,40;216,09;56,35;26,10;99,88;58,93;90,23;0,80;232,96;109,73;69,10;16,56;2560,34;64,94;10,68;6,06;0,38;8,94;165,98;15,38;58,13;17,70
Barru;2015;150,20;39,80;226,40;47,60;70,50;105,80;24,30;195,30;58,10;918,00;1082,53;66,84;157,14;3,42;3,42;394,07;234,68;61,16;27,25;116,55;67,94;97,34;0,82;246,28;117,19;74,10;17,78;2768,52;65,73;9,59;8,13;0,38;5,75;167,51;17,40;33,05;16,10
"""

# ----------------------------------------------------------------------------
# FUNGSI UTILITAS
# ----------------------------------------------------------------------------
def _to_numeric_id(series: pd.Series) -> pd.Series:
    """Konversi string format Indonesia (koma desimal, titik ribuan) -> float."""
    if series.dtype == object:
        cleaned = (
            series.astype(str)
            .str.strip()
            .str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)
            .str.replace(",", ".", regex=False)
        )
        return pd.to_numeric(cleaned, errors="coerce")
    return pd.to_numeric(series, errors="coerce")


def normalize_github_url(url: str) -> str:
    url = url.strip()
    m = re.match(r"https?://github\.com/([^/]+)/([^/]+)/blob/(.+)", url)
    if m:
        user, repo, rest = m.groups()
        return f"https://raw.githubusercontent.com/{user}/{repo}/{rest}"
    return url


@st.cache_data(show_spinner=False)
def load_data_from_url(url: str) -> pd.DataFrame:
    url = normalize_github_url(url)
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    content = resp.content
    if url.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(content))
    else:
        text = content.decode("utf-8", errors="ignore")
        first_line = text.split("\n")[0]
        sep = ";" if first_line.count(";") > first_line.count(",") else ","
        df = pd.read_csv(io.StringIO(text), sep=sep)
    return _postprocess_df(df)


@st.cache_data(show_spinner=False)
def load_data(file) -> pd.DataFrame:
    if file is None:
        df = pd.read_csv(io.StringIO(SAMPLE_CSV), sep=";")
    else:
        name = file.name.lower()
        if name.endswith(".csv"):
            raw = file.read()
            file.seek(0)
            sep = ";" if raw.decode("utf-8", errors="ignore").split("\n")[0].count(";") > raw.decode("utf-8", errors="ignore").split("\n")[0].count(",") else ","
            df = pd.read_csv(file, sep=sep)
        else:
            df = pd.read_excel(file)
    return _postprocess_df(df)


def _postprocess_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip() for c in df.columns]
    numeric_candidates = (
        list(SECTOR_COLS.keys()) +
        list(GOV_FUNC_COLS.keys()) +
        ["Total"] +
        list(SOCIAL_COLS.keys()) +
        ["PDRB"]
    )
    for col in numeric_candidates:
        if col in df.columns:
            df[col] = _to_numeric_id(df[col])

    if "tahun" in df.columns:
        df["tahun"] = pd.to_numeric(df["tahun"], errors="coerce").astype("Int64")
    if "kabupaten" in df.columns:
        df["kabupaten"] = df["kabupaten"].astype(str).str.strip()

    df = df.dropna(subset=["kabupaten", "tahun"]) if {"kabupaten", "tahun"}.issubset(df.columns) else df
    return df


def available_sector_cols(df: pd.DataFrame) -> list:
    return [c for c in SECTOR_COLS if c in df.columns]


def available_gov_cols(df: pd.DataFrame) -> list:
    return [c for c in GOV_FUNC_COLS if c in df.columns]


def fmt_number(x, dec=2):
    if pd.isna(x):
        return "-"
    return f"{x:,.{dec}f}".replace(",", "#").replace(".", ",").replace("#", ".")


def kpi_delta(current, previous):
    if previous in (None, 0) or pd.isna(previous) or pd.isna(current):
        return None
    return (current - previous) / previous * 100


# ----------------------------------------------------------------------------
# SIDEBAR — UPLOAD & FILTER
# ----------------------------------------------------------------------------
st.sidebar.title("📊 Panel Kontrol")
st.sidebar.markdown("Upload dataset panel (.xlsx/.csv).")

sumber_data = st.sidebar.radio("Sumber data", ["Upload file", "Link Dataset"], horizontal=True)

df_raw = None
load_error = None

if sumber_data == "Upload file":
    uploaded_file = st.sidebar.file_uploader("Dataset", type=["xlsx", "xls", "csv"])
    if uploaded_file is None:
        st.sidebar.info("Belum ada file diunggah — menampilkan **data contoh** (Bantaeng & Barru). "
                         "Unggah file lengkap Anda untuk analisis penuh seluruh kabupaten/kota.")
    try:
        df_raw = load_data(uploaded_file)
    except Exception as e:
        load_error = str(e)
else:
    st.sidebar.caption(
        "Paste link **Dataset** (mis. `https://raw.githubusercontent.com/user/repo/main/data.xlsx`) "
        "atau link biasa `https://github.com/user/repo/blob/main/data.xlsx."
    )
    github_url = st.sidebar.text_input(
        "Link dataset",
        value=DEFAULT_GITHUB_URL,
        placeholder="https://github.com/user/repo/blob/main/data.xlsx",
    )
    muat_btn = st.sidebar.button("🔄 Tampilkan Data", use_container_width=True)

    if "github_df" not in st.session_state:
        st.session_state["github_df"] = None

    if muat_btn and github_url:
        with st.spinner("Loading dataset..."):
            try:
                st.session_state["github_df"] = load_data_from_url(github_url)
                st.sidebar.success("Dataset berhasil dimuat.")
            except Exception as e:
                st.session_state["github_df"] = None
                load_error = f"Gagal memuat dari link tersebut: {e}"

    if st.session_state.get("github_df") is not None:
        df_raw = st.session_state["github_df"]
    else:
        st.sidebar.info("Belum ada data dimuat — menampilkan **data contoh** (Bantaeng & Barru).")
        df_raw = load_data(None)

if load_error:
    st.sidebar.error(load_error)

if df_raw.empty or not {"kabupaten", "tahun"}.issubset(df_raw.columns):
    st.error("Dataset tidak valid. Pastikan minimal terdapat kolom 'kabupaten' dan 'tahun'.")
    st.stop()

sector_cols = available_sector_cols(df_raw)
gov_cols = available_gov_cols(df_raw)

st.sidebar.markdown("---")
kab_list = sorted(df_raw["kabupaten"].dropna().unique().tolist())
selected_kabs = st.sidebar.multiselect("Kabupaten/Kota", kab_list, default=kab_list)

tahun_min, tahun_max = int(df_raw["tahun"].min()), int(df_raw["tahun"].max())
selected_years = st.sidebar.slider("Rentang Tahun", tahun_min, tahun_max, (tahun_min, tahun_max))

st.sidebar.markdown("---")
st.sidebar.caption("Dashboard Ketimpangan · Data Panel Kab/Kota Sulsel")

mask = (
    df_raw["kabupaten"].isin(selected_kabs)
    & df_raw["tahun"].between(selected_years[0], selected_years[1])
)
df = df_raw.loc[mask].copy()

if df.empty:
    st.warning("Tidak ada data pada kombinasi filter yang dipilih.")
    st.stop()

# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
st.title("⚖️ Dashboard Ketimpangan Ekonomi")
st.markdown(
    "Analisis panel PDRB sektoral, pengeluaran pemerintah (fungsi), dan **indikator ketimpangan** "
    "beserta faktor-faktor pendorongnya di kabupaten/kota Sulawesi Selatan."
)

tab_overview, tab_sektoral, tab_gov, tab_sosial, tab_perbandingan, tab_regresi = st.tabs(
    ["🏠 Ringkasan", "🏭 Kontribusi Sektoral", "🏛️ Belanja Pemerintah",
     "👥 Indikator Sosial", "⚖️ Perbandingan Wilayah", "📈 Model Regresi Panel"]
)

# ----------------------------------------------------------------------------
# TAB 1 — RINGKASAN
# ----------------------------------------------------------------------------
with tab_overview:
    latest_year = df["tahun"].max()
    df_latest = df[df["tahun"] == latest_year]
    df_prev = df[df["tahun"] == latest_year - 1] if (latest_year - 1) in df["tahun"].values else pd.DataFrame()

    c1, c2, c3, c4, c5 = st.columns(5)
    avg_pdrb = df_latest["PDRB"].mean()
    avg_pdrb_prev = df_prev["PDRB"].mean() if not df_prev.empty else None
    c1.metric(f"Rata-rata PDRB ({latest_year})", fmt_number(avg_pdrb, 1),
              f"{kpi_delta(avg_pdrb, avg_pdrb_prev):.1f}%" if kpi_delta(avg_pdrb, avg_pdrb_prev) is not None else None)

    if "ipm" in df.columns:
        avg_ipm = df_latest["ipm"].mean()
        avg_ipm_prev = df_prev["ipm"].mean() if not df_prev.empty else None
        c2.metric("Rata-rata IPM", fmt_number(avg_ipm, 2),
                  f"{kpi_delta(avg_ipm, avg_ipm_prev):.2f}%" if kpi_delta(avg_ipm, avg_ipm_prev) is not None else None)

    if "miskin" in df.columns:
        avg_miskin = df_latest["miskin"].mean()
        avg_miskin_prev = df_prev["miskin"].mean() if not df_prev.empty else None
        c3.metric("Rata-rata Kemiskinan", f"{fmt_number(avg_miskin, 2)}%",
                  f"{kpi_delta(avg_miskin, avg_miskin_prev):.2f}%" if kpi_delta(avg_miskin, avg_miskin_prev) is not None else None,
                  delta_color="inverse")

    if "gini" in df.columns:
        avg_gini = df_latest["gini"].mean()
        avg_gini_prev = df_prev["gini"].mean() if not df_prev.empty else None
        c4.metric("Rata-rata Rasio Gini", fmt_number(avg_gini, 3),
                  f"{kpi_delta(avg_gini, avg_gini_prev):.2f}%" if kpi_delta(avg_gini, avg_gini_prev) is not None else None,
                  delta_color="inverse")

    if "unemploy" in df.columns:
        avg_unemp = df_latest["unemploy"].mean()
        avg_unemp_prev = df_prev["unemploy"].mean() if not df_prev.empty else None
        c5.metric("Tingkat Pengangguran", f"{fmt_number(avg_unemp, 2)}%",
                  f"{kpi_delta(avg_unemp, avg_unemp_prev):.2f}%" if kpi_delta(avg_unemp, avg_unemp_prev) is not None else None,
                  delta_color="inverse")

    st.markdown("---")

    col_a, col_b = st.columns([3, 2])
    with col_a:
        st.subheader("Tren PDRB & Rasio Gini")
        agg = df.groupby("tahun").agg(PDRB=("PDRB", "mean"), gini=("gini", "mean")).reset_index()
        fig = go.Figure()
        fig.add_trace(go.Bar(x=agg["tahun"], y=agg["PDRB"], name="Rata-rata PDRB", marker_color="#4c72b0", opacity=0.7, yaxis="y1"))
        fig.add_trace(go.Scatter(x=agg["tahun"], y=agg["gini"], name="Rata-rata Rasio Gini",
                                  mode="lines+markers", line=dict(color="#dd4b39", width=3), yaxis="y2"))
        fig.update_layout(
            yaxis=dict(title="PDRB (Miliar Rp)"),
            yaxis2=dict(title="Rasio Gini", overlaying="y", side="right"),
            legend=dict(orientation="h", y=1.12),
            margin=dict(t=30, b=10), height=420,
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        st.subheader(f"Distribusi Rasio Gini ({latest_year})")
        if "gini" in df.columns:
            df_gini_dist = df_latest[["kabupaten", "gini"]].sort_values("gini", ascending=True)
            fig_bar_gini = px.bar(
                df_gini_dist, x="gini", y="kabupaten", orientation="h",
                title=f"Urutan Ketimpangan (Gini) — Tahun {latest_year}",
                labels={"gini": "Rasio Gini", "kabupaten": ""},
                color="gini", color_continuous_scale="RdYlGn_r",
            )
            fig_bar_gini.update_traces(texttemplate="%{x:.3f}", textposition="outside")
            fig_bar_gini.update_layout(yaxis=dict(categoryorder="total ascending"), 
                                       height=max(350, 28 * len(df_gini_dist)), margin=dict(t=20))
            st.plotly_chart(fig_bar_gini, use_container_width=True)

    st.markdown("---")
    st.subheader(f"Peringkat Kabupaten/Kota berdasarkan Ketimpangan Tertinggi ({latest_year})")
    if "gini" in df.columns:
        rank_df = df_latest[["kabupaten", "gini", "PDRB"]].sort_values("gini", ascending=False).reset_index(drop=True)
        rank_df.index += 1
        fig_rank = px.bar(
            rank_df, x="gini", y="kabupaten", orientation="h",
            color="gini", color_continuous_scale="RdYlGn_r",
            text="gini",
            labels={"gini": "Rasio Gini", "kabupaten": ""},
        )
        fig_rank.update_traces(texttemplate="%{text:.3f}", textposition="outside")
        fig_rank.update_layout(yaxis=dict(categoryorder="total ascending"), 
                               height=max(350, 28 * len(rank_df)), margin=dict(t=20))
        st.plotly_chart(fig_rank, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 2 — KONTRIBUSI SEKTORAL (fokus utama)
# ----------------------------------------------------------------------------
with tab_sektoral:
    st.subheader("Struktur & Kontribusi Sektor Lapangan Usaha")

    focus_kab = st.selectbox("Pilih kabupaten/kota untuk analisis struktur sektoral", options=["(Semua - rata-rata)"] + kab_list)

    if focus_kab == "(Semua - rata-rata)":
        df_focus = df.groupby("tahun")[sector_cols].mean().reset_index()
    else:
        df_focus = df[df["kabupaten"] == focus_kab][["tahun"] + sector_cols].sort_values("tahun")

    df_share = df_focus.copy()
    df_share[sector_cols] = df_share[sector_cols].div(df_share[sector_cols].sum(axis=1), axis=0) * 100
    df_share_long = df_share.melt(id_vars="tahun", value_vars=sector_cols, var_name="Sektor", value_name="Kontribusi (%)")
    df_share_long["Sektor"] = df_share_long["Sektor"].map(SECTOR_COLS)

    fig_area = px.area(
        df_share_long, x="tahun", y="Kontribusi (%)", color="Sektor",
        title=f"Komposisi Sektoral PDRB — {focus_kab}",
    )
    fig_area.update_layout(height=480, legend=dict(font=dict(size=10)))
    st.plotly_chart(fig_area, use_container_width=True)

    latest_share = df_share[df_share["tahun"] == df_share["tahun"].max()][sector_cols].T
    latest_share.columns = ["Kontribusi (%)"]
    latest_share["Sektor"] = [SECTOR_COLS[i] for i in latest_share.index]
    latest_share = latest_share.sort_values("Kontribusi (%)", ascending=True)
    fig_bar_share = px.bar(latest_share, x="Kontribusi (%)", y="Sektor", orientation="h",
                            title=f"Kontribusi Sektor terhadap PDRB Tahun {int(df_share['tahun'].max())}",
                            text="Kontribusi (%)", color="Kontribusi (%)", color_continuous_scale="Blues")
    fig_bar_share.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig_bar_share.update_layout(height=550, margin=dict(t=40))
    st.plotly_chart(fig_bar_share, use_container_width=True)

    st.markdown("---")
    st.subheader("🔎 Sektor Mana yang Paling Berasosiasi dengan Ketimpangan (Gini)?")
    st.caption(
        "Korelasi dihitung antara **pertumbuhan tahunan (%) tiap sektor** dengan **Rasio Gini**, "
        "menggunakan seluruh observasi panel (kabupaten × tahun) sesuai filter aktif."
    )

    growth_df = df.sort_values(["kabupaten", "tahun"]).copy()
    for col in sector_cols:
        growth_df[f"g_{col}"] = growth_df.groupby("kabupaten")[col].pct_change() * 100

    corr_rows = []
    for col in sector_cols:
        gcol = f"g_{col}"
        valid = growth_df[[gcol, "gini"]].dropna()
        corr = valid[gcol].corr(valid["gini"]) if len(valid) >= 3 else np.nan
        corr_rows.append({"Sektor": SECTOR_COLS[col], "Korelasi dengan Gini": corr})

    corr_df = pd.DataFrame(corr_rows).dropna().sort_values("Korelasi dengan Gini")
    if not corr_df.empty:
        fig_corr = px.bar(
            corr_df, x="Korelasi dengan Gini", y="Sektor", orientation="h",
            color="Korelasi dengan Gini", color_continuous_scale="RdYlGn",
            range_color=[-1, 1],
        )
        fig_corr.update_layout(height=550, margin=dict(t=20))
        st.plotly_chart(fig_corr, use_container_width=True)

        top_pos = corr_df.iloc[-1]
        top_neg = corr_df.iloc[0]
        colx, coly = st.columns(2)
        colx.success(f"**Korelasi positif terkuat:** {top_pos['Sektor']} (r = {top_pos['Korelasi dengan Gini']:.2f})")
        coly.error(f"**Korelasi negatif terkuat:** {top_neg['Sektor']} (r = {top_neg['Korelasi dengan Gini']:.2f})")
    else:
        st.info("Data belum cukup (minimal 2 tahun berurutan per kabupaten) untuk menghitung korelasi pertumbuhan sektoral.")

    st.markdown("---")
    st.subheader("Peta Korelasi Antar Sektor & Ketimpangan (Gini)")
    corr_matrix_cols = sector_cols + ["gini"]
    corr_matrix = df[corr_matrix_cols].corr()
    corr_matrix.index = [SECTOR_COLS.get(i, i) for i in corr_matrix.index]
    corr_matrix.columns = [SECTOR_COLS.get(c, c) for c in corr_matrix.columns]
    fig_heat = px.imshow(corr_matrix, color_continuous_scale="RdBu_r", zmin=-1, zmax=1, aspect="auto",
                          text_auto=".2f")
    fig_heat.update_layout(height=650, margin=dict(t=20))
    st.plotly_chart(fig_heat, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 3 — BELANJA PEMERINTAH (BARU)
# ----------------------------------------------------------------------------
with tab_gov:
    st.subheader("🏛️ Komposisi & Alokasi Belanja Pemerintah per Fungsi")
    st.caption("Menganalisis alokasi anggaran pemerintah daerah berdasarkan 9 fungsi utama.")

    if not gov_cols:
        st.warning("Dataset ini tidak memiliki kolom fungsi pengeluaran pemerintah "
                    "(Pelayanan_Umum, Ekonomi, Kesehatan, Pendidikan, dst.).")
    else:
        focus_kab_gov = st.selectbox(
            "Pilih kabupaten/kota untuk analisis belanja",
            options=["(Semua - rata-rata)"] + kab_list,
            key="focus_kab_gov",
        )

        if focus_kab_gov == "(Semua - rata-rata)":
            df_gov_focus = df.groupby("tahun")[gov_cols + (["Total"] if "Total" in df.columns else [])].mean().reset_index()
        else:
            df_gov_focus = df[df["kabupaten"] == focus_kab_gov][
                ["tahun"] + gov_cols + (["Total"] if "Total" in df.columns else [])
            ].sort_values("tahun")

        # Komposisi persentase
        df_gov_share = df_gov_focus.copy()
        base_sum = df_gov_focus[gov_cols].sum(axis=1)
        df_gov_share[gov_cols] = df_gov_focus[gov_cols].div(base_sum.replace(0, np.nan), axis=0) * 100

        df_gov_long = df_gov_share.melt(id_vars="tahun", value_vars=gov_cols, var_name="Fungsi", value_name="Proporsi (%)")
        df_gov_long["Fungsi"] = df_gov_long["Fungsi"].map(GOV_FUNC_COLS)

        fig_gov_area = px.area(
            df_gov_long, x="tahun", y="Proporsi (%)", color="Fungsi",
            title=f"Komposisi Belanja Pemerintah — {focus_kab_gov}",
        )
        fig_gov_area.update_layout(height=480, legend=dict(font=dict(size=10)))
        st.plotly_chart(fig_gov_area, use_container_width=True)

        # Bar horizontal tahun terakhir
        last_gov = df_gov_share[df_gov_share["tahun"] == df_gov_share["tahun"].max()][gov_cols].T
        last_gov.columns = ["Proporsi (%)"]
        last_gov["Fungsi"] = [GOV_FUNC_COLS[i] for i in last_gov.index]
        last_gov = last_gov.sort_values("Proporsi (%)", ascending=True)
        fig_gov_bar = px.bar(
            last_gov, x="Proporsi (%)", y="Fungsi", orientation="h",
            title=f"Alokasi Belanja Tahun {int(df_gov_share['tahun'].max())} (%)",
            text="Proporsi (%)", color="Proporsi (%)", color_continuous_scale="Purples",
        )
        fig_gov_bar.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig_gov_bar.update_layout(height=500, margin=dict(t=40))
        st.plotly_chart(fig_gov_bar, use_container_width=True)

        st.markdown("---")
        st.subheader("🔍 Fungsi Belanja yang Berasosiasi dengan Ketimpangan (Gini)")
        st.caption("Korelasi antara **pertumbuhan tahunan (%) tiap fungsi belanja** dan **Rasio Gini**.")

        g_gov = df.sort_values(["kabupaten", "tahun"]).copy()
        for col in gov_cols:
            g_gov[f"g_{col}"] = g_gov.groupby("kabupaten")[col].pct_change() * 100

        corr_gov_rows = []
        for col in gov_cols:
            gcol = f"g_{col}"
            valid = g_gov[[gcol, "gini"]].dropna()
            corr = valid[gcol].corr(valid["gini"]) if len(valid) >= 3 else np.nan
            corr_gov_rows.append({"Fungsi": GOV_FUNC_COLS[col], "Korelasi dengan Gini": corr})

        corr_gov_df = pd.DataFrame(corr_gov_rows).dropna().sort_values("Korelasi dengan Gini")
        if not corr_gov_df.empty:
            fig_gov_corr = px.bar(
                corr_gov_df, x="Korelasi dengan Gini", y="Fungsi", orientation="h",
                color="Korelasi dengan Gini", color_continuous_scale="RdYlGn", range_color=[-1, 1],
            )
            fig_gov_corr.update_layout(height=450, margin=dict(t=20))
            st.plotly_chart(fig_gov_corr, use_container_width=True)

            tp = corr_gov_df.iloc[-1]
            tn = corr_gov_df.iloc[0]
            gx, gy = st.columns(2)
            gx.success(f"**Korelasi positif terkuat:** {tp['Fungsi']} (r = {tp['Korelasi dengan Gini']:.2f})")
            gy.error(f"**Korelasi negatif terkuat:** {tn['Fungsi']} (r = {tn['Korelasi dengan Gini']:.2f})")
        else:
            st.info("Belum cukup data berurutan untuk menghitung korelasi pertumbuhan belanja.")

        # Perbandingan total belanja antar kabupaten (tahun terakhir)
        st.markdown("---")
        st.subheader(f"Perbandingan Total Belanja Pemerintah ({int(df['tahun'].max())})")
        if "Total" in df.columns:
            df_latest_gov = df[df["tahun"] == df["tahun"].max()].sort_values("Total", ascending=True)
            fig_total = px.bar(
                df_latest_gov, x="Total", y="kabupaten", orientation="h",
                color="Total", color_continuous_scale="Purples",
                text="Total",
                labels={"Total": "Total Belanja (Miliar Rp)", "kabupaten": ""},
            )
            fig_total.update_traces(texttemplate="%{text:,.1f}", textposition="outside")
            fig_total.update_layout(height=max(350, 28 * len(df_latest_gov)), margin=dict(t=20))
            st.plotly_chart(fig_total, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 4 — INDIKATOR SOSIAL
# ----------------------------------------------------------------------------
with tab_sosial:
    st.subheader("Perkembangan Indikator Sosial-Ekonomi")
    social_available = [c for c in SOCIAL_COLS if c in df.columns]
    metric_choice = st.selectbox("Pilih indikator", social_available, format_func=lambda c: SOCIAL_COLS[c])

    fig_line = px.line(
        df.sort_values("tahun"), x="tahun", y=metric_choice, color="kabupaten", markers=True,
        labels={metric_choice: SOCIAL_COLS[metric_choice], "tahun": "Tahun"},
        title=f"Tren {SOCIAL_COLS[metric_choice]} per Kabupaten/Kota",
    )
    fig_line.update_layout(height=480)
    st.plotly_chart(fig_line, use_container_width=True)

    st.markdown("---")
    colp, colq = st.columns(2)
    with colp:
        if {"gini", "ipm"}.issubset(df.columns):
            fig_sc = px.scatter(
                df, x="gini", y="ipm", size="PDRB", hover_name="kabupaten", hover_data=["tahun"],
                title="Rasio Gini vs IPM (ukuran gelembung = PDRB)",
                labels={"gini": "Rasio Gini (Ketimpangan)", "ipm": "IPM"},
            )
            fig_sc.update_layout(height=450)
            st.plotly_chart(fig_sc, use_container_width=True)
    with colq:
        if {"miskin", "unemploy"}.issubset(df.columns):
            fig_sc2 = px.scatter(
                df, x="unemploy", y="miskin", size="pop_ribujiwa" if "pop_ribujiwa" in df.columns else None,
                hover_name="kabupaten", hover_data=["tahun"],
                title="Pengangguran vs Kemiskinan (ukuran gelembung = populasi)",
                labels={"unemploy": "Tingkat Pengangguran (%)", "miskin": "Tingkat Kemiskinan (%)"},
            )
            fig_sc2.update_layout(height=450)
            st.plotly_chart(fig_sc2, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 5 — PERBANDINGAN WILAYAH
# ----------------------------------------------------------------------------
with tab_perbandingan:
    st.subheader("Perbandingan Antar Kabupaten/Kota")

    pivot_metric = st.selectbox(
        "Indikator untuk dibandingkan",
        ["PDRB", "gini"] + social_available if 'social_available' in dir() else ["PDRB", "gini"],
        format_func=lambda c: SOCIAL_COLS.get(c, c),
    )

    pivot_df = df.pivot_table(index="kabupaten", columns="tahun", values=pivot_metric, aggfunc="mean")
    fig_hm = px.imshow(pivot_df, aspect="auto", 
                        color_continuous_scale="Viridis" if pivot_metric != "gini" else "RdYlGn_r",
                        labels=dict(color=SOCIAL_COLS.get(pivot_metric, pivot_metric)),
                        title=f"Heatmap {SOCIAL_COLS.get(pivot_metric, pivot_metric)} — Kabupaten × Tahun")
    fig_hm.update_layout(height=max(400, 25 * len(pivot_df)), margin=dict(t=40))
    st.plotly_chart(fig_hm, use_container_width=True)

    st.markdown("---")
    st.subheader("PDRB vs Rasio Gini (seluruh observasi)")
    if "gini" in df.columns:
        fig_scatter_all = px.scatter(
            df, x="PDRB", y="gini", color="kabupaten",
            hover_name="kabupaten", hover_data=["tahun"], trendline="ols" if len(df) > 5 else None,
            labels={"gini": "Rasio Gini (Ketimpangan)"},
        )
        fig_scatter_all.update_layout(height=500)
        st.plotly_chart(fig_scatter_all, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 6 — MODEL REGRESI PANEL
# ----------------------------------------------------------------------------
with tab_regresi:
    st.subheader("📈 Model Regresi Data Panel")
    st.markdown(
        "Pilih **Variabel Dependen (Y)** dari daftar berikut: *Pertumbuhan Ekonomi, Rasio Gini (Ketimpangan), "
        "Tingkat Pengangguran, Rata-rata Lama Pendidikan Penduduk Miskin,* atau *Tingkat Kemiskinan*. "
        "Seluruh variabel lain (sektoral & fungsi belanja) otomatis menjadi kandidat **Variabel Independen (X)**."
    )

    dep_options_available = {k: v for k, v in DEPENDENT_VAR_OPTIONS.items() if k in df.columns}

    if not dep_options_available:
        st.warning("Tidak ada variabel dependen yang tersedia pada dataset ini.")
    else:
        col_sel1, col_sel2 = st.columns([2, 1])
        with col_sel1:
            dep_var = st.selectbox(
                "Variabel Dependen (Y)",
                list(dep_options_available.keys()),
                format_func=lambda c: dep_options_available[c],
            )
        with col_sel2:
            model_type = st.selectbox(
                "Jenis Model Panel",
                [
                    "Pooled OLS",
                    "Fixed Effect - Entity (Kabupaten/Kota)",
                    "Fixed Effect - Two Way (Kabupaten/Kota & Tahun)",
                ],
            )

        label_map = {**SECTOR_COLS, **GOV_FUNC_COLS, **SOCIAL_COLS, "PDRB": "PDRB", "Total": "Total Belanja"}
        
        # Variabel Independen: 17 sektor + 9 fungsi belanja pemerintah
        all_numeric_candidates = sector_cols + gov_cols
        indep_candidates = [c for c in all_numeric_candidates if c != dep_var]

        selected_indep = st.multiselect(
            "Variabel Independen (X)",
            indep_candidates,
            default=indep_candidates,
            format_func=lambda c: label_map.get(c, c),
        )

        if len(selected_indep) == 0:
            st.info("Pilih minimal satu variabel independen.")
        else:
            reg_df = df[["kabupaten", "tahun", dep_var] + selected_indep].dropna().copy()

            if len(reg_df) < len(selected_indep) + 2:
                st.warning("Jumlah observasi tidak cukup. Kurangi variabel independen atau perluas filter.")
            else:
                Y = reg_df[dep_var].astype(float)
                X = reg_df[selected_indep].copy()

                if model_type == "Fixed Effect - Entity (Kabupaten/Kota)":
                    entity_dummies = pd.get_dummies(reg_df["kabupaten"], drop_first=True, prefix="FE_kab")
                    X = pd.concat([X, entity_dummies], axis=1)
                elif model_type == "Fixed Effect - Two Way (Kabupaten/Kota & Tahun)":
                    entity_dummies = pd.get_dummies(reg_df["kabupaten"], drop_first=True, prefix="FE_kab")
                    time_dummies = pd.get_dummies(reg_df["tahun"].astype(str), drop_first=True, prefix="FE_thn")
                    X = pd.concat([X, entity_dummies, time_dummies], axis=1)

                X = X.astype(float)
                X = sm.add_constant(X)

                try:
                    model = sm.OLS(Y, X).fit(cov_type="cluster", cov_kwds={"groups": reg_df["kabupaten"]})
                except Exception:
                    model = sm.OLS(Y, X).fit()

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("N Observasi", f"{int(model.nobs)}")
                m2.metric("R-squared", f"{model.rsquared:.3f}")
                m3.metric("Adj. R-squared", f"{model.rsquared_adj:.3f}")
                m4.metric("F-statistic", f"{model.fvalue:.2f}" if model.fvalue is not None else "-")

                st.markdown("---")
                st.subheader("Hasil Estimasi Koefisien (Variabel Utama)")

                main_vars = ["const"] + selected_indep
                coef_df = pd.DataFrame({
                    "Variabel": ["Konstanta" if v == "const" else label_map.get(v, v) for v in main_vars],
                    "Koefisien": [model.params[v] for v in main_vars],
                    "Std. Error": [model.bse[v] for v in main_vars],
                    "t-stat": [model.tvalues[v] for v in main_vars],
                    "p-value": [model.pvalues[v] for v in main_vars],
                })
                coef_df["Signifikan (p<0,05)"] = coef_df["p-value"] < 0.05

                def _highlight_sig(row):
                    color = "background-color: #d4edda" if row["Signifikan (p<0,05)"] else ""
                    return [color] * len(row)

                st.dataframe(
                    coef_df.style.apply(_highlight_sig, axis=1).format({
                        "Koefisien": "{:.4f}", "Std. Error": "{:.4f}", "t-stat": "{:.3f}", "p-value": "{:.4f}",
                    }),
                    use_container_width=True,
                    height=min(450, 45 * len(coef_df) + 40),
                )

                if model_type != "Pooled OLS":
                    with st.expander("Lihat koefisien Fixed Effect (dummy kabupaten/tahun)"):
                        fe_vars = [v for v in X.columns if v.startswith("FE_")]
                        if fe_vars:
                            fe_df = pd.DataFrame({
                                "Dummy": fe_vars,
                                "Koefisien": [model.params[v] for v in fe_vars],
                                "p-value": [model.pvalues[v] for v in fe_vars],
                            })
                            st.dataframe(fe_df, use_container_width=True, height=300)

                st.markdown("---")
                st.subheader("Visualisasi Koefisien Variabel Independen")
                plot_df = coef_df[coef_df["Variabel"] != "Konstanta"].sort_values("Koefisien")
                fig_coef = px.bar(
                    plot_df, x="Koefisien", y="Variabel", orientation="h",
                    color="Signifikan (p<0,05)",
                    color_discrete_map={True: "#1a9850", False: "#bdbdbd"},
                    error_x=plot_df["Std. Error"] * 1.96,
                    title=f"Koefisien Regresi terhadap {dep_options_available[dep_var]} ({model_type})",
                )
                fig_coef.update_layout(height=max(350, 30 * len(plot_df)), margin=dict(t=40))
                st.plotly_chart(fig_coef, use_container_width=True)

                st.markdown("---")
                st.subheader("Prediksi vs Aktual")
                pred_df = reg_df[["kabupaten", "tahun"]].copy()
                pred_df["Aktual"] = Y.values
                pred_df["Prediksi"] = model.predict(X).values
                fig_pred = px.scatter(
                    pred_df, x="Aktual", y="Prediksi", color="kabupaten", hover_data=["tahun"],
                    title="Perbandingan Nilai Aktual vs Prediksi Model",
                )
                min_v = float(pred_df[["Aktual", "Prediksi"]].min().min())
                max_v = float(pred_df[["Aktual", "Prediksi"]].max().max())
                fig_pred.add_shape(type="line", x0=min_v, y0=min_v, x1=max_v, y1=max_v,
                                    line=dict(color="gray", dash="dash"))
                fig_pred.update_layout(height=480)
                st.plotly_chart(fig_pred, use_container_width=True)

                with st.expander("📋 Ringkasan Model Lengkap (statsmodels summary)"):
                    st.text(model.summary().as_text())

st.markdown("---")
st.caption("Dashboard dibuat dengan Cinta & Python · Data diolah TIM ISEI MAKASSAR.")
