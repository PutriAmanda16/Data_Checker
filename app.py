import streamlit as st
import pandas as pd
from io import BytesIO
from html import escape
import hashlib
import hmac
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

st.set_page_config(
    page_title=" DATA QUALITY SYSTEM",
    page_icon=":material/database:",
    layout="wide"
)

# =====================================================
# SYSTEM AUTHENTICATION / LOGIN STATE
# =====================================================

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

def load_login_credentials():
    """Ambil email & password dari .streamlit/secrets.toml
    (tidak ditulis di kode supaya aman dibagikan / didaftarkan HKI)."""
    try:
        return str(st.secrets["auth"]["email"]), str(st.secrets["auth"]["password"])
    except Exception:
        return None, None


def render_login_page():
    # 1. CSS khusus halaman login (tidak ikut ke halaman utama setelah login)
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

    :root {
        --slate-900: #0F172A; --slate-800: #1E293B; --slate-700: #334155;
        --slate-600: #475569; --slate-400: #94A3B8; --slate-300: #CBD5E1;
        --slate-200: #E2E8F0; --slate-100: #F1F5F9; --slate-50: #F8FAFC;
        --ok: #059669;
    }

    /* ---------- Reset layout Streamlit ---------- */
    [data-testid="stHeader"], [data-testid="stSidebar"], [data-testid="stToolbar"] {
        display: none !important;
    }
    .stApp { background: var(--slate-50) !important; }
    .block-container,
    [data-testid="stMainBlockContainer"] {
        padding: 0 !important;
        max-width: 100% !important;
    }
    [data-testid="stHorizontalBlock"] { gap: 0 !important; }
    [data-testid="stColumn"] > [data-testid="stVerticalBlock"] { gap: 0 !important; }
    html, body, .stApp, .stApp * { font-family: 'Inter', sans-serif; }

    /* ---------- Hero (kiri) ---------- */
    .login-hero {
        background: var(--slate-900);
        border-right: 1px solid rgba(255,255,255,.08);
        padding: 3.5rem 4rem;
        min-height: 100vh;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-sizing: border-box;
    }
    .hero-brand {
        display: flex; align-items: center; gap: .85rem; margin-bottom: 3rem;
    }
    .hero-logo {
        width: 40px; height: 40px; border-radius: 9px; background: #FFFFFF;
        display: flex; align-items: center; justify-content: center; flex-shrink: 0;
    }
    .hero-brand-name {
        display: flex; align-items: center; gap: .5rem;
        font-weight: 700; font-size: 1rem; letter-spacing: .06em; color: #FFFFFF;
    }
    .hero-ver {
        background: rgba(255,255,255,.12); color: #FFFFFF;
        font-size: .65rem; font-weight: 500; letter-spacing: 0;
        padding: 2px 7px; border-radius: 4px;
    }
    .hero-brand-sub {
        font-size: .68rem; color: #94A3B8; letter-spacing: .1em; margin-top: 2px;
    }
    .hero-badge {
        display: inline-flex; align-items: center; gap: .5rem;
        background: rgba(255,255,255,.06); color: var(--slate-300);
        border: 1px solid rgba(255,255,255,.12);
        padding: 5px 14px; border-radius: 9999px;
        font-size: .75rem; font-weight: 600;
    }
    .hero-dot { width: 7px; height: 7px; border-radius: 50%; background: #10B981; }
    .hero-title {
        font-size: 2.6rem; font-weight: 700; line-height: 1.15;
        letter-spacing: -0.02em; margin: 1.4rem 0 1rem; color: #FFFFFF;
    }
    .hero-desc {
        color: var(--slate-400); font-size: .95rem; line-height: 1.7; max-width: 480px;
    }
    .hero-card {
        background: rgba(255,255,255,.05);
        border: 1px solid rgba(255,255,255,.1);
        border-radius: 14px; padding: 1.3rem 1.4rem; margin-top: 2.2rem; max-width: 520px;
    }
    .hero-card-top {
        display: flex; justify-content: space-between; align-items: center;
        font-family: 'JetBrains Mono', monospace; font-size: .75rem; color: #94A3B8;
    }
    .hero-card-ok {
        color: #10B981; background: rgba(16,185,129,.12);
        padding: 3px 8px; border-radius: 4px; font-weight: 500;
    }
    .hero-stats {
        display: grid; grid-template-columns: repeat(3, auto);
        justify-content: start; gap: 2rem; margin-top: 1.1rem;
    }
    .hero-stat-l { color: #94A3B8; font-size: .65rem; letter-spacing: .05em; }
    .hero-stat-v { color: #10B981; font-weight: 700; font-size: .85rem; margin-top: 2px; }
    .hero-foot {
        font-size: .72rem; color: #64748B; display: flex; gap: .8rem; margin-top: 2.5rem;
    }

    .stApp .hero-card-top, .stApp .hero-card-top * {
        font-family: 'JetBrains Mono', ui-monospace, monospace;
    }

    /* ---------- Form (kanan) ---------- */
    [data-testid="stColumn"]:has(.login-form-container) {
        background: var(--slate-50);
    }
    [data-testid="stColumn"]:has(.login-form-container) > [data-testid="stVerticalBlock"] {
        min-height: 100vh;
        justify-content: center;
        align-items: center;
        padding: 2rem 1.5rem;
        box-sizing: border-box;
    }
    .login-form-container,
    [data-testid="stForm"],
    [data-testid="stColumn"]:has(.login-form-container) [data-testid="stAlert"] {
        width: 100% !important;
        max-width: 400px !important;
        margin-left: auto; margin-right: auto;
    }
    .sso-chip {
        display: inline-block; background: #DAE2FD; color: #3F465C;
        font-size: .68rem; font-weight: 700; letter-spacing: .06em; padding: 4px 11px; border-radius: 6px;
    }
    .welcome-title {
        font-size: 2rem; font-weight: 700; color: #0F172A;
        letter-spacing: -0.02em; margin-top: 1rem;
    }
    .welcome-sub {
        color: var(--slate-600); font-size: .9rem; line-height: 1.6;
        margin: .4rem 0 1.8rem;
    }

    [data-testid="stForm"] {
        border: none !important; padding: 0 !important; background: transparent !important;
    }
    [data-testid="stForm"] [data-testid="stVerticalBlock"] { gap: 1.1rem !important; }

    [data-testid="stTextInput"] label p {
        font-size: .72rem !important; font-weight: 600 !important;
        letter-spacing: .06em; color: #475569 !important;
    }
    [data-testid="stTextInput"] div[data-baseweb="input"] {
        background: #FFFFFF !important;
        border: 1px solid var(--slate-200) !important;
        border-radius: 10px !important;
        height: 3rem;
        transition: border-color .15s, box-shadow .15s;
    }
    [data-testid="stTextInput"] div[data-baseweb="input"]:focus-within {
        border-color: var(--slate-900) !important;
        box-shadow: 0 0 0 3px rgba(15,23,42,.10);
    }
    [data-testid="stTextInput"] input {
        background: transparent !important; color: #0F172A !important;
        font-size: .92rem !important;
    }
    [data-testid="stTextInput"] input::placeholder { color: #94A3B8 !important; }
    [data-testid="stTextInput"] button { background: transparent !important; color: #64748B !important; }

    [data-testid="stFormSubmitButton"] button {
        background: var(--slate-900) !important; color: #FFFFFF !important;
        border: none !important; border-radius: 10px !important;
        height: 3rem; font-weight: 600; font-size: .95rem;
        margin-top: .4rem; transition: background .15s;
    }
    [data-testid="stFormSubmitButton"] button:hover { background: var(--slate-800) !important; }
    [data-testid="stFormSubmitButton"] button p { color: #FFFFFF !important; }

    .login-note {
        text-align: center; color: #94A3B8; font-size: .72rem; margin-top: 1.5rem;
    }

    /* ---------- Responsif ---------- */
    @media (max-width: 900px) {
        .login-hero { min-height: auto; padding: 2rem 1.5rem; }
        .hero-title { font-size: 2rem; }
        [data-testid="stColumn"]:has(.login-form-container) > [data-testid="stVerticalBlock"] {
            min-height: auto; padding: 2.5rem 1.5rem;
        }
    }
    </style>
    """, unsafe_allow_html=True)

    # 2. Dua kolom: hero (kiri) dan form (kanan)
    col_hero, col_form = st.columns([1.1, 1], gap="small")

    with col_hero:
        st.html("""
        <div class="login-hero">
            <div>
                <div class="hero-brand">
                    <div class="hero-logo"><img src="data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgd2lkdGg9IjI0IiBoZWlnaHQ9IjI0IiBmaWxsPSIjMEYxNzJBIj48cGF0aCBkPSJNMTIgM0M3LjU4IDMgNCA0Ljc5IDQgN3YxMGMwIDIuMjEgMy41OSA0IDggNHM4LTEuNzkgOC00VjdjMC0yLjIxLTMuNTgtNC04LTR6bTYgMTRjMCAuNS0yLjEzIDItNiAycy02LTEuNS02LTJ2LTIuMjNjMS42MS43OCAzLjcyIDEuMjMgNiAxLjIzczQuMzktLjQ1IDYtMS4yM1YxN3ptMC00LjU1Yy0xLjMuOTUtMy41OCAxLjU1LTYgMS41NXMtNC43LS42LTYtMS41NVY5LjY0YzEuNDcuODMgMy42MSAxLjM2IDYgMS4zNnM0LjUzLS41MyA2LTEuMzZ2Mi44MXpNMTIgOUM4LjEzIDkgNiA3LjUgNiA3czIuMTMtMiA2LTIgNiAxLjUgNiAyLTIuMTMgMi02IDJ6Ii8+PC9zdmc+" width="24" height="24" alt="logo"/></div>
                    <div>
                        <div class="hero-brand-name">DATA CHECKER <span class="hero-ver">v2.4 PRO</span></div>
                        <div class="hero-brand-sub">BAPENDA DATA QUALITY SYSTEM</div>
                    </div>
                </div>

                <div class="hero-badge"><span class="hero-dot"></span>Pemeriksaan Kualitas Data Otomatis</div>
                <div class="hero-title">Flawless datasets,<br/>zero pipeline friction.</div>
                <div class="hero-desc">Deteksi missing values, duplikasi, dan ketidaksesuaian tipe data pada dataset pendapatan daerah, lengkap dengan pembersihan dan ekspor laporan.</div>

                <div class="hero-card">
                    <div class="hero-card-top">
                        <span>contoh_dataset.xlsx</span>
                        <span class="hero-card-ok">CONTOH TAMPILAN</span>
                    </div>
                    <div class="hero-stats">
                        <div><div class="hero-stat-l">MISSING CELLS</div><div class="hero-stat-v">0.0%</div></div>
                        <div><div class="hero-stat-l">DUPLICATES</div><div class="hero-stat-v">0 Clean</div></div>
                        <div><div class="hero-stat-l">TYPE COERCION</div><div class="hero-stat-v">PASS</div></div>
                    </div>
                </div>
            </div>

            <div class="hero-foot">
                <span>Dikembangkan selama magang di Bapenda Kota Surabaya</span>
            </div>
        </div>
        """)

    with col_form:
        st.html("""
        <div class="login-form-container">
            <div class="sso-chip">🔑 Akses Workspace Internal</div>
            <div class="welcome-title">Selamat Datang</div>
            <div class="welcome-sub">Masukkan kredensial Anda untuk mengakses workspace analisis data.</div>
        </div>
        """)

        with st.form("login_form", clear_on_submit=False):
            email = st.text_input("EMAIL WORKSPACE", placeholder="nama@instansi.go.id")
            password = st.text_input("PASSWORD", type="password", placeholder="••••••••••••")

            submitted = st.form_submit_button("Sign In to Workspace →", use_container_width=True, type="primary")

            if submitted:
                if email == "" or password == "":
                    st.warning("Mohon isi Email dan Password.")
                else:
                    valid_email, valid_password = load_login_credentials()

                    if valid_email is None:
                        st.error(
                            "Kredensial login belum dikonfigurasi. "
                            "Isi .streamlit/secrets.toml (lihat secrets.toml.example)."
                        )
                    elif (
                        hmac.compare_digest(email.strip().encode(), valid_email.encode())
                        and hmac.compare_digest(password.encode(), valid_password.encode())
                    ):
                        st.session_state["authenticated"] = True
                        st.success("Login Berhasil!")
                        st.rerun()
                    else:
                        st.error("Email atau Password salah!")

        st.html('<div class="login-form-container"><div class="login-note">Akses terbatas untuk pegawai</div></div>')

# Cek Status Autentikasi
if not st.session_state["authenticated"]:
    render_login_page()
    st.stop()  # Menghentikan eksekusi kode di bawahnya sampai user berhasil login

# =====================================================
# CUSTOM UI (Precision Metric design system)
# =====================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,400,0,0&display=block');

:root {
    --slate-900: #0F172A;
    --slate-800: #1E293B;
    --slate-700: #334155;
    --slate-600: #475569;
    --slate-400: #94A3B8;
    --slate-300: #CBD5E1;
    --slate-200: #E2E8F0;
    --slate-100: #F1F5F9;
    --slate-50:  #F8FAFC;
    --white: #FFFFFF;
    --ok: #059669;     --ok-bg: #ECFDF5;   --ok-bd: #A7F3D0;
    --warn: #D97706;   --warn-bg: #FFFBEB; --warn-bd: #FDE68A;
    --err: #E11D48;    --err-bg: #FFF1F2;  --err-bd: #FECDD3;
    --mono: 'JetBrains Mono', ui-monospace, monospace;
}

html, body, .stApp, .stApp p, .stApp label, .stApp input, .stApp textarea,
.stApp button, .stApp th, .stApp td, .stApp li, .stApp h1, .stApp h2, .stApp h3 {
    font-family: 'Inter', sans-serif;
}
[data-testid="stIconMaterial"], [data-testid="stIconMaterial"] * {
    font-family: 'Material Symbols Rounded' !important;
}

.stApp {
    background: var(--slate-50);
    color: var(--slate-900);
}

/* Streamlit chrome */
#MainMenu, footer, [data-testid="stDecoration"] { display: none; }
[data-testid="stHeader"] { background: transparent; height: 2.5rem; }

.block-container {
    max-width: 1560px;
    padding: 1.4rem 2rem 3rem;
}

h1, h2, h3, h4 {
    color: var(--slate-900) !important;
    font-family: 'Inter', sans-serif !important;
    letter-spacing: -0.02em;
    font-weight: 600 !important;
}
h3 { font-size: 1.05rem !important; }
p, label, li { color: var(--slate-900); }

.material-symbols-outlined {
    font-family: 'Material Symbols Rounded', 'Material Symbols Outlined' !important;
    font-weight: 400;
    font-style: normal;
    line-height: 1;
    letter-spacing: normal;
    text-transform: none;
    display: inline-block;
    white-space: nowrap;
    -webkit-font-feature-settings: 'liga';
    font-feature-settings: 'liga';
    vertical-align: middle;
}

/* ---------- SIDEBAR ---------- */
[data-testid="stSidebar"][aria-expanded="true"] {
    width: 260px !important;
    min-width: 260px !important;
}
[data-testid="stSidebar"] {
    background: var(--slate-900);
    border-right: 1px solid rgba(255,255,255,.08);
}
[data-testid="stSidebar"] > div:first-child { padding-top: .6rem; }
[data-testid="stSidebar"] * { color: #E2E8F0; }
[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,.08);
    margin: 1rem 0;
}

.brand {
    display: flex; align-items: center; gap: .65rem;
    padding: .4rem .2rem 1.1rem;
    border-bottom: 1px solid rgba(255,255,255,.08);
    margin-bottom: 1rem;
}
.brand-logo {
    width: 34px; height: 34px; border-radius: 8px;
    background: #FFFFFF; color: var(--slate-900) !important;
    display: flex; align-items: center; justify-content: center;
}
.brand-logo span { color: var(--slate-900) !important; font-size: 20px; }
.brand-name {
    font-size: .95rem; font-weight: 700; letter-spacing: .04em;
    color: #FFFFFF !important; line-height: 1.1;
}
.brand-sub {
    font-size: .6rem; letter-spacing: .14em; color: var(--slate-400) !important;
    margin-top: 3px; font-weight: 600;
}
.nav-label {
    font-size: .65rem; font-weight: 700; letter-spacing: .12em;
    color: var(--slate-400) !important; margin: .2rem 0 .5rem .3rem;
}

[data-testid="stSidebar"] [data-testid="stRadio"] > label,
[data-testid="stSidebar"] [data-testid="stRadio"] [data-testid="stWidgetLabel"] { display: none; }
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] { gap: .15rem; }
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"],
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"] {
    padding: .55rem .75rem;
    border-radius: 8px;
    width: 100%;
    margin: 0;
    transition: background .15s;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"] > div:first-child,
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"] > div > div:first-child {
    display: none;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label p {
    font-size: .86rem; font-weight: 500; color: #CBD5E1;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
    background: rgba(255,255,255,.06);
}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked),
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-selected="true"] {
    background: rgba(255,255,255,.12);
}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) p,
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-selected="true"] p {
    color: #FFFFFF; font-weight: 600;
}

.side-card {
    background: rgba(255,255,255,.05);
    border: 1px solid rgba(255,255,255,.08);
    border-radius: 10px;
    padding: .8rem .85rem;
    font-size: .75rem;
    line-height: 1.5;
}
.side-card .t { font-weight: 600; color: #FFFFFF !important; font-size: .75rem; margin-bottom: 4px; }
.side-card .m { color: var(--slate-400) !important; font-family: var(--mono); font-size: .68rem; }
.side-foot {
    display: flex; align-items: center; gap: .45rem;
    font-size: .65rem; letter-spacing: .1em; font-weight: 600;
    color: var(--slate-400) !important; margin-top: 1rem;
}
.dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; }
.dot.ok { background: #34D399; }
.dot.warn { background: #F59E0B; }
.dot.err { background: var(--err); }
.dot.off { background: var(--slate-400); }

/* ---------- PAGE HEADER ---------- */
.eyebrow {
    display: inline-flex; align-items: center; gap: .4rem;
    padding: 3px 10px; border-radius: 6px;
    background: #DAE2FD; color: #3F465C;
    font-size: .66rem; font-weight: 700; letter-spacing: .08em;
    text-transform: uppercase;
}
.eyebrow .material-symbols-outlined { font-size: 14px; }
.page-title {
    font-size: 1.6rem; font-weight: 600; letter-spacing: -.025em;
    color: var(--slate-900); line-height: 1.25; margin: .55rem 0 .3rem;
}
.page-desc {
    color: var(--slate-600); font-size: .85rem; line-height: 1.55; max-width: 780px;
}
.chip-row { display: flex; gap: .35rem; align-items: center; justify-content: flex-end; flex-wrap: wrap; }
.chip-row .lbl {
    font-size: .65rem; font-weight: 600; letter-spacing: .08em;
    color: var(--slate-600); margin-right: .2rem;
}
.code-chip {
    font-family: var(--mono); font-size: .7rem; font-weight: 500;
    padding: 2px 8px; border-radius: 4px;
    background: #E6E8EA; color: var(--slate-900);
}

/* ---------- FILE STRIP ---------- */
.file-strip {
    background: var(--white);
    border: 1px solid var(--slate-200);
    border-radius: 16px;
    padding: .9rem 1.1rem;
    display: flex; align-items: center; gap: .9rem;
    min-height: 92px;
    box-shadow: 0 1px 3px rgba(15,23,42,.04);
}
.file-ico {
    width: 46px; height: 46px; border-radius: 10px;
    background: var(--slate-100); color: var(--slate-800);
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.file-ico .material-symbols-outlined { font-size: 26px; }
.file-name { font-size: .95rem; font-weight: 600; color: var(--slate-900); word-break: break-all; }
.file-meta {
    display: flex; flex-wrap: wrap; align-items: center; gap: .5rem;
    margin-top: 4px; font-size: .74rem; color: var(--slate-600);
}
.file-meta .mono { font-family: var(--mono); font-size: .7rem; }
.size-pill {
    font-family: var(--mono); font-size: .68rem; font-weight: 500;
    padding: 1px 7px; border-radius: 4px; background: #85F8C4; color: #002114;
}

/* ---------- KPI ---------- */
.kpi {
    background: var(--white);
    border: 1px solid var(--slate-200);
    border-radius: 16px;
    padding: .9rem 1.1rem 1rem;
    box-shadow: 0 1px 3px rgba(15,23,42,.04);
    min-height: 132px;
    display: flex; flex-direction: column; justify-content: space-between;
}
.kpi-top { display: flex; align-items: center; justify-content: space-between; }
.kpi-label {
    font-size: .66rem; font-weight: 600; letter-spacing: .07em;
    color: var(--slate-600); text-transform: uppercase;
}
.kpi-ico .material-symbols-outlined { font-size: 18px; color: var(--slate-400); }
.kpi.warn .kpi-ico .material-symbols-outlined { color: var(--warn); }
.kpi.err  .kpi-ico .material-symbols-outlined { color: var(--err); }
.kpi.ok   .kpi-ico .material-symbols-outlined { color: var(--ok); }
.kpi-value {
    font-size: 2rem; font-weight: 600; letter-spacing: -.025em;
    color: var(--slate-900); line-height: 1.1; margin-top: .6rem;
    font-variant-numeric: tabular-nums;
}
.kpi.err .kpi-value { color: var(--err); }
.kpi-note { display: flex; align-items: center; gap: .4rem; margin-top: .3rem; font-size: .72rem; color: var(--slate-600); }

/* ---------- BADGES ---------- */
.badge {
    display: inline-block; border-radius: 9999px;
    font-size: .62rem; font-weight: 700; letter-spacing: .05em;
    padding: 2px 9px; text-transform: uppercase; border: 1px solid;
    font-family: 'Inter', sans-serif;
}
.badge.ok   { background: var(--ok-bg);   color: var(--ok);   border-color: var(--ok-bd); }
.badge.warn { background: var(--warn-bg); color: var(--warn); border-color: var(--warn-bd); }
.badge.err  { background: var(--err-bg);  color: var(--err);  border-color: var(--err-bd); }
.badge.neutral { background: var(--slate-100); color: var(--slate-700); border-color: var(--slate-200); }
.mono-tag {
    font-family: var(--mono); font-size: .68rem; font-weight: 500;
    padding: 2px 6px; border-radius: 4px; background: var(--slate-100); color: var(--slate-700);
}

/* ---------- STEPPER ---------- */
.step {
    background: var(--white);
    border: 1px solid var(--slate-200);
    border-radius: 16px;
    padding: .85rem 1.1rem;
    display: flex; align-items: center; gap: .9rem;
    box-shadow: 0 1px 3px rgba(15,23,42,.04);
}
.step-ico {
    width: 40px; height: 40px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.step-ico .material-symbols-outlined { font-size: 20px; }
.step-ico.done   { background: #82F5C1; color: #00714E; }
.step-ico.active { background: #DAE2FD; color: #131B2E; }
.step-ico.warn   { background: var(--warn-bg); color: var(--warn); }
.step-ico.idle   { background: #E6E8EA; color: var(--slate-400); }
.step-k { display: flex; gap: .5rem; align-items: baseline; flex-wrap: wrap; }
.step-k .n { font-size: .64rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: var(--slate-600); }
.step-k .s { font-family: var(--mono); font-size: .68rem; color: var(--slate-600); }
.step-t { font-size: .9rem; font-weight: 600; color: var(--slate-900); }
.step-t.idle { color: var(--slate-600); }

/* ---------- PANELS (bordered containers) ---------- */
[data-testid="stVerticalBlockBorderWrapper"] {
    background: var(--white);
    border-color: var(--slate-200) !important;
    border-radius: 16px !important;
    box-shadow: 0 1px 3px rgba(15,23,42,.04);
}
.panel-head { display: flex; align-items: center; justify-content: space-between; gap: .6rem; }
.panel-title { font-size: .95rem; font-weight: 600; color: var(--slate-900); letter-spacing: -.01em; }
.panel-sub { font-size: .76rem; color: var(--slate-600); margin: 2px 0 .6rem; line-height: 1.45; }
.kicker {
    font-family: var(--mono); font-size: .64rem; font-weight: 500;
    letter-spacing: .06em; color: var(--slate-400); text-transform: uppercase;
    margin-bottom: 2px;
}

.bar-row { margin: .55rem 0; }
.bar-head { display: flex; justify-content: space-between; font-size: .74rem; margin-bottom: 4px; }
.bar-head .c { font-family: var(--mono); font-weight: 500; color: var(--slate-900); }
.bar-head .v { font-family: var(--mono); color: var(--slate-600); }
.bar-track { height: 6px; border-radius: 9999px; background: var(--slate-200); overflow: hidden; }
.bar-fill { height: 6px; border-radius: 9999px; transition: width 0.3s ease; }

.sheet-card {
    background: var(--slate-50);
    border: 1px solid var(--slate-200);
    border-radius: 12px;
    padding: .85rem 1rem; min-height: 92px;
}
.sheet-card .n { font-family: var(--mono); font-size: .66rem; color: var(--slate-400); font-weight: 600; }
.sheet-card .t { font-size: .88rem; font-weight: 600; color: var(--slate-900); margin: 3px 0; }
.sheet-card .d { font-size: .74rem; color: var(--slate-600); line-height: 1.45; }

.ready-strip {
    display: flex; align-items: center; gap: .8rem;
    background: var(--ok-bg); border: 1px solid var(--ok-bd);
    border-radius: 12px; padding: .8rem 1rem;
}
.ready-strip .material-symbols-outlined { color: var(--ok); font-size: 26px; }
.ready-strip .t { font-size: .9rem; font-weight: 600; color: var(--slate-900); }
.ready-strip .d { font-size: .74rem; color: var(--slate-600); font-family: var(--mono); }

.empty-state {
    background: var(--white); border: 1px dashed var(--slate-300);
    border-radius: 16px; padding: 2.2rem 1.5rem; text-align: center;
}
.empty-state .material-symbols-outlined { font-size: 34px; color: var(--slate-400); }
.empty-state .t { font-weight: 600; margin-top: .4rem; font-size: 1rem; color: var(--slate-900); }
.empty-state .d { color: var(--slate-600); font-size: .82rem; margin-top: .25rem; }

.sec-gap { height: 1rem; }

/* ---------- KUSTOMISASI TEKS UPLOADER ---------- */
[data-testid="stFileUploaderDropzoneInstructions"] div span {
    display: none !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] div::before {
    content: "Pilih File";
    font-weight: 500;
    font-size: 0.88rem;
    color: var(--slate-700);
}

/* ---------- STREAMLIT WIDGETS ---------- */
[data-testid="stMetric"] {
    background: var(--slate-50);
    border: 1px solid var(--slate-200);
    border-radius: 12px;
    padding: .75rem .95rem;
}
[data-testid="stMetricLabel"] p {
    color: var(--slate-600) !important;
    font-size: .68rem !important; font-weight: 600 !important;
    text-transform: uppercase; letter-spacing: .06em;
}
[data-testid="stMetricValue"] {
    color: var(--slate-900) !important;
    font-size: 1.6rem !important; font-weight: 600 !important;
    font-variant-numeric: tabular-nums;
}

.stButton > button, .stDownloadButton > button {
    min-height: 2.3rem;
    border-radius: 8px;
    background: var(--slate-900);
    border: 1px solid var(--slate-900);
    color: #FFFFFF;
    font-weight: 500; font-size: .82rem;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.15);
}
.stButton > button:hover, .stDownloadButton > button:hover {
    background: var(--slate-800); border-color: var(--slate-800); color: #FFFFFF;
}
.stButton > button p, .stDownloadButton > button p { color: #FFFFFF !important; }
.stButton > button:focus:not(:active), .stDownloadButton > button:focus:not(:active) {
    color: #FFFFFF; border-color: var(--slate-900);
}
.stButton > button[kind="secondary"] {
    background: #FFFFFF; color: var(--slate-900); border: 1px solid var(--slate-200);
    box-shadow: none;
}
.stButton > button[kind="secondary"] p { color: var(--slate-900) !important; }
.stButton > button[kind="secondary"]:hover { background: var(--slate-100); border-color: var(--slate-300); }
.stButton > button[kind="primary"] { background: var(--slate-900); }
.stDownloadButton > button { width: 100%; }

[data-testid="stFileUploader"] section {
    background: var(--slate-50) !important;
    border: 1px dashed var(--slate-300) !important;
    border-radius: 12px !important;
    padding: .7rem .9rem !important;
}
[data-testid="stFileUploader"] section * { color: var(--slate-700) !important; }
[data-testid="stFileUploader"] small { font-size: .7rem; }
[data-testid="stFileUploader"] button {
    background: #FFFFFF !important; border: 1px solid var(--slate-200) !important;
    color: var(--slate-900) !important;
}
[data-testid="stFileUploader"] button p { color: var(--slate-900) !important; }
[data-testid="stFileUploaderFile"] * { color: var(--slate-700) !important; }
[data-testid="stFileUploaderFile"] svg { color: var(--slate-600) !important; }
[data-testid="stFileUploaderFile"] > div:first-child, [data-testid="stFileUploaderFileData"] > div:first-child {
    background: var(--slate-100) !important; border-radius: 8px;
}

[data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="base-input"] {
    border-radius: 8px !important;
}
[data-testid="stSelectbox"] label p, [data-testid="stMultiSelect"] label p,
[data-testid="stTextInput"] label p, [data-testid="stRadio"] label p {
    font-size: .76rem; font-weight: 500; color: var(--slate-700);
}

[data-testid="stDataFrame"] {
    border: 1px solid var(--slate-200);
    border-radius: 12px;
    overflow: hidden;
}

[data-baseweb="tab-list"] {
    gap: .25rem;
    background: var(--slate-100);
    padding: .4rem .5rem 0;
    border-radius: 12px 12px 0 0;
    border-bottom: 1px solid var(--slate-200);
}
[data-baseweb="tab"] {
    height: 2.6rem;
    padding: 0 1rem;
    border-radius: 8px 8px 0 0;
    background: transparent;
}
[data-baseweb="tab"] p { font-size: .84rem; font-weight: 500; color: var(--slate-600); }
[data-baseweb="tab"][aria-selected="true"] { background: #FFFFFF; }
[data-baseweb="tab"][aria-selected="true"] p { color: var(--slate-900); font-weight: 600; }
[data-baseweb="tab-highlight"] { background: var(--slate-900) !important; height: 2px !important; }
[data-baseweb="tab-border"] { display: none; }
[data-testid="stTabs"] [role="tabpanel"] {
    background: #FFFFFF;
    border: 1px solid var(--slate-200);
    border-top: none;
    border-radius: 0 0 16px 16px;
    padding: 1.1rem 1.2rem 1.2rem;
}

[data-testid="stExpander"] {
    background: #FFFFFF; border: 1px solid var(--slate-200); border-radius: 12px;
}
[data-testid="stAlert"] { border-radius: 10px; }
hr { border-color: var(--slate-200); margin: 1rem 0; }

@media (max-width: 768px) {
    .block-container { padding: 1rem; }
    .page-title { font-size: 1.3rem; }
    .chip-row { justify-content: flex-start; }
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# FUNGSI PENGOLAHAN DATA
# =====================================================

BOOLEAN_VALUES = {
    "true", "false", "1", "0", "yes", "no", "ya", "tidak"
}
TYPE_OPTIONS = ["Teks", "Numerik", "Tanggal", "Boolean"]


def load_dataset(uploaded_file):
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    try:
        if name.endswith(".csv"):
            try:
                return pd.read_csv(BytesIO(raw), encoding="utf-8-sig")
            except UnicodeDecodeError:
                return pd.read_csv(BytesIO(raw), encoding="latin-1")

        if name.endswith((".xlsx", ".xls")):
            return pd.read_excel(BytesIO(raw))

        raise ValueError("Format file tidak didukung.")

    except Exception as error:
        st.error(f"Gagal membaca dataset: {error}")
        return None


def get_missing_mask(data):
    mask = data.isna().copy()

    for position in range(data.shape[1]):
        series = data.iloc[:, position]
        blank = series.map(
            lambda value: isinstance(value, str) and value.strip() == ""
        )
        mask.iloc[:, position] = mask.iloc[:, position] | blank

    return mask


def build_missing_table(data):
    mask = get_missing_mask(data)
    counts = mask.sum()

    percentages = (
        counts / len(data) * 100
        if len(data) > 0
        else pd.Series(0.0, index=data.columns)
    )

    return pd.DataFrame({
        "Nama Kolom": [str(c) for c in data.columns],
        "Jumlah Missing Values": counts.to_numpy(),
        "Persentase Missing Values": percentages.round(2).to_numpy()
    })


def detect_expected_type(series):
    values = series.dropna()
    values = values[
        ~values.map(
            lambda value: isinstance(value, str) and value.strip() == ""
        )
    ]

    if values.empty:
        return "Teks"

    normalized = values.astype(str).str.strip().str.lower()

    if normalized.isin(BOOLEAN_VALUES).mean() >= 0.8:
        return "Boolean"

    numeric = pd.to_numeric(values, errors="coerce")
    if numeric.notna().mean() >= 0.8:
        return "Numerik"

    dates = pd.to_datetime(values, errors="coerce")
    if dates.notna().mean() >= 0.8:
        return "Tanggal"

    return "Teks"


def get_expected_type(position, series):
    saved = st.session_state.get("saved_types", {})
    return saved.get(position) or detect_expected_type(series)


def check_data_types(data):
    results = []
    details = []
    missing = get_missing_mask(data)

    for position, column in enumerate(data.columns):
        series = data.iloc[:, position]
        expected = get_expected_type(position, series)

        actual = str(series.dtype)
        non_missing = ~missing.iloc[:, position]

        if expected == "Numerik":
            converted = pd.to_numeric(series, errors="coerce")
            invalid = non_missing & converted.isna()

        elif expected == "Tanggal":
            converted = pd.to_datetime(series, errors="coerce")
            invalid = non_missing & converted.isna()

        elif expected == "Boolean":
            normalized = series.astype(str).str.strip().str.lower()
            invalid = non_missing & ~normalized.isin(BOOLEAN_VALUES)

        else:
            is_text = series.map(lambda value: isinstance(value, str))
            invalid = non_missing & ~is_text

        count = int(invalid.sum())

        results.append({
            "Nama Kolom": str(column),
            "Tipe Data Aktual": actual,
            "Tipe Data Diharapkan": expected,
            "Jumlah Nilai Tidak Sesuai": count,
            "Status": "Sesuai" if count == 0 else "Perlu Pemeriksaan"
        })

        invalid_indices = data.index[invalid].tolist()
        for row_pos in invalid_indices:
            details.append({
                "Nama Kolom": str(column),
                "Nomor Baris": row_pos + 2,
                "Nilai Tidak Sesuai": str(series.iloc[row_pos]),
                "Tipe Data Diharapkan": expected
            })

    type_table = pd.DataFrame(results, columns=[
        "Nama Kolom", "Tipe Data Aktual", "Tipe Data Diharapkan",
        "Jumlah Nilai Tidak Sesuai", "Status"
    ])

    detail_table = pd.DataFrame(details, columns=[
        "Nama Kolom", "Nomor Baris", "Nilai Tidak Sesuai",
        "Tipe Data Diharapkan"
    ])

    return type_table, detail_table


def check_duplicates(data, mode, selected_columns):
    if mode == "Kolom tertentu":
        if not selected_columns:
            return pd.Series(False, index=data.index), data.iloc[0:0].copy()

        mask = data.duplicated(subset=selected_columns, keep=False)
    else:
        mask = data.duplicated(keep=False)

    return mask, data.loc[mask].copy()


def clean_missing_values(data, selected_columns):
    cleaned = data.copy()

    for column in selected_columns:
        cleaned[column] = cleaned[column].replace(
            r"^\s*$", pd.NA, regex=True
        )

    return cleaned.dropna(subset=selected_columns)


def impute_missing_values(data, column, strategy, custom_value=""):
    cleaned = data.copy()
    series = cleaned[column].replace(r"^\s*$", pd.NA, regex=True)

    if strategy == "Mean (Rata-rata)":
        num = pd.to_numeric(series, errors="coerce")
        cleaned[column] = series.fillna(num.mean())
    elif strategy == "Median (Nilai Tengah)":
        num = pd.to_numeric(series, errors="coerce")
        cleaned[column] = series.fillna(num.median())
    elif strategy == "Modus (Nilai Sering Muncul)":
        mode_val = series.mode()
        if not mode_val.empty:
            cleaned[column] = series.fillna(mode_val[0])
    elif strategy == "Nilai Kustom":
        cleaned[column] = series.fillna(custom_value)

    return cleaned


def normalize_text_column(data, column, action):
    cleaned = data.copy()
    series = cleaned[column].astype(str)

    if action == "Trim Whitespace":
        cleaned[column] = series.str.strip()
    elif action == "UPPERCASE":
        cleaned[column] = series.str.upper()
    elif action == "lowercase":
        cleaned[column] = series.str.lower()
    elif action == "Title Case":
        cleaned[column] = series.str.title()

    return cleaned


def clean_duplicates(data, mode, selected_columns):
    cleaned = data.copy()

    if mode == "Kolom tertentu":
        if not selected_columns:
            return cleaned

        cleaned = cleaned.drop_duplicates(
            subset=selected_columns,
            keep="first"
        )
    else:
        cleaned = cleaned.drop_duplicates(keep="first")

    return cleaned


def convert_column_type(data, column, target_type):
    cleaned = data.copy()
    series = cleaned[column]

    if target_type == "Numerik":
        converted = pd.to_numeric(series, errors="coerce")

    elif target_type == "Tanggal":
        converted = pd.to_datetime(series, errors="coerce")

    elif target_type == "Boolean":
        mapping = {
            "true": True, "false": False,
            "1": True, "0": False,
            "yes": True, "no": False,
            "ya": True, "tidak": False
        }

        converted = series.map(
            lambda value: mapping.get(str(value).strip().lower())
            if pd.notna(value) else pd.NA
        )

    else:
        converted = series.astype("string")

    original_non_missing = ~get_missing_mask(
        data[[column]]
    ).iloc[:, 0]

    failed = original_non_missing & converted.isna()

    if failed.any():
        return None, int(failed.sum())

    cleaned[column] = converted
    return cleaned, 0


def generate_clean_excel(data):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(
            writer,
            sheet_name="Data Clean",
            index=False
        )

        worksheet = writer.sheets["Data Clean"]
        worksheet.freeze_panes = "A2"

        if worksheet.max_row and worksheet.max_column:
            worksheet.auto_filter.ref = worksheet.dimensions

        for cells in worksheet.columns:
            letter = cells[0].column_letter
            max_length = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in cells
            )
            worksheet.column_dimensions[letter].width = min(
                max_length + 3, 40
            )

    output.seek(0)
    return output.getvalue()


def generate_excel_report(
    data, file_name, duplicate_mode, selected_columns,
    type_table, detail_table
):
    output = BytesIO()

    missing_table = build_missing_table(data)
    total_missing = int(get_missing_mask(data).to_numpy().sum())
    total_cells = data.shape[0] * data.shape[1]
    missing_pct = total_missing / total_cells * 100 if total_cells else 0

    duplicate_mask, duplicate_table = check_duplicates(
        data, duplicate_mode, selected_columns
    )
    duplicate_count = int(duplicate_mask.sum())

    if duplicate_mode == "Kolom tertentu" and selected_columns:
        duplicate_method = "Kolom tertentu: " + ", ".join(selected_columns)
    elif duplicate_mode == "Kolom tertentu":
        duplicate_method = "Kolom tertentu: belum dipilih"
    else:
        duplicate_method = "Seluruh kolom"

    type_errors = int(type_table["Jumlah Nilai Tidak Sesuai"].sum())

    summary = pd.DataFrame({
        "Indikator": [
            "Nama File", "Jumlah Baris", "Jumlah Kolom", "Total Sel Data",
            "Total Missing Values", "Persentase Missing Values",
            "Metode Pemeriksaan Duplikasi",
            "Jumlah Baris Terindikasi Duplikat",
            "Jumlah Nilai Tidak Sesuai Tipe Data"
        ],
        "Hasil": [
            file_name, len(data), len(data.columns), total_cells,
            total_missing, f"{missing_pct:.2f}%", duplicate_method,
            duplicate_count, type_errors
        ]
    })

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Ringkasan", index=False)
        missing_table.to_excel(writer, sheet_name="Missing Values", index=False)
        duplicate_table.to_excel(writer, sheet_name="Data Duplikat", index=False)
        type_table.to_excel(writer, sheet_name="Tipe Data", index=False)
        detail_table.to_excel(writer, sheet_name="Detail Tipe Data", index=False)
        data.to_excel(writer, sheet_name="Data Asli", index=False)

        red_fill = PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid")

        ws_type = writer.sheets["Tipe Data"]
        ws_type.conditional_formatting.add(
            f"D2:D{len(type_table)+1}",
            CellIsRule(operator="greaterThan", formula=["0"], stopIfTrue=True, fill=red_fill)
        )

        ws_missing = writer.sheets["Missing Values"]
        ws_missing.conditional_formatting.add(
            f"B2:B{len(missing_table)+1}",
            CellIsRule(operator="greaterThan", formula=["0"], stopIfTrue=True, fill=red_fill)
        )

        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"

            if worksheet.max_row and worksheet.max_column:
                worksheet.auto_filter.ref = worksheet.dimensions

            for cells in worksheet.columns:
                letter = cells[0].column_letter
                max_length = max(
                    (
                        len(str(cell.value))
                        if cell.value is not None else 0
                    )
                    for cell in cells
                )
                worksheet.column_dimensions[letter].width = min(
                    max_length + 3, 40
                )

    output.seek(0)
    return output.getvalue()


# =====================================================
# KOMPONEN TAMPILAN
# =====================================================

PROBLEM_COLUMNS_TAIL = ["Jenis Masalah", "Keterangan", "Konfirmasi OPD"]


def build_problem_data(data, duplicate_mode, selected_columns):
    """Kumpulkan baris bermasalah (missing, tipe tidak sesuai, duplikat)
    dari data asli, untuk dikonfirmasi ke OPD."""
    columns = list(data.columns)
    missing_mask = get_missing_mask(data)
    _, type_detail = check_data_types(data)
    dup_mask, _ = check_duplicates(data, duplicate_mode, selected_columns)

    issues = []  # setiap isi: (posisi_baris, posisi_kolom, jenis, nilai, keterangan)

    for col_pos in range(len(columns)):
        flags = missing_mask.iloc[:, col_pos].tolist()
        for row_pos, flag in enumerate(flags):
            if flag:
                issues.append((
                    row_pos, col_pos, "Missing Value", "", "Sel kosong"
                ))

    name_to_pos = {}
    for col_pos, name in enumerate(columns):
        name_to_pos.setdefault(str(name), col_pos)

    for _, item in type_detail.iterrows():
        issues.append((
            int(item["Nomor Baris"]) - 2,
            name_to_pos.get(item["Nama Kolom"], -1),
            "Tipe Data Tidak Sesuai",
            item["Nilai Tidak Sesuai"],
            f"Seharusnya {item['Tipe Data Diharapkan']}"
        ))

    if dup_mask.any():
        if duplicate_mode == "Kolom tertentu" and selected_columns:
            dup_cols = list(selected_columns)
            group_id = data.groupby(dup_cols, dropna=False, sort=False).ngroup()
            dup_label = "Duplikat pada kolom: " + ", ".join(dup_cols)
        else:
            group_id = pd.Series(
                data.astype(object).apply(lambda r: "||".join(map(str, r)), axis=1)
                .factorize()[0],
                index=data.index
            )
            dup_label = "Duplikat seluruh kolom"

        flags = dup_mask.tolist()
        for row_pos, flag in enumerate(flags):
            if flag:
                issues.append((
                    row_pos, -1, "Data Duplikat", "",
                    f"{dup_label} (kelompok {int(group_id.iloc[row_pos]) + 1})"
                ))

    issue_columns = [
        "Nomor Baris", "Nama Kolom", "Jenis Masalah", "Nilai", "Keterangan"
    ]
    issues_df = pd.DataFrame(
        [
            (
                r + 2,
                str(columns[c]) if c >= 0 else "(seluruh baris)",
                jenis, nilai, ket
            )
            for r, c, jenis, nilai, ket in sorted(issues, key=lambda x: (x[0], x[1]))
        ],
        columns=issue_columns
    )

    if not issues:
        empty = data.iloc[0:0].copy()
        empty.insert(0, "Nomor Baris", pd.Series(dtype="int64"))
        for name in PROBLEM_COLUMNS_TAIL:
            empty[name] = pd.Series(dtype="object")
        return empty, issues_df, issues

    row_positions = sorted({item[0] for item in issues})
    problem = data.iloc[row_positions].copy()
    problem.insert(0, "Nomor Baris", [p + 2 for p in row_positions])

    jenis_per_row, ket_per_row = [], []
    for row_pos in row_positions:
        row_issues = [i for i in issues if i[0] == row_pos]
        jenis_per_row.append(", ".join(dict.fromkeys(i[2] for i in row_issues)))

        parts = []
        miss = [str(columns[i[1]]) for i in row_issues if i[2] == "Missing Value"]
        tipe = [
            f"{columns[i[1]]} ('{i[3]}')" for i in row_issues
            if i[2] == "Tipe Data Tidak Sesuai" and i[1] >= 0
        ]
        dup = [i[4] for i in row_issues if i[2] == "Data Duplikat"]
        if miss:
            parts.append("Kosong: " + ", ".join(miss))
        if tipe:
            parts.append("Tipe tidak sesuai: " + ", ".join(tipe))
        if dup:
            parts.append(dup[0])
        ket_per_row.append("; ".join(parts))

    problem["Jenis Masalah"] = jenis_per_row
    problem["Keterangan"] = ket_per_row
    problem["Konfirmasi OPD"] = ""

    return problem, issues_df, issues


def generate_problem_excel(problem, issues_df, issues):
    output = BytesIO()
    n_orig = problem.shape[1] - 1 - len(PROBLEM_COLUMNS_TAIL)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        problem.to_excel(writer, sheet_name="Data Bermasalah", index=False)
        issues_df.to_excel(writer, sheet_name="Rincian Masalah", index=False)

        ws = writer.sheets["Data Bermasalah"]
        ws.freeze_panes = "B2"

        fill_missing = PatternFill("solid", start_color="FDE68A", end_color="FDE68A")
        fill_type = PatternFill("solid", start_color="FECDD3", end_color="FECDD3")
        fill_dup = PatternFill("solid", start_color="FED7AA", end_color="FED7AA")
        fill_head = PatternFill("solid", start_color="E2E8F0", end_color="E2E8F0")

        for cell in ws[1]:
            cell.fill = fill_head

        row_to_excel = {
            int(num) - 2: idx + 2
            for idx, num in enumerate(problem["Nomor Baris"].tolist())
        }

        for row_pos, col_pos, jenis, _, _ in issues:
            excel_row = row_to_excel.get(row_pos)
            if excel_row is None:
                continue
            if jenis == "Data Duplikat":
                ws.cell(row=excel_row, column=1).fill = fill_dup
            elif col_pos >= 0:
                ws.cell(row=excel_row, column=col_pos + 2).fill = (
                    fill_missing if jenis == "Missing Value" else fill_type
                )

        for sheet in (ws, writer.sheets["Rincian Masalah"]):
            sheet.auto_filter.ref = sheet.dimensions
            for cells in sheet.columns:
                letter = cells[0].column_letter
                longest = max(
                    len(str(c.value)) if c.value is not None else 0
                    for c in cells
                )
                sheet.column_dimensions[letter].width = min(longest + 3, 45)

        ws.column_dimensions[ws.cell(row=1, column=n_orig + 4).column_letter].width = 22

    output.seek(0)
    return output.getvalue()


def icon(name):
    return f'<span class="material-symbols-outlined">{name}</span>'


def fmt_size(num_bytes):
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 ** 2:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / 1024 ** 2:.1f} MB"


def page_header(title, description, eyebrow, right_html=""):
    left, right = st.columns([1.5, 1], gap="large", vertical_alignment="center")

    with left:
        st.markdown(
            f'<div class="eyebrow">{icon("verified")}{escape(eyebrow)}</div>'
            f'<div class="page-title">{escape(title)}</div>'
            f'<div class="page-desc">{escape(description)}</div>',
            unsafe_allow_html=True
        )

    with right:
        if right_html:
            st.markdown(right_html, unsafe_allow_html=True)


def panel_heading(kicker, title, description=None):
    html = (
        f'<div class="kicker">{escape(kicker)}</div>'
        f'<div class="panel-title">{escape(title)}</div>'
    )
    if description:
        html += f'<div class="panel-sub">{escape(description)}</div>'
    st.markdown(html, unsafe_allow_html=True)


def kpi_card(label, value, note="", ico="info", tone="neutral"):
    st.markdown(
        f"""
        <div class="kpi {tone}">
            <div class="kpi-top">
                <span class="kpi-label">{label}</span>
                <span class="kpi-ico">{icon(ico)}</span>
            </div>
            <div>
                <div class="kpi-value">{value}</div>
                <div class="kpi-note">{note}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def badge(text, tone="neutral"):
    return f'<span class="badge {tone}">{escape(str(text))}</span>'


def step_card(number, sub, title, state, ico):
    title_class = "idle" if state == "idle" else ""
    st.markdown(
        f"""
        <div class="step">
            <div class="step-ico {state}">{icon(ico)}</div>
            <div style="min-width:0">
                <div class="step-k">
                    <span class="n">Langkah {number}</span>
                    <span class="s">{sub}</span>
                </div>
                <div class="step-t {title_class}">{title}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def empty_state(title, description, ico="upload_file"):
    st.markdown(
        f"""
        <div class="empty-state">
            {icon(ico)}
            <div class="t">{title}</div>
            <div class="d">{description}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def gap():
    st.markdown('<div class="sec-gap"></div>', unsafe_allow_html=True)


def style_status(table):
    colors = {
        "Lolos": "background-color:#ECFDF5;color:#059669;font-weight:600",
        "Peringatan": "background-color:#FFFBEB;color:#D97706;font-weight:600",
        "Error": "background-color:#FFF1F2;color:#E11D48;font-weight:600",
    }
    styler = table.style
    mapper = getattr(styler, "map", None) or styler.applymap
    return mapper(lambda v: colors.get(v, ""), subset=["Status"])


def flash(message):
    st.session_state["_flash"] = message


def show_flash():
    message = st.session_state.pop("_flash", None)
    if message:
        st.success(message)


# =====================================================
# SIDEBAR
# =====================================================

NAV = {
    "dashboard": ":material/dashboard: Dashboard",
    "inspect": ":material/rule: Pemeriksaan Data",
    "clean": ":material/auto_fix_high: Studio Pembersihan",
    "report": ":material/description: Laporan Excel",
}

st.sidebar.markdown(f"""
<div class="brand">
    <div class="brand-logo">{icon("database")}</div>
    <div>
        <div class="brand-name">DATA CHECKER</div>
        <div class="brand-sub">BAPENDA DATA QUALITY</div>
    </div>
</div>
<div class="nav-label">NAVIGASI UTAMA</div>
""", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "Navigasi",
    list(NAV.keys()),
    format_func=lambda key: NAV[key],
    label_visibility="collapsed"
)

sidebar_status = st.sidebar.empty()


# =====================================================
# HEADER DAN UPLOAD
# =====================================================

if menu == "dashboard":
    page_header(
        "Sistem Validasi & Higienitas Data Pendapatan Daerah",
        "Modul pra-pemrosesan data target dan realisasi dari Organisasi Perangkat Daerah (OPD). "
        "Mengidentifikasi missing values, duplikasi entitas, serta ketidaksesuaian tipe data "
        "sebelum data dipublikasikan ke Web Potensi Pendapatan.",
        "Regional Revenue Data Quality Systems",
        '<div class="chip-row"><span class="lbl">FORMAT:</span>'
        '<span class="code-chip">.CSV</span>'
        '<span class="code-chip">.XLSX</span>'
        '<span class="code-chip">.XLS</span></div>'
    )
elif menu == "inspect":
    page_header(
        "Pemeriksaan Kualitas Data",
        "Identifikasi masalah kelengkapan, duplikasi, dan konsistensi tipe data.",
        "Data Inspection"
    )
elif menu == "clean":
    page_header(
        "Studio Pembersihan Data",
        "Susun transformasi secara bertahap. Data asli tetap dipertahankan.",
        "Data Transformation"
    )
else:
    page_header(
        "Laporan Hasil Pemeriksaan",
        "Dokumentasi hasil pemeriksaan kualitas dataset dalam format Excel.",
        "Report Center"
    )

gap()

uploaded_file = st.session_state.get("uploaded_file_obj", None)

if uploaded_file is None:
    info_col, upload_col = st.columns([1.15, 1], gap="medium", vertical_alignment="center")
    with upload_col:
        new_file = st.file_uploader(
            "Pilih File",
            type=["csv", "xlsx", "xls"],
            help="Format yang didukung: CSV, XLSX, dan XLS.",
            key="main_dataset_uploader"
        )
        if new_file is not None:
            st.session_state["uploaded_file_obj"] = new_file
            st.rerun()

    with info_col:
        st.markdown(
            f"""
            <div class="file-strip">
                <div class="file-ico">{icon("upload_file")}</div>
                <div>
                    <div class="file-name">Belum ada dataset</div>
                    <div class="file-meta">Unggah berkas CSV atau Excel pada panel di samping untuk memulai pemeriksaan.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
else:
    file_col, action_col = st.columns([4.2, 1], gap="medium", vertical_alignment="center")

    with file_col:
        st.markdown(
            f"""
            <div class="file-strip">
                <div class="file-ico">{icon("description")}</div>
                <div style="min-width:0; flex-grow: 1;">
                    <div style="display:flex;gap:.5rem;align-items:center;flex-wrap:wrap">
                        <span class="file-name">{escape(uploaded_file.name)}</span>
                        <span class="size-pill">{fmt_size(len(uploaded_file.getvalue()))}</span>
                    </div>
                    <div class="file-meta">
                        <span class="dot ok"></span>
                        <span>Dataset berhasil dimuat</span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with action_col:
        if st.button("Ganti Dataset", icon=":material/delete:", use_container_width=True, type="secondary"):
            st.session_state.pop("uploaded_file_obj", None)
            st.session_state.pop("active_file_key", None)
            st.session_state.pop("cleaned_data", None)
            st.rerun()

df = None
file_key = None

if uploaded_file is not None:
    with st.spinner("Memuat dan menganalisis dataset..."):
        df = load_dataset(uploaded_file)
        file_key = hashlib.md5(uploaded_file.getvalue()).hexdigest()

if df is not None and df.empty:
    st.warning("Berkas berhasil dibaca, tetapi tidak terdapat baris data.")

# ---- Reset state ketika dataset berganti ----
if st.session_state.get("active_file_key") != file_key:
    st.session_state["active_file_key"] = file_key
    st.session_state["saved_types"] = {}
    st.session_state["saved_dup_mode"] = "Seluruh kolom"
    st.session_state["saved_dup_cols"] = []

    for state_key in list(st.session_state.keys()):
        if (
            state_key.startswith("expected_type_v2_")
            or state_key in (
                "selected_duplicate_columns", "clean_missing_columns",
                "clean_duplicate_columns", "clean_type_column"
            )
        ):
            del st.session_state[state_key]

    if df is not None:
        st.session_state["clean_file_key"] = file_key
        st.session_state["cleaned_data"] = df.copy()

has_data = df is not None and not df.empty

# ---- Ringkasan status di sidebar ----
if has_data:
    sidebar_status.markdown(
        f"""
        <hr/>
        <div class="side-card">
            <div class="t">DATASET AKTIF</div>
            <div class="m">{escape(uploaded_file.name)}</div>
            <div class="m">{len(df):,} baris · {len(df.columns):,} kolom</div>
        </div>
        <div class="side-foot"><span class="dot ok"></span>SIAP DIPERIKSA</div>
        """,
        unsafe_allow_html=True
    )
else:
    sidebar_status.markdown(
        """
        <hr/>
        <div class="side-card">
            <div class="t">DATA QUALITY WORKSPACE</div>
            <div class="m">Pemeriksaan, pembersihan, dan dokumentasi kualitas dataset.</div>
        </div>
        <div class="side-foot"><span class="dot off"></span>MENUNGGU DATASET</div>
        """,
        unsafe_allow_html=True
    )

gap()


# =====================================================
# DASBOR
# =====================================================

if menu == "dashboard":

    if df is None or df.empty:
        step_cols = st.columns(3, gap="small")
        guides = [
            ("1", "Pemuatan", "Unggah Dataset", "active", "upload_file"),
            ("2", "Validasi Aturan", "Periksa Kualitas", "idle", "troubleshoot"),
            ("3", "Remediation", "Bersihkan & Unduh", "idle", "cleaning_services"),
        ]
        for col, (num, sub, title, state, ico) in zip(step_cols, guides):
            with col:
                step_card(num, sub, title, state, ico)

        gap()
        empty_state(
            "Dataset belum tersedia",
            "Unggah dataset melalui panel di atas untuk menampilkan "
            "ringkasan struktur dan indikator kualitas data.",
            "dataset"
        )

        gap()
        module_cols = st.columns(4, gap="small")
        modules = [
            ("01", "Dashboard", "Ringkasan dan struktur dataset."),
            ("02", "Pemeriksaan Data", "Deteksi missing values, duplikat, dan tipe data."),
            ("03", "Studio Pembersihan", "Pembersihan dan konversi data."),
            ("04", "Laporan Excel", "Dokumentasi hasil pemeriksaan dalam Excel."),
        ]
        for col, (num, title, desc) in zip(module_cols, modules):
            with col:
                st.markdown(
                    f"""
                    <div class="sheet-card">
                        <div class="n">{num}</div>
                        <div class="t">{title}</div>
                        <div class="d">{desc}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

    else:
        missing_mask = get_missing_mask(df)
        total_missing = int(missing_mask.to_numpy().sum())
        total_cells = df.shape[0] * df.shape[1]
        missing_pct = total_missing / total_cells * 100 if total_cells else 0

        duplicate_mask, _ = check_duplicates(df, "Seluruh kolom", [])
        duplicate_count = int(duplicate_mask.sum())
        duplicate_pct = duplicate_count / len(df) * 100

        type_table, _ = check_data_types(df)
        type_errors = int(type_table["Jumlah Nilai Tidak Sesuai"].sum())

        # ---- KPI ----
        kpis = st.columns(5, gap="small")

        with kpis[0]:
            kpi_card(
                "Total Baris", f"{len(df):,}",
                "Jumlah baris data", "table_rows"
            )
        with kpis[1]:
            kpi_card(
                "Total Kolom", f"{len(df.columns):,}",
                "Jumlah variabel", "view_column"
            )
        with kpis[2]:
            kpi_card(
                "Missing Values", f"{total_missing:,}",
                f'{badge(f"{missing_pct:.2f}%", "warn" if total_missing else "ok")} sel kosong',
                "indeterminate_check_box",
                "warn" if total_missing else "ok"
            )
        with kpis[3]:
            kpi_card(
                "Data Duplikat", f"{duplicate_count:,}",
                f'{badge(f"{duplicate_pct:.2f}%", "warn" if duplicate_count else "ok")} baris terindikasi',
                "content_copy",
                "warn" if duplicate_count else "ok"
            )
        with kpis[4]:
            kpi_card(
                "Ketidaksesuaian Tipe", f"{type_errors:,}",
                f'{badge("Perlu Cek" if type_errors else "Sesuai", "err" if type_errors else "ok")} nilai tidak sesuai',
                "error" if type_errors else "check_circle",
                "err" if type_errors else "ok"
            )

        gap()

        # ---- STEPPER ----
        cleaned_now = st.session_state.get("cleaned_data")
        cleaned_changed = (
            cleaned_now is not None
            and st.session_state.get("clean_file_key") == file_key
            and (
                cleaned_now.shape != df.shape
                or not cleaned_now.equals(df)
            )
        )
        issue_total = total_missing + duplicate_count + type_errors

        s1, s2, s3 = st.columns(3, gap="small")
        with s1:
            step_card("1", "Pemuatan & Profiling", "Dataset Dimuat", "done", "check")
        with s2:
            if issue_total:
                step_card(
                    "2", "Validasi Aturan",
                    f"{issue_total:,} temuan perlu ditinjau", "warn", "troubleshoot"
                )
            else:
                step_card("2", "Validasi Aturan", "Tidak Ada Temuan", "done", "check")
        with s3:
            if cleaned_changed:
                step_card(
                    "3", "Remediation",
                    f"Data clean: {len(cleaned_now):,} baris", "done", "cleaning_services"
                )
            else:
                step_card(
                    "3", "Remediation", "Pembersihan & Ekspor Studio",
                    "idle", "cleaning_services"
                )

        gap()

        # ---- SKEMA + PROFIL MISSING ----
        schema_col, side_col = st.columns([1.7, 1], gap="medium")

        profile = []
        for position, column in enumerate(df.columns):
            series = df.iloc[:, position]
            miss_count = int(missing_mask.iloc[:, position].sum())
            try:
                unique_count = int(series.nunique(dropna=True))
            except TypeError:
                unique_count = int(series.astype(str).nunique())

            type_issue = int(
                type_table.loc[type_table["Nama Kolom"] == str(column), "Jumlah Nilai Tidak Sesuai"].values[0]
            )

            if type_issue:
                status = "Error"
            elif miss_count:
                status = "Peringatan"
            else:
                status = "Lolos"

            profile.append({
                "Nama Kolom": str(column),
                "Tipe Data": str(series.dtype),
                "Tipe Diharapkan": type_table.loc[position, "Tipe Data Diharapkan"],
                "Missing": miss_count,
                "% Missing": round(miss_count / len(df) * 100, 2),
                "Nilai Unik": unique_count,
                "Tipe Salah": type_issue,
                "Status": status,
            })

        profile_df = pd.DataFrame(profile)[[
            "Nama Kolom", "Status", "Tipe Data", "Tipe Diharapkan",
            "Missing", "% Missing", "Nilai Unik", "Tipe Salah"
        ]]

        with schema_col:
            with st.container(border=True):
                panel_heading(
                    "DATA SCHEMA",
                    "Gambaran Umum dan Struktur Dataset",
                    "Tipe, kelengkapan, dan status setiap kolom."
                )

                f1, f2, f3 = st.columns([1.6, 1, 1], gap="small")
                query = f1.text_input(
                    "Cari kolom", placeholder="Cari nama kolom...",
                    label_visibility="collapsed", key="schema_search"
                )
                type_filter = f2.selectbox(
                    "Filter tipe", ["Semua Tipe"] + TYPE_OPTIONS,
                    label_visibility="collapsed", key="schema_type_filter"
                )
                status_filter = f3.selectbox(
                    "Filter status",
                    ["Semua Status", "Lolos", "Peringatan", "Error"],
                    label_visibility="collapsed", key="schema_status_filter"
                )

                view = profile_df
                if query:
                    view = view[
                        view["Nama Kolom"].str.contains(
                            query, case=False, regex=False
                        )
                    ]
                if type_filter != "Semua Tipe":
                    view = view[view["Tipe Diharapkan"] == type_filter]
                if status_filter != "Semua Status":
                    view = view[view["Status"] == status_filter]

                st.caption(
                    f"Menampilkan {len(view)} dari {len(profile_df)} kolom."
                )

                st.dataframe(
                    style_status(view),
                    use_container_width=True,
                    hide_index=True,
                    height=min(420, 38 + 35 * max(len(view), 3)),
                    column_config={
                        "% Missing": st.column_config.ProgressColumn(
                            "% Missing", format="%.2f%%",
                            min_value=0, max_value=100
                        ),
                        "Missing": st.column_config.NumberColumn(format="%d"),
                        "Nilai Unik": st.column_config.NumberColumn(format="%d"),
                        "Tipe Salah": st.column_config.NumberColumn(format="%d"),
                    }
                )

        with side_col:
            with st.container(border=True):
                panel_heading(
                    "COMPLETENESS",
                    "Profil Missing Values",
                    "Kolom dengan missing values tertinggi."
                )

                top_missing = (
                    profile_df[profile_df["Missing"] > 0]
                    .sort_values("% Missing", ascending=False)
                    .head(6)
                )

                if top_missing.empty:
                    st.success("Tidak ditemukan missing values.")
                else:
                    bars = ""
                    for _, row in top_missing.iterrows():
                        pct = row["% Missing"]
                        if pct < 5:
                            bar_color = "background: var(--ok);"
                        elif pct <= 20:
                            bar_color = "background: var(--warn);"
                        else:
                            bar_color = "background: var(--err);"

                        width = max(pct, 1)
                        bars += (
                            f'<div class="bar-row">'
                            f'<div class="bar-head">'
                            f'<span class="c">{escape(row["Nama Kolom"])}</span>'
                            f'<span class="v">{pct:.2f}% ({int(row["Missing"]):,})</span></div>'
                            f'<div class="bar-track"><div class="bar-fill" '
                            f'style="width:{width}%; {bar_color}"></div></div></div>'
                        )
                    st.markdown(bars, unsafe_allow_html=True)

            gap()

            with st.container(border=True):
                panel_heading(
                    "TYPE CONSISTENCY",
                    "Ringkasan Tipe Data",
                    "Kolom dengan nilai tidak sesuai tipe."
                )
                bad_types = profile_df[profile_df["Tipe Salah"] > 0]

                if bad_types.empty:
                    st.success("Seluruh nilai sesuai tipe yang diharapkan.")
                else:
                    rows = ""
                    for _, row in bad_types.head(6).iterrows():
                        rows += (
                            f'<div class="bar-head" style="margin:.5rem 0">'
                            f'<span class="c">{escape(row["Nama Kolom"])} '
                            f'<span class="mono-tag">{escape(row["Tipe Diharapkan"])}</span></span>'
                            f'{badge(str(int(row["Tipe Salah"])) + " nilai", "err")}</div>'
                        )
                    st.markdown(rows, unsafe_allow_html=True)

        gap()

        # ---- PREVIEW ----
        with st.container(border=True):
            panel_heading(
                "DATA PREVIEW",
                "Pratinjau Dataset",
                f"Menampilkan maksimal 10 baris pertama dari {len(df):,} baris."
            )
            st.dataframe(
                df.head(10),
                use_container_width=True,
                hide_index=True
            )


# =====================================================
# PEMERIKSAAN DATA
# =====================================================

elif menu == "inspect":

    if not has_data:
        empty_state(
            "Dataset belum tersedia",
            "Unggah dataset melalui panel di atas untuk memulai pemeriksaan.",
            "rule"
        )

    else:
        tab_missing, tab_dup, tab_type = st.tabs([
            ":material/indeterminate_check_box: Missing Values",
            ":material/content_copy: Duplikat",
            ":material/rule: Tipe Data",
        ])

        # ---------- MISSING ----------
        with tab_missing:
            panel_heading(
                "01 / COMPLETENESS CHECK",
                "Deteksi Missing Values",
                "Mengidentifikasi NaN, None, teks kosong, dan teks berisi spasi."
            )

            missing_mask = get_missing_mask(df)
            missing_count = missing_mask.sum()
            total_missing = int(missing_mask.to_numpy().sum())

            total_cells = df.shape[0] * df.shape[1]
            missing_pct = (
                total_missing / total_cells * 100
                if total_cells else 0
            )
            affected_columns = int((missing_count > 0).sum())

            cols = st.columns(3)
            cols[0].metric("Total Missing Values", f"{total_missing:,}")
            cols[1].metric("Persentase Missing Values", f"{missing_pct:.2f}%")
            cols[2].metric("Kolom Terdampak", affected_columns)

            gap()

            st.markdown("**Rincian Missing Values**")
            st.dataframe(
                build_missing_table(df),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Persentase Missing Values": st.column_config.ProgressColumn(
                        "Persentase Missing Values", format="%.2f%%",
                        min_value=0, max_value=100
                    )
                }
            )

            if total_missing:
                st.warning(f"Ditemukan {total_missing:,} missing values.")
            else:
                st.success("Tidak ditemukan missing values.")

        # ---------- DUPLICATES ----------
        with tab_dup:
            panel_heading(
                "02 / DUPLICATE CHECK",
                "Deteksi Data Duplikat",
                "Memeriksa baris yang memiliki nilai identik berdasarkan kriteria."
            )

            saved_mode = st.session_state.get("saved_dup_mode", "Seluruh kolom")
            duplicate_mode = st.radio(
                "Metode pemeriksaan duplikasi:",
                ["Seluruh kolom", "Kolom tertentu"],
                index=["Seluruh kolom", "Kolom tertentu"].index(saved_mode),
                horizontal=True,
                key="duplicate_mode"
            )

            selected_columns = []

            if duplicate_mode == "Kolom tertentu":
                selected_columns = st.multiselect(
                    "Pilih kolom sebagai dasar pemeriksaan:",
                    options=list(df.columns),
                    default=[
                        c for c in st.session_state.get("saved_dup_cols", [])
                        if c in list(df.columns)
                    ],
                    key="selected_duplicate_columns"
                )

            st.session_state["saved_dup_mode"] = duplicate_mode
            st.session_state["saved_dup_cols"] = selected_columns

            duplicate_mask, duplicate_data = check_duplicates(
                df, duplicate_mode, selected_columns
            )
            duplicate_count = int(duplicate_mask.sum())

            cols = st.columns(2)
            cols[0].metric(
                "Baris Terindikasi Duplikat", f"{duplicate_count:,}"
            )
            cols[1].metric(
                "Baris Tidak Terindikasi Duplikat",
                f"{len(df) - duplicate_count:,}"
            )

            gap()

            if duplicate_mode == "Kolom tertentu" and not selected_columns:
                st.info("Pilih minimal satu kolom untuk menjalankan pemeriksaan.")
            elif duplicate_count:
                st.warning(
                    f"Ditemukan {duplicate_count:,} baris terindikasi duplikat."
                )
                st.markdown("**Detail Data Duplikat**")
                st.dataframe(
                    duplicate_data,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.success(
                    "Tidak ditemukan data duplikat berdasarkan kriteria."
                )

            st.caption(
                "Seluruh anggota kelompok duplikat ditampilkan, "
                "termasuk kemunculan pertama."
            )

        # ---------- DATA TYPES ----------
        with tab_type:
            panel_heading(
                "03 / TYPE CONSISTENCY CHECK",
                "Deteksi Tipe Data",
                "Tipe data disarankan secara otomatis dan dapat disesuaikan."
            )

            saved_types = st.session_state.setdefault("saved_types", {})

            with st.expander(
                "Atur tipe data yang diharapkan per kolom",
                expanded=False,
                icon=":material/tune:"
            ):
                grid = st.columns(3, gap="medium")

                for position, column in enumerate(df.columns):
                    default_type = detect_expected_type(df.iloc[:, position])
                    key = f"expected_type_v2_{position}"
                    current = saved_types.get(position, default_type)

                    if current not in TYPE_OPTIONS:
                        current = default_type

                    with grid[position % 3]:
                        saved_types[position] = st.selectbox(
                            f"Kolom: {column}",
                            options=TYPE_OPTIONS,
                            index=TYPE_OPTIONS.index(current),
                            key=key
                        )

            type_table, detail_table = check_data_types(df)
            total_type_errors = int(
                type_table["Jumlah Nilai Tidak Sesuai"].sum()
            )

            cols = st.columns(3)
            cols[0].metric("Kolom Diperiksa", len(type_table))
            cols[1].metric(
                "Kolom Bermasalah",
                int((type_table["Jumlah Nilai Tidak Sesuai"] > 0).sum())
            )
            cols[2].metric("Nilai Tidak Sesuai", f"{total_type_errors:,}")

            gap()

            st.markdown("**Ringkasan Pemeriksaan Tipe Data**")
            st.dataframe(
                type_table,
                use_container_width=True,
                hide_index=True
            )

            if total_type_errors:
                st.warning(
                    f"Ditemukan {total_type_errors:,} nilai yang tidak sesuai."
                )
                st.markdown("**Detail Nilai Tidak Sesuai**")
                st.dataframe(
                    detail_table,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.success(
                    "Seluruh nilai sesuai dengan tipe data yang diharapkan."
                )

            st.caption(
                "Nilai missing values tidak dihitung sebagai kesalahan tipe data."
            )


# =====================================================
# STUDIO PEMBERSIHAN
# =====================================================

elif menu == "clean":

    if not has_data:
        empty_state(
            "Dataset belum tersedia",
            "Unggah dataset melalui panel di atas untuk memulai pembersihan.",
            "auto_fix_high"
        )

    else:
        if "cleaned_data" not in st.session_state:
            st.session_state.cleaned_data = df.copy()

        show_flash()

        bar_info, bar_reset = st.columns([3, 1], gap="medium", vertical_alignment="center")

        with bar_info:
            st.markdown(
                '<div class="panel-title">Saluran Pembersihan</div>'
                '<div class="panel-sub" style="margin:0">'
                'Terapkan transformasi pada salinan dataset. '
                'Data asli tidak diubah.</div>',
                unsafe_allow_html=True
            )

        with bar_reset:
            if st.button(
                "Reset ke Data Asli",
                icon=":material/refresh:",
                use_container_width=True,
                type="secondary"
            ):
                st.session_state.cleaned_data = df.copy()
                st.rerun()

        cleaned = st.session_state.cleaned_data.copy()

        cols = st.columns(3)
        cols[0].metric("Baris Data Asli", f"{len(df):,}")
        cols[1].metric(
            "Baris Data Clean", f"{len(cleaned):,}",
            delta=f"{len(cleaned) - len(df):,}" if len(cleaned) != len(df) else None
        )
        cols[2].metric("Jumlah Kolom", f"{len(cleaned.columns):,}")

        gap()

        # BARIS 1 STUDIO PEMBERSIHAN (3 MODUL)
        c_missing, c_impute, c_dup = st.columns(3, gap="medium")

        # ---- HAPUS MISSING ----
        with c_missing:
            with st.container(border=True):
                panel_heading(
                    "01 / MISSING VALUES",
                    "Hapus Missing Values",
                    "Hapus baris dengan missing values pada kolom terpilih."
                )

                missing_columns = st.multiselect(
                    "Kolom sasaran:",
                    options=list(cleaned.columns),
                    key="clean_missing_columns"
                )

                if missing_columns:
                    mask = get_missing_mask(cleaned[missing_columns])
                    affected_rows = int(mask.any(axis=1).sum())
                    st.caption(f"ℹ️ **Dampak:** Akan menghapus {affected_rows:,} baris.")

                if st.button(
                    "Hapus Baris dengan Missing Values",
                    key="btn_clean_missing",
                    use_container_width=True
                ):
                    if not missing_columns:
                        st.warning("Pilih minimal satu kolom terlebih dahulu.")
                    else:
                        before = len(cleaned)
                        cleaned = clean_missing_values(cleaned, missing_columns)
                        st.session_state.cleaned_data = cleaned
                        removed = before - len(cleaned)
                        flash(f"Berhasil menghapus {removed:,} baris.")
                        st.rerun()

        # ---- IMPUTASI NILAI KOSONG ----
        with c_impute:
            with st.container(border=True):
                panel_heading(
                    "02 / IMPUTASI",
                    "Isi Nilai Kosong",
                    "Ganti sel kosong dengan nilai statistik atau kustom."
                )

                impute_col = st.selectbox(
                    "Kolom sasaran:",
                    options=list(cleaned.columns),
                    key="clean_impute_column"
                )

                impute_strategy = st.selectbox(
                    "Strategi imputasi:",
                    ["Mean (Rata-rata)", "Median (Nilai Tengah)", "Modus (Nilai Sering Muncul)", "Nilai Kustom"],
                    key="clean_impute_strategy"
                )

                custom_val = ""
                if impute_strategy == "Nilai Kustom":
                    custom_val = st.text_input("Nilai pengganti:", placeholder="Tuliskan nilai...", key="impute_custom_val")

                if st.button(
                    "Terapkan Imputasi",
                    key="btn_impute",
                    use_container_width=True
                ):
                    cleaned = impute_missing_values(cleaned, impute_col, impute_strategy, custom_val)
                    st.session_state.cleaned_data = cleaned
                    flash(f"Berhasil menerapkan imputasi pada kolom {impute_col}.")
                    st.rerun()

        # ---- DUPLICATES ----
        with c_dup:
            with st.container(border=True):
                panel_heading(
                    "03 / DUPLIKASI",
                    "Hapus Data Duplikat",
                    "Pertahankan kemunculan pertama, hapus sisanya."
                )

                duplicate_mode_clean = st.radio(
                    "Dasar pemeriksaan duplikasi:",
                    ["Seluruh kolom", "Kolom tertentu"],
                    horizontal=True,
                    key="clean_duplicate_mode"
                )

                selected_duplicate_clean = []
                if duplicate_mode_clean == "Kolom tertentu":
                    selected_duplicate_clean = st.multiselect(
                        "Kolom pemeriksaan:",
                        options=list(cleaned.columns),
                        key="clean_duplicate_columns"
                    )

                if st.button(
                    "Hapus Data Duplikat",
                    key="btn_clean_duplicates",
                    use_container_width=True
                ):
                    if duplicate_mode_clean == "Kolom tertentu" and not selected_duplicate_clean:
                        st.warning("Pilih minimal satu kolom.")
                    else:
                        before = len(cleaned)
                        cleaned = clean_duplicates(cleaned, duplicate_mode_clean, selected_duplicate_clean)
                        st.session_state.cleaned_data = cleaned
                        removed = before - len(cleaned)
                        flash(f"Berhasil menghapus {removed:,} baris duplikat.")
                        st.rerun()

        gap()

        # BARIS 2 STUDIO PEMBERSIHAN (2 MODUL)
        c_text, c_type = st.columns(2, gap="medium")

        # ---- NORMALISASI TEKS & TRIM ----
        with c_text:
            with st.container(border=True):
                panel_heading(
                    "04 / TEXT NORMALIZATION",
                    "Penanganan Teks & Spasi",
                    "Bersihkan spasi liar (trim) dan sesuaikan format huruf."
                )

                norm_col = st.selectbox(
                    "Kolom teks:",
                    options=list(cleaned.columns),
                    key="clean_norm_column"
                )

                norm_action = st.selectbox(
                    "Tindakan transformasi:",
                    ["Trim Whitespace", "UPPERCASE", "lowercase", "Title Case"],
                    key="clean_norm_action"
                )

                if st.button(
                    "Terapkan Normalisasi Teks",
                    key="btn_normalize_text",
                    use_container_width=True
                ):
                    cleaned = normalize_text_column(cleaned, norm_col, norm_action)
                    st.session_state.cleaned_data = cleaned
                    flash(f"Berhasil menerapkan {norm_action} pada kolom {norm_col}.")
                    st.rerun()

        # ---- KONVERSI TIPE ----
        with c_type:
            with st.container(border=True):
                panel_heading(
                    "05 / TIPE DATA",
                    "Ubah Tipe Data",
                    "Konversi dengan validasi nilai sebelum perubahan."
                )

                selected_column = st.selectbox(
                    "Kolom asal:",
                    options=list(cleaned.columns),
                    key="clean_type_column"
                )

                target_type = st.selectbox(
                    "Tipe data tujuan:",
                    TYPE_OPTIONS,
                    key="clean_target_type"
                )

                if st.button(
                    "Terapkan Perubahan Tipe Data",
                    key="btn_convert_type",
                    use_container_width=True
                ):
                    converted, failed_count = convert_column_type(
                        cleaned,
                        selected_column,
                        target_type
                    )

                    if converted is None:
                        st.error(
                            f"Konversi dibatalkan. Terdapat {failed_count:,} "
                            "nilai yang tidak dapat dikonversi. Data tidak diubah."
                        )
                    else:
                        st.session_state.cleaned_data = converted
                        flash(
                            f"Tipe data kolom {selected_column} berhasil "
                            f"diubah menjadi {target_type}."
                        )
                        st.rerun()

        gap()

        # ---- OUTPUT ----
        cleaned = st.session_state.cleaned_data.copy()

        with st.container(border=True):
            panel_heading(
                "06 / OUTPUT",
                "Data Clean Siap Diunduh",
                "Periksa hasil pembersihan sebelum mengunduh dataset."
            )

            st.markdown(
                f"""
                <div class="ready-strip">
                    {icon("cloud_done")}
                    <div>
                        <div class="t">Dataset Bersih</div>
                        <div class="d">{len(cleaned):,} baris · {len(cleaned.columns):,} kolom</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            gap()

            csv_file = cleaned.to_csv(index=False).encode("utf-8-sig")
            excel_file = generate_clean_excel(cleaned)

            d1, d2 = st.columns(2, gap="medium")

            with d1:
                st.download_button(
                    "Download CSV",
                    data=csv_file,
                    file_name="DATA_CHECKER_Clean.csv",
                    mime="text/csv",
                    icon=":material/table_view:",
                    use_container_width=True
                )

            with d2:
                st.download_button(
                    "Download Excel",
                    data=excel_file,
                    file_name="DATA_CHECKER_Clean.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    use_container_width=True
                )

            gap()

            st.markdown("**Preview Data Clean**")
            st.dataframe(
                cleaned,
                use_container_width=True,
                hide_index=True
            )

            st.caption(
                "Data asli tidak berubah. Hasil pembersihan hanya "
                "diterapkan pada salinan dataset."
            )

        gap()

        # ---- DATA BERMASALAH (untuk konfirmasi OPD) ----
        problem_mode = st.session_state.get("saved_dup_mode", "Seluruh kolom")
        problem_cols = [
            c for c in st.session_state.get("saved_dup_cols", [])
            if c in list(df.columns)
        ]
        problem_df, problem_issues, problem_raw = build_problem_data(
            df, problem_mode, problem_cols
        )

        with st.container(border=True):
            panel_heading(
                "07 / DATA BERMASALAH",
                "Data Bermasalah untuk Konfirmasi OPD",
                "Baris yang tidak boleh dihapus sepihak. Unduh dan kirim ke OPD terkait."
            )

            if problem_df.empty:
                st.markdown(
                    f'''
                    <div class="ready-strip">
                        {icon("check_circle")}
                        <div>
                            <div class="t">Tidak ada data bermasalah</div>
                            <div class="d">Tidak ditemukan missing values, duplikat, atau tipe tidak sesuai</div>
                        </div>
                    </div>
                    ''',
                    unsafe_allow_html=True
                )
            else:
                n_missing = int((problem_issues["Jenis Masalah"] == "Missing Value").sum())
                n_type = int((problem_issues["Jenis Masalah"] == "Tipe Data Tidak Sesuai").sum())
                n_dup = int((problem_issues["Jenis Masalah"] == "Data Duplikat").sum())

                st.markdown(
                    f'''
                    <div class="ready-strip" style="background:var(--warn-bg); border-color:var(--warn-bd)">
                        {icon("report")}
                        <div>
                            <div class="t">{len(problem_df):,} baris perlu konfirmasi</div>
                            <div class="d">{n_missing:,} missing · {n_type:,} tipe tidak sesuai · {n_dup:,} baris duplikat</div>
                        </div>
                    </div>
                    ''',
                    unsafe_allow_html=True
                )

                gap()

                problem_csv = problem_df.to_csv(index=False).encode("utf-8-sig")
                problem_excel = generate_problem_excel(
                    problem_df, problem_issues, problem_raw
                )

                p1, p2 = st.columns(2, gap="medium")

                with p1:
                    st.download_button(
                        "Download CSV Bermasalah",
                        data=problem_csv,
                        file_name="DATA_CHECKER_Bermasalah.csv",
                        mime="text/csv",
                        icon=":material/table_view:",
                        use_container_width=True
                    )

                with p2:
                    st.download_button(
                        "Download Excel Bermasalah",
                        data=problem_excel,
                        file_name="DATA_CHECKER_Bermasalah.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        icon=":material/download:",
                        use_container_width=True
                    )

                gap()

                st.markdown("**Preview Data Bermasalah**")
                st.dataframe(
                    problem_df,
                    use_container_width=True,
                    hide_index=True
                )

                st.caption(
                    "Dihitung dari data asli, jadi tetap lengkap walau baris sudah "
                    "dihapus di Data Clean. Excel berisi sheet Data Bermasalah (sel "
                    "bermasalah diwarnai, ada kolom Konfirmasi OPD) dan sheet "
                    "Rincian Masalah. Nomor Baris sama dengan nomor baris di file asli."
                )


# =====================================================
# LAPORAN
# =====================================================

elif menu == "report":

    if not has_data:
        empty_state(
            "Dataset belum tersedia",
            "Unggah dataset melalui panel di atas untuk membuat laporan.",
            "description"
        )

    else:
        duplicate_mode = st.session_state.get("saved_dup_mode", "Seluruh kolom")
        selected_columns = [
            c for c in st.session_state.get("saved_dup_cols", [])
            if c in list(df.columns)
        ]

        type_table, detail_table = check_data_types(df)

        missing_mask = get_missing_mask(df)
        total_missing = int(missing_mask.to_numpy().sum())

        total_cells = df.shape[0] * df.shape[1]
        missing_pct = (
            total_missing / total_cells * 100
            if total_cells else 0
        )

        duplicate_mask, _ = check_duplicates(
            df, duplicate_mode, selected_columns
        )
        duplicate_count = int(duplicate_mask.sum())

        total_type_errors = int(
            type_table["Jumlah Nilai Tidak Sesuai"].sum()
        )

        # ---- KPI ----
        k = st.columns(5, gap="small")
        with k[0]:
            kpi_card("Jumlah Baris", f"{len(df):,}", "Dataset yang diunggah", "table_rows")
        with k[1]:
            kpi_card(
                "Missing Values", f"{total_missing:,}", "Sel kosong",
                "indeterminate_check_box", "warn" if total_missing else "ok"
            )
        with k[2]:
            kpi_card(
                "Data Duplikat", f"{duplicate_count:,}", "Terindikasi duplikat",
                "content_copy", "warn" if duplicate_count else "ok"
            )
        with k[3]:
            kpi_card(
                "Missing Values (%)", f"{missing_pct:.2f}%", "Dari total sel",
                "percent", "warn" if total_missing else "ok"
            )
        with k[4]:
            kpi_card(
                "Kesalahan Tipe", f"{total_type_errors:,}", "Nilai tidak sesuai",
                "error" if total_type_errors else "check_circle",
                "err" if total_type_errors else "ok"
            )

        gap()

        left, right = st.columns([1.8, 1], gap="medium")

        with left:
            with st.container(border=True):
                panel_heading(
                    "EXCEL REPORT STRUCTURE",
                    "Isi Laporan Excel",
                    "Laporan terdiri dari enam sheet."
                )

                report_sheets = [
                    ("01", "Ringkasan", "Informasi umum dan hasil pemeriksaan dataset."),
                    ("02", "Missing Values", "Rincian data kosong pada setiap kolom."),
                    ("03", "Data Duplikat", "Baris yang terindikasi duplikat."),
                    ("04", "Tipe Data", "Hasil pemeriksaan kesesuaian tipe data."),
                    ("05", "Detail Tipe Data", "Rincian nilai yang tidak sesuai."),
                    ("06", "Data Asli", "Salinan dataset yang diunggah."),
                ]

                for row_start in range(0, 6, 3):
                    sheet_cols = st.columns(3, gap="small")
                    for col, (num, title, desc) in zip(
                        sheet_cols, report_sheets[row_start:row_start + 3]
                    ):
                        with col:
                            st.markdown(
                                f"""
                                <div class="sheet-card">
                                    <div class="n">SHEET {num}</div>
                                    <div class="t">{title}</div>
                                    <div class="d">{desc}</div>
                                </div>
                                """,
                                unsafe_allow_html=True
                            )
                    gap()

        with right:
            with st.container(border=True):
                panel_heading(
                    "EXPORT",
                    "Unduh Laporan",
                    "Dibuat dari dataset dan pengaturan pemeriksaan terpilih."
                )

                if duplicate_mode == "Kolom tertentu" and selected_columns:
                    method = "Kolom tertentu: " + ", ".join(selected_columns)
                elif duplicate_mode == "Kolom tertentu":
                    method = "Kolom tertentu: belum dipilih"
                else:
                    method = "Seluruh kolom"

                st.markdown(
                    f"""
                    <div class="bar-head" style="margin:.4rem 0">
                        <span class="c">Berkas</span>
                        <span class="v">{escape(uploaded_file.name)}</span>
                    </div>
                    <div class="bar-head" style="margin:.4rem 0">
                        <span class="c">Metode duplikasi</span>
                        <span class="v">{escape(method)}</span>
                    </div>
                    <div class="bar-head" style="margin:.4rem 0 1rem">
                        <span class="c">Format</span>
                        <span class="v">.xlsx · 6 sheet</span>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                excel_file = generate_excel_report(
                    data=df,
                    file_name=uploaded_file.name,
                    duplicate_mode=duplicate_mode,
                    selected_columns=selected_columns,
                    type_table=type_table,
                    detail_table=detail_table
                )

                st.download_button(
                    label="Download Laporan Excel",
                    data=excel_file,
                    file_name="Laporan_DATA_CHECKER.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    use_container_width=True
                )

                st.caption(
                    "Metode duplikasi dan tipe data mengikuti pengaturan "
                    "di menu Pemeriksaan Data."
                )