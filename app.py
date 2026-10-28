import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.oauth2 import id_token
from google.auth.transport import requests as google_requests

st.set_page_config(page_title="Portal TND BPKP Papua", page_icon="📜", layout="wide")

# 1. SETUP WHITELIST EMAILS
# Ganti dan tambahkan email tim yang diizinkan mengakses aplikasi ini
WHITELIST_EMAILS = [
    "mi.ramadhani123@gmail.com",  
    "afri.rodhiatul@bpkp.go.id"
]

# 2. KONFIGURASI OAUTH
client_config = {
    "web": {
        "client_id": st.secrets["google_oauth"]["client_id"],
        "client_secret": st.secrets["google_oauth"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
    }
}

flow = Flow.from_client_config(
    client_config,
    scopes=["openid", "https://www.googleapis.com/auth/userinfo.email"],
    redirect_uri=st.secrets["google_oauth"]["redirect_uri"]
)

# 3. INISIALISASI SESSION STATE
if "user_email" not in st.session_state:
    st.session_state.user_email = None

# 4. TANGKAP KODE REDIRECT DARI GOOGLE
if "code" in st.query_params and st.session_state.user_email is None:
    code = st.query_params["code"]
    try:
        flow.fetch_token(code=code)
        credentials = flow.credentials

        # Verifikasi token dan ekstrak email
        request = google_requests.Request()
        id_info = id_token.verify_oauth2_token(
            credentials.id_token, request, st.secrets["google_oauth"]["client_id"]
        )
        st.session_state.user_email = id_info.get("email")

        # Bersihkan URL agar rapi
        st.query_params.clear()
        st.rerun()
    except Exception as e:
        st.error(f"Gagal memverifikasi login: {e}")

# 5. HALAMAN LOGIN & VALIDASI WHITELIST
if st.session_state.user_email is None:
    # Tampilan jika belum login
    st.title("Portal TND BPKP Papua")
    st.warning("Aplikasi ini dilindungi. Silakan login menggunakan email BPKP yang terdaftar.")

    auth_url, _ = flow.authorization_url(prompt='consent')
    st.link_button("Login with Google", auth_url)

else:
    # Jika sudah login, cek whitelist
    email = st.session_state.user_email
    if email in WHITELIST_EMAILS:
        # --- USER DIIZINKAN MASUK ---
        col1, col2 = st.columns([0.8, 0.2])
        with col1:
            st.success(f"Login berhasil. Selamat datang, {email}")
        with col2:
            if st.button("Logout"):
                st.session_state.user_email = None
                st.rerun()

        st.divider()

        # PANGGIL MENU NAVIGASI UTAMA LO DI SINI
        # (Pastikan variabel 'pages' lo disiapkan sebelum ini atau di dalam blok ini sesuai kode lo sebelumnya)

        import tnd_generator # Sesuaikan dengan nama modul lo

        pages = {
            "Navigasi TND": [
                st.Page("tnd_generator.py", title="Surat Penegasan (SP2)", icon="✍️"),
                # Tambahkan halaman lain jika ada
            ]
        }
        st.navigation(pages).run()

    else:
        # --- USER DIBLOKIR ---
        st.error(f"AKSES DITOLAK: Email '{email}' tidak terdaftar dalam whitelist sistem.")
        if st.button("Kembali / Ganti Akun"):
            st.session_state.user_email = None
            st.rerun()
