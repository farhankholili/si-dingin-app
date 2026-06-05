import streamlit as st
import pandas as pd
import os
from database import db, Order, DetailOrder, User, Layanan
from test_input import buat_order_baru

# Konfigurasi Halaman Utama
st.set_page_config(page_title="Si Dingin - Sistem Integrasi", layout="wide")

# Hubungkan Database
db.connect(reuse_if_open=True)

# Folder untuk simpan logo
if not os.path.exists("assets"):
    os.makedirs("assets")

# Default Data Toko
if "nama_usaha" not in st.session_state:
    st.session_state.nama_usaha = "SI DINGIN COOLING SYSTEM"
if "alamat_usaha" not in st.session_state:
    st.session_state.alamat_usaha = "Jl. Teknisi AC No. 1, Kota Pempek"
if "kontak_usaha" not in st.session_state:
    st.session_state.kontak_usaha = "0812-7000-xxxx"

# Session State untuk Login & Keranjang
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "user_fullname" not in st.session_state:
    st.session_state.user_fullname = ""
if "user_id" not in st.session_state:
    st.session_state.user_id = None
if "keranjang_kerja" not in st.session_state:
    st.session_state.keranjang_kerja = []

# ==================================================
# HALAMAN LOGIN (SISTEM KEAMANAN PINTU UTAMA)
# ==================================================
def halaman_login():
    st.markdown("<h1 style='text-align: center; color: #0E7490;'>❄️ SISTEM OPERASI SI DINGIN</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Silakan masuk dengan akun Anda untuk melanjutkan</p>", unsafe_allow_html=True)
    
    kol1, kol2, kol3 = st.columns([1, 2, 1])
    with kol2:
        with st.container(border=True):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            
            if st.button("🚪 MASUK KE SISTEM", type="primary", use_container_width=True):
                # Proteksi khusus owner bypass cepat atau via db
                if username == "farhan" and password == "owner123":
                    st.session_state.authenticated = True
                    st.session_state.user_role = "owner"
                    st.session_state.user_fullname = "Farhan Kholili"
                    st.rerun()
                else:
                    # Cek ke database untuk teknisi
                    user_match = User.select().where(User.username == username, User.password == password).first()
                    if user_match:
                        st.session_state.authenticated = True
                        st.session_state.user_role = user_match.role
                        st.session_state.user_fullname = user_match.nama_lengkap
                        st.session_state.user_id = user_match.id
                        st.rerun()
                    else:
                        st.error("Username atau Password salah, Bos!")

# ==================================================
# VIEW 1: INTERFACE KHUSUS TEKNISI LAPANGAN
# ==================================================
def tampilan_teknisi():
    st.sidebar.markdown(f"### 🧑‍🔧 Teknisi: **{st.session_state.user_fullname}**")
    if st.sidebar.button("🔒 Keluar Aplikasi", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
        
    st.title("❄️ SI DINGIN - LAYANAN TEKNISI LAPANGAN")
    
    semua_layanan = Layanan.select()
    kol1, kol2 = st.columns([1, 1])

    with kol1:
        st.subheader("📋 1. Data Pelanggan")
        nama = st.text_input("Nama Pelanggan")
        alamat = st.text_area("Alamat Rumah")
        telp = st.text_input("No. WhatsApp")

        st.divider()
        st.subheader("🛠️ 2. Input Item Pekerjaan")
        opsi = {l.id: f"{l.nama_layanan} (Rp {l.harga:,})" for l in semua_layanan}
        layanan_pilihan = st.selectbox("Pilih Jenis Pekerjaan", options=list(opsi.keys()), format_func=lambda x: opsi[x])
        ruangan = st.text_input("Lokasi / Ruangan (Misal: Kamar Utama)")
        qty = st.number_input("Jumlah (Qty)", min_value=1, value=1, step=1)

        if st.button("➕ Tambah Pekerjaan", use_container_width=True):
            if ruangan == "": st.warning("Isi lokasi ruangan dulu!")
            else:
                layanan_obj = [l for l in semua_layanan if l.id == layanan_pilihan][0]
                st.session_state.keranjang_kerja.append({
                    "id_layanan": layanan_pilihan, "nama_layanan": layanan_obj.nama_layanan,
                    "harga": layanan_obj.harga, "jumlah": qty, "lokasi_ruang": ruangan
                })
                st.rerun()

    with kol2:
        st.subheader("🛒 3. Keranjang Kerja")
        if not st.session_state.keranjang_kerja:
            st.info("Keranjang kosong.")
        else:
            total_nota = 0
            for idx, item in enumerate(st.session_state.keranjang_kerja):
                subtotal = item["harga"] * item["jumlah"]
                total_nota += subtotal
                st.markdown(f"**{idx+1}. {item['nama_layanan']}** (x{item['jumlah']})")
                st.caption(f"📌 Ruang: {item['lokasi_ruang']} | Subtotal: Rp {subtotal:,}")
            
            st.metric(label="TOTAL NOTA", value=f"Rp {total_nota:,}")
            if st.button("🗑️ Reset Keranjang"):
                st.session_state.keranjang_kerja = []
                st.rerun()

            st.divider()
            if st.button("🔥 PROSES & BUAT NOTA INVOICE", type="primary", use_container_width=True):
                if not nama: st.error("Nama pelanggan wajib diisi!")
                else:
                    id_tek = st.session_state.user_id if st.session_state.user_id else 2
                    order = buat_order_baru(id_teknisi=id_tek, nama_pelanggan=nama, alamat_pelanggan=alamat, no_telp=telp, list_pekerjaan=st.session_state.keranjang_kerja)
                    if order:
                        st.balloons()
                        st.success("🚀 INVOICE SUKSES DISIMPAN!")
                        
                        # --- INVOICE STYLE NOTA PREMIUM KASIR THERMAL (HTML/CSS) ---
                        logo_html = ""
                        if os.path.exists("assets/logo.png"):
                            logo_html = f"<center><img src='app/static/assets/logo.png' width='100'></center>"
                            
                        html_nota = f"""
                        <div style='background-color:#fff; padding:20px; border:2px dashed #000; color:#000; font-family:monospace;'>
                            {logo_html}
                            <h3 style='text-align:center; margin:0;'>{st.session_state.nama_usaha}</h3>
                            <p style='text-align:center; font-size:12px; margin:5px 0;'>📍 {st.session_state.alamat_usaha}<br>📞 WA: {st.session_state.kontak_usaha}</p>
                            <hr style='border-top:1px dashed #000;'>
                            <p style='font-size:13px; margin:3px 0;'><b>Inv:</b> {order.no_invoice}</p>
                            <p style='font-size:13px; margin:3px 0;'><b>Tgl:</b> {order.tanggal_kerja}</p>
                            <p style='font-size:13px; margin:3px 0;'><b>Cust:</b> {order.nama_pelanggan}</p>
                            <p style='font-size:13px; margin:3px 0;'><b>Teknisi:</b> {st.session_state.user_fullname}</p>
                            <hr style='border-top:1px dashed #000;'>
                            <table style='width:100%; font-size:13px;'>
                        """
                        for item in st.session_state.keranjang_kerja:
                            html_nota += f"<tr><td colspan='2'><b>{item['nama_layanan']}</b> ({item['lokasi_ruang']})</td></tr>"
                            html_nota += f"<tr><td>{item['jumlah']} x Rp {item['harga']:,}</td><td style='text-align:right;'>Rp {item['harga']*item['jumlah']:,}</td></tr>"
                        
                        html_nota += f"""
                            </table>
                            <hr style='border-top:1px dashed #000;'>
                            <h3 style='text-align:right; margin:5px 0;'>TOTAL: Rp {order.total_bayar:,}</h3>
                            <p style='text-align:center; font-size:11px; margin-top:15px;'>* STATUS: BELUM BAYAR (MENUNGGU VALIDASI OWNER) *<br>Terima kasih atas kepercayaan Anda!</p>
                        </div>
                        """
                        st.markdown(html_nota, unsafe_allow_html=True)
                        st.session_state.keranjang_kerja = []

# ==================================================
# VIEW 2: INTERFACE KHUSUS OWNER (PAK FARHAN)
# ==================================================
def tampilan_owner():
    st.sidebar.markdown(f"### 👑 Tingkat Akses: **{st.session_state.user_fullname}**")
    if st.sidebar.button("🔒 Keluar Aplikasi", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

    tab_keuangan, tab_layanan, tab_teknisi, tab_toko = st.tabs([
        "📈 Keuangan & Invoice", "⚙️ Kelola Harga Layanan", "🧑‍🔧 Kelola Data Teknisi", "🏨 Pengaturan Toko & Logo"
    ])
    
    # --- LOGIKA KEUANGAN ---
    with tab_keuangan:
        query_order = (Order.select(Order, User).join(User, on=(Order.teknisi_id == User.id)).order_by(Order.tanggal_kerja.desc()))
        orders_data = []
        for o in query_order:
            orders_data.append({
                "ID": o.id, "No Invoice": o.no_invoice, "Tanggal": o.tanggal_kerja,
                "Pelanggan": o.nama_pelanggan, "Alamat": o.alamat_pelanggan,
                "Teknisi": o.teknisi.nama_lengkap if o.teknisi else "Tanpa Teknisi",
                "Total Bayar": o.total_bayar, "Status": getattr(o, "status_bayar", "Belum Dibayar")
            })
        df = pd.DataFrame(orders_data)
        if df.empty: st.info("Belum ada transaksi.")
        else:
            t_omset = df[df["Status"] == "Lunas"]["Total Bayar"].sum()
            t_piutang = df[df["Status"] == "Belum Dibayar"]["Total Bayar"].sum()
            st.columns(3)[0].metric("💰 TOTAL OMSET MASUK", f"Rp {t_omset:,}")
            st.columns(3)[1].metric("⚠️ TOTAL PIUTANG", f"Rp {t_piutang:,}")
            st.columns(3)[2].metric("📦 TOTAL NOTA", f"{len(df)} Nota")
            st.divider()
            st.dataframe(df.drop(columns=["ID"]), use_container_width=True, hide_index=True)
            
            st.subheader("⚡ Aksi Pelunasan Nota")
            df_blm = df[df["Status"] == "Belum Dibayar"]
            if not df_blm.empty:
                opsi_nota = {row["ID"]: f"{row['No Invoice']} - {row['Pelanggan']} (Rp {row['Total Bayar']:,})" for _, row in df_blm.iterrows()}
                nota_sel = st.selectbox("Pilih Invoice:", options=list(opsi_nota.keys()), format_func=lambda x: opsi_nota[x])
                if st.button("✅ SET JADI LUNAS", type="primary"):
                    ord_obj = Order.get_by_id(nota_sel)
                    setattr(ord_obj, "status_bayar", "Lunas")
                    ord_obj.save()
                    st.rerun()

    # --- KELOLA HARGA ---
    with tab_layanan:
        st.subheader("🛠️ Master Harga Jasa Jual")
        semua_layanan = list(Layanan.select())
        df_lay = pd.DataFrame([{"ID": l.id, "Nama Layanan": l.nama_layanan, "Harga Jual": f"Rp {l.harga:,}"} for l in semua_layanan])
        st.dataframe(df_lay, use_container_width=True, hide_index=True)
        st.divider()
        k_t, k_e = st.columns(2)
        with k_t:
            n_jasa = st.text_input("Nama Layanan Baru")
            h_jasa = st.number_input("Harga Tarif (Rp)", min_value=0, step=5000, value=65000)
            if st.button("Simpan Jasa Baru"):
                if n_jasa: Layanan.create(nama_layanan=n_jasa, harga=h_jasa); st.rerun()
        with k_e:
            if semua_layanan:
                opsi_edit = {l.id: f"{l.nama_layanan}" for l in semua_layanan}
                j_pilih = st.selectbox("Pilih Jasa Mau Diedit:", options=list(opsi_edit.keys()), format_func=lambda x: opsi_edit[x])
                hrg_baru = st.number_input("Harga Baru", min_value=0, step=5000, value=75000, key="edit_h_b")
                if st.button("Update Tarif Harga"):
                    l_obj = Layanan.get_by_id(j_pilih)
                    l_obj.harga = hrg_baru
                    l_obj.save()
                    st.rerun()

    # --- KELOLA DATA TEKNISI (SUDAH BISA INTIP PASSWORD) ---
    with tab_teknisi:
        st.subheader("🧑‍🔧 Akun Tim Teknisi & Password")
        
        # Mengambil data teknisi termasuk password dari database
        semua_tek = list(User.select().where(User.role == "teknisi"))
        
        # SEKARANG KOLOM PASSWORD KITA MUNCULKAN DI SINI, BOS!
        df_tek = pd.DataFrame([{
            "ID": t.id, 
            "Nama Lengkap": t.nama_lengkap,
            "Username Login": t.username, 
            "🔑 Password Aktif": t.password  # <--- Ini dia kuncinya!
        } for t in semua_tek])
        
        if df_tek.empty:
            st.info("Belum ada teknisi yang terdaftar.")
        else:
            # Tampilkan tabel yang sudah ada kolom password-nya
            st.dataframe(df_tek, use_container_width=True, hide_index=True)
            
        st.divider()
        st.markdown("### ➕ Daftarkan Teknisi Baru")
        kt1, kt2, kt3 = st.columns(3)
        with kt1: tk_n = st.text_input("Nama Lengkap Karyawan")
        with kt2: tk_u = st.text_input("Username Login (Huruf Kecil, Tanpa Spasi)")
        with kt3: tk_p = st.text_input("Password Login", type="default", value="123") # type=default biar kelihatan pas ngetik
        
        if st.button("Daftarkan Teknisi", type="primary", key="btn_reg_tek"):
            if tk_n and tk_u:
                # Cek apakah username sudah dipakai orang lain
                cek_user = User.select().where(User.username == tk_u).first()
                if cek_user:
                    st.error(f"Username '{tk_u}' sudah dipakai, Bos! Ganti yang lain.")
                else:
                    User.create(username=tk_u, nama_lengkap=tk_n, password=tk_p, role="teknisi")
                    st.success(f"Siiap! Teknisi baru bernama {tk_n} berhasil didaftarkan!")
                    st.rerun()
    # --- PROFILE TOKO ---
    with tab_toko:
        st.subheader("🏨 Profil Usaha")
        st.session_state.nama_usaha = st.text_input("Nama Usaha/Bengkel AC", value=st.session_state.nama_usaha)
        st.session_state.alamat_usaha = st.text_area("Alamat Workshop", value=st.session_state.alamat_usaha)
        st.session_state.kontak_usaha = st.text_input("Nomor HP / WhatsApp Bisnis", value=st.session_state.kontak_usaha)
        if st.button("💾 Simpan Informasi Toko"): st.success("Informasi Toko Disimpan!")

# ==================================================
# LOGIKA ROUTING UTAMA APLIKASI
# ==================================================
if not st.session_state.authenticated:
    halaman_login()
else:
    if st.session_state.user_role == "owner":
        tampilan_owner()
    else:
        tampilan_teknisi()