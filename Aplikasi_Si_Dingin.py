import streamlit as st
import pandas as pd
import os
import io  # Untuk memproses file Excel di memori RAM
import openpyxl
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as ExcelImage
from database import db, Order, DetailOrder, User, Layanan
from test_input import buat_order_baru

# Konfigurasi Halaman Utama
st.set_page_config(page_title="Si Dingin - Sistem Integrasi", layout="wide")

# Hubungkan Database
db.connect(reuse_if_open=True)

# Default Data Toko di Session State
if "nama_usaha" not in st.session_state:
    st.session_state.nama_usaha = "NOTA DIGITAL JASA REPARASI & SERVICE AC"
if "alamat_usaha" not in st.session_state:
    st.session_state.alamat_usaha = "AC Dingin, Kerja Transparan, Keluarga Nyaman"
if "kontak_usaha" not in st.session_state:
    st.session_state.kontak_usaha = "0812-XXXX-XXXX"

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
if "nota_siap_download" not in st.session_state:
    st.session_state.nota_siap_download = None

# STATE TAMBAHAN: Fitur Lupa Password
if "forgot_mode" not in st.session_state:
    st.session_state.forgot_mode = False
if "otp_terkirim" not in st.session_state:
    st.session_state.otp_terkirim = None
if "otp_terverifikasi" not in st.session_state:
    st.session_state.otp_terverifikasi = False

# ==================================================
# FUNGSI ENGINERING: KIRIM OTP EMAIL
# ==================================================
def kirim_email_otp(email_tujuan, otp_code):
    """Mengirimkan email kode verifikasi OTP menggunakan server SMTP Gmail"""
    email_pengirim = "owner.sidingin@gmail.com" 
    password_aplikasi = "xxxx xxxx xxxx xxxx" # ⚠️ GANTI PAKE 16 DIGIT APP PASSWORD GOOGLE BOS!

    msg = MIMEMultipart()
    msg['From'] = email_pengirim
    msg['To'] = email_tujuan
    msg['Subject'] = "🔐 KODE OTP: Pemulihan Sistem Si Dingin"

    body = f"""
    Halo Bos Farhan,
    
    Sistem mendeteksi permintaan pemulihan akun untuk dashboard Owner Si Dingin.
    Berikut adalah kode verifikasi OTP Bos:
    
    👉 {otp_code} 👈
    
    Masukkan kode di atas ke dalam aplikasi untuk membuka akses perubahan username dan password baru.
    
    Salam Sukses,
    Sistem Integrasi Si Dingin ❄️
    """
    msg.attach(MIMEText(body, 'plain'))
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_pengirim, password_aplikasi)
        server.sendmail(email_pengirim, email_tujuan, msg.as_string())
        server.quit()
        return True
    except Exception:
        return False

# ==================================================
# FUNGSI SAKTI: CETAK NOTA PREMIUM (DINAMIS & LOGO)
# ==================================================
def buat_file_excel_nota(order_obj, list_detail_pekerjaan):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nota Digital"
    ws.views.sheetView[0].showGridLines = True
    
    font_judul = Font(name="Arial", size=16, bold=True, color="1E3A8A")
    font_sub_judul = Font(name="Arial", size=9, italic=True, color="555555")
    font_section = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_header_tabel = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_data = Font(name="Arial", size=10)
    font_garansi = Font(name="Arial", size=9, italic=True)
    
    fill_biru_tua = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
    fill_biru_tabel = PatternFill(start_color="1D4ED8", end_color="1D4ED8", fill_type="solid")
    
    border_thin = Border(
        left=Side(style='thin', color='B0B0B0'), right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'), bottom=Side(style='thin', color='B0B0B0')
    )
    
    # 1. ATUR TINGGI BARIS HEADER SUPAYA LOGO PAS
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 15
    
    # 2. MASUKKAN LOGO DI POJOK KIRI ATAS (CELL A1)
    if os.path.exists("logo.png"):
        try:
            img = ExcelImage("logo.png")
            # Set ukuran proporsional (tinggi sekitar 75-80px agar pas dengan tinggi baris header)
            img.width = 80
            img.height = 80
            ws.add_image(img, 'A1')
        except Exception:
            pass
        
    # 3. KOP SURAT DIGESER KE KOLOM B AGAR TIDAK TERTUTUP LOGO DI KOLOM A
    ws.merge_cells("B2:F2")
    ws["B2"] = st.session_state.nama_usaha.upper()
    ws["B2"].font = font_judul
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    
    ws.merge_cells("B3:F3")
    ws["B3"] = f"{st.session_state.alamat_usaha} | Telp: {st.session_state.kontak_usaha}"
    ws["B3"].font = font_sub_judul
    ws["B3"].alignment = Alignment(horizontal="left", vertical="center")
    
    # 4. DATA NOTA & PELANGGAN DIMULAI DARI BARIS 6 (TURUN BEBERAPA BARIS DARI LOGO)
    ws.merge_cells("A6:F6")
    ws["A6"] = " DATA NOTA & PELANGGAN"
    ws["A6"].font = font_section
    ws["A6"].fill = fill_biru_tua
    ws["A6"].alignment = Alignment(vertical="center")
    ws.row_dimensions[6].height = 22
    
    nama_tek = order_obj.teknisi.nama_lengkap if order_obj.teknisi else "Tanpa Teknisi"
    hp_tek = "-"
    if order_obj.teknisi:
        hp_tek = getattr(order_obj.teknisi, 'no_hp', getattr(order_obj.teknisi, 'no_telp', '-'))
    
    status_db = getattr(order_obj, "status_pembayaran", "Belum Dibayar")
    status_text = "LUNAS" if status_db == "Lunas" else "BELUM LUNAS"
    
    nomor_telepon_pelanggan = "-"
    if hasattr(order_obj, 'no_telp_pelanggan'): nomor_telepon_pelanggan = order_obj.no_telp_pelanggan
    elif hasattr(order_obj, 'no_telp'): nomor_telepon_pelanggan = order_obj.no_telp
    elif hasattr(order_obj, 'no_hp'): nomor_telepon_pelanggan = order_obj.no_hp
        
    ws["A7"] = "No. Nota / Invoice"; ws["B7"] = order_obj.no_invoice
    ws["A8"] = "Tanggal Pekerjaan"; ws["B8"] = str(order_obj.tanggal_kerja)
    ws["A9"] = "Teknisi Lapangan"; ws["B9"] = f"{nama_tek} ({hp_tek})"
    ws["A10"] = "Status Pembayaran"; ws["B10"] = status_text
    
    alamat_p_text = getattr(order_obj, 'alamat_pelanggan', '-')
        
    ws["D7"] = "Nama Pelanggan"; ws["E7"] = order_obj.nama_pelanggan
    ws["D8"] = "No. Telepon / WA"; ws["E8"] = nomor_telepon_pelanggan
    ws["D9"] = "Alamat Pelanggan"; ws["E9"] = alamat_p_text
    ws["D10"] = "Metode Pembayaran"; ws["E10"] = "Cash / Transfer"
    
    for r in range(7, 11):
        ws.row_dimensions[r].height = 18
        ws[f"A{r}"].font = font_bold; ws[f"B{r}"].font = font_data
        ws[f"D{r}"].font = font_bold; ws[f"E{r}"].font = font_data
    
    # 5. HEADER TABEL DIMULAI DARI BARIS 12
    ws.merge_cells("A12:F12")
    ws["A12"] = " DETAIL PEKERJAAN, SPAREPART & BIAYA"
    ws["A12"].font = font_section
    ws["A12"].fill = fill_biru_tua
    ws["A12"].alignment = Alignment(vertical="center")
    ws.row_dimensions[12].height = 22
    
    headers = ["No", "Deskripsi Pekerjaan / Sparepart", "Qty (Unit/Pcs)", "Harga Satuan", "Total Jasa & Bahan"]
    ws.cell(row=13, column=1, value=headers[0])
    ws.cell(row=13, column=2, value=headers[1])
    ws.cell(row=13, column=3, value=headers[2])
    ws.merge_cells("C13:D13")
    ws.cell(row=13, column=5, value=headers[3])
    ws.cell(row=13, column=6, value=headers[4])
    
    ws.row_dimensions[13].height = 25
    for col_idx in [1, 2, 3, 5, 6]:
        c = ws.cell(row=13, column=col_idx)
        c.font = font_header_tabel; c.fill = fill_biru_tabel
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = border_thin
    
    # 6. DATA ITEM TABEL DIMULAI DARI BARIS 14
    current_row = 14
    total_akhir = 0
    for idx, item in enumerate(list_detail_pekerjaan, start=1):
        ws.row_dimensions[current_row].height = 20
        sub_item = item["harga"] * item["jumlah"]
        total_akhir += sub_item
        
        ws.cell(row=current_row, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
        
        teks_ruangan = f" ({item['lokasi_ruang']})" if item['lokasi_ruang'] and item['lokasi_ruang'] != "-" else ""
        ws.cell(row=current_row, column=2, value=f"{item['nama_layanan']}{teks_ruangan}").alignment = Alignment(vertical="center")
        
        ws.cell(row=current_row, column=3, value=item["jumlah"]).alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4)
        
        c_harga = ws.cell(row=current_row, column=5, value=item["harga"])
        c_harga.number_format = '"Rp "#,##0'; c_harga.alignment = Alignment(horizontal="right", vertical="center")
        
        c_sub = ws.cell(row=current_row, column=6, value=sub_item)
        c_sub.number_format = '"Rp "#,##0'; c_sub.alignment = Alignment(horizontal="right", vertical="center")
        
        for col_idx in range(1, 7):
            ws.cell(row=current_row, column=col_idx).font = font_data
            ws.cell(row=current_row, column=col_idx).border = border_thin
        current_row += 1
    
    # TOTAL AKHIR
    ws.row_dimensions[current_row].height = 22
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
    c_tot_label = ws.cell(row=current_row, column=1, value="TOTAL PEMBAYARAN :")
    c_tot_label.font = font_bold; c_tot_label.alignment = Alignment(horizontal="right", vertical="center")
    
    c_tot_val = ws.cell(row=current_row, column=6, value=total_akhir)
    c_tot_val.font = font_bold; c_tot_val.number_format = '"Rp "#,##0'
    c_tot_val.alignment = Alignment(horizontal="right", vertical="center"); c_tot_val.border = border_thin
    
    # GARANSI
    current_row += 2
    ws.row_dimensions[current_row].height = 22
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
    
    cell_garansi_header = ws.cell(row=current_row, column=1, value=" SYARAT & KETENTUAN GARANSI REPARASI")
    cell_garansi_header.font = font_section
    cell_garansi_header.fill = fill_biru_tua
    cell_garansi_header.alignment = Alignment(vertical="center")
    
    syarat_teks = [
        "1. Garansi reparasi berlaku selama 30 HARI terhitung sejak tanggal selesai pengerjaan.",
        "2. Garansi HANYA BERLAKU untuk komponen, sparepart, atau jenis kerusakan yang sama pada nota ini.",
        "3. Nota ini dianggap SAH apabila dikirim langsung oleh pihak Manajemen/Teknisi Resmi kami dalam bentuk GAMBAR/PDF.",
        "4. Segala bentuk manipulasi, perubahan tulisan, atau penyalahgunaan nota fisik/digital di luar sistem kami",
        "   bukan merupakan tanggung jawab kami dan dapat dibawa ke jalur hukum demi kenyamanan bersama."
    ]
    for teks in syarat_teks:
        current_row += 1
        ws.row_dimensions[current_row].height = 16
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
        ws.cell(row=current_row, column=1, value=teks).font = font_garansi
        ws.cell(row=current_row, column=1).alignment = Alignment(vertical="center")
    
    current_row += 2
    ws.row_dimensions[current_row].height = 20
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
    ws.cell(row=current_row, column=1, value=f"- {st.session_state.nama_usaha} -").font = Font(name="Arial", size=10, bold=True, color="1E40AF")
    ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    
    # ATUR LEBAR KOLOM AGAR PRESISI
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 22
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ==================================================
# HALAMAN LOGIN & RECOVERY PASSWORD
# ==================================================
def halaman_login():
    st.markdown("<h1 style='text-align: center; color: #0E7490;'>❄️ SISTEM OPERASI SI DINGIN</h1>", unsafe_allow_html=True)
    kol1, col2, kol3 = st.columns([1, 2, 1])
    
    with col2:
        if st.session_state.forgot_mode:
            with st.container(border=True):
                st.subheader("🔑 Pemulihan Akun Owner via Email")
                
                if not st.session_state.otp_terverifikasi:
                    email_input = st.text_input("Masukkan Alamat Email Owner Terdaftar:")
                    
                    if st.session_state.otp_terkirim is not None:
                        otp_user = st.text_input("Masukkan 6-Digit Kode OTP dari Email:", max_chars=6)
                        if st.button("🔑 VERIFIKASI KODE OTP", use_container_width=True, type="primary"):
                            if otp_user == st.session_state.otp_terkirim:
                                st.session_state.otp_terverifikasi = True
                                st.success("Verifikasi Berhasil! Silakan atur ulang akses masuk Bos.")
                                st.rerun()
                            else:
                                st.error("Kode OTP salah atau tidak sesuai, Bos!")
                    else:
                        if st.button("📨 KIRIM KODE OTP SEKARANG", use_container_width=True, type="primary"):
                            owner_db = User.select().where(User.role == "owner").first()
                            email_di_db = getattr(owner_db, 'no_hp', '') if owner_db else ""
                            
                            if not email_input or email_input != email_di_db:
                                st.error("Maaf Bos, email tersebut tidak cocok dengan data Owner di database!")
                            else:
                                raw_otp = str(random.randint(100000, 999999))
                                st.session_state.otp_terkirim = raw_otp
                                if kirim_email_otp(email_input, raw_otp):
                                    st.success("Kode OTP berhasil dikirim! Silakan periksa Kotak Masuk Email Bos.")
                                    st.rerun()
                                else:
                                    st.error("Gagal mengirim email. Periksa koneksi internet atau setelan App Password Google.")
                else:
                    owner_db = User.select().where(User.role == "owner").first()
                    st.info(f"Username Anda saat ini di sistem: **{owner_db.username if owner_db else 'farhan'}**")
                    
                    user_baru_otp = st.text_input("Buat Username Baru:", value=owner_db.username if owner_db else "farhan")
                    pass_baru_otp = st.text_input("Buat Password Baru:", type="password")
                    
                    if st.button("💾 SIMPAN & PERBARUI AKUN", use_container_width=True, type="primary"):
                        if user_baru_otp and pass_baru_otp and owner_db:
                            User.update(username=user_baru_otp, password=pass_baru_otp).where(User.id == owner_db.id).execute()
                            st.success("Akun Owner berhasil diselamatkan! Silakan Login kembali.")
                            st.session_state.forgot_mode = False
                            st.session_state.otp_terkirim = None
                            st.session_state.otp_terverifikasi = False
                            st.rerun()
                
                if st.button("⬅️ Kembali ke Halaman Login", use_container_width=True):
                    st.session_state.forgot_mode = False
                    st.session_state.otp_terkirim = None
                    st.session_state.otp_terverifikasi = False
                    st.rerun()
        
        else:
            with st.container(border=True):
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                if st.button("🚪 MASUK KE SISTEM", type="primary", use_container_width=True):
                    if username == "farhan" and password == "owner123":
                        st.session_state.authenticated = True
                        st.session_state.user_role = "owner"
                        st.session_state.user_fullname = "Farhan Kholili"
                        owner_init = User.select().where(User.role == "owner").first()
                        if owner_init: st.session_state.user_id = owner_init.id
                        st.rerun()
                    else:
                        user_match = User.select().where(User.username == username, User.password == password).first()
                        if user_match:
                            st.session_state.authenticated = True
                            st.session_state.user_role = user_match.role
                            st.session_state.user_fullname = user_match.nama_lengkap
                            st.session_state.user_id = user_match.id
                            st.rerun()
                        else: 
                            st.error("Username atau Password salah, Bos!")
                
                st.divider()
                if st.button("🔒 Lupa Username / Password Owner?", use_container_width=True):
                    st.session_state.forgot_mode = True
                    st.rerun()

# ==================================================
# VIEW 1: INTERFACE KHUSUS TEKNISI LAPANGAN
# ==================================================
def tampilan_teknisi():
    st.sidebar.markdown(f"### 🧑‍🔧 Teknisi: **{st.session_state.user_fullname}**")
    if st.sidebar.button("🔒 Keluar Aplikasi", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
        
    tab_kerja, tab_riwayat_tek = st.tabs(["🛠️ Input Pekerjaan Lapangan", "🖨️ Riwayat & Cetak Ulang Nota Anda"])
    
    with tab_kerja:
        st.title("❄️ INPUT DETAIL NOTA BARU")
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
                if ruangan == "": 
                    st.warning("Isi lokasi ruangan dulu!")
                else:
                    layanan_obj = [l for l in semua_layanan if l.id == layanan_pilihan][0]
                    st.session_state.keranjang_kerja.append({
                        "id_layanan": layanan_pilihan, "nama_layanan": layanan_obj.nama_layanan,
                        "harga": layanan_obj.harga, "jumlah": qty, "lokasi_ruang": ruangan
                    })
                    st.session_state.nota_siap_download = None 
                    st.rerun()

        with kol2:
            st.subheader("🛒 3. Keranjang Kerja")
            if not st.session_state.keranjang_kerja: 
                st.info("Keranjang kosong. Silakan tambah item pekerjaan di sebelah kiri.")
            else:
                total_nota = 0
                for idx, item in enumerate(st.session_state.keranjang_kerja):
                    subtotal = item["harga"] * item["jumlah"]
                    total_nota += subtotal
                    st.markdown(f"**{idx+1}. {item['nama_layanan']}** ({item['lokasi_ruang']}) x{item['jumlah']}")
                
                st.metric(label="TOTAL NOTA TEMPORER", value=f"Rp {total_nota:,}")
                
                if st.button("🗑️ Reset Keranjang", use_container_width=True):
                    st.session_state.keranjang_kerja = []
                    st.session_state.nota_siap_download = None
                    st.rerun()

                st.divider()
                if st.button("🔥 PROSES & BUAT NOTA INVOICE EXCEL", type="primary", use_container_width=True):
                    if not nama: 
                        st.error("Nama pelanggan wajib diisi!")
                    else:
                        try:
                            id_tek = st.session_state.user_id if st.session_state.user_id else 2
                            
                            order = buat_order_baru(
                                id_teknisi=id_tek, 
                                nama_pelanggan=nama, 
                                alamat=alamat,
                                no_telp=telp, 
                                list_pekerjaan=st.session_state.keranjang_kerja
                            )
                            
                            if order:
                                buffer_nota = buat_file_excel_nota(order, st.session_state.keranjang_kerja)
                                st.session_state.nota_siap_download = {
                                    "label": f"📥 DOWNLOAD NOTA EXCEL ({order.no_invoice})",
                                    "data": buffer_nota.getvalue(),
                                    "filename": f"Nota_{order.no_invoice}_{order.nama_pelanggan}.xlsx"
                                }
                                st.session_state.keranjang_kerja = []
                                st.balloons()
                                st.rerun()
                            else:
                                st.error("❌ Gagal membuat order baru.")
                        except Exception as e:
                            st.error(f"⚠️ Terjadi kesalahan sistem database: {str(e)}")
                
                if st.session_state.nota_siap_download is not None:
                    st.success("✅ DATA BERHASIL MASUK KE SISTEM DATABASE SI DINGIN, BOS!")
                    st.download_button(
                        label=st.session_state.nota_siap_download["label"],
                        data=st.session_state.nota_siap_download["data"],
                        file_name=st.session_state.nota_siap_download["filename"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary", use_container_width=True
                    )
                    if st.button("🔄 Buat Nota Baru Lagi"):
                        st.session_state.nota_siap_download = None
                        st.rerun()

    with tab_riwayat_tek:
        st.subheader(f"📊 Daftar Riwayat Kerja Hasil Inputan: {st.session_state.user_fullname}")
        id_tek_aktif = st.session_state.user_id if st.session_state.user_id else 2
        
        query_riwayat_tek = Order.select().where(Order.teknisi_id == id_tek_aktif).order_by(Order.tanggal_kerja.desc())
        
        if not query_riwayat_tek.exists():
            st.info("Anda belum memiliki riwayat pengerjaan di dalam sistem.")
        else:
            data_tabel_tek = []
            for o in query_riwayat_tek:
                almt_p = getattr(o, 'alamat_pelanggan', '-')
                data_tabel_tek.append({
                    "No Invoice": o.no_invoice,
                    "Tanggal": o.tanggal_kerja,
                    "Nama Pelanggan": o.nama_pelanggan,
                    "Alamat": almt_p,
                    "Status": getattr(o, "status_pembayaran", "Belum Dibayar")
                })
            df_tek_monitor = pd.DataFrame(data_tabel_tek)
            st.dataframe(df_tek_monitor, use_container_width=True, hide_index=True)
            
            st.divider()
            st.markdown("#### 🖨️ Cetak Ulang Nota Lama")
            opsi_riwayat = {o.id: f"{o.no_invoice} - {o.nama_pelanggan} [{getattr(o, 'status_pembayaran', 'Belum Dibayar')}]" for o in query_riwayat_tek}
            nota_tek_sel = st.selectbox("Pilih Struk yang Ingin Didownload Ulang:", options=list(opsi_riwayat.keys()), format_func=lambda x: opsi_riwayat[x], key="sb_riwayat_tek")
            
            if nota_tek_sel:
                ord_tek_obj = Order.get_by_id(nota_tek_sel)
                items_tek = DetailOrder.select().where(DetailOrder.order_id == nota_tek_sel)
                list_items_tek = []
                for it in items_tek:
                    txt_ruang = getattr(it, 'lokasi_ruang', getattr(it, 'lokasi_ruangan', getattr(it, 'ruangan', '-')))
                    list_items_tek.append({
                        "nama_layanan": it.layanan.nama_layanan if it.layanan else "Jasa Service",
                        "harga": it.harga_snapshot, "jumlah": it.jumlah, "lokasi_ruang": txt_ruang
                    })
                
                if st.button(f"🖨️ BUAT EXCEL UNTUK NOTA {ord_tek_obj.no_invoice}", use_container_width=True):
                    buffer_tek_cetak = buat_file_excel_nota(ord_tek_obj, list_items_tek)
                    st.session_state.nota_tek_download_file = {
                        "data": buffer_tek_cetak.getvalue(),
                        "filename": f"Nota_{ord_tek_obj.no_invoice}.xlsx"
                    }
                    st.rerun()
                
                if "nota_tek_download_file" in st.session_state:
                    st.download_button(
                        label="📥 AMBIL FILE NOTA",
                        data=st.session_state.nota_tek_download_file["data"],
                        file_name=st.session_state.nota_tek_download_file["filename"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    if st.button("❌ Selesai"):
                        del st.session_state.nota_tek_download_file
                        st.rerun()

# ==================================================
# VIEW 2: INTERFACE KHUSUS OWNER (BOS FARHAN)
# ==================================================
def tampilan_owner():
    st.sidebar.markdown(f"### 👑 Tingkat Akses: **{st.session_state.user_fullname}**")
    if st.sidebar.button("🔒 Keluar Aplikasi", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

    tab_keuangan, tab_layanan, tab_teknisi, tab_toko, tab_keamanan = st.tabs([
        "📈 Keuangan & Invoice", "⚙️ Kelola Harga Layanan", "🧑‍🔧 Kelola Data Teknisi", "🏨 Pengaturan Toko & Logo", "🔐 Keamanan Akun"
    ])
    
   # --------------------------------------------------
    # TAB KEUANGAN: FIX BUG DUPLICATE FORM KEY & FULL MANAGEMENT
    # --------------------------------------------------
    with tab_keuangan:
        query_order = (Order.select(Order, User).join(User, on=(Order.teknisi_id == User.id)).order_by(Order.tanggal_kerja.desc()))
        orders_data = []
        for o in query_order:
            almt_str = getattr(o, 'alamat_pelanggan', '-')
            orders_data.append({
                "ID": o.id, "No Invoice": o.no_invoice, "Tanggal": o.tanggal_kerja,
                "Pelanggan": o.nama_pelanggan, "Alamat": almt_str,
                "Teknisi": o.teknisi.nama_lengkap if o.teknisi else "Tanpa Teknisi",
                "Total Bayar": o.total_bayar, "Status": getattr(o, "status_pembayaran", "Belum Dibayar")
            })
        df = pd.DataFrame(orders_data)
        
        if df.empty: 
            st.info("Belum ada transaksi di database.")
        else:
            t_omset = df[df["Status"] == "Lunas"]["Total Bayar"].sum()
            t_piutang = df[df["Status"] == "Belum Dibayar"]["Total Bayar"].sum()
            
            kol_m1, kol_m2, kol_m3 = st.columns(3)
            kol_m1.metric("💰 TOTAL OMSET MASUK", f"Rp {t_omset:,}")
            kol_m2.metric("⚠️ TOTAL PIUTANG", f"Rp {t_piutang:,}")
            kol_m3.metric("📦 TOTAL NOTA", f"{len(df)} Nota")
            st.divider()
            
            st.dataframe(df.drop(columns=["ID"]), use_container_width=True, hide_index=True)
            
            st.divider()
            st.markdown("### 🛠️ PANEL KONTROL & EDIT INVOICE")
            opsi_nota = {row["ID"]: f"{row['No Invoice']} - {row['Pelanggan']} (Rp {row['Total Bayar']:,}) [{row['Status']}]" for _, row in df.iterrows()}
            nota_sel = st.selectbox("Pilih Invoice Yang Akan Dikelola/Diedit:", options=list(opsi_nota.keys()), format_func=lambda x: opsi_nota[x], key="sb_owner_sel")
            
            if nota_sel:
                ord_obj = Order.get_by_id(nota_sel)
                query_items = DetailOrder.select().where(DetailOrder.order_id == nota_sel)
                
                list_items = []
                for item in query_items:
                    txt_ruang = getattr(item, 'lokasi_ruang', getattr(item, 'lokasi_ruangan', getattr(item, 'ruangan', '-')))
                    list_items.append({
                        "id_detail": item.id,
                        "nama_layanan": item.layanan.nama_layanan if item.layanan else "Jasa Service Terhapus",
                        "id_layanan": item.layanan.id if item.layanan else None,
                        "harga": item.harga_snapshot, 
                        "jumlah": item.jumlah, 
                        "lokasi_ruang": txt_ruang
                    })
                
                # --------------------------------------------------
                # BAGIAN A: AKSES UTAMA (LUNAS, CETAK, HAPUS INVOICE)
                # --------------------------------------------------
                kol_btn1, kol_btn2, kol_btn3 = st.columns(3)
                
                with kol_btn1:
                    status_sekarang = getattr(ord_obj, "status_pembayaran", "Belum Dibayar")
                    if status_sekarang != "Lunas":
                        if st.button("✅ SET JADI LUNAS", type="primary", use_container_width=True, key=f"lunas_btn_{ord_obj.id}"):
                            Order.update(status_pembayaran="Lunas").where(Order.id == ord_obj.id).execute()
                            st.success("🔥 Invoice berhasil diset Lunas!")
                            st.rerun()
                    else:
                        st.info("👌 Nota Ini Sudah Selesai / Lunas.")
                        
                with kol_btn2:
                    if st.button(f"🖨️ PROSES CETAK STRUK ({ord_obj.no_invoice})", use_container_width=True):
                        buffer_cetak = buat_file_excel_nota(ord_obj, list_items)
                        st.session_state.owner_download_file = {
                            "data": buffer_cetak.getvalue(),
                            "filename": f"Nota_{ord_obj.no_invoice}.xlsx"
                        }
                        st.rerun()
                        
                    if "owner_download_file" in st.session_state:
                        st.download_button(
                            label="📥 DOWNLOAD FILE SEKARANG", 
                            data=st.session_state.owner_download_file["data"], 
                            file_name=st.session_state.owner_download_file["filename"], 
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                            use_container_width=True, type="primary"
                        )
                        if st.button("❌ Selesai Cetak"):
                            del st.session_state.owner_download_file
                            st.rerun()

                with kol_btn3:
                    if st.button(f"🚨 HAPUS INVOICE TOTAL", use_container_width=True, key=f"del_inv_{ord_obj.id}"):
                        DetailOrder.delete().where(DetailOrder.order_id == ord_obj.id).execute()
                        Order.delete().where(Order.id == ord_obj.id).execute()
                        st.warning(f"Invoice {ord_obj.no_invoice} Berhasil Dihapus Permanen!")
                        st.rerun()
                
                st.divider()
                
                # Ambil master layanan untuk opsi selectbox Jasa/Sparepart
                master_layanan = list(Layanan.select())
                dict_master_lay = {l.id: f"{l.nama_layanan} (Rp {l.harga:,})" for l in master_layanan}
                
                # --------------------------------------------------
                # BAGIAN B: FITUR TAMBAH / SUNTIK ITEM BARU (KEY DIJAMIN UNIK)
                # --------------------------------------------------
                st.markdown("#### ➕ Tambah Item Pekerjaan Baru ke Dalam Nota Ini")
                with st.form(key=f"form_suntik_item_baru_nota_{ord_obj.id}", clear_on_submit=True):
                    kol_add1, kol_add2, kol_add3, kol_add4 = st.columns([2, 1.5, 1, 1])
                    
                    with kol_add1:
                        lay_suntik_id = st.selectbox("Pilih Layanan/Bahan Tambahan:", options=list(dict_master_lay.keys()), format_func=lambda x: dict_master_lay[x])
                    with kol_add2:
                        ruang_suntik = st.text_input("Lokasi Ruangan:", placeholder="Contoh: R. Tamu", key=f"txt_ruang_suntik_{ord_obj.id}")
                    with kol_add3:
                        qty_suntik = st.number_input("Qty Unit:", min_value=1, value=1, step=1, key=f"num_qty_suntik_{ord_obj.id}")
                    with kol_add4:
                        st.markdown("<div style='padding-top:28px;'></div>", unsafe_allow_html=True)
                        btn_suntik = st.form_submit_button("➕ Sisipkan Item", type="primary", use_container_width=True)
                        
                    if btn_suntik:
                        if not ruang_suntik:
                            st.error("Gagal! Lokasi ruangan harus diisi agar nota jelas, Bos!")
                        else:
                            lay_terpilih_obj = Layanan.get_by_id(lay_suntik_id)
                            DetailOrder.create(
                                order_id=ord_obj.id,
                                layanan_id=lay_suntik_id,
                                lokasi_ruang=ruang_suntik,
                                harga_snapshot=lay_terpilih_obj.harga,
                                jumlah=qty_suntik
                            )
                            
                            semua_detail_baru = DetailOrder.select().where(DetailOrder.order_id == ord_obj.id)
                            total_baru = sum(d.harga_snapshot * d.jumlah for d in semua_detail_baru)
                            Order.update(total_bayar=total_baru).where(Order.id == ord_obj.id).execute()
                            
                            st.success("✅ Item pekerjaan baru berhasil disisipkan!")
                            st.rerun()

                st.divider()
                
                # --------------------------------------------------
                # BAGIAN C: EDIT & HAPUS LIST ITEM YANG SUDAH ADA (KEY DIJAMIN UNIK)
                # --------------------------------------------------
                st.markdown("#### ⚙️ Kelola/Edit List Detail Pekerjaan Terpasang")
                st.caption("Klik untuk membuka item di bawah jika ingin mengubah data teknis atau menghapusnya.")
                
                dict_master_pure_nama = {l.id: l.nama_layanan for l in master_layanan}
                for idx, item in enumerate(list_items):
                    with st.expander(f"📦 Item ke-{idx+1}: {item['nama_layanan']} ({item['lokasi_ruang']}) - {item['jumlah']} Unit @ Rp {item['harga']:,}"):
                        kol_edit1, kol_edit2 = st.columns(2)
                        
                        with kol_edit1:
                            with st.form(key=f"form_item_edit_detail_id_{item['id_detail']}"):
                                st.markdown("**📝 Ubah Data Item**")
                                index_default_lay = 0
                                if item['id_layanan'] in dict_master_pure_nama:
                                    index_default_lay = list(dict_master_pure_nama.keys()).index(item['id_layanan'])
                                
                                lay_edit_pilih = st.selectbox("Jenis Jasa/Sparepart:", options=list(dict_master_pure_nama.keys()), format_func=lambda x: dict_master_pure_nama[x], index=index_default_lay, key=f"sb_edit_lay_{item['id_detail']}")
                                ruang_edit = st.text_input("Lokasi Ruangan:", value=item['lokasi_ruang'], key=f"txt_edit_ruang_{item['id_detail']}")
                                harga_edit = st.number_input("Harga Jual Snapshot (Rp):", min_value=0, value=int(item['harga']), step=5000, key=f"num_edit_harga_{item['id_detail']}")
                                qty_edit = st.number_input("Jumlah Unit (Qty):", min_value=1, value=int(item['jumlah']), step=1, key=f"num_edit_qty_{item['id_detail']}")
                                
                                btn_simpan_item = st.form_submit_button("🔄 Update Item Ini", use_container_width=True)
                                if btn_simpan_item:
                                    DetailOrder.update(
                                        layanan_id=lay_edit_pilih,
                                        lokasi_ruang=ruang_edit,
                                        harga_snapshot=harga_edit,
                                        jumlah=qty_edit
                                    ).where(DetailOrder.id == item['id_detail']).execute()
                                    
                                    semua_detail_baru = DetailOrder.select().where(DetailOrder.order_id == ord_obj.id)
                                    total_baru = sum(d.harga_snapshot * d.jumlah for d in semua_detail_baru)
                                    Order.update(total_bayar=total_baru).where(Order.id == ord_obj.id).execute()
                                    
                                    st.success("Item pekerjaan diperbarui!")
                                    st.rerun()
                        
                        with kol_edit2:
                            st.markdown("**🗑️ Tindakan Hapus Item**")
                            st.caption("Hapus item ini jika terjadi salah input fatal oleh tim lapangan.")
                            if st.button("❌ HAPUS ITEM INI SAJA", key=f"btn_del_item_id_{item['id_detail']}", use_container_width=True, type="primary"):
                                DetailOrder.delete().where(DetailOrder.id == item['id_detail']).execute()
                                
                                semua_detail_baru = DetailOrder.select().where(DetailOrder.order_id == ord_obj.id)
                                total_baru = sum(d.harga_snapshot * d.jumlah for d in semua_detail_baru) if semua_detail_baru.exists() else 0
                                Order.update(total_bayar=total_baru).where(Order.id == ord_obj.id).execute()
                                
                                st.warning("Item berhasil dibuang!")
                                st.rerun()
                
                st.divider()
                
                # --------------------------------------------------
                # BAGIAN D: FORM EDIT DATA INDUK INVOICE (NAMA & ALAMAT - FIXED INDEPENDENT KEY)
                # --------------------------------------------------
                st.markdown("##### 📝 Panel Edit Informasi Pelanggan Utama:")
                with st.form(key=f"form_edit_induk_invoice_utama_id_{ord_obj.id}"):
                    edit_nama = st.text_input("Ubah Nama Pelanggan:", value=ord_obj.nama_pelanggan, key=f"txt_main_nama_{ord_obj.id}")
                    edit_alamat = st.text_area("Ubah Alamat Rumah:", value=getattr(ord_obj, 'alamat_pelanggan', '-'), key=f"txt_main_alamat_{ord_obj.id}")
                    edit_status = st.selectbox("Sesuaikan Status Pembayaran:", options=["Belum Dibayar", "Lunas"], index=0 if status_sekarang != "Lunas" else 1, key=f"sb_main_status_{ord_obj.id}")
                    
                    if st.form_submit_button("🔄 UPDATE DATA NOTA UTAMA", use_container_width=True):
                        Order.update(
                            nama_pelanggan=edit_nama,
                            alamat_pelanggan=edit_alamat,
                            status_pembayaran=edit_status
                        ).where(Order.id == ord_obj.id).execute()
                        st.success("Data Invoice berhasil diperbarui, Bos!")
                        st.rerun()

    # --------------------------------------------------
    # TAB LAYANAN: SEKARANG BISA TAMBAH & HAPUS ITEM LAYANAN
    # --------------------------------------------------
    with tab_layanan:
        st.subheader("🛠️ Master Harga Jasa Jual")
        semua_layanan = list(Layanan.select())
        df_lay = pd.DataFrame([{"ID": l.id, "Nama Layanan": l.nama_layanan, "Harga Jual": l.harga} for l in semua_layanan])
        
        df_tampil = df_lay.copy()
        if not df_tampil.empty:
            df_tampil["Harga Jual"] = df_tampil["Harga Jual"].apply(lambda x: f"Rp {x:,}")
        st.dataframe(df_tampil, use_container_width=True, hide_index=True)

        st.divider()
        kol_lay1, kol_lay2 = st.columns(2)
        
        with kol_lay1:
            st.markdown("##### ➕ Tambah Item Layanan Baru")
            with st.form("form_tambah_layanan", clear_on_submit=True):
                nama_lay_baru = st.text_input("Nama Layanan / Sparepart Baru:")
                harga_lay_baru = st.number_input("Harga Jual (Rp):", min_value=0, step=25000)
                submit_lay = st.form_submit_button("💾 Simpan Layanan Baru", type="primary", use_container_width=True)
                if submit_lay and nama_lay_baru:
                    Layanan.create(nama_layanan=nama_lay_baru, harga=harga_lay_baru)
                    st.success(f"Berhasil menambah layanan: {nama_lay_baru}")
                    st.rerun()
                    
        with kol_lay2:
            st.markdown("##### 🗑️ Hapus Item Layanan")
            if semua_layanan:
                opsi_hapus_lay = {l.id: f"{l.nama_layanan} (Rp {l.harga:,})" for l in semua_layanan}
                lay_target_hapus = st.selectbox("Pilih Layanan yang Ingin Dihapus Permanen:", options=list(opsi_hapus_lay.keys()), format_func=lambda x: opsi_hapus_lay[x])
                if st.button("❌ HAPUS LAYANAN DARI SISTEM", type="primary", use_container_width=True):
                    Layanan.delete().where(Layanan.id == lay_target_hapus).execute()
                    st.success("Layanan berhasil dihapus dari sistem!")
                    st.rerun()
            else:
                st.info("Belum ada layanan yang terdaftar.")

    # --------------------------------------------------
    # TAB DATA TEKNISI: SEKARANG BISA HAPUS AKUN TEKNISI
    # --------------------------------------------------
    with tab_teknisi:
        st.subheader("🧑‍🔧 Akun Tim Teknisi & Password")
        semua_tek = list(User.select().where(User.role == "teknisi"))
        
        data_tek_list = []
        for t in semua_tek:
            nomor_hp_tampil = getattr(t, 'no_hp', getattr(t, 'no_telp', '-'))
            data_tek_list.append({
                "ID": t.id, "Nama Lengkap": t.nama_lengkap, "No HP/WA": nomor_hp_tampil, 
                "Username Login": t.username, "🔑 Password": t.password
            })
            
        df_tek = pd.DataFrame(data_tek_list)
        if not df_tek.empty: 
            st.dataframe(df_tek, use_container_width=True, hide_index=True)
            
        st.divider()
        kol_tek_f1, kol_tek_f2 = st.columns(2)
        
        with kol_tek_f1:
            st.markdown("##### ➕ Daftarkan Akun Teknisi Baru")
            input_nama_tek = st.text_input("Nama Lengkap Teknisi:")
            input_hp_tek = st.text_input("Nomor HP/WA Teknisi:")
            input_user_tek = st.text_input("Username Akun:")
            input_pass_tek = st.text_input("Password Akun:")
                
            if st.button("💾 Simpan Akun Teknisi", use_container_width=True, type="primary"):
                if input_nama_tek and input_user_tek and input_pass_tek:
                    simpan_data = {
                        "nama_lengkap": input_nama_tek, "username": input_user_tek,
                        "password": input_pass_tek, "role": "teknisi"
                    }
                    if hasattr(User, 'no_hp'): simpan_data["no_hp"] = input_hp_tek
                    elif hasattr(User, 'no_telp'): simpan_data["no_telp"] = input_hp_tek
                    
                    User.create(**simpan_data)
                    st.success("Akun teknisi baru berhasil didaftarkan!")
                    st.rerun()
                else: 
                    st.error("Gagal! Nama, Username, dan Password wajib diisi!")
                    
        with kol_tek_f2:
            st.markdown("##### 🗑️ Hapus Akun Teknisi (Nonaktifkan)")
            if semua_tek:
                opsi_hapus_tek = {t.id: f"{t.nama_lengkap} ({t.username})" for t in semua_tek}
                tek_target_hapus = st.selectbox("Pilih Akun Teknisi Yang Ingin Dihapus:", options=list(opsi_hapus_tek.keys()), format_func=lambda x: opsi_hapus_tek[x])
                if st.button("🚨 HAPUS AKUN TEKNISI PERMANEN", type="primary", use_container_width=True):
                    User.delete().where(User.id == tek_target_hapus).execute()
                    st.success("Akun teknisi telah sukses dihapus dari database.")
                    st.rerun()
            else:
                st.info("Tidak ada data teknisi.")

    # --------------------------------------------------
    # TAB PENGATURAN IDENTITAS TOKO & LOGO
    # --------------------------------------------------
    with tab_toko:
        st.subheader("🏨 Pengaturan Identitas Usaha & Cetakan Nota")
        
        col_toko1, col_toko2 = st.columns(2)
        with col_toko1:
            input_nama_usaha = st.text_input("Nama Usaha (Muncul di KOP Atas):", value=st.session_state.nama_usaha)
            input_alamat_usaha = st.text_area("Slogan / Alamat Usaha:", value=st.session_state.alamat_usaha)
            input_kontak_usaha = st.text_input("Nomor Kontak Resmi Usaha (WA/Telp):", value=st.session_state.kontak_usaha)
        
        with col_toko2:
            st.markdown("#### 🖼️ Pasang Logo Usaha (Gambar JPG/PNG)")
            file_logo = st.file_uploader("Pilih file foto logo toko Bos:", type=["jpg", "jpeg", "png"])
            
            if file_logo is not None:
                with open("logo.png", "wb") as f:
                    f.write(file_logo.getbuffer())
                st.success("✨ Logo Usaha Berhasil Diunggah & Disimpan!")
            
            if os.path.exists("logo.png"):
                st.image("logo.png", caption="Logo Aktif Saat Ini", width=120)
        
        if st.button("💾 SIMPAN IDENTITAS BARU", use_container_width=True, type="primary"):
            st.session_state.nama_usaha = input_nama_usaha
            st.session_state.alamat_usaha = input_alamat_usaha
            st.session_state.kontak_usaha = input_kontak_usaha
            st.success("Data Toko berhasil diupdate!")
            st.rerun()

   # --------------------------------------------------
    # TAB KEAMANAN AKUN: BYPASS VERIFIKASI PASSWORD LAMA
    # --------------------------------------------------
    with tab_keamanan:
        st.subheader("🔐 Kelola Kredensial Akses Owner")
        owner_now = User.select().where(User.role == "owner").first()
        
        kolom_email_tersedia = 'no_hp' if hasattr(User, 'no_hp') else ('no_telp' if hasattr(User, 'no_telp') else None)
        current_email = getattr(owner_now, kolom_email_tersedia, '') if (owner_now and kolom_email_tersedia) else ""
        
        with st.form("form_ganti_kredensial"):
            st.markdown("##### 📝 Perbarui Profil Akun")
            u_baru = st.text_input("Username Baru Owner:", value=owner_now.username if owner_now else "farhan")
            
            # NOTIFIKASI BYPASS
            st.info("💡 Jalur Pintas Aktif: Bos bisa langsung isi password baru tanpa perlu input password lama.")
            p_baru = st.text_input("Masukkan Password Baru Langsung:", type="password")
            
            st.markdown("---")
            st.markdown("##### 📨 Email Darurat Pemulihan")
            e_baru = st.text_input("Email Valid Owner (Digunakan saat lupa sandi):", value=current_email)
            
            submit_keamanan = st.form_submit_button("💾 SIMPAN & PAKSA UPDATE SISTEM", type="primary", use_container_width=True)
            
            if submit_keamanan:
                if not u_baru or not e_baru:
                    st.error("⚠️ Kolom Username dan Email darurat tidak boleh kosong, Bos!")
                elif not p_baru:
                    st.error("⚠️ Silakan isi Password Baru terlebih dahulu, Bos!")
                else:
                    # Langsung pakai password baru tanpa cek password lama
                    update_fields = {"username": u_baru, "password": p_baru}
                    if kolom_email_tersedia:
                        update_fields[kolom_email_tersedia] = e_baru
                    
                    if owner_now:
                        User.update(**update_fields).where(User.id == owner_now.id).execute()
                    else:
                        update_fields["role"] = "owner"
                        update_fields["nama_lengkap"] = "Farhan Kholili"
                        User.create(**update_fields)
                        
                    st.success("✨ Sukses Besar! Akun Owner berhasil dipaksa update ke database.")
                    st.session_state.authenticated = False
                    st.rerun()

# ==================================================
# ALUR KONTROL ROUTING UTAMA
# ==================================================
if not st.session_state.authenticated:
    halaman_login()
else:
    if st.session_state.user_role == "owner":
        tampilan_owner()
    else:
        tampilan_teknisi()
