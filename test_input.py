from database import db, User, Layanan, Order, DetailOrder
import datetime

# 1. FUNGSI INVOICE OTOMATIS
def generate_no_invoice():
    hari_ini = datetime.date.today().strftime("%Y%m%d")
    prefix = f"INV-{hari_ini}-"
    order_terakhir = Order.select().where(Order.no_invoice.startswith(prefix)).order_by(Order.no_invoice.desc()).first()
    if order_terakhir:
        nomor_urut_terakhir = int(order_terakhir.no_invoice.split("-")[-1])
        nomor_baru = nomor_urut_terakhir + 1
    else:
        nomor_baru = 1
    return f"{prefix}{nomor_baru:04d}"

# 2. FUNGSI BUAT ORDER
def buat_order_baru(id_teknisi, nama_pelanggan, alamat, no_telp, list_pekerjaan):
    with db.atomic() as transaction:
        try:
            no_inv = generate_no_invoice()
            order = Order.create(
                no_invoice=no_inv,
                nama_pelanggan=nama_pelanggan,
                alamat_pelanggan=alamat,
                no_telp_pelanggan=no_telp,
                teknisi_id=id_teknisi
            )
            total_keseluruhan = 0
            for item in list_pekerjaan:
                layanan = Layanan.get_by_id(item["id_layanan"])
                subtotal_item = layanan.harga * item["jumlah"]
                total_keseluruhan += subtotal_item
                
                DetailOrder.create(
                    order=order,
                    layanan=layanan,
                    harga_snapshot=layanan.harga,
                    jumlah=item["jumlah"],
                    lokasi_ruang=item["lokasi_ruang"],
                    subtotal=subtotal_item
                )
            order.total_bayar = total_keseluruhan
            order.save()
            print(f"\n==========================================")
            print(f"🔥 SUKSES! Order {no_inv} Berhasil Dibuat!")
            print(f"Pelanggan : {nama_pelanggan}")
            print(f"Total Nota: Rp {total_keseluruhan:,}")
            print(f"==========================================\n")
            return order
        except Exception as e:
            transaction.rollback()
            print(f"Gagal membuat order: {e}")
            return None

# ==================================================
# JALANKAN SIMULASI LANGSUNG DI SINI
# ==================================================
if __name__ == '__main__':
    # SEKARANG SUDAH DIPERBAIKI MENGGUNAKAN reuse_if_open=True
    db.connect(reuse_if_open=True)
    
    if User.select().count() == 0:
        User.create(username="owner1", password="123", nama_lengkap="Farhan Kholili", role="admin")
        User.create(id=2, username="budi", password="123", nama_lengkap="Budi Teknisi", role="teknisi")
        print("-> Sukses menginput data User awal.")
        
    if Layanan.select().count() == 0:
        Layanan.create(id=1, nama_layanan="Cuci AC 1/2 PK - 1 PK", harga=75000)
        Layanan.create(id=2, nama_layanan="Cuci AC 1.5 PK - 2 PK", harga=90000)
        Layanan.create(id=3, nama_layanan="Isi/Tambah Freon R32", harga=150000)
        Layanan.create(id=4, nama_layanan="Ganti Kapasitor", harga=125000)
        print("-> Sukses menginput data Master Harga Layanan.")

    # Data simulasi kerjaan si Budi
    pekerjaan_budi = [
        {"id_layanan": 1, "jumlah": 1, "lokasi_ruang": "Kamar Utama Lt. 1"},
        {"id_layanan": 1, "jumlah": 1, "lokasi_ruang": "Kamar Anak Lt. 2"},
        {"id_layanan": 3, "jumlah": 1, "lokasi_ruang": "Kamar Utama Lt. 1"}
    ]

    print("\nMencoba membuat nota digital untuk Pak Ahmad...")
    buat_order_baru(
        id_teknisi=2, 
        nama_pelanggan="Pak Ahmad", 
        alamat="Jl. Merdeka No. 10", 
        no_telp="08123456789", 
        list_pekerjaan=pekerjaan_budi
    )