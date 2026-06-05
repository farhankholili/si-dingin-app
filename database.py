from peewee import *
import datetime

# Membuat database lokal SQLite
db = SqliteDatabase('si_dingin.db')

class BaseModel(Model):
    class Meta:
        database = db

# 1. TABEL USER (Admin/Owner & Teknisi)
class User(BaseModel):
    username = CharField(unique=True)
    password = CharField()  
    nama_lengkap = CharField()
    role = CharField()  # 'admin' atau 'teknisi'
    is_active = BooleanField(default=True)

# 2. TABEL MASTER LAYANAN/HARGA
class Layanan(BaseModel):
    nama_layanan = CharField()  # Contoh: "Cuci AC 1 PK", "Isi Freon R32"
    harga = IntegerField()      
    keterangan = TextField(null=True)

# 3. TABEL ORDER / PEKERJAAN UTAMA
class Order(BaseModel):
    no_invoice = CharField(unique=True)  
    nama_pelanggan = CharField()
    alamat_pelanggan = TextField()
    no_telp_pelanggan = CharField()
    tanggal_kerja = DateField(default=datetime.date.today)
    teknisi = ForeignKeyField(User, backref='orders')
    status_pengerjaan = CharField(default='Proses') # 'Proses', 'Selesai'
    status_pembayaran = CharField(default='Belum Dibayar') # 'Belum Dibayar', 'Lunas'
    total_bayar = IntegerField(default=0)
    catatan_internal = TextField(null=True) 

# 4. TABEL DETAIL ORDER (Solusi Item Sama Beda Ruangan)
class DetailOrder(BaseModel):
    order = ForeignKeyField(Order, backref='items', on_delete='CASCADE')
    layanan = ForeignKeyField(Layanan, backref='detail_orders')
    harga_snapshot = IntegerField() 
    jumlah = IntegerField(default=1)
    lokasi_ruang = CharField() # KUNCI: "Kamar Utama", "Ruang Tamu", dll.
    subtotal = IntegerField()

# Fungsi untuk membuat tabel otomatis
def init_db():
    db.connect()
    db.create_tables([User, Layanan, Order, DetailOrder])
    print("Database 'Si Dingin' dan tabel-tabel berhasil dibuat!")

if __name__ == '__main__':
    init_db()