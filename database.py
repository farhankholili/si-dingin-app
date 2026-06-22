from peewee import *
import datetime

# Membuat database lokal SQLite
db = SqliteDatabase('si_dingin.db')

class BaseModel(Model):
    class Meta:
        database = db

# ==========================================
# 0. TABEL BENGKEL (Kunci Utama Multi-Tenant)
# ==========================================
class Bengkel(BaseModel):
    nama_bengkel = CharField(unique=True)
    alamat = TextField(null=True)
    no_telp = CharField(null=True)
    owner_name = CharField()
    is_active = BooleanField(default=True)

# 1. TABEL USER (Ditambahkan relasi ke Bengkel)
class User(BaseModel):
    bengkel = ForeignKeyField(Bengkel, backref='users', null=True) # Terhubung ke Bengkel
    username = CharField(unique=True)
    password = CharField()  
    nama_lengkap = CharField()
    role = CharField()  # 'super_admin', 'owner', atau 'teknisi'
    is_active = BooleanField(default=True)

# 2. TABEL MASTER LAYANAN/HARGA (Ditambahkan relasi ke Bengkel)
class Layanan(BaseModel):
    bengkel = ForeignKeyField(Bengkel, backref='layanan_bengkel', null=True) # Milik bengkel mana
    nama_layanan = CharField()  # Contoh: "Cuci AC 1 PK", "Isi Freon R32"
    harga = IntegerField()      
    keterangan = TextField(null=True)

# 3. TABEL ORDER / PEKERJAAN UTAMA (Ditambahkan relasi ke Bengkel)
class Order(BaseModel):
    bengkel = ForeignKeyField(Bengkel, backref='orders_bengkel', null=True) # Orderan bengkel mana
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

# 4. TABEL DETAIL ORDER
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
    # Tambahkan Bengkel ke dalam list antrean pembuatan tabel
    db.create_tables([Bengkel, User, Layanan, Order, DetailOrder])
    print("Database Multi-Tenant 'Si Dingin' dan tabel-tabel berhasil dibuat!")

if __name__ == '__main__':
    init_db()
