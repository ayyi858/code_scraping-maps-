import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.exc import GeocoderTimedOut, GeocoderServiceError
import re
import time

# 1. Baca data Excel
file_path = "oce.xlsx"
try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    print(f"Error: File '{file_path}' tidak ditemukan.")
    exit()

# 2. Perbaiki kolom Kecamatan berdasarkan teks di Alamat
def perbaiki_kecamatan(row):
    alamat = str(row['alamat'])
    match = re.search(r'Kec\.\s*([^,]+)', alamat)
    if match:
        return match.group(1).strip()
    return row['kecamatan']

df['kecamatan_revisi'] = df.apply(perbaiki_kecamatan, axis=1)

# 3. Geocoding dengan pengaturan lebih kuat
# Menambahkan timeout yang lebih besar saat inisialisasi geolocator
geolocator = Nominatim(user_agent="makassar_geocoder_v2", timeout=10)

# Menambah delay dan jumlah retry pada RateLimiter
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.5, max_retries=3, error_wait_seconds=5)

def bersihkan_alamat(alamat):
    """Menghapus kode Plus dan menyederhanakan alamat agar lebih mudah dicari."""
    alamat = str(alamat)
    # Hapus kode Plus (misal: VC39+M4P)
    alamat = re.sub(r'^[A-Z0-9\+]+\,\s*', '', alamat)
    
    # Hapus kata-kata dalam kurung yang sering membingungkan geocoder
    alamat = re.sub(r'\([^)]*\)', '', alamat)
    
    # Hapus teks spesifik yang mungkin tidak dikenali
    alamat = alamat.replace('Inspeksi Kanal 1', '')
    
    # Opsional: Jika masih sering gagal, kita bisa ambil bagian akhir alamat (Kecamatan & Kota)
    # sebagai fall-back (lihat fungsi dapatkan_koordinat)
    
    # Rapikan spasi ganda
    alamat = " ".join(alamat.split())
    return alamat

def dapatkan_koordinat(row):
    alamat_asli = row['alamat']
    alamat_bersih = bersihkan_alamat(alamat_asli)
    
    try:
        # Coba alamat lengkap (yang sudah dibersihkan dari kode Plus/kurung)
        location = geocode(alamat_bersih)
        if location:
             return pd.Series([location.latitude, location.longitude])
        
        # JIKA GAGAL (lokasi tidak ditemukan), coba Fallback 1: Hanya nama Jalan dan Kota
        # Mengambil teks sebelum koma pertama (biasanya nama jalan) dan menggabungkannya dengan kota
        jalan = alamat_bersih.split(',')[0].strip()
        if "Jl." in jalan or "Jalan" in jalan:
             alamat_fallback = f"{jalan}, Makassar, Sulawesi Selatan"
             print(f"  --> Coba fallback: {alamat_fallback}")
             location = geocode(alamat_fallback)
             if location:
                 return pd.Series([location.latitude, location.longitude])
                 
        # JIKA MASIH GAGAL, coba Fallback 2: Koordinat Kecamatan (titik tengah kecamatan)
        kec = row['kecamatan_revisi']
        if pd.notna(kec):
            alamat_kecamatan = f"Kecamatan {kec}, Makassar, Sulawesi Selatan"
            print(f"  --> Coba kecamatan: {alamat_kecamatan}")
            location = geocode(alamat_kecamatan)
            if location:
                 return pd.Series([location.latitude, location.longitude])
                 
        return pd.Series([None, None])

    except (GeocoderTimedOut, GeocoderServiceError) as e:
        print(f"Error pada alamat '{alamat_bersih}': {e}. Melewati baris ini.")
        return pd.Series([None, None])

print("Sedang mencari koordinat... Proses ini membutuhkan kesabaran.")
df[['latitude', 'longitude']] = df.apply(dapatkan_koordinat, axis=1)

# 4. Evaluasi Hasil
jumlah_sukses = df['latitude'].notna().sum()
print(f"\nSelesai! Berhasil mendapatkan koordinat untuk {jumlah_sukses} dari {len(df)} lokasi.")

# 5. Simpan ke file Excel baru
output_file = "Data_Makassar_Diperbaiki.xlsx"
df.to_excel(output_file, index=False)
print(f"Data telah disimpan di {output_file}")