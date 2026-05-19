import os
from supabase import create_client, Client

# DİKKAT: Supabase bilgilerini doğrudan buraya da yazabilirsiniz 
# ya da bilgisayarınızın çevre değişkenlerinden okutabilirsiniz.
# Testi kolaylaştırmak için aşağıdaki iki tırnağın içine kendi şifrelerinizi yapıştırabilirsiniz:
SUPABASE_URL = "https://pmlgahbdpzlhzlbvfpxf.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBtbGdhaGJkcHpsaHpsYnZmcHhmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkyMjI4MzQsImV4cCI6MjA5NDc5ODgzNH0.6-OdE2chR29IJHKV1lCHYDkxE5HkkMtvMmXEkFYnbH0"

if SUPABASE_URL.startswith("5._ADIMDA") or SUPABASE_KEY.startswith("5._ADIMDA"):
    print("LÜTFEN ÖNCE SUPABASE URL VE KEY BİLGİLERİNİZİ KODA YAZINIZ!")
    exit()

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def donem_olustur_ve_aktif_et(donem_adi):
    """Sistemde aktif bir eğitim yılı dönemi açar."""
    print(f"'{donem_adi}' kontrol ediliyor...")
    mevcut = supabase.table("donemler").select("id").eq("donem_adi", donem_adi).execute()
    
    if len(mevcut.data) == 0:
        # Önceki tüm dönemleri pasif yap
        supabase.table("donemler").update({"aktif_mi": False}).eq("aktif_mi", True).execute()
        # Yeni dönemi aktif olarak ekle
        yeni = supabase.table("donemler").insert({"donem_adi": donem_adi, "aktif_mi": True}).execute()
        print(f"Yeni aktif dönem başarıyla açıldı: {donem_adi}")
        return yeni.data[0]['id']
    else:
        print(f"'{donem_adi}' zaten mevcut.")
        return mevcut.data[0]['id']

def sinif_ve_ogrencileri_yukle(sinif_adi, ogrenci_listesi):
    """
    Belirtilen sınıfa ait öğrencileri toplu halde yükler.
    ogrenci_listesi: [{"okul_no": 120, "ad_soyad": "Ahmet Yılmaz"}, ...] formatında olmalıdır.
    """
    # 1. Sınıf veritabanında var mı kontrol et, yoksa oluştur
    sinif_kontrol = supabase.table("siniflar").select("id").eq("sinif_adi", sinif_adi).execute()
    
    if len(sinif_kontrol.data) == 0:
        yeni_sinif = supabase.table("siniflar").insert({"sinif_adi": sinif_adi}).execute()
        sinif_id = yeni_sinif.data[0]['id']
        print(f"'{sinif_adi}' sınıfı sisteme yeni eklendi.")
    else:
        sinif_id = sinif_kontrol.data[0]['id']
        print(f"'{sinif_adi}' sınıfı zaten mevcut, ID'si alındı.")

    # 2. Öğrencileri toplu olarak veritabanına ekle
    eklenen_sayisi = 0
    for ogrenci in ogrenci_listesi:
        try:
            # Öğrenciyi eklemeyi dene
            supabase.table("ogrenciler").insert({
                "okul_no": ogrenci["okul_no"],
                "ad_soyad": ogrenci["ad_soyad"],
                "sinif_id": sinif_id
            }).execute()
            eklenen_sayisi += 1
        except Exception as e:
            # Eğer numara zaten varsa hata verecektir, atla
            print(f"Hata: {ogrenci['okul_no']} numaralı öğrenci zaten kayıtlı olabilir.")

    print(f"--> {sinif_adi} sınıfına {eklenen_sayisi} yeni öğrenci başarıyla yüklendi!\n")

# ==========================================
# ÇALIŞTIRMA VE TEST ALANI
# ==========================================
if __name__ == "__main__":
    # 1. Önce aktif eğitim yılını tanımlıyoruz (Bu her sene 1 kez yapılır)
    # Bu sayede her senenin raporu birbirine karışmadan kalıcı saklanacak
    aktif_donem_id = donem_olustur_ve_aktif_et("2025-2026 Eğitim Öğretim Yılı")

    # 2. ÖRNEK DENEME LİSTELERİ
    # Normalde e-okuldan aldığınız excel verisini buraya döngüyle bağlayacağız.
    # Sistemin çalıştığını görmek için aşağıdaki deneme verilerini yükleyelim:
    
    ornek_6A_listesi = [
        {"okul_no": 101, "ad_soyad": "Ali Yılmaz"},
        {"okul_no": 102, "ad_soyad": "Fatma Demir"},
        {"okul_no": 103, "ad_soyad": "Sıraç Aksan"}
    ]
    
    ornek_6B_listesi = [
        {"okul_no": 201, "ad_soyad": "Zeynep Kaya"},
        {"okul_no": 202, "ad_soyad": "Mehmet Çelik"}
    ]

    # Fonksiyonları çağırıp veritabanına yükleme yapıyoruz
    sinif_ve_ogrencileri_yukle("6/A", ornek_6A_listesi)
    sinif_ve_ogrencileri_yukle("6/B", ornek_6B_listesi)
    
    print("Tüm yükleme işlemi tamamlandı!")