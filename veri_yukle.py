from supabase import create_client, Client

# =====================================================================
# SUPABASE BAĞLANTI AYARLARI
# =====================================================================
# Lütfen aşağıdaki tırnakların içine 5. Adımda aldığınız kendi bilgilerinizi yazın:
SUPABASE_URL = "https://your-project-url.supabase.co"
SUPABASE_KEY = "your-anon-key"

if SUPABASE_URL.startswith("https://your-") or SUPABASE_KEY == "your-anon-key":
    print("HATA: Lütfen önce SUPABASE_URL ve SUPABASE_KEY bilgilerinizi giriniz!")
    exit()

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# =====================================================================
# SEÇENEK 1: TÜM VERİTABANINI TEMİZLEME (ÖĞRENCİ SİLME) FONKSİYONU
# =====================================================================
def veritabanini_sifirla():
    """Veritabanındaki tüm öğrencileri, sınıfları ve cevapları kalıcı olarak siler."""
    print("\n⚠️ VERİTABANI SİLME İŞLEMİ BAŞLADI...")
    try:
        # 1. Önce anket cevaplarını silelim (Bağlantılı tablo olduğu için)
        supabase.table("riba_cevaplari").delete().neq("id", 0).execute()
        print("-> Tüm anket cevapları silindi.")

        # 2. Öğrencileri silelim
        supabase.table("ogrenciler").delete().neq("id", 0).execute()
        print("-> Tüm öğrenciler veritabanından silindi.")

        # 3. Sınıfları silelim
        supabase.table("siniflar").delete().neq("id", 0).execute()
        print("-> Tüm sınıflar silindi.")
        
        print("✅ VERİTABANI TAMAMEN TEMİZLENDİ! SİSTEM SIFIR DURUMDA.\n")
    except Exception as e:
        print(f"Silme işlemi sırasında hata oluştu: {str(e)}")

# =====================================================================
# SEÇENEK 2: EXCEL / LİSTE MANTIĞIYLA TOPLU ÖĞRENCİ YÜKLEME FONKSİYONU
# =====================================================================
def donem_ve_sinif_kur(donem_adi, sinif_adi, ogrenci_listesi):
    """Belirtilen dönemi, sınıfı açar ve öğrencileri toplu yükler."""
    try:
        # Dönem Kontrolü
        donem_kontrol = supabase.table("donemler").select("id").eq("donem_adi", donem_adi).execute()
        if len(donem_kontrol.data) == 0:
            supabase.table("donemler").update({"aktif_mi": False}).eq("aktif_mi", True).execute()
            yeni_donem = supabase.table("donemler").insert({"donem_adi": donem_adi, "aktif_mi": True}).execute()
            donem_id = yeni_donem.data[0]['id']
        else:
            donem_id = donem_kontrol.data[0]['id']

        # Sınıf Kontrolü
        sinif_kontrol = supabase.table("siniflar").select("id").eq("sinif_adi", sinif_adi).execute()
        if len(sinif_kontrol.data) == 0:
            yeni_sinif = supabase.table("siniflar").insert({"sinif_adi": sinif_adi}).execute()
            sinif_id = yeni_sinif.data[0]['id']
        else:
            sinif_id = sinif_kontrol.data[0]['id']

        # Öğrencileri Toplu Ekleme
        eklenen = 0
        for ogrenci in ogrenci_listesi:
            try:
                supabase.table("ogrenciler").insert({
                    "okul_no": ogrenci["okul_no"],
                    "ad_soyad": ogrenci["ad_soyad"],
                    "sinif_id": sinif_id
                }).execute()
                eklenen += 1
            except:
                pass # Aynı numara varsa hata vermemesi için atla
                
        print(f"✅ {sinif_adi} sınıfı için {eklenen} öğrenci başarıyla yüklendi.")
    except Exception as e:
        print(f"Yükleme hatası: {str(e)}")

# =====================================================================
# KONTROL PANELİ (NE YAPMAK İSTİYORSANIZ AŞAĞIDAN AYARLAYIN)
# =====================================================================
if __name__ == "__main__":
    
    # ❌ EĞER TÜM ÖĞRENCİLERİ VE ESKİ VERİLERİ SİLMEK İSTİYORSANIZ:
    # Aşağıdaki satırın başındaki '#' işaretini kaldırın ve dosyayı çalıştırın:
    # veritabanini_sifirla()


    # 📝 EĞER SİSTEME YENİ ÖRNEK ÖĞRENCİLER YÜKLEMEK İSTİYORSANIZ:
    # (Silme işlemi kapalıyken burası çalışır)
    
    ornek_6A = [
        {"okul_no": 101, "ad_soyad": "Ali Yılmaz"},
        {"okul_no": 102, "ad_soyad": "Fatma Demir"},
        {"okul_no": 103, "ad_soyad": "Sıraç Aksan"}
    ]
    
    ornek_6B = [
        {"okul_no": 201, "ad_soyad": "Zeynep Kaya"},
        {"okul_no": 202, "ad_soyad": "Mehmet Çelik"}
    ]

    print("--- Veri Yükleme İşlemi Başlatılıyor ---")
    donem_ve_sinif_kur("2025-2026 Eğitim Öğretim Yılı", "6/A", ornek_6A)
    donem_ve_sinif_kur("2025-2026 Eğitim Öğretim Yılı", "6/B", ornek_6B)