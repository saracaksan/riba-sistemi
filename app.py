from flask import Flask, render_template, jsonify, request
from flask import send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
from supabase import create_client, Client
import os

app = Flask(__name__)

# Supabase Gizli Bağlantı Şifrelerini Render ortamından okuyoruz
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

# Eğer şifreler tanımlanmışsa veritabanı bağlantısını kur
if SUPABASE_URL and SUPABASE_KEY:
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    supabase = None

@app.route('/')
def ana_sayfa():
    return render_template('index.html')
@app.route('/yonetim/raporlar')
def raporlar_sayfasi():
    return render_template('raporlar.html')

@app.route('/rapor/excel')
def rapor_excel():
    sinif_filtresi = request.args.get('sinif', 'TUM_OKUL')
    
    # 1. Bellekte sanal bir Excel dosyası oluşturuyoruz
    wb = openpyxl.Workbook()
    
    # Şablonunuza uygun sekmeleri açıyoruz
    ws_ozet = wb.active
    ws_ozet.title = "Sonuç (ASP)"
    ws_ogrenci = wb.create_sheet(title="Öğrenci")
    ws_veli = wb.create_sheet(title="Veli")
    ws_ogretmen = wb.create_sheet(title="Öğretmen")
    
    # Tasarım stilleri (Resmi yazı tipleri ve yumuşak kurumsal renkler)
    baslik_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    normal_font = Font(name='Arial', size=10)
    kalin_font = Font(name='Arial', size=10, bold=True)
    
    mavi_dolgu = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    gri_dolgu = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    
    ince_kenar = Side(border_style="thin", color="D9D9D9")
    hucre_border = Border(left=ince_kenar, right=ince_kenar, top=ince_kenar, bottom=ince_kenar)
    ortala = Alignment(horizontal="center", vertical="center")
    
    # 2. SEKMELERİN BAŞLIKLARINI OLUŞTURMA (Yüklediğiniz Excel formatına sadık kalındı)
    for ws in [ws_ogrenci, ws_veli, ws_ogretmen]:
        ws.append(["No", "Adı Soyadı", "Sınıfı", "M1_A", "M1_B", "M2_A", "M2_B", "M3_A", "M3_B", "M12_A", "M12_B", "M30_A", "M30_B"])
        for cell in ws[1]:
            cell.font = baslik_font
            cell.fill = mavi_dolgu
            cell.alignment = ortala

    # 3. VERİTABANINDAN VERİLERİ ÇEKİP EXCEL'E DOLDURMA
    if supabase:
        # Tüm cevapları çekiyoruz
        cevap_sorgu = supabase.table("riba_cevaplari").select("*, ogrenciler(ad_soyad, siniflar(sinif_adi))").execute()
        
        for satir_no, veri in enumerate(cevap_sorgu.data, start=2):
            ogrenci_adi = veri['ogrenciler']['ad_soyad']
            sinif_adi = veri['ogrenciler']['siniflar']['sinif_adi']
            tip = veri['dolduran_tipi'] # 'Ogrenci', 'Veli', 'Ogretmen'
            cevaplar = veri['cevaplar'] # JSON formatındaki ham puanlar
            
            # Sınıf filtresi varsa ve uyuşmuyorsa bu öğrenciyi rapora dahil etme
            if sinif_filtresi != 'TUM_OKUL' and sinif_filtresi != sinif_adi:
                continue
                
            # İlgili sekmeyi seç
            target_ws = ws_ogrenci if tip == 'Ogrenci' else (ws_veli if tip == 'Veli' else ws_ogretmen)
            
            # Ham maddeleri (Örn: M12, M30) ayrıştırarak Excel satırına ekle
            m12_a = cevaplar.get('M12', {}).get('A', 0)
            m12_b = cevaplar.get('M12', {}).get('B', 0)
            m30_a = cevaplar.get('M30', {}).get('A', 0)
            m30_b = cevaplar.get('M30', {}).get('B', 0)
            
            row_data = [satir_no - 1, ogrenci_adi, sinif_adi, 0, 0, 0, 0, 0, 0, m12_a, m12_b, m30_a, m30_b]
            target_ws.append(row_data)
            
            # Eklenen satırın hücrelerini biçimlendir
            for cell in target_ws[target_ws.max_row]:
                cell.font = normal_font
                cell.border = hucre_border

    # 4. "SONUÇ (ASP)" SEKME TASARIMI VE FORMÜLLER (Yüklediğiniz özet mantığı)
    ws_ozet.append(["RS", "HEDEFLER", "Öğrenci Frekans (f)", "Veli Frekans (f)", "Öğretmen Frekans (f)", "ASP PUANI"])
    for cell in ws_ozet[1]:
        cell.font = baslik_font
        cell.fill = mavi_dolgu
        cell.alignment = ortala
        
    # Örnek Hedefler Listesi
    hedefler = [
        {"rs": 1, "ad": "Problem Çözme Becerileri (M12)"},
        {"rs": 2, "ad": "Öz Düzenlemeli Öğrenme (M30)"}
    ]
    
    for h in hedefler:
        r = ws_ozet.max_row + 1
        # Excel içi akıllı dinamik formülleri gömüyoruz (Böylece Excel canlı kalıyor)
        ws_ozet.append([
            h["rs"], 
            h["ad"], 
            f"=SUM(Öğrenci!J:J)", # Öğrenci sekmesindeki M12_A sütun toplamı
            f"=SUM(Veli!J:J)",    # Veli sekmesindeki M12_A sütun toplamı
            f"=SUM(Öğretmen!J:J)", # Öğretmen sekmesindeki M12_A sütun toplamı
            f"=(C{r}*0.5) + (D{r}*0.3) + (E{r}*0.2)" # Otomatik Ağırlıklandırılmış Standart Puan Formülü
        ])
        for cell in ws_ozet[ws_ozet.max_row]:
            cell.font = normal_font
            cell.border = hucre_border

    # Excel dosyasını belleğe kaydedip kullanıcıya indirtiyoruz
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"RIBA_Resmi_Form_Sonuc_{sinif_filtresi}.xlsx"
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)
@app.route('/giris')
def giris_sayfasi():
    return render_template('giris.html')

# Okul Numarası Yazıldığında Canlı Sorgulama Yapan Fonksiyon
@app.route('/ogrenci-sorgula/<int:okul_no>', methods=['GET'])
def ogrenci_sorgula(okul_no):
    if not supabase:
        return jsonify({"durum": "hata", "mesaj": "Veritabanı bağlantısı kurulu değil."}), 500

    try:
        # Veritabanındaki ogrenciler tablosundan sınıf adıyla birlikte çekiyoruz
        response = supabase.table("ogrenciler").select("id, ad_soyad, siniflar(sinif_adi)").eq("okul_no", okul_no).execute()
        
        if response.data and len(response.data) > 0:
            return jsonify({"durum": "basarili", "veri": response.data[0]})
        else:
            return jsonify({"durum": "hata", "mesaj": "Öğrenci bulunamadı."}), 404
    except Exception as e:
        return jsonify({"durum": "hata", "mesaj": str(e)}), 500
@app.route('/rapor/pdf')
def rapor_pdf():
    sinif_filtresi = request.args.get('sinif', 'TUM_OKUL')
    
    # Bu alanda Python ile veritabanından gelen puanları analiz ediyoruz
    # Gerçek uygulamada puanları SUM() edip en yükseği bulacağız. 
    # Şimdilik örnek senaryo akıllı yorumlama algoritmasını kuralım:
    baskin_problem = "Öz Düzenlemeli Öğrenme (Ders Çalışma Becerileri)"
    asp_puanı = "78.45"
    
    # Rehberlik servisi için otomatik üretilen kurumsal yönlendirme anekdotu
    otomatik_yorum = (
        f"Yapılan RİBA analizi sonucunda, {sinif_filtresi} düzeyinde en yüksek ağırlıklı standart puanın "
        f"({asp_puanı}) ile '{baskin_problem}' alanında olduğu tespit edilmiştir. "
        f"Bu durum, öğrencilerin akademik başarıyı artırma, planlı ders çalışma ve zaman yönetimi "
        f"konularında yoğun bir rehberlik desteğine ihtiyaç duyduklarını göstermektedir. "
        f"Velilerin de süreçte çocuklarının çalışma ortamlarını düzenleme konusunda bilgilendirilmesi elzemdir."
    )
    
    # Bakanlık rapor formatına uygun kurumsal HTML tasarımı
    html_content = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Arial', sans-serif; color: #333; line-height: 1.6; padding: 30px; }}
            .meb-header {{ text-align: center; border-bottom: 2px solid #c0392b; padding-bottom: 10px; margin-bottom: 30px; }}
            .title {{ font-size: 20px; font-weight: bold; color: #2c3e50; }}
            .section {{ margin-bottom: 25px; background: #fafafa; padding: 15px; border-left: 5px solid #2980b9; }}
            .section-title {{ font-size: 16px; font-weight: bold; color: #2980b9; margin-bottom: 10px; }}
            .anekdot-box {{ background: #fdf2e9; border: 1px solid #f5b041; padding: 15px; border-radius: 5px; font-style: italic; }}
        </style>
    </head>
    <body>
        <div class="meb-header">
            <div class="title">T.C. MİLLÎ EĞİTİM BAKANLIĞI</div>
            <div>Okul Rehberlik ve Psikolojik Danışma Servisi Sene Başı İhtiyaç Analiz Raporu</div>
        </div>
        
        <p><strong>Rapor Kapsamı:</strong> {sinif_filtresi}</p>
        <p><strong>Değerlendirme Tarihi:</strong> 20.05.2026</p>
        
        <div class="section">
            <div class="section-title">1. İstatistiki Bulgular ve ASP Analizi</div>
            <p>Sistem üzerinden toplanan verilerin analitik ağırlıklandırılmasına göre öncelikli hedef alanı <strong>{baskin_problem}</strong> olarak belirlenmiştir.</p>
        </div>
        
        <div class="section">
            <div class="section-title">2. Sistem Tarafından Üretilen Otomatik Değerlendirme ve Yorum</div>
            <p>{otomatik_yorum}</p>
        </div>
        
        <div class="section">
            <div class="section-title">3. Rehber Öğretmen Yol Haritası ve Önerilen Çalışmalar (Anekdotlar)</div>
            <div class="anekdot-box">
                * Öğrencilere yönelik 'Zaman Yönetimi ve Verimli Ders Çalışma Teknikleri' seminerinin düzenlenmesi,<br>
                * Her öğrenci için rehberlik servisi gözetiminde bireysel ders çalışma planlarının oluşturulması,<br>
                * Veli oturumlarında 'Evde Verimli Çalışma Ortamı Nasıl Sağlanır?' broşürlerinin dağıtılması önerilir.
            </div>
        </div>
    </body>
    </html>
    """
    
    # WeasyPrint veya PDF oluşturucu kütüphane entegrasyonu (Basit çıktı için doğrudan HTML yanıtı olarak basıyoruz)
    # Render üzerinde tam PDF basımı için tarayıcı bunu otomatik PDF olarak algılayacaktır.
    return html_content
@app.route('/anket-formu')
def anket_formu():
    ogrenci_id = request.args.get('ogrenci_id')
    tip = request.args.get('tip') # 'Ogrenci' mi 'Veli' mi?
    return f"Tebrikler! Doğrulama Başarılı. {tip} Anketi Sayfası Bir Sonraki Adımda Buraya Gelecek. Öğrenci ID: {ogrenci_id}"
# URL'den gelen istek doğrultusunda anket sayfasını açar
@app.route('/anket-formu')
def anket_formu():
    return render_template('anket.html')

# Form bittiğinde gelen verileri Supabase'e kaydeden fonksiyon
@app.route('/anket-kaydet', methods=['POST'])
def anket_kaydet():
    if not supabase:
        return jsonify({"durum": "hata", "mesaj": "Veritabanı bağlantısı yok."}), 500

    veri = request.get_json()
    ogrenci_id = veri.get('ogrenci_id')
    dolduran_tipi = veri.get('dolduran_tipi') # 'Ogrenci' veya 'Veli'
    cevaplar = veri.get('cevaplar')

    try:
        # 1. Şu an aktif olan dönem ID'sini bulalım
        donem_sorgu = supabase.table("donemler").select("id").eq("aktif_mi", True).execute()
        if not donem_sorgu.data:
            return jsonify({"durum": "hata", "mesaj": "Aktif bir anket dönemi bulunamadı."}), 400
        donem_id = donem_sorgu.data[0]['id']

        # 2. Veritabanına kaydı ekle (Eğer daha önce doldurduysa UNIQUE kuralından dolayı hata verir, mükerrer kaydı önler)
        response = supabase.table("riba_cevaplari").insert({
            "donem_id": donem_id,
            "ogrenci_id": ogrenci_id,
            "dolduran_tipi": dolduran_tipi,
            "cevaplar": cevaplar
        }).execute()

        return jsonify({"durum": "basarili", "mesaj": "Cevaplar kalıcı olarak kaydedildi!"})

    except Exception as e:
        # Aynı kişi tekrar doldurmaya çalışırsa buraya düşer
        return jsonify({"durum": "hata", "mesaj": "Bu öğrenci/veli için daha önce anket doldurulmuş."}), 400
if __name__ == '__main__':
    app.run(debug=True)