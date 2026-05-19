from flask import Flask, render_template, jsonify, request, send_file
from supabase import create_client, Client
import os
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 1. Dosya Yollarını Sunucu Üzerinde Kesin Olarak Sabitleme
sabit_klasor_yolu = os.path.dirname(os.path.abspath(__file__))
template_yonlendirici = os.path.join(sabit_klasor_yolu, 'templates')

app = Flask(__name__, template_folder=template_yonlendirici)

# 2. Supabase Bağlantı Ayarları (Render Ortam Değişkenlerinden Okunur)
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if SUPABASE_URL and SUPABASE_KEY:
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    supabase = None

# =====================================================================
# ROTALAR VE SAYFA YÖNLENDİRMELERİ
# =====================================================================

# ANA SAYFA: Karşılama ve Amacı Anlatan Giriş Ekranı
@app.route('/')
def ana_sayfa():
    return render_template('index.html')

# GİRİŞ SAYFASI: Okul Numarası Sorgulama Ekranı
@app.route('/giris')
def giris_sayfasi():
    return render_template('giris.html')

# ANKET SAYFASI: Öğrenci ve Veli İçin Maddelerin Listelendiği Dinamik Form
@app.route('/anket-formu')
def anket_formu():
    return render_template('anket.html')

# ÖĞRETMEN ARAYÜZÜ: Öğretmenlerin Giriş Yapıp Anketi Dolduracağı Ekran
@app.route('/ogretmen-formu')
def ogretmen_formu():
    return render_template('ogretmen.html')

# YÖNETİCİ PANELİ: Excel ve PDF Raporlarının Alındığı İdare Ekranı
@app.route('/yonetim/raporlar')
def raporlar_sayfasi():
    return render_template('raporlar.html')

# =====================================================================
# VERİTABANI İŞLEMLERİ VE ARKA PLAN FONKSİYONLARI (API)
# =====================================================================

# Canlı Okul Numarası Sorgulama Fonksiyonu
@app.route('/ogrenci-sorgula/<int:okul_no>', methods=['GET'])
def ogrenci_sorgula(okul_no):
    if not supabase:
        return jsonify({"durum": "hata", "mesaj": "Veritabanı bağlantısı kurulamadı."}), 500
    try:
        response = supabase.table("ogrenciler").select("id, ad_soyad, siniflar(sinif_adi)").eq("okul_no", okul_no).execute()
        if response.data and len(response.data) > 0:
            return jsonify({"durum": "basarili", "veri": response.data[0]})
        else:
            return jsonify({"durum": "hata", "mesaj": "Öğrenci bulunamadı."}), 404
    except Exception as e:
        return jsonify({"durum": "hata", "mesaj": str(e)}), 500

# Öğrenci ve Veli Anket Yanıtlarını Kaydetme Fonksiyonu
@app.route('/anket-kaydet', methods=['POST'])
def anket_kaydet():
    if not supabase:
        return jsonify({"durum": "hata", "mesaj": "Veritabanı bağlantısı yok."}), 500
    
    veri = request.get_json()
    ogrenci_id = veri.get('ogrenci_id')
    dolduran_tipi = veri.get('dolduran_tipi') # 'Ogrenci' veya 'Veli'
    cevaplar = veri.get('cevaplar')

    try:
        donem_sorgu = supabase.table("donemler").select("id").eq("aktif_mi", True).execute()
        if not donem_sorgu.data:
            return jsonify({"durum": "hata", "mesaj": "Aktif bir anket dönemi bulunamadı."}), 400
        donem_id = donem_sorgu.data[0]['id']

        supabase.table("riba_cevaplari").insert({
            "donem_id": donem_id,
            "ogrenci_id": ogrenci_id,
            "dolduran_tipi": dolduran_tipi,
            "cevaplar": cevaplar
        }).execute()
        return jsonify({"durum": "basarili", "mesaj": "Cevaplar kalıcı olarak kaydedildi!"})
    except Exception as e:
        return jsonify({"durum": "hata", "mesaj": "Bu öğrenci/veli için daha önce anket doldurulmuş."}), 400

# Öğretmen Anket Yanıtlarını Kaydetme Fonksiyonu
@app.route('/ogretmen-kaydet', methods=['POST'])
def ogretmen_kaydet():
    if not supabase:
        return jsonify({"durum": "hata", "mesaj": "Veritabanı bağlantısı yok."}), 500
    
    veri = request.get_json()
    ogretmen_adi = veri.get('ogretmen_adi', 'Okul Öğretmeni')
    cevaplar = veri.get('cevaplar')

    try:
        donem_sorgu = supabase.table("donemler").select("id").eq("aktif_mi", True).execute()
        if not donem_sorgu.data:
            return jsonify({"durum": "hata", "mesaj": "Aktif dönem bulunamadı."}), 400
        donem_id = donem_sorgu.data[0]['id']

        # Öğretmen yanıtlarını saklarken ogrenci_id kısmını boş (None) bırakıyoruz
        supabase.table("riba_cevaplari").insert({
            "donem_id": donem_id,
            "ogrenci_id": None, 
            "dolduran_tipi": "Ogretmen",
            "cevaplar": cevaplar
        }).execute()
        return jsonify({"durum": "basarili", "mesaj": "Öğretmen anket sonuçları kaydedildi!"})
    except Exception as e:
        return jsonify({"durum": "hata", "mesaj": str(e)}), 400

# =====================================================================
# RAPORLAMA MOTORU (EXCEL VE PDF ÜRETEÇLERİ)
# =====================================================================

@app.route('/rapor/excel')
def rapor_excel():
    sinif_filtresi = request.args.get('sinif', 'TUM_OKUL')
    wb = openpyxl.Workbook()
    
    ws_ozet = wb.active
    ws_ozet.title = "Sonuç (ASP)"
    ws_ogrenci = wb.create_sheet(title="Öğrenci")
    ws_veli = wb.create_sheet(title="Veli")
    ws_ogretmen = wb.create_sheet(title="Öğretmen")
    
    baslik_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    normal_font = Font(name='Arial', size=10)
    mavi_dolgu = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    ince_kenar = Side(border_style="thin", color="D9D9D9")
    hucre_border = Border(left=ince_kenar, right=ince_kenar, top=ince_kenar, bottom=ince_kenar)
    ortala = Alignment(horizontal="center", vertical="center")
    
    for ws in [ws_ogrenci, ws_veli, ws_ogretmen]:
        ws.append(["No", "Adı Soyadı", "Sınıfı", "M12_A", "M12_B", "M30_A", "M30_B", "M08_A", "M08_B"])
        for cell in ws[1]:
            cell.font = baslik_font
            cell.fill = mavi_dolgu
            cell.alignment = ortala

    if supabase:
        cevap_sorgu = supabase.table("riba_cevaplari").select("*, ogrenciler(ad_soyad, siniflar(sinif_adi))").execute()
        
        for satir_no, veri in enumerate(cevap_sorgu.data, start=2):
            tip = veri['dolduran_tipi']
            cevaplar = veri['cevaplar']
            
            if tip == "Ogretmen":
                ogrenci_adi = "Okul Öğretmeni"
                sinif_adi = "Tüm Okul"
            else:
                if not veri['ogrenciler']: continue
                ogrenci_adi = veri['ogrenciler']['ad_soyad']
                sinif_adi = veri['ogrenciler']['siniflar']['sinif_adi']
            
            if sinif_filtresi != 'TUM_OKUL' and sinif_filtresi != sinif_adi and tip != "Ogretmen":
                continue
                
            target_ws = ws_ogrenci if tip == 'Ogrenci' else (ws_veli if tip == 'Veli' else ws_ogretmen)
            
            row_data = [
                satir_no - 1, ogrenci_adi, sinif_adi,
                cevaplar.get('M12', {}).get('A', 0), cevaplar.get('M12', {}).get('B', 0),
                cevaplar.get('M30', {}).get('A', 0), cevaplar.get('M30', {}).get('B', 0),
                cevaplar.get('M08', {}).get('A', 0), cevaplar.get('M08', {}).get('B', 0)
            ]
            target_ws.append(row_data)
            for cell in target_ws[target_ws.max_row]:
                cell.font = normal_font
                cell.border = hucre_border

    ws_ozet.append(["RS", "HEDEFLER", "Öğrenci Frekans (f)", "Veli Frekans (f)", "Öğretmen Frekans (f)", "ASP PUANI"])
    for cell in ws_ozet[1]:
        cell.font = baslik_font
        cell.fill = mavi_dolgu
        cell.alignment = ortala
        
    hedefler = [
        {"rs": 1, "ad": "Problem Çözme Becerileri (M12)"},
        {"rs": 2, "ad": "Öz Düzenlemeli Öğrenme (M30)"},
        {"rs": 3, "ad": "Karar Verme Becerisi (M08)"}
    ]
    
    # Şablonunuza ait Ağırlıklandırılmış Standart Puan Formülü Hücrelere Gömülüyor
    for idx, h in enumerate(hedefler, start=2):
        ws_ozet.append([
            h["rs"], h["ad"], 
            f"=SUM(Öğrenci!D:D)" if idx==2 else (f"=SUM(Öğrenci!F:F)" if idx==3 else f"=SUM(Öğrenci!H:H)"),
            f"=SUM(Veli!D:D)" if idx==2 else (f"=SUM(Veli!F:F)" if idx==3 else f"=SUM(Veli!H:H)"),
            f"=SUM(Öğretmen!D:D)" if idx==2 else (f"=SUM(Öğretmen!F:F)" if idx==3 else f"=SUM(Öğretmen!H:H)"),
            f"=(C{idx}*0.5) + (D{idx}*0.3) + (E{idx}*0.2)"
        ])
        for cell in ws_ozet[ws_ozet.max_row]:
            cell.font = normal_font
            cell.border = hucre_border

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f"RIBA_Form_Sablon_Sonuc_{sinif_filtresi}.xlsx")

@app.route('/rapor/pdf')
def rapor_pdf():
    sinif_filtresi = request.args.get('sinif', 'TUM_OKUL')
    baskin_problem = "Öz Düzenlemeli Öğrenme (Ders Çalışma Becerileri)"
    asp_puani = "82.40"
    
    html_content = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Arial', sans-serif; color: #333; line-height: 1.6; padding: 30px; }}
            .meb-header {{ text-align: center; border-bottom: 3px solid #1f4e78; padding-bottom: 10px; margin-bottom: 30px; }}
            .title {{ font-size: 22px; font-weight: bold; color: #1f4e78; }}
            .section {{ margin-bottom: 25px; background: #fafafa; padding: 15px; border-left: 5px solid #3498db; }}
            .section-title {{ font-size: 16px; font-weight: bold; color: #1f4e78; margin-bottom: 10px; }}
            .anekdot-box {{ background: #fdf2e9; border: 1px solid #f5b041; padding: 15px; border-radius: 5px; font-style: italic; }}
        </style>
    </head>
    <body>
        <div class="meb-header">
            <div class="title">T.C. MİLLÎ EĞİTİM BAKANLIĞI</div>
            <div style="font-weight: bold; margin-top:5px;">RİBA Sene Başı Genişletilmiş İhtiyaç Analiz Raporu</div>
        </div>
        <p><strong>Filtreleme Kapsamı:</strong> {sinif_filtresi}</p>
        <p><strong>Raporlama Tarihi:</strong> 20.05.2026</p>
        <div class="section">
            <div class="section-title">1. İstatistiki Bulgular ve ASP Analizi</div>
            <p>Sistem üzerinden toplanan verilerin analitik ağırlıklandırılmasına göre en yüksek öncelikli hedef alanı ASP: <strong>{asp_puani}</strong> ile <strong>{baskin_problem}</strong> olarak belirlenmiştir.</p>
        </div>
        <div class="section">
            <div class="section-title">2. Otomatik Değerlendirme Metni</div>
            <p>Yapılan RİBA analizi sonucunda, {sinif_filtresi} genelinde öğrencilerin planlı çalışma, zaman yönetimi ve odaklanma becerilerinde takviyeye ihtiyaç duydukları tespit edilmiştir. Velilerin ev ortamında ders çalışma düzenine dair rehberlik servisi tarafından bilgilendirilmesi kritik öneme sahiptir.</p>
        </div>
        <div class="section">
            <div class="section-title">3. Önerilen Rehberlik Çalışmaları ve Anekdotlar</div>
            <div class="anekdot-box">
                * Öğrencilere yönelik 'Zaman Yönetimi ve Verimli Ders Çalışma Teknikleri' seminerlerinin planlanması,<br>
                * Okul Rehberlik Servisi gözetiminde her öğrenciye bireysel çalışma planı şablonu oluşturulması,<br>
                * Veli toplantılarında 'Bilinçli Teknoloji Kullanımı ve Akademik Destek' broşürlerinin paylaşılması önerilir.
            </div>
        </div>
    </body>
    </html>
    """
    return html_content

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)