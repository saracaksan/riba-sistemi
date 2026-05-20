from flask import Flask, render_template, jsonify, request, session, send_file
from supabase import create_client, Client
import os
import io
import csv
import openpyxl
from openpyxl.styles import Font, PatternFill

sabit_klasor_yolu = os.path.dirname(os.path.abspath(__file__))
template_yonlendirici = os.path.join(sabit_klasor_yolu, 'templates')

app = Flask(__name__, template_folder=template_yonlendirici)
app.secret_key = "ilce_mem_riba_sistem_guvenlik_kilidi_2026_gazi"

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://pmlgahbdpzlhzlbvfpxf.supabase.co").strip()
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBtbGdhaGJkcHpsaHpsYnZmcHhmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkyMjI4MzQsImV4cCI6MjA5NDc5ODgzNH0.6-OdE2chR29IJHKV1lCHYDkxE5HkkMtvMmXEkFYnbH0").strip()
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

ILCE_USER = "Gazi"
ILCE_PASS = "Gazi.47"

# ================= 🌐 SAYFA YÖNLENDİRMELERİ =================
@app.route('/')
def ana_sayfa(): return render_template('index.html')

@app.route('/anket')
def anket_formu(): return render_template('anket.html')

@app.route('/panel/ilce-mem')
def panel_ilce(): 
    if session.get('rol') != 'ilce': return redirect('/')
    return render_template('panel_ilce.html')

@app.route('/panel/okul-idare')
def panel_okul():
    if session.get('rol') != 'okul': return redirect('/')
    return render_template('panel_okul.html')

@app.route('/panel/ogretmen')
def panel_ogretmen():
    if session.get('rol') != 'ogretmen': return redirect('/')
    return render_template('panel_ogretmen.html')

# ================= 🔐 KİMLİK DOĞRULAMA =================
@app.route('/api/personel-login', methods=['POST'])
def api_personel_login():
    veri = request.get_json()
    k_adi = str(veri.get('kullanici_adi', '')).strip()
    sifre = str(veri.get('sifre', '')).strip()
    
    if k_adi == ILCE_USER and sifre == ILCE_PASS:
        session.clear(); session['rol'] = 'ilce'
        return jsonify({"durum": "basarili", "yonlendir": "/panel/ilce-mem"})
    
    try:
        res = supabase.table("kullanicilar").select("*, okullar(okul_adi, kademe)").eq("kullanici_adi", k_adi).eq("sifre", sifre).execute()
        if res.data:
            kul = res.data[0]
            session.clear(); session['rol'] = kul['rol']; session['okul_id'] = kul['okul_id']; session['kullanici_adi'] = kul['kullanici_adi']
            if kul['rol'] == 'okul': return jsonify({"durum": "basarili", "yonlendir": "/panel/okul-idare"})
            if kul['rol'] == 'ogretmen': return jsonify({"durum": "basarili", "yonlendir": "/panel/ogretmen"})
        return jsonify({"durum": "hata", "mesaj": "Kullanıcı adı veya şifre hatalı!"}), 401
    except Exception as e: return jsonify({"durum": "hata", "mesaj": str(e)}), 500

@app.route('/api/ogrenci-giris', methods=['POST'])
def api_ogrenci_giris():
    veri = request.get_json()
    try:
        res = supabase.table("ogrenciler").select("id, ad, soyad, siniflar(sinif_adi)").eq("okul_id", veri.get('okul_id')).eq("okul_no", veri.get('okul_no')).execute()
        if res.data:
            ogr = res.data[0]
            tam_ad = f"{ogr.get('ad','')} {ogr.get('soyad','')}".strip()
            return jsonify({"durum": "basarili", "id": ogr['id'], "ad_soyad": tam_ad, "sinif": ogr['siniflar']['sinif_adi']})
        return jsonify({"durum": "hata", "mesaj": "Kayıt bulunamadı!"}), 404
    except: return jsonify({"durum": "hata"}), 500

# ================= 🏛️ İLÇE MEM: OKUL İŞLEMLERİ =================
@app.route('/api/ilce/okul-ekle', methods=['POST'])
def api_okul_ekle():
    veri = request.get_json()
    try:
        y_okul = supabase.table("okullar").insert({"okul_adi": veri['okul_adi'], "kademe": veri['kademe'], "kurum_kodu": str(veri['kurum_kodu'])}).execute()
        supabase.table("kullanicilar").insert({"kullanici_adi": str(veri['kurum_kodu']), "sifre": veri['sifre'], "rol": "okul", "okul_id": y_okul.data[0]['id']}).execute()
        return jsonify({"durum": "basarili", "mesaj": "Okul eklendi!"})
    except: return jsonify({"durum": "hata"})

@app.route('/api/ilce/okul-guncelle/<int:id>', methods=['PUT'])
def api_okul_guncelle(id):
    veri = request.get_json()
    try:
        supabase.table("okullar").update({"okul_adi": veri['okul_adi'], "kademe": veri['kademe'], "kurum_kodu": str(veri['kurum_kodu'])}).eq("id", id).execute()
        supabase.table("kullanicilar").update({"kullanici_adi": str(veri['kurum_kodu']), "sifre": veri['sifre']}).eq("okul_id", id).eq("rol", "okul").execute()
        return jsonify({"durum": "basarili"})
    except: return jsonify({"durum": "hata"})

@app.route('/api/ilce/okul-sil/<int:id>', methods=['DELETE'])
def api_okul_sil(id):
    try: supabase.table("okullar").delete().eq("id", id).execute(); return jsonify({"durum": "basarili"})
    except: return jsonify({"durum": "hata"})

@app.route('/api/okullari-getir')
@app.route('/api/ilce/okullari-getir')
def api_okullari_listele():
    try:
        res = supabase.table("okullar").select("id, okul_adi, kademe, kurum_kodu, kullanicilar(sifre)").execute()
        return jsonify([{"id": o['id'], "okul_adi": o['okul_adi'], "kademe": o['kademe'], "kurum_kodu": o['kurum_kodu'], "sifre": o['kullanicilar'][0]['sifre'] if o.get('kullanicilar') else ""} for o in res.data])
    except: return jsonify([])

# ================= 👨‍🏫 OKUL İDARESİ: ÖĞRETMEN VE ŞUBE YÖNETİMİ =================
@app.route('/api/okul/ogretmen-ekle', methods=['POST'])
def api_ogretmen_ekle():
    veri = request.get_json(); okul_id = session.get('okul_id')
    try:
        supabase.table("kullanicilar").insert({
            "kullanici_adi": veri['ad_soyad'], "sifre": veri['sifre'], "rol": "ogretmen", 
            "okul_id": okul_id, "ad_soyad": veri['ad_soyad'], "atanan_siniflar": veri.get('atanan_siniflar', 'TÜMÜ')
        }).execute()
        return jsonify({"durum": "basarili", "mesaj": "Öğretmen eklendi ve şubeleri atandı!"})
    except Exception as e: return jsonify({"durum": "hata", "mesaj": str(e)})

@app.route('/api/okul/ogretmenleri-getir')
def api_ogretmenleri_getir():
    try: return jsonify(supabase.table("kullanicilar").select("id, ad_soyad, sifre, atanan_siniflar, anket_doldurdu").eq("okul_id", session.get('okul_id')).eq("rol", "ogretmen").execute().data)
    except: return jsonify([])

@app.route('/api/okul/ogretmen-sil/<int:id>', methods=['DELETE'])
def api_ogretmen_sil(id):
    try: supabase.table("kullanicilar").delete().eq("id", id).execute(); return jsonify({"durum": "basarili"})
    except: return jsonify({"durum": "hata"})

# ================= 📋 OKUL İDARESİ: ÖĞRENCİ LİSTELEME =================
@app.route('/api/okul/ogrencileri-getir')
def api_ogrencileri_getir():
    try: return jsonify(supabase.table("ogrenciler").select("id, okul_no, ad, soyad, cinsiyet, siniflar(sinif_adi)").eq("okul_id", session.get('okul_id')).order("okul_no").execute().data)
    except: return jsonify([])

@app.route('/api/okul/ogrenci-ekle-manuel', methods=['POST'])
def api_ogrenci_ekle_manuel():
    veri = request.get_json(); okul_id = session.get('okul_id')
    try:
        s_k = supabase.table("siniflar").select("id").eq("okul_id", okul_id).eq("sinif_adi", veri['sinif_adi']).execute()
        s_id = s_k.data[0]['id'] if s_k.data else supabase.table("siniflar").insert({"okul_id": okul_id, "sinif_adi": veri['sinif_adi']}).execute().data[0]['id']
        supabase.table("ogrenciler").insert({"okul_id": okul_id, "sinif_id": s_id, "okul_no": int(veri['okul_no']), "ad": veri['ad'], "soyad": veri['soyad'], "cinsiyet": veri.get('cinsiyet','Belirtilmemiş')}).execute()
        return jsonify({"durum": "basarili", "mesaj": "Öğrenci Eklendi!"})
    except: return jsonify({"durum": "hata"})

@app.route('/api/okul/ogrenci-guncelle/<int:id>', methods=['PUT'])
def api_ogrenci_guncelle(id):
    veri = request.get_json(); okul_id = session.get('okul_id')
    try:
        s_k = supabase.table("siniflar").select("id").eq("okul_id", okul_id).eq("sinif_adi", veri['sinif_adi']).execute()
        s_id = s_k.data[0]['id'] if s_k.data else supabase.table("siniflar").insert({"okul_id": okul_id, "sinif_adi": veri['sinif_adi']}).execute().data[0]['id']
        supabase.table("ogrenciler").update({"okul_no": int(veri['okul_no']), "ad": veri['ad'], "soyad": veri['soyad'], "cinsiyet": veri['cinsiyet'], "sinif_id": s_id}).eq("id", id).execute()
        return jsonify({"durum": "basarili", "mesaj": "Güncellendi!"})
    except: return jsonify({"durum": "hata"})

@app.route('/api/okul/ogrenci-sil/<int:id>', methods=['DELETE'])
def api_ogrenci_sil(id):
    try: supabase.table("ogrenciler").delete().eq("id", id).execute(); return jsonify({"durum": "basarili"})
    except: return jsonify({"durum": "hata"})

# ================= ⚡ IŞIK HIZINDA TOPLU İŞLEM MOTORLARI (BULK) =================
@app.route('/api/okul/ogrenci-toplu-sinif-guncelle', methods=['POST'])
def api_ogrenci_toplu_sinif_guncelle():
    veri = request.get_json()
    idler = veri.get('idler', [])
    yeni_sinif_adi = veri.get('yeni_sinif_adi')
    okul_id = session.get('okul_id')

    if not idler or not yeni_sinif_adi: 
        return jsonify({"durum": "hata", "mesaj": "Eksik bilgi gönderildi."})

    try:
        s_k = supabase.table("siniflar").select("id").eq("okul_id", okul_id).eq("sinif_adi", yeni_sinif_adi).execute()
        s_id = s_k.data[0]['id'] if s_k.data else supabase.table("siniflar").insert({"okul_id": okul_id, "sinif_adi": yeni_sinif_adi}).execute().data[0]['id']
        supabase.table("ogrenciler").update({"sinif_id": s_id}).in_("id", idler).execute()
        return jsonify({"durum": "basarili", "mesaj": f"{len(idler)} öğrencinin sınıfı ışık hızında '{yeni_sinif_adi}' olarak güncellendi!"})
    except Exception as e: return jsonify({"durum": "hata", "mesaj": str(e)})

@app.route('/api/okul/ogrenci-toplu-sil', methods=['POST'])
def api_ogrenci_toplu_sil():
    veri = request.get_json()
    idler = veri.get('idler', [])
    if not idler: return jsonify({"durum": "hata", "mesaj": "Silinecek öğrenci bulunamadı."})
    try:
        supabase.table("ogrenciler").delete().in_("id", idler).execute()
        return jsonify({"durum": "basarili", "mesaj": f"{len(idler)} öğrenci tek seferde başarıyla silindi!"})
    except Exception as e: return jsonify({"durum": "hata", "mesaj": str(e)})

# ================= 📥 ŞABLON İNDİRME VE YÜKLEME =================
@app.route('/api/okul/sablon-indir')
def api_okul_sablon_indir():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ogrenciler"
    ws.append(["Okul No", "Adı", "Soyadı", "Cinsiyeti", "Sınıfı"])
    ws.append([101, "Örnek Ad", "Örnek Soyad", "Erkek", "5/A"])
    for col in range(1, 6):
        ws.cell(row=1, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=1, column=col).fill = openpyxl.styles.PatternFill(start_color="1E3A8A", fill_type="solid")
    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name="Okul_Veri_Yukleme_Sablonu.xlsx")

@app.route('/api/sablon-liste-yukle', methods=['POST'])
def api_sablon_yukle():
    if 'file' not in request.files: return jsonify({"durum": "hata", "mesaj": "Dosya yok"}), 400
    file = request.files['file']
    okul_id = session.get('okul_id')
    filename = file.filename.lower()
    eklenen = 0
    try:
        if filename.endswith('.csv'):
            stream = io.StringIO(file.stream.read().decode("UTF8"), newline=None)
            csv_reader = csv.reader(stream, delimiter=',')
            next(csv_reader, None)
            for row in csv_reader:
                if len(row) < 5: continue
                no, ad, soyad, cinsiyet, sinif = [str(e).strip() for e in row[:5]]
                if not no.isdigit(): continue
                s_k = supabase.table("siniflar").select("id").eq("okul_id", okul_id).eq("sinif_adi", sinif).execute()
                s_id = s_k.data[0]['id'] if s_k.data else supabase.table("siniflar").insert({"okul_id": okul_id, "sinif_adi": sinif}).execute().data[0]['id']
                try: supabase.table("ogrenciler").insert({"okul_id": okul_id, "sinif_id": s_id, "okul_no": int(no), "ad": ad, "soyad": soyad, "cinsiyet": cinsiyet}).execute(); eklenen += 1
                except: pass
                
        elif filename.endswith(('.xls', '.xlsx')):
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0] or not row[1] or not row[4]: continue
                no = str(row[0]).strip(); ad = str(row[1]).strip(); soyad = str(row[2]).strip() if row[2] else ""; cinsiyet = str(row[3]).strip() if row[3] else "Belirtilmemiş"; sinif = str(row[4]).strip()
                if not no.isdigit(): continue
                s_k = supabase.table("siniflar").select("id").eq("okul_id", okul_id).eq("sinif_adi", sinif).execute()
                s_id = s_k.data[0]['id'] if s_k.data else supabase.table("siniflar").insert({"okul_id": okul_id, "sinif_adi": sinif}).execute().data[0]['id']
                try: supabase.table("ogrenciler").insert({"okul_id": okul_id, "sinif_id": s_id, "okul_no": int(no), "ad": ad, "soyad": soyad, "cinsiyet": cinsiyet}).execute(); eklenen += 1
                except: pass
                
        return jsonify({"durum": "basarili", "mesaj": f"{eklenen} öğrenci başarıyla sisteme aktarıldı!"})
    except Exception as e: return jsonify({"durum": "hata", "mesaj": str(e)}), 500

# ================= 📝 ANKET KAYDETME MOTORU =================
@app.route('/api/anket-kaydet', methods=['POST'])
def api_anket_kaydet():
    veri = request.get_json()
    try:
        supabase.table("riba_cevaplari").insert({
            "ogrenci_id": veri.get('ogrenci_id'),
            "dolduran_tipi": veri['dolduran_tipi'],
            "cevaplar": veri['cevaplar']
        }).execute()
        
        # Öğretmen anket doldurduysa hesabını işaretle
        if veri['dolduran_tipi'] == 'Ogretmen':
            k_adi = session.get('kullanici_adi')
            if k_adi:
                supabase.table("kullanicilar").update({"anket_doldurdu": True}).eq("kullanici_adi", k_adi).execute()
                
        return jsonify({"durum": "basarili"})
    except Exception as e:
        return jsonify({"durum": "hata", "mesaj": str(e)})

# ================= 👨‍🏫 ÖĞRETMEN TAKİP MOTORU =================
@app.route('/api/ogretmen/sinif-takip')
def api_ogretmen_takip():
    okul_id = session.get('okul_id')
    k_adi = session.get('kullanici_adi')
    try:
        kul = supabase.table("kullanicilar").select("atanan_siniflar").eq("kullanici_adi", k_adi).execute().data[0]
        atanan_str = kul.get('atanan_siniflar', 'TÜMÜ')
        atanan_liste = [s.strip() for s in atanan_str.split(',')] if atanan_str != 'TÜMÜ' else []
        
        ogrenciler = supabase.table("ogrenciler").select("id, okul_no, ad, soyad, siniflar(sinif_adi)").eq("okul_id", okul_id).execute().data
        if atanan_str != 'TÜMÜ':
            ogrenciler = [o for o in ogrenciler if o['siniflar']['sinif_adi'] in atanan_liste]
            
        cevaplar = supabase.table("riba_cevaplari").select("ogrenci_id, dolduran_tipi").execute().data
        o_cev = {c['ogrenci_id']: True for c in cevaplar if c['dolduran_tipi'] == 'Ogrenci'}
        v_cev = {c['ogrenci_id']: True for c in cevaplar if c['dolduran_tipi'] == 'Veli'}
        
        liste = []
        for o in ogrenciler:
            liste.append({
                "no": o['okul_no'], "ad_soyad": f"{o['ad']} {o['soyad']}", "sinif": o['siniflar']['sinif_adi'],
                "ogrenci_durum": "Tamamlandı" if o['id'] in o_cev else "Eksik",
                "veli_durum": "Tamamlandı" if o['id'] in v_cev else "Eksik"
            })
        return jsonify({"ogrenciler": liste, "atanan_siniflar": atanan_str})
    except Exception as e: return jsonify({"durum": "hata", "mesaj": str(e)}), 500

# ================= 📈 KATILIM RADARI (İDARE VE İLÇE) =================
@app.route('/api/takip/katilim-durumu')
def api_katilim_durumu():
    rol = session.get('rol'); okul_id = request.args.get('okul_id') if rol == 'ilce' else session.get('okul_id')
    try:
        q = supabase.table("ogrenciler").select("id, okul_no, ad, soyad, sinif_id, siniflar(sinif_adi)")
        if okul_id: q = q.eq("okul_id", okul_id)
        ogrenciler = q.execute().data
        cevaplar = supabase.table("riba_cevaplari").select("ogrenci_id, dolduran_tipi").execute().data
        
        o_cev = {c['ogrenci_id']: True for c in cevaplar if c['dolduran_tipi'] == 'Ogrenci'}
        v_cev = {c['ogrenci_id']: True for c in cevaplar if c['dolduran_tipi'] == 'Veli'}
        s_mevcut = {}; s_katilim = {"ogrenci": {}, "veli": {}}
        for o in ogrenciler:
            s_adi = o['siniflar']['sinif_adi']
            s_mevcut[s_adi] = s_mevcut.get(s_adi, 0) + 1
            if o['id'] in o_cev: s_katilim["ogrenci"][s_adi] = s_katilim["ogrenci"].get(s_adi, 0) + 1
            if o['id'] in v_cev: s_katilim["veli"][s_adi] = s_katilim["veli"].get(s_adi, 0) + 1

        eksik = [{"sinif_adi": s, "mevcut": m, "ogrenci_katilim": s_katilim["ogrenci"].get(s,0), "veli_katilim": s_katilim["veli"].get(s,0)} for s, m in s_mevcut.items() if s_katilim["ogrenci"].get(s,0) < m or s_katilim["veli"].get(s,0) < m]
        return jsonify({"eksik_siniflar": eksik, "toplam_mevcut": len(ogrenciler), "toplam_ogrenci_katilim": len(o_cev), "toplam_veli_katilim": len(v_cev)})
    except: return jsonify({"durum": "hata"}), 500

# ================= 🏆 MEB ORİJİNAL RİBA SONUÇ RAPORLAMA MOTORU =================
@app.route('/api/rapor/analiz-veri')
def api_rapor_analiz_veri():
    rol = session.get('rol')
    kademe_filtre = request.args.get('kademe')
    okul_id = request.args.get('okul_id') if rol == 'ilce' else session.get('okul_id')
    sinif_id = request.args.get('sinif_id') 
    
    try:
        # MEB riba_formlari.xlsx ve Sonuc.csv tablosundan çekilen 36 Orijinal Hedef 
        hedefler_map = {
            "Problem Çözme Becerileri": ["M12"], "Öz Düzenlemeli Öğrenme": ["M30"], "Karar Verme Becerisi": ["M08"],
            "Motivasyon/Devamsızlığı Önleme": ["M29"], "Sosyal Beceriler": ["M07"], "Üst Öğrenime Geçiş Sınavları": ["M36"],
            "Okul ve Çevresindeki Sosyokültürel İmkanlar": ["M34", "M35"], "Atılganlık": ["M15"], "Yardım Arama": ["M48"],
            "Öz Disiplin Geliştirme": ["M04"], "Duygu Düzenleme": ["M42"], "Bilinçli Teknoloji Kullanımı": ["M11"],
            "Zaman Yönetimi/Öz Düzenlemeli Öğrenme": ["M31"], "Gelişim Dönemi Özellikleri": ["M13"], "Hak ve Sorumluluklarını Bilme": ["M14"],
            "Meslek ile İlgi, Değer, Yetenek ve Kişisel Özellik": ["M23", "M22"], "Üst Öğrenim Kurumlarının Tanıtılması": ["M32"],
            "Psikolojik Sağlamlık": ["M47"], "Dikkat Geliştirme Çalışmaları": ["M33"], "Bağımlılıkla Mücadele": ["M16"],
            "İletişim Becerileri": ["M03", "M10"], "Mesleki Benlik": ["M39"], "Çatışma Çözme Becerileri": ["M19"],
            "Bireysel Farklılıklara Saygı": ["M45"], "Aile İçi İletişim": ["M46"], "Öfke Yönetimi": ["M06"],
            "Sınır Koyma": ["M09"], "Okula ve Çevreye Uyum/Okul Kuralları": ["M27"], "Rehberlik ve Psikolojik Danışma Servisinin Tanıtılması": ["M38"],
            "Özgüven Geliştirme": ["M01"], "Sınav Kaygısı": ["M37"], "İhmal ve İstismardan Korunma": ["M41"],
            "Yaşam Becerileri": ["M18"], "Sağlıklı Yaşam": ["M44"], "Akran Zorbalığı": ["M05"], "Duygu Farkındalığı/Duygu Düzenleme": ["M02"]
        }
        
        frekanslar = {isim: {"Ogrenci": 0, "Veli": 0, "Ogretmen": 0} for isim in hedefler_map.keys()}
        
        sorgu = supabase.table("riba_cevaplari").select("*, ogrenciler(okul_id, sinif_id, okullar(kademe))")
        cevaplar = sorgu.execute().data
        
        for c in cevaplar:
            ogr_baglanti = c.get('ogrenciler')
            if not ogr_baglanti and c['dolduran_tipi'] != 'Ogretmen': continue
            
            if ogr_baglanti:
                if okul_id and str(ogr_baglanti['okul_id']) != str(okul_id): continue
                if sinif_id and str(ogr_baglanti['sinif_id']) != str(sinif_id): continue
                if kademe_filtre and kademe_filtre != "TUM" and ogr_baglanti['okullar']['kademe'] != kademe_filtre: continue
            
            tip = c['dolduran_tipi'] 
            c_json = c.get('cevaplar', {})
            
            for hedef_adi, kodlar in hedefler_map.items():
                for kod in kodlar:
                    if kod in c_json:
                        frekanslar[hedef_adi][tip] += c_json[kod].get("A", 0)

        rapor_matrisi = []
        for hedef_adi, frek in frekanslar.items():
            f_o = frek["Ogrenci"]
            f_v = frek["Veli"]
            f_og = frek["Ogretmen"]
            
            # Ağırlıklandırılmış Standart Puan (ASP) Hesaplaması
            asp = round((f_o * 0.5) + (f_v * 0.3) + (f_og * 0.2), 2)
            
            rapor_matrisi.append({
                "hedef_adi": hedef_adi, "ogrenci_f": f_o, "veli_f": f_v, "ogretmen_f": f_og, "asp": asp
            })
            
        rapor_matrisi.sort(key=lambda x: x['asp'], reverse=True)
        return jsonify(rapor_matrisi)
    except Exception as e: return jsonify({"durum": "hata", "mesaj": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)