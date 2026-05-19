from flask import Flask, render_template, jsonify, request
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