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

if __name__ == '__main__':
    app.run(debug=True)