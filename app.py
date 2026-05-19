from flask import Flask, render_template

app = Flask(__name__)

@app.route('/')
def ana_sayfa():
    return render_template('index.html')

@app.route('/giris')
def giris_sayfasi():
    return "Giriş ekranı (Okul No yazılacak yer) bir sonraki adımda buraya eklenecek!"

if __name__ == '__main__':
    app.run(debug=True)