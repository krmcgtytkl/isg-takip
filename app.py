import os
import io
from flask import Flask, render_template_string, request, redirect, url_for, send_file, session
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from werkzeug.utils import secure_filename
import base64
import qrcode

app = Flask(__name__)
app.config['SECRET_KEY'] = 'isg-gizli-anahtar'

# ---------------- KULLANICI GİRİŞİ ----------------
KULLANICI_ADI = 'admin'
SITE_SIFRESI = '1071isg'
app.config['PERMANENT_SESSION_LIFETIME'] = 86400

# ---------------- VARSAYILAN FİRMA LİSTESİ ----------------
VARSAYILAN_FIRMALAR = [
    "Tat Nişasta.",
    "Sitrik Asit.",
    "Etanol",
    "Tezmak",
]

# ---------------- VERİTABANI BAĞLANTISI ----------------
app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql+psycopg2://postgres.vykhokdrahdftzwlizdu:1071isgkatip@aws-0-eu-west-1.pooler.supabase.com:5432/postgres'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

UPLOAD_FOLDER = os.path.join('static', 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

db = SQLAlchemy(app)

# ---------------- MODELLER ----------------
class Egitim(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    firma_adi = db.Column(db.String(150), nullable=False)
    personel_adi = db.Column(db.String(120), nullable=False)
    ise_giris = db.Column(db.Date, nullable=True)
    gorevi = db.Column(db.String(100), nullable=True)
    resim = db.Column(db.String(200))
    egitim_adi = db.Column(db.String(150), nullable=False)
    egitim_tarihi = db.Column(db.Date, nullable=False)
    gecerlilik_ay = db.Column(db.Integer, default=12)
    sertifika_no = db.Column(db.String(50))

    @property
    def son_gecerlilik(self):
        return self.egitim_tarihi + relativedelta(months=self.gecerlilik_ay)

    @property
    def durum(self):
        kalan = (self.son_gecerlilik - date.today()).days
        if kalan < 0: return "SÜRESİ GEÇMİŞ"
        elif kalan <= 30: return "YAKLAŞIYOR"
        return "GEÇERLİ"

class IsKazasi(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    firma_adi = db.Column(db.String(150), nullable=False)
    personel_adi = db.Column(db.String(120), nullable=False)
    tarih = db.Column(db.Date, nullable=False, default=date.today)
    yer = db.Column(db.String(120))
    aciklama = db.Column(db.Text)
    yaralanma_turu = db.Column(db.String(80))
    kok_neden = db.Column(db.Text)
    duzeltici_faaliyet = db.Column(db.Text)
    sgk_bildirim_tarihi = db.Column(db.Date, nullable=True)
    rapor_gun_sayisi = db.Column(db.Integer, nullable=True)
    kaza_sonrasi_egitim = db.Column(db.Date, nullable=True)

class RiskPlan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    firma_adi = db.Column(db.String(150), nullable=False)
    rd_tehlike_sinifi = db.Column(db.String(50), nullable=True)
    rd_tarihi = db.Column(db.Date, nullable=True)
    rd_gecerlilik = db.Column(db.Date, nullable=True)
    adp_tehlike_sinifi = db.Column(db.String(50), nullable=True)
    adp_tarihi = db.Column(db.Date, nullable=True)
    adp_gecerlilik = db.Column(db.Date, nullable=True)
    egitim_tehlike_sinifi = db.Column(db.String(50), nullable=True)
    egitim_plani_tarihi = db.Column(db.Date, nullable=True)
    egitim_plani_gecerlilik = db.Column(db.Date, nullable=True)
    calisma_tehlike_sinifi = db.Column(db.String(50), nullable=True)
    calisma_plani_tarihi = db.Column(db.Date, nullable=True)
    calisma_plani_gecerlilik = db.Column(db.Date, nullable=True)

class Firma(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    adi = db.Column(db.String(150), nullable=False, unique=True)
    aktif = db.Column(db.Boolean, default=True)
    olusturma = db.Column(db.Date, default=date.today)

# ---------------- FİRMA LİSTESİ HELPER ----------------
def get_firmalar_listesi():
    try:
        firmalar = Firma.query.filter_by(aktif=True).order_by(Firma.adi).all()
        if not firmalar:
            for f_adi in VARSAYILAN_FIRMALAR:
                if not Firma.query.filter_by(adi=f_adi).first():
                    db.session.add(Firma(adi=f_adi))
            db.session.commit()
            firmalar = Firma.query.filter_by(aktif=True).order_by(Firma.adi).all()
        return [f.adi for f in firmalar]
    except Exception as e:
        print(f"⚠️ Firma listesi hatası: {e}")
        return list(VARSAYILAN_FIRMALAR)

# ---------------- HTML ŞABLONU (ORTAK) ----------------
BASE_HTML = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>İSG TAKİP</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css">
    <style>
        body { font-family: 'Poppins', sans-serif; background: #f4f7f6; color: #2c3e50; overflow-x: hidden; }
        .sidebar { height: 100vh; width: 250px; position: fixed; top: 0; left: 0; background: linear-gradient(180deg, #1e3c72 0%, #2a5298 100%); color: white; padding-top: 20px; border-top-right-radius: 20px; border-bottom-right-radius: 20px; box-shadow: 4px 0 15px rgba(0,0,0,0.1); display: flex; flex-direction: column; }
        .sidebar h4 { text-align: center; font-weight: 700; margin-bottom: 20px; letter-spacing: 1px; font-size: 1.2rem; }
        .sidebar a { padding: 12px 25px; text-decoration: none; font-size: 14px; color: rgba(255, 255, 255, 0.8); display: block; transition: 0.3s; }
        .sidebar a:hover, .sidebar a.active { background: rgba(255, 255, 255, 0.1); color: white; border-left: 4px solid #fff; }
        .sidebar a i { margin-right: 10px; }
        .sidebar-menu { flex: 1; overflow-y: auto; }
        .sidebar-bottom { padding-bottom: 20px; }
        .main-content { margin-left: 250px; padding: 30px; min-height: 100vh; background: #f4f7f6; }
        
        .firma-selector { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 15px; padding: 20px 25px; margin-bottom: 25px; color: white; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 8px 25px rgba(102, 126, 234, 0.3); flex-wrap: wrap; gap: 15px; }
        .firma-selector .label { font-size: 0.85rem; opacity: 0.9; margin-bottom: 5px; display: block; }
        .firma-selector select { border-radius: 10px; border: none; padding: 10px 15px; font-size: 1rem; font-weight: 600; min-width: 300px; background: white; color: #2c3e50; }
        .firma-selector .info { font-size: 0.85rem; opacity: 0.9; }
        
        .widget-card { background: white; border-radius: 15px; padding: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.03); transition: transform 0.3s; height: 100%; border-left: 5px solid transparent; }
        .widget-card:hover { transform: translateY(-3px); box-shadow: 0 8px 25px rgba(0,0,0,0.08); }
        .widget-rd { border-left-color: #3498db; }
        .widget-adp { border-left-color: #e74c3c; }
        .widget-egitim-plan { border-left-color: #1abc9c; }
        .widget-calisma-plan { border-left-color: #34495e; }
        .widget-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; }
        .widget-title { font-weight: 600; font-size: 1rem; display: flex; align-items: center; gap: 8px; }
        .widget-icon { font-size: 1.3rem; }
        .mini-card { background: white; border-radius: 10px; padding: 12px 15px; margin-bottom: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.02); display: flex; justify-content: space-between; align-items: center; border-left: 4px solid transparent; }
        .mini-card.danger { border-left-color: #e74c3c; }
        .mini-card.warning { border-left-color: #f39c12; }
        .mini-card-title { font-weight: 600; font-size: 0.85rem; margin-bottom: 2px; }
        .mini-card-sub { font-size: 0.7rem; color: #7f8c8d; }
        .mini-card-date { font-weight: 700; font-size: 0.8rem; }
        .table-custom { background: transparent; font-size: 0.8rem; margin-bottom: 0; }
        .table-custom th { border: none; color: #7f8c8d; font-weight: 500; padding: 8px 5px; }
        .table-custom td { border-bottom: 1px solid #f1f2f6; padding: 8px 5px; vertical-align: middle; }
        .badge-glass { border-radius: 6px; padding: 4px 8px; font-weight: 500; font-size: 0.7rem; }
        .btn-glass { background: #f4f7f6; border: none; border-radius: 8px; padding: 6px 12px; font-size: 0.8rem; font-weight: 500; transition: 0.3s; }
        .btn-glass:hover { background: #e0e6ed; }
        .alert-panel { background: #fff3cd; border-left: 5px solid #ffc107; border-radius: 10px; padding: 15px 20px; margin-bottom: 25px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); }
        .alert-panel-title { font-weight: 700; font-size: 1rem; color: #856404; margin-bottom: 10px; display: flex; align-items: center; gap: 10px; }
        .alert-item { font-size: 0.85rem; padding: 6px 0; border-bottom: 1px dashed rgba(0,0,0,0.1); display: flex; justify-content: space-between; }
        .alert-item:last-child { border-bottom: none; }
        @media (max-width: 768px) {
            .sidebar { width: 100%; height: auto; position: relative; border-radius: 0; }
            .main-content { margin-left: 0; padding: 15px; }
            .firma-selector select { min-width: 100%; }
        }
    </style>
</head>
<body>

    <div class="sidebar">
        <h4>İSG TAKİP</h4>
        <div class="sidebar-menu">
            <a href="/" class="active"><i class="bi bi-grid-1x2-fill"></i> Panel</a>
            <a href="/firma_degistir"><i class="bi bi-arrow-left-right"></i> Firma Değiştir</a>
            <a href="/firma_yonetim"><i class="bi bi-gear-fill"></i> Firma Yönetimi</a>
            <a href="/tum_firmalar"><i class="bi bi-buildings"></i> Tüm Firmalar Özet</a>
            <a href="/personeller"><i class="bi bi-people-fill"></i> Personel Listesi</a>
            <a href="/egitimler"><i class="bi bi-mortarboard-fill"></i> Eğitimler</a>
            <a href="/egitim/ekle"><i class="bi bi-plus-circle-fill"></i> + Eğitim Ekle</a>
            <a href="/kazalar"><i class="bi bi-exclamation-triangle-fill"></i> Kazalar</a>
            <a href="/kaza/ekle"><i class="bi bi-plus-circle-fill"></i> + Kaza Ekle</a>
            <a href="/planlar"><i class="bi bi-journal-check"></i> Tüm Planlar</a>
            <a href="/plan/ekle"><i class="bi bi-plus-circle-fill"></i> + Plan Ekle</a>
        </div>
        <div class="sidebar-bottom" style="border-top: 1px solid rgba(255,255,255,0.2); padding-top: 15px;">
            <a href="/cikis" style="color: #ffb3b3;"><i class="bi bi-box-arrow-right"></i> Çıkış Yap</a>
        </div>
    </div>

    <div class="main-content">
        {% if session.get('firma_adi') %}
        <div class="firma-selector">
            <div>
                <span class="label"><i class="bi bi-building"></i> Aktif Firma</span>
                <select onchange="window.location='/firma_sec/'+encodeURIComponent(this.value)">
                    {% for f in firmalar %}
                        <option value="{{ f }}" {% if session.get('firma_adi') == f %}selected{% endif %}>{{ f }}</option>
                    {% endfor %}
                    <option value="TÜM FİRMALAR" {% if session.get('firma_adi') == 'TÜM FİRMALAR' %}selected{% endif %}>⭐ TÜM FİRMALAR</option>
                </select>
            </div>
            <div class="info">
                <i class="bi bi-info-circle"></i> Veriler sadece seçili firmaya aittir. Değiştirmek için yukarıdaki menüyü kullanın.
            </div>
        </div>
        {% endif %}
        
        {% block content %}{% endblock %}
    </div>

</body>
</html>
"""

# ---------------- GİRİŞ EKRANI ----------------
GIRIS_HTML = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <title>Giriş - İSG TAKİP</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); min-height: 100vh; display: flex; align-items: center; justify-content: center; font-family: 'Poppins', sans-serif; }
        .login-card { background: white; border-radius: 20px; padding: 40px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); width: 100%; max-width: 400px; }
        .login-card h2 { text-align: center; font-weight: 700; margin-bottom: 10px; color: #1e3c72; }
        .login-card .subtitle { text-align: center; color: #7f8c8d; font-size: 0.85rem; margin-bottom: 30px; }
        .login-card label { font-size: 0.85rem; font-weight: 500; color: #34495e; }
        .login-card input { border-radius: 10px; padding: 12px; margin-bottom: 15px; border: 1px solid #e0e6ed; width: 100%; }
        .login-card input:focus { border-color: #1e3c72; box-shadow: 0 0 0 0.2rem rgba(30,60,114,0.15); }
        .login-card button { border-radius: 10px; padding: 12px; font-weight: 600; background: #1e3c72; border: none; color: white; width: 100%; transition: 0.3s; }
        .login-card button:hover { background: #2a5298; }
    </style>
</head>
<body>
    <div class="login-card">
        <h2>İSG TAKİP</h2>
        <p class="subtitle">Kerim Çağatay Toklu</p>
        <form method="post" action="/giris">
            <label>Kullanıcı Adı</label>
            <input type="text" name="kullanici_adi" placeholder="Kullanıcı adınızı girin" required autofocus>
            <label>Şifre</label>
            <input type="password" name="sifre" placeholder="Şifrenizi girin" required>
            {% if hata %}<div class="alert alert-danger py-2" style="font-size:0.85rem;">{{ hata }}</div>{% endif %}
            <button type="submit" class="mt-2">Giriş Yap</button>
        </form>
    </div>
</body>
</html>
"""

# ---------------- FİRMA SEÇİM SAYFASI ----------------
FIRMA_SECIM_HTML = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <title>Firma Seç - İSG TAKİP</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        body { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); min-height: 100vh; padding: 40px 15px; font-family: 'Poppins', sans-serif; }
        h2 { text-align: center; color: white; font-weight: 700; margin-bottom: 10px; }
        .subtitle { text-align: center; color: rgba(255,255,255,0.7); margin-bottom: 40px; font-size: 0.9rem; }
        .firma-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 20px; max-width: 900px; margin: 0 auto; }
        .firma-card { background: white; border-radius: 20px; padding: 30px 25px; text-align: center; cursor: pointer; transition: 0.3s; box-shadow: 0 10px 30px rgba(0,0,0,0.2); }
        .firma-card:hover { transform: translateY(-5px); box-shadow: 0 20px 40px rgba(0,0,0,0.3); }
        .firma-card .icon { font-size: 3rem; color: #2a5298; margin-bottom: 15px; }
        .firma-card .name { font-weight: 700; font-size: 1.1rem; color: #2c3e50; }
        .firma-card .desc { font-size: 0.8rem; color: #7f8c8d; margin-top: 8px; }
        .firma-card.super { background: linear-gradient(135deg, #f39c12 0%, #e67e22 100%); }
        .firma-card.super .icon, .firma-card.super .name, .firma-card.super .desc { color: white; }
        .firma-card.super .desc { color: rgba(255,255,255,0.9); }
        .logout-link { text-align: center; margin-top: 40px; }
        .logout-link a { color: rgba(255,255,255,0.8); text-decoration: none; font-size: 0.9rem; }
        .logout-link a:hover { color: white; }
    </style>
</head>
<body>
    <div class="container">
        <h2>İSG TAKİP</h2>
        <p class="subtitle">Yönetmek istediğiniz firmayı seçin</p>
        <div class="firma-grid">
            {% for f in firmalar %}
            <a href="/firma_sec/{{ f }}" class="text-decoration-none">
                <div class="firma-card">
                    <div class="icon"><i class="bi bi-building"></i></div>
                    <div class="name">{{ f }}</div>
                    <div class="desc">Verileri yönetmek için tıkla</div>
                </div>
            </a>
            {% endfor %}
            <a href="/firma_sec/TÜM FİRMALAR" class="text-decoration-none">
                <div class="firma-card super">
                    <div class="icon"><i class="bi bi-buildings"></i></div>
                    <div class="name">TÜM FİRMALAR</div>
                    <div class="desc">Bütün verileri bir arada görüntüle</div>
                </div>
            </a>
        </div>
        <div class="logout-link">
            <a href="/cikis"><i class="bi bi-box-arrow-right"></i> Çıkış Yap</a>
        </div>
    </div>
</body>
</html>
"""

# ---------------- ROUTE'LAR VE YARDIMCI FONKSİYONLAR ----------------
@app.route('/giris', methods=['GET', 'POST'])
def giris():
    if request.method == 'POST':
        kullanici = request.form.get('kullanıcı_adi', request.form.get('kullanici_adi', '')).strip()
        sifre = request.form.get('sifre', '')
        if kullanici == KULLANICI_ADI and sifre == SITE_SIFRESI:
            session['giris_yapildi'] = True
            session['kullanici_adi'] = kullanici
            return redirect(url_for('firma_secim_sayfasi'))
        else:
            return render_template_string(GIRIS_HTML, hata="Kullanıcı adı veya şifre hatalı!")
    return render_template_string(GIRIS_HTML, hata=None)

@app.route('/cikis')
def cikis():
    session.clear()
    return redirect(url_for('giris'))

@app.route('/firma_secim')
def firma_secim_sayfasi():
    if not session.get('giris_yapildi'):
        return redirect(url_for('giris'))
    return render_template_string(FIRMA_SECIM_HTML, firmalar=get_firmalar_listesi())

@app.route('/firma_sec/<path:firma>')
def firma_sec(firma):
    if not session.get('giris_yapildi'):
        return redirect(url_for('giris'))
    if firma != 'TÜM FİRMALAR' and firma not in get_firmalar_listesi():
        return redirect(url_for('firma_secim_sayfasi'))
    session['firma_adi'] = firma
    return redirect(url_for('dashboard'))

@app.route('/firma_degistir')
def firma_degistir():
    return redirect(url_for('firma_secim_sayfasi'))

@app.before_request
def giris_kontrol():
    izinli = ['giris', 'static', 'firma_secim_sayfasi', 'firma_sec', 'firma_degistir', 'sertifika_dogrula']
    if request.endpoint in izinli:
        return None
    if not session.get('giris_yapildi'):
        return redirect(url_for('giris'))
    if not session.get('firma_adi') and request.endpoint != 'firma_secim_sayfasi':
        return redirect(url_for('firma_secim_sayfasi'))

def aktif_firma():
    return session.get('firma_adi', '')

def firma_filtre(query):
    firma = aktif_firma()
    if firma != 'TÜM FİRMALAR':
        return query.filter_by(firma_adi=firma)
    return query

def get_personeller(firma=None):
    if not firma: firma = aktif_firma()
    q1 = db.session.query(Egitim.personel_adi)
    q2 = db.session.query(IsKazasi.personel_adi)
    if firma != 'TÜM FİRMALAR':
        q1 = q1.filter(Egitim.firma_adi == firma)
        q2 = q2.filter(IsKazasi.firma_adi == firma)
    personeller = q1.distinct().all()
    kaza_personeller = q2.distinct().all()
    tum = set([p[0] for p in personeller if p[0]] + [p[0] for p in kaza_personeller if p[0]])
    return sorted(list(tum))

def hesapla_gecerlilik(tarih, tehlike_sinifi):
    if not tarih or not tehlike_sinifi: return None
    if tehlike_sinifi == "Az Tehlikeli": return tarih + relativedelta(years=6)
    elif tehlike_sinifi == "Tehlikeli": return tarih + relativedelta(years=4)
    elif tehlike_sinifi == "Çok Tehlikeli": return tarih + relativedelta(years=2)
    return None

def resim_kaydet(resim_dosyasi):
    if resim_dosyasi and resim_dosyasi.filename != '':
        # Dosya uzantısını al (örn: .jpg, .png)
        uzanti = os.path.splitext(resim_dosyasi.filename)[1].lower()
        # Benzersiz bir isim oluştur
        benzersiz_isim = f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}{uzanti}"
        
        # Klasörün var olduğundan emin ol (yoksa oluştur)
        upload_dir = os.path.join(app.root_path, 'static', 'uploads')
        os.makedirs(upload_dir, exist_ok=True)
        
        # Tam kayıt yolu
        kayit_yolu = os.path.join(upload_dir, benzersiz_isim)
        
        # Dosyayı kaydet
        resim_dosyasi.save(kayit_yolu)
        return benzersiz_isim
    return None

def resim_sil(resim_adi):
    if resim_adi:
        dosya_yolu = os.path.join(app.config['UPLOAD_FOLDER'], resim_adi)
        if os.path.exists(dosya_yolu):
            os.remove(dosya_yolu)
# ---------------- SAYFA İÇERİKLERİ ----------------
DASHBOARD_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    {% if uyarilar %}
    <div class="alert-panel">
        <div class="alert-panel-title"><i class="bi bi-exclamation-triangle-fill"></i> Güncelleme / Yenileme Uyarıları</div>
        {% for uyari in uyarilar %}
        <div class="alert-item">
            <span><strong>{{ uyari.firma }}</strong> - {{ uyari.mesaj }}</span>
            <span class="badge bg-danger">{{ uyari.durum }}</span>
        </div>
        {% endfor %}
    </div>
    {% endif %}

    <div class="row g-3 mb-4">
        <div class="col-md-6">
            <div class="widget-card widget-rd">
                <div class="widget-header">
                    <div class="widget-title"><i class="bi bi-shield-check text-primary widget-icon"></i> Risk Değerlendirmesi</div>
                    <a href="/planlar" class="btn-glass text-decoration-none text-dark">Detaylar</a>
                </div>
                <div class="table-responsive">
                    <table class="table table-custom align-middle">
                        <tbody>
                            {% for plan in risk_planlari if plan.rd_tarihi %}
                            <tr>
                                <td><strong>{{ plan.firma_adi }}</strong></td>
                                <td class="text-end text-muted">{{ plan.rd_tehlike_sinifi or '-' }}</td>
                                <td class="text-end">
                                    {% if plan.rd_gecerlilik < bugun %}
                                        <span class="badge bg-danger badge-glass">{{ plan.rd_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                    {% else %}
                                        <span class="badge bg-success badge-glass">{{ plan.rd_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                    {% endif %}
                                </td>
                            </tr>
                            {% else %}
                            <tr><td colspan="3" class="text-center text-muted py-3" style="font-size:0.75rem;">Kayıt yok.</td></tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="widget-card widget-adp">
                <div class="widget-header">
                    <div class="widget-title"><i class="bi bi-fire text-danger widget-icon"></i> Acil Durum Planı</div>
                    <a href="/planlar" class="btn-glass text-decoration-none text-dark">Detaylar</a>
                </div>
                <div class="table-responsive">
                    <table class="table table-custom align-middle">
                        <tbody>
                            {% for plan in risk_planlari if plan.adp_tarihi %}
                            <tr>
                                <td><strong>{{ plan.firma_adi }}</strong></td>
                                <td class="text-end text-muted">{{ plan.adp_tehlike_sinifi or '-' }}</td>
                                <td class="text-end">
                                    {% if plan.adp_gecerlilik < bugun %}
                                        <span class="badge bg-danger badge-glass">{{ plan.adp_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                    {% else %}
                                        <span class="badge bg-success badge-glass">{{ plan.adp_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                    {% endif %}
                                </td>
                            </tr>
                            {% else %}
                            <tr><td colspan="3" class="text-center text-muted py-3" style="font-size:0.75rem;">Kayıt yok.</td></tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="widget-card widget-egitim-plan">
                <div class="widget-header">
                    <div class="widget-title"><i class="bi bi-journal-bookmark-fill text-success widget-icon"></i> Eğitim Planı</div>
                    <a href="/planlar" class="btn-glass text-decoration-none text-dark">Detaylar</a>
                </div>
                <div class="table-responsive">
                    <table class="table table-custom align-middle">
                        <tbody>
                            {% for plan in risk_planlari if plan.egitim_plani_tarihi %}
                            <tr>
                                <td><strong>{{ plan.firma_adi }}</strong></td>
                                <td class="text-end text-muted">{{ plan.egitim_tehlike_sinifi or '-' }}</td>
                                <td class="text-end">
                                    {% if plan.egitim_plani_gecerlilik < bugun %}
                                        <span class="badge bg-danger badge-glass">{{ plan.egitim_plani_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                    {% else %}
                                        <span class="badge bg-success badge-glass">{{ plan.egitim_plani_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                    {% endif %}
                                </td>
                            </tr>
                            {% else %}
                            <tr><td colspan="3" class="text-center text-muted py-3" style="font-size:0.75rem;">Kayıt yok.</td></tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="widget-card widget-calisma-plan">
                <div class="widget-header">
                    <div class="widget-title"><i class="bi bi-calendar-check-fill text-dark widget-icon"></i> Çalışma Planı</div>
                    <a href="/planlar" class="btn-glass text-decoration-none text-dark">Detaylar</a>
                </div>
                <div class="table-responsive">
                    <table class="table table-custom align-middle">
                        <tbody>
                            {% for plan in risk_planlari if plan.calisma_plani_tarihi %}
                            <tr>
                                <td><strong>{{ plan.firma_adi }}</strong></td>
                                <td class="text-end text-muted">{{ plan.calisma_tehlike_sinifi or '-' }}</td>
                                <td class="text-end">
                                    {% if plan.calisma_plani_gecerlilik < bugun %}
                                        <span class="badge bg-danger badge-glass">{{ plan.calisma_plani_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                    {% else %}
                                        <span class="badge bg-success badge-glass">{{ plan.calisma_plani_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                    {% endif %}
                                </td>
                            </tr>
                            {% else %}
                            <tr><td colspan="3" class="text-center text-muted py-3" style="font-size:0.75rem;">Kayıt yok.</td></tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <h5 class="mb-3 fw-bold" style="font-size:1rem;">Eğitim Takip Özeti</h5>
    <div class="row g-3">
        <div class="col-md-6">
            <div class="widget-card" style="border-left-color: #e74c3c; padding: 15px;">
                <div class="widget-header mb-2">
                    <div class="widget-title" style="font-size:0.95rem;"><i class="bi bi-exclamation-octagon text-danger"></i> Süresi Geçenler</div>
                    <span class="badge bg-danger" style="font-size:0.7rem;">{{ suresi_gecenler|length }}</span>
                </div>
                {% for e in suresi_gecenler %}
                <div class="mini-card danger">
                    <div><div class="mini-card-title">{{ e.personel_adi }}</div><div class="mini-card-sub">{{ e.egitim_adi }}</div></div>
                    <div class="text-danger mini-card-date">{{ e.son_gecerlilik.strftime('%d.%m.%Y') }}</div>
                </div>
                {% else %}
                <p class="text-muted text-center py-2" style="font-size:0.8rem;">Kayıt bulunmuyor.</p>
                {% endfor %}
            </div>
        </div>
        <div class="col-md-6">
            <div class="widget-card" style="border-left-color: #f39c12; padding: 15px;">
                <div class="widget-header mb-2">
                    <div class="widget-title" style="font-size:0.95rem;"><i class="bi bi-hourglass-split text-warning"></i> Yaklaşanlar (30 Gün)</div>
                    <span class="badge bg-warning text-dark" style="font-size:0.7rem;">{{ yaklasan_egitimler|length }}</span>
                </div>
                {% for e in yaklasan_egitimler %}
                <div class="mini-card warning">
                    <div><div class="mini-card-title">{{ e.personel_adi }}</div><div class="mini-card-sub">{{ e.egitim_adi }}</div></div>
                    <div class="text-warning mini-card-date">{{ e.son_gecerlilik.strftime('%d.%m.%Y') }}</div>
                </div>
                {% else %}
                <p class="text-muted text-center py-2" style="font-size:0.8rem;">Kayıt bulunmuyor.</p>
                {% endfor %}
            </div>
        </div>
    </div>
""")

PLANLAR_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <h3 class="mb-0 fw-bold" style="font-size: 1.1rem;">Tüm Planlar</h3>
        <a href="/plan/ekle" class="btn btn-primary rounded-pill px-3 py-1" style="font-size:0.8rem;">+ Yeni Plan Ekle</a>
    </div>
    <div class="widget-card p-0 overflow-hidden" style="border-left:none;">
        <div class="table-responsive">
            <table class="table table-custom align-middle mb-0">
                <thead style="background:#f8f9fa;">
                    <tr>
                        <th rowspan="2" class="align-middle border-end ps-3" style="font-size:0.75rem;">Firma Adı</th>
                        <th colspan="3" class="text-center border-end" style="font-size:0.75rem;">Risk Değerlendirmesi</th>
                        <th colspan="3" class="text-center border-end" style="font-size:0.75rem;">Acil Durum Planı</th>
                        <th colspan="3" class="text-center border-end" style="font-size:0.75rem;">Eğitim Planı</th>
                        <th colspan="3" class="text-center border-end" style="font-size:0.75rem;">Çalışma Planı</th>
                        <th rowspan="2" class="align-middle text-center pe-3" style="font-size:0.75rem;">İşlemler</th>
                    </tr>
                    <tr>
                        <th class="text-center" style="font-size:0.7rem;">Tarih</th><th class="text-center" style="font-size:0.7rem;">Sınıf</th><th class="text-center border-end" style="font-size:0.7rem;">Geçerlilik</th>
                        <th class="text-center" style="font-size:0.7rem;">Tarih</th><th class="text-center" style="font-size:0.7rem;">Sınıf</th><th class="text-center border-end" style="font-size:0.7rem;">Geçerlilik</th>
                        <th class="text-center" style="font-size:0.7rem;">Tarih</th><th class="text-center" style="font-size:0.7rem;">Sınıf</th><th class="text-center border-end" style="font-size:0.7rem;">Geçerlilik</th>
                        <th class="text-center" style="font-size:0.7rem;">Tarih</th><th class="text-center" style="font-size:0.7rem;">Sınıf</th><th class="text-center border-end" style="font-size:0.7rem;">Geçerlilik</th>
                    </tr>
                </thead>
                <tbody>
                    {% for plan in planlar %}
                    <tr>
                        <td class="border-end ps-3"><strong>{{ plan.firma_adi }}</strong></td>
                        <td class="text-center">{{ plan.rd_tarihi.strftime('%d.%m.%Y') if plan.rd_tarihi else '-' }}</td>
                        <td class="text-center">{{ plan.rd_tehlike_sinifi or '-' }}</td>
                        <td class="text-center border-end">
                            {% if plan.rd_gecerlilik %}
                                {% if plan.rd_gecerlilik < bugun %}<span class="badge bg-danger badge-glass">{{ plan.rd_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                {% else %}<span class="badge bg-success badge-glass">{{ plan.rd_gecerlilik.strftime('%d.%m.%Y') }}</span>{% endif %}
                            {% else %}-{% endif %}
                        </td>
                        <td class="text-center">{{ plan.adp_tarihi.strftime('%d.%m.%Y') if plan.adp_tarihi else '-' }}</td>
                        <td class="text-center">{{ plan.adp_tehlike_sinifi or '-' }}</td>
                        <td class="text-center border-end">
                            {% if plan.adp_gecerlilik %}
                                {% if plan.adp_gecerlilik < bugun %}<span class="badge bg-danger badge-glass">{{ plan.adp_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                {% else %}<span class="badge bg-success badge-glass">{{ plan.adp_gecerlilik.strftime('%d.%m.%Y') }}</span>{% endif %}
                            {% else %}-{% endif %}
                        </td>
                        <td class="text-center">{{ plan.egitim_plani_tarihi.strftime('%d.%m.%Y') if plan.egitim_plani_tarihi else '-' }}</td>
                        <td class="text-center">{{ plan.egitim_tehlike_sinifi or '-' }}</td>
                        <td class="text-center border-end">
                            {% if plan.egitim_plani_gecerlilik %}
                                {% if plan.egitim_plani_gecerlilik < bugun %}<span class="badge bg-danger badge-glass">{{ plan.egitim_plani_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                {% else %}<span class="badge bg-success badge-glass">{{ plan.egitim_plani_gecerlilik.strftime('%d.%m.%Y') }}</span>{% endif %}
                            {% else %}-{% endif %}
                        </td>
                        <td class="text-center">{{ plan.calisma_plani_tarihi.strftime('%d.%m.%Y') if plan.calisma_plani_tarihi else '-' }}</td>
                        <td class="text-center">{{ plan.calisma_tehlike_sinifi or '-' }}</td>
                        <td class="text-center border-end">
                            {% if plan.calisma_plani_gecerlilik %}
                                {% if plan.calisma_plani_gecerlilik < bugun %}<span class="badge bg-danger badge-glass">{{ plan.calisma_plani_gecerlilik.strftime('%d.%m.%Y') }}</span>
                                {% else %}<span class="badge bg-success badge-glass">{{ plan.calisma_plani_gecerlilik.strftime('%d.%m.%Y') }}</span>{% endif %}
                            {% else %}-{% endif %}
                        </td>
                        <td class="text-center pe-3">
                            <a href="/plan/duzenle/{{ plan.id }}" class="btn-glass text-decoration-none text-primary"><i class="bi bi-pencil-square"></i></a>
                            <a href="/plan/sil/{{ plan.id }}" class="btn-glass text-decoration-none text-danger" onclick="return confirm('Silmek istediğinize emin misiniz?')"><i class="bi bi-trash"></i></a>
                        </td>
                    </tr>
                    {% else %}
                    <tr><td colspan="14" class="text-center text-muted py-3">Kayıtlı plan bulunmuyor.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
""")

EGITIMLER_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <h3 class="mb-0 fw-bold" style="font-size: 1.1rem;">Eğitim Listesi</h3>
        <div class="d-flex gap-2 align-items-center flex-nowrap">
            <form method="get" class="d-flex gap-2 align-items-center m-0">
                <input type="text" name="arama" class="form-control border-0 shadow-sm form-control-sm" placeholder="Personel veya Eğitim Ara..." value="{{ arama or '' }}" style="width: 250px; border-radius: 20px;">
                <button type="submit" class="btn btn-primary rounded-pill px-3" style="font-size:0.8rem;"><i class="bi bi-search"></i></button>
            </form>
            <form method="get" class="d-flex gap-2 align-items-center m-0">
                <select name="siralama" class="form-select border-0 shadow-sm" style="width: auto; font-size:0.8rem; padding: 6px 12px; border-radius: 8px;" onchange="this.form.submit()">
                    <option value="yeni" {% if siralama == 'yeni' %}selected{% endif %}>En Yeni</option>
                    <option value="eski" {% if siralama == 'eski' %}selected{% endif %}>En Eski</option>
                </select>
            </form>
            <a href="/egitimler/excel?siralama={{ siralama }}&arama={{ arama or '' }}" class="btn btn-success text-nowrap rounded-pill px-3" style="font-size:0.8rem;"><i class="bi bi-file-earmark-excel"></i> Excel</a>
            <a href="/egitim/ekle" class="btn btn-primary text-nowrap rounded-pill px-3" style="font-size:0.8rem;">+ Eğitim Ekle</a>
        </div>
    </div>
    <div class="widget-card p-0 overflow-hidden" style="border-left:none;">
        <div class="table-responsive">
            <table class="table table-custom align-middle table-hover mb-0">
                <thead style="background:#f8f9fa;">
                    <tr>
                        <th class="ps-3" style="font-size:0.75rem;">Sıra</th>
                        <th style="font-size:0.75rem;">Resim</th>
                        <th style="font-size:0.75rem;">Firma</th>
                        <th style="font-size:0.75rem;">Personel</th>
                        <th style="font-size:0.75rem;">Görevi</th>
                        <th style="font-size:0.75rem;">İşe Giriş</th>
                        <th style="font-size:0.75rem;">Eğitim Adı</th>
                        <th style="font-size:0.75rem;">Eğitim Tarihi</th>
                        <th style="font-size:0.75rem;">Son Geçerlilik</th>
                        <th style="font-size:0.75rem;">Durum</th>
                        <th style="font-size:0.75rem;">Kaza</th>
                        <th class="pe-3 text-center" style="font-size:0.75rem;">İşlemler</th>
                    </tr>
                </thead>
                <tbody>
                    {% for e in egitimler %}
                    <tr class="clickable-row" onclick="window.location='/personel_detay/{{ e.personel_adi }}'">
                        <td class="ps-3">{{ loop.index }}</td>
                        <td>
                            {% if e.resim %}<img src="{{ url_for('static', filename='uploads/' + e.resim) }}" class="rounded-circle" style="width:30px; height:30px; object-fit:cover;">
                            {% else %}<div class="bg-secondary text-white rounded-circle d-flex align-items-center justify-content-center" style="width:30px; height:30px; font-size:10px;">?</div>{% endif %}
                        </td>
                        <td>{{ e.firma_adi }}</td>
                        <td class="fw-bold">{{ e.personel_adi }}</td>
                        <td>{{ e.gorevi or '-' }}</td>
                        <td>{{ e.ise_giris.strftime('%d.%m.%Y') if e.ise_giris else '-' }}</td>
                        <td>{{ e.egitim_adi }}</td>
                        <td>{{ e.egitim_tarihi.strftime('%d.%m.%Y') }}</td>
                        <td>{{ e.son_gecerlilik.strftime('%d.%m.%Y') }}</td>
                        <td>
                            {% if e.durum == 'SÜRESİ GEÇMİŞ' %}<span class="badge bg-danger badge-glass">Geçmiş</span>
                            {% elif e.durum == 'YAKLAŞIYOR' %}<span class="badge bg-warning text-dark badge-glass">Yaklaşıyor</span>
                            {% else %}<span class="badge bg-success badge-glass">Geçerli</span>{% endif %}
                        </td>
                        <td>
                            {% set kaza_var = kazalar_dict.get(e.personel_adi.lower().strip()) %}
                            {% if kaza_var %}<span class="badge bg-danger badge-glass">⚠️ ({{ kaza_var }})</span>
                            {% else %}<span class="badge bg-light text-dark border badge-glass">Yok</span>{% endif %}
                        </td>
                        <td class="pe-3 text-center" onclick="event.stopPropagation();">
                            <a href="/sertifika/{{ e.id }}" class="btn-glass text-decoration-none text-success" title="Sertifikayı Görüntüle"><i class="bi bi-award"></i></a>
                            <a href="/egitim/duzenle/{{ e.id }}" class="btn-glass text-decoration-none text-primary" title="Düzenle"><i class="bi bi-pencil-square"></i></a>
                            <a href="/egitim/sil/{{ e.id }}" class="btn-glass text-decoration-none text-danger" title="Sil" onclick="return confirm('Silmek istediğinize emin misiniz?')"><i class="bi bi-trash"></i></a>
                        </td>
                    </tr>
                    {% else %}
                    <tr><td colspan="12" class="text-center text-muted py-3">Aramanıza uygun veya kayıtlı eğitim bulunmuyor.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
""")

KAZALAR_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <h3 class="mb-0 fw-bold" style="font-size: 1.1rem;">İş Kazaları Listesi</h3>
        <div class="d-flex gap-2 align-items-center flex-nowrap">
            <form method="get" class="d-flex gap-2 align-items-center m-0">
                <select name="siralama" class="form-select border-0 shadow-sm" style="width: auto; font-size:0.8rem; padding: 6px 12px; border-radius: 8px;" onchange="this.form.submit()">
                    <option value="yeni" {% if siralama == 'yeni' %}selected{% endif %}>En Yeni</option>
                    <option value="eski" {% if siralama == 'eski' %}selected{% endif %}>En Eski</option>
                </select>
            </form>
            <a href="/kazalar/excel?siralama={{ siralama }}" class="btn btn-success text-nowrap rounded-pill px-3" style="font-size:0.8rem;"><i class="bi bi-file-earmark-excel"></i> Excel</a>
            <a href="/kaza/ekle" class="btn btn-danger text-nowrap rounded-pill px-3" style="font-size:0.8rem;">+ Kaza</a>
        </div>
    </div>
    <div class="widget-card p-0 overflow-hidden" style="border-left:none;">
        <div class="table-responsive">
            <table class="table table-custom align-middle table-hover mb-0">
                <thead style="background:#f8f9fa;">
                    <tr>
                        <th class="ps-3" style="font-size:0.75rem;">Sıra</th><th style="font-size:0.75rem;">Firma</th><th style="font-size:0.75rem;">Personel</th><th style="font-size:0.75rem;">Kaza Tarihi</th><th style="font-size:0.75rem;">Yaralanma</th>
                        <th style="font-size:0.75rem;">SGK Bildirim</th><th style="font-size:0.75rem;">Rapor Gün</th><th style="font-size:0.75rem;">Kaza Sonrası Eğitim</th><th class="pe-3" style="font-size:0.75rem;">İşlem</th>
                    </tr>
                </thead>
                <tbody>
                    {% for k in kazalar %}
                    <tr>
                        <td class="ps-3">{{ loop.index }}</td>
                        <td>{{ k.firma_adi }}</td>
                        <td class="fw-bold">{{ k.personel_adi }}</td>
                        <td>{{ k.tarih.strftime('%d.%m.%Y') }}</td>
                        <td><span class="badge bg-danger badge-glass">{{ k.yaralanma_turu }}</span></td>
                        <td>{{ k.sgk_bildirim_tarihi.strftime('%d.%m.%Y') if k.sgk_bildirim_tarihi else '-' }}</td>
                        <td>{{ k.rapor_gun_sayisi or '-' }}</td>
                        <td>{{ k.kaza_sonrasi_egitim.strftime('%d.%m.%Y') if k.kaza_sonrasi_egitim else '-' }}</td>
                        <td class="pe-3">
                            <a href="/kaza/duzenle/{{ k.id }}" class="btn-glass text-decoration-none text-primary"><i class="bi bi-pencil-square"></i></a>
                            <a href="/kaza/sil/{{ k.id }}" class="btn-glass text-decoration-none text-danger" onclick="return confirm('Silmek istediğinize emin misiniz?')"><i class="bi bi-trash"></i></a>
                        </td>
                    </tr>
                    {% else %}
                    <tr><td colspan="9" class="text-center text-muted py-3">Kayıtlı kaza bulunmuyor.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
""")

PERSONEL_DETAY_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <a href="/egitimler" class="btn btn-light mb-3 rounded-pill px-3" style="font-size:0.8rem;"><i class="bi bi-arrow-left"></i> Listeye Dön</a>
    <div class="widget-card mb-3" style="border-left: 5px solid #3498db;">
        <div class="d-flex align-items-center">
            {% set kisi = personel_bilgi %}
            {% if kisi and kisi.resim %}<img src="{{ url_for('static', filename='uploads/' + kisi.resim) }}" class="rounded-circle me-3" style="width:70px; height:70px; object-fit:cover; border:3px solid #f0f2f5;">
            {% else %}<div class="bg-secondary text-white rounded-circle d-flex align-items-center justify-content-center me-3" style="width:70px; height:70px; font-size:24px;">?</div>{% endif %}
            <div>
                <h4 class="mb-1 fw-bold">{{ personel_adi }}</h4>
                <h6 class="text-primary mb-1" style="font-size:0.85rem;">{{ personel_bilgi.firma_adi if personel_bilgi else '-' }} - {{ personel_bilgi.gorevi if personel_bilgi and personel_bilgi.gorevi else 'Görev Belirtilmemiş' }}</h6>
                <p class="mb-0 text-muted" style="font-size:0.8rem;"><i class="bi bi-calendar-event"></i> İşe Giriş: {{ personel_bilgi.ise_giris.strftime('%d.%m.%Y') if personel_bilgi and personel_bilgi.ise_giris else '-' }}</p>
            </div>
        </div>
    </div>
    <div class="row">
        <div class="col-md-12">
            <h6 class="text-info border-bottom pb-2 mb-2 fw-bold" style="font-size:0.9rem;"><i class="bi bi-mortarboard-fill"></i> Eğitim Geçmişi</h6>
            {% if egitimler %}
                <div class="widget-card p-0 overflow-hidden mb-3" style="border-left:none;">
                    <div class="table-responsive">
                        <table class="table table-custom align-middle mb-0">
                            <thead style="background:#f8f9fa;"><tr>
                                <th class="ps-3" style="font-size:0.75rem;">Eğitim Adı</th><th style="font-size:0.75rem;">Eğitim Tarihi</th><th style="font-size:0.75rem;">Geçerlilik (Ay)</th><th style="font-size:0.75rem;">Son Geçerlilik</th><th style="font-size:0.75rem;">Sertifika No</th><th class="pe-3" style="font-size:0.75rem;">Durum</th>
                            </tr></thead>
                            <tbody>
                                {% for e in egitimler %}
                                <tr>
                                    <td class="ps-3">{{ e.egitim_adi }}</td><td>{{ e.egitim_tarihi.strftime('%d.%m.%Y') }}</td><td>{{ e.gecerlilik_ay }}</td><td>{{ e.son_gecerlilik.strftime('%d.%m.%Y') }}</td><td>{{ e.sertifika_no or '-' }}</td>
                                    <td class="pe-3">
                                        {% if e.durum == 'SÜRESİ GEÇMİŞ' %}<span class="badge bg-danger badge-glass">Süresi Geçmiş</span>
                                        {% elif e.durum == 'YAKLAŞIYOR' %}<span class="badge bg-warning text-dark badge-glass">Yaklaşıyor</span>
                                        {% else %}<span class="badge bg-success badge-glass">Geçerli</span>{% endif %}
                                    </td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            {% else %}<div class="alert alert-info py-2" style="font-size:0.8rem;">Bu personelin kayıtlı eğitimi bulunmuyor.</div>{% endif %}
        </div>
    </div>
    <div class="row mt-3">
        <div class="col-md-12">
            <h6 class="text-danger border-bottom pb-2 mb-2 fw-bold" style="font-size:0.9rem;"><i class="bi bi-exclamation-triangle-fill"></i> İş Kazası Geçmişi</h6>
            {% if kazalar %}
                <div class="widget-card p-0 overflow-hidden" style="border-left:none;">
                    <div class="table-responsive">
                        <table class="table table-custom align-middle mb-0">
                            <thead style="background:#f8f9fa;"><tr>
                                <th class="ps-3" style="font-size:0.75rem;">Tarih</th><th style="font-size:0.75rem;">Kaza Yeri</th><th style="font-size:0.75rem;">Yaralanma Türü</th><th style="font-size:0.75rem;">Açıklama</th><th style="font-size:0.75rem;">Kök Neden</th><th style="font-size:0.75rem;">Düzeltici Faaliyet</th>
                                <th style="font-size:0.75rem;">SGK Bildirim</th><th style="font-size:0.75rem;">Rapor Gün</th><th class="pe-3" style="font-size:0.75rem;">Kaza Sonrası Eğitim</th>
                            </tr></thead>
                            <tbody>
                                {% for k in kazalar %}
                                <tr>
                                    <td class="ps-3">{{ k.tarih.strftime('%d.%m.%Y') }}</td><td>{{ k.yer }}</td><td><span class="badge bg-danger badge-glass">{{ k.yaralanma_turu }}</span></td><td>{{ k.aciklama or '-' }}</td><td>{{ k.kok_neden or '-' }}</td><td>{{ k.duzeltici_faaliyet or '-' }}</td>
                                    <td>{{ k.sgk_bildirim_tarihi.strftime('%d.%m.%Y') if k.sgk_bildirim_tarihi else '-' }}</td>
                                    <td>{{ k.rapor_gun_sayisi or '-' }}</td>
                                    <td class="pe-3">{{ k.kaza_sonrasi_egitim.strftime('%d.%m.%Y') if k.kaza_sonrasi_egitim else '-' }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            {% else %}<div class="alert alert-success py-2" style="font-size:0.8rem;"><i class="bi bi-check-circle-fill"></i> Bu personelin kayıtlı bir iş kazası bulunmuyor.</div>{% endif %}
        </div>
    </div>
""")

EGITIM_EKLE_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <h4 class="mb-3 fw-bold" style="font-size: 1.1rem;">Yeni Eğitim ve Personel Kaydı</h4>
    <div class="widget-card" style="border-left:none;">
        <form method="post" enctype="multipart/form-data">
            <div class="row">
                <div class="col-md-6 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Personel ve Firma</h6>
                    <label class="form-label" style="font-size:0.8rem;">Firma</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" value="{{ session.get('firma_adi') }}" disabled style="background:#e8f4ff; font-weight:600;">

                    <label class="form-label" style="font-size:0.8rem;">Personel Adı Soyadı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="personel_adi" list="personel_listesi" placeholder="Personel seçin veya yazın" required>
                    <datalist id="personel_listesi">{% for p in personeller %}<option value="{{ p }}">{% endfor %}</datalist>

                    <label class="form-label" style="font-size:0.8rem;">Görevi</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="gorevi" placeholder="Örn: Operatör">
                    
                    <label class="form-label" style="font-size:0.8rem;">İşe Giriş Tarihi</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="ise_giris">
                    
                    <label class="form-label" style="font-size:0.8rem;">Fotoğraf</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" type="file" name="resim" accept="image/*">
                </div>
                <div class="col-md-6 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Eğitim Bilgileri</h6>
                    <label class="form-label" style="font-size:0.8rem;">Eğitim Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="egitim_adi" placeholder="Örn: Temel İSG" required>
                    
                    <label class="form-label" style="font-size:0.8rem;">Eğitim Tarihi</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="egitim_tarihi" required>
                    
                    <label class="form-label" style="font-size:0.8rem;">Geçerlilik Süresi (Ay)</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="number" name="gecerlilik_ay" value="12" required>
                    
                    <label class="form-label" style="font-size:0.8rem;">Sertifika No</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" name="sertifika_no" placeholder="Varsa sertifika no">
                </div>
            </div>
            <button class="btn btn-primary rounded-pill px-4 py-2 mt-2" style="font-size:0.85rem;">Kaydet</button>
        </form>
    </div>
""")

EGITIM_DUZENLE_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <h4 class="mb-3 fw-bold" style="font-size: 1.1rem;">Eğitim ve Personel Düzenle</h4>
    <div class="widget-card" style="border-left:none;">
        <form method="post" enctype="multipart/form-data">
            <div class="row">
                <div class="col-md-6 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Personel ve Firma</h6>
                    {% if e.resim %}
                        <div class="mb-2 text-center">
                            <img src="{{ url_for('static', filename='uploads/' + e.resim) }}" class="rounded-circle mb-2" style="width:70px; height:70px; object-fit:cover;">
                            <div class="form-check d-flex justify-content-center">
                                <input class="form-check-input me-2" type="checkbox" name="resim_sil" id="resimSil">
                                <label class="form-check-label text-danger" for="resimSil" style="font-size:0.75rem;">Mevcut fotoğrafı sil</label>
                            </div>
                        </div>
                    {% endif %}
                    <label class="form-label" style="font-size:0.8rem;">Yeni Fotoğraf</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="file" name="resim" accept="image/*">
                    <label class="form-label" style="font-size:0.8rem;">Firma</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" value="{{ e.firma_adi }}" disabled style="background:#e8f4ff; font-weight:600;">
                    <label class="form-label" style="font-size:0.8rem;">Personel Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="personel_adi" list="personel_listesi" value="{{ e.personel_adi }}" required>
                    <datalist id="personel_listesi">{% for p in personeller %}<option value="{{ p }}">{% endfor %}</datalist>
                    <label class="form-label" style="font-size:0.8rem;">Görevi</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="gorevi" value="{{ e.gorevi or '' }}">
                    <label class="form-label" style="font-size:0.8rem;">İşe Giriş</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" type="date" name="ise_giris" value="{{ e.ise_giris }}">
                </div>
                <div class="col-md-6 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Eğitim Bilgileri</h6>
                    <label class="form-label" style="font-size:0.8rem;">Eğitim Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="egitim_adi" value="{{ e.egitim_adi }}" required>
                    <label class="form-label" style="font-size:0.8rem;">Eğitim Tarihi</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="egitim_tarihi" value="{{ e.egitim_tarihi }}" required>
                    <label class="form-label" style="font-size:0.8rem;">Geçerlilik (Ay)</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="number" name="gecerlilik_ay" value="{{ e.gecerlilik_ay }}" required>
                    <label class="form-label" style="font-size:0.8rem;">Sertifika No</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" name="sertifika_no" value="{{ e.sertifika_no or '' }}">
                </div>
            </div>
            <button class="btn btn-primary rounded-pill px-4 py-2 mt-2" style="font-size:0.85rem;">Güncelle</button>
            <a href="/egitimler" class="btn btn-light rounded-pill px-4 py-2 mt-2 ms-2" style="font-size:0.85rem;">İptal</a>
        </form>
    </div>
""")

KAZA_EKLE_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <h4 class="mb-3 fw-bold" style="font-size: 1.1rem;">İş Kazası Bildirimi</h4>
    <div class="widget-card" style="border-left:none;">
        <form method="post">
            <div class="row">
                <div class="col-md-6 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Kaza Bilgileri</h6>
                    <label class="form-label" style="font-size:0.8rem;">Firma</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" value="{{ session.get('firma_adi') }}" disabled style="background:#e8f4ff; font-weight:600;">
                    <label class="form-label" style="font-size:0.8rem;">Personel Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="personel_adi" list="personel_listesi" placeholder="Personel seçin veya yazın" required>
                    <datalist id="personel_listesi">{% for p in personeller %}<option value="{{ p }}">{% endfor %}</datalist>
                    <label class="form-label" style="font-size:0.8rem;">Tarih</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="tarih" required>
                    <label class="form-label" style="font-size:0.8rem;">Kaza Yeri</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="yer" placeholder="Kaza yeri">
                    <label class="form-label" style="font-size:0.8rem;">Yaralanma Türü</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" name="yaralanma_turu" placeholder="Örn: El kesiği">
                </div>
                <div class="col-md-6 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Yasal ve Sonrası Süreçler</h6>
                    <label class="form-label" style="font-size:0.8rem;">SGK Bildirim Tarihi</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="sgk_bildirim_tarihi">
                    <label class="form-label" style="font-size:0.8rem;">Rapor Gün Sayısı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="number" name="rapor_gun_sayisi" placeholder="Örn: 5">
                    <label class="form-label" style="font-size:0.8rem;">Kaza Sonrası Eğitim</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" type="date" name="kaza_sonrasi_egitim">
                </div>
                <div class="col-md-12 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Açıklama ve Analiz</h6>
                    <label class="form-label" style="font-size:0.8rem;">Açıklama</label>
                    <textarea class="form-control mb-2 border-0 shadow-sm form-control-sm" name="aciklama" rows="2"></textarea>
                    <label class="form-label" style="font-size:0.8rem;">Kök Neden</label>
                    <textarea class="form-control mb-2 border-0 shadow-sm form-control-sm" name="kok_neden" rows="2"></textarea>
                    <label class="form-label" style="font-size:0.8rem;">Düzeltici Faaliyet</label>
                    <textarea class="form-control border-0 shadow-sm form-control-sm" name="duzeltici_faaliyet" rows="2"></textarea>
                </div>
            </div>
            <button class="btn btn-danger rounded-pill px-4 py-2 mt-2" style="font-size:0.85rem;">Bildir</button>
        </form>
    </div>
""")

KAZA_DUZENLE_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <h4 class="mb-3 fw-bold" style="font-size: 1.1rem;">İş Kazası Düzenle</h4>
    <div class="widget-card" style="border-left:none;">
        <form method="post">
            <div class="row">
                <div class="col-md-6 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Kaza Bilgileri</h6>
                    <label class="form-label" style="font-size:0.8rem;">Firma</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" value="{{ k.firma_adi }}" disabled style="background:#e8f4ff; font-weight:600;">
                    <label class="form-label" style="font-size:0.8rem;">Personel Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="personel_adi" list="personel_listesi" value="{{ k.personel_adi }}" required>
                    <datalist id="personel_listesi">{% for p in personeller %}<option value="{{ p }}">{% endfor %}</datalist>
                    <label class="form-label" style="font-size:0.8rem;">Tarih</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="tarih" value="{{ k.tarih }}" required>
                    <label class="form-label" style="font-size:0.8rem;">Kaza Yeri</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="yer" value="{{ k.yer }}">
                    <label class="form-label" style="font-size:0.8rem;">Yaralanma Türü</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" name="yaralanma_turu" value="{{ k.yaralanma_turu }}">
                </div>
                <div class="col-md-6 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Yasal ve Sonrası Süreçler</h6>
                    <label class="form-label" style="font-size:0.8rem;">SGK Bildirim Tarihi</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="sgk_bildirim_tarihi" value="{{ k.sgk_bildirim_tarihi }}">
                    <label class="form-label" style="font-size:0.8rem;">Rapor Gün Sayısı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="number" name="rapor_gun_sayisi" value="{{ k.rapor_gun_sayisi }}">
                    <label class="form-label" style="font-size:0.8rem;">Kaza Sonrası Eğitim</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" type="date" name="kaza_sonrasi_egitim" value="{{ k.kaza_sonrasi_egitim }}">
                </div>
                <div class="col-md-12 mb-3">
                    <h6 class="text-secondary border-bottom pb-2 mb-3 fw-bold" style="font-size:0.85rem;">Açıklama ve Analiz</h6>
                    <label class="form-label" style="font-size:0.8rem;">Açıklama</label>
                    <textarea class="form-control mb-2 border-0 shadow-sm form-control-sm" name="aciklama" rows="2">{{ k.aciklama }}</textarea>
                    <label class="form-label" style="font-size:0.8rem;">Kök Neden</label>
                    <textarea class="form-control mb-2 border-0 shadow-sm form-control-sm" name="kok_neden" rows="2">{{ k.kok_neden }}</textarea>
                    <label class="form-label" style="font-size:0.8rem;">Düzeltici Faaliyet</label>
                    <textarea class="form-control border-0 shadow-sm form-control-sm" name="duzeltici_faaliyet" rows="2">{{ k.duzeltici_faaliyet }}</textarea>
                </div>
            </div>
            <button class="btn btn-primary rounded-pill px-4 py-2 mt-2" style="font-size:0.85rem;">Güncelle</button>
            <a href="/kazalar" class="btn btn-light rounded-pill px-4 py-2 mt-2 ms-2" style="font-size:0.85rem;">İptal</a>
        </form>
    </div>
""")

PLAN_EKLE_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <h4 class="mb-3 fw-bold" style="font-size: 1.1rem;">Yeni Plan Ekle</h4>
    <div class="widget-card" style="border-left:none;">
        <form method="post">
            <div class="row mb-3">
                <div class="col-md-12">
                    <label class="form-label fw-bold" style="font-size:0.85rem;">Firma</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" value="{{ session.get('firma_adi') }}" disabled style="background:#e8f4ff; font-weight:600;">
                </div>
            </div>
            <div class="row">
                <div class="col-md-6 mb-3">
                    <div class="p-3 rounded-3" style="background: #f8f9fa; border-left: 4px solid #3498db;">
                        <h6 class="text-primary border-bottom pb-2 mb-2 fw-bold" style="font-size:0.85rem;">Risk Değerlendirmesi</h6>
                        <label class="form-label" style="font-size:0.75rem;">RD Tarihi</label>
                        <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="rd_tarihi">
                        <label class="form-label" style="font-size:0.75rem;">Tehlike Sınıfı</label>
                        <select class="form-select mb-1 border-0 shadow-sm form-select-sm" name="rd_tehlike_sinifi">
                            <option value="">Seçiniz...</option><option value="Az Tehlikeli">Az Tehlikeli (6 Yıl)</option><option value="Tehlikeli">Tehlikeli (4 Yıl)</option><option value="Çok Tehlikeli">Çok Tehlikeli (2 Yıl)</option>
                        </select>
                    </div>
                </div>
                <div class="col-md-6 mb-3">
                    <div class="p-3 rounded-3" style="background: #f8f9fa; border-left: 4px solid #e74c3c;">
                        <h6 class="text-danger border-bottom pb-2 mb-2 fw-bold" style="font-size:0.85rem;">Acil Durum Planı</h6>
                        <label class="form-label" style="font-size:0.75rem;">ADP Tarihi</label>
                        <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="adp_tarihi">
                        <label class="form-label" style="font-size:0.75rem;">Tehlike Sınıfı</label>
                        <select class="form-select mb-1 border-0 shadow-sm form-select-sm" name="adp_tehlike_sinifi">
                            <option value="">Seçiniz...</option><option value="Az Tehlikeli">Az Tehlikeli (6 Yıl)</option><option value="Tehlikeli">Tehlikeli (4 Yıl)</option><option value="Çok Tehlikeli">Çok Tehlikeli (2 Yıl)</option>
                        </select>
                    </div>
                </div>
                <div class="col-md-6 mb-3">
                    <div class="p-3 rounded-3" style="background: #f8f9fa; border-left: 4px solid #1abc9c;">
                        <h6 class="text-info border-bottom pb-2 mb-2 fw-bold" style="font-size:0.85rem;">Eğitim Planı</h6>
                        <label class="form-label" style="font-size:0.75rem;">Plan Tarihi</label>
                        <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="egitim_plani_tarihi">
                        <label class="form-label" style="font-size:0.75rem;">Tehlike Sınıfı</label>
                        <select class="form-select mb-1 border-0 shadow-sm form-select-sm" name="egitim_tehlike_sinifi">
                            <option value="">Seçiniz...</option><option value="Az Tehlikeli">Az Tehlikeli (6 Yıl)</option><option value="Tehlikeli">Tehlikeli (4 Yıl)</option><option value="Çok Tehlikeli">Çok Tehlikeli (2 Yıl)</option>
                        </select>
                    </div>
                </div>
                <div class="col-md-6 mb-3">
                    <div class="p-3 rounded-3" style="background: #f8f9fa; border-left: 4px solid #34495e;">
                        <h6 class="text-dark border-bottom pb-2 mb-2 fw-bold" style="font-size:0.85rem;">Çalışma Planı</h6>
                        <label class="form-label" style="font-size:0.75rem;">Plan Tarihi</label>
                        <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="calisma_plani_tarihi">
                        <label class="form-label" style="font-size:0.75rem;">Tehlike Sınıfı</label>
                        <select class="form-select mb-1 border-0 shadow-sm form-select-sm" name="calisma_tehlike_sinifi">
                            <option value="">Seçiniz...</option><option value="Az Tehlikeli">Az Tehlikeli (6 Yıl)</option><option value="Tehlikeli">Tehlikeli (4 Yıl)</option><option value="Çok Tehlikeli">Çok Tehlikeli (2 Yıl)</option>
                        </select>
                    </div>
                </div>
            </div>
            <button class="btn btn-primary rounded-pill px-4 py-2 mt-2" style="font-size:0.85rem;">Planı Kaydet</button>
            <a href="/planlar" class="btn btn-light rounded-pill px-4 py-2 mt-2 ms-2" style="font-size:0.85rem;">İptal</a>
        </form>
    </div>
""")

PLAN_DUZENLE_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <h4 class="mb-3 fw-bold" style="font-size: 1.1rem;">Plan Düzenle</h4>
    <div class="widget-card" style="border-left:none;">
        <form method="post">
            <div class="row mb-3">
                <div class="col-md-12">
                    <label class="form-label fw-bold" style="font-size:0.85rem;">Firma</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" value="{{ plan.firma_adi }}" disabled style="background:#e8f4ff; font-weight:600;">
                </div>
            </div>
            <div class="row">
                <div class="col-md-6 mb-3">
                    <div class="p-3 rounded-3" style="background: #f8f9fa; border-left: 4px solid #3498db;">
                        <h6 class="text-primary border-bottom pb-2 mb-2 fw-bold" style="font-size:0.85rem;">Risk Değerlendirmesi</h6>
                        <label class="form-label" style="font-size:0.75rem;">RD Tarihi</label>
                        <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="rd_tarihi" value="{{ plan.rd_tarihi }}">
                        <label class="form-label" style="font-size:0.75rem;">Tehlike Sınıfı</label>
                        <select class="form-select mb-1 border-0 shadow-sm form-select-sm" name="rd_tehlike_sinifi">
                            <option value="">Seçiniz...</option>
                            <option value="Az Tehlikeli" {% if plan.rd_tehlike_sinifi == 'Az Tehlikeli' %}selected{% endif %}>Az Tehlikeli (6 Yıl)</option>
                            <option value="Tehlikeli" {% if plan.rd_tehlike_sinifi == 'Tehlikeli' %}selected{% endif %}>Tehlikeli (4 Yıl)</option>
                            <option value="Çok Tehlikeli" {% if plan.rd_tehlike_sinifi == 'Çok Tehlikeli' %}selected{% endif %}>Çok Tehlikeli (2 Yıl)</option>
                        </select>
                    </div>
                </div>
                <div class="col-md-6 mb-3">
                    <div class="p-3 rounded-3" style="background: #f8f9fa; border-left: 4px solid #e74c3c;">
                        <h6 class="text-danger border-bottom pb-2 mb-2 fw-bold" style="font-size:0.85rem;">Acil Durum Planı</h6>
                        <label class="form-label" style="font-size:0.75rem;">ADP Tarihi</label>
                        <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="adp_tarihi" value="{{ plan.adp_tarihi }}">
                        <label class="form-label" style="font-size:0.75rem;">Tehlike Sınıfı</label>
                        <select class="form-select mb-1 border-0 shadow-sm form-select-sm" name="adp_tehlike_sinifi">
                            <option value="">Seçiniz...</option>
                            <option value="Az Tehlikeli" {% if plan.adp_tehlike_sinifi == 'Az Tehlikeli' %}selected{% endif %}>Az Tehlikeli (6 Yıl)</option>
                            <option value="Tehlikeli" {% if plan.adp_tehlike_sinifi == 'Tehlikeli' %}selected{% endif %}>Tehlikeli (4 Yıl)</option>
                            <option value="Çok Tehlikeli" {% if plan.adp_tehlike_sinifi == 'Çok Tehlikeli' %}selected{% endif %}>Çok Tehlikeli (2 Yıl)</option>
                        </select>
                    </div>
                </div>
                <div class="col-md-6 mb-3">
                    <div class="p-3 rounded-3" style="background: #f8f9fa; border-left: 4px solid #1abc9c;">
                        <h6 class="text-info border-bottom pb-2 mb-2 fw-bold" style="font-size:0.85rem;">Eğitim Planı</h6>
                        <label class="form-label" style="font-size:0.75rem;">Plan Tarihi</label>
                        <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="egitim_plani_tarihi" value="{{ plan.egitim_plani_tarihi }}">
                        <label class="form-label" style="font-size:0.75rem;">Tehlike Sınıfı</label>
                        <select class="form-select mb-1 border-0 shadow-sm form-select-sm" name="egitim_tehlike_sinifi">
                            <option value="">Seçiniz...</option>
                            <option value="Az Tehlikeli" {% if plan.egitim_tehlike_sinifi == 'Az Tehlikeli' %}selected{% endif %}>Az Tehlikeli (6 Yıl)</option>
                            <option value="Tehlikeli" {% if plan.egitim_tehlike_sinifi == 'Tehlikeli' %}selected{% endif %}>Tehlikeli (4 Yıl)</option>
                            <option value="Çok Tehlikeli" {% if plan.egitim_tehlike_sinifi == 'Çok Tehlikeli' %}selected{% endif %}>Çok Tehlikeli (2 Yıl)</option>
                        </select>
                    </div>
                </div>
                <div class="col-md-6 mb-3">
                    <div class="p-3 rounded-3" style="background: #f8f9fa; border-left: 4px solid #34495e;">
                        <h6 class="text-dark border-bottom pb-2 mb-2 fw-bold" style="font-size:0.85rem;">Çalışma Planı</h6>
                        <label class="form-label" style="font-size:0.75rem;">Plan Tarihi</label>
                        <input class="form-control mb-2 border-0 shadow-sm form-control-sm" type="date" name="calisma_plani_tarihi" value="{{ plan.calisma_plani_tarihi }}">
                        <label class="form-label" style="font-size:0.75rem;">Tehlike Sınıfı</label>
                        <select class="form-select mb-1 border-0 shadow-sm form-select-sm" name="calisma_tehlike_sinifi">
                            <option value="">Seçiniz...</option>
                            <option value="Az Tehlikeli" {% if plan.calisma_tehlike_sinifi == 'Az Tehlikeli' %}selected{% endif %}>Az Tehlikeli (6 Yıl)</option>
                            <option value="Tehlikeli" {% if plan.calisma_tehlike_sinifi == 'Tehlikeli' %}selected{% endif %}>Tehlikeli (4 Yıl)</option>
                            <option value="Çok Tehlikeli" {% if plan.calisma_tehlike_sinifi == 'Çok Tehlikeli' %}selected{% endif %}>Çok Tehlikeli (2 Yıl)</option>
                        </select>
                    </div>
                </div>
            </div>
            <button class="btn btn-primary rounded-pill px-4 py-2 mt-2" style="font-size:0.85rem;">Güncelle</button>
            <a href="/planlar" class="btn btn-light rounded-pill px-4 py-2 mt-2 ms-2" style="font-size:0.85rem;">İptal</a>
        </form>
    </div>
""")

# ---------------- FİRMA YÖNETİM SAYFASI ----------------
FIRMA_YONETIM_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <div class="d-flex justify-content-between align-items-center mb-3">
        <h3 class="mb-0 fw-bold" style="font-size: 1.1rem;">🏢 Firma Yönetimi</h3>
        <button class="btn btn-primary rounded-pill px-3" data-bs-toggle="modal" data-bs-target="#firmaEkleModal">+ Yeni Firma</button>
    </div>
    
    <div class="widget-card">
        <div class="table-responsive">
            <table class="table table-custom align-middle mb-0">
                <thead style="background:#f8f9fa;">
                    <tr>
                        <th class="ps-3">#</th>
                        <th>Firma Adı</th>
                        <th>Durum</th>
                        <th>Eklenme Tarihi</th>
                        <th class="pe-3 text-end">İşlemler</th>
                    </tr>
                </thead>
                <tbody>
                    {% for f in firmalar_db %}
                    <tr>
                        <td class="ps-3">{{ loop.index }}</td>
                        <td><strong>{{ f.adi }}</strong></td>
                        <td>
                            {% if f.aktif %}
                                <span class="badge bg-success">Aktif</span>
                            {% else %}
                                <span class="badge bg-secondary">Pasif</span>
                            {% endif %}
                        </td>
                        <td>{{ f.olusturma.strftime('%d.%m.%Y') if f.olusturma else '-' }}</td>
                        <td class="pe-3 text-end">
                            <a href="/firma/sil/{{ f.id }}" class="btn btn-sm btn-outline-danger rounded-pill" onclick="return confirm('{{ f.adi }} firmasını silmek istediğinize emin misiniz? Tüm verileri de silinecek!')">Sil</a>
                        </td>
                    </tr>
                    {% else %}
                    <tr><td colspan="5" class="text-center text-muted py-3">Henüz firma eklenmemiş.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>

    <!-- Yeni Firma Modal -->
    <div class="modal fade" id="firmaEkleModal" tabindex="-1">
        <div class="modal-dialog">
            <div class="modal-content">
                <form method="post" action="/firma/ekle">
                    <div class="modal-header">
                        <h5 class="modal-title">Yeni Firma Ekle</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <label class="form-label">Firma Adı</label>
                        <input class="form-control" name="firma_adi" required placeholder="Örn: ABC İnşaat A.Ş.">
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-light" data-bs-dismiss="modal">İptal</button>
                        <button type="submit" class="btn btn-primary">Kaydet</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
""")

# ---------------- TÜM FİRMALAR ÖZET ----------------
TUM_FIRMALAR_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <h3 class="mb-4 fw-bold" style="font-size: 1.2rem;">📊 Tüm Firmalar ve İSG İstatistik Paneli</h3>
    
    <!-- ÜST İSTATİSTİK KARTLARI -->
    <div class="row g-3 mb-4">
        <div class="col-md-3">
            <div class="widget-card text-center" style="border-left: none; border-bottom: 4px solid #3498db;">
                <h4 class="mb-1 fw-bold text-primary">{{ genel_istatistikler.toplam_personel }}</h4>
                <p class="text-muted mb-0" style="font-size: 0.85rem;">Toplam Personel</p>
            </div>
        </div>
        <div class="col-md-3">
            <div class="widget-card text-center" style="border-left: none; border-bottom: 4px solid #e74c3c;">
                <h4 class="mb-1 fw-bold text-danger">{{ genel_istatistikler.toplam_kaza }}</h4>
                <p class="text-muted mb-0" style="font-size: 0.85rem;">Toplam İş Kazası</p>
            </div>
        </div>
        <div class="col-md-3">
            <div class="widget-card text-center" style="border-left: none; border-bottom: 4px solid #1abc9c;">
                <h4 class="mb-1 fw-bold text-success">{{ genel_istatistikler.toplam_egitim }}</h4>
                <p class="text-muted mb-0" style="font-size: 0.85rem;">Verilen Eğitimler</p>
            </div>
        </div>
        <div class="col-md-3">
            <div class="widget-card text-center" style="border-left: none; border-bottom: 4px solid #f39c12;">
                <h4 class="mb-1 fw-bold text-warning">{{ genel_istatistikler.suresi_gecen }}</h4>
                <p class="text-muted mb-0" style="font-size: 0.85rem;">Süresi Geçen Eğitim</p>
            </div>
        </div>
    </div>

    <!-- GRAFİKLER -->
    <div class="row g-3 mb-4">
        <div class="col-md-6">
            <div class="widget-card">
                <h6 class="fw-bold mb-3">Eğitim Durumları Dağılımı</h6>
                <canvas id="egitimGrafik" style="max-height: 250px;"></canvas>
            </div>
        </div>
        <div class="col-md-6">
            <div class="widget-card">
                <h6 class="fw-bold mb-3">Son 6 Ayın Kaza Grafiği</h6>
                <canvas id="kazaGrafik" style="max-height: 250px;"></canvas>
            </div>
        </div>
    </div>
    <!-- İSG PERFORMANS METRİKLERİ -->
    <div class="row g-3 mb-4 mt-2">
        <div class="col-md-12">
            <h5 class="mb-3 fw-bold" style="font-size: 1rem;">📈 İSG Performans Metrikleri</h5>
        </div>
        <div class="col-md-3">
            <div class="widget-card text-center bg-light">
                <h4 class="mb-1 fw-bold text-dark">{{ isg_metrikleri.toplam_rapor_gunu }}</h4>
                <p class="text-muted mb-0" style="font-size: 0.8rem;">Toplam Kayıp Gün</p>
            </div>
        </div>
        <div class="col-md-3">
            <div class="widget-card text-center bg-light">
                <h4 class="mb-1 fw-bold text-dark">%{{ isg_metrikleri.kaza_orani }}</h4>
                <p class="text-muted mb-0" style="font-size: 0.8rem;">Kaza Olabilirlik Oranı</p>
            </div>
        </div>
        <div class="col-md-3">
            <div class="widget-card text-center bg-light">
                <h4 class="mb-1 fw-bold text-dark">{{ isg_metrikleri.ortalama_rapor }} Gün</h4>
                <p class="text-muted mb-0" style="font-size: 0.8rem;">Kaza Başına Ort. İş Göremezlik</p>
            </div>
        </div>
        <div class="col-md-3">
            <div class="widget-card text-center bg-light">
                <h4 class="mb-1 fw-bold text-dark">%{{ isg_metrikleri.egitim_tamamlama_orani }}</h4>
                <p class="text-muted mb-0" style="font-size: 0.8rem;">Eğitim Geçerlilik Oranı</p>
            </div>
        </div>
    </div>
    <!-- FİRMA ÖZET TABLOSU -->
    <h5 class="mb-3 fw-bold" style="font-size: 1rem;">🏢 Firma Bazlı Özet</h5>
    <div class="widget-card p-0 overflow-hidden mb-4" style="border-left:none;">
        <div class="table-responsive">
            <table class="table table-custom align-middle mb-0">
                <thead style="background:#f8f9fa;">
                    <tr>
                        <th class="ps-3" style="font-size:0.8rem;">Firma Adı</th>
                        <th class="text-center" style="font-size:0.8rem;">Personel Sayısı</th>
                        <th class="text-center" style="font-size:0.8rem;">Eğitim</th>
                        <th class="text-center" style="font-size:0.8rem;">Süresi Geçen</th>
                        <th class="text-center" style="font-size:0.8rem;">Yaklaşan (30 Gün)</th>
                        <th class="text-center" style="font-size:0.8rem;">Kaza</th>
                        <th class="text-center pe-3" style="font-size:0.8rem;">İşlem</th>
                    </tr>
                </thead>
                <tbody>
                    {% for f in firma_ozetleri %}
                    <tr>
                        <td class="ps-3"><strong>{{ f.firma }}</strong></td>
                        <td class="text-center"><span class="badge bg-primary">{{ f.personel_sayisi }}</span></td>
                        <td class="text-center">{{ f.egitim_sayisi }}</td>
                        <td class="text-center">
                            {% if f.suresi_gecen > 0 %}<span class="badge bg-danger">{{ f.suresi_gecen }}</span>
                            {% else %}<span class="badge bg-success">0</span>{% endif %}
                        </td>
                        <td class="text-center">
                            {% if f.yaklasan > 0 %}<span class="badge bg-warning text-dark">{{ f.yaklasan }}</span>
                            {% else %}<span class="badge bg-light text-dark">0</span>{% endif %}
                        </td>
                        <td class="text-center">
                            {% if f.kaza_sayisi > 0 %}<span class="badge bg-danger">{{ f.kaza_sayisi }}</span>
                            {% else %}<span class="badge bg-light text-dark">0</span>{% endif %}
                        </td>
                        <td class="text-center pe-3">
                            <a href="/firma_sec/{{ f.firma }}" class="btn btn-sm btn-primary rounded-pill px-3">Git →</a>
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>

    <!-- GRAFİK ÇALIŞTIRMA SCRIPTI -->
    <script>
        const ctxEgitim = document.getElementById('egitimGrafik');
        new Chart(ctxEgitim, {
            type: 'doughnut',
            data: {
                labels: ['Geçerli', 'Yaklaşıyor', 'Süresi Geçmiş'],
                datasets: [{
                    data: [{{ grafik_egitim.gecerli }}, {{ grafik_egitim.yaklasan }}, {{ grafik_egitim.gecmis }}],
                    backgroundColor: ['#1abc9c', '#f39c12', '#e74c3c'],
                    borderWidth: 0
                }]
            },
            options: { cutout: '70%', plugins: { legend: { position: 'bottom' } } }
        });

        const ctxKaza = document.getElementById('kazaGrafik');
        new Chart(ctxKaza, {
            type: 'bar',
            data: {
                labels: {{ grafik_kaza.aylar | safe }},
                datasets: [{
                    label: 'Kaza Sayısı',
                    data: {{ grafik_kaza.sayilar | safe }},
                    backgroundColor: '#e74c3c',
                    borderRadius: 5
                }]
            },
            options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } } }
        });
    </script>
""")
   

# ---------------- ROTALAR ----------------
@app.route('/')
def dashboard():
    bugun = date.today()
    tum_egitimler = firma_filtre(Egitim.query).all()
    yaklasan_egitimler, suresi_gecenler = [], []
    for e in tum_egitimler:
        kalan = (e.son_gecerlilik - bugun).days
        if kalan < 0: suresi_gecenler.append(e)
        elif kalan <= 30: yaklasan_egitimler.append(e)
    yaklasan_egitimler.sort(key=lambda x: x.personel_adi)
    suresi_gecenler.sort(key=lambda x: x.personel_adi)
    risk_planlari = firma_filtre(RiskPlan.query).all()
    uyarilar = []
    for plan in risk_planlari:
        for field, label in [(plan.rd_gecerlilik, 'Risk Değerlendirmesi'), (plan.adp_gecerlilik, 'Acil Durum Planı'), (plan.egitim_plani_gecerlilik, 'Eğitim Planı'), (plan.calisma_plani_gecerlilik, 'Çalışma Planı')]:
            if field:
                kalan = (field - bugun).days
                if kalan < 0:
                    uyarilar.append({'firma': plan.firma_adi, 'mesaj': f'{label} süresi doldu! Güncellenmeli.', 'durum': 'Süresi Geçmiş'})
                elif kalan <= 30:
                    uyarilar.append({'firma': plan.firma_adi, 'mesaj': f'{label} süresi {kalan} gün içinde doluyor.', 'durum': 'Yaklaşıyor'})
    kaza_q = firma_filtre(IsKazasi.query).all()
    son_kazalar = [k for k in kaza_q if 0 <= (bugun - k.tarih).days <= 30]
    for k in son_kazalar:
        uyarilar.append({'firma': k.firma_adi, 'mesaj': f'{k.personel_adi} için son 30 günde iş kazası var. RD gözden geçirilmeli!', 'durum': 'Acil Aksiyon'})
    return render_template_string(DASHBOARD_HTML, yaklasan_egitimler=yaklasan_egitimler, suresi_gecenler=suresi_gecenler, risk_planlari=risk_planlari, bugun=bugun, uyarilar=uyarilar, firmalar=get_firmalar_listesi())

@app.route('/tum_firmalar')
def tum_firmalar():
    bugun = date.today()
    firma_ozetleri = []
    tum_egitimler = Egitim.query.all()
    tum_kazalar = IsKazasi.query.all()
    
    # 1. Firma Bazlı Özet Verileri
    for f in get_firmalar_listesi():
        egitimler = [e for e in tum_egitimler if e.firma_adi == f]
        kazalar = [k for k in tum_kazalar if k.firma_adi == f]
        personeller = set([e.personel_adi for e in egitimler] + [k.personel_adi for k in kazalar])
        
        suresi_gecen = sum(1 for e in egitimler if e.durum == 'SÜRESİ GEÇMİŞ')
        yaklasan = sum(1 for e in egitimler if e.durum == 'YAKLAŞIYOR')
        
        firma_ozetleri.append({
            'firma': f,
            'personel_sayisi': len(personeller),
            'egitim_sayisi': len(egitimler),
            'suresi_gecen': suresi_gecen,
            'yaklasan': yaklasan,
            'kaza_sayisi': len(kazalar)
        })

    # 2. Genel İstatistikler
    tum_personeller = set([e.personel_adi for e in tum_egitimler] + [k.personel_adi for k in tum_kazalar])
    genel_istatistikler = {
        'toplam_personel': len(tum_personeller),
        'toplam_kaza': len(tum_kazalar),
        'toplam_egitim': len(tum_egitimler),
        'suresi_gecen': sum(1 for e in tum_egitimler if e.durum == 'SÜRESİ GEÇMİŞ')
    }

    # 3. Grafik Verileri (Eğitim Pasta Grafiği)
    grafik_egitim = {
        'gecerli': sum(1 for e in tum_egitimler if e.durum == 'GEÇERLİ'),
        'yaklasan': sum(1 for e in tum_egitimler if e.durum == 'YAKLAŞIYOR'),
        'gecmis': sum(1 for e in tum_egitimler if e.durum == 'SÜRESİ GEÇMİŞ')
    }

    # 4. Grafik Verileri (Son 6 Ay Kaza Bar Grafiği)
    aylar = []
    kaza_sayilari = []
    for i in range(5, -1, -1):
        ay_tarihi = bugun - relativedelta(months=i)
        aylar.append(ay_tarihi.strftime('%B'))
        sayi = sum(1 for k in tum_kazalar if k.tarih.year == ay_tarihi.year and k.tarih.month == ay_tarihi.month)
        kaza_sayilari.append(sayi)
        
    grafik_kaza = {
        'aylar': str(aylar).replace("'", '"'),
        'sayilar': str(kaza_sayilari)
    }

    # 5. İSG Performans Metrikleri (Kaza Sıklık, Ağırlık vb.)
    toplam_rapor_gunu = sum(k.rapor_gun_sayisi or 0 for k in tum_kazalar)
    toplam_personel_sayisi = len(tum_personeller)
    
    kaza_orani = (len(tum_kazalar) / toplam_personel_sayisi * 100) if toplam_personel_sayisi > 0 else 0
    ortalama_rapor = (toplam_rapor_gunu / len(tum_kazalar)) if len(tum_kazalar) > 0 else 0
    egitim_tamamlama_orani = (grafik_egitim['gecerli'] / len(tum_egitimler) * 100) if len(tum_egitimler) > 0 else 0

    isg_metrikleri = {
        'toplam_rapor_gunu': toplam_rapor_gunu,
        'kaza_orani': round(kaza_orani, 1),
        'ortalama_rapor': round(ortalama_rapor, 1),
        'egitim_tamamlama_orani': round(egitim_tamamlama_orani, 1)
    }

    return render_template_string(
        TUM_FIRMALAR_HTML, 
        firma_ozetleri=firma_ozetleri, 
        firmalar=get_firmalar_listesi(),
        genel_istatistikler=genel_istatistikler,
        grafik_egitim=grafik_egitim,
        grafik_kaza=grafik_kaza,
        isg_metrikleri=isg_metrikleri
    )

@app.route('/planlar')
def plan_listesi():
    planlar = firma_filtre(RiskPlan.query).all()
    return render_template_string(PLANLAR_HTML, planlar=planlar, bugun=date.today(), firmalar=get_firmalar_listesi())

@app.route('/plan/ekle', methods=['GET', 'POST'])
def plan_ekle():
    if request.method == 'POST':
        firma = aktif_firma()
        if firma == 'TÜM FİRMALAR':
            return "Süper admin modundasınız. Lütfen önce belirli bir firma seçin.", 400
        rd_str = request.form.get('rd_tarihi'); rd_sinif = request.form.get('rd_tehlike_sinifi')
        rd_tarihi = datetime.strptime(rd_str, '%Y-%m-%d').date() if rd_str else None
        rd_gecerlilik = hesapla_gecerlilik(rd_tarihi, rd_sinif)
        adp_str = request.form.get('adp_tarihi'); adp_sinif = request.form.get('adp_tehlike_sinifi')
        adp_tarihi = datetime.strptime(adp_str, '%Y-%m-%d').date() if adp_str else None
        adp_gecerlilik = hesapla_gecerlilik(adp_tarihi, adp_sinif)
        egt_str = request.form.get('egitim_plani_tarihi'); egt_sinif = request.form.get('egitim_tehlike_sinifi')
        egt_tarihi = datetime.strptime(egt_str, '%Y-%m-%d').date() if egt_str else None
        egt_gecerlilik = hesapla_gecerlilik(egt_tarihi, egt_sinif)
        cls_str = request.form.get('calisma_plani_tarihi'); cls_sinif = request.form.get('calisma_tehlike_sinifi')
        cls_tarihi = datetime.strptime(cls_str, '%Y-%m-%d').date() if cls_str else None
        cls_gecerlilik = hesapla_gecerlilik(cls_tarihi, cls_sinif)
        plan = RiskPlan(firma_adi=firma, rd_tarihi=rd_tarihi, rd_tehlike_sinifi=rd_sinif, rd_gecerlilik=rd_gecerlilik, adp_tarihi=adp_tarihi, adp_tehlike_sinifi=adp_sinif, adp_gecerlilik=adp_gecerlilik, egitim_plani_tarihi=egt_tarihi, egitim_tehlike_sinifi=egt_sinif, egitim_plani_gecerlilik=egt_gecerlilik, calisma_plani_tarihi=cls_tarihi, calisma_tehlike_sinifi=cls_sinif, calisma_plani_gecerlilik=cls_gecerlilik)
        db.session.add(plan); db.session.commit()
        return redirect(url_for('plan_listesi'))
    return render_template_string(PLAN_EKLE_HTML, firmalar=get_firmalar_listesi())

@app.route('/plan/duzenle/<int:id>', methods=['GET', 'POST'])
def plan_duzenle(id):
    plan = RiskPlan.query.get_or_404(id)
    if request.method == 'POST':
        rd_str = request.form.get('rd_tarihi')
        plan.rd_tarihi = datetime.strptime(rd_str, '%Y-%m-%d').date() if rd_str else None
        plan.rd_tehlike_sinifi = request.form.get('rd_tehlike_sinifi')
        plan.rd_gecerlilik = hesapla_gecerlilik(plan.rd_tarihi, plan.rd_tehlike_sinifi)
        adp_str = request.form.get('adp_tarihi')
        plan.adp_tarihi = datetime.strptime(adp_str, '%Y-%m-%d').date() if adp_str else None
        plan.adp_tehlike_sinifi = request.form.get('adp_tehlike_sinifi')
        plan.adp_gecerlilik = hesapla_gecerlilik(plan.adp_tarihi, plan.adp_tehlike_sinifi)
        egt_str = request.form.get('egitim_plani_tarihi')
        plan.egitim_plani_tarihi = datetime.strptime(egt_str, '%Y-%m-%d').date() if egt_str else None
        plan.egitim_tehlike_sinifi = request.form.get('egitim_tehlike_sinifi')
        plan.egitim_plani_gecerlilik = hesapla_gecerlilik(plan.egitim_plani_tarihi, plan.egitim_tehlike_sinifi)
        cls_str = request.form.get('calisma_plani_tarihi')
        plan.calisma_plani_tarihi = datetime.strptime(cls_str, '%Y-%m-%d').date() if cls_str else None
        plan.calisma_tehlike_sinifi = request.form.get('calisma_tehlike_sinifi')
        plan.calisma_plani_gecerlilik = hesapla_gecerlilik(plan.calisma_plani_tarihi, plan.calisma_tehlike_sinifi)
        db.session.commit()
        return redirect(url_for('plan_listesi'))
    return render_template_string(PLAN_DUZENLE_HTML, plan=plan, firmalar=get_firmalar_listesi())

@app.route('/plan/sil/<int:id>')
def plan_sil(id):
    plan = RiskPlan.query.get_or_404(id); db.session.delete(plan); db.session.commit()
    return redirect(url_for('plan_listesi'))

@app.route('/egitimler')
def egitim_listesi():
    siralama = request.args.get('siralama', 'yeni')
    query = firma_filtre(Egitim.query)
    if siralama == 'eski': egitimler = query.order_by(Egitim.egitim_tarihi.asc()).all()
    else: egitimler = query.order_by(Egitim.egitim_tarihi.desc()).all()
    kazalar = firma_filtre(IsKazasi.query).all()
    kazalar_dict = {}
    for k in kazalar: 
        p = k.personel_adi.lower().strip()
        kazalar_dict[p] = kazalar_dict.get(p, 0) + 1
    return render_template_string(EGITIMLER_HTML, egitimler=egitimler, siralama=siralama, kazalar_dict=kazalar_dict, firmalar=get_firmalar_listesi())

@app.route('/personel_detay/<path:personel_adi>')
def personel_detay(personel_adi):
    p = personel_adi.lower().strip()
    firma = aktif_firma()
    q_e = Egitim.query; q_k = IsKazasi.query
    if firma != 'TÜM FİRMALAR':
        q_e = q_e.filter(Egitim.firma_adi == firma); q_k = q_k.filter(IsKazasi.firma_adi == firma)
    egitimler = [e for e in q_e.all() if e.personel_adi.lower().strip() == p]
    egitimler.sort(key=lambda x: x.egitim_tarihi, reverse=True)
    kazalar = [k for k in q_k.all() if k.personel_adi.lower().strip() == p]
    kazalar.sort(key=lambda x: x.tarih, reverse=True)
    personel_bilgi = egitimler[0] if egitimler else (kazalar[0] if kazalar else None)
    return render_template_string(PERSONEL_DETAY_HTML, personel_adi=personel_adi, personel_bilgi=personel_bilgi, egitimler=egitimler, kazalar=kazalar, firmalar=get_firmalar_listesi())

@app.route('/egitimler/excel')
def egitim_excel():
    siralama = request.args.get('siralama', 'yeni')
    query = firma_filtre(Egitim.query)
    if siralama == 'eski': egitimler = query.order_by(Egitim.egitim_tarihi.asc()).all()
    else: egitimler = query.order_by(Egitim.egitim_tarihi.desc()).all()
    wb = Workbook(); ws = wb.active; ws.title = "Eğitim Listesi"
    headers = ["Sıra No", "Firma Adı", "Personel", "Görevi", "İşe Giriş", "Eğitim Adı", "Eğitim Tarihi", "Geçerlilik (Ay)", "Son Geçerlilik", "Durum"]
    ws.append(headers)
    for index, e in enumerate(egitimler, start=1):
        ws.append([index, e.firma_adi, e.personel_adi, e.gorevi or '-', e.ise_giris.strftime('%d.%m.%Y') if e.ise_giris else '-', e.egitim_adi, e.egitim_tarihi.strftime('%d.%m.%Y'), e.gecerlilik_ay, e.son_gecerlilik.strftime('%d.%m.%Y'), e.durum])
    output = io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output, download_name=f"{aktif_firma()}_Egitim_Listesi.xlsx".replace(" ", "_"), as_attachment=True)
        
@app.route('/egitim/ekle', methods=['GET', 'POST'])
def egitim_ekle():
    firma = aktif_firma()
    if request.method == 'POST':
        if firma == 'TÜM FİRMALAR': return "Süper admin modundasınız. Lütfen önce belirli bir firma seçin.", 400
        resim_dosyasi = request.files.get('resim'); resim_adi = resim_kaydet(resim_dosyasi)
        ise_giris_str = request.form.get('ise_giris'); ise_giris = datetime.strptime(ise_giris_str, '%Y-%m-%d').date() if ise_giris_str else None
        e = Egitim(firma_adi=firma, personel_adi=request.form['personel_adi'], gorevi=request.form.get('gorevi', ''), ise_giris=ise_giris, resim=resim_adi, egitim_adi=request.form['egitim_adi'], egitim_tarihi=datetime.strptime(request.form['egitim_tarihi'], '%Y-%m-%d').date(), gecerlilik_ay=int(request.form['gecerlilik_ay']), sertifika_no=request.form.get('sertifika_no', ''))
        db.session.add(e); db.session.commit()
        return redirect(url_for('egitim_listesi'))
    return render_template_string(EGITIM_EKLE_HTML, personeller=get_personeller(firma), firmalar=get_firmalar_listesi())

@app.route('/egitim/duzenle/<int:id>', methods=['GET', 'POST'])
def egitim_duzenle(id):
    e = Egitim.query.get_or_404(id)
    firma = aktif_firma()
    if request.method == 'POST':
        e.personel_adi = request.form['personel_adi']; e.gorevi = request.form.get('gorevi', ''); e.egitim_adi = request.form['egitim_adi']
        e.egitim_tarihi = datetime.strptime(request.form['egitim_tarihi'], '%Y-%m-%d').date(); e.gecerlilik_ay = int(request.form['gecerlilik_ay']); e.sertifika_no = request.form.get('sertifika_no', '')
        ise_giris_str = request.form.get('ise_giris'); e.ise_giris = datetime.strptime(ise_giris_str, '%Y-%m-%d').date() if ise_giris_str else None
        if request.form.get('resim_sil') == 'on': resim_sil(e.resim); e.resim = None
        yeni_resim = request.files.get('resim')
        if yeni_resim and yeni_resim.filename != '':
            resim_sil(e.resim); e.resim = resim_kaydet(yeni_resim)
        db.session.commit()
        return redirect(url_for('egitim_listesi'))
    return render_template_string(EGITIM_DUZENLE_HTML, e=e, personeller=get_personeller(firma), firmalar=get_firmalar_listesi())

@app.route('/egitim/sil/<int:id>')
def egitim_sil(id):
    e = Egitim.query.get_or_404(id); resim_sil(e.resim); db.session.delete(e); db.session.commit()
    return redirect(url_for('egitim_listesi'))

@app.route('/kazalar')
def kaza_listesi():
    siralama = request.args.get('siralama', 'yeni')
    query = firma_filtre(IsKazasi.query)
    if siralama == 'eski': kazalar = query.order_by(IsKazasi.tarih.asc()).all()
    else: kazalar = query.order_by(IsKazasi.tarih.desc()).all()
    return render_template_string(KAZALAR_HTML, kazalar=kazalar, siralama=siralama, firmalar=get_firmalar_listesi())

@app.route('/kazalar/excel')
def kaza_excel():
    siralama = request.args.get('siralama', 'yeni')
    query = firma_filtre(IsKazasi.query)
    if siralama == 'eski': kazalar = query.order_by(IsKazasi.tarih.asc()).all()
    else: kazalar = query.order_by(IsKazasi.tarih.desc()).all()
    wb = Workbook(); ws = wb.active; ws.title = "Kaza Listesi"
    headers = ["Sıra No", "Firma", "Personel", "Kaza Tarihi", "Yaralanma Türü", "Kaza Yeri", "SGK Bildirim Tarihi", "Rapor Gün Sayısı", "Kaza Sonrası Eğitim", "Kök Neden", "Düzeltici Faaliyet"]
    ws.append(headers)
    for index, k in enumerate(kazalar, start=1):
        ws.append([index, k.firma_adi, k.personel_adi, k.tarih.strftime('%d.%m.%Y'), k.yaralanma_turu, k.yer, k.sgk_bildirim_tarihi.strftime('%d.%m.%Y') if k.sgk_bildirim_tarihi else '-', k.rapor_gun_sayisi or '-', k.kaza_sonrasi_egitim.strftime('%d.%m.%Y') if k.kaza_sonrasi_egitim else '-', k.kok_neden or '-', k.duzeltici_faaliyet or '-'])
    output = io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output, download_name=f"{aktif_firma()}_Kaza_Listesi.xlsx".replace(" ", "_"), as_attachment=True)

@app.route('/kaza/ekle', methods=['GET', 'POST'])
def kaza_ekle():
    firma = aktif_firma()
    if request.method == 'POST':
        if firma == 'TÜM FİRMALAR': return "Süper admin modundasınız. Lütfen önce belirli bir firma seçin.", 400
        sgk_str = request.form.get('sgk_bildirim_tarihi'); rapor_str = request.form.get('rapor_gun_sayisi'); sonraki_egitim_str = request.form.get('kaza_sonrasi_egitim')
        k = IsKazasi(firma_adi=firma, personel_adi=request.form['personel_adi'], tarih=datetime.strptime(request.form['tarih'], '%Y-%m-%d').date(), yer=request.form['yer'], aciklama=request.form['aciklama'], yaralanma_turu=request.form['yaralanma_turu'], kok_neden=request.form.get('kok_neden', ''), duzeltici_faaliyet=request.form.get('duzeltici_faaliyet', ''), sgk_bildirim_tarihi=datetime.strptime(sgk_str, '%Y-%m-%d').date() if sgk_str else None, rapor_gun_sayisi=int(rapor_str) if rapor_str else None, kaza_sonrasi_egitim=datetime.strptime(sonraki_egitim_str, '%Y-%m-%d').date() if sonraki_egitim_str else None)
        db.session.add(k); db.session.commit()
        return redirect(url_for('kaza_listesi'))
    return render_template_string(KAZA_EKLE_HTML, personeller=get_personeller(firma), firmalar=get_firmalar_listesi())

@app.route('/kaza/duzenle/<int:id>', methods=['GET', 'POST'])
def kaza_duzenle(id):
    k = IsKazasi.query.get_or_404(id)
    if request.method == 'POST':
        k.personel_adi = request.form['personel_adi']; k.tarih = datetime.strptime(request.form['tarih'], '%Y-%m-%d').date()
        k.yer = request.form['yer']; k.aciklama = request.form['aciklama']; k.yaralanma_turu = request.form['yaralanma_turu']
        k.kok_neden = request.form.get('kok_neden', ''); k.duzeltici_faaliyet = request.form.get('duzeltici_faaliyet', '')
        sgk_str = request.form.get('sgk_bildirim_tarihi'); rapor_str = request.form.get('rapor_gun_sayisi'); sonraki_egitim_str = request.form.get('kaza_sonrasi_egitim')
        k.sgk_bildirim_tarihi = datetime.strptime(sgk_str, '%Y-%m-%d').date() if sgk_str else None
        k.rapor_gun_sayisi = int(rapor_str) if rapor_str else None
        k.kaza_sonrasi_egitim = datetime.strptime(sonraki_egitim_str, '%Y-%m-%d').date() if sonraki_egitim_str else None
        db.session.commit()
        return redirect(url_for('kaza_listesi'))
    return render_template_string(KAZA_DUZENLE_HTML, k=k, personeller=get_personeller(aktif_firma()), firmalar=get_firmalar_listesi())

@app.route('/kaza/sil/<int:id>')
def kaza_sil(id):
    k = IsKazasi.query.get_or_404(id); db.session.delete(k); db.session.commit()
    return redirect(url_for('kaza_listesi'))

# ---------------- FİRMA YÖNETİM ROUTE'LARI ----------------
@app.route('/firma_yonetim')
def firma_yonetim():
    firmalar_db = Firma.query.order_by(Firma.adi).all()
    return render_template_string(FIRMA_YONETIM_HTML, firmalar_db=firmalar_db, firmalar=get_firmalar_listesi())

@app.route('/firma/ekle', methods=['POST'])
def firma_ekle_route():
    firma_adi = request.form.get('firma_adi', '').strip()
    if firma_adi:
        if not Firma.query.filter_by(adi=firma_adi).first():
            db.session.add(Firma(adi=firma_adi))
            db.session.commit()
    return redirect(url_for('firma_yonetim'))

@app.route('/firma/sil/<int:id>', methods=['GET', 'POST'])
def firma_sil_route(id):
    firma = Firma.query.get_or_404(id)
    if firma.adi in VARSAYILAN_FIRMALAR:
        return "Varsayılan firma silinemez.", 400
    try:
        Egitim.query.filter_by(firma_adi=firma.adi).delete()
        IsKazasi.query.filter_by(firma_adi=firma.adi).delete()
        RiskPlan.query.filter_by(firma_adi=firma.adi).delete()
        db.session.delete(firma)
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        return f"Silme hatası: {e}", 500
    if session.get('firma_adi') == firma.adi:
        session.pop('firma_adi', None)
    return redirect(url_for('firma_yonetim'))

# ---------------- VERİTABANI BAŞLATMA ----------------
with app.app_context():
    try:
        db.create_all()
        print("✅ Veritabanı tabloları başarıyla oluşturuldu veya zaten mevcut.")
        try:
            get_firmalar_listesi()
            print("✅ Firma listesi hazır.")
        except Exception as fe:
            print(f"⚠️ Firma tohumlama uyarısı: {fe}")
    except Exception as e:
        print(f"❌ Tablo oluşturma hatası: {e}")
# ---------------- PERSONEL LİSTESİ ----------------
PERSONELLER_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <h3 class="mb-0 fw-bold" style="font-size: 1.1rem;">👥 Personel Listesi</h3>
        <form method="get" class="d-flex gap-2 align-items-center m-0">
            <input type="text" name="arama" class="form-control border-0 shadow-sm form-control-sm" placeholder="Personel Ara..." value="{{ arama or '' }}" style="width: 250px; border-radius: 20px;">
            <button type="submit" class="btn btn-primary rounded-pill px-3" style="font-size:0.8rem;"><i class="bi bi-search"></i></button>
        </form>
    </div>
    <div class="widget-card p-0 overflow-hidden" style="border-left:none;">
        <div class="table-responsive">
            <table class="table table-custom align-middle table-hover mb-0">
                <thead style="background:#f8f9fa;">
                    <tr>
                        <th class="ps-3" style="font-size:0.75rem;">Sıra</th>
                        <th style="font-size:0.75rem;">Resim</th>
                        <th style="font-size:0.75rem;">Personel Adı</th>
                        <th style="font-size:0.75rem;">Firma</th>
                        <th style="font-size:0.75rem;">Görevi</th>
                        <th style="font-size:0.75rem;">İşe Giriş</th>
                        <th class="pe-3 text-center" style="font-size:0.75rem;">İşlem</th>
                    </tr>
                </thead>
                <tbody>
                    {% for p in personeller %}
                    <tr class="clickable-row" onclick="window.location='/personel_detay/{{ p.adi }}'">
                        <td class="ps-3">{{ loop.index }}</td>
                        <td>
                            {% if p.resim %}<img src="{{ url_for('static', filename='uploads/' + p.resim) }}" class="rounded-circle" style="width:30px; height:30px; object-fit:cover;">
                            {% else %}<div class="bg-secondary text-white rounded-circle d-flex align-items-center justify-content-center" style="width:30px; height:30px; font-size:10px;">?</div>{% endif %}
                        </td>
                        <td class="fw-bold">{{ p.adi }}</td>
                        <td>{{ p.firma_adi }}</td>
                        <td>{{ p.gorevi or '-' }}</td>
                        <td>{{ p.ise_giris.strftime('%d.%m.%Y') if p.ise_giris else '-' }}</td>
                        <td class="pe-3 text-center">
                            <a href="/personel_detay/{{ p.adi }}" class="btn btn-sm btn-primary rounded-pill px-3">Detay →</a>
                        </td>
                    </tr>
                    {% else %}
                    <tr><td colspan="7" class="text-center text-muted py-3">Kayıtlı personel bulunmuyor.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
""")

@app.route('/personeller')
def personeller_listesi():
    arama = request.args.get('arama', '').strip()
    firma = aktif_firma()
    
    personel_dict = {}
    egitimler = firma_filtre(Egitim.query).all()
    
    for e in egitimler:
        if arama and arama.lower() not in e.personel_adi.lower(): continue
        if e.personel_adi not in personel_dict:
            personel_dict[e.personel_adi] = {
                'adi': e.personel_adi,
                'firma_adi': e.firma_adi,
                'gorevi': e.gorevi,
                'ise_giris': e.ise_giris,
                'resim': e.resim
            }
            
    personeller = list(personel_dict.values())
    personeller.sort(key=lambda x: x['adi'])
    
    return render_template_string(PERSONELLER_HTML, personeller=personeller, arama=arama, firmalar=get_firmalar_listesi())
# ---------------- PDF OLUŞTURMA YARDIMCISI ----------------
def pdf_olustur(html_icerik):
    result = io.BytesIO()
    pdf = pisa.CreatePDF(io.StringIO(html_icerik), dest=result)
    if not pdf.err:
        return result
    return None
# ---------------- DİJİTAL SERTİFİKA HTML ----------------
SERTIFIKA_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <div class="d-flex justify-content-between align-items-center mb-3">
        <h3 class="mb-0 fw-bold" style="font-size: 1.1rem;">📜 Dijital Eğitim Sertifikası</h3>
        <a href="/egitimler" class="btn btn-light rounded-pill px-3"><i class="bi bi-arrow-left"></i> Listeye Dön</a>
    </div>
    <div class="widget-card" style="max-width: 800px; margin: 0 auto; border-top: 8px solid #1e3c72; padding: 40px; background: #fff;">
        <div class="text-center mb-4">
            <h2 class="fw-bold" style="color: #1e3c72; letter-spacing: 2px; font-family: serif;">EĞİTİM SERTİFİKASI</h2>
            <p class="text-muted">Bu belge, aşağıda bilgileri yer alan personelin eğitimi başarıyla tamamladığını onaylar.</p>
        </div>
        <div class="row align-items-center">
            <div class="col-md-8">
                <table class="table table-borderless">
                    <tr><th style="width: 160px; color: #555;">Personel Adı:</th><td style="font-size: 1.1rem;"><strong>{{ egitim.personel_adi }}</strong></td></tr>
                    <tr><th style="color: #555;">Firma:</th><td>{{ egitim.firma_adi }}</td></tr>
                    <tr><th style="color: #555;">Görevi:</th><td>{{ egitim.gorevi or '-' }}</td></tr>
                    <tr><th style="color: #555;">Eğitim Adı:</th><td>{{ egitim.egitim_adi }}</td></tr>
                    <tr><th style="color: #555;">Eğitim Tarihi:</th><td>{{ egitim.egitim_tarihi.strftime('%d.%m.%Y') }}</td></tr>
                    <tr><th style="color: #555;">Son Geçerlilik:</th><td>{{ egitim.son_gecerlilik.strftime('%d.%m.%Y') }}</td></tr>
                    <tr><th style="color: #555;">Durum:</th>
                        <td>
                            {% if egitim.durum == 'GEÇERLİ' %}<span class="badge bg-success" style="font-size: 0.9rem; padding: 8px 15px;">GEÇERLİ</span>
                            {% elif egitim.durum == 'YAKLAŞIYOR' %}<span class="badge bg-warning text-dark" style="font-size: 0.9rem; padding: 8px 15px;">YAKLAŞIYOR</span>
                            {% else %}<span class="badge bg-danger" style="font-size: 0.9rem; padding: 8px 15px;">SÜRESİ GEÇMİŞ</span>{% endif %}
                        </td>
                    </tr>
                </table>
            </div>
            <div class="col-md-4 text-center">
                <img src="data:image/png;base64,{{ qr_kod }}" alt="QR Kod" style="width: 180px; height: 180px; border: 2px solid #ddd; border-radius: 10px; padding: 5px;">
                <p class="text-muted mt-2" style="font-size: 0.75rem;"><i class="bi bi-phone"></i> Doğrulama için okutun</p>
            </div>
        </div>
    </div>
""")
@app.route('/sertifika/<int:id>')
def sertifika_goruntule(id):
    egitim = Egitim.query.get_or_404(id)
    
        # Canlı sitenin adresi (Render'daki linkin)
    CANLI_SITE_URL = "https://isg-takip.onrender.com"  # <-- Kendi Render linkinle değiştir

    dogrulama_url = f"{CANLI_SITE_URL}/sertifika/dogrula/{egitim.id}"
    
    # QR Kodu oluştur
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(dogrulama_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    # QR kodu Base64'e çevir (HTML'de <img> etiketiyle göstermek için)
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    qr_kod_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
    
    return render_template_string(SERTIFIKA_HTML, egitim=egitim, qr_kod=qr_kod_base64, firmalar=get_firmalar_listesi())

@app.route('/sertifika/dogrula/<int:id>')
def sertifika_dogrula(id):
    egitim = Egitim.query.get_or_404(id)
    # Bu sayfa QR okutulduğunda açılır (Sade bir doğrulama sayfası)
    return f"""
    <html><head><title>Sertifika Doğrulama</title><meta charset='utf-8'>
    <style>body {{ font-family: sans-serif; text-align: center; padding-top: 50px; background-color: #f4f7f6; }}
    .card {{ background: white; padding: 40px; border-radius: 15px; display: inline-block; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }}</style>
    </head>
    <body>
        <div class="card">
            <h1 style='color: #1abc9c;'>✅ Sertifika Doğrulandı</h1>
            <h3>{egitim.personel_adi} - {egitim.egitim_adi}</h3>
            <p>Bu sertifika <strong>{egitim.firma_adi}</strong> firmasına ait olup, geçerlilik tarihi <strong>{egitim.son_gecerlilik.strftime('%d.%m.%Y')}</strong> dir.</p>
            <p style="color: #999; font-size: 0.8rem; margin-top: 30px;">İSG TAKİP Sistemi</p>
        </div>
    </body></html>
    """
if __name__ == '__main__':
    app.run(debug=True, port=5000)
