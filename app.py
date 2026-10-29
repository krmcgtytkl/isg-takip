import os
import io
from flask import Flask, render_template_string, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['SECRET_KEY'] = 'isg-gizli-anahtar'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///isg.db'
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
    
    # Risk Değerlendirme
    rd_tehlike_sinifi = db.Column(db.String(50), nullable=True)
    rd_tarihi = db.Column(db.Date, nullable=True)
    rd_gecerlilik = db.Column(db.Date, nullable=True)
    
    # Acil Durum Planı
    adp_tehlike_sinifi = db.Column(db.String(50), nullable=True)
    adp_tarihi = db.Column(db.Date, nullable=True)
    adp_gecerlilik = db.Column(db.Date, nullable=True)
    
    # Eğitim Planı
    egitim_tehlike_sinifi = db.Column(db.String(50), nullable=True)
    egitim_plani_tarihi = db.Column(db.Date, nullable=True)
    egitim_plani_gecerlilik = db.Column(db.Date, nullable=True)
    
    # Çalışma Planı
    calisma_tehlike_sinifi = db.Column(db.String(50), nullable=True)
    calisma_plani_tarihi = db.Column(db.Date, nullable=True)
    calisma_plani_gecerlilik = db.Column(db.Date, nullable=True)

# ---------------- HTML ŞABLONU (YENİ TASARIM) ----------------
BASE_HTML = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>İSG TAKİP</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css">
    <style>
        body {
            font-family: 'Poppins', sans-serif;
            background: #f4f7f6;
            color: #2c3e50;
            overflow-x: hidden;
        }
        .sidebar {
            height: 100vh;
            width: 250px;
            position: fixed;
            top: 0;
            left: 0;
            background: linear-gradient(180deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding-top: 20px;
            border-top-right-radius: 20px;
            border-bottom-right-radius: 20px;
            box-shadow: 4px 0 15px rgba(0,0,0,0.1);
        }
        .sidebar h4 {
            text-align: center;
            font-weight: 700;
            margin-bottom: 30px;
            letter-spacing: 1px;
            font-size: 1.2rem;
        }
        .sidebar a {
            padding: 12px 25px;
            text-decoration: none;
            font-size: 14px;
            color: rgba(255, 255, 255, 0.8);
            display: block;
            transition: 0.3s;
        }
        .sidebar a:hover, .sidebar a.active {
            background: rgba(255, 255, 255, 0.1);
            color: white;
            border-left: 4px solid #fff;
        }
        .sidebar a i { margin-right: 10px; }
        
        .main-content {
            margin-left: 250px;
            padding: 30px;
            min-height: 100vh;
            background: #f4f7f6;
        }
        .top-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
        }
        .top-bar h3 { font-weight: 600; margin: 0; font-size: 1.3rem; }
        
        .widget-card {
            background: white;
            border-radius: 15px;
            padding: 20px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.03);
            transition: transform 0.3s;
            height: 100%;
            border-left: 5px solid transparent;
        }
        .widget-card:hover { transform: translateY(-3px); box-shadow: 0 8px 25px rgba(0,0,0,0.08); }
        
        .widget-rd { border-left-color: #3498db; }
        .widget-adp { border-left-color: #e74c3c; }
        .widget-egitim-plan { border-left-color: #1abc9c; }
        .widget-calisma-plan { border-left-color: #34495e; }
        
        .widget-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }
        .widget-title { font-weight: 600; font-size: 1rem; display: flex; align-items: center; gap: 8px; }
        .widget-icon { font-size: 1.3rem; }
        
        .mini-card {
            background: white;
            border-radius: 10px;
            padding: 12px 15px;
            margin-bottom: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.02);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-left: 4px solid transparent;
        }
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
        
        .alert-panel {
            background: #fff3cd;
            border-left: 5px solid #ffc107;
            border-radius: 10px;
            padding: 15px 20px;
            margin-bottom: 25px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05);
        }
        .alert-panel-title {
            font-weight: 700;
            font-size: 1rem;
            color: #856404;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .alert-item {
            font-size: 0.85rem;
            padding: 6px 0;
            border-bottom: 1px dashed rgba(0,0,0,0.1);
            display: flex;
            justify-content: space-between;
        }
        .alert-item:last-child { border-bottom: none; }
        
        @media (max-width: 768px) {
            .sidebar { width: 100%; height: auto; position: relative; border-radius: 0; }
            .main-content { margin-left: 0; padding: 15px; }
        }
    </style>
</head>
<body>

    <div class="sidebar">
        <h4>🦺 İSG TAKİP</h4>
        <a href="/" class="active"><i class="bi bi-grid-1x2-fill"></i> Panel</a>
        <a href="/egitimler"><i class="bi bi-mortarboard-fill"></i> Eğitimler</a>
        <a href="/egitim/ekle"><i class="bi bi-plus-circle-fill"></i> + Eğitim Ekle</a>
        <a href="/kazalar"><i class="bi bi-exclamation-triangle-fill"></i> Kazalar</a>
        <a href="/kaza/ekle"><i class="bi bi-plus-circle-fill"></i> + Kaza Ekle</a>
        <a href="/planlar"><i class="bi bi-journal-check"></i> Tüm Planlar</a>
        <a href="/plan/ekle"><i class="bi bi-plus-circle-fill"></i> + Plan Ekle</a>
    </div>

    <div class="main-content">
        <div class="top-bar">
            <h3>Dashboard</h3>
            <div class="d-flex align-items-center gap-3">
                <span class="text-muted" style="font-size:0.85rem;">Hoş Geldiniz</span>
                <div class="bg-primary text-white rounded-circle d-flex align-items-center justify-content-center" style="width:35px; height:35px; font-weight:bold; font-size:0.8rem;">İSG</div>
            </div>
        </div>
        {% block content %}{% endblock %}
    </div>

</body>
</html>
"""

# ---------------- SAYFA İÇERİKLERİ ----------------
DASHBOARD_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
    <!-- UYARILAR PANELİ -->
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
        <!-- 1. WIDGET: RİSK DEĞERLENDİRME -->
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

        <!-- 2. WIDGET: ACİL DURUM PLANI -->
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

        <!-- 3. WIDGET: EĞİTİM PLANI -->
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

        <!-- 4. WIDGET: ÇALIŞMA PLANI -->
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

    <!-- EĞİTİM DURUMLARI -->
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
                    <div>
                        <div class="mini-card-title">{{ e.personel_adi }}</div>
                        <div class="mini-card-sub">{{ e.egitim_adi }}</div>
                    </div>
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
                    <div>
                        <div class="mini-card-title">{{ e.personel_adi }}</div>
                        <div class="mini-card-sub">{{ e.egitim_adi }}</div>
                    </div>
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
                <select name="firma" class="form-select border-0 shadow-sm" style="width: auto; font-size:0.8rem; padding: 6px 12px; border-radius: 8px;" onchange="this.form.submit()">
                    <option value="">Tüm Firmalar</option>
                    {% for f in firmalar %}
                        <option value="{{ f }}" {% if secili_firma == f %}selected{% endif %}>{{ f }}</option>
                    {% endfor %}
                </select>
                
                <select name="siralama" class="form-select border-0 shadow-sm" style="width: auto; font-size:0.8rem; padding: 6px 12px; border-radius: 8px;" onchange="this.form.submit()">
                    <option value="yeni" {% if siralama == 'yeni' %}selected{% endif %}>En Yeni</option>
                    <option value="eski" {% if siralama == 'eski' %}selected{% endif %}>En Eski</option>
                </select>
            </form>
            <a href="/egitimler/excel?firma={{ secili_firma }}&siralama={{ siralama }}" class="btn btn-success text-nowrap rounded-pill px-3" style="font-size:0.8rem;"><i class="bi bi-file-earmark-excel"></i> Excel</a>
            <a href="/egitim/ekle" class="btn btn-primary text-nowrap rounded-pill px-3" style="font-size:0.8rem;">+ Eğitim</a>
        </div>
    </div>
    
    <div class="widget-card p-0 overflow-hidden" style="border-left:none;">
        <div class="table-responsive">
            <table class="table table-custom align-middle table-hover mb-0">
                <thead style="background:#f8f9fa;">
                    <tr>
                        <th class="ps-3" style="font-size:0.75rem;">Sıra</th><th style="font-size:0.75rem;">Resim</th><th style="font-size:0.75rem;">Firma</th><th style="font-size:0.75rem;">Personel</th><th style="font-size:0.75rem;">Görevi</th><th style="font-size:0.75rem;">İşe Giriş</th>
                        <th style="font-size:0.75rem;">Eğitim Adı</th><th style="font-size:0.75rem;">Eğitim Tarihi</th><th style="font-size:0.75rem;">Son Geçerlilik</th><th style="font-size:0.75rem;">Durum</th><th style="font-size:0.75rem;">Kaza</th><th class="pe-3" style="font-size:0.75rem;">İşlem</th>
                    </tr>
                </thead>
                <tbody>
                    {% for e in egitimler %}
                    <tr class="clickable-row" onclick="window.location='/personel_detay/{{ e.personel_adi }}'">
                        <td class="ps-3">{{ loop.index }}</td>
                        <td>
                            {% if e.resim %}
                                <img src="{{ url_for('static', filename='uploads/' + e.resim) }}" class="rounded-circle" style="width:30px; height:30px; object-fit:cover;">
                            {% else %}
                                <div class="bg-secondary text-white rounded-circle d-flex align-items-center justify-content-center" style="width:30px; height:30px; font-size:10px;">?</div>
                            {% endif %}
                        </td>
                        <td>{{ e.firma_adi }}</td>
                        <td class="fw-bold">{{ e.personel_adi }}</td>
                        <td>{{ e.gorevi or '-' }}</td>
                        <td>{{ e.ise_giris.strftime('%d.%m.%Y') if e.ise_giris else '-' }}</td>
                        <td>{{ e.egitim_adi }}</td>
                        <td>{{ e.egitim_tarihi.strftime('%d.%m.%Y') }}</td>
                        <td>{{ e.son_gecerlilik.strftime('%d.%m.%Y') }}</td>
                        <td>
                            {% if e.durum == 'SÜRESİ GEÇMİŞ' %}
                                <span class="badge bg-danger badge-glass">Geçmiş</span>
                            {% elif e.durum == 'YAKLAŞIYOR' %}
                                <span class="badge bg-warning text-dark badge-glass">Yaklaşıyor</span>
                            {% else %}
                                <span class="badge bg-success badge-glass">Geçerli</span>
                            {% endif %}
                        </td>
                        <td>
                            <!-- DÜZELTME: Kaza kontrolü için personel adı küçük harflere çevrilerek karşılaştırılıyor -->
                            {% set kaza_var = kazalar_dict.get(e.personel_adi.lower().strip()) %}
                            {% if kaza_var %}
                                <span class="badge bg-danger badge-glass">⚠️ Kaza ({{ kaza_var }})</span>
                            {% else %}
                                <span class="badge bg-light text-dark border badge-glass">Yok</span>
                            {% endif %}
                        </td>
                        <td class="pe-3" onclick="event.stopPropagation();">
                            <a href="/egitim/duzenle/{{ e.id }}" class="btn-glass text-decoration-none text-primary"><i class="bi bi-pencil-square"></i></a>
                            <a href="/egitim/sil/{{ e.id }}" class="btn-glass text-decoration-none text-danger" onclick="return confirm('Silmek istediğinize emin misiniz?')"><i class="bi bi-trash"></i></a>
                        </td>
                    </tr>
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
                <select name="firma" class="form-select border-0 shadow-sm" style="width: auto; font-size:0.8rem; padding: 6px 12px; border-radius: 8px;" onchange="this.form.submit()">
                    <option value="">Tüm Firmalar</option>
                    {% for f in firmalar %}
                        <option value="{{ f }}" {% if secili_firma == f %}selected{% endif %}>{{ f }}</option>
                    {% endfor %}
                </select>

                <select name="siralama" class="form-select border-0 shadow-sm" style="width: auto; font-size:0.8rem; padding: 6px 12px; border-radius: 8px;" onchange="this.form.submit()">
                    <option value="yeni" {% if siralama == 'yeni' %}selected{% endif %}>En Yeni</option>
                    <option value="eski" {% if siralama == 'eski' %}selected{% endif %}>En Eski</option>
                </select>
            </form>
            <a href="/kazalar/excel?firma={{ secili_firma }}&siralama={{ siralama }}" class="btn btn-success text-nowrap rounded-pill px-3" style="font-size:0.8rem;"><i class="bi bi-file-earmark-excel"></i> Excel</a>
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
            {% if kisi and kisi.resim %}
                <img src="{{ url_for('static', filename='uploads/' + kisi.resim) }}" class="rounded-circle me-3" style="width:70px; height:70px; object-fit:cover; border:3px solid #f0f2f5;">
            {% else %}
                <div class="bg-secondary text-white rounded-circle d-flex align-items-center justify-content-center me-3" style="width:70px; height:70px; font-size:24px;">?</div>
            {% endif %}
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
                            <thead style="background:#f8f9fa;">
                                <tr>
                                    <th class="ps-3" style="font-size:0.75rem;">Eğitim Adı</th><th style="font-size:0.75rem;">Eğitim Tarihi</th><th style="font-size:0.75rem;">Geçerlilik (Ay)</th><th style="font-size:0.75rem;">Son Geçerlilik</th><th style="font-size:0.75rem;">Sertifika No</th><th class="pe-3" style="font-size:0.75rem;">Durum</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for e in egitimler %}
                                <tr>
                                    <td class="ps-3">{{ e.egitim_adi }}</td><td>{{ e.egitim_tarihi.strftime('%d.%m.%Y') }}</td><td>{{ e.gecerlilik_ay }}</td><td>{{ e.son_gecerlilik.strftime('%d.%m.%Y') }}</td><td>{{ e.sertifika_no or '-' }}</td>
                                    <td class="pe-3">
                                        {% if e.durum == 'SÜRESİ GEÇMİŞ' %}
                                            <span class="badge bg-danger badge-glass">Süresi Geçmiş</span>
                                        {% elif e.durum == 'YAKLAŞIYOR' %}
                                            <span class="badge bg-warning text-dark badge-glass">Yaklaşıyor</span>
                                        {% else %}
                                            <span class="badge bg-success badge-glass">Geçerli</span>
                                        {% endif %}
                                    </td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            {% else %}
                <div class="alert alert-info py-2" style="font-size:0.8rem;">Bu personelin kayıtlı eğitimi bulunmuyor.</div>
            {% endif %}
        </div>
    </div>

    <div class="row mt-3">
        <div class="col-md-12">
            <h6 class="text-danger border-bottom pb-2 mb-2 fw-bold" style="font-size:0.9rem;"><i class="bi bi-exclamation-triangle-fill"></i> İş Kazası Geçmişi</h6>
            
            {% if kazalar %}
                <div class="widget-card p-0 overflow-hidden" style="border-left:none;">
                    <div class="table-responsive">
                        <table class="table table-custom align-middle mb-0">
                            <thead style="background:#f8f9fa;">
                                <tr>
                                    <th class="ps-3" style="font-size:0.75rem;">Tarih</th><th style="font-size:0.75rem;">Kaza Yeri</th><th style="font-size:0.75rem;">Yaralanma Türü</th><th style="font-size:0.75rem;">Açıklama</th><th style="font-size:0.75rem;">Kök Neden</th><th style="font-size:0.75rem;">Düzeltici Faaliyet</th>
                                    <th style="font-size:0.75rem;">SGK Bildirim</th><th style="font-size:0.75rem;">Rapor Gün</th><th class="pe-3" style="font-size:0.75rem;">Kaza Sonrası Eğitim</th>
                                </tr>
                            </thead>
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
            {% else %}
                <div class="alert alert-success py-2" style="font-size:0.8rem;">
                    <i class="bi bi-check-circle-fill"></i> Bu personelin kayıtlı bir iş kazası bulunmuyor.
                </div>
            {% endif %}
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
                    <label class="form-label" style="font-size:0.8rem;">Firma Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="firma_adi" list="firma_listesi" placeholder="Firma seçin veya yazın" required>
                    <datalist id="firma_listesi">{% for f in firmalar %}<option value="{{ f }}">{% endfor %}</datalist>

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
                    
                    <label class="form-label" style="font-size:0.8rem;">Firma Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="firma_adi" list="firma_listesi" value="{{ e.firma_adi }}" required>
                    <datalist id="firma_listesi">{% for f in firmalar %}<option value="{{ f }}">{% endfor %}</datalist>

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
                    <label class="form-label" style="font-size:0.8rem;">Firma Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="firma_adi" list="firma_listesi" placeholder="Firma seçin veya yazın" required>
                    <datalist id="firma_listesi">{% for f in firmalar %}<option value="{{ f }}">{% endfor %}</datalist>

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
                    <label class="form-label" style="font-size:0.8rem;">Firma Adı</label>
                    <input class="form-control mb-2 border-0 shadow-sm form-control-sm" name="firma_adi" list="firma_listesi" value="{{ k.firma_adi }}" required>
                    <datalist id="firma_listesi">{% for f in firmalar %}<option value="{{ f }}">{% endfor %}</datalist>

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
                    <label class="form-label fw-bold" style="font-size:0.85rem;">Firma Adı</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" name="firma_adi" list="firma_listesi" placeholder="Firma adı seçin veya yazın" required>
                    <datalist id="firma_listesi">{% for f in firmalar %}<option value="{{ f }}">{% endfor %}</datalist>
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
                            <option value="">Seçiniz...</option>
                            <option value="Az Tehlikeli">Az Tehlikeli (6 Yıl)</option>
                            <option value="Tehlikeli">Tehlikeli (4 Yıl)</option>
                            <option value="Çok Tehlikeli">Çok Tehlikeli (2 Yıl)</option>
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
                            <option value="">Seçiniz...</option>
                            <option value="Az Tehlikeli">Az Tehlikeli (6 Yıl)</option>
                            <option value="Tehlikeli">Tehlikeli (4 Yıl)</option>
                            <option value="Çok Tehlikeli">Çok Tehlikeli (2 Yıl)</option>
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
                            <option value="">Seçiniz...</option>
                            <option value="Az Tehlikeli">Az Tehlikeli (6 Yıl)</option>
                            <option value="Tehlikeli">Tehlikeli (4 Yıl)</option>
                            <option value="Çok Tehlikeli">Çok Tehlikeli (2 Yıl)</option>
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
                            <option value="">Seçiniz...</option>
                            <option value="Az Tehlikeli">Az Tehlikeli (6 Yıl)</option>
                            <option value="Tehlikeli">Tehlikeli (4 Yıl)</option>
                            <option value="Çok Tehlikeli">Çok Tehlikeli (2 Yıl)</option>
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
                    <label class="form-label fw-bold" style="font-size:0.85rem;">Firma Adı</label>
                    <input class="form-control border-0 shadow-sm form-control-sm" name="firma_adi" list="firma_listesi" value="{{ plan.firma_adi }}" required>
                    <datalist id="firma_listesi">{% for f in firmalar %}<option value="{{ f }}">{% endfor %}</datalist>
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

# ---------------- YARDIMCI FONKSİYONLAR ----------------
def resim_kaydet(resim_dosyasi):
    if resim_dosyasi and resim_dosyasi.filename != '':
        filename = secure_filename(resim_dosyasi.filename)
        benzersiz_isim = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}"
        resim_dosyasi.save(os.path.join(app.config['UPLOAD_FOLDER'], benzersiz_isim))
        return benzersiz_isim
    return None

def resim_sil(resim_adi):
    if resim_adi:
        dosya_yolu = os.path.join(app.config['UPLOAD_FOLDER'], resim_adi)
        if os.path.exists(dosya_yolu):
            os.remove(dosya_yolu)

def get_firmalar():
    firmalar_egitim = db.session.query(Egitim.firma_adi).distinct().all()
    firmalar_kaza = db.session.query(IsKazasi.firma_adi).distinct().all()
    firmalar_plan = db.session.query(RiskPlan.firma_adi).distinct().all()
    tum_firmalar = set([f[0] for f in firmalar_egitim if f[0]] + [f[0] for f in firmalar_kaza if f[0]] + [f[0] for f in firmalar_plan if f[0]])
    return sorted(list(tum_firmalar))

def get_personeller():
    personeller = db.session.query(Egitim.personel_adi).distinct().all()
    kaza_personeller = db.session.query(IsKazasi.personel_adi).distinct().all()
    tum_personeller = set([p[0] for p in personeller if p[0]] + [p[0] for p in kaza_personeller if p[0]])
    return sorted(list(tum_personeller))

def hesapla_gecerlilik(tarih, tehlike_sinifi):
    if not tarih or not tehlike_sinifi:
        return None
    if tehlike_sinifi == "Az Tehlikeli":
        return tarih + relativedelta(years=6)
    elif tehlike_sinifi == "Tehlikeli":
        return tarih + relativedelta(years=4)
    elif tehlike_sinifi == "Çok Tehlikeli":
        return tarih + relativedelta(years=2)
    return None

# ---------------- ROTALAR ----------------
@app.route('/')
def dashboard():
    bugun = date.today()
    
    tum_egitimler = Egitim.query.all()
    yaklasan_egitimler = []
    suresi_gecenler = []
    
    for e in tum_egitimler:
        kalan = (e.son_gecerlilik - bugun).days
        if kalan < 0: suresi_gecenler.append(e)
        elif kalan <= 30: yaklasan_egitimler.append(e)
            
    yaklasan_egitimler.sort(key=lambda x: x.personel_adi)
    suresi_gecenler.sort(key=lambda x: x.personel_adi)
    risk_planlari = RiskPlan.query.all()
    
    # UYARILAR (30 Gün yaklaşan veya süresi geçmiş planlar)
    uyarilar = []
    for plan in risk_planlari:
        if plan.rd_gecerlilik:
            kalan_rd = (plan.rd_gecerlilik - bugun).days
            if kalan_rd < 0:
                uyarilar.append({'firma': plan.firma_adi, 'mesaj': 'Risk Değerlendirmesi süresi doldu! Güncellenmeli.', 'durum': 'Süresi Geçmiş'})
            elif kalan_rd <= 30:
                uyarilar.append({'firma': plan.firma_adi, 'mesaj': f'Risk Değerlendirmesi süresi {kalan_rd} gün içinde doluyor.', 'durum': 'Yaklaşıyor'})
                
        if plan.adp_gecerlilik:
            kalan_adp = (plan.adp_gecerlilik - bugun).days
            if kalan_adp < 0:
                uyarilar.append({'firma': plan.firma_adi, 'mesaj': 'Acil Durum Planı süresi doldu! Güncellenmeli.', 'durum': 'Süresi Geçmiş'})
            elif kalan_adp <= 30:
                uyarilar.append({'firma': plan.firma_adi, 'mesaj': f'Acil Durum Planı süresi {kalan_adp} gün içinde doluyor.', 'durum': 'Yaklaşıyor'})
        
        if plan.egitim_plani_gecerlilik:
            kalan_egt = (plan.egitim_plani_gecerlilik - bugun).days
            if kalan_egt < 0:
                uyarilar.append({'firma': plan.firma_adi, 'mesaj': 'Eğitim Planı süresi doldu! Güncellenmeli.', 'durum': 'Süresi Geçmiş'})
            elif kalan_egt <= 30:
                uyarilar.append({'firma': plan.firma_adi, 'mesaj': f'Eğitim Planı süresi {kalan_egt} gün içinde doluyor.', 'durum': 'Yaklaşıyor'})
                
        if plan.calisma_plani_gecerlilik:
            kalan_cls = (plan.calisma_plani_gecerlilik - bugun).days
            if kalan_cls < 0:
                uyarilar.append({'firma': plan.firma_adi, 'mesaj': 'Çalışma Planı süresi doldu! Güncellenmeli.', 'durum': 'Süresi Geçmiş'})
            elif kalan_cls <= 30:
                uyarilar.append({'firma': plan.firma_adi, 'mesaj': f'Çalışma Planı süresi {kalan_cls} gün içinde doluyor.', 'durum': 'Yaklaşıyor'})

    # İş Kazası Sonrası RD Güncelleme Uyarısı
    son_kazalar = [k for k in IsKazasi.query.all() if (bugun - k.tarih).days <= 30 and (bugun - k.tarih).days >= 0]
    for k in son_kazalar:
        uyarilar.append({'firma': k.firma_adi, 'mesaj': f'{k.personel_adi} için son 30 günde iş kazası kaydı var. Risk Değerlendirmesi gözden geçirilmeli!', 'durum': 'Acil Aksiyon'})

    return render_template_string(DASHBOARD_HTML, 
        yaklasan_egitimler=yaklasan_egitimler, 
        suresi_gecenler=suresi_gecenler, 
        risk_planlari=risk_planlari, 
        bugun=bugun,
        uyarilar=uyarilar)

@app.route('/planlar')
def plan_listesi():
    planlar = RiskPlan.query.all()
    return render_template_string(PLANLAR_HTML, planlar=planlar, bugun=date.today())

@app.route('/plan/ekle', methods=['GET', 'POST'])
def plan_ekle():
    if request.method == 'POST':
        rd_str = request.form.get('rd_tarihi')
        rd_sinif = request.form.get('rd_tehlike_sinifi')
        rd_tarihi = datetime.strptime(rd_str, '%Y-%m-%d').date() if rd_str else None
        rd_gecerlilik = hesapla_gecerlilik(rd_tarihi, rd_sinif)
        
        adp_str = request.form.get('adp_tarihi')
        adp_sinif = request.form.get('adp_tehlike_sinifi')
        adp_tarihi = datetime.strptime(adp_str, '%Y-%m-%d').date() if adp_str else None
        adp_gecerlilik = hesapla_gecerlilik(adp_tarihi, adp_sinif)

        egitim_plan_str = request.form.get('egitim_plani_tarihi')
        egitim_sinif = request.form.get('egitim_tehlike_sinifi')
        egitim_plan_tarihi = datetime.strptime(egitim_plan_str, '%Y-%m-%d').date() if egitim_plan_str else None
        egitim_plan_gecerlilik = hesapla_gecerlilik(egitim_plan_tarihi, egitim_sinif)

        calisma_plan_str = request.form.get('calisma_plani_tarihi')
        calisma_sinif = request.form.get('calisma_tehlike_sinifi')
        calisma_plan_tarihi = datetime.strptime(calisma_plan_str, '%Y-%m-%d').date() if calisma_plan_str else None
        calisma_plan_gecerlilik = hesapla_gecerlilik(calisma_plan_tarihi, calisma_sinif)
        
        plan = RiskPlan(
            firma_adi=request.form['firma_adi'],
            rd_tarihi=rd_tarihi, rd_tehlike_sinifi=rd_sinif, rd_gecerlilik=rd_gecerlilik,
            adp_tarihi=adp_tarihi, adp_tehlike_sinifi=adp_sinif, adp_gecerlilik=adp_gecerlilik,
            egitim_plani_tarihi=egitim_plan_tarihi, egitim_tehlike_sinifi=egitim_sinif, egitim_plani_gecerlilik=egitim_plan_gecerlilik,
            calisma_plani_tarihi=calisma_plan_tarihi, calisma_tehlike_sinifi=calisma_sinif, calisma_plani_gecerlilik=calisma_plan_gecerlilik
        )
        db.session.add(plan)
        db.session.commit()
        return redirect(url_for('plan_listesi'))
    return render_template_string(PLAN_EKLE_HTML, firmalar=get_firmalar())

@app.route('/plan/duzenle/<int:id>', methods=['GET', 'POST'])
def plan_duzenle(id):
    plan = RiskPlan.query.get_or_404(id)
    if request.method == 'POST':
        plan.firma_adi = request.form['firma_adi']
        
        rd_str = request.form.get('rd_tarihi')
        plan.rd_tarihi = datetime.strptime(rd_str, '%Y-%m-%d').date() if rd_str else None
        plan.rd_tehlike_sinifi = request.form.get('rd_tehlike_sinifi')
        plan.rd_gecerlilik = hesapla_gecerlilik(plan.rd_tarihi, plan.rd_tehlike_sinifi)
        
        adp_str = request.form.get('adp_tarihi')
        plan.adp_tarihi = datetime.strptime(adp_str, '%Y-%m-%d').date() if adp_str else None
        plan.adp_tehlike_sinifi = request.form.get('adp_tehlike_sinifi')
        plan.adp_gecerlilik = hesapla_gecerlilik(plan.adp_tarihi, plan.adp_tehlike_sinifi)

        egitim_plan_str = request.form.get('egitim_plani_tarihi')
        plan.egitim_plani_tarihi = datetime.strptime(egitim_plan_str, '%Y-%m-%d').date() if egitim_plan_str else None
        plan.egitim_tehlike_sinifi = request.form.get('egitim_tehlike_sinifi')
        plan.egitim_plani_gecerlilik = hesapla_gecerlilik(plan.egitim_plani_tarihi, plan.egitim_tehlike_sinifi)

        calisma_plan_str = request.form.get('calisma_plani_tarihi')
        plan.calisma_plani_tarihi = datetime.strptime(calisma_plan_str, '%Y-%m-%d').date() if calisma_plan_str else None
        plan.calisma_tehlike_sinifi = request.form.get('calisma_tehlike_sinifi')
        plan.calisma_plani_gecerlilik = hesapla_gecerlilik(plan.calisma_plani_tarihi, plan.calisma_tehlike_sinifi)
        
        db.session.commit()
        return redirect(url_for('plan_listesi'))
    return render_template_string(PLAN_DUZENLE_HTML, plan=plan, firmalar=get_firmalar())

@app.route('/plan/sil/<int:id>')
def plan_sil(id):
    plan = RiskPlan.query.get_or_404(id)
    db.session.delete(plan)
    db.session.commit()
    return redirect(url_for('plan_listesi'))

@app.route('/egitimler')
def egitim_listesi():
    secili_firma = request.args.get('firma', '')
    siralama = request.args.get('siralama', 'yeni')
    query = Egitim.query
    if secili_firma: query = query.filter_by(firma_adi=secili_firma)
    if siralama == 'eski': egitimler = query.order_by(Egitim.egitim_tarihi.asc()).all()
    else: egitimler = query.order_by(Egitim.egitim_tarihi.desc()).all()
        
    kazalar = IsKazasi.query.all()
    kazalar_dict = {}
    for k in kazalar: 
        p_adi_kucuk = k.personel_adi.lower().strip()
        kazalar_dict[p_adi_kucuk] = kazalar_dict.get(p_adi_kucuk, 0) + 1
        
    return render_template_string(EGITIMLER_HTML, egitimler=egitimler, siralama=siralama, secili_firma=secili_firma, firmalar=get_firmalar(), kazalar_dict=kazalar_dict)

@app.route('/personel_detay/<path:personel_adi>')
def personel_detay(personel_adi):
    p_adi_kucuk = personel_adi.lower().strip()
    tum_egitimler = Egitim.query.all()
    egitimler = [e for e in tum_egitimler if e.personel_adi.lower().strip() == p_adi_kucuk]
    egitimler.sort(key=lambda x: x.egitim_tarihi, reverse=True)
    
    tum_kazalar = IsKazasi.query.all()
    kazalar = [k for k in tum_kazalar if k.personel_adi.lower().strip() == p_adi_kucuk]
    kazalar.sort(key=lambda x: x.tarih, reverse=True)
    
    personel_bilgi = egitimler[0] if egitimler else (kazalar[0] if kazalar else None)
    return render_template_string(PERSONEL_DETAY_HTML, personel_adi=personel_adi, personel_bilgi=personel_bilgi, egitimler=egitimler, kazalar=kazalar)

@app.route('/egitimler/excel')
def egitim_excel():
    secili_firma = request.args.get('firma', '')
    siralama = request.args.get('siralama', 'yeni')
    query = Egitim.query
    if secili_firma: query = query.filter_by(firma_adi=secili_firma)
    if siralama == 'eski': egitimler = query.order_by(Egitim.egitim_tarihi.asc()).all()
    else: egitimler = query.order_by(Egitim.egitim_tarihi.desc()).all()

    wb = Workbook(); ws = wb.active; ws.title = "Eğitim Listesi"
    headers = ["Sıra No", "Firma Adı", "Personel", "Görevi", "İşe Giriş", "Eğitim Adı", "Eğitim Tarihi", "Geçerlilik (Ay)", "Son Geçerlilik", "Durum"]
    ws.append(headers)
    for index, e in enumerate(egitimler, start=1):
        ws.append([index, e.firma_adi, e.personel_adi, e.gorevi or '-', e.ise_giris.strftime('%d.%m.%Y') if e.ise_giris else '-', e.egitim_adi, e.egitim_tarihi.strftime('%d.%m.%Y'), e.gecerlilik_ay, e.son_gecerlilik.strftime('%d.%m.%Y'), e.durum])
    output = io.BytesIO(); wb.save(output); output.seek(0)
    dosya_adi = f"{secili_firma}_Egitim_Listesi.xlsx" if secili_firma else "Tum_Firmalar_Egitim_Listesi.xlsx"
    return send_file(output, download_name=dosya_adi, as_attachment=True)

@app.route('/egitim/ekle', methods=['GET', 'POST'])
def egitim_ekle():
    if request.method == 'POST':
        resim_dosyasi = request.files.get('resim'); resim_adi = resim_kaydet(resim_dosyasi)
        ise_giris_str = request.form.get('ise_giris'); ise_giris = datetime.strptime(ise_giris_str, '%Y-%m-%d').date() if ise_giris_str else None
        e = Egitim(firma_adi=request.form['firma_adi'], personel_adi=request.form['personel_adi'], gorevi=request.form.get('gorevi', ''), ise_giris=ise_giris, resim=resim_adi, egitim_adi=request.form['egitim_adi'], egitim_tarihi=datetime.strptime(request.form['egitim_tarihi'], '%Y-%m-%d').date(), gecerlilik_ay=int(request.form['gecerlilik_ay']), sertifika_no=request.form.get('sertifika_no', ''))
        db.session.add(e); db.session.commit()
        return redirect(url_for('egitim_listesi'))
    return render_template_string(EGITIM_EKLE_HTML, firmalar=get_firmalar(), personeller=get_personeller())

@app.route('/egitim/duzenle/<int:id>', methods=['GET', 'POST'])
def egitim_duzenle(id):
    e = Egitim.query.get_or_404(id)
    if request.method == 'POST':
        e.firma_adi = request.form['firma_adi']; e.personel_adi = request.form['personel_adi']; e.gorevi = request.form.get('gorevi', ''); e.egitim_adi = request.form['egitim_adi']
        e.egitim_tarihi = datetime.strptime(request.form['egitim_tarihi'], '%Y-%m-%d').date(); e.gecerlilik_ay = int(request.form['gecerlilik_ay']); e.sertifika_no = request.form.get('sertifika_no', '')
        ise_giris_str = request.form.get('ise_giris'); e.ise_giris = datetime.strptime(ise_giris_str, '%Y-%m-%d').date() if ise_giris_str else None
        if request.form.get('resim_sil') == 'on': resim_sil(e.resim); e.resim = None
        yeni_resim = request.files.get('resim')
        if yeni_resim and yeni_resim.filename != '':
            resim_sil(e.resim); e.resim = resim_kaydet(yeni_resim)
        db.session.commit()
        return redirect(url_for('egitim_listesi'))
    return render_template_string(EGITIM_DUZENLE_HTML, e=e, firmalar=get_firmalar(), personeller=get_personeller())

@app.route('/egitim/sil/<int:id>')
def egitim_sil(id):
    e = Egitim.query.get_or_404(id); resim_sil(e.resim); db.session.delete(e); db.session.commit()
    return redirect(url_for('egitim_listesi'))

@app.route('/kazalar')
def kaza_listesi():
    secili_firma = request.args.get('firma', '')
    siralama = request.args.get('siralama', 'yeni')
    
    query = IsKazasi.query
    if secili_firma: query = query.filter_by(firma_adi=secili_firma)
    if siralama == 'eski': kazalar = query.order_by(IsKazasi.tarih.asc()).all()
    else: kazalar = query.order_by(IsKazasi.tarih.desc()).all()
        
    return render_template_string(KAZALAR_HTML, kazalar=kazalar, siralama=siralama, secili_firma=secili_firma, firmalar=get_firmalar())

@app.route('/kazalar/excel')
def kaza_excel():
    secili_firma = request.args.get('firma', '')
    siralama = request.args.get('siralama', 'yeni')
    
    query = IsKazasi.query
    if secili_firma: query = query.filter_by(firma_adi=secili_firma)
    if siralama == 'eski': kazalar = query.order_by(IsKazasi.tarih.asc()).all()
    else: kazalar = query.order_by(IsKazasi.tarih.desc()).all()
        
    wb = Workbook(); ws = wb.active; ws.title = "Kaza Listesi"
    headers = ["Sıra No", "Firma", "Personel", "Kaza Tarihi", "Yaralanma Türü", "Kaza Yeri", "SGK Bildirim Tarihi", "Rapor Gün Sayısı", "Kaza Sonrası Eğitim", "Kök Neden", "Düzeltici Faaliyet"]
    ws.append(headers)
    for index, k in enumerate(kazalar, start=1):
        ws.append([
            index, k.firma_adi, k.personel_adi, k.tarih.strftime('%d.%m.%Y'), k.yaralanma_turu, k.yer,
            k.sgk_bildirim_tarihi.strftime('%d.%m.%Y') if k.sgk_bildirim_tarihi else '-', k.rapor_gun_sayisi or '-',
            k.kaza_sonrasi_egitim.strftime('%d.%m.%Y') if k.kaza_sonrasi_egitim else '-',
            k.kok_neden or '-', k.duzeltici_faaliyet or '-'
        ])
    output = io.BytesIO(); wb.save(output); output.seek(0)
    dosya_adi = f"{secili_firma}_Kaza_Listesi.xlsx" if secili_firma else "Tum_Firmalar_Kaza_Listesi.xlsx"
    return send_file(output, download_name=dosya_adi, as_attachment=True)

@app.route('/kaza/ekle', methods=['GET', 'POST'])
def kaza_ekle():
    if request.method == 'POST':
        sgk_str = request.form.get('sgk_bildirim_tarihi'); rapor_str = request.form.get('rapor_gun_sayisi'); sonraki_egitim_str = request.form.get('kaza_sonrasi_egitim')
        k = IsKazasi(
            firma_adi=request.form['firma_adi'], personel_adi=request.form['personel_adi'], tarih=datetime.strptime(request.form['tarih'], '%Y-%m-%d').date(),
            yer=request.form['yer'], aciklama=request.form['aciklama'], yaralanma_turu=request.form['yaralanma_turu'],
            kok_neden=request.form.get('kok_neden', ''), duzeltici_faaliyet=request.form.get('duzeltici_faaliyet', ''),
            sgk_bildirim_tarihi=datetime.strptime(sgk_str, '%Y-%m-%d').date() if sgk_str else None,
            rapor_gun_sayisi=int(rapor_str) if rapor_str else None,
            kaza_sonrasi_egitim=datetime.strptime(sonraki_egitim_str, '%Y-%m-%d').date() if sonraki_egitim_str else None
        )
        db.session.add(k); db.session.commit()
        return redirect(url_for('kaza_listesi'))
    return render_template_string(KAZA_EKLE_HTML, firmalar=get_firmalar(), personeller=get_personeller())

@app.route('/kaza/duzenle/<int:id>', methods=['GET', 'POST'])
def kaza_duzenle(id):
    k = IsKazasi.query.get_or_404(id)
    if request.method == 'POST':
        k.firma_adi = request.form['firma_adi']; k.personel_adi = request.form['personel_adi']; k.tarih = datetime.strptime(request.form['tarih'], '%Y-%m-%d').date()
        k.yer = request.form['yer']; k.aciklama = request.form['aciklama']; k.yaralanma_turu = request.form['yaralanma_turu']
        k.kok_neden = request.form.get('kok_neden', ''); k.duzeltici_faaliyet = request.form.get('duzeltici_faaliyet', '')
        sgk_str = request.form.get('sgk_bildirim_tarihi'); rapor_str = request.form.get('rapor_gun_sayisi'); sonraki_egitim_str = request.form.get('kaza_sonrasi_egitim')
        k.sgk_bildirim_tarihi = datetime.strptime(sgk_str, '%Y-%m-%d').date() if sgk_str else None
        k.rapor_gun_sayisi = int(rapor_str) if rapor_str else None
        k.kaza_sonrasi_egitim = datetime.strptime(sonraki_egitim_str, '%Y-%m-%d').date() if sonraki_egitim_str else None
        db.session.commit()
        return redirect(url_for('kaza_listesi'))
    return render_template_string(KAZA_DUZENLE_HTML, k=k, firmalar=get_firmalar(), personeller=get_personeller())

@app.route('/kaza/sil/<int:id>')
def kaza_sil(id):
    k = IsKazasi.query.get_or_404(id); db.session.delete(k); db.session.commit()
    return redirect(url_for('kaza_listesi'))

# Veritabanı tablolarını uygulama başlarken otomatik oluştur
with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
