# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
import threading
from dotenv import load_dotenv
from entegra_cek import excel_cek, INDIRME_KLASORU, durum_mesaj as entegra_canlı_mesaj
import entegra_cek

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# Entegra çekme durumu (thread-safe)
entegra_durum = {'durum': 'bosta', 'mesaj': '', 'dosya': None, 'detay': None}
# Genel siparişler için ayrı durum
genel_entegra_durum = {'durum': 'bosta', 'mesaj': '', 'dosya': None, 'detay': None}

def sutun_indeksi(sutun_adi):
    """Excel sütun adını indekse çevirir"""
    indeks = 0
    for i, harf in enumerate(reversed(sutun_adi.upper())):
        indeks += (ord(harf) - ord('A') + 1) * (26 ** i)
    return indeks - 1


def basliga_gore_sutun_bul(df, aranan_basliklar):
    """İlk satırdaki başlıklara göre sütun indeksini bulur.
    aranan_basliklar: aranacak başlık isimlerinin listesi (ilk eşleşen döner)
    Bulunamazsa None döner.
    """
    for col_idx in range(df.shape[1]):
        baslik = str(df.iloc[0, col_idx]).strip()
        for aranan in aranan_basliklar:
            if baslik == aranan:
                return col_idx
    return None

def analiz_yap(df):
    """Excel verisini analiz eder"""
    bn_idx = sutun_indeksi('BN')  # Ürün adı
    bs_idx = sutun_indeksi('BS')  # Sipariş adedi
    c_idx = sutun_indeksi('C')    # Sipariş numarası
    cg_idx = sutun_indeksi('CG')  # Platform (Trendyol filtresi)
    s_idx = sutun_indeksi('S')    # Durum (Kargoya verilecek filtresi)

    if df.shape[1] <= max(bn_idx, bs_idx, c_idx, cg_idx):
        return None, "Excel dosyasında yeterli sütun yok!"

    urun_sutunu = df.iloc[:, bn_idx]
    adet_sutunu = df.iloc[:, bs_idx]
    siparis_sutunu = df.iloc[:, c_idx]
    platform_sutunu = df.iloc[:, cg_idx]
    durum_sutunu = df.iloc[:, s_idx]

    siparis_detay = {}  # Sipariş numarasına göre ürünleri grupla

    # Önce sipariş detaylarını oluştur (aynı üründen birden fazla satır varsa birleştir)
    for i in range(len(df)):
        urun = str(urun_sutunu.iloc[i]).strip()
        siparis_no = str(siparis_sutunu.iloc[i]).strip()

        if urun == '' or urun == 'nan' or pd.isna(urun_sutunu.iloc[i]):
            continue

        # Başlık satırını atla
        if urun == 'Ürün İsmi':
            continue

        # CG sütunu: Sadece Trendyol ve trendyol.micro siparişleri
        platform = str(platform_sutunu.iloc[i]).strip().lower()
        if 'trendyol' not in platform and 'trendyol.micro' not in platform:
            continue

        # S sütunu: Sadece "Kargoya verilecek" durumundakiler
        durum = str(durum_sutunu.iloc[i]).strip().lower()
        if 'kargoya verilecek' not in durum:
            continue

        try:
            adet = int(float(adet_sutunu.iloc[i]))
        except (ValueError, TypeError):
            continue

        # Sipariş detayı (aynı ürünleri birleştir)
        if siparis_no and siparis_no != 'nan':
            if siparis_no not in siparis_detay:
                siparis_detay[siparis_no] = []
            # Aynı ürün var mı kontrol et, varsa adetini artır
            urun_bulundu = False
            for item in siparis_detay[siparis_no]:
                if item['urun'] == urun:
                    item['adet'] += adet
                    urun_bulundu = True
                    break
            if not urun_bulundu:
                siparis_detay[siparis_no].append({
                    'urun': urun,
                    'adet': adet
                })

    # Sipariş detaylarından ürün özetini oluştur (birleştirilmiş adetlerle)
    urun_ozeti = {}
    for siparis_no, urunler in siparis_detay.items():
        for u in urunler:
            urun = u['urun']
            adet = u['adet']
            if urun not in urun_ozeti:
                urun_ozeti[urun] = {
                    'toplam_adet': 0,
                    'siparis_sayisi': 0,
                    'paketler': {}
                }
            urun_ozeti[urun]['toplam_adet'] += adet
            urun_ozeti[urun]['siparis_sayisi'] += 1
            if adet in urun_ozeti[urun]['paketler']:
                urun_ozeti[urun]['paketler'][adet] += 1
            else:
                urun_ozeti[urun]['paketler'][adet] = 1

    # Karma siparişleri bul (birden fazla ürün içeren)
    karma_siparisler_raw = []
    karma_urun_adetleri = {}  # Karma siparişlerdeki ürün adetlerini takip et

    for siparis_no, urunler in siparis_detay.items():
        if len(urunler) > 1:
            karma_siparisler_raw.append({
                'siparis_no': siparis_no,
                'urunler': urunler
            })
            # Karma siparişlerdeki ürün adetlerini topla
            for u in urunler:
                urun_adi = u['urun']
                adet = u['adet']
                if urun_adi not in karma_urun_adetleri:
                    karma_urun_adetleri[urun_adi] = {
                        'toplam_adet': 0,
                        'siparis_sayisi': 0,
                        'paketler': {}
                    }
                karma_urun_adetleri[urun_adi]['toplam_adet'] += adet
                karma_urun_adetleri[urun_adi]['siparis_sayisi'] += 1
                if adet in karma_urun_adetleri[urun_adi]['paketler']:
                    karma_urun_adetleri[urun_adi]['paketler'][adet] += 1
                else:
                    karma_urun_adetleri[urun_adi]['paketler'][adet] = 1

    # Aynı içerikli karma siparişleri grupla
    karma_gruplar = {}
    for siparis in karma_siparisler_raw:
        # İçeriği anahtar olarak kullan (ürün adı ve adet sıralı)
        icerik_key = tuple(sorted((u['urun'], u['adet']) for u in siparis['urunler']))
        if icerik_key not in karma_gruplar:
            karma_gruplar[icerik_key] = {
                'urunler': siparis['urunler'],
                'siparis_nolar': [],
                'adet': 0
            }
        karma_gruplar[icerik_key]['siparis_nolar'].append(siparis['siparis_no'])
        karma_gruplar[icerik_key]['adet'] += 1

    # Grupları listeye dönüştür
    karma_siparisler = []
    for icerik_key, grup in karma_gruplar.items():
        karma_siparisler.append({
            'urunler': grup['urunler'],
            'siparis_nolar': grup['siparis_nolar'],
            'adet': grup['adet']
        })

    # Karma siparişlerdeki adetleri ana özetten çıkar
    for urun_adi, karma_bilgi in karma_urun_adetleri.items():
        if urun_adi in urun_ozeti:
            urun_ozeti[urun_adi]['toplam_adet'] -= karma_bilgi['toplam_adet']
            urun_ozeti[urun_adi]['siparis_sayisi'] -= karma_bilgi['siparis_sayisi']
            # Paketlerden de çıkar
            for adet, sayi in karma_bilgi['paketler'].items():
                if adet in urun_ozeti[urun_adi]['paketler']:
                    urun_ozeti[urun_adi]['paketler'][adet] -= sayi
                    if urun_ozeti[urun_adi]['paketler'][adet] <= 0:
                        del urun_ozeti[urun_adi]['paketler'][adet]

    # Sonuçları liste olarak döndür
    sonuclar = []
    toplam_siparis = 0
    toplam_urun = 0

    for urun in sorted(urun_ozeti.keys()):
        bilgi = urun_ozeti[urun]

        # Tekli siparişi kalmayan ürünleri atla
        if bilgi['toplam_adet'] <= 0:
            continue

        toplam_siparis += bilgi['siparis_sayisi']
        toplam_urun += bilgi['toplam_adet']

        paket_listesi = []
        for adet in sorted(bilgi['paketler'].keys()):
            if bilgi['paketler'][adet] > 0:
                paket_listesi.append({
                    'adet': adet,
                    'sayi': bilgi['paketler'][adet]
                })

        sonuclar.append({
            'urun': urun,
            'toplam': bilgi['toplam_adet'],
            'siparis_sayisi': bilgi['siparis_sayisi'],
            'paketler': paket_listesi
        })

    # Karma siparişlerdeki toplam ürün sayısını hesapla
    karma_toplam_urun = sum(
        sum(u['adet'] for u in siparis['urunler'])
        for siparis in karma_siparisler
    )

    ozet = {
        'urun_cesidi': len([u for u in urun_ozeti.keys() if urun_ozeti[u]['toplam_adet'] > 0]),
        'toplam_siparis': toplam_siparis + len(karma_siparisler),
        'toplam_urun': toplam_urun + karma_toplam_urun,
        'karma_siparis_sayisi': len(karma_siparisler)
    }

    return {'urunler': sonuclar, 'ozet': ozet, 'karma_siparisler': karma_siparisler}, None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analiz', methods=['POST'])
def analiz():
    if 'file' not in request.files:
        return jsonify({'error': 'Dosya seçilmedi!'})

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Dosya seçilmedi!'})

    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': 'Sadece .xlsx dosyaları desteklenir!'})

    try:
        df = pd.read_excel(file, header=None)
        sonuc, hata = analiz_yap(df)

        if hata:
            return jsonify({'error': hata})

        return jsonify(sonuc)
    except Exception as e:
        return jsonify({'error': f'Hata: {str(e)}'})

@app.route('/genel-analiz', methods=['POST'])
def genel_analiz():
    """Genel sipariş için Excel'den barkod verilerini çıkarır"""
    if 'file' not in request.files:
        return jsonify({'error': 'Dosya seçilmedi!'})

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Dosya seçilmedi!'})

    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': 'Sadece .xlsx dosyaları desteklenir!'})

    try:
        df = pd.read_excel(file, header=None)

        # Sütun indeksleri: AN=barkod, BN=ürün ismi, BS=adet
        an_idx = sutun_indeksi('AN')  # Barkod
        bn_idx = sutun_indeksi('BN')  # Ürün adı
        bs_idx = sutun_indeksi('BS')  # Adet
        cg_idx = sutun_indeksi('CG')  # Platform (Trendyol filtresi)
        s_idx = sutun_indeksi('S')    # Durum (Kargoya verilecek filtresi)

        if df.shape[1] <= max(an_idx, bn_idx, bs_idx, cg_idx):
            return jsonify({'error': 'Excel dosyasında yeterli sütun yok!'})

        barkod_sutunu = df.iloc[:, an_idx]
        urun_sutunu = df.iloc[:, bn_idx]
        adet_sutunu = df.iloc[:, bs_idx]
        platform_sutunu = df.iloc[:, cg_idx]
        durum_sutunu = df.iloc[:, s_idx]

        barkodlar = {}

        for i in range(len(df)):
            barkod = str(barkod_sutunu.iloc[i]).strip()
            urun = str(urun_sutunu.iloc[i]).strip()

            if barkod == '' or barkod == 'nan' or pd.isna(barkod_sutunu.iloc[i]):
                continue

            # Başlık satırını atla
            if 'Barkod' in barkod or 'barkod' in barkod:
                continue

            # CG sütunu: Sadece Trendyol ve trendyol.micro siparişleri
            platform = str(platform_sutunu.iloc[i]).strip().lower()
            if 'trendyol' not in platform and 'trendyol.micro' not in platform:
                continue

            # S sütunu: Sadece "Kargoya verilecek" durumundakiler
            durum = str(durum_sutunu.iloc[i]).strip().lower()
            if 'kargoya verilecek' not in durum:
                continue

            try:
                adet = int(float(adet_sutunu.iloc[i]))
            except (ValueError, TypeError):
                adet = 1  # Varsayılan 1 adet

            # Bir barkodda birden fazla ürün olabilir
            if barkod not in barkodlar:
                barkodlar[barkod] = []

            barkodlar[barkod].append({
                'urun': urun,
                'adet': adet
            })

        return jsonify({'barkodlar': barkodlar})

    except Exception as e:
        return jsonify({'error': f'Hata: {str(e)}'})

@app.route('/entegra-cek', methods=['POST'])
def entegra_cek_route():
    """Entegra'dan Selenium ile Excel çeker"""
    global entegra_durum

    if entegra_durum['durum'] == 'calisiyor':
        return jsonify({'error': 'Zaten bir indirme işlemi devam ediyor!'})

    data = request.get_json()
    email = data.get('email', '').strip()
    sifre = data.get('sifre', '').strip()

    # .env'den oku yoksa formdan al
    if not email:
        email = os.getenv('ENTEGRA_EMAIL', '')
    if not sifre:
        sifre = os.getenv('ENTEGRA_SIFRE', '')

    if not email or not sifre:
        return jsonify({'error': 'E-posta ve şifre gerekli!'})

    entegra_durum = {'durum': 'calisiyor', 'mesaj': 'Entegra paneline bağlanılıyor... CAPTCHA cikarsa acilan Chrome penceresinde manuel tiklayin.', 'dosya': None, 'detay': None}

    def cek_thread():
        global entegra_durum
        try:
            sonuc = excel_cek(email, sifre, headless=False)
            if sonuc['basarili']:
                entegra_durum = {
                    'durum': 'tamamlandi',
                    'mesaj': sonuc['mesaj'],
                    'dosya': sonuc['dosya'],
                    'detay': None
                }
            else:
                entegra_durum = {
                    'durum': 'hata',
                    'mesaj': sonuc['mesaj'],
                    'dosya': None,
                    'detay': sonuc.get('sayfa_bilgisi')
                }
        except Exception as e:
            entegra_durum = {
                'durum': 'hata',
                'mesaj': f'Beklenmeyen hata: {str(e)}',
                'dosya': None,
                'detay': None
            }

    thread = threading.Thread(target=cek_thread)
    thread.start()

    return jsonify({'mesaj': 'İndirme başlatıldı, tarayıcı açılacak...'})


@app.route('/entegra-durum')
def entegra_durum_route():
    """Entegra indirme durumunu döndürür"""
    # Çalışıyor durumundaysa canlı mesajı da ekle
    sonuc = dict(entegra_durum)
    if sonuc['durum'] == 'calisiyor' and entegra_cek.durum_mesaj:
        sonuc['mesaj'] = entegra_cek.durum_mesaj
    return jsonify(sonuc)


def entegra_analiz_yap(df, durum_filtre='kargoya_verilecek'):
    """Entegra Excel formatını analiz eder - sütun isimlerine göre dinamik tespit yapar

    Args:
        df: Excel DataFrame
        durum_filtre: 'kargoya_verilecek', 'yeni_siparis' veya 'hepsi'

    Desteklenen formatlar:
    - Ayrıntılı Excel (Türkçe başlıklar): Platform Referans No, Entegrasyon, Pazaryeri Durumu, Ürün İsmi, Adet
    - Normal Excel (İngilizce başlıklar): order_number, entegration, store_order_status_name, product_name, total_product_quantity
    """
    # İlk satıra bakarak formatı algıla
    ilk_hucre = str(df.iloc[0, 0]).strip()

    if ilk_hucre == 'ID' or 'Tarih' in str(df.iloc[0, 1]):
        # Format 1 - Ayrıntılı Excel (Türkçe) - başlık ismine göre bul
        siparis_idx = basliga_gore_sutun_bul(df, ['Platform Referans No'])
        platform_idx = basliga_gore_sutun_bul(df, ['Entegrasyon'])
        durum_idx = basliga_gore_sutun_bul(df, ['Pazaryeri Durumu'])
        urun_idx = basliga_gore_sutun_bul(df, ['Ürün İsmi', 'Ürün ismi'])
        adet_idx = basliga_gore_sutun_bul(df, ['Adet'])

        eksik = []
        if siparis_idx is None: eksik.append('Platform Referans No')
        if platform_idx is None: eksik.append('Entegrasyon')
        if durum_idx is None: eksik.append('Pazaryeri Durumu')
        if urun_idx is None: eksik.append('Ürün İsmi')
        if adet_idx is None: eksik.append('Adet')
        if eksik:
            return None, f"Excel'de şu sütunlar bulunamadı: {', '.join(eksik)}"
    else:
        # Format 2 - Normal Excel (İngilizce) - başlık ismine göre bul
        siparis_idx = basliga_gore_sutun_bul(df, ['order_number'])
        platform_idx = basliga_gore_sutun_bul(df, ['entegration'])
        durum_idx = basliga_gore_sutun_bul(df, ['store_order_status_name'])
        urun_idx = basliga_gore_sutun_bul(df, ['product_name'])
        adet_idx = basliga_gore_sutun_bul(df, ['total_product_quantity'])

        eksik = []
        if siparis_idx is None: eksik.append('order_number')
        if platform_idx is None: eksik.append('entegration')
        if durum_idx is None: eksik.append('store_order_status_name')
        if urun_idx is None: eksik.append('product_name')
        if adet_idx is None: eksik.append('total_product_quantity')
        if eksik:
            return None, f"Excel'de şu sütunlar bulunamadı: {', '.join(eksik)}"

    urun_sutunu = df.iloc[:, urun_idx]
    adet_sutunu = df.iloc[:, adet_idx]
    siparis_sutunu = df.iloc[:, siparis_idx]
    platform_sutunu = df.iloc[:, platform_idx]
    durum_sutunu = df.iloc[:, durum_idx]

    siparis_detay = {}

    # Önce sipariş detaylarını oluştur (aynı üründen birden fazla satır varsa birleştir)
    for i in range(len(df)):
        urun = str(urun_sutunu.iloc[i]).strip()
        siparis_no = str(siparis_sutunu.iloc[i]).strip()
        platform = str(platform_sutunu.iloc[i]).strip().lower()
        durum = str(durum_sutunu.iloc[i]).strip().lower()

        if not urun or urun == 'nan' or pd.isna(urun_sutunu.iloc[i]):
            continue

        # Başlık satırını atla
        if urun == 'Ürün İsmi' or 'product' in urun.lower() or urun == 'Ürün ismi':
            continue

        # Sadece Trendyol siparişleri
        if 'trendyol' not in platform:
            continue

        # Pazaryeri durumu filtresi
        if durum_filtre == 'kargoya_verilecek':
            if 'kargoya verilecek' not in durum:
                continue
        elif durum_filtre == 'yeni_siparis':
            if 'yeni' not in durum:
                continue
        elif durum_filtre == 'hepsi':
            if 'kargoya verilecek' not in durum and 'yeni' not in durum:
                continue

        try:
            adet = int(float(adet_sutunu.iloc[i]))
        except (ValueError, TypeError):
            adet = 1

        # Sipariş detayı (aynı ürünleri birleştir)
        if siparis_no and siparis_no != 'nan':
            if siparis_no not in siparis_detay:
                siparis_detay[siparis_no] = []
            # Aynı ürün var mı kontrol et, varsa adetini artır
            urun_bulundu = False
            for item in siparis_detay[siparis_no]:
                if item['urun'] == urun:
                    item['adet'] += adet
                    urun_bulundu = True
                    break
            if not urun_bulundu:
                siparis_detay[siparis_no].append({'urun': urun, 'adet': adet})

    # Sipariş detaylarından ürün özetini oluştur (birleştirilmiş adetlerle)
    urun_ozeti = {}
    for siparis_no, urunler in siparis_detay.items():
        for u in urunler:
            urun = u['urun']
            adet = u['adet']
            if urun not in urun_ozeti:
                urun_ozeti[urun] = {'toplam_adet': 0, 'siparis_sayisi': 0, 'paketler': {}}
            urun_ozeti[urun]['toplam_adet'] += adet
            urun_ozeti[urun]['siparis_sayisi'] += 1
            if adet in urun_ozeti[urun]['paketler']:
                urun_ozeti[urun]['paketler'][adet] += 1
            else:
                urun_ozeti[urun]['paketler'][adet] = 1

    # Karma siparişleri bul
    karma_siparisler_raw = []
    karma_urun_adetleri = {}

    for siparis_no, urunler in siparis_detay.items():
        if len(urunler) > 1:
            karma_siparisler_raw.append({'siparis_no': siparis_no, 'urunler': urunler})
            for u in urunler:
                urun_adi = u['urun']
                a = u['adet']
                if urun_adi not in karma_urun_adetleri:
                    karma_urun_adetleri[urun_adi] = {'toplam_adet': 0, 'siparis_sayisi': 0, 'paketler': {}}
                karma_urun_adetleri[urun_adi]['toplam_adet'] += a
                karma_urun_adetleri[urun_adi]['siparis_sayisi'] += 1
                if a in karma_urun_adetleri[urun_adi]['paketler']:
                    karma_urun_adetleri[urun_adi]['paketler'][a] += 1
                else:
                    karma_urun_adetleri[urun_adi]['paketler'][a] = 1

    # Aynı içerikli karma siparişleri grupla
    karma_gruplar = {}
    for siparis in karma_siparisler_raw:
        icerik_key = tuple(sorted((u['urun'], u['adet']) for u in siparis['urunler']))
        if icerik_key not in karma_gruplar:
            karma_gruplar[icerik_key] = {
                'urunler': siparis['urunler'],
                'siparis_nolar': [],
                'adet': 0
            }
        karma_gruplar[icerik_key]['siparis_nolar'].append(siparis['siparis_no'])
        karma_gruplar[icerik_key]['adet'] += 1

    karma_siparisler = []
    for icerik_key, grup in karma_gruplar.items():
        karma_siparisler.append({
            'urunler': grup['urunler'],
            'siparis_nolar': grup['siparis_nolar'],
            'adet': grup['adet']
        })

    # Karma adetleri ana özetten çıkar
    for urun_adi, karma_bilgi in karma_urun_adetleri.items():
        if urun_adi in urun_ozeti:
            urun_ozeti[urun_adi]['toplam_adet'] -= karma_bilgi['toplam_adet']
            urun_ozeti[urun_adi]['siparis_sayisi'] -= karma_bilgi['siparis_sayisi']
            for a, sayi in karma_bilgi['paketler'].items():
                if a in urun_ozeti[urun_adi]['paketler']:
                    urun_ozeti[urun_adi]['paketler'][a] -= sayi
                    if urun_ozeti[urun_adi]['paketler'][a] <= 0:
                        del urun_ozeti[urun_adi]['paketler'][a]

    # Sonuçları oluştur
    sonuclar = []
    toplam_siparis = 0
    toplam_urun = 0

    for urun in sorted(urun_ozeti.keys()):
        bilgi = urun_ozeti[urun]
        if bilgi['toplam_adet'] <= 0:
            continue
        toplam_siparis += bilgi['siparis_sayisi']
        toplam_urun += bilgi['toplam_adet']
        paket_listesi = []
        for a in sorted(bilgi['paketler'].keys()):
            if bilgi['paketler'][a] > 0:
                paket_listesi.append({'adet': a, 'sayi': bilgi['paketler'][a]})
        sonuclar.append({
            'urun': urun,
            'toplam': bilgi['toplam_adet'],
            'siparis_sayisi': bilgi['siparis_sayisi'],
            'paketler': paket_listesi
        })

    karma_toplam_urun = sum(
        sum(u['adet'] for u in siparis['urunler'])
        for siparis in karma_siparisler
    )

    ozet = {
        'urun_cesidi': len([u for u in urun_ozeti.keys() if urun_ozeti[u]['toplam_adet'] > 0]),
        'toplam_siparis': toplam_siparis + len(karma_siparisler),
        'toplam_urun': toplam_urun + karma_toplam_urun,
        'karma_siparis_sayisi': len(karma_siparisler)
    }

    return {'urunler': sonuclar, 'ozet': ozet, 'karma_siparisler': karma_siparisler}, None


@app.route('/entegra-analiz', methods=['POST'])
def entegra_analiz():
    """İndirilen Entegra Excel dosyasını analiz eder"""
    if not entegra_durum.get('dosya'):
        return jsonify({'error': 'Henüz indirilmiş dosya yok!'})

    dosya_yolu = entegra_durum['dosya']
    if not os.path.exists(dosya_yolu):
        return jsonify({'error': 'Dosya bulunamadı!'})

    durum_filtre = 'kargoya_verilecek'
    if request.is_json and request.json:
        durum_filtre = request.json.get('durum_filtre', 'kargoya_verilecek')

    try:
        df = pd.read_excel(dosya_yolu, header=None)
        sonuc, hata = entegra_analiz_yap(df, durum_filtre)

        if hata:
            return jsonify({'error': hata})

        return jsonify(sonuc)
    except Exception as e:
        return jsonify({'error': f'Hata: {str(e)}'})


@app.route('/genel-entegra-cek', methods=['POST'])
def genel_entegra_cek_route():
    """Genel siparişler için Entegra'dan Selenium ile Excel çeker (filtresiz)"""
    global genel_entegra_durum

    if genel_entegra_durum['durum'] == 'calisiyor':
        return jsonify({'error': 'Zaten bir indirme işlemi devam ediyor!'})

    email = os.getenv('ENTEGRA_EMAIL', '')
    sifre = os.getenv('ENTEGRA_SIFRE', '')

    if not email or not sifre:
        return jsonify({'error': 'E-posta ve şifre .env dosyasında tanımlı değil!'})

    genel_entegra_durum = {'durum': 'calisiyor', 'mesaj': 'Entegra paneline bağlanılıyor...', 'dosya': None, 'detay': None}

    def cek_thread():
        global genel_entegra_durum
        try:
            # Tarih filtresi ile çek (ama trendyol/kargoya verilecek filtresi yok)
            sonuc = excel_cek(email, sifre, headless=False, tarih_filtresi=True)
            if sonuc['basarili']:
                genel_entegra_durum = {
                    'durum': 'tamamlandi',
                    'mesaj': sonuc['mesaj'],
                    'dosya': sonuc['dosya'],
                    'detay': None
                }
            else:
                genel_entegra_durum = {
                    'durum': 'hata',
                    'mesaj': sonuc['mesaj'],
                    'dosya': None,
                    'detay': sonuc.get('sayfa_bilgisi')
                }
        except Exception as e:
            genel_entegra_durum = {
                'durum': 'hata',
                'mesaj': f'Beklenmeyen hata: {str(e)}',
                'dosya': None,
                'detay': None
            }

    thread = threading.Thread(target=cek_thread)
    thread.start()

    return jsonify({'mesaj': 'İndirme başlatıldı, tarayıcı açılacak...'})


@app.route('/genel-entegra-durum')
def genel_entegra_durum_route():
    """Genel siparişler için Entegra indirme durumunu döndürür"""
    sonuc = dict(genel_entegra_durum)
    if sonuc['durum'] == 'calisiyor' and entegra_cek.durum_mesaj:
        sonuc['mesaj'] = entegra_cek.durum_mesaj
    return jsonify(sonuc)


def genel_entegra_analiz_yap(df):
    """Genel siparişler için Entegra Excel analizi - sütun isimlerine göre dinamik tespit yapar

    Desteklenen formatlar:
    - Ayrıntılı Excel (Türkçe): Platform Referans No, Kargo Kodu, Ürün İsmi, Adet, Barkod
    - Normal Excel (İngilizce): order_number, cargo_code, product_name, total_product_quantity, barcode
    """
    # İlk satıra bakarak formatı algıla
    ilk_hucre = str(df.iloc[0, 0]).strip()

    if ilk_hucre == 'ID' or 'Tarih' in str(df.iloc[0, 1]):
        # Format 1 - Ayrıntılı Excel (Türkçe) - başlık ismine göre bul
        siparis_idx = basliga_gore_sutun_bul(df, ['Platform Referans No'])
        kargo_idx = basliga_gore_sutun_bul(df, ['Kargo Kodu'])
        urun_idx = basliga_gore_sutun_bul(df, ['Ürün İsmi', 'Ürün ismi'])
        adet_idx = basliga_gore_sutun_bul(df, ['Adet'])
        barkod_idx = basliga_gore_sutun_bul(df, ['Barkod'])

        eksik = []
        if siparis_idx is None: eksik.append('Platform Referans No')
        if kargo_idx is None: eksik.append('Kargo Kodu')
        if urun_idx is None: eksik.append('Ürün İsmi')
        if adet_idx is None: eksik.append('Adet')
        if barkod_idx is None: eksik.append('Barkod')
        if eksik:
            return None, f"Excel'de şu sütunlar bulunamadı: {', '.join(eksik)}"
    else:
        # Format 2 - Normal Excel (İngilizce) - başlık ismine göre bul
        siparis_idx = basliga_gore_sutun_bul(df, ['order_number'])
        kargo_idx = basliga_gore_sutun_bul(df, ['cargo_code'])
        urun_idx = basliga_gore_sutun_bul(df, ['product_name'])
        adet_idx = basliga_gore_sutun_bul(df, ['total_product_quantity'])
        barkod_idx = basliga_gore_sutun_bul(df, ['barcode'])

        eksik = []
        if siparis_idx is None: eksik.append('order_number')
        if kargo_idx is None: eksik.append('cargo_code')
        if urun_idx is None: eksik.append('product_name')
        if adet_idx is None: eksik.append('total_product_quantity')
        if barkod_idx is None: eksik.append('barcode')
        if eksik:
            return None, f"Excel'de şu sütunlar bulunamadı: {', '.join(eksik)}"

    siparis_sutunu = df.iloc[:, siparis_idx]
    kargo_sutunu = df.iloc[:, kargo_idx]
    urun_sutunu = df.iloc[:, urun_idx]
    adet_sutunu = df.iloc[:, adet_idx]
    barkod_sutunu = df.iloc[:, barkod_idx]

    # Kargo Kodu -> ürün bilgisi eşleştirmesi (barkod okuyucu ile kargo kodu okutulur)
    barkodlar = {}

    for i in range(len(df)):
        siparis_no = str(siparis_sutunu.iloc[i]).strip()
        urun = str(urun_sutunu.iloc[i]).strip()
        kargo_kodu = str(kargo_sutunu.iloc[i]).strip()
        barkod = str(barkod_sutunu.iloc[i]).strip()

        if kargo_kodu == '' or kargo_kodu == 'nan' or pd.isna(kargo_sutunu.iloc[i]):
            continue

        # Başlık satırını atla
        if kargo_kodu in ['cargo_code', 'Kargo Kodu']:
            continue

        if not urun or urun == 'nan' or pd.isna(urun_sutunu.iloc[i]):
            continue

        # Adet bilgisini al
        try:
            adet = int(float(adet_sutunu.iloc[i]))
        except (ValueError, TypeError):
            adet = 1

        if kargo_kodu not in barkodlar:
            barkodlar[kargo_kodu] = []

        barkodlar[kargo_kodu].append({
            'urun': urun,
            'adet': adet,
            'barkod': barkod,
            'siparis_no': siparis_no
        })

    return {'barkodlar': barkodlar}, None


@app.route('/genel-entegra-analiz', methods=['POST'])
def genel_entegra_analiz():
    """İndirilen Entegra Excel dosyasını genel siparişler için analiz eder"""
    if not genel_entegra_durum.get('dosya'):
        return jsonify({'error': 'Henüz indirilmiş dosya yok!'})

    dosya_yolu = genel_entegra_durum['dosya']
    if not os.path.exists(dosya_yolu):
        return jsonify({'error': 'Dosya bulunamadı!'})

    try:
        df = pd.read_excel(dosya_yolu, header=None)
        sonuc, hata = genel_entegra_analiz_yap(df)

        if hata:
            return jsonify({'error': hata})

        return jsonify(sonuc)
    except Exception as e:
        return jsonify({'error': f'Hata: {str(e)}'})


if __name__ == '__main__':
    app.run(debug=True, port=5000)
