# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, jsonify
import pandas as pd
import os

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

def sutun_indeksi(sutun_adi):
    """Excel sütun adını indekse çevirir"""
    indeks = 0
    for i, harf in enumerate(reversed(sutun_adi.upper())):
        indeks += (ord(harf) - ord('A') + 1) * (26 ** i)
    return indeks - 1

def analiz_yap(df):
    """Excel verisini analiz eder"""
    bn_idx = sutun_indeksi('BN')  # Ürün adı
    bs_idx = sutun_indeksi('BS')  # Sipariş adedi
    c_idx = sutun_indeksi('C')    # Sipariş numarası

    if df.shape[1] <= max(bn_idx, bs_idx, c_idx):
        return None, "Excel dosyasında yeterli sütun yok!"

    urun_sutunu = df.iloc[:, bn_idx]
    adet_sutunu = df.iloc[:, bs_idx]
    siparis_sutunu = df.iloc[:, c_idx]

    urun_ozeti = {}
    siparis_detay = {}  # Sipariş numarasına göre ürünleri grupla

    for i in range(len(df)):
        urun = str(urun_sutunu.iloc[i]).strip()
        siparis_no = str(siparis_sutunu.iloc[i]).strip()

        if urun == '' or urun == 'nan' or pd.isna(urun_sutunu.iloc[i]):
            continue

        # Başlık satırını atla
        if urun == 'Ürün İsmi':
            continue

        try:
            adet = int(float(adet_sutunu.iloc[i]))
        except (ValueError, TypeError):
            continue

        # Ürün özeti
        if urun not in urun_ozeti:
            urun_ozeti[urun] = {
                'toplam_adet': 0,
                'siparis_sayisi': 0,
                'paketler': {}
            }

        urun_ozeti[urun]['toplam_adet'] += adet
        urun_ozeti[urun]['siparis_sayisi'] += 1

        # Paket gruplama
        if adet in urun_ozeti[urun]['paketler']:
            urun_ozeti[urun]['paketler'][adet] += 1
        else:
            urun_ozeti[urun]['paketler'][adet] = 1

        # Sipariş detayı (karma siparişler için)
        if siparis_no and siparis_no != 'nan':
            if siparis_no not in siparis_detay:
                siparis_detay[siparis_no] = []
            siparis_detay[siparis_no].append({
                'urun': urun,
                'adet': adet
            })

    # Karma siparişleri bul (birden fazla ürün içeren)
    karma_siparisler = []
    for siparis_no, urunler in siparis_detay.items():
        if len(urunler) > 1:
            karma_siparisler.append({
                'siparis_no': siparis_no,
                'urunler': urunler
            })

    # Sonuçları liste olarak döndür
    sonuclar = []
    toplam_siparis = 0
    toplam_urun = 0

    for urun in sorted(urun_ozeti.keys()):
        bilgi = urun_ozeti[urun]
        toplam_siparis += bilgi['siparis_sayisi']
        toplam_urun += bilgi['toplam_adet']

        paket_listesi = []
        for adet in sorted(bilgi['paketler'].keys()):
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

    ozet = {
        'urun_cesidi': len(urun_ozeti),
        'toplam_siparis': toplam_siparis,
        'toplam_urun': toplam_urun,
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

if __name__ == '__main__':
    app.run(debug=True, port=5000)
