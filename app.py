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
    karma_urun_adetleri = {}  # Karma siparişlerdeki ürün adetlerini takip et

    for siparis_no, urunler in siparis_detay.items():
        if len(urunler) > 1:
            karma_siparisler.append({
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

        if df.shape[1] <= max(an_idx, bn_idx, bs_idx):
            return jsonify({'error': 'Excel dosyasında yeterli sütun yok!'})

        barkod_sutunu = df.iloc[:, an_idx]
        urun_sutunu = df.iloc[:, bn_idx]
        adet_sutunu = df.iloc[:, bs_idx]

        barkodlar = {}

        for i in range(len(df)):
            barkod = str(barkod_sutunu.iloc[i]).strip()
            urun = str(urun_sutunu.iloc[i]).strip()

            if barkod == '' or barkod == 'nan' or pd.isna(barkod_sutunu.iloc[i]):
                continue

            # Başlık satırını atla
            if 'Barkod' in barkod or 'barkod' in barkod:
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

if __name__ == '__main__':
    app.run(debug=True, port=5000)
