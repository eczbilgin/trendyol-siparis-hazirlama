# Trendyol Siparis Hazirlama

Trendyol siparis Excel dosyalarindan siparis ozeti olusturan web uygulamasi.

## Ozellikler

- Excel dosyasi (.xlsx) yukleme
- Urun bazinda siparis ozeti
- Paket hazirlama listesi (kac adetlik siparisten kac tane var)
- Yazdirilabilir kompakt tablo formati

## Kurulum

```bash
pip install flask pandas openpyxl
```

## Kullanim

1. `BASLAT.bat` dosyasina cift tiklayin
2. Chrome'da acilan sayfada "Dosya Sec" butonuna tiklayin
3. Trendyol Excel dosyanizi secin (.xlsx)
4. "ANALIZ ET" butonuna basin
5. Sonuclari gorun ve yazdir butonuyla cikti alin

## Excel Dosya Formati

Program asagidaki sutunlari okur:
- **BN sutunu (66. sutun):** Urun adi
- **BS sutunu (71. sutun):** Siparis adedi
- **C sutunu (3. sutun):** Siparis numarasi

## Ekran Goruntusu

Uygulama Chrome'da calisir ve kolay kullanim icin gorsel arayuz sunar.

## Lisans

MIT
