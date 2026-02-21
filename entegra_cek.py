# -*- coding: utf-8 -*-
import os
import time
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

INDIRME_KLASORU = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'indirilenler')

# Durum takibi
durum_mesaj = ''

def durum_guncelle(mesaj):
    global durum_mesaj
    durum_mesaj = mesaj


def tarayici_baslat():
    """Chrome tarayıcısını Selenium ile başlatır"""
    os.makedirs(INDIRME_KLASORU, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument('--window-position=0,0')

    prefs = {
        'download.default_directory': INDIRME_KLASORU,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
    }
    chrome_options.add_experimental_option('prefs', prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)

    # Chrome penceresini öne getir
    one_getir(driver)

    return driver


def one_getir(driver):
    """Chrome penceresini öne getirir"""
    try:
        driver.switch_to.window(driver.current_window_handle)
        driver.execute_script("window.focus();")
    except:
        pass


def giris_yap(driver, email, sifre):
    """Adım 1: Entegra paneline giriş yapar"""
    durum_guncelle('Entegra giris sayfasi aciliyor...')
    driver.get('https://web.entegrabilisim.com/')
    wait = WebDriverWait(driver, 15)

    durum_guncelle('Kullanici bilgileri dolduruluyor...')
    email_input = wait.until(EC.presence_of_element_located((By.ID, 'input-username')))
    email_input.clear()
    email_input.send_keys(email)

    sifre_input = driver.find_element(By.ID, 'input-password')
    sifre_input.clear()
    sifre_input.send_keys(sifre)

    durum_guncelle('CAPTCHA bekleniyor - Chrome penceresinde "Ben robot degilim" tiklayin ve "Oturum Ac" butonuna basin!')

    # Chrome'u öne getir (kullanıcı görsün)
    one_getir(driver)

    # Kullanıcının CAPTCHA çözüp giriş yapmasını bekle (max 120 saniye)
    for i in range(60):
        time.sleep(2)
        try:
            # do_login butonu kaybolursa giriş yapılmış demektir
            login_btn = driver.find_elements(By.ID, 'do_login')
            if not login_btn:
                durum_guncelle('Giris basarili!')
                return True, "Giriş başarılı!"
            # Veya dashboard elementleri göründüyse
            body_text = driver.find_element(By.TAG_NAME, 'body').text
            if 'Raporlar' in body_text or 'Ana Men' in body_text:
                durum_guncelle('Giris basarili!')
                return True, "Giriş başarılı!"
        except:
            pass

    return False, "Giriş zaman aşımına uğradı. 2 dakika içinde giriş yapılmadı."


def siparislere_git(driver):
    """Adım 2: Dashboard'dan Siparişler ikonuna tıklar"""
    durum_guncelle('Siparisler ikonuna tiklaniyor...')
    time.sleep(2)

    # URL'den token'ı al
    url = driver.current_url
    token = ''
    if 'token=' in url:
        token = url.split('token=')[1].split('&')[0]

    # Doğrudan siparişler sayfasına git (URL'den biliyoruz)
    siparis_url = f'https://web.entegrabilisim.com/index.php?route=order/order&token={token}'
    driver.get(siparis_url)
    time.sleep(5)

    # Siparişler Listesi yüklendi mi kontrol et
    body_text = driver.find_element(By.TAG_NAME, 'body').text
    if 'Sipari' in body_text and 'Liste' in body_text:
        durum_guncelle('Siparisler sayfasi acildi.')
        return True, "Siparişler sayfasına gidildi."

    # URL ile gidemediyse, dashboard'daki ikona tıklamayı dene
    durum_guncelle('Siparisler ikonu araniyor...')
    driver.get(f'https://web.entegrabilisim.com/index.php?route=common/dashboard&token={token}')
    time.sleep(3)

    # "Siparişler" metnini içeren tıklanabilir element bul
    try:
        # Tüm tıklanabilir elementlerden "Siparişler" içereni bul
        elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Sipari')]")
        for el in elements:
            try:
                if el.is_displayed():
                    el.click()
                    time.sleep(5)
                    body = driver.find_element(By.TAG_NAME, 'body').text
                    if 'Liste' in body:
                        durum_guncelle('Siparisler sayfasi acildi.')
                        return True, "Siparişler sayfasına gidildi."
            except:
                continue
    except:
        pass

    return False, "Siparişler sayfası bulunamadı."


def excel_indir(driver):
    """Adım 3-4: Toplu İşlemler → Ayrıntılı Excel → İndir"""
    durum_guncelle('Toplu Islemler butonuna tiklaniyor...')
    wait = WebDriverWait(driver, 15)

    # Mevcut dosyaları kaydet (yeni indirilen dosyayı ayırt etmek için)
    onceki_dosyalar = set(glob.glob(os.path.join(INDIRME_KLASORU, '*.xls*')))

    # "Toplu İşlemler" butonuna tıkla
    toplu_islem_bulundu = False
    selectors = [
        "//button[contains(text(), 'Toplu')]",
        "//a[contains(text(), 'Toplu')]",
        "//*[contains(text(), 'Toplu')]",
        "//button[contains(@class, 'toplu')]",
    ]
    for selector in selectors:
        try:
            elements = driver.find_elements(By.XPATH, selector)
            for el in elements:
                if el.is_displayed() and 'Toplu' in el.text:
                    el.click()
                    toplu_islem_bulundu = True
                    time.sleep(3)
                    break
            if toplu_islem_bulundu:
                break
        except:
            continue

    if not toplu_islem_bulundu:
        return False, "'Toplu İşlemler' butonu bulunamadı."

    # "Ayrıntılı Excel" butonuna tıkla
    durum_guncelle('Ayrintili Excel butonuna tiklaniyor...')
    excel_bulundu = False
    excel_selectors = [
        "//button[contains(text(), 'Excel')]",
        "//a[contains(text(), 'Excel')]",
        "//*[contains(text(), 'Ayr') and contains(text(), 'Excel')]",
        "//button[contains(text(), 'Ayr')]",
        "//a[contains(text(), 'Ayr')]",
    ]
    for selector in excel_selectors:
        try:
            elements = driver.find_elements(By.XPATH, selector)
            for el in elements:
                if el.is_displayed() and 'Excel' in el.text:
                    el.click()
                    excel_bulundu = True
                    time.sleep(3)
                    break
            if excel_bulundu:
                break
        except:
            continue

    if not excel_bulundu:
        return False, "'Ayrıntılı Excel' butonu bulunamadı."

    # Dosyanın indirilmesini bekle (max 60 saniye)
    durum_guncelle('Excel dosyasi indiriliyor...')
    for i in range(30):
        time.sleep(2)
        mevcut_dosyalar = set(glob.glob(os.path.join(INDIRME_KLASORU, '*.xls*')))
        yeni_dosyalar = mevcut_dosyalar - onceki_dosyalar

        # .crdownload (indirme devam ediyor) dosyalarını filtrele
        tamamlanan = [f for f in yeni_dosyalar if not f.endswith('.crdownload')]
        if tamamlanan:
            dosya = max(tamamlanan, key=os.path.getmtime)
            durum_guncelle('Excel basariyla indirildi!')
            return True, dosya

    # Son çare: en son değişen dosyayı döndür
    dosya = son_indirilen_dosya()
    if dosya and dosya not in onceki_dosyalar:
        return True, dosya

    return False, "Excel dosyası indirilemedi. İndirme zaman aşımına uğradı."


def son_indirilen_dosya():
    """İndirme klasöründeki en son .xlsx/.xls dosyasını bulur"""
    xlsx_dosyalar = glob.glob(os.path.join(INDIRME_KLASORU, '*.xlsx'))
    xls_dosyalar = glob.glob(os.path.join(INDIRME_KLASORU, '*.xls'))
    dosyalar = [f for f in xlsx_dosyalar + xls_dosyalar if not f.endswith('.crdownload')]

    if not dosyalar:
        return None

    return max(dosyalar, key=os.path.getmtime)


def sayfa_bilgisi_al(driver):
    """Mevcut sayfanın bilgilerini döndürür (debug için)"""
    bilgi = {
        'url': driver.current_url,
        'title': driver.title,
        'linkler': [],
        'butonlar': []
    }

    try:
        linkler = driver.find_elements(By.TAG_NAME, 'a')
        for link in linkler[:30]:
            text = link.text.strip()
            href = link.get_attribute('href') or ''
            if text:
                bilgi['linkler'].append({'text': text, 'href': href})
    except:
        pass

    try:
        butonlar = driver.find_elements(By.TAG_NAME, 'button')
        for btn in butonlar[:20]:
            text = btn.text.strip()
            if text:
                bilgi['butonlar'].append(text)
    except:
        pass

    return bilgi


def excel_cek(email, sifre, headless=False):
    """Ana fonksiyon: Entegra'ya giriş yap → Siparişler → Toplu İşlemler → Ayrıntılı Excel"""
    driver = None
    try:
        durum_guncelle('Chrome tarayici baslatiliyor...')
        driver = tarayici_baslat()

        # Adım 1: Giriş yap
        basarili, mesaj = giris_yap(driver, email, sifre)
        if not basarili:
            sayfa = sayfa_bilgisi_al(driver)
            return {'basarili': False, 'mesaj': mesaj, 'sayfa_bilgisi': sayfa}

        # Adım 2: Siparişler sayfasına git
        basarili, mesaj = siparislere_git(driver)
        if not basarili:
            sayfa = sayfa_bilgisi_al(driver)
            return {'basarili': False, 'mesaj': mesaj, 'sayfa_bilgisi': sayfa, 'adim': 'siparisler'}

        # Adım 3-4: Toplu İşlemler → Ayrıntılı Excel → İndir
        basarili, sonuc = excel_indir(driver)
        if not basarili:
            sayfa = sayfa_bilgisi_al(driver)
            return {'basarili': False, 'mesaj': sonuc, 'sayfa_bilgisi': sayfa, 'adim': 'excel_indir'}

        durum_guncelle('Excel basariyla indirildi!')
        return {
            'basarili': True,
            'mesaj': 'Excel başarıyla indirildi!',
            'dosya': sonuc
        }

    except Exception as e:
        return {'basarili': False, 'mesaj': f'Hata: {str(e)}'}

    finally:
        if driver:
            driver.quit()
