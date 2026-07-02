# TDR Generator

ALPORT Terminal Departure Report (TDR) oluşturmak için Streamlit tabanlı basit uygulama.

## GitHub dosya yapısı

Klasör kullanmadan sadece şu 4 dosyayı yükleyin:

```text
app.py
requirements.txt
README.md
tdr_template.xlsx
```

## Kurulum

```bash
pip install -r requirements.txt
```

## Çalıştırma

```bash
streamlit run app.py
```

## Kullanım

1. Vessel Name, Voyage No ve Flag bilgilerini girin.
2. Varsa Discharging, Loading, Reefer, IMDG ve Damage List Excel listelerini yükleyin.
3. **TDR Excel Oluştur** butonuna basın.
4. Oluşan Excel dosyasını indirin.

Not: `tdr_template.xlsx` dosyası `app.py` ile aynı yerde olmalıdır.
