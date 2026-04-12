# Fikstür Talep Uygulaması

Elektrooptik sistem fikstür taleplerini yönetmek için geliştirilmiş masaüstü uygulaması.

## Özellikler

- Talep formu doldurma (Yeni Talep / Revizyon)
- Outlook üzerinden otomatik mail gönderimi (Windows)
- Excel log dosyasına kayıt (aciliyet bazlı renklendirme)
- PDF özet oluşturma
- Talep klasörü ve dosya yönetimi

## Kurulum

```bash
pip install -r requirements.txt
```

> Not: `pywin32` yalnızca Windows'ta gereklidir (Outlook entegrasyonu).

## Çalıştırma

```bash
python fikstur_talep2.py
```

İlk çalıştırmadan önce `fikstur_talep2.py` içindeki `DEFAULT_BASE_DIR` değişkenini
kayıt dizinine göre ayarlayın (varsayılan: `~/Documents/FixtureRequests`).

## EXE Oluşturma (Windows)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "FiksturTalep" fikstur_talep2.py
```

Oluşan dosya: `dist\FiksturTalep.exe`

Detaylı kurulum ve dağıtım talimatı için `KURULUM_VE_DAGITIM_TALIMATI.txt` dosyasına bakın.

## Sürüm

v2.0
