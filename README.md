# Axın - Excel və CSV Məlumat Analitikası

Axın, Excel və CSV fayllarını brauzerdə yükləmək, təhlil etmək, təmizləmək,
birləşdirmək, doğrulamaq və eksport etmək üçün yerli FastAPI tətbiqidir.

## Xüsusiyyətlər

- `.xlsx`, `.xls` və `.csv` fayllarının yüklənməsi
- Bir neçə vərəq və bir neçə fayl arasında məlumat seçimi
- Təkrar sətirlərin silinməsi
- Boş xanaların sıfır və ya sütun ortalaması ilə doldurulması
- Boş sətirlərin silinməsi
- Azərbaycan mətnlərinin standartlaşdırılması
- Tarixlərin `YYYY-MM-DD` formatına çevrilməsi
- Mənfi dəyərlərin və IQR kənar dəyərlərinin aşkarlanması
- Cədvəllərin birləşdirilməsi
- XLSX, SQL və mətn hesabatı kimi eksport
- Excel şrift rənglərinə görə XLSX sətirlərinin ayrılması

## Tələblər

- Python 3.10 və ya daha yeni versiya
- pip

## İşə salma

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --host 127.0.0.1 --port 8000
```

Brauzerdə `http://127.0.0.1:8000` ünvanını açın.

## Testlər

```bash
python -m unittest discover -s tests -v
```

## Layihə strukturu

- `main.py` - FastAPI backend və məlumat emalı funksiyaları
- `static/index.html` - Web interfeysi
- `tests/test_color_extraction.py` - Rəng ayrılması testləri
- `test_data_full.xlsx` - İstifadə və yoxlama üçün nümunə fayl
