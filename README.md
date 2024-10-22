# Aplikace pro Adama

## Prerekvizity

- Python
- (volitelný) git

## Jak spustit aplikaci

1. stáhnout repozitář

Buďto stáhnout zip soubor a rozbalit ho, nebo pomocí gitu
```
git clone https://github.com/zdenecek/xlsx-merger-adamko.git
```

1. vytvořit si virtuální prostředí

```
python -m venv venv
```

1. aktivovat virtuální prostředí


pro Windows
```
.\venv\Scripts\Activate.bat
```

1. nainstalovat potřebné knihovny

```
pip install -r requirements.txt
```

1. spustit aplikaci

```
streamlit run src/app.py
```