# DTC Generatorius

Automobilių diagnostikos kodų (DTC) paieška naršyklėje, su vietiniu lietuvišku techniniu vertimu.  
Web lookup for automotive diagnostic trouble codes, with a local Lithuanian technical translation.

Veikianti svetainė / live site: https://dtc-generatorius-clom4.ondigitalocean.app

---

## Lietuviškai

### Apie projektą

DTC Generatorius ieško OBD-II stiliaus gedimų kodų (`P`, `B`, `C` arba `U` ir keturi skaitmenys). Įvedus kodą, pavyzdžiui `P0101`, puslapis parodo:

- originalų anglišką aprašymą iš šaltinio Word failo
- lietuvišką techninį vertimą, sudarytą vietoje pagal frazių žodyną (mokamo vertimo API nereikia)
- ar kodas patenka į standartinį OBD-II diapazoną, gamintojo valdomą diapazoną, ar į projekto Ford specifinių kodų sąrašą
- tris orientacines galimas priežastis

Gedimų sąrašas yra tik orientyras. Tai nėra remonto procedūra ir nėra galutinė diagnozė. Originali formuluotė lieka šalia vertimo, nes gamintojo specifinės frazės gali būti išverstos netiksliai.

Kodai įkeliami paleidžiant programą iš `DTC generatoriaus saltinis(2).docx` projekto šaknyje (dabartiniame faile jų yra 3717). Patvirtintą Ford formuluotę konkrečiam modeliui, metams ir varikliui galima įrašyti į `FORD_SPECIFIC_CODES` faile `engine.py`.

Palaikoma sąsaja yra ši žiniatinklio programa. Ankstesnis Jupyter notebook (`DTC Kodu generatorius V7.ipynb`) iš projekto pašalintas. Jei prireiks, jis lieka Git istorijoje:

```bash
git log -- "DTC Kodu generatorius V7.ipynb"
git checkout <commit> -- "DTC Kodu generatorius V7.ipynb"
```

### Kur veikia

DigitalOcean App Platform (Frankfurtas) diegia GitHub saugyklą `WildWolfLT/DTC-Generatorius`, šaką `main`. Įkelimas į `main` iš naujo paleidžia svetainę.

Procesas yra Gunicorn, prievadas 8080:

```text
gunicorn --worker-tmp-dir /dev/shm --bind 0.0.0.0:8080 --workers 2 --preload app:app
```

CSS ir JavaScript pateikiami iš `/assets`. Platformos kraštas į `/static` atsako HTTP 530, todėl Flask `static_url_path` lieka `/assets`.

### Paleidimas savo kompiuteryje

Python 3.14 (žr. `runtime.txt`). Word šaltinis turi gulėti šalia `app.py`.

Windows (`cmd`):

```bat
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
python app.py
```

macOS arba Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Atidarykite http://127.0.0.1:8080 . Numatytasis prievadas yra `8080`; jį keičia aplinkos kintamasis `PORT`.

Patikrinti paieškos taisykles nepaleidžiant serverio:

```bat
python test_engine.py
```

### HTTP API

| Metodas ir kelias | Paskirtis |
| --- | --- |
| `GET /` | Paieškos puslapis |
| `GET /api/meta` | Kodų skaičius ir ar Word failas įsikėlė |
| `GET /api/lookup?code=P0101` | Pilnas vieno kodo rezultatas |
| `GET /api/suggest?q=P01&limit=8` | Pasiūlymai pagal pradžią (riba 1–12) |

---

## English

### What this is

DTC Generatorius looks up OBD-II style trouble codes (`P`, `B`, `C`, or `U` plus four digits). Type a code such as `P0101` and the page shows:

- the original English description from the source Word file
- a Lithuanian technical translation, built locally from a phrase dictionary (no paid translation API)
- whether the code falls in the standard OBD-II range, a manufacturer-controlled range, or the project's Ford-specific list
- three indicative possible causes

The fault list is orientation only. It is not a repair procedure and not a final diagnosis. The original wording stays on screen next to the translation, because manufacturer-specific phrases can be translated imperfectly.

Codes are loaded at startup from `DTC generatoriaus saltinis(2).docx` in the project root (3717 codes in the current file). Confirmed Ford wording for a specific model, year, and engine can be added to `FORD_SPECIFIC_CODES` in `engine.py`.

The supported interface is this web app. The earlier Jupyter notebook (`DTC Kodu generatorius V7.ipynb`) has been removed from the tree. It remains in Git history if you need to restore it:

```bash
git log -- "DTC Kodu generatorius V7.ipynb"
git checkout <commit> -- "DTC Kodu generatorius V7.ipynb"
```

### How it is hosted

DigitalOcean App Platform (Frankfurt) deploys the GitHub repository `WildWolfLT/DTC-Generatorius`, branch `main`. A push to `main` redeploys the site.

The process is Gunicorn on port 8080:

```text
gunicorn --worker-tmp-dir /dev/shm --bind 0.0.0.0:8080 --workers 2 --preload app:app
```

CSS and JavaScript are served from `/assets`. The platform edge returns HTTP 530 for `/static`, so Flask `static_url_path` stays `/assets`.

### Run it yourself

Python 3.14 (see `runtime.txt`). The Word source file must sit next to `app.py`.

Windows (`cmd`):

```bat
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
python app.py
```

macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:8080 . The port defaults to `8080`; set the `PORT` environment variable to change it.

Check the lookup rules without starting the server:

```bat
python test_engine.py
```

### HTTP API

| Method and path | Purpose |
| --- | --- |
| `GET /` | Search page |
| `GET /api/meta` | Code count and whether the Word file loaded |
| `GET /api/lookup?code=P0101` | Full result for one code |
| `GET /api/suggest?q=P01&limit=8` | Prefix suggestions (limit 1–12) |
