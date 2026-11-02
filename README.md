# MÉH Cégelemző

OKIR hulladékadatok és céges forrásadatok elemzése, közös Python-szolgáltatásokkal:

- **Asztali alkalmazás:** a meglévő PySide6 felület, Excel-import, grafikonok,
  KTJ-adatok, helyi mentések és `.okirpkg` csomagok megmaradnak.
- **Web / iOS PWA:** reszponzív, magyar nyelvű felület Safariban, iPhone-on,
  iPaden és asztali böngészőben. A számítás és az OKIR-lekérés a szerveren fut;
  az iOS-eszközre nem kell Python vagy Qt.

Python **3.11 vagy újabb** szükséges. A verziózott `dist/` korábbi, kész
asztali build: **nem tartalmazza az új funkciókat**. Az új kódot forrásból
indítsd, vagy készíts új platformfüggő buildet.

## Webes indítás

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements-web.txt
# A jelszó nem kerül a parancselőzményekbe:
read -rs -p "Munkaterület jelszava (legalább 12 karakter): " MEH_WEB_PASSWORD
echo
export MEH_WEB_PASSWORD
python web_app.py
```

Windows PowerShell esetén az aktiválás `.venv\Scripts\Activate.ps1`;
a `MEH_WEB_PASSWORD` környezeti változót a helyi titokkezelésből állítsd be.

Alapértelmezett cím: **http://127.0.0.1:8080**. A webes függőségek nem
telepítik a PySide6-ot vagy a matplotlibet. A szerver Waitress-t használ,
nem a Flask fejlesztői szerverét.

### iPhone / iPad és éles HTTPS

1. A Python-szervert egy tartósan elérhető gépen kell futtatni. Az iPhone-on
   a `localhost` az iPhone-t jelenti, nem a szervert.
2. Tegyél elé HTTPS reverse proxyt, például Caddyt. A DNS-név mutasson a
   szerverre; a 80/443 port legyen elérhető. Példa `Caddyfile`:

   ```caddyfile
   cegelemzo.example.hu {
       reverse_proxy 127.0.0.1:8080
   }
   ```

3. HTTPS mögött állítsd be `MEH_HTTPS=1`-re a biztonságos munkamenet-cookie-hoz.
   A backend maradjon loopback címen. A folyamatot szolgáltatáskezelővel
   (például systemd) indítsd, nem interaktív terminálban.
4. Safariban nyisd meg a **saját HTTPS-címet**, jelentkezz be, majd
   **Megosztás → Főképernyőhöz adás**. Az ikonról önálló ablakban használható.

A főképernyős telepítéshez HTTPS kell (fejlesztéskor a localhost kivétel).
Az app nem natív App Store-alkalmazás, és **nem offline elemző**. Hálózat
nélkül külön tájékoztató jelenik meg; cégadatok, API-válaszok, exportok és
hitelesített HTML-oldalak nem kerülnek a service worker gyorsítótárába.
A letöltött fájlokat a felhasználó saját eszköze tárolja.

| Környezeti változó | Jelentés |
|---|---|
| `MEH_WEB_PASSWORD` | Kötelező, legalább 12 karakteres közös munkaterület-jelszó |
| `MEH_SECRET_KEY` | Opcionális, titokként kezelt véletlen kulcs; nélküle újraindításkor mindenki újra belép |
| `MEH_DATA_HOME` | Írható, tartós munkakönyvtár; alapérték az indítás könyvtára |
| `MEH_HOST` / `MEH_PORT` | Alapérték `127.0.0.1` / `8080` |
| `MEH_HTTPS` | `1`: csak HTTPS-en küldött session-cookie |

Ez **egyetlen közös munkaterület**, nem felhasználónként elkülönített SaaS.
A jelszó birtokosai ugyanazokat az adatokat és mentéseket látják és módosítják.
Külön csapatoknak külön példány, jelszó és `MEH_DATA_HOME` szükséges. Egy
példányhoz **egy szerverfolyamatot** használj; a szálak és a háttérfeladatok
azon belül működnek. Ne indíts több worker-folyamatot vagy több konténert
ugyanarra az adatkönyvtárra.

A `MEH_DATA_HOME/data/okir_cache.db` tárolja a mentéseket, metaadatokat,
OKIR-cache-t és az aktív webes mentés azonosítóját. Rendszeres biztonsági
mentéshez SQLite backupot használj, vagy állítsd le a szervert a teljes
adatkönyvtár másolásakor. Futó WAL-adatbázisból a `.db` önmagában nem elég.
Az asztali és webes appot ne futtasd egyszerre ugyanazon adatkönyvtáron.

## Használat

**Import:** `.xlsx` forrás vagy az app által készített MÉH `.json`.
Az Excel kötelező oszlopai: `Teljes név` / `Cégnév`, `Adószám`, `Megye`,
`Irányítószám`, `Település`, `Cím`. A korábbi ékezet nélküli
fejlécváltozatok és az opcionális kapcsolattartási/pénzügyi mezők továbbra
is működnek. Az OKIR-cégkeresés az adószám első nyolc számjegyét használja.

Weben maximum 20 MB-os feltöltés, 100 MB kicsomagolt XLSX és 2000
Excel-forráscég engedélyezett egy importban. JSON-táblánként legfeljebb
250 000 sor tölthető be. A hibás sorok és sikertelen lekérések a
**Kimaradt / hibák** nézetben láthatók és exportálhatók. Részleges importnál
ellenőrizd a hibajegyzéket: a hiányzó adatok nem nulla hulladékot jelentenek.

Az import a háttérben fut; közben a korábbi adatok olvashatók. Egyszerre egy
módosító feladat futhat. Sikeres import után automatikus mentés készül.
Hibás fájl nem cseréli le a korábbi adatokat. Részleges, akár kizárólag
hibákat tartalmazó OKIR-eredmény új munkamenetként mentődik; a korábbi mentés
visszatölthető. Újraindításkor az aktív mentés áll vissza, a félbeszakadt
háttérlekérést újra kell indítani.

**Szűrés:** cég, év, termékcsoport, HAK-kódrészlet (`;`-vel elválasztva);
weben adattípus is választható. Üres kijelölés: összes. Induláskor és új
adat betöltésekor a legutóbbi két adatév aktív. Weben a „Szűrők törlése”
minden évet mutat; az asztali alaphelyzet a korábbi két legfrissebb év.
A grafikonokra kattintva továbbszűrhetsz. A cégadatok nézetére csak a
cégnévszűrő, a hibajegyzékre semmilyen elemzési szűrő nem vonatkozik.

**KTJ:** a nyers telephelyi sorok azonnal megjelennek. A cím/metaadatok
betöltése weben külön gombbal indítható háttérfeladat. A KTJ-export nem
indít rejtett webes hálózati lekérést; előbb töltsd be a metaadatokat.

**Mennyiségek:** az alapegység kg. Egy HAK-kód több termékcsoportba tartozhat;
a termékcsoport-összegek ezért átfednek, de a teljes mennyiségből a
kategóriabővítés duplikátumai kimaradnak. Külön adattípusok és külön
hulladékleírások nem vesznek el a részletes exportban. Az összesítő tonnára
kerekítése az összegzés után történik. Több adattípus összege nem nettó
anyagáram vagy anyagmérleg.

## Exportok és átjárás a két felület között

| Export | Tartalom |
|---|---|
| Excel összesítő | Cég / év / termékcsoport, cég- és pénzügyi adatok |
| Excel részletes | HAK-részletek, adattípus, kg/t és összecsukható kimutatás |
| Excel KTJ | Telephelyi tételek és betöltött metaadatok, kg |
| CSV adatcsere | Szűrt részletes adatok; UTF-8 BOM, pontosvesszős elválasztás |
| JSON szűrt adatok | Szűrt nyers rekordok és hozzájuk tartozó cégek; teljes hibajegyzék |
| Teljes JSON mentés | Összes forrásadat, nyers rekord és hiba, szűréstől függetlenül |
| Hibajegyzék CSV | Teljes munkamenet kimaradt/hibás lekérései |
| Vezetői riport | Nyomtatható HTML; a böngészőből PDF-ként menthető |

PDF: asztali böngészőben **Nyomtatás → Mentés PDF-ként**. iOS-en
**Megosztás → Nyomtatás → előnézet megnyitása → Megosztás → Mentés a
Fájlokba**. Ez böngészős PDF-mentés, nem szerveroldali PDF-generálás.
Letöltéskor Safariban a Fájlok / Letöltések mappát is ellenőrizd.

A JSON `format: "meh-cegelemzo"`, `version: 1`, `quantity_unit: "kg"`
mezőket és `source`, `raw`, `missing` tömböket tartalmaz. Az azonosítók
szövegek, a mennyiségek számok. A JSON adatcsere nem tárol nézeti
szűrőállapotot vagy KTJ-metaadat-cache-t; ezek asztali `.okirpkg` esetén
továbbra is megmaradnak. Régi csomag átviteléhez az asztali appban nyisd meg
az `.okirpkg`-t, majd exportálj teljes JSON-t. A web **nem olvas pickle-t**.
Régi `.okirpkg`-t asztalon is kizárólag megbízható forrásból nyiss meg.

## Asztali indítás

```bash
pip install -r requirements.txt
python app.py
```

A teljes forrásimport és hulladéklekérés háttérszálon fut. A táblázatok
modellalapúak: csak a látható cellák szövege készül el, nem soronként több
ezer Qt-elem. A nem látható nézetek frissítése a megnyitásukig halasztott.
Az asztali KTJ-metaadatbetöltés és a nagy Excel-export még szinkron
művelet; nagy adathalmaznál erre a webes háttérbetöltés használható.

A kategóriakonfiguráció a forráshoz/buildhez képest kerül feloldásra;
PyInstaller esetén a spec csomagolja a `config/product_groups.json` fájlt.

## Felépítés és sebesség

- `services/`: közös import, korlátozott párhuzamos OKIR-lekérés, szűrés,
  aggregálás, export, tárolás. Nincs Qt-függőség.
- `web/`: Flask API, jelszavas munkamenet/CSRF, egyfeladatos háttérfeldolgozó,
  lapozott táblázatok, reszponzív felület és adatokat nem cache-elő PWA.
- `ui/`, `workers/`: asztali megjelenítés és vékony Qt háttérszáladapter.
- `tests/`: szintetikus, külső OKIR-hálózatot nem használó regressziós és
  Chromium/WebKit böngészős ellenőrzések.

Gyorsítás: indexelt HAK-besorolás; importonként egyszeri adószám-keresés;
legfeljebb 6 OKIR-szál, megőrzött közös lekérési ütemezés; kötegelt SQLite
mentés és snapshot-indexek; legfeljebb 8 szűrt webes eredmény gyorsítótárazása;
50 soros oldalak (API-maximum 100). A böngésző nem tölti le a teljes nyers
adatállományt egy táblanézethez.

## Fejlesztői ellenőrzés

```bash
pip install -r requirements.txt -r requirements-test.txt
QT_QPA_PLATFORM=offscreen python -m unittest discover -s tests -v
python -m playwright install --with-deps chromium webkit
python -m tests.browser_check
python -m tools.benchmark
```

A böngészős futás asztali Chromiumot, iPhone- és iPad-méretű WebKitet fed le:
belépés, import, szűrés, lapozás, CSV-letöltés, riport, offline visszaállás.
Ez **nem helyettesíti a fizikai iOS-eszközön végzett főképernyős telepítés
és PDF-megosztás ellenőrzését**. Az OKIR élő elérhetősége külön külső
függőség. A benchmark 100 000 szintetikus soron méri az aggregálást,
szűrést, mentést és visszatöltést; nem hálózati sebességígéret.

A CI mindkét ellenőrzést futtatja. A PWA PNG-ikonjai újragenerálhatók:
`python tools/generate_icons.py` (Pillow, az asztali matplotlib függősége).
