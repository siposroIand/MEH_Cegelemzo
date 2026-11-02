# MÉH Cégelemző

Magyar nyelvű **asztali alkalmazás** OKIR hulladékadatok és céges
Excel-forrásadatok elemzéséhez. Az alapértelmezett `python app.py` ismét
natív PySide6/Qt ablakot nyit: nem indít webszervert, nem kell hozzá
böngésző vagy webes jelszó.

## Indítás

Python **3.11 vagy újabb** szükséges.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Linux/macOS alatt az aktiválás: `source .venv/bin/activate`.
Linuxon a Qt rendszerfüggőségei is szükségesek; Ubuntu alatt például
`libegl1` és `libopengl0`.

## Használat

1. **Excel betöltése:** válassz `.xlsx` forrást és lekérdezési típust.
   A forrásimport, az OKIR-lekérés, az aggregálás és az automatikus helyi
   mentés háttérben fut. Sikertelen betöltés nem cseréli le az aktív adatokat.
2. **Szűrés:** cég, év, termékcsoport, adattípus és HAK szerint. Több érték
   is kijelölhető; üres kijelölés az összeset jelenti. A HAK-mező részletekre
   is keres, több kód `;`-vel választható el. A grafikonokra kattintva is
   szűrhetsz. Betöltéskor a legutóbbi két adatév aktív; a **Szűrők törlése**
   minden évet visszahoz.
3. **Cég × Év / HAK mátrix:** a felső kiemelt gomb vagy az **Exportálás**
   menü elkészíti az aktuális szűrés szerinti Excel-kimutatást.

A világos, egységes felületen külön áttekintés, adattáblák, KTJ-nézet és
hibajegyzék található. A táblázatok a látható cellákat modellből jelenítik
meg; a szűrőgépelés késleltetett, a nézetek csak megnyitáskor frissülnek.
A hosszabb Excel-exportok, mentésbetöltések és KTJ-lekérések háttérben
futnak. Egyszerre egy ilyen művelet indulhat; bezárás előtt az alkalmazás
megváratja a befejezést. Az export idején a szűrők tovább használhatók,
de a fájl az export indításakor kijelölt adatokat tartalmazza.

### Új Cég–Év–HAK Excel-mátrix

A cégek **függőlegesen**, az év–HAK párok **vízszintesen** szerepelnek.
Három fejlécsor teszi egyértelművé, melyik évhez és hulladékhoz tartozik
egy mennyiség. Példa:

| | B | C | D | E |
|---|---|---|---|---|
| ÉV | 2024 | 2025 | 2025 | Összes év |
| HAK | 150101 | 150101 | 170405 | Összes hulladék |
| Cég / HAK MEGNEVEZÉS (t) | Papír csomagolás | Papír csomagolás | Vas és acél | Összes hulladék (t) |
| Minta Cég A | 1,250 | – | 0,500 | 1,750 |
| Minta Cég B | – | 0,750 | 2,000 | 2,750 |

- Mértékegység: **tonna**; az alapadatok kg-ban vannak.
- Nulla és hiányzó kombináció esetén **`-`** jelenik meg, kitöltés nélkül.
  Ez számformátum: a mögöttes érték 0 marad, ezért az Excelben továbbra is
  összegezhető. Sikertelen OKIR-lekérés esetén a jel nem bizonyít nulla
  hulladékot; ellenőrizd a **Kimaradt** nézetet.
- Pozitív érték: **zöld cella**; negatív korrekció: **narancssárga cella**.
  A kis, nem nulla értékek is színt kapnak.
- A cellák teljes számértéket tárolnak, legalább 3, legfeljebb 9 tizedest
  jelenítenek meg. A kódok szövegek, így a kezdő nullák megmaradnak.
- Az utolsó oszlop a sor összes hulladékát tartalmazza a kiválasztott
  évekre, HAK-kódokra és adattípusokra. Több adattípus összege nem nettó
  anyagmérleg. Az azonos nevű cégek KÜJ-azonosítóinak mennyiségei összeadódnak.
- Az átfedő termékcsoportok nem többszörözik a mennyiségeket.
  Egy év–HAK pár több megnevezése pontosvesszővel elválasztva marad meg.
  Hiányzó év külön **Ismeretlen év** oszlopcsoportba kerül.
- A három fejlécsor és a cégoszlop rögzített. A **Szűrés** munkalap
  dokumentálja a szűrőket, az egységet és a jelmagyarázatot.
- Üres eredmény vagy az Excel sor-/oszlopkorlátjának túllépése esetén
  egyértelmű hibaüzenet jelenik meg. A fájl ideiglenes helyre készül,
  és csak sikeres írás után váltja fel a célfájlt.

A termékcsoportos összesítő, a részletes HAK-lista, a KTJ-export és a
kimaradt lekérések Excel-exportja továbbra is elérhető. A KTJ-export a
**KTJ szerint** nézetben már betöltött metaadatokat használja; önmagában
nem indít új hálózati lekérést.
A vezetői riport HTML-ként menthető; böngészőből nyomtatással PDF-be menthető.

### Megosztható PKG és további adatexportok

Az **Exportálás** menüben, illetve a webes **Megosztható csomag és
adatexportok** szakaszban az alábbi lehetőségek is elérhetők:

| Export | Tartalom és szűrés |
|---|---|
| Megosztható `.okirpkg` | Teljes forrás-, nyers- és hibatábla, a rendelkezésre álló KTJ-metaadatok; asztali exportnál az aktuális szűrők és nézet is |
| MÉH `.json` | Teljes forrás-, nyers- és hibatábla; visszaimportálható, KTJ-metaadatok és nézetbeállítások nélkül |
| Teljes adatállomány `.xlsx` | Három munkalap: Forráscégek, Nyers adatok (kg), Kimaradt lekérések; szűréstől függetlenül |
| Nyers adatok `.csv` | Minden nyers sor, amely megfelel az aktuális cég-, év-, termékcsoport-, HAK- és adattípus-szűrésnek |
| Forráscégek `.csv` | A cégszűrésnek megfelelő forrásadatok; a többi szűrő nem vonatkozik rájuk |
| Kimaradt lekérések `.csv` | A teljes munkamenet hibalistája |

A PKG, JSON és teljes Excel **a szűrőkkel elrejtett adatokat is tartalmazza**;
megosztás előtt ezt vedd figyelembe. A PKG nem titkosított.
A CSV UTF-8 BOM-os, pontosvesszős, tizedesvesszős fájl, a mennyiségek kg-ban
maradnak. A képletnek értelmezhető szövegek elé védő aposztróf kerül.
Az azonosítókat CSV megnyitásakor szövegoszlopként importáld Excelbe, hogy
a kezdő nullák megmaradjanak. A JSON/PKG és az új teljes Excel ezeket
szövegként őrzi meg. A teljes Excel adatexport, nem forrásimport-sablon;
visszatöltéshez a PKG-t vagy a MÉH JSON-t használd.

Az **Importálás → Megosztható PKG betöltése** megnyitja az új és a régi
`.okirpkg` csomagokat is (`.pkg` és `.zip` kiterjesztés is választható).
Az új formátum ZIP-ben tárolt, verziózott JSON-adatokat használ, nem pickle-t.
Az import nem indít OKIR-lekérést, lecseréli az aktív adatokat, helyi mentést
készít, és asztalon visszaállítja a csomagban tárolt szűrőket.
Sikertelen importnál az aktív adatállomány változatlan marad.

A régi, pickle-alapú csomagokat **csak megbízható forrásból**, az asztali
alkalmazásban, külön megerősítés után töltsd be. A korlátozott olvasó a régi
exportáló szokásos pandas/NumPy táblatípusait támogatja; idegen vagy nem
támogatott objektumokat elutasít. Ez nem általános pickle-olvasó vagy
biztonsági sandbox. A betöltött csomagot újraexportálva már az új formátumot
kapod. A webes változat csak az új PKG-t fogadja el; az asztali nézetbeállításokat
nem alkalmazza. Az új PKG-t a régi alkalmazásverziók nem tudják megnyitni.
PKG-nél a fájl és a kicsomagolt tartalom is legfeljebb 100 MB,
JSON/PKG esetén táblánként legfeljebb 250 000 sor támogatott.

## Adatok és helyi mentések

Az Excel kötelező oszlopai: `Teljes név` / `Cégnév`, `Adószám`, `Megye`,
`Irányítószám`, `Település`, `Cím`. Az ékezet nélküli fejlécváltozatok és az
opcionális kapcsolattartási/pénzügyi mezők is működnek. Az OKIR-cégkeresés
az adószám első nyolc számjegyét használja.

Az OKIR lekérdezéshez internet szükséges. A már mentett adatok helyben,
internet nélkül is megnyithatók, szűrhetők és exportálhatók. Részleges
eredmény is mentődik; a hiányokat a **Kimaradt** nézet mutatja.

A meglévő `data/okir_cache.db` sémája változatlan. Az asztali alkalmazás
automatikusan visszatölti az utoljára aktív asztali mentést; ennek hiányában
az aktív webes, majd a legújabb mentést. Az **Exportálás → Helyi mentés**
elnevezett mentést készít, a **Korábbi mentés** gomb a régebbieket nyitja meg.
A korábbi és új MÉH JSON mentések egyaránt betölthetők. Másik gépre
átvihető, visszatölthető mentéshez az új megosztható PKG ajánlott.

Alapértelmezésben az indítás könyvtára tartalmazza a `data/`, `logs/` és
`output/` mappákat. A `MEH_DATA_HOME` környezeti változóval másik, írható,
tartós munkakönyvtár adható meg. A meglévő mentések használatához ugyanarra
a könyvtárra mutasson, mint korábban. Ne futtass egyszerre több példányt
vagy webes és asztali importot ugyanazzal az adatkönyvtárral.

Biztonsági mentéshez állítsd le az alkalmazást, majd másold a teljes
adatkönyvtárat, vagy használj SQLite backupot. Futó WAL-adatbázisból a
`.db` fájl önmagában nem elég.

## Windows-csomag készítése

Windows alatt, az aktivált virtuális környezetben:

```powershell
pip install -r requirements-build.txt
pyinstaller --noconfirm MEH_Cegprofil.spec
```

Az eredmény a `dist/MEH_Cegprofil/` könyvtár; a teljes könyvtárat kell
átadni, nem csak az `.exe` fájlt. Az adatkönyvtár legyen írható, ne
a `Program Files` alatt tárold. A csomaghoz nem kell külön Python.
A **Windows desktop build** GitHub Actions munkafolyamat pull requestnél
automatikusan fut, és kézzel is indítható:
teszteli és ZIP-ként feltölti a Windows-csomagot. Linuxos buildből
nem lesz Windows EXE; a Windows-kiadást Windows gépen is ellenőrizni kell.

## Opcionális webes változat

A meglévő webes változat megmarad, de nem ez az alapértelmezett alkalmazás.
Külön környezetben telepíthető, Qt és matplotlib nélkül:

```bash
pip install -r requirements-web.txt
read -rs -p "Munkaterület jelszava (legalább 12 karakter): " MEH_WEB_PASSWORD
echo
export MEH_WEB_PASSWORD
python web_app.py
```

Windows PowerShellben a `MEH_WEB_PASSWORD` változót a helyi titokkezelésből
állítsd be. Alapértelmezett cím: **http://127.0.0.1:8080**.
A webes felület korábbi, két fejlécsoros cég–HAK exportja nem változott;
az új év szerinti mátrix az asztali alkalmazásban érhető el.

| Környezeti változó | Jelentés |
|---|---|
| `MEH_WEB_PASSWORD` | Kötelező, legalább 12 karakteres közös munkaterület-jelszó |
| `MEH_SECRET_KEY` | Opcionális titkos kulcs; nélküle újraindításkor új belépés szükséges |
| `MEH_DATA_HOME` | Írható, tartós munkakönyvtár |
| `MEH_HOST` / `MEH_PORT` | Alapérték `127.0.0.1` / `8080` |
| `MEH_HTTPS` | HTTPS reverse proxy mögött `1` |

A webapp közös munkaterület, nem felhasználónként elkülönített SaaS.
Egy adatkönyvtárhoz egy szerverfolyamatot használj. Távoli használathoz
HTTPS reverse proxy szükséges; a backend maradjon loopback címen.
iPhone/iPad Safariban HTTPS-en a főképernyőhöz adható, de nem offline
elemző. Céges adatok nem kerülnek a service worker gyorsítótárába.
Webes feltöltés: maximum 20 MB, 100 MB kicsomagolt XLSX, 2000 forráscég;
JSON/PKG-importnál táblánként 250 000 sor. A 20 MB-os webes feltöltési korlát
a JSON- és PKG-fájlokra is vonatkozik.

## Fejlesztés

`ui/`: Qt felület és táblamodellek. `workers/`: asztali háttérműveletek.
`services/`: közös import, OKIR, aggregálás, szűrés, export és SQLite-tárolás.
`web/`: opcionális Flask API és PWA.

```bash
pip install -r requirements-test.txt
QT_QPA_PLATFORM=offscreen python -m unittest discover -s tests -v
python -m tools.benchmark
python -m playwright install --with-deps chromium webkit
python -m tests.browser_check
```

PowerShellben a tesztparancs előtt `$env:QT_QPA_PLATFORM = "offscreen"`.
A tesztek szintetikus adatokkal, mockolt OKIR-hálózattal ellenőrzik a
mátrix tényleges celláit, színeit, számformátumát, évbontását és szűrését,
a háttérműveletek alatti Qt eseménykezelést, a hibás export sértetlen
célfájlját és a helyi mentések visszatöltését. A benchmark 100 000 soron
mér; nem jelent hálózati sebességígéretet.
