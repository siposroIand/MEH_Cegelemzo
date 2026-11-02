# MÉH Cégelemző

Teljesen webes, magyar nyelvű alkalmazás OKIR hulladékadatok és céges
forrásadatok elemzéséhez. Asztali böngészőben, iPhone-on és iPaden is
használható; Safariból a főképernyőre tehető. A feldolgozás a szerveren fut.
Nincs telepítendő asztali kliens, Qt, matplotlib vagy Windows-build.

## Három lépés

1. **Excel betöltése:** tölts fel egy céges `.xlsx` forrásfájlt, válaszd ki
   a lekérendő adattípusokat, majd indítsd a betöltést.
2. **Szűrés:** válassz céget, évet, termékcsoportot, adattípust vagy HAK-kódot.
   A grafikonokra kattintva is szűrhetsz. Az aktuális szűrés a mutatók
   fölött látható.
3. **Kimutatás letöltése:** a kiemelt gomb a cég–HAK Excel-mátrixot készíti
   el; mellette nyomtatható vezetői riport érhető el.

Az import a háttérben fut; közben a korábbi adatok olvashatók. Egyszerre egy
módosító feladat futhat. Sikeres import után automatikus szerveroldali mentés
készül; a **Korábbi mentések** panelen elnevezett mentést is készíthetsz.
Újraindításkor az aktív mentés áll vissza.

### Cég–HAK Excel-mátrix

A `Cég–HAK mátrix` munkalap pontos elrendezése:

| | A | B | C |
|---|---|---|---|
| 1. sor | Cég | 150101 | 170405 |
| 2. sor | Mennyiség (t) | Papír csomagolás | Vas és acél |
| 3. sor | Minta Cég A | 1,250 | 0,500 |
| 4. sor | Minta Cég B | 0,000 | 2,750 |

- A cégek függőlegesen, a HAK-kódok vízszintesen jelennek meg.
- A HAK-kód alatt a megnevezése szerepel. Ha a szűrt adatokban ugyanahhoz
  a kódhoz több megnevezés tartozik, mindegyik megmarad, pontosvesszővel
  elválasztva; hiányzó névnél „Nincs megnevezés” látható.
- A cellák **tonnában összesített, valódi számértékek**. A kijelölt évek,
  adattípusok és azonos nevű cégek KÜJ-azonosítóinak mennyiségei összeadódnak.
  A több termékcsoportba tartozó tétel nem számít többször.
- Három tizedes jelenik meg, de a tárolt számértéket nem kerekítjük.
  A hiányzó cég–HAK kombináció értéke 0; ez nem jelenti azt, hogy egy
  sikertelen OKIR-lekéréshez bizonyítottan nulla hulladék tartozik.
- A két fejlécsor és a cégnevek görgetéskor rögzítve maradnak.
- A külön `Szűrés` munkalap rögzíti az export szűrőit és mértékegységét.
  Üres eredményből nem készül félrevezető mátrix.

A **További Excel-kimutatások** összecsukható részben megmaradt a
termékcsoportos összesítő, a részletes lista, a KTJ-kimutatás és a kimaradt
lekérések listája. **CSV- és JSON-export nincs**, a korábbi export-URL-ek
404-et adnak. A hibajegyzék is `.xlsx` formátumú.

PDF: a nyomtatható riportot böngészőben a **Nyomtatás → Mentés PDF-ként**
útvonalon mentheted. iOS-en **Megosztás → Nyomtatás → előnézet megnyitása →
Megosztás → Mentés a Fájlokba**. Ez böngészős PDF-mentés, nem szerveroldali
PDF-generálás.

## Adatok és szűrés

Az Excel kötelező oszlopai: `Teljes név` / `Cégnév`, `Adószám`, `Megye`,
`Irányítószám`, `Település`, `Cím`. Az ékezet nélküli fejlécváltozatok és az
opcionális kapcsolattartási/pénzügyi mezők is működnek. Az OKIR-cégkeresés
az adószám első nyolc számjegyét használja.

Maximum 20 MB-os feltöltés, 100 MB kicsomagolt XLSX és 2000 Excel-forráscég
engedélyezett egy importban. Hibás fájl nem cseréli le az aktív adatokat.
Részleges OKIR-eredmény új munkamenetként mentődik; a korábbi mentés
visszatölthető. A sikertelen lekérések a **Kimaradt / hibák** nézetben
láthatók. Félbeszakadt háttérlekérést újra kell indítani.

Több érték is kijelölhető. Üres kijelölés: összes. Induláskor és új adat
betöltésekor a legutóbbi két adatév aktív; a **Szűrők törlése** minden évet
visszahoz. A HAK-mező részletekre is keres, több kód `;`-vel választható el.
A cégadatok nézetére csak a cégnévszűrő, a hibajegyzékre semmilyen elemzési
szűrő nem vonatkozik. A KTJ-címadatok külön gombbal tölthetők be; a
KTJ-export a már betöltött metaadatokat tartalmazza.

Az alapadatok mértékegysége kg; az Excel-mátrix tonnára vált. A termékcsoportok
átfedhetnek, ezért csoportösszegeik nem adhatók össze. Több adattípus összege
nem nettó anyagmérleg.

### Korábbi adatok megőrzése

A meglévő `data/okir_cache.db` mentései és cache-e változatlan sémával
tovább használhatók. Adatot vagy helyi mentést az átállás nem töröl.
Korábbi verzióból már elkészített MÉH JSON csak **importként** maradt meg:
**Korábbi mentések → Adatkezelési beállítások → Régi mentés betöltése**.
Ez nem új JSON-export funkció. A 20 MB-os feltöltési korlát és táblánként
250 000 soros importkorlát erre is vonatkozik.

Az asztali forrás, a régi Windows-build és az elavult asztali kézikönyv
eltávolításra került. A korábbi változat a Git-előzményekben elérhető.
Pickle/`.okirpkg` olvasó vagy író nincs a webes alkalmazásban.

## Indítás

Python **3.11 vagy újabb** szükséges.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
read -rs -p "Munkaterület jelszava (legalább 12 karakter): " MEH_WEB_PASSWORD
echo
export MEH_WEB_PASSWORD
python app.py
```

Windows PowerShell esetén az aktiválás `.venv\Scripts\Activate.ps1`;
a `MEH_WEB_PASSWORD` változót a helyi titokkezelésből állítsd be.

Alapértelmezett cím: **http://127.0.0.1:8080**. A szerver Waitress-t használ.
A korábbi `python web_app.py` indítás és `requirements-web.txt` telepítés
kompatibilitási belépési pontként továbbra is ugyanazt a webappot indítja.
Régi környezet helyett új virtuális környezetet használva a korábbi asztali
függőségek sem maradnak telepítve.

### iPhone / iPad és HTTPS

1. A Python-szervert egy tartósan elérhető gépen kell futtatni. Az iPhone-on
   a `localhost` az iPhone-t jelenti, nem a szervert.
2. Tegyél elé HTTPS reverse proxyt, például Caddyt. A saját DNS-név mutasson
   a szerverre; a 80/443 port legyen elérhető:

   ```caddyfile
   cegelemzo.example.hu {
       reverse_proxy 127.0.0.1:8080
   }
   ```

3. HTTPS mögött állítsd be `MEH_HTTPS=1`-re a biztonságos session-cookie-hoz.
   A backend maradjon loopback címen; a folyamatot szolgáltatáskezelővel
   (például systemd) indítsd.
4. Safariban nyisd meg a saját HTTPS-címet, jelentkezz be, majd válaszd a
   **Megosztás → Főképernyőhöz adás** lehetőséget.

A főképernyős használathoz HTTPS kell (fejlesztéskor a localhost kivétel).
Az app nem natív App Store-alkalmazás, és nem offline elemző. Hálózat
nélkül tájékoztató jelenik meg; cégadatok, API-válaszok, exportok és
hitelesített HTML nem kerülnek a service worker gyorsítótárába.

| Környezeti változó | Jelentés |
|---|---|
| `MEH_WEB_PASSWORD` | Kötelező, legalább 12 karakteres közös munkaterület-jelszó |
| `MEH_SECRET_KEY` | Opcionális, titokként kezelt véletlen kulcs; nélküle újraindításkor új belépés szükséges |
| `MEH_DATA_HOME` | Írható, tartós munkakönyvtár; alapérték az indítás könyvtára |
| `MEH_HOST` / `MEH_PORT` | Alapérték `127.0.0.1` / `8080` |
| `MEH_HTTPS` | `1`: csak HTTPS-en küldött session-cookie |

Ez egy **közös, jelszóval védett munkaterület**, nem felhasználónként
elkülönített SaaS. Külön csapatoknak külön példány, jelszó és adatkönyvtár
szükséges. Egy példányhoz **egy szerverfolyamatot** használj; ne indíts több
workert/konténert ugyanarra az adatkönyvtárra.

A `MEH_DATA_HOME/data/okir_cache.db` tárolja a mentéseket, metaadatokat,
OKIR-cache-t és az aktív mentés azonosítóját. Biztonsági mentéshez SQLite
backupot használj, vagy állítsd le a szervert az adatkönyvtár másolásakor.
Futó WAL-adatbázisból a `.db` önmagában nem elég.

## Fejlesztés

`services/`: import, ütemezett párhuzamos OKIR-lekérés, szűrés, aggregálás,
Excel-export, riport és SQLite-tárolás. `web/`: Flask API, jelszavas
munkamenet/CSRF, háttérfeladatok, lapozott táblázatok és PWA-felület.

```bash
pip install -r requirements-test.txt
python -m unittest discover -s tests -v
python -m playwright install --with-deps chromium webkit
python -m tests.browser_check
python -m tools.benchmark
```

A regresszió az Excel-mátrix tényleges celláit, két fejlécsorát, tonnára
váltását, szűrését és duplázásmentes összegzését is ellenőrzi. A böngészős
folyamat Chromiumon, iPhone- és iPad-méretű WebKiten fut, valódi `.xlsx`
letöltéssel. A tesztadatok szintetikusak; az OKIR-hálózat mockolt.
Fizikai iOS-telepítést és PDF-megosztást külön kell ellenőrizni.

A mátrix Excel-író soronként ír, nem épít nagy cellamátrixot a memóriában.
A benchmark 100 000 soron méri az aggregálást és tárolást; ez nem hálózati
sebességígéret. Ikonok újragenerálása: `python tools/generate_icons.py`.
