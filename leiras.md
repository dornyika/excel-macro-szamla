# Számla VB scriptek — Leírás

## Áttekintés

Két egymás után futtatandó script alkotja a feldolgozási folyamatot:

| Sorrend | Script | Modul neve | Mire való? |
|---------|--------|-----------|------------|
| 1. | `szamla-pre.vb` | `SzamlaPreProcess` | Sablonfájl feltöltése forrásfájlokból, átnevezés |
| 2. | `szamla.vb` | `SzamlaProcess` | Formázás, rendezés, szűrés, lookup-ok |

Mindkettő a `Personal.xlsb`-be kerül modulként.

---

## ⚠️ Mit kell módosítani?

### szamla-pre.vb — kötelező ellenőrzés minden új sablonnál

A script egyetlen hardkódolt konstanst tartalmaz, amelyet szükség esetén frissíteni kell:

```vb
' ~36. sor, SzamlaPreProcess modul tetején:
Const TARGET_FILENAME As String = "pipalo_v20250827_tol_sablonom.xlsx"
```

**Ha új pipalo sablon készül** (pl. `pipalo_v20260101_tol_sablonom.xlsx`), ezt a sort kell átírni az új fájlnévre. Különben a script nem találja meg a sablont az almappákban, és minden mappa `[WARN]`-nal vagy hibával végez.

### szamla.vb — nincs rendszeres módosítás

Nincsenek felső szintű konstansok. A sheetnevek (`"vissz"`, `"xls"`, `"tsz"`, `"regi_tsz_utolsó4honap"`, `"utólagos"`, `"Nemzeti TAJ"`, `"reszletes"`) hardkódolva vannak az egyes Sub-okban. Ha a sablon sheetnevei változnak, minden érintett Sub-ban kell keresni és cserélni.

---

---

# szamla-pre.vb — Leírás

## Mire való?

Ez a script a `szamla.vb` **előtt** fut. Feladata: minden kiválasztott almappában megtalálja a sablonfájlt (`pipalo_v20250827_tol_sablonom.xlsx`), és feltölti adatokkal az adott almappában lévő forrásfájlokból.

A script a `Personal.xlsb`-be kerül, és két makróval indítható:
- `Run_PreProcess_Batch` — teljes batch feldolgozás (összes almappa)
- `Run_PreProcess_Single` — egymappa-feldolgozás hibakereséshez

---

## Futtatás menete

1. **Periódus megadása** — két InputBox: 2 jegyű év (pl. `26`), majd hónap (1–12). Ebből lesz a perióduskód, pl. `v2602`.
2. **Szülőmappa kiválasztása** — mappapicker dialógus. Az almappáit fogja végigpásztázni.
3. **Almappák szűrése** — csak azok az almappák kerülnek feldolgozásra, amelyek tartalmazzák a sablonfájlt, és amelyek neve **nem** végződik `-nincs`-re. A feldolgozás sorrendje ABC (almappa neve szerint).
4. **Megerősítés** — megmutatja: hány almappa, melyik periódus, és a `regi_tsz` szűrésnél melyik 4 YYMM kódot keresi az S oszlopban.
5. **Batch feldolgozás** — almappánként:
   - Megnyitja a sablonfájlt
   - Feltölti a 7 munkalapot forrásfájlokból
   - Menti, majd **átnevezi** a sablonfájlt: `[almappa neve]_vYYMM.xlsx` (pl. `Arpad_v2602.xlsx`) — azonos nevű meglévő fájlt felülírja
   - Naplózza az eredményt: `[OK]`, `[WARN]` (ha valamely forrásfájl hiányzott), vagy `[ERROR]` (futási hiba — a fájl nem kerül mentésre, nem lesz átnevezve)
6. **Napló** — a szülőmappában létrehoz egy `preprocess_log_YYYYMMDD_HHMMSS.txt` fájlt.
7. **Összegzés** — MsgBox: hány sikerült, hány hibás, a naplófájl elérési útja.

---

## Feltöltött munkalapok

A sablon 5 eredeti lapot tartalmaz. A script eléjük másolja a `reszletes` és `Nemzeti TAJ` lapokat, így a végső sorrend:

| # | Lap neve | Forrásfájl mintája | Honnan tölt? |
|---|----------|--------------------|--------------|
| 1 | Nemzeti TAJ | `*nemztaj*` | Sheets(1) teljes másolata |
| 2 | reszletes | `*reszletes*` | Sheets(1) teljes másolata |
| 3 | vissz | `vissz_*` | A1:AN, értékként |
| 4 | xls | `xls_*TITKOS*` (fő) + többi `xls_*` | 2. sortól, 50 oszlop; másodlagos fájlok 3. sortól, piros betűvel |
| 5 | tsz | `tsz_*` | A2-től, 22 oszlop, értékként |
| 6 | regi_tsz_utolsó4honap | `*_tsz.xls*` | S oszlop alapján szűrt sorok (előző 4 hónap YYMM kódjai) |
| 7 | utólagos | `*_vYYMM.xls*` (nem `p_` kezdetű) | A2-től, 15 oszlop, értékként |

### Megjegyzések

- **Nemzeti TAJ és reszletes**: teljes lapmásolás (formátumok megőrződnek).
- **Összes többi lap**: csak értékek kerülnek át (`.Value = .Value`), nincs formázásmásolás.
- **tsz forrásfájl**: a `Munka1` nevű sheetből olvas.
- **regi_tsz forrásfájl**: ha csak 1 sheet van, azt használja. Ha több sheet van, az A1 cellájában `Törzs-szám` vagy `Törzs-számok` szöveget tartalmazó sheetet keresi (sortörés és kis/nagybetű figyelmen kívül hagyva).
- **xls TITKOS fájl**: az alapnevének `TITKOS`-ra kell végződnie (kis/nagybetű mindegy). A többi `xls_*` fájl másodlagos — adatai piros betűszínnel kerülnek be a fő adatok alá.
- **regi_tsz szűrés**: az S oszlop (19.) értékében keresi az előző 4 hónap YYMM kódjait (pl. `2601`, `2512`), `InStr`-rel — kis/nagybetű-érzéketlen.
- **NumberFormat**: TAJ-számokat és egyéb hosszú numerikus azonosítókat tartalmazó oszlopok szöveg formátumra (`"@"`) vannak állítva, hogy elkerüljük a tudományos jelölést.

---

## Előző 4 hónap kiszámítása

A `regi_tsz` szűréshez a script visszaszámol 4 hónapot a megadott periódusból:

- `v2602` → szűr: `2601`, `2512`, `2511`, `2510`
- Évváltásnál automatikusan kezeli (pl. 00 után 99 lesz)

---

## Napló formátuma

```
Pre-Process Log
Date: 07/03/2026 14:23:11
Period: v2602
--------------------------------------------------
[OK]    Arpad
[WARN]  Bajcsy
    WARN: utólagos source not found (not p_*, ending with _v2602)
[ERROR] Keleti - ERROR:1004 ...
--------------------------------------------------
Processed: 2
Errors:    1
Duration:  1:23
```

- `[OK]` — minden forrásfájl megtalálva, feltöltés sikeres
- `[WARN]` — egy vagy több forrásfájl hiányzott, de ami megvolt, az feltöltődött; a fájl mentésre kerül
- `[ERROR]` — futási hiba történt; a fájl **nem** kerül mentésre

---

## Egymappa-feldolgozás (`Run_PreProcess_Single`)

Hibakereséshez: a felhasználó közvetlenül kiválaszt egy almappát (nem szülőmappát), és csak az kerül feldolgozásra. Ugyanazt a logikát használja, mint a batch — a napló és az eredményüzenet szintén megjelenik. Hálózati probléma gyanúja esetén érdemes először lokálisan futtatni.

---

## Technikai megjegyzések

- A script `Personal.xlsb`-be kerül, `SzamlaPreProcess` modul névvel.
- Batch futás közben `ScreenUpdating`, `Calculation`, `EnableEvents`, `DisplayAlerts` és `Interactive` ki van kapcsolva — mentés előtt minden visszaállítódik. Az `Interactive = False` megakadályozza, hogy véletlen kattintás vagy gépelés megzavarja a futást.
- Váratlan hiba esetén (`FatalError` ág) az összes beállítás garantáltan visszaáll, a nyitott munkafüzet bezárul.
- Egymásba ágyazott `Dir()` helyett `FileSystemObject`-et használ az almappák iterálásához.
- Forrásworkbookok `ReadOnly:=True` módban nyílnak meg.
- `SetForegroundWindow` (Win32 API) gondoskodik arról, hogy az Excel az előtérbe kerüljön a feldolgozás végén, így az összegző MsgBox és a hangjelzés nem marad háttérben.

---

---

# szamla.vb — Leírás

## Mire való?

Ez a script a `szamla-pre.vb` **után** fut. Feladata: a feltöltött sablonfájlok munkalapjait formázza, rendezi, szűri és lookup-okkal egészíti ki.

A script a `Personal.xlsb`-be kerül, és több makróval indítható:

**Egyedi munkafüzethez:**
- `RunAll` — mind a 7 lap feldolgozása (utólagos-szal együtt)
- `RunAll_WithoutUtolagos` — utólagos lap kihagyása (ha manuálisan szerkesztik)

**Batch módban (mappában lévő összes fájlhoz):**
- `Run_BatchFolder` — a kiválasztott mappában lévő összes `.xls*` fájl feldolgozása
- `Run_BatchFolder_WithoutUtolagos` — ugyanaz, utólagos nélkül

**Egyedi lapok hibakereséshez:**
- `Run_Vissz`, `Run_Xls`, `Run_NemzetiTAJ`, `Run_Reszletes`, `Run_Tsz`, `Run_Regi_tsz`, `Run_Utolagos`

---

## Mit csinál laponként?

| Lap | Főbb műveletek |
|-----|---------------|
| **vissz** | Fejléc félkövér; rendezés TAJ + FELV szerint; panelek rögzítése B2-nél; másolat lap (`vissz (2)`) létrehozása szöveg formátummal; `0000`, `ARES`, `KRON` értékű M oszlop sorok törlése a másolatból |
| **xls** | Rendezés TAJ + FELV_IDO szerint; panelek rögzítése B2-nél |
| **Nemzeti TAJ** | F oszlopba `/` jel beillesztése a 4. karakter után (pl. `123456789` → `1234/56789`); sárga háttér az F oszlopon; üres sorok törlése; oszlopszélesség igazítás |
| **reszletes** | MEGYE és INTKOD oszlopok törlése; új A oszlop beillesztése TSZ-szel (piros betűszín, `/` jellel); P10 és TRAUMA oszlopok sárga kiemelése; rendezés TSZ + TRAUMA (csökkenő) + P10 (csökkenő) szerint; üres sorok törlése |
| **tsz** | 2–4. sorok és N,O,P,Q,V oszlopok törlése; L:M oszlopok törlése és átnevezése `Pipáló` / `Vissz2`-re; lookup-ok (ld. lent); AutoFilter: I/N = `I` ÉS Vissz2 ≠ `HIÁNYZIK`; szűrt sorokon jelölés |
| **regi_tsz_utolsó4honap** | N,O,P,Q,W oszlopok törlése; ugyanolyan lookup-ok mint tsz-nél; AutoFilter: Vissz2 ≠ `HIÁNYZIK`; VISSZA oszlop értékének másolása, majd törlése |
| **utólagos** | Tsz (piros) és regi_tsz (kék) szűrt sorainak összegyűjtése; rendezés TAJ szerint; K oszlop feltöltése (előző sor értékének propagálása); formázás: Calibri 11pt, középre, keret nélkül; felesleges üres sorok törlése |

---

## Pipáló és Vissz2 lookup

A `tsz` és `regi_tsz_utolsó4honap` lapokon a script két lookup oszlopot tölt ki:

| Oszlop | Neve | Keresés helye | Mit vesz át? |
|--------|------|--------------|--------------|
| L | **Pipáló** | `utólagos` lap A oszlopa | C oszlop értéke (ha megtalálta a TAJ-t) |
| M | **Vissz2** | `vissz (2)` lap A oszlopa | M oszlop értéke (ha megtalálta a TAJ-t) |

Ha a TAJ nem található: `"HIÁNYZIK"` értéket kap.

Az AutoFilter ezután a Vissz2 = `"HIÁNYZIK"` sorokat szűri ki. A szűrt (látható) sorokon a Pipáló értéke alapján jelöl:
- Ha Pipáló nem üres: `"online"` → a sor online úton érkezett
- Ha Pipáló üres: `"utólagos"` (tsz lapon) vagy a VISSZA oszlop értéke (regi_tsz lapon)

---

## Batch mód

A `Run_BatchFolder` / `Run_BatchFolder_WithoutUtolagos` makrók:
1. Mappapicker dialógus — a felhasználó kiválaszt egy mappát
2. Megkeresi az összes `.xls*` fájlt a mappában
3. Minden fájlt megnyit, feldolgoz, majd a `Modified\` almappába ment (ha nem létezik, létrehozza)
4. Az eredeti fájlok nem módosulnak

---

## Technikai megjegyzések

- Minden lapmakrónak van `Optional silent As Boolean = False` paramétere — batch módban `True`-val hívódnak, hogy ne jelenjenek meg MsgBox-ok.
- Futás közben `ScreenUpdating`, `Calculation`, `EnableEvents` ki van kapcsolva; batch mentés előtt visszaállnak (fájlkorrupció elkerülése érdekében).
- Lookup-ok `.Find()` metódussal dolgoznak (`LookIn:=xlValues, LookAt:=xlWhole` — pontos egyezés).
- Szöveg formátum (`"@"`) alkalmazva TAJ és TSZ azonosítókat tartalmazó oszlopokon, mielőtt az értékek bekerülnek.
- AutoFilter használat előtt mindig `ws.AutoFilterMode = False` hívódik.
