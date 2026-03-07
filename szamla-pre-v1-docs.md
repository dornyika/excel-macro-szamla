# szamla-pre-v1.vb — Teljes leírás

**Verzió:** 2
**Utolsó frissítés:** 2026-03-05
**Futtatás sorrendje:** Ez a script fut **ELŐBB**, mielőtt a `szamla-v107.vb` feldolgozó makrók futnak.

---

## Áttekintés

A script célja: egy kiválasztott szülőmappa almappáit végigjárja, és minden almappában megtalált sablonfájlt (`pipalo_v20250827_tol_sablonom.xlsx`) feltölt az ugyanabban az almappában lévő forrásfájlokból.

**Munkafolyamat:**

```
Forrásfájlok (almappánként)
    ↓
[szamla-pre-v1.vb] → Kitölti a sablon munkalapjait
    ↓
Kitöltött pipalo_*.xlsx
    ↓
[szamla-v107.vb] → Feldolgozza, létrehozza a végleges outputot
```

---

## Belépési pontok

| Makró neve | Leírás |
|---|---|
| `Run_PreProcess_Batch` | Futtatható gombhoz rendelhető belépési pont — meghívja a `RunPreProcess_BatchFolders`-t |

---

## Főfolyamat: `RunPreProcess_BatchFolders`

### 1. lépés — Időszak bekérése (`GetPeriodInput`)
- Két InputBox jelenik meg: év (2 jegy, pl. `26`) és hónap (`1–12`)
- Ebből épül fel a perióduskód: `v2602` (2026. február esetén)
- Ha az input érvénytelen vagy üres → megszakítás

### 2. lépés — Szülőmappa kiválasztása
- Mappaválasztó dialógus jelenik meg
- A felhasználó azt a mappát választja ki, amelynek **almappái** tartalmazzák a feldolgozandó fájlokat

### 3. lépés — Érvényes almappák összegyűjtése
- Az összes almappát bejárja
- Csak azokat tartja meg, amelyekben megtalálható a `TARGET_FILENAME` (`pipalo_v20250827_tol_sablonom.xlsx`)
- Ha egy almappában nincs sablonfájl → kihagyja (nem hiba)

### 4. lépés — Megerősítés
- MsgBox megmutatja: hány mappa kerül feldolgozásra, melyik periódus, regi_tsz filter tartomány
- `Nem` → megszakítás

### 5. lépés — Teljesítményoptimalizálás
```
ScreenUpdating = False
Calculation = Manual
EnableEvents = False
DisplayAlerts = False
DisplayStatusBar = True  ← státuszsor látható marad
```

### 6. lépés — Almappák feldolgozása (ciklus)

Minden érvényes almappára:

1. Megnyitja a sablonfájlt (`wb`) — írható módban
2. Meghívja a `TryFillAllSheets`-t → kitölti mind a 7 munkalapot
3. Visszaállítja a beállításokat (ScreenUpdating stb.) **mentés előtt** (kritikus!)
4. Eredmény szerint:
   - `[OK]` → ment és zár
   - `[WARN]` → ment és zár (néhány forrás hiányzott, de a folyamat lefutott)
   - `[ERROR]` → **nem ment** (sablon érintetlen marad újrafuttatáshoz), zár
5. Újra bekapcsolja az optimalizációkat a következő mappához

### 7. lépés — Napló mentése és összegzés
- Időbélyeges `.txt` naplófájl kerül a szülőmappába: `preprocess_log_YYYYMMDD_HHMMSS.txt`
- MsgBox összegzi az eredményt és megmutatja a napló elérési útját

---

## Munkalapkitöltő szubrutinok

A `FillAllSheets` orchestrátor hívja meg őket sorban. **A sorrend fontos:** a másolással beillesztett lapokat (reszletes, Nemzeti TAJ) először kell beilleszteni, hogy a végleges lapsorend helyes legyen.

```
1. FillReszletes    → beilleszti pos. 1-re
2. FillNemzetiTAJ   → beilleszti pos. 1-re → reszletes pos. 2-re tolódik
3. FillVissz
4. FillXLS
5. FillTsz
6. FillRegiTsz
7. FillUtolagos
```

### `FillNemzetiTAJ`
- **Forrásfájl:** `*nemztaj*` (bármilyen fájl, amelynek neve tartalmazza a `nemztaj` szót)
- **Módszer:** a forrásfájl 1. munkalapját **átmásolja** a sablonba (pos. 1), átnevezi `Nemzeti TAJ`-ra
- **Célmunkalap:** újonnan létrehozott másolat

### `FillReszletes`
- **Forrásfájl:** `*reszletes*`
- **Módszer:** a forrásfájl 1. munkalapját **átmásolja** a sablonba (pos. 1), átnevezi `reszletes`-re
- **Célmunkalap:** újonnan létrehozott másolat

### `FillVissz`
- **Forrásfájl:** `vissz_*`
- **Módszer:** értékek beillesztése (PasteSpecial xlPasteValues) az A1 cellától
- **Forrásadat:** 1. sortól az utolsó sorig, 40 oszlop (A:AN)
- **Célmunkalap:** `vissz` (már létezik a sablonban)
- **Megjegyzés:** A oszlop Text formátumra állítva beillesztés előtt (tudományos jelölés megelőzése)

### `FillXLS`
- **Forrásfájlok:** `xls_*` (több fájl is lehetséges)
- **Fő fájl:** amelynek alapneve `TITKOS`-ra végződik → adatok a 2. sortól, 50 oszlop → célba A1-től
- **Másodlagos fájlok:** a 3. sortól másolja, 50 oszlop → hozzáfűzi az adatokhoz, **piros betűszínnel** jelöli
- **Célmunkalap:** `xls` (már létezik a sablonban)

### `FillTsz`
- **Forrásfájl:** `tsz_*`
- **Módszer:** értékek beillesztése a 2. sortól, 22 oszlop (A:V) → célba A1-től
- **Célmunkalap:** `tsz` (már létezik a sablonban)
- **Megjegyzés:** A oszlop Text formátumra állítva

### `FillRegiTsz`
- **Forrásfájl:** bármely `*.xls*` amelynek alapneve `_tsz`-re végződik (pl. `valami_tsz.xlsx`)
- **Szűrés:** S oszlop (19.) mező) → az előző 4 hónap perióduskódjai (pl. `v2601`, `v2512`, ...)
- **Módszer:** AutoFilter → csak látható sorok másolása (fejléc + szűrt adatok) → PasteSpecial → célba A1-től
- **Célmunkalap:** `regi_tsz_utolsó4honap` (már létezik a sablonban)
- **Megjegyzés:** A oszlop Text formátumra állítva

### `FillUtolagos`
- **Forrásfájl:** `*.xls*`, amelynek neve **nem** kezdődik `p_`-vel, és alapneve `_vYYMM`-re végződik (pl. `valami_v2602.xlsx`)
- **Módszer:** értékek beillesztése a 2. sortól, 15 oszlop (A:O) → célba A1-től
- **Célmunkalap:** `utólagos` (már létezik a sablonban)
- **Megjegyzés:** A oszlop Text formátumra állítva (TAJ számok)

---

## Forrásfájl-mintázatok összefoglalása

| Munkalap | Keresési minta | Megjegyzés |
|---|---|---|
| Nemzeti TAJ | `*nemztaj*` | Bármilyen kiterjesztés |
| reszletes | `*reszletes*` | Bármilyen kiterjesztés |
| vissz | `vissz_*` | |
| xls | `xls_*` (több db) | Fő fájl: alapnév `TITKOS`-ra végződik |
| tsz | `tsz_*` | |
| regi_tsz | `*_tsz.xls*` | Alapnév `_tsz`-re végződik |
| utólagos | `*_vYYMM.xls*` | Nem kezdődik `p_`-vel |

---

## Hibanapló formátuma

A naplófájl a szülőmappába kerül: `preprocess_log_YYYYMMDD_HHMMSS.txt`

```
Pre-Process Log
Date: 05/03/2026 14:32:11
Period: v2602
--------------------------------------------------
[OK]    Arpad
[WARN]  Bajcsy
    WARN: utólagos source not found (not p_*, ending with _v2602)
    WARN: regi_tsz source not found (pattern: *_tsz.xls*)
[ERROR] Keleti - ERROR:1004 Range osztály AutoFilter metódusa hibás
--------------------------------------------------
Processed: 2
Errors:    1
```

| Jelölés | Jelentés | Mentve? |
|---|---|---|
| `[OK]` | Minden munkalap sikeresen kitöltve | Igen |
| `[WARN]` | Lefutott, de valamely forrás hiányzott (az érintett munkalap üres marad) | Igen |
| `[ERROR]` | Kivétel keletkezett — a sablon érintetlen marad | Nem |

---

## Hibakezelés

- **Megnyitási hiba** (`wb`): naplózva mint `[ERROR]`, a következő almappa feldolgozása folytatódik
- **Kitöltési hiba** (bármely Fill szubrutin kivételt dob): a `TryFillAllSheets` elkapja, visszaadja az `ERROR:` prefixű hibaüzenetet; `wb` mindig bezárásra kerül
- **Hiányzó forrásfájl / célmunkalap**: `[WARN]` bejegyzés a naplóba, az adott munkalap kihagyásra kerül, a többi Fill szubrutin tovább fut
- **Teljesítménybeállítások**: mindig visszaállításra kerülnek mentés előtt, még hiba esetén is

---

## Konstansok és globális változók

| Név | Típus | Leírás |
|---|---|---|
| `TARGET_FILENAME` | Const String | A sablonfájl neve — frissíteni kell új sablon verzió esetén |
| `g_warnings` | Private String | Egy mappa feldolgozása alatt gyűlt figyelmeztető üzenetek |

---

## Függvények és szubrutinok áttekintése

| Név | Típus | Leírás |
|---|---|---|
| `Run_PreProcess_Batch` | Sub | Belépési pont |
| `RunPreProcess_BatchFolders` | Sub | Főfolyamat |
| `GetPeriodInput` | Function | Év/hónap bekérése és validálása |
| `FillAllSheets` | Sub | Orchestrátor — mind a 7 Fill szubrutint sorban hívja |
| `FillNemzetiTAJ` | Sub | Nemzeti TAJ munkalap kitöltése |
| `FillReszletes` | Sub | reszletes munkalap kitöltése |
| `FillVissz` | Sub | vissz munkalap kitöltése |
| `FillXLS` | Sub | xls munkalap kitöltése (fő + másodlagos fájlok) |
| `FillTsz` | Sub | tsz munkalap kitöltése |
| `FillRegiTsz` | Sub | regi_tsz_utolsó4honap munkalap kitöltése (szűréssel) |
| `FillUtolagos` | Sub | utólagos munkalap kitöltése |
| `TryFillAllSheets` | Function | `FillAllSheets` hívása hibavédelemmel; visszaadja a hibát/figyelmeztetéseket |
| `WriteLogFile` | Function | Naplófájl írása a szülőmappába |
| `Warn` | Sub | Figyelmeztető üzenet hozzáfűzése a `g_warnings`-hoz |
| `FindFile` | Function | Első illeszkedő fájlnév visszaadása (temp fájlok kizárva) |
| `FindFileEndsWith` | Function | Első `*.xls*` fájl, amelynek alapneve adott suffixre végződik |
| `FindUtolagosFile` | Function | Utólagos forrásfájl megkeresése (nem `p_*`, periódusra végződik) |
