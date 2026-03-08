# szamla-pre-v1.vb — Leírás

## Mire való?

Ez a script a `szamla-v107.vb` előtt fut. Feladata: minden kiválasztott almappában megtalálja a sablonfájlt (`pipalo_v20250827_tol_sablonom.xlsx`), és feltölti adatokkal az adott almappában lévő forrásfájlokból.

A script a `Personal.xlsb`-be kerül, és a `Run_PreProcess_Batch` nevű makróval indítható.

---

## Futtatás menete

1. **Periódus megadása** — két InputBox: 2 jegyű év (pl. `26`), majd hónap (1–12). Ebből lesz a perióduskód, pl. `v2602`.
2. **Szülőmappa kiválasztása** — mappapicker dialógus. Az almappáit fogja végigpásztázni.
3. **Almappák szűrése** — csak azok az almappák kerülnek feldolgozásra, amelyek tartalmazzák a sablonfájlt.
4. **Megerősítés** — megmutatja: hány almappa, melyik periódus, és a `regi_tsz` szűrésnél melyik 4 YYMM kódot keresi az S oszlopban.
5. **Batch feldolgozás** — almappánként:
   - Megnyitja a sablonfájlt
   - Feltölti a 7 munkalapot forrásfájlokból
   - Menti és zárja a fájlt
   - Naplózza az eredményt: `[OK]`, `[WARN]` (ha valamely forrásfájl hiányzott), vagy `[ERROR]` (futási hiba — a fájl nem kerül mentésre)
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
- **tsz és regi_tsz forrásfájlok**: ha a fájlban rejtett lap van az 1. pozícióban (pl. HBCS1608), a script az első **látható** lapot használja (`GetFirstVisibleSheet`).
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
```

- `[OK]` — minden forrásfájl megtalálva, feltöltés sikeres
- `[WARN]` — egy vagy több forrásfájl hiányzott, de ami megvolt, az feltöltődött; a fájl mentésre kerül
- `[ERROR]` — futási hiba történt; a fájl **nem** kerül mentésre

---

## Technikai megjegyzések

- A script `Personal.xlsb`-be kerül, `SzamlaPreProcess` modul névvel.
- Batch futás közben `ScreenUpdating`, `Calculation`, `EnableEvents`, `DisplayAlerts` ki van kapcsolva a gyorsabb futásért — mentés előtt minden visszaállítódik.
- Egymásba ágyazott `Dir()` helyett `FileSystemObject`-et használ az almappák iterálásához.
- Forrásworkbookok `ReadOnly:=True` módban nyílnak meg.
