# VESC tagdíjak

Tagdíj-nyilvántartás a VESC triatlon klubnak. A felület GitHub Pages-en fut, az adatok egy Google Sheetben vannak (Apps Script webalkalmazáson keresztül), ugyanúgy, mint a közös edzésnaptárnál.

**Tud:** új befizetés rögzítése űrlapon, tagonkénti fennmaradó összeg és időarányos tartozás, havi rács (mint az Excelben), tagok és éves csomagok kezelése, CSV-export.

## Fájlok

| Fájl | Mire való |
|---|---|
| `index.html` | A weboldal (GitHub Pages) |
| `Code.gs` | A Google Sheet mögötti Apps Script (nem kell a repóba, de itt a helye a verziókövetésnek) |

## Beállítás (egyszer)

### 1. Google Sheet + Apps Script (a vesctriatlon@gmail.com fiókban)
1. Hozz létre egy új Google Sheetet, például *VESC tagdíjak* néven.
2. **Bővítmények → Apps Script**. Töröld a mintakódot, és másold be a `Code.gs` teljes tartalmát.
3. A fájl elején írd át a `KEZDO_JELSZO` értékét a klub jelszavára, majd mentsd.
4. A függvényválasztóban válaszd a `setup` függvényt, és kattints a **Futtatás** gombra. Engedélyezd a hozzáférést.
   Ekkor létrejön a három lap (*Csomagok*, *Tagok*, *Befizetések*), betöltődnek a 2026-os adatok, és elmentődik a jelszó.
5. **Telepítés → Új telepítés → típus: Webalkalmazás**
   - Végrehajtás mint: **Én**
   - Hozzáférés: **Bárki**
6. Másold ki a kapott URL-t (`https://script.google.com/macros/s/…/exec`).

### 2. GitHub repó (Gombi75 fiók)
1. Új repó: **tagdij** (Public).
2. Töltsd fel az `index.html` és a `README.md` fájlt.
3. Az `index.html` elején az `API_URL` sorba illeszd be az előbb kimásolt URL-t:
   ```js
   const API_URL='https://script.google.com/macros/s/…/exec';
   ```
4. **Settings → Pages → Branch: main / (root) → Save**.
5. Pár perc múlva elérhető itt: `https://gombi75.github.io/tagdij/`

## Használat
- Belépéskor az oldal kéri a jelszót, és a böngésző megjegyzi. A **Kilépés** gombbal törölheted.
- A Sheetet közvetlenül is szerkesztheted (például régi adatok tömeges bemásolásához). Az oldal a **Frissítés** gombra vagy újbóli megnyitáskor tölti be a változásokat.
  - *Tagok* lap: ID, Név, Aktív (igen/nem), és évenként egy oszlop (`2026`, `2027`…) a csomag nevével.
  - *Befizetések* lap: a `TagID` oszlopnak a *Tagok* lap ID-jára kell mutatnia; a `Tól (hó)` és `Ig (hó)` oszlopban 1–12 közötti szám áll.
- **Új év:** a tagok csomagja automatikusan öröklődik az előző évből. A Tagok fülön évenként átállítható.
- **Jelszócsere:** az Apps Scriptben írd át a `jelszoCsere()` függvényben az új jelszót, és futtasd.
- **A Code.gs módosítása után:** Telepítés → Telepítések kezelése → szerkesztés → *Új verzió*. Így az URL változatlan marad.

## Adatvédelem
A repó nyilvános, de a tagdíjadatok nem: azok a Google Sheetben vannak, és csak jelszóval kérhetők le. A jelszót ne írd be a repóba.
