# PUP1 – import av DRE-export med en knapp

Makrot `PUP_Import.bas` ersätter steg 3–5 i instruktionerna på bladet *Data*:

1. Du väljer DRE-exporten (xlsx/xls/csv).
2. Kolumnerna Tracking Number, Form ID, Dest City Name, Dest Cntry,
   Dest Postal Cd, Base Service och Pkg Type hittas via rubrikerna, så ordningen
   i exporten spelar ingen roll.
3. Tabellen A–G töms och fylls med exakt så många rader som behövs (inga extra
   rader att ta bort).
4. Du får bekräfta Track Date (MMdd) och Track Time (HHmm). Förifyllt med nu.
   LocID, Pickup, Employee number och Stop Type ligger kvar som de står.
5. Allt uppdateras (Refresh All) och bladet *PUP* visas.

## Installation (en gång)

1. Öppna PUP1 i Excel för skrivbordet.
2. `Alt + F11` → **File → Import File…** → välj `PUP_Import.bas`.
3. Stäng VBA-fönstret. **Arkiv → Spara som** → *Excel-arbetsbok med makron (\*.xlsm)*.
4. Lägg till en knapp: **Infoga → Former** → rita en rektangel på bladet *Data*
   → högerklicka → **Tilldela makro…** → `ImporteraDRE`.

## Varje gång

1. Kör rapporten i DRE och spara/exportera den som Excel.
2. Klicka på knappen i PUP1 → välj filen → bekräfta datum/tid. Klart.

## Om något strular

- **"Följande kolumner saknas"** – rubriken heter något annat i exporten. Lägg
  till namnet i `IsAlias` i makrot (versaler, mellanslag och punkter ignoreras).
- **Makron blockerade** – vissa företagsdatorer stänger av makron. Då går det
  inte att lösa med VBA; säg till så görs en variant utan makron.
