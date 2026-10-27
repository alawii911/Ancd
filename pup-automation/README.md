# PUP1 – DRE-import

Två sätt att fylla PUP1 från en DRE-rapport. Båda gör samma sak som Power Query-frågorna i PUP1.

## 1. Makro i PUP1 (rekommenderas) – `PUP_Import.bas`

Knapp i PUP1-filen: välj DRE-fil, bekräfta datum och tid, så fylls Data-bladet och PUP uppdateras.

Installation (en gång, i den riktiga PUP1.xlsm):
1. `Alt + F11` → **File → Import File…** → välj `PUP_Import.bas`.
2. Lägg en knapp på bladet Data eller PUP: **Developer → Insert → Button (Form Control)** → tilldela makrot `ImporteraDRE`.
   (Eller kör det med `Alt + F8`.)

Makrot:
- tömmer och fyller Table1 (A–G) med exakt så många rader som behövs
- sätter Track Date och Track Time i Table2 (J–O); övriga fält där lämnas som de är
- kör Refresh All så att PUP-bladet byggs av de befintliga frågorna
- fyller ut Track Time till fyra tecken (frågan gör annars 0905 till 905)
- säger till om några rader inte får någon Base Service

Siffror i DRE-filen skrivs som siffror (som vid kopiera–klistra), annars matchar inte
Base Service mot Service-tabellen.

## 2. Fristående `PUP.html`

Öppna i Chrome/Edge, dra in DRE-filen och kopiera tabellen (PUP-bladets kolumn A–M) till A3.
Uppdaterar inte PUP-bladet själv – en webbsida kan inte skriva i en öppen Excel-fil.

Varken makrot eller HTML-sidan skickar scans eller hanterar inloggning.
