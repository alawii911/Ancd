# PUP1 – DRE-import

Två sätt att fylla PUP1 från en DRE-rapport. Båda gör samma sak som Power Query-frågorna i PUP1.

## 1. Makro i PUP1 (rekommenderas) – `PUP_Import.bas`

Välj DRE-fil, bekräfta datum och tid, så fylls Data-bladet och PUP uppdateras.

Installation (en gång). PUP1:s VBA-projekt är lösenordsskyddat, så makrot läggs i en egen fil:
1. Skapa en ny tom arbetsbok och spara den som **PUP_Import.xlsm** (Excel-arbetsbok med makron).
2. `Alt + F11` → **File → Import File…** → välj `PUP_Import.bas`.

Varje gång:
1. Öppna både PUP_Import.xlsm och PUP1.xlsm.
2. Klicka på PUP1-fönstret så att den är aktiv, tryck `Alt + F8` och kör `PUP_Import.xlsm!ImporteraDRE`.

Makrot:
- tömmer och fyller Table1 (A–G) med exakt så många rader som behövs
- sätter Track Date och Track Time i Table2 (J–O); övriga fält där lämnas som de är
- kör Refresh All så att PUP-bladet byggs av de befintliga frågorna
- fyller ut Track Time till fyra tecken (frågan gör annars 0905 till 905)
- säger till om några rader inte får någon Base Service

- varnar om Track Date ligger 7 dagar eller mer från idag (valideringen ger FAIL på det)
- frågar efter Dropbox Number när Stop Type är R eller D och skriver det i kolumn T på PUP

## Varför rader blir FAIL i Validate Scans

Reglerna ligger i PUP1:s egen kod (modulen FWSMain):
- Stop Type R eller D kräver ett numeriskt Dropbox Number (1–7 siffror) i kolumn T.
- Track Date får vara högst 6 dagar från idag, åt båda håll.
- Base Service måste vara 01, 04, 06, 18, 70 eller 86 (EC, RL, RM, 2A, 2P godkänns inte).
- Pkg Type måste vara 01, 02, 03, 04, 06, 15 eller 25.
- Dest City (första bokstaven) och Dest Cntry får bara innehålla A–Z.
- Track Time: timme 01–24.
- Validate och Send hanterar bara rad 3–1000.

Siffror i DRE-filen skrivs som siffror (som vid kopiera–klistra), annars matchar inte
Base Service mot Service-tabellen.

## 2. Fristående `PUP.html`

Öppna i Chrome/Edge, dra in DRE-filen och kopiera tabellen (PUP-bladets kolumn A–M) till A3.
Uppdaterar inte PUP-bladet själv – en webbsida kan inte skriva i en öppen Excel-fil.

Varken makrot eller HTML-sidan skickar scans eller hanterar inloggning.
