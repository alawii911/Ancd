# PUP.html

Fristående ersättare för att fylla bladen *Data* och *PUP* i PUP1.xlsm. Öppna `PUP.html`
i Chrome/Edge (dubbelklick), dra in DRE-rapporten och kopiera resultatet.

Den gör samma sak som Power Query-frågorna i PUP1:

| PUP-kolumn | Källa |
| --- | --- |
| AWB Number | Tracking Number |
| Form ID | Form ID, fylls ut med nollor till 4 tecken |
| Track Location ID, Pickup Route Number, Emploee number, Stop Type | Inställningar på sidan |
| Track Date (MMdd), Track Time (HHmm) | Inställningar på sidan (förifyllda med nu), 4 tecken |
| Dest City | Första bokstaven i Dest City Name |
| Dest Cntry, Dest Postal Cd | Direkt från DRE |
| Base Service | Slås upp i Service-tabellen (1→01, 4→04, 6→06, 18, 70, 86, EC, RL, RM, 2A, 2P) |
| Package type | "0" + Pkg Type |

Raderna sorteras fallande på AWB, som i PUP-bladet.

Sidan skickar inga scans och hanterar inga inloggningsuppgifter. Den förbereder bara tabellen.
