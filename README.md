# Partijcontrole Veal

Klein Windows-programma om een partij vleeskalveren te controleren bij SKV.
Je vult **één UBN** (Uniek Bedrijfsnummer) en **1 tot 5 AM91-nummers** in. Het programma haalt dan twee dingen op:

1. **SKV-certificaat** (SKV API): is het kwaliteitscertificaat van de combinatie UBN + AM91 geldig?
2. **VKI** (Voedselketeninformatie, SKV InfoKalf-webservice): het VKI-formulier van de kalverhouder, met de antwoorden op de 20 vragen, de diergeneesmiddelen en de diernummers.

- Per regel: groen ✓ gelukt, oranje ! let op, rood ! fout. Dat staat er los voor het certificaat en voor de VKI.
- Dubbelklik op een regel voor alle details:
  - tabblad **SKV-certificaat**: het volledige JSON-antwoord, en knoppen voor de certificaat- en stallijst-PDF;
  - tabblad **VKI**: de gegevens, de 20 vragen, de diergeneesmiddelen, en knoppen voor het **VKI-formulier (PDF)** en het **VKI-antwoord (XML)**.
- Exporteren van het overzicht naar een `.txt`-bestand.

Geschreven in Windows PowerShell 5.1 met Windows Forms. Er hoeft niets
geïnstalleerd te worden. De VKI-PDF wordt gemaakt met Microsoft Edge, dat standaard op Windows 10/11 staat.

> Voorheen heette dit programma **SKV Certificaat Checker** (`SKV-Checker.ps1`).

## Installeren

1. Download de code: groene knop **Code** → **Download ZIP**.
2. Klik met de rechtermuisknop op de zip → **Eigenschappen** → vink **Blokkering opheffen** aan (als die optie er staat) → **OK**.
3. Pak de zip uit naar een map die alleen van jou is, bijvoorbeeld `C:\Tools\Partijcontrole-Veal`.
4. Kopieer `appsettings.example.json` naar `appsettings.json` en vul daarin de URL's, je SKV-gebruikersnaam en -wachtwoord, en bij `Vki` je VKI-gebruikersnaam en -wachtwoord in.
   (Vergeet je dit, dan maakt het programma bij de eerste start zelf een `appsettings.json` aan en vraagt het je die in te vullen.)
5. Dubbelklik op **Start Partijcontrole Veal.bat**.

**Had je de SKV Checker al?** Je kunt je oude `appsettings.json` gewoon blijven gebruiken. Voeg alleen het blok `"Vki": { ... }`
uit `appsettings.example.json` toe. Zonder dat blok werkt het programma zoals vroeger: alleen de certificaatcontrole. Zie [LEESMIJ.txt](LEESMIJ.txt).

## Bestanden

| Bestand | Doel |
|---|---|
| `Start Partijcontrole Veal.bat` | Start het programma |
| `Start Partijcontrole Veal (met foutmeldingen).bat` | Start het programma en laat foutmeldingen zien |
| `Partijcontrole-Veal.ps1` | Het programma zelf |
| `appsettings.example.json` | Voorbeeld van de instellingen (zonder wachtwoorden) |
| `appsettings.json` | Jouw eigen instellingen (staat niet op GitHub) |
| `test\Test-Offline.ps1` | Test van de VKI-verwerking zonder netwerk |
| `test\*.xml` | Testantwoorden. Bestanden met `nagemaakt` in de naam zijn zelf gemaakt, de rest is een geanonimiseerd of echt foutantwoord |

## Instellingen (`appsettings.json`)

| Veld | Uitleg |
|---|---|
| `TokenUrl` | Adres waar het bearer-token wordt opgehaald |
| `CertificaatUrl` | Adres van de certificaatcontrole |
| `Username` / `Password` | Je SKV-inloggegevens (certificaat) |
| `TokenMethod` | `GET`, `POST` of `AUTO` (eerst POST, dan GET) |
| `TimeoutSeconds` | Maximaal wachten op SKV, in seconden |
| `Vki.Ingeschakeld` | `true` = ook de VKI ophalen, `false` = alleen het certificaat |
| `Vki.Url` | Adres van de VKI-webservice (productie: `https://infokalf.skv.info/service/VKIService/V2/VKIwebservice.svc`) |
| `Vki.Username` / `Vki.Password` | Je VKI-inloggegevens (SKV InfoKalf) |
| `Vki.TimeoutSeconden` | Maximaal wachten op de VKI-webservice, in seconden |

Benodigde URL's, gebruikersnamen en wachtwoorden zijn op te vragen bij de auteur. Deze gegevens zijn aan een bedrijf gebonden!

**Let op:** SKV blokkeert het VKI-account 10 minuten na 3 foute inlogpogingen. Het programma probeert het daarom na een
inlogfout (`ERR0002`) niet opnieuw tot je het opnieuw start.

## Hoe de VKI wordt beoordeeld (voorlopig)

| Kleur | Wanneer |
|---|---|
| Rood ! | geen formulier gevonden / fout van SKV / niet "naar waarheid ingevuld" / vraag 1 niet JA / een dier op het formulier zonder AM91-nummer |
| Oranje ! | een van de vragen 2 t/m 9 is JA / er staan diergeneesmiddelen op het formulier / er zijn diagnoses |
| Groen ✓ | al het andere |

**Zichtkeuring: ja** als vraag 10 t/m 13 JA zijn, 14 t/m 16 NEE, en (als vraag 17 JA is) ook 18 en 20 JA en 19 JA of NVT.

Deze regels zijn **voorlopig** en moeten nog door de key-user bevestigd worden. De regel voor diergeneesmiddelen
staat op één plek in de code (`Get-VkiMiddelenStatus`), zodat die later makkelijk aan te passen is.

De VKI-PDF wordt door dit programma zelf samengesteld uit de VKI-gegevens, met dezelfde opbouw als de PDF uit SKV InfoKalf.
Onderaan staat dat het geen officiële SKV-print is.

## Testen zonder netwerk

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\Test-Offline.ps1
# met HTML/PDF van de testformulieren in een map:
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\Test-Offline.ps1 -HtmlMap C:\Temp\vki-test -Pdf
```
