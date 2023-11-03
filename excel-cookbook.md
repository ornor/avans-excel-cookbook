# Excel cookbook

*versie 03-11-2023*

Avans Hogeschool, Academie voor Duurzaam Gebouwde Omgeving, cursus Parametrisch Ontwerpen (DG-MI-PAO).

----



[TOC]

## Celverwijzingen vastzetten

Het is erg handig om formules te kunnen kopiëren/hergebruiken. Hiertoe is het vaak efficiënt *dollartekens* `$` [&Eopf;](https://support.microsoft.com/nl-nl/office/formules-verplaatsen-of-kopi%C3%ABren-1f5cf825-9b07-41b1-8719-bf88b07450c6) te gebruiken om bepaald rijen en/of kolommen vast te zetten bij het kopiëren en plakken. Normaal gesproken wordt een verwijzing automatisch aangepast op de doelcel. Bij gebruik van dollartekens blijft de verwijzing hetzelfde als in de cel waar die is gekopieerd.

In onderstaande voorbeeld is de formule in cel C5 gekopieerd en geplakt in de overige blauwe cellen. Cel B2 moet overal gebruikt worden. Bij het plakken één cel rechts van C5 moet deze verwijzing naar B2 NIET worden aangepast naar C2 (dit zou anders gebeuren). Door gebruik van dollartekens zal ALTIJD de nieuwe cel verwijzen naar cel B2. 

In de blauwe cel zijn zowel de rij als kolom vastgezet. In de rode cel is alleen de rij 4 vastgezet (de kolommen moeten juist WEL automatisch aangepast worden om de juiste kolomletters te krijgen). En de paarse cel is juist alleen de kolom B vastgezet (en de rijen niet).

![image-20230925121248932](./assets/image-20230925121248932.png)

## Werken met namen en gegevensvalidatie

Er kan in een formule worden verwijzen naar andere cellen d.m.v. het adres (bijvoorbeeld `$A$2`). Het is ook mogelijk om een bepaalde cel, of een bereik meerdere aaneengesloten cellen een *naam* te geven [&Eopf;](https://support.microsoft.com/nl-nl/office/namen-defini%C3%ABren-en-gebruiken-in-formules-4d0f13ac-53b7-422e-afd2-abd7ff379c64).

![image-20231102161344086](./assets/image-20231102161344086.png)

In bovenstaande voorbeeld wordt de inhoud van een balk uitgerekend. Cel C2 krijgt nu de naam `b` door deze te selecteren en vervolgens in het witte invoervakje links naast de formule werkbalk, de nieuwe naam te typen (en op `ENTER` te drukken).

![image-20231102161622009](./assets/image-20231102161622009.png)

Als je nu de nieuwe formule (met namen) in cel C6 gaat typen, dan zie je ingevoerde naam ook in de keuzelijst terug komen (tijdens het typen). 

Het voordeel van het geven van een naam, is dat het niet uitmaakt waar de breedte-waarde staat. Zolang het bereik van de naam `b` maar naar juiste invoerveldje gaat, komt het goed. Mocht je later de formules willen veranderen, omdat je een anders invoerveld wilt gebruiken voor de breedte-waarde, dan hoef je NIET alle formules aan te passen (want naam klopt dan nog steeds), maar je hoeft alleen het bereik van de naam aan te passen. Dit kan je doen via het tabblad "Formules" bij de knop "Namen beheren". Selecteer de naam en pas het bereik aan via de knop "Bewerken".

![image-20231102162100683](./assets/image-20231102162100683.png)

Een ander voordeel van namen geven is dat de formules leesbaarder kunnen worden.

Een nadeel van namen geven is, dat als je alle cellen op werkblad wilt kopiëren (om voor een tweede keer te gebruiken), je maar één waarde voor één naam kan hebben. Je moet dus dan het gekopieerde nieuwe (unieke) waarden geven.

Namen zijn erg handig in combinatie met *gegevensvalidatie* [&Eopf;](https://support.microsoft.com/nl-nl/office/gegevensvalidatie-toepassen-op-cellen-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249). 

Stel je hebt een lijst met namen van personen. En je wilt dat je in een andere cel (eventueel op een ander tabblad) in een cel kan kiezen tussen deze personen (in onderstaande situatie is dat cel F8). Dan is het handig om een keuzelijst te hebben (om te zorgen dat je nooit een niet geldige naam kan typen). Om dat te doen, ga je naar tabblad "Gegevens" naar de knop "Gegevensvalidatie". 

![image-20231102163005981](./assets/image-20231102163005981.png)

Je krijgt dan een popup menu met de volgende opties. Bij het vakje "toestaan" kan je restricties meegeven welke waarden gebruikt mogen worden. In dit geval kies je voor je voor "lijst" (i.p.v. "alle waarden"). Je kan dan naast het vakje van "Bron" klikken op het pijltje omhoog om de cellen F3:F6 te selecteren. Je hebt nu een keuzelijst gemaakt.

![image-20231102163222779](./assets/image-20231102163222779.png)![image-20231102163359507](./assets/image-20231102163359507.png)

Een nadeel is dat als de lijst met later nog wordt aangepast, je ook alle gegevensvalidaties moet gaan aanpassen. Ook is het een nadeel dat je geen cellen kan selecteren van andere tabbladen (bij de invoeroptie "Bron"). Beiden kunnen opgelost worden door de invoercellen een naam te geven. Vervolgens type je bij "Bron" een `=` met daarachter de naam die je aan invoerbereik hebt gegeven.

![image-20231102162707285](./assets/image-20231102162707285.png) ![image-20231102163730260](./assets/image-20231102163730260.png)

Een andere interessante functie bij gegevensvalidatie, is het toevoegen van helptekst als een gebruiker een bepaalde cel selecteert. In dit geval kunnen we hierboven bij het tabblad "Invoerbericht" dit bericht nog invoeren.

![image-20231102164051290](./assets/image-20231102164051290.png) ![image-20231102164123604](./assets/image-20231102164123604.png)

## Waarde zoeken in tabel

Functies als `HORIZ.ZOEKEN` en `VERT.ZOEKEN` worden veel gebruikt om waarden in een datatabel op te zoeken. Echter bij deze functie moet een totaal bereik van hele tabel worden opgegeven. Ook moet een rij- of kolom index getal worden aangegeven. Als datakolommen of -rijen worden toegevoegd of aangepast, dan is het niet altijd duidelijk of de zoekfuncties nog juist zijn.

Een andere manier om waarden op te zoeken is de combinatie van de functies `INDEX` [&Eopf;](https://support.microsoft.com/nl-nl/office/index-functie-a5dcf0dd-996d-40a4-a822-b56b061328bd) en `VERGELIJKEN` [&Eopf;](https://support.microsoft.com/nl-nl/office/vergelijken-functie-e8dffd45-c762-47d6-bf89-533f4a37673a).

![image-20230922114521125](./assets/image-20230922114521125.png)

In dit geval kan in een geel vakje een naam worden ingegeven (met behulp van keuzemenu door 'lijst' in *gegevensvalidatie* [&Eopf;](https://support.microsoft.com/nl-nl/office/een-vervolgkeuzelijst-maken-7693307a-59ef-400a-b769-c5402dce407b)) en wordt de bijhorende woonplaats opgezocht. De functie `vergelijken` zoekt de hoeveelste cel de naam (in gele vakje) in de deelnemerslijst is. De functie `index` haalt de zoveelste waarde uit het bereik van de woonplaatsen.

Gebruik bij de functie `vergelijken` altijd de '0' als 3e parameter om betrouwbare resultaten te krijgen.

Je ziet nu heel duidelijk van welke cellen deze formule afhankelijk is (dat minder het geval is bij `VERT.ZOEKEN`).

Je kan het bereik van cellen nog verbeteren met gebruik van *dollar-tekens* om celrichtingen vast te zetten bij kopiëren [&Eopf;](https://support.microsoft.com/nl-nl/office/formules-verplaatsen-of-kopi%C3%ABren-1f5cf825-9b07-41b1-8719-bf88b07450c6). Ook kan gekozen worden om bereik een *naam* te geven [&Eopf;](https://support.microsoft.com/nl-nl/office/namen-defini%C3%ABren-en-gebruiken-in-formules-4d0f13ac-53b7-422e-afd2-abd7ff379c64).

## Functies met 'ALS'

Excel heeft een aantal functies met 'ALS' daarin zoals: `SOM.ALS`, `PRODUCT.ALS` en `AANTAL.ALS`. Deze functies zijn erg handig om twee stappen in één keer te doen (je hebt dan geen hulpcellen nodig als tussenuitkomst om iets uit te rekenen).

![image-20230922123525428](./assets/image-20230922123525428.png)

In dit geval wordt alle blauwe cellen opgeteld (SOM) wanneer deze groter of gelijk zijn aan 5,5 (in dit geval 7,0 + 6,9). Deze wordt gedeeld door het aantal voldoendes (in dit geval 2). Dit geeft het gemiddelde van alleen de voldoendes.

Deze formule had trouwens nog korter kunnen worden opgesteld door gebruik te maken van de functie `GEMIDDELDE.ALS`.

## Berekenen van gewogen gemiddelde

Soms wil je niet een normaal gemiddelde berekenen, maar een gemiddelde gewogen naar en bepaalde weegfactor.

![image-20230922145429614](./assets/image-20230922145429614.png)

De functie `SOMPRODUCT` [&Eopf;](https://support.microsoft.com/nl-nl/office/somproduct-functie-16753e75-9f68-4874-94ac-4d2145a2fd2e) gaat per element de ene lijst vermenigvuldigen met de andere lijst, om vervolgens deze resultaten op te tellen. In dit geval: 4 * 3 + 7 * 1 + 8 * 2. Het gewogen gemiddelde is in dit voorbeeld een 5,83.

## Celverwijzingen ontkoppelen of relatief maken

De functie `INDIRECT` [&Eopf;](https://support.microsoft.com/nl-nl/office/indirect-functie-474b3a3a-8a26-4f44-b491-92b6306fa261) kan een waarde van een andere cel verkrijgen zonder dat deze gekoppeld is. Als in onderstaande voorbeeld de cel C3 wordt verplaats (d.m.v. knippen en plakken), zal de verwijzing in cel F3 dan ook NIET mee veranderen (wat normaal wel het geval is).  Normaal is dat ongewenst maar er zijn situaties waarin dat juist gewenst kan zijn.

![image-20230922125415160](./assets/image-20230922125415160.png)

Stel de data in kolommen B t/m D is uit een externe bron gekopieerd en hier geplakt. Vervolgens wil je hier nieuwe data overheen plakken (omdat deze is veranderd). Het kan dan zijn dat je per ongeluk de oorspronkelijke data eerst verwijderd en hiermee ook alle koppelingen verwijderd. Vervolgens plak je de nieuwe data maar alle verwijzingen kloppen niet meer. Om dit probleem te tackelen kan je gebruik maken van de functie `INDIRECT`. Als je dan nieuwe data erin plakt, weet je zeker dat de cel F3 de juiste waarde ophaalt aangezien deze niet afhankelijk was van koppelingen. Wel is het belangrijk dat je de data structuur (bijvoorbeeld aantal kolommen) niet aanpast, want anders klopt de formule niet meer.

Je kan ook gebruik maken van verwijzingen naar andere cellen om indirect (relatief) een koppeling te maken met een cel. Een voorbeeld is hieronder gegeven.

![image-20230922130424403](./assets/image-20230922130424403.png)

De functie `ADRES` [&Eopf;](https://support.microsoft.com/nl-nl/office/adres-functie-d0c26c0d-3991-446b-8de4-ab46431d4f89) geeft een tekst-celverwijzing gebaseerd op een rij- en kolomnummer. Hiertoe wordt nu de functie `RIJ` [&Eopf;](https://support.microsoft.com/nl-nl/office/rij-functie-3a63b74a-c4d0-4093-b49a-e76eb49a6d8d) gebruikt om vanuit de blauwe verwijzing een rijnummer (3) te verkrijgen en met de functie `KOLOM` [&Eopf;](https://support.microsoft.com/nl-nl/office/kolom-functie-44e8c754-711c-4df3-9da4-47a55042554b) een kolomnummer (2 (= B)) van rode verwijzing te verkrijgen. De leeftijd van Jan is nu niet direct gekoppeld aan de cel waar deze waarde staat (deze kan dus geknipt/verwijderd worden zonder dat er foutmeldingen ontstaan). Mocht er een kolom tussen A en B gevoegd worden, zal het (in tegenstelling tot 1e voorbeeld) ook goed blijven gaan omdat deze de juist kolomnummer vanuit rode cel mee krijgt.

Je kan ook een verwijzing maken naar bijvoorbeeld een cel rechts van een andere kolom, door het kolomnummer te verhogen met 1.

![image-20230922132239693](./assets/image-20230922132239693.png)

Of je kan aangeven dat je het laatste kolomnummer van een bepaald bereik wilt hebben.

![image-20230922133110065](./assets/image-20230922133110065.png)

## Werken met tekst

In Excel formules kan gewerkt worden met tekst. Alles tussen dubbele aanhalingstekens is tekst. Zie volgende overzichtspagina met alle tekst functies: [&Eopf;](https://support.microsoft.com/nl-nl/office/tekstfuncties-overzicht-cccd86ad-547d-4ea9-a065-7bb697c2a56e)

Tekst kan worden gecombineerd (aan elkaar geplakt) door middel van de operator: `&` [&Eopf;](https://support.microsoft.com/nl-nl/office/tekst-samenvoegen-functie-8f8ae884-2ca8-4f7a-b093-75d702bea31d).

![image-20230922134751237](./assets/image-20230922134751237.png)

Een getal (of andere niet-tekstuele waarde) kan worden omgezet naar tekst met de functie `TEKST` [&Eopf;](https://support.microsoft.com/nl-nl/office/tekst-functie-20d5ac4d-7b94-49fd-bb38-93d29371225c). Hierbij moet worden aangegeven in welk format de waarde moet worden weergegeven. Voorbeelden van dit format zijn ook te vinden bij *celeigenschappen* (ctrl+1), tabblad 'getal', optie 'aangepast'.

![image-20230922135317453](./assets/image-20230922135317453.png)

In bovenstaande voorbeeld berekent de functie `RIJEN` het aantal rijen in het blauwe bereik. Het argument "0" geeft aan dat dit als een geheel getal moet worden weergegeven.

De functie `WAARDE` [&Eopf;](https://support.microsoft.com/nl-nl/office/waarde-functie-257d0108-07dc-437d-ae1c-bc2d3953d8c2) doet het omgekeerde en zet een stuk tekst om naar een getal. Onderstaande voorbeeld geeft het getal 5.

![image-20230922140343442](./assets/image-20230922140343442.png)

Er kan ook geknipt worden in een stuk tekst. Zodat alleen een deel van tekst gebruikt word. Zie functies `DEEL` [&Eopf;](https://support.microsoft.com/nl-nl/office/deel-deelb-functie-d5f9e25c-d7d6-472e-b568-4ecb12433028), `LINKS` [&Eopf;](https://support.microsoft.com/nl-nl/office/links-linksb-functie-9203d2d2-7960-479b-84c6-1ea52b99640c) en `RECHTS` [&Eopf;](https://support.microsoft.com/nl-nl/office/links-linksb-functie-9203d2d2-7960-479b-84c6-1ea52b99640c). De functie `LENGTE` [&Eopf;](https://support.microsoft.com/nl-nl/office/lengte-lengteb-functie-29236f94-cedc-429d-affd-b5e33d2c67cb) retourneert het aantal karakters.

![image-20230922141108466](./assets/image-20230922141108466.png)

![image-20230922141518942](./assets/image-20230922141518942.png)

Laatste formule zoekt een deel van het woord "Achmed" op, beginnend bij letter 3. Het aantal resterende letters is het aantal totale letters van dit woord minus 2 (de eerste twee letters van het woord). Dit resulteert in: "Daarna volgen de letters: hmed."

## Werkbladbeveiliging

Soms is het handig om te zorgen dat gebruikers bepaalde cellen niet per ongeluk aan kunnen passen. In onderstaande voorbeeld mag een gebruiker cel C2 aanpassen maar C4 niet. Het is mogelijk om een werkblad te beveiligen. Hierdoor worden restricties opgelegd aan de gebruiker. 

Standaard zijn alle cellen `geblokkeerd`. Dat kan je zien/aanpassen als je rechter muis klikt op die cel en dan naar "Celeigenschappen" gaat, bij laatste tabblad "Bescherming". In dit geval willen we in cel C2 het vakje weer uitvinken (zodat deze cel niet wordt geblokkeerd).

![image-20231102164715759](./assets/image-20231102164715759.png)

Rechter muis klik vervolgens op het tabblad onder (het werkblad) die je wilt beveiligen en klik op "Blad beveiligen". 

![image-20231102165215758](./assets/image-20231102165215758.png) 

Je ziet vervolgens onderstaande menu. Hier kan je in de meeste situaties, gewoon de standaard instellingen laten staan. Deze instellingen laten het toe om alle cellen te selecteren (geblokkeerd en niet geblokkeerde) om zodoende de formules ook te kunnen zien. Het is NIET toegestaan om opmaak aan te passen. 

![image-20231102165345478](./assets/image-20231102165345478.png)

Je kan hierna alleen de waardes in de cellen die NIET geblokkeerd zijn veranderen. Als je een geblokkeerde cel probeert te veranderen dan krijg je de onderstaande waarschuwing.

![image-20231102165401839](./assets/image-20231102165401839.png)

Tip: laat het invoerveld voor een wachtwoord altijd leeg. Zodoende kan iedereen het werkblad ook weer aanpassen. Zeker als collega's vertrekken bij bedrijf, kan het zijn dat bepaalde wachtwoorden niet meer beschikbaar zijn. Dan kunnen de rekensheets niet meer aangepast worden. Je kan de beveiliging er weer afhalen door weer rechter muis knop op tabblad onderaan en kies voor "Beveiliging blad opheffen".

## ==Voorwaardelijke opmaak==



## ==Formules evalueren==

bron- en doelcellen



## ==Werken met foutmeldingen==

ISFOUT, ISLEEG, ISNB, ISGETAL etc.



## Matrix formules

Als je een formule wilt repeteren, dan kan je de formule kopiëren en hergebruiken door te plakken. Wat ook mogelijk is, is om maar één keer deze formule te typen en deze op een aantal cellen tegelijk van toepassing te laten zijn. In onderstaande voorbeeld zijn de cellen F2 t/m F5 geselecteerd, vervolgens is de gegeven formule getypt, en is deze afgesloten door `CTRL + SHIFT + ENTER` te typen. Hierbij is een matrix formule gecreëerd [&Eopf;](https://support.microsoft.com/nl-nl/office/matrixformules-richtlijnen-en-voorbeelden-7d94a64e-3ff3-4686-9372-ecfd5caa57c7). 

![image-20230922144136615](./assets/image-20230922144136615.png)

Je kan ook zien dat het een matrix functie is, door op één van deze vier afzonderlijke cellen te klikken. Je ziet de formule dan tussen accolades `{ }` staan.

![image-20230922144632625](./assets/image-20230922144632625.png)

Het idee is dat je de matrix formule niet in één van de afzonderlijke cellen kan aanpassen. Want deze vier cellen delen namen een gemeenschappelijke formule. Je kan deze alleen aanpassen door al deze vier de cellen te selecteren en vervolgens de formule te wijzigen.

Het voordeel van deze aanpak is dat je een formule maar altijd op één plek hebt staan. Je kan nooit vergeten een formule ook aan te passen bij de andere soortgelijke cellen.

Matrix formules kunnen ook gebruikt worden voor een stelsel vergelijkingen met aantal onbekenden op te lossen.

![image-20230922143453320](./assets/image-20230922143453320.png)

![image-20230922143520796](./assets/image-20230922143520796.png)

In dit geval blijkt bijvoorbeeld dat de volgende vergelijking klopt:   -9 * ==-6.8== + 2 * ==143,5== + 3 * ==-114,6== = 4

## ==Tekeningen maken==

 grafiek type 'spreiding'



## ==Werken met macro's==



## ==MVC model toegepast op Excel==



### Model

adsf

### View

1D leesrichting (in niet 2D, zo werken onze hersenen niet)

waarde of formule?; nooit leeg veld achter laten

### Controller

asdf
