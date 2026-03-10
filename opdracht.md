## Opdracht Beschrijving

Je gaat een interactief Excel-dashboard bouwen met behulp van **AI (ChatGPT, Claude, Copilot etc)**.

De AI genereert de VBA-code, jij voert deze in en test het resultaat.

**Eindresultaat:**
- Knop "Genereer Data" → vult kolom A en B met random data
- Knop "Maak Dropdown" → maakt selectiemenu in D2
- Knop "Bereken Verkoop" → berekent totaal in E2

---

## Voorbereiding (1 minuut)

### Wat je nodig hebt:
- Excel geopend
- Een blad genaamd "Sheet1"
- VBA-editor ingeschakeld (Alt + F11)
- ChatGPT, Claude of Copilot geopend in je browser
- Dit document met de prompts

### Layout:
```
Kolom A: Productnamen (header: "Product")
Kolom B: Verkoopbedragen (header: "Verkoop")
Kolom D: Dropdown voor productselectie
Kolom E: Totale verkoop per product
```

---

## Stap 1: Genereer Data (3 minuten)

### Wat je doet:
1. Open **ChatGPT, Claude of Copilot** in je browser
2. Kopieer de prompt hieronder
3. Plak de prompt in de AI
4. Kopieer de gegenereerde VBA-code
5. Plak deze in Excel VBA-editor (Alt + F11)

### PROMPT 1: Genereer Data

```
Ik ben beginner in VBA en wil een knop maken in Excel die:

1. In kolom A (rijen 2-11) willekeurig deze producten genereert: 
   "Product A", "Product B", "Product C"
2. In kolom B (rijen 2-11) willekeurig getallen tussen 100 en 1000 genereert
3. De header in A1 moet "Product" zijn en B1 moet "Verkoop" zijn

Geef me ALLEEN de VBA-code die ik in een Sub-procedure kan plakken.
Zorg dat de code begint met: Sub GenereerData()

Maak de code stap voor stap uit met commentaar zodat ik begrijp wat er gebeurt.
```

### Hoe je dit invoert:
1. Druk op **Alt + F11** om VBA-editor te openen
2. Dubbelklik op **Sheet1** in het linkerpaneel
3. Plak de gegenereerde code in het editorvenster
4. Druk op **Ctrl + S** om op te slaan

### Test:
- [ ] Ga terug naar Excel (Alt + F11)
- [ ] Maak een knop en koppel aan **GenereerData**
- [ ] Klik op de knop
- [ ] Kolom A en B moeten nu data bevatten

---

## Stap 2: Maak Dropdown (4 minuten)

### PROMPT 2: Dropdown Menu

```
Ik wil een dropdown-menu maken in Excel cel D2 die:

1. Alleen de unieke productnamen toont die in kolom A staan
2. De producten zijn: "Product A", "Product B", "Product C"
3. De dropdown mag maar 1 product tegelijk selecteren

Geef me de stappen om dit in Excel in te stellen in VBA.
Zorg dat je het stap-voor-stap uitlegt zodat een beginner het kan volgen.

De gegenereerde VBA-code moet:
- Beginnen met: Sub MaakDropdownUniekeProducten()
- Alle unieke waarden uit A2:A11 automatisch detecteren
- Een dropdown maken in D2
- Commentaar hebben zodat ik begrijp wat er gebeurt
```

### Hoe je dit invoert:
1. Kopieer de gegenereerde code van de AI
2. Plak in VBA-editor (onder de vorige code)
3. Druk op **Ctrl + S** om op te slaan

### Test:
- [ ] Maak een knop en koppel aan **MaakDropdownUniekeProducten**
- [ ] Klik op de knop
- [ ] Klik op cel D2
- [ ] Er moet een dropdown-pijltje verschijnen
- [ ] Klik erop en selecteer een product

---

## Stap 3: SOMIF Berekening (4 minuten)

### PROMPT 3: SOMIF in VBA

```
Ik wil een VBA-functie die:

1. Kijkt naar wat ik in cel D2 heb geselecteerd (de dropdown)
2. In kolom A zoekt naar dat product
3. Alle bijbehorende verkoopbedragen uit kolom B optelt
4. Het totaal in cel E2 zet

De data staat in:
- Kolom A (rijen 2-11): Productnamen
- Kolom B (rijen 2-11): Verkoopbedragen

Geef me ALLEEN de VBA-code voor een Sub-procedure genaamd: BerekenSomIf()

Zorg dat:
- De code begint met: Sub BerekenSomIf()
- Er commentaar bij staat zodat ik begrijp wat er gebeurt
- De code niet meer dan 15 regels is
- Het resultaat in E2 komt
```

### Hoe je dit invoert:
1. Kopieer de gegenereerde code van de AI
2. Plak in VBA-editor (onder de vorige code)
3. Druk op **Ctrl + S** om op te slaan

### Test:
- [ ] Selecteer een product in D2 (bijv. "Product A")
- [ ] Maak een knop en koppel aan **BerekenSomIf**
- [ ] Klik op de knop
- [ ] Cel E2 moet nu het totaal tonen
- [ ] Wissel van product en klik opnieuw → getal in E2 verandert

---

## Stap 4: Knop voor Berekening (3 minuten)

### PROMPT 4: Knop Maken

```
Ik heb al een VBA-functie genaamd BerekenSomIf(), deze staat op "Sheet1" in mijn Excel-bestand.
Nu wil ik een knop maken die deze functie aanroept.

Dit is wat ik wil:
1. Een knop maken in Excel (op het zichtbare blad)
2. De knop een duidelijke naam geven: "Bereken Verkoop"
3. De knop koppelen aan de bestaande Sub BerekenSomIf()
4. Wanneer ik op de knop klik, moet de BerekenSomIf()-functie uitvoeren

Geef me de VBA-code die automatisch deze knop aanmaakt en koppelt.

Zorg dat:
- De code begint met: Sub MaakKnopVoorBerekening()
- De knop de tekst "Bereken Verkoop" heeft
- De knop gekoppeld is aan BerekenSomIf
- Er commentaar bij staat
```

### Hoe je dit invoert:
1. Kopieer de gegenereerde code van de AI
2. Plak in VBA-editor (onder de vorige code)
3. Druk op **Ctrl + S** om op te slaan
4. Voer **MaakKnopVoorBerekening()** uit (F5 in VBA-editor)
5. Wijzig "Sheet1" naar "Blad1" wanneer de taal Nederlands is

        knop.OnAction = "Sheet1.BerekenSomIf"
        
### Test:
- [ ] De knop "Bereken Verkoop" moet nu zichtbaar zijn op je blad
- [ ] Klik erop en test of het werkt

---
## Antwoorden
    Sub GenereerData()
        ' Zorg dat we met het actieve werkblad werken (je kunt dit ook expliciet maken, bv. met Worksheets("Blad1"))
        Dim ws As Worksheet
        Set ws = ActiveSheet
        
        ' Array met mogelijke productnamen
        Dim producten As Variant
        producten = Array("Product A", "Product B", "Product C")
        
        ' Willekeurige generator initialiseren (zodat de uitkomsten niet elke keer hetzelfde zijn)
        Randomize
        
        ' Headers instellen in rij 1
        ws.Range("A1").Value = "Product"   ' Header voor kolom A
        ws.Range("B1").Value = "Verkoop"   ' Header voor kolom B
        
        ' Variabelen voor de lus
        Dim i As Long          ' Rownummer
        Dim randomIndex As Long ' Index voor het product in de array
        Dim randomGetal As Long ' Willekeurig getal voor verkoop
        
        ' Lus over de rijen 2 t/m 11
        For i = 2 To 11
            ' Willekeurige index kiezen voor het product (0 t/m 2, omdat de array 3 items heeft)
            randomIndex = Int((UBound(producten) - LBound(producten) + 1) * Rnd) + LBound(producten)
            
            ' Productnaam in kolom A zetten
            ws.Cells(i, 1).Value = producten(randomIndex)
            
            ' Willekeurig getal tussen 100 en 1000 genereren
            ' Formule: Int((boven - onder + 1) * Rnd) + onder
            randomGetal = Int((1000 - 100 + 1) * Rnd) + 100
            
            ' Willekeurig getal in kolom B zetten
            ws.Cells(i, 2).Value = randomGetal
        Next i
    End Sub
---
    Sub MaakDropdownUniekeProducten()
        ' Werkbladvariabele declareren en instellen op het actieve werkblad
        Dim ws As Worksheet
        Set ws = ActiveSheet
        
        ' Bereik met brondata (producten in kolom A, rijen 2 t/m 11)
        Dim rngBron As Range
        Set rngBron = ws.Range("A2:A11")
        
        ' Collection om unieke waarden in op te slaan
        Dim uniekeProducten As New Collection
        
        ' Variabelen voor lus
        Dim cel As Range
        Dim waarde As Variant
        
        ' Foutafhandeling tijdelijk aanzetten om dubbele keys in Collection op te vangen
        On Error Resume Next
        
        ' Door elke cel in de bronlijst lopen
        For Each cel In rngBron
            waarde = cel.Value
            
            ' Alleen iets doen als de cel niet leeg is
            If waarde <> "" Then
                ' Proberen toe te voegen aan de Collection
                ' Als de waarde al bestaat, geeft dit een fout, die we negeren
                uniekeProducten.Add Item:=waarde, Key:=CStr(waarde)
            End If
        Next cel
        
        ' Foutafhandeling weer uitzetten
        On Error GoTo 0
        
        ' Hulpkolom leegmaken (bijvoorbeeld kolom F) zodat we een schone lijst hebben
        ws.Range("F:F").ClearContents
        
        ' Unieke waarden naar kolom F schrijven, start in F2
        Dim i As Long
        For i = 1 To uniekeProducten.Count
            ws.Cells(i + 1, "F").Value = uniekeProducten(i) ' i+1 omdat we in rij 2 beginnen
        Next i
        
        ' Laatste rij van de unieke lijst bepalen
        Dim laatsteRijUniek As Long
        laatsteRijUniek = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
        
        ' Eventuele bestaande data validation in D2 verwijderen
        ws.Range("D2").Validation.Delete
        
        ' Data validation (dropdown) toevoegen in D2
        ' Type:=xlValidateList betekent een keuzelijst
        ' Formula1 verwijst naar het bereik met unieke producten in kolom F
        With ws.Range("D2").Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="=$F$2:$F$" & laatsteRijUniek
            .IgnoreBlank = True
            .InCellDropdown = True ' Zorgt ervoor dat het pijltje zichtbaar is
            .ShowError = True
        End With
        
        ' Opmerking:
        ' Een standaard Data Validation-lijst laat maar één waarde tegelijk toe,
        ' dus hiermee kan de gebruiker slechts één product per keer kiezen.
    End Sub
---
    Sub BerekenSomIf()
        ' Werkblad instellen
        Dim ws As Worksheet: Set ws = ActiveSheet
        
        ' Geselecteerd product uit D2 ophalen
        Dim gekozenProduct As String
        gekozenProduct = ws.Range("D2").Value
        
        ' Variabelen voor lus en som
        Dim i As Long, totaal As Double
        
        ' Door rijen 2 t/m 11 lopen en verkoop optellen als product overeenkomt
        For i = 2 To 11
            If ws.Cells(i, 1).Value = gekozenProduct Then
                totaal = totaal + ws.Cells(i, 2).Value
            End If
        Next i
        
        ' Resultaat in E2 zetten
        ws.Range("E2").Value = totaal
    End Sub
---
    Sub MaakKnopVoorBerekening()
        ' Verwijzing naar het actieve werkblad
        Dim ws As Worksheet: Set ws = ActiveSheet
        
        ' Een nieuwe knop (Form Control) toevoegen op het werkblad
        Dim knop As Shape
        Set knop = ws.Shapes.AddFormControl(Type:=xlButtonControl, _
                                            Left:=100, Top:=50, Width:=120, Height:=30)
        
        ' Tekst op de knop zetten
        knop.TextFrame.Characters.Text = "Bereken Verkoop"
        
        ' De knop koppelen aan de bestaande macro BerekenSomIf
        knop.OnAction = "Sheet1.BerekenSomIf"
    End Sub


