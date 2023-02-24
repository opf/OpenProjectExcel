Attribute VB_Name = "GanttModul"
' Code Origin: Jakbob
' Changed by Horst Schmid 2023

Option Explicit
Public Const MyExcel8 = 56 'xlExcel8 is undefined in Office2003! Eventuell max. nur 255 Spalten

Type LabelInfoType
 srcColum As Integer
 srcNameIntl As String
 srcName As String
End Type


Sub myPageSetup(oWs As Worksheet, titleRowCount As Integer)
  Application.PrintCommunication = False
  With oWs.PageSetup
    .LeftMargin = Application.CentimetersToPoints(1.6)
    .RightMargin = Application.CentimetersToPoints(1.6)
    .TopMargin = Application.CentimetersToPoints(1.6)
    .BottomMargin = Application.CentimetersToPoints(1.6)
    .HeaderMargin = Application.CentimetersToPoints(0.8)
    .FooterMargin = Application.CentimetersToPoints(0.8)
      
    .PrintTitleRows = oWs.Rows("1:" & titleRowCount).Address
    '.PrintTitleColumns = "$D:$D"
    
    .Orientation = xlLandscape
    .FitToPagesWide = 1
    .FitToPagesTall = 1
  End With
  Application.PrintCommunication = True
End Sub

    
  Public Function DateOfFirstDayInCalenderWeek(myKW As Long, Optional myJahr As Variant, Optional firstDayInWeek As VbDayOfWeek = vbUseSystemDayOfWeek, Optional firstWeekInYear As VbFirstWeekOfYear = vbUseSystem) As Date
    ' gibt das Datum des ersten Tags der übergebenen Kalenderwoche im übergebenen oder aktuellen Kalenderjahr zurück
    ' Funktioniert für die jeweils gewählte Varianten des Beginns der 1. Woche im Jahr, egal ob 1. Woche am 29.12. oder am 7.1. beginnt
    Dim myDate As Date
    If IsMissing(myJahr) Then myJahr = Year(Date) 'IsMissing works only for Variant!
    myDate = DateSerial(myJahr, 1, 1) 'Start mit 1.Jan im Jahr
    'Suche jetzt nach dem Beginn der 1. Woche:
    Do While DatePart("ww", myDate, firstDayInWeek, firstWeekInYear) > 50 'Letzte KW des Vorjahrs reicht noch in diese Jahr
      myDate = myDate + 1
    Loop
    Do While DatePart("ww", myDate - 1, firstDayInWeek, firstWeekInYear) = 1 'wenn der Tag zuvor noch KW 1 ist , dann gehe zurück ins vorige Jahr
      myDate = myDate - 1
    Loop
    DateOfFirstDayInCalenderWeek = myDate + 7 * (myKW - 1) ' jetzt noch die Tage entspr. der gewünschten Woche hinzuaddieren
  End Function


  'Gibt die Kalenderwoche aus Datum zurück
  'geg. werden die Wochen-Anzahl der Vorjahre hinzuaddiert!
  Public Function myKW(DasDatum As Date, Optional StartJahr As Integer = 0, Optional firstDayInWeek As VbDayOfWeek = vbUseSystemDayOfWeek, Optional firstWeekInYear As VbFirstWeekOfYear = vbUseSystem) As Integer
    ' Hint 1:
    ' The Bug in DatePart("ww", ...), which is described under
    ' https://learn.microsoft.com/de-de/office/troubleshoot/access/functions-return-wrong-week-number
    ' seems to be valid only in MS-Access, not in Excel!
    '
    ' What's considerd here:
    ' For ISO 8601 the CW 1 can start on 29th Dec (e.g. year 2014) and Dates in January can have CW52 or CW53
    Dim kw As Integer
    Dim jahr As Integer
    
    kw = DatePart("ww", DasDatum, firstDayInWeek, firstWeekInYear)
    'kw = WorksheetFunction.IsoWeekNum(DasDatum) 'ISO 8601 would be a Solution for Europe
    
    If StartJahr > 0 Then
      ' Wochen der Vorjahre hinzuaddieren:
      jahr = Year(DasDatum)
      If (kw > 51) And (Month(DasDatum) = 1) Then 'Anfang Januar: wir sind mit Zählung schon im Vorjahr!
        jahr = jahr - 1
      ElseIf (kw = 1) And (Month(DasDatum) = 12) Then 'Ende Dezember: Wir sind mit Zählung schon im nächsten Jahr!
        jahr = jahr + 1
      End If
      Do While jahr > StartJahr 'geg. die Wochen der Vorjahr noch hinzurechnen:
        jahr = jahr - 1
        If firstWeekInYear = vbFirstFullWeek Then
          ' z.B. am 30.12 kann nochmal eine neue Woche beginnen
          kw = kw + DatePart("ww", DateSerial(jahr, 12, 31), firstDayOfWeek:=firstDayInWeek, firstweekofyear:=firstWeekInYear)
        Else
          'bei ISO8601 kann am 29.12. die KW 1 beginnen, der 28.12. gehört immer zur letzten Woche 52 oder 53. Der 29.12. kann zu KW 1 gehören
          kw = kw + DatePart("ww", DateSerial(jahr, 12, 28), firstDayOfWeek:=firstDayInWeek, firstweekofyear:=firstWeekInYear)
        End If
      Loop
    End If
    myKW = kw
  End Function


  Public Function myMonth(DasDatum As Date, Optional StartJahr As Integer = 0) As Integer
    Dim m As Integer
    Dim jahr As Integer
    m = Month(DasDatum)
    If StartJahr > 0 Then
      jahr = Year(DasDatum)
      m = m + 12 * (jahr - StartJahr)
    End If
    myMonth = m
  End Function


    Sub FeiertageHolen(jahr As Integer, ByRef Feiertage As Dictionary)
      'https://feiertage-api.de
      'https://feiertage-api.de/api/?jahr=2019&nur_land=BW
      Dim bundesLand As String
      Dim url As String
      Dim json As String
      Dim wr As String
      Dim res As Dictionary
      Dim res2() As Variant
      Dim res3 As Dictionary
      Dim d2() As String
      Dim res4
      Dim n As Integer
      Dim d1 As Date
        
      bundesLand = ThisWorkbook.Worksheets("GanttConfig").Range("GanttFeiertagsBundesland").Value
      If bundesLand <> "" Then
        bundesLand = "&nur_land=" & Left(bundesLand, 2)
      End If
      url = "http://feiertage-api.de/api/?jahr=" & CStr(jahr) & bundesLand
      
      'ParseJson(ByVal JsonString As String) As Object
      wr = HttpGet(url)
      Set res = JsonConverter.ParseJson(wr) 'Dictionary mit Feiertagsnamen
      'res2 = res.Items
      For n = 0 To res.Count - 1
        Set res3 = res.Items(n)
        res4 = res3.Items
'        Debug.Print res4(1) 'Kommentar
        If res4(1) = "" Then
          'd1 = CDate(res4(0)) 'was ist ein gültiger Datumsausdruck, der hier problemlos konvertiert wird?
          d2 = Split(res4(0), "-")
          d1 = DateSerial(d2(0), d2(1), d2(2))
          Feiertage.Add d1, res.Keys(n) ' Datum und Name des Feiertags
'        Else
'          Debug.Print res4(1) 'Kommentar bei weiteren schulfreien Tagen
        End If
      Next n
      'Set FeiertageHolen = Feiertage
    End Sub


    'Possibly there is in this Package already a similar funtion, which may be used instead of this!
    Function HttpGet(s1 As String, Optional DstFileName As String = "") As String
      'Download a file from Internet
      ' - without DstFileName: File Content is returned as result of function
      ' - with DstFileName: Content is stored to that file and the FunctionResult is the status "200 - OK" or similar
    
      Dim d() As Byte
      Dim WHttpR As Object
      Dim fNr As Integer
      
      Set WHttpR = CreateObject("WinHttp.WinHttpRequest.5.1") 'requires WinHttp.dll
      If WHttpR Is Nothing Then Set WHttpR = CreateObject("WinHttp.WinHttpRequest")
      If WHttpR Is Nothing Then Set WHttpR = CreateObject("MSXML2.ServerXMLHTTP")
      If WHttpR Is Nothing Then Set WHttpR = CreateObject("Microsoft.XMLHTTP")
      'alternatively: Set WHttpR = CreateObject("WinHttp.WinHttpRequest.5") 'requires WinHttp5.dll
    
      WHttpR.SetAutoLogonPolicy AutoLogonPolicy_Always '5.1 and also ???, e.g. for SharePoint-Server
      WHttpR.Open "GET", s1, False
      'If DstFileName <> "" Then
        'WHttpR.SetRequestHeader "content-type", "application/x-msdownload" 'application/octet-stream
      'End If
      WHttpR.Send
      If DstFileName = "" Then
        'Debug.Print WHttpR.Status & " - " & WHttpR.StatusText
        HttpGet = WHttpR.ResponseText
      Else
        HttpGet = WHttpR.Status & " - " & WHttpR.StatusText
        If WHttpR.Status = 200 Then
          fNr = FreeFile()
          Open DstFileName For Binary As fNr
          d() = WHttpR.ResponseBody
          Put fNr, 1, d()
          Close fNr
        End If
      End If
    End Function



Sub Terminplan()
    Dim oWs As Worksheet
    Dim WorkpackagesSheet As Worksheet
    Dim GanttConfigSheet As Worksheet
    Dim kw0 As Long 'Kalenderwoche der ersten Spalte
    Dim month0 As Integer 'Monat der ersten Spalte

    Dim jahr0 As Integer 'Jahr, in dem die KalenderWochen-Zählung beginnt
    Dim startdatum As Date
    Dim enddatum As Date
    Dim startMonth As Integer ' für Balken
    Dim endMonth As Integer ' für Balken
    Dim mileStoneDatum As Date
    Dim Bereich As Range
    Dim myColumnStart As Integer
    Dim myColumnEnde As Integer
    Dim DiagramHeadlineRange As Range
    Dim reihe As Long
    Dim spaltenNummerAnfang As Integer
    Dim spaltenNummerEnde As Integer
    Dim spaltenNummerMilestoneDatum As Integer
    Dim spaltenNummerFuerBalkenFarbe As Integer 'z.B Spaltennummer der Kategorie oder des Typs
    Dim spaltenNummerFuerBalkenMuster As Integer
    Dim spaltenNummerTyp As Integer ' für MileStone-Erkennung
    Dim nextBarColor As Integer 'im Range 'GanttBarColors' das nächste, bisher nicht benutzte Feld
    Dim nextBarPattern As Integer 'im Range 'GanttBarPatterns' das nächste, bisher nicht benutzte Feld
    Dim barColorCell As Range ' Eine Zelle aus dem Range 'GanttBarColors'
    Dim barPatternCell As Range ' Eine Zelle aus dem Range 'GanttBarPatterns'
    'Dim arr_over As Variant, blnFound As Boolean
    Dim letzteZeile As Long 'Workpackages-Sheet
    Dim letzteZeileAttributes As Long
    Dim AttrTranslationLookupRange As Range
    Dim spalteID As Integer
    Dim differenz As Integer 'Spalten für Gatt
    Dim c As Long
    Dim cw As Long
    Dim weekStartDate As Date
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim n As Integer
    Dim r As Variant
    Dim firstDayOfWeek As Date
    Dim firstDay As Date
    Dim enddatumID As Date
    Dim startdatumID As Date
    Dim labelSourceColumnNumbers(1, 2) As LabelInfoType 'Left (0,0) ... (0,2), right (1,0) ... (1,2)
    Dim LabelInfo As LabelInfoType
    Dim labelCfgCellNames(1) As String
    Dim labelText(1) As String '0=links, 1=rechts
    Dim hint As String
    Dim myDate As Date
    Dim oRngBalken As Range
    Dim startKW As Integer
    Dim endeKW As Integer
    Dim balkenStartSpalte As Integer
    Dim balkenEndeSpalte As Integer
    Dim SpaltenDauerTage As Integer ' 1 (Tag) oder 7 (Woche) oder 0 für Monate
    Dim SpaltenDauerMonate As Integer ' 0 (für Tage, Wochen), 1 (Monate) 3 (Quartale)
    Dim foundCell As Range
    Dim myItem As String
    Dim myTemplateCell As Range
    Dim colWidth As Double
    Dim Feiertage As New Dictionary
    Dim BalkenFarbe As New Dictionary
    Dim BalkenMuster As New Dictionary
    Dim CellComment
    
    'Application.ScreenUpdating = False
    labelCfgCellNames(0) = "GanttLeftLabel"
    labelCfgCellNames(1) = "GanttRightLabel"
    Application.EnableEvents = False ' kein CreateDropdownsInRows
    
    'Get and Check Configuration:
    Set GanttConfigSheet = ThisWorkbook.Worksheets("GanttConfig")
    SpaltenDauerMonate = 0
    SpaltenDauerTage = 0
    If WorkpackagesSheet Is Nothing Then
      Set WorkpackagesSheet = ThisWorkbook.Worksheets("Workpackages")
    End If
    letzteZeile = LastUsedRow(WorkpackagesSheet)
    If letzteZeile < 3 Then
      MsgBox "Please load first your workpackges from OpenProject to this Excel Workbook!", , "Gantt Error"
      Exit Sub
    End If
    ThisWorkbook.Worksheets("Attributes").Activate
    letzteZeileAttributes = LastUsedRow(ThisWorkbook.Worksheets("Attributes"))
    WorkpackagesSheet.Activate
    If letzteZeileAttributes = 0 Then
      MsgBox "Please load first your workpackges and Attributs from OpenProject to this Excel Workbook!", , "Gantt Error"
      Exit Sub
    End If
    Application.Cursor = xlWait
    With GanttConfigSheet
      If .Range("GanttResolution") = "Day" Then
        SpaltenDauerTage = 1
      ElseIf .Range("GanttResolution") = "Week" Then
        SpaltenDauerTage = 7
      ElseIf .Range("GanttResolution") = "Month" Then
        SpaltenDauerMonate = 1
      ElseIf .Range("GanttResolution") = "Quarter" Then
        SpaltenDauerMonate = 3
      End If
      If .Range("GanttDiagramLocation") = "Workpackages Sheet" Then
        Set oWs = WorkpackagesSheet
      ElseIf .Range("GanttDiagramLocation") = "New Sheet" Then
        WorkpackagesSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set oWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
      ElseIf .Range("GanttDiagramLocation") = "New Workbook" Then
        WorkpackagesSheet.Copy
        Set oWs = ActiveWorkbook.Worksheets(1)
      End If
    End With
    
    letzteZeile = LastUsedRow(oWs)
    Set AttrTranslationLookupRange = ThisWorkbook.Worksheets("Attributes").Range(ThisWorkbook.Worksheets("Attributes").Cells(2, 1), ThisWorkbook.Worksheets("Attributes").Cells(letzteZeileAttributes, 2))
    'Const sucheInReihe = 1
    Set foundCell = oWs.Cells(1, 1).EntireRow.Find(What:="ID", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not foundCell Is Nothing Then
      spalteID = foundCell.Column 'Mehrere Zeilen mit gleicher ID könnten vorhanden sein: Bisher nicht bearbeitet
    End If
    
    ' LookUp der sprachspezifischen Spalten-Namen:
    myItem = WorksheetFunction.VLookup("startDate", AttrTranslationLookupRange, 2, False)
    Set foundCell = oWs.Cells(1, 1).EntireRow.Find(What:=myItem, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If foundCell Is Nothing Then
      Application.Cursor = xlDefault
      MsgBox "Column 'startDate' (='" & myItem & " ') not found in Workpackages sheet", , "Gantt Error"
      Exit Sub
    Else
      spaltenNummerAnfang = foundCell.Column
    End If
    myItem = WorksheetFunction.VLookup("dueDate", AttrTranslationLookupRange, 2, False) 'z.B. "EndTermin"
    Set foundCell = oWs.Cells(1, 1).EntireRow.Find(What:=myItem, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If foundCell Is Nothing Then
      Application.Cursor = xlDefault
      MsgBox "Column 'dueDate' (='" & myItem & " ') not found in Workpackages sheet", , "Gantt Error"
      Exit Sub
    Else
      spaltenNummerEnde = foundCell.Column
    End If
    
    ' Datumsbereich holen:
    enddatum = CDate(WorksheetFunction.Max(oWs.Range(oWs.Cells(1 + 1, spaltenNummerEnde), oWs.Cells(letzteZeile, spaltenNummerEnde))))
    startdatum = CDate(WorksheetFunction.Min(oWs.Range(oWs.Cells(1 + 1, spaltenNummerAnfang), oWs.Cells(letzteZeile, spaltenNummerEnde))))
    'Milestone-Datum steht in extra Spalte, falls vor bisher erstem Datum oder nach bisher letztem Datum, dann Bereich korrigieren:
    myItem = WorksheetFunction.VLookup("date", AttrTranslationLookupRange, 2, False)
    Set foundCell = oWs.Cells(1, 1).EntireRow.Find(What:="Datum", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not foundCell Is Nothing Then
      spaltenNummerMilestoneDatum = foundCell.Column
      mileStoneDatum = CDate(WorksheetFunction.Min(oWs.Range(oWs.Cells(1 + 1, spaltenNummerMilestoneDatum), oWs.Cells(letzteZeile, spaltenNummerMilestoneDatum))))
      If mileStoneDatum < startdatum Then
        startdatum = mileStoneDatum
      End If
      mileStoneDatum = CDate(WorksheetFunction.Max(oWs.Range(oWs.Cells(1 + 1, spaltenNummerMilestoneDatum), oWs.Cells(letzteZeile, spaltenNummerMilestoneDatum))))
      If mileStoneDatum > enddatum Then
        enddatum = mileStoneDatum
      End If
    End If
    
    ' Anfang geg. z.B. auf Wochen-, Monats-, Quartalsanfang verringern:
    jahr0 = Year(startdatum) 'Jahr, in dem die Kalenderwochenzählung beginnt
    kw0 = myKW(startdatum)
    If (Month(startdatum) = 1) And (kw0 > 51) Then ' 1.Jan gives kw=52 or 53!
      jahr0 = jahr0 - 1
    End If
    firstDayOfWeek = DateOfFirstDayInCalenderWeek(kw0, jahr0)
    If SpaltenDauerTage = 7 Then
      differenz = DateDiff("w", startdatum, enddatum)
      firstDay = firstDayOfWeek
    ElseIf SpaltenDauerTage = 1 Then
      differenz = DateDiff("d", startdatum, enddatum)
      firstDay = startdatum
    ElseIf SpaltenDauerMonate = 1 Then
      differenz = (1 + DateDiff("m", startdatum, enddatum)) / SpaltenDauerMonate ' ist plus 1 hier notwendig?
      startdatum = DateSerial(Year(startdatum), Month(startdatum), 1)
    ElseIf SpaltenDauerMonate > 1 Then
      ' SpaltenDauerMonate=3: differenz = DateDiff("q", startdatum, enddatum)
      ' Für Quartale, Halbjahre, 12 Monate=Jahr:
      'StartDatum auf z.B. Quartals oder Halbjahresanfang abrunden, z.B. 20.Nov auf Quartalsanfang 1.Okt, HalbjahresAnfang 1.Juli, Jahresanfang 1.Jan:
      Dim st As Date
      month0 = 1 + SpaltenDauerMonate * Round((Month(startdatum) - 1) \ SpaltenDauerMonate)
      st = DateSerial(Year(startdatum), month0, 1)
      differenz = Application.WorksheetFunction.RoundUp((1 + DateDiff("m", st, enddatum)) / SpaltenDauerMonate, 0)
      startdatum = DateSerial(Year(startdatum), Month(startdatum), 1)
    End If
    
    'Überflüssige, leere Spalten löschen:
    ' sollte dies bei vielleicht bei "Diagram Location"="Workpackages Sheet" übersprungen werden???
    m = LastUsedColumn(oWs)
    n = oWs.Cells.SpecialCells(xlCellTypeLastCell).Column
    oWs.Range(oWs.Cells(1, m + 1), oWs.Cells(1, n)).EntireColumn.Delete 'Überflüssige Spalten löschen
    
    ' Fügt rechts neben ausgewählter Spalte so viele Spalten ein wie für die Wochen von Start- bis zum Enddatum benötigt
    If GanttConfigSheet.Range("GanttColumnInsertPosition") = "after DueDate column" Then
      myColumnStart = 1 + spaltenNummerEnde
    Else
      myColumnStart = 1 + m
    End If
    
    'Spalten (incl. je einer zusätzl Spalte für Labels davor und danach) einfügen:
    On Error GoTo ColumnInsertFailure
    oWs.Range(oWs.Cells(1, myColumnStart), oWs.Cells(1, myColumnStart + differenz + 1 + 2)).EntireColumn.Insert xlShiftToRight
    On Error GoTo 0
    oWs.Range(oWs.Cells(1, myColumnStart), oWs.Cells(1, myColumnStart + differenz + 1 + 2)).EntireColumn.ClearFormats 'z.B. 'Shrink to Fit'
    
    myColumnStart = myColumnStart + 1 ' Labelspalte
    myColumnEnde = myColumnStart + differenz + 1
    Set DiagramHeadlineRange = oWs.Range(oWs.Cells(1, myColumnStart), oWs.Cells(1, myColumnEnde))
    DiagramHeadlineRange.EntireColumn.Validation.Delete
    
    'Headline(s) mit Datums- bzw KW-Woche füllen:
    If Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 3) = "No" Then
      DiagramHeadlineRange.Rows(1).Orientation = 90
    ElseIf Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 3) = "Add" Then
      ' Zeile 1 für Kalenderwoche und Zeile 2 fürs Datum
      If GanttConfigSheet.Range("GanttDiagramLocation") = "Workpackages Sheet" Then
        MsgBox "Sorry, additionally to a date row a row with the calender week is only for new sheets possible"
      ElseIf (GanttConfigSheet.Range("GanttResolution") <> "Day") And (GanttConfigSheet.Range("GanttResolution") <> "Week") Then
        MsgBox "Sorry, additionally to a date row a row with the calender week is only for the resolutions DAY and WEEK possible"
      Else
        'Einfügen einer zusätzlichen Zeile:
        oWs.Cells(2, 1).EntireRow.Insert xlShiftDown
        oWs.Cells(1, 1).RowHeight = 30 ' sonst sind Wochenzahlen hinter den Autofilter-DropDownflächen verborgen
        DiagramHeadlineRange.Rows(1).Orientation = 0
        letzteZeile = LastUsedRow(oWs)
        With DiagramHeadlineRange
          .NumberFormat = "General" ' format Standard
          .VerticalAlignment = xlTop
          .HorizontalAlignment = xlCenter
        End With
        '.Borders(xlEdgeTop).LineStyle = xlNone
        '.Borders(xlEdgeBottom).LineStyle = xlNone
        '.Borders(xlInsideVertical).LineStyle = xlNone
        '.Borders(xlInsideHorizontal).LineStyle = xlNone
        Set DiagramHeadlineRange = Nothing
        Set DiagramHeadlineRange = oWs.Range(oWs.Cells(2, myColumnStart), oWs.Cells(2, myColumnEnde))
        DiagramHeadlineRange.Orientation = 90
      End If
    End If
    If Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 3) <> "YES" Then
      DiagramHeadlineRange.NumberFormatLocal = GanttConfigSheet.Range("GanttDateDisplayFormat").Value
    End If
    
    'Vorbereitungen für Balken-Farben, Balken-Muster und Balken-Labels: Welche Spalten müssen ausgelesen werden:
    myItem = GanttConfigSheet.Range("GanttBarColorFrom") ' e.g "Type", "assignee" or "category"
    myItem = WorksheetFunction.VLookup(myItem, AttrTranslationLookupRange, 2, False) ' e.g. "Typ", "Zugewiesen an", "Kategorie"
    Set foundCell = oWs.Cells(1, 1).EntireRow.Find(What:=myItem, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not foundCell Is Nothing Then
      spaltenNummerFuerBalkenFarbe = foundCell.Column 'Spalte, deren Inhalt die Balkenfarbe bestimmt
    End If
    
    myItem = GanttConfigSheet.Range("GanttBarPatternFrom") ' e.g "Version", "assignee" or "category"
    myItem = WorksheetFunction.VLookup(myItem, AttrTranslationLookupRange, 2, False) ' e.g. "Version", "Zugewiesen an", "Kategorie"
    Set foundCell = oWs.Cells(1, 1).EntireRow.Find(What:=myItem, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If Not foundCell Is Nothing Then
      spaltenNummerFuerBalkenMuster = foundCell.Column 'Spalte, deren Inhalt das Balkenmuster bestimmt
    End If
    
    
    myItem = WorksheetFunction.VLookup("Type", AttrTranslationLookupRange, 2, False) ' e.g. "Typ", "Zugewiesen an", "Kategorie"
    Set foundCell = oWs.Cells(1, 1).EntireRow.Find(What:=myItem, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    spaltenNummerTyp = 0 ' für Erkennung von Milestone
    If Not foundCell Is Nothing Then
      spaltenNummerTyp = foundCell.Column
    End If
    
    'spaltenNummern mit Angaben für die Labels holen:
    For m = 1 To 3
      For n = 0 To 1 ' 0=Left, 1=Right
        myItem = GanttConfigSheet.Range(labelCfgCellNames(n) & m).Value
        If myItem <> "" Then
          LabelInfo.srcNameIntl = myItem
          myItem = WorksheetFunction.VLookup(myItem, AttrTranslationLookupRange, 2, False)
          Set foundCell = oWs.Cells(1, 1).EntireRow.Find(What:=myItem, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
          If Not foundCell Is Nothing Then
            LabelInfo.srcColum = foundCell.Column
            LabelInfo.srcName = myItem
            labelSourceColumnNumbers(n, m - 1) = LabelInfo
          End If
        End If
      Next n
    Next m
    
    'Spaltenbreite:
    If differenz < 40 Then
      colWidth = 3
    ElseIf differenz < 140 Then
      colWidth = 2
    Else
      colWidth = 1.1
    End If
    DiagramHeadlineRange.ColumnWidth = colWidth
    
    'je x-te Zeile für bessere Lesbakeit Hellgrau markieren:
    For m = 5 To letzteZeile Step 4
      oWs.Cells(m, 1).EntireRow.Interior.Color = RGB(248, 248, 248)
    Next m
    
    'Arbeitsfreie Spalten (Wochenenden, Deutsche Feiertage) markieren:
    Feiertage.RemoveAll
    For m = Year(startdatum) To Year(enddatum)
      FeiertageHolen m, Feiertage ' Feiertage aus Internet holen und dem Dictionary Feiertage hinzufügen
    Next m
    
    'Trägt in eingefügten Bereich in erste Zeile das Datum ein:
    ' Bei Resolution=Day und mehr als xx Spalten die Montage als Datum ein
    For j = 0 To differenz + 1
      myDate = DateSerial(Year(firstDay), Month(firstDay) + j * SpaltenDauerMonate, Day(firstDay) + j * SpaltenDauerTage)
      If SpaltenDauerTage = 1 Then
        If Feiertage.Exists(myDate) Or Weekday(myDate, vbMonday) > 5 Then
          oWs.Range(oWs.Cells(2, j + myColumnStart), oWs.Cells(letzteZeile, j + myColumnStart)).Interior.Color = RGB(235, 235, 235)
          If Feiertage.Exists(myDate) Then 'not working as expected !!??
            If Feiertage(myDate) <> "" Then ' workaround
              With oWs.Cells(2, j + myColumnStart)
                Set CellComment = .AddComment
                .Comment.Visible = False
                .Comment.Text Text:=Feiertage(myDate)
                CellComment.Shape.Width = 70
                CellComment.Shape.Height = 12
              End With
            End If
          End If
        End If
        'Kalenderwoche eintragen:
        If (Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 3) <> "No") Then
          cw = myKW(myDate, Year(myDate))
          weekStartDate = DateOfFirstDayInCalenderWeek(cw, Year(myDate))
          If (weekStartDate = myDate) Or ((j = 0) And (myDate - weekStartDate > 0)) Then
            
            m = j + myColumnStart + 6 - (myDate - weekStartDate)
            If m > myColumnEnde Then
              m = myColumnEnde '.MergeCells macht sonst Problem falls am Ende keine ganze Woche mehr!
            End If
            With oWs.Range(oWs.Cells(1, j + myColumnStart), oWs.Cells(1, m))
              .HorizontalAlignment = xlCenter
              .MergeCells = True
              .Borders(xlEdgeLeft).LineStyle = xlContinuous
              .Borders(xlEdgeLeft).Weight = xlThin
              .Borders(xlEdgeRight).LineStyle = xlContinuous
              .Borders(xlEdgeRight).Weight = xlThin
            End With
            oWs.Cells(1, j + myColumnStart) = myKW(myDate)
          End If
        End If
      End If
      If colWidth > 2.8 Then
'        If Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 3) <> "No" Then ' Yes or Add
'          oWs.Cells(1, j + myColumnStart) = DIN_KW(myDate)
'        End If
        If Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 3) <> "Yes" Then 'No or Add
          oWs.Cells(DiagramHeadlineRange.Row, j + myColumnStart) = myDate
        End If
      ElseIf SpaltenDauerTage = 1 Then
        ' nur jeden 7. Tag eintragen:
        If j Mod 7 = 0 Then
          m = j + myColumnStart + 6
          If m > myColumnEnde Then
            m = myColumnEnde '.MergeCells macht sonst Problem falls am Ende keine ganze Woche mehr!
          End If
          With oWs.Range(oWs.Cells(DiagramHeadlineRange.Row, j + myColumnStart), oWs.Cells(DiagramHeadlineRange.Row, m))
            .HorizontalAlignment = xlLeft
            .MergeCells = True
          End With
          If Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 2) <> "No" Then
            oWs.Cells(1, j + myColumnStart) = myKW(myDate)
          End If
          If Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 3) <> "Yes" Then
            oWs.Cells(DiagramHeadlineRange.Row, j + myColumnStart) = myDate
          End If
        End If
      Else 'z.B. 1 Woche je Spalte und sehr schmale Spalten
        ' Eine gute Lösung fehlt noch !!!!
        ' vorläufig:
        If Left(GanttConfigSheet.Range("GanttCalenderWeekDisplay"), 3) = "Yes" Then
          oWs.Cells(1, j + myColumnStart) = myKW(myDate)
        Else
          oWs.Cells(1, j + myColumnStart) = myDate
        End If
      End If
    Next j
    With DiagramHeadlineRange.Rows(1)
      '.EntireRow.AutoFit 'funktioniert nicht!!
      .Rows.RowHeight = 8 * Len(GanttConfigSheet.Range("GanttDateDisplayFormat").Value)
      'vertikale Ausrichtung oben!!!!!
      .ShrinkToFit = True
    End With
    
    'Fest zugeordnete Balkenfarben vorbereiten:
    BalkenFarbe.RemoveAll
    BalkenFarbe.CompareMode = TextCompare
    For j = 1 To GanttConfigSheet.Range("GanttBarColors").Columns.Count
      If GanttConfigSheet.Range("GanttBarColors").Cells(2, j).Text <> "" Then
        BalkenFarbe.Add Key:=GanttConfigSheet.Range("GanttBarColors").Cells(2, j).Text, Item:=GanttConfigSheet.Range("GanttBarColors").Cells(1, j).Interior
      End If
    Next j
    'Fest zugeordnete Balkenmuster vorbereiten:
    BalkenMuster.RemoveAll
    BalkenMuster.CompareMode = TextCompare
    For j = 1 To GanttConfigSheet.Range("GanttBarPatterns").Columns.Count
      If GanttConfigSheet.Range("GanttBarPatterns").Cells(2, j).Text <> "" Then
        BalkenMuster.Add GanttConfigSheet.Range("GanttBarPatterns").Cells(2, j).Text, GanttConfigSheet.Range("GanttBarPatterns").Cells(1, j).Interior
      End If
    Next j
    
    
    '==================================
    'Workpackages als Balken eintragen und links & rechts Label dazu:
    nextBarColor = 0 ' Spalte in GanttConfigSheet.Range("GanttBarColors") mit den Füllfarben
    nextBarPattern = 0 ' Spalte in GanttConfigSheet.Range("GanttBarPatterns") mit den Mustern
    For m = 1 + DiagramHeadlineRange.Row To letzteZeile 'für jedes Workpackage
      
      'Start- und Ende-Datum für den Balken bzw Milestone holen:
      enddatumID = oWs.Cells(m, spaltenNummerEnde).Value
      startdatumID = oWs.Cells(m, spaltenNummerAnfang).Value
      If startdatumID = 0 Then 'kein Startdatum, möglicherweise Milestone:
        startdatumID = oWs.Cells(m, spaltenNummerMilestoneDatum).Value
        enddatumID = oWs.Cells(m, spaltenNummerMilestoneDatum).Value
      ElseIf enddatumID = 0 Then
        enddatumID = enddatum
      End If
      If startdatumID = 0 Then
        startdatumID = startdatum 'Balken für WorkPackage ab Diagramm-Beginn
      End If
      If enddatumID = 0 Then
        enddatumID = enddatum 'Balken für WorkPackage bis zum Diagramm-Ende
      End If
      
      'Datum in Spalte umrechnen:
      If SpaltenDauerTage = 7 Then
        startKW = myKW(startdatumID, jahr0) 'für weitere Jahre auch >52!
        endeKW = myKW(enddatumID, jahr0)
        balkenStartSpalte = myColumnStart + startKW - kw0
        balkenEndeSpalte = myColumnStart + endeKW - kw0
      ElseIf SpaltenDauerTage = 1 Then
        balkenStartSpalte = myColumnStart + startdatumID - firstDay
        balkenEndeSpalte = myColumnStart + enddatumID - firstDay
      ElseIf SpaltenDauerMonate > 0 Then ' 1 Monat, 3 Monate, 6 Monate, 12 Monate, ...
        startMonth = 1 + SpaltenDauerMonate * Round((myMonth(startdatumID, jahr0) - 1) \ SpaltenDauerMonate)
        endMonth = 1 + SpaltenDauerMonate * Round((myMonth(enddatumID, jahr0) - 1) \ SpaltenDauerMonate)
        balkenStartSpalte = myColumnStart + (startMonth - month0) / SpaltenDauerMonate
        balkenEndeSpalte = myColumnStart + (endMonth - month0) / SpaltenDauerMonate
      End If
      Set oRngBalken = oWs.Range(oWs.Cells(m, balkenStartSpalte), oWs.Cells(m, balkenEndeSpalte))
      
      'Balkenfarbe:
      myItem = oWs.Cells(m, spaltenNummerFuerBalkenFarbe).Text ' Den Wert der Eigenschaft, welcher die Farbe bestimmen soll, holen
      If Not BalkenFarbe.Exists(myItem) Then ' Für diesen Wert (z.B. für diese Person) ist noch keine Farbe festgelegt
        ' Interior aus einer nicht fix zugeordneten Spalte von GanttConfigSheet.Range("GanttBarColors") ins Dictionary holen:
        Do
          nextBarColor = nextBarColor + 1
        Loop While (nextBarColor < GanttConfigSheet.Range("GanttBarColors").Columns.Count) And (GanttConfigSheet.Range("GanttBarColors").Cells(2, nextBarColor) <> "") 'noch nicht letzte Farbe
        BalkenFarbe.Add myItem, GanttConfigSheet.Range("GanttBarColors").Cells(1, nextBarColor).Interior ' Die Farbe für diesen Wert (z.B. Person) in Dictionary eintragen
      End If
      CopyInteriorColor src:=BalkenFarbe(myItem), dst:=oRngBalken.Interior 'Aus dem Interior-Object die Füllfarbe in die Zellen des Balkens kopieren
      
      'Balken-Muster:
      myItem = oWs.Cells(m, spaltenNummerFuerBalkenMuster).Text ' Den Wert der Eigenschaft, die das Muster bestimmen soll, holen
      If Not BalkenMuster.Exists(myItem) Then ' Für diesen Wert (z.B. für diese Version) ist noch kein Muster festgelegt
        ' Interior aus einer Spalte von GanttConfigSheet.Range("GanttBarPatterns") ins Dictionary holen:
        Do 'noch nicht das letzte Muster
          nextBarPattern = nextBarPattern + 1
        Loop While (nextBarPattern < GanttConfigSheet.Range("GanttBarPatterns").Columns.Count) And (GanttConfigSheet.Range("GanttBarPatterns").Cells(2, nextBarPattern).Text <> "")
        BalkenMuster.Add myItem, GanttConfigSheet.Range("GanttBarPatterns").Cells(1, nextBarPattern).Interior ' Das Muster für diesen Wert (z.B. Version) in Dictionary eintragen
      End If
      CopyInteriorPattern src:=BalkenMuster(myItem), dst:=oRngBalken.Interior 'Aus dem Interior-Object das Muster in die Zellen des Balkens kopieren
      
      If spaltenNummerTyp > 0 Then 'Typ-Spalte vorhanden
        If oWs.Cells(m, spaltenNummerTyp) = "Milestone" Then
          'Set oRngBalken.Borders = GanttConfigSheet.Range("GanttMilestoneBorders").Cells(1, 1).Borders
          CopyBorders src:=GanttConfigSheet.Range("GanttMilestoneBorders"), dst:=oRngBalken.Cells(1, 1)
        End If
      End If
      
      'Labels erstellen:
      For j = 0 To 1 '0=left and 1=right
        labelText(j) = ""
        For k = 0 To 2 'up to 3 items
          c = labelSourceColumnNumbers(j, k).srcColum
          If c > 0 Then 'Spalte ist wirklich vorhanden
            If (oWs.Cells(m, c).Text = "") And (spaltenNummerMilestoneDatum > 0) Then
              'falls Start oder EndeDatum und kein Wert: Auf 'Datum'-Spalte ausweichen, sofern diese vorhanden:
              If (labelSourceColumnNumbers(j, k).srcNameIntl = "startDate") Or (labelSourceColumnNumbers(j, k).srcNameIntl = "dueDate") Then
                c = spaltenNummerMilestoneDatum
              End If
            End If
            labelText(j) = labelText(j) & ";" & CStr(oWs.Cells(m, c).Text)
          End If
        Next k
      Next j
      If labelText(0) <> "" Then
        With oWs.Cells(m, balkenStartSpalte - 1)
          .NumberFormat = "@" 'Text
          .Formula = Mid(labelText(0), 2)
          .HorizontalAlignment = xlRight
        End With
      End If
      If labelText(1) <> "" Then
        With oWs.Cells(m, balkenEndeSpalte + 1)
          .NumberFormat = "@"
          .Formula = Mid(labelText(1), 2)
          '.HorizontalAlignment = xlLeft
        End With
      End If
    Next m 'nächstes Workpackage

    'Rahmen ums Diagramm:
    Dim b() As Variant
    b = Array(xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlEdgeTop)
    For i = 0 To UBound(b)
      With Range(Cells(1, myColumnStart - 1), Cells(m - 1, myColumnStart + differenz + 2)).Borders(b(i))
        .LineStyle = xlContinuous
        .Color = xlAutomatic
        .Color = RGB(0, 0, 0) 'Black
        .Weight = xlMedium
      End With
    Next i
    
    'Seiteneinrichtung:
    myPageSetup oWs, DiagramHeadlineRange.Rows.Count
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    Exit Sub
    
ColumnInsertFailure:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    r = oWs.UsedRange()
    c = oWs.UsedRange.Cells.SpecialCells(xlCellTypeLastCell).Column
    hint = ""
    If c + differenz > 254 Then
      hint = vbCrLf & "This may be a problem with the Excel limitation to 255 columns in some 32-bit versions and in elder file formats. "
    End If
    Application.Cursor = xlDefault
    MsgBox "Failed to add " & differenz & " additional columns to the Workpackages sheet." & hint & vbCrLf & "You may solve it by another resolution, e.g. not 'Day' but 'Week' or 'Month' for your chart.", , "Gantt Error"
End Sub


Sub CopyBorders(src As Range, dst As Range) 'for Milestones
  'xlDiagonalDown  5   Rahmen, der von der oberen linken Ecke zur unteren rechten Ecke jeder Zelle im Bereich verläuft.
  'xlDiagonalUp  6   Rahmen, der von der unteren linken Ecke zur oberen rechten Ecke jeder Zelle im Bereich verläuft.
  'xlEdgeBottom  9   Rahmen am unteren Rand des Bereichs
  'xlEdgeLeft  7   Rahmen am linken Rand des Bereichs.
  'xlEdgeRight   10  Rahmen am rechten Rand des Bereichs.
  'xlEdgeTop   8   Rahmen am oberen Rand des Bereichs
  'xlInsideHorizontal  12  Horizontale Rahmenlinien für alle Zellen im Bereich, mit Ausnahme von Rahmen am Außenrand des Bereichs
  'xlInsideVertical  11  Vertikale Rahmenlinien für alle Zellen im Bereich, mit Ausnahme von Rahmen am Außenrand des Bereichs
  Dim b() As Variant
  Dim i As Integer
  b = Array(xlDiagonalDown, xlDiagonalUp, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlInsideHorizontal, xlInsideVertical)
  For i = 0 To UBound(b)
    'dst.Borders(b(i)) = src.Borders(b(i)) ' not possible
    '  .LineStyle = xlContinuous
    '  .ColorIndex = xlAutomatic
    '  .TintAndShade = 0
    '  .Weight = xlThin
    dst.Borders(b(i)).LineStyle = src.Borders(b(i)).LineStyle
    'dst.Borders(b(i)).ColorIndex = src.Borders(b(i)).ColorIndex
    dst.Borders(b(i)).Color = src.Borders(b(i)).Color
    dst.Borders(b(i)).TintAndShade = src.Borders(b(i)).TintAndShade
    dst.Borders(b(i)).Weight = src.Borders(b(i)).Weight
  Next i
End Sub


Sub CopyInteriorColor(src As Interior, dst As Interior)
  'dst.ThemeColor = src.ThemeColor
  'dst.ColorIndex = src.ColorIndex '56 Colors
  dst.Color = src.Color
  dst.TintAndShade = src.TintAndShade
  'dst.Gradient = src.Gradient
  'dst.Pattern = src.Pattern
  'dst.PatternColor = src.PatternColor
  'dst.PatternColorIndex = src.PatternColorIndex
  'dst.PatternThemeColor = src.PatternThemeColor
  'dst.PatternTintAndShade = src.PatternTintAndShade
End Sub


Sub CopyInteriorPattern(src As Interior, dst As Interior)
  dst.Pattern = src.Pattern
  ' dst.PatternThemeColor = src.PatternThemeColor
  'dst.PatternColorIndex = src.PatternColorIndex
  dst.PatternTintAndShade = src.PatternTintAndShade
  dst.PatternColor = src.PatternColor
End Sub


Public Function LastUsedRow(oWs As Object) As Long
  'Workaround for buggy oWs.Cells.SpecialCells(xlCellTypeLastCell).Row
  Dim rowMax As Long, rowMax0 As Long
  Dim row2 As Long
  Dim n2 As Long
  
  oWs.UsedRange 'Reset if some Cell contents was cleared. Without this you will get the last row, which ever has been used in this worksheet, even if the content was in between deleted.
  rowMax = oWs.Cells.SpecialCells(xlCellTypeLastCell).Row
  If rowMax = 0 Then rowMax = 1 'rowMax is the number of the last visible row, if there are via grouping hidden row below, then they re not included
  rowMax0 = rowMax + 1000 'assume less than 1000 hidden rows
  If oWs.Parent.FileFormat = MyExcel8 And rowMax0 > 65535 Then rowMax0 = 65535
  'add further rows to rowMax while rowMax+1 ... rowMax0 is not total empty:
  Do While Application.WorksheetFunction.CountA(oWs.Range(oWs.Cells(rowMax + 1, 1), oWs.Cells(rowMax0, 1)).EntireRow) > 0 And rowMax < rowMax0
    rowMax = rowMax + 1
  Loop
  'oWS.UsedRange is not always working well: reduce rowMax while that row is empty:
  Do While Application.WorksheetFunction.CountA(oWs.Range(oWs.Cells(rowMax, 1), oWs.Cells(rowMax0, 1)).EntireRow) = 0
    rowMax = rowMax - 1
    If rowMax = 0 Then
      Exit Do
    End If
  Loop
  LastUsedRow = rowMax
End Function


Public Function LastUsedColumn(oWs As Object, Optional ByVal row1st As Long = 1) As Long
  'Workaround for buggy .Cells.SpecialCells(xlCellTypeLastCell).Row
  Dim colMax As Long, colMax0 As Long, rowMax As Long
  Dim col2 As Long
  'Dim n2 As Long
  Dim r1 As Long
  
  oWs.UsedRange 'Reset if some Cell contents was cleared
  colMax = oWs.Cells.SpecialCells(xlCellTypeLastCell).Column
  If colMax = 0 Then colMax = 1
  colMax0 = colMax + 1000
  If Val(Application.Version) <= 11 And colMax0 > 255 Then colMax0 = 255
  'If oWs.Parent.FileFormat = MyExcel8 And colMax0 > 255 Then colMax0 = 255
  'often this gives a too high value without the UsedRange, in case of Grouping and "hidden" Rows, also too low!!!
  If row1st = 1 Then
    Do While Application.WorksheetFunction.CountA(oWs.Range(oWs.Cells(row1st, colMax + 1), oWs.Cells(row1st, colMax0)).EntireColumn) > 0 And colMax < colMax0
      colMax = colMax + 1
    Loop
    'oWS.UsedRange is not always working well: reduce rowMax while that row is empty:
    Do While Application.WorksheetFunction.CountA(oWs.Range(oWs.Cells(1, colMax), oWs.Cells(1, colMax0)).EntireColumn) = 0
      colMax = colMax - 1
    Loop
  Else
    rowMax = LastUsedRow(oWs)
    Do While Application.WorksheetFunction.CountA(oWs.Range(oWs.Cells(row1st, colMax + 1), oWs.Cells(rowMax, colMax0)).EntireColumn) > 0 And colMax < colMax0
      colMax = colMax + 1
    Loop
    'oWS.UsedRange is not always working well: reduce rowMax while that row is empty:
    Do While Application.WorksheetFunction.CountA(oWs.Range(oWs.Cells(row1st, colMax), oWs.Cells(rowMax, colMax0)).EntireColumn) = 0
      colMax = colMax - 1
      If colMax = 0 Then
        Exit Do
      End If
    Loop
  End If
  LastUsedColumn = colMax
End Function

