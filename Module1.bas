Attribute VB_Name = "Module1"
 Option Explicit
 Dim DateDay As String, DateMonth As String, DateYear As String
Dim FirstSlash As Integer
Dim toCheckUrl As String
Dim convertFrom, convertTo As String
Dim StartUSD, convertFromCurrency, convertToCurrency As Range



Sub Openform()
Dim i As Integer, DateArray As Variant, TodayDate As String
Worksheets("Instructions").Select
Range("A1").Select

For i = 1 To Application.WorksheetFunction.CountA(Columns("A:A"))
    CurrencyConverter.cbxConvertFrom.AddItem ActiveCell.Offset(i - 1, 0) & " - " & ActiveCell.Offset(i - 1, 1)
    CurrencyConverter.cbxConvertTo.AddItem ActiveCell.Offset(i - 1, 0) & " - " & ActiveCell.Offset(i - 1, 1)
Next i
CurrencyConverter.cbxConvertFrom.Text = Range("A1")
CurrencyConverter.cbxConvertTo.Text = Range("A1")
CurrencyConverter.txtDate.Text = Date
With CurrencyConverter
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
End With
'CurrencyConverter.Show
End Sub

Sub PlotData()
Dim i As Integer

Dim DateDay As String, DateMonth As String, DateYear, dateOnUserForm As String
Dim FirstSlash As Integer
Dim URL As String
Dim convertFrom, convertTo As String
Dim StartUSD, convertFromCurrency, convertToCurrency As Range

With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With

Worksheets("Plots").Visible = True

For i = 1 To Sheets.Count
    If Sheets(i).Name = "Plot Chart" Then
         Sheets(i).Delete
         Exit For
    End If
Next i


Worksheets("Plots").Select
Worksheets("Plots").Cells.Clear
Range("A15").Select

' last 15 days
For i = 15 To 1 Step -1
    Worksheets("Plots").Range("A" & 15 - i + 1) = VBA.DateAdd("d", -i + 1, CurrencyConverter.txtDate.Text)
Next i

For i = 1 To 15

    FirstSlash = InStr(Range("A" & i), "/")
    DateMonth = Left(Range("A" & i), FirstSlash - 1)
    If Len(DateMonth) = 1 Then DateMonth = 0 & DateMonth
    DateDay = Mid(Range("A" & i), FirstSlash + 1, 2)
   
    If InStr(DateDay, "/") Then DateDay = 0 & Replace(DateDay, "/", "")
  
    DateYear = Right(Range("A" & i), 4)
    
    dateOnUserForm = DateYear & "-" & DateMonth & "-" & DateDay
    Call Module1.URL(dateOnUserForm)
    
    Set StartUSD = Worksheets("Sheet1").Range("A:A").Find(What:="Currency code", Lookat:=xlPart)
    
    convertFrom = CurrencyConverter.cbxConvertFrom.Value
    convertTo = CurrencyConverter.cbxConvertTo.Value
    
    Set convertFromCurrency = Worksheets("Sheet1").Range("A:A").Find(What:=Left(convertFrom, 3), After:=StartUSD, Lookat:=xlPart)
    Set convertToCurrency = Worksheets("Sheet1").Range("A:A").Find(What:=Left(convertTo, 3), After:=StartUSD, Lookat:=xlPart)

       
    Worksheets("Plots").Range("B" & i) = convertToCurrency.Offset(0, 2) / convertFromCurrency.Offset(0, 2)

Next i



Range("A1:B15").Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=Range("Plots!$A$1:$B$15")
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Plot Chart"
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    Selection.Delete

MsgBox "Plot Chart is done. Thank you for waiting!"
Worksheets("Plots").Visible = False

With Application
.DisplayAlerts = True
.ScreenUpdating = True
End With
End Sub

Function dateOnUserForm() As String
Dim DateArr() As String
    FirstSlash = InStr(CurrencyConverter.txtDate, "/")
    DateMonth = Left(CurrencyConverter.txtDate, FirstSlash - 1)
        If DateMonth > 12 Then
        MsgBox "Wrong format! Please, enter in MM/DD/YYYY"
        Exit Function
        End If
        
    If Len(DateMonth) = 1 Then DateMonth = 0 & DateMonth
    
    DateYear = Right(CurrencyConverter.txtDate, 4)
        If Len(DateYear) < 4 Then
            
             Exit Function
            ElseIf DateYear < 1999 Or DateYear > 2021 Then
  
             Exit Function
        End If
    DateArr = Split(CurrencyConverter.txtDate, "/")
    DateDay = DateArr(1)
    
    If Len(DateDay) = 1 Then DateDay = 0 & DateDay
    dateOnUserForm = DateYear & "-" & DateMonth & "-" & DateDay
 
End Function


Sub Calculation()
   Set StartUSD = Worksheets("Sheet1").Range("A:A").Find(What:="Currency code", Lookat:=xlPart)
                
                convertFrom = CurrencyConverter.cbxConvertFrom.Value
                convertTo = CurrencyConverter.cbxConvertTo.Value
                
                Set convertFromCurrency = Worksheets("Sheet1").Range("A:A").Find(What:=Left(convertFrom, 3), After:=StartUSD, Lookat:=xlPart)
                Set convertToCurrency = Worksheets("Sheet1").Range("A:A").Find(What:=Left(convertTo, 3), After:=StartUSD, Lookat:=xlPart)
                If convertFromCurrency Is Nothing Or convertToCurrency Is Nothing Then
                    MsgBox " Apology, this currency is not availible at this time period. Try enter different date?"
                    CurrencyConverter.txtAmountConverted.Value = 0
                    Exit Sub
                 End If
                CurrencyConverter.txtAmountConverted.Value = Round(CurrencyConverter.txtAmountToConvert.Value * convertToCurrency.Offset(0, 2).Value / convertFromCurrency.Offset(0, 2).Value, 2)
                Worksheets("Sheet1").Visible = False
End Sub

Sub URL(x As String)
Worksheets("Sheet1").Cells.Clear
toCheckUrl = "URL;https://www.xe.com/currencytables/?from=USD&date=" & x
                With Worksheets("Sheet1").QueryTables.Add(Connection:=toCheckUrl, Destination:=Worksheets("Sheet1").Range("A1"))
                    .Name = "My Query"
                    .RowNumbers = False
                    .FillAdjacentFormulas = False
                    .PreserveFormatting = True
                    .RefreshOnFileOpen = False
                    .BackgroundQuery = False
                    .RefreshStyle = xlOverwriteCells
                    .SavePassword = False
                    .SaveData = True
                    .AdjustColumnWidth = True
                    .RefreshPeriod = 0
                    .WebSelectionType = xlEntirePage
                    .WebFormatting = xlWebFormattingNone
                    .WebPreFormattedTextToColumns = True
                    .WebConsecutiveDelimitersAsOne = True
                    .WebSingleBlockTextImport = False
                    .WebDisableDateRecognition = False
                    .WebDisableRedirections = False
                    .Refresh BackgroundQuery:=False
                End With
End Sub
