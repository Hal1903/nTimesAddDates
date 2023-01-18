Attribute VB_Name = "FromHaruku"
'
' If you do not understand what those opening file codes below do,
' please refer to HW3.2. I explained about them step by step.
'
' Opening Files
'

Sub openFile(path As String)
'H:\MFG Reports\Production Analysis SheetBuilding1\ETHR QAS PAS\auto 3\
    'declaring the variable/object and its type
    Dim wb As Workbook
'setting the object and opens the file
    Set wb = Workbooks.Open(path)
'selecting "Summary" sheet
    Sheets("Summary").Select
End Sub
Public Function monthLetters() As String
    Dim DD As Date
    Dim MM As Integer
    DD = Now
    'Debug.Print MonthName(month(DD), True) 'MonthName(MM) 'UCase(Left(MonthName(MM), 3))
    monthLetters = UCase(Left(MonthName(month(DD)), 3))
End Function
Public Function EGRmonthNum() As String
    Dim d As Date
    d = Now
    m = Format(d, "MM")
    EGRmonthNum = m
End Function

Public Function YearNum() As String
    Dim d As Date
    d = Now
    YearNum = Year(d)
End Function
Public Function YearNum2() As String
    Dim d As Date
    d = Now
    YearNum2 = Right(Year(d), 2)
End Function
Sub testY()
    Debug.Print monthLetters()
End Sub
Public Function stFile()
    stFile = "H:\MFG Reports\Production Analysis SheetBuilding1\ETHR QAS PAS\"
End Function
Sub auto3thr_Click()
    Dim path As String
    path = stFile & "auto 3\6QAS PAS -MM5L-   " & monthLetters & YearNum2() & "(AUTO3) ET42 - MASS.xlsm"
    openFile (path)
End Sub
Sub auto4_zh2k1_Click()
    Dim path As String
    path = stFile & "auto 4\6SSQAS PAS -  " & monthLetters & YearNum2() & " -TRIAL (AUTO4) --ZH2k1 ET7 only.xlsm"
    Call openFile(path)
End Sub
Sub auto4_tr2k3_Click()
    openFile (stFile & "\auto 4\6SSQAS PAS -  " & monthLetters & YearNum2() & " -TRIAL (AUTO4) --TR2k3 ET7 only.xlsm")
End Sub
Sub auto5_kh5t_Click()
    openFile (stFile & "\auto 5\6QAS PAS -  " & monthLetters & YearNum2() & " -TRIAL (AUTO5) --KH5T only.xlsm")
End Sub
Sub auto7_tgna_Click()
    openFile (stFile & "\auto 7\6QAS PAS -MM5L-   " & monthLetters & YearNum2() & "(AUTO7) TGNA.xlsm")
End Sub
Sub auto6_Click()
    openFile (stFile & "\auto 6\6QAS PAS -MM5L-   " & monthLetters & YearNum2() & "(AUTO6) TGNA.xlsm")
End Sub
Public Function EGRstFile()
    EGRstFile = "H:\MFG Reports\Production Analysis SheetBuilding1\EGR QASPAS\2022 EGR"
End Function
Sub EGRV1_open()
    ' Call openFile ("H:\MFG Reports\Production Analysis SheetBuilding1\EGR QASPAS\2022 EGR\07JUL2022 EGR\6SSQASPAS MM10L - EGR Assy L01.xlsm")
    Call openFile("H:\MFG Reports\Production Analysis SheetBuilding1\EGR QASPAS\2022 EGR\" & EGRmonthNum() & monthLetters & YearNum() & " EGR\6SSQASPAS MM10L - EGR Assy L01.xlsm")
End Sub
Sub EGRV2_open()
    ' Call openFile("H:\MFG Reports\Production Analysis SheetBuilding1\EGR QASPAS\2022 EGR\07JUL2022 EGR\6QASPAS MM10L - EGR Assy L02.xlsm")
    Call openFile("H:\MFG Reports\Production Analysis SheetBuilding1\EGR QASPAS\2022 EGR\" & EGRmonthNum() & monthLetters & YearNum() & " EGR\6QASPAS MM10L - EGR Assy L02.xlsm")
End Sub
Sub openAllFiles()
    Call auto3thr_Click
    Call auto4_zh2k1_Click
    Call auto5_kh5t_Click
    Call auto7_tgna_Click
    Call auto6_Click
    Call EGRV1_open
    Call EGRV2_open
End Sub

'
' Adding Dates Except For Weekend
'

Sub nthAddDate()
'take an input of date and how many dates to add
    Dim start As Date
    Dim nth As Integer
'setting value from the cells for start date and how many dates to add
    start = Format(Range("F1").Value, "d-mmm")
    'nth = Range("AU2").Value + 1
    nth = Range("F2").Value + 1
    Dim count As Integer
'increment for dates from yesterday, needed to generate the dates for tomorrow, day after tomorrow, and so on.
    Dim incr As Integer
'container of next 5 days
    ReDim dayArray(nth) As Integer
    incr = 1
    ReDim dateArray(nth) As Date
'array index that increases after successful addition to the array
    count = 1
    Dim st As String
    st = ""
'loop for completing the array
    While count < nth
        Dim dat As Date
        dat = start + incr
        wday_d = Weekday(dat)
        If wday_d <> 1 And wday_d <> 7 Then
            dateArray(count) = dat
            dayArray(count) = wday_d
            count = count + 1
            st = st & dat & ", "
        End If
        incr = incr + 1
    Wend
    Debug.Print (st)
'increment to move one right column and fill the date
    Dim columnIncr As Integer
    Dim c As Integer
    c = 0
    For Each cell In ActiveSheet.UsedRange
        columnIncr = 1
'conditional statement determining the variable type of a cell value
'also determining the cell value is equal to yesterday's date
        If VarType(cell) = 7 And Format(cell.Value, "m/d/yyyy") = start Then
'Filling the column by dates
            For d = 1 To nth - 1
                cell.Offset(0, columnIncr).NumberFormat = "d-mmm"
                cell.Offset(0, columnIncr).Value = Format(dateArray(d), "d-mmm")
                columnIncr = columnIncr + 1
            Next d
        End If
    Next cell
End Sub
'
' Changing all charts' ranges
'
Sub ChangeChartRange()
    Dim i As Integer, r As Integer, n As Integer, p1 As Integer, p2 As Integer, p3 As Integer
    Dim rng As Range
    Dim ax As Range
    Dim incrBy As Integer
    incrBy = Range("P2").Value
    For Each ch In ActiveSheet.ChartObjects
        ch.Activate
        'Cycles through each series
        For n = 1 To ActiveChart.SeriesCollection.count Step 1
            r = 0
    
            'Finds the current range of the series and the axis
            For i = 1 To Len(ActiveChart.SeriesCollection(n).Formula) Step 1
                If Mid(ActiveChart.SeriesCollection(n).Formula, i, 1) = "," Then
                    r = r + 1
                    If r = 1 Then p1 = i + 1
                    If r = 2 Then p2 = i
                    If r = 3 Then p3 = i
                End If
            Next i
    
            ' Debug.Print r & ", " & p1 & ", " & p2 & ", " & p3
            'Defines new range
            Set rng = Range(Mid(ActiveChart.SeriesCollection(n).Formula, p2 + 1, p3 - p2 - 1))
            Set rng = Range(rng, rng.Offset(0, incrBy))
    
            'Sets new range for each series
            ActiveChart.SeriesCollection(n).Values = rng
    
            'Updates axis
            Debug.Print p1 & ", " & p2 & ", " & p3
            Debug.Print Mid(ActiveChart.SeriesCollection(n).Formula, p1, p2 - p1)
            Set ax = Range(Mid(ActiveChart.SeriesCollection(n).Formula, p1, p2 - p1))
            Set ax = Range(ax, ax.Offset(0, incrBy))
            ActiveChart.SeriesCollection(n).XValues = ax
    
        Next n
    Next ch
End Sub
'
' Automating autofills for all the tables
'
Sub autoFill()
    ' F1: start date, F2: nth
    Dim start As Date
    Dim nth, incr As Integer
    Dim production As Double
    Dim emp
    Dim bl As Boolean
    bl = False
    'Sheets("Sheet1 (3)").Activate
    start = Range("F1").Value
    nth = Range("F2").Value
    ' finding target date by iterating all the cells
    For Each cell In ActiveSheet.UsedRange
        If VarType(cell) = 10 Then
            ' MsgBox
            Debug.Print "Error cell found at " & cell.Address & vbNewLine & "If error found many times, consider stop program"
        ElseIf cell = start Then
            Range(cell.Offset(1, 0).Address & ":" & cell.Offset(3, 0).Address).autoFill Destination:=Range(cell.Offset(1, 0).Address & ":" & cell.Offset(3, nth).Address), Type:=xlFillDefault
'If production is 0
            For i = 0 To nth
' if production was blank, not even 0, set production to 0
                If VarType(Range(cell.Address).Offset(1, i).Value) = 8 Then
                    production = 0
                Else
                    production = Range(cell.Address).Offset(1, i).Value
                End If
' Also deletes when production was 0.
'                If production = 0 Then
'                    Range(cell.Offset(0, i).Address & ":" & cell.Offset(0, i).End(xlDown).Address).Clear
'                    bl = True
'                End If
            Next i
        End If
    Next cell
    If bl Then
        MsgBox "One or more table had production 0. Please check the tables; delete the column and move your data, or change your column reference of its formula"
    End If
    
End Sub
Sub updateColumn()
Attribute updateColumn.VB_ProcData.VB_Invoke_Func = "w\n14"
    Dim c, a, s, p As String
    Dim dol, exc As Long
    a = ActiveCell.Formula
    s = "!" & Range("AA1").Value
    p = ""
    dol = InStr(a, "$")
    exc = InStr(a, "!")
    p = Mid(a, exc, dol - exc)
    a = Replace(a, p, s)
    ActiveCell.Formula = a
End Sub
