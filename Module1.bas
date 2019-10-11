Attribute VB_Name = "Module1"
Dim TimerActive As Boolean
Dim DT As Variant
Dim DTT As Variant
Dim linie As Integer
Dim Coloana As Integer
Dim IntervalHH As Integer
Dim IntervalMM As Integer
Dim IntervalSS As Integer

Dim lin As Integer


Sub StartTimer()
If Not (TimerActive) Then
        linie = Worksheets("INIT").Cells(1, 2).Value
        Coloana = Worksheets("INIT").Cells(2, 2).Value
        IntervalSS = Worksheets("INIT").Cells(3, 2).Value
        
        Worksheets("INIT").Cells(1, 4).Value = "INREGISTREAZA..."
        Worksheets("DATE").Cells(linie, Coloana + 0).Value = "Data"
        Worksheets("DATE").Cells(linie, Coloana + 1).Value = "Temp"
        Worksheets("DATE").Cells(linie, Coloana + 2).Value = "Pres"
        
        Worksheets("DATE").Columns(Coloana + 0).NumberFormat = "dd.mm.yyyy h:mm:ss"
        
        IntervalHH = IntervalSS / 3600
        IntervalMM = (IntervalSS Mod 3600) / 60
        IntervalSS = (IntervalSS Mod 3600) Mod 60
        
        linie = linie + 1
        DT = Now()
        DTT = DT + TimeSerial(0, 0, 60 - (Second(DT) Mod 60))
        TimerActive = True
        Application.OnTime DTT, "Timer"
End If
End Sub

Sub ResetDate()
    TimerActive = False
    linie = 0
    Range("A10").Select
    Worksheets("INIT").Cells(1, 1).Value = "Linie:"
    Worksheets("INIT").Cells(2, 1).Value = "Coloana:"
    Worksheets("INIT").Cells(1, 2).Value = 1
    Worksheets("INIT").Cells(2, 2).Value = 1
    
    Worksheets("INIT").Cells(3, 1).Value = "Interval Secunde:"
    Worksheets("INIT").Cells(3, 2).Value = 5
    
    Worksheets("INIT").Cells(1, 4).Value = "OPRIT"
    Worksheets("DATE").Cells.Clear
    Worksheets("DATE").Rows("1:1").HorizontalAlignment = xlCenter
End Sub

Sub AutoWidth()
    Worksheets("DATE").Cells.EntireColumn.AutoFit
End Sub


Sub StopTimer()
    If TimerActive Then
         TimerActive = False
       'Worksheets("INIT").Cells(1, 2).Value = Linie
        Worksheets("INIT").Cells(2, 2).Value = Coloana + 4
        Worksheets("INIT").Cells(1, 4).Value = "OPRIT"
        ' AutoWidth
    End If
End Sub

Private Sub Timer()
DT = Now()
    If TimerActive Then
        'If (Second(DT) Mod IntervalSS) = 0 Then
        Worksheets("DATE").Cells(linie, Coloana + 0).Value = DT
        Worksheets("DATE").Cells(linie, Coloana + 1).Value = Worksheets("CAF").Cells(6, 2).Value
        Worksheets("DATE").Cells(linie, Coloana + 2).Value = Worksheets("CAF").Cells(8, 2).Value
        linie = linie + 1
        'End If
        DTT = DTT + TimeSerial(IntervalHH, IntervalMM, IntervalSS)
        Application.OnTime DTT, "Timer"
    End If
End Sub

Sub ChangeData()
'Static linie As Integer
Coloana = Worksheets("INIT").Cells(2, 2).Value
If linie = 0 Then
linie = Worksheets("INIT").Cells(1, 2).Value + 1
Worksheets("DATE").Columns(Coloana + 0).NumberFormat = "dd.mm.yyyy h:mm:ss"
End If
DT = Now()
    'If TimerActive Then
        'If (Second(DT) Mod IntervalSS) = 0 Then
        'linie = linie + 1
        Worksheets("DATE").Cells(linie, Coloana + 0).Value = DT
        Worksheets("DATE").Cells(linie, Coloana + 1).Value = Worksheets("CAF").Cells(6, 2).Value
        Worksheets("DATE").Cells(linie, Coloana + 2).Value = Worksheets("CAF").Cells(8, 2).Value
        linie = linie + 1
        'End If
        'DTT = DTT + TimeSerial(IntervalHH, IntervalMM, IntervalSS)
        'Application.OnTime DTT, "Timer"
    'End If
End Sub



Private Sub UnusedCode()
' Worksheets("TIME").Cells(1, 1).Value = Time
' Worksheets("TIME").Cells(1, 2).Value = Second(Time)
' Worksheets("TIME").Cells(1, 3).Value = Second(Time)

'Start_Timer
'DT = Now()
'For i = 0 To 10
'Worksheets("test").Cells(i + 1, 1).Value = 10 - (Second(DT) Mod 10)
'Next i
' Do Until Not (Second(Now()) Mod 5)
' Application.Wait "00:00:00.500"
' Loop

'DTT = DT + TimeValue("00:00:" & CStr(10 - (Second(DT) Mod 10)))

' Worksheets("DATE").UsedRange.ShrinkToFit = True

' Application.OnTime DT + TimeValue("00:00:01"), "Timer"
'Application.OnTime DT + TimeValue("00:00:" & CStr((DT Mod 10) + Second(DT))), "Timer"
'Application.OnTime DT + TimeValue("00:00:" & CStr(10 - (Second(DT) Mod 10))), "Timer"

' Worksheets("DATE").Range("A1").Select
'DTT = DT + TimeSerial(0, 0, 10 - (Second(DT) Mod IntervalSS))
'DTT = DT + TimeSerial(24 - (Hour(DT) Mod IntervalHH), 60 - (Minute(DT) Mod IntervalMM), 10 - (Second(DT) Mod IntervalSS))

' in sheet, on a1 change
'If Not Intersect(Target, Range("A1")) Is Nothing Then
'Target.EntireRow.Interior.ColorIndex = 15
'Worksheets("test2").Cells(1, 1).Value = "bum"
'Module1.test

'End If

End Sub

