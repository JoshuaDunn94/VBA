Attribute VB_Name = "Module1"
Option Explicit

Sub UpdateQuery()

Application.Calculation = xlCalculationManual

StopSub = False

'1: check cell A1 contains 1, 2 or 3
If Not Range("A1") = 1 Then
    If Not Range("A1") = 2 Then
    If Not Range("A1") = 3 Then MsgBox ("Error: Cell A1 should contain '1' for JWD1, '2' for JWD2 or '3' for JWD3")
End If
End If
JWDPeriod = Range("A1").Value

'2: Clear contents, check there's some DOPs in the list
Worksheets("JWD1 Results").Range("A2:BV1000000").ClearContents
Worksheets("JWD2 Results").Range("A2:CM1000000").ClearContents
Worksheets("JWD3 Results").Range("A2:CA1000000").ClearContents
If IsEmpty(Worksheets("List of DOPs").Range("B5")) Then End

'3:
KeepPowerOn

ErrorTotal = 0
TotalMeasures = WorksheetFunction.CountA(Worksheets("List of DOPs").Range("B5:B1000000"))

'4: Start loop
For Counter = 1 To TotalMeasures

'5: Read DOP and enter it into relevant sheet
    CurrentDOP = Worksheets("List of DOPs").Range("B4").Offset(Counter, 0).Value
    If JWDPeriod = 1 Then
        Worksheets("LEW Pivot & Workings").Range("F2").Value = CurrentDOP
    End If

    If JWDPeriod = 2 Then
        Worksheets("LEW Pivot & Workings").Range("F23").Value = CurrentDOP
    End If
    
    If JWDPeriod = 3 Then
        Worksheets("LEW Pivot & Workings").Range("F45").Value = CurrentDOP
    End If


'6: Calculate workings sheet
    Worksheets("LEW Pivot & Workings").Calculate
    
    DoEvents

'7: Update relevant table with new query
If JWDPeriod = 1 Then
    On Error GoTo ErrorFound
        With Worksheets("LEW Returns").Range("A3").ListObject.QueryTable
            .Connection = Array( _
            "OLEDB;Provider=MSOLAP.4;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=JWD;Data Source=lonp-JWDBe01;MDX Compatibi" _
            , "lity=1;Safety Options=2;MDX Missing Member Mode=Error")
            .CommandType = xlCmdDefault
            .BackgroundQuery = False
            .CommandText = Worksheets("LEW Pivot & Workings").Range("F8").Value
            .Refresh
        End With
End If

If JWDPeriod = 2 Then
    On Error GoTo ErrorFound
        With Worksheets("JWD2 Returns").Range("A3").ListObject.QueryTable
            .Connection = Array( _
            "Provider=MSOLAP.4;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=JWD2;Data Source=lonp-JWDBE01;MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error")
            .CommandType = xlCmdDefault
            .BackgroundQuery = False
            .CommandText = Worksheets("LEW Pivot & Workings").Range("F32").Value
            .Refresh
        End With
End If

If JWDPeriod = 3 Then
    On Error GoTo ErrorFound
        With Worksheets("JWD3 Returns").Range("A3").ListObject.QueryTable
            .Connection = Array( _
            "Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=JWD3;Data Source=####;MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Update Isolation Level=2")
            .CommandType = xlCmdDefault
            .BackgroundQuery = False
            .CommandText = Worksheets("LEW Pivot & Workings").Range("F50").Value
            .Refresh
        End With
End If
        
    DoEvents

    PasteRow = Counter + 1
    
    DoEvents
    
'8: Copy results out
    If JWDPeriod = 1 Then
        ResultsPaste = "A" & PasteRow & ":BV" & PasteRow
        Worksheets("LEW Returns").Range("A4:BV4").Copy
        Worksheets("JWD1 Results").Range(ResultsPaste).PasteSpecial xlPasteValues
        Worksheets("JWD1 Results").Visible = True
        Worksheets("JWD1 Results").Activate
    End If
    
    If JWDPeriod = 2 Then
        ResultsPaste = "A" & PasteRow & ":CM" & PasteRow
        Worksheets("JWD2 Returns").Range("A4:CM4").Copy
        Worksheets("JWD2 Results").Range(ResultsPaste).PasteSpecial xlPasteValues
        Worksheets("JWD2 Results").Visible = True
        Worksheets("JWD2 Results").Activate
    End If
    
    If JWDPeriod = 3 Then
        ResultsPaste = "A" & PasteRow & ":BX" & PasteRow
        Worksheets("JWD3 Returns").Range("A4:CA4").Copy
        Worksheets("JWD3 Results").Range(ResultsPaste).PasteSpecial xlPasteValues
        Worksheets("JWD3 Results").Visible = True
        Worksheets("JWD3 Results").Activate
    End If

    DoEvents
    GoTo GoNext
    
'9: Error handling
ErrorFound:
    ErrorTotal = ErrorTotal + 1
    PasteRow = Counter + 1
    
    If JWDPeriod = 1 Then
        ResultsPaste = "A" & PasteRow & ":BV" & PasteRow
        Worksheets("JWD1 Results").Range(ResultsPaste).Value = "NOT FOUND"
        Worksheets("JWD1 Results").Range("D" & PasteRow).Value = Worksheets("List of DOPs").Range("B4").Offset(Counter, 0).Value
        Worksheets("JWD1 Results").Visible = True
        Worksheets("JWD1 Results").Activate
    End If
    
    If JWDPeriod = 2 Then
        ResultsPaste = "A" & PasteRow & ":CM" & PasteRow
        Worksheets("JWD2 Results").Range(ResultsPaste).Value = "NOT FOUND"
        Worksheets("JWD2 Results").Range("D" & PasteRow).Value = Worksheets("List of DOPs").Range("B4").Offset(Counter, 0).Value
        Worksheets("JWD2 Results").Visible = True
        Worksheets("JWD2 Results").Activate
    End If
    
    If JWDPeriod = 3 Then
        ResultsPaste = "A" & PasteRow & ":CA" & PasteRow
        Worksheets("JWD3 Results").Range(ResultsPaste).Value = "NOT FOUND"
        Worksheets("JWD3 Results").Range("D" & PasteRow).Value = Worksheets("List of DOPs").Range("B4").Offset(Counter, 0).Value
        Worksheets("JWD3 Results").Visible = True
        Worksheets("JWD3 Results").Activate
    End If
    
    Resume GoNext

'10 Update progress indicator
GoNext:

Progress.MeasuresRemaining = TotalMeasures - Counter
Progress.PercentComplete2 = (Counter / TotalMeasures) * 100
Progress.Repaint
If StopSub = True Then GoTo EndofSub

Next Counter

EndofSub:
If ErrorTotal > 0 Then MsgBox (ErrorTotal & " DOPs not found in LEW, marked as NOT FOUND")
StopTimer
Progress.Hide

End Sub

