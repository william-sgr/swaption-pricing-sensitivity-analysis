Attribute VB_Name = "Cash_IRR_fwd_prm_strategies"
Option Explicit

Public Sub Build_CashIRR_Strategies_Shift005_STRICT()

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastRow As Long

    Set wb = ThisWorkbook
    Set wsSrc = wb.Worksheets("CASH_IRR_FWD_PREM_LONG")

    '--- columns
    Dim cBlock As Long, cExp As Long, cTen As Long, cMon As Long
    Dim cShift As Long, cPrem As Long

    cBlock = FindHeaderCol(wsSrc, "SourceBlock")
    cExp = FindHeaderCol(wsSrc, "ExpiryLbl")
    cTen = FindHeaderCol(wsSrc, "TenorLbl")
    cMon = FindHeaderCol(wsSrc, "MoneynessBP")
    cShift = FindHeaderCol(wsSrc, "Shift")
    cPrem = FindHeaderCol(wsSrc, "CashIRRPremiumBP")

    If cBlock = 0 Or cExp = 0 Or cTen = 0 Or cMon = 0 Or cShift = 0 Or cPrem = 0 Then
        MsgBox "Mancano colonne. Servono: SourceBlock, ExpiryLbl, TenorLbl, MoneynessBP, Shift, CashIRRPremiumBP.", vbCritical
        Exit Sub
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, cExp).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Foglio sorgente vuoto.", vbCritical
        Exit Sub
    End If

    Dim targetShift As Double
    targetShift = 0.05   ' <<< FIX: 0,05

    '--- check existence of shift=0.05
    If Not ExistsShift(wsSrc, cShift, lastRow, targetShift) Then
        MsgBox "STOP: nel foglio non esistono righe con Shift=0,05 (0.05).", vbCritical
        Exit Sub
    End If

    Dim mons(1 To 6) As Long
    mons(1) = 50: mons(2) = 100: mons(3) = 150: mons(4) = 200: mons(5) = 300: mons(6) = 400

    ' OTM pairs (SourceBlock<>ATM)
    Dim pairs() As String, nPairs As Long
    pairs = UniquePairs_ByBlockAndMons(wsSrc, cBlock, "<>ATM", cExp, cTen, cMon, lastRow, mons, nPairs)

    If nPairs = 0 Then
        MsgBox "Non trovo righe OTM (SourceBlock<>ATM) con |MoneynessBP| in {50,100,150,200,300,400}.", vbCritical
        Exit Sub
    End If

    ' Expiry/Tenor lists for ATM matrix
    Dim expList() As String, tenList() As String
    Dim nExp As Long, nTen As Long
    expList = UniqueList_ByBlock(wsSrc, cBlock, "ATM", cExp, lastRow, nExp)
    tenList = UniqueList_ByBlock(wsSrc, cBlock, "ATM", cTen, lastRow, nTen)

    If nExp = 0 Or nTen = 0 Then
        MsgBox "Non trovo righe ATM (SourceBlock=ATM).", vbCritical
        Exit Sub
    End If

    ' Recreate output sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("CashIRR_Strategies_005").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.name = "CashIRR_Strategies_005"
    wsOut.Activate

    Dim top1 As Long, top2 As Long
    top1 = 1

    WriteStrategiesTable_STRICT wsSrc, wsOut, top1, _
        "CashIRRPremiumBP Shift=0.05", _
        cBlock, cExp, cTen, cMon, cShift, cPrem, lastRow, pairs, nPairs, mons, targetShift

    top2 = top1 + (nPairs + 3) + 4

    WriteATMStraddleMatrix_STRICT wsSrc, wsOut, top2, _
        "Straddle Shift=0.05 (Expiry x Tenor)", _
        cBlock, cExp, cTen, cMon, cShift, cPrem, lastRow, expList, tenList, nExp, nTen, targetShift

    wsOut.Columns.AutoFit
    MsgBox "Creato foglio 'CashIRR_Strategies_005' (STRICT Shift=0,05).", vbInformation

End Sub


'=========================================================
' TABLE 1: Collar | Straddle | Strangle  (STRICT Shift=0.05)
' OTM: Call = +m, Put = -m (signed match)
'=========================================================
Private Sub WriteStrategiesTable_STRICT(wsSrc As Worksheet, wsOut As Worksheet, topRow As Long, _
                                       title As String, _
                                       cBlock As Long, cExp As Long, cTen As Long, cMon As Long, cShift As Long, cPrem As Long, _
                                       lastRow As Long, pairs() As String, nPairs As Long, mons() As Long, targetShift As Double)

    Dim i As Long, j As Long
    Dim exp As String, ten As String
    Dim callV As Variant, putV As Variant, atmV As Variant

    wsOut.Cells(topRow, 1).Value = title
    wsOut.Cells(topRow, 1).Font.Bold = True

    ' Header row 1
    wsOut.Range(wsOut.Cells(topRow + 1, 2), wsOut.Cells(topRow + 1, 7)).Merge
    wsOut.Cells(topRow + 1, 2).Value = "Collar"
    wsOut.Cells(topRow + 1, 2).HorizontalAlignment = xlCenter
    wsOut.Cells(topRow + 1, 2).Font.Bold = True

    wsOut.Cells(topRow + 1, 8).Value = "Straddle"
    wsOut.Cells(topRow + 1, 8).HorizontalAlignment = xlCenter
    wsOut.Cells(topRow + 1, 8).Font.Bold = True

    wsOut.Range(wsOut.Cells(topRow + 1, 9), wsOut.Cells(topRow + 1, 14)).Merge
    wsOut.Cells(topRow + 1, 9).Value = "Strangle"
    wsOut.Cells(topRow + 1, 9).HorizontalAlignment = xlCenter
    wsOut.Cells(topRow + 1, 9).Font.Bold = True

    ' Header row 2
    wsOut.Cells(topRow + 2, 1).Value = "ExpiryTenor"
    wsOut.Cells(topRow + 2, 1).Font.Bold = True

    For j = 1 To 6
        wsOut.Cells(topRow + 2, 1 + j).Value = mons(j)
        wsOut.Cells(topRow + 2, 1 + j).Font.Bold = True
    Next j

    wsOut.Cells(topRow + 2, 8).Value = "ATM"
    wsOut.Cells(topRow + 2, 8).Font.Bold = True

    For j = 1 To 6
        wsOut.Cells(topRow + 2, 8 + j).Value = mons(j)
        wsOut.Cells(topRow + 2, 8 + j).Font.Bold = True
    Next j

    ' Body
    For i = 1 To nPairs
        exp = Split(pairs(i), "|")(0)
        ten = Split(pairs(i), "|")(1)

        wsOut.Cells(topRow + 2 + i, 1).Value = exp & ten

        ' ATM Straddle = 2 * Premium (ATM, mon=0, shift=0.05)
        atmV = SumPrem_STRICT(wsSrc, cBlock, "ATM", cExp, exp, cTen, ten, cMon, 0, cShift, targetShift, cPrem, lastRow)
        If IsEmpty(atmV) Then
            wsOut.Cells(topRow + 2 + i, 8).Value = ""
        Else
            wsOut.Cells(topRow + 2 + i, 8).Value = 2# * CDbl(atmV)
        End If

        ' OTM
        For j = 1 To 6
            callV = SumPrem_STRICT(wsSrc, cBlock, "<>ATM", cExp, exp, cTen, ten, cMon, mons(j), cShift, targetShift, cPrem, lastRow)     ' +m
            putV = SumPrem_STRICT(wsSrc, cBlock, "<>ATM", cExp, exp, cTen, ten, cMon, -mons(j), cShift, targetShift, cPrem, lastRow)     ' -m

            If IsEmpty(callV) Or IsEmpty(putV) Then
                wsOut.Cells(topRow + 2 + i, 1 + j).Value = ""
                wsOut.Cells(topRow + 2 + i, 8 + j).Value = ""
            Else
                wsOut.Cells(topRow + 2 + i, 1 + j).Value = CDbl(callV) - CDbl(putV)
                wsOut.Cells(topRow + 2 + i, 8 + j).Value = CDbl(callV) + CDbl(putV)
            End If
        Next j
    Next i

End Sub


'=========================================================
' TABLE 2: ATM Straddle matrix (2*Premium) STRICT Shift=0.05
'=========================================================
Private Sub WriteATMStraddleMatrix_STRICT(wsSrc As Worksheet, wsOut As Worksheet, topRow As Long, _
                                         title As String, _
                                         cBlock As Long, cExp As Long, cTen As Long, cMon As Long, cShift As Long, cPrem As Long, _
                                         lastRow As Long, expList() As String, tenList() As String, _
                                         nExp As Long, nTen As Long, targetShift As Double)

    Dim i As Long, j As Long
    Dim atmV As Variant

    wsOut.Cells(topRow, 1).Value = title
    wsOut.Cells(topRow, 1).Font.Bold = True

    wsOut.Cells(topRow + 1, 1).Value = "ExpiryLbl"
    wsOut.Cells(topRow + 1, 1).Font.Bold = True

    For j = 1 To nTen
        wsOut.Cells(topRow + 1, j + 1).Value = tenList(j)
        wsOut.Cells(topRow + 1, j + 1).Font.Bold = True
    Next j

    For i = 1 To nExp
        wsOut.Cells(topRow + 1 + i, 1).Value = expList(i)
        For j = 1 To nTen
            atmV = SumPrem_STRICT(wsSrc, cBlock, "ATM", cExp, expList(i), cTen, tenList(j), cMon, 0, cShift, targetShift, cPrem, lastRow)
            If IsEmpty(atmV) Then
                wsOut.Cells(topRow + 1 + i, j + 1).Value = ""
            Else
                wsOut.Cells(topRow + 1 + i, j + 1).Value = 2# * CDbl(atmV)
            End If
        Next j
    Next i

End Sub


'=========================================================
' Sum premium STRICT (exact shift only)
' Parses numbers with comma correctly.
'=========================================================
Private Function SumPrem_STRICT(ws As Worksheet, _
                               cBlock As Long, blockCriteria As String, _
                               cExp As Long, exp As String, _
                               cTen As Long, ten As String, _
                               cMon As Long, monSigned As Long, _
                               cShift As Long, targetShift As Double, _
                               cPrem As Long, lastRow As Long) As Variant

    Dim r As Long
    Dim sb As String
    Dim M As Long
    Dim sh As Double, prem As Double
    Dim sumV As Double
    Dim found As Boolean

    sumV = 0#
    found = False

    For r = 2 To lastRow

        sb = Trim$(CStr(ws.Cells(r, cBlock).Value))
        If Not BlockMatch(sb, blockCriteria) Then GoTo NextR

        If Trim$(CStr(ws.Cells(r, cExp).Value)) <> exp Then GoTo NextR
        If Trim$(CStr(ws.Cells(r, cTen).Value)) <> ten Then GoTo NextR

        M = CLng(val(ws.Cells(r, cMon).Value))
        If M <> monSigned Then GoTo NextR

        If Not TryParseDouble(ws.Cells(r, cShift).Value, sh) Then GoTo NextR
        If Abs(sh - targetShift) > 0.0000000001 Then GoTo NextR

        If Not TryParseDouble(ws.Cells(r, cPrem).Value, prem) Then GoTo NextR

        sumV = sumV + prem
        found = True

NextR:
    Next r

    If found Then
        SumPrem_STRICT = sumV
    Else
        SumPrem_STRICT = Empty
    End If

End Function

Private Function ExistsShift(ws As Worksheet, cShift As Long, lastRow As Long, targetShift As Double) As Boolean
    Dim r As Long, sh As Double
    For r = 2 To lastRow
        If TryParseDouble(ws.Cells(r, cShift).Value, sh) Then
            If Abs(sh - targetShift) < 0.0000000001 Then
                ExistsShift = True
                Exit Function
            End If
        End If
    Next r
    ExistsShift = False
End Function

Private Function FindHeaderCol(ws As Worksheet, headerName As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerName Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Private Function UniquePairs_ByBlockAndMons(ws As Worksheet, _
                                           cBlock As Long, blockCriteria As String, _
                                           cExp As Long, cTen As Long, cMon As Long, _
                                           lastRow As Long, mons() As Long, ByRef nOut As Long) As String()
    Dim tmp() As String
    ReDim tmp(1 To lastRow)
    nOut = 0

    Dim r As Long, sb As String, exp As String, ten As String, key As String
    Dim M As Long, absM As Long

    For r = 2 To lastRow
        sb = Trim$(CStr(ws.Cells(r, cBlock).Value))
        If BlockMatch(sb, blockCriteria) Then
            M = CLng(val(ws.Cells(r, cMon).Value))
            absM = Abs(M)
            If M <> 0 And InMons(absM, mons) Then
                exp = Trim$(CStr(ws.Cells(r, cExp).Value))
                ten = Trim$(CStr(ws.Cells(r, cTen).Value))
                If Len(exp) > 0 And Len(ten) > 0 Then
                    key = exp & "|" & ten
                    If Not InArray(key, tmp, nOut) Then
                        nOut = nOut + 1
                        tmp(nOut) = key
                    End If
                End If
            End If
        End If
    Next r

    If nOut = 0 Then
        ReDim UniquePairs_ByBlockAndMons(1 To 1)
        Exit Function
    End If

    ReDim Preserve tmp(1 To nOut)
    UniquePairs_ByBlockAndMons = tmp
End Function

Private Function UniqueList_ByBlock(ws As Worksheet, cBlock As Long, blockValue As String, _
                                   targetCol As Long, lastRow As Long, ByRef nOut As Long) As String()
    Dim tmp() As String
    ReDim tmp(1 To lastRow)
    nOut = 0

    Dim r As Long, sb As String, v As String
    For r = 2 To lastRow
        sb = Trim$(CStr(ws.Cells(r, cBlock).Value))
        If UCase$(sb) = UCase$(blockValue) Then
            v = Trim$(CStr(ws.Cells(r, targetCol).Value))
            If Len(v) > 0 Then
                If Not InArray(v, tmp, nOut) Then
                    nOut = nOut + 1
                    tmp(nOut) = v
                End If
            End If
        End If
    Next r

    If nOut = 0 Then
        ReDim UniqueList_ByBlock(1 To 1)
        Exit Function
    End If

    ReDim Preserve tmp(1 To nOut)
    UniqueList_ByBlock = tmp
End Function

Private Function InMons(absMon As Long, mons() As Long) As Boolean
    Dim k As Long
    For k = LBound(mons) To UBound(mons)
        If absMon = mons(k) Then
            InMons = True
            Exit Function
        End If
    Next k
    InMons = False
End Function

Private Function InArray(val As String, arr() As String, n As Long) As Boolean
    Dim i As Long
    For i = 1 To n
        If arr(i) = val Then
            InArray = True
            Exit Function
        End If
    Next i
    InArray = False
End Function

Private Function BlockMatch(sourceBlockValue As String, criteria As String) As Boolean
    If criteria = "ATM" Then
        BlockMatch = (UCase$(sourceBlockValue) = "ATM")
    ElseIf criteria = "<>ATM" Then
        BlockMatch = (UCase$(sourceBlockValue) <> "ATM" And Len(sourceBlockValue) > 0)
    Else
        BlockMatch = (sourceBlockValue = criteria)
    End If
End Function

Private Function TryParseDouble(v As Variant, ByRef outD As Double) As Boolean
    On Error GoTo fail
    If IsNumeric(v) Then
        outD = CDbl(v)
        TryParseDouble = True
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then GoTo fail
    s = Replace(s, ",", ".")
    If IsNumeric(s) Then
        outD = CDbl(s)
        TryParseDouble = True
        Exit Function
    End If

fail:
    outD = 0#
    TryParseDouble = False
End Function

