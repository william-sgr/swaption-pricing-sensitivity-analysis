Attribute VB_Name = "Sensitivities_strategies_ANA"
Option Explicit

' ==========================================================
' ALL-IN-ONE (ANALYTIC) – NO ANNUITY BLOCK PRINTED
' - Reads SENS_LONG_005
' - Produces ONE sheet named: "Sensitivities Analitical"
'
'   (ATM) 3 blocks:
'     1) DeltaPrice_per1bp (ATM STRADDLE) = 2*DeltaCall - 1e-4*Annuity
'     2) VegaPrice_per1pct (ATM STRADDLE) = 2*VegaCall
'     3) ParDeltaPrice_per1bp (ATM)       = as-is
'
'   (OTM) 3 blocks:
'     4) DeltaPrice_per1bp (OTM) - Collars/Strangles
'     5) VegaPrice_per1pct (OTM) - Collars/Strangles
'     6) ParDeltaPrice_per1bp (OTM) - Collars/Strangles
'
' NOTE:
' - Annuity is STILL used inside ATM STRADDLE Delta formula,
'   but NOT printed as a standalone block.
' ==========================================================

Public Sub Build_Sensitivities_Analitical()

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastRow As Long

    Set wb = ThisWorkbook
    Set wsSrc = wb.Worksheets("SENS_LONG_005")

    ' --------- Columns
    Dim cBlock As Long, cExp As Long, cTen As Long, cMon As Long
    Dim cDelta As Long, cVega As Long, cPar As Long, cAnn As Long

    cBlock = FindHeaderCol(wsSrc, "SourceBlock")
    cExp = FindHeaderCol(wsSrc, "ExpiryLbl")
    cTen = FindHeaderCol(wsSrc, "TenorLbl")
    cMon = FindHeaderCol(wsSrc, "MoneynessBP")

    cDelta = FindHeaderCol(wsSrc, "DeltaPrice_per1bp")    ' CALL leg for ATM; signed for OTM
    cVega = FindHeaderCol(wsSrc, "VegaPrice_per1pct")     ' CALL leg for ATM; signed for OTM

    cPar = FindHeaderCol(wsSrc, "ParDeltaPrice_per1bp")
    If cPar = 0 Then cPar = FindHeaderCol(wsSrc, "ParDeltaAnn_per1bp") ' fallback

    cAnn = FindHeaderCol(wsSrc, "Annuity_Te") ' used only in STRADDLE delta formula

    If cExp = 0 Or cTen = 0 Or cDelta = 0 Or cVega = 0 Or cPar = 0 Then
        MsgBox "STOP: mancano colonne nel foglio SENS_LONG_005 (ExpiryLbl, TenorLbl, DeltaPrice_per1bp, VegaPrice_per1pct, ParDeltaPrice_per1bp/ParDeltaAnn_per1bp).", vbCritical
        Exit Sub
    End If
    If cBlock = 0 Or cAnn = 0 Then
        MsgBox "STOP: mancano colonne per ATM STRADDLE (SourceBlock e/o Annuity_Te).", vbCritical
        Exit Sub
    End If
    If cMon = 0 Then
        MsgBox "STOP: manca colonna MoneynessBP (necessaria per i blocchi OTM).", vbCritical
        Exit Sub
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, cExp).End(xlUp).Row

    ' --------- ATM Expiry/Tenor lists
    Dim expList() As String, tenList() As String
    Dim nExp As Long, nTen As Long

    expList = UniqueListFiltered(wsSrc, cBlock, "ATM", cExp, lastRow, nExp)
    tenList = UniqueListFiltered(wsSrc, cBlock, "ATM", cTen, lastRow, nTen)

    If nExp = 0 Or nTen = 0 Then
        MsgBox "Nessun Expiry/Tenor ATM trovato (SourceBlock='ATM').", vbCritical
        Exit Sub
    End If

    ' --------- OTM pairs
    Dim mons(1 To 6) As Long
    mons(1) = 50: mons(2) = 100: mons(3) = 150: mons(4) = 200: mons(5) = 300: mons(6) = 400

    Dim pairs() As String, nPairs As Long
    pairs = UniquePairsFromMoneyness(wsSrc, cExp, cTen, cMon, lastRow, mons, nPairs)

    If nPairs = 0 Then
        MsgBox "Non trovo righe OTM con |MoneynessBP| in {50,100,150,200,300,400}.", vbCritical
        Exit Sub
    End If

    ' --------- Recreate output sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Sensitivities Analitical").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.name = "SENS_STRAT_ANA"
    wsOut.Activate

    ' --------- Stack blocks (NO Annuity block)
    Dim top As Long
    top = 1

    ' ATM 1: STRADDLE Delta
    WriteBlock_STRADDLE_Delta wsSrc, wsOut, top, _
        "DeltaPrice_per1bp (ATM STRADDLE)", _
        cBlock, cExp, cTen, cDelta, cAnn, expList, tenList, nExp, nTen, lastRow
    top = top + (nExp + 2) + 2

    ' ATM 2: STRADDLE Vega
    WriteBlock_STRADDLE_Vega wsSrc, wsOut, top, _
        "VegaPrice_per1pct (ATM STRADDLE)", _
        cBlock, cExp, cTen, cVega, expList, tenList, nExp, nTen, lastRow
    top = top + (nExp + 2) + 2

    ' ATM 3: ParDelta
    WriteBlock_ATM wsSrc, wsOut, top, _
        "ParDeltaPrice_per1bp (ATM)", _
        cBlock, cExp, cTen, cPar, expList, tenList, nExp, nTen, lastRow
    top = top + (nExp + 2) + 3  ' gap before OTM

    ' OTM 4: Delta
    WriteStrategyBlock wsSrc, wsOut, top, _
        "DeltaPrice_per1bp (OTM) - Collars/Strangles", _
        cExp, cTen, cMon, cDelta, lastRow, pairs, nPairs, mons
    top = top + (nPairs + 2) + 3

    ' OTM 5: Vega
    WriteStrategyBlock wsSrc, wsOut, top, _
        "VegaPrice_per1pct (OTM) - Collars/Strangles", _
        cExp, cTen, cMon, cVega, lastRow, pairs, nPairs, mons
    top = top + (nPairs + 2) + 3

    ' OTM 6: ParDelta
    WriteStrategyBlock wsSrc, wsOut, top, _
        "ParDeltaPrice_per1bp (OTM) - Collars/Strangles", _
        cExp, cTen, cMon, cPar, lastRow, pairs, nPairs, mons

    wsOut.Columns.AutoFit
    MsgBox "Creato foglio 'Sensitivities Analitical' (senza blocco Annuity).", vbInformation

End Sub

' ==========================================================
' ATM writers
' ==========================================================

' STRADDLE Delta: 2*Delta_call - 1e-4*Annuity
Private Sub WriteBlock_STRADDLE_Delta(wsSrc As Worksheet, wsOut As Worksheet, topRow As Long, _
                                      title As String, _
                                      cBlock As Long, cExp As Long, cTen As Long, _
                                      cDeltaCall As Long, cAnn As Long, _
                                      expList() As String, tenList() As String, nExp As Long, nTen As Long, _
                                      lastRow As Long)

    Dim i As Long, j As Long
    Dim dCall As Double, a As Double

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
            dCall = SumIfsSafe(wsSrc, cDeltaCall, cBlock, "ATM", cExp, expList(i), cTen, tenList(j), lastRow)
            a = SumIfsSafe(wsSrc, cAnn, cBlock, "ATM", cExp, expList(i), cTen, tenList(j), lastRow)
            wsOut.Cells(topRow + 1 + i, j + 1).Value = 2# * dCall - 0.0001 * a
        Next j
    Next i
End Sub

' STRADDLE Vega: 2*Vega_call
Private Sub WriteBlock_STRADDLE_Vega(wsSrc As Worksheet, wsOut As Worksheet, topRow As Long, _
                                     title As String, _
                                     cBlock As Long, cExp As Long, cTen As Long, _
                                     cVegaCall As Long, _
                                     expList() As String, tenList() As String, nExp As Long, nTen As Long, _
                                     lastRow As Long)

    Dim i As Long, j As Long
    Dim vCall As Double

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
            vCall = SumIfsSafe(wsSrc, cVegaCall, cBlock, "ATM", cExp, expList(i), cTen, tenList(j), lastRow)
            wsOut.Cells(topRow + 1 + i, j + 1).Value = 2# * vCall
        Next j
    Next i
End Sub

' ATM generic block (filters SourceBlock="ATM")
Private Sub WriteBlock_ATM(wsSrc As Worksheet, wsOut As Worksheet, topRow As Long, _
                           title As String, _
                           cBlock As Long, cExp As Long, cTen As Long, cVal As Long, _
                           expList() As String, tenList() As String, nExp As Long, nTen As Long, _
                           lastRow As Long)

    Dim i As Long, j As Long

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
            wsOut.Cells(topRow + 1 + i, j + 1).Value = _
                SumIfsSafe(wsSrc, cVal, cBlock, "ATM", cExp, expList(i), cTen, tenList(j), lastRow)
        Next j
    Next i
End Sub

' ==========================================================
' OTM writer (Collars / Strangles)
' ==========================================================

Private Sub WriteStrategyBlock(wsSrc As Worksheet, wsOut As Worksheet, topRow As Long, _
                               title As String, _
                               cExp As Long, cTen As Long, cMon As Long, cVal As Long, _
                               lastRow As Long, pairs() As String, nPairs As Long, mons() As Long)

    Dim j As Long, i As Long
    Dim exp As String, ten As String
    Dim callV As Double, putV As Double

    wsOut.Cells(topRow, 1).Value = title
    wsOut.Cells(topRow, 1).Font.Bold = True

    wsOut.Range(wsOut.Cells(topRow + 1, 2), wsOut.Cells(topRow + 1, 7)).Merge
    wsOut.Cells(topRow + 1, 2).Value = "Collars"
    wsOut.Cells(topRow + 1, 2).HorizontalAlignment = xlCenter
    wsOut.Cells(topRow + 1, 2).Font.Bold = True

    wsOut.Cells(topRow + 1, 8).Value = "ATM"
    wsOut.Cells(topRow + 1, 8).HorizontalAlignment = xlCenter
    wsOut.Cells(topRow + 1, 8).Font.Bold = True

    wsOut.Range(wsOut.Cells(topRow + 1, 9), wsOut.Cells(topRow + 1, 14)).Merge
    wsOut.Cells(topRow + 1, 9).Value = "Strangles"
    wsOut.Cells(topRow + 1, 9).HorizontalAlignment = xlCenter
    wsOut.Cells(topRow + 1, 9).Font.Bold = True

    wsOut.Cells(topRow + 2, 1).Value = ""
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

    For i = 1 To nPairs
        exp = Split(pairs(i), "|")(0)
        ten = Split(pairs(i), "|")(1)

        wsOut.Cells(topRow + 2 + i, 1).Value = exp & ten

        For j = 1 To 6
            callV = GetValueByMoneyness(wsSrc, cExp, cTen, cMon, cVal, lastRow, exp, ten, mons(j))
            putV = GetValueByMoneyness(wsSrc, cExp, cTen, cMon, cVal, lastRow, exp, ten, -mons(j))

            If IsMissingBoth(callV, putV) Then
                wsOut.Cells(topRow + 2 + i, 1 + j).Value = ""
                wsOut.Cells(topRow + 2 + i, 8 + j).Value = ""
            Else
                wsOut.Cells(topRow + 2 + i, 1 + j).Value = callV - putV  ' Collar
                wsOut.Cells(topRow + 2 + i, 8 + j).Value = callV + putV  ' Strangle
            End If
        Next j

        wsOut.Cells(topRow + 2 + i, 8).Value = "" ' ATM blank
    Next i
End Sub

' ==========================================================
' Helpers
' ==========================================================

Private Function FindHeaderCol(ws As Worksheet, headerName As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value2)) = headerName Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Private Function UniqueListFiltered(ws As Worksheet, _
    filterCol As Long, filterVal As String, _
    targetCol As Long, lastRow As Long, _
    ByRef nOut As Long) As String()

    Dim tmp() As String
    ReDim tmp(1 To lastRow)
    nOut = 0

    Dim r As Long, v As String
    For r = 2 To lastRow
        If CStr(ws.Cells(r, filterCol).Value) = filterVal Then
            v = Trim$(CStr(ws.Cells(r, targetCol).Value))
            If Len(v) > 0 Then
                If Not InArray(v, tmp, nOut) Then
                    nOut = nOut + 1
                    tmp(nOut) = v
                End If
            End If
        End If
    Next r

    ReDim Preserve tmp(1 To nOut)
    UniqueListFiltered = tmp
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

Private Function SumIfsSafe(ws As Worksheet, cVal As Long, _
                            c1 As Long, v1 As String, _
                            c2 As Long, v2 As String, _
                            c3 As Long, v3 As String, _
                            lastRow As Long) As Double
    On Error GoTo fail
    SumIfsSafe = Application.WorksheetFunction.SumIfs( _
        ws.Range(ws.Cells(2, cVal), ws.Cells(lastRow, cVal)), _
        ws.Range(ws.Cells(2, c1), ws.Cells(lastRow, c1)), v1, _
        ws.Range(ws.Cells(2, c2), ws.Cells(lastRow, c2)), v2, _
        ws.Range(ws.Cells(2, c3), ws.Cells(lastRow, c3)), v3)
    Exit Function
fail:
    SumIfsSafe = 0#
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

Private Function UniquePairsFromMoneyness(ws As Worksheet, cExp As Long, cTen As Long, cMon As Long, _
                                         lastRow As Long, mons() As Long, ByRef nOut As Long) As String()
    Dim tmp() As String
    ReDim tmp(1 To lastRow)
    nOut = 0

    Dim r As Long, exp As String, ten As String, key As String
    Dim M As Long, absM As Long

    For r = 2 To lastRow
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
    Next r

    If nOut = 0 Then
        ReDim UniquePairsFromMoneyness(1 To 1)
        Exit Function
    End If

    ReDim Preserve tmp(1 To nOut)
    UniquePairsFromMoneyness = tmp
End Function

Private Function GetValueByMoneyness(ws As Worksheet, cExp As Long, cTen As Long, cMon As Long, cVal As Long, _
                                     lastRow As Long, exp As String, ten As String, mon As Long) As Double
    On Error GoTo fail
    GetValueByMoneyness = Application.WorksheetFunction.SumIfs( _
        ws.Range(ws.Cells(2, cVal), ws.Cells(lastRow, cVal)), _
        ws.Range(ws.Cells(2, cExp), ws.Cells(lastRow, cExp)), exp, _
        ws.Range(ws.Cells(2, cTen), ws.Cells(lastRow, cTen)), ten, _
        ws.Range(ws.Cells(2, cMon), ws.Cells(lastRow, cMon)), mon)
    Exit Function
fail:
    GetValueByMoneyness = 0#
End Function

Private Function IsMissingBoth(a As Double, b As Double) As Boolean
    IsMissingBoth = (Abs(a) < 0.000000000001 And Abs(b) < 0.000000000001)
End Function


