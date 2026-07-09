Attribute VB_Name = "Delta_strat"
Option Explicit

' ==========================================================
' FD DELTA -> per 1bp, printed like analytic sheet
' Source: SENS_LONG_005
' Uses columns (FD derivative in Price):
'   DeltaPrice_FD_df=0,5bps
'   DeltaPrice_FD_df=1bps
'   DeltaPrice_FD_df=2bps
'
' Conversion to per 1bp:
'   DeltaPrice_FD_per1bp = DeltaPrice_FD * 1e-4
'
' ATM STRADDLE per 1bp:
'   (2*DeltaPrice_FD - Annuity) * 1e-4
'
' OTM collars/strangles per 1bp:
'   Collar   = (call - put) * 1e-4
'   Strangle = (call + put) * 1e-4
'
' Output: ShiftGreeks_FD_per1bp
' ==========================================================

Public Sub Build_Delta_FD_per1bp()

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastRow As Long

    Set wb = ThisWorkbook
    Set wsSrc = wb.Worksheets("SENS_LONG_005")

    ' --- Required columns
    Dim cBlock As Long, cExp As Long, cTen As Long, cMon As Long, cAnn As Long
    Dim cD05 As Long, cD1 As Long, cD2 As Long

    cBlock = FindHeaderCol(wsSrc, "SourceBlock")
    cExp = FindHeaderCol(wsSrc, "ExpiryLbl")
    cTen = FindHeaderCol(wsSrc, "TenorLbl")
    cMon = FindHeaderCol(wsSrc, "MoneynessBP")
    cAnn = FindHeaderCol(wsSrc, "Annuity_Te")

    cD05 = FindHeaderCol(wsSrc, "DeltaPrice_FD_df=0,5bps")
    cD1 = FindHeaderCol(wsSrc, "DeltaPrice_FD_df=1bps")
    cD2 = FindHeaderCol(wsSrc, "DeltaPrice_FD_df=2bps")

    If cBlock * cExp * cTen * cMon * cAnn * cD05 * cD1 * cD2 = 0 Then
        MsgBox "STOP: mancano colonne in SENS_LONG_005 (SourceBlock, ExpiryLbl, TenorLbl, MoneynessBP, Annuity_Te, DeltaPrice_FD_df=0,5bps/1bps/2bps).", vbCritical
        Exit Sub
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, cExp).End(xlUp).Row

    ' --- ATM grids
    Dim expList() As String, tenList() As String
    Dim nExp As Long, nTen As Long
    expList = UniqueListFiltered(wsSrc, cBlock, "ATM", cExp, lastRow, nExp)
    tenList = UniqueListFiltered(wsSrc, cBlock, "ATM", cTen, lastRow, nTen)

    If nExp = 0 Or nTen = 0 Then
        MsgBox "Nessun Expiry/Tenor ATM trovato (SourceBlock='ATM').", vbCritical
        Exit Sub
    End If

    ' --- OTM pairs
    Dim mons(1 To 6) As Long
    mons(1) = 50: mons(2) = 100: mons(3) = 150: mons(4) = 200: mons(5) = 300: mons(6) = 400

    Dim pairs() As String, nPairs As Long
    pairs = UniquePairsFromMoneyness(wsSrc, cExp, cTen, cMon, lastRow, mons, nPairs)
    If nPairs = 0 Then
        MsgBox "Non trovo righe OTM con |MoneynessBP| in {50,100,150,200,300,400}.", vbCritical
        Exit Sub
    End If

    ' --- recreate output
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("ShiftGreeks_FD_per1bp").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.name = "DELTA_FD_per1bp"

    Dim top As Long: top = 1

    ' =========================
    ' ATM STRADDLE (3 bumps)
    ' =========================
    WriteBlock_STRADDLE_Delta_FD_per1bp wsSrc, wsOut, top, _
        "DeltaPrice_FD_df=0,5bps -> per1bp (ATM STRADDLE)", _
        cBlock, cExp, cTen, cD05, cAnn, expList, tenList, nExp, nTen, lastRow
    top = top + (nExp + 2) + 2

    WriteBlock_STRADDLE_Delta_FD_per1bp wsSrc, wsOut, top, _
        "DeltaPrice_FD_df=1bps -> per1bp (ATM STRADDLE)", _
        cBlock, cExp, cTen, cD1, cAnn, expList, tenList, nExp, nTen, lastRow
    top = top + (nExp + 2) + 2

    WriteBlock_STRADDLE_Delta_FD_per1bp wsSrc, wsOut, top, _
        "DeltaPrice_FD_df=2bps -> per1bp (ATM STRADDLE)", _
        cBlock, cExp, cTen, cD2, cAnn, expList, tenList, nExp, nTen, lastRow
    top = top + (nExp + 2) + 3

    ' =========================
    ' OTM (3 bumps)
    ' =========================
    WriteStrategyBlock_FD_per1bp wsSrc, wsOut, top, _
        "DeltaPrice_FD_df=0,5bps -> per1bp (OTM) - Collars/Strangles", _
        cExp, cTen, cMon, cD05, lastRow, pairs, nPairs, mons
    top = top + (nPairs + 2) + 3

    WriteStrategyBlock_FD_per1bp wsSrc, wsOut, top, _
        "DeltaPrice_FD_df=1bps -> per1bp (OTM) - Collars/Strangles", _
        cExp, cTen, cMon, cD1, lastRow, pairs, nPairs, mons
    top = top + (nPairs + 2) + 3

    WriteStrategyBlock_FD_per1bp wsSrc, wsOut, top, _
        "DeltaPrice_FD_df=2bps -> per1bp (OTM) - Collars/Strangles", _
        cExp, cTen, cMon, cD2, lastRow, pairs, nPairs, mons

    wsOut.Columns.AutoFit
    MsgBox "Creato foglio 'ShiftGreeks_FD_per1bp' (FD Delta -> per 1bp) con layout identico.", vbInformation

End Sub

' ==========================================================
' ATM STRADDLE from FD derivative -> per1bp
' Straddle_per1bp = (2*DeltaPrice_FD - Annuity) * 1e-4
' ==========================================================
Private Sub WriteBlock_STRADDLE_Delta_FD_per1bp(wsSrc As Worksheet, wsOut As Worksheet, topRow As Long, _
                                                title As String, _
                                                cBlock As Long, cExp As Long, cTen As Long, _
                                                cDeltaFD As Long, cAnn As Long, _
                                                expList() As String, tenList() As String, nExp As Long, nTen As Long, _
                                                lastRow As Long)

    Dim i As Long, j As Long
    Dim dCallFD As Double, a As Double

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
            ' FD delta is derivative in Price wrt F (CALL leg only at ATM)
            dCallFD = SumIfsSafe(wsSrc, cDeltaFD, cBlock, "ATM", cExp, expList(i), cTen, tenList(j), lastRow)
            a = SumIfsSafe(wsSrc, cAnn, cBlock, "ATM", cExp, expList(i), cTen, tenList(j), lastRow)

            ' Convert to per1bp AND build straddle via parity adjustment:
            wsOut.Cells(topRow + 1 + i, j + 1).Value = (2# * dCallFD - a) * 0.0001
        Next j
    Next i
End Sub

' ==========================================================
' OTM collars/strangles from FD derivative -> per1bp
' Collar_per1bp   = (call - put) * 1e-4
' Strangle_per1bp = (call + put) * 1e-4
' ==========================================================
Private Sub WriteStrategyBlock_FD_per1bp(wsSrc As Worksheet, wsOut As Worksheet, topRow As Long, _
                                         title As String, _
                                         cExp As Long, cTen As Long, cMon As Long, cValFD As Long, _
                                         lastRow As Long, pairs() As String, nPairs As Long, mons() As Long)

    Dim j As Long, i As Long
    Dim exp As String, ten As String
    Dim callV As Double, putV As Double

    wsOut.Cells(topRow, 1).Value = title
    wsOut.Cells(topRow, 1).Font.Bold = True

    ' header row 1
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

    ' header row 2
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

    ' body
    For i = 1 To nPairs
        exp = Split(pairs(i), "|")(0)
        ten = Split(pairs(i), "|")(1)

        wsOut.Cells(topRow + 2 + i, 1).Value = exp & ten

        For j = 1 To 6
            callV = GetValueByMoneyness(wsSrc, cExp, cTen, cMon, cValFD, lastRow, exp, ten, mons(j))
            putV = GetValueByMoneyness(wsSrc, cExp, cTen, cMon, cValFD, lastRow, exp, ten, -mons(j))

            If IsMissingBoth(callV, putV) Then
                wsOut.Cells(topRow + 2 + i, 1 + j).Value = ""
                wsOut.Cells(topRow + 2 + i, 8 + j).Value = ""
            Else
                wsOut.Cells(topRow + 2 + i, 1 + j).Value = (callV - putV) * 0.0001  ' Collar per1bp
                wsOut.Cells(topRow + 2 + i, 8 + j).Value = (callV + putV) * 0.0001  ' Strangle per1bp
            End If
        Next j

        wsOut.Cells(topRow + 2 + i, 8).Value = "" ' ATM blank (as analytic)
    Next i

End Sub

' ==========================================================
' Helpers (same as your analytic module)
' ==========================================================

Public Function FindHeaderCol(ws As Worksheet, headerName As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(1, c).Value)) = headerName Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Public Function UniqueListFiltered(ws As Worksheet, _
    filterCol As Long, filterVal As String, _
    targetCol As Long, lastRow As Long, _
    ByRef nOut As Long) As String()

    Dim tmp() As String
    ReDim tmp(1 To lastRow)
    nOut = 0

    Dim r As Long, v As String
    For r = 2 To lastRow
        If CStr(ws.Cells(r, filterCol).Value) = filterVal Then
            v = Trim(CStr(ws.Cells(r, targetCol).Value))
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

Public Function UniquePairsFromMoneyness(ws As Worksheet, cExp As Long, cTen As Long, cMon As Long, _
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
            exp = Trim(CStr(ws.Cells(r, cExp).Value))
            ten = Trim(CStr(ws.Cells(r, cTen).Value))
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

