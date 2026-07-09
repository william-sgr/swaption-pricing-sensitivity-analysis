Attribute VB_Name = "Sensitivities"
Option Explicit

' ==========================================================
' MODULE: SENSITIVITIES (Exercise 2b) + Parallel Delta
' - Filters ONLY Model="BLACK_SHIFT" AND Shift=0.05 from "VOL_LONG_ALL"
' - Computes Delta and Vega:
'     (i) Analytical (shifted-lognormal / Black-76 on shifted forwards)
'     (ii) Finite Differences (central) with multiple bumps (3 + 3)
' - Adds FD greeks ALSO "per annuity" (DeltaAnn_FD, VegaAnn_FD)
' - Adds Parallel Delta (OIS curve +/-1bp) via annuity bump:
'     Price = Annuity * PremAnn
'     ParDeltaPrice_per1bp = (Price(+1bp) - Price(-1bp))/2
' ==========================================================

Private Const valDate As Date = #10/31/2019#
Private Const TARGET_SHIFT As Double = 0.05
Private Const SHIFT_TOL As Double = 0.000000001

Private Const PAR_BUMP As Double = 0.0001 ' 1bp

' ============================
' Global curve + holidays
' ============================
Private gT() As Double   ' years ACT/365
Private gR() As Double   ' cont zero rates
Private gN As Long

Private gHol() As Long   ' Excel serial dates
Private gHolN As Long

' ==========================================================
' ENTRY
' ==========================================================
Public Sub BUILD_SENS_SHIFTED_BLACK_005()

    ' --- FD bumps (3 + 3) ---
    ' Forward bumps in bps (converted to decimal rate shift: bps/10000)
    Dim bumpsF_bps As Variant
    bumpsF_bps = Array(0.5, 1, 2)

    ' Vol bumps in abs vol (0.01 = +1% abs vol)
    Dim bumpsV As Variant
    bumpsV = Array(0.0025, 0.005, 0.01)

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsIn As Worksheet: Set wsIn = wb.Worksheets("VOL_LONG_ALL")

    ' load curve + holidays (needed for parallel delta)
    Load_OIS_Curve wb.Worksheets("IR Yield Curves")
    Load_Holidays wb.Worksheets("Calendar")

    Build_SENS_LONG_005 wsIn, TARGET_SHIFT, bumpsF_bps, bumpsV, PAR_BUMP
    Build_SENS_FD_CHECK_005 wsIn, TARGET_SHIFT, bumpsF_bps, bumpsV, PAR_BUMP

    MsgBox "Done: SENS_LONG_005 and SENS_FD_CHECK_005 (with DeltaAnn_FD + Parallel Delta).", vbInformation
End Sub

' ==========================================================
' 1) SENS_LONG_005
' One row per leg, wide layout:
' - analytic greeks
' - parallel delta (annuity + price) per 1bp
' - FD greeks for each bump:
'     for each df: DeltaAnn_FD + DeltaPrice_FD + errors (Ann + Price)
'     for each dv: VegaAnn_FD  + VegaPrice_FD  + errors (Ann + Price)
' ==========================================================
Private Sub Build_SENS_LONG_005(ByVal wsIn As Worksheet, ByVal targetShift As Double, _
                                ByVal bumpsF_bps As Variant, ByVal bumpsV As Variant, _
                                ByVal parBump As Double)

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsOut As Worksheet: Set wsOut = GetOrCreateSheet(wb, "SENS_LONG_005")
    wsOut.Cells.Clear

    Dim lastRow As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim cStatus As Long: cStatus = GetCol(wsIn, "Status")
    Dim baseCols As Long: baseCols = cStatus

    ' required columns
    Dim cTe As Long, cTenY As Long, cF As Long, cK As Long, cOpt As Long, cAnn As Long, cModel As Long, cShift As Long, cVol As Long
    cTe = GetCol(wsIn, "Te")
    cTenY = GetCol(wsIn, "SwapTenorY")
    cF = GetCol(wsIn, "FwdRate")
    cK = GetCol(wsIn, "Strike")
    cOpt = GetCol(wsIn, "OptType")
    cAnn = GetCol(wsIn, "Annuity_Te")
    cModel = GetCol(wsIn, "Model")
    cShift = GetCol(wsIn, "Shift")
    cVol = GetCol(wsIn, "ImplVol")

    Dim rngIn As Range
    Set rngIn = wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(lastRow, baseCols))
    Dim arrIn As Variant: arrIn = rngIn.Value2

    ' count kept rows
    Dim i As Long, nKeep As Long
    nKeep = 0
    For i = 2 To UBound(arrIn, 1)
        If UCase$(CStr(arrIn(i, cModel))) = "BLACK_SHIFT" Then
            If Abs(CDbl0(arrIn(i, cShift)) - targetShift) < SHIFT_TOL Then
                nKeep = nKeep + 1
            End If
        End If
    Next i
    If nKeep = 0 Then Exit Sub

    Dim nbF As Long: nbF = UBound(bumpsF_bps) - LBound(bumpsF_bps) + 1
    Dim nbV As Long: nbV = UBound(bumpsV) - LBound(bumpsV) + 1

    ' fixed extras:
    ' d1 d2 (2)
    ' analytic: PremAnn, DeltaAnn, VegaAnn, Price, DeltaPrice, VegaPrice, scaled (8)
    ' parallel: ParDeltaAnn_per1bp, ParDeltaPrice_per1bp (2)
    ' FD per df: (DeltaAnn_FD, abs, rel) + (DeltaPrice_FD, abs, rel) => 6 per bump
    ' FD per dv: (VegaAnn_FD, abs, rel) + (VegaPrice_FD, abs, rel)  => 6 per bump
    Dim EXTRA As Long
    EXTRA = 2 + 8 + 2 + (nbF * 6) + (nbV * 6)

    Dim outCols As Long: outCols = baseCols + EXTRA
    Dim arrOut As Variant
    ReDim arrOut(1 To nKeep + 1, 1 To outCols)

    ' copy base header
    Dim c As Long
    For c = 1 To baseCols
        arrOut(1, c) = arrIn(1, c)
    Next c

    Dim colX As Long
    colX = baseCols

    ' fixed extra headers
    colX = colX + 1: arrOut(1, colX) = "d1"
    colX = colX + 1: arrOut(1, colX) = "d2"

    colX = colX + 1: arrOut(1, colX) = "PremAnn_ANA"
    colX = colX + 1: arrOut(1, colX) = "DeltaAnn_ANA"
    colX = colX + 1: arrOut(1, colX) = "VegaAnn_ANA"

    colX = colX + 1: arrOut(1, colX) = "Price_ANA"
    colX = colX + 1: arrOut(1, colX) = "DeltaPrice_ANA"
    colX = colX + 1: arrOut(1, colX) = "VegaPrice_ANA"

    colX = colX + 1: arrOut(1, colX) = "DeltaPrice_per1bp"
    colX = colX + 1: arrOut(1, colX) = "VegaPrice_per1pct"

    colX = colX + 1: arrOut(1, colX) = "ParDeltaAnn_per1bp"
    colX = colX + 1: arrOut(1, colX) = "ParDeltaPrice_per1bp"

    ' FD headers: Delta (3 bumps) -> Ann + Price
    Dim j As Long, bump As Double
    For j = LBound(bumpsF_bps) To UBound(bumpsF_bps)
        bump = CDbl(bumpsF_bps(j))

        colX = colX + 1: arrOut(1, colX) = "DeltaAnn_FD_df=" & CStr(bump) & "bps"
        colX = colX + 1: arrOut(1, colX) = "DeltaAnn_FD_AbsErr_df=" & CStr(bump) & "bps"
        colX = colX + 1: arrOut(1, colX) = "DeltaAnn_FD_RelErr_df=" & CStr(bump) & "bps"

        colX = colX + 1: arrOut(1, colX) = "DeltaPrice_FD_df=" & CStr(bump) & "bps"
        colX = colX + 1: arrOut(1, colX) = "DeltaPrice_FD_AbsErr_df=" & CStr(bump) & "bps"
        colX = colX + 1: arrOut(1, colX) = "DeltaPrice_FD_RelErr_df=" & CStr(bump) & "bps"
    Next j

    ' FD headers: Vega (3 bumps) -> Ann + Price
    Dim dv As Double
    For j = LBound(bumpsV) To UBound(bumpsV)
        dv = CDbl(bumpsV(j))

        colX = colX + 1: arrOut(1, colX) = "VegaAnn_FD_dv=" & CStr(dv)
        colX = colX + 1: arrOut(1, colX) = "VegaAnn_FD_AbsErr_dv=" & CStr(dv)
        colX = colX + 1: arrOut(1, colX) = "VegaAnn_FD_RelErr_dv=" & CStr(dv)

        colX = colX + 1: arrOut(1, colX) = "VegaPrice_FD_dv=" & CStr(dv)
        colX = colX + 1: arrOut(1, colX) = "VegaPrice_FD_AbsErr_dv=" & CStr(dv)
        colX = colX + 1: arrOut(1, colX) = "VegaPrice_FD_RelErr_dv=" & CStr(dv)
    Next j

    ' fill
    Dim outR As Long: outR = 1

    Dim t As Double, TenY As Long, F As Double, k As Double, sh As Double, sig As Double, a As Double, omega As Double
    Dim d1 As Double, d2 As Double
    Dim premAnn As Double, dAnn As Double, vAnn As Double
    Dim priceA As Double, deltaPxA As Double, vegaPxA As Double

    Dim fdDeltaAnn As Double, fdDeltaPx As Double
    Dim fdVegaAnn As Double, fdVegaPx As Double
    Dim df As Double
    Dim absErr As Double
    Dim relErr As Variant

    Dim Aup As Double, Adn As Double
    Dim parDelA_per1bp As Double, parDelPx_per1bp As Double

    For i = 2 To UBound(arrIn, 1)

        If UCase$(CStr(arrIn(i, cModel))) <> "BLACK_SHIFT" Then GoTo NextI
        If Abs(CDbl0(arrIn(i, cShift)) - targetShift) >= SHIFT_TOL Then GoTo NextI

        outR = outR + 1

        ' copy base
        For c = 1 To baseCols
            arrOut(outR, c) = arrIn(i, c)
        Next c

        t = CDbl0(arrIn(i, cTe))
        TenY = CLng(CDbl0(arrIn(i, cTenY)))
        F = CDbl0(arrIn(i, cF))
        k = CDbl0(arrIn(i, cK))
        sh = targetShift
        sig = CDbl0(arrIn(i, cVol))
        a = CDbl0(arrIn(i, cAnn))
        omega = OmegaFromOptType(CStr(arrIn(i, cOpt)))

        colX = baseCols

        If omega = 0# Then GoTo NextI
        If Not BlackShift_d1d2(F, k, sig, t, sh, d1, d2) Then GoTo NextI

        premAnn = BlackShift_PremAnn(F, k, sig, t, sh, omega)
        dAnn = omega * StdNormCDF(omega * d1)
        vAnn = (F + sh) * StdNormPDF(d1) * Sqr(t)

        priceA = a * premAnn
        deltaPxA = a * dAnn
        vegaPxA = a * vAnn

        ' parallel delta via annuity bump (+/-1bp)
        Aup = OIS_Annuity_Bumped(t, TenY, parBump)
        Adn = OIS_Annuity_Bumped(t, TenY, -parBump)
        parDelA_per1bp = (Aup - Adn) / 2#
        parDelPx_per1bp = parDelA_per1bp * premAnn

        ' write fixed extras
        colX = colX + 1: arrOut(outR, colX) = d1
        colX = colX + 1: arrOut(outR, colX) = d2

        colX = colX + 1: arrOut(outR, colX) = premAnn
        colX = colX + 1: arrOut(outR, colX) = dAnn
        colX = colX + 1: arrOut(outR, colX) = vAnn

        colX = colX + 1: arrOut(outR, colX) = priceA
        colX = colX + 1: arrOut(outR, colX) = deltaPxA
        colX = colX + 1: arrOut(outR, colX) = vegaPxA

        colX = colX + 1: arrOut(outR, colX) = deltaPxA * 0.0001   ' per 1bp rate shift on F
        colX = colX + 1: arrOut(outR, colX) = vegaPxA * 0.01       ' per +1% abs vol shift

        colX = colX + 1: arrOut(outR, colX) = parDelA_per1bp
        colX = colX + 1: arrOut(outR, colX) = parDelPx_per1bp

        ' Delta FD for each df (3 bumps): store Ann + Price
        For j = LBound(bumpsF_bps) To UBound(bumpsF_bps)
            bump = CDbl(bumpsF_bps(j))
            df = bump / 10000#

            fdDeltaAnn = BlackShift_DeltaFD_Annuity(F, k, sig, t, sh, omega, df)
            fdDeltaPx = a * fdDeltaAnn

            ' Ann errors vs analytic DeltaAnn
            absErr = fdDeltaAnn - dAnn
            relErr = RelErrSafe(absErr, dAnn)
            colX = colX + 1: arrOut(outR, colX) = fdDeltaAnn
            colX = colX + 1: arrOut(outR, colX) = absErr
            colX = colX + 1: arrOut(outR, colX) = relErr

            ' Price errors vs analytic DeltaPrice
            absErr = fdDeltaPx - deltaPxA
            relErr = RelErrSafe(absErr, deltaPxA)
            colX = colX + 1: arrOut(outR, colX) = fdDeltaPx
            colX = colX + 1: arrOut(outR, colX) = absErr
            colX = colX + 1: arrOut(outR, colX) = relErr
        Next j

        ' Vega FD for each dv (3 bumps): store Ann + Price
        For j = LBound(bumpsV) To UBound(bumpsV)
            dv = CDbl(bumpsV(j))

            fdVegaAnn = BlackShift_VegaFD_Annuity(F, k, sig, t, sh, omega, dv)
            fdVegaPx = a * fdVegaAnn

            ' Ann errors vs analytic VegaAnn
            absErr = fdVegaAnn - vAnn
            relErr = RelErrSafe(absErr, vAnn)
            colX = colX + 1: arrOut(outR, colX) = fdVegaAnn
            colX = colX + 1: arrOut(outR, colX) = absErr
            colX = colX + 1: arrOut(outR, colX) = relErr

            ' Price errors vs analytic VegaPrice
            absErr = fdVegaPx - vegaPxA
            relErr = RelErrSafe(absErr, vegaPxA)
            colX = colX + 1: arrOut(outR, colX) = fdVegaPx
            colX = colX + 1: arrOut(outR, colX) = absErr
            colX = colX + 1: arrOut(outR, colX) = relErr
        Next j

NextI:
    Next i

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(outR, outCols)).Value2 = arrOut
    wsOut.Rows(1).Font.Bold = True
    wsOut.Columns.AutoFit
End Sub

' ==========================================================
' 2) SENS_FD_CHECK_005
' Long format:
' For each instrument -> 3 delta bumps + 3 vega bumps (vs analytic)
' Adds parallel delta columns (same for all bumps of the leg)
' ==========================================================
Private Sub Build_SENS_FD_CHECK_005(ByVal wsIn As Worksheet, ByVal targetShift As Double, _
                                    ByVal bumpsF_bps As Variant, ByVal bumpsV As Variant, _
                                    ByVal parBump As Double)

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsOut As Worksheet: Set wsOut = GetOrCreateSheet(wb, "SENS_FD_CHECK_005")
    wsOut.Cells.Clear

    Dim lastRow As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim cStatus As Long: cStatus = GetCol(wsIn, "Status")
    Dim baseCols As Long: baseCols = cStatus

    Dim cTe As Long, cTenY As Long, cF As Long, cK As Long, cOpt As Long, cAnn As Long, cModel As Long, cShift As Long, cVol As Long
    cTe = GetCol(wsIn, "Te")
    cTenY = GetCol(wsIn, "SwapTenorY")
    cF = GetCol(wsIn, "FwdRate")
    cK = GetCol(wsIn, "Strike")
    cOpt = GetCol(wsIn, "OptType")
    cAnn = GetCol(wsIn, "Annuity_Te")
    cModel = GetCol(wsIn, "Model")
    cShift = GetCol(wsIn, "Shift")
    cVol = GetCol(wsIn, "ImplVol")

    Dim rngIn As Range
    Set rngIn = wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(lastRow, baseCols))
    Dim arrIn As Variant: arrIn = rngIn.Value2

    Dim nbF As Long: nbF = UBound(bumpsF_bps) - LBound(bumpsF_bps) + 1
    Dim nbV As Long: nbV = UBound(bumpsV) - LBound(bumpsV) + 1

    ' count instruments kept
    Dim i As Long, nKeep As Long
    nKeep = 0
    For i = 2 To UBound(arrIn, 1)
        If UCase$(CStr(arrIn(i, cModel))) = "BLACK_SHIFT" Then
            If Abs(CDbl0(arrIn(i, cShift)) - targetShift) < SHIFT_TOL Then
                nKeep = nKeep + 1
            End If
        End If
    Next i
    If nKeep = 0 Then Exit Sub

    Dim nOut As Long: nOut = nKeep * (nbF + nbV)
    Dim outCols As Long: outCols = baseCols + 12

    Dim arrOut As Variant
    ReDim arrOut(1 To nOut + 1, 1 To outCols)

    Dim c As Long
    For c = 1 To baseCols: arrOut(1, c) = arrIn(1, c): Next c
    arrOut(1, baseCols + 1) = "Greek"
    arrOut(1, baseCols + 2) = "BumpType"
    arrOut(1, baseCols + 3) = "Bump"
    arrOut(1, baseCols + 4) = "Value_Annuity_FD"
    arrOut(1, baseCols + 5) = "Value_Price_FD"
    arrOut(1, baseCols + 6) = "Value_Annuity_ANA"
    arrOut(1, baseCols + 7) = "Value_Price_ANA"
    arrOut(1, baseCols + 8) = "AbsErr_Price"
    arrOut(1, baseCols + 9) = "RelErr_Price"
    arrOut(1, baseCols + 10) = "ParDeltaAnn_per1bp"
    arrOut(1, baseCols + 11) = "ParDeltaPrice_per1bp"
    arrOut(1, baseCols + 12) = "Note"

    Dim outR As Long: outR = 1

    Dim t As Double, TenY As Long, F As Double, k As Double, sh As Double, sig As Double, a As Double, omega As Double
    Dim d1 As Double, d2 As Double
    Dim premAnn As Double
    Dim dAnn As Double, vAnn As Double, dPx As Double, vPx As Double

    Dim bump As Double, df As Double, dv As Double
    Dim fdValAnn As Double, fdValPx As Double
    Dim anaValAnn As Double, anaValPx As Double
    Dim absErr As Double, relErr As Variant, note As String
    Dim j As Long

    Dim Aup As Double, Adn As Double
    Dim parDelA_per1bp As Double, parDelPx_per1bp As Double

    For i = 2 To UBound(arrIn, 1)

        If UCase$(CStr(arrIn(i, cModel))) <> "BLACK_SHIFT" Then GoTo NextI
        If Abs(CDbl0(arrIn(i, cShift)) - targetShift) >= SHIFT_TOL Then GoTo NextI

        t = CDbl0(arrIn(i, cTe))
        TenY = CLng(CDbl0(arrIn(i, cTenY)))
        F = CDbl0(arrIn(i, cF))
        k = CDbl0(arrIn(i, cK))
        sh = targetShift
        sig = CDbl0(arrIn(i, cVol))
        a = CDbl0(arrIn(i, cAnn))
        omega = OmegaFromOptType(CStr(arrIn(i, cOpt)))
        If omega = 0# Then GoTo NextI

        If Not BlackShift_d1d2(F, k, sig, t, sh, d1, d2) Then GoTo NextI

        premAnn = BlackShift_PremAnn(F, k, sig, t, sh, omega)
        dAnn = omega * StdNormCDF(omega * d1)
        vAnn = (F + sh) * StdNormPDF(d1) * Sqr(t)

        dPx = a * dAnn
        vPx = a * vAnn

        ' parallel delta via annuity bump (+/-1bp)
        Aup = OIS_Annuity_Bumped(t, TenY, parBump)
        Adn = OIS_Annuity_Bumped(t, TenY, -parBump)
        parDelA_per1bp = (Aup - Adn) / 2#
        parDelPx_per1bp = parDelA_per1bp * premAnn

        ' ---- DELTA vs 3 df ----
        For j = LBound(bumpsF_bps) To UBound(bumpsF_bps)
            bump = CDbl(bumpsF_bps(j))
            df = bump / 10000#

            outR = outR + 1
            For c = 1 To baseCols: arrOut(outR, c) = arrIn(i, c): Next c

            anaValAnn = dAnn
            anaValPx = dPx

            fdValAnn = BlackShift_DeltaFD_Annuity(F, k, sig, t, sh, omega, df)
            fdValPx = a * fdValAnn

            absErr = fdValPx - anaValPx
            relErr = RelErrSafe(absErr, anaValPx)

            arrOut(outR, baseCols + 1) = "DELTA"
            arrOut(outR, baseCols + 2) = "FWD_BPS"
            arrOut(outR, baseCols + 3) = bump
            arrOut(outR, baseCols + 4) = fdValAnn
            arrOut(outR, baseCols + 5) = fdValPx
            arrOut(outR, baseCols + 6) = anaValAnn
            arrOut(outR, baseCols + 7) = anaValPx
            arrOut(outR, baseCols + 8) = absErr
            arrOut(outR, baseCols + 9) = relErr
            arrOut(outR, baseCols + 10) = parDelA_per1bp
            arrOut(outR, baseCols + 11) = parDelPx_per1bp
            arrOut(outR, baseCols + 12) = vbNullString
        Next j

        ' ---- VEGA vs 3 dv ----
        For j = LBound(bumpsV) To UBound(bumpsV)
            dv = CDbl(bumpsV(j))

            outR = outR + 1
            For c = 1 To baseCols: arrOut(outR, c) = arrIn(i, c): Next c

            anaValAnn = vAnn
            anaValPx = vPx

            note = vbNullString
            If dv <= 0# Then note = "dv<=0"
            If sig <= dv Then note = "sig<=dv (unstable/skip)"

            fdValAnn = BlackShift_VegaFD_Annuity(F, k, sig, t, sh, omega, dv)
            fdValPx = a * fdValAnn

            absErr = fdValPx - anaValPx
            relErr = RelErrSafe(absErr, anaValPx)

            arrOut(outR, baseCols + 1) = "VEGA"
            arrOut(outR, baseCols + 2) = "VOL"
            arrOut(outR, baseCols + 3) = dv
            arrOut(outR, baseCols + 4) = fdValAnn
            arrOut(outR, baseCols + 5) = fdValPx
            arrOut(outR, baseCols + 6) = anaValAnn
            arrOut(outR, baseCols + 7) = anaValPx
            arrOut(outR, baseCols + 8) = absErr
            arrOut(outR, baseCols + 9) = relErr
            arrOut(outR, baseCols + 10) = parDelA_per1bp
            arrOut(outR, baseCols + 11) = parDelPx_per1bp
            arrOut(outR, baseCols + 12) = note
        Next j

NextI:
    Next i

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(outR, outCols)).Value2 = arrOut
    wsOut.Rows(1).Font.Bold = True
    wsOut.Columns.AutoFit
End Sub

' ==========================================================
' Shifted-Black core
' ==========================================================
Private Function BlackShift_d1d2(ByVal F As Double, ByVal k As Double, ByVal sig As Double, ByVal t As Double, ByVal shift As Double, _
                                 ByRef d1 As Double, ByRef d2 As Double) As Boolean
    BlackShift_d1d2 = False
    If t <= 0# Or sig <= 0# Then Exit Function
    Dim Fp As Double: Fp = F + shift
    Dim Kp As Double: Kp = k + shift
    If Fp <= 0# Or Kp <= 0# Then Exit Function

    Dim sqrtT As Double: sqrtT = Sqr(t)
    d1 = (Log(Fp / Kp) + 0.5 * sig * sig * t) / (sig * sqrtT)
    d2 = d1 - sig * sqrtT
    BlackShift_d1d2 = True
End Function

Private Function BlackShift_PremAnn(ByVal F As Double, ByVal k As Double, ByVal sig As Double, ByVal t As Double, ByVal shift As Double, ByVal omega As Double) As Double
    Dim d1 As Double, d2 As Double
    If Not BlackShift_d1d2(F, k, sig, t, shift, d1, d2) Then
        BlackShift_PremAnn = 0#
        Exit Function
    End If
    Dim Fp As Double: Fp = F + shift
    Dim Kp As Double: Kp = k + shift
    BlackShift_PremAnn = omega * (Fp * StdNormCDF(omega * d1) - Kp * StdNormCDF(omega * d2))
End Function

Private Function BlackShift_DeltaFD_Annuity(ByVal F As Double, ByVal k As Double, ByVal sig As Double, ByVal t As Double, ByVal shift As Double, ByVal omega As Double, ByVal df As Double) As Double
    If df <= 0# Then
        BlackShift_DeltaFD_Annuity = 0#
        Exit Function
    End If
    Dim up As Double, dn As Double
    up = BlackShift_PremAnn(F + df, k, sig, t, shift, omega)
    dn = BlackShift_PremAnn(F - df, k, sig, t, shift, omega)
    BlackShift_DeltaFD_Annuity = (up - dn) / (2# * df)
End Function

Private Function BlackShift_VegaFD_Annuity(ByVal F As Double, ByVal k As Double, ByVal sig As Double, ByVal t As Double, ByVal shift As Double, ByVal omega As Double, ByVal dv As Double) As Double
    If dv <= 0# Or sig <= dv Then
        BlackShift_VegaFD_Annuity = 0#
        Exit Function
    End If
    Dim up As Double, dn As Double
    up = BlackShift_PremAnn(F, k, sig + dv, t, shift, omega)
    dn = BlackShift_PremAnn(F, k, sig - dv, t, shift, omega)
    BlackShift_VegaFD_Annuity = (up - dn) / (2# * dv)
End Function

' ==========================================================
' Parallel Delta helpers: OIS curve + annuity builder
' ==========================================================

Private Sub Load_OIS_Curve(ByVal ws As Worksheet)
    ' Expected format (as in your workbook):
    '   col E: tenor in days (1,4,5,11,...)  ; col F: cont zero rate
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    If lastR < 2 Then Err.Raise vbObjectError + 700, "Load_OIS_Curve", "Empty OIS curve range."

    ' count numeric points
    Dim n As Long: n = 0
    For r = 2 To lastR
        If IsNumeric(ws.Cells(r, 5).Value2) And IsNumeric(ws.Cells(r, 6).Value2) Then
            If ws.Cells(r, 5).Value2 > 0 Then n = n + 1
        End If
    Next r
    If n = 0 Then Err.Raise vbObjectError + 701, "Load_OIS_Curve", "No OIS points found."

    ReDim gT(1 To n)
    ReDim gR(1 To n)
    gN = n

    Dim i As Long: i = 0
    For r = 2 To lastR
        If IsNumeric(ws.Cells(r, 5).Value2) And IsNumeric(ws.Cells(r, 6).Value2) Then
            If ws.Cells(r, 5).Value2 > 0 Then
                i = i + 1
                gT(i) = CDbl(ws.Cells(r, 5).Value2) / 365#
                gR(i) = CDbl(ws.Cells(r, 6).Value2)
            End If
        End If
    Next r
End Sub

Private Sub Load_Holidays(ByVal ws As Worksheet)
    ' Expected format (as in your workbook):
    '   col E: holiday dates (excluding weekends), from row 2 down
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    If lastR < 2 Then
        gHolN = 0
        Exit Sub
    End If

    Dim n As Long: n = 0
    For r = 2 To lastR
        If IsDate(ws.Cells(r, 5).Value2) Then n = n + 1
    Next r

    If n = 0 Then
        gHolN = 0
        Exit Sub
    End If

    ReDim gHol(1 To n)
    gHolN = n

    Dim i As Long: i = 0
    For r = 2 To lastR
        If IsDate(ws.Cells(r, 5).Value2) Then
            i = i + 1
            gHol(i) = CLng(CDate(ws.Cells(r, 5).Value2))
        End If
    Next r
End Sub

Private Function OIS_ZeroRate(ByVal tYears As Double) As Double
    If gN <= 0 Then Err.Raise vbObjectError + 710, "OIS_ZeroRate", "Curve not loaded."
    If tYears <= gT(1) Then OIS_ZeroRate = gR(1): Exit Function
    If tYears >= gT(gN) Then OIS_ZeroRate = gR(gN): Exit Function

    Dim i As Long
    For i = 1 To gN - 1
        If tYears >= gT(i) And tYears <= gT(i + 1) Then
            Dim w As Double
            w = (tYears - gT(i)) / (gT(i + 1) - gT(i))
            OIS_ZeroRate = gR(i) + w * (gR(i + 1) - gR(i))
            Exit Function
        End If
    Next i
End Function

Private Function OIS_DF_Bumped(ByVal tYears As Double, ByVal bumpRate As Double) As Double
    Dim r As Double
    r = OIS_ZeroRate(tYears) + bumpRate
    OIS_DF_Bumped = exp(-r * tYears)
End Function

Private Function OIS_Annuity_Bumped(ByVal Te As Double, ByVal swapTenorY As Long, ByVal bumpRate As Double) As Double
    ' Build annual fixed schedule from start = expiry + 2bd, Following
    ' Accrual: 30E/360
    ' DF from bumped OIS curve

    If swapTenorY <= 0 Then
        OIS_Annuity_Bumped = 0#
        Exit Function
    End If

    Dim expiryDate As Date
    expiryDate = DateAdd("d", CLng(Application.WorksheetFunction.Round(Te * 365#, 0)), valDate)

    Dim startDate As Date
    startDate = BusinessDayAdd(expiryDate, 2)

    Dim ann As Double: ann = 0#
    Dim prevDate As Date: prevDate = startDate

    Dim y As Long
    For y = 1 To swapTenorY
        Dim payDate As Date
        payDate = DateAdd("yyyy", y, startDate)
        payDate = FollowingBusinessDay(payDate)

        Dim accr As Double
        accr = YearFrac_30E360(prevDate, payDate)

        Dim tPay As Double
        tPay = (CLng(payDate) - CLng(valDate)) / 365#

        ann = ann + accr * OIS_DF_Bumped(tPay, bumpRate)

        prevDate = payDate
    Next y

    OIS_Annuity_Bumped = ann
End Function

Private Function IsHoliday(ByVal d As Date) As Boolean
    If gHolN <= 0 Then
        IsHoliday = False
        Exit Function
    End If
    Dim x As Long: x = CLng(d)
    Dim i As Long
    For i = 1 To gHolN
        If gHol(i) = x Then
            IsHoliday = True
            Exit Function
        End If
    Next i
    IsHoliday = False
End Function

Private Function IsBusinessDay(ByVal d As Date) As Boolean
    Dim wd As VbDayOfWeek
    wd = Weekday(d, vbMonday)
    If wd >= 6 Then
        IsBusinessDay = False
        Exit Function
    End If
    If IsHoliday(d) Then
        IsBusinessDay = False
        Exit Function
    End If
    IsBusinessDay = True
End Function

Private Function FollowingBusinessDay(ByVal d As Date) As Date
    Dim x As Date: x = d
    Do While Not IsBusinessDay(x)
        x = DateAdd("d", 1, x)
    Loop
    FollowingBusinessDay = x
End Function

Private Function BusinessDayAdd(ByVal d As Date, ByVal nBusDays As Long) As Date
    Dim x As Date: x = d
    Dim n As Long: n = 0
    Do While n < nBusDays
        x = DateAdd("d", 1, x)
        If IsBusinessDay(x) Then n = n + 1
    Loop
    BusinessDayAdd = x
End Function

Private Function YearFrac_30E360(ByVal dt1 As Date, ByVal dt2 As Date) As Double
    Dim y1 As Long, m1 As Long, d1 As Long
    Dim y2 As Long, m2 As Long, d2 As Long

    y1 = Year(dt1): m1 = Month(dt1): d1 = Day(dt1)
    y2 = Year(dt2): m2 = Month(dt2): d2 = Day(dt2)

    If d1 = 31 Then d1 = 30
    If d2 = 31 Then d2 = 30

    YearFrac_30E360 = (360# * (y2 - y1) + 30# * (m2 - m1) + (d2 - d1)) / 360#
End Function


' ==========================================================
' Utilities (minimal set)
' ==========================================================
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(shName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.name = shName
    End If
End Function

Private Function GetCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value2)) = headerName Then
            GetCol = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 513, "GetCol", "Header not found: " & headerName
End Function

Private Function OmegaFromOptType(ByVal s As String) As Double
    Dim u As String: u = UCase$(s)
    If InStr(u, "CALL") > 0 Or InStr(u, "PAYER") > 0 Then
        OmegaFromOptType = 1#
    ElseIf InStr(u, "PUT") > 0 Or InStr(u, "RECEIVER") > 0 Then
        OmegaFromOptType = -1#
    Else
        OmegaFromOptType = 0#
    End If
End Function

Private Function CDbl0(ByVal v As Variant) As Double
    On Error GoTo fail
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Or IsObject(v) Then GoTo fail
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then GoTo fail
    End If
    CDbl0 = CDbl(v)
    Exit Function
fail:
    CDbl0 = 0#
End Function

Private Function StdNormCDF(ByVal x As Double) As Double
    On Error GoTo useOld
    StdNormCDF = Application.WorksheetFunction.Norm_S_Dist(x, True)
    Exit Function
useOld:
    On Error GoTo 0
    StdNormCDF = Application.WorksheetFunction.NormSDist(x)
End Function

Private Function StdNormPDF(ByVal x As Double) As Double
    On Error GoTo useOld
    StdNormPDF = Application.WorksheetFunction.Norm_S_Dist(x, False)
    Exit Function
useOld:
    On Error GoTo 0
    StdNormPDF = Application.WorksheetFunction.NormDist(x, 0#, 1#, False)
End Function

Private Function RelErrSafe(ByVal absErr As Double, ByVal refVal As Double) As Variant
    If Abs(refVal) < 0.000000000001 Then
        RelErrSafe = vbNullString
    Else
        RelErrSafe = absErr / refVal
    End If
End Function


