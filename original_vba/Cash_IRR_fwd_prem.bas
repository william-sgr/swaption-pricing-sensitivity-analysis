Attribute VB_Name = "Cash_IRR_fwd_prem"
Option Explicit

' ==========================================================
' MODULE: CASH IRR FORWARD PREMIUM (from VOL_LONG_ALL ONLY)
' Output sheet: CASH_IRR_FWD_PREM_LONG
'
' Conventions aligned with VOL_LONG_ALL:
' - Valuation date fixed: 31/10/2019
' - Expiry date from label (1M, 3M, 1Y, ...)
' - Te = ACT/365(valuation, expiry)
' - Underlying swap start = expiry + 2 business days, then Following
' - Fixed leg: annual pay, Following adjustment, daycount 30E/360
'
' Cash IRR annuity Ã(R) built from the fixed-leg schedule (NOT constant tau):
'   Ã(R) = Sum_i tau_i / ?_{j=1..i} (1 + R * tau_j)
'
' Forward premium in bp:
'   CashIRRPremiumBP = 10000 * Ã(FwdRate) * PricePerCashIRRAnnuity(model, vol)
'
' IMPORTANT (as requested):
' - Input is ONLY VOL_LONG_ALL.
' - If ImplVol is missing/non-numeric/<=0 -> NO PREMIUM (leave hole), do NOT infer.
' ==========================================================

Private Const INPUT_SHEET As String = "VOL_LONG_ALL"
Private Const OUTPUT_SHEET As String = "CASH_IRR_FWD_PREM_LONG"
Private Const CAL_SHEET As String = "Calendar"
Private Const VAL_Y As Long = 2019, VAL_M As Long = 10, VAL_D As Long = 31

' Fixed valuation date
Private gValDate As Date

' Holidays storage (serial dates sorted)
Private gHol() As Long
Private gHolN As Long

Public Sub CIRR_BuildCashIRRForwardPremiumLong()

    Dim wb As Workbook
    Dim wsIn As Worksheet, wsOut As Worksheet, wsCal As Worksheet
    Dim lastRow As Long
    Dim inData As Variant, outData As Variant
    Dim r As Long

    Dim expLbl As String
    Dim TeIn As Double, swapTenY As Double, F As Double, mBP As Double, k As Double
    Dim optType As String, model As String
    Dim shift As Double, implVol As Variant, status As String

    Dim expDate As Date, Te As Double
    Dim isCall As Boolean
    Dim cashA As Double, cashAStatus As String
    Dim pricePerCashA As Variant, priceStatus As String
    Dim premCashBP As Variant
    Dim outStatus As String

    Set wb = ThisWorkbook

    If Not SheetExists(wb, INPUT_SHEET) Then
        MsgBox "Missing input sheet: " & INPUT_SHEET, vbExclamation
        Exit Sub
    End If
    If Not SheetExists(wb, CAL_SHEET) Then
        MsgBox "Missing sheet: " & CAL_SHEET, vbExclamation
        Exit Sub
    End If

    Set wsIn = wb.Worksheets(INPUT_SHEET)
    Set wsCal = wb.Worksheets(CAL_SHEET)
    Set wsOut = GetOrCreateAndReset(wb, OUTPUT_SHEET)

    gValDate = DateSerial(VAL_Y, VAL_M, VAL_D)
    CIRR_Load_Holidays wsCal

    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox INPUT_SHEET & " is empty.", vbExclamation
        Exit Sub
    End If

    ' Read input (A:Q = 17 cols)
    inData = wsIn.Range("A1").Resize(lastRow, 17).Value2
    ReDim outData(1 To lastRow, 1 To 21)

    ' Header (21 cols)
    outData(1, 1) = "SourceBlock"
    outData(1, 2) = "Instrument"
    outData(1, 3) = "ExpiryLbl"
    outData(1, 4) = "TenorLbl"
    outData(1, 5) = "Te"
    outData(1, 6) = "SwapTenorY"
    outData(1, 7) = "FwdRate"
    outData(1, 8) = "MoneynessBP"
    outData(1, 9) = "Strike"
    outData(1, 10) = "OptType"
    outData(1, 11) = "PremiumBP"
    outData(1, 12) = "Annuity_Te"
    outData(1, 13) = "PricePerAnnuity"
    outData(1, 14) = "Model"
    outData(1, 15) = "Shift"
    outData(1, 16) = "ImplVol"
    outData(1, 17) = "Status"
    outData(1, 18) = "CashIRR_Annuity"
    outData(1, 19) = "PricePerCashIRRAnnuity"
    outData(1, 20) = "CashIRRPremiumBP"
    outData(1, 21) = "CashIRRStatus"

    ' Loop rows
    For r = 2 To lastRow

        ' Copy original 17 cols (but Te will be overwritten with recomputed Te below)
        outData(r, 1) = inData(r, 1)
        outData(r, 2) = inData(r, 2)
        outData(r, 3) = inData(r, 3)
        outData(r, 4) = inData(r, 4)
        outData(r, 5) = inData(r, 5)
        outData(r, 6) = inData(r, 6)
        outData(r, 7) = inData(r, 7)
        outData(r, 8) = inData(r, 8)
        outData(r, 9) = inData(r, 9)
        outData(r, 10) = inData(r, 10)
        outData(r, 11) = inData(r, 11)
        outData(r, 12) = inData(r, 12)
        outData(r, 13) = inData(r, 13)
        outData(r, 14) = inData(r, 14)
        outData(r, 15) = inData(r, 15)
        outData(r, 16) = inData(r, 16)
        outData(r, 17) = inData(r, 17)

        expLbl = CStr(inData(r, 3))
        TeIn = CDbl(Nz(inData(r, 5), 0#))
        swapTenY = CDbl(Nz(inData(r, 6), 0#))
        F = CDbl(Nz(inData(r, 7), 0#))
        mBP = CDbl(Nz(inData(r, 8), 0#)) ' not used here, but keep
        k = CDbl(Nz(inData(r, 9), 0#))
        optType = CStr(inData(r, 10))
        model = CStr(inData(r, 14))
        shift = CDbl(Nz(inData(r, 15), 0#))
        implVol = inData(r, 16)
        status = CStr(inData(r, 17))

        isCall = CIRR_IsCall(optType)

        ' Expiry date + Te
        Te = TeIn
        On Error Resume Next
        expDate = CIRR_ExpiryDate(gValDate, expLbl)
        If Err.Number = 0 Then
            Te = CIRR_YF_ACT365(gValDate, expDate)
            If Te <= 0# Then Te = TeIn
        Else
            Err.Clear
        End If
        On Error GoTo 0

        outData(r, 5) = Te

        ' Cash IRR annuity
        cashAStatus = "OK"
        cashA = CIRR_CashIRRAnnuity_Dates(F, gValDate, expDate, swapTenY, cashAStatus)

        ' Price per Cash IRR annuity from vol (NO VOL => hole)
        priceStatus = "OK"
        pricePerCashA = CIRR_PricePerAnnuityFromVol(model, F, k, Te, implVol, isCall, shift, priceStatus)

        ' Premium in bp
        premCashBP = ""
        If cashAStatus = "OK" And (Not IsError(pricePerCashA)) Then
            premCashBP = 10000# * cashA * CDbl(pricePerCashA)
        End If

        ' Write extra cols
        If cashAStatus = "OK" Then
            outData(r, 18) = cashA
        Else
            outData(r, 18) = ""
        End If

        If IsError(pricePerCashA) Then
            outData(r, 19) = ""
        Else
            outData(r, 19) = CDbl(pricePerCashA)
        End If

        outData(r, 20) = premCashBP

        outStatus = CIRR_CombineStatus(cashAStatus, priceStatus)
        outData(r, 21) = outStatus

    Next r

    ' Dump output once
    wsOut.Range("A1").Resize(lastRow, 21).Value2 = outData
    wsOut.Columns.AutoFit

    MsgBox OUTPUT_SHEET & " built from " & INPUT_SHEET & ". Rows: " & (lastRow - 1), vbInformation

End Sub

' ==========================================================
' CASH IRR ANNUITY Ã(R) via schedule (annual, Following, 30E/360)
' start = expiry + 2 business days, then Following
' Ã = Sum_i tau_i / ?_{j=1..i}(1 + R*tau_j)
' ==========================================================
Private Function CIRR_CashIRRAnnuity_Dates(ByVal r As Double, ByVal valDate As Date, ByVal expDate As Date, _
                                          ByVal swapTenY As Double, ByRef st As String) As Double
    Dim nPay As Long, i As Long
    Dim startDate As Date
    Dim prev As Date, payD As Date
    Dim tau As Double
    Dim disc As Double, onePlus As Double
    Dim acc As Double

    st = "OK"
    CIRR_CashIRRAnnuity_Dates = 0#

    nPay = CLng(swapTenY + 0.0001)
    If nPay <= 0 Then
        st = "BAD_TENOR"
        Exit Function
    End If

    startDate = CIRR_AddBusinessDays(expDate, 2)
    startDate = CIRR_AdjustFollowing(startDate)

    prev = startDate
    disc = 1#
    acc = 0#

    For i = 1 To nPay
        payD = DateAdd("yyyy", i, startDate)
        payD = CIRR_AdjustFollowing(payD)

        tau = CIRR_YF_30E360(prev, payD)
        If tau <= 0# Then
            st = "BAD_TAU"
            Exit Function
        End If

        onePlus = 1# + r * tau
        If onePlus <= 0# Then
            st = "DENOM_LE0"
            Exit Function
        End If

        disc = disc * onePlus
        acc = acc + tau / disc

        prev = payD
    Next i

    CIRR_CashIRRAnnuity_Dates = acc
End Function

' ==========================================================
' Price per annuity from vol (model-dependent)
' NOTE: if vol missing/non-numeric/<=0 => NO_VOL (leave hole)
' ==========================================================
Private Function CIRR_PricePerAnnuityFromVol(ByVal modelName As String, ByVal F As Double, ByVal k As Double, _
                                             ByVal t As Double, ByVal vol As Variant, ByVal isCall As Boolean, _
                                             ByVal sh As Double, ByRef st As String) As Variant
    Dim M As String
    Dim sig As Double

    st = "OK"
    M = UCase$(Trim$(modelName))

    If (Not IsNumeric(vol)) Then
        st = "NO_VOL"
        CIRR_PricePerAnnuityFromVol = CVErr(xlErrNA)
        Exit Function
    End If

    sig = CDbl(vol)
    If sig <= 0# Or t <= 0# Then
        st = "NO_VOL"
        CIRR_PricePerAnnuityFromVol = CVErr(xlErrNA)
        Exit Function
    End If

    If M = "NORMAL" Then
        CIRR_PricePerAnnuityFromVol = CIRR_Price_Bachelier(F, k, t, sig, isCall)
    ElseIf (M = "BLACK_SHIFT") Or (M = "SHIFTED_BLACK") Or (M = "BLACK") Then
        CIRR_PricePerAnnuityFromVol = CIRR_Price_ShiftedBlack(F, k, t, sig, isCall, sh)
        If IsError(CIRR_PricePerAnnuityFromVol) Then st = "BLACK_INVALID"
    Else
        st = "UNKNOWN_MODEL"
        CIRR_PricePerAnnuityFromVol = CVErr(xlErrNA)
    End If

End Function

' =========================
' Pricing: Bachelier
' =========================
Private Function CIRR_Price_Bachelier(ByVal F As Double, ByVal k As Double, ByVal t As Double, _
                                      ByVal sigma As Double, ByVal isCall As Boolean) As Double
    Dim srt As Double, d As Double

    srt = sigma * Sqr(t)
    If srt <= 0# Then
        If isCall Then
            CIRR_Price_Bachelier = WorksheetFunction.Max(F - k, 0#)
        Else
            CIRR_Price_Bachelier = WorksheetFunction.Max(k - F, 0#)
        End If
        Exit Function
    End If

    d = (F - k) / srt

    If isCall Then
        CIRR_Price_Bachelier = (F - k) * CIRR_Ncdf(d) + srt * CIRR_Npdf(d)
    Else
        CIRR_Price_Bachelier = (k - F) * CIRR_Ncdf(-d) + srt * CIRR_Npdf(d)
    End If
End Function

' =========================
' Pricing: Shifted Black
' =========================
Private Function CIRR_Price_ShiftedBlack(ByVal F As Double, ByVal k As Double, ByVal t As Double, _
                                         ByVal sigma As Double, ByVal isCall As Boolean, ByVal sh As Double) As Variant
    Dim Fs As Double, Ks As Double
    Dim volT As Double, d1 As Double, d2 As Double

    Fs = F + sh
    Ks = k + sh

    If Fs <= 0# Or Ks <= 0# Then
        CIRR_Price_ShiftedBlack = CVErr(xlErrNum)
        Exit Function
    End If

    volT = sigma * Sqr(t)
    If volT <= 0# Then
        CIRR_Price_ShiftedBlack = CVErr(xlErrNum)
        Exit Function
    End If

    d1 = (Log(Fs / Ks) + 0.5 * volT * volT) / volT
    d2 = d1 - volT

    If isCall Then
        CIRR_Price_ShiftedBlack = Fs * CIRR_Ncdf(d1) - Ks * CIRR_Ncdf(d2)
    Else
        CIRR_Price_ShiftedBlack = Ks * CIRR_Ncdf(-d2) - Fs * CIRR_Ncdf(-d1)
    End If
End Function

' =========================
' Normal CDF/PDF
' =========================
Private Function CIRR_Ncdf(ByVal x As Double) As Double
    CIRR_Ncdf = WorksheetFunction.Norm_S_Dist(x, True)
End Function

Private Function CIRR_Npdf(ByVal x As Double) As Double
    CIRR_Npdf = WorksheetFunction.Norm_S_Dist(x, False)
End Function

' ==========================================================
' DATE / CALENDAR (weekend + holidays) and daycount
' ==========================================================
Private Function CIRR_YF_ACT365(ByVal d1 As Date, ByVal d2 As Date) As Double
    CIRR_YF_ACT365 = DateDiff("d", d1, d2) / 365#
End Function

' 30E/360 (Eurobond basis): D1=31->30, D2=31->30
Private Function CIRR_YF_30E360(ByVal dt1 As Date, ByVal dt2 As Date) As Double
    Dim d1 As Integer, m1 As Integer, y1 As Integer
    Dim d2 As Integer, m2 As Integer, y2 As Integer

    d1 = Day(dt1): m1 = Month(dt1): y1 = Year(dt1)
    d2 = Day(dt2): m2 = Month(dt2): y2 = Year(dt2)

    If d1 = 31 Then d1 = 30
    If d2 = 31 Then d2 = 30

    CIRR_YF_30E360 = (360# * (y2 - y1) + 30# * (m2 - m1) + (d2 - d1)) / 360#
End Function

Private Function CIRR_IsWeekend(ByVal d As Date) As Boolean
    Dim wd As Integer
    wd = Weekday(d, vbMonday) ' 1=Mon ... 7=Sun
    CIRR_IsWeekend = (wd >= 6)
End Function

Private Sub CIRR_Load_Holidays(ByVal wsCal As Worksheet)
    ' expects holiday dates in column E from row 2 down
    Dim r As Long, n As Long
    Dim tmp() As Long

    ReDim tmp(1 To 5000)
    n = 0
    r = 2

    Do While wsCal.Cells(r, 5).Value <> ""
        If IsDate(wsCal.Cells(r, 5).Value) Then
            n = n + 1
            tmp(n) = CLng(CDate(wsCal.Cells(r, 5).Value))
        End If
        r = r + 1
        If r > wsCal.Rows.Count Then Exit Do
    Loop

    If n = 0 Then
        gHolN = 0
        Exit Sub
    End If

    gHolN = n
    ReDim gHol(1 To gHolN)

    Dim i As Long
    For i = 1 To gHolN
        gHol(i) = tmp(i)
    Next i

    CIRR_SortLongArray gHol, 1, gHolN
End Sub

Private Sub CIRR_SortLongArray(ByRef a() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Long, t As Long
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then CIRR_SortLongArray a, lo, j
    If i < hi Then CIRR_SortLongArray a, i, hi
End Sub

Private Function CIRR_IsHoliday(ByVal d As Date) As Boolean
    If gHolN = 0 Then CIRR_IsHoliday = False: Exit Function

    Dim x As Long: x = CLng(d)
    Dim lo As Long, hi As Long, mid As Long

    lo = 1: hi = gHolN
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If gHol(mid) = x Then
            CIRR_IsHoliday = True
            Exit Function
        ElseIf gHol(mid) < x Then
            lo = mid + 1
        Else
            hi = mid - 1
        End If
    Loop

    CIRR_IsHoliday = False
End Function

Private Function CIRR_IsBusinessDay(ByVal d As Date) As Boolean
    CIRR_IsBusinessDay = (Not CIRR_IsWeekend(d)) And (Not CIRR_IsHoliday(d))
End Function

Private Function CIRR_AdjustFollowing(ByVal d As Date) As Date
    Dim x As Date: x = d
    Do While Not CIRR_IsBusinessDay(x)
        x = x + 1
    Loop
    CIRR_AdjustFollowing = x
End Function

Private Function CIRR_AddBusinessDays(ByVal d As Date, ByVal n As Long) As Date
    Dim x As Date: x = d
    Dim k As Long: k = 0
    Do While k < n
        x = x + 1
        If CIRR_IsBusinessDay(x) Then k = k + 1
    Loop
    CIRR_AddBusinessDays = x
End Function

' Expiry date from label like "1M", "18M", "2Y", ...
Private Function CIRR_ExpiryDate(ByVal valDate As Date, ByVal expLbl As String) As Date
    Dim s As String: s = LCase$(Trim$(expLbl))
    Dim n As Long: n = CLng(val(s))
    Dim u As String: u = Right$(s, 1)

    If u = "m" Then
        CIRR_ExpiryDate = DateAdd("m", n, valDate)
    ElseIf u = "y" Then
        CIRR_ExpiryDate = DateAdd("yyyy", n, valDate)
    Else
        Err.Raise vbObjectError + 210, , "Bad expiry label: " & expLbl
    End If
End Function

' =========================
' Utilities
' =========================
Private Function CIRR_IsCall(ByVal optType As String) As Boolean
    Dim s As String
    s = UCase$(Trim$(optType))
    CIRR_IsCall = (InStr(1, s, "CALL", vbTextCompare) > 0 Or InStr(1, s, "PAYER", vbTextCompare) > 0)
End Function

Private Function CIRR_CombineStatus(ByVal s1 As String, ByVal s2 As String) As String
    If s1 = "OK" And s2 = "OK" Then
        CIRR_CombineStatus = "OK"
    Else
        CIRR_CombineStatus = s1 & "|" & s2
    End If
End Function

Private Function Nz(ByVal v As Variant, ByVal defaultValue As Variant) As Variant
    If IsError(v) Then
        Nz = defaultValue
    ElseIf IsEmpty(v) Then
        Nz = defaultValue
    ElseIf Trim$(CStr(v)) = "" Then
        Nz = defaultValue
    Else
        Nz = v
    End If
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

Private Function GetOrCreateAndReset(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = sheetName
    End If

    ws.Cells.Clear
    Set GetOrCreateAndReset = ws
End Function


