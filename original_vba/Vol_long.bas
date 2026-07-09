Attribute VB_Name = "Vol_long"
Option Explicit

' ==========================================================
' VOL_LONG_ALL BUILDER (NORMAL + SHIFTED-BLACK) from Swaption Physical forward premium cube
'
' Data:
' - Market quotes date (valuation): 31/10/2019 (fixed, dataset date)
' - OIS curve: "IR Yield Curves" cols E:F (days, cont. zero rate)
' - Holidays: sheet "calendar", column A (dates), from row 2 down
'
' Conventions:
' - Expiry date from label (1M, 3M, 1Y, ...)
' - Te = ACT/365(valuation, expiry)
' - Underlying swap start = expiry + 2 business days (calendar-aware)
' - Fixed leg: annual pay, Following adjustment, 30E/360
' - OIS discounting; annuity forwardized: A(Te)=A(0)/DF(0,Te)
'
' Notes:
' - Floating schedule is not needed because ATM forwards are read from Swaptions Physical.
' ==========================================================

' ============================
' Global OIS curve storage
' ============================
Private gT() As Double   ' maturities in years (ACT/365)
Private gR() As Double   ' continuous zero rates
Private gN As Long

' ============================
' Fixed valuation date
' ============================
Private gValDate As Date

' ============================
' Holidays storage (serial dates sorted)
' ============================
Private gHol() As Long
Private gHolN As Long

' ============================
' USER CONFIG
' ============================
Private Function VL2_ShiftsList() As Variant
    VL2_ShiftsList = Array(0.01, 0.02, 0.03, 0.05)
End Function

' ============================
' MAIN ENTRY
' ============================
Public Sub VL2_BuildVolLong_All()

    Dim wsP As Worksheet, wsC As Worksheet, wsCal As Worksheet, wsOut As Worksheet
    Set wsP = ThisWorkbook.Worksheets("Swaptions Physical")
    Set wsC = ThisWorkbook.Worksheets("IR Yield Curves")
    Set wsCal = ThisWorkbook.Worksheets("Calendar")
    Set wsOut = VL2_GetOrCreateAndReset(ThisWorkbook, "VOL_LONG_ALL")

    ' Market data date (fixed)
    gValDate = DateSerial(2019, 10, 31)

    ' Load inputs
    VL2_Load_OIS_Curve wsC
    VL2_Load_Holidays wsCal

    ' Output
    VL2_WriteHeader wsOut

    Dim NextRow As Long
    NextRow = 2

    NextRow = VL2_Append_ATM_Long_All(wsP, wsOut, NextRow)

    ' Titles must match your sheet
    NextRow = VL2_Append_Skew_Long_All(wsP, wsOut, NextRow, "EUR Gamma - Strangles", "GAMMA")
    NextRow = VL2_Append_Skew_Long_All(wsP, wsOut, NextRow, "EUR Vega - Strangles", "VEGA")

    wsOut.Columns.AutoFit
    MsgBox "VOL_LONG_ALL built. Rows: " & (NextRow - 2), vbInformation

End Sub

' ==========================================================
' SHEET HELPERS
' ==========================================================
Private Function VL2_GetOrCreateAndReset(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = sheetName
    End If

    ws.Cells.Clear
    Set VL2_GetOrCreateAndReset = ws
End Function

' ==========================================================
' OUTPUT
' ==========================================================
Private Sub VL2_WriteHeader(ByVal ws As Worksheet)
    ws.Range("A1").Resize(1, 17).Value = Array( _
        "SourceBlock", "Instrument", "ExpiryLbl", "TenorLbl", "Te", "SwapTenorY", "FwdRate", _
        "MoneynessBP", "Strike", "OptType", "PremiumBP", "Annuity_Te", "PricePerAnnuity", _
        "Model", "Shift", "ImplVol", "Status")
End Sub

Private Function VL2_AppendRow( _
    ByVal wsOut As Worksheet, ByVal r As Long, _
    ByVal sourceBlock As String, ByVal instName As String, _
    ByVal expLbl As String, ByVal tenLbl As String, ByVal Te As Double, ByVal swapTenY As Double, _
    ByVal F As Double, ByVal mBP As Double, ByVal strikeK As Double, ByVal optType As String, _
    ByVal premBP As Double, ByVal ATe As Double, ByVal pricePA As Double, _
    ByVal model As String, ByVal sh As Double, ByVal vol As Variant, ByVal status As String _
) As Long

    With wsOut
        .Cells(r, 1).Value = sourceBlock
        .Cells(r, 2).Value = instName
        .Cells(r, 3).Value = expLbl
        .Cells(r, 4).Value = tenLbl
        .Cells(r, 5).Value = Te
        .Cells(r, 6).Value = swapTenY
        .Cells(r, 7).Value = F
        .Cells(r, 8).Value = mBP
        .Cells(r, 9).Value = strikeK
        .Cells(r, 10).Value = optType
        .Cells(r, 11).Value = premBP
        .Cells(r, 12).Value = ATe
        .Cells(r, 13).Value = pricePA
        .Cells(r, 14).Value = model
        .Cells(r, 15).Value = sh
        If IsError(vol) Then
            .Cells(r, 16).Value = ""
        Else
            .Cells(r, 16).Value = vol
        End If
        .Cells(r, 17).Value = status
    End With

    VL2_AppendRow = r + 1
End Function

' ==========================================================
' SAFE NUMERIC PARSING (comma/dot)
' ==========================================================
Private Function VL2_TryDbl(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = False
    If IsError(v) Or IsEmpty(v) Then Exit Function

    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Or s = "-" Then Exit Function

    s = Replace(s, " ", "")
    s = Replace(s, ".", Application.International(xlDecimalSeparator))
    s = Replace(s, ",", Application.International(xlDecimalSeparator))

    If Not IsNumeric(s) Then Exit Function

    VL2_TryDbl = CDbl(s)
    ok = True
End Function

' ==========================================================
' DATE CONVENTIONS + CALENDAR (weekend + holidays)
' ==========================================================
Private Function VL2_YF_ACT365(ByVal d1 As Date, ByVal d2 As Date) As Double
    VL2_YF_ACT365 = DateDiff("d", d1, d2) / 365#
End Function

' 30E/360 (Eurobond basis): D1=31->30, D2=31->30
Private Function VL2_YF_30E360(ByVal dt1 As Date, ByVal dt2 As Date) As Double
    Dim d1 As Integer, m1 As Integer, y1 As Integer
    Dim d2 As Integer, m2 As Integer, y2 As Integer

    d1 = Day(dt1): m1 = Month(dt1): y1 = Year(dt1)
    d2 = Day(dt2): m2 = Month(dt2): y2 = Year(dt2)

    If d1 = 31 Then d1 = 30
    If d2 = 31 Then d2 = 30

    VL2_YF_30E360 = (360# * (y2 - y1) + 30# * (m2 - m1) + (d2 - d1)) / 360#
End Function

Private Function VL2_IsWeekend(ByVal d As Date) As Boolean
    Dim wd As Integer
    wd = Weekday(d, vbMonday) ' 1=Mon ... 7=Sun
    VL2_IsWeekend = (wd >= 6)
End Function

Private Sub VL2_Load_Holidays(ByVal wsCal As Worksheet)
    ' expects holiday dates in column E from row 2 down
    Dim r As Long, n As Long
    Dim tmp() As Long
    ReDim tmp(1 To 5000)

    n = 0
    r = 2
    Do While wsCal.Cells(r, 5).Value <> ""
        If IsDate(wsCal.Cells(r, 5).Value) Then
            n = n + 1
            tmp(n) = CLng(CDate(wsCal.Cells(r, 5).Value)) ' serial date
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

    VL2_SortLongArray gHol, 1, gHolN
End Sub

Private Sub VL2_SortLongArray(ByRef a() As Long, ByVal lo As Long, ByVal hi As Long)
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
    If lo < j Then VL2_SortLongArray a, lo, j
    If i < hi Then VL2_SortLongArray a, i, hi
End Sub

Private Function VL2_IsHoliday(ByVal d As Date) As Boolean
    If gHolN = 0 Then
        VL2_IsHoliday = False
        Exit Function
    End If

    Dim x As Long
    x = CLng(d)

    Dim lo As Long, hi As Long, mid As Long
    lo = 1: hi = gHolN
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If gHol(mid) = x Then
            VL2_IsHoliday = True
            Exit Function
        ElseIf gHol(mid) < x Then
            lo = mid + 1
        Else
            hi = mid - 1
        End If
    Loop

    VL2_IsHoliday = False
End Function

Private Function VL2_IsBusinessDay(ByVal d As Date) As Boolean
    VL2_IsBusinessDay = (Not VL2_IsWeekend(d)) And (Not VL2_IsHoliday(d))
End Function

Private Function VL2_AdjustFollowing(ByVal d As Date) As Date
    Dim x As Date: x = d
    Do While Not VL2_IsBusinessDay(x)
        x = x + 1
    Loop
    VL2_AdjustFollowing = x
End Function

Private Function VL2_AddBusinessDays(ByVal d As Date, ByVal n As Long) As Date
    Dim x As Date: x = d
    Dim k As Long: k = 0
    Do While k < n
        x = x + 1
        If VL2_IsBusinessDay(x) Then k = k + 1
    Loop
    VL2_AddBusinessDays = x
End Function

' Expiry date from label like "1M", "18M", "2Y", ...
Private Function VL2_ExpiryDate(ByVal valDate As Date, ByVal expLbl As String) As Date
    Dim s As String: s = LCase$(Trim$(expLbl))
    Dim n As Long: n = CLng(val(s))
    Dim u As String: u = Right$(s, 1)

    If u = "m" Then
        VL2_ExpiryDate = DateAdd("m", n, valDate)
    ElseIf u = "y" Then
        VL2_ExpiryDate = DateAdd("yyyy", n, valDate)
    Else
        Err.Raise vbObjectError + 210, , "Bad expiry label: " & expLbl
    End If
End Function

' ==========================================================
' CURVE: OIS (cols E:F, row 3 down) -> DF = exp(-r*T)
' col E: days, col F: cont. zero rate
' ==========================================================
Private Sub VL2_Load_OIS_Curve(ByVal wsCurve As Worksheet)
    Dim r As Long, n As Long
    Dim tmpT() As Double, tmpR() As Double

    r = 3
    n = 0
    ReDim tmpT(1 To 5000)
    ReDim tmpR(1 To 5000)

    Do While wsCurve.Cells(r, 5).Value <> "" And wsCurve.Cells(r, 6).Value <> ""
        If IsNumeric(wsCurve.Cells(r, 5).Value) And IsNumeric(wsCurve.Cells(r, 6).Value) Then
            n = n + 1
            tmpT(n) = CDbl(wsCurve.Cells(r, 5).Value) / 365#   ' days -> years ACT/365
            tmpR(n) = CDbl(wsCurve.Cells(r, 6).Value)          ' continuous zero rate
        End If
        r = r + 1
        If r > wsCurve.Rows.Count Then Exit Do
    Loop

    If n < 2 Then Err.Raise vbObjectError + 101, , "OIS curve not found / too short."

    gN = n
    ReDim gT(1 To gN)
    ReDim gR(1 To gN)

    Dim i As Long
    For i = 1 To gN
        gT(i) = tmpT(i)
        gR(i) = tmpR(i)
    Next i
End Sub

Private Function VL2_ZeroRate_OIS(ByVal t As Double) As Double
    If t <= gT(1) Then VL2_ZeroRate_OIS = gR(1): Exit Function
    If t >= gT(gN) Then VL2_ZeroRate_OIS = gR(gN): Exit Function

    Dim i As Long
    For i = 1 To gN - 1
        If t >= gT(i) And t <= gT(i + 1) Then
            Dim w As Double
            w = (t - gT(i)) / (gT(i + 1) - gT(i))
            VL2_ZeroRate_OIS = gR(i) + w * (gR(i + 1) - gR(i))
            Exit Function
        End If
    Next i
End Function

Private Function VL2_DF_OIS(ByVal t As Double) As Double
    VL2_DF_OIS = exp(-VL2_ZeroRate_OIS(t) * t)
End Function

' ==========================================================
' TENOR / CODE PARSING
' ==========================================================
Private Function VL2_TenorToYears(ByVal lbl As String) As Double
    Dim s As String, u As String
    Dim n As Double

    s = LCase$(Trim$(lbl))
    s = Replace(s, "opt", "")
    s = Replace(s, " ", "")
    s = Replace(s, "x", "")

    u = Right$(s, 1)
    n = CDbl(val(s))

    If u = "y" Then
        VL2_TenorToYears = n
    ElseIf u = "m" Then
        VL2_TenorToYears = n / 12#
    Else
        Err.Raise vbObjectError + 201, , "Cannot parse tenor: " & lbl
    End If
End Function

Private Sub VL2_ParseInstrumentCode(ByVal code As String, ByRef expLbl As String, ByRef tenLbl As String)
    Dim s As String
    Dim i As Long, posExpUnit As Long
    Dim ch As String

    s = LCase$(Trim$(code))
    s = Replace(s, " ", "")
    s = Replace(s, "x", "")

    posExpUnit = 0
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        If ch = "m" Or ch = "y" Then
            posExpUnit = i
            Exit For
        End If
    Next i
    If posExpUnit = 0 Then Err.Raise vbObjectError + 301, , "Bad instrument code: " & code

    expLbl = UCase$(Left$(s, posExpUnit))
    tenLbl = UCase$(mid$(s, posExpUnit + 1))
    If Right$(LCase$(tenLbl), 1) <> "y" Then Err.Raise vbObjectError + 302, , "Bad tenor in code: " & code
End Sub

Private Function VL2_IsInstrumentCode(ByVal v As Variant) As Boolean
    If VarType(v) <> vbString Then Exit Function
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    If s = "" Then Exit Function
    s = Replace(s, " ", "")
    s = Replace(s, "x", "")
    VL2_IsInstrumentCode = (s Like "*m*y" Or s Like "*y*y")
End Function

Private Function VL2_IsBlockTitle(ByVal v As Variant) As Boolean
    If VarType(v) <> vbString Then Exit Function
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    If s = "" Then Exit Function

    If Left$(s, 3) <> "eur" Then Exit Function
    If InStr(1, s, "strangle", vbTextCompare) = 0 Then Exit Function

    If (InStr(1, s, "gamma", vbTextCompare) > 0) Or (InStr(1, s, "vega", vbTextCompare) > 0) Then
        VL2_IsBlockTitle = True
    End If
End Function

' ==========================================================
' FORWARD RATE LOOKUP (reads from Swaptions Physical "ATM swaption forwards" block)
' Returns decimal rate (e.g., 0.0123)
' ==========================================================
Private Function VL2_ForwardRate_FromSheet(ByVal ws As Worksheet, ByVal expLbl As String, ByVal tenLbl As String) As Double
    Dim ur As Range
    Set ur = ws.UsedRange

    Dim baseRow As Long
    baseRow = 0

    Dim r As Long, c As Long
    For r = ur.Row To ur.Row + ur.Rows.Count - 1
        Dim hasEUR As Boolean, hasATM As Boolean, hasSwap As Boolean, hasFwd As Boolean
        hasEUR = False: hasATM = False: hasSwap = False: hasFwd = False

        For c = 1 To 12
            Dim v As Variant
            v = ws.Cells(r, c).Value
            If VarType(v) = vbString Then
                Dim s As String
                s = LCase$(Trim$(CStr(v)))
                If s Like "*eur*" Then hasEUR = True
                If s Like "*atm*" Then hasATM = True
                If s Like "*swaption*" Then hasSwap = True
                If s Like "*forwards*" Or s Like "*forward*" Then hasFwd = True
            End If
        Next c

        If hasEUR And hasATM And hasSwap And hasFwd Then
            baseRow = r
            Exit For
        End If
    Next r

    If baseRow = 0 Then Err.Raise vbObjectError + 401, , "Cannot locate forward rates block."

    Dim hdrRow As Long
    hdrRow = baseRow + 1

    Dim tenCol As Long
    tenCol = 0
    For c = 1 To ur.Column + ur.Columns.Count
        If UCase$(Trim$(CStr(ws.Cells(hdrRow, c).Value))) = UCase$(tenLbl) Then
            tenCol = c
            Exit For
        End If
    Next c
    If tenCol = 0 Then Err.Raise vbObjectError + 403, , "Cannot find tenor col: " & tenLbl

    Dim expRow As Long
    expRow = 0
    For r = hdrRow + 1 To hdrRow + 600
        If Trim$(CStr(ws.Cells(r, 2).Value)) = "" Then Exit For
        If UCase$(Trim$(CStr(ws.Cells(r, 2).Value))) = UCase$(expLbl) And LCase$(Trim$(CStr(ws.Cells(r, 3).Value))) = "opt" Then
            expRow = r
            Exit For
        End If
    Next r
    If expRow = 0 Then Err.Raise vbObjectError + 404, , "Cannot find expiry row: " & expLbl

    Dim ok As Boolean, x As Double
    x = VL2_TryDbl(ws.Cells(expRow, tenCol).Value, ok)
    If Not ok Then Err.Raise vbObjectError + 405, , "Forward cell not numeric for " & expLbl & " " & tenLbl

    VL2_ForwardRate_FromSheet = x / 100#
End Function

' ==========================================================
' ANNUITY FORWARD at Te using DATES + schedule (calendar-aware)
' Fixed leg:
' - start = expiry + 2 business days (weekend+holiday)
' - annual payments
' - Following adjustment (weekend+holiday)
' - daycount 30E/360
'
' A0 = sum DF(0,payDate_i) * tau_i
' A(Te) = A0 / DF(0,Te)
' ==========================================================
Private Function VL2_AnnuityForward_ATe_Dates(ByVal valDate As Date, ByVal expDate As Date, ByVal swapTenY As Double) As Double

    Dim nPay As Long
    nPay = CLng(swapTenY + 0.0001)
    If nPay <= 0 Then Err.Raise vbObjectError + 520, , "Invalid swap tenor years."

    Dim startDate As Date
    startDate = VL2_AddBusinessDays(expDate, 2)
    startDate = VL2_AdjustFollowing(startDate)

    Dim a0 As Double: a0 = 0#
    Dim i As Long
    Dim prev As Date, payD As Date
    prev = startDate

    For i = 1 To nPay
        payD = DateAdd("yyyy", i, startDate)
        payD = VL2_AdjustFollowing(payD)

        Dim tau As Double
        tau = VL2_YF_30E360(prev, payD)

        Dim tPay As Double
        tPay = VL2_YF_ACT365(valDate, payD)

        a0 = a0 + VL2_DF_OIS(tPay) * tau
        prev = payD
    Next i

    Dim TeYF As Double
    TeYF = VL2_YF_ACT365(valDate, expDate)

    Dim P0Te As Double
    P0Te = VL2_DF_OIS(TeYF)
    If P0Te <= 0# Then Err.Raise vbObjectError + 501, , "Invalid DF(0,Te)."

    VL2_AnnuityForward_ATe_Dates = a0 / P0Te
End Function

' ==========================================================
' PRICING + IMPLIED VOLS
' ==========================================================
Private Function VL2_Ncdf(ByVal x As Double) As Double
    VL2_Ncdf = WorksheetFunction.Norm_S_Dist(x, True)
End Function

Private Function VL2_Npdf(ByVal x As Double) As Double
    VL2_Npdf = WorksheetFunction.Norm_S_Dist(x, False)
End Function

Private Function VL2_Price_Bachelier(ByVal F As Double, ByVal strikeK As Double, ByVal t As Double, ByVal sigma As Double, ByVal isCall As Boolean) As Double
    If t <= 0# Or sigma <= 0# Then
        If isCall Then
            VL2_Price_Bachelier = Application.Max(F - strikeK, 0#)
        Else
            VL2_Price_Bachelier = Application.Max(strikeK - F, 0#)
        End If
        Exit Function
    End If

    Dim srt As Double, d As Double
    srt = sigma * Sqr(t)
    d = (F - strikeK) / srt

    If isCall Then
        VL2_Price_Bachelier = (F - strikeK) * VL2_Ncdf(d) + srt * VL2_Npdf(d)
    Else
        VL2_Price_Bachelier = (strikeK - F) * VL2_Ncdf(-d) + srt * VL2_Npdf(d)
    End If
End Function

Private Function VL2_Price_ShiftedBlack(ByVal F As Double, ByVal strikeK As Double, ByVal t As Double, ByVal sigma As Double, ByVal isCall As Boolean, ByVal sh As Double) As Variant
    Dim Fs As Double, Ks As Double
    Fs = F + sh
    Ks = strikeK + sh

    If Fs <= 0# Or Ks <= 0# Then
        VL2_Price_ShiftedBlack = CVErr(xlErrNum)
        Exit Function
    End If

    If t <= 0# Or sigma <= 0# Then
        If isCall Then
            VL2_Price_ShiftedBlack = Application.Max(F - strikeK, 0#)
        Else
            VL2_Price_ShiftedBlack = Application.Max(strikeK - F, 0#)
        End If
        Exit Function
    End If

    Dim volT As Double, d1 As Double, d2 As Double
    volT = sigma * Sqr(t)
    d1 = (Log(Fs / Ks) + 0.5 * volT * volT) / volT
    d2 = d1 - volT

    If isCall Then
        VL2_Price_ShiftedBlack = Fs * VL2_Ncdf(d1) - Ks * VL2_Ncdf(d2)
    Else
        VL2_Price_ShiftedBlack = Ks * VL2_Ncdf(-d2) - Fs * VL2_Ncdf(-d1)
    End If
End Function

Private Function VL2_ImplVol_Bachelier(ByVal target As Double, ByVal F As Double, ByVal strikeK As Double, ByVal t As Double, ByVal isCall As Boolean, ByRef status As String) As Variant
    status = "OK"

    Dim intrinsic As Double
    If isCall Then
        intrinsic = Application.Max(F - strikeK, 0#)
    Else
        intrinsic = Application.Max(strikeK - F, 0#)
    End If

    If target < intrinsic - 0.000000000001 Then
        status = "BelowIntrinsic"
        VL2_ImplVol_Bachelier = CVErr(xlErrNum)
        Exit Function
    End If

    If Abs(target - intrinsic) <= 0.000000000001 Then
        VL2_ImplVol_Bachelier = 0#
        Exit Function
    End If

    Dim lo As Double, hi As Double, pHi As Double
    lo = 0.00000001
    hi = 0.5
    pHi = VL2_Price_Bachelier(F, strikeK, t, hi, isCall)

    Do While pHi < target
        hi = hi * 2#
        If hi > 50# Then Exit Do
        pHi = VL2_Price_Bachelier(F, strikeK, t, hi, isCall)
    Loop

    If pHi < target Then
        status = "NoBracket"
        VL2_ImplVol_Bachelier = CVErr(xlErrNum)
        Exit Function
    End If

    Dim i As Long
    For i = 1 To 80
        Dim mid As Double, pMid As Double
        mid = 0.5 * (lo + hi)
        pMid = VL2_Price_Bachelier(F, strikeK, t, mid, isCall)
        If pMid > target Then
            hi = mid
        Else
            lo = mid
        End If
        If Abs(hi - lo) < 0.0000000001 Then Exit For
    Next i

    VL2_ImplVol_Bachelier = 0.5 * (lo + hi)
End Function

Private Function VL2_ImplVol_ShiftedBlack(ByVal target As Double, ByVal F As Double, ByVal strikeK As Double, ByVal t As Double, ByVal isCall As Boolean, ByVal sh As Double, ByRef status As String) As Variant
    status = "OK"

    Dim Fs As Double, Ks As Double
    Fs = F + sh
    Ks = strikeK + sh

    If Fs <= 0# Or Ks <= 0# Then
        status = "ShiftTooSmall"
        VL2_ImplVol_ShiftedBlack = CVErr(xlErrNum)
        Exit Function
    End If

    Dim intrinsic As Double
    If isCall Then
        intrinsic = Application.Max(F - strikeK, 0#)
    Else
        intrinsic = Application.Max(strikeK - F, 0#)
    End If

    If target < intrinsic - 0.000000000001 Then
        status = "BelowIntrinsic"
        VL2_ImplVol_ShiftedBlack = CVErr(xlErrNum)
        Exit Function
    End If

    If Abs(target - intrinsic) <= 0.000000000001 Then
        VL2_ImplVol_ShiftedBlack = 0#
        Exit Function
    End If

    If isCall Then
        If target > Fs + 0.000000000001 Then
            status = "AboveMax"
            VL2_ImplVol_ShiftedBlack = CVErr(xlErrNum)
            Exit Function
        End If
    Else
        If target > Ks + 0.000000000001 Then
            status = "AboveMax"
            VL2_ImplVol_ShiftedBlack = CVErr(xlErrNum)
            Exit Function
        End If
    End If

    Dim lo As Double, hi As Double
    Dim pHi As Variant
    lo = 0.00000001
    hi = 0.5
    pHi = VL2_Price_ShiftedBlack(F, strikeK, t, hi, isCall, sh)

    Do While Not IsError(pHi) And CDbl(pHi) < target
        hi = hi * 2#
        If hi > 50# Then Exit Do
        pHi = VL2_Price_ShiftedBlack(F, strikeK, t, hi, isCall, sh)
    Loop

    If IsError(pHi) Or CDbl(pHi) < target Then
        status = "NoBracket"
        VL2_ImplVol_ShiftedBlack = CVErr(xlErrNum)
        Exit Function
    End If

    Dim i As Long
    For i = 1 To 80
        Dim mid As Double
        Dim pMid As Variant
        mid = 0.5 * (lo + hi)
        pMid = VL2_Price_ShiftedBlack(F, strikeK, t, mid, isCall, sh)

        If IsError(pMid) Then
            hi = mid
        ElseIf CDbl(pMid) > target Then
            hi = mid
        Else
            lo = mid
        End If

        If Abs(hi - lo) < 0.0000000001 Then Exit For
    Next i

    VL2_ImplVol_ShiftedBlack = 0.5 * (lo + hi)
End Function

' ==========================================================
' ATM LONG APPEND (top matrix)
' quote is STRADDLE bp -> payer premium = straddle/2
' ==========================================================
Private Function VL2_Append_ATM_Long_All(ByVal wsP As Worksheet, ByVal wsOut As Worksheet, ByVal NextRow As Long) As Long

    Dim shifts As Variant
    shifts = VL2_ShiftsList()

    Dim lastTenCol As Long
    lastTenCol = 4
    Do While Trim$(CStr(wsP.Cells(4, lastTenCol).Value)) <> ""
        lastTenCol = lastTenCol + 1
    Loop
    lastTenCol = lastTenCol - 1

    Dim r As Long, c As Long
    For r = 5 To 2000

        If Trim$(CStr(wsP.Cells(r, 2).Value)) = "" Then Exit For
        If LCase$(Trim$(CStr(wsP.Cells(r, 3).Value))) <> "opt" Then GoTo NextR

        Dim expLbl As String
        expLbl = UCase$(Trim$(CStr(wsP.Cells(r, 2).Value)))

        Dim expDate As Date
        expDate = VL2_ExpiryDate(gValDate, expLbl)

        Dim Te As Double
        Te = VL2_YF_ACT365(gValDate, expDate)

        For c = 4 To lastTenCol

            Dim tenLbl As String
            tenLbl = UCase$(Trim$(CStr(wsP.Cells(4, c).Value)))
            If tenLbl = "" Then GoTo NextC

            Dim swapTenY As Double
            swapTenY = VL2_TenorToYears(tenLbl)

            Dim ok As Boolean
            Dim straddleBP As Double
            straddleBP = VL2_TryDbl(wsP.Cells(r, c).Value, ok)

            Dim premCallBP As Double
            Dim pricePA As Double
            Dim statusBase As String
            statusBase = "OK"

            If ok Then
                premCallBP = straddleBP / 2#
            Else
                premCallBP = 0#
                statusBase = "MISSING_QUOTE"
            End If

            Dim F As Double, ATe As Double
            On Error Resume Next
            F = VL2_ForwardRate_FromSheet(wsP, expLbl, tenLbl)
            If Err.Number <> 0 Then
                Err.Clear
                statusBase = IIf(statusBase = "OK", "FWD_NOT_FOUND", statusBase & "|FWD_NOT_FOUND")
                F = 0#
            End If

            ATe = VL2_AnnuityForward_ATe_Dates(gValDate, expDate, swapTenY)
            If Err.Number <> 0 Or ATe <= 0# Then
                Err.Clear
                statusBase = IIf(statusBase = "OK", "ANN_NOT_FOUND", statusBase & "|ANN_NOT_FOUND")
                ATe = 0#
            End If
            On Error GoTo 0

            If statusBase = "OK" Then
                pricePA = (premCallBP / 10000#) / ATe
            Else
                pricePA = 0#
            End If

            Dim instName As String
            instName = "ATM " & expLbl & " x " & tenLbl

            Dim stN As String
            Dim volN As Variant
            If statusBase <> "OK" Then
                volN = ""
                stN = statusBase
            Else
                volN = VL2_ImplVol_Bachelier(pricePA, F, F, Te, True, stN)
            End If

            NextRow = VL2_AppendRow(wsOut, NextRow, "ATM", instName, expLbl, tenLbl, Te, swapTenY, F, 0#, F, _
                                    "CALL/PAYER", premCallBP, ATe, pricePA, "NORMAL", 0#, volN, stN)

            Dim i As Long
            For i = LBound(shifts) To UBound(shifts)
                Dim sh As Double
                sh = CDbl(shifts(i))

                Dim stB As String
                Dim volB As Variant
                If statusBase <> "OK" Then
                    volB = ""
                    stB = statusBase
                Else
                    volB = VL2_ImplVol_ShiftedBlack(pricePA, F, F, Te, True, sh, stB)
                End If

                NextRow = VL2_AppendRow(wsOut, NextRow, "ATM", instName, expLbl, tenLbl, Te, swapTenY, F, 0#, F, _
                                        "CALL/PAYER", premCallBP, ATe, pricePA, "BLACK_SHIFT", sh, volB, stB)
            Next i

NextC:
        Next c

NextR:
    Next r

    VL2_Append_ATM_Long_All = NextRow
End Function

' ==========================================================
' SKEW LONG APPEND (scans ALL occurrences of titleText)
' collar/strangle -> payer/receiver premiums
' ==========================================================
Private Function VL2_Append_Skew_Long_All(ByVal wsP As Worksheet, ByVal wsOut As Worksheet, ByVal NextRow As Long, _
                                         ByVal titleText As String, ByVal tag As String) As Long

    Dim shifts As Variant
    shifts = VL2_ShiftsList()

    Dim firstCell As Range, cell As Range, firstAddr As String
    Set firstCell = wsP.Cells.Find(wHat:=titleText, LookIn:=xlValues, LookAt:=xlPart)
    If firstCell Is Nothing Then
        VL2_Append_Skew_Long_All = NextRow
        Exit Function
    End If

    firstAddr = firstCell.Address
    Set cell = firstCell

    Do
        Dim titleRow As Long, mRow As Long, dataRow As Long
        titleRow = cell.Row
        mRow = titleRow + 2
        dataRow = titleRow + 3

        Dim nm As Long
        nm = 0

        Dim mVals() As Double, colColl() As Long, colStr() As Long
        Dim c As Long, ok As Boolean, tmp As Double

        For c = 3 To 8
            tmp = VL2_TryDbl(wsP.Cells(mRow, c).Value, ok)
            If ok Then
                nm = nm + 1
                ReDim Preserve mVals(1 To nm)
                ReDim Preserve colColl(1 To nm)
                ReDim Preserve colStr(1 To nm)
                mVals(nm) = tmp
                colColl(nm) = c
                colStr(nm) = c + 7
            End If
        Next c

        If nm = 0 Then
            Set cell = wsP.Cells.FindNext(cell)
            If cell Is Nothing Then Exit Do
            GoTo ContinueOuter
        End If

        Dim blankStreak As Long
        blankStreak = 0

        Do While dataRow < wsP.Rows.Count

            Dim vB As Variant
            vB = wsP.Cells(dataRow, 2).Value

            If VL2_IsBlockTitle(vB) Then Exit Do

            If Trim$(CStr(vB)) = "" Then
                blankStreak = blankStreak + 1
                If blankStreak >= 4 Then Exit Do
                dataRow = dataRow + 1
                GoTo ContinueLoop
            Else
                blankStreak = 0
            End If

            If Not VL2_IsInstrumentCode(vB) Then
                dataRow = dataRow + 1
                GoTo ContinueLoop
            End If

            Dim code As String
            code = LCase$(Trim$(CStr(vB)))
            code = Replace(code, " ", "")
            code = Replace(code, "x", "")

            Dim expLbl As String, tenLbl As String
            On Error GoTo SkipRow
            VL2_ParseInstrumentCode code, expLbl, tenLbl
            On Error GoTo 0

            Dim expDate As Date
            expDate = VL2_ExpiryDate(gValDate, expLbl)

            Dim Te As Double, swapTenY As Double
            Te = VL2_YF_ACT365(gValDate, expDate)
            swapTenY = VL2_TenorToYears(tenLbl)

            Dim F As Double, ATe As Double
            Dim statusBase As String
            statusBase = "OK"

            On Error Resume Next
            F = VL2_ForwardRate_FromSheet(wsP, expLbl, tenLbl)
            If Err.Number <> 0 Then
                Err.Clear
                statusBase = "FWD_NOT_FOUND"
                F = 0#
            End If

            ATe = VL2_AnnuityForward_ATe_Dates(gValDate, expDate, swapTenY)
            If Err.Number <> 0 Or ATe <= 0# Then
                Err.Clear
                statusBase = IIf(statusBase = "OK", "ANN_NOT_FOUND", statusBase & "|ANN_NOT_FOUND")
                ATe = 0#
            End If
            On Error GoTo 0

            Dim instName As String
            instName = tag & " " & code

            Dim kk As Long
            For kk = 1 To nm

                Dim collarOk As Boolean, strangleOk As Boolean
                Dim collarBP As Double, strangleBP As Double

                collarBP = VL2_TryDbl(wsP.Cells(dataRow, colColl(kk)).Value, collarOk)
                strangleBP = VL2_TryDbl(wsP.Cells(dataRow, colStr(kk)).Value, strangleOk)

                Dim mBP As Double
                mBP = mVals(kk)

                Dim payerBP As Double, recvBP As Double
                Dim Kp As Double, Kr As Double
                Dim pricePA_P As Double, pricePA_R As Double

                Dim statusQuote As String
                statusQuote = statusBase

                If (Not collarOk) Or (Not strangleOk) Then
                    payerBP = 0#
                    recvBP = 0#
                    statusQuote = IIf(statusQuote = "OK", "MISSING_QUOTE", statusQuote & "|MISSING_QUOTE")
                Else
                    payerBP = (strangleBP + collarBP) / 2#
                    recvBP = (strangleBP - collarBP) / 2#
                End If

                Kp = F + mBP / 10000#
                Kr = F - mBP / 10000#

                If statusQuote = "OK" Then
                    pricePA_P = (payerBP / 10000#) / ATe
                    pricePA_R = (recvBP / 10000#) / ATe
                Else
                    pricePA_P = 0#
                    pricePA_R = 0#
                End If

                Dim stNP As String, stNR As String
                Dim volNP As Variant, volNR As Variant

                If statusQuote <> "OK" Then
                    volNP = ""
                    volNR = ""
                    stNP = statusQuote
                    stNR = statusQuote
                Else
                    volNP = VL2_ImplVol_Bachelier(pricePA_P, F, Kp, Te, True, stNP)
                    volNR = VL2_ImplVol_Bachelier(pricePA_R, F, Kr, Te, False, stNR)
                End If

                NextRow = VL2_AppendRow(wsOut, NextRow, tag, instName, expLbl, tenLbl, Te, swapTenY, F, mBP, Kp, _
                                        "CALL/PAYER", payerBP, ATe, pricePA_P, "NORMAL", 0#, volNP, stNP)

                NextRow = VL2_AppendRow(wsOut, NextRow, tag, instName, expLbl, tenLbl, Te, swapTenY, F, -mBP, Kr, _
                                        "PUT/RECEIVER", recvBP, ATe, pricePA_R, "NORMAL", 0#, volNR, stNR)

                Dim i As Long
                For i = LBound(shifts) To UBound(shifts)
                    Dim sh As Double
                    sh = CDbl(shifts(i))

                    Dim stBP As String, stBR As String
                    Dim volBP As Variant, volBR As Variant

                    If statusQuote <> "OK" Then
                        volBP = ""
                        volBR = ""
                        stBP = statusQuote
                        stBR = statusQuote
                    Else
                        volBP = VL2_ImplVol_ShiftedBlack(pricePA_P, F, Kp, Te, True, sh, stBP)
                        volBR = VL2_ImplVol_ShiftedBlack(pricePA_R, F, Kr, Te, False, sh, stBR)
                    End If

                    NextRow = VL2_AppendRow(wsOut, NextRow, tag, instName, expLbl, tenLbl, Te, swapTenY, F, mBP, Kp, _
                                            "CALL/PAYER", payerBP, ATe, pricePA_P, "BLACK_SHIFT", sh, volBP, stBP)

                    NextRow = VL2_AppendRow(wsOut, NextRow, tag, instName, expLbl, tenLbl, Te, swapTenY, F, -mBP, Kr, _
                                            "PUT/RECEIVER", recvBP, ATe, pricePA_R, "BLACK_SHIFT", sh, volBR, stBR)
                Next i

            Next kk

            dataRow = dataRow + 1
            GoTo ContinueLoop

SkipRow:
            On Error GoTo 0
            dataRow = dataRow + 1

ContinueLoop:
        Loop

        Set cell = wsP.Cells.FindNext(cell)
        If cell Is Nothing Then Exit Do

ContinueOuter:
    Loop While cell.Address <> firstAddr

    VL2_Append_Skew_Long_All = NextRow
End Function


