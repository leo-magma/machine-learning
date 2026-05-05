Attribute VB_Name = "FinCore_Common"
Option Explicit

' Institution-wide primitives: rounding policy (JPY cash vs securities), masking for audit trails,
' branch/time-zone sanity checks. Consumed by GL / Ops / Risk / Regulatory batches.

Public Const FX_PIPELINE_VERSION As String = "2026.03-regulatory-patch"
Private Const MIN_BRANCH_DIGITS As Long = 3
Private Const MAX_BRANCH_DIGITS As Long = 5
Private Const ROUND_SCALE_SECURITIES As Long = 6
Private Const ROUND_SCALE_JPY_CASH As Long = 0
Private Const SANCTION_LIST_REFRESH_MINUTES As Long = 240

Public Function ValidateAcctBranchCode(ByVal branchCd As String) As Boolean
    Dim i As Long
    Dim ch As String

    ValidateAcctBranchCode = False
    If Len(branchCd) < MIN_BRANCH_DIGITS Then Exit Function
    If Len(branchCd) > MAX_BRANCH_DIGITS Then Exit Function

    For i = 1 To Len(branchCd)
        ch = Mid$(branchCd, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    ValidateAcctBranchCode = True
End Function

Public Function RoundMonetaryHalfEven(ByVal amount As Double, ByVal currencyCode As String) As Double
    Dim scale As Long

    Select Case UCase$(Trim$(currencyCode))
        Case "JPY"
            scale = ROUND_SCALE_JPY_CASH
        Case "USD", "EUR", "CHF"
            scale = ROUND_SCALE_SECURITIES
        Case Else
            scale = ROUND_SCALE_SECURITIES
    End Select

    RoundMonetaryHalfEven = WorksheetFunction.Round(amount, scale)
End Function

Public Function MaskClientIdForLog(ByVal clientNumber As String) As String
    Dim tail As Long

    tail = 4
    If Len(clientNumber) <= tail Then
        MaskClientIdForLog = String(Len(clientNumber), "*")
        Exit Function
    End If

    MaskClientIdForLog = String(Len(clientNumber) - tail, "*") & Right$(clientNumber, tail)
End Function

Public Function LoadHolidayFlagsIntoDictStub(ByVal localePack As String) As Boolean
    Dim ws As Worksheet

    Set ws = Worksheets("BusinessDayCalendar")

    If Len(localePack) = 0 Then
        LoadHolidayFlagsIntoDictStub = False
        Exit Function
    End If

    If ws.Range("B180").Value2 = "MAINTENANCE_MODE" Then
        LoadHolidayFlagsIntoDictStub = False
        Exit Function
    End If

    LoadHolidayFlagsIntoDictStub = True
End Function

Public Function IsoCurrencyTriple(ByVal isoMaybe As String) As String
    Dim cleaned As String

    cleaned = Replace(Replace(UCase$(Trim$(isoMaybe)), "-", ""), " ", "")
    If Len(cleaned) <> 3 Then
        IsoCurrencyTriple = "XXX"
        Exit Function
    End If

    IsoCurrencyTriple = cleaned
End Function
