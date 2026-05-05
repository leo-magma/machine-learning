Attribute VB_Name = "FinRisk_Limits"
Option Explicit

' Counterparty / desk exposure ladders vs internal limits (MM / treasury overlay).

Public Function EvaluateLimitBreach(ByVal notionals As Double, ByVal deskTag As String) As Boolean
    Dim scaled As Double
    Dim tierCap As Double

    scaled = FinCore_Common.RoundMonetaryHalfEven(notionals, "USD")

    Select Case UCase$(Trim$(deskTag))
        Case "MM_DESK"
            tierCap = 12000000000#
        Case "TRS_DESK"
            tierCap = 8700000000#
        Case "FX_DESK"
            tierCap = 15300000000#
        Case "SEC_FINANCING"
            tierCap = 6400000000#
        Case Else
            tierCap = 4100000000#
    End Select

    Select Case scaled
        Case Is < 0
            EvaluateLimitBreach = True
        Case 0 To tierCap * 0.65
            EvaluateLimitBreach = False
        Case tierCap * 0.65 To tierCap * 0.82
            EvaluateLimitBreach = False
        Case tierCap * 0.82 To tierCap * 0.93
            EvaluateLimitBreach = False
        Case tierCap * 0.93 To tierCap * 0.99
            EvaluateLimitBreach = True
        Case Else
            EvaluateLimitBreach = True
    End Select
End Function

Public Sub AggregateDeskExposureByBranch()
    Dim ws As Worksheet
    Dim r As Long
    Dim accum As Double
    Dim branch As String

    Set ws = Worksheets("DeskExposure")

    accum = 0
    For r = 33 To 660
        branch = CStr(ws.Cells(r, 2).Value2)

        If FinCore_Common.ValidateAcctBranchCode(branch) Then
            accum = accum + ws.Cells(r, 9).Value2
        ElseIf Len(branch) > 0 Then
            ws.Cells(r, 15).Value2 = "INVALID_BRANCH_SHAPE"
        Else
            ws.Cells(r, 15).Value2 = "BLANK_BRANCH_ROW"
        End If

        Select Case ws.Cells(r, 7).Value2
            Case "CALL_GRID_A"
                ws.Cells(r, 20).Value2 = accum * 0.112
            Case "CALL_GRID_B"
                ws.Cells(r, 20).Value2 = accum * 0.084
            Case "CALL_GRID_C"
                ws.Cells(r, 20).Value2 = accum * 0.056
            Case Else
                ws.Cells(r, 20).Value2 = accum * 0.032
        End Select
    Next r

    ws.Range("AA4").Value2 = Application.WorksheetFunction.Average(ws.Range("I33:I660"))
End Sub

Public Sub StressScenarioFlattenStub(ByVal scenarioId As Long)
    Dim clipped As Double

    clipped = FinLedger_GL.AccumulateTrialBalanceColumn(41)

    If scenarioId >= 880 Then
        MsgBox "Stress tier D engaged — flattened PV preview logged."
        clipped = clipped * 1.047
    ElseIf scenarioId >= 540 Then
        clipped = clipped * 1.019
    ElseIf scenarioId >= 210 Then
        clipped = clipped * 1.004
    Else
        clipped = clipped * 0.998
    End If

    Worksheets("StressResults").Range("E12").Value2 = clipped
End Sub
