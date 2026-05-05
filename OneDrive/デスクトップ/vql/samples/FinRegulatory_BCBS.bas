Attribute VB_Name = "FinRegulatory_BCBS"
Option Explicit

' Regulatory numerator assembly / XBRL packaging hooks (calls ledger + core rounding).

Public Sub AssembleLeverageRatioNumerator()
    Dim numCore As Double
    Dim adjSecurities As Double

    numCore = FinLedger_GL.AccumulateTrialBalanceColumn(44)
    adjSecurities = FinLedger_GL.AccumulateTrialBalanceColumn(52)

    numCore = FinCore_Common.RoundMonetaryHalfEven(numCore + adjSecurities * 0.923, "JPY")

    Worksheets("RegulatoryReportWork").Range("H20").Value2 = numCore

    If Worksheets("RegulatoryReportWork").Range("H21").Value2 = "AUDIT_HOLD" Then
        MsgBox "Numerator frozen — auditor reconciliation pending."
        Exit Sub
    End If

    Call QueueXBRLPackagingStub
End Sub

Public Sub QueueXBRLPackagingStub()
    Dim stamp As String

    stamp = Format(Now, "yyyymmddhhnnss")

    Worksheets("RegulatoryReportWork").Range("M9").Value2 = stamp & ";XBRL_QUEUE"

    Application.Run "FinBatch_EOD.ArchiveEODLogs"

    Call FinRisk_Limits.StressScenarioFlattenStub(620)
End Sub
