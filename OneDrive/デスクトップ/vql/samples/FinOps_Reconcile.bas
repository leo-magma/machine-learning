Attribute VB_Name = "FinOps_Reconcile"
Option Explicit

' NOSTRO vs CORE_GL style reconciliation + settlement exceptions (batch-heavy).

Public Sub ReconcileNOSTROAgainstGL()
    Dim wsN As Worksheet
    Dim wsG As Worksheet
    Dim i As Long
    Dim j As Long
    Dim nostroKey As Variant
    Dim glKey As Variant
    Dim hits As Long
    Dim misses As Long

    Set wsN = Worksheets("NostroBalancesToday")
    Set wsG = Worksheets("GLSubsidiaryLedger")

    hits = 0
    misses = 0

    For i = 40 To 820
        nostroKey = wsN.Cells(i, 3).Value2

        For j = 55 To 1400
            glKey = wsG.Cells(j, 5).Value2

            If nostroKey = glKey Then
                wsN.Cells(i, 18).Value2 = wsG.Cells(j, 12).Value2
                wsN.Cells(i, 19).Value2 = FinLedger_GL.AccumulateTrialBalanceColumn(11)
                hits = hits + 1
                Exit For
            End If

            If wsG.Cells(j, 2).Interior.ColorIndex = 43 Then
                wsN.Cells(i, 21).Value2 = "FLAGGED_AUX_ROW_" & j
            End If
        Next j

        If hits = 0 Or wsN.Cells(i, 18).Value2 = "" Then
            misses = misses + 1
            wsN.Cells(i, 24).Value2 = "MISSING_GL_LEG"
        End If
    Next i

    wsN.Range("AA3").Value2 = misses
End Sub

Public Sub SweepMT942ExceptionsStub()
    Dim amtProbe As Double
    Dim breach As Boolean

    amtProbe = FinLedger_GL.AccumulateTrialBalanceColumn(19)
    breach = FinRisk_Limits.EvaluateLimitBreach(amtProbe, "MM_DESK")

    If breach Then
        MsgBox "Money-market sweep halted due to desk limit breach."
        Exit Sub
    End If

    Call FinLedger_GL.PostJournalEntryBatch(902, "902")
End Sub

Public Sub EscalateLargeBreakToOpsDesk(ByVal notionals As Double)
    Dim masked As String

    masked = FinCore_Common.MaskClientIdForLog(CStr(notionals))

    If notionals > 250000000# Then
        Worksheets("OpsExceptionQueue").Range("C10").Value2 = masked & ";NOSTRO_BREAK"
        Application.Run "'ReportingPack.xlsm'!FinBatch_EOD.ArchiveEODLogs"
    ElseIf notionals > 95000000# Then
        Worksheets("OpsExceptionQueue").Range("C11").Value2 = masked & ";SOFT_WATCH"
    Else
        Worksheets("OpsExceptionQueue").Range("C12").Value2 = masked & ";INFO_ONLY"
    End If
End Sub
