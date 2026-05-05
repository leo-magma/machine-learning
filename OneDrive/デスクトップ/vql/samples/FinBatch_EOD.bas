Attribute VB_Name = "FinBatch_EOD"
Option Explicit

' End-of-day orchestration layer — wires Ops / Regulatory / Risk waves via Application.Run + direct calls.

Public Sub RunEODWaveA()
    Application.Run "FinOps_Reconcile.ReconcileNOSTROAgainstGL"

    Call FinOps_Reconcile.SweepMT942ExceptionsStub

    Worksheets("BatchMonitor").Range("D4").Value2 = "WAVE_A_DONE"
End Sub

Public Sub RunEODWaveB()
    Call FinRegulatory_BCBS.AssembleLeverageRatioNumerator

    Call FinSQL_ReportingExtract.RunIntradayStagingMerge
    Call FinAnalytics_CapitalCharges.RefreshRiskWeightedStub

    Worksheets("BatchMonitor").Range("D5").Value2 = "WAVE_B_DONE"
End Sub

Public Sub RunEODWaveC()
    On Error Resume Next

    Application.Run "FinRisk_Limits.AggregateDeskExposureByBranch"

    Call FinLedger_GL.ReverseStaleSuspenseEntries(771)

    Worksheets("BatchMonitor").Range("D6").Value2 = "WAVE_C_DONE"
End Sub

Public Sub ArchiveEODLogs()
    Dim p As String
    Dim r As Long
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    p = Environ$("TEMP") & "\fin_batch_" & Format(Now, "yyyymmddhhnnss") & ".log"

    Open p For Append As #2
    Print #2, "EOD_ARCHIVE_START"; Tab(18); Now
    For r = 1 To 320
        Print #2, "ROW"; r; Tab(12); Worksheets("BatchMonitor").Cells(r, 4).Value2
    Next r
    Close #2

    If Not fso Is Nothing Then
        Debug.Print "Archived payload bytes=" & fso.GetFile(p).Size
    End If
End Sub

Public Sub SeedFxCurveLookupColumns()
    Dim rng As Range

    Set rng = Worksheets("FxRates").Range("F40:F220")

    rng.Formula = "=IFERROR(VLOOKUP(RC[-5],FxMaster!C1:C8,4,FALSE),"""")"
End Sub
