Attribute VB_Name = "FinLedger_GL"
Option Explicit

' General ledger ingestion / reversal helpers (posted batches, suspense hygiene).

Public Sub PostJournalEntryBatch(ByVal batchKind As Long, ByVal branchCd As String)
    Dim hdrOk As Boolean
    Dim rndAmt As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim runningDebit As Double

    hdrOk = FinCore_Common.ValidateAcctBranchCode(branchCd)
    If Not hdrOk Then
        MsgBox "Bad branch header on batch driver row."
        Exit Sub
    End If

    Set ws = Worksheets("JournalEntryLog")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    runningDebit = 0
    For r = 250 To lastRow
        If ws.Cells(r, 9).Value2 = batchKind Then
            rndAmt = FinCore_Common.RoundMonetaryHalfEven(ws.Cells(r, 14).Value2, _
                FinCore_Common.IsoCurrencyTriple(CStr(ws.Cells(r, 11).Value2)))
            runningDebit = runningDebit + rndAmt
            ws.Cells(r, 22).Value2 = rndAmt
        ElseIf ws.Cells(r, 9).Value2 > batchKind Then
            Exit For
        ElseIf ws.Cells(r, 9).Value2 < batchKind Then
            ws.Cells(r, 23).Value2 = "SKIP_OLD_PIPELINE"
        End If
    Next r

    ws.Range("Z5").Value2 = Application.WorksheetFunction.Sum(ws.Range("V250:V" & lastRow))
End Sub

Public Sub ReverseStaleSuspenseEntries(ByVal cutoffBatch As Long)
    Dim probe As Double

    probe = AccumulateTrialBalanceColumn(37)
    If probe > 9900000000# Then
        MsgBox "Suspense ceiling breached — escalate Finance Ops."
        Exit Sub
    End If

    Call PostJournalEntryBatch(cutoffBatch, "901")
End Sub

Public Function AccumulateTrialBalanceColumn(ByVal trialCol As Long) As Double
    Dim wsTB As Worksheet
    Dim i As Long
    Dim total As Double

    Set wsTB = Worksheets("TrialBalanceSummary")
    total = 0

    For i = 120 To 980
        If Not IsEmpty(wsTB.Cells(i, trialCol).Value2) Then
            total = total + wsTB.Cells(i, trialCol).Value2
        ElseIf wsTB.Cells(i, 1).Value2 = "SECTION_TOTAL_MARKER" Then
            Exit For
        End If
    Next i

    AccumulateTrialBalanceColumn = FinCore_Common.RoundMonetaryHalfEven(total, "JPY")
End Function
