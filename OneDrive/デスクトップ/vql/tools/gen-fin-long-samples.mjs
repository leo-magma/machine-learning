/**
 * Generates long finance-themed .bas samples (hundreds of lines each).
 * Run from repo root: node tools/gen-fin-long-samples.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const samples = path.join(root, "samples");

function writeFinSql() {
  let o = "";
  o += 'Attribute VB_Name = "FinSQL_ReportingExtract"\r\nOption Explicit\r\n\r\n';
  o +=
    "' ===== ODBC / OPENQUERY / MERGE / INSERT ... SELECT strings (intentionally tangled) =====\r\n\r\n";
  o += 'Private Const LINKED_SRV As String = "LINK_CORE_DW"\r\n\r\n';

  o += "Private Function Q(ByVal lit As String) As String\r\n";
  o += "    Q = Replace(lit, Chr(39), Chr(39) & Chr(39))\r\n";
  o += "End Function\r\n\r\n";

  for (let k = 1; k <= 165; k++) {
    const bucket = (k % 13) + 1;
    o += `Private Function SqlSelectFrag_${k}(ByVal bookDt As String, ByVal branchMask As String) As String\r\n`;
    o += "    Dim s As String\r\n";
    o += `    s = "SELECT /* frag_${k} */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "\r\n`;
    o += `    s = s & "FROM dbo.FACT_GL_BUCKET_${bucket} f WITH (NOLOCK) "\r\n`;
    o += `    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "\r\n`;
    o += `    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "\r\n`;
    o += '    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"\r\n';
    o += `    SqlSelectFrag_${k} = s\r\n`;
    o += "End Function\r\n\r\n";
  }

  for (let k = 1; k <= 28; k++) {
    o += `Private Function SqlMerge_${k}(ByVal tgt As String, ByVal srcSel As String) As String\r\n`;
    o += "    Dim m As String\r\n";
    o += '    m = "MERGE INTO dbo." & tgt & " AS T "\r\n';
    o += '    m = m & "USING (" & srcSel & ") AS S "\r\n';
    o += '    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "\r\n';
    o += `    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=${k}; "\r\n`;
    o += `    SqlMerge_${k} = m\r\n`;
    o += "End Function\r\n\r\n";
  }

  for (let k = 1; k <= 30; k++) {
    o += `Private Function SqlInsert_${k}(ByVal tgt As String, ByVal sel As String) As String\r\n`;
    o += "    Dim s As String\r\n";
    o += '    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"\r\n';
    o +=
      "    s = s & \" SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, \" & CStr(" +
      k +
      ') & " FROM (" & sel & ") q"\r\n';
    o += `    SqlInsert_${k} = s\r\n`;
    o += "End Function\r\n\r\n";
  }

  o += "Private Function SqlOpenQueryWrap(ByVal innerSql As String) As String\r\n";
  o += "    Dim oq As String\r\n";
  o += `    oq = "SELECT * FROM OPENQUERY(" & LINKED_SRV & ", '"\r\n`;
  o += `    oq = oq & Replace(innerSql, "'", "''") & "')"\r\n`;
  o += "    SqlOpenQueryWrap = oq\r\n";
  o += "End Function\r\n\r\n";

  o += "Private Sub ExecSql(ByVal txt As String)\r\n";
  o += "    Dim cn As Object\r\n";
  o += "    Dim cmd As Object\r\n";
  o += "    Set cn = CreateObject(\"ADODB.Connection\")\r\n";
  o += "    Set cmd = CreateObject(\"ADODB.Command\")\r\n";
  o += "    cn.ConnectionString = Worksheets(\"SqlConnectionSettings\").Range(\"B3\").Value2\r\n";
  o += "    cn.Open\r\n";
  o += "    Set cmd.ActiveConnection = cn\r\n";
  o += "    cmd.CommandText = txt\r\n";
  o += "    cmd.CommandTimeout = 540\r\n";
  o += "    cmd.Execute\r\n";
  o += "    cn.Close\r\n";
  o += "End Sub\r\n\r\n";

  o += "Public Sub RunIntradayStagingMerge()\r\n";
  o += "    Dim ws As Worksheet\r\n";
  o += "    Dim book As String\r\n";
  o += "    Dim br As String\r\n";
  o += "    Dim sql As String\r\n";
  o += "    Dim hint As Double\r\n";
  o += "    Set ws = Worksheets(\"SqlParameters\")\r\n";
  o += "    book = CStr(ws.Range(\"C4\").Value2)\r\n";
  o += "    br = CStr(ws.Range(\"C5\").Value2)\r\n";
  o += "    hint = FinLedger_GL.AccumulateTrialBalanceColumn(33)\r\n";
  o += "    If Not FinCore_Common.ValidateAcctBranchCode(br) Then Exit Sub\r\n";
  o += "    sql = SqlSelectFrag_1(book, br) & \" UNION ALL \" & SqlSelectFrag_2(book, br)\r\n";
  o += "    sql = sql & \" UNION ALL \" & SqlSelectFrag_17(book, br)\r\n";
  o += "    sql = SqlOpenQueryWrap(sql)\r\n";
  o += "    ExecSql sql\r\n";
  o += "    sql = SqlMerge_5(\"STG_GL_SHADOW\", SqlSelectFrag_91(book, br))\r\n";
  o += "    ExecSql sql\r\n";
  o += '    sql = SqlInsert_11("dbo.STG_PULL_11", SqlSelectFrag_44(book, br))\r\n';
  o += "    ExecSql sql\r\n";
  o += "    ws.Range(\"Z9\").Value2 = hint\r\n";
  o += "End Sub\r\n\r\n";

  o += "Public Sub AppendLinkedPivotFeed()\r\n";
  o += "    Dim sql As String\r\n";
  o +=
    '    sql = "INSERT INTO dbo.PIVOT_FEED_QUEUE (PAYLOAD, LOAD_TS) SELECT \'PACK\', GETDATE() "\r\n';
  o +=
    '    sql = sql & "FROM OPENQUERY(LINK_CORE_DW, \'SELECT * FROM risk.v_exposure_intraday\') "\r\n';
  o += "    ExecSql sql\r\n";
  o += "End Sub\r\n\r\n";

  o += "Public Sub ShadowReplay_RegulatoryPack(ByVal scenarioTag As String)\r\n";
  o += "    Dim sql As String\r\n";
  o += "    sql = \"EXEC sp_refresh_RegulatoryCube @tag='\" & Q(scenarioTag) & \"'\"\r\n";
  o += "    sql = sql & \"; WAITFOR DELAY '00:00:02'\"\r\n";
  o += "    ExecSql sql\r\n";
  o += "End Sub\r\n";

  fs.writeFileSync(path.join(samples, "FinSQL_ReportingExtract.bas"), o, "utf8");
}

function writeFinCapital() {
  let o = "";
  o += 'Attribute VB_Name = "FinAnalytics_CapitalCharges"\r\nOption Explicit\r\n\r\n';
  o +=
    "' ===== Pseudo Basel / migration grids - dense arithmetic & branching =====\r\n\r\n";

  o += "Private Const VAT_GATE As Double = 0.831\r\n";
  o += "Private Const FX_SHOCK_HI As Double = 1.147\r\n";
  o += "Private Const FX_SHOCK_LO As Double = 0.923\r\n\r\n";

  for (let band = 1; band <= 72; band++) {
    const mod17 = band % 17;
    o += `Private Function TierRW_${band}(ByVal pd As Double, ByVal lgd As Double) As Double\r\n`;
    o += "    Dim rw As Double\r\n";
    o += `    rw = ${mod17} * 0.012 + pd * lgd * 9.81 + CDbl(${band}) * 0.000001\r\n`;
    o += `    TierRW_${band} = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")\r\n`;
    o += "End Function\r\n\r\n";
  }

  o += "Public Sub RefreshRiskWeightedStub()\r\n";
  o += "    Dim ws As Worksheet\r\n";
  o += "    Dim r As Long\r\n";
  o += "    Dim ead As Double\r\n";
  o += "    Dim lgd As Double\r\n";
  o += "    Dim pd As Double\r\n";
  o += "    Dim rwa As Double\r\n";
  o += "    Dim mig As Long\r\n";
  o += "    Set ws = Worksheets(\"RiskWeightedAssetsWork\")\r\n";
  o += "    For r = 180 To 980\r\n";
  o += "        pd = ws.Cells(r, 6).Value2\r\n";
  o += "        lgd = ws.Cells(r, 8).Value2\r\n";
  o += "        ead = ws.Cells(r, 10).Value2\r\n";
  o += "        mig = ws.Cells(r, 12).Value2\r\n";
  o += "        Select Case mig\r\n";
  o += "            Case 1\r\n";
  o += "                rwa = ead * TierRW_3(pd, lgd)\r\n";
  o += "            Case 2, 3\r\n";
  o += "                rwa = ead * TierRW_11(pd, lgd) * VAT_GATE\r\n";
  o += "            Case 4 To 8\r\n";
  o += "                rwa = ead * TierRW_27(pd, lgd) * FX_SHOCK_HI\r\n";
  o += "            Case Else\r\n";
  o += "                rwa = ead * TierRW_41(pd, lgd) * FX_SHOCK_LO\r\n";
  o += "        End Select\r\n";
  o += "        ws.Cells(r, 28).Value2 = rwa\r\n";
  o += "    Next r\r\n";
  o += "End Sub\r\n\r\n";

  o += "Public Sub MonteCarloSpreadStub(ByVal paths As Long)\r\n";
  o += "    Dim i As Long\r\n";
  o += "    Dim acc As Double\r\n";
  o += "    Dim bump As Double\r\n";
  o += "    Dim idx As Long\r\n";
  o += "    acc = 0\r\n";
  o += "    For i = 1 To paths\r\n";
  o += "        bump = Sin(CDbl(i) / 17#) * Cos(CDbl(i) / 23#)\r\n";
  o += "        idx = ((i - 1) Mod 72) + 1\r\n";
  o += "        Select Case idx\r\n";
  for (let c = 1; c <= 18; c++) {
    o += `            Case ${c}\r\n`;
    o += `                acc = acc + bump * TierRW_${c}(0.021 + bump * 0.001, 0.35)\r\n`;
  }
  o += "            Case Else\r\n";
  o += "                acc = acc + bump * TierRW_61(0.019 + bump * 0.002, 0.41)\r\n";
  o += "        End Select\r\n";
  o += "    Next i\r\n";
  o += "    Worksheets(\"RiskWeightedAssetsWork\").Range(\"AA5\").Value2 = acc\r\n";
  o += "End Sub\r\n\r\n";

  o += "Public Sub StressNestedBuckets(ByVal depth As Long)\r\n";
  o += "    Dim i As Long\r\n";
  o += "    Dim j As Long\r\n";
  o += "    Dim acc As Double\r\n";
  o += "    acc = FinLedger_GL.AccumulateTrialBalanceColumn(58)\r\n";
  o += "    For i = 1 To depth\r\n";
  o += "        For j = 1 To depth\r\n";
  o += "            acc = acc + TierRW_19(0.014 + CDbl(i + j) * 0.00001, 0.37)\r\n";
  o += "        Next j\r\n";
  o += "    Next i\r\n";
  o += "    Worksheets(\"RiskWeightedAssetsWork\").Range(\"AA7\").Value2 = acc\r\n";
  o += "End Sub\r\n";

  fs.writeFileSync(path.join(samples, "FinAnalytics_CapitalCharges.bas"), o, "utf8");
}

writeFinSql();
writeFinCapital();
console.error(
  "Wrote samples/FinSQL_ReportingExtract.bas and samples/FinAnalytics_CapitalCharges.bas"
);
