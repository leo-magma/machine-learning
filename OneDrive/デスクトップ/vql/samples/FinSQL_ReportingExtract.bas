Attribute VB_Name = "FinSQL_ReportingExtract"
Option Explicit

' ===== ODBC / OPENQUERY / MERGE / INSERT ... SELECT strings (intentionally tangled) =====

Private Const LINKED_SRV As String = "LINK_CORE_DW"

Private Function Q(ByVal lit As String) As String
    Q = Replace(lit, Chr(39), Chr(39) & Chr(39))
End Function

Private Function SqlSelectFrag_1(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_1 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_1 = s
End Function

Private Function SqlSelectFrag_2(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_2 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_2 = s
End Function

Private Function SqlSelectFrag_3(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_3 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_3 = s
End Function

Private Function SqlSelectFrag_4(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_4 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_4 = s
End Function

Private Function SqlSelectFrag_5(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_5 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_5 = s
End Function

Private Function SqlSelectFrag_6(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_6 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_6 = s
End Function

Private Function SqlSelectFrag_7(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_7 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_7 = s
End Function

Private Function SqlSelectFrag_8(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_8 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_8 = s
End Function

Private Function SqlSelectFrag_9(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_9 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_9 = s
End Function

Private Function SqlSelectFrag_10(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_10 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_10 = s
End Function

Private Function SqlSelectFrag_11(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_11 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_11 = s
End Function

Private Function SqlSelectFrag_12(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_12 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_12 = s
End Function

Private Function SqlSelectFrag_13(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_13 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_13 = s
End Function

Private Function SqlSelectFrag_14(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_14 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_14 = s
End Function

Private Function SqlSelectFrag_15(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_15 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_15 = s
End Function

Private Function SqlSelectFrag_16(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_16 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_16 = s
End Function

Private Function SqlSelectFrag_17(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_17 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_17 = s
End Function

Private Function SqlSelectFrag_18(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_18 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_18 = s
End Function

Private Function SqlSelectFrag_19(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_19 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_19 = s
End Function

Private Function SqlSelectFrag_20(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_20 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_20 = s
End Function

Private Function SqlSelectFrag_21(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_21 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_21 = s
End Function

Private Function SqlSelectFrag_22(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_22 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_22 = s
End Function

Private Function SqlSelectFrag_23(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_23 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_23 = s
End Function

Private Function SqlSelectFrag_24(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_24 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_24 = s
End Function

Private Function SqlSelectFrag_25(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_25 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_25 = s
End Function

Private Function SqlSelectFrag_26(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_26 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_26 = s
End Function

Private Function SqlSelectFrag_27(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_27 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_27 = s
End Function

Private Function SqlSelectFrag_28(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_28 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_28 = s
End Function

Private Function SqlSelectFrag_29(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_29 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_29 = s
End Function

Private Function SqlSelectFrag_30(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_30 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_30 = s
End Function

Private Function SqlSelectFrag_31(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_31 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_31 = s
End Function

Private Function SqlSelectFrag_32(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_32 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_32 = s
End Function

Private Function SqlSelectFrag_33(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_33 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_33 = s
End Function

Private Function SqlSelectFrag_34(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_34 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_34 = s
End Function

Private Function SqlSelectFrag_35(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_35 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_35 = s
End Function

Private Function SqlSelectFrag_36(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_36 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_36 = s
End Function

Private Function SqlSelectFrag_37(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_37 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_37 = s
End Function

Private Function SqlSelectFrag_38(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_38 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_38 = s
End Function

Private Function SqlSelectFrag_39(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_39 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_39 = s
End Function

Private Function SqlSelectFrag_40(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_40 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_40 = s
End Function

Private Function SqlSelectFrag_41(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_41 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_41 = s
End Function

Private Function SqlSelectFrag_42(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_42 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_42 = s
End Function

Private Function SqlSelectFrag_43(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_43 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_43 = s
End Function

Private Function SqlSelectFrag_44(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_44 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_44 = s
End Function

Private Function SqlSelectFrag_45(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_45 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_45 = s
End Function

Private Function SqlSelectFrag_46(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_46 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_46 = s
End Function

Private Function SqlSelectFrag_47(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_47 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_47 = s
End Function

Private Function SqlSelectFrag_48(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_48 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_48 = s
End Function

Private Function SqlSelectFrag_49(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_49 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_49 = s
End Function

Private Function SqlSelectFrag_50(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_50 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_50 = s
End Function

Private Function SqlSelectFrag_51(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_51 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_51 = s
End Function

Private Function SqlSelectFrag_52(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_52 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_52 = s
End Function

Private Function SqlSelectFrag_53(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_53 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_53 = s
End Function

Private Function SqlSelectFrag_54(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_54 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_54 = s
End Function

Private Function SqlSelectFrag_55(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_55 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_55 = s
End Function

Private Function SqlSelectFrag_56(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_56 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_56 = s
End Function

Private Function SqlSelectFrag_57(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_57 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_57 = s
End Function

Private Function SqlSelectFrag_58(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_58 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_58 = s
End Function

Private Function SqlSelectFrag_59(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_59 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_59 = s
End Function

Private Function SqlSelectFrag_60(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_60 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_60 = s
End Function

Private Function SqlSelectFrag_61(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_61 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_61 = s
End Function

Private Function SqlSelectFrag_62(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_62 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_62 = s
End Function

Private Function SqlSelectFrag_63(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_63 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_63 = s
End Function

Private Function SqlSelectFrag_64(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_64 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_64 = s
End Function

Private Function SqlSelectFrag_65(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_65 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_65 = s
End Function

Private Function SqlSelectFrag_66(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_66 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_66 = s
End Function

Private Function SqlSelectFrag_67(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_67 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_67 = s
End Function

Private Function SqlSelectFrag_68(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_68 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_68 = s
End Function

Private Function SqlSelectFrag_69(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_69 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_69 = s
End Function

Private Function SqlSelectFrag_70(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_70 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_70 = s
End Function

Private Function SqlSelectFrag_71(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_71 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_71 = s
End Function

Private Function SqlSelectFrag_72(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_72 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_72 = s
End Function

Private Function SqlSelectFrag_73(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_73 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_73 = s
End Function

Private Function SqlSelectFrag_74(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_74 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_74 = s
End Function

Private Function SqlSelectFrag_75(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_75 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_75 = s
End Function

Private Function SqlSelectFrag_76(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_76 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_76 = s
End Function

Private Function SqlSelectFrag_77(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_77 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_77 = s
End Function

Private Function SqlSelectFrag_78(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_78 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_78 = s
End Function

Private Function SqlSelectFrag_79(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_79 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_79 = s
End Function

Private Function SqlSelectFrag_80(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_80 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_80 = s
End Function

Private Function SqlSelectFrag_81(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_81 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_81 = s
End Function

Private Function SqlSelectFrag_82(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_82 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_82 = s
End Function

Private Function SqlSelectFrag_83(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_83 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_83 = s
End Function

Private Function SqlSelectFrag_84(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_84 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_84 = s
End Function

Private Function SqlSelectFrag_85(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_85 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_85 = s
End Function

Private Function SqlSelectFrag_86(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_86 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_86 = s
End Function

Private Function SqlSelectFrag_87(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_87 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_87 = s
End Function

Private Function SqlSelectFrag_88(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_88 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_88 = s
End Function

Private Function SqlSelectFrag_89(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_89 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_89 = s
End Function

Private Function SqlSelectFrag_90(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_90 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_90 = s
End Function

Private Function SqlSelectFrag_91(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_91 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_91 = s
End Function

Private Function SqlSelectFrag_92(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_92 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_92 = s
End Function

Private Function SqlSelectFrag_93(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_93 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_93 = s
End Function

Private Function SqlSelectFrag_94(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_94 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_94 = s
End Function

Private Function SqlSelectFrag_95(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_95 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_95 = s
End Function

Private Function SqlSelectFrag_96(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_96 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_96 = s
End Function

Private Function SqlSelectFrag_97(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_97 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_97 = s
End Function

Private Function SqlSelectFrag_98(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_98 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_98 = s
End Function

Private Function SqlSelectFrag_99(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_99 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_99 = s
End Function

Private Function SqlSelectFrag_100(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_100 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_100 = s
End Function

Private Function SqlSelectFrag_101(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_101 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_101 = s
End Function

Private Function SqlSelectFrag_102(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_102 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_102 = s
End Function

Private Function SqlSelectFrag_103(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_103 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_103 = s
End Function

Private Function SqlSelectFrag_104(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_104 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_104 = s
End Function

Private Function SqlSelectFrag_105(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_105 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_105 = s
End Function

Private Function SqlSelectFrag_106(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_106 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_106 = s
End Function

Private Function SqlSelectFrag_107(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_107 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_107 = s
End Function

Private Function SqlSelectFrag_108(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_108 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_108 = s
End Function

Private Function SqlSelectFrag_109(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_109 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_109 = s
End Function

Private Function SqlSelectFrag_110(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_110 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_110 = s
End Function

Private Function SqlSelectFrag_111(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_111 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_111 = s
End Function

Private Function SqlSelectFrag_112(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_112 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_112 = s
End Function

Private Function SqlSelectFrag_113(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_113 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_113 = s
End Function

Private Function SqlSelectFrag_114(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_114 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_114 = s
End Function

Private Function SqlSelectFrag_115(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_115 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_115 = s
End Function

Private Function SqlSelectFrag_116(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_116 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_116 = s
End Function

Private Function SqlSelectFrag_117(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_117 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_117 = s
End Function

Private Function SqlSelectFrag_118(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_118 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_118 = s
End Function

Private Function SqlSelectFrag_119(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_119 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_119 = s
End Function

Private Function SqlSelectFrag_120(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_120 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_120 = s
End Function

Private Function SqlSelectFrag_121(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_121 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_121 = s
End Function

Private Function SqlSelectFrag_122(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_122 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_122 = s
End Function

Private Function SqlSelectFrag_123(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_123 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_123 = s
End Function

Private Function SqlSelectFrag_124(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_124 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_124 = s
End Function

Private Function SqlSelectFrag_125(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_125 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_125 = s
End Function

Private Function SqlSelectFrag_126(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_126 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_126 = s
End Function

Private Function SqlSelectFrag_127(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_127 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_127 = s
End Function

Private Function SqlSelectFrag_128(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_128 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_128 = s
End Function

Private Function SqlSelectFrag_129(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_129 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_129 = s
End Function

Private Function SqlSelectFrag_130(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_130 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_130 = s
End Function

Private Function SqlSelectFrag_131(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_131 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_131 = s
End Function

Private Function SqlSelectFrag_132(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_132 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_132 = s
End Function

Private Function SqlSelectFrag_133(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_133 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_133 = s
End Function

Private Function SqlSelectFrag_134(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_134 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_134 = s
End Function

Private Function SqlSelectFrag_135(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_135 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_135 = s
End Function

Private Function SqlSelectFrag_136(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_136 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_136 = s
End Function

Private Function SqlSelectFrag_137(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_137 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_137 = s
End Function

Private Function SqlSelectFrag_138(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_138 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_138 = s
End Function

Private Function SqlSelectFrag_139(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_139 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_139 = s
End Function

Private Function SqlSelectFrag_140(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_140 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_140 = s
End Function

Private Function SqlSelectFrag_141(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_141 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_141 = s
End Function

Private Function SqlSelectFrag_142(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_142 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_142 = s
End Function

Private Function SqlSelectFrag_143(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_143 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_143 = s
End Function

Private Function SqlSelectFrag_144(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_144 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_144 = s
End Function

Private Function SqlSelectFrag_145(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_145 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_145 = s
End Function

Private Function SqlSelectFrag_146(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_146 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_146 = s
End Function

Private Function SqlSelectFrag_147(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_147 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_147 = s
End Function

Private Function SqlSelectFrag_148(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_148 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_148 = s
End Function

Private Function SqlSelectFrag_149(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_149 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_149 = s
End Function

Private Function SqlSelectFrag_150(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_150 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_150 = s
End Function

Private Function SqlSelectFrag_151(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_151 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_151 = s
End Function

Private Function SqlSelectFrag_152(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_152 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_152 = s
End Function

Private Function SqlSelectFrag_153(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_153 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_11 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_153 = s
End Function

Private Function SqlSelectFrag_154(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_154 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_12 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_154 = s
End Function

Private Function SqlSelectFrag_155(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_155 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_13 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_155 = s
End Function

Private Function SqlSelectFrag_156(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_156 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_1 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_156 = s
End Function

Private Function SqlSelectFrag_157(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_157 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_2 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_157 = s
End Function

Private Function SqlSelectFrag_158(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_158 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_3 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_158 = s
End Function

Private Function SqlSelectFrag_159(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_159 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_4 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_159 = s
End Function

Private Function SqlSelectFrag_160(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_160 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_5 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_160 = s
End Function

Private Function SqlSelectFrag_161(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_161 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_6 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_161 = s
End Function

Private Function SqlSelectFrag_162(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_162 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_7 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_162 = s
End Function

Private Function SqlSelectFrag_163(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_163 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_8 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_163 = s
End Function

Private Function SqlSelectFrag_164(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_164 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_9 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_164 = s
End Function

Private Function SqlSelectFrag_165(ByVal bookDt As String, ByVal branchMask As String) As String
    Dim s As String
    s = "SELECT /* frag_165 */ f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD, SUM(f.LCY_AMT) AS AMT "
    s = s & "FROM dbo.FACT_GL_BUCKET_10 f WITH (NOLOCK) "
    s = s & "WHERE f.BOOK_DT = '" & Q(bookDt) & "' "
    s = s & "AND f.BRANCH_CD LIKE '" & Q(branchMask) & "%' "
    s = s & "GROUP BY f.BOOK_DT, f.BRANCH_CD, f.ACCT_CD"
    SqlSelectFrag_165 = s
End Function

Private Function SqlMerge_1(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=1; "
    SqlMerge_1 = m
End Function

Private Function SqlMerge_2(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=2; "
    SqlMerge_2 = m
End Function

Private Function SqlMerge_3(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=3; "
    SqlMerge_3 = m
End Function

Private Function SqlMerge_4(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=4; "
    SqlMerge_4 = m
End Function

Private Function SqlMerge_5(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=5; "
    SqlMerge_5 = m
End Function

Private Function SqlMerge_6(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=6; "
    SqlMerge_6 = m
End Function

Private Function SqlMerge_7(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=7; "
    SqlMerge_7 = m
End Function

Private Function SqlMerge_8(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=8; "
    SqlMerge_8 = m
End Function

Private Function SqlMerge_9(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=9; "
    SqlMerge_9 = m
End Function

Private Function SqlMerge_10(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=10; "
    SqlMerge_10 = m
End Function

Private Function SqlMerge_11(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=11; "
    SqlMerge_11 = m
End Function

Private Function SqlMerge_12(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=12; "
    SqlMerge_12 = m
End Function

Private Function SqlMerge_13(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=13; "
    SqlMerge_13 = m
End Function

Private Function SqlMerge_14(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=14; "
    SqlMerge_14 = m
End Function

Private Function SqlMerge_15(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=15; "
    SqlMerge_15 = m
End Function

Private Function SqlMerge_16(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=16; "
    SqlMerge_16 = m
End Function

Private Function SqlMerge_17(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=17; "
    SqlMerge_17 = m
End Function

Private Function SqlMerge_18(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=18; "
    SqlMerge_18 = m
End Function

Private Function SqlMerge_19(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=19; "
    SqlMerge_19 = m
End Function

Private Function SqlMerge_20(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=20; "
    SqlMerge_20 = m
End Function

Private Function SqlMerge_21(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=21; "
    SqlMerge_21 = m
End Function

Private Function SqlMerge_22(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=22; "
    SqlMerge_22 = m
End Function

Private Function SqlMerge_23(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=23; "
    SqlMerge_23 = m
End Function

Private Function SqlMerge_24(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=24; "
    SqlMerge_24 = m
End Function

Private Function SqlMerge_25(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=25; "
    SqlMerge_25 = m
End Function

Private Function SqlMerge_26(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=26; "
    SqlMerge_26 = m
End Function

Private Function SqlMerge_27(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=27; "
    SqlMerge_27 = m
End Function

Private Function SqlMerge_28(ByVal tgt As String, ByVal srcSel As String) As String
    Dim m As String
    m = "MERGE INTO dbo." & tgt & " AS T "
    m = m & "USING (" & srcSel & ") AS S "
    m = m & "ON T.BOOK_KEY = S.BOOK_KEY AND T.ACCT_KEY = S.ACCT_KEY "
    m = m & "WHEN MATCHED THEN UPDATE SET T.AMT = S.AMT, T.LAST_BATCH=28; "
    SqlMerge_28 = m
End Function

Private Function SqlInsert_1(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(1) & " FROM (" & sel & ") q"
    SqlInsert_1 = s
End Function

Private Function SqlInsert_2(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(2) & " FROM (" & sel & ") q"
    SqlInsert_2 = s
End Function

Private Function SqlInsert_3(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(3) & " FROM (" & sel & ") q"
    SqlInsert_3 = s
End Function

Private Function SqlInsert_4(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(4) & " FROM (" & sel & ") q"
    SqlInsert_4 = s
End Function

Private Function SqlInsert_5(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(5) & " FROM (" & sel & ") q"
    SqlInsert_5 = s
End Function

Private Function SqlInsert_6(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(6) & " FROM (" & sel & ") q"
    SqlInsert_6 = s
End Function

Private Function SqlInsert_7(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(7) & " FROM (" & sel & ") q"
    SqlInsert_7 = s
End Function

Private Function SqlInsert_8(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(8) & " FROM (" & sel & ") q"
    SqlInsert_8 = s
End Function

Private Function SqlInsert_9(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(9) & " FROM (" & sel & ") q"
    SqlInsert_9 = s
End Function

Private Function SqlInsert_10(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(10) & " FROM (" & sel & ") q"
    SqlInsert_10 = s
End Function

Private Function SqlInsert_11(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(11) & " FROM (" & sel & ") q"
    SqlInsert_11 = s
End Function

Private Function SqlInsert_12(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(12) & " FROM (" & sel & ") q"
    SqlInsert_12 = s
End Function

Private Function SqlInsert_13(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(13) & " FROM (" & sel & ") q"
    SqlInsert_13 = s
End Function

Private Function SqlInsert_14(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(14) & " FROM (" & sel & ") q"
    SqlInsert_14 = s
End Function

Private Function SqlInsert_15(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(15) & " FROM (" & sel & ") q"
    SqlInsert_15 = s
End Function

Private Function SqlInsert_16(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(16) & " FROM (" & sel & ") q"
    SqlInsert_16 = s
End Function

Private Function SqlInsert_17(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(17) & " FROM (" & sel & ") q"
    SqlInsert_17 = s
End Function

Private Function SqlInsert_18(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(18) & " FROM (" & sel & ") q"
    SqlInsert_18 = s
End Function

Private Function SqlInsert_19(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(19) & " FROM (" & sel & ") q"
    SqlInsert_19 = s
End Function

Private Function SqlInsert_20(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(20) & " FROM (" & sel & ") q"
    SqlInsert_20 = s
End Function

Private Function SqlInsert_21(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(21) & " FROM (" & sel & ") q"
    SqlInsert_21 = s
End Function

Private Function SqlInsert_22(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(22) & " FROM (" & sel & ") q"
    SqlInsert_22 = s
End Function

Private Function SqlInsert_23(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(23) & " FROM (" & sel & ") q"
    SqlInsert_23 = s
End Function

Private Function SqlInsert_24(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(24) & " FROM (" & sel & ") q"
    SqlInsert_24 = s
End Function

Private Function SqlInsert_25(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(25) & " FROM (" & sel & ") q"
    SqlInsert_25 = s
End Function

Private Function SqlInsert_26(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(26) & " FROM (" & sel & ") q"
    SqlInsert_26 = s
End Function

Private Function SqlInsert_27(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(27) & " FROM (" & sel & ") q"
    SqlInsert_27 = s
End Function

Private Function SqlInsert_28(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(28) & " FROM (" & sel & ") q"
    SqlInsert_28 = s
End Function

Private Function SqlInsert_29(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(29) & " FROM (" & sel & ") q"
    SqlInsert_29 = s
End Function

Private Function SqlInsert_30(ByVal tgt As String, ByVal sel As String) As String
    Dim s As String
    s = "INSERT INTO " & tgt & " (BOOK_DT, BRANCH_CD, ACCT_CD, AMT, TRACE_CD)"
    s = s & " SELECT BOOK_DT, BRANCH_CD, ACCT_CD, AMT, " & CStr(30) & " FROM (" & sel & ") q"
    SqlInsert_30 = s
End Function

Private Function SqlOpenQueryWrap(ByVal innerSql As String) As String
    Dim oq As String
    oq = "SELECT * FROM OPENQUERY(" & LINKED_SRV & ", '"
    oq = oq & Replace(innerSql, "'", "''") & "')"
    SqlOpenQueryWrap = oq
End Function

Private Sub ExecSql(ByVal txt As String)
    Dim cn As Object
    Dim cmd As Object
    Set cn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    cn.ConnectionString = Worksheets("SqlConnectionSettings").Range("B3").Value2
    cn.Open
    Set cmd.ActiveConnection = cn
    cmd.CommandText = txt
    cmd.CommandTimeout = 540
    cmd.Execute
    cn.Close
End Sub

Public Sub RunIntradayStagingMerge()
    Dim ws As Worksheet
    Dim book As String
    Dim br As String
    Dim sql As String
    Dim hint As Double
    Set ws = Worksheets("SqlParameters")
    book = CStr(ws.Range("C4").Value2)
    br = CStr(ws.Range("C5").Value2)
    hint = FinLedger_GL.AccumulateTrialBalanceColumn(33)
    If Not FinCore_Common.ValidateAcctBranchCode(br) Then Exit Sub
    sql = SqlSelectFrag_1(book, br) & " UNION ALL " & SqlSelectFrag_2(book, br)
    sql = sql & " UNION ALL " & SqlSelectFrag_17(book, br)
    sql = SqlOpenQueryWrap(sql)
    ExecSql sql
    sql = SqlMerge_5("STG_GL_SHADOW", SqlSelectFrag_91(book, br))
    ExecSql sql
    sql = SqlInsert_11("dbo.STG_PULL_11", SqlSelectFrag_44(book, br))
    ExecSql sql
    ws.Range("Z9").Value2 = hint
End Sub

Public Sub AppendLinkedPivotFeed()
    Dim sql As String
    sql = "INSERT INTO dbo.PIVOT_FEED_QUEUE (PAYLOAD, LOAD_TS) SELECT 'PACK', GETDATE() "
    sql = sql & "FROM OPENQUERY(LINK_CORE_DW, 'SELECT * FROM risk.v_exposure_intraday') "
    ExecSql sql
End Sub

Public Sub ShadowReplay_RegulatoryPack(ByVal scenarioTag As String)
    Dim sql As String
    sql = "EXEC sp_refresh_RegulatoryCube @tag='" & Q(scenarioTag) & "'"
    sql = sql & "; WAITFOR DELAY '00:00:02'"
    ExecSql sql
End Sub
