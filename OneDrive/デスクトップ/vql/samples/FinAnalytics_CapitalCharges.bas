Attribute VB_Name = "FinAnalytics_CapitalCharges"
Option Explicit

' ===== Pseudo Basel / migration grids - dense arithmetic & branching =====

Private Const VAT_GATE As Double = 0.831
Private Const FX_SHOCK_HI As Double = 1.147
Private Const FX_SHOCK_LO As Double = 0.923

Private Function TierRW_1(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 1 * 0.012 + pd * lgd * 9.81 + CDbl(1) * 0.000001
    TierRW_1 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_2(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 2 * 0.012 + pd * lgd * 9.81 + CDbl(2) * 0.000001
    TierRW_2 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_3(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 3 * 0.012 + pd * lgd * 9.81 + CDbl(3) * 0.000001
    TierRW_3 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_4(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 4 * 0.012 + pd * lgd * 9.81 + CDbl(4) * 0.000001
    TierRW_4 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_5(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 5 * 0.012 + pd * lgd * 9.81 + CDbl(5) * 0.000001
    TierRW_5 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_6(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 6 * 0.012 + pd * lgd * 9.81 + CDbl(6) * 0.000001
    TierRW_6 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_7(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 7 * 0.012 + pd * lgd * 9.81 + CDbl(7) * 0.000001
    TierRW_7 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_8(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 8 * 0.012 + pd * lgd * 9.81 + CDbl(8) * 0.000001
    TierRW_8 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_9(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 9 * 0.012 + pd * lgd * 9.81 + CDbl(9) * 0.000001
    TierRW_9 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_10(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 10 * 0.012 + pd * lgd * 9.81 + CDbl(10) * 0.000001
    TierRW_10 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_11(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 11 * 0.012 + pd * lgd * 9.81 + CDbl(11) * 0.000001
    TierRW_11 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_12(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 12 * 0.012 + pd * lgd * 9.81 + CDbl(12) * 0.000001
    TierRW_12 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_13(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 13 * 0.012 + pd * lgd * 9.81 + CDbl(13) * 0.000001
    TierRW_13 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_14(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 14 * 0.012 + pd * lgd * 9.81 + CDbl(14) * 0.000001
    TierRW_14 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_15(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 15 * 0.012 + pd * lgd * 9.81 + CDbl(15) * 0.000001
    TierRW_15 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_16(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 16 * 0.012 + pd * lgd * 9.81 + CDbl(16) * 0.000001
    TierRW_16 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_17(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 0 * 0.012 + pd * lgd * 9.81 + CDbl(17) * 0.000001
    TierRW_17 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_18(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 1 * 0.012 + pd * lgd * 9.81 + CDbl(18) * 0.000001
    TierRW_18 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_19(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 2 * 0.012 + pd * lgd * 9.81 + CDbl(19) * 0.000001
    TierRW_19 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_20(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 3 * 0.012 + pd * lgd * 9.81 + CDbl(20) * 0.000001
    TierRW_20 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_21(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 4 * 0.012 + pd * lgd * 9.81 + CDbl(21) * 0.000001
    TierRW_21 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_22(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 5 * 0.012 + pd * lgd * 9.81 + CDbl(22) * 0.000001
    TierRW_22 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_23(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 6 * 0.012 + pd * lgd * 9.81 + CDbl(23) * 0.000001
    TierRW_23 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_24(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 7 * 0.012 + pd * lgd * 9.81 + CDbl(24) * 0.000001
    TierRW_24 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_25(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 8 * 0.012 + pd * lgd * 9.81 + CDbl(25) * 0.000001
    TierRW_25 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_26(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 9 * 0.012 + pd * lgd * 9.81 + CDbl(26) * 0.000001
    TierRW_26 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_27(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 10 * 0.012 + pd * lgd * 9.81 + CDbl(27) * 0.000001
    TierRW_27 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_28(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 11 * 0.012 + pd * lgd * 9.81 + CDbl(28) * 0.000001
    TierRW_28 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_29(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 12 * 0.012 + pd * lgd * 9.81 + CDbl(29) * 0.000001
    TierRW_29 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_30(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 13 * 0.012 + pd * lgd * 9.81 + CDbl(30) * 0.000001
    TierRW_30 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_31(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 14 * 0.012 + pd * lgd * 9.81 + CDbl(31) * 0.000001
    TierRW_31 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_32(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 15 * 0.012 + pd * lgd * 9.81 + CDbl(32) * 0.000001
    TierRW_32 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_33(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 16 * 0.012 + pd * lgd * 9.81 + CDbl(33) * 0.000001
    TierRW_33 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_34(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 0 * 0.012 + pd * lgd * 9.81 + CDbl(34) * 0.000001
    TierRW_34 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_35(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 1 * 0.012 + pd * lgd * 9.81 + CDbl(35) * 0.000001
    TierRW_35 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_36(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 2 * 0.012 + pd * lgd * 9.81 + CDbl(36) * 0.000001
    TierRW_36 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_37(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 3 * 0.012 + pd * lgd * 9.81 + CDbl(37) * 0.000001
    TierRW_37 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_38(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 4 * 0.012 + pd * lgd * 9.81 + CDbl(38) * 0.000001
    TierRW_38 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_39(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 5 * 0.012 + pd * lgd * 9.81 + CDbl(39) * 0.000001
    TierRW_39 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_40(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 6 * 0.012 + pd * lgd * 9.81 + CDbl(40) * 0.000001
    TierRW_40 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_41(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 7 * 0.012 + pd * lgd * 9.81 + CDbl(41) * 0.000001
    TierRW_41 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_42(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 8 * 0.012 + pd * lgd * 9.81 + CDbl(42) * 0.000001
    TierRW_42 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_43(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 9 * 0.012 + pd * lgd * 9.81 + CDbl(43) * 0.000001
    TierRW_43 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_44(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 10 * 0.012 + pd * lgd * 9.81 + CDbl(44) * 0.000001
    TierRW_44 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_45(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 11 * 0.012 + pd * lgd * 9.81 + CDbl(45) * 0.000001
    TierRW_45 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_46(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 12 * 0.012 + pd * lgd * 9.81 + CDbl(46) * 0.000001
    TierRW_46 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_47(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 13 * 0.012 + pd * lgd * 9.81 + CDbl(47) * 0.000001
    TierRW_47 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_48(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 14 * 0.012 + pd * lgd * 9.81 + CDbl(48) * 0.000001
    TierRW_48 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_49(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 15 * 0.012 + pd * lgd * 9.81 + CDbl(49) * 0.000001
    TierRW_49 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_50(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 16 * 0.012 + pd * lgd * 9.81 + CDbl(50) * 0.000001
    TierRW_50 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_51(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 0 * 0.012 + pd * lgd * 9.81 + CDbl(51) * 0.000001
    TierRW_51 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_52(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 1 * 0.012 + pd * lgd * 9.81 + CDbl(52) * 0.000001
    TierRW_52 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_53(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 2 * 0.012 + pd * lgd * 9.81 + CDbl(53) * 0.000001
    TierRW_53 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_54(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 3 * 0.012 + pd * lgd * 9.81 + CDbl(54) * 0.000001
    TierRW_54 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_55(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 4 * 0.012 + pd * lgd * 9.81 + CDbl(55) * 0.000001
    TierRW_55 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_56(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 5 * 0.012 + pd * lgd * 9.81 + CDbl(56) * 0.000001
    TierRW_56 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_57(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 6 * 0.012 + pd * lgd * 9.81 + CDbl(57) * 0.000001
    TierRW_57 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_58(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 7 * 0.012 + pd * lgd * 9.81 + CDbl(58) * 0.000001
    TierRW_58 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_59(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 8 * 0.012 + pd * lgd * 9.81 + CDbl(59) * 0.000001
    TierRW_59 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_60(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 9 * 0.012 + pd * lgd * 9.81 + CDbl(60) * 0.000001
    TierRW_60 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_61(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 10 * 0.012 + pd * lgd * 9.81 + CDbl(61) * 0.000001
    TierRW_61 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_62(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 11 * 0.012 + pd * lgd * 9.81 + CDbl(62) * 0.000001
    TierRW_62 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_63(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 12 * 0.012 + pd * lgd * 9.81 + CDbl(63) * 0.000001
    TierRW_63 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_64(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 13 * 0.012 + pd * lgd * 9.81 + CDbl(64) * 0.000001
    TierRW_64 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_65(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 14 * 0.012 + pd * lgd * 9.81 + CDbl(65) * 0.000001
    TierRW_65 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_66(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 15 * 0.012 + pd * lgd * 9.81 + CDbl(66) * 0.000001
    TierRW_66 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_67(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 16 * 0.012 + pd * lgd * 9.81 + CDbl(67) * 0.000001
    TierRW_67 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_68(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 0 * 0.012 + pd * lgd * 9.81 + CDbl(68) * 0.000001
    TierRW_68 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_69(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 1 * 0.012 + pd * lgd * 9.81 + CDbl(69) * 0.000001
    TierRW_69 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_70(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 2 * 0.012 + pd * lgd * 9.81 + CDbl(70) * 0.000001
    TierRW_70 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_71(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 3 * 0.012 + pd * lgd * 9.81 + CDbl(71) * 0.000001
    TierRW_71 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Private Function TierRW_72(ByVal pd As Double, ByVal lgd As Double) As Double
    Dim rw As Double
    rw = 4 * 0.012 + pd * lgd * 9.81 + CDbl(72) * 0.000001
    TierRW_72 = FinCore_Common.RoundMonetaryHalfEven(rw, "JPY")
End Function

Public Sub RefreshRiskWeightedStub()
    Dim ws As Worksheet
    Dim r As Long
    Dim ead As Double
    Dim lgd As Double
    Dim pd As Double
    Dim rwa As Double
    Dim mig As Long
    Set ws = Worksheets("RiskWeightedAssetsWork")
    For r = 180 To 980
        pd = ws.Cells(r, 6).Value2
        lgd = ws.Cells(r, 8).Value2
        ead = ws.Cells(r, 10).Value2
        mig = ws.Cells(r, 12).Value2
        Select Case mig
            Case 1
                rwa = ead * TierRW_3(pd, lgd)
            Case 2, 3
                rwa = ead * TierRW_11(pd, lgd) * VAT_GATE
            Case 4 To 8
                rwa = ead * TierRW_27(pd, lgd) * FX_SHOCK_HI
            Case Else
                rwa = ead * TierRW_41(pd, lgd) * FX_SHOCK_LO
        End Select
        ws.Cells(r, 28).Value2 = rwa
    Next r
End Sub

Public Sub MonteCarloSpreadStub(ByVal paths As Long)
    Dim i As Long
    Dim acc As Double
    Dim bump As Double
    Dim idx As Long
    acc = 0
    For i = 1 To paths
        bump = Sin(CDbl(i) / 17#) * Cos(CDbl(i) / 23#)
        idx = ((i - 1) Mod 72) + 1
        Select Case idx
            Case 1
                acc = acc + bump * TierRW_1(0.021 + bump * 0.001, 0.35)
            Case 2
                acc = acc + bump * TierRW_2(0.021 + bump * 0.001, 0.35)
            Case 3
                acc = acc + bump * TierRW_3(0.021 + bump * 0.001, 0.35)
            Case 4
                acc = acc + bump * TierRW_4(0.021 + bump * 0.001, 0.35)
            Case 5
                acc = acc + bump * TierRW_5(0.021 + bump * 0.001, 0.35)
            Case 6
                acc = acc + bump * TierRW_6(0.021 + bump * 0.001, 0.35)
            Case 7
                acc = acc + bump * TierRW_7(0.021 + bump * 0.001, 0.35)
            Case 8
                acc = acc + bump * TierRW_8(0.021 + bump * 0.001, 0.35)
            Case 9
                acc = acc + bump * TierRW_9(0.021 + bump * 0.001, 0.35)
            Case 10
                acc = acc + bump * TierRW_10(0.021 + bump * 0.001, 0.35)
            Case 11
                acc = acc + bump * TierRW_11(0.021 + bump * 0.001, 0.35)
            Case 12
                acc = acc + bump * TierRW_12(0.021 + bump * 0.001, 0.35)
            Case 13
                acc = acc + bump * TierRW_13(0.021 + bump * 0.001, 0.35)
            Case 14
                acc = acc + bump * TierRW_14(0.021 + bump * 0.001, 0.35)
            Case 15
                acc = acc + bump * TierRW_15(0.021 + bump * 0.001, 0.35)
            Case 16
                acc = acc + bump * TierRW_16(0.021 + bump * 0.001, 0.35)
            Case 17
                acc = acc + bump * TierRW_17(0.021 + bump * 0.001, 0.35)
            Case 18
                acc = acc + bump * TierRW_18(0.021 + bump * 0.001, 0.35)
            Case Else
                acc = acc + bump * TierRW_61(0.019 + bump * 0.002, 0.41)
        End Select
    Next i
    Worksheets("RiskWeightedAssetsWork").Range("AA5").Value2 = acc
End Sub

Public Sub StressNestedBuckets(ByVal depth As Long)
    Dim i As Long
    Dim j As Long
    Dim acc As Double
    acc = FinLedger_GL.AccumulateTrialBalanceColumn(58)
    For i = 1 To depth
        For j = 1 To depth
            acc = acc + TierRW_19(0.014 + CDbl(i + j) * 0.00001, 0.37)
        Next j
    Next i
    Worksheets("RiskWeightedAssetsWork").Range("AA7").Value2 = acc
End Sub
