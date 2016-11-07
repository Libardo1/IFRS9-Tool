Imports IFRS9_Tool_MBTN_v0._1.Method
Public Class FICalculator

    ' Approximation!
    Public Shared Function CalculateEIR(Transaction_Date As Date, Maturity_Date As Date, TransactionValue As Decimal, Coupon As Double, FaceValue As Decimal, eIDN As DayCountConvention.eIDN) As Double
        Dim EIR As Double
        Dim NumPeriod As Double = DayCountConvention.Yearfraction(Transaction_Date, Maturity_Date, eIDN)
        Dim I As Double = (Coupon * FaceValue)

        EIR = (I + ((FaceValue - TransactionValue) / NumPeriod)) / ((FaceValue + TransactionValue) / 2)

        If EIR < -1 Then
            EIR = 0
        End If
        Return EIR
    End Function

    'does not include ReferenceDate
    Public Shared Function CalculateNPV(ReferenceDate As Date, CashflowScheme As CashFlowScheme, EIR As Double, eIDN As DayCountConvention.eIDN) As Double
        Dim NPV As Double = 0
        Dim YearFrac As Double
        Dim CF_Date As Date

        For i As Integer = 0 To CashflowScheme.GetCashFlows().Count - 1
            CF_Date = CashflowScheme.GetCashFlows(i).GetCFDate

            If CF_Date <= ReferenceDate Then
                NPV += 0
            Else
                YearFrac = DayCountConvention.Yearfraction(ReferenceDate, CF_Date, eIDN)
                NPV += CashflowScheme.GetCashFlows(i).GetCF_Amount / ((1 + EIR) ^ (YearFrac))
            End If

        Next

        Return NPV
    End Function

    'include ReferenceDate
    Public Shared Function CalculateNPV_Plus(ReferenceDate As Date, CashflowScheme As CashFlowScheme, EIR As Double, eIDN As DayCountConvention.eIDN) As Double

        Dim NPV As Double = CalculateNPV(ReferenceDate, CashflowScheme, EIR, eIDN)

        Dim CashFlow As CashFlow = CashflowScheme.GetCashFlowByDate(ReferenceDate)

        If ReferenceDate = CashFlow.GetCFDate Then
            NPV += CashFlow.GetCF_Amount
        End If

        Return NPV
    End Function

    'present value CF
    Public Shared Function CalculatePVCashFlow(ReferenceDate As Date, CF_Date As Date, CF_Amount As Double, EIR As Double, eIDN As DayCountConvention.eIDN) As Double
        Dim PVCashFlow As Double
        Dim YearFrac As Double

        YearFrac = DayCountConvention.Yearfraction(ReferenceDate, CF_Date, eIDN)
        PVCashFlow = CF_Amount / ((1 + EIR) ^ (YearFrac))

        Return PVCashFlow
    End Function

    'discount factor
    Public Shared Function CalculateDF(ReportDate As Date, EIR As Double, CashFlowDate As Date, eIDN As DayCountConvention.eIDN)
        Dim DF As Double
        Dim YearFrac As Double = DayCountConvention.Yearfraction(ReportDate, CashFlowDate, eIDN)

        DF = 1 / ((1 + EIR) ^ (YearFrac))

        Return DF
    End Function

    'ugly
    Public Shared Function CalculateCashFlowScheme(ReportDate As Date, Maturity_Date As Date, Coupon As Double, Facevalue As Double, CFSchemeType As CashFlowScheme.CFSchemeType, CouponFrequency As CouponFrequency, eIDN As DayCountConvention.eIDN) As CashFlowScheme
        Dim CashFlowScheme As New CashFlowScheme(CFSchemeType)
        Dim Maturity = DayCountConvention.Yearfraction(ReportDate, Maturity_Date, eIDN)
        Dim CashFlowDate = ReportDate
        Dim CF_t As Double
        Dim OutstandingPrincipal As Double
        Dim CashFlowType As CashFlow.CFType
        Dim CashFlow As CashFlow

        Dim count As Integer = 0
        'Include CF_0 (transaction cost?)
        '?

        'CF_(t<T) = Coupon * Facevalue
        For i As Integer = 1 To Math.Truncate(Maturity)
            CashFlowDate = CashFlowDate.AddYears(1)
            CF_t = Coupon * Facevalue
            OutstandingPrincipal = Facevalue
            CashFlowType = CashFlow.CFType.Interest

            CashFlow = New CashFlow(CashFlowDate, CF_t, OutstandingPrincipal, CashFlowType)
            CashFlowScheme.AddCashFlows(CashFlow)

            count += 1
        Next

        'CF_(T) = FaceValue + (Coupon * Facevalue) (2 steps)
        'Step2: Facevalue
        CashFlowDate = Maturity_Date
        CF_t = Facevalue
        OutstandingPrincipal = 0
        CashFlowType = CashFlow.CFType.Principal

        CashFlow = New CashFlow(CashFlowDate, CF_t, OutstandingPrincipal, CashFlowType)
        CashFlowScheme.AddCashFlows(CashFlow)
        count += 1

        ''Step1: Coupon * Facevalue * FractionalPart
        'Dim FractionalPart As Double = SplitDecimal(Maturity)
        'CF_t = (Coupon * Facevalue) * FractionalPart
        'OutstandingPrincipal = 0
        'CashFlowType = CashFlow.CFType.Interest

        'CashFlow = New CashFlow(CashFlowDate, CF_t, OutstandingPrincipal, CashFlowType)
        'CashFlowScheme.AddCashFlows(CashFlow)
        'count += 1


        Return CashFlowScheme
    End Function

    'Public Shared Function CalculateSemiAnnualCashFlowScheme(Transaction_Date As Date, Maturity_Date As Date, Coupon As Double, Facevalue As Double, eIDN As DayCountConvention.eIDN) As CashFlowScheme
    '    Dim CashFlowScheme As New CashFlowScheme
    '    Dim Maturity = DayCountConvention.Yearfraction(Transaction_Date, Maturity_Date, eIDN)

    '    'CF_(t<T) = Coupon * Facevalue
    '    Dim CF_t As Double
    '    Dim CashFlow As CashFlow
    '    For i As Integer = 1 To (Maturity * 2) - 1
    '        CF_t = Coupon / 2 * Facevalue
    '        CashFlow = New CashFlow(CF_t)
    '        CashFlowScheme.AddCashFlows(CashFlow)
    '    Next

    '    'CF_(T) = FaceValue + (Coupon * Facevalue)
    '    Dim CF_last = Facevalue + (Coupon / 2 * Facevalue)
    '    CashFlow = New CashFlow(CF_last)
    '    CashFlowScheme.AddCashFlows(CashFlow)

    '    Return CashFlowScheme
    'End Function

    Public Shared Function SplitDecimal(ByVal number As Double)
        Dim wholePart = Math.Truncate(number)
        Dim fractionalPart = number - wholePart

        Return fractionalPart
    End Function

End Class