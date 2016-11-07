Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.Method

Public Class EADCalculator
    Private ReportDate As Date
    Private Model As PDMethod

    Public Sub New(ReportDate As Date, Method As Method.LGDMethod)
        Me.ReportDate = ReportDate
        Me.Model = Method
    End Sub

    Public Function CalculatePeriodEAD(EAD_Date As Date, Instrument As FixedIncomeTransaction, eIDN As DayCountConvention.eIDN) As Double
        Dim PeriodEAD As Double

        Select Case Model
            Case EADMethod.ACEndOfPeriod
                PeriodEAD = CalcACEndOfPeriodEAD(EAD_Date, Instrument, eIDN)

            Case EADMethod.ACEndOfPeriod_PlusCFO
                PeriodEAD = CalcACEndOfPeriod_PlusCFOEAD(EAD_Date, Instrument, eIDN)

            Case EADMethod.ACMidPeriod
                PeriodEAD = CalcACMidPeriodEAD(EAD_Date, Instrument, eIDN)

            Case EADMethod.ACAvgPeriod
                PeriodEAD = CalcACAvgPeriodEAD()

            Case EADMethod.FaceValueConstant
                PeriodEAD = CalcFaceValueConstantEAD(Instrument)

            Case EADMethod.BookValueConstant
                PeriodEAD = CalcBookValueConstantEAD(Instrument)
        End Select

        Return PeriodEAD
    End Function

    'CF at date of default is paid
    Private Function CalcACEndOfPeriodEAD(EAD_Date As Date, Instrument As FixedIncomeTransaction, eIDN As DayCountConvention.eIDN) As Double
        Dim CFScheme As CashFlowScheme = Instrument.GetCashFlowScheme
        Dim EIR As Double = Instrument.GetEIR

        Dim EAD_Value As Double = FICalculator.CalculateNPV(EAD_Date, CFScheme, EIR, eIDN)

        Return EAD_Value
    End Function

    'CF at date of default not yet paid
    Private Function CalcACEndOfPeriod_PlusCFOEAD(EAD_Date As Date, Instrument As FixedIncomeTransaction, eIDN As DayCountConvention.eIDN) As Double
        Dim CFScheme As CashFlowScheme = Instrument.GetCashFlowScheme
        Dim EIR As Double = Instrument.GetEIR

        Dim EAD_Value As Double = FICalculator.CalculateNPV_Plus(EAD_Date, CFScheme, EIR, eIDN)

        Return EAD_Value
    End Function

    'loss occured in the "middle of the time interval" 
    Private Function CalcACMidPeriodEAD(EAD_Date As Date, Instrument As FixedIncomeTransaction, eIDN As DayCountConvention.eIDN) As Double
        Dim CFSchemeOriginal As CashFlowScheme = Instrument.GetCashFlowScheme
        Dim EIR As Double = Instrument.GetEIR

        'create new CFScheme with middle values
        Dim MiddleCFScheme As CashFlowScheme = CFSchemeOriginal
        Dim MiddleCFAmount As Double
        For Each cashflow In CFSchemeOriginal.GetCashFlows
            MiddleCFAmount = cashflow.GetCF_Amount / 2
            MiddleCFScheme.SetCashFlowsAmounts(MiddleCFAmount)
        Next

        'CF at date of default has been paid
        Dim EAD_Value As Double = FICalculator.CalculateNPV(EAD_Date, MiddleCFScheme, EIR, eIDN)

        Return EAD_Value
    End Function

    'average EAD ?
    Private Function CalcACAvgPeriodEAD() As Double

        Return 0
    End Function

    'facevalue
    Private Function CalcFaceValueConstantEAD(Instrument As FixedIncomeTransaction) As Double
        Dim FaceValue As Double = Instrument.GetFaceValue
        Dim Quantity As Long = Instrument.GetQuantity

        Return FaceValue * Quantity
    End Function

    'bookvalue
    Private Function CalcBookValueConstantEAD(Instrument As FixedIncomeTransaction) As Double
        Dim BookValue As Double = Instrument.GetTransactionValue

        Return BookValue
    End Function
End Class
