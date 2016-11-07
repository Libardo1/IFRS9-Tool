Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.Method

Public Class FixedIncomeTransaction
    Inherits Instrument

    Private ID_Perso As String 'ISIN number for bonds
    Private FaceValue As Decimal
    Private Coupon As Double

    Private EIR As Double
    Private Instrument_Quantity As Long
    Private Unit_Price As Decimal
    Private Transaction_Value As Decimal
    Private Interest_Term As String
    Private CouponFreq As CouponFrequency
    Private CFSchemeType As CashFlowScheme.CFSchemeType
    Private MeasurementBase As Base

    Private CashFlowScheme As CashFlowScheme
    Private PaymentFrequency As Frequency

    Private eIDN As DayCountConvention.eIDN = 7
    Private GrossCarryingValue As Double
    Private FICalculator As FICalculator

    Private CalculationMethod As CalculationMethod
    Private Timing As ProvisioningMoments
    Private PDMethod As PDMethod
    Private EADMethod As EADMethod
    Private LGDMethod As LGDMethod
    Private StagingMethod As StagingMethod
    Private ThresholdX As Double

    'Temp for Anadolu
    Public CollateralValue As Decimal

    'Can't chose the stage so set the first one here (to get eaither 1year or LT ECL)
    'Private TemporaryStage As Stage.IFRSStage = Stage.IFRSStage.IncreasedRisk


    Public Enum Base
        HistoricalCost = 1
        AmortisedCost = 2
        FairValue = 3
    End Enum

    Public Enum Frequency
        Monthly = 0
        Yearly = 1
    End Enum

    Public Sub New(CalculationMethod As CalculationMethod, Timing As ProvisioningMoments, PDMethod As PDMethod, EADMethod As EADMethod, LGDMethod As LGDMethod, MasterRatingMethod As MasterRatingMethod,
                   StagingMethod As StagingMethod, ThresholdX As Double)
        Me.CalculationMethod = CalculationMethod
        Me.Timing = Timing
        Me.PDMethod = PDMethod
        Me.EADMethod = EADMethod
        Me.LGDMethod = LGDMethod
        Me.StagingMethod = StagingMethod
        Me.ThresholdX = ThresholdX
    End Sub

    'retrieve data from DB
    Public Function Retrieve(ReportDate As Date, DB As AccessIO) As FixedIncomeTransaction
        MyBase.TheReportDate = ReportDate

        MyBase.TheID = DB.Reader("ID")
        ID_Perso = DB.Reader("LOAN_REF_NO")
        MyBase.TheID_Counterparty = DB.Reader("ID_COUNTERPARTY")
        MyBase.TheName_Counterparty = DB.Reader("COUNTERPARTY_NAME")
        MyBase.TheCurrency = DB.Reader("CURRENCY")
        TheTransactionDate = DB.Reader("TRANSACTION_DATE")
        MyBase.TheMaturityDate = DB.Reader("MATURITY_DATE")
        FaceValue = DB.Reader("FaceValue")
        If Not DBNull.Value.Equals(DB.Reader("Transaction_Value")) Then
            Transaction_Value = DB.Reader("Transaction_Value")
        Else
            Transaction_Value = FaceValue
        End If

        Interest_Term = DB.Reader("Interest_Term")
        CouponFreq = SetCouponFrequency(Interest_Term)

        'NEED TO ADD INTEREST_TYPE (FIX, FLO...)

        Coupon = DB.Reader("COUPON") / 100

        'for whatever reason these two don't work... Find out later
        CFSchemeType = 1 'DB.Reader("CFSchemeType")
        MeasurementBase = 2 'DB.Reader("MeasurementBase")


        FICalculator = New FICalculator

        Return Me
    End Function

    Public Sub CalculateCashFlowScheme()
        CashFlowScheme = FICalculator.CalculateCashFlowScheme(TheReportDate, TheMaturityDate, Coupon, FaceValue, CFSchemeType, GetCouponFrequency, eIDN)
    End Sub

    'PUT IT IN INSTRUMENT CLASS ?
    Public Sub RetrieveCFScheme(DB As AccessIO)
        'fill CashFlowScheme with all the cashflows
        CashFlowScheme = New CashFlowScheme()
        CashFlowScheme.RetrieveCFs(MyBase.TheID, DB)

    End Sub

    Public Sub CalculateGrossCarryingValue()
        Select Case MeasurementBase
            Case Base.HistoricalCost
                GrossCarryingValue = Transaction_Value
            Case Base.AmortisedCost
                GrossCarryingValue = CalculateGrossCarryingValueUsingCFs(MyBase.TheReportDate, CashFlowScheme)
            Case Else
                GrossCarryingValue = Decimal.MinValue
        End Select
    End Sub

    Private Function CalculateGrossCarryingValueUsingCFs(ReferenceDate As Date, CFScheme As CashFlowScheme) As Double
        GrossCarryingValue = FICalculator.CalculateNPV(ReferenceDate, CFScheme, EIR, eIDN)

        Return GrossCarryingValue
    End Function

    Public Sub CalculateEIR()
        EIR = FICalculator.CalculateEIR(TheTransactionDate, TheMaturityDate, Transaction_Value, Coupon, FaceValue, eIDN)
    End Sub

    'calculate provision
    Public Sub CalculateProvision(DB As AccessIO)
        Dim ProvisionCalc As ProvisionCalculatorFI = New ProvisionCalculatorFI(TheReportDate)

        'determine instrument stage
        TheStage = New Stage
        If CalculationMethod = CalculationMethod.AllOneYear Then
            TheInstrumentStage = Stage.IFRSStage.Standard

        ElseIf CalculationMethod = CalculationMethod.AllLifetime Then
            TheInstrumentStage = Stage.IFRSStage.IncreasedRisk

        ElseIf CalculationMethod = CalculationMethod.StageDependent Then
            'retrieve current stage
            TheInstrumentStage = GetStage(TheOriginationDate, TheReportDate)

        Else Throw New Exception("ERROR: Instrument stage unknown")
            Stop
        End If

        'launch either 1y or LT provision based on CurrentStage
        ProvisionCalc.CalculateProvision(Me, DB, TheInstrumentStage)

        TheProvision = ProvisionCalc.GetProvision
    End Sub

    ''calculate provision DELETE AFTER
    'Public Sub CalculateProvision(DB As AccessIO)
    '    Dim ProvisionCalc As ProvisionCalculatorFI = New ProvisionCalculatorFI(TheReportDate)

    '    'retrieve current stage of the instrument
    '    TheInstrumentStage = GetStage(TheOriginationDate, TheReportDate)

    '    'launch either 1y or LT provision based on CurrentStage
    '    ProvisionCalc.CalculateProvision(Me, DB, TheInstrumentStage)

    '    TheProvision = ProvisionCalc.GetProvision
    'End Sub

    'NOT IMPLEMENTED
    Public ReadOnly Property CalcEIR(ReferenceDate As Date) As Double
        'calc EIR using instru info
        Get
            Return FICalculator.CalculateEIR(TheTransactionDate, TheMaturityDate, TheAmount / GetQuantity, Coupon, FaceValue, eIDN)
        End Get
    End Property

    Public Sub SaveInstrumentResultAccess(DB As AccessIO, filepath As String, ID_Report As Integer)
        DB = New AccessIO(filepath)
        DB.OpenConnection()

        Dim IDCalc As Integer = 1 'create Method containing all the possible PD/EAD/LGD combinations ?

        Dim query As String
        Dim InstrumentType As String = "Res_Provision yallah"
        Select Case InstrumentType
            Case "Security"
                query = "INSERT INTO Res_Provisions_Security([ReportDate], [No_Instrument], [Counterparty], [Exposure], [Maturity], [ID_Calculation], [ECL_TYPE], [ECL_VALUE])
                                values(#" & TheReportDate & "#, '" & ID_Perso & "', '" & TheName_Counterparty & "', '" & Convert.ToDecimal(FaceValue) & "', #" & TheMaturityDate & "#, " & IDCalc & ", '" & OneYearOrLifeTime() & "', '" & Convert.ToDecimal(TheProvision) & "')"

            Case "Loan"
                query = "INSERT INTO Res_Provisions_Loan([ReportDate], [No_Instrument], [Counterparty], [Exposure], [Maturity], [ID_Calculation], [ECL_TYPE], [ECL_VALUE])
                                values(#" & TheReportDate & "#, '" & ID_Perso & "', '" & TheName_Counterparty & "', '" & Convert.ToDecimal(FaceValue) & "', #" & TheMaturityDate & "#, " & IDCalc & ", '" & OneYearOrLifeTime() & "', '" & Convert.ToDecimal(TheProvision) & "')"

            Case "Account Receivable"
                '...

            Case Else
                query = "INSERT INTO Res_Provisions([ReportDate], [No_Instrument], [Counterparty], [Exposure], [Maturity], [ID_Calculation], [ECL_TYPE], [ECL_VALUE])
                                values(#" & TheReportDate & "#, '" & ID_Perso & "', '" & TheName_Counterparty & "', '" & Convert.ToDecimal(FaceValue) & "', #" & TheMaturityDate & "#, " & IDCalc & ", '" & OneYearOrLifeTime() & "', '" & Convert.ToDecimal(TheProvision) & "')"
        End Select

        DB.CreateCommand(query)
        DB.ExecuteNonQuery()

        DB.CloseConnection()
    End Sub

    Private Function OneYearOrLifeTime() As String
        Select Case TheInstrumentStage
            Case Stage.IFRSStage.Standard
                Return "1Year_ECL"
            Case Stage.IFRSStage.IncreasedRisk
                Return "Lifetime_ECL"
            Case Stage.IFRSStage.Impaired
                Return "Lifetime_ECL"
            Case Else
                Return "ERROR"
        End Select
    End Function


    Public Sub SaveInstrumentResultExcel()

    End Sub

    'Gets functions
    Public ReadOnly Property GetGrossCarryingValue()
        Get
            Return GrossCarryingValue
        End Get
    End Property

    Public ReadOnly Property GetNetCarryingValue() As Decimal
        Get
            GetNetCarryingValue = GrossCarryingValue - TheProvision
        End Get
    End Property

    Public ReadOnly Property GetEIR() As Double
        Get
            Return EIR
        End Get
    End Property

    Public ReadOnly Property GetCounterParty(DB As AccessIO) As Counterparty
        Get
            Dim Counterparty As New Counterparty(MyBase.TheID_Counterparty)
            GetCounterParty = Counterparty.Retrieve(MyBase.TheID_Counterparty, DB)
        End Get
    End Property

    Public ReadOnly Property GetID_Perso() As String
        Get
            GetID_Perso = ID_Perso
        End Get
    End Property

    Public ReadOnly Property GetFaceValue() As Decimal
        Get
            GetFaceValue = FaceValue
        End Get
    End Property

    Public ReadOnly Property GetCoupon() As Double
        Get
            GetCoupon = Coupon
        End Get
    End Property

    Public ReadOnly Property GetCashFlowScheme() As CashFlowScheme
        Get
            Return CashFlowScheme
        End Get
    End Property

    Public ReadOnly Property GetUnitPrice() As Double
        Get
            Return Unit_Price
        End Get
    End Property

    Public ReadOnly Property GetQuantity() As Long
        Get
            Return Instrument_Quantity
        End Get
    End Property

    Public ReadOnly Property GetTransactionValue() As Decimal
        Get
            Return Transaction_Value
        End Get
    End Property

    'put somewhere else
    Private Function SetCouponFrequency(Interest_Term As String) As CouponFrequency
        Select Case Interest_Term
            Case "12Months"
                CouponFreq = CouponFrequency.Annual
            Case "6Months"
                CouponFreq = CouponFrequency.SemiAnnual
            Case Else
                'Throw New Exception("Error in Coupon Frequency")
                CouponFreq = CouponFrequency.Annual
        End Select

        Return CouponFreq
    End Function

    Public ReadOnly Property GetCouponFrequency() As CouponFrequency
        Get
            Return CouponFreq
        End Get
    End Property

    Public ReadOnly Property GetMeasurementBase() As Base
        Get
            GetMeasurementBase = MeasurementBase
        End Get
    End Property

    Public ReadOnly Property GetPDMethod() As PDMethod
        Get
            GetPDMethod = PDMethod
        End Get
    End Property

    Public ReadOnly Property GetEADMethod() As EADMethod
        Get
            GetEADMethod = EADMethod
        End Get
    End Property

    Public ReadOnly Property GetLGDMethod() As LGDMethod
        Get
            GetLGDMethod = LGDMethod
        End Get
    End Property

    Public ReadOnly Property GetTiming() As ProvisioningMoments
        Get
            GetTiming = Timing
        End Get
    End Property

End Class

