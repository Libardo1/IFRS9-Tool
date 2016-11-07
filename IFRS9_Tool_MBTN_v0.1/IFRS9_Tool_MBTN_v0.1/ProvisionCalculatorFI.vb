Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.Method

Public Class ProvisionCalculatorFI
    ' create Base Class ProvisionCalculator having all the common parameters and functions?
    Private ReportDate As Date
    Private TSet As List(Of Date)
    Private ECL As ECL

    Public Sub New(ReportDate As Date)
        Me.ReportDate = ReportDate
    End Sub

    Public Sub CalculateProvision(Instrument As FixedIncomeTransaction, DB As AccessIO, InstrumentStage As Stage.IFRSStage)
        ECL = New ECL(Instrument.TheID, Instrument.TheReportDate)
        'some observation have a maturity date lower than today --> set ECL to 0 and exit sub
        If Instrument.TheMaturityDate <= ReportDate Then
            ECL.TheECLValue = 0
            Exit Sub
        End If

        Dim TotalProvision As Decimal = 0
        Dim PDCalc As PDCalculator = New PDCalculator(ReportDate, Instrument.GetPDMethod, DB)
        Dim EADCalc As EADCalculator = New EADCalculator(ReportDate, Instrument.GetEADMethod)
        Dim LGDCalc As LGDCalculator = New LGDCalculator(ReportDate, Instrument.GetLGDMethod)
        Dim eIDN As DayCountConvention.eIDN = DayCountConvention.eIDN.eACT365 ' Daycount set to eACT365 by default
        Dim EIR As Double = Instrument.GetEIR()

        'only Anadolubank rating
        Dim Rating As String = Instrument.GetRatingBySystem(RatingMasterData.RatingSystem.Fitch)
        Dim InstrumentTiming As ProvisioningMoments = Instrument.GetTiming

        DetermineTset(Instrument.TheMaturityDate, Instrument.GetTiming)
        Dim N As Long = TSet.Count

        'launch calculation based on current stage
        Select Case InstrumentStage
            Case Stage.IFRSStage.Standard
                ECL.TheECLType = ECL.ECLType.OneYearECL
                TotalProvision = CalculateOneYearProvision(Instrument, PDCalc, EADCalc, LGDCalc, Rating, InstrumentTiming, eIDN, EIR)

            Case Stage.IFRSStage.IncreasedRisk
                ECL.TheECLType = ECL.ECLType.LifeTimeECL
                TotalProvision = CalculateLifetimeProvision(Instrument, PDCalc, EADCalc, LGDCalc, Rating, InstrumentTiming, eIDN, EIR, N)

            Case Stage.IFRSStage.Impaired
                'NOT IMPLEMENTED
                ECL.TheECLType = ECL.ECLType.LifeTimeECL
                TotalProvision = 0

            Case Else
                Throw New Exception("ERROR: stage unknown")
                Stop
        End Select

        ECL.TheECLValue = TotalProvision
    End Sub

    Private Function CalculateOneYearProvision(Instrument As FixedIncomeTransaction, PDCalc As PDCalculator, EADCalc As EADCalculator, LGDCalc As LGDCalculator,
                                               Rating As String, Timing As ProvisioningMoments, eIDN As DayCountConvention.eIDN, EIR As Double) As Decimal
        Dim provision As Decimal

        Dim PD1y As Double = PDCalc.CalcConditionalDefaultProbability(TSet, TSet(0), TSet(1), Rating, Timing, eIDN)
        Dim EAD1y As Double = EADCalc.CalculatePeriodEAD(TSet(1), Instrument, eIDN)
        Dim LGD1y As Double = LGDCalc.CalculatePeriodLGD(TSet(1), Instrument)
        Dim DF1y As Double = FICalculator.CalculateDF(ReportDate, EIR, TSet(1), eIDN)

        provision = PD1y * EAD1y * LGD1y * DF1y

        Return provision
    End Function

    Private Function CalculateLifetimeProvision(Instrument As FixedIncomeTransaction, PDCalc As PDCalculator, EADCalc As EADCalculator, LGDCalc As LGDCalculator,
                                               Rating As String, Timing As ProvisioningMoments, eIDN As DayCountConvention.eIDN, EIR As Double, N As Long) As Decimal
        Dim provision As Decimal

        'components arrays
        Dim PD(N - 2), EAD(N - 2), LGD(N - 2), Prov(N - 2), DF(N - 2) As Double
        For i = 0 To N - 2
            PD(i) = PDCalc.CalcConditionalDefaultProbability(TSet, TSet(i), TSet(i + 1), Rating, Instrument.GetTiming, eIDN)
            EAD(i) = EADCalc.CalculatePeriodEAD(TSet(i + 1), Instrument, eIDN)
            LGD(i) = LGDCalc.CalculatePeriodLGD(TSet(i + 1), Instrument)
            DF(i) = FICalculator.CalculateDF(ReportDate, EIR, TSet(i + 1), eIDN)

            Prov(i) = (PD(i) * EAD(i) * LGD(i) * DF(i))
            provision += Prov(i)
        Next

        Return provision
    End Function

    Private Sub DetermineTset(EndDate As Date, Timing As ProvisioningMoments)
        Dim RelevantDate As Date = ReportDate
        TSet = New List(Of Date)

        Select Case Timing
            Case ProvisioningMoments.EndOfYear
                'first date (ReportDate)
                TSet.Add(RelevantDate)
                RelevantDate = New Date(Year(ReportDate), 12, 31) 'last day of current year -> currentyear/12/31
                'RelevantDate = New Date(Year(ReportDate) + 1, 1, 1) 'first day of next year - > 01/01/currentyear+1

                'first year date
                If TSet.Item(0) <> RelevantDate Then
                    TSet.Add(RelevantDate)
                End If

                'middle year dates
                RelevantDate = DateAdd(DateInterval.Year, 1, RelevantDate)
                Do While RelevantDate < EndDate
                    TSet.Add(RelevantDate)
                    RelevantDate = DateAdd(DateInterval.Year, 1, RelevantDate)
                Loop

                'last year date: added twice because principal and interest are separeted in the DB
                If TSet.Item(TSet.Count - 1) <> EndDate Then
                    TSet.Add(EndDate)
                    TSet.Add(EndDate)
                End If

            Case ProvisioningMoments.YearlyFromReferenceDate
                Do While RelevantDate < EndDate
                    TSet.Add(RelevantDate)
                    RelevantDate = DateAdd(DateInterval.Year, 1, RelevantDate)
                Loop

                If TSet.Item(TSet.Count - 1) <> EndDate Then
                    TSet.Add(EndDate)
                End If

            Case Else
                Throw New Exception("Unknown timing model")
        End Select
    End Sub

    Public ReadOnly Property GetProvision() As Decimal
        Get
            Return ECL.TheECLValue
        End Get
    End Property

    'Public Sub CalcProvisionOLD()
    '    Dim Provision As Provision
    '    Dim ProvisionAmount As Double

    '    'loop through all the fixe income
    '    For i As Integer = 0 To FIPortfolio.GetAllFixedIncomes.Count - 1
    '        'loop through all ProvisioningMethods
    '        Dim values = [Enum].GetValues(GetType(ProvisioningMethod))
    '        For Each value In values

    '            'Need to retrieve anything before? The CFs?
    '            Dim ID As Long = FIPortfolio.GetAllFixedIncomes(i).TheID
    '            Dim Transaction_Date As Date = FIPortfolio.GetAllFixedIncomes(i).GetTransactionDate
    '            Dim CarryingValue As Double = FIPortfolio.GetAllFixedIncomes(i).GetGrossCarryingValue(ReportDate) ' the GrossCarryingValue can also be reach using ReportDate

    '            Dim PDt As Double
    '            Dim Rating_Counterparty As String = FIPortfolio.GetAllFixedIncomes(i).GetCounterParty(DB).GetRatingSnP(ReportDate, DB)
    '            If Method = ProvisioningMethod.FixedPD Then
    '                'retrieve PD based on rating and reportDate: should be OK
    '                PDConstant = PDCalculator.GetLifeTimePDConstant(Rating_Counterparty, ReportDate, DB)

    '                ProvisionAmount = PDConstant * FIPortfolio.GetAllFixedIncomes(i).GetGrossCarryingValue()

    '            ElseIf Method = ProvisioningMethod.MatrixPD Then
    '                '1)need to keep track of every NPVs separetely 


    '                '2)loop through all the CFs, reach the PD for each. PRB if it is not annual interest!?
    '                Dim NumYear As Integer = FIPortfolio.GetAllFixedIncomes(i).GetCashFlowScheme.GetCashFlows.Count
    '                For j As Integer = 1 To NumYear
    '                    PDt = PDCalculator.GetLifeTimePDTM(Rating_Counterparty, j)
    '                    PDs.Add(PDt)
    '                Next

    '            Else
    '                Throw New Exception("PD model not correct. Verify the PD model")
    '            End If

    '            'Create new provision and save in ACCESS
    '            Provision = New Provision(ID, ReportDate, Method, ProvisionAmount)
    '            Provision.SaveProvisions(DB)
    '        Next
    '    Next
    'End Sub

    'Public Property TheMethod() As ProvisioningMethod
    '    Get
    '        TheMethod = Method
    '    End Get
    '    Set(ByVal value As ProvisioningMethod)
    '        Method = value
    '    End Set
    'End Property

End Class



