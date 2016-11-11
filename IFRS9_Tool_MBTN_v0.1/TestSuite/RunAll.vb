Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports IFRS9_Tool_MBTN_v0._1
Imports IFRS9_Tool_MBTN_v0._1.Method
Imports DataManager
Imports System
Imports System.IO

<TestClass()> Public Class RunAll

    'datapath
    Dim strDataPath As String = "|DataDirectory|\AnadolubankDatabase_v0.0.accdb"

    'calculation method (1year, LT, StageDependent)
    Dim CalculationMethod As CalculationMethod = CalculationMethod.AllLifetime

    Dim Report As New Report(2, #12/31/2015#, strDataPath)


    Dim Timing As ProvisioningMoments = ProvisioningMoments.YearlyFromReferenceDate
    Dim PDMethod As PDMethod = PDMethod.ConstantConditionalPD
    Dim EADMethod As EADMethod = EADMethod.ACEndOfPeriod_PlusCFO
    Dim LGDMethod As LGDMethod = LGDMethod.CollToExposureRatio
    Dim RatingMethod As MasterRatingMethod = MasterRatingMethod.Average

    Dim DownGrade As DowngradeNotches = DowngradeNotches.None 'not used

    'threshold: why using two? --> remove ThresholdX ?
    Dim StagingMethod As StagingMethod = StagingMethod.HigherThresholdX
    Dim ThresholdX As Double = 1.0

    <TestMethod()> Public Sub Receivables()
        RunCalculationsReceivables()
    End Sub

    <TestMethod()> Public Sub FixedIncome()
        RunCalculationsFixedIncome()
    End Sub

    <TestMethod()> Public Sub RunAll()
        RunCalculationsReceivables()
        RunCalculationsFixedIncome()
    End Sub

    Public Sub RunCalculationsReceivables()
        Report.CalcReceivablesProvisions()
    End Sub

    Public Sub RunCalculationsFixedIncome()
        '***Start user input***'
        'ask user for calculation method (COMMENT to skip it in debug mode, otherwise doesn't work)
        'Dim prompt As String = String.Empty
        'Dim title As String = String.Empty
        'Dim defaultResponse As String = String.Empty
        'Dim answer As Object

        'prompt = "Please chose a Calculation Method (0 = AllOneYear, 1 = AllLifetime, 2 = StageDependent)"
        'title = "Getting user input"
        'defaultResponse = "Calculation Method here"

        'answer = InputBox(prompt, title, defaultResponse)
        'CalculationMethod = answer
        '***End user input***'

        'set methods
        Report.SetMehods(CalculationMethod, Timing, PDMethod, EADMethod, LGDMethod, RatingMethod, StagingMethod, ThresholdX, DownGrade) 'ok

        Report.RetrieveDataFI() 'ok

        Report.CalculateCashFlowScheme() 'ok

        Report.RetrieveRating() 'ok

        Report.CalcFIGrossCarryingValue()

        Report.CalcFIProvisions()

        Report.CalcFINetCarryingValue()

        Report.SaveResults()

    End Sub
End Class