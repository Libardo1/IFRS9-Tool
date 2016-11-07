Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.Method

Public Class LGDCalculator
    Private ReportDate As Date
    Private Model As PDMethod


    Public Sub New(ReportDate As Date, Method As Method.LGDMethod)
        Me.ReportDate = ReportDate
        Me.Model = Method
    End Sub

    Public Function CalculatePeriodLGD(LGD_Date As Date, Instrument As FixedIncomeTransaction) As Double
        Dim PeriodLGD As Double

        Select Case Model
            Case LGDMethod.Fixed
                PeriodLGD = CalcFixedLGD(LGD_Date, Instrument)
            Case LGDMethod.CollToExposureRatio
                PeriodLGD = 1 - (Instrument.CollateralValue / Instrument.GetFaceValue)
        End Select

        Return PeriodLGD
    End Function


    Private Function CalcFixedLGD(LGD_Date As Date, Instrument As FixedIncomeTransaction)
        Dim LGD As Double = 0.4

        Return LGD
    End Function

End Class
