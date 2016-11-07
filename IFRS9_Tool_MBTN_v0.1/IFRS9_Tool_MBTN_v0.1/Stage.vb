Public Class Stage

    Public Stage As IFRSStage
    Private Threshold As Double

    Public Enum IFRSStage
        Standard = 1
        IncreasedRisk = 2
        Impaired = 3
        'originatedOrPurchasedImpaired = 4
    End Enum

    Public ReadOnly Property GetStage(MasterScoreRefDate1 As Double, MasterScoreRefDate2 As Double) As IFRSStage
        Get

            Select Case Threshold
                Case 0
                    If MasterScoreRefDate2 <= (MasterScoreRefDate1 + Threshold) Then
                        Stage = IFRSStage.Standard
                    ElseIf MasterScoreRefDate2 > (MasterScoreRefDate1 + Threshold) Then
                        Stage = IFRSStage.IncreasedRisk
                    End If

                Case Else
                    If MasterScoreRefDate2 < (MasterScoreRefDate1 + Threshold) Then
                        Stage = IFRSStage.Standard
                    ElseIf MasterScoreRefDate2 >= (MasterScoreRefDate1 + Threshold) Then
                        Stage = IFRSStage.IncreasedRisk
                    End If
            End Select

            Return Stage
        End Get
    End Property

    Property TheThreshold() As Double
        Get
            Return Threshold
        End Get
        Set(ByVal Value As Double)
            Threshold = Value
        End Set
    End Property

    Property TheStage() As IFRSStage
        Get
            Return Stage
        End Get
        Set(ByVal Value As IFRSStage)
            Stage = Value
        End Set
    End Property

End Class
