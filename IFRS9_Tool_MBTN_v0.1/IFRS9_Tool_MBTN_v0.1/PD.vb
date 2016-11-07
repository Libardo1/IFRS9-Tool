Public Class PD
    Private PD_Value As Double
    Private ID_Counterparty As Long

    Public Sub New()

    End Sub

    Property ThePDValue() As Double
        Get
            Return PD_Value
        End Get

        Set(ByVal Value As Double)
            PD_Value = Value
        End Set
    End Property

End Class

