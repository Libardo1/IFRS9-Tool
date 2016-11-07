Public Class LGD
    Private ID_Instrument As Long
    Private Ref_Date As Date
    Private LGD_Value As Double

    Public Sub New(Instrument_ID As Long, Ref_Date As Date)
        Me.ID_Instrument = ID_Instrument
        Me.Ref_Date = Ref_Date
    End Sub

    Public ReadOnly Property GetLGD() As Double
        Get
            Return LGD_Value
        End Get
    End Property
End Class