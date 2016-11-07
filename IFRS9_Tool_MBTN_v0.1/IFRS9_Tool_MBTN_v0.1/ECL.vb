Public Class ECL
    Private ID_Instrument As Long
    Private Ref_Date As Date
    Private ECL_Type As ECLType
    Private ECL_Value As Double

    Public Enum ECLType
        OneYearECL = 0
        LifeTimeECL = 1
    End Enum

    'used at the portfolio level
    Public Sub New(Ref_Date As Date)
        Me.Ref_Date = Ref_Date

    End Sub

    'used at the instrument level
    Public Sub New(ID_Instrument As Long, Ref_Date As Date)
        Me.ID_Instrument = ID_Instrument
        Me.Ref_Date = Ref_Date

    End Sub

    Property TheECLType() As ECLType
        Get
            Return ECL_Type
        End Get

        Set(ByVal Value As ECLType)
            ECL_Type = Value
        End Set
    End Property

    Property TheECLValue() As Double
        Get
            Return ECL_Value
        End Get

        Set(ByVal Value As Double)
            ECL_Value = Value
        End Set
    End Property


End Class
