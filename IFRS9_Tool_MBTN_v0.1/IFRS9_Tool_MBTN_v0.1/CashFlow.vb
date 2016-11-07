Imports DataManager

Public Class CashFlow
    Private ID As Long
    Private ID_CFScheme As Long
    Private CashFlowDate As Date
    Private CF_Amount As Decimal
    Private OutstandingPrincipal As Double
    Private Type As CFType

    Public Enum CFType
        InitialCost = 1
        Interest = 2
        Principal = 3
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(CashFlowDate As Date, CF_Amount As Double, OutStandingPrincipal As Double, Type As CFType)
        Me.CashFlowDate = CashFlowDate
        Me.CF_Amount = CF_Amount
        Me.OutstandingPrincipal = OutStandingPrincipal
        Me.Type = Type
    End Sub

    'retrieve data from DB
    Public Function Retrieve(DB As AccessIO) As CashFlow
        ID = DB.Reader("ID")
        ID_CFScheme = DB.Reader("ID_CFScheme")
        CashFlowDate = DB.Reader("Cf_Date")
        CF_Amount = DB.Reader("CF_Amount")
        Type = DB.Reader("ID_CFType")
        OutstandingPrincipal = DB.Reader("OutstandingPrincipal")

        Return Me
    End Function

    Public ReadOnly Property GetID() As Long
        Get
            Return ID
        End Get
    End Property


    Public ReadOnly Property GetID_CFScheme() As Long
        Get
            Return ID_CFScheme
        End Get
    End Property


    Public ReadOnly Property GetCFDate() As Date
        Get
            Return CashFlowDate
        End Get
    End Property


    Public ReadOnly Property GetCF_Amount() As Double
        Get
            Return CF_Amount
        End Get
    End Property


    Public ReadOnly Property GetOutstandingPrincipal() As Decimal
        Get
            Return OutstandingPrincipal
        End Get
    End Property


    Public ReadOnly Property GetCashFlowType() As CFType
        Get
            Return Type
        End Get
    End Property

    Public WriteOnly Property SetCashFlowAmount() As Double
        Set(ByVal value As Double)
            CF_Amount = value
        End Set
    End Property

End Class

