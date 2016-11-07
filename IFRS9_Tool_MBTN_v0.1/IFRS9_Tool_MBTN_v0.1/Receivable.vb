Imports DataManager

Public Class Receivable
    Inherits Instrument

    'retrieve data from DB
    Public Function Retrieve(DB As AccessIO) As Receivable
        MyBase.TheID = DB.Reader("ID")
        MyBase.TheID_Counterparty = DB.Reader("ID_Counterparty")
        MyBase.TheAmount = DB.Reader("Amount")
        MyBase.TheCurrency = DB.Reader("Currency")
        MyBase.TheOriginationDate = DB.Reader("Creation_Date")
        MyBase.TheMaturityDate = DB.Reader("Due_Date")

        Return Me
    End Function

    'return number of days past due 
    Public ReadOnly Property GetNumDaysPastDue(ReferenceDate As Date) As Long
        Get
            GetNumDaysPastDue = CalcNumDaysPastDue(ReferenceDate)
        End Get
    End Property

    'calc numberod days past due
    Private Function CalcNumDaysPastDue(ReferenceDate As Date) As Long
        Dim NumDaysPastDue = DateDiff(DateInterval.Day, MyBase.TheMaturityDate, ReferenceDate)
        If NumDaysPastDue <= 0 Then
            NumDaysPastDue = 0
        End If

        Return NumDaysPastDue
    End Function

    Public ReadOnly Property GetCounterParty(DB As AccessIO) As Counterparty
        Get
            Dim Counterparty As New Counterparty(MyBase.TheID_Counterparty)
            GetCounterParty = Counterparty.Retrieve(MyBase.TheID_Counterparty, DB)
        End Get
    End Property

End Class
