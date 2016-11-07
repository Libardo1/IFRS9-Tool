Imports DataManager

Public Class ReceivablePortFolio
    Inherits Portfolio

    Private Receivables As List(Of Receivable)

    Public Sub New(ReportDate As Date)
        MyBase.TheReportDate = ReportDate
    End Sub

    ' Retrieve Data
    Public Sub RetrieveData(DB As AccessIO)
        DB.OpenConnection()
        Receivables = New List(Of Receivable)

        'Retrieve all receivables data
        Dim Query As String = "Select * 
                               FROM [Dat_Receivable]"

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        Dim Receivable As Receivable
        While (DB.Reader.Read())
            Receivable = New Receivable
            Receivable.Retrieve(DB)
            Receivables.Add(Receivable)
        End While

        DB.CloseReader()
        DB.CloseConnection()
    End Sub

    Public ReadOnly Property GetAllReceivables() As List(Of Receivable)
        Get
            Return Receivables
        End Get
    End Property

End Class
