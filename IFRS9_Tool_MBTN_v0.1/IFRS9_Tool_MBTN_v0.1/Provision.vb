Imports DataManager

''' <summary>
''' used with Receivables provisions. Adapt it to ECL 
''' </summary>
Public Class Provision
    Private ID As Long
    Private ID_Receivable As Long
    Private Ref_Date As Date
    Private ID_ProvisionModel As Long
    Private ProvisionAmount As Double

    Public Sub New(ID_Receivable As Long, Ref_Date As Date, ID_ProvisionModel As Long, ProvisionAmount As Double)
        Me.ID_Receivable = ID_Receivable
        Me.Ref_Date = Ref_Date
        Me.ID_ProvisionModel = ID_ProvisionModel
        Me.ProvisionAmount = ProvisionAmount
    End Sub

    Public Sub SaveProvisions(DB As AccessIO)
        DB.OpenConnection()

        ' Insert data into Provisions_Results ACCESS table
        Dim query As String = "INSERT INTO Res_Provision_Receivables(ID_Receivable, Ref_Date, ID_ProvisionModel, Provision)
                               values (" + ID_Receivable.ToString + ",#" + Ref_Date + "#," + ID_ProvisionModel.ToString + "," + ProvisionAmount.ToString + ")"

        DB.CreateCommand(query)
        DB.ExecuteNonQuery()

        DB.CloseConnection()
    End Sub
End Class
