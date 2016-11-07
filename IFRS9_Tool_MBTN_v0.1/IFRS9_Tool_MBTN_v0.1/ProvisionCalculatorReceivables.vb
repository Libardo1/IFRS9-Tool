Imports DataManager

Public Class ProvisionCalculatorReceivables
    'MAXIME: create Base Class ProvisionCalculator having all the common parameters and functions?
    Private ReportDate As Date
    Private Receivables As ReceivablePortFolio
    Private Method As ProvisioningMethod
    Private ProvisionRuleSets As List(Of ProvisionRuleSetReceivables)

    Public Sub New(ReportDate As Date, Receivables As ReceivablePortFolio)
        Me.ReportDate = ReportDate
        Me.Receivables = Receivables
    End Sub

    Public Enum ProvisioningMethod
        SimpleLossProvisioning = 0
        ProvisioningByRatingSnP = 1
        ProvisioningByRatingZanders = 2
    End Enum

    Public Sub RetrieveRules(DB As AccessIO)
        'Create set of rules
        ProvisionRuleSets = New List(Of ProvisionRuleSetReceivables)
        Dim ProvisionRuleSet As ProvisionRuleSetReceivables

        'loop through all ProvisioningMethods
        Dim values = [Enum].GetValues(GetType(ProvisioningMethod))
        For Each value In values
            ProvisionRuleSet = New ProvisionRuleSetReceivables(Method, Receivables)
            ProvisionRuleSet.RetrieveProvSetDes(DB) 'retrieve description
            ProvisionRuleSet.RetrieveProvRules(DB) 'retrieve rules for this set

            'Add it to the list of sets
            ProvisionRuleSets.Add(ProvisionRuleSet)
        Next
    End Sub

    Public Sub CalcProvision(DB As AccessIO)
        Dim Provision As Provision

        'Loop through all the receivables
        For Each receivable In Receivables.GetAllReceivables
            'loop through all ProvisioningMethods
            Dim values = [Enum].GetValues(GetType(ProvisioningMethod))
            For Each value In values

                ' Retrieve receivables info
                Dim ID_Receivable As Long = receivable.TheID
                Dim amount As Double = receivable.TheAmount
                Dim DaysPastDue As Integer = receivable.GetNumDaysPastDue(ReportDate)

                'Retrieve rating base on provisioning model
                Dim Rating As String
                If Method = ProvisioningMethod.SimpleLossProvisioning Then
                    Rating = "None"
                ElseIf Method = ProvisioningMethod.ProvisioningByRatingSnP Then
                    Rating = receivable.GetCounterParty(DB).GetRatingSnP(ReportDate, DB)
                ElseIf Method = ProvisioningMethod.ProvisioningByRatingZanders Then
                    Rating = receivable.GetCounterParty(DB).GetRatingZanders(ReportDate, DB)
                Else
                    Throw New Exception("Provisioning method not correct. Verify the provisioning method.")
                End If

                Dim LossRate As Double = ProvisionRuleSets(value).GetLossRate(DaysPastDue, Rating, DB)
                Dim ProvisionAmount As Double = amount * LossRate

                'Create new provision and save in DB
                Provision = New Provision(ID_Receivable, ReportDate, Method, ProvisionAmount)
                Provision.SaveProvisions(DB)
            Next
        Next
    End Sub

    Public ReadOnly Property GetRuleSets() As List(Of ProvisionRuleSetReceivables)
        Get
            Return ProvisionRuleSets
        End Get
    End Property

    Public Property TheMethod() As ProvisioningMethod
        Get
            TheMethod = Method
        End Get
        Set(ByVal value As ProvisioningMethod)
            Method = value
        End Set
    End Property

End Class
