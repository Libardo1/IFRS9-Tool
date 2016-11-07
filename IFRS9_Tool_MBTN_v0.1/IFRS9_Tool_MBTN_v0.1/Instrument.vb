Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.RatingMasterData
Imports IFRS9_Tool_MBTN_v0._1.Method

Public Class Instrument
    Private ID As Long
    Private ID_Counterparty As Long
    Private Name_Counterparty As String
    Private Amount As Double
    Private Currency As String
    Private OriginationDate As Date
    Private TransactionDate As Date
    Private MaturityDate As Date
    Private ReportDate As Date

    Private Provision As Decimal
    Private ProvisionType As String

    Private HistoricalCost As Decimal 'necessary?

    Private RatingInfo As RatingInfo
    Private MasterRating As MasterRatingMethod

    Private InstrumentStage As Stage

    Private DB As AccessIO

    Public Sub RetrieveRatingInfo(DB As AccessIO, MasterRatingData As RatingMasterData, RatingMethod As MasterRatingMethod)
        RatingInfo = New RatingInfo(DB, MasterRatingData, RatingMethod)
        RatingInfo.Retrieve(TheName_Counterparty, TheReportDate)
    End Sub

    Public ReadOnly Property GetMasterRating() As String
        Get
            GetMasterRating = RatingInfo.GetOrCalculateMasterRating
        End Get
    End Property

    Public ReadOnly Property GetRatingBySystem(Method As RatingSystem) As String
        Get
            GetRatingBySystem = RatingInfo.GetRatingBySystem(Method).RatingName
        End Get
    End Property

    Public ReadOnly Property GetRemainingLifetime(PositionDate As Date) As Double
        ' to be defined
        Get
            Dim EndDate As Date = EndDate

            '  GetRemainingLifetime = FICalculator.CalculateRemainingLife(PositionDate, EndDate)
            Return 0.0
        End Get
    End Property

    Public ReadOnly Property GetRiskScore(ReferenceDate) As Long
        ' to be defined
        Get
            GetRiskScore = Long.MinValue
        End Get
    End Property

    Public ReadOnly Property GetStage(ReferenceDate1 As Date, ReferenceDate2 As Date) As Stage.IFRSStage
        Get
            'get master scores for ReferenceDate1 and ReferenceDate2
            Dim MasterScoreReferenceDate1 As Double = GetMasterScore(ReferenceDate1)
            Dim MasterScoreReferenceDate2 As Double = GetMasterScore(ReferenceDate2)

            'get the stage based on these two MasterScores
            GetStage = InstrumentStage.GetStage(MasterScoreReferenceDate1, MasterScoreReferenceDate2)
        End Get

    End Property

    'is it a problem to re-use the RatingInfo object ? 
    Public ReadOnly Property GetMasterScore(RefDate As Date) As Double
        Get
            'retrieve the rating scores at the RefDate
            RatingInfo.Retrieve(TheName_Counterparty, RefDate)

            'calculate the master score
            GetMasterScore = RatingInfo.CalculateMasterScore
        End Get
    End Property

    Public Property TheRatingInfo() As RatingInfo
        Get
            TheRatingInfo = RatingInfo
        End Get
        Set(ByVal value As RatingInfo)
            RatingInfo = value
        End Set
    End Property

    'Public Overridable ReadOnly Property GetCarryingValue(PositionDate As Date) As Decimal
    '     to be defined
    '    Get
    '        GetCarryingValue = Decimal.MinValue
    '    End Get
    'End Property


    ''' <summary>
    ''' GENERAL CHARACTERISTICS
    ''' </summary>
    ''' <returns></returns>
    Public Property TheProvision() As Decimal
        Get
            TheProvision = Provision
        End Get
        Set(ByVal value As Decimal)
            Provision = value
        End Set
    End Property

    Public Property TheReportDate() As Date
        Get
            TheReportDate = ReportDate
        End Get
        Set(ByVal value As Date)
            ReportDate = value
        End Set
    End Property

    Public ReadOnly Property GetHistoricalCost() As Decimal
        ' to be defined
        Get
            GetHistoricalCost = HistoricalCost
        End Get
    End Property

    Public ReadOnly Property GetOriginationDate() As Date
        ' to be defined
        Get
            GetOriginationDate = OriginationDate
        End Get
    End Property

    Property TheID() As Long
        Get
            Return ID
        End Get
        Set(ByVal Value As Long)
            ID = Value
        End Set
    End Property

    Property TheID_Counterparty() As Long
        Get
            Return ID_Counterparty
        End Get
        Set(ByVal Value As Long)
            ID_Counterparty = Value
        End Set
    End Property

    Property TheName_Counterparty() As String
        Get
            Return Name_Counterparty
        End Get
        Set(ByVal Value As String)
            Name_Counterparty = Value
        End Set
    End Property

    Property TheAmount() As Double
        Get
            Return Amount
        End Get
        Set(ByVal Value As Double)
            Amount = Value
        End Set
    End Property

    Property TheCurrency() As String
        Get
            Return Currency
        End Get
        Set(ByVal Value As String)
            Currency = Value
        End Set
    End Property

    Property TheOriginationDate() As Date
        Get
            Return OriginationDate
        End Get
        Set(ByVal Value As Date)
            OriginationDate = Value
        End Set
    End Property

    Property TheTransactionDate() As Date
        Get
            Return TransactionDate
        End Get
        Set(ByVal Value As Date)
            TransactionDate = Value
        End Set
    End Property

    Property TheMaturityDate() As Date
        Get
            Return MaturityDate
        End Get
        Set(ByVal Value As Date)
            MaturityDate = Value
        End Set
    End Property

    Property TheStage() As Stage
        Get
            Return InstrumentStage
        End Get
        Set(ByVal Value As Stage)
            InstrumentStage = Value
        End Set
    End Property

    Property TheInstrumentStage() As Stage.IFRSStage
        Get
            Return InstrumentStage.TheStage
        End Get
        Set(ByVal Value As Stage.IFRSStage)
            InstrumentStage.TheStage = Value
        End Set
    End Property

    Public Property DBConnection As AccessIO
        Get
            DBConnection = DB
        End Get
        Set(value As AccessIO)
            DB = value
        End Set
    End Property

End Class

