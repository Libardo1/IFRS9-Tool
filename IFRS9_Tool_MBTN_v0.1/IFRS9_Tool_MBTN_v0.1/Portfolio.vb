Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.Method

Public Class Portfolio
    Private PortfolioType As Type
    Private ReportDate As Date
    Private DB As AccessIO
    Private MasterRatingMethod As MasterRatingMethod

    Public Enum Type
        Receivable = 1
        Bond = 2
        Loan = 3
        Guarantee = 4
        Commitment = 5
    End Enum

    Public Property ThePortfolioType() As Type
        Get
            ThePortfolioType = PortfolioType
        End Get
        Set(ByVal value As Type)
            PortfolioType = value
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

    Public Property DBConnection As AccessIO
        Get
            DBConnection = DB
        End Get
        Set(value As AccessIO)
            DB = value
        End Set
    End Property

    Public Property TheMasterRatingMethod() As MasterRatingMethod
        Get
            TheMasterRatingMethod = MasterRatingMethod
        End Get
        Set(ByVal value As MasterRatingMethod)
            MasterRatingMethod = value
        End Set
    End Property

End Class
