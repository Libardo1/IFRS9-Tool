Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.Method
Imports System.Math

Public Class PDCalculator
    Private ReportDate As Date
    Private DictionaryRatingSnP As Dictionary(Of String, Integer)
    Private DictionaryRatingInternal As Dictionary(Of String, Integer)
    Private TransitionMatrices As List(Of Matrix)
    Private Model As PDMethod

    Private Downgrade As DowngradeNotches

    Private DB As AccessIO

    Public Sub New(ReportDate As Date, PDMethod As PDMethod, DB As AccessIO)
        Me.ReportDate = ReportDate
        Me.Model = PDMethod
        Me.Downgrade = Downgrade
        Me.DB = DB
    End Sub

    Public Function CalcConditionalDefaultProbability(Tset As List(Of Date), StartDate As Date, EndDate As Date, Rating As String, Timing As ProvisioningMoments, eIDN As DayCountConvention.eIDN) As Double
        Dim PD As New PD

        Select Case Model
            Case PDMethod.ConstantConditionalPD
                PD = CalcConstantConditionalPDAnadolubank(Tset, StartDate, EndDate, Rating, eIDN)

            Case PDMethod.FixedAnnualTM
                PD = CalcFixedAnnualTM(Tset, StartDate, EndDate, Rating, Timing, eIDN)

            Case PDMethod.ConstantContinousTM
                PD = CalcConstantContinuousTM()

            Case PDMethod.TimeVaryingContinousTM
                PD = CalcTimeVaryingContinuousTM()
        End Select

        Return PD.ThePDValue
    End Function

    'PD method1: constant annual conditional PD
    Private Function CalcConstantConditionalPDAnadolubank(Tset As List(Of Date), StartDate_ As Date, EndDate_ As Date, Rating As String, eIDN As DayCountConvention.eIDN) As PD
        Dim PD As New PD
        Dim NumPeriod As Integer = Tset.Count
        Dim InitialDate As Date = Tset(0)
        Dim StartDate As Date = Tset.Find(Function(x As Date) x = StartDate_)
        Dim EndDate As Date = Tset.Find(Function(x) x = EndDate_)

        Select Case Downgrade
            Case DowngradeNotches.None


            Case DowngradeNotches.ThreeNotchesDown
        End Select


        'return pd value fom DB using rating
        Dim PD_Value As Double
        PD_Value = GetPDValueFromRatingAnadolubank(Rating)

        'return intensity factor
        Dim DefaultIntensity As Double
        DefaultIntensity = CalculateDefaultIntensityFromPD(PD_Value)

        'calc marginal PD
        Dim t1 As Double
        If StartDate = EndDate Then
            t1 = DayCountConvention.Yearfraction(ReportDate, Tset(NumPeriod - 3), eIDN) 'need the date just before the last one. But last one is counted twice (principal + interest)
        Else
            t1 = DayCountConvention.Yearfraction(ReportDate, StartDate, eIDN)
        End If

        Dim t2 As Double
        t2 = DayCountConvention.Yearfraction(ReportDate, EndDate, eIDN)

        Dim ConditionalPD_Value As Double

        ConditionalPD_Value = Math.Exp(-DefaultIntensity * t1) - Math.Exp(-DefaultIntensity * t2)

        'set PD value to PD object
        PD.ThePDValue = ConditionalPD_Value

        Return PD
    End Function

    Private Function CalcConstantConditionalPD(Tset As List(Of Date), StartDate_ As Date, EndDate_ As Date, Rating As String, eIDN As DayCountConvention.eIDN) As PD
        Dim PD As New PD
        Dim NumPeriod As Integer = Tset.Count
        Dim InitialDate As Date = Tset(0)
        Dim StartDate As Date = Tset.Find(Function(x As Date) x = StartDate_)
        Dim EndDate As Date = Tset.Find(Function(x) x = EndDate_)

        'return pd value fom DB using rating
        Dim PD_Value As Double
        PD_Value = GetPDValueFromRating(Rating, InitialDate)

        'return intensity factor
        Dim DefaultIntensity As Double
        DefaultIntensity = CalculateDefaultIntensityFromPD(PD_Value)

        'calc marginal PD
        Dim t1 As Double
        Dim t2 As Double
        If StartDate = EndDate Then
            t1 = DayCountConvention.Yearfraction(ReportDate, Tset(NumPeriod - 3), eIDN) 'need the date just before the last one. But last one is counted twice (principal + interest)
        Else
            t1 = DayCountConvention.Yearfraction(ReportDate, StartDate, eIDN)
        End If

        t2 = DayCountConvention.Yearfraction(ReportDate, EndDate, eIDN)

        Dim ConditionalPD_Value As Double

        ConditionalPD_Value = Math.Exp(-DefaultIntensity * t1) - Math.Exp(-DefaultIntensity * t2)

        'set PD value to PD object
        PD.ThePDValue = ConditionalPD_Value

        Return PD
    End Function

    Private Function CalcFixedAnnualTM(TSet As List(Of Date), StartDate_ As Date, EndDate_ As Date, Rating As String, Timing As ProvisioningMoments, eIDN As DayCountConvention.eIDN) As PD
        Dim PD As New PD
        Dim ConditionalPD_Value As Double

        Dim StartDate As Date = TSet.Find(Function(x As Date) x = StartDate_)
        Dim EndDate As Date = TSet.Find(Function(x) x = EndDate_)

        Dim NumDates As Integer = TSet.Count
        Dim LastDate As Date = TSet(NumDates - 1)

        'time periods
        Dim t1 As Double = DayCountConvention.Yearfraction(ReportDate, TSet(1), eIDN)
        Dim t_Middle1 As Double
        Dim t_Middle2 As Double
        Dim t_BeforeLast As Double
        Dim t_E

        'used to compare time interval with report date
        Dim compare1 As Integer = DateTime.Compare(StartDate, ReportDate)
        Dim compare2 As Integer = DateTime.Compare(EndDate, LastDate)

        ''Scale Ratings
        'ScaleRatings()

        Select Case Timing
            Case ProvisioningMoments.EndOfYear
                If StartDate = ReportDate Then
                    'get TM^1
                    GetTM1(ReportDate)

                    ConditionalPD_Value = GetPDFromTM(Rating, 1) * t1

                ElseIf compare1 > 0 And compare2 < 0 Then
                    t_Middle1 = DayCountConvention.Yearfraction(ReportDate, StartDate, eIDN)
                    t_Middle2 = DayCountConvention.Yearfraction(ReportDate, EndDate, eIDN)

                    'calc TM^n until TM^N-1
                    CalcTransitionMatrices(ReportDate, NumDates - 2)

                    Dim m1 As Double = GetPDFromTM(Rating, t_Middle2 - t1)
                    Dim m2 As Double = GetPDFromTM(Rating, t_Middle1 - t1)
                    Dim m3 As Double = GetPDFromTM(Rating, 1) * t1

                    ConditionalPD_Value = (m1 - m2) * (1 - m3)

                Else
                    t_BeforeLast = DayCountConvention.Yearfraction(ReportDate, StartDate, eIDN)
                    t_E = DayCountConvention.Yearfraction(ReportDate, LastDate, eIDN)

                    'calc TM^n until TM^N
                    CalcTransitionMatrices(ReportDate, NumDates - 1)

                    ConditionalPD_Value = (GetPDFromTM(Rating, t_BeforeLast - t1 + 1) - GetPDFromTM(Rating, t_BeforeLast - t1)) * (t_E - t_BeforeLast) * (1 - (GetPDFromTM(Rating, 1) * t1))
                End If

                PD.ThePDValue = ConditionalPD_Value

            Case ProvisioningMoments.YearlyFromReferenceDate
                If DateTime.Compare(EndDate, LastDate) < 0 Then
                    t_Middle1 = DayCountConvention.Yearfraction(ReportDate, StartDate, eIDN)
                    t_Middle2 = DayCountConvention.Yearfraction(ReportDate, EndDate, eIDN)

                    'calc TM^n until TM^N-1
                    CalcTransitionMatrices(ReportDate, NumDates - 2)

                    Dim m1 As Double = GetPDFromTM(Rating, t_Middle2)
                    Dim m2 As Double = GetPDFromTM(Rating, t_Middle1)

                    ConditionalPD_Value = m1 - m2

                Else
                    t_BeforeLast = DayCountConvention.Yearfraction(ReportDate, StartDate, eIDN)
                    t_E = DayCountConvention.Yearfraction(ReportDate, LastDate, eIDN)

                    'calc TM^n until TM^N
                    CalcTransitionMatrices(ReportDate, NumDates - 1)

                    Dim m1 As Double = GetPDFromTM(Rating, t_BeforeLast + 1)
                    Dim m2 As Double = GetPDFromTM(Rating, t_BeforeLast)

                    ConditionalPD_Value = (m1 - m2) * (t_E - t_BeforeLast)

                End If

                PD.ThePDValue = ConditionalPD_Value

            Case Else
                Throw New Exception("Provisioning timing doesn't exist.")
        End Select

        Return PD
    End Function

    'PD method3: constant continuous TM
    Private Function CalcConstantContinuousTM() As PD
        'to do: use TM in DB. Calc exponential of this matrix and return good row (see below)


        Return New PD()
    End Function

    'PD method4: time varying continuous TM
    Private Function CalcTimeVaryingContinuousTM() As PD
        'to do: use TM in DB. Calc exponential of this matrix and return good row (see below)


        Return New PD()
    End Function

    'return PD_Value from the transition probability table
    'could be reach via another table
    'DEBUG: OK
    Private Function GetPDValueFromRatingAnadolubank(Rating As String) As Double
        Dim PD_Value As Double

        DB.OpenConnection()

        Dim Query As String = "Select *
                               FROM [Dat_RatingsPDAnadolubank]
                               WHERE Rating = '" & Rating.ToString & "'"

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        While (DB.Reader.Read())
            PD_Value = DB.Reader("PD")
        End While

        DB.CloseReader()
        DB.CloseConnection()

        Return PD_Value
    End Function

    Private Function GetPDValueFromRating(Rating As String, Ref_Year As Date) As Double
        Dim PD_Value As Double

        DB.OpenConnection()

        Dim ID_Matrix As Integer = 1 'needed for the moment
        Dim Query As String = "Select Prob 
                               FROM [TransitionProbabilities]
                               WHERE ID_Matrix = " + ID_Matrix.ToString + " AND
                                     Date_Start <= #" + Ref_Year.ToString + "# AND Date_End > #" + Ref_Year.ToString + "# AND
                                     RatingFrom = '" + Rating.ToString + "' AND RatingTo = '4' "

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        While (DB.Reader.Read())
            PD_Value = DB.Reader("Prob")
        End While

        DB.CloseReader()
        DB.CloseConnection()

        Return PD_Value
    End Function

    'default intensity param from PD
    Private Function CalculateDefaultIntensityFromPD(PD_Value As Double) As Double
        Dim DefaultIntensity As Double

        DefaultIntensity = -Math.Log(1 - PD_Value)

        Return DefaultIntensity
    End Function

    'PD from default intensity param
    Public Shared Function CalculatePDFromDefaultIntensity(DefaultIntensity As Double) As Double
        Dim PD_Value As Double

        PD_Value = 1 - Math.Exp(-DefaultIntensity)

        Return PD_Value
    End Function

    'DEBUG OK
    Public Sub GetTM1(Ref_Date)
        TransitionMatrices = New List(Of Matrix)

        'retrive TM^1 fron DB
        'Note: do we need a ID_Matrix here?
        Dim TM1 As New Matrix
        TM1.RetrieveTPs(Ref_Date, DB)
        TransitionMatrices.Add(TM1)
    End Sub

    Public Sub CalcTransitionMatrices(Ref_Date As Date, NumYear As Integer) 'NumYear is the number of period ahead for the lifetimePD
        'get TM1
        GetTM1(Ref_Date)

        'calc TM^n (2 -> n)
        Dim TM As New Matrix
        TM = TransitionMatrices(0)
        Dim TMn As Matrix
        For i As Integer = 2 To NumYear
            TMn = New Matrix
            TMn = TM.MultiplyPower(i)
            TransitionMatrices.Add(TMn)
        Next
    End Sub

    'Return last column of TM^n, given initial rating Ro (S&P rating is used)
    Public Function GetPDFromTM(Rating_Counterparty As String, NumYear As Integer) As Double
        If NumYear = 0 Then
            Return 0
        End If

        Dim PD_TM As Double

        'go to TM^n last column, row "Rank"
        Dim Col As Integer = TransitionMatrices(0).NumCol
        Dim Row As Integer = Rating_Counterparty
        PD_TM = TransitionMatrices(NumYear - 1).GetElement(Col - 1, Row - 1) 'index starts at 0 in Matrix Class

        Return PD_TM
    End Function


    'Scale the ratings (0 -> N): works only for S&P
    'DEBUG: OK
    'Public Sub ScaleRatings()
    '    DB.OpenConnection()

    '    'fill S&P rating dictonary
    '    Dim Query As String = "Select * 
    '                           FROM [Dat_Rating]
    '                           WHERE Rating_Type = 'S&P'"

    '    DB.CreateCommand(Query)
    '    DB.ExecuteReader()

    '    DictionaryRatingSnP = New Dictionary(Of String, Integer)
    '    While (DB.Reader.Read())
    '        Dim Rating As String = DB.Reader("Rating")
    '        Dim Rank As Integer = DB.Reader("Rank")
    '        DictionaryRatingSnP.Add(Rating, Rank)
    '    End While

    '    DB.CloseReader()

    '    'fill Internal rating dictonary
    '    Query = "Select * 
    '             FROM [Dat_Rating]
    '             WHERE Rating_Type = 'Internal'"

    '    DB.CreateCommand(Query)
    '    DB.ExecuteReader()

    '    DictionaryRatingInternal = New Dictionary(Of String, Integer)
    '    While (DB.Reader.Read())
    '        Dim Rating As String = DB.Reader("Rating")
    '        Dim Rank As Integer = DB.Reader("Rank")

    '        DictionaryRatingInternal.Add(Rating, Rank)
    '    End While

    '    DB.CloseReader()

    '    DB.CloseConnection()
    'End Sub

    ''return the rank given the counterparty rating
    ''VERIFY WITH ANOTHER RANKING
    'Private Function GetRank(Rating_Counterparty As String, DictionaryRatings As Dictionary(Of String, Integer))
    '    Dim Rank As Integer

    '    Dim pair As KeyValuePair(Of String, Integer)
    '    For Each pair In DictionaryRatings
    '        If pair.Key = Rating_Counterparty Then
    '            Rank = pair.Value
    '            Return Rank
    '        End If
    '    Next

    '    Return New Exception("Not rating found")
    'End Function

    ''return constant PD 
    'Public Function GetLifeTimePDConstant(Rating_Counterparty As String, Ref_Year As Date, DB As AccessIO) As Double
    '    DB.OpenConnection()

    '    Dim ID_Matrix As Integer = 1
    '    Dim Query As String = "Select Prob 
    '                           FROM [TransitionProbabilities]
    '                           WHERE Matrix_ID = " + ID_Matrix.ToString + " AND
    '                                 #Date_Start# <= " + Ref_Year.ToString + " AND #Date_End# > " + Ref_Year.ToString + " AND
    '                                 RatingFrom = 'A' AND RatingTo = 'D' "

    '    DB.CreateCommand(Query)
    '    DB.ExecuteReader()

    '    Dim Prob As Double
    '    While (DB.Reader.Read())
    '        Prob = DB.Reader("Prob")
    '    End While

    '    DB.CloseReader()
    '    DB.CloseConnection()

    '    Return Prob
    'End Function

    Public Property TheModel() As Method.PDMethod
        Get
            TheModel = Model
        End Get
        Set(ByVal value As Method.PDMethod)
            Model = value
        End Set
    End Property

End Class
