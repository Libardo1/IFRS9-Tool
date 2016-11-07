Imports System.Math

Public Class DayCountConvention
    '<summary>
    '   The IDN (interest day count convention) table contains the possible interest day count conventions.
    '</summary>
    '<remarks>
    '   The main interest day count conventions are ACT/365 (typical for fixed maturities) and 30/360
    '   (typical for mortgages).
    '</remarks>

    Public ID As eIDN           'The ID of the day count convention

    Public Enum eIDN
        None = 0
        eACTACT = 1
        eACT360 = 2
        e30360 = 3
        e30365 = 4
        eACT365N = 5
        eACT360N = 6
        eACT365 = 7
        eACT366 = 8
        e366366 = 9
        e30e360 = 10
    End Enum

    Public Shared Function Yearfraction(StartDate As Date, EndDate As Date, DaycountConvention As eIDN) As Double

        Dim Result As Double, lEndDay As Long, lStartDay As Long

        Select Case DaycountConvention
            Case eIDN.eACT365 : Result = (EndDate - StartDate).TotalDays / 365
            Case eIDN.eACT360 : Result = (EndDate - StartDate).TotalDays / 360

                'SP 08-09-2011
               'Deze functie was er nog niet voor de day count conventions ACT/366 en 366/366
            Case eIDN.eACT366 : Result = (EndDate - StartDate).TotalDays / 366
            Case eIDN.e366366 : Result = (EndDate - StartDate).TotalDays / 366 'Note that this is not correct yet

            Case eIDN.eACTACT

                Dim lDaysInStartYear As Double, lDaysInEndYear As Double, lLastDayInStartYear As Date, lFirstDayInEndYear As Date
                REM                 lDaysInStartYear = DateSerial(Year(StartDate), 12, 31).d - DateSerial(Year(StartDate), 1, 1).AddDays(1)
                lDaysInStartYear = (DateSerial(Year(StartDate), 12, 31) - DateSerial(Year(StartDate), 1, 1)).Days + 1
                REM lDaysInEndYear = DateSerial(Year(EndDate), 12, 31) - DateSerial(Year(EndDate), 1, 1) + 1
                lDaysInEndYear = (DateSerial(Year(EndDate), 12, 31) - DateSerial(Year(EndDate), 1, 1)).Days + 1
                REM lLastDayInStartYear = DateSerial(Year(StartDate), 12, 31)
                lLastDayInStartYear = DateSerial(Year(StartDate), 12, 31)
                lFirstDayInEndYear = DateSerial(Year(EndDate), 1, 1)

                If Year(StartDate) = Year(EndDate) Then
                    Result = (EndDate - StartDate).TotalDays / lDaysInStartYear
                Else
                    Result = ((lLastDayInStartYear - StartDate).TotalDays / lDaysInStartYear) +
                                    (((EndDate - lFirstDayInEndYear).Days + 1) / lDaysInEndYear) +
                                    Max(Year(EndDate) - Year(StartDate) - 1, 0)
                End If

            Case eIDN.e30360

                If Month(EndDate) = 2 And Microsoft.VisualBasic.DateAndTime.Day(EndDate) + System.DateTime.IsLeapYear(Year(EndDate)) = 28 Then lEndDay = 30 Else lEndDay = Microsoft.VisualBasic.DateAndTime.Day(EndDate)
                If Month(StartDate) = 2 And Microsoft.VisualBasic.DateAndTime.Day(StartDate) + System.DateTime.IsLeapYear(Year(StartDate)) = 28 Then lStartDay = 30 Else lStartDay = Microsoft.VisualBasic.DateAndTime.Day(StartDate)

                Result = (Max(30 - lStartDay, 0) + Min(lEndDay, 30) + 360 * (Year(EndDate) - Year(StartDate)) + 30 * (Month(EndDate) - Month(StartDate) - 1))
                Result = Result / 360

            Case eIDN.e30365

                If Month(EndDate) = 2 And Microsoft.VisualBasic.DateAndTime.Day(EndDate) + System.DateTime.IsLeapYear(Year(EndDate)) = 28 Then lEndDay = 30 Else lEndDay = Microsoft.VisualBasic.DateAndTime.Day(EndDate)
                If Month(StartDate) = 2 And Microsoft.VisualBasic.DateAndTime.Day(StartDate) + System.DateTime.IsLeapYear(Year(StartDate)) = 28 Then lStartDay = 30 Else lStartDay = Microsoft.VisualBasic.DateAndTime.Day(StartDate)

                Result = (Max(30 - lStartDay, 0) + Min(lEndDay, 30) + 365 * (Year(EndDate) - Year(StartDate)) + 30 * (Month(EndDate) - Month(StartDate) - 1))
                Result = Result / 365

            Case eIDN.e30e360
                lEndDay = Microsoft.VisualBasic.DateAndTime.Day(EndDate)
                lStartDay = Microsoft.VisualBasic.DateAndTime.Day(StartDate)

                Result = (Min(lEndDay, 30) - Min(30, lStartDay)) + 360 * (Year(EndDate) - Year(StartDate)) + 30 * (Month(EndDate) - Month(StartDate))
                Result = Result / 360

            Case Else : Result = 0

        End Select
        Yearfraction = Result
    End Function

End Class