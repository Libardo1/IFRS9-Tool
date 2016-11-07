Public Class Method

    Public Enum CalculationMethod
        AllOneYear = 0
        AllLifetime = 1
        StageDependent = 2
    End Enum

    Public Enum ProvisioningMoments
        EndOfYear = 0
        YearlyFromReferenceDate = 1
    End Enum

    Public Enum MasterRatingMethod
        Average = 1
        Min = 2
    End Enum

    Public Enum PDMethod
        ConstantConditionalPD = 0
        FixedAnnualTM = 1
        ConstantContinousTM = 2
        TimeVaryingContinousTM = 3
    End Enum

    Public Enum DowngradeNotches
        None = 0
        'OneNotchesDown = 1
        'TwoNotchesDown = 2
        ThreeNotchesDown = 3

    End Enum

    Public Enum EADMethod
        ACEndOfPeriod = 0
        ACEndOfPeriod_PlusCFO = 1
        ACMidPeriod = 2
        ACAvgPeriod = 3
        FaceValueConstant = 4
        BookValueConstant = 5
    End Enum

    Public Enum LGDMethod
        Fixed = 0
        CollToExposureRatio = 1
    End Enum

    Public Enum StagingMethod
        HigherThresholdX = 0
    End Enum

    Public Enum CouponFrequency
        Annual = 0
        SemiAnnual = 1
        Quarterly = 2
    End Enum
End Class
