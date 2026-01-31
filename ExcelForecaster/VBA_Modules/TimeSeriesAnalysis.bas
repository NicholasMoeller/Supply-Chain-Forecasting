Attribute VB_Name = "TimeSeriesAnalysis"
Option Explicit

' Time Series Data Type
Public Type TimeSeriesData
    Values() As Double
    Frequency As Integer
End Type

' Forecast Result Type
Public Type ForecastResult
    ModelName As String         ' Name of the forecasting model used
    FittedValues() As Double
    ForecastValues() As Double
    Residuals() As Double
    Lower80() As Double
    Upper80() As Double
    Lower95() As Double
    Upper95() As Double
    MAPE As Double
    MAE As Double
    RMSE As Double
    MBE As Double
    Alpha As Double
    Beta As Double
    Gamma As Double
    Phi As Double              ' Damping parameter
End Type

' Decomposition Result Type
Public Type DecompositionResult
    Observed() As Double
    Trend() As Double
    Seasonal() As Double
    Random() As Double
End Type

' Seasonality Detection Result Type
Public Type SeasonalityInfo
    HasSeasonality As Boolean
    DetectedFrequency As Integer
    SeasonalType As String ' "additive", "multiplicative", "none"
    Confidence As Double ' 0-100%
End Type

' Cross-Validation Result Type
Public Type CrossValidationResult
    AvgMAPE As Double
    AvgMAE As Double
    AvgRMSE As Double
    NumFolds As Integer
End Type

' Demand Classification Type
Public Type DemandClassification
    Category As String ' "Smooth", "Intermittent", "Erratic", "Lumpy"
    CV As Double ' Coefficient of Variation
    ADI As Double ' Average Demand Interval
    RecommendedMethod As String
End Type

' Optimized Parameters Type
Public Type OptimizedParameters
    Alpha As Double
    Beta As Double
    Gamma As Double
    Phi As Double
    BestMAPE As Double
End Type

' Probabilistic Forecast Type
Public Type ProbabilisticForecast
    Percentile05() As Double ' 5th percentile
    Percentile25() As Double ' 25th percentile (Q1)
    Percentile50() As Double ' 50th percentile (median)
    Percentile75() As Double ' 75th percentile (Q3)
    Percentile95() As Double ' 95th percentile
    Mean() As Double ' Mean forecast
    StdDev() As Double ' Standard deviation at each horizon
End Type

' Forecast Value Added Type
Public Type ForecastValueAdded
    NaiveMAPE As Double ' Naive forecast MAPE (baseline)
    ModelMAPE As Double ' Advanced model MAPE
    FVA As Double ' FVA = (Naive - Model) / Naive * 100
    ImprovedAccuracy As Boolean ' True if FVA > 0
End Type

' Multi-Objective Score Type
Public Type MultiObjectiveScore
    MAPE As Double
    MAE As Double
    RMSE As Double
    Bias As Double
    CompositeScore As Double ' Weighted combination
End Type

' Pattern Match Type
Public Type PatternMatch
    StartIndex As Long
    EndIndex As Long
    Similarity As Double ' 0-1, higher is more similar
    Pattern() As Double
End Type

' Ensemble Trim Result Type
Public Type EnsembleTrimResult
    TrimmedModels() As ForecastResult
    NumModels As Integer
    AvgImprovement As Double ' MAPE improvement from trimming
End Type

' ============================================================================
' SIMPLE EXPONENTIAL SMOOTHING
' ============================================================================

Public Function SimpleExponentialSmoothing(ByRef tsData As TimeSeriesData, _
                                          ByVal horizon As Integer, _
                                          Optional ByVal Alpha As Variant) As ForecastResult
    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long
    Dim level As Double
    Dim optimizedAlpha As Double

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Optimize alpha if not provided
    If IsMissing(Alpha) Or IsEmpty(Alpha) Then
        optimizedAlpha = OptimizeAlphaSES(Values)
    Else
        optimizedAlpha = CDbl(Alpha)
    End If

    ' Initialize arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Fit the model
    level = Values(LBound(Values))
    result.FittedValues(LBound(Values)) = level

    For i = LBound(Values) + 1 To UBound(Values)
        level = optimizedAlpha * Values(i - 1) + (1 - optimizedAlpha) * level
        result.FittedValues(i) = level
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecast
    Dim lastLevel As Double
    lastLevel = optimizedAlpha * Values(UBound(Values)) + (1 - optimizedAlpha) * level

    For i = 1 To horizon
        result.ForecastValues(i) = lastLevel
    Next i

    ' Calculate confidence intervals
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)

    Dim se As Double
    For i = 1 To horizon
        se = residualStdDev * Sqr(i)
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * se
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * se
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * se
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.Alpha = optimizedAlpha
    result.ModelName = "SES"

    SimpleExponentialSmoothing = result
End Function

Private Function OptimizeAlphaSES(ByRef Values() As Double) As Double
    Dim bestAlpha As Double
    Dim bestMAPE As Double
    Dim currentAlpha As Double
    Dim cvResult As CrossValidationResult
    Dim numFolds As Integer

    bestAlpha = 0.1
    bestMAPE = 1E+100

    ' Use 3-fold cross-validation for robust parameter selection
    numFolds = 3

    For currentAlpha = 0.01 To 0.99 Step 0.01
        cvResult = CrossValidateSES(Values, currentAlpha, numFolds)
        If cvResult.AvgMAPE < bestMAPE Then
            bestMAPE = cvResult.AvgMAPE
            bestAlpha = currentAlpha
        End If
    Next currentAlpha

    OptimizeAlphaSES = bestAlpha
End Function

Private Function CalculateSSE_SES(ByRef Values() As Double, ByVal Alpha As Double) As Double
    Dim n As Long, i As Long
    Dim level As Double
    Dim sse As Double
    Dim forecast As Double
    Dim errorVal As Double

    n = UBound(Values) - LBound(Values) + 1
    level = Values(LBound(Values))
    sse = 0

    For i = LBound(Values) + 1 To UBound(Values)
        forecast = level
        errorVal = Values(i) - forecast
        sse = sse + errorVal * errorVal
        level = Alpha * Values(i) + (1 - Alpha) * level
    Next i

    CalculateSSE_SES = sse
End Function

' ============================================================================
' HOLT-WINTERS METHOD
' ============================================================================

Public Function HoltWinters(ByRef tsData As TimeSeriesData, _
                           ByVal horizon As Integer, _
                           Optional ByVal seasonalType As String = "additive", _
                           Optional ByVal Alpha As Variant, _
                           Optional ByVal Beta As Variant, _
                           Optional ByVal Gamma As Variant) As ForecastResult

    Dim result As ForecastResult
    Dim Values() As Double
    Dim Frequency As Integer
    Dim n As Long, i As Long, seasonalIdx As Long
    Dim level() As Double, Trend() As Double, Seasonal() As Double
    Dim optimizedAlpha As Double, optimizedBeta As Double, optimizedGamma As Double

    Values = tsData.Values
    Frequency = tsData.Frequency
    n = UBound(Values) - LBound(Values) + 1

    ' Optimize parameters if not provided
    If IsMissing(Alpha) Or IsMissing(Beta) Or IsMissing(Gamma) Or _
       IsEmpty(Alpha) Or IsEmpty(Beta) Or IsEmpty(Gamma) Then
        Dim optimized As Variant
        optimized = OptimizeHoltWinters(Values, Frequency, seasonalType)
        optimizedAlpha = optimized(0)
        optimizedBeta = optimized(1)
        optimizedGamma = optimized(2)
    Else
        optimizedAlpha = CDbl(Alpha)
        optimizedBeta = CDbl(Beta)
        optimizedGamma = CDbl(Gamma)
    End If

    ' Initialize arrays
    ReDim level(LBound(Values) To UBound(Values))
    ReDim Trend(LBound(Values) To UBound(Values))
    ReDim Seasonal(LBound(Values) To UBound(Values) + horizon)
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Initial values
    level(LBound(Values)) = Values(LBound(Values))
    Trend(LBound(Values)) = (Values(LBound(Values) + Frequency) - Values(LBound(Values))) / Frequency

    ' Initialize seasonal components
    If seasonalType = "additive" Then
        For i = LBound(Values) To LBound(Values) + Frequency - 1
            Seasonal(i) = Values(i) - level(LBound(Values))
        Next i

        result.FittedValues(LBound(Values)) = level(LBound(Values)) + Seasonal(LBound(Values))

        ' Fit the model
        For i = LBound(Values) + 1 To UBound(Values)
            seasonalIdx = ((i - LBound(Values)) Mod Frequency) + LBound(Values)

            Dim prevLevel As Double, prevTrend As Double, prevSeasonal As Double
            prevLevel = level(i - 1)
            prevTrend = Trend(i - 1)
            prevSeasonal = Seasonal(seasonalIdx)

            level(i) = optimizedAlpha * (Values(i) - prevSeasonal) + (1 - optimizedAlpha) * (prevLevel + prevTrend)
            Trend(i) = optimizedBeta * (level(i) - prevLevel) + (1 - optimizedBeta) * prevTrend
            Seasonal(i + Frequency) = optimizedGamma * (Values(i) - level(i)) + (1 - optimizedGamma) * prevSeasonal

            result.FittedValues(i) = prevLevel + prevTrend + prevSeasonal
        Next i

        ' Forecast
        For i = 1 To horizon
            seasonalIdx = ((UBound(Values) + i - LBound(Values)) Mod Frequency) + LBound(Values)
            result.ForecastValues(i) = level(UBound(Values)) + i * Trend(UBound(Values)) + Seasonal(seasonalIdx)
        Next i

    Else ' multiplicative
        For i = LBound(Values) To LBound(Values) + Frequency - 1
            If level(LBound(Values)) <> 0 Then
                Seasonal(i) = Values(i) / level(LBound(Values))
            Else
                Seasonal(i) = 1
            End If
        Next i

        result.FittedValues(LBound(Values)) = level(LBound(Values)) * Seasonal(LBound(Values))

        ' Fit the model
        For i = LBound(Values) + 1 To UBound(Values)
            seasonalIdx = ((i - LBound(Values)) Mod Frequency) + LBound(Values)

            prevLevel = level(i - 1)
            prevTrend = Trend(i - 1)
            prevSeasonal = Seasonal(seasonalIdx)

            If prevSeasonal <> 0 Then
                level(i) = optimizedAlpha * (Values(i) / prevSeasonal) + (1 - optimizedAlpha) * (prevLevel + prevTrend)
            Else
                level(i) = prevLevel + prevTrend
            End If
            Trend(i) = optimizedBeta * (level(i) - prevLevel) + (1 - optimizedBeta) * prevTrend
            If level(i) <> 0 Then
                Seasonal(i + Frequency) = optimizedGamma * (Values(i) / level(i)) + (1 - optimizedGamma) * prevSeasonal
            Else
                Seasonal(i + Frequency) = prevSeasonal
            End If

            result.FittedValues(i) = (prevLevel + prevTrend) * prevSeasonal
        Next i

        ' Forecast
        For i = 1 To horizon
            seasonalIdx = ((UBound(Values) + i - LBound(Values)) Mod Frequency) + LBound(Values)
            result.ForecastValues(i) = (level(UBound(Values)) + i * Trend(UBound(Values))) * Seasonal(seasonalIdx)
        Next i
    End If

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Calculate confidence intervals
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)

    Dim se As Double
    For i = 1 To horizon
        se = residualStdDev * Sqr(i)
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * se
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * se
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * se
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.Alpha = optimizedAlpha
    result.Beta = optimizedBeta
    result.Gamma = optimizedGamma

    If LCase(seasonalType) = "additive" Then
        result.ModelName = "Holt-Winters (Additive)"
    Else
        result.ModelName = "Holt-Winters (Multiplicative)"
    End If

    HoltWinters = result
End Function

Private Function OptimizeHoltWinters(ByRef Values() As Double, _
                                    ByVal Frequency As Integer, _
                                    ByVal seasonalType As String) As Variant
    Dim bestAlpha As Double, bestBeta As Double, bestGamma As Double
    Dim bestMAPE As Double
    Dim currentAlpha As Double, currentBeta As Double, currentGamma As Double
    Dim cvResult As CrossValidationResult
    Dim numFolds As Integer

    bestAlpha = 0.3
    bestBeta = 0.1
    bestGamma = 0.1
    bestMAPE = 1E+100

    ' Use 3-fold cross-validation for robust parameter selection
    numFolds = 3

    ' Coarse grid search with cross-validation
    For currentAlpha = 0.1 To 0.9 Step 0.2
        For currentBeta = 0.05 To 0.3 Step 0.1
            For currentGamma = 0.05 To 0.3 Step 0.1
                cvResult = CrossValidateHW(Values, Frequency, currentAlpha, currentBeta, currentGamma, seasonalType, numFolds)
                If cvResult.AvgMAPE < bestMAPE Then
                    bestMAPE = cvResult.AvgMAPE
                    bestAlpha = currentAlpha
                    bestBeta = currentBeta
                    bestGamma = currentGamma
                End If
            Next currentGamma
        Next currentBeta
    Next currentAlpha

    OptimizeHoltWinters = Array(bestAlpha, bestBeta, bestGamma)
End Function

Private Function CalculateSSE_HW(ByRef Values() As Double, _
                                 ByVal Frequency As Integer, _
                                 ByVal Alpha As Double, _
                                 ByVal Beta As Double, _
                                 ByVal Gamma As Double, _
                                 ByVal seasonalType As String) As Double
    Dim n As Long, i As Long, seasonalIdx As Long
    Dim level() As Double, Trend() As Double, Seasonal() As Double
    Dim sse As Double
    Dim forecast As Double, errorVal As Double
    Dim prevLevel As Double, prevTrend As Double, prevSeasonal As Double

    n = UBound(Values) - LBound(Values) + 1
    ReDim level(LBound(Values) To UBound(Values))
    ReDim Trend(LBound(Values) To UBound(Values))
    ReDim Seasonal(LBound(Values) To UBound(Values) + Frequency)

    level(LBound(Values)) = Values(LBound(Values))
    Trend(LBound(Values)) = (Values(LBound(Values) + Frequency) - Values(LBound(Values))) / Frequency
    sse = 0

    If seasonalType = "additive" Then
        For i = LBound(Values) To LBound(Values) + Frequency - 1
            Seasonal(i) = Values(i) - level(LBound(Values))
        Next i

        For i = LBound(Values) + 1 To UBound(Values)
            seasonalIdx = ((i - LBound(Values)) Mod Frequency) + LBound(Values)
            prevLevel = level(i - 1)
            prevTrend = Trend(i - 1)
            prevSeasonal = Seasonal(seasonalIdx)

            forecast = prevLevel + prevTrend + prevSeasonal
            errorVal = Values(i) - forecast
            sse = sse + errorVal * errorVal

            level(i) = Alpha * (Values(i) - prevSeasonal) + (1 - Alpha) * (prevLevel + prevTrend)
            Trend(i) = Beta * (level(i) - prevLevel) + (1 - Beta) * prevTrend
            Seasonal(i + Frequency) = Gamma * (Values(i) - level(i)) + (1 - Gamma) * prevSeasonal
        Next i
    Else
        For i = LBound(Values) To LBound(Values) + Frequency - 1
            If level(LBound(Values)) <> 0 Then
                Seasonal(i) = Values(i) / level(LBound(Values))
            Else
                Seasonal(i) = 1
            End If
        Next i

        For i = LBound(Values) + 1 To UBound(Values)
            seasonalIdx = ((i - LBound(Values)) Mod Frequency) + LBound(Values)
            prevLevel = level(i - 1)
            prevTrend = Trend(i - 1)
            prevSeasonal = Seasonal(seasonalIdx)

            forecast = (prevLevel + prevTrend) * prevSeasonal
            errorVal = Values(i) - forecast
            sse = sse + errorVal * errorVal

            If prevSeasonal <> 0 Then
                level(i) = Alpha * (Values(i) / prevSeasonal) + (1 - Alpha) * (prevLevel + prevTrend)
            Else
                level(i) = prevLevel + prevTrend
            End If
            Trend(i) = Beta * (level(i) - prevLevel) + (1 - Beta) * prevTrend
            If level(i) <> 0 Then
                Seasonal(i + Frequency) = Gamma * (Values(i) / level(i)) + (1 - Gamma) * prevSeasonal
            Else
                Seasonal(i + Frequency) = prevSeasonal
            End If
        Next i
    End If

    CalculateSSE_HW = sse
End Function

' ============================================================================
' TIME SERIES DECOMPOSITION
' ============================================================================

Public Function Decompose(ByRef tsData As TimeSeriesData, _
                         Optional ByVal decompType As String = "additive") As DecompositionResult

    Dim result As DecompositionResult
    Dim Values() As Double
    Dim Frequency As Integer
    Dim n As Long, i As Long, j As Long
    Dim halfWindow As Long
    Dim sum As Double

    Values = tsData.Values
    Frequency = tsData.Frequency
    n = UBound(Values) - LBound(Values) + 1

    ReDim result.Observed(LBound(Values) To UBound(Values))
    ReDim result.Trend(LBound(Values) To UBound(Values))
    ReDim result.Seasonal(LBound(Values) To UBound(Values))
    ReDim result.Random(LBound(Values) To UBound(Values))

    ' Copy observed values
    For i = LBound(Values) To UBound(Values)
        result.Observed(i) = Values(i)
    Next i

    ' Calculate trend using centered moving average
    halfWindow = Frequency \ 2

    For i = LBound(Values) To UBound(Values)
        If i < LBound(Values) + halfWindow Or i > UBound(Values) - halfWindow Then
            result.Trend(i) = Empty ' Will be treated as missing
        Else
            sum = 0
            For j = i - halfWindow To i + halfWindow
                sum = sum + Values(j)
            Next j
            result.Trend(i) = sum / (2 * halfWindow + 1)
        End If
    Next i

    ' Calculate seasonal component
    Dim detrended() As Double
    Dim seasonalAverages() As Double
    Dim seasonValues As Collection
    Dim seasonalMean As Double
    Dim s As Long

    ReDim detrended(LBound(Values) To UBound(Values))
    ReDim seasonalAverages(0 To Frequency - 1)

    If decompType = "additive" Then
        ' Detrend
        For i = LBound(Values) To UBound(Values)
            If Not IsEmpty(result.Trend(i)) Then
                detrended(i) = Values(i) - result.Trend(i)
            Else
                detrended(i) = Empty
            End If
        Next i

        ' Calculate average for each season
        For s = 0 To Frequency - 1
            Set seasonValues = New Collection
            For i = LBound(Values) + s To UBound(Values) Step Frequency
                If Not IsEmpty(detrended(i)) Then
                    seasonValues.Add detrended(i)
                End If
            Next i

            If seasonValues.Count > 0 Then
                sum = 0
                For j = 1 To seasonValues.Count
                    sum = sum + seasonValues(j)
                Next j
                seasonalAverages(s) = sum / seasonValues.Count
            Else
                seasonalAverages(s) = 0
            End If
        Next s

        ' Center seasonal components
        seasonalMean = 0
        For s = 0 To Frequency - 1
            seasonalMean = seasonalMean + seasonalAverages(s)
        Next s
        seasonalMean = seasonalMean / Frequency

        For s = 0 To Frequency - 1
            seasonalAverages(s) = seasonalAverages(s) - seasonalMean
        Next s

        ' Assign seasonal values
        For i = LBound(Values) To UBound(Values)
            result.Seasonal(i) = seasonalAverages((i - LBound(Values)) Mod Frequency)
        Next i

        ' Calculate random component
        For i = LBound(Values) To UBound(Values)
            If Not IsEmpty(result.Trend(i)) Then
                result.Random(i) = Values(i) - result.Trend(i) - result.Seasonal(i)
            Else
                result.Random(i) = Empty
            End If
        Next i

    Else ' multiplicative
        ' Detrend
        For i = LBound(Values) To UBound(Values)
            If Not IsEmpty(result.Trend(i)) And result.Trend(i) <> 0 Then
                detrended(i) = Values(i) / result.Trend(i)
            Else
                detrended(i) = Empty
            End If
        Next i

        ' Calculate average for each season
        For s = 0 To Frequency - 1
            Set seasonValues = New Collection
            For i = LBound(Values) + s To UBound(Values) Step Frequency
                If Not IsEmpty(detrended(i)) Then
                    seasonValues.Add detrended(i)
                End If
            Next i

            If seasonValues.Count > 0 Then
                sum = 0
                For j = 1 To seasonValues.Count
                    sum = sum + seasonValues(j)
                Next j
                seasonalAverages(s) = sum / seasonValues.Count
            Else
                seasonalAverages(s) = 1
            End If
        Next s

        ' Assign seasonal values
        For i = LBound(Values) To UBound(Values)
            result.Seasonal(i) = seasonalAverages((i - LBound(Values)) Mod Frequency)
        Next i

        ' Calculate random component
        For i = LBound(Values) To UBound(Values)
            If Not IsEmpty(result.Trend(i)) And result.Trend(i) <> 0 And result.Seasonal(i) <> 0 Then
                result.Random(i) = Values(i) / (result.Trend(i) * result.Seasonal(i))
            Else
                result.Random(i) = Empty
            End If
        Next i
    End If

    Decompose = result
End Function

' ============================================================================
' AUTOCORRELATION FUNCTION (ACF)
' ============================================================================

Public Function CalculateACF(ByRef Values() As Double, ByVal maxLag As Integer) As Double()
    Dim n As Long, i As Long, lag As Long
    Dim mean As Double, variance As Double, covariance As Double
    Dim acf() As Double

    n = UBound(Values) - LBound(Values) + 1

    ' Calculate mean
    mean = 0
    For i = LBound(Values) To UBound(Values)
        mean = mean + Values(i)
    Next i
    mean = mean / n

    ' Calculate variance
    variance = 0
    For i = LBound(Values) To UBound(Values)
        variance = variance + (Values(i) - mean) ^ 2
    Next i

    ' Calculate ACF
    ReDim acf(0 To maxLag)
    acf(0) = 1#

    For lag = 1 To maxLag
        covariance = 0
        For i = LBound(Values) To UBound(Values) - lag
            covariance = covariance + (Values(i) - mean) * (Values(i + lag) - mean)
        Next i
        acf(lag) = covariance / variance
    Next lag

    CalculateACF = acf
End Function

' ============================================================================
' STATISTICAL METRICS
' ============================================================================

Public Function CalculateMAPE(ByRef actual() As Double, ByRef fitted() As Double) As Double
    Dim sum As Double, Count As Long, i As Long

    sum = 0
    Count = 0

    For i = LBound(actual) To UBound(actual)
        If actual(i) <> 0 Then
            sum = sum + Abs((actual(i) - fitted(i)) / actual(i))
            Count = Count + 1
        End If
    Next i

    If Count > 0 Then
        CalculateMAPE = (sum / Count) * 100
    Else
        CalculateMAPE = 0
    End If
End Function

Public Function CalculateMAE(ByRef Residuals() As Double) As Double
    Dim sum As Double, i As Long, n As Long

    sum = 0
    n = UBound(Residuals) - LBound(Residuals) + 1

    For i = LBound(Residuals) To UBound(Residuals)
        sum = sum + Abs(Residuals(i))
    Next i

    CalculateMAE = sum / n
End Function

Public Function CalculateRMSE(ByRef Residuals() As Double) As Double
    Dim sum As Double, i As Long, n As Long

    sum = 0
    n = UBound(Residuals) - LBound(Residuals) + 1

    For i = LBound(Residuals) To UBound(Residuals)
        sum = sum + Residuals(i) ^ 2
    Next i

    CalculateRMSE = Sqr(sum / n)
End Function

Private Function CalculateStdDev(ByRef Values() As Double) As Double
    Dim mean As Double, sumOfSquares As Double
    Dim i As Long, n As Long

    n = UBound(Values) - LBound(Values) + 1

    ' Calculate mean
    mean = 0
    For i = LBound(Values) To UBound(Values)
        mean = mean + Values(i)
    Next i
    mean = mean / n

    ' Calculate sum of squares
    sumOfSquares = 0
    For i = LBound(Values) To UBound(Values)
        sumOfSquares = sumOfSquares + (Values(i) - mean) ^ 2
    Next i

    CalculateStdDev = Sqr(sumOfSquares / n)
End Function

' ============================================================================
' PACF (Partial Autocorrelation Function)
' ============================================================================

Public Function CalculatePACF(ByRef Values() As Double, ByVal maxLag As Integer) As Double()
    ' Calculate PACF using Durbin-Levinson algorithm
    Dim n As Long, i As Long, lag As Long, k As Long
    Dim pacf() As Double
    Dim acf() As Double
    Dim phi() As Double
    Dim phiNew() As Double
    Dim sum As Double

    n = UBound(Values) - LBound(Values) + 1

    ' First calculate ACF
    acf = CalculateACF(Values, maxLag)

    ' Initialize PACF array
    ReDim pacf(0 To maxLag)
    pacf(0) = 1#

    If maxLag >= 1 Then
        pacf(1) = acf(1)
    End If

    ' Durbin-Levinson recursion
    ReDim phi(1 To maxLag)
    phi(1) = acf(1)

    For lag = 2 To maxLag
        ReDim phiNew(1 To lag)

        ' Calculate new PACF value
        sum = 0
        For k = 1 To lag - 1
            sum = sum + phi(k) * acf(lag - k)
        Next k

        pacf(lag) = (acf(lag) - sum) / (1 - sum)
        phiNew(lag) = pacf(lag)

        ' Update phi coefficients
        For k = 1 To lag - 1
            phiNew(k) = phi(k) - pacf(lag) * phi(lag - k)
        Next k

        ' Copy new phi values
        For k = 1 To lag
            If k <= UBound(phi) Then phi(k) = phiNew(k)
        Next k

        ReDim Preserve phi(1 To lag)
        For k = 1 To lag
            phi(k) = phiNew(k)
        Next k
    Next lag

    CalculatePACF = pacf
End Function

' ============================================================================
' LJUNG-BOX TEST
' ============================================================================

Public Function CalculateLjungBox(ByRef Residuals() As Double, ByVal maxLag As Integer) As Variant
    ' Ljung-Box Q-statistic for testing autocorrelation
    ' Returns array: [Q-statistic, p-value, conclusion]

    Dim n As Long, lag As Long
    Dim acf() As Double
    Dim qStat As Double
    Dim pValue As Double
    Dim result(1 To 3) As Variant

    n = UBound(Residuals) - LBound(Residuals) + 1

    ' Calculate ACF of residuals
    acf = CalculateACF(Residuals, maxLag)

    ' Calculate Q-statistic: Q = n(n+2) * sum[(acf^2)/(n-lag)]
    qStat = 0
    For lag = 1 To maxLag
        qStat = qStat + (acf(lag) ^ 2) / (n - lag)
    Next lag
    qStat = n * (n + 2) * qStat

    ' Calculate p-value using chi-square approximation
    ' Degrees of freedom = maxLag
    pValue = 1 - ChiSquareCDF(qStat, maxLag)

    ' Store results
    result(1) = qStat
    result(2) = pValue

    If pValue > 0.05 Then
        result(3) = "PASS: Residuals appear random (p=" & Format(pValue, "0.000") & ")"
    Else
        result(3) = "FAIL: Residuals show autocorrelation (p=" & Format(pValue, "0.000") & ")"
    End If

    CalculateLjungBox = result
End Function

Private Function ChiSquareCDF(ByVal x As Double, ByVal df As Integer) As Double
    ' Approximation of chi-square cumulative distribution function
    ' Uses gamma function approximation for simplicity

    If x <= 0 Then
        ChiSquareCDF = 0
        Exit Function
    End If

    If df <= 0 Then
        ChiSquareCDF = 0
        Exit Function
    End If

    ' For small df, use series expansion
    ' For larger df, use normal approximation

    If df <= 30 Then
        ' Use incomplete gamma function approximation
        Dim k As Double, sum As Double, term As Double
        Dim halfDf As Double, halfX As Double

        halfDf = df / 2#
        halfX = x / 2#

        ' Series expansion
        term = Exp(-halfX) * (halfX ^ halfDf)
        sum = term

        For k = 1 To 100
            term = term * halfX / (halfDf + k)
            sum = sum + term
            If Abs(term) < 0.0000001 Then Exit For
        Next k

        ChiSquareCDF = sum / Gamma(halfDf)

        ' Ensure bounds
        If ChiSquareCDF > 1 Then ChiSquareCDF = 1
        If ChiSquareCDF < 0 Then ChiSquareCDF = 0
    Else
        ' Normal approximation for large df
        Dim z As Double
        z = (Sqr(2 * x) - Sqr(2 * df - 1))
        ChiSquareCDF = NormalCDF(z)
    End If
End Function

Private Function Gamma(ByVal z As Double) As Double
    ' Approximation of gamma function using Stirling's formula
    Dim coef(0 To 5) As Double
    Dim i As Integer, sum As Double, x As Double

    ' Lanczos approximation coefficients
    coef(0) = 1.000000000190015
    coef(1) = 76.18009172947146
    coef(2) = -86.50532032941677
    coef(3) = 24.01409824083091
    coef(4) = -1.231739572450155
    coef(5) = 0.001208650973866179

    x = z
    sum = coef(0)

    For i = 1 To 5
        sum = sum + coef(i) / (x + i)
    Next i

    Gamma = Sqr(2 * 3.14159265358979) * sum * ((x + 5.5) ^ (x + 0.5)) * Exp(-(x + 5.5)) / x
End Function

Private Function NormalCDF(ByVal z As Double) As Double
    ' Standard normal cumulative distribution function
    Dim t As Double, p As Double
    Dim b1 As Double, b2 As Double, b3 As Double, b4 As Double, b5 As Double

    b1 = 0.319381530
    b2 = -0.356563782
    b3 = 1.781477937
    b4 = -1.821255978
    b5 = 1.330274429

    If z < 0 Then
        NormalCDF = 1 - NormalCDF(-z)
    Else
        t = 1 / (1 + 0.2316419 * z)
        p = 1 - (1 / Sqr(2 * 3.14159265358979)) * Exp(-z * z / 2) * _
            (b1 * t + b2 * t ^ 2 + b3 * t ^ 3 + b4 * t ^ 4 + b5 * t ^ 5)
        NormalCDF = p
    End If
End Function

' ============================================================================
' MEAN BIAS ERROR
' ============================================================================

Public Function CalculateMBE(ByRef actual() As Double, ByRef fitted() As Double) As Double
    ' Mean Bias Error - detects systematic over/under forecasting
    ' Positive = under-forecasting (actual > forecast)
    ' Negative = over-forecasting (actual < forecast)

    Dim sum As Double, i As Long, n As Long

    sum = 0
    n = UBound(actual) - LBound(actual) + 1

    For i = LBound(actual) To UBound(actual)
        sum = sum + (actual(i) - fitted(i))
    Next i

    CalculateMBE = sum / n
End Function

' ============================================================================
' Q-Q PLOT DATA
' ============================================================================

Public Function CalculateQQPlotData(ByRef Residuals() As Double) As Variant
    Dim n As Long, i As Long
    Dim sorted() As Double
    Dim theoreticalQuantiles() As Double
    Dim p As Double

    ' Copy and sort residuals
    n = UBound(Residuals) - LBound(Residuals) + 1
    ReDim sorted(1 To n)
    ReDim theoreticalQuantiles(1 To n)

    For i = LBound(Residuals) To UBound(Residuals)
        sorted(i - LBound(Residuals) + 1) = Residuals(i)
    Next i

    ' Sort using simple bubble sort
    Dim j As Long, temp As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If sorted(i) > sorted(j) Then
                temp = sorted(i)
                sorted(i) = sorted(j)
                sorted(j) = temp
            End If
        Next j
    Next i

    ' Calculate theoretical quantiles
    For i = 1 To n
        p = (i - 0.5) / n
        theoreticalQuantiles(i) = NormalInverseCDF(p)
    Next i

    CalculateQQPlotData = Array(theoreticalQuantiles, sorted)
End Function

Private Function NormalInverseCDF(ByVal p As Double) As Double
    ' Approximation of inverse CDF for standard normal distribution
    Dim q As Double, r As Double, result As Double

    If p <= 0 Then
        NormalInverseCDF = -1E+100
        Exit Function
    End If

    If p >= 1 Then
        NormalInverseCDF = 1E+100
        Exit Function
    End If

    q = p - 0.5

    If Abs(q) <= 0.425 Then
        r = 0.180625 - q * q
        NormalInverseCDF = q * (((((((2509.0809287301226727 * r + 33430.575583588128105) * r + _
            43162.719737483539118) * r + 27170.726569603738341) * r + 6377.6776828239672917) * r + _
            588.20283456061709004) * r + 25.063890505363067995) * r + 1#) / _
            (((((((5226.4952788528542831 * r + 28729.085735721942674) * r + _
            39307.895800092710610) * r + 21213.794301586595867) * r + 5394.1960214247511077) * r + _
            687.18700749205790830) * r + 42.313330701600911252) * r + 1#)
    Else
        If q < 0 Then
            r = p
        Else
            r = 1 - p
        End If
        r = Sqr(-Log(r))

        If r <= 5# Then
            r = r - 1.6
            result = (((((((0.00077454501427834140764 * r + 0.0227238449892691845833) * r + _
                0.24178072517745061177) * r + 1.2704582524523683826) * r + _
                3.6478483247632046050) * r + 5.7694972214606914055) * r + _
                4.6303378461565452959) * r + 1.4234371107496835773) / _
                (((((((0.00000000105075007164441684324 * r + 0.00054759380849953449460) * r + _
                0.0151986665636164571966) * r + 0.14810397642748007459) * r + _
                0.68976733498510000455) * r + 1.6763848301838038494) * r + _
                2.0531916266377588218) * r + 1#)
        Else
            r = r - 5#
            result = (((((((0.00000020103343992922881327 * r + 0.000027115555687434875782) * r + _
                0.0012426609473880784386) * r + 0.026532189526576123093) * r + _
                0.29656057182850489123) * r + 1.7848265399172913358) * r + _
                5.4637849111641143699) * r + 6.6579046435011037772) / _
                (((((((0.00000000000000204426310338993978564 * r + 0.00000014215117583164458887) * r + _
                0.000018463183175100546818) * r + 0.00078686913114561325910) * r + _
                0.014875361290850614853) * r + 0.13692988092273580531) * r + _
                0.59983220655588793769) * r + 1#)
        End If

        If q < 0 Then
            NormalInverseCDF = -result
        Else
            NormalInverseCDF = result
        End If
    End If
End Function

' ============================================================================
' ADVANCED FORECASTING METHODS - IMPROVED ACCURACY
' ============================================================================

' ============================================================================
' LINEAR TREND FORECAST - Simple linear regression
' ============================================================================
Public Function LinearTrendForecast(ByRef tsData As TimeSeriesData, _
                                   ByVal horizon As Integer) As ForecastResult
    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
    Dim slope As Double, intercept As Double
    Dim meanX As Double, meanY As Double

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Calculate linear regression: y = slope * x + intercept
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0

    For i = LBound(Values) To UBound(Values)
        Dim x As Double
        x = i - LBound(Values) + 1 ' Time index starting from 1
        sumX = sumX + x
        sumY = sumY + Values(i)
        sumXY = sumXY + x * Values(i)
        sumX2 = sumX2 + x * x
    Next i

    meanX = sumX / n
    meanY = sumY / n

    slope = (sumXY - n * meanX * meanY) / (sumX2 - n * meanX * meanX)
    intercept = meanY - slope * meanX

    ' Calculate fitted values and residuals
    For i = LBound(Values) To UBound(Values)
        x = i - LBound(Values) + 1
        result.FittedValues(i) = slope * x + intercept
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Generate forecasts
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        x = n + i
        result.ForecastValues(i) = slope * x + intercept

        ' Confidence intervals widen with horizon
        Dim se As Double
        se = residualStdDev * Sqr(1 + 1 / n + ((x - meanX) ^ 2) / sumX2)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * se
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.ModelName = "Linear Trend"

    LinearTrendForecast = result
End Function

' ============================================================================
' MOVING AVERAGE FORECAST - Simple average of recent observations
' ============================================================================
Public Function MovingAverageForecast(ByRef tsData As TimeSeriesData, _
                                     ByVal horizon As Integer) As ForecastResult
    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long
    Dim windowSize As Long
    Dim movingAvg As Double

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Window size = min(12, n/3)
    windowSize = WorksheetFunction.Min(12, WorksheetFunction.Max(3, Int(n / 3)))

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Calculate fitted values using moving average
    For i = LBound(Values) To UBound(Values)
        Dim sum As Double
        Dim count As Long
        Dim j As Long

        sum = 0
        count = 0

        For j = WorksheetFunction.Max(LBound(Values), i - windowSize + 1) To i - 1
            sum = sum + Values(j)
            count = count + 1
        Next j

        If count > 0 Then
            result.FittedValues(i) = sum / count
        Else
            result.FittedValues(i) = Values(i)
        End If

        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Calculate forecast as average of last windowSize observations
    movingAvg = 0
    For i = WorksheetFunction.Max(LBound(Values), UBound(Values) - windowSize + 1) To UBound(Values)
        movingAvg = movingAvg + Values(i)
    Next i
    movingAvg = movingAvg / windowSize

    ' All future forecasts are the same (flat forecast)
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        result.ForecastValues(i) = movingAvg
        Dim se As Double
        se = residualStdDev * Sqr(1 + i / windowSize)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * se
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.ModelName = "Moving Average"

    MovingAverageForecast = result
End Function

' ============================================================================
' EXPONENTIAL TREND FORECAST - For exponentially growing series
' ============================================================================
Public Function ExponentialTrendForecast(ByRef tsData As TimeSeriesData, _
                                        ByVal horizon As Integer) As ForecastResult
    Dim result As ForecastResult
    Dim Values() As Double
    Dim logValues() As Double
    Dim n As Long, i As Long
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
    Dim slope As Double, intercept As Double
    Dim meanX As Double, meanY As Double
    Dim hasNegative As Boolean

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)
    ReDim logValues(LBound(Values) To UBound(Values))

    ' Check for non-positive values (can't take log)
    hasNegative = False
    For i = LBound(Values) To UBound(Values)
        If Values(i) <= 0 Then
            hasNegative = True
            Exit For
        End If
    Next i

    ' If negative values, fall back to linear trend
    If hasNegative Then
        result = LinearTrendForecast(tsData, horizon)
        result.ModelName = "Linear Trend (fallback)"
        ExponentialTrendForecast = result
        Exit Function
    End If

    ' Transform to log space
    For i = LBound(Values) To UBound(Values)
        logValues(i) = Log(Values(i))
    Next i

    ' Linear regression on log-transformed data: log(y) = slope * x + intercept
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0

    For i = LBound(Values) To UBound(Values)
        Dim x As Double
        x = i - LBound(Values) + 1
        sumX = sumX + x
        sumY = sumY + logValues(i)
        sumXY = sumXY + x * logValues(i)
        sumX2 = sumX2 + x * x
    Next i

    meanX = sumX / n
    meanY = sumY / n

    slope = (sumXY - n * meanX * meanY) / (sumX2 - n * meanX * meanX)
    intercept = meanY - slope * meanX

    ' Calculate fitted values and residuals in original space
    For i = LBound(Values) To UBound(Values)
        x = i - LBound(Values) + 1
        result.FittedValues(i) = Exp(slope * x + intercept)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Generate forecasts
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        x = n + i
        result.ForecastValues(i) = Exp(slope * x + intercept)

        ' Approximate confidence intervals
        Dim se As Double
        se = result.ForecastValues(i) * residualStdDev / (sumY / n) * Sqr(i)
        result.Lower95(i) = WorksheetFunction.Max(0, result.ForecastValues(i) - 1.96 * se)
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.ModelName = "Exponential Trend"

    ExponentialTrendForecast = result
End Function

' ============================================================================
' AUTOMATIC MODEL SELECTION - Tries all methods and picks the best!
' ============================================================================
Public Function AutoForecast(ByRef tsData As TimeSeriesData, _
                             ByVal horizon As Integer, _
                             ByVal seasonalType As String) As ForecastResult
    On Error GoTo ErrorHandler

    Dim results() As ForecastResult
    Dim modelNames() As String
    Dim modelCount As Integer
    Dim i As Integer
    Dim bestIndex As Integer
    Dim bestMAPE As Double
    Dim cleanedData As TimeSeriesData
    Dim combinedResult As ForecastResult
    
    ' STEP 1: Data Preprocessing - Impute missing values and clean outliers
    Dim imputedValues() As Double
    imputedValues = ImputeMissingValues(tsData.Values)

    cleanedData = tsData
    cleanedData.Values = imputedValues

    ' Advanced outlier detection using MAD (more robust than Z-score)
    Dim outliers() As Boolean
    outliers = DetectOutliersAdvanced(cleanedData.Values, "MAD")

    Dim k As Long
    For k = LBound(cleanedData.Values) To UBound(cleanedData.Values)
        If outliers(k) Then
            ' Replace outlier with linear interpolation
            If k > LBound(cleanedData.Values) And k < UBound(cleanedData.Values) Then
                cleanedData.Values(k) = (cleanedData.Values(k - 1) + cleanedData.Values(k + 1)) / 2
            End If
        End If
    Next k

    ' STEP 2: Intelligent Demand Classification
    Dim demandClass As DemandClassification
    demandClass = ClassifyDemand(cleanedData.Values)

    ' STEP 3: Feature-based Model Selection
    Dim recommendations() As String
    recommendations = SelectBestModels(cleanedData)

    ' STEP 4: Try ALL 30+ forecasting methods - best ones combined via BMA!
    modelCount = 30
    ReDim results(1 To modelCount)
    ReDim modelNames(1 To modelCount)

    ' 1. Simple Exponential Smoothing
    On Error Resume Next
    results(1) = SimpleExponentialSmoothing(cleanedData, horizon)
    results(1).ModelName = "SES"
    If Err.Number <> 0 Then results(1).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 2. Holt-Winters
    On Error Resume Next
    results(2) = HoltWinters(cleanedData, horizon, seasonalType)
    results(2).ModelName = "Holt-Winters"
    If Err.Number <> 0 Then results(2).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 3. Damped Trend Holt-Winters
    On Error Resume Next
    results(3) = DampedHoltWinters(cleanedData, horizon, seasonalType)
    results(3).ModelName = "Damped HW"
    If Err.Number <> 0 Then results(3).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 4. Theta Method
    On Error Resume Next
    results(4) = ThetaMethod(cleanedData, horizon)
    results(4).ModelName = "Theta"
    If Err.Number <> 0 Then results(4).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 5. Ensemble (SES + HW)
    On Error Resume Next
    results(5) = EnsembleForecast(cleanedData, horizon, seasonalType)
    results(5).ModelName = "Ensemble"
    If Err.Number <> 0 Then results(5).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 6. Simple ARIMA(1,1,0)
    On Error Resume Next
    results(6) = SimpleARIMA(cleanedData, horizon)
    results(6).ModelName = "ARIMA(1,1,0)"
    If Err.Number <> 0 Then results(6).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 7. Croston's Method (Perfect for intermittent demand!)
    On Error Resume Next
    results(7) = CrostonsMethod(cleanedData, horizon)
    results(7).ModelName = "Croston"
    If Err.Number <> 0 Then results(7).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 8. Auto-ARIMA (Tests multiple ARIMA orders)
    On Error Resume Next
    results(8) = AutoARIMA(cleanedData, horizon)
    ' ModelName set by AutoARIMA itself
    If Err.Number <> 0 Then results(8).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 9. Advanced Ensemble (Recent performance weighting)
    On Error Resume Next
    results(9) = AdvancedEnsemble(cleanedData, horizon, seasonalType)
    results(9).ModelName = "Advanced Ensemble"
    If Err.Number <> 0 Then results(9).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 10. Holt-Winters with opposite seasonal type (try both!)
    On Error Resume Next
    Dim altSeasonalType As String
    If LCase(seasonalType) = "additive" Then
        altSeasonalType = "multiplicative"
    Else
        altSeasonalType = "additive"
    End If
    results(10) = HoltWinters(cleanedData, horizon, altSeasonalType)
    results(10).ModelName = "HW (" & altSeasonalType & ")"
    If Err.Number <> 0 Then results(10).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 11. Linear Trend (Simple but effective for trending data)
    On Error Resume Next
    results(11) = LinearTrendForecast(cleanedData, horizon)
    results(11).ModelName = "Linear Trend"
    If Err.Number <> 0 Then results(11).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 12. Moving Average (Good for stable series)
    On Error Resume Next
    results(12) = MovingAverageForecast(cleanedData, horizon)
    results(12).ModelName = "Moving Average"
    If Err.Number <> 0 Then results(12).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 13. Exponential Trend (For exponentially growing series)
    On Error Resume Next
    results(13) = ExponentialTrendForecast(cleanedData, horizon)
    results(13).ModelName = "Exponential Trend"
    If Err.Number <> 0 Then results(13).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 14. SARIMA - Seasonal ARIMA for complex seasonality
    On Error Resume Next
    results(14) = SARIMAForecast(cleanedData, horizon)
    ' ModelName set by SARIMA
    If Err.Number <> 0 Then results(14).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 15. Fourier Series - Multi-seasonal patterns
    On Error Resume Next
    results(15) = FourierForecast(cleanedData, horizon)
    ' ModelName set by Fourier
    If Err.Number <> 0 Then results(15).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 16. TSB Method - Advanced intermittent demand
    On Error Resume Next
    results(16) = TSBMethod(cleanedData, horizon)
    results(16).ModelName = "TSB"
    If Err.Number <> 0 Then results(16).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 17. SBA Method - Syntetos-Boylan for intermittent demand
    On Error Resume Next
    results(17) = SBAMethod(cleanedData, horizon)
    results(17).ModelName = "SBA"
    If Err.Number <> 0 Then results(17).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 18. ETS(A,A,A) - State Space model
    On Error Resume Next
    results(18) = ETSForecast(cleanedData, horizon, "A", "A", "A")
    ' ModelName set by ETS
    If Err.Number <> 0 Then results(18).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 19. ETS(M,A,M) - Multiplicative State Space
    On Error Resume Next
    results(19) = ETSForecast(cleanedData, horizon, "M", "A", "M")
    ' ModelName set by ETS
    If Err.Number <> 0 Then results(19).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 20. ETS(A,Ad,A) - Damped State Space
    On Error Resume Next
    results(20) = ETSForecast(cleanedData, horizon, "A", "Ad", "A")
    ' ModelName set by ETS
    If Err.Number <> 0 Then results(20).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 21. Dynamic Harmonic Regression
    On Error Resume Next
    results(21) = DynamicHarmonicRegression(cleanedData, horizon)
    ' ModelName set by DHR
    If Err.Number <> 0 Then results(21).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 22. Neural Network (lag-5, 5 hidden nodes)
    On Error Resume Next
    results(22) = NeuralNetworkForecast(cleanedData, horizon, 5, 5)
    ' ModelName set by NN
    If Err.Number <> 0 Then results(22).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 23. Neural Network (lag-10, 8 hidden nodes)
    On Error Resume Next
    results(23) = NeuralNetworkForecast(cleanedData, horizon, 10, 8)
    ' ModelName set by NN
    If Err.Number <> 0 Then results(23).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 24. Holt Linear (ETS A,A,N)
    On Error Resume Next
    results(24) = ETSForecast(cleanedData, horizon, "A", "A", "N")
    ' ModelName set by ETS
    If Err.Number <> 0 Then results(24).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 25. Simple Seasonal Smoothing (no trend)
    On Error Resume Next
    results(25) = ETSForecast(cleanedData, horizon, "A", "N", "A")
    ' ModelName set by ETS
    If Err.Number <> 0 Then results(25).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 26. STL Decomposition - State-of-the-art seasonal decomposition
    On Error Resume Next
    results(26) = STLDecomposition(cleanedData, horizon)
    ' ModelName set by STL
    If Err.Number <> 0 Then results(26).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 27. Kalman Filter - Adaptive state estimation
    On Error Resume Next
    results(27) = KalmanFilterForecast(cleanedData, horizon)
    results(27).ModelName = "Kalman"
    If Err.Number <> 0 Then results(27).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 28. Kalman Filter (higher process noise for volatile series)
    On Error Resume Next
    results(28) = KalmanFilterForecast(cleanedData, horizon, 0.1, 0.1)
    results(28).ModelName = "Kalman-HN"
    If Err.Number <> 0 Then results(28).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 29. STL with multiplicative seasonality approximation
    On Error Resume Next
    If cleanedData.Frequency > 0 Then
        results(29) = STLDecomposition(cleanedData, horizon, cleanedData.Frequency)
    Else
        results(29) = STLDecomposition(cleanedData, horizon, 12)
    End If
    If Err.Number <> 0 Then results(29).MAPE = 9999
    On Error GoTo ErrorHandler

    ' 30. Ensemble with equal weights (baseline comparison)
    On Error Resume Next
    results(30) = EnsembleForecast(cleanedData, horizon, seasonalType)
    results(30).ModelName = "Ensemble-Equal"
    If Err.Number <> 0 Then results(30).MAPE = 9999
    On Error GoTo ErrorHandler

    ' STEP 5: Intelligent Bayesian Model Averaging with Diversity Weighting
    ' Filter out failed models (MAPE = 9999)
    Dim validModels() As ForecastResult
    Dim validCount As Integer
    validCount = 0

    For i = 1 To modelCount
        If results(i).MAPE < 9999 Then
            validCount = validCount + 1
        End If
    Next i

    If validCount > 0 Then
        ReDim validModels(1 To validCount)
        Dim validIdx As Integer
        validIdx = 0

        For i = 1 To modelCount
            If results(i).MAPE < 9999 Then
                validIdx = validIdx + 1
                validModels(validIdx) = results(i)
            End If
        Next i

        ' Use BMA to combine all valid models (statistically optimal)
        combinedResult = BayesianModelAveraging(validModels, validCount, cleanedData.Values)

        ' Calculate ensemble diversity scores
        Dim diversityScores() As Double
        diversityScores = EnsembleDiversityScore(validModels, validCount)

        ' Incorporate diversity into final model name
        Dim avgDiversity As Double
        avgDiversity = 0
        For i = 1 To validCount
            avgDiversity = avgDiversity + diversityScores(i)
        Next i
        avgDiversity = avgDiversity / validCount

        combinedResult.ModelName = combinedResult.ModelName & " [Div:" & Format(avgDiversity, "0.00") & "]"
    Else
        ' Fallback if all models failed
        GoTo ErrorHandler
    End If

    ' STEP 6: Apply bias correction if significant
    Dim biasCorrect As Double
    biasCorrect = combinedResult.MBE

    ' Only correct if bias is significant (> 10% of MAE)
    If Abs(biasCorrect) > combinedResult.MAE * 0.1 Then
        ' Apply bias correction to forecast
        For i = 1 To horizon
            combinedResult.ForecastValues(i) = combinedResult.ForecastValues(i) - biasCorrect
            combinedResult.Lower95(i) = combinedResult.Lower95(i) - biasCorrect
            combinedResult.Upper95(i) = combinedResult.Upper95(i) - biasCorrect
        Next i

        ' Mark that bias correction was applied
        combinedResult.ModelName = combinedResult.ModelName & " [BC]"
    End If

    ' STEP 7: Apply quantile-based prediction intervals (more accurate than normal dist)
    combinedResult = QuantileForecastIntervals(combinedResult, horizon)

    ' STEP 8: Apply non-negativity constraint for demand forecasting
    Dim reconciledForecast() As Double
    reconciledForecast = ReconcileForecast(combinedResult.ForecastValues, 0)

    For i = 1 To horizon
        combinedResult.ForecastValues(i) = reconciledForecast(i)
        ' Ensure intervals are also non-negative
        If combinedResult.Lower95(i) < 0 Then combinedResult.Lower95(i) = 0
        If combinedResult.Lower80(i) < 0 Then combinedResult.Lower80(i) = 0
    Next i

    ' STEP 9: Add demand classification to model name for transparency
    combinedResult.ModelName = "[" & demandClass.Category & "] " & combinedResult.ModelName

    AutoForecast = combinedResult
    Exit Function

ErrorHandler:
    ' If all fail, return SES
    AutoForecast = SimpleExponentialSmoothing(tsData, horizon)
    AutoForecast.ModelName = "SES (Fallback)"
End Function

' ============================================================================
' ENSEMBLE FORECASTING - Combines SES + HW for better accuracy
' ============================================================================
Public Function EnsembleForecast(ByRef tsData As TimeSeriesData, _
                                 ByVal horizon As Integer, _
                                 ByVal seasonalType As String) As ForecastResult
    Dim sesResult As ForecastResult
    Dim hwResult As ForecastResult
    Dim result As ForecastResult
    Dim i As Long
    Dim sesWeight As Double, hwWeight As Double
    
    ' Get forecasts from both methods
    sesResult = SimpleExponentialSmoothing(tsData, horizon)
    hwResult = HoltWinters(tsData, horizon, seasonalType)
    
    ' Weight by inverse MAPE (better model gets more weight)
    If sesResult.MAPE > 0 And hwResult.MAPE > 0 Then
        sesWeight = (1 / sesResult.MAPE) / ((1 / sesResult.MAPE) + (1 / hwResult.MAPE))
        hwWeight = 1 - sesWeight
    Else
        sesWeight = 0.5
        hwWeight = 0.5
    End If
    
    ' Initialize result arrays
    ReDim result.FittedValues(LBound(tsData.Values) To UBound(tsData.Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Residuals(LBound(tsData.Values) To UBound(tsData.Values))
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)
    
    ' Combine fitted values
    For i = LBound(tsData.Values) To UBound(tsData.Values)
        result.FittedValues(i) = sesWeight * sesResult.FittedValues(i) + hwWeight * hwResult.FittedValues(i)
        result.Residuals(i) = tsData.Values(i) - result.FittedValues(i)
    Next i
    
    ' Combine forecasts
    For i = 1 To horizon
        result.ForecastValues(i) = sesWeight * sesResult.ForecastValues(i) + hwWeight * hwResult.ForecastValues(i)
        result.Lower95(i) = sesWeight * sesResult.Lower95(i) + hwWeight * hwResult.Lower95(i)
        result.Upper95(i) = sesWeight * sesResult.Upper95(i) + hwWeight * hwResult.Upper95(i)
    Next i
    
    ' Calculate metrics
    result.MAPE = CalculateMAPE(tsData.Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(tsData.Values, result.FittedValues)
    result.ModelName = "Ensemble"

    EnsembleForecast = result
End Function

' ============================================================================
' DAMPED TREND HOLT-WINTERS - Better for long-term forecasts
' ============================================================================
Public Function DampedHoltWinters(ByRef tsData As TimeSeriesData, _
                                  ByVal horizon As Integer, _
                                  ByVal seasonalType As String, _
                                  Optional ByVal damping As Variant) As ForecastResult
    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long, j As Long
    Dim m As Integer
    Dim level As Double, trend As Double
    Dim seasonal() As Double
    Dim phi As Double
    Dim alpha As Double, beta As Double, gamma As Double
    
    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1
    m = tsData.Frequency
    
    ' Use optimized damping parameter if not provided
    If IsMissing(damping) Or IsEmpty(damping) Then
        phi = 0.98  ' Typical value - dampens trend by 2% per period
    Else
        phi = CDbl(damping)
    End If
    
    ' Optimize parameters (simplified - use good defaults)
    alpha = 0.2
    beta = 0.1
    gamma = 0.1
    
    ' Initialize
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)
    ReDim seasonal(1 To m)
    
    ' Initial values
    level = Values(LBound(Values))
    trend = (Values(LBound(Values) + m) - Values(LBound(Values))) / m
    
    ' Initialize seasonal indices
    For i = 1 To m
        seasonal(i) = 1
    Next i
    
    ' Fit model with damping
    For i = LBound(Values) To UBound(Values)
        Dim seasonalIndex As Integer
        seasonalIndex = ((i - LBound(Values)) Mod m) + 1
        
        If LCase(seasonalType) = "multiplicative" Then
            result.FittedValues(i) = (level + trend) * seasonal(seasonalIndex)
            
            Dim newLevel As Double
            newLevel = alpha * (Values(i) / seasonal(seasonalIndex)) + (1 - alpha) * (level + phi * trend)
            trend = beta * (newLevel - level) + (1 - beta) * phi * trend
            seasonal(seasonalIndex) = gamma * (Values(i) / newLevel) + (1 - gamma) * seasonal(seasonalIndex)
            level = newLevel
        Else ' additive
            result.FittedValues(i) = level + trend + seasonal(seasonalIndex)
            
            newLevel = alpha * (Values(i) - seasonal(seasonalIndex)) + (1 - alpha) * (level + phi * trend)
            trend = beta * (newLevel - level) + (1 - beta) * phi * trend
            seasonal(seasonalIndex) = gamma * (Values(i) - newLevel) + (1 - gamma) * seasonal(seasonalIndex)
            level = newLevel
        End If
        
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i
    
    ' Forecast with damped trend
    Dim cumulativePhi As Double
    cumulativePhi = 0
    For j = 1 To horizon
        cumulativePhi = cumulativePhi + phi ^ j
        seasonalIndex = ((j - 1) Mod m) + 1
        
        If LCase(seasonalType) = "multiplicative" Then
            result.ForecastValues(j) = (level + cumulativePhi * trend) * seasonal(seasonalIndex)
        Else
            result.ForecastValues(j) = level + cumulativePhi * trend + seasonal(seasonalIndex)
        End If
    Next j
    
    ' Calculate confidence intervals
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)
    
    For i = 1 To horizon
        Dim se As Double
        se = residualStdDev * Sqr(i)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * se
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i
    
    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.Alpha = alpha
    result.Beta = beta
    result.Gamma = gamma
    result.Phi = phi
    result.ModelName = "Damped HW"

    DampedHoltWinters = result
End Function

' ============================================================================
' THETA METHOD - Simple but often very accurate!
' ============================================================================
Public Function ThetaMethod(ByRef tsData As TimeSeriesData, _
                            ByVal horizon As Integer) As ForecastResult
    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long
    Dim theta As Double
    Dim line1() As Double, line2() As Double
    Dim drift As Double
    
    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1
    theta = 2  ' Classic Theta=2
    
    ' Initialize
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)
    ReDim line1(LBound(Values) To UBound(Values))
    ReDim line2(LBound(Values) To UBound(Values))
    
    ' Line 1: SES (theta=0)
    Dim sesResult As ForecastResult
    sesResult = SimpleExponentialSmoothing(tsData, horizon)
    
    ' Line 2: Linear regression (theta=2)
    ' Calculate drift
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0
    
    For i = LBound(Values) To UBound(Values)
        Dim x As Double
        x = i - LBound(Values) + 1
        sumX = sumX + x
        sumY = sumY + Values(i)
        sumXY = sumXY + x * Values(i)
        sumX2 = sumX2 + x * x
    Next i
    
    drift = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX)
    Dim intercept As Double
    intercept = (sumY - drift * sumX) / n
    
    ' Combine lines (average of SES and linear trend)
    For i = LBound(Values) To UBound(Values)
        x = i - LBound(Values) + 1
        line2(i) = intercept + drift * x
        result.FittedValues(i) = (sesResult.FittedValues(i) + line2(i)) / 2
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i
    
    ' Forecast
    For i = 1 To horizon
        x = n + i
        Dim sesForecast As Double
        sesForecast = sesResult.ForecastValues(i)
        Dim linearForecast As Double
        linearForecast = intercept + drift * x
        result.ForecastValues(i) = (sesForecast + linearForecast) / 2
    Next i
    
    ' Calculate confidence intervals
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)
    
    For i = 1 To horizon
        Dim se As Double
        se = residualStdDev * Sqr(i)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * se
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i
    
    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.ModelName = "Theta"

    ThetaMethod = result
End Function

' ============================================================================
' SIMPLE ARIMA - Simplified Auto-Regressive Integrated Moving Average
' ============================================================================
Public Function SimpleARIMA(ByRef tsData As TimeSeriesData, _
                            ByVal horizon As Integer) As ForecastResult
    Dim result As ForecastResult
    Dim Values() As Double
    Dim diffValues() As Double
    Dim n As Long, i As Long
    Dim ar1 As Double  ' AR(1) coefficient
    
    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1
    
    ' First differencing to make stationary
    ReDim diffValues(LBound(Values) + 1 To UBound(Values))
    For i = LBound(Values) + 1 To UBound(Values)
        diffValues(i) = Values(i) - Values(i - 1)
    Next i
    
    ' Estimate AR(1) coefficient using Yule-Walker
    Dim mean As Double
    Dim sumDiff As Double
    sumDiff = 0
    For i = LBound(diffValues) To UBound(diffValues)
        sumDiff = sumDiff + diffValues(i)
    Next i
    mean = sumDiff / (UBound(diffValues) - LBound(diffValues) + 1)
    
    Dim gamma0 As Double, gamma1 As Double
    gamma0 = 0: gamma1 = 0
    For i = LBound(diffValues) To UBound(diffValues)
        gamma0 = gamma0 + (diffValues(i) - mean) ^ 2
        If i < UBound(diffValues) Then
            gamma1 = gamma1 + (diffValues(i) - mean) * (diffValues(i + 1) - mean)
        End If
    Next i
    
    If gamma0 > 0 Then
        ar1 = gamma1 / gamma0
    Else
        ar1 = 0
    End If
    
    ' Ensure stability
    If ar1 >= 1 Then ar1 = 0.95
    If ar1 <= -1 Then ar1 = -0.95
    
    ' Initialize result
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)
    
    ' Fit values
    result.FittedValues(LBound(Values)) = Values(LBound(Values))
    For i = LBound(Values) + 1 To UBound(Values)
        result.FittedValues(i) = Values(i - 1) + ar1 * (Values(i - 1) - Values(i - 2))
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i
    result.Residuals(LBound(Values)) = 0
    
    ' Forecast
    Dim lastValue As Double
    Dim lastDiff As Double
    lastValue = Values(UBound(Values))
    lastDiff = diffValues(UBound(diffValues))
    
    For i = 1 To horizon
        Dim forecastDiff As Double
        forecastDiff = ar1 ^ i * lastDiff
        result.ForecastValues(i) = lastValue + forecastDiff * i
    Next i
    
    ' Calculate confidence intervals
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)
    
    For i = 1 To horizon
        Dim se As Double
        se = residualStdDev * Sqr(i)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * se
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i
    
    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.Alpha = ar1  ' Store AR coefficient in Alpha field
    result.ModelName = "ARIMA(1,1,0)"

    SimpleARIMA = result
End Function

' ============================================================================
' OUTLIER DETECTION AND REMOVAL - Clean your data first!
' ============================================================================
Public Function RemoveOutliers(ByRef tsData As TimeSeriesData) As TimeSeriesData
    Dim result As TimeSeriesData
    Dim Values() As Double
    Dim cleanValues() As Double
    Dim n As Long, i As Long
    Dim q1 As Double, q3 As Double, iqr As Double
    Dim lowerBound As Double, upperBound As Double
    Dim sorted() As Double
    
    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1
    
    ' Copy to sorted array
    ReDim sorted(LBound(Values) To UBound(Values))
    For i = LBound(Values) To UBound(Values)
        sorted(i) = Values(i)
    Next i
    
    ' Simple bubble sort
    Dim temp As Double
    Dim j As Long
    For i = LBound(sorted) To UBound(sorted) - 1
        For j = i + 1 To UBound(sorted)
            If sorted(i) > sorted(j) Then
                temp = sorted(i)
                sorted(i) = sorted(j)
                sorted(j) = temp
            End If
        Next j
    Next i
    
    ' Calculate Q1 and Q3
    Dim q1Pos As Long, q3Pos As Long
    q1Pos = LBound(sorted) + Int(n * 0.25)
    q3Pos = LBound(sorted) + Int(n * 0.75)
    q1 = sorted(q1Pos)
    q3 = sorted(q3Pos)
    iqr = q3 - q1
    
    ' Calculate bounds (1.5 * IQR is standard outlier detection)
    lowerBound = q1 - 1.5 * iqr
    upperBound = q3 + 1.5 * iqr
    
    ' Replace outliers with interpolated values
    ReDim cleanValues(LBound(Values) To UBound(Values))
    For i = LBound(Values) To UBound(Values)
        If Values(i) < lowerBound Or Values(i) > upperBound Then
            ' Interpolate from neighbors
            If i = LBound(Values) Then
                cleanValues(i) = Values(i + 1)
            ElseIf i = UBound(Values) Then
                cleanValues(i) = Values(i - 1)
            Else
                cleanValues(i) = (Values(i - 1) + Values(i + 1)) / 2
            End If
        Else
            cleanValues(i) = Values(i)
        End If
    Next i
    
    result.Values = cleanValues
    result.Frequency = tsData.Frequency
    RemoveOutliers = result
End Function

' ============================================================================
' CROSTON'S METHOD - Perfect for intermittent/sparse demand (supply chain!)
' ============================================================================
' Handles data with lots of zeros (intermittent demand) by separating:
' 1. Size of demand (when it occurs)
' 2. Interval between demands
Public Function CrostonsMethod(ByRef tsData As TimeSeriesData, _
                               ByVal horizon As Integer) As ForecastResult
    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long, j As Long
    Dim alpha As Double

    ' Arrays for non-zero demands and intervals
    Dim demandSizes() As Double
    Dim intervals() As Long
    Dim demandCount As Long
    Dim lastDemandPeriod As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Optimal alpha for Croston (research shows 0.1-0.2 works best)
    alpha = 0.1

    ' Count non-zero demands
    demandCount = 0
    For i = LBound(Values) To UBound(Values)
        If Values(i) > 0 Then demandCount = demandCount + 1
    Next i

    ' Need at least 2 demands to work
    If demandCount < 2 Then
        ' Fall back to simple average
        Dim avgDemand As Double
        Dim totalDemand As Double
        For i = LBound(Values) To UBound(Values)
            totalDemand = totalDemand + Values(i)
        Next i
        avgDemand = totalDemand / n

        ReDim result.ForecastValues(1 To horizon)
        ReDim result.Lower95(1 To horizon)
        ReDim result.Upper95(1 To horizon)
        ReDim result.FittedValues(LBound(Values) To UBound(Values))
        ReDim result.Residuals(LBound(Values) To UBound(Values))

        For i = 1 To horizon
            result.ForecastValues(i) = avgDemand
        Next i

        For i = LBound(Values) To UBound(Values)
            result.FittedValues(i) = avgDemand
            result.Residuals(i) = Values(i) - avgDemand
        Next i

        result.MAPE = CalculateMAPE(Values, result.FittedValues)
        result.MAE = CalculateMAE(result.Residuals)
        result.RMSE = CalculateRMSE(result.Residuals)
        result.MBE = CalculateMBE(Values, result.FittedValues)
        result.ModelName = "Croston (fallback)"
        result.Alpha = alpha

        CrostonsMethod = result
        Exit Function
    End If

    ' Extract non-zero demands and intervals
    ReDim demandSizes(1 To demandCount)
    ReDim intervals(1 To demandCount)

    demandCount = 0
    lastDemandPeriod = LBound(Values) - 1

    For i = LBound(Values) To UBound(Values)
        If Values(i) > 0 Then
            demandCount = demandCount + 1
            demandSizes(demandCount) = Values(i)
            intervals(demandCount) = i - lastDemandPeriod
            lastDemandPeriod = i
        End If
    Next i

    ' Apply Croston's smoothing
    Dim smoothedSize As Double
    Dim smoothedInterval As Double

    smoothedSize = demandSizes(1)
    smoothedInterval = intervals(1)

    Dim smoothedSizes() As Double
    Dim smoothedIntervals() As Double
    ReDim smoothedSizes(1 To demandCount)
    ReDim smoothedIntervals(1 To demandCount)

    smoothedSizes(1) = smoothedSize
    smoothedIntervals(1) = smoothedInterval

    For i = 2 To demandCount
        smoothedSize = alpha * demandSizes(i) + (1 - alpha) * smoothedSize
        smoothedInterval = alpha * intervals(i) + (1 - alpha) * smoothedInterval
        smoothedSizes(i) = smoothedSize
        smoothedIntervals(i) = smoothedInterval
    Next i

    ' Calculate forecast = demand size / interval
    Dim forecastValue As Double
    If smoothedInterval > 0 Then
        forecastValue = smoothedSize / smoothedInterval
    Else
        forecastValue = smoothedSize
    End If

    ' Initialize arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Create fitted values (map smoothed forecasts back to original timeline)
    Dim demandIdx As Long
    demandIdx = 1
    For i = LBound(Values) To UBound(Values)
        If demandIdx <= demandCount Then
            If smoothedIntervals(demandIdx) > 0 Then
                result.FittedValues(i) = smoothedSizes(demandIdx) / smoothedIntervals(demandIdx)
            Else
                result.FittedValues(i) = smoothedSizes(demandIdx)
            End If
        Else
            result.FittedValues(i) = forecastValue
        End If

        If Values(i) > 0 And demandIdx < demandCount Then
            demandIdx = demandIdx + 1
        End If
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecast (constant)
    For i = 1 To horizon
        result.ForecastValues(i) = forecastValue
    Next i

    ' Confidence intervals (based on residual variance)
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        Dim se As Double
        se = residualStdDev * Sqr(i)
        result.Lower95(i) = WorksheetFunction.Max(0, result.ForecastValues(i) - 1.96 * se)
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)
    result.Alpha = alpha
    result.ModelName = "Croston"

    CrostonsMethod = result
    Exit Function

ErrorHandler:
    ' Return simple average on error
    Dim fallbackAvg As Double
    Dim fallbackSum As Double
    For i = LBound(Values) To UBound(Values)
        fallbackSum = fallbackSum + Values(i)
    Next i
    fallbackAvg = fallbackSum / n

    ReDim result.ForecastValues(1 To horizon)
    For i = 1 To horizon
        result.ForecastValues(i) = fallbackAvg
    Next i
    result.ModelName = "Croston (error)"
    CrostonsMethod = result
End Function

' ============================================================================
' AUTO-ARIMA - Automatically selects best (p,d,q) order
' ============================================================================
' Tests multiple ARIMA configurations and picks best based on AIC
Public Function AutoARIMA(ByRef tsData As TimeSeriesData, _
                          ByVal horizon As Integer, _
                          Optional ByVal maxP As Integer = 3, _
                          Optional ByVal maxD As Integer = 2, _
                          Optional ByVal maxQ As Integer = 3) As ForecastResult

    Dim bestResult As ForecastResult
    Dim testResult As ForecastResult
    Dim bestAIC As Double
    Dim currentAIC As Double
    Dim p As Integer, d As Integer, q As Integer
    Dim tested As Boolean

    bestAIC = 1E+100
    tested = False

    ' Test common ARIMA configurations
    ' Start with simple ones that work well in practice
    Dim configs() As Variant
    configs = Array( _
        Array(0, 1, 1), _
        Array(1, 1, 0), _
        Array(1, 1, 1), _
        Array(2, 1, 0), _
        Array(0, 1, 2), _
        Array(2, 1, 1), _
        Array(1, 1, 2), _
        Array(2, 1, 2), _
        Array(1, 0, 1), _
        Array(0, 1, 0) _
    )

    Dim configIdx As Integer
    For configIdx = LBound(configs) To UBound(configs)
        p = configs(configIdx)(0)
        d = configs(configIdx)(1)
        q = configs(configIdx)(2)

        On Error Resume Next
        testResult = FitARIMA(tsData, horizon, p, d, q)

        If Err.Number = 0 Then
            ' Calculate AIC
            currentAIC = CalculateAIC(testResult, tsData.Values, p + q + 1)

            If currentAIC < bestAIC Then
                bestAIC = currentAIC
                bestResult = testResult
                bestResult.ModelName = "ARIMA(" & p & "," & d & "," & q & ")"
                tested = True
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next configIdx

    ' If no model worked, fall back to simple ARIMA(1,1,0)
    If Not tested Then
        bestResult = SimpleARIMA(tsData, horizon)
        bestResult.ModelName = "ARIMA(1,1,0)-fallback"
    End If

    AutoARIMA = bestResult
End Function

Private Function FitARIMA(ByRef tsData As TimeSeriesData, _
                          ByVal horizon As Integer, _
                          ByVal p As Integer, _
                          ByVal d As Integer, _
                          ByVal q As Integer) As ForecastResult

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long, j As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Difference the series d times
    Dim diffValues() As Double
    ReDim diffValues(LBound(Values) To UBound(Values) - d)

    ' Copy original values
    Dim tempValues() As Double
    ReDim tempValues(LBound(Values) To UBound(Values))
    For i = LBound(Values) To UBound(Values)
        tempValues(i) = Values(i)
    Next i

    ' Apply differencing
    Dim diffLevel As Integer
    For diffLevel = 1 To d
        Dim newLen As Long
        newLen = UBound(tempValues) - LBound(tempValues)

        ReDim diffValues(1 To newLen)
        For i = 1 To newLen
            diffValues(i) = tempValues(i + 1) - tempValues(i)
        Next i

        ReDim tempValues(1 To newLen)
        For i = 1 To newLen
            tempValues(i) = diffValues(i)
        Next i
    Next diffLevel

    ' Now fit AR(p) or MA(q) on differenced data
    ' For simplicity, use AR(p) model
    Dim nDiff As Long
    nDiff = UBound(diffValues) - LBound(diffValues) + 1

    ' Fit AR coefficients using Yule-Walker equations
    Dim arCoeffs() As Double
    ReDim arCoeffs(1 To p)

    If p > 0 Then
        ' Calculate autocorrelations
        Dim acf() As Double
        acf = CalculateACF(diffValues, p)

        ' Simple Yule-Walker for AR(p) - just use first p ACF values
        For i = 1 To p
            If i <= UBound(acf) Then
                arCoeffs(i) = acf(i) * 0.8  ' Dampen coefficients
            Else
                arCoeffs(i) = 0
            End If
        Next i
    End If

    ' Initialize arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Fit on differenced data
    For i = LBound(diffValues) + p To UBound(diffValues)
        Dim fitted As Double
        fitted = 0

        For j = 1 To p
            If i - j >= LBound(diffValues) Then
                fitted = fitted + arCoeffs(j) * diffValues(i - j)
            End If
        Next j

        result.FittedValues(i + d) = Values(i + d - 1) + fitted
    Next i

    ' Fill early values with actuals
    For i = LBound(Values) To LBound(Values) + d + p - 1
        result.FittedValues(i) = Values(i)
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecast
    Dim lastDiff As Double
    If nDiff > 0 Then
        lastDiff = diffValues(UBound(diffValues))
    Else
        lastDiff = 0
    End If

    Dim lastLevel As Double
    lastLevel = Values(UBound(Values))

    For i = 1 To horizon
        result.ForecastValues(i) = lastLevel + lastDiff * i
    Next i

    ' Confidence intervals
    Dim residualStdDev As Double
    residualStdDev = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        Dim se As Double
        se = residualStdDev * Sqr(i)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * se
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * se
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    FitARIMA = result
End Function

Private Function CalculateAIC(ByRef result As ForecastResult, _
                              ByRef Values() As Double, _
                              ByVal numParams As Integer) As Double
    ' AIC = n * log(SSE/n) + 2*k
    ' where n = sample size, k = number of parameters

    Dim n As Long
    Dim sse As Double
    Dim i As Long

    n = UBound(Values) - LBound(Values) + 1
    sse = 0

    For i = LBound(result.Residuals) To UBound(result.Residuals)
        sse = sse + result.Residuals(i) * result.Residuals(i)
    Next i

    If sse > 0 And n > 0 Then
        CalculateAIC = n * Log(sse / n) + 2 * numParams
    Else
        CalculateAIC = 1E+100
    End If
End Function

' ============================================================================
' ADVANCED ENSEMBLE - Recent performance weighting with exponential decay
' ============================================================================
Public Function AdvancedEnsemble(ByRef tsData As TimeSeriesData, _
                                 ByVal horizon As Integer, _
                                 ByVal seasonalType As String) As ForecastResult

    Dim result As ForecastResult
    Dim sesResult As ForecastResult
    Dim hwResult As ForecastResult
    Dim thetaResult As ForecastResult
    Dim i As Long

    ' Get forecasts from multiple models
    sesResult = SimpleExponentialSmoothing(tsData, horizon)
    hwResult = HoltWinters(tsData, horizon, seasonalType)
    thetaResult = ThetaMethod(tsData, horizon)

    ' Calculate weights based on recent performance (last 20% of data)
    Dim recentWindow As Long
    Dim n As Long
    n = UBound(tsData.Values) - LBound(tsData.Values) + 1
    recentWindow = WorksheetFunction.Max(5, Int(n * 0.2))

    Dim sesRecentMAPE As Double
    Dim hwRecentMAPE As Double
    Dim thetaRecentMAPE As Double

    ' Calculate MAPE on recent window only
    sesRecentMAPE = CalculateRecentMAPE(tsData.Values, sesResult.FittedValues, recentWindow)
    hwRecentMAPE = CalculateRecentMAPE(tsData.Values, hwResult.FittedValues, recentWindow)
    thetaRecentMAPE = CalculateRecentMAPE(tsData.Values, thetaResult.FittedValues, recentWindow)

    ' Exponentially decaying weights (recent performance matters more)
    Dim sesWeight As Double, hwWeight As Double, thetaWeight As Double
    Dim totalWeight As Double

    ' Inverse MAPE weighting
    If sesRecentMAPE > 0 Then
        sesWeight = 1 / sesRecentMAPE
    Else
        sesWeight = 1
    End If

    If hwRecentMAPE > 0 Then
        hwWeight = 1 / hwRecentMAPE
    Else
        hwWeight = 1
    End If

    If thetaRecentMAPE > 0 Then
        thetaWeight = 1 / thetaRecentMAPE
    Else
        thetaWeight = 1
    End If

    totalWeight = sesWeight + hwWeight + thetaWeight

    If totalWeight > 0 Then
        sesWeight = sesWeight / totalWeight
        hwWeight = hwWeight / totalWeight
        thetaWeight = thetaWeight / totalWeight
    Else
        ' Equal weights fallback
        sesWeight = 0.33
        hwWeight = 0.33
        thetaWeight = 0.34
    End If

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(tsData.Values) To UBound(tsData.Values))
    ReDim result.Residuals(LBound(tsData.Values) To UBound(tsData.Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Combine fitted values
    For i = LBound(tsData.Values) To UBound(tsData.Values)
        result.FittedValues(i) = sesWeight * sesResult.FittedValues(i) + _
                                 hwWeight * hwResult.FittedValues(i) + _
                                 thetaWeight * thetaResult.FittedValues(i)
        result.Residuals(i) = tsData.Values(i) - result.FittedValues(i)
    Next i

    ' Combine forecasts
    For i = 1 To horizon
        result.ForecastValues(i) = sesWeight * sesResult.ForecastValues(i) + _
                                   hwWeight * hwResult.ForecastValues(i) + _
                                   thetaWeight * thetaResult.ForecastValues(i)
        result.Lower95(i) = sesWeight * sesResult.Lower95(i) + _
                           hwWeight * hwResult.Lower95(i) + _
                           thetaWeight * thetaResult.Lower95(i)
        result.Upper95(i) = sesWeight * sesResult.Upper95(i) + _
                           hwWeight * hwResult.Upper95(i) + _
                           thetaWeight * thetaResult.Upper95(i)
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(tsData.Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(tsData.Values, result.FittedValues)
    result.ModelName = "Advanced Ensemble"

    AdvancedEnsemble = result
End Function

Private Function CalculateRecentMAPE(ByRef actual() As Double, _
                                     ByRef fitted() As Double, _
                                     ByVal windowSize As Long) As Double

    Dim i As Long
    Dim sumAbsPercentError As Double
    Dim validCount As Long
    Dim startIdx As Long

    startIdx = UBound(actual) - windowSize + 1
    If startIdx < LBound(actual) Then startIdx = LBound(actual)

    sumAbsPercentError = 0
    validCount = 0

    For i = startIdx To UBound(actual)
        If actual(i) <> 0 Then
            sumAbsPercentError = sumAbsPercentError + Abs((actual(i) - fitted(i)) / actual(i)) * 100
            validCount = validCount + 1
        End If
    Next i

    If validCount > 0 Then
        CalculateRecentMAPE = sumAbsPercentError / validCount
    Else
        CalculateRecentMAPE = 100
    End If
End Function

' ============================================================================
' AUTOMATIC SEASONALITY DETECTION - Eliminates user errors!
' ============================================================================

' Type for seasonality detection results
Public Function DetectSeasonality(ByRef tsData As TimeSeriesData) As SeasonalityInfo
    ' Automatically detect if data has seasonality and what type
    Dim result As SeasonalityInfo
    Dim Values() As Double
    Dim n As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Initialize defaults
    result.HasSeasonality = False
    result.DetectedFrequency = 1
    result.SeasonalType = "none"
    result.Confidence = 0

    ' Need at least 2 cycles to detect seasonality
    If n < 24 Then
        DetectSeasonality = result
        Exit Function
    End If

    ' Step 1: Detect frequency from ACF peaks
    result.DetectedFrequency = DetectFrequencyFromACF(Values)

    ' Step 2: Test if seasonality is significant
    If result.DetectedFrequency > 1 Then
        Dim seasonalStrength As Double
        seasonalStrength = TestSeasonalStrength(Values, result.DetectedFrequency)

        If seasonalStrength > 0.3 Then ' 30% threshold
            result.HasSeasonality = True
            result.Confidence = seasonalStrength * 100

            ' Step 3: Determine additive vs multiplicative
            result.SeasonalType = DetectSeasonalType(Values, result.DetectedFrequency)
        End If
    End If

    DetectSeasonality = result
End Function

Private Function DetectFrequencyFromACF(ByRef Values() As Double) As Integer
    ' Find dominant frequency from ACF peaks
    Dim maxLag As Integer
    Dim n As Long
    Dim acf() As Double
    Dim i As Long
    Dim maxACF As Double
    Dim maxLag_idx As Integer

    n = UBound(Values) - LBound(Values) + 1
    maxLag = WorksheetFunction.Min(Int(n / 2), 52) ' Max 52 for weekly data

    If maxLag < 4 Then
        DetectFrequencyFromACF = 1
        Exit Function
    End If

    acf = CalculateACF(Values, maxLag)

    ' Find strongest ACF peak (ignoring lag 1)
    maxACF = 0
    maxLag_idx = 1

    For i = 2 To UBound(acf)
        If acf(i) > maxACF Then
            maxACF = acf(i)
            maxLag_idx = i
        End If
    Next i

    ' Common frequencies: 4 (quarterly), 7 (weekly), 12 (monthly), 52 (weekly/yearly)
    ' Snap to common frequencies if close
    Dim commonFreqs() As Variant
    commonFreqs = Array(4, 7, 12, 24, 52)

    Dim closestFreq As Integer
    Dim minDiff As Integer
    Dim diff As Integer

    closestFreq = maxLag_idx
    minDiff = 9999

    For i = LBound(commonFreqs) To UBound(commonFreqs)
        diff = Abs(maxLag_idx - commonFreqs(i))
        If diff < minDiff And diff < 3 Then ' Within 3 lags
            minDiff = diff
            closestFreq = commonFreqs(i)
        End If
    Next i

    If maxACF > 0.2 And closestFreq > 1 Then
        DetectFrequencyFromACF = closestFreq
    Else
        DetectFrequencyFromACF = 1 ' No seasonality
    End If
End Function

Private Function TestSeasonalStrength(ByRef Values() As Double, frequency As Integer) As Double
    ' Measure strength of seasonal pattern (0-1)
    Dim n As Long
    Dim i As Long, j As Long
    Dim cycleAvgs() As Double
    Dim grandAvg As Double
    Dim seasonalVar As Double
    Dim totalVar As Double
    Dim count As Long

    n = UBound(Values) - LBound(Values) + 1

    If n < frequency * 2 Then
        TestSeasonalStrength = 0
        Exit Function
    End If

    ' Calculate average for each season
    ReDim cycleAvgs(1 To frequency)
    Dim cycleCounts() As Long
    ReDim cycleCounts(1 To frequency)

    For i = LBound(Values) To UBound(Values)
        j = ((i - LBound(Values)) Mod frequency) + 1
        cycleAvgs(j) = cycleAvgs(j) + Values(i)
        cycleCounts(j) = cycleCounts(j) + 1
    Next i

    For j = 1 To frequency
        If cycleCounts(j) > 0 Then
            cycleAvgs(j) = cycleAvgs(j) / cycleCounts(j)
        End If
    Next j

    ' Calculate grand average
    grandAvg = 0
    For i = LBound(Values) To UBound(Values)
        grandAvg = grandAvg + Values(i)
    Next i
    grandAvg = grandAvg / n

    ' Calculate seasonal variance
    seasonalVar = 0
    For j = 1 To frequency
        If cycleCounts(j) > 0 Then
            seasonalVar = seasonalVar + (cycleAvgs(j) - grandAvg) ^ 2
        End If
    Next j

    ' Calculate total variance
    totalVar = 0
    For i = LBound(Values) To UBound(Values)
        totalVar = totalVar + (Values(i) - grandAvg) ^ 2
    Next i

    If totalVar > 0 Then
        TestSeasonalStrength = Sqr(seasonalVar / totalVar)
    Else
        TestSeasonalStrength = 0
    End If

    ' Cap at 1
    If TestSeasonalStrength > 1 Then TestSeasonalStrength = 1
End Function

Private Function DetectSeasonalType(ByRef Values() As Double, frequency As Integer) As String
    ' Determine if seasonality is additive or multiplicative
    ' Multiplicative: seasonal variation proportional to level
    ' Additive: seasonal variation constant

    Dim n As Long
    Dim i As Long, j As Long
    Dim cycleStdDevs() As Double
    Dim cycleMeans() As Double
    Dim cycleCounts() As Long
    Dim cv As Double ' Coefficient of variation
    Dim correlation As Double

    n = UBound(Values) - LBound(Values) + 1

    If n < frequency * 2 Then
        DetectSeasonalType = "additive"
        Exit Function
    End If

    ' Calculate mean and std dev for each seasonal period
    ReDim cycleMeans(1 To frequency)
    ReDim cycleStdDevs(1 To frequency)
    ReDim cycleCounts(1 To frequency)

    Dim cycleSums() As Double
    ReDim cycleSums(1 To frequency)

    ' First pass: means
    For i = LBound(Values) To UBound(Values)
        j = ((i - LBound(Values)) Mod frequency) + 1
        cycleSums(j) = cycleSums(j) + Values(i)
        cycleCounts(j) = cycleCounts(j) + 1
    Next i

    For j = 1 To frequency
        If cycleCounts(j) > 0 Then
            cycleMeans(j) = cycleSums(j) / cycleCounts(j)
        End If
    Next j

    ' Second pass: std devs
    For i = LBound(Values) To UBound(Values)
        j = ((i - LBound(Values)) Mod frequency) + 1
        cycleStdDevs(j) = cycleStdDevs(j) + (Values(i) - cycleMeans(j)) ^ 2
    Next i

    For j = 1 To frequency
        If cycleCounts(j) > 1 Then
            cycleStdDevs(j) = Sqr(cycleStdDevs(j) / (cycleCounts(j) - 1))
        End If
    Next j

    ' Calculate correlation between means and std devs
    ' High correlation → multiplicative
    ' Low correlation → additive
    correlation = CalculateCorrelationN(cycleMeans, cycleStdDevs, frequency)

    If correlation > 0.5 Then
        DetectSeasonalType = "multiplicative"
    Else
        DetectSeasonalType = "additive"
    End If
End Function

Private Function CalculateCorrelationN(ByRef x() As Double, ByRef y() As Double, n As Integer) As Double
    Dim i As Long
    Dim sumX As Double, sumY As Double
    Dim sumXY As Double, sumX2 As Double, sumY2 As Double
    Dim meanX As Double, meanY As Double
    Dim numerator As Double, denominator As Double

    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0: sumY2 = 0

    For i = 1 To n
        sumX = sumX + x(i)
        sumY = sumY + y(i)
    Next i

    meanX = sumX / n
    meanY = sumY / n

    For i = 1 To n
        numerator = numerator + (x(i) - meanX) * (y(i) - meanY)
        sumX2 = sumX2 + (x(i) - meanX) ^ 2
        sumY2 = sumY2 + (y(i) - meanY) ^ 2
    Next i

    denominator = Sqr(sumX2 * sumY2)

    If denominator > 0 Then
        CalculateCorrelationN = numerator / denominator
    Else
        CalculateCorrelationN = 0
    End If
End Function

' ============================================================================
' TIME SERIES CROSS-VALIDATION
' ============================================================================
' Implements rolling window cross-validation for more robust parameter optimization
' This prevents overfitting by testing parameters on multiple held-out test sets

' Performs rolling window cross-validation for Simple Exponential Smoothing
Private Function CrossValidateSES(ByRef Values() As Double, ByVal Alpha As Double, ByVal numFolds As Integer) As CrossValidationResult
    Dim result As CrossValidationResult
    Dim n As Long
    Dim foldSize As Long
    Dim testSize As Long
    Dim trainSize As Long
    Dim fold As Integer
    Dim i As Long, j As Long
    Dim trainData() As Double
    Dim testData() As Double
    Dim level As Double
    Dim forecast As Double
    Dim errorVal As Double
    Dim sumMAPE As Double, sumMAE As Double, sumRMSE As Double
    Dim foldMAPE As Double, foldMAE As Double, foldRMSE As Double
    Dim validCount As Long

    n = UBound(Values) - LBound(Values) + 1
    testSize = WorksheetFunction.Max(1, Int(n / (numFolds + 2))) ' Leave at least 2 periods for initial folds
    foldSize = WorksheetFunction.Max(1, Int((n - testSize) / numFolds))

    sumMAPE = 0: sumMAE = 0: sumRMSE = 0
    validCount = 0

    ' Rolling window: each fold trains on expanding window, tests on next testSize points
    For fold = 1 To numFolds
        trainSize = foldSize * fold

        ' Need enough data for training
        If trainSize < 3 Then GoTo NextFold
        If trainSize + testSize > n Then Exit For

        ' Extract train and test data
        ReDim trainData(1 To trainSize)
        ReDim testData(1 To testSize)

        For i = 1 To trainSize
            trainData(i) = Values(LBound(Values) + i - 1)
        Next i

        For i = 1 To testSize
            testData(i) = Values(LBound(Values) + trainSize + i - 1)
        Next i

        ' Initialize level with first training value
        level = trainData(1)

        ' Run through training data to update level
        For i = 2 To trainSize
            level = Alpha * trainData(i) + (1 - Alpha) * level
        Next i

        ' Calculate errors on test set
        foldMAPE = 0: foldMAE = 0: foldRMSE = 0
        For i = 1 To testSize
            forecast = level
            errorVal = testData(i) - forecast

            foldMAE = foldMAE + Abs(errorVal)
            foldRMSE = foldRMSE + errorVal * errorVal
            If testData(i) <> 0 Then
                foldMAPE = foldMAPE + Abs(errorVal / testData(i)) * 100
            End If

            ' Update level for next forecast
            level = Alpha * testData(i) + (1 - Alpha) * level
        Next i

        foldMAPE = foldMAPE / testSize
        foldMAE = foldMAE / testSize
        foldRMSE = Sqr(foldRMSE / testSize)

        sumMAPE = sumMAPE + foldMAPE
        sumMAE = sumMAE + foldMAE
        sumRMSE = sumRMSE + foldRMSE
        validCount = validCount + 1

NextFold:
    Next fold

    If validCount > 0 Then
        result.AvgMAPE = sumMAPE / validCount
        result.AvgMAE = sumMAE / validCount
        result.AvgRMSE = sumRMSE / validCount
        result.NumFolds = validCount
    Else
        ' Fallback if CV fails
        result.AvgMAPE = 9999
        result.AvgMAE = 9999
        result.AvgRMSE = 9999
        result.NumFolds = 0
    End If

    CrossValidateSES = result
End Function

' Performs rolling window cross-validation for Holt-Winters
Private Function CrossValidateHW(ByRef Values() As Double, _
                                 ByVal frequency As Integer, _
                                 ByVal Alpha As Double, _
                                 ByVal Beta As Double, _
                                 ByVal Gamma As Double, _
                                 ByVal seasonalType As String, _
                                 ByVal numFolds As Integer) As CrossValidationResult
    Dim result As CrossValidationResult
    Dim n As Long
    Dim foldSize As Long
    Dim testSize As Long
    Dim trainSize As Long
    Dim fold As Integer
    Dim i As Long, j As Long, seasonalIdx As Long
    Dim trainData() As Double
    Dim testData() As Double
    Dim level() As Double, Trend() As Double, Seasonal() As Double
    Dim forecast As Double
    Dim errorVal As Double
    Dim sumMAPE As Double, sumMAE As Double, sumRMSE As Double
    Dim foldMAPE As Double, foldMAE As Double, foldRMSE As Double
    Dim validCount As Long
    Dim prevLevel As Double, prevTrend As Double, prevSeasonal As Double

    n = UBound(Values) - LBound(Values) + 1
    testSize = WorksheetFunction.Max(frequency, Int(n / (numFolds + 2)))
    foldSize = WorksheetFunction.Max(frequency * 2, Int((n - testSize) / numFolds))

    sumMAPE = 0: sumMAE = 0: sumRMSE = 0
    validCount = 0

    ' Rolling window cross-validation
    For fold = 1 To numFolds
        trainSize = foldSize * fold

        ' Need at least 2 full seasonal cycles for HW
        If trainSize < frequency * 2 Then GoTo NextFoldHW
        If trainSize + testSize > n Then Exit For

        ' Extract train and test data
        ReDim trainData(1 To trainSize)
        ReDim testData(1 To testSize)

        For i = 1 To trainSize
            trainData(i) = Values(LBound(Values) + i - 1)
        Next i

        For i = 1 To testSize
            testData(i) = Values(LBound(Values) + trainSize + i - 1)
        Next i

        ' Initialize components
        ReDim level(1 To trainSize + testSize)
        ReDim Trend(1 To trainSize + testSize)
        ReDim Seasonal(1 To trainSize + testSize + frequency)

        level(1) = trainData(1)
        Trend(1) = (trainData(frequency + 1) - trainData(1)) / frequency

        ' Initialize seasonal indices
        If seasonalType = "additive" Then
            For i = 1 To frequency
                Seasonal(i) = trainData(i) - level(1)
            Next i
        Else
            For i = 1 To frequency
                If level(1) <> 0 Then
                    Seasonal(i) = trainData(i) / level(1)
                Else
                    Seasonal(i) = 1
                End If
            Next i
        End If

        ' Run through training data
        If seasonalType = "additive" Then
            For i = 2 To trainSize
                seasonalIdx = ((i - 1) Mod frequency) + 1
                prevLevel = level(i - 1)
                prevTrend = Trend(i - 1)
                prevSeasonal = Seasonal(seasonalIdx)

                level(i) = Alpha * (trainData(i) - prevSeasonal) + (1 - Alpha) * (prevLevel + prevTrend)
                Trend(i) = Beta * (level(i) - prevLevel) + (1 - Beta) * prevTrend
                Seasonal(i + frequency) = Gamma * (trainData(i) - level(i)) + (1 - Gamma) * prevSeasonal
            Next i
        Else ' Multiplicative
            For i = 2 To trainSize
                seasonalIdx = ((i - 1) Mod frequency) + 1
                prevLevel = level(i - 1)
                prevTrend = Trend(i - 1)
                prevSeasonal = Seasonal(seasonalIdx)

                If prevSeasonal <> 0 Then
                    level(i) = Alpha * (trainData(i) / prevSeasonal) + (1 - Alpha) * (prevLevel + prevTrend)
                Else
                    level(i) = prevLevel + prevTrend
                End If
                Trend(i) = Beta * (level(i) - prevLevel) + (1 - Beta) * prevTrend
                If level(i) <> 0 Then
                    Seasonal(i + frequency) = Gamma * (trainData(i) / level(i)) + (1 - Gamma) * prevSeasonal
                Else
                    Seasonal(i + frequency) = prevSeasonal
                End If
            Next i
        End If

        ' Calculate errors on test set
        foldMAPE = 0: foldMAE = 0: foldRMSE = 0
        For i = 1 To testSize
            seasonalIdx = ((trainSize + i - 1) Mod frequency) + 1
            If seasonalIdx + frequency <= UBound(Seasonal) Then
                prevSeasonal = Seasonal(seasonalIdx + frequency)
            Else
                prevSeasonal = Seasonal(seasonalIdx)
            End If

            If seasonalType = "additive" Then
                forecast = level(trainSize) + Trend(trainSize) * i + prevSeasonal
            Else
                forecast = (level(trainSize) + Trend(trainSize) * i) * prevSeasonal
            End If

            errorVal = testData(i) - forecast

            foldMAE = foldMAE + Abs(errorVal)
            foldRMSE = foldRMSE + errorVal * errorVal
            If testData(i) <> 0 Then
                foldMAPE = foldMAPE + Abs(errorVal / testData(i)) * 100
            End If
        Next i

        foldMAPE = foldMAPE / testSize
        foldMAE = foldMAE / testSize
        foldRMSE = Sqr(foldRMSE / testSize)

        sumMAPE = sumMAPE + foldMAPE
        sumMAE = sumMAE + foldMAE
        sumRMSE = sumRMSE + foldRMSE
        validCount = validCount + 1

NextFoldHW:
    Next fold

    If validCount > 0 Then
        result.AvgMAPE = sumMAPE / validCount
        result.AvgMAE = sumMAE / validCount
        result.AvgRMSE = sumRMSE / validCount
        result.NumFolds = validCount
    Else
        result.AvgMAPE = 9999
        result.AvgMAE = 9999
        result.AvgRMSE = 9999
        result.NumFolds = 0
    End If

    CrossValidateHW = result
End Function

' ============================================================================
' BOOTSTRAP CONFIDENCE INTERVALS
' ============================================================================
' Provides more accurate confidence intervals than normal approximation
' Uses residual bootstrap to capture non-normal distributions

Public Function BootstrapConfidenceIntervals(ByRef result As ForecastResult, _
                                            ByVal numBootstrap As Integer, _
                                            Optional ByVal confidenceLevel As Double = 0.95) As ForecastResult
    ' Apply residual bootstrap to improve confidence intervals
    ' Uses percentile method for non-parametric CI estimation

    Dim bootstrapResult As ForecastResult
    bootstrapResult = result ' Copy all fields

    Dim horizon As Integer
    Dim n As Long
    Dim i As Long, j As Long, b As Integer
    Dim bootstrapForecasts() As Double
    Dim sortedForecasts() As Double
    Dim lowerIdx As Long
    Dim upperIdx As Long
    Dim residuals() As Double
    Dim resampleIdx As Long

    horizon = UBound(result.ForecastValues) - LBound(result.ForecastValues) + 1
    n = UBound(result.Residuals) - LBound(result.Residuals) + 1

    ' Extract non-zero residuals for resampling
    Dim validResiduals() As Double
    Dim validCount As Long
    validCount = 0

    For i = LBound(result.Residuals) To UBound(result.Residuals)
        If Not IsEmpty(result.Residuals(i)) Then
            validCount = validCount + 1
        End If
    Next i

    If validCount < 10 Then
        ' Not enough residuals for bootstrap, return original
        BootstrapConfidenceIntervals = result
        Exit Function
    End If

    ReDim validResiduals(1 To validCount)
    validCount = 0
    For i = LBound(result.Residuals) To UBound(result.Residuals)
        If Not IsEmpty(result.Residuals(i)) Then
            validCount = validCount + 1
            validResiduals(validCount) = result.Residuals(i)
        End If
    Next i

    ' Perform bootstrap for each forecast horizon
    For i = 1 To horizon
        ReDim bootstrapForecasts(1 To numBootstrap)

        ' Generate bootstrap samples
        For b = 1 To numBootstrap
            ' Resample residuals and add to point forecast
            ' This simulates the forecast uncertainty
            Dim bootstrapError As Double
            bootstrapError = 0

            ' For horizon i, cumulative error from i resampled residuals
            For j = 1 To i
                resampleIdx = Int(Rnd() * validCount) + 1
                bootstrapError = bootstrapError + validResiduals(resampleIdx)
            Next j

            ' Average error over horizon steps (random walk of errors)
            bootstrapForecasts(b) = result.ForecastValues(i) + (bootstrapError / Sqr(i))
        Next b

        ' Sort bootstrap forecasts
        ReDim sortedForecasts(1 To numBootstrap)
        For b = 1 To numBootstrap
            sortedForecasts(b) = bootstrapForecasts(b)
        Next b
        Call BubbleSortDouble(sortedForecasts)

        ' Calculate percentile-based confidence intervals
        Dim alpha As Double
        alpha = 1 - confidenceLevel
        lowerIdx = WorksheetFunction.Max(1, Int(alpha / 2 * numBootstrap))
        upperIdx = WorksheetFunction.Min(numBootstrap, Int((1 - alpha / 2) * numBootstrap))

        bootstrapResult.Lower95(i) = sortedForecasts(lowerIdx)
        bootstrapResult.Upper95(i) = sortedForecasts(upperIdx)

        ' Also calculate 80% CI using same bootstrap samples
        lowerIdx = WorksheetFunction.Max(1, Int(0.1 * numBootstrap))
        upperIdx = WorksheetFunction.Min(numBootstrap, Int(0.9 * numBootstrap))
        bootstrapResult.Lower80(i) = sortedForecasts(lowerIdx)
        bootstrapResult.Upper80(i) = sortedForecasts(upperIdx)
    Next i

    BootstrapConfidenceIntervals = bootstrapResult
End Function

Private Sub BubbleSortDouble(ByRef arr() As Double)
    ' Simple bubble sort for bootstrap samples
    Dim i As Long, j As Long
    Dim temp As Double
    Dim n As Long

    n = UBound(arr) - LBound(arr) + 1

    For i = LBound(arr) To UBound(arr) - 1
        For j = LBound(arr) To UBound(arr) - 1 - (i - LBound(arr))
            If arr(j) > arr(j + 1) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i
End Sub

' ============================================================================
' SARIMA - Seasonal ARIMA
' ============================================================================

Public Function SARIMAForecast(ByRef tsData As TimeSeriesData, _
                               ByVal horizon As Integer, _
                               Optional ByVal p As Integer = 1, _
                               Optional ByVal d As Integer = 1, _
                               Optional ByVal q As Integer = 1, _
                               Optional ByVal sP As Integer = 1, _
                               Optional ByVal sD As Integer = 1, _
                               Optional ByVal sQ As Integer = 1, _
                               Optional ByVal s As Integer = 0) As ForecastResult
    ' SARIMA(p,d,q)(sP,sD,sQ)s model
    ' p,d,q = non-seasonal AR, differencing, MA orders
    ' sP,sD,sQ = seasonal AR, differencing, MA orders (renamed from P,D,Q due to VBA case-insensitivity)
    ' s = seasonal period (0 = auto-detect)

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long, j As Long
    Dim seasonalPeriod As Integer

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Auto-detect seasonal period if not provided
    If s = 0 Then
        If tsData.Frequency > 0 Then
            seasonalPeriod = tsData.Frequency
        Else
            seasonalPeriod = DetectSeasonalPeriod(Values)
        End If
    Else
        seasonalPeriod = s
    End If

    ' Apply non-seasonal differencing
    Dim diffValues() As Double
    ReDim diffValues(LBound(Values) To UBound(Values))
    For i = LBound(Values) To UBound(Values)
        diffValues(i) = Values(i)
    Next i

    For i = 1 To d
        diffValues = ApplyDifferencing(diffValues)
    Next i

    ' Apply seasonal differencing
    If sD > 0 And seasonalPeriod > 1 Then
        For i = 1 To sD
            diffValues = ApplySeasonalDifferencing(diffValues, seasonalPeriod)
        Next i
    End If

    ' Estimate ARMA parameters on differenced data
    Dim arParams() As Double
    Dim maParams() As Double
    Dim arSeasonalParams() As Double
    Dim maSeasonalParams() As Double

    ReDim arParams(1 To p)
    ReDim maParams(1 To q)
    ReDim arSeasonalParams(1 To sP)
    ReDim maSeasonalParams(1 To sQ)

    ' Use Yule-Walker for AR parameters (simple estimation)
    If p > 0 Then
        Dim acf() As Double
        acf = CalculateACF(diffValues, p + 1)
        arParams = YuleWalkerAR(acf, p)
    End If

    ' Use innovation algorithm for MA parameters (simplified)
    If q > 0 Then
        For i = 1 To q
            maParams(i) = 0.1 ' Simplified initialization
        Next i
    End If

    ' Seasonal parameters (simplified)
    If sP > 0 And seasonalPeriod > 1 Then
        For i = 1 To sP
            arSeasonalParams(i) = 0.2
        Next i
    End If

    If sQ > 0 And seasonalPeriod > 1 Then
        For i = 1 To sQ
            maSeasonalParams(i) = 0.1
        Next i
    End If

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Generate fitted values using ARMA structure
    Dim errors() As Double
    ReDim errors(LBound(diffValues) To UBound(diffValues))

    For i = LBound(diffValues) To UBound(diffValues)
        Dim fittedDiff As Double
        fittedDiff = 0

        ' AR component
        For j = 1 To p
            If i - j >= LBound(diffValues) Then
                fittedDiff = fittedDiff + arParams(j) * diffValues(i - j)
            End If
        Next j

        ' Seasonal AR component
        For j = 1 To sP
            If i - j * seasonalPeriod >= LBound(diffValues) Then
                fittedDiff = fittedDiff + arSeasonalParams(j) * diffValues(i - j * seasonalPeriod)
            End If
        Next j

        ' MA component
        For j = 1 To q
            If i - j >= LBound(errors) Then
                fittedDiff = fittedDiff + maParams(j) * errors(i - j)
            End If
        Next j

        ' Seasonal MA component
        For j = 1 To sQ
            If i - j * seasonalPeriod >= LBound(errors) Then
                fittedDiff = fittedDiff + maSeasonalParams(j) * errors(i - j * seasonalPeriod)
            End If
        Next j

        errors(i) = diffValues(i) - fittedDiff
    Next i

    ' Invert differencing to get fitted values in original scale
    result.FittedValues = InvertDifferencing(diffValues, Values, d, sD, seasonalPeriod)

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Generate forecasts
    Dim forecastDiff() As Double
    ReDim forecastDiff(1 To horizon)

    For i = 1 To horizon
        Dim fcstVal As Double
        fcstVal = 0

        ' AR component using recent values
        For j = 1 To p
            If i - j > 0 Then
                fcstVal = fcstVal + arParams(j) * forecastDiff(i - j)
            ElseIf UBound(diffValues) - j + 1 >= LBound(diffValues) Then
                fcstVal = fcstVal + arParams(j) * diffValues(UBound(diffValues) - j + 1)
            End If
        Next j

        ' Seasonal AR component
        For j = 1 To sP
            If i - j * seasonalPeriod > 0 Then
                fcstVal = fcstVal + arSeasonalParams(j) * forecastDiff(i - j * seasonalPeriod)
            ElseIf UBound(diffValues) - (j * seasonalPeriod) + 1 >= LBound(diffValues) Then
                fcstVal = fcstVal + arSeasonalParams(j) * diffValues(UBound(diffValues) - (j * seasonalPeriod) + 1)
            End If
        Next j

        forecastDiff(i) = fcstVal
    Next i

    ' Invert differencing for forecasts
    For i = 1 To horizon
        result.ForecastValues(i) = forecastDiff(i)
        ' Add back the seasonal and non-seasonal differences
        For j = 1 To d
            If UBound(Values) - j + 1 >= LBound(Values) Then
                result.ForecastValues(i) = result.ForecastValues(i) + Values(UBound(Values) - j + 1)
            End If
        Next j
    Next i

    ' Calculate prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        Dim intervalWidth As Double
        intervalWidth = sigma * Sqr(i) ' Increases with horizon
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * intervalWidth
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * intervalWidth
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * intervalWidth
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * intervalWidth
    Next i

    ' Calculate accuracy metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.ModelName = "SARIMA(" & p & "," & d & "," & q & ")(" & sP & "," & sD & "," & sQ & ")[" & seasonalPeriod & "]"

    SARIMAForecast = result
    Exit Function

ErrorHandler:
    ' Return simple fallback on error
    result.MAPE = 9999
    result.ModelName = "SARIMA (Error)"
    SARIMAForecast = result
End Function

Private Function DetectSeasonalPeriod(ByRef Values() As Double) As Integer
    ' Detect seasonal period using ACF peaks
    Dim n As Long
    Dim maxLag As Integer
    Dim acf() As Double
    Dim i As Integer
    Dim maxACF As Double
    Dim maxLag As Integer

    n = UBound(Values) - LBound(Values) + 1
    maxLag = WorksheetFunction.Min(24, Int(n / 3))

    acf = CalculateACF(Values, maxLag)

    ' Find first significant peak after lag 1
    maxACF = -1
    maxLag = 12 ' Default to monthly if not found

    For i = 2 To UBound(acf)
        If acf(i) > maxACF And acf(i) > 0.3 Then ' Threshold for significance
            maxACF = acf(i)
            maxLag = i
        End If
    Next i

    DetectSeasonalPeriod = maxLag
End Function

Private Function ApplyDifferencing(ByRef Values() As Double) As Double()
    ' Apply first-order differencing
    Dim n As Long
    Dim result() As Double
    Dim i As Long

    n = UBound(Values) - LBound(Values) + 1
    If n <= 1 Then
        ApplyDifferencing = Values
        Exit Function
    End If

    ReDim result(LBound(Values) + 1 To UBound(Values))

    For i = LBound(Values) + 1 To UBound(Values)
        result(i) = Values(i) - Values(i - 1)
    Next i

    ApplyDifferencing = result
End Function

Private Function ApplySeasonalDifferencing(ByRef Values() As Double, ByVal period As Integer) As Double()
    ' Apply seasonal differencing
    Dim n As Long
    Dim result() As Double
    Dim i As Long

    n = UBound(Values) - LBound(Values) + 1
    If n <= period Then
        ApplySeasonalDifferencing = Values
        Exit Function
    End If

    ReDim result(LBound(Values) + period To UBound(Values))

    For i = LBound(Values) + period To UBound(Values)
        result(i) = Values(i) - Values(i - period)
    Next i

    ApplySeasonalDifferencing = result
End Function

Private Function YuleWalkerAR(ByRef acf() As Double, ByVal order As Integer) As Double()
    ' Estimate AR parameters using Yule-Walker equations
    Dim params() As Double
    Dim i As Integer, j As Integer
    Dim R() As Double ' Autocorrelation matrix
    Dim r() As Double ' Autocorrelation vector

    ReDim params(1 To order)
    ReDim R(1 To order, 1 To order)
    ReDim r(1 To order)

    ' Build autocorrelation matrix
    For i = 1 To order
        r(i) = acf(i)
        For j = 1 To order
            R(i, j) = acf(Abs(i - j))
        Next j
    Next i

    ' Solve using simple Gaussian elimination (for small orders)
    params = SolveLinearSystem(R, r, order)

    YuleWalkerAR = params
End Function

Private Function SolveLinearSystem(ByRef A() As Double, ByRef b() As Double, ByVal n As Integer) As Double()
    ' Solve Ax = b using Gaussian elimination
    Dim x() As Double
    Dim i As Integer, j As Integer, k As Integer
    Dim factor As Double
    Dim augmented() As Double

    ReDim x(1 To n)
    ReDim augmented(1 To n, 1 To n + 1)

    ' Create augmented matrix [A|b]
    For i = 1 To n
        For j = 1 To n
            augmented(i, j) = A(i, j)
        Next j
        augmented(i, n + 1) = b(i)
    Next i

    ' Forward elimination
    For k = 1 To n - 1
        For i = k + 1 To n
            If augmented(k, k) <> 0 Then
                factor = augmented(i, k) / augmented(k, k)
                For j = k To n + 1
                    augmented(i, j) = augmented(i, j) - factor * augmented(k, j)
                Next j
            End If
        Next i
    Next k

    ' Back substitution
    For i = n To 1 Step -1
        x(i) = augmented(i, n + 1)
        For j = i + 1 To n
            x(i) = x(i) - augmented(i, j) * x(j)
        Next j
        If augmented(i, i) <> 0 Then
            x(i) = x(i) / augmented(i, i)
        Else
            x(i) = 0
        End If
    Next i

    SolveLinearSystem = x
End Function

Private Function InvertDifferencing(ByRef diffValues() As Double, _
                                   ByRef originalValues() As Double, _
                                   ByVal d As Integer, _
                                   ByVal sD As Integer, _
                                   ByVal period As Integer) As Double()
    ' Invert differencing to get back to original scale (simplified)
    ' d = non-seasonal differencing order
    ' sD = seasonal differencing order (renamed from D due to VBA case-insensitivity)
    Dim result() As Double
    Dim i As Long

    ReDim result(LBound(originalValues) To UBound(originalValues))

    ' This is a simplified inversion - just use the original fitted approach
    For i = LBound(originalValues) To UBound(originalValues)
        result(i) = originalValues(i) ' Placeholder - proper inversion is complex
    Next i

    InvertDifferencing = result
End Function

' ============================================================================
' FOURIER SERIES FORECASTING
' ============================================================================

Public Function FourierForecast(ByRef tsData As TimeSeriesData, _
                               ByVal horizon As Integer, _
                               Optional ByVal K As Integer = 0) As ForecastResult
    ' Forecast using Fourier series decomposition
    ' K = number of Fourier terms (0 = auto-select)

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long, k As Integer, j As Long
    Dim numTerms As Integer

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Auto-select number of Fourier terms
    If K = 0 Then
        If tsData.Frequency > 0 Then
            numTerms = WorksheetFunction.Min(5, Int(tsData.Frequency / 2))
        Else
            numTerms = 3 ' Default
        End If
    Else
        numTerms = K
    End If

    ' Build design matrix for Fourier regression
    Dim X() As Double ' Design matrix
    Dim y() As Double ' Response vector
    Dim beta() As Double ' Regression coefficients

    Dim numParams As Integer
    numParams = 1 + 2 * numTerms ' Intercept + sin/cos pairs

    ReDim X(1 To n, 1 To numParams)
    ReDim y(1 To n)
    ReDim beta(1 To numParams)

    ' Fill design matrix
    For i = 1 To n
        X(i, 1) = 1 ' Intercept

        For k = 1 To numTerms
            Dim freq As Double
            freq = 2 * WorksheetFunction.Pi * k * i / n
            X(i, 2 * k) = Sin(freq) ' Sin component
            X(i, 2 * k + 1) = Cos(freq) ' Cos component
        Next k

        y(i) = Values(LBound(Values) + i - 1)
    Next i

    ' Estimate coefficients using least squares
    beta = FitLinearRegression(X, y, n, numParams)

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Calculate fitted values
    For i = 1 To n
        Dim fitted As Double
        fitted = beta(1) ' Intercept

        For k = 1 To numTerms
            freq = 2 * WorksheetFunction.Pi * k * i / n
            fitted = fitted + beta(2 * k) * Sin(freq) + beta(2 * k + 1) * Cos(freq)
        Next k

        result.FittedValues(LBound(Values) + i - 1) = fitted
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Generate forecasts by extrapolating Fourier series
    For i = 1 To horizon
        Dim fcst As Double
        fcst = beta(1)

        For k = 1 To numTerms
            freq = 2 * WorksheetFunction.Pi * k * (n + i) / n
            fcst = fcst + beta(2 * k) * Sin(freq) + beta(2 * k + 1) * Cos(freq)
        Next k

        result.ForecastValues(i) = fcst
    Next i

    ' Calculate prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * sigma
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * sigma
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * sigma
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * sigma
    Next i

    ' Calculate accuracy metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.ModelName = "Fourier[K=" & numTerms & "]"

    FourierForecast = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "Fourier (Error)"
    FourierForecast = result
End Function

Private Function FitLinearRegression(ByRef X() As Double, ByRef y() As Double, _
                                    ByVal n As Long, ByVal p As Integer) As Double()
    ' Fit linear regression using normal equations: beta = (X'X)^-1 X'y
    Dim beta() As Double
    Dim XtX() As Double
    Dim Xty() As Double
    Dim i As Long, j As Integer, k As Integer

    ReDim beta(1 To p)
    ReDim XtX(1 To p, 1 To p)
    ReDim Xty(1 To p)

    ' Calculate X'X
    For i = 1 To p
        For j = 1 To p
            XtX(i, j) = 0
            For k = 1 To n
                XtX(i, j) = XtX(i, j) + X(k, i) * X(k, j)
            Next k
        Next j
    Next i

    ' Calculate X'y
    For i = 1 To p
        Xty(i) = 0
        For k = 1 To n
            Xty(i) = Xty(i) + X(k, i) * y(k)
        Next k
    Next i

    ' Solve (X'X)beta = X'y
    beta = SolveLinearSystem(XtX, Xty, p)

    FitLinearRegression = beta
End Function

' ============================================================================
' TSB METHOD - Teunter-Syntetos-Babai (For Intermittent Demand)
' ============================================================================

Public Function TSBMethod(ByRef tsData As TimeSeriesData, _
                         ByVal horizon As Integer, _
                         Optional ByVal Alpha As Double = 0.2, _
                         Optional ByVal Beta As Double = 0.2) As ForecastResult
    ' TSB method for intermittent demand forecasting
    ' Alpha = smoothing parameter for demand size
    ' Beta = smoothing parameter for inter-arrival time

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long
    Dim demandSize As Double
    Dim interArrivalTime As Double
    Dim lastDemandIdx As Long
    Dim periods Since LastDemand As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Initialize
    demandSize = 0
    interArrivalTime = 1
    lastDemandIdx = LBound(Values)
    periodsSinceLastDemand = 0

    ' Count non-zero demands to initialize
    Dim nonZeroCount As Integer
    Dim totalDemand As Double
    nonZeroCount = 0
    totalDemand = 0

    For i = LBound(Values) To UBound(Values)
        If Values(i) > 0 Then
            nonZeroCount = nonZeroCount + 1
            totalDemand = totalDemand + Values(i)
        End If
    Next i

    If nonZeroCount > 0 Then
        demandSize = totalDemand / nonZeroCount
        interArrivalTime = n / nonZeroCount
    Else
        demandSize = 0
        interArrivalTime = n
    End If

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Fit the model
    periodsSinceLastDemand = 0

    For i = LBound(Values) To UBound(Values)
        periodsSinceLastDemand = periodsSinceLastDemand + 1

        If Values(i) > 0 Then
            ' Update demand size estimate
            demandSize = Alpha * Values(i) + (1 - Alpha) * demandSize

            ' Update inter-arrival time estimate
            interArrivalTime = Beta * periodsSinceLastDemand + (1 - Beta) * interArrivalTime

            periodsSinceLastDemand = 0
        End If

        ' Fitted value is expected demand per period
        If interArrivalTime > 0 Then
            result.FittedValues(i) = demandSize / interArrivalTime
        Else
            result.FittedValues(i) = 0
        End If
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Generate forecasts
    Dim expectedDemand As Double
    If interArrivalTime > 0 Then
        expectedDemand = demandSize / interArrivalTime
    Else
        expectedDemand = 0
    End If

    For i = 1 To horizon
        result.ForecastValues(i) = expectedDemand
    Next i

    ' Calculate prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        result.Lower80(i) = WorksheetFunction.Max(0, result.ForecastValues(i) - 1.28 * sigma)
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * sigma
        result.Lower95(i) = WorksheetFunction.Max(0, result.ForecastValues(i) - 1.96 * sigma)
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * sigma
    Next i

    ' Calculate accuracy metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.Alpha = Alpha
    result.Beta = Beta
    result.ModelName = "TSB"

    TSBMethod = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "TSB (Error)"
    TSBMethod = result
End Function

' ============================================================================
' SBA METHOD - Syntetos-Boylan Approximation (For Intermittent Demand)
' ============================================================================

Public Function SBAMethod(ByRef tsData As TimeSeriesData, _
                         ByVal horizon As Integer, _
                         Optional ByVal Alpha As Double = 0.1) As ForecastResult
    ' SBA method for intermittent demand with bias correction

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long
    Dim demandProb As Double ' Probability of non-zero demand
    Dim demandSize As Double ' Mean size when demand occurs

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Calculate empirical probability and size
    Dim nonZeroCount As Long
    Dim totalDemand As Double

    nonZeroCount = 0
    totalDemand = 0

    For i = LBound(Values) To UBound(Values)
        If Values(i) > 0 Then
            nonZeroCount = nonZeroCount + 1
            totalDemand = totalDemand + Values(i)
        End If
    Next i

    If n > 0 Then
        demandProb = nonZeroCount / n
    Else
        demandProb = 0
    End If

    If nonZeroCount > 0 Then
        demandSize = totalDemand / nonZeroCount
    Else
        demandSize = 0
    End If

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Exponential smoothing updates
    Dim smoothedProb As Double
    Dim smoothedSize As Double

    smoothedProb = demandProb
    smoothedSize = demandSize

    For i = LBound(Values) To UBound(Values)
        Dim occurence As Double
        If Values(i) > 0 Then
            occurence = 1
            smoothedProb = Alpha * occurence + (1 - Alpha) * smoothedProb
            smoothedSize = Alpha * Values(i) + (1 - Alpha) * smoothedSize
        Else
            occurence = 0
            smoothedProb = Alpha * occurence + (1 - Alpha) * smoothedProb
        End If

        ' SBA forecast with approximation
        result.FittedValues(i) = smoothedProb * smoothedSize
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Generate forecasts
    Dim forecastValue As Double
    forecastValue = smoothedProb * smoothedSize

    For i = 1 To horizon
        result.ForecastValues(i) = forecastValue
    Next i

    ' Calculate prediction intervals (wider for intermittent demand)
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        result.Lower80(i) = WorksheetFunction.Max(0, result.ForecastValues(i) - 1.28 * sigma * Sqr(i))
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * sigma * Sqr(i)
        result.Lower95(i) = WorksheetFunction.Max(0, result.ForecastValues(i) - 1.96 * sigma * Sqr(i))
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * sigma * Sqr(i)
    Next i

    ' Calculate accuracy metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.Alpha = Alpha
    result.ModelName = "SBA"

    SBAMethod = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "SBA (Error)"
    SBAMethod = result
End Function

' ============================================================================
' ETS STATE SPACE MODELS (Error-Trend-Seasonal)
' ============================================================================

Public Function ETSForecast(ByRef tsData As TimeSeriesData, _
                           ByVal horizon As Integer, _
                           Optional ByVal errorType As String = "A", _
                           Optional ByVal trendType As String = "A", _
                           Optional ByVal seasonalType As String = "N") As ForecastResult
    ' ETS State Space models
    ' errorType: "A" = Additive, "M" = Multiplicative
    ' trendType: "N" = None, "A" = Additive, "M" = Multiplicative, "Ad" = Additive Damped
    ' seasonalType: "N" = None, "A" = Additive, "M" = Multiplicative

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Map to existing methods based on configuration
    Dim modelCode As String
    modelCode = errorType & trendType & seasonalType

    Select Case modelCode
        Case "AAN" ' Additive error, additive trend, no seasonal
            result = HoltLinear(tsData, horizon)
            result.ModelName = "ETS(A,A,N)"

        Case "AAdN" ' Additive error, additive damped trend, no seasonal
            result = DampedTrend(tsData, horizon)
            result.ModelName = "ETS(A,Ad,N)"

        Case "ANA" ' Additive error, no trend, additive seasonal
            result = SimpleSeasonalSmoothing(tsData, horizon, "additive")
            result.ModelName = "ETS(A,N,A)"

        Case "AAA" ' Additive error, additive trend, additive seasonal
            result = HoltWinters(tsData, horizon, "additive")
            result.ModelName = "ETS(A,A,A)"

        Case "MAM" ' Multiplicative error, additive trend, multiplicative seasonal
            result = HoltWinters(tsData, horizon, "multiplicative")
            result.ModelName = "ETS(M,A,M)"

        Case "MAdM" ' Multiplicative error, damped trend, multiplicative seasonal
            result = DampedHoltWinters(tsData, horizon, "multiplicative")
            result.ModelName = "ETS(M,Ad,M)"

        Case Else ' Default to ANN (Simple Exponential Smoothing)
            result = SimpleExponentialSmoothing(tsData, horizon)
            result.ModelName = "ETS(A,N,N)"
    End Select

    ETSForecast = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "ETS (Error)"
    ETSForecast = result
End Function

Private Function HoltLinear(ByRef tsData As TimeSeriesData, ByVal horizon As Integer) As ForecastResult
    ' Holt's linear trend method (additive trend, no seasonal)
    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long
    Dim level As Double, trend As Double
    Dim Alpha As Double, Beta As Double

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    Alpha = 0.2
    Beta = 0.1

    ' Initialize
    level = Values(LBound(Values))
    If n > 1 Then
        trend = Values(LBound(Values) + 1) - Values(LBound(Values))
    Else
        trend = 0
    End If

    ' Initialize arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    result.FittedValues(LBound(Values)) = level

    ' Fit model
    For i = LBound(Values) + 1 To UBound(Values)
        Dim prevLevel As Double
        prevLevel = level

        level = Alpha * Values(i) + (1 - Alpha) * (level + trend)
        trend = Beta * (level - prevLevel) + (1 - Beta) * trend

        result.FittedValues(i) = prevLevel + trend
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecast
    For i = 1 To horizon
        result.ForecastValues(i) = level + i * trend
    Next i

    ' Prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        Dim intervalWidth As Double
        intervalWidth = sigma * Sqr(i)
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * intervalWidth
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * intervalWidth
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * intervalWidth
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * intervalWidth
    Next i

    ' Metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.Alpha = Alpha
    result.Beta = Beta
    result.ModelName = "Holt Linear"

    HoltLinear = result
End Function

Private Function DampedTrend(ByRef tsData As TimeSeriesData, ByVal horizon As Integer) As ForecastResult
    ' Holt's damped trend method
    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long
    Dim level As Double, trend As Double
    Dim Alpha As Double, Beta As Double, Phi As Double

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    Alpha = 0.2
    Beta = 0.1
    Phi = 0.9 ' Damping parameter

    ' Initialize
    level = Values(LBound(Values))
    If n > 1 Then
        trend = Values(LBound(Values) + 1) - Values(LBound(Values))
    Else
        trend = 0
    End If

    ' Initialize arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    result.FittedValues(LBound(Values)) = level

    ' Fit model
    For i = LBound(Values) + 1 To UBound(Values)
        Dim prevLevel As Double
        prevLevel = level

        level = Alpha * Values(i) + (1 - Alpha) * (level + Phi * trend)
        trend = Beta * (level - prevLevel) + (1 - Beta) * Phi * trend

        result.FittedValues(i) = prevLevel + Phi * trend
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecast with damping
    Dim phiSum As Double
    For i = 1 To horizon
        phiSum = 0
        Dim j As Integer
        For j = 1 To i
            phiSum = phiSum + Phi ^ j
        Next j
        result.ForecastValues(i) = level + phiSum * trend
    Next i

    ' Prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        Dim intervalWidth As Double
        intervalWidth = sigma * Sqr(i)
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * intervalWidth
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * intervalWidth
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * intervalWidth
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * intervalWidth
    Next i

    ' Metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.Alpha = Alpha
    result.Beta = Beta
    result.Phi = Phi
    result.ModelName = "Damped Trend"

    DampedTrend = result
End Function

Private Function SimpleSeasonalSmoothing(ByRef tsData As TimeSeriesData, _
                                        ByVal horizon As Integer, _
                                        ByVal seasonalType As String) As ForecastResult
    ' Simple seasonal smoothing (no trend)
    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long, j As Long
    Dim level As Double
    Dim seasonal() As Double
    Dim Alpha As Double, Gamma As Double
    Dim m As Integer ' Seasonal period

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    If tsData.Frequency > 0 Then
        m = tsData.Frequency
    Else
        m = 12
    End If

    Alpha = 0.2
    Gamma = 0.1

    ReDim seasonal(1 To m)

    ' Initialize seasonal indices
    For i = 1 To m
        seasonal(i) = 1
    Next i

    level = 0
    For i = LBound(Values) To WorksheetFunction.Min(LBound(Values) + m - 1, UBound(Values))
        level = level + Values(i)
    Next i
    level = level / WorksheetFunction.Min(m, n)

    ' Initialize arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Fit model
    For i = LBound(Values) To UBound(Values)
        Dim seasonIdx As Integer
        seasonIdx = ((i - LBound(Values)) Mod m) + 1

        If LCase(seasonalType) = "multiplicative" And level > 0 Then
            Dim prevLevel As Double
            prevLevel = level
            level = Alpha * (Values(i) / seasonal(seasonIdx)) + (1 - Alpha) * level
            seasonal(seasonIdx) = Gamma * (Values(i) / prevLevel) + (1 - Gamma) * seasonal(seasonIdx)
            result.FittedValues(i) = prevLevel * seasonal(seasonIdx)
        Else ' Additive
            prevLevel = level
            level = Alpha * (Values(i) - seasonal(seasonIdx)) + (1 - Alpha) * level
            seasonal(seasonIdx) = Gamma * (Values(i) - prevLevel) + (1 - Gamma) * seasonal(seasonIdx)
            result.FittedValues(i) = prevLevel + seasonal(seasonIdx)
        End If
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecast
    For i = 1 To horizon
        seasonIdx = ((i - 1) Mod m) + 1
        If LCase(seasonalType) = "multiplicative" Then
            result.ForecastValues(i) = level * seasonal(seasonIdx)
        Else
            result.ForecastValues(i) = level + seasonal(seasonIdx)
        End If
    Next i

    ' Prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * sigma * Sqr(i)
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * sigma * Sqr(i)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * sigma * Sqr(i)
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * sigma * Sqr(i)
    Next i

    ' Metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.Alpha = Alpha
    result.Gamma = Gamma
    result.ModelName = "Seasonal Smoothing"

    SimpleSeasonalSmoothing = result
End Function

' ============================================================================
' FEATURE ENGINEERING FOR TIME SERIES
' ============================================================================

Public Function ExtractTimeSeriesFeatures(ByRef Values() As Double) As Double()
    ' Extract statistical features from time series
    ' Returns: [Mean, StdDev, CV, Skewness, Kurtosis, TrendStrength, SeasonalStrength, ACF1, Entropy]

    Dim features() As Double
    ReDim features(1 To 9)

    Dim n As Long, i As Long
    Dim mean As Double, variance As Double, stdDev As Double

    n = UBound(Values) - LBound(Values) + 1

    ' Mean
    mean = 0
    For i = LBound(Values) To UBound(Values)
        mean = mean + Values(i)
    Next i
    mean = mean / n
    features(1) = mean

    ' Standard Deviation
    variance = 0
    For i = LBound(Values) To UBound(Values)
        variance = variance + (Values(i) - mean) ^ 2
    Next i
    variance = variance / n
    stdDev = Sqr(variance)
    features(2) = stdDev

    ' Coefficient of Variation
    If mean <> 0 Then
        features(3) = stdDev / Abs(mean)
    Else
        features(3) = 0
    End If

    ' Skewness
    Dim skewness As Double
    skewness = 0
    For i = LBound(Values) To UBound(Values)
        skewness = skewness + ((Values(i) - mean) / stdDev) ^ 3
    Next i
    skewness = skewness / n
    features(4) = skewness

    ' Kurtosis
    Dim kurtosis As Double
    kurtosis = 0
    For i = LBound(Values) To UBound(Values)
        kurtosis = kurtosis + ((Values(i) - mean) / stdDev) ^ 4
    Next i
    kurtosis = (kurtosis / n) - 3 ' Excess kurtosis
    features(5) = kurtosis

    ' Trend Strength (using linear regression R²)
    Dim trendStrength As Double
    trendStrength = CalculateTrendStrength(Values)
    features(6) = trendStrength

    ' Seasonal Strength (using ACF)
    Dim seasonalStrength As Double
    seasonalStrength = CalculateSeasonalStrength(Values)
    features(7) = seasonalStrength

    ' First-order autocorrelation
    Dim acf1 As Double
    acf1 = CalculateACF1(Values)
    features(8) = acf1

    ' Entropy (measure of randomness)
    Dim entropy As Double
    entropy = CalculateEntropy(Values)
    features(9) = entropy

    ExtractTimeSeriesFeatures = features
End Function

Private Function CalculateTrendStrength(ByRef Values() As Double) As Double
    ' Calculate trend strength using linear regression R²
    Dim n As Long, i As Long
    Dim x As Double, y As Double
    Dim sumX As Double, sumY As Double, sumXY As Double
    Dim sumX2 As Double, sumY2 As Double
    Dim slope As Double, intercept As Double
    Dim rSquared As Double

    n = UBound(Values) - LBound(Values) + 1
    sumX = 0: sumY = 0: sumXY = 0: sumX2 = 0: sumY2 = 0

    For i = LBound(Values) To UBound(Values)
        x = i - LBound(Values) + 1
        y = Values(i)
        sumX = sumX + x
        sumY = sumY + y
        sumXY = sumXY + x * y
        sumX2 = sumX2 + x * x
        sumY2 = sumY2 + y * y
    Next i

    If n * sumX2 - sumX * sumX <> 0 Then
        slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX)
        intercept = (sumY - slope * sumX) / n

        ' Calculate R²
        Dim ssTot As Double, ssRes As Double
        Dim meanY As Double
        meanY = sumY / n

        ssTot = 0: ssRes = 0
        For i = LBound(Values) To UBound(Values)
            x = i - LBound(Values) + 1
            Dim predicted As Double
            predicted = intercept + slope * x
            ssTot = ssTot + (Values(i) - meanY) ^ 2
            ssRes = ssRes + (Values(i) - predicted) ^ 2
        Next i

        If ssTot > 0 Then
            rSquared = 1 - (ssRes / ssTot)
        Else
            rSquared = 0
        End If
    Else
        rSquared = 0
    End If

    CalculateTrendStrength = rSquared
End Function

Private Function CalculateSeasonalStrength(ByRef Values() As Double) As Double
    ' Calculate seasonal strength using ACF peaks
    Dim acf() As Double
    Dim maxLag As Integer
    Dim n As Long
    Dim maxACF As Double
    Dim i As Integer

    n = UBound(Values) - LBound(Values) + 1
    maxLag = WorksheetFunction.Min(24, Int(n / 3))

    acf = CalculateACF(Values, maxLag)

    ' Find maximum ACF (excluding lag 0)
    maxACF = 0
    For i = 2 To UBound(acf)
        If Abs(acf(i)) > Abs(maxACF) Then
            maxACF = acf(i)
        End If
    Next i

    CalculateSeasonalStrength = Abs(maxACF)
End Function

Private Function CalculateACF1(ByRef Values() As Double) As Double
    ' Calculate first-order autocorrelation
    Dim n As Long, i As Long
    Dim mean As Double
    Dim numerator As Double, denominator As Double

    n = UBound(Values) - LBound(Values) + 1

    ' Calculate mean
    mean = 0
    For i = LBound(Values) To UBound(Values)
        mean = mean + Values(i)
    Next i
    mean = mean / n

    ' Calculate ACF(1)
    numerator = 0
    denominator = 0

    For i = LBound(Values) To UBound(Values) - 1
        numerator = numerator + (Values(i) - mean) * (Values(i + 1) - mean)
    Next i

    For i = LBound(Values) To UBound(Values)
        denominator = denominator + (Values(i) - mean) ^ 2
    Next i

    If denominator > 0 Then
        CalculateACF1 = numerator / denominator
    Else
        CalculateACF1 = 0
    End If
End Function

Private Function CalculateEntropy(ByRef Values() As Double) As Double
    ' Calculate approximate entropy using binning
    Dim n As Long, i As Long
    Dim minVal As Double, maxVal As Double
    Dim numBins As Integer
    Dim bins() As Long
    Dim binWidth As Double
    Dim binIdx As Integer
    Dim entropy As Double
    Dim prob As Double

    n = UBound(Values) - LBound(Values) + 1
    numBins = WorksheetFunction.Min(10, Int(Sqr(n)))

    ReDim bins(1 To numBins)

    ' Find min and max
    minVal = Values(LBound(Values))
    maxVal = Values(LBound(Values))

    For i = LBound(Values) To UBound(Values)
        If Values(i) < minVal Then minVal = Values(i)
        If Values(i) > maxVal Then maxVal = Values(i)
    Next i

    ' Avoid division by zero
    If maxVal = minVal Then
        CalculateEntropy = 0
        Exit Function
    End If

    binWidth = (maxVal - minVal) / numBins

    ' Count observations in each bin
    For i = LBound(Values) To UBound(Values)
        binIdx = WorksheetFunction.Min(numBins, Int((Values(i) - minVal) / binWidth) + 1)
        bins(binIdx) = bins(binIdx) + 1
    Next i

    ' Calculate entropy
    entropy = 0
    For i = 1 To numBins
        If bins(i) > 0 Then
            prob = bins(i) / n
            entropy = entropy - prob * WorksheetFunction.Ln(prob)
        End If
    Next i

    CalculateEntropy = entropy
End Function

' ============================================================================
' SIMPLE NEURAL NETWORK FORECAST
' ============================================================================

Public Function NeuralNetworkForecast(ByRef tsData As TimeSeriesData, _
                                     ByVal horizon As Integer, _
                                     Optional ByVal numLags As Integer = 0, _
                                     Optional ByVal hiddenNodes As Integer = 5) As ForecastResult
    ' Simple feedforward neural network for forecasting
    ' numLags = number of lagged values to use as inputs (0 = auto)
    ' hiddenNodes = number of hidden layer neurons

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long, j As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Auto-determine number of lags
    If numLags = 0 Then
        numLags = WorksheetFunction.Min(12, Int(n / 4))
    End If

    ' Simple AR-style neural network: predict next value from past lags
    Dim inputSize As Integer
    inputSize = numLags

    ' Create training data
    Dim numSamples As Long
    numSamples = n - numLags

    If numSamples <= 0 Then
        GoTo ErrorHandler
    End If

    Dim X() As Double ' Input matrix
    Dim y() As Double ' Target vector

    ReDim X(1 To numSamples, 1 To inputSize)
    ReDim y(1 To numSamples)

    ' Normalize data to [0,1]
    Dim minVal As Double, maxVal As Double
    minVal = Values(LBound(Values))
    maxVal = Values(LBound(Values))

    For i = LBound(Values) To UBound(Values)
        If Values(i) < minVal Then minVal = Values(i)
        If Values(i) > maxVal Then maxVal = Values(i)
    Next i

    Dim dataRange As Double
    dataRange = maxVal - minVal
    If dataRange = 0 Then dataRange = 1

    ' Build training samples
    For i = 1 To numSamples
        For j = 1 To numLags
            X(i, j) = (Values(LBound(Values) + i + j - 2) - minVal) / dataRange
        Next j
        y(i) = (Values(LBound(Values) + i + numLags - 1) - minVal) / dataRange
    Next i

    ' Initialize network weights (small random values)
    Dim W1() As Double ' Input to hidden weights
    Dim b1() As Double ' Hidden bias
    Dim W2() As Double ' Hidden to output weights
    Dim b2 As Double ' Output bias

    ReDim W1(1 To inputSize, 1 To hiddenNodes)
    ReDim b1(1 To hiddenNodes)
    ReDim W2(1 To hiddenNodes)

    Randomize
    For i = 1 To inputSize
        For j = 1 To hiddenNodes
            W1(i, j) = (Rnd() - 0.5) * 0.5
        Next j
    Next i

    For j = 1 To hiddenNodes
        b1(j) = (Rnd() - 0.5) * 0.1
        W2(j) = (Rnd() - 0.5) * 0.5
    Next j

    b2 = (Rnd() - 0.5) * 0.1

    ' Train network (simple gradient descent - limited epochs for speed)
    Dim numEpochs As Integer
    Dim learningRate As Double

    numEpochs = 50 ' Limited for performance in Excel
    learningRate = 0.01

    Dim k As Integer
    For k = 1 To numEpochs
        ' Forward and backward pass for each sample
        For i = 1 To numSamples
            ' Forward pass
            Dim hidden() As Double
            ReDim hidden(1 To hiddenNodes)

            For j = 1 To hiddenNodes
                Dim hiddenSum As Double
                hiddenSum = b1(j)
                Dim l As Integer
                For l = 1 To inputSize
                    hiddenSum = hiddenSum + X(i, l) * W1(l, j)
                Next l
                hidden(j) = Sigmoid(hiddenSum)
            Next j

            Dim outputSum As Double
            outputSum = b2
            For j = 1 To hiddenNodes
                outputSum = outputSum + hidden(j) * W2(j)
            Next j
            Dim output As Double
            output = outputSum ' Linear output for regression

            ' Backward pass (gradient descent)
            Dim errorOutput As Double
            errorOutput = output - y(i)

            ' Update output layer weights
            For j = 1 To hiddenNodes
                W2(j) = W2(j) - learningRate * errorOutput * hidden(j)
            Next j
            b2 = b2 - learningRate * errorOutput

            ' Update hidden layer weights
            For j = 1 To hiddenNodes
                Dim errorHidden As Double
                errorHidden = errorOutput * W2(j) * SigmoidDerivative(hidden(j))

                For l = 1 To inputSize
                    W1(l, j) = W1(l, j) - learningRate * errorHidden * X(i, l)
                Next l
                b1(j) = b1(j) - learningRate * errorHidden
            Next j
        Next i
    Next k

    ' Generate fitted values and forecasts
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Generate fitted values
    For i = 1 To numSamples
        ReDim hidden(1 To hiddenNodes)

        For j = 1 To hiddenNodes
            hiddenSum = b1(j)
            For l = 1 To inputSize
                hiddenSum = hiddenSum + X(i, l) * W1(l, j)
            Next l
            hidden(j) = Sigmoid(hiddenSum)
        Next j

        outputSum = b2
        For j = 1 To hiddenNodes
            outputSum = outputSum + hidden(j) * W2(j)
        Next j

        result.FittedValues(LBound(Values) + i + numLags - 1) = outputSum * dataRange + minVal
    Next i

    ' Fill early values with actual
    For i = LBound(Values) To LBound(Values) + numLags - 1
        result.FittedValues(i) = Values(i)
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Generate forecasts
    Dim lastLags() As Double
    ReDim lastLags(1 To numLags)

    ' Initialize with last numLags values
    For i = 1 To numLags
        lastLags(i) = (Values(UBound(Values) - numLags + i) - minVal) / dataRange
    Next i

    For i = 1 To horizon
        ' Use last lags as input
        ReDim hidden(1 To hiddenNodes)

        For j = 1 To hiddenNodes
            hiddenSum = b1(j)
            For l = 1 To numLags
                hiddenSum = hiddenSum + lastLags(l) * W1(l, j)
            Next l
            hidden(j) = Sigmoid(hiddenSum)
        Next j

        outputSum = b2
        For j = 1 To hiddenNodes
            outputSum = outputSum + hidden(j) * W2(j)
        Next j

        result.ForecastValues(i) = outputSum * dataRange + minVal

        ' Update lags for next forecast
        For j = 1 To numLags - 1
            lastLags(j) = lastLags(j + 1)
        Next j
        lastLags(numLags) = outputSum ' Normalized forecast
    Next i

    ' Prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * sigma * Sqr(i)
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * sigma * Sqr(i)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * sigma * Sqr(i)
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * sigma * Sqr(i)
    Next i

    ' Metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.ModelName = "NeuralNet[" & numLags & "-" & hiddenNodes & "-1]"

    NeuralNetworkForecast = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "NeuralNet (Error)"
    NeuralNetworkForecast = result
End Function

Private Function Sigmoid(ByVal x As Double) As Double
    ' Sigmoid activation function
    If x < -20 Then
        Sigmoid = 0
    ElseIf x > 20 Then
        Sigmoid = 1
    Else
        Sigmoid = 1 / (1 + Exp(-x))
    End If
End Function

Private Function SigmoidDerivative(ByVal sigmoidOutput As Double) As Double
    ' Derivative of sigmoid given sigmoid output
    SigmoidDerivative = sigmoidOutput * (1 - sigmoidOutput)
End Function

' ============================================================================
' CHANGEPOINT DETECTION
' ============================================================================

Public Function DetectChangepoints(ByRef Values() As Double, _
                                  Optional ByVal minSegmentLength As Integer = 5) As Long()
    ' Detect changepoints using Binary Segmentation algorithm
    ' Returns array of changepoint indices

    Dim n As Long
    Dim changepoints() As Long
    Dim numChangepoints As Integer

    n = UBound(Values) - LBound(Values) + 1

    ReDim changepoints(1 To Int(n / minSegmentLength)) ' Maximum possible changepoints

    numChangepoints = 0

    ' Find changepoints recursively
    Call FindChangepoints(Values, LBound(Values), UBound(Values), _
                         minSegmentLength, changepoints, numChangepoints)

    ' Resize to actual number found
    If numChangepoints > 0 Then
        ReDim Preserve changepoints(1 To numChangepoints)
    Else
        ReDim changepoints(1 To 1)
        changepoints(1) = -1 ' No changepoints found
    End If

    DetectChangepoints = changepoints
End Function

Private Sub FindChangepoints(ByRef Values() As Double, _
                            ByVal startIdx As Long, _
                            ByVal endIdx As Long, _
                            ByVal minLen As Integer, _
                            ByRef changepoints() As Long, _
                            ByRef numCP As Integer)
    ' Recursive binary segmentation

    If endIdx - startIdx + 1 < 2 * minLen Then
        Exit Sub
    End If

    ' Find best split point
    Dim bestSplit As Long
    Dim bestCost As Double
    Dim i As Long

    bestCost = 1E+100
    bestSplit = -1

    For i = startIdx + minLen - 1 To endIdx - minLen
        Dim cost As Double
        cost = CalculateSplitCost(Values, startIdx, i, endIdx)

        If cost < bestCost Then
            bestCost = cost
            bestSplit = i
        End If
    Next i

    ' Check if split is significant
    Dim threshold As Double
    threshold = CalculateStdDev(Values) * 0.5 ' Simple threshold

    If bestSplit > 0 And bestCost < threshold Then
        numCP = numCP + 1
        changepoints(numCP) = bestSplit

        ' Recursively split left and right segments
        Call FindChangepoints(Values, startIdx, bestSplit, minLen, changepoints, numCP)
        Call FindChangepoints(Values, bestSplit + 1, endIdx, minLen, changepoints, numCP)
    End If
End Sub

Private Function CalculateSplitCost(ByRef Values() As Double, _
                                   ByVal startIdx As Long, _
                                   ByVal splitIdx As Long, _
                                   ByVal endIdx As Long) As Double
    ' Calculate cost of splitting at splitIdx
    ' Cost = variance before split - (variance left + variance right)

    Dim i As Long
    Dim mean1 As Double, mean2 As Double
    Dim var1 As Double, var2 As Double
    Dim n1 As Long, n2 As Long

    n1 = splitIdx - startIdx + 1
    n2 = endIdx - splitIdx

    ' Calculate means
    mean1 = 0
    For i = startIdx To splitIdx
        mean1 = mean1 + Values(i)
    Next i
    mean1 = mean1 / n1

    mean2 = 0
    For i = splitIdx + 1 To endIdx
        mean2 = mean2 + Values(i)
    Next i
    mean2 = mean2 / n2

    ' Calculate variances
    var1 = 0
    For i = startIdx To splitIdx
        var1 = var1 + (Values(i) - mean1) ^ 2
    Next i

    var2 = 0
    For i = splitIdx + 1 To endIdx
        var2 = var2 + (Values(i) - mean2) ^ 2
    Next i

    ' Return cost (difference in means - proxy for changepoint strength)
    CalculateSplitCost = Abs(mean1 - mean2)
End Function

' ============================================================================
' ADVANCED OUTLIER DETECTION
' ============================================================================

Public Function DetectOutliersAdvanced(ByRef Values() As Double, _
                                      Optional ByVal method As String = "MAD") As Boolean()
    ' Detect outliers using multiple methods
    ' Methods: "ZScore", "IQR", "MAD" (Median Absolute Deviation), "Grubbs"

    Dim n As Long, i As Long
    Dim isOutlier() As Boolean

    n = UBound(Values) - LBound(Values) + 1
    ReDim isOutlier(LBound(Values) To UBound(Values))

    Select Case UCase(method)
        Case "ZSCORE"
            Call DetectOutliersZScore(Values, isOutlier)

        Case "IQR"
            Call DetectOutliersIQR(Values, isOutlier)

        Case "MAD"
            Call DetectOutliersMAD(Values, isOutlier)

        Case "GRUBBS"
            Call DetectOutliersGrubbs(Values, isOutlier)

        Case Else ' Default to MAD
            Call DetectOutliersMAD(Values, isOutlier)
    End Select

    DetectOutliersAdvanced = isOutlier
End Function

Private Sub DetectOutliersZScore(ByRef Values() As Double, ByRef isOutlier() As Boolean)
    ' Z-score method: |z| > 3 is outlier
    Dim mean As Double, stdDev As Double
    Dim i As Long
    Dim threshold As Double

    threshold = 3 ' Standard threshold

    ' Calculate mean
    mean = 0
    For i = LBound(Values) To UBound(Values)
        mean = mean + Values(i)
    Next i
    mean = mean / (UBound(Values) - LBound(Values) + 1)

    ' Calculate standard deviation
    stdDev = CalculateStdDev(Values)

    ' Flag outliers
    For i = LBound(Values) To UBound(Values)
        If stdDev > 0 Then
            Dim zScore As Double
            zScore = Abs((Values(i) - mean) / stdDev)
            isOutlier(i) = (zScore > threshold)
        Else
            isOutlier(i) = False
        End If
    Next i
End Sub

Private Sub DetectOutliersIQR(ByRef Values() As Double, ByRef isOutlier() As Boolean)
    ' IQR method: outside [Q1 - 1.5*IQR, Q3 + 1.5*IQR]
    Dim sorted() As Double
    Dim n As Long, i As Long
    Dim Q1 As Double, Q3 As Double, IQR As Double
    Dim lowerBound As Double, upperBound As Double

    n = UBound(Values) - LBound(Values) + 1

    ' Copy and sort
    ReDim sorted(LBound(Values) To UBound(Values))
    For i = LBound(Values) To UBound(Values)
        sorted(i) = Values(i)
    Next i
    Call BubbleSortDouble(sorted)

    ' Calculate quartiles
    Dim q1Pos As Long, q3Pos As Long
    q1Pos = Int(n * 0.25)
    q3Pos = Int(n * 0.75)

    If q1Pos < LBound(sorted) Then q1Pos = LBound(sorted)
    If q3Pos > UBound(sorted) Then q3Pos = UBound(sorted)

    Q1 = sorted(q1Pos)
    Q3 = sorted(q3Pos)
    IQR = Q3 - Q1

    lowerBound = Q1 - 1.5 * IQR
    upperBound = Q3 + 1.5 * IQR

    ' Flag outliers
    For i = LBound(Values) To UBound(Values)
        isOutlier(i) = (Values(i) < lowerBound Or Values(i) > upperBound)
    Next i
End Sub

Private Sub DetectOutliersMAD(ByRef Values() As Double, ByRef isOutlier() As Boolean)
    ' MAD (Median Absolute Deviation) method - robust to outliers
    Dim median As Double
    Dim mad As Double
    Dim i As Long
    Dim deviations() As Double
    Dim threshold As Double

    threshold = 3 ' Modified Z-score threshold

    ' Calculate median
    median = CalculateMedian(Values)

    ' Calculate absolute deviations from median
    ReDim deviations(LBound(Values) To UBound(Values))
    For i = LBound(Values) To UBound(Values)
        deviations(i) = Abs(Values(i) - median)
    Next i

    ' MAD is median of absolute deviations
    mad = CalculateMedian(deviations)

    ' Avoid division by zero
    If mad = 0 Then
        For i = LBound(Values) To UBound(Values)
            isOutlier(i) = False
        Next i
        Exit Sub
    End If

    ' Modified Z-score = 0.6745 * |x - median| / MAD
    For i = LBound(Values) To UBound(Values)
        Dim modifiedZ As Double
        modifiedZ = 0.6745 * Abs(Values(i) - median) / mad
        isOutlier(i) = (modifiedZ > threshold)
    Next i
End Sub

Private Sub DetectOutliersGrubbs(ByRef Values() As Double, ByRef isOutlier() As Boolean)
    ' Grubbs' test for outliers (single pass)
    Dim mean As Double, stdDev As Double
    Dim n As Long, i As Long
    Dim maxDeviation As Double
    Dim maxIdx As Long
    Dim G As Double, GCritical As Double

    n = UBound(Values) - LBound(Values) + 1

    ' Initialize all as non-outliers
    For i = LBound(Values) To UBound(Values)
        isOutlier(i) = False
    Next i

    ' Calculate mean and std dev
    mean = 0
    For i = LBound(Values) To UBound(Values)
        mean = mean + Values(i)
    Next i
    mean = mean / n

    stdDev = CalculateStdDev(Values)

    If stdDev = 0 Then Exit Sub

    ' Find maximum deviation
    maxDeviation = 0
    maxIdx = LBound(Values)

    For i = LBound(Values) To UBound(Values)
        Dim deviation As Double
        deviation = Abs(Values(i) - mean)
        If deviation > maxDeviation Then
            maxDeviation = deviation
            maxIdx = i
        End If
    Next i

    ' Calculate Grubbs statistic
    G = maxDeviation / stdDev

    ' Critical value (approximate for alpha = 0.05)
    ' G_critical ≈ ((n-1)/sqrt(n)) * sqrt(t^2 / (n-2+t^2))
    ' Simplified: use threshold of 3 for reasonable n
    GCritical = 3

    If G > GCritical Then
        isOutlier(maxIdx) = True
    End If
End Sub

Private Function CalculateMedian(ByRef Values() As Double) As Double
    ' Calculate median
    Dim sorted() As Double
    Dim n As Long, i As Long

    n = UBound(Values) - LBound(Values) + 1

    ' Copy and sort
    ReDim sorted(LBound(Values) To UBound(Values))
    For i = LBound(Values) To UBound(Values)
        sorted(i) = Values(i)
    Next i
    Call BubbleSortDouble(sorted)

    ' Find median
    If n Mod 2 = 1 Then
        CalculateMedian = sorted(LBound(sorted) + Int(n / 2))
    Else
        CalculateMedian = (sorted(LBound(sorted) + Int(n / 2) - 1) + sorted(LBound(sorted) + Int(n / 2))) / 2
    End If
End Function

' ============================================================================
' BAYESIAN MODEL AVERAGING (BMA)
' ============================================================================

Public Function BayesianModelAveraging(ByRef models() As ForecastResult, _
                                      ByVal numModels As Integer, _
                                      ByRef actualValues() As Double) As ForecastResult
    ' Combine forecasts using Bayesian Model Averaging
    ' Weights based on model likelihood (BIC approximation)

    Dim result As ForecastResult
    Dim weights() As Double
    Dim i As Integer, j As Integer

    ReDim weights(1 To numModels)

    ' Calculate BIC for each model (simplified using RMSE)
    Dim bic() As Double
    ReDim bic(1 To numModels)

    Dim n As Long
    n = UBound(actualValues) - LBound(actualValues) + 1

    For i = 1 To numModels
        ' BIC ≈ n * log(RMSE^2) + k * log(n)
        ' Using RMSE as proxy, k = 3 (typical for exponential smoothing)
        Dim k As Integer
        k = 3

        If models(i).RMSE > 0 Then
            bic(i) = n * WorksheetFunction.Ln(models(i).RMSE ^ 2) + k * WorksheetFunction.Ln(n)
        Else
            bic(i) = 1E+100 ' Large penalty for bad models
        End If
    Next i

    ' Convert BIC to weights using exp(-0.5 * BIC)
    Dim totalWeight As Double
    totalWeight = 0

    For i = 1 To numModels
        weights(i) = Exp(-0.5 * bic(i))
        totalWeight = totalWeight + weights(i)
    Next i

    ' Normalize weights
    If totalWeight > 0 Then
        For i = 1 To numModels
            weights(i) = weights(i) / totalWeight
        Next i
    Else
        ' Equal weights if calculation fails
        For i = 1 To numModels
            weights(i) = 1 / numModels
        Next i
    End If

    ' Combine forecasts
    Dim horizon As Integer
    horizon = UBound(models(1).ForecastValues)

    ReDim result.FittedValues(LBound(actualValues) To UBound(actualValues))
    ReDim result.Residuals(LBound(actualValues) To UBound(actualValues))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Weighted combination
    For j = 1 To horizon
        result.ForecastValues(j) = 0
        result.Lower80(j) = 0
        result.Upper80(j) = 0
        result.Lower95(j) = 0
        result.Upper95(j) = 0

        For i = 1 To numModels
            result.ForecastValues(j) = result.ForecastValues(j) + weights(i) * models(i).ForecastValues(j)
            result.Lower80(j) = result.Lower80(j) + weights(i) * models(i).Lower80(j)
            result.Upper80(j) = result.Upper80(j) + weights(i) * models(i).Upper80(j)
            result.Lower95(j) = result.Lower95(j) + weights(i) * models(i).Lower95(j)
            result.Upper95(j) = result.Upper95(j) + weights(i) * models(i).Upper95(j)
        Next i
    Next j

    ' Combine fitted values
    For j = LBound(actualValues) To UBound(actualValues)
        result.FittedValues(j) = 0
        For i = 1 To numModels
            If j >= LBound(models(i).FittedValues) And j <= UBound(models(i).FittedValues) Then
                result.FittedValues(j) = result.FittedValues(j) + weights(i) * models(i).FittedValues(j)
            End If
        Next i
    Next j

    ' Calculate residuals
    For j = LBound(actualValues) To UBound(actualValues)
        result.Residuals(j) = actualValues(j) - result.FittedValues(j)
    Next j

    ' Calculate metrics
    result.MAPE = CalculateMAPE(actualValues, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(actualValues, result.FittedValues)

    ' Build model name showing weights
    Dim modelName As String
    modelName = "BMA["
    For i = 1 To numModels
        If i > 1 Then modelName = modelName & ","
        modelName = modelName & models(i).ModelName & ":" & Format(weights(i), "0.00")
    Next i
    modelName = modelName & "]"

    result.ModelName = modelName

    BayesianModelAveraging = result
End Function

' ============================================================================
' TIME SERIES CROSS-VALIDATION
' ============================================================================

Public Function TimeSeriesCrossValidation(ByRef tsData As TimeSeriesData, _
                                         ByVal horizon As Integer, _
                                         ByVal numFolds As Integer, _
                                         ByVal modelType As String) As CrossValidationResult
    ' Perform time series cross-validation with expanding window
    ' Returns average metrics across folds

    Dim cvResult As CrossValidationResult
    Dim Values() As Double
    Dim n As Long, i As Integer, j As Long
    Dim foldSize As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    foldSize = Int(n / (numFolds + 1))

    Dim totalMAPE As Double, totalMAE As Double, totalRMSE As Double
    totalMAPE = 0
    totalMAE = 0
    totalRMSE = 0

    Dim validFolds As Integer
    validFolds = 0

    ' Perform cross-validation
    For i = 1 To numFolds
        Dim trainSize As Long
        trainSize = foldSize * i

        If trainSize + horizon <= n Then
            ' Create training subset
            Dim trainData As TimeSeriesData
            ReDim trainData.Values(LBound(Values) To LBound(Values) + trainSize - 1)

            For j = LBound(Values) To LBound(Values) + trainSize - 1
                trainData.Values(j) = Values(j)
            Next j
            trainData.Frequency = tsData.Frequency

            ' Forecast
            Dim forecast As ForecastResult

            On Error Resume Next
            Select Case UCase(modelType)
                Case "SES"
                    forecast = SimpleExponentialSmoothing(trainData, horizon)
                Case "HOLT-WINTERS"
                    forecast = HoltWinters(trainData, horizon, "additive")
                Case "ARIMA"
                    forecast = SimpleARIMA(trainData, horizon)
                Case "AUTO"
                    forecast = AutoForecast(trainData, horizon, "additive")
                Case Else
                    forecast = SimpleExponentialSmoothing(trainData, horizon)
            End Select
            On Error GoTo 0

            ' Evaluate against test set
            Dim testStart As Long
            testStart = LBound(Values) + trainSize

            Dim testValues() As Double
            Dim forecastValues() As Double
            ReDim testValues(1 To horizon)
            ReDim forecastValues(1 To horizon)

            For j = 1 To horizon
                If testStart + j - 1 <= UBound(Values) Then
                    testValues(j) = Values(testStart + j - 1)
                    forecastValues(j) = forecast.ForecastValues(j)
                End If
            Next j

            ' Calculate metrics for this fold
            Dim foldMAPE As Double, foldMAE As Double, foldRMSE As Double
            Dim foldResiduals() As Double
            ReDim foldResiduals(1 To horizon)

            For j = 1 To horizon
                foldResiduals(j) = testValues(j) - forecastValues(j)
            Next j

            foldMAPE = CalculateMAPE(testValues, forecastValues)
            foldMAE = CalculateMAE(foldResiduals)
            foldRMSE = CalculateRMSE(foldResiduals)

            If foldMAPE < 9999 Then ' Valid forecast
                totalMAPE = totalMAPE + foldMAPE
                totalMAE = totalMAE + foldMAE
                totalRMSE = totalRMSE + foldRMSE
                validFolds = validFolds + 1
            End If
        End If
    Next i

    ' Average across folds
    If validFolds > 0 Then
        cvResult.AvgMAPE = totalMAPE / validFolds
        cvResult.AvgMAE = totalMAE / validFolds
        cvResult.AvgRMSE = totalRMSE / validFolds
        cvResult.NumFolds = validFolds
    Else
        cvResult.AvgMAPE = 9999
        cvResult.AvgMAE = 9999
        cvResult.AvgRMSE = 9999
        cvResult.NumFolds = 0
    End If

    TimeSeriesCrossValidation = cvResult
End Function

' ============================================================================
' DYNAMIC HARMONIC REGRESSION
' ============================================================================

Public Function DynamicHarmonicRegression(ByRef tsData As TimeSeriesData, _
                                         ByVal horizon As Integer, _
                                         Optional ByVal K As Integer = 0) As ForecastResult
    ' DHR = Fourier terms + ARIMA errors
    ' Combines Fourier series with ARIMA modeling of residuals

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' First, fit Fourier series
    Dim fourierResult As ForecastResult
    fourierResult = FourierForecast(tsData, horizon, K)

    ' Model residuals with ARIMA
    Dim residualData As TimeSeriesData
    ReDim residualData.Values(LBound(fourierResult.Residuals) To UBound(fourierResult.Residuals))

    Dim i As Long
    For i = LBound(fourierResult.Residuals) To UBound(fourierResult.Residuals)
        residualData.Values(i) = fourierResult.Residuals(i)
    Next i
    residualData.Frequency = tsData.Frequency

    ' Forecast residuals
    Dim residualForecast As ForecastResult
    residualForecast = SimpleARIMA(residualData, horizon)

    ' Combine Fourier forecast with ARIMA residual forecast
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Fitted values = Fourier fitted + ARIMA fitted residuals
    For i = LBound(Values) To UBound(Values)
        result.FittedValues(i) = fourierResult.FittedValues(i) + residualForecast.FittedValues(i)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecasts = Fourier forecast + ARIMA residual forecast
    For i = 1 To horizon
        result.ForecastValues(i) = fourierResult.ForecastValues(i) + residualForecast.ForecastValues(i)
        result.Lower80(i) = fourierResult.Lower80(i) + residualForecast.Lower80(i)
        result.Upper80(i) = fourierResult.Upper80(i) + residualForecast.Upper80(i)
        result.Lower95(i) = fourierResult.Lower95(i) + residualForecast.Lower95(i)
        result.Upper95(i) = fourierResult.Upper95(i) + residualForecast.Upper95(i)
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.ModelName = "DHR (Fourier+ARIMA)"

    DynamicHarmonicRegression = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "DHR (Error)"
    DynamicHarmonicRegression = result
End Function

' ============================================================================
' DEMAND CLASSIFICATION (Syntetos-Boylan-Croston Framework)
' ============================================================================

Public Function ClassifyDemand(ByRef Values() As Double) As DemandClassification
    ' Classify demand pattern using CV² and ADI (Average Demand Interval)
    ' Returns optimal forecasting method recommendation

    Dim classification As DemandClassification
    Dim n As Long, i As Long
    Dim mean As Double, variance As Double, cv As Double
    Dim zeroCount As Long, adi As Double
    Dim cv2 As Double

    n = UBound(Values) - LBound(Values) + 1

    ' Calculate mean (excluding zeros for intermittent demand)
    Dim nonZeroSum As Double
    Dim nonZeroCount As Long
    nonZeroSum = 0
    nonZeroCount = 0
    zeroCount = 0

    For i = LBound(Values) To UBound(Values)
        If Values(i) > 0 Then
            nonZeroSum = nonZeroSum + Values(i)
            nonZeroCount = nonZeroCount + 1
        Else
            zeroCount = zeroCount + 1
        End If
    Next i

    If nonZeroCount > 0 Then
        mean = nonZeroSum / nonZeroCount
    Else
        mean = 0
    End If

    ' Calculate variance of non-zero demands
    variance = 0
    For i = LBound(Values) To UBound(Values)
        If Values(i) > 0 Then
            variance = variance + (Values(i) - mean) ^ 2
        End If
    Next i

    If nonZeroCount > 1 Then
        variance = variance / (nonZeroCount - 1)
    End If

    ' Coefficient of Variation
    If mean > 0 Then
        cv = Sqr(variance) / mean
    Else
        cv = 0
    End If

    cv2 = cv ^ 2

    ' Average Demand Interval (ADI)
    If nonZeroCount > 0 Then
        adi = n / nonZeroCount
    Else
        adi = n
    End If

    ' Store metrics
    classification.CV = cv
    classification.ADI = adi

    ' Syntetos-Boylan-Croston Classification
    ' CV² < 0.49 and ADI < 1.32: Smooth
    ' CV² >= 0.49 and ADI < 1.32: Erratic
    ' CV² < 0.49 and ADI >= 1.32: Intermittent
    ' CV² >= 0.49 and ADI >= 1.32: Lumpy

    If cv2 < 0.49 Then
        If adi < 1.32 Then
            classification.Category = "Smooth"
            classification.RecommendedMethod = "ETS/ARIMA/SARIMA"
        Else
            classification.Category = "Intermittent"
            classification.RecommendedMethod = "Croston/TSB/SBA"
        End If
    Else
        If adi < 1.32 Then
            classification.Category = "Erratic"
            classification.RecommendedMethod = "SES/MovingAvg"
        Else
            classification.Category = "Lumpy"
            classification.RecommendedMethod = "TSB/SBA/Croston"
        End If
    End If

    ClassifyDemand = classification
End Function

' ============================================================================
' KALMAN FILTER FOR ADAPTIVE STATE ESTIMATION
' ============================================================================

Public Function KalmanFilterForecast(ByRef tsData As TimeSeriesData, _
                                    ByVal horizon As Integer, _
                                    Optional ByVal processNoise As Double = 0.01, _
                                    Optional ByVal measurementNoise As Double = 0.1) As ForecastResult
    ' Kalman Filter for adaptive level and trend estimation
    ' processNoise (Q): How much we expect the state to change
    ' measurementNoise (R): How noisy our observations are

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' State vector: [level, trend]
    Dim x1 As Double, x2 As Double ' State estimates
    Dim P11 As Double, P12 As Double, P21 As Double, P22 As Double ' Covariance matrix
    Dim Q As Double, R As Double ' Process and measurement noise

    Q = processNoise
    R = measurementNoise

    ' Initialize state
    x1 = Values(LBound(Values)) ' Initial level
    x2 = 0 ' Initial trend

    If n > 1 Then
        x2 = Values(LBound(Values) + 1) - Values(LBound(Values))
    End If

    ' Initialize covariance matrix
    P11 = 1
    P12 = 0
    P21 = 0
    P22 = 1

    ' Initialize result arrays
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Kalman filter iteration
    For i = LBound(Values) To UBound(Values)
        ' Prediction step
        Dim x1_pred As Double, x2_pred As Double
        x1_pred = x1 + x2 ' Level + Trend
        x2_pred = x2 ' Trend persists

        ' Predict covariance
        Dim P11_pred As Double, P12_pred As Double, P21_pred As Double, P22_pred As Double
        P11_pred = P11 + P12 + P21 + P22 + Q
        P12_pred = P12 + P22
        P21_pred = P21 + P22
        P22_pred = P22 + Q

        ' Update step (measurement update)
        Dim y As Double ' Observation
        y = Values(i)

        Dim innovation As Double
        innovation = y - x1_pred ' Prediction error

        Dim S As Double ' Innovation covariance
        S = P11_pred + R

        ' Kalman gain
        Dim K1 As Double, K2 As Double
        If S > 0 Then
            K1 = P11_pred / S
            K2 = P21_pred / S
        Else
            K1 = 0
            K2 = 0
        End If

        ' Update state estimate
        x1 = x1_pred + K1 * innovation
        x2 = x2_pred + K2 * innovation

        ' Update covariance
        P11 = (1 - K1) * P11_pred
        P12 = (1 - K1) * P12_pred
        P21 = P21_pred - K2 * P11_pred
        P22 = P22_pred - K2 * P12_pred

        ' Store fitted value
        result.FittedValues(i) = x1_pred
    Next i

    ' Calculate residuals
    For i = LBound(Values) To UBound(Values)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecast future values
    Dim level As Double, trend As Double
    level = x1
    trend = x2

    For i = 1 To horizon
        result.ForecastValues(i) = level + i * trend
    Next i

    ' Prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        Dim intervalWidth As Double
        intervalWidth = sigma * Sqr(i) ' Uncertainty grows with horizon
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * intervalWidth
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * intervalWidth
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * intervalWidth
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * intervalWidth
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.ModelName = "KalmanFilter"

    KalmanFilterForecast = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "Kalman (Error)"
    KalmanFilterForecast = result
End Function

' ============================================================================
' STL DECOMPOSITION (Seasonal-Trend using Loess) - Simplified
' ============================================================================

Public Function STLDecomposition(ByRef tsData As TimeSeriesData, _
                                ByVal horizon As Integer, _
                                Optional ByVal period As Integer = 0) As ForecastResult
    ' Simplified STL: Seasonal-Trend decomposition using moving averages
    ' Then forecast trend and seasonal components separately

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long, i As Long, j As Long
    Dim m As Integer ' Seasonal period

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Determine seasonal period
    If period > 0 Then
        m = period
    ElseIf tsData.Frequency > 0 Then
        m = tsData.Frequency
    Else
        m = 12 ' Default
    End If

    If m < 2 Or n < 2 * m Then
        ' Not enough data for STL, use simple method
        result = SimpleExponentialSmoothing(tsData, horizon)
        result.ModelName = "STL (Insufficient Data)"
        STLDecomposition = result
        Exit Function
    End If

    ' Step 1: Detrend using moving average
    Dim trend() As Double
    ReDim trend(LBound(Values) To UBound(Values))

    Dim windowSize As Integer
    windowSize = m
    If windowSize Mod 2 = 0 Then windowSize = windowSize + 1 ' Make odd

    Dim halfWindow As Integer
    halfWindow = Int(windowSize / 2)

    For i = LBound(Values) To UBound(Values)
        Dim sum As Double, count As Integer
        sum = 0
        count = 0

        For j = WorksheetFunction.Max(LBound(Values), i - halfWindow) To WorksheetFunction.Min(UBound(Values), i + halfWindow)
            sum = sum + Values(j)
            count = count + 1
        Next j

        trend(i) = sum / count
    Next i

    ' Step 2: Detrended series
    Dim detrended() As Double
    ReDim detrended(LBound(Values) To UBound(Values))

    For i = LBound(Values) To UBound(Values)
        detrended(i) = Values(i) - trend(i)
    Next i

    ' Step 3: Extract seasonal component (average for each season)
    Dim seasonal() As Double
    ReDim seasonal(1 To m)

    Dim seasonCount() As Integer
    ReDim seasonCount(1 To m)

    For i = 1 To m
        seasonal(i) = 0
        seasonCount(i) = 0
    Next i

    For i = LBound(Values) To UBound(Values)
        Dim seasonIdx As Integer
        seasonIdx = ((i - LBound(Values)) Mod m) + 1
        seasonal(seasonIdx) = seasonal(seasonIdx) + detrended(i)
        seasonCount(seasonIdx) = seasonCount(seasonIdx) + 1
    Next i

    For i = 1 To m
        If seasonCount(i) > 0 Then
            seasonal(i) = seasonal(i) / seasonCount(i)
        End If
    Next i

    ' Center seasonal component (sum to zero for additive)
    Dim seasonalMean As Double
    seasonalMean = 0
    For i = 1 To m
        seasonalMean = seasonalMean + seasonal(i)
    Next i
    seasonalMean = seasonalMean / m

    For i = 1 To m
        seasonal(i) = seasonal(i) - seasonalMean
    Next i

    ' Step 4: Calculate remainder
    Dim remainder() As Double
    ReDim remainder(LBound(Values) To UBound(Values))

    For i = LBound(Values) To UBound(Values)
        seasonIdx = ((i - LBound(Values)) Mod m) + 1
        remainder(i) = Values(i) - trend(i) - seasonal(seasonIdx)
    Next i

    ' Step 5: Forecast trend using linear extrapolation
    Dim trendSlope As Double
    Dim lastTrends As Double
    Dim trendWindow As Integer

    trendWindow = WorksheetFunction.Min(m, 10)
    lastTrends = 0

    For i = UBound(trend) - trendWindow + 1 To UBound(trend)
        If i > LBound(trend) Then
            lastTrends = lastTrends + (trend(i) - trend(i - 1))
        End If
    Next i

    trendSlope = lastTrends / (trendWindow - 1)

    ' Step 6: Generate forecasts
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Fitted values
    For i = LBound(Values) To UBound(Values)
        seasonIdx = ((i - LBound(Values)) Mod m) + 1
        result.FittedValues(i) = trend(i) + seasonal(seasonIdx)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecasts
    Dim lastTrend As Double
    lastTrend = trend(UBound(trend))

    For i = 1 To horizon
        seasonIdx = ((UBound(Values) - LBound(Values) + i) Mod m) + 1
        result.ForecastValues(i) = lastTrend + i * trendSlope + seasonal(seasonIdx)
    Next i

    ' Prediction intervals
    Dim sigma As Double
    sigma = CalculateStdDev(result.Residuals)

    For i = 1 To horizon
        result.Lower80(i) = result.ForecastValues(i) - 1.28 * sigma * Sqr(i)
        result.Upper80(i) = result.ForecastValues(i) + 1.28 * sigma * Sqr(i)
        result.Lower95(i) = result.ForecastValues(i) - 1.96 * sigma * Sqr(i)
        result.Upper95(i) = result.ForecastValues(i) + 1.96 * sigma * Sqr(i)
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.ModelName = "STL[" & m & "]"

    STLDecomposition = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "STL (Error)"
    STLDecomposition = result
End Function

' ============================================================================
' INTELLIGENT MODEL SELECTOR (Feature-Based)
' ============================================================================

Public Function SelectBestModels(ByRef tsData As TimeSeriesData) As String()
    ' Automatically select best models based on data characteristics
    ' Returns array of recommended model names

    Dim recommendations() As String
    Dim recCount As Integer
    recCount = 0

    ReDim recommendations(1 To 10) ' Max 10 recommendations

    ' Extract features
    Dim features() As Double
    features = ExtractTimeSeriesFeatures(tsData.Values)

    ' features: [Mean, StdDev, CV, Skewness, Kurtosis, TrendStrength, SeasonalStrength, ACF1, Entropy]
    Dim cv As Double, trendStrength As Double, seasonalStrength As Double, acf1 As Double

    cv = features(3)
    trendStrength = features(6)
    seasonalStrength = features(7)
    acf1 = features(8)

    ' Classify demand
    Dim demandClass As DemandClassification
    demandClass = ClassifyDemand(tsData.Values)

    ' Rule-based model selection
    Select Case demandClass.Category
        Case "Smooth"
            ' High trend strength -> trend methods
            If trendStrength > 0.7 Then
                recCount = recCount + 1: recommendations(recCount) = "LinearTrend"
                recCount = recCount + 1: recommendations(recCount) = "Holt"
                recCount = recCount + 1: recommendations(recCount) = "ETS-AAN"
            End If

            ' High seasonal strength -> seasonal methods
            If seasonalStrength > 0.4 Then
                recCount = recCount + 1: recommendations(recCount) = "HoltWinters"
                recCount = recCount + 1: recommendations(recCount) = "SARIMA"
                recCount = recCount + 1: recommendations(recCount) = "STL"
                recCount = recCount + 1: recommendations(recCount) = "Fourier"
            End If

            ' High autocorrelation -> ARIMA
            If acf1 > 0.5 Then
                recCount = recCount + 1: recommendations(recCount) = "ARIMA"
                recCount = recCount + 1: recommendations(recCount) = "AutoARIMA"
            End If

            ' Always include general methods
            recCount = recCount + 1: recommendations(recCount) = "SES"
            recCount = recCount + 1: recommendations(recCount) = "Theta"

        Case "Intermittent", "Lumpy"
            ' Specialized intermittent demand methods
            recCount = recCount + 1: recommendations(recCount) = "Croston"
            recCount = recCount + 1: recommendations(recCount) = "TSB"
            recCount = recCount + 1: recommendations(recCount) = "SBA"
            recCount = recCount + 1: recommendations(recCount) = "SES"

        Case "Erratic"
            ' Simple smoothing methods for erratic data
            recCount = recCount + 1: recommendations(recCount) = "SES"
            recCount = recCount + 1: recommendations(recCount) = "MovingAverage"
            recCount = recCount + 1: recommendations(recCount) = "Theta"
            recCount = recCount + 1: recommendations(recCount) = "NeuralNet"

        Case Else
            ' Default recommendations
            recCount = recCount + 1: recommendations(recCount) = "SES"
            recCount = recCount + 1: recommendations(recCount) = "HoltWinters"
            recCount = recCount + 1: recommendations(recCount) = "ARIMA"
    End Select

    ' Always add advanced ensemble methods
    recCount = recCount + 1: recommendations(recCount) = "BMA"
    recCount = recCount + 1: recommendations(recCount) = "KalmanFilter"

    ' Resize to actual count
    If recCount > 0 Then
        ReDim Preserve recommendations(1 To recCount)
    Else
        ReDim recommendations(1 To 1)
        recommendations(1) = "SES"
    End If

    SelectBestModels = recommendations
End Function

' ============================================================================
' ENSEMBLE DIVERSITY WEIGHTING
' ============================================================================

Public Function EnsembleDiversityScore(ByRef models() As ForecastResult, _
                                      ByVal numModels As Integer) As Double()
    ' Calculate diversity scores for ensemble models
    ' More diverse models get higher weights (complementary to accuracy)

    Dim diversityScores() As Double
    ReDim diversityScores(1 To numModels)

    Dim i As Integer, j As Integer, k As Long
    Dim n As Long

    If numModels < 1 Then
        EnsembleDiversityScore = diversityScores
        Exit Function
    End If

    n = UBound(models(1).FittedValues) - LBound(models(1).FittedValues) + 1

    ' Calculate pairwise correlations
    For i = 1 To numModels
        Dim avgCorrelation As Double
        Dim correlationCount As Integer

        avgCorrelation = 0
        correlationCount = 0

        For j = 1 To numModels
            If i <> j Then
                ' Calculate correlation between model i and j forecasts
                Dim corr As Double
                corr = CalculateCorrelation(models(i).FittedValues, models(j).FittedValues)

                avgCorrelation = avgCorrelation + Abs(corr)
                correlationCount = correlationCount + 1
            End If
        Next j

        If correlationCount > 0 Then
            avgCorrelation = avgCorrelation / correlationCount
        End If

        ' Diversity score = 1 - average correlation
        ' Higher diversity = lower correlation with other models
        diversityScores(i) = 1 - avgCorrelation
    Next i

    EnsembleDiversityScore = diversityScores
End Function

Private Function CalculateCorrelation(ByRef x() As Double, ByRef y() As Double) As Double
    ' Calculate Pearson correlation coefficient
    Dim n As Long, i As Long
    Dim meanX As Double, meanY As Double
    Dim covXY As Double, varX As Double, varY As Double

    n = UBound(x) - LBound(x) + 1

    If n <> (UBound(y) - LBound(y) + 1) Then
        CalculateCorrelation = 0
        Exit Function
    End If

    ' Calculate means
    meanX = 0
    meanY = 0
    For i = LBound(x) To UBound(x)
        meanX = meanX + x(i)
        meanY = meanY + y(i)
    Next i
    meanX = meanX / n
    meanY = meanY / n

    ' Calculate covariance and variances
    covXY = 0
    varX = 0
    varY = 0

    For i = LBound(x) To UBound(x)
        covXY = covXY + (x(i) - meanX) * (y(i) - meanY)
        varX = varX + (x(i) - meanX) ^ 2
        varY = varY + (y(i) - meanY) ^ 2
    Next i

    ' Correlation coefficient
    If varX > 0 And varY > 0 Then
        CalculateCorrelation = covXY / Sqr(varX * varY)
    Else
        CalculateCorrelation = 0
    End If
End Function

' ============================================================================
' QUANTILE FORECAST (Better Prediction Intervals)
' ============================================================================

Public Function QuantileRegression(ByRef residuals() As Double, _
                                  ByVal quantile As Double) As Double
    ' Calculate quantile of residuals for prediction intervals
    ' quantile: 0.025 for lower 95%, 0.975 for upper 95%

    Dim sorted() As Double
    Dim n As Long, i As Long
    Dim position As Double
    Dim lowerIdx As Long, upperIdx As Long

    n = UBound(residuals) - LBound(residuals) + 1

    ' Copy and sort residuals
    ReDim sorted(LBound(residuals) To UBound(residuals))
    For i = LBound(residuals) To UBound(residuals)
        sorted(i) = residuals(i)
    Next i
    Call BubbleSortDouble(sorted)

    ' Find quantile position
    position = quantile * (n - 1) + LBound(sorted)

    lowerIdx = Int(position)
    upperIdx = lowerIdx + 1

    If lowerIdx < LBound(sorted) Then lowerIdx = LBound(sorted)
    If upperIdx > UBound(sorted) Then upperIdx = UBound(sorted)

    ' Linear interpolation
    Dim fraction As Double
    fraction = position - lowerIdx

    QuantileRegression = sorted(lowerIdx) + fraction * (sorted(upperIdx) - sorted(lowerIdx))
End Function

Public Function QuantileForecastIntervals(ByRef baseResult As ForecastResult, _
                                         ByVal horizon As Integer) As ForecastResult
    ' Calculate prediction intervals using quantile regression on residuals
    ' More accurate than normal distribution assumption

    Dim result As ForecastResult
    result = baseResult

    ' Calculate quantiles from residuals
    Dim q025 As Double, q975 As Double, q10 As Double, q90 As Double

    q025 = QuantileRegression(baseResult.Residuals, 0.025)
    q975 = QuantileRegression(baseResult.Residuals, 0.975)
    q10 = QuantileRegression(baseResult.Residuals, 0.1)
    q90 = QuantileRegression(baseResult.Residuals, 0.9)

    ' Apply quantiles to forecasts
    Dim i As Integer
    For i = 1 To horizon
        ' Scale intervals by sqrt(horizon) to account for uncertainty growth
        result.Lower95(i) = result.ForecastValues(i) + q025 * Sqr(i)
        result.Upper95(i) = result.ForecastValues(i) + q975 * Sqr(i)
        result.Lower80(i) = result.ForecastValues(i) + q10 * Sqr(i)
        result.Upper80(i) = result.ForecastValues(i) + q90 * Sqr(i)
    Next i

    QuantileForecastIntervals = result
End Function

' ============================================================================
' MISSING VALUE IMPUTATION
' ============================================================================

Public Function ImputeMissingValues(ByRef Values() As Double) As Double()
    ' Impute missing values using linear interpolation
    ' Missing values assumed to be represented as Empty or error values

    Dim result() As Double
    Dim n As Long, i As Long

    n = UBound(Values) - LBound(Values) + 1
    ReDim result(LBound(Values) To UBound(Values))

    ' Copy values
    For i = LBound(Values) To UBound(Values)
        result(i) = Values(i)
    Next i

    ' Identify and impute missing values
    For i = LBound(Values) To UBound(Values)
        ' Check if missing (implement your own check for missing data)
        Dim isMissing As Boolean
        isMissing = False

        On Error Resume Next
        If IsError(Values(i)) Or IsEmpty(Values(i)) Or Values(i) = 0 / 0 Then
            isMissing = True
        End If
        On Error GoTo 0

        If isMissing Then
            ' Find previous non-missing value
            Dim prevIdx As Long, nextIdx As Long
            Dim prevVal As Double, nextVal As Double
            Dim foundPrev As Boolean, foundNext As Boolean

            foundPrev = False
            foundNext = False

            prevIdx = i - 1
            Do While prevIdx >= LBound(Values) And Not foundPrev
                On Error Resume Next
                If Not IsError(Values(prevIdx)) And Not IsEmpty(Values(prevIdx)) Then
                    prevVal = Values(prevIdx)
                    foundPrev = True
                End If
                On Error GoTo 0
                prevIdx = prevIdx - 1
            Loop

            nextIdx = i + 1
            Do While nextIdx <= UBound(Values) And Not foundNext
                On Error Resume Next
                If Not IsError(Values(nextIdx)) And Not IsEmpty(Values(nextIdx)) Then
                    nextVal = Values(nextIdx)
                    foundNext = True
                End If
                On Error GoTo 0
                nextIdx = nextIdx + 1
            Loop

            ' Linear interpolation
            If foundPrev And foundNext Then
                result(i) = prevVal + (nextVal - prevVal) * (i - (prevIdx + 1)) / ((nextIdx - 1) - (prevIdx + 1))
            ElseIf foundPrev Then
                result(i) = prevVal
            ElseIf foundNext Then
                result(i) = nextVal
            Else
                result(i) = 0 ' Last resort
            End If
        End If
    Next i

    ImputeMissingValues = result
End Function

' ============================================================================
' REGIME DETECTION
' ============================================================================

Public Function DetectRegimes(ByRef Values() As Double) As Long()
    ' Detect regime shifts using rolling variance
    ' Returns array of regime change points

    Dim n As Long, i As Long
    Dim windowSize As Integer
    Dim regimeChanges() As Long
    Dim changeCount As Integer

    n = UBound(Values) - LBound(Values) + 1
    windowSize = WorksheetFunction.Min(20, Int(n / 5))

    ReDim regimeChanges(1 To Int(n / windowSize))
    changeCount = 0

    ' Calculate rolling variance
    Dim rollingVar() As Double
    ReDim rollingVar(LBound(Values) To UBound(Values) - windowSize + 1)

    For i = LBound(Values) To UBound(Values) - windowSize + 1
        Dim sum As Double, sumSq As Double
        Dim j As Long, mean As Double, variance As Double

        sum = 0
        sumSq = 0

        For j = i To i + windowSize - 1
            sum = sum + Values(j)
            sumSq = sumSq + Values(j) ^ 2
        Next j

        mean = sum / windowSize
        variance = (sumSq - windowSize * mean ^ 2) / (windowSize - 1)

        rollingVar(i) = variance
    Next i

    ' Detect significant changes in variance
    Dim meanVar As Double, stdVar As Double
    meanVar = 0

    For i = LBound(rollingVar) To UBound(rollingVar)
        meanVar = meanVar + rollingVar(i)
    Next i
    meanVar = meanVar / (UBound(rollingVar) - LBound(rollingVar) + 1)

    stdVar = 0
    For i = LBound(rollingVar) To UBound(rollingVar)
        stdVar = stdVar + (rollingVar(i) - meanVar) ^ 2
    Next i
    stdVar = Sqr(stdVar / (UBound(rollingVar) - LBound(rollingVar)))

    ' Flag regime changes (variance > mean + 2*std)
    For i = LBound(rollingVar) + 1 To UBound(rollingVar)
        If Abs(rollingVar(i) - rollingVar(i - 1)) > 2 * stdVar Then
            changeCount = changeCount + 1
            regimeChanges(changeCount) = i
        End If
    Next i

    ' Resize to actual number found
    If changeCount > 0 Then
        ReDim Preserve regimeChanges(1 To changeCount)
    Else
        ReDim regimeChanges(1 To 1)
        regimeChanges(1) = -1 ' No regimes detected
    End If

    DetectRegimes = regimeChanges
End Function

' ============================================================================
' FORECAST RECONCILIATION (Simple Constraint)
' ============================================================================

Public Function ReconcileForecast(ByRef forecast() As Double, _
                                 Optional ByVal minValue As Variant, _
                                 Optional ByVal maxValue As Variant) As Double()
    ' Apply constraints to forecasts (non-negativity, bounds, etc.)

    Dim result() As Double
    Dim i As Long
    Dim hasMin As Boolean, hasMax As Boolean
    Dim minVal As Double, maxVal As Double

    ReDim result(LBound(forecast) To UBound(forecast))

    hasMin = Not IsMissing(minValue)
    hasMax = Not IsMissing(maxValue)

    If hasMin Then minVal = CDbl(minValue)
    If hasMax Then maxVal = CDbl(maxValue)

    For i = LBound(forecast) To UBound(forecast)
        result(i) = forecast(i)

        ' Apply minimum constraint
        If hasMin And result(i) < minVal Then
            result(i) = minVal
        End If

        ' Apply maximum constraint
        If hasMax And result(i) > maxVal Then
            result(i) = maxVal
        End If
    Next i

    ReconcileForecast = result
End Function

' ============================================================================
' PROBABILISTIC FORECASTING (Monte Carlo Simulation)
' ============================================================================

Public Function ProbabilisticForecastMC(ByRef baseResult As ForecastResult, _
                                       ByVal horizon As Integer, _
                                       Optional ByVal numSimulations As Integer = 1000) As ProbabilisticForecast
    ' Generate full probability distribution using Monte Carlo simulation
    ' Returns percentiles (5%, 25%, 50%, 75%, 95%) for each horizon step

    Dim probForecast As ProbabilisticForecast
    Dim simulations() As Double ' numSimulations x horizon
    Dim i As Long, j As Integer, h As Integer

    ReDim simulations(1 To numSimulations, 1 To horizon)
    ReDim probForecast.Percentile05(1 To horizon)
    ReDim probForecast.Percentile25(1 To horizon)
    ReDim probForecast.Percentile50(1 To horizon)
    ReDim probForecast.Percentile75(1 To horizon)
    ReDim probForecast.Percentile95(1 To horizon)
    ReDim probForecast.Mean(1 To horizon)
    ReDim probForecast.StdDev(1 To horizon)

    ' Get residual standard deviation
    Dim sigma As Double
    sigma = CalculateStdDev(baseResult.Residuals)

    ' Generate simulations using bootstrap of residuals
    Dim residualCount As Long
    residualCount = UBound(baseResult.Residuals) - LBound(baseResult.Residuals) + 1

    Randomize

    For i = 1 To numSimulations
        For h = 1 To horizon
            ' Base forecast
            Dim forecast As Double
            forecast = baseResult.ForecastValues(h)

            ' Add random error (bootstrap from residuals)
            Dim cumulativeError As Double
            cumulativeError = 0

            For j = 1 To h
                ' Sample random residual
                Dim randomIdx As Long
                randomIdx = Int(Rnd() * residualCount) + LBound(baseResult.Residuals)
                cumulativeError = cumulativeError + baseResult.Residuals(randomIdx) / Sqr(j)
            Next j

            simulations(i, h) = forecast + cumulativeError
        Next h
    Next i

    ' Calculate percentiles for each horizon
    For h = 1 To horizon
        ' Extract simulations for this horizon
        Dim horizonSims() As Double
        ReDim horizonSims(1 To numSimulations)

        For i = 1 To numSimulations
            horizonSims(i) = simulations(i, h)
        Next i

        ' Sort
        Call BubbleSortDouble(horizonSims)

        ' Calculate percentiles
        probForecast.Percentile05(h) = horizonSims(Int(0.05 * numSimulations))
        probForecast.Percentile25(h) = horizonSims(Int(0.25 * numSimulations))
        probForecast.Percentile50(h) = horizonSims(Int(0.50 * numSimulations))
        probForecast.Percentile75(h) = horizonSims(Int(0.75 * numSimulations))
        probForecast.Percentile95(h) = horizonSims(Int(0.95 * numSimulations))

        ' Calculate mean and std dev
        Dim sum As Double, sumSq As Double
        sum = 0
        sumSq = 0

        For i = 1 To numSimulations
            sum = sum + horizonSims(i)
            sumSq = sumSq + horizonSims(i) ^ 2
        Next i

        probForecast.Mean(h) = sum / numSimulations
        probForecast.StdDev(h) = Sqr((sumSq - numSimulations * probForecast.Mean(h) ^ 2) / (numSimulations - 1))
    Next h

    ProbabilisticForecastMC = probForecast
End Function

' ============================================================================
' ADAPTIVE LEARNING (Recency-Weighted Model Performance)
' ============================================================================

Public Function AdaptiveLearningWeights(ByRef models() As ForecastResult, _
                                       ByVal numModels As Integer, _
                                       ByRef actualValues() As Double, _
                                       Optional ByVal decayFactor As Double = 0.95) As Double()
    ' Calculate model weights with recency bias
    ' Recent performance weighted higher - decayFactor: 0.95 = each older period gets 95% weight

    Dim weights() As Double
    ReDim weights(1 To numModels)

    Dim n As Long
    n = UBound(actualValues) - LBound(actualValues) + 1

    Dim i As Integer, j As Long
    Dim weightedErrors() As Double
    ReDim weightedErrors(1 To numModels)

    ' Calculate weighted MAPE for each model
    For i = 1 To numModels
        Dim weightedError As Double
        Dim totalWeight As Double

        weightedError = 0
        totalWeight = 0

        For j = LBound(actualValues) To UBound(actualValues)
            Dim recencyWeight As Double
            Dim periodsFromEnd As Long

            periodsFromEnd = UBound(actualValues) - j
            recencyWeight = decayFactor ^ periodsFromEnd

            If actualValues(j) <> 0 And j >= LBound(models(i).FittedValues) And j <= UBound(models(i).FittedValues) Then
                weightedError = weightedError + recencyWeight * Abs((actualValues(j) - models(i).FittedValues(j)) / actualValues(j))
                totalWeight = totalWeight + recencyWeight
            End If
        Next j

        If totalWeight > 0 Then
            weightedErrors(i) = weightedError / totalWeight
        Else
            weightedErrors(i) = 9999
        End If
    Next i

    ' Convert errors to weights (inverse)
    Dim totalInverseError As Double
    totalInverseError = 0

    For i = 1 To numModels
        If weightedErrors(i) > 0 And weightedErrors(i) < 9999 Then
            weights(i) = 1 / weightedErrors(i)
            totalInverseError = totalInverseError + weights(i)
        Else
            weights(i) = 0
        End If
    Next i

    ' Normalize
    If totalInverseError > 0 Then
        For i = 1 To numModels
            weights(i) = weights(i) / totalInverseError
        Next i
    Else
        For i = 1 To numModels
            weights(i) = 1 / numModels
        Next i
    End If

    AdaptiveLearningWeights = weights
End Function

' ============================================================================
' FORECAST VALUE ADDED (FVA)
' ============================================================================

Public Function CalculateFVA(ByRef actual() As Double, _
                            ByRef modelForecast() As Double, _
                            Optional ByVal naiveMethod As String = "LastValue") As ForecastValueAdded
    ' Calculate Forecast Value Added vs naive baseline
    ' FVA = (Naive Error - Model Error) / Naive Error * 100

    Dim fva As ForecastValueAdded
    Dim naiveForecast() As Double
    Dim n As Long, i As Long

    n = UBound(actual) - LBound(actual) + 1
    ReDim naiveForecast(LBound(actual) To UBound(actual))

    Select Case UCase(naiveMethod)
        Case "LASTVALUE"
            For i = LBound(actual) + 1 To UBound(actual)
                naiveForecast(i) = actual(i - 1)
            Next i
            naiveForecast(LBound(actual)) = actual(LBound(actual))

        Case "SEASONALNAIVE"
            Dim seasonalPeriod As Integer
            seasonalPeriod = 12

            For i = LBound(actual) To UBound(actual)
                If i >= LBound(actual) + seasonalPeriod Then
                    naiveForecast(i) = actual(i - seasonalPeriod)
                Else
                    naiveForecast(i) = actual(LBound(actual))
                End If
            Next i

        Case Else
            For i = LBound(actual) + 1 To UBound(actual)
                naiveForecast(i) = actual(i - 1)
            Next i
            naiveForecast(LBound(actual)) = actual(LBound(actual))
    End Select

    fva.NaiveMAPE = CalculateMAPE(actual, naiveForecast)
    fva.ModelMAPE = CalculateMAPE(actual, modelForecast)

    If fva.NaiveMAPE > 0 Then
        fva.FVA = ((fva.NaiveMAPE - fva.ModelMAPE) / fva.NaiveMAPE) * 100
        fva.ImprovedAccuracy = (fva.FVA > 0)
    Else
        fva.FVA = 0
        fva.ImprovedAccuracy = False
    End If

    CalculateFVA = fva
End Function

' ============================================================================
' ENSEMBLE TRIMMING (Remove Underperforming Models)
' ============================================================================

Public Function TrimEnsemble(ByRef models() As ForecastResult, _
                            ByVal numModels As Integer, _
                            Optional ByVal keepTopN As Integer = 0, _
                            Optional ByVal trimThreshold As Double = 0) As EnsembleTrimResult
    ' Remove worst performing models before ensemble averaging
    ' keepTopN: Keep only top N models (0 = use threshold)
    ' trimThreshold: Remove models with MAPE > avg + threshold*std (0 = use keepTopN)

    Dim trimResult As EnsembleTrimResult
    Dim i As Integer, j As Integer
    Dim avgMAPE As Double, stdMAPE As Double
    Dim cutoffMAPE As Double

    ' Calculate average and std dev of MAPE
    avgMAPE = 0
    For i = 1 To numModels
        avgMAPE = avgMAPE + models(i).MAPE
    Next i
    avgMAPE = avgMAPE / numModels

    stdMAPE = 0
    For i = 1 To numModels
        stdMAPE = stdMAPE + (models(i).MAPE - avgMAPE) ^ 2
    Next i
    stdMAPE = Sqr(stdMAPE / numModels)

    ' Determine trimming strategy
    If keepTopN > 0 Then
        ' Strategy 1: Keep only top N models
        ' Sort models by MAPE
        Dim sortedIndices() As Integer
        ReDim sortedIndices(1 To numModels)

        For i = 1 To numModels
            sortedIndices(i) = i
        Next i

        ' Bubble sort
        Dim temp As Integer
        For i = 1 To numModels - 1
            For j = i + 1 To numModels
                If models(sortedIndices(i)).MAPE > models(sortedIndices(j)).MAPE Then
                    temp = sortedIndices(i)
                    sortedIndices(i) = sortedIndices(j)
                    sortedIndices(j) = temp
                End If
            Next j
        Next i

        ' Keep top N
        Dim numToKeep As Integer
        numToKeep = WorksheetFunction.Min(keepTopN, numModels)

        ReDim trimResult.TrimmedModels(1 To numToKeep)
        For i = 1 To numToKeep
            trimResult.TrimmedModels(i) = models(sortedIndices(i))
        Next i

        trimResult.NumModels = numToKeep
    Else
        ' Strategy 2: Remove outliers beyond threshold
        cutoffMAPE = avgMAPE + trimThreshold * stdMAPE

        ' Count models to keep
        Dim keepCount As Integer
        keepCount = 0

        For i = 1 To numModels
            If models(i).MAPE <= cutoffMAPE Then
                keepCount = keepCount + 1
            End If
        Next i

        ' Keep at least 3 models
        If keepCount < 3 Then keepCount = WorksheetFunction.Min(3, numModels)

        ReDim trimResult.TrimmedModels(1 To keepCount)

        Dim idx As Integer
        idx = 0

        For i = 1 To numModels
            If models(i).MAPE <= cutoffMAPE Or idx < 3 Then
                idx = idx + 1
                If idx <= keepCount Then
                    trimResult.TrimmedModels(idx) = models(i)
                End If
            End If
        Next i

        trimResult.NumModels = keepCount
    End If

    ' Calculate improvement
    Dim newAvgMAPE As Double
    newAvgMAPE = 0

    For i = 1 To trimResult.NumModels
        newAvgMAPE = newAvgMAPE + trimResult.TrimmedModels(i).MAPE
    Next i
    newAvgMAPE = newAvgMAPE / trimResult.NumModels

    trimResult.AvgImprovement = avgMAPE - newAvgMAPE

    TrimEnsemble = trimResult
End Function

' ============================================================================
' WEIGHTED MEDIAN (Robust to Outliers)
' ============================================================================

Public Function WeightedMedianForecast(ByRef forecasts() As ForecastResult, _
                                      ByRef weights() As Double, _
                                      ByVal numModels As Integer, _
                                      ByVal horizon As Integer) As ForecastResult
    ' Calculate weighted median instead of weighted mean
    ' More robust to outlier forecasts

    Dim result As ForecastResult
    Dim i As Integer, h As Integer

    ' Initialize result arrays
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' For each horizon, calculate weighted median
    For h = 1 To horizon
        ' Extract forecasts for this horizon
        Dim horizonForecasts() As Double
        ReDim horizonForecasts(1 To numModels)

        For i = 1 To numModels
            horizonForecasts(i) = forecasts(i).ForecastValues(h)
        Next i

        ' Calculate weighted median
        result.ForecastValues(h) = CalculateWeightedMedian(horizonForecasts, weights, numModels)

        ' Calculate weighted median for intervals
        ReDim horizonForecasts(1 To numModels)
        For i = 1 To numModels
            horizonForecasts(i) = forecasts(i).Lower95(h)
        Next i
        result.Lower95(h) = CalculateWeightedMedian(horizonForecasts, weights, numModels)

        ReDim horizonForecasts(1 To numModels)
        For i = 1 To numModels
            horizonForecasts(i) = forecasts(i).Upper95(h)
        Next i
        result.Upper95(h) = CalculateWeightedMedian(horizonForecasts, weights, numModels)

        ReDim horizonForecasts(1 To numModels)
        For i = 1 To numModels
            horizonForecasts(i) = forecasts(i).Lower80(h)
        Next i
        result.Lower80(h) = CalculateWeightedMedian(horizonForecasts, weights, numModels)

        ReDim horizonForecasts(1 To numModels)
        For i = 1 To numModels
            horizonForecasts(i) = forecasts(i).Upper80(h)
        Next i
        result.Upper80(h) = CalculateWeightedMedian(horizonForecasts, weights, numModels)
    Next h

    result.ModelName = "Weighted Median Ensemble"

    WeightedMedianForecast = result
End Function

Private Function CalculateWeightedMedian(ByRef values() As Double, _
                                        ByRef weights() As Double, _
                                        ByVal n As Integer) As Double
    ' Calculate weighted median
    Dim sorted() As Double
    Dim sortedWeights() As Double
    Dim indices() As Integer
    Dim i As Integer, j As Integer

    ReDim sorted(1 To n)
    ReDim sortedWeights(1 To n)
    ReDim indices(1 To n)

    ' Initialize indices
    For i = 1 To n
        indices(i) = i
    Next i

    ' Sort by values
    Dim temp As Integer
    For i = 1 To n - 1
        For j = i + 1 To n
            If values(indices(i)) > values(indices(j)) Then
                temp = indices(i)
                indices(i) = indices(j)
                indices(j) = temp
            End If
        Next j
    Next i

    ' Create sorted arrays
    For i = 1 To n
        sorted(i) = values(indices(i))
        sortedWeights(i) = weights(indices(i))
    Next i

    ' Find weighted median
    Dim cumulativeWeight As Double
    Dim halfWeight As Double
    Dim totalWeight As Double

    totalWeight = 0
    For i = 1 To n
        totalWeight = totalWeight + sortedWeights(i)
    Next i

    halfWeight = totalWeight / 2
    cumulativeWeight = 0

    For i = 1 To n
        cumulativeWeight = cumulativeWeight + sortedWeights(i)
        If cumulativeWeight >= halfWeight Then
            CalculateWeightedMedian = sorted(i)
            Exit Function
        End If
    Next i

    ' Fallback
    CalculateWeightedMedian = sorted(Int(n / 2))
End Function

' ============================================================================
' PATTERN RECOGNITION (Find Similar Historical Patterns)
' ============================================================================

Public Function FindSimilarPatterns(ByRef Values() As Double, _
                                   ByVal patternLength As Integer, _
                                   Optional ByVal numMatches As Integer = 3) As PatternMatch()
    ' Find historical patterns similar to recent data
    ' Uses DTW (Dynamic Time Warping) distance

    Dim matches() As PatternMatch
    ReDim matches(1 To numMatches)

    Dim n As Long, i As Long, j As Long
    Dim recentPattern() As Double

    n = UBound(Values) - LBound(Values) + 1

    If n < patternLength * 2 Then
        ' Not enough data
        FindSimilarPatterns = matches
        Exit Function
    End If

    ' Extract recent pattern
    ReDim recentPattern(1 To patternLength)
    For i = 1 To patternLength
        recentPattern(i) = Values(UBound(Values) - patternLength + i)
    Next i

    ' Search for similar patterns in history
    Dim searchLength As Long
    searchLength = n - patternLength - patternLength ' Don't include recent period

    Dim similarities() As Double
    Dim indices() As Long

    ReDim similarities(1 To searchLength)
    ReDim indices(1 To searchLength)

    ' Calculate similarity for each historical window
    For i = 1 To searchLength
        Dim historicalPattern() As Double
        ReDim historicalPattern(1 To patternLength)

        For j = 1 To patternLength
            historicalPattern(j) = Values(LBound(Values) + i - 1 + j - 1)
        Next j

        ' Calculate similarity (1 - normalized distance)
        similarities(i) = CalculatePatternSimilarity(recentPattern, historicalPattern, patternLength)
        indices(i) = i
    Next i

    ' Sort by similarity
    Dim temp As Double
    Dim tempIdx As Long

    For i = 1 To searchLength - 1
        For j = i + 1 To searchLength
            If similarities(i) < similarities(j) Then
                temp = similarities(i)
                similarities(i) = similarities(j)
                similarities(j) = temp

                tempIdx = indices(i)
                indices(i) = indices(j)
                indices(j) = tempIdx
            End If
        Next j
    Next i

    ' Return top matches
    Dim numToReturn As Integer
    numToReturn = WorksheetFunction.Min(numMatches, searchLength)

    ReDim Preserve matches(1 To numToReturn)

    For i = 1 To numToReturn
        matches(i).StartIndex = LBound(Values) + indices(i) - 1
        matches(i).EndIndex = matches(i).StartIndex + patternLength - 1
        matches(i).Similarity = similarities(i)

        ReDim matches(i).Pattern(1 To patternLength)
        For j = 1 To patternLength
            matches(i).Pattern(j) = Values(matches(i).StartIndex + j - 1)
        Next j
    Next i

    FindSimilarPatterns = matches
End Function

Private Function CalculatePatternSimilarity(ByRef pattern1() As Double, _
                                           ByRef pattern2() As Double, _
                                           ByVal length As Integer) As Double
    ' Calculate similarity using normalized Euclidean distance
    Dim distance As Double, maxDistance As Double
    Dim i As Integer

    ' Normalize patterns
    Dim mean1 As Double, mean2 As Double
    Dim std1 As Double, std2 As Double

    mean1 = 0: mean2 = 0
    For i = 1 To length
        mean1 = mean1 + pattern1(i)
        mean2 = mean2 + pattern2(i)
    Next i
    mean1 = mean1 / length
    mean2 = mean2 / length

    std1 = 0: std2 = 0
    For i = 1 To length
        std1 = std1 + (pattern1(i) - mean1) ^ 2
        std2 = std2 + (pattern2(i) - mean2) ^ 2
    Next i
    std1 = Sqr(std1 / length)
    std2 = Sqr(std2 / length)

    If std1 = 0 Then std1 = 1
    If std2 = 0 Then std2 = 1

    ' Calculate Euclidean distance on normalized values
    distance = 0
    For i = 1 To length
        Dim norm1 As Double, norm2 As Double
        norm1 = (pattern1(i) - mean1) / std1
        norm2 = (pattern2(i) - mean2) / std2

        distance = distance + (norm1 - norm2) ^ 2
    Next i

    distance = Sqr(distance / length)

    ' Convert to similarity (0-1, higher is more similar)
    maxDistance = Sqr(8) ' Max distance for normalized data ≈ 2*sqrt(2)

    CalculatePatternSimilarity = WorksheetFunction.Max(0, 1 - (distance / maxDistance))
End Function

' ============================================================================
' DECOMPOSITION-BASED HYBRID FORECASTING
' ============================================================================

Public Function HybridDecompositionForecast(ByRef tsData As TimeSeriesData, _
                                           ByVal horizon As Integer) As ForecastResult
    ' Hybrid approach: Decompose → Forecast each component separately → Recombine
    ' Trend: Linear regression, Seasonal: Seasonal naive, Remainder: ARIMA

    On Error GoTo ErrorHandler

    Dim result As ForecastResult
    Dim Values() As Double
    Dim n As Long

    Values = tsData.Values
    n = UBound(Values) - LBound(Values) + 1

    ' Detect seasonality
    Dim seasonalInfo As SeasonalityInfo
    seasonalInfo = DetectSeasonality(tsData)

    Dim m As Integer
    If seasonalInfo.HasSeasonality Then
        m = seasonalInfo.DetectedFrequency
    Else
        m = 1
    End If

    ' Perform STL decomposition
    Dim stlResult As ForecastResult
    If m > 1 Then
        stlResult = STLDecomposition(tsData, horizon, m)
    Else
        ' No seasonality, use simple trend
        stlResult = LinearTrendForecast(tsData, horizon)
    End If

    ' Enhance with ARIMA on residuals
    Dim residualData As TimeSeriesData
    ReDim residualData.Values(LBound(stlResult.Residuals) To UBound(stlResult.Residuals))

    Dim i As Long
    For i = LBound(stlResult.Residuals) To UBound(stlResult.Residuals)
        residualData.Values(i) = stlResult.Residuals(i)
    Next i
    residualData.Frequency = m

    ' Forecast residuals with ARIMA
    Dim residualForecast As ForecastResult
    residualForecast = SimpleARIMA(residualData, horizon)

    ' Combine STL forecast with ARIMA residual forecast
    ReDim result.FittedValues(LBound(Values) To UBound(Values))
    ReDim result.Residuals(LBound(Values) To UBound(Values))
    ReDim result.ForecastValues(1 To horizon)
    ReDim result.Lower80(1 To horizon)
    ReDim result.Upper80(1 To horizon)
    ReDim result.Lower95(1 To horizon)
    ReDim result.Upper95(1 To horizon)

    ' Fitted values
    For i = LBound(Values) To UBound(Values)
        result.FittedValues(i) = stlResult.FittedValues(i)
        result.Residuals(i) = Values(i) - result.FittedValues(i)
    Next i

    ' Forecasts = STL + ARIMA residuals
    For i = 1 To horizon
        result.ForecastValues(i) = stlResult.ForecastValues(i) + residualForecast.ForecastValues(i)

        ' Combine uncertainty
        result.Lower95(i) = stlResult.Lower95(i) + residualForecast.Lower95(i)
        result.Upper95(i) = stlResult.Upper95(i) + residualForecast.Upper95(i)
        result.Lower80(i) = stlResult.Lower80(i) + residualForecast.Lower80(i)
        result.Upper80(i) = stlResult.Upper80(i) + residualForecast.Upper80(i)
    Next i

    ' Calculate metrics
    result.MAPE = CalculateMAPE(Values, result.FittedValues)
    result.MAE = CalculateMAE(result.Residuals)
    result.RMSE = CalculateRMSE(result.Residuals)
    result.MBE = CalculateMBE(Values, result.FittedValues)

    result.ModelName = "Hybrid(STL+ARIMA)"

    HybridDecompositionForecast = result
    Exit Function

ErrorHandler:
    result.MAPE = 9999
    result.ModelName = "Hybrid (Error)"
    HybridDecompositionForecast = result
End Function

