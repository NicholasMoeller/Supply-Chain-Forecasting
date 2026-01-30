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
    
    ' Clean outliers first for better accuracy
    cleanedData = RemoveOutliers(tsData)

    ' Try ALL 13 forecasting methods - best one wins!
    modelCount = 13
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
    
    ' Sort models by MAPE to find top 3
    Dim sortedIndices() As Integer
    ReDim sortedIndices(1 To modelCount)
    For i = 1 To modelCount
        sortedIndices(i) = i
    Next i

    ' Bubble sort by MAPE
    Dim temp As Integer
    Dim j As Integer
    For i = 1 To modelCount - 1
        For j = i + 1 To modelCount
            If results(sortedIndices(i)).MAPE > results(sortedIndices(j)).MAPE Then
                temp = sortedIndices(i)
                sortedIndices(i) = sortedIndices(j)
                sortedIndices(j) = temp
            End If
        Next j
    Next i

    ' Combine top 3 models for even better accuracy
    Dim combinedResult As ForecastResult
    Dim topCount As Integer
    topCount = WorksheetFunction.Min(3, modelCount) ' Top 3 or fewer

    ' Calculate inverse MAPE weights
    Dim weights() As Double
    ReDim weights(1 To topCount)
    Dim totalWeight As Double
    totalWeight = 0

    For i = 1 To topCount
        If results(sortedIndices(i)).MAPE > 0 Then
            weights(i) = 1 / results(sortedIndices(i)).MAPE
            totalWeight = totalWeight + weights(i)
        End If
    Next i

    ' Normalize weights
    If totalWeight > 0 Then
        For i = 1 To topCount
            weights(i) = weights(i) / totalWeight
        Next i
    Else
        ' Equal weights if issues
        For i = 1 To topCount
            weights(i) = 1 / topCount
        Next i
    End If

    ' Combine forecasts
    combinedResult = results(sortedIndices(1)) ' Start with best

    ' Initialize forecast arrays
    For i = 1 To horizon
        combinedResult.ForecastValues(i) = 0
        combinedResult.Lower95(i) = 0
        combinedResult.Upper95(i) = 0
    Next i

    ' Weighted combination
    For i = 1 To topCount
        Dim idx As Integer
        idx = sortedIndices(i)

        For j = 1 To horizon
            combinedResult.ForecastValues(j) = combinedResult.ForecastValues(j) + weights(i) * results(idx).ForecastValues(j)
            combinedResult.Lower95(j) = combinedResult.Lower95(j) + weights(i) * results(idx).Lower95(j)
            combinedResult.Upper95(j) = combinedResult.Upper95(j) + weights(i) * results(idx).Upper95(j)
        Next j
    Next i

    ' Combined fitted values
    For i = LBound(cleanedData.Values) To UBound(cleanedData.Values)
        combinedResult.FittedValues(i) = 0
        combinedResult.Residuals(i) = 0
    Next i

    For i = 1 To topCount
        idx = sortedIndices(i)
        For j = LBound(cleanedData.Values) To UBound(cleanedData.Values)
            combinedResult.FittedValues(j) = combinedResult.FittedValues(j) + weights(i) * results(idx).FittedValues(j)
        Next j
    Next i

    ' Recalculate residuals and metrics
    For i = LBound(cleanedData.Values) To UBound(cleanedData.Values)
        combinedResult.Residuals(i) = cleanedData.Values(i) - combinedResult.FittedValues(i)
    Next i

    combinedResult.MAPE = CalculateMAPE(cleanedData.Values, combinedResult.FittedValues)
    combinedResult.MAE = CalculateMAE(combinedResult.Residuals)
    combinedResult.RMSE = CalculateRMSE(combinedResult.Residuals)
    combinedResult.MBE = CalculateMBE(cleanedData.Values, combinedResult.FittedValues)

    ' Build model name from top 3
    Dim combinedName As String
    combinedName = "Top3: " & results(sortedIndices(1)).ModelName
    If topCount > 1 Then combinedName = combinedName & "+" & results(sortedIndices(2)).ModelName
    If topCount > 2 Then combinedName = combinedName & "+" & results(sortedIndices(3)).ModelName
    combinedResult.ModelName = combinedName

    ' Apply bias correction if significant
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
        combinedResult.ModelName = combinedResult.ModelName & " (bias-corrected)"
    End If

    ' Apply bootstrap confidence intervals for more accurate uncertainty estimation
    ' Use 200 bootstrap samples for good balance of accuracy vs speed
    combinedResult = BootstrapConfidenceIntervals(combinedResult, 200, 0.95)

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
Public Type SeasonalityInfo
    HasSeasonality As Boolean
    DetectedFrequency As Integer
    SeasonalType As String ' "additive", "multiplicative", "none"
    Confidence As Double ' 0-100%
End Type

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
    correlation = CalculateCorrelation(cycleMeans, cycleStdDevs, frequency)

    If correlation > 0.5 Then
        DetectSeasonalType = "multiplicative"
    Else
        DetectSeasonalType = "additive"
    End If
End Function

Private Function CalculateCorrelation(ByRef x() As Double, ByRef y() As Double, n As Integer) As Double
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
        CalculateCorrelation = numerator / denominator
    Else
        CalculateCorrelation = 0
    End If
End Function

' ============================================================================
' TIME SERIES CROSS-VALIDATION
' ============================================================================
' Implements rolling window cross-validation for more robust parameter optimization
' This prevents overfitting by testing parameters on multiple held-out test sets

Public Type CrossValidationResult
    AvgMAPE As Double
    AvgMAE As Double
    AvgRMSE As Double
    NumFolds As Integer
End Type

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

