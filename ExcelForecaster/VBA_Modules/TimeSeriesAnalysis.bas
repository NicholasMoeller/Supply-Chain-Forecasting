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
    Dim bestSSE As Double
    Dim currentAlpha As Double
    Dim currentSSE As Double

    bestAlpha = 0.1
    bestSSE = 1E+100

    For currentAlpha = 0.01 To 0.99 Step 0.01
        currentSSE = CalculateSSE_SES(Values, currentAlpha)
        If currentSSE < bestSSE Then
            bestSSE = currentSSE
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
    Dim bestSSE As Double
    Dim currentAlpha As Double, currentBeta As Double, currentGamma As Double
    Dim currentSSE As Double

    bestAlpha = 0.3
    bestBeta = 0.1
    bestGamma = 0.1
    bestSSE = 1E+100

    ' Coarse grid search
    For currentAlpha = 0.1 To 0.9 Step 0.2
        For currentBeta = 0.05 To 0.3 Step 0.1
            For currentGamma = 0.05 To 0.3 Step 0.1
                currentSSE = CalculateSSE_HW(Values, Frequency, currentAlpha, currentBeta, currentGamma, seasonalType)
                If currentSSE < bestSSE Then
                    bestSSE = currentSSE
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
    
    ' Try all forecasting methods
    modelCount = 6
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
    
    ' 6. ARIMA
    On Error Resume Next
    results(6) = SimpleARIMA(cleanedData, horizon)
    results(6).ModelName = "ARIMA"
    If Err.Number <> 0 Then results(6).MAPE = 9999
    On Error GoTo ErrorHandler
    
    ' Find best model (lowest MAPE)
    bestMAPE = results(1).MAPE
    bestIndex = 1
    For i = 2 To modelCount
        If results(i).MAPE < bestMAPE And results(i).MAPE > 0 Then
            bestMAPE = results(i).MAPE
            bestIndex = i
        End If
    Next i
    
    AutoForecast = results(bestIndex)
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
    result.MAE = CalculateMAE(tsData.Values, result.FittedValues)
    result.RMSE = CalculateRMSE(tsData.Values, result.FittedValues)
    result.MBE = CalculateMBE(result.Residuals)
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
    result.MAE = CalculateMAE(Values, result.FittedValues)
    result.RMSE = CalculateRMSE(Values, result.FittedValues)
    result.MBE = CalculateMBE(result.Residuals)
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
    result.MAE = CalculateMAE(Values, result.FittedValues)
    result.RMSE = CalculateRMSE(Values, result.FittedValues)
    result.MBE = CalculateMBE(result.Residuals)
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
    result.MAE = CalculateMAE(Values, result.FittedValues)
    result.RMSE = CalculateRMSE(Values, result.FittedValues)
    result.MBE = CalculateMBE(result.Residuals)
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
