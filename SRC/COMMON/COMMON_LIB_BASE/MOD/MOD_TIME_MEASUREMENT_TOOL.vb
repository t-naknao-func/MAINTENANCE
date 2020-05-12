'時間計測用モジュール
Public Module MOD_TIME_MEASUREMENT_TOOL
    Private stwTIME_MEASUREMENT As System.Diagnostics.Stopwatch
    Private dblTIME_MEASUREMENT_TOTAL_SECONDS As Double
    Private Const CST_DEFAULT_METHOD_NAME As String = "不明な処理"

    '計測の開始
    Public Sub SUB_TIME_MEASUREMEN_START()
        dblTIME_MEASUREMENT_TOTAL_SECONDS = 0
        stwTIME_MEASUREMENT = Nothing
        stwTIME_MEASUREMENT = New System.Diagnostics.Stopwatch
        Call stwTIME_MEASUREMENT.Start()
    End Sub

    '計測の開始(再)
    Public Sub SUB_TIME_MEASUREMEN_RESTART()
        dblTIME_MEASUREMENT_TOTAL_SECONDS = 0
        If stwTIME_MEASUREMENT Is Nothing Then
            stwTIME_MEASUREMENT = New System.Diagnostics.Stopwatch
        End If
        Call stwTIME_MEASUREMENT.Start()
    End Sub

    '計測の終了
    Public Sub SUB_TIME_MEASUREMEN_STOP()
        Dim tpnTIME_SPAN As TimeSpan

        Call stwTIME_MEASUREMENT.Stop()
        tpnTIME_SPAN = stwTIME_MEASUREMENT.Elapsed
        dblTIME_MEASUREMENT_TOTAL_SECONDS = tpnTIME_SPAN.TotalSeconds

        'stwTIME_MEASUREMENT = Nothing
    End Sub

    'コンソールへ表示
    Public Sub SUB_TIME_MEASUREMEN_PUT_LOG(Optional ByVal strMETHOD_NAME As String = CST_DEFAULT_METHOD_NAME)
        Dim strLOG As String

        strLOG = ""
        strLOG &= strMETHOD_NAME & ":"
        strLOG &= String.Format("{0:0.00}", dblTIME_MEASUREMENT_TOTAL_SECONDS) & "秒"

        Call Console.WriteLine(strLOG)
    End Sub

    '計測終了&コンソール表示
    Public Sub SUB_TIME_MEASUREMENT_STOP_AND_PUT_LOG(Optional ByVal strMETHOD_NAME As String = CST_DEFAULT_METHOD_NAME)
        Call SUB_TIME_MEASUREMEN_STOP()
        Call SUB_TIME_MEASUREMEN_PUT_LOG(strMETHOD_NAME)
    End Sub

End Module
