Public Module MOD_DATE_TOOL

#Region "日付関連"

    Public Function FUNC_GET_YEAR_ADD( _
    ByVal datDATE_BASE As DateTime, ByVal intADD As Integer _
    ) As DateTime
        Dim intYEAR As Integer
        Dim intMONTH As Integer
        Dim intDAY As Integer
        Dim datRET As DateTime

        intYEAR = datDATE_BASE.Year
        intMONTH = datDATE_BASE.Month
        intDAY = datDATE_BASE.Day

        intYEAR += intADD

        Try
            datRET = New DateTime(intYEAR, intMONTH, intDAY)
        Catch ex As Exception
            Return datDATE_BASE
        End Try

        Return datRET
    End Function

    'クエリ用に日付型を最小日付→NULL、通常日付「'」付与日付のみ文字列
    Public Function FUNC_GET_SQL_DATE(ByVal datDATE_BASE As DateTime) As String
        Dim strRET As String

        If datDATE_BASE < cstVB_DATE_MIN Then
            datDATE_BASE = cstVB_DATE_MIN
        End If

        If datDATE_BASE = cstVB_DATE_MIN Then
            strRET = "NULL"
        Else
            strRET = FUNC_ADD_ENCLOSED_SCOT(datDATE_BASE.ToShortDateString)
        End If

        Return strRET
    End Function
#End Region

#Region "期間関連"
    '特定の日付間の日数を取得
    Public Function FUNC_GET_DAYS_FROM_DATE_INTERVAL( _
    ByVal datDATE_FROM As DateTime, ByVal datDATE_TO As DateTime _
    ) As Integer
        Dim tsSPAN As TimeSpan
        Dim intRET As Integer

        tsSPAN = (datDATE_TO - datDATE_FROM)
        intRET = tsSPAN.Days

        Return intRET
    End Function

    '特定の日付間の年数を取得
    Public Function FUNC_GET_YEARS_FROM_DATE_INTERVAL( _
    ByVal datDATE_FROM As DateTime, ByVal datDATE_TO As DateTime _
    ) As Integer
        Dim intRET As Integer
        Dim intYEARS As Integer
        Dim intYEAR_INTERVAL As Integer
        Dim intMONTH_INTERVAL As Integer
        Dim intDAY_INTERVAL As Integer

        intYEAR_INTERVAL = datDATE_TO.Year - datDATE_FROM.Year
        intMONTH_INTERVAL = datDATE_TO.Month - datDATE_FROM.Month
        intDAY_INTERVAL = datDATE_TO.Day - datDATE_FROM.Day

        intYEARS = intYEAR_INTERVAL '基準年数
        Select Case intMONTH_INTERVAL
            Case 0 '同月の場合は日付基準
                If intDAY_INTERVAL < 0 Then '日付が過ぎていない場合
                    intYEARS -= 1 'ひとつ少なくなる
                End If
            Case Is < 0 '月が過ぎていない場合
                intYEARS -= 1 'ひとつ少なくなる
            Case Else 'それ以外(月が過ぎている場合)
                'スルー(そのままの年数)
        End Select

        intRET = intYEARS
        Return intRET
    End Function
#End Region

#Region "年月(YYYYMM)関連"

    '年月チェック　
    '例：TRUE="2004/05","2004/05/","2004年05","2004年05月"
    Public Function FUNC_CHECK_YEAR_AND_MONTH_STR( _
    ByVal strYEAR_MONTH As String _
    ) As Boolean
        Dim intLoopIndex As Integer
        Dim strTemp As String
        Dim strChar As String

        strTemp = ""
        For intLoopIndex = 1 To Len(strYEAR_MONTH)
            strChar = Mid(strYEAR_MONTH, intLoopIndex, 1)
            If IsNumeric(strChar) Then
                strTemp = strTemp & strChar
            Else
                Select Case strChar
                    Case "/", "年", "月"
                        'スルー
                    Case Else
                        Return False
                End Select
            End If
        Next

        If Not IsNumeric(strTemp) Then
            Return False
        End If

        Return FUNC_CHECK_YEAR_AND_MONTH(CInt(strTemp))

    End Function

    '年月チェック
    Public Function FUNC_CHECK_YEAR_AND_MONTH( _
    ByVal intYearAndMonth As Integer _
    ) As Boolean
        Dim strTemp As String
        Dim intYear As Integer
        Dim intMonth As Integer

        strTemp = Format(intYearAndMonth, "000000")

        If Len(strTemp) <> 6 Then
            Return False
        End If

        intYear = Mid(strTemp, 1, 4)
        intMonth = Mid(strTemp, 5, 2)

        If Not FUNC_CHECK_MONTH(intMonth) Then
            Return False
        End If

        Return True
    End Function

    '日付型から年月を数値で抜出
    Public Function FUNC_GET_YYYYMM_FROM_DATE( _
    ByVal datDATE_BASE As DateTime _
    ) As Integer
        Dim intRET As Integer
        Dim intYEAR As Integer
        Dim intMONTH As Integer
        Dim strTEMP As String

        intYEAR = datDATE_BASE.Year
        intMONTH = datDATE_BASE.Month

        strTEMP = Format(intYEAR, "0000") & Format(intMONTH, "00")

        If Not IsNumeric(strTEMP) Then
            Return 0
        End If

        intRET = CInt(strTEMP)
        Return intRET
    End Function

    '年月の形の数値から年の情報を抜出す
    Public Function FUNC_GET_YYYY_FROM_YYYYMM( _
    ByVal intCODE_YYYYMM As Integer _
    ) As Integer
        Dim intRET As Integer
        Dim strTEMP As String

        strTEMP = Format(intCODE_YYYYMM, "000000")
        Try
            strTEMP = strTEMP.Substring(0, 4)
        Catch ex As Exception
            strTEMP = ""
        End Try

        If Not IsNumeric(strTEMP) Then
            Return 0
        End If

        intRET = CInt(strTEMP)
        Return intRET
    End Function

    '年月の形の数値から月の情報を抜出す
    Public Function FUNC_GET_MM_FROM_YYYYMM( _
    ByVal intCODE_YYYYMM As Integer _
    ) As Integer
        Dim intRET As Integer
        Dim strTEMP As String

        strTEMP = Format(intCODE_YYYYMM, "000000")
        Try
            strTEMP = strTEMP.Substring(4, 2)
        Catch ex As Exception
            strTEMP = ""
        End Try

        If Not IsNumeric(strTEMP) Then
            Return 0
        End If

        intRET = CInt(strTEMP)
        Return intRET
    End Function

    '数値型年月を+-する
    Public Function FUNC_ADD_MONTH_YYYYMM( _
    ByVal intCODE_YYYYMM As Integer, ByVal intADD_MONTH As Integer _
    ) As Integer
        Dim intRET As Integer
        Dim intCODE_YEAR As Integer
        Dim intCODE_MONTH As Integer
        Dim datBASE As DateTime

        intCODE_YEAR = FUNC_GET_YYYY_FROM_YYYYMM(intCODE_YYYYMM)
        intCODE_MONTH = FUNC_GET_MM_FROM_YYYYMM(intCODE_YYYYMM)

        Try
            datBASE = New DateTime(intCODE_YEAR, intCODE_MONTH, 1) '月初で生成
        Catch ex As Exception
            Return intCODE_YYYYMM
        End Try

        datBASE = datBASE.AddMonths(intADD_MONTH)

        intRET = FUNC_GET_YYYYMM_FROM_DATE(datBASE)

        Return intRET
    End Function

    '年月の形の数値を文字列型(/付)に変換
    Public Function FUNC_CONVERT_STR_FROM_YYYYMM( _
    ByVal intCODE_YYYYMM As Integer _
    ) As String
        Dim intYEAR As Integer
        Dim intMONTH As Integer
        Dim strRET As String

        intYEAR = FUNC_GET_YYYY_FROM_YYYYMM(intCODE_YYYYMM)
        intMONTH = FUNC_GET_MM_FROM_YYYYMM(intCODE_YYYYMM)

        strRET = Format(intYEAR, "0000") & "/" & Format(intMONTH, "00")

        Return strRET
    End Function

    '年月の形の文字列(/付)を数値に変換
    Public Function FUNC_CONVERT_YYYYMM_FROM_STR( _
    ByVal strCODE_YYYYMM As String _
    ) As Integer
        Dim intRET As Integer
        Dim strTEMP As String

        strTEMP = strCODE_YYYYMM.Replace("/", "")

        If IsNumeric(strTEMP) Then
            intRET = CInt(strTEMP)
        Else
            intRET = 0
        End If

        Return intRET
    End Function

    '数値年月のFromとToから月数の差分を取得する
    '例：200804～200903→11　200804～200804→0　200810～200903→5　200903～200804→-11
    Public Function FUNC_GET_MONTH_FROM_TO( _
    ByVal intCODE_YYYYMM_FROM As Integer, ByVal intCODE_YYYYMM_TO As Integer _
    ) As Integer
        Dim intYEAR_FROM As Integer
        Dim intYEAR_TO As Integer
        Dim intMONTH_FROM As Integer
        Dim intMONTH_TO As Integer
        Dim intYEAR_DIFF As Integer
        Dim intMONTH_DIFF As Integer
        Dim intRET As Integer

        intYEAR_FROM = FUNC_GET_YYYY_FROM_YYYYMM(intCODE_YYYYMM_FROM)
        intMONTH_FROM = FUNC_GET_MM_FROM_YYYYMM(intCODE_YYYYMM_FROM)
        intYEAR_TO = FUNC_GET_YYYY_FROM_YYYYMM(intCODE_YYYYMM_TO)
        intMONTH_TO = FUNC_GET_MM_FROM_YYYYMM(intCODE_YYYYMM_TO)

        intYEAR_DIFF = (intYEAR_TO - intYEAR_FROM)
        intMONTH_DIFF = (intMONTH_TO - intMONTH_FROM)

        intRET = (intYEAR_DIFF * 12) + intMONTH_DIFF

        Return intRET
    End Function

    '数値年月のFromとToから月数の期間を取得する
    '例：200804～200903→12　200804～200804→1　200810～200903→6　200804～200803→0　200903～200804→0
    Public Function FUNC_GET_MONTH_KIKAN_FROM_TO( _
    ByVal intCODE_YYYYMM_FROM As Integer, ByVal intCODE_YYYYMM_TO As Integer _
    ) As Integer
        Dim intMONTH_SABUN As Integer
        Dim intRET As Integer

        intMONTH_SABUN = FUNC_GET_MONTH_FROM_TO(intCODE_YYYYMM_FROM, intCODE_YYYYMM_TO)

        If intMONTH_SABUN < 0 Then
            intRET = 0
        Else
            intRET = intMONTH_SABUN + 1
        End If

        Return intRET
    End Function

    '数値年月から日付型(月初)の日付型の変数を返す
    Public Function FUNC_GET_DATE_FROM_YEARMONTH( _
    ByVal intCODE_YYYYMM As Integer _
    ) As DateTime
        Dim datRET As DateTime
        Dim intYEAR As Integer
        Dim intMONTH As Integer
        Dim intDAY As Integer

        intYEAR = FUNC_GET_YYYY_FROM_YYYYMM(intCODE_YYYYMM)
        intMONTH = FUNC_GET_MM_FROM_YYYYMM(intCODE_YYYYMM)
        intDAY = 1

        Try
            datRET = New DateTime(intYEAR, intMONTH, intDAY)
        Catch ex As Exception
            Return cstVB_DATE_MIN
        End Try

        Return datRET
    End Function

    '数値年月から日付型(月末)の日付型の変数を返す
    Public Function FUNC_GET_DATE_LAST_FROM_YEARMONTH( _
    ByVal intCODE_YYYYMM As Integer _
    ) As DateTime
        Dim datFIRST As DateTime
        Dim datRET As DateTime


        datFIRST = FUNC_GET_DATE_FROM_YEARMONTH(intCODE_YYYYMM)

        datRET = FUNC_GET_DATE_LASTMONTH(datFIRST)

        Return datRET
    End Function

    '年度と月を年月に変換
    '例：2012,4→201204  2012,3→201303
    Public Function FUNC_GET_YYYYMM_FROM_YYYY_MM( _
    ByVal intCODE_YYYY As Integer, _
    ByVal intCODE_MM As Integer _
    ) As Integer
        Dim intYearMonth As Integer
        Dim intYear As Integer

        If intCODE_MM = 1 Or intCODE_MM = 2 Or intCODE_MM = 3 Then
            intYear = (intCODE_YYYY + 1) * 100
        Else
            intYear = intCODE_YYYY * 100
        End If

        intYear = Format(intYear, "000000")
        intYearMonth = intYear + intCODE_MM

        If Not FUNC_CHECK_YEAR_AND_MONTH(intYearMonth) Then
            Return 0
        End If

        Return intYearMonth
    End Function

    '文字列年月を足し算する
    Public Function FUNC_GET_PLUS_YEAR_AND_MONTH(ByVal strYEAR_MONTH As String, ByVal datDATE_MIN As DateTime, ByVal datDATE_MAX As DateTime) As String
        Dim intYEAR_MONTH As Integer
        Dim intYEAR_MONTH_MIN As Integer
        Dim intYEAR_MONTH_MAX As Integer

        If Not FUNC_CHECK_YEAR_AND_MONTH_STR(strYEAR_MONTH) Then
            Return FUNC_CONVERT_STR_FROM_YYYYMM(FUNC_GET_YYYYMM_FROM_DATE(datDATE_MIN))
        End If

        intYEAR_MONTH = FUNC_CONVERT_YYYYMM_FROM_STR(strYEAR_MONTH)
        intYEAR_MONTH_MIN = FUNC_GET_YYYYMM_FROM_DATE(datDATE_MIN)
        intYEAR_MONTH_MAX = FUNC_GET_YYYYMM_FROM_DATE(datDATE_MAX)

        If intYEAR_MONTH >= intYEAR_MONTH_MAX Then
            Return strYEAR_MONTH
        End If

        Return FUNC_CONVERT_STR_FROM_YYYYMM(FUNC_ADD_MONTH_YYYYMM(intYEAR_MONTH, 1))
    End Function

    '文字列年月を引き算する
    Public Function FUNC_GET_MINUS_YEAR_AND_MONTH(ByVal strYEAR_MONTH As String, ByVal datDATE_MIN As DateTime, ByVal datDATE_MAX As DateTime) As String
        Dim intYEAR_MONTH As Integer
        Dim intYEAR_MONTH_MIN As Integer
        Dim intYEAR_MONTH_MAX As Integer

        If Not FUNC_CHECK_YEAR_AND_MONTH_STR(strYEAR_MONTH) Then
            Return FUNC_CONVERT_STR_FROM_YYYYMM(FUNC_GET_YYYYMM_FROM_DATE(datDATE_MAX))
        End If

        intYEAR_MONTH = FUNC_CONVERT_YYYYMM_FROM_STR(strYEAR_MONTH)
        intYEAR_MONTH_MIN = FUNC_GET_YYYYMM_FROM_DATE(datDATE_MIN)
        intYEAR_MONTH_MAX = FUNC_GET_YYYYMM_FROM_DATE(datDATE_MAX)

        If intYEAR_MONTH <= intYEAR_MONTH_MIN Then
            Return strYEAR_MONTH
        End If

        Return FUNC_CONVERT_STR_FROM_YYYYMM(FUNC_ADD_MONTH_YYYYMM(intYEAR_MONTH, -1))
    End Function

#End Region

#Region "数値型和暦日付(gyyMMdd)関連"

    '数値型和暦日付(201225等)から年を抜出
    Public Function FUNC_GET_YEAR_FROM_WAREKI( _
    ByVal intDATE_WAREKI As Integer _
    ) As Integer
        Dim strTEMP As String
        Dim intRET As Integer

        strTEMP = Format(intDATE_WAREKI, "00000000") '一旦8桁で書式整形
        strTEMP = strTEMP.Substring(strTEMP.Length - 8, 8) '8桁になるように右端から間引き
        strTEMP = strTEMP.Substring(0, 4) '1～4文字が年

        If IsNumeric(strTEMP) Then
            intRET = CInt(strTEMP)
        Else
            intRET = 0
        End If

        Return intRET
    End Function

    '数値型和暦日付(201225等)から月を抜出
    Public Function FUNC_GET_MONTH_FROM_WAREKI( _
    ByVal intDATE_WAREKI As Integer _
    ) As Integer
        Dim strTEMP As String
        Dim intRET As Integer

        strTEMP = Format(intDATE_WAREKI, "00000000") '一旦8桁で書式整形
        strTEMP = strTEMP.Substring(strTEMP.Length - 8, 8) '8桁になるように右端から間引き
        strTEMP = strTEMP.Substring(4, 2) '5～6文字が年

        If IsNumeric(strTEMP) Then
            intRET = CInt(strTEMP)
        Else
            intRET = 0
        End If

        Return intRET
    End Function

    '数値型和暦日付(201225等)から日を抜出
    Public Function FUNC_GET_DAY_FROM_WAREKI( _
    ByVal intDATE_WAREKI As Integer _
    ) As Integer
        Dim strTEMP As String
        Dim intRET As Integer

        strTEMP = Format(intDATE_WAREKI, "00000000") '一旦8桁で書式整形
        strTEMP = strTEMP.Substring(strTEMP.Length - 8, 8) '8桁になるように右端から間引き
        strTEMP = strTEMP.Substring(6, 2) '7～8文字が年

        If IsNumeric(strTEMP) Then
            intRET = CInt(strTEMP)
        Else
            intRET = 0
        End If

        Return intRET
    End Function

#End Region

#Region "数値型日付関連"

    '数値型日付(yyyyMMdd)→日付型への変換
    Public Function FUNC_CONVERT_NUMERIC_DATE_TO_DATETIME( _
    ByVal intDATE_NUM As Integer _
    ) As DateTime
        Dim strTEMP As String
        Dim intYEAR As Integer
        Dim intMONTH As Integer
        Dim intDAY As Integer
        Dim datRET As DateTime

        strTEMP = Format(intDATE_NUM, "00000000")
        strTEMP = strTEMP.Substring(strTEMP.Length - 8, 8) '8桁になるように右端から間引き
        intYEAR = CInt(strTEMP.Substring(0, 4))
        intMONTH = CInt(strTEMP.Substring(4, 2))
        intDAY = CInt(strTEMP.Substring(6, 2))

        If Not (1 <= intMONTH And intMONTH <= 12) Then
            Return cstVB_DATE_MIN
        End If

        If Not (1 <= intDAY And intDAY <= 31) Then
            Return cstVB_DATE_MIN
        End If

        Try
            datRET = New DateTime(intYEAR, intMONTH, intDAY)
        Catch ex As Exception
            Return cstVB_DATE_MIN
        End Try

        Return datRET
    End Function

    Public Function FUNC_CONVERT_DATETIME_TO_NUMERIC_DATE( _
    ByVal datDATE As DateTime _
    ) As Integer
        Dim intRET As Integer
        Dim intYEAR As Integer
        Dim intMONTH As Integer
        Dim intDAY As Integer

        intYEAR = datDATE.Year
        intMONTH = datDATE.Month
        intDAY = datDATE.Day

        intRET = (intYEAR * 10000) + (intMONTH * 100) + (intDAY)

        Return intRET
    End Function
#End Region

#Region "月関連"
    '月チェック(1～12ならOK)
    Public Function FUNC_CHECK_MONTH(ByVal intMonth As Integer) As Boolean
        Select Case intMonth
            Case 1 To 12
                Return True
            Case Else
                Return False
        End Select
    End Function

    '月初日付取得
    Public Function FUNC_GET_DATE_FIRSMONTH( _
    ByVal datDATE_BASE As DateTime _
    ) As DateTime
        Dim datTEMP As DateTime
        Dim datRET As DateTime

        Try
            datTEMP = New DateTime(datDATE_BASE.Year, datDATE_BASE.Month, 1) '1日に固定する
            datRET = datTEMP
        Catch ex As Exception
            Return datDATE_BASE
        End Try

        Return datRET

    End Function

    '月末日付取得
    Public Function FUNC_GET_DATE_LASTMONTH( _
    ByVal datDATE_BASE As DateTime _
    ) As DateTime
        Dim datNEXT_MONTH As DateTime
        Dim datTEMP As DateTime
        Dim datRET As DateTime

        Try
            datTEMP = datDATE_BASE.AddMonths(1) '翌月を求める
            datNEXT_MONTH = New DateTime(datTEMP.Year, datTEMP.Month, 1) '1日に固定する
            datRET = datNEXT_MONTH.AddDays(-1) '1日戻す(月末になる)
        Catch ex As Exception
            Return datDATE_BASE
        End Try

        Return datRET

    End Function

    '月末日チェック
    Public Function FUNC_CHECK_DATE_LASTMONTH( _
    ByVal datDATE_BASE As DateTime _
    ) As Boolean
        Dim datTEMP As DateTime
        Dim intDAY As Integer

        Try
            datTEMP = datDATE_BASE.AddDays(1) '翌日を求める
        Catch ex As Exception
            Return False
        End Try

        intDAY = datTEMP.Day

        Return (intDAY = 1)
    End Function
#End Region

#Region "和暦関連"

    Public Function FUNC_GET_ERA_WAREKI_FROM_DATE(ByVal datBASE As DateTime) As Integer
        Dim calJAP As System.Globalization.JapaneseCalendar
        Dim intRET As Integer

        calJAP = New System.Globalization.JapaneseCalendar
        intRET = calJAP.GetEra(datBASE)

        Return intRET
    End Function

    Public Function FUNC_GET_YEAR_WAREKI_FROM_DATE(ByVal datBASE As DateTime) As Integer
        Dim calJAP As System.Globalization.JapaneseCalendar
        Dim intRET As Integer

        calJAP = New System.Globalization.JapaneseCalendar
        intRET = calJAP.GetYear(datBASE)

        Return intRET
    End Function

    '和暦でのフォーマット(yymmddの6桁)
    Public Function FUNC_GET_WAREKI_NUM_YYMMDD(ByVal intYEAR As Integer, ByVal intMONTH As Integer, ByVal intDAY As Integer) As Integer
        Dim intRET As Integer

        intRET = (intYEAR * 10000) + (intMONTH * 100) + intDAY

        Return intRET
    End Function
#End Region

End Module
