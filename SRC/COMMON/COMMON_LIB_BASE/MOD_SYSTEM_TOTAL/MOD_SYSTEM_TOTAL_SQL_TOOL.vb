Public Module MOD_SYSTEM_TOTAL_SQL_TOOL

#Region "CONNECTION関連"

    Public Function FUNC_SYSTEM_SET_CONNECTION_TIMEOUT( _
    Optional ByVal intTIMEOUT As Integer = 15 _
    ) As Boolean

        'sql_SYSTEM_PUBLIC_CONNECTION.ConnectionTimeout = intTIMEOUT

        Return True
    End Function
#End Region

#Region "SELECT関連"
    '結果セットをSqlDataReader型で取得する
    Public Function FUNC_SYSTEM_GET_SQL_DATA_READER( _
    ByVal strSql As String, _
    ByRef sdrDATA_READER_RET As SqlClient.SqlDataReader, _
    Optional ByVal hiReadOpt As CommandBehavior = CommandBehavior.Default, _
    Optional ByVal intTIME_OUT As Integer = -1 _
    ) As Boolean
        Dim blnRET As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        Call SUB_START_TICK()
        blnRET = FUNC_GET_SQL_DATA_READER(strSql, sdrDATA_READER_RET, hiReadOpt, intTIME_OUT, sql_SYSTEM_PUBLIC_TRANSACTION)
        Call SUB_PUT_LOG_TICK(strSql)

        Call SUB_REMOVE_SQL_CONNECTION()

        If Not blnRET Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING & Environment.NewLine & strSql.ToString) 'デバッグログを追記
        End If

        Return blnRET
    End Function

    '任意の1列1行の値を取得する
    Public Function FUNC_SYSTEM_GET_SQL_SINGLE_VALUE(ByVal strSql As String, _
    Optional ByVal intTIME_OUT As Integer = -1 _
    ) As Object
        Dim objRET As Object
        Dim blnSQL_DONE As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        Call SUB_START_TICK()
        objRET = FUNC_GET_SQL_SINGLE_VALUE(strSql, intTIME_OUT, sql_SYSTEM_PUBLIC_TRANSACTION, blnSQL_DONE)
        Call SUB_PUT_LOG_TICK(strSql)

        Call SUB_REMOVE_SQL_CONNECTION()

        If Not blnSQL_DONE Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING & Environment.NewLine & strSql.ToString) 'デバッグログを追記
        End If

        Return objRET
    End Function

    '任意の1列1行の値を取得する(整数用)
    Public Function FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC( _
    ByVal strSQL As String, Optional ByVal lngMISS_VALUE As Long = -1, _
    Optional ByVal intTIME_OUT As Integer = -1 _
    ) As Long
        Dim lngRET As Long
        Dim blnSQL_DONE As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        Call SUB_START_TICK()
        lngRET = FUNC_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, lngMISS_VALUE, intTIME_OUT, sql_SYSTEM_PUBLIC_TRANSACTION, blnSQL_DONE)
        Call SUB_PUT_LOG_TICK(strSQL)

        Call SUB_REMOVE_SQL_CONNECTION()

        If Not blnSQL_DONE Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING & Environment.NewLine & strSQL.ToString) 'デバッグログを追記
        End If

        Return lngRET
    End Function

    '任意の1列1行の値を取得する(金額、量値、少数用)
    Public Function FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_DECIMAL( _
    ByVal strSQL As String, Optional ByVal decMISS_VALUE As Decimal = 0, _
    Optional ByVal intTIME_OUT As Integer = -1 _
    ) As Double
        Dim dblRET As Double
        Dim blnSQL_DONE As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        Call SUB_START_TICK()
        dblRET = FUNC_GET_SQL_SINGLE_VALUE_DECIMAL(strSQL, decMISS_VALUE, intTIME_OUT, sql_SYSTEM_PUBLIC_TRANSACTION, blnSQL_DONE)
        Call SUB_PUT_LOG_TICK(strSQL)

        Call SUB_REMOVE_SQL_CONNECTION()

        If Not blnSQL_DONE Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING & Environment.NewLine & strSQL.ToString) 'デバッグログを追記
        End If

        Return dblRET
    End Function

    '任意の1列1行の値を取得する(文字列用)
    Public Function FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING( _
    ByVal strSQL As String, Optional ByVal strMISS_VALUE As String = "", _
    Optional ByVal intTIME_OUT As Integer = -1 _
    ) As String
        Dim strRET As String
        Dim blnSQL_DONE As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        Call SUB_START_TICK()
        strRET = FUNC_GET_SQL_SINGLE_VALUE_STRING(strSQL, strMISS_VALUE, intTIME_OUT, sql_SYSTEM_PUBLIC_TRANSACTION, blnSQL_DONE)
        Call SUB_PUT_LOG_TICK(strSQL)

        Call SUB_REMOVE_SQL_CONNECTION()

        If Not blnSQL_DONE Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING & Environment.NewLine & strSQL.ToString) 'デバッグログを追記
        End If

        Return strRET
    End Function

    '任意の1列1行の値を取得する(日付用)
    Public Function FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_DATETIME( _
    ByVal strSQL As String, Optional ByVal datMISS_VALUE As DateTime = cstVB_DATE_MIN, _
    Optional ByVal intTIME_OUT As Integer = -1 _
    ) As DateTime
        Dim datRET As DateTime
        Dim blnSQL_DONE As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        Call SUB_START_TICK()
        datRET = FUNC_GET_SQL_SINGLE_VALUE_DATETIME(strSQL, datMISS_VALUE, intTIME_OUT, sql_SYSTEM_PUBLIC_TRANSACTION, blnSQL_DONE)
        Call SUB_PUT_LOG_TICK(strSQL)

        Call SUB_REMOVE_SQL_CONNECTION()

        If Not blnSQL_DONE Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING & Environment.NewLine & strSQL.ToString) 'デバッグログを追記
        End If

        Return datRET
    End Function
#End Region

#Region "INSERT,UPDATE関連"

    'クエリの実行
    Public Function FUNC_SYSTEM_DO_SQL_EXECUTE( _
    ByVal strSql As String, _
    Optional ByRef intHRetRow As Integer = 0, _
    Optional ByVal intTIME_OUT As Integer = -1 _
    ) As Boolean
        Dim blnRET As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        Call SUB_START_TICK()
        blnRET = FUNC_DO_SQL_EXECUTE(strSql, intHRetRow, intTIME_OUT, sql_SYSTEM_PUBLIC_TRANSACTION)
        Call SUB_PUT_LOG_TICK(strSql)

        Call SUB_REMOVE_SQL_CONNECTION()

        If Not blnRET Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING & Environment.NewLine & strSql.ToString) 'デバッグログを追記
        End If

        Return blnRET
    End Function

    'クエリの実行(複数回)
    Public Function FUNC_SYSTEM_DO_SQL_EXECUTE_RETRY( _
    ByVal strSql As String, _
    Optional ByVal intRETRY As Integer = 1, _
    Optional ByRef intHRetRow As Integer = 0, _
    Optional ByVal intTIME_OUT As Integer = -1 _
    ) As Boolean
        Dim blnRET As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        blnRET = FUNC_DO_SQL_EXECUTE_RETRY(strSql, intRETRY, intHRetRow, intTIME_OUT, sql_SYSTEM_PUBLIC_TRANSACTION)

        Call SUB_REMOVE_SQL_CONNECTION()

        If Not blnRET Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING & Environment.NewLine & strSql.ToString) 'デバッグログを追記
        End If

        Return blnRET
    End Function
#End Region

#Region "トランザクション関連"

    'トランザクション開始(BEGIN TRANS)
    Public Function FUNC_SYSTEM_BEGIN_TRANSACTION( _
    Optional ByVal lvlWIsoLationalLevel As IsolationLevel = IsolationLevel.ReadUncommitted _
    ) As Boolean
        Dim blnRET As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        blnRET = FUNC_BEGIN_TRANSACTION(sql_SYSTEM_PUBLIC_TRANSACTION, lvlWIsoLationalLevel)

        If Not blnRET Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING) 'デバッグログを追記
        End If

        Call SUB_REMOVE_SQL_CONNECTION()

        Return blnRET
    End Function

    'トランザクション確定(COMMIT)
    Public Function FUNC_SYSTEM_COMMIT_TRANSACTION() As Boolean
        Dim blnRET As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        blnRET = FUNC_COMMIT_TRANSACTION(sql_SYSTEM_PUBLIC_TRANSACTION)

        If Not blnRET Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING) 'デバッグログを追記
        End If

        Call SUB_REMOVE_SQL_CONNECTION()

        Return blnRET
    End Function

    'トランザクションキャンセル(ROLLBACK)
    Public Function FUNC_SYSTEM_ROLLBACK_TRANSACTION() As Boolean
        Dim blnRET As Boolean

        Call SUB_SET_SQL_CONNECTION(sql_SYSTEM_PUBLIC_CONNECTION)

        blnRET = FUNC_ROLLBACK_TRANSACTION(sql_SYSTEM_PUBLIC_TRANSACTION)

        If Not blnRET Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG(str_SQL_TOOL_LAST_ERR_STRING) 'デバッグログを追記
        End If

        Call SUB_REMOVE_SQL_CONNECTION()

        Return blnRET
    End Function

#End Region

#Region "その他共通等"
    'クエリ・トランザクションエラー時のメッセージを取得
    Public Function FUNC_SYSTEM_SQLGET_ERR_MESSAGE() As String
        Return str_SQL_TOOL_LAST_ERR_STRING
    End Function
#End Region

#Region "時間計測タイマー"
    Private intSTART_TICK As Integer
    '時間計測を開始
    Private Sub SUB_START_TICK()
        intSTART_TICK = System.Environment.TickCount
    End Sub

    '時間計測開始から何ミリ秒経過したかを取得
    Private Function FUNC_GET_TICK()
        Dim intRET As Integer
        Dim intCURRENT_TICK As Integer
        intCURRENT_TICK = System.Environment.TickCount
        intRET = intCURRENT_TICK - intSTART_TICK

        Return intRET
    End Function

    '経過時間が閾値以上の場合にログを出力
    'SUB_START_TICKと合わせて使用すること
    Private Const cstTICK_BASE As Integer = 400 '閾値基準(0.4秒)
    Private Const cstENABLED_SUB_PUT_LOG_TICK As Boolean = False  'TRUE：ログ出力有効 FALSE：ログ出力無効
    Private Sub SUB_PUT_LOG_TICK( _
    ByVal strLOG As String, _
    Optional ByVal intTICK_BASE As Integer = cstTICK_BASE _
    )
        Dim intTICK As Integer

        If Not cstENABLED_SUB_PUT_LOG_TICK Then
            Exit Sub
        End If

        If intSTART_TICK = 0 Then
            Exit Sub '初期化エラーフック
        End If

        intTICK = FUNC_GET_TICK()
        If intTICK >= intTICK_BASE Then
            Call SUB_SYSTEM_LOG_PUT_DEBUG("--------------------------------" & Environment.NewLine & "経過時間：" & intTICK / 1000 & "秒" & Environment.NewLine & "--------------------------------" & Environment.NewLine & strLOG)
        End If

    End Sub
#End Region

End Module
