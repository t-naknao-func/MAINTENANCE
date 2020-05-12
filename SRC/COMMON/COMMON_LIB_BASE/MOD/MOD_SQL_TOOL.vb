'SQLサーバ用クエリ発行モジュール
Public Module MOD_SQL_TOOL

#Region "外出変数"
    'TEST TEST
    Public ecp_SQL_TOOL_LAST_ERR_EXCEPTION As SqlClient.SqlException 'エラー時の最終エラー保持用
    Public str_SQL_TOOL_LAST_ERR_STRING As String 'エラー時の最終エラー文字列
#End Region

#Region "共通オブジェクト"
    Private sql_SQL_TOOL_PUBLIC_CONNECTION As SqlClient.SqlConnection '共有DBコネクション(このモジュール内のすべての関数はこのコネクションがオープンされていることが前提です)
#End Region

#Region "コネクションとトランザクションの受渡"
    Public Sub SUB_SET_SQL_CONNECTION( _
    ByRef objCONNECTION As SqlClient.SqlConnection _
    )
        sql_SQL_TOOL_PUBLIC_CONNECTION = objCONNECTION
    End Sub

    Public Sub SUB_REMOVE_SQL_CONNECTION()
        sql_SQL_TOOL_PUBLIC_CONNECTION = Nothing
    End Sub
#End Region

#Region "SELECT関連"
    '結果セットを返す
    Friend Function FUNC_GET_SQL_DATA_READER( _
    ByVal strSql As String, _
    ByRef sdrDATA_READER_RET As SqlClient.SqlDataReader, _
    Optional ByVal hiReadOpt As CommandBehavior = CommandBehavior.Default, _
    Optional ByVal intTIME_OUT As Integer = -1, _
    Optional ByRef objTRANSACTION As SqlClient.SqlTransaction = Nothing _
    ) As Boolean
        Dim cmdSqlCommand As SqlClient.SqlCommand

        If intTIME_OUT = -1 Then
            intTIME_OUT = 30
        End If

        cmdSqlCommand = New SqlClient.SqlCommand

        Try
            With cmdSqlCommand
                .Connection = sql_SQL_TOOL_PUBLIC_CONNECTION
                If Not IsNothing(objTRANSACTION) Then
                    .Transaction = objTRANSACTION
                End If

                .CommandType = CommandType.Text
                .CommandText = strSql

                .CommandTimeout = intTIME_OUT
                sdrDATA_READER_RET = .ExecuteReader(hiReadOpt)
            End With
        Catch ex As SqlClient.SqlException
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = ex
            cmdSqlCommand = Nothing
            Return False
        Catch ex As Exception
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = Nothing
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            Return False
        End Try

        cmdSqlCommand.Dispose()
        cmdSqlCommand = Nothing

        Return True
    End Function

    '一行・一列の値を返す
    Friend Function FUNC_GET_SQL_SINGLE_VALUE( _
    ByVal strSql As String, _
    Optional ByVal intTIME_OUT As Integer = -1, _
    Optional ByRef objTRANSACTION As SqlClient.SqlTransaction = Nothing, _
    Optional ByRef blnSQL_DONE As Boolean = True _
    ) As Object
        Dim cmdSqlCommand As SqlClient.SqlCommand
        Dim objRET As Object

        If intTIME_OUT = -1 Then
            intTIME_OUT = 30
        End If

        cmdSqlCommand = New SqlClient.SqlCommand
        objRET = Nothing
        blnSQL_DONE = True '初期化
        Try
            With cmdSqlCommand
                .Connection = sql_SQL_TOOL_PUBLIC_CONNECTION
                If Not IsNothing(objTRANSACTION) Then
                    .Transaction = objTRANSACTION 'デフォルトトランザクション
                End If

                .CommandType = CommandType.Text
                .CommandText = strSql
                .CommandTimeout = intTIME_OUT
                objRET = .ExecuteScalar
            End With
        Catch ex As SqlClient.SqlException
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = ex
            blnSQL_DONE = False
        Catch ex As Exception
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = Nothing
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            blnSQL_DONE = False
        End Try

        cmdSqlCommand.Dispose()
        cmdSqlCommand = Nothing

        Return objRET
    End Function

    'FUNC_GET_SQL_SINGLE_VALUEのラップ(整数値用)
    Friend Function FUNC_GET_SQL_SINGLE_VALUE_NUMERIC( _
    ByVal strSQL As String, Optional ByVal lngMISS_VALUE As Long = -1, _
    Optional ByVal intTIME_OUT As Integer = -1, _
    Optional ByRef objTRANSACTION As SqlClient.SqlTransaction = Nothing, _
    Optional ByRef blnSQL_DONE As Boolean = True _
    ) As Long
        Dim objTEMP As Object
        Dim lngRET As Long

        objTEMP = FUNC_GET_SQL_SINGLE_VALUE(strSQL, intTIME_OUT, objTRANSACTION, blnSQL_DONE)

        If Not IsNumeric(objTEMP) Then
            Return lngMISS_VALUE
        End If

        lngRET = CLng(objTEMP)
        objTEMP = Nothing

        Return lngRET

    End Function

    'FUNC_GET_SQL_SINGLE_VALUEのラップ(小数値・数量値用)
    Friend Function FUNC_GET_SQL_SINGLE_VALUE_DECIMAL( _
    ByVal strSQL As String, Optional ByVal decMISS_VALUE As Decimal = 0, _
    Optional ByVal intTIME_OUT As Integer = -1, _
    Optional ByRef objTRANSACTION As SqlClient.SqlTransaction = Nothing, _
    Optional ByRef blnSQL_DONE As Boolean = True _
    ) As Decimal
        Dim objTEMP As Object
        Dim decRET As Decimal

        objTEMP = FUNC_GET_SQL_SINGLE_VALUE(strSQL, intTIME_OUT, objTRANSACTION, blnSQL_DONE)

        If Not IsNumeric(objTEMP) Then
            Return decMISS_VALUE
        End If

        decRET = CDec(objTEMP)
        objTEMP = Nothing

        Return decRET

    End Function

    'FUNC_GET_SQL_SINGLE_VALUEのラップ(文字列用)
    Friend Function FUNC_GET_SQL_SINGLE_VALUE_STRING( _
    ByVal strSQL As String, Optional ByVal strMISS_VALUE As String = "", _
    Optional ByVal intTIME_OUT As Integer = -1, _
    Optional ByRef objTRANSACTION As SqlClient.SqlTransaction = Nothing, _
    Optional ByRef blnSQL_DONE As Boolean = True _
    ) As String
        Dim objTEMP As Object
        Dim strRET As String

        objTEMP = FUNC_GET_SQL_SINGLE_VALUE(strSQL, intTIME_OUT, objTRANSACTION, blnSQL_DONE)

        If IsNothing(objTEMP) Then
            Return strMISS_VALUE
        End If

        If IsDBNull(objTEMP) Then
            Return strMISS_VALUE
        End If

        strRET = CStr(objTEMP)
        objTEMP = Nothing

        Return strRET

    End Function

    'FUNC_GET_SQL_SINGLE_VALUEのラップ(日付用)
    Friend Function FUNC_GET_SQL_SINGLE_VALUE_DATETIME( _
    ByVal strSQL As String, Optional ByVal datMISS_VALUE As DateTime = cstVB_DATE_MIN, _
    Optional ByVal intTIME_OUT As Integer = -1, _
    Optional ByRef objTRANSACTION As SqlClient.SqlTransaction = Nothing, _
    Optional ByRef blnSQL_DONE As Boolean = True _
    ) As DateTime
        Dim objTEMP As Object
        Dim datRET As DateTime

        objTEMP = FUNC_GET_SQL_SINGLE_VALUE(strSQL, intTIME_OUT, objTRANSACTION, blnSQL_DONE)

        If Not IsDate(objTEMP) Then
            Return datMISS_VALUE
        End If

        datRET = CDate(objTEMP)
        objTEMP = Nothing

        Return datRET

    End Function

#End Region

#Region "INSERT,UPDATE関連"
    'SQL実行
    Friend Function FUNC_DO_SQL_EXECUTE( _
    ByVal strSql As String, _
    Optional ByRef intHRetRow As Integer = 0, _
    Optional ByVal intTIME_OUT As Integer = -1, _
    Optional ByRef objTRANSACTION As SqlClient.SqlTransaction = Nothing _
    ) As Boolean

        Dim cmdSqlCommand As SqlClient.SqlCommand

        If intTIME_OUT = -1 Then
            intTIME_OUT = 30
        End If

        cmdSqlCommand = New SqlClient.SqlCommand

        Try
            With cmdSqlCommand
                .Connection = sql_SQL_TOOL_PUBLIC_CONNECTION
                .CommandText = strSql
                .CommandTimeout = intTIME_OUT
                If Not IsNothing(objTRANSACTION) Then
                    .Transaction = objTRANSACTION 'デフォルトトランザクション
                End If
            End With

            intHRetRow = cmdSqlCommand.ExecuteNonQuery()
        Catch ex As SqlClient.SqlException
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = ex
            cmdSqlCommand = Nothing
            Return False
        Catch ex As Exception
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = Nothing
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            Return False
        End Try

        cmdSqlCommand.Dispose()
        cmdSqlCommand = Nothing

        Return True
    End Function

    '上記SQL実行関数の繰り返し用
    Friend Function FUNC_DO_SQL_EXECUTE_RETRY( _
    ByVal strSql As String, _
    Optional ByVal intRETRY As Integer = 1, _
    Optional ByRef intHRetRow As Integer = 0, _
    Optional ByVal intTIME_OUT As Integer = -1, _
    Optional ByRef objTRANSACTION As SqlClient.SqlTransaction = Nothing _
    ) As Boolean
        Dim intLOOP_INDEX As Integer
        Dim blnRET As Boolean

        blnRET = False

        For intLOOP_INDEX = 1 To intRETRY
            If FUNC_DO_SQL_EXECUTE(strSql, intHRetRow, intTIME_OUT, objTRANSACTION) Then
                blnRET = True
                Exit For
            End If
            intHRetRow = 0
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = Nothing
            str_SQL_TOOL_LAST_ERR_STRING = ""
            Call System.Windows.Forms.Application.DoEvents()
        Next

        Return blnRET
    End Function

#End Region

#Region "トランザクション関連"

    Friend Function FUNC_BEGIN_TRANSACTION( _
    ByRef objTRANSACTION As SqlClient.SqlTransaction, _
    Optional ByVal lvlWIsoLationalLevel As IsolationLevel = IsolationLevel.ReadUncommitted _
    ) As Boolean
        Dim strAPPL_NAME As String

        strAPPL_NAME = FUNC_PATH_TO_FILENAME(System.Windows.Forms.Application.ExecutablePath)
        strAPPL_NAME = FUNC_GET_FILENAME_REMOVE_EXCTENT(strAPPL_NAME)

        Try
            objTRANSACTION = sql_SQL_TOOL_PUBLIC_CONNECTION.BeginTransaction(lvlWIsoLationalLevel, strAPPL_NAME)
        Catch ex As SqlClient.SqlException
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = ex
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            objTRANSACTION = Nothing
            Return False
        Catch ex As Exception
            objTRANSACTION = Nothing
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            Return False
        End Try

        Return True
    End Function

    Friend Function FUNC_COMMIT_TRANSACTION( _
    ByRef objTRANSACTION As SqlClient.SqlTransaction _
    ) As Boolean

        Try
            If IsNothing(objTRANSACTION) Then
                Return False
            End If
            Call objTRANSACTION.Commit()
            objTRANSACTION = Nothing
        Catch ex As SqlClient.SqlException
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = ex
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            objTRANSACTION = Nothing
            Return False
        Catch ex As Exception
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = Nothing
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            objTRANSACTION = Nothing
            Return False
        End Try

        Return True
    End Function

    Friend Function FUNC_ROLLBACK_TRANSACTION( _
    ByRef objTRANSACTION As SqlClient.SqlTransaction _
    ) As Boolean

        Try
            If IsNothing(objTRANSACTION) Then
                Return False
            End If
            Call objTRANSACTION.Rollback()
            objTRANSACTION = Nothing
        Catch ex As SqlClient.SqlException
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = ex
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            objTRANSACTION = Nothing
            Return False
        Catch ex As Exception
            ecp_SQL_TOOL_LAST_ERR_EXCEPTION = Nothing
            str_SQL_TOOL_LAST_ERR_STRING = ex.Message.ToString
            objTRANSACTION = Nothing
            Return False
        End Try

        Return True
    End Function

#End Region

#Region "クエリ作成補助関連"
    Public Structure SRT_SQL_TOOL_SELECT_ONE_COL_WHERE
        Public COL_NAME As String
        Public VALUE As Object 'ENUMを設定すると誤読するため、列挙定数を代入する場合は標準型にキャストしてください
    End Structure

    Public Structure SRT_SQL_TOOL_SELECT_ONE_COL
        Public TABLE_NAME As String
        Public COL_NAME As String
        Public ORDER_KEY As String
        Public WHERE() As SRT_SQL_TOOL_SELECT_ONE_COL_WHERE
    End Structure

    Public Structure SRT_SQL_TOOL_DELETE
        Public TABLE_NAME As String
        Public WHERE() As SRT_SQL_TOOL_SELECT_ONE_COL_WHERE
    End Structure

    Public Function FUNC_GET_SQL_TOOL_SELECT_ONE_COL(ByRef srtDATA As SRT_SQL_TOOL_SELECT_ONE_COL) As String
        Dim strSQL As System.Text.StringBuilder
        Dim strTEMP As String
        Dim intLOOP_INDEX As Integer

        strSQL = New System.Text.StringBuilder
        With strSQL
            Call .Append("SELECT" & Environment.NewLine)
            Call .Append(srtDATA.COL_NAME & Environment.NewLine)
            Call .Append("FROM" & Environment.NewLine)
            Call .Append(srtDATA.TABLE_NAME & " " & "WITH(NOLOCK)" & Environment.NewLine)
            Call .Append("WHERE" & Environment.NewLine)
            Call .Append("1=1" & Environment.NewLine)

            If Not (srtDATA.WHERE Is Nothing) Then
                For intLOOP_INDEX = 1 To (srtDATA.WHERE.Length - 1)
                    Select Case True
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is Integer
                            strTEMP = srtDATA.WHERE(intLOOP_INDEX).VALUE.ToString
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is Long
                            strTEMP = srtDATA.WHERE(intLOOP_INDEX).VALUE.ToString
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is Decimal
                            strTEMP = srtDATA.WHERE(intLOOP_INDEX).VALUE.ToString
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is String
                            strTEMP = FUNC_ADD_ENCLOSED_SCOT(srtDATA.WHERE(intLOOP_INDEX).VALUE)
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is DateTime
                            strTEMP = FUNC_ADD_ENCLOSED_SCOT(srtDATA.WHERE(intLOOP_INDEX).VALUE.ToShortDateString)
                        Case Else
                            strTEMP = srtDATA.WHERE(intLOOP_INDEX).VALUE.ToString
                    End Select

                    Call .Append("AND" & " " & srtDATA.WHERE(intLOOP_INDEX).COL_NAME & "=" & strTEMP & Environment.NewLine)
                Next
            End If

            If srtDATA.ORDER_KEY <> "" Then
                Call .Append("ORDER BY" & Environment.NewLine)
                Call .Append(srtDATA.ORDER_KEY & Environment.NewLine)
            End If
        End With

        Return strSQL.ToString
    End Function

    Public Function FUNC_GET_SQL_TOOL_DELETE(ByRef srtDATA As SRT_SQL_TOOL_DELETE) As String
        Dim strSQL As System.Text.StringBuilder
        Dim strTEMP As String
        Dim intLOOP_INDEX As Integer

        strSQL = New System.Text.StringBuilder

        With strSQL
            Call .Append("DELETE" & Environment.NewLine)
            Call .Append("FROM" & Environment.NewLine)
            Call .Append(srtDATA.TABLE_NAME & " " & "WITH(ROWLOCK)" & Environment.NewLine)
            Call .Append("WHERE" & Environment.NewLine)
            Call .Append("1=1" & Environment.NewLine)

            If Not (srtDATA.WHERE Is Nothing) Then
                For intLOOP_INDEX = 1 To (srtDATA.WHERE.Length - 1)
                    Select Case True
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is Integer
                            strTEMP = srtDATA.WHERE(intLOOP_INDEX).VALUE.ToString
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is Long
                            strTEMP = srtDATA.WHERE(intLOOP_INDEX).VALUE.ToString
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is Decimal
                            strTEMP = srtDATA.WHERE(intLOOP_INDEX).VALUE.ToString
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is String
                            strTEMP = FUNC_ADD_ENCLOSED_SCOT(srtDATA.WHERE(intLOOP_INDEX).VALUE)
                        Case TypeOf srtDATA.WHERE(intLOOP_INDEX).VALUE Is DateTime
                            strTEMP = FUNC_ADD_ENCLOSED_SCOT(srtDATA.WHERE(intLOOP_INDEX).VALUE.ToShortDateString)
                        Case Else
                            strTEMP = srtDATA.WHERE(intLOOP_INDEX).VALUE.ToString
                    End Select

                    Call .Append("AND" & " " & srtDATA.WHERE(intLOOP_INDEX).COL_NAME & "=" & strTEMP & Environment.NewLine)
                Next
            End If
        End With

        Return strSQL.ToString
    End Function

    '特定の値からSQLサーバークエリで使用する文字列を取得
    Public Function FUNC_GET_VALUE_SQL_STRING(ByVal objVALUE As Object)
        Dim strRET As String

        Select Case True
            Case TypeOf objVALUE Is Integer
                strRET = objVALUE.ToString
            Case TypeOf objVALUE Is Long
                strRET = objVALUE.ToString
            Case TypeOf objVALUE Is Decimal
                strRET = objVALUE.ToString
            Case TypeOf objVALUE Is String
                'strRET = "N" & FUNC_ADD_ENCLOSED_SCOT(objVALUE)
                strRET = FUNC_ADD_ENCLOSED_SCOT(objVALUE)
            Case TypeOf objVALUE Is DateTime
                strRET = FUNC_ADD_ENCLOSED_SCOT(objVALUE.ToShortDateString)
            Case Else
                strRET = objVALUE.ToString
        End Select

        Return strRET
    End Function

    '日時をすべて挿入する場合専用
    Public Function FUNC_GET_STR_DATETIME(ByVal datVALUE As DateTime) As String
        Dim strRET As String

        strRET = ""
        strRET &= datVALUE.ToShortDateString & " "
        strRET &= datVALUE.ToLongTimeString

        Return strRET
    End Function

#End Region

End Module
