'SQLサーバ用モジュール
Public Module MOD_DB_TOOL

#Region "外出変数"
    Public ecp_DB_TOOL_LAST_ERR_EXCEPTION As SqlClient.SqlException '更新エラー時の最終エラー保持用
    Public str_DB_TOOL_LAST_ERR_STRING As String 'エラー時の最終エラー文字列
#End Region

#Region "セッション関連"
    'オープン
    Public Function FUNC_CONNECT_SESSION_DB( _
    ByRef sqlConect As SqlClient.SqlConnection, _
    ByVal strSERVER As String, ByVal strDB_NAME As String, _
    ByVal strUSER_ID As String, ByVal strPASSWORD As String, _
    ByVal strAPPL_NAME As String, ByVal strTERM_NAME As String, _
    ByVal intTIME_OUT As Integer _
    ) As Boolean

        Dim strWConnection As System.Text.StringBuilder

        strWConnection = New System.Text.StringBuilder
        strWConnection.Append("Server=" & strSERVER & ";")
        strWConnection.Append("Initial Catalog=" & strDB_NAME & ";")
        strWConnection.Append("User ID=" & strUSER_ID & ";")
        strWConnection.Append("Password=" & strPASSWORD & ";")

        strWConnection.Append("Application Name=" & strAPPL_NAME & ";")
        strWConnection.Append("Workstation ID=" & strTERM_NAME & ";")

        strWConnection.Append("Connect Timeout=" & intTIME_OUT & ";")

        If Not IsNothing(sqlConect) Then
            If sqlConect.State <> ConnectionState.Closed Then
                Call sqlConect.Close()
            End If
            Call sqlConect.Dispose()
            sqlConect = Nothing
        End If

        sqlConect = New SqlClient.SqlConnection

        Try
            sqlConect.ConnectionString = strWConnection.ToString()
            Call sqlConect.Open()
        Catch ex As SqlClient.SqlException
            ecp_DB_TOOL_LAST_ERR_EXCEPTION = ex
            str_DB_TOOL_LAST_ERR_STRING = ex.Message.ToString
            Return False
        Catch ex As Exception
            ecp_DB_TOOL_LAST_ERR_EXCEPTION = Nothing
            str_DB_TOOL_LAST_ERR_STRING = ex.Message.ToString
            Return False
        End Try

        Return True
    End Function

    'クローズ
    Public Function FUNC_CLOSE_SESSION_DB( _
    ByRef sqlConect As SqlClient.SqlConnection _
    ) As Boolean

        If IsNothing(sqlConect) Then
            Return True
        End If

        Try
            If Not sqlConect.State = ConnectionState.Closed Then
                Call sqlConect.Close()
            End If
        Catch ex As SqlClient.SqlException
            ecp_DB_TOOL_LAST_ERR_EXCEPTION = ex
            str_DB_TOOL_LAST_ERR_STRING = ex.Message.ToString
            Return False
        Catch ex As Exception
            ecp_DB_TOOL_LAST_ERR_EXCEPTION = Nothing
            str_DB_TOOL_LAST_ERR_STRING = ex.Message.ToString
            Return False
        End Try

        Call sqlConect.Dispose()

        sqlConect = Nothing

        Return True
    End Function
#End Region

End Module
