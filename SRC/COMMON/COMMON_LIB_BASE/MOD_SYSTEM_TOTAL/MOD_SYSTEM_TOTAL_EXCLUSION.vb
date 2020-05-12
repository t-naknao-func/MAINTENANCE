Public Module MOD_SYSTEM_TOTAL_EXCLUSION

    'ロックされているかのチェック
    Public Function FUNC_SYSTEM_CHECK_EXCLUSION(ByVal intCODE_SYSTEM As Integer, ByVal intCODE_PROGRAM As Integer) As Boolean
        Dim blnRET As Boolean
        Dim srtPROGRAM As SRT_TABLE_MNG_M_PROGRAM
        Dim blnFLAG_LOCK As Boolean

        With srtPROGRAM.KEY
            .CODE_SYSTEM = intCODE_SYSTEM
            .CODE_PROGRAM = intCODE_PROGRAM
        End With
        srtPROGRAM.DATA = Nothing
        If Not FUNC_SELECT_TABLE_MNG_M_PROGRAM(srtPROGRAM.KEY, srtPROGRAM.DATA) Then 'プログラム管理レコードがない場合は
            Return False 'ロックチェックをしない(ロックされてないとする)
        End If

        blnFLAG_LOCK = FUNC_CAST_INT_TO_BOOL(srtPROGRAM.DATA.FLAG_LOCK)
        If Not blnFLAG_LOCK Then 'ロック対象プログラムでない場合は
            Return False 'ロックチェックをしない(ロックされてないとする)
        End If

        blnRET = FUNC_CHECK_EXCLUSION(intCODE_SYSTEM, intCODE_PROGRAM)
        Return blnRET
    End Function

    'ロックを行う
    Public Function FUNC_SYSTEM_LOCK_EXCLUSION(ByVal intCODE_SYSTEM As Integer, ByVal intCODE_PROGRAM As Integer, ByVal intCODE_STAFF As Integer) As Boolean
        Dim blnRET As Boolean
        Dim srtPROGRAM As SRT_TABLE_MNG_M_PROGRAM
        Dim blnFLAG_LOCK As Boolean

        With srtPROGRAM.KEY
            .CODE_SYSTEM = intCODE_SYSTEM
            .CODE_PROGRAM = intCODE_PROGRAM
        End With
        srtPROGRAM.DATA = Nothing
        If Not FUNC_SELECT_TABLE_MNG_M_PROGRAM(srtPROGRAM.KEY, srtPROGRAM.DATA) Then 'プログラム管理レコードがない場合は
            Return True 'ロック処理を行わない(正常にロック処理が行われた事にしてスルーする)
        End If

        blnFLAG_LOCK = FUNC_CAST_INT_TO_BOOL(srtPROGRAM.DATA.FLAG_LOCK)
        If Not blnFLAG_LOCK Then 'ロック対象プログラムでない場合は
            Return True 'ロック処理を行わない(正常にロック処理が行われた事にしてスルーする)
        End If

        blnRET = FUNC_LOCK_EXCLUSION(intCODE_SYSTEM, intCODE_PROGRAM, intCODE_STAFF)
        Return blnRET
    End Function

    'ロックを解除する
    Public Function FUNC_SYSTEM_UNLOCK_EXCLUSION(ByVal intCODE_SYSTEM As Integer, ByVal intCODE_PROGRAM As Integer, ByVal intCODE_STAFF As Integer) As Boolean
        Dim blnRET As Boolean

        Dim srtPROGRAM As SRT_TABLE_MNG_M_PROGRAM
        Dim blnFLAG_LOCK As Boolean

        With srtPROGRAM.KEY
            .CODE_SYSTEM = intCODE_SYSTEM
            .CODE_PROGRAM = intCODE_PROGRAM
        End With
        srtPROGRAM.DATA = Nothing
        If Not FUNC_SELECT_TABLE_MNG_M_PROGRAM(srtPROGRAM.KEY, srtPROGRAM.DATA) Then 'プログラム管理レコードがない場合は
            Return True 'ロック解除処理を行わない(正常にロック解除処理が行われた事にしてスルーする)
        End If

        blnFLAG_LOCK = FUNC_CAST_INT_TO_BOOL(srtPROGRAM.DATA.FLAG_LOCK)
        If Not blnFLAG_LOCK Then 'ロック対象プログラムでない場合は
            Return True 'ロック解除処理を行わない(正常にロック解除処理が行われた事にしてスルーする)
        End If

        blnRET = FUNC_UNLOCK_EXCLUSION(intCODE_SYSTEM, intCODE_PROGRAM, intCODE_STAFF)
        Return blnRET
    End Function

#Region "内部関数"

    'チェック実処理
    Public Function FUNC_CHECK_EXCLUSION(ByVal intCODE_SYSTEM As Integer, ByVal intCODE_PROGRAM As Integer) As Boolean
        Dim srtTABLE As SRT_TABLE_MNG_M_EXCLUSION
        Dim blnRET As Boolean

        With srtTABLE.KEY
            .CODE_SYSTEM = intCODE_SYSTEM
            .CODE_PROGRAM = intCODE_PROGRAM
        End With

        If Not FUNC_CHECK_TABLE_MNG_M_EXCLUSION(srtTABLE.KEY) Then
            Return False 'レコードがない場合はロックされない
        End If

        srtTABLE.DATA = Nothing
        If Not FUNC_SELECT_TABLE_MNG_M_EXCLUSION(srtTABLE.KEY, srtTABLE.DATA) Then
            Return False 'データ部が取得できない場合はロックされない
        End If

        blnRET = FUNC_CAST_INT_TO_BOOL(srtTABLE.DATA.FLAG_STATE) '状態フラグをBOOL変換

        Return blnRET
    End Function

    'ロック実処理
    Private Function FUNC_LOCK_EXCLUSION(ByVal intCODE_SYSTEM As Integer, ByVal intCODE_PROGRAM As Integer, ByVal intCODE_STAFF As Integer) As Boolean
        Dim intFLAG_STATE As Integer

        intFLAG_STATE = FUNC_CAST_BOOL_TO_INT(True)
        If Not FUNC_UPDATE_FLAG_STATE(intCODE_SYSTEM, intCODE_PROGRAM, intFLAG_STATE, intCODE_STAFF) Then
            Return False
        End If

        Return True
    End Function

    'ロック解除実処理
    Private Function FUNC_UNLOCK_EXCLUSION(ByVal intCODE_SYSTEM As Integer, ByVal intCODE_PROGRAM As Integer, ByVal intCODE_STAFF As Integer) As Boolean
        Dim intFLAG_STATE As Integer

        intFLAG_STATE = FUNC_CAST_BOOL_TO_INT(False)
        If Not FUNC_UPDATE_FLAG_STATE(intCODE_SYSTEM, intCODE_PROGRAM, intFLAG_STATE, intCODE_STAFF) Then
            Return False
        End If

        Return True
    End Function

    'フラグ更新
    Private Function FUNC_UPDATE_FLAG_STATE(ByVal intCODE_SYSTEM As Integer, ByVal intCODE_PROGRAM As Integer, ByVal intFLAG_STATE As Integer, ByVal intCODE_STAFF As Integer) As Boolean
        Dim strSQL As System.Text.StringBuilder
        Dim strDATE_UPDATE As String

        strDATE_UPDATE = System.DateTime.Today.ToShortDateString & " " & System.DateTime.Now.ToLongTimeString
        strSQL = New System.Text.StringBuilder
        With strSQL
            .Append("UPDATE" & System.Environment.NewLine)
            .Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & "MNG_M_EXCLUSION" & System.Environment.NewLine)
            .Append("SET" & System.Environment.NewLine)
            .Append("FLAG_STATE=" & FUNC_GET_VALUE_SQL_STRING(intFLAG_STATE) & "," & System.Environment.NewLine)
            .Append("CODE_STAFF=" & FUNC_GET_VALUE_SQL_STRING(intCODE_STAFF) & "," & System.Environment.NewLine)
            .Append("DATE_UPDATE=" & FUNC_GET_VALUE_SQL_STRING(strDATE_UPDATE) & System.Environment.NewLine)

            .Append("WHERE" & System.Environment.NewLine)
            .Append("CODE_SYSTEM=" & intCODE_SYSTEM & System.Environment.NewLine)
            .Append("AND CODE_PROGRAM=" & intCODE_PROGRAM & System.Environment.NewLine)
        End With

        If Not FUNC_SYSTEM_BEGIN_TRANSACTION() Then
            Return False
        End If

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Call FUNC_SYSTEM_ROLLBACK_TRANSACTION()
            Return False
        End If

        If Not FUNC_SYSTEM_COMMIT_TRANSACTION() Then
            Return False
        End If

        Return True
    End Function
#End Region
    
End Module
