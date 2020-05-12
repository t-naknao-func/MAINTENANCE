#Region "COMMON(全体共通用)"

Public Structure SRT_TABLE_INPRINT 'インプリント用構造体
    Public DATE_INSERT As DateTime
    Public DATE_UPDATE As DateTime
End Structure

#End Region

#Region "MNG_M_USER"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_USER

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_USER"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_USER_KEY
        Public CODE_STAFF As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_USER_DATA
        Public FLAG_GRANT As Integer
        Public USER_ID As String
        Public PASS_WORD As String
        Public NAME_STAFF As String
        Public FLAG_DELETE As Integer
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_USER
        Public KEY As SRT_TABLE_MNG_M_USER_KEY
        Public DATA As SRT_TABLE_MNG_M_USER_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_USER

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_USER, ByRef srtKEY As SRT_TABLE_MNG_M_USER_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_STAFF = srtKEY.CODE_STAFF Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_USER, ByRef srtCASH_ON As SRT_TABLE_MNG_M_USER _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_USER( _
    ByRef srtDATA As SRT_TABLE_MNG_M_USER_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_USER_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_USER

        With srtRET
            .FLAG_GRANT = -1
            .USER_ID = ""
            .PASS_WORD = ""
            .NAME_STAFF = ""
            .FLAG_DELETE = 0
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STAFF"
            .WHERE(1).VALUE = srtDATA.CODE_STAFF
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .FLAG_GRANT = CInt(sdrREADER.Item("FLAG_GRANT"))
            .USER_ID = CStr(sdrREADER.Item("USER_ID"))
            .PASS_WORD = CStr(sdrREADER.Item("PASS_WORD"))
            .NAME_STAFF = CStr(sdrREADER.Item("NAME_STAFF"))
            .FLAG_DELETE = CInt(sdrREADER.Item("FLAG_DELETE"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_USER( _
    ByRef srtDATA As SRT_TABLE_MNG_M_USER_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STAFF"
            .WHERE(1).VALUE = srtDATA.CODE_STAFF
        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_USER( _
    ByRef srtDATA As SRT_TABLE_MNG_M_USER _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_STAFF) & "," & Environment.NewLine)
        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_GRANT) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.USER_ID) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.PASS_WORD) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_STAFF) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_DELETE) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_USER( _
    ByRef srtDATA As SRT_TABLE_MNG_M_USER_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STAFF"
            .WHERE(1).VALUE = srtDATA.CODE_STAFF
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_STORE"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_STORE

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_STORE"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_STORE_KEY
        Public CODE_STORE As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_STORE_DATA
        Public FLAG_STORE As Integer
        Public NAME_STORE As String
        Public SHORT_STORE As String
        Public KANA_STORE As String
        Public CODE_CITY As Integer
        Public FLAG_FDMAKE As Integer
        Public CODE_POST As String
        Public ADDRESS_STORE01 As String
        Public ADDRESS_STORE02 As String
        Public NUMBER_TEL01 As String
        Public NUMBER_FAX01 As String
        Public NAME_MANAGER As String
        Public NAME_PRESIDENT As String
        Public NAME_SENDFINANCE As String
        Public NAME_SENDSTORE As String
        Public CODE_KAMOKU As Integer
        Public CODE_KOUZA As Long
        Public FLAG_DELETE As Integer
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_STORE
        Public KEY As SRT_TABLE_MNG_M_STORE_KEY
        Public DATA As SRT_TABLE_MNG_M_STORE_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_STORE

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_STORE, ByRef srtKEY As SRT_TABLE_MNG_M_STORE_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_STORE = srtKEY.CODE_STORE Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_STORE, ByRef srtCASH_ON As SRT_TABLE_MNG_M_STORE _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_STORE( _
    ByRef srtDATA As SRT_TABLE_MNG_M_STORE_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_STORE_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_STORE

        With srtRET
            .FLAG_STORE = 0
            .NAME_STORE = ""
            .SHORT_STORE = ""
            .KANA_STORE = ""
            .CODE_CITY = 0
            .FLAG_FDMAKE = 0
            .CODE_POST = ""
            .ADDRESS_STORE01 = ""
            .ADDRESS_STORE02 = ""
            .NUMBER_TEL01 = ""
            .NUMBER_FAX01 = ""
            .NAME_MANAGER = ""
            .NAME_PRESIDENT = ""
            .NAME_SENDFINANCE = ""
            .NAME_SENDSTORE = ""
            .CODE_KAMOKU = 0
            .CODE_KOUZA = 0
            .FLAG_DELETE = 0
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = srtDATA.CODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .FLAG_STORE = CInt(sdrREADER.Item("FLAG_STORE"))
            .NAME_STORE = CStr(sdrREADER.Item("NAME_STORE"))
            .SHORT_STORE = CStr(sdrREADER.Item("SHORT_STORE"))
            .KANA_STORE = CStr(sdrREADER.Item("KANA_STORE"))
            .CODE_CITY = CInt(sdrREADER.Item("CODE_CITY"))
            .FLAG_FDMAKE = CInt(sdrREADER.Item("FLAG_FDMAKE"))
            .CODE_POST = CStr(sdrREADER.Item("CODE_POST"))
            .ADDRESS_STORE01 = CStr(sdrREADER.Item("ADDRESS_STORE01"))
            .ADDRESS_STORE02 = CStr(sdrREADER.Item("ADDRESS_STORE02"))
            .NUMBER_TEL01 = CStr(sdrREADER.Item("NUMBER_TEL01"))
            .NUMBER_FAX01 = CStr(sdrREADER.Item("NUMBER_FAX01"))
            .NAME_MANAGER = CStr(sdrREADER.Item("NAME_MANAGER"))
            .NAME_PRESIDENT = CStr(sdrREADER.Item("NAME_PRESIDENT"))
            .NAME_SENDFINANCE = CStr(sdrREADER.Item("NAME_SENDFINANCE"))
            .NAME_SENDSTORE = CStr(sdrREADER.Item("NAME_SENDSTORE"))
            .CODE_KAMOKU = CInt(sdrREADER.Item("CODE_KAMOKU"))
            .CODE_KOUZA = CLng(sdrREADER.Item("CODE_KOUZA"))
            .FLAG_DELETE = CInt(sdrREADER.Item("FLAG_DELETE"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_STORE( _
    ByRef srtDATA As SRT_TABLE_MNG_M_STORE_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = srtDATA.CODE_STORE
        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_STORE( _
    ByRef srtDATA As SRT_TABLE_MNG_M_STORE _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_STORE) & "," & Environment.NewLine)
        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_STORE) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_STORE) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.SHORT_STORE) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.KANA_STORE) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_CITY) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_FDMAKE) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_POST) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.ADDRESS_STORE01) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.ADDRESS_STORE02) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NUMBER_TEL01) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NUMBER_FAX01) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_MANAGER) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_PRESIDENT) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_SENDFINANCE) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_SENDSTORE) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_KAMOKU) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_KOUZA) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_DELETE) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region


#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_STORE( _
    ByRef srtDATA As SRT_TABLE_MNG_M_STORE_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = srtDATA.CODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_CITY"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_CITY

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_CITY"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_CITY_KEY
        Public CODE_CITY As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_CITY_DATA
        Public NAME_CITY As String
        Public FLAG_DELETE As Integer
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_CITY
        Public KEY As SRT_TABLE_MNG_M_CITY_KEY
        Public DATA As SRT_TABLE_MNG_M_CITY_DATA
    End Structure


#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_CITY

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_CITY, ByRef srtKEY As SRT_TABLE_MNG_M_CITY_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_CITY = srtKEY.CODE_CITY Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_CITY, ByRef srtCASH_ON As SRT_TABLE_MNG_M_CITY _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_CITY( _
    ByRef srtDATA As SRT_TABLE_MNG_M_CITY_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_CITY_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_CITY

        With srtRET
            .NAME_CITY = ""
            .FLAG_DELETE = 0
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_CITY"
            .WHERE(1).VALUE = srtDATA.CODE_CITY
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .NAME_CITY = CStr(sdrREADER.Item("NAME_CITY"))
            .FLAG_DELETE = CInt(sdrREADER.Item("FLAG_DELETE"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_CITY( _
    ByRef srtDATA As SRT_TABLE_MNG_M_CITY_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_CITY"
            .WHERE(1).VALUE = srtDATA.CODE_CITY
        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_CITY( _
    ByRef srtDATA As SRT_TABLE_MNG_M_CITY _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_CITY) & "," & Environment.NewLine)
        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_CITY) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_DELETE) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_CITY( _
    ByRef srtDATA As SRT_TABLE_MNG_M_CITY_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_CITY"
            .WHERE(1).VALUE = srtDATA.CODE_CITY
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_TRUST"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_TRUST

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_TRUST"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_TRUST_KEY
        Public CODE_TRUST As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_TRUST_DATA
        Public NAME_TRUST As String
        Public FLAG_DELETE As Integer
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_TRUST
        Public KEY As SRT_TABLE_MNG_M_TRUST_KEY
        Public DATA As SRT_TABLE_MNG_M_TRUST_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_TRUST

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_TRUST, ByRef srtKEY As SRT_TABLE_MNG_M_TRUST_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_TRUST = srtKEY.CODE_TRUST Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_TRUST, ByRef srtCASH_ON As SRT_TABLE_MNG_M_TRUST _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_TRUST( _
    ByRef srtDATA As SRT_TABLE_MNG_M_TRUST_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_TRUST_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_TRUST

        With srtRET
            .NAME_TRUST = ""
            .FLAG_DELETE = 0
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_TRUST"
            .WHERE(1).VALUE = srtDATA.CODE_TRUST
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .NAME_TRUST = CStr(sdrREADER.Item("NAME_TRUST"))
            .FLAG_DELETE = CInt(sdrREADER.Item("FLAG_DELETE"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_TRUST( _
    ByRef srtDATA As SRT_TABLE_MNG_M_TRUST_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_TRUST"
            .WHERE(1).VALUE = srtDATA.CODE_TRUST
        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_TRUST( _
    ByRef srtDATA As SRT_TABLE_MNG_M_TRUST _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_TRUST) & "," & Environment.NewLine)
        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_TRUST) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_DELETE) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_TRUST( _
    ByRef srtDATA As SRT_TABLE_MNG_M_TRUST_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_TRUST"
            .WHERE(1).VALUE = srtDATA.CODE_TRUST
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_ACTIVE"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_ACTIVE

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_ACTIVE"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_ACTIVE_KEY
        Public CODE_ACTIVE As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_ACTIVE_DATA
        Public DATE_ACTIVE As DateTime
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_ACTIVE
        Public KEY As SRT_TABLE_MNG_M_ACTIVE_KEY
        Public DATA As SRT_TABLE_MNG_M_ACTIVE_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_ACTIVE

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_ACTIVE, ByRef srtKEY As SRT_TABLE_MNG_M_ACTIVE_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_ACTIVE = srtKEY.CODE_ACTIVE Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_ACTIVE, ByRef srtCASH_ON As SRT_TABLE_MNG_M_ACTIVE _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_ACTIVE( _
    ByRef srtDATA As SRT_TABLE_MNG_M_ACTIVE_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_ACTIVE_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_ACTIVE

        With srtRET
            .DATE_ACTIVE = cstVB_DATE_MIN
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_ACTIVE"
            .WHERE(1).VALUE = srtDATA.CODE_ACTIVE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .DATE_ACTIVE = CDate(sdrREADER.Item("DATE_ACTIVE"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_ACTIVE( _
    ByRef srtDATA As SRT_TABLE_MNG_M_ACTIVE_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_ACTIVE"
            .WHERE(1).VALUE = srtDATA.CODE_ACTIVE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_ERA"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_ERA

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_ERA"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_ERA_KEY
        Public CODE_ERA As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_ERA_DATA
        Public NAME_ERA As String
        Public MARK_ERA As String
        Public KANA_ERA As String
        Public DATE_START As DateTime
        Public DATE_END As DateTime
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_ERA
        Public KEY As SRT_TABLE_MNG_M_ERA_KEY
        Public DATA As SRT_TABLE_MNG_M_ERA_DATA
    End Structure

End Module

#End Region

#Region "MNG_M_SYSTEM"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_SYSTEM

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_SYSTEM"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_SYSTEM_KEY
        Public CODE_SYSTEM As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_SYSTEM_DATA
        Public SYSTEM_ID As String
        Public NAME_SYSTEM As String
        Public SHORT_SYSTEM As String
        Public KANA_SYSTEM As String
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_SYSTEM
        Public KEY As SRT_TABLE_MNG_M_SYSTEM_KEY
        Public DATA As SRT_TABLE_MNG_M_SYSTEM_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_SYSTEM

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_SYSTEM, ByRef srtKEY As SRT_TABLE_MNG_M_SYSTEM_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_SYSTEM = srtKEY.CODE_SYSTEM Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_SYSTEM, ByRef srtCASH_ON As SRT_TABLE_MNG_M_SYSTEM _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_SYSTEM( _
    ByRef srtDATA As SRT_TABLE_MNG_M_SYSTEM_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_SYSTEM_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_SYSTEM

        With srtRET
            .SYSTEM_ID = ""
            .NAME_SYSTEM = ""
            .SHORT_SYSTEM = ""
            .KANA_SYSTEM = ""
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .SYSTEM_ID = CStr(sdrREADER.Item("SYSTEM_ID"))
            .NAME_SYSTEM = CStr(sdrREADER.Item("NAME_SYSTEM"))
            .SHORT_SYSTEM = CStr(sdrREADER.Item("SHORT_SYSTEM"))
            .KANA_SYSTEM = CStr(sdrREADER.Item("KANA_SYSTEM"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_SYSTEM( _
    ByRef srtDATA As SRT_TABLE_MNG_M_SYSTEM_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_SYSTEM( _
    ByRef srtDATA As SRT_TABLE_MNG_M_SYSTEM _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_SYSTEM) & "," & Environment.NewLine)
        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.SYSTEM_ID) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_SYSTEM) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.SHORT_SYSTEM) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.KANA_SYSTEM) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_SYSTEM( _
    ByRef srtDATA As SRT_TABLE_MNG_M_SYSTEM_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_PROGRAM"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_PROGRAM

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_PROGRAM"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_PROGRAM_KEY
        Public CODE_SYSTEM As Integer
        Public CODE_PROGRAM As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_PROGRAM_DATA
        Public FLAG_LOCK As Integer
        Public PROGRAM_ID As String
        Public NAME_PROGRAM As String
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_PROGRAM
        Public KEY As SRT_TABLE_MNG_M_PROGRAM_KEY
        Public DATA As SRT_TABLE_MNG_M_PROGRAM_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_PROGRAM

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_PROGRAM, ByRef srtKEY As SRT_TABLE_MNG_M_PROGRAM_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_SYSTEM = srtKEY.CODE_SYSTEM _
                And .KEY.CODE_PROGRAM = srtKEY.CODE_PROGRAM _
                Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_PROGRAM, ByRef srtCASH_ON As SRT_TABLE_MNG_M_PROGRAM _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_PROGRAM( _
    ByRef srtDATA As SRT_TABLE_MNG_M_PROGRAM_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_PROGRAM_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_PROGRAM

        With srtRET
            .FLAG_LOCK = 0
            .PROGRAM_ID = ""
            .NAME_PROGRAM = ""
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"

            ReDim .WHERE(0)
            ReDim Preserve .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            ReDim Preserve .WHERE(2)
            .WHERE(2).COL_NAME = "CODE_PROGRAM"
            .WHERE(2).VALUE = srtDATA.CODE_PROGRAM

            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .FLAG_LOCK = CInt(sdrREADER.Item("FLAG_LOCK"))
            .PROGRAM_ID = CStr(sdrREADER.Item("PROGRAM_ID"))
            .NAME_PROGRAM = CStr(sdrREADER.Item("NAME_PROGRAM"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_PROGRAM( _
    ByRef srtDATA As SRT_TABLE_MNG_M_PROGRAM_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT

            ReDim .WHERE(0)
            ReDim Preserve .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            ReDim Preserve .WHERE(2)
            .WHERE(2).COL_NAME = "CODE_PROGRAM"
            .WHERE(2).VALUE = srtDATA.CODE_PROGRAM

        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_PROGRAM( _
    ByRef srtDATA As SRT_TABLE_MNG_M_PROGRAM _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY

            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_SYSTEM) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_PROGRAM) & "," & Environment.NewLine)

        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_LOCK) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.PROGRAM_ID) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.NAME_PROGRAM) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_PROGRAM( _
    ByRef srtDATA As SRT_TABLE_MNG_M_PROGRAM_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"

            ReDim .WHERE(0)
            ReDim Preserve .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            ReDim Preserve .WHERE(2)
            .WHERE(2).COL_NAME = "CODE_PROGRAM"
            .WHERE(2).VALUE = srtDATA.CODE_PROGRAM

            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_EXCLUSION"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_EXCLUSION

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_EXCLUSION"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_EXCLUSION_KEY
        Public CODE_SYSTEM As Integer
        Public CODE_PROGRAM As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_EXCLUSION_DATA
        Public FLAG_STATE As Integer
        Public CODE_STAFF As Integer
        Public DATE_UPDATE As String '共通メソッド「FUNC_GET_VALUE_SQL_STRING」の動作(時刻情報の切捨て)の為、型をStringとしている
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_EXCLUSION
        Public KEY As SRT_TABLE_MNG_M_EXCLUSION_KEY
        Public DATA As SRT_TABLE_MNG_M_EXCLUSION_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_EXCLUSION

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_EXCLUSION, ByRef srtKEY As SRT_TABLE_MNG_M_EXCLUSION_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_SYSTEM = srtKEY.CODE_SYSTEM _
                  And .KEY.CODE_PROGRAM = srtKEY.CODE_PROGRAM _
                  Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_EXCLUSION, ByRef srtCASH_ON As SRT_TABLE_MNG_M_EXCLUSION _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_EXCLUSION( _
    ByRef srtDATA As SRT_TABLE_MNG_M_EXCLUSION_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_EXCLUSION_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_EXCLUSION

        With srtRET
            .FLAG_STATE = 0
            .CODE_STAFF = 0
            .DATE_UPDATE = cstVB_DATE_MIN
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"

            ReDim .WHERE(0)
            ReDim Preserve .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            ReDim Preserve .WHERE(2)
            .WHERE(2).COL_NAME = "CODE_PROGRAM"
            .WHERE(2).VALUE = srtDATA.CODE_PROGRAM

            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .FLAG_STATE = CInt(sdrREADER.Item("FLAG_STATE"))
            .CODE_STAFF = CInt(sdrREADER.Item("CODE_STAFF"))
            .DATE_UPDATE = CDate(sdrREADER.Item("DATE_UPDATE"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_EXCLUSION( _
    ByRef srtDATA As SRT_TABLE_MNG_M_EXCLUSION_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT

            ReDim .WHERE(0)
            ReDim Preserve .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            ReDim Preserve .WHERE(2)
            .WHERE(2).COL_NAME = "CODE_PROGRAM"
            .WHERE(2).VALUE = srtDATA.CODE_PROGRAM

        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_EXCLUSION( _
    ByRef srtDATA As SRT_TABLE_MNG_M_EXCLUSION _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_SYSTEM) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_PROGRAM) & "," & Environment.NewLine)
        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.FLAG_STATE) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_STAFF) & "," & Environment.NewLine)
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.DATE_UPDATE) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_EXCLUSION( _
    ByRef srtDATA As SRT_TABLE_MNG_M_EXCLUSION_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"

            ReDim .WHERE(0)
            ReDim Preserve .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            ReDim Preserve .WHERE(2)
            .WHERE(2).COL_NAME = "CODE_PROGRAM"
            .WHERE(2).VALUE = srtDATA.CODE_PROGRAM

            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_MONTH"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_MONTH

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_MONTH"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_MONTH_KEY
        Public CODE_SYSTEM As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_MONTH_DATA
        Public CODE_YYYYMM As Integer
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_MONTH
        Public KEY As SRT_TABLE_MNG_M_MONTH_KEY
        Public DATA As SRT_TABLE_MNG_M_MONTH_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_MONTH

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_MONTH, ByRef srtKEY As SRT_TABLE_MNG_M_MONTH_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_SYSTEM = srtKEY.CODE_SYSTEM Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_MONTH, ByRef srtCASH_ON As SRT_TABLE_MNG_M_MONTH _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_MONTH( _
    ByRef srtDATA As SRT_TABLE_MNG_M_MONTH_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_MONTH_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_MONTH

        With srtRET
            .CODE_YYYYMM = 0
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .CODE_YYYYMM = CInt(sdrREADER.Item("CODE_YYYYMM"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_MONTH( _
    ByRef srtDATA As SRT_TABLE_MNG_M_MONTH_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_MONTH( _
    ByRef srtDATA As SRT_TABLE_MNG_M_MONTH _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_SYSTEM) & "," & Environment.NewLine)
        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_YYYYMM) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_MONTH( _
    ByRef srtDATA As SRT_TABLE_MNG_M_SYSTEM_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_FISCAL_YEAR"

Public Module MOD_SYSTEM_TOTAL_TABLE_STRUCTURE_MNG_M_FISCAL_YEAR

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_FISCAL_YEAR"
#End Region

#Region "KEY"
    Public Structure SRT_TABLE_MNG_M_FISCAL_YEAR_KEY
        Public CODE_SYSTEM As Integer
    End Structure
#End Region

#Region "DATA"
    Public Structure SRT_TABLE_MNG_M_FISCAL_YEAR_DATA
        Public CODE_NENDO As Integer
    End Structure
#End Region

    Public Structure SRT_TABLE_MNG_M_FISCAL_YEAR
        Public KEY As SRT_TABLE_MNG_M_FISCAL_YEAR_KEY
        Public DATA As SRT_TABLE_MNG_M_FISCAL_YEAR_DATA
    End Structure

#Region "CASH"
    Private CASH_TABLE() As SRT_TABLE_MNG_M_FISCAL_YEAR

    Private Function FUNC_SEARCH_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_FISCAL_YEAR, ByRef srtKEY As SRT_TABLE_MNG_M_FISCAL_YEAR_KEY _
    ) As Integer
        Dim intLOOP_INDEX As Integer

        If IsNothing(srtCASH) Then
            Return -1
        End If

        For intLOOP_INDEX = LBound(srtCASH) To UBound(srtCASH)
            With srtCASH(intLOOP_INDEX)
                If .KEY.CODE_SYSTEM = srtKEY.CODE_SYSTEM Then
                    Return intLOOP_INDEX
                End If
            End With
        Next

        Return -1
    End Function

    Private Sub SUB_ADD_CASH( _
    ByRef srtCASH() As SRT_TABLE_MNG_M_FISCAL_YEAR, ByRef srtCASH_ON As SRT_TABLE_MNG_M_FISCAL_YEAR _
    )
        Dim intSERACH As Integer
        Dim intINDEX As Integer

        intSERACH = FUNC_SEARCH_CASH(srtCASH, srtCASH_ON.KEY)
        If intSERACH <> -1 Then
            Exit Sub
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If
        ReDim Preserve srtCASH(intINDEX)

        srtCASH(intINDEX) = srtCASH_ON
    End Sub
#End Region

#Region "SELECT"
    Public Function FUNC_SELECT_TABLE_MNG_M_FISCAL_YEAR( _
    ByRef srtDATA As SRT_TABLE_MNG_M_FISCAL_YEAR_KEY, _
    ByRef srtRET As SRT_TABLE_MNG_M_FISCAL_YEAR_DATA, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim sdrREADER As SqlClient.SqlDataReader
        Dim intCASH_INDEX As Integer
        Dim srtCASH_ONE As SRT_TABLE_MNG_M_FISCAL_YEAR

        With srtRET
            .CODE_NENDO = 0
        End With

        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH(CASH_TABLE, srtDATA)
            If intCASH_INDEX <> -1 Then
                srtRET = CASH_TABLE(intCASH_INDEX).DATA
                Return True
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "*"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        sdrREADER = Nothing
        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL, sdrREADER, CommandBehavior.SingleRow) Then
            Return False
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Return True
        End If

        Call sdrREADER.Read()

        With srtRET
            .CODE_NENDO = CInt(sdrREADER.Item("CODE_NENDO"))
        End With

        Call sdrREADER.Close()
        sdrREADER = Nothing

        srtCASH_ONE.KEY = srtDATA
        srtCASH_ONE.DATA = srtRET

        Call SUB_ADD_CASH(CASH_TABLE, srtCASH_ONE)

        Return True
    End Function
#End Region

#Region "DELETE"
    Public Function FUNC_DELETE_TABLE_MNG_M_FISCAL_YEAR( _
    ByRef srtDATA As SRT_TABLE_MNG_M_FISCAL_YEAR_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_DELETE
        Dim strSQL As String

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
        End With

        strSQL = FUNC_GET_SQL_TOOL_DELETE(srtSQL)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "INSERT"
    Public Function FUNC_INSERT_TABLE_MNG_M_FISCAL_YEAR( _
    ByRef srtDATA As SRT_TABLE_MNG_M_FISCAL_YEAR _
    ) As Boolean
        Dim strSQL As System.Text.StringBuilder

        strSQL = New System.Text.StringBuilder
        Call strSQL.Append("INSERT" & Environment.NewLine)
        Call strSQL.Append("INTO" & Environment.NewLine)
        Call strSQL.Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(ROWLOCK)" & Environment.NewLine)
        Call strSQL.Append("VALUES" & Environment.NewLine)
        Call strSQL.Append("(" & Environment.NewLine)
        With srtDATA.KEY
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_SYSTEM) & "," & Environment.NewLine)
        End With
        With srtDATA.DATA
            Call strSQL.Append(FUNC_GET_VALUE_SQL_STRING(.CODE_NENDO) & Environment.NewLine)
        End With
        Call strSQL.Append(")" & Environment.NewLine)

        If Not FUNC_SYSTEM_DO_SQL_EXECUTE(strSQL.ToString) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region "CHECK"
    Public Function FUNC_CHECK_TABLE_MNG_M_MONTH( _
    ByRef srtDATA As SRT_TABLE_MNG_M_SYSTEM_KEY _
    ) As Boolean
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String
        Dim intCNT As Integer
        Dim blnRET As Boolean

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "COUNT(*)"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = srtDATA.CODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)

        intCNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)
        blnRET = (intCNT > 0)
        Return blnRET
    End Function
#End Region

End Module

#End Region
