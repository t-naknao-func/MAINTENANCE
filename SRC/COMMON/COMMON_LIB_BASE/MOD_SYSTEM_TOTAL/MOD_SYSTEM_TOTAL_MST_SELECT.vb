
#Region "MNG_M_FLAG"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_FLAG

End Module

#End Region

#Region "MNG_M_KIND"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_KIND

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_KIND"
#End Region

    Private srtCASH_GET_MNG_M_KIND_NAME_KIND() As SRT_CASH_INT_INT_STR
    Public Function FUNC_GET_MNG_M_KIND_NAME_KIND( _
    ByVal enmCODE_FLAG As ENM_MNG_M_KIND_CODE_FLAG, _
    ByVal intCODE_KIND As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_KIND_NAME_KIND
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT_STR(srtMY_CASH, enmCODE_FLAG, intCODE_KIND)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_KIND"
            ReDim .WHERE(2)
            .WHERE(1).COL_NAME = "CODE_FLAG"
            .WHERE(1).VALUE = CInt(enmCODE_FLAG.GetHashCode)
            .WHERE(2).COL_NAME = "CODE_KIND"
            .WHERE(2).VALUE = intCODE_KIND
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_INT_STR(srtCASH_GET_MNG_M_KIND_NAME_KIND, enmCODE_FLAG, intCODE_KIND, strRET)

        Return strRET
    End Function

End Module

#End Region

#Region "MNG_M_STORE"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_SOTRE

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_STORE"
#End Region

    Private srtCASH_GET_MNG_M_STORE_NAME_STORE() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_STORE_NAME_STORE( _
    ByVal intCODE_STORE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_STORE_NAME_STORE
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_STORE)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_STORE"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = intCODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_STORE, strRET)

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_STORE_CODE_CITY() As SRT_CASH_INT_INT
    Public Function FUNC_GET_MNG_M_STORE_CODE_CITY( _
    ByVal intCODE_STORE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Long
        Dim lngRET As Long
        Dim srtMY_CASH() As SRT_CASH_INT_INT
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        lngRET = -1

        srtMY_CASH = srtCASH_GET_MNG_M_STORE_CODE_CITY
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT(srtMY_CASH, intCODE_STORE)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "CODE_CITY"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = intCODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        lngRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL)

        Call SUB_ADD_CASH_INT_INT(srtMY_CASH, intCODE_STORE, lngRET)

        Return lngRET
    End Function

    Private srtCASH_GET_MNG_M_STORE_CODE_KAMOKU() As SRT_CASH_INT_INT
    Public Function FUNC_GET_MNG_M_STORE_CODE_KAMOKU( _
    ByVal intCODE_STORE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Integer
        Dim intRET As Integer
        Dim srtMY_CASH() As SRT_CASH_INT_INT
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        intRET = -1

        srtMY_CASH = srtCASH_GET_MNG_M_STORE_CODE_KAMOKU
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT(srtMY_CASH, intCODE_STORE)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "CODE_KAMOKU"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = intCODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        intRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL)

        Call SUB_ADD_CASH_INT_INT(srtMY_CASH, intCODE_STORE, intRET)

        Return intRET
    End Function

    Private srtCASH_GET_MNG_M_STORE_CODE_KOUZA() As SRT_CASH_INT_DEC
    Public Function FUNC_GET_MNG_M_STORE_CODE_KOUZA( _
    ByVal intCODE_STORE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Long
        Dim lngRET As Long
        Dim srtMY_CASH() As SRT_CASH_INT_DEC
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        lngRET = -1

        srtMY_CASH = srtCASH_GET_MNG_M_STORE_CODE_KOUZA
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_DEC(srtMY_CASH, intCODE_STORE)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "CODE_KOUZA"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = intCODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        lngRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL)

        Call SUB_ADD_CASH_INT_DEC(srtMY_CASH, intCODE_STORE, lngRET)

        Return lngRET
    End Function

    Private srtCASH_GET_MNG_M_STORE_NAME_SENDFINANCE() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_STORE_NAME_SENDFINANCE( _
    ByVal intCODE_STORE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_STORE_NAME_SENDFINANCE
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_STORE)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_SENDFINANCE"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = intCODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_STORE, strRET)

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_STORE_NAME_SENDSTORE() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_STORE_NAME_SENDSTORE( _
    ByVal intCODE_STORE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_STORE_NAME_SENDSTORE
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_STORE)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_SENDSTORE"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = intCODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_STORE, strRET)

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_STORE_FLAG_DELETE() As SRT_CASH_INT_INT
    Public Function FUNC_GET_MNG_M_STORE_FLAG_DELETE( _
    ByVal intCODE_STORE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Integer
        Dim intRET As Integer
        Dim srtMY_CASH() As SRT_CASH_INT_INT
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        intRET = -1

        srtMY_CASH = srtCASH_GET_MNG_M_STORE_FLAG_DELETE
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT(srtMY_CASH, intCODE_STORE)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "FLAG_DELETE"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STORE"
            .WHERE(1).VALUE = intCODE_STORE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        intRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL)

        Call SUB_ADD_CASH_INT_INT(srtMY_CASH, intCODE_STORE, intRET)

        Return intRET
    End Function

    Public Function FUNC_CHECK_MNG_M_STORE( _
    ByVal intCODE_STORE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtKEY As SRT_TABLE_MNG_M_STORE_KEY
        Dim blnRET As Boolean

        srtKEY.CODE_STORE = intCODE_STORE

        blnRET = FUNC_CHECK_TABLE_MNG_M_STORE(srtKEY)

        Return blnRET
    End Function

End Module

#End Region

#Region "MNG_M_USER"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_USER

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_USER"
#End Region

    Public Function FUNC_GET_MNG_M_USER_CODE_STAFF( _
    ByVal strUSER_ID As String _
    ) As Integer
        Dim intRET As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        intRET = -1

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "CODE_STAFF"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "USER_ID"
            .WHERE(1).VALUE = strUSER_ID
            .ORDER_KEY = "CODE_STAFF"
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        intRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL)

        Return intRET
    End Function

    Public Function FUNC_CHECK_MNG_M_USER_USER_ID(
    ByVal STR_USER_ID As String, ByVal INT_CODE_STAFF As Integer
    ) As Boolean
        Dim STR_SQL As System.Text.StringBuilder
        STR_SQL = New System.Text.StringBuilder
        With STR_SQL
            Call .Append("SELECT" & Environment.NewLine)
            Call .Append("COUNT(*)" & Environment.NewLine)
            Call .Append("FROM" & Environment.NewLine)
            Call .Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT & " " & "WITH(NOLOCK)" & Environment.NewLine)
            Call .Append("WHERE" & Environment.NewLine)
            Call .Append("1=1" & Environment.NewLine)
            Call .Append("AND USER_ID" & "=" & FUNC_ADD_ENCLOSED_SCOT(STR_USER_ID))
            Call .Append("AND CODE_STAFF" & "<>" & INT_CODE_STAFF)
        End With

        Dim INT_COUNT As Integer

        INT_COUNT = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(STR_SQL.ToString, 0)

        Return (INT_COUNT > 0)
    End Function

    Private srtCASH_GET_MNG_M_USER_USER_ID() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_USER_USER_ID( _
    ByVal intCODE_STAFF As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_USER_USER_ID
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_STAFF)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "USER_ID"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STAFF"
            .WHERE(1).VALUE = intCODE_STAFF
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_STAFF, strRET)

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_USER_PASS_WORD() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_USER_PASS_WORD( _
    ByVal intCODE_STAFF As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_USER_PASS_WORD
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_STAFF)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "PASS_WORD"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STAFF"
            .WHERE(1).VALUE = intCODE_STAFF
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_STAFF, strRET)

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_USER_NAME_STAFF() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_USER_NAME_STAFF( _
    ByVal intCODE_STAFF As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_USER_NAME_STAFF
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_STAFF)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_STAFF"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STAFF"
            .WHERE(1).VALUE = intCODE_STAFF
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_STAFF, strRET)

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_USER_FLAG_DELETE() As SRT_CASH_INT_INT
    Public Function FUNC_GET_MNG_M_USER_FLAG_DELETE( _
    ByVal intCODE_STAFF As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Integer
        Dim intRET As Integer
        Dim srtMY_CASH() As SRT_CASH_INT_INT
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        intRET = 0

        srtMY_CASH = srtCASH_GET_MNG_M_USER_FLAG_DELETE
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT(srtMY_CASH, intCODE_STAFF)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "FLAG_DELETE"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STAFF"
            .WHERE(1).VALUE = intCODE_STAFF
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        intRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)

        Call SUB_ADD_CASH_INT_INT(srtMY_CASH, intCODE_STAFF, intRET)

        Return intRET
    End Function

    Private srtCASH_GET_MNG_M_USER_FLAG_GRANT() As SRT_CASH_INT_INT
    Public Function FUNC_GET_MNG_M_USER_FLAG_GRANT( _
    ByVal intCODE_STAFF As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Integer
        Dim intRET As Integer
        Dim srtMY_CASH() As SRT_CASH_INT_INT
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        intRET = 0

        srtMY_CASH = srtCASH_GET_MNG_M_USER_FLAG_GRANT
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT(srtMY_CASH, intCODE_STAFF)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "FLAG_GRANT"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_STAFF"
            .WHERE(1).VALUE = intCODE_STAFF
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        intRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_NUMERIC(strSQL, 0)

        Call SUB_ADD_CASH_INT_INT(srtMY_CASH, intCODE_STAFF, intRET)

        Return intRET
    End Function

End Module

#End Region

#Region "MNG_M_ACTIVE"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_ACTIVE

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_ACTIVE"
#End Region

    Public Function FUNC_GET_MNG_M_ACTIVE_DATE_ACTIVE( _
    ByVal intCODE_ACTIVE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "DATE_ACTIVE"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_ACTIVE"
            .WHERE(1).VALUE = intCODE_ACTIVE
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_DATETIME(strSQL)

        Return strRET
    End Function

    Public Function FUNC_CHECK_MNG_M_ACTIVE( _
    ByVal intCODE_ACTIVE As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtKEY As SRT_TABLE_MNG_M_ACTIVE_KEY
        Dim blnRET As Boolean

        srtKEY.CODE_ACTIVE = intCODE_ACTIVE

        blnRET = FUNC_CHECK_TABLE_MNG_M_ACTIVE(srtKEY)

        Return blnRET
    End Function

End Module

#End Region

#Region "MNG_M_CITY"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_CITY

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_CITY"
#End Region

    Private srtCASH_GET_MNG_M_CITY_NAME_CITY() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_CITY_NAME_CITY( _
    ByVal intCODE_CITY As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_CITY_NAME_CITY
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_CITY)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_CITY"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_CITY"
            .WHERE(1).VALUE = intCODE_CITY
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_CITY, strRET)

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_CITY_FLAG_DELETE() As SRT_CASH_INT_INT
    Public Function FUNC_GET_MNG_M_CITY_FLAG_DELETE( _
    ByVal intCODE_CITY As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_INT
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_CITY_FLAG_DELETE
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT(srtMY_CASH, intCODE_CITY)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "FLAG_DELETE"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_CITY"
            .WHERE(1).VALUE = intCODE_CITY
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_INT(srtMY_CASH, intCODE_CITY, strRET)

        Return strRET
    End Function

    Public Function FUNC_CHECK_MNG_M_CITY( _
    ByVal intCODE_CITY As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtKEY As SRT_TABLE_MNG_M_CITY_KEY
        Dim blnRET As Boolean

        srtKEY.CODE_CITY = intCODE_CITY

        blnRET = FUNC_CHECK_TABLE_MNG_M_CITY(srtKEY)

        Return blnRET
    End Function

End Module

#End Region

#Region "MNG_M_TRUST"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_TRUST

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_TRUST"
#End Region

    Private srtCASH_FUNC_CHECK_MNG_M_TRUST() As SRT_CASH_INT_BOOL
    Public Function FUNC_CHECK_MNG_M_TRUST( _
    ByVal intCODE_TRUST As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtKEY As SRT_TABLE_MNG_M_TRUST_KEY
        Dim blnRET As Boolean
        Dim srtMY_CASH() As SRT_CASH_INT_BOOL
        Dim intCASH_INDEX As Integer

        srtMY_CASH = srtCASH_FUNC_CHECK_MNG_M_TRUST
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_BOOL(srtMY_CASH, intCODE_TRUST)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        srtKEY.CODE_TRUST = intCODE_TRUST

        blnRET = FUNC_CHECK_TABLE_MNG_M_TRUST(srtKEY)

        Call SUB_ADD_CASH_INT_BOOL(srtMY_CASH, intCODE_TRUST, blnRET)

        Return blnRET
    End Function

End Module

#End Region

#Region "MNG_M_ERA"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_ERA

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_ERA"
#End Region

    Private srtCASH_GET_MNG_M_ERA_NAME_ERA() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_ERA_NAME_ERA( _
    ByVal intCODE_ERA As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_ERA_NAME_ERA
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_ERA)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_ERA"
            .WHERE = FUNC_GET_WHERE_KEY(intCODE_ERA)
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_ERA, strRET)
        srtCASH_GET_MNG_M_ERA_NAME_ERA = srtMY_CASH '4.5.1対応

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_ERA_MARK_ERA() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_ERA_MARK_ERA( _
    ByVal intCODE_ERA As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_ERA_MARK_ERA
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_ERA)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "MARK_ERA"
            .WHERE = FUNC_GET_WHERE_KEY(intCODE_ERA)
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_ERA, strRET)
        srtCASH_GET_MNG_M_ERA_MARK_ERA = srtMY_CASH '4.5.1対応

        Return strRET
    End Function

    Private srtCASH_GET_MNG_M_ERA_DATE_START() As SRT_CASH_INT_DATE
    Public Function FUNC_GET_MNG_M_ERA_DATE_START( _
    ByVal intCODE_ERA As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As DateTime
        Dim datRET As DateTime
        Dim srtMY_CASH() As SRT_CASH_INT_DATE
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        datRET = cstVB_DATE_MIN

        srtMY_CASH = srtCASH_GET_MNG_M_ERA_DATE_START
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_DATE(srtMY_CASH, intCODE_ERA)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "DATE_START"
            .WHERE = FUNC_GET_WHERE_KEY(intCODE_ERA)
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        datRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_DATETIME(strSQL)

        Call SUB_ADD_CASH_INT_DATE(srtMY_CASH, intCODE_ERA, datRET)
        srtCASH_GET_MNG_M_ERA_DATE_START = srtMY_CASH '4.5.1対応

        Return datRET
    End Function

#Region "WHERE"
    Private Function FUNC_GET_WHERE_KEY(ByVal intCODE_ERA As Integer) As SRT_SQL_TOOL_SELECT_ONE_COL_WHERE()
        Dim whrRET() As SRT_SQL_TOOL_SELECT_ONE_COL_WHERE
        Dim intINDEX As Integer

        ReDim whrRET(0)

        intINDEX = whrRET.Length
        ReDim Preserve whrRET(intINDEX)
        whrRET(intINDEX).COL_NAME = "CODE_ERA"
        whrRET(intINDEX).VALUE = intCODE_ERA

        Return whrRET
    End Function
#End Region

End Module

#End Region

#Region "MNG_M_SYSTEM"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_SYSTEM

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_SYSTEM"
#End Region

    Private srtCASH_GET_MNG_M_SYSTEM_NAME_SYSTEM() As SRT_CASH_INT_STR
    Public Function FUNC_GET_MNG_M_SYSTEM_NAME_SYSTEM( _
    ByVal intCODE_SYSTEM As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_SYSTEM_NAME_SYSTEM
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_STR(srtMY_CASH, intCODE_SYSTEM)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_SYSTEM"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = intCODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_STR(srtMY_CASH, intCODE_SYSTEM, strRET)

        Return strRET
    End Function

    Private srtCASH_FUNC_CHECK_MNG_M_SYSTEM() As SRT_CASH_INT_BOOL
    Public Function FUNC_CHECK_MNG_M_SYSTEM( _
    ByVal intCODE_SYSTEM As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtKEY As SRT_TABLE_MNG_M_SYSTEM_KEY
        Dim blnRET As Boolean
        Dim srtMY_CASH() As SRT_CASH_INT_BOOL
        Dim intCASH_INDEX As Integer

        srtMY_CASH = srtCASH_FUNC_CHECK_MNG_M_SYSTEM
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_BOOL(srtMY_CASH, intCODE_SYSTEM)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        srtKEY.CODE_SYSTEM = intCODE_SYSTEM

        blnRET = FUNC_CHECK_TABLE_MNG_M_SYSTEM(srtKEY)

        Call SUB_ADD_CASH_INT_BOOL(srtMY_CASH, intCODE_SYSTEM, blnRET)

        Return blnRET
    End Function

End Module

#End Region

#Region "MNG_M_PROGRAM"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_PROGRAM

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_PROGRAM"
#End Region

    Private srtCASH_GET_MNG_M_PROGRAM_NAME_PROGRAM() As SRT_CASH_INT_INT_STR
    Public Function FUNC_GET_MNG_M_PROGRAM_NAME_PROGRAM( _
    ByVal intCODE_SYSTEM As Integer, _
    ByVal intCODE_PROGRAM As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_INT_STR
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_PROGRAM_NAME_PROGRAM
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT_STR(srtMY_CASH, intCODE_SYSTEM, intCODE_PROGRAM)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "NAME_PROGRAM"

            ReDim .WHERE(0)
            ReDim Preserve .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = intCODE_SYSTEM
            ReDim Preserve .WHERE(2)
            .WHERE(2).COL_NAME = "CODE_PROGRAM"
            .WHERE(2).VALUE = intCODE_PROGRAM

            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_INT_STR(srtMY_CASH, intCODE_SYSTEM, intCODE_PROGRAM, strRET)

        Return strRET
    End Function

    Private srtCASH_FUNC_CHECK_MNG_M_PROGRAM() As SRT_CASH_INT_INT_BOOL
    Public Function FUNC_CHECK_MNG_M_PROGRAM( _
    ByVal intCODE_SYSTEM As Integer, _
    ByVal intCODE_PROGRAM As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As Boolean
        Dim srtKEY As SRT_TABLE_MNG_M_PROGRAM_KEY
        Dim blnRET As Boolean
        Dim srtMY_CASH() As SRT_CASH_INT_INT_BOOL
        Dim intCASH_INDEX As Integer

        srtMY_CASH = srtCASH_FUNC_CHECK_MNG_M_PROGRAM
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT_BOOL(srtMY_CASH, intCODE_SYSTEM, intCODE_PROGRAM)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        srtKEY.CODE_SYSTEM = intCODE_SYSTEM
        srtKEY.CODE_PROGRAM = intCODE_PROGRAM
        blnRET = FUNC_CHECK_TABLE_MNG_M_PROGRAM(srtKEY)

        Call SUB_ADD_CASH_INT_INT_BOOL(srtMY_CASH, intCODE_SYSTEM, intCODE_PROGRAM, blnRET)

        Return blnRET
    End Function

End Module

#End Region

#Region "MNG_M_MONTH"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_MONTH

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_MONTH"
#End Region

    Private srtCASH_GET_MNG_M_MONTH_CODE_YYYYMM() As SRT_CASH_INT_INT
    Public Function FUNC_GET_MNG_M_MONTH_CODE_YYYYMM( _
    ByVal intCODE_SYSTEM As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_INT
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_MONTH_CODE_YYYYMM
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT(srtMY_CASH, intCODE_SYSTEM)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "CODE_YYYYMM"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = intCODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_INT(srtMY_CASH, intCODE_SYSTEM, strRET)

        Return strRET
    End Function
End Module

#End Region

#Region "MNG_M_FISCAL_YEAR"

Public Module MOD_SYSTEM_TOTAL_MST_SELECT_MNG_M_FISCAL_YEAR

#Region "モジュール用・定数"
    Private Const CST_TABLE_NAME_DEFAULT As String = "MNG_M_FISCAL_YEAR"
#End Region

    Private srtCASH_GET_MNG_M_FISCAL_YEAR_CODE_NENDO() As SRT_CASH_INT_INT
    Public Function FUNC_GET_MNG_M_FISCAL_YEAR_CODE_NENDO( _
    ByVal intCODE_SYSTEM As Integer, _
    Optional ByVal blnCASH As Boolean = False _
    ) As String
        Dim strRET As String
        Dim srtMY_CASH() As SRT_CASH_INT_INT
        Dim intCASH_INDEX As Integer
        Dim srtSQL As SRT_SQL_TOOL_SELECT_ONE_COL
        Dim strSQL As String

        strRET = ""

        srtMY_CASH = srtCASH_GET_MNG_M_FISCAL_YEAR_CODE_NENDO
        If blnCASH Then
            intCASH_INDEX = FUNC_SEARCH_CASH_INT_INT(srtMY_CASH, intCODE_SYSTEM)
            If intCASH_INDEX <> -1 Then
                Return srtMY_CASH(intCASH_INDEX).VALUE
            End If
        End If

        With srtSQL
            .TABLE_NAME = strSYSTEM_PUBLIC_MNGDB_PREFIX & CST_TABLE_NAME_DEFAULT
            .COL_NAME = "CODE_NENDO"
            ReDim .WHERE(1)
            .WHERE(1).COL_NAME = "CODE_SYSTEM"
            .WHERE(1).VALUE = intCODE_SYSTEM
            .ORDER_KEY = ""
        End With

        strSQL = FUNC_GET_SQL_TOOL_SELECT_ONE_COL(srtSQL)
        strRET = FUNC_SYSTEM_GET_SQL_SINGLE_VALUE_STRING(strSQL)

        Call SUB_ADD_CASH_INT_INT(srtMY_CASH, intCODE_SYSTEM, strRET)

        Return strRET
    End Function
End Module

#End Region