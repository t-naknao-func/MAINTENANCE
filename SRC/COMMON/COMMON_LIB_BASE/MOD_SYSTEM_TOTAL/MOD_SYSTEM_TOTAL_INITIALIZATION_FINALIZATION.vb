Public Module MOD_SYSTEM_TOTAL_INITIALIZATION_FINALIZATION

#Region "外出変数"
    Public str_FUNC_SYSTEM_TOTAL_INITIALIZATION_MSG As String 'システム初期化処理のメッセージ
    Public str_FUNC_SYSTEM_TOTAL_FINALIZATION_MSG As String 'システム終了処理のメッセージ
#End Region

#Region "初期化処理関連"
    '初期化処理-呼出
    Public Function FUNC_SYSTEM_TOTAL_INITIALIZATION( _
    ByRef strCOMMAND_LINE() As String _
    ) As Boolean

        str_FUNC_SYSTEM_TOTAL_INITIALIZATION_MSG = ""

        If Not FUNC_SYSTEM_TOTAL_GET_COMMANDLINE(srtSYSTEM_TOTAL_COMMANDLINE, strCOMMAND_LINE) Then
            str_FUNC_SYSTEM_TOTAL_INITIALIZATION_MSG = "コマンドラインが不正です。"
            Return False
        End If

        If Not FUNC_SYSTEM_TOTAL_GET_CONFIG(srtSYSTEM_TOTAL_CONFIG_SETTINGS) Then
            str_FUNC_SYSTEM_TOTAL_INITIALIZATION_MSG = "システム設定情報が取得できませんでした"
            Return False
        End If
        
        If Not FUNC_MAKE_VARIABLE() Then
            str_FUNC_SYSTEM_TOTAL_INITIALIZATION_MSG = "システム内部設定の作成に失敗しました"
            Return False
        End If

        If Not FUNC_SYSTEM_CONNECT_DB() Then
            str_FUNC_SYSTEM_TOTAL_INITIALIZATION_MSG = "SQLServerに接続できませんでした"
            Return False
        End If

        If Not FUNC_GET_DB_VALUE() Then
            str_FUNC_SYSTEM_TOTAL_INITIALIZATION_MSG = "DB設定値の取得に失敗しました"
            Return False
        End If

        Return True
    End Function

    '初期化処理-フォントサイズ用
    Public Sub SUB_SYSTEM_TOTAL_INIT_FONT(ByRef SNG_FONT_SIZE As Single, ByVal SNG_FONT_SIZE_DEF As Single)

        If SNG_FONT_SIZE <> 0 Then
            Exit Sub
        End If
        'SNG_FONT_SIZE = 0
        'SNG_FONT_SIZE = 9
        'SNG_FONT_SIZE = 10
        SNG_FONT_SIZE = 11
    End Sub

#Region "内部処理"

    Private Function FUNC_GET_DB_VALUE() As Boolean
        Const cstCODE_ACTIVE As Integer = 1 '処理日取得用
        Dim blnCHECK As Boolean

        blnCHECK = FUNC_CHECK_MNG_M_ACTIVE(cstCODE_ACTIVE)
        If blnCHECK Then
            datSYSTEM_TOTAL_DATE_ACTIVE = FUNC_GET_MNG_M_ACTIVE_DATE_ACTIVE(cstCODE_ACTIVE)
        Else
            datSYSTEM_TOTAL_DATE_ACTIVE = System.DateTime.Today 'システム日付
        End If

        Return True
    End Function

    Private Function FUNC_MAKE_VARIABLE() As Boolean

        If srtSYSTEM_TOTAL_CONFIG_SETTINGS.LOCAL.MNGDB_CATALOG = "" Then
            strSYSTEM_PUBLIC_MNGDB_PREFIX = ""
        Else
            strSYSTEM_PUBLIC_MNGDB_PREFIX = srtSYSTEM_TOTAL_CONFIG_SETTINGS.LOCAL.MNGDB_CATALOG & "." & CST_SYSTEM_DB_OWNER & "."
        End If

        Return True
    End Function
#End Region

#End Region

#Region "終了処理関連"
    '終了処理-呼出
    Public Function FUNC_SYSTEM_TOTAL_FINALIZATION() As Boolean

        str_FUNC_SYSTEM_TOTAL_FINALIZATION_MSG = ""

        If Not FUNC_SYSTEM_UNCONNECT_DB() Then 'DB切断
            str_FUNC_SYSTEM_TOTAL_FINALIZATION_MSG = "SQLServerとの切断がエラーで終了しました"
            Return False
        End If

        Return True

    End Function
#End Region

#Region "DB接続関連"
    Public str_SYSTEM_TOOL_FUNC_SYSTEM_CONNECT_DB_LAST_ERR As String '当該関数の最終エラー
    Public Function FUNC_SYSTEM_CONNECT_DB( _
    ) As Boolean 'システムデータベースへの接続
        Dim strSERVER As String
        Dim strCATLOG As String
        Dim strUSER_ID As String
        Dim strPASSWORD As String
        Dim intTIMEOUT As Integer
        Dim strAPPL_NAME As String
        Dim strTERM_NAME As String

        strSERVER = srtSYSTEM_TOTAL_CONFIG_SETTINGS.DB.SERVER
        strCATLOG = srtSYSTEM_TOTAL_CONFIG_SETTINGS.DB.CATALOG
        strUSER_ID = srtSYSTEM_TOTAL_CONFIG_SETTINGS.DB.USER
        strPASSWORD = srtSYSTEM_TOTAL_CONFIG_SETTINGS.DB.PASSWORD
        intTIMEOUT = srtSYSTEM_TOTAL_CONFIG_SETTINGS.DB.TIMEOUT

        strAPPL_NAME = FUNC_PATH_TO_FILENAME(System.Windows.Forms.Application.ExecutablePath)
        strAPPL_NAME = FUNC_GET_FILENAME_REMOVE_EXCTENT(strAPPL_NAME)
        strTERM_NAME = ""

        If Not FUNC_CONNECT_SESSION_DB(sql_SYSTEM_PUBLIC_CONNECTION, strSERVER, strCATLOG, strUSER_ID, strPASSWORD, strAPPL_NAME, strTERM_NAME, intTIMEOUT) Then
            str_SYSTEM_TOOL_FUNC_SYSTEM_CONNECT_DB_LAST_ERR = str_DB_TOOL_LAST_ERR_STRING
            Return False
        End If

        Return True
    End Function

    Public str_SYSTEM_TOOL_FUNC_SYSTEM_UNCONNECT_DB_LAST_ERR As String '当該関数の最終エラー
    Public Function FUNC_SYSTEM_UNCONNECT_DB() As Boolean 'システムデータベースからの切断

        If Not FUNC_CLOSE_SESSION_DB(sql_SYSTEM_PUBLIC_CONNECTION) Then
            str_SYSTEM_TOOL_FUNC_SYSTEM_UNCONNECT_DB_LAST_ERR = str_DB_TOOL_LAST_ERR_STRING
            Return False
        End If

        Return True
    End Function
#End Region

End Module
