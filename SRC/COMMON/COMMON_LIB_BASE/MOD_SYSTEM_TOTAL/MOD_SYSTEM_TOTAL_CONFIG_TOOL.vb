Public Module MOD_SYSTEM_TOTAL_CONFIG_TOOL

    Public Structure SRT_SYSTEM_COMMON_SETTINGS
        Public DB As SRT_SYSTEM_COMMON_SETTINGS_DB_CONNECTION
        Public LIST As SRT_SYSTEM_COMMON_SETTINGS_LIST
        Public LOCAL As SRT_SYSTEM_COMMON_SETTINGS_LOCAL
    End Structure

    Public Structure SRT_SYSTEM_COMMON_SETTINGS_DB_CONNECTION
        Public SERVER As String
        Public CATALOG As String
        Public USER As String
        Public PASSWORD As String
        Public TIMEOUT As String
    End Structure

    Public Structure SRT_SYSTEM_COMMON_SETTINGS_LIST
        Public DIR_BIP_EXE As String
        Public DIR_ASSETS As String
        Public DIR_DATA As String
        Public DIR_ASSETS_SERVER As String
    End Structure

    Public Structure SRT_SYSTEM_COMMON_SETTINGS_LOCAL
        Public MNGDB_CATALOG As String
        Public DATE_SYSTEM_START As DateTime
        Public DATE_SYSTEM_REPLACE As DateTime
    End Structure

    '設定ファイル共通部分の取得
    Function FUNC_SYSTEM_TOTAL_GET_CONFIG(ByRef srtSETTINGS As SRT_SYSTEM_COMMON_SETTINGS) As Boolean
        Try
            With srtSETTINGS.DB
                .SERVER = FUNC_GET_APP_SETTINGS("common_server")
                .CATALOG = FUNC_GET_APP_SETTINGS("common_catalog")
                .USER = FUNC_GET_APP_SETTINGS("common_user")
                .PASSWORD = FUNC_GET_APP_SETTINGS("common_password")
                .TIMEOUT = FUNC_GET_APP_SETTINGS("common_timeout")
            End With

            With srtSETTINGS.LIST
                .DIR_BIP_EXE = FUNC_GET_APP_SETTINGS("common_dir_bip_exe")
                .DIR_ASSETS = FUNC_GET_APP_SETTINGS("common_dir_assets")
                .DIR_DATA = FUNC_GET_APP_SETTINGS("common_dir_data")
                .DIR_ASSETS_SERVER = FUNC_GET_APP_SETTINGS("common_dir_assets_server")
            End With

            With srtSETTINGS.LOCAL
                .MNGDB_CATALOG = FUNC_GET_APP_SETTINGS("common_mngdb_catalog")
                .DATE_SYSTEM_START = CDate(FUNC_GET_APP_SETTINGS("common_date_system_start"))
                .DATE_SYSTEM_REPLACE = CDate(FUNC_GET_APP_SETTINGS("common_date_system_replace"))
            End With
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    'コンフィグの確認
    Public Sub SUB_SYSTEM_TOTAL_SHOW_SETTINGS(ByVal srtSETTINGS As SRT_SYSTEM_COMMON_SETTINGS)
        Dim strMSG As String

        strMSG = ""
        With srtSETTINGS
            strMSG &= "DB" & Environment.NewLine
            strMSG &= "SERVER:" & .DB.SERVER & Environment.NewLine
            strMSG &= "CATALOG:" & .DB.CATALOG & Environment.NewLine
            strMSG &= "USER:" & .DB.USER & Environment.NewLine
            strMSG &= "PASSWORD:" & .DB.PASSWORD & Environment.NewLine
            strMSG &= "TIMEOUT:" & .DB.TIMEOUT & Environment.NewLine
        End With

        Call System.Windows.Forms.MessageBox.Show(strMSG, "コンフィグ内容確認")

    End Sub

End Module
