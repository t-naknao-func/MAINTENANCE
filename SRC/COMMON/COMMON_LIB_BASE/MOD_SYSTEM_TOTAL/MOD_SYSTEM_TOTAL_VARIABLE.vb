Public Module MOD_SYSTEM_TOTAL_VARIABLE

    Public srtSYSTEM_TOTAL_COMMANDLINE As SRT_SYSTEM_COMMANDLINE 'コマンドラインの内容
    Public srtSYSTEM_TOTAL_CONFIG_SETTINGS As SRT_SYSTEM_COMMON_SETTINGS '_SystemSettings.configの内容

#Region "システムの標準DBコネクション&トランザクション"
    Public sql_SYSTEM_PUBLIC_CONNECTION As SqlClient.SqlConnection '共有DBコネクション(このモジュール内のすべての関数はこのコネクションがオープンされていることが前提です)
    Public sql_SYSTEM_PUBLIC_TRANSACTION As SqlClient.SqlTransaction 'システム共有トランザクション(上記セッションと必ずセット)
#End Region

#Region "システムのローカル設定等"
    Public strSYSTEM_PUBLIC_MNGDB_PREFIX As String 'MNGDBへの接頭語(例：MNGDB.dbo.)
#End Region

#Region "初期処理でDBから取得するもの"
    Public datSYSTEM_TOTAL_DATE_ACTIVE As DateTime '処理日付(管理レコードがない場合は当日が入る)
#End Region

End Module
