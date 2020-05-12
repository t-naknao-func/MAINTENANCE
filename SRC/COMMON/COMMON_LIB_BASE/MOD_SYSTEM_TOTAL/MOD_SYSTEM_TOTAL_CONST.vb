Public Module MOD_SYSTEM_TOTAL_CONST

#Region "DB関連"
    Public Const CST_SYSTEM_DB_OWNER As String = "dbo"
#End Region

#Region "最小値・最大値関連"
    Public Const CST_SYSTEM_KINGAKU_MAX_VALUE As Long = 999999999999 '金額最大値(整数12桁)
    Public Const CST_SYSTEM_KINGAKU_THOUSAND_MAX_VALUE As Long = 999999999 '金額(千円)最大値(整数9桁)

    '職員コード(ユーザー)
    Public Const CST_SYSTEM_CODE_STAFF_MIN_VALUE As Integer = 1
    Public Const CST_SYSTEM_CODE_STAFF_MAX_VALUE As Integer = 999999999
    Public INT_SYSTEM_CODE_STAFF_MAX_LENGTH As Integer = (CST_SYSTEM_CODE_STAFF_MAX_VALUE.ToString.Length)

    '店舗コード
    Public Const CST_SYSTEM_CODE_STORE_MIN_VALUE As Integer = 1
    Public Const CST_SYSTEM_CODE_STORE_MAX_VALUE As Integer = 999

    '市町村コード
    Public Const CST_SYSTEM_CODE_CITY_MIN_VALUE As Integer = 1
    Public Const CST_SYSTEM_CODE_CITY_MAX_VALUE As Integer = 999

    '委託者コード
    Public Const CST_SYSTEM_CODE_TRUST_MIN_VALUE As Integer = 1
    Public Const CST_SYSTEM_CODE_TRUST_MAX_VALUE As Integer = 99999

#End Region

End Module
