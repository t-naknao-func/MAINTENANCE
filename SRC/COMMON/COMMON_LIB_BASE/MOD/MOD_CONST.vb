Public Module MOD_CONST

    Public STR_TXT_DB_DRIVER_NAME As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ="
    Public STR_TXT_DB_DRIVER_INI_NAME As String = "schema.ini"

    Public Const cstVB_DATE_MIN As DateTime = #1/1/1753# 'DateTimePicker上の最小日付≒VBコード上で扱いうる最小日付(基本的に日付型のエラートラップに使用する)
    Public Const cstVB_DATE_MAX As DateTime = #12/31/9998# 'DateTimePicker上の最大日付≒VBコード上で扱いうる最大日付(基本的に永久に到達しない日付型として使用する)
	'Public Const cstvbCrlf As String = Environment.NewLine
End Module
