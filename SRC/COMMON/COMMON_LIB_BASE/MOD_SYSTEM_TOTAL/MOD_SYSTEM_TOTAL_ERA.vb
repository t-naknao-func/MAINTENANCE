Public Module MOD_SYSTEM_TOTAL_ERA
    '全てのメソッドにおいて、共通システム変数「srtSYSTEM_TOTAL_MNG_M_ERA」に値が入っていることが前提となります
    '内部的なテーブル取得関係は全てキャッシュ処理を行っている為、レコードの値を変更した場合は、システムを一度終了してください

    '和暦名称の取得(明治、大正、昭和、平成)
    Public Function FUNC_GET_SYSTEM_TOTAL_NAME_ERA(ByVal datDATE_SEIREKi As DateTime) As String
        Dim intCODE_ERA As Integer
        Dim strRET As String

        intCODE_ERA = FUNC_GET_CODE_ERA(datDATE_SEIREKi)

        strRET = FUNC_GET_MNG_M_ERA_NAME_ERA(intCODE_ERA, True)

        Return strRET
    End Function

    '和暦略称の取得(m、t、s、h)
    Public Function FUNC_GET_SYSTEM_TOTAL_MARK_ERA(ByVal datDATE_SEIREKi As DateTime) As String
        Dim intCODE_ERA As Integer
        Dim strRET As String

        intCODE_ERA = FUNC_GET_CODE_ERA(datDATE_SEIREKi)

        strRET = FUNC_GET_MNG_M_ERA_MARK_ERA(intCODE_ERA, True)

        Return strRET
    End Function

    '和暦の取得(数値型6桁 例：260131)
    Public Function FUNC_GET_SYSTEM_TOTAL_NUMBER_ERA(ByVal datDATE_SEIREKi As DateTime) As Integer
        Dim intCODE_ERA As Integer
        Dim datDATE_START As DateTime
        Dim intRET As Integer
        Dim intYEAR_ERA As Integer

        intCODE_ERA = FUNC_GET_CODE_ERA(datDATE_SEIREKi)

        datDATE_START = FUNC_GET_MNG_M_ERA_DATE_START(intCODE_ERA, True)

        intYEAR_ERA = (datDATE_SEIREKi.Year - datDATE_START.Year) + 1

        intRET = intYEAR_ERA * 10000 + datDATE_SEIREKi.Month * 100 + datDATE_SEIREKi.Day

        Return intRET
    End Function

    '和暦連番の取得
    Private Function FUNC_GET_CODE_ERA(ByVal datDATE_IN As DateTime) As Integer
        Return 0
    End Function
End Module
