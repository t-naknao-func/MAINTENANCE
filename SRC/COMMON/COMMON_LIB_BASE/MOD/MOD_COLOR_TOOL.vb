Public Module MOD_COLOR_TOOL


    Private Enum ENM_COLOR_CODE_INDEX
        R = 0
        G = 1
        B = 2
    End Enum
    '"x,y,z"形式の文字列をカラー情報に変換
    Public Function FUNC_GET_COLOR_FROM_STR(ByVal strCOLOR As String) As System.Drawing.Color
        Const cstSEP As String = ","
        Dim strCOLOR_SEP() As String
        Dim strR As String
        Dim strG As String
        Dim strB As String
        Dim intR As Integer
        Dim intG As Integer
        Dim intB As Integer
        Dim colRET As System.Drawing.Color

        If strCOLOR Is Nothing Then
            Return System.Drawing.Color.Empty
        End If

        strCOLOR_SEP = strCOLOR.Split(cstSEP)

        strR = FUNC_GET_STR_VALUE_FROM_ROW(strCOLOR_SEP, ENM_COLOR_CODE_INDEX.R, "0")
        strG = FUNC_GET_STR_VALUE_FROM_ROW(strCOLOR_SEP, ENM_COLOR_CODE_INDEX.G, "0")
        strB = FUNC_GET_STR_VALUE_FROM_ROW(strCOLOR_SEP, ENM_COLOR_CODE_INDEX.B, "0")

        Try
            intR = CInt(strR)
            intG = CInt(strG)
            intB = CInt(strB)
        Catch ex As Exception
            Return System.Drawing.Color.Empty
        End Try

        Try
            colRET = System.Drawing.Color.FromArgb(intR, intG, intB)
        Catch ex As Exception
            Return System.Drawing.Color.Empty
        End Try

        Return colRET
    End Function

    'カラー情報を"x,y,z"形式の文字列に変換
    Public Function FUNC_GET_STR_FROM_COLOR(ByVal colCOLOR As System.Drawing.Color) As String
        Const cstSEP As String = ","
        Dim intR As Integer
        Dim intG As Integer
        Dim intB As Integer
        Dim strRET As String

        intR = colCOLOR.R
        intG = colCOLOR.G
        intB = colCOLOR.B

        strRET = intR & cstSEP & intG & cstSEP & intB

        Return strRET
    End Function

    Private Function FUNC_GET_STR_VALUE_FROM_ROW(ByRef strROW() As String, ByVal intINDEX As Integer, Optional ByVal strMISS As String = "") As String
        Dim strRET As String

        If strROW.Length > intINDEX Then
            strRET = strROW(intINDEX)
        Else
            strRET = strMISS
        End If

        Return strRET
    End Function

End Module
