Public Module MOD_SYSTEM_TOTAL_XLS_TOOL

    'CSVファイルに「セパレータを","→TABへ置換」「ヘッダ付与」を行い、汎用的なXLSファイルに変換する
    Public strLAST_ERR_FUNC＿SYSTEM_TOTAL_MAKE_OPTIONAL_XLS_FILE As String
    Public Function FUNC＿SYSTEM_TOTAL_MAKE_OPTIONAL_XLS_FILE( _
    ByVal strSOURCE_PATH As String, ByVal strMAKE_PATH As String, ByVal strHEADER As String, _
    Optional ByVal strSOURCE_SEP As String = "," _
    ) As Boolean
        Dim strONE_ROW As String
        Dim strCSV_READER As System.IO.StreamReader
        Dim stwXLS_WRITER As System.IO.StreamWriter
        Dim strARRAY() As String
        Dim intLOOP_INDEX As Integer
        Dim strXLS_SEP As String

        strLAST_ERR_FUNC＿SYSTEM_TOTAL_MAKE_OPTIONAL_XLS_FILE = ""

        strXLS_SEP = CStr(Convert.ToChar(9))

        Try
            strCSV_READER = New System.IO.StreamReader(strSOURCE_PATH, System.Text.Encoding.Default)
        Catch ex As Exception
            strLAST_ERR_FUNC＿SYSTEM_TOTAL_MAKE_OPTIONAL_XLS_FILE = ex.Message
            Return False
        End Try

        intLOOP_INDEX = 0
        ReDim strARRAY(0)
        strARRAY(0) = strHEADER.Replace(strSOURCE_SEP, strXLS_SEP)

        Do
            strONE_ROW = strCSV_READER.ReadLine()
            If strONE_ROW = "" Then '対象レコードがなくなったら、ループを抜ける
                Exit Do
            End If
            strONE_ROW = strONE_ROW.Replace(strSOURCE_SEP, strXLS_SEP)
            intLOOP_INDEX += 1
            ReDim Preserve strARRAY(intLOOP_INDEX)
            strARRAY(intLOOP_INDEX) = strONE_ROW
        Loop

        Call strCSV_READER.Close()

        Try
            stwXLS_WRITER = New System.IO.StreamWriter(strMAKE_PATH, False, System.Text.Encoding.Default)
        Catch ex As Exception
            strLAST_ERR_FUNC＿SYSTEM_TOTAL_MAKE_OPTIONAL_XLS_FILE = ex.Message
            Return False
        End Try

        For intLOOP_INDEX = 0 To (strARRAY.Length - 1)
            stwXLS_WRITER.WriteLine(strARRAY(intLOOP_INDEX))
        Next

        stwXLS_WRITER.Close()

        Return True
    End Function

End Module
