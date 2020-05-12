Public Module MOD_SYSTEM_TOTAL_BIP_COMMAND

#Region "定数・帳票"
    Private Const CST_CSV_SEPARATOR As String = ","
#End Region

#Region "定数・帳票用エラーメッセージ"
    Public Const CST_SYSTEM_TOTAL_LIST_ERR_MSG_8001 As String = "データファイルを作成できませんでした。"
    Public Const CST_SYSTEM_TOTAL_LIST_ERR_MSG_8002 As String = "帳票定義体を取得できませんでした。"
    Public Const CST_SYSTEM_TOTAL_LIST_ERR_MSG_8003 As String = "印刷処理に失敗しました。"
    Public Const CST_SYSTEM_TOTAL_LIST_ERR_MSG_8004 As String = "プレビュー処理に失敗しました。"
    Public Const CST_SYSTEM_TOTAL_LIST_ERR_MSG_8005 As String = "ファイル出力処理に失敗しました。"
    Public Const CST_SYSTEM_TOTAL_LIST_ERR_MSG_8006 As String = "印刷可能なレコード件数を超えています。"

#End Region

#Region "構造体"
    Private Structure SRT_BIP_INFO
        Public ENVIRONMENT As SRT_BIP_INFO_ENVIRONMENT
        Public DATA_FORMAT() As SRT_BIP_INFO_DATA_FORMAT
    End Structure

    Private Structure SRT_BIP_INFO_ENVIRONMENT
        Public COMMENT As String
    End Structure

    Private Structure SRT_BIP_INFO_DATA_FORMAT
        Public NAME_COL As String
        Public OPERAND As String
    End Structure
#End Region

    'プレビュー処理
    Public Function FUNC_SHOW_PREVIEW_BIP(
    ByVal strDIR_BIP_EXE As String,
    ByVal strDIR_ASSETS As String,
    ByVal strDEFINITION As String,
    ByVal strPATH_DATA As String
    ) As Boolean
        Dim STR_EXE_PATH As String
        Const CST_PREVIEW As String = "Bipview.exe"
        STR_EXE_PATH = ""
        STR_EXE_PATH &= strDIR_BIP_EXE & "\"
        STR_EXE_PATH &= CST_PREVIEW

        Dim STR_COMMAND_LINE As String
        STR_COMMAND_LINE = ""
        STR_COMMAND_LINE &= """" & strDEFINITION & """"
        STR_COMMAND_LINE &= " " & "-assetsdir" & " " & strDIR_ASSETS
        STR_COMMAND_LINE &= " " & "-f" & " " & strPATH_DATA
        STR_COMMAND_LINE &= " " & "-indatacode" & " " & "SJIS"

        If Not FUNC_CALL_EXE_FILE_SHELL(STR_EXE_PATH, STR_COMMAND_LINE) Then
            Return False
        End If

        Return True
    End Function

    '印刷処理
    Public Function FUNC_PUT_LIST_BIP(
    ByVal strDIR_BIP_EXE As String,
    ByVal strDIR_ASSETS As String,
    ByVal strDEFINITION As String,
    ByVal strPATH_DATA As String
    ) As Boolean

        Dim STR_EXE_PATH As String
        Const CST_PRINT As String = "Bipprint.exe"
        STR_EXE_PATH = ""
        STR_EXE_PATH &= strDIR_BIP_EXE & "\"
        STR_EXE_PATH &= CST_PRINT

        Dim STR_COMMAND_LINE As String
        STR_COMMAND_LINE = ""
        STR_COMMAND_LINE &= """" & strDEFINITION & """"
        STR_COMMAND_LINE &= " " & "-assetsdir" & " " & strDIR_ASSETS
        STR_COMMAND_LINE &= " " & "-f" & " " & strPATH_DATA
        STR_COMMAND_LINE &= " " & "-indatacode" & " " & "SJIS"

        If Not FUNC_CALL_EXE_FILE_SHELL(STR_EXE_PATH, STR_COMMAND_LINE) Then
            Return False
        End If

        Return True
    End Function

    'ファイル保存ダイアログ表示
    Public Function FUNC_SHOW_PUT_FILE_DIALOG(ByRef STR_FILE_PATH As String, ByVal STR_FILE_NAME As String) As Boolean
        Dim SFD_FILE As System.Windows.Forms.SaveFileDialog

        SFD_FILE = New System.Windows.Forms.SaveFileDialog
        If Not (STR_FILE_NAME = "") Then
            SFD_FILE.FileName = STR_FILE_NAME
        End If
        SFD_FILE.DefaultExt = "csv"
        Dim STR_FILTER As String
        STR_FILTER = ""
        STR_FILTER &= "" & "CSVカンマ区切り(*.csv)|*.csv"
        'STR_FILTER &= "|" & "Excelブック(*.xlsx)|*.xlsx"
        'STR_FILTER &= "|" & "PDF(*.pdf)|*.pdf"
        SFD_FILE.Filter = STR_FILTER
        SFD_FILE.OverwritePrompt = True

        Dim RST_MSG As System.Windows.Forms.DialogResult
        RST_MSG = SFD_FILE.ShowDialog()
        If RST_MSG = Windows.Forms.DialogResult.Cancel Then
            Return False
        End If

        STR_FILE_PATH = SFD_FILE.FileName
        Call SFD_FILE.Dispose()
        SFD_FILE = Nothing

        Return True
    End Function

    'ファイル出力処理
    Public Function FUNC_PUT_FILE_BIP(
    ByVal strDIR_BIP_EXE As String,
    ByVal strDIR_ASSETS As String,
    ByVal strDEFINITION As String,
    ByVal strPATH_DATA As String,
    ByVal strPATH_PUT As String
    ) As Boolean

        Call FUNC_FILE_DELETE(strPATH_PUT)

        Dim STR_EXTENT As String
        STR_EXTENT = FUNC_GET_FILE_EXTENT(strPATH_PUT)
        STR_EXTENT = STR_EXTENT.ToLower

        Dim BLN_RET As Boolean
        Select Case STR_EXTENT
            Case ".xlsx"
                BLN_RET = FUNC_PUT_EXCEL_BIP(strDIR_BIP_EXE, strDIR_ASSETS, strDEFINITION, strPATH_DATA, strPATH_PUT)
            Case ".csv"
                BLN_RET = FUNC_FILE_COPY(strPATH_DATA, strPATH_PUT)
            Case ".pdf"
                BLN_RET = FUNC_PUT_PDF_BIP(strDIR_BIP_EXE, strDIR_ASSETS, strDEFINITION, strPATH_DATA, strPATH_PUT)
            Case Else
                BLN_RET = False
        End Select

        Return BLN_RET
    End Function

    '定義体のコピー処理
    Public Function FUNC_COPY_LIST_DEFINITION_BIP(
    ByVal strDEFINITION As String, ByVal strDIR_ASSETS_SERVER As String, ByVal strDIR_ASSETS_LOCAL As String
    ) As Boolean

        If strDIR_ASSETS_SERVER = strDIR_ASSETS_LOCAL Then '同一の場合はコピー不要
            Return True
        End If

        Dim strEXTENT() As String
        ReDim strEXTENT(4)
        strEXTENT(1) = "bip"
        strEXTENT(2) = "ovd"
        strEXTENT(3) = "pmd"
        strEXTENT(4) = "psf"

        Dim strFILE_NAME As String
        Dim intLOOP_INDEX As Integer
        Dim strFROM As String
        Dim strTO As String
        strFILE_NAME = ""
        For intLOOP_INDEX = 1 To (strEXTENT.Length - 1)
            strFILE_NAME = strDEFINITION & "." & strEXTENT(intLOOP_INDEX)
            strFROM = strDIR_ASSETS_SERVER & "\" & strFILE_NAME
            strTO = strDIR_ASSETS_LOCAL & "\" & strFILE_NAME
            If Not FUNC_FILE_COPY(strFROM, strTO) Then
                Return False
            End If
        Next

        Return True
    End Function

    '文字列配列をCSV1行に変換
    Public Function FUNC_GET_ONE_ROW_LIST_CSV(ByRef strROW() As String) As String
        Dim intLOOP_INDEX As Integer
        Dim intINDEX As Integer
        Dim strSEP As String
        Dim strONE_ROW As System.Text.StringBuilder

        If strROW Is Nothing Then
            Return ""
        End If

        strONE_ROW = New System.Text.StringBuilder
        intINDEX = (strROW.Length - 1)
        For intLOOP_INDEX = 1 To intINDEX
            strSEP = If(intLOOP_INDEX = intINDEX, "", CST_CSV_SEPARATOR)
            strONE_ROW.Append(strROW(intLOOP_INDEX) & strSEP)
        Next

        Return strONE_ROW.ToString
    End Function

    'プレビュープロセスにクローズ命令
    Public Sub SUB_LIST_PREVIEW_WINDOW_CLOSE_ALL()

        Dim p As Process
        p = Nothing

        Dim INT_PROCESS() As Integer
        ReDim INT_PROCESS(0)
        Do
            p = FUNC_GET_BIP_PREVIEW_PROCESS(, INT_PROCESS)
            If p Is Nothing Then
                Exit Do
            End If
            Dim INT_INDEX As Integer
            INT_INDEX = INT_PROCESS.Length
            ReDim Preserve INT_PROCESS(INT_INDEX)
            INT_PROCESS(INT_INDEX) = p.Id

            Dim BLN_CLOSE As Boolean
            BLN_CLOSE = p.CloseMainWindow() 'クローズメッセージを送信する
            If Not BLN_CLOSE Then
                Continue Do
            End If
            p.WaitForExit(100)
            If p.HasExited Then

            End If
        Loop
    End Sub

    'BIPファイルから項目名を取得
    Public Function FUNC_GET_NAME_COLS_FROM_BIP(ByVal STR_DIR_ASSETS As String, ByVal STR_DEFINITION As String) As String()
        Dim STR_RET() As String
        ReDim STR_RET(0)

        Dim SRT_BIP As SRT_BIP_INFO
        SRT_BIP = FUNC_GET_BIP_INFO(STR_DIR_ASSETS, STR_DEFINITION)

        For i = 1 To (SRT_BIP.DATA_FORMAT.Length - 1)
            If Not (SRT_BIP.DATA_FORMAT(i).OPERAND = "") Then
                Continue For
            End If

            Dim INT_INDEX As Integer
            INT_INDEX = STR_RET.Length
            ReDim Preserve STR_RET(INT_INDEX)
            STR_RET(INT_INDEX) = SRT_BIP.DATA_FORMAT(i).NAME_COL
        Next
        Return STR_RET
    End Function

#Region "内部処理"
    Private Function FUNC_CONVRET_CSV(ByVal STR_PATH_DATA As String, ByVal STR_PATH_MAKE As String) As Boolean
        Dim STR_CSV_READER As System.IO.StreamReader 'ファイル入力用のIOオブジェクト
        Try
            STR_CSV_READER = New System.IO.StreamReader(STR_PATH_DATA, System.Text.Encoding.UTF8)   'ファイルライターを開く
        Catch ex As Exception
            Return False
        End Try

        Dim STR_FILE() As String
        ReDim STR_FILE(0)
        While STR_CSV_READER.Peek() > -1
            Dim INT_INDEX As Integer
            INT_INDEX = STR_FILE.Length
            ReDim Preserve STR_FILE(INT_INDEX)
            STR_FILE(INT_INDEX) = STR_CSV_READER.ReadLine()
        End While

        Call STR_CSV_READER.Close() '閉じる

        Dim STR_PUT() As String
        ReDim STR_PUT(0)
        For i = 1 To STR_FILE.Length - 1
            Dim STR_LINE As String
            STR_LINE = STR_FILE(i)
            Dim STR_VALUE_ROW() As String
            STR_VALUE_ROW = STR_LINE.Split(CST_CSV_SEPARATOR)
            Dim STR_PUT_ROW() As String
            ReDim STR_PUT_ROW(STR_VALUE_ROW.Length - 1)
            For j = 0 To STR_VALUE_ROW.Length - 1
                Dim STR_VALUE As String
                STR_VALUE = FUNC_CONVERT_UTF_QUESTION(STR_VALUE_ROW(j))
                STR_PUT_ROW(j) = STR_VALUE
            Next

            STR_LINE = ""
            For j = 0 To STR_PUT_ROW.Length - 1
                STR_LINE &= STR_PUT_ROW(j) & If(j = (STR_PUT_ROW.Length - 1), "", CST_CSV_SEPARATOR)
            Next

            ReDim Preserve STR_PUT(i)
            STR_PUT(i) = STR_LINE
        Next

        Dim STW_CSV_WRITER As System.IO.StreamWriter 'ファイル出力用のIOオブジェクト
        Try
            STW_CSV_WRITER = New System.IO.StreamWriter(STR_PATH_MAKE, False, System.Text.Encoding.UTF8)   'ファイルライターを開く
        Catch ex As Exception
            Return False
        End Try

        For i = 1 To (STR_PUT.Length - 1)
            STW_CSV_WRITER.WriteLine(STR_PUT(i)) 'CSVﾌｧｲﾙ書き込み
        Next

        Call STW_CSV_WRITER.Close() 'ファイルライターを閉じる

        Return True
    End Function

    Private Function FUNC_CONVERT_UTF_QUESTION(ByRef STR_VALUE As String) As String

        Dim STR_RET As String
        STR_RET = ""
        For i = 0 To STR_VALUE.Length - 1
            Dim STR_ONE As String

            STR_ONE = STR_VALUE.Substring(i, 1)
            If FUNC_CHECK_SHIFT_JIS(STR_ONE) Then
                STR_RET &= STR_ONE
            Else
                STR_RET &= "？"
            End If
        Next

        Return STR_RET
    End Function

    'EXCEL出力処理
    Private Function FUNC_PUT_EXCEL_BIP(
    ByVal strDIR_BIP_EXE As String,
    ByVal strDIR_ASSETS As String,
    ByVal strDEFINITION As String,
    ByVal strPATH_DATA As String,
    ByVal strPATH_PUT As String
    ) As Boolean

        'Dim STR_PATH_DECT_EXCEL As String
        'STR_PATH_DECT_EXCEL = strPATH_DATA & "_"
        'If Not FUNC_CONVRET_CSV(strPATH_DATA, STR_PATH_DECT_EXCEL) Then
        '    Return False
        'End If

        Dim STR_EXE_PATH As String
        Const CST_PRINT As String = "Bipprint.exe"
        STR_EXE_PATH = ""
        STR_EXE_PATH &= strDIR_BIP_EXE & "\"
        STR_EXE_PATH &= CST_PRINT

        Dim STR_COMMAND_LINE As String
        STR_COMMAND_LINE = ""
        STR_COMMAND_LINE &= """" & strDEFINITION & """"
        STR_COMMAND_LINE &= " " & "-assetsdir" & " " & strDIR_ASSETS
        STR_COMMAND_LINE &= " " & "-f" & " " & strPATH_DATA
        STR_COMMAND_LINE &= " " & "-atdirect" & " " & "excel"
        STR_COMMAND_LINE &= " " & "-keepxlsx" & " " & strPATH_PUT
        STR_COMMAND_LINE &= " " & "-indatacode" & " " & "SJIS"

        If Not FUNC_EXE_FILE_SHELL(STR_EXE_PATH, STR_COMMAND_LINE) = 0 Then
            Return False
        End If

        'Call System.Threading.Thread.Sleep(500)
        'If Not FUNC_FILE_CHECK(strPATH_PUT) Then
        '    Return False
        'End If

        Return True
    End Function

    'PDF出力処理
    Private Function FUNC_PUT_PDF_BIP(
    ByVal strDIR_BIP_EXE As String,
    ByVal strDIR_ASSETS As String,
    ByVal strDEFINITION As String,
    ByVal strPATH_DATA As String,
    ByVal strPATH_PUT As String
    ) As Boolean

        Dim STR_EXE_PATH As String
        Const CST_PRINT As String = "Bipprint.exe"
        STR_EXE_PATH = ""
        STR_EXE_PATH &= strDIR_BIP_EXE & "\"
        STR_EXE_PATH &= CST_PRINT

        Dim STR_COMMAND_LINE As String
        STR_COMMAND_LINE = ""
        STR_COMMAND_LINE &= """" & strDEFINITION & """"
        STR_COMMAND_LINE &= " " & "-assetsdir" & " " & strDIR_ASSETS
        STR_COMMAND_LINE &= " " & "-f" & " " & strPATH_DATA
        STR_COMMAND_LINE &= " " & "-atdirect" & " " & "file"
        STR_COMMAND_LINE &= " " & "-keeppdf" & " " & strPATH_PUT
        STR_COMMAND_LINE &= " " & "-gpdfsubtitle" & " " & "aaa"
        STR_COMMAND_LINE &= " " & "-gpdfauthor" & " " & "aaa"
        STR_COMMAND_LINE &= " " & "-gpdfprint" & " " & "Y"
        STR_COMMAND_LINE &= " " & "-gpdfmodify" & " " & "Y"
        STR_COMMAND_LINE &= " " & "-gpdfselect" & " " & "Y"
        STR_COMMAND_LINE &= " " & "-gpdfannotate" & " " & "Y"

        If Not FUNC_CALL_EXE_FILE_SHELL(STR_EXE_PATH, STR_COMMAND_LINE) Then
            Return False
        End If

        Return True
    End Function

    'BIPファイルのパース
    Private Function FUNC_GET_BIP_INFO(ByVal STR_DIR_ASSETS As String, ByVal STR_DEFINITION As String) As SRT_BIP_INFO
        Dim SRT_RET As SRT_BIP_INFO
        With SRT_RET.ENVIRONMENT
            .COMMENT = ""
        End With
        ReDim SRT_RET.DATA_FORMAT(0)
        With SRT_RET.DATA_FORMAT(0)
            .NAME_COL = ""
            .OPERAND = ""
        End With

        Dim STR_PATH_FILE As String
        STR_PATH_FILE = STR_DIR_ASSETS & "\" & STR_DEFINITION & ".bip"

        Dim STR_FILE_ROW() As String
        STR_FILE_ROW = Nothing
        If Not FUNC_GET_FILE_ROW(STR_PATH_FILE, STR_FILE_ROW) Then
            Return SRT_RET
        End If

        Dim BLN_ENVIRONMENT As Boolean
        Dim BLN_DATA_FORMAT As Boolean
        BLN_ENVIRONMENT = False
        BLN_DATA_FORMAT = False

        Dim STR_ENVIRONMENT() As String
        Dim STR_DATA_FORMAT() As String
        ReDim STR_ENVIRONMENT(0)
        ReDim STR_DATA_FORMAT(0)
        For i = 0 To (STR_FILE_ROW.Length - 1)
            If FUNC_CHECK_CATEGORY(STR_FILE_ROW(i)) Then
                BLN_ENVIRONMENT = False
                BLN_DATA_FORMAT = False

                Dim STR_CATEGORY As String
                STR_CATEGORY = STR_FILE_ROW(i)
                STR_CATEGORY = STR_CATEGORY.Replace("[", "")
                STR_CATEGORY = STR_CATEGORY.Replace("]", "")

                Select Case STR_CATEGORY
                    Case "ENVIRONMENT"
                        BLN_ENVIRONMENT = True
                    Case "DATA FORMAT"
                        BLN_DATA_FORMAT = True
                    Case Else
                End Select
            Else
                If BLN_ENVIRONMENT Then
                    Dim INT_INDEX As Integer
                    INT_INDEX = STR_ENVIRONMENT.Length
                    ReDim Preserve STR_ENVIRONMENT(INT_INDEX)
                    STR_ENVIRONMENT(INT_INDEX) = STR_FILE_ROW(i)
                End If

                If BLN_DATA_FORMAT Then
                    Dim INT_INDEX As Integer
                    INT_INDEX = STR_DATA_FORMAT.Length
                    ReDim Preserve STR_DATA_FORMAT(INT_INDEX)
                    STR_DATA_FORMAT(INT_INDEX) = STR_FILE_ROW(i)
                End If
            End If
        Next

        For i = 1 To (STR_ENVIRONMENT.Length - 1)
            Dim STR_SEP() As String
            STR_SEP = STR_ENVIRONMENT(i).Split(" ")
            Select Case STR_SEP(0)
                Case "COMMENT"
                    SRT_RET.ENVIRONMENT.COMMENT = STR_SEP(5).Replace("""", "")
                Case Else

            End Select
        Next

        For i = 1 To (STR_DATA_FORMAT.Length - 1)
            Dim STR_SEP() As String
            STR_SEP = STR_DATA_FORMAT(i).Split(" ")
            Dim INT_INDEX As Integer
            INT_INDEX = SRT_RET.DATA_FORMAT.Length
            ReDim Preserve SRT_RET.DATA_FORMAT(INT_INDEX)
            With SRT_RET.DATA_FORMAT(INT_INDEX)
                .NAME_COL = STR_SEP(0)
                .OPERAND = STR_SEP(2)
            End With
        Next

        Return SRT_RET
    End Function

    'ファイル全情報を取得
    Private Function FUNC_GET_FILE_ROW(ByVal STR_PATH_FILE As String, ByRef STR_FILE_ROW() As String) As Boolean
        If STR_FILE_ROW Is Nothing Then
            ReDim STR_FILE_ROW(0)
        End If

        Dim ENC_FILE As System.Text.Encoding
        ENC_FILE = System.Text.Encoding.GetEncoding("shift_jis")

        Dim STR_READER As System.IO.StreamReader
        Try
            STR_READER = New System.IO.StreamReader(STR_PATH_FILE, ENC_FILE)
        Catch ex As Exception
            Return False
        End Try
        Const CST_MAX_INDEX As Integer = 200000
        ReDim STR_FILE_ROW(CST_MAX_INDEX)
        Dim INT_INDEX As Integer
        INT_INDEX = 0
        While (STR_READER.Peek() > -1)
            INT_INDEX += 1
            Dim STR_ROW As String
            STR_ROW = STR_READER.ReadLine()
            STR_FILE_ROW(INT_INDEX) = STR_ROW
        End While

        ReDim Preserve STR_FILE_ROW(INT_INDEX)

        Call STR_READER.Close()
        Call STR_READER.Dispose()
        STR_READER = Nothing

        Return True
    End Function

    '文字列がカテゴリかどうか判断
    Private Function FUNC_CHECK_CATEGORY(ByVal STR_ROW As String) As Boolean

        If STR_ROW = "" Then
            Return False
        End If

        Dim STR_ONE As String
        STR_ONE = STR_ROW.Substring(0, 1)
        If STR_ONE = "[" Then
            Return True
        End If

        Return False
    End Function

    Private Const cstDEFAULT_PROCESS_NAME As String = "Bipview"
    Private Function FUNC_GET_BIP_PREVIEW_PROCESS(Optional ByVal strFFXIV_PROCESS_NAME As String = cstDEFAULT_PROCESS_NAME, Optional ByRef INT_PROCESS_ID_EXCEPT() As Integer = Nothing) As Process
        Dim p As Process
        For Each p In Process.GetProcesses()
            If (p.MainWindowHandle <> IntPtr.Zero) Then
                If (p.ProcessName = strFFXIV_PROCESS_NAME) Then
                    If FUNC_CHECK_INT_ROW(INT_PROCESS_ID_EXCEPT, p.Id) Then
                        Continue For
                    End If
                    Return p
                End If
            End If
        Next

        Return Nothing
    End Function

    Private Function FUNC_CHECK_INT_ROW(ByRef INT_CHECK() As Integer, ByVal INT_VALUE As Integer) As Boolean

        If INT_CHECK Is Nothing Then
            Return False
        End If

        For i = 0 To (INT_CHECK.Length - 1)
            If INT_VALUE = INT_CHECK(i) Then
                Return True
            End If
        Next

        Return False
    End Function
#End Region

End Module
