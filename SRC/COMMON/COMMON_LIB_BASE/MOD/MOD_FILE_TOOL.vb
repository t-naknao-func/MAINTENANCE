Public Module MOD_FILE_TOOL

#Region "外出変数"
    Public str_FILE_TOOL_LAST_ERR_STRING As String
#End Region

    'パスから絶対パスを取得する
    Public Function FUNC_GET_ABB_PATH(ByVal strFILE_PATH As String) As String
        Dim strRET As String

        Try
            strRET = System.IO.Path.GetFullPath(strFILE_PATH)
        Catch ex As Exception
            Return ""
        End Try

        Return strRET
    End Function

    'パスからファイル名を抜き出す
    Public Function FUNC_PATH_TO_FILENAME( _
    ByVal strHFilePath As String _
    ) As String
        Dim intWLoopIndex As Integer
        Dim strWRetFileName As String

        strWRetFileName = ""
        intWLoopIndex = Len(strHFilePath)

        Do Until intWLoopIndex = 0
            If Mid(strHFilePath, intWLoopIndex, 1) = "\" Then
                strWRetFileName = Mid(strHFilePath, intWLoopIndex + 1, Len(strHFilePath) - intWLoopIndex)
                Exit Do
            End If
            intWLoopIndex = intWLoopIndex - 1
        Loop

        Return strWRetFileName
    End Function

    'パスからディレクトリパスを抜き出す
    Public Function FUNC_PATH_TO_DIR_PATH( _
    ByVal strPATH As String _
    ) As String
        Dim strRET As String
        Dim intLEN As Integer
        Dim strCHECK As String
        Dim intLOOP_INDEX As Integer

        intLEN = Len(strPATH)
        strRET = ""
        For intLOOP_INDEX = 1 To intLEN
            strCHECK = Mid(strPATH, intLEN - intLOOP_INDEX, 1)
            If strCHECK = "\" Then
                strRET = Mid(strPATH, 1, intLEN - intLOOP_INDEX)
                Exit For
            End If
        Next

        Return strRET
    End Function

    'ファイル名から拡張子を省く(半/全角混在,複数ピリオド対応)
    Public Function FUNC_GET_FILENAME_REMOVE_EXCTENT( _
    ByVal strHFileStr As String _
    ) As String
        Dim intWPoint As Integer
        Dim strWTemp As String
        Dim intWLen As Integer
        Dim intWIdx As Integer
        Dim strRET As String

        intWLen = Len(strHFileStr)
        intWPoint = 0

        For intWIdx = 1 To intWLen - 1
            strWTemp = Mid(strHFileStr, intWLen - (intWIdx - 1), 1)
            If strWTemp = "." Then
                intWPoint = intWIdx
                Exit For
            End If
        Next intWIdx

        strRET = Mid(strHFileStr, 1, intWLen - intWPoint)

        Return strRET
    End Function

    'EXE実行処理(呼出用)
    Public Function FUNC_CALL_EXE_FILE_SHELL( _
    ByVal strEXE_PATH As String, ByVal strCOMMAND_LINE As String _
    ) As Boolean
        Dim intACCESSID As Integer

        str_FILE_TOOL_LAST_ERR_STRING = ""

        If Not FUNC_FILE_CHECK(strEXE_PATH) Then 'チェックを行う
            str_FILE_TOOL_LAST_ERR_STRING = strEXE_PATH & Environment.NewLine & "ファイルがありません"
            Return False
        End If

        intACCESSID = FUNC_EXE_FILE_SHELL(strEXE_PATH, strCOMMAND_LINE) '実呼出
        If intACCESSID = -1 Then
            Return False
        End If

        Call System.Windows.Forms.Application.DoEvents()

        Call FUNC_APP_ACTIVE(intACCESSID) '画面をアクティブにする

        Return True
    End Function

    'ファイルチェック処理
    Public Function FUNC_FILE_CHECK( _
    ByVal strFILE_PATH As String _
    ) As Boolean
        Dim filBASE As System.IO.FileInfo
        Dim blnRET As Boolean

        If strFILE_PATH Is Nothing Then
            Return False
        End If

        If strFILE_PATH = "" Then
            Return False
        End If

        filBASE = New System.IO.FileInfo(strFILE_PATH)
        blnRET = filBASE.Exists

        Call GC.ReRegisterForFinalize(filBASE)
        filBASE = Nothing

        Return blnRET
    End Function

    'ファイルの拡張子を取得
    Public Function FUNC_GET_FILE_EXTENT( _
    ByVal strFILE_PATH As String _
    ) As String
        Dim filBASE As System.IO.FileInfo
        Dim strRET As String

        If strFILE_PATH Is Nothing Then
            Return ""
        End If

        If strFILE_PATH = "" Then
            Return ""
        End If

        filBASE = New System.IO.FileInfo(strFILE_PATH)
        strRET = filBASE.Extension

        Call GC.ReRegisterForFinalize(filBASE)
        filBASE = Nothing

        Return strRET
    End Function

    'ファイルパスでチェックし、同一のファイル名がある場合は連番付与の判断を行い、その連番を返す
    '返却連番=0はオリジナルファイルが存在しない、連番は1～99
    Public Function FUNC_FILE_CHECK_NUMBER( _
    ByVal strFILE_PATH As String _
    ) As Boolean
        Dim strTEMP As String
        Dim strEXTENT As String
        Dim intRET As Integer
        Dim intLOOP_INDEX As Integer
        Const cstMAX_INDEX As Integer = 99

        intRET = 0
        If Not FUNC_FILE_CHECK(strFILE_PATH) Then
            Return 0
        End If

        strEXTENT = FUNC_GET_FILE_EXTENT(strFILE_PATH)
        For intLOOP_INDEX = 1 To cstMAX_INDEX
            strTEMP = FUNC_GET_FILENAME_REMOVE_EXCTENT(strFILE_PATH) & CStr(intLOOP_INDEX) & strEXTENT
            If Not FUNC_FILE_CHECK(strTEMP) Then
                Exit For
            End If
        Next

        intRET = intLOOP_INDEX

        Return intRET
    End Function

    'EXE実行処理
    Public Function FUNC_EXE_FILE_SHELL(
    ByVal strEXE_PATH As String, ByVal strCOMMAND_LINE As String
    ) As Integer
        Dim strPATH As String
        Dim intRET As Integer
        Dim strEXE_PATH_ABB As String

        Dim psiSET As System.Diagnostics.ProcessStartInfo
        str_FILE_TOOL_LAST_ERR_STRING = ""

        psiSET = New System.Diagnostics.ProcessStartInfo

        strEXE_PATH_ABB = FUNC_GET_ABB_PATH(strEXE_PATH)

        With psiSET
            .FileName = strEXE_PATH_ABB
            .Arguments = strCOMMAND_LINE
            .WorkingDirectory = FUNC_PATH_TO_DIR_PATH(strEXE_PATH_ABB)
        End With

        strPATH = strEXE_PATH & " " & strCOMMAND_LINE
        Try
            ' パラメータを指定して実行
            'Process.Start(strEXE_PATH, strCOMMAND_LINE)
            intRET = 0
            Process.Start(psiSET)
        Catch ex As Exception
            intRET = -1
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
        End Try

        Return intRET
    End Function

    '他のEXEの画面をアクティブにする(繰返あり)
    Public Function FUNC_APP_ACTIVE( _
    ByVal intACCESS_ID As Integer, _
    Optional ByVal intNUMBER_OF_TIMES As Integer = 1 _
    ) As Boolean
        Dim blnRET As Boolean
        Dim intLOOP_INDEX As Integer

        blnRET = False
        For intLOOP_INDEX = 1 To intNUMBER_OF_TIMES
            If FUNC_APP_ACTIVE_MAIN(intACCESS_ID) Then
                blnRET = True
                Exit For
            End If
            Call System.Windows.Forms.Application.DoEvents()
        Next

        Return blnRET
    End Function

    '他のEXEの画面をアクティブにする
    Private Function FUNC_APP_ACTIVE_MAIN( _
    ByVal intACCESS_ID As Integer _
    ) As Boolean
        Try
            'Call Microsoft.VisualBasic.Interaction.AppActivate(intACCESS_ID)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    'ディレクトリのチェック
    Public Function FUNC_DIR_CHECK(ByVal strPATH_DIR As String) As Boolean
        Dim dirCHECK As System.IO.DirectoryInfo
        Dim blnRET As Boolean

        str_FILE_TOOL_LAST_ERR_STRING = ""

        Try
            dirCHECK = New System.IO.DirectoryInfo(strPATH_DIR)
        Catch ex As Exception
            dirCHECK = Nothing
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
            Return True '無理やりあるものとして返して上位関数でのエラーを誘う
        End Try

        blnRET = dirCHECK.Exists
        dirCHECK = Nothing

        Return blnRET
    End Function

    'ディレクトリの作成
    Public Function FUNC_DIR_MAKE(ByVal strPATH_DIR As String) As Boolean
        Dim dirMake As System.IO.DirectoryInfo
        Dim strPATH_DIR_TOP As String

        str_FILE_TOOL_LAST_ERR_STRING = ""

        strPATH_DIR_TOP = FUNC_DIR_CONVERT_ONE_TOP(strPATH_DIR)

        If Not FUNC_DIR_CHECK(strPATH_DIR_TOP) Then
            If Not FUNC_DIR_MAKE(strPATH_DIR_TOP) Then
                Return False
            End If
        End If

        If FUNC_DIR_CHECK(strPATH_DIR) Then 'すでに存在するなら
            Return True 'そのまま
        End If

        Try
            dirMake = New System.IO.DirectoryInfo(strPATH_DIR)
        Catch ex As Exception
            dirMake = Nothing
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
            Return False
        End Try

        Try
            Call dirMake.Create() '作成
            dirMake = Nothing
        Catch ex As Exception
            dirMake = Nothing
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
            Return False
        End Try

        Return True
    End Function

    '上位ディレクトリへ変換
    Public Function FUNC_DIR_CONVERT_ONE_TOP(ByVal strPATH_DIR As String) As String
        Dim strRET As String
        Dim intLEN As Integer
        Dim strCHECK As String
        Dim intLOOP_INDEX As Integer

        intLEN = strPATH_DIR.Length
        strRET = ""
        For intLOOP_INDEX = 1 To intLEN
            strCHECK = Mid(strPATH_DIR, intLEN - intLOOP_INDEX, 1)
            If strCHECK = "\" Then
                strRET = Mid(strPATH_DIR, 1, intLEN - intLOOP_INDEX)
                Exit For
            End If
        Next

        Return strRET
    End Function

    'ディレクトリの再作成
    Public Function FUNC_DIR_REMAKE(ByVal strPATH_DIR As String) As Boolean
        Dim dirReMake As System.IO.DirectoryInfo

        str_FILE_TOOL_LAST_ERR_STRING = ""

        Try
            dirReMake = New System.IO.DirectoryInfo(strPATH_DIR)
            If dirReMake.Exists Then 'すでに存在する場合は
                dirReMake.Delete(True) '中身ごと削除
            End If

            Call dirReMake.Create() '再作成
            dirReMake = Nothing
        Catch ex As Exception
            dirReMake = Nothing
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
            Return False
        End Try

        Return True

    End Function

    'ディレクトリの強制削除
    Public Function FUNC_DIR_DELETE_FORCE(ByVal strPATH_DIR As String) As Boolean
        Dim dirReMove As System.IO.DirectoryInfo

        str_FILE_TOOL_LAST_ERR_STRING = ""

        Try
            dirReMove = New System.IO.DirectoryInfo(strPATH_DIR)
            If dirReMove.Exists Then 'すでに存在する場合は
                dirReMove.Delete(True) '中身ごと削除
            End If
            dirReMove = Nothing
        Catch ex As Exception
            dirReMove = Nothing
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
            Return False
        End Try

        Return True
    End Function

    'ファイルに追記書き出しを行う
    Public Function FUNC_FILE_APPEND_WRITE(ByVal strFilePath As String, ByVal strOneRow As String) As Boolean
        Dim stwCsvWriter As System.IO.StreamWriter

        str_FILE_TOOL_LAST_ERR_STRING = ""

        Try
            stwCsvWriter = New System.IO.StreamWriter(strFilePath, True, System.Text.Encoding.Default)
            stwCsvWriter.WriteLine(strOneRow)
            stwCsvWriter.Close()
        Catch ex As Exception
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
            Return False
        End Try

        Return True
    End Function

    'ファイルの削除を行う
    Public Function FUNC_FILE_DELETE(ByVal strFilePath As String) As Boolean

        If Not FUNC_FILE_CHECK(strFilePath) Then
            Return True 'ファイルがなければ消去する必要なし
        End If

        Try
            Call System.IO.File.Delete(strFilePath)
        Catch ex As Exception
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
            Return False
        End Try

        Return True
    End Function

    'ファイルのコピーを行う
    Public Function FUNC_FILE_COPY( _
    ByVal strFILE_SOURCE As String, ByVal strFILE_DEST As String _
    ) As Boolean
        If Not FUNC_FILE_CHECK(strFILE_SOURCE) Then
            str_FILE_TOOL_LAST_ERR_STRING = "基本ファイル：" & strFILE_SOURCE & "が存在しません"
            Return False 'ファイルがなければコピーできない
        End If

        If FUNC_FILE_CHECK(strFILE_DEST) Then 'コピー先ファイルが存在する場合は
            If Not FUNC_FILE_DELETE(strFILE_DEST) Then '強制削除
                'エラーは中で
                Return False
            End If
        End If

        Try
            Call System.IO.File.Copy(strFILE_SOURCE, strFILE_DEST)
        Catch ex As Exception
            str_FILE_TOOL_LAST_ERR_STRING = ex.Message
            Return False
        End Try

        Return True
    End Function

    '指定のディレクトリのファイルを取得する(代表一つ)
    Public Function FUNC_FILE_NAME_GETS(ByVal strDIR As String, ByVal strFILETER As String) As String
        Dim strFILES() As String
        Dim strRET As String

        strFILES = System.IO.Directory.GetFiles(strDIR, "*" & strFILETER & "*")

        If strFILES Is Nothing Then
            Return ""
        End If

        If strFILES.Length < 1 Then
            Return ""
        End If

        strRET = strFILES(0)

        Return strRET
    End Function

    '指定のディレクトリのファイルを取得する(全て)
    Public Function FUNC_FILE_NAME_GETS_ALL(ByVal strDIR As String, ByVal strFILETER As String) As String()
        Dim strFILES() As String
        Dim strRET() As String
        Dim intLOOP_INDEX As Integer
        Dim intINDEX As Integer

        ReDim strRET(0)

        strFILES = System.IO.Directory.GetFiles(strDIR, "*" & strFILETER & "*")

        If strFILES Is Nothing Then
            Return strRET
        End If

        If strFILES.Length < 1 Then
            Return strRET
        End If

        For intLOOP_INDEX = 0 To (strFILES.Length - 1)
            intINDEX = strRET.Length
            ReDim Preserve strRET(intINDEX)
            strRET(intINDEX) = strFILES(intLOOP_INDEX)
        Next

        Return strRET
    End Function
End Module
