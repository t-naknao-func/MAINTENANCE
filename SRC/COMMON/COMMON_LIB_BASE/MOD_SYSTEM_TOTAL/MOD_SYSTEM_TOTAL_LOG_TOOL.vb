Public Module MOD_SYSTEM_TOTAL_LOG_TOOL

    'デバッグログの出力判断・及び出力
    Public Sub SUB_SYSTEM_LOG_PUT_DEBUG( _
    ByVal strLOG As String _
    )
        Dim cstEXTEND_LOG As String = ".LOG"
        Dim strLOG_DIR As String
        Dim strLOG_FILE_NAME As String
        Dim strLOG_PATH As String
        Dim blnPUT As Boolean
        Dim strHEAD As String
        Dim strEXE_NAME As String

        blnPUT = True
        If Not blnPUT Then
            Exit Sub
        End If

        strLOG_DIR = FUNC_PATH_TO_DIR_PATH(System.Windows.Forms.Application.ExecutablePath) & "\"
        If Not FUNC_DIR_MAKE(strLOG_DIR) Then '出力用ディレクトリを作成
            Exit Sub
        End If

        strEXE_NAME = FUNC_GET_FILENAME_REMOVE_EXCTENT(FUNC_PATH_TO_FILENAME(System.Windows.Forms.Application.ExecutablePath))
        strLOG_FILE_NAME = strEXE_NAME & cstEXTEND_LOG

        strLOG_PATH = strLOG_DIR & strLOG_FILE_NAME

        strHEAD = System.DateTime.Today.ToShortDateString & " " & System.DateTime.Today.ToShortTimeString & " " & strEXE_NAME
        If Not FUNC_FILE_APPEND_WRITE(strLOG_PATH, strHEAD) Then
            Exit Sub
        End If

        If Not FUNC_FILE_APPEND_WRITE(strLOG_PATH, strLOG) Then
            Exit Sub
        End If

    End Sub

    'バッチ進捗ログの出力
    Public Sub SUB_SYSTEM_LOG_PUT_BATCH( _
    ByVal strLOG As String _
    )
        Dim cstEXTEND_LOG As String = ".LOG"
        'Dim strTEMP As String
        'Dim strREG_KEY As String
        'Dim strREG_VALUE As String
        'Dim strLOG_DIR As String
        'Dim strLOG_FILE_NAME As String
        'Dim strLOG_PATH As String
        'Dim blnPUT As Boolean
        'Dim strHEAD As String
        'Dim strEXE_NAME As String

    End Sub

    'バッチ時刻計測ログの出力
    Public Function FUNC_SYSTEM_PUT_TIME_LOG( _
    ByVal strPROCESS_NAME As String, _
    ByVal dblSEC As Double _
    ) As Boolean
        Dim cstEXTEND_LOG As String = ".LOG"
        'Dim strHEAD As String
        'Dim strDETAIL As String
        'Dim strLOG_PATH As String
        'Dim strEXE_NAME As String
        'Dim strLOG_DIR As String
        'Dim strLOG_FILE_NAME As String
        'Dim strTEMP As String

        '処理時間計測用の為、本番環境では本処理は不要
        Return True

        'strTEMP = FUNC_PATH_TO_DIR_PATH(Application.ExecutablePath)
        'strLOG_DIR = strTEMP & "\" & "BATCH_TIME" & "\"

        'If Not FUNC_DIR_MAKE(strLOG_DIR) Then '出力用ディレクトリを作成
        '    Return False
        'End If

        'strEXE_NAME = FUNC_GET_FILENAME_REMOVE_EXCTENT(FUNC_PATH_TO_FILENAME(Application.ExecutablePath))
        'strLOG_FILE_NAME = strEXE_NAME & cstEXTEND_LOG

        'strLOG_PATH = strLOG_DIR & strLOG_FILE_NAME

        'strHEAD = Today.ToShortDateString & vbTab & Now.ToLongTimeString

        'If dblSEC = 0 Then
        '    strDETAIL = strPROCESS_NAME
        'Else
        '    strDETAIL = strPROCESS_NAME & "：" & dblSEC & "秒"
        'End If

        'If Not FUNC_FILE_APPEND_WRITE(strLOG_PATH, strHEAD & vbTab & strDETAIL) Then
        '    Return False
        'End If

        'Return True
    End Function
End Module
