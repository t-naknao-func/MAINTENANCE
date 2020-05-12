Public Module MOD_FILE_OPEN_TOOL

    'ファイルセーブダイアログを開く
    Public Function FUNC_VIEW_SAVE_FILE_DIALOG(ByVal strDEFAULT_EXT As String, ByVal strFILTER As String, Optional ByVal strFILE_NAME As String = "") As String
        Dim ofdFILE As System.Windows.Forms.SaveFileDialog
        Dim rstMSG As System.Windows.Forms.DialogResult
        Dim strRET As String

        ofdFILE = New System.Windows.Forms.SaveFileDialog()

        If strFILE_NAME <> "" Then
            ofdFILE.FileName = strFILE_NAME
        End If
        ofdFILE.DefaultExt = strDEFAULT_EXT
        ofdFILE.Filter = strFILTER

        rstMSG = ofdFILE.ShowDialog
        If rstMSG = Windows.Forms.DialogResult.Cancel Then
            Return ""
        End If

        strRET = ofdFILE.FileName
        Call ofdFILE.Dispose()
        ofdFILE = Nothing

        Return strRET
    End Function

    'ファイルオープンダイアログを開く
    Public Function FUNC_VIEW_OPEN_FILE_DIALOG(ByVal strDEFAULT_EXT As String, ByVal strFILTER As String) As String
        Dim ofdFILE As System.Windows.Forms.OpenFileDialog
        Dim rstMSG As System.Windows.Forms.DialogResult
        Dim strRET As String

        ofdFILE = New System.Windows.Forms.OpenFileDialog()

        ofdFILE.DefaultExt = strDEFAULT_EXT
        ofdFILE.Filter = strFILTER

        rstMSG = ofdFILE.ShowDialog
        If rstMSG = Windows.Forms.DialogResult.Cancel Then
            Return ""
        End If

        strRET = ofdFILE.FileName
        Call ofdFILE.Dispose()
        ofdFILE = Nothing

        Return strRET
    End Function

    '単一のデフォルト拡張子からフィルターを返却
    Public Function FUNC_GET_FILETER_DEFAULT_EXT(ByVal strDEFAULT_EXT As String) As String
        Dim strRET As String

        strRET = "|*" & strDEFAULT_EXT

        Return strRET
    End Function
End Module
