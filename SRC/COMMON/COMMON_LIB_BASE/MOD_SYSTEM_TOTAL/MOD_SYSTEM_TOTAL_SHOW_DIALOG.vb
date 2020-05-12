Public Module MOD_SYSTEM_TOTAL_SHOW_DIALOG

    Public Function FUNC_SHOW_DIALOG_FONT_SETTING(ByRef sngFONT_SIZE As Single) As Boolean
        Dim frmDIALOG As FRM_SYSTEM_TOTAL_FONT_SETTING
        Dim rstRESULT As System.Windows.Forms.DialogResult
        Dim objTEMP As Object
        Dim intTEMP As Integer

        frmDIALOG = New FRM_SYSTEM_TOTAL_FONT_SETTING
        frmDIALOG.cmbFONT_SIZE.SelectedIndex = 0
        frmDIALOG.MY_RESULT = System.Windows.Forms.DialogResult.Cancel
        Call frmDIALOG.ShowDialog()
        rstRESULT = frmDIALOG.MY_RESULT

        If rstRESULT = System.Windows.Forms.DialogResult.Cancel Then
            Call frmDIALOG.Dispose()
            Return False
        End If

        objTEMP = frmDIALOG.cmbFONT_SIZE.SelectedItem()
        If objTEMP Is Nothing Then
            objTEMP = ""
        End If
        intTEMP = FUNC_VALUE_CONVERT_NUMERIC_INT(objTEMP.ToString, CInt(sngFONT_SIZE))
        sngFONT_SIZE = CSng(intTEMP)
        Call SUB_SET_APP_CONFIG_COLOR(frmDIALOG.picCOLOR_01.BackColor, CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCF)
        Call SUB_SET_APP_CONFIG_COLOR(frmDIALOG.picCOLOR_02.BackColor, CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCG)
        Call SUB_SET_APP_CONFIG_COLOR(frmDIALOG.picCOLOR_03.BackColor, CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCP)
        Call SUB_SET_APP_CONFIG_COLOR(frmDIALOG.picCOLOR_04.BackColor, CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCO)


        Call frmDIALOG.Dispose()
        frmDIALOG = Nothing
        Return True
    End Function

    Private Sub SUB_SET_APP_CONFIG_COLOR(ByRef colCOLOR As System.Drawing.Color, ByVal strKEY As String)
        Dim strTEMP As String

        If FUNC_CHECK_COLOR_EMPTY(colCOLOR) Then
            Call FUNC_SYSTEM_TOTAL_DELETE_APP_CONFIG(strKEY)
        Else
            strTEMP = FUNC_GET_STR_FROM_COLOR(colCOLOR)
            Call FUNC_SYSTEM_TOTAL_WRITE_APP_CONFIG(strKEY, strTEMP)
        End If

    End Sub

    Private Function FUNC_CHECK_COLOR_EMPTY(ByRef colCOLOR As System.Drawing.Color) As Boolean
        Dim intA As Integer

        If colCOLOR = Drawing.Color.Empty Then
            Return True
        End If

        intA = colCOLOR.A

        If intA = 0 Then
            Return True
        End If

        Return False
    End Function
End Module
