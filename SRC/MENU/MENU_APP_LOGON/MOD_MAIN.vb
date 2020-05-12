Module MOD_MAIN
    Private Const cstAPPL_NAME As String = "ログオン"
    Friend blnSHOW_SETTING As Boolean

    Public Sub MAIN(ByVal strCOMMAND_LINE() As String)
        Dim frmMAIN As Form
        Dim sngFONT_SIZE As Single

        If Not FUNC_SYSTEM_TOTAL_INITIALIZATION(strCOMMAND_LINE) Then
            Call MessageBox.Show(str_FUNC_SYSTEM_TOTAL_INITIALIZATION_MSG, cstAPPL_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        blnSHOW_SETTING = False
        sngFONT_SIZE = 0.0
        Do
            If blnSHOW_SETTING Then
                Call FUNC_SHOW_DIALOG_FONT_SETTING(sngFONT_SIZE)
                blnSHOW_SETTING = False
            End If

            frmMAIN = New FRM_MAIN
            Call SUB_SYSTEM_TOTAL_INIT_FONT(sngFONT_SIZE, frmMAIN.Font.Size)
            Dim srtPARAM As SRT_FORM_SETTING_PARAM
            With srtPARAM
                .TEXT = cstAPPL_NAME
                .FONT_SIZE = sngFONT_SIZE
                Erase .BACK_COLOR
                Call SUB_SET_FORM_SETTING_COLOR(.BACK_COLOR)
            End With
            Call SUB_INIT_FORM_SETTING(frmMAIN, srtPARAM)
            Call frmMAIN.ShowDialog()
            Call frmMAIN.Dispose()
            If Not blnSHOW_SETTING Then
                Exit Do
            End If
        Loop

        Call FUNC_SYSTEM_TOTAL_FINALIZATION()
    End Sub

    Private Sub SUB_SET_FORM_SETTING_COLOR(ByRef srtBACK_COLOR() As System.Drawing.Color)
        Dim strTEMP As String

        ReDim srtBACK_COLOR(ENM_COLOR_SET_LEVEL.UBOUND)
        strTEMP = FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCF)
        srtBACK_COLOR(ENM_COLOR_SET_LEVEL.FORM) = FUNC_GET_COLOR_FROM_STR(strTEMP)

        strTEMP = FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCG)
        srtBACK_COLOR(ENM_COLOR_SET_LEVEL.GROUP) = FUNC_GET_COLOR_FROM_STR(strTEMP)

        strTEMP = FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCP)
        srtBACK_COLOR(ENM_COLOR_SET_LEVEL.PANEL) = FUNC_GET_COLOR_FROM_STR(strTEMP)

        strTEMP = FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCO)
        srtBACK_COLOR(ENM_COLOR_SET_LEVEL.OBJ) = FUNC_GET_COLOR_FROM_STR(strTEMP)

    End Sub
End Module
