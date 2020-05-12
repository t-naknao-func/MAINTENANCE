Public Class FRM_SYSTEM_TOTAL_FONT_SETTING

#Region "画面用・変数"
    Private blnPROCESS_DOING As Boolean
    Private picSETTING As System.Windows.Forms.PictureBox

    Private colDEFAULT_01 As System.Drawing.Color
    Private colDEFAULT_02 As System.Drawing.Color
    Private colDEFAULT_03 As System.Drawing.Color
    Private colDEFAULT_04 As System.Drawing.Color
#End Region

#Region "画面用・列挙定数"
    Private Enum ENM_MY_EXEC_DO
        DO_OK = 1
        DO_CANCEL = 2
        DO_SHOW_COLOR_DIALOG = 3
    End Enum
#End Region

#Region "プロパティ用変数"
    Private RST_PROPERTY_MY_RESULT As System.Windows.Forms.DialogResult
#End Region

#Region "プロパティ"
    Public Property MY_RESULT As System.Windows.Forms.DialogResult
        Get
            Return RST_PROPERTY_MY_RESULT
        End Get
        Set(ByVal value As System.Windows.Forms.DialogResult)
            RST_PROPERTY_MY_RESULT = value
        End Set
    End Property
#End Region

#Region "各処理呼出元"
    Private Sub SUB_EXEC_DO( _
    ByVal enmEXEC_DO As ENM_MY_EXEC_DO _
    )
        If blnPROCESS_DOING Then
            Exit Sub
        End If

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        blnPROCESS_DOING = True
        Call System.Windows.Forms.Application.DoEvents()

        Select Case enmEXEC_DO
            Case ENM_MY_EXEC_DO.DO_OK
                Call SUB_OK()
            Case ENM_MY_EXEC_DO.DO_CANCEL
                Call SUB_CANCEL()
            Case ENM_MY_EXEC_DO.DO_SHOW_COLOR_DIALOG
                Call SUB_SHOW_COLOR_DIALOG()
        End Select

        Call System.Windows.Forms.Application.DoEvents()
        blnPROCESS_DOING = False
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
#End Region

#Region "実行処理群"
    Private Sub SUB_OK()
        Me.MY_RESULT = Windows.Forms.DialogResult.OK
        Call Me.Close()
    End Sub

    Private Sub SUB_CANCEL()
        Me.MY_RESULT = Windows.Forms.DialogResult.Cancel
        Call Me.Close()
    End Sub

    Private Sub SUB_SHOW_COLOR_DIALOG()
        Dim rstMSG As System.Windows.Forms.DialogResult

        If picSETTING Is Nothing Then
            Exit Sub
        End If

        cldCOLOR_SETTING.Color = picSETTING.BackColor
        rstMSG = cldCOLOR_SETTING.ShowDialog
        If rstMSG = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        picSETTING.BackColor = cldCOLOR_SETTING.Color
    End Sub
#End Region

#Region "初期化・終了処理"
    Private Sub SUB_CTRL_NEW_INIT()

    End Sub

    Private Sub SUB_CTRL_DISPOSED_FIN() '画面破棄時の追記処理(Dispose時)

    End Sub

    Private Sub SUB_CTRL_VIEW_INIT()
        colDEFAULT_01 = System.Drawing.Color.FromArgb(230, 230, 240)
        colDEFAULT_02 = System.Drawing.Color.FromArgb(230, 230, 240)
        colDEFAULT_03 = System.Drawing.Color.FromArgb(225, 225, 235)
        colDEFAULT_04 = System.Drawing.Color.FromArgb(225, 225, 235)

        'colDEFAULT_01 = Drawing.Color.CornflowerBlue
        'colDEFAULT_02 = Drawing.Color.CornflowerBlue
        'colDEFAULT_03 = System.Drawing.Color.FromArgb(80, 140, 240)
        'colDEFAULT_04 = System.Drawing.Color.FromArgb(80, 140, 240)

        'colDEFAULT_01 = System.Drawing.Color.FromKnownColor(Drawing.KnownColor.ControlLightLight)
        'colDEFAULT_02 = System.Drawing.Color.FromKnownColor(Drawing.KnownColor.ControlLightLight)
        'colDEFAULT_03 = System.Drawing.Color.FromKnownColor(Drawing.KnownColor.ControlLight)
        '        colDEFAULT_04 = System.Drawing.Color.FromKnownColor(Drawing.KnownColor.Control)
    End Sub

    Private Sub SUB_CTRL_VALUE_INIT()
        Call SUB_CONTROL_CLEAR_FORM(Me)

        Dim strTEMP As String
        Dim colTEMP As System.Drawing.Color

        strTEMP = FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCF)
        colTEMP = FUNC_GET_COLOR_FROM_STR(strTEMP)
        If colTEMP = System.Drawing.Color.Empty Then
            colTEMP = colDEFAULT_01
        End If
        picCOLOR_01.BackColor = colTEMP

        strTEMP = FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCG)
        colTEMP = FUNC_GET_COLOR_FROM_STR(strTEMP)
        If colTEMP = System.Drawing.Color.Empty Then
            colTEMP = colDEFAULT_02
        End If
        picCOLOR_02.BackColor = colTEMP

        strTEMP = FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCP)
        colTEMP = FUNC_GET_COLOR_FROM_STR(strTEMP)
        If colTEMP = System.Drawing.Color.Empty Then
            colTEMP = colDEFAULT_03
        End If
        picCOLOR_03.BackColor = colTEMP

        strTEMP = FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCO)
        colTEMP = FUNC_GET_COLOR_FROM_STR(strTEMP)
        If colTEMP = System.Drawing.Color.Empty Then
            colTEMP = colDEFAULT_04
        End If
        picCOLOR_04.BackColor = colTEMP
    End Sub
#End Region

#Region "イベント-ボタンクリック"
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_OK)
    End Sub

    Private Sub btnCANCEL_Click(sender As Object, e As EventArgs) Handles btnCANCEL.Click
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_CANCEL)
    End Sub

    Private Sub btnCOLOR_01_INIT_Click(sender As Object, e As EventArgs) Handles btnCOLOR_01_INIT.Click
        picCOLOR_01.BackColor = colDEFAULT_01
        'picCOLOR_01.BackColor = System.Drawing.Color.FromArgb(0, picCOLOR_01.BackColor)
    End Sub

    Private Sub btnCOLOR_02_INIT_Click(sender As Object, e As EventArgs) Handles btnCOLOR_02_INIT.Click
        picCOLOR_02.BackColor = colDEFAULT_02
        'picCOLOR_02.BackColor = System.Drawing.Color.FromArgb(0, picCOLOR_02.BackColor)
    End Sub

    Private Sub btnCOLOR_03_INIT_Click(sender As Object, e As EventArgs) Handles btnCOLOR_03_INIT.Click
        picCOLOR_03.BackColor = colDEFAULT_03
        'picCOLOR_03.BackColor = System.Drawing.Color.FromArgb(0, picCOLOR_03.BackColor)
    End Sub

    Private Sub btnCOLOR_04_INIT_Click(sender As Object, e As EventArgs) Handles btnCOLOR_04_INIT.Click
        picCOLOR_04.BackColor = colDEFAULT_04
        ' picCOLOR_04.BackColor = System.Drawing.Color.FromArgb(0, picCOLOR_04.BackColor)
    End Sub
#End Region

#Region "イベント-ピクチャークリック"
    Private Sub picCOLOR_01_Click(sender As Object, e As EventArgs) Handles picCOLOR_01.Click
        picSETTING = sender
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_SHOW_COLOR_DIALOG)
        picSETTING = Nothing
    End Sub

    Private Sub picCOLOR_02_Click(sender As Object, e As EventArgs) Handles picCOLOR_02.Click
        picSETTING = sender
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_SHOW_COLOR_DIALOG)
        picSETTING = Nothing
    End Sub

    Private Sub picCOLOR_03_Click(sender As Object, e As EventArgs) Handles picCOLOR_03.Click
        picSETTING = sender
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_SHOW_COLOR_DIALOG)
        picSETTING = Nothing
    End Sub

    Private Sub picCOLOR_04_Click(sender As Object, e As EventArgs) Handles picCOLOR_04.Click
        picSETTING = sender
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_SHOW_COLOR_DIALOG)
        picSETTING = Nothing
    End Sub

#End Region

    Private Sub FRM_SYSTEM_TOTAL_FONT_SETTING_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call SUB_CTRL_VIEW_INIT()
        Call SUB_CTRL_VALUE_INIT()
    End Sub

    Private Sub FRM_SYSTEM_TOTAL_FONT_SETTING_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub

    Private Sub FRM_SYSTEM_TOTAL_FONT_SETTING_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Enabled = False
        Call System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub FRM_SYSTEM_TOTAL_FONT_SETTING_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    End Sub

End Class