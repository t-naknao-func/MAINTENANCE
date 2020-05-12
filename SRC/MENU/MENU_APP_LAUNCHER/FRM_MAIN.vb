Public Class FRM_MAIN

#Region "画面用・変数"
    Private blnPROCESS_DOING As Boolean
    Private srtMY_TAB_GROUP() As SRT_MY_TAB
    Private srtMY_LAUNCHER_GROUP() As SRT_MY_LAUNCHER
    Private tbpLAUNCHER_TABPAGE() As TabPage
    Private btnLAUNCHER_DETAIL() As Button
    Private lblLAUNCHER_DETAIL() As Label
    Private blnFIRST_ACTIVE As Boolean
#End Region

#Region "画面用・列挙定数"
    Private Enum ENM_MY_EXEC_DO
        DO_LOG_OFF = 1
        DO_BACK
        DO_SHELL

        DO_END = 81
        DO_SHOW_SETTING
        DO_SHOW_COMMANDLINE
        DO_SHOW_CONFIG_SETTINGS
    End Enum

    Private Enum ENM_MY_KIND_TABPAGE
        KIND_NONE = 0
        KIND_TABPAGE = 1
    End Enum

    Private Enum ENM_MY_KIND_LAUNCHER
        KIND_NONE = 0
        KIND_LABEL = 1
        KIND_BUTTON = 2
    End Enum
#End Region

#Region "画面用・構造体"
    Private Structure SRT_MY_TAB
        Public KIND_TABPAGE As ENM_MY_KIND_TABPAGE
        Public TEXT As String
    End Structure

    Private Structure SRT_MY_LAUNCHER
        Public KIND_LAUNCHER As ENM_MY_KIND_LAUNCHER
        Public TEXT As String
        Public COMMAND As String
        Public LEVEL As Integer
    End Structure
#End Region

#Region "初期化・終了処理"
    Private Sub SUB_CTRL_NEW_INIT()
        
    End Sub

    Private Sub SUB_CTRL_VIEW_INIT()
        Call SUB_GET_TAB_CONFIG()
        Call SUB_SET_WINDOW_TAB()
        Call SUB_GET_LAUNCHER_CONFIG()
        Call SUB_SET_WINDOW_LAUNCHER()
        Dim intLEVEL_USER As Integer 'ユーザーのレベル(小さい程高い)
        intLEVEL_USER = FUNC_GET_MNG_M_USER_FLAG_GRANT(srtSYSTEM_TOTAL_COMMANDLINE.CODE_STAFF)
        Call SUB_BUTTONS_ENABLED_REFRESH(intLEVEL_USER)

        blnFIRST_ACTIVE = True
    End Sub

    Private Sub SUB_CTRL_VALUE_INIT()
        lblNAME_USER.Text = FUNC_GET_MNG_M_USER_NAME_STAFF(srtSYSTEM_TOTAL_COMMANDLINE.CODE_STAFF)

        lblDATE_ACTIVE.Text = Format(datSYSTEM_TOTAL_DATE_ACTIVE, "yyyy年MM月dd日")
    End Sub

    Private Sub SUB_FIRST_ACTIVE()
        Dim intID_FOCUS As Integer
        Dim intLOOP_INDEX As Integer
        Dim strTAG As String
        Dim intTAG As Integer
        Dim intTAB As Integer

        If Not blnFIRST_ACTIVE Then
            Exit Sub
        End If
        blnFIRST_ACTIVE = False
        Call Application.DoEvents()

        intID_FOCUS = If(blnTOTAL_MODE, srtSYSTEM_TOTAL_COMMANDLINE.ID_FOCUS_MENU_TOTAL, srtSYSTEM_TOTAL_COMMANDLINE.ID_FOCUS_MENU)
        For intLOOP_INDEX = 1 To (btnLAUNCHER_DETAIL.Length - 1)
            strTAG = btnLAUNCHER_DETAIL(intLOOP_INDEX).Tag
            intTAG = CInt(strTAG)
            If intTAG = intID_FOCUS Then
                Call btnLAUNCHER_DETAIL(intLOOP_INDEX).Focus()
                intTAB = intTAG \ 100
                tabLAUNCHER_GROUP.SelectedIndex = (intTAB - 1)
            End If
        Next

        Call Application.DoEvents()
    End Sub
#End Region

#Region "各処理呼出元"
    Private Sub SUB_EXEC_DO( _
    ByVal enmEXEC_DO As ENM_MY_EXEC_DO, Optional ByVal intINDEX As Integer = 0 _
    )
        If blnPROCESS_DOING Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        blnPROCESS_DOING = True
        Call Application.DoEvents()

        Select Case enmEXEC_DO
            Case ENM_MY_EXEC_DO.DO_LOG_OFF
                Call SUB_LOG_OFF()
            Case ENM_MY_EXEC_DO.DO_BACK
                Call SUB_BACK()
            Case ENM_MY_EXEC_DO.DO_SHELL
                Call SUB_SHELL(intINDEX)

            Case ENM_MY_EXEC_DO.DO_END
                Call SUB_END()
            Case ENM_MY_EXEC_DO.DO_SHOW_SETTING
                Call SUB_SHOW_SETTING()
            Case ENM_MY_EXEC_DO.DO_SHOW_COMMANDLINE
                Call SUB_SHOW_COMMANDLINE()
            Case ENM_MY_EXEC_DO.DO_SHOW_CONFIG_SETTINGS
                Call SUB_SHOW_CONFIG_SETTINGS()
        End Select

        Call Application.DoEvents()
        blnPROCESS_DOING = False
        Me.Cursor = Cursors.Default
    End Sub
#End Region

#Region "実行処理群"
    Private Sub SUB_LOG_OFF()
        Dim strSHELL As String
        
        strSHELL = FUNC_GET_APP_SETTINGS("logoff")

        If Not FUNC_CALL_EXE_FILE_SHELL(strSHELL, "") Then
            Call MessageBox.Show(str_FILE_TOOL_LAST_ERR_STRING, Me.Text)
            Exit Sub
        End If
        Call Application.DoEvents()
        Call Me.Close()
    End Sub

    Private Sub SUB_BACK()
        Dim strSHELL As String
        Dim strCOMMAND_LINE As String

        strSHELL = FUNC_GET_APP_SETTINGS("back")
        strCOMMAND_LINE = FUNC_SYSTEM_TOTAL_MAKE_COMMANDLINE(Application.ExecutablePath, srtSYSTEM_TOTAL_COMMANDLINE.CODE_STAFF, srtSYSTEM_TOTAL_COMMANDLINE.ID_FOCUS_MENU_TOTAL, 0)

        If Not FUNC_CALL_EXE_FILE_SHELL(strSHELL, strCOMMAND_LINE) Then
            Call MessageBox.Show(str_FILE_TOOL_LAST_ERR_STRING, Me.Text)
            Exit Sub
        End If
        Call Application.DoEvents()
        Call Me.Close()
    End Sub

    Private Sub SUB_END()
        Call Me.Close()
    End Sub

    Private Sub SUB_SHOW_SETTING()
        blnSHOW_SETTING = True
        Call Me.Close()
    End Sub

    Private Sub SUB_SHOW_COMMANDLINE()
        Call SUB_SYSTEM_TOTAL_SHOW_COMMANDLINE(srtSYSTEM_TOTAL_COMMANDLINE)
    End Sub

    Private Sub SUB_SHOW_CONFIG_SETTINGS()
        Call SUB_SYSTEM_TOTAL_SHOW_SETTINGS(srtSYSTEM_TOTAL_CONFIG_SETTINGS)
    End Sub

    Private Sub SUB_SHELL(ByVal intINDEX As Integer)
        Dim strSHELL As String
        Dim strCOMMAND_LINE As String
        Dim intID_FOCUS_MENU_TOTAL As Integer
        Dim intID_FOCUS_MENU As Integer

        If blnTOTAL_MODE Then
            intID_FOCUS_MENU_TOTAL = intINDEX
            intID_FOCUS_MENU = 0
        Else
            intID_FOCUS_MENU_TOTAL = srtSYSTEM_TOTAL_COMMANDLINE.ID_FOCUS_MENU_TOTAL
            intID_FOCUS_MENU = intINDEX
        End If

        strSHELL = srtMY_LAUNCHER_GROUP(intINDEX).COMMAND
        strCOMMAND_LINE = FUNC_SYSTEM_TOTAL_MAKE_COMMANDLINE(Application.ExecutablePath, srtSYSTEM_TOTAL_COMMANDLINE.CODE_STAFF, intID_FOCUS_MENU_TOTAL, intID_FOCUS_MENU)

        If Not FUNC_CALL_EXE_FILE_SHELL(strSHELL, strCOMMAND_LINE) Then
            Call MessageBox.Show(str_FILE_TOOL_LAST_ERR_STRING, Me.Text)
            Exit Sub
        End If
        Call Application.DoEvents()
        Call Me.Close()
    End Sub
#End Region

#Region "キー制御処理"
    '通常のコマンドキー制御(シフトマスク無し)
    Private Sub SUB_KEY_DOWN(ByVal enmKEY_CODE As Windows.Forms.Keys, ByRef blnHandled As Boolean)
        
    End Sub

    '通常のコマンドキー制御(ALT)
    Private Sub SUB_KEY_DOWN_ALT(ByVal enmKEY_CODE As Windows.Forms.Keys, ByRef blnHandled As Boolean)
        Select Case enmKEY_CODE
            Case Keys.C
                Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_SHOW_COMMANDLINE)
            Case Keys.V
                Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_SHOW_CONFIG_SETTINGS)
            Case Keys.S
                Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_SHOW_SETTING)
            Case Else
                'スルー
        End Select
    End Sub
#End Region

#Region "その他処理"
    Private Sub SUB_GET_TAB_CONFIG()
        Const cstTAB As String = "tab"
        Dim intLOOP_INDEX As Integer
        Dim strKEY As String
        Dim strTEMP As String

        ReDim srtMY_TAB_GROUP(0)

        For intLOOP_INDEX = 1 To 99
            strKEY = cstTAB & Format(intLOOP_INDEX, "00")
            strTEMP = ""
            strTEMP = System.Configuration.ConfigurationManager.AppSettings(strKEY)
            ReDim Preserve srtMY_TAB_GROUP(intLOOP_INDEX)
            If strTEMP = Nothing Then
                With srtMY_TAB_GROUP(intLOOP_INDEX)
                    .KIND_TABPAGE = ENM_MY_KIND_TABPAGE.KIND_NONE
                    .TEXT = ""
                End With
                'スルー
            Else
                With srtMY_TAB_GROUP(intLOOP_INDEX)
                    .KIND_TABPAGE = ENM_MY_KIND_TABPAGE.KIND_TABPAGE
                    .TEXT = strTEMP
                End With
            End If
        Next

    End Sub

    Private Sub SUB_SET_WINDOW_TAB()
        Dim tbpTEMP As TabPage
        For intLOOP_INDEX = 1 To UBound(srtMY_TAB_GROUP)
            ReDim Preserve tbpLAUNCHER_TABPAGE(intLOOP_INDEX)
            tbpLAUNCHER_TABPAGE(intLOOP_INDEX) = New TabPage
            If srtMY_TAB_GROUP(intLOOP_INDEX).KIND_TABPAGE = ENM_MY_KIND_TABPAGE.KIND_TABPAGE Then
                tbpTEMP = tbpLAUNCHER_TABPAGE(intLOOP_INDEX)
                tbpTEMP.Text = srtMY_TAB_GROUP(intLOOP_INDEX).TEXT
                tbpTEMP.BackColor = tabLAUNCHER_GROUP.Parent.BackColor
                tabLAUNCHER_GROUP.TabPages.Add(tbpTEMP)
            End If
        Next
    End Sub

    Private Sub SUB_GET_LAUNCHER_CONFIG()
        Const cstLAUNECHER As String = "launcher"
        Dim intLOOP_INDEX As Integer
        Dim strKEY As String
        Dim strTEMP As String
        Const cstLEVEL_DEFAULT As Integer = 99

        ReDim srtMY_LAUNCHER_GROUP(0)
        For intLOOP_INDEX = 1 To 9999
            ReDim Preserve srtMY_LAUNCHER_GROUP(intLOOP_INDEX)
            strKEY = cstLAUNECHER & Format(intLOOP_INDEX, "0000")
            strTEMP = ""
            strTEMP = System.Configuration.ConfigurationManager.AppSettings(strKEY)
            If strTEMP = Nothing Then
                With srtMY_LAUNCHER_GROUP(intLOOP_INDEX)
                    .KIND_LAUNCHER = ENM_MY_KIND_LAUNCHER.KIND_NONE
                    .TEXT = ""
                    .COMMAND = ""
                    .LEVEL = cstLEVEL_DEFAULT
                End With
            Else
                With srtMY_LAUNCHER_GROUP(intLOOP_INDEX)
                    .KIND_LAUNCHER = CInt(FUNC_SPLIT_STR(strTEMP, 0))
                    .TEXT = FUNC_SPLIT_STR(strTEMP, 1)
                    .COMMAND = FUNC_SPLIT_STR(strTEMP, 2)
                    strTEMP = FUNC_SPLIT_STR(strTEMP, 3)
                    If IsNumeric(strTEMP) Then
                        .LEVEL = CInt(strTEMP)
                    Else
                        .LEVEL = cstLEVEL_DEFAULT
                    End If
                End With
            End If
        Next
    End Sub

    Private Sub SUB_SET_WINDOW_LAUNCHER()
        Dim intLOOP_INDEX As Integer
        Dim intINDEX As Integer

        Dim strTEMP As String
        Dim intPAGE As Integer
        Dim intX As Integer
        Dim intY As Integer
        Dim tbpADD As TabPage

        ReDim btnLAUNCHER_DETAIL(0)
        ReDim lblLAUNCHER_DETAIL(0)

        For intLOOP_INDEX = 1 To UBound(srtMY_LAUNCHER_GROUP)
            strTEMP = Format(intLOOP_INDEX, "0000")
            intPAGE = CInt(strTEMP.Substring(0, 2))
            intX = CInt(strTEMP.Substring(2, 1))
            intY = CInt(strTEMP.Substring(3, 1))

            tbpADD = Nothing
            Call SUB_GET_ADD_TABPAGE(tbpADD, intPAGE)

            Select Case srtMY_LAUNCHER_GROUP(intLOOP_INDEX).KIND_LAUNCHER
                Case ENM_MY_KIND_LAUNCHER.KIND_LABEL
                    intINDEX = UBound(lblLAUNCHER_DETAIL) + 1
                    ReDim Preserve lblLAUNCHER_DETAIL(intINDEX)
                    lblLAUNCHER_DETAIL(intINDEX) = New Label
                    Call SUB_LABEL_INIT(lblLAUNCHER_DETAIL(intINDEX), intX, intY, srtMY_LAUNCHER_GROUP(intLOOP_INDEX).TEXT)
                    Call tbpADD.Controls.Add(lblLAUNCHER_DETAIL(intINDEX))
                Case ENM_MY_KIND_LAUNCHER.KIND_BUTTON
                    intINDEX = UBound(btnLAUNCHER_DETAIL) + 1
                    ReDim Preserve btnLAUNCHER_DETAIL(intINDEX)
                    btnLAUNCHER_DETAIL(intINDEX) = New Button
                    Call SUB_BUTTON_INIT(btnLAUNCHER_DETAIL(intINDEX), intX, intY, srtMY_LAUNCHER_GROUP(intLOOP_INDEX).TEXT, intLOOP_INDEX)
                    btnLAUNCHER_DETAIL(intINDEX).Tag = intLOOP_INDEX
                    Call tbpADD.Controls.Add(btnLAUNCHER_DETAIL(intINDEX))
                    AddHandler btnLAUNCHER_DETAIL(intINDEX).Click, AddressOf btnLAUNCHER_BASE_Click
                Case ENM_MY_KIND_LAUNCHER.KIND_NONE
                    'スルー
                Case Else
                    'スルー
            End Select
        Next
    End Sub

    Private Sub SUB_LABEL_INIT( _
    ByRef lblBASE As Label, _
    ByVal intX As Integer, ByVal intY As Integer, _
    ByVal strTITLE As String _
    )
        Const cstLEFT_START As Integer = 10
        Const cstTOP_START As Integer = 10
        Dim intLEFT As Integer
        Dim intTOP As Integer

        intLEFT = cstLEFT_START + ((intX - 1) * (lblLAUNCHER_BASE.Width + 5))
        intTOP = cstTOP_START + ((intY - 1) * (lblLAUNCHER_BASE.Height + 10))

        With lblBASE
            .Font = lblLAUNCHER_BASE.Font
            .Anchor = lblLAUNCHER_BASE.Anchor
            .BorderStyle = lblLAUNCHER_BASE.BorderStyle
            .BackColor = lblLAUNCHER_BASE.BackColor

            .Left = intLEFT
            .Top = intTOP
            .Width = lblLAUNCHER_BASE.Width
            .Height = lblLAUNCHER_BASE.Height
            .Text = strTITLE

            .ForeColor = lblLAUNCHER_BASE.ForeColor
            .TextAlign = lblLAUNCHER_BASE.TextAlign

            .Visible = True
        End With
    End Sub

    Private Sub SUB_BUTTON_INIT( _
    ByRef btnBASE As Button, _
    ByVal intX As Integer, ByVal intY As Integer, _
    ByVal strTITLE As String, ByVal intINDEX As Integer _
    )

        Const cstLEFT_START As Integer = 10
        Const cstTOP_START As Integer = 10
        Dim intLEFT As Integer
        Dim intTOP As Integer

        intLEFT = cstLEFT_START + ((intX - 1) * (lblLAUNCHER_BASE.Width + 5))
        intTOP = cstTOP_START + ((intY - 1) * (lblLAUNCHER_BASE.Height + 10))

        With btnBASE
            .Font = btnLAUNCHER_BASE.Font
            .Anchor = btnLAUNCHER_BASE.Anchor
            .FlatStyle = btnLAUNCHER_BASE.FlatStyle

            .Left = intLEFT
            .Top = intTOP
            .Width = btnLAUNCHER_BASE.Width
            .Height = btnLAUNCHER_BASE.Height
            .Text = strTITLE
            .AutoEllipsis = btnLAUNCHER_BASE.AutoEllipsis

            .TextAlign = btnLAUNCHER_BASE.TextAlign
            .ForeColor = btnLAUNCHER_BASE.ForeColor
            .BackColor = btnLAUNCHER_BASE.BackColor
            .UseVisualStyleBackColor = btnLAUNCHER_BASE.UseVisualStyleBackColor
            .Tag = intINDEX
            .Visible = True
        End With

    End Sub

    Private Sub SUB_GET_ADD_TABPAGE(ByRef tbpRET As TabPage, ByVal intINDEX As Integer)
        tbpRET = tbpLAUNCHER_TABPAGE(intINDEX)
    End Sub

    Private Sub SUB_BUTTONS_ENABLED_REFRESH(ByVal intLEVEL As Integer)
        Dim intLOOP_INDEX As Integer
        Dim strTAG As String
        Dim intINDEX As Integer
        Dim blnENABLED As Boolean

        If IsNothing(btnLAUNCHER_DETAIL) Then
            Exit Sub
        End If

        For intLOOP_INDEX = 1 To (btnLAUNCHER_DETAIL.Length - 1)
            With btnLAUNCHER_DETAIL(intLOOP_INDEX)
                strTAG = .Tag
                If IsNumeric(strTAG) Then
                    intINDEX = CInt(strTAG)
                Else
                    intINDEX = 0
                End If

            End With

            blnENABLED = (intLEVEL <= srtMY_LAUNCHER_GROUP(intINDEX).LEVEL)
            btnLAUNCHER_DETAIL(intLOOP_INDEX).Enabled = blnENABLED
        Next

        Call Application.DoEvents()
    End Sub
#End Region

#Region "NEW"
    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Call SUB_CTRL_NEW_INIT()
    End Sub
#End Region

#Region "イベント-ボタンクリック"
    Private Sub btnLOG_OFF_Click(sender As Object, e As EventArgs) Handles btnLOG_OFF.Click
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_LOG_OFF)
    End Sub

    Private Sub btnBACK_Click(sender As Object, e As EventArgs)
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_BACK)
    End Sub

    Private Sub btnEND_Click(sender As Object, e As EventArgs) Handles btnEND.Click
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_END)
    End Sub

    Private Sub btnLAUNCHER_BASE_Click(sender As Object, e As EventArgs) Handles btnLAUNCHER_BASE.Click
        Dim intINDEX As Integer

        If Not IsNumeric(sender.Tag) Then
            Exit Sub
        End If
        intINDEX = CInt(sender.Tag)
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_SHELL, intINDEX)
    End Sub

#End Region

    Private Sub FRM_MAIN_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Call SUB_FIRST_ACTIVE()
    End Sub

    Private Sub FRM_MAIN_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call SUB_CTRL_VIEW_INIT()
        Call SUB_CTRL_VALUE_INIT()
    End Sub

    Private Sub FRM_MAIN_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Enabled = False
        Call Application.DoEvents()
    End Sub

    Private Sub FRM_MAIN_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case True
            Case e.Control
            Case e.Alt
                Call SUB_KEY_DOWN_ALT(e.KeyCode, e.Handled)
            Case e.Shift
            Case Else
                Call SUB_KEY_DOWN(e.KeyCode, e.Handled)
        End Select
    End Sub

End Class
