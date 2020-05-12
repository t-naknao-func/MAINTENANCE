Public Class FRM_MAIN

#Region "画面用・変数"
    Private blnPROCESS_DOING As Boolean
#End Region

#Region "画面用・列挙定数"
    Private Enum ENM_MY_EXEC_DO
        DO_LOG_ON = 1

        DO_END = 81
        DO_SHOW_SETTING
        DO_SHOW_COMMANDLINE
        DO_SHOW_CONFIG_SETTINGS
    End Enum
#End Region

#Region "初期化・終了処理"
    Private Sub SUB_CTRL_NEW_INIT()
        Call SUB_CTRL_EVENT_HANDLED_ADD()
    End Sub

    Private Sub SUB_CTRL_VIEW_INIT()

    End Sub

    Private Sub SUB_CTRL_VALUE_INIT()
        Call SUB_CONTROL_CLEAR_FORM(Me)
    End Sub
#End Region

#Region "各処理呼出元"
    Private Sub SUB_EXEC_DO( _
    ByVal enmEXEC_DO As ENM_MY_EXEC_DO _
    )
        If blnPROCESS_DOING Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        blnPROCESS_DOING = True
        Call Application.DoEvents()

        Select Case enmEXEC_DO
            Case ENM_MY_EXEC_DO.DO_LOG_ON
                Call SUB_LOG_ON()
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
    Private Sub SUB_LOG_ON()
        Dim strSHELL As String
        Dim strCOMMANDLINE As String
        Dim strUSER_ID As String
        Dim intCODE_STAFF As Integer

        If Not FUNC_CHECK_INPUT() Then
            Exit Sub
        End If

        strUSER_ID = TXT_ID_USER.Text
        intCODE_STAFF = FUNC_GET_MNG_M_USER_CODE_STAFF(strUSER_ID)

        strSHELL = "MENU_APP_LAUNCHER.EXE"
        strCOMMANDLINE = FUNC_SYSTEM_TOTAL_MAKE_COMMANDLINE(Application.ExecutablePath, intCODE_STAFF, 0, 0)
        If Not FUNC_CALL_EXE_FILE_SHELL(strSHELL, strCOMMANDLINE) Then
            Call MessageBox.Show("メニューを呼び出せません")
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

#End Region

#Region "チェック処理"
    Private Function FUNC_CHECK_INPUT() As Boolean
        Dim ctlCONTROL As Control
        Dim enmERR_CODE As CONTROL_CHECK_ERR_CODE
        Dim strERR_MSG As String
        Dim strTEXT As String
        Dim strUSER_ID As String
        Dim strUSER_ID_GET As String
        Dim strPASSWORD As String
        Dim strPASSWORD_GET As String
        Dim intCODE_STAFF As Integer
        Dim intFLAG_DELETE As Integer

        'Enable = True の入力項目すべてチェック対象(TAG=Check_Head)
        ctlCONTROL = Nothing
        If Not FUNC_CONTROL_CHECK_INPUT_FORM_CONTROLS(grpBODY, ctlCONTROL, enmERR_CODE, "Check") Then
            strERR_MSG = FUNC_GET_MESSAGE_CTRL_CHECK(enmERR_CODE, FUNC_GET_TEXT_GUIDE_LABEL(ctlCONTROL))
            Call MessageBox.Show(strERR_MSG, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Call ctlCONTROL.Focus()
            Return False
        End If

        'ユーザーID(この項目のチェック完了後はintCODE_STAFFが有効となる)
        ctlCONTROL = TXT_ID_USER
        strTEXT = ctlCONTROL.Text
        strUSER_ID = strTEXT
        intCODE_STAFF = FUNC_GET_MNG_M_USER_CODE_STAFF(strUSER_ID)
        If intCODE_STAFF <= 0 Then
            strERR_MSG = FUNC_GET_TEXT_GUIDE_LABEL(ctlCONTROL) & "にはマスタに登録されている値を入力してください"
            Call MessageBox.Show(strERR_MSG, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Call ctlCONTROL.Focus()
            Return False
        End If
        strUSER_ID_GET = FUNC_GET_MNG_M_USER_USER_ID(intCODE_STAFF) '再取得(照合順序によっては大文字小文字などの区別がされずマスタチェックはスルーする為)
        If Not (strUSER_ID.Equals(strUSER_ID_GET)) Then
            strERR_MSG = FUNC_GET_TEXT_GUIDE_LABEL(ctlCONTROL) & "にはマスタに登録されている値を入力してください"
            Call MessageBox.Show(strERR_MSG, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Call ctlCONTROL.Focus()
            Return False
        End If

        'パスワード
        ctlCONTROL = TXT_PASSWORD
        strTEXT = ctlCONTROL.Text
        strPASSWORD = strTEXT
        strPASSWORD_GET = FUNC_GET_MNG_M_USER_PASS_WORD(intCODE_STAFF)
        If Not (strPASSWORD.Equals(strPASSWORD_GET)) Then
            strERR_MSG = FUNC_GET_TEXT_GUIDE_LABEL(ctlCONTROL) & "が不正です"
            Call MessageBox.Show(strERR_MSG, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Call ctlCONTROL.Focus()
            Return False
        End If

        'その他入力項目との関連性のないチェック
        ctlCONTROL = TXT_ID_USER
        intFLAG_DELETE = FUNC_GET_MNG_M_USER_FLAG_DELETE(intCODE_STAFF)
        If FUNC_CAST_INT_TO_BOOL(intFLAG_DELETE) Then
            strERR_MSG = "入力された" & FUNC_GET_TEXT_GUIDE_LABEL(ctlCONTROL) & "は使用不可に設定されています"
            Call MessageBox.Show(strERR_MSG, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Call ctlCONTROL.Focus()
            Return False
        End If

        Return True
    End Function
#End Region

#Region "キー制御処理"
    '通常のコマンドキー制御(シフトマスク無し)
    Private Sub SUB_KEY_DOWN(ByVal enmKEY_CODE As Windows.Forms.Keys, ByRef blnHandled As Boolean)
        Select Case enmKEY_CODE
            Case Keys.Enter
                blnHandled = True
                If Not FUNC_RETURN_KEYDOWN() Then
                    Exit Sub
                End If
                Call SUB_CONTROL_FOCUS_MOVE(ENM_MOVE_FOCUS_TYPE.FOCUS_NEXT)
            Case Else
                'スルー
        End Select
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

    Private Function FUNC_RETURN_KEYDOWN() As Boolean
        Dim ctlACTIVE As Control
        Dim blnRET As Boolean

        If IsNothing(Me.ActiveControl) Then
            Return False
        End If

        ctlACTIVE = Me.ActiveControl

        Select Case True
            Case Else
                blnRET = True
        End Select

        ctlACTIVE = Nothing

        Return blnRET

        Return True
    End Function
#End Region

#Region "NEW"
    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Call SUB_CTRL_NEW_INIT()
    End Sub
#End Region

#Region "イベントハンドル(フォーカス制御)"
    Private Sub SUB_CTRL_EVENT_HANDLED_ADD() '共通イベントハンドルの追加
        Call SUB_CTRL_EVENT_HANDLED_ADD_MAIN(Me)
    End Sub

    Private Sub SUB_CTRL_EVENT_HANDLED_ADD_MAIN(ByRef objCONTENA As Object)
        Dim ctlCUR_CTRL As Control

        For Each ctlCUR_CTRL In objCONTENA.Controls
            Select Case True
                Case TypeOf ctlCUR_CTRL Is GroupBox
                    Call SUB_CTRL_EVENT_HANDLED_ADD_MAIN(ctlCUR_CTRL)
                Case TypeOf ctlCUR_CTRL Is Panel
                    Call SUB_CTRL_EVENT_HANDLED_ADD_MAIN(ctlCUR_CTRL)
                Case TypeOf ctlCUR_CTRL Is TextBox _
                  Or TypeOf ctlCUR_CTRL Is ComboBox
                    AddHandler ctlCUR_CTRL.GotFocus, AddressOf SUB_CTRL_GOTFOCUS    'フォーカス取得
                    AddHandler ctlCUR_CTRL.LostFocus, AddressOf SUB_CTRL_LOSTFOCUS  'フォーカス喪失
                Case Else

            End Select
        Next
    End Sub

    Private Sub SUB_CTRL_EVENT_HANDLED_REMOVE() '共通イベントハンドルの削除
        Call SUB_CTRL_EVENT_HANDLED_REMOVE_MAIN(Me)
    End Sub

    Private Sub SUB_CTRL_EVENT_HANDLED_REMOVE_MAIN(ByRef objCONTENA As Object)
        Dim ctlCUR_CTRL As Control

        For Each ctlCUR_CTRL In objCONTENA
            Select Case True
                Case TypeOf ctlCUR_CTRL Is GroupBox
                    Call SUB_CTRL_EVENT_HANDLED_REMOVE_MAIN(ctlCUR_CTRL)
                Case TypeOf ctlCUR_CTRL Is Panel
                    Call SUB_CTRL_EVENT_HANDLED_REMOVE_MAIN(ctlCUR_CTRL)
                Case TypeOf ctlCUR_CTRL Is TextBox Or TypeOf ctlCUR_CTRL Is ComboBox
                    RemoveHandler ctlCUR_CTRL.GotFocus, AddressOf SUB_CTRL_GOTFOCUS   'フォーカス取得
                    RemoveHandler ctlCUR_CTRL.LostFocus, AddressOf SUB_CTRL_LOSTFOCUS  'フォーカス喪失
                Case Else

            End Select
        Next
    End Sub
#End Region

#Region "イベント-フォーカス取得"
    Private Sub SUB_CTRL_GOTFOCUS(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SUB_COMMON_EVENT_GOTFOCUS(sender)
    End Sub
#End Region

#Region "イベント-フォーカス喪失"
    Private Sub SUB_CTRL_LOSTFOCUS(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SUB_COMMON_EVENT_LOSTFOCUS(sender)
    End Sub
#End Region

#Region "イベント-ボタンクリック"
    Private Sub btnLOG_ON_Click(sender As Object, e As EventArgs) Handles btnLOG_ON.Click
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_LOG_ON)
    End Sub

    Private Sub btnEND_Click(sender As Object, e As EventArgs) Handles btnEND.Click
        Call SUB_EXEC_DO(ENM_MY_EXEC_DO.DO_END)
    End Sub
#End Region

    Private Sub FRM_MAIN_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call SUB_CTRL_VIEW_INIT()
        Call SUB_CTRL_VALUE_INIT()
    End Sub

    Private Sub FRM_MAIN_Activated(sender As Object, e As EventArgs) Handles Me.Activated

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
