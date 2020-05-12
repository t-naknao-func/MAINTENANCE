Public Module MOD_CTRL_TOOLS

#Region "列挙定数"
    Public Enum ENM_MOVE_FOCUS_TYPE
        FOCUS_NEXT = 1 '次へ進む
        FOCUS_PREV = 2 '前へ戻る
    End Enum
#End Region

#Region "接頭語・接尾語関係定数"
    Private Const cst_COMMON_CONTROL_PREFIX_TEXTBOX As String = "TXT"
    Private Const cst_COMMON_CONTROL_PREFIX_COMBOBOX As String = "CMB"
    Private Const cst_COMMON_CONTROL_PREFIX_DATETIMEPICKER As String = "DTP"
    Private Const cst_COMMON_CONTROL_PREFIX_LABEL As String = "LBL"

    Private Const cst_COMMON_CONTROL_SUFFIX_FROM As String = "FROM"
    Private Const cst_COMMON_CONTROL_SUFFIX_TO As String = "TO"
    Private Const cst_COMMON_CONTROL_SUFFIX_FIRST As String = "FIRST"
    Private Const cst_COMMON_CONTROL_SUFFIX_LAST As String = "LAST"

    Private Const cst_COMMON_CONTROL_SUFFIX_GUIDE As String = "GUIDE"
    Private Const cst_COMMON_CONTROL_SUFFIX_NAME As String = "NAME"
#End Region

#Region "フォーカス制御部"

    'フォーカスの移動
    Public Sub SUB_CONTROL_FOCUS_MOVE(ByVal enmFocusType As ENM_MOVE_FOCUS_TYPE)
        Select Case enmFocusType
            Case ENM_MOVE_FOCUS_TYPE.FOCUS_PREV
                Call System.Windows.Forms.SendKeys.SendWait("+{TAB}")
            Case ENM_MOVE_FOCUS_TYPE.FOCUS_NEXT
                Call System.Windows.Forms.SendKeys.SendWait("{TAB}")
        End Select

        Call System.Windows.Forms.Application.DoEvents()
    End Sub

    'フォーカスの移動-イロイロ無視した高速版
    Public Sub SUB_CONTROL_FOCUS_MOVE_HIGH_SPEED(ByVal enmFocusType As ENM_MOVE_FOCUS_TYPE)
        Select Case enmFocusType
            Case ENM_MOVE_FOCUS_TYPE.FOCUS_PREV
                Call System.Windows.Forms.SendKeys.Send("+{TAB}")
            Case ENM_MOVE_FOCUS_TYPE.FOCUS_NEXT
                Call System.Windows.Forms.SendKeys.Send("{TAB}")
        End Select

        Call System.Windows.Forms.Application.DoEvents()
    End Sub

    '入力コントロールで一番TAB順の若い物を探索する(フレーム非対応)
    Public Function FUNC_SEARCH_INPUT_CONTROL_TAB_FIRST( _
    ByRef objCONTENA As Object _
    ) As System.Windows.Forms.Control
        Dim ctlRET As System.Windows.Forms.Control
        Dim ctlCUR_CONTROL As System.Windows.Forms.Control
        Dim intLOOP_INDEX As Integer
        Dim intCUR_TAB_INDEX_PANEL As Integer
        Dim intCUR_TAB_INDEX As Integer
        Dim ctlTEMP As System.Windows.Forms.Control

        intCUR_TAB_INDEX = 65536
        intCUR_TAB_INDEX_PANEL = -1

        ctlRET = Nothing
        intLOOP_INDEX = -1
        For Each ctlCUR_CONTROL In objCONTENA.Controls
            intLOOP_INDEX += 1

            Select Case True
                Case TypeOf ctlCUR_CONTROL Is System.Windows.Forms.Panel _
                Or TypeOf ctlCUR_CONTROL Is System.Windows.Forms.GroupBox
                    ctlTEMP = FUNC_SEARCH_INPUT_CONTROL_TAB_FIRST(ctlCUR_CONTROL)
                    If Not (ctlTEMP Is Nothing) Then
                        intCUR_TAB_INDEX_PANEL = ctlCUR_CONTROL.TabIndex
                        If intCUR_TAB_INDEX > ctlCUR_CONTROL.TabIndex Then 'タブ順が小さければ
                            intCUR_TAB_INDEX = ctlCUR_CONTROL.TabIndex
                            ctlRET = ctlTEMP
                        End If
                    End If
                Case TypeOf ctlCUR_CONTROL Is System.Windows.Forms.TextBox _
                Or TypeOf ctlCUR_CONTROL Is System.Windows.Forms.ComboBox _
                Or TypeOf ctlCUR_CONTROL Is System.Windows.Forms.DateTimePicker _
                Or TypeOf ctlCUR_CONTROL Is System.Windows.Forms.CheckBox
                    If intCUR_TAB_INDEX > ctlCUR_CONTROL.TabIndex Then 'タブ順が小さければ
                        intCUR_TAB_INDEX = ctlCUR_CONTROL.TabIndex
                        ctlRET = ctlCUR_CONTROL
                    End If
                Case Else
                    'スルー(非入力コントロール)
            End Select
        Next

        Return ctlRET
    End Function

    'もっともTAB順の若いコントロールにフォーカスする
    Public Sub SUB_FOCUS_FIRST_INPUT_CONTROL( _
    ByRef objCONTENA As Object _
    )
        Dim ctlFOCUS As System.Windows.Forms.Control

        ctlFOCUS = FUNC_SEARCH_INPUT_CONTROL_TAB_FIRST(objCONTENA)

        If Not IsNothing(ctlFOCUS) Then
            Call ctlFOCUS.Focus()
            Call System.Windows.Forms.Application.DoEvents()
        End If

    End Sub
#End Region

#Region "プルダウン制御"
    'コンボの擬似プルダウン
    Public Sub SUB_CONTROL_COMBO_BOX_PULLDOWN( _
    ByRef frmOwner As System.Windows.Forms.Form _
    )
        Dim ctlCommbo As System.Windows.Forms.ComboBox

        If frmOwner.ActiveControl Is Nothing Then
            Exit Sub
        End If

        If Not TypeOf frmOwner.ActiveControl Is System.Windows.Forms.ComboBox Then
            Exit Sub
        End If

        ctlCommbo = frmOwner.ActiveControl

        ctlCommbo.DroppedDown = True
        Call System.Windows.Forms.Application.DoEvents()
    End Sub
#End Region

#Region "ボタンクリックの仮想制御"
    'ボタンコントロールの仮想クリック
    Public blnPubFocus As Boolean '仮想クリック連打対応用
    Public Sub SUB_CONTROL_BUTTON_CLICK( _
    ByRef btnClickButton As System.Windows.Forms.Button _
    )
        Dim blnFocus As Boolean

        If blnPubFocus Then
            Exit Sub '並列イベントは禁止
        End If

        If Not btnClickButton.Enabled Then
            Exit Sub '無効なら無視
        End If

        If Not btnClickButton.Visible Then
            Exit Sub '非表示なら無視
        End If

        blnFocus = btnClickButton.Focus()
        If blnFocus = False Then
            Exit Sub
        End If

        blnPubFocus = True
        System.Windows.Forms.Application.DoEvents()
        Call System.Windows.Forms.SendKeys.SendWait("{ENTER}")
        System.Windows.Forms.Application.DoEvents()
        blnPubFocus = False

    End Sub
#End Region

#Region "フォーカス制御、Enabled制御用のダミーコントロール関連"

    Public Sub SUB_ADD_TEXTBOX_AND_MOVE_FOCUS( _
    ByRef txtTextBox As System.Windows.Forms.TextBox, _
    ByRef frmParent As System.Windows.Forms.Form _
    )
        txtTextBox = New System.Windows.Forms.TextBox

        With txtTextBox
            .Top = -.Height
            .Left = -.Width
            .Visible = True
            .Enabled = True
        End With

        frmParent.Controls.Add(txtTextBox) 'ダミーのテキストボックスを生成
        txtTextBox.Focus() 'ダミーのテキストボックスにフォーカスを移動

    End Sub

    Public Sub SUB_REMOVE_TEXTBOX( _
    ByRef txtTextBox As System.Windows.Forms.TextBox, _
    ByRef frmParent As System.Windows.Forms.Form _
    )
        txtTextBox.Enabled = False
        frmParent.Controls.Remove(txtTextBox)
        txtTextBox.Dispose()
    End Sub

#End Region

#Region "DateTimePicker制御用"

    Public Sub SUB_CONTROL_INITALIZE_DateTimePicker( _
    ByRef dtpControl As System.Windows.Forms.DateTimePicker, _
    ByVal datMinDate As DateTime, ByVal datMaxDate As DateTime _
    )
        Try
            If datMinDate.Date > datMaxDate.Date Then
                Exit Sub
            End If

            dtpControl.MinDate = cstVB_DATE_MIN
            dtpControl.MaxDate = cstVB_DATE_MAX

            If datMinDate > cstVB_DATE_MIN Then
                dtpControl.MinDate = datMinDate
            End If
            If datMaxDate < cstVB_DATE_MAX Then
                dtpControl.MaxDate = datMaxDate
            End If
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Public Sub SUB_CONTROL_SET_VALUE_FORCE_DateTimePicker(ByRef dtpControl As System.Windows.Forms.DateTimePicker, ByVal datValueDate As DateTime)
        Dim datMinDate As DateTime
        Dim datMaxDate As DateTime
        Dim datSetDate As DateTime

        datMinDate = dtpControl.MinDate
        datMaxDate = dtpControl.MaxDate

        datSetDate = datValueDate
        If datValueDate > datMaxDate Then
            datSetDate = datMaxDate
        End If
        If datValueDate < datMinDate Then
            datSetDate = datMinDate
        End If

        Try
            dtpControl.Value = datSetDate
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Public Sub SUB_CONTROL_SET_VALUE_DateTimePicker(ByRef dtpControl As System.Windows.Forms.DateTimePicker, ByVal datValueDate As DateTime)
        Dim datMinDate As DateTime
        Dim datMaxDate As DateTime

        datMinDate = dtpControl.MinDate
        datMaxDate = dtpControl.MaxDate
        Try
            dtpControl.MinDate = cstVB_DATE_MIN
            dtpControl.MaxDate = cstVB_DATE_MAX

            dtpControl.Value = datValueDate
            dtpControl.MinDate = datMinDate
            dtpControl.MaxDate = datMaxDate
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

#End Region

#Region "コンボボックス制御用"
    Public Sub SUB_CONTROL_SET_ITEM_ComboBox( _
    ByRef ctlCONTROL As System.Windows.Forms.ComboBox, _
    Optional ByVal intSELECT_ITEM As Integer = 0 _
    )
        Dim intITEM_COUNT As Integer
        Dim intITEM_MAX_INDEX As Integer
        If IsNothing(ctlCONTROL) Then
            Exit Sub
        End If

        If IsNothing(ctlCONTROL.Items) Then
            Exit Sub
        End If

        intITEM_COUNT = ctlCONTROL.Items.Count
        intITEM_MAX_INDEX = (intITEM_COUNT - 1)

        If intSELECT_ITEM > intITEM_MAX_INDEX Then
            Exit Sub
        End If

        ctlCONTROL.SelectedIndex = intSELECT_ITEM
        Call ctlCONTROL.Refresh()
        Call System.Windows.Forms.Application.DoEvents()
    End Sub
#End Region

#Region "画面Window制御用"
    Public Function FUNC_WINDOW_ACTIVE(ByVal frmBASE As System.Windows.Forms.Form) As Boolean
        Dim strCAPTION As String

        strCAPTION = frmBASE.Text

        Try
            Call Microsoft.VisualBasic.Interaction.AppActivate(strCAPTION)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function
#End Region

#Region "接頭語・接尾語を基にした変換・探索など"

    '特定の入力コントロールの接頭語をはずした名称を取得する
    Public Function FUNC_GET_INPUT_CONTROL_MAIN_NAME( _
    ByRef ctlCONTROL As System.Windows.Forms.Control, _
    Optional ByVal blnREMOVE_SUFFIX As Boolean = True _
    ) As String
        Dim strRET As String

        Select Case True
            Case TypeOf ctlCONTROL Is System.Windows.Forms.TextBox
                strRET = FUNC_L_REMOVE_STRING(ctlCONTROL.Name, cst_COMMON_CONTROL_PREFIX_TEXTBOX & "_")
                strRET = If(blnREMOVE_SUFFIX, FUNC_REMOVE_SUFFIX(strRET), strRET)
            Case TypeOf ctlCONTROL Is System.Windows.Forms.ComboBox
                strRET = FUNC_L_REMOVE_STRING(ctlCONTROL.Name, cst_COMMON_CONTROL_PREFIX_COMBOBOX & "_")
                strRET = If(blnREMOVE_SUFFIX, FUNC_REMOVE_SUFFIX(strRET), strRET)
            Case TypeOf ctlCONTROL Is System.Windows.Forms.DateTimePicker
                strRET = FUNC_L_REMOVE_STRING(ctlCONTROL.Name, cst_COMMON_CONTROL_PREFIX_DATETIMEPICKER & "_")
                strRET = If(blnREMOVE_SUFFIX, FUNC_REMOVE_SUFFIX(strRET), strRET)
            Case Else
                strRET = ""
        End Select

        Return strRET
    End Function

    '特定のコントロール名称からサフィックスを省略する
    Private Function FUNC_REMOVE_SUFFIX( _
    ByVal strNAME_CONTROL As String _
    ) As String
        Dim strRET As String

        strRET = strNAME_CONTROL
        strRET = FUNC_R_REMOVE_STRING(strRET, "_" & cst_COMMON_CONTROL_SUFFIX_FROM) '"_FROM"を外す
        strRET = FUNC_R_REMOVE_STRING(strRET, "_" & cst_COMMON_CONTROL_SUFFIX_TO) '"_TO"を外す
        strRET = FUNC_R_REMOVE_STRING(strRET, "_" & cst_COMMON_CONTROL_SUFFIX_FIRST) '"_FIRST"を外す
        strRET = FUNC_R_REMOVE_STRING(strRET, "_" & cst_COMMON_CONTROL_SUFFIX_LAST) '"_LAST"を外す
        Return strRET
    End Function

    '特定の名称からラベルガイドのコントロール名称に変換する
    Public Function FUNC_GET_LABEL_GUIDE_CONTROL_NAME( _
    ByVal strMAIN_CONTROL_NAME As String _
    ) As String
        Dim strRET As String

        strRET = cst_COMMON_CONTROL_PREFIX_LABEL & "_" & strMAIN_CONTROL_NAME & "_" & cst_COMMON_CONTROL_SUFFIX_GUIDE

        Return strRET
    End Function

    '特定の名称からラベル名称のコントロール名称に変換する
    Public Function FUNC_GET_LABEL_NAME_CONTROL_NAME( _
    ByVal strMAIN_CONTROL_NAME As String _
    ) As String
        Dim strRET As String

        strRET = cst_COMMON_CONTROL_PREFIX_LABEL & "_" & strMAIN_CONTROL_NAME & "_" & cst_COMMON_CONTROL_SUFFIX_NAME

        Return strRET
    End Function

    '任意の名称のコントロールを画面から探索する
    Public Function FUNC_SEARCH_CONTROL_NAME( _
    ByRef objFORM As Object, ByVal strSERACH_NAME As String _
    ) As Integer
        Dim intRET As Integer
        Dim intLOOP_INDEX As Integer
        Dim ctlCURRENT_CONTROL As System.Windows.Forms.Control

        intRET = -1
        intLOOP_INDEX = -1
        For Each ctlCURRENT_CONTROL In objFORM.Controls
            intLOOP_INDEX += 1
            If ctlCURRENT_CONTROL.Name = strSERACH_NAME Then
                intRET = intLOOP_INDEX
                Exit For
            End If
        Next

        Return intRET
    End Function
#End Region

#Region "ラベルのフォントサイズ変更処理"
    Public Sub SUB_LABEL_FONT_CHANGE( _
ByRef lblBASE As System.Windows.Forms.Label _
)
        Dim strTITLE As String
        Dim sngFONT_SIZE As Integer

        Dim intLength1 As Integer
        Dim intLength2 As Integer
        Dim intLength3 As Integer
        Dim intLength4 As Integer

        Select Case lblBASE.Size.Width
            Case 128
                intLength1 = 7
                intLength2 = 8
                intLength3 = 9
                intLength4 = 10
            Case Else
                intLength1 = 7
                intLength2 = 8
                intLength3 = 9
                intLength4 = 10
        End Select


        strTITLE = lblBASE.Text
        With lblBASE
            Select Case strTITLE.Length
                Case Is <= intLength1
                    sngFONT_SIZE = 12
                Case Is <= intLength2
                    sngFONT_SIZE = 11
                Case Is <= intLength3
                    sngFONT_SIZE = 10
                Case Is <= intLength4
                    sngFONT_SIZE = 9
                Case Else
                    sngFONT_SIZE = 8
            End Select

            .Font = New System.Drawing.Font(lblBASE.Font.Name, sngFONT_SIZE)
        End With


    End Sub
#End Region

#Region "入力関連の補助"

    '入力された文字列を数値型(INT)のコードとして取得
    Public Function FUNC_GET_CODE_NUMERIC_INPUT_TEXT_INT(ByRef ctlINPUT_CONTROL As System.Windows.Forms.Control, Optional ByVal intMISS_VALUE As Integer = -1) As Integer
        Dim strTEXT As String
        Dim intCODE As Integer

        strTEXT = ctlINPUT_CONTROL.Text
        If IsNumeric(strTEXT) Then
            intCODE = CInt(strTEXT)
        Else
            intCODE = intMISS_VALUE
        End If

        Return intCODE
    End Function

    '特定の入力コントロールの名称ラベルに文字列を設定
    Public Sub SUB_SET_TEXT_TO_NAME_LABEL(ByRef ctlINPUT_CONTROL As System.Windows.Forms.Control, ByVal strTEXT As String, Optional ByVal blnREMOVE_SUFFIX As Boolean = False)
        Dim lblNAME As System.Windows.Forms.Label

        lblNAME = FUNC_GET_CONTROL_NAME_LABEL(ctlINPUT_CONTROL, blnREMOVE_SUFFIX)
        If lblNAME Is Nothing Then
            Exit Sub
        End If

        lblNAME.Text = strTEXT
        lblNAME = Nothing
    End Sub
#End Region

#Region "フォームの初期化"
    Public Enum ENM_COLOR_SET_LEVEL
        FORM = 0
        GROUP = 1
        PANEL = 2
        OBJ = 3
        UBOUND = OBJ
    End Enum

    Public Structure SRT_FORM_SETTING_PARAM
        Public TEXT As String
        Public FONT_SIZE As Single
        Public BACK_COLOR() As System.Drawing.Color
    End Structure

    Public Sub SUB_INIT_FORM_SETTING(ByRef frmMAIN As System.Windows.Forms.Form, ByRef srtPARAM As SRT_FORM_SETTING_PARAM)
        Dim sngFONT_SIZE As Single
        Dim setCOLOR As System.Drawing.Color
        With frmMAIN

            If srtPARAM.FONT_SIZE <= 0.0 Then
                sngFONT_SIZE = frmMAIN.Font.Size
            Else
                sngFONT_SIZE = srtPARAM.FONT_SIZE
            End If
            .Font = New System.Drawing.Font(frmMAIN.Font.Name, sngFONT_SIZE)
            .Text = srtPARAM.TEXT
            setCOLOR = srtPARAM.BACK_COLOR(ENM_COLOR_SET_LEVEL.FORM)
            If Not (setCOLOR = Drawing.Color.Empty) Then
                .BackColor = setCOLOR
            End If
            Dim ctlROW() As System.Windows.Forms.Control
            ctlROW = Nothing
            Call SUB_GET_ALL_CONTROL_FORM(frmMAIN, ctlROW)

            If ctlROW Is Nothing Then
                Exit Sub
            End If

            Dim intLOOP_INDEX As Integer
            Dim ctlCONTROL As System.Windows.Forms.Control
            ctlCONTROL = Nothing
            For intLOOP_INDEX = 0 To (ctlROW.Length - 1)
                ctlCONTROL = ctlROW(intLOOP_INDEX)
                Select Case True
                    Case TypeOf ctlCONTROL Is System.Windows.Forms.GroupBox
                        setCOLOR = srtPARAM.BACK_COLOR(ENM_COLOR_SET_LEVEL.GROUP)
                        If Not (setCOLOR = Drawing.Color.Empty) Then
                            ctlCONTROL.BackColor = setCOLOR
                        End If
                    Case TypeOf ctlCONTROL Is System.Windows.Forms.Panel
                        Dim ctlPANEL As System.Windows.Forms.Panel
                        ctlPANEL = ctlCONTROL
                        If ctlPANEL.BorderStyle = System.Windows.Forms.BorderStyle.None Then
                            Continue For
                        End If

                        Dim ctlPARENT As System.Windows.Forms.Control
                        ctlPARENT = ctlCONTROL.Parent
                        If TypeOf ctlPARENT Is System.Windows.Forms.Panel Then
                            setCOLOR = srtPARAM.BACK_COLOR(ENM_COLOR_SET_LEVEL.OBJ)
                            If Not (setCOLOR = Drawing.Color.Empty) Then
                                ctlCONTROL.BackColor = setCOLOR
                            End If
                        Else
                            setCOLOR = srtPARAM.BACK_COLOR(ENM_COLOR_SET_LEVEL.PANEL)
                            If Not (setCOLOR = Drawing.Color.Empty) Then
                                ctlCONTROL.BackColor = setCOLOR
                            End If
                        End If
                            Case Else
                        'スルー
                End Select
            Next

        End With
    End Sub

    Private Sub SUB_GET_ALL_CONTROL_FORM(ByRef objFORM As Object, ByRef ctlGET() As System.Windows.Forms.Control)
        
        For Each ctlCURRENT_CONTROL In objFORM.Controls
            Call SUB_ADD_CONTROL_ROW(ctlGET, ctlCURRENT_CONTROL)
            Select Case True
                Case TypeOf ctlCURRENT_CONTROL Is System.Windows.Forms.GroupBox
                    Call SUB_GET_ALL_CONTROL_FORM(ctlCURRENT_CONTROL, ctlGET)
                Case TypeOf ctlCURRENT_CONTROL Is System.Windows.Forms.Panel
                    Call SUB_GET_ALL_CONTROL_FORM(ctlCURRENT_CONTROL, ctlGET)
                Case Else
                    'スルー
            End Select
        Next

    End Sub

    Private Sub SUB_ADD_CONTROL_ROW(ByRef ctlROW() As System.Windows.Forms.Control, ByRef ctlCONTROL As System.Windows.Forms.Control)

        If ctlROW Is Nothing Then
            ReDim ctlROW(0)
            ctlROW(0) = ctlCONTROL
            Exit Sub
        End If

        Dim intINDEX As Integer
        intINDEX = ctlROW.Length
        ReDim Preserve ctlROW(intINDEX)
        ctlROW(intINDEX) = ctlCONTROL
    End Sub
#End Region

End Module
