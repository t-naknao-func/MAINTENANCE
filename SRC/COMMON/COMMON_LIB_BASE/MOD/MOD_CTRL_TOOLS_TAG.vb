Public Module MOD_CTRL_TOOLS_TAG
#Region "列挙定数"
    Public Enum CONTROL_CHECK_ERR_CODE
        CHECK_OK = 0
        ErrNull
        ErrNotNum
        ErrNotDate
        ErrBadChar
        ErrByteOver
        ErrZero
        ErrNotPlus
        ErrOverMaxValue
        ErrOverMinValue
        ErrBadFloat
        ErrNotSelected
        ErrNotHalfKana
        ErrNotAllKana
        ErrNotMonth
        ErrNotYear
        ErrNotYearMonth
        ErrNotAscii
    End Enum

#End Region

#Region "定数"
    Private Const cstDeffaultMojiCode As String = "Shift_JIS"

    Private Const cstTabooChar As String = ", ， ' ’" '禁止文字(半角スペース区切)"

    Private cstColor_Normal_TextBox As System.Drawing.Color = Drawing.SystemColors.Window
    Private cstColor_Focus_TextBox As System.Drawing.Color = Drawing.Color.Aqua
    Private cstColor_Disable_TextBox As System.Drawing.Color = Drawing.SystemColors.Control

    Private cstColor_Normal_ComboBox As System.Drawing.Color = Drawing.SystemColors.Window
    Private cstColor_Focus_ComboBox As System.Drawing.Color = Drawing.Color.Aqua
    Private cstColor_Disable_ComboBox As System.Drawing.Color = Drawing.SystemColors.Control

    Private cstColor_Normal_CheckBox As System.Drawing.Color = Drawing.SystemColors.Window
    Private cstColor_Focus_CheckBox As System.Drawing.Color = Drawing.SystemColors.ActiveCaption
#End Region

#Region "TAGプロパティ用の固定文字列"
    Private Const cstTagEraseWord As String = "Initialize" '初期化を行う
    Private Const cstTagClearWord As String = "Clear" '初期化を行う
    Private Const cstTagCheckWord As String = "Check" '一括チェックを行う
    Private Const cstTagCtxtWord As String = "Ctxt" '+キー入力時に文字情報のクリアを行う

    Private Const cstTagNumericWord As String = "Numeric" '数値項目,数値チェック
    Private Const cstTagNumberWord As String = "Number" '数値列項目,数値列チェック
    Private Const cstTagNotNullWord As String = "NotNull" 'Nullチェックを行う
    Private Const cstTagDateWord As String = "Check_Date" '日付チェックを行う
    Private Const cstTagCharWord As String = "Char" '禁止文字チェックを行う
    Private Const cstTagAsciiWord As String = "Ascii" '半角英数字チェックを行う
    Private Const cstTagNotZeroWord As String = "NotZero" '零チェックを行う
    Private Const cstTagPlusWord As String = "Plus" '負数チェックを行う
    Private Const cstTagByteWord As String = "Byte=" 'バイトチェックを行う
    Private Const cstTagFormatWord As String = "Format=" 'フォーカス消失時、書式整形を行う
    Private Const cstTagPadLeftWord As String = "PadLeft=" 'フォーカス消失時、左に文字を埋め込む
    Private Const cstTagPadRightWord As String = "PadRight=" 'フォーカス消失時、右に文字を埋め込む
    Private Const cstTagPadCharWord As String = "PadChar=" '埋め込み文字の指定
    Private Const cstTagMaxValueWord As String = "MaxValue=" '最大値チェック(数値用)
    Private Const cstTagMinValueWord As String = "MinValue=" '最小値チェック(数値用)
    Private Const cstTagFloatWord As String = "Float=" '小数項目,精度チェック
    Private Const cstTagSelectedWord As String = "Selected" '強制選択チェック
    Private Const cstTagYearMonthWord As String = "YearMonth" '年月(フォーカス時「YYYYMM」 未フォーカス時「YYYY/MM」)
    Private Const cstTagYearMonthDayWord As String = "DateSlush" '年月日(フォーカス時「YYYYMMDD」 未フォーカス時「YYYY/MM/DD」)
    Private Const cstTagYearWord As String = "Check_Year" '年チェック
    Private Const cstTagMonthWord As String = "Check_Month" '月チェック
    Private Const cstTagFocusColorWord As String = "FocusColor" 'フォーカス時色変更(通常時：白　フォーカス時：シアン)
    Private Const cstTagHalfKanaWord As String = "HalfKana" '半角カナ
    Private Const cstTagAllKanaWord As String = "AllKana" '全角カナ

#End Region

#Region "画面の初期化・クリア等"
    '初期化
    Public Sub SUB_CONTROL_INITIALIZE_FORM( _
    ByRef objHContena As Object, _
    Optional ByVal strHInitKeyword As String = cstTagEraseWord _
    )
        Dim ctlWCurCtrl As System.Windows.Forms.Control

        Call System.Windows.Forms.Application.DoEvents() '初期化前の待機イベントをすべて吐出す

        For Each ctlWCurCtrl In objHContena.Controls
            Select Case True
                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.Panel
                    Call SUB_CONTROL_INITIALIZE_FORM(ctlWCurCtrl) '再帰呼出
                Case Else
                    If InStr(ctlWCurCtrl.Tag, strHInitKeyword) > 0 Then
                        With ctlWCurCtrl
                            Select Case True
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.TextBox
                                    .Text = ""
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.Button
                                    .Text = ""
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.Label
                                    .Text = ""
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.ComboBox
                                    Dim ctlCurCombo As New System.Windows.Forms.ComboBox
                                    ctlCurCombo = ctlWCurCtrl
                                    ctlCurCombo.Text = ""
                                    ctlCurCombo.SelectedIndex = -1
                                    Call ctlCurCombo.Items.Clear()
                                    ctlCurCombo = Nothing
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.DateTimePicker
                                    Dim ctlCurDateTimePicker As New System.Windows.Forms.DateTimePicker
                                    ctlCurDateTimePicker = ctlWCurCtrl
                                    Call SUB_CONTROL_INITALIZE_DateTimePicker(ctlCurDateTimePicker, cstVB_DATE_MIN, cstVB_DATE_MAX)
                                    ctlCurDateTimePicker = Nothing
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.CheckBox
                                    Dim ctlCurCheckBox As New System.Windows.Forms.CheckBox
                                    ctlCurCheckBox = ctlWCurCtrl
                                    ctlCurCheckBox.Checked = False
                                    ctlCurCheckBox = Nothing
                            End Select
                        End With
                    End If
            End Select
        Next ctlWCurCtrl

        ctlWCurCtrl = Nothing
        Call System.Windows.Forms.Application.DoEvents() '初期化後のチェンジなどのイベントをすべて吐出す
    End Sub

    'クリア
    Public Sub SUB_CONTROL_CLEAR_FORM( _
    ByRef objHContena As Object, _
    Optional ByVal strHInitKeyword As String = cstTagClearWord _
    )
        Dim ctlWCurCtrl As System.Windows.Forms.Control

        Call System.Windows.Forms.Application.DoEvents() '初期化前の待機イベントをすべて吐出す

        For Each ctlWCurCtrl In objHContena.Controls
            Select Case True
                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.Panel
                    Call SUB_CONTROL_CLEAR_FORM(ctlWCurCtrl)
                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.GroupBox
                    Call SUB_CONTROL_CLEAR_FORM(ctlWCurCtrl)
                Case Else
                    If InStr(ctlWCurCtrl.Tag, strHInitKeyword) > 0 Then
                        With ctlWCurCtrl
                            Select Case True
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.TextBox
                                    .Text = ""
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.Button
                                    .Text = ""
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.Label
                                    .Text = ""
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.ComboBox
                                    Dim ctlCurCombo As New System.Windows.Forms.ComboBox
                                    ctlCurCombo = ctlWCurCtrl
                                    ctlCurCombo.Text = ""
                                    ctlCurCombo.SelectedIndex = -1
                                    ctlCurCombo = Nothing
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.DateTimePicker
                                    Dim ctlCurDateTimePicker As New System.Windows.Forms.DateTimePicker
                                    ctlCurDateTimePicker = ctlWCurCtrl
                                    ctlCurDateTimePicker.Value = ctlCurDateTimePicker.MaxDate
                                    ctlCurDateTimePicker = Nothing
                                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.CheckBox
                                    Dim ctlCurCheckBox As New System.Windows.Forms.CheckBox
                                    ctlCurCheckBox = ctlWCurCtrl
                                    ctlCurCheckBox.Checked = False
                                    ctlCurCheckBox = Nothing
                            End Select
                        End With
                    End If
            End Select
        Next ctlWCurCtrl

        ctlWCurCtrl = Nothing
        Call System.Windows.Forms.Application.DoEvents() '初期化後のチェンジなどのイベントをすべて吐出す
    End Sub
#End Region

#Region "共通フォーカスイベント用"
    Public Sub SUB_COMMON_EVENT_GOTFOCUS(ByRef ctlFocusCtrl As System.Windows.Forms.Control)
        With ctlFocusCtrl
            Select Case True
                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.TextBox

                    With ctlFocusCtrl
                        If InStr(.Tag, cstTagFocusColorWord) Then
                            .BackColor = cstColor_Focus_TextBox
                        End If

                        If InStr(.Tag, cstTagYearMonthWord) > 0 Then
                            .Text = .Text.Replace("/", "")
                        End If

                        If InStr(.Tag, cstTagYearMonthDayWord) > 0 Then
                            .Text = .Text.Replace("/", "")
                        End If

                        If InStr(.Tag, cstTagNumericWord) > 0 Then
                            If IsNumeric(.Text) And InStr(.Text, ",") > 0 Then
                                .Text = CDec(.Text)
                            End If
                        End If

                    End With

                    Dim ctlFocusText As New System.Windows.Forms.TextBox
                    ctlFocusText = ctlFocusCtrl
                    ctlFocusText.SelectAll()
                    ctlFocusText = Nothing
                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.Button

                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.DateTimePicker

                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.ComboBox
                    With ctlFocusCtrl
                        If InStr(.Tag, cstTagFocusColorWord) > 0 Then
                            .BackColor = cstColor_Focus_ComboBox
                        End If
                    End With
                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.CheckBox
                    With ctlFocusCtrl
                        If InStr(.Tag, cstTagFocusColorWord) > 0 Then
                            .BackColor = cstColor_Focus_CheckBox
                        End If
                    End With
            End Select
        End With
    End Sub

    Public Sub SUB_COMMON_EVENT_LOSTFOCUS(ByRef ctlFocusCtrl As System.Windows.Forms.Control)
        Dim intTemp As Integer
        Dim strTempString As String

        With ctlFocusCtrl
            Select Case True
                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.TextBox

                    With ctlFocusCtrl
                        If InStr(.Tag, cstTagFocusColorWord) Then
                            .BackColor = cstColor_Normal_TextBox
                        End If
                    End With

                    If InStr(.Tag, cstTagNumericWord) > 0 And .Text <> "" Then
                        If Not IsNumeric(.Text) Then
                            .Text = ""
                        End If
                    End If

                    If InStr(.Tag, cstTagYearMonthWord) > 0 Then
                        strTempString = .Text.Replace("/", "")
                        If IsNumeric(strTempString) Then
                            intTemp = CInt(strTempString)
                            .Text = Format(intTemp, "0000/00")
                        Else
                            .Text = ""
                        End If
                    End If

                    If InStr(.Tag, cstTagYearMonthDayWord) > 0 Then
                        strTempString = .Text.Replace("/", "")
                        If IsNumeric(strTempString) Then
                            intTemp = CInt(strTempString)
                            .Text = Format(intTemp, "0000/00/00")
                        Else
                            .Text = ""
                        End If
                    End If


                    If InStr(.Tag, cstTagFormatWord) > 0 And .Text <> "" Then
                        strTempString = FUNC_GET_TAG_EQUAL_AFTER_STR(.Tag, cstTagFormatWord)
                        Select Case True
                            Case InStr(.Tag, cstTagNumericWord)
                                If IsNumeric(.Text) Then
                                    .Text = Format(CDbl(.Text), strTempString)
                                End If
                            Case InStr(.Tag, cstTagDateWord)
                                If IsDate(.Text) Then
                                    .Text = Format(CDate(.Text), strTempString)
                                End If
                            Case Else
                                .Text = Format(.Text, strTempString)
                        End Select

                        If strTempString = "" Then
                            .Text = ""
                        End If
                    End If

                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.Button

                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.DateTimePicker

                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.ComboBox

                    With ctlFocusCtrl
                        If InStr(.Tag, cstTagFocusColorWord) Then
                            .BackColor = cstColor_Normal_ComboBox
                        End If
                    End With

                    If InStr(.Tag, cstTagFormatWord) > 0 And .Text <> "" Then
                        strTempString = FUNC_GET_TAG_EQUAL_AFTER_STR(.Tag, cstTagFormatWord)
                        Select Case True
                            Case InStr(.Tag, cstTagNumericWord)
                                If IsNumeric(.Text) Then
                                    .Text = Format(CDbl(.Text), strTempString)
                                End If
                            Case InStr(.Tag, cstTagDateWord)
                                If IsDate(.Text) Then
                                    .Text = Format(CDate(.Text), strTempString)
                                End If
                            Case Else
                                .Text = Format(.Text, strTempString)
                        End Select
                    End If
                Case TypeOf ctlFocusCtrl Is System.Windows.Forms.CheckBox
                    With ctlFocusCtrl
                        If InStr(.Tag, cstTagFocusColorWord) Then
                            '.BackColor = IIf(InStr(.Tag, Color_Chk_Main), MY_SYSTEM_CONTROL_COLOR.Chk_Main.BACK_COLOR, cstColor_Normal_CheckBox)
                        End If
                    End With
            End Select
        End With
    End Sub
#End Region

#Region "共通キープレスイベント用"
    'キープレス時の特殊操作
    Public Sub SUB_COMMON_EVENT_KEYPRESS( _
    ByRef frmWindow As System.Windows.Forms.Form, ByRef chrKeyChar As Char, ByRef blnHandled As Boolean _
    )
        Dim ctlFocusCtrl As System.Windows.Forms.Control
        ctlFocusCtrl = frmWindow.ActiveControl
        Try
            With ctlFocusCtrl
                Select Case True
                    Case TypeOf ctlFocusCtrl Is System.Windows.Forms.TextBox 'テキストボックス
                        If InStr(.Tag, cstTagCtxtWord) > 0 Then
                            Select Case chrKeyChar
                                Case "+"
                                    .Text = ""
                                    blnHandled = True
                                Case Else
                                    'スルー
                            End Select
                        End If

                        Select Case True
                            Case InStr(.Tag, cstTagNumericWord) > 0, _
                             InStr(.Tag, cstTagYearMonthWord) > 0, _
                             InStr(.Tag, cstTagNumberWord) > 0
                                Select Case chrKeyChar
                                    Case "0" To "9"
                                        'スルー
                                    Case "."
                                        If InStr(.Tag, cstTagFloatWord) > 0 Then
                                            'スルー
                                        Else
                                            blnHandled = True
                                        End If
                                    Case "-"
                                        If InStr(.Tag, cstTagPlusWord) > 0 Then
                                            blnHandled = True
                                        Else
                                            'スルー
                                        End If
                                    Case Else
                                        Dim intASCII As Integer
                                        intASCII = Convert.ToInt32(chrKeyChar)
                                        Select Case intASCII
                                            Case 8 'バックスペース
                                                'スルー
                                            Case Else
                                                blnHandled = True
                                        End Select
                                End Select

                            Case InStr(.Tag, cstTagHalfKanaWord) > 0 '半角文字
                                blnHandled = Not FUNC_CHECK_HALF_CHAR(chrKeyChar)

                            Case Else
                        End Select
                    Case TypeOf ctlFocusCtrl Is System.Windows.Forms.ComboBox 'コンボボックス
                        Select Case True
                            Case InStr(.Tag, cstTagNumericWord) > 0 '数値項目
                                Select Case chrKeyChar
                                    Case "0" To "9"
                                        'スルー
                                    Case "."
                                        If InStr(.Tag, cstTagFloatWord) > 0 Then
                                            'スルー
                                        Else
                                            blnHandled = True
                                        End If
                                    Case Else
                                        Dim intASCII As Integer
                                        intASCII = Convert.ToInt32(chrKeyChar)
                                        Select Case intASCII
                                            Case 8 'バックスペース
                                                'スルー
                                            Case Else
                                                blnHandled = True
                                        End Select
                                End Select
                            Case Else
                        End Select
                End Select
            End With
        Catch ex As Exception
            'スルー
        End Try

    End Sub
#End Region

#Region "コントロールのTAGプロパティにおける単純チェック"

    'チェック処理
    Public Function FUNC_CONTROL_CHECK_INPUT(ByRef ctlCHECK_CONTROL As System.Windows.Forms.Control) As CONTROL_CHECK_ERR_CODE
        Dim enmRET_ERR_CODE As CONTROL_CHECK_ERR_CODE

        Select Case True
            Case TypeOf ctlCHECK_CONTROL Is System.Windows.Forms.TextBox
                enmRET_ERR_CODE = FUNC_CONTROL_CHECK_INPUT_TEXT_BOX(ctlCHECK_CONTROL)
            Case TypeOf ctlCHECK_CONTROL Is System.Windows.Forms.ComboBox
                enmRET_ERR_CODE = FUNC_CONTROL_CHECK_INPUT_COMBO_BOX(ctlCHECK_CONTROL)
            Case Else
                enmRET_ERR_CODE = CONTROL_CHECK_ERR_CODE.CHECK_OK
        End Select

        Call Debug.WriteLine(ctlCHECK_CONTROL.Name & " " & ":" & enmRET_ERR_CODE.ToString)

        Return enmRET_ERR_CODE
    End Function

    Public Function FUNC_GET_MESSAGE_CTRL_CHECK( _
    ByVal enmCODE_ERR As CONTROL_CHECK_ERR_CODE, _
    Optional ByVal strNAME_VALUE As String = "" _
    ) As String
        Dim strRET As String

        Select Case enmCODE_ERR
            Case CONTROL_CHECK_ERR_CODE.CHECK_OK
                strRET = ""
            Case CONTROL_CHECK_ERR_CODE.ErrNull
                If strNAME_VALUE = "" Then
                    strRET = "入力してください" '***を入力してください
                Else
                    strRET = strNAME_VALUE & "を" & "入力してください"  '***を入力してください
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotNum
                If strNAME_VALUE = "" Then
                    strRET = "数値を入力してください" '***は数値を入力してください
                Else
                    strRET = strNAME_VALUE & "には" & "数値を入力してください" '***は数値を入力してください
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotDate
                If strNAME_VALUE = "" Then
                    strRET = "日付を入力してください"
                Else
                    strRET = strNAME_VALUE & "には" & "正しい日付を入力してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrBadChar
                If strNAME_VALUE = "" Then
                    strRET = "禁則文字が使用されています"
                Else
                    strRET = strNAME_VALUE & "に" & "禁則文字が使用されています"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrByteOver
                If strNAME_VALUE = "" Then
                    strRET = "入力制限バイト数をオーバーしています"
                Else
                    strRET = strNAME_VALUE & "が" & "入力制限バイト数をオーバーしています"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrZero
                If strNAME_VALUE = "" Then
                    strRET = "「0」は入力できません" '***には0は入力できません
                Else
                    strRET = strNAME_VALUE & "には" & "「0」は入力できません"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotPlus
                If strNAME_VALUE = "" Then
                    strRET = "負数は入力できません" '***には負数は入力できません
                Else
                    strRET = strNAME_VALUE & "には" & "負数は入力できません"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrOverMaxValue
                If strNAME_VALUE = "" Then
                    strRET = "入力された数値が入力可能な最大値を超えています" & Environment.NewLine & "入力可能な範囲を確認してください"
                Else
                    strRET = strNAME_VALUE & "に" & "入力された数値が入力可能な最大値を超えています" & Environment.NewLine & "入力可能な範囲を確認してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrOverMinValue
                If strNAME_VALUE = "" Then
                    strRET = "入力された数値が入力可能な最小値を下回っています" & Environment.NewLine & "入力可能な範囲を確認してください"
                Else
                    strRET = strNAME_VALUE & "に" & "入力された数値が入力可能な最小値を下回っています" & Environment.NewLine & "入力可能な範囲を確認してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrBadFloat
                If strNAME_VALUE = "" Then
                    strRET = "入力された数値の精度が不正です" & Environment.NewLine & "小数部、または整数部の桁数を確認してください"
                Else
                    strRET = strNAME_VALUE & "に" & "入力された数値の精度が不正です" & Environment.NewLine & "小数部、または整数部の桁数を確認してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotSelected
                If strNAME_VALUE = "" Then
                    strRET = "コンボボックスから値を選択してください" '***を選択してください
                Else
                    strRET = strNAME_VALUE & "の" & "値が選択されていません" '***を選択してください
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotHalfKana
                If strNAME_VALUE = "" Then
                    strRET = "半角カナを入力してください"
                Else
                    strRET = strNAME_VALUE & "には" & "半角カナを入力してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotAllKana
                If strNAME_VALUE = "" Then
                    strRET = "全角カナを入力してください"
                Else
                    strRET = strNAME_VALUE & "には" & "全角カナを入力してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotYearMonth
                If strNAME_VALUE = "" Then
                    strRET = "年月を入力してください"
                Else
                    strRET = strNAME_VALUE & "には" & "正しい年月を入力してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotYear
                If strNAME_VALUE = "" Then
                    strRET = "年を入力してください"
                Else
                    strRET = strNAME_VALUE & "には" & "正しい年を入力してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotMonth
                If strNAME_VALUE = "" Then
                    strRET = "月を入力してください"
                Else
                    strRET = strNAME_VALUE & "には" & "正しい月を入力してください"
                End If
            Case CONTROL_CHECK_ERR_CODE.ErrNotAscii
                If strNAME_VALUE = "" Then
                    strRET = "英数字を入力してください"
                Else
                    strRET = strNAME_VALUE & "には" & "英数字を入力してください"
                End If
            Case Else
                strRET = ""
        End Select

        Return strRET
    End Function

#Region "各種類コントロール用単純チェック"
    'テキストボックスチェック処理
    Public Function FUNC_CONTROL_CHECK_INPUT_TEXT_BOX( _
    ByRef ctlCurCtrl As System.Windows.Forms.TextBox _
    ) As CONTROL_CHECK_ERR_CODE
        Dim strTEMP_STRING As String

        With ctlCurCtrl
            If .Visible = False Then
                Return CONTROL_CHECK_ERR_CODE.CHECK_OK '不可視のものはチェックしない
            End If

            If .Enabled = False Then
                Return CONTROL_CHECK_ERR_CODE.CHECK_OK 'Disabledのものはチェックしない
            End If

            If InStr(.Tag, cstTagNotNullWord) > 0 Then
                If .Text = "" Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNull 'NullStringエラー
                End If
            End If

            If .Text = "" Then
                Return CONTROL_CHECK_ERR_CODE.CHECK_OK
            End If

            If InStr(.Tag, cstTagNumericWord) > 0 Then
                If IsNumeric(.Text) = False Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotNum '非数値エラー
                Else
                    If InStr(.Tag, cstTagNotZeroWord) > 0 Then
                        If CDbl(.Text) = 0 Then
                            Return CONTROL_CHECK_ERR_CODE.ErrZero '数値零エラー
                        End If
                    End If

                    If InStr(.Tag, cstTagPlusWord) > 0 Then
                        If CDbl(.Text < 0) Then
                            Return CONTROL_CHECK_ERR_CODE.ErrNotPlus '負数エラー
                        End If
                    End If

                    If InStr(.Tag, cstTagMaxValueWord) > 0 Then
                        If Not FUNC_CHECK_MAX_VALUE(.Text, .Tag) Then
                            Return CONTROL_CHECK_ERR_CODE.ErrOverMaxValue '最大値オーバーエラー
                        End If
                    End If

                    If InStr(.Tag, cstTagMinValueWord) > 0 Then
                        If Not FUNC_CHECK_MIN_VALUE(.Text, .Tag) Then
                            Return CONTROL_CHECK_ERR_CODE.ErrOverMinValue '最小値オーバーエラー
                        End If
                    End If

                    If InStr(.Tag, cstTagFloatWord) > 0 Then
                        If functionCheckFloat(.Text, .Tag) = False Then
                            Return CONTROL_CHECK_ERR_CODE.ErrBadFloat '小数精度エラー
                        End If
                    End If
                End If
            End If

            If InStr(.Tag, cstTagNumberWord) > 0 Then
                If Not IsNumeric(.Text) Then
                    If Not (InStr(.Text, "-") > 0) Then
                        Return CONTROL_CHECK_ERR_CODE.ErrNotNum '非数値エラー
                    End If
                End If
            End If

            If InStr(.Tag, cstTagDateWord) > 0 Then
                If IsNumeric(.Text) Then
                    strTEMP_STRING = Format(CLng(.Text), "0000/00/00")
                Else
                    strTEMP_STRING = .Text
                End If

                If IsDate(strTEMP_STRING) = False Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotDate '日付エラー
                Else
                    .Text = Format(CDate(strTEMP_STRING), "yyyy/MM/dd")
                End If
            End If

            If InStr(.Tag, cstTagYearMonthWord) > 0 Then
                strTEMP_STRING = .Text.Replace("/", String.Empty)
                If IsNumeric(strTEMP_STRING) Then
                    strTEMP_STRING = Format(CLng(strTEMP_STRING), "0000/00") & "/01"
                Else
                    strTEMP_STRING = .Text
                End If

                If IsDate(strTEMP_STRING) = False Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotYearMonth '年月エラー
                Else
                    .Text = Format(CDate(strTEMP_STRING), "yyyy/MM")
                End If

                If FUNC_GET_YYYYMM_FROM_DATE(strTEMP_STRING) < FUNC_GET_YYYYMM_FROM_DATE(cstVB_DATE_MIN) Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotYearMonth  '有効稼動年月チェック
                End If
            Else
                If InStr(.Tag, cstTagYearWord) > 0 Then
                    strTEMP_STRING = .Text

                    If CInt(strTEMP_STRING) < FUNC_GET_YYYY_FROM_YYYYMM(FUNC_GET_YYYYMM_FROM_DATE(cstVB_DATE_MIN)) Then
                        Return CONTROL_CHECK_ERR_CODE.ErrNotYear  '有効稼動年度チェック
                    End If

                    If IsNumeric(strTEMP_STRING) Then
                        strTEMP_STRING = Format(CLng(strTEMP_STRING), "0000") & "/01/01"
                    End If

                    If IsDate(strTEMP_STRING) = False Then
                        Return CONTROL_CHECK_ERR_CODE.ErrNotYear  '年エラー
                    End If
                End If

                If InStr(.Tag, cstTagMonthWord) > 0 Then
                    strTEMP_STRING = .Text
                    Dim intMONTH As Integer

                    intMONTH = CInt(strTEMP_STRING)
                    Select Case intMONTH
                        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
                            'スルー
                        Case Else
                            Return CONTROL_CHECK_ERR_CODE.ErrNotMonth  '月エラー
                    End Select
                End If

            End If

            If InStr(.Tag, cstTagCharWord) > 0 Then
                If funcCheckTabooChar(.Text) = False Then
                    Return CONTROL_CHECK_ERR_CODE.ErrBadChar '禁則文字エラー
                End If
            End If

            If InStr(.Tag, cstTagHalfKanaWord) > 0 Then
                If Not FUNC_CHECK_HALF_CHAR(.Text) Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotHalfKana '半角エラー
                End If
            End If

            If InStr(.Tag, cstTagAllKanaWord) > 0 Then

                If Not funcCheckAllKana(.Text) Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotAllKana '全角カナエラー
                End If
            End If

            If InStr(.Tag, cstTagByteWord) > 0 Then
                If funcCheckByteCnt(.Text, .Tag) = False Then
                    Return CONTROL_CHECK_ERR_CODE.ErrByteOver '文字列制限超エラー
                End If
            End If

            If InStr(.Tag, cstTagAsciiWord) > 0 Then
                If FUNC_CHECK_ASCII(.Text) = False Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotAscii '英数字エラー
                End If
            End If
        End With

        Return CONTROL_CHECK_ERR_CODE.CHECK_OK
    End Function

    'コンボボックスチェック処理
    Public Function FUNC_CONTROL_CHECK_INPUT_COMBO_BOX( _
    ByRef ctlCurCtrl As System.Windows.Forms.ComboBox _
    ) As CONTROL_CHECK_ERR_CODE
        Dim strTEMP_STRING As String

        With ctlCurCtrl
            If .Visible = False Then
                Return CONTROL_CHECK_ERR_CODE.CHECK_OK   '不可視のものはチェックしない
            End If

            If .Enabled = False Then
                Return CONTROL_CHECK_ERR_CODE.CHECK_OK   'Disabledのものはチェックしない
            End If

            If InStr(.Tag, cstTagNotNullWord) > 0 Then
                If .Text = "" Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNull 'NullStringエラー
                End If
            End If

            If .Text = "" And .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown Then
                Return CONTROL_CHECK_ERR_CODE.CHECK_OK
            End If

            If InStr(.Tag, cstTagSelectedWord) > 0 Then
                If .SelectedIndex = -1 Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotSelected '非選択エラー
                End If
            End If

            If .SelectedIndex >= 0 Then
                Return CONTROL_CHECK_ERR_CODE.CHECK_OK 'コンボから選択されていればOK
            End If

            If .Text = "" Then
                Return CONTROL_CHECK_ERR_CODE.CHECK_OK
            End If

            If InStr(.Tag, cstTagNumericWord) > 0 Then
                If IsNumeric(.Text) = False Then
                    Return CONTROL_CHECK_ERR_CODE.ErrNotNum '非数値エラー
                Else
                    If InStr(.Tag, cstTagNotZeroWord) > 0 Then
                        If CDbl(.Text) = 0 Then
                            Return CONTROL_CHECK_ERR_CODE.ErrZero '数値零エラー
                        End If
                    End If

                    If InStr(.Tag, cstTagPlusWord) > 0 Then
                        If CDbl(.Text < 0) Then
                            Return CONTROL_CHECK_ERR_CODE.ErrNotPlus '負数エラー
                        End If
                    End If

                    If InStr(.Tag, cstTagFloatWord) > 0 Then
                        If functionCheckFloat(.Text, .Tag) = False Then
                            Return CONTROL_CHECK_ERR_CODE.ErrBadFloat '小数精度エラー
                        End If
                    End If

                End If

                If InStr(.Tag, cstTagDateWord) > 0 Then
                    strTEMP_STRING = Format(Format(.Text, "00000000"), "0000/00/00")

                    If IsDate(strTEMP_STRING) = False Then
                        Return CONTROL_CHECK_ERR_CODE.ErrNotDate '日付エラー
                    Else
                        .Text = Format(CDate(strTEMP_STRING), "YYYYMMDD")
                    End If
                End If

                If InStr(.Tag, cstTagCharWord) > 0 Then
                    If funcCheckTabooChar(.Text) = False Then
                        Return CONTROL_CHECK_ERR_CODE.ErrBadChar '禁則文字エラー
                    End If
                End If

                If InStr(.Tag, cstTagByteWord) > 0 Then
                    If funcCheckByteCnt(.Text, .Tag) = False Then
                        Return CONTROL_CHECK_ERR_CODE.ErrByteOver '文字列制限超エラー
                    End If
                End If
            End If
        End With

        Return CONTROL_CHECK_ERR_CODE.CHECK_OK
    End Function
#End Region

#End Region

#Region "TAGプロパティ単純チェック-画面用"

    Private Structure SRT_ERR_CODE_CONTROLS
        Public CONTROL As System.Windows.Forms.Control
        Public ERR_CODE As CONTROL_CHECK_ERR_CODE
        Public TAB_INDEX As Long
    End Structure

    Public Function FUNC_CONTROL_CHECK_INPUT_FORM_CONTROLS( _
    ByRef objFORM As Object, _
    ByRef ctlRET_CONTROL As System.Windows.Forms.Control, _
    ByRef enmRET_ERR_CODE As CONTROL_CHECK_ERR_CODE, _
    Optional ByVal strCHECK_TAG As String = cstTagCheckWord _
    ) As Boolean
        Dim intLOOP_INDEX As Integer
        Dim srtDATA() As SRT_ERR_CODE_CONTROLS
        Dim intCURRENT_TAB_INDEX As Long

        srtDATA = Nothing
        Call SUB_CONTROL_CHECK_INPUT_FORM_CONTROLS_MAIN(objFORM, srtDATA, strCHECK_TAG)

        If UBound(srtDATA) <= 0 Then
            Return True 'ALL_CHECK_OK
        End If

        intCURRENT_TAB_INDEX = 9999999999
        For intLOOP_INDEX = 1 To UBound(srtDATA)
            With srtDATA(intLOOP_INDEX)
                If intCURRENT_TAB_INDEX >= .TAB_INDEX Then
                    ctlRET_CONTROL = .CONTROL
                    enmRET_ERR_CODE = .ERR_CODE
                    intCURRENT_TAB_INDEX = .TAB_INDEX
                End If
            End With
        Next

        Return False 'CHECK_NG
    End Function

    Private Sub SUB_CONTROL_CHECK_INPUT_FORM_CONTROLS_MAIN( _
    ByRef objFORM As Object, _
    ByRef srtDATA() As SRT_ERR_CODE_CONTROLS, _
    ByVal strCHECK_TAG As String _
    )
        Dim intLOOP_INDEX As Integer
        Dim ctlWCurCtrl As System.Windows.Forms.Control
        Dim enmERR_CODE As CONTROL_CHECK_ERR_CODE
        Dim intINDEX As Integer

        Call Debug.WriteLine("-START GROUP CONTROL CHECK-" & " " & objFORM.Name)

        intLOOP_INDEX = -1
        If srtDATA Is Nothing Then
            ReDim srtDATA(0)
        End If

        For Each ctlWCurCtrl In objFORM.Controls
            Select Case True
                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.Panel
                    Call SUB_CONTROL_CHECK_INPUT_FORM_CONTROLS_MAIN(ctlWCurCtrl, srtDATA, strCHECK_TAG)
                Case TypeOf ctlWCurCtrl Is System.Windows.Forms.GroupBox
                    Call SUB_CONTROL_CHECK_INPUT_FORM_CONTROLS_MAIN(ctlWCurCtrl, srtDATA, strCHECK_TAG)
            End Select
            intLOOP_INDEX += 1

            If InStr(ctlWCurCtrl.Tag, strCHECK_TAG) <= 0 Then
                GoTo NEXT_FOR
            End If

            enmERR_CODE = FUNC_CONTROL_CHECK_INPUT(ctlWCurCtrl)
            If enmERR_CODE <> CONTROL_CHECK_ERR_CODE.CHECK_OK Then
                intINDEX = UBound(srtDATA) + 1
                ReDim Preserve srtDATA(intINDEX)
                With srtDATA(intINDEX)
                    .CONTROL = ctlWCurCtrl
                    .ERR_CODE = enmERR_CODE
                    '.TAB_INDEX = ctlWCurCtrl.TabIndex
                    .TAB_INDEX = FUNC_GET_TAB_INDEX_FORM_GROUP(ctlWCurCtrl)
                End With
            End If
NEXT_FOR:
        Next ctlWCurCtrl

        Call Debug.WriteLine("-END   GROUP CONTROL CHECK-")
        Call Debug.WriteLine("")

        ctlWCurCtrl = Nothing
    End Sub

    '特定の入力コントロールのガイドラベルの文字列を取得する
    Public Function FUNC_GET_TEXT_GUIDE_LABEL( _
    ByRef ctlINPUT As System.Windows.Forms.Control, _
    Optional ByVal blnREMOVE_SUFFIX As Boolean = True _
    ) As String
        Dim strRET As String
        Dim strGUIDE_NAME As String
        Dim intCONTROL_INDEX As Integer
        Dim ctlGUIDE As System.Windows.Forms.Control

        strGUIDE_NAME = FUNC_GET_LABEL_GUIDE_CONTROL_NAME(FUNC_GET_INPUT_CONTROL_MAIN_NAME(ctlINPUT, blnREMOVE_SUFFIX))
        intCONTROL_INDEX = FUNC_SEARCH_CONTROL_NAME(ctlINPUT.Parent, strGUIDE_NAME)
        If intCONTROL_INDEX = -1 Then
            Return ""
        End If

        ctlGUIDE = ctlINPUT.Parent.Controls(intCONTROL_INDEX)
        strRET = ctlGUIDE.Text
        ctlGUIDE = Nothing
        Return strRET
    End Function

    '特定の入力コントロールの名称ラベルのコントロールを取得する
    Public Function FUNC_GET_CONTROL_NAME_LABEL( _
    ByRef ctlINPUT As System.Windows.Forms.Control, _
    Optional ByVal blnREMOVE_SUFFIX As Boolean = True _
    ) As System.Windows.Forms.Label
        Dim strCONTROL_NAME As String
        Dim intCONTROL_INDEX As Integer
        Dim ctlRET As System.Windows.Forms.Control

        strCONTROL_NAME = FUNC_GET_LABEL_NAME_CONTROL_NAME(FUNC_GET_INPUT_CONTROL_MAIN_NAME(ctlINPUT, blnREMOVE_SUFFIX))
        intCONTROL_INDEX = FUNC_SEARCH_CONTROL_NAME(ctlINPUT.Parent, strCONTROL_NAME)
        If intCONTROL_INDEX = -1 Then
            Return Nothing
        End If

        ctlRET = ctlINPUT.Parent.Controls(intCONTROL_INDEX)

        Return ctlRET
    End Function

    '画面総合のタブインデックスを取得
    '1フレームにコントロール99個、5フレーム入れ子まで対応
    Private Function FUNC_GET_TAB_INDEX_FORM_GROUP(ByRef ctlMAIN As System.Windows.Forms.Control) As Long
        Dim lngRET As Long
        Dim strTAB As String
        Dim strTEMP As String
        Dim intTAB_INDEX() As Integer
        Dim intLOOP_INDEX As Integer
        Dim intTAB As Integer
        Dim intINDEX As Integer

        intTAB_INDEX = Nothing
        Call SUB_MAKE_TAB_ROW(ctlMAIN, intTAB_INDEX)

        strTAB = ""
        For intLOOP_INDEX = 1 To 5
            intINDEX = (intTAB_INDEX.Length - 1)
            intINDEX -= (intLOOP_INDEX - 1)
            If intINDEX >= 1 Then
                intTAB = intTAB_INDEX(intINDEX)
            Else
                intTAB = 0
            End If
            strTEMP = Format(intTAB, "00")
            strTAB &= strTEMP
        Next

        If strTAB = "" Then
            lngRET = 0
        Else
            lngRET = CLng(strTAB)
        End If

        Return lngRET
    End Function

    Private Sub SUB_MAKE_TAB_ROW(ByRef ctlCONTROL As System.Windows.Forms.Control, ByRef intTAB_INDEX() As Integer)
        Dim objPARENT As Object
        Dim intINDEX As Integer

        If intTAB_INDEX Is Nothing Then
            ReDim intTAB_INDEX(0)
        End If

        intINDEX = (intTAB_INDEX.Length)
        ReDim Preserve intTAB_INDEX(intINDEX)
        intTAB_INDEX(intINDEX) = ctlCONTROL.TabIndex

        objPARENT = ctlCONTROL.Parent
        If TypeOf objPARENT Is System.Windows.Forms.Form Then
            objPARENT = Nothing
            Exit Sub
        End If

        Call SUB_MAKE_TAB_ROW(objPARENT, intTAB_INDEX) '再帰呼出
    End Sub
#End Region

#Region "TAGに関する共通部"
    '「=」付のタグの「=」の後ろの部分を返す
    Public Function FUNC_GET_TAG_EQUAL_AFTER_STR( _
    ByVal strTAG As String, ByVal strSEARCH As String _
    ) As String
        Dim strTAG_ONE As String
        Dim strRET As String

        strTAG_ONE = FUNC_GET_STRING_FROM_DELIMTED_STR_SEARCH(strTAG, ",", strSEARCH)
        strRET = FUNC_GET_STRING_FROM_DELIMTED_STR(strTAG_ONE, "=", 1)
        strRET = strRET.Replace("@", ",")

        Return strRET
    End Function
#End Region

#Region "内部用の共通チェック等"
    '最大値チェック
    Private Function FUNC_CHECK_MAX_VALUE(ByVal strTEXT As String, ByVal strTAG As String) As Boolean
        Dim strCHECK_MAX_VALUE As String
        Dim decCHECK_MAX_VALUE As Decimal
        Dim decVALUE As Decimal

        strCHECK_MAX_VALUE = FUNC_GET_TAG_EQUAL_AFTER_STR(strTAG, cstTagMaxValueWord)
        Try
            decCHECK_MAX_VALUE = CDec(strCHECK_MAX_VALUE)
        Catch ex As Exception
            Return False
        End Try

        Try
            decVALUE = CDec(strTEXT)
        Catch ex As Exception
            Return False
        End Try

        If decVALUE > decCHECK_MAX_VALUE Then
            Return False
        End If

        Return True
    End Function

    '最小値チェック
    Private Function FUNC_CHECK_MIN_VALUE(ByVal strTEXT As String, ByVal strTAG As String) As Boolean
        Dim strCHECK_MIN_VALUE As String
        Dim decCHECK_MIN_VALUE As Decimal
        Dim decVALUE As Decimal

        strCHECK_MIN_VALUE = FUNC_GET_TAG_EQUAL_AFTER_STR(strTAG, cstTagMinValueWord)
        Try
            decCHECK_MIN_VALUE = CDec(strCHECK_MIN_VALUE)
        Catch ex As Exception
            Return False
        End Try

        Try
            decVALUE = CDec(strTEXT)
        Catch ex As Exception
            Return False
        End Try

        If decVALUE < decCHECK_MIN_VALUE Then
            Return False
        End If

        Return True
    End Function

    '小数などのチェック
    Private Function functionCheckFloat(ByVal strText As String, ByVal strTag As String) As Boolean
        Dim strCheckFloat As String
        Dim objTemp As Object
        'Dim lngLeftVal As Long
        'Dim strRightVal As String
        Dim intMaxLeftLen As Integer
        Dim intMaxRightLen As Integer

        strCheckFloat = FUNC_GET_TAG_EQUAL_AFTER_STR(strTag, cstTagFloatWord)

        objTemp = Split(strCheckFloat, ".")
        intMaxLeftLen = CInt(objTemp(0))
        intMaxRightLen = CInt(objTemp(1))
        objTemp = Nothing

        Dim blnRET As Boolean

        blnRET = FUNC_CHECK_FLOAT(strText, intMaxLeftLen, intMaxRightLen)
        Return blnRET

        'If InStr(strText, ".") > 0 Then
        '    objTemp = Split(strText, ".")
        '    lngLeftVal = CLng(objTemp(0))
        '    strRightVal = CStr(objTemp(1))

        '    If Len(CStr(lngLeftVal)) > intMaxLeftLen Then
        '        Return False
        '    End If

        '    If Len(strRightVal) > intMaxRightLen Then
        '        Return False
        '    End If

        '    Return True
        'Else
        '    lngLeftVal = CLng(strText)
        '    Return (Len(CStr(lngLeftVal)) <= intMaxLeftLen)
        'End If

    End Function

    '小数精度のチェック
    Public Function FUNC_CHECK_FLOAT( _
    ByVal strVALUE_CHECK As String, _
    ByVal intLEN_MAX_LEFT As Integer, ByVal intLEN_MAX_RIGHT As Integer _
    ) As Boolean
        Dim lngVALUE_LEFT As Long
        Dim strVALUE_RIGHT As String
        Dim objTEMP As Object
        Dim blnRET As Boolean
        Dim strTEMP As String

        If Not IsNumeric(strVALUE_CHECK) Then '数値でなければ
            Return False 'エラー
        End If

        If InStr(strVALUE_CHECK, ".") <= 0 Then '"."がなければ整数部のみチェック
            lngVALUE_LEFT = CLng(strVALUE_CHECK)
            blnRET = (Len(CStr(lngVALUE_LEFT)) <= intLEN_MAX_LEFT)
            Return blnRET
        End If

        objTEMP = Split(strVALUE_CHECK, ".")
        strTEMP = CStr(objTEMP(0))
        If IsNumeric(strTEMP) Then
            lngVALUE_LEFT = CLng(strTEMP)
        Else
            lngVALUE_LEFT = 0
        End If
        lngVALUE_LEFT = Math.Abs(lngVALUE_LEFT) '符号がある場合に外す
        strVALUE_RIGHT = CStr(objTEMP(1))

        If Len(CStr(lngVALUE_LEFT)) > intLEN_MAX_LEFT Then
            Return False
        End If

        If Len(strVALUE_RIGHT) > intLEN_MAX_RIGHT Then
            Return False
        End If

        Return True
    End Function

    'バイト数のチェック
    Private Function funcCheckByteCnt(ByVal strText As String, ByVal strTag As String) As Boolean
        Dim intCheckByte As Integer
        Dim strCheckByte As String


        strCheckByte = FUNC_GET_TAG_EQUAL_AFTER_STR(strTag, cstTagByteWord)

        If IsNumeric(strCheckByte) = False Then
            Return True
        End If

        intCheckByte = CInt(strCheckByte)

        If intCheckByte = 0 Then
            Return True
        End If

        If Not FUNC_CHECK_BYTE_CNT_MAIN(strCheckByte, intCheckByte) Then
            Return False
        End If
        'Dim encEncode As System.Text.Encoding
        'encEncode = System.Text.Encoding.GetEncoding(cstDeffaultMojiCode)
        'If encEncode.GetByteCount(strText) > intCheckByte Then
        '    Return False
        'End If

        Return True
    End Function

    'バイト数のチェック(メイン)
    'グリッド入力で単体使用するためスコープをpublicへ
    Public Function FUNC_CHECK_BYTE_CNT_MAIN(ByVal strTEXT As String, ByVal intCHECK_BYTE As Integer) As Boolean
        Dim encENCODE As System.Text.Encoding

        encENCODE = System.Text.Encoding.GetEncoding(cstDeffaultMojiCode)

        If encENCODE.GetByteCount(strTEXT) > intCHECK_BYTE Then
            Return False
        End If

        Return True

    End Function

    '禁止文字のチェック
    'グリッド入力で単体使用するためスコープをpublicへ
    Public Function funcCheckTabooChar(ByVal strText As String) As Boolean
        Dim objTabooChar As Object
        Dim strCharTemp As String
        Dim intLoopIndex As Integer

        objTabooChar = Split(cstTabooChar, " ")

        For intLoopIndex = LBound(objTabooChar) To UBound(objTabooChar)
            strCharTemp = CStr(objTabooChar(intLoopIndex))
            If InStr(strText, strCharTemp) > 0 Then
                Return False
            End If
        Next intLoopIndex

        Return True

    End Function

    '全角カナのチェック
    Private Function funcCheckAllKana(ByVal strText As String) As Boolean
        Dim intLoopIndex As Integer

        For intLoopIndex = 1 To Len(strText)
            Dim chrOneChar As String

            chrOneChar = CChar(Mid(strText, intLoopIndex, 1))
            Select Case chrOneChar
                Case "ア" To "ン"
                    'スルー
                Case "ァ" To "ョ"
                    'スルー
                Case "ー", "―", "。", "＿", "　"
                    'スルー
                Case Else
                    Return False
            End Select
        Next

        Return True
    End Function

    '半角カナのチェック
    Private Function funcCheckHalfKana(ByVal strText As String) As Boolean
        Dim chrOneChar As String
        Dim intLoopIndex As Integer

        For intLoopIndex = 1 To Len(strText)
            chrOneChar = CChar(Mid(strText, intLoopIndex, 1))
            Select Case chrOneChar
                Case "ｱ" To "ﾝ"
                    'スルー
                Case "ｧ" To "ｮ"
                    'スルー
                Case "ﾞ", "ﾟ"
                    'スルー
                Case "ｰ", "-", ".", "_", " "
                    'スルー
                Case "\b" 'バックスペース
                    'スルー
                Case Else
                    Return False
            End Select
        Next

        Return True
    End Function

    '半角のチェック
    Private Function FUNC_CHECK_HALF_CHAR(ByVal strText As String) As Boolean
        Dim blnRET As Boolean

        blnRET = System.Text.Encoding.GetEncoding(cstDeffaultMojiCode).GetByteCount(strText) = strText.Length 'バイト数と文字数が同じ場合は半角文字（全て半角の場合はTrueを返す）
        Return blnRET
    End Function

    '英数字チェック
    Private Function FUNC_CHECK_ASCII(ByVal STR_TEXT As String) As Boolean

        For i = 1 To STR_TEXT.Length
            Dim CHR_ONE As String
            CHR_ONE = CChar(Mid(STR_TEXT, i, 1))
            If Not FUNC_CHECK_ASCII_CHAR(CHR_ONE) Then
                Return False
            End If
        Next

        Return True
    End Function

    '英数字チェック(1文字)
    Private Function FUNC_CHECK_ASCII_CHAR(ByVal CHR_TEXT As Char) As Boolean

        Select Case CHR_TEXT
            Case "a" To "z"
                    'スルー
            Case "A" To "Z"
                    'スルー
            Case "0" To "9"
                    'スルー
            Case ".", "-", "!", "#", "$", "%", "&", "(", ")", "=", "~", "|", "\", "@", "+", "*", "?", "/", ";", ":"
                'スルー
            Case Else
                Return False
        End Select

        Return True
    End Function
#End Region

End Module
