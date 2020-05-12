Public Module MOD_COMBO_TOOL

#Region "定数"
    Private Const cstComboBoxGetMiss_NUM As Long = -1
    Private Const cstComboBoxGetMiss_STR As String = ""
#End Region

#Region "構造体"
    Public Structure srtLComboDataArray
        Public strLOneRecord() As String
    End Structure
#End Region

#Region "ITEM置換用クラス"
    Public Class clsGComboBoxData
        Public strGComboBoxData() As String
        Public intGValueIndex As Integer
        Public intGWidth() As Integer

        Public Sub New(ByRef strHCmbData() As String, ByRef intHWidth() As Integer, ByVal intHRetIndex As Integer)
            Dim intWIndex As Integer
            Dim intWloopIndex As Integer

            intWIndex = UBound(strHCmbData)
            ReDim strGComboBoxData(intWIndex)
            ReDim intGWidth(intWIndex)
            For intWloopIndex = LBound(strGComboBoxData) To UBound(strGComboBoxData)
                strGComboBoxData(intWloopIndex) = strHCmbData(intWloopIndex)
                If intHWidth Is Nothing Then
                    intGWidth(intWloopIndex) = Nothing
                Else
                    intGWidth(intWloopIndex) = intHWidth(intWloopIndex)
                End If
            Next
            intGValueIndex = intHRetIndex

        End Sub

        Public Overrides Function ToString() As String
            Return strGComboBoxData(intGValueIndex)
        End Function

        Public Sub glbDispose()
            ReDim strGComboBoxData(0)
            Erase strGComboBoxData
            ReDim intGWidth(0)
            Erase intGWidth
            intGValueIndex = 0
        End Sub

    End Class
#End Region

#Region "DRAWITEM置換用メソッド"

    Public Sub glbSubComboBox_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)
        Dim brsWBrushesText As System.Drawing.Brush
        Dim cmbDataWdata As clsGComboBoxData
        Dim intWCurLeft As Integer
        Dim penWLine As New System.Drawing.Pen(System.Drawing.Color.Brown)
        Dim intWLoopIndex As Integer
        Dim blnWLineDraw As Boolean

        Const intcstWMargin As Integer = 8

        With e

            cmbDataWdata = sender.Items.Item(.Index)

            If .State = System.Windows.Forms.DrawItemState.Selected Then
                brsWBrushesText = System.Drawing.SystemBrushes.HighlightText
            Else
                brsWBrushesText = System.Drawing.SystemBrushes.WindowText
            End If

            Call .DrawBackground()

            If IsNothing(cmbDataWdata.strGComboBoxData) Then
                Exit Sub
            End If

        End With

        blnWLineDraw = False
        intWCurLeft = 0 'cstWMargin
        With cmbDataWdata
            For intWLoopIndex = LBound(.strGComboBoxData) To UBound(.strGComboBoxData)
                If blnWLineDraw Then
                    e.Graphics.DrawLine(penWLine, intWCurLeft, e.Bounds.Top, intWCurLeft, e.Bounds.Bottom)
                    intWCurLeft = intWCurLeft + intcstWMargin
                Else
                    blnWLineDraw = True
                End If

                e.Graphics.DrawString(.strGComboBoxData(intWLoopIndex), e.Font, brsWBrushesText, intWCurLeft, e.Bounds.Y)
                intWCurLeft = intWCurLeft + cmbDataWdata.intGWidth(intWLoopIndex) + intcstWMargin
            Next
        End With

    End Sub

    Public Sub glbSubComboBox_DrawItem_ONE(ByVal sender As Object, _
                                           ByVal e As System.Windows.Forms.DrawItemEventArgs)
        Dim brsWBrushesText As System.Drawing.Brush
        Dim cmbDataWdata As clsGComboBoxData
        Dim intWCurLeft As Integer
        Dim blnWLineDraw As Boolean
        Dim intWSelectIndex As Integer
        Const intcstWMargin As Integer = 8

        With e

            If .Index < 0 Then
                intWSelectIndex = 0
            Else
                intWSelectIndex = .Index
            End If

            cmbDataWdata = sender.Items.Item(intWSelectIndex)

            If .State = System.Windows.Forms.DrawItemState.Selected Then
                brsWBrushesText = System.Drawing.SystemBrushes.HighlightText
            Else
                brsWBrushesText = System.Drawing.SystemBrushes.WindowText
            End If

            Call .DrawBackground()

            If IsNothing(cmbDataWdata.strGComboBoxData) Then
                Exit Sub
            End If

        End With

        blnWLineDraw = False
        intWCurLeft = 0 'cstMargin
        With cmbDataWdata
            e.Graphics.DrawString(.strGComboBoxData(0), e.Font, brsWBrushesText, intWCurLeft, e.Bounds.Y)
            intWCurLeft = intWCurLeft + cmbDataWdata.intGWidth(0) + intcstWMargin
        End With

    End Sub

#End Region

#Region "TOOL的な共通メソッド"

    Private Function prvFuncGetValue_Initial(ByRef cmbHComboBox As System.Windows.Forms.ComboBox) As Boolean
        With cmbHComboBox
            If .Text = "" Then
                .SelectedIndex = -1
            End If

            If .Items.Count = 0 Then
                Return False
            End If

            If Not TypeOf .Items.Item(0) Is clsGComboBoxData Then
                Return False
            End If
        End With
        Return True
    End Function

    Public Sub glbSubCOPY_COMBO_ITEM_NORMAL(ByRef cmbSOURCE_COMBO As System.Windows.Forms.ComboBox, ByRef cmbDESTINATION_COMBO As System.Windows.Forms.ComboBox)
        Dim intLoopIndex As Integer

        With cmbDESTINATION_COMBO

            For intLoopIndex = 0 To cmbSOURCE_COMBO.Items.Count - 1
                .Items.Add(cmbSOURCE_COMBO.Items(intLoopIndex))
            Next

        End With

    End Sub

#End Region

#Region "コード・名称コンボに関する共通部"

    'コンボボックスの終了処理
    Public Sub SUB_END_COMBO_VIEW(ByRef cmbComboBox As System.Windows.Forms.ComboBox)
        'クリア部
        With cmbComboBox
            .SelectedIndex = -1
            .Text = ""
            .Items.Clear()

            Try
                RemoveHandler .DrawItem, AddressOf glbSubComboBox_DrawItem
            Catch ex As Exception
                'スルー
            End Try
        End With
    End Sub

    '現在選択されているアイテムのコード部分を取得する
    Public Function FUNC_GET_COMBO_VIEW_CODE( _
    ByRef cmbComboBox As System.Windows.Forms.ComboBox, _
    Optional ByVal lngVALUE_MISS As Long = cstComboBoxGetMiss_NUM _
    ) As Long
        Dim intIndex As Integer
        Dim lngRet As Long
        Dim strTemp As String

        With cmbComboBox
            If .Text = "" Then
                .SelectedIndex = -1
            End If

            intIndex = .SelectedIndex
            strTemp = ""

            If .Items.Count = 0 Then
                If IsNumeric(.Text) Then
                    Return CLng(.Text)
                Else
                    Return lngVALUE_MISS
                End If
            End If

            If Not TypeOf .Items.Item(0) Is clsGComboBoxData Then
                Return lngVALUE_MISS
            End If

            If intIndex > -1 Then
                Try
                    strTemp = .Items.Item(intIndex).strGComboBoxData(0)
                Catch ex As Exception
                    strTemp = ""
                End Try
            End If

            If IsNumeric(strTemp) Then
                lngRet = CLng(strTemp)
            Else
                lngRet = lngVALUE_MISS
            End If

            If lngRet = lngVALUE_MISS Then
                strTemp = .Text

                If IsNumeric(strTemp) Then
                    lngRet = CLng(strTemp)
                Else
                    lngRet = lngVALUE_MISS
                End If

            End If

            Return lngRet
        End With
    End Function

    '現在選択されているアイテムの名称部分を取得する
    Public Function FUNC_GET_COMBO_VIEW_NAME(ByRef cmbComboBox As System.Windows.Forms.ComboBox) As String
        Dim intIndex As Integer
        Dim strTemp As String
        Dim lngCODE As Long
        Dim intLoopIndex As Integer
        Dim lngTemp As Long

        With cmbComboBox
            If .Text = "" Then
                .SelectedIndex = -1
            End If

            intIndex = .SelectedIndex
            strTemp = cstComboBoxGetMiss_STR

            If .Items.Count = 0 Then
                Return cstComboBoxGetMiss_STR
            End If

            If Not TypeOf .Items.Item(0) Is clsGComboBoxData Then
                Return cstComboBoxGetMiss_STR
            End If

            If intIndex > -1 Then
                Try
                    strTemp = .Items.Item(intIndex).strGComboBoxData(1)
                Catch ex As Exception
                    strTemp = cstComboBoxGetMiss_STR
                End Try
                Return strTemp
            End If

            lngCODE = FUNC_GET_COMBO_VIEW_CODE(cmbComboBox)

            If lngCODE = cstComboBoxGetMiss_NUM Then
                Return cstComboBoxGetMiss_STR
            End If

            For intLoopIndex = 0 To .Items.Count - 1
                lngTemp = CLng(.Items.Item(intLoopIndex).strGComboBoxData(0))
                If lngCODE = lngTemp Then
                    strTemp = .Items.Item(intLoopIndex).strGComboBoxData(1)
                    Return strTemp
                End If
            Next

        End With

        Return cstComboBoxGetMiss_STR

    End Function

    'アイテム内の任意のコードを選択する
    Public Sub SUB_SET_COMBO_VIEW(ByRef cmbComboBox As System.Windows.Forms.ComboBox, ByVal lngCODE As Long)
        Dim intLoopIndex As Integer
        Dim lngTemp_CODE As Long
        Dim intSetIndex As Integer

        With cmbComboBox
            .SelectedIndex = -1
            intSetIndex = -1

            If .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown And .Text <> "" Then 'インデックスを回してもテキストが残った場合
                .Text = ""
            End If

            If .Items.Count = 0 Then
                Exit Sub
            End If

            For intLoopIndex = 0 To .Items.Count - 1
                lngTemp_CODE = CLng(.Items.Item(intLoopIndex).strGComboBoxData(0))
                If lngCODE = lngTemp_CODE Then
                    intSetIndex = intLoopIndex
                End If
            Next

            .SelectedIndex = intSetIndex
        End With
    End Sub

    'すでにアイテムが展開されているコンボボックスと同様の内容をほかのコンボボックスに展開する
    Public Sub SUB_COPY_COMBO_ITEM(ByRef cmbSOURCE_COMBO As System.Windows.Forms.ComboBox, ByRef cmbDESTINATION_COMBO As System.Windows.Forms.ComboBox)
        Dim intLoopIndex As Integer

        Call SUB_END_COMBO_VIEW(cmbDESTINATION_COMBO)

        With cmbDESTINATION_COMBO
            .DropDownWidth = cmbSOURCE_COMBO.DropDownWidth
            .DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
            .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown

            For intLoopIndex = 0 To cmbSOURCE_COMBO.Items.Count - 1
                .Items.Add(cmbSOURCE_COMBO.Items(intLoopIndex))
            Next

            AddHandler .DrawItem, AddressOf glbSubComboBox_DrawItem
        End With

    End Sub

#End Region

#Region "名称のみのコンボに関する共通部"

    Public Sub SUB_END_COMBO_KIND(ByRef cmbComboBox As Object)
        'クリア部
        With cmbComboBox
            .SelectedIndex = -1
            .Text = ""
            .Items.Clear()
        End With
    End Sub

    Public Function FUNC_GET_COMBO_KIND_CODE(ByRef cmbComboBox As Object) As Long
        Dim intIndex As Integer
        Dim lngRet As Long
        Dim strTemp As String

        With cmbComboBox
            'If .Text = "" Then
            '    .SelectedIndex = -1
            'End If

            intIndex = .SelectedIndex
            strTemp = ""

            If .Items.Count = 0 Then
                Return cstComboBoxGetMiss_NUM
            End If

            If Not TypeOf .Items.Item(0) Is clsGComboBoxData Then
                Return cstComboBoxGetMiss_NUM
            End If

            If intIndex > -1 Then
                Try
                    strTemp = .Items.Item(intIndex).strGComboBoxData(0)
                Catch ex As Exception
                    strTemp = ""
                End Try
            End If

            If IsNumeric(strTemp) Then
                lngRet = CLng(strTemp)
            Else
                lngRet = cstComboBoxGetMiss_NUM
            End If

            If lngRet = cstComboBoxGetMiss_NUM Then
                strTemp = .Text

                If IsNumeric(strTemp) Then
                    lngRet = CLng(strTemp)
                Else
                    lngRet = cstComboBoxGetMiss_NUM
                End If

            End If

            Return lngRet
        End With
    End Function

    Public Sub SUB_SET_COMBO_KIND_CODE(ByRef cmbComboBox As Object, ByVal lngCODE As Long)
        Dim intLoopIndex As Integer
        Dim lngTemp_CODE As Long
        Dim intSetIndex As Integer

        With cmbComboBox
            .SelectedIndex = -1
            intSetIndex = -1

            If .Items.Count = 0 Then
                Exit Sub
            End If

            For intLoopIndex = 0 To .Items.Count - 1
                lngTemp_CODE = CInt(.Items.Item(intLoopIndex).strGComboBoxData(0))
                If lngCODE = lngTemp_CODE Then
                    intSetIndex = intLoopIndex
                End If
            Next

            .SelectedIndex = intSetIndex
        End With
    End Sub

    Public Sub SUB_SET_COMBO_KIND_CODE_FIRST(ByRef cmbComboBox As Object)

        With cmbComboBox

            If .Items.Count <= 0 Then
                Exit Sub
            End If

            If .SelectedIndex = 0 Then
                Exit Sub
            End If

            .SelectedIndex = -1

            .SelectedIndex = 0
            Call System.Windows.Forms.Application.DoEvents()

        End With
    End Sub

    Public Sub SUB_SET_COMBO_KIND_CODE_LAST(ByRef cmbComboBox As Object)

        With cmbComboBox

            If .Items.Count <= 0 Then
                Exit Sub
            End If

            .SelectedIndex = -1

            .SelectedIndex = (.Items.Count - 1)
            Call System.Windows.Forms.Application.DoEvents()

        End With
    End Sub
#End Region

#Region "キャッシュ"

#Region "アイテムの転写"
    Private Sub SUB_COPY_ITEM( _
    ByRef srtITEM_BASE() As srtLComboDataArray, ByRef srtITEM_DEST() As srtLComboDataArray _
    )
        Dim intINDEX As Integer
        Dim intLOOP_INDEX As Integer

        ReDim srtITEM_DEST(-1)
        srtITEM_DEST = Nothing
        If IsNothing(srtITEM_BASE) Then
            Exit Sub
        End If

        intINDEX = (srtITEM_BASE.Length - 1)
        ReDim srtITEM_DEST(intINDEX)

        For intLOOP_INDEX = 0 To (srtITEM_BASE.Length - 1)
            srtITEM_DEST(intLOOP_INDEX) = srtITEM_BASE(intLOOP_INDEX)
        Next

    End Sub
#End Region

#Region "INT_ITEM"

    Public Structure SRT_CASH_INT_ITEM
        Public KEY01 As Integer
        Public ITEM() As srtLComboDataArray
    End Structure

    Public Function FUNC_SEARCH_CASH_INT_ITEM( _
    ByRef srtSEARCH() As SRT_CASH_INT_ITEM, _
    ByVal intKEY01 As Integer _
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

    Public Sub SUB_ADD_CASH_INT_ITEM( _
    ByRef srtCASH() As SRT_CASH_INT_ITEM, _
    ByVal intKEY01 As Integer, ByVal srtITEM() As srtLComboDataArray _
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_ITEM(srtCASH, intKEY01) <> -1 Then 'すでに存在するなら
            Exit Sub '追加しない
        End If

        If IsNothing(srtCASH) Then
            intINDEX = 0
        Else
            intINDEX = UBound(srtCASH) + 1
        End If

        ReDim Preserve srtCASH(intINDEX)
        With srtCASH(intINDEX)
            .KEY01 = intKEY01
            Call SUB_COPY_ITEM(srtITEM, .ITEM)
        End With

    End Sub
#End Region

#End Region

End Module
