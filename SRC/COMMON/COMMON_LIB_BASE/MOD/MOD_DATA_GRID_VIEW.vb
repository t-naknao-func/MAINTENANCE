Public Module MOD_DATA_GRID_VIEW

    '文字数でのラッパー
    Public Sub SUB_DGV_COLUMN_WIDTH_INIT_COUNT_FONT( _
    ByRef dgvCONTROL As System.Windows.Forms.DataGridView, _
    ByVal strCOLUMN_COUNT As String, _
    ByVal strALIGNMENT As String, _
    Optional ByVal strSORT_ENABLED As String = "" _
    )
        Dim sngFONT_SIZE As Single
        Dim sngBASE_POINT As Single
        Dim objCOUNT As Object
        Dim intLOOP_INDEX As Integer
        Dim intCOUNT As Integer
        Dim strWIDTH_ROW() As String
        Dim strWIDTHS As String

        sngFONT_SIZE = dgvCONTROL.Font.Size
        sngBASE_POINT = (sngFONT_SIZE + (sngFONT_SIZE / 9)) * 2 '全角としてとりあえず2倍
        sngBASE_POINT *= 1.05
        sngBASE_POINT = CInt(sngBASE_POINT)

        'objCOUNT = Split(strCOLUMN_COUNT, ",")
        'ReDim strWIDTH_ROW(UBound(objCOUNT))
        'For intLOOP_INDEX = LBound(objCOUNT) To UBound(objCOUNT)
        '    intCOUNT = CInt(objCOUNT(intLOOP_INDEX))
        '    strWIDTH_ROW(intLOOP_INDEX) = CStr(sngBASE_POINT * intCOUNT)
        'Next

        objCOUNT = strCOLUMN_COUNT.Split(",")
        ReDim strWIDTH_ROW(objCOUNT.Length)
        For intLOOP_INDEX = 0 To (objCOUNT.Length - 1)
            intCOUNT = CInt(objCOUNT(intLOOP_INDEX))
            strWIDTH_ROW(intLOOP_INDEX + 1) = CStr(sngBASE_POINT * intCOUNT)
        Next

        strWIDTHS = FUNC_CONVERT_STRING_ROW_TO_STRING(strWIDTH_ROW)

        Call SUB_DGV_COLUMN_WIDTH_INIT(dgvCONTROL, strWIDTHS, strALIGNMENT, strSORT_ENABLED)

    End Sub

    Public Sub SUB_DGV_COLUMN_WIDTH_INIT( _
    ByRef dgvCONTROL As System.Windows.Forms.DataGridView, _
    ByVal strCOLUMN_WIDTH As String, _
    ByVal strALIGNMENT As String, _
    Optional ByVal strSORT_ENABLED As String = "" _
    )
        Dim objWIDTH As Object
        Dim intLOOP_INDEX As Integer
        Dim strCHAR As String

        objWIDTH = Split(strCOLUMN_WIDTH, ",")

        For intLOOP_INDEX = LBound(objWIDTH) To UBound(objWIDTH)
            Try
                dgvCONTROL.Columns(intLOOP_INDEX).Width = CInt(objWIDTH(intLOOP_INDEX))

                If strSORT_ENABLED = "" Then
                    dgvCONTROL.Columns(intLOOP_INDEX).SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
                Else
                    strCHAR = Mid(strSORT_ENABLED, intLOOP_INDEX + 1, 1)
                    dgvCONTROL.Columns(intLOOP_INDEX).SortMode = FUNC_GET_SORT_MODE(strCHAR)
                End If

                strCHAR = Mid(strALIGNMENT, intLOOP_INDEX + 1, 1)
                dgvCONTROL.Columns(intLOOP_INDEX).DefaultCellStyle.Alignment = FUNC_GET_ALIGNMENT_VALUE(strCHAR)

            Catch ex As Exception
                Exit Sub
            End Try
        Next

    End Sub

    Public Sub SUB_DGV_COLUMN_WIDTH_ADJUST(ByRef dgvCONTROL As System.Windows.Forms.DataGridView)
        Dim intLOOP_INDEX As Integer
        Dim intLAST_INDEX As Integer
        Dim intWIDTH_COL As Integer
        Dim intWIDTH_ADJUST As Integer
        Const cstWIDTH_LINE As Integer = 1
        Dim intWIDTH_SCROLL As Integer 

        intWIDTH_SCROLL = System.Windows.Forms.SystemInformation.HorizontalScrollBarArrowWidth
        intLAST_INDEX = dgvCONTROL.Columns.Count - 1
        intWIDTH_COL = 0
        For intLOOP_INDEX = 0 To intLAST_INDEX
            intWIDTH_COL += dgvCONTROL.Columns(intLOOP_INDEX).Width + cstWIDTH_LINE '+境界線分
        Next

        intWIDTH_ADJUST = (dgvCONTROL.Width - intWIDTH_COL)
        intWIDTH_ADJUST -= intWIDTH_SCROLL 'スクロールバー分を除く

        If intWIDTH_ADJUST < 0 Then 'オーバーした場合（定義がはみ出ている)は調整しない
            Exit Sub
        End If

        dgvCONTROL.Columns(intLAST_INDEX).Width += intWIDTH_ADJUST
    End Sub

    '行の高さ調整。フォントサイズに対する自動調整。(独自デザインの場合は呼び出さない)
    Public Sub SUB_DGV_ROW_HEIGHT_ADJUST(ByRef dgvCONTROL As System.Windows.Forms.DataGridView)
        'dgvCONTROL.RowTemplate.Height = 20
    End Sub

    Private Function FUNC_GET_SORT_MODE(ByVal strMODE As String) As System.Windows.Forms.DataGridViewColumnSortMode
        Dim enmRET As System.Windows.Forms.DataGridViewColumnSortMode

        If strMODE = "Y" Then
            enmRET = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Else
            enmRET = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        End If

        Return enmRET
    End Function

    Public Sub SUB_DGV_COLUMN_WIDTH_INIT_COPY( _
    ByRef dgvCONTROL As System.Windows.Forms.DataGridView, _
    ByRef dgvCONTROL_SOURCE As System.Windows.Forms.DataGridView _
    )
        Dim intLOOP_INDEX As Integer

        For intLOOP_INDEX = 0 To (dgvCONTROL_SOURCE.Columns.Count - 1)
            Try
                dgvCONTROL.Columns(intLOOP_INDEX).Width = dgvCONTROL_SOURCE.Columns(intLOOP_INDEX).Width
                dgvCONTROL.Columns(intLOOP_INDEX).SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
                dgvCONTROL.Columns(intLOOP_INDEX).DefaultCellStyle.Alignment = dgvCONTROL_SOURCE.Columns(intLOOP_INDEX).DefaultCellStyle.Alignment
            Catch ex As Exception
                Exit Sub
            End Try
        Next

    End Sub

    Private Function FUNC_GET_ALIGNMENT_VALUE(ByVal strCHAR As String) As System.Windows.Forms.DataGridViewContentAlignment
        Dim enmRET As System.Windows.Forms.DataGridViewContentAlignment

        Select Case strCHAR
            Case "R"
                enmRET = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
            Case "L"
                enmRET = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            Case "C"
                enmRET = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case Else
                enmRET = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        End Select
        Return enmRET
    End Function

    Public Function FUNC_GET_SELECT_ROW_INDEX(ByRef dgvCONTROL As System.Windows.Forms.DataGridView) As Integer
        Dim tblTEMP As DataTable
        Dim intRET As Integer
        Dim dgvROW As System.Windows.Forms.DataGridViewRow
        Dim dgvCELL As System.Windows.Forms.DataGridViewCell

        tblTEMP = dgvCONTROL.DataSource

        If IsNothing(tblTEMP) Then
            Return -1
        End If

        If tblTEMP.Rows.Count <= 0 Then
            Return 0
        End If

        Select Case dgvCONTROL.SelectionMode
            Case Windows.Forms.DataGridViewSelectionMode.FullRowSelect
                If dgvCONTROL.SelectedRows.Count > 0 Then
                    dgvROW = dgvCONTROL.SelectedRows(0)
                    intRET = dgvROW.Cells(0).RowIndex
                Else
                    intRET = -1
                End If
            Case Windows.Forms.DataGridViewSelectionMode.CellSelect
                dgvCELL = dgvCONTROL.SelectedCells(0)
                intRET = dgvCELL.RowIndex
            Case Windows.Forms.DataGridViewSelectionMode.RowHeaderSelect
                dgvCELL = dgvCONTROL.SelectedCells(0)
                intRET = dgvCELL.RowIndex
            Case Else
                intRET = 0
        End Select

        intRET += 1
        Return intRET
    End Function

    Public Sub SUB_SET_SELECT_ROW_INDEX(ByRef dgvCONTROL As System.Windows.Forms.DataGridView, ByVal intSELECT As Integer)
        Dim tblTEMP As DataTable

        tblTEMP = dgvCONTROL.DataSource

        If IsNothing(tblTEMP) Then
            Exit Sub
        End If

        If tblTEMP.Rows.Count <= 0 Then
            Exit Sub
        End If

        If intSELECT <= 0 Then
            dgvCONTROL.CurrentCell = Nothing
        Else
            dgvCONTROL.CurrentCell = dgvCONTROL.Rows(intSELECT - 1).Cells(0)
        End If
        'SelectionModeプロパティが "RowHeaderSelect" なら、
        'dgvCONTROL.Rows(1).Selected = True
        Call System.Windows.Forms.Application.DoEvents()

    End Sub

    Public Sub SUB_DATA_GRID_SORT_CLEAR(ByRef dgvCONTROL As System.Windows.Forms.DataGridView)
        Dim tblTEMP As DataTable
        Dim davTEMP As DataView

        tblTEMP = dgvCONTROL.DataSource

        If IsNothing(tblTEMP) Then
            Exit Sub
        End If

        davTEMP = tblTEMP.DefaultView 'テーブルからビューを読み取る
        davTEMP.Sort = "" 'ソート条件に空（Null）を代入
        Call System.Windows.Forms.Application.DoEvents()
        'dgvCONTROL.HeaderCell.SortGlyphDirection = Windows.Forms.SortOrder.None 'ソートを解除
    End Sub

    'グリッド単位の読取専用をセル単位に切り替える
    'グリッド内のデータが表示されてから呼び出す事
    'また、明細の件数に増減があった場合は、呼び直す事
    Public Sub SUB_DATA_GRID_SET_READ_ONLY_MODE_CELL(ByRef dgvCONTROL As System.Windows.Forms.DataGridView)
        Dim intMAX_INDEX_ROW As Integer '行数
        Dim intMAX_INDEX_COL As Integer '列数

        If dgvCONTROL.DataSource Is Nothing Then
            Exit Sub
        End If

        dgvCONTROL.ReadOnly = False

        intMAX_INDEX_ROW = (dgvCONTROL.RowCount - 1)
        intMAX_INDEX_COL = (dgvCONTROL.ColumnCount - 1)

        For intROW_INDEX = 0 To intMAX_INDEX_ROW
            For intCOL_INDEX = 0 To intMAX_INDEX_COL
                dgvCONTROL.Rows(intROW_INDEX).Cells(intCOL_INDEX).ReadOnly = True '
            Next
        Next

    End Sub

    'グリッドの特定の列を入力可能とする
    'グリッド内のデータが表示されてから呼び出す事
    'また、明細の件数に増減があった場合は、呼び直す事
    Public Sub SUB_DATA_GRID_CELL_READ_ONLY_MODE(ByRef dgvCONTROL As System.Windows.Forms.DataGridView, intEDIT_COL_INDEX As Integer, ByVal blnREAD_ONLY As Boolean)
        Dim intMAX_INDEX_ROW As Integer '行数
       
        If dgvCONTROL.DataSource Is Nothing Then
            Exit Sub
        End If

        intMAX_INDEX_ROW = (dgvCONTROL.RowCount - 1)

        For intROW_INDEX = 0 To intMAX_INDEX_ROW
            dgvCONTROL.Rows(intROW_INDEX).Cells(intEDIT_COL_INDEX).ReadOnly = blnREAD_ONLY
        Next
    End Sub

    'グリッドのチラツキを抑えるため、選択モードを変更して、リフレッシュ
    Public Sub SUB_DATA_GRID_REFRESH_CHG_SELECTION_MODE(ByRef dgvCONTROL As System.Windows.Forms.DataGridView)
        Dim intSELECTION_MODE_BEFORE As Integer

        intSELECTION_MODE_BEFORE = dgvCONTROL.SelectionMode 'モードを保持
        dgvCONTROL.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect 'チラツキを抑えるため、セル選択モードへ変更
        Call dgvCONTROL.Refresh() 'グリッドの表示を更新
        dgvCONTROL.SelectionMode = intSELECTION_MODE_BEFORE '元のモードへ戻す

    End Sub
End Module
