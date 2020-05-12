Public Module MOD_DATA_TABLE_TOOL

    'データテーブルの作成
    Public Sub glbSubMakeDataTable(ByRef tblHBase As DataTable, ByVal strHColumnString As String, ByVal strHTypeString As String)
        Dim objWGuideColumn As Object
        Dim intWLoopIndex As Integer
        Dim colWTblColumn As DataColumn

        tblHBase = New DataTable

		objWGuideColumn = Split(strHColumnString, ",")

        For intWLoopIndex = LBound(objWGuideColumn) To UBound(objWGuideColumn)
            colWTblColumn = New DataColumn
            colWTblColumn.ColumnName = objWGuideColumn(intWLoopIndex)
            colWTblColumn.DataType = funcConvStrToType(Mid(strHTypeString, intWLoopIndex + 1, 1))
            tblHBase.Columns.Add(colWTblColumn)
        Next

    End Sub

    'データテーブルの作成(別のデータテーブルから定義をコピー)
    Public Sub glbSubMakeDataTableCopy(ByRef tblHBase As DataTable, ByRef tblSource As DataTable)
        Dim intWLoopIndex As Integer
        Dim colWTblColumn As DataColumn

        tblHBase = New DataTable

        
        For intWLoopIndex = 0 To (tblSource.Columns.Count - 1)
            colWTblColumn = New DataColumn
            colWTblColumn.ColumnName = tblSource.Columns(intWLoopIndex).ColumnName
            colWTblColumn.DataType = tblSource.Columns(intWLoopIndex).DataType
            tblHBase.Columns.Add(colWTblColumn)
        Next

    End Sub

    '列型の変換
    Private Function funcConvStrToType(ByVal chrHTypeChar As Char) As Type
        Select Case chrHTypeChar
            Case "L"
                Return System.Type.GetType("System.Int64")
            Case "S"
                Return System.Type.GetType("System.String")
            Case "C"
                Return System.Type.GetType("System.Decimal")
            Case "D"
                Return System.Type.GetType("System.DateTime")
            Case "I"
                Return System.Type.GetType("System.Int32")
            Case "Z"
                Return System.Type.GetType("System.Double")
            Case "B"
                Return System.Type.GetType("System.Boolean")
            Case Else
                Return System.Type.GetType("System.String")
        End Select
    End Function

    'データテーブルに行を追加
    Public Sub glbSubAddRowDataTable(ByRef tblHBase As DataTable, ByRef objHRow() As Object)
        Call tblHBase.Rows.Add(objHRow)
    End Sub

    'データテーブルに行を追加
    Public Sub SUB_ADD_ROW_DATA_TABLE( _
    ByRef tblHBase As DataTable, ByRef objHRow() As Object _
    )
        Dim dtrROW As DataRow

        dtrROW = tblHBase.NewRow()

        dtrROW.ItemArray = objHRow

        Call tblHBase.Rows.Add(dtrROW)
        dtrROW = Nothing
    End Sub

    'データテーブルからデータテーブルへ値のコピー
    Public Sub SUB_COPY_ROW_DATA_TABLE( _
    ByRef tblSOURCE As DataTable, ByRef tblDEST As DataTable _
    )
        Dim intLOOP_INDEX As Integer

        Call tblDEST.Clear()

        For intLOOP_INDEX = 0 To tblSOURCE.Rows.Count - 1
            Call tblDEST.Rows.Add(tblSOURCE.Rows(intLOOP_INDEX).ItemArray)
        Next

    End Sub

    'データテーブル行の値をクリア
    Public Sub SUB_DATAROW_CLEAR( _
    ByRef dtrCLEAR As DataRow _
    )
        Dim intLOOP_INDEX As Integer

        For intLOOP_INDEX = 0 To (dtrCLEAR.Table.Columns.Count - 1)
            dtrCLEAR.Item(intLOOP_INDEX) = DBNull.Value
        Next

    End Sub

    'データテーブル行の値をコピー
    Public Sub SUB_DATAROW_COPY( _
    ByRef dtrSOURCE As DataRow, ByRef dtrDEST As DataRow _
    )
        Dim intLOOP_INDEX As Integer

        For intLOOP_INDEX = 0 To (dtrSOURCE.Table.Columns.Count - 1)
            dtrDEST.Item(intLOOP_INDEX) = dtrSOURCE.Item(intLOOP_INDEX)
        Next

    End Sub

    'DataTableの中身を改行付き任意の区切文字列へ変換
	Public Function FUNC_GET_STR_FROM_DATA_TABLE( _
	ByRef dtTABLE As DataTable, _
	Optional ByVal strPUT_INDEX As String = "", _
	Optional ByVal blnAPPEND_HEADER As Boolean = True, _
	Optional ByVal strSEP As String = "" _
	) As String
		Dim strRET As String
		Dim intLOOP_INDEX As Integer
		Dim intPUT_INDEX() As Integer
		strRET = ""

		If strSEP = "" Then
			strSEP = Convert.ToChar(9)
		End If

		If IsNothing(dtTABLE) Then
			Return ""
		End If

		If blnAPPEND_HEADER Then
			For intLOOP_INDEX = 0 To (dtTABLE.Columns.Count - 1)
				intPUT_INDEX = FUNC_SPLIT_STR_CONV_INT_ROW(strPUT_INDEX)
				If strPUT_INDEX = "" Or FUNC_CHECK_INT_ROW(intPUT_INDEX, intLOOP_INDEX) Then
					strRET &= dtTABLE.Columns(intLOOP_INDEX).ColumnName & If(intLOOP_INDEX = (dtTABLE.Columns.Count - 1), "", strSEP)
				End If
			Next
			strRET &= Environment.NewLine
		End If

		For intLOOP_INDEX = 0 To (dtTABLE.Rows.Count - 1)
			strRET &= FUNC_GET_STR_FROM_DATA_ROW(dtTABLE.Rows(intLOOP_INDEX), strPUT_INDEX, strSEP) & Environment.NewLine
		Next

		Return strRET
	End Function

	'DataRowを任意の区切文字列へ変換
	Public Function FUNC_GET_STR_FROM_DATA_ROW( _
	ByRef dtrROW As DataRow, _
	Optional ByVal strPUT_INDEX As String = "", _
	Optional ByVal strSEP As String = "" _
	) As String
		Dim strRET As String
		Dim intLOOP_INDEX As Integer
		Dim intPUT_INDEX() As Integer

		If strSEP = "" Then
			strSEP = Convert.ToChar(9)
		End If

		strRET = ""
		For intLOOP_INDEX = 0 To (dtrROW.Table.Columns.Count - 1)
			intPUT_INDEX = FUNC_SPLIT_STR_CONV_INT_ROW(strPUT_INDEX)
			If strPUT_INDEX = "" Or FUNC_CHECK_INT_ROW(intPUT_INDEX, intLOOP_INDEX) Then
				strRET &= dtrROW.Item(intLOOP_INDEX) & If(intLOOP_INDEX = (dtrROW.Table.Columns.Count - 1), "", strSEP)
			End If
		Next

		Return strRET
	End Function
End Module
