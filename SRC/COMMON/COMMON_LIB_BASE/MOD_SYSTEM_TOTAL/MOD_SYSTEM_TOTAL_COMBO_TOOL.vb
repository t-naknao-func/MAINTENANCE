Public Module MOD_SYSTEM_TOTAL_COMBO_TOOL

#Region "MNG_M_KIND"
    Public Sub SUB_SYSTEM_COMMBO_MNG_M_KIND( _
    ByRef cmbCOMBO_BOX As Object, _
    ByVal enmCODE_FLAG As ENM_MNG_M_KIND_CODE_FLAG, _
    Optional ByVal blnVIEW_NULL As Boolean = False, _
    Optional ByVal strNAME_ALL As String = "", _
    Optional ByVal strCODE_KIND_IN As String = ""
    )
        Dim strSQL As System.Text.StringBuilder
        Dim sdrREADER As SqlClient.SqlDataReader 'データリーダー
        Dim datCOMBO_ITEM As clsGComboBoxData
        Dim strComboString() As srtLComboDataArray
        Dim intLOOP_INDEX As Integer

        Call SUB_END_COMBO_KIND(cmbCOMBO_BOX) '渡されたコンボを初期化(データリンクなどの破棄)

        strSQL = New System.Text.StringBuilder
        With strSQL
            Call .Append("SELECT" & Environment.NewLine)
            Call .Append("CODE_KIND AS CODE," & Environment.NewLine)
            Call .Append("NAME_KIND AS NAME" & Environment.NewLine)
            Call .Append("FROM" & Environment.NewLine)
            Call .Append(strSYSTEM_PUBLIC_MNGDB_PREFIX & "MNG_M_KIND" & Environment.NewLine)
            Call .Append("WHERE" & Environment.NewLine)
            Call .Append("CODE_FLAG=" & enmCODE_FLAG & Environment.NewLine)
            If Not (strCODE_KIND_IN = "") Then
                Call .Append("AND CODE_KIND IN(" & strCODE_KIND_IN & ")" & Environment.NewLine)
            End If
            Call .Append("ORDER BY" & Environment.NewLine)
            Call .Append("CODE_FLAG,CODE_KIND" & Environment.NewLine)
        End With
        sdrREADER = Nothing '初期化

        If Not FUNC_SYSTEM_GET_SQL_DATA_READER(strSQL.ToString, sdrREADER) Then
            sdrREADER = Nothing
            Exit Sub
        End If

        If Not sdrREADER.HasRows Then
            Call sdrREADER.Close()
            sdrREADER = Nothing
            Exit Sub
        End If

        ReDim strComboString(-1)
        intLOOP_INDEX = 0

        If blnVIEW_NULL Then
            ReDim strComboString(0)
            intLOOP_INDEX += 1
            ReDim Preserve strComboString(intLOOP_INDEX)
            ReDim strComboString(intLOOP_INDEX).strLOneRecord(1)
            With strComboString(intLOOP_INDEX)
                .strLOneRecord(0) = -1
                .strLOneRecord(1) = strNAME_ALL
            End With
        End If

        Do While sdrREADER.Read
            intLOOP_INDEX += 1
            ReDim Preserve strComboString(intLOOP_INDEX)
            ReDim strComboString(intLOOP_INDEX).strLOneRecord(1)
            With strComboString(intLOOP_INDEX)
                .strLOneRecord(0) = CInt(sdrREADER.Item("CODE"))
                .strLOneRecord(1) = CStr(sdrREADER.Item("NAME"))
            End With
        Loop

        Call sdrREADER.Close()
        sdrREADER = Nothing

        For intLOOP_INDEX = 1 To UBound(strComboString)
            datCOMBO_ITEM = New clsGComboBoxData(strComboString(intLOOP_INDEX).strLOneRecord, Nothing, 1)
            cmbCOMBO_BOX.Items.Add(datCOMBO_ITEM)
            datCOMBO_ITEM = Nothing
        Next

        ReDim strComboString(0)
        Erase strComboString
    End Sub
#End Region

End Module
