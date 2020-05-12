Public Module MOD_CASH_TOOL

#Region "キャッシュ用構造体の定義"
    Public Structure SRT_CASH_INT_DEC
        Public KEY01 As Integer
        Public VALUE As Decimal
    End Structure

    Public Structure SRT_CASH_INT_INT
        Public KEY01 As Integer
        Public VALUE As Integer
    End Structure

    Public Structure SRT_CASH_INT_BOOL
        Public KEY01 As Integer
        Public VALUE As Boolean
    End Structure

    Public Structure SRT_CASH_INT_STR
        Public KEY01 As Integer
        Public VALUE As String
    End Structure

    Public Structure SRT_CASH_INT_DATE
        Public KEY01 As Integer
        Public VALUE As DateTime
    End Structure

    Public Structure SRT_CASH_INT_INT_STR
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public VALUE As String
    End Structure

    Public Structure SRT_CASH_INT_INT_BOOL
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public VALUE As Boolean
    End Structure

    Public Structure SRT_CASH_INT_INT_INT
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public VALUE As Integer
    End Structure

    Public Structure SRT_CASH_INT_INT_DATE
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public VALUE As DateTime
    End Structure

    Public Structure SRT_CASH_INT_INT_DEC
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public VALUE As Decimal
    End Structure

    Public Structure SRT_CASH_INT_INT_INT_STR
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public KEY03 As Integer
        Public VALUE As String
    End Structure

    Public Structure SRT_CASH_INT_INT_INT_BOOL
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public KEY03 As Integer
        Public VALUE As Boolean
    End Structure

    Public Structure SRT_CASH_INT_INT_INT_INT
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public KEY03 As Integer
        Public VALUE As Integer
    End Structure

    Public Structure SRT_CASH_INT_INT_INT_LONG
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public KEY03 As Integer
        Public VALUE As Long
    End Structure

    Public Structure SRT_CASH_INT_INT_INT_INT_STR
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public KEY03 As Integer
        Public KEY04 As Integer
        Public VALUE As String
    End Structure

    Public Structure SRT_CASH_INT_INT_INT_INT_BOOL
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public KEY03 As Integer
        Public KEY04 As Integer
        Public VALUE As Boolean
    End Structure

    Public Structure SRT_CASH_INT_INT_INT_INT_INT_STR
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public KEY03 As Integer
        Public KEY04 As Integer
        Public KEY05 As Integer
        Public VALUE As String
    End Structure

    Public Structure SRT_CASH_INT_INT_INT_INT_INT_BOOL
        Public KEY01 As Integer
        Public KEY02 As Integer
        Public KEY03 As Integer
        Public KEY04 As Integer
        Public KEY05 As Integer
        Public VALUE As Boolean
    End Structure
#End Region

#Region "INT_BOOL"
    Public Sub SUB_ADD_CASH_INT_BOOL(
    ByRef srtCASH() As SRT_CASH_INT_BOOL,
    ByVal intKEY01 As Integer, ByVal blnVALUE As Boolean
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_BOOL(srtCASH, intKEY01) <> -1 Then 'すでに存在するなら
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
            .VALUE = blnVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_BOOL(
    ByRef srtSEARCH() As SRT_CASH_INT_BOOL,
    ByVal intKEY01 As Integer
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
#End Region

#Region "INT_INT"
    Public Sub SUB_ADD_CASH_INT_INT(
    ByRef srtCASH() As SRT_CASH_INT_INT,
    ByVal intKEY01 As Integer, ByVal intVALUE As Integer
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT(srtCASH, intKEY01) <> -1 Then 'すでに存在するなら
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
            .VALUE = intVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT(
    ByRef srtSEARCH() As SRT_CASH_INT_INT,
    ByVal intKEY01 As Integer
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
#End Region

#Region "INT_DEC"
    Public Sub SUB_ADD_CASH_INT_DEC(
    ByRef srtCASH() As SRT_CASH_INT_DEC,
    ByVal intKEY01 As Integer, ByVal decVALUE As Decimal
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_DEC(srtCASH, intKEY01) <> -1 Then 'すでに存在するなら
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
            .VALUE = decVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_DEC(
    ByRef srtSEARCH() As SRT_CASH_INT_DEC,
    ByVal intKEY01 As Integer
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
#End Region

#Region "INT_STR"
    Public Sub SUB_ADD_CASH_INT_STR(
    ByRef srtCASH() As SRT_CASH_INT_STR,
    ByVal intKEY01 As Integer, ByVal strVALUE As String
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_STR(srtCASH, intKEY01) <> -1 Then 'すでに存在するなら
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
            .VALUE = strVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_STR(
    ByRef srtSEARCH() As SRT_CASH_INT_STR,
    ByVal intKEY01 As Integer
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
#End Region

#Region "INT_DATE"
    Public Sub SUB_ADD_CASH_INT_DATE(
    ByRef srtCASH() As SRT_CASH_INT_DATE,
    ByVal intKEY01 As Integer, ByVal datVALUE As DateTime
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_DATE(srtCASH, intKEY01) <> -1 Then 'すでに存在するなら
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
            .VALUE = datVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_DATE(
    ByRef srtSEARCH() As SRT_CASH_INT_DATE,
    ByVal intKEY01 As Integer
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
#End Region

#Region "INT_INT_STR"
    Public Sub SUB_ADD_CASH_INT_INT_STR(
    ByRef srtCASH() As SRT_CASH_INT_INT_STR,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal strVALUE As String
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_STR(srtCASH, intKEY01, intKEY02) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .VALUE = strVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_STR(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_STR,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_BOOL"
    Public Sub SUB_ADD_CASH_INT_INT_BOOL(
    ByRef srtCASH() As SRT_CASH_INT_INT_BOOL,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal blnVALUE As Boolean
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_BOOL(srtCASH, intKEY01, intKEY02) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .VALUE = blnVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_BOOL(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_BOOL,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT"
    Public Sub SUB_ADD_CASH_INT_INT_INT(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intVALUE As Integer
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT(srtCASH, intKEY01, intKEY02) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .VALUE = intVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_DATE"
    Public Sub SUB_ADD_CASH_INT_INT_DATE(
    ByRef srtCASH() As SRT_CASH_INT_INT_DATE,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal datVALUE As DateTime
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_DATE(srtCASH, intKEY01, intKEY02) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .VALUE = datVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_DATE(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_DATE,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_DEC"
    Public Sub SUB_ADD_CASH_INT_INT_DEC(
    ByRef srtCASH() As SRT_CASH_INT_INT_DEC,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal decVALUE As Decimal
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_DEC(srtCASH, intKEY01, intKEY02) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .VALUE = decVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_DEC(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_DEC,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT_STR"
    Public Sub SUB_ADD_CASH_INT_INT_INT_STR(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT_STR,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intKEY03 As Integer, ByVal strVALUE As String
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT_STR(srtCASH, intKEY01, intKEY02, intKEY03) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .KEY03 = intKEY03
            .VALUE = strVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT_STR(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT_STR,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer,
    ByVal intKEY03 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 _
                And .KEY03 = intKEY03 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT_BOOL"
    Public Sub SUB_ADD_CASH_INT_INT_INT_BOOL(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT_BOOL,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intKEY03 As Integer, ByVal blnVALUE As Boolean
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT_BOOL(srtCASH, intKEY01, intKEY02, intKEY03) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .KEY03 = intKEY03
            .VALUE = blnVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT_BOOL(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT_BOOL,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer,
    ByVal intKEY03 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 _
                And .KEY03 = intKEY03 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT_INT"
    Public Sub SUB_ADD_CASH_INT_INT_INT_INT(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT_INT,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intKEY03 As Integer, ByVal intVALUE As Integer
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT_INT(srtCASH, intKEY01, intKEY02, intKEY03) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .KEY03 = intKEY03
            .VALUE = intVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT_INT(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT_INT,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer,
    ByVal intKEY03 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 _
                And .KEY03 = intKEY03 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT_LONG"
    Public Sub SUB_ADD_CASH_INT_INT_INT_LONG(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT_LONG,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intKEY03 As Integer, ByVal lngVALUE As Long
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT_LONG(srtCASH, intKEY01, intKEY02, intKEY03) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .KEY03 = intKEY03
            .VALUE = lngVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT_LONG(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT_LONG,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer,
    ByVal intKEY03 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 _
                And .KEY03 = intKEY03 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT_INT_STR"
    Public Sub SUB_ADD_CASH_INT_INT_INT_INT_STR(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT_INT_STR,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intKEY03 As Integer, ByVal intKEY04 As Integer, ByVal strVALUE As String
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT_INT_STR(srtCASH, intKEY01, intKEY02, intKEY03, intKEY04) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .KEY03 = intKEY03
            .KEY04 = intKEY04
            .VALUE = strVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT_INT_STR(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT_INT_STR,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer,
    ByVal intKEY03 As Integer,
    ByVal intKEY04 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 _
                And .KEY03 = intKEY03 _
                And .KEY04 = intKEY04 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT_INT_BOOL"
    Public Sub SUB_ADD_CASH_INT_INT_INT_INT_BOOL(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT_INT_BOOL,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intKEY03 As Integer, ByVal intKEY04 As Integer, ByVal blnVALUE As Boolean
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT_INT_BOOL(srtCASH, intKEY01, intKEY02, intKEY03, intKEY04) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .KEY03 = intKEY03
            .KEY04 = intKEY04
            .VALUE = blnVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT_INT_BOOL(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT_INT_BOOL,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer,
    ByVal intKEY03 As Integer,
    ByVal intKEY04 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 _
                And .KEY03 = intKEY03 _
                And .KEY04 = intKEY04 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT_INT_INT_STR"
    Public Sub SUB_ADD_CASH_INT_INT_INT_INT_INT_STR(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT_INT_INT_STR,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intKEY03 As Integer, ByVal intKEY04 As Integer, ByVal intKEY05 As Integer, ByVal strVALUE As String
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT_INT_INT_STR(srtCASH, intKEY01, intKEY02, intKEY03, intKEY04, intKEY05) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .KEY03 = intKEY03
            .KEY04 = intKEY04
            .KEY05 = intKEY05
            .VALUE = strVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT_INT_INT_STR(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT_INT_INT_STR,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer,
    ByVal intKEY03 As Integer,
    ByVal intKEY04 As Integer,
    ByVal intKEY05 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 _
                And .KEY03 = intKEY03 _
                And .KEY04 = intKEY04 _
                And .KEY05 = intKEY05 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

#Region "INT_INT_INT_INT_INT_BOOL"
    Public Sub SUB_ADD_CASH_INT_INT_INT_INT_INT_BOOL(
    ByRef srtCASH() As SRT_CASH_INT_INT_INT_INT_INT_BOOL,
    ByVal intKEY01 As Integer, ByVal intKEY02 As Integer, ByVal intKEY03 As Integer, ByVal intKEY04 As Integer, ByVal intKEY05 As Integer, ByVal blnVALUE As Boolean
    )
        Dim intINDEX As Integer

        If FUNC_SEARCH_CASH_INT_INT_INT_INT_INT_BOOL(srtCASH, intKEY01, intKEY02, intKEY03, intKEY04, intKEY05) <> -1 Then 'すでに存在するなら
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
            .KEY02 = intKEY02
            .KEY03 = intKEY03
            .KEY04 = intKEY04
            .KEY05 = intKEY05
            .VALUE = blnVALUE
        End With

    End Sub

    Public Function FUNC_SEARCH_CASH_INT_INT_INT_INT_INT_BOOL(
    ByRef srtSEARCH() As SRT_CASH_INT_INT_INT_INT_INT_BOOL,
    ByVal intKEY01 As Integer,
    ByVal intKEY02 As Integer,
    ByVal intKEY03 As Integer,
    ByVal intKEY04 As Integer,
    ByVal intKEY05 As Integer
    ) As Integer
        Dim intLOOP_INDEX As Integer
        Dim intRET As Integer

        If IsNothing(srtSEARCH) Then
            Return -1
        End If

        intRET = -1

        For intLOOP_INDEX = LBound(srtSEARCH) To UBound(srtSEARCH)
            With srtSEARCH(intLOOP_INDEX)
                If .KEY01 = intKEY01 _
                And .KEY02 = intKEY02 _
                And .KEY03 = intKEY03 _
                And .KEY04 = intKEY04 _
                And .KEY05 = intKEY05 Then
                    intRET = intLOOP_INDEX
                    Exit For
                End If
            End With
        Next

        Return intRET
    End Function

#End Region

End Module
