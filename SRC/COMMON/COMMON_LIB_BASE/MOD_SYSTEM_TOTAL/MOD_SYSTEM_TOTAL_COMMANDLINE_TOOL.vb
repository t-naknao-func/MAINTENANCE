Public Module MOD_SYSTEM_TOTAL_COMMANDLINE_TOOL

    Public Enum ENM_SYSTEM_COMMANDLINE
        CALL_EXE_PATH = 0
        CODE_STAFF
        ID_FOCUS_MENU_TOTAL
        ID_FOCUS_MENU
    End Enum

    Public Structure SRT_SYSTEM_COMMANDLINE
        Public CALL_EXE_PATH As String '呼出元EXEパス
        Public CODE_STAFF As Integer
        Public ID_FOCUS_MENU_TOTAL As Integer
        Public ID_FOCUS_MENU As Integer
    End Structure

    'コマンドライン情報の取得
    Public Function FUNC_SYSTEM_TOTAL_GET_COMMANDLINE(ByRef srtGET_CMMANDLINE As SRT_SYSTEM_COMMANDLINE, ByVal strCOMMANDLINE() As String) As Boolean

        '初期化
        With srtGET_CMMANDLINE
            .CALL_EXE_PATH = ""
            .CODE_STAFF = 999999998
            .ID_FOCUS_MENU_TOTAL = 0
            .ID_FOCUS_MENU = 0
        End With

        If strCOMMANDLINE Is Nothing Then
            Return True '省略はOKとする
        End If

        If strCOMMANDLINE.Length <= 0 Then
            Return True '省略はOKとする
        End If

        With srtGET_CMMANDLINE
            Try
                .CALL_EXE_PATH = strCOMMANDLINE(ENM_SYSTEM_COMMANDLINE.CALL_EXE_PATH)
                .CODE_STAFF = strCOMMANDLINE(ENM_SYSTEM_COMMANDLINE.CODE_STAFF)
                .ID_FOCUS_MENU_TOTAL = strCOMMANDLINE(ENM_SYSTEM_COMMANDLINE.ID_FOCUS_MENU_TOTAL)
                .ID_FOCUS_MENU = strCOMMANDLINE(ENM_SYSTEM_COMMANDLINE.ID_FOCUS_MENU)
            Catch ex As Exception
                Return False
            End Try
        End With

        Return True
    End Function

    'コマンドラインの確認
    Public Sub SUB_SYSTEM_TOTAL_SHOW_COMMANDLINE(ByVal srtCOMMANDLINE As SRT_SYSTEM_COMMANDLINE)
        Dim strMSG As String

        strMSG = ""
        With srtCOMMANDLINE
            strMSG &= "CALL_EXE_PATH:" & .CALL_EXE_PATH & Environment.NewLine
            strMSG &= "CODE_STAFF:" & .CODE_STAFF & Environment.NewLine
            strMSG &= "ID_FOCUS_MENU_TOTAL:" & .ID_FOCUS_MENU_TOTAL & Environment.NewLine
            strMSG &= "ID_FOCUS_MENU:" & .ID_FOCUS_MENU & Environment.NewLine
        End With

        Call System.Windows.Forms.MessageBox.Show(strMSG, "コマンドライン内容確認")

    End Sub

    'コマンドラインの作成
    Public Function FUNC_SYSTEM_TOTAL_MAKE_COMMANDLINE( _
    ByVal strPATH_EXE As String, ByVal intCODE_STAFF As Integer, _
    ByVal intID_FOCUS_MENU_TOTAL As Integer, ByVal intID_FOCUS_MENU As Integer _
    )
        Dim strRET As String

        strRET = ""
        strRET &= FUNC_ADD_ENCLOSED(strPATH_EXE, """") & " "
        strRET &= intCODE_STAFF & " "
        strRET &= intID_FOCUS_MENU_TOTAL & " "
        strRET &= intID_FOCUS_MENU & " "

        Return strRET
    End Function
End Module
