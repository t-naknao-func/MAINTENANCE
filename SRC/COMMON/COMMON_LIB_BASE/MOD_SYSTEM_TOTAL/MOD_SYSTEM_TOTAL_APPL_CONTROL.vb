Public Module MOD_SYSTEM_TOTAL_APPL_CONTROL
    '一つ前の機能を呼出す
    Public Function FUNC_SYSTEM_TOTAL_CALL_BACK_APPL() As Boolean
        Dim strSHELL As String
        Dim strCOMMAND_LINE As String
        Dim strPATH_EXE As String
        Dim intCODE_STAFF As Integer
        Dim intID_FOCUS_MENU_TOTAL As Integer
        Dim intID_FOCUS_MENU As Integer

        strSHELL = srtSYSTEM_TOTAL_COMMANDLINE.CALL_EXE_PATH

        strPATH_EXE = System.Windows.Forms.Application.ExecutablePath
        intCODE_STAFF = srtSYSTEM_TOTAL_COMMANDLINE.CODE_STAFF
        intID_FOCUS_MENU_TOTAL = srtSYSTEM_TOTAL_COMMANDLINE.ID_FOCUS_MENU_TOTAL
        intID_FOCUS_MENU = srtSYSTEM_TOTAL_COMMANDLINE.ID_FOCUS_MENU

        strCOMMAND_LINE = FUNC_SYSTEM_TOTAL_MAKE_COMMANDLINE(strPATH_EXE, intCODE_STAFF, intID_FOCUS_MENU_TOTAL, intID_FOCUS_MENU)

        If Not FUNC_CALL_EXE_FILE_SHELL(strSHELL, strCOMMAND_LINE) Then
            Return False
        End If

        Return False
    End Function
End Module
