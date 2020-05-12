Public Module MOD_SYSTEM_TOTAL_APP_CONFIG

#Region "設定文字列定数"
    Public Const CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCF As String = "window_setting_back_color_form"
    Public Const CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCG As String = "window_setting_back_color_group"
    Public Const CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCP As String = "window_setting_back_color_panel"
    Public Const CST_SYSTEM_TOTAL_APP_CONFIG_STR_BCO As String = "window_setting_back_color_object"
#End Region

    Public Function FUNC_SYSTEM_TOTAL_WRITE_APP_CONFIG(ByVal strKEY_NAME As String, ByVal strVALUE As String)
        Dim asm As System.Reflection.Assembly
        Dim appConfigPath As String
        Dim doc As System.Xml.XmlDocument
        Dim node As System.Xml.XmlNode
        Dim newNode As System.Xml.XmlElement
        Dim blnMOD As Boolean
        Dim strNAME_APPL As String

        Const cstCONFIG_EXTENT As String = ".config"
        strNAME_APPL = System.IO.Path.GetFileName(System.Windows.Forms.Application.ExecutablePath)
        asm = System.Reflection.Assembly.GetExecutingAssembly()
        appConfigPath = System.IO.Path.GetDirectoryName(asm.Location) & "\" & strNAME_APPL & cstCONFIG_EXTENT

        doc = New System.Xml.XmlDocument

        Try
            Call doc.Load(appConfigPath)
        Catch ex As Exception
            doc = Nothing
            Return False
        End Try

        node = doc("configuration")("appSettings")

        blnMOD = False
        For Each n In doc("configuration")("appSettings")
            If n.Name = "add" Then
                If n.Attributes.GetNamedItem("key").Value = strKEY_NAME Then
                    n.Attributes.GetNamedItem("value").Value = strVALUE
                    blnMOD = True
                End If
            End If
        Next

        If Not blnMOD Then
            newNode = doc.CreateElement("add")
            newNode.SetAttribute("key", strKEY_NAME)
            newNode.SetAttribute("value", strVALUE)
            node.AppendChild(newNode)
        End If

        doc.Save(appConfigPath)
        doc = Nothing

        Call System.Configuration.ConfigurationManager.RefreshSection("appSettings")

        Return True
    End Function

    Public Function FUNC_SYSTEM_TOTAL_DELETE_APP_CONFIG(ByVal strKEY_NAME As String)

        Dim asm As System.Reflection.Assembly
        Dim appConfigPath As String
        Dim doc As System.Xml.XmlDocument
        Dim node As System.Xml.XmlNode
        Dim strNAME_APPL As String

        Const cstCONFIG_EXTENT As String = ".config"
        strNAME_APPL = System.IO.Path.GetFileName(System.Windows.Forms.Application.ExecutablePath)
        asm = System.Reflection.Assembly.GetExecutingAssembly()
        appConfigPath = System.IO.Path.GetDirectoryName(asm.Location) & "\" & strNAME_APPL & cstCONFIG_EXTENT

        doc = New System.Xml.XmlDocument

        Try
            Call doc.Load(appConfigPath)
        Catch ex As Exception
            doc = Nothing
            Return False
        End Try

        node = doc("configuration")("appSettings")

        For Each n In doc("configuration")("appSettings")
            If n.Name = "add" Then
                If n.Attributes.GetNamedItem("key").Value = strKEY_NAME Then
                    doc("configuration")("appSettings").RemoveChild(n)
                End If
            End If
        Next

        doc.Save(appConfigPath)
        doc = Nothing

        Call System.Configuration.ConfigurationManager.RefreshSection("appSettings")

        Return True
    End Function

    Public Function FUNC_SYSTEM_TOTAL_SAVE_APP_CONFIG(ByVal strKEY_NAME As String, ByVal strVALUE As String)
        Dim config As System.Configuration.Configuration

        config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None)
        Call config.AppSettings.Settings.Add(strKEY_NAME, strVALUE)
        Call config.Save(System.Configuration.ConfigurationSaveMode.Modified)
        Call System.Configuration.ConfigurationManager.RefreshSection("appSettings")

        Return True
    End Function

    Public Function FUNC_SYSTEM_TOTAL_GET_APP_CONFIG(ByVal strKEY_NAME As String) As String
        Dim strRET As String

        Try
            strRET = System.Configuration.ConfigurationManager.AppSettings(strKEY_NAME)
        Catch ex As Exception
            Return ""
        End Try

        Return strRET
    End Function

End Module
