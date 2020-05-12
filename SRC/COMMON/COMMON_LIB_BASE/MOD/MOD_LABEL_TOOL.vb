Public Module MOD_LABEL_TOOL

    'Public Sub SUB_INIT_ANTI_ALIAS_LABEL(ByRef lblCONTROL As Label)
    '    With lblCONTROL
    '        RemoveHandler .OnPaint, AddressOf OnPaint
    '    End With

    'End Sub

    Public Sub OnPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
        sender.OnPaint(e)
    End Sub

End Module
