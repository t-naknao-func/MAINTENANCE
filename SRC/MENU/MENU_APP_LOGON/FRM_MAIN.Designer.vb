<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FRM_MAIN
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.grpHEAD = New System.Windows.Forms.GroupBox()
        Me.lblTITLE = New System.Windows.Forms.Label()
        Me.grpBODY = New System.Windows.Forms.GroupBox()
        Me.pnlPASSWORD = New System.Windows.Forms.Panel()
        Me.TXT_PASSWORD = New System.Windows.Forms.TextBox()
        Me.LBL_PASSWORD_GUIDE = New System.Windows.Forms.Label()
        Me.pnlID_USER = New System.Windows.Forms.Panel()
        Me.TXT_ID_USER = New System.Windows.Forms.TextBox()
        Me.LBL_ID_USER_GUIDE = New System.Windows.Forms.Label()
        Me.grpFOOT = New System.Windows.Forms.GroupBox()
        Me.pnlFUNCTION_GROUP = New System.Windows.Forms.Panel()
        Me.btnEND = New System.Windows.Forms.Button()
        Me.btnLOG_ON = New System.Windows.Forms.Button()
        Me.grpHEAD.SuspendLayout()
        Me.grpBODY.SuspendLayout()
        Me.pnlPASSWORD.SuspendLayout()
        Me.pnlID_USER.SuspendLayout()
        Me.grpFOOT.SuspendLayout()
        Me.pnlFUNCTION_GROUP.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpHEAD
        '
        Me.grpHEAD.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpHEAD.Controls.Add(Me.lblTITLE)
        Me.grpHEAD.Location = New System.Drawing.Point(10, 10)
        Me.grpHEAD.Name = "grpHEAD"
        Me.grpHEAD.Size = New System.Drawing.Size(760, 40)
        Me.grpHEAD.TabIndex = 0
        Me.grpHEAD.TabStop = False
        '
        'lblTITLE
        '
        Me.lblTITLE.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTITLE.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTITLE.Location = New System.Drawing.Point(10, 15)
        Me.lblTITLE.Name = "lblTITLE"
        Me.lblTITLE.Size = New System.Drawing.Size(736, 16)
        Me.lblTITLE.TabIndex = 1
        Me.lblTITLE.Text = "基幹業務システム"
        Me.lblTITLE.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grpBODY
        '
        Me.grpBODY.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpBODY.Controls.Add(Me.pnlPASSWORD)
        Me.grpBODY.Controls.Add(Me.pnlID_USER)
        Me.grpBODY.Location = New System.Drawing.Point(10, 50)
        Me.grpBODY.Name = "grpBODY"
        Me.grpBODY.Size = New System.Drawing.Size(760, 439)
        Me.grpBODY.TabIndex = 1
        Me.grpBODY.TabStop = False
        '
        'pnlPASSWORD
        '
        Me.pnlPASSWORD.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlPASSWORD.BackColor = System.Drawing.Color.FromArgb(CType(CType(225, Byte), Integer), CType(CType(225, Byte), Integer), CType(CType(235, Byte), Integer))
        Me.pnlPASSWORD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPASSWORD.Controls.Add(Me.TXT_PASSWORD)
        Me.pnlPASSWORD.Controls.Add(Me.LBL_PASSWORD_GUIDE)
        Me.pnlPASSWORD.Location = New System.Drawing.Point(8, 229)
        Me.pnlPASSWORD.Name = "pnlPASSWORD"
        Me.pnlPASSWORD.Size = New System.Drawing.Size(740, 32)
        Me.pnlPASSWORD.TabIndex = 1
        '
        'TXT_PASSWORD
        '
        Me.TXT_PASSWORD.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TXT_PASSWORD.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.TXT_PASSWORD.Location = New System.Drawing.Point(360, 2)
        Me.TXT_PASSWORD.MaxLength = 10
        Me.TXT_PASSWORD.Name = "TXT_PASSWORD"
        Me.TXT_PASSWORD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TXT_PASSWORD.Size = New System.Drawing.Size(260, 25)
        Me.TXT_PASSWORD.TabIndex = 2
        Me.TXT_PASSWORD.Tag = "Clear,Check,NotNull"
        Me.TXT_PASSWORD.Text = "WWWWWWWWWW"
        '
        'LBL_PASSWORD_GUIDE
        '
        Me.LBL_PASSWORD_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LBL_PASSWORD_GUIDE.AutoEllipsis = True
        Me.LBL_PASSWORD_GUIDE.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(250, Byte), Integer))
        Me.LBL_PASSWORD_GUIDE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LBL_PASSWORD_GUIDE.ForeColor = System.Drawing.Color.White
        Me.LBL_PASSWORD_GUIDE.Location = New System.Drawing.Point(280, 2)
        Me.LBL_PASSWORD_GUIDE.Name = "LBL_PASSWORD_GUIDE"
        Me.LBL_PASSWORD_GUIDE.Size = New System.Drawing.Size(79, 25)
        Me.LBL_PASSWORD_GUIDE.TabIndex = 1
        Me.LBL_PASSWORD_GUIDE.Text = "パスワード"
        Me.LBL_PASSWORD_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlID_USER
        '
        Me.pnlID_USER.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlID_USER.BackColor = System.Drawing.Color.FromArgb(CType(CType(225, Byte), Integer), CType(CType(225, Byte), Integer), CType(CType(235, Byte), Integer))
        Me.pnlID_USER.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlID_USER.Controls.Add(Me.TXT_ID_USER)
        Me.pnlID_USER.Controls.Add(Me.LBL_ID_USER_GUIDE)
        Me.pnlID_USER.Location = New System.Drawing.Point(8, 189)
        Me.pnlID_USER.Name = "pnlID_USER"
        Me.pnlID_USER.Size = New System.Drawing.Size(740, 32)
        Me.pnlID_USER.TabIndex = 0
        '
        'TXT_ID_USER
        '
        Me.TXT_ID_USER.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TXT_ID_USER.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.TXT_ID_USER.Location = New System.Drawing.Point(360, 2)
        Me.TXT_ID_USER.MaxLength = 30
        Me.TXT_ID_USER.Name = "TXT_ID_USER"
        Me.TXT_ID_USER.Size = New System.Drawing.Size(260, 25)
        Me.TXT_ID_USER.TabIndex = 2
        Me.TXT_ID_USER.Tag = "Clear,Check,NotNull"
        Me.TXT_ID_USER.Text = "WWWWWWWWWWWWWWWWWWWW"
        '
        'LBL_ID_USER_GUIDE
        '
        Me.LBL_ID_USER_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LBL_ID_USER_GUIDE.AutoEllipsis = True
        Me.LBL_ID_USER_GUIDE.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(250, Byte), Integer))
        Me.LBL_ID_USER_GUIDE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LBL_ID_USER_GUIDE.ForeColor = System.Drawing.Color.White
        Me.LBL_ID_USER_GUIDE.Location = New System.Drawing.Point(280, 2)
        Me.LBL_ID_USER_GUIDE.Name = "LBL_ID_USER_GUIDE"
        Me.LBL_ID_USER_GUIDE.Size = New System.Drawing.Size(79, 25)
        Me.LBL_ID_USER_GUIDE.TabIndex = 1
        Me.LBL_ID_USER_GUIDE.Text = "ユーザーID"
        Me.LBL_ID_USER_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpFOOT
        '
        Me.grpFOOT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpFOOT.Controls.Add(Me.pnlFUNCTION_GROUP)
        Me.grpFOOT.Location = New System.Drawing.Point(10, 489)
        Me.grpFOOT.Name = "grpFOOT"
        Me.grpFOOT.Size = New System.Drawing.Size(760, 60)
        Me.grpFOOT.TabIndex = 2
        Me.grpFOOT.TabStop = False
        '
        'pnlFUNCTION_GROUP
        '
        Me.pnlFUNCTION_GROUP.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlFUNCTION_GROUP.AutoScroll = True
        Me.pnlFUNCTION_GROUP.Controls.Add(Me.btnEND)
        Me.pnlFUNCTION_GROUP.Controls.Add(Me.btnLOG_ON)
        Me.pnlFUNCTION_GROUP.Location = New System.Drawing.Point(280, 16)
        Me.pnlFUNCTION_GROUP.MinimumSize = New System.Drawing.Size(190, 40)
        Me.pnlFUNCTION_GROUP.Name = "pnlFUNCTION_GROUP"
        Me.pnlFUNCTION_GROUP.Size = New System.Drawing.Size(190, 40)
        Me.pnlFUNCTION_GROUP.TabIndex = 0
        '
        'btnEND
        '
        Me.btnEND.AutoSize = True
        Me.btnEND.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnEND.Location = New System.Drawing.Point(100, 4)
        Me.btnEND.MinimumSize = New System.Drawing.Size(80, 30)
        Me.btnEND.Name = "btnEND"
        Me.btnEND.Size = New System.Drawing.Size(80, 30)
        Me.btnEND.TabIndex = 1
        Me.btnEND.Text = "終了"
        Me.btnEND.UseVisualStyleBackColor = False
        '
        'btnLOG_ON
        '
        Me.btnLOG_ON.AutoSize = True
        Me.btnLOG_ON.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnLOG_ON.Location = New System.Drawing.Point(10, 4)
        Me.btnLOG_ON.MinimumSize = New System.Drawing.Size(80, 30)
        Me.btnLOG_ON.Name = "btnLOG_ON"
        Me.btnLOG_ON.Size = New System.Drawing.Size(80, 30)
        Me.btnLOG_ON.TabIndex = 0
        Me.btnLOG_ON.Text = "ログオン"
        Me.btnLOG_ON.UseVisualStyleBackColor = False
        '
        'FRM_MAIN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(230, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.grpFOOT)
        Me.Controls.Add(Me.grpBODY)
        Me.Controls.Add(Me.grpHEAD)
        Me.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(800, 600)
        Me.Name = "FRM_MAIN"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "***"
        Me.grpHEAD.ResumeLayout(False)
        Me.grpBODY.ResumeLayout(False)
        Me.pnlPASSWORD.ResumeLayout(False)
        Me.pnlPASSWORD.PerformLayout()
        Me.pnlID_USER.ResumeLayout(False)
        Me.pnlID_USER.PerformLayout()
        Me.grpFOOT.ResumeLayout(False)
        Me.pnlFUNCTION_GROUP.ResumeLayout(False)
        Me.pnlFUNCTION_GROUP.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpHEAD As System.Windows.Forms.GroupBox
    Friend WithEvents grpBODY As System.Windows.Forms.GroupBox
    Friend WithEvents grpFOOT As System.Windows.Forms.GroupBox
    Friend WithEvents pnlFUNCTION_GROUP As System.Windows.Forms.Panel
    Friend WithEvents btnEND As System.Windows.Forms.Button
    Friend WithEvents btnLOG_ON As System.Windows.Forms.Button
    Friend WithEvents pnlID_USER As System.Windows.Forms.Panel
    Friend WithEvents LBL_ID_USER_GUIDE As System.Windows.Forms.Label
    Friend WithEvents TXT_ID_USER As System.Windows.Forms.TextBox
    Friend WithEvents pnlPASSWORD As System.Windows.Forms.Panel
    Friend WithEvents TXT_PASSWORD As System.Windows.Forms.TextBox
    Friend WithEvents LBL_PASSWORD_GUIDE As System.Windows.Forms.Label
    Friend WithEvents lblTITLE As System.Windows.Forms.Label

End Class
