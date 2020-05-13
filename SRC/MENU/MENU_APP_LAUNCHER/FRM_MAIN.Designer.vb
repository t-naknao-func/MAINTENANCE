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
        Me.pnlDATE_ACTIVE = New System.Windows.Forms.Panel()
        Me.lblDATE_ACTIVE = New System.Windows.Forms.Label()
        Me.lblDATE_ACTIVE_GUIDE = New System.Windows.Forms.Label()
        Me.pnlNAME_USER = New System.Windows.Forms.Panel()
        Me.lblNAME_USER = New System.Windows.Forms.Label()
        Me.lblNAME_USER_GUIDE = New System.Windows.Forms.Label()
        Me.grpFOOT = New System.Windows.Forms.GroupBox()
        Me.pnlFUNCTION_GROUP = New System.Windows.Forms.Panel()
        Me.btnEND = New System.Windows.Forms.Button()
        Me.btnLOG_OFF = New System.Windows.Forms.Button()
        Me.grpBODY = New System.Windows.Forms.GroupBox()
        Me.pnlLAUNCHER_GROUP = New System.Windows.Forms.Panel()
        Me.tabLAUNCHER_GROUP = New System.Windows.Forms.TabControl()
        Me.lblLAUNCHER_BASE = New System.Windows.Forms.Label()
        Me.btnLAUNCHER_BASE = New System.Windows.Forms.Button()
        Me.grpHEAD.SuspendLayout()
        Me.pnlDATE_ACTIVE.SuspendLayout()
        Me.pnlNAME_USER.SuspendLayout()
        Me.grpFOOT.SuspendLayout()
        Me.pnlFUNCTION_GROUP.SuspendLayout()
        Me.grpBODY.SuspendLayout()
        Me.pnlLAUNCHER_GROUP.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpHEAD
        '
        Me.grpHEAD.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpHEAD.Controls.Add(Me.pnlDATE_ACTIVE)
        Me.grpHEAD.Controls.Add(Me.pnlNAME_USER)
        Me.grpHEAD.Location = New System.Drawing.Point(10, 10)
        Me.grpHEAD.Name = "grpHEAD"
        Me.grpHEAD.Size = New System.Drawing.Size(760, 50)
        Me.grpHEAD.TabIndex = 0
        Me.grpHEAD.TabStop = False
        '
        'pnlDATE_ACTIVE
        '
        Me.pnlDATE_ACTIVE.BackColor = System.Drawing.Color.FromArgb(CType(CType(225, Byte), Integer), CType(CType(225, Byte), Integer), CType(CType(235, Byte), Integer))
        Me.pnlDATE_ACTIVE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlDATE_ACTIVE.Controls.Add(Me.lblDATE_ACTIVE)
        Me.pnlDATE_ACTIVE.Controls.Add(Me.lblDATE_ACTIVE_GUIDE)
        Me.pnlDATE_ACTIVE.Location = New System.Drawing.Point(10, 12)
        Me.pnlDATE_ACTIVE.Name = "pnlDATE_ACTIVE"
        Me.pnlDATE_ACTIVE.Size = New System.Drawing.Size(365, 34)
        Me.pnlDATE_ACTIVE.TabIndex = 2
        '
        'lblDATE_ACTIVE
        '
        Me.lblDATE_ACTIVE.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblDATE_ACTIVE.AutoEllipsis = True
        Me.lblDATE_ACTIVE.BackColor = System.Drawing.Color.White
        Me.lblDATE_ACTIVE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDATE_ACTIVE.ForeColor = System.Drawing.Color.Black
        Me.lblDATE_ACTIVE.Location = New System.Drawing.Point(80, 3)
        Me.lblDATE_ACTIVE.Name = "lblDATE_ACTIVE"
        Me.lblDATE_ACTIVE.Size = New System.Drawing.Size(120, 25)
        Me.lblDATE_ACTIVE.TabIndex = 2
        Me.lblDATE_ACTIVE.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblDATE_ACTIVE_GUIDE
        '
        Me.lblDATE_ACTIVE_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblDATE_ACTIVE_GUIDE.AutoEllipsis = True
        Me.lblDATE_ACTIVE_GUIDE.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(250, Byte), Integer))
        Me.lblDATE_ACTIVE_GUIDE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDATE_ACTIVE_GUIDE.ForeColor = System.Drawing.Color.Black
        Me.lblDATE_ACTIVE_GUIDE.Location = New System.Drawing.Point(1, 3)
        Me.lblDATE_ACTIVE_GUIDE.Name = "lblDATE_ACTIVE_GUIDE"
        Me.lblDATE_ACTIVE_GUIDE.Size = New System.Drawing.Size(79, 25)
        Me.lblDATE_ACTIVE_GUIDE.TabIndex = 1
        Me.lblDATE_ACTIVE_GUIDE.Text = "処理日"
        Me.lblDATE_ACTIVE_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlNAME_USER
        '
        Me.pnlNAME_USER.BackColor = System.Drawing.Color.FromArgb(CType(CType(225, Byte), Integer), CType(CType(225, Byte), Integer), CType(CType(235, Byte), Integer))
        Me.pnlNAME_USER.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlNAME_USER.Controls.Add(Me.lblNAME_USER)
        Me.pnlNAME_USER.Controls.Add(Me.lblNAME_USER_GUIDE)
        Me.pnlNAME_USER.Location = New System.Drawing.Point(385, 12)
        Me.pnlNAME_USER.Name = "pnlNAME_USER"
        Me.pnlNAME_USER.Size = New System.Drawing.Size(365, 34)
        Me.pnlNAME_USER.TabIndex = 1
        '
        'lblNAME_USER
        '
        Me.lblNAME_USER.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblNAME_USER.AutoEllipsis = True
        Me.lblNAME_USER.BackColor = System.Drawing.Color.White
        Me.lblNAME_USER.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNAME_USER.ForeColor = System.Drawing.Color.Black
        Me.lblNAME_USER.Location = New System.Drawing.Point(80, 3)
        Me.lblNAME_USER.Name = "lblNAME_USER"
        Me.lblNAME_USER.Size = New System.Drawing.Size(240, 25)
        Me.lblNAME_USER.TabIndex = 2
        Me.lblNAME_USER.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblNAME_USER_GUIDE
        '
        Me.lblNAME_USER_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblNAME_USER_GUIDE.AutoEllipsis = True
        Me.lblNAME_USER_GUIDE.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(250, Byte), Integer))
        Me.lblNAME_USER_GUIDE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNAME_USER_GUIDE.ForeColor = System.Drawing.Color.Black
        Me.lblNAME_USER_GUIDE.Location = New System.Drawing.Point(1, 3)
        Me.lblNAME_USER_GUIDE.Name = "lblNAME_USER_GUIDE"
        Me.lblNAME_USER_GUIDE.Size = New System.Drawing.Size(79, 25)
        Me.lblNAME_USER_GUIDE.TabIndex = 1
        Me.lblNAME_USER_GUIDE.Text = "ログオン者"
        Me.lblNAME_USER_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
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
        Me.pnlFUNCTION_GROUP.Controls.Add(Me.btnLOG_OFF)
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
        Me.btnEND.TabIndex = 2
        Me.btnEND.Text = "終了"
        Me.btnEND.UseVisualStyleBackColor = False
        '
        'btnLOG_OFF
        '
        Me.btnLOG_OFF.AutoSize = True
        Me.btnLOG_OFF.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnLOG_OFF.Location = New System.Drawing.Point(10, 4)
        Me.btnLOG_OFF.MinimumSize = New System.Drawing.Size(80, 30)
        Me.btnLOG_OFF.Name = "btnLOG_OFF"
        Me.btnLOG_OFF.Size = New System.Drawing.Size(80, 30)
        Me.btnLOG_OFF.TabIndex = 0
        Me.btnLOG_OFF.Text = "ログオフ"
        Me.btnLOG_OFF.UseVisualStyleBackColor = False
        '
        'grpBODY
        '
        Me.grpBODY.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpBODY.Controls.Add(Me.pnlLAUNCHER_GROUP)
        Me.grpBODY.Location = New System.Drawing.Point(10, 60)
        Me.grpBODY.Name = "grpBODY"
        Me.grpBODY.Size = New System.Drawing.Size(760, 420)
        Me.grpBODY.TabIndex = 1
        Me.grpBODY.TabStop = False
        '
        'pnlLAUNCHER_GROUP
        '
        Me.pnlLAUNCHER_GROUP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlLAUNCHER_GROUP.Controls.Add(Me.tabLAUNCHER_GROUP)
        Me.pnlLAUNCHER_GROUP.Controls.Add(Me.lblLAUNCHER_BASE)
        Me.pnlLAUNCHER_GROUP.Controls.Add(Me.btnLAUNCHER_BASE)
        Me.pnlLAUNCHER_GROUP.Location = New System.Drawing.Point(10, 30)
        Me.pnlLAUNCHER_GROUP.Name = "pnlLAUNCHER_GROUP"
        Me.pnlLAUNCHER_GROUP.Size = New System.Drawing.Size(736, 360)
        Me.pnlLAUNCHER_GROUP.TabIndex = 0
        '
        'tabLAUNCHER_GROUP
        '
        Me.tabLAUNCHER_GROUP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabLAUNCHER_GROUP.Location = New System.Drawing.Point(0, 0)
        Me.tabLAUNCHER_GROUP.Name = "tabLAUNCHER_GROUP"
        Me.tabLAUNCHER_GROUP.SelectedIndex = 0
        Me.tabLAUNCHER_GROUP.Size = New System.Drawing.Size(736, 360)
        Me.tabLAUNCHER_GROUP.TabIndex = 0
        Me.tabLAUNCHER_GROUP.TabStop = False
        '
        'lblLAUNCHER_BASE
        '
        Me.lblLAUNCHER_BASE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblLAUNCHER_BASE.AutoEllipsis = True
        Me.lblLAUNCHER_BASE.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(250, Byte), Integer))
        Me.lblLAUNCHER_BASE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLAUNCHER_BASE.ForeColor = System.Drawing.Color.Black
        Me.lblLAUNCHER_BASE.Location = New System.Drawing.Point(8, 0)
        Me.lblLAUNCHER_BASE.Name = "lblLAUNCHER_BASE"
        Me.lblLAUNCHER_BASE.Size = New System.Drawing.Size(175, 30)
        Me.lblLAUNCHER_BASE.TabIndex = 1
        Me.lblLAUNCHER_BASE.Text = "***"
        Me.lblLAUNCHER_BASE.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLAUNCHER_BASE.Visible = False
        '
        'btnLAUNCHER_BASE
        '
        Me.btnLAUNCHER_BASE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnLAUNCHER_BASE.AutoEllipsis = True
        Me.btnLAUNCHER_BASE.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.btnLAUNCHER_BASE.Location = New System.Drawing.Point(8, 30)
        Me.btnLAUNCHER_BASE.Name = "btnLAUNCHER_BASE"
        Me.btnLAUNCHER_BASE.Size = New System.Drawing.Size(175, 30)
        Me.btnLAUNCHER_BASE.TabIndex = 2
        Me.btnLAUNCHER_BASE.Text = "***"
        Me.btnLAUNCHER_BASE.UseVisualStyleBackColor = False
        Me.btnLAUNCHER_BASE.Visible = False
        '
        'FRM_MAIN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(230, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.grpBODY)
        Me.Controls.Add(Me.grpFOOT)
        Me.Controls.Add(Me.grpHEAD)
        Me.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(800, 600)
        Me.Name = "FRM_MAIN"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "***"
        Me.grpHEAD.ResumeLayout(False)
        Me.pnlDATE_ACTIVE.ResumeLayout(False)
        Me.pnlNAME_USER.ResumeLayout(False)
        Me.grpFOOT.ResumeLayout(False)
        Me.pnlFUNCTION_GROUP.ResumeLayout(False)
        Me.pnlFUNCTION_GROUP.PerformLayout()
        Me.grpBODY.ResumeLayout(False)
        Me.pnlLAUNCHER_GROUP.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpHEAD As System.Windows.Forms.GroupBox
    Friend WithEvents grpFOOT As System.Windows.Forms.GroupBox
    Friend WithEvents pnlFUNCTION_GROUP As System.Windows.Forms.Panel
    Friend WithEvents btnEND As System.Windows.Forms.Button
    Friend WithEvents btnLOG_OFF As System.Windows.Forms.Button
    Friend WithEvents grpBODY As System.Windows.Forms.GroupBox
    Friend WithEvents pnlLAUNCHER_GROUP As System.Windows.Forms.Panel
    Friend WithEvents lblLAUNCHER_BASE As System.Windows.Forms.Label
    Friend WithEvents btnLAUNCHER_BASE As System.Windows.Forms.Button
    Friend WithEvents tabLAUNCHER_GROUP As System.Windows.Forms.TabControl
    Friend WithEvents pnlDATE_ACTIVE As Panel
    Friend WithEvents lblDATE_ACTIVE As Label
    Friend WithEvents lblDATE_ACTIVE_GUIDE As Label
    Friend WithEvents pnlNAME_USER As Panel
    Friend WithEvents lblNAME_USER As Label
    Friend WithEvents lblNAME_USER_GUIDE As Label
End Class
