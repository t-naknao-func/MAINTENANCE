<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FRM_SYSTEM_TOTAL_FONT_SETTING
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
        Me.pnlFUNCTION_GROUP = New System.Windows.Forms.Panel()
        Me.btnCANCEL = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.pnlSIZE_FONT = New System.Windows.Forms.Panel()
        Me.cmbFONT_SIZE = New System.Windows.Forms.ComboBox()
        Me.lblSIZE_FONT_GUIDE = New System.Windows.Forms.Label()
        Me.pnlCOLOR_01 = New System.Windows.Forms.Panel()
        Me.btnCOLOR_01_INIT = New System.Windows.Forms.Button()
        Me.picCOLOR_01 = New System.Windows.Forms.PictureBox()
        Me.lblCOLOR_01_GUIDE = New System.Windows.Forms.Label()
        Me.cldCOLOR_SETTING = New System.Windows.Forms.ColorDialog()
        Me.pnlCOLOR_02 = New System.Windows.Forms.Panel()
        Me.btnCOLOR_02_INIT = New System.Windows.Forms.Button()
        Me.picCOLOR_02 = New System.Windows.Forms.PictureBox()
        Me.lblCOLOR_02_GUIDE = New System.Windows.Forms.Label()
        Me.pnlCOLOR_03 = New System.Windows.Forms.Panel()
        Me.btnCOLOR_03_INIT = New System.Windows.Forms.Button()
        Me.picCOLOR_03 = New System.Windows.Forms.PictureBox()
        Me.lblCOLOR_03_GUIDE = New System.Windows.Forms.Label()
        Me.pnlCOLOR_04 = New System.Windows.Forms.Panel()
        Me.btnCOLOR_04_INIT = New System.Windows.Forms.Button()
        Me.picCOLOR_04 = New System.Windows.Forms.PictureBox()
        Me.lblCOLOR_04_GUIDE = New System.Windows.Forms.Label()
        Me.pnlMANUAL = New System.Windows.Forms.Panel()
        Me.lblMANUAL = New System.Windows.Forms.Label()
        Me.pnlFUNCTION_GROUP.SuspendLayout()
        Me.pnlSIZE_FONT.SuspendLayout()
        Me.pnlCOLOR_01.SuspendLayout()
        CType(Me.picCOLOR_01, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCOLOR_02.SuspendLayout()
        CType(Me.picCOLOR_02, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCOLOR_03.SuspendLayout()
        CType(Me.picCOLOR_03, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCOLOR_04.SuspendLayout()
        CType(Me.picCOLOR_04, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMANUAL.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlFUNCTION_GROUP
        '
        Me.pnlFUNCTION_GROUP.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlFUNCTION_GROUP.AutoScroll = True
        Me.pnlFUNCTION_GROUP.Controls.Add(Me.btnCANCEL)
        Me.pnlFUNCTION_GROUP.Controls.Add(Me.btnOK)
        Me.pnlFUNCTION_GROUP.Location = New System.Drawing.Point(50, 315)
        Me.pnlFUNCTION_GROUP.MinimumSize = New System.Drawing.Size(190, 40)
        Me.pnlFUNCTION_GROUP.Name = "pnlFUNCTION_GROUP"
        Me.pnlFUNCTION_GROUP.Size = New System.Drawing.Size(190, 40)
        Me.pnlFUNCTION_GROUP.TabIndex = 5
        '
        'btnCANCEL
        '
        Me.btnCANCEL.AutoSize = True
        Me.btnCANCEL.Location = New System.Drawing.Point(100, 4)
        Me.btnCANCEL.MinimumSize = New System.Drawing.Size(80, 30)
        Me.btnCANCEL.Name = "btnCANCEL"
        Me.btnCANCEL.Size = New System.Drawing.Size(80, 30)
        Me.btnCANCEL.TabIndex = 1
        Me.btnCANCEL.Text = "キャンセル"
        Me.btnCANCEL.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.AutoSize = True
        Me.btnOK.BackColor = System.Drawing.Color.FromArgb(CType(CType(140, Byte), Integer), CType(CType(200, Byte), Integer), CType(CType(140, Byte), Integer))
        Me.btnOK.Location = New System.Drawing.Point(10, 4)
        Me.btnOK.MinimumSize = New System.Drawing.Size(80, 30)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 30)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'pnlSIZE_FONT
        '
        Me.pnlSIZE_FONT.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlSIZE_FONT.BackColor = System.Drawing.SystemColors.ControlDark
        Me.pnlSIZE_FONT.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSIZE_FONT.Controls.Add(Me.cmbFONT_SIZE)
        Me.pnlSIZE_FONT.Controls.Add(Me.lblSIZE_FONT_GUIDE)
        Me.pnlSIZE_FONT.Location = New System.Drawing.Point(10, 10)
        Me.pnlSIZE_FONT.Name = "pnlSIZE_FONT"
        Me.pnlSIZE_FONT.Size = New System.Drawing.Size(260, 32)
        Me.pnlSIZE_FONT.TabIndex = 0
        '
        'cmbFONT_SIZE
        '
        Me.cmbFONT_SIZE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFONT_SIZE.FormattingEnabled = True
        Me.cmbFONT_SIZE.Items.AddRange(New Object() {"9", "10", "11"})
        Me.cmbFONT_SIZE.Location = New System.Drawing.Point(140, 2)
        Me.cmbFONT_SIZE.Name = "cmbFONT_SIZE"
        Me.cmbFONT_SIZE.Size = New System.Drawing.Size(50, 26)
        Me.cmbFONT_SIZE.TabIndex = 1
        '
        'lblSIZE_FONT_GUIDE
        '
        Me.lblSIZE_FONT_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblSIZE_FONT_GUIDE.AutoEllipsis = True
        Me.lblSIZE_FONT_GUIDE.AutoSize = True
        Me.lblSIZE_FONT_GUIDE.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSIZE_FONT_GUIDE.ForeColor = System.Drawing.Color.White
        Me.lblSIZE_FONT_GUIDE.Location = New System.Drawing.Point(40, 5)
        Me.lblSIZE_FONT_GUIDE.Name = "lblSIZE_FONT_GUIDE"
        Me.lblSIZE_FONT_GUIDE.Size = New System.Drawing.Size(92, 18)
        Me.lblSIZE_FONT_GUIDE.TabIndex = 0
        Me.lblSIZE_FONT_GUIDE.Text = "フォントサイズ"
        Me.lblSIZE_FONT_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlCOLOR_01
        '
        Me.pnlCOLOR_01.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlCOLOR_01.BackColor = System.Drawing.SystemColors.ControlDark
        Me.pnlCOLOR_01.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCOLOR_01.Controls.Add(Me.btnCOLOR_01_INIT)
        Me.pnlCOLOR_01.Controls.Add(Me.picCOLOR_01)
        Me.pnlCOLOR_01.Controls.Add(Me.lblCOLOR_01_GUIDE)
        Me.pnlCOLOR_01.Location = New System.Drawing.Point(10, 45)
        Me.pnlCOLOR_01.Name = "pnlCOLOR_01"
        Me.pnlCOLOR_01.Size = New System.Drawing.Size(260, 32)
        Me.pnlCOLOR_01.TabIndex = 1
        '
        'btnCOLOR_01_INIT
        '
        Me.btnCOLOR_01_INIT.AutoSize = True
        Me.btnCOLOR_01_INIT.BackColor = System.Drawing.Color.FromArgb(CType(CType(140, Byte), Integer), CType(CType(200, Byte), Integer), CType(CType(140, Byte), Integer))
        Me.btnCOLOR_01_INIT.Location = New System.Drawing.Point(200, 0)
        Me.btnCOLOR_01_INIT.MinimumSize = New System.Drawing.Size(55, 30)
        Me.btnCOLOR_01_INIT.Name = "btnCOLOR_01_INIT"
        Me.btnCOLOR_01_INIT.Size = New System.Drawing.Size(55, 30)
        Me.btnCOLOR_01_INIT.TabIndex = 1
        Me.btnCOLOR_01_INIT.Text = "初期化"
        Me.btnCOLOR_01_INIT.UseVisualStyleBackColor = True
        '
        'picCOLOR_01
        '
        Me.picCOLOR_01.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picCOLOR_01.Location = New System.Drawing.Point(140, 5)
        Me.picCOLOR_01.Name = "picCOLOR_01"
        Me.picCOLOR_01.Size = New System.Drawing.Size(60, 20)
        Me.picCOLOR_01.TabIndex = 5
        Me.picCOLOR_01.TabStop = False
        '
        'lblCOLOR_01_GUIDE
        '
        Me.lblCOLOR_01_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCOLOR_01_GUIDE.AutoEllipsis = True
        Me.lblCOLOR_01_GUIDE.AutoSize = True
        Me.lblCOLOR_01_GUIDE.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblCOLOR_01_GUIDE.ForeColor = System.Drawing.Color.White
        Me.lblCOLOR_01_GUIDE.Location = New System.Drawing.Point(40, 5)
        Me.lblCOLOR_01_GUIDE.Name = "lblCOLOR_01_GUIDE"
        Me.lblCOLOR_01_GUIDE.Size = New System.Drawing.Size(51, 18)
        Me.lblCOLOR_01_GUIDE.TabIndex = 0
        Me.lblCOLOR_01_GUIDE.Text = "背景色1"
        Me.lblCOLOR_01_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlCOLOR_02
        '
        Me.pnlCOLOR_02.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlCOLOR_02.BackColor = System.Drawing.SystemColors.ControlDark
        Me.pnlCOLOR_02.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCOLOR_02.Controls.Add(Me.btnCOLOR_02_INIT)
        Me.pnlCOLOR_02.Controls.Add(Me.picCOLOR_02)
        Me.pnlCOLOR_02.Controls.Add(Me.lblCOLOR_02_GUIDE)
        Me.pnlCOLOR_02.Location = New System.Drawing.Point(10, 80)
        Me.pnlCOLOR_02.Name = "pnlCOLOR_02"
        Me.pnlCOLOR_02.Size = New System.Drawing.Size(260, 32)
        Me.pnlCOLOR_02.TabIndex = 2
        '
        'btnCOLOR_02_INIT
        '
        Me.btnCOLOR_02_INIT.AutoSize = True
        Me.btnCOLOR_02_INIT.BackColor = System.Drawing.Color.FromArgb(CType(CType(140, Byte), Integer), CType(CType(200, Byte), Integer), CType(CType(140, Byte), Integer))
        Me.btnCOLOR_02_INIT.Location = New System.Drawing.Point(200, 0)
        Me.btnCOLOR_02_INIT.MinimumSize = New System.Drawing.Size(55, 30)
        Me.btnCOLOR_02_INIT.Name = "btnCOLOR_02_INIT"
        Me.btnCOLOR_02_INIT.Size = New System.Drawing.Size(55, 30)
        Me.btnCOLOR_02_INIT.TabIndex = 1
        Me.btnCOLOR_02_INIT.Text = "初期化"
        Me.btnCOLOR_02_INIT.UseVisualStyleBackColor = True
        '
        'picCOLOR_02
        '
        Me.picCOLOR_02.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picCOLOR_02.Location = New System.Drawing.Point(140, 5)
        Me.picCOLOR_02.Name = "picCOLOR_02"
        Me.picCOLOR_02.Size = New System.Drawing.Size(60, 20)
        Me.picCOLOR_02.TabIndex = 5
        Me.picCOLOR_02.TabStop = False
        '
        'lblCOLOR_02_GUIDE
        '
        Me.lblCOLOR_02_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCOLOR_02_GUIDE.AutoEllipsis = True
        Me.lblCOLOR_02_GUIDE.AutoSize = True
        Me.lblCOLOR_02_GUIDE.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblCOLOR_02_GUIDE.ForeColor = System.Drawing.Color.White
        Me.lblCOLOR_02_GUIDE.Location = New System.Drawing.Point(40, 5)
        Me.lblCOLOR_02_GUIDE.Name = "lblCOLOR_02_GUIDE"
        Me.lblCOLOR_02_GUIDE.Size = New System.Drawing.Size(51, 18)
        Me.lblCOLOR_02_GUIDE.TabIndex = 0
        Me.lblCOLOR_02_GUIDE.Text = "背景色2"
        Me.lblCOLOR_02_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlCOLOR_03
        '
        Me.pnlCOLOR_03.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlCOLOR_03.BackColor = System.Drawing.SystemColors.ControlDark
        Me.pnlCOLOR_03.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCOLOR_03.Controls.Add(Me.btnCOLOR_03_INIT)
        Me.pnlCOLOR_03.Controls.Add(Me.picCOLOR_03)
        Me.pnlCOLOR_03.Controls.Add(Me.lblCOLOR_03_GUIDE)
        Me.pnlCOLOR_03.Location = New System.Drawing.Point(10, 115)
        Me.pnlCOLOR_03.Name = "pnlCOLOR_03"
        Me.pnlCOLOR_03.Size = New System.Drawing.Size(260, 32)
        Me.pnlCOLOR_03.TabIndex = 3
        '
        'btnCOLOR_03_INIT
        '
        Me.btnCOLOR_03_INIT.AutoSize = True
        Me.btnCOLOR_03_INIT.BackColor = System.Drawing.Color.FromArgb(CType(CType(140, Byte), Integer), CType(CType(200, Byte), Integer), CType(CType(140, Byte), Integer))
        Me.btnCOLOR_03_INIT.Location = New System.Drawing.Point(200, 0)
        Me.btnCOLOR_03_INIT.MinimumSize = New System.Drawing.Size(55, 30)
        Me.btnCOLOR_03_INIT.Name = "btnCOLOR_03_INIT"
        Me.btnCOLOR_03_INIT.Size = New System.Drawing.Size(55, 30)
        Me.btnCOLOR_03_INIT.TabIndex = 1
        Me.btnCOLOR_03_INIT.Text = "初期化"
        Me.btnCOLOR_03_INIT.UseVisualStyleBackColor = True
        '
        'picCOLOR_03
        '
        Me.picCOLOR_03.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picCOLOR_03.Location = New System.Drawing.Point(140, 5)
        Me.picCOLOR_03.Name = "picCOLOR_03"
        Me.picCOLOR_03.Size = New System.Drawing.Size(60, 20)
        Me.picCOLOR_03.TabIndex = 5
        Me.picCOLOR_03.TabStop = False
        '
        'lblCOLOR_03_GUIDE
        '
        Me.lblCOLOR_03_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCOLOR_03_GUIDE.AutoEllipsis = True
        Me.lblCOLOR_03_GUIDE.AutoSize = True
        Me.lblCOLOR_03_GUIDE.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblCOLOR_03_GUIDE.ForeColor = System.Drawing.Color.White
        Me.lblCOLOR_03_GUIDE.Location = New System.Drawing.Point(40, 5)
        Me.lblCOLOR_03_GUIDE.Name = "lblCOLOR_03_GUIDE"
        Me.lblCOLOR_03_GUIDE.Size = New System.Drawing.Size(51, 18)
        Me.lblCOLOR_03_GUIDE.TabIndex = 0
        Me.lblCOLOR_03_GUIDE.Text = "背景色3"
        Me.lblCOLOR_03_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlCOLOR_04
        '
        Me.pnlCOLOR_04.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlCOLOR_04.BackColor = System.Drawing.SystemColors.ControlDark
        Me.pnlCOLOR_04.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCOLOR_04.Controls.Add(Me.btnCOLOR_04_INIT)
        Me.pnlCOLOR_04.Controls.Add(Me.picCOLOR_04)
        Me.pnlCOLOR_04.Controls.Add(Me.lblCOLOR_04_GUIDE)
        Me.pnlCOLOR_04.Location = New System.Drawing.Point(10, 150)
        Me.pnlCOLOR_04.Name = "pnlCOLOR_04"
        Me.pnlCOLOR_04.Size = New System.Drawing.Size(260, 32)
        Me.pnlCOLOR_04.TabIndex = 4
        '
        'btnCOLOR_04_INIT
        '
        Me.btnCOLOR_04_INIT.AutoSize = True
        Me.btnCOLOR_04_INIT.BackColor = System.Drawing.Color.FromArgb(CType(CType(140, Byte), Integer), CType(CType(200, Byte), Integer), CType(CType(140, Byte), Integer))
        Me.btnCOLOR_04_INIT.Location = New System.Drawing.Point(200, 0)
        Me.btnCOLOR_04_INIT.MinimumSize = New System.Drawing.Size(55, 30)
        Me.btnCOLOR_04_INIT.Name = "btnCOLOR_04_INIT"
        Me.btnCOLOR_04_INIT.Size = New System.Drawing.Size(55, 30)
        Me.btnCOLOR_04_INIT.TabIndex = 1
        Me.btnCOLOR_04_INIT.Text = "初期化"
        Me.btnCOLOR_04_INIT.UseVisualStyleBackColor = True
        '
        'picCOLOR_04
        '
        Me.picCOLOR_04.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picCOLOR_04.Location = New System.Drawing.Point(140, 5)
        Me.picCOLOR_04.Name = "picCOLOR_04"
        Me.picCOLOR_04.Size = New System.Drawing.Size(60, 20)
        Me.picCOLOR_04.TabIndex = 5
        Me.picCOLOR_04.TabStop = False
        '
        'lblCOLOR_04_GUIDE
        '
        Me.lblCOLOR_04_GUIDE.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCOLOR_04_GUIDE.AutoEllipsis = True
        Me.lblCOLOR_04_GUIDE.AutoSize = True
        Me.lblCOLOR_04_GUIDE.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblCOLOR_04_GUIDE.ForeColor = System.Drawing.Color.White
        Me.lblCOLOR_04_GUIDE.Location = New System.Drawing.Point(40, 5)
        Me.lblCOLOR_04_GUIDE.Name = "lblCOLOR_04_GUIDE"
        Me.lblCOLOR_04_GUIDE.Size = New System.Drawing.Size(51, 18)
        Me.lblCOLOR_04_GUIDE.TabIndex = 0
        Me.lblCOLOR_04_GUIDE.Text = "背景色4"
        Me.lblCOLOR_04_GUIDE.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlMANUAL
        '
        Me.pnlMANUAL.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlMANUAL.BackColor = System.Drawing.SystemColors.ControlDark
        Me.pnlMANUAL.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlMANUAL.Controls.Add(Me.lblMANUAL)
        Me.pnlMANUAL.Location = New System.Drawing.Point(10, 185)
        Me.pnlMANUAL.Name = "pnlMANUAL"
        Me.pnlMANUAL.Size = New System.Drawing.Size(260, 120)
        Me.pnlMANUAL.TabIndex = 6
        '
        'lblMANUAL
        '
        Me.lblMANUAL.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblMANUAL.AutoEllipsis = True
        Me.lblMANUAL.AutoSize = True
        Me.lblMANUAL.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblMANUAL.ForeColor = System.Drawing.Color.White
        Me.lblMANUAL.Location = New System.Drawing.Point(40, 10)
        Me.lblMANUAL.Name = "lblMANUAL"
        Me.lblMANUAL.Size = New System.Drawing.Size(200, 90)
        Me.lblMANUAL.TabIndex = 0
        Me.lblMANUAL.Text = "フォントサイズの設定は、" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "一時設定となり、保存されません。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "背景色は、初期化を行うと、" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "保存された設定をクリアし、" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "デザイン時の色へ戻します。"
        Me.lblMANUAL.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FRM_SYSTEM_TOTAL_FONT_SETTING
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.ClientSize = New System.Drawing.Size(284, 362)
        Me.Controls.Add(Me.pnlMANUAL)
        Me.Controls.Add(Me.pnlCOLOR_04)
        Me.Controls.Add(Me.pnlCOLOR_03)
        Me.Controls.Add(Me.pnlCOLOR_02)
        Me.Controls.Add(Me.pnlCOLOR_01)
        Me.Controls.Add(Me.pnlSIZE_FONT)
        Me.Controls.Add(Me.pnlFUNCTION_GROUP)
        Me.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FRM_SYSTEM_TOTAL_FONT_SETTING"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "設定画面"
        Me.pnlFUNCTION_GROUP.ResumeLayout(False)
        Me.pnlFUNCTION_GROUP.PerformLayout()
        Me.pnlSIZE_FONT.ResumeLayout(False)
        Me.pnlSIZE_FONT.PerformLayout()
        Me.pnlCOLOR_01.ResumeLayout(False)
        Me.pnlCOLOR_01.PerformLayout()
        CType(Me.picCOLOR_01, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCOLOR_02.ResumeLayout(False)
        Me.pnlCOLOR_02.PerformLayout()
        CType(Me.picCOLOR_02, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCOLOR_03.ResumeLayout(False)
        Me.pnlCOLOR_03.PerformLayout()
        CType(Me.picCOLOR_03, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCOLOR_04.ResumeLayout(False)
        Me.pnlCOLOR_04.PerformLayout()
        CType(Me.picCOLOR_04, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMANUAL.ResumeLayout(False)
        Me.pnlMANUAL.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlFUNCTION_GROUP As System.Windows.Forms.Panel
    Friend WithEvents btnCANCEL As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents pnlSIZE_FONT As System.Windows.Forms.Panel
    Friend WithEvents lblSIZE_FONT_GUIDE As System.Windows.Forms.Label
    Friend WithEvents cmbFONT_SIZE As System.Windows.Forms.ComboBox
    Friend WithEvents pnlCOLOR_01 As System.Windows.Forms.Panel
    Friend WithEvents lblCOLOR_01_GUIDE As System.Windows.Forms.Label
    Friend WithEvents picCOLOR_01 As System.Windows.Forms.PictureBox
    Friend WithEvents cldCOLOR_SETTING As System.Windows.Forms.ColorDialog
    Friend WithEvents pnlCOLOR_02 As System.Windows.Forms.Panel
    Friend WithEvents picCOLOR_02 As System.Windows.Forms.PictureBox
    Friend WithEvents lblCOLOR_02_GUIDE As System.Windows.Forms.Label
    Friend WithEvents pnlCOLOR_03 As System.Windows.Forms.Panel
    Friend WithEvents picCOLOR_03 As System.Windows.Forms.PictureBox
    Friend WithEvents lblCOLOR_03_GUIDE As System.Windows.Forms.Label
    Friend WithEvents pnlCOLOR_04 As System.Windows.Forms.Panel
    Friend WithEvents picCOLOR_04 As System.Windows.Forms.PictureBox
    Friend WithEvents lblCOLOR_04_GUIDE As System.Windows.Forms.Label
    Friend WithEvents btnCOLOR_01_INIT As System.Windows.Forms.Button
    Friend WithEvents btnCOLOR_02_INIT As System.Windows.Forms.Button
    Friend WithEvents btnCOLOR_03_INIT As System.Windows.Forms.Button
    Friend WithEvents btnCOLOR_04_INIT As System.Windows.Forms.Button
    Friend WithEvents pnlMANUAL As System.Windows.Forms.Panel
    Friend WithEvents lblMANUAL As System.Windows.Forms.Label
End Class
