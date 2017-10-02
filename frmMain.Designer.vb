<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
	Inherits frmCommon

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnSetting = New System.Windows.Forms.Button()
        Me.grpTargetForm = New System.Windows.Forms.GroupBox()
        Me.rbtnResultText = New System.Windows.Forms.RadioButton()
        Me.rbtnResultGridView = New System.Windows.Forms.RadioButton()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnSelectFolder = New System.Windows.Forms.Button()
        Me.txtFolderPath = New System.Windows.Forms.TextBox()
        Me.grpMaxCountInPage = New System.Windows.Forms.GroupBox()
        Me.txtMaxCountInPage = New Global.FolderFileList.NurmericTextBox()
        Me.lblMaxCountInPage = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkRunSaveFile = New System.Windows.Forms.CheckBox()
        Me.lblDocument = New System.Windows.Forms.Label()
        Me.grpTargetForm.SuspendLayout()
        Me.grpMaxCountInPage.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSetting
        '
        Me.btnSetting.Location = New System.Drawing.Point(563, 10)
        Me.btnSetting.Name = "btnSetting"
        Me.btnSetting.Size = New System.Drawing.Size(80, 23)
        Me.btnSetting.TabIndex = 3
        Me.btnSetting.Text = "設定(Ctrl+S)"
        Me.btnSetting.UseVisualStyleBackColor = True
        '
        'grpTargetForm
        '
        Me.grpTargetForm.Controls.Add(Me.rbtnResultText)
        Me.grpTargetForm.Controls.Add(Me.rbtnResultGridView)
        Me.grpTargetForm.Location = New System.Drawing.Point(12, 54)
        Me.grpTargetForm.Name = "grpTargetForm"
        Me.grpTargetForm.Size = New System.Drawing.Size(250, 50)
        Me.grpTargetForm.TabIndex = 4
        Me.grpTargetForm.TabStop = False
        Me.grpTargetForm.Text = "対象フォーム"
        '
        'rbtnResultText
        '
        Me.rbtnResultText.AutoSize = True
        Me.rbtnResultText.Location = New System.Drawing.Point(6, 19)
        Me.rbtnResultText.Name = "rbtnResultText"
        Me.rbtnResultText.Size = New System.Drawing.Size(124, 16)
        Me.rbtnResultText.TabIndex = 0
        Me.rbtnResultText.Text = "出力文字列(Ctrl+O)"
        Me.rbtnResultText.UseVisualStyleBackColor = True
        '
        'rbtnResultGridView
        '
        Me.rbtnResultGridView.AutoSize = True
        Me.rbtnResultGridView.Location = New System.Drawing.Point(136, 19)
        Me.rbtnResultGridView.Name = "rbtnResultGridView"
        Me.rbtnResultGridView.Size = New System.Drawing.Size(110, 16)
        Me.rbtnResultGridView.TabIndex = 1
        Me.rbtnResultGridView.Text = "リスト表示(Ctrl+L)"
        Me.rbtnResultGridView.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(477, 10)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(80, 23)
        Me.btnRun.TabIndex = 2
        Me.btnRun.Text = "実行(Ctrl+R)"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'btnSelectFolder
        '
        Me.btnSelectFolder.Location = New System.Drawing.Point(356, 10)
        Me.btnSelectFolder.Name = "btnSelectFolder"
        Me.btnSelectFolder.Size = New System.Drawing.Size(115, 23)
        Me.btnSelectFolder.TabIndex = 1
        Me.btnSelectFolder.Text = "フォルダ選択(Ctrl+F)"
        Me.btnSelectFolder.UseVisualStyleBackColor = True
        '
        'txtFolderPath
        '
        Me.txtFolderPath.AllowDrop = True
        Me.txtFolderPath.Location = New System.Drawing.Point(12, 12)
        Me.txtFolderPath.Name = "txtFolderPath"
        Me.txtFolderPath.Size = New System.Drawing.Size(337, 19)
        Me.txtFolderPath.TabIndex = 0
        '
        'grpMaxCountInPage
        '
        Me.grpMaxCountInPage.Controls.Add(Me.txtMaxCountInPage)
        Me.grpMaxCountInPage.Controls.Add(Me.lblMaxCountInPage)
        Me.grpMaxCountInPage.Location = New System.Drawing.Point(421, 54)
        Me.grpMaxCountInPage.Name = "grpMaxCountInPage"
        Me.grpMaxCountInPage.Size = New System.Drawing.Size(215, 50)
        Me.grpMaxCountInPage.TabIndex = 5
        Me.grpMaxCountInPage.TabStop = False
        Me.grpMaxCountInPage.Text = "最大表示ファイル数（リスト表示）(Ctrl+M)"
        '
        'txtMaxCountInPage
        '
        Me.txtMaxCountInPage.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtMaxCountInPage.Location = New System.Drawing.Point(6, 17)
        Me.txtMaxCountInPage.Name = "txtMaxCountInPage"
        Me.txtMaxCountInPage.NumberOfDigit = 5
        Me.txtMaxCountInPage.Size = New System.Drawing.Size(180, 19)
        Me.txtMaxCountInPage.TabIndex = 0
        Me.txtMaxCountInPage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblMaxCountInPage
        '
        Me.lblMaxCountInPage.AutoSize = True
        Me.lblMaxCountInPage.Location = New System.Drawing.Point(190, 20)
        Me.lblMaxCountInPage.Name = "lblMaxCountInPage"
        Me.lblMaxCountInPage.Size = New System.Drawing.Size(17, 12)
        Me.lblMaxCountInPage.TabIndex = 1
        Me.lblMaxCountInPage.Text = "件"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkRunSaveFile)
        Me.GroupBox1.Location = New System.Drawing.Point(268, 54)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(147, 50)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "保存ファイル実行(Ctrl+K)"
        '
        'chkRunSaveFile
        '
        Me.chkRunSaveFile.AutoSize = True
        Me.chkRunSaveFile.Location = New System.Drawing.Point(7, 20)
        Me.chkRunSaveFile.Name = "chkRunSaveFile"
        Me.chkRunSaveFile.Size = New System.Drawing.Size(130, 16)
        Me.chkRunSaveFile.TabIndex = 0
        Me.chkRunSaveFile.Text = "保存ファイル即時実行"
        Me.chkRunSaveFile.UseVisualStyleBackColor = True
        '
        'lblDocument
        '
        Me.lblDocument.AutoSize = True
        Me.lblDocument.ForeColor = System.Drawing.Color.Blue
        Me.lblDocument.Location = New System.Drawing.Point(10, 111)
        Me.lblDocument.Name = "lblDocument"
        Me.lblDocument.Size = New System.Drawing.Size(213, 12)
        Me.lblDocument.TabIndex = 7
        Me.lblDocument.Text = "FolderFileListのドキュメントはこちら(Ctrl+D)"
        '
        'frmMain
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(662, 132)
        Me.Controls.Add(Me.lblDocument)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpMaxCountInPage)
        Me.Controls.Add(Me.btnSetting)
        Me.Controls.Add(Me.grpTargetForm)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnSelectFolder)
        Me.Controls.Add(Me.txtFolderPath)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(670, 80)
        Me.Name = "frmMain"
        Me.Opacity = 0R
        Me.Text = "フォルダファイルリスト"
        Me.grpTargetForm.ResumeLayout(False)
        Me.grpTargetForm.PerformLayout()
        Me.grpMaxCountInPage.ResumeLayout(False)
        Me.grpMaxCountInPage.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtFolderPath As System.Windows.Forms.TextBox
	Friend WithEvents btnSelectFolder As System.Windows.Forms.Button
	Friend WithEvents btnRun As System.Windows.Forms.Button
	Friend WithEvents grpTargetForm As System.Windows.Forms.GroupBox
	Friend WithEvents rbtnResultText As System.Windows.Forms.RadioButton
	Friend WithEvents rbtnResultGridView As System.Windows.Forms.RadioButton
	Friend WithEvents btnSetting As System.Windows.Forms.Button
	Friend WithEvents grpMaxCountInPage As System.Windows.Forms.GroupBox
	Friend WithEvents lblMaxCountInPage As System.Windows.Forms.Label
	Friend WithEvents txtMaxCountInPage As NurmericTextBox
	Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
	Friend WithEvents chkRunSaveFile As System.Windows.Forms.CheckBox
	Friend WithEvents lblDocument As System.Windows.Forms.Label

End Class
