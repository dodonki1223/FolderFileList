<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmResultText
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
        Me.btnResultGridView = New System.Windows.Forms.Button()
        Me.btnTextOutput = New System.Windows.Forms.Button()
        Me.txtFolderFileList = New System.Windows.Forms.TextBox()
        Me.btnHtmlOutput = New System.Windows.Forms.Button()
        Me.grpDisplayTarget = New System.Windows.Forms.GroupBox()
        Me.rbtnFolder = New System.Windows.Forms.RadioButton()
        Me.rbtnAll = New System.Windows.Forms.RadioButton()
        Me.grpExtension = New System.Windows.Forms.GroupBox()
        Me.cmbExtension = New System.Windows.Forms.ComboBox()
        Me.grpName = New System.Windows.Forms.GroupBox()
        Me.btnNameSearch = New System.Windows.Forms.Button()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.grpDisplayTarget.SuspendLayout()
        Me.grpExtension.SuspendLayout()
        Me.grpName.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnResultGridView
        '
        Me.btnResultGridView.Location = New System.Drawing.Point(17, 612)
        Me.btnResultGridView.Name = "btnResultGridView"
        Me.btnResultGridView.Size = New System.Drawing.Size(120, 23)
        Me.btnResultGridView.TabIndex = 5
        Me.btnResultGridView.Text = "リスト表示フォーム(O)"
        Me.btnResultGridView.UseVisualStyleBackColor = True
        '
        'btnTextOutput
        '
        Me.btnTextOutput.Location = New System.Drawing.Point(563, 612)
        Me.btnTextOutput.Name = "btnTextOutput"
        Me.btnTextOutput.Size = New System.Drawing.Size(85, 23)
        Me.btnTextOutput.TabIndex = 6
        Me.btnTextOutput.Text = "TEXT出力(T)"
        Me.btnTextOutput.UseVisualStyleBackColor = True
        '
        'txtFolderFileList
        '
        Me.txtFolderFileList.BackColor = System.Drawing.Color.Black
        Me.txtFolderFileList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtFolderFileList.ForeColor = System.Drawing.Color.White
        Me.txtFolderFileList.Location = New System.Drawing.Point(17, 65)
        Me.txtFolderFileList.MinimumSize = New System.Drawing.Size(630, 540)
        Me.txtFolderFileList.Multiline = True
        Me.txtFolderFileList.Name = "txtFolderFileList"
        Me.txtFolderFileList.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtFolderFileList.Size = New System.Drawing.Size(630, 540)
        Me.txtFolderFileList.TabIndex = 4
        '
        'btnHtmlOutput
        '
        Me.btnHtmlOutput.Location = New System.Drawing.Point(471, 612)
        Me.btnHtmlOutput.Name = "btnHtmlOutput"
        Me.btnHtmlOutput.Size = New System.Drawing.Size(85, 23)
        Me.btnHtmlOutput.TabIndex = 6
        Me.btnHtmlOutput.Text = "HTML出力(H)"
        Me.btnHtmlOutput.UseVisualStyleBackColor = True
        '
        'grpDisplayTarget
        '
        Me.grpDisplayTarget.Controls.Add(Me.rbtnFolder)
        Me.grpDisplayTarget.Controls.Add(Me.rbtnAll)
        Me.grpDisplayTarget.Location = New System.Drawing.Point(17, 12)
        Me.grpDisplayTarget.Name = "grpDisplayTarget"
        Me.grpDisplayTarget.Size = New System.Drawing.Size(220, 45)
        Me.grpDisplayTarget.TabIndex = 0
        Me.grpDisplayTarget.TabStop = False
        Me.grpDisplayTarget.Text = "表示対象"
        '
        'rbtnFolder
        '
        Me.rbtnFolder.AutoSize = True
        Me.rbtnFolder.Location = New System.Drawing.Point(99, 18)
        Me.rbtnFolder.Name = "rbtnFolder"
        Me.rbtnFolder.Size = New System.Drawing.Size(118, 16)
        Me.rbtnFolder.TabIndex = 0
        Me.rbtnFolder.Text = "フォルダのみ表示(K)"
        Me.rbtnFolder.UseVisualStyleBackColor = True
        '
        'rbtnAll
        '
        Me.rbtnAll.AutoSize = True
        Me.rbtnAll.Checked = True
        Me.rbtnAll.Location = New System.Drawing.Point(9, 18)
        Me.rbtnAll.Name = "rbtnAll"
        Me.rbtnAll.Size = New System.Drawing.Size(84, 16)
        Me.rbtnAll.TabIndex = 1
        Me.rbtnAll.TabStop = True
        Me.rbtnAll.Text = "全て表示(A)"
        Me.rbtnAll.UseVisualStyleBackColor = True
        '
        'grpExtension
        '
        Me.grpExtension.Controls.Add(Me.cmbExtension)
        Me.grpExtension.Location = New System.Drawing.Point(243, 12)
        Me.grpExtension.Name = "grpExtension"
        Me.grpExtension.Size = New System.Drawing.Size(111, 45)
        Me.grpExtension.TabIndex = 1
        Me.grpExtension.TabStop = False
        Me.grpExtension.Text = "拡張子指定(E)"
        '
        'cmbExtension
        '
        Me.cmbExtension.BackColor = System.Drawing.SystemColors.Window
        Me.cmbExtension.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbExtension.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbExtension.FormattingEnabled = True
        Me.cmbExtension.Location = New System.Drawing.Point(8, 15)
        Me.cmbExtension.Name = "cmbExtension"
        Me.cmbExtension.Size = New System.Drawing.Size(95, 20)
        Me.cmbExtension.TabIndex = 0
        '
        'grpName
        '
        Me.grpName.Controls.Add(Me.btnNameSearch)
        Me.grpName.Controls.Add(Me.txtName)
        Me.grpName.Location = New System.Drawing.Point(360, 12)
        Me.grpName.Name = "grpName"
        Me.grpName.Size = New System.Drawing.Size(287, 45)
        Me.grpName.TabIndex = 3
        Me.grpName.TabStop = False
        Me.grpName.Text = "名前（LIKE検索）(Ctrl+F)"
        '
        'btnNameSearch
        '
        Me.btnNameSearch.Location = New System.Drawing.Point(241, 13)
        Me.btnNameSearch.Name = "btnNameSearch"
        Me.btnNameSearch.Size = New System.Drawing.Size(40, 23)
        Me.btnNameSearch.TabIndex = 1
        Me.btnNameSearch.Text = "検索"
        Me.btnNameSearch.UseVisualStyleBackColor = True
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(7, 15)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(228, 19)
        Me.txtName.TabIndex = 0
        '
        'frmResultText
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.ClientSize = New System.Drawing.Size(664, 642)
        Me.Controls.Add(Me.grpName)
        Me.Controls.Add(Me.grpExtension)
        Me.Controls.Add(Me.grpDisplayTarget)
        Me.Controls.Add(Me.btnHtmlOutput)
        Me.Controls.Add(Me.btnResultGridView)
        Me.Controls.Add(Me.btnTextOutput)
        Me.Controls.Add(Me.txtFolderFileList)
        Me.MinimumSize = New System.Drawing.Size(680, 680)
        Me.Name = "frmResultText"
        Me.Opacity = 0.0R
        Me.Text = "出力文字列フォーム"
        Me.grpDisplayTarget.ResumeLayout(False)
        Me.grpDisplayTarget.PerformLayout()
        Me.grpExtension.ResumeLayout(False)
        Me.grpName.ResumeLayout(False)
        Me.grpName.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
	Friend WithEvents txtFolderFileList As System.Windows.Forms.TextBox
	Friend WithEvents btnTextOutput As System.Windows.Forms.Button
	Friend WithEvents btnResultGridView As Button
    Friend WithEvents btnHtmlOutput As System.Windows.Forms.Button
    Friend WithEvents grpDisplayTarget As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnFolder As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnAll As System.Windows.Forms.RadioButton
    Friend WithEvents grpExtension As System.Windows.Forms.GroupBox
    Friend WithEvents cmbExtension As System.Windows.Forms.ComboBox
    Friend WithEvents grpName As System.Windows.Forms.GroupBox
    Friend WithEvents btnNameSearch As System.Windows.Forms.Button
    Friend WithEvents txtName As System.Windows.Forms.TextBox
End Class
