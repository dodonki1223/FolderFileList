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
		Me.btnOutput = New System.Windows.Forms.Button()
		Me.txtFolderFileList = New System.Windows.Forms.TextBox()
		Me.btnHtmlOutput = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'btnResultGridView
		'
		Me.btnResultGridView.Location = New System.Drawing.Point(17, 520)
		Me.btnResultGridView.Name = "btnResultGridView"
		Me.btnResultGridView.Size = New System.Drawing.Size(120, 23)
		Me.btnResultGridView.TabIndex = 2
		Me.btnResultGridView.Text = "リスト表示フォーム(O)"
		Me.btnResultGridView.UseVisualStyleBackColor = True
		'
		'btnOutput
		'
		Me.btnOutput.Location = New System.Drawing.Point(457, 520)
		Me.btnOutput.Name = "btnOutput"
		Me.btnOutput.Size = New System.Drawing.Size(60, 23)
		Me.btnOutput.TabIndex = 1
		Me.btnOutput.Text = "出力(T)"
		Me.btnOutput.UseVisualStyleBackColor = True
		'
		'txtFolderFileList
		'
		Me.txtFolderFileList.BackColor = System.Drawing.Color.Black
		Me.txtFolderFileList.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtFolderFileList.ForeColor = System.Drawing.Color.White
		Me.txtFolderFileList.Location = New System.Drawing.Point(17, 10)
		Me.txtFolderFileList.MinimumSize = New System.Drawing.Size(500, 500)
		Me.txtFolderFileList.Multiline = True
		Me.txtFolderFileList.Name = "txtFolderFileList"
		Me.txtFolderFileList.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.txtFolderFileList.Size = New System.Drawing.Size(500, 500)
		Me.txtFolderFileList.TabIndex = 0
		'
		'btnHtmlOutput
		'
		Me.btnHtmlOutput.Location = New System.Drawing.Point(366, 520)
		Me.btnHtmlOutput.Name = "btnHtmlOutput"
		Me.btnHtmlOutput.Size = New System.Drawing.Size(80, 23)
		Me.btnHtmlOutput.TabIndex = 3
		Me.btnHtmlOutput.Text = "html出力(H)"
		Me.btnHtmlOutput.UseVisualStyleBackColor = True
		'
		'frmResultText
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.Color.DeepSkyBlue
		Me.ClientSize = New System.Drawing.Size(534, 551)
		Me.Controls.Add(Me.btnHtmlOutput)
		Me.Controls.Add(Me.btnResultGridView)
		Me.Controls.Add(Me.btnOutput)
		Me.Controls.Add(Me.txtFolderFileList)
		Me.MinimumSize = New System.Drawing.Size(300, 39)
		Me.Name = "frmResultText"
		Me.Opacity = 0.0R
		Me.Text = "出力文字列フォーム"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
    Friend WithEvents txtFolderFileList As System.Windows.Forms.TextBox
	Friend WithEvents btnOutput As System.Windows.Forms.Button
	Friend WithEvents btnResultGridView As Button
	Friend WithEvents btnHtmlOutput As System.Windows.Forms.Button
End Class
