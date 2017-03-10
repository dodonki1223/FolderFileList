<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmResultGridView
    Inherits frmCommon

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
		Me.lblRangeStart = New System.Windows.Forms.Label()
		Me.txtRangeStart = New System.Windows.Forms.TextBox()
		Me.txtMaxSearchCount = New System.Windows.Forms.TextBox()
		Me.lblRangeEnd = New System.Windows.Forms.Label()
		Me.txtRangeEnd = New System.Windows.Forms.TextBox()
		Me.lblMaxSearchCount = New System.Windows.Forms.Label()
		Me.lblRange = New System.Windows.Forms.Label()
		Me.btnNextPage = New System.Windows.Forms.Button()
		Me.btnPrevPage = New System.Windows.Forms.Button()
		Me.grpFileSize = New System.Windows.Forms.GroupBox()
		Me.chkFileSizeLevel = New System.Windows.Forms.CheckBox()
		Me.cmbFileSizeLevel = New System.Windows.Forms.ComboBox()
		Me.btnResultTextForm = New System.Windows.Forms.Button()
		Me.btnTsvOutput = New System.Windows.Forms.Button()
		Me.btnCsvOutput = New System.Windows.Forms.Button()
		Me.grpName = New System.Windows.Forms.GroupBox()
		Me.btnNameSearch = New System.Windows.Forms.Button()
		Me.txtName = New System.Windows.Forms.TextBox()
		Me.grpExtension = New System.Windows.Forms.GroupBox()
		Me.cmbExtension = New System.Windows.Forms.ComboBox()
		Me.grpDisplayTarget = New System.Windows.Forms.GroupBox()
		Me.rbtnFile = New System.Windows.Forms.RadioButton()
		Me.rbtnFolder = New System.Windows.Forms.RadioButton()
		Me.rbtnAll = New System.Windows.Forms.RadioButton()
		Me.dgvFolderFileList = New System.Windows.Forms.DataGridView()
		Me.btnHtmlOutput = New System.Windows.Forms.Button()
		Me.grpFileSize.SuspendLayout()
		Me.grpName.SuspendLayout()
		Me.grpExtension.SuspendLayout()
		Me.grpDisplayTarget.SuspendLayout()
		CType(Me.dgvFolderFileList, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'lblRangeStart
		'
		Me.lblRangeStart.AutoSize = True
		Me.lblRangeStart.Location = New System.Drawing.Point(785, 65)
		Me.lblRangeStart.Name = "lblRangeStart"
		Me.lblRangeStart.Size = New System.Drawing.Size(17, 12)
		Me.lblRangeStart.TabIndex = 20
		Me.lblRangeStart.Text = "件"
		'
		'txtRangeStart
		'
		Me.txtRangeStart.Location = New System.Drawing.Point(740, 62)
		Me.txtRangeStart.Name = "txtRangeStart"
		Me.txtRangeStart.ReadOnly = True
		Me.txtRangeStart.Size = New System.Drawing.Size(40, 19)
		Me.txtRangeStart.TabIndex = 19
		Me.txtRangeStart.Text = "2000"
		Me.txtRangeStart.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		'
		'txtMaxSearchCount
		'
		Me.txtMaxSearchCount.Location = New System.Drawing.Point(664, 62)
		Me.txtMaxSearchCount.Name = "txtMaxSearchCount"
		Me.txtMaxSearchCount.ReadOnly = True
		Me.txtMaxSearchCount.Size = New System.Drawing.Size(40, 19)
		Me.txtMaxSearchCount.TabIndex = 18
		Me.txtMaxSearchCount.Text = "34232"
		Me.txtMaxSearchCount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		'
		'lblRangeEnd
		'
		Me.lblRangeEnd.AutoSize = True
		Me.lblRangeEnd.Location = New System.Drawing.Point(859, 65)
		Me.lblRangeEnd.Name = "lblRangeEnd"
		Me.lblRangeEnd.Size = New System.Drawing.Size(17, 12)
		Me.lblRangeEnd.TabIndex = 16
		Me.lblRangeEnd.Text = "件"
		'
		'txtRangeEnd
		'
		Me.txtRangeEnd.Location = New System.Drawing.Point(816, 62)
		Me.txtRangeEnd.Name = "txtRangeEnd"
		Me.txtRangeEnd.ReadOnly = True
		Me.txtRangeEnd.Size = New System.Drawing.Size(40, 19)
		Me.txtRangeEnd.TabIndex = 15
		Me.txtRangeEnd.Text = "3000"
		Me.txtRangeEnd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		'
		'lblMaxSearchCount
		'
		Me.lblMaxSearchCount.AutoSize = True
		Me.lblMaxSearchCount.Location = New System.Drawing.Point(710, 65)
		Me.lblMaxSearchCount.Name = "lblMaxSearchCount"
		Me.lblMaxSearchCount.Size = New System.Drawing.Size(29, 12)
		Me.lblMaxSearchCount.TabIndex = 14
		Me.lblMaxSearchCount.Text = "件中"
		'
		'lblRange
		'
		Me.lblRange.AutoSize = True
		Me.lblRange.Location = New System.Drawing.Point(803, 65)
		Me.lblRange.Name = "lblRange"
		Me.lblRange.Size = New System.Drawing.Size(11, 12)
		Me.lblRange.TabIndex = 13
		Me.lblRange.Text = "-"
		'
		'btnNextPage
		'
		Me.btnNextPage.Location = New System.Drawing.Point(939, 60)
		Me.btnNextPage.Name = "btnNextPage"
		Me.btnNextPage.Size = New System.Drawing.Size(55, 23)
		Me.btnNextPage.TabIndex = 10
		Me.btnNextPage.Text = "次へ(N)"
		Me.btnNextPage.UseVisualStyleBackColor = True
		'
		'btnPrevPage
		'
		Me.btnPrevPage.Location = New System.Drawing.Point(879, 60)
		Me.btnPrevPage.Name = "btnPrevPage"
		Me.btnPrevPage.Size = New System.Drawing.Size(55, 23)
		Me.btnPrevPage.TabIndex = 9
		Me.btnPrevPage.Text = "前へ(P)"
		Me.btnPrevPage.UseVisualStyleBackColor = True
		'
		'grpFileSize
		'
		Me.grpFileSize.Controls.Add(Me.chkFileSizeLevel)
		Me.grpFileSize.Controls.Add(Me.cmbFileSizeLevel)
		Me.grpFileSize.Location = New System.Drawing.Point(480, 12)
		Me.grpFileSize.Name = "grpFileSize"
		Me.grpFileSize.Size = New System.Drawing.Size(220, 45)
		Me.grpFileSize.TabIndex = 2
		Me.grpFileSize.TabStop = False
		Me.grpFileSize.Text = "ファイルサイズ指定(D)"
		'
		'chkFileSizeLevel
		'
		Me.chkFileSizeLevel.AutoSize = True
		Me.chkFileSizeLevel.Location = New System.Drawing.Point(82, 17)
		Me.chkFileSizeLevel.Name = "chkFileSizeLevel"
		Me.chkFileSizeLevel.Size = New System.Drawing.Size(136, 16)
		Me.chkFileSizeLevel.TabIndex = 1
		Me.chkFileSizeLevel.Text = "レベルごと色を付ける(L)"
		Me.chkFileSizeLevel.UseVisualStyleBackColor = True
		'
		'cmbFileSizeLevel
		'
		Me.cmbFileSizeLevel.BackColor = System.Drawing.SystemColors.Window
		Me.cmbFileSizeLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbFileSizeLevel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.cmbFileSizeLevel.FormattingEnabled = True
		Me.cmbFileSizeLevel.Location = New System.Drawing.Point(6, 15)
		Me.cmbFileSizeLevel.Name = "cmbFileSizeLevel"
		Me.cmbFileSizeLevel.Size = New System.Drawing.Size(70, 20)
		Me.cmbFileSizeLevel.TabIndex = 0
		'
		'btnResultTextForm
		'
		Me.btnResultTextForm.Location = New System.Drawing.Point(17, 530)
		Me.btnResultTextForm.Name = "btnResultTextForm"
		Me.btnResultTextForm.Size = New System.Drawing.Size(130, 23)
		Me.btnResultTextForm.TabIndex = 5
		Me.btnResultTextForm.Text = "出力文字列フォーム(O)"
		Me.btnResultTextForm.UseVisualStyleBackColor = True
		'
		'btnTsvOutput
		'
		Me.btnTsvOutput.Location = New System.Drawing.Point(914, 530)
		Me.btnTsvOutput.Name = "btnTsvOutput"
		Me.btnTsvOutput.Size = New System.Drawing.Size(80, 23)
		Me.btnTsvOutput.TabIndex = 7
		Me.btnTsvOutput.Text = "TSV出力(T)"
		Me.btnTsvOutput.UseVisualStyleBackColor = True
		'
		'btnCsvOutput
		'
		Me.btnCsvOutput.Location = New System.Drawing.Point(829, 530)
		Me.btnCsvOutput.Name = "btnCsvOutput"
		Me.btnCsvOutput.Size = New System.Drawing.Size(80, 23)
		Me.btnCsvOutput.TabIndex = 6
		Me.btnCsvOutput.Text = "CSV出力(C)"
		Me.btnCsvOutput.UseVisualStyleBackColor = True
		'
		'grpName
		'
		Me.grpName.Controls.Add(Me.btnNameSearch)
		Me.grpName.Controls.Add(Me.txtName)
		Me.grpName.Location = New System.Drawing.Point(706, 12)
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
		'grpExtension
		'
		Me.grpExtension.Controls.Add(Me.cmbExtension)
		Me.grpExtension.Location = New System.Drawing.Point(363, 12)
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
		Me.cmbExtension.Location = New System.Drawing.Point(6, 15)
		Me.cmbExtension.Name = "cmbExtension"
		Me.cmbExtension.Size = New System.Drawing.Size(95, 20)
		Me.cmbExtension.TabIndex = 0
		'
		'grpDisplayTarget
		'
		Me.grpDisplayTarget.Controls.Add(Me.rbtnFile)
		Me.grpDisplayTarget.Controls.Add(Me.rbtnFolder)
		Me.grpDisplayTarget.Controls.Add(Me.rbtnAll)
		Me.grpDisplayTarget.Location = New System.Drawing.Point(17, 12)
		Me.grpDisplayTarget.Name = "grpDisplayTarget"
		Me.grpDisplayTarget.Size = New System.Drawing.Size(340, 45)
		Me.grpDisplayTarget.TabIndex = 0
		Me.grpDisplayTarget.TabStop = False
		Me.grpDisplayTarget.Text = "表示対象"
		'
		'rbtnFile
		'
		Me.rbtnFile.AutoSize = True
		Me.rbtnFile.Location = New System.Drawing.Point(220, 18)
		Me.rbtnFile.Name = "rbtnFile"
		Me.rbtnFile.Size = New System.Drawing.Size(117, 16)
		Me.rbtnFile.TabIndex = 2
		Me.rbtnFile.Text = "ファイルのみ表示(F)"
		Me.rbtnFile.UseVisualStyleBackColor = True
		'
		'rbtnFolder
		'
		Me.rbtnFolder.AutoSize = True
		Me.rbtnFolder.Location = New System.Drawing.Point(96, 18)
		Me.rbtnFolder.Name = "rbtnFolder"
		Me.rbtnFolder.Size = New System.Drawing.Size(118, 16)
		Me.rbtnFolder.TabIndex = 1
		Me.rbtnFolder.Text = "フォルダのみ表示(K)"
		Me.rbtnFolder.UseVisualStyleBackColor = True
		'
		'rbtnAll
		'
		Me.rbtnAll.AutoSize = True
		Me.rbtnAll.Checked = True
		Me.rbtnAll.Location = New System.Drawing.Point(6, 18)
		Me.rbtnAll.Name = "rbtnAll"
		Me.rbtnAll.Size = New System.Drawing.Size(84, 16)
		Me.rbtnAll.TabIndex = 0
		Me.rbtnAll.TabStop = True
		Me.rbtnAll.Text = "全て表示(A)"
		Me.rbtnAll.UseVisualStyleBackColor = True
		'
		'dgvFolderFileList
		'
		Me.dgvFolderFileList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgvFolderFileList.Location = New System.Drawing.Point(17, 89)
		Me.dgvFolderFileList.Name = "dgvFolderFileList"
		Me.dgvFolderFileList.RowTemplate.Height = 21
		Me.dgvFolderFileList.Size = New System.Drawing.Size(976, 429)
		Me.dgvFolderFileList.TabIndex = 4
		'
		'btnHtmlOutput
		'
		Me.btnHtmlOutput.Location = New System.Drawing.Point(743, 530)
		Me.btnHtmlOutput.Name = "btnHtmlOutput"
		Me.btnHtmlOutput.Size = New System.Drawing.Size(80, 23)
		Me.btnHtmlOutput.TabIndex = 21
		Me.btnHtmlOutput.Text = "html出力(H)"
		Me.btnHtmlOutput.UseVisualStyleBackColor = True
		'
		'frmResultGridView
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1009, 562)
		Me.Controls.Add(Me.btnHtmlOutput)
		Me.Controls.Add(Me.lblRangeStart)
		Me.Controls.Add(Me.txtRangeStart)
		Me.Controls.Add(Me.txtMaxSearchCount)
		Me.Controls.Add(Me.lblRangeEnd)
		Me.Controls.Add(Me.txtRangeEnd)
		Me.Controls.Add(Me.lblMaxSearchCount)
		Me.Controls.Add(Me.lblRange)
		Me.Controls.Add(Me.btnNextPage)
		Me.Controls.Add(Me.btnPrevPage)
		Me.Controls.Add(Me.grpFileSize)
		Me.Controls.Add(Me.btnResultTextForm)
		Me.Controls.Add(Me.btnTsvOutput)
		Me.Controls.Add(Me.btnCsvOutput)
		Me.Controls.Add(Me.grpName)
		Me.Controls.Add(Me.grpExtension)
		Me.Controls.Add(Me.grpDisplayTarget)
		Me.Controls.Add(Me.dgvFolderFileList)
		Me.KeyPreview = True
		Me.MinimumSize = New System.Drawing.Size(300, 39)
		Me.Name = "frmResultGridView"
		Me.Opacity = 0.0R
		Me.Text = "リスト表示フォーム"
		Me.grpFileSize.ResumeLayout(False)
		Me.grpFileSize.PerformLayout()
		Me.grpName.ResumeLayout(False)
		Me.grpName.PerformLayout()
		Me.grpExtension.ResumeLayout(False)
		Me.grpDisplayTarget.ResumeLayout(False)
		Me.grpDisplayTarget.PerformLayout()
		CType(Me.dgvFolderFileList, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents dgvFolderFileList As System.Windows.Forms.DataGridView
	Friend WithEvents grpDisplayTarget As System.Windows.Forms.GroupBox
	Friend WithEvents rbtnFile As System.Windows.Forms.RadioButton
	Friend WithEvents rbtnFolder As System.Windows.Forms.RadioButton
	Friend WithEvents rbtnAll As System.Windows.Forms.RadioButton
	Friend WithEvents grpExtension As System.Windows.Forms.GroupBox
	Friend WithEvents cmbExtension As System.Windows.Forms.ComboBox
	Friend WithEvents grpName As System.Windows.Forms.GroupBox
	Friend WithEvents txtName As System.Windows.Forms.TextBox
	Friend WithEvents btnCsvOutput As System.Windows.Forms.Button
	Friend WithEvents btnTsvOutput As System.Windows.Forms.Button
    Friend WithEvents btnNameSearch As System.Windows.Forms.Button
    Friend WithEvents btnResultTextForm As Button
    Friend WithEvents grpFileSize As GroupBox
    Friend WithEvents chkFileSizeLevel As CheckBox
    Friend WithEvents cmbFileSizeLevel As ComboBox
    Friend WithEvents btnPrevPage As System.Windows.Forms.Button
    Friend WithEvents btnNextPage As System.Windows.Forms.Button
    Friend WithEvents lblRange As System.Windows.Forms.Label
    Friend WithEvents lblMaxSearchCount As System.Windows.Forms.Label
    Friend WithEvents txtRangeEnd As System.Windows.Forms.TextBox
    Friend WithEvents lblRangeEnd As System.Windows.Forms.Label
    Friend WithEvents txtMaxSearchCount As System.Windows.Forms.TextBox
    Friend WithEvents lblRangeStart As System.Windows.Forms.Label
	Friend WithEvents txtRangeStart As System.Windows.Forms.TextBox
	Friend WithEvents btnHtmlOutput As System.Windows.Forms.Button
End Class
