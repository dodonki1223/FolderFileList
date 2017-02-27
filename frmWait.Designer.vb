<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWait
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
		Me.lblMaking = New System.Windows.Forms.Label()
		Me.ssProgress = New System.Windows.Forms.StatusStrip()
		Me.tspbProgressRate = New System.Windows.Forms.ToolStripProgressBar()
		Me.tsslStatus = New System.Windows.Forms.ToolStripStatusLabel()
		Me.ssDisplayFoloderFile = New System.Windows.Forms.StatusStrip()
		Me.tsslProcessingFolderFile = New System.Windows.Forms.ToolStripStatusLabel()
		Me.ssProgress.SuspendLayout()
		Me.ssDisplayFoloderFile.SuspendLayout()
		Me.SuspendLayout()
		'
		'lblMaking
		'
		Me.lblMaking.AutoSize = True
		Me.lblMaking.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me.lblMaking.Location = New System.Drawing.Point(6, 9)
		Me.lblMaking.Name = "lblMaking"
		Me.lblMaking.Size = New System.Drawing.Size(266, 16)
		Me.lblMaking.TabIndex = 0
		Me.lblMaking.Text = "フォルダファイルリスト作成中．．．．．．"
		'
		'ssProgress
		'
		Me.ssProgress.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tspbProgressRate, Me.tsslStatus})
		Me.ssProgress.Location = New System.Drawing.Point(0, 59)
		Me.ssProgress.Name = "ssProgress"
		Me.ssProgress.Size = New System.Drawing.Size(284, 22)
		Me.ssProgress.TabIndex = 1
		Me.ssProgress.Text = "StatusStrip1"
		'
		'tspbProgressRate
		'
		Me.tspbProgressRate.Name = "tspbProgressRate"
		Me.tspbProgressRate.Size = New System.Drawing.Size(100, 16)
		'
		'tsslStatus
		'
		Me.tsslStatus.Name = "tsslStatus"
		Me.tsslStatus.Size = New System.Drawing.Size(136, 17)
		Me.tsslStatus.Spring = True
		Me.tsslStatus.Text = "0％完了"
		Me.tsslStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
		'
		'ssDisplayFoloderFile
		'
		Me.ssDisplayFoloderFile.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslProcessingFolderFile})
		Me.ssDisplayFoloderFile.Location = New System.Drawing.Point(0, 37)
		Me.ssDisplayFoloderFile.Name = "ssDisplayFoloderFile"
		Me.ssDisplayFoloderFile.Size = New System.Drawing.Size(284, 22)
		Me.ssDisplayFoloderFile.TabIndex = 3
		Me.ssDisplayFoloderFile.Text = "StatusStrip1"
		'
		'tsslProcessingFolderFile
		'
		Me.tsslProcessingFolderFile.Name = "tsslProcessingFolderFile"
		Me.tsslProcessingFolderFile.Size = New System.Drawing.Size(269, 17)
		Me.tsslProcessingFolderFile.Text = "処理中フォルダ・ファイルを表示が表示されます"
		'
		'frmWait
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(284, 81)
		Me.Controls.Add(Me.ssDisplayFoloderFile)
		Me.Controls.Add(Me.ssProgress)
		Me.Controls.Add(Me.lblMaking)
		Me.MinimumSize = New System.Drawing.Size(290, 100)
		Me.Name = "frmWait"
		Me.Opacity = 0.0R
		Me.Text = "リスト作成中"
		Me.ssProgress.ResumeLayout(False)
		Me.ssProgress.PerformLayout()
		Me.ssDisplayFoloderFile.ResumeLayout(False)
		Me.ssDisplayFoloderFile.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
    Friend WithEvents lblMaking As System.Windows.Forms.Label
	Friend WithEvents ssProgress As System.Windows.Forms.StatusStrip
	Friend WithEvents tspbProgressRate As System.Windows.Forms.ToolStripProgressBar
	Friend WithEvents tsslStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ssDisplayFoloderFile As StatusStrip
    Friend WithEvents tsslProcessingFolderFile As ToolStripStatusLabel
End Class
