Option Explicit On

Imports FolderFileList.DebugWatch

''' <summary>デバッグモード時の機能を提供する</summary>
''' <remarks></remarks>
<Conditional("DEBUG")> _
Public Class DebugMode

#Region "共通メソッド"

	''' <summary>フォームタイトルにデバッグモードの文言を追加</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Shared Sub SetDebugModeTitle(ByVal pForm As Form)

		pForm.Text = pForm.Text & "（デバッグモード）"

	End Sub

	''' <summary>時間の計測開始</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Shared Sub StartDebugWatch()

		'時間の計測をリセット
		'※２回目以降が前の最終計測時間後から使われてしまうため
		DebugWatch.Instance.ResetTime()

		'時間の計測開始
		DebugWatch.Instance.StartTime()

	End Sub

	''' <summary>時間の計測終了、経過時間を表示</summary>
	''' <param name="pTitle">ポップアップメッセージのタイトル</param>
	''' <param name="pMessage">ポップアップメッセージに表示するメッセージ</param>
	''' <remarks></remarks>
	<Conditional("DEBUG")>
	Public Shared Sub StopDebugWatchShowProcessingTime(ByVal pTitle As String, ByVal pMessage As String)

		'時間の計測終了
		DebugWatch.Instance.StopTime()

		'経過時間を表示
		DebugWatch.Instance.DebugPrintTimeToPopupMessage(pTitle, pMessage)

	End Sub

	''' <summary>Messenger風通知メッセージが閉じるまで処理を待機</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")>
	Public Shared Sub WaitTillClosingPopupMessage()

		'Messenger風通知メッセージが閉じるまで待機する
		System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime)

	End Sub

#End Region

#Region "メインフォーム"

	''' <summary>コマンドライン引数を表示</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Shared Sub ShowCommandLineArg()

		'コマンドライン引数を表示
		CommandLine.ShowCommandLineArgListToMessageBox()

	End Sub

#End Region

#Region "リスト表示フォーム"

	''' <summary>非表示のカラムを表示する</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Shared Sub DisplayAllColumnForFrmResultGridView(ByVal pResultGridView As frmResultGridView)

		'デバッグモード時は全てのカラムを表示する
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.No).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.Name).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.UpdateDate).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.FileSystemType).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.FileSystemTypeName).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.Extension).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.Size).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.SizeAndUnit).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.SizeLevel).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.SizeLevelName).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.SizeAndUnit).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.DirectoryLevel).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.ParentFolder).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.ParentFolderFullPath).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.UnderTargetFolder).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.FullPath).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.IsLastFileInFolder).Visible = True
		pResultGridView.dgvFolderFileList.Columns(FolderFileList.FolderFileListColumn.DispString).Visible = True

	End Sub

#End Region

End Class
