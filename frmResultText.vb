Option Explicit On

Imports FolderFileList.FolderFileList
Imports FolderFileList.CommandLine

''' <summary>フォルダファイルリストの出力文字列を表示するフォームを提供する</summary>
''' <remarks>フォームの共通処理のInterface（IFormCommonProcess）を実装する</remarks>
Public Class frmResultText
	Implements IFormCommonProcess

#Region "定数"

	''' <summary>Text出力ファイル形式</summary>
	''' <remarks></remarks>
	Private Const _cOutputTextFileFormat As String = "TXT"

	''' <summary>Html出力ファイル形式</summary>
	''' <remarks></remarks>
	Private Const _cOutputHtmlFileFormat As String = "HTML"

	''' <summary>出力文字列フォームで使用するメッセージを提供する</summary>
	''' <remarks></remarks>
	Private Class _cMessage

		''' <summary>Messenger風通知メッセージタイトル</summary>
		''' <remarks></remarks>
		Public Const NoticeMessageTitle As String = "出力文字列フォーム"

		''' <summary>Messenger風通知メッセージ内容</summary>
		''' <remarks></remarks>
		Public Const NoticeMessageDetail As String = "出力文字列をクリップボードにコピーしました"

	End Class

#End Region

#Region "変数"

	''' <summary>フォルダファイルリストデータ</summary>
	''' <remarks>FolderFileListプロパティにてデータをセットする</remarks>
	Private _FolderFileList As FolderFileList

	''' <summary>リサイズ前ウインドウサイズ</summary>
	''' <remarks></remarks>
	Private _WindowSize As System.Drawing.Size

	''' <summary>リサイズ後の変更分ウインドウサイズ</summary>
	''' <remarks></remarks>
	Private _ChangedWindowSize As System.Drawing.Size

	''' <summary>フォームのインスタンスを保持する変数</summary>
	''' <remarks></remarks>
	Private Shared _instance As frmResultText

#End Region

#Region "プロパティ"

	''' <summary>フォームにアクセスするためのプロパティ</summary>
	''' <remarks>※デザインパターンのSingletonパターンです
	'''            インスタンスがただ１つであることを保証する</remarks>
	Public Shared ReadOnly Property Instance() As frmResultText

		Get

			'インスタンスが作成されてなかったらインスタンスを作成
			If _instance Is Nothing Then _instance = New frmResultText

			Return _instance

		End Get

	End Property

	''' <summary>出力文字列フォームインスタンス存在プロパティ</summary>
	''' <remarks></remarks>
	Public Shared ReadOnly Property HasInstance() As Boolean

		Get

			If _instance Is Nothing Then

				Return False

			Else

				Return True

			End If

		End Get

	End Property

	''' <summary>フォルダファイルリストデータセットプロパティ</summary>
	''' <remarks></remarks>
	Public WriteOnly Property FolderFileListData() As FolderFileList

		Set(value As FolderFileList)

			_FolderFileList = value

		End Set

	End Property

#End Region

#Region "コンストラクタ"

	''' <summary>コンストラクタ（外部非公開）</summary>
	''' <remarks>引数付きのコンストラクタのみ許容させるため</remarks>
	Private Sub New()

		' この呼び出しはデザイナーで必要です。
		InitializeComponent()

		' InitializeComponent() 呼び出しの後で初期化を追加します。

	End Sub

#End Region

#Region "イベント"

#Region "フォーム"

	''' <summary>フォームロードイベント</summary>
	''' <param name="sender">Formオブジェクト</param>
	''' <param name="e">Loadイベント</param>
	''' <remarks></remarks>
	Private Sub frmResultText_Load(sender As Object, e As EventArgs) Handles Me.Load

		'初期コントロール設定
		_SetInitialControlInScreen()

	End Sub

	''' <summary>フォームShownイベント</summary>
	''' <param name="sender">Formオブジェクト</param>
	''' <param name="e">Shownイベント</param>
	''' <remarks></remarks>
	Private Sub frmResultText_Shown(sender As Object, e As EventArgs) Handles Me.Shown

		'処理タイプにより処理を分岐
		Select Case CommandLine.Instance.ProcessMode

			Case CommandLine.ProcessType.Command

				'コマンド処理を実行
				Call _RunCommandProcess()

			Case Else

				'フォームを透明状態から元に戻す
				Me.Opacity = 1

		End Select

	End Sub

	''' <summary>フォームリサイズイベント</summary>
	''' <param name="sender">Formオブジェクト</param>
	''' <param name="e">Resizeイベント</param>
	''' <remarks></remarks>
	Private Sub frmResultText_Resize(sender As Object, e As EventArgs) Handles Me.Resize

		'ウインドウの状態が最小化以外の時
		If Me.WindowState <> FormWindowState.Minimized Then

			'リサイズ前ウインドウサイズから変更された分のサイズを取得
			_ChangedWindowSize = Me.Size - _WindowSize

			'コントロールのサイズ変更・再配置
			_ResizeAndRealignmentControls(_ChangedWindowSize)

			'リサイズ後のウインドウサイズをセット
			_WindowSize = Me.Size

		End If

	End Sub

	''' <summary>フォームのKeyDownイベント</summary>
	''' <param name="sender">Formオブジェクト</param>
	''' <param name="e">KeyDownイベント</param>
	''' <remarks></remarks>
	Private Sub frmResultText_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

		If e.Shift Then

			'-------------------------------
			' Shiftキーのショートカットキー
			'-------------------------------
			Select Case e.KeyCode.ToString

				Case "O" 'Shift+O：「リスト表示フォーム」ボタンをクリック

					btnResultGridView.PerformClick()

				Case "H" 'Shift+H：「html出力」ボタンをクリック

					btnHtmlOutput.PerformClick()

				Case "T" 'Shift+T：「Text出力」ボタンをクリック

					btnTextOutput.PerformClick()

			End Select

		End If

	End Sub

	''' <summary>フォームのClosedイベント</summary>
	''' <param name="sender">Formオブジェクト</param>
	''' <param name="e">FormClosedイベント</param>
	''' <remarks>フォームが閉じられようとした際に閉じる処理をキャンセルさせる時に使用するのがFormClosing
	'''          フォームが閉じる前提で後処理するのがFormClosed                                          </remarks>
	Private Sub frmResultText_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

		'インスタンスへの参照を破棄
		'※もう一度、表示する時はインスタンスを再作成させるため
		_instance = Nothing

		'リスト表示フォームのインスタンスが存在したら
		'※インスタンスがあるってことはメインフォームの閉じる処理でないってこと
		If frmResultGridView.HasInstance Then

			'設定ファイルへ書き込み処理
			Settings.Instance.TargetForm = CommandLine.FormType.Text
			Settings.SaveToXmlFile()

			'メインフォームに設定ファイルの内容を適用する
			DirectCast(Me.Owner, frmMain).SetXmlFileSettingToScreen()

			Select Case CommandLine.Instance.ProcessMode

				Case CommandLine.ProcessType.Normarl

					'メインフォームをモードレスで表示
					Me.Owner.Show()

				Case CommandLine.ProcessType.Command

					'メインフォームのリソースを破棄する
					'※Closeをすると無限ループしてしまうのでDisposeで対応（これでいいのか不明……）
					'  メインフォームを表示させず閉じるため
					Me.Owner.Dispose()

			End Select

		End If

		'リスト表示フォームのインスタンス変数を破棄
		'※出力文字列フォームが閉じるタイミングでリスト表示フォームも閉じる仕様のため
		frmResultGridView.DisposeInstance()

	End Sub

#End Region

#Region "ボタン"

	''' <summary>Text出力ボタンクリックイベント</summary>
	''' <param name="sender">Text出力ボタン</param>
	''' <param name="e">Clickイベント</param>
	''' <remarks></remarks>
	Private Async Sub btnTextOutput_Click(sender As Object, e As EventArgs) Handles btnTextOutput.Click

		'メッセージボックスから押されたボタンにより処理を分岐
		Select Case MyBase.ShowDialogueMessage(_cOutputTextFileFormat)

			Case Windows.Forms.DialogResult.Yes

				'TXTファイル保存処理
				Call _SaveOutputFile(OutputFileFormat.TEXT, _FolderFileList.TargetPathFolderName, cEncording.ShiftJis)

			Case Windows.Forms.DialogResult.No

				'TXTテキストクリップボードコピー
				Clipboard.SetText(_FolderFileList.OutputText)

				'クリップボードにコピーしました通知を表示
				Dim mFrmPopupMessage As New frmPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageDetail)
				mFrmPopupMessage.Show()

				'非同期でMessenger風通知メッセージが非表示になるまで待機
				Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

		End Select

	End Sub

	''' <summary>html出力ボタンクリックイベント</summary>
	''' <param name="sender">html出力ボタン</param>
	''' <param name="e">Clickイベント</param>
	''' <remarks></remarks>
	Private Sub btnHtmlOutput_Click(sender As Object, e As EventArgs) Handles btnHtmlOutput.Click

		'Htmlファイル保存処理
		Call _SaveOutputFile(OutputFileFormat.HTML, _FolderFileList.TargetPathFolderName, cEncording.UTF8)

	End Sub

	''' <summary>リスト表示フォームへボタンクリックイベント</summary>
	''' <param name="sender">リスト表示フォームへボタン</param>
	''' <param name="e">Clickイベント</param>
	''' <remarks></remarks>
	Private Sub btnResultGridView_Click(sender As Object, e As EventArgs) Handles btnResultGridView.Click

		'出力文字列フォームを非表示
		Me.Hide()

		'リスト表示フォームをモードレスで表示
		frmResultGridView.Instance.Show()

	End Sub

#End Region

#End Region

#Region "メソッド"

#Region "コントロール設定メソッド"

	''' <summary>初期画面コントロール設定</summary>
	''' <remarks></remarks>
	Private Sub _SetInitialControlInScreen() Implements IFormCommonProcess._SetInitialControlInScreen

		'フォームを透明状態にする
		Me.Opacity = 0

		'フォームでもキーイベントを取得可に設定
		'※ショートカットキーを設定出来るようにするため
		Me.KeyPreview = True

		'出力用文字列テキストボックスを編集不可に
		txtFolderFileList.ReadOnly = True

		'出力用文字列をセットを画面にセット
		txtFolderFileList.Text = _FolderFileList.OutputText

		'フォームのタイトルにコマンドモード文言を追加
		CommandLine.SetCommandModeToTitle(Me)

		'フォームのタイトルにデバッグモード文言を追加
		'※デバッグモード時のみ実行される
		DebugMode.SetDebugModeTitle(Me)

	End Sub

	''' <summary>コントロールのサイズ変更・再配置</summary>
	''' <param name="pChangedSize">フォームの変更後サイズ</param>
	''' <remarks>全てのコントロールを取得しコントロールごと処理を行う</remarks>
	Private Sub _ResizeAndRealignmentControls(pChangedSize As Size) Implements IFormCommonProcess._ResizeAndRealignmentControls

		'全てのコントロールを取得
		Dim mAllControls As Control() = GetControlsInTarget(Me)

		For Each mTargetControl As Control In mAllControls

			Select Case True

				Case TypeOf mTargetControl Is TextBox

					'テキストボックスの場合はフォームのサイズの変更分、サイズを変更する
					mTargetControl.Size = New System.Drawing.Size(mTargetControl.Size + pChangedSize)

				Case TypeOf mTargetControl Is Button

					Select Case mTargetControl.Name

						Case btnHtmlOutput.Name, btnTextOutput.Name

							'Html出力ボタン、Text出力ボタンの場合はフォームのサイズの変更分、位置を変更する
							mTargetControl.Location = New System.Drawing.Point(mTargetControl.Location + pChangedSize)

						Case btnResultGridView.Name

							'リスト表示フォームへボタンの場合はフォームのサイズの変更分、縦位置を変更する
							mTargetControl.Location = New System.Drawing.Point(mTargetControl.Location.X, mTargetControl.Location.Y + pChangedSize.Height)

					End Select

			End Select

		Next

	End Sub

#End Region

#Region "「Html・TEXT」出力用メソッド"

	''' <summary>出力ファイルの保存処理</summary>
	''' <param name="pFileFormat">出力形式</param>
	''' <param name="pDefalutFileName">保存ファイルのデフォルトのファイル名</param>
	''' <param name="pEncording">エンコード</param>
	''' <remarks></remarks>
	Private Sub _SaveOutputFile(ByVal pFileFormat As OutputFileFormat, ByVal pDefalutFileName As String, ByVal pEncording As System.Text.Encoding)

		'名前を付けて保存ダイアログを表示
		Dim mDailog As SaveFileDialog = MyBase.GetSaveAsDialog(pDefalutFileName, pFileFormat.ToString.ToLower)

		'名前をつけて保存ダイアログでOKが押されたら
		If mDailog.ShowDialog = Windows.Forms.DialogResult.OK Then

			'ファイルの保存処理
			MyBase.WriteTextToOutputFile(_GetOutputFileText(pFileFormat), mDailog.FileName, pEncording)

		End If

	End Sub

	''' <summary>出力用ファイルのテキストデータを取得</summary>
	''' <param name="pFileFormat">出力形式</param>
	''' <returns>対象出力形式のテキストデータ</returns>
	''' <remarks></remarks>
	Private Function _GetOutputFileText(ByVal pFileFormat As OutputFileFormat) As String

		Dim mOutputText As String = String.Empty

		'出力形式ごと処理を分岐
		Select Case pFileFormat

			Case OutputFileFormat.TEXT

				'フォルダファイルリストの出力文字列を取得
				mOutputText = _FolderFileList.OutputText

			Case OutputFileFormat.HTML

				'フォルダファイルリストから出力用のHtml文字列を取得
				Dim mOutputHtml As New OutputHtml(_FolderFileList, CommandLine.FormType.Text)
				mOutputText = mOutputHtml.HtmlSentence

		End Select

		Return mOutputText

	End Function

#End Region

#Region "コマンド処理メソッド"

	''' <summary>コマンド処理を実行</summary>
	''' <remarks>コマンドタイプによりそれぞれの処理を行う</remarks>
	Private Async Sub _RunCommandProcess()

		'出力形式コマンドが指定なしの時
		If CommandLine.Instance.Output = CommandLine.OutputType.None Then

			'フォームを透明状態から元に戻す
			Me.Opacity = 1

		Else

			'ウインドウをAlt+Tabに表示させない
			MyBase.SetShowHideAltTabWindow(AltTabType.Hide)

			Select Case CommandLine.Instance.Output

				Case CommandLine.OutputType.ClipBoard

					'TXTテキストクリップボードコピー
					Clipboard.SetText(txtFolderFileList.Text)

					'クリップボードにコピーしました通知を表示
					Dim mFrmPopupMessage As New frmPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageDetail)
					mFrmPopupMessage.Show()

					'非同期でMessenger風通知メッセージが非表示になるまで待機
					Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

					'メインフォームのリソースを破棄する
					'※Closeをすると無限ループしてしまうのでDisposeで対応（これでいいのか不明……）
					'  メインフォームを表示させず閉じるため
					Me.Owner.Dispose()

				Case CommandLine.OutputType.SaveDialog

					'拡張子がHtmlの場合はエンコードを「UTF-8」に変更
					'※デフォルトは「Shift-Jis」
					Dim mSaveFileEncord As System.Text.Encoding = cEncording.ShiftJis
					If CommandLine.Instance.Extension = CommandLine.OutputFileFormat.HTML Then mSaveFileEncord = cEncording.UTF8

					'出力用テキスト保存処理
					Call _SaveOutputFile(CommandLine.Instance.Extension, _FolderFileList.TargetPathFolderName, mSaveFileEncord)

					'メインフォームのリソースを破棄する
					'※Closeをすると無限ループしてしまうのでDisposeで対応（これでいいのか不明……）
					'  メインフォームを表示させず閉じるため
					Me.Owner.Dispose()

			End Select

		End If

	End Sub

#End Region

#End Region

#Region "外部公開メソッド"

	''' <summary>インスタンスを保持する変数の破棄処理</summary>
	''' <remarks></remarks>
	Public Shared Sub DisposeInstance()

		_instance = Nothing

	End Sub

#End Region

End Class