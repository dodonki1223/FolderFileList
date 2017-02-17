Option Explicit On

Imports FolderFileList.FolderFileList

''' <summary>フォルダファイルリストの出力文字列を表示するフォームを提供する</summary>
''' <remarks>フォームの共通処理のInterface（IFormCommonProcess）を実装する</remarks>
Public Class frmResultText
	Implements IFormCommonProcess

#Region "定数"

	''' <summary>出力ファイル形式</summary>
	''' <remarks></remarks>
	Private Const _cOutputFileFormat As String = "TXT"

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

				Case "T" 'Shift+T：「出力」ボタンをクリック

					btnOutput.PerformClick()

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
            Settings.Instance.TargetForm = Settings.FormType.Text
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

	''' <summary>出力ボタンクリックイベント</summary>
	''' <param name="sender">出力ボタン</param>
	''' <param name="e">Clickイベント</param>
	''' <remarks></remarks>
	Private Async Sub btnOutput_Click(sender As Object, e As EventArgs) Handles btnOutput.Click

		'メッセージボックスから押されたボタンにより処理を分岐
		Select Case _ShowDialogueMessage(_cOutputFileFormat)

			Case Windows.Forms.DialogResult.Yes

				'TXTファイル保存処理
				Call _SaveTextFile()

			Case Windows.Forms.DialogResult.No

				'TXTテキストクリップボードコピー
				Clipboard.SetText(txtFolderFileList.Text)

				'クリップボードにコピーしました通知を表示
				Dim mFrmPopupMessage As New frmPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageDetail)
				mFrmPopupMessage.Show()

				'非同期でMessenger風通知メッセージが非表示になるまで待機
				Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

		End Select

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
		txtFolderFileList.Text = _FolderFileList.OutPutText

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

						Case btnOutput.Name

							'出力ボタンの場合はフォームのサイズの変更分、位置を変更する
							mTargetControl.Location = New System.Drawing.Point(mTargetControl.Location + pChangedSize)

						Case btnResultGridView.Name

							'リスト表示フォームへボタンの場合はフォームのサイズの変更分、縦位置を変更する
							mTargetControl.Location = New System.Drawing.Point(mTargetControl.Location.X, mTargetControl.Location.Y + pChangedSize.Height)

					End Select

			End Select

		Next

	End Sub

	''' <summary>Alt+Tabのウインドウ切り替え時非表示</summary>
	''' <remarks></remarks>
	Private Sub _SetNotShowWindowToAltTab()

		'Alt+Tabに表示させない
		Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.ShowInTaskbar = False

	End Sub

#End Region

#Region "「出力文字列」出力用メソッド"

	''' <summary>TXT保存処理</summary>
	''' <remarks></remarks>
	Private Sub _SaveTextFile()

		'名前を付けて保存ダイアログを表示
		Dim mDailog As SaveFileDialog = _GetSaveAsDialog(_cOutputFileFormat)

		If mDailog.ShowDialog = Windows.Forms.DialogResult.OK Then

			'ファイルの保存処理
			Call _WriteTextToOutputFile(txtFolderFileList.Text, mDailog.FileName, cEncording.ShiftJis)

		End If

	End Sub

	''' <summary>ユーザーに対話メッセージボックス</summary>
	''' <param name="pFileFormat">出力形式</param>
	''' <returns>ダイアログボックスの戻り値</returns>
	''' <remarks></remarks>
	Private Function _ShowDialogueMessage(ByVal pFileFormat As String) As Windows.Forms.DialogResult

		'メッセージボックスを表示しユーザーに対話
		Dim mMsgBoxTitle As String = pFileFormat & "出力"
		Dim mMsgBoxText As String = "はい　：" & pFileFormat & "ファイルを保存します" & ControlChars.CrLf &
									"いいえ：クリップボードに" & pFileFormat & "形式のテキストを保存します"

		Return MessageBox.Show(mMsgBoxText, mMsgBoxTitle, MessageBoxButtons.YesNo)

	End Function

	''' <summary>名前を付けて保存ダイアログを取得</summary>
	''' <param name="pFileFormat ">ファイル形式</param>
	''' <returns>ファイル形式にあった設定の名前を付けて保存ダイアログ</returns>
	''' <remarks>区切り文字からその区切り文字用の名前をつけて保存ダイアログを作成し返す</remarks>
	Private Function _GetSaveAsDialog(ByVal pFileFormat As String) As SaveFileDialog

		'大文字、小文字の区切り文字を取得
		Dim mFileFormatUpperCase As String = pFileFormat
		Dim mFileFormatLowerCase As String = pFileFormat.ToLower

		'名前を付けて保存ダイアログを表示
		Dim mDailog As New SaveFileDialog
		With mDailog

			'デフォルトファイル設定（対象パスフォルダ名＋.区切り文字）
			.FileName = _FolderFileList.TargetPathFolderName & "." & mFileFormatLowerCase

			'表示ファイル設定
			.Filter = mFileFormatUpperCase & "ファイル(*." & mFileFormatLowerCase & ")|*." & mFileFormatLowerCase & "|すべてのファイル(*.*)|*.*"

		End With

		Return mDailog

	End Function

	''' <summary>文字列を指定ファイルに書き込む</summary>
	''' <param name="pWriteText">書き込むテキスト</param>
	''' <param name="pOutputPath">指定ファイル（出力先フルパス）</param>
	''' <param name="pEncoding">使用する文字エンコーディング</param>
	''' <remarks>指定されたファイルが存在しない場合はファイルを作成して書き込む
	'''          指定されたファイルが存在する場合はファイルを上書きして書き込む</remarks>
	Private Sub _WriteTextToOutputFile(ByVal pWriteText As String, ByVal pOutputPath As String, ByVal pEncoding As System.Text.Encoding)

		Using mSW As New System.IO.StreamWriter(pOutputPath, False, pEncoding)

			mSW.Write(pWriteText)

		End Using

	End Sub

#End Region

#Region "コマンド処理メソッド"

	''' <summary>コマンド処理を実行</summary>
	''' <remarks>コマンドタイプによりそれぞれの処理を行う</remarks>
	Private Async Sub _RunCommandProcess()

		'出力形式コマンドが指定なしの時
		If CommandLine.Instance.OutPut = CommandLine.OutPutType.None Then

			'フォームを透明状態から元に戻す
			Me.Opacity = 1

		Else

			Select Case CommandLine.Instance.OutPut

				Case CommandLine.OutPutType.ClipBoard

					'ウインドウをAlt+Tabに表示させない
					MyBase.SetShowHideAltTabWindow(AltTabType.Hide)

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

				Case CommandLine.OutPutType.SaveDialog

					'ALT+TABに表示させない
					Call _SetNotShowWindowToAltTab()

					'TXTファイル保存処理
					Call _SaveTextFile()

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