Option Explicit On

Imports FolderFileList.FolderFileList
Imports FolderFileList.CommandLine
Imports System.Linq

''' <summary>フォルダファイルリストをGridViewに表示するフォームを提供する</summary>
''' <remarks>フォームの共通処理のInterface（IFormCommonProcess）を実装する</remarks>
Public Class frmResultGridView
    Implements IFormCommonProcess

#Region "定数"

    ''' <summary>フォルダファイルリストのヘッダー情報を提供する</summary>
    ''' <remarks></remarks>
    Public Class cFolderFileListHeaders

        ''' <summary>区切り文字用ヘッダー名</summary>
        ''' <remarks></remarks>
        Public Enum DelimiterHeaderName

            ファイル名

            更新日時

            ファイルシステムタイプ

            拡張子

            サイズ

            サイズレベル

            親フォルダ

            対象パス以下

            フルパス

        End Enum

        ''' <summary>ヘッダー名を取得</summary>
        ''' <returns>ヘッダー名</returns>
        ''' <remarks>ヘッダー名をArrayListで返す</remarks>
		Public Shared Function HeaderNames() As ArrayList

			Dim mHeaderNames As New ArrayList

			'ヘッダー用の名前をArrayListに追加していく
			For Each mColumnName As String In System.Enum.GetNames(GetType(FolderFileListJapaneseColumn))

				mHeaderNames.Add(mColumnName)

			Next

			Return mHeaderNames

		End Function

		''' <summary>区切り文字用ヘッダー名を取得</summary>
		''' <returns>区切り文字用ヘッダー名</returns>
		''' <remarks>ヘッダー名をArrayListで返す</remarks>
		Public Shared Function DelimiterHeaderNames() As ArrayList

			Dim mHeaderNames As New ArrayList

			'ヘッダー用の名前をArrayListに追加していく
			For Each mColumnName As String In System.Enum.GetNames(GetType(DelimiterHeaderName))

				mHeaderNames.Add(mColumnName)

			Next

			Return mHeaderNames

		End Function

	End Class

	''' <summary>リスト表示フォームで使用するメッセージを提供する</summary>
	''' <remarks></remarks>
	Private Class _cMessage

		''' <summary>Messenger風通知メッセージタイトル</summary>
		''' <remarks></remarks>
		Public Const NoticeMessageTitle As String = "リスト表示フォーム"

        ''' <summary>Messenger風通知メッセージ内容：クリップボードにコピーしました</summary>
        ''' <remarks>コマンド処理用</remarks>
        Public Shared NoticeCommandMessageCopyToClipBoard As String = "リスト表示の" & CommandLine.Instance.Extension.ToString & "形式を" _
                                                                    & Environment.NewLine & "クリップボードにコピーしました"

        ''' <summary>Messenger風通知メッセージ内容：クリップボードにコピーしました</summary>
        ''' <param name="pFileFormat">ファイルの出力形式</param>
        ''' <returns>ファイル出力形式に対応したクリップボードにコピーしました文言</returns>
        ''' <remarks></remarks>
        Public Shared Function NoticeMessageCopyToClipBoard(ByVal pFileFormat As OutputFileFormat) As String

			Return "リスト表示の" & pFileFormat.ToString & "形式を" & Environment.NewLine _
				 & "クリップボードにコピーしました"

        End Function

        ''' <summary>Messenger風通知メッセージ：保存したファイルを実行しました</summary>
        ''' <remarks></remarks>
        Public Const NoticeMessageRunSaveFile As String = "保存したファイルを実行しました"

    End Class

	''' <summary>データグリッドビューで使用するカラー</summary>
	''' <remarks></remarks>
	Private Class _cGridViewColor

		''' <summary>実行出来るカラムのフォントカラー</summary>
		''' <remarks>Blue</remarks>
		Public Shared ReadOnly RunColumnFont As System.Drawing.Color = Color.Blue

		''' <summary>ヘッダーセル昇順カラー</summary>
		''' <remarks>HotPink</remarks>
		Public Shared ReadOnly HeaderAsceding As System.Drawing.Color = Color.HotPink

		''' <summary>ヘッダーセル降順カラー</summary>
		''' <remarks>LightSkyBlue</remarks>
		Public Shared ReadOnly HeaderDesceding As System.Drawing.Color = Color.LightSkyBlue

	End Class

	''' <summary>Html出力ファイル形式</summary>
	''' <remarks></remarks>
	Private Const _cOutputHtmlFileFormat As String = "HTML"

	''' <summary>１ページ内に表示できるファイル数</summary>
	''' <remarks>フォルダファイルリストGridViewの1ページに表示される最大件数（デフォルト値）</remarks>
	Public Const cGridViewDataMaxCountInPage As Integer = 1000

#End Region

#Region "変数"

	''' <summary>フォルダファイルリストデータ</summary>
	''' <remarks></remarks>
	Private _FolderFileList As FolderFileList

	''' <summary>リサイズ前ウインドウサイズ</summary>
	''' <remarks></remarks>
	Private _WindowSize As System.Drawing.Size

	''' <summary>リサイズ後の変更分ウインドウサイズ</summary>
	''' <remarks></remarks>
	Private _ChangedWindowSize As System.Drawing.Size

	''' <summary>GridViewに現在表示されている情報</summary>
	''' <remarks></remarks>
	Private _CurrentGridView As _GridViewState

	''' <summary>GridViewのSizeAndUnitヘッダーの状態</summary>
	''' <remarks>デフォルトは「なし」</remarks>
	Private _SortOrderFromByteSizeAndUnit As SortOrder = SortOrder.None

	''' <summary>現在表示しているページ番号</summary>
	''' <remarks>GridViewのページング処理で使用する</remarks>
	Private _CurrentPage As Integer = 1

	''' <summary>１ページ内に表示できるファイル数</summary>
	''' <remarks>フォルダファイルリストGridViewの1ページに表示される最大件数</remarks>
	Private _GridViewDataMaxCountInPage As Integer = cGridViewDataMaxCountInPage

	''' <summary>フォームのインスタンスを保持する変数</summary>
	''' <remarks></remarks>
	Private Shared _instance As frmResultGridView

#End Region

#Region "構造体"

	''' <summary>GridViewの状態</summary>
	''' <remarks></remarks>
	Private Structure _GridViewState

#Region "変数"

		''' <summary>GridView表示データ</summary>
		Private _Data As DataTable

		''' <summary>ソートカラム</summary>
		Private _SortColumn As FolderFileListColumn

		''' <summary>ソートカラム</summary>
		''' <remarks>実際にソートするカラム
		'''          ※サイズ付きファイルサイズカラムの場合はファイルサイズカラムで並び替える必要があるため</remarks>
		Private _ActuallySortColumn As FolderFileListColumn

		''' <summary>ソート状態</summary>
		Private _SortOrder As SortOrder

		''' <summary>ソート情報有無</summary>
		Private Shared _HasSortInfo As Boolean = False

#End Region

#Region "プロパティ"

		''' <summary>GridView表示データプロパティ</summary>
		Public Property Data As DataTable

			Set(value As DataTable)

				_Data = value

			End Set
			Get

				Return _Data

			End Get

		End Property

		''' <summary>ソートカラムプロパティ</summary>
		Public Property SortColumn As FolderFileListColumn

			Set(value As FolderFileListColumn)

				_SortColumn = value

			End Set
			Get

				Return _SortColumn

			End Get

		End Property

		''' <summary>ソートカラムプロパティ</summary>
		''' <remarks>実際にソートするカラム
		'''          ※サイズ付きファイルサイズカラムの場合はファイルサイズカラムで並び替える必要があるため</remarks>
		Public Property ActuallySortColumn As FolderFileListColumn

			Set(value As FolderFileListColumn)

				_ActuallySortColumn = value

			End Set
			Get

				Return _ActuallySortColumn

			End Get

		End Property

		''' <summary>ソート状態プロパティ</summary>
		Public Property SortOrder As SortOrder

			Set(value As SortOrder)

				_SortOrder = value

			End Set
			Get

				Return _SortOrder

			End Get

		End Property

		''' <summary>ソート情報有無プロパティ</summary>
		Public Property HasSortInfo As Boolean

			Set(value As Boolean)

				_HasSortInfo = value

			End Set
			Get

				Return HasSortInfo

			End Get

		End Property

#End Region

#Region "メソッド"

		''' <summary>ソート情報を取得</summary>
		''' <param name="pSortColumn">ソートカラム</param>
		''' <param name="pSortOrder">並び替え状態</param>
		''' <remarks></remarks>
		Public Sub SetSortInfo(ByVal pSortColumn As FolderFileListColumn, ByVal pSortOrder As SortOrder)

			'ソート対象カラムを取得する
			Dim mSortColumn As FolderFileListColumn
			Select Case pSortColumn

				Case FolderFileListColumn.SizeAndUnit

					'「単位付きサイズ」項目の場合は「サイズ」項目に変更
					mSortColumn = FolderFileListColumn.Size

				Case Else

					mSortColumn = pSortColumn

			End Select

			'ソート情報をセット
			_SortColumn = pSortColumn
			_ActuallySortColumn = mSortColumn
			_SortOrder = pSortOrder

		End Sub

		''' <summary>並び替えを切り替える</summary>
		''' <remarks></remarks>
		Public Sub ToggleSortOrder()

			'選択されている列のソート状態より処理を分岐
			Select Case _SortOrder

				Case SortOrder.Ascending

					'昇順の場合は降順をセット
					_SortOrder = SortOrder.Descending

				Case SortOrder.Descending, SortOrder.None

					'降順またはなしの場合は昇順をセット
					_SortOrder = SortOrder.Ascending

			End Select

		End Sub

		''' <summary>並び替え</summary>
		''' <remarks></remarks>
		Public Sub Sort()

			'並び替え処理
			_Data = FolderFileList.GetSpecifiedColumnFolderFileList(_Data, _ActuallySortColumn, _SortOrder)

		End Sub

#End Region

	End Structure

#End Region

#Region "プロパティ"

	''' <summary>フォームにアクセスするためのプロパティ</summary>
	''' <remarks>※デザインパターンのSingletonパターンです
	'''            インスタンスがただ１つであることを保証する</remarks>
	Public Shared ReadOnly Property Instance() As frmResultGridView

		Get

			'インスタンスが作成されてなかったらインスタンスを作成
			If _instance Is Nothing Then _instance = New frmResultGridView

			Return _instance

		End Get

	End Property

	''' <summary>リスト表示フォームインスタンス存在プロパティ</summary>
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

	''' <summary>フォルダファイルリスト１ページ内に表示できるファイル数セットプロパティ</summary>
	''' <remarks></remarks>
	Public WriteOnly Property MaxCountInPage As Integer

		Set(value As Integer)

			_GridViewDataMaxCountInPage = value

		End Set

	End Property

#End Region

#Region "コンストラクタ"

	''' <summary>コンストラクタ</summary>
	''' <remarks>引数無しは外部に公開しない</remarks>>
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
	Private Sub frmResultGridView_Load(sender As Object, e As EventArgs) Handles Me.Load

		'初期コントロール設定
		_SetInitialControlInScreen()

	End Sub

	''' <summary>フォームShownイベント</summary>
	''' <param name="sender">Formオブジェクト</param>
	''' <param name="e">Shownイベント</param>
	''' <remarks></remarks>
	Private Sub frmResultGridView_Shown(sender As Object, e As EventArgs) Handles Me.Shown

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
	Private Sub frmResultGridView_Resize(sender As Object, e As EventArgs) Handles Me.Resize

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
	Private Sub frmResultGridView_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

		Dim mForcusedControl As Control = Me.ActiveControl

		'名前（LIKE検索）コントロールにフォーカスがある時は処理を終了
		'※Shiftキーを押して文字入力するとショートカットキーとして認識されてしまうための対応
		If mForcusedControl.Name = txtName.Name Then Exit Sub

		If e.Control Then

			'-------------------------------
			' Ctrlキーのショートカットキー
			'-------------------------------
			Select Case e.KeyCode.ToString

				Case "F" 'Ctrl+F：「名前（LIKE検索）」へフォーカスをセット

					txtName.Focus()

			End Select

		ElseIf e.Shift Then

			'-------------------------------
			' Shiftキーのショートカットキー
			'-------------------------------
			Select Case e.KeyCode.ToString

				Case "A" 'Shift+A：「全て表示」を選択

					rbtnAll.Checked = True

				Case "K" 'Shift+K：「フォルダのみ表示」を選択

					rbtnFolder.Checked = True

				Case "F" 'Shift+F：「ファイルのみ表示」を選択

					rbtnFile.Checked = True

				Case "E" 'Shift+E：「拡張子指定」へフォーカスをセット

					cmbExtension.Focus()

					'「拡張子指定」を開いた状態で表示
					cmbExtension.DroppedDown = True

				Case "D" 'Shift+D：「ファイルサイズ指定」へフォーカスをセット

					cmbFileSizeLevel.Focus()

					'「ファイルサイズ指定」を開いた状態で表示
					cmbFileSizeLevel.DroppedDown = True

				Case "L" 'Shift+L：「レベルごと色を付ける」のチェックを反転

					chkFileSizeLevel.Checked = Not (chkFileSizeLevel.Checked)

				Case "O" 'Shift+O：「出力文字列フォーム」ボタンをクリック

					btnResultTextForm.PerformClick()

				Case "H" 'Shift+H：「Html出力」ボタンをクリック

					btnHtmlOutput.PerformClick()

				Case "C" 'Shift+C：「CSV出力」ボタンをクリック

					btnCsvOutput.PerformClick()

				Case "T" 'Shift+T：「TSV出力」ボタンをクリック

					btnTsvOutput.PerformClick()

				Case "P" 'Shift+P：「前へ」ボタンをクリック

					btnPrevPage.PerformClick()

				Case "N" 'Shift+N：「次へ」ボタンをクリック

					btnNextPage.PerformClick()

			End Select

		End If

	End Sub

	''' <summary>フォームのClosedイベント</summary>
	''' <param name="sender">Formオブジェクト</param>
	''' <param name="e">FormClosedイベント</param>
	''' <remarks>フォームが閉じられようとした際に閉じる処理をキャンセルさせる時に使用するのがFormClosing
	'''          フォームが閉じる前提で後処理するのがFormClosed                                          </remarks>
	Private Sub frmResultGridView_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

		'インスタンスへの参照を破棄
		'※もう一度、表示する時はインスタンスを再作成させるため
		_instance = Nothing

		'出力文字列フォームのインスタンスが存在したら
		'※インスタンスがあるってことはメインフォームの閉じる処理でないってこと
		If frmResultText.HasInstance Then

			'設定ファイルへ書き込み処理
			Settings.Instance.TargetForm = CommandLine.FormType.List
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

		'出力文字列フォームのインスタンス変数を破棄
		'※リスト表示フォームが閉じるタイミングで出力文字列フォームも閉じる仕様のため
		frmResultText.DisposeInstance()

	End Sub

#End Region

#Region "絞り込み関係"

	''' <summary>絞り込みイベント</summary>
	''' <param name="sender">イベント対象コントロール</param>
	''' <param name="e">イベント対象コントロールのイベント</param>
	''' <remarks>ラジオボタンのCheckedDhanged、コンボボックスのTextChanged、ボタンのClick</remarks>
	Private Sub RefineByFolderFileList_Event(sender As Object, e As EventArgs) Handles rbtnAll.CheckedChanged _
																					 , rbtnFolder.CheckedChanged _
																					 , rbtnFile.CheckedChanged _
																					 , cmbExtension.TextChanged _
																					 , cmbFileSizeLevel.TextChanged _
																					 , btnNameSearch.Click

		'フォルダファイルリストを再表示
		Call _ReDisplayRefineByFolderFileList()

	End Sub

	''' <summary>コントロールのKeyDownイベント</summary>
	''' <param name="sender">イベント対象コントロール</param>
	''' <param name="e">イベント対象コントロールのKeyDownイベント</param>
	''' <remarks></remarks>
	Private Sub Controls_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbExtension.KeyDown _
																			, cmbFileSizeLevel.KeyDown _
																			, txtName.KeyDown

		'Enterキーが押された時
		If e.KeyCode = Keys.Enter Then

			'フォルダファイルリストを再表示
			Call _ReDisplayRefineByFolderFileList()

		End If

	End Sub

#End Region

#Region "レベルごと色を付けるチェックボックス"

	''' <summary>レベルごと色を付けるチェックボックスのチェックイベント</summary>
	''' <param name="sender">レベルごと色を付けるチェックボックス</param>
	''' <param name="e">チェックイベント</param>
	''' <remarks></remarks>
	Private Sub chkFileSizeLevel_CheckedChanged(sender As Object, e As EventArgs) Handles chkFileSizeLevel.CheckedChanged

		'ファイルサイズレベルによって色を付ける
		Call _ToColorRowByFileSizeLevel()

	End Sub

#End Region

#Region "ボタン"

	''' <summary>次へボタンクリックイベント</summary>
	''' <param name="sender">次へボタン</param>
	''' <param name="e">Clickイベント</param>
	''' <remarks></remarks>
	Private Sub btnNextPage_Click(sender As Object, e As EventArgs) Handles btnNextPage.Click

		'現在ページをカウントアップしてフォルダファイルリストページング処理
		_SetPaging(_CurrentPage + 1)

		'ヘッダーセルにソート状態をセット
		dgvFolderFileList.Columns(_CurrentGridView.SortColumn).HeaderCell.SortGlyphDirection = _CurrentGridView.SortOrder

		'画面コントロール設定
		_SetControlInScreen()

	End Sub

	''' <summary>前へボタンクリックイベント</summary>
	''' <param name="sender">前へボタン</param>
	''' <param name="e">Clickイベント</param>
	''' <remarks></remarks>
	Private Sub btnPrevPage_Click(sender As Object, e As EventArgs) Handles btnPrevPage.Click

		'現在ページをカウントダウンしてフォルダファイルリストページング処理
		_SetPaging(_CurrentPage - 1)

		'ヘッダーセルにソート状態をセット
		dgvFolderFileList.Columns(_CurrentGridView.SortColumn).HeaderCell.SortGlyphDirection = _CurrentGridView.SortOrder

		'画面コントロール設定
		_SetControlInScreen()

	End Sub

	''' <summary>CSV出力ボタンクリックイベント</summary>
	''' <param name="sender">CSV出力ボタン</param>
	''' <param name="e">Clickイベント</param>
	''' <remarks></remarks>
	Private Async Sub btnCsvOutput_Click(sender As Object, e As EventArgs) Handles btnCsvOutput.Click

		'メッセージボックスから押されたボタンにより処理を分岐
		Select Case MyBase.ShowDialogueMessage(OutputDelimiterText.Delimiter.CSV.ToString)

			Case Windows.Forms.DialogResult.Yes

                'CSVファイル保存・実行処理が行われた時
                If _SaveRunOutputFile(OutputFileFormat.CSV, _FolderFileList.TargetPathFolderName, cEncording.ShiftJis) Then

                    '保存したファイルを実行しました通知を表示
                    MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageRunSaveFile)

                    '非同期でMessenger風通知メッセージが非表示になるまで待機
                    Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

                End If

            Case Windows.Forms.DialogResult.No

				'出力用テキストをクリップボードコピー
				Clipboard.SetText(_GetOutputFileText(OutputFileFormat.CSV))

                'クリップボードにコピーしました通知を表示
				MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageCopyToClipBoard(OutputFileFormat.CSV))

                '非同期でMessenger風通知メッセージが非表示になるまで待機
                Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

		End Select

	End Sub

	''' <summary>TSV出力ボタンクリックイベント</summary>
	''' <param name="sender">TSV出力ボタン</param>
	''' <param name="e">Clickイベント</param>
	''' <remarks></remarks>
	Private Async Sub btnTsvOutput_Click(sender As Object, e As EventArgs) Handles btnTsvOutput.Click

		'メッセージボックスから押されたボタンにより処理を分岐
		Select Case MyBase.ShowDialogueMessage(OutputDelimiterText.Delimiter.TSV.ToString)

			Case Windows.Forms.DialogResult.Yes

                'TSVファイル保存・実行処理が行われた時
                If _SaveRunOutputFile(OutputFileFormat.TSV, _FolderFileList.TargetPathFolderName, cEncording.ShiftJis) Then

                    '保存したファイルを実行しました通知を表示
                    MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageRunSaveFile)

                    '非同期でMessenger風通知メッセージが非表示になるまで待機
                    Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

                End If

            Case Windows.Forms.DialogResult.No

				'出力用テキストをクリップボードコピー
				Clipboard.SetText(_GetOutputFileText(OutputFileFormat.TSV))

                'クリップボードにコピーしました通知を表示
				Dim mFrmPopupMessage As New frmPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageCopyToClipBoard(OutputFileFormat.TSV))
                mFrmPopupMessage.Show()

				'非同期でMessenger風通知メッセージが非表示になるまで待機
				Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

		End Select

	End Sub

    ''' <summary>html出力ボタンクリックイベント</summary>
    ''' <param name="sender">html出力ボタン</param>
    ''' <param name="e">Clickイベント</param>
    ''' <remarks></remarks>
    Private Async Sub btnHtmlOutput_Click(sender As Object, e As EventArgs) Handles btnHtmlOutput.Click

        'HTMLファイル保存・実行処理が行われた時
        If _SaveRunOutputFile(OutputFileFormat.HTML, _FolderFileList.TargetPathFolderName, cEncording.UTF8) Then

            '保存したファイルを実行しました通知を表示
            MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageRunSaveFile)

            '非同期でMessenger風通知メッセージが非表示になるまで待機
            Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

        End If

    End Sub

    ''' <summary>出力文字列フォームボタンクリックイベント</summary>
    ''' <param name="sender">出力文字列フォームボタン</param>
    ''' <param name="e">Clickイベント</param>
    ''' <remarks></remarks>
    Private Sub btnResultTextForm_Click(sender As Object, e As EventArgs) Handles btnResultTextForm.Click

		'リスト表示フォームを非表示
		Me.Hide()

		'出力文字列フォームをモードレスで表示
		frmResultText.Instance.Show()

	End Sub

#End Region

#Region "テキストボックス"

	''' <summary>名前テキストボックスGotFocusイベント</summary>
	''' <param name="sender">名前テキストボックス</param>
	''' <param name="e">GotFocusイベント</param>
	''' <remarks></remarks>
	Private Sub txtName_GotFocus(sender As Object, e As EventArgs) Handles txtName.GotFocus

		'テキストボックスの内容を全選択状態に
		txtName.SelectAll()

	End Sub

	''' <summary>名前テキストボックスMouseDownイベント</summary>
	''' <param name="sender">名前テキストボックス</param>
	''' <param name="e">MouseDownイベント</param>
	''' <remarks></remarks>
	Private Sub txtName_MouseDown(sender As Object, e As MouseEventArgs) Handles txtName.MouseDown

		'テキストボックスの内容を全選択状態に
		txtName.SelectAll()

	End Sub

#End Region

#Region "データグリッドビュー"

	''' <summary>ヘッダーカラムのクリックイベント</summary>
	''' <param name="sender">フォルダファイルリストGridView</param>
	''' <param name="e">ColumnHeaderMouseClickイベント</param>
	''' <remarks></remarks>
	Private Sub dgvFolderFileList_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvFolderFileList.ColumnHeaderMouseClick

		Dim mDgv As DataGridView = DirectCast(sender, DataGridView)

		'クリックされたヘッダーセルのインデックスを取得
		Dim mClickHeaderCell As FolderFileListColumn = DirectCast(e.ColumnIndex, FolderFileListColumn)

		'フォルダファイルリストの並び替え
		Call _SortFolderFileList(mDgv, mClickHeaderCell)

	End Sub

	''' <summary>キーダウンイベント</summary>
	''' <param name="sender">フォルダファイルリストGridView</param>
	''' <param name="e">KeyDownイベント</param>
	Private Sub dgvFolderFileList_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvFolderFileList.KeyDown

		Dim mDgv As DataGridView = DirectCast(sender, DataGridView)

		If e.Shift Then

			'-------------------------------
			' Shiftキーのショートカットキー
			'-------------------------------
			Select Case e.KeyCode.ToString

				Case "S" 'Shift+S：「フォルダファイルリストの並び替えを切り替える」

					'ソートカラムを取得
					Dim mSortColumn As FolderFileListColumn = DirectCast(mDgv.SelectedCells(0).ColumnIndex, FolderFileListColumn)

					'フォルダファイルリストの並び替え
					Call _SortFolderFileList(mDgv, mSortColumn)

				Case "M" 'Shift+M：「セルの幅を自動調整（ヘッダー、表示内容で）の切り替え」

					Select Case mDgv.AutoSizeColumnsMode

						Case DataGridViewAutoSizeColumnsMode.AllCells

							'セルの幅を自動調整（ヘッダー、表示内容で）をリセット
							mDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

						Case DataGridViewAutoSizeColumnsMode.None

							'セルの幅を自動調整（ヘッダー、表示内容で）をセット
							mDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

					End Select

			End Select

		Else

			If e.KeyCode = Keys.Enter Then

				'-----------------------
				' Enterキーが押された時
				'-----------------------

				'選択されているセル分繰り返す
				For Each mSelectedCell As DataGridViewCell In mDgv.SelectedCells

					'選択セルからファイルのフルパスを取得
					Dim mFileFullPath As String = _GetFullPathFromGridViewCell(mDgv, mSelectedCell)

					'フルパスを取得出来た時、ファイルを実行
					If Not String.IsNullOrEmpty(mFileFullPath) Then MyBase.RunFile(mFileFullPath)

				Next

				'下のフォーカスへ移動する処理を無効化
				e.Handled = True

			End If

		End If

	End Sub

	''' <summary>セルのダブルクリックイベント</summary>
	''' <param name="sender">フォルダファイルリストGridView</param>
	''' <param name="e">CellDoubleClick</param>
	Private Sub dgvFolderFileList_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvFolderFileList.CellDoubleClick

		Dim mDgv As DataGridView = DirectCast(sender, DataGridView)

		'ヘッダーセル以外の時 ※ヘッダーセルをダブルクリックするとe.RowIndexが-1になる
		If e.RowIndex > -1 Then

			'選択セルからファイルのフルパスを取得
			Dim mFileFullPath As String = _GetFullPathFromGridViewCell(mDgv, mDgv.CurrentCell)

			'フルパスを取得出来た時、ファイルを実行
			If Not String.IsNullOrEmpty(mFileFullPath) Then MyBase.RunFile(mFileFullPath)

		End If

	End Sub

#End Region

#End Region

#Region "メソッド"

#Region "コントロール設定メソッド"

	''' <summary>初期画面コントロール設定</summary>
	''' <remarks>※この関数でコンボボックスのDroppedDownをTrueにセットすると
	'''            フォームがまだ表示されていないのにも関わらずリストが勝手
	'''            に表示されてしまう                                       </remarks>
	Private Sub _SetInitialControlInScreen() Implements IFormCommonProcess._SetInitialControlInScreen

		'フォームを透明状態にする
		Me.Opacity = 0

		'フォームでもキーイベントを取得可に設定
		'※ショートカットキーを設定出来るようにするため
		Me.KeyPreview = True

		'全て表示ラジオボタンをチェック状態に
		rbtnAll.Checked = True

		'拡張子コンボボックス設定（フォルダファイルリスト内の拡張子）
		cmbExtension.Items.AddRange(_FolderFileList.ExtensionList().ToArray)

		'ファイルサイズレベルコンボボックス設定
		cmbFileSizeLevel.DataSource = FolderFileList.cFileSizeLevel.LevelList
		cmbFileSizeLevel.ValueMember = FolderFileList.cFileSizeLevel.LevelListColum.ID.ToString
		cmbFileSizeLevel.DisplayMember = FolderFileList.cFileSizeLevel.LevelListColum.NAME.ToString

		'GridViewに表示させるデータを変数にセット
		_CurrentGridView.Data = _FolderFileList.FolderFileList

		'現在ページを初期化してフォルダファイルリストページング処理
		Call _SetPaging(1)

		'画面コントロール設定
		Call _SetControlInScreen()

		'フォームロード時にフォルダファイルリストへフォーカスをセットする
		'※フォームロード時のフォーカスはタブ・オーダー順で１番のものにセットされる
		Me.ActiveControl = dgvFolderFileList

		'フォームのタイトルにコマンドモード文言を追加
		CommandLine.SetCommandModeToTitle(Me)

		'フォームのタイトルにデバッグモード文言を追加（デバッグモード時のみ実行される）
		DebugMode.SetDebugModeTitle(Me)

	End Sub

	''' <summary>画面コントロール設定</summary>
	''' <remarks>再描画時にコントロールの再設定を行う</remarks>
	Private Sub _SetControlInScreen()

		'----------------------------------------
		' 拡張子コンボボックス設定
		'----------------------------------------
		'チェック済みのラジオボタンを取得
		Dim mRbtnButton As RadioButton = _GetCheckedRadioButton(grpDisplayTarget)

		'ラジオボタンの状態により拡張子コンボボックスの設定を行う
		Select Case mRbtnButton.Name

			Case rbtnAll.Name

				'コンボボックスの使用を可に
				cmbExtension.Enabled = True

			Case rbtnFolder.Name

				'コンボボックスの使用を不可に、値を初期値（空文字）に設定
				cmbExtension.Enabled = False
				cmbExtension.SelectedIndex = 0

			Case rbtnFile.Name

				'コンボボックスの使用を可に
				cmbExtension.Enabled = True

		End Select

		'----------------------------------------
		' ファイルサイズレベルで行に色を付ける
		' ※レベルごと色を付けるにチェックがつい
		'   ていた場合のみ
		'----------------------------------------
		Call _ToColorRowByFileSizeLevel()

		'----------------------------------------
		' フォルダファイルリスト表示GridView設定
		'----------------------------------------
		Call _SetLayoutForDgvFolderFileList()

		'フォーカスをフォルダファイルリストへセット
		dgvFolderFileList.Focus()

	End Sub

	''' <summary>フォルダファイルリストGridViewのレイアウト設定</summary>
	''' <remarks>セルスタイルの優先順位
	'''            ①DataGridViewCell.Style
	'''            ②DataGridViewRow.DefaultCellStyle
	'''            ③DataGridView.AlternatingRowsDefaultCellStyle
	'''            ④DataGridView.RowsDefaultCellStyle
	'''            ⑤DataGridViewColumn.DefaultCellStyle
	'''            ⑥DataGridView.DefaultCellStyle
	'''          ※継承関係で上記の優先順位となる                 </remarks>
	Private Sub _SetLayoutForDgvFolderFileList()

		'----------------------------------------
		' カラム設定の表示設定
		'----------------------------------------
		dgvFolderFileList.Columns(FolderFileListColumn.No).Visible = False
		dgvFolderFileList.Columns(FolderFileListColumn.Name).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.UpdateDate).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.FileSystemType).Visible = False
		dgvFolderFileList.Columns(FolderFileListColumn.FileSystemTypeName).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.Extension).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.Size).Visible = False
		dgvFolderFileList.Columns(FolderFileListColumn.SizeAndUnit).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.SizeLevel).Visible = False
		dgvFolderFileList.Columns(FolderFileListColumn.SizeLevelName).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.SizeAndUnit).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.DirectoryLevel).Visible = False
		dgvFolderFileList.Columns(FolderFileListColumn.ParentFolder).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.ParentFolderFullPath).Visible = False
		dgvFolderFileList.Columns(FolderFileListColumn.UnderTargetFolder).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.FullPath).Visible = True
		dgvFolderFileList.Columns(FolderFileListColumn.IsLastFileInFolder).Visible = False
		dgvFolderFileList.Columns(FolderFileListColumn.DispString).Visible = False

		'全てのカラムを表示する（デバッグモード時のみ実行される）
		DebugMode.DisplayAllColumnForFrmResultGridView(CType(Me, frmResultGridView))

		'----------------------------------------
		' カラムのフォントカラー設定
		'----------------------------------------
		dgvFolderFileList.Columns(FolderFileListColumn.Name).DefaultCellStyle.ForeColor = _cGridViewColor.RunColumnFont
		dgvFolderFileList.Columns(FolderFileListColumn.ParentFolder).DefaultCellStyle.ForeColor = _cGridViewColor.RunColumnFont
		dgvFolderFileList.Columns(FolderFileListColumn.UnderTargetFolder).DefaultCellStyle.ForeColor = _cGridViewColor.RunColumnFont
		dgvFolderFileList.Columns(FolderFileListColumn.FullPath).DefaultCellStyle.ForeColor = _cGridViewColor.RunColumnFont

		'----------------------------------------
		' GridViewプロパティ設定
		'----------------------------------------
		'読み込み専用（編集不可）に設定
		dgvFolderFileList.ReadOnly = True

		'行ヘッダーを非表示
		dgvFolderFileList.RowHeadersVisible = False

		'追加行を非表示
		dgvFolderFileList.AllowUserToAddRows = False

		'列の幅の変更を可に設定
		dgvFolderFileList.AllowUserToResizeColumns = True

		'行の高さの変更を不可に設定
		dgvFolderFileList.AllowUserToResizeRows = False

		'ヘッダーのデザインを変更できるようにする
		dgvFolderFileList.EnableHeadersVisualStyles = False

		'ヘッダー名を取得
		Dim mHeaderNames As ArrayList = cFolderFileListHeaders.HeaderNames

		For Each mColumn As DataGridViewColumn In dgvFolderFileList.Columns

			'ヘッダー背景色を灰色に
			mColumn.HeaderCell.Style.BackColor = Color.LightGray

			'ヘッダー名に日本語名を設定
			mColumn.HeaderText = mHeaderNames(mColumn.Index)

			'ヘッダーセルの最小幅を設定
			'※セルの幅を自動調整を実行した時、ヘッダーセルが２行になってしまうための対応
			mColumn.MinimumWidth = GetHeaderCellWidth(mColumn.Index)

			'ヘッダーセルクリックでの自動並び替えをプログラムで行う
			mColumn.SortMode = DataGridViewColumnSortMode.Programmatic

			'ヘッダーセルのソート状態により色を付ける
			Call _ToColorSortHeaderCell(mColumn.HeaderCell)

		Next

		'ヘッダーのフォントサイズを変更、文字を太字に
		dgvFolderFileList.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 11, FontStyle.Bold)

	End Sub

	''' <summary>コントロールのサイズ変更・再配置</summary>
	''' <param name="pChangedSize">フォームの変更後サイズ</param>
	''' <remarks>全てのコントロールを取得しコントロールごと処理を行う</remarks>
	Private Sub _ResizeAndRealignmentControls(ByVal pChangedSize As System.Drawing.Size) Implements IFormCommonProcess._ResizeAndRealignmentControls

		'Form以外の全てのコントロールを取得
		Dim mAllControls As Control() = MyBase.GetControlsInTarget(Me)

		For Each mTargetControl As Control In mAllControls

			Select Case mTargetControl.Name

				Case dgvFolderFileList.Name

					'フォルダファイルリストGridViewの場合はフォームのサイズの変更分、サイズを変更する
					mTargetControl.Size = New System.Drawing.Size(mTargetControl.Size + pChangedSize)

				Case btnPrevPage.Name, btnNextPage.Name, txtMaxSearchCount.Name, txtRangeStart.Name, txtRangeEnd.Name _
				   , lblMaxSearchCount.Name, lblRangeStart.Name, lblRange.Name, lblRangeEnd.Name

					'ページング関係のコントロールの場合はフォームの幅の変更分、左の位置を変更する
					mTargetControl.Left = mTargetControl.Left + pChangedSize.Width

				Case btnHtmlOutput.Name, btnCsvOutput.Name, btnTsvOutput.Name

					'Html出力ボタン、CSV出力ボタン、TSV出力ボタンの場合はフォームのサイズの変更分、表示位置を変更する
					mTargetControl.Location = New System.Drawing.Point(mTargetControl.Location + pChangedSize)

				Case btnResultTextForm.Name

					'文字列出力フォームへボタンの場合はフォームの高さの変更分、高さの位置を変更する
					mTargetControl.Top = mTargetControl.Top + pChangedSize.Height

			End Select

		Next

	End Sub

	''' <summary>チェックされているラジオボタンを取得</summary>
	''' <param name="pGrpBox">チェックされているラジオボタンを探すグループボックス</param>
	''' <returns>チェックされているラジオボタン</returns>
	''' <remarks></remarks>
	Private Function _GetCheckedRadioButton(ByVal pGrpBox As GroupBox) As RadioButton

		'対象のグループボックス内のラジオボタン分繰り返す
		For Each rbtnButton As Windows.Forms.RadioButton In pGrpBox.Controls

			'チェックされているラジオボタンが見つかったらそれを返す
			'※グループボックス内のラジオボタンの内どれか1つが必ずチェックされている
			If rbtnButton.Checked Then Return rbtnButton

		Next

		Return Nothing

	End Function

	''' <summary>絞り込みデータを取得</summary>
	''' <returns>絞り込みデータ</returns>
	''' <remarks>画面内の条件に一致する絞り込みデータを返す</remarks>
	Private Function _GetRefineByFolderFileListForDisplay() As DataTable

		Dim mDisplayData As New DataTable

		'----------------------------------
		' 表示対象データを取得
		'----------------------------------
		'チェック済みのラジオボタンを取得
		Dim mRbtnButton As RadioButton = _GetCheckedRadioButton(grpDisplayTarget)

		Select Case mRbtnButton.Name

			Case rbtnAll.Name

				mDisplayData = _FolderFileList.FolderFileList

			Case rbtnFolder.Name

				mDisplayData = _FolderFileList.FolderFileListOnlyFolder

			Case rbtnFile.Name

				mDisplayData = _FolderFileList.FolderFileListOnlyFile

		End Select

		'----------------------------------
		' 表示対象データを拡張子で絞り込んだ
		' データを取得
		'----------------------------------
		Dim mExtensionWhere As String = cmbExtension.SelectedItem
		mDisplayData = _FolderFileList.GetRefineByFolderFileList(mDisplayData, FolderFileListColumn.Extension, mExtensionWhere)

		'----------------------------------
		' 表示対象データをファイルサイズレ
		' ベルで絞り込んだデータを取得
		'----------------------------------
		Dim mFileSizeLevelWhere As String = cmbFileSizeLevel.SelectedItem(FolderFileList.cFileSizeLevel.LevelListColum.ID)
		mDisplayData = _FolderFileList.GetRefineByFolderFileList(mDisplayData, FolderFileListColumn.SizeLevel, mFileSizeLevelWhere)

		'----------------------------------
		' 表示対象データを名前(LIKE検索)で
		' 絞り込んだデータを取得
		'----------------------------------
		Dim mNameWhere As String = txtName.Text
		mDisplayData = _FolderFileList.GetRefineByFolderFileList(mDisplayData, FolderFileListColumn.Name, mNameWhere)

		Return mDisplayData

	End Function

	''' <summary>ヘッダーセルの最小幅を取得する</summary>
	''' <param name="pHeaderColumnIndex">カラムインデックス</param>
	''' <returns>対象ヘッダーセルの最小幅</returns>
	''' <remarks>DataGridViewAutoSizeColumnsMode.AllCellsプロパティ（表示内容と
	'''          ヘッダーの内容で幅を自動調整）だけだとヘッダー名が折り返してし
	'''          まったのでヘッダーが折り返さない最小幅を設定することで折り返し
	'''          を回避することが出来た                                        </remarks>
	Private Function GetHeaderCellWidth(ByVal pHeaderColumnIndex As Integer) As Integer

		Select Case pHeaderColumnIndex

			Case FolderFileListColumn.No

				'「ＮＯ」は６０を返す
				Return 60

			Case FolderFileListColumn.Extension

				'「拡張子」は９０を返す
				Return 90

			Case FolderFileListColumn.UpdateDate, FolderFileListColumn.SizeAndUnit

				'「更新日時」、「単位付きサイズ」は１１０を返す
				Return 110

			Case FolderFileListColumn.DirectoryLevel, FolderFileListColumn.IsLastFileInFolder

				'「ディレクトリレベル」、「最後のファイルか」は１３０を返す
				Return 130

			Case FolderFileListColumn.FileSystemTypeName, FolderFileListColumn.SizeLevelName, FolderFileListColumn.DispString

				'「ファイルシステムタイプ名」、「ファイルサイズレベル名」、「表示文字列」は１５０を返す
				Return 150

			Case FolderFileListColumn.FileSystemType, FolderFileListColumn.SizeLevel

				'「ファイルシステムタイプ」、「ファイルサイズレベル」は１６０を返す
				Return 160

			Case FolderFileListColumn.Size

				'「ファイルサイズ」は１７０を返す
				Return 170

			Case FolderFileListColumn.Name, FolderFileListColumn.ParentFolder, FolderFileListColumn.ParentFolderFullPath, FolderFileListColumn.UnderTargetFolder, FolderFileListColumn.FullPath

				'「ファイル名」、「親フォルダ」、「親フォルダフルパス」、「対象フォルダ以下」、「フルパス」は３００を返す
				Return 300

			Case Else

				'デフォルトは１００を返す
				Return 100

		End Select

	End Function

	''' <summary>ヘッダーセルの並び状態により色を付ける</summary>
	''' <param name="pHeaderCell">グリッドビューのヘッダーセル</param>
	''' <remarks></remarks>
	Private Sub _ToColorSortHeaderCell(ByVal pHeaderCell As DataGridViewColumnHeaderCell)

		Select Case pHeaderCell.SortGlyphDirection

			Case SortOrder.Ascending

				pHeaderCell.Style.BackColor = _cGridViewColor.HeaderAsceding

			Case SortOrder.Descending

				pHeaderCell.Style.BackColor = _cGridViewColor.HeaderDesceding

		End Select

	End Sub

    ''' <summary>ファイルサイズレベルで行に色を付ける</summary>
    ''' <remarks>レベルごと色を付けるチェックボックスが
    '''            チェックされている時は色をつける
    '''            チェックされていない時は色をなくす  </remarks>
    Private Sub _ToColorRowByFileSizeLevel()

        For Each mRow As DataGridViewRow In dgvFolderFileList.Rows

            Dim mFileSizeLevel As FolderFileList.cFileSizeLevel.SizeLevel

            'ファイルサイズレベルを取得
            If chkFileSizeLevel.Checked = True Then

                mFileSizeLevel = Integer.Parse(mRow.Cells(FolderFileListColumn.SizeLevel).Value)

            Else

                mFileSizeLevel = Nothing

            End If

            '背景色をファイルサイズレベルに応じた色にする
            mRow.DefaultCellStyle.BackColor = FolderFileList.cFileSizeLevel.LevelColor(mFileSizeLevel)

        Next

    End Sub

#End Region

#Region "ページングメソッド"

    ''' <summary>ページング処理</summary>
    ''' <param name="pPageNo">ページ番号</param>
    ''' <remarks>対象ページ番号のページング設定を行う</remarks>
    Private Sub _SetPaging(ByVal pPageNo As Integer)

		'現在ページにセット
		_CurrentPage = pPageNo

		'フォルダファイルリストページング処理
		Call _SetPagingControl(_CurrentPage)

		'フォルダファイルリストに対象の範囲データをセット
		dgvFolderFileList.DataSource = _GetRefineByFolderFileListForPaging(_CurrentGridView.Data)

	End Sub

	''' <summary>ページング関係コントロール設定</summary>
	''' <param name="pPageNo">ページ番号</param>
	''' <remarks></remarks>
	Private Sub _SetPagingControl(ByVal pPageNo As Integer)

		'------------------------------
		' 最大件数
		'------------------------------
		'最大件数を取得
		Dim mMaxSearchCount As Integer = _CurrentGridView.Data.Rows.Count

		'------------------------------
		' ページ数範囲（最初）件数
		'------------------------------
		Dim mRangeStartPage As Integer

		If mMaxSearchCount = 0 Then

			mRangeStartPage = 0

		ElseIf pPageNo = 1 Then

			mRangeStartPage = 1

		Else

			'「ページ番号 - 1」 * 「１ページ内に表示できるファイル数」をセット
			mRangeStartPage = (pPageNo - 1) * _GridViewDataMaxCountInPage

		End If

		'------------------------------
		' ページ数範囲（最後）件数計算
		'------------------------------
		Dim mRangeEndPage As Integer

		'「現在表示最大ページ数」 = 「ページ番号 - 1」 * 「１ページ内に表示できるファイル数」 + 「１ページ内に表示できるファイル数」
		Dim mCurrentMaxPageCount As Integer = ((pPageNo - 1) * _GridViewDataMaxCountInPage) + _GridViewDataMaxCountInPage

		'「フォルダファイルリストの件数」が「１ページ内に表示できるファイル数」以下の時
		'または
		'「フォルダファイルリストの件数」が「現在表示最大ページ数」より小さい時
		If mMaxSearchCount <= _GridViewDataMaxCountInPage OrElse mMaxSearchCount < mCurrentMaxPageCount Then

			'「最大件数」をセット
			mRangeEndPage = mMaxSearchCount

		Else

			'「現在表示最大ページ数」をセット
			mRangeEndPage = mCurrentMaxPageCount

		End If

		'------------------------------
		' 最大ページ数計算
		'------------------------------
		'「フォルダファイルリスト件数」 / 「１ページ内に表示できるファイル数」
		Dim mMaxPageCount As Integer = Math.Floor(mMaxSearchCount / _GridViewDataMaxCountInPage)

		'「フォルダファイルリスト件数」 / 「１ページ内に表示できるファイル数」の余りが０でない時は＋１する
		'「１ページ内に表示できるファイル数」の倍数でない時
		If mMaxSearchCount Mod _GridViewDataMaxCountInPage <> 0 Then mMaxPageCount = mMaxPageCount + 1

		'------------------------------
		' コントロール設定
		'------------------------------
		'最大件数
		txtMaxSearchCount.Text = mMaxSearchCount

		'ページ数範囲（最初）
		txtRangeStart.Text = mRangeStartPage

		'ページ数範囲（最後）
		txtRangeEnd.Text = mRangeEndPage

		'前へボタンEnable設定
		If pPageNo = 1 Then

			btnPrevPage.Enabled = False

		Else

			btnPrevPage.Enabled = True

		End If

		'次へボタンEnable設定
		If pPageNo = mMaxPageCount Then

			btnNextPage.Enabled = False

		Else

			btnNextPage.Enabled = True

		End If

	End Sub

	''' <summary>ページング後の範囲のデータを取得</summary>
	''' <param name="pFolderFileList">フォルダファイルリスト</param>
	''' <returns>表示範囲のフォルダファイルリスト</returns>
	''' <remarks></remarks>
	Private Function _GetRefineByFolderFileListForPaging(ByVal pFolderFileList As DataTable) As DataTable

		'フォルダファイルリストの件数が０件の時はそのまま返す
		If pFolderFileList.Rows.Count = 0 Then Return pFolderFileList

		'抽出する範囲を取得（始まり位置）
		Dim mStartIndex As Integer = Integer.Parse(txtRangeStart.Text)
		If mStartIndex = 1 Then mStartIndex = 0

		'抽出する範囲を取得（始まり位置から取得する件数）
		'※1ページ内に表示できるファイル数
		Dim mEndIndex As Integer = _GridViewDataMaxCountInPage

		'抽出する範囲のデータを取得
		Dim mRefineByFolderFileListForPaging = pFolderFileList.AsEnumerable().Skip(mStartIndex).Take(mEndIndex)

		'DataRowの配列をDataTableに変換して返す
		Return mRefineByFolderFileListForPaging.CopyToDataTable()

	End Function

#End Region

#Region "グリッドビューメソッド"

	''' <summary>絞り込み条件でフォルダフォルダファイルリストを再表示</summary>
	''' <remarks>絞り込み条件で絞り込んだフォルダファイルリストを再表示させる</remarks>
	Private Sub _ReDisplayRefineByFolderFileList()

		'フォルダファイルリストデータが存在しない場合は処理を終了
		If _FolderFileList Is Nothing Then Exit Sub

		'絞り込んだ状態のフォルダファイルリストを取得
		_CurrentGridView.Data = _GetRefineByFolderFileListForDisplay()

		'現在ページを初期化してフォルダファイルリストページング処理
		Call _SetPaging(1)

		'画面コントロール設定
		Call _SetControlInScreen()

	End Sub

	''' <summary>フォルダファイルリストの並び替え</summary>
	''' <param name="pDgv">フォルダファイルリスト</param>
	''' <param name="pSortColumn">並び替えカラム</param>
	''' <remarks></remarks>
	Private Sub _SortFolderFileList(ByVal pDgv As DataGridView, Optional ByVal pSortColumn As FolderFileListColumn = FolderFileListColumn.Name)

		'選択されている行と列のインデックスを取得する。複数選択されている場合は最初の行と列を取得する
		'※２つ目以降は無視する
		Dim mSelectedColumn As FolderFileListColumn = DirectCast(pDgv.SelectedCells(0).ColumnIndex, FolderFileListColumn)
		Dim mSelectedRow As Integer = pDgv.SelectedCells(0).RowIndex

		'ソート情報をセット
		_CurrentGridView.SetSortInfo(pSortColumn, pDgv.Columns(pSortColumn).HeaderCell.SortGlyphDirection)

		'並び替えを切り替える
		_CurrentGridView.ToggleSortOrder()

		'GridViewに表示されているデータの並び替え
		_CurrentGridView.Sort()

		'現在ページを初期化してフォルダファイルリストページング処理
		_SetPaging(1)

		'ヘッダーセルにソート状態をセット
		pDgv.Columns(pSortColumn).HeaderCell.SortGlyphDirection = _CurrentGridView.SortOrder

		'コントロール設定
		Call _SetControlInScreen()

		'フォーカスを選択されていたセルへ
		pDgv.CurrentCell = pDgv(mSelectedColumn, mSelectedRow)

	End Sub

	''' <summary>グリッドビューセルからフルパスを取得</summary>
	''' <param name="pDgv">フォルダファイルリストGridView</param>
	''' <param name="pCell">対象セル</param>
	''' <returns></returns>
	Private Function _GetFullPathFromGridViewCell(ByVal pDgv As DataGridView, ByVal pCell As DataGridViewCell) As String

		Select Case pCell.ColumnIndex

            Case FolderFileListColumn.Name, FolderFileListColumn.UnderTargetFolder, FolderFileListColumn.FullPath

                'ファイル名、対象フォルダ以下、ファイルのフルパスの時はフルパスを返す
                Return pDgv.Rows(pCell.RowIndex).Cells(FolderFileListColumn.FullPath).Value.ToString

			Case FolderFileListColumn.ParentFolder

                '親フォルダの時は親フォルダのフルパスを返す
                Return pDgv.Rows(pCell.RowIndex).Cells(FolderFileListColumn.ParentFolderFullPath).Value.ToString

			Case Else

				Return String.Empty

		End Select

	End Function

#End Region

#Region "「Html・CSV・TSV」出力用メソッド"

    ''' <summary>出力ファイルの保存・実行処理</summary>
    ''' <param name="pFileFormat">出力形式</param>
    ''' <param name="pDefalutFileName">保存ファイルのデフォルトのファイル名</param>
    ''' <param name="pEncording">エンコード</param>
    ''' <returns>True ：保存ファイルの実行
    '''          False：保存ファイルを実行しない</returns>
    ''' <remarks></remarks>
    Private Function _SaveRunOutputFile(ByVal pFileFormat As OutputFileFormat, ByVal pDefalutFileName As String, ByVal pEncording As System.Text.Encoding) As Boolean

        '名前を付けて保存ダイアログを表示
        Dim mDialog As SaveFileDialog = MyBase.GetSaveAsDialog(pDefalutFileName, pFileFormat.ToString.ToLower)

        '名前をつけて保存ダイアログでOKが押されたら
        If mDialog.ShowDialog = Windows.Forms.DialogResult.OK Then

            'ファイルの保存処理
            MyBase.WriteTextToOutputFile(_GetOutputFileText(pFileFormat), mDialog.FileName, pEncording)

            '実行タイプが実行の時、保存したファイルを関連付けで実行する
			If Settings.Instance.Execute = ExecuteType.Run Then

				MyBase.RunFile(mDialog.FileName)
				Return True

			End If

        End If

        Return False

    End Function

    ''' <summary>出力用ファイルのテキストデータを取得</summary>
    ''' <param name="pFileFormat">出力形式</param>
    ''' <returns>対象出力形式のテキストデータ</returns>
    ''' <remarks></remarks>
    Private Function _GetOutputFileText(ByVal pFileFormat As OutputFileFormat) As String

		Dim mOutputText As String = String.Empty

		'出力形式ごと処理を分岐
		Select Case pFileFormat

			Case OutputFileFormat.CSV

				'画面に表示しているGridViewデータからCSV用テキストのデータを作成
				Dim mOutputCSV As OutputDelimiterText = _CreateOutputTextData(_CurrentGridView.Data.Copy(), OutputDelimiterText.Delimiter.CSV)
				mOutputText = mOutputCSV.AllOutputText

			Case OutputFileFormat.TSV

				'画面に表示しているGridViewデータからTSV用テキストのデータを作成
				Dim mOutputTSV As OutputDelimiterText = _CreateOutputTextData(_CurrentGridView.Data.Copy(), OutputDelimiterText.Delimiter.TSV)
				mOutputText = mOutputTSV.AllOutputText

			Case OutputFileFormat.HTML

				'フォルダファイルリストから出力用のHtml文字列を取得
				Dim mOutputHtml As New OutputHtml(_FolderFileList, CommandLine.FormType.List)
				mOutputText = mOutputHtml.HtmlSentence

		End Select

		Return mOutputText

	End Function

	''' <summary>出力用のテキストデータを取得</summary>
	''' <param name="pGridViewData">画面に表示されているGridViewData</param>
	''' <param name="pDelimiter">区切り文字</param>
	''' <returns>出力用のテキストデータ</returns>
	''' <remarks></remarks>
	Private Function _CreateOutputTextData(ByVal pGridViewData As DataTable, ByVal pDelimiter As OutputDelimiterText.Delimiter) As OutputDelimiterText

		'GridViewに表示しているデータを出力用テキストのデータに変換
		Dim mOutputTextData As DataTable = pGridViewData.Copy
		mOutputTextData = _ConvertOutputTextData(mOutputTextData)

		'出力用テキストのデータを出力用テキストに変換
		Dim mDelimiteText As New OutputDelimiterText(mOutputTextData, pDelimiter)

		'出力用テキストヘッダー情報を取得する
		Dim mHeaderArray As String() = CType(cFolderFileListHeaders.DelimiterHeaderNames.ToArray(GetType(String)), String())

		'ヘッダーをセット
		mDelimiteText.HeaderText = mDelimiteText.ConvertArrayToDelimiterRowData(mHeaderArray, pDelimiter)

		Return mDelimiteText

	End Function

	''' <summary>出力テキスト用のデータに変換</summary>
	''' <param name="pData">GridViewに表示しているデータ</param>
	''' <returns>出力テキスト用のデータ</returns>
	''' <remarks>出力テキスト用に不必要なカラムを削除して返す
	'''          GridViewに表示されていないカラムを削除する  </remarks>
	Private Function _ConvertOutputTextData(ByVal pData As DataTable)

		'出力テキスト用に不必要なカラム情報を削除する
		pData.Columns.Remove(FolderFileListColumn.No.ToString)
		pData.Columns.Remove(FolderFileListColumn.FileSystemType.ToString)
		pData.Columns.Remove(FolderFileListColumn.Size.ToString)
		pData.Columns.Remove(FolderFileListColumn.SizeLevel.ToString)
		pData.Columns.Remove(FolderFileListColumn.DirectoryLevel.ToString)
		pData.Columns.Remove(FolderFileListColumn.ParentFolderFullPath.ToString)
		pData.Columns.Remove(FolderFileListColumn.IsLastFileInFolder.ToString)
		pData.Columns.Remove(FolderFileListColumn.DispString.ToString)

		Return pData

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

			'Alt+Tabに表示させない
			MyBase.SetShowHideAltTabWindow(AltTabType.Hide)

			Select Case CommandLine.Instance.Output

				Case CommandLine.OutputType.ClipBoard

					'出力用テキストをクリップボードコピー
					Clipboard.SetText(_GetOutputFileText(CommandLine.Instance.Extension))

                    'クリップボードにコピーしました通知を表示
                    MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeCommandMessageCopyToClipBoard)

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

                    '出力用テキスト保存・実行処理
                    If _SaveRunOutputFile(CommandLine.Instance.Extension, _FolderFileList.TargetPathFolderName, mSaveFileEncord) Then

                        '保存したファイルを実行しました通知を表示
                        MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageRunSaveFile)

                        '非同期でMessenger風通知メッセージが非表示になるまで待機
                        Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

                    End If

                    'メインフォームのリソースを破棄する
                    '※Closeをすると無限ループしてしまうのでDisposeで対応（これでいいのか不明……）
                    '  メインフォームを表示させず閉じるため
                    Me.Owner.Dispose()

			End Select

		End If

	End Sub

#End Region

#Region "外部公開メソッド"

	''' <summary>インスタンスを保持する変数の破棄処理</summary>
	''' <remarks></remarks>
	Public Shared Sub DisposeInstance()

		_instance = Nothing

	End Sub

#End Region

#End Region

End Class