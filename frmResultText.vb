Option Explicit On

Imports FolderFileList.FolderFileList
Imports FolderFileList.CommandLine

''' <summary>
'''   フォルダファイルリストの出力文字列を表示するフォームを提供する
''' </summary>
''' <remarks>
'''   フォームの共通処理のInterface（IFormCommonProcess）を実装する
''' </remarks>
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

        ''' <summary>Messenger風通知メッセージ：クリップボードへコピーしました</summary>
        ''' <remarks></remarks>
        Public Const NoticeMessageCopyToClipBoard As String = "出力文字列をクリップボードにコピーしました"

        ''' <summary>Messenger風通知メッセージ：保存したファイルを実行しました</summary>
        ''' <remarks></remarks>
        Public Const NoticeMessageRunSaveFile As String = "保存したファイルを実行しました"

    End Class

#End Region

#Region "変数"

    ''' <summary>フォルダファイルリストデータ</summary>
    ''' <remarks>FolderFileListプロパティにてデータをセットする</remarks>
    Private _FolderFileList As FolderFileList

    ''' <summary>現在表示されているフォルダファイルリストデータ</summary>
    ''' <remarks></remarks>
    Private _CurrentFolderFileList As DataTable

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

    ''' <summary>
    '''   フォームにアクセスするためのプロパティ
    ''' </summary>
    ''' <remarks>
    '''   ※デザインパターンのSingletonパターンです
    '''     インスタンスがただ１つであることを保証する
    ''' </remarks>
    Public Shared ReadOnly Property Instance() As frmResultText

        Get

            'インスタンスが作成されてなかったらインスタンスを作成
            If _instance Is Nothing Then _instance = New frmResultText

            Return _instance

        End Get

    End Property

    ''' <summary>
    '''   出力文字列フォームインスタンス存在プロパティ
    ''' </summary>
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

    ''' <summary>
    '''   フォルダファイルリストデータセットプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property FolderFileListData() As FolderFileList

        Set(value As FolderFileList)

            _FolderFileList = value
            _CurrentFolderFileList = _FolderFileList.FolderFileList

        End Set

    End Property

    ''' <summary>
    '''   フォルダファイルリストの出力文字列プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property OutputText() As String

        Get

            Return _FolderFileList.GetOutputTextString(_CurrentFolderFileList)

        End Get

    End Property

#End Region

#Region "コンストラクタ"

    ''' <summary>
    '''   コンストラクタ（外部非公開）
    ''' </summary>
    ''' <remarks>
    '''   引数付きのコンストラクタのみ許容させるため
    ''' </remarks>
    Private Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

    End Sub

#End Region

#Region "イベント"

#Region "フォーム"

    ''' <summary>
    '''   フォームロードイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">Loadイベント</param>
    ''' <remarks></remarks>
    Private Sub frmResultText_Load(sender As Object, e As EventArgs) Handles Me.Load

        '初期コントロール設定
        _SetInitialControlInScreen()

    End Sub

    ''' <summary>
    '''   フォームShownイベント
    ''' </summary>
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

        '出力用文字列表示テキストボックスへフォーカスをセット
        Call _FocusTxtFolderFileList()

    End Sub

    ''' <summary>
    '''   フォームリサイズイベント
    ''' </summary>
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

    ''' <summary>
    '''   フォームのKeyDownイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">KeyDownイベント</param>
    ''' <remarks></remarks>
    Private Sub frmResultText_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        'フォーカスのあるコントロールを取得
        Dim mForcusedControl As Control = Me.ActiveControl

        If e.Control Then

            '-------------------------------
            ' Ctrlキーのショートカットキー
            '-------------------------------
            Select Case e.KeyCode.ToString

                Case "F" 'Ctrl+F：「名前（LIKE検索）」へフォーカスをセット

                    txtName.Focus()

                Case "A" 'Ctrl+A：「フォルダファイルリストテキストボックス」の全選択

                    txtFolderFileList.SelectAll()

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

                Case "E" 'Shift+E：「拡張子指定」へフォーカスをセット

                    cmbExtension.Focus()

                    '「拡張子指定」を開いた状態で表示
                    cmbExtension.DroppedDown = True

                Case "O" 'Shift+O：「リスト表示フォーム」ボタンをクリック

                    btnResultGridView.PerformClick()

                Case "H" 'Shift+H：「html出力」ボタンをクリック

                    btnHtmlOutput.PerformClick()

                Case "T" 'Shift+T：「Text出力」ボタンをクリック

                    btnTextOutput.PerformClick()

            End Select

        End If

    End Sub

    ''' <summary>
    '''   フォームのClosedイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">FormClosedイベント</param>
    ''' <remarks>
    '''   フォームが閉じられようとした際に閉じる処理をキャンセルさせる時に使用するのがFormClosing
    '''   フォームが閉じる前提で後処理するのがFormClosed
    ''' </remarks>
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

#Region "絞り込み関係"

    ''' <summary>
    '''   絞り込みイベント
    ''' </summary>
    ''' <param name="sender">イベント対象コントロール</param>
    ''' <param name="e">イベント対象コントロールのイベント</param>
    ''' <remarks>
    '''   ラジオボタンのCheckedDhanged、コンボボックスのTextChanged、ボタンのClick
    ''' </remarks>
    Private Sub RefineByFolderFileList_Event(sender As Object, e As EventArgs) Handles rbtnAll.CheckedChanged _
                                                                                     , rbtnFolder.CheckedChanged _
                                                                                     , cmbExtension.TextChanged _
                                                                                     , btnNameSearch.Click

        'フォルダファイルリストを再表示
        Call _ReDisplayRefineByFolderFileList()

    End Sub

    ''' <summary>
    '''   コントロールのKeyDownイベント
    ''' </summary>
    ''' <param name="sender">イベント対象コントロール</param>
    ''' <param name="e">イベント対象コントロールのKeyDownイベント</param>
    ''' <remarks></remarks>
    Private Sub Controls_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbExtension.KeyDown _
                                                                            , txtName.KeyDown

        'Enterキーが押された時
        If e.KeyCode = Keys.Enter Then

            'フォルダファイルリストを再表示
            Call _ReDisplayRefineByFolderFileList()

        End If

    End Sub

#End Region

#Region "ボタン"

    ''' <summary>
    '''   TEXT出力ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender">Text出力ボタン</param>
    ''' <param name="e">Clickイベント</param>
    ''' <remarks></remarks>
    Private Async Sub btnTextOutput_Click(sender As Object, e As EventArgs) Handles btnTextOutput.Click

        'メッセージボックスから押されたボタンにより処理を分岐
        Select Case MyBase.ShowDialogueMessage(_cOutputTextFileFormat)

            Case Windows.Forms.DialogResult.Yes

                'TXTファイル保存・実行処理
                If _SaveRunOutputFile(OutputFileFormat.TXT, _FolderFileList.TargetPathFolderName, cEncording.ShiftJis) Then

                    '保存したファイルを実行しました通知を表示
                    MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageRunSaveFile)

                    '非同期でMessenger風通知メッセージが非表示になるまで待機
                    Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

                End If

            Case Windows.Forms.DialogResult.No

                'TXTテキストクリップボードコピー
                Clipboard.SetText(Me.OutputText)

                'クリップボードにコピーしました通知を表示
                MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageCopyToClipBoard)

                '非同期でMessenger風通知メッセージが非表示になるまで待機
                Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

        End Select

    End Sub

    ''' <summary>
    '''   HTML出力ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender">html出力ボタン</param>
    ''' <param name="e">Clickイベント</param>
    ''' <remarks></remarks>
    Private Async Sub btnHtmlOutput_Click(sender As Object, e As EventArgs) Handles btnHtmlOutput.Click

        'HTMLファイル保存・実行処理
        If _SaveRunOutputFile(OutputFileFormat.HTML, _FolderFileList.TargetPathFolderName, cEncording.UTF8) Then

            '保存したファイルを実行しました通知を表示
            MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageRunSaveFile)

            '非同期でMessenger風通知メッセージが非表示になるまで待機
            Await Task.Run(Sub() System.Threading.Thread.Sleep(frmPopupMessage._cMessageDisplayTotalTime))

        End If

    End Sub

    ''' <summary>
    '''   リスト表示フォームへボタンクリックイベント
    ''' </summary>
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

    ''' <summary>
    '''   初期画面コントロール設定
    ''' </summary>
    ''' <remarks></remarks>
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

        '出力用文字列テキストボックスを編集不可に
        txtFolderFileList.ReadOnly = True

        '出力用文字列をセットを画面にセット
        txtFolderFileList.Text = _FolderFileList.OutputText

        'フォームのタイトルにコマンドモード文言を追加
        CommandLine.SetCommandModeToTitle(Me)

        'フォームのタイトルにデバッグモード文言を追加（デバッグモード時のみ実行される）
        DebugMode.SetDebugModeTitle(Me)

    End Sub

    ''' <summary>
    '''   画面コントロール設定
    ''' </summary>
    ''' <remarks>
    '''   再描画時にコントロールの再設定を行う
    ''' </remarks>
    Private Sub _SetControlInScreen()

        'チェック済みのラジオボタンを取得
        Dim mRbtnButton As RadioButton = _GetCheckedRadioButton(grpDisplayTarget)

        'ラジオボタンの状態により拡張子コンボボックスとファイル名検索テキストボックスの設定を行う
        Select Case mRbtnButton.Name

            Case rbtnAll.Name

                '拡張子コンボボックスの使用を可に
                cmbExtension.Enabled = True
                cmbExtension.BackColor = Color.White

                '検索ボタンの使用を可に
                btnNameSearch.Enabled = True

                '名前検索文字列入力テキストボックスの入力を可能に
                txtName.ReadOnly = False
                txtName.BackColor = Color.White

            Case rbtnFolder.Name

                '拡張子コンボボックスの使用を不可に、値を初期値（空文字）に設定
                cmbExtension.Enabled = False
                cmbExtension.SelectedIndex = 0
                cmbExtension.BackColor = Color.Gray

                '検索ボタンの使用を不可に
                btnNameSearch.Enabled = False

                '名前検索文字列入力テキストボックスの入力を不可に
                txtName.ReadOnly = True
                txtName.Text = ""
                txtName.BackColor = Color.Gray

        End Select

        '出力用文字列をセット
        txtFolderFileList.Text = Me.OutputText

        '出力用文字列表示テキストボックスへフォーカスをセット
        Call _FocusTxtFolderFileList()

    End Sub

    ''' <summary>
    '''   コントロールのサイズ変更・再配置
    ''' </summary>
    ''' <param name="pChangedSize">フォームの変更後サイズ</param>
    ''' <remarks>
    '''   全てのコントロールを取得しコントロールごと処理を行う
    ''' </remarks>
    Private Sub _ResizeAndRealignmentControls(pChangedSize As Size) Implements IFormCommonProcess._ResizeAndRealignmentControls

        '全てのコントロールを取得
        Dim mAllControls As Control() = GetControlsInTarget(Me)

        For Each mTargetControl As Control In mAllControls

            Select Case True

                Case TypeOf mTargetControl Is TextBox

                    'テキストボックスが「txtFolderFileList」の時
                    If mTargetControl.Name = txtFolderFileList.Name Then

                        'テキストボックスの場合はフォームのサイズの変更分、サイズを変更する
                        mTargetControl.Size = New System.Drawing.Size(mTargetControl.Size + pChangedSize)

                    End If

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

    ''' <summary>
    '''   出力用文字列表示テキストボックスへのフォーカス処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub _FocusTxtFolderFileList()

        'フォーカスを出力用文字列表示テキストボックスへ
        txtFolderFileList.Focus()

        '選択位置をなしに設定 ※全選択状態の回避のため
        txtFolderFileList.Select(0, 0)

    End Sub

    ''' <summary>
    '''   チェックされているラジオボタンを取得
    ''' </summary>
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

    ''' <summary>
    '''   絞り込み条件でフォルダフォルダファイルリストを再表示
    ''' </summary>
    ''' <remarks>
    '''   絞り込み条件で絞り込んだフォルダファイルリストを再表示させる
    ''' </remarks>
    Private Sub _ReDisplayRefineByFolderFileList()

        'フォルダファイルリストデータが存在しない場合は処理を終了
        If _FolderFileList Is Nothing Then Exit Sub

        '絞り込んだ状態のフォルダファイルリストを取得
        _CurrentFolderFileList = _GetRefineByFolderFileListForDisplay()

        '画面コントロール設定
        Call _SetControlInScreen()

    End Sub

    ''' <summary>
    '''   絞り込みデータを取得
    ''' </summary>
    ''' <returns>絞り込みデータ</returns>
    ''' <remarks>
    '''   画面内の条件に一致する絞り込みデータを返す
    ''' </remarks>
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

        End Select

        '----------------------------------
        ' 表示対象データを拡張子で絞り込んだ
        ' データを取得
        '----------------------------------
        Dim mExtensionWhere As String = cmbExtension.SelectedItem
        mDisplayData = _FolderFileList.GetRefineByFolderFileListIncludeFolder(mDisplayData, FolderFileListColumn.Extension, mExtensionWhere)

        '----------------------------------
        ' 表示対象データを名前(LIKE検索)で
        ' 絞り込んだデータを取得
        '----------------------------------
        Dim mNameWhere As String = txtName.Text
        mDisplayData = _FolderFileList.GetRefineByFolderFileListIncludeFolder(mDisplayData, FolderFileListColumn.Name, mNameWhere)

        '絞り込みデータの「フォルダ内で最後のファイルかどうか」を再セットする
        _FolderFileList.SetIsLastFileInFolder(mDisplayData)

        '出力用の文字列を再作成する
        _FolderFileList.CreateOutputTextString(mDisplayData)

        Return mDisplayData

    End Function

#End Region

#Region "「Html・TEXT」出力用メソッド"

    ''' <summary>
    '''   出力ファイルの保存・実行処理
    ''' </summary>
    ''' <param name="pFileFormat">出力形式</param>
    ''' <param name="pDefalutFileName">保存ファイルのデフォルトのファイル名</param>
    ''' <param name="pEncording">エンコード</param>
    ''' <returns>True ：保存ファイルの実行、False：保存ファイルを実行しない</returns>
    ''' <remarks></remarks>
    Private Function _SaveRunOutputFile(ByVal pFileFormat As OutputFileFormat, ByVal pDefalutFileName As String, ByVal pEncording As System.Text.Encoding) As Boolean

        '名前を付けて保存ダイアログを表示
        Dim mDialog As SaveFileDialog = MyBase.GetSaveAsDialog(pDefalutFileName, pFileFormat.ToString.ToLower)

        '名前をつけて保存ダイアログでOKが押されたら
        If mDialog.ShowDialog = Windows.Forms.DialogResult.OK Then

            'ファイルの保存処理
            MyBase.WriteTextToOutputFile(_GetOutputFileText(pFileFormat), mDialog.FileName, pEncording)

            '実行タイプが実行の時
            If Settings.Instance.Execute = ExecuteType.Run Then

                '保存したファイルを関連付けで実行する
                MyBase.RunFile(mDialog.FileName)
                Return True

            End If

        End If

        Return False

    End Function

    ''' <summary>
    '''   出力用ファイルのテキストデータを取得
    ''' </summary>
    ''' <param name="pFileFormat">出力形式</param>
    ''' <returns>対象出力形式のテキストデータ</returns>
    ''' <remarks></remarks>
    Private Function _GetOutputFileText(ByVal pFileFormat As OutputFileFormat) As String

        Dim mOutputText As String = String.Empty

        '出力形式ごと処理を分岐
        Select Case pFileFormat

            Case OutputFileFormat.TXT

                'フォルダファイルリストの出力文字列を取得
                mOutputText = Me.OutputText

            Case OutputFileFormat.HTML

                'フォルダファイルリストから出力用のHtml文字列を取得
                Dim mOutputHtml As New OutputHtml(_FolderFileList, CommandLine.FormType.Text)
                mOutputText = mOutputHtml.HtmlSentence

        End Select

        Return mOutputText

    End Function

#End Region

#Region "コマンド処理メソッド"

    ''' <summary>
    '''   コマンド処理を実行
    ''' </summary>
    ''' <remarks>
    '''   コマンドタイプによりそれぞれの処理を行う
    ''' </remarks>
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
                    MyBase.ShowPopupMessage(_cMessage.NoticeMessageTitle, _cMessage.NoticeMessageCopyToClipBoard)

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

#End Region

#Region "外部公開メソッド"

    ''' <summary>
    '''   インスタンスを保持する変数の破棄処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub DisposeInstance()

        _instance = Nothing

    End Sub

#End Region

End Class