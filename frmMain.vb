Option Explicit On

Imports System.IO
Imports System.Threading
Imports FolderFileList.FolderFileList

''' <summary>
'''   フォルダ選択画面を表示するフォームを提供する
''' </summary>
''' <remarks>
'''   フォームの共通処理のInterface（IFormCommonProcess）を実装する
''' </remarks>
Public Class frmMain
    Implements IFormCommonProcess

#Region "定数"

    ''' <summary>初期表示時のウインドウサイズ</summary>
    Private Shared ReadOnly _cDefaultWindowSize As New System.Drawing.Size(New System.Drawing.Point(670, 80))

    ''' <summary>設定ボタン押下後のウインドウサイズ</summary>
    Private Shared ReadOnly _cOpenSettingWindowSize As New System.Drawing.Size(New System.Drawing.Point(670, 170))

    ''' <summary>フォルダ選択画面で使用するメッセージを提供する</summary>
    ''' <remarks></remarks>
    Private Class _cMessage

        ''' <summary>メッセージボックスタイトル文字列（エラー）</summary>
        ''' <remarks></remarks>
        Public Const MessageBoxTitleError As String = "エラー"

        ''' <summary>フォルダパスが未入力メッセージ</summary>
        ''' <remarks></remarks>
        Public Const FolderPathNothing As String = "フォルダパスが入力されていません"

        ''' <summary>選択フォルダ数超過メッセージ</summary>
        ''' <remarks></remarks>
        Public Const FolderCountExceeded As String = "選択できるフォルダは１件です"

        ''' <summary>フォルダパスでないメッセージ</summary>
        ''' <remarks></remarks>
        Public Const NotFolderPath As String = "選択されたパスはフォルダではありません"

        ''' <summary>リスト表示の最大表示ファイル数の形式が正しくありませんメッセージ</summary>
        ''' <remarks></remarks>
        Public Const InCorrectMaxCountInPage As String = "リスト表示の最大表示ファイル数の値が正しくありません"

    End Class

#End Region

#Region "列挙体"

    ''' <summary>メッセージボックス表示区分</summary>
    Public Enum ShowMessageBox

        ''' <summary>表示させる</summary>
        Yes

        ''' <summary>表示させない</summary>
        No

    End Enum

#End Region

#Region "変数"

    ''' <summary>リサイズ前ウインドウサイズ</summary>
    ''' <remarks></remarks>
    Private _WindowSize As System.Drawing.Size

    ''' <summary>リサイズ後の変更分ウインドウサイズ</summary>
    ''' <remarks></remarks>
    Private _ChangedWindowSize As System.Drawing.Size

#End Region

#Region "イベント"

#Region "フォーム"

    ''' <summary>
    '''   フォームロードイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">Loadイベント</param>
    ''' <remarks></remarks>
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        '初期コントロール設定
        _SetInitialControlInScreen()

    End Sub

    ''' <summary>
    '''   フォームShownイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">Shownイベント</param>
    ''' <remarks></remarks>
    Private Sub frmMain_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        '処理タイプにより処理を分岐
        Select Case CommandLine.Instance.ProcessMode

            Case CommandLine.ProcessType.Help

                'コマンドリストを表示しプログラムを終了
                CommandLine.ShowCommnadList()
                Me.Close()

            Case CommandLine.ProcessType.Command

                'コマンド処理を実行
                Call _RunCommandProcess()

            Case Else

                'フォームを透明状態から元に戻す
                Me.Opacity = 1

        End Select

    End Sub

    ''' <summary>
    '''   フォームリサイズイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">Resizeイベント</param>
    ''' <remarks></remarks>
    Private Sub frmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize

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
    '''   フォームのClosedイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">FormClosedイベント</param>
    ''' <remarks></remarks>
    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        '画面の内容を設定ファイルへ書き込み
        Call WriteScreenSettingToXmlFile()

    End Sub

    ''' <summary>
    '''   フォームのKeyDownイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">KeyDownイベント</param>
    ''' <remarks></remarks>
    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.Control Then

            '-------------------------------
            ' Ctrlキーのショートカットキー
            '-------------------------------
            Select Case e.KeyCode.ToString

                Case "P" 'Ctrl+P：「フォルダパス」へフォーカスをセット

                    txtFolderPath.Focus()

                Case "F" 'Ctrl+F：「フォルダ選択」ボタンを実行

                    btnSelectFolder.PerformClick()

                Case "R" 'Ctrl+R：「実行」ボタンを実行

                    btnRun.PerformClick()

                Case "S" 'Ctrl+S：「設定」ボタンを実行

                    btnSetting.PerformClick()

                Case "O" 'Ctrl+O：「出力文字列」を選択

                    '設定ボタンを押した時のウインドウサイズの時、出力文字列をチェック
                    If Me.Size = _cOpenSettingWindowSize Then rbtnResultText.Checked = True

                Case "L" 'Ctrl+L：「リスト表示」を選択

                    '設定ボタンを押した時のウインドウサイズの時、リスト表示をチェック
                    If Me.Size = _cOpenSettingWindowSize Then rbtnResultGridView.Checked = True

                Case "K" 'Ctrl+K：「保存ファイル即時実行」のチェック

                    '設定ボタンを押した時のウインドウサイズの時、保存ファイルの即時実行をチェック・アンチェック
                    If Me.Size = _cOpenSettingWindowSize Then chkRunSaveFile.Checked = Not (chkRunSaveFile.Checked)

                Case "M" 'Ctrl+M：「最大表示ファイル数（リスト表示）」へフォーカスをセット

                    '設定ボタンを押した時のウインドウサイズの時、実行
                    If Me.Size = _cOpenSettingWindowSize Then txtMaxCountInPage.Focus()

                Case "D" 'Ctrl+D：「FolderFileListのドキュメント」URLへ遷移（既定のブラウザで開く）

                    '設定ボタンを押した時のウインドウサイズの時、「FolderFileListのドキュメント」URLへ遷移（既定のブラウザで開く）
                    If Me.Size = _cOpenSettingWindowSize Then MyBase.RunFile(OutputHtml.cFolderFileListDocumentURL)

            End Select

        End If

    End Sub

    ''' <summary>
    '''   フォームのDragEnterイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">DragEnterイベント</param>
    ''' <remarks></remarks>
    Private Sub frmMain_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter

        'ファイルをドラッグ中のマウス・カーソルがフォーム上に入ったときの処理
        Call _FileDragEnter(e)

    End Sub

    ''' <summary>
    '''   フォームのDragDropイベント
    ''' </summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">DragDropイベント</param>
    ''' <remarks></remarks>
    Private Sub frmMain_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop

        'ドラッグ＆ドロップされたフォルダパスを取得
        Dim mFolders As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'フォルダのドラッグ＆ドロップ処理
        Call _FolderDragDrop(mFolders)

    End Sub

#End Region

#Region "テキストボックス"

    ''' <summary>
    '''   フォルダ選択テキストボックスGotFocusイベント
    ''' </summary>
    ''' <param name="sender">フォルダ選択テキストボックス</param>
    ''' <param name="e">GotFocusイベント</param>
    ''' <remarks></remarks>
    Private Sub txtFolderPath_GotFocus(sender As Object, e As EventArgs) Handles txtFolderPath.GotFocus

        'テキストボックスの内容を全選択状態に
        txtFolderPath.SelectAll()

    End Sub

    ''' <summary>
    '''   フォルダ選択テキストボックスのKeyDownイベント
    ''' </summary>
    ''' <param name="sender">フォルダ選択テキストボックス</param>
    ''' <param name="e">KeyDownイベント</param>
    ''' <remarks></remarks>
    Private Sub txtFolderPath_KeyDown(sender As Object, e As KeyEventArgs) Handles txtFolderPath.KeyDown

        'Enterキーが押された時
        If e.KeyCode = Keys.Enter Then

            '「実行ボタン」をクリック
            btnRun.PerformClick()

        End If

    End Sub

    ''' <summary>
    '''   フォルダ選択テキストボックスのKeyPressイベント
    ''' </summary>
    ''' <param name="sender">フォルダ選択テキストボックス</param>
    ''' <param name="e">KeyPressイベント</param>
    ''' <remarks></remarks>
    Private Sub txtFolderPath_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtFolderPath.KeyPress

        '押されたキーがEnterまたはEscapeの時
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) _
        OrElse e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Escape) Then

            'キーの入力イベントを処理済みにする（ビープ音を鳴らさない）
            e.Handled = True

        End If

    End Sub

    ''' <summary>
    '''   フォルダ選択テキストボックスのMouseDownイベント
    ''' </summary>
    ''' <param name="sender">フォルダ選択テキストボックス</param>
    ''' <param name="e">MouseDownイベント</param>
    ''' <remarks></remarks>
    Private Sub txtFolderPath_MouseDown(sender As Object, e As MouseEventArgs) Handles txtFolderPath.MouseDown

        'テキストボックスの内容を全選択状態に
        txtFolderPath.SelectAll()

    End Sub

    ''' <summary>
    '''   フォルダ選択テキストボックスのDragEnterイベント
    ''' </summary>
    ''' <param name="sender">txtFolderPathオブジェクト</param>
    ''' <param name="e">DragEnterイベント</param>
    ''' <remarks></remarks>
    Private Sub txtFolderPath_DragEnter(sender As Object, e As DragEventArgs) Handles txtFolderPath.DragEnter

        'ファイルをドラッグ中のマウス・カーソルがフォーム上に入ったときの処理
        Call _FileDragEnter(e)

    End Sub

    ''' <summary>
    '''   フォルダ選択テキストボックスのDragDropイベント
    ''' </summary>
    ''' <param name="sender">txtFolderPathオブジェクト</param>
    ''' <param name="e">DragDropイベント</param>
    ''' <remarks></remarks>
    Private Sub txtFolderPath_DragDrop(sender As Object, e As DragEventArgs) Handles txtFolderPath.DragDrop

        'ドラッグ＆ドロップされたフォルダパスを取得
        Dim mFolders As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'フォルダのドラッグ＆ドロップ処理
        Call _FolderDragDrop(mFolders)

    End Sub

#End Region

#Region "ボタン"

    ''' <summary>
    '''   フォルダ選択ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender">フォルダ選択ボタン</param>
    ''' <param name="e">Clickイベント</param>
    ''' <remarks></remarks>
    Private Sub btnSelectFolder_Click(sender As Object, e As EventArgs) Handles btnSelectFolder.Click

        Dim mDialog = New FolderBrowserDialog

        '「新しいフォルダーの作成(N)」ボタンの非表示
        mDialog.ShowNewFolderButton = False

        '「OK」ボタンが押された時
        If mDialog.ShowDialog() = DialogResult.OK Then

            '選択されたパスを対象フォルダテキストボックスにセット
            txtFolderPath.Text = mDialog.SelectedPath

        End If

        mDialog = Nothing

    End Sub

    ''' <summary>
    '''   実行ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender">実行ボタン</param>
    ''' <param name="e">Clickイベント</param>
    ''' <remarks></remarks>
    Private Async Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click

        Dim mFolderPath As String = txtFolderPath.Text

        'フォルダパスが存在する時 かつ リスト表示の最大表示ファイル数の値が正しい時
        If _IsExistFolderPath(mFolderPath, ShowMessageBox.Yes) AndAlso _ValidateMaxCountInPageValue(ShowMessageBox.Yes) Then

            'メインフォームを非表示
            Me.Hide()

            '非同期でフォルダファイルリスト作成
            Dim mFolderFileList As FolderFileList = Await _CreateFolderFileList(mFolderPath)

            'フォルダファイルリストのデータを子フォームにセット
            frmResultText.Instance.FolderFileListData = mFolderFileList
            frmResultGridView.Instance.FolderFileListData = mFolderFileList

            'リスト表示フォームに最大表示ファイル数をセット
            frmResultGridView.Instance.MaxCountInPage = txtMaxCountInPage.Text

            '子フォームの親フォームにメインフォームを設定
            frmResultText.Instance.Owner = Me
            frmResultGridView.Instance.Owner = Me

            '画面の内容を設定ファイルへ書き込み 
            Call WriteScreenSettingToXmlFile()

            '対処フォームのラジオボタンのチェック状態により表示するフォームを変更する
            '※★フォームの表示について★
            '  ①モードレス表示:Form.Show()
            '    開いているフォームを閉じなくても、ほかのフォームとの間で自由にフォーカスを移動できる
            '  ②モーダル表示  :Form.ShowDialog()
            '    アプリケーションのほかの作業に移行する前に、終了する（非表示にする）必要がある
            '    「Me」を渡すとメインフォームをオーナーに設定出来る（オーナーが閉じると全てのフォームが閉じられる）
            Select Case True

                Case rbtnResultText.Checked = True

                    '出力文字列フォームをモードレスで表示
                    frmResultText.Instance.Show()

                Case rbtnResultGridView.Checked = True

                    'リスト表示フォームをモードレスで表示
                    frmResultGridView.Instance.Show()

            End Select

        End If

    End Sub

    ''' <summary>
    '''   設定ボタンクリックイベント
    ''' </summary>
    ''' <param name="sender">設定ボタン</param>
    ''' <param name="e">Clickイベント</param>
    ''' <remarks></remarks>
    Private Sub btnSetting_Click(sender As Object, e As EventArgs) Handles btnSetting.Click

        '現在のウインドウサイズがデフォルトのウインドウサイズだったら
        If Me.Size = _cDefaultWindowSize Then

            '設定ボタン押下後のウインドウサイズに変更
            Me.Size = _cOpenSettingWindowSize

        Else

            'デフォルトのウインドウサイズに変更
            Me.Size = _cDefaultWindowSize

        End If

    End Sub

#End Region

#Region "ラベル"

    ''' <summary>
    '''   「FolderFileListのドキュメントはこちら」ラベルのClickイベント
    ''' </summary>
    ''' <param name="sender">「FolderFileListのドキュメントはこちら」ラベル</param>
    ''' <param name="e">Clickイベント</param>
    ''' <remarks></remarks>
    Private Sub lblDocument_Click(sender As Object, e As EventArgs) Handles lblDocument.Click

        'FolderFileListのドキュメントURLへ遷移（既定のブラウザで開く）
        MyBase.RunFile(OutputHtml.cFolderFileListDocumentURL)

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

        'ウインドウの最小サイズを再設定
        '※ウインドウの最小サイズはfrmCommonで最初に表示したサイズで設定されるため改めて再設定する
        Me.MinimumSize = _cDefaultWindowSize

        'ウインドウサイズをデフォルトウインドウサイズに設定
        Me.Size = _cDefaultWindowSize

        'ユーザーのウインドウサイズの変更を不可に
        Me.FormBorderStyle = FormBorderStyle.FixedSingle

        '最大化ボタンの使用を禁止に
        Me.MaximizeBox = Not Me.MaximizeBox

        'フォームへのドラッグ＆ドロップを許可する
        Me.AllowDrop = True

        'フォルダ選択テキストボックスへのドラッグ＆ドロップを許可する
        txtFolderPath.AllowDrop = True

        '設定ファイルの内容を画面に適用
        SetXmlFileSettingToScreen()

        'フォームのタイトルにデバッグモード文言を追加（デバッグモード時のみ実行される）
        DebugMode.SetDebugModeTitle(Me)

    End Sub

    ''' <summary>
    '''   コントロールのサイズ変更・再配置
    ''' </summary>
    ''' <param name="pChangedSize">フォームの変更後サイズ</param>
    ''' <remarks></remarks>
    Public Sub _ResizeAndRealignmentControls(pChangedSize As Size) Implements IFormCommonProcess._ResizeAndRealignmentControls

        '※ウインドウのサイズを変更不可にしたのでこの処理は行わない

    End Sub

#End Region

#Region "ドラッグ＆ドロップメソッド"

    ''' <summary>
    '''   ファイルをドラッグ中のマウス・カーソルがフォーム上に入ったときの処理
    ''' </summary>
    ''' <param name="e">DragEventArgsオブジェクト</param>
    ''' <remarks>
    '''   ファイルのときのみ受け付ける
    ''' </remarks>
    Private Sub _FileDragEnter(ByVal e As DragEventArgs)

        'コントロール内にドラッグされたとき実行される
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            'ドラッグされたデータ形式を調べ、ファイルのときはコピーとする
            e.Effect = DragDropEffects.Copy

        Else

            'ファイル以外は受け付けない
            e.Effect = DragDropEffects.None

        End If

    End Sub

    ''' <summary>
    '''   フォルダのドラッグドロップ処理
    ''' </summary>
    ''' <param name="pFolders">ドラッグされたフォルダパス</param>
    ''' <remarks></remarks>
    Private Sub _FolderDragDrop(ByVal pFolders As String())

        'ドラッグされたフォルダが１件以上の時
        If pFolders.GetUpperBound(0) >= 1 Then

            '「選択できるフォルダは１件です」メッセージを表示
            MessageBox.Show(_cMessage.FolderCountExceeded, _cMessage.MessageBoxTitleError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub

        End If

        'フォルダのパスを取得
        Dim mFolderPath As String = pFolders(0)

        'フォルダのパスにファイルが存在する時
        If System.IO.File.Exists(mFolderPath) Then

            '「選択されたパスはフォルダパスではありません」メッセージを表示
            MessageBox.Show(_cMessage.NotFolderPath, _cMessage.MessageBoxTitleError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub

        End If

        'テキストボックスにフォルダパスをセット
        txtFolderPath.Text = mFolderPath

    End Sub

#End Region

#Region "値チェックメソッド"

    ''' <summary>
    '''   フォルダパスが存在するか
    ''' </summary>
    ''' <param name="pPath">対象パス</param>
    ''' <param name="pShowMessageBox">メッセージボックスを表示させるかどうか</param>
    ''' <returns>True：フォルダパスが存在する、False：フォルダパスが存在しない</returns>
    ''' <remarks></remarks>
    Private Function _IsExistFolderPath(ByVal pPath As String, Optional ByVal pShowMessageBox As ShowMessageBox = ShowMessageBox.No) As Boolean

        '入力パスが空の時
        If String.IsNullOrEmpty(pPath) Then

            'メッセージボックス表示区分がYesだったら
            If pShowMessageBox = ShowMessageBox.Yes Then

                '「フォルダパスが入力されていません」メッセージを表示
                MessageBox.Show(_cMessage.FolderPathNothing, _cMessage.MessageBoxTitleError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End If

            Return False

        End If

        '入力パスがディレクトリで無かった時
        If Not System.IO.Directory.Exists(pPath) Then

            'メッセージボックス表示区分がYesだったら
            If pShowMessageBox = ShowMessageBox.Yes Then

                '「選択されたパスがフォルダではありません」メッセージを表示
                MessageBox.Show(_cMessage.NotFolderPath, _cMessage.MessageBoxTitleError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End If

            Return False

        End If

        Return True

    End Function

    ''' <summary>
    '''   リスト表示の最大表示ファイル数の値が正しいか
    ''' </summary>
    ''' <param name="pShowMessageBox">メッセージボックスを表示させるかどうか</param>
    ''' <returns>True：正しい値、False：正しくない値</returns>
    ''' <remarks></remarks>
    Private Function _ValidateMaxCountInPageValue(Optional ByVal pShowMessageBox As ShowMessageBox = ShowMessageBox.No) As Boolean

        'リスト表示の最大表示ファイル数が空 または リスト表示の最大表示ファイル数が０の時
        If String.IsNullOrEmpty(txtMaxCountInPage.Text) OrElse txtMaxCountInPage.Text = 0 Then

            'メッセージボックス表示区分がYesだったら
            If pShowMessageBox = ShowMessageBox.Yes Then

                '「リスト表示の最大表示ファイル数の値が正しくありません」メッセージを表示
                MessageBox.Show(_cMessage.InCorrectMaxCountInPage, _cMessage.MessageBoxTitleError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End If

            Return False

        Else

            Return True

        End If

    End Function

#End Region

#Region "フォルダファイルリスト作成"

    ''' <summary>
    '''   フォルダファイルリストを作成
    ''' </summary>
    ''' <param name="pPath">対象フォルダパス</param>
    ''' <returns>対象フォルダパスのフォルダファイルリスト</returns>
    ''' <remarks>
    '''   フォルダファイルリストを非同期で作成し、作成中の時は作成中フォームを表示する
    ''' </remarks>
    Private Async Function _CreateFolderFileList(ByVal pPath As String) As Task(Of FolderFileList)

        '時間の計測開始（デバッグモード時のみ実行される）
        DebugMode.StartDebugWatch()

        '子フォームの親フォームにメインフォームを設定
        frmWait.Instance.Owner = Me

        '作成中フォームをモードレスで表示
        frmWait.Instance.Show()

        'フォルダファイルリストの入れ物を作成
        Dim mFolderFileList As New FolderFileList(pPath)

        '処理進捗プロパティをセット
        mFolderFileList.ProcessProgress = New Progress(Of FolderFileListProgress)(AddressOf ShowFolderFileListProgress)

        Try

            '非同期でフォルダファイルリストを作成する
            Await Task.Run(
                           Sub()

                               'フォルダファイルリスト作成の進捗報告をONにする
                               mFolderFileList.IsReportProcessProgress = True

                               'フォルダファイルリストを作成
                               mFolderFileList.CreateFolderFileList()

                               '出力文字列を作成
                               mFolderFileList.CreateOutputTextString()

                           End Sub
                          )

        Catch ex As Exception

            Throw

        Finally

            '作成中フォームを非表示
            frmWait.Instance.Hide()

            '作成中フォームのインスタンスを破棄
            frmWait.DisposeInstance()

        End Try

        '時間の計測終了、経過時間を表示（デバッグモード時のみ実行される）
        DebugMode.StopDebugWatchShowProcessingTime("ﾌｫﾙﾀﾞﾌｧｲﾙﾘｽﾄ作成時間", "フォルダファイルリストの作成時間は")

        '非同期でMessenger風通知メッセージが非表示になるまで待機（デバッグモード時のみ実行される）
        Await Task.Run(Sub() DebugMode.WaitTillClosingPopupMessage())

        Return mFolderFileList

    End Function

    ''' <summary>
    '''   フォルダファイルリスト作成中フォームへ進捗状況を表示
    ''' </summary>
    ''' <param name="pProgress">フォルダファイルリスト進捗状況報告用</param>
    ''' <remarks></remarks>
    Public Sub ShowFolderFileListProgress(ByVal pProgress As FolderFileListProgress)

        'フォルダファイルリスト作成中ラベルを取得
        '※「．」文字列を除いた部分
        Dim mMakingText As String = frmWait.Instance.lblMaking.Text.Replace(frmWait._cMessage.MaikingDot, "")

        'フォルダファイルリスト作成中ラベルが「対象フォルダ内のフォルダ・ファイル数を計算しています」の時
        If mMakingText = frmWait._cMessage.Calculating Then

            '「対象フォルダ内のフォルダ・ファイル数を計算しています」文言から「フォルダファイルリスト作成中」に変更
            frmWait.Instance.lblMaking.Text = frmWait._cMessage.Making

        End If

        'フォルダファイルリスト作成中フォームのプログレスバーに進捗率をセット
        frmWait.Instance.tspbProgressRate.Value = pProgress.Percent

        'フォルダファイルリスト作成中フォームのステータスを表示するラベルに進捗率をセット
        frmWait.Instance.tsslStatus.Text = pProgress.Percent & "％完了"

        '処理フォルダファイルをセット
        frmWait.Instance.tsslProcessingFolderFile.Text = pProgress.ProcessingFolderFile

    End Sub

#End Region

#Region "設定ファイルメソッド"

    ''' <summary>
    '''   設定ファイルの内容を画面にセット
    ''' </summary>
    ''' <remarks>
    '''   Config.xmlファイルから取得した設定内容を画面に適用する
    ''' </remarks>
    Public Sub SetXmlFileSettingToScreen()

        '設定ファイルの読み込み
        Settings.LoadFromXmlFile()

        '表示フォームを画面に適用
        Select Case Settings.Instance.TargetForm

            Case CommandLine.FormType.Text

                rbtnResultText.Checked = True
                rbtnResultGridView.Checked = False

            Case CommandLine.FormType.List

                rbtnResultText.Checked = False
                rbtnResultGridView.Checked = True

        End Select

        '保存ファイルの即時実行区分を画面に適用
        Select Case Settings.Instance.Execute

            Case CommandLine.ExecuteType.None

                '保存ファイルの即時実行区分にアンチェック
                chkRunSaveFile.Checked = False

            Case CommandLine.ExecuteType.Run

                '保存ファイルの即時実行区分にチェック
                chkRunSaveFile.Checked = True

        End Select

        'リスト表示の最大表示ファイル数をセット
        txtMaxCountInPage.Text = Settings.Instance.MaxCountInPage

    End Sub

    ''' <summary>
    '''   設定ファイルへ画面の内容を書き込む
    ''' </summary>
    ''' <remarks>
    '''   Config.xmlファイルへ画面の内容を書き込む
    ''' </remarks>
    Public Sub WriteScreenSettingToXmlFile()

        '対象フォーム設定を取得し設定クラスにセット
        Select Case True

            Case rbtnResultText.Checked = True

                Settings.Instance.TargetForm = CommandLine.FormType.Text

            Case rbtnResultGridView.Checked = True

                Settings.Instance.TargetForm = CommandLine.FormType.List

        End Select

        '保存ファイルの即時実行区分を取得し設定クラスにセット
        Select Case True

            Case chkRunSaveFile.Checked = False

                Settings.Instance.Execute = CommandLine.ExecuteType.None

            Case chkRunSaveFile.Checked = True

                Settings.Instance.Execute = CommandLine.ExecuteType.Run

        End Select

        'リスト表示の最大表示ファイル数が正しい値の時
        If _ValidateMaxCountInPageValue() Then

            'リスト表示の最大表示ファイル数テキストボックスの内容を設定クラスにセット
            Settings.Instance.MaxCountInPage = txtMaxCountInPage.Text

        Else

            'リスト表示の最大表示ファイル数のデフォルトを設定クラスにセット
            Settings.Instance.MaxCountInPage = frmResultGridView.cGridViewDataMaxCountInPage

        End If

        '設定ファイルへの書き込み処理
        Settings.SaveToXmlFile()

    End Sub

#End Region

#Region "コマンド処理メソッド"

    ''' <summary>
    '''   コマンド処理を実行
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub _RunCommandProcess()

        '出力形式が指定なし以外の時
        If CommandLine.Instance.Output <> CommandLine.OutputType.None Then

            'ウインドウをAlt+Tabに表示させない
            MyBase.SetShowHideAltTabWindow(AltTabType.Hide)

        End If

        'コマンドライン引数を表示（デバッグモード時のみ実行される）
        DebugMode.ShowCommandLineArg()

        'コマンドライン引数の内容を画面にセット
        Call _SetCommandLineArgToScreen()

        '実行ボタンをクリック
        btnRun.PerformClick()

        'フォームを透明状態から元に戻す
        Me.Opacity = 1

    End Sub

    ''' <summary>
    '''   コマンドライン引数の内容を画面にセット
    ''' </summary>
    ''' <remarks>
    '''   コマンドライン引数から取得した内容を画面に適用する
    ''' </remarks>
    Private Sub _SetCommandLineArgToScreen()

        'フォルダパスをセット
        txtFolderPath.Text = CommandLine.Instance.FolderPath

        '表示フォームを画面に適用
        Select Case CommandLine.Instance.TargetForm

            Case CommandLine.FormType.Text

                rbtnResultText.Checked = True
                rbtnResultGridView.Checked = False

            Case CommandLine.FormType.List

                rbtnResultText.Checked = False
                rbtnResultGridView.Checked = True

        End Select

        '保存ファイルの即時実行区分を画面に適用
        Select Case CommandLine.Instance.Execute

            Case CommandLine.ExecuteType.None

                chkRunSaveFile.Checked = False

            Case CommandLine.ExecuteType.Run

                chkRunSaveFile.Checked = True

        End Select

        'リスト表示の最大表示ファイル数を画面に適用
        txtMaxCountInPage.Text = CommandLine.Instance.MaxCountInPage

    End Sub

#End Region

#End Region

End Class