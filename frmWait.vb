Option Explicit On

''' <summary>フォルダファイルリスト作成中を表示するフォームを提供する</summary>
''' <remarks>フォームの共通処理のInterface（IFormCommonProcess）を実装する</remarks>
Public Class frmWait
    Implements IFormCommonProcess

#Region "定数"

    ''' <summary>フォームが表示されるまでの時間</summary>
    ''' <remarks>３秒</remarks>
    Private Const _cTimeToDisplayForForm As Integer = 3000

    ''' <summary>作成中文字列の表示切り替えるまでの時間</summary>
    ''' <remarks>１秒</remarks>
    Private Const _cTimeToSwitchingMakingString As Integer = 1000

    ''' <summary>フォルダファイルリスト作成中フォーム画面で使用するメッセージを提供する</summary>
    ''' <remarks></remarks>
    Public Class _cMessage

        ''' <summary>作成中文字列の共通メッセージ</summary>
        ''' <remarks></remarks>
        Public Const Making As String = "フォルダファイルリスト作成中"

        ''' <summary>共通メッセージの後に表示する文字列「．」</summary>
        ''' <remarks>作成中文字列の表示切り替えるまでの時間ごとこの文字列を増やしていく</remarks>
        Public Const MaikingDot As String = "．"

        ''' <summary>共通メッセージの後に表示する文字列の最高カウント数</summary>
        ''' <remarks></remarks>
        Public Const MakingDotMaxCount As Integer = 6

        ''' <summary>フォルダ・ファイル数計算中</summary>
        ''' <remarks></remarks>
        Public Const Calculating As String = "対象フォルダ内のフォルダ・ファイル数を計算しています"

    End Class

#End Region

#Region "変数"

    ''' <summary>フォームのインスタンスを保持する変数</summary>
    ''' <remarks></remarks>
    Private Shared _instance As frmWait

    ''' <summary>フォーム表示用のタイマー</summary>
    ''' <remarks></remarks>
    Private _FormDisplayTimer As Timer

    ''' <summary>作成中文字列の表示切り替え用タイマー</summary>
    ''' <remarks></remarks>
    Private _SwitchingMakingStringTimer As Timer

#End Region

#Region "プロパティ"

    ''' <summary>フォームにアクセスするためのプロパティ</summary>
    ''' <remarks>※デザインパターンのSingletonパターンです
    '''            インスタンスがただ１つであることを保証する</remarks>
    Public Shared ReadOnly Property Instance() As frmWait

        Get

            'インスタンスが作成されてなかったらインスタンスを作成
            If _instance Is Nothing Then _instance = New frmWait

            Return _instance

        End Get

    End Property

    ''' <summary>フォルダファイルリスト作成中を表示するフォームインスタンス存在プロパティ</summary>
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

#End Region

#Region "イベント"

#Region "フォーム"

    ''' <summary>フォームロードイベント</summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">Loadイベント</param>
    ''' <remarks></remarks>
    Private Sub frmWait_Load(sender As Object, e As EventArgs) Handles Me.Load

        '初期コントロール設定
        Call _SetInitialControlInScreen()

        'フォーム表示用タイマー設定
        Call _SetFormDisplayTimer()

        '作成中文字列の表示切り替え用タイマー設定
        Call _SetSwitchingMakingStringTimer()

    End Sub

    ''' <summary>フォームのFormClosingイベント</summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">FormClosingイベント</param>
    ''' <remarks>フォームが閉じられようとした際に閉じる処理をキャンセルさせる時に使用するのがFormClosing
    '''          フォームが閉じる前提で後処理するのがFormClosed                                          </remarks>
    Private Sub frmWait_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        '作成中フォームのインスタンスが存在しない時は処理を終了
        If Not frmWait.HasInstance Then Exit Sub

        Dim mMsgBoxTitle As String = "プログラムの終了"
        Dim mMsgBoxText As String = "プログラムを終了しますか？"

        'メッセージボックスを表示し押されたボタンにより処理を分岐
        Select Case MessageBox.Show(mMsgBoxTitle, mMsgBoxText, MessageBoxButtons.YesNo)

            Case Windows.Forms.DialogResult.Yes

                'フォームの閉じる処理を続行

            Case Windows.Forms.DialogResult.No

                'フォームの閉じる処理をキャンセル
                e.Cancel = True

        End Select

    End Sub

    ''' <summary>フォームのClosedイベント</summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">FormClosedイベント</param>
    ''' <remarks>フォームが閉じられようとした際に閉じる処理をキャンセルさせる時に使用するのがFormClosing
    '''          フォームが閉じる前提で後処理するのがFormClosed                                          </remarks>
    Private Sub frmWait_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        '作成中フォームのインスタンスが存在しない時は処理を終了
        If Not frmWait.HasInstance Then Exit Sub

        'インスタンスへの参照を破棄
        _instance = Nothing

        '画面の内容を設定ファイルへ書き込み 
        Dim mFrmMain As frmMain = DirectCast(Me.Owner, frmMain)
        mFrmMain.WriteScreenSettingToXmlFile()

        'メインフォームのリソースを破棄する
        '※Closeをすると無限ループしてしまうのでDisposeで対応（これでいいのか不明……）
        Me.Owner.Dispose()

    End Sub

#End Region

#Region "フォーム表示用タイマー"

    ''' <summary>フォーム表示イベント</summary>
    ''' <param name="sender">Timerオブジェクト</param>
    ''' <param name="e">Tickイベント</param>
    ''' <remarks></remarks>
    Public Sub _DisplayForm(ByVal sender As Object, ByVal e As EventArgs)

        'ウインドウをAlt+Tabに表示させる
        MyBase.SetShowHideAltTabWindow(AltTabType.Show)

        'フォームを透明状態から元に戻す
        Me.Opacity = 1

        'フォーム表示用タイマーを停止
        _FormDisplayTimer.Stop()

        'フォームk表示用タイマーを破棄
        _FormDisplayTimer.Dispose()

    End Sub

#End Region

#Region "作成中文字列の表示切り替え用タイマー"

    ''' <summary>作成中文字列切り替え</summary>
    ''' <param name="sender">Timerオブジェクト</param>
    ''' <param name="e">Tickイベント</param>
    ''' <remarks></remarks>
    Public Sub _SwitchingMakingString(ByVal sender As Object, ByVal e As EventArgs)

        '「．」文字列をカウントする
        Dim mDotCount As Integer = _GetCountCharForTargetChar(lblMaking.Text, _cMessage.MaikingDot)

        '「．」文字列のカウント数が「．」文字列の最高カウント数と同じだったら
        If mDotCount = _cMessage.MakingDotMaxCount Then

            'フォームに現在表示されている文字列（「．」を含まない）＋「．」をセット
            lblMaking.Text = lblMaking.Text.Replace(_cMessage.MaikingDot, "") & _cMessage.MaikingDot

        Else

            'フォームに現在表示されている文字列（「．」を含む）＋「．」をセット
            lblMaking.Text = lblMaking.Text & _cMessage.MaikingDot

        End If

    End Sub

#End Region

#End Region

#Region "メソッド"

    ''' <summary>初期画面コントロール設定</summary>
    ''' <remarks></remarks>
    Public Sub _SetInitialControlInScreen() Implements IFormCommonProcess._SetInitialControlInScreen

        'フォームを透明状態にする
        Me.Opacity = 0

        'フォームサイズを初期表示と同じ大きさから変更出来なくする
        Me.MaximumSize = Me.Size
        Me.MinimumSize = Me.Size

        '最大化ボタンを非表示にする
        Me.MaximizeBox = False

        'メッセージ設定
        Me.lblMaking.Text = _cMessage.Calculating
        Me.tsslProcessingFolderFile.Text = String.Empty

        'ウインドウをAlt+Tabに表示させない
        MyBase.SetShowHideAltTabWindow(AltTabType.Hide)

        'フォームのタイトルにデバッグモード文言を追加（デバッグモード時のみ実行される）
        DebugMode.SetDebugModeTitle(Me)

    End Sub

    ''' <summary>コントロールのサイズ変更・再配置</summary>
    ''' <param name="pChangedSize">フォームの変更後サイズ</param>
    ''' <remarks>全てのコントロールを取得しコントロールごと処理を行う</remarks>
    Public Sub _ResizeAndRealignmentControls(pChangedSize As Size) Implements IFormCommonProcess._ResizeAndRealignmentControls

        'フォームの大きさを固定なので処理を行わない

    End Sub

    ''' <summary>フォーム表示用タイマー設定</summary>
    ''' <remarks></remarks>
    Private Sub _SetFormDisplayTimer()

        'インスタンスを作成
        _FormDisplayTimer = New Timer

        'フォーム表示用タイマーにTickイベントを設定
        AddHandler _FormDisplayTimer.Tick, New EventHandler(AddressOf _DisplayForm)

        'Tickイベントが発生する間隔を設定
        _FormDisplayTimer.Interval = _cTimeToDisplayForForm

        'フォーム表示用タイマーを起動
        _FormDisplayTimer.Start()

    End Sub

    ''' <summary>作成中文字列の表示切り替え用タイマー設定</summary>
    ''' <remarks></remarks>
    Private Sub _SetSwitchingMakingStringTimer()

        'インスタンスを作成
        _SwitchingMakingStringTimer = New Timer

        '作成中文字列の表示切り替え用タイマーにTickイベントを設定
        AddHandler _SwitchingMakingStringTimer.Tick, New EventHandler(AddressOf _SwitchingMakingString)

        'Tickイベントが発生する間隔を設定
        _SwitchingMakingStringTimer.Interval = _cTimeToSwitchingMakingString

        'フォーム表示用タイマーを起動
        _SwitchingMakingStringTimer.Start()

    End Sub

    ''' <summary>対象文字列の中に特定の文字の出現回数を取得</summary>
    ''' <param name="pTargetString">対象文字列</param>
    ''' <param name="pCountString">カウント文字列</param>
    ''' <returns>対象文字列内にあるカウント文字列の出現回数</returns>
    ''' <remarks></remarks>
    Private Function _GetCountCharForTargetChar(ByVal pTargetString As String, ByVal pCountString As String) As Integer

        '「対象文字列数 - 対象文字列からカウント文字列を引いた数」を返す
        Return pTargetString.Length - pTargetString.Replace(pCountString, "").Length

    End Function

#End Region

#Region "外部公開メソッド"

    ''' <summary>インスタンスを保持する変数の破棄処理</summary>
    ''' <remarks></remarks>
    Public Shared Sub DisposeInstance()

        _instance = Nothing

    End Sub

#End Region

End Class