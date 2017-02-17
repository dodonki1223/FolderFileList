Option Explicit On

''' <summary>フォルダファイルリスト作成中を表示するフォームを提供する</summary>
''' <remarks>フォームの共通処理のInterface（IFormCommonProcess）を実装する</remarks>
Public Class frmWait
	Implements IFormCommonProcess

#Region "定数"

    ''' <summary>フォームが表示されるまでの時間</summary>
    ''' <remarks></remarks>
    Private Const _cTimeToDisplayForForm As Integer = 3000

#End Region

#Region "変数"

    ''' <summary>フォームのインスタンスを保持する変数</summary>
    ''' <remarks></remarks>
    Private Shared _instance As frmWait

    ''' <summary>フォーム表示用のタイマー</summary>
    ''' <remarks></remarks>
    Private _FormDisplayTimer As Timer

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
    Public Sub _DisplayForm(sender As Object, e As EventArgs)

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

		'ウインドウをAlt+Tabに表示させない
		MyBase.SetShowHideAltTabWindow(AltTabType.Hide)

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

#End Region

#Region "外部公開メソッド"

    ''' <summary>インスタンスを保持する変数の破棄処理</summary>
    ''' <remarks></remarks>
    Public Shared Sub DisposeInstance()

		_instance = Nothing

	End Sub

#End Region

End Class