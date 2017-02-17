Option Explicit On

Imports System.Threading

''' <summary>Messenger風通知メッセージを表示するウインドウの機能を提供する</summary>
''' <remarks>ディスプレイの幅取得用のメモ
'''            WorkingArea：デスクトップ全体の範囲（タスクバーとか除く）
'''            Bounds　　 ：デスクトップ全体の範囲（タスクバーとか含む）</remarks>
Public Class frmPopupMessage
	Implements IFormCommonProcess

#Region "定数"

	''' <summary>ウインドウサイズ</summary>
	Private Shared ReadOnly _cDefaultWindowSize As New System.Drawing.Size(New System.Drawing.Point(230, 240))

	''' <summary>ポップアップ・アンポップアップのインターバル時間</summary>
	''' <remarks>０．０１秒</remarks>
	Private Const _cPopupIntervalTime As Integer = 10

	''' <summary>ウインドウが移動する刻み</summary>
	''' <remarks></remarks>
	Private Const _cMoveWindowNotch As Integer = 5

    ''' <summary>Messenger風通知メッセージの表示時間</summary>
    ''' <remarks>３秒</remarks>
    Public Const _cMessageDisplayTime As Integer = 3000

    ''' <summary>Messenger風通知メッセージが表示されるトータル時間</summary>
    ''' <remarks>５秒</remarks>
    Public Const _cMessageDisplayTotalTime As Integer = 5000

#End Region

#Region "変数"

    ''' <summary>フォームポップアップ用タイマー</summary>
    ''' <remarks></remarks>
    Private _PopupTimer As System.Windows.Forms.Timer

	''' <summary>フォームアンポップアップ用タイマー</summary>
	''' <remarks></remarks>
	Private _UnPopupTimer As System.Windows.Forms.Timer

	''' <summary>ウインドウのタイトル</summary>
	''' <remarks>プロパティ用変数</remarks>
	Private _FormTitle As String

	''' <summary>表示するメッセージ</summary>
	''' <remarks>プロパティ用変数</remarks>
	Private _Message As String

#End Region

#Region "プロパティ"

	''' <summary>Messenger風通知メッセージのタイトル</summary>
	''' <remarks></remarks>
	Public WriteOnly Property Title As String

		Set(value As String)

			_FormTitle = value

		End Set

	End Property

	''' <summary>画面に表示するメッセージ</summary>
	''' <remarks></remarks>
	Public WriteOnly Property Message As String

		Set(value As String)

			_Message = value

		End Set

	End Property

#End Region

#Region "コンストラクタ"

	''' <summary>コンストラクタ</summary>
	''' <remarks></remarks>
	Public Sub New()

		' この呼び出しはデザイナーで必要です。
		InitializeComponent()

		' InitializeComponent() 呼び出しの後で初期化を追加します。

	End Sub

	''' <summary>コンストラクタ</summary>
	''' <param name="pMessage">表示させるメッセージ</param>
	''' <remarks></remarks>
	Public Sub New(ByVal pMessage As String)

		' この呼び出しはデザイナーで必要です。
		InitializeComponent()

		' InitializeComponent() 呼び出しの後で初期化を追加します。

		'表示メッセージを設定
		_Message = pMessage

	End Sub

	''' <summary>コンストラクタ</summary>
	''' <param name="pMessage">表示させるメッセージ</param>
	''' <remarks></remarks>
	Public Sub New(ByVal pTitle As String, ByVal pMessage As String)

		' この呼び出しはデザイナーで必要です。
		InitializeComponent()

		' InitializeComponent() 呼び出しの後で初期化を追加します。

		'Messenger風通知メッセージのタイトル設定
		_FormTitle = pTitle

		'表示メッセージを設定
		_Message = pMessage

	End Sub

#End Region

#Region "イベント"

#Region "フォーム"

	''' <summary>フォームロードイベント</summary>
	''' <param name="sender">Formオブジェクト</param>
	''' <param name="e">Loadイベント</param>
	''' <remarks></remarks>
	Private Sub frmPopupMessage_Load(sender As Object, e As EventArgs) Handles MyBase.Load

		'初期コントロール設定
		Call _SetInitialControlInScreen()

		'フォームポップアップ用タイマー設定
		Call _SetPopupTimer()

	End Sub

#End Region

#Region "タイマー"

	''' <summary>ポップアップ用タイマーTickイベント</summary>
	''' <param name="sender">Timerオブジェクト</param>
	''' <param name="e">Tickイベント</param>
	''' <remarks></remarks>
	Private Sub PopupMessage_Tick(sender As Object, e As EventArgs)

		'プライマリーディスプレイの高さ（タスクバーとか除く）を取得
		'※プライマリーディスプレイのタスクバーの上の位置
		Dim mDisplayY As Integer = Screen.PrimaryScreen.WorkingArea.Height

		'タスクバーの上の位置よりウインドウの下の位置が大きい時
		If mDisplayY < Me.Top + Me.Height Then

			'ウインドウのY座標を-5する
			'※ウインドウが上昇
			Me.Top -= _cMoveWindowNotch

		Else

			'Messenger風通知メッセージを閉じる
			Call UnPopupMessage()

		End If

	End Sub

	''' <summary>アンポップアップ用タイマーTickイベント</summary>
	''' <param name="sender">Timerオブジェクト</param>
	''' <param name="e">Tickイベント</param>
	Private Sub UnPopupMessage_Tick(sender As Object, e As EventArgs)

		'プライマリーディスプレイの高さ（タスクバーとか含む）を取得
		'※プライマリーディスプレイの一番下の位置
		Dim mDisplayY As Integer = Screen.PrimaryScreen.Bounds.Height

		'ウインドウの上がプリマリーディスプレイの一番下の位置より小さい時
		If Me.Top < mDisplayY Then

			'ウインドウのY座標を+5する
			'※ウインドウが下降
			Me.Top += _cMoveWindowNotch

		Else

			'ウインドウを閉じる
			Me.Close()

		End If

	End Sub

#End Region

#End Region

#Region "メソッド"

#Region "コントロール設定メソッド"

	''' <summary>初期画面コントロール設定</summary>
	''' <remarks></remarks>
	Private Sub _SetInitialControlInScreen() Implements IFormCommonProcess._SetInitialControlInScreen

		'--------------------------
		' フォームの表示位置を設定
		'--------------------------
		'表示X座標 = プライマリーディスプレイの幅 - frmPopupMessageウインドウの幅
		Dim mDisplayX As Integer = Screen.PrimaryScreen.Bounds.Width - Me.Width

		'表示Y座標 = プライマリーディスプレイの高さ
		Dim mDisplayY As Integer = Screen.PrimaryScreen.Bounds.Height

		'フォームの表示位置を設定
		Me.DesktopLocation = New Point(mDisplayX, mDisplayY)

		'--------------------------
		' フォーム設定
		'--------------------------
		'フォームのサイズを設定
		Me.Size = _cDefaultWindowSize

		'フォームを最前面化
		Me.TopMost = True

		'最大化ボタンを非表示
		Me.MaximizeBox = False

		'最小化ボタンを非表示
		Me.MinimizeBox = False

		'タイトルを設定
		Me.Text = _FormTitle

		'メッセージを設定
		lblMessage.Text = _Message

		'ウインドウをAlt+Tabに表示させない
		MyBase.SetShowHideAltTabWindow(AltTabType.Hide)

	End Sub

	''' <summary>フォームポップアップ用タイマー設定</summary>
	''' <remarks></remarks>
	Private Sub _SetPopupTimer()

		'インスタンスを作成
		_PopupTimer = New System.Windows.Forms.Timer

		'フォームポップアップ用タイマーにTickイベントを設定
		AddHandler _PopupTimer.Tick, New EventHandler(AddressOf PopupMessage_Tick)

		'Tickイベントが発生する間隔を設定
		_PopupTimer.Interval = _cPopupIntervalTime

		'フォームポップアップ用タイマーを起動
		_PopupTimer.Start()

	End Sub

	''' <summary>フォームアンポップアップ用タイマー設定</summary>
	''' <remarks></remarks>
	Private Sub _SetUnPopupTimer()

		'インスタンスを作成
		_UnPopupTimer = New System.Windows.Forms.Timer

		'フォームアンポップアップ用タイマーにTickイベントを設定
		AddHandler _UnPopupTimer.Tick, New EventHandler(AddressOf UnPopupMessage_Tick)

		'Tickイベントが発生する間隔を設定
		_UnPopupTimer.Interval = _cPopupIntervalTime

		'フォームアンポップアップ用タイマーを起動
		_UnPopupTimer.Start()

	End Sub

	''' <summary>コントロールのサイズ変更・再配置</summary>
	''' <param name="pChangedSize">フォームの変更後サイズ</param>
	''' <remarks>全てのコントロールを取得しコントロールごと処理を行う</remarks>
	Private Sub _ResizeAndRealignmentControls(pChangedSize As Size) Implements IFormCommonProcess._ResizeAndRealignmentControls

		'※ウインドウのサイズを変更不可にしたのでこの処理は行わない

	End Sub

#End Region

#Region "ポップアップメソッド"

	''' <summary>Messenger風通知メッセージを閉じる</summary>
	''' <remarks>非同期でポップアップメッセージを閉じる</remarks>
	Private Async Sub UnPopupMessage()

		'フォームポップアップ用タイマーを破棄
		_PopupTimer.Stop()
		_PopupTimer.Dispose()

        '非同期で指定時間待機
        Await Task.Run(Sub() Thread.Sleep(_cMessageDisplayTime))

		'フォームアンポップアップ用タイマー設定
		Call _SetUnPopupTimer()

	End Sub

#End Region

#End Region

End Class