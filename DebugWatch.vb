Option Explicit On

''' <summary>時間計測用のストップウォッチの機能を提供します</summary>
''' <remarks>
''' 時間計測用のストップウォッチです（デバッグ用)
''' 基本的にはConditional("DEBUG")によってリリース時はコンパイルされないため、処理速度に影響しません。
''' ※このクラスを使用する時はStartTimeメソッドを呼び出してから使用すること
''' </remarks>
Public Class DebugWatch

#Region "変数"

	''' <summary>時間計測用のメソッドとプロパティを提供するクラス</summary>
	Private sw As New System.Diagnostics.Stopwatch

	''' <summary>時間計測用のストップウォッチ</summary>
	Private Shared _DebugWatch As DebugWatch

#End Region

#Region "コンストラクタ"

	''' <summary>コンストラクタ</summary>
	''' <remarks>引数無しは外部に公開しない</remarks>
	Private Sub New()

	End Sub

#End Region

#Region "プロパティ"

	''' <summary>ストップウォッチ情報を返却します。</summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks>
	''' デバック用クラスですが、どこでも使用される可能性があるので
	''' 自分以外には設定させ無い目的があります。
	''' ※デザインパターンのSingletonパターンです。
	''' </remarks>
	Public Shared ReadOnly Property Instance() As DebugWatch

		Get

			'インスタンスが作成されてなかったらインスタンスを作成
			If _DebugWatch Is Nothing Then _DebugWatch = New DebugWatch

			Return _DebugWatch

		End Get

	End Property

#End Region

#Region "メソッド"

	'Conditional属性は引数の定義の時のみ実行される。
	'定義のシンボルが無い場合はコンパイル対象外となるため
	'コンパイル後のサイズや挙動に影響を及ぼさない。

	''' <summary>ストップウォッチ開始</summary>
	''' <remarks>時間を計測する時は</remarks>
	<Conditional("DEBUG")> _
	Public Sub StartTime()

		Me.sw.Start()

	End Sub

	''' <summary>ストップウォッチの停止</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Sub StopTime()

		Me.sw.Stop()

	End Sub

	''' <summary>ストップウォッチのリセット</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Sub ResetTime()

		Me.sw.Reset()

	End Sub

	''' <summary>経過時間間隔を取得する</summary>
	''' <returns>時間間隔</returns>
	''' <remarks>インスタンスが作成されてからの時間間隔を返す</remarks>
	Private Function PutTime() As TimeSpan

		Return Me.sw.Elapsed()

	End Function

	''' <summary>経過時間を出力ウインドウに出力</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Sub DebugPrintTime(ByVal putMessage As String)

		Debug.Print(putMessage & ": " & Me.PutTime.ToString())

	End Sub

	''' <summary>経過時間をメッセージボックスで表示</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Sub DebugPrintTimeToMessegeBox(ByVal putMessage As String)

		MessageBox.Show(putMessage & ": " & Me.PutTime.ToString(), "経過時間", MessageBoxButtons.OK, MessageBoxIcon.Information)

	End Sub

	''' <summary>経過時間をMessenger風通知メッセージで表示</summary>
	''' <remarks></remarks>
	<Conditional("DEBUG")> _
	Public Sub DebugPrintTimeToPopupMessage(ByVal pTitle As String, ByVal pPutMessage As String)

		'表示するメッセージを作成
		Dim mMessage As String = pPutMessage & Environment.NewLine & Me.PutTime.ToString()

		'Messenger風通知メッセージを表示
		Dim mPopupMessage As New frmPopupMessage(pTitle, mMessage)
		mPopupMessage.Show()

	End Sub

#End Region

End Class
