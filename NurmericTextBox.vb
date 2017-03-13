Option Explicit On

Imports System.ComponentModel
Imports System.Windows.Forms

''' <summary>数値のみの入力を許可したテキストボックスの機能を提供する</summary>
''' <remarks></remarks>
Public Class NurmericTextBox

#Region "プロパティ"

	''' <summary>有効桁数プロパティ</summary>
	''' <remarks>デフォルトは１００００まで</remarks>
	Public Property NumberOfDigit As Integer = 5

#End Region

#Region "コンストラクタ"

	''' <summary>コンストラクタ</summary>
	''' <remarks></remarks>
	Public Sub New()

		' この呼び出しはデザイナーで必要です。
		InitializeComponent()

		' InitializeComponent() 呼び出しの後で初期化を追加します。

		'IMEのONの無効化
		ImeMode = ImeMode.Disable

	End Sub

#End Region

#Region "イベント"

	''' <summary>OnPaintイベント</summary>
	''' <param name="e">Paintイベント</param>
	''' <remarks>描画が必要なとき</remarks>
	Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)

		'Paintイベントを発生させる
		MyBase.OnPaint(e)

	End Sub

	''' <summary>KeyPressイベント</summary>
	''' <param name="e">KeyPressイベント</param>
	''' <remarks>キーが押されたとき</remarks>
	Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)

		'KeyPressイベントを発生させる
		MyBase.OnKeyPress(e)

		'入力されたキーが数値以外 か テキストの内容の桁数が有効桁数よりも以上の時
		If e.KeyChar > "9"c OrElse e.KeyChar < "0"c OrElse Text.Length >= NumberOfDigit Then

			'入力キーがBackSpace以外の時のキーの入力イベントを処理済みにする
			If e.KeyChar <> Chr(8) Then e.Handled = True

		End If

	End Sub

	''' <summary>OnValidatingイベント</summary>
	''' <param name="e">Validatingイベント</param>
	''' <remarks>評価が必要なとき(ロストフォーカス)</remarks>
	Protected Overrides Sub OnValidating(e As CancelEventArgs)

		'Validatingイベントを発生させる
		MyBase.OnValidating(e)

		'入力内容が空欄なら何もしない
		If String.IsNullOrEmpty(Text) Then Return

		'入力内容が数値以外 または 有効桁数より多い場合、ロストフォーカスしない
		If Not IsNumeric(Text) OrElse Text.Length > NumberOfDigit Then

			e.Cancel = True

		End If
	End Sub

	''' <summary>OnEnterイベント</summary>
	''' <param name="e">Enterイベント</param>
	''' <remarks>フォーカスが当たった時</remarks>
	Protected Overrides Sub OnEnter(e As EventArgs)

		'Enterイベントを発生させる
		MyBase.OnEnter(e)

		'テキストボックスの中身を選択状態にする
		SelectAll()

	End Sub

	''' <summary>OnMouseDownイベント</summary>
	''' <param name="e">MouseDownイベント</param>
	''' <remarks>マウスボタンが下がった時</remarks>
	Protected Overrides Sub OnMouseDown(e As MouseEventArgs)

		'MouseDownイベントを発生させる
		MyBase.OnMouseDown(e)

		'テキストボックスの中身を選択状態にする
		SelectAll()

	End Sub

#End Region

End Class