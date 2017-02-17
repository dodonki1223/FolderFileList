Option Explicit On

Imports System.Threading

''' <summary>集約例外クラス</summary>
''' <remarks>未処理例外を処理する機能を提供する</remarks>
Public Class AggregateException

#Region "列挙体"

	''' <summary>エラーメッセージタイプ</summary>
	''' <remarks></remarks>
	Public Enum ErrorMessageType

		''' <summary>メッセージボックス</summary>
		MessageBox

		''' <summary>イベントログ</summary>
		EventLog

	End Enum

	''' <summary>イベントハンドラタイプ</summary>
	''' <remarks></remarks>
	Public Enum EventHandlerType

		''' <summary>Windowsフォームアプリケーションで捕捉されなかった例外</summary>
		ThreadException

		''' <summary>UIスレッドで例外がスローされた時はThreadExceptionイベントが発生しそれ以外の例外</summary>
		UnhandledException

	End Enum

#End Region

#Region "メイン処理"

	''' <summary>メイン処理</summary>
	''' <remarks></remarks>
	<STAThread()> _
 Shared Sub Main()

		'ThreadExceptionイベントハンドラを登録する
		AddHandler Application.ThreadException, AddressOf Application_ThreadException

		'UnhandledExceptionイベントハンドラを登録する
		AddHandler Thread.GetDomain().UnhandledException, AddressOf Application_UnhandledException

		'コントロ―ルの外観をビジュアルスタイル（XPスタイル）にする
		Application.EnableVisualStyles()

		'フォルダ選択画面を表示するフォームの実行
		'※ここで実行されるメインアプリケーションスレッドの例外はApplication_ThreadExceptionでハンドルできる
		Application.Run(New frmMain)

	End Sub

#End Region

#Region "イベント"

	''' <summary>未処理例外をキャッチするイベントハンドラ（ThreadException）</summary>
	''' <param name="sender"></param>
	''' <param name="e"></param>
	''' <remarks>Windowsフォームアプリケーションでは、捕捉されなかった例外がス
	'''          ローされるとApplication.ThreadExceptionイベントが発生します。
	'''          ThreadExceptionイベントが発生するのは、Windowsフォームが作成、
	'''          所有しているスレッド（UIスレッド）で例外がスローされた時だけ
	'''          です。例えば、Thread.Startメソッドや、デリゲートのBeginInvoke
	'''          メソッドなどで開始されたスレッドで発生した例外では発生しません。
	'''          ThreadExceptionイベントハンドラが呼び出された時に何もしないと、
	'''          アプリケーションは続行され、例外は無視されます。例外を無視した
	'''          ままアプリケーションを続行させるのは危険ですので、通常は
	'''          ThreadExceptionイベントハンドラにアプリケーションを終了させる
	'''          コードを記述します。                                          </remarks>
	Public Shared Sub Application_ThreadException(ByVal sender As Object, ByVal e As ThreadExceptionEventArgs)

		'画面へメッセージの表示
		Call ShowErrorMessage(e.Exception, EventHandlerType.ThreadException)

		'エラーをイベントログへ書き込む
		Call WriteErrorToEventLog(e.Exception, EventHandlerType.ThreadException)

		'アプリケーションの終了
		Application.Exit()

	End Sub

	''' <summary>未処理例外をキャッチするイベントハンドラ（UnhandledException）</summary>
	''' <param name="sender"></param>
	''' <param name="e"></param>
	''' <remarks>UnhandledExceptionイベントはThreadExceptionイベントとは違い、
	'''          UIスレッド以外のスレッドで例外がスローされた場合でも発生します。
	'''          また、Windowsフォームアプリケーションだけでなく、コンソールアプ
	'''          リケーションでも使用できます。さらにThreadExceptionイベントとは
	'''          違い、UnhandledExceptionイベントハンドラが呼び出された後も例外
	'''          が無視されるということはなく、そのままにすると通常は「（アプリ
	'''          ケーション）は動作を停止しました。」のようなダイアログが表示さ
	'''          れ、アプリケーションは終了します。                              </remarks>
	Public Shared Sub Application_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)

		Dim mEx As Exception = CType(e.ExceptionObject, Exception)

		If Not mEx Is Nothing Then

			'画面へメッセージの表示
			Call ShowErrorMessage(mEx, EventHandlerType.UnhandledException)

			'エラーをイベントログへ書き込む
			Call WriteErrorToEventLog(mEx, EventHandlerType.UnhandledException)

			'アプリケーションの終了
			Application.Exit()

		End If

	End Sub

#End Region

#Region "メソッド"

	''' <summary>エラーメッセージを取得する</summary>
	''' <param name="pEx">Exceptionクラス</param>
	''' <param name="pMessageType">エラーメッセージタイプ</param>
	''' <param name="pEventHandler">イベントハンドラタイプ</param>
	''' <returns>エラーメッセージ</returns>
	''' <remarks>エラーメッセージタイプに応じたエラーメッセージを返す</remarks>
	Public Shared Function GetErrorMessage(ByVal pEx As Exception, ByVal pMessageType As ErrorMessageType, ByVal pEventHandler As EventHandlerType) As String

		Dim mErrorMessage As New System.Text.StringBuilder

		'エラーメッセージタイプにより処理を分岐
		Select Case pMessageType

			Case ErrorMessageType.MessageBox

				With mErrorMessage

					.AppendLine("エラーが発生しました。開発元にお知らせ下さい。")
					.AppendLine()
					.AppendLine("【エラー内容】")
					.AppendLine(pEx.Message)

				End With

			Case ErrorMessageType.EventLog

				With mErrorMessage

					.AppendLine("【例外通知イベントハンドラ】")

					Select Case pEventHandler

						Case EventHandlerType.ThreadException

							.AppendLine("Application.ThreadException")

						Case EventHandlerType.UnhandledException

							.AppendLine("Application.UnhandledException")

					End Select

					.AppendLine()
					.AppendLine("【エラー内容】")
					.AppendLine(pEx.Message)
					.AppendLine()
					.AppendLine("【スタックトレース】")
					.AppendLine(pEx.StackTrace)

				End With

		End Select

		Return mErrorMessage.ToString

	End Function

	''' <summary>エラーメッセージを画面に表示</summary>
	''' <param name="pEx">Exceptionクラス</param>
	''' <param name="pEventHandler">イベントハンドラタイプ</param>
	''' <remarks></remarks>
	Public Shared Sub ShowErrorMessage(ByVal pEx As Exception, ByVal pEventHandler As EventHandlerType)

		'メッセージボックスのエラーメッセージを取得
		Dim mMessgeBoxErrorString As String = GetErrorMessage(pEx, ErrorMessageType.MessageBox, pEventHandler)

		'メッセージボックスの表示処理
		MessageBox.Show(mMessgeBoxErrorString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

	End Sub

	''' <summary>エラーをイベントログに書き込み</summary>
	''' <param name="pEx">Exceptionクラス</param>
	''' <param name="pEventHandler">イベントハンドラタイプ</param>
	''' <remarks></remarks>
	Public Shared Sub WriteErrorToEventLog(ByVal pEx As Exception, ByVal pEventHandler As EventHandlerType)

		'メッセージボックスのエラーメッセージを取得
		Dim mEventLogString As String = GetErrorMessage(pEx, ErrorMessageType.EventLog, pEventHandler)

		'イベントログにメッセージを書き込む
		Call WriteEventLog(pEx, mEventLogString)

	End Sub

	''' <summary>イベントログに書き込む</summary>
	''' <param name="pEx">Exceptionクラス</param>
	''' <remarks></remarks>
	Public Shared Sub WriteEventLog(ByVal pEx As Exception, ByVal pErrorMessage As String)

		'ソース名を取得
		Dim mSourceName As String = pEx.Source

		'ソースが存在していない時は作成する
		If Not EventLog.SourceExists(mSourceName) Then EventLog.CreateEventSource(mSourceName, "")

		'イベントログにエントリを書き込む
		EventLog.WriteEntry(mSourceName, pErrorMessage, EventLogEntryType.Error)

	End Sub

#End Region

End Class
