Option Explicit On

''' <summary>コマンドラインの機能を提供する</summary>
''' <remarks>※処理タイプの判断基準
'''            ①ヘルプキーワードが存在したらヘルプ処理
'''            ②フォルダパスが取得できたらコマンド処理
'''            ③フォルダパスを取得出来なかったら通常処理 </remarks>
Public Class CommandLine

#Region "定数"

	''' <summary>コマンドライン引数のキーワードを提供する</summary>
	''' <remarks></remarks>
	Public Class cKeyWords

		''' <summary>ヘルプキーワード</summary>
		''' <remarks></remarks>
		Public Const Help As String = "/?"

		''' <summary>フォームタイプキーワード</summary>
		''' <remarks></remarks>
		Public Const Form As String = "/Form"

		''' <summary>出力形式キーワード</summary>
		''' <remarks></remarks>
		Public Const Output As String = "/Output"

		''' <summary>拡張子キーワード</summary>
		''' <remarks></remarks>
		Public Const Extension As String = "/Extension"

		''' <summary>最大表示ファイル数キーワード</summary>
		''' <remarks>リスト表示フォームの１ページ内の最大表示ファイル数</remarks>
		Public Const PageSize As String = "/PageSize"

	End Class

	''' <summary>出力文字列フォーム対象拡張子リスト</summary>
	''' <remarks></remarks>
	Public Shared cTextFormExtensionList As New ArrayList From {OutputFileFormat.TEXT, OutputFileFormat.HTML}

	''' <summary>リスト表示フォーム対象拡張子リスト</summary>
	''' <remarks></remarks>
	Public Shared cListFormExtensionList As New ArrayList From {OutputFileFormat.CSV, OutputFileFormat.TSV, OutputFileFormat.HTML}

#End Region

#Region "列挙体"

	''' <summary>フォームの種類</summary>
	''' <remarks></remarks>
	Public Enum FormType

		''' <summary>フォルダファイルリストの出力文字列を表示するフォーム</summary>
		''' <remarks></remarks>
		Text

		''' <summary>フォルダファイルリストをGridViewに表示するフォーム</summary>
		''' <remarks></remarks>
		List

	End Enum

	''' <summary>出力形式タイプ</summary>
	''' <remarks></remarks>
	Public Enum OutputType

		''' <summary>指定なし</summary>
		''' <remarks></remarks>
		None

		''' <summary>クリップボードにコピー</summary>
		''' <remarks></remarks>
		ClipBoard

		''' <summary>名前を付けて保存ダイアログ</summary>
		''' <remarks></remarks>
		SaveDialog

	End Enum

	''' <summary>処理タイプ</summary>
	''' <remarks></remarks>
	Public Enum ProcessType

		''' <summary>通常処理</summary>
		''' <remarks></remarks>
		Normarl

		''' <summary>コマンド処理</summary>
		''' <remarks></remarks>
		Command

		''' <summary>ヘルプ処理</summary>
		''' <remarks></remarks>
		Help


	End Enum

	''' <summary>ファイルの出力形式</summary>
	''' <remarks></remarks>
	Public Enum OutputFileFormat

		''' <summary>HTML</summary>
		''' <remarks></remarks>
		HTML

		''' <summary>TEXT</summary>
		''' <remarks></remarks>
		TEXT

		''' <summary>CSV</summary>
		''' <remarks></remarks>
		CSV

		''' <summary>TSV</summary>
		''' <remarks></remarks>
		TSV

	End Enum

#End Region

#Region "変数"

	''' <summary>コマンドライン格納変数</summary>
	Private _Commands As CommandList

	''' <summary>フォームのインスタンスを保持する変数</summary>
	''' <remarks></remarks>
	Private Shared _instance As CommandLine

#End Region

#Region "構造体"

	''' <summary>コマンドライン構造体</summary>
	''' <remarks></remarks>
	Private Structure CommandList

		''' <summary>フォルダパス</summary>
		Private Shared _FolderPath As String = String.Empty

		''' <summary>ヘルプコマンドが存在するか</summary>
		Private Shared _HasHelpCommand As Boolean = False

		''' <summary>フォームタイプコマンド</summary>
		''' <remarks>デフォルトで出力文字列フォームをセット</remarks>
		Private Shared _FormType As FormType = FormType.Text

		''' <summary>出力形式タイプコマンド</summary>
		''' <remarks>デフォルトは指定なしをセット</remarks>
		Private Shared _Output As OutputType = OutputType.None

		''' <summary>拡張子コマンド</summary>
		''' <remarks>デフォルトはTEXTをセット</remarks>
		Private Shared _Extension As OutputFileFormat = OutputFileFormat.TEXT

		''' <summary>最大表示ファイル数コマンド</summary>
		''' <remarks>デフォルトで1000をセット</remarks>
		Private Shared _PageSize As Integer = frmResultGridView.cGridViewDataMaxCountInPage

		''' <summary>フォルダパスプロパティ</summary>
		''' <remarks></remarks>
		Public Property FolderPath As String

			Set(value As String)

				_FolderPath = value

			End Set
			Get

				Return _FolderPath

			End Get

		End Property

		''' <summary>ヘルプコマンドが存在するかプロパティ</summary>
		''' <remarks></remarks>
		Public Property HasHelpCommand As Boolean

			Set(value As Boolean)

				_HasHelpCommand = value

			End Set
			Get

				Return _HasHelpCommand

			End Get

		End Property

		''' <summary>フォームタイププロパティ</summary>
		''' <remarks></remarks>
		Public Property FormType As CommandLine.FormType

			Set(value As CommandLine.FormType)

				_FormType = value

			End Set
			Get

				Return _FormType

			End Get

		End Property

		''' <summary>出力形式プロパティ</summary>
		''' <remarks></remarks>
		Public Property Output As OutputType

			Set(value As OutputType)

				_Output = value

			End Set
			Get

				Return _Output

			End Get

		End Property

		''' <summary>拡張子プロパティ</summary>
		''' <remarks></remarks>
		Public Property Extension As OutputFileFormat

			Set(value As OutputFileFormat)

				_Extension = value

			End Set
			Get

				Return _Extension

			End Get

		End Property

		''' <summary>最大表示ファイル数プロパティ</summary>
		''' <remarks></remarks>
		Public Property PageSize As Integer

			Set(value As Integer)

				_PageSize = value

			End Set
			Get

				Return _PageSize

			End Get

		End Property

	End Structure

	''' <summary>コマンド</summary>
	''' <remarks>コマンドキーワードとその値を提供する</remarks>
	Private Structure Command

		''' <summary>キーワード</summary>
		Private Shared _KeyWord As String = String.Empty

		''' <summary>値</summary>
		Private Shared _Value As String = String.Empty

		''' <summary>キーワードプロパティ</summary>
		''' <remarks></remarks>
		Public Property KeyWord As String

			Set(value As String)

				_KeyWord = value

			End Set
			Get

				Return _KeyWord

			End Get

		End Property

		''' <summary>値プロパティ</summary>
		''' <remarks></remarks>
		Public Property Value As String

			Set(value As String)

				_Value = value

			End Set
			Get

				Return _Value

			End Get

		End Property

		''' <summary>引数付きコンストラクタ</summary>
		''' <param name="pCommandLineArg">コマンドライン引数</param>
		''' <param name="pKeyWord">キーワード</param>
		''' <remarks></remarks>
		Sub New(ByVal pCommandLineArg As String, ByVal pKeyWord As String)

			'       コマンドライン引数よりもキーワードの方が文字数が多い時
			'または 'コマンドライン引数の文字数よりキーワード＋１の文字数の方が大きかったら
			If pCommandLineArg.Length < pKeyWord.Length _
			OrElse pCommandLineArg.Length < pKeyWord.Length + 1 Then

				'空文字をセット
				KeyWord = String.Empty
				Value = String.Empty

			Else

				'コマンドライン引数からキーワード部分をセット
				KeyWord = pCommandLineArg.Substring(0, pKeyWord.Length)

				'KeyWord=○○○から値プロパティに「○○○」の部分をセットする
				Value = pCommandLineArg.Substring(KeyWord.Length + 1)

			End If

		End Sub

	End Structure

#End Region

#Region "プロパティ"

	''' <summary>CommandLineクラスにアクセスするためのプロパティ</summary>
	''' <remarks>※デザインパターンのSingletonパターンです
	'''            インスタンスがただ１つであることを保証する</remarks>
	Public Shared ReadOnly Property Instance() As CommandLine

		Get

			'インスタンスが作成されてなかったらインスタンスを作成
			If _instance Is Nothing Then _instance = New CommandLine

			Return _instance

		End Get

	End Property

	''' <summary>フォルダパスプロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property FolderPath As String

		Get

			Return _Commands.FolderPath

		End Get

	End Property

	''' <summary>ヘルプコマンドが存在するかプロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property HasHelpCommand As Boolean

		Get

			Return _Commands.HasHelpCommand

		End Get

	End Property

	''' <summary>フォームタイププロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property TargetForm As CommandLine.FormType

		Get

			Return _Commands.FormType

		End Get

	End Property

	''' <summary>出力形式プロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property Output As OutputType

		Get

			Return _Commands.Output

		End Get

	End Property

	''' <summary>拡張子プロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property Extension As OutputFileFormat

		Get

			Return _Commands.Extension

		End Get

	End Property

	''' <summary>最大表示ファイル数プロパティ</summary>
	''' <remarks>リスト表示の１ページ内最大表示ファイル数</remarks>
	Public ReadOnly Property MaxCountInPage As Integer

		Get

			Return _Commands.PageSize

		End Get

	End Property

	''' <summary>処理タイププロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property ProcessMode As ProcessType

		Get

			Return _GetProcessMode(_Commands)

		End Get

	End Property

#End Region

#Region "コンストラクタ"

	''' <summary>コンストラクタ</summary>
	''' <remarks></remarks>
	Private Sub New()

		'コマンドライン引数分繰り返す
		For Each mCommand As String In My.Application.CommandLineArgs

			'ヘルプコマンドが存在するかをセット、ヘルプコマンドがあった場合はヘルプコマンド以外がセットされていても無視して処理を終了
			If _SetHelpCommand(mCommand, _Commands) Then Exit Sub

			'フォルダパスをセット、セット出来た時は次の繰り返しへ
			If _SetFolderPath(mCommand, _Commands) Then Continue For

			'フォームタイプコマンドをセット、セット出来た時は次の繰り返しへ
			If _SetFormTypeCommand(mCommand, _Commands) Then Continue For

			'出力形式コマンドをセット、セット出来た時は次の繰り返しへ
			If _SetOutputCommand(mCommand, _Commands) Then Continue For

			'拡張子コマンドをセット、セット出来た時は次の繰り返しへ
			If _SetExtensionCommand(mCommand, _Commands) Then Continue For

			'最大表示ファイル数コマンドをセット、セット出来た時は次の繰り返しへ
			If _SetPageSizeCommand(mCommand, _Commands) Then Continue For

		Next

		'拡張子コマンドの値が不正だった時、正しい値をセットする
		_Commands.Extension = _GetCorrectExtension(_Commands.FormType, _Commands.Extension)

	End Sub

#End Region

#Region "メソッド"

	''' <summary>フォルダパスをセット</summary>
	''' <param name="pCommandLineArg">対象コマンドライン引数</param>
	''' <param name="pCommands">コマンドライン格納変数</param>
	''' <returns>True ：フォルダパスをセット 
	'''          False：フォルダパスを未セット</returns>
	''' <remarks>対象のコマンドライン引数がフォルダパスの時、
	'''          コマンドライン格納変数にセットする         </remarks>
	Private Function _SetFolderPath(ByVal pCommandLineArg As String, ByRef pCommands As CommandList) As Boolean

		'入力パスが空の時、フォルダパスを未セット
		If String.IsNullOrEmpty(pCommandLineArg) Then Return False

		'入力パスがディレクトリで無かった時、フォルダパスを未セット
		If Not System.IO.Directory.Exists(pCommandLineArg) Then Return False

		'フォルダパスをセット
		pCommands.FolderPath = pCommandLineArg

		Return True

	End Function

	''' <summary>ヘルプコマンドが存在するかをセット</summary>
	''' <param name="pCommandLineArg">対象コマンドライン引数</param>
	''' <param name="pCommands">コマンドライン格納変数</param>
	''' <returns>True ：ヘルプコマンドが存在するかをセット
	'''          False：ヘルプコマンドが存在するかを未セット</returns>
	''' <remarks>対象のコマンドライン引数がヘルプコマンドの時、
	'''          コマンドライン格納変数にセットする           </remarks>
	Private Function _SetHelpCommand(ByVal pCommandLineArg As String, ByRef pCommands As CommandList) As Boolean

		'コマンドがヘルプだったらコマンドライン格納変数にセット
		If pCommandLineArg = cKeyWords.Help Then

			_Commands.HasHelpCommand = True
			Return True

		Else

			_Commands.HasHelpCommand = False
			Return False

		End If

	End Function

	''' <summary>フォームタイプコマンドをセット</summary>
	''' <param name="pCommandLineArg">対象コマンドライン引数</param>
	''' <param name="pCommands">コマンドライン格納変数</param>
	''' <returns>True ：フォームタイプコマンドをセット 
	'''          False：フォームタイプコマンドを未セット</returns>
	''' <remarks>対象のコマンドライン引数がフォームタイプコマンドの時、
	'''          コマンドライン格納変数にセットする                   </remarks>
	Private Function _SetFormTypeCommand(ByVal pCommandLineArg As String, ByRef pCommands As CommandList) As Boolean

		'コマンドライン引数から拡張子コマンドと拡張子を取得
		Dim mCommand As New Command(pCommandLineArg, cKeyWords.Form)

		'対象コマンドがフォームタイプコマンドの時
		If mCommand.KeyWord.ToLower = cKeyWords.Form.ToLower Then

			Select Case mCommand.Value.ToLower

				Case FormType.Text.ToString.ToLower

					pCommands.FormType = FormType.Text

				Case FormType.List.ToString.ToLower

					pCommands.FormType = FormType.List

				Case Else

					Return False

			End Select

			Return True

		Else

			Return False

		End If

	End Function

	''' <summary>出力形式コマンドをセット</summary>
	''' <param name="pCommandLineArg">対象コマンドライン引数</param>
	''' <param name="pCommands">コマンドライン格納変数</param>
	''' <returns>True ：出力形式コマンドをセット 
	'''          False：出力形式コマンドを未セット</returns>
	''' <remarks>対象のコマンドライン引数が出力形式コマンドの時、
	'''          コマンドライン格納変数にセットする                   </remarks>
	Private Function _SetOutputCommand(ByVal pCommandLineArg As String, ByRef pCommands As CommandList) As Boolean

		'コマンドライン引数から拡張子コマンドと拡張子を取得
		Dim mCommand As New Command(pCommandLineArg, cKeyWords.Output)

		'対象コマンドがフォームタイプコマンドの時
		If mCommand.KeyWord.ToLower = cKeyWords.Output.ToLower Then

			Select Case mCommand.Value.ToLower

				Case OutputType.ClipBoard.ToString.ToLower

					pCommands.Output = OutputType.ClipBoard

				Case OutputType.SaveDialog.ToString.ToLower

					pCommands.Output = OutputType.SaveDialog

				Case Else

					Return False

			End Select

			Return True

		Else

			Return False

		End If

	End Function

	''' <summary>拡張子コマンドをセット</summary>
	''' <param name="pCommandLineArg">対象コマンドライン引数</param>
	''' <param name="pCommands">コマンドライン格納変数</param>
	''' <returns>True ：拡張子コマンドをセット 
	'''          False：拡張子コマンドを未セット</returns>
	''' <remarks>対象のコマンドライン引数が出力形式コマンドの時、
	'''          コマンドライン格納変数にセットする                   </remarks>
	Private Function _SetExtensionCommand(ByVal pCommandLineArg As String, ByRef pCommands As CommandList) As Boolean

		'コマンドライン引数から拡張子コマンドと拡張子を取得
		Dim mCommand As New Command(pCommandLineArg, cKeyWords.Extension)

		'対象コマンドが拡張子コマンドの時
		If mCommand.KeyWord.ToLower = cKeyWords.Extension.ToLower Then

			Select Case mCommand.Value.ToLower

				Case OutputFileFormat.TEXT.ToString.ToLower

					pCommands.Extension = OutputFileFormat.TEXT

				Case OutputFileFormat.CSV.ToString.ToLower

					pCommands.Extension = OutputFileFormat.CSV

				Case OutputFileFormat.TSV.ToString.ToLower

					pCommands.Extension = OutputFileFormat.TSV

				Case OutputFileFormat.HTML.ToString.ToLower

					pCommands.Extension = OutputFileFormat.HTML

				Case Else

					Return False

			End Select

			Return True

		Else

			Return False

		End If

	End Function

	''' <summary>最大表示ファイル数コマンドをセット</summary>
	''' <param name="pCommandLineArg">対象コマンドライン引数</param>
	''' <param name="pCommands">コマンドライン格納変数</param>
	''' <returns>True ：最大表示ファイル数コマンドをセット 
	'''          False：最大表示ファイル数コマンドを未セット</returns>
	''' <remarks>対象のコマンドライン引数が最大表示ファイル数コマンドの時、
	'''          コマンドライン格納変数にセットする                   </remarks>
	Private Function _SetPageSizeCommand(ByVal pCommandLineArg As String, ByRef pCommands As CommandList) As Boolean

		'コマンドライン引数から拡張子コマンドと拡張子を取得
		Dim mCommand As New Command(pCommandLineArg, cKeyWords.PageSize)

		'対象コマンドがフォームタイプコマンドの時
		If mCommand.KeyWord.ToLower = cKeyWords.PageSize.ToLower Then

			Dim mSetPageSize As Integer

			'取得した最大表示ファイル数が数値に変換できる時
			If Integer.TryParse(mCommand.Value, mSetPageSize) Then

				'文字列を数値に変換したものをセット
				pCommands.PageSize = mSetPageSize
				Return True

			Else

				Return False

			End If

		Else

			Return False

		End If

	End Function

	''' <summary>処理タイプを取得</summary>
	''' <param name="pCommands">コマンドライン格納変数</param>
	''' <returns>処理タイプ</returns>
	''' <remarks></remarks>
	Private Function _GetProcessMode(ByVal pCommands As CommandList) As ProcessType

		'コマンドライン引数にヘルプコマンドがあった時はヘルプ処理を返す
		If pCommands.HasHelpCommand Then Return ProcessType.Help

		'フォルダパス項目が空文字の時
		If String.IsNullOrEmpty(pCommands.FolderPath) Then

			'通常処理を返す
			Return ProcessType.Normarl

		Else

			'コマンド処理を返す
			Return ProcessType.Command

		End If

	End Function

	''' <summary>フォームにあった正しい拡張子の値を取得</summary>
	''' <param name="pFormType">フォームタイプ</param>
	''' <param name="pExtension">拡張子コマンドの値</param>
	''' <returns>フォームにあった正しい拡張子の値</returns>
	''' <remarks></remarks>
	Private Function _GetCorrectExtension(ByVal pFormType As FormType, ByVal pExtension As OutputFileFormat) As OutputFileFormat

		Dim CorrectExtension As OutputFileFormat = pExtension

		'フォームタイプごと処理を分岐
		Select Case pFormType

			Case FormType.Text

				'拡張子コマンドが出力文字列フォームで使用出来る拡張子と一致しなかったら
				If Not cTextFormExtensionList.Contains(pExtension) Then

					'出力文字列フォームのデフォルトの拡張子をセット
					CorrectExtension = OutputFileFormat.TEXT

				End If

			Case FormType.List

				'拡張子コマンドがリスト表示フォームで使用出来る拡張子と一致しなかったら
				If Not cListFormExtensionList.Contains(pExtension) Then

					'リスト表示フォームのデフォルトの拡張子をセット
					CorrectExtension = OutputFileFormat.CSV

				End If

		End Select

		Return CorrectExtension

	End Function

#End Region

#Region "外部公開メソッド"

	''' <summary>コマンドリストを表示</summary>
	''' <remarks></remarks>
	Public Shared Sub ShowCommnadList()

		'コマンドリストの文字列を作成する
		Dim mCommnadListString As New System.Text.StringBuilder

		With mCommnadListString

			.AppendLine("ﾍﾙﾌﾟ　　　　　：/?                                           ")
			.AppendLine("ﾌｫｰﾑﾀｲﾌﾟ　　　：/Form=(Text | List)                          ")
			.AppendLine("　　　　　　　　※ﾃﾞﾌｫﾙﾄは「/Form=Text」                     ")
			.AppendLine("　　　　　　　　　Text：出力文字列ﾌｫｰﾑ                       ")
			.AppendLine("　　　　　　　　　List：ﾘｽﾄ表示ﾌｫｰﾑ                          ")
			.AppendLine("出力形式　　　：/Output=(ClipBoard | SaveDialog)             ")
			.AppendLine("　　　　　　　　　ClipBoard ：ｸﾘｯﾌﾟﾎﾞｰﾄﾞにｺﾋﾟｰ               ")
			.AppendLine("　　　　　　　　　SaveDialog：名前をつけて保存ﾀﾞｲｱﾛｸﾞで保存  ")
			.AppendLine("拡張子　　　　：/Extension=(txt | csv | tsv | html)          ")
			.AppendLine("　　　　　　　　　txt ：TXT形式で出力                        ")
			.AppendLine("　　　　　　　　　csv ：CSV形式で出力                        ")
			.AppendLine("　　　　　　　　　tsv ：TSV形式で出力                        ")
			.AppendLine("　　　　　　　　　html：HTML形式で出力                       ")
			.AppendLine("　　　　　　　　※ﾃﾞﾌｫﾙﾄの拡張子はﾌｫｰﾑによって異なります     ")
			.AppendLine("　　　　　　　　  「/Extension=txt」(/Form=Text)             ")
			.AppendLine("　　　　　　　　  「/Extension=csv」(/Form=List)             ")
			.AppendLine("　　　　　　　　※ﾌｫｰﾑによって使用できる拡張子が違います     ")
			.AppendLine("　　　　　　　　  txt,html(/Form=Text)                       ")
			.AppendLine("　　　　　　　　  csv,tsv,html(/Form=List)                   ")
			.AppendLine("　　　　　　　　※htmlは以下の設定では使用出来ません         ")
			.AppendLine("　　　　　　　　  「/Output=ClipBoard」                      ")
			.AppendLine("最大表示ﾌｧｲﾙ数：/PageSize=数値                               ")
			.AppendLine("　　　　　　　　※ﾃﾞﾌｫﾙﾄは「/PageSize=1000」                 ")
			.AppendLine("　　　　　　　　　数値のみ有効でそれ以外は無視されます       ")
			.AppendLine("　　　　　　　　　「/Form=List」の時しか効きません           ")
			.AppendLine("                                                             ")
			.AppendLine("★不正なｺﾏﾝﾄﾞﾗｲﾝ引数を使用した場合★　　　　　　　　         ")
			.AppendLine("　不正なｷｰﾜｰﾄﾞの場合は無視されます                           ")
			.AppendLine("　不正な値の場合はﾃﾞﾌｫﾙﾄ設定で実行されます                   ")
			.AppendLine("                                                             ")
			.AppendLine("★使用例★                                                   ")
			.AppendLine("　FolderFileList.exe ""ﾌｫﾙﾀﾞﾊﾟｽ"" /Form=List                 ")
			.AppendLine("　※ﾌｫﾙﾀﾞﾊﾟｽをﾘｽﾄ表示ﾌｫｰﾑで表示                              ")

		End With

		MessageBox.Show(mCommnadListString.ToString, "コマンドリスト", MessageBoxButtons.OK, MessageBoxIcon.Information)

		mCommnadListString = Nothing

	End Sub

	''' <summary>コマンドモードの文言をフォームタイトルにセットする</summary>
	''' <param name="pForm">対象フォーム</param>
	''' <remarks></remarks>
	Public Shared Sub SetCommandModeToTitle(ByVal pForm As Form)

		'処理タイプがコマンド処理の時
		If CommandLine.Instance.ProcessMode = ProcessType.Command Then

			'フォームのタイトルに（コマンドモード）を付加する
			pForm.Text = pForm.Text & "（コマンドモード）"

		End If

	End Sub

	''' <summary>アプリケーションに渡されたコマンドライン引数を表示（メッセージボックス）</summary>
	''' <remarks>※Conditional("DEBUG")によってリリース時はコンパイルされません</remarks>
	<Conditional("DEBUG")> _
	Public Shared Sub ShowCommandLineArgListToMessageBox()

		'コマンドリストの文字列を作成する
		Dim mCommnadLineArgListString As New System.Text.StringBuilder

		With mCommnadLineArgListString

			.AppendLine("①フォルダパス                         ")
			.AppendLine("→" & Instance.FolderPath & "          ")
			.AppendLine("                                       ")
			.AppendLine("②ヘルプ                               ")
			.AppendLine("→" & Instance.HasHelpCommand & "      ")
			.AppendLine("                                       ")
			.AppendLine("③フォームタイプ                       ")
			.AppendLine("→" & Instance.TargetForm.ToString & " ")
			.AppendLine("                                       ")
			.AppendLine("④出力形式                             ")
			.AppendLine("→" & Instance.Output.ToString & "     ")
			.AppendLine("                                       ")
			.AppendLine("⑤拡張子                               ")
			.AppendLine("→" & Instance.Extension.ToString & "  ")
			.AppendLine("                                       ")
			.AppendLine("⑥最大表示ファイル数                   ")
			.AppendLine("→" & Instance.MaxCountInPage & "      ")

		End With

		MessageBox.Show(mCommnadLineArgListString.ToString, "コマンドライン引数", MessageBoxButtons.OK, MessageBoxIcon.Information)

		mCommnadLineArgListString = Nothing

	End Sub

	''' <summary>アプリケーションに渡されたコマンドライン引数を表示（Messenger風通知メッセージ）</summary>
	''' <remarks>※Conditional("DEBUG")によってリリース時はコンパイルされません</remarks>
	<Conditional("DEBUG")> _
	Public Shared Sub ShowCommandLineArgListToPopupMessage()

		'コマンドリストの文字列を作成する
		Dim mCommnadLineArgListString As New System.Text.StringBuilder

		With mCommnadLineArgListString

			.AppendLine("①フォルダパス                         ")
			.AppendLine("→" & Instance.FolderPath & "          ")
			.AppendLine("②ヘルプ                               ")
			.AppendLine("→" & Instance.HasHelpCommand & "      ")
			.AppendLine("③フォームタイプ                       ")
			.AppendLine("→" & Instance.TargetForm.ToString & " ")
			.AppendLine("④出力形式                             ")
			.AppendLine("→" & Instance.Output.ToString & "     ")
			.AppendLine("⑤拡張子                               ")
			.AppendLine("→" & Instance.Extension.ToString & "  ")
			.AppendLine("⑥最大表示ファイル数                   ")
			.AppendLine("→" & Instance.MaxCountInPage & "      ")

		End With

		'Messenger風通知メッセージを表示
		Dim mPopupMessage As New frmPopupMessage("コマンドライン引数", mCommnadLineArgListString.ToString)
		mPopupMessage.Show()

		mCommnadLineArgListString = Nothing

	End Sub

#End Region

End Class