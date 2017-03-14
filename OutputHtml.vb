Option Explicit On

Imports System.Data
Imports System.Text
Imports FolderFileList.FolderFileList

''' <summary>Html出力の機能を提供する</summary>
''' <remarks></remarks>
Public Class OutputHtml

#Region "定数"

	''' <summary>階層構造表示用文字列</summary>
	''' <remarks>フォルダの階層構造を見た目でわかるようにファイル名
	'''          の前に全角スペースを２つ付加する                  </remarks>
	Private Const _cHierarchySpace As String = "　　"

	''' <summary>フォルダファイルリストのドキュメントURL</summary>
	''' <remarks></remarks>
	Public Const cFolderFileListDocumentURL As String = "https://github.com/dodonki1223/FolderFileList/tree/master/docs"

	''' <summary>Selectタグ初期値</summary>
	''' <remarks></remarks>
	Private Const _cInitialSelectValue As String = "未選択"

#End Region

#Region "列挙体"

	''' <summary>リンクタイプ</summary>
	''' <remarks></remarks>
	Public Enum LinkType

		''' <summary>ローカルファイルパス</summary>
		Local

		''' <summary>URL</summary>
		Url

	End Enum

#End Region

#Region "変数"

	''' <summary>フォルダ・ファイルリスト</summary>
	''' <remarks></remarks>
	Private _FolderFileList As FolderFileList

	Private _FormType As CommandLine.FormType

#End Region

#Region "プロパティ"

	''' <summary>フォルダファイルリストプロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property FolderFileList As FolderFileList

		Get

			Return _FolderFileList

		End Get

	End Property

	''' <summary>フォームタイププロパティ</summary>
	''' <remarks></remarks>
	Public Property FormType As CommandLine.FormType

		Get

			Return _FormType

		End Get

		Set(value As CommandLine.FormType)

			_FormType = value

		End Set

	End Property

	''' <summary>Html文書</summary>
	''' <remarks>FolderFileListからHtml文書に変換したものを取得</remarks>
	Public ReadOnly Property HtmlSentence() As String

		Get

			Return _ConvertFolderFileListToHtmlSentence(_FolderFileList)

		End Get

	End Property

#End Region

#Region "コンストラクタ"

	''' <summary>コンストラクタ</summary>
	''' <remarks>引数無しは外部に公開しない</remarks>
	Private Sub New()

	End Sub

	''' <summary>コンストラクタ</summary>
	''' <param name="pFolderFileList">フォルダファイルリスト</param>
	''' <param name="pFormType">フォームタイプ</param>
	''' <remarks>引数付きのコンストラクタのみを公開</remarks>
	Public Sub New(ByVal pFolderFileList As FolderFileList, Optional ByVal pFormType As CommandLine.FormType = CommandLine.FormType.Text)

		'フォームタイプをセット
		_FormType = pFormType

		'フォルダファイルリストをセット
		_FolderFileList = pFolderFileList

	End Sub

#End Region

#Region "メソッド"

	''' <summary>フォルダファイルリストをHtml文書に変換</summary>
	''' <param name="pFolderFileList">フォルダファイルリスト</param>
	''' <returns>フォルダファイルリストをHtml文書に変換した文字列</returns>
	''' <remarks></remarks>
	Private Function _ConvertFolderFileListToHtmlSentence(ByVal pFolderFileList As FolderFileList) As String

		Dim mHtmlSentence As New StringBuilder

		With mHtmlSentence

			.AppendLine("<!DOCTYPE html>                                                                                                                                ")
			.AppendLine("<html lang=""ja"">                                                                                                                             ")
			.AppendLine("  <head>                                                                                                                                       ")
			.AppendLine("    <meta charset=""UTF-8"">                                                                                                                   ")

			'CSS設定
			.Append(_GetCssSetting())

			.AppendLine("  </head>                                                                                                                                      ")
			.AppendLine("  <body>                                                                                                                                       ")
			.AppendLine("    <header id=""header"">                                                                                                                     ")
			.AppendLine("      <h1>" & _FolderFileList.TargetPathFolderName & "</h1>                                                                                    ")
			.AppendLine("    </header>                                                                                                                                  ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("    <div class=""refineConditions"">                                                                                                           ")
			.AppendLine("      <fieldset>                                                                                                                               ")
			.AppendLine("        <legend>絞り込み条件</legend>                                                                                                          ")
			.AppendLine("        <div class=""refineBox"">                                                                                                              ")
			.AppendLine("          <fieldset>                                                                                                                           ")
			.AppendLine("            <legend>対象ファイル</legend>                                                                                                      ")
			.AppendLine("            <input type=""radio"" name=""displayTarget"" value=""All"" checked>全て表示</input>                                                ")
			.AppendLine("            <input type=""radio"" name=""displayTarget"" value=""Folder"">フォルダのみ表示</input>                                             ")
			.AppendLine("            <input type=""radio"" name=""displayTarget"" value=""File"">ファイルのみ表示</input>                                               ")
			.AppendLine("          </fieldset>                                                                                                                          ")
			.AppendLine("        </div>                                                                                                                                 ")
			.AppendLine("        <div class=""refineBox"">                                                                                                              ")
			.AppendLine("          <fieldset>                                                                                                                           ")
			.AppendLine("            <legend>拡張子指定</legend>                                                                                                        ")
			.AppendLine("            <select name=""extensionTarget"" >                                                                                                 ")

			'拡張子リスト分繰り返す
			For Each mExtension As String In _FolderFileList.ExtensionList(_cInitialSelectValue)

				.AppendLine("              <option value=""" & mExtension & """" & _GetNotSelectClass(mExtension) & ">" & mExtension & "</option>                           ")

			Next

			.AppendLine("            </select>                                                                                                                          ")
			.AppendLine("          </fieldset>                                                                                                                          ")
			.AppendLine("        </div>                                                                                                                                 ")
			.AppendLine("        <div class=""refineBox"">                                                                                                              ")
			.AppendLine("          <fieldset>                                                                                                                           ")
			.AppendLine("            <legend>ファイルサイズ指定</legend>                                                                                                ")
			.AppendLine("            <select name=""sizeLevelTarget"" >                                                                                                 ")

			'ファイルサイズレベル分繰り返す
			For Each mLevel As DataRow In FolderFileList.cFileSizeLevel.LevelList(_cInitialSelectValue).Rows

				Dim mLevelValue As String = mLevel(cFileSizeLevel.LevelListColum.NAME)
				.AppendLine("              <option value=""" & mLevelValue & """" & _GetNotSelectClass(mLevelValue) & ">" & mLevelValue & "</option>                        ")

			Next

			.AppendLine("            </select>                                                                                                                          ")
			.AppendLine("          </fieldset>                                                                                                                          ")
			.AppendLine("        </div>                                                                                                                                 ")
			.AppendLine("        <div class=""refineBox"">                                                                                                              ")
			.AppendLine("          <fieldset>                                                                                                                           ")
			.AppendLine("            <legend>名前（LIKE検索）</legend>                                                                                                  ")
			.AppendLine("            <input type=""text"" name=""fileNameSearch"" placeholder=""検索したい文字列を入力して下さい""></input>                             ")
			.AppendLine("          </fieldset>                                                                                                                          ")
			.AppendLine("        </div>                                                                                                                                 ")
			.AppendLine("        <div class=""refineButtonBox"">                                                                                                        ")
			.AppendLine("          <a class=""refineButton"" href=""javascript:(function() {refineFolderFileList()})();"">絞り込み</a>                                  ")
			.AppendLine("        </div>                                                                                                                                 ")
			.AppendLine("      </fieldset>                                                                                                                              ")
			.AppendLine("    </div>                                                                                                                                     ")
			.AppendLine("    <div>                                                                                                                                      ")
			.AppendLine("      <p class=""folderFileCount""></p>                                                                                                        ")
			.AppendLine("    </div>                                                                                                                                     ")
			.AppendLine("    <div>                                                                                                                                      ")
			.AppendLine("      <table class=""folderFileList"">                                                                                                         ")
			.AppendLine("        <thead>                                                                                                                                ")
			.AppendLine("          <tr>                                                                                                                                 ")
			.AppendLine("            <th>" & StrConv(FolderFileListJapaneseColumn.ﾌｧｲﾙ名.ToString, VbStrConv.Wide) & "</th>                                             ")
			.AppendLine("            <th>" & FolderFileListJapaneseColumn.更新日時.ToString & "</th>                                                                    ")
			.AppendLine("            <th>" & StrConv(FolderFileListJapaneseColumn.ﾌｧｲﾙｻｲｽﾞ.ToString, VbStrConv.Wide) & "</th>                                           ")
			.AppendLine("            <th>" & StrConv(FolderFileListJapaneseColumn.ﾌｧｲﾙｼｽﾃﾑﾀｲﾌﾟ.ToString, VbStrConv.Wide) & "</th>                                       ")
			.AppendLine("            <th>" & FolderFileListJapaneseColumn.拡張子.ToString & "</th>                                                                      ")
			.AppendLine("            <th>" & StrConv(FolderFileListJapaneseColumn.ﾌｧｲﾙｻｲｽﾞﾚﾍﾞﾙ.ToString, VbStrConv.Wide) & "</th>                                       ")
			.AppendLine("            <th>" & StrConv(FolderFileListJapaneseColumn.親ﾌｫﾙﾀﾞ.ToString, VbStrConv.Wide) & "</th>                                            ")
			.AppendLine("            <th>" & StrConv(FolderFileListJapaneseColumn.対象ﾌｫﾙﾀﾞ以下.ToString, VbStrConv.Wide) & "</th>                                      ")
			.AppendLine("            <th>" & StrConv(FolderFileListJapaneseColumn.ﾌﾙﾊﾟｽ.ToString, VbStrConv.Wide) & "</th>                                              ")
			.AppendLine("          </tr>                                                                                                                                ")
			.AppendLine("        </thead>                                                                                                                               ")
			.AppendLine("        <tbody>                                                                                                                                ")

			'フォルダファイルリストの行数分繰り返す
			For Each mDr As DataRow In _FolderFileList.FolderFileList.Rows

				'行IDを作成
				Dim mRowID As String = "row" & mDr(FolderFileListColumn.No)

				'ファイル名を取得
				Dim mFileName As String = _GetFileName(mDr(FolderFileListColumn.Name), mDr(FolderFileListColumn.DirectoryLevel), _FormType)

				'ファイルパスのURLを作成
				Dim mFilePathUrl As String = _CreateWindowOpenURL(mDr(FolderFileListColumn.FullPath), LinkType.Local)

				'親フォルダパスのURLを作成
				Dim mParentFolderPathUrl As String = _CreateWindowOpenURL(mDr(FolderFileListColumn.ParentFolderFullPath), LinkType.Local)

				.AppendLine("          <tr id=""" & mRowID & """>                                                                                                           ")
				.AppendLine("            <td class=""fileName"">                                                                                                            ")
				.AppendLine("              <a href=""" & mFilePathUrl & """>" & mFileName & "</a>                                                                           ")
				.AppendLine("            </td>                                                                                                                              ")
				.AppendLine("            <td>" & mDr(FolderFileListColumn.UpdateDate) & "</td>                                                                              ")
				.AppendLine("            <td>" & mDr(FolderFileListColumn.SizeAndUnit) & "</td>                                                                             ")
				.AppendLine("            <td class=""fileSystemType"">" & mDr(FolderFileListColumn.FileSystemTypeName) & "</td>                                             ")
				.AppendLine("            <td class=""extension"">" & mDr(FolderFileListColumn.Extension) & "</td>                                                           ")
				.AppendLine("            <td class=""sizeLevel"">" & mDr(FolderFileListColumn.SizeLevelName) & "</td>                                                       ")
				.AppendLine("            <td>                                                                                                                               ")
				.AppendLine("              <a href=""" & mParentFolderPathUrl & """>" & mDr(FolderFileListColumn.ParentFolder) & "</a>                                      ")
				.AppendLine("            </td>                                                                                                                              ")
				.AppendLine("            <td>                                                                                                                               ")
				.AppendLine("              <a href=""" & mFilePathUrl & """>" & mDr(FolderFileListColumn.UnderTargetFolder) & "</a>                                         ")
				.AppendLine("            </td>                                                                                                                              ")
				.AppendLine("            <td>                                                                                                                               ")
				.AppendLine("              <a href=""" & mFilePathUrl & """>" & mDr(FolderFileListColumn.FullPath) & "</a>                                                  ")
				.AppendLine("            </td>                                                                                                                              ")
				.AppendLine("          </tr>                                                                                                                                ")

			Next

			.AppendLine("        </tbody>                                                                                                                               ")
			.AppendLine("      </table>                                                                                                                                 ")
			.AppendLine("    </div>                                                                                                                                     ")
			.AppendLine("    <footer id=""footer"">                                                                                                                     ")
			.AppendLine("      <p>                                                                                                                                      ")

			'フォルダ・ファイルリストのドキュメントURLを取得
			Dim mDocumentUrl As String = _CreateWindowOpenURL(cFolderFileListDocumentURL, LinkType.Url)
			.AppendLine("        <a href=""" & mDocumentUrl & """>FolderFileListのドキュメントはこちら</a>                                                              ")

			.AppendLine("      </p>                                                                                                                                     ")
			.AppendLine("    </footer>                                                                                                                                  ")

			'JavaScript設定
			.Append(_GetJavaScriptSetting())

			.AppendLine("  </body>                                                                                                                                      ")
			.AppendLine("</html>                                                                                                                                        ")

		End With

		Return mHtmlSentence.ToString

	End Function

	''' <summary>ファイル名を取得</summary>
	''' <param name="pFileName">ファイル名</param>
	''' <param name="pDirectoryLevel">ディレクトリレベル</param>
	''' <param name="pFormType">フォームタイプ</param>
	''' <returns>フォームタイプに応じてファイル名を返す
	'''          出力文字列：ディレクトリレベル分、階層構造用の文字列をファイル名の前に付加して返す
	'''          リスト表示：ファイル名をそのまま返す                                              </returns>
	''' <remarks></remarks>
	Private Function _GetFileName(ByVal pFileName As String, ByVal pDirectoryLevel As Integer, ByVal pFormType As CommandLine.FormType) As String

		Dim mFileName As String = String.Empty

		'フォームタイプごと処理を分岐
		Select Case pFormType

			Case CommandLine.FormType.Text

				mFileName = pFileName

				'ディレクトリレベル分繰り返す
				For i As Integer = 0 To pDirectoryLevel

					'階層構造用の文字列をファイル名の前に追加
					mFileName = _cHierarchySpace & mFileName

				Next

			Case CommandLine.FormType.List

				mFileName = pFileName

		End Select

		Return mFileName

	End Function

	''' <summary>Html出力時のCSS部分を取得</summary>
	''' <returns>Htmlに設定するCSS</returns>
	''' <remarks></remarks>
	Private Function _GetCssSetting() As String

		Dim mCssString As New System.Text.StringBuilder

		With mCssString

			.AppendLine("    <style type=""text/css"">                                                                                                                  ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* aタグ設定 */                                                                                                                          ")
			.AppendLine("      a {                                                                                                                                      ")
			.AppendLine("        text-decoration: none;                                                                                                                 ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* tableタグ設定 */                                                                                                                      ")
			.AppendLine("      table {                                                                                                                                  ")
			.AppendLine("          *border-collapse: collapse;                                                                                                          ")
			.AppendLine("          border-spacing: 0;                                                                                                                   ")
			.AppendLine("          width: 100%;                                                                                                                         ")
			.AppendLine("          font-size: 10px;                                                                                                                     ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* Selectタグ設定 */                                                                                                                     ")
			.AppendLine("      select {                                                                                                                                 ")
			.AppendLine("        width :120px;                                                                                                                          ")
			.AppendLine("        height:21px;                                                                                                                           ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* Selectタグの項目の文字色は黒色 */                                                                                                     ")
			.AppendLine("      option {                                                                                                                                 ")
			.AppendLine("        color:#333;                                                                                                                            ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* TextBox設定 */                                                                                                                        ")
			.AppendLine("      input[type=""text""] {                                                                                                                   ")
			.AppendLine("        width :300px;                                                                                                                          ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* ボタン設定 */                                                                                                                         ")
            .AppendLine("      .refineButton {                                                                                                                          ")
			.AppendLine("        margin-top:10px;                                                                                                                       ")
            .AppendLine("        position: relative;                                                                                                                    ")
			.AppendLine("        background-color: #1abc9c;                                                                                                             ")
			.AppendLine("        border-radius: 4px;                                                                                                                    ")
			.AppendLine("        color: #fff;                                                                                                                           ")
			.AppendLine("        line-height: 52px;                                                                                                                     ")
			.AppendLine("        -webkit-transition: none;                                                                                                              ")
			.AppendLine("        transition: none;                                                                                                                      ")
			.AppendLine("        box-shadow: 0 3px 0 #0e8c73;                                                                                                           ")
			.AppendLine("        text-shadow: 0 1px 1px rgba(0, 0, 0, .3);                                                                                              ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .refineButton:hover {                                                                                                                    ")
			.AppendLine("        background-color: #31c8aa;                                                                                                             ")
			.AppendLine("        box-shadow: 0 3px 0 #23a188;                                                                                                           ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .refineButton:active {                                                                                                                   ")
			.AppendLine("        top: 3px;                                                                                                                              ")
			.AppendLine("        box-shadow: none;                                                                                                                      ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* 絞り込み領域 */                                                                                                                       ")
			.AppendLine("      .refineConditions {                                                                                                                      ")
			.AppendLine("        width: 1270px;                                                                                                                         ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* 絞り込み条件ボックス */                                                                                                               ")
			.AppendLine("      .refineBox {                                                                                                                             ")
			.AppendLine("        float: left;                                                                                                                           ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* 絞り込み条件ボックス */                                                                                                               ")
            .AppendLine("      .refineButtonBox {                                                                                                                       ")
            .AppendLine("        text-align: right;                                                                                                                     ")
            .AppendLine("        float: left;                                                                                                                           ")
            .AppendLine("        width: 135px;                                                                                                                          ")
            .AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* 絞り込みボタン設定 */                                                                                                                 ")
			.AppendLine("      .refineConditions .refineButton {                                                                                                        ")
			.AppendLine("        display: inline-block;                                                                                                                 ")
			.AppendLine("        width: 130px;                                                                                                                          ")
			.AppendLine("        height: 50px;                                                                                                                          ")
			.AppendLine("        text-align: center;                                                                                                                    ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /* フォルダファイルリスト設定 */                                                                                                         ")
			.AppendLine("      .folderFileList {                                                                                                                        ")
			.AppendLine("          border: solid #ccc 1px;                                                                                                              ")
			.AppendLine("          -moz-border-radius: 6px;                                                                                                             ")
			.AppendLine("          -webkit-border-radius: 6px;                                                                                                          ")
			.AppendLine("          border-radius: 6px;                                                                                                                  ")
			.AppendLine("          -webkit-box-shadow: 0 1px 1px #ccc;                                                                                                  ")
			.AppendLine("          -moz-box-shadow: 0 1px 1px #ccc;                                                                                                     ")
			.AppendLine("          box-shadow: 0 1px 1px #ccc;                                                                                                          ")
			.AppendLine("          margin-top: 10px;                                                                                                                    ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList tr:hover {                                                                                                               ")
			.AppendLine("          background: #e8bebe;                                                                                                                 ")
			.AppendLine("          -o-transition: all 0.1s ease-in-out;                                                                                                 ")
			.AppendLine("          -webkit-transition: all 0.1s ease-in-out;                                                                                            ")
			.AppendLine("          -moz-transition: all 0.1s ease-in-out;                                                                                               ")
			.AppendLine("          -ms-transition: all 0.1s ease-in-out;                                                                                                ")
			.AppendLine("          transition: all 0.1s ease-in-out;                                                                                                    ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList td, .folderFileList th {                                                                                                 ")
			.AppendLine("          border-left: 1px solid #ccc;                                                                                                         ")
			.AppendLine("          border-top: 1px solid #ccc;                                                                                                          ")
			.AppendLine("          padding: 10px;                                                                                                                       ")
			.AppendLine("          text-align: left;                                                                                                                    ")
			.AppendLine("          white-space: nowrap;                                                                                                                 ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList th {                                                                                                                     ")
			.AppendLine("          background-color: #dce9f9;                                                                                                           ")
			.AppendLine("          background-image: -webkit-gradient(linear, left top, left bottom, from(#ebf3fc), to(#dce9f9));                                       ")
			.AppendLine("          background-image: -webkit-linear-gradient(top, #ebf3fc, #dce9f9);                                                                    ")
			.AppendLine("          background-image:    -moz-linear-gradient(top, #ebf3fc, #dce9f9);                                                                    ")
			.AppendLine("          background-image:     -ms-linear-gradient(top, #ebf3fc, #dce9f9);                                                                    ")
			.AppendLine("          background-image:      -o-linear-gradient(top, #ebf3fc, #dce9f9);                                                                    ")
			.AppendLine("          background-image:         linear-gradient(top, #ebf3fc, #dce9f9);                                                                    ")
			.AppendLine("          -webkit-box-shadow: 0 1px 0 rgba(255,255,255,.8) inset;                                                                              ")
			.AppendLine("          -moz-box-shadow:0 1px 0 rgba(255,255,255,.8) inset;                                                                                  ")
			.AppendLine("          box-shadow: 0 1px 0 rgba(255,255,255,.8) inset;                                                                                      ")
			.AppendLine("          border-top: none;                                                                                                                    ")
			.AppendLine("          text-shadow: 0 1px 0 rgba(255,255,255,.5);                                                                                           ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList td:first-child, .folderFileList th:first-child {                                                                         ")
			.AppendLine("          border-left: none;                                                                                                                   ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList th:first-child {                                                                                                         ")
			.AppendLine("          -moz-border-radius: 6px 0 0 0;                                                                                                       ")
			.AppendLine("          -webkit-border-radius: 6px 0 0 0;                                                                                                    ")
			.AppendLine("          border-radius: 6px 0 0 0;                                                                                                            ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList th:last-child {                                                                                                          ")
			.AppendLine("          -moz-border-radius: 0 6px 0 0;                                                                                                       ")
			.AppendLine("          -webkit-border-radius: 0 6px 0 0;                                                                                                    ")
			.AppendLine("          border-radius: 0 6px 0 0;                                                                                                            ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList th:only-child{                                                                                                           ")
			.AppendLine("          -moz-border-radius: 6px 6px 0 0;                                                                                                     ")
			.AppendLine("          -webkit-border-radius: 6px 6px 0 0;                                                                                                  ")
			.AppendLine("          border-radius: 6px 6px 0 0;                                                                                                          ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList tr:last-child td:first-child {                                                                                           ")
			.AppendLine("          -moz-border-radius: 0 0 0 6px;                                                                                                       ")
			.AppendLine("          -webkit-border-radius: 0 0 0 6px;                                                                                                    ")
			.AppendLine("          border-radius: 0 0 0 6px;                                                                                                            ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("      .folderFileList tr:last-child td:last-child {                                                                                            ")
			.AppendLine("          -moz-border-radius: 0 0 6px 0;                                                                                                       ")
			.AppendLine("          -webkit-border-radius: 0 0 6px 0;                                                                                                    ")
			.AppendLine("          border-radius: 0 0 6px 0;                                                                                                            ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("    </style>                                                                                                                                   ")

		End With

		Return mCssString.ToString

	End Function

	''' <summary>Html出力時のJavaScript部分を取得</summary>
	''' <returns>Htmlに設定するJavaScript</returns>
	''' <remarks></remarks>
	Private Function _GetJavaScriptSetting() As String

		Dim mJavaScriptString As New System.Text.StringBuilder

		With mJavaScriptString

			.AppendLine("    <script src=""https://code.jquery.com/jquery-2.2.0.min.js"" type=""text/javascript""></script>                                             ")
			.AppendLine("    <script type=""text/javascript"">                                                                                                          ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /**                                                                                                                                      ")
			.AppendLine("       * フォルダファイルリストの絞り込み                                                                                                      ")
			.AppendLine("       */                                                                                                                                      ")
			.AppendLine("      function refineFolderFileList(){                                                                                                         ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //選択されているラジオボタンを取得し対象ファイルシステムタイプの行を表示                                                               ")
			.AppendLine("        var checkedRadioButton = $('input[name=displayTarget]:checked').val();                                                                 ")
			.AppendLine("        refineByFileSystemType(checkedRadioButton);                                                                                            ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //選択されている拡張子を取得し対象拡張子のファイルに一致する行を表示                                                                   ")
			.AppendLine("        var selectExtension = $('[name=extensionTarget]').val();                                                                               ")
			.AppendLine("        refineBySelectTag(selectExtension, 'extension');                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //選択されているサイズレベルを取得し対象サイズレベルのファイルに一致する行を表示                                                       ")
			.AppendLine("        var selectSizeLevel = $('[name=sizeLevelTarget]').val();                                                                               ")
			.AppendLine("        refineBySelectTag(selectSizeLevel, 'sizeLevel');                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //入力されている「名前（LIKE検索）」の値を取得し対象文字列にLIKEで一致するファイル名の行を表示                                         ")
			.AppendLine("        var fileNameSearch = $('[name=fileNameSearch]').val();                                                                                 ")
			.AppendLine("        refineByTextBox(fileNameSearch, 'fileName');                                                                                           ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //フォルダファイルリストの件数を画面に表示                                                                                             ")
			.AppendLine("        setFolderFileCount();                                                                                                                  ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /**                                                                                                                                      ")
			.AppendLine("       * 対象のファイルシステムタイプの行を表示                                                                                                ")
			.AppendLine("       * @param  {string} target 表示対象のファイルシステムタイプ                                                                              ")
			.AppendLine("       *     ※All：すべて表示、Folder：フォルダのみ表示、File：ファイルのみ表示                                                               ")
			.AppendLine("       */                                                                                                                                      ")
			.AppendLine("      function refineByFileSystemType(target) {                                                                                                ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //全て表示定数                                                                                                                         ")
			.AppendLine("        const ALL = 'All';                                                                                                                     ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //Class名に「fileSystemType」と名の付くタグ分繰り返す                                                                                  ")
			.AppendLine("        $('.fileSystemType').each(function() {                                                                                                 ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //表示対象ファイルシステムタイプがALLの時                                                                                            ")
			.AppendLine("          if (target === ALL) {                                                                                                                ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //親要素を表示し次の繰り返しへ                                                                                                     ")
			.AppendLine("            $(this).parent().show();                                                                                                           ")
			.AppendLine("            return true;                                                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          }                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //「fileSystemType」と名の付くタグのテキストが「target」と一致したら                                                                 ")
			.AppendLine("          if ($(this).text() === target) {                                                                                                     ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「fileSystemType」と名の付くタグの親要素を表示する                                                                               ")
			.AppendLine("            $(this).parent().show();                                                                                                           ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          } else {                                                                                                                             ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「fileSystemType」と名の付くタグの親要素を非表示にする                                                                           ")
			.AppendLine("            $(this).parent().hide();                                                                                                           ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          }                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        });                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /**                                                                                                                                      ")
			.AppendLine("       *  選択されているSelectタグの値の行を表示                                                                                               ")
			.AppendLine("       * @param  {string} target    Selectタグで選択されている値                                                                               ")
			.AppendLine("       * @param  {string} className 対象のtdタグのクラス名                                                                                     ")
			.AppendLine("       */                                                                                                                                      ")
			.AppendLine("      function refineBySelectTag(target, className) {                                                                                          ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //対象のSelectタグの選択されている値が未選択の時は処理を終了                                                                           ")
			.AppendLine("        if (target === '未選択') return;                                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //Class名に「className」と名の付くタグ分繰り返す                                                                                       ")
			.AppendLine("        $('.' + className).each(function() {                                                                                                   ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //「className」と名の付くタグのテキストと親要素のdisplayの状態を取得                                                                 ")
			.AppendLine("          var targetTag = $(this).text().toLowerCase() ,                                                                                       ")
			.AppendLine("              parentDisplayState = $(this).parent().css('display');                                                                            ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //     「className」と名の付くタグのテキストが「target」と一致                                                                       ")
			.AppendLine("          // かつ「className」と名の付くタグの親要素のdisplayの状態が「none」でない時                                                          ")
			.AppendLine("          if (targetTag === target && parentDisplayState != 'none') {                                                                          ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("                //「className」と名の付くタグの親要素を表示                                                                                    ")
			.AppendLine("                $(this).parent().show();                                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          } else {                                                                                                                             ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「className」と名の付くタグの親要素を非表示                                                                                      ")
			.AppendLine("            $(this).parent().hide();                                                                                                           ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          }                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        });                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /**                                                                                                                                      ")
			.AppendLine("       *  テキストボックスに入力されている値にLIKE一致する行の表示                                                                             ")
			.AppendLine("       * @param  {string} target    input(text)タグで入力されている値                                                                          ")
			.AppendLine("       * @param  {string} className 対象のtdタグのクラス名                                                                                     ")
			.AppendLine("       */                                                                                                                                      ")
			.AppendLine("      function refineByTextBox(target, className) {                                                                                            ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //Class名に「className」と名の付くタグ分繰り返す                                                                                       ")
			.AppendLine("        $('.' + className).each(function() {                                                                                                   ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //「className」と名の付くタグのテキスト（小文字変換後）と親要素のdisplayの状態を取得                                                 ")
			.AppendLine("          //「 対象文字が含まれているか」変数にFalseをセット（デフォルト）                                                                     ")
			.AppendLine("          var targetString = $(this).text().toLowerCase() ,                                                                                    ")
			.AppendLine("              parentDisplayState = $(this).parent().css('display') ,                                                                           ")
			.AppendLine("              isIncludedString = false;                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //「className」と名の付くタグのテキスト（小文字変換後）に対象となる文字列が存在する時は                                              ")
			.AppendLine("          //「 対象文字が含まれているか」変数にTrueをセット                                                                                    ")
			.AppendLine("          if ( targetString.indexOf(target.toLowerCase()) != -1) isIncludedString = true;                                                      ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //     「 対象文字が含まれているか」変数がTrueで                                                                                     ")
			.AppendLine("          // かつ「className」と名の付くタグの親要素のdisplayの状態が「none」でない時                                                          ")
			.AppendLine("          if (isIncludedString && parentDisplayState != 'none') {                                                                              ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("                //「className」と名の付くタグの親要素を表示                                                                                    ")
			.AppendLine("                $(this).parent().show();                                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          } else {                                                                                                                             ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「className」と名の付くタグの親要素を非表示                                                                                      ")
			.AppendLine("            $(this).parent().hide();                                                                                                           ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          }                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        });                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /**                                                                                                                                      ")
			.AppendLine("       * Selectタグで最初に表示する項目の文字色を灰色表示にします                                                                              ")
			.AppendLine("       */                                                                                                                                      ")
			.AppendLine("      function setSelectTagColor() {                                                                                                           ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //Selectタグ内のOptionタグでSlectedされているタグ分繰り返す                                                                            ")
			.AppendLine("        $('select').find('option:selected').each(function() {                                                                                  ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //Slectedされているタグで「not-select」というクラス名を持っていた時                                                                  ")
			.AppendLine("          if ($(this).hasClass('not-select')) {                                                                                                ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            // 文字色を灰色に変更                                                                                                              ")
			.AppendLine("            $(this).parent().css({'color':'#A9A9A9'});                                                                                         ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          } else {                                                                                                                             ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            // 文字色を黒色に変更                                                                                                              ")
			.AppendLine("            $(this).parent().css({'color':'#333'});                                                                                            ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          }                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        });                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //項目が変更された時、条件によって色変更                                                                                               ")
			.AppendLine("        $('select').on('change', function(){                                                                                                   ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //Selectタグ内のOptionタグでSlectedされているタグ分繰り返す                                                                          ")
			.AppendLine("          $('select').find('option:selected').each(function() {                                                                                ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //Slectedされているタグで「not-select」というクラス名を持っていた時                                                                ")
			.AppendLine("            if ($(this).hasClass('not-select')) {                                                                                              ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("              // 文字色を灰色に変更                                                                                                            ")
			.AppendLine("              $(this).parent().css({'color':'#A9A9A9'});                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            } else {                                                                                                                           ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("              // 文字色を黒色に変更                                                                                                            ")
			.AppendLine("              $(this).parent().css({'color':'#333'});                                                                                          ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            }                                                                                                                                  ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          });                                                                                                                                  ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        });                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /**                                                                                                                                      ")
			.AppendLine("       * 拡張子使用不可設定                                                                                                                    ")
			.AppendLine("       * ※対象ファイルのラジオボタンイベントに拡張子使用不可の処理を追加する                                                                  ")
			.AppendLine("       */                                                                                                                                      ")
			.AppendLine("      function setExtensionDisable() {                                                                                                         ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        $('input[name=displayTarget]:radio').change(function() {                                                                               ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //対象ファイルFolder定数                                                                                                             ")
			.AppendLine("          const FOLDER = 'Folder';                                                                                                             ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //選択されているラジオボタンを取得                                                                                                   ")
			.AppendLine("          var checkedRadioButton = $('input[name=displayTarget]:checked').val();                                                               ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //選択されているラジオボタンが「Folder」の時                                                                                         ")
			.AppendLine("          if (checkedRadioButton === FOLDER) {                                                                                                 ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「extensionTarget」Selectタグの値を「未選択」でセット                                                                            ")
			.AppendLine("            $('select[name=extensionTarget]').val('未選択');                                                                                   ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「extensionTarget」Selectタグの使用を不可                                                                                        ")
			.AppendLine("            $('select[name=extensionTarget]').attr('disabled', true);                                                                          ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「extensionTarget」Selectタグの背景色を灰色に変更                                                                                ")
			.AppendLine("            $('select[name=extensionTarget]').css({'background-color':'#808080'});                                                             ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //Selectタグで最初に表示する項目の文字色を灰色表示                                                                                 ")
			.AppendLine("            setSelectTagColor();                                                                                                               ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          } else {                                                                                                                             ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「extensionTarget」Selectタグの「disabled」属性を削除                                                                            ")
			.AppendLine("            $('select[name=extensionTarget]').removeAttr('disabled');                                                                          ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("            //「extensionTarget」Selectタグの背景色を白に変更                                                                                  ")
			.AppendLine("            $('select[name=extensionTarget]').css({'background-color':'#ffffff'});                                                             ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          }                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        });                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /**                                                                                                                                      ")
			.AppendLine("       * フォルダファイルリストの件数を取得する                                                                                                ")
			.AppendLine("       * @return {number}  フォルダファイルリストの表示されている件数                                                                          ")
			.AppendLine("       */                                                                                                                                      ")
			.AppendLine("      function getFolderFileListRowCount(){                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //行数を数える変数                                                                                                                     ")
			.AppendLine("        var rowCount = 0;                                                                                                                      ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //クラス名が「fileName」と名の付くタグ分繰り返す                                                                                       ")
			.AppendLine("        $('.fileName').each(function() {                                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //「fileName」と名の付くタグの親要素のdisplayの状態を取得                                                                            ")
			.AppendLine("          var parentDisplayState = $(this).parent().css('display');                                                                            ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("          //親要素のdisplayの状態が「display:none;」で無かった時、行数を＋１                                                                   ")
			.AppendLine("          if (parentDisplayState != 'none') rowCount++;                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        });                                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        return rowCount;                                                                                                                       ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      /**                                                                                                                                      ")
			.AppendLine("       * フォルダ・ファイルリストの件数を画面にセット                                                                                          ")
			.AppendLine("       */                                                                                                                                      ")
			.AppendLine("      function setFolderFileCount() {                                                                                                          ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("        //「folderFileCount」クラスの名の付く要素のテキストに「フォルダ・ファイルリスト件数：○件」形式で表示させる                            ")
			.AppendLine("        $('.folderFileCount').text('フォルダ・ファイルリスト件数：' + getFolderFileListRowCount() + '件');                                     ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      }                                                                                                                                        ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      //Selectタグの最初に表示する項目の文字色を灰色表示に変更                                                                                 ")
			.AppendLine("      setSelectTagColor();                                                                                                                     ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      //対象ファイルのラジオボタンイベントに拡張子使用不可の処理を追加する                                                                     ")
			.AppendLine("      setExtensionDisable();                                                                                                                   ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("      //フォルダ・ファイルリストの件数を画面に表示                                                                                             ")
			.AppendLine("      setFolderFileCount();                                                                                                                    ")
			.AppendLine("                                                                                                                                               ")
			.AppendLine("    </script>                                                                                                                                  ")

		End With

		Return mJavaScriptString.ToString

	End Function

	''' <summary>Select句の未選択状態のクラスを取得する</summary>
	''' <param name="pValue">Option句に表示する値</param>
	''' <returns>not-selectのクラス設定</returns>
	''' <remarks></remarks>
	Private Function _GetNotSelectClass(ByVal pValue As String) As String

		'Option句に表示する値がSelectタグの初期値
		If pValue = _cInitialSelectValue Then

			'未選択クラス設定を返す
			Return " class=""not-select"""

		Else

			Return String.Empty

		End If

	End Function

	''' <summary>window.openで開く用のURLを作成</summary>
	''' <param name="pPath">対象パス</param>
	''' <returns>window.openで開く用のURL文字列</returns>
	''' <remarks>※別タブでリンクを開く時のパスは以下のどちらかの形式でなければならない
	'''              ①\\\\FileServer\\Folder1\\Gorira.txt
	'''              ②file://FileServer/Folder1/Gorira.txt                            </remarks>
	Private Function _CreateWindowOpenURL(ByVal pPath As String, ByVal pLinkType As LinkType) As String

		Dim mLink As String = String.Empty

		Select Case pLinkType

			Case LinkType.Local

				'パス文字列から「\」マークを「\\」に変換し、
				'パス文字列の前と後にjavascriptのwindow.open用の文字列を付加して返す
				mLink = "javascript:window.open('" & "file:\\\\" & Replace(pPath, "\", "\\") & "');"

			Case LinkType.Url

				'URLをwindow.open用の文字列を付加して返す
				mLink = "javascript:window.open('" & pPath & "');"

		End Select

		Return mLink

	End Function

#End Region

End Class