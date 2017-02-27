Option Explicit On

Imports System.IO
Imports System.Data.Linq

''' <summary>フォルダファイルリストの機能を提供する</summary>
''' <remarks></remarks>
Public Class FolderFileList

#Region "定数"

	''' <summary>出力用の文字列を提供する</summary>
	''' <remarks></remarks>
	Private Class _cOutputString

		''' <summary>ディレクトリ内の最後のファイル</summary>
		Public Const LastFile As String = " └ "

		''' <summary>ディレクトリ内のファイル</summary>
		Public Const File As String = " ├ "

		''' <summary>ディレクトリ内にファイルがまだ存在する</summary>
		Public Const ExistsFile As String = " ｜ "

		''' <summary>ディレクトリ内にファイルが存在しない</summary>
		Public Const NotExistsFile As String = "    "

	End Class

	''' <summary>ファイルサイズレベルを提供する</summary>
	''' <remarks></remarks>
	Public Class cFileSizeLevel

		''' <summary>10GB</summary>
		''' <remarks></remarks>
		Private Const _10GB As Long = 10737418240

		''' <summary>1GB</summary>
		''' <remarks></remarks>
		Private Const _1GB As Long = 1073741824

		''' <summary>100MB</summary>
		''' <remarks></remarks>
		Private Const _100MB As Long = 104857600

		''' <summary>1MB</summary>
		''' <remarks></remarks>
		Private Const _1MB As Long = 1048576

		''' <summary>1KB</summary>
		''' <remarks></remarks>
		Private Const _1KB As Long = 1024

		''' <summary>ファイルサイズレベルリストのからデータID</summary>
		''' <remarks></remarks>
		Public Const LevelListNoneDataID As Integer = 0

		''' <summary>ファイルサイズのレベル</summary>
		''' <remarks></remarks>
		Public Enum SizeLevel

			''' <summary>レベル０</summary>
			''' <remarks>０B以上１KBより小さい</remarks>
			Level0 = 1

			''' <summary>レベル１</summary>
			''' <remarks>１KB以上１MBより小さい</remarks>
			Level1

			''' <summary>レベル２</summary>
			''' <remarks>１MB以上１００MBより小さい</remarks>
			Level2

			''' <summary>レベル３</summary>
			''' <remarks>１００MB以上１GBより小さい</remarks>
			Level3

			''' <summary>レベル４</summary>
			''' <remarks>１GB以上１０GBより小さい</remarks>
			Level4

			''' <summary>レベル５</summary>
			''' <remarks>１０GB以上</remarks>
			Level5

		End Enum

		''' <summary>ファイルサイズのレベルリストカラム</summary>
		''' <remarks>コンボボックスの表示用のテキストと内部で持つIDの設定用</remarks>
		Public Enum LevelListColum

			''' <summary>ID</summary>
			''' <remarks></remarks>
			ID

			''' <summary>名前</summary>
			''' <remarks></remarks>
			NAME

		End Enum

		''' <summary>ファイルサイズレベルを取得</summary>
		''' <param name="pFileSize">ファイルサイズ（Byte）</param>
		''' <returns>ファイルサイズのレベル</returns>
		''' <remarks></remarks>
		Public Shared Function GetLevel(ByVal pFileSize As Long) As SizeLevel

			Select Case True

				Case pFileSize >= _10GB

					Return SizeLevel.Level5

				Case pFileSize >= _1GB

					Return SizeLevel.Level4

				Case pFileSize >= _100MB

					Return SizeLevel.Level3

				Case pFileSize >= _1MB

					Return SizeLevel.Level2

				Case pFileSize >= _1KB

					Return SizeLevel.Level1

				Case Else

					Return SizeLevel.Level0

			End Select

		End Function

		''' <summary>ファイルサイズレベルのカラーを取得</summary>
		''' <param name="pLevel">ファイルサイズレベル</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function LevelColor(ByVal pLevel As SizeLevel) As System.Drawing.Color

			Select Case pLevel

				Case SizeLevel.Level0

					Return Nothing

				Case SizeLevel.Level1

					Return System.Drawing.ColorTranslator.FromHtml("#b7ffff")

				Case SizeLevel.Level2

					Return System.Drawing.ColorTranslator.FromHtml("#b7ffb7")

				Case SizeLevel.Level3

					Return System.Drawing.ColorTranslator.FromHtml("#ffffb7")

				Case SizeLevel.Level4

					Return System.Drawing.ColorTranslator.FromHtml("#b7b7ff")

				Case SizeLevel.Level5

					Return System.Drawing.ColorTranslator.FromHtml("#ffb7b7")

			End Select

		End Function

		''' <summary>ファイルサイズレベルリストを取得</summary>
		''' <returns>ファイルサイズレベルのリスト</returns>
		''' <remarks></remarks>
		Public Shared Function LevelList() As DataTable

			Dim mLevelList As New DataTable()

			'カラムを作成
			mLevelList.Columns.Add(LevelListColum.ID.ToString, GetType(Integer))
			mLevelList.Columns.Add(LevelListColum.NAME.ToString, GetType(String))

			'空データを追加
			Dim mNoneRow As DataRow = mLevelList.NewRow

			mNoneRow(LevelListColum.ID) = 0
			mNoneRow(LevelListColum.NAME) = String.Empty

			mLevelList.Rows.Add(mNoneRow)

			'ファイルサイズレベル列挙体分繰り返す
			For Each mLevel As SizeLevel In System.Enum.GetValues(GetType(SizeLevel))

				'ファイルサイズレベル列挙体データを追加
				Dim mNewRow As DataRow = mLevelList.NewRow

				mNewRow(LevelListColum.ID) = mLevel
				mNewRow(LevelListColum.NAME) = mLevel.ToString

				mLevelList.Rows.Add(mNewRow)

			Next

			Return mLevelList

		End Function

	End Class

	''' <summary>エンコーディングを提供する</summary>
	''' <remarks></remarks>
	Public Class cEncording

		''' <summary>Shift-Jis</summary>
		Public Shared ReadOnly ShiftJis As System.Text.Encoding = System.Text.Encoding.GetEncoding(932)

		''' <summary>EUC-JP</summary>
		Public Shared ReadOnly EucJp As System.Text.Encoding = System.Text.Encoding.GetEncoding(20932)

		''' <summary>iso-2022-jp</summary>
		Public Shared ReadOnly JisCode As System.Text.Encoding = System.Text.Encoding.GetEncoding(50220)

		''' <summary>UTF-8</summary>
		Public Shared ReadOnly UTF8 As System.Text.Encoding = System.Text.Encoding.UTF8

	End Class

#End Region

#Region "列挙体"

	''' <summary>ファイルシステムタイプ</summary>
	''' <remarks></remarks>
	Public Enum FileSystemType

		''' <summary>フォルダ</summary>
		Folder = 1

		''' <summary>ファイル</summary>
		File

		''' <summary>不明</summary>
		None

	End Enum

	''' <summary>フォルダファイルリストデータカラム</summary>
	''' <remarks></remarks>
	Public Enum FolderFileListColumn

		''' <summary>データID</summary>
		No

		''' <summary>ファイル名</summary>
		Name

		''' <summary>更新日時</summary>
		UpdateDate

		''' <summary>ファイルシステムタイプ</summary>
		''' <remarks>フォルダとかファイル</remarks>
		FileSystemType

		''' <summary>ファイルシステムタイプ名</summary>
		''' <remarks>FileとかFolder</remarks>
		FileSystemTypeName

		''' <summary>拡張子</summary>
		Extension

		''' <summary>ファイルサイズ</summary>
		Size

		''' <summary>単位付きファイルサイズ</summary>
		''' <remarks>単位付きのサイズ表示</remarks>
		SizeAndUnit

		''' <summary>ファイルサイズレベル</summary>
		SizeLevel

		''' <summary>ファイルサイズレベル名</summary>
		SizeLevelName

		''' <summary>ディレクトリのレベル</summary>
		DirectoryLevel

		''' <summary>親フォルダ</summary>
		ParentFolder

		''' <summary>親フォルダのフルパス</summary>
		ParentFolderFullPath

		''' <summary>対象フォルダ以下</summary>
		UnderTargetFolder

		''' <summary>ファイルのフルパス</summary>
		FullPath

		''' <summary>フォルダ内で最後のファイルか</summary>
		IsLastFileInFolder

		''' <summary>表示文字列</summary>
		''' <remarks>出力するときの文字列</remarks>
		DispString

	End Enum

	''' <summary>フォルダファイルリスト並び順</summary>
	''' <remarks></remarks>
	Public Enum FolderFileListSortOrder

		''' <summary>昇順</summary>
		''' <remarks></remarks>
		ASC

		''' <summary>降順</summary>
		''' <remarks></remarks>
		DESC

	End Enum

#End Region

#Region "変数"

	''' <summary>フォルダ・ファイルリスト</summary>
	''' <remarks></remarks>
	Private _FolderFileList As DataTable

    ''' <summary>対象となるフォルダパス</summary>
    ''' <remarks></remarks>
    Private _TargetPath As String

    ''' <summary>フォルダ・ファイルリストのファイル数</summary>
    ''' <remarks>フォルダとファイルのファイル数を保持する</remarks>
    Private _FolderFileListCount As Integer

    ''' <summary>フォルダファイルリストの進捗状況を報告する変数</summary>
    ''' <remarks>フォルダファイルリスト作成中の進捗状況報告用</remarks>
    Private _ProcessProgresss As IProgress(Of FolderFileListProgress) = Nothing

#End Region

#Region "構造体"

    ''' <summary>フォルダファイルリスト進捗率構造体</summary>
    ''' <remarks>フォルダファイルリストの進捗率の内容を保持する構造体</remarks>
    Public Structure FolderFileListProgress

        ''' <summary>進捗率</summary>
        ''' <remarks></remarks>
        Private _Percent As Integer

        ''' <summary>処理対象フォルダファイル</summary>
        ''' <remarks></remarks>
        Private _ProcessingFolderFile As String

        ''' <summary>進捗率プロパティ</summary>
        ''' <remarks></remarks>
        Public Property Percent As Integer

            Set(value As Integer)

                _Percent = value

            End Set
            Get

                Return _Percent

            End Get

        End Property

        ''' <summary>対象フォルダファイルプロパティ</summary>
        ''' <remarks></remarks>
        Public Property ProcessingFolderFile As String

            Set(value As String)

                _ProcessingFolderFile = value

            End Set
            Get

                Return _ProcessingFolderFile

            End Get

        End Property

        ''' <summary>コンストラクタ</summary>
        ''' <param name="pPercent">進捗率</param>
        ''' <param name="pProcessingFolderFile">処理フォルダファイル</param>
        Public Sub New(ByVal pPercent As Integer, ByVal pProcessingFolderFile As String)

            '進捗率をセット
            _Percent = pPercent

            '処理フォルダファイル
            _ProcessingFolderFile = pProcessingFolderFile

        End Sub

    End Structure

#End Region

#Region "プロパティ"

    ''' <summary>フォルダファイルリストプロパティ</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property FolderFileList As DataTable

		Get

			Return _FolderFileList

		End Get

	End Property

	''' <summary>対象となるフォルダパスプロパティ</summary>
	''' <remarks>フルパスで返す</remarks>
	Public ReadOnly Property TargetPath As String

		Get

			Return _TargetPath

		End Get

	End Property

	''' <summary>対象となるフォルダ名プロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property TargetPathFolderName As String

		Get

			Dim mFolder As New System.IO.FileInfo(_TargetPath)
			Return mFolder.Name

		End Get

	End Property

	''' <summary>フォルダリストプロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property FolderFileListOnlyFolder As DataTable

		Get

			'フォルダファイルリストからフォルダのみを絞り込んで取得するクエリーを作成
			Dim mQuery = From FolderFileList In _FolderFileList.AsEnumerable
						 Where FolderFileList(FolderFileListColumn.FileSystemType.ToString) = FileSystemType.Folder
						 Order By FolderFileList(FolderFileListColumn.No) Ascending
						 Select FolderFileList

			'クエリーからDataViewを作成
			Dim mQueryToDataView As DataView = mQuery.AsDataView

			Return mQueryToDataView.ToTable

		End Get

	End Property

	''' <summary>ファイルリストプロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property FolderFileListOnlyFile As DataTable

		Get

			'フォルダファイルリストからファイルのみを絞り込んで取得するクエリーを作成
			Dim mQuery = From FolderFileList In _FolderFileList.AsEnumerable
						 Where FolderFileList(FolderFileListColumn.FileSystemType.ToString) = FileSystemType.File
						 Order By FolderFileList(FolderFileListColumn.No) Ascending
						 Select FolderFileList

			'クエリーからDataViewを作成
			Dim mQueryToDataView As DataView = mQuery.AsDataView

			Return mQueryToDataView.ToTable

		End Get

	End Property

	''' <summary>出力用文字列プロパティ</summary>
	''' <remarks></remarks>
	Public ReadOnly Property OutPutText As String

		Get

			'すべての行を取得し出力用の文字列を取得するQueryを作成（ラムダ式）
			Dim mQuery = _FolderFileList.AsEnumerable() _
						.Select(Function(FolderFileListRow As DataRow) New With _
										   { _
											   .DispString = FolderFileListRow.Field(Of String)(FolderFileListColumn.DispString) _
										   } _
							   )

			'出力用文字列を結合していく
			Dim mOutPutText As New System.Text.StringBuilder
			For Each FolderFileListRow In mQuery

				mOutPutText.AppendLine(FolderFileListRow.DispString)

			Next

			Return mOutPutText.ToString

		End Get

	End Property

    ''' <summary>拡張子リストプロパティ</summary>
    ''' <remarks>対象となるフォルダパスで取得されたファイル群から
    '''          作成した拡張子リストを返す                      </remarks>
    Public ReadOnly Property ExtensionList As ArrayList

        Get

            Return _GetExtensionList(_FolderFileList)

        End Get

    End Property

    ''' <summary>処理進捗プロパティ</summary>
	''' <remarks></remarks>
    Public WriteOnly Property ProcessProgress As IProgress(Of FolderFileListProgress)

        Set(value As IProgress(Of FolderFileListProgress))

			_ProcessProgresss = value

        End Set

	End Property

	''' <summary>対象フォルダ内のサブディレクトリとファイル数プロパティ</summary>
	''' <remarks>処理進捗を表示する場合はこのプロパティに値をセットすること</remarks>
	Public WriteOnly Property FileCountForTargetFolder As Integer

		Set(value As Integer)

			'対象フォルダ内のサブディレクトリとファイル数をセット
			_FolderFileListCount = GetFileCountForTargetFolder()

		End Set

	End Property

#End Region

#Region "コンストラクタ"

    ''' <summary>コンストラクタ</summary>
    ''' <remarks>引数無しは外部に公開しない</remarks>
    Private Sub New()

	End Sub

	''' <summary>コンストラクタ</summary>
	''' <param name="pPath">フォルダファイルリストを取得したいフォルダパス</param>
	''' <remarks>引数付きのコンストラクタのみを公開</remarks>
	Public Sub New(ByVal pPath As String)

        '対象パスをセット
        _TargetPath = pPath

        'フォルダファイルリストのDataTableのカラムを作成
        _FolderFileList = _CreateFolderFileListColumns()

    End Sub

#End Region

#Region "メソッド"

#Region "フォルダファイルリストDataTableのカラムを作成する"

	''' <summary>フォルダファイルリストDataTableのカラムを作成する</summary>
	''' <returns>フォルダファイルリストのカラム作成後のDataTable</returns>
	''' <remarks>FolderFileListColumn列挙体の項目分カラムを作成する</remarks>
	Private Function _CreateFolderFileListColumns() As DataTable

		Dim mFolderFileList As New DataTable

		'フォルダーファイルリストカラム列挙体分繰り返す
		For Each mColumnName As String In System.Enum.GetNames(GetType(FolderFileListColumn))

			'カラムによりデータ型を設定する
			Select Case mColumnName

				Case FolderFileListColumn.No.ToString, FolderFileListColumn.FileSystemType.ToString, FolderFileListColumn.DirectoryLevel.ToString, FolderFileListColumn.SizeLevel.ToString

					'「NO」、「ファイルシステムタイプ」、「ディレクトリレベル」、「ファイルサイズレベル」はInteger型で作成
					mFolderFileList.Columns.Add(mColumnName, Type.GetType("System.Int32"))

				Case FolderFileListColumn.Size.ToString

					'「サイズ」はLong型で作成
					mFolderFileList.Columns.Add(mColumnName, Type.GetType("System.Int64"))

				Case FolderFileListColumn.IsLastFileInFolder.ToString

					'「フォルダ内で最後のファイルか」はBoolean型で作成
					mFolderFileList.Columns.Add(mColumnName, Type.GetType("System.Boolean"))

				Case Else

					'デフォルトはString型で作成
					mFolderFileList.Columns.Add(mColumnName, Type.GetType("System.String"))

			End Select

		Next

		Return mFolderFileList

	End Function

#End Region

#Region "対象フォルダ内のファイル数を取得"

	''' <summary>対象フォルダ内のサブディレクトリとファイル数を取得</summary>
	''' <returns>対象フォルダ内のサブディレクトリとファイル数</returns>
	''' <remarks>
	'''  メモ：DirectoryクラスとDirectoryInfoクラスの違い
	'''    Directoryクラスはstaticなメソッドばかりからなるユーティリティ的なクラ
	'''    スであるのに対して、DirectoryInfoクラスではまず特定のディレクトリを指
	'''    定してインスタンスを作成し、それに対して各メソッドの呼び出しを行う。
	'''    1つのディレクトリに対して一連の操作を行う場合には、DirectoryInfoクラス
	'''    を使用すべきだろう。リファレンス・マニュアルには、Directoryクラスの
	'''    staticなメソッドでは、毎回ファイルに対するセキュリティのチェックが実行
	'''    されるが、DirectoryInfoクラスのインスタンス・メソッドでは必ずしもそう
	'''    ではないといったことも記述されている                                  
	''' </remarks>
	Public Function GetFileCountForTargetFolder() As Integer

		'対象フォルダ内のサブディレクトリとファイル数を取得
		Dim mDI As New DirectoryInfo(_TargetPath)
		Dim mFolderFileCount As Integer = mDI.GetFileSystemInfos("*", SearchOption.AllDirectories).Length + 1
		mDI = Nothing

		'対象フォルダパスが含まれていないので＋１する
		mFolderFileCount = mFolderFileCount + 1

		Return mFolderFileCount

	End Function

#End Region

#Region "ディレクトリのレベルを取得"

    ''' <summary>ディレクトリのレベルを取得</summary>
    ''' <param name="pStartingPointPath">起点となるパス</param>
    ''' <param name="pNowPath">調べたいパス</param>
    ''' <returns>パスの階層レベル</returns>
    ''' <remarks>起点となるパスからみて調べたいパスの階層レベルを返す</remarks>
    Private Function _GetDirectoryLevel(ByVal pStartingPointPath As String, ByVal pNowPath As String) As Integer

		Return pNowPath.Split("\").Length - pStartingPointPath.Split("\").Length

	End Function

#End Region

#Region "ファイルシステムタイプを取得"

	''' <summary>ファイルシステムタイプを取得</summary>
	''' <param name="pPath">対象ファイルのフルパス</param>
	''' <returns>ファイルのシステムタイプ</returns>
	''' <remarks></remarks>
	Private Function _GetFileSystemType(ByVal pPath As String) As FileSystemType

		If System.IO.Directory.Exists(pPath) Then

			Return FileSystemType.Folder

		ElseIf System.IO.File.Exists(pPath) Then

			Return FileSystemType.File

		Else

			Return FileSystemType.None

		End If

	End Function

#End Region

#Region "フォルダサイズを取得"

	''' <summary>フォルダサイズを取得</summary>
	''' <param name="pPath">対象フォルダ</param>
	''' <returns>フォルダサイズ</returns>
	''' <remarks>再帰処理を行い全てのファイルのサイズを足して計算していく
	'''          ※引数「pPath」がファイルだった時のことを考慮していないので注意</remarks>
	Private Function _GetFolderSize(ByVal pPath As String) As Long

		Dim mDI As New System.IO.DirectoryInfo(pPath)
		Dim mFolderSize As Long = 0

		'サブフォルダ内のサイズを計算
		Dim mSubFolders As System.IO.FileSystemInfo() = mDI.GetDirectories("*", System.IO.SearchOption.TopDirectoryOnly)
		For Each mFolder As System.IO.FileSystemInfo In mSubFolders

			'再帰処理を行いファイルサイズを計算していく
			mFolderSize = mFolderSize + _GetFolderSize(mFolder.FullName)

		Next

		'フォルダ内のファイルのサイズを計算
		Dim mFiles As System.IO.FileInfo() = mDI.GetFiles("*", IO.SearchOption.TopDirectoryOnly)
		For Each mFile As System.IO.FileInfo In mFiles

			mFolderSize = mFolderSize + mFile.Length

		Next

		Return mFolderSize

	End Function

#End Region

#Region "対象フォルダ以下のパスを取得"

	''' <summary>対象フォルダ以下のパスを取得</summary>
	''' <param name="pPath">対象ファイルのフルパス</param>
	''' <returns>対象ファイルのフルパスから対象フォルダの部分のパスを除いたもの</returns>
	''' <remarks>対象フォルダ          ：C:\Users\○○○○○\Desktop\test\
	'''          対象ファイルのフルパス：C:\Users\○○○○○\Desktop\test\ゴリラ.txt
	'''          最終的な結果のパス    ：test\ゴリラ.txt                             </remarks>
	Private Function _GetUnderTargetPath(ByVal pPath As String) As String

		'対象フォルダのフルパスから対象フォルダ名分引いた数を取得
		Dim mStartPlace As Integer = Me.TargetPath.Length - Me.TargetPathFolderName.Length

		Return pPath.Substring(mStartPlace)

	End Function

#End Region

#Region "フォルダ・ファイルリストを取得"

	''' <summary>フォルダ・ファイルリストを取得</summary>
	''' <param name="pPath">対象パス</param>
	''' <param name="pFolderFileList">フォルダファイルリストを格納するDataTable</param>
	''' <remarks>フォルダ・ファイルリストをDataTableに格納する
	'''          ※DataTableは参照型だからByRefにしなくても値を保持したままであった……。</remarks>
	Public Sub GetFolderFileList(ByVal pPath As String, ByVal pFolderFileList As DataTable)

		'対象となるフォルダパスと対象パスが同じ（対象フォルダ）だった時、フォルダ・ファイルリストへ対象フォルダ情報をセット
		If _TargetPath = pPath Then _TargetPathToList(_TargetPath, pFolderFileList, -1)

		'フォルダ内のフォルダとファイルのリストを取得
		Dim mDI As New System.IO.DirectoryInfo(pPath)
		Dim mSubFolders As System.IO.FileSystemInfo() = mDI.GetDirectories("*", System.IO.SearchOption.TopDirectoryOnly)
		Dim mFiles As System.IO.FileSystemInfo() = mDI.GetFiles("*", IO.SearchOption.TopDirectoryOnly)

		'ToDo:IComparerインターフェースを使ってフォルダとファイルの並びを自然な並び替えになるようにするのも有りかも
		'     「1,2,3,10,11,12,4,5」とつくファイルがあったら並び順としては「1,10,11,12,2,3,4,5」となってしまう……。

		'現在のパスのディレクトリレベルを取得
		Dim mDirectoryLevel As Integer = _GetDirectoryLevel(_TargetPath, pPath)

		'フォルダ情報をフォルダファイルリストへセット
		Call _AddFolderFileDataToList(mSubFolders, pFolderFileList, mDirectoryLevel, FileSystemType.Folder, mFiles.Length)

		'ファイル情報をフォルダファイルリストへセット
		Call _AddFolderFileDataToList(mFiles, pFolderFileList, mDirectoryLevel, FileSystemType.File)

	End Sub

#End Region

#Region "フォルダファイル情報をフォルダファイルリストへセット"

    ''' <summary>フォルダファイル情報をフォルダファイルリストへセット</summary>
    ''' <param name="pFolderFile">フォルダファイル情報</param>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <param name="pDirectoryLevel">ディレクトリレベル</param>
    ''' <param name="pFolderFileType">フォルダファイルタイプ</param>
    ''' <param name="pFilesCount">ファイル数</param>
    ''' <remarks>対象のフォルダファイル情報をフォルダファイルリストへセットする</remarks>
    Private Sub _AddFolderFileDataToList(ByVal pFolderFile() As FileSystemInfo, ByVal pFolderFileList As DataTable, ByVal pDirectoryLevel As Integer _
									   , ByVal pFolderFileType As FileSystemType, Optional pFilesCount As Integer = 0)

		For mFolderFileCounter As Integer = 0 To pFolderFile.GetUpperBound(0)

			'フォルダ内で最後のファイルかフラグ
			Dim mIsLastFileInFolder As Boolean = False

			Select Case pFolderFileType

				Case FileSystemType.Folder

					'サブフォルダにファイルが１件も無くかつ最後のフォルダの時、フォルダ内で最後のファイルかフラグにTrueをセット
					If pFilesCount <= 0 AndAlso mFolderFileCounter = pFolderFile.GetUpperBound(0) Then mIsLastFileInFolder = True

				Case FileSystemType.File

					'最後のファイルの時、フォルダ内で最後のファイルかフラグにTrueをセット
					If mFolderFileCounter = pFolderFile.GetUpperBound(0) Then mIsLastFileInFolder = True

			End Select

			'対象パス情報をフォルダ・ファイルリストへ追加
			Call _TargetPathToList(pFolderFile(mFolderFileCounter).FullName, pFolderFileList, pDirectoryLevel, mIsLastFileInFolder)

			If pFolderFileType = FileSystemType.Folder Then

				'再帰処理（サブディレクトリ内の処理）
				Call GetFolderFileList(pFolderFile(mFolderFileCounter).FullName, pFolderFileList)

			End If

		Next

	End Sub

#End Region

#Region "対象パス情報をフォルダ・ファイルリストへセット"

	''' <summary>対象パス情報をフォルダ・ファイルリストへセット</summary>
	''' <param name="pTargetPath">対象パス</param>
	''' <param name="pFolderFileList">フォルダ・ファイルリスト</param>
	''' <param name="pDirectoryLevel">ディレクトリレベル</param>
	''' <remarks></remarks>
	Private Sub _TargetPathToList(ByVal pTargetPath As String, ByVal pFolderFileList As DataTable, ByVal pDirectoryLevel As Integer)

		'最後のファイルかどうかのフラグ、最後のフォルダかどうかのフラグをfalseでセットして引数違いの関数の呼び出し
		Call _TargetPathToList(pTargetPath, pFolderFileList, pDirectoryLevel, False)

	End Sub

	''' <summary>対象パス情報をフォルダ・ファイルリストへセット</summary>
	''' <param name="pTargetPath">対象パス</param>
	''' <param name="pFolderFileList">フォルダ・ファイルリスト</param>
	''' <param name="pDirectoryLevel">ディレクトリレベル</param>
	''' <param name="pIsLastFileInFolder">最後のファイルかどうかのフラグ</param>
	''' <remarks>※出力用の文字列情報はここではセットしない</remarks>
	Private Sub _TargetPathToList(ByVal pTargetPath As String, ByVal pFolderFileList As DataTable, ByVal pDirectoryLevel As Integer, ByVal pIsLastFileInFolder As Boolean)

		Dim mFolderFile As New System.IO.FileInfo(pTargetPath)

		'ファイルシステムタイプを取得
		Dim mFileSystemType As FileSystemType = _GetFileSystemType(mFolderFile.FullName)
		Dim mFileSystemTypeName As String = mFileSystemType.ToString

		'拡張子、ファイルサイズを取得
		Dim mExtension As String = String.Empty
		Dim mFileSize As Long
		Select Case mFileSystemType

			Case FileSystemType.Folder

				'対象フォルダ以下のすべてのファイルサイズを合計した値を取得
				mFileSize = _GetFolderSize(mFolderFile.FullName)

			Case FileSystemType.File

				mFileSize = mFolderFile.Length

				'拡張子は「.○○○」の形式で取得されるので「.」を削除した形で取得
				mExtension = mFolderFile.Extension.Replace(".", "")

		End Select

		'ファイルサイズレベルを取得
		Dim mFileSizeLevel As cFileSizeLevel.SizeLevel = cFileSizeLevel.GetLevel(mFileSize)

		'新規行データを作成（空）
		Dim mAddDataRow As DataRow = pFolderFileList.NewRow

		'追加行にデータをセット
		mAddDataRow(FolderFileListColumn.No) = pFolderFileList.Rows.Count
		mAddDataRow(FolderFileListColumn.Name) = mFolderFile.Name
		mAddDataRow(FolderFileListColumn.UpdateDate) = mFolderFile.LastWriteTime
		mAddDataRow(FolderFileListColumn.FileSystemType) = mFileSystemType
		mAddDataRow(FolderFileListColumn.FileSystemTypeName) = mFileSystemTypeName
		mAddDataRow(FolderFileListColumn.Extension) = mExtension
		mAddDataRow(FolderFileListColumn.Size) = mFileSize
		mAddDataRow(FolderFileListColumn.SizeAndUnit) = ByteUnit.GetByteSizeAndUnit(mFileSize)
		mAddDataRow(FolderFileListColumn.SizeLevel) = mFileSizeLevel
		mAddDataRow(FolderFileListColumn.SizeLevelName) = mFileSizeLevel.ToString
		mAddDataRow(FolderFileListColumn.DirectoryLevel) = pDirectoryLevel
		mAddDataRow(FolderFileListColumn.ParentFolder) = mFolderFile.Directory.Name
		mAddDataRow(FolderFileListColumn.ParentFolderFullPath) = mFolderFile.Directory.FullName
		mAddDataRow(FolderFileListColumn.UnderTargetFolder) = _GetUnderTargetPath(mFolderFile.FullName)
		mAddDataRow(FolderFileListColumn.FullPath) = mFolderFile.FullName
		mAddDataRow(FolderFileListColumn.IsLastFileInFolder) = pIsLastFileInFolder

        '行データをDataTableにセット
        pFolderFileList.Rows.Add(mAddDataRow)

        'フォルダファイルリスト作業進捗報告
        '※処理進捗プロパティをセットしている時のみ報告を行う
        Call _ReportProcessProgress(pTargetPath)

    End Sub

#End Region

#Region "出力用の文字列をセット"

	''' <summary>出力用の文字列をセットする</summary>
	''' <param name="pFolderFileList">フォルダ・ファイルリスト</param>
	''' <remarks>画面に出力した際にフォルダ階層がひと目でわかるような文字列を作成していく</remarks>
	Public Sub SetOutputTextStringToFolderFileList(ByVal pFolderFileList As DataTable)

		'「｜」を出力するかのフラグ
		Dim mNotPrintVerticalBar() As Boolean

		'フォルダファイルリストの行数分繰り返す
		For Each mFolderFileListRow As DataRow In pFolderFileList.Rows

			Dim mDispString As New System.Text.StringBuilder												   '出力用文字列
			Dim mCurrentRowDirectoryLevel As Integer = mFolderFileListRow(FolderFileListColumn.DirectoryLevel) '現在行のディレクトリレベル

			'現在行のディレクトリレベルが「-1」の時（対象パス）、出力用文字列にファイル名をセットする
			If mCurrentRowDirectoryLevel = -1 Then

				mFolderFileListRow(FolderFileListColumn.DispString) = mFolderFileListRow(FolderFileListColumn.Name)
				Continue For

			End If

			'ディレクトリレベル分、枠を作成
			ReDim Preserve mNotPrintVerticalBar(mCurrentRowDirectoryLevel)

			'ディレクトリレベル分繰り返す
			For i As Integer = 0 To mCurrentRowDirectoryLevel

				'フォルダファイルリストの現在行のディレクトリレベルと同じ時
				If i = mFolderFileListRow(FolderFileListColumn.DirectoryLevel) Then

					'----------------------------------
					' ファイル名作成処理
					'----------------------------------
					'※フォルダ内の最後のファイルかどうかでファイル名の前に付加する文字列を変更する
					'  最後のファイルじゃない：「├」、最後のファイル：「└」
					'  例：├ ゴリラ.txt,└ ブタゴリラ.txt
					If mFolderFileListRow(FolderFileListColumn.IsLastFileInFolder) Then

						mDispString.Append(_cOutputString.LastFile & mFolderFileListRow(FolderFileListColumn.Name))

					Else

						mDispString.Append(_cOutputString.File & mFolderFileListRow(FolderFileListColumn.Name))

					End If

					'「｜」を出力するかのフラグに値をセット
					'※１番下の階層のファイルから見て１番上までの階層のファイルの「｜」を出力するかどうかの状態を保持する
					mNotPrintVerticalBar(i) = mFolderFileListRow(FolderFileListColumn.IsLastFileInFolder)

					'出力用文字列項目に値をセット
					mFolderFileListRow(FolderFileListColumn.DispString) = mDispString.ToString

				Else

					'----------------------------------
					' ファイル名の前文字列作成処理
					'----------------------------------
					'※１階層目、２階層目……それぞれで「｜」、「  」のどちらかを印字するかどうかを
					'  持ったフラグによりファイル名の前の文字列を付加していく
					If mNotPrintVerticalBar(i) = True Then

						mDispString.Append(_cOutputString.NotExistsFile)

					Else

						mDispString.Append(_cOutputString.ExistsFile)

					End If

				End If

			Next

		Next

	End Sub

#End Region

#Region "拡張子リストを取得"

	''' <summary>拡張子リストを取得</summary>
	''' <param name="pFolderFileList"></param>
	''' <returns>拡張子リスト</returns>
	''' <remarks>引数で渡されたリストの拡張子リストを返します</remarks>
	Private Function _GetExtensionList(ByVal pFolderFileList As DataTable) As ArrayList

		'初期値のリスト（空文字）を追加
		Dim mExtensioinListArray As New ArrayList
		mExtensioinListArray.Add(String.Empty)

		'拡張子項目でグループ化（ファイル項目のみを抽出）し、空文字を除くクエリーを作成
		'※拡張子はすべて小文字に変換する（大文字と小文字違いの拡張子をリストに表示させないため）
		Dim mQuery = From Extensions In pFolderFileList
					 Where Extensions(FolderFileListColumn.FileSystemType.ToString) = FileSystemType.File
					 Group By Extension = Extensions(FolderFileListColumn.Extension).ToString.ToLower()
					 Into Group
					 Where Extension <> String.Empty


		'初期値以外のリストを追加
		For Each mDr In mQuery

			mExtensioinListArray.Add(mDr.Extension)

		Next

		Return mExtensioinListArray

	End Function

#End Region

#Region "指定カラムで並び替えたフォルダファイルリストを取得"

	''' <summary>指定カラムで並び替えたフォルダファイルリストを取得</summary>
	''' <param name="pFolderFileList">フォルダファイルリスト</param>
	''' <param name="pColumn">対象カラム</param>
	''' <param name="pSortOrder">ソート条件</param>
	''' <returns>対象カラムをソート条件で並び替えたフォルダファイルリスト</returns>
	''' <remarks></remarks>
	Public Shared Function GetSpecifiedColumnFolderFileList(ByVal pFolderFileList As DataTable, ByVal pColumn As FolderFileListColumn, ByVal pSortOrder As SortOrder) As DataTable

		'指定カラムで並び替え
		Select Case pSortOrder

			Case SortOrder.Ascending

				Return GetSpecifiedColumnFolderFileList(pFolderFileList, pColumn, FolderFileListSortOrder.ASC)

			Case SortOrder.Descending

				Return GetSpecifiedColumnFolderFileList(pFolderFileList, pColumn, FolderFileListSortOrder.DESC)

			Case Else

				Return GetSpecifiedColumnFolderFileList(pFolderFileList, pColumn, FolderFileListSortOrder.ASC)

		End Select

	End Function

	''' <summary>指定カラムで並び替えたフォルダファイルリストを取得</summary>
	''' <param name="pFolderFileList">フォルダファイルリスト</param>
	''' <param name="pColumn">対象カラム</param>
	''' <param name="pSortOrder">ソート条件</param>
	''' <returns>対象カラムをソート条件で並び替えたフォルダファイルリスト</returns>
	''' <remarks></remarks>
	Public Shared Function GetSpecifiedColumnFolderFileList(ByVal pFolderFileList As DataTable, ByVal pColumn As FolderFileListColumn, ByVal pSortOrder As FolderFileListSortOrder) As DataTable

		'指定カラムで並び替え
		Select Case pSortOrder

			Case FolderFileListSortOrder.ASC

				Dim mQuery = From RefineByFolderFileList In pFolderFileList.AsEnumerable
							 Order By RefineByFolderFileList(pColumn) Ascending
							 Select RefineByFolderFileList

				'クエリーからDataViewを作成
				Dim mQueryToDataView As DataView = mQuery.AsDataView

				Return mQueryToDataView.ToTable

			Case FolderFileListSortOrder.DESC

				Dim mQuery = From RefineByFolderFileList In pFolderFileList.AsEnumerable
							 Order By RefineByFolderFileList(pColumn) Descending
							 Select RefineByFolderFileList

				'クエリーからDataViewを作成
				Dim mQueryToDataView As DataView = mQuery.AsDataView

				Return mQueryToDataView.ToTable

			Case Else

				'デフォルトは昇順
				Dim mQuery = From RefineByFolderFileList In pFolderFileList.AsEnumerable
							 Order By RefineByFolderFileList(pColumn) Ascending
							 Select RefineByFolderFileList

				'クエリーからDataViewを作成
				Dim mQueryToDataView As DataView = mQuery.AsDataView

				Return mQueryToDataView.ToTable

		End Select

	End Function

#End Region

#Region "フォルダファイルリスト絞り込みデータを取得"

	''' <summary>フォルダファイルリスト絞り込みデータを取得</summary>
	''' <param name="pFolderFileList">フォルダファイルリスト</param>
	''' <param name="pColumn">フォルダファイルリストカラム</param>
	''' <param name="pWhereString">絞り込み条件</param>
	''' <returns>フォルダファイルリストを絞り込み条件で絞り込んだデータ</returns>
	''' <remarks></remarks>
	Public Function GetRefineByFolderFileList(ByVal pFolderFileList As DataTable, ByVal pColumn As FolderFileListColumn, ByVal pWhereString As String) As DataTable

		'条件が空だったら引数で渡ってきたリストをそのまま返す
		If String.IsNullOrEmpty(pWhereString) Then Return pFolderFileList

		'絞り込む文字列を小文字に変換
		Dim mLowerWhere As String = pWhereString.ToLower()

		'カラムごと絞り込み条件を作成する
		Select Case pColumn

			Case FolderFileListColumn.Name

				'引数で渡ってきたリストから名前で条件を絞り込んでデータを取得するクエリーを作成
				'※大文字小文字は区別しないでLike検索
				Dim mQuery = From RefineByFolderFileList In pFolderFileList.AsEnumerable
							 Where RefineByFolderFileList(pColumn).ToString.ToLower().Contains(mLowerWhere)
							 Order By RefineByFolderFileList(FolderFileListColumn.No) Ascending
							 Select RefineByFolderFileList

				'クエリーからDataViewを作成
				Dim mQueryToDataView As DataView = mQuery.AsDataView

				Return mQueryToDataView.ToTable

			Case FolderFileListColumn.Extension

				'引数で渡ってきたリストから拡張子で条件を絞り込んでデータを取得するクエリーを作成
				Dim mQuery = From RefineByFolderFileList In pFolderFileList.AsEnumerable
							 Where RefineByFolderFileList(pColumn).ToString.ToLower().ToString = mLowerWhere
							 Order By RefineByFolderFileList(FolderFileListColumn.No) Ascending
							 Select RefineByFolderFileList

				'クエリーからDataViewを作成
				Dim mQueryToDataView As DataView = mQuery.AsDataView

				Return mQueryToDataView.ToTable

			Case FolderFileListColumn.SizeLevel

				'空指定の時
				If mLowerWhere = cFileSizeLevel.LevelListNoneDataID Then

					'NOの昇順で並び替えたものを返す
					'※①何も指定しない→②ヘッダーセルで並び替え→③ファイルサイズレベル指定→④何も指定しない
					'  ②の時の並び替えを保持したままになってしまうのでその対応
					Return GetSpecifiedColumnFolderFileList(pFolderFileList, FolderFileListColumn.No, FolderFileListSortOrder.ASC)

				Else

					'引数で渡ってきたリストからファイルサイズレベルで条件を絞り込んでデータを取得するクエリーを作成
					Dim mQuery = From RefineByFolderFileList In pFolderFileList.AsEnumerable
								 Where RefineByFolderFileList(pColumn).ToString = mLowerWhere
								 Order By RefineByFolderFileList(FolderFileListColumn.SizeLevel) Ascending
								 Select RefineByFolderFileList

					'クエリーからDataViewを作成
					Dim mQueryToDataView As DataView = mQuery.AsDataView

					Return mQueryToDataView.ToTable

				End If

			Case Else

				'名前カラム、拡張子カラム、ファイルサイズレベルカラム以外はそのまま返す
				Return pFolderFileList

		End Select

	End Function

#End Region

#Region "フォルダファイルリスト進捗率報告"

    ''' <summary>フォルダファイルリスト進捗報告</summary>
    ''' <param name="pPath">処理フォルダファイル</param>
    ''' <remarks></remarks>
    Private Sub _ReportProcessProgress(ByVal pPath As String)

        'フォルダファイルリストの進捗状況を報告する変数がNothingで無かった時
        If Not _ProcessProgresss Is Nothing Then

            '現在の処理進捗率を計算
            Dim mProcessPercent As Integer = (_FolderFileList.Rows.Count / _FolderFileListCount) * 100

            '進捗率を報告
            _ProcessProgresss.Report(New FolderFileListProgress(mProcessPercent, pPath))

        End If

    End Sub

#End Region

#End Region

End Class