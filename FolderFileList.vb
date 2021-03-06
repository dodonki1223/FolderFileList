﻿Option Explicit On

Imports System.IO
Imports System.Data.Linq

''' <summary>
'''   フォルダファイルリストの機能を提供する
''' </summary>
''' <remarks></remarks>
Public Class FolderFileList

#Region "定数"

    ''' <summary>
    '''   出力用の文字列を提供する
    ''' </summary>
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

    ''' <summary>
    '''   ファイルサイズレベルを提供する
    ''' </summary>
    ''' <remarks></remarks>
    Public Class cFileSizeLevel

        ''' <summary>10GB</summary>
        Private Const _10GB As Long = 10737418240

        ''' <summary>1GB</summary>
        Private Const _1GB As Long = 1073741824

        ''' <summary>100MB</summary>
        Private Const _100MB As Long = 104857600

        ''' <summary>1MB</summary>
        Private Const _1MB As Long = 1048576

        ''' <summary>1KB</summary>
        Private Const _1KB As Long = 1024

        ''' <summary>ファイルサイズレベルリストのからデータID</summary>
        Public Const LevelListNoneDataID As Integer = 0

        ''' <summary>
        '''   ファイルサイズのレベル
        ''' </summary>
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

        ''' <summary>
        '''   ファイルサイズのレベルリストカラム
        ''' </summary>
        ''' <remarks>
        '''   コンボボックスの表示用のテキストと内部で持つIDの設定用
        ''' </remarks>
        Public Enum LevelListColum

            ''' <summary>ID</summary>
            ''' <remarks></remarks>
            ID

            ''' <summary>名前</summary>
            ''' <remarks></remarks>
            NAME

        End Enum

        ''' <summary>
        '''   ファイルサイズレベルを取得
        ''' </summary>
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

        ''' <summary>
        '''   ファイルサイズレベルのカラーを取得
        ''' </summary>
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

        ''' <summary>
        '''   ファイルサイズレベルリストを取得
        ''' </summary>
        ''' <param name="pInitialString">リスト初期表示文字列</param>
        ''' <returns>ファイルサイズレベルのリスト</returns>
        ''' <remarks></remarks>
        Public Shared Function LevelList(Optional ByVal pInitialString As String = "") As DataTable

            Dim mLevelList As New DataTable()

            'カラムを作成
            mLevelList.Columns.Add(LevelListColum.ID.ToString, GetType(Integer))
            mLevelList.Columns.Add(LevelListColum.NAME.ToString, GetType(String))

            '空データを追加
            Dim mNoneRow As DataRow = mLevelList.NewRow

            mNoneRow(LevelListColum.ID) = 0
            mNoneRow(LevelListColum.NAME) = pInitialString

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

    ''' <summary>
    '''   エンコーディングを提供する
    ''' </summary>
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

    ''' <summary>
    '''   ファイルシステムタイプ
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FileSystemType

        ''' <summary>ドライブ</summary>
        Drive = 1

        ''' <summary>フォルダ</summary>
        Folder

        ''' <summary>ファイル</summary>
        File

        ''' <summary>不明</summary>
        None

    End Enum

    ''' <summary>
    '''   フォルダファイルリストデータカラム
    ''' </summary>
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

    ''' <summary>
    '''   フォルダファイルリストデータカラム（日本語名）
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FolderFileListJapaneseColumn

        NO

        ﾌｧｲﾙ名

        更新日時

        ﾌｧｲﾙｼｽﾃﾑﾀｲﾌﾟ値

        ﾌｧｲﾙｼｽﾃﾑﾀｲﾌﾟ

        拡張子

        ﾌｧｲﾙｻｲｽﾞ単位無し

        ﾌｧｲﾙｻｲｽﾞ

        ﾌｧｲﾙｻｲｽﾞﾚﾍﾞﾙ値

        ﾌｧｲﾙｻｲｽﾞﾚﾍﾞﾙ

        ﾃﾞｨﾚｸﾄﾘﾚﾍﾞﾙ

        親ﾌｫﾙﾀﾞ

        親ﾌｫﾙﾀﾞﾌﾙﾊﾟｽ

        対象ﾌｫﾙﾀﾞ以下

        ﾌﾙﾊﾟｽ

        最後のﾌｧｲﾙか

        表示文字列

    End Enum

    ''' <summary>
    '''   フォルダファイルリスト並び順
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FolderFileListSortOrder

        ''' <summary>昇順</summary>
        ASC

        ''' <summary>降順</summary>
        DESC

    End Enum

    ''' <summary>
    '''   フォルダファイルリスト処理状態
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FolderFileListProcessState

        ''' <summary>ファイル数をセット</summary>
        SetFileCount

        ''' <summary>フォルダファイルリストを作成</summary>
        CreateFolderFileList

        ''' <summary>「最後のファイルかどうか」にセット</summary>
        SetIsLastFileInFolder

    End Enum

#End Region

#Region "変数"

    ''' <summary>フォルダファイルリスト</summary>
    Private _FolderFileList As DataTable

    ''' <summary>対象となるフォルダパス</summary>
    Private _TargetPath As String

    ''' <summary>処理進捗報告をするかどうか</summary>
    Private _IsReportProcessProgress As Boolean = False

    ''' <summary>フォルダファイルリストのファイル数</summary>
    ''' <remarks>フォルダとファイルのファイル数を保持する</remarks>
    Private _FolderFileListCount As Integer

    ''' <summary>フォルダファイルリストの進捗状況を報告する変数</summary>
    ''' <remarks>フォルダファイルリスト作成中の進捗状況報告用</remarks>
    Private _ProcessProgresss As IProgress(Of FolderFileListProgress) = Nothing

#End Region

#Region "構造体"

    ''' <summary>
    '''   フォルダファイルリスト進捗率構造体
    ''' </summary>
    ''' <remarks>
    '''   フォルダファイルリストの進捗率の内容を保持する構造体
    ''' </remarks>
    Public Structure FolderFileListProgress

        ''' <summary>進捗率</summary>
        Private _Percent As Integer

        ''' <summary>処理対象フォルダファイル</summary>
        Private _ProcessingFolderFile As String

        ''' <summary>フォルダファイルリスト処理状態</summary>
        Private _FolderFileListProcessState As FolderFileListProcessState

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

        ''' <summary>フォルダファイルリスト処理状態</summary>
        ''' <remarks></remarks>
        Public Property FolderFileListProcessState As FolderFileListProcessState

            Set(value As FolderFileListProcessState)

                _FolderFileListProcessState = value

            End Set
            Get

                Return _FolderFileListProcessState

            End Get

        End Property

        ''' <summary>
        '''   コンストラクタ
        ''' </summary>
        ''' <param name="pPercent">進捗率</param>
        ''' <param name="pProcessingFolderFile">処理フォルダファイル</param>
        ''' <param name="pFolderFileListProcessState">フォルダファイルリスト処理状態</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pPercent As Integer, ByVal pProcessingFolderFile As String, ByVal pFolderFileListProcessState As FolderFileListProcessState)

            '進捗率をセットをセット
            _Percent = pPercent

            '処理フォルダファイルをセット
            _ProcessingFolderFile = pProcessingFolderFile

            'フォルダファイルリスト処理状態をセット
            _FolderFileListProcessState = pFolderFileListProcessState

        End Sub

    End Structure

#End Region

#Region "プロパティ"

    ''' <summary>
    '''   フォルダファイルリストプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property FolderFileList As DataTable

        Get

            Return _FolderFileList

        End Get

    End Property

    ''' <summary>
    '''   対象となるフォルダパスプロパティ（フルパス）
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property TargetPath As String

        Get

            Return _TargetPath

        End Get

    End Property

    ''' <summary>
    '''   対象となるフォルダ名プロパティ（フォルダ名のみ）
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property TargetPathFolderName As String

        Get

            Dim mFolder As New System.IO.FileInfo(_TargetPath)
            Return mFolder.Name

        End Get

    End Property

    ''' <summary>
    '''   対象となるフォルダのファイルリスト（文字列のリスト）
    ''' </summary>
    ''' <remarks>
    '''   システムで使用するファイルを除いてリストを取得する
    '''   ※「System Volume Information」などにアクセスすると例外で落ちるため
    ''' </remarks>
    Public ReadOnly Property TargetPathFilesList As List(Of String)

        Get

            Return _GetFilesForTargetPath(_TargetPath)

        End Get

    End Property

    ''' <summary>
    '''   フォルダリストプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property FolderFileListOnlyFolder As DataTable

        Get

            'フォルダファイルリストからフォルダのみを絞り込んで取得するクエリーを作成
            Dim mQuery = From FolderFileList In _FolderFileList.AsEnumerable
                         Where FolderFileList(FolderFileListColumn.FileSystemType.ToString) = FileSystemType.Folder _
                            Or FolderFileList(FolderFileListColumn.FileSystemType.ToString) = FileSystemType.Drive
                         Order By FolderFileList(FolderFileListColumn.No) Ascending
                         Select FolderFileList

            'クエリーからDataViewを作成
            Dim mQueryToDataView As DataView = mQuery.AsDataView

            Return mQueryToDataView.ToTable

        End Get

    End Property

    ''' <summary>
    '''   ファイルリストプロパティ
    ''' </summary>
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

    ''' <summary>
    '''   出力用文字列プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property OutputText As String

        Get

            Return GetOutputTextString(_FolderFileList)

        End Get

    End Property

    ''' <summary>
    '''   拡張子リストプロパティ
    ''' </summary>
    ''' <param name="pInitialString">リスト初期表示文字列</param>
    ''' <remarks>
    '''   対象となるフォルダパスで取得されたファイル群から作成した拡張子リストを返す
    ''' </remarks>
    Public ReadOnly Property ExtensionList(Optional ByVal pInitialString As String = "") As ArrayList

        Get

            '初期値のリスト（初期表示文字列）を追加
            Dim mExtensioinListArray As New ArrayList
            mExtensioinListArray.Add(pInitialString)

            '拡張子項目でグループ化（ファイル項目のみを抽出）し、初期表示文字列と空文字を除くクエリーを作成
            '※拡張子はすべて小文字に変換する（大文字と小文字違いの拡張子をリストに表示させないため）
            Dim mQuery = From Extensions In _FolderFileList
                         Order By Extensions(FolderFileListColumn.Extension) Ascending
                         Where Extensions(FolderFileListColumn.FileSystemType.ToString) = FileSystemType.File
                         Group By Extension = Extensions(FolderFileListColumn.Extension).ToString.ToLower()
                         Into Group
                         Where Extension <> pInitialString
                         Where Extension <> String.Empty


            '初期値以外のリストを追加
            For Each mDr In mQuery

                mExtensioinListArray.Add(mDr.Extension)

            Next

            Return mExtensioinListArray

        End Get

    End Property

    ''' <summary>
    '''   処理進捗報告をするかどうか（デフォルトはOFF）
    ''' </summary>
    ''' <remarks>
    '''   Trueをセットしても処理進捗用変数がNothingの時はONにならない
    '''   ※ファイル数が多い場合は時間がかかるので注意
    ''' </remarks>
    Public WriteOnly Property IsReportProcessProgress As Boolean

        Set(value As Boolean)

            '処理進捗用の変数がNothingでない時
            If Not _ProcessProgresss Is Nothing Then

                _IsReportProcessProgress = value

                '処理進捗報告を行う時
                If _IsReportProcessProgress Then

                    'フォルダファイルリストのファイル数のセットの進捗報告
                    _ReportSetFileCount()

                    'フォルダファイルリストのファイル数をセット ※ファイル数が多い場合は時間がかかるので注意
                    _FolderFileListCount = Me.TargetPathFilesList.Count

                End If

            End If

        End Set

    End Property

    ''' <summary>
    '''   処理進捗プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property ProcessProgress As IProgress(Of FolderFileListProgress)

        Set(value As IProgress(Of FolderFileListProgress))

            _ProcessProgresss = value

        End Set

    End Property

#End Region

#Region "コンストラクタ"

    ''' <summary>
    '''   コンストラクタ
    ''' </summary>
    ''' <remarks>
    '''   引数無しは外部に公開しない
    ''' </remarks>
    Private Sub New()

    End Sub

    ''' <summary>
    '''   コンストラクタ
    ''' </summary>
    ''' <param name="pPath">フォルダファイルリストを取得したいフォルダパス</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pPath As String)

        '対象パスをセット
        _TargetPath = pPath

        'フォルダファイルリストのDataTableのカラムを作成
        _FolderFileList = _CreateFolderFileListColumns()

    End Sub

    ''' <summary>
    '''   コンストラクタ
    ''' </summary>
    ''' <param name="pPath">フォルダファイルリストを取得したいフォルダパス</param>
    ''' <param name="pProcessProgress">フォルダファイルリスト処理進捗</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pPath As String, ByVal pProcessProgress As IProgress(Of FolderFileListProgress))

        '対象パスをセット
        _TargetPath = pPath

        'フォルダファイルリストのDataTableのカラムを作成
        _FolderFileList = _CreateFolderFileListColumns()

        'フォルダファイルリストの進捗をセット
        _ProcessProgresss = pProcessProgress

    End Sub

#End Region

#Region "メソッド"

#Region "フォルダファイルリストDataTableのカラムを作成する"

    ''' <summary>
    '''   フォルダファイルリストDataTableのカラムを作成する
    ''' </summary>
    ''' <returns>フォルダファイルリストのカラム作成後のDataTable</returns>
    ''' <remarks>
    '''   FolderFileListColumn列挙体の項目分カラムを作成する
    ''' </remarks>
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

#Region "対象パスのファイル一覧をListで取得"

    ''' <summary>
    '''   対象パスのファイルの一覧をListで取得
    ''' </summary>
    ''' <param name="pPath">対象パス</param>
    ''' <returns>対象パスのフルパスのファイルリスト</returns>
    ''' <remarks>
    '''   システムで使用するファイルを除いてリストを取得する
    '''   ※「System Volume Information」などにアクセスすると例外で落ちるため
    ''' </remarks>
    Private Function _GetFilesForTargetPath(ByVal pPath As String) As List(Of String)

        Dim mFilesForTargetPath As New List(Of String)
        Dim mDI As New System.IO.DirectoryInfo(pPath)

        Try

            '引数のパスをセット
            mFilesForTargetPath.Add(pPath)

            '対象パス内のフォルダを取得
            Dim mSubFolders As System.IO.FileSystemInfo() = mDI.GetDirectories("*", System.IO.SearchOption.TopDirectoryOnly)

            '対象パス内のフォルダ数分繰り返す
            For Each mFolder As System.IO.FileSystemInfo In mSubFolders

                'システムファイル以外の時
                If (mFolder.Attributes And FileAttributes.System) <> FileAttributes.System Then

                    'フォルダ内のファイルを取得（再帰処理）
                    Dim mSubFolderFiles As New List(Of String)
                    mSubFolderFiles = _GetFilesForTargetPath(mFolder.FullName)

                    'ファイル数分繰り返す
                    For Each mSubFolderFile As String In mSubFolderFiles

                        'ファイルをセット
                        mFilesForTargetPath.Add(mSubFolderFile)

                    Next

                End If

            Next

            '対象パス内のファイルを取得
            Dim mFiles As System.IO.FileInfo() = mDI.GetFiles("*", IO.SearchOption.TopDirectoryOnly)

            '対象パス内のファイル数分繰り返す
            For Each mFile As System.IO.FileInfo In mFiles

                'システムファイル以外の時、ファイルをセット
                If (mFile.Attributes And FileAttributes.System) <> FileAttributes.System Then mFilesForTargetPath.Add(mFile.FullName)

            Next

        Catch mGeneratedExceptionName As System.UnauthorizedAccessException

            '例外が発生した場合は無視する
            '※「System Volume Information」等にアクセスした時にここにくる
            '  システムで管理しているファイルはリストに不必要なので何もしない

        End Try

        Return mFilesForTargetPath

    End Function

#End Region

#Region "ディレクトリのレベルを取得"

    ''' <summary>
    '''   ディレクトリのレベルを取得
    ''' </summary>
    ''' <param name="pStartingPointPath">起点となるパス</param>
    ''' <param name="pTargetPath">調べたいパス</param>
    ''' <returns>パスの階層レベル</returns>
    ''' <remarks>
    '''   起点となるパスからみて調べたいパスの階層レベルを返す
    ''' </remarks>
    Public Function _GetDirectoryLevel(ByVal pStartingPointPath As String, ByVal pTargetPath As String) As Integer

        Return _GetDirectoryLevel(pTargetPath) - _GetDirectoryLevel(pStartingPointPath)

    End Function

    ''' <summary>
    '''   ディレクトリのレベルを取得
    ''' </summary>
    ''' <param name="pPath">対象パス</param>
    ''' <returns>パスの階層レベル</returns>
    ''' <remarks>
    '''   パス文字列からSplitメソッドを使い「\」で切り分けて配列数に応じて階層レベルを取得する
    '''   ドライブ文字列「D:/」の時は配列数が２つ取得できてしまうので２つ目が空文字であった場合は削除する
    ''' </remarks>
    Public Function _GetDirectoryLevel(ByVal pPath As String) As Integer

        'パス文字列を「\」で切り分ける
        Dim mPath As String() = pPath.Split("\")

        '切り分けた要素中の最後の値が空文字の時
        If String.IsNullOrEmpty(mPath(mPath.GetUpperBound(0))) Then

            '配列の要素数を１つ減らす
            ReDim Preserve mPath(mPath.GetUpperBound(0) - 1)

        End If

        Return mPath.Length

    End Function

#End Region

#Region "ファイルシステムタイプを取得"

    ''' <summary>
    '''   ファイルシステムタイプを取得
    ''' </summary>
    ''' <param name="pPath">対象ファイルのフルパス</param>
    ''' <returns>ファイルのシステムタイプ</returns>
    ''' <remarks></remarks>
    Private Function _GetFileSystemType(ByVal pPath As String) As FileSystemType

        If System.IO.Directory.Exists(pPath) Then

            'パス文字列が３文字以内の時 例：D:\、E:\
            If pPath.Length <= 3 Then

                Return FileSystemType.Drive

            Else

                Return FileSystemType.Folder

            End If

        ElseIf System.IO.File.Exists(pPath) Then

            Return FileSystemType.File

        Else

            Return FileSystemType.None

        End If

    End Function

#End Region

#Region "フォルダサイズを取得"

    ''' <summary>
    '''   フォルダサイズを取得
    ''' </summary>
    ''' <param name="pPath">対象フォルダ</param>
    ''' <returns>フォルダサイズ</returns>
    ''' <remarks>
    '''   再帰処理を行い全てのファイルのサイズを足して計算していく、ただしシステムファイルは計算しない
    '''   ※引数「pPath」がファイルだった時のことを考慮していないので注意
    ''' </remarks>
    Private Function _GetFolderSize(ByVal pPath As String) As Long

        Dim mDI As New System.IO.DirectoryInfo(pPath)
        Dim mFolderSize As Long = 0

        Try

            '----------------------------------------
            ' フォルダ内のサブフォルダのサイズを計算
            '----------------------------------------
            '対象フォルダ内のサブフォルダを取得
            Dim mSubFolders As System.IO.FileSystemInfo() = mDI.GetDirectories("*", System.IO.SearchOption.TopDirectoryOnly)

            'サブフォルダ分繰り返す
            For Each mFolder As System.IO.FileSystemInfo In mSubFolders

                '対象フォルダが「システムファイル」の時は次の繰り返しへ
                If (mFolder.Attributes And FileAttributes.System) = FileAttributes.System Then Continue For

                '再帰処理を行いファイルサイズを計算
                mFolderSize = mFolderSize + _GetFolderSize(mFolder.FullName)

            Next

            '----------------------------------------
            ' フォルダ内のファイルのサイズを計算
            '----------------------------------------
            'フォルダ内のファイルを取得
            Dim mFiles As System.IO.FileInfo() = mDI.GetFiles("*", IO.SearchOption.TopDirectoryOnly)

            'ファイル分繰り返す
            For Each mFile As System.IO.FileInfo In mFiles

                '対象ファイルが「システムファイル」の時は次の繰り返しへ
                If (mFile.Attributes And FileAttributes.System) = FileAttributes.System Then Continue For

                'ファイルサイズを計算
                mFolderSize = mFolderSize + mFile.Length

            Next

        Catch mGeneratedExceptionName As System.UnauthorizedAccessException

            '例外が発生した場合は無視する
            '※「System Volume Information」等にアクセスした時にここにくる
            '  システムで管理しているファイルはリストに不必要なので何もしない

        End Try

        Return mFolderSize

    End Function

#End Region

#Region "対象フォルダ以下のパスを取得"

    ''' <summary>
    '''   対象フォルダ以下のパスを取得
    ''' </summary>
    ''' <param name="pPath">対象ファイルのフルパス</param>
    ''' <returns>対象ファイルのフルパスから対象フォルダの部分のパスを除いたもの</returns>
    ''' <remarks>
    '''   対象フォルダ          ：C:\Users\○○○○○\Desktop\test\
    '''   対象ファイルのフルパス：C:\Users\○○○○○\Desktop\test\ゴリラ.txt
    '''   最終的な結果のパス    ：test\ゴリラ.txt                             
    ''' </remarks>
    Private Function _GetUnderTargetPath(ByVal pPath As String) As String

        '対象フォルダのフルパスから対象フォルダ名分引いた数を取得
        Dim mStartPlace As Integer = Me.TargetPath.Length - Me.TargetPathFolderName.Length

        Return pPath.Substring(mStartPlace)

    End Function

#End Region

#Region "フォルダファイルリストを作成"

    ''' <summary>
    '''   フォルダファイルリストを作成
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateFolderFileList()

        'フォルダファイルリストを作成
        Call _GetFolderFileList(_TargetPath, _FolderFileList)

        '最後のファイルかどうかをセット
        Call SetIsLastFileInFolder(_FolderFileList)

    End Sub

    ''' <summary>
    '''   フォルダファイルリストを取得
    ''' </summary>
    ''' <param name="pPath">対象パス</param>
    ''' <param name="pFolderFileList">フォルダファイルリストを格納するDataTable</param>
    ''' <remarks>
    '''   フォルダ・ファイルリストをDataTableに格納する
    ''' </remarks>
    Private Sub _GetFolderFileList(ByVal pPath As String, ByVal pFolderFileList As DataTable)

        '対象となるフォルダパスと対象パスが同じ（対象フォルダ）だった時、フォルダ・ファイルリストへ対象フォルダ情報をセット
        If _TargetPath = pPath Then _TargetPathToList(_TargetPath, pFolderFileList, -1)

        Try

            'フォルダ内のフォルダとファイルのリストを取得
            Dim mDI As New System.IO.DirectoryInfo(pPath)
            Dim mSubFolders As System.IO.FileSystemInfo() = mDI.GetDirectories("*", System.IO.SearchOption.TopDirectoryOnly)
            Dim mFiles As System.IO.FileSystemInfo() = mDI.GetFiles("*", IO.SearchOption.TopDirectoryOnly)

            'ToDo:IComparerインターフェースを使ってフォルダとファイルの並びを自然な並び替えになるようにするのも有りかも
            '     「1,2,3,10,11,12,4,5」とつくファイルがあったら並び順としては「1,10,11,12,2,3,4,5」となってしまう……。

            '現在のパスのディレクトリレベルを取得
            Dim mDirectoryLevel As Integer = _GetDirectoryLevel(_TargetPath, pPath)

            'フォルダ情報をフォルダファイルリストへセット
            Call _AddFolderFileDataToList(mSubFolders, pFolderFileList, mDirectoryLevel, FileSystemType.Folder)

            'ファイル情報をフォルダファイルリストへセット
            Call _AddFolderFileDataToList(mFiles, pFolderFileList, mDirectoryLevel, FileSystemType.File)

        Catch mGeneratedExceptionName As System.UnauthorizedAccessException

            '例外が発生した場合は無視する
            '※「System Volume Information」等にアクセスした時にここにくる
            '  システムで管理しているファイルはリストに不必要なので何もしない

        End Try

    End Sub

    ''' <summary>
    '''   フォルダファイル情報をフォルダファイルリストへセット
    ''' </summary>
    ''' <param name="pFolderFile">フォルダファイル情報</param>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <param name="pDirectoryLevel">ディレクトリレベル</param>
    ''' <param name="pFolderFileType">フォルダファイルタイプ</param>
    ''' <remarks>
    '''   対象のフォルダファイル情報をフォルダファイルリストへセットする
    ''' </remarks>
    Private Sub _AddFolderFileDataToList(ByVal pFolderFile() As FileSystemInfo, ByVal pFolderFileList As DataTable, ByVal pDirectoryLevel As Integer _
                                       , ByVal pFolderFileType As FileSystemType)

        'フォルダファイル情報分繰り返す
        For mFolderFileCounter As Integer = 0 To pFolderFile.GetUpperBound(0)

            '対象ファイルが「システムファイル」の時は次の繰り返しへ
            If (pFolderFile(mFolderFileCounter).Attributes And FileAttributes.System) = FileAttributes.System Then Continue For

            '対象パス情報をフォルダ・ファイルリストへ追加
            Call _TargetPathToList(pFolderFile(mFolderFileCounter).FullName, pFolderFileList, pDirectoryLevel)

            If pFolderFileType = FileSystemType.Folder Then

                '再帰処理（サブディレクトリ内の処理）
                Call _GetFolderFileList(pFolderFile(mFolderFileCounter).FullName, pFolderFileList)

            End If

        Next

    End Sub

    ''' <summary>
    '''   対象パス情報をフォルダファイルリストへセット
    ''' </summary>
    ''' <param name="pTargetPath">対象パス</param>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <param name="pDirectoryLevel">ディレクトリレベル</param>
    ''' <remarks>
    '''   ※出力用の文字列情報はここではセットしない
    ''' </remarks>
    Private Sub _TargetPathToList(ByVal pTargetPath As String, ByVal pFolderFileList As DataTable, ByVal pDirectoryLevel As Integer)

        Try

            Dim mFolderFile As New System.IO.FileInfo(pTargetPath)

            '------------------------------
            ' ファイルシステムタイプを取得
            '------------------------------
            Dim mFileSystemType As FileSystemType = _GetFileSystemType(mFolderFile.FullName)
            Dim mFileSystemTypeName As String = mFileSystemType.ToString

            '------------------------------
            ' ファイル名を取得
            '------------------------------
            Dim mFileName As String = String.Empty
            If mFileSystemType = FileSystemType.Drive Then

                'ディレクトリのフルパスをセット
                mFileName = mFolderFile.FullName

            Else

                'ファイルの名前をセットする
                mFileName = mFolderFile.Name

            End If

            '------------------------------
            ' 拡張子、ファイルサイズを取得
            '------------------------------
            Dim mExtension As String = String.Empty
            Dim mFileSize As Long
            Select Case mFileSystemType

                Case FileSystemType.Drive, FileSystemType.Folder

                    '対象フォルダ以下のすべてのファイルサイズを合計した値を取得
                    mFileSize = _GetFolderSize(mFolderFile.FullName)

                Case FileSystemType.File

                    mFileSize = mFolderFile.Length

                    '拡張子は「.○○○」の形式で取得されるので「.」を削除した形で取得
                    mExtension = mFolderFile.Extension.Replace(".", "")

            End Select

            '------------------------------
            ' ファイルサイズレベルを取得
            '------------------------------
            Dim mFileSizeLevel As cFileSizeLevel.SizeLevel = cFileSizeLevel.GetLevel(mFileSize)

            '------------------------------
            ' 親フォルダ情報を取得
            '------------------------------
            Dim mParentFolder As String = String.Empty
            Dim mParentFolderFullPath As String = String.Empty

            '親ディレクトリのインスタンスが存在する時
            If Not mFolderFile.Directory Is Nothing Then

                '親フォルダ名、親フォルダフルパスをセット
                mParentFolder = mFolderFile.Directory.Name
                mParentFolderFullPath = mFolderFile.Directory.FullName

            End If

            '------------------------------
            ' 新規行にデータをセット
            '------------------------------
            '新規行データを作成（空）
            Dim mAddDataRow As DataRow = pFolderFileList.NewRow

            '追加行にデータをセット
            mAddDataRow(FolderFileListColumn.No) = pFolderFileList.Rows.Count

            mAddDataRow(FolderFileListColumn.Name) = mFileName
            mAddDataRow(FolderFileListColumn.UpdateDate) = mFolderFile.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss") '※時間に前0がつかないため、ToStringで前0が付くようにする
            mAddDataRow(FolderFileListColumn.FileSystemType) = mFileSystemType
            mAddDataRow(FolderFileListColumn.FileSystemTypeName) = mFileSystemTypeName
            mAddDataRow(FolderFileListColumn.Extension) = mExtension
            mAddDataRow(FolderFileListColumn.Size) = mFileSize
            mAddDataRow(FolderFileListColumn.SizeAndUnit) = ByteUnit.GetByteSizeAndUnit(mFileSize)
            mAddDataRow(FolderFileListColumn.SizeLevel) = mFileSizeLevel
            mAddDataRow(FolderFileListColumn.SizeLevelName) = mFileSizeLevel.ToString
            mAddDataRow(FolderFileListColumn.DirectoryLevel) = pDirectoryLevel
            mAddDataRow(FolderFileListColumn.ParentFolder) = mParentFolder
            mAddDataRow(FolderFileListColumn.ParentFolderFullPath) = mParentFolderFullPath
            mAddDataRow(FolderFileListColumn.UnderTargetFolder) = _GetUnderTargetPath(mFolderFile.FullName)
            mAddDataRow(FolderFileListColumn.FullPath) = mFolderFile.FullName
            mAddDataRow(FolderFileListColumn.IsLastFileInFolder) = False                                             '※最後のファイルかどうかは後でセットするためここではデフォルト値のFlaseをセットする

            '行データをDataTableにセット
            pFolderFileList.Rows.Add(mAddDataRow)

            'フォルダファイルリスト作業進捗報告 ※処理進捗報告を行うときのみ実行
            If _IsReportProcessProgress Then _ReportCreateFolderFileList(mFolderFile.Name)

        Catch ex As PathTooLongException

            Throw ex

        Catch ex As DirectoryNotFoundException

            Throw ex

        End Try

    End Sub

#End Region

#Region "フォルダ内で最後のファイルかどうかをセットする"

    ''' <summary>
    '''   フォルダファイルリストデータ内の最後のファイルかどうかに値をセットする
    ''' </summary>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <remarks></remarks>
    Public Sub SetIsLastFileInFolder(ByVal pFolderFileList As DataTable)

        'フォルダファイルリストデータで無い時は処理を終了
        If IsFolderFileListData(pFolderFileList) = False Then Exit Sub

        '「最後のファイルかどうか」にFalseを一括でセット
        pFolderFileList.AsEnumerable _
                       .Select(Function(UpdateRow As DataRow) InlineAssignHelper(UpdateRow(FolderFileListColumn.IsLastFileInFolder), False)).ToList()

        Dim mQuery = pFolderFileList.AsEnumerable() _
                     .Where(Function(ExclusionRow As DataRow) ExclusionRow(FolderFileListColumn.No) <> 0) _
                     .GroupBy(Function(FolderFileListRow As DataRow) FolderFileListRow(FolderFileListColumn.ParentFolderFullPath)) _
                     .Select(Function(ParentFolderGroupBy) New With {
                                                                         Key .No = ParentFolderGroupBy.Max(Function(ParentFolder) ParentFolder(FolderFileListColumn.No)),
                                                                         Key .PanretFolder = ParentFolderGroupBy.Key
                                                                     })

        '「最後のファイルかどうか」にTrueをセットし終わった数を保持する変数
        Dim mProcessCount As Integer = 1

        'フォルダ内で最後のファイルであるデータ分繰り返す
        For Each mDr In mQuery

            '対象「No」の値に「最後のファイルかどうか」にTrueをセット
            pFolderFileList.AsEnumerable() _
                           .Where(Function(WhereRow As DataRow) WhereRow(FolderFileListColumn.No) = mDr.No) _
                           .Select(Function(UpdateRow As DataRow) InlineAssignHelper(Of Boolean)(UpdateRow(FolderFileListColumn.IsLastFileInFolder), True)).ToList()

            '「最後のファイルかどうか」にTrueをセットする処理の作業進捗報告 ※処理進捗報告を行うときのみ実行
            If _IsReportProcessProgress Then _ReportSetIsLastFileInFolder(mQuery.Count, mProcessCount, mDr.PanretFolder)

            '「最後のファイルかどうか」にTrueをセットし終わった数を更新
            mProcessCount = mProcessCount + 1

        Next

    End Sub

    ''' <summary>
    '''   Linq内で値を更新するようメソッド
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="pTarget">更新対象の値</param>
    ''' <param name="value">更新後の値</param>
    ''' <returns>更新後の値</returns>
    ''' <remarks></remarks>
    Private Function InlineAssignHelper(Of T)(ByRef pTarget As T, value As T) As T

        pTarget = value

        Return value

    End Function

#End Region

#Region "出力用の文字列を作成"

    ''' <summary>
    '''   出力用の文字列を作成する
    ''' </summary>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <remarks>
    '''   画面に出力した際にフォルダ階層がひと目でわかるような文字列を作成していく
    '''   ※引数が存在しない場合は、指定されたパスのフォルダファイルリストで出力用の文字列を作成する
    ''' </remarks>
    Public Sub CreateOutputTextString(Optional ByVal pFolderFileList As DataTable = Nothing)

        '引数の「フォルダファイルリストがNothing」または「フォルダファイルリストデータで無かった」時
        If pFolderFileList Is Nothing OrElse Not IsFolderFileListData(pFolderFileList) Then

            '指定されたパスのフォルダファイルリストをセット
            pFolderFileList = _FolderFileList

        End If

        '「｜」を出力するかのフラグ
        Dim mNotPrintVerticalBar() As Boolean

        'フォルダファイルリストの行数分繰り返す
        For Each mFolderFileListRow As DataRow In pFolderFileList.Rows

            Dim mDispString As New System.Text.StringBuilder                                                   '出力用文字列
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

#Region "出力用の文字列を取得"

    ''' <summary>
    '''   出力用文字列を取得
    ''' </summary>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <returns>フォルダファイルリストの出力用文字列</returns>
    ''' <remarks></remarks>
    Public Function GetOutputTextString(ByVal pFolderFileList As DataTable) As String

        'フォルダファイルリストデータで無い時は「String.Empty」を返す
        If IsFolderFileListData(pFolderFileList) = False Then Return String.Empty

        'すべての行を取得し出力用の文字列を取得するQueryを作成（ラムダ式）
        Dim mQuery = pFolderFileList.AsEnumerable() _
                     .Select(Function(FolderFileListRow As DataRow) New With _
                                     { _
                                         .DispString = FolderFileListRow.Field(Of String)(FolderFileListColumn.DispString) _
                                     } _
                            )

        '出力用文字列を結合していく
        Dim mOutputText As New System.Text.StringBuilder
        For Each mRowData In mQuery

            mOutputText.AppendLine(mRowData.DispString)

        Next

        Return mOutputText.ToString

    End Function

#End Region

#Region "指定カラムで並び替えたフォルダファイルリストを取得"

    ''' <summary>
    '''   指定カラムで並び替えたフォルダファイルリストを取得
    ''' </summary>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <param name="pColumn">対象カラム</param>
    ''' <param name="pSortOrder">ソート条件</param>
    ''' <returns>対象カラムをソート条件で並び替えたフォルダファイルリスト</returns>
    ''' <remarks>
    '''   pFolderFileListがフォルダファイルリストデータで無かった場合はNothingを返す
    ''' </remarks>
    Public Function GetSpecifiedColumnFolderFileList(ByVal pFolderFileList As DataTable, ByVal pColumn As FolderFileListColumn, ByVal pSortOrder As SortOrder) As DataTable

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

    ''' <summary>
    '''   指定カラムで並び替えたフォルダファイルリストを取得
    ''' </summary>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <param name="pColumn">対象カラム</param>
    ''' <param name="pSortOrder">ソート条件</param>
    ''' <returns>対象カラムをソート条件で並び替えたフォルダファイルリスト</returns>
    ''' <remarks>
    '''   pFolderFileListがフォルダファイルリストデータで無かった場合はNothingを返す
    ''' </remarks>
    Public Function GetSpecifiedColumnFolderFileList(ByVal pFolderFileList As DataTable, ByVal pColumn As FolderFileListColumn, ByVal pSortOrder As FolderFileListSortOrder) As DataTable

        'フォルダファイルリストデータで無い時はNothingを返す
        If IsFolderFileListData(pFolderFileList) = False Then Return Nothing

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

    ''' <summary>
    '''   フォルダファイルリスト絞り込みデータを取得
    ''' </summary>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <param name="pColumn">フォルダファイルリストカラム</param>
    ''' <param name="pWhereString">絞り込み条件</param>
    ''' <returns>フォルダファイルリストを絞り込み条件で絞り込んだデータ</returns>
    ''' <remarks></remarks>
    Public Function GetRefineByFolderFileList(ByVal pFolderFileList As DataTable, ByVal pColumn As FolderFileListColumn, ByVal pWhereString As String) As DataTable

        'フォルダファイルリストデータで無い時はNothingを返す
        If IsFolderFileListData(pFolderFileList) = False Then Return Nothing

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

    ''' <summary>
    '''   フォルダファイルリスト絞り込みデータを取得 ファイルのみの絞り込みバージョン
    ''' </summary>
    ''' <param name="pFolderFileList">フォルダファイルリスト</param>
    ''' <param name="pColumn">フォルダファイルリストカラム</param>
    ''' <param name="pWhereString">絞り込み条件</param>
    ''' <returns>フォルダファイルリストを絞り込み条件で絞り込んだデータ</returns>
    ''' <remarks></remarks>
    Public Function GetRefineByFolderFileListIncludeFolder(ByVal pFolderFileList As DataTable, ByVal pColumn As FolderFileListColumn, ByVal pWhereString As String) As DataTable

        'フォルダファイルリストデータで無い時はNothingを返す
        If IsFolderFileListData(pFolderFileList) = False Then Return Nothing

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
                             Where RefineByFolderFileList(pColumn).ToString.ToLower().Contains(mLowerWhere) _
                                Or RefineByFolderFileList(FolderFileListColumn.FileSystemType) <> FileSystemType.File
                             Order By RefineByFolderFileList(FolderFileListColumn.No) Ascending

                'クエリーからDataViewを作成
                Dim mQueryToDataView As DataView = mQuery.AsDataView

                Return mQueryToDataView.ToTable

            Case FolderFileListColumn.Extension

                '引数で渡ってきたリストから拡張子で条件を絞り込んでデータを取得するクエリーを作成
                Dim mQuery = From RefineByFolderFileList In pFolderFileList.AsEnumerable
                             Where RefineByFolderFileList(pColumn).ToString.ToLower().ToString = mLowerWhere _
                                Or RefineByFolderFileList(FolderFileListColumn.FileSystemType) <> FileSystemType.File
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
                                 Where RefineByFolderFileList(pColumn).ToString = mLowerWhere _
                                Or RefineByFolderFileList(FolderFileListColumn.FileSystemType) <> FileSystemType.File
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

    ''' <summary>
    '''   フォルダファイルリストのファイル数のセットの進捗報告
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub _ReportSetFileCount()

        'フォルダファイルリストの進捗状況を報告する変数がNothingで無かった時
        If Not _ProcessProgresss Is Nothing Then

            '進捗率を報告
            _ProcessProgresss.Report(New FolderFileListProgress(0, "", FolderFileListProcessState.SetFileCount))

        End If

    End Sub

    ''' <summary>
    '''   フォルダファイルリスト作成の進捗報告
    ''' </summary>
    ''' <param name="pPath">処理フォルダファイル</param>
    ''' <remarks></remarks>
    Private Sub _ReportCreateFolderFileList(ByVal pPath As String)

        'フォルダファイルリストの進捗状況を報告する変数がNothingで無かった時
        If Not _ProcessProgresss Is Nothing Then

            '現在の処理進捗率を計算
            Dim mProcessPercent As Integer = (_FolderFileList.Rows.Count / _FolderFileListCount) * 100

            '進捗率を報告
            _ProcessProgresss.Report(New FolderFileListProgress(mProcessPercent, pPath, FolderFileListProcessState.CreateFolderFileList))

        End If

    End Sub

    ''' <summary>
    '''   「最後のファイルかどうか」にTrueをセットする処理の作業進捗
    ''' </summary>
    ''' <param name="pTargetFileCount">処理対象数</param>
    ''' <param name="pProcessCount">処理済み数</param>
    ''' <param name="pFolderPath">対象フォルダパス</param>
    ''' <remarks></remarks>
    Private Sub _ReportSetIsLastFileInFolder(ByVal pTargetFileCount As Integer, ByVal pProcessCount As Integer, ByVal pFolderPath As String)

        'フォルダファイルリストの進捗状況を報告する変数がNothingで無かった時
        If Not _ProcessProgresss Is Nothing Then

            '現在の処理進捗率を計算
            Dim mProcessPercent As Integer = (pProcessCount / pTargetFileCount) * 100

            '進捗率を報告
            _ProcessProgresss.Report(New FolderFileListProgress(mProcessPercent, pFolderPath, FolderFileListProcessState.SetIsLastFileInFolder))

        End If

    End Sub

#End Region

#Region "フォルダファイルリストデータかどうか"

    ''' <summary>
    '''   フォルダファイルリストデータかどうか
    ''' </summary>
    ''' <param name="pDt">チェック対象DataTable</param>
    ''' <returns>True：フォルダファイルリストデータ、False：フォルダファイルリストデータでない</returns>
    ''' <remarks>
    '''   カラム数が同じかつカラム名がすべて一致し場合はフォルダファイルリストデータと判断する
    '''   ※カラムの位置が違うが同じカラム名が存在する時は同じと判断されるので注意
    ''' </remarks>
    Public Function IsFolderFileListData(ByVal pDt As DataTable) As Boolean

        '------------------
        ' カラム数チェック
        '------------------
        'DataTableとフォルダファイルリストカラムの数が一致しない時は、フォルダファイルリストデータでないとする
        If pDt.Columns.Count <> System.Enum.GetNames(GetType(FolderFileListColumn)).Count Then Return False

        '------------------
        ' カラム名チェック
        '------------------
        'フォルダーファイルリストカラム列挙体分繰り返す
        For Each mColumnName As String In System.Enum.GetNames(GetType(FolderFileListColumn))

            Dim mIsMatchColumnName As Boolean = False

            'DataTableのカラム数分繰り返す
            For Each mDc As DataColumn In pDt.Columns

                'フォルダファイルリストカラムと一致するカラム名があった時は次の繰り返し
                If mDc.ColumnName = mColumnName Then

                    mIsMatchColumnName = True
                    Continue For

                End If

            Next

            '一致するカラムが無かった時はフォルダファイルリストデータでないとする
            If mIsMatchColumnName = False Then Return False

        Next

        Return True

    End Function

#End Region

#End Region

End Class

''' <summary>
'''   単位付きバイトサイズの機能を提供する
''' </summary>
''' <remarks></remarks>
Public Class ByteUnit

#Region "定数"

    ''' <summary>バイト単位計算用</summary>
    ''' <remarks>2の10乗</remarks>
    Private Const _cByteCalculationSize As Double = 1024

    ''' <summary>バイトサイズが上がるかどうか判断用定数</summary>
    ''' <remarks>値がこの数値以上だったら次のバイトサイズ名へ変更するための目安</remarks>
    Private Const _cIsUpByteSize As Double = 1001

    ''' <summary>四捨五入対象表示桁数</summary>
    ''' <remarks></remarks>
    Private Const _cRoundDigit As Integer = 2

#End Region

#Region "列挙体"

    ''' <summary>情報の単位</summary>
    ''' <remarks></remarks>
    Private Enum ByteSizeName

        ''' <summary>バイト</summary>
        B = 1

        ''' <summary>キロバイト</summary>
        KB

        ''' <summary>メガバイト</summary>
        MB

        ''' <summary>ギガバイト</summary>
        GB

        ''' <summary>テラバイト</summary>
        TB

    End Enum

#End Region

#Region "コンストラクタ"

    ''' <summary>
    '''   コンストラクタ
    ''' </summary>
    ''' <remarks>引数無しは外部に公開しない</remarks>
    Private Sub New()

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>
    '''   単位を付加したバイトサイズを取得
    ''' </summary>
    ''' <param name="pByte">ファイルサイズ</param>
    ''' <returns>単位を付加したバイトサイズ</returns>
    ''' <remarks>
    '''   ※テラバイト以上には対応していない
    ''' </remarks>
    Public Shared Function GetByteSizeAndUnit(ByVal pByte As Long) As String

        Dim mByte As Double = CType(pByte, Double)
        Dim mSizeName As ByteSizeName = ByteSizeName.B

        'バイトサイズが1001以上の時（上位サイズへ変換可能）、繰り返す
        Do While mByte >= _cIsUpByteSize

            '上位サイズへ変換後の値を計算
            mByte = mByte / _cByteCalculationSize

            'サイズ名を上位サイズへ
            mSizeName = mSizeName + 1

        Loop

        '四捨五入後のバイトサイズを取得
        mByte = Math.Round(mByte, _cRoundDigit)

        'バイト数＋サイズ名を返す
        Return mByte.ToString & mSizeName.ToString

    End Function

    ''' <summary>
    '''   単位を付加したバイトサイズをバイトサイズに変換する
    ''' </summary>
    ''' <param name="pByteSizeAndUnit">単位を付加したバイトサイズ</param>
    ''' <returns>バイトサイズ</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertSizeAndUnitToSize(ByVal pByteSizeAndUnit As String) As Long

        '単位の一覧を配列で取得し順番を逆に変更
        Dim mByteSizeNames As String() = System.Enum.GetNames(GetType(ByteSizeName))
        Array.Reverse(mByteSizeNames)

        '単位を選定する
        Dim mByteUnit As ByteSizeName = ByteSizeName.B
        For Each mByteSizeName As ByteSizeName In mByteSizeNames

            '単位を付加したバイトサイズに現在の単位文字列が含まれている時
            If pByteSizeAndUnit.Contains(mByteSizeName.ToString) Then

                mByteUnit = mByteSizeName
                Exit For

            End If

        Next

        '単位を付加したバイトサイズより数値の部分を取得
        Dim mByteSize As Long = CType(pByteSizeAndUnit.Replace(mByteUnit.ToString, ""), Long)

        '単位変換前の値に変換
        For i As Integer = 1 To mByteUnit

            mByteSize = mByteSize * _cByteCalculationSize

        Next

        Return mByteSize

    End Function

#End Region

End Class