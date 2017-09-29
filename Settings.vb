Option Explicit On

Imports System
Imports System.IO
Imports System.Text
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Windows.Forms
Imports Microsoft.Win32

''' <summary>設定を読み書きする機能を提供します</summary>
''' <remarks></remarks>
<Serializable()> _
Public Class Settings

#Region "変数"

    ''' <summary>設定を読み書きする機能を提供するクラス</summary>
    ''' <remarks>インスタンスは１つのみ作られるようにしている</remarks>
    <NonSerialized()> _
    Private Shared _Settings As Settings

    ''' <summary>フォームの種類</summary>
    ''' <remarks></remarks>
    Private _TargetForm As CommandLine.FormType

    ''' <summary>実行タイプ</summary>
    ''' <remarks>保存したファイルを即時で実行するかどうか</remarks>
    Private _Execute As CommandLine.ExecuteType

    ''' <summary>１ページ内表示最大件数（リスト表示）</summary>
    ''' <remarks></remarks>
    Private _MaxCountInPage As Integer

#End Region

#Region "プロパティ"

    ''' <summary>ターゲットフォームプロパティ</summary>
    ''' <remarks>※シリアライズを行うプロパティ</remarks>
    Public Property TargetForm() As CommandLine.FormType

        Get

            Return _TargetForm

        End Get

        Set(value As CommandLine.FormType)

            _TargetForm = value

        End Set

    End Property

    ''' <summary>実行タイププロパティ</summary>
    ''' <remarks>保存したファイルを即時で実行するかどうか
    '''          ※シリアライズを行うプロパティ                          </remarks>
    Public Property Execute() As CommandLine.ExecuteType

        Get

            Return _Execute

        End Get

        Set(value As CommandLine.ExecuteType)

            _Execute = value

        End Set

    End Property

    ''' <summary>1ページ内に表示できるファイル数プロパティ</summary>
    ''' <remarks>リスト表示フォームの１ページ内に表示できる最大ファイル数
    '''          ※シリアライズを行うプロパティ                          </remarks>
    Public Property MaxCountInPage() As Integer

        Get

            Return _MaxCountInPage

        End Get

        Set(value As Integer)

            _MaxCountInPage = value

        End Set

    End Property

    ''' <summary>設定を読み書きするクラスを返却します</summary>
    ''' <remarks>※デザインパターンのSingletonパターンです
    '''            シリアライズを行わないプロパティ       </remarks>
    <System.Xml.Serialization.XmlIgnore()> _
    Public Shared Property Instance() As Settings

        Get

            'インスタンスが作成されてなかったらインスタンスを作成
            If _Settings Is Nothing Then _Settings = New Settings

            Return _Settings

        End Get

        Set(value As Settings)

            _Settings = value

        End Set

    End Property

#End Region

#Region "コンストラクタ"

    ''' <summary>コンストラクタ</summary>
    ''' <remarks>外部に公開しない</remarks>
    Private Sub New()

        'ターゲットフォームのデフォルトはフォルダファイルリストの出力文字列を表示するフォーム
        _TargetForm = CommandLine.FormType.Text

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>設定をXMLファイルから読み込み復元する</summary>
    ''' <remarks></remarks>
    Public Shared Sub LoadFromXmlFile()

        '設定ファイルのパスを取得
        Dim mSettingPath As String = GetSettingPath()

        '設定ファイルが存在しない時は処理を終了
        If Not File.Exists(mSettingPath) Then Exit Sub

        '設定ファイルを読み込み
        Dim mSR As New StreamReader(mSettingPath, New UTF8Encoding(False))

        '設定ファイルを読み込んでデシリアライズする（オブジェクトからXML形式に変換）
        Dim mXS As New System.Xml.Serialization.XmlSerializer(GetType(Settings))
        Dim mDeserializeXML As Object = mXS.Deserialize(mSR)

        mSR.Close()

        Instance = CType(mDeserializeXML, Settings)

    End Sub

    ''' <summary>現在の設定をXMLファイルに書き込む</summary>
    ''' <remarks></remarks>
    Public Shared Sub SaveToXmlFile()

        '設定ファイルのパスを取得
        Dim mSettingPath As String = GetSettingPath()

        '書き込み対象のXMLファイルのTextWriterを作成
        Dim mSW As New StreamWriter(mSettingPath, False, New UTF8Encoding(False))
        Dim mSerializeXML As New System.Xml.Serialization.XmlSerializer(GetType(Settings))

        'シリアライズ（オブジェクトからXML形式に変換）し、設定ファイルに書き込む
        mSerializeXML.Serialize(mSW, Instance)

        mSW.Close()

    End Sub

    ''' <summary>設定ファイルの保存場所を取得</summary>
    ''' <returns>設定ファイルの保存場所パス</returns>
    ''' <remarks>「EXE実行パスの親フォルダ + \ + Config.xml」形式のパスを返す</remarks>
    Private Shared Function GetSettingPath() As String

        'EXEの実行パスを取得
        Dim mExePath As New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)

        '設定ファイル保存パス（EXE実行パスの親フォルダ + \ + Config.xml）を返す
        Return mExePath.DirectoryName & "\" & "Config.xml"

    End Function

#End Region

End Class