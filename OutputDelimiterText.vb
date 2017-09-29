Option Explicit On

Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Text

''' <summary>
'''   出力テキストの機能を提供する
''' </summary>
''' <remarks></remarks>
Public Class OutputDelimiterText

#Region "定数"

    ''' <summary>出力用ファイルの区切り文字を提供する</summary>
    ''' <remarks></remarks>
    Public Class cDelimiters

        ''' <summary>CSV用区切り文字</summary>
        Public Const CSV As String = ","

        ''' <summary>TSV用区切り文字</summary>
        Public Const TSV As String = ControlChars.Tab

        ''' <summary>SPACE用区切り文字</summary>
        Public Const SPACE As String = " "

    End Class

#End Region

#Region "列挙体"

    ''' <summary>区切り文字</summary>
    ''' <remarks></remarks>
    Public Enum Delimiter

        ''' <summary>CSV</summary>
        ''' <remarks>「,」カンマ区切り</remarks>
        CSV

        ''' <summary>TSV</summary>
        ''' <remarks>「TAB」タブ区切り</remarks>
        TSV

    End Enum

#End Region

#Region "変数"

    ''' <summary>出力文字列</summary>
    ''' <remarks>クラス内で使用する変数</remarks>
    Private _OutputText As String

    ''' <summary>ヘッダー文字列</summary>
    ''' <remarks>クラス内で使用する変数</remarks>
    Private _HeaderText As String

    ''' <summary>出力文字列の前に出力する文字列</summary>
    ''' <remarks>クラス内で使用する変数</remarks>
    Private _BeforeOutputText As String

    ''' <summary>出力文字列の後に出力する文字列</summary>
    ''' <remarks>クラス内で使用する変数</remarks>
    Private _AfterOutputText As String

#End Region

#Region "プロパティ"

    ''' <summary>
    '''   出力文字列プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public Property OutputText As String
        Get

            Return _OutputText

        End Get
        Set(value As String)

            _OutputText = _AddString(_OutputText, value)

        End Set
    End Property

    ''' <summary>
    '''   ヘッダープロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public Property HeaderText As String
        Get

            Return _HeaderText

        End Get
        Set(value As String)

            _HeaderText = _AddString(_HeaderText, value)

        End Set
    End Property

    ''' <summary>
    '''   出力文字列前に出力する文字列プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public Property BeforeOutputText As String
        Get

            Return _BeforeOutputText

        End Get
        Set(value As String)

            _BeforeOutputText = _AddString(_BeforeOutputText, value)

        End Set
    End Property

    ''' <summary>
    '''   出力文字列後に出力する文字列プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public Property AfterOutputText As String
        Get

            Return _AfterOutputText

        End Get
        Set(value As String)

            _AfterOutputText = _AddString(_AfterOutputText, value)

        End Set
    End Property

    ''' <summary>
    '''   最終的な出力文字列プロパティ
    ''' </summary>
    ''' <remarks>  
    '''     出力文字列前に出力する文字列
    '''   ＋ヘッダー文字列
    '''   ＋出力文字列
    '''   ＋出力文字列後に出力する文字列
    ''' </remarks>
    Public ReadOnly Property AllOutputText As String
        Get

            Dim mAllOutputText As New StringBuilder
            If Not String.IsNullOrEmpty(_BeforeOutputText) Then mAllOutputText.AppendLine(_BeforeOutputText)
            If Not String.IsNullOrEmpty(_HeaderText) Then mAllOutputText.AppendLine(_HeaderText)
            If Not String.IsNullOrEmpty(_OutputText) Then mAllOutputText.AppendLine(_OutputText)
            If Not String.IsNullOrEmpty(_AfterOutputText) Then mAllOutputText.AppendLine(_AfterOutputText)

            Return mAllOutputText.ToString

        End Get
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
    ''' <param name="pData">対象のDataTable</param>
    ''' <param name="pDelimiter">区切り文字</param>
    ''' <remarks>
    '''   引数付きのコンストラクタのみを公開
    ''' </remarks>
    Public Sub New(ByVal pData As DataTable, ByVal pDelimiter As Delimiter)

        '区切り文字を取得
        Dim mDelimiterString As String = _GetDelimiterString(pDelimiter)

        '対象DataTtableの値を区切り文字で区切ったテキストデータに変換
        _OutputText = _ConvertDataTableToOutputText(pData, mDelimiterString)

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>
    '''   DataTableのデータを出力テキストに変換する
    ''' </summary>
    ''' <param name="pTargetDt">対象のDataTable</param>
    ''' <param name="pDelimiter">区切り文字</param>
    ''' <returns>区切り文字で区切られた文字列</returns>
    ''' <remarks></remarks>
    Private Function _ConvertDataTableToOutputText(ByVal pTargetDt As DataTable, ByVal pDelimiter As String) As String

        Dim mOutputText As New StringBuilder

        'DataTableの行数分繰り返す
        For Each dr As DataRow In pTargetDt.Rows

            Dim mRowWriteString As String = String.Empty

            'DataTableの列数分繰り返す
            For Each dc As DataColumn In pTargetDt.Columns

                '対象カラムデータを追加していく
                mRowWriteString = _GetOutputTextForColumnData(mRowWriteString, pDelimiter, dr(dc.ColumnName).ToString)

            Next

            mOutputText.AppendLine(mRowWriteString)

        Next

        '末尾の改行文字を除いた文字列を返す
        Return System.Text.RegularExpressions.Regex.Replace(mOutputText.ToString, "[\r\n]+$", "")

    End Function

    ''' <summary>
    '''   対象カラム出力テキストデータを取得
    ''' </summary>
    ''' <param name="pRowWriteString">１行分の出力テキストデータ</param>
    ''' <param name="pDelimiter">区切り文字</param>
    ''' <param name="pWriteString">対象カラムデータ</param>
    ''' <returns>対象カラムデータを追加して返す</returns>
    ''' <remarks>
    '''   １番始めのカラムデータの時はそのまま返す
    '''   １番始め以外は区切り文字を追加して返す
    ''' </remarks>
    Private Function _GetOutputTextForColumnData(ByVal pRowWriteString As String, ByVal pDelimiter As String, ByVal pWriteString As String) As String

        '1行分の文字列を格納する変数が空だったら
        If String.IsNullOrEmpty(pRowWriteString) Then

            'ダブルクォーテーションで囲み値をそのまま返す
            Return EncloseDoubleQuotes(pWriteString)

        Else

            '対象の文字列をダブルクォーテーションで囲み区切り文字を追加して値を返す
            Return pRowWriteString & pDelimiter & EncloseDoubleQuotes(pWriteString)

        End If

    End Function

    ''' <summary>
    '''   区切り文字列を取得
    ''' </summary>
    ''' <param name="pDelimiter">区切り文字列挙体</param>
    ''' <returns>区切り文字に一致する文字列</returns>
    ''' <remarks>
    '''   区切り文字列挙体に一致する文字列が無かった場合は空文字を返す
    ''' </remarks>
    Private Function _GetDelimiterString(ByVal pDelimiter As Delimiter) As String

        Select Case pDelimiter

            Case Delimiter.CSV

                Return cDelimiters.CSV

            Case Delimiter.TSV

                Return cDelimiters.TSV

            Case Else

                Return ""

        End Select

    End Function

    ''' <summary>
    '''   対象文字列に追加文字列を追加
    ''' </summary>
    ''' <param name="pTargetText">対象文字列</param>
    ''' <param name="pAddString">文字列の追加</param>
    ''' <returns>対象文字列に追加後の文字列</returns>
    ''' <remarks></remarks>
    Private Function _AddString(ByVal pTargetText As String, ByVal pAddString As String) As String

        Dim mAfterAddString As New System.Text.StringBuilder

        '対象の文字列が空で無かった時、文字列を追加
        If Not String.IsNullOrEmpty(pTargetText) Then mAfterAddString.AppendLine(pTargetText)
        mAfterAddString.Append(pAddString)

        Return mAfterAddString.ToString

    End Function

    ''' <summary>
    '''   配列から１行分のデータを作成する
    ''' </summary>
    ''' <param name="pArrayData">対象配列</param>
    ''' <param name="pDelimiter">区切り文字</param>
    ''' <returns>配列の文字列を区切り文字で区切って返す</returns>
    ''' <remarks></remarks>
    Public Function ConvertArrayToDelimiterRowData(ByVal pArrayData As String(), ByVal pDelimiter As Delimiter) As String

        '区切り文字を取得
        Dim mDelimiter As String = _GetDelimiterString(pDelimiter)

        Dim mRowData As String = String.Empty
        For Each mColumnData As String In pArrayData

            mRowData = _GetOutputTextForColumnData(mRowData, mDelimiter, mColumnData)

        Next

        Return mRowData

    End Function

#End Region

#Region "外部公開メソッド"

    ''' <summary>
    '''   文字列をダブルクォーテションで囲む
    ''' </summary>
    ''' <param name="pField">対象文字列</param>
    ''' <returns>ダブルクォーテーションで囲んだ文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function EncloseDoubleQuotes(pField As String) As String

        If pField.IndexOf(""""c) > -1 Then

            '「"」を「""」とする
            pField = pField.Replace("""", """""")

        End If

        Return """" & pField & """"

    End Function

#End Region

End Class