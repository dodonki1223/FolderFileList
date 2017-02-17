Option Explicit On

''' <summary>単位付きバイトサイズの機能を提供する</summary>
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

    ''' <summary>コンストラクタ</summary>
    ''' <remarks>引数無しは外部に公開しない</remarks>
    Private Sub New()

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>単位を付加したバイトサイズを取得</summary>
    ''' <param name="pByte">ファイルサイズ</param>
    ''' <returns>単位を付加したバイトサイズ</returns>
    ''' <remarks>テラバイト以上には対応していない</remarks>
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

    ''' <summary>単位を付加したバイトサイズをバイトサイズに変換する</summary>
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
