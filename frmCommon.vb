Option Explicit On

#Region "インターフェース"

''' <summary>フォームの共通処理を定義</summary>
''' <remarks>インターフェースを使用したことがないからとりあえず実装してみる！！</remarks>
Interface IFormCommonProcess

    ''' <summary>コントロールのサイズ変更・再配置</summary>
    ''' <param name="pChangedSize">フォームの変更後サイズ</param>
    ''' <remarks>全てのコントロールを取得しコントロールごと処理を行う</remarks>
    Sub _ResizeAndRealignmentControls(ByVal pChangedSize As System.Drawing.Size)

    ''' <summary>初期画面コントロール設定</summary>
    ''' <remarks></remarks>
    Sub _SetInitialControlInScreen()

End Interface

#End Region

''' <summary>継承用のフォーム</summary>
''' <remarks>フォームの共通処理を記述</remarks>
Public Class frmCommon

#Region "定数"

    ''' <summary>バージョン番号</summary>
    ''' <remarks>Chromeのバージョン表記を参考に付ける
    '''            Chrome 5.0.307.11
    '''                5：メジャーバージョン
    '''                0：マイナーバージョン。この場合は、バージョン5の最初のもの
    '''              307：累積バージョン。ビルドバージョン。Chromeが出てから307番目に作られたもの
    '''               11：累積バージョンのマイナー番号（子番号）。リビジョン（修正版）
    ''' </remarks>
    Public Const cVersionNumber As String = "Ver2.0.0.0"

#End Region

#Region "列挙体"

    ''' <summary>Alt+Tabウインドウタイプ</summary>
    ''' <remarks></remarks>
    Public Enum AltTabType

        ''' <summary>表示</summary>
        Show

        ''' <summary>非表示</summary>
        Hide

    End Enum

#End Region

#Region "イベント"

    ''' <summary>フォームのロードイベント</summary>
    ''' <param name="sender">Formオブジェクト</param>
    ''' <param name="e">Loadイベント</param>
    Private Sub CommonForm_Load(sender As Object, e As EventArgs) Handles Me.Load

        'バージョン番号をフォームのテキストに表示
        Call SetVerNumber(Me)

        '全てのコントロールの最小サイズを設定
        Call SetControlsForMinimumSize()

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>バージョン番号をセット</summary>
    ''' <param name="pForm">対象フォーム</param>
    ''' <remarks></remarks>
    Public Sub SetVerNumber(ByVal pForm As Form)

        'フォームのテキストにバージョン番号を表記
        Me.Text = Me.Text & "（" & cVersionNumber & "）"

    End Sub

    ''' <summary>対象コントロール内のコントロールを取得する</summary>
    ''' <param name="pTopControl">対象コントロール</param>
    ''' <returns>対象コントロール内のコントロールを全て格納したArrayList</returns>
    ''' <remarks>再帰処理を行いコントロールを取得していく</remarks>
    Public Function GetControlsInTarget(ByVal pTopControl As Control) As Control()

        Dim mAllControls As ArrayList = New ArrayList

        'pTopControl配下のコントロールを追加していく
        For Each mUnderControl As Control In pTopControl.Controls

            mAllControls.Add(mUnderControl)

            'mUnderControl配下のコントロールも追加していく（再帰処理）
            mAllControls.AddRange(GetControlsInTarget(mUnderControl))

        Next

        Return CType(mAllControls.ToArray(GetType(Control)), Control())

    End Function

    ''' <summary>コントロールの最小サイズ設定</summary>
    ''' <remarks></remarks>
    Public Sub SetControlsForMinimumSize()

        'Formの最小サイズを設定
        Me.MinimumSize = Me.Size

        'Form以外の全てのコントロールを取得
        Dim mAllControls As Control() = GetControlsInTarget(Me)

        For Each mTargetControl As Control In mAllControls

            'Form以外のコントロールの最小サイズを設定していく
            mTargetControl.MinimumSize = mTargetControl.Size

        Next

    End Sub

    ''' <summary>フォームをAlt+Tabに表示・非表示</summary>
    ''' <param name="pAltTabType">Alt+Tabに表示させるかどうかの区分</param>
    ''' <remarks>Alt+Tabに表示させるかどうか設定する</remarks>
    Public Sub SetShowHideAltTabWindow(ByVal pAltTabType As AltTabType)

        Select Case pAltTabType

            Case AltTabType.Show

                'Alt+Tabに表示させる
                Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
                Me.ShowInTaskbar = True

            Case AltTabType.Hide

                'Alt+Tabに表示させない
                Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
                Me.ShowInTaskbar = False

        End Select

    End Sub

    ''' <summary>ユーザーに対話メッセージボックス</summary>
    ''' <param name="pFileFormat">出力形式</param>
    ''' <returns>ダイアログボックスの戻り値</returns>
    ''' <remarks></remarks>
    Public Function ShowDialogueMessage(ByVal pFileFormat As String) As Windows.Forms.DialogResult

        'メッセージボックスを表示しユーザーに対話
        Dim mMsgBoxTitle As String = pFileFormat & "出力"
        Dim mMsgBoxText As String = "はい　：" & pFileFormat & "ファイルを保存します" & System.Environment.NewLine &
                                    "いいえ：クリップボードに" & pFileFormat & "形式のテキストを保存します"

        Return MessageBox.Show(mMsgBoxText, mMsgBoxTitle, MessageBoxButtons.YesNo)

    End Function

    ''' <summary>名前を付けて保存ダイアログを取得</summary>
    ''' <param name="pSaveFileName">保存ファイル名</param>
    ''' <param name="pFileFormat">ファイル形式</param>
    ''' <param name="pDirectory">起動ディレクトリ</param>
    ''' <returns>ファイル形式にあった設定の名前を付けて保存ダイアログ</returns>
    ''' <remarks>ファイル形式用の名前を付けて保存ダイアログを作成し返す</remarks>
    Public Function GetSaveAsDialog(ByVal pSaveFileName As String, ByVal pFileFormat As String, Optional ByVal pDirectory As String = "") As SaveFileDialog

        '大文字、小文字の区切り文字を取得
        Dim mFileFormatUpperCase As String = pFileFormat.ToUpper
        Dim mFileFormatLowerCase As String = pFileFormat.ToLower

        '名前を付けて保存ダイアログを表示
        Dim mDailog As New SaveFileDialog
        With mDailog

            'デフォルトファイル設定（保存ファイル名＋.区切り文字）
            .FileName = pSaveFileName & "." & mFileFormatLowerCase

            '表示ファイル設定
            .Filter = mFileFormatUpperCase & "ファイル(*." & mFileFormatLowerCase & ")|*." & mFileFormatLowerCase & "|すべてのファイル(*.*)|*.*"

            '起動ディレクトリ設定
            .InitialDirectory = pDirectory

        End With

        Return mDailog

    End Function

    ''' <summary>文字列を指定ファイルに書き込む</summary>
    ''' <param name="pWriteText">書き込むテキスト</param>
    ''' <param name="pOutputPath">指定ファイル（出力先フルパス）</param>
    ''' <param name="pEncoding">使用する文字エンコーディング</param>
    ''' <remarks>指定されたファイルが存在しない場合はファイルを作成して書き込む
    '''          指定されたファイルが存在する場合はファイルを上書きして書き込む</remarks>
    Public Sub WriteTextToOutputFile(ByVal pWriteText As String, ByVal pOutputPath As String, ByVal pEncoding As System.Text.Encoding)

        Using mSW As New System.IO.StreamWriter(pOutputPath, False, pEncoding)

            mSW.Write(pWriteText)

        End Using

    End Sub

    ''' <summary>ファイルの実行処理</summary>
    ''' <param name="pPath">実行ファイルのフルパス</param>
    ''' <remarks></remarks>
    Public Sub RunFile(ByVal pPath As String)

        Dim mPsi As New System.Diagnostics.ProcessStartInfo()

        '関連付けで実行するファイルのフルパスを指定
        mPsi.FileName = pPath

        'ファイルの実行処理
        System.Diagnostics.Process.Start(mPsi)

    End Sub

    ''' <summary>Messenger風通知メッセージを表示するウインドウを表示</summary>
    ''' <param name="pTitle">タイトル</param>
    ''' <param name="pMessage">メッセージ</param>
    Public Sub ShowPopupMessage(ByVal pTitle As String, ByVal pMessage As String)

        'Messenger風通知メッセージを表示するウインドウのインスタンスを作成
        Dim mFrmPopupMessage As New frmPopupMessage(pTitle, pMessage)

        'Messenger風通知メッセージを表示
        mFrmPopupMessage.Show()

    End Sub

#End Region

End Class