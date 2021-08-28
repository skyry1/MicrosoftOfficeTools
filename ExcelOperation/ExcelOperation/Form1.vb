Imports System.IO
Imports System.Text

Public Class Form1

    Private ReadOnly PropertyFileName As String = "ExcelOperation.properties"
    Private ReadOnly DirPath_PropertiesName As String = "実行ディレクトリ："
    Private ReadOnly ActiveCell_PropertiesName As String = "アクティブセル："
    Private ReadOnly SheetZoom_PropertiesName As String = "表示倍率："
    Private ReadOnly Filter_PropertiesName As String = "フィルター解除："

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReadProperties()
    End Sub


    ''' <summary>
    ''' 実行ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ProgressBar.Minimum = 0

        Try
            If DirPath_TextBox.Text.Equals("") Then
                Exit Try
            End If

            Dim di As New IO.DirectoryInfo(DirPath_TextBox.Text)
            Dim files As IO.FileInfo() = di.GetFiles("*.*", IO.SearchOption.AllDirectories)
            ProgressBar.Maximum = files.Length
            For Each f As IO.FileInfo In files
                If f.Extension.Equals(".xlsx") Or f.Extension.Equals(".xlsm") Or f.Extension.Equals(".xls") Then
                    SheetOperation.bookName = f.FullName
                    SheetOperation.activeCell = ActiveCell_TextBox.Text
                    SheetOperation.zoom = Integer.Parse(SheetZoom_TextBox.Text)
                    SheetOperation.filterLift = Filter_CheckBox.Checked
                    SheetOperation.ActiveCellSet()
                End If
                ProgressBar.Value = ProgressBar.Value + 1
            Next
        Catch ex As Exception
            Logger.WriteTraceLog("ERROR", ex.Message.ToString)
            Throw
        Finally
            MsgBox("処理が完了しました。")
            ProgressBar.Value = ProgressBar.Minimum
            WriteProperties()
        End Try
    End Sub


    ''' <summary>
    ''' 参照ボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'FolderBrowserDialogクラスのインスタンスを作成
        Dim fbd As New FolderBrowserDialog

        '上部に表示する説明テキストを指定する
        fbd.Description = "フォルダを指定してください。"
        'ルートフォルダを指定する
        'デフォルトでDesktop
        fbd.RootFolder = Environment.SpecialFolder.Desktop
        'ユーザーが新しいフォルダを作成できるようにする
        'デフォルトでTrue
        fbd.ShowNewFolderButton = True

        'ダイアログを表示する
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            DirPath_TextBox.Text = fbd.SelectedPath
        End If
    End Sub

    ''' <summary>
    ''' 前回実行履歴を読み込む
    ''' </summary>
    Private Sub ReadProperties()
        'Shift_JISを使用可能にする
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        Try
            Dim sr As New StreamReader(PropertyFileName, Encoding.GetEncoding("Shift_JIS"))
            Dim line As String = sr.ReadLine()
            Do Until line Is Nothing
                If InStr(line, DirPath_PropertiesName) <> 0 Then
                    DirPath_TextBox.Text = line.Replace(DirPath_PropertiesName, "")
                ElseIf InStr(line, ActiveCell_PropertiesName) <> 0 Then
                    ActiveCell_TextBox.Text = line.Replace(ActiveCell_PropertiesName, "")
                ElseIf InStr(line, SheetZoom_PropertiesName) <> 0 Then
                    SheetZoom_TextBox.Text = line.Replace(SheetZoom_PropertiesName, "")
                ElseIf InStr(line, Filter_PropertiesName) <> 0 Then
                    Filter_CheckBox.Checked = Boolean.Parse(line.Replace(Filter_PropertiesName, ""))
                End If
                line = sr.ReadLine()
            Loop
            sr.Close()
        Catch ex As Exception
            WriteProperties()
        Finally
        End Try
    End Sub

    ''' <summary>
    ''' 実行履歴を書き込む
    ''' </summary>
    Private Sub WriteProperties()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        Dim sw As New StreamWriter(PropertyFileName, False, Encoding.GetEncoding("Shift_JIS"))
        sw.WriteLine(DirPath_PropertiesName + DirPath_TextBox.Text)
        sw.WriteLine(ActiveCell_PropertiesName + ActiveCell_TextBox.Text)
        sw.WriteLine(SheetZoom_PropertiesName + SheetZoom_TextBox.Text)
        sw.WriteLine(Filter_PropertiesName + Filter_CheckBox.Checked.ToString)
        sw.Close()
    End Sub
End Class
