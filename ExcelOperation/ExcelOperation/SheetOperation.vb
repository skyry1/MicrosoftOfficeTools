Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Module SheetOperation
    ''' <summary>
    ''' 操作対象のExcelファイル
    ''' </summary>
    Public bookName As String
    Public activeCell As String = "A1"
    Public zoom As Integer = 100
    Public filterLift As Boolean = False

    Public Sub ActiveCellSet()

        '1シートずつ処理
        Dim excelApp As New Excel.Application()
        Dim excelBook = excelApp.Workbooks.Open(bookName)
        Logger.WriteTraceLog("INFO", excelBook.Name + "を開く")
        For Each sheet As Excel.Worksheet In excelBook.Sheets

            '非表示シートを表示にする
            Dim sheetVisible As Boolean = sheet.Visible
            If Not sheet.Visible Then
                Logger.WriteTraceLog("INFO", sheet.Name + "シートを非表示から表示に変更")
                sheet.Visible = False
            End If

            'アクティブシートに設定
            Logger.WriteTraceLog("INFO", sheet.Name + "シートをアクティブシートに変更")
            sheet.Activate()

            'アクティブセルを変更
            sheet.Range(activeCell).Select()
            Logger.WriteTraceLog("INFO", sheet.Name + "シートのアクティブセルを" + activeCell + "に変更")

            '表示倍率を変更
            excelApp.ActiveWindow.Zoom = zoom
            Logger.WriteTraceLog("INFO", sheet.Name + "シートの表示倍率" + zoom.ToString + "に変更")

            'フィルターを解除
            If filterLift And sheet.FilterMode Then
                sheet.ShowAllData()
                Logger.WriteTraceLog("INFO", sheet.Name + "シートのフィルターを解除")
            End If

            'シートの選択範囲を値貼り付け

            'シートの選択範囲をクリア

            'シートの選択範囲を削除

            'シートの表示/非表示を元に戻す
            sheet.Visible = sheet.Visible

            'シートを削除

        Next

        excelBook.Sheets.Select(1)
        excelBook.Save()
        Logger.WriteTraceLog("INFO", excelBook.Name + "を保存")
        excelBook.Close()
        excelApp.Quit()
        KillExcelApp()
    End Sub

    Private Sub KillExcelApp()
        'Excelプロセスを必ずKillする
        For Each p As Process In Process.GetProcessesByName("EXCEL")
            If p.MainWindowTitle.Equals("") Then
                p.Kill()
            End If
        Next
    End Sub

End Module
