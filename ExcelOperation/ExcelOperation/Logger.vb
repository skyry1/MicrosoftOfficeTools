Imports System.IO

Public Class Logger

    Private ReadOnly sw As StreamWriter

    Sub New(filePath As String)
        'Shift_JISを使用可能にする
        Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)
        sw = New StreamWriter(filePath, True, Text.Encoding.GetEncoding("Shift_JIS"))
    End Sub

    Public Sub WriteTraceLog(kind As String, msg As String)
        sw.WriteLine(kind + " " + Date.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + msg)
    End Sub

    Public Sub Quit()
        sw.Close()
    End Sub
End Class
