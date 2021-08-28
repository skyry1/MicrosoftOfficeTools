Imports System.IO

Module Logger

    Private ReadOnly sw As StreamWriter

    Sub New()
        'Shift_JISを使用可能にする
        Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)
    End Sub

    Public Sub WriteTraceLog(kind As String, msg As String)
        Dim sw As New StreamWriter(Date.Now.ToString("yyyyMMdd") + ".log", True, Text.Encoding.GetEncoding("Shift_JIS"))
        sw.WriteLine(kind + " " + Date.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + msg)
        sw.Close()
    End Sub

    Public Sub Quit()

    End Sub
End Module
