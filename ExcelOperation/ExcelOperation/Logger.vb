Imports System.IO

Module Logger
    Sub New()
        'Shift_JISを使用可能にする
        Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)
    End Sub

    Public Sub WriteTraceLog(kind As String, msg As String)
        Dim sw As New StreamWriter(Date.Now.ToString("yyyyMMdd") + ".log", True, Text.Encoding.GetEncoding("Shift_JIS"))
        sw.WriteLine(kind + " " + Date.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + msg)
        sw.Close()
    End Sub
End Module
