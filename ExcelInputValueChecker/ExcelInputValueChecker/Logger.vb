Module Logger

    Private ReadOnly Encoding As String = "Shift_JIS"
    Private ReadOnly LogFileName As String = "ExcelInputValueChecker_" + Date.Now.ToString("yyyyMMdd") + ".log"

    Sub New()
        'Shift_JISを使用可能にする
        Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)
    End Sub

    ''' <summary>
    ''' 情報を提供するメッセージレベルです。
    ''' </summary>
    ''' <param name="msg"></param>
    Public Sub WriteInfoTraceLog(msg As String)
        WriteTraceLog("INFO ", msg)
    End Sub

    ''' <summary>
    ''' 潜在的な問題を示すメッセージレベルです。
    ''' </summary>
    ''' <param name="msg"></param>
    Public Sub WriteWarnTraceLog(msg As String)
        WriteTraceLog("WARN ", msg)
    End Sub

    ''' <summary>
    ''' 重大な障害を示すメッセージレベルです。
    ''' </summary>
    ''' <param name="msg"></param>
    Public Sub WriteErrorTraceLog(msg As String)
        WriteTraceLog("ERROR", msg)
    End Sub

    Private Sub WriteTraceLog(kind As String, msg As String)
        Try
            Dim sw As New IO.StreamWriter(LogFileName, True, Text.Encoding.GetEncoding(Encoding))
            Dim line As String = kind + " " + Date.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + msg
            sw.WriteLine(line)
#If DEBUG Then
            Console.WriteLine(line)
#End If
            sw.Close()

        Catch ex As Exception
            Console.WriteLine("ログの書き込みに失敗しました。")
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
