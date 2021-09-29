Module CsvFileUtil

    Private ReadOnly Encoding As String = "Shift_JIS"

    Sub New()
        'Shift_JISを使用可能にする
        Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)
    End Sub

    ''' <summary>
    ''' CSVファイルを1行目から読み込む
    ''' </summary>
    ''' <param name="fileName">ファイル名</param>
    ''' <returns>レコードリスト</returns>
    Public Function Read(fileName As String) As String()
        Return Read(fileName, 0)
    End Function

    ''' <summary>
    ''' CSVファイルを指定した行から読み込む
    ''' </summary>
    ''' <param name="fileName">ファイル名</param>
    ''' <param name="startLineNo">読み込み開始行</param>
    ''' 0:すべて読み込む
    ''' 1;1行目から読み込む
    ''' <returns>レコードリスト</returns>
    Public Function Read(fileName As String, startLineNo As Integer) As String()
        Dim sr As New System.IO.StreamReader(fileName, Text.Encoding.GetEncoding(Encoding))
        Dim line As String = sr.ReadLine()
        Dim size As Integer = 0
        Dim i As Integer = 0
        Dim records(size) As String
        Do Until line Is Nothing Or line Is String.Empty
            If i > startLineNo - 1 Then


                ReDim Preserve records(size)
                records(size) = line
#If DEBUG Then
                Console.WriteLine(line)
#End If
                size += 1
            End If
            i += 1
            line = sr.ReadLine()
        Loop
        sr.Close()
        Return records
    End Function


    ''' <summary>
    ''' CSVファイルを出力する
    ''' </summary>
    ''' <param name="fileName">ファイル名</param>
    ''' <param name="columnList">カラム名リスト(カンマ区切り)</param>
    ''' <param name="records">レコード(カンマ区切り)</param>
    Public Sub Write(fileName As String, columnList As String, records As String())
        Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)
        Dim sw As New IO.StreamWriter(fileName, False, Text.Encoding.GetEncoding(Encoding))

        'カラム書き込み
        sw.WriteLine(columnList)

        'レコード書き込み
        For Each record As String In records
            sw.WriteLine(record)
#If DEBUG Then
            Console.WriteLine(record)
#End If
        Next
        sw.Close()
    End Sub

End Module
