Imports System.IO

''' <summary>
''' 参照設定に以下を追加
''' Microsoft Scripting Runtime
''' Microsoft Excel 16.0 Object Library
''' </summary>
Module Program

    Private ReadOnly PropertyCsvFileName As String = "ExcelInputValueChecker.csv"
    Private ReadOnly PropertyColumnList As String = "シート名,セル名称,セル番号"
    Private sheetList() As String
    Private cellNameList() As String
    Private cellList() As String

    Private ReadOnly ResultFileName As String = "result_" + Date.Now.ToString("yyyyMMdd") + ".csv"
    Private ReadOnly ResultFilecolumnList As String = "No,ファイル名,シート名,セル名称,設定値"
    Private ResultFileRecords(0) As String

    Private ReadOnly fi As New Scripting.FileSystemObject
    Private ReadOnly targetExtensions As String = "*.xls?"

    Sub Main(args As String())
        Console.WriteLine("Excelの指定したセルの設定値をCSVファイルで出力します。")
        WriteInfoTraceLog("START")
        If ReadProperties() Then
            ReadExcelFile(args)
        End If
        WriteInfoTraceLog("END")
        Console.WriteLine("続行するには何かキーを押してください．．．")
        Console.ReadKey()
    End Sub

    Private Sub ReadExcelFile(args As String())

        Dim targetFiles(0) As String

        Dim i As Integer = 0
        For Each arg As String In args
            If fi.FolderExists(arg) Then
                Dim fileList As String() = Directory.GetFiles(arg, targetExtensions, SearchOption.AllDirectories)
                For Each fileName As String In fileList
                    ReDim Preserve targetFiles(i)
                    targetFiles(i) = fileName
                    i += 1
                Next
            ElseIf fi.FileExists(arg) Then
                ReDim Preserve targetFiles(i)
                targetFiles(i) = arg
                i += 1
            End If
        Next

        For Each file As String In targetFiles
            OpenExcelBook(file)
        Next

        KillExcelApp()
        ReDim Preserve ResultFileRecords(ResultFileRecords.Length - 1)
        CsvFileUtil.Write(ResultFileName, ResultFilecolumnList, ResultFileRecords)
    End Sub

    Private Sub KillExcelApp()
        'Excelプロセスを必ずKillする
        For Each p As Process In Process.GetProcessesByName("EXCEL")
            If p.MainWindowTitle.Equals(String.Empty) Then
                p.Kill()
            End If
        Next
    End Sub

    Private Sub OpenExcelBook(fileName As String)

        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
        Dim excelBook = excelApp.Workbooks.Open(fileName)
        WriteInfoTraceLog("OPEN " + excelBook.Name)

        For Each sheet As Microsoft.Office.Interop.Excel.Worksheet In excelBook.Sheets
            '非表示シートを表示にする
            Dim sheetVisible As Boolean = sheet.Visible
            If Not sheet.Visible Then
                WriteInfoTraceLog(sheet.Name + "シートを非表示から表示に変更")
                sheet.Visible = False
            End If
            OpenExcelSheet(excelBook.Name, sheet)
            'シートの表示/非表示を元に戻す
            sheet.Visible = sheet.Visible
        Next
        WriteInfoTraceLog("CLOSE " + excelBook.Name)
        excelApp.Workbooks.Close()
    End Sub


    Private Sub OpenExcelSheet(excelBookName As String, sheet As Microsoft.Office.Interop.Excel.Worksheet)

        WriteInfoTraceLog(sheet.Name + "シートを確認")

        Dim i As Integer = 0
        For Each sh As String In sheetList
            If sheet.Name.Equals(sh) Then

                Dim inputValue As String = sheet.Range(cellList(i)).Value
                If inputValue Is Nothing Then
                    inputValue = String.Empty
                End If

                '改行を無くす
                inputValue = Replace(inputValue, vbLf, "")


                Dim no As Integer = ResultFileRecords.Length
                ReDim Preserve ResultFileRecords(no)
                Dim record As String = (no).ToString + "," + excelBookName + "," + sheet.Name + "," + cellNameList(i) + "," + inputValue
#If DEBUG Then
                Console.WriteLine(record)
#End If
                ResultFileRecords(no - 1) = record

            End If
            i += 1
        Next
    End Sub


    ''' <summary>
    ''' 前回実行履歴を読み込む
    ''' </summary>
    Private Function ReadProperties() As Boolean
        Dim result As Boolean
        Try
            Dim records() As String = CsvFileUtil.Read(PropertyCsvFileName, 1)
            Console.WriteLine("読み込み情報一覧")
            Console.WriteLine("-----------------------")

            Dim size As Integer = records.Length
            ReDim sheetList(size)
            ReDim cellNameList(size)
            ReDim cellList(size)

            Dim i As Integer = 0
            For Each record As String In records
                Console.WriteLine(record)
                WriteInfoTraceLog(record)
                Dim data() As String = record.Split(",")
                sheetList(i) = data(0)
                cellNameList(i) = data(1)
                cellList(i) = data(2)
                i += 1
            Next
            Console.WriteLine("-----------------------")
            result = True
        Catch ex As Exception
            Console.WriteLine("設定ファイルが存在しないため作成します。")
            Console.WriteLine("設定ファイルの内容を変更の上、再度実行をお願いします。")
            WriteErrorTraceLog("設定ファイルが存在しない。")
            WriteWarnTraceLog("設定ファイルが存在しないため作成します。")
            Dim noRecords(0) As String
            CsvFileUtil.Write(PropertyCsvFileName, PropertyColumnList, noRecords)
        Finally
        End Try
        Return result
    End Function

End Module
