Imports ClosedXML.Excel

Namespace Njc.Common

    ''' <summary>
    ''' 紐付設定ファイル(xlsx)の読み書き
    ''' 以前は ACE.OLEDB (Microsoft.ACE.OLEDB.12.0, HDR=YES) で行っていた処理を ClosedXML で置き換えたもの。
    ''' Access Database Engine が不要になり、32/64ビットどちらでも動作する。
    ''' ・1行目を見出し行、2行目以降をデータ行として扱う(HDR=YES と同じ)
    ''' ・空のセルは DBNull として返す(ACE.OLEDB と同じ)
    ''' </summary>
    Public Class RelationExcelFile

        Private Const HEADER_ROW As Integer = 1

        ''' <summary>
        ''' シートの内容を DataTable で取得する (旧: SELECT * FROM [シート名$])
        ''' </summary>
        Public Shared Function ReadSheet(ByVal filepath As String, ByVal sheetname As String) As DataTable
            Using wbook As New XLWorkbook(filepath)
                Return ReadSheet(wbook.Worksheet(sheetname))
            End Using
        End Function

        ''' <summary>
        ''' シートの内容を DataTable で取得する (ブックを開いたまま複数シートを読む場合用)
        ''' </summary>
        Public Shared Function ReadSheet(ByVal wsheet As IXLWorksheet) As DataTable

            Dim dt As New DataTable()

            Dim lastcol As Integer = GetLastColumn(wsheet)
            Dim lastrow As Integer = GetLastRow(wsheet)

            '見出し行から列を作成(空の見出しは F1, F2… / 重複する見出しは末尾に連番を付ける)
            For col = 1 To lastcol
                Dim colname As String = wsheet.Cell(HEADER_ROW, col).GetString().Trim()
                If colname = "" Then
                    colname = "F" & col.ToString()
                End If
                Dim basename As String = colname
                Dim seq As Integer = 1
                While dt.Columns.Contains(colname)
                    colname = basename & seq.ToString()
                    seq += 1
                End While
                dt.Columns.Add(colname, GetType(Object))
            Next

            'データ行
            For row = HEADER_ROW + 1 To lastrow
                Dim values(lastcol - 1) As Object
                For col = 1 To lastcol
                    values(col - 1) = ToObject(wsheet.Cell(row, col).Value)
                Next
                dt.Rows.Add(values)
            Next

            Return dt

        End Function

        ''' <summary>
        ''' 指定したシートにデータ行(見出し行以外)が1行以上あるか
        ''' </summary>
        Public Shared Function HasAnyData(ByVal wbook As XLWorkbook, ByVal sheetname As String) As Boolean
            Return GetLastRow(wbook.Worksheet(sheetname)) > HEADER_ROW
        End Function

        ''' <summary>
        ''' シートの最終行の後ろに行を追加する (旧: INSERT INTO [シート名$] VALUES(...))
        ''' 値はすべて文字列として書き込む(ACE.OLEDB で文字列を INSERT していたのと同じ)
        ''' </summary>
        Public Shared Sub AppendRows(ByVal filepath As String, ByVal sheetname As String, ByVal rows As List(Of String()))

            Using wbook As New XLWorkbook(filepath)

                Dim wsheet As IXLWorksheet = wbook.Worksheet(sheetname)
                Dim writerow As Integer = Math.Max(GetLastRow(wsheet), HEADER_ROW) + 1

                For Each rowvalues In rows
                    For col = 1 To rowvalues.Length
                        Dim cell As IXLCell = wsheet.Cell(writerow, col)
                        If String.IsNullOrEmpty(rowvalues(col - 1)) Then
                            cell.Value = Blank.Value
                        Else
                            cell.Value = rowvalues(col - 1)
                        End If
                    Next
                    writerow += 1
                Next

                wbook.Save()

            End Using

        End Sub

        ''' <summary>
        ''' 値が入っている最終行(書式だけのセルは含めない)。データが無ければ 0
        ''' </summary>
        Private Shared Function GetLastRow(ByVal wsheet As IXLWorksheet) As Integer
            Dim lastrow As IXLRow = wsheet.LastRowUsed(XLCellsUsedOptions.Contents)
            If lastrow Is Nothing Then
                Return 0
            End If
            Return lastrow.RowNumber()
        End Function

        ''' <summary>
        ''' 値が入っている最終列(書式だけのセルは含めない)。データが無ければ 0
        ''' </summary>
        Private Shared Function GetLastColumn(ByVal wsheet As IXLWorksheet) As Integer
            Dim lastcol As IXLColumn = wsheet.LastColumnUsed(XLCellsUsedOptions.Contents)
            If lastcol Is Nothing Then
                Return 0
            End If
            Return lastcol.ColumnNumber()
        End Function

        ''' <summary>
        ''' セルの値を ACE.OLEDB で読んだときと同じ型の値に変換する
        ''' </summary>
        Private Shared Function ToObject(ByVal value As XLCellValue) As Object
            Select Case value.Type
                Case XLDataType.Blank
                    Return DBNull.Value
                Case XLDataType.Text
                    Dim text As String = value.GetText()
                    If text = "" Then
                        Return DBNull.Value
                    End If
                    Return text
                Case XLDataType.Number
                    Return value.GetNumber()
                Case XLDataType.Boolean
                    Return value.GetBoolean()
                Case XLDataType.DateTime
                    Return value.GetDateTime()
                Case XLDataType.TimeSpan
                    Return value.GetTimeSpan()
                Case Else
                    'エラー値など
                    Return DBNull.Value
            End Select
        End Function

    End Class

End Namespace
