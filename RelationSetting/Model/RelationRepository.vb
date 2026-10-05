Imports RelationSetting.Njc.Common
Imports System.Data.SqlClient
Imports RelationSetting.Njc.Model
Imports System.IO

Namespace Njc.Repository

#Region "仮テーブル設定"

    Public Class TmpTblSetting

        ''' <summary>
        ''' 仮テーブルセッティング
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Setting_TmpTbl(ByRef cnn As System.Data.SqlClient.SqlConnection, _
                                  ByVal tmptblname As String, _
                                  ByVal typestr As String, _
                                  Optional ByVal item As Object = Nothing)

            Dim getqry As New Njc.Query.CommonQuery
            Dim strsql As String = ""

            Select Case typestr
                Case "DROP"
                    strsql = Query.CommonQuery.Qry_TmpTbl_Drop(tmptblname)
                Case "CREATE"
                    strsql = Query.CommonQuery.Qry_TmpTbl_Create(tmptblname, item)
                Case "INSERT"
                    Dim value As String = ""
                    Dim syusyokucnn As String = "',"
                    Dim syusyoku As String = "'"

                    For cntii = 0 To UBound(item)
                        If cntii < UBound(item) Then
                            value = value & syusyoku & item(cntii) & syusyokucnn
                        Else
                            value = value & syusyoku & item(cntii) & syusyoku
                        End If
                    Next

                    strsql = Query.CommonQuery.Qry_TmpTbl_Insert(tmptblname, value)
            End Select

            Dim tmpsqlcom As New SqlCommand(strsql, cnn)
            Dim rowsAffected As Integer = tmpsqlcom.ExecuteNonQuery()

        End Sub

    End Class

#End Region

#Region "CSVファイル関連"

    'Public Class CSVSetting

    '    ''' <summary>
    '    ''' CSV書込
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public Sub Write_CSV(ByVal dgv As DataGridView, ByVal tblname As String)

    '        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS") 'CSVファイルに書き込むときに使うEncoding
    '        Dim csvPath As String = RelDirPath & "\" & tblname & ".csv"
    '        Dim outfile As New StreamWriter(csvPath, False, enc)
    '        Dim colcnt As Integer = dgv.Columns.Count
    '        Dim lastcol As Integer = colcnt - 1
    '        Dim cntii As Integer
    '        Dim cntjj As Integer
    '        '2015.07.06 sol レビュー後修正_レビュー№35 -add
    '        Dim logset As New Njc.Common.LogSetting 'ログ出力用

    '        '2015.07.06 sol レビュー後修正_レビュー№35 -add
    '        Call logset.OutPutLog(LogFilePath, SetLogData.Set_LogData(7, tblname, ""))

    '        'ヘッダを書き込む
    '        For cntjj = 0 To colcnt - 1
    '            'ヘッダの取得
    '            Dim field As String = dgv.Columns(cntjj).HeaderText

    '            '2015.07.06 sol レビュー後修正_レビュー№6 -add
    '            Try
    '                'フィールドを書き込む
    '                outfile.Write(field)
    '                'カンマを書き込む
    '                If lastcol > cntjj Then
    '                    outfile.Write(","c)
    '                End If
    '                '2015.07.06 sol レビュー後修正_レビュー№6 -add sta
    '            Catch ex As Exception
    '                Dim tmp_str = field & " → "
    '                Call logset.OutPutLog(LogFilePath, SetLogData.Set_LogData(21, tmp_str & ex.Message, ""))
    '                Exit Sub
    '            End Try
    '            '2015.07.06 sol レビュー後修正_レビュー№6 -add end
    '        Next
    '        '改行する
    '        outfile.Write(vbCrLf)

    '        'レコードを書き込む
    '        For cntii = 0 To dgv.RowCount - 1
    '            '2015.07.06 sol レビュー後修正_レビュー№34 -add
    '            Application.DoEvents()
    '            For cntjj = 0 To colcnt - 1
    '                'フィールドの取得
    '                Dim field As String = dgv(cntjj, cntii).Value
    '                '2015.07.06 sol レビュー後修正_レビュー№6 -add
    '                Try
    '                    'フィールドを書き込む
    '                    outfile.Write(field)
    '                    'カンマを書き込む
    '                    If lastcol > cntjj Then
    '                        outfile.Write(","c)
    '                    End If
    '                    '2015.07.06 sol レビュー後修正_レビュー№6 -add sta
    '                Catch ex As Exception
    '                    Dim tmp_str As String = dgv.Columns(cntjj).HeaderText & " " & field & " → "
    '                    Call logset.OutPutLog(LogFilePath, SetLogData.Set_LogData(21, tmp_str & ex.Message, ""))
    '                End Try
    '                '2015.07.06 sol レビュー後修正_レビュー№6 -add end
    '            Next
    '            '改行する
    '            outfile.Write(vbCrLf)
    '        Next

    '        '閉じる
    '        outfile.Close()



    '        '2015.07.06 sol レビュー後修正_レビュー№35 -add
    '        Call logset.OutPutLog(LogFilePath, SetLogData.Set_LogData(8, tblname, ""))

    '    End Sub

    '    ''' <summary>
    '    ''' CSV読込
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public Sub Read_CSV(ByRef cnn As System.Data.SqlClient.SqlConnection, ByVal tblname As String, ByVal tmptblname As String)

    '        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
    '        Dim strtmp() As String
    '        Dim cntii As Integer = 0

    '        Dim csvPath As String = RelDirPath & "\" & tblname & ".csv"
    '        Dim readfile As New StreamReader(csvPath, enc)

    '        Dim writetmptbl As New Njc.Repository.TmpTblSetting
    '        '2015.07.06 sol レビュー後修正_レビュー№35 -add
    '        Dim logset As New Njc.Common.LogSetting 'ログ出力用

    '        '2015.07.06 sol レビュー後修正_レビュー№35 -add
    '        Call logset.OutPutLog(LogFilePath, SetLogData.Set_LogData(13, tblname, ""))

    '        Do Until readfile.EndOfStream
    '            '2015.07.06 sol レビュー後修正_レビュー№34 -add
    '            Application.DoEvents()
    '            strtmp = Split(readfile.ReadLine, ",")
    '            'ヘッダをスキップ
    '            If cntii <> 0 Then
    '                'クエリ作成→INSERT処理へ
    '                Call writetmptbl.Setting_TmpTbl(cnn, tmptblname, "INSERT", strtmp)
    '            End If
    '            cntii = cntii + 1
    '        Loop

    '        '2015.07.06 sol レビュー後修正_レビュー№35 -add
    '        Call logset.OutPutLog(LogFilePath, SetLogData.Set_LogData(14, tblname, ""))

    '    End Sub

    'End Class

#End Region

    ' ''' <summary>
    ' ''' 出力ログデータ作成 2015.07.06 sol レビュー後修正_レビュー№35
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Public Class SetLogData

    '    ''' <summary>
    '    ''' ログ内容セット
    '    ''' </summary>
    '    ''' <param name="num">ログNo</param>
    '    ''' <param name="tblname">関連テーブル名</param>
    '    ''' <param name="syosai">詳細内容</param>
    '    ''' <param name="subinfo">他に付加する情報</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    Public Shared Function Set_LogData(ByVal num As Integer, ByVal tblname As String, ByVal syosai As String, Optional ByVal subinfo As String = "") As String()

    '        Dim logmsg() As String = Nothing
    '        ReDim logmsg(12)

    '        logmsg(1) = String.Format(Now)
    '        logmsg(5) = tblname
    '        logmsg(6) = syosai
    '        logmsg(8) = "-"
    '        logmsg(9) = "-"
    '        logmsg(10) = "-"
    '        logmsg(11) = "-"
    '        logmsg(12) = num.ToString

    '        Select Case num
    '            Case 0
    '                logmsg(1) = ""
    '                logmsg(2) = ""
    '                logmsg(3) = ""
    '                logmsg(4) = ""
    '                logmsg(5) = ""
    '                logmsg(6) = ""
    '                logmsg(7) = ""
    '                logmsg(8) = ""
    '                logmsg(9) = ""
    '                logmsg(10) = ""
    '                logmsg(11) = ""
    '                logmsg(12) = ""

    '            Case 1  '処理開始
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = LOG_SYU_BEGIN
    '                logmsg(4) = LOG_NAIYO_REL_STA
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 2 '処理終了
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = LOG_SYU_FIN
    '                logmsg(4) = LOG_NAIYO_REL_END
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 3  'DB接続OK
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = LOG_SYU_CON
    '                logmsg(4) = LOG_NAIYO_CON_OK
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 4  'DB接続NG
    '                logmsg(2) = LOG_RUI_KEIKOKU
    '                logmsg(3) = LOG_SYU_CON
    '                logmsg(4) = LOG_NAIYO_CON_NG
    '                logmsg(7) = LOG_HUBI_CON
    '                logmsg(8) = "-"
    '            Case 5  '接続情報
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "接続情報"
    '                logmsg(4) = "設定内容"
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 6  '実行者出力
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "選択ユーザ(システム)"
    '                logmsg(4) = "ユーザ選択"
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 7  'CSVファイル出力開始(項目毎)
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "CSV作成開始"
    '                logmsg(4) = "紐付項目"
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 8  'CSVファイル出力完了(項目毎)
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "CSV作成終了"
    '                logmsg(4) = "紐付項目"
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 9  '仮テーブル削除開始
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "仮テーブル削除開始"
    '                logmsg(4) = ""
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 10  '仮テーブル削除終了
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "仮テーブル削除終了"
    '                logmsg(4) = ""
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 11  '仮テーブル作成開始
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "仮テーブル作成開始"
    '                logmsg(4) = ""
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 12  '仮テーブル作成終了
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "仮テーブル作成終了"
    '                logmsg(4) = ""
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 13 'CSVファイル読込開始(項目毎)
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "CSV読込開始"
    '                logmsg(4) = "紐付項目"
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"
    '            Case 14 'CSVファイル読込終了(項目毎)
    '                logmsg(2) = LOG_RUI_RIREKI
    '                logmsg(3) = "CSV読込終了"
    '                logmsg(4) = "紐付項目"
    '                logmsg(7) = "-"
    '                logmsg(8) = "-"


    '            Case 21
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "CSV書込処理"
    '                logmsg(4) = "移行失敗"
    '                logmsg(7) = ""
    '                logmsg(8) = ""

    '            Case 22
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "データベース初期化"
    '                logmsg(4) = "初期化失敗"
    '                logmsg(7) = "該当テーブルの初期化に失敗しました(サムネイル画像テーブル)。"
    '                logmsg(8) = "※要検証"
    '            Case 23
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "データベース初期化"
    '                logmsg(4) = "初期化成功"
    '                logmsg(7) = "該当テーブルを初期化しました(サムネイル画像画像テーブル)。"
    '                logmsg(8) = "※要検証"
    '            Case 24
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "データベース初期化"
    '                logmsg(4) = "初期化失敗"
    '                logmsg(7) = "該当テーブルの初期化に失敗しました(実画像テーブル)。"
    '                logmsg(8) = "※要検証"
    '                logmsg(5) = "imgdata"   '変更
    '            Case 25
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "データベース初期化"
    '                logmsg(4) = "初期化成功"
    '                logmsg(7) = "該当テーブルを初期化しました(実画像画像テーブル)。"
    '                logmsg(8) = "-"
    '                logmsg(5) = "imgdata"   '変更

    '            Case 31
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "移行対象判定"
    '                logmsg(4) = "移行対象外"
    '                logmsg(7) = "ファイル名が移行ルールに一致しません。"
    '                logmsg(8) = "お客様にて確認をお願いします。"
    '            Case 32
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "移行対象判定"
    '                logmsg(4) = "移行失敗"
    '                logmsg(7) = "データ重複チェックに失敗しました。"
    '                logmsg(8) = "※要検証"
    '            Case 33
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "移行対象判定"
    '                logmsg(4) = "移行回避"
    '                logmsg(7) = "既にデータが登録されています。"
    '                logmsg(8) = "お客様にて確認をお願いします。"
    '            Case 34
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "移行対象判定"
    '                logmsg(4) = "移行失敗"
    '                logmsg(7) = "既存データの上書きに失敗しました(サムネイル画像既存データ削除失敗)。"
    '                logmsg(8) = "※要検証"
    '            Case 35
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "移行対象判定"
    '                logmsg(4) = "移行失敗"
    '                logmsg(7) = "既存データの上書きに失敗しました(実画像既存データ削除失敗)。"
    '                logmsg(8) = "※要検証"

    '            Case 41
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "登録画像No取得"
    '                logmsg(4) = "移行失敗"
    '                logmsg(7) = "画像Noを取得できません。"
    '                logmsg(8) = "※要検証"
    '            Case 42
    '                logmsg(2) = LOG_RUI_TYUUI
    '                logmsg(3) = "データベース登録"
    '                logmsg(4) = "移行失敗"
    '                logmsg(7) = "基本情報(上位情報)が登録されていません。"
    '                logmsg(8) = "先に基本情報を登録する必要があります。"

    '            Case 99
    '                logmsg(1) = LOG_HEADER1
    '                logmsg(2) = LOG_HEADER2
    '                logmsg(3) = LOG_HEADER3
    '                logmsg(4) = LOG_HEADER4
    '                logmsg(5) = LOG_HEADER5
    '                logmsg(6) = LOG_HEADER6
    '                logmsg(7) = LOG_HEADER7
    '                logmsg(8) = LOG_HEADER8
    '                logmsg(9) = LOG_HEADER9
    '                logmsg(10) = LOG_HEADER10
    '                logmsg(11) = LOG_HEADER11
    '                logmsg(12) = LOG_HEADER12

    '            Case Else

    '        End Select

    '        Dim tmpstr As String = ""
    '        tmpstr = tblname
    '        If tmpstr <> "" Then
    '            logmsg(5) = tmpstr
    '        End If

    '        Return logmsg

    '    End Function

    'End Class

End Namespace


