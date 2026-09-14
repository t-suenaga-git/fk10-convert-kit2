Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

    '20160929 汎用→既存コピー処理改善対応 -add sta
    '作り直し分

#Region "部屋駐車場情報"

    Public Class Hy_cyusyajo_mid_Repository

        ''' <summary>
        ''' 部屋駐車場情報のコピー処理
        ''' '20161021 既存中間ファイルを無くすことによる速度改善処理2 引数にSQLSERVERへの接続オブジェクト(sqlcnnv10)と仮テーブル名(tblname)およびデータテーブルオブジェクト(dt)を追加
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="dt"></param>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Read_BaseMidFile(ByVal sqlcnnv10 As SqlConnection, ByVal dt As DataTable, ByVal tblname As String) As Boolean

            Dim rtn As Boolean = True
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim basemidfilename As String = "中間ファイル"
            'Dim basemidfilesheetname As String = "部屋情報"
            'Dim existmidfilename As String = "部屋情報"
            'Dim existmidfilesheetname As String = "部屋駐車場情報"
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            Dim qry As String = ""

            Dim da As New OleDbDataAdapter()
            Dim ds As DataSet = New DataSet()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del
            'Dim dt As New DataTable() 
            Dim dr As OleDbDataReader

            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
            'Dim tmpextend As String = "Excel 8.0;HDR=YES;"

            'Dim con_read As New OleDbConnection()
            'Dim cmd_read As New OleDbCommand()
            'Dim con_write As New OleDbConnection()
            'Dim cmd_write As New OleDbCommand()

            'Dim basemidpath As String = EtcMethod.Set_Path(BaseMidDirPath, basemidfilename & ".xlsx")
            'Dim existmidpath As String = EtcMethod.Set_Path(MiddleDirPath, existmidfilename & ".xlsx")
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            '20161014 実行時の進捗表示対応 -add sta
            Dim obj_pgb As New ProgressBarManager
            Dim obj_com As New Njc.Common.CommonRepository
            '20161014 実行時の進捗表示対応 -add end
            Try

                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
                ''================================================================================
                '' ●接続設定●
                ''================================================================================
                ''--------------------------------------------------
                '' ①EXCEL接続 (ファイル読込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_read.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & basemidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_read.Connection = con_read

                ''--------------------------------------------------
                '' ②EXCEL接続 (ファイル書込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_write.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & existmidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_write.Connection = con_write

                ''--------------------------------------------------
                '' 読込→書込
                ''--------------------------------------------------
                ''接続オープン処理
                'con_read.Open()
                'con_write.Open()

                'qry = "SELECT * FROM [" & basemidfilesheetname & "$] "

                'cmd_read = con_read.CreateCommand
                'cmd_read.CommandText = qry
                'dt = New DataTable
                'dr = cmd_read.ExecuteReader
                'dt.Load(dr)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
                Dim colcnt As Integer = dt.Columns.Count
                Dim readrowcnt As Integer = 1
                Dim solist_header As New SortedList(Of Integer, String)

                '20161014 実行時の進捗表示対応 -add sta
                'プログレスバー初期化
                Dim pgbtotalcnt_part As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If dt.Rows.Count <= pgbbasecnt Then
                    pgbtotalcnt_part = dt.Rows.Count
                Else
                    pgbtotalcnt_part = Math.Ceiling(dt.Rows.Count / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt_part)
                '20161014 実行時の進捗表示対応 -add end

                '1行ずつ取得
                Dim row As DataRow
                For Each row In dt.Rows

                    Dim tmp_hash As New Hashtable
                    Dim tmp_solist As New SortedList(Of Integer, String)        '作業用オブジェクト
                    Dim solist_value As New SortedList(Of Integer, String)      'ヘッダー格納用オブジェクト

                    '行データを取得して作業用オブジェクトへ格納
                    For cntjj = 1 To colcnt - 1
                        tmp_solist.Add(cntjj, row(cntjj).ToString)
                    Next

                    '行数から判断する
                    If readrowcnt = 1 Then

                        '1行目はヘッダーなのでヘッダー用のオブジェクトへ格納
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'solist_header = tmp_solist
                        For cntjj = 1 To tmp_solist.Count

                            If tmp_solist(cntjj) <> "" Then
                                solist_header.Add(cntjj, tmp_solist(cntjj))
                            End If

                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                    ElseIf readrowcnt >= BASEMIDFILE_READWRITE_ROW - 1 Then

                        '実データ開始行からデータ格納用オブジェクトへ格納
                        solist_value = tmp_solist

                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'For cntkk = 1 To colcnt - 1
                        '    Dim tmp_baseheader As String = solist_header(cntkk)
                        '    Dim tmp_basevalue As String = solist_value(cntkk)
                        '    tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                        'Next
                        For cntkk = 1 To solist_header.Count
                            Dim tmp_baseheader As String = solist_header(cntkk)
                            Dim tmp_basevalue As String = solist_value(cntkk)
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add sta
                            If tmp_basevalue Is Nothing Then
                                tmp_basevalue = ""
                            End If
                            tmp_basevalue.Replace("'", "’")
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add end
                            If tmp_baseheader <> "" Then
                                tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                            End If
                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                        For cntkbn = 1 To 1 '1 To 3→1 To 1

                            Dim insertvalue_sb As New System.Text.StringBuilder
                            Dim chk_sb As New System.Text.StringBuilder
                            Dim insertflg As Boolean = True

                            insertvalue_sb.Append("'" & tmp_hash("物件NO") & "'")
                            insertvalue_sb.Append(",'" & tmp_hash("部屋NO") & "'")
                            insertvalue_sb.Append(",'" & cntkbn.ToString & "'")

                            Select Case cntkbn
                                Case 1

                                    '空文字チェック
                                    chk_sb.Append(tmp_hash("駐車場の空き数(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐車場の空き有無(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐車場料金区分(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐車場料金(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐車場料金税区分(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐車場の賃貸可能数"))

                                    '空の場合は挿入処理を行わない
                                    If chk_sb.ToString.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("駐車場の空き数(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐車場の空き有無(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐車場料金区分(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐車場料金(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐車場料金税区分(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐車場の賃貸可能数") & "'")
                                    End If

                                Case 2

                                    '空文字チェック
                                    chk_sb.Append(tmp_hash("バイクの空き数(手動設定用)"))
                                    chk_sb.Append(tmp_hash("バイク駐車場の空き有無(手動設定用)"))
                                    chk_sb.Append(tmp_hash("バイク駐車場料金区分(手動設定用)"))
                                    chk_sb.Append(tmp_hash("バイク駐車場料金(手動設定用)"))
                                    chk_sb.Append(tmp_hash("バイク駐車場料金税区分(手動設定用)"))
                                    chk_sb.Append(tmp_hash("バイク駐車場の賃貸可能数"))

                                    '空の場合は挿入処理を行わない
                                    If chk_sb.ToString.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("バイクの空き数(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("バイク駐車場の空き有無(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("バイク駐車場料金区分(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("バイク駐車場料金(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("バイク駐車場料金税区分(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("バイク駐車場の賃貸可能数") & "'")
                                    End If

                                Case 3

                                    '空文字チェック
                                    chk_sb.Append(tmp_hash("駐輪場駐車場の空き数(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐輪場駐車場の空き有無(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐輪場駐車場料金区分(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐輪場駐車場料金(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐輪場駐車場料金税区分(手動設定用)"))
                                    chk_sb.Append(tmp_hash("駐輪場駐車場の賃貸可能数"))

                                    '空の場合は挿入処理を行わない
                                    If chk_sb.ToString.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("駐輪場駐車場の空き数(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐輪場駐車場の空き有無(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐輪場駐車場料金区分(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐輪場駐車場料金(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐輪場駐車場料金税区分(手動設定用)") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("駐輪場駐車場の賃貸可能数") & "'")
                                    End If

                            End Select

                            '挿入
                            If insertflg Then

                                Dim taisyo As String = "[物件NO],[部屋NO],[駐車場区分],[駐車場の空き数(手動設定用)],[駐車場の空き有無(手動設定用)],[駐車場料金区分(手動設定用)],[駐車場料金(手動設定用)],[駐車場料金税区分(手動設定用)],[駐車場の賃貸可能数]"

                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg sta
                                'qry = ""
                                'qry += " INSERT INTO [" & existmidfilesheetname & "$] "
                                'qry += "(" & taisyo & ")"
                                'qry += " VALUES "
                                'qry += "(" & insertvalue_sb.ToString & ")"

                                'cmd_write.CommandText = qry
                                'cmd_write.ExecuteNonQuery()
                                Dim tmp_cnt As Integer = 0
                                qry = ""
                                qry += " INSERT INTO " & tblname
                                qry += "(" & taisyo & ")"
                                qry += " VALUES "
                                qry += "(" & insertvalue_sb.ToString & ")"
                                DBExec.Exec_NonQuery(sqlcnnv10, qry, tmp_cnt)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg end

                            End If

                        Next

                    End If

                    '20161014 実行時の進捗表示対応 -add sta
                    '-------------------------------
                    'プログレスバー更新/進捗率表示
                    '-------------------------------
                    Dim tmp_pgbcnt As Integer = 0
                    If dt.Rows.Count <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        tmp_pgbcnt = readrowcnt
                    ElseIf dt.Rows.Count > readrowcnt Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        Call EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, tmp_pgbcnt)
                    ElseIf dt.Rows.Count <= readrowcnt Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        tmp_pgbcnt = pgbtotalcnt_part
                    End If

                    '表示
                    If tmp_pgbcnt <> 0 Then
                        Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                        Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                    End If
                    '20161017 進捗表示処理による速度低下の修正 -del
                    'Njc.Frm.MainFrm.Refresh()
                    '20161014 実行時の進捗表示対応 -add end

                    readrowcnt = readrowcnt + 1

                Next

            Catch ex As Exception

                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sat
                ''クローズ処理
                'con_read.Close()
                'con_write.Close()
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                rtn = False
                Return rtn

            End Try

            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
            ''クローズ処理
            'con_read.Close()
            'con_write.Close()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end

            Return rtn

        End Function

    End Class

#End Region

#Region "部屋特約情報"

    Public Class Hy_tokuyaku_mid_Repository

        ''' <summary>
        ''' 部屋特約情報コピー処理
        ''' '20161021 既存中間ファイルを無くすことによる速度改善処理2 引数にSQLSERVERへの接続オブジェクト(sqlcnnv10)と仮テーブル名(tblname)およびデータテーブルオブジェクト(dt)を追加
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="dt"></param>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Read_BaseMidFile(ByVal sqlcnnv10 As SqlConnection, ByVal dt As DataTable, ByVal tblname As String) As Boolean

            Dim rtn As Boolean = True
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim basemidfilename As String = "中間ファイル"
            'Dim basemidfilesheetname As String = "部屋情報"
            'Dim existmidfilename As String = "部屋情報"
            'Dim existmidfilesheetname As String = "部屋特約情報"
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            Dim qry As String = ""

            Dim da As New OleDbDataAdapter()
            Dim ds As DataSet = New DataSet()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del
            'Dim dt As New DataTable()   
            Dim dr As OleDbDataReader
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
            'Dim tmpextend As String = "Excel 8.0;HDR=YES;"

            'Dim con_read As New OleDbConnection()
            'Dim cmd_read As New OleDbCommand()
            'Dim con_write As New OleDbConnection()
            'Dim cmd_write As New OleDbCommand()

            'Dim basemidpath As String = EtcMethod.Set_Path(BaseMidDirPath, basemidfilename & ".xlsx")
            'Dim existmidpath As String = EtcMethod.Set_Path(MiddleDirPath, existmidfilename & ".xlsx")
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            '20161014 実行時の進捗表示対応 -add sta
            Dim obj_pgb As New ProgressBarManager
            Dim obj_com As New Njc.Common.CommonRepository
            '20161014 実行時の進捗表示対応 -add end

            Try
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
                ''================================================================================
                '' ●接続設定●
                ''================================================================================
                ''--------------------------------------------------
                '' ①EXCEL接続 (ファイル読込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_read.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & basemidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_read.Connection = con_read

                ''--------------------------------------------------
                '' ②EXCEL接続 (ファイル書込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_write.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & existmidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_write.Connection = con_write

                ''--------------------------------------------------
                '' 読込→書込
                ''--------------------------------------------------
                ''接続オープン処理
                'con_read.Open()
                'con_write.Open()

                'qry = "SELECT * FROM [" & basemidfilesheetname & "$] "

                'cmd_read = con_read.CreateCommand
                'cmd_read.CommandText = qry
                'dt = New DataTable
                'dr = cmd_read.ExecuteReader
                'dt.Load(dr)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
                Dim colcnt As Integer = dt.Columns.Count
                Dim readrowcnt As Integer = 1
                Dim solist_header As New SortedList(Of Integer, String)

                '20161014 実行時の進捗表示対応 -add sta
                'プログレスバー初期化
                Dim pgbtotalcnt_part As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If dt.Rows.Count <= pgbbasecnt Then
                    pgbtotalcnt_part = dt.Rows.Count
                Else
                    pgbtotalcnt_part = Math.Ceiling(dt.Rows.Count / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt_part)
                '20161014 実行時の進捗表示対応 -add end

                '1行ずつ取得
                Dim row As DataRow
                For Each row In dt.Rows

                    Dim tmp_hash As New Hashtable
                    Dim tmp_solist As New SortedList(Of Integer, String)        '作業用オブジェクト
                    Dim solist_value As New SortedList(Of Integer, String)      'ヘッダー格納用オブジェクト

                    '行データを取得して作業用オブジェクトへ格納
                    For cntjj = 1 To colcnt - 1
                        tmp_solist.Add(cntjj, row(cntjj).ToString)
                    Next

                    '行数から判断する
                    If readrowcnt = 1 Then

                        '1行目はヘッダーなのでヘッダー用のオブジェクトへ格納
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'solist_header = tmp_solist
                        For cntjj = 1 To tmp_solist.Count

                            If tmp_solist(cntjj) <> "" Then
                                solist_header.Add(cntjj, tmp_solist(cntjj))
                            End If

                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                    ElseIf readrowcnt >= BASEMIDFILE_READWRITE_ROW - 1 Then

                        '実データ開始行からデータ格納用オブジェクトへ格納
                        solist_value = tmp_solist

                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'For cntkk = 1 To colcnt - 1
                        '    Dim tmp_baseheader As String = solist_header(cntkk)
                        '    Dim tmp_basevalue As String = solist_value(cntkk)
                        '    tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                        'Next
                        For cntkk = 1 To solist_header.Count
                            Dim tmp_baseheader As String = solist_header(cntkk)
                            Dim tmp_basevalue As String = solist_value(cntkk)
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add sta
                            If tmp_basevalue Is Nothing Then
                                tmp_basevalue = ""
                            End If
                            tmp_basevalue.Replace("'", "’")
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add end
                            If tmp_baseheader <> "" Then
                                tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                            End If
                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                        For cntkbn = 1 To 3

                            Dim insertvalue_sb As New System.Text.StringBuilder
                            Dim chk_str As String = ""
                            Dim insertflg As Boolean = True

                            insertvalue_sb.Append("'" & tmp_hash("物件NO") & "'")
                            insertvalue_sb.Append(",'" & tmp_hash("部屋NO") & "'")
                            insertvalue_sb.Append(",'" & cntkbn.ToString & "'")

                            Select Case cntkbn
                                Case 1

                                    '空文字チェック
                                    chk_str = IIf(tmp_hash("その他特約内容") Is Nothing, "", tmp_hash("その他特約内容"))

                                    '空の場合は挿入処理を行わない
                                    If chk_str.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("その他特約内容") & "'")
                                    End If

                                Case 2

                                    '空文字チェック
                                    chk_str = IIf(tmp_hash("原状回復特約内容") Is Nothing, "", tmp_hash("原状回復特約内容"))

                                    '空の場合は挿入処理を行わない
                                    If chk_str.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("原状回復特約内容") & "'")
                                    End If

                                Case 3

                                    '空文字チェック
                                    chk_str = IIf(tmp_hash("入居中修繕特約内容") Is Nothing, "", tmp_hash("入居中修繕特約内容"))

                                    '空の場合は挿入処理を行わない
                                    If chk_str.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("入居中修繕特約内容") & "'")
                                    End If

                            End Select

                            '挿入
                            If insertflg Then

                                Dim taisyo As String = "[物件NO],[部屋NO],[特約区分],[内容1]"

                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg sta
                                'qry = ""
                                'qry += " INSERT INTO [" & existmidfilesheetname & "$] "
                                'qry += "(" & taisyo & ")"
                                'qry += " VALUES "
                                'qry += "(" & insertvalue_sb.ToString & ")"

                                'cmd_write.CommandText = qry
                                'cmd_write.ExecuteNonQuery()
                                Dim tmp_cnt As Integer = 0
                                qry = ""
                                qry += " INSERT INTO " & tblname
                                qry += "(" & taisyo & ")"
                                qry += " VALUES "
                                qry += "(" & insertvalue_sb.ToString & ")"
                                DBExec.Exec_NonQuery(sqlcnnv10, qry, tmp_cnt)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg end

                            End If

                        Next

                    End If

                    '20161014 実行時の進捗表示対応 -add sta
                    '-------------------------------
                    'プログレスバー更新/進捗率表示
                    '-------------------------------
                    Dim tmp_pgbcnt As Integer = 0
                    If dt.Rows.Count <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        tmp_pgbcnt = readrowcnt
                    ElseIf dt.Rows.Count > readrowcnt Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        Call EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, tmp_pgbcnt)
                    ElseIf dt.Rows.Count <= readrowcnt Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        tmp_pgbcnt = pgbtotalcnt_part
                    End If

                    '表示
                    If tmp_pgbcnt <> 0 Then
                        Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                        Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                    End If
                    '20161017 進捗表示処理による速度低下の修正 -del
                    'Njc.Frm.MainFrm.Refresh()
                    '20161014 実行時の進捗表示対応 -add end

                    readrowcnt = readrowcnt + 1

                Next

            Catch ex As Exception
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'con_read.Close()
                'con_write.Close()
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                rtn = False
                Return rtn

            End Try
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
            ''クローズ処理
            'con_read.Close()
            'con_write.Close()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
            Return rtn

        End Function

    End Class

#End Region

#Region "契約契約者情報"

    Public Class Ky_kys_mid_Repository

        ''' <summary>
        ''' 契約契約者情報のコピー処理
        ''' '20161021 既存中間ファイルを無くすことによる速度改善処理2 引数にSQLSERVERへの接続オブジェクト(sqlcnnv10)と仮テーブル名(tblname)およびデータテーブルオブジェクト(dt)を追加
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="dt"></param>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Read_BaseMidFile(ByVal sqlcnnv10 As SqlConnection, ByVal dt As DataTable, ByVal tblname As String) As Boolean

            Dim rtn As Boolean = True
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim basemidfilename As String = "中間ファイル"
            'Dim basemidfilesheetname As String = "契約契約者保証人情報"
            'Dim existmidfilename As String = "契約情報"
            'Dim existmidfilesheetname As String = "契約契約者情報"
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            Dim qry As String = ""
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            Dim da As New OleDbDataAdapter()
            Dim ds As DataSet = New DataSet()
            'Dim dt As New DataTable()   '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del
            Dim dr As OleDbDataReader
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
            'Dim tmpextend As String = "Excel 8.0;HDR=YES;"

            'Dim con_read As New OleDbConnection()
            'Dim cmd_read As New OleDbCommand()
            'Dim con_write As New OleDbConnection()
            'Dim cmd_write As New OleDbCommand()

            'Dim basemidpath As String = EtcMethod.Set_Path(BaseMidDirPath, basemidfilename & ".xlsx")
            'Dim existmidpath As String = EtcMethod.Set_Path(MiddleDirPath, existmidfilename & ".xlsx")
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            '20161014 実行時の進捗表示対応 -add sta
            Dim obj_pgb As New ProgressBarManager
            Dim obj_com As New Njc.Common.CommonRepository
            '20161014 実行時の進捗表示対応 -add end

            Try
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
                ''================================================================================
                '' ●接続設定●
                ''================================================================================
                ''--------------------------------------------------
                '' ①EXCEL接続 (ファイル読込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_read.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & basemidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_read.Connection = con_read

                ''--------------------------------------------------
                '' ②EXCEL接続 (ファイル書込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_write.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & existmidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_write.Connection = con_write

                ''--------------------------------------------------
                '' 読込→書込
                ''--------------------------------------------------
                ''接続オープン処理
                'con_read.Open()
                'con_write.Open()

                'qry = "SELECT * FROM [" & basemidfilesheetname & "$] "

                'cmd_read = con_read.CreateCommand
                'cmd_read.CommandText = qry
                'dt = New DataTable
                'dr = cmd_read.ExecuteReader
                'dt.Load(dr)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
                Dim colcnt As Integer = dt.Columns.Count
                Dim readrowcnt As Integer = 1
                Dim solist_header As New SortedList(Of Integer, String)

                '20161014 実行時の進捗表示対応 -add sta
                'プログレスバー初期化
                Dim pgbtotalcnt_part As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If dt.Rows.Count <= pgbbasecnt Then
                    pgbtotalcnt_part = dt.Rows.Count
                Else
                    pgbtotalcnt_part = Math.Ceiling(dt.Rows.Count / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt_part)
                '20161014 実行時の進捗表示対応 -add end

                '1行ずつ取得
                Dim row As DataRow
                For Each row In dt.Rows

                    Dim tmp_hash As New Hashtable
                    Dim tmp_solist As New SortedList(Of Integer, String)        '作業用オブジェクト
                    Dim solist_value As New SortedList(Of Integer, String)      'ヘッダー格納用オブジェクト

                    '行データを取得して作業用オブジェクトへ格納
                    For cntjj = 1 To colcnt - 1
                        tmp_solist.Add(cntjj, row(cntjj).ToString)
                    Next

                    '行数から判断する
                    If readrowcnt = 1 Then

                        '1行目はヘッダーなのでヘッダー用のオブジェクトへ格納
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'solist_header = tmp_solist
                        For cntjj = 1 To tmp_solist.Count

                            If tmp_solist(cntjj) <> "" Then
                                solist_header.Add(cntjj, tmp_solist(cntjj))
                            End If

                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                    ElseIf readrowcnt >= BASEMIDFILE_READWRITE_ROW - 1 Then

                        '実データ開始行からデータ格納用オブジェクトへ格納
                        solist_value = tmp_solist

                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'For cntkk = 1 To colcnt - 1
                        '    Dim tmp_baseheader As String = solist_header(cntkk)
                        '    Dim tmp_basevalue As String = solist_value(cntkk)
                        '    tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                        'Next
                        For cntkk = 1 To solist_header.Count
                            Dim tmp_baseheader As String = solist_header(cntkk)
                            Dim tmp_basevalue As String = solist_value(cntkk)
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add sta
                            If tmp_basevalue Is Nothing Then
                                tmp_basevalue = ""
                            End If
                            tmp_basevalue.Replace("'", "’")
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add end
                            If tmp_baseheader <> "" Then
                                tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                            End If
                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                        For cntkbn = 1 To 3

                            Dim insertvalue_sb As New System.Text.StringBuilder
                            Dim chk_sb As New System.Text.StringBuilder
                            Dim insertflg As Boolean = True

                            insertvalue_sb.Append("'" & tmp_hash("物件No") & "'")
                            insertvalue_sb.Append(",'" & tmp_hash("部屋No") & "'")
                            insertvalue_sb.Append(",'" & cntkbn.ToString & "'")

                            Select Case cntkbn
                                Case 1

                                    '空文字チェック
                                    chk_sb.Append(tmp_hash("契約者No1"))
                                    chk_sb.Append(tmp_hash("入居フラグ1"))

                                    '空の場合は挿入処理を行わない
                                    If chk_sb.ToString.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("契約者No1") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("入居フラグ1") & "'")
                                    End If

                                Case 2

                                    '空文字チェック
                                    chk_sb.Append(tmp_hash("契約者No2"))
                                    chk_sb.Append(tmp_hash("入居フラグ3"))

                                    '空の場合は挿入処理を行わない
                                    If chk_sb.ToString.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("契約者No2") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("入居フラグ2") & "'")
                                    End If

                                Case 3

                                    '空文字チェック
                                    chk_sb.Append(tmp_hash("契約者No3"))
                                    chk_sb.Append(tmp_hash("入居フラグ3"))

                                    '空の場合は挿入処理を行わない
                                    If chk_sb.ToString.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("契約者No3") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("入居フラグ3") & "'")
                                    End If

                            End Select

                            '挿入
                            If insertflg Then

                                Dim taisyo As String = "[物件NO],[部屋NO],[並び順No],[契約者No],[入居フラグ]"

                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg sta
                                'qry = ""
                                'qry += " INSERT INTO [" & existmidfilesheetname & "$] "
                                'qry += "(" & taisyo & ")"
                                'qry += " VALUES "
                                'qry += "(" & insertvalue_sb.ToString & ")"

                                'cmd_write.CommandText = qry
                                'cmd_write.ExecuteNonQuery()
                                Dim tmp_cnt As Integer = 0
                                qry = ""
                                qry += " INSERT INTO " & tblname
                                qry += "(" & taisyo & ")"
                                qry += " VALUES "
                                qry += "(" & insertvalue_sb.ToString & ")"
                                DBExec.Exec_NonQuery(sqlcnnv10, qry, tmp_cnt)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg end

                            End If

                        Next

                    End If

                    '20161014 実行時の進捗表示対応 -add sta
                    '-------------------------------
                    'プログレスバー更新/進捗率表示
                    '-------------------------------
                    Dim tmp_pgbcnt As Integer = 0
                    If dt.Rows.Count <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        tmp_pgbcnt = readrowcnt
                    ElseIf dt.Rows.Count > readrowcnt Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        Call EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, tmp_pgbcnt)
                    ElseIf dt.Rows.Count <= readrowcnt Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        tmp_pgbcnt = pgbtotalcnt_part
                    End If

                    '表示
                    If tmp_pgbcnt <> 0 Then
                        Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                        Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                    End If
                    '20161017 進捗表示処理による速度低下の修正 -del
                    'Njc.Frm.MainFrm.Refresh()
                    '20161014 実行時の進捗表示対応 -add end

                    readrowcnt = readrowcnt + 1

                Next

            Catch ex As Exception
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'con_read.Close()
                'con_write.Close()
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                rtn = False
                Return rtn

            End Try
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
            ''クローズ処理
            'con_read.Close()
            'con_write.Close()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
            Return rtn

        End Function

    End Class

#End Region

#Region "契約保証人情報"

    Public Class Ky_hosyonin_mid_Repository

        ''' <summary>
        ''' 契約保証人情報のコピー処理
        ''' '20161021 既存中間ファイルを無くすことによる速度改善処理2 引数にSQLSERVERへの接続オブジェクト(sqlcnnv10)と仮テーブル名(tblname)およびデータテーブルオブジェクト(dt)を追加
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="dt"></param>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Read_BaseMidFile(ByVal sqlcnnv10 As SqlConnection, ByVal dt As DataTable, ByVal tblname As String) As Boolean

            Dim rtn As Boolean = True
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim basemidfilename As String = "中間ファイル"
            'Dim basemidfilesheetname As String = "契約契約者保証人情報"
            'Dim existmidfilename As String = "契約情報"
            'Dim existmidfilesheetname As String = "契約保証人情報"
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            Dim qry As String = ""

            Dim da As New OleDbDataAdapter()
            Dim ds As DataSet = New DataSet()
            'Dim dt As New DataTable()   '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del
            Dim dr As OleDbDataReader
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
            'Dim tmpextend As String = "Excel 8.0;HDR=YES;"

            'Dim con_read As New OleDbConnection()
            'Dim cmd_read As New OleDbCommand()
            'Dim con_write As New OleDbConnection()
            'Dim cmd_write As New OleDbCommand()

            'Dim basemidpath As String = EtcMethod.Set_Path(BaseMidDirPath, basemidfilename & ".xlsx")
            'Dim existmidpath As String = EtcMethod.Set_Path(MiddleDirPath, existmidfilename & ".xlsx")
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            '20161014 実行時の進捗表示対応 -add sta
            Dim obj_pgb As New ProgressBarManager
            Dim obj_com As New Njc.Common.CommonRepository
            '20161014 実行時の進捗表示対応 -add end

            Try
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
                ''================================================================================
                '' ●接続設定●
                ''================================================================================
                ''--------------------------------------------------
                '' ①EXCEL接続 (ファイル読込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_read.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & basemidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_read.Connection = con_read

                ''--------------------------------------------------
                '' ②EXCEL接続 (ファイル書込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_write.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & existmidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_write.Connection = con_write

                ''--------------------------------------------------
                '' 読込→書込
                ''--------------------------------------------------
                ''接続オープン処理
                'con_read.Open()
                'con_write.Open()

                'qry = "SELECT * FROM [" & basemidfilesheetname & "$] "

                'cmd_read = con_read.CreateCommand
                'cmd_read.CommandText = qry
                'dt = New DataTable
                'dr = cmd_read.ExecuteReader
                'dt.Load(dr)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
                Dim colcnt As Integer = dt.Columns.Count
                Dim readrowcnt As Integer = 1
                Dim solist_header As New SortedList(Of Integer, String)

                '20161014 実行時の進捗表示対応 -add sta
                'プログレスバー初期化
                Dim pgbtotalcnt_part As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If dt.Rows.Count <= pgbbasecnt Then
                    pgbtotalcnt_part = dt.Rows.Count
                Else
                    pgbtotalcnt_part = Math.Ceiling(dt.Rows.Count / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt_part)
                '20161014 実行時の進捗表示対応 -add end

                '1行ずつ取得
                Dim row As DataRow
                For Each row In dt.Rows

                    Dim tmp_hash As New Hashtable
                    Dim tmp_solist As New SortedList(Of Integer, String)        '作業用オブジェクト
                    Dim solist_value As New SortedList(Of Integer, String)      'ヘッダー格納用オブジェクト

                    '行データを取得して作業用オブジェクトへ格納
                    For cntjj = 1 To colcnt - 1
                        tmp_solist.Add(cntjj, row(cntjj).ToString)
                    Next

                    '行数から判断する
                    If readrowcnt = 1 Then

                        '1行目はヘッダーなのでヘッダー用のオブジェクトへ格納
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'solist_header = tmp_solist
                        For cntjj = 1 To tmp_solist.Count

                            If tmp_solist(cntjj) <> "" Then
                                solist_header.Add(cntjj, tmp_solist(cntjj))
                            End If

                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                    ElseIf readrowcnt >= BASEMIDFILE_READWRITE_ROW - 1 Then

                        '実データ開始行からデータ格納用オブジェクトへ格納
                        solist_value = tmp_solist

                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'For cntkk = 1 To colcnt - 1
                        '    Dim tmp_baseheader As String = solist_header(cntkk)
                        '    Dim tmp_basevalue As String = solist_value(cntkk)
                        '    tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                        'Next
                        For cntkk = 1 To solist_header.Count
                            Dim tmp_baseheader As String = solist_header(cntkk)
                            Dim tmp_basevalue As String = solist_value(cntkk)
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add sta
                            If tmp_basevalue Is Nothing Then
                                tmp_basevalue = ""
                            End If
                            tmp_basevalue.Replace("'", "’")
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add end
                            If tmp_baseheader <> "" Then
                                tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                            End If
                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                        For cntkbn = 1 To 2

                            Dim insertvalue_sb As New System.Text.StringBuilder
                            Dim chk_sb As New System.Text.StringBuilder
                            Dim insertflg As Boolean = True

                            insertvalue_sb.Append("'" & tmp_hash("物件No") & "'")
                            insertvalue_sb.Append(",'" & tmp_hash("部屋No") & "'")
                            insertvalue_sb.Append(",'" & cntkbn.ToString & "'")

                            Select Case cntkbn
                                Case 1

                                    '空文字チェック
                                    chk_sb.Append(tmp_hash("保証人使用契約者No"))
                                    chk_sb.Append(tmp_hash("保証人No1"))

                                    '空の場合は挿入処理を行わない
                                    If chk_sb.ToString.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("保証人使用契約者No") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("保証人No1") & "'")
                                    End If

                                Case 2

                                    '空文字チェック
                                    chk_sb.Append(tmp_hash("保証人使用契約者No"))
                                    chk_sb.Append(tmp_hash("保証人No2"))

                                    '空の場合は挿入処理を行わない
                                    If chk_sb.ToString.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("保証人使用契約者No") & "'")
                                        insertvalue_sb.Append(",'" & tmp_hash("保証人No2") & "'")
                                    End If

                            End Select

                            '挿入
                            If insertflg Then

                                Dim taisyo As String = "[物件NO],[部屋NO],[並び順],[契約者No],[保証人No]"

                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg sta
                                'qry = ""
                                'qry += " INSERT INTO [" & existmidfilesheetname & "$] "
                                'qry += "(" & taisyo & ")"
                                'qry += " VALUES "
                                'qry += "(" & insertvalue_sb.ToString & ")"

                                'cmd_write.CommandText = qry
                                'cmd_write.ExecuteNonQuery()
                                Dim tmp_cnt As Integer = 0
                                qry = ""
                                qry += " INSERT INTO " & tblname
                                qry += "(" & taisyo & ")"
                                qry += " VALUES "
                                qry += "(" & insertvalue_sb.ToString & ")"
                                DBExec.Exec_NonQuery(sqlcnnv10, qry, tmp_cnt)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg end

                            End If

                        Next

                    End If

                    '20161014 実行時の進捗表示対応 -add sta
                    '-------------------------------
                    'プログレスバー更新/進捗率表示
                    '-------------------------------
                    Dim tmp_pgbcnt As Integer = 0
                    If dt.Rows.Count <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        tmp_pgbcnt = readrowcnt
                    ElseIf dt.Rows.Count > readrowcnt Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        Call EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, tmp_pgbcnt)
                    ElseIf dt.Rows.Count <= readrowcnt Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        tmp_pgbcnt = pgbtotalcnt_part
                    End If

                    '表示
                    If tmp_pgbcnt <> 0 Then
                        Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                        Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                    End If
                    '20161017 進捗表示処理による速度低下の修正 -del
                    'Njc.Frm.MainFrm.Refresh()
                    '20161014 実行時の進捗表示対応 -add end

                    readrowcnt = readrowcnt + 1

                Next

            Catch ex As Exception
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'con_read.Close()
                'con_write.Close()
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                rtn = False
                Return rtn

            End Try
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
            ''クローズ処理
            'con_read.Close()
            'con_write.Close()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
            Return rtn

        End Function

    End Class

#End Region

#Region "契約特約情報"

    Public Class Ky_tokuyaku_mid_Repository

        ''' <summary>
        ''' 契約特約情報コピー処理
        ''' '20161021 既存中間ファイルを無くすことによる速度改善処理2 引数にSQLSERVERへの接続オブジェクト(sqlcnnv10)と仮テーブル名(tblname)およびデータテーブルオブジェクト(dt)を追加
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="dt"></param>
        ''' <param name="tblname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Read_BaseMidFile(ByVal sqlcnnv10 As SqlConnection, ByVal dt As DataTable, ByVal tblname As String) As Boolean

            Dim rtn As Boolean = True
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim basemidfilename As String = "中間ファイル"
            'Dim basemidfilesheetname As String = "契約特約およびメモ情報"
            'Dim existmidfilename As String = "契約情報"
            'Dim existmidfilesheetname As String = "契約特約事項情報"
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            Dim qry As String = ""

            Dim da As New OleDbDataAdapter()
            Dim ds As DataSet = New DataSet()
            'Dim dt As New DataTable()   '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del
            Dim dr As OleDbDataReader
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
            'Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
            'Dim tmpextend As String = "Excel 8.0;HDR=YES;"

            'Dim con_read As New OleDbConnection()
            'Dim cmd_read As New OleDbCommand()
            'Dim con_write As New OleDbConnection()
            'Dim cmd_write As New OleDbCommand()

            'Dim basemidpath As String = EtcMethod.Set_Path(BaseMidDirPath, basemidfilename & ".xlsx")
            'Dim existmidpath As String = EtcMethod.Set_Path(MiddleDirPath, existmidfilename & ".xlsx")
            '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
            '20161014 実行時の進捗表示対応 -add sta
            Dim obj_pgb As New ProgressBarManager
            Dim obj_com As New Njc.Common.CommonRepository
            '20161014 実行時の進捗表示対応 -add end

            Try
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del sta
                ''================================================================================
                '' ●接続設定●
                ''================================================================================
                ''--------------------------------------------------
                '' ①EXCEL接続 (ファイル読込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_read.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & basemidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_read.Connection = con_read

                ''--------------------------------------------------
                '' ②EXCEL接続 (ファイル書込用)
                ''--------------------------------------------------

                ''接続文字列生成
                'con_write.ConnectionString = _
                '    "Provider=" & tmpprovider & _
                '    "Data Source=" & existmidpath & ";" & _
                '    "Extended Properties=" & """" & tmpextend & """"

                ''接続設定
                'cmd_write.Connection = con_write

                ''--------------------------------------------------
                '' 読込→書込
                ''--------------------------------------------------
                ''接続オープン処理
                'con_read.Open()
                'con_write.Open()

                'qry = "SELECT * FROM [" & basemidfilesheetname & "$] "

                'cmd_read = con_read.CreateCommand
                'cmd_read.CommandText = qry
                'dt = New DataTable
                'dr = cmd_read.ExecuteReader
                'dt.Load(dr)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_不要処理削除 -del end
                Dim colcnt As Integer = dt.Columns.Count
                Dim readrowcnt As Integer = 1
                Dim solist_header As New SortedList(Of Integer, String)

                '20161014 実行時の進捗表示対応 -add sta
                'プログレスバー初期化
                Dim pgbtotalcnt_part As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If dt.Rows.Count <= pgbbasecnt Then
                    pgbtotalcnt_part = dt.Rows.Count
                Else
                    pgbtotalcnt_part = Math.Ceiling(dt.Rows.Count / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt_part)
                '20161014 実行時の進捗表示対応 -add end

                '1行ずつ取得
                Dim row As DataRow
                For Each row In dt.Rows

                    Dim tmp_hash As New Hashtable
                    Dim tmp_solist As New SortedList(Of Integer, String)        '作業用オブジェクト
                    Dim solist_value As New SortedList(Of Integer, String)      'ヘッダー格納用オブジェクト

                    '行データを取得して作業用オブジェクトへ格納
                    For cntjj = 1 To colcnt - 1
                        tmp_solist.Add(cntjj, row(cntjj).ToString)
                    Next

                    '行数から判断する
                    If readrowcnt = 1 Then

                        '1行目はヘッダーなのでヘッダー用のオブジェクトへ格納
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'solist_header = tmp_solist
                        For cntjj = 1 To tmp_solist.Count

                            If tmp_solist(cntjj) <> "" Then
                                solist_header.Add(cntjj, tmp_solist(cntjj))
                            End If

                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                    ElseIf readrowcnt >= BASEMIDFILE_READWRITE_ROW - 1 Then

                        '実データ開始行からデータ格納用オブジェクトへ格納
                        solist_value = tmp_solist

                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg sta
                        'For cntkk = 1 To colcnt - 1
                        '    Dim tmp_baseheader As String = solist_header(cntkk)
                        '    Dim tmp_basevalue As String = solist_value(cntkk)
                        '    tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                        'Next
                        For cntkk = 1 To solist_header.Count
                            Dim tmp_baseheader As String = solist_header(cntkk)
                            Dim tmp_basevalue As String = solist_value(cntkk)
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add sta
                            If tmp_basevalue Is Nothing Then
                                tmp_basevalue = ""
                            End If
                            tmp_basevalue.Replace("'", "’")
                            '20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -add end
                            If tmp_baseheader <> "" Then
                                tmp_hash.Add(tmp_baseheader, tmp_basevalue)
                            End If
                        Next
                        '20161007 汎用中間ファイルのゴミデータ残存時の対応 -chg end

                        For cntkbn = 1 To 3

                            Dim insertvalue_sb As New System.Text.StringBuilder
                            Dim chk_str As String
                            Dim insertflg As Boolean = True

                            insertvalue_sb.Append("'" & tmp_hash("物件No") & "'")
                            insertvalue_sb.Append(",'" & tmp_hash("部屋No") & "'")
                            insertvalue_sb.Append(",'" & cntkbn.ToString & "'")

                            Select Case cntkbn
                                Case 1

                                    '空文字チェック
                                    chk_str = IIf(tmp_hash("その他特約内容") Is Nothing, "", tmp_hash("その他特約内容"))

                                    '空の場合は挿入処理を行わない
                                    If chk_str.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("その他特約内容") & "'")
                                    End If

                                Case 2

                                    '空文字チェック
                                    chk_str = IIf(tmp_hash("原状回復特約内容") Is Nothing, "", tmp_hash("原状回復特約内容"))

                                    '空の場合は挿入処理を行わない
                                    If chk_str.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("原状回復特約内容") & "'")
                                    End If

                                Case 3

                                    '空文字チェック
                                    chk_str = IIf(tmp_hash("入居中修繕特約内容") Is Nothing, "", tmp_hash("入居中修繕特約内容"))

                                    '空の場合は挿入処理を行わない
                                    If chk_str.Trim = "" Then
                                        insertflg = False
                                    Else
                                        insertvalue_sb.Append(",'" & tmp_hash("入居中修繕特約内容") & "'")
                                    End If

                            End Select

                            '挿入
                            If insertflg Then

                                Dim taisyo As String = "[物件NO],[部屋NO],[特約区分],[内容1]"

                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg sta
                                'qry = ""
                                'qry += " INSERT INTO [" & existmidfilesheetname & "$] "
                                'qry += "(" & taisyo & ")"
                                'qry += " VALUES "
                                'qry += "(" & insertvalue_sb.ToString & ")"

                                'cmd_write.CommandText = qry
                                'cmd_write.ExecuteNonQuery()
                                Dim tmp_cnt As Integer = 0
                                qry = ""
                                qry += " INSERT INTO " & tblname
                                qry += "(" & taisyo & ")"
                                qry += " VALUES "
                                qry += "(" & insertvalue_sb.ToString & ")"
                                DBExec.Exec_NonQuery(sqlcnnv10, qry, tmp_cnt)
                                '20161021 既存中間ファイルを無くすことによる速度改善処理_個別コピー処理_書込処理修正 -chg end

                            End If

                        Next

                    End If

                    '20161014 実行時の進捗表示対応 -add sta
                    '-------------------------------
                    'プログレスバー更新/進捗率表示
                    '-------------------------------
                    Dim tmp_pgbcnt As Integer = 0
                    If dt.Rows.Count <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        tmp_pgbcnt = readrowcnt
                    ElseIf dt.Rows.Count > readrowcnt Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        Call EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, tmp_pgbcnt)
                    ElseIf dt.Rows.Count <= readrowcnt Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        tmp_pgbcnt = pgbtotalcnt_part
                    End If

                    '表示
                    If tmp_pgbcnt <> 0 Then
                        Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
                        Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
                    End If
                    '20161017 進捗表示処理による速度低下の修正 -del
                    'Njc.Frm.MainFrm.Refresh()
                    '20161014 実行時の進捗表示対応 -add end

                    readrowcnt = readrowcnt + 1

                Next

            Catch ex As Exception
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'con_read.Close()
                'con_write.Close()
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                rtn = False
                Return rtn

            End Try
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
            ''クローズ処理
            'con_read.Close()
            'con_write.Close()
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
            Return rtn

        End Function

    End Class

#End Region

    '20160929 汎用→既存コピー処理改善対応 -add end




    '20160929 汎用→既存コピー処理改善対応 -del sta
    '作り直しのためコメントアウト
    '#Region "部屋駐車場情報"

    '    Public Class Hy_cyusyajo_mid_Repository

    '        ''' <summary>
    '        ''' 汎用用中間ファイル読込→オブジェクトへ格納
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Read_BaseMidFile(ByVal filename As String, ByVal sheetname As String, ByRef model_basemiditem As Object) As Boolean

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値
    '            model_basemiditem = New Njc.Model.Hy_cyusyajo_mid_Model         '汎用用中間ファイルデータ格納用モデル

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If rowcnt = 0 Then
    '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '            'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
    '            Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 1), wsheet.Cells(headerrow, columncnt)).Value
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '            'データ取得
    '            With model_basemiditem

    '                For cntjj = 1 To columncnt
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '                    'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '                    Dim fldvalue As New Object

    '                    '列データを取得して
    '                    fldvalue = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value

    '                    Select Case fldname

    '                        Case "物件NO"
    '                            .ObjCol_物件NO = fldvalue
    '                            .ItemCnt = fldvalue.Length
    '                        Case "部屋NO"
    '                            .ObjCol_部屋NO = fldvalue
    '                        Case "駐車場の空き数(手動設定用)"
    '                            .ObjCol_駐車場の空き数手動設定用 = fldvalue
    '                        Case "駐車場の空き有無(手動設定用)"
    '                            .ObjCol_駐車場の空き有無手動設定用 = fldvalue
    '                        Case "駐車場料金区分(手動設定用)"
    '                            .ObjCol_駐車場料金区分手動設定用 = fldvalue
    '                        Case "駐車場料金(手動設定用)"
    '                            .ObjCol_駐車場料金手動設定用 = fldvalue
    '                        Case "駐車場料金税区分(手動設定用)"
    '                            .ObjCol_駐車場料金税区分手動設定用 = fldvalue
    '                        Case "駐車場の賃貸可能数"
    '                            .ObjCol_駐車場の賃貸可能数 = fldvalue
    '                        Case "バイクの空き数(手動設定用)"
    '                            .ObjCol_バイクの空き数手動設定用 = fldvalue
    '                        Case "バイク駐車場の空き有無(手動設定用)"
    '                            .ObjCol_バイク駐車場の空き有無手動設定用 = fldvalue
    '                        Case "バイク駐車場料金区分(手動設定用)"
    '                            .ObjCol_バイク駐車場料金区分手動設定用 = fldvalue
    '                        Case "バイク駐車場料金(手動設定用)"
    '                            .ObjCol_バイク駐車場料金手動設定用 = fldvalue
    '                        Case "バイク駐車場料金税区分(手動設定用)"
    '                            .ObjCol_バイク駐車場料金税区分手動設定用 = fldvalue
    '                        Case "バイク駐車場の賃貸可能数"
    '                            .ObjCol_バイク駐車場の賃貸可能数 = fldvalue
    '                        Case "駐輪場駐車場の空き数(手動設定用)"
    '                            .ObjCol_駐輪場駐車場の空き数手動設定用 = fldvalue
    '                        Case "駐輪場駐車場の空き有無(手動設定用)"
    '                            .ObjCol_駐輪場駐車場の空き有無手動設定用 = fldvalue
    '                        Case "駐輪場駐車場料金区分(手動設定用)"
    '                            .ObjCol_駐輪場駐車場料金区分手動設定用 = fldvalue
    '                        Case "駐輪場駐車場料金(手動設定用)"
    '                            .ObjCol_駐輪場駐車場料金手動設定用 = fldvalue
    '                        Case "駐輪場駐車場料金税区分(手動設定用)"
    '                            .ObjCol_駐輪場駐車場料金税区分手動設定用 = fldvalue
    '                        Case "駐輪場駐車場の賃貸可能数"
    '                            .ObjCol_駐輪場駐車場の賃貸可能数 = fldvalue

    '                    End Select

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '        ''' <summary>
    '        ''' 汎用用中間ファイルデータ→既存用中間ファイルへ書込み
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Set_ExistMidFile(ByVal filename As String, ByVal sheetname As String, ByVal model_basemiditem As Object)

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If model_basemiditem.ItemCnt = 0 Then
    '                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '            '書込開始行初期値設定
    '            Dim writestartrow As Integer = startrow


    '            'データ取得
    '            With model_basemiditem

    '                For cntii = 1 To 3

    '                    '書込行数セット
    '                    Dim writerowcnt As Integer = model_basemiditem.ItemCnt

    '                    '書込みデータの振り分け
    '                    Dim tmp_akisu As New Object
    '                    Dim tmp_akiumu As New Object
    '                    Dim tmp_ryokinkbn As New Object
    '                    Dim tmp_ryokin As New Object
    '                    Dim tmp_ryokinzeikbn As New Object
    '                    Dim tmp_tintaikanosu As New Object

    '                    Select Case cntii
    '                        Case 1
    '                            tmp_akisu = .ObjCol_駐車場の空き数手動設定用
    '                            tmp_akiumu = .ObjCol_駐車場の空き有無手動設定用
    '                            tmp_ryokinkbn = .ObjCol_駐車場料金区分手動設定用
    '                            tmp_ryokin = .ObjCol_駐車場料金手動設定用
    '                            tmp_ryokinzeikbn = .ObjCol_駐車場料金税区分手動設定用
    '                            tmp_tintaikanosu = .ObjCol_駐車場の賃貸可能数
    '                        Case 2
    '                            tmp_akisu = .ObjCol_バイクの空き数手動設定用
    '                            tmp_akiumu = .ObjCol_バイク駐車場の空き有無手動設定用
    '                            tmp_ryokinkbn = .ObjCol_バイク駐車場料金区分手動設定用
    '                            tmp_ryokin = .ObjCol_バイク駐車場料金手動設定用
    '                            tmp_ryokinzeikbn = .ObjCol_バイク駐車場料金税区分手動設定用
    '                            tmp_tintaikanosu = .ObjCol_バイク駐車場の賃貸可能数
    '                        Case 3
    '                            tmp_akisu = .ObjCol_駐輪場駐車場の空き数手動設定用
    '                            tmp_akiumu = .ObjCol_駐輪場駐車場の空き有無手動設定用
    '                            tmp_ryokinkbn = .ObjCol_駐輪場駐車場料金区分手動設定用
    '                            tmp_ryokin = .ObjCol_駐輪場駐車場料金手動設定用
    '                            tmp_ryokinzeikbn = .ObjCol_駐輪場駐車場料金税区分手動設定用
    '                            tmp_tintaikanosu = .ObjCol_駐輪場駐車場の賃貸可能数
    '                    End Select

    '                    '書込処理
    '                    For cntjj = 1 To columncnt

    '                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim

    '                        Select Case fldname

    '                            Case "物件NO"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_物件NO
    '                            Case "部屋NO"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_部屋NO
    '                            Case "駐車場区分"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = cntii.ToString
    '                            Case "駐車場の空き数(手動設定用)"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_akisu
    '                            Case "駐車場の空き有無(手動設定用)"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_akiumu
    '                            Case "駐車場料金区分(手動設定用)"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_ryokinkbn
    '                            Case "駐車場料金(手動設定用)"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_ryokin
    '                            Case "駐車場料金税区分(手動設定用)"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_ryokinzeikbn
    '                            Case "駐車場の賃貸可能数"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_tintaikanosu
    '                        End Select

    '                    Next

    '                    writestartrow = (writestartrow - 1) + writerowcnt + 1

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '    End Class

    '#End Region

    '#Region "部屋特約情報"

    '    Public Class Hy_tokuyaku_mid_Repository

    '        ''' <summary>
    '        ''' 汎用用中間ファイル読込→オブジェクトへ格納
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Read_BaseMidFile(ByVal filename As String, ByVal sheetname As String, ByRef model_basemiditem As Object) As Boolean

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値
    '            model_basemiditem = New Njc.Model.Hy_tokuyaku_mid_Model         '汎用用中間ファイルデータ格納用モデル

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If rowcnt = 0 Then
    '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '            'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
    '            Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 1), wsheet.Cells(headerrow, columncnt)).Value
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '            'データ取得
    '            With model_basemiditem

    '                For cntjj = 1 To columncnt
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '                    'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '                    Dim fldvalue As New Object

    '                    '列データを取得して
    '                    fldvalue = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value

    '                    Select Case fldname

    '                        Case "物件NO"
    '                            .ObjCol_物件NO = fldvalue
    '                            .ItemCnt = fldvalue.Length
    '                        Case "部屋NO"
    '                            .ObjCol_部屋NO = fldvalue
    '                        Case "原状回復特約内容"
    '                            .ObjCol_原状回復特約内容 = fldvalue
    '                        Case "入居中修繕特約内容"
    '                            .ObjCol_入居中修繕特約内容 = fldvalue
    '                        Case "その他特約内容"
    '                            .ObjCol_その他特約内容 = fldvalue
    '                    End Select

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '        ''' <summary>
    '        ''' 汎用用中間ファイルデータ→既存用中間ファイルへ書込み
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Set_ExistMidFile(ByVal filename As String, ByVal sheetname As String, ByVal model_basemiditem As Object)

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If model_basemiditem.ItemCnt = 0 Then
    '                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '            '書込開始行初期値設定
    '            Dim writestartrow As Integer = startrow


    '            'データ取得
    '            With model_basemiditem

    '                For cntii = 1 To 3

    '                    '書込行数セット
    '                    Dim writerowcnt As Integer = model_basemiditem.ItemCnt

    '                    '書込みデータの振り分け
    '                    Dim tmp_tokuyaku As New Object

    '                    Select Case cntii
    '                        Case 1
    '                            tmp_tokuyaku = .ObjCol_その他特約内容
    '                        Case 2
    '                            tmp_tokuyaku = .ObjCol_原状回復特約内容
    '                        Case 3
    '                            tmp_tokuyaku = .ObjCol_入居中修繕特約内容
    '                    End Select

    '                    '書込処理
    '                    For cntjj = 1 To columncnt

    '                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim

    '                        Select Case fldname

    '                            Case "物件NO"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_物件NO
    '                            Case "部屋NO"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_部屋NO
    '                            Case "特約区分"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = cntii.ToString
    '                            Case "内容1"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_tokuyaku
    '                        End Select

    '                    Next

    '                    writestartrow = (writestartrow - 1) + writerowcnt + 1

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '    End Class

    '#End Region

    '#Region "部屋契約解約確認事項情報"

    '    Public Class Hy_kykaikakuninjiko_mid_Repository

    '        ''' <summary>
    '        ''' 汎用用中間ファイル読込→オブジェクトへ格納
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Read_BaseMidFile(ByVal filename As String, ByVal sheetname As String, ByRef model_basemiditem As Object) As Boolean

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値
    '            model_basemiditem = New Njc.Model.Hy_kykaikakuninjiko_mid_Model '汎用用中間ファイルデータ格納用モデル

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If rowcnt = 0 Then
    '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '            'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
    '            Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 1), wsheet.Cells(headerrow, columncnt)).Value
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '            'データ取得
    '            With model_basemiditem

    '                For cntjj = 1 To columncnt
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '                    'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '                    Dim fldvalue As New Object

    '                    '列データを取得して
    '                    fldvalue = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value

    '                    Select Case fldname
    '                        Case "物件NO"
    '                            .ObjCol_物件NO = fldvalue
    '                            .ItemCnt = fldvalue.Length
    '                        Case "部屋NO"
    '                            .ObjCol_部屋NO = fldvalue
    '                        Case "他業者による客付け時（広告料の相殺処理）_募集"
    '                            .ObjCol_他業者による客付け時広告料の相殺処理募集 = fldvalue
    '                        Case "賃料交渉"
    '                            .ObjCol_賃料交渉 = fldvalue
    '                        Case "案内時の照明器具"
    '                            .ObjCol_案内時の照明器具 = fldvalue
    '                        Case "空室の清掃"
    '                            .ObjCol_空室の清掃 = fldvalue
    '                        Case "案内時の営業車駐車可能スペース"
    '                            .ObjCol_案内時の営業車駐車可能スペース = fldvalue
    '                        Case "入居申込書"
    '                            .ObjCol_入居申込書 = fldvalue
    '                        Case "入居前に必要な工事箇所"
    '                            .ObjCol_入居前に必要な工事箇所 = fldvalue
    '                        Case "入居時立会い"
    '                            .ObjCol_入居時立会い = fldvalue
    '                        Case "他業者による客付け時の重要事項説明の確認"
    '                            .ObjCol_他業者による客付け時の重要事項説明の確認 = fldvalue
    '                        Case "他業者による客付け時（広告料の相殺処理）_契約"
    '                            .ObjCol_他業者による客付け時広告料の相殺処理契約 = fldvalue
    '                        Case "家主への契約金支払い"
    '                            .ObjCol_家主への契約金支払い = fldvalue
    '                        Case "敷金預かり先"
    '                            .ObjCol_敷金預かり先 = fldvalue
    '                        Case "広告料"
    '                            .ObjCol_広告料 = fldvalue
    '                        Case "解約違約金　（予告と短期）"
    '                            .ObjCol_解約違約金予告と短期 = fldvalue
    '                        Case "退去時の立会い"
    '                            .ObjCol_退去時の立会い = fldvalue
    '                    End Select

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '        ''' <summary>
    '        ''' 汎用用中間ファイルデータ→既存用中間ファイルへ書込み
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Set_ExistMidFile(ByVal filename As String, ByVal sheetname As String, ByVal model_basemiditem As Object)

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If model_basemiditem.ItemCnt = 0 Then
    '                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '            '書込開始行初期値設定
    '            Dim writestartrow As Integer = startrow

    '            'データ取得
    '            With model_basemiditem

    '                For cntii = 1 To 4

    '                    '書込行数セット
    '                    Dim writerowcnt As Integer = model_basemiditem.ItemCnt

    '                    '書込みデータの振り分け
    '                    Dim tmp_tokuyaku As New Object

    '                    '各区分毎の最大行Noセット
    '                    Dim maxcntinkbn As Integer = 0
    '                    Select Case cntii
    '                        Case 1
    '                            maxcntinkbn = 5
    '                        Case 2
    '                            maxcntinkbn = 5
    '                        Case 3
    '                            maxcntinkbn = 3
    '                        Case 4
    '                            maxcntinkbn = 2
    '                    End Select

    '                    '各区分毎の書込みデータセット
    '                    Dim tmp_naiyo As New Object
    '                    For cntjj = 1 To maxcntinkbn
    '                        Select Case cntii
    '                            Case 1
    '                                Select Case cntjj
    '                                    Case 1
    '                                        tmp_naiyo = .ObjCol_他業者による客付け時広告料の相殺処理募集
    '                                    Case 2
    '                                        tmp_naiyo = .ObjCol_賃料交渉
    '                                    Case 3
    '                                        tmp_naiyo = .ObjCol_案内時の照明器具
    '                                    Case 4
    '                                        tmp_naiyo = .ObjCol_空室の清掃
    '                                    Case 5
    '                                        tmp_naiyo = .ObjCol_案内時の営業車駐車可能スペース
    '                                End Select
    '                            Case 2
    '                                Select Case cntjj
    '                                    Case 1
    '                                        tmp_naiyo = .ObjCol_入居申込書
    '                                    Case 2
    '                                        tmp_naiyo = .ObjCol_入居前に必要な工事箇所
    '                                    Case 3
    '                                        tmp_naiyo = .ObjCol_入居時立会い
    '                                    Case 4
    '                                        tmp_naiyo = .ObjCol_他業者による客付け時の重要事項説明の確認
    '                                    Case 5
    '                                        tmp_naiyo = .ObjCol_他業者による客付け時広告料の相殺処理契約
    '                                End Select
    '                            Case 3
    '                                Select Case cntjj
    '                                    Case 1
    '                                        tmp_naiyo = .ObjCol_家主への契約金支払い
    '                                    Case 2
    '                                        tmp_naiyo = .ObjCol_敷金預かり先
    '                                    Case 3
    '                                        tmp_naiyo = .ObjCol_広告料
    '                                End Select
    '                            Case 4
    '                                Select Case cntjj
    '                                    Case 1
    '                                        tmp_naiyo = .ObjCol_解約違約金予告と短期
    '                                    Case 2
    '                                        tmp_naiyo = .ObjCol_退去時の立会い
    '                                End Select
    '                        End Select

    '                        '書込処理
    '                        For cntkk = 1 To columncnt

    '                            Dim fldname As String = headervalue(startrow - 1, cntkk).ToString.Trim

    '                            Select Case fldname

    '                                Case "物件NO"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = .ObjCol_物件NO
    '                                Case "部屋NO"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = .ObjCol_部屋NO
    '                                Case "確認事項区分"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = cntii.ToString
    '                                Case "確認事項No"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = cntjj.ToString
    '                                Case "内容"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = tmp_naiyo
    '                            End Select

    '                        Next

    '                        writestartrow = (writestartrow - 1) + writerowcnt + 1

    '                    Next

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '    End Class

    '#End Region

    '#Region "物件部屋所有者情報"

    '    Public Class Bkhy_syoyusya_mid_Repository

    '        ''' <summary>
    '        ''' 汎用用中間ファイル読込→オブジェクトへ格納
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Read_BaseMidFile(ByVal filename As String, ByVal sheetname As String, ByRef model_basemiditem As Object) As Boolean

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値
    '            model_basemiditem = New Njc.Model.Hy_kykaikakuninjiko_mid_Model '汎用用中間ファイルデータ格納用モデル

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If rowcnt = 0 Then
    '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '            'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
    '            Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 1), wsheet.Cells(headerrow, columncnt)).Value
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '            'データ取得
    '            With model_basemiditem

    '                For cntjj = 1 To columncnt
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '                    'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '                    Dim fldvalue As New Object

    '                    '列データを取得して
    '                    fldvalue = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value

    '                    Select Case fldname
    '                        Case "物件NO"
    '                            .ObjCol_物件NO = fldvalue
    '                            .ItemCnt = fldvalue.Length
    '                        Case "部屋NO"
    '                            .ObjCol_部屋NO = fldvalue
    '                        Case "他業者による客付け時（広告料の相殺処理）_募集"
    '                            .ObjCol_他業者による客付け時広告料の相殺処理募集 = fldvalue
    '                        Case "賃料交渉"
    '                            .ObjCol_賃料交渉 = fldvalue
    '                        Case "案内時の照明器具"
    '                            .ObjCol_案内時の照明器具 = fldvalue
    '                        Case "空室の清掃"
    '                            .ObjCol_空室の清掃 = fldvalue
    '                        Case "案内時の営業車駐車可能スペース"
    '                            .ObjCol_案内時の営業車駐車可能スペース = fldvalue
    '                        Case "入居申込書"
    '                            .ObjCol_入居申込書 = fldvalue
    '                        Case "入居前に必要な工事箇所"
    '                            .ObjCol_入居前に必要な工事箇所 = fldvalue
    '                        Case "入居時立会い"
    '                            .ObjCol_入居時立会い = fldvalue
    '                        Case "他業者による客付け時の重要事項説明の確認"
    '                            .ObjCol_他業者による客付け時の重要事項説明の確認 = fldvalue
    '                        Case "他業者による客付け時（広告料の相殺処理）_契約"
    '                            .ObjCol_他業者による客付け時広告料の相殺処理契約 = fldvalue
    '                        Case "家主への契約金支払い"
    '                            .ObjCol_家主への契約金支払い = fldvalue
    '                        Case "敷金預かり先"
    '                            .ObjCol_敷金預かり先 = fldvalue
    '                        Case "広告料"
    '                            .ObjCol_広告料 = fldvalue
    '                        Case "解約違約金　（予告と短期）"
    '                            .ObjCol_解約違約金予告と短期 = fldvalue
    '                        Case "退去時の立会い"
    '                            .ObjCol_退去時の立会い = fldvalue
    '                    End Select

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '        ''' <summary>
    '        ''' 汎用用中間ファイルデータ→既存用中間ファイルへ書込み
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Set_ExistMidFile(ByVal filename As String, ByVal sheetname As String, ByVal model_basemiditem As Object)

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If model_basemiditem.ItemCnt = 0 Then
    '                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '            '書込開始行初期値設定
    '            Dim writestartrow As Integer = startrow

    '            'データ取得
    '            With model_basemiditem

    '                For cntii = 1 To 4

    '                    '書込行数セット
    '                    Dim writerowcnt As Integer = model_basemiditem.ItemCnt

    '                    '書込みデータの振り分け
    '                    Dim tmp_tokuyaku As New Object

    '                    '各区分毎の最大行Noセット
    '                    Dim maxcntinkbn As Integer = 0
    '                    Select Case cntii
    '                        Case 1
    '                            maxcntinkbn = 5
    '                        Case 2
    '                            maxcntinkbn = 5
    '                        Case 3
    '                            maxcntinkbn = 3
    '                        Case 4
    '                            maxcntinkbn = 2
    '                    End Select

    '                    '各区分毎の書込みデータセット
    '                    Dim tmp_naiyo As New Object
    '                    For cntjj = 1 To maxcntinkbn
    '                        Select Case cntii
    '                            Case 1
    '                                Select Case cntjj
    '                                    Case 1
    '                                        tmp_naiyo = .ObjCol_他業者による客付け時広告料の相殺処理募集
    '                                    Case 2
    '                                        tmp_naiyo = .ObjCol_賃料交渉
    '                                    Case 3
    '                                        tmp_naiyo = .ObjCol_案内時の照明器具
    '                                    Case 4
    '                                        tmp_naiyo = .ObjCol_空室の清掃
    '                                    Case 5
    '                                        tmp_naiyo = .ObjCol_案内時の営業車駐車可能スペース
    '                                End Select
    '                            Case 2
    '                                Select Case cntjj
    '                                    Case 1
    '                                        tmp_naiyo = .ObjCol_入居申込書
    '                                    Case 2
    '                                        tmp_naiyo = .ObjCol_入居前に必要な工事箇所
    '                                    Case 3
    '                                        tmp_naiyo = .ObjCol_入居時立会い
    '                                    Case 4
    '                                        tmp_naiyo = .ObjCol_他業者による客付け時の重要事項説明の確認
    '                                    Case 5
    '                                        tmp_naiyo = .ObjCol_他業者による客付け時広告料の相殺処理契約
    '                                End Select
    '                            Case 3
    '                                Select Case cntjj
    '                                    Case 1
    '                                        tmp_naiyo = .ObjCol_家主への契約金支払い
    '                                    Case 2
    '                                        tmp_naiyo = .ObjCol_敷金預かり先
    '                                    Case 3
    '                                        tmp_naiyo = .ObjCol_広告料
    '                                End Select
    '                            Case 4
    '                                Select Case cntjj
    '                                    Case 1
    '                                        tmp_naiyo = .ObjCol_解約違約金予告と短期
    '                                    Case 2
    '                                        tmp_naiyo = .ObjCol_退去時の立会い
    '                                End Select
    '                        End Select

    '                        '書込処理
    '                        For cntkk = 1 To columncnt

    '                            Dim fldname As String = headervalue(startrow - 1, cntkk).ToString.Trim

    '                            Select Case fldname

    '                                Case "物件NO"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = .ObjCol_物件NO
    '                                Case "部屋NO"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = .ObjCol_部屋NO
    '                                Case "確認事項区分"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = cntii.ToString
    '                                Case "確認事項No"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = cntjj.ToString
    '                                Case "内容"
    '                                    wsheet.Range(wsheet.Cells(writestartrow, cntkk), wsheet.Cells((writestartrow - 1) + writerowcnt, cntkk)).Value = tmp_naiyo
    '                            End Select

    '                        Next

    '                        writestartrow = (writestartrow - 1) + writerowcnt + 1

    '                    Next

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '    End Class

    '#End Region

    '#Region "契約契約者情報"

    '    Public Class Ky_kys_mid_Repository

    '        ''' <summary>
    '        ''' 汎用用中間ファイル読込→オブジェクトへ格納
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Read_BaseMidFile(ByVal filename As String, ByVal sheetname As String, ByRef model_basemiditem As Object) As Boolean

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値
    '            model_basemiditem = New Njc.Model.Ky_kys_mid_Model              '汎用用中間ファイルデータ格納用モデル

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If rowcnt = 0 Then
    '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '            'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
    '            Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 1), wsheet.Cells(headerrow, columncnt)).Value
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '            'データ取得
    '            With model_basemiditem

    '                For cntjj = 1 To columncnt
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '                    'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '                    Dim fldvalue As New Object

    '                    '列データを取得して
    '                    fldvalue = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value

    '                    Select Case fldname

    '                        Case "物件No"
    '                            .ObjCol_物件No = fldvalue
    '                            .ItemCnt = fldvalue.Length
    '                        Case "部屋No"
    '                            .ObjCol_部屋No = fldvalue
    '                        Case "契約No"
    '                            .ObjCol_契約No = fldvalue
    '                        Case "契約管理レコードNo"
    '                            .ObjCol_契約管理レコードNo = fldvalue
    '                        Case "契約者No1"
    '                            .ObjCol_契約者No1 = fldvalue
    '                        Case "入居フラグ1"
    '                            .ObjCol_入居フラグ1 = fldvalue
    '                        Case "契約者No2"
    '                            .ObjCol_契約者No2 = fldvalue
    '                        Case "入居フラグ2"
    '                            .ObjCol_入居フラグ2 = fldvalue
    '                        Case "契約者No3"
    '                            .ObjCol_契約者No3 = fldvalue
    '                        Case "入居フラグ3"
    '                            .ObjCol_入居フラグ3 = fldvalue
    '                    End Select

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '        ''' <summary>
    '        ''' 汎用用中間ファイルデータ→既存用中間ファイルへ書込み
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Set_ExistMidFile(ByVal filename As String, ByVal sheetname As String, ByVal model_basemiditem As Object)

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If model_basemiditem.ItemCnt = 0 Then
    '                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '            '書込開始行初期値設定
    '            Dim writestartrow As Integer = startrow


    '            'データ取得
    '            With model_basemiditem

    '                For cntii = 1 To 3

    '                    '書込行数セット
    '                    Dim writerowcnt As Integer = model_basemiditem.ItemCnt

    '                    '書込みデータの振り分け
    '                    Dim tmp_kysno As New Object
    '                    Dim tmp_nyukyoflg As New Object

    '                    Select Case cntii
    '                        Case 1
    '                            tmp_kysno = .ObjCol_契約者No1
    '                            tmp_nyukyoflg = .ObjCol_入居フラグ1
    '                        Case 2
    '                            tmp_kysno = .ObjCol_契約者No2
    '                            tmp_nyukyoflg = .ObjCol_入居フラグ2
    '                        Case 3
    '                            tmp_kysno = .ObjCol_契約者No3
    '                            tmp_nyukyoflg = .ObjCol_入居フラグ3
    '                    End Select

    '                    '書込処理
    '                    For cntjj = 1 To columncnt

    '                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim

    '                        Select Case fldname

    '                            Case "物件No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_物件No
    '                            Case "部屋No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_部屋No
    '                            Case "契約No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_契約No
    '                            Case "契約管理レコードNo"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_契約管理レコードNo
    '                            Case "並び順No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = cntii.ToString
    '                            Case "契約者No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_kysno
    '                            Case "入居フラグ"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_nyukyoflg

    '                        End Select

    '                    Next

    '                    writestartrow = (writestartrow - 1) + writerowcnt + 1

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '    End Class

    '#End Region

    '#Region "契約保証人情報"

    '    Public Class Ky_hosyonin_mid_Repository

    '        ''' <summary>
    '        ''' 汎用用中間ファイル読込→オブジェクトへ格納
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Read_BaseMidFile(ByVal filename As String, ByVal sheetname As String, ByRef model_basemiditem As Object) As Boolean

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値
    '            model_basemiditem = New Njc.Model.Ky_hosyonin_mid_Model         '汎用用中間ファイルデータ格納用モデル

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If rowcnt = 0 Then
    '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '            'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
    '            Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 1), wsheet.Cells(headerrow, columncnt)).Value
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '            'データ取得
    '            With model_basemiditem

    '                For cntjj = 1 To columncnt
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '                    'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '                    Dim fldvalue As New Object

    '                    '列データを取得して
    '                    fldvalue = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value

    '                    Select Case fldname

    '                        Case "物件No"
    '                            .ObjCol_物件No = fldvalue
    '                            .ItemCnt = fldvalue.Length
    '                        Case "部屋No"
    '                            .ObjCol_部屋No = fldvalue
    '                        Case "契約No"
    '                            .ObjCol_契約No = fldvalue
    '                        Case "契約管理レコードNo"
    '                            .ObjCol_契約管理レコードNo = fldvalue
    '                        Case "保証人使用契約者No"
    '                            .ObjCol_保証人使用契約者No = fldvalue
    '                        Case "保証人No1"
    '                            .ObjCol_保証人No1 = fldvalue
    '                        Case "保証人No2"
    '                            .ObjCol_保証人No2 = fldvalue

    '                    End Select

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '        ''' <summary>
    '        ''' 汎用用中間ファイルデータ→既存用中間ファイルへ書込み
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Set_ExistMidFile(ByVal filename As String, ByVal sheetname As String, ByVal model_basemiditem As Object)

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If model_basemiditem.ItemCnt = 0 Then
    '                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '            '書込開始行初期値設定
    '            Dim writestartrow As Integer = startrow


    '            'データ取得
    '            With model_basemiditem

    '                For cntii = 1 To 2

    '                    '書込行数セット
    '                    Dim writerowcnt As Integer = model_basemiditem.ItemCnt

    '                    '書込みデータの振り分け
    '                    Dim tmp_kysno As New Object
    '                    Dim tmp_hosyoninno As New Object

    '                    Select Case cntii
    '                        Case 1
    '                            tmp_kysno = .ObjCol_保証人使用契約者No
    '                            tmp_hosyoninno = .ObjCol_保証人No1
    '                        Case 2
    '                            tmp_kysno = .ObjCol_保証人使用契約者No
    '                            tmp_hosyoninno = .ObjCol_保証人No2
    '                    End Select

    '                    '書込処理
    '                    For cntjj = 1 To columncnt

    '                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim

    '                        Select Case fldname

    '                            Case "物件No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_物件No
    '                            Case "部屋No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_部屋No
    '                            Case "契約No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_契約No
    '                            Case "契約管理レコードNo"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_契約管理レコードNo
    '                            Case "並び順"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = cntii.ToString
    '                            Case "契約者No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_kysno
    '                            Case "保証人No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_hosyoninno

    '                        End Select

    '                    Next

    '                    writestartrow = (writestartrow - 1) + writerowcnt + 1

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '    End Class

    '#End Region

    '#Region "契約特約情報"

    '    Public Class Ky_tokuyaku_mid_Repository

    '        ''' <summary>
    '        ''' 汎用用中間ファイル読込→オブジェクトへ格納
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Read_BaseMidFile(ByVal filename As String, ByVal sheetname As String, ByRef model_basemiditem As Object) As Boolean

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値
    '            model_basemiditem = New Njc.Model.Ky_tokuyaku_mid_Model         '汎用用中間ファイルデータ格納用モデル

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If rowcnt = 0 Then
    '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '            'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
    '            Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 1), wsheet.Cells(headerrow, columncnt)).Value
    '            '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '            'データ取得
    '            With model_basemiditem

    '                For cntjj = 1 To columncnt
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg sta
    '                    'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                    Dim fldname As String = headervalue(1, cntjj).ToString.Trim
    '                    '20160913_2 中間ファイル新規作成分の修正を個別コピー処理へ反映 -chg end
    '                    Dim fldvalue As New Object

    '                    '列データを取得して
    '                    fldvalue = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value

    '                    Select Case fldname

    '                        Case "物件No"
    '                            .ObjCol_物件No = fldvalue
    '                            .ItemCnt = fldvalue.Length
    '                        Case "部屋No"
    '                            .ObjCol_部屋No = fldvalue
    '                        Case "契約No"
    '                            .ObjCol_契約No = fldvalue
    '                        Case "契約管理レコードNo"
    '                            .ObjCol_契約管理レコードNo = fldvalue
    '                        Case "原状回復特約内容"
    '                            .ObjCol_原状回復特約内容 = fldvalue
    '                        Case "入居中修繕特約内容"
    '                            .ObjCol_入居中修繕特約内容 = fldvalue
    '                        Case "その他特約内容"
    '                            .ObjCol_その他特約内容 = fldvalue

    '                    End Select

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '        ''' <summary>
    '        ''' 汎用用中間ファイルデータ→既存用中間ファイルへ書込み
    '        ''' </summary>
    '        ''' <param name="filename"></param>
    '        ''' <param name="sheetname"></param>
    '        ''' <param name="model_basemiditem"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Set_ExistMidFile(ByVal filename As String, ByVal sheetname As String, ByVal model_basemiditem As Object)

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
    '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '            Dim rtn As Boolean = True                                       '戻り値

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add sta
    '            If model_basemiditem.ItemCnt = 0 Then
    '                Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
    '                Return rtn
    '            End If
    '            '20160905 中間ファイルが空の場合の個別コピー処理修正 -add end
    '            'ヘッダー行取得
    '            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '            '書込開始行初期値設定
    '            Dim writestartrow As Integer = startrow


    '            'データ取得
    '            With model_basemiditem

    '                For cntii = 1 To 3

    '                    '書込行数セット
    '                    Dim writerowcnt As Integer = model_basemiditem.ItemCnt

    '                    '書込みデータの振り分け
    '                    Dim tmp_tokuyaku As New Object

    '                    Select Case cntii
    '                        Case 1
    '                            tmp_tokuyaku = .ObjCol_その他特約内容
    '                        Case 2
    '                            tmp_tokuyaku = .ObjCol_原状回復特約内容
    '                        Case 3
    '                            tmp_tokuyaku = .ObjCol_入居中修繕特約内容
    '                    End Select

    '                    '書込処理
    '                    For cntjj = 1 To columncnt

    '                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim

    '                        Select Case fldname

    '                            Case "物件No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_物件No
    '                            Case "部屋No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_部屋No
    '                            Case "契約No"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_契約No
    '                            Case "契約管理レコードNo"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = .ObjCol_契約管理レコードNo
    '                            Case "特約区分"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = cntii.ToString
    '                            Case "内容1"
    '                                wsheet.Range(wsheet.Cells(writestartrow, cntjj), wsheet.Cells((writestartrow - 1) + writerowcnt, cntjj)).Value = tmp_tokuyaku

    '                        End Select

    '                    Next

    '                    writestartrow = (writestartrow - 1) + writerowcnt + 1

    '                Next

    '            End With

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

    '            Return rtn

    '        End Function

    '    End Class

    '#End Region
    '20160929 汎用→既存コピー処理改善対応 -del end

End Namespace
