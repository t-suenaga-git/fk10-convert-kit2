Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "部屋基本情報"

    Public Class Hydata_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_Model                  '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '20160720 連動情報構築 -chg sta
                'Call Me.Get_RelData()
                Call SetRelItemToObject.Set_RelData_Hyrui()
                '20160720 連動情報構築 -chg end
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'テーブル名/フィールド名セット
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "hydata"
                '20161012 革命10アップデートに伴う修正 -add (最後尾にwmp_idを追加)
                '20160519 EXEUpdateに伴う修正 部屋基本情報 -add (最後尾にsort_hy_no、lastupdateを追加)
                Dim fldnamegrp As String = "hy_guid,bk_guid,hy_no,hy_gaibuno,hy_deleteflg," & _
                                           "delete_guid,delete_day,delete_cnt,hy_ruinokbn,madori_cnt," & _
                                           "madori_typekbn,madori_biko,men_senyujitu,men_senyujitutubo,men_yuka," & _
                                           "men_yukatubo,men_senyutouki,men_senyutoukitubo,men_toki,men_tokitubo," & _
                                           "men_balcony,men_balconytubo,men_tempo,men_tempotubo,men_jutaku," & _
                                           "men_jutakutubo,history,rowid,delete_cause,kaiyaku_uketukekbn," & _
                                           "kaiyaku_months,kaiyaku_days,kaiyaku_day,sort_hy_no,lastupdate," & _
                                           "wmp_id"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub As String = readtbl.Columns(keycol_sub - 1).ColumnName.Trim

                    For cntii As Integer = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""
                        Dim tmp_madori As String = ""

                        'キー値取得
                        Dim fldvalue_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim fldvalue_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub - 1)).Trim
                        Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
                                Case "部屋NO"
                                    .Vari_Hy_no = fldvalue.Trim
                                Case "部屋分類名"
                                    .Vari_Hy_ruinokbn = fldvalue.Trim
                                Case "間取り"
                                    tmp_madori = fldvalue.Trim
                                Case "間取り備考"
                                    .Vari_Madori_biko = fldvalue.Trim
                                Case "専有実面積(m2)"
                                    .Vari_Men_senyujitu = fldvalue.Trim
                                Case "専有実面積(坪)"
                                    .Vari_Men_senyujitutubo = fldvalue.Trim
                                Case "床面積(m2)"
                                    .Vari_Men_yuka = fldvalue.Trim
                                Case "床面積(坪)"
                                    .Vari_Men_yukatubo = fldvalue.Trim
                                Case "専有登記面積(m2)"
                                    .Vari_Men_senyutouki = fldvalue.Trim
                                Case "専有登記面積(坪)"
                                    .Vari_Men_senyutoukitubo = fldvalue.Trim
                                Case "登記延床面積(m2)"
                                    .Vari_Men_toki = fldvalue.Trim
                                Case "登記延床面積(坪)"
                                    .Vari_Men_tokitubo = fldvalue.Trim
                                Case "バルコニー面積(m2)"
                                    .Vari_Men_balcony = fldvalue.Trim
                                Case "バルコニー面積(坪)"
                                    .Vari_Men_balconytubo = fldvalue.Trim
                                Case "店舗付き住宅店舗部分面積(m2)"
                                    .Vari_Men_tempo = fldvalue.Trim
                                    .Vari_Men_tempotubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(fldvalue.Trim)), RoundingTypes.Kirisute, 2)
                                Case "店舗付き住宅店舗部分面積(坪)"
                                    'm2を変換
                                Case "店舗付き住宅住宅部分面積(m2)"
                                    .Vari_Men_jutaku = fldvalue.Trim
                                    .Vari_Men_jutakutubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(fldvalue.Trim)), RoundingTypes.Kirisute, 2)
                                Case "店舗付き住宅住宅部分面積(坪)"
                                    'm2を変換
                                Case "解約受付区分"
                                    .Vari_Kaiyaku_uketukekbn = fldvalue.Trim
                                Case "解約受付月数"
                                    .Vari_Kaiyaku_months = fldvalue.Trim
                                Case "解約受付日数"
                                    .Vari_Kaiyaku_days = fldvalue.Trim
                                Case "解約受付日にち"
                                    .Vari_Kaiyaku_day = fldvalue.Trim
                                Case "自社Web内部キー採番ID"    '20161012 革命10アップデートに伴う修正 -add
                                    .Vari_Wmp_id = fldvalue.Trim
                            End Select

                        Next
                        '20161004 汎用コンバート時の平米→坪変換処理の追加 -add sta
                        '汎用の場合は平米から坪を算出する
                        If CNVNO = ConvertTypes._汎用 Then
                            .Vari_Men_senyujitutubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_senyujitu)), RoundingTypes.Kirisute, 2)
                            .Vari_Men_yukatubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_yuka)), RoundingTypes.Kirisute, 2)
                            .Vari_Men_balconytubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_balcony)), RoundingTypes.Kirisute, 2)
                        End If
                        '20161004 汎用コンバート時の平米→坪変換処理の追加 -add end
                        '間取変換
                        Dim tmp_madoricnt As String = ""
                        Dim tmp_madorinaiyo As String = ""

                        Call MadoriConv.SplitMadori(tmp_madori, tmp_madoricnt, tmp_madorinaiyo)

                        .Vari_Madori_cnt = tmp_madoricnt
                        .Vari_Madori_typekbn = tmp_madorinaiyo

                        '固定値
                        .Vari_Hy_guid = Guid.NewGuid.ToString
                        .Vari_Hy_gaibuno = 0    '後で一括更新する
                        .Vari_Hy_deleteflg = 0
                        .Vari_Delete_guid = Nothing
                        .Vari_Delete_day = Nothing
                        .Vari_Delete_cnt = Nothing
                        .Vari_History = DefHistory
                        .Vari_Rowid = Guid.NewGuid.ToString
                        .Vari_Delete_cause = Nothing
                        '20160519 EXEUpdateに伴う修正 部屋基本情報 -add sta
                        'とりあえずNULLにしておく (必要であればここで値を設定)
                        .Vari_Sort_hy_no = Nothing
                        .Vari_Lastupdate = Nothing
                        '20160519 EXEUpdateに伴う修正 部屋基本情報 -add end

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '部屋基本情報テーブル外部No一括更新
                Dim tmpcnt As Integer = 0
                Dim gaibunoupdatesql As String = Me.Get_UseQry_Update()
                DBExec.Exec_NonQuery(sqlcnnv10, gaibunoupdatesql, tmpcnt)

                '外部No管理テーブル一括更新
                Dim obj_hydatagaibuno As New Njc.Repository.Hydata_gaibuno_Repository.SubConv
                rtn = obj_hydatagaibuno.Set_Hydata_Gaibuno(sqlcnnv10)

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT bk_no,bk_guid FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)

            End Sub

            ''' <summary>
            ''' [hydata].[hy_gaibuno]の一括更新クエリ
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Update() As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " UPDATE hydata SET "
                tmp_sql = tmp_sql & " 	hy_gaibuno = HYGAIBUNO.GAIBUNO "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 ROW_NUMBER()OVER(PARTITION BY bk_guid ORDER BY bk_guid,hy_no) AS GAIBUNO "
                tmp_sql = tmp_sql & " 			,bk_guid "
                tmp_sql = tmp_sql & " 			,hy_guid "
                tmp_sql = tmp_sql & " 		FROM hydata "
                tmp_sql = tmp_sql & " 	) AS HYGAIBUNO "
                tmp_sql = tmp_sql & " ON  HY.bk_guid = HYGAIBUNO.bk_guid "
                tmp_sql = tmp_sql & " AND HY.hy_guid = HYGAIBUNO.hy_guid "
                tmp_sql = tmp_sql & " ; "

                Return tmp_sql

            End Function
            '20160720 連動情報構築 -del sta
            ' ''' <summary>
            ' ''' 部屋分類紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "部屋分類マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldhyruiname As String = ""
            '        Dim tmp_newhyruino As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元部屋分類名称"
            '                    tmp_oldhyruiname = fldvalue.Trim
            '                Case "賃貸革命10部屋分類No"
            '                    tmp_newhyruino = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newhyruino <> "" Then
            '            Hash_Rel_Hyrui.Add(tmp_oldhyruiname, tmp_newhyruino)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            'End Sub
            '20160720 連動情報構築 -del end
        End Class

    End Class

#End Region

#Region "部屋外部No管理情報"

    '部屋基本情報作成直後に呼び出して処理する

    Public Class Hydata_gaibuno_Repository

        Public Class SubConv

            ''' <summary>
            ''' 部屋基本情報(hydata)から外部Noを取得して外部No管理テーブル(hydata_gaibuno)へ挿入する
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Hydata_Gaibuno(ByVal sqlcnnv10 As SqlConnection) As Boolean

                Dim rtn As Boolean = True
                Dim hash_gaibuno As New Hashtable
                Dim tmp_hash As New Hashtable
                Dim model_cvitem As New Njc.Model.Hydata_gaibuno_Model
                Dim normalflg As Boolean = True
                Dim tmp_cnt As Integer = 0

                '対象テーブル初期化
                Dim tmp_sql_delete As String = " DELETE FROM hydata_gaibuno "
                rtn = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_delete, tmp_cnt)

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_gaibuno"
                Dim fldnamegrp As String = "bk_guid,current_no,history"

                'hydataから外部No最大値を集約して必要なデータを抽出
                Dim tmp_sql_select As String = Get_UseQry_Select()

                '抽出結果をハッシュテーブルへ格納
                DBExec.Exec_DataReader_Col_Hash(tmp_sql_select, sqlcnnv10, hash_gaibuno)

                'データ取得→挿入処理
                With model_cvitem

                    For Each gaibunodata In hash_gaibuno

                        .Vari_Bk_guid = gaibunodata.Key
                        .Vari_Current_no = gaibunodata.Value
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        '挿入処理
                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, normalflg)

                    Next

                End With

                Return rtn

            End Function

            ''' <summary>
            ''' 最大外部Noを持つ物件情報取得クエリ作成
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Select() As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 bk_guid "
                tmp_sql = tmp_sql & " 	,hy_gaibuno "
                tmp_sql = tmp_sql & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_guid "
                tmp_sql = tmp_sql & " 		,ROW_NUMBER()OVER(PARTITION BY bk_guid ORDER BY hy_gaibuno DESC) AS 抽出用連番 "
                tmp_sql = tmp_sql & " 		,hy_gaibuno "
                tmp_sql = tmp_sql & " 	FROM hydata "
                tmp_sql = tmp_sql & " ) AS VW "
                tmp_sql = tmp_sql & " WHERE 抽出用連番 = 1 "

                Return tmp_sql

            End Function

        End Class

    End Class

#End Region

#Region "部屋詳細情報"

    Public Class Hydata_detail_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_detail_Model           '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '2016.04.06 紐付データ取得処理を外出し -chg sta
                'Call Me.Get_RelData()
                Call SetRelItemToObject.Set_RelData_Toritaiyo()
                '2016.04.06 紐付データ取得処理を外出し -chg end

                '2016.04.06 契約分類を紐付データを元に移行 -add
                Call SetRelItemToObject.Set_RelData_Kyrui()
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_detail"
                '20160519 EXEUpdateに伴う修正 部屋詳細情報 -chg sta
                'Dim fldnamegrp As String = "hy_guid,hygyomu_setflg,syozai_kaisu1,syozai_kaisu2,syozai_kaisu3," & _
                '                           "syozai_tikaflg1,syozai_tikaflg2,syozai_tikaflg3,mukikbn,kadoheya," & _
                '                           "saikokbn,balcony_mukikbn,nyukyo_jokyokbn,nyukyo_jokyomemo,nyukyo_syunflg," & _
                '                           "nyukyo_ym,nyukyo_syunkbn,nyukyo_joken,kakunin_ymd,bosyu_startflg," & _
                '                           "freerent_flg,freerent_month,freerent_detaill,hoken_kbn,hoken_kikan," & _
                '                           "hoken_gak,hoken_biko,torihiki_taiyokbn,torihiki_kyakutuke,torihiki_tesumoto," & _
                '                           "torihiki_tesukyaku,torihiki_futankasi,torihiki_futankari,torihiki_kyakutukecomment,torihiki_gykokokukatudokbn," & _
                '                           "moto_gy_fudono,koukoku_ryoukbn,koukoku_jogengak,koukoku_jokennaiyo,biko," & _
                '                           "addr_replaceflg,addr_replacecyome,addr_replacecyomeptn,addr_replacebanti,addr_replaceetc," & _
                '                           "hosyo_gyno,hosyo_naiyo,kyrui_nokbn,keiyaku_kikan,keiyaku_kijitu," & _
                '                           "parking_biko,parking_bikebiko,parking_cyurinbiko,shared_salespoint,history," & _
                '                           "keisai_bkflg,keisai_hyflg,keisai_bantiflg,keisai_mapflg,hosyo_kbn," & _
                '                           "koukoku_jogengakkbn,koukoku_jogenrit,koukoku_jogentaxkbn,commonsalespoint_useflg,floors_flg," & _
                '                           "jisya_no,jisya_tanto,confirmky_sekininsya,confirmky_ymd,confirmky_print," & _
                '                           "confirmkai_sekininsya,confirmkai_ymd,confirmkai_print,nextnkin_kosindefault,toki_ymd," & _
                '                           "syo_kenriflg,syo_kenrikbn,other_kenriflg,btob_groupkbn,jisyaweb_osusumebk," & _
                '                           "kaiyaku_ym,taikyo_ym,svbunrui_no,krbunrui_no,siyo_mokuteki," & _
                '                           "nyukyo_sintikukbn,nyukyo_jikikbn,torihiki_jisyakbn,keiyaku_kikankbn"
                '20160829 革命10バージョンアップに伴う修正 saikokbnを削除 -del
                Dim fldnamegrp As String = "hy_guid,hygyomu_setflg,syozai_kaisu1,syozai_kaisu2,syozai_kaisu3," & _
                                           "syozai_tikaflg1,syozai_tikaflg2,syozai_tikaflg3,mukikbn,kadoheya," & _
                                           "balcony_mukikbn,nyukyo_jokyokbn,nyukyo_jokyomemo,nyukyo_syunflg," & _
                                           "nyukyo_ym,nyukyo_syunkbn,nyukyo_joken,kakunin_ymd,bosyu_startflg," & _
                                           "freerent_flg,freerent_month,freerent_detaill,hoken_kbn,hoken_kikan," & _
                                           "hoken_gak,hoken_biko,torihiki_taiyokbn,torihiki_kyakutuke,torihiki_tesumoto," & _
                                           "torihiki_tesukyaku,torihiki_futankasi,torihiki_futankari,torihiki_kyakutukecomment,torihiki_gykokokukatudokbn," & _
                                           "moto_gy_fudono,koukoku_ryoukbn,koukoku_jogengak,koukoku_jokennaiyo,biko," & _
                                           "addr_replaceflg,addr_replacecyome,addr_replacecyomeptn,addr_replacebanti,addr_replaceetc," & _
                                           "hosyo_gyno,hosyo_naiyo,kyrui_nokbn,keiyaku_kikan,keiyaku_kijitu," & _
                                           "parking_biko,parking_bikebiko,parking_cyurinbiko,shared_salespoint,history," & _
                                           "keisai_bkflg,keisai_hyflg,keisai_bantiflg,keisai_mapflg,hosyo_kbn," & _
                                           "koukoku_jogengakkbn,koukoku_jogenrit,koukoku_jogentaxkbn,commonsalespoint_useflg,floors_flg," & _
                                           "jisya_no,jisya_tanto,nextnkin_kosindefault,toki_ymd,syo_kenriflg," & _
                                           "syo_kenrikbn,other_kenriflg,btob_groupkbn,jisyaweb_osusumebk,kaiyaku_ym," & _
                                           "taikyo_ym,svbunrui_no,krbunrui_no,siyo_mokuteki,nyukyo_sintikukbn," & _
                                           "nyukyo_jikikbn,torihiki_jisyakbn,keiyaku_kikankbn"
                '20160519 EXEUpdateに伴う修正 部屋詳細情報 -chg end

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub As String = readtbl.Columns(keycol_sub - 1).ColumnName.Trim

                    For cntii As Integer = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""

                        'キー値取得
                        Dim fldvalue_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim fldvalue_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub - 1)).Trim
                        Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub
                        .Vari_Hy_guid = fldvalue_key

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "部屋業務期間有無"
                                    .Vari_Hygyomu_setflg = fldvalue.Trim
                                Case "所在階１"
                                    .Vari_Syozai_kaisu1 = fldvalue.Trim
                                Case "所在階２"
                                    .Vari_Syozai_kaisu2 = fldvalue.Trim
                                Case "所在階３"
                                    .Vari_Syozai_kaisu3 = fldvalue.Trim
                                Case "地下フラグ１"
                                    .Vari_Syozai_tikaflg1 = fldvalue.Trim
                                Case "地下フラグ２"
                                    .Vari_Syozai_tikaflg2 = fldvalue.Trim
                                Case "地下フラグ３"
                                    .Vari_Syozai_tikaflg3 = fldvalue.Trim
                                Case "向き"
                                    .Vari_Mukikbn = fldvalue.Trim
                                Case "角部屋フラグ"
                                    .Vari_Kadoheya = fldvalue.Trim
                                    '20160829 革命10バージョンアップに伴う修正 -del sta
                                    'Case "採光"
                                    '    .Vari_Saikokbn = fldvalue.Trim
                                    '20160829 革命10バージョンアップに伴う修正 -del end
                                Case "バルコニー方向"
                                    .Vari_Balcony_mukikbn = fldvalue.Trim
                                Case "状況"
                                    .Vari_Nyukyo_jokyokbn = fldvalue.Trim
                                Case "入居状況備考"
                                    .Vari_Nyukyo_jokyomemo = fldvalue.Trim
                                Case "入居可能日を時期で指定"
                                    .Vari_Nyukyo_syunflg = fldvalue.Trim
                                Case "入居可能時期(年月)"
                                    .Vari_Nyukyo_ym = fldvalue.Trim
                                Case "入居可能時期(上旬・中旬・下旬)"
                                    .Vari_Nyukyo_syunkbn = fldvalue.Trim
                                Case "入居条件"
                                    .Vari_Nyukyo_joken = fldvalue.Trim
                                Case "広告内容確認日"
                                    .Vari_Kakunin_ymd = fldvalue.Trim
                                Case "募集可能種別(しない・する)"
                                    .Vari_Bosyu_startflg = fldvalue.Trim
                                Case "フリーレント有無"
                                    .Vari_Freerent_flg = fldvalue.Trim
                                Case "フリーレントカ月"
                                    .Vari_Freerent_month = fldvalue.Trim
                                Case "フリーレント詳細"
                                    .Vari_Freerent_detaill = fldvalue.Trim
                                Case "保険-保険の利用"
                                    .Vari_Hoken_kbn = fldvalue.Trim
                                Case "保険-保険期間"
                                    .Vari_Hoken_kikan = fldvalue.Trim
                                Case "保険-保険料"
                                    .Vari_Hoken_gak = fldvalue.Trim
                                Case "保険-保険の備考"
                                    .Vari_Hoken_biko = fldvalue.Trim
                                Case "取引形態種別"
                                    .Vari_Torihiki_taiyokbn = fldvalue.Trim
                                Case "客付け状態可否"
                                    .Vari_Torihiki_kyakutuke = fldvalue.Trim
                                Case "手数料負担割合貸主(%)"
                                    .Vari_Torihiki_futankasi = fldvalue.Trim
                                Case "手数料負担割合借主(%)"
                                    .Vari_Torihiki_futankari = fldvalue.Trim
                                Case "手数料配分元付(%)"
                                    .Vari_Torihiki_tesumoto = fldvalue.Trim
                                Case "手数料配分先物(%)"
                                    .Vari_Torihiki_tesukyaku = fldvalue.Trim
                                Case "客付会社への物件コメント"
                                    .Vari_Torihiki_kyakutukecomment = fldvalue.Trim
                                Case "業者間広告広告活動種別"
                                    .Vari_Torihiki_gykokokukatudokbn = fldvalue.Trim
                                Case "情報元業者No"
                                    .Vari_Moto_gy_fudono = fldvalue.Trim
                                Case "広告料有無"
                                    .Vari_Koukoku_ryoukbn = fldvalue.Trim
                                Case "広告料上限額"
                                    .Vari_Koukoku_jogengak = fldvalue.Trim
                                Case "広告料条件内容"
                                    .Vari_Koukoku_jokennaiyo = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko = fldvalue.Trim
                                Case "部屋住所の変更フラグ"
                                    .Vari_Addr_replaceflg = fldvalue.Trim
                                Case "部屋住所(丁目)"
                                    .Vari_Addr_replacecyome = fldvalue.Trim
                                Case "部屋住所(丁目区分)"
                                    .Vari_Addr_replacecyomeptn = fldvalue.Trim
                                Case "部屋住所(町地域)"
                                    .Vari_Addr_replacebanti = fldvalue.Trim
                                Case "部屋住所(その他)"
                                    .Vari_Addr_replaceetc = fldvalue.Trim
                                Case "保証会社"
                                    .Vari_Hosyo_gyno = fldvalue.Trim
                                Case "保証内容"
                                    .Vari_Hosyo_naiyo = fldvalue.Trim
                                Case "契約の種類"
                                    .Vari_Kyrui_nokbn = fldvalue.Trim
                                Case "定期借家契約(期間)"
                                    .Vari_Keiyaku_kikan = fldvalue.Trim
                                Case "定期借家契約(期日)"
                                    .Vari_Keiyaku_kijitu = fldvalue.Trim
                                Case "駐車場備考"
                                    .Vari_Parking_biko = fldvalue.Trim
                                Case "バイク駐車場備考"
                                    .Vari_Parking_bikebiko = fldvalue.Trim
                                Case "駐輪場備考"
                                    .Vari_Parking_cyurinbiko = fldvalue.Trim
                                Case "共通セールスポイント"
                                    .Vari_Shared_salespoint = fldvalue.Trim
                                Case "不動産検索サイト掲載設定-物件名"
                                    .Vari_Keisai_bkflg = fldvalue.Trim
                                Case "不動産検索サイト掲載設定-部屋NO"
                                    .Vari_Keisai_hyflg = fldvalue.Trim
                                Case "不動産検索サイト掲載設定-丁番地以下"
                                    .Vari_Keisai_bantiflg = fldvalue.Trim
                                Case "不動産検索サイト掲載設定-地図上"
                                    .Vari_Keisai_mapflg = fldvalue.Trim
                                Case "賃貸保証利用区分"
                                    .Vari_Hosyo_kbn = fldvalue.Trim
                                Case "広告料上限額区分"
                                    .Vari_Koukoku_jogengakkbn = fldvalue.Trim
                                Case "広告料上限率"
                                    .Vari_Koukoku_jogenrit = fldvalue.Trim
                                Case "広告料上限額税区分"
                                    .Vari_Koukoku_jogentaxkbn = fldvalue.Trim
                                Case "共通セールスポイント使用フラグ"
                                    .Vari_Commonsalespoint_useflg = fldvalue.Trim
                                Case "複数階有りのフラグ"
                                    .Vari_Floors_flg = fldvalue.Trim
                                Case "支店NO"
                                    .Vari_Jisya_no = fldvalue.Trim
                                Case "自社担当者No"
                                    .Vari_Jisya_tanto = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 部屋詳細情報 -del sta
                                    'Case "確認事項(契約)責任者"
                                    '    .Vari_Confirmky_sekininsya = fldvalue.Trim
                                    'Case "確認事項(契約)内容確認日"
                                    '    .Vari_Confirmky_ymd = fldvalue.Trim
                                    'Case "確認事項(契約)印刷時"
                                    '    .Vari_Confirmky_print = fldvalue.Trim
                                    'Case "確認事項(解約)責任者"
                                    '    .Vari_Confirmkai_sekininsya = fldvalue.Trim
                                    'Case "確認事項(解約)内容確認日"
                                    '    .Vari_Confirmkai_ymd = fldvalue.Trim
                                    'Case "確認事項(解約)印刷時"
                                    '    .Vari_Confirmkai_print = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 部屋詳細情報 -del end
                                Case "次回更新時の初期値"
                                    .Vari_Nextnkin_kosindefault = fldvalue.Trim
                                Case "登記情報の日付"
                                    .Vari_Toki_ymd = fldvalue.Trim
                                Case "所有権にかかる権利有無"
                                    .Vari_Syo_kenriflg = fldvalue.Trim
                                Case "所有権にかかる権利の種類"
                                    .Vari_Syo_kenrikbn = fldvalue.Trim
                                Case "所有権以外の権利有無"
                                    .Vari_Other_kenriflg = fldvalue.Trim
                                Case "BtoBプラグイングループ設定区分"
                                    .Vari_Btob_groupkbn = fldvalue.Trim
                                Case "自社Webオススメ物件表示"
                                    .Vari_Jisyaweb_osusumebk = fldvalue.Trim
                                Case "解約日"
                                    .Vari_Kaiyaku_ym = fldvalue.Trim
                                Case "退去日"
                                    .Vari_Taikyo_ym = fldvalue.Trim
                                Case "使用目的"
                                    .Vari_Siyo_mokuteki = fldvalue.Trim
                                Case "新築区分"
                                    .Vari_Nyukyo_sintikukbn = fldvalue.Trim
                                Case "入居時期区分"
                                    .Vari_Nyukyo_jikikbn = fldvalue.Trim
                                Case "取引自社区分"
                                    .Vari_Torihiki_jisyakbn = fldvalue.Trim
                                Case "契約期間区分"
                                    .Vari_Keiyaku_kikankbn = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_Svbunrui_no = 0   '京王カスタマイズ分
                        .Vari_Krbunrui_no = 0   '使用箇所不明なためデフォルト値を設定
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20161012 部屋所在階チェック処理の追加 -add
                        Call Me.Chk_Kaisu(sqlcnnv10, tblname, hash_cvitem, hash_log)

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end
            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 部屋所在階のチェック '20161012 部屋所在階チェック処理の追加
            ''' 例)物件階数=5階の場合で部屋階数=6階はあり得ない
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_log"></param>
            ''' <remarks></remarks>
            Public Sub Chk_Kaisu(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_cvitem As Hashtable, ByRef hash_log As Hashtable)

                Dim tmp_key() As String = hash_cvitem("hy_guid").ToString.Split("-")
                Dim bkno As String = tmp_key(0)
                'Dim kaisu1 As Integer = (IIf(hash_cvitem("syozai_kaisu1") = "", 0, Int32.Parse(hash_cvitem("syozai_kaisu1"))))
                'Dim kaisu2 As Integer = (IIf(hash_cvitem("syozai_kaisu2") = "", 0, Int32.Parse(hash_cvitem("syozai_kaisu2"))))
                'Dim kaisu3 As Integer = (IIf(hash_cvitem("syozai_kaisu3") = "", 0, Int32.Parse(hash_cvitem("syozai_kaisu3"))))
                Dim tmp_kaisu1 As String = IIf(hash_cvitem("syozai_kaisu1") Is Nothing, "", hash_cvitem("syozai_kaisu1"))
                Dim tmp_kaisu2 As String = IIf(hash_cvitem("syozai_kaisu2") Is Nothing, "", hash_cvitem("syozai_kaisu2"))
                Dim tmp_kaisu3 As String = IIf(hash_cvitem("syozai_kaisu3") Is Nothing, "", hash_cvitem("syozai_kaisu3"))
                Dim tikaflg1 As Integer = (IIf(Int32.Parse(hash_cvitem("syozai_tikaflg1")) = 1, -1, 1))
                Dim tikaflg2 As Integer = (IIf(Int32.Parse(hash_cvitem("syozai_tikaflg2")) = 1, -1, 1))
                Dim tikaflg3 As Integer = (IIf(Int32.Parse(hash_cvitem("syozai_tikaflg3")) = 1, -1, 1))
                Dim kaisu1 As Integer = 0
                Dim kaisu2 As Integer = 0
                Dim kaisu3 As Integer = 0
                Dim bkkai As Integer = 0
                Dim tmp_bkkai As String = ""
                Dim tmp_hubi As String = "物件情報の階建てを超えているため移行できません。"     'ログ出力用

                '階数調整
                If tmp_kaisu1 <> "" Then
                    kaisu1 = Int32.Parse(tmp_kaisu1) * tikaflg1
                End If
                If tmp_kaisu2 <> "" Then
                    kaisu2 = Int32.Parse(tmp_kaisu2) * tikaflg2
                End If
                If tmp_kaisu3 <> "" Then
                    kaisu3 = Int32.Parse(tmp_kaisu3) * tikaflg3
                End If

                '物件No有無/無効データチェック(念の為)
                If bkno = "" OrElse Typ.toInt(bkno) = 0 Then
                    Exit Sub
                End If

                '物件情報の階建てを取得
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT ISNULL(kaidate,0) FROM bkdata_detail AS BKD "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON BKD.bk_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " WHERE bk_no = " & bkno
                If DBExec.Exec_Scalar(tmp_sql, sqlcnnv10) IsNot Nothing Then
                    bkkai = Int32.Parse(DBExec.Exec_Scalar(tmp_sql, sqlcnnv10))
                End If

                '階建てが設定されていない場合は処理を抜ける(未設定の場合は部屋所在階を自由に設定できるため)
                If bkkai = 0 Then
                    Exit Sub
                End If

                '所在階1チェック
                If kaisu1 <> 0 Then
                    If bkkai < kaisu1 Then
                        'ログ
                        Dim log_key As String = tblname & "-" & "syozai_kaisu1"
                        Dim errstr As String = LOG_NAIYO_ERR_OUTOFRANGE & "-" & tmp_hubi & "-" & LOG_TAISYO_DEFAULT
                        hash_log.Add(log_key, errstr)
                        '値の再格納
                        hash_cvitem("syozai_kaisu1") = ""
                        hash_cvitem("syozai_tikaflg1") = "0"
                    End If
                End If

                '所在階2チェック
                If kaisu2 <> 0 Then
                    If bkkai < kaisu2 Then
                        'ログ
                        Dim log_key As String = tblname & "-" & "syozai_kaisu2"
                        Dim errstr As String = LOG_NAIYO_ERR_OUTOFRANGE & "-" & tmp_hubi & "-" & LOG_TAISYO_DEFAULT
                        hash_log.Add(log_key, errstr)
                        '値の再格納
                        hash_cvitem("syozai_kaisu2") = ""
                        hash_cvitem("syozai_tikaflg2") = "0"
                    End If
                End If

                '所在階3チェック
                If kaisu3 <> 0 Then
                    If bkkai < kaisu3 Then
                        'ログ
                        Dim log_key As String = tblname & "-" & "syozai_kaisu3"
                        Dim errstr As String = LOG_NAIYO_ERR_OUTOFRANGE & "-" & tmp_hubi & "-" & LOG_TAISYO_DEFAULT
                        hash_log.Add(log_key, errstr)
                        '値の再格納
                        hash_cvitem("syozai_kaisu3") = ""
                        hash_cvitem("syozai_tikaflg3") = "0"
                    End If
                End If

            End Sub

            '2016.04.06 紐付データ取得処理を外出し -del sta
            ' ''' <summary>
            ' ''' 取引態様マスタ紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "取引態様マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldtoritaiyoname As String = ""
            '        Dim tmp_newtoritaiyono As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元取引態様名称"
            '                    tmp_oldtoritaiyoname = fldvalue.Trim
            '                Case "賃貸革命10取引態様No"
            '                    tmp_newtoritaiyono = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newtoritaiyono <> "" Then
            '            Hash_Rel_Toritaiyo.Add(tmp_oldtoritaiyoname, tmp_newtoritaiyono)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub
            '2016.04.06 紐付データ取得処理を外出し -del end

        End Class

    End Class

#End Region

#Region "部屋所有者情報"

    Public Class Hydata_syo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_syo_Model              '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_syo"
                Dim fldnamegrp As String = "hy_guid,kn_no,sorule_guid,kasi1_ow_no,syo1_ow_no," & _
                                           "kasi2_ow_no,syo2_ow_no,syo_startymd,syo_endymd,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub1 As String = readtbl.Columns(keycol_sub1 - 1).ColumnName.Trim
                    Dim fldname_keysub2 As String = readtbl.Columns(keycol_sub2 - 1).ColumnName.Trim

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "管理No"
                                    .Vari_Kn_no = fldvalue.Trim
                                Case "貸主１No"
                                    .Vari_Kasi1_ow_no = fldvalue.Trim
                                Case "所有者１No"
                                    .Vari_Syo1_ow_no = fldvalue.Trim
                                Case "貸主２No"
                                    .Vari_Kasi2_ow_no = fldvalue.Trim
                                Case "所有者２No"
                                    .Vari_Syo2_ow_no = fldvalue.Trim
                                Case "所有期間開始"
                                    .Vari_Syo_startymd = fldvalue.Trim
                                Case "所有期間終了"
                                    .Vari_Syo_endymd = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_Sorule_guid = Guid.NewGuid.ToString
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,kn_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋駐車場情報"

    Public Class Hydata_parking_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_parking_Model          '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_parking"
                Dim fldnamegrp As String = "hy_guid,parking_kbn,parking_akisu,parking_status,parking_gakkbn," & _
                                           "parking_gak,parking_gakzei,history,parking_tintaisu"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub1 As String = readtbl.Columns(keycol_sub1 - 1).ColumnName.Trim
                    Dim fldname_keysub2 As String = readtbl.Columns(keycol_sub2 - 1).ColumnName.Trim

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "駐車場区分"
                                    .Vari_Parking_kbn = fldvalue.Trim
                                Case "駐車場の空き数(手動設定用)"
                                    .Vari_Parking_akisu = fldvalue.Trim
                                Case "駐車場の空き有無(手動設定用)"
                                    .Vari_Parking_status = fldvalue.Trim
                                Case "駐車場料金区分(手動設定用)"
                                    .Vari_Parking_gakkbn = fldvalue.Trim
                                Case "駐車場料金(手動設定用)"
                                    .Vari_Parking_gak = fldvalue.Trim
                                Case "駐車場料金税区分(手動設定用)"
                                    .Vari_Parking_gakzei = fldvalue.Trim
                                Case "駐車場の賃貸可能数"
                                    .Vari_Parking_tintaisu = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parking_kbn) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋特約情報"

    Public Class Hydata_tokuyaku_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_tokuyaku_Model         '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"
                Dim totalrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

                '************************
                '作業準備
                '************************
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_tokuyaku"
                Dim fldnamegrp As String = "hy_guid,tokuyaku_grpno,naiyo,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '移行対象件数取得用
                Dim cvtaisyocnt As Integer = 0

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub1 As String = readtbl.Columns(keycol_sub1 - 1).ColumnName.Trim
                    Dim fldname_keysub2 As String = readtbl.Columns(keycol_sub2 - 1).ColumnName.Trim

                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '移行判別用
                        Dim cvflg As Boolean = True
                        Dim keychkflg As Boolean = True

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        Dim tmp_keyoya As String = tmp_keymain & "-" & tmp_keysub1
                        Dim tmp_keydup As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1
                        .Vari_Tokuyaku_grpno = tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2

                        '特約カウント初期化
                        Dim cvitemcnt As Integer = 0
                        Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

                        'データ有無チェック
                        Dim tmp_fldvalueumuchk As String = ""
                        For cntjj = 3 To readtbl.Columns.Count - 1
                            tmp_fldvalueumuchk = tmp_fldvalueumuchk & Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                        Next
                        If tmp_fldvalueumuchk = "" Then
                            cvflg = False
                        End If

                        If cvflg Then

                            '移行対象件数カウント
                            cvtaisyocnt = cvtaisyocnt + 1

                            Dim hash_keylog As New Hashtable
                            Dim log_keyout As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1
                            '親データ有無チェック
                            If list_basekeydata.Contains(tmp_keyoya) = False Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "hydata_tokuyaku-hy_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(tmp_keydup) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "hydata_tokuyaku-hy_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            If keychkflg Then

                                '作業用変数作成
                                Dim tmp_tokuyaku As String = ""

                                '移行値取得
                                For cntjj = 3 To readtbl.Columns.Count - 1

                                    'サブキー取得
                                    Dim tmp_keysub3 As String = (cntjj - 2).ToString

                                    'ログ出力用データ格納(サブキーフィールド)
                                    Dim fldname_keysub3 As String = readtbl.Columns(cntjj).ColumnName.Trim

                                    '登録値取得
                                    Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                                    If fldvalue <> "" Then
                                        tmp_tokuyaku = tmp_tokuyaku & LINE_BREAK & fldvalue
                                    End If

                                Next

                                '成形
                                '2016.04.26 メインの方へも反映させる修正 -chg sta
                                '.Vari_Naiyo = tmp_tokuyaku.Remove(0, LINE_BREAK.Length)
                                If tmp_tokuyaku <> "" Then
                                    .Vari_Naiyo = tmp_tokuyaku.Remove(0, LINE_BREAK.Length)
                                End If
                                '2016.04.26 メインの方へも反映させる修正 -chg end

                                '固定値
                                .Vari_History = DefHistory

                                '特約にデータが存在する場合に書込処理を行う
                                If .Vari_Naiyo <> "" Then

                                    '20160928 メモ関連の移行件数表示修正 -add
                                    tmp_cvrowcnt = tmp_cvrowcnt + 1

                                    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                                    'データチェック
                                    Dim skipflg As Boolean = False
                                    Dim hash_cvitem As New Hashtable
                                    Dim hash_log As New Hashtable
                                    skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                                    '書込処理
                                    If Not skipflg Then

                                        'キーをguidへ変換
                                        '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                                        ''2016.04.26 メインの方へも反映させる修正 -chg sta
                                        ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                                        'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                                        ''2016.04.26 メインの方へも反映させる修正 -chg end
                                        hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                                        '20160819 部屋の半角変換処理によるエラー修正 -chg end

                                        '挿入処理
                                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                        '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                        If normalflg Then
                                            If list_chkduplicate.Contains(tmp_keydup) = False Then
                                                list_chkduplicate.Add(tmp_keydup)
                                            End If
                                            cvitemcnt = cvitemcnt + 1
                                        End If

                                    End If

                                    'ログ出力メッセージ整形
                                    Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                                    'ログ出力
                                    If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                                        Dim tmp_cnt As Integer = 0
                                        '挿入
                                        For Each logvalue In sortlist_log
                                            Dim tmp_sql_insert As String = logvalue.Value
                                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                        Next
                                        '初期化
                                        tmp_logcnt = 0
                                        sortlist_log.Clear()
                                    End If

                                End If

                            Else

                                '親データ無し、重複チェックエラーログ出力
                                'ログ出力メッセージ整形
                                tmp_hash.Clear()
                                Call LogSetting.Set_Log_Value_KomkErr(hash_keylog, tmp_hash, tmp_hash, log_keyout, tblname, tmp_logcnt, sortlist_log)  '※引数のtmp_hashは使用しないが仮で入れておく

                                'ログ出力
                                If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = totalrowcnt - 1) Then
                                    Dim tmp_cnt As Integer = 0
                                    '挿入
                                    For Each logvalue In sortlist_log
                                        Dim tmp_sql_insert As String = logvalue.Value
                                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                    Next
                                    '初期化
                                    tmp_logcnt = 0
                                    sortlist_log.Clear()
                                End If

                            End If


                        End If

                        '--------------------------------------------------------------------
                        '移行した備考が1データ以上ある場合移行したレコードの数を更新する
                        '--------------------------------------------------------------------
                        '20160928 メモ関連の移行件数表示修正 -chg sta
                        'tmp_cvcnt = tmp_cvcnt + 1
                        If tmp_cvrowcnt > 0 Then
                            totalrowcnt = totalrowcnt + 1
                        End If
                        If cvitemcnt > 0 Then
                            tmp_cvcnt = tmp_cvcnt + 1
                        End If
                        '20160928 メモ関連の移行件数表示修正 -chg end
                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = cvtaisyocnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,tokuyaku_grpno) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋鍵情報"

    Public Class Hydata_kagi_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_kagi_Model             '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                Dim cnt_cvcnt As Integer = 0                                    '20160531 鍵情報移行処理の修正 専用鍵の件数取得用 -add

                '************************
                '作業準備
                '************************

                '20160613 鍵情報の取得処理修正 -chg sta
                ''10鍵タイトルマスタから鍵Noとタイトル名を紐付けたデータを取得
                'Call EtcMethod.Set_KagiTitleMst(sqlcnnv10, 1)
                'Call EtcMethod.Set_KagiTitleMst(sqlcnnv10, 2)

                'V7鍵Noと10鍵No(鍵区分を含む)を紐付けたハッシュテーブルを作成
                Call SetRelItemToObject.Set_RelData_KagiInfo(sqlcnnv10)
                '20160613 鍵情報の取得処理修正 -chg end
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_kagi"
                Dim fldnamegrp As String = "hy_guid,kagi_no,honsu,biko,history," & _
                                           "hokan,gyshare"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub1 As String = readtbl.Columns(keycol_sub1 - 1).ColumnName.Trim
                    Dim fldname_keysub2 As String = readtbl.Columns(keycol_sub2 - 1).ColumnName.Trim

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "鍵No"      '20160613 鍵情報の取得処理修正 -chg sta 鍵タイトル名→鍵No
                                    .Vari_Kagi_no = fldvalue.Trim
                                Case "鍵本数"
                                    .Vari_Honsu = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko = fldvalue.Trim
                                Case "保管場所"
                                    .Vari_Hokan = fldvalue.Trim
                                Case "業者間での情報共有"
                                    .Vari_Gyshare = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory
                        '20160927 汎用CV時の鍵情報の取得処理修正 -chg sta
                        ''20160613 鍵情報の取得処理修正 -chg sta
                        ' ''20160531 鍵情報移行処理の修正 -chg sta
                        ' ''共用鍵の場合に移行処理を行う
                        '' ''フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        ' ''tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        '' ''データチェック
                        ' ''Dim skipflg As Boolean = False
                        ' ''Dim hash_cvitem As New Hashtable
                        ' ''Dim hash_log As New Hashtable
                        ' ''skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '' ''書込処理
                        ' ''If Not skipflg Then

                        ' ''    'キーをguidへ変換
                        ' ''    '2016.04.26 メインの方へも反映させる修正 -chg sta
                        ' ''    'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        ' ''    hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        ' ''    '2016.04.26 メインの方へも反映させる修正 -chg end

                        ' ''    '挿入処理
                        ' ''    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                        ' ''    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        ' ''    If normalflg Then
                        ' ''        list_chkduplicate.Add(fldvalue_key)
                        ' ''        tmp_cvcnt = tmp_cvcnt + 1
                        ' ''    End If

                        ' ''End If

                        '' ''-------------------
                        '' ''ログ出力
                        '' ''-------------------

                        '' ''ログ出力メッセージ整形
                        ' ''Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        '' ''ログ出力
                        ' ''If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                        ' ''    Dim tmp_cnt As Integer = 0
                        ' ''    '挿入
                        ' ''    For Each logvalue In sortlist_log
                        ' ''        Dim tmp_sql_insert As String = logvalue.Value
                        ' ''        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                        ' ''    Next
                        ' ''    '初期化
                        ' ''    tmp_logcnt = 0
                        ' ''    sortlist_log.Clear()
                        ' ''End If

                        ''If Hash_KagiTitleSenyo.Contains(.Vari_Kagi_no) Then

                        ''    '件数カウント
                        ''    cnt_cvcnt = cnt_cvcnt + 1

                        ''    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        ''    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        ''    'データチェック
                        ''    Dim skipflg As Boolean = False
                        ''    Dim hash_cvitem As New Hashtable
                        ''    Dim hash_log As New Hashtable
                        ''    skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        ''    '書込処理
                        ''    If Not skipflg Then

                        ''        'キーをguidへ変換
                        ''        '2016.04.26 メインの方へも反映させる修正 -chg sta
                        ''        'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        ''        hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        ''        '2016.04.26 メインの方へも反映させる修正 -chg end

                        ''        '挿入処理
                        ''        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                        ''        '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        ''        If normalflg Then
                        ''            list_chkduplicate.Add(fldvalue_key)
                        ''            tmp_cvcnt = tmp_cvcnt + 1
                        ''        End If

                        ''    End If

                        ''    '-------------------
                        ''    'ログ出力
                        ''    '-------------------

                        ''    'ログ出力メッセージ整形
                        ''    Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        ''    'ログ出力
                        ''    If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                        ''        Dim tmp_cnt As Integer = 0
                        ''        '挿入
                        ''        For Each logvalue In sortlist_log
                        ''            Dim tmp_sql_insert As String = logvalue.Value
                        ''            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                        ''        Next
                        ''        '初期化
                        ''        tmp_logcnt = 0
                        ''        sortlist_log.Clear()
                        ''    End If

                        ''End If
                        ' ''20160531 鍵情報移行処理の修正 -chg end

                        ''フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        'tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        ''データチェック
                        'Dim skipflg As Boolean = False
                        'Dim hash_cvitem As New Hashtable
                        'Dim hash_log As New Hashtable
                        'skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        'If hash_cvitem.Item("kagi_no") <> "" Then

                        '    '件数カウント
                        '    cnt_cvcnt = cnt_cvcnt + 1

                        '    '書込処理
                        '    If Not skipflg Then

                        '        'キーをguidへ変換
                        '        '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        '        ''2016.04.26 メインの方へも反映させる修正 -chg sta
                        '        ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        '        'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        '        ''2016.04.26 メインの方へも反映させる修正 -chg end
                        '        hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        '        '20160819 部屋の半角変換処理によるエラー修正 -chg end

                        '        '挿入処理
                        '        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                        '        '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        '        If normalflg Then
                        '            list_chkduplicate.Add(fldvalue_key)
                        '            tmp_cvcnt = tmp_cvcnt + 1
                        '        End If

                        '    End If

                        '    '-------------------
                        '    'ログ出力
                        '    '-------------------

                        '    'ログ出力メッセージ整形
                        '    Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        '    'ログ出力
                        '    If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                        '        Dim tmp_cnt As Integer = 0
                        '        '挿入
                        '        For Each logvalue In sortlist_log
                        '            Dim tmp_sql_insert As String = logvalue.Value
                        '            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                        '        Next
                        '        '初期化
                        '        tmp_logcnt = 0
                        '        sortlist_log.Clear()
                        '    End If

                        'End If
                        ''20160613 鍵情報の取得処理修正 -chg end

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        Select Case CNVNO
                            Case ConvertTypes._既存ユーザ用

                                If hash_cvitem.Item("kagi_no") <> "" Then

                                    '件数カウント
                                    cnt_cvcnt = cnt_cvcnt + 1

                                    '書込処理
                                    If Not skipflg Then

                                        'キーをguidへ変換
                                        hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))

                                        '挿入処理
                                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                        '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                        If normalflg Then
                                            list_chkduplicate.Add(fldvalue_key)
                                            tmp_cvcnt = tmp_cvcnt + 1
                                        End If

                                    End If

                                    '-------------------
                                    'ログ出力
                                    '-------------------

                                    'ログ出力メッセージ整形
                                    Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                                    'ログ出力
                                    If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                                        Dim tmp_cnt As Integer = 0
                                        '挿入
                                        For Each logvalue In sortlist_log
                                            Dim tmp_sql_insert As String = logvalue.Value
                                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                        Next
                                        '初期化
                                        tmp_logcnt = 0
                                        sortlist_log.Clear()
                                    End If

                                End If

                            Case ConvertTypes._汎用

                                '書込処理
                                If Not skipflg Then

                                    'キーをguidへ変換
                                    hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))

                                    '挿入処理
                                    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                    If normalflg Then
                                        list_chkduplicate.Add(fldvalue_key)
                                        tmp_cvcnt = tmp_cvcnt + 1
                                    End If

                                End If

                                '-------------------
                                'ログ出力
                                '-------------------

                                'ログ出力メッセージ整形
                                Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                                'ログ出力
                                If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                                    Dim tmp_cnt As Integer = 0
                                    '挿入
                                    For Each logvalue In sortlist_log
                                        Dim tmp_sql_insert As String = logvalue.Value
                                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                    Next
                                    '初期化
                                    tmp_logcnt = 0
                                    sortlist_log.Clear()
                                End If

                        End Select
                        '20160927 汎用CV時の鍵情報の取得処理修正 -chg end
                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************
                '20160927 汎用CV時の鍵情報の取得処理修正 -chg sta
                ''20160531 鍵情報移行処理の修正 -chg sta
                ''共用鍵の場合に中間ファイルの件数をカウントする
                ''中間ファイル件数を取得
                ''midrowcnt = rowcnt
                'midrowcnt = cnt_cvcnt
                ''20160531 鍵情報移行処理の修正 -chg end
                Select Case CNVNO
                    Case ConvertTypes._既存ユーザ用
                        midrowcnt = cnt_cvcnt
                    Case ConvertTypes._汎用
                        midrowcnt = rowcnt
                End Select
                '20160927 汎用CV時の鍵情報の取得処理修正 -chg end
                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,kagi_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋面積情報"

    Public Class Hydata_othermenseki_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
                Dim startrow As Integer                                         '書込開始行
                Dim columncnt As Integer                                        '列数
                Dim maxrowcnt As Integer                                        '既存データの行数
                Dim rowcnt As Integer                                           '書込行数
                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納

                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_othermenseki_Model     '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_othermenseki"
                Dim fldnamegrp As String = "hy_guid,mensekitype,no,mensekikbn,name," & _
                                           "meter,tubo,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 1 To rowcnt

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                            Return rtn
                        End If

                        '行取得
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                        'キー値取得
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "面積タイプ"
                                    .Vari_Mensekitype = fldvalue.Trim
                                Case "面積No"
                                    .Vari_No = fldvalue.Trim
                                Case "面積区分"
                                    .Vari_Mensekikbn = fldvalue.Trim
                                Case "区画名"
                                    .Vari_Name = fldvalue.Trim
                                Case "面積(m2)"
                                    .Vari_Meter = fldvalue.Trim
                                Case "坪数"
                                    .Vari_Tubo = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii
                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,mensekitype)  + '-' + CONVERT(varchar,no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋間取内訳情報"

    Public Class Hydata_madoriutiwake_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_madoriutiwake_Model    '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_madoriutiwake"
                Dim fldnamegrp As String = "hy_guid,no,hykbn,jo,history," & _
                                           "syozaikai"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub As String = readtbl.Columns(keycol_sub - 1).ColumnName.Trim

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""

                        'キー値取得
                        Dim fldvalue_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim fldvalue_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub - 1)).Trim
                        Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub
                        .Vari_Hy_guid = fldvalue_key

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

                        '間取内訳カウント初期化
                        Dim cvitemcnt As Integer = 0

                        '作業用変数作成
                        Dim tmp_madori As String = ""
                        Dim tmp_madorino() As String = Nothing
                        Dim tmp_madorikbn() As String = Nothing
                        Dim tmp_madorijo() As String = Nothing

                        '格納
                        '移行値取得
                        '20160915 間取内訳情報の汎用/既存処理の分岐 -chg sta
                        'For cntjj = 1 To columncnt

                        '    Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                        '    Dim fldvalue As String = ""
                        '    If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                        '        fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                        '    End If

                        '    If fldname = "既存間取内訳データ" Then

                        '        '間取の連結文字列を取得
                        '        tmp_madori = fldvalue.Trim

                        '        '間取変換
                        '        Call MadoriConv.SplitMadori(tmp_madori, tmp_madorino, tmp_madorikbn, tmp_madorijo)

                        '        '変数格納→挿入処理
                        '        For cntkk = 0 To UBound(tmp_madorino)

                        '            .Vari_No = tmp_madorino(cntkk)
                        '            .Vari_Hykbn = tmp_madorikbn(cntkk)
                        '            .Vari_Jo = tmp_madorijo(cntkk)

                        '            '固定値
                        '            .Vari_History = DefHistory
                        '            .Vari_Syozaikai = ""

                        '            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        '            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        '            'データチェック
                        '            Dim skipflg As Boolean = False
                        '            Dim hash_cvitem As New Hashtable
                        '            Dim hash_log As New Hashtable
                        '            skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '            '書込処理
                        '            If Not skipflg Then

                        '                'キーをguidへ変換
                        '                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        '                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                        '                ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        '                'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        '                ''2016.04.26 メインの方へも反映させる修正 -chg end
                        '                hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        '                '20160819 部屋の半角変換処理によるエラー修正 -chg end

                        '                '挿入処理
                        '                Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                        '                '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        '                If normalflg Then
                        '                    list_chkduplicate.Add(fldvalue_key)
                        '                    cvitemcnt = cvitemcnt + 1
                        '                End If

                        '            End If

                        '            '-------------------
                        '            'ログ出力
                        '            '-------------------

                        '            'ログ出力メッセージ整形
                        '            Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        '            'ログ出力
                        '            If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                        '                Dim tmp_cnt As Integer = 0
                        '                '挿入
                        '                For Each logvalue In sortlist_log
                        '                    Dim tmp_sql_insert As String = logvalue.Value
                        '                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                        '                Next
                        '                '初期化
                        '                tmp_logcnt = 0
                        '                sortlist_log.Clear()
                        '            End If

                        '        Next

                        '    End If

                        'Next

                        ''--------------------------------------------------------------------
                        ''移行した間取内訳が1データ以上ある場合移行したレコードの数を更新する
                        ''--------------------------------------------------------------------
                        'If cvitemcnt > 0 Then
                        '    tmp_cvcnt = tmp_cvcnt + 1
                        'End If

                        Select Case CNVNO
                            Case ConvertTypes._既存ユーザ用

                                'データ取得
                                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                                    '項目名取得
                                    fldname = readtbl.Columns(cntjj).ColumnName.Trim

                                    '登録値取得
                                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                                    If fldname = "既存間取内訳データ" Then

                                        '間取の連結文字列を取得
                                        tmp_madori = fldvalue.Trim

                                        '間取変換
                                        Call MadoriConv.SplitMadori(tmp_madori, tmp_madorino, tmp_madorikbn, tmp_madorijo)

                                        '変数格納→挿入処理
                                        If tmp_madori <> "" Then

                                            For cntkk = 0 To UBound(tmp_madorino)

                                                .Vari_No = tmp_madorino(cntkk)
                                                .Vari_Hykbn = tmp_madorikbn(cntkk)
                                                .Vari_Jo = tmp_madorijo(cntkk)

                                                '固定値
                                                .Vari_History = DefHistory
                                                .Vari_Syozaikai = ""

                                                'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                                                tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                                                'データチェック
                                                Dim skipflg As Boolean = False
                                                Dim hash_cvitem As New Hashtable
                                                Dim hash_log As New Hashtable
                                                skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                                                '書込処理
                                                If Not skipflg Then

                                                    'キーをguidへ変換
                                                    hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))

                                                    '挿入処理
                                                    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                                    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                                    If normalflg Then
                                                        list_chkduplicate.Add(fldvalue_key)
                                                        cvitemcnt = cvitemcnt + 1
                                                    End If

                                                End If

                                                '-------------------
                                                'ログ出力
                                                '-------------------

                                                'ログ出力メッセージ整形
                                                Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                                                'ログ出力
                                                If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                                                    Dim tmp_cnt As Integer = 0
                                                    '挿入
                                                    For Each logvalue In sortlist_log
                                                        Dim tmp_sql_insert As String = logvalue.Value
                                                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                                    Next
                                                    '初期化
                                                    tmp_logcnt = 0
                                                    sortlist_log.Clear()
                                                End If

                                            Next

                                        End If

                                    End If

                                Next

                                '--------------------------------------------------------------------
                                '移行した間取内訳が1データ以上ある場合移行したレコードの数を更新する
                                '--------------------------------------------------------------------
                                If cvitemcnt > 0 Then
                                    tmp_cvcnt = tmp_cvcnt + 1
                                End If

                            Case ConvertTypes._汎用

                                'データ取得
                                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                                    '項目名取得
                                    fldname = readtbl.Columns(cntjj).ColumnName.Trim

                                    '登録値取得
                                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                                    Select Case fldname
                                        Case "間取り内訳No"
                                            .Vari_No = fldvalue.Trim
                                        Case "間取り内訳区分"
                                            .Vari_Hykbn = fldvalue.Trim
                                        Case "畳数"
                                            .Vari_Jo = fldvalue.Trim
                                        Case "所在階"
                                            .Vari_Syozaikai = fldvalue.Trim
                                    End Select

                                Next

                                '固定値
                                .Vari_History = DefHistory

                                'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                                tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                                'データチェック
                                Dim skipflg As Boolean = False
                                Dim hash_cvitem As New Hashtable
                                Dim hash_log As New Hashtable
                                skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                                '書込処理
                                If Not skipflg Then

                                    'キーをguidへ変換
                                    hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))

                                    '挿入処理
                                    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                    If normalflg Then
                                        list_chkduplicate.Add(fldvalue_key)
                                        tmp_cvcnt = tmp_cvcnt + 1
                                    End If

                                End If

                                '-------------------
                                'ログ出力
                                '-------------------

                                'ログ出力メッセージ整形
                                Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                                'ログ出力
                                If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                                    Dim tmp_cnt As Integer = 0
                                    '挿入
                                    For Each logvalue In sortlist_log
                                        Dim tmp_sql_insert As String = logvalue.Value
                                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                    Next
                                    '初期化
                                    tmp_logcnt = 0
                                    sortlist_log.Clear()
                                End If

                        End Select
                        '20160915 間取内訳情報の汎用/既存処理の分岐 -chg end

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 間取内訳行Noの一括更新クエリ
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Update() As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " UPDATE hydata_madoriutiwake SET "
                tmp_sql = tmp_sql & " 	no = MADORINO.連番 "
                tmp_sql = tmp_sql & " FROM hydata_madoriutiwake AS HYMADORI "
                tmp_sql = tmp_sql & " LEFT JOIN "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 hy_guid "
                tmp_sql = tmp_sql & " 		,no "
                tmp_sql = tmp_sql & " 		,ROW_NUMBER()OVER(PARTITION BY hy_guid ORDER BY no) AS 連番 "
                tmp_sql = tmp_sql & " 	FROM hydata_madoriutiwake "
                tmp_sql = tmp_sql & " ) AS MADORINO "
                tmp_sql = tmp_sql & " ON  HYMADORI.hy_guid = MADORINO.hy_guid "
                tmp_sql = tmp_sql & " AND HYMADORI.no = MADORINO.no "
                tmp_sql = tmp_sql & " ; "

                Return tmp_sql

            End Function

        End Class

    End Class

#End Region

#Region "部屋修繕維持管理連絡先情報"

    Public Class Hydata_szeniji_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
                Dim startrow As Integer                                         '書込開始行
                Dim columncnt As Integer                                        '列数
                Dim maxrowcnt As Integer                                        '既存データの行数
                Dim rowcnt As Integer                                           '書込行数
                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納

                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_szeniji_Model          '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_szeniji"
                Dim fldnamegrp As String = "hy_guid,syuzenijikanri_no,syuzenijikanri_kasyo,syuzenijikanri_taisyokbn,syuzenijikanri_gyno," & _
                                           "syuzenijikanri_simei,syuzenijikanri_address,syuzenijikanri_tel,syuzenijikanri_jisyano,syuzenijikanri_kasino," & _
                                           "syuzenijikanri_syono,syuzenijikanri_simeiu"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 1 To rowcnt

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                            Return rtn
                        End If

                        '行取得
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                        'キー値取得
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        '作業用変数 
                        Dim tmp_kasyoname As String = ""    '20160829 箇所未登録データにデフォルト値を設定 -add

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "修繕及び維持管理No"
                                    .Vari_Syuzenijikanri_no = fldvalue.Trim
                                Case "修繕及び維持管理の箇所"
                                    '20160829 箇所未登録データにデフォルト値を設定 -chg sta
                                    '.Vari_Syuzenijikanri_kasyo = fldvalue.Trim
                                    tmp_kasyoname = fldvalue.Trim
                                    '20160829 箇所未登録データにデフォルト値を設定 -chg end
                                Case "対象区分"
                                    .Vari_Syuzenijikanri_taisyokbn = fldvalue.Trim
                                Case "業者no"
                                    .Vari_Syuzenijikanri_gyno = fldvalue.Trim
                                Case "氏名(商号または名称)(SJIS)"
                                    .Vari_Syuzenijikanri_simei = fldvalue.Trim
                                Case "住所(主たる事務所の所在地)"
                                    .Vari_Syuzenijikanri_address = fldvalue.Trim
                                Case "連絡先電話番号"
                                    .Vari_Syuzenijikanri_tel = fldvalue.Trim
                                Case "自社no"
                                    .Vari_Syuzenijikanri_jisyano = fldvalue.Trim
                                Case "貸主No"
                                    .Vari_Syuzenijikanri_kasino = fldvalue.Trim
                                Case "所有者No"
                                    .Vari_Syuzenijikanri_syono = fldvalue.Trim
                                Case "氏名(商号または名称)"
                                    .Vari_Syuzenijikanri_simeiu = fldvalue.Trim
                            End Select

                        Next

                        '20160829 箇所未登録データにデフォルト値を設定 -add
                        .Vari_Syuzenijikanri_kasyo = tmp_kasyoname & .Vari_Syuzenijikanri_no

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20160829 箇所未登録データにデフォルト値を設定 -add sta
                        If tmp_kasyoname = "" Then
                            tmp_hash("syuzenijikanri_kasyo") = ""
                        End If
                        '20160829 箇所未登録データにデフォルト値を設定 -add end

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii
                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,syuzenijikanri_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋設備情報"

    Public Class Hydata_setubilst_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
                Dim startrow As Integer                                         '書込開始行
                Dim columncnt As Integer                                        '列数
                Dim maxrowcnt As Integer                                        '既存データの行数
                Dim rowcnt As Integer                                           '書込行数
                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納

                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_chkduplicateguid As New List(Of String)                '重複チェック用リスト (移行後の重複チェック用)   '2016.04.26 メインの方へも反映させる修正
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_setubilst_Model        '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                Dim cnt_cvcnt As Integer = 0                                    '20160603 ユーザーデータ検証による修正 -add

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '2016.04.06 部屋設備情報の取得処理修正 -chg sta
                'Call Me.Get_RelData()
                Call SetRelItemToObject.Set_RelData_Setubi_V7To10No()
                Call EtcMethod.Set_SetubiMst(sqlcnnv10)
                Call EtcMethod.Set_SetubiMst_UserMake(sqlcnnv10) '20160829 設備の新規挿入処理を追加 -add
                Call Me.Set_RelData_Setubi_V7To10Guid()
                '2016.04.06 部屋設備情報の取得処理修正 -chg end

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_setubilst"
                Dim fldnamegrp As String = "hy_guid,komok_guid,override_kbn,komok_name,disp1name," & _
                                           "disp2name,disp3name,disp1iconguid,disp2iconguid,disp3iconguid," & _
                                           "history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicateguid)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 1 To rowcnt

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                            Return rtn
                        End If

                        '行取得
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                        'キー値取得
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        '作業用変数
                        Dim tmp_setubikomkname As String = ""
                        Dim tmp_setubinaiyo As String = ""

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "設備項目名"
                                    tmp_setubikomkname = fldvalue.Trim
                                Case "設備内容"
                                    tmp_setubinaiyo = fldvalue.Trim
                            End Select

                        Next

                        '項目guidへ一時格納
                        .Vari_Komok_guid = tmp_setubikomkname & "-" & tmp_setubinaiyo

                        '固定値
                        .Vari_Override_kbn = 0
                        .Vari_Komok_name = ""
                        .Vari_Disp1name = ""
                        .Vari_Disp2name = ""
                        .Vari_Disp3name = ""
                        .Vari_Disp1iconguid = "00000000-0000-0000-0000-000000000000"
                        .Vari_Disp2iconguid = "00000000-0000-0000-0000-000000000000"
                        .Vari_Disp3iconguid = "00000000-0000-0000-0000-000000000000"
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            '2016.04.26 メインの方へも反映させる修正 -chg sta
                            'guidへ変換した値での重複チェックを行う
                            '↓↓↓旧srcコメントアウト↓↓↓
                            ''キーをguidへ変換
                            'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))

                            ''挿入処理
                            'Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            ''挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            'If normalflg Then
                            '    list_chkduplicate.Add(fldvalue_key)
                            '    tmp_cvcnt = tmp_cvcnt + 1
                            'End If
                            '↑↑↑旧srcコメントアウト↑↑↑

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            'guidでの重複チェック
                            Dim tmp_hyguid As String = hash_cvitem("hy_guid")
                            Dim tmp_komkguid As String = hash_cvitem("komok_guid")

                            '20160829 エレベーター移行対応 -chg sta
                            ''20160603 ユーザーデータ検証による修正 -add sta
                            ''紐付けられた値のみ移行件数をカウントする
                            'If tmp_komkguid <> "" Then
                            '    cnt_cvcnt = cnt_cvcnt + 1
                            'End If
                            ''20160603 ユーザーデータ検証による修正 -add end

                            'If list_chkduplicateguid.Contains(tmp_hyguid & "-" & tmp_komkguid) = False Then

                            '    '挿入処理
                            '    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            '    If normalflg Then
                            '        list_chkduplicate.Add(fldvalue_key)
                            '        list_chkduplicateguid.Add(tmp_hyguid & "-" & tmp_komkguid)
                            '        tmp_cvcnt = tmp_cvcnt + 1
                            '    End If

                            'Else

                            '    'ログ出力
                            '    hash_log.Clear()    '重複エラーを優先する
                            '    Dim tmp_tblfldvalue As String = tblname & "-" & "komok_guid"
                            '    hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)

                            'End If
                            ''2016.04.26 メインの方へも反映させる修正 -chg end

                            'エレーベータの値をセット
                            Dim list_elv As New List(Of String) From {"1", "2", "3"}
                            If list_elv.Contains(tmp_komkguid) Then
                                '物件情報のエレベーターを更新する
                                Call Set_Bkdatadetail_ElvFlg(sqlcnnv10, tmp_hyguid, tmp_komkguid)
                            Else
                                '紐付けられた値のみ移行件数をカウントする
                                If tmp_komkguid <> "" Then
                                    cnt_cvcnt = cnt_cvcnt + 1
                                End If

                                If list_chkduplicateguid.Contains(tmp_hyguid & "-" & tmp_komkguid) = False Then

                                    '挿入処理
                                    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                    If normalflg Then
                                        list_chkduplicate.Add(fldvalue_key)
                                        list_chkduplicateguid.Add(tmp_hyguid & "-" & tmp_komkguid)
                                        tmp_cvcnt = tmp_cvcnt + 1
                                    End If

                                Else

                                    'ログ出力
                                    hash_log.Clear()    '重複エラーを優先する
                                    Dim tmp_tblfldvalue As String = tblname & "-" & "komok_guid"
                                    hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)

                                End If

                            End If
                            '20160829 エレベーター移行対応 -chg end
                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii
                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '20160603 ユーザーデータ検証による修正 -chg sta
                ''中間ファイル件数を取得
                'midrowcnt = rowcnt
                '移行件数を取得
                midrowcnt = cnt_cvcnt
                '20160603 ユーザーデータ検証による修正 -chg end

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                'tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parking_kbn) FROM " & tblname
                'tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                'Dim flg As Boolean = True

                'flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

                '2016.04.26 メインの方へも反映させる修正 -add sta
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(VARCHAR(200),hy_guid) + '-' + CONVERT(VARCHAR(200),komok_guid) "
                tmp_sql = tmp_sql & " FROM hydata_setubilst "
                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)
                '2016.04.26 メインの方へも反映させる修正 -add end


            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            '2016.04.06 部屋設備情報の取得処理修正 -del sta
            ' ''' <summary>
            ' ''' 設備紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "設備マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_setubigrpname As String = ""
            '        Dim tmp_setubiname As String = ""
            '        Dim tmp_setubiguid As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元設備グループ名称"
            '                    tmp_setubigrpname = fldvalue.Trim
            '                Case "移行元設備名称"
            '                    tmp_setubiname = fldvalue.Trim
            '                Case "賃貸革命10項目GUID"
            '                    tmp_setubiguid = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_setubiguid <> "" Then
            '            Hash_Rel_Setubi.Add(tmp_setubigrpname & "-" & tmp_setubiname, tmp_setubiguid)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub
            '2016.04.06 部屋設備情報の取得処理修正 -del end

            ''' <summary>
            ''' V7設備データと10設備項目guidを紐付 2016.04.06 部屋設備情報の取得処理修正 -add sta
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub Set_RelData_Setubi_V7To10Guid()

                '初期化
                Hash_Rel_Setubi.Clear()

                For Each setubiitem In Hash_SetubiMid

                    Dim tmp_V7setubi As String = setubiitem.Key
                    Dim tmp_10setubino As String = setubiitem.Value

                    If Hash_SetubiMst.Contains(tmp_10setubino) Then
                        Dim tmp_10setubiguid As String = Hash_SetubiMst(tmp_10setubino)
                        Hash_Rel_Setubi.Add(tmp_V7setubi, tmp_10setubiguid)
                    ElseIf tmp_10setubino = "" Then     '20160829 設備の新規挿入処理を追加 -add
                        If Hash_SetubiMst_UserMake.Contains(tmp_V7setubi) Then
                            Dim tmp_10setubiguid_usermake As String = Hash_SetubiMst_UserMake(tmp_V7setubi)
                            Hash_Rel_Setubi.Add(tmp_V7setubi, tmp_10setubiguid_usermake)
                        End If
                    End If

                Next

            End Sub
            '2016.04.06 部屋設備情報の取得処理修正 -add end

            ''' <summary>
            ''' 物件詳細情報のエレベーターフラグを更新 '20160829 エレベーター移行対応
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hyguid"></param>
            ''' <remarks></remarks>
            Public Sub Set_Bkdatadetail_ElvFlg(ByVal sqlcnnv10 As SqlConnection, ByVal hyguid As String, ByVal elvvalue As String)

                '部屋guidが存在しない場合は処理を抜ける
                If hyguid = "" Then
                    Exit Sub
                End If

                'bkguidを取得
                Dim tmp_sql As String = " SELECT CONVERT(VARCHAR(MAX),bk_guid) FROM hydata WHERE hy_guid = '" & hyguid & "'"
                Dim tmp_bkguid As String = DBExec.Exec_Scalar(tmp_sql, sqlcnnv10)

                '更新クエリを作成
                Dim tmp_updatefld As String = ""
                Select Case elvvalue
                    Case "1"
                        tmp_updatefld = "elevator_flg"
                    Case "2"
                        tmp_updatefld = "homeelevator_flg"
                    Case "3"
                        tmp_updatefld = "emergencyelevator_flg"
                End Select

                '更新クエリ作成
                Dim tmp_sqlupdate As String = ""
                tmp_sqlupdate = tmp_sqlupdate & " UPDATE bkdata_detail SET " & tmp_updatefld & " = 1 "
                tmp_sqlupdate = tmp_sqlupdate & " WHERE bk_guid = '" & tmp_bkguid & "'"

                '実行
                Dim tmpcnt As Integer = 0
                Dim flg As Boolean = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqlupdate, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋入金項目情報"

    Public Class Hydata_nkin_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_nkin_Model                 '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                      'メインキー列
                Dim keycol_sub1 As Integer = 2                                      'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                      'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                      'サブ3キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '2016.04.06 入金項目読込処理の修正 -chg sta
                'Call Me.Get_RelData()
                Call SetRelItemToObject.Set_RelData_Nkinkomk()
                '2016.04.06 入金項目読込処理の修正 -chg end
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_nkin"
                Dim fldnamegrp As String = "hy_guid,tuki_kbn,nkin_no,nkin_recno,nkin_kbn," & _
                                           "sq_gak,sq_zeikbn,calc_kbn,calc_nkinno,calc_monthcnt," & _
                                           "sqsaki_no,sq_mmkbn,sqstart_ymd,sq_ptn,sq_interval," & _
                                           "sq_nen,sq_tuki,biko,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub1 As String = readtbl.Columns(keycol_sub1 - 1).ColumnName.Trim
                    Dim fldname_keysub2 As String = readtbl.Columns(keycol_sub2 - 1).ColumnName.Trim
                    Dim fldname_keysub3 As String = readtbl.Columns(keycol_sub3 - 1).ColumnName.Trim

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim tmp_keysub3 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub3 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        '20161125 ログに入金項目名を追加 -chg sta
                        'Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2 & "、" & fldname_keysub3 & " = " & tmp_keysub3
                        '20161125 ログに入金項目名を追加 -chg end

                        '2016.04.06 入金項目読込処理の修正 -add
                        '作業用変数
                        Dim tmp_nkinname As String = ""

                        '20160530 ログ修正 -add
                        Dim tmp_nkinname_kijyun As String = ""


                        '2016.04.26 メインの方へも反映させる修正 -add
                        Dim tmp_sqkankaku As String = ""

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "月区分"
                                    .Vari_Tuki_kbn = fldvalue
                                Case "入金項目名"
                                    '2016.04.06 入金項目読込処理の修正 -chg sta
                                    '.Vari_Nkin_no = fldvalue.Trim
                                    tmp_nkinname = fldvalue
                                    '2016.04.06 入金項目読込処理の修正 -chg end
                                Case "入金項目行No"
                                    .Vari_Nkin_recno = fldvalue
                                Case "入金項目区分"
                                    .Vari_Nkin_kbn = fldvalue
                                Case "請求額"
                                    .Vari_Sq_gak = fldvalue
                                Case "税区分"
                                    .Vari_Sq_zeikbn = fldvalue
                                Case "算出区分"
                                    .Vari_Calc_kbn = fldvalue
                                Case "算出基準入金項目名"
                                    '20160530 ログ修正 -chg sta
                                    '.Vari_Calc_nkinno = fldvalue
                                    tmp_nkinname_kijyun = fldvalue
                                    '20160530 ログ修正 -chg end
                                Case "算出ヶ月"
                                    .Vari_Calc_monthcnt = fldvalue
                                Case "請求先No"
                                    .Vari_Sqsaki_no = fldvalue
                                Case "請求月区分"
                                    .Vari_Sq_mmkbn = fldvalue
                                Case "請求開始月"
                                    .Vari_Sqstart_ymd = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -del sta
                                    '必要な情報をまとめて格納するためここではコメントアウト
                                    'Case "請求パターン"
                                    '    .Vari_Sq_ptn = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -del end
                                Case "固定公共料金で使用"
                                    '2016.04.26 メインの方へも反映させる修正 -chg sta
                                    '請求間隔は変換して移行するため作業用変数へ格納しておく
                                    '.Vari_Sq_interval = fldvalue
                                    tmp_sqkankaku = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -chg end

                                    '2016.04.26 メインの方へも反映させる修正 -del sta
                                    '必要な情報をまとめて格納するためここではコメントアウト
                                    'Case "請求発生年"
                                    '    .Vari_Sq_nen = fldvalue
                                    'Case "請求発生月"
                                    '    .Vari_Sq_tuki = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -del end
                                Case "備考"
                                    .Vari_Biko = fldvalue
                            End Select

                        Next

                        '2016.04.06 入金項目読込処理の修正 -add sta
                        '紐付用に成形
                        .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname(.Vari_Nkin_kbn)
                        '2016.04.06 入金項目読込処理の修正 -add end

                        '20160530 ログ修正 -add sta
                        If tmp_nkinname_kijyun <> "" Then
                            .Vari_Calc_nkinno = tmp_nkinname_kijyun & "-" & EtcMethod.Get_Nkinruiname("1")
                        Else
                            .Vari_Calc_nkinno = ""
                        End If
                        '20160530 ログ修正 -add end

                        '2016.04.26 メインの方へも反映させる修正 -add sta
                        '変動費で固定項目の請求間隔、請求発生年、発生月を取得するために必要なデータを格納しておく
                        '必要な情報
                        '移行項目(部屋と契約入金項目情報で移行内容が異なるため)、請求パターン、請求月、請求開始月、請求間隔
                        .Vari_Sq_ptn = sheetname & "-" & .Vari_Sq_mmkbn & "-" & .Vari_Sqstart_ymd & "-" & tmp_sqkankaku
                        .Vari_Sq_interval = sheetname & "-" & .Vari_Sq_mmkbn & "-" & .Vari_Sqstart_ymd & "-" & tmp_sqkankaku
                        .Vari_Sq_nen = sheetname & "-" & .Vari_Sq_mmkbn & "-" & .Vari_Sqstart_ymd & "-" & tmp_sqkankaku
                        .Vari_Sq_tuki = sheetname & "-" & .Vari_Sq_mmkbn & "-" & .Vari_Sqstart_ymd & "-" & tmp_sqkankaku
                        '2016.04.26 メインの方へも反映させる修正 -add end

                        '固定値
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                '入金項目の重複チェックは特殊なので考慮すること

                'Dim tmp_sql As String = ""
                'tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parking_kbn) FROM " & tblname
                'tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                'Dim flg As Boolean = True

                'flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            '2016.04.06 入金項目読込処理の修正 -del sta
            ''' <summary>
            ''' 入金項目紐付情報取得
            ''' </summary>
            ''' <remarks></remarks>
            'Public Sub Get_RelData()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "入金項目マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    '既存データ有無確認(入金項目は他の項目でも参照するため)
            '    If Hash_Rel_Nkinkomk.Count <> 0 Then
            '        Exit Sub
            '    End If

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldnkinname As String = ""
            '        Dim tmp_newnkinno As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元入金項目名称"
            '                    tmp_oldnkinname = fldvalue.Trim
            '                Case "賃貸革命10入金項目No"
            '                    tmp_newnkinno = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newnkinno <> "" Then
            '            Hash_Rel_Nkinkomk.Add(tmp_oldnkinname, tmp_newnkinno)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub

            '2016.04.06 入金項目読込処理の修正 -del end

        End Class

    End Class

#End Region

#Region "部屋変動費各戸メーター情報"

    Public Class Hydata_hendo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
                Dim startrow As Integer                                         '書込開始行
                Dim columncnt As Integer                                        '列数
                Dim maxrowcnt As Integer                                        '既存データの行数
                Dim rowcnt As Integer                                           '書込行数
                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納

                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_hendo_Model            '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************


                '入金項目の重複チェックは特殊なので考慮すること






                '紐付けデータ取得
                '2016.04.06 入金項目読込処理の修正 -chg sta
                'Call Me.Get_RelData_Nkinkomk()
                'Call Me.Get_RelData_Hendometer()
                Call SetRelItemToObject.Set_RelData_Nkinkomk()
                '2016.04.06 入金項目読込処理の修正 -chg end

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_hendo"
                Dim fldnamegrp As String = "hy_guid,hyhendo_guid,rec_no,hendo_kbn,meter_name," & _
                                           "nkin_no,biko,hendorule_no,hendorule_biko,history," & _
                                           "sq_mmkbn,useflg"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 1 To rowcnt

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                            Return rtn
                        End If

                        '行取得
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                        'キー値取得
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "行No"
                                    .Vari_Rec_no = fldvalue
                                Case "メーター分類"
                                    '.Vari_Hendo_kbn = fldvalue.Trim
                                Case "メーター名"
                                    .Vari_Meter_name = fldvalue
                                Case "入金項目名"      '2016.04.01 他の移行項目との統一のためヘッダー名を変更 変動費入金項目 → 入金項目名
                                    '2016.04.06 入金項目読込処理の修正 -chg sta
                                    '.Vari_Nkin_no = fldvalue.Trim
                                    '.Vari_Hendo_kbn = fldvalue.Trim     '紐付けたメーター分類を取得するため入金項目名を格納しておく
                                    .Vari_Nkin_no = fldvalue & "-" & EtcMethod.Get_Nkinruiname("5")
                                    .Vari_Hendo_kbn = fldvalue & "-" & EtcMethod.Get_Nkinruiname("5")     '紐付けたメーター分類を取得するため入金項目名を格納しておく
                                    '2016.04.06 入金項目読込処理の修正 -chg end
                                Case "備考"
                                    .Vari_Biko = fldvalue
                                Case "変動費請求ルールNo"
                                    .Vari_Hendorule_no = fldvalue
                                Case "変動ルール備考"
                                    .Vari_Hendorule_biko = fldvalue
                                Case "請求月区分"
                                    .Vari_Sq_mmkbn = fldvalue
                            End Select

                        Next

                        '固定値
                        .Vari_Hyhendo_guid = Guid.NewGuid.ToString
                        .Vari_History = DefHistory
                        .Vari_Useflg = 1        '変動費検針データに必要なため1をセット

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii
                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '2016.04.06 物件情報の変動費検針有無の一括更新処理追加 -add sta
                '変動費が存在する物件の検針業務有無フラグを一括更新
                Dim tmpcnt As Integer = 0
                Dim kensinumuupdatesql As String = Me.Get_UseQry_Update()
                DBExec.Exec_NonQuery(sqlcnnv10, kensinumuupdatesql, tmpcnt)
                '2016.04.06 物件情報の変動費検針有無の一括更新処理追加 -add end

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                '入金項目の重複チェックは特殊なので考慮すること

                'Dim tmp_sql As String = ""
                'tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parking_kbn) FROM " & tblname
                'tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                'Dim flg As Boolean = True

                'flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 入金項目紐付情報取得
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub Get_RelData_Nkinkomk()

                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
                Dim startrow As Integer                                         '書込開始行
                Dim columncnt As Integer                                        '列数
                Dim maxrowcnt As Integer                                        '既存データの行数
                Dim rowcnt As Integer                                           '書込行数
                Dim relsheetname As String = "入金項目マスタ"
                Dim rtn As Boolean = True

                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

                '既存データ有無確認(入金項目は他の項目でも参照するため)
                If Hash_Rel_Nkinkomk.Count <> 0 Then
                    Exit Sub
                End If

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

                For cntii = 0 To rowcnt - 1

                    '作業用変数作成
                    Dim tmp_oldnkinname As String = ""
                    Dim tmp_newnkinno As String = ""
                    Dim tmp_nkinkbn As String = ""
                    Dim tmp_hendometerno As String = ""

                    '紐付設定値取得
                    For cntjj = 1 To columncnt

                        'ヘッダー格納
                        Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                        '移行値格納
                        Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                        Select Case fldname
                            Case "移行元入金項目名称"
                                tmp_oldnkinname = fldvalue.Trim
                            Case "賃貸革命10入金項目No"
                                tmp_newnkinno = fldvalue.Trim
                                '2016.04.06 入金項目読込処理の修正 -add sta
                            Case "移行元入金項目区分"
                                tmp_nkinkbn = fldvalue.Trim
                            Case "賃貸革命10変動費メーター分類No"
                                tmp_hendometerno = fldvalue.Trim
                                '2016.04.06 入金項目読込処理の修正 -add end
                        End Select

                    Next

                    '入金項目をハッシュテーブル格納
                    If tmp_newnkinno <> "" Then
                        Hash_Rel_Nkinkomk.Add(tmp_oldnkinname, tmp_newnkinno)
                    End If

                    '2016.04.06 入金項目読込処理の修正 -add sta
                    '変動費メーター分類をハッシュテーブル格納
                    If tmp_nkinkbn = "5.随時変動" And tmp_hendometerno <> "" Then
                        Hash_Rel_Hendometer.Add(tmp_oldnkinname, tmp_hendometerno)
                    End If
                    '2016.04.06 入金項目読込処理の修正 -add end

                Next

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            End Sub

            '2016.04.06 入金項目読込処理の修正 -del sta
            ''' <summary>
            ''' 各戸メーター分類紐付情報取得
            ''' </summary>
            ''' <remarks></remarks>
            'Public Sub Get_RelData_Hendometer()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "各戸メーター分類"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    '既存データ有無確認(入金項目は他の項目でも参照するため)
            '    If Hash_Rel_Hendometer.Count <> 0 Then
            '        Exit Sub
            '    End If

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldnkinname As String = ""
            '        Dim tmp_newmeterno As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元入金項目名称"
            '                    tmp_oldnkinname = fldvalue.Trim
            '                Case "賃貸革命10メーター分類No"
            '                    tmp_newmeterno = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newmeterno <> "" Then
            '            Hash_Rel_Hendometer.Add(tmp_oldnkinname, tmp_newmeterno)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub
            '2016.04.06 入金項目読込処理の修正 -del end

            ''' <summary>
            ''' [bkdata_detail].[kensingyomu_umu]の一括更新クエリ  2016.04.06 物件情報の変動費検針有無の一括更新処理追加
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Update() As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " UPDATE bkdata_detail SET "
                tmp_sql = tmp_sql & " 	kensingyomu_umu = 1 "
                tmp_sql = tmp_sql & " WHERE bk_guid IN "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT bk_guid FROM hydata AS HY "
                tmp_sql = tmp_sql & " 	WHERE hy_guid IN "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT hy_guid FROM hydata_hendo AS HYH "
                tmp_sql = tmp_sql & " 		) "
                tmp_sql = tmp_sql & " ); "

                Return tmp_sql

            End Function

        End Class

    End Class

#End Region

#Region "部屋メモ情報"

    Public Class Hydata_memo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_memo_Model             '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列
                Dim viewname As String = "物件部屋キー情報"
                Dim totalrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

                '************************
                '作業準備
                '************************
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_memo"
                Dim fldnamegrp As String = "hy_guid,memo_no,memo,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '移行対象件数取得用
                Dim cvtaisyocnt As Integer = 0

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub1 As String = readtbl.Columns(keycol_sub - 1).ColumnName.Trim

                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '移行判別用
                        Dim cvflg As Boolean = True
                        Dim keychkflg As Boolean = True

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub - 1)).Trim
                        Dim tmp_keytotal As String = tmp_keymain & "-" & tmp_keysub1
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        '備考カウント初期化
                        Dim cvitemcnt As Integer = 0
                        Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

                        'データ有無チェック
                        Dim tmp_fldvalueumuchk As String = ""
                        For cntjj = 2 To readtbl.Columns.Count - 1
                            tmp_fldvalueumuchk = tmp_fldvalueumuchk & Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                        Next
                        If tmp_fldvalueumuchk = "" Then
                            cvflg = False
                        End If

                        If cvflg Then

                            '移行対象件数カウント
                            cvtaisyocnt = cvtaisyocnt + 1

                            Dim hash_keylog As New Hashtable
                            Dim log_keyout As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1
                            '親データ有無チェック
                            If list_basekeydata.Contains(tmp_keytotal) = False Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "hydata_memo-hy_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(tmp_keytotal) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "hydata_memo-hy_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            If keychkflg Then

                                '移行値取得
                                For cntjj = 2 To readtbl.Columns.Count - 1

                                    'サブキー取得
                                    Dim tmp_keysub2 As String = (cntjj - 1).ToString

                                    '全キー取得
                                    Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                                    'ログ出力用データ格納(サブキーフィールド)
                                    Dim fldname_keysub2 As String = readtbl.Columns(cntjj).ColumnName.Trim
                                    Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2

                                    '登録値取得
                                    Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                                    .Vari_Memo = fldvalue

                                    '固定値
                                    .Vari_Memo_no = (cntjj - 1).ToString
                                    .Vari_History = DefHistory

                                    'メモにデータが存在する場合に書込処理を行う
                                    If .Vari_Memo <> "" Then

                                        '20160928 メモ関連の移行件数表示修正 -add
                                        tmp_cvrowcnt = tmp_cvrowcnt + 1

                                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                                        'データチェック
                                        Dim skipflg As Boolean = False
                                        Dim hash_cvitem As New Hashtable
                                        Dim hash_log As New Hashtable
                                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                                        '書込処理
                                        If Not skipflg Then

                                            'キーをguidへ変換
                                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                                            '挿入処理
                                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                            If normalflg Then
                                                If list_chkduplicate.Contains(tmp_keytotal) = False Then
                                                    list_chkduplicate.Add(tmp_keytotal)
                                                End If
                                                cvitemcnt = cvitemcnt + 1
                                            End If

                                        End If

                                        'ログ出力メッセージ整形
                                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                                        'ログ出力
                                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                                            Dim tmp_cnt As Integer = 0
                                            '挿入
                                            For Each logvalue In sortlist_log
                                                Dim tmp_sql_insert As String = logvalue.Value
                                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                            Next
                                            '初期化
                                            tmp_logcnt = 0
                                            sortlist_log.Clear()
                                        End If

                                    End If

                                Next

                            Else

                                '親データ無し、重複チェックエラーログ出力
                                'ログ出力メッセージ整形
                                tmp_hash.Clear()
                                Call LogSetting.Set_Log_Value_KomkErr(hash_keylog, tmp_hash, tmp_hash, log_keyout, tblname, tmp_logcnt, sortlist_log)  '※引数のtmp_hashは使用しないが仮で入れておく

                                'ログ出力
                                If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = totalrowcnt - 1) Then
                                    Dim tmp_cnt As Integer = 0
                                    '挿入
                                    For Each logvalue In sortlist_log
                                        Dim tmp_sql_insert As String = logvalue.Value
                                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                    Next
                                    '初期化
                                    tmp_logcnt = 0
                                    sortlist_log.Clear()
                                End If

                            End If

                        End If

                        '--------------------------------------------------------------------
                        '移行した備考が1データ以上ある場合移行したレコードの数を更新する
                        '--------------------------------------------------------------------
                        '20160928 メモ関連の移行件数表示修正 -chg sta
                        'tmp_cvcnt = tmp_cvcnt + 1
                        If tmp_cvrowcnt > 0 Then
                            totalrowcnt = totalrowcnt + 1
                        End If
                        If cvitemcnt > 0 Then
                            tmp_cvcnt = tmp_cvcnt + 1
                        End If
                        '20160928 メモ関連の移行件数表示修正 -chg end
                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = cvtaisyocnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,memo_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋共通セールスポイント情報"

    Public Class Hydata_commonsalespointparts_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_commonsalespointparts_Model    '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                          'メインキー列
                Dim keycol_sub1 As Integer = 2                                          'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                          'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                ''Excelファイル初期設定                  
                'Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                'Dim readtbl As New DataTable()
                'Dim con_read As New OleDbConnection()

                ''オープン処理
                'rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                ''オープン処理失敗時は処理を抜ける
                'If rtn = False Then
                '    excelfile.ExcelFile_ReadClose(con_read)
                '    Return rtn
                'End If

                ''行数取得
                'Dim rowcnt As Integer = readtbl.Rows.Count

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_commonsalespointparts"
                Dim fldnamegrp As String = "hy_guid,parts_no,salespoint_parts"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub1 As String = readtbl.Columns(keycol_sub1 - 1).ColumnName.Trim
                    Dim fldname_keysub2 As String = readtbl.Columns(keycol_sub2 - 1).ColumnName.Trim

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                            'excelfile.ExcelFile_ReadClose(con_read)
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "パーツNO"
                                    .Vari_Parts_no = fldvalue.Trim
                                Case "セールスポイントパーツ"
                                    .Vari_Salespoint_parts = fldvalue.Trim
                            End Select

                        Next

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        'Dim pgbcnt As Integer = 0
                        'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                        '    pgbcnt = cntii
                        'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii + 1
                        ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parts_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋契約解約確認事項情報"

    Public Class Hydata_confirm_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
                Dim startrow As Integer                                         '書込開始行
                Dim columncnt As Integer                                        '列数
                Dim maxrowcnt As Integer                                        '既存データの行数
                Dim rowcnt As Integer                                           '書込行数
                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納

                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_confirm_Model          '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_confirm"
                Dim fldnamegrp As String = "hy_guid,confirm_kbn,confirm_no,naiyo"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)
                    Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 1 To rowcnt

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                            Return rtn
                        End If

                        '行取得
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                        'キー値取得
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim tmp_keysub3 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub3))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "確認事項区分"
                                    .Vari_Confirm_kbn = fldvalue.Trim
                                Case "確認事項No"
                                    .Vari_Confirm_no = fldvalue.Trim
                                Case "内容"
                                    .Vari_Naiyo = fldvalue.Trim
                            End Select

                        Next

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii
                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,confirm_kbn)  + '-' + CONVERT(varchar,confirm_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋権利情報"

    Public Class Hydata_kenri_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim appli As Excel.Application = Nothing                                'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                                   'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                                 'Excelオブジェクト
                Dim startrow As Integer                                                 '書込開始行
                Dim columncnt As Integer                                                '列数
                Dim maxrowcnt As Integer                                                '既存データの行数
                Dim rowcnt As Integer                                                   '書込行数
                Dim tmp_cvcnt As Integer                                                '移行件数格納
                Dim tmp_condcnt As Integer                                              '調整件数格納

                Dim excelfile As New Njc.Common.ExcelFileManager                        'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                        'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                          '共通処理用
                Dim list_chkduplicate As New List(Of String)                            '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                             '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)                  'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                           'ログ出力時のソート用
                Dim hash_guid As New Hashtable                                          'guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                           '作業用ハッシュテーブル
                Dim normalflg As Boolean                                                'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                               '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_kenri_Model                    '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                          'メインキー列
                Dim keycol_sub1 As Integer = 2                                          'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                          'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_kenri"
                Dim fldnamegrp As String = "hy_guid,kenri_no,other_kenrirui,other_kenribiko,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 1 To rowcnt

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                            Return rtn
                        End If

                        '行取得
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                        'キー値取得
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "権利情報No"
                                    .Vari_Kenri_no = fldvalue.Trim
                                Case "所有権以外の権利の種類"
                                    .Vari_Other_kenrirui = fldvalue.Trim
                                Case "所有権以外の権利備考"
                                    .Vari_Other_kenribiko = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii
                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,kenri_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋参照ファイル情報"

    Public Class Hydata_relfile_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim appli As Excel.Application = Nothing                                'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                                   'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                                 'Excelオブジェクト
                Dim startrow As Integer                                                 '書込開始行
                Dim columncnt As Integer                                                '列数
                Dim maxrowcnt As Integer                                                '既存データの行数
                Dim rowcnt As Integer                                                   '書込行数
                Dim tmp_cvcnt As Integer                                                '移行件数格納
                Dim tmp_condcnt As Integer                                              '調整件数格納

                Dim excelfile As New Njc.Common.ExcelFileManager                        'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                        'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                          '共通処理用
                Dim list_chkduplicate As New List(Of String)                            '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                             '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)                  'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                           'ログ出力時のソート用
                Dim hash_guid As New Hashtable                                          'guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                           '作業用ハッシュテーブル
                Dim normalflg As Boolean                                                'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                               '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_relfile_Model                  '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                          'メインキー列
                Dim keycol_sub1 As Integer = 2                                          'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                          'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_relfile"
                Dim fldnamegrp As String = "hy_guid,file_no,fullpath,addtime,biko," & _
                                           "history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)


                '************************
                '処理開始
                '************************
                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 1 To rowcnt

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                            Return rtn
                        End If

                        '行取得
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                        'キー値取得
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "ファイルNo"
                                    .Vari_File_no = fldvalue.Trim
                                Case "フルパス"
                                    .Vari_Fullpath = fldvalue.Trim
                                Case "追加時間"
                                    .Vari_Addtime = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii
                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,file_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋原状回復目安単価情報"

    Public Class Hydata_szen_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim appli As Excel.Application = Nothing                                'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                                   'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                                 'Excelオブジェクト
                Dim startrow As Integer                                                 '書込開始行
                Dim columncnt As Integer                                                '列数
                Dim maxrowcnt As Integer                                                '既存データの行数
                Dim rowcnt As Integer                                                   '書込行数
                Dim tmp_cvcnt As Integer                                                '移行件数格納
                Dim tmp_condcnt As Integer                                              '調整件数格納

                Dim excelfile As New Njc.Common.ExcelFileManager                        'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                        'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                          '共通処理用
                Dim list_chkduplicate As New List(Of String)                            '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                             '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)                  'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                           'ログ出力時のソート用
                Dim hash_guid As New Hashtable                                          'guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                           '作業用ハッシュテーブル
                Dim normalflg As Boolean                                                'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                               '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_szen_Model                     '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                          'メインキー列
                Dim keycol_sub1 As Integer = 2                                          'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                          'サブ2キー列
                Dim viewname As String = "物件部屋キー情報"

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

                'テーブル名/フィールド名セット
                Dim viewname_base As String = PRE_VIEW_NAME & viewname
                Dim tblname As String = "hydata_szen"
                Dim fldnamegrp As String = "hy_guid,szen_no,szen_name,ryo,unit_name," & _
                                           "tanka"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 1 To rowcnt

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                            Return rtn
                        End If

                        '行取得
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                        'キー値取得
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "修繕No"
                                    .Vari_Szen_no = fldvalue.Trim
                                Case "修繕項目名"
                                    .Vari_Szen_name = fldvalue.Trim
                                Case "数量"
                                    .Vari_Ryo = fldvalue.Trim
                                Case "単位"
                                    .Vari_Unit_name = fldvalue.Trim
                                Case "単価"
                                    .Vari_Tanka = fldvalue.Trim
                            End Select

                        Next

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                        End If

                        '-------------------
                        'ログ出力
                        '-------------------

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                        '-------------------------------
                        'プログレスバー更新/進捗率表示
                        '-------------------------------
                        '件数取得
                        Dim pgbcnt As Integer = 0
                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                            pgbcnt = cntii
                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                            pgbcnt = pgbtotalcnt
                        End If

                        '表示
                        If pgbcnt <> 0 Then
                            Call obj_pgb.pgbsettingPart(pgbcnt)
                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        End If

                    Next

                End With

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_basedata"></param>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT キー FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,szen_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' guid取得→ハッシュテーブル格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="hash_guid"></param>
            ''' <remarks></remarks>
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

                Dim tmp_sql As String = "SELECT * FROM " & tblname
                Dim flg As Boolean = True

                '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                ''2016.04.26 メインの方へも反映させる修正 -chg sta
                ''flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                'Dim tmp_hash As New Hashtable
                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                ''取得した部屋Noを半角変換して照合用に統一するrtn_hash
                'hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                ''2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                '20160819 部屋の半角変換処理によるエラー修正 -chg end

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

                Dim viewname As String = "物件部屋キー情報"
                Dim tmpcnt As Integer = 0

                'VIEW初期化
                '20160531 部屋仮VIEWの初期化処理修正 -chg sta
                'Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(PRE_VIEW_NAME & viewname, False)
                '20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                'VIEW作成
                Dim tmp_sql_create As String = ""
                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region



#Region "部屋所有者情報"

#End Region

#Region "部屋画像情報"

#End Region

#Region "部屋業務期間情報"

#End Region

#Region "部屋BtoBグループ情報"

#End Region

#Region "部屋KeyValue情報"

#End Region

#Region "部屋広告補足情報(作り直し中)"

#End Region

#Region "部屋hydata_kys_atenaprint"

#End Region

#Region "部屋ポータル最終送信履歴情報"

#End Region

#Region "部屋セールスポイント情報"

#End Region

#Region "部屋連動サイト毎地図表示情報"

#End Region

#Region "部屋hydata_panorama情報"

#End Region

#Region "部屋支店情報"

#End Region

#Region "部屋hydata_sosin情報"

#End Region

#Region "部屋hydata_youtube情報"

#End Region

#Region "部屋親メーター関連情報"

#End Region

End Namespace


