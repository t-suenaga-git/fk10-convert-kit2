Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "送金ルール基本情報"

    Public Class Sorule_Repository

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
                Dim hash_bkhyguid As New Hashtable                              'bkguid格納用ハッシュテーブル
                Dim hash_bkhysoguid As New Hashtable                            'bkguid/soruleguid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Sorule_Model                  '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewnamebkhytoguid As String = PRE_VIEW_NAME & "物件部屋キー情報"
                Dim viewnamebkhytosoruleguid As String = PRE_VIEW_NAME & "物件部屋送金キー情報"

                '************************
                '作業準備
                '************************
                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytoguid)
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytosoruleguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "sorule"
                Dim fldnamegrp As String = "sorule_guid,sorule_no,relation_guid,kubunsyo,sorule_startymd," & _
                                           "sorule_endymd,kanri_keitaikbn,ikkatu_kanriflg,cyukai_kyflg,cyukai_koflg," & _
                                           "cyukai_kaiflg,soymd1_gaitoukbn,soymd1_simekbn,soymd1_sokintukikbn,soymd1_sokinsimekbn," & _
                                           "soymd2_gaitoukbn,soymd2_simekbn,soymd2_sokintukikbn,soymd2_sokinsimekbn,soymd3_gaitoukbn," & _
                                           "soymd3_simekbn,soymd3_sokintukikbn,soymd3_sokinsimekbn,soymd4_gaitoukbn,soymd4_simekbn," & _
                                           "soymd4_sokintukikbn,soymd4_sokinsimekbn,soymd5_gaitoukbn,soymd5_simekbn,soymd5_sokintukikbn," & _
                                           "soymd5_sokinsimekbn,sosaki_multikbn,sosaki_koteiflg,sosaki_koteisu,sosaki_anbunflg," & _
                                           "sosaki_hasuuketorisaki,sosaki_hasuadjustmentflg,ikkatu_kbn,ikkatu_bkgak,ikkatu_menseki," & _
                                           "ikkatu_mensekimonth,kanritesu_kbn,kanritesu_cyosyukbn,kanri_reigaiky,kanri_reigaikai," & _
                                           "hyteigaku_kbn,hyteigaku_kanrigak,hyteigaku_hiwariflg,bkteigaku_kanrigak,kanri_taxflg," & _
                                           "biko_basic,sh_daihyohyflg,sh_kozabetuflg,sh_kozabetuListNokbn,hyikkatu_nkin_no1kbn," & _
                                           "hyikkatu_nkin_no2kbn,hyikkatu_nkin_no3kbn,hyikkatu_nkin_no4kbn,hyikkatu_nkin_no5kbn,kanritesu_zeiumukbn," & _
                                           "history,rowid,no_soruleflg,cyukai_zuijisokin,ikkatu_calckbn," & _
                                           "ikkatu_taxflg,kanri_utizeiflg,soymd_basis"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytoguid, hash_bkhyguid)
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytosoruleguid, hash_bkhysoguid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, viewnamebkhytoguid, list_basekeydata)

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
                        .Vari_Sorule_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Relation_guid = tmp_keymain & "-" & tmp_keysub1

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
                                Case "送金ルール管理No"
                                    .Vari_Sorule_no = fldvalue.Trim
                                Case "一所有形態区分-棟/区分"
                                    .Vari_Kubunsyo = fldvalue.Trim
                                Case "送金ルール適用開始日"
                                    .Vari_Sorule_startymd = fldvalue.Trim
                                Case "送金ルール適用終了日"
                                    .Vari_Sorule_endymd = fldvalue.Trim
                                Case "管理形態"
                                    .Vari_Kanri_keitaikbn = fldvalue.Trim
                                Case "一括借上 一部管理"
                                    .Vari_Ikkatu_kanriflg = fldvalue.Trim
                                Case "仲介物件 - 新規契約業務"
                                    .Vari_Cyukai_kyflg = fldvalue.Trim
                                Case "仲介物件 - 契約更新業務"
                                    .Vari_Cyukai_koflg = fldvalue.Trim
                                Case "仲介物件 - 解約業務"
                                    .Vari_Cyukai_kaiflg = fldvalue.Trim
                                Case "送金日決定方法1 - 該当年月"
                                    .Vari_Soymd1_gaitoukbn = fldvalue.Trim
                                Case "送金日決定方法1 - 締日"
                                    .Vari_Soymd1_simekbn = fldvalue.Trim
                                Case "送金日決定方法1 - 送金月"
                                    .Vari_Soymd1_sokintukikbn = fldvalue.Trim
                                Case "送金日決定方法1 - 送金締日"
                                    .Vari_Soymd1_sokinsimekbn = fldvalue.Trim
                                Case "送金日決定方法2 - 該当年月"
                                    .Vari_Soymd2_gaitoukbn = fldvalue.Trim
                                Case "送金日決定方法2 - 締日"
                                    .Vari_Soymd2_simekbn = fldvalue.Trim
                                Case "送金日決定方法2 - 送金月"
                                    .Vari_Soymd2_sokintukikbn = fldvalue.Trim
                                Case "送金日決定方法2 - 送金締日"
                                    .Vari_Soymd2_sokinsimekbn = fldvalue.Trim
                                Case "送金日決定方法3 - 該当年月"
                                    .Vari_Soymd3_gaitoukbn = fldvalue.Trim
                                Case "送金日決定方法3 - 締日"
                                    .Vari_Soymd3_simekbn = fldvalue.Trim
                                Case "送金日決定方法3 - 送金月"
                                    .Vari_Soymd3_sokintukikbn = fldvalue.Trim
                                Case "送金日決定方法3 - 送金締日"
                                    .Vari_Soymd3_sokinsimekbn = fldvalue.Trim
                                Case "送金日決定方法4 - 該当年月"
                                    .Vari_Soymd4_gaitoukbn = fldvalue.Trim
                                Case "送金日決定方法4 - 締日"
                                    .Vari_Soymd4_simekbn = fldvalue.Trim
                                Case "送金日決定方法4 - 送金月"
                                    .Vari_Soymd4_sokintukikbn = fldvalue.Trim
                                Case "送金日決定方法4 - 送金締日"
                                    .Vari_Soymd4_sokinsimekbn = fldvalue.Trim
                                Case "送金日決定方法5 - 該当年月"
                                    .Vari_Soymd5_gaitoukbn = fldvalue.Trim
                                Case "送金日決定方法5 - 締日"
                                    .Vari_Soymd5_simekbn = fldvalue.Trim
                                Case "送金日決定方法5 - 送金月"
                                    .Vari_Soymd5_sokintukikbn = fldvalue.Trim
                                Case "送金日決定方法5 - 送金締日"
                                    .Vari_Soymd5_sokinsimekbn = fldvalue.Trim
                                Case "送金先単独/複数指定"
                                    .Vari_Sosaki_multikbn = fldvalue.Trim
                                Case "送金固定額使用フラグ"
                                    .Vari_Sosaki_koteiflg = fldvalue.Trim
                                Case "送金固定数"
                                    .Vari_Sosaki_koteisu = fldvalue.Trim
                                Case "均等案分フラグ"
                                    .Vari_Sosaki_anbunflg = fldvalue.Trim
                                Case "端数受取先"
                                    .Vari_Sosaki_hasuuketorisaki = fldvalue.Trim
                                Case "送金額案分端数調整フラグ"
                                    .Vari_Sosaki_hasuadjustmentflg = fldvalue.Trim
                                Case "一括借上 物件毎/部屋毎"
                                    .Vari_Ikkatu_kbn = fldvalue.Trim
                                Case "一括借上 物件毎設定額"
                                    .Vari_Ikkatu_bkgak = fldvalue.Trim
                                Case "一括借上 免責期間の設定フラグ"
                                    .Vari_Ikkatu_menseki = fldvalue.Trim
                                Case "一括借上 免責期間(解約翌月から何カ月)"
                                    .Vari_Ikkatu_mensekimonth = fldvalue.Trim
                                Case "管理手数料区分"
                                    .Vari_Kanritesu_kbn = fldvalue.Trim
                                Case "管理手数料徴収区分"
                                    .Vari_Kanritesu_cyosyukbn = fldvalue.Trim
                                Case "管理手数料 例外：契約金は対象外とする"
                                    .Vari_Kanri_reigaiky = fldvalue.Trim
                                Case "管理手数料 例外：解約金は対象外とする"
                                    .Vari_Kanri_reigaikai = fldvalue.Trim
                                Case "部屋毎定額 - 全部屋一律管理手数料"
                                    .Vari_Hyteigaku_kanrigak = fldvalue.Trim
                                Case "部屋毎定額 - 部屋毎日割りフラグ"
                                    .Vari_Hyteigaku_hiwariflg = fldvalue.Trim
                                Case "物件毎定額 - 物件管理手数料"
                                    .Vari_Bkteigaku_kanrigak = fldvalue.Trim
                                Case "管理手数料 消費税適用フラグ"
                                    .Vari_Kanri_taxflg = fldvalue.Trim
                                Case "備考 - 基本情報"
                                    .Vari_Biko_basic = fldvalue.Trim
                                Case "支払明細書関連 - 同時契約の場合は代表する部屋を表示"
                                    .Vari_Sh_daihyohyflg = fldvalue.Trim
                                Case "支払明細書関連 - 口座毎に支払明細書を作成1"
                                    .Vari_Sh_kozabetuflg = fldvalue.Trim
                                Case "支払明細書関連 - 口座毎に支払明細書を作成2"
                                    .Vari_Sh_kozabetuListNokbn = fldvalue.Trim
                                Case "部屋毎一括借上額 - 入金項目No1"
                                    .Vari_Hyikkatu_nkin_no1kbn = fldvalue.Trim
                                Case "部屋毎一括借上額 - 入金項目No2"
                                    .Vari_Hyikkatu_nkin_no2kbn = fldvalue.Trim
                                Case "部屋毎一括借上額 - 入金項目No3"
                                    .Vari_Hyikkatu_nkin_no3kbn = fldvalue.Trim
                                Case "部屋毎一括借上額 - 入金項目No4"
                                    .Vari_Hyikkatu_nkin_no4kbn = fldvalue.Trim
                                Case "部屋毎一括借上額 - 入金項目No5"
                                    .Vari_Hyikkatu_nkin_no5kbn = fldvalue.Trim
                                Case "管理手数料計算基準区分"
                                    .Vari_Kanritesu_zeiumukbn = fldvalue.Trim
                                Case "満額入金-送金フラグ"
                                    .Vari_No_soruleflg = fldvalue.Trim
                                Case "仲介物件 - 随時送金"
                                    .Vari_Cyukai_zuijisokin = fldvalue.Trim
                                Case "一括借上 部屋毎一括借上額の計算方法"
                                    .Vari_Ikkatu_calckbn = fldvalue.Trim
                                Case "一括借上 消費税適用フラグ"
                                    .Vari_Ikkatu_taxflg = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory
                        .Vari_Rowid = Guid.NewGuid.ToString
                        .Vari_Kanri_utizeiflg = 0
                        .Vari_Soymd_basis = 0

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
                            hash_cvitem("sorule_guid") = hash_bkhysoguid.Item(hash_cvitem.Item("sorule_guid"))
                            hash_cvitem("relation_guid") = hash_bkhyguid.Item(hash_cvitem.Item("relation_guid"))

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

                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		CONVERT(varchar,bk_no) + '-' + '' + '-' + CONVERT(varchar,sorule_no) AS キー "
                tmp_sql = tmp_sql & " 	FROM " & tblname
                tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON " & tblname & ".relation_guid = bkdata.bk_guid "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,sorule_no) "
                tmp_sql = tmp_sql & " 	FROM " & tblname
                tmp_sql = tmp_sql & " 	LEFT JOIN hydata ON " & tblname & ".relation_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                tmp_sql = tmp_sql & " ) AS VW "
                tmp_sql = tmp_sql & " WHERE キー IS NOT NULL"


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

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                Dim tmp_viewmainname As String = viewname.Replace(PRE_VIEW_NAME, "")

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                Select Case tmp_viewmainname
                    Case "物件部屋キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,bk_no)  + '-' AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,bk_guid "
                        tmp_sql_create = tmp_sql_create & " FROM bkdata "
                        tmp_sql_create = tmp_sql_create & " UNION "
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                    Case "物件部屋送金キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + '-' + CONVERT(varchar,BKS.kn_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,sorule_guid "
                        tmp_sql_create = tmp_sql_create & " FROM bkdata_syo AS BKS "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON BKS.bk_guid = BK.bk_guid "
                        tmp_sql_create = tmp_sql_create & " UNION "
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no)  + '-' + CONVERT(varchar,HYS.kn_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,sorule_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata_syo AS HYS "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON HYS.hy_guid = HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                End Select

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "送金ルール送金先情報"

    Public Class Sorule_sosaki_Repository

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
                Dim hash_bkhyguid As New Hashtable                                'bkguid格納用ハッシュテーブル
                Dim hash_bkhysoguid As New Hashtable                              'bkguid/soruleguid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Sorule_sosaki_Model           '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim viewnamebkhytoguid As String = PRE_VIEW_NAME & "物件部屋キー情報"
                Dim viewnamebkhytosoruleguid As String = PRE_VIEW_NAME & "物件部屋送金キー情報"

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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytoguid)
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytosoruleguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "sorule_sosaki"
                Dim fldnamegrp As String = "sorule_guid,sorule_no,sosaki_recno,so_ow_no,so_ow_kozano," & _
                                           "so_rit,so_fixsogak"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytoguid, hash_bkhyguid)
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytosoruleguid, hash_bkhysoguid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, viewnamebkhytosoruleguid, list_basekeydata)

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
                        .Vari_Sorule_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

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
                                Case "送金ルール管理No"
                                    .Vari_Sorule_no = fldvalue.Trim
                                Case "送金先行No"
                                    .Vari_Sosaki_recno = fldvalue.Trim
                                Case "送金先家主NO"
                                    .Vari_So_ow_no = fldvalue.Trim
                                Case "送金先口座NO"
                                    .Vari_So_ow_kozano = fldvalue.Trim
                                Case "送金率"
                                    .Vari_So_rit = fldvalue.Trim
                                Case "固定送金額"
                                    .Vari_So_fixsogak = fldvalue.Trim
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
                            hash_cvitem("sorule_guid") = hash_bkhysoguid.Item(hash_cvitem.Item("sorule_guid"))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            Else    '20161014 重複エラーのログ出力処理を追加 -add
                                Dim log_key As String = "sorule_sosaki-sorule_guid" & "/" & "sorule_sosaki-sorule_no" & "/" & "sorule_sosaki-sosaki_recno"
                                Dim errstr As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
                                hash_log.Clear()
                                hash_log.Add(log_key, errstr)
                                tmp_hash("sorule_guid") = ""
                                hash_cvitem("sorule_guid") = ""
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

                '20161108_2 送金ルール送金先情報の移行制御処理修正 -add sta
                '移行したデータのうち、管理形態が自社物件、送金保留以外の場合は送金先と送金先Noが必須なのでデータ調整を行う
                '※ここで簡易的な例外処理を入れておく(20161108 例外処理追加)
                If CNVNO = ConvertTypes._汎用 Then
                    Try
                        Call Me.Chk_SoruleSosaki_Sono(sqlcnnv10, tblname)
                        Call Me.Chk_SoruleSosaki_SoKozano(sqlcnnv10, tblname)
                        Call Me.Chk_SoruleBase(sqlcnnv10, tblname)
                        Call Me.Set_SoruleSosakiDummyRecord(sqlcnnv10)
                    Catch ex As Exception
                        '----- ログ出力 -----
                        Dim tmptmpcnt As Integer = 0
                        Dim tmptmpstr As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(41, sheetname, "データ調整", "データ調整中にエラーが発生しました。"), False)
                        DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, tmptmpcnt)
                    End Try
                End If
                '20161108_2 送金ルール送金先情報の移行制御処理修正 -add end

                '20161108_2 送金ルール送金先情報の移行制御処理修正 -del sta
                '↑に統合
                ''20161011 送金ルール送金先有無による送金ルール基本情報の移行制御 -add sta
                ''送金ルール基本情報と送金ルール送金先情報を照合し移行制御を行う
                'If CNVNO = ConvertTypes._汎用 Then
                '    Call Me.Chk_SoruleBase(sqlcnnv10, tblname)
                'End If
                ''20161011 送金ルール送金先有無による送金ルール基本情報の移行制御 -add end
                '20161108_2 送金ルール送金先情報の移行制御処理修正 -del end
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

                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + '' + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON sorule.relation_guid = bkdata.bk_guid "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                tmp_sql = tmp_sql & " 	LEFT JOIN hydata ON sorule.relation_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                tmp_sql = tmp_sql & " ) AS VW "
                tmp_sql = tmp_sql & " WHERE キー IS NOT NULL "


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

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                Dim tmp_viewmainname As String = viewname.Replace(PRE_VIEW_NAME, "")

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                Select Case tmp_viewmainname
                    Case "物件部屋キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,bk_no)  + '-' AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,bk_guid "
                        tmp_sql_create = tmp_sql_create & " FROM bkdata "
                        tmp_sql_create = tmp_sql_create & " UNION "
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                    Case "物件部屋送金キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + '-' + CONVERT(varchar,BKS.kn_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,sorule_guid "
                        tmp_sql_create = tmp_sql_create & " FROM bkdata_syo AS BKS "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON BKS.bk_guid = BK.bk_guid "
                        tmp_sql_create = tmp_sql_create & " UNION "
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no)  + '-' + CONVERT(varchar,HYS.kn_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,sorule_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata_syo AS HYS "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON HYS.hy_guid = HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                End Select

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 送金先情報で送金先が存在しないデータを抽出、削除
            ''' </summary>
            ''' <param name="sqlcnnv10">DB接続用オブジェクト</param>
            ''' <param name="tblname">送金先情報のテーブル名</param>
            ''' <remarks>20161108_2 送金ルール送金先情報の移行制御処理修正 新規追加</remarks>
            Public Sub Chk_SoruleSosaki_Sono(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String)

                Dim list_err As New List(Of String)

                '----------------
                '不正データ抽出
                '----------------
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) AS [ログ対象データ] "
                tmp_sql = tmp_sql & " 		FROM sorule_sosaki AS SOS "
                tmp_sql = tmp_sql & " 		LEFT JOIN sorule AS SO ON SOS.sorule_guid = SO.sorule_guid "
                tmp_sql = tmp_sql & " 		LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " 		WHERE SOS.so_ow_no IS NULL AND SO.kanri_keitaikbn NOT IN (4,5) "
                tmp_sql = tmp_sql & " 	) AS ITTO "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) + '、' + '部屋No = ' + CONVERT(VARCHAR(MAX),hy_no) AS [ログ対象データ] "
                tmp_sql = tmp_sql & " 		FROM sorule_sosaki AS SOS "
                tmp_sql = tmp_sql & " 		LEFT JOIN sorule AS SO ON SOS.sorule_guid = SO.sorule_guid "
                tmp_sql = tmp_sql & " 		LEFT JOIN hydata AS HY ON SO.relation_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " 		LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " 		WHERE SOS.so_ow_no IS NULL AND SO.kanri_keitaikbn NOT IN (4,5) "
                tmp_sql = tmp_sql & " 	) AS KBN "
                tmp_sql = tmp_sql & " ) AS VW "
                tmp_sql = tmp_sql & " WHERE [ログ対象データ] IS NOT NULL "
                DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_err)

                '----------------------------------------------
                '不正データが存在する場合、ログ出力→削除処理
                '----------------------------------------------
                If list_err.Count <> 0 Then

                    '作業用変数
                    Dim tmp_hash As New Hashtable       'ログ作成用の仮オブジェクト(引数として使用するだけ)
                    Dim tmp_logcnt As Integer = 0       'ログ作成用の仮変数(引数として使用するだけ)
                    Dim tmp_cnt As Integer = 0          '仮変数(引数として使用するだけ)
                    Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                    Dim tmp_loghash As New Hashtable
                    Dim log_key As String = tblname & "-" & "so_ow_no"
                    Dim tmp_hubi As String = "物件管理情報の送金先が設定されていないため移行できません。"
                    Dim errstr As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & tmp_hubi & "-" & LOG_TAISYO_DEFAULT
                    tmp_loghash.Add(log_key, errstr)

                    'ログ出力
                    For Each str_logkey In list_err

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(tmp_loghash, tmp_hash, tmp_hash, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        '挿入
                        For Each logvalue In sortlist_log
                            Dim tmp_sql_insert As String = logvalue.Value
                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                        Next
                        '初期化
                        tmp_logcnt = 0
                        sortlist_log.Clear()

                    Next

                    '不正データ削除
                    Dim tmp_sql_delete As String = ""
                    tmp_sql_delete = tmp_sql_delete & " DELETE FROM sorule_sosaki "
                    tmp_sql_delete = tmp_sql_delete & " WHERE so_ow_no IS NULL "
                    tmp_sql_delete = tmp_sql_delete & " AND   sorule_guid IN "
                    tmp_sql_delete = tmp_sql_delete & " (SELECT sorule_guid FROM sorule WHERE kanri_keitaikbn NOT IN (4,5)) "
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_delete, tmp_cnt)

                End If

            End Sub

            ''' <summary>
            ''' 送金先情報で送金先口座が存在しないデータを抽出、削除
            ''' 送金先が存在しないデータは調整済み
            ''' </summary>
            ''' <param name="sqlcnnv10">DB接続用オブジェクト</param>
            ''' <param name="tblname">送金先情報のテーブル名</param>
            ''' <remarks>20161108_2 送金ルール送金先情報の移行制御処理修正 新規追加</remarks>
            Public Sub Chk_SoruleSosaki_SoKozano(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String)

                Dim list_err As New List(Of String)

                '----------------
                '不正データ抽出
                '----------------
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) AS [ログ対象データ] "
                tmp_sql = tmp_sql & " 		FROM sorule_sosaki AS SOS "
                tmp_sql = tmp_sql & " 		LEFT JOIN sorule AS SO ON SOS.sorule_guid = SO.sorule_guid "
                tmp_sql = tmp_sql & " 		LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " 		WHERE SOS.so_ow_no IS NOT NULL AND SOS.so_ow_kozano IS NULL AND SO.kanri_keitaikbn NOT IN (4,5) "
                tmp_sql = tmp_sql & " 	) AS ITTO "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) + '、' + '部屋No = ' + CONVERT(VARCHAR(MAX),hy_no) AS [ログ対象データ] "
                tmp_sql = tmp_sql & " 		FROM sorule_sosaki AS SOS "
                tmp_sql = tmp_sql & " 		LEFT JOIN sorule AS SO ON SOS.sorule_guid = SO.sorule_guid "
                tmp_sql = tmp_sql & " 		LEFT JOIN hydata AS HY ON SO.relation_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " 		LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " 		WHERE SOS.so_ow_no IS NOT NULL AND SOS.so_ow_kozano IS NULL AND SO.kanri_keitaikbn NOT IN (4,5) "
                tmp_sql = tmp_sql & " 	) AS KBN "
                tmp_sql = tmp_sql & " ) AS VW "
                tmp_sql = tmp_sql & " WHERE [ログ対象データ] IS NOT NULL "
                DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_err)

                '----------------------------------------------
                '不正データが存在する場合、ログ出力→削除処理
                '----------------------------------------------
                If list_err.Count <> 0 Then

                    '作業用変数
                    Dim tmp_hash As New Hashtable       'ログ作成用の仮オブジェクト(引数として使用するだけ)
                    Dim tmp_logcnt As Integer = 0       'ログ作成用の仮変数(引数として使用するだけ)
                    Dim tmp_cnt As Integer = 0          '仮変数(引数として使用するだけ)
                    Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                    Dim tmp_loghash As New Hashtable
                    Dim log_key As String = tblname & "-" & "so_ow_kozano"
                    Dim tmp_hubi As String = "物件管理情報の送金先口座が設定されていないため移行できません。"
                    Dim errstr As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & tmp_hubi & "-" & LOG_TAISYO_DEFAULT
                    tmp_loghash.Add(log_key, errstr)

                    'ログ出力
                    For Each str_logkey In list_err

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(tmp_loghash, tmp_hash, tmp_hash, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        '挿入
                        For Each logvalue In sortlist_log
                            Dim tmp_sql_insert As String = logvalue.Value
                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                        Next
                        '初期化
                        tmp_logcnt = 0
                        sortlist_log.Clear()

                    Next

                    '不正データ削除
                    Dim tmp_sql_delete As String = ""
                    tmp_sql_delete = tmp_sql_delete & " DELETE FROM sorule_sosaki "
                    tmp_sql_delete = tmp_sql_delete & " WHERE so_ow_no IS NOT NULL AND so_ow_kozano IS NULL "
                    tmp_sql_delete = tmp_sql_delete & " AND   sorule_guid IN "
                    tmp_sql_delete = tmp_sql_delete & " (SELECT sorule_guid FROM sorule WHERE kanri_keitaikbn NOT IN (4,5)) "
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_delete, tmp_cnt)

                End If

            End Sub

            ''' <summary>
            ''' 送金ルール基本情報が存在するが送金先が存在しないデータのチェック '20161011 送金ルール送金先有無による送金ルール基本情報の移行制御 -add
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <remarks></remarks>
            Public Sub Chk_SoruleBase(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String)

                Dim list_err As New List(Of String)

                '----------------
                '不正データ抽出
                '----------------

                '送金ルール基本情報が存在するが送金先が存在しないデータを抽出(ログ出力用に成形済み)
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) AS [ログ対象データ] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		/*20161107 自社物件の送金ルール作成方法修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*SELECT * FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kubunsyo = 1*/ "
                tmp_sql = tmp_sql & " 		SELECT * FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kubunsyo = 1 AND kanri_keitaikbn <> 4 "
                tmp_sql = tmp_sql & " 		/*20161107 自社物件の送金ルール作成方法修正 chg end*/ "
                tmp_sql = tmp_sql & " 	) AS ITTO "
                tmp_sql = tmp_sql & " 	LEFT JOIN bkdata AS BK "
                tmp_sql = tmp_sql & " 	ON ITTO.relation_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) + '、' + '部屋No = ' + CONVERT(VARCHAR(MAX),hy_no) AS [ログ対象データ] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		/*20161107 自社物件の送金ルール作成方法修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*SELECT * FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kubunsyo = 2*/ "
                tmp_sql = tmp_sql & " 		SELECT * FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kubunsyo = 2 AND kanri_keitaikbn <> 4 "
                tmp_sql = tmp_sql & " 		/*20161107 自社物件の送金ルール作成方法修正 chg end*/ "
                tmp_sql = tmp_sql & " 	) AS KBN "
                tmp_sql = tmp_sql & " 	LEFT JOIN hydata AS HY ON KBN.relation_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " 	LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " ) AS VW "
                DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_err)

                '----------------------------------------------
                '不正データが存在する場合、ログ出力→削除処理
                '----------------------------------------------
                If list_err.Count <> 0 Then

                    '作業用変数
                    Dim tmp_hash As New Hashtable       'ログ作成用の仮オブジェクト(引数として使用するだけ)
                    Dim tmp_logcnt As Integer = 0       'ログ作成用の仮変数(引数として使用するだけ)
                    Dim tmp_cnt As Integer = 0          '仮変数(引数として使用するだけ)
                    Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                    Dim tmp_loghash As New Hashtable
                    Dim log_key As String = tblname & "-" & "so_ow_no"
                    Dim tmp_hubi As String = "物件管理情報の送金先が設定されていないため移行できません。"
                    Dim errstr As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & tmp_hubi & "-" & LOG_TAISYO_DEFAULT
                    tmp_loghash.Add(log_key, errstr)

                    'ログ出力
                    For Each str_logkey In list_err

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(tmp_loghash, tmp_hash, tmp_hash, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        '挿入
                        For Each logvalue In sortlist_log
                            Dim tmp_sql_insert As String = logvalue.Value
                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                        Next
                        '初期化
                        tmp_logcnt = 0
                        sortlist_log.Clear()

                    Next

                    '不正データ削除
                    '20161107 自社物件の送金ルール作成方法修正 -chg sta
                    'Dim tmp_sql_ittodelete As String = " DELETE FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) "
                    Dim tmp_sql_ittodelete As String = " DELETE FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kanri_keitaikbn <> 4 "
                    '20161107 自社物件の送金ルール作成方法修正 -chg end
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_ittodelete, tmp_cnt)

                End If

            End Sub

            ''' <summary>
            ''' 送金先情報のダミーレコード作成処理
            ''' 画面上で送金ルールを作成した際に送金先情報が自動で3レコード作成されるため、それに合わせる
            ''' </summary>
            ''' <param name="sqlcnnv10">DB接続用オブジェクト</param>
            ''' <remarks>20161108_2 送金ルール送金先情報のダミーレコード作成処理の追加 新規追加</remarks>
            Public Sub Set_SoruleSosakiDummyRecord(ByVal sqlcnnv10 As SqlConnection)

                'ダミーレコード作成用クエリ
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " INSERT INTO sorule_sosaki "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 			DUMMYREC.* "
                tmp_sql = tmp_sql & " 		,NULL AS so_ow_no "
                tmp_sql = tmp_sql & " 		,NULL AS so_ow_kozano "
                tmp_sql = tmp_sql & " 		,NULL AS so_rit "
                tmp_sql = tmp_sql & " 		,NULL AS so_fixsogak "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT DISTINCT "
                tmp_sql = tmp_sql & " 				sorule_guid "
                tmp_sql = tmp_sql & " 			,1 AS sorule_no "
                tmp_sql = tmp_sql & " 			,1 AS sosaki_recno "
                tmp_sql = tmp_sql & " 		FROM sorule_sosaki "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT DISTINCT "
                tmp_sql = tmp_sql & " 				sorule_guid "
                tmp_sql = tmp_sql & " 			,1 AS sorule_no "
                tmp_sql = tmp_sql & " 			,2 AS sosaki_recno "
                tmp_sql = tmp_sql & " 		FROM sorule_sosaki "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT DISTINCT "
                tmp_sql = tmp_sql & " 				sorule_guid "
                tmp_sql = tmp_sql & " 			,1 AS sorule_no "
                tmp_sql = tmp_sql & " 			,3 AS sosaki_recno "
                tmp_sql = tmp_sql & " 		FROM sorule_sosaki "
                tmp_sql = tmp_sql & " 	) AS DUMMYREC "
                tmp_sql = tmp_sql & " 	WHERE NOT EXISTS "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT * FROM sorule_sosaki AS SOS "
                tmp_sql = tmp_sql & " 			WHERE DUMMYREC.sorule_guid = SOS.sorule_guid "
                tmp_sql = tmp_sql & " 			AND   DUMMYREC.sorule_no = SOS.sorule_no "
                tmp_sql = tmp_sql & " 			AND   DUMMYREC.sosaki_recno = SOS.sosaki_recno "
                tmp_sql = tmp_sql & " 		) "

                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

        End Class

    End Class

#End Region

#Region "送金ルール入金項目情報"

    Public Class Sorule_nk_cmrule_Repository

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
                Dim hash_bkhyguid As New Hashtable                              'bkguid格納用ハッシュテーブル
                Dim hash_bkhysoguid As New Hashtable                            'bkguid/soruleguid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Sorule_nk_cmrule_Model        '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列                    
                Dim viewnamebkhytoguid As String = PRE_VIEW_NAME & "物件部屋キー情報"
                Dim viewnamebkhytosoruleguid As String = PRE_VIEW_NAME & "物件部屋送金キー情報"

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Nkinkomk()
                
                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytoguid)
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytosoruleguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "sorule_nk_cmrule"
                Dim fldnamegrp As String = "sorule_guid,sorule_no,taisyokbn,nkin_sortorder,nkin_no," & _
                                           "sokin_rit,kanrigak_rit,hosyo_flg,so_no"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytoguid, hash_bkhyguid)
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytosoruleguid, hash_bkhysoguid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, viewnamebkhytosoruleguid, list_basekeydata)

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
                    Dim fldname_keysub4 As String = readtbl.Columns(keycol_sub4 - 1).ColumnName.Trim

                    '---------------
                    'データ部処理
                    '---------------
                    For cntii = 0 To rowcnt - 1

                        '中断処理
                        Application.DoEvents()
                        If CancelFlg Then
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
                        Dim tmp_keysub4 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub4 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4
                        .Vari_Sorule_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4

                        '作業用変数
                        Dim tmp_nkinname As String = ""

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "送金ルール管理No"
                                    .Vari_Sorule_no = fldvalue
                                Case "月々/契約時/更新時区分"
                                    .Vari_Taisyokbn = fldvalue
                                Case "行NO"
                                    .Vari_Nkin_sortorder = fldvalue
                                Case "入金項目名"
                                    tmp_nkinname = fldvalue
                                Case "送金率"
                                    .Vari_Sokin_rit = fldvalue
                                Case "管理手数料率"
                                    .Vari_Kanrigak_rit = fldvalue
                                Case "滞納保証有無"
                                    .Vari_Hosyo_flg = fldvalue
                            End Select

                        Next

                        '紐付用に成形
                        .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname(.Vari_Taisyokbn)

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
                            hash_cvitem("sorule_guid") = hash_bkhysoguid.Item(hash_cvitem.Item("sorule_guid"))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            Else    '20161014 重複エラーのログ出力処理を追加 -add
                                Dim log_key As String = "sorule_nk_cmrule-sorule_guid" & "/" & "sorule_nk_cmrule-sorule_no" & "/" & _
                                                        "sorule_nk_cmrule-taisyokbn" & "/" & "sorule_nk_cmrule-nkin_sortorder" & "/" & "sorule_nk_cmrule-nkin_no"
                                Dim errstr As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
                                hash_log.Clear()
                                hash_log.Add(log_key, errstr)
                                tmp_hash("sorule_guid") = ""
                                hash_cvitem("sorule_guid") = ""
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

                'その他請求レコード一括挿入処理
                Dim tmpcnt As Integer = 0
                Dim othernkkomksql As String = Me.Get_UseQry_Insert()
                DBExec.Exec_NonQuery(sqlcnnv10, othernkkomksql, tmpcnt)

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

                'tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + '' + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                'tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                'tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON sorule.relation_guid = bkdata.bk_guid "
                'tmp_sql = tmp_sql & " 	UNION "
                'tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                'tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                'tmp_sql = tmp_sql & " 	LEFT JOIN hydata ON sorule.relation_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                'tmp_sql = tmp_sql & " ) AS VW "
                'tmp_sql = tmp_sql & " WHERE キー IS NOT NULL "


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

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                Dim tmp_viewmainname As String = viewname.Replace(PRE_VIEW_NAME, "")

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                Select Case tmp_viewmainname
                    Case "物件部屋キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,bk_no)  + '-' AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,bk_guid "
                        tmp_sql_create = tmp_sql_create & " FROM bkdata "
                        tmp_sql_create = tmp_sql_create & " UNION "
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                    Case "物件部屋送金キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + '-' + CONVERT(varchar,BKS.kn_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,sorule_guid "
                        tmp_sql_create = tmp_sql_create & " FROM bkdata_syo AS BKS "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON BKS.bk_guid = BK.bk_guid "
                        tmp_sql_create = tmp_sql_create & " UNION "
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no)  + '-' + CONVERT(varchar,HYS.kn_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,sorule_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata_syo AS HYS "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON HYS.hy_guid = HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                End Select

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 送金ルールのその他請求レコード挿入クエリ '20160928 送金ルールその他請求取得処理修正
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Insert() As String

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " INSERT INTO sorule_nk_cmrule "
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 sorule_guid "
                tmp_sql = tmp_sql & " 	,sorule_no "
                tmp_sql = tmp_sql & " 	,7 AS taisyokbn "
                tmp_sql = tmp_sql & " 	,0 AS nkin_sortorder "
                tmp_sql = tmp_sql & " 	,0 AS nkin_no "
                tmp_sql = tmp_sql & " 	,0 AS sokin_rit "
                tmp_sql = tmp_sql & " 	,0 AS kanrigak_rit "
                tmp_sql = tmp_sql & " 	,hosyo_flg "
                tmp_sql = tmp_sql & " 	,NULL AS so_no "
                tmp_sql = tmp_sql & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT DISTINCT sorule_guid,sorule_no,hosyo_flg FROM sorule_nk_cmrule "
                tmp_sql = tmp_sql & " ) AS VW "
                Return tmp_sql


            End Function

            '2016.04.06 入金項目読込処理の修正 -del sta
            ' ''' <summary>
            ' ''' 入金項目紐付情報取得
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

#Region "送金ルール控除項目情報"

    Public Class Sorule_kojo_cmrule_Repository

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
                Dim hash_bkhyguid As New Hashtable                                'bkguid格納用ハッシュテーブル
                Dim hash_bkhysoguid As New Hashtable                              'bkguid/soruleguid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Sorule_kojo_cmrule_Model      '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim viewnamebkhytoguid As String = PRE_VIEW_NAME & "物件部屋キー情報"
                Dim viewnamebkhytosoruleguid As String = PRE_VIEW_NAME & "物件部屋送金キー情報"

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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytoguid)
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytosoruleguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "sorule_kojo_cmrule"
                Dim fldnamegrp As String = "sorule_guid,sorule_no,kojo_taisyokbn,nkin_sortorder,taisyonkin_no," & _
                                           "sosai_flg,kojocalc_kbn,kojo_gak,kojo_gakzeikbn,kojonkin_no," & _
                                           "kojo_rit,kojo_ritzeikbn,kojo_ritutizei,zei_rit,tateazu_flg," & _
                                           "kojo_zeigak,sotaisyo_flg,sotaisyonkin_no"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytoguid, hash_bkhyguid)
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytosoruleguid, hash_bkhysoguid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, viewnamebkhytosoruleguid, list_basekeydata)

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
                    Dim fldname_keysub4 As String = readtbl.Columns(keycol_sub4 - 1).ColumnName.Trim

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
                        Dim tmp_keysub4 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub4 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4
                        .Vari_Sorule_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4


                        '2016.04.06 入金項目読込処理の修正 -add
                        '作業用変数
                        Dim tmp_nkinname As String = ""

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "送金ルール管理No"
                                    .Vari_Sorule_no = fldvalue.Trim
                                Case "契約時/更新時区分"
                                    .Vari_Kojo_taisyokbn = fldvalue.Trim
                                Case "行NO"
                                    .Vari_Nkin_sortorder = fldvalue.Trim
                                Case "控除入金項目名"
                                    '2016.04.06 入金項目読込処理の修正 -chg sta
                                    '.Vari_Taisyonkin_no = fldvalue.Trim
                                    tmp_nkinname = fldvalue
                                    '2016.04.06 入金項目読込処理の修正 -chg end
                                Case "相殺予定フラグ"
                                    .Vari_Sosai_flg = fldvalue.Trim
                                Case "控除額基準"
                                    .Vari_Kojocalc_kbn = fldvalue.Trim
                                Case "控除額"
                                    .Vari_Kojo_gak = fldvalue.Trim
                                Case "控除額税区分"
                                    .Vari_Kojo_gakzeikbn = fldvalue.Trim
                                Case "控除対象入金項目名"
                                    .Vari_Kojonkin_no = fldvalue.Trim
                                Case "控除率"
                                    .Vari_Kojo_rit = fldvalue.Trim
                                Case "控除率税区分"
                                    .Vari_Kojo_ritzeikbn = fldvalue.Trim
                                Case "控除率税有無"
                                    .Vari_Kojo_ritutizei = fldvalue.Trim
                                Case "適用税率"
                                    .Vari_Zei_rit = fldvalue.Trim
                                Case "立替回収または預り金"
                                    .Vari_Tateazu_flg = fldvalue.Trim
                                Case "控除税額"
                                    .Vari_Kojo_zeigak = fldvalue.Trim
                                Case "送金対象フラグ"
                                    .Vari_Sotaisyo_flg = fldvalue.Trim
                                Case "送金対象入金項目名"
                                    .Vari_Sotaisyonkin_no = fldvalue.Trim
                            End Select

                        Next

                        '2016.04.06 入金項目読込処理の修正 -add sta
                        '紐付用に成形
                        .Vari_Taisyonkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname("9")
                        '2016.04.06 入金項目読込処理の修正 -add end

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
                            hash_cvitem("sorule_guid") = hash_bkhysoguid.Item(hash_cvitem.Item("sorule_guid"))

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

                'tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + '' + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                'tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                'tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON sorule.relation_guid = bkdata.bk_guid "
                'tmp_sql = tmp_sql & " 	UNION "
                'tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                'tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                'tmp_sql = tmp_sql & " 	LEFT JOIN hydata ON sorule.relation_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                'tmp_sql = tmp_sql & " ) AS VW "
                'tmp_sql = tmp_sql & " WHERE キー IS NOT NULL "


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

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)

            End Sub

            ''' <summary>
            ''' キー/guidをセットにしたVIEWの作成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <remarks></remarks>
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

                Dim tmp_viewmainname As String = viewname.Replace(PRE_VIEW_NAME, "")

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                Select Case tmp_viewmainname
                    Case "物件部屋キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,bk_no)  + '-' AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,bk_guid "
                        tmp_sql_create = tmp_sql_create & " FROM bkdata "
                        tmp_sql_create = tmp_sql_create & " UNION "
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                    Case "物件部屋送金キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + '-' + CONVERT(varchar,BKS.kn_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,sorule_guid "
                        tmp_sql_create = tmp_sql_create & " FROM bkdata_syo AS BKS "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON BKS.bk_guid = BK.bk_guid "
                        tmp_sql_create = tmp_sql_create & " UNION "
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no)  + '-' + CONVERT(varchar,HYS.kn_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,sorule_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata_syo AS HYS "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON HYS.hy_guid = HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                End Select

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            '2016.04.06 入金項目読込処理の修正 -del sta
            ' ''' <summary>
            ' ''' 入金項目紐付情報取得
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

End Namespace


