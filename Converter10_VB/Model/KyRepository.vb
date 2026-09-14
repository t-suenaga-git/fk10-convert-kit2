Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "契約基本情報"

    Public Class Kydata_Repository

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
                Dim hash_bkguid As New Hashtable                                'bkguid格納用ハッシュテーブル
                Dim hash_hyguid As New Hashtable                                'bkguid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_Model                  '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim viewname As String = PRE_VIEW_NAME & "物件部屋キー情報"

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
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewname)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata"
                Dim fldnamegrp As String = "ky_guid,bk_guid,hy_guid,ky_no,ky_deleteflg," & _
                                           "delete_guid,delete_day,delete_cnt,syokai_kyymd,status," & _
                                           "cancelriyu,status_ymd,cyukai_gy_fudono,syunin_logonuser_no,tetuke_gak1," & _
                                           "tetuke_ymd1,tetuke_biko1,kaiyaku_flg,history,rowid," & _
                                           "movefrom_kyguid,moveto_kyguid,svbunrui_no,krbunrui_no,kaiyaku_uketukekbn," & _
                                           "kaiyaku_months,kaiyaku_days,kaiyaku_day,cyukai_tantoname,cyukai_tantonamesjis"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewname, hash_bkguid, 0)  'bk_guid取得
                Call Me.Get_Guid(sqlcnnv10, viewname, hash_hyguid, 1)  'hy_guid取得

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewname, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)
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
                        Dim tmp_kystartymd As String = ""
                        Dim tmp_kyendymd As String = ""

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Bk_guid = tmp_keymain & "-" & tmp_keysub1
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
                                Case "契約No"
                                    .Vari_Ky_no = fldvalue
                                Case "初回契約日"
                                    .Vari_Syokai_kyymd = fldvalue
                                Case "契約状況(ステータス)"
                                    .Vari_Status = fldvalue
                                Case "キャンセル理由"
                                    .Vari_Cancelriyu = fldvalue
                                Case "ステータス変更日"
                                    .Vari_Status_ymd = fldvalue
                                Case "仲介業者No"
                                    .Vari_Cyukai_gy_fudono = fldvalue
                                Case "取引主任者No"
                                    .Vari_Syunin_logonuser_no = fldvalue
                                Case "手付預り額①"
                                    .Vari_Tetuke_gak1 = fldvalue
                                Case "手付預り日①"
                                    .Vari_Tetuke_ymd1 = fldvalue
                                Case "手付預り備考①"
                                    .Vari_Tetuke_biko1 = fldvalue
                                Case "解約フラグ"
                                    .Vari_Kaiyaku_flg = fldvalue
                                Case "サービス分類"
                                    .Vari_Svbunrui_no = fldvalue
                                Case "会計グループ分類"
                                    .Vari_Krbunrui_no = fldvalue
                                Case "解約受付区分"
                                    .Vari_Kaiyaku_uketukekbn = fldvalue
                                Case "解約受付月数"
                                    .Vari_Kaiyaku_months = fldvalue
                                Case "解約受付日数"
                                    .Vari_Kaiyaku_days = fldvalue
                                Case "解約受付日にち"
                                    .Vari_Kaiyaku_day = fldvalue
                                Case "仲介担当者名"
                                    .Vari_Cyukai_tantoname = fldvalue
                                Case "仲介担当者名SJIS"
                                    .Vari_Cyukai_tantonamesjis = fldvalue
                                    '2016.04.06 仮契約情報の移行制御処理を追加 -add sta
                                Case "契約開始日"
                                    tmp_kystartymd = fldvalue
                                Case "契約終了日"
                                    tmp_kyendymd = fldvalue
                                    '2016.04.06 仮契約情報の移行制御処理を追加 -add end
                            End Select

                        Next

                        '固定値
                        .Vari_Ky_guid = Guid.NewGuid.ToString
                        .Vari_Ky_deleteflg = 0
                        .Vari_Delete_guid = ""
                        .Vari_Delete_day = ""
                        .Vari_Delete_cnt = ""
                        .Vari_Movefrom_kyguid = ""
                        .Vari_Moveto_kyguid = ""
                        .Vari_History = DefHistory
                        .Vari_Rowid = Guid.NewGuid.ToString

                        '20170530 契約状況確定日項目追加対応 -add sta
                        '契約状況とステータス変更日を見て値の有無を確認・設定する
                        If .Vari_Status = "2" OrElse .Vari_Status = "3" Then
                            If .Vari_Status_ymd = "" Then
                                '空の場合は契約開始日を設定する
                                .Vari_Status_ymd = tmp_kystartymd
                            End If
                        End If
                        '20170530 契約状況確定日項目追加対応 -add end

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '2016.04.06 仮契約情報の移行制御処理を追加 -add sta
                        '*****************************************************************************************
                        '※制御自体は可能だが、ログの出力に個別に対応が必要
                        '※契約基本情報のみにチェックを入れた場合はテーブル名とフィールド名が出力されない
                        '  (契約履歴情報のテーブル名、フィールド名をログに出力するため)
                        '※ただし、革命上では契約基本情報と契約履歴情報は同時に作成される。そのため
                        '  ・どちらかのみの移行を不可にする等の制御を設ける (必ずセットで移行する)
                        '  ・移行項目の統合で自動で両方移行するようにする (統合する予定なのでおそらくこれになる)
                        '*****************************************************************************************
                        If tmp_kystartymd = "" Then
                            hash_log.Add("kydata_kihon-kystart_ymd", LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED)
                            skipflg = True
                        ElseIf tmp_kyendymd = "" Then
                            hash_log.Add("kydata_kihon-kyend_ymd", LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED)
                            skipflg = True
                        End If
                        '2016.04.06 仮契約情報の移行制御処理を追加 -add end

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("bk_guid") = hash_bkguid.Item(hash_cvitem.Item("bk_guid"))
                            ''hash_cvitem("hy_guid") = hash_hyguid.Item(hash_cvitem.Item("hy_guid"))
                            'hash_cvitem("bk_guid") = hash_bkguid.Item(StrConv(hash_cvitem.Item("bk_guid"), VbStrConv.Narrow))
                            'hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("bk_guid") = hash_bkguid.Item(hash_cvitem.Item("bk_guid"))
                            hash_cvitem("hy_guid") = hash_hyguid.Item(hash_cvitem.Item("hy_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                ''クローズ処理
                'excelfile.ExcelFile_ReadClose(con_read)
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                '20160929 汎用CV時のステータス変更日の一括更新処理 -add sta
                '仮契約の場合はステータス変更日NULLに一括更新
                If CNVNO = ConvertTypes._汎用 Then
                    Dim tmpcnt As Integer = 0
                    Dim stymdupdatesql As String = " UPDATE kydata SET status_ymd = NULL WHERE [status] = 1 "
                    DBExec.Exec_NonQuery(sqlcnnv10, stymdupdatesql, tmpcnt)
                End If
                '20160929 汎用CV時のステータス変更日の一括更新処理 -add end
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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
            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable, ByVal bkhytype As Integer)

                Dim fldname As String = ""

                Select Case bkhytype
                    Case 0
                        fldname = " キー,bk_guid "
                    Case 1
                        fldname = " キー,hy_guid "
                End Select

                Dim tmp_sql As String = " SELECT " & fldname & " FROM " & tblname

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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,HY.bk_guid "
                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約履歴情報"

    Public Class Kydata_kihon_Repository

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
                Dim hash_bkhykyguid As New Hashtable                              'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_kihon_Model       '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '2016.04.06 紐付データ取得処理を外出し -chg sta
                'Call Me.Get_RelData_NkinKbn()
                Call SetRelItemToObject.Set_RelData_Nkinkbn()
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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_kihon"
                '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add (torihiki_tesumoto,torihiki_tesukyaku,torihiki_tesugak,torihiki_tesuzeikbn,torihiki_tesuzeigak)を追加
                '20160519 EXEUpdateに伴う修正 契約履歴情報 -chg sta
                'Dim fldnamegrp As String = "ky_guid,ky_recno,ko_no,henko_no,sqdata_guid," & _
                '                           "ky_bango,ky_ymd,kystart_ymd,kyend_ymd,henko_ymd," & _
                '                           "tuti_ymd,print_ymd,kanri_gy_fudono,gy_hosyono,hosyo_naiyo," & _
                '                           "kokyaku_bango,kyrui_no,siyo_mokuteki,kosin_umu,yatin_kbn," & _
                '                           "fkae_startym,fkae_willstartflg,yatin_kozakbn,yatin_kozano,maitukiyatin_kozano," & _
                '                           "yokugetu_uketoriflg,biko,tougetu_sagakuflg,nextky_startymd,nextky_endymd," & _
                '                           "kosin_hiwariflg,sokojorule_kbn,soyotei_ymdflg,soyotei_ymd,kanritesu_flg," & _
                '                           "kanritesu_gak,nextnkinset_kbn,sqdata_startymd,sqdata_endymd,history," & _
                '                           "biko2,hoken_biko,nkinsime_ymd,yatin_jisansaki,hoken_kikan," & _
                '                           "hoken_gak,confirmky_sekininsya,confirmky_ymd,confirmky_print,confirmkai_sekininsya," & _
                '                           "confirmkai_ymd,confirmkai_print,hikiuke_name,hikiuke_addr,hikiuke_tel," & _
                '                           "kohokennkin_ymd,kokanryo_ymd,kokanryotuti_ymd,kotuti_ymd,kosaisoku_ymd," & _
                '                           "kosyoruiuke_ymd,kokairenraku_umu,kokairenraku_ymd,kokairenraku_logonuser_no,kokairenraku_biko," & _
                '                           "maitukisq_umu,hikiuke_namesjis,kofkae_ymd,yokugetu_uketorimonth,ky_logonuser_no," & _
                '                           "sq_logonuser_no"
                Dim fldnamegrp As String = "ky_guid,ky_recno,ko_no,henko_no,sqdata_guid," & _
                                           "ky_bango,ky_ymd,kystart_ymd,kyend_ymd,henko_ymd," & _
                                           "tuti_ymd,print_ymd,kanri_gy_fudono,gy_hosyono,hosyo_naiyo," & _
                                           "kokyaku_bango,kyrui_no,siyo_mokuteki,kosin_umu,yatin_kbn," & _
                                           "fkae_startym,fkae_willstartflg,yatin_kozakbn,yatin_kozano,maitukiyatin_kozano," & _
                                           "yokugetu_uketoriflg,biko,tougetu_sagakuflg,nextky_startymd,nextky_endymd," & _
                                           "kosin_hiwariflg,sokojorule_kbn,soyotei_ymdflg,soyotei_ymd,kanritesu_flg," & _
                                           "kanritesu_gak,nextnkinset_kbn,sqdata_startymd,sqdata_endymd,history," & _
                                           "biko2,hoken_biko,nkinsime_ymd,yatin_jisansaki,hoken_kikan," & _
                                           "hoken_gak,hikiuke_name,hikiuke_addr,hikiuke_tel,kohokennkin_ymd," & _
                                           "kokanryo_ymd,kokanryotuti_ymd,kotuti_ymd,kosaisoku_ymd,kosyoruiuke_ymd," & _
                                           "maitukisq_umu,hikiuke_namesjis,kofkae_ymd,yokugetu_uketorimonth,ky_logonuser_no," & _
                                           "sq_logonuser_no,torihiki_tesumoto,torihiki_tesukyaku,torihiki_tesugak,torihiki_tesuzeikbn," & _
                                           "torihiki_tesuzeigak"
                '20160519 EXEUpdateに伴う修正 契約履歴情報 -chg end
                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "更新No"
                                    .Vari_Ko_no = fldvalue.Trim
                                Case "改定No"
                                    .Vari_Henko_no = fldvalue.Trim
                                Case "契約番号"
                                    .Vari_Ky_bango = fldvalue.Trim
                                Case "契約日"
                                    .Vari_Ky_ymd = fldvalue.Trim
                                Case "契約開始日"
                                    .Vari_Kystart_ymd = fldvalue.Trim
                                Case "契約終了日"
                                    .Vari_Kyend_ymd = fldvalue.Trim
                                Case "条件変更日"
                                    .Vari_Henko_ymd = fldvalue.Trim
                                Case "契約更新通知日"
                                    .Vari_Tuti_ymd = fldvalue.Trim
                                Case "通知書印刷日"
                                    .Vari_Print_ymd = fldvalue.Trim
                                Case "管理業者No"
                                    .Vari_Kanri_gy_fudono = fldvalue.Trim
                                Case "賃貸保証業者No"
                                    .Vari_Gy_hosyono = fldvalue.Trim
                                Case "賃貸保証内容"
                                    .Vari_Hosyo_naiyo = fldvalue.Trim
                                Case "賃貸保証顧客番号"
                                    .Vari_Kokyaku_bango = fldvalue.Trim
                                Case "契約分類名"   '2016.04.06 契約分類を紐付データを元に移行 「契約分類No」→「契約分類名」へ変更
                                    .Vari_Kyrui_no = fldvalue.Trim
                                Case "使用目的"
                                    .Vari_Siyo_mokuteki = fldvalue.Trim
                                Case "契約更新業務有無"
                                    .Vari_Kosin_umu = fldvalue.Trim
                                Case "家賃入金区分"
                                    .Vari_Yatin_kbn = fldvalue.Trim
                                Case "口座振替開始日"
                                    .Vari_Fkae_startym = fldvalue.Trim
                                Case "口座振替開始待ちフラグ"
                                    .Vari_Fkae_willstartflg = fldvalue.Trim
                                Case "家賃入金口座区分"
                                    .Vari_Yatin_kozakbn = fldvalue.Trim
                                Case "契約一時金入金口座No"
                                    .Vari_Yatin_kozano = fldvalue.Trim
                                Case "毎月分入金口座No"
                                    .Vari_Maitukiyatin_kozano = fldvalue.Trim
                                Case "翌月分受取り有無"
                                    .Vari_Yokugetu_uketoriflg = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko = fldvalue.Trim
                                Case "当月分の差額受取り有無"
                                    .Vari_Tougetu_sagakuflg = fldvalue.Trim
                                Case "次回契約開始日"
                                    .Vari_Nextky_startymd = fldvalue.Trim
                                Case "次回契約終了日"
                                    .Vari_Nextky_endymd = fldvalue.Trim
                                Case "更新時日割り有無"
                                    .Vari_Kosin_hiwariflg = fldvalue.Trim
                                Case "送金控除ルール適用有無"
                                    .Vari_Sokojorule_kbn = fldvalue.Trim
                                Case "送金予定日使用フラグ"
                                    .Vari_Soyotei_ymdflg = fldvalue.Trim
                                Case "送金予定日"
                                    .Vari_Soyotei_ymd = fldvalue.Trim
                                Case "部屋固定管理手数料フラグ"
                                    .Vari_Kanritesu_flg = fldvalue.Trim
                                Case "部屋固定管理手数料額"
                                    .Vari_Kanritesu_gak = fldvalue.Trim
                                Case "次回更新設定区分"
                                    .Vari_Nextnkinset_kbn = fldvalue.Trim
                                Case "請求データ作成開始日"
                                    .Vari_Sqdata_startymd = fldvalue.Trim
                                Case "請求データ作成終了日"
                                    .Vari_Sqdata_endymd = fldvalue.Trim
                                Case "備考2(基本)"
                                    .Vari_Biko2 = fldvalue.Trim
                                Case "備考(保険)"
                                    .Vari_Hoken_biko = fldvalue.Trim
                                Case "入金締め日"
                                    .Vari_Nkinsime_ymd = fldvalue.Trim
                                Case "家賃持参先"
                                    .Vari_Yatin_jisansaki = fldvalue.Trim
                                Case "保険期間"
                                    .Vari_Hoken_kikan = fldvalue.Trim
                                Case "保険額"
                                    .Vari_Hoken_gak = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約履歴情報 -del sta
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
                                    '20160519 EXEUpdateに伴う修正 契約履歴情報 -del end
                                Case "身元引受人名"
                                    .Vari_Hikiuke_name = fldvalue.Trim
                                Case "身元引受人住所"
                                    .Vari_Hikiuke_addr = fldvalue.Trim
                                Case "身元引受人連絡先"
                                    .Vari_Hikiuke_tel = fldvalue.Trim
                                Case "保険料入金日"
                                    .Vari_Kohokennkin_ymd = fldvalue.Trim
                                Case "更新完了日"
                                    .Vari_Kokanryo_ymd = fldvalue.Trim
                                Case "更新完了通知日"
                                    .Vari_Kokanryotuti_ymd = fldvalue.Trim
                                Case "更新通知日"
                                    .Vari_Kotuti_ymd = fldvalue.Trim
                                Case "催促実施日"
                                    .Vari_Kosaisoku_ymd = fldvalue.Trim
                                Case "書類返送受取日"
                                    .Vari_Kosyoruiuke_ymd = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約履歴情報 -del sta
                                    'Case "更新時の解約検討連絡有無"
                                    '    .Vari_Kokairenraku_umu = fldvalue.Trim
                                    'Case "更新時の解約検討連絡受付日"
                                    '    .Vari_Kokairenraku_ymd = fldvalue.Trim
                                    'Case "更新時の解約検討連絡受付担当者"
                                    '    .Vari_Kokairenraku_logonuser_no = fldvalue.Trim
                                    'Case "更新時の解約検討連絡備考"
                                    '    .Vari_Kokairenraku_biko = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約履歴情報 -del end
                                Case "毎月分請求有無"
                                    .Vari_Maitukisq_umu = fldvalue.Trim
                                Case "身元引受人名SJIS"
                                    .Vari_Hikiuke_namesjis = fldvalue.Trim
                                Case "口座振替日"
                                    .Vari_Kofkae_ymd = fldvalue.Trim
                                Case "翌月分受取り月"
                                    .Vari_Yokugetu_uketorimonth = fldvalue.Trim
                                Case "契約担当者No"
                                    .Vari_Ky_logonuser_no = fldvalue.Trim
                                Case "請求担当者No"
                                    .Vari_Sq_logonuser_no = fldvalue.Trim
                                    '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add sta
                                Case "配分割合_元付"
                                    .Vari_Torihiki_tesumoto = fldvalue
                                Case "配分割合_客付"
                                    .Vari_Torihiki_tesukyaku = fldvalue
                                Case "客付会社の手数料額"
                                    .Vari_Torihiki_tesugak = fldvalue
                                Case "客付会社の手数料税区分"
                                    .Vari_Torihiki_tesuzeikbn = fldvalue
                                Case "客付会社の手数料税額"
                                    .Vari_Torihiki_tesuzeigak = fldvalue
                                    '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add end
                            End Select

                        Next

                        '固定値
                        .Vari_Sqdata_guid = Guid.NewGuid.ToString
                        .Vari_History = DefHistory

                        '20170525 契約情報が自動更新の場合の次回契約期間を空にする修正対応 -add sta
                        '自動更新の場合は次回契約期間を空にする
                        If .Vari_Kosin_umu = "2" Then
                            .Vari_Nextky_startymd = ""
                            .Vari_Nextky_endymd = ""
                        End If
                        '20170525 契約情報が自動更新の場合の次回契約期間を空にする修正対応 -add end

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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 入金区分紐付情報取得
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub Get_RelData_NkinKbn()

                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
                Dim startrow As Integer                                         '書込開始行
                Dim columncnt As Integer                                        '列数
                Dim maxrowcnt As Integer                                        '既存データの行数
                Dim rowcnt As Integer                                           '書込行数
                Dim relsheetname As String = "入金区分マスタ"
                Dim rtn As Boolean = True

                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

                '既存データ有無確認(入金項目は他の項目でも参照するため)
                If Hash_Rel_Nkinkbn.Count <> 0 Then
                    Exit Sub
                End If

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

                For cntii = 0 To rowcnt - 1

                    '作業用変数作成
                    Dim tmp_oldnkinkbnname As String = ""
                    Dim tmp_newnkinkbnno As String = ""

                    '紐付設定値取得
                    For cntjj = 1 To columncnt

                        'ヘッダー格納
                        Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                        '移行値格納
                        Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                        Select Case fldname
                            Case "移行元入金区分名称"
                                tmp_oldnkinkbnname = fldvalue.Trim
                            Case "賃貸革命10入金区分No"
                                tmp_newnkinkbnno = fldvalue.Trim
                        End Select

                    Next

                    'ハッシュテーブル格納
                    If tmp_newnkinkbnno <> "" Then
                        Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
                    End If

                Next

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約契約者情報"

    Public Class Kydata_kys_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_kys_Model         '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_kys"
                Dim fldnamegrp As String = "ky_guid,ky_recno,kys_cnt,kys_no,nyukyo_flg," & _
                                           "history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "並び順No"
                                    .Vari_Kys_cnt = fldvalue.Trim
                                Case "契約者No"
                                    .Vari_Kys_no = fldvalue.Trim
                                Case "入居フラグ"
                                    .Vari_Nyukyo_flg = fldvalue.Trim
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

                        '20161130 契約者入居フラグの移行処理修正_再 -del sta
                        ''20161128 契約者入居フラグの移行処理修正 -add sta
                        ''入居フラグが0の場合は値をそのまま移行する
                        'If tmp_hash("nyukyo_flg") = "0" Then
                        '    hash_cvitem("nyukyo_flg") = "0"
                        'End If
                        ''20161128 契約者入居フラグの移行処理修正 -add end
                        '20161130 契約者入居フラグの移行処理修正_再 -del end

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '20170116 契約情報の契約者情報ダミーレコード作成処理の追加 -add sta
                If CNVNO = ConvertTypes._汎用 Then
                    Try
                        Call Me.Set_KysDummyRecord(sqlcnnv10)
                    Catch ex As Exception
                        '----- ログ出力 -----
                        Dim tmptmpcnt As Integer = 0
                        Dim tmptmpstr As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(41, sheetname, "データ調整", "データ調整中にエラーが発生しました。"), False)
                        DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, tmptmpcnt)
                    End Try
                End If
                '20170116 契約情報の契約者情報ダミーレコード作成処理の追加 -add end

                '************************
                '終了処理
                '************************

                '中間ファイル件数を取得
                midrowcnt = rowcnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

            ''' <summary>
            ''' 契約情報の契約者情報へダミーレコードを作成する
            ''' </summary>
            ''' <param name="sqlcnnv10">賃貸10DB接続用オブジェクト</param>
            ''' <remarks>
            ''' 20170116 契約情報の契約者情報ダミーレコード作成処理の追加 新規追加
            ''' 　契約情報登録時に契約者No1～3まで自動でレコードが作成されるためこれの対応を行う
            ''' 　※契約者No = 2、3が未設定でも作成される
            ''' 　TSから頂いたクエリをそのまま流用する
            ''' </remarks>
            Public Sub Set_KysDummyRecord(ByVal sqlcnnv10 As SqlConnection)

                Dim tmp_sql As String = ""
                Dim tmpcnt As Integer = 0

                '①kys_cnt=1しかないデータに対してkys_cnt=2を作る
                tmp_sql = tmp_sql & " INSERT INTO  "
                tmp_sql = tmp_sql & " 	kydata_kys "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		 ky_guid "
                tmp_sql = tmp_sql & " 		,ky_recno "
                tmp_sql = tmp_sql & " 		,kys_cnt "
                tmp_sql = tmp_sql & " 		,kys_no "
                tmp_sql = tmp_sql & " 		,nyukyo_flg "
                tmp_sql = tmp_sql & " 		,history "
                tmp_sql = tmp_sql & " 	) "
                tmp_sql = tmp_sql & " select  "
                tmp_sql = tmp_sql & " 	 ky_guid "
                tmp_sql = tmp_sql & " 	,ky_recno "
                tmp_sql = tmp_sql & " 	,2 kys_cnt "
                tmp_sql = tmp_sql & " 	,null "
                tmp_sql = tmp_sql & " 	,0 "
                tmp_sql = tmp_sql & " 	,history "
                tmp_sql = tmp_sql & " from kydata_kys "
                tmp_sql = tmp_sql & " WHERE ky_guid not in (select distinct ky_guid from kydata_kys where kys_cnt in (2,3)) "
                tmp_sql = tmp_sql & " AND kys_cnt = 1 "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmpcnt)
                tmp_sql = ""

                '②kys_cnt=3がないデータに対してkys_cnt=3を作る
                '※過去データ調整で対応したクエリを一部修正して流用する
                tmp_sql = tmp_sql & " INSERT INTO  "
                tmp_sql = tmp_sql & " 	kydata_kys "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		 ky_guid "
                tmp_sql = tmp_sql & " 		,ky_recno "
                tmp_sql = tmp_sql & " 		,kys_cnt "
                tmp_sql = tmp_sql & " 		,kys_no "
                tmp_sql = tmp_sql & " 		,nyukyo_flg "
                tmp_sql = tmp_sql & " 		,history "
                tmp_sql = tmp_sql & " 	) "
                tmp_sql = tmp_sql & " select  "
                tmp_sql = tmp_sql & " 	 ky_guid "
                tmp_sql = tmp_sql & " 	,ky_recno "
                tmp_sql = tmp_sql & " 	,3 kys_cnt "
                tmp_sql = tmp_sql & " 	,null "
                tmp_sql = tmp_sql & " 	,0 "
                tmp_sql = tmp_sql & " 	,history "
                tmp_sql = tmp_sql & " from kydata_kys "
                tmp_sql = tmp_sql & " WHERE ky_guid not in (select distinct ky_guid from kydata_kys where kys_cnt in (3)) "
                tmp_sql = tmp_sql & " AND kys_cnt = 1 "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmpcnt)
                tmp_sql = ""

            End Sub

        End Class

    End Class

#End Region

#Region "契約入居者情報"

    Public Class Kydata_nyukyo_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_nyukyo_Model      '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_nyukyo"
                Dim fldnamegrp As String = "ky_guid,ky_recno,nyukyosya_cnt,kys_no,name," & _
                                           "nameu,kana,gender,aidagara,birthday," & _
                                           "kinmusaki,tel,mobiletel,biko,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "入居者No"
                                    .Vari_Nyukyosya_cnt = fldvalue.Trim
                                Case "契約者No"
                                    .Vari_Kys_no = fldvalue.Trim
                                Case "氏名"
                                    '"氏名"はsjisのフィールドと思われるが[name]へ移行しても10に反映されないため[nameu]へ移行する
                                    .Vari_Nameu = fldvalue.Trim
                                Case "氏名Unicode"
                                    '"氏名"の値を[nameu]へ移行するためここでは何もしない
                                Case "カナ"
                                    .Vari_Kana = fldvalue.Trim
                                Case "性別"
                                    .Vari_Gender = fldvalue.Trim
                                Case "続柄"
                                    .Vari_Aidagara = fldvalue.Trim
                                Case "生年月日"
                                    .Vari_Birthday = fldvalue.Trim
                                Case "勤務先"
                                    .Vari_Kinmusaki = fldvalue.Trim
                                Case "連絡先"
                                    .Vari_Tel = fldvalue.Trim
                                Case "携帯電話番号"
                                    .Vari_Mobiletel = fldvalue.Trim
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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,nyukyosya_cnt) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約保証人情報"

    Public Class Kydata_hosyonin_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_hosyonin_Model    '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_hosyonin"
                Dim fldnamegrp As String = "ky_guid,ky_recno,hosyonin_cnt,kys_no,hosyonin_no," & _
                                           "history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "並び順"
                                    .Vari_Hosyonin_cnt = fldvalue.Trim
                                Case "契約者No"
                                    .Vari_Kys_no = fldvalue.Trim
                                Case "保証人No"
                                    .Vari_Hosyonin_no = fldvalue.Trim
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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,hosyonin_cnt) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約車情報"

    Public Class Kydata_car_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_car_Model         '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

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
                '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_car"
                Dim fldnamegrp As String = "ky_guid,ky_recno,car_cnt,carmaker,carname," & _
                                           "carcolor,carnumber,biko,history,car_kukaku"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "車情報No"
                                    .Vari_Car_cnt = fldvalue.Trim
                                Case "メーカー"
                                    .Vari_Carmaker = fldvalue.Trim
                                Case "車名"
                                    .Vari_Carname = fldvalue.Trim
                                Case "車色"
                                    .Vari_Carcolor = fldvalue.Trim
                                Case "ナンバー"
                                    .Vari_Carnumber = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko = fldvalue.Trim
                                Case "駐車区画"
                                    .Vari_Car_kukaku = fldvalue.Trim
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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,car_cnt) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約保険情報 (10でも履歴管理されていない)"

    Public Class Kydata_hoken_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_hoken_Model       '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_hoken"
                Dim fldnamegrp As String = "ky_guid,hoken_no,gy_hokenno,ky_ymd,kystart_ymd," & _
                                           "kyend_ymd,hoken_gak,mankituti_flg,biko,history," & _
                                           "syoken_bango"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "保険種類No"
                                    .Vari_Hoken_no = fldvalue.Trim
                                Case "保険業者No"
                                    .Vari_Gy_hokenno = fldvalue.Trim
                                Case "契約日"
                                    .Vari_Ky_ymd = fldvalue.Trim
                                Case "適用開始年月日"
                                    .Vari_Kystart_ymd = fldvalue.Trim
                                Case "適用終了年月日"
                                    .Vari_Kyend_ymd = fldvalue.Trim
                                Case "保険金額"
                                    .Vari_Hoken_gak = fldvalue.Trim
                                Case "満期案内通知有無"
                                    .Vari_Mankituti_flg = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko = fldvalue.Trim
                                Case "証券番号"
                                    .Vari_Syoken_bango = fldvalue.Trim
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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,hoken_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約特約事項情報"

    Public Class Kydata_tokuyaku_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_tokuyaku_Model    '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"
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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_tokuyaku"
                Dim fldnamegrp As String = "ky_guid,ky_recno,tokuyaku_grpno,naiyo,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                    Dim fldname_keysub3 As String = readtbl.Columns(keycol_sub3 - 1).ColumnName.Trim
                    Dim fldname_keysub4 As String = readtbl.Columns(keycol_sub4 - 1).ColumnName.Trim

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
                        Dim tmp_keysub3 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub3 - 1)).Trim
                        Dim tmp_keysub4 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub4 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4
                        Dim tmp_keyoya As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        Dim tmp_keydup As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4

                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Ky_recno = tmp_keysub3
                        .Vari_Tokuyaku_grpno = tmp_keysub4

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4

                        '特約カウント初期化
                        Dim cvitemcnt As Integer = 0
                        Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

                        'データ有無チェック
                        Dim tmp_fldvalueumuchk As String = ""
                        For cntjj = 5 To readtbl.Columns.Count - 1
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
                                Dim log_key As String = "kydata_tokuyaku-ky_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(tmp_keydup) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "kydata_tokuyaku-ky_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            If keychkflg Then

                                '作業用変数作成
                                Dim tmp_tokuyaku As String = ""

                                '移行値取得
                                For cntjj = 5 To readtbl.Columns.Count - 1

                                    'サブキー取得
                                    Dim tmp_keysub5 As String = (cntjj - 4).ToString

                                    'ログ出力用データ格納(サブキーフィールド)
                                    Dim fldname_keysub As String = readtbl.Columns(cntjj).ColumnName.Trim

                                    '登録値取得
                                    Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                                    If fldvalue <> "" Then
                                        tmp_tokuyaku = tmp_tokuyaku & LINE_BREAK & fldvalue
                                    End If

                                Next

                                '2016.04.26 メインの方へも反映させる修正 -chg sta
                                '成形
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
                                        ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                                        'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                                        ''2016.04.26 メインの方へも反映させる修正 -chg end
                                        hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,tokuyaku_grpno) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約メモ情報"

    Public Class Kydata_memo_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_memo_Model        '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"
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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_memo"
                Dim fldnamegrp As String = "ky_guid,ky_recno,memo_no,memo,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                    Dim fldname_keysub3 As String = readtbl.Columns(keycol_sub3 - 1).ColumnName.Trim

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
                        Dim tmp_keysub3 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub3 - 1)).Trim
                        Dim tmp_keyoya As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        Dim tmp_keydup As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                        .Vari_Ky_recno = tmp_keysub3

                        '備考カウント初期化
                        Dim cvitemcnt As Integer = 0
                        Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

                        'データ有無チェック
                        Dim tmp_fldvalueumuchk As String = ""
                        For cntjj = 4 To readtbl.Columns.Count - 1
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
                                Dim log_key As String = "kydata_memo-ky_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(tmp_keydup) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "kydata_memo-ky_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            If keychkflg Then

                                '移行値取得
                                For cntjj = 4 To readtbl.Columns.Count - 1

                                    'サブキー取得
                                    Dim tmp_keysub4 As String = (cntjj - 3).ToString

                                    '全キー取得
                                    Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4

                                    'ログ出力用データ格納(サブキーフィールド)
                                    Dim fldname_keysub4 As String = readtbl.Columns(cntjj).ColumnName.Trim
                                    Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                               fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                               fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                               fldname_keysub3 & " = " & tmp_keysub3

                                    '登録値取得
                                    Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                                    .Vari_Memo = fldvalue

                                    '固定値
                                    .Vari_Memo_no = (cntjj - 3).ToString
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
                                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,memo_no) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約入金項目情報"

    Public Class Kydata_nkin_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_nkin_Model        '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim keycol_sub5 As Integer = 6                                  'サブ5キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '2016.04.06 入金項目読込処理の修正 -chg sta
                'Call Me.Get_RelData_NkinKomk()
                'Call Me.Get_RelData_NkinKbn()
                Call SetRelItemToObject.Set_RelData_Nkinkomk()
                Call SetRelItemToObject.Set_RelData_Nkinkbn()
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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_nkin"
                Dim fldnamegrp As String = "ky_guid,ky_recno,tuki_kbn,nkin_no,nkin_recno," & _
                                           "nkin_sortorder,nkin_kbn,sq_gak,sq_zeikbn,sq_zeigak," & _
                                           "calc_kbn,calc_nkinno,calc_monthcnt,sqsaki_no,nkbn_yotei," & _
                                           "sq_mmkbn,frstart_ymd,frend_ymd,frsq_gak,sqstart_ymd," & _
                                           "sq_ptn,sq_interval,sq_nen,sq_tuki,zei_rit," & _
                                           "biko,nkin_guid,history,fr_kbn"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                    Dim fldname_keysub4 As String = readtbl.Columns(keycol_sub4 - 1).ColumnName.Trim
                    Dim fldname_keysub5 As String = readtbl.Columns(keycol_sub5 - 1).ColumnName.Trim

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
                        Dim tmp_nkinname As String = ""         '2016.04.06 入金項目読込処理の修正 -add
                        Dim tmp_nkinname_kijyun As String = ""  '20160530 ログ修正 -add
                        Dim tmp_sqkankaku As String = ""        '2016.04.26 メインの方へも反映させる修正 -add

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim tmp_keysub3 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub3 - 1)).Trim
                        Dim tmp_keysub4 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub4 - 1)).Trim
                        Dim tmp_keysub5 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub5 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4 & "-" & tmp_keysub5
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4 & "、" & _
                                                   fldname_keysub5 & " = " & tmp_keysub5

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue
                                Case "月区分"
                                    .Vari_Tuki_kbn = fldvalue
                                Case "入金項目名"
                                    '2016.04.06 入金項目読込処理の修正 -chg sta
                                    '.Vari_Nkin_no = fldvalue
                                    tmp_nkinname = fldvalue
                                    '2016.04.06 入金項目読込処理の修正 -chg end
                                Case "入金項目レコードNo"
                                    .Vari_Nkin_recno = fldvalue
                                Case "並び順No"
                                    .Vari_Nkin_sortorder = fldvalue
                                Case "入金項目区分"
                                    .Vari_Nkin_kbn = fldvalue
                                Case "請求額"
                                    .Vari_Sq_gak = fldvalue
                                Case "税区分"
                                    .Vari_Sq_zeikbn = fldvalue
                                Case "請求税額"
                                    .Vari_Sq_zeigak = fldvalue
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
                                Case "入金方法"
                                    .Vari_Nkbn_yotei = fldvalue
                                Case "請求月区分"
                                    .Vari_Sq_mmkbn = fldvalue
                                Case "フリーレント適用開始日"
                                    .Vari_Frstart_ymd = fldvalue
                                Case "フリーレント適用終了日"
                                    .Vari_Frend_ymd = fldvalue
                                Case "フリーレント終了月請求額"
                                    .Vari_Frsq_gak = fldvalue
                                Case "請求開始月"
                                    .Vari_Sqstart_ymd = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -del sta
                                    '必要な情報をまとめて格納するためここではコメントアウト
                                    'Case "請求パターン"
                                    '    .Vari_Sq_ptn = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -del end
                                Case "請求間隔"
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
                                Case "適用税率"
                                    .Vari_Zei_rit = fldvalue
                                Case "備考"
                                    .Vari_Biko = fldvalue
                                Case "フリーレント適用区分"
                                    .Vari_Fr_kbn = fldvalue
                            End Select

                        Next

                        '2016.04.26 メインの方へも反映させる修正 -chg sta
                        ''2016.04.06 入金項目読込処理の修正 -add sta
                        ''紐付用に成形
                        '.Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname(.Vari_Nkin_kbn)
                        ''2016.04.06 入金項目読込処理の修正 -add end

                        '紐付用に成形
                        If .Vari_Tuki_kbn = "2" And .Vari_Nkin_kbn = "4" Then       '解約時の通常月入金項目は通常月の入金項目へ紐付ける
                            .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname("1")
                        Else
                            .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname(.Vari_Nkin_kbn)
                        End If
                        '2016.04.26 メインの方へも反映させる修正 -chg end

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
                        .Vari_Nkin_guid = Guid.NewGuid.ToString
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        '20160617 マイナス金額移行処理修正 -add sta
                        'マイナス金額を予めチェックしておく
                        Call Me.Chk_MinusSqgak(tmp_hash)
                        '20160617 マイナス金額移行処理修正 -add end

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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                'tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                'tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            '2016.04.06 入金項目読込処理の修正 -del sta
            ' ''' <summary>
            ' ''' 入金項目紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKomk()

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

            ' ''' <summary>
            ' ''' 入金区分紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKbn()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "入金区分マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    '既存データ有無確認(入金項目は他の項目でも参照するため)
            '    If Hash_Rel_Nkinkbn.Count <> 0 Then
            '        Exit Sub
            '    End If

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldnkinkbnname As String = ""
            '        Dim tmp_newnkinkbnno As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元入金区分名称"
            '                    tmp_oldnkinkbnname = fldvalue.Trim
            '                Case "賃貸革命10入金区分No"
            '                    tmp_newnkinkbnno = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newnkinkbnno <> "" Then
            '            Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub
            '2016.04.06 入金項目読込処理の修正 -del end

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

            ''' <summary>
            ''' マイナス請求の補正 '20160617 マイナス金額移行処理修正 -add
            ''' 革命10にはマイナス金額を設定することができないためV7のマイナス金額を補正する
            ''' 「敷金差額」「保証金差額」「敷金戻し」「保証金戻し」の属性に紐付いていることが条件
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Chk_MinusSqgak(ByRef hash As Hashtable)

                Dim chg_sqgak As String = ""
                Dim tmp_sqgak As String = hash.Item("sq_gak")
                Dim tmp_nkinname As String = hash.Item("nkin_no")
                Dim tmp_nkinzkseino As String = ""
                '20161012 敷金、保証金差額のマイナス値移行対応 -chg sta
                'Dim list_taisyozkseino As New List(Of String) From {"400", "420"}       '「敷金戻し」「保証金戻し」の属性No
                Dim list_taisyozkseino As New List(Of String) From {"320", "330", "400", "420"}       '「敷金差額」「保証金差額」「敷金戻し」「保証金戻し」の属性No
                '20161012 敷金、保証金差額のマイナス値移行対応 -chg end
                'マイナスの値か判別→それ以外は処理を抜ける
                Dim tmp_dbl As Double = 0
                If Double.TryParse(tmp_sqgak, tmp_dbl) = False Then
                    Exit Sub
                ElseIf tmp_dbl >= 0 Then
                    Exit Sub
                End If

                '入金項目属性Noとの照合→金額補正
                If Hash_Rel_NkinkomkZksei.Contains(tmp_nkinname) Then
                    tmp_nkinzkseino = Hash_Rel_NkinkomkZksei.Item(tmp_nkinname)

                    If list_taisyozkseino.Contains(tmp_nkinzkseino) Then

                        '「敷金戻し」「保証金戻し」の属性Noに該当する場合は絶対値を取得して再格納(マイナスをプラスへ変更)
                        chg_sqgak = Math.Abs(tmp_dbl)
                        hash("sq_gak") = chg_sqgak.ToString

                    End If

                End If

            End Sub

        End Class

    End Class

#End Region

#Region "契約次回入金項目情報"

    Public Class Kydata_nkin_nx_Repository

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
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_nkin_nx_Model     '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim keycol_sub5 As Integer = 6                                  'サブ5キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '2016.04.06 入金項目読込処理の修正 -chg sta
                'Call Me.Get_RelData_NkinKomk()
                'Call Me.Get_RelData_NkinKbn()
                Call SetRelItemToObject.Set_RelData_Nkinkomk()
                Call SetRelItemToObject.Set_RelData_Nkinkbn()
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
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_nkin_nx"
                Dim fldnamegrp As String = "ky_guid,ky_recno,tuki_kbn,nkin_no,nkin_recno," & _
                                           "nkin_sortorder,nkin_kbn,sq_gak,sq_zeikbn,sq_zeigak," & _
                                           "calc_kbn,calc_nkinno,calc_monthcnt,sqsaki_no,nkbn_yotei," & _
                                           "sq_mmkbn,sqstart_ymd,sq_ptn,sq_interval,sq_nen," & _
                                           "sq_tuki,zei_rit,biko,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                    Dim fldname_keysub4 As String = readtbl.Columns(keycol_sub4 - 1).ColumnName.Trim
                    Dim fldname_keysub5 As String = readtbl.Columns(keycol_sub5 - 1).ColumnName.Trim

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
                        Dim tmp_nkinname As String = ""         '2016.04.06 入金項目読込処理の修正 -add
                        Dim tmp_nkinname_kijyun As String = ""  '20160530 ログ修正 -add

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub1 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub1 - 1)).Trim
                        Dim tmp_keysub2 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub2 - 1)).Trim
                        Dim tmp_keysub3 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub3 - 1)).Trim
                        Dim tmp_keysub4 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub4 - 1)).Trim
                        Dim tmp_keysub5 As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub5 - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4 & "-" & tmp_keysub5
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4 & "、" & _
                                                   fldname_keysub5 & " = " & tmp_keysub5

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue
                                Case "月区分"
                                    .Vari_Tuki_kbn = fldvalue
                                Case "入金項目名"
                                    '2016.04.06 入金項目読込処理の修正 -chg sta
                                    '.Vari_Nkin_no = fldvalue
                                    tmp_nkinname = fldvalue
                                    '2016.04.06 入金項目読込処理の修正 -chg end
                                Case "入金項目レコードNo"
                                    .Vari_Nkin_recno = fldvalue
                                Case "並び順No"
                                    .Vari_Nkin_sortorder = fldvalue
                                Case "入金項目区分"
                                    .Vari_Nkin_kbn = fldvalue
                                Case "請求額"
                                    .Vari_Sq_gak = fldvalue
                                Case "税区分"
                                    .Vari_Sq_zeikbn = fldvalue
                                Case "請求税額"
                                    .Vari_Sq_zeigak = fldvalue
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
                                Case "入金方法"
                                    .Vari_Nkbn_yotei = fldvalue
                                Case "請求月区分"
                                    .Vari_Sq_mmkbn = fldvalue
                                Case "請求開始月"
                                    .Vari_Sqstart_ymd = fldvalue
                                Case "請求パターン"
                                    .Vari_Sq_ptn = fldvalue
                                Case "請求発生間隔"
                                    .Vari_Sq_interval = fldvalue
                                Case "請求対象年"
                                    .Vari_Sq_nen = fldvalue
                                Case "請求対象月"
                                    .Vari_Sq_tuki = fldvalue
                                Case "適用税率"
                                    .Vari_Zei_rit = fldvalue
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

                        '固定値
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        '20161012 敷金、保証金差額のマイナス値移行対応 -add sta
                        'マイナス金額を予めチェックしておく
                        Call Me.Chk_MinusSqgak(tmp_hash)
                        '20161012 敷金、保証金差額のマイナス値移行対応 -add end

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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
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
                'tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                'tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            '2016.04.06 入金項目読込処理の修正 -del sta
            ' ''' <summary>
            ' ''' 入金項目紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKomk()

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

            ' ''' <summary>
            ' ''' 入金区分紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKbn()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "入金区分マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    '既存データ有無確認(入金項目は他の項目でも参照するため)
            '    If Hash_Rel_Nkinkbn.Count <> 0 Then
            '        Exit Sub
            '    End If

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldnkinkbnname As String = ""
            '        Dim tmp_newnkinkbnno As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元入金区分名称"
            '                    tmp_oldnkinkbnname = fldvalue.Trim
            '                Case "賃貸革命10入金区分No"
            '                    tmp_newnkinkbnno = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newnkinkbnno <> "" Then
            '            Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub
            '2016.04.06 入金項目読込処理の修正 -del end

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

            ''' <summary>
            ''' マイナス請求の補正 '20161012 敷金、保証金差額のマイナス値移行対応 -add
            ''' 革命10にはマイナス金額を設定することができないためV7のマイナス金額を補正する
            ''' 「敷金差額」「保証金差額」「敷金戻し」「保証金戻し」の属性に紐付いていることが条件
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Chk_MinusSqgak(ByRef hash As Hashtable)

                Dim chg_sqgak As String = ""
                Dim tmp_sqgak As String = hash.Item("sq_gak")
                Dim tmp_nkinname As String = hash.Item("nkin_no")
                Dim tmp_nkinzkseino As String = ""
                Dim list_taisyozkseino As New List(Of String) From {"320", "330", "400", "420"}       '「敷金差額」「保証金差額」「敷金戻し」「保証金戻し」の属性No

                'マイナスの値か判別→それ以外は処理を抜ける
                Dim tmp_dbl As Double = 0
                If Double.TryParse(tmp_sqgak, tmp_dbl) = False Then
                    Exit Sub
                ElseIf tmp_dbl >= 0 Then
                    Exit Sub
                End If

                '入金項目属性Noとの照合→金額補正
                If Hash_Rel_NkinkomkZksei.Contains(tmp_nkinname) Then
                    tmp_nkinzkseino = Hash_Rel_NkinkomkZksei.Item(tmp_nkinname)

                    If list_taisyozkseino.Contains(tmp_nkinzkseino) Then

                        '「敷金戻し」「保証金戻し」の属性Noに該当する場合は絶対値を取得して再格納(マイナスをプラスへ変更)
                        chg_sqgak = Math.Abs(tmp_dbl)
                        hash("sq_gak") = chg_sqgak.ToString

                    End If

                End If

            End Sub

        End Class

    End Class

#End Region

#Region "契約変動費各戸メーター情報"

    Public Class Kydata_hendo_Repository

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
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト(一棟所有)
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                '20160603 ユーザーデータ検証による修正 -del
                'Dim hash_bkhyhendoguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_hendo_Model       '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"
                '20160603 ユーザーデータ検証による修正 -del
                'Dim viewnamebkhytohyhendoguid As String = PRE_VIEW_NAME & "物件部屋変動費キー情報"

                '************************
                '作業準備
                '************************



                '入金項目の重複チェックは特殊なので考慮すること

                '紐付けデータ取得
                '2016.04.06 入金項目読込処理の修正 -chg sta
                'Call Me.Get_RelData_NkinKomk()
                'Call Me.Get_RelData_NkinKbn()
                'Call Me.Get_RelData_Hendometer()
                Call SetRelItemToObject.Set_RelData_Nkinkomk()
                Call SetRelItemToObject.Set_RelData_Nkinkbn()
                '2016.04.06 入金項目読込処理の修正 -chg end

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)
                '20160603 ユーザーデータ検証による修正 -del
                'Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyhendoguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_hendo"
                Dim fldnamegrp As String = "ky_guid,ky_recno,hyhendo_guid,hendo_sortorder,useflg," & _
                                           "hendo_kbn,meter_name,nkin_no,hendorule_no,hendorule_biko," & _
                                           "sqsaki_no,nkbn_yotei,biko,history,sq_mmkbn"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)
                '20160603 ユーザーデータ検証による修正 -del
                'Call Me.Get_Guid(sqlcnnv10, viewnamebkhytohyhendoguid, hash_bkhyhendoguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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

                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)
                    Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)
                    Dim fldname_keysub4 As String = headervalue(startrow - 1, keycol_sub4)

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
                        Dim tmp_keysub4 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub4))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "変動費No"
                                    .Vari_Hendo_sortorder = fldvalue.Trim
                                Case "請求対象フラグ"
                                    .Vari_Useflg = fldvalue.Trim
                                Case "メーター分類"
                                    '.Vari_Hendo_kbn = fldvalue.Trim
                                Case "メーター名"
                                    .Vari_Meter_name = fldvalue.Trim
                                Case "入金項目名"
                                    '2016.04.06 入金項目読込処理の修正 -chg sta
                                    '.Vari_Nkin_no = fldvalue.Trim
                                    '.Vari_Hendo_kbn = fldvalue.Trim     '紐付けたメーター分類を取得するため入金項目名を格納しておく
                                    .Vari_Nkin_no = fldvalue & "-" & EtcMethod.Get_Nkinruiname("5")
                                    .Vari_Hendo_kbn = fldvalue & "-" & EtcMethod.Get_Nkinruiname("5")     '紐付けたメーター分類を取得するため入金項目名を格納しておく
                                    '2016.04.06 入金項目読込処理の修正 -chg end
                                Case "変動費請求ルールNo"
                                    .Vari_Hendorule_no = fldvalue.Trim
                                Case "変動費請求ルール備考"
                                    .Vari_Hendorule_biko = fldvalue.Trim
                                Case "請求先No"
                                    .Vari_Sqsaki_no = fldvalue.Trim
                                Case "入金方法"
                                    .Vari_Nkbn_yotei = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko = fldvalue.Trim
                                Case "請求月"
                                    .Vari_Sq_mmkbn = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_Hyhendo_guid = Guid.NewGuid.ToString
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

                            '20160603 ユーザーデータ検証による修正 -del sta
                            ''変動区分を取得
                            'Dim hendo_guidkey As String = tmp_keymain & "-" & tmp_keysub1 & "-" & hash_cvitem.Item("hendo_kbn")
                            '20160603 ユーザーデータ検証による修正 -del end

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''hash_cvitem("hyhendo_guid") = hash_bkhyhendoguid.Item(hendo_guidkey)        '20160603 ユーザーデータ検証による修正 -del
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            '20160819 部屋の半角変換処理によるエラー修正 -chg end

                            '20160603 ユーザーデータ検証による修正 -add sta
                            Dim kyguid As String = hash_cvitem("ky_guid")
                            Dim hendokbn As String = hash_cvitem("hendo_kbn")
                            Dim nkinno As String = hash_cvitem("nkin_no")
                            Dim tmp_hyhendoguid As String = Me.Get_HyHendoguid(sqlcnnv10, kyguid, hendokbn, nkinno)
                            hash_cvitem("hyhendo_guid") = tmp_hyhendoguid
                            '20160603 ユーザーデータ検証による修正 -add end

                            '挿入処理
                            '20160819 部屋変動費guidが存在しない場合の処理を追加 -chg sta
                            'Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)
                            If tmp_hyhendoguid = "" Then
                                normalflg = False
                                'ログ出力
                                Dim log_key As String = tblname & "-" & "hyhendo_guid"
                                Dim errstr As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                                hash_log.Clear()
                                hash_log.Add(log_key, errstr)
                            Else
                                Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)
                            End If
                            '20160819 部屋変動費guidが存在しない場合の処理を追加 -chg end

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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)

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
                'tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                'tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                Dim tmp_viewmainname As String = viewname.Replace(PRE_VIEW_NAME, "")

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                Select Case tmp_viewmainname
                    Case "物件部屋変動費キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,HYHENDO.hendo_kbn) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,hyhendo_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata_hendo AS HYHENDO "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON HYHENDO.hy_guid = HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                    Case "物件部屋契約キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                        tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "
                End Select

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            '2016.04.06 入金項目読込処理の修正 -del sta
            ' ''' <summary>
            ' ''' 入金項目紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKomk()

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

            ' ''' <summary>
            ' ''' 入金区分紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKbn()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "入金区分マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    '既存データ有無確認(入金項目は他の項目でも参照するため)
            '    If Hash_Rel_Nkinkbn.Count <> 0 Then
            '        Exit Sub
            '    End If

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldnkinkbnname As String = ""
            '        Dim tmp_newnkinkbnno As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元入金区分名称"
            '                    tmp_oldnkinkbnname = fldvalue.Trim
            '                Case "賃貸革命10入金区分No"
            '                    tmp_newnkinkbnno = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newnkinkbnno <> "" Then
            '            Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub

            ' ''' <summary>
            ' ''' 各戸メーター分類紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
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
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

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
                tmp_sql = tmp_sql & " 	SELECT bk_guid FROM kydata AS KY "
                tmp_sql = tmp_sql & " 	WHERE ky_guid IN "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT ky_guid FROM kydata_hendo "
                tmp_sql = tmp_sql & " 		) "
                tmp_sql = tmp_sql & " ); "

                Return tmp_sql

            End Function

            ''' <summary>
            ''' 部屋変動費情報から変動費guidを取得 '20160603 ユーザーデータ検証による修正 新規追加
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hyguid"></param>
            ''' <param name="hendokbn"></param>
            ''' <param name="nkinno"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_HyHendoguid(ByVal sqlcnnv10 As SqlConnection, ByVal kyguid As String, ByVal hendokbn As String, ByVal nkinno As String) As String

                Dim rtn_guid As String = ""
                Dim tmp_sql As String = ""

                '20160616 契約変動費情報の抽出条件不足時の処理の追加 -add sta
                '抽出条件が不足している場合は処理を抜ける
                If kyguid.Trim = "" Or hendokbn.Trim = "" Or nkinno.Trim = "" Then
                    Return rtn_guid
                End If
                '20160616 契約変動費情報の抽出条件不足時の処理の追加 -add end

                tmp_sql = tmp_sql & " SELECT CONVERT(VARCHAR(50),hyhendo_guid) FROM hydata_hendo "
                tmp_sql = tmp_sql & " WHERE hy_guid =  "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT hy_guid FROM kydata WHERE ky_guid = '" & kyguid & "' "
                tmp_sql = tmp_sql & " 	) "
                tmp_sql = tmp_sql & " AND   hendo_kbn = " & hendokbn
                tmp_sql = tmp_sql & " AND   nkin_no = " & nkinno
                rtn_guid = DBExec.Exec_Scalar(tmp_sql, sqlcnnv10)

                Return rtn_guid

            End Function

        End Class

    End Class

#End Region

#Region "契約控除ルール情報"

    Public Class Kydata_kojorule_Repository

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
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト(一棟所有)
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_kojorule_Model    '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim keycol_sub5 As Integer = 6                                  'サブ5キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                '************************
                '作業準備
                '************************

                '入金項目の重複チェックは特殊なので考慮すること

                '紐付けデータ取得
                '2016.04.06 入金項目読込処理の修正 -chg sta
                'Call Me.Get_RelData_NkinKomk()
                'Call Me.Get_RelData_NkinKbn()
                Call SetRelItemToObject.Set_RelData_Nkinkomk()
                Call SetRelItemToObject.Set_RelData_Nkinkbn()
                '2016.04.06 入金項目読込処理の修正 -chg end

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_kojorule"
                Dim fldnamegrp As String = "ky_guid,ky_recno,taisyokbn,nkin_no,nkin_sortorder," & _
                                           "sosai_flg,calc_kbn,kojo_gak,kojo_gakzeikbn,kojonkin_no," & _
                                           "kojo_rit,kojo_ritzeikbn,kojo_ritutizei,zei_rit,tateazu_flg," & _
                                           "kojo_zeigak,sotaisyo_flg,sotaisyonkin_no"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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

                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)
                    Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)
                    Dim fldname_keysub4 As String = headervalue(startrow - 1, keycol_sub4)
                    Dim fldname_keysub5 As String = headervalue(startrow - 1, keycol_sub5)

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
                        Dim tmp_keysub4 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub4))
                        Dim tmp_keysub5 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub5))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4 & "-" & tmp_keysub5
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4 & "、" & _
                                                   fldname_keysub5 & " = " & tmp_keysub5

                        '2016.04.06 入金項目読込処理の修正 -add
                        '作業用変数
                        Dim tmp_nkinname As String = ""

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "対象区分"
                                    .Vari_Taisyokbn = fldvalue.Trim
                                Case "控除入金項目名"
                                    '2016.04.06 入金項目読込処理の修正 -chg sta
                                    '.Vari_Nkin_no = fldvalue.Trim
                                    tmp_nkinname = fldvalue
                                    '2016.04.06 入金項目読込処理の修正 -chg end
                                Case "表示順"
                                    .Vari_Nkin_sortorder = fldvalue.Trim
                                Case "相殺予定フラグ"
                                    .Vari_Sosai_flg = fldvalue.Trim
                                Case "控除額算出基準"
                                    .Vari_Calc_kbn = fldvalue.Trim
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
                                Case "控除率内税"
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
                        .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname("9")
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
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            ''2016.04.26 メインの方へも反映させる修正 -chg sta
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)

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

                '入金項目の重複チェックは特殊なので考慮すること

                'Dim tmp_sql As String = ""
                'tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                'tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            '2016.04.06 入金項目読込処理の修正 -del sta
            ' ''' <summary>
            ' ''' 入金項目紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKomk()

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

            ' ''' <summary>
            ' ''' 入金区分紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKbn()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "入金区分マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    '既存データ有無確認(入金項目は他の項目でも参照するため)
            '    If Hash_Rel_Nkinkbn.Count <> 0 Then
            '        Exit Sub
            '    End If

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldnkinkbnname As String = ""
            '        Dim tmp_newnkinkbnno As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元入金区分名称"
            '                    tmp_oldnkinkbnname = fldvalue.Trim
            '                Case "賃貸革命10入金区分No"
            '                    tmp_newnkinkbnno = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newnkinkbnno <> "" Then
            '            Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub
            '2016.04.06 入金項目読込処理の修正 -del end

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

        End Class

    End Class

#End Region

#Region "契約送金ルール情報"

    Public Class Kydata_sorule_Repository

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
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト(一棟所有)
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_sorule_Model      '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim keycol_sub5 As Integer = 6                                  'サブ5キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                '************************
                '作業準備
                '************************



                '入金項目の重複チェックは特殊なので考慮すること


                '紐付けデータ取得
                '2016.04.06 入金項目読込処理の修正 -chg sta
                'Call Me.Get_RelData_NkinKomk()
                Call SetRelItemToObject.Set_RelData_Nkinkomk()
                '2016.04.06 入金項目読込処理の修正 -chg end

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_sorule"
                Dim fldnamegrp As String = "ky_guid,ky_recno,taisyokbn,nkin_no,nkin_sortorder," & _
                                           "sokin_rit,kanrigak_rit,hosyo_flg"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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

                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)
                    Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)
                    Dim fldname_keysub4 As String = headervalue(startrow - 1, keycol_sub4)
                    Dim fldname_keysub5 As String = headervalue(startrow - 1, keycol_sub5)

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
                        Dim tmp_keysub4 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub4))
                        Dim tmp_keysub5 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub5))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4 & "-" & tmp_keysub5
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4 & "、" & _
                                                   fldname_keysub5 & " = " & tmp_keysub5

                        '2016.04.06 入金項目読込処理の修正 -add
                        '作業用変数
                        Dim tmp_nkinname As String = ""

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "対象区分"
                                    .Vari_Taisyokbn = fldvalue.Trim
                                Case "入金項目名"
                                    '2016.04.06 入金項目読込処理の修正 -chg sta
                                    '.Vari_Nkin_no = fldvalue.Trim
                                    tmp_nkinname = fldvalue
                                    '2016.04.06 入金項目読込処理の修正 -chg end
                                Case "表示順"
                                    .Vari_Nkin_sortorder = fldvalue.Trim
                                Case "送金率"
                                    .Vari_Sokin_rit = fldvalue.Trim
                                Case "管理手数料率"
                                    .Vari_Kanrigak_rit = fldvalue.Trim
                                Case "滞納保証有無"
                                    .Vari_Hosyo_flg = fldvalue.Trim
                            End Select

                        Next

                        '2016.04.26 メインの方へも反映させる修正 -chg sta
                        ''2016.04.06 入金項目読込処理の修正 -add sta
                        ''紐付用に成形
                        '.Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname(.Vari_Taisyokbn)
                        ''2016.04.06 入金項目読込処理の修正 -add end

                        '紐付用に成形
                        Select Case .Vari_Taisyokbn
                            Case "4"
                                .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname("1")     '解約時の通常月入金項目は通常月の入金項目へ紐付ける
                            Case "6"
                                .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname("4")     '解約費の入金項目は解約時の入金項目へ紐付ける
                            Case Else
                                .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname(.Vari_Taisyokbn)
                        End Select
                        '2016.04.26 メインの方へも反映させる修正 -chg end

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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '契約送金ルール区分一括更新
                Dim tmpcnt As Integer = 0
                Dim gaibunoupdatesql As String = Me.Get_UseQry_Update()
                DBExec.Exec_NonQuery(sqlcnnv10, gaibunoupdatesql, tmpcnt)

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
                'tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                'tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                'tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            '2016.04.06 入金項目読込処理の修正 -del sta
            ' ''' <summary>
            ' ''' 入金項目紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_NkinKomk()

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

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

            ''' <summary>
            ''' [kydata_kihon].[sokojorule_kbn]の一括更新クエリ
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Update() As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " UPDATE kydata_kihon SET "
                tmp_sql = tmp_sql & " 	sokojorule_kbn = 2 "
                tmp_sql = tmp_sql & " ; "

                Return tmp_sql

            End Function

        End Class

    End Class

#End Region

#Region "契約解約情報"

    Public Class Kydata_kai_Repository

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
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト(一棟所有)
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                'Dim hash_bkhyguid As New Hashtable                                'bkguid格納用ハッシュテーブル
                Dim hash_bkhykyguid As New Hashtable                              'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_kai_Model         '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                '************************
                '作業準備
                '************************

                '20160629 修繕検証後修正 入金区分の紐付情報取得 -add
                Call SetRelItemToObject.Set_RelData_Nkinkbn()

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_kai"
                '20160620 EXEUpdateに伴う修正2 -chg sta
                ''20160519 EXEUpdateに伴う修正 契約解約情報 -chg sta
                ''Dim fldnamegrp As String = "ky_guid,ky_recno,kai_ymd,seisan_ymd,seisan_completeymd," & _
                ''                           "uketuke_ymd,riyu,tatiai_ymd,uketuke_logonuser_no,tatiai_logonuser_no," & _
                ''                           "seisan_logonuser_no,szen_logonuser_no,szen_jisyano,szen_endymd,szen_kojiyoteistartymd," & _
                ''                           "szen_kojiyoteiendymd,szen_kojibasyo,szen_kojigaiyo,next_name,next_namesjis," & _
                ''                           "next_postcode,next_addr1,next_addr2,next_tel1,yatin_kozano," & _
                ''                           "biko,biko_tatiai,seisan_henkinymd,seisan_sqymd,kanritesu_flg," & _
                ''                           "kanritesu_gak,rowid,history,so_yoteiymd,tatiai," & _
                ''                           "tatiai_yoteiymd,bosyu_jokenymd,ow_logonuser_no,bosyujoken,confirmkai_sekininsya," & _
                ''                           "confirmkai_ymd,confirmkai_print,szen_jisyafutanumuflg,szen_jisyafutangak,szen_jisyafutanzeikbn," & _
                ''                           "szen_jisyafutanzeigak,szen_bikojisyafutan,szen_sonotafutanumuflg,szen_sonotafutangak,szen_sonotafutanzeikbn," & _
                ''                           "szen_sonotafutanzeigak,szen_sonotafutanname,szen_bikosonotafutan,tatiai_yoteitime,next_keisyo"
                'Dim fldnamegrp As String = "ky_guid,ky_recno,kai_ymd,seisan_ymd,seisan_completeymd," & _
                '                           "uketuke_ymd,riyu,tatiai_ymd,uketuke_logonuser_no,tatiai_logonuser_no," & _
                '                           "seisan_logonuser_no,szen_logonuser_no,szen_jisyano,szen_endymd,szen_kojiyoteistartymd," & _
                '                           "szen_kojiyoteiendymd,szen_kojibasyo,szen_kojigaiyo,next_name,next_namesjis," & _
                '                           "next_postcode,next_addr1,next_addr2,next_tel1,yatin_kozano," & _
                '                           "biko,biko_tatiai,seisan_henkinymd,seisan_sqymd,kanritesu_flg," & _
                '                           "kanritesu_gak,rowid,history,so_yoteiymd,tatiai," & _
                '                           "tatiai_yoteiymd,bosyu_jokenymd,ow_logonuser_no,bosyujoken,szen_jisyafutanumuflg," & _
                '                           "szen_jisyafutangak,szen_jisyafutanzeikbn,szen_jisyafutanzeigak,szen_bikojisyafutan,szen_sonotafutanumuflg," & _
                '                           "szen_sonotafutangak,szen_sonotafutanzeikbn,szen_sonotafutanzeigak,szen_sonotafutanname,szen_bikosonotafutan," & _
                '                           "tatiai_yoteitime,next_keisyo,nkin_kaiyakutukinoprintflg"
                ''20160519 EXEUpdateに伴う修正 契約解約情報 -chg end
                '20160829 革命10バージョンアップに伴う修正 szen_owkojoymd,szen_owsqymd を追加 -add
                Dim fldnamegrp As String = "ky_guid,ky_recno,kai_ymd,seisan_ymd,seisan_completeymd," & _
                                           "uketuke_ymd,riyu,tatiai_ymd,uketuke_logonuser_no,tatiai_logonuser_no," & _
                                           "seisan_logonuser_no,szen_logonuser_no,szen_jisyano,szen_endymd,szen_kojiyoteistartymd," & _
                                           "szen_kojiyoteiendymd,szen_kojibasyo,szen_kojigaiyo,next_name,next_namesjis," & _
                                           "next_postcode,next_addr1,next_addr2,next_tel1,yatin_kozano," & _
                                           "biko,biko_tatiai,seisan_henkinymd,seisan_sqymd,kanritesu_flg," & _
                                           "kanritesu_gak,rowid,history,so_yoteiymd,tatiai," & _
                                           "tatiai_yoteiymd,bosyu_jokenymd,ow_logonuser_no,bosyujoken,tatiai_yoteitime," & _
                                           "next_keisyo,nkin_kaiyakutukinoprintflg,szen_kysno,szen_kysnkbn,szen_kyssorit," & _
                                           "szen_kyssogak,szen_owno,szen_owkaisyukbn,szen_owsosakisoruleguid,szen_owsosakisoruleno," & _
                                           "szen_owsokozano,szen_owsqsimeymd,szen_ownkbn,szen_owyatinkozano,szen_owsqsakino," & _
                                           "szen_owkojoymd,szen_owsqymd"
                '20160620 EXEUpdateに伴う修正2 -chg end

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

                '親マスタ取得
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                'Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)
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
                        .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "契約管理レコードNo"
                                    .Vari_Ky_recno = fldvalue.Trim
                                Case "解約日"
                                    .Vari_Kai_ymd = fldvalue.Trim
                                Case "解約精算費用決定日"
                                    .Vari_Seisan_ymd = fldvalue.Trim
                                Case "解約精算業務完了日"
                                    .Vari_Seisan_completeymd = fldvalue.Trim
                                Case "解約受付日"
                                    .Vari_Uketuke_ymd = fldvalue.Trim
                                Case "解約理由"
                                    .Vari_Riyu = fldvalue.Trim
                                Case "退去日"
                                    .Vari_Tatiai_ymd = fldvalue.Trim
                                Case "受付担当者No"
                                    .Vari_Uketuke_logonuser_no = fldvalue.Trim
                                Case "立会担当者No"
                                    .Vari_Tatiai_logonuser_no = fldvalue.Trim
                                Case "精算担当者No"
                                    .Vari_Seisan_logonuser_no = fldvalue.Trim
                                Case "修繕担当者No"
                                    .Vari_Szen_logonuser_no = fldvalue.Trim
                                Case "自社・支店No"
                                    .Vari_Szen_jisyano = fldvalue.Trim
                                Case "修繕完了日"
                                    .Vari_Szen_endymd = fldvalue.Trim
                                Case "工事予定期間開始日"
                                    .Vari_Szen_kojiyoteistartymd = fldvalue.Trim
                                Case "工事予定期間終了日"
                                    .Vari_Szen_kojiyoteiendymd = fldvalue.Trim
                                Case "工事場所"
                                    .Vari_Szen_kojibasyo = fldvalue.Trim
                                Case "工事概要"
                                    .Vari_Szen_kojigaiyo = fldvalue.Trim
                                Case "退去後宛名"
                                    .Vari_Next_name = fldvalue.Trim
                                Case "退去後宛名Shift-jis"
                                    .Vari_Next_namesjis = fldvalue.Trim
                                Case "退去後郵便番号"
                                    .Vari_Next_postcode = fldvalue.Trim
                                Case "退去後住所①"
                                    .Vari_Next_addr1 = fldvalue.Trim
                                Case "退去後住所②"
                                    .Vari_Next_addr2 = fldvalue.Trim
                                Case "退去後電話番号①"
                                    .Vari_Next_tel1 = fldvalue.Trim
                                Case "不足時入金口座"
                                    .Vari_Yatin_kozano = fldvalue.Trim
                                Case "備考(解約精算)"
                                    .Vari_Biko = fldvalue.Trim
                                Case "備考(立会)"
                                    .Vari_Biko_tatiai = fldvalue.Trim
                                Case "返金予定日"
                                    .Vari_Seisan_henkinymd = fldvalue.Trim
                                Case "請求締め日"
                                    .Vari_Seisan_sqymd = fldvalue.Trim
                                Case "管理手数料フラグ"
                                    .Vari_Kanritesu_flg = fldvalue.Trim
                                Case "管理手数料額"
                                    .Vari_Kanritesu_gak = fldvalue.Trim
                                Case "送金予定日"
                                    .Vari_So_yoteiymd = fldvalue.Trim
                                Case "立会い"
                                    .Vari_Tatiai = fldvalue.Trim
                                Case "退去予定日"
                                    .Vari_Tatiai_yoteiymd = fldvalue.Trim
                                Case "募集条件確認日"
                                    .Vari_Bosyu_jokenymd = fldvalue.Trim
                                Case "家主連絡担当者No"
                                    .Vari_Ow_logonuser_no = fldvalue.Trim
                                Case "募集条件内容"
                                    .Vari_Bosyujoken = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約解約情報 -del sta
                                    'Case "確認事項入力責任者"
                                    '    .Vari_Confirmkai_sekininsya = fldvalue.Trim
                                    'Case "確認事項確認日"
                                    '    .Vari_Confirmkai_ymd = fldvalue.Trim
                                    'Case "印刷時の確認事項"
                                    '    .Vari_Confirmkai_print = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約解約情報 -del end
                                    '20160620 EXEUpdateに伴う修正2 -del sta
                                    'Case "修繕自社負担フラグ"
                                    '    .Vari_Szen_jisyafutanumuflg = fldvalue.Trim
                                    'Case "修繕自社負担額"
                                    '    .Vari_Szen_jisyafutangak = fldvalue.Trim
                                    'Case "修繕自社負担税区分"
                                    '    .Vari_Szen_jisyafutanzeikbn = fldvalue.Trim
                                    'Case "修繕自社負担税額"
                                    '    .Vari_Szen_jisyafutanzeigak = fldvalue.Trim
                                    'Case "修繕自社負担備考"
                                    '    .Vari_Szen_bikojisyafutan = fldvalue.Trim
                                    'Case "修繕その他負担フラグ"
                                    '    .Vari_Szen_sonotafutanumuflg = fldvalue.Trim
                                    'Case "修繕その他負担額"
                                    '    .Vari_Szen_sonotafutangak = fldvalue.Trim
                                    'Case "修繕その他負担税区分"
                                    '    .Vari_Szen_sonotafutanzeikbn = fldvalue.Trim
                                    'Case "修繕その他負担税額"
                                    '    .Vari_Szen_sonotafutanzeigak = fldvalue.Trim
                                    'Case "修繕その他負担者"
                                    '    .Vari_Szen_sonotafutanname = fldvalue.Trim
                                    'Case "修繕その他負担備考"
                                    '    .Vari_Szen_bikosonotafutan = fldvalue.Trim
                                    '20160620 EXEUpdateに伴う修正2 -del end
                                Case "退去予定時刻"
                                    .Vari_Tatiai_yoteitime = fldvalue.Trim
                                Case "退去後敬称"
                                    .Vari_Next_keisyo = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約解約情報 -add sta
                                Case "解約月賃料は扱わないフラグ"
                                    .Vari_Nkin_kaiyakutukinoprintflg = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約解約情報 -add end
                                    '20160620 EXEUpdateに伴う修正2 -add sta
                                Case "契約者修繕請求先No"
                                    .Vari_Szen_kysno = fldvalue
                                Case "契約者修繕入金区分"
                                    .Vari_Szen_kysnkbn = fldvalue
                                Case "契約者修繕送金率"
                                    .Vari_Szen_kyssorit = fldvalue
                                Case "契約者修繕送金額"
                                    .Vari_Szen_kyssogak = fldvalue
                                Case "家主修繕控除先No"
                                    .Vari_Szen_owno = fldvalue
                                Case "家主修繕控除請求区分"
                                    .Vari_Szen_owkaisyukbn = fldvalue
                                Case "家主修繕控除先送金ルールGuid"
                                    .Vari_Szen_owsosakisoruleguid = fldvalue
                                Case "家主修繕控除先送金ルールNo"
                                    .Vari_Szen_owsosakisoruleno = fldvalue
                                Case "家主修繕控除先口座No"
                                    .Vari_Szen_owsokozano = fldvalue
                                Case "家主修繕請求締日"
                                    .Vari_Szen_owsqsimeymd = fldvalue
                                Case "家主修繕請求入金区分"
                                    .Vari_Szen_ownkbn = fldvalue
                                Case "家主修繕請求振込先口座No"
                                    .Vari_Szen_owyatinkozano = fldvalue
                                Case "家主修繕請求先No"
                                    .Vari_Szen_owsqsakino = fldvalue
                                    '20160620 EXEUpdateに伴う修正2 -add end
                                    '20160829 革命10バージョンアップに伴う修正 -add sta
                                Case "家主修繕控除予定日"
                                    .Vari_Szen_owkojoymd = fldvalue
                                Case "家主修繕請求書発行予定日"
                                    .Vari_Szen_owsqymd = fldvalue
                                    '20160829 革命10バージョンアップに伴う修正 -add end
                            End Select

                        Next

                        '固定値
                        .Vari_Rowid = Guid.NewGuid.ToString
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
                            ''hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            ''2016.04.26 メインの方へも反映させる修正 -chg end
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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

                '作業用VIEWの削除
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)

                '契約基本情報の解約フラグを一括更新
                Dim tmpcnt As Integer = 0
                Dim kaiflgupdatesql As String = Me.Get_UseQry_Update()
                DBExec.Exec_NonQuery(sqlcnnv10, kaiflgupdatesql, tmpcnt)
                '20160829 革命10バージョンアップに伴う修正 -add sta
                '控除予定日の一括更新
                Dim kojodayupdatesql As String = Me.Get_UseQry_Kojo()
                DBExec.Exec_NonQuery(sqlcnnv10, kojodayupdatesql, tmpcnt)

                '請求書発行予定日の一括更新
                Dim sqdayupdatesql As String = Me.Get_UseQry_Sq()
                DBExec.Exec_NonQuery(sqlcnnv10, sqdayupdatesql, tmpcnt)
                '20160829 革命10バージョンアップに伴う修正 -add end
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

            ''' <summary>
            ''' 控除予定日の一括更新クエリ 20160829 革命10バージョンアップに伴う修正
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Kojo() As String

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " UPDATE kydata_kai "
                tmp_sql = tmp_sql & " 	SET szen_owkojoymd = LEFT(CONVERT(VARCHAR,DATEADD(MONTH,SO.soymd1_sokintukikbn - 1,GETDATE()),112),6) + CONVERT(VARCHAR,SO.soymd1_sokinsimekbn) "
                tmp_sql = tmp_sql & " FROM kydata_kai AS KYK "
                tmp_sql = tmp_sql & " LEFT JOIN kydata AS KY ON KYK.ky_guid = KY.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN sorule AS SO ON KY.bk_guid = SO.relation_guid "
                tmp_sql = tmp_sql & " WHERE szen_owkaisyukbn = 1 "
                tmp_sql = tmp_sql & " AND   szen_owkojoymd IS NULL "

                Return tmp_sql

            End Function

            ''' <summary>
            ''' 請求書発行予定日の一括更新クエリ 20160829 革命10バージョンアップに伴う修正
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Sq() As String

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " UPDATE kydata_kai "
                tmp_sql = tmp_sql & " 	SET szen_owsqymd = LEFT(CONVERT(VARCHAR,DATEADD(MONTH,SO.soymd1_sokintukikbn - 1,GETDATE()),112),6) + CONVERT(VARCHAR,SO.soymd1_sokinsimekbn) "
                tmp_sql = tmp_sql & " FROM kydata_kai AS KYK "
                tmp_sql = tmp_sql & " LEFT JOIN kydata AS KY ON KYK.ky_guid = KY.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN sorule AS SO ON KY.bk_guid = SO.relation_guid "
                tmp_sql = tmp_sql & " WHERE szen_owkaisyukbn = 2 "
                tmp_sql = tmp_sql & " AND   szen_owsqymd IS NULL "

                Return tmp_sql

            End Function

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) FROM " & tblname
                tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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
            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                'VIEW初期化
                Call Me.Drop_TmpView(sqlcnnv10, viewname)

                'VIEW作成
                Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

                tmp_sql_create = tmp_sql_create & " SELECT "
                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            End Sub

            ''' <summary>
            ''' 作業用VIEWの削除
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="viewname"></param>
            ''' <remarks></remarks>
            Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

                Dim tmpcnt As Integer = 0

                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            End Sub

            ''' <summary>
            ''' [kydata].[kaiyaku_flg]の一括更新クエリ
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Update() As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " UPDATE kydata SET "
                tmp_sql = tmp_sql & " 	kaiyaku_flg = KYK.ky_recno "
                tmp_sql = tmp_sql & " FROM kydata_kai AS KYK "
                tmp_sql = tmp_sql & " LEFT JOIN kydata AS KY ON KYK.ky_guid = KY.ky_guid "

                Return tmp_sql

            End Function

        End Class

    End Class

#End Region

#Region "契約修繕見積情報"

    Public Class Kydata_kaiszen_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
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
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト(一棟所有)
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_kaiszen_Model          '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'guid取得
                Call Me.Set_KyGuid(sqlcnnv10, hash_bkhykyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_kaiszen"
                Dim fldnamegrp As String = "ky_guid,ky_recno,szen_mituno,mitu_title,mitu_bango," & _
                                           "mitu_ymd,seiyaku_flg,seiyaku_ymd,futan_kbn,kys_futanrit," & _
                                           "ow_futanrit,jisya_futanrit,zei_kbn,gokei_zeikbn,gokei_zeirit," & _
                                           "kys_gokeizeigak,ow_gokeizeigak,jisya_gokeizeigak,hasu_futankbn"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)

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

                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)
                    Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)
                    Dim fldname_keysub4 As String = headervalue(startrow - 1, keycol_sub4)

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
                        Dim tmp_keysub4 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub4))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""
                        Dim tmp_kyno As String = ""

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "物件No"
                                    tmp_bkno = fldvalue
                                Case "部屋No"
                                    tmp_hyno = fldvalue
                                Case "契約No"
                                    tmp_kyno = fldvalue
                                Case "更新No"
                                    .Vari_Ky_recno = fldvalue
                                Case "見積No"
                                    .Vari_Szen_mituno = fldvalue
                                Case "見積タイトル"
                                    .Vari_Mitu_title = fldvalue
                                Case "見積番号"
                                    .Vari_Mitu_bango = fldvalue
                                Case "見積日"
                                    .Vari_Mitu_ymd = fldvalue
                                Case "成約フラグ"
                                    .Vari_Seiyaku_flg = fldvalue
                                Case "成約日"
                                    .Vari_Seiyaku_ymd = fldvalue
                                Case "負担区分(全体・個別)"
                                    .Vari_Futan_kbn = fldvalue
                                Case "契約者負担率(全体用)"
                                    .Vari_Kys_futanrit = fldvalue
                                Case "家主負担率(全体用)"
                                    .Vari_Ow_futanrit = fldvalue
                                Case "自社負担率(全体用)"
                                    .Vari_Jisya_futanrit = fldvalue
                                Case "税適用区分(全体・個別)"
                                    .Vari_Zei_kbn = fldvalue
                                Case "税区分(全体用)"
                                    .Vari_Gokei_zeikbn = fldvalue
                                Case "適用税率"
                                    .Vari_Gokei_zeirit = fldvalue
                                Case "契約者税額(全体用)"
                                    .Vari_Kys_gokeizeigak = fldvalue
                                Case "家主税額(全体用)"
                                    .Vari_Ow_gokeizeigak = fldvalue
                                Case "自社税額(全体用)"
                                    .Vari_Jisya_gokeizeigak = fldvalue
                                Case "端数負担者区分"
                                    .Vari_Hasu_futankbn = fldvalue
                            End Select

                        Next

                        'guid取得用に成形
                        .Vari_Ky_guid = tmp_bkno & "-" & tmp_hyno & "-" & tmp_kyno

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20160629 修繕検証後修正 -add sta
                        If skipflg = False Then
                            Call Me.Chk_Futanritu(tblname, hash_cvitem, hash_log)
                        End If
                        '20160629 修繕検証後修正 -add end

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,ky_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,ky_recno) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,szen_mituno) "
                tmp_sql = tmp_sql & " FROM kydata_kaiszen AS KYSZEN "
                tmp_sql = tmp_sql & " LEFT JOIN kydata AS KY ON KYSZEN.ky_guid = KY.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 契約guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_KyGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql = tmp_sql & " 	,ky_guid "
                tmp_sql = tmp_sql & " FROM kydata AS KY "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 負担率の総和チェック '20160629 修繕検証後修正
            ''' </summary>
            ''' <param name="tblname"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_log"></param>
            ''' <remarks></remarks>
            Public Sub Chk_Futanritu(ByVal tblname As String, ByRef hash_cvitem As Hashtable, ByRef hash_log As Hashtable)

                Dim futan_kys As Double = 0
                Dim futan_ow As Double = 0
                Dim futan_jisya As Double = 0
                Dim futan_total As Double = 0

                Dim tmp_futan_kys As String = IIf(hash_cvitem("kys_futanrit") Is Nothing, "0", hash_cvitem("kys_futanrit"))
                Dim tmp_futan_ow As String = IIf(hash_cvitem("ow_futanrit") Is Nothing, "0", hash_cvitem("ow_futanrit"))
                Dim tmp_futan_jisya As String = IIf(hash_cvitem("jisya_futanrit") Is Nothing, "0", hash_cvitem("jisya_futanrit"))

                If tmp_futan_kys.Trim = "" And tmp_futan_ow.Trim = "" And tmp_futan_jisya.Trim = "" Then
                    Exit Sub
                End If

                '※データチェックで数値確認済み
                futan_kys = Double.Parse(tmp_futan_kys)
                futan_ow = Double.Parse(tmp_futan_ow)
                futan_jisya = Double.Parse(tmp_futan_jisya)
                futan_total = futan_kys + futan_ow + futan_jisya

                If futan_total > 100 Then

                    '負担率の総和が100を超えた場合は0に変換して移行する
                    hash_cvitem("kys_futanrit") = "0"
                    hash_cvitem("ow_futanrit") = "0"
                    hash_cvitem("jisya_futanrit") = "0"

                    'ログ出力
                    Dim log_key As String = tblname & "-" & "kys_futanrit"
                    Dim errstr As String = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                    If hash_log.Contains(log_key) Then
                        hash_log(log_key) = errstr
                    Else
                        hash_log.Add(log_key, errstr)
                    End If

                End If

            End Sub

        End Class

    End Class

#End Region

#Region "契約修繕見積詳細情報"

    Public Class Kydata_kaiszenmeisai_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
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
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト(一棟所有)
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Kydata_kaiszenmeisai_Model    '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                Dim keycol_sub5 As Integer = 6                                  'サブ5キー列

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'guid取得
                Call Me.Set_KyGuid(sqlcnnv10, hash_bkhykyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "kydata_kaiszenmeisai"
                Dim fldnamegrp As String = "ky_guid,ky_recno,szen_mituno,szen_meisaino,szen_name," & _
                                           "tekiyo,mitu_suryo,mitu_tani,mitu_tanka,mitu_zeikbn," & _
                                           "mitu_zeigak,kys_futanrit,ow_futanrit,jisya_futanrit,szen_gyno," & _
                                           "jikko_suryo,jikko_tani,jikko_tanka,jikko_zeikbn,jikko_zeigak"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_bkhykyguid, list_basekeydata)

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

                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)
                    Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)
                    Dim fldname_keysub4 As String = headervalue(startrow - 1, keycol_sub4)
                    Dim fldname_keysub5 As String = headervalue(startrow - 1, keycol_sub5)

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
                        Dim tmp_keysub4 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub4))
                        Dim tmp_keysub5 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub5))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3 & "-" & tmp_keysub4 & "-" & tmp_keysub5

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                                                   fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                                                   fldname_keysub4 & " = " & tmp_keysub4 & "、" & _
                                                   fldname_keysub5 & " = " & tmp_keysub5

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""
                        Dim tmp_kyno As String = ""

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "物件No"
                                    tmp_bkno = fldvalue
                                Case "部屋No"
                                    tmp_hyno = fldvalue
                                Case "契約No"
                                    tmp_kyno = fldvalue
                                Case "更新No"
                                    .Vari_Ky_recno = fldvalue
                                Case "見積No"
                                    .Vari_Szen_mituno = fldvalue
                                Case "見積明細No"
                                    .Vari_Szen_meisaino = fldvalue
                                Case "修繕項目名"
                                    .Vari_Szen_name = fldvalue
                                Case "摘要"
                                    .Vari_Tekiyo = fldvalue
                                Case "見積数量"
                                    .Vari_Mitu_suryo = fldvalue
                                Case "見積単位"
                                    .Vari_Mitu_tani = fldvalue
                                Case "見積単価"
                                    .Vari_Mitu_tanka = fldvalue
                                Case "見積税区分"
                                    .Vari_Mitu_zeikbn = fldvalue
                                Case "見積税額(税入力用)"
                                    .Vari_Mitu_zeigak = fldvalue
                                Case "契約者負担率(個別用)"
                                    .Vari_Kys_futanrit = fldvalue
                                Case "家主負担率(個別用)"
                                    .Vari_Ow_futanrit = fldvalue
                                Case "自社負担率(個別用)"
                                    .Vari_Jisya_futanrit = fldvalue
                                Case "発注業者No"
                                    .Vari_Szen_gyno = fldvalue
                                Case "実行数量"
                                    .Vari_Jikko_suryo = fldvalue
                                Case "実行単位"
                                    .Vari_Jikko_tani = fldvalue
                                Case "実行単価"
                                    .Vari_Jikko_tanka = fldvalue
                                Case "実行税区分"
                                    .Vari_Jikko_zeikbn = fldvalue
                                Case "実行税額(税入力用)"
                                    .Vari_Jikko_zeigak = fldvalue
                            End Select

                        Next

                        'guid取得用に成形
                        .Vari_Ky_guid = tmp_bkno & "-" & tmp_hyno & "-" & tmp_kyno

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20160629 修繕検証後修正 -add sta
                        If skipflg = False Then
                            skipflg = Not (Me.Chk_Futanritu(tblname, hash_cvitem, hash_log))
                        End If
                        '20160629 修繕検証後修正 -add end

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                            'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
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
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,ky_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,ky_recno) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,szen_mituno) "
                tmp_sql = tmp_sql & " FROM kydata_kaiszen AS KYSZEN "
                tmp_sql = tmp_sql & " LEFT JOIN kydata AS KY ON KYSZEN.ky_guid = KY.ky_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 契約guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_KyGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                tmp_sql = tmp_sql & " 	,ky_guid "
                tmp_sql = tmp_sql & " FROM kydata AS KY "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 負担率の総和チェック '20160629 修繕検証後修正
            ''' </summary>
            ''' <param name="tblname"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_log"></param>
            ''' <remarks></remarks>
            Public Function Chk_Futanritu(ByVal tblname As String, ByRef hash_cvitem As Hashtable, ByRef hash_log As Hashtable) As Boolean

                Dim rtn As Boolean = True

                Dim futan_kys As Double = 0
                Dim futan_ow As Double = 0
                Dim futan_jisya As Double = 0
                Dim futan_total As Double = 0

                Dim tmp_futan_kys As String = IIf(hash_cvitem("kys_futanrit") Is Nothing, "0", hash_cvitem("kys_futanrit"))
                Dim tmp_futan_ow As String = IIf(hash_cvitem("ow_futanrit") Is Nothing, "0", hash_cvitem("ow_futanrit"))
                Dim tmp_futan_jisya As String = IIf(hash_cvitem("jisya_futanrit") Is Nothing, "0", hash_cvitem("jisya_futanrit"))

                If tmp_futan_kys.Trim = "" And tmp_futan_ow.Trim = "" And tmp_futan_jisya.Trim = "" Then
                    Return rtn
                End If

                '※データチェックで数値確認済み
                futan_kys = Double.Parse(tmp_futan_kys)
                futan_ow = Double.Parse(tmp_futan_ow)
                futan_jisya = Double.Parse(tmp_futan_jisya)
                futan_total = futan_kys + futan_ow + futan_jisya

                If futan_total > 100 Then

                    '負担率の総和が100を超えた場合は0に変換して移行する
                    hash_cvitem("kys_futanrit") = "0"
                    hash_cvitem("ow_futanrit") = "0"
                    hash_cvitem("jisya_futanrit") = "0"

                    'ログ出力
                    Dim log_key As String = tblname & "-" & "kys_futanrit"
                    Dim errstr As String = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                    If hash_log.Contains(log_key) Then
                        hash_log(log_key) = errstr
                    Else
                        hash_log.Add(log_key, errstr)
                    End If

                    rtn = False

                End If

                Return rtn

            End Function

        End Class

    End Class

#End Region





    '20160531 鍵情報移行処理の修正 -del sta
    '#Region "契約鍵情報"

    '    Public Class Kydata_kagi_Repository

    '        Public Class SubConv
    '            Implements IConv

    '            ''' <summary>
    '            ''' 【中間ファイル→変数】
    '            ''' </summary>
    '            ''' <param name="sqlcnnv10"></param>
    '            ''' <param name="syorikomok"></param>
    '            ''' <returns></returns>
    '            ''' <remarks></remarks>
    '            Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

    '                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '                Dim startrow As Integer                                         '書込開始行
    '                Dim columncnt As Integer                                        '列数
    '                Dim maxrowcnt As Integer                                        '既存データの行数
    '                Dim rowcnt As Integer                                           '書込行数
    '                Dim tmp_cvcnt As Integer                                        '移行件数格納
    '                Dim tmp_condcnt As Integer                                      '調整件数格納

    '                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
    '                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
    '                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
    '                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
    '                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
    '                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
    '                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
    '                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル

    '                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
    '                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
    '                Dim rtn As Boolean = True                                       '戻り値

    '                Dim model_cvitem As New Njc.Model.Hydata_kagi_Model             '移行値格納用モデル初期化
    '                Dim keycol_main As Integer = 1                                  'メインキー列
    '                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
    '                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
    '                Dim viewname As String = "物件部屋キー情報"

    '                '************************
    '                '作業準備
    '                '************************

    '                '10鍵タイトルマスタから鍵Noとタイトル名を紐付けたデータを取得
    '                Call EtcMethod.Set_KagiTitleMst(sqlcnnv10, 1)
    '                Call EtcMethod.Set_KagiTitleMst(sqlcnnv10, 2)

    '                'Excelファイル初期設定
    '                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '                'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '                If Not rtn Then
    '                    Return rtn
    '                End If

    '                'キー取得用のVIEWを作成
    '                Call Me.Create_View_KeyAndGuid(sqlcnnv10)

    '                'テーブル名/フィールド名セット
    '                Dim viewname_base As String = PRE_VIEW_NAME & viewname
    '                Dim tblname As String = "hydata_kagi"
    '                Dim fldnamegrp As String = "hy_guid,kagi_no,honsu,biko,history," & _
    '                                           "hokan,gyshare"

    '                '追加コンバート時の重複チェック用に既存データのキーを取得
    '                If Not InitDBFlg Then
    '                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
    '                End If

    '                'guidを取得
    '                Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

    '                '親マスタ取得
    '                '2016.04.26 メインの方へも反映させる修正 -chg sta
    '                'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
    '                Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
    '                '2016.04.26 メインの方へも反映させる修正 -chg end

    '                'プログレスバー初期化
    '                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
    '                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

    '                If rowcnt <= pgbbasecnt Then
    '                    pgbtotalcnt = rowcnt
    '                Else
    '                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
    '                End If
    '                Call obj_pgb.pgbInitPart(pgbtotalcnt)

    '                '************************
    '                '処理開始
    '                '************************
    '                With model_cvitem

    '                    '---------------
    '                    'ヘッダー処理
    '                    '---------------
    '                    'ヘッダー行取得
    '                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '                    'キーヘッダー名取得
    '                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
    '                    Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
    '                    Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

    '                    '---------------
    '                    'データ部処理
    '                    '---------------
    '                    For cntii = 1 To rowcnt

    '                        '中断処理
    '                        Application.DoEvents()
    '                        If CancelFlg Then
    '                            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                            Return rtn
    '                        End If

    '                        '行取得
    '                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

    '                        'キー値取得
    '                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
    '                        Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
    '                        Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
    '                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
    '                        .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

    '                        'ログ出力用
    '                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

    '                        '移行値取得
    '                        For cntjj = 1 To columncnt

    '                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                            Dim fldvalue As String = ""
    '                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
    '                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
    '                            End If

    '                            Select Case fldname
    '                                Case "鍵タイトル名"
    '                                    .Vari_Kagi_no = fldvalue.Trim
    '                                Case "鍵本数"
    '                                    .Vari_Honsu = fldvalue.Trim
    '                                Case "備考"
    '                                    .Vari_Biko = fldvalue.Trim
    '                                Case "保管場所"
    '                                    .Vari_Hokan = fldvalue.Trim
    '                                Case "業者間での情報共有"
    '                                    .Vari_Gyshare = fldvalue.Trim
    '                            End Select

    '                        Next

    '                        '固定値
    '                        .Vari_History = DefHistory

    '                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
    '                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

    '                        'データチェック
    '                        Dim skipflg As Boolean = False
    '                        Dim hash_cvitem As New Hashtable
    '                        Dim hash_log As New Hashtable
    '                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

    '                        '書込処理
    '                        If Not skipflg Then

    '                            'キーをguidへ変換
    '                            hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))

    '                            '挿入処理
    '                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

    '                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
    '                            If normalflg Then
    '                                list_chkduplicate.Add(fldvalue_key)
    '                                tmp_cvcnt = tmp_cvcnt + 1
    '                            End If

    '                        End If

    '                        '-------------------
    '                        'ログ出力
    '                        '-------------------

    '                        'ログ出力メッセージ整形
    '                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

    '                        'ログ出力
    '                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
    '                            Dim tmp_cnt As Integer = 0
    '                            '挿入
    '                            For Each logvalue In sortlist_log
    '                                Dim tmp_sql_insert As String = logvalue.Value
    '                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
    '                            Next
    '                            '初期化
    '                            tmp_logcnt = 0
    '                            sortlist_log.Clear()
    '                        End If

    '                        '-------------------------------
    '                        'プログレスバー更新/進捗率表示
    '                        '-------------------------------
    '                        '件数取得
    '                        Dim pgbcnt As Integer = 0
    '                        If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
    '                            pgbcnt = cntii
    '                        ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
    '                            Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
    '                        ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
    '                            pgbcnt = pgbtotalcnt
    '                        End If

    '                        '表示
    '                        If pgbcnt <> 0 Then
    '                            Call obj_pgb.pgbsettingPart(pgbcnt)
    '                            Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
    '                        End If

    '                    Next

    '                End With

    '                '************************
    '                '終了処理
    '                '************************

    '                '中間ファイル件数を取得
    '                midrowcnt = rowcnt

    '                '移行件数を取得
    '                cvrowcnt = tmp_cvcnt

    '                '調整件数を取得
    '                conditioncnt = tmp_condcnt

    '                'Excelファイル終了設定
    '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '                '返却
    '                Return rtn

    '            End Function

    '            ''' <summary>
    '            ''' 親マスタ取得→リスト格納
    '            ''' </summary>
    '            ''' <param name="sqlcnnv10"></param>
    '            ''' <param name="tblname"></param>
    '            ''' <param name="list_basedata"></param>
    '            ''' <remarks></remarks>
    '            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

    '                Dim tmp_sql As String = " SELECT キー FROM " & tblname
    '                Dim flg As Boolean = True

    '                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

    '            End Sub

    '            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

    '            End Function

    '            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

    '            End Function

    '            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

    '                Dim tmp_sql As String = ""
    '                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,kagi_no) FROM " & tblname
    '                tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
    '                tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
    '                Dim flg As Boolean = True

    '                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

    '            End Sub

    '            ''' <summary>
    '            ''' guid取得→ハッシュテーブル格納
    '            ''' </summary>
    '            ''' <param name="sqlcnnv10"></param>
    '            ''' <param name="tblname"></param>
    '            ''' <param name="hash_guid"></param>
    '            ''' <remarks></remarks>
    '            Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

    '                Dim tmp_sql As String = "SELECT * FROM " & tblname
    '                Dim flg As Boolean = True

    '                '2016.04.26 メインの方へも反映させる修正 -chg sta
    '                'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
    '                Dim tmp_hash As New Hashtable
    '                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

    '                '取得した部屋Noを半角変換して照合用に統一するrtn_hash
    '                hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
    '                '2016.04.26 メインの方へも反映させる修正 -chg end

    '            End Sub

    '            ''' <summary>
    '            ''' キー/guidをセットにしたVIEWの作成
    '            ''' </summary>
    '            ''' <param name="sqlcnnv10"></param>
    '            ''' <remarks></remarks>
    '            Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

    '                Dim viewname As String = "物件部屋キー情報"
    '                Dim tmpcnt As Integer = 0

    '                'VIEW初期化
    '                Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
    '                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

    '                'VIEW作成
    '                Dim tmp_sql_create As String = ""
    '                tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
    '                tmp_sql_create = tmp_sql_create & " SELECT "
    '                tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
    '                tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
    '                tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
    '                tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
    '                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

    '            End Sub

    '        End Class

    '    End Class

    '#End Region
    '20160531 鍵情報移行処理の修正 -del end

End Namespace


