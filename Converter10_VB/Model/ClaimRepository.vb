Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "クレーム基本情報"

    Public Class Claimdata_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

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

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Claimdata_Model               '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列
                Dim hash_bkguid As New Hashtable                                'キー/guid格納用ハッシュテーブル
                Dim hash_bkhyguid As New Hashtable                              'キー/guid格納用ハッシュテーブル
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim viewnamebktobkguid As String = PRE_VIEW_NAME & "物件キー情報"
                Dim viewnamebkhytohyguid As String = PRE_VIEW_NAME & "物件部屋キー情報"
                Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                '************************
                '作業準備
                '************************

                '10クレーム分類マスタから分類Noと分類名を紐付けたデータを取得
                Call EtcMethod.Set_ClaimBruiMst(sqlcnnv10, 1)
                Call EtcMethod.Set_ClaimBruiMst(sqlcnnv10, 2)

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'キー取得用のVIEWを作成
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebktobkguid)
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyguid)
                Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "claimdata"
                Dim fldnamegrp As String = "claim_no,status,title,uke_ymd,uke_logonuser_no," & _
                                           "uke_name,uke_tel1,uke_tel2,renraku_timestart,renraku_timeend," & _
                                           "kinkyu_kbn,taio_limitymd,kasyo_ruino,claim_ruino,uke_report," & _
                                           "taio_logonuser_no,emailbiko,bk_guid,hy_guid,ky_guid," & _
                                           "ky_recno,reform_guid,taiosaki_kbn,taiosaki_gy_no,taiosaki_tantoname," & _
                                           "taiosaki_tel,taio_finishymd,taio_report,hutan1_kbn,hutan1_gaketc," & _
                                           "hutan2_kbn,hutan2_gaketc,hutan3_kbn,hutan3_gaketc,hutanbiko," & _
                                           "jisya_no,jisya_name,history,rowid,uke_hm," & _
                                           "taio_finishhm,uke_namesjis,taiosaki_tantonamesjis,ow_no"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, viewnamebktobkguid, hash_bkguid)
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhytohyguid, hash_bkhyguid)
                Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

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
                    Dim fldname_key As String = headervalue(startrow - 1, keycol)

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
                        Dim fldvalue_key As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))

                        'ログ出力用
                        Dim str_logkey As String = fldname_key & " = " & fldvalue_key

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
                                Case "クレームNo"
                                    .Vari_Claim_no = fldvalue
                                Case "対応状況"
                                    .Vari_Status = fldvalue
                                Case "タイトル"
                                    .Vari_Title = fldvalue
                                Case "受付日時"
                                    .Vari_Uke_ymd = fldvalue
                                Case "受付担当者"
                                    .Vari_Uke_logonuser_no = fldvalue
                                Case "連絡者名"
                                    .Vari_Uke_name = fldvalue
                                Case "連絡者TEL1"
                                    .Vari_Uke_tel1 = fldvalue
                                Case "連絡者TEL2"
                                    .Vari_Uke_tel2 = fldvalue
                                Case "連絡可能時間（開始）"
                                    .Vari_Renraku_timestart = fldvalue
                                Case "連絡可能時間（終了）"
                                    .Vari_Renraku_timeend = fldvalue
                                Case "緊急度"
                                    .Vari_Kinkyu_kbn = fldvalue
                                Case "対応期限"
                                    .Vari_Taio_limitymd = fldvalue
                                Case "箇所分類No"
                                    .Vari_Kasyo_ruino = fldvalue
                                Case "クレーム分類No"
                                    .Vari_Claim_ruino = fldvalue
                                Case "内容"
                                    .Vari_Uke_report = fldvalue
                                Case "対応担当者"
                                    .Vari_Taio_logonuser_no = fldvalue
                                Case "電子メール備考"
                                    .Vari_Emailbiko = fldvalue
                                Case "物件No"
                                    '.Vari_Bk_guid = fldvalue
                                    tmp_bkno = fldvalue
                                Case "部屋No"
                                    '.Vari_Hy_guid = fldvalue
                                    tmp_hyno = fldvalue
                                Case "契約No"
                                    '.Vari_Ky_guid = fldvalue
                                    tmp_kyno = fldvalue
                                Case "契約レコードNo"
                                    .Vari_Ky_recno = fldvalue
                                Case "対応先区分"
                                    .Vari_Taiosaki_kbn = fldvalue
                                Case "依頼業者No"
                                    .Vari_Taiosaki_gy_no = fldvalue
                                Case "業者担当者"
                                    .Vari_Taiosaki_tantoname = fldvalue
                                Case "業者担当者TEL"
                                    .Vari_Taiosaki_tel = fldvalue
                                Case "完了日時"
                                    .Vari_Taio_finishymd = fldvalue
                                Case "結果入力"
                                    .Vari_Taio_report = fldvalue
                                Case "負担者1"
                                    .Vari_Hutan1_kbn = fldvalue
                                Case "負担者1金額等"
                                    .Vari_Hutan1_gaketc = fldvalue
                                Case "負担者2"
                                    .Vari_Hutan2_kbn = fldvalue
                                Case "負担者2金額等"
                                    .Vari_Hutan2_gaketc = fldvalue
                                Case "負担者3"
                                    .Vari_Hutan3_kbn = fldvalue
                                Case "負担者3金額等"
                                    .Vari_Hutan3_gaketc = fldvalue
                                Case "備考"
                                    .Vari_Hutanbiko = fldvalue
                                Case "自社・支店No"
                                    .Vari_Jisya_no = fldvalue
                                Case "自社・支店名"
                                    .Vari_Jisya_name = fldvalue
                                Case "受付時刻"
                                    .Vari_Uke_hm = fldvalue
                                Case "対応終了時刻"
                                    .Vari_Taio_finishhm = fldvalue
                                Case "連絡者名SJIS"
                                    .Vari_Uke_namesjis = fldvalue
                                Case "対応担当者名SJIS"
                                    .Vari_Taiosaki_tantonamesjis = fldvalue
                                Case "家主No"
                                    .Vari_Ow_no = fldvalue
                            End Select

                        Next

                        'guid取得用に成形
                        .Vari_Bk_guid = tmp_bkno
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno
                        .Vari_Ky_guid = tmp_bkno & "-" & tmp_hyno & "-" & tmp_kyno

                        '固定値
                        .Vari_Reform_guid = "00000000-0000-0000-0000-000000000000"  '修繕関連は構築中のため現時点でダミーを設定しておく
                        .Vari_History = DefHistory
                        .Vari_Rowid = Guid.NewGuid.ToString

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
                            hash_cvitem("bk_guid") = hash_bkguid.Item(hash_cvitem.Item("bk_guid"))
                            hash_cvitem("hy_guid") = hash_bkhyguid.Item(hash_cvitem.Item("hy_guid"))
                            hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))

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
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebktobkguid)
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid)
                Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

                '返却
                Return rtn

            End Function

            Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            ''' <summary>
            ''' 既存データのキーを取得してリストへ格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT claim_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                    Case "物件キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT bk_no AS キー,bk_guid FROM bkdata "
                    Case "物件部屋キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT "
                        tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 	,hy_guid "
                        tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                        tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                    Case "物件部屋契約キー情報"
                        tmp_sql_create = tmp_sql_create & " SELECT * FROM "
                        tmp_sql_create = tmp_sql_create & " ( "
                        tmp_sql_create = tmp_sql_create & " 	SELECT "
                        tmp_sql_create = tmp_sql_create & " 		 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                        tmp_sql_create = tmp_sql_create & " 		,ky_guid "
                        tmp_sql_create = tmp_sql_create & " 	FROM kydata AS KY "
                        tmp_sql_create = tmp_sql_create & " 	LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                        tmp_sql_create = tmp_sql_create & " 	LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "
                        tmp_sql_create = tmp_sql_create & " ) AS VW "
                        tmp_sql_create = tmp_sql_create & " WHERE [キー] IS NOT NULL "
                End Select

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

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

#Region "クレーム対応履歴情報"

    Public Class Claim_taio_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

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

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Claim_taio_Model              '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'テーブル名/フィールド名セット
                Dim tblname_base As String = "claimdata"
                Dim tblname As String = "claim_taio"
                Dim fldnamegrp As String = "claim_no,taio_no,taio_ymd,taio_logonuser_no,taio_report," & _
                                           "history,taio_hm"

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

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
                    Dim fldname_keysub As String = headervalue(startrow - 1, keycol_sub)

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
                        Dim fldvalue_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim fldvalue_keysub As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub))
                        Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "クレームNo"
                                    .Vari_Claim_no = fldvalue
                                Case "対応履歴No"
                                    .Vari_Taio_no = fldvalue
                                Case "対応日時"
                                    .Vari_Taio_ymd = fldvalue
                                Case "対応担当者"
                                    .Vari_Taio_logonuser_no = fldvalue
                                Case "内容"
                                    .Vari_Taio_report = fldvalue
                                Case "対応時刻"
                                    .Vari_Taio_hm = fldvalue
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
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT claim_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            ''' <summary>
            ''' 既存データのキーを取得してリストへ格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,claim_no) + '-' + CONVERT(varchar,taio_no) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "クレーム関連ファイル情報"

    Public Class Claimdata_relfile_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="syorikomok"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

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

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Claimdata_relfile_Model       '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列

                '************************
                '作業準備
                '************************

                '20160525 クレーム関連ファイルの画像判別処理実装 -add sta
                '画像ファイルの拡張子リストを作成
                Call EtcMethod.Set_ImportableImageFileAttributesList()
                '20160525 クレーム関連ファイルの画像判別処理実装 -add end

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'テーブル名/フィールド名セット
                Dim tblname_base As String = "claimdata"
                Dim tblname As String = "claimdata_relfile"
                Dim fldnamegrp As String = "claim_no,file_no,fullpath,addtime,biko," & _
                                           "history"

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

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
                    Dim fldname_keysub As String = headervalue(startrow - 1, keycol_sub)

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
                        Dim fldvalue_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim fldvalue_keysub As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub))
                        Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                case "クレームNo"	
                                    .Vari_Claim_no = fldvalue
                                Case "ファイルNo"
                                    .Vari_File_no = fldvalue
                                Case "ファイルパス"
                                    .Vari_Fullpath = fldvalue
                                Case "追加日"
                                    .Vari_Addtime = fldvalue
                                Case "備考"
                                    .Vari_Biko = fldvalue
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
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT claim_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            ''' <summary>
            ''' 既存データのキーを取得してリストへ格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,claim_no) + '-' + CONVERT(varchar,file_no) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

End Namespace


