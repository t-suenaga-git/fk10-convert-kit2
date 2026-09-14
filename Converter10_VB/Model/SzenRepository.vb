Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "修繕基本情報"

    Public Class Szendata_Repository

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
                Dim hash_bkguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim hash_bkhyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim hash_bkhykyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Szendata_Model                '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                'Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                '************************
                '作業準備
                '************************

                '入金区分紐付情報取得
                Call SetRelItemToObject.Set_RelData_Nkinkbn()

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'guid取得
                Call Me.Set_BkGuid(sqlcnnv10, hash_bkguid)
                Call Me.Set_HyGuid(sqlcnnv10, hash_bkhyguid)
                Call Me.Set_KyGuid(sqlcnnv10, hash_bkhykyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "szendata"
                Dim fldnamegrp As String = "bk_guid,szen_no,hy_useflg,hy_guid,ky_useflg," & _
                                           "ky_guid,szen_name,szen_ukeymd,szen_endymd,tanto_no," & _
                                           "jisya_no,koji_yoteistartymd,koji_yoteiendymd,koji_basyo,koji_gaiyo," & _
                                           "koji_kagi,koji_tatiai,biko,kys_futanrit,ow_futanrit," & _
                                           "rowid,history,jisya_futanrit,kys_sqflg,kys_no," & _
                                           "kys_gtym,kys_sqsimeymd,kys_nkbn,kys_sorit,kys_yatinkozano," & _
                                           "kys_biko,ow_sqflg,ow_no,ow_kaisyukbn,ow_kojoymd," & _
                                           "ow_kojogtym,ow_sosaki_soruleguid,ow_sosaki_sorule_no,ow_sokozano,ow_sqymd," & _
                                           "ow_sqgtym,ow_sqsqsimeymd,ow_sqnkbn,ow_sqyatinkozano,ow_sqsakino," & _
                                           "ow_biko"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_bkguid, list_basekeydata)

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
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1

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
                                    .Vari_Bk_guid = fldvalue
                                    tmp_bkno = fldvalue
                                Case "修繕No"
                                    .Vari_Szen_no = fldvalue
                                Case "部屋使用フラグ"
                                    .Vari_Hy_useflg = fldvalue
                                Case "部屋No"
                                    tmp_hyno = fldvalue
                                Case "契約使用フラグ"
                                    .Vari_Ky_useflg = fldvalue
                                Case "契約No"
                                    tmp_kyno = fldvalue
                                Case "修繕名"
                                    .Vari_Szen_name = fldvalue
                                Case "修繕受付日"
                                    .Vari_Szen_ukeymd = fldvalue
                                Case "修繕終了日"
                                    .Vari_Szen_endymd = fldvalue
                                Case "修繕担当者"
                                    .Vari_Tanto_no = fldvalue
                                Case "自社支店No"
                                    .Vari_Jisya_no = fldvalue
                                Case "工事開始予定日"
                                    .Vari_Koji_yoteistartymd = fldvalue
                                Case "工事終了予定日"
                                    .Vari_Koji_yoteiendymd = fldvalue
                                Case "工事場所"
                                    .Vari_Koji_basyo = fldvalue
                                Case "工事概要"
                                    .Vari_Koji_gaiyo = fldvalue
                                Case "鍵"
                                    .Vari_Koji_kagi = fldvalue
                                Case "立会者"
                                    .Vari_Koji_tatiai = fldvalue
                                Case "備考"
                                    .Vari_Biko = fldvalue
                                Case "契約者負担率"
                                    .Vari_Kys_futanrit = fldvalue
                                Case "家主負担率"
                                    .Vari_Ow_futanrit = fldvalue
                                Case "自社負担率"
                                    .Vari_Jisya_futanrit = fldvalue
                                Case "契約者請求作成フラグ"
                                    .Vari_Kys_sqflg = fldvalue
                                Case "契約者請求先No"
                                    .Vari_Kys_no = fldvalue
                                Case "契約者該当年月"
                                    .Vari_Kys_gtym = fldvalue
                                Case "契約者請求締日"
                                    .Vari_Kys_sqsimeymd = fldvalue
                                Case "契約者入金区分"
                                    .Vari_Kys_nkbn = fldvalue
                                Case "契約者送金率"
                                    .Vari_Kys_sorit = fldvalue
                                Case "契約者振込先口座No"
                                    .Vari_Kys_yatinkozano = fldvalue
                                Case "契約者請求備考"
                                    .Vari_Kys_biko = fldvalue
                                Case "家主請求作成フラグ"
                                    .Vari_Ow_sqflg = fldvalue
                                Case "家主控除先No"
                                    .Vari_Ow_no = fldvalue
                                Case "家主控除請求区分"
                                    .Vari_Ow_kaisyukbn = fldvalue
                                Case "家主控除予定日"
                                    .Vari_Ow_kojoymd = fldvalue
                                Case "家主控除該当年月"
                                    .Vari_Ow_kojogtym = fldvalue
                                Case "家主控除先送金ルールNo"
                                    .Vari_Ow_sosaki_sorule_no = fldvalue
                                Case "家主控除先口座No"
                                    .Vari_Ow_sokozano = fldvalue
                                Case "家主請求書発行予定日"
                                    .Vari_Ow_sqymd = fldvalue
                                Case "家主請求該当年月"
                                    .Vari_Ow_sqgtym = fldvalue
                                Case "家主請求締日"
                                    .Vari_Ow_sqsqsimeymd = fldvalue
                                Case "家主請求入金区分"
                                    .Vari_Ow_sqnkbn = fldvalue
                                Case "家主請求振込先口座No"
                                    .Vari_Ow_sqyatinkozano = fldvalue
                                Case "家主請求先No"
                                    .Vari_Ow_sqsakino = fldvalue
                                Case "家主控除請求備考"
                                    .Vari_Ow_biko = fldvalue
                            End Select

                        Next

                        'guid取得用に成形
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno
                        .Vari_Ky_guid = tmp_bkno & "-" & tmp_hyno & "-" & tmp_kyno

                        '固定値
                        .Vari_Ow_sosaki_soruleguid = ""
                        .Vari_Rowid = Guid.NewGuid.ToString
                        .Vari_History = DefHistory

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
                            hash_cvitem("bk_guid") = hash_bkguid.Item(StrConv(hash_cvitem.Item("bk_guid"), VbStrConv.Narrow))
                            If tmp_hyno <> "" Then
                                hash_cvitem("hy_guid") = hash_bkhyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                            Else
                                hash_cvitem("hy_guid") = ""
                            End If
                            If tmp_kyno <> "" Then
                                hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                            Else
                                hash_cvitem("ky_guid") = ""
                            End If
                            
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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,szen_no) FROM szendata AS SZ "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            '20160621 del sta
            ' ''' <summary>
            ' ''' guid取得→ハッシュテーブル格納
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="tblname"></param>
            ' ''' <param name="hash_guid"></param>
            ' ''' <remarks></remarks>
            'Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As Hashtable)

            '    Dim tmp_sql As String = "SELECT * FROM " & tblname
            '    Dim flg As Boolean = True

            '    '2016.04.26 メインの方へも反映させる修正 -chg sta
            '    'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
            '    Dim tmp_hash As New Hashtable
            '    flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

            '    '取得した部屋Noを半角変換して照合用に統一するrtn_hash
            '    hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
            '    '2016.04.26 メインの方へも反映させる修正 -chg end

            'End Sub

            ' ''' <summary>
            ' ''' キー/guidをセットにしたVIEWの作成
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <remarks></remarks>
            'Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

            '    Dim tmpcnt As Integer = 0

            '    'VIEW初期化
            '    Call Me.Drop_TmpView(sqlcnnv10, viewname)

            '    'VIEW作成
            '    Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

            '    tmp_sql_create = tmp_sql_create & " SELECT "
            '    tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
            '    tmp_sql_create = tmp_sql_create & " 	,ky_guid "
            '    tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
            '    tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
            '    tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

            '    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

            'End Sub

            ' ''' <summary>
            ' ''' 作業用VIEWの削除
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="viewname"></param>
            ' ''' <remarks></remarks>
            'Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

            '    Dim tmpcnt As Integer = 0

            '    Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
            '    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

            'End Sub
            '20160621 del end

            ''' <summary>
            ''' 物件guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_BkGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = "SELECT bk_no,bk_guid FROM bkdata "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 部屋guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_HyGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) "
                tmp_sql = tmp_sql & " 	,hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

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

#Region "修繕見積情報"

    Public Class Szendata_szen_Repository

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
                Dim hash_bkguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Szendata_szen_Model           '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列

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
                Call Me.Set_BkGuid(sqlcnnv10, hash_bkguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "szendata_szen"
                Dim fldnamegrp As String = "bk_guid,szen_no,szen_mituno,mitu_title,mitu_bango," & _
                                           "mitu_ymd,seiyaku_flg,seiyaku_ymd,futan_kbn,kys_futanrit," & _
                                           "ow_futanrit,jisya_futanrit,zei_kbn,gokei_zeikbn,gokei_zeirit," & _
                                           "kys_gokeizeigak,ow_gokeizeigak,jisya_gokeizeigak,hasu_futankbn"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_bkguid, list_basekeydata)

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

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                case "物件No"	
                                    .Vari_Bk_guid = fldvalue
                                Case "修繕No"
                                    .Vari_Szen_no = fldvalue
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
                            hash_cvitem("bk_guid") = hash_bkguid.Item(StrConv(hash_cvitem.Item("bk_guid"), VbStrConv.Narrow))

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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,szen_no) FROM szendata_szen AS SZ "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 物件guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_BkGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = "SELECT bk_no,bk_guid FROM bkdata "
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

#Region "修繕見積詳細情報"

    Public Class Szendata_szenmeisai_Repository

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
                Dim hash_bkguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Szendata_szenmeisai_Model     '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                Dim keycol_sub3 As Integer = 4                                  'サブ3キー列

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
                Call Me.Set_BkGuid(sqlcnnv10, hash_bkguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "szendata_szenmeisai"
                Dim fldnamegrp As String = "bk_guid,szen_no,szen_mituno,szen_meisaino,szen_name," & _
                                           "tekiyo,mitu_suryo,mitu_tani,mitu_tanka,mitu_zeikbn," & _
                                           "mitu_zeigak,kys_futanrit,ow_futanrit,jisya_futanrit,szen_gyno," & _
                                           "jikko_suryo,jikko_tani,jikko_tanka,jikko_zeikbn,jikko_zeigak"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_bkguid, list_basekeydata)

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
                                Case "物件No"
                                    .Vari_Bk_guid = fldvalue
                                Case "修繕No"
                                    .Vari_Szen_no = fldvalue
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
                            hash_cvitem("bk_guid") = hash_bkguid.Item(StrConv(hash_cvitem.Item("bk_guid"), VbStrConv.Narrow))

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
                tmp_sql = tmp_sql & " 	CONVERT(varchar,bk_no) + '-' + "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,szen_no) + '-' + "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,szen_mituno) + '-' + "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,szen_meisaino) "
                tmp_sql = tmp_sql & " FROM szendata_szenmeisai AS SZ "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 物件guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_BkGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = "SELECT bk_no,bk_guid FROM bkdata "
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

#Region "修繕クレーム関連付け情報"

    Public Class Szendata_claim_Repository

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
                Dim hash_bkguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Szendata_claim_Model          '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列

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
                Call Me.Set_BkGuid(sqlcnnv10, hash_bkguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "szendata_claim"
                Dim fldnamegrp As String = "bk_guid,szen_no,claim_no,claimdata_sortorder,szendata_sortorder," & _
                                           "history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_bkguid, list_basekeydata)

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

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "物件No"
                                    .Vari_Bk_guid = fldvalue
                                Case "修繕No"
                                    .Vari_Szen_no = fldvalue
                                Case "クレームNo"
                                    .Vari_Claim_no = fldvalue
                                Case "クレーム画面での並び順"
                                    .Vari_Claimdata_sortorder = fldvalue
                                Case "修繕情報登録画面での並び順"
                                    .Vari_Szendata_sortorder = fldvalue
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
                            hash_cvitem("bk_guid") = hash_bkguid.Item(StrConv(hash_cvitem.Item("bk_guid"), VbStrConv.Narrow))

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
                tmp_sql = tmp_sql & " 	CONVERT(varchar,bk_no) + '-' + "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,szen_no) + '-' + "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,claim_no) "
                tmp_sql = tmp_sql & " FROM szendata_claim AS SZ "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 物件guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_BkGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = "SELECT bk_no,bk_guid FROM bkdata "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

        End Class

    End Class

#End Region

#Region "修繕関連ファイル情報"

    Public Class Szendata_relfile_Repository

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
                Dim hash_bkguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Szendata_relfile_Model        '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                  'サブ2キー列

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
                Call Me.Set_BkGuid(sqlcnnv10, hash_bkguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "szendata_relfile"
                Dim fldnamegrp As String = "bk_guid,szen_no,file_no,fullpath,addtime," & _
                                           "biko,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_bkguid, list_basekeydata)

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

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                                                   fldname_keysub2 & " = " & tmp_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                case "物件No"	
                                    .Vari_Bk_guid = fldvalue
                                Case "修繕No"
                                    .Vari_Szen_no = fldvalue
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

                            'キーをguidへ変換
                            hash_cvitem("bk_guid") = hash_bkguid.Item(StrConv(hash_cvitem.Item("bk_guid"), VbStrConv.Narrow))

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
                tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,szen_no) FROM szendata_szen AS SZ "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 物件guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_BkGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = "SELECT bk_no,bk_guid FROM bkdata "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

        End Class

    End Class

#End Region

#Region "修繕メモ情報"

    Public Class Szendata_memo_Repository

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
                Dim hash_bkguid As New Hashtable                            'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Szendata_memo_Model           '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列
                Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

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
                Call Me.Set_BkGuid(sqlcnnv10, hash_bkguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "szendata_memo"
                Dim fldnamegrp As String = "bk_guid,szen_no,memo_no,memo,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_bkguid, list_basekeydata)

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
                        Dim tmp_keysub As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub))
                        .Vari_Bk_guid = tmp_keymain
                        .Vari_Szen_no = tmp_keysub

                        '備考カウント初期化
                        Dim cvitemcnt As Integer = 0

                        '移行値取得
                        For cntjj = 3 To columncnt

                            'サブ2キー取得
                            Dim tmp_keysub2 As String = (cntjj - 2).ToString

                            '全キー取得
                            Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub & "-" & tmp_keysub2

                            'ログ出力用データ格納(サブキーフィールド)
                            Dim fldname_keysub2 As String = headervalue(startrow - 1, cntjj)
                            Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub & " = " & tmp_keysub & "、" & fldname_keysub2

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If
                            .Vari_Memo = fldvalue

                            '固定値
                            .Vari_Memo_no = (cntjj - 2).ToString
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
                                    hash_cvitem("bk_guid") = hash_bkguid.Item(hash_cvitem.Item("bk_guid"))

                                    '挿入処理
                                    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                    If normalflg Then
                                        list_chkduplicate.Add(fldvalue_key)
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

                        '--------------------------------------------------------------------
                        '移行した備考が1データ以上ある場合移行したレコードの数を更新する
                        '--------------------------------------------------------------------
                        '20160928 メモ関連の移行件数表示修正 -chg sta
                        'tmp_cvcnt = tmp_cvcnt + 1
                        If cvitemcnt > 0 Then
                            tmp_cvcnt = tmp_cvcnt + 1
                        End If
                        '20160928 メモ関連の移行件数表示修正 -chg end
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
                '20160928 メモ関連の移行件数表示修正 -chg sta
                'midrowcnt = rowcnt
                midrowcnt = tmp_cvrowcnt
                '20160928 メモ関連の移行件数表示修正 -chg end

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
                tmp_sql = tmp_sql & " 	CONVERT(varchar,bk_no) + '-' + "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,szen_no) + '-' + "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,memo_no) "
                tmp_sql = tmp_sql & " FROM szendata_memo AS SZ "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 物件guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_BkGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = "SELECT bk_no,bk_guid FROM bkdata "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

        End Class

    End Class

#End Region

End Namespace


