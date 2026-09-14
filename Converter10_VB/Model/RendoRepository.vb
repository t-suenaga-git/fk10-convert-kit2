Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "送信設定基本情報"

    Public Class M_sendsetting_Repository

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

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.M_sendsetting_Model           '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列

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
                Dim tblname As String = "m_sendsetting"
                Dim fldnamegrp As String = "setting_guid,setting_sortorder,setting_name,keisai_siten,datalink_id," & _
                                           "adconfirm_limitday,history,basedate,basedate_flg,emailaddr1," & _
                                           "emailaddr2,emailaddr3,emailaddr4,emailaddr5"

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
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol)

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
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))
                        Dim fldvalue_key As String = tmp_keymain

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "送信設定順"
                                    .Vari_Setting_sortorder = fldvalue
                                Case "送信設定名"
                                    .Vari_Setting_name = fldvalue
                                Case "掲載支店"
                                    .Vari_Keisai_siten = fldvalue
                                Case "連動ID"
                                    .Vari_Datalink_id = fldvalue
                                Case "広告確認からの確認期間"
                                    .Vari_Adconfirm_limitday = fldvalue
                                Case "基準日"
                                    .Vari_Basedate = fldvalue
                                Case "基準日区分"
                                    .Vari_Basedate_flg = fldvalue
                                Case "Eメールアドレス1"
                                    .Vari_Emailaddr1 = fldvalue
                                Case "Eメールアドレス2"
                                    .Vari_Emailaddr2 = fldvalue
                                Case "Eメールアドレス3"
                                    .Vari_Emailaddr3 = fldvalue
                                Case "Eメールアドレス4"
                                    .Vari_Emailaddr4 = fldvalue
                                Case "Eメールアドレス5"
                                    .Vari_Emailaddr5 = fldvalue
                            End Select

                        Next

                        '固定値
                        .Vari_Setting_guid = Guid.NewGuid.ToString
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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                'Dim tmp_sql As String = ""
                'Dim flg As Boolean = True
                'flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

                'キーをプログラム内で作成しているため重複不可情報を考慮する必要あり

            End Sub

        End Class

    End Class

#End Region

#Region "送信設定自社web情報"

    Public Class M_site_sendsetting_jisyaweb_Repository

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
                Dim hash_settingguid As New Hashtable                           'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.M_site_sendsetting_Model      '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Nkinkomk()

                '革命10の入金項目マスタを取得
                Dim hash_nkinkomkmst As New Hashtable
                Dim tmp_sql As String = "SELECT nkin_no,nkin_name FROM m_nkin"
                Call DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_nkinkomkmst)

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                '送信設定guid取得
                Call Me.Set_SettingGuid(sqlcnnv10, hash_settingguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_site_sendsetting"
                Dim fldnamegrp As String = "sitesetting_guid,setting_guid,site_no,site_id,site_sendumu," & _
                                           "site_data,site_password"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_settingguid, list_basekeydata)

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
                        Dim hash_cvxmlitem As New Hashtable

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "送信設定順"
                                    .Vari_Setting_guid = fldvalue
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case "ポータルサイトID"
                                    .Vari_Site_id = fldvalue
                                Case "サイト別送信有無"
                                    .Vari_Site_sendumu = fldvalue
                                Case "ポータルサイトパスワード"
                                    .Vari_Site_password = IIf(fldvalue = "", " ", fldvalue)
                                Case Else
                                    If fldname <> "" Then
                                        hash_cvxmlitem.Add(fldname, fldvalue)
                                    End If
                            End Select

                        Next

                        '固定値
                        .Vari_Sitesetting_guid = Guid.NewGuid.ToString

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
                            hash_cvitem("setting_guid") = hash_settingguid.Item(StrConv(hash_cvitem.Item("setting_guid"), VbStrConv.Narrow))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                'xmlの移行
                                Call Me.Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem, hash_nkinkomkmst)

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

                'ポータルサイトパスワード一括更新 (半角スペース " " を設定していた分を空文字へ変換)
                Dim tmpcnt As Integer = 0
                Dim passwordupdateqry As String = " UPDATE " & tblname & " SET site_password = '' WHERE site_no = 10 AND site_password = ' ' "
                DBExec.Exec_NonQuery(sqlcnnv10, passwordupdateqry, tmpcnt)

                '共通設定情報へ個別設定No1を反映
                '20160908 個別処理を共通処理へコピーする処理の制御修正 -chg sta
                'Dim tmp_sortstr As String = ""  '使用しない
                'Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                'DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                If cvrowcnt <> 0 Then
                    Dim tmp_sortstr As String = ""  '使用しない
                    Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                    DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                End If
                '20160908 個別処理を共通処理へコピーする処理の制御修正 -chg end
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
            ''' 共通設定情報更新クエリの作成
            ''' </summary>
            ''' <param name="sortstr"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " UPDATE m_site_sendsetting SET "
                tmp_sql = tmp_sql & " 	 site_id = (SELECT site_id FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 10) "
                tmp_sql = tmp_sql & " 	,site_sendumu = (SELECT site_sendumu FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 10) "
                tmp_sql = tmp_sql & " 	,site_data = (SELECT site_data FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 10) "
                tmp_sql = tmp_sql & " 	,site_password = (SELECT site_password FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 10) "
                tmp_sql = tmp_sql & " WHERE setting_guid = "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0 "
                tmp_sql = tmp_sql & " ) "
                tmp_sql = tmp_sql & " AND site_no = 10 "

                Return tmp_sql

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) "
                tmp_sql = tmp_sql & " FROM hydata_kokoku AS HYKOKOKU "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 送信設定guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_SettingGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = " SELECT setting_sortorder,setting_guid FROM m_sendsetting "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 送信設定のxml項目の移行
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <remarks></remarks>
            Public Sub Set_XmlItem(ByVal sqlcnnv10 As SqlConnection, ByVal hash_cvitem As Hashtable, ByVal hash_xmlitem As Hashtable, ByVal hash_nkinkomkmst As Hashtable)

                Dim sitesettingguid As String = hash_cvitem("sitesetting_guid")   '更新対象選択用

                '「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                Dim xmlupdateqry_xmlpre As String = ""
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre & "<WmpSendSettingModel xmlns:i=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & ">"
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre & "<__identity xmlns=" & """" & "http://schemas.datacontract.org/2004/07/System" & """" & " i:nil=" & """" & "true" & """" & " />"
                Dim xmlupdateqry_xmlpost As String = ""
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost & "<ProductModelList />"
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost & "</WmpSendSettingModel>"

                'グループ毎にXml文字列を作成
                '画像
                Dim tmp_xml_gazo As String = Me.Get_XmlStr_Gazo(hash_xmlitem)
                '入金項目
                Dim tmp_xml_nkin As String = Me.Get_XmlStr_Nkinkomk(hash_xmlitem, hash_nkinkomkmst)
                'その他
                Dim tmp_xml_other As String = Me.Get_XmlStr_Other(hash_cvitem, hash_xmlitem)

                'Xml文字列の結合
                Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & tmp_xml_gazo & tmp_xml_nkin & tmp_xml_other & xmlupdateqry_xmlpost

                '更新用クエリの作成
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " DECLARE @tmp_xml XML "
                tmp_sql = tmp_sql & " SET @tmp_xml = CAST('" & xmlupdateqry_main & "' AS XML) "
                tmp_sql = tmp_sql & " UPDATE m_site_sendsetting SET site_data = @tmp_xml "
                tmp_sql = tmp_sql & " WHERE sitesetting_guid = '" & sitesettingguid & "' "

                '更新
                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

            ''' <summary>
            ''' 画像に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Gazo(ByVal hash As Hashtable) As String

                Dim rtn_str As String = ""
                Dim xml_pre As String = "<GazoSettingList xmlns:a=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">"
                Dim xml_post As String = "</GazoSettingList>"

                Dim gazocnt As Integer = 30
                Dim tmp_gazono(gazocnt) As String
                Dim tmp_gazosyuV7(gazocnt) As String
                Dim tmp_gazonoV7(gazocnt) As String
                Dim tmp_gazosyurendo(gazocnt) As String

                '画像に関する情報を配列へ格納
                For Each item In hash

                    Dim fldname As String = item.Key
                    Dim value As String = item.Value

                    Select Case True
                        Case fldname.Replace("画像No", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("画像No", ""))
                            tmp_gazono(gazono) = value
                        Case fldname.Replace("革命側の画像種別", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("革命側の画像種別", ""))
                            tmp_gazosyuV7(gazono) = value
                        Case fldname.Replace("革命側の画像タイトルNo", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("革命側の画像タイトルNo", ""))
                            tmp_gazonoV7(gazono) = value
                        Case fldname.Replace("連動側の画像種別", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("連動側の画像種別", ""))
                            tmp_gazosyurendo(gazono) = value
                    End Select

                Next

                '文字列へ成形
                Dim xml_gazoitem As New System.Text.StringBuilder
                For cntii = 1 To gazocnt
                    If tmp_gazosyurendo(cntii) <> "0" And tmp_gazosyuV7(cntii) <> "0" And tmp_gazonoV7(cntii) <> "0" Then
                        xml_gazoitem.Append("<a:PortalGazoSettingModel>")
                        xml_gazoitem.Append("<a:FK8LocationKbn>" & tmp_gazosyuV7(cntii) & "</a:FK8LocationKbn>")
                        xml_gazoitem.Append("<a:FK8LocationNo>" & tmp_gazonoV7(cntii) & "</a:FK8LocationNo>")
                        xml_gazoitem.Append("<a:No>" & tmp_gazono(cntii) & "</a:No>")
                        xml_gazoitem.Append("<a:PortalGazoSyubetu>" & tmp_gazosyurendo(cntii) & "</a:PortalGazoSyubetu>")
                        xml_gazoitem.Append("</a:PortalGazoSettingModel>")
                    End If
                Next

                '結合
                Dim xml_gazo As String = xml_pre & xml_gazoitem.ToString & xml_post
                rtn_str = xml_gazo

                Return rtn_str

            End Function

            ''' <summary>
            ''' 入金項目に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Nkinkomk(ByVal hash_xmlitem As Hashtable, ByVal hash_nkinkomkmst As Hashtable) As String

                Dim rtn_str As String = ""
                Dim xml_pre As String = "<NkinKomokDict xmlns:a=" & """" & "http://schemas.microsoft.com/2003/10/Serialization/Arrays" & """" & ">"
                Dim xml_post As String = "</NkinKomokDict>"
                Dim nkincnt As Integer = 11
                Dim cntindex As Integer = 1
                Dim xml_nkinkomk(nkincnt) As String
                Dim hash_nkinkomkitem As New Hashtable

                '入金項目を移行用に成形する
                Dim hash_cvxmlitem As New Hashtable
                hash_cvxmlitem = Me.Chg_NkinKomk(hash_xmlitem, hash_nkinkomkmst)

                '賃料
                hash_nkinkomkitem.Add("item_key", "1")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_賃料"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_賃料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_賃料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '管理費共益費
                hash_nkinkomkitem.Add("item_key", "2")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_共益費/管理費"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_共益費管理費")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_共益費/管理費No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '敷金
                hash_nkinkomkitem.Add("item_key", "6")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_敷金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("敷金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_敷金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_敷金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '礼金
                hash_nkinkomkitem.Add("item_key", "5")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_礼金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("礼金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_礼金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_礼金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '保証金
                hash_nkinkomkitem.Add("item_key", "7")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_保証金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("保証金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_保証金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_保証金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '償却金
                hash_nkinkomkitem.Add("item_key", "9")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_償却金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("償却金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_償却金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_償却金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '敷引(解約引き)
                hash_nkinkomkitem.Add("item_key", "12")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_敷引(解約引き)"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("敷引(解約引き)表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_敷引金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_敷引(解約引き)No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '更新料
                hash_nkinkomkitem.Add("item_key", "14")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_更新料"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("更新料表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_更新料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_更新料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '仲介手数料
                hash_nkinkomkitem.Add("item_key", "15")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_仲介手数料"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("仲介手数料表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_仲介手数料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_仲介手数料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                'その他一時金
                hash_nkinkomkitem.Add("item_key", "18")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_その他一時金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("その他一時金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_その他一時金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_その他一時金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()


                'その他月額費用
                hash_nkinkomkitem.Add("item_key", "17")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_その他月額費用"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_その他費用")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_その他月額費用No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '結合
                Dim xml_total As String = ""
                Dim xml_nkintotal As String = ""
                For cntii = 1 To nkincnt
                    xml_nkintotal = xml_nkintotal & xml_nkinkomk(cntii)
                Next
                xml_total = xml_pre & xml_nkintotal & xml_post

                rtn_str = xml_total

                Return rtn_str

            End Function

            ''' <summary>
            ''' 入金項目の編集
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_xml"></param>
            ''' <param name="hash_nkinkomkmst"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chg_NkinKomk(ByVal hash_xml As Hashtable, ByVal hash_nkinkomkmst As Hashtable) As Hashtable

                Dim rtn_hash As New Hashtable
                Dim tmp_hash As New Hashtable

                Dim tmp_nkincnt As Integer = 16
                Dim tmp_nkinkomkname(tmp_nkincnt) As String
                Dim tmp_nkinkomkkbn(tmp_nkincnt) As String
                Dim tmp_fldname(tmp_nkincnt) As String
                '-----------------------
                '入金項目取得
                '-----------------------
                For Each item In hash_xml

                    Dim midfldname As String = item.Key
                    Dim value As String = item.Value

                    Select Case midfldname
                        Case "入金項目_賃料1" : tmp_nkinkomkname(1) = value : tmp_fldname(1) = midfldname
                        Case "入金項目区分_賃料1" : tmp_nkinkomkkbn(1) = value
                        Case "入金項目_賃料2" : tmp_nkinkomkname(2) = value : tmp_fldname(2) = midfldname
                        Case "入金項目区分_賃料2" : tmp_nkinkomkkbn(2) = value
                        Case "入金項目_賃料3" : tmp_nkinkomkname(3) = value : tmp_fldname(3) = midfldname
                        Case "入金項目区分_賃料3" : tmp_nkinkomkkbn(3) = value
                        Case "入金項目_共益費/管理費1" : tmp_nkinkomkname(4) = value : tmp_fldname(4) = midfldname
                        Case "入金項目区分_共益費/管理費1" : tmp_nkinkomkkbn(4) = value
                        Case "入金項目_共益費/管理費2" : tmp_nkinkomkname(5) = value : tmp_fldname(5) = midfldname
                        Case "入金項目区分_共益費/管理費2" : tmp_nkinkomkkbn(5) = value
                        Case "入金項目_共益費/管理費3" : tmp_nkinkomkname(6) = value : tmp_fldname(6) = midfldname
                        Case "入金項目区分_共益費/管理費3" : tmp_nkinkomkkbn(6) = value
                        Case "入金項目_敷金" : tmp_nkinkomkname(7) = value : tmp_fldname(7) = midfldname
                        Case "入金項目区分_敷金" : tmp_nkinkomkkbn(7) = value
                        Case "入金項目_礼金" : tmp_nkinkomkname(8) = value : tmp_fldname(8) = midfldname
                        Case "入金項目区分_礼金" : tmp_nkinkomkkbn(8) = value
                        Case "入金項目_保証金" : tmp_nkinkomkname(9) = value : tmp_fldname(9) = midfldname
                        Case "入金項目区分_保証金" : tmp_nkinkomkkbn(9) = value
                        Case "入金項目_償却金" : tmp_nkinkomkname(10) = value : tmp_fldname(10) = midfldname
                        Case "入金項目区分_償却金" : tmp_nkinkomkkbn(10) = value
                        Case "入金項目_敷引(解約引き)" : tmp_nkinkomkname(11) = value : tmp_fldname(11) = midfldname
                        Case "入金項目区分_敷引(解約引き)" : tmp_nkinkomkkbn(11) = value
                        Case "入金項目_更新料" : tmp_nkinkomkname(12) = value : tmp_fldname(12) = midfldname
                        Case "入金項目区分_更新料" : tmp_nkinkomkkbn(12) = value
                        Case "入金項目_仲介手数料" : tmp_nkinkomkname(13) = value : tmp_fldname(13) = midfldname
                        Case "入金項目区分_仲介手数料" : tmp_nkinkomkkbn(13) = value
                        Case "入金項目_その他一時金" : tmp_nkinkomkname(14) = value : tmp_fldname(14) = midfldname
                        Case "入金項目区分_その他一時金" : tmp_nkinkomkkbn(14) = value
                        Case "入金項目_その他月額費用1" : tmp_nkinkomkname(15) = value : tmp_fldname(15) = midfldname
                        Case "入金項目区分_その他月額費用1" : tmp_nkinkomkkbn(15) = value
                        Case "入金項目_その他月額費用2" : tmp_nkinkomkname(16) = value : tmp_fldname(16) = midfldname
                        Case "入金項目区分_その他月額費用2" : tmp_nkinkomkkbn(16) = value
                        Case Else
                            tmp_hash.Add(midfldname, value)
                    End Select

                Next

                '-----------------------
                '入金項目をXml移行用に成形
                '-----------------------
                Dim tmp_nkinno(tmp_nkincnt) As String
                Dim tmp_dispnkinname(tmp_nkincnt) As String
                For cntii = 1 To tmp_nkincnt

                    Dim tmp_relnkinstr As String = tmp_nkinkomkname(cntii) & "-" & EtcMethod.Get_Nkinruiname(tmp_nkinkomkkbn(cntii))
                    tmp_nkinno(cntii) = ""
                    tmp_dispnkinname(cntii) = ""

                    If Hash_Rel_Nkinkomk.Contains(tmp_relnkinstr) Then
                        tmp_nkinno(cntii) = Hash_Rel_Nkinkomk(tmp_relnkinstr)
                        tmp_dispnkinname(cntii) = tmp_nkinno(cntii) & ":" & hash_nkinkomkmst(tmp_nkinno(cntii))
                    End If

                Next

                '-----------------------
                '入金項目の再格納
                '-----------------------
                '結合する項目
                Call Me.Set_NkinkomkAddToHash(1, 3, tmp_nkinno, tmp_dispnkinname, "入金項目_賃料", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(4, 6, tmp_nkinno, tmp_dispnkinname, "入金項目_共益費/管理費", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(15, 16, tmp_nkinno, tmp_dispnkinname, "入金項目_その他月額費用", tmp_hash)

                '結合不要の項目
                For cntii = 7 To 14
                    tmp_hash.Add(tmp_fldname(cntii), tmp_dispnkinname(cntii))
                    tmp_hash.Add(tmp_fldname(cntii) & "No", tmp_nkinno(cntii))
                Next

                rtn_hash = tmp_hash

                Return rtn_hash

            End Function

            ''' <summary>
            ''' 入金項目を結合して1つのデータにまとめる
            ''' まとめたデータをハッシュテーブルへ格納する
            ''' </summary>
            ''' <param name="cntsta"></param>
            ''' <param name="cntend"></param>
            ''' <param name="tmp_nkinno"></param>
            ''' <param name="tmp_dispnkinname"></param>
            ''' <param name="hashkey"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_NkinkomkAddToHash(ByVal cntsta As Integer, ByVal cntend As Integer, ByVal tmp_nkinno As Object, ByVal tmp_dispnkinname As Object, ByVal hashkey As String, ByRef hash As Hashtable)

                Dim nkinno As String = ""
                Dim dispname As String = ""
                For cntii = cntsta To cntend
                    If tmp_nkinno(cntii) <> "" Then
                        nkinno = nkinno & "," & tmp_nkinno(cntii)
                        dispname = dispname & " / " & tmp_dispnkinname(cntii)
                    End If
                Next
                If nkinno <> "" Then
                    nkinno = nkinno.Remove(0, 1)
                End If
                If dispname <> "" Then
                    dispname = dispname.Remove(0, 3)
                End If
                hash.Add(hashkey, dispname)
                hash.Add(hashkey & "No", nkinno)

            End Sub

            ''' <summary>
            ''' 各入金項目毎にXml文字列作成
            ''' </summary>
            ''' <param name="hash_nkin"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chg_NkinkomkToXml(ByVal hash_nkin As Hashtable) As String

                Dim rtn_str As String = ""

                '置換用/ハッシュテーブル値取得用文字列
                Dim item_key As String = "item_key"
                Dim item_disp As String = "item_disp"
                Dim item_tani As String = "item_tani"
                Dim item_komk As String = "item_komk"
                Dim item_nkinno As String = "item_nkinno"

                'Xmlのベースを作成
                Dim xml_nkinkomkbase As New System.Text.StringBuilder
                xml_nkinkomkbase.Append("<a:KeyValueOfintPortalNkinKomokLjh4bohd>")
                xml_nkinkomkbase.Append("<a:Key>" & item_key & "</a:Key>")
                xml_nkinkomkbase.Append("<a:Value>")
                xml_nkinkomkbase.Append("<DisplayString>" & item_disp & "</DisplayString>")
                xml_nkinkomkbase.Append("<EnKagetu>" & item_tani & "</EnKagetu>")
                xml_nkinkomkbase.Append("<NkinKomok>" & item_komk & "</NkinKomok>")
                xml_nkinkomkbase.Append("<NkinNoList>")
                xml_nkinkomkbase.Append(item_nkinno)
                xml_nkinkomkbase.Append("</NkinNoList>")
                xml_nkinkomkbase.Append("</a:Value>")
                xml_nkinkomkbase.Append("</a:KeyValueOfintPortalNkinKomokLjh4bohd>")

                '各要素で置換してい成形
                Dim xml_nkinkomk As String = xml_nkinkomkbase.ToString
                xml_nkinkomk = xml_nkinkomk.Replace(item_key, hash_nkin(item_key))
                xml_nkinkomk = xml_nkinkomk.Replace(item_disp, hash_nkin(item_disp))
                xml_nkinkomk = xml_nkinkomk.Replace(item_tani, hash_nkin(item_tani))
                xml_nkinkomk = xml_nkinkomk.Replace(item_komk, hash_nkin(item_komk))
                Dim tmp_nkinnogrp_tinryo As String = hash_nkin(item_nkinno)
                Dim tmp_nkinno_tinryo() As String = tmp_nkinnogrp_tinryo.Split(",")
                Dim nkinno_tinryo As String = ""
                For cntii = 0 To UBound(tmp_nkinno_tinryo)
                    If tmp_nkinno_tinryo(cntii) <> "" Then
                        nkinno_tinryo = nkinno_tinryo & "<a:int>" & tmp_nkinno_tinryo(cntii) & "</a:int>"
                    End If
                Next
                xml_nkinkomk = xml_nkinkomk.Replace(item_nkinno, nkinno_tinryo)

                rtn_str = xml_nkinkomk

                Return rtn_str

            End Function

            ''' <summary>
            ''' その他項目に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_xmlitem"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Other(ByVal hash_cvitem As Hashtable, ByVal hash_xmlitem As Hashtable) As String

                Dim rtn_str As String = ""
                Dim othercnt As Integer = 19
                Dim xml_other(othercnt) As String
                Dim cntindex As Integer = 1

                xml_other(cntindex) = "<Password>" & hash_cvitem("site_password") & "</Password>" : cntindex = cntindex + 1           'パスワード
                xml_other(cntindex) = "<SiteID>" & hash_cvitem("site_id") & "</SiteID>" : cntindex = cntindex + 1                       'ポータルサイトID
                xml_other(cntindex) = "<SiteSendUmu>" & IIf(hash_cvitem("site_sendumu") = "1", "true", "false") & "</SiteSendUmu>" : cntindex = cntindex + 1           'サイト別送信有無
                xml_other(cntindex) = "<UseCommonSetting>" & "false" & "</UseCommonSetting>" : cntindex = cntindex + 1                         '共通設定使用有無
                xml_other(cntindex) = "<BantiDispSetting>" & hash_xmlitem("番地以降の表示") & "</BantiDispSetting>" : cntindex = cntindex + 1   '番地以降の表示
                xml_other(cntindex) = "<BkHyDispSetting>" & hash_xmlitem("物件部屋の表示") & "</BkHyDispSetting>" : cntindex = cntindex + 1   '物件部屋の表示
                xml_other(cntindex) = "<EmailAddr1>" & hash_xmlitem("Eメールアドレス1") & "</EmailAddr1>" : cntindex = cntindex + 1          'Eメールアドレス1
                xml_other(cntindex) = "<EmailAddr2>" & hash_xmlitem("Eメールアドレス2") & "</EmailAddr2>" : cntindex = cntindex + 1          'Eメールアドレス2
                xml_other(cntindex) = "<EmailAddr3>" & hash_xmlitem("Eメールアドレス3") & "</EmailAddr3>" : cntindex = cntindex + 1          'Eメールアドレス3
                xml_other(cntindex) = "<EmailAddr4>" & hash_xmlitem("Eメールアドレス4") & "</EmailAddr4>" : cntindex = cntindex + 1         'Eメールアドレス4
                xml_other(cntindex) = "<EmailAddr5>" & hash_xmlitem("Eメールアドレス5") & "</EmailAddr5>" : cntindex = cntindex + 1         'Eメールアドレス5

                '固定分
                Dim tmp_strb As New System.Text.StringBuilder
                tmp_strb.Append("<GazoSettingList>")
                tmp_strb.Append("<WmpGazoSettingModel>")
                tmp_strb.Append("<FK8LocationKbn xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</FK8LocationKbn>")
                tmp_strb.Append("<FK8LocationNo xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</FK8LocationNo>")
                tmp_strb.Append("<No xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</No>")
                tmp_strb.Append("<PortalGazoSyubetu xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</PortalGazoSyubetu>")
                tmp_strb.Append("</WmpGazoSettingModel>")
                tmp_strb.Append("<WmpGazoSettingModel>")
                tmp_strb.Append("<FK8LocationKbn xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">2</FK8LocationKbn>")
                tmp_strb.Append("<FK8LocationNo xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</FK8LocationNo>")
                tmp_strb.Append("<No xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">16</No>")
                tmp_strb.Append("<PortalGazoSyubetu xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">5</PortalGazoSyubetu>")
                tmp_strb.Append("</WmpGazoSettingModel>")
                tmp_strb.Append("</GazoSettingList>")
                xml_other(cntindex) = tmp_strb.ToString : cntindex = cntindex + 1

                xml_other(cntindex) = "<KagiInfo1>" & hash_xmlitem("鍵情報1") & "</KagiInfo1>" : cntindex = cntindex + 1                  '鍵情報1
                xml_other(cntindex) = "<KagiInfo2>" & hash_xmlitem("鍵情報2") & "</KagiInfo2>" : cntindex = cntindex + 1                  '鍵情報2
                xml_other(cntindex) = "<KagiInfo3>" & hash_xmlitem("鍵情報3") & "</KagiInfo3>" : cntindex = cntindex + 1                  '鍵情報3
                xml_other(cntindex) = "<KurasapoIgaiMapDispSetting>" & hash_xmlitem("くらさぽ以外への地図表示") & "</KurasapoIgaiMapDispSetting>" : cntindex = cntindex + 1   'くらさぽ以外への地図表示
                xml_other(cntindex) = "<KurasapoMapDispSetting>" & hash_xmlitem("くらさぽへの地図表示") & "</KurasapoMapDispSetting>" : cntindex = cntindex + 1              'くらさぽへの地図表示
                xml_other(cntindex) = "<MovieDispSetting>" & hash_xmlitem("動画の表示") & "</MovieDispSetting>" : cntindex = cntindex + 1                       '動画の表示

                '動画の表示順序
                Dim tmp_replacedoga As String = "item_doga"
                Dim tmp_xml_gazojyunjo As String = "<MovieYusenList xmlns:a=" & """" & "http://schemas.datacontract.org/2004/07/Njc.N3Lib.Utys" & """" & ">" & tmp_replacedoga & "</MovieYusenList>"

                Dim dogacnt As Integer = 4
                Dim xml_doga(dogacnt) As String
                Dim tmp_dogastrb As New System.Text.StringBuilder
                For cntii = 1 To dogacnt
                    Dim dogavalue As String = ""
                    Select Case cntii
                        Case 1
                            dogavalue = "物件動画１"
                        Case 2
                            dogavalue = "物件動画２"
                        Case 3
                            dogavalue = "部屋動画１"
                        Case 4
                            dogavalue = "部屋動画２"
                    End Select
                    tmp_dogastrb.Append("<a:DataKeyValue>")
                    tmp_dogastrb.Append("<__identity xmlns=" & """" & "http://schemas.datacontract.org/2004/07/System" & """" & " i:nil=" & """" & "true" & """" & " />")
                    tmp_dogastrb.Append("<a:_alwaysVisible>true</a:_alwaysVisible>")
                    tmp_dogastrb.Append("<a:_key>" & cntii.ToString & "</a:_key>")
                    tmp_dogastrb.Append("<a:_keyDescription i:nil=" & """" & "true" & """" & " />")
                    tmp_dogastrb.Append("<a:_obj i:nil=" & """" & "true" & """" & " />")
                    tmp_dogastrb.Append("<a:_value>" & dogavalue & "</a:_value>")
                    tmp_dogastrb.Append("</a:DataKeyValue>")
                    xml_doga(cntii) = tmp_dogastrb.ToString
                    tmp_dogastrb.Clear()
                Next
                If hash_xmlitem("動画の表示") = "2" Then
                    tmp_xml_gazojyunjo = tmp_xml_gazojyunjo.Replace(tmp_replacedoga, xml_doga(Int32.Parse(hash_xmlitem("動画の優先項目1"))) & xml_doga(Int32.Parse(hash_xmlitem("動画の優先項目2"))) & _
                                                                           xml_doga(Int32.Parse(hash_xmlitem("動画の優先項目3"))) & xml_doga(Int32.Parse(hash_xmlitem("動画の優先項目4"))))
                Else
                    tmp_xml_gazojyunjo = tmp_xml_gazojyunjo.Replace(tmp_replacedoga, xml_doga(1) & xml_doga(2) & xml_doga(3) & xml_doga(4))
                End If
                xml_other(cntindex) = tmp_xml_gazojyunjo

                '結合
                Dim tmp_str As String = ""
                For cntii = 1 To othercnt
                    tmp_str = tmp_str & xml_other(cntii)
                Next

                rtn_str = tmp_str

                Return rtn_str

            End Function

        End Class

    End Class

#End Region

#Region "送信設定HOMES情報"

    Public Class M_site_sendsetting_homes_Repository

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
                Dim hash_settingguid As New Hashtable                           'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.M_site_sendsetting_Model      '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Nkinkomk()

                '革命10の入金項目マスタを取得
                Dim hash_nkinkomkmst As New Hashtable
                Dim tmp_sql As String = "SELECT nkin_no,nkin_name FROM m_nkin"
                Call DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_nkinkomkmst)

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                '送信設定guid取得
                Call Me.Set_SettingGuid(sqlcnnv10, hash_settingguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_site_sendsetting"
                Dim fldnamegrp As String = "sitesetting_guid,setting_guid,site_no,site_id,site_sendumu," & _
                                           "site_data,site_password"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_settingguid, list_basekeydata)

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
                        Dim hash_cvxmlitem As New Hashtable

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "送信設定順"
                                    .Vari_Setting_guid = fldvalue
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case "ポータルサイトID"
                                    .Vari_Site_id = fldvalue
                                Case "サイト別送信有無"
                                    .Vari_Site_sendumu = fldvalue
                                Case "ポータルサイトパスワード"
                                    .Vari_Site_password = IIf(fldvalue = "", " ", fldvalue)
                                Case Else
                                    If fldname <> "" Then
                                        hash_cvxmlitem.Add(fldname, fldvalue)
                                    End If
                            End Select

                        Next

                        '固定値
                        .Vari_Sitesetting_guid = Guid.NewGuid.ToString

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        'FTPユーザーID、パスワード有無チェック
                        Call Me.Chk_FTPData(sqlcnnv10, tblname, hash_cvitem, hash_cvxmlitem)

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            hash_cvitem("setting_guid") = hash_settingguid.Item(StrConv(hash_cvitem.Item("setting_guid"), VbStrConv.Narrow))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                'xmlの移行
                                Call Me.Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem, hash_nkinkomkmst)

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

                'ポータルサイトパスワード一括更新 (半角スペース " " を設定していた分を空文字へ変換)
                Dim tmpcnt As Integer = 0
                Dim passwordupdateqry As String = " UPDATE " & tblname & " SET site_password = '' WHERE site_no = 20 AND site_password = ' ' "
                DBExec.Exec_NonQuery(sqlcnnv10, passwordupdateqry, tmpcnt)

                '共通設定情報へ個別設定No1を反映
                '20160908 個別処理を共通処理へコピーする処理の制御修正 -chg sta
                'Dim tmp_sortstr As String = ""  '使用しない
                'Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                'DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                If cvrowcnt <> 0 Then
                    Dim tmp_sortstr As String = ""  '使用しない
                    Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                    DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                End If
                '20160908 個別処理を共通処理へコピーする処理の制御修正 -chg end

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
            ''' 共通設定情報更新クエリの作成
            ''' </summary>
            ''' <param name="sortstr"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " UPDATE m_site_sendsetting SET "
                tmp_sql = tmp_sql & " 	 site_id = (SELECT site_id FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 20) "
                tmp_sql = tmp_sql & " 	,site_sendumu = (SELECT site_sendumu FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 20) "
                tmp_sql = tmp_sql & " 	,site_data = (SELECT site_data FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 20) "
                tmp_sql = tmp_sql & " 	,site_password = (SELECT site_password FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 20) "
                tmp_sql = tmp_sql & " WHERE setting_guid = "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0 "
                tmp_sql = tmp_sql & " ) "
                tmp_sql = tmp_sql & " AND site_no = 20 "

                Return tmp_sql

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) "
                tmp_sql = tmp_sql & " FROM hydata_kokoku AS HYKOKOKU "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 送信設定guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_SettingGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = " SELECT setting_sortorder,setting_guid FROM m_sendsetting "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 送信設定のxml項目の移行
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <remarks></remarks>
            Public Sub Set_XmlItem(ByVal sqlcnnv10 As SqlConnection, ByVal hash_cvitem As Hashtable, ByVal hash_xmlitem As Hashtable, ByVal hash_nkinkomkmst As Hashtable)

                Dim sitesettingguid As String = hash_cvitem("sitesetting_guid")   '更新対象選択用

                '「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                Dim xmlupdateqry_xmlpre As String = ""
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre & "<HomesSendSettingModel xmlns:i=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & ">"
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre & "<__identity xmlns=" & """" & "http://schemas.datacontract.org/2004/07/System" & """" & " i:nil=" & """" & "true" & """" & " />"
                Dim xmlupdateqry_xmlpost As String = ""
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost & "</HomesSendSettingModel>"

                'グループ毎にXml文字列を作成
                '画像
                Dim tmp_xml_gazo As String = Me.Get_XmlStr_Gazo(hash_xmlitem)
                '入金項目
                Dim tmp_xml_nkin As String = Me.Get_XmlStr_Nkinkomk(hash_xmlitem, hash_nkinkomkmst)
                'その他
                Dim tmp_xml_other As String = Me.Get_XmlStr_Other(hash_cvitem, hash_xmlitem)

                'Xml文字列の結合
                Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & tmp_xml_gazo & tmp_xml_nkin & tmp_xml_other & xmlupdateqry_xmlpost
                'Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & tmp_xml_nkin & xmlupdateqry_xmlpost

                '更新用クエリの作成
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " DECLARE @tmp_xml XML "
                tmp_sql = tmp_sql & " SET @tmp_xml = CAST('" & xmlupdateqry_main & "' AS XML) "
                tmp_sql = tmp_sql & " UPDATE m_site_sendsetting SET site_data = @tmp_xml "
                tmp_sql = tmp_sql & " WHERE sitesetting_guid = '" & sitesettingguid & "' "

                '更新
                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

            ''' <summary>
            ''' 画像に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Gazo(ByVal hash As Hashtable) As String

                Dim rtn_str As String = ""
                Dim xml_pre As String = "<GazoSettingList xmlns:a=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">"
                Dim xml_post As String = "</GazoSettingList>"

                Dim gazocnt As Integer = 30
                Dim tmp_gazono(gazocnt) As String
                Dim tmp_gazosyuV7(gazocnt) As String
                Dim tmp_gazonoV7(gazocnt) As String
                Dim tmp_gazosyurendo(gazocnt) As String

                '画像に関する情報を配列へ格納
                For Each item In hash

                    Dim fldname As String = item.Key
                    Dim value As String = item.Value

                    Select Case True
                        Case fldname.Replace("画像No", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("画像No", ""))
                            tmp_gazono(gazono) = value
                        Case fldname.Replace("革命側の画像種別", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("革命側の画像種別", ""))
                            tmp_gazosyuV7(gazono) = value
                        Case fldname.Replace("革命側の画像タイトルNo", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("革命側の画像タイトルNo", ""))
                            tmp_gazonoV7(gazono) = value
                        Case fldname.Replace("連動側の画像種別", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("連動側の画像種別", ""))
                            tmp_gazosyurendo(gazono) = value
                    End Select

                Next

                '文字列へ成形
                Dim xml_gazoitem As New System.Text.StringBuilder
                For cntii = 1 To gazocnt
                    If tmp_gazosyurendo(cntii) <> "0" And tmp_gazosyuV7(cntii) <> "0" And tmp_gazonoV7(cntii) <> "0" Then
                        xml_gazoitem.Append("<a:PortalGazoSettingModel>")
                        xml_gazoitem.Append("<a:FK8LocationKbn>" & tmp_gazosyuV7(cntii) & "</a:FK8LocationKbn>")
                        xml_gazoitem.Append("<a:FK8LocationNo>" & tmp_gazonoV7(cntii) & "</a:FK8LocationNo>")
                        xml_gazoitem.Append("<a:No>" & tmp_gazono(cntii) & "</a:No>")
                        xml_gazoitem.Append("<a:PortalGazoSyubetu>" & tmp_gazosyurendo(cntii) & "</a:PortalGazoSyubetu>")
                        xml_gazoitem.Append("</a:PortalGazoSettingModel>")
                    End If
                Next

                '結合
                Dim xml_gazo As String = xml_pre & xml_gazoitem.ToString & xml_post
                rtn_str = xml_gazo

                Return rtn_str

            End Function

            ''' <summary>
            ''' 入金項目に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Nkinkomk(ByVal hash_xmlitem As Hashtable, ByVal hash_nkinkomkmst As Hashtable) As String

                Dim rtn_str As String = ""
                Dim xml_pre As String = "<NkinKomokDict xmlns:a=" & """" & "http://schemas.microsoft.com/2003/10/Serialization/Arrays" & """" & ">"
                Dim xml_post As String = "</NkinKomokDict>"
                Dim nkincnt As Integer = 13
                Dim cntindex As Integer = 1
                Dim xml_nkinkomk(nkincnt) As String
                Dim hash_nkinkomkitem As New Hashtable

                '入金項目を移行用に成形する
                Dim hash_cvxmlitem As New Hashtable
                hash_cvxmlitem = Me.Chg_NkinKomk(hash_xmlitem, hash_nkinkomkmst)

                '賃料
                hash_nkinkomkitem.Add("item_key", "1")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_賃料"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_賃料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_賃料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '管理費/共益費
                hash_nkinkomkitem.Add("item_key", "2")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_管理費/共益費"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_共益費管理費")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_管理費/共益費No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '敷金
                hash_nkinkomkitem.Add("item_key", "6")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_敷金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("敷金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_敷金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_敷金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '礼金
                hash_nkinkomkitem.Add("item_key", "5")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_礼金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("礼金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_礼金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_礼金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '保証金
                hash_nkinkomkitem.Add("item_key", "7")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_保証金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("保証金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_保証金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_保証金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '権利金
                hash_nkinkomkitem.Add("item_key", "8")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_権利金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("権利金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_権利金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_権利金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '償却/敷引金
                hash_nkinkomkitem.Add("item_key", "10")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_償却/敷引金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("償却/敷引金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_償却敷引金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_償却/敷引金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '更新料
                hash_nkinkomkitem.Add("item_key", "14")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_更新料"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("更新料表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_更新料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_更新料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '造作譲渡金
                hash_nkinkomkitem.Add("item_key", "16")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_造作譲渡金"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_造作譲渡金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_造作譲渡金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '仲介手数料
                hash_nkinkomkitem.Add("item_key", "15")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_仲介手数料"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_仲介手数料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_仲介手数料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '鍵交換費用
                hash_nkinkomkitem.Add("item_key", "20")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_鍵交換費用"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_鍵交換代等")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_鍵交換費用No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '室内清掃費用
                hash_nkinkomkitem.Add("item_key", "23")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_室内清掃費用"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_室内清掃費")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_室内清掃費用No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                'その他月額費用
                hash_nkinkomkitem.Add("item_key", "17")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_その他月額費用"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_その他費用")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_その他月額費用No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '結合
                Dim xml_total As String = ""
                Dim xml_nkintotal As String = ""
                For cntii = 1 To nkincnt
                    xml_nkintotal = xml_nkintotal & xml_nkinkomk(cntii)
                Next
                xml_total = xml_pre & xml_nkintotal & xml_post

                rtn_str = xml_total

                Return rtn_str

            End Function

            ''' <summary>
            ''' 入金項目の編集
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_xml"></param>
            ''' <param name="hash_nkinkomkmst"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chg_NkinKomk(ByVal hash_xml As Hashtable, ByVal hash_nkinkomkmst As Hashtable) As Hashtable

                Dim rtn_hash As New Hashtable
                Dim tmp_hash As New Hashtable

                Dim tmp_nkincnt As Integer = 19
                Dim tmp_nkinkomkname(tmp_nkincnt) As String
                Dim tmp_nkinkomkkbn(tmp_nkincnt) As String
                Dim tmp_fldname(tmp_nkincnt) As String
                '-----------------------
                '入金項目取得
                '-----------------------
                For Each item In hash_xml

                    Dim midfldname As String = item.Key
                    Dim value As String = item.Value

                    Select Case midfldname
                        Case "入金項目_賃料1" : tmp_nkinkomkname(1) = value : tmp_fldname(1) = midfldname
                        Case "入金項目区分_賃料1" : tmp_nkinkomkkbn(1) = value
                        Case "入金項目_賃料2" : tmp_nkinkomkname(2) = value : tmp_fldname(2) = midfldname
                        Case "入金項目区分_賃料2" : tmp_nkinkomkkbn(2) = value
                        Case "入金項目_賃料3" : tmp_nkinkomkname(3) = value : tmp_fldname(3) = midfldname
                        Case "入金項目区分_賃料3" : tmp_nkinkomkkbn(3) = value
                        Case "入金項目_管理費/共益費1" : tmp_nkinkomkname(4) = value : tmp_fldname(4) = midfldname
                        Case "入金項目区分_管理費/共益費1" : tmp_nkinkomkkbn(4) = value
                        Case "入金項目_管理費/共益費2" : tmp_nkinkomkname(5) = value : tmp_fldname(5) = midfldname
                        Case "入金項目区分_管理費/共益費2" : tmp_nkinkomkkbn(5) = value
                        Case "入金項目_管理費/共益費3" : tmp_nkinkomkname(6) = value : tmp_fldname(6) = midfldname
                        Case "入金項目区分_管理費/共益費3" : tmp_nkinkomkkbn(6) = value
                        Case "入金項目_敷金" : tmp_nkinkomkname(7) = value : tmp_fldname(7) = midfldname
                        Case "入金項目区分_敷金" : tmp_nkinkomkkbn(7) = value
                        Case "入金項目_礼金" : tmp_nkinkomkname(8) = value : tmp_fldname(8) = midfldname
                        Case "入金項目区分_礼金" : tmp_nkinkomkkbn(8) = value
                        Case "入金項目_保証金" : tmp_nkinkomkname(9) = value : tmp_fldname(9) = midfldname
                        Case "入金項目区分_保証金" : tmp_nkinkomkkbn(9) = value
                        Case "入金項目_権利金" : tmp_nkinkomkname(10) = value : tmp_fldname(10) = midfldname
                        Case "入金項目区分_権利金" : tmp_nkinkomkkbn(10) = value
                        Case "入金項目_償却/敷引金" : tmp_nkinkomkname(11) = value : tmp_fldname(11) = midfldname
                        Case "入金項目区分_償却/敷引金" : tmp_nkinkomkkbn(11) = value
                        Case "入金項目_更新料" : tmp_nkinkomkname(12) = value : tmp_fldname(12) = midfldname
                        Case "入金項目区分_更新料" : tmp_nkinkomkkbn(12) = value
                        Case "入金項目_造作譲渡金" : tmp_nkinkomkname(13) = value : tmp_fldname(13) = midfldname
                        Case "入金項目区分_造作譲渡金" : tmp_nkinkomkkbn(13) = value
                        Case "入金項目_仲介手数料" : tmp_nkinkomkname(14) = value : tmp_fldname(14) = midfldname
                        Case "入金項目区分_仲介手数料" : tmp_nkinkomkkbn(14) = value
                        Case "入金項目_鍵交換費用" : tmp_nkinkomkname(15) = value : tmp_fldname(15) = midfldname
                        Case "入金項目区分_鍵交換費用" : tmp_nkinkomkkbn(15) = value
                        Case "入金項目_室内清掃費用" : tmp_nkinkomkname(16) = value : tmp_fldname(16) = midfldname
                        Case "入金項目区分_室内清掃費用" : tmp_nkinkomkkbn(16) = value
                        Case "入金項目_その他月額費用1" : tmp_nkinkomkname(17) = value : tmp_fldname(17) = midfldname
                        Case "入金項目区分_その他費用1" : tmp_nkinkomkkbn(17) = value
                        Case "入金項目_その他月額費用2" : tmp_nkinkomkname(18) = value : tmp_fldname(18) = midfldname
                        Case "入金項目区分_その他費用2" : tmp_nkinkomkkbn(18) = value
                        Case "入金項目_その他月額費用3" : tmp_nkinkomkname(19) = value : tmp_fldname(19) = midfldname
                        Case "入金項目区分_その他費用3" : tmp_nkinkomkkbn(19) = value
                        Case Else
                            tmp_hash.Add(midfldname, value)
                    End Select

                Next

                '-----------------------
                '入金項目をXml移行用に成形
                '-----------------------
                Dim tmp_nkinno(tmp_nkincnt) As String
                Dim tmp_dispnkinname(tmp_nkincnt) As String
                For cntii = 1 To tmp_nkincnt

                    Dim tmp_relnkinstr As String = tmp_nkinkomkname(cntii) & "-" & EtcMethod.Get_Nkinruiname(tmp_nkinkomkkbn(cntii))
                    tmp_nkinno(cntii) = ""
                    tmp_dispnkinname(cntii) = ""

                    If Hash_Rel_Nkinkomk.Contains(tmp_relnkinstr) Then
                        tmp_nkinno(cntii) = Hash_Rel_Nkinkomk(tmp_relnkinstr)
                        tmp_dispnkinname(cntii) = tmp_nkinno(cntii) & ":" & hash_nkinkomkmst(tmp_nkinno(cntii))
                    End If

                Next

                '-----------------------
                '入金項目の再格納
                '-----------------------
                '結合する項目
                Call Me.Set_NkinkomkAddToHash(1, 3, tmp_nkinno, tmp_dispnkinname, "入金項目_賃料", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(4, 6, tmp_nkinno, tmp_dispnkinname, "入金項目_管理費/共益費", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(17, 19, tmp_nkinno, tmp_dispnkinname, "入金項目_その他月額費用", tmp_hash)

                '結合不要の項目
                For cntii = 7 To 16
                    tmp_hash.Add(tmp_fldname(cntii), tmp_dispnkinname(cntii))
                    tmp_hash.Add(tmp_fldname(cntii) & "No", tmp_nkinno(cntii))
                Next

                rtn_hash = tmp_hash

                Return rtn_hash

            End Function

            ''' <summary>
            ''' 入金項目を結合して1つのデータにまとめる
            ''' まとめたデータをハッシュテーブルへ格納する
            ''' </summary>
            ''' <param name="cntsta"></param>
            ''' <param name="cntend"></param>
            ''' <param name="tmp_nkinno"></param>
            ''' <param name="tmp_dispnkinname"></param>
            ''' <param name="hashkey"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_NkinkomkAddToHash(ByVal cntsta As Integer, ByVal cntend As Integer, ByVal tmp_nkinno As Object, ByVal tmp_dispnkinname As Object, ByVal hashkey As String, ByRef hash As Hashtable)

                Dim nkinno As String = ""
                Dim dispname As String = ""
                For cntii = cntsta To cntend
                    If tmp_nkinno(cntii) <> "" Then
                        nkinno = nkinno & "," & tmp_nkinno(cntii)
                        dispname = dispname & " / " & tmp_dispnkinname(cntii)
                    End If
                Next
                If nkinno <> "" Then
                    nkinno = nkinno.Remove(0, 1)
                End If
                If dispname <> "" Then
                    dispname = dispname.Remove(0, 3)
                End If
                hash.Add(hashkey, dispname)
                hash.Add(hashkey & "No", nkinno)

            End Sub

            ''' <summary>
            ''' 各入金項目毎にXml文字列作成
            ''' </summary>
            ''' <param name="hash_nkin"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chg_NkinkomkToXml(ByVal hash_nkin As Hashtable) As String

                Dim rtn_str As String = ""

                '置換用/ハッシュテーブル値取得用文字列
                Dim item_key As String = "item_key"
                Dim item_disp As String = "item_disp"
                Dim item_tani As String = "item_tani"
                Dim item_komk As String = "item_komk"
                Dim item_nkinno As String = "item_nkinno"

                'Xmlのベースを作成
                Dim xml_nkinkomkbase As New System.Text.StringBuilder
                xml_nkinkomkbase.Append("<a:KeyValueOfintPortalNkinKomokLjh4bohd>")
                xml_nkinkomkbase.Append("<a:Key>" & item_key & "</a:Key>")
                xml_nkinkomkbase.Append("<a:Value>")
                xml_nkinkomkbase.Append("<DisplayString>" & item_disp & "</DisplayString>")
                xml_nkinkomkbase.Append("<EnKagetu>" & item_tani & "</EnKagetu>")
                xml_nkinkomkbase.Append("<NkinKomok>" & item_komk & "</NkinKomok>")
                xml_nkinkomkbase.Append("<NkinNoList>")
                xml_nkinkomkbase.Append(item_nkinno)
                xml_nkinkomkbase.Append("</NkinNoList>")
                xml_nkinkomkbase.Append("</a:Value>")
                xml_nkinkomkbase.Append("</a:KeyValueOfintPortalNkinKomokLjh4bohd>")

                '各要素で置換してい成形
                Dim xml_nkinkomk As String = xml_nkinkomkbase.ToString
                xml_nkinkomk = xml_nkinkomk.Replace(item_key, hash_nkin(item_key))
                xml_nkinkomk = xml_nkinkomk.Replace(item_disp, hash_nkin(item_disp))
                xml_nkinkomk = xml_nkinkomk.Replace(item_tani, hash_nkin(item_tani))
                xml_nkinkomk = xml_nkinkomk.Replace(item_komk, hash_nkin(item_komk))
                Dim tmp_nkinnogrp_tinryo As String = hash_nkin(item_nkinno)
                Dim tmp_nkinno_tinryo() As String = tmp_nkinnogrp_tinryo.Split(",")
                Dim nkinno_tinryo As String = ""
                For cntii = 0 To UBound(tmp_nkinno_tinryo)
                    If tmp_nkinno_tinryo(cntii) <> "" Then
                        nkinno_tinryo = nkinno_tinryo & "<a:int>" & tmp_nkinno_tinryo(cntii) & "</a:int>"
                    End If
                Next
                xml_nkinkomk = xml_nkinkomk.Replace(item_nkinno, nkinno_tinryo)

                rtn_str = xml_nkinkomk

                Return rtn_str

            End Function

            ''' <summary>
            ''' その他項目に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_xmlitem"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Other(ByVal hash_cvitem As Hashtable, ByVal hash_xmlitem As Hashtable) As String

                Dim rtn_str As String = ""
                Dim othercnt As Integer = 100
                Dim xml_other(othercnt) As String
                Dim cntindex As Integer = 1

                xml_other(cntindex) = "<Password>" & hash_cvitem("site_password") & "</Password>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<SiteID>" & hash_cvitem("site_id") & "</SiteID>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<SiteSendUmu>" & IIf(hash_cvitem("site_sendumu") = "1", "true", "false") & "</SiteSendUmu>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<UseCommonSetting>" & "false" & "</UseCommonSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<AssignedSyuhenKbn>" & "15" & "</AssignedSyuhenKbn>" : cntindex = cntindex + 1   '要確認
                xml_other(cntindex) = "<BantiDispSetting>" & hash_xmlitem("番地以降の表示") & "</BantiDispSetting>" : cntindex = cntindex + 1

                'グループ設定 sta
                Dim fldnamegrp As String = "所属グループ設定"
                Dim grpname As String = "grpname"
                Dim xml_grp As String = ""
                Dim tmp_strbgrp As New System.Text.StringBuilder
                For cntgrp = 1 To 10
                    tmp_strbgrp.Append("    <HomesSendSettingModel.BelongGroupModel>")
                    tmp_strbgrp.Append("      <__identity xmlns=" & """" & "http://schemas.datacontract.org/2004/07/System" & """" & " i:nil=" & """" & "true" & """" & " />")
                    tmp_strbgrp.Append("      <Id>" & grpname & "</Id>")
                    tmp_strbgrp.Append("      <IdNo>ID" & cntgrp.ToString & "</IdNo>")
                    tmp_strbgrp.Append("    </HomesSendSettingModel.BelongGroupModel>")
                    Dim tmp_grpstr As String = tmp_strbgrp.ToString
                    tmp_grpstr = tmp_grpstr.Replace(grpname, hash_xmlitem(fldnamegrp & cntgrp.ToString))
                    xml_grp = xml_grp & tmp_grpstr
                    tmp_strbgrp.Clear()
                Next
                '結合
                xml_other(cntindex) = "<BelongGroupList>" & xml_grp & "</BelongGroupList>" : cntindex = cntindex + 1
                'グループ設定 end

                xml_other(cntindex) = "<BkHyDispSetting>" & hash_xmlitem("物件部屋の表示") & "</BkHyDispSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<FtpPassword>" & hash_xmlitem("FTPパスワード") & "</FtpPassword>" : cntindex = cntindex + 1  '要追加
                xml_other(cntindex) = "<FtpUserId>" & hash_xmlitem("FTPユーザーID") & "</FtpUserId>" : cntindex = cntindex + 1       '要追加

                '固定分
                Dim tmp_strb As New System.Text.StringBuilder
                tmp_strb.Append("<GazoSettingList>")
                tmp_strb.Append("<HomesGazoSettingModel>")
                tmp_strb.Append("<FK8LocationKbn xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</FK8LocationKbn>")
                tmp_strb.Append("<FK8LocationNo xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</FK8LocationNo>")
                tmp_strb.Append("<No xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</No>")
                tmp_strb.Append("<PortalGazoSyubetu xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</PortalGazoSyubetu>")
                tmp_strb.Append("</HomesGazoSettingModel>")
                tmp_strb.Append("<HomesGazoSettingModel>")
                tmp_strb.Append("<FK8LocationKbn xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">2</FK8LocationKbn>")
                tmp_strb.Append("<FK8LocationNo xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">1</FK8LocationNo>")
                tmp_strb.Append("<No xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">16</No>")
                tmp_strb.Append("<PortalGazoSyubetu xmlns=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">5</PortalGazoSyubetu>")
                tmp_strb.Append("</HomesGazoSettingModel>")
                tmp_strb.Append("</GazoSettingList>")
                xml_other(cntindex) = tmp_strb.ToString : cntindex = cntindex + 1

                xml_other(cntindex) = "<IsAsteriskOsusumePointScore>" & hash_xmlitem("特別広告ポイント数") & "</IsAsteriskOsusumePointScore>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<IsAsteriskPanorama>" & hash_xmlitem("パノラマセット") & "</IsAsteriskPanorama>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<IsAsteriskStaffComment>" & hash_xmlitem("スタッフコメント") & "</IsAsteriskStaffComment>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<KeisaiKakuninBi>" & hash_xmlitem("掲載確認日") & "</KeisaiKakuninBi>" : cntindex = cntindex + 1

                '結合
                Dim tmp_str As String = ""
                For cntii = 1 To othercnt
                    If xml_other(cntii) Is Nothing Then
                        Exit For
                    End If
                    tmp_str = tmp_str & xml_other(cntii)
                Next

                rtn_str = tmp_str

                Return rtn_str

            End Function

            ''' <summary>
            ''' FTPユーザーID、パスワードの有無チェック
            ''' </summary>
            ''' <param name="tblname"></param>
            ''' <param name="hash_xmlitem"></param>
            ''' <param name="hash_log"></param>
            ''' <remarks></remarks>
            Public Sub Chk_FTPData(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_cvitem As Hashtable, ByRef hash_xmlitem As Hashtable)

                Dim ftpuserid As String = hash_xmlitem("FTPユーザーID")
                Dim ftppassword As String = hash_xmlitem("FTPパスワード")
                Dim log_key As String = tblname & "-" & "site_data"
                Dim errstr As String = ""

                'ログ出力用の項目設定
                Dim syorikomok As String = "送信設定HOMES情報"
                Dim taisyodata As String = "送信設定順 = " & hash_cvitem("setting_guid") & "、" & "サイトNo = " & hash_cvitem("site_no")

                'FTPユーザーIDチェック
                If ftpuserid = "" Then

                    '任意の文字列を格納
                    Dim tmp_ftpuserid As String = "ftpuserid"

                    '再格納
                    hash_xmlitem("FTPユーザーID") = tmp_ftpuserid

                    'ログ出力
                    Dim syoriitem As String = "FTPユーザーID"
                    Dim syorikekka As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED
                    Dim hubigein As String = LOG_HUBI_NOTEXISTDATA_REQUIRED_ADDSTR
                    Dim taisyo As String = LOG_TAISYO_DEFAULT
                    Dim beforechgvalue As String = ftpuserid
                    Dim afterchgvalue As String = tmp_ftpuserid

                    'ログ挿入用へ加工
                    Dim logitem As String = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname)
                    '挿入用クエリへ加工
                    Dim log_sql As String = LogSetting.Get_LogTblInsertQry(logitem, False)
                    'ログ挿入処理
                    Dim tmp_cnt As Integer = 0
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, tmp_cnt)

                End If

                'FTPパスワードチェック
                If ftppassword = "" Then

                    '任意の文字列を格納
                    Dim tmp_ftppassword As String = "ftppassword"

                    '再格納
                    hash_xmlitem("FTPパスワード") = tmp_ftppassword

                    'ログ出力
                    Dim syoriitem As String = "FTPパスワード"
                    Dim syorikekka As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED
                    Dim hubigein As String = LOG_HUBI_NOTEXISTDATA_REQUIRED_ADDSTR
                    Dim taisyo As String = LOG_TAISYO_DEFAULT
                    Dim beforechgvalue As String = ftppassword
                    Dim afterchgvalue As String = tmp_ftppassword

                    'ログ挿入用へ加工
                    Dim logitem As String = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname)
                    '挿入用クエリへ加工
                    Dim log_sql As String = LogSetting.Get_LogTblInsertQry(logitem, False)
                    'ログ挿入処理
                    Dim tmp_cnt As Integer = 0
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, tmp_cnt)

                End If

            End Sub

        End Class

    End Class

#End Region

#Region "送信設定athome情報"

    Public Class M_site_sendsetting_athome_Repository

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
                Dim hash_settingguid As New Hashtable                           'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.M_site_sendsetting_Model      '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Nkinkomk()

                '革命10の入金項目マスタを取得
                Dim hash_nkinkomkmst As New Hashtable
                Dim tmp_sql As String = "SELECT nkin_no,nkin_name FROM m_nkin"
                Call DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_nkinkomkmst)

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                '送信設定guid取得
                Call Me.Set_SettingGuid(sqlcnnv10, hash_settingguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_site_sendsetting"
                Dim fldnamegrp As String = "sitesetting_guid,setting_guid,site_no,site_id,site_sendumu," & _
                                           "site_data,site_password"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_settingguid, list_basekeydata)

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
                        Dim hash_cvxmlitem As New Hashtable

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "送信設定順"
                                    .Vari_Setting_guid = fldvalue
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case "ポータルサイトID"
                                    .Vari_Site_id = fldvalue
                                Case "サイト別送信有無"
                                    .Vari_Site_sendumu = fldvalue
                                Case "ポータルサイトパスワード"
                                    .Vari_Site_password = IIf(fldvalue = "", " ", fldvalue)
                                Case Else
                                    If fldname <> "" Then
                                        hash_cvxmlitem.Add(fldname, fldvalue)
                                    End If
                            End Select

                        Next

                        '固定値
                        .Vari_Sitesetting_guid = Guid.NewGuid.ToString

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '会社支店コード桁数チェック
                        Call Me.Chk_Tenpocode(sqlcnnv10, tblname, hash_cvitem, hash_cvxmlitem)

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            hash_cvitem("setting_guid") = hash_settingguid.Item(StrConv(hash_cvitem.Item("setting_guid"), VbStrConv.Narrow))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                'xmlの移行
                                Call Me.Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem, hash_nkinkomkmst)

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

                'ポータルサイトパスワード一括更新 (半角スペース " " を設定していた分を空文字へ変換)
                Dim tmpcnt As Integer = 0
                Dim passwordupdateqry As String = " UPDATE " & tblname & " SET site_password = '' WHERE site_no = 30 AND site_password = ' ' "
                DBExec.Exec_NonQuery(sqlcnnv10, passwordupdateqry, tmpcnt)

                '共通設定情報へ個別設定No1を反映
                '20160908 個別処理を共通処理へコピーする処理の制御修正 -chg sta
                'Dim tmp_sortstr As String = ""  '使用しない
                'Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                'DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                If cvrowcnt <> 0 Then
                    Dim tmp_sortstr As String = ""  '使用しない
                    Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                    DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                End If
                '20160908 個別処理を共通処理へコピーする処理の制御修正 -chg end
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
            ''' 共通設定情報更新クエリの作成
            ''' </summary>
            ''' <param name="sortstr"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " UPDATE m_site_sendsetting SET "
                tmp_sql = tmp_sql & " 	 site_id = (SELECT site_id FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 30) "
                tmp_sql = tmp_sql & " 	,site_sendumu = (SELECT site_sendumu FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 30) "
                tmp_sql = tmp_sql & " 	,site_data = (SELECT site_data FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 30) "
                tmp_sql = tmp_sql & " 	,site_password = (SELECT site_password FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 30) "
                tmp_sql = tmp_sql & " WHERE setting_guid = "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0 "
                tmp_sql = tmp_sql & " ) "
                tmp_sql = tmp_sql & " AND site_no = 30 "

                Return tmp_sql

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) "
                tmp_sql = tmp_sql & " FROM hydata_kokoku AS HYKOKOKU "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 送信設定guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_SettingGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = " SELECT setting_sortorder,setting_guid FROM m_sendsetting "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 送信設定のxml項目の移行
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <remarks></remarks>
            Public Sub Set_XmlItem(ByVal sqlcnnv10 As SqlConnection, ByVal hash_cvitem As Hashtable, ByVal hash_xmlitem As Hashtable, ByVal hash_nkinkomkmst As Hashtable)

                Dim sitesettingguid As String = hash_cvitem("sitesetting_guid")   '更新対象選択用

                '「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                Dim xmlupdateqry_xmlpre As String = ""
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre & "<AthomeSendSettingModel xmlns:i=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & ">"
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre & "<__identity xmlns=" & """" & "http://schemas.datacontract.org/2004/07/System" & """" & " i:nil=" & """" & "true" & """" & " />"
                Dim xmlupdateqry_xmlpost As String = ""
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost & "</AthomeSendSettingModel>"

                'グループ毎にXml文字列を作成
                '画像
                Dim tmp_xml_gazo As String = Me.Get_XmlStr_Gazo(hash_xmlitem)
                '入金項目
                Dim tmp_xml_nkin As String = Me.Get_XmlStr_Nkinkomk(hash_xmlitem, hash_nkinkomkmst)
                'その他
                Dim tmp_xml_other As String = Me.Get_XmlStr_Other(hash_cvitem, hash_xmlitem)

                'Xml文字列の結合
                Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & tmp_xml_gazo & tmp_xml_nkin & tmp_xml_other & xmlupdateqry_xmlpost

                '更新用クエリの作成
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " DECLARE @tmp_xml XML "
                tmp_sql = tmp_sql & " SET @tmp_xml = CAST('" & xmlupdateqry_main & "' AS XML) "
                tmp_sql = tmp_sql & " UPDATE m_site_sendsetting SET site_data = @tmp_xml "
                tmp_sql = tmp_sql & " WHERE sitesetting_guid = '" & sitesettingguid & "' "

                '更新
                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

            ''' <summary>
            ''' 画像に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Gazo(ByVal hash As Hashtable) As String

                Dim rtn_str As String = ""
                Dim xml_pre As String = "<GazoSettingList xmlns:a=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">"
                Dim xml_post As String = "</GazoSettingList>"

                Dim gazocnt As Integer = 16
                Dim tmp_gazono(gazocnt) As String
                Dim tmp_gazosyuV7(gazocnt) As String
                Dim tmp_gazonoV7(gazocnt) As String
                Dim tmp_gazosyurendo(gazocnt) As String

                '画像に関する情報を配列へ格納
                For Each item In hash

                    Dim fldname As String = item.Key
                    Dim value As String = item.Value

                    Select Case True
                        Case fldname.Replace("画像No", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("画像No", ""))
                            tmp_gazono(gazono) = value
                        Case fldname.Replace("革命側の画像種別", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("革命側の画像種別", ""))
                            tmp_gazosyuV7(gazono) = value
                        Case fldname.Replace("革命側の画像タイトルNo", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("革命側の画像タイトルNo", ""))
                            tmp_gazonoV7(gazono) = value
                        Case fldname.Replace("連動側の画像種別", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("連動側の画像種別", ""))
                            tmp_gazosyurendo(gazono) = value
                    End Select

                Next

                '文字列へ成形
                Dim xml_gazoitem As New System.Text.StringBuilder
                For cntii = 1 To gazocnt
                    If tmp_gazosyurendo(cntii) <> "0" And tmp_gazosyuV7(cntii) <> "0" And tmp_gazonoV7(cntii) <> "0" Then
                        xml_gazoitem.Append("<a:PortalGazoSettingModel>")
                        xml_gazoitem.Append("<a:FK8LocationKbn>" & tmp_gazosyuV7(cntii) & "</a:FK8LocationKbn>")
                        xml_gazoitem.Append("<a:FK8LocationNo>" & tmp_gazonoV7(cntii) & "</a:FK8LocationNo>")
                        xml_gazoitem.Append("<a:No>" & tmp_gazono(cntii) & "</a:No>")
                        xml_gazoitem.Append("<a:PortalGazoSyubetu>" & tmp_gazosyurendo(cntii) & "</a:PortalGazoSyubetu>")
                        xml_gazoitem.Append("</a:PortalGazoSettingModel>")
                    End If
                Next

                '固定部分の追加
                Dim xml_gazokotei As New System.Text.StringBuilder
                Dim koteicnt As Integer = 8
                For cntjj = 1 To koteicnt
                    xml_gazokotei.Append("<a:PortalGazoSettingModel>")
                    xml_gazokotei.Append("<a:FK8LocationKbn>" & "10" & "</a:FK8LocationKbn>")
                    xml_gazokotei.Append("<a:FK8LocationNo>" & cntjj.ToString & "</a:FK8LocationNo>")
                    xml_gazokotei.Append("<a:No>" & "10" & cntjj.ToString & "</a:No>")
                    xml_gazokotei.Append("<a:PortalGazoSyubetu>" & "-1" & "</a:PortalGazoSyubetu>")
                    xml_gazokotei.Append("</a:PortalGazoSettingModel>")
                Next

                '結合
                Dim xml_gazo As String = xml_pre & xml_gazoitem.ToString & xml_gazokotei.ToString & xml_post
                rtn_str = xml_gazo

                Return rtn_str

            End Function

            ''' <summary>
            ''' 入金項目に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Nkinkomk(ByVal hash_xmlitem As Hashtable, ByVal hash_nkinkomkmst As Hashtable) As String

                Dim rtn_str As String = ""
                Dim xml_pre As String = "<NkinKomokDict xmlns:a=" & """" & "http://schemas.microsoft.com/2003/10/Serialization/Arrays" & """" & ">"
                Dim xml_post As String = "</NkinKomokDict>"
                Dim nkincnt As Integer = 16
                Dim cntindex As Integer = 1
                Dim xml_nkinkomk(nkincnt) As String
                Dim hash_nkinkomkitem As New Hashtable

                '入金項目を移行用に成形する
                Dim hash_cvxmlitem As New Hashtable
                hash_cvxmlitem = Me.Chg_NkinKomk(hash_xmlitem, hash_nkinkomkmst)

                '賃料
                hash_nkinkomkitem.Add("item_key", "1")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_賃料"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_賃料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_賃料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '共益費
                hash_nkinkomkitem.Add("item_key", "3")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_共益費"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_共益費")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_共益費No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '管理費
                hash_nkinkomkitem.Add("item_key", "4")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_管理費"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_管理費")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_管理費No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '敷金
                hash_nkinkomkitem.Add("item_key", "6")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_敷金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("敷金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_敷金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_敷金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '礼金
                hash_nkinkomkitem.Add("item_key", "5")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_礼金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("礼金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_礼金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_礼金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '保証金
                hash_nkinkomkitem.Add("item_key", "7")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_保証金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("保証金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_保証金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_保証金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '敷引金
                hash_nkinkomkitem.Add("item_key", "12")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_敷引金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("敷引金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_敷引金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_敷引金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '保証金償却
                hash_nkinkomkitem.Add("item_key", "9")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_保証金償却"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("保証金償却表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_償却金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_保証金償却No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '更新料
                hash_nkinkomkitem.Add("item_key", "14")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_更新料"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("更新料表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_更新料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_更新料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '仲介手数料
                hash_nkinkomkitem.Add("item_key", "15")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_仲介手数料"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("仲介手数料表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_仲介手数料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_仲介手数料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '雑費
                hash_nkinkomkitem.Add("item_key", "19")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_雑費"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_雑費")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_雑費No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '鍵交換代等
                hash_nkinkomkitem.Add("item_key", "20")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_鍵交換代等"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_鍵交換代等")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_鍵交換代等No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                'その他一時金
                hash_nkinkomkitem.Add("item_key", "18")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_その他一時金"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_その他一時金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_その他一時金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                'その他月額費用
                hash_nkinkomkitem.Add("item_key", "17")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_その他月額費用"))
                hash_nkinkomkitem.Add("item_tani", "_円")
                hash_nkinkomkitem.Add("item_komk", "_その他費用")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_その他月額費用No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '駐車場敷金
                hash_nkinkomkitem.Add("item_key", "21")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_駐車場敷金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("駐車場敷金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_駐車場敷金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_駐車場敷金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '駐車場礼金
                hash_nkinkomkitem.Add("item_key", "22")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_駐車場礼金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("駐車場礼金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_駐車場礼金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_駐車場礼金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '結合
                Dim xml_total As String = ""
                Dim xml_nkintotal As String = ""
                For cntii = 1 To nkincnt
                    xml_nkintotal = xml_nkintotal & xml_nkinkomk(cntii)
                Next
                xml_total = xml_pre & xml_nkintotal & xml_post

                rtn_str = xml_total

                Return rtn_str

            End Function

            ''' <summary>
            ''' 入金項目の編集
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_xml"></param>
            ''' <param name="hash_nkinkomkmst"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chg_NkinKomk(ByVal hash_xml As Hashtable, ByVal hash_nkinkomkmst As Hashtable) As Hashtable

                Dim rtn_hash As New Hashtable
                Dim tmp_hash As New Hashtable

                Dim tmp_nkincnt As Integer = 32
                Dim tmp_nkinkomkname(tmp_nkincnt) As String
                Dim tmp_nkinkomkkbn(tmp_nkincnt) As String
                Dim tmp_fldname(tmp_nkincnt) As String
                '-----------------------
                '入金項目取得
                '-----------------------
                For Each item In hash_xml

                    Dim midfldname As String = item.Key
                    Dim value As String = item.Value

                    Select Case midfldname
                        Case "入金項目_賃料1" : tmp_nkinkomkname(1) = value : tmp_fldname(1) = midfldname
                        Case "入金項目区分_賃料1" : tmp_nkinkomkkbn(1) = value
                        Case "入金項目_賃料2" : tmp_nkinkomkname(2) = value : tmp_fldname(2) = midfldname
                        Case "入金項目区分_賃料2" : tmp_nkinkomkkbn(2) = value
                        Case "入金項目_賃料3" : tmp_nkinkomkname(3) = value : tmp_fldname(3) = midfldname
                        Case "入金項目区分_賃料3" : tmp_nkinkomkkbn(3) = value
                        Case "入金項目_共益費1" : tmp_nkinkomkname(4) = value : tmp_fldname(4) = midfldname
                        Case "入金項目区分_共益費1" : tmp_nkinkomkkbn(4) = value
                        Case "入金項目_共益費2" : tmp_nkinkomkname(5) = value : tmp_fldname(5) = midfldname
                        Case "入金項目区分_共益費2" : tmp_nkinkomkkbn(5) = value
                        Case "入金項目_共益費3" : tmp_nkinkomkname(6) = value : tmp_fldname(6) = midfldname
                        Case "入金項目区分_共益費3" : tmp_nkinkomkkbn(6) = value
                        Case "入金項目_管理費1" : tmp_nkinkomkname(7) = value : tmp_fldname(7) = midfldname
                        Case "入金項目区分_管理費1" : tmp_nkinkomkkbn(7) = value
                        Case "入金項目_管理費2" : tmp_nkinkomkname(8) = value : tmp_fldname(8) = midfldname
                        Case "入金項目区分_管理費2" : tmp_nkinkomkkbn(8) = value
                        Case "入金項目_管理費3" : tmp_nkinkomkname(9) = value : tmp_fldname(9) = midfldname
                        Case "入金項目区分_管理費3" : tmp_nkinkomkkbn(9) = value
                        Case "入金項目_敷金" : tmp_nkinkomkname(10) = value : tmp_fldname(10) = midfldname
                        Case "入金項目区分_敷金" : tmp_nkinkomkkbn(10) = value
                        Case "入金項目_礼金" : tmp_nkinkomkname(11) = value : tmp_fldname(11) = midfldname
                        Case "入金項目区分_礼金" : tmp_nkinkomkkbn(11) = value
                        Case "入金項目_保証金" : tmp_nkinkomkname(12) = value : tmp_fldname(12) = midfldname
                        Case "入金項目区分_保証金" : tmp_nkinkomkkbn(12) = value
                        Case "入金項目_保証金償却" : tmp_nkinkomkname(13) = value : tmp_fldname(13) = midfldname
                        Case "入金項目区分_保証金償却" : tmp_nkinkomkkbn(13) = value
                        Case "入金項目_敷引金" : tmp_nkinkomkname(14) = value : tmp_fldname(14) = midfldname
                        Case "入金項目区分_敷引金" : tmp_nkinkomkkbn(14) = value
                        Case "入金項目_更新料" : tmp_nkinkomkname(15) = value : tmp_fldname(15) = midfldname
                        Case "入金項目区分_更新料" : tmp_nkinkomkkbn(15) = value
                        Case "入金項目_仲介手数料" : tmp_nkinkomkname(16) = value : tmp_fldname(16) = midfldname
                        Case "入金項目区分_仲介手数料" : tmp_nkinkomkkbn(16) = value
                        Case "入金項目_鍵交換代等" : tmp_nkinkomkname(17) = value : tmp_fldname(17) = midfldname
                        Case "入金項目区分_鍵交換代等" : tmp_nkinkomkkbn(17) = value
                        Case "入金項目_雑費1" : tmp_nkinkomkname(18) = value : tmp_fldname(18) = midfldname
                        Case "入金項目区分_雑費1" : tmp_nkinkomkkbn(18) = value
                        Case "入金項目_雑費2" : tmp_nkinkomkname(19) = value : tmp_fldname(19) = midfldname
                        Case "入金項目区分_雑費2" : tmp_nkinkomkkbn(19) = value
                        Case "入金項目_雑費3" : tmp_nkinkomkname(20) = value : tmp_fldname(20) = midfldname
                        Case "入金項目区分_雑費3" : tmp_nkinkomkkbn(20) = value
                        Case "入金項目_駐車場敷金" : tmp_nkinkomkname(21) = value : tmp_fldname(21) = midfldname
                        Case "入金項目区分_駐車場敷金" : tmp_nkinkomkkbn(21) = value
                        Case "入金項目_駐車場礼金" : tmp_nkinkomkname(22) = value : tmp_fldname(22) = midfldname
                        Case "入金項目区分_駐車場礼金" : tmp_nkinkomkkbn(22) = value
                        Case "入金項目_その他月額費用1" : tmp_nkinkomkname(23) = value : tmp_fldname(23) = midfldname
                        Case "入金項目区分_その他月額費用1" : tmp_nkinkomkkbn(23) = value
                        Case "入金項目_その他月額費用2" : tmp_nkinkomkname(24) = value : tmp_fldname(24) = midfldname
                        Case "入金項目区分_その他月額費用2" : tmp_nkinkomkkbn(24) = value
                        Case "入金項目_その他月額費用3" : tmp_nkinkomkname(25) = value : tmp_fldname(25) = midfldname
                        Case "入金項目区分_その他月額費用3" : tmp_nkinkomkkbn(25) = value
                        Case "入金項目_その他月額費用4" : tmp_nkinkomkname(26) = value : tmp_fldname(26) = midfldname
                        Case "入金項目区分_その他月額費用4" : tmp_nkinkomkkbn(26) = value
                        Case "入金項目_その他月額費用5" : tmp_nkinkomkname(27) = value : tmp_fldname(27) = midfldname
                        Case "入金項目区分_その他月額費用5" : tmp_nkinkomkkbn(27) = value
                        Case "入金項目_その他一時金1" : tmp_nkinkomkname(28) = value : tmp_fldname(28) = midfldname
                        Case "入金項目区分_その他一時金1" : tmp_nkinkomkkbn(28) = value
                        Case "入金項目_その他一時金2" : tmp_nkinkomkname(29) = value : tmp_fldname(29) = midfldname
                        Case "入金項目区分_その他一時金2" : tmp_nkinkomkkbn(29) = value
                        Case "入金項目_その他一時金3" : tmp_nkinkomkname(30) = value : tmp_fldname(30) = midfldname
                        Case "入金項目区分_その他一時金3" : tmp_nkinkomkkbn(30) = value
                        Case "入金項目_その他一時金4" : tmp_nkinkomkname(31) = value : tmp_fldname(31) = midfldname
                        Case "入金項目区分_その他一時金4" : tmp_nkinkomkkbn(31) = value
                        Case "入金項目_その他一時金5" : tmp_nkinkomkname(32) = value : tmp_fldname(32) = midfldname
                        Case "入金項目区分_その他一時金5" : tmp_nkinkomkkbn(32) = value
                        Case Else
                            tmp_hash.Add(midfldname, value)
                    End Select

                Next

                '-----------------------
                '入金項目をXml移行用に成形
                '-----------------------
                Dim tmp_nkinno(tmp_nkincnt) As String
                Dim tmp_dispnkinname(tmp_nkincnt) As String
                For cntii = 1 To tmp_nkincnt

                    Dim tmp_relnkinstr As String = tmp_nkinkomkname(cntii) & "-" & EtcMethod.Get_Nkinruiname(tmp_nkinkomkkbn(cntii))
                    tmp_nkinno(cntii) = ""
                    tmp_dispnkinname(cntii) = ""

                    If Hash_Rel_Nkinkomk.Contains(tmp_relnkinstr) Then
                        tmp_nkinno(cntii) = Hash_Rel_Nkinkomk(tmp_relnkinstr)
                        tmp_dispnkinname(cntii) = tmp_nkinno(cntii) & ":" & hash_nkinkomkmst(tmp_nkinno(cntii))
                    End If

                Next

                '-----------------------
                '入金項目の再格納
                '-----------------------
                '結合する項目
                Call Me.Set_NkinkomkAddToHash(1, 3, tmp_nkinno, tmp_dispnkinname, "入金項目_賃料", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(4, 6, tmp_nkinno, tmp_dispnkinname, "入金項目_共益費", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(7, 9, tmp_nkinno, tmp_dispnkinname, "入金項目_管理費", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(18, 20, tmp_nkinno, tmp_dispnkinname, "入金項目_雑費", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(23, 27, tmp_nkinno, tmp_dispnkinname, "入金項目_その他月額費用", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(28, 32, tmp_nkinno, tmp_dispnkinname, "入金項目_その他一時金", tmp_hash)

                '結合不要の項目
                For cntii = 10 To 22
                    tmp_hash.Add(tmp_fldname(cntii), tmp_dispnkinname(cntii))
                    tmp_hash.Add(tmp_fldname(cntii) & "No", tmp_nkinno(cntii))
                Next

                rtn_hash = tmp_hash

                Return rtn_hash

            End Function

            ''' <summary>
            ''' 入金項目を結合して1つのデータにまとめる
            ''' まとめたデータをハッシュテーブルへ格納する
            ''' </summary>
            ''' <param name="cntsta"></param>
            ''' <param name="cntend"></param>
            ''' <param name="tmp_nkinno"></param>
            ''' <param name="tmp_dispnkinname"></param>
            ''' <param name="hashkey"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_NkinkomkAddToHash(ByVal cntsta As Integer, ByVal cntend As Integer, ByVal tmp_nkinno As Object, ByVal tmp_dispnkinname As Object, ByVal hashkey As String, ByRef hash As Hashtable)

                Dim nkinno As String = ""
                Dim dispname As String = ""
                For cntii = cntsta To cntend
                    If tmp_nkinno(cntii) <> "" Then
                        nkinno = nkinno & "," & tmp_nkinno(cntii)
                        dispname = dispname & " / " & tmp_dispnkinname(cntii)
                    End If
                Next
                If nkinno <> "" Then
                    nkinno = nkinno.Remove(0, 1)
                End If
                If dispname <> "" Then
                    dispname = dispname.Remove(0, 3)
                End If
                hash.Add(hashkey, dispname)
                hash.Add(hashkey & "No", nkinno)

            End Sub

            ''' <summary>
            ''' 各入金項目毎にXml文字列作成
            ''' </summary>
            ''' <param name="hash_nkin"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chg_NkinkomkToXml(ByVal hash_nkin As Hashtable) As String

                Dim rtn_str As String = ""

                '置換用/ハッシュテーブル値取得用文字列
                Dim item_key As String = "item_key"
                Dim item_disp As String = "item_disp"
                Dim item_tani As String = "item_tani"
                Dim item_komk As String = "item_komk"
                Dim item_nkinno As String = "item_nkinno"

                'Xmlのベースを作成
                Dim xml_nkinkomkbase As New System.Text.StringBuilder
                xml_nkinkomkbase.Append("<a:KeyValueOfintPortalNkinKomokLjh4bohd>")
                xml_nkinkomkbase.Append("<a:Key>" & item_key & "</a:Key>")
                xml_nkinkomkbase.Append("<a:Value>")
                xml_nkinkomkbase.Append("<DisplayString>" & item_disp & "</DisplayString>")
                xml_nkinkomkbase.Append("<EnKagetu>" & item_tani & "</EnKagetu>")
                xml_nkinkomkbase.Append("<NkinKomok>" & item_komk & "</NkinKomok>")
                xml_nkinkomkbase.Append("<NkinNoList>")
                xml_nkinkomkbase.Append(item_nkinno)
                xml_nkinkomkbase.Append("</NkinNoList>")
                xml_nkinkomkbase.Append("</a:Value>")
                xml_nkinkomkbase.Append("</a:KeyValueOfintPortalNkinKomokLjh4bohd>")

                '各要素で置換してい成形
                Dim xml_nkinkomk As String = xml_nkinkomkbase.ToString
                xml_nkinkomk = xml_nkinkomk.Replace(item_key, hash_nkin(item_key))
                xml_nkinkomk = xml_nkinkomk.Replace(item_disp, hash_nkin(item_disp))
                xml_nkinkomk = xml_nkinkomk.Replace(item_tani, hash_nkin(item_tani))
                xml_nkinkomk = xml_nkinkomk.Replace(item_komk, hash_nkin(item_komk))
                Dim tmp_nkinnogrp_tinryo As String = hash_nkin(item_nkinno)
                Dim tmp_nkinno_tinryo() As String = tmp_nkinnogrp_tinryo.Split(",")
                Dim nkinno_tinryo As String = ""
                For cntii = 0 To UBound(tmp_nkinno_tinryo)
                    If tmp_nkinno_tinryo(cntii) <> "" Then
                        nkinno_tinryo = nkinno_tinryo & "<a:int>" & tmp_nkinno_tinryo(cntii) & "</a:int>"
                    End If
                Next
                xml_nkinkomk = xml_nkinkomk.Replace(item_nkinno, nkinno_tinryo)

                rtn_str = xml_nkinkomk

                Return rtn_str

            End Function

            ''' <summary>
            ''' その他項目に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_xmlitem"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Other(ByVal hash_cvitem As Hashtable, ByVal hash_xmlitem As Hashtable) As String

                Dim rtn_str As String = ""
                Dim othercnt As Integer = 100
                Dim xml_other(othercnt) As String
                Dim cntindex As Integer = 1

                xml_other(cntindex) = "<Password>" & hash_cvitem("site_password") & "</Password>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<SiteID>" & hash_cvitem("site_id") & "</SiteID>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<SiteSendUmu>" & IIf(hash_cvitem("site_sendumu") = "1", "true", "false") & "</SiteSendUmu>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<UseCommonSetting>" & "false" & "</UseCommonSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<BantiDispSetting>" & hash_xmlitem("番地以降の表示") & "</BantiDispSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<BkHyDispSetting>" & hash_xmlitem("物件名・部屋Noの表示(一般ユーザー向け)") & "</BkHyDispSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<BkHyDispSettingKaiin>" & hash_xmlitem("物件名・部屋Noの表示(不動産会社向け)") & "</BkHyDispSettingKaiin>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<KeisaiKakuninBi>" & hash_xmlitem("掲載確認日") & "</KeisaiKakuninBi>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<MapDispSetting>" & hash_xmlitem("地図表示") & "</MapDispSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<TenpoCode>" & hash_xmlitem("店舗ID") & "</TenpoCode>" : cntindex = cntindex + 1

                '結合
                Dim tmp_str As String = ""
                For cntii = 1 To othercnt
                    If xml_other(cntii) Is Nothing Then
                        Exit For
                    End If
                    tmp_str = tmp_str & xml_other(cntii)
                Next

                rtn_str = tmp_str

                Return rtn_str

            End Function

            ''' <summary>
            ''' 店舗コードの桁数(4桁)チェック
            ''' V7の店舗IDはathomeIDと店舗コード(下4桁)が結合された状態になっている
            ''' 抽出時に下4桁を店舗ID、残りをathomeIDとして取得している
            ''' </summary>
            ''' <param name="tblname"></param>
            ''' <param name="hash_xmlimte"></param>
            ''' <param name="hash_log"></param>
            ''' <remarks></remarks>
            Public Sub Chk_Tenpocode(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_cvitem As Hashtable, ByRef hash_xmlitem As Hashtable)

                Dim errstr As String = ""
                Dim log_key As String = tblname & "-" & "site_data"
                Dim athomeid As String = hash_cvitem("site_id")
                Dim tenpocode As String = hash_xmlitem("店舗ID")

                'ログ出力用の項目設定
                Dim syorikomok As String = "送信設定athome情報"
                Dim taisyodata As String = "送信設定順 = " & hash_cvitem("setting_guid") & "、" & "サイトNo = " & hash_cvitem("site_no")

                '桁数チェック
                If tenpocode.Length < 4 Then

                    '4桁へ成形
                    Dim tmp_tenpocode As String = tenpocode
                    tmp_tenpocode = "0000" & tmp_tenpocode
                    tmp_tenpocode = tmp_tenpocode.Substring(tmp_tenpocode.Length - 4)

                    '再格納
                    hash_xmlitem("店舗ID") = tmp_tenpocode

                    'ログ出力
                    Dim syoriitem As String = "店舗ID"
                    Dim syorikekka As String = LOG_NAIYO_ERR_OUTOFRANGE_STR
                    Dim hubigein As String = LOG_HUBI_OUTOFRANGE_STR_SHORTAGE
                    Dim taisyo As String = LOG_TAISYO_OUTOOFRANGE_STR
                    Dim beforechgvalue As String = tenpocode
                    Dim afterchgvalue As String = tmp_tenpocode

                    'ログ挿入用へ加工
                    Dim logitem As String = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname)
                    '挿入用クエリへ加工
                    Dim log_sql As String = LogSetting.Get_LogTblInsertQry(logitem, False)
                    'ログ挿入処理
                    Dim tmp_cnt As Integer = 0
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, tmp_cnt)

                End If

                '-----------------------------------------------
                'athomeIDが空の場合の対応
                '※tenpocodeが4桁以下の場合は空文字になっている
                '※念の為空文字チェックは行っておく
                '-----------------------------------------------
                If athomeid = "" Then

                    '任意の文字列を格納
                    Dim tmp_athomeid As String = "athomeid"

                    '再格納
                    hash_cvitem("site_id") = tmp_athomeid

                    'ログ出力
                    Dim syoriitem As String = "athomeID"
                    Dim syorikekka As String = LOG_NAIYO_ERR_OUTOFRANGE_STR
                    Dim hubigein As String = LOG_HUBI_OUTOFRANGE_STR_SHORTAGE
                    Dim taisyo As String = LOG_TAISYO_OUTOOFRANGE_STR
                    Dim beforechgvalue As String = athomeid
                    Dim afterchgvalue As String = tmp_athomeid

                    'ログ挿入用へ加工
                    Dim logitem As String = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname)
                    '挿入用クエリへ加工
                    Dim log_sql As String = LogSetting.Get_LogTblInsertQry(logitem, False)
                    'ログ挿入処理
                    Dim tmp_cnt As Integer = 0
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, tmp_cnt)

                End If

            End Sub

        End Class

    End Class

#End Region

#Region "送信設定SUUMO情報"

    Public Class M_site_sendsetting_suumo_Repository

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
                Dim hash_settingguid As New Hashtable                           'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.M_site_sendsetting_Model      '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブ1キー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Nkinkomk()

                '革命10の入金項目マスタを取得
                Dim hash_nkinkomkmst As New Hashtable
                Dim tmp_sql As String = "SELECT nkin_no,nkin_name FROM m_nkin"
                Call DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_nkinkomkmst)

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                '送信設定guid取得
                Call Me.Set_SettingGuid(sqlcnnv10, hash_settingguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_site_sendsetting"
                Dim fldnamegrp As String = "sitesetting_guid,setting_guid,site_no,site_id,site_sendumu," & _
                                           "site_data,site_password"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_settingguid, list_basekeydata)

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
                        Dim hash_cvxmlitem As New Hashtable

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "送信設定順"
                                    .Vari_Setting_guid = fldvalue
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case "ポータルサイトID"
                                    .Vari_Site_id = fldvalue
                                Case "サイト別送信有無"
                                    .Vari_Site_sendumu = fldvalue
                                Case "ポータルサイトパスワード"
                                    .Vari_Site_password = IIf(fldvalue = "", " ", fldvalue)
                                Case Else
                                    If fldname <> "" Then
                                        hash_cvxmlitem.Add(fldname, fldvalue)
                                    End If
                            End Select

                        Next

                        '固定値
                        .Vari_Sitesetting_guid = Guid.NewGuid.ToString

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '会社支店コード桁数チェック
                        Call Me.Chk_Kaisyasitencode(sqlcnnv10, tblname, hash_cvitem)

                        '書込処理
                        If Not skipflg Then

                            'キーをguidへ変換
                            hash_cvitem("setting_guid") = hash_settingguid.Item(StrConv(hash_cvitem.Item("setting_guid"), VbStrConv.Narrow))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                'xmlの移行
                                Call Me.Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem, hash_nkinkomkmst)

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

                'ポータルサイトパスワード一括更新 (半角スペース " " を設定していた分を空文字へ変換)
                Dim tmpcnt As Integer = 0
                Dim passwordupdateqry As String = " UPDATE " & tblname & " SET site_password = '' WHERE site_no = 40 AND site_password = ' ' "
                DBExec.Exec_NonQuery(sqlcnnv10, passwordupdateqry, tmpcnt)

                '共通設定情報へ個別設定No1を反映
                '20160908 個別処理を共通処理へコピーする処理の制御修正 -chg sta
                'Dim tmp_sortstr As String = ""  '使用しない
                'Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                'DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                If cvrowcnt <> 0 Then
                    Dim tmp_sortstr As String = ""  '使用しない
                    Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                    DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                End If
                '20160908 個別処理を共通処理へコピーする処理の制御修正 -chg end

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
            ''' 共通設定情報更新クエリの作成
            ''' </summary>
            ''' <param name="sortstr"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " UPDATE m_site_sendsetting SET "
                tmp_sql = tmp_sql & " 	 site_id = (SELECT site_id FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 40) "
                tmp_sql = tmp_sql & " 	,site_sendumu = (SELECT site_sendumu FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 40) "
                tmp_sql = tmp_sql & " 	,site_data = (SELECT site_data FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 40) "
                tmp_sql = tmp_sql & " 	,site_password = (SELECT site_password FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 40) "
                tmp_sql = tmp_sql & " WHERE setting_guid = "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0 "
                tmp_sql = tmp_sql & " ) "
                tmp_sql = tmp_sql & " AND site_no = 40 "

                Return tmp_sql

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) "
                tmp_sql = tmp_sql & " FROM hydata_kokoku AS HYKOKOKU "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 送信設定guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_SettingGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = " SELECT setting_sortorder,setting_guid FROM m_sendsetting "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 送信設定のxml項目の移行
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <remarks></remarks>
            Public Sub Set_XmlItem(ByVal sqlcnnv10 As SqlConnection, ByVal hash_cvitem As Hashtable, ByVal hash_xmlitem As Hashtable, ByVal hash_nkinkomkmst As Hashtable)

                Dim sitesettingguid As String = hash_cvitem("sitesetting_guid")   '更新対象選択用

                '「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                Dim xmlupdateqry_xmlpre As String = ""
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre & "<SuumoSendSettingModel xmlns:i=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & ">"
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre & "<__identity xmlns=" & """" & "http://schemas.datacontract.org/2004/07/System" & """" & " i:nil=" & """" & "true" & """" & " />"
                Dim xmlupdateqry_xmlpost As String = ""
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost & "</SuumoSendSettingModel>"

                'グループ毎にXml文字列を作成
                '画像
                Dim tmp_xml_gazo As String = Me.Get_XmlStr_Gazo(hash_xmlitem)
                '入金項目
                Dim tmp_xml_nkin As String = Me.Get_XmlStr_Nkinkomk(hash_xmlitem, hash_nkinkomkmst)
                'その他
                Dim tmp_xml_other As String = Me.Get_XmlStr_Other(hash_cvitem, hash_xmlitem)

                'Xml文字列の結合
                Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & tmp_xml_gazo & tmp_xml_nkin & tmp_xml_other & xmlupdateqry_xmlpost

                '更新用クエリの作成
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " DECLARE @tmp_xml XML "
                tmp_sql = tmp_sql & " SET @tmp_xml = CAST('" & xmlupdateqry_main & "' AS XML) "
                tmp_sql = tmp_sql & " UPDATE m_site_sendsetting SET site_data = @tmp_xml "
                tmp_sql = tmp_sql & " WHERE sitesetting_guid = '" & sitesettingguid & "' "

                '更新
                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

            ''' <summary>
            ''' 画像に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Gazo(ByVal hash As Hashtable) As String

                Dim rtn_str As String = ""
                Dim xml_pre As String = "<GazoSettingList xmlns:a=" & """" & "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" & """" & ">"
                Dim xml_post As String = "</GazoSettingList>"

                Dim gazocnt As Integer = 17
                Dim tmp_gazono(gazocnt) As String
                Dim tmp_gazosyuV7(gazocnt) As String
                Dim tmp_gazonoV7(gazocnt) As String
                Dim tmp_gazosyurendo(gazocnt) As String

                '画像に関する情報を配列へ格納
                For Each item In hash

                    Dim fldname As String = item.Key
                    Dim value As String = item.Value

                    Select Case True
                        Case fldname.Replace("画像No", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("画像No", ""))
                            tmp_gazono(gazono) = value
                        Case fldname.Replace("革命側の画像種別", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("革命側の画像種別", ""))
                            tmp_gazosyuV7(gazono) = value
                        Case fldname.Replace("革命側の画像タイトルNo", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("革命側の画像タイトルNo", ""))
                            tmp_gazonoV7(gazono) = value
                        Case fldname.Replace("連動側の画像種別", "") <> fldname
                            Dim gazono As Integer = Int32.Parse(fldname.Replace("連動側の画像種別", ""))
                            tmp_gazosyurendo(gazono) = value
                    End Select

                Next

                '文字列へ成形
                Dim xml_gazoitem As New System.Text.StringBuilder
                For cntii = 1 To gazocnt
                    If tmp_gazosyurendo(cntii) <> "0" And tmp_gazosyuV7(cntii) <> "0" And tmp_gazonoV7(cntii) <> "0" Then
                        xml_gazoitem.Append("<a:PortalGazoSettingModel>")
                        xml_gazoitem.Append("<a:FK8LocationKbn>" & tmp_gazosyuV7(cntii) & "</a:FK8LocationKbn>")
                        xml_gazoitem.Append("<a:FK8LocationNo>" & tmp_gazonoV7(cntii) & "</a:FK8LocationNo>")
                        xml_gazoitem.Append("<a:No>" & tmp_gazono(cntii) & "</a:No>")
                        xml_gazoitem.Append("<a:PortalGazoSyubetu>" & tmp_gazosyurendo(cntii) & "</a:PortalGazoSyubetu>")
                        xml_gazoitem.Append("</a:PortalGazoSettingModel>")
                    End If
                Next

                '結合
                Dim xml_gazo As String = xml_pre & xml_gazoitem.ToString & xml_post
                rtn_str = xml_gazo

                Return rtn_str

            End Function

            ''' <summary>
            ''' 入金項目に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Nkinkomk(ByVal hash_xmlitem As Hashtable, ByVal hash_nkinkomkmst As Hashtable) As String

                Dim rtn_str As String = ""
                Dim xml_pre As String = "<NkinKomokDict xmlns:a=" & """" & "http://schemas.microsoft.com/2003/10/Serialization/Arrays" & """" & ">"
                Dim xml_post As String = "</NkinKomokDict>"
                Dim nkincnt As Integer = 10
                Dim cntindex As Integer = 1
                Dim xml_nkinkomk(nkincnt) As String
                Dim hash_nkinkomkitem As New Hashtable

                '入金項目を移行用に成形する
                Dim hash_cvxmlitem As New Hashtable
                hash_cvxmlitem = Me.Chg_NkinKomk(hash_xmlitem, hash_nkinkomkmst)

                '賃料
                hash_nkinkomkitem.Add("item_key", "1")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_賃料"))
                hash_nkinkomkitem.Add("item_tani", "_万円")
                hash_nkinkomkitem.Add("item_komk", "_賃料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_賃料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '管理費
                hash_nkinkomkitem.Add("item_key", "4")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_管理費"))
                hash_nkinkomkitem.Add("item_tani", "_万円")
                hash_nkinkomkitem.Add("item_komk", "_管理費")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_管理費No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '礼金
                hash_nkinkomkitem.Add("item_key", "5")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_礼金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("礼金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_礼金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_礼金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '敷金
                hash_nkinkomkitem.Add("item_key", "6")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_敷金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("敷金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_敷金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_敷金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '償却金
                hash_nkinkomkitem.Add("item_key", "9")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_償却金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("償却金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_償却金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_償却金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '保証金
                hash_nkinkomkitem.Add("item_key", "7")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_保証金"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("保証金表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_保証金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_保証金No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '敷引
                hash_nkinkomkitem.Add("item_key", "12")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_敷引"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("敷引表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_敷引金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_敷引No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '仲介手数料
                hash_nkinkomkitem.Add("item_key", "15")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_仲介手数料"))
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem("仲介手数料表示単位"))
                hash_nkinkomkitem.Add("item_komk", "_仲介手数料")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_仲介手数料No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                'ほか初期費用
                hash_nkinkomkitem.Add("item_key", "18")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_ほか初期費用"))
                hash_nkinkomkitem.Add("item_tani", "_万円")
                hash_nkinkomkitem.Add("item_komk", "_その他一時金")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_ほか初期費用No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                'その他諸費用
                hash_nkinkomkitem.Add("item_key", "17")
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem("入金項目_その他諸費用"))
                hash_nkinkomkitem.Add("item_tani", "_万円")
                hash_nkinkomkitem.Add("item_komk", "_その他費用")
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem("入金項目_その他諸費用No"))
                xml_nkinkomk(cntindex) = Me.Chg_NkinkomkToXml(hash_nkinkomkitem)
                cntindex = cntindex + 1
                hash_nkinkomkitem.Clear()

                '結合
                Dim xml_total As String = ""
                Dim xml_nkintotal As String = ""
                For cntii = 1 To nkincnt
                    xml_nkintotal = xml_nkintotal & xml_nkinkomk(cntii)
                Next
                xml_total = xml_pre & xml_nkintotal & xml_post

                rtn_str = xml_total

                Return rtn_str

            End Function

            ''' <summary>
            ''' 入金項目の編集
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_xml"></param>
            ''' <param name="hash_nkinkomkmst"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chg_NkinKomk(ByVal hash_xml As Hashtable, ByVal hash_nkinkomkmst As Hashtable) As Hashtable

                Dim rtn_hash As New Hashtable
                Dim tmp_hash As New Hashtable

                Dim tmp_nkincnt As Integer = 14
                Dim tmp_nkinkomkname(tmp_nkincnt) As String
                Dim tmp_nkinkomkkbn(tmp_nkincnt) As String
                Dim tmp_fldname(tmp_nkincnt) As String
                '-----------------------
                '入金項目取得
                '-----------------------
                For Each item In hash_xml

                    Dim midfldname As String = item.Key
                    Dim value As String = item.Value

                    Select Case midfldname
                        Case "入金項目_賃料1" : tmp_nkinkomkname(1) = value : tmp_fldname(1) = midfldname
                        Case "入金項目区分_賃料1" : tmp_nkinkomkkbn(1) = value
                        Case "入金項目_賃料2" : tmp_nkinkomkname(2) = value : tmp_fldname(2) = midfldname
                        Case "入金項目区分_賃料2" : tmp_nkinkomkkbn(2) = value
                        Case "入金項目_賃料3" : tmp_nkinkomkname(3) = value : tmp_fldname(3) = midfldname
                        Case "入金項目区分_賃料3" : tmp_nkinkomkkbn(3) = value
                        Case "入金項目_管理費1" : tmp_nkinkomkname(4) = value : tmp_fldname(4) = midfldname
                        Case "入金項目区分_管理費1" : tmp_nkinkomkkbn(4) = value
                        Case "入金項目_管理費2" : tmp_nkinkomkname(5) = value : tmp_fldname(5) = midfldname
                        Case "入金項目区分_管理費2" : tmp_nkinkomkkbn(5) = value
                        Case "入金項目_管理費3" : tmp_nkinkomkname(6) = value : tmp_fldname(6) = midfldname
                        Case "入金項目区分_管理費3" : tmp_nkinkomkkbn(6) = value
                        Case "入金項目_敷金" : tmp_nkinkomkname(7) = value : tmp_fldname(7) = midfldname
                        Case "入金項目区分_敷金" : tmp_nkinkomkkbn(7) = value
                        Case "入金項目_礼金" : tmp_nkinkomkname(8) = value : tmp_fldname(8) = midfldname
                        Case "入金項目区分_礼金" : tmp_nkinkomkkbn(8) = value
                        Case "入金項目_保証金" : tmp_nkinkomkname(9) = value : tmp_fldname(9) = midfldname
                        Case "入金項目区分_保証金" : tmp_nkinkomkkbn(9) = value
                        Case "入金項目_償却金" : tmp_nkinkomkname(10) = value : tmp_fldname(10) = midfldname
                        Case "入金項目区分_償却金" : tmp_nkinkomkkbn(10) = value
                        Case "入金項目_敷引" : tmp_nkinkomkname(11) = value : tmp_fldname(11) = midfldname
                        Case "入金項目区分_敷引" : tmp_nkinkomkkbn(11) = value
                        Case "入金項目_仲介手数料" : tmp_nkinkomkname(12) = value : tmp_fldname(12) = midfldname
                        Case "入金項目区分_仲介手数料" : tmp_nkinkomkkbn(12) = value
                        Case "入金項目_その他諸費用" : tmp_nkinkomkname(13) = value : tmp_fldname(13) = midfldname
                        Case "入金項目区分_その他諸費用" : tmp_nkinkomkkbn(13) = value
                        Case "入金項目_ほか初期費用" : tmp_nkinkomkname(14) = value : tmp_fldname(14) = midfldname
                        Case "入金項目区分_ほか初期費用" : tmp_nkinkomkkbn(14) = value
                        Case Else
                            tmp_hash.Add(midfldname, value)
                    End Select

                Next

                '-----------------------
                '入金項目をXml移行用に成形
                '-----------------------
                Dim tmp_nkinno(tmp_nkincnt) As String
                Dim tmp_dispnkinname(tmp_nkincnt) As String
                For cntii = 1 To tmp_nkincnt

                    Dim tmp_relnkinstr As String = tmp_nkinkomkname(cntii) & "-" & EtcMethod.Get_Nkinruiname(tmp_nkinkomkkbn(cntii))
                    tmp_nkinno(cntii) = ""
                    tmp_dispnkinname(cntii) = ""

                    If Hash_Rel_Nkinkomk.Contains(tmp_relnkinstr) Then
                        tmp_nkinno(cntii) = Hash_Rel_Nkinkomk(tmp_relnkinstr)
                        tmp_dispnkinname(cntii) = tmp_nkinno(cntii) & ":" & hash_nkinkomkmst(tmp_nkinno(cntii))
                    End If

                Next

                '-----------------------
                '入金項目の再格納
                '-----------------------
                '結合する項目
                Call Me.Set_NkinkomkAddToHash(1, 3, tmp_nkinno, tmp_dispnkinname, "入金項目_賃料", tmp_hash)
                Call Me.Set_NkinkomkAddToHash(4, 6, tmp_nkinno, tmp_dispnkinname, "入金項目_管理費", tmp_hash)

                '結合不要の項目
                For cntii = 7 To 14
                    tmp_hash.Add(tmp_fldname(cntii), tmp_dispnkinname(cntii))
                    tmp_hash.Add(tmp_fldname(cntii) & "No", tmp_nkinno(cntii))
                Next

                rtn_hash = tmp_hash

                Return rtn_hash

            End Function

            ''' <summary>
            ''' 入金項目を結合して1つのデータにまとめる
            ''' まとめたデータをハッシュテーブルへ格納する
            ''' </summary>
            ''' <param name="cntsta"></param>
            ''' <param name="cntend"></param>
            ''' <param name="tmp_nkinno"></param>
            ''' <param name="tmp_dispnkinname"></param>
            ''' <param name="hashkey"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_NkinkomkAddToHash(ByVal cntsta As Integer, ByVal cntend As Integer, ByVal tmp_nkinno As Object, ByVal tmp_dispnkinname As Object, ByVal hashkey As String, ByRef hash As Hashtable)

                Dim nkinno As String = ""
                Dim dispname As String = ""
                For cntii = cntsta To cntend
                    If tmp_nkinno(cntii) <> "" Then
                        nkinno = nkinno & "," & tmp_nkinno(cntii)
                        dispname = dispname & " / " & tmp_dispnkinname(cntii)
                    End If
                Next
                If nkinno <> "" Then
                    nkinno = nkinno.Remove(0, 1)
                End If
                If dispname <> "" Then
                    dispname = dispname.Remove(0, 3)
                End If
                hash.Add(hashkey, dispname)
                hash.Add(hashkey & "No", nkinno)

            End Sub

            ''' <summary>
            ''' 各入金項目毎にXml文字列作成
            ''' </summary>
            ''' <param name="hash_nkin"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chg_NkinkomkToXml(ByVal hash_nkin As Hashtable) As String

                Dim rtn_str As String = ""

                '置換用/ハッシュテーブル値取得用文字列
                Dim item_key As String = "item_key"
                Dim item_disp As String = "item_disp"
                Dim item_tani As String = "item_tani"
                Dim item_komk As String = "item_komk"
                Dim item_nkinno As String = "item_nkinno"

                'Xmlのベースを作成
                Dim xml_nkinkomkbase As New System.Text.StringBuilder
                xml_nkinkomkbase.Append("<a:KeyValueOfintPortalNkinKomokLjh4bohd>")
                xml_nkinkomkbase.Append("<a:Key>" & item_key & "</a:Key>")
                xml_nkinkomkbase.Append("<a:Value>")
                xml_nkinkomkbase.Append("<DisplayString>" & item_disp & "</DisplayString>")
                xml_nkinkomkbase.Append("<EnKagetu>" & item_tani & "</EnKagetu>")
                xml_nkinkomkbase.Append("<NkinKomok>" & item_komk & "</NkinKomok>")
                xml_nkinkomkbase.Append("<NkinNoList>")
                xml_nkinkomkbase.Append(item_nkinno)
                xml_nkinkomkbase.Append("</NkinNoList>")
                xml_nkinkomkbase.Append("</a:Value>")
                xml_nkinkomkbase.Append("</a:KeyValueOfintPortalNkinKomokLjh4bohd>")

                '各要素で置換してい成形
                Dim xml_nkinkomk As String = xml_nkinkomkbase.ToString
                xml_nkinkomk = xml_nkinkomk.Replace(item_key, hash_nkin(item_key))
                xml_nkinkomk = xml_nkinkomk.Replace(item_disp, hash_nkin(item_disp))
                xml_nkinkomk = xml_nkinkomk.Replace(item_tani, hash_nkin(item_tani))
                xml_nkinkomk = xml_nkinkomk.Replace(item_komk, hash_nkin(item_komk))
                Dim tmp_nkinnogrp_tinryo As String = hash_nkin(item_nkinno)
                Dim tmp_nkinno_tinryo() As String = tmp_nkinnogrp_tinryo.Split(",")
                Dim nkinno_tinryo As String = ""
                For cntii = 0 To UBound(tmp_nkinno_tinryo)
                    If tmp_nkinno_tinryo(cntii) <> "" Then
                        nkinno_tinryo = nkinno_tinryo & "<a:int>" & tmp_nkinno_tinryo(cntii) & "</a:int>"
                    End If
                Next
                xml_nkinkomk = xml_nkinkomk.Replace(item_nkinno, nkinno_tinryo)

                rtn_str = xml_nkinkomk

                Return rtn_str

            End Function

            ''' <summary>
            ''' その他項目に関するXml文字列の作成
            ''' </summary>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_xmlitem"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlStr_Other(ByVal hash_cvitem As Hashtable, ByVal hash_xmlitem As Hashtable) As String

                Dim rtn_str As String = ""
                Dim othercnt As Integer = 100
                Dim xml_other(othercnt) As String
                Dim cntindex As Integer = 1

                xml_other(cntindex) = "<Password>" & hash_cvitem("site_password") & "</Password>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<SiteID>" & hash_cvitem("site_id") & "</SiteID>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<SiteSendUmu>" & IIf(hash_cvitem("site_sendumu") = "1", "true", "false") & "</SiteSendUmu>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<UseCommonSetting>" & "false" & "</UseCommonSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<BantiDispSetting>" & hash_xmlitem("番地以降の表示(会社間図面および会社間流通向け)") & "</BantiDispSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<BkHyDispSetting>" & hash_xmlitem("物件名・部屋Noの表示(一般ユーザー向け)") & "</BkHyDispSetting>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<BkHyDispSettingKaisya>" & hash_xmlitem("物件名・部屋Noの表示(会社間図面および会社間流通向け)") & "</BkHyDispSettingKaisya>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<FtpPassword>" & hash_xmlitem("FTPパスワード") & "</FtpPassword>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<FtpUserId>" & hash_xmlitem("FTPユーザーID") & "</FtpUserId>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<IsKaiyakuYoteiAsKusitu>" & hash_xmlitem("空室状況設定") & "</IsKaiyakuYoteiAsKusitu>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<KeisaiKakuninBi>" & hash_xmlitem("元付確認日(先物物件のみ)") & "</KeisaiKakuninBi>" : cntindex = cntindex + 1
                xml_other(cntindex) = "<KusituJokyoSetting>" & "false" & "</KusituJokyoSetting>" : cntindex = cntindex + 1   '要確認
                xml_other(cntindex) = "<MapDispSetting>" & hash_xmlitem("地図表示") & "</MapDispSetting>" : cntindex = cntindex + 1

                '結合
                Dim tmp_str As String = ""
                For cntii = 1 To othercnt
                    If xml_other(cntii) Is Nothing Then
                        Exit For
                    End If
                    tmp_str = tmp_str & xml_other(cntii)
                Next

                rtn_str = tmp_str

                Return rtn_str

            End Function

            ''' <summary>
            ''' 会社支店コードの桁数(9桁)チェック
            ''' </summary>
            ''' <param name="tblname"></param>
            ''' <param name="hash_xmlimte"></param>
            ''' <param name="hash_log"></param>
            ''' <remarks></remarks>
            Public Sub Chk_Kaisyasitencode(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_cvitem As Hashtable)

                Dim errstr As String = ""
                Dim log_key As String = tblname & "-" & "site_data"
                Dim kaisyasitencode As String = hash_cvitem("site_id")

                '桁数チェック(V7画面上で最大9桁の制御が設けられているため9桁に満たないかチェックする)
                If kaisyasitencode.Length < 9 Then

                    '9桁へ成形
                    Dim tmp_kaisyasitencode As String = kaisyasitencode
                    tmp_kaisyasitencode = "000000000" & tmp_kaisyasitencode
                    tmp_kaisyasitencode = tmp_kaisyasitencode.Substring(tmp_kaisyasitencode.Length - 9)

                    '再格納
                    hash_cvitem("site_id") = tmp_kaisyasitencode

                    'ログ出力(※xmlデータのログは直接出力する)
                    Dim syorikomok As String = "送信設定SUUMO情報"
                    Dim syoriitem As String = "会社支店コード"
                    Dim syorikekka As String = LOG_NAIYO_ERR_OUTOFRANGE_STR
                    Dim taisyodata As String = "送信設定順 = " & hash_cvitem("setting_guid") & "、" & "サイトNo = " & hash_cvitem("site_no")
                    Dim hubigein As String = LOG_HUBI_OUTOFRANGE_STR_SHORTAGE
                    Dim taisyo As String = LOG_TAISYO_OUTOOFRANGE_STR
                    Dim beforechgvalue As String = kaisyasitencode
                    Dim afterchgvalue As String = tmp_kaisyasitencode

                    'ログ挿入用へ加工
                    Dim logitem As String = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname)
                    '挿入用クエリへ加工
                    Dim log_sql As String = LogSetting.Get_LogTblInsertQry(logitem, False)
                    'ログ挿入処理
                    Dim tmp_cnt As Integer = 0
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, tmp_cnt)

                End If

            End Sub

        End Class

    End Class

#End Region

#Region "広告補足自社web情報"

    Public Class Hydata_kokoku_jisyaweb_Repository

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
                Dim hash_hyguid As New Hashtable                                'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_kokoku_Model           '移行値格納用モデル初期化
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

                '部屋guid取得
                Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_kokoku"
                Dim fldnamegrp As String = "hy_guid,site_no,item"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

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

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""
                        Dim hash_cvxmlitem As New Hashtable

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
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case Else
                                    If fldname <> "" Then
                                        hash_cvxmlitem.Add(fldname, fldvalue)
                                    End If
                            End Select

                        Next

                        '部屋guid取得用
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno

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
                            hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                'xmlの移行
                                Call Me.Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem)

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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) "
                tmp_sql = tmp_sql & " FROM hydata_kokoku AS HYKOKOKU "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql = tmp_sql & " 	,HY.hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 広告補足のxml項目の移行
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <remarks></remarks>
            Public Sub Set_XmlItem(ByVal sqlcnnv10 As SqlConnection, ByVal hash_cvitem As Hashtable, ByVal hash_cvxmlitem As Hashtable)

                Dim hyguid As String = hash_cvitem("hy_guid")   '更新対象選択用
                Dim siteno As String = hash_cvitem("site_no")   '更新対象選択用
                Dim hash_xmlindex As New Hashtable              'xml要素の順番決定用

                'xml要素の順番設定
                hash_xmlindex = Me.Get_Hash_XmlIndex()

                '「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                Dim xmlupdateqry_xmlpre As String = "<WmpKokokuModel xmlns:i=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & ">"
                Dim xmlupdateqry_xmlpost As String = "</WmpKokokuModel>"
                Dim xmlitem As String = Get_XmlValue(hash_cvxmlitem, hash_xmlindex)
                Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & xmlitem & xmlupdateqry_xmlpost

                '更新用クエリの作成
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " DECLARE @tmp_xml XML "
                tmp_sql = tmp_sql & " SET @tmp_xml = CAST('" & xmlupdateqry_main & "' AS XML) "
                tmp_sql = tmp_sql & " UPDATE hydata_kokoku SET item = @tmp_xml "
                tmp_sql = tmp_sql & " WHERE hy_guid = '" & hyguid & "' "
                tmp_sql = tmp_sql & " AND site_no = " & siteno

                '更新
                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

            ''' <summary>
            ''' xml要素の順番決定
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_Hash_XmlIndex() As Hashtable

                Dim rtn_hash As New Hashtable

                rtn_hash.Add("AdditionalSalesPoint", 1)
                rtn_hash.Add("Biko", 2)
                rtn_hash.Add("CelingHeight", 3)
                rtn_hash.Add("ChinryoNegotiable", 4)
                rtn_hash.Add("RaisedFloor", 5)
                rtn_hash.Add("TantoComment", 6)
                rtn_hash.Add("Tokki", 7)

                Return rtn_hash

            End Function

            ''' <summary>
            ''' xml移行用の文字列の生成
            ''' </summary>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlValue(ByVal hash_cvxmlitem As Hashtable, ByVal hash_xmlindex As Hashtable) As String

                Dim rtn_str As New System.Text.StringBuilder
                Dim tmp_slist As New SortedList(Of Integer, String)

                For Each midvalue In hash_cvxmlitem

                    Dim midfldname As String = midvalue.Key
                    Dim cvvalue As String = midvalue.Value
                    Dim tmp_xmlelement As String = ""

                    Select Case midfldname
                        Case "セールスポイント補足"
                            tmp_xmlelement = "AdditionalSalesPoint"
                        Case "担当者コメント"
                            tmp_xmlelement = "TantoComment"
                        Case "特記事項"
                            tmp_xmlelement = "Tokki"
                        Case "広告備考"
                            tmp_xmlelement = "Biko"
                        Case "天井高"
                            tmp_xmlelement = "CelingHeight"
                        Case "OAフロアの高さ"
                            tmp_xmlelement = "RaisedFloor"
                        Case "賃料応相談"
                            tmp_xmlelement = "ChinryoNegotiable"
                    End Select

                    Dim xmlelement As String = "<" & tmp_xmlelement & ">" & cvvalue & "</" & tmp_xmlelement & ">"
                    Dim keyno As Integer = hash_xmlindex(tmp_xmlelement)
                    tmp_slist.Add(keyno, xmlelement)

                Next

                For Each item In tmp_slist
                    Dim xmlvalue As String = item.Value
                    rtn_str.Append(xmlvalue)
                Next

                Return rtn_str.ToString

            End Function

        End Class

    End Class

#End Region

#Region "広告補足HOMES情報"

    Public Class Hydata_kokoku_homes_Repository

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
                Dim hash_hyguid As New Hashtable                                'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_kokoku_Model           '移行値格納用モデル初期化
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

                '部屋guid取得
                Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_kokoku"
                Dim fldnamegrp As String = "hy_guid,site_no,item"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

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

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""
                        Dim hash_cvxmlitem As New Hashtable

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
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case Else
                                    If fldname <> "" Then
                                        hash_cvxmlitem.Add(fldname, fldvalue)
                                    End If
                            End Select

                        Next

                        '部屋guid取得用
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno

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
                            hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                'xmlの移行
                                Call Me.Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem)

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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) "
                tmp_sql = tmp_sql & " FROM hydata_kokoku AS HYKOKOKU "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql = tmp_sql & " 	,HY.hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 広告補足のxml項目の移行
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <remarks></remarks>
            Public Sub Set_XmlItem(ByVal sqlcnnv10 As SqlConnection, ByVal hash_cvitem As Hashtable, ByVal hash_cvxmlitem As Hashtable)

                Dim hyguid As String = hash_cvitem("hy_guid")   '更新対象選択用
                Dim siteno As String = hash_cvitem("site_no")   '更新対象選択用
                Dim hash_xmlindex As New Hashtable              'xml要素の順番決定用

                'xml要素の順番設定
                hash_xmlindex = Me.Get_Hash_XmlIndex()

                '「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                Dim xmlupdateqry_xmlpre As String = "<HomesKokokuModel xmlns:i=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & ">"
                Dim xmlupdateqry_xmlpost As String = "</HomesKokokuModel>"
                Dim xmlitem As String = Get_XmlValue(hash_cvxmlitem, hash_xmlindex)
                Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & xmlitem & xmlupdateqry_xmlpost

                '更新用クエリの作成
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " DECLARE @tmp_xml XML "
                tmp_sql = tmp_sql & " SET @tmp_xml = CAST('" & xmlupdateqry_main & "' AS XML) "
                tmp_sql = tmp_sql & " UPDATE hydata_kokoku SET item = @tmp_xml "
                tmp_sql = tmp_sql & " WHERE hy_guid = '" & hyguid & "' "
                tmp_sql = tmp_sql & " AND site_no = " & siteno

                '更新
                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

            ''' <summary>
            ''' xml要素の順番決定
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_Hash_XmlIndex() As Hashtable

                Dim rtn_hash As New Hashtable

                rtn_hash.Add("BaikaiKeiyakuDate", 1)
                rtn_hash.Add("Etc", 2)
                rtn_hash.Add("EtcOemA", 3)
                rtn_hash.Add("EtcOemB", 4)
                rtn_hash.Add("GazoNotDownload", 5)
                rtn_hash.Add("GenteiKokai", 6)
                rtn_hash.Add("IsAkiyaBank", 7)
                rtn_hash.Add("KagiBiko", 8)
                rtn_hash.Add("KagiHokan", 9)
                rtn_hash.Add("Memo", 10)
                rtn_hash.Add("MonthlyEnable", 11)
                rtn_hash.Add("MustContractParking", 12)
                rtn_hash.Add("NaisoCustomize", 13)
                rtn_hash.Add("NaisoCustomizeJoken", 14)
                rtn_hash.Add("NaisoCustomizeNaiyo", 15)
                rtn_hash.Add("OsusumePointScore", 16)
                rtn_hash.Add("ReformBiko", 17)
                rtn_hash.Add("ReformDate", 18)
                rtn_hash.Add("ReformPartBathroom", 19)
                rtn_hash.Add("ReformPartCloth", 20)
                rtn_hash.Add("ReformPartEtc", 21)
                rtn_hash.Add("ReformPartFloor", 22)
                rtn_hash.Add("ReformPartInteriorDesign", 23)
                rtn_hash.Add("ReformPartKitchen", 24)
                rtn_hash.Add("ReformPartOutside", 25)
                rtn_hash.Add("ReformPartToilet", 26)
                rtn_hash.Add("ReformPartWashRoom", 27)
                rtn_hash.Add("RenovationDate", 28)
                rtn_hash.Add("RenovationMemo", 29)
                rtn_hash.Add("SalesStaff", 30)
                rtn_hash.Add("StaffComment", 31)
                rtn_hash.Add("StaffCommentKbn", 32)
                rtn_hash.Add("StaffCommentSettingMethod", 33)
                rtn_hash.Add("StaffCommentTemplateNo", 34)
                rtn_hash.Add("Tasyatorikomi", 35)
                rtn_hash.Add("TokucyouA", 36)
                rtn_hash.Add("TokucyouB", 37)
                rtn_hash.Add("TokuyutinBiko", 38)
                rtn_hash.Add("TokuyutinEnable", 39)
                rtn_hash.Add("TokuyutinJyogen", 40)
                rtn_hash.Add("TokuyutinJyousyoritu", 41)
                rtn_hash.Add("TokuyutinKagen", 42)
                rtn_hash.Add("TokuyutinRyokinHendo", 43)
                rtn_hash.Add("TokuyutinYatinHojyoNensu", 44)
                rtn_hash.Add("Url", 45)
                rtn_hash.Add("UrlKind", 46)
                rtn_hash.Add("YotakusakiKaisya", 47)

                Return rtn_hash

            End Function

            ''' <summary>
            ''' xml移行用の文字列の生成
            ''' </summary>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlValue(ByVal hash_cvxmlitem As Hashtable, ByVal hash_xmlindex As Hashtable) As String

                Dim rtn_str As New System.Text.StringBuilder
                Dim tmp_slist As New SortedList(Of Integer, String)

                For Each midvalue In hash_cvxmlitem

                    Dim midfldname As String = midvalue.Key
                    Dim cvvalue As String = midvalue.Value
                    Dim tmp_xmlelement As String = ""

                    Select Case midfldname
                        Case "物件の特徴_自社HPに表示"
                            tmp_xmlelement = "TokucyouA"
                        Case "物件の特徴_他社HPに表示"
                            tmp_xmlelement = "TokucyouB"
                        Case "広告備考"
                            tmp_xmlelement = "Etc"
                        Case "広告備考_自社HPに表示"
                            tmp_xmlelement = "EtcOemA"
                        Case "広告備考_他社HPに表示"
                            tmp_xmlelement = "EtcOemB"
                        Case "リンク先区分"
                            tmp_xmlelement = "UrlKind"
                        Case "リンク先URL"
                            tmp_xmlelement = "Url"
                        Case "社内用メモ"
                            tmp_xmlelement = "Memo"
                        Case "リフォーム実施年月"
                            tmp_xmlelement = "ReformDate"
                        Case "リフォームその他箇所"
                            tmp_xmlelement = "ReformPartEtc"
                        Case "リフォーム備考"
                            tmp_xmlelement = "ReformBiko"
                        Case "リノベーション施工完了年月"
                            tmp_xmlelement = "RenovationDate"
                        Case "リノベーション内容"
                            tmp_xmlelement = "RenovationMemo"
                        Case "カスタマイズ可否"
                            tmp_xmlelement = "NaisoCustomize"
                        Case "カスタマイズ内容"
                            tmp_xmlelement = "NaisoCustomizeNaiyo"
                        Case "カスタマイズ条件"
                            tmp_xmlelement = "NaisoCustomizeJoken"
                        Case "特定優遇賃貸住宅チェック"
                            tmp_xmlelement = "TokuyutinEnable"
                        Case "特優賃下限"
                            tmp_xmlelement = "TokuyutinKagen"
                        Case "特優賃上限"
                            tmp_xmlelement = "TokuyutinJyogen"
                        Case "特優賃料金変動区分"
                            tmp_xmlelement = "TokuyutinRyokinHendo"
                        Case "特優賃料金上昇率"
                            tmp_xmlelement = "TokuyutinJyousyoritu"
                        Case "特優賃家賃補助年数"
                            tmp_xmlelement = "TokuyutinYatinHojyoNensu"
                        Case "特優賃備考"
                            tmp_xmlelement = "TokuyutinBiko"
                        Case "鍵保管場所"
                            tmp_xmlelement = "KagiHokan"
                        Case "預託先会社名"
                            tmp_xmlelement = "YotakusakiKaisya"
                        Case "鍵備考"
                            tmp_xmlelement = "KagiBiko"
                        Case "限定公開"
                            tmp_xmlelement = "GenteiKokai"
                        Case "画像ダウンロード不可"
                            tmp_xmlelement = "GazoNotDownload"
                        Case "営業スタッフ名"
                            tmp_xmlelement = "SalesStaff"
                        Case "営業スタッフコメント設定方法"
                            tmp_xmlelement = "StaffCommentSettingMethod"
                        Case "コメントテンプレート"
                            tmp_xmlelement = "StaffCommentTemplateNo"
                        Case "コメント種別"
                            tmp_xmlelement = "StaffCommentKbn"
                        Case "コメント"
                            tmp_xmlelement = "StaffComment"
                        Case "マンスリー可"
                            tmp_xmlelement = "MonthlyEnable"
                        Case "駐車場契約必須"
                            tmp_xmlelement = "MustContractParking"
                        Case "他社取込"
                            tmp_xmlelement = "Tasyatorikomi"
                        Case "媒介契約年月日"
                            tmp_xmlelement = "BaikaiKeiyakuDate"
                        Case "特別広告ポイント数"
                            tmp_xmlelement = "OsusumePointScore"
                        Case "空き家バンク登録物件"
                            tmp_xmlelement = "IsAkiyaBank"
                    End Select

                    Dim xmlelement As String = "<" & tmp_xmlelement & ">" & cvvalue & "</" & tmp_xmlelement & ">"
                    Dim keyno As Integer = hash_xmlindex(tmp_xmlelement)
                    tmp_slist.Add(keyno, xmlelement)

                Next

                For Each item In tmp_slist
                    Dim xmlvalue As String = item.Value
                    rtn_str.Append(xmlvalue)
                Next

                Return rtn_str.ToString

            End Function

        End Class

    End Class

#End Region

#Region "広告補足athome情報"

    Public Class Hydata_kokoku_athome_Repository

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
                Dim hash_hyguid As New Hashtable                                'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_kokoku_Model           '移行値格納用モデル初期化
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

                '部屋guid取得
                Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_kokoku"
                Dim fldnamegrp As String = "hy_guid,site_no,item"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

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

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""
                        Dim hash_cvxmlitem As New Hashtable

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
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case Else
                                    If fldname <> "" Then
                                        hash_cvxmlitem.Add(fldname, fldvalue)
                                    End If
                            End Select

                        Next

                        '部屋guid取得用
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno

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
                            hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                'xmlの移行
                                Call Me.Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem)

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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) "
                tmp_sql = tmp_sql & " FROM hydata_kokoku AS HYKOKOKU "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql = tmp_sql & " 	,HY.hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 広告補足のxml項目の移行
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <remarks></remarks>
            Public Sub Set_XmlItem(ByVal sqlcnnv10 As SqlConnection, ByVal hash_cvitem As Hashtable, ByVal hash_cvxmlitem As Hashtable)

                Dim hyguid As String = hash_cvitem("hy_guid")   '更新対象選択用
                Dim siteno As String = hash_cvitem("site_no")   '更新対象選択用
                Dim hash_xmlindex As New Hashtable              'xml要素の順番決定用

                'xml要素の順番設定
                hash_xmlindex = Me.Get_Hash_XmlIndex()

                '「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                Dim xmlupdateqry_xmlpre As String = "<AthomeKokokuModel xmlns:i=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & ">"
                Dim xmlupdateqry_xmlpost As String = "</AthomeKokokuModel>"
                Dim xmlitem As String = Get_XmlValue(hash_cvxmlitem, hash_xmlindex)
                Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & xmlitem & xmlupdateqry_xmlpost

                '更新用クエリの作成
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " DECLARE @tmp_xml XML "
                tmp_sql = tmp_sql & " SET @tmp_xml = CAST('" & xmlupdateqry_main & "' AS XML) "
                tmp_sql = tmp_sql & " UPDATE hydata_kokoku SET item = @tmp_xml "
                tmp_sql = tmp_sql & " WHERE hy_guid = '" & hyguid & "' "
                tmp_sql = tmp_sql & " AND site_no = " & siteno

                '更新
                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

            ''' <summary>
            ''' xml要素の順番決定
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_Hash_XmlIndex() As Hashtable

                Dim rtn_hash As New Hashtable

                rtn_hash.Add("BkNameEnabled", 1)
                rtn_hash.Add("CanDiscountInitialCost", 2)
                rtn_hash.Add("CanDiscountTinryo", 3)
                rtn_hash.Add("Etc", 4)
                rtn_hash.Add("EtcCardPayment", 5)
                rtn_hash.Add("EtcCardPaymentUmu", 6)
                rtn_hash.Add("EtcPet", 7)
                rtn_hash.Add("EtcReform", 8)
                rtn_hash.Add("EtcRenovation", 9)
                rtn_hash.Add("EtcSyokyakuJoken", 10)
                rtn_hash.Add("ExteriorReformKbn", 11)
                rtn_hash.Add("HyNoEnabled", 12)
                rtn_hash.Add("InteriorReformKbn", 13)
                rtn_hash.Add("KagiGentiTaio", 14)
                rtn_hash.Add("KosinRyoKbn", 15)
                rtn_hash.Add("NearBikeOkibaRyokin", 16)
                rtn_hash.Add("NearBikeOkibaTax", 17)
                rtn_hash.Add("NearTyurinjoRyokin", 18)
                rtn_hash.Add("NearTyurinjoTax", 19)
                rtn_hash.Add("PanoramaContentsId", 20)
                rtn_hash.Add("ProComment", 21)
                rtn_hash.Add("ReformUmu", 22)
                rtn_hash.Add("ReformYm", 23)
                rtn_hash.Add("RenovationUmu", 24)
                rtn_hash.Add("RenovationYm", 25)
                rtn_hash.Add("StaffID", 26)
                rtn_hash.Add("YahooEnabled", 27)

                Return rtn_hash

            End Function

            ''' <summary>
            ''' xml移行用の文字列の生成
            ''' </summary>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlValue(ByVal hash_cvxmlitem As Hashtable, ByVal hash_xmlindex As Hashtable) As String

                Dim rtn_str As New System.Text.StringBuilder
                Dim tmp_slist As New SortedList(Of Integer, String)
                Dim xml_null As String = "i:nil=" & """" & "true" & """"

                For Each midvalue In hash_cvxmlitem

                    Dim midfldname As String = midvalue.Key
                    Dim cvvalue As String = midvalue.Value
                    Dim tmp_xmlelement As String = ""

                    Select Case midfldname
                        Case "プロのコメント"
                            tmp_xmlelement = "ProComment"
                        Case "プロのコメントスタッフID"
                            tmp_xmlelement = "StaffID"
                        Case "広告備考"
                            tmp_xmlelement = "Etc"
                        Case "その他ペット可内容"
                            tmp_xmlelement = "EtcPet"
                        Case "更新料区分"
                            tmp_xmlelement = "KosinRyoKbn"
                        Case "保証金その他償却条件内容"
                            tmp_xmlelement = "EtcSyokyakuJoken"
                        Case "その他クレジット決済可"
                            tmp_xmlelement = "EtcCardPaymentUmu"
                        Case "その他クレジット決済可能条件等"
                            tmp_xmlelement = "EtcCardPayment"
                        Case "リフォーム有無"
                            tmp_xmlelement = "ReformUmu"
                        Case "リフォーム年月"
                            tmp_xmlelement = "ReformYm"
                        Case "対象(内装関連)"
                            tmp_xmlelement = "InteriorReformKbn"
                        Case "対象(外装関連)"
                            tmp_xmlelement = "ExteriorReformKbn"
                        Case "リフォーム内容"
                            tmp_xmlelement = "EtcReform"
                        Case "リノベーション有無"
                            tmp_xmlelement = "RenovationUmu"
                        Case "リノベーション実施年月"
                            tmp_xmlelement = "RenovationYm"
                        Case "リノベーション内容"
                            tmp_xmlelement = "EtcRenovation"
                        Case "建物名表示"
                            tmp_xmlelement = "BkNameEnabled"
                        Case "部屋番号表示"
                            tmp_xmlelement = "HyNoEnabled"
                        Case "鍵現地対応"
                            tmp_xmlelement = "KagiGentiTaio"
                        Case "賃料値下げ可"
                            tmp_xmlelement = "CanDiscountTinryo"
                        Case "初期費用値下げ可"
                            tmp_xmlelement = "CanDiscountInitialCost"
                        Case "パノラマコンテンツID"
                            tmp_xmlelement = "PanoramaContentsId"
                        Case "近隣駐輪場料金"
                            tmp_xmlelement = "NearTyurinjoRyokin"
                        Case "近隣駐輪場料金_税区分"
                            tmp_xmlelement = "NearTyurinjoTax"
                        Case "近隣バイク置き場料金"
                            tmp_xmlelement = "NearBikeOkibaRyokin"
                        Case "近隣バイク置き場料金_税区分"
                            tmp_xmlelement = "NearBikeOkibaTax"
                        Case "Yahoo公開不可"
                            tmp_xmlelement = "YahooEnabled"
                    End Select

                    Dim xmlelement As String = "<" & tmp_xmlelement & ">" & cvvalue & "</" & tmp_xmlelement & ">"
                    If (tmp_xmlelement = "NearTyurinjoRyokin" Or tmp_xmlelement = "NearBikeOkibaRyokin") And cvvalue = "" Then
                        xmlelement = xmlelement.Replace("<" & tmp_xmlelement & ">", "<" & tmp_xmlelement & " " & xml_null & ">")
                    End If
                    Dim keyno As Integer = hash_xmlindex(tmp_xmlelement)
                    tmp_slist.Add(keyno, xmlelement)

                Next

                For Each item In tmp_slist
                    Dim xmlvalue As String = item.Value
                    rtn_str.Append(xmlvalue)
                Next

                Return rtn_str.ToString

            End Function

        End Class

    End Class

#End Region

#Region "広告補足SUUMO情報"

    Public Class Hydata_kokoku_suumo_Repository

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
                Dim hash_hyguid As New Hashtable                                'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_kokoku_Model           '移行値格納用モデル初期化
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

                '部屋guid取得
                Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_kokoku"
                Dim fldnamegrp As String = "hy_guid,site_no,item"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

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

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""
                        Dim hash_cvxmlitem As New Hashtable

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
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case Else
                                    If fldname <> "" Then
                                        hash_cvxmlitem.Add(fldname, fldvalue)
                                    End If
                            End Select

                        Next

                        '部屋guid取得用
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno

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
                            hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                'xmlの移行
                                Call Me.Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem)

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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) "
                tmp_sql = tmp_sql & " FROM hydata_kokoku AS HYKOKOKU "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql = tmp_sql & " 	,HY.hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 広告補足のxml項目の移行
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <remarks></remarks>
            Public Sub Set_XmlItem(ByVal sqlcnnv10 As SqlConnection, ByVal hash_cvitem As Hashtable, ByVal hash_cvxmlitem As Hashtable)

                Dim hyguid As String = hash_cvitem("hy_guid")   '更新対象選択用
                Dim siteno As String = hash_cvitem("site_no")   '更新対象選択用
                Dim hash_xmlindex As New Hashtable              'xml要素の順番決定用

                'xml要素の順番設定
                hash_xmlindex = Me.Get_Hash_XmlIndex()

                '「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                Dim xmlupdateqry_xmlpre As String = "<SuumoKokokuModel xmlns:i=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & ">"
                Dim xmlupdateqry_xmlpost As String = "</SuumoKokokuModel>"
                Dim xmlitem As String = Get_XmlValue(hash_cvxmlitem, hash_xmlindex)
                Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & xmlitem & xmlupdateqry_xmlpost

                '更新用クエリの作成
                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " DECLARE @tmp_xml XML "
                tmp_sql = tmp_sql & " SET @tmp_xml = CAST('" & xmlupdateqry_main & "' AS XML) "
                tmp_sql = tmp_sql & " UPDATE hydata_kokoku SET item = @tmp_xml "
                tmp_sql = tmp_sql & " WHERE hy_guid = '" & hyguid & "' "
                tmp_sql = tmp_sql & " AND site_no = " & siteno

                '更新
                Dim tmp_cnt As Integer = 0
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)

            End Sub

            ''' <summary>
            ''' xml要素の順番決定
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_Hash_XmlIndex() As Hashtable

                Dim rtn_hash As New Hashtable
                Dim xml_namesp As String = "xmlns:a=" & """" & "http://schemas.datacontract.org/2004/07/Njc.N3Lib.Utys" & """"

                rtn_hash.Add("AddressEnabled", 1)
                rtn_hash.Add("AllowCopySearchForKaisyakan", 2)
                rtn_hash.Add("BkNameEnabled", 3)
                rtn_hash.Add("EtcCarParking", 4)
                rtn_hash.Add("EtcReform", 5)
                rtn_hash.Add("EtcTokuyutin", 6)
                rtn_hash.Add("FreeCommentForKaisyaKan", 7)
                rtn_hash.Add("FreeCommentForNet", 8)
                rtn_hash.Add("HyNoEnabled", 9)
                rtn_hash.Add("JosyoRitu", 10)
                rtn_hash.Add("JutakuSienSite", 11)
                rtn_hash.Add("KengakuYoyakuKino", 12)
                rtn_hash.Add("KisyaKanriCode1", 13)
                rtn_hash.Add("KisyaKanriCode2", 14)
                rtn_hash.Add("MainCatch1", 15)
                rtn_hash.Add("MainCatch2", 16)
                rtn_hash.Add("NetCatch", 17)
                rtn_hash.Add("NyukyoSyaFutanFrom", 18)
                rtn_hash.Add("NyukyosyaFutanTo", 19)
                rtn_hash.Add("OsusumePickUp1", 20)
                rtn_hash.Add("OsusumePickUp2", 21)
                rtn_hash.Add("OsusumePickUp3", 22)
                rtn_hash.Add("OsusumePickUpRui1", 23)
                rtn_hash.Add("OsusumePickUpRui2", 24)
                rtn_hash.Add("OsusumePickUpRui3", 25)
                rtn_hash.Add("ReformJiki", 26)
                rtn_hash.Add("ReformKasyo", 27)
                rtn_hash.Add("RyokinHendoKbn", 28)
                rtn_hash.Add("SelectedKaguKadenList" & " " & xml_namesp, 29)
                rtn_hash.Add("SelectedNyukyoJokenList" & " " & xml_namesp, 30)
                rtn_hash.Add("SelectedReformList" & " " & xml_namesp, 31)
                rtn_hash.Add("SelectedReformNaiyoList" & " " & xml_namesp, 32)
                rtn_hash.Add("SelectedSecurityList" & " " & xml_namesp, 33)
                rtn_hash.Add("SelectedSitunaiMadoriList" & " " & xml_namesp, 34)
                rtn_hash.Add("SetubiKankyo1", 35)
                rtn_hash.Add("SetubiKankyo1F", 36)
                rtn_hash.Add("SetubiKankyo2", 37)
                rtn_hash.Add("SetubiKankyoKyori1", 38)
                rtn_hash.Add("SetubiKankyoKyori2", 39)
                rtn_hash.Add("SetubiKankyoRinsetu1", 40)
                rtn_hash.Add("SetubiKankyoRinsetu2", 41)
                rtn_hash.Add("SetubiKankyoRinsetuTani1", 42)
                rtn_hash.Add("SetubiKankyoRinsetuTani2", 43)
                rtn_hash.Add("SikikinTumimasiGaku", 44)
                rtn_hash.Add("SikikinTumimasiGakuTani", 45)
                rtn_hash.Add("SikikinTumimasiJoken", 46)
                rtn_hash.Add("SubCatch1", 47)
                rtn_hash.Add("SubCatch10", 48)
                rtn_hash.Add("SubCatch2", 49)
                rtn_hash.Add("SubCatch3", 50)
                rtn_hash.Add("SubCatch4", 51)
                rtn_hash.Add("SubCatch5", 52)
                rtn_hash.Add("SubCatch6", 53)
                rtn_hash.Add("SubCatch7", 54)
                rtn_hash.Add("SubCatch8", 55)
                rtn_hash.Add("SubCatch9", 56)
                rtn_hash.Add("SuumoGazoYusen", 57)
                rtn_hash.Add("TokuyutinUmu", 58)
                rtn_hash.Add("YatinHojoNensu", 59)
                rtn_hash.Add("ZumenPattern", 60)

                Return rtn_hash

            End Function

            ''' <summary>
            ''' xml移行用の文字列の生成
            ''' </summary>
            ''' <param name="hash_cvxmlitem"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_XmlValue(ByVal hash_cvxmlitem As Hashtable, ByVal hash_xmlindex As Hashtable) As String

                Dim rtn_str As New System.Text.StringBuilder
                Dim tmp_slist As New SortedList(Of Integer, String)

                For Each midvalue In hash_cvxmlitem

                    Dim midfldname As String = midvalue.Key
                    Dim cvvalue As String = midvalue.Value
                    Dim tmp_xmlelement As String = ""
                    Dim xml_namesp As String = "xmlns:a=" & """" & "http://schemas.datacontract.org/2004/07/Njc.N3Lib.Utys" & """"
                    Dim xml_null As String = "i:nil=" & """" & "true" & """"

                    Select Case midfldname
                        Case "物件の特徴_ネット用キャッチ"
                            tmp_xmlelement = "NetCatch"
                        Case "物件の特徴_フリーコメント"
                            tmp_xmlelement = "FreeCommentForNet"
                        Case "敷金積増条件"
                            tmp_xmlelement = "SikikinTumimasiJoken"
                        Case "敷金積増後総額"
                            tmp_xmlelement = "SikikinTumimasiGaku"
                        Case "敷金積増後総額_単位"
                            tmp_xmlelement = "SikikinTumimasiGakuTani"
                        Case "駐車場備考"
                            tmp_xmlelement = "EtcCarParking"
                        Case "特優賃有無"
                            tmp_xmlelement = "TokuyutinUmu"
                        Case "特優賃入居者負担額_下限"
                            tmp_xmlelement = "NyukyoSyaFutanFrom"
                        Case "特優賃入居者負担額_上限"
                            tmp_xmlelement = "NyukyosyaFutanTo"
                        Case "特優賃料金変動区分"
                            tmp_xmlelement = "RyokinHendoKbn"
                        Case "特優賃料金変動区分上昇率"
                            tmp_xmlelement = "JosyoRitu"
                        Case "特優賃家賃補助年数"
                            tmp_xmlelement = "YatinHojoNensu"
                        Case "特優賃補足"
                            tmp_xmlelement = "EtcTokuyutin"
                        Case "リフォーム時期"
                            tmp_xmlelement = "ReformJiki"
                        Case "リフォーム箇所"
                            tmp_xmlelement = "ReformKasyo"
                        Case "リフォーム補足"
                            tmp_xmlelement = "EtcReform"
                        Case "設備環境1"
                            tmp_xmlelement = "SetubiKankyo1"
                        Case "設備環境1_距離"
                            tmp_xmlelement = "SetubiKankyoKyori1"
                        Case "設備環境2"
                            tmp_xmlelement = "SetubiKankyo2"
                        Case "設備環境2_距離"
                            tmp_xmlelement = "SetubiKankyoKyori2"
                        Case "周辺環境隣接1"
                            tmp_xmlelement = "SetubiKankyoRinsetu1"
                        Case "周辺環境隣接1_単位"
                            tmp_xmlelement = "SetubiKankyoRinsetuTani1"
                        Case "周辺環境隣接2"
                            tmp_xmlelement = "SetubiKankyoRinsetu2"
                        Case "周辺環境隣接2_単位"
                            tmp_xmlelement = "SetubiKankyoRinsetuTani2"
                        Case "設備環境1F"
                            tmp_xmlelement = "SetubiKankyo1F"
                        Case "セキュリティ"
                            tmp_xmlelement = "SelectedSecurityList" & " " & xml_namesp
                        Case "室内間取"
                            tmp_xmlelement = "SelectedSitunaiMadoriList" & " " & xml_namesp
                        Case "家具・家電"
                            tmp_xmlelement = "SelectedKaguKadenList" & " " & xml_namesp
                        Case "リフォーム"
                            tmp_xmlelement = "SelectedReformList" & " " & xml_namesp
                        Case "リフォーム(内容)"
                            tmp_xmlelement = "SelectedReformNaiyoList" & " " & xml_namesp
                        Case "入居条件"
                            tmp_xmlelement = "SelectedNyukyoJokenList" & " " & xml_namesp
                        Case "SUUMO以外の災害時住宅支援サイトへの掲載"
                            tmp_xmlelement = "JutakuSienSite"
                        Case "おすすめピックアップピクト指定1"
                            tmp_xmlelement = "OsusumePickUpRui1"
                        Case "特徴1"
                            tmp_xmlelement = "OsusumePickUp1"
                        Case "おすすめピックアップピクト指定2"
                            tmp_xmlelement = "OsusumePickUpRui2"
                        Case "特徴2"
                            tmp_xmlelement = "OsusumePickUp2"
                        Case "おすすめピックアップピクト指定3"
                            tmp_xmlelement = "OsusumePickUpRui3"
                        Case "特徴3"
                            tmp_xmlelement = "OsusumePickUp3"
                        Case "物件名公開"
                            tmp_xmlelement = "BkNameEnabled"
                        Case "部屋番号公開"
                            tmp_xmlelement = "HyNoEnabled"
                        Case "詳細住所公開"
                            tmp_xmlelement = "AddressEnabled"
                        Case "会社間物件検索コピー許可"
                            tmp_xmlelement = "AllowCopySearchForKaisyakan"
                        Case "図面パターン"
                            tmp_xmlelement = "ZumenPattern"
                        Case "メインキャッチ1"
                            tmp_xmlelement = "MainCatch1"
                        Case "メインキャッチ2"
                            tmp_xmlelement = "MainCatch2"
                        Case "サブキャッチ1"
                            tmp_xmlelement = "SubCatch1"
                        Case "サブキャッチ2"
                            tmp_xmlelement = "SubCatch2"
                        Case "サブキャッチ3"
                            tmp_xmlelement = "SubCatch3"
                        Case "サブキャッチ4"
                            tmp_xmlelement = "SubCatch4"
                        Case "サブキャッチ5"
                            tmp_xmlelement = "SubCatch5"
                        Case "サブキャッチ6"
                            tmp_xmlelement = "SubCatch6"
                        Case "サブキャッチ7"
                            tmp_xmlelement = "SubCatch7"
                        Case "サブキャッチ8"
                            tmp_xmlelement = "SubCatch8"
                        Case "サブキャッチ9"
                            tmp_xmlelement = "SubCatch9"
                        Case "サブキャッチ10"
                            tmp_xmlelement = "SubCatch10"
                        Case "会社間用補足フリーコメント"
                            tmp_xmlelement = "FreeCommentForKaisyaKan"
                        Case "SUUMO内優先画像"
                            tmp_xmlelement = "SuumoGazoYusen"
                        Case "貴社管理コード1"
                            tmp_xmlelement = "KisyaKanriCode1"
                        Case "貴社管理コード2"
                            tmp_xmlelement = "KisyaKanriCode2"
                        Case "見学予約機能を利用する"
                            tmp_xmlelement = "KengakuYoyakuKino"
                    End Select

                    Dim xmlelement As String = "<" & tmp_xmlelement & ">" & cvvalue & "</" & tmp_xmlelement.Replace(xml_namesp, "") & ">"
                    If (tmp_xmlelement = "JosyoRitu" Or _
                        tmp_xmlelement = "NyukyoSyaFutanFrom" Or _
                        tmp_xmlelement = "NyukyosyaFutanTo" Or _
                        tmp_xmlelement = "SetubiKankyoKyori1" Or _
                        tmp_xmlelement = "SetubiKankyoKyori2" Or _
                        tmp_xmlelement = "SikikinTumimasiGaku" Or _
                        tmp_xmlelement = "YatinHojoNensu") And _
                        cvvalue = "" Then
                        xmlelement = xmlelement.Replace("<" & tmp_xmlelement & ">", "<" & tmp_xmlelement & " " & xml_null & ">")
                    End If
                    Dim keyno As Integer = hash_xmlindex(tmp_xmlelement)
                    tmp_slist.Add(keyno, xmlelement)

                Next

                For Each item In tmp_slist
                    Dim xmlvalue As String = item.Value
                    rtn_str.Append(xmlvalue)
                Next

                Return rtn_str.ToString

            End Function

        End Class

    End Class

#End Region

#Region "ポータル連動部屋分類情報"

    Public Class M_hy_ruisite_Repository

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

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.M_hy_ruisite_Model            '移行値格納用モデル初期化
                Dim keycol_main As Integer = 2                                  'メインキー列
                Dim keycol_sub1 As Integer = 3                                  'サブ1キー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Hyrui()

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_hy_ruisite"
                Dim fldnamegrp As String = "hy_ruino,site_no,hy_ruisiteitemdata"

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

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "部屋分類名"
                                    .Vari_Hy_ruino = fldvalue
                                Case "サイトNo"
                                    .Vari_Site_no = fldvalue
                                Case "サイト用部屋分類No"
                                    .Vari_Hy_ruisiteitemdata = fldvalue
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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_jisyano) "
                tmp_sql = tmp_sql & " FROM hydata_sosin AS HYSOSIN "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYSOSIN.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 部屋分類取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_Hyrui(sqlcnnv10 As SqlConnection, ByRef hash_hyrui As Hashtable)

                Dim tmp_sql As String = "SELECT hy_ruiname,hy_ruino FROM m_hy_rui ORDER BY hy_ruino"
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_hyrui)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋毎送信情報"

    Public Class Hydata_sosin_Repository

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
                Dim hash_hyguid As New Hashtable                                'キー/guid格納用ハッシュテーブル
                Dim hash_komkguid As New Hashtable

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_sosin_Model            '移行値格納用モデル初期化
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
                Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)          '部屋guid
                Call Me.Set_KomokGuid(sqlcnnv10, hash_komkguid)     '項目guid

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_sosin"
                Dim fldnamegrp As String = "hy_guid,site_no,site_jisyano,select_no,komok_guid," & _
                                           "history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

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

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""
                        Dim tmp_siteno As String = ""
                        Dim tmp_sitejisyano As String = ""
                        Dim tmp_komok As String = ""

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
                                Case "ポータルサイトNo"
                                    .Vari_Site_no = fldvalue
                                    tmp_siteno = fldvalue
                                Case "自社サービスNo"
                                    .Vari_Site_jisyano = fldvalue
                                    tmp_sitejisyano = fldvalue
                                Case "選択No"
                                    .Vari_Select_no = fldvalue
                                Case "項目値"
                                    tmp_komok = fldvalue
                            End Select

                        Next

                        'guid取得用
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno                                   '部屋guid
                        .Vari_Komok_guid = tmp_siteno & "-" & tmp_sitejisyano & "-" & tmp_komok     '項目guid

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
                            hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))

                            '項目をguidへ変換
                            hash_cvitem("komok_guid") = hash_komkguid.Item(StrConv(hash_cvitem.Item("komok_guid"), VbStrConv.Narrow))

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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_jisyano) "
                tmp_sql = tmp_sql & " FROM hydata_sosin AS HYSOSIN "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYSOSIN.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql = tmp_sql & " 	,HY.hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

            ''' <summary>
            ''' 項目guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_KomokGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,site_no) + '-' + CONVERT(varchar,site_jisyano)  + '-' + CONVERT(varchar,komok_value) AS キー "
                tmp_sql = tmp_sql & " 	,komok_guid "
                tmp_sql = tmp_sql & " FROM m_site_komok "
                tmp_sql = tmp_sql & " WHERE site_no IN (10,20,30,40) "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

        End Class

    End Class

#End Region

#Region "BtoBグループ設定情報"

    Public Class Hydata_btobgroup_Repository

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
                Dim hash_hyguid As New Hashtable                                'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_btobgroup_Model        '移行値格納用モデル初期化
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

                '部屋guid取得
                Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_btobgroup"
                Dim fldnamegrp As String = "hy_guid,group_no,display_kbn,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

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

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""

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
                                Case "BtoBプラグイングループNo"
                                    .Vari_Group_no = fldvalue
                                Case "グループ公開有無"
                                    .Vari_Display_kbn = fldvalue
                            End Select

                        Next

                        '部屋guid取得用
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno

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
                            hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))

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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,group_no) "
                tmp_sql = tmp_sql & " FROM hydata_btobgroup AS HYBB "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYBB.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql = tmp_sql & " 	,HY.hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

        End Class

    End Class

#End Region

#Region "地図表示詳細設定情報"

    Public Class Hydata_mapdisp_Repository

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
                Dim hash_hyguid As New Hashtable                                'キー/guid格納用ハッシュテーブル

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_mapdisp_Model          '移行値格納用モデル初期化
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

                '部屋guid取得
                Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_mapdisp"
                Dim fldnamegrp As String = "hy_guid,site_no,site_jisyano,map_dispflg,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '親マスタ取得
                Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

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

                        '作業用変数
                        Dim tmp_bkno As String = ""
                        Dim tmp_hyno As String = ""

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
                                Case "外部サイトNo（FK）"
                                    .Vari_Site_no = fldvalue
                                Case "自社サイトNo"
                                    .Vari_Site_jisyano = fldvalue
                                Case "地図上に表示フラグ"
                                    .Vari_Map_dispflg = fldvalue
                            End Select

                        Next

                        '部屋guid取得用
                        .Vari_Hy_guid = tmp_bkno & "-" & tmp_hyno

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
                            hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))

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

            ''' <summary>
            ''' 既存データ取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,BK.bk_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,HY.hy_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_no) + '-' +  "
                tmp_sql = tmp_sql & " 	CONVERT(varchar,site_jisyano) "
                tmp_sql = tmp_sql & " FROM hydata_mapdisp AS HYMAP "
                tmp_sql = tmp_sql & " LEFT JOIN hydata AS HY ON HYMAP.hy_guid = HY.hy_guid "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "

                Dim flg As Boolean = True
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                tmp_sql = tmp_sql & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                tmp_sql = tmp_sql & " 	,HY.hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

        End Class

    End Class

#End Region

End Namespace


