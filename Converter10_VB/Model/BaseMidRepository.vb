Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "自社口座情報(汎用用)"

    Public Class Jisyadata_koza_BaseMid_Repository

        Public Class SubConv
            Implements IConv
            '20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -add sta
            ''' <summary>
            ''' 【中間ファイル→変数】 
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="filename"></param>
            ''' <param name="sheetname"></param>
            ''' <param name="midrowcnt"></param>
            ''' <param name="cvrowcnt"></param>
            ''' <param name="conditioncnt"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

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

                Dim model_cvitem As New Njc.Model.Jisyadata_koza_BaseMid_Model  '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call Me.Get_RelData_KozaSyubetu()

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If

                'テーブル名/フィールド名セット
                Dim tblname_base As String = "jisyadata"
                Dim tblname As String = "jisyadata_koza"
                Dim fldnamegrp As String = "jisya_no,jisya_kozano,kinyu_no,kinyu_tenno,koza_syubetu," & _
                                           "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," & _
                                           "yucyokoza_bango,biko_koza,history,useflg"

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

                ''ヘッダー格納用
                'Dim fldname_keymain As String = ""
                'Dim fldname_keysub As String = ""

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
                            Return rtn
                        End If

                        '作業用変数
                        Dim fldname As String = ""
                        Dim fldvalue As String = ""

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

                            Select Case fldname
                                Case "自社・支店No"
                                    .Vari_Jisya_no = fldvalue.Trim
                                Case "自社・支店口座No"
                                    .Vari_Jisya_kozano = fldvalue.Trim
                                Case "金融機関No"
                                    .Vari_Kinyu_no = fldvalue.Trim
                                Case "金融機関店No"
                                    .Vari_Kinyu_tenno = fldvalue.Trim
                                Case "口座種別"
                                    .Vari_Koza_syubetu = fldvalue.Trim
                                Case "口座番号"
                                    .Vari_Koza_bango = fldvalue.Trim
                                Case "口座名義"
                                    .Vari_Koza_meigi = fldvalue.Trim
                                Case "口座名義カナ"
                                    .Vari_Koza_meigikana = fldvalue.Trim
                                Case "ゆうちょ口座記号１"
                                    .Vari_Yucyokoza_kigo1 = fldvalue.Trim
                                Case "ゆうちょ口座記号２"
                                    .Vari_Yucyokoza_kigo2 = fldvalue.Trim
                                Case "ゆうちょ口座番号"
                                    .Vari_Yucyokoza_bango = fldvalue.Trim
                                Case "備考(口座情報)"
                                    .Vari_Biko_koza = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory
                        .Vari_Useflg = 1

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20161125 自社口座情報で仮口座作成時にゆうちょが存在する場合は設定しないようにする処理を追加 -add sta
                        'ゆうちょ関連情報にデータが存在する場合は設定しておいた仮の金融機関とログを削除する
                        Dim tmp_yucyoinfo As String = .Vari_Yucyokoza_kigo1 & .Vari_Yucyokoza_kigo2 & .Vari_Yucyokoza_bango
                        If tmp_yucyoinfo <> "" Then
                            'ゆうちょ銀行にデータがある場合はそちらを優先する
                            '金融機関、支店Noを削除
                            hash_cvitem("kinyu_no") = ""
                            hash_cvitem("kinyu_tenno") = ""

                            'ログを削除
                            hash_log.Remove("jisyadata_koza-kinyu_no")
                            hash_log.Remove("jisyadata_koza-kinyu_tenno")
                        End If
                        '20161125 自社口座情報で仮口座作成時にゆうちょが存在する場合は設定しないようにする処理を追加 -add end

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
            '20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -add end

            '20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -del sta
            ' ''' <summary>
            ' ''' 【中間ファイル→変数】
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="filename"></param>
            ' ''' <param name="sheetname"></param>
            ' ''' <param name="midrowcnt"></param>
            ' ''' <param name="cvrowcnt"></param>
            ' ''' <param name="conditioncnt"></param>
            ' ''' <returns></returns>
            ' ''' <remarks></remarks>
            'Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim tmp_cvcnt As Integer                                        '移行件数格納
            '    Dim tmp_condcnt As Integer                                      '調整件数格納

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            '    Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            '    Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
            '    Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
            '    Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
            '    Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
            '    Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

            '    Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
            '    Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            '    Dim rtn As Boolean = True                                       '戻り値

            '    Dim model_cvitem As New Njc.Model.Jisyadata_koza_BaseMid_Model  '移行値格納用モデル初期化
            '    Dim keycol_main As Integer = 1                                  'メインキー列
            '    Dim keycol_sub As Integer = 2                                   'サブキー列

            '    '************************
            '    '作業準備
            '    '************************

            '    '紐付けデータ取得
            '    Call Me.Get_RelData_KozaSyubetu()

            '    'Excelファイル初期設定
            '    '汎用コンバートの自社口座は汎用用中間ファイルから取得する
            '    Dim readfilename As String = CV_FROM_MIDDLE            '汎用用中間ファイル名
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, BaseMidDirPath, readfilename, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If Not rtn Then
            '        Return rtn
            '    End If

            '    'テーブル名/フィールド名セット
            '    Dim tblname_base As String = "jisyadata"
            '    Dim tblname As String = "jisyadata_koza"
            '    Dim fldnamegrp As String = "jisya_no,jisya_kozano,kinyu_no,kinyu_tenno,koza_syubetu," & _
            '                               "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," & _
            '                               "yucyokoza_bango,biko_koza,history,useflg"

            '    '親マスタ取得
            '    Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

            '    '追加コンバート時の重複チェック用に既存データのキーを取得
            '    If Not InitDBFlg Then
            '        Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
            '    End If

            '    'プログレスバー初期化
            '    Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
            '    Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

            '    If rowcnt <= pgbbasecnt Then
            '        pgbtotalcnt = rowcnt
            '    Else
            '        pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
            '    End If
            '    Call obj_pgb.pgbInitPart(pgbtotalcnt)

            '    '************************
            '    '処理開始
            '    '************************
            '    With model_cvitem

            '        'ヘッダー行取得
            '        Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
            '        Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 2), wsheet.Cells(headerrow, columncnt)).Value

            '        'キーヘッダー名取得
            '        Dim fldname_keymain As String = headervalue(1, keycol_main)
            '        Dim fldname_keysub As String = headervalue(1, keycol_sub)

            '        '---------------
            '        'データ部処理
            '        '---------------
            '        For cntii = 1 To rowcnt

            '            '中断処理
            '            Application.DoEvents()
            '            If CancelFlg Then
            '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            '                Return rtn
            '            End If

            '            '行取得
            '            Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 2), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

            '            'キー値取得
            '            Dim fldvalue_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(1, keycol_main))
            '            Dim fldvalue_keysub As String = EtcMethod.Get_ShapeStr(datarowvalue(1, keycol_sub))
            '            Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

            '            'ログ出力用
            '            Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

            '            '移行値取得
            '            For cntjj = 1 To columncnt - 1

            '                Dim fldname As String = headervalue(1, cntjj).ToString.Trim
            '                Dim fldvalue As String = ""
            '                If datarowvalue(1, cntjj) IsNot Nothing Then
            '                    fldvalue = datarowvalue(1, cntjj).ToString.Trim
            '                End If

            '                Select Case fldname
            '                    Case "自社・支店No"
            '                        .Vari_Jisya_no = fldvalue.Trim
            '                    Case "自社・支店口座No"
            '                        .Vari_Jisya_kozano = fldvalue.Trim
            '                    Case "金融機関No"
            '                        .Vari_Kinyu_no = fldvalue.Trim
            '                    Case "金融機関店No"
            '                        .Vari_Kinyu_tenno = fldvalue.Trim
            '                    Case "口座種別"
            '                        .Vari_Koza_syubetu = fldvalue.Trim
            '                    Case "口座番号"
            '                        .Vari_Koza_bango = fldvalue.Trim
            '                    Case "口座名義"
            '                        .Vari_Koza_meigi = fldvalue.Trim
            '                    Case "口座名義カナ"
            '                        .Vari_Koza_meigikana = fldvalue.Trim
            '                    Case "ゆうちょ口座記号１"
            '                        .Vari_Yucyokoza_kigo1 = fldvalue.Trim
            '                    Case "ゆうちょ口座記号２"
            '                        .Vari_Yucyokoza_kigo2 = fldvalue.Trim
            '                    Case "ゆうちょ口座番号"
            '                        .Vari_Yucyokoza_bango = fldvalue.Trim
            '                    Case "備考(口座情報)"
            '                        .Vari_Biko_koza = fldvalue.Trim
            '                End Select

            '            Next

            '            '固定値
            '            .Vari_History = DefHistory
            '            .Vari_Useflg = 1

            '            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            '            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '            'データチェック
            '            Dim skipflg As Boolean = False
            '            Dim hash_cvitem As New Hashtable
            '            Dim hash_log As New Hashtable
            '            skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

            '            '書込処理
            '            If Not skipflg Then

            '                '挿入処理
            '                Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

            '                '挿入処理が正常終了したレコードのキーを重複チェック用に格納
            '                If normalflg Then
            '                    list_chkduplicate.Add(fldvalue_key)
            '                    tmp_cvcnt = tmp_cvcnt + 1
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

            '            '-------------------------------
            '            'プログレスバー更新/進捗率表示
            '            '-------------------------------
            '            '件数取得
            '            Dim pgbcnt As Integer = 0
            '            If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
            '                pgbcnt = cntii
            '            ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
            '                Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
            '            ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
            '                pgbcnt = pgbtotalcnt
            '            End If

            '            '表示
            '            If pgbcnt <> 0 Then
            '                Call obj_pgb.pgbsettingPart(pgbcnt)
            '                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
            '            End If

            '        Next

            '    End With

            '    '************************
            '    '終了処理
            '    '************************

            '    '中間ファイル件数を取得
            '    midrowcnt = rowcnt

            '    '移行件数を取得
            '    cvrowcnt = tmp_cvcnt

            '    '調整件数を取得
            '    conditioncnt = tmp_condcnt

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            '    '返却
            '    Return rtn

            'End Function
            '20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -del end

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT jisya_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,ow_no) + '-' + CONVERT(varchar,ow_kozano) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 口座種別紐付情報取得
            ''' </summary>
            ''' <remarks></remarks>
            Public Sub Get_RelData_KozaSyubetu()

                Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
                Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
                Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
                Dim startrow As Integer                                         '書込開始行
                Dim columncnt As Integer                                        '列数
                Dim maxrowcnt As Integer                                        '既存データの行数
                Dim rowcnt As Integer                                           '書込行数
                Dim relsheetname As String = "口座種別マスタ"
                Dim rtn As Boolean = True

                Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

                '既存データ有無確認(口座種別は他の項目でも参照するため)
                If Hash_Rel_Kozasyubetu.Count <> 0 Then
                    Exit Sub
                End If

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

                For cntii = 0 To rowcnt - 1

                    '作業用変数作成
                    Dim tmp_oldkozasyubetuname As String = ""
                    Dim tmp_newkozasyubetuno As String = ""

                    '紐付設定値取得
                    For cntjj = 1 To columncnt

                        'ヘッダー格納
                        Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

                        '移行値格納
                        Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

                        Select Case fldname
                            Case "移行元口座種別名称"
                                tmp_oldkozasyubetuname = fldvalue.Trim
                            Case "賃貸革命10口座種別No"
                                tmp_newkozasyubetuno = fldvalue.Trim
                        End Select

                    Next

                    'ハッシュテーブル格納
                    If tmp_newkozasyubetuno <> "" Then
                        Hash_Rel_Kozasyubetu.Add(tmp_oldkozasyubetuname, tmp_newkozasyubetuno)
                    End If

                Next

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            End Sub

        End Class

    End Class

#End Region

#Region "部屋設備情報(汎用用)"

    Public Class Hydata_setubilst_BaseMid_Repository

        Public Class SubConv
            Implements IConv
            '20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -add sta
            ''' <summary>
            ''' 【中間ファイル→変数】 
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
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
                Dim hash_hyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Hydata_setubilst_BaseMid_Model    '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                  'サブ1キー列

                '************************
                '作業準備
                '************************

                '革命10から設備マスタを取得
                Dim hash_setubimst As New Hashtable
                Call EtcMethod.Set_SetubiMst_BaseMId(sqlcnnv10, hash_setubimst)

                'データ取得
                Dim tmptblname As String = PRE_TBL_NAME & sheetname
                Dim tmp_sql As String = " SELECT * FROM " & tmptblname
                Dim readtbl As New DataTable
                Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

                'オープン処理失敗時は処理を抜ける
                If rtn = False Then
                    Return rtn
                End If

                'guid取得
                Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

                'テーブル名/フィールド名セット
                Dim tblname As String = "hydata_setubilst"
                Dim fldnamegrp As String = "hy_guid,komok_guid,override_kbn,komok_name,disp1name," & _
                                           "disp2name,disp3name,disp1iconguid,disp2iconguid,disp3iconguid," & _
                                           "history"

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

                '移行対象件数取得用
                Dim cvtaisyocnt As Integer = 0

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
                            Return rtn
                        End If

                        '作業用変数
                        Dim hash_log As New Hashtable
                        Dim hash_cvitem As New Hashtable
                        Dim setubicvcnt As Integer = 0

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim tmp_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub - 1)).Trim
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                                                   fldname_keysub & " = " & tmp_keysub

                        '移行判別用
                        Dim cvflg As Boolean = True
                        Dim keychkflg As Boolean = True

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

                            '親データ有無チェック
                            If list_basekeydata.Contains(fldvalue_key) = False Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "hydata_setubilst-hy_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_log.Clear()
                                hash_log.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(fldvalue_key) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "hydata_setubilst-hy_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
                                hash_log.Clear()
                                hash_log.Add(log_key, log_value)
                            End If

                            'データ取得、移行処理
                            If keychkflg Then

                                '作業用変数
                                Dim fldname As String = ""
                                Dim fldvalue As String = ""

                                '移行値取得
                                For cntjj = 2 To readtbl.Columns.Count - 1

                                    '項目名取得
                                    fldname = readtbl.Columns(cntjj).ColumnName.Trim

                                    '登録値取得
                                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                                    Dim value_total As String = fldname.Replace("-", STR_SPLIT_1) & STR_SPLIT_1 & fldvalue

                                    '20170116 部屋設備情報未設定項目のログ出力制御修正 -chg sta
                                    '項目が存在するかで処理を分岐するように修正する
                                    ''設備マスターと一致した場合
                                    'If hash_setubimst.Contains(value_total) Then

                                    '    '項目guidを取得
                                    '    .Vari_Komok_guid = hash_setubimst(value_total)

                                    '    '固定値
                                    '    .Vari_Override_kbn = 0
                                    '    .Vari_Komok_name = ""
                                    '    .Vari_Disp1name = ""
                                    '    .Vari_Disp2name = ""
                                    '    .Vari_Disp3name = ""
                                    '    .Vari_Disp1iconguid = "00000000-0000-0000-0000-000000000000"
                                    '    .Vari_Disp2iconguid = "00000000-0000-0000-0000-000000000000"
                                    '    .Vari_Disp3iconguid = "00000000-0000-0000-0000-000000000000"
                                    '    .Vari_History = DefHistory

                                    '    '部屋guid変換
                                    '    .Vari_Hy_guid = hash_hyguid(fldvalue_key)

                                    '    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                                    '    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                                    '    '挿入処理
                                    '    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, normalflg)

                                    '    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                    '    If normalflg Then
                                    '        If list_chkduplicate.Contains(fldvalue_key) = False Then
                                    '            list_chkduplicate.Add(fldvalue_key)
                                    '        End If
                                    '        setubicvcnt = setubicvcnt + 1
                                    '    End If

                                    'Else

                                    '    'マスタ不一致ログ出力
                                    '    Dim log_key As String = "hydata_setubilst-komok_guid"
                                    '    Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & "設備" & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                                    '    Dim hash_komklog As New Hashtable From {{log_key, log_value}}

                                    '    'ログ出力メッセージ整形
                                    '    tmp_hash.Clear()
                                    '    hash_cvitem.Clear()
                                    '    Call LogSetting.Set_Log_Value_KomkErr(hash_komklog, tmp_hash, hash_cvitem, str_logkey & "、" & fldname, tblname, tmp_logcnt, sortlist_log)

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
                                    If fldvalue <> "" Then

                                        '設備マスターと一致した場合
                                        If hash_setubimst.Contains(value_total) Then

                                            '項目guidを取得
                                            .Vari_Komok_guid = hash_setubimst(value_total)

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

                                            '部屋guid変換
                                            .Vari_Hy_guid = hash_hyguid(fldvalue_key)

                                            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                                            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                                            '挿入処理
                                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, normalflg)

                                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                            If normalflg Then
                                                If list_chkduplicate.Contains(fldvalue_key) = False Then
                                                    list_chkduplicate.Add(fldvalue_key)
                                                End If
                                                setubicvcnt = setubicvcnt + 1
                                            End If

                                        Else

                                            'マスタ不一致ログ出力
                                            Dim log_key As String = "hydata_setubilst-komok_guid"
                                            Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & "設備" & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                                            Dim hash_komklog As New Hashtable From {{log_key, log_value}}

                                            'ログ出力メッセージ整形
                                            tmp_hash.Clear()
                                            hash_cvitem.Clear()
                                            Call LogSetting.Set_Log_Value_KomkErr(hash_komklog, tmp_hash, hash_cvitem, str_logkey & "、" & fldname, tblname, tmp_logcnt, sortlist_log)

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

                                    End If
                                    '20170116 部屋設備情報未設定項目のログ出力制御修正 -chg end
                                Next

                            Else

                                '親データ無し、重複チェックエラーログ出力
                                'ログ出力メッセージ整形
                                tmp_hash.Clear()
                                hash_cvitem.Clear()
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

                        End If

                        '-------------------------------
                        '件数取得
                        '-------------------------------
                        If setubicvcnt > 0 Then
                            tmp_cvcnt = tmp_cvcnt + 1
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
                midrowcnt = cvtaisyocnt

                '移行件数を取得
                cvrowcnt = tmp_cvcnt

                '調整件数を取得
                conditioncnt = tmp_condcnt

                '返却
                Return rtn

            End Function
            '20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -add end

            '20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -del sta
            '        ''' <summary>
            '        ''' 【中間ファイル→変数】 
            '        ''' </summary>
            '        ''' <param name="sqlcnnv10"></param>
            '        ''' <returns></returns>
            '        ''' <remarks></remarks>
            '        Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

            '            Dim tmp_cvcnt As Integer                                        '移行件数格納
            '            Dim tmp_condcnt As Integer                                      '調整件数格納
            '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            '            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            '            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
            '            Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
            '            Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
            '            Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
            '            Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
            '            Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
            '            Dim hash_hyguid As New Hashtable                            'キー/guid格納用ハッシュテーブル
            '            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            '            Dim rtn As Boolean = True                                       '戻り値

            '            Dim model_cvitem As New Njc.Model.Hydata_setubilst_BaseMid_Model    '移行値格納用モデル初期化
            '            Dim keycol_main As Integer = 1                                  'メインキー列
            '            Dim keycol_sub As Integer = 2                                  'サブ1キー列

            '            '************************
            '            '作業準備
            '            '************************

            '            '革命10から設備マスタを取得
            '            Dim hash_setubimst As New Hashtable
            '            Call EtcMethod.Set_SetubiMst_BaseMId(sqlcnnv10, hash_setubimst)

            '            'Excelファイル初期設定                  
            '            Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
            '            Dim readtbl As New DataTable()
            '            Dim con_read As New OleDbConnection()

            '            'オープン処理
            '            Dim readfilename As String = CV_FROM_MIDDLE            '汎用用中間ファイル名
            '            rtn = excelfile.ExcelFile_ReadOpen(BaseMidDirPath, readfilename, tmp_sql, readtbl, con_read)

            '            'オープン処理失敗時は処理を抜ける
            '            If rtn = False Then
            '                excelfile.ExcelFile_ReadClose(con_read)
            '                Return rtn
            '            End If

            '            '行数取得
            '            Dim totalrowcnt As Integer = readtbl.Rows.Count     '全件数
            '            Dim rowcnt As Integer = readtbl.Rows.Count - 10     'データ部件数

            '            'guid取得
            '            Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

            '            'テーブル名/フィールド名セット
            '            Dim tblname As String = "hydata_setubilst"
            '            Dim fldnamegrp As String = "hy_guid,komok_guid,override_kbn,komok_name,disp1name," & _
            '                                       "disp2name,disp3name,disp1iconguid,disp2iconguid,disp3iconguid," & _
            '                                       "history"

            '            '追加コンバート時の重複チェック用に既存データのキーを取得
            '            'If Not InitDBFlg Then
            '            '    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
            '            'End If

            '            '親マスタ取得
            '            Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

            '            'プログレスバー初期化
            '            Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
            '            Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

            '            If totalrowcnt <= pgbbasecnt Then
            '                pgbtotalcnt = totalrowcnt
            '            Else
            '                pgbtotalcnt = Math.Ceiling(totalrowcnt / pgbbasecnt)
            '            End If
            '            Call obj_pgb.pgbInitPart(pgbtotalcnt)

            '            'ヘッダー格納用
            '            Dim fldname_keymain As String = ""
            '            Dim fldname_keysub As String = ""

            '            '移行対象件数取得用
            '            Dim cvtaisyocnt As Integer = 0

            '            '************************
            '            '処理開始
            '            '************************
            '            With model_cvitem

            '                '-------------------------------------------------------------------
            '                'データ部処理(汎用中間ファイルからはヘッダーを取得できないので行Noから取得する)
            '                '-------------------------------------------------------------------
            '                For cntii = 0 To totalrowcnt - 1

            '                    '中断処理
            '                    Application.DoEvents()
            '                    If CancelFlg Then
            '                        Call excelfile.ExcelFile_ReadClose(con_read)
            '                        Return rtn
            '                    End If

            '                    'キー格納用
            '                    Dim fldvalue_key As String = ""
            '                    Dim str_logkey As String = ""
            '                    Dim str_taihilogkey As String = ""
            '                    Dim str_taihilog As String = ""
            '                    Dim setubicvcnt As Integer = 0

            '                    Dim hash_log As New Hashtable
            '                    Dim hash_cvitem As New Hashtable

            '                    '移行判別用
            '                    Dim cvflg As Boolean = True
            '                    Dim keychkflg As Boolean = True

            '                    If cntii = 0 Then

            '                        'ヘッダー取得
            '                        fldname_keymain = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main)).Trim
            '                        fldname_keysub = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub)).Trim

            '                    ElseIf cntii >= BASEMIDFILE_READWRITE_ROW - 2 Then

            '                        'キー値取得
            '                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main)).Trim
            '                        Dim tmp_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub)).Trim
            '                        fldvalue_key = tmp_keymain & "-" & tmp_keysub
            '                        '.Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub

            '                        'ログ出力用
            '                        str_logkey = fldname_keymain & " = " & tmp_keymain & "、" & _
            '                                     fldname_keysub & " = " & tmp_keysub

            '                        'データ有無チェック
            '                        Dim tmp_fldvalueumuchk As String = ""
            '                        For cntjj = 3 To readtbl.Columns.Count - 1
            '                            tmp_fldvalueumuchk = tmp_fldvalueumuchk & Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
            '                        Next
            '                        If tmp_fldvalueumuchk = "" Then
            '                            cvflg = False
            '                        End If

            '                        If cvflg Then

            '                            '移行対象件数カウント
            '                            cvtaisyocnt = cvtaisyocnt + 1

            '                            '親データ有無チェック
            '                            If list_basekeydata.Contains(fldvalue_key) = False Then
            '                                keychkflg = False
            '                                'ログ出力
            '                                Dim log_key As String = "hydata_setubilst-hy_guid"
            '                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
            '                                hash_log.Clear()
            '                                hash_log.Add(log_key, log_value)
            '                            End If

            '                            '重複チェック
            '                            If keychkflg And list_chkduplicate.Contains(fldvalue_key) Then
            '                                keychkflg = False
            '                                'ログ出力
            '                                Dim log_key As String = "hydata_setubilst-hy_guid"
            '                                Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
            '                                hash_log.Clear()
            '                                hash_log.Add(log_key, log_value)
            '                            End If

            '                            'データ取得、移行処理
            '                            If keychkflg Then

            '                                '作業用変数
            '                                Dim tmp_fldname As String = ""
            '                                Dim fldname_syogo As String = ""
            '                                Dim fldname_log As String = ""
            '                                Dim fldvalue As String = ""

            '                                '移行値取得
            '                                For cntjj = 0 To readtbl.Columns.Count - 1

            '                                    '項目名取得
            '                                    tmp_fldname = Typ.ToStr(readtbl.Rows(0).Item(cntjj)).Trim
            '                                    fldname_log = tmp_fldname.Replace(vbLf, "-")
            '                                    fldname_syogo = tmp_fldname.Replace(vbLf, STR_SPLIT_1)

            '                                    '登録値取得
            '                                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

            '                                    If fldname_syogo <> "項目名" And fldname_syogo <> "物件NO" And fldname_syogo <> "部屋NO" And fldvalue <> "" Then

            '                                        Dim value_total As String = fldname_syogo & STR_SPLIT_1 & fldvalue

            '                                        '設備マスターと一致した場合
            '                                        If hash_setubimst.Contains(value_total) Then

            '                                            '項目guidを取得
            '                                            .Vari_Komok_guid = hash_setubimst(value_total)

            '                                            '固定値
            '                                            .Vari_Override_kbn = 0
            '                                            .Vari_Komok_name = ""
            '                                            .Vari_Disp1name = ""
            '                                            .Vari_Disp2name = ""
            '                                            .Vari_Disp3name = ""
            '                                            .Vari_Disp1iconguid = "00000000-0000-0000-0000-000000000000"
            '                                            .Vari_Disp2iconguid = "00000000-0000-0000-0000-000000000000"
            '                                            .Vari_Disp3iconguid = "00000000-0000-0000-0000-000000000000"
            '                                            .Vari_History = DefHistory

            '                                            '部屋guid変換
            '                                            .Vari_Hy_guid = hash_hyguid(fldvalue_key)

            '                                            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            '                                            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '                                            '挿入処理
            '                                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, normalflg)

            '                                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
            '                                            If normalflg Then
            '                                                If list_chkduplicate.Contains(fldvalue_key) = False Then
            '                                                    list_chkduplicate.Add(fldvalue_key)
            '                                                End If
            '                                                setubicvcnt = setubicvcnt + 1
            '                                            End If

            '                                        Else

            '                                            'マスタ不一致ログ出力
            '                                            Dim log_key As String = "hydata_setubilst-komok_guid"
            '                                            Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & "設備" & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            '                                            Dim hash_komklog As New Hashtable From {{log_key, log_value}}

            '                                            'ログ出力メッセージ整形
            '                                            tmp_hash.Clear()
            '                                            hash_cvitem.Clear()
            '                                            Call LogSetting.Set_Log_Value_KomkErr(hash_komklog, tmp_hash, hash_cvitem, str_logkey & "、" & fldname_log, tblname, tmp_logcnt, sortlist_log)

            '                                            'ログ出力
            '                                            If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = totalrowcnt - 1) Then
            '                                                Dim tmp_cnt As Integer = 0
            '                                                '挿入
            '                                                For Each logvalue In sortlist_log
            '                                                    Dim tmp_sql_insert As String = logvalue.Value
            '                                                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
            '                                                Next
            '                                                '初期化
            '                                                tmp_logcnt = 0
            '                                                sortlist_log.Clear()
            '                                            End If

            '                                        End If

            '                                    End If

            '                                Next

            '                            Else

            '                                '親データ無し、重複チェックエラーログ出力
            '                                'ログ出力メッセージ整形
            '                                tmp_hash.Clear()
            '                                hash_cvitem.Clear()
            '                                Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

            '                                'ログ出力
            '                                If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = totalrowcnt - 1) Then
            '                                    Dim tmp_cnt As Integer = 0
            '                                    '挿入
            '                                    For Each logvalue In sortlist_log
            '                                        Dim tmp_sql_insert As String = logvalue.Value
            '                                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
            '                                    Next
            '                                    '初期化
            '                                    tmp_logcnt = 0
            '                                    sortlist_log.Clear()
            '                                End If

            '                            End If

            '                        End If

            '                    End If

            '                    '-------------------------------
            '                    '件数取得
            '                    '-------------------------------
            '                    If setubicvcnt > 0 Then
            '                        tmp_cvcnt = tmp_cvcnt + 1
            '                    End If

            '                    '-------------------------------
            '                    'プログレスバー更新/進捗率表示
            '                    '-------------------------------
            '                    '件数取得
            '                    'Dim pgbcnt As Integer = 0
            '                    'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
            '                    '    pgbcnt = cntii
            '                    'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
            '                    '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
            '                    'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
            '                    '    pgbcnt = pgbtotalcnt
            '                    'End If

            '                    Dim pgbcnt As Integer = 0
            '                    If totalrowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
            '                        pgbcnt = cntii + 1
            '                    ElseIf totalrowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
            '                        Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
            '                    ElseIf totalrowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
            '                        pgbcnt = pgbtotalcnt
            '                    End If

            '                    '表示
            '                    If pgbcnt <> 0 Then
            '                        Call obj_pgb.pgbsettingPart(pgbcnt)
            '                        Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
            '                    End If

            '                Next

            '            End With

            '            '************************
            '            '終了処理
            '            '************************

            '            '中間ファイル件数を取得
            '            midrowcnt = cvtaisyocnt

            '            '移行件数を取得
            '            cvrowcnt = tmp_cvcnt

            '            '調整件数を取得
            '            conditioncnt = tmp_condcnt

            '            'クローズ処理
            '            Call excelfile.ExcelFile_ReadClose(con_read)

            '            '返却
            '            Return rtn

            '        End Function
            '20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -del end

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
            ''' 部屋guidの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="hash"></param>
            ''' <remarks></remarks>
            Public Sub Set_HyGuid(ByVal sqlcnnv10 As SqlConnection, ByRef hash As Hashtable)

                Dim tmp_sql As String = ""
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 CONVERT(VARCHAR,bk_no) + '-' + CONVERT(VARCHAR,hy_no) "
                tmp_sql = tmp_sql & " 	,hy_guid "
                tmp_sql = tmp_sql & " FROM hydata AS HY "
                tmp_sql = tmp_sql & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash)

            End Sub

        End Class

    End Class

#End Region
    '20161028 物件/部屋鍵取得方法修正 -add sta
#Region "鍵タイトルマスタ(汎用用)"

    Public Class M_kagi_title_BaseMid_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】 
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="filename"></param>
            ''' <param name="sheetname"></param>
            ''' <param name="midrowcnt"></param>
            ''' <param name="cvrowcnt"></param>
            ''' <param name="conditioncnt"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

                Dim tmp_cvcnt As Integer                                        '移行件数格納
                Dim tmp_condcnt As Integer                                      '調整件数格納
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値
                Dim tmpcnt As Integer = 0                                       'クエリ実行用作業変数

                Dim model_cvitem As New Njc.Model.M_kagi_title_BaseMid_Model            '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列

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

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_kagi_title"
                Dim fldnamegrp As String = "kagi_kbn,kagi_no,kagi_name,history"

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

                '鍵タイトルマスタをデフォルト値で一括更新する
                If rowcnt > 0 Then
                    Dim tmp_sql_kagititle_def As String = " UPDATE m_kagi_title SET kagi_name = '鍵' + CONVERT(VARCHAR,kagi_no) WHERE kagi_kbn IN (1,2) "
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_kagititle_def, tmpcnt)
                End If

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol_main - 1).ColumnName.Trim
                    Dim fldname_keysub As String = readtbl.Columns(keycol_sub - 1).ColumnName.Trim

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
                        Dim fldvalue_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        Dim fldvalue_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub - 1)).Trim
                        Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub
                        '鍵区分はプログラム内部で使用する項目のためログ出力用に成形する
                        str_logkey = str_logkey.Replace("鍵区分 = 1", "物件鍵タイトルマスタ")
                        str_logkey = str_logkey.Replace("鍵区分 = 2", "部屋鍵タイトルマスタ")

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            Select Case fldname
                                Case "鍵区分"
                                    .Vari_Kagi_kbn = fldvalue
                                Case "鍵No"
                                    .Vari_Kagi_no = fldvalue
                                Case "鍵名称"
                                    .Vari_Kagi_name = fldvalue
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
                        If skipflg = False Then

                            '更新処理
                            Dim tmp_sql_kagititle As String = " UPDATE m_kagi_title SET kagi_name = '" & hash_cvitem("kagi_name") & "' WHERE kagi_kbn = " & hash_cvitem("kagi_kbn") & " AND kagi_no = " & hash_cvitem("kagi_no")
                            normalflg = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_kagititle, tmpcnt)

                            '更新処理が正常終了したレコードのキーを重複チェック用に格納
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
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT jisya_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,kagi_kbn) + '-' + CONVERT(varchar,kagi_no) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region
    '20161028 物件/部屋鍵取得方法修正 -add end
End Namespace


