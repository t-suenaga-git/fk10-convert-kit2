Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "自社情報"

    Public Class Jisyadata_Repository

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

                Dim model_cvitem As New Njc.Model.Jisyadata_Model               '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列

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
                'テーブル名/フィールド名セット
                Dim tblname As String = "jisyadata"
                Dim fldnamegrp As String = "jisya_no,jisya_name,jisya_namesjis,jisya_kana,jisya_name2," & _
                                           "jisya_name2sjis,jisya_name3,jisya_name3sjis,daihyo_yakusyoku,daihyo_name," & _
                                           "post_code,addr_kenno,addr_sino,addr_cyo,addr_cyome," & _
                                           "addr_cyomeptn,addr_banti,addr_etc,tel1,tel2," & _
                                           "mobiletel1,mobiletel2,fax,mail,mobilemail," & _
                                           "url,biko_kihon,menkyo_bango,menkyo_ymd,syunin_bango," & _
                                           "syunin_name,kanrikyokai_bango,keieikanrisi_bango,keieikanrisi_name,kanrigy_bango," & _
                                           "kanyu_dantai1,kanyu_dantai2,kanyu_dantai3,kanyu_dantai4,kanyu_dantai5," & _
                                           "kanyu_dantai6,kanyu_dantai7,kanyu_dantai8,kanyu_dantai9,kanyu_dantai10," & _
                                           "history,useflg,daihyo_namesjis,syunin_namesjis,keieikanrisi_namesjis"

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

                    'キーヘッダー名取得
                    Dim fldname_key As String = readtbl.Columns(keycol - 1).ColumnName.Trim

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
                        Dim tmp_address1 As String = ""
                        Dim tmp_address2 As String = ""

                        'キー値取得
                        Dim fldvalue_key As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol - 1)).Trim

                        'ログ出力用
                        Dim str_logkey As String = fldname_key & " = " & fldvalue_key

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "自社・支店No"
                                    .Vari_Jisya_no = fldvalue.Trim
                                Case "自社・支店名"
                                    .Vari_Jisya_name = fldvalue.Trim
                                Case "自社・支店名（SJIS）"
                                    .Vari_Jisya_namesjis = fldvalue.Trim
                                Case "自社・支店カナ"
                                    .Vari_Jisya_kana = fldvalue.Trim
                                Case "自社・支店名２"
                                    .Vari_Jisya_name2 = fldvalue.Trim
                                Case "自社・支店名２（SJIS）"
                                    .Vari_Jisya_name2sjis = fldvalue.Trim
                                Case "自社・支店名３"
                                    .Vari_Jisya_name3 = fldvalue.Trim
                                Case "自社・支店名３（SJIS）"
                                    .Vari_Jisya_name3sjis = fldvalue.Trim
                                Case "代表者役職"
                                    .Vari_Daihyo_yakusyoku = fldvalue.Trim
                                Case "代表者名"
                                    .Vari_Daihyo_name = fldvalue.Trim
                                Case "郵便番号"
                                    .Vari_Post_code = fldvalue.Trim
                                Case "住所1"
                                    tmp_address1 = fldvalue.Trim
                                Case "住所2"
                                    tmp_address2 = fldvalue.Trim
                                Case "TEL1"
                                    .Vari_Tel1 = fldvalue.Trim
                                Case "TEL2"
                                    .Vari_Tel2 = fldvalue.Trim
                                Case "携帯TEL1"
                                    .Vari_Mobiletel1 = fldvalue.Trim
                                Case "携帯TEL2"
                                    .Vari_Mobiletel2 = fldvalue.Trim
                                Case "FAX"
                                    .Vari_Fax = fldvalue.Trim
                                Case "メールアドレス"
                                    .Vari_Mail = fldvalue.Trim
                                Case "携帯メールアドレス"
                                    .Vari_Mobilemail = fldvalue.Trim
                                Case "Webアドレス"
                                    .Vari_Url = fldvalue.Trim
                                Case "備考（基本情報）"
                                    .Vari_Biko_kihon = fldvalue.Trim
                                Case "免許証番号"
                                    .Vari_Menkyo_bango = fldvalue.Trim
                                Case "免許年月日"
                                    .Vari_Menkyo_ymd = fldvalue.Trim
                                Case "取引主任者登録番号（代表）"
                                    .Vari_Syunin_bango = fldvalue.Trim
                                Case "取引主任者名（代表）"
                                    .Vari_Syunin_name = fldvalue.Trim
                                Case "全国賃貸不動産管理業協会会員番号"
                                    .Vari_Kanrikyokai_bango = fldvalue.Trim
                                Case "賃貸不動産経営管理士登録番号"
                                    .Vari_Keieikanrisi_bango = fldvalue.Trim
                                Case "賃貸不動産経営管理士名"
                                    .Vari_Keieikanrisi_name = fldvalue.Trim
                                Case "賃貸住宅管理業者登録番号"
                                    .Vari_Kanrigy_bango = fldvalue.Trim
                                Case "加入団体１"
                                    .Vari_Kanyu_dantai1 = fldvalue.Trim
                                Case "加入団体２"
                                    .Vari_Kanyu_dantai2 = fldvalue.Trim
                                Case "加入団体３"
                                    .Vari_Kanyu_dantai3 = fldvalue.Trim
                                Case "加入団体４"
                                    .Vari_Kanyu_dantai4 = fldvalue.Trim
                                Case "加入団体５"
                                    .Vari_Kanyu_dantai5 = fldvalue.Trim
                                Case "加入団体６"
                                    .Vari_Kanyu_dantai6 = fldvalue.Trim
                                Case "加入団体７"
                                    .Vari_Kanyu_dantai7 = fldvalue.Trim
                                Case "加入団体８"
                                    .Vari_Kanyu_dantai8 = fldvalue.Trim
                                Case "加入団体９"
                                    .Vari_Kanyu_dantai9 = fldvalue.Trim
                                Case "加入団体１０"
                                    .Vari_Kanyu_dantai10 = fldvalue.Trim
                                Case "代表者名（SJIS）"
                                    .Vari_Daihyo_namesjis = fldvalue.Trim
                                Case "取引主任者名（代表）（SJIS）"
                                    .Vari_Syunin_namesjis = fldvalue.Trim
                                Case "賃貸不動産経営管理士名（SJIS）"
                                    .Vari_Keieikanrisi_namesjis = fldvalue.Trim
                            End Select

                        Next

                        '住所変換/格納
                        Dim tmp_address As String = tmp_address1 & " " & tmp_address2
                        Dim hash_address As New Hashtable
                        '20160622 住所分割処理対応 -chg sta
                        'hash_address = AddressChange.Get_Address(tmp_address, sqlcnnv10)
                        hash_address = AddressChange_old.Get_Address(tmp_address, sqlcnnv10)
                        '20160622 住所分割処理対応 -chg end
                        .Vari_Addr_kenno = hash_address("ken_no")
                        .Vari_Addr_sino = hash_address("si_no")
                        .Vari_Addr_cyo = hash_address("mati")
                        .Vari_Addr_cyome = hash_address("cyome")
                        .Vari_Addr_cyomeptn = hash_address("chomeptn")
                        .Vari_Addr_banti = hash_address("banti")
                        .Vari_Addr_etc = hash_address("etc")

                        '固定値
                        .Vari_Useflg = 1
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
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

            ' ''' <summary>
            ' ''' 【中間ファイル→変数】
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="syorikomok"></param>
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

            '    Dim model_cvitem As New Njc.Model.Jisyadata_Model               '移行値格納用モデル初期化
            '    Dim keycol As Integer = 1                                       'キー列

            '    '************************
            '    '作業準備
            '    '************************

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If Not rtn Then
            '        Return rtn
            '    End If

            '    'テーブル名/フィールド名セット
            '    Dim tblname As String = "jisyadata"
            '    Dim fldnamegrp As String = "jisya_no,jisya_name,jisya_namesjis,jisya_kana,jisya_name2," & _
            '                               "jisya_name2sjis,jisya_name3,jisya_name3sjis,daihyo_yakusyoku,daihyo_name," & _
            '                               "post_code,addr_kenno,addr_sino,addr_cyo,addr_cyome," & _
            '                               "addr_cyomeptn,addr_banti,addr_etc,tel1,tel2," & _
            '                               "mobiletel1,mobiletel2,fax,mail,mobilemail," & _
            '                               "url,biko_kihon,menkyo_bango,menkyo_ymd,syunin_bango," & _
            '                               "syunin_name,kanrikyokai_bango,keieikanrisi_bango,keieikanrisi_name,kanrigy_bango," & _
            '                               "kanyu_dantai1,kanyu_dantai2,kanyu_dantai3,kanyu_dantai4,kanyu_dantai5," & _
            '                               "kanyu_dantai6,kanyu_dantai7,kanyu_dantai8,kanyu_dantai9,kanyu_dantai10," & _
            '                               "history,useflg,daihyo_namesjis,syunin_namesjis,keieikanrisi_namesjis"

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

            '        '---------------
            '        'ヘッダー処理
            '        '---------------
            '        'ヘッダー行取得
            '        Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '        'キーヘッダー名取得
            '        Dim fldname_key As String = headervalue(startrow - 1, keycol)

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
            '            Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

            '            'キー値取得
            '            Dim fldvalue_key As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))

            '            'ログ出力用
            '            Dim str_logkey As String = fldname_key & " = " & fldvalue_key

            '            '作業用変数
            '            Dim tmp_address1 As String = ""
            '            Dim tmp_address2 As String = ""

            '            '移行値取得
            '            For cntjj = 1 To columncnt

            '                Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            '                Dim fldvalue As String = ""
            '                If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            '                End If

            '                Select Case fldname
            '                    Case "自社・支店No"
            '                        .Vari_Jisya_no = fldvalue.Trim
            '                    Case "自社・支店名"
            '                        .Vari_Jisya_name = fldvalue.Trim
            '                    Case "自社・支店名（SJIS）"
            '                        .Vari_Jisya_namesjis = fldvalue.Trim
            '                    Case "自社・支店カナ"
            '                        .Vari_Jisya_kana = fldvalue.Trim
            '                    Case "自社・支店名２"
            '                        .Vari_Jisya_name2 = fldvalue.Trim
            '                    Case "自社・支店名２（SJIS）"
            '                        .Vari_Jisya_name2sjis = fldvalue.Trim
            '                    Case "自社・支店名３"
            '                        .Vari_Jisya_name3 = fldvalue.Trim
            '                    Case "自社・支店名３（SJIS）"
            '                        .Vari_Jisya_name3sjis = fldvalue.Trim
            '                    Case "代表者役職"
            '                        .Vari_Daihyo_yakusyoku = fldvalue.Trim
            '                    Case "代表者名"
            '                        .Vari_Daihyo_name = fldvalue.Trim
            '                    Case "郵便番号"
            '                        .Vari_Post_code = fldvalue.Trim
            '                    Case "住所1"
            '                        tmp_address1 = fldvalue.Trim
            '                    Case "住所2"
            '                        tmp_address2 = fldvalue.Trim
            '                    Case "TEL1"
            '                        .Vari_Tel1 = fldvalue.Trim
            '                    Case "TEL2"
            '                        .Vari_Tel2 = fldvalue.Trim
            '                    Case "携帯TEL1"
            '                        .Vari_Mobiletel1 = fldvalue.Trim
            '                    Case "携帯TEL2"
            '                        .Vari_Mobiletel2 = fldvalue.Trim
            '                    Case "FAX"
            '                        .Vari_Fax = fldvalue.Trim
            '                    Case "メールアドレス"
            '                        .Vari_Mail = fldvalue.Trim
            '                    Case "携帯メールアドレス"
            '                        .Vari_Mobilemail = fldvalue.Trim
            '                    Case "Webアドレス"
            '                        .Vari_Url = fldvalue.Trim
            '                    Case "備考（基本情報）"
            '                        .Vari_Biko_kihon = fldvalue.Trim
            '                    Case "免許証番号"
            '                        .Vari_Menkyo_bango = fldvalue.Trim
            '                    Case "免許年月日"
            '                        .Vari_Menkyo_ymd = fldvalue.Trim
            '                    Case "取引主任者登録番号（代表）"
            '                        .Vari_Syunin_bango = fldvalue.Trim
            '                    Case "取引主任者名（代表）"
            '                        .Vari_Syunin_name = fldvalue.Trim
            '                    Case "全国賃貸不動産管理業協会会員番号"
            '                        .Vari_Kanrikyokai_bango = fldvalue.Trim
            '                    Case "賃貸不動産経営管理士登録番号"
            '                        .Vari_Keieikanrisi_bango = fldvalue.Trim
            '                    Case "賃貸不動産経営管理士名"
            '                        .Vari_Keieikanrisi_name = fldvalue.Trim
            '                    Case "賃貸住宅管理業者登録番号"
            '                        .Vari_Kanrigy_bango = fldvalue.Trim
            '                    Case "加入団体１"
            '                        .Vari_Kanyu_dantai1 = fldvalue.Trim
            '                    Case "加入団体２"
            '                        .Vari_Kanyu_dantai2 = fldvalue.Trim
            '                    Case "加入団体３"
            '                        .Vari_Kanyu_dantai3 = fldvalue.Trim
            '                    Case "加入団体４"
            '                        .Vari_Kanyu_dantai4 = fldvalue.Trim
            '                    Case "加入団体５"
            '                        .Vari_Kanyu_dantai5 = fldvalue.Trim
            '                    Case "加入団体６"
            '                        .Vari_Kanyu_dantai6 = fldvalue.Trim
            '                    Case "加入団体７"
            '                        .Vari_Kanyu_dantai7 = fldvalue.Trim
            '                    Case "加入団体８"
            '                        .Vari_Kanyu_dantai8 = fldvalue.Trim
            '                    Case "加入団体９"
            '                        .Vari_Kanyu_dantai9 = fldvalue.Trim
            '                    Case "加入団体１０"
            '                        .Vari_Kanyu_dantai10 = fldvalue.Trim
            '                    Case "代表者名（SJIS）"
            '                        .Vari_Daihyo_namesjis = fldvalue.Trim
            '                    Case "取引主任者名（代表）（SJIS）"
            '                        .Vari_Syunin_namesjis = fldvalue.Trim
            '                    Case "賃貸不動産経営管理士名（SJIS）"
            '                        .Vari_Keieikanrisi_namesjis = fldvalue.Trim
            '                End Select

            '            Next

            '            '住所変換/格納
            '            Dim tmp_address As String = tmp_address1 & " " & tmp_address2
            '            Dim hash_address As New Hashtable
            '            '20160622 住所分割処理対応 -chg sta
            '            'hash_address = AddressChange.Get_Address(tmp_address, sqlcnnv10)
            '            hash_address = AddressChange_old.Get_Address(tmp_address, sqlcnnv10)
            '            '20160622 住所分割処理対応 -chg end
            '            .Vari_Addr_kenno = hash_address("ken_no")
            '            .Vari_Addr_sino = hash_address("si_no")
            '            .Vari_Addr_cyo = hash_address("mati")
            '            .Vari_Addr_cyome = hash_address("cyome")
            '            .Vari_Addr_cyomeptn = hash_address("chomeptn")
            '            .Vari_Addr_banti = hash_address("banti")
            '            .Vari_Addr_etc = hash_address("etc")

            '            '固定値
            '            .Vari_Useflg = 1
            '            .Vari_History = DefHistory

            '            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            '            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '            'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
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


                Dim tmp_sql As String = " SELECT jisya_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)


            End Sub

        End Class

    End Class

#End Region

#Region "自社口座情報"

    'V7には自社口座情報が存在しない
    '家賃入金口座マスタに登録されている口座情報から判別
    '判別は紐付けツールで行う
    '紐付け材料は名称、口座名義、カナ等

    Public Class Jisyadata_koza_Repository

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

                Dim model_cvitem As New Njc.Model.Jisyadata_koza_Model          '移行値格納用モデル初期化
                Dim keycol_main As Integer = 15                                  'メインキー列
                Dim keycol_sub As Integer = 17                                   'サブキー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call Me.Get_RelData_KozaSyubetu()

                'Excelファイル初期設定
                '2016.04.26 メインの方へも反映させる修正 -chg sta
                '自社口座は紐付けファイルから取得する
                'rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)
                Dim readfiledirpath As String = RelationDirPath         '紐付けファイル格納先
                Dim readfilename As String = MIDFILE_RELNAME            '紐付けファイル名
                Dim readsheetname As String = "自社口座マスタ"          '紐付けシート名
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, readfiledirpath, readfilename, readsheetname)
                '2016.04.26 メインの方へも反映させる修正 -chg end

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
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
                                Case "自社No"
                                    .Vari_Jisya_no = fldvalue.Trim
                                Case "自社口座No"
                                    .Vari_Jisya_kozano = fldvalue.Trim
                                Case "金融機関No"
                                    .Vari_Kinyu_no = fldvalue.Trim
                                Case "支店No"
                                    .Vari_Kinyu_tenno = fldvalue.Trim
                                Case "口座種別"
                                    .Vari_Koza_syubetu = fldvalue.Trim
                                Case "口座番号"
                                    .Vari_Koza_bango = fldvalue.Trim
                                Case "口座名義"
                                    .Vari_Koza_meigi = fldvalue.Trim
                                    '20160516 FB構築に伴う自社口座情報の修正 -add sta
                                Case "口座名義カナ"
                                    .Vari_Koza_meigikana = fldvalue.Trim
                                Case "ゆうちょ記号1"
                                    .Vari_Yucyokoza_kigo1 = fldvalue.Trim
                                Case "ゆうちょ記号2"
                                    .Vari_Yucyokoza_kigo2 = fldvalue.Trim
                                Case "ゆうちょ口座番号"
                                    .Vari_Yucyokoza_bango = fldvalue.Trim
                                Case "備考(口座情報)"
                                    .Vari_Biko_koza = fldvalue.Trim
                                    '20160516 FB構築に伴う自社口座情報の修正 -add end
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

#Region "自社担当者情報"   '20160608 自社担当者情報構築

    Public Class Profile_logonuser_Repository

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

                Dim model_cvitem As New Njc.Model.Profile_logonuser_Model       '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列

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
                'テーブル名/フィールド名セット
                Dim tblname As String = "profile_logonuser"
                Dim fldnamegrp As String = "logonuser_no,logonuser_name,logonuser_namesjis,logonuser_kana,logongroup_no," & _
                                           "tel1,tel2,mobiletel1,fax1,mailaddress1," & _
                                           "mobilemailaddress1,takkenmenkyo,biko,denylogon,noassign," & _
                                           "password,behaviourdefine_base64,colordefine_base64,n3spreadcolordefine_base64,menudesigndefine_base64," & _
                                           "smtpdefine_base64,history,sendtargetsiteid,windowsuser_name,hysearchdefine_base64," & _
                                           "csvsetting_base64,dontshowfeedbackagain"

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

                    'キーヘッダー名取得
                    Dim fldname_key As String = readtbl.Columns(keycol - 1).ColumnName.Trim

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
                        Dim fldvalue_key As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol - 1)).Trim

                        'ログ出力用
                        Dim str_logkey As String = fldname_key & " = " & fldvalue_key

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "賃貸革命ログオンユーザNo"
                                    .Vari_Logonuser_no = fldvalue
                                Case "賃貸革命ログオンユーザ名"
                                    .Vari_Logonuser_name = fldvalue
                                Case "賃貸革命ログオンユーザ名(SJIS)"
                                    .Vari_Logonuser_namesjis = fldvalue
                                Case "賃貸革命ログオンユーザ名カナ"
                                    .Vari_Logonuser_kana = fldvalue
                                Case "賃貸革命ログオングループNo"
                                    '20160928 自社担当者移行処理の修正 -chg sta
                                    '.Vari_Logongroup_no = fldvalue
                                    .Vari_Logongroup_no = IIf(fldvalue = "", "99001", fldvalue)
                                    '20160928 自社担当者移行処理の修正 -chg end
                                Case "TEL1"
                                    .Vari_Tel1 = fldvalue
                                Case "TEL2"
                                    .Vari_Tel2 = fldvalue
                                Case "携帯電話1"
                                    .Vari_Mobiletel1 = fldvalue
                                Case "FAX1"
                                    .Vari_Fax1 = fldvalue
                                Case "メールアドレス1"
                                    .Vari_Mailaddress1 = fldvalue
                                Case "携帯メールアドレス1"
                                    .Vari_Mobilemailaddress1 = fldvalue
                                Case "宅建免許番号"
                                    .Vari_Takkenmenkyo = fldvalue
                                Case "備考"
                                    .Vari_Biko = fldvalue
                                Case "担当者ログオン禁止フラグ"
                                    .Vari_Denylogon = fldvalue
                                Case "使用しないフラグ"
                                    .Vari_Noassign = fldvalue
                                Case "ログオンパスワード"
                                    .Vari_Password = fldvalue
                                Case "windowsログオンユーザー名"
                                    .Vari_Windowsuser_name = fldvalue
                                Case "意見画面表示フラグ"
                                    .Vari_Dontshowfeedbackagain = fldvalue
                            End Select

                        Next

                        '固定値
                        .Vari_Behaviourdefine_base64 = "H4sIAAAAAAAEAI1WW2/aMBR+n7T/gHgvIbTd1IpSQRgr4tKutNL6VLnJIXhx7MyXptmvnzFJSUxI+wbnu5zvOMdR+tdvMWm9AheY0au22+m2rwdfv/RHsEGvmCk+hjWm0NIsKi7fBL5qb6RMLh0nTdNOetphPHR63a7r/F7MV/4GYnSCqZCI+tB+VwUfq9q6a6vV9xCdMF+JB46owFKHGmU/qAQ+g2wguYK+00g5avKYjFlKPcUF+8DLZlYsR4wHwD1GGPc2iIZgGR3iRm5jQx6+DE7c3tlFt/fd7Tu1eNH4l8JSjyh8lEAlug0Ugl1nj1HJGTF+e00NVsimIWUcpjRR8h7+Kswh2OvqwEJ4r/SxTc4r2Uq1gqZ/Dzkg3ZSgRJROzgaMYIEwQUHAQQgPCPEI9qMHtq0CX0nE5WCNiNAGHxNzP6F3ZM5YpJJtKtc922qrxRLzNgG6450XvKK0G0gfIbzJGyC5XU/PUq0ZnjWc4fZO+05N3fCHSrIFYIHwKsVS3w1eDFqDGMVUTED/HRIyRhIVbKuaPwRCIJgwHh/v0sg5klCzl0xPII5H3VPyzHp5/egnK93d9+CHkNHcISFSfUdusJCMZy2nWp7rZ7RgAV5jCJ7ioID1ItAFUDWNUQgjJODbmQ2ZLVmiWK8kS3wkIdT2zy+R2S2LURHe43AjVziAiaL+9u0xcPeaQzB/FXB/e//WmMe31JARVR5hAtMwvxPNnJ1NxOKFEnLG/qEVEPDlXB+LPdn4boJePaJPyIRf/vE7+dUXneWpQbulyBbdshpFDU6jqGpUJh/4dN3nGaIczwBLhOuDz7KGbrOs2q1MPvD5XDfNmopHfREzvTz5cziOW/InJDFtyGvwamRLUmf4meAeaWjrkWrPMtnyWWUswu+o6dN3rK+AwX9Db0j/LggAAA=="
                        .Vari_Colordefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E9zzs8rKcrPcc7PyS9ySU3LzEtVAKrPK7aqKM60VcooKSmw0tcvLy/XKzfWyy9K1zcyMDDUj/D1CU7OSM1N1M3MKy5JzEtOVYLrSiGsS0lB3w4AaXGCjpcAAAA="
                        .Vari_N3spreadcolordefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zMw4uKEpNTAnNS0kt8snMS3XOz8kvcklNAzIVgDrziq0qijNtlTJKSgqs9PXLy8v1yo318ovS9Y0MDAz1I3x9gpMzUnMTdTPziksS85JTleC6UgjrUlLQtwMAoBcTOKEAAAA="
                        .Vari_Menudesigndefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zTc0rdUktzkzPc0lNy8xLVQCqziu2qijOtFXKKCkpsNLXLy8v1ys31ssvStc3MjAw1I/w9QlOzkjNTdTNzCsuScxLTlWC60ohrEtJQd8OAJAcjj6VAAAA"
                        .Vari_Smtpdefine_base64 = "H4sIAAAAAAAEAIWRz2vCMBTH7wP/h5C7fa1jF4mRqRdhboU62DXU17WgeSXJFv3vTSPtSi+75fvj8XmPiPX1cma/aGxDesWzJOVrOXsSxeGY77BqNLJQ0HZ5tc2K1861SwDvfeKfEzLfsEjTDL4Ob0VZ40XNG22d0iXyYer0/xQPQMYiskATdmHw5+RknFy8CBhEjLakNZbueGtR7ki7rTIoYOzG2uuPq6N6p+4pYDAGwKedAJW1nsyp9/Y2p3aDFRnsYlmpsw2oqR27+cfkgmDEnbMsFdCLPhmTu2wMftz7+AF5B/390ZmlAQAA"
                        .Vari_Sendtargetsiteid = ""
                        .Vari_Hysearchdefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zqAxOTSxKzggAKsjPS8xxzs9LySwBKlUAassrtqoozrRVyigpKbDS1y8vL9crN9bLL0rXNzIwMNSP8PUJTs5IzU3UzcwrLknMS05VgutKIaxLSUHfDgDu3pVrngAAAA=="
                        .Vari_Csvsetting_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E9zDg4LTi0pycxLVwCqyyu2qijOtFXKKCkpsNLXLy8v1ys31ssvStc3MjAw1I/w9QlOzkjNTdTNzCsuScxLTlWC60ohrEtJQd8OAIz7H0GPAAAA"
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
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

            ' ''' <summary>
            ' ''' 【中間ファイル→変数】
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="syorikomok"></param>
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

            '    Dim model_cvitem As New Njc.Model.Profile_logonuser_Model       '移行値格納用モデル初期化
            '    Dim keycol As Integer = 1                                       'キー列

            '    '************************
            '    '作業準備
            '    '************************

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If Not rtn Then
            '        Return rtn
            '    End If

            '    'テーブル名/フィールド名セット
            '    Dim tblname As String = "profile_logonuser"
            '    Dim fldnamegrp As String = "logonuser_no,logonuser_name,logonuser_namesjis,logonuser_kana,logongroup_no," & _
            '                               "tel1,tel2,mobiletel1,fax1,mailaddress1," & _
            '                               "mobilemailaddress1,takkenmenkyo,biko,denylogon,noassign," & _
            '                               "password,behaviourdefine_base64,colordefine_base64,n3spreadcolordefine_base64,menudesigndefine_base64," & _
            '                               "smtpdefine_base64,history,sendtargetsiteid,windowsuser_name,hysearchdefine_base64," & _
            '                               "csvsetting_base64,dontshowfeedbackagain"

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

            '        '---------------
            '        'ヘッダー処理
            '        '---------------
            '        'ヘッダー行取得
            '        Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '        'キーヘッダー名取得
            '        Dim fldname_key As String = headervalue(startrow - 1, keycol)

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
            '            Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

            '            'キー値取得
            '            Dim fldvalue_key As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))

            '            'ログ出力用
            '            Dim str_logkey As String = fldname_key & " = " & fldvalue_key

            '            '移行値取得
            '            For cntjj = 1 To columncnt

            '                Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            '                Dim fldvalue As String = ""
            '                If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            '                End If

            '                Select Case fldname
            '                    Case "賃貸革命ログオンユーザNo"
            '                        .Vari_Logonuser_no = fldvalue
            '                    Case "賃貸革命ログオンユーザ名"
            '                        .Vari_Logonuser_name = fldvalue
            '                    Case "賃貸革命ログオンユーザ名(SJIS)"
            '                        .Vari_Logonuser_namesjis = fldvalue
            '                    Case "賃貸革命ログオンユーザ名カナ"
            '                        .Vari_Logonuser_kana = fldvalue
            '                    Case "賃貸革命ログオングループNo"
            '                        '20160928 自社担当者移行処理の修正 -chg sta
            '                        '.Vari_Logongroup_no = fldvalue
            '                        .Vari_Logongroup_no = IIf(fldvalue = "", "99001", fldvalue)
            '                        '20160928 自社担当者移行処理の修正 -chg end
            '                    Case "TEL1"
            '                        .Vari_Tel1 = fldvalue
            '                    Case "TEL2"
            '                        .Vari_Tel2 = fldvalue
            '                    Case "携帯電話1"
            '                        .Vari_Mobiletel1 = fldvalue
            '                    Case "FAX1"
            '                        .Vari_Fax1 = fldvalue
            '                    Case "メールアドレス1"
            '                        .Vari_Mailaddress1 = fldvalue
            '                    Case "携帯メールアドレス1"
            '                        .Vari_Mobilemailaddress1 = fldvalue
            '                    Case "宅建免許番号"
            '                        .Vari_Takkenmenkyo = fldvalue
            '                    Case "備考"
            '                        .Vari_Biko = fldvalue
            '                    Case "担当者ログオン禁止フラグ"
            '                        .Vari_Denylogon = fldvalue
            '                    Case "使用しないフラグ"
            '                        .Vari_Noassign = fldvalue
            '                    Case "ログオンパスワード"
            '                        .Vari_Password = fldvalue
            '                    Case "windowsログオンユーザー名"
            '                        .Vari_Windowsuser_name = fldvalue
            '                    Case "意見画面表示フラグ"
            '                        .Vari_Dontshowfeedbackagain = fldvalue
            '                End Select

            '            Next

            '            '固定値
            '            .Vari_Behaviourdefine_base64 = "H4sIAAAAAAAEAI1WW2/aMBR+n7T/gHgvIbTd1IpSQRgr4tKutNL6VLnJIXhx7MyXptmvnzFJSUxI+wbnu5zvOMdR+tdvMWm9AheY0au22+m2rwdfv/RHsEGvmCk+hjWm0NIsKi7fBL5qb6RMLh0nTdNOetphPHR63a7r/F7MV/4GYnSCqZCI+tB+VwUfq9q6a6vV9xCdMF+JB46owFKHGmU/qAQ+g2wguYK+00g5avKYjFlKPcUF+8DLZlYsR4wHwD1GGPc2iIZgGR3iRm5jQx6+DE7c3tlFt/fd7Tu1eNH4l8JSjyh8lEAlug0Ugl1nj1HJGTF+e00NVsimIWUcpjRR8h7+Kswh2OvqwEJ4r/SxTc4r2Uq1gqZ/Dzkg3ZSgRJROzgaMYIEwQUHAQQgPCPEI9qMHtq0CX0nE5WCNiNAGHxNzP6F3ZM5YpJJtKtc922qrxRLzNgG6450XvKK0G0gfIbzJGyC5XU/PUq0ZnjWc4fZO+05N3fCHSrIFYIHwKsVS3w1eDFqDGMVUTED/HRIyRhIVbKuaPwRCIJgwHh/v0sg5klCzl0xPII5H3VPyzHp5/egnK93d9+CHkNHcISFSfUdusJCMZy2nWp7rZ7RgAV5jCJ7ioID1ItAFUDWNUQgjJODbmQ2ZLVmiWK8kS3wkIdT2zy+R2S2LURHe43AjVziAiaL+9u0xcPeaQzB/FXB/e//WmMe31JARVR5hAtMwvxPNnJ1NxOKFEnLG/qEVEPDlXB+LPdn4boJePaJPyIRf/vE7+dUXneWpQbulyBbdshpFDU6jqGpUJh/4dN3nGaIczwBLhOuDz7KGbrOs2q1MPvD5XDfNmopHfREzvTz5cziOW/InJDFtyGvwamRLUmf4meAeaWjrkWrPMtnyWWUswu+o6dN3rK+AwX9Db0j/LggAAA=="
            '            .Vari_Colordefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E9zzs8rKcrPcc7PyS9ySU3LzEtVAKrPK7aqKM60VcooKSmw0tcvLy/XKzfWyy9K1zcyMDDUj/D1CU7OSM1N1M3MKy5JzEtOVYLrSiGsS0lB3w4AaXGCjpcAAAA="
            '            .Vari_N3spreadcolordefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zMw4uKEpNTAnNS0kt8snMS3XOz8kvcklNAzIVgDrziq0qijNtlTJKSgqs9PXLy8v1yo318ovS9Y0MDAz1I3x9gpMzUnMTdTPziksS85JTleC6UgjrUlLQtwMAoBcTOKEAAAA="
            '            .Vari_Menudesigndefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zTc0rdUktzkzPc0lNy8xLVQCqziu2qijOtFXKKCkpsNLXLy8v1ys31ssvStc3MjAw1I/w9QlOzkjNTdTNzCsuScxLTlWC60ohrEtJQd8OAJAcjj6VAAAA"
            '            .Vari_Smtpdefine_base64 = "H4sIAAAAAAAEAIWRz2vCMBTH7wP/h5C7fa1jF4mRqRdhboU62DXU17WgeSXJFv3vTSPtSi+75fvj8XmPiPX1cma/aGxDesWzJOVrOXsSxeGY77BqNLJQ0HZ5tc2K1861SwDvfeKfEzLfsEjTDL4Ob0VZ40XNG22d0iXyYer0/xQPQMYiskATdmHw5+RknFy8CBhEjLakNZbueGtR7ki7rTIoYOzG2uuPq6N6p+4pYDAGwKedAJW1nsyp9/Y2p3aDFRnsYlmpsw2oqR27+cfkgmDEnbMsFdCLPhmTu2wMftz7+AF5B/390ZmlAQAA"
            '            .Vari_Sendtargetsiteid = ""
            '            .Vari_Hysearchdefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zqAxOTSxKzggAKsjPS8xxzs9LySwBKlUAassrtqoozrRVyigpKbDS1y8vL9crN9bLL0rXNzIwMNSP8PUJTs5IzU3UzcwrLknMS05VgutKIaxLSUHfDgDu3pVrngAAAA=="
            '            .Vari_Csvsetting_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E9zDg4LTi0pycxLVwCqyyu2qijOtFXKKCkpsNLXLy8v1ys31ssvStc3MjAw1I/w9QlOzkjNTdTNzCsuScxLTlWC60ohrEtJQd8OAIz7H0GPAAAA"
            '            .Vari_History = DefHistory

            '            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            '            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '            'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
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


                Dim tmp_sql As String = " SELECT logonuser_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)


            End Sub

        End Class

    End Class

#End Region

#Region "自社メモ情報"

    Public Class Jisyadata_memo_Repository

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

                Dim model_cvitem As New Njc.Model.Jisyadata_memo_Model          '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列
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
                'テーブル名/フィールド名セット
                Dim tblname_base As String = "jisyadata"
                Dim tblname As String = "jisyadata_memo"
                Dim fldnamegrp As String = "jisya_no,memo_no,memo,history"

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

                '移行対象件数取得用
                Dim cvtaisyocnt As Integer = 0

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_keymain As String = readtbl.Columns(keycol - 1).ColumnName.Trim

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
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol - 1)).Trim
                        .Vari_Jisya_no = tmp_keymain

                        '備考カウント初期化
                        Dim cvitemcnt As Integer = 0
                        Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

                        'データ有無チェック
                        Dim tmp_fldvalueumuchk As String = ""
                        For cntjj = 1 To readtbl.Columns.Count - 1
                            tmp_fldvalueumuchk = tmp_fldvalueumuchk & Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                        Next
                        If tmp_fldvalueumuchk = "" Then
                            cvflg = False
                        End If

                        If cvflg Then

                            '移行対象件数カウント
                            cvtaisyocnt = cvtaisyocnt + 1

                            Dim hash_keylog As New Hashtable
                            Dim log_keyout As String = fldname_keymain & " = " & tmp_keymain
                            '親データ有無チェック
                            If list_basekeydata.Contains(tmp_keymain) = False Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "jisyadata_memo-jisya_no"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(tmp_keymain) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "jisyadata_memo-jisya_no"
                                Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            If keychkflg Then

                                '移行値取得
                                For cntjj = 1 To readtbl.Columns.Count - 1

                                    'サブキー取得
                                    Dim tmp_keysub As String = (cntjj).ToString

                                    '全キー取得
                                    Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

                                    'ログ出力用データ格納(サブキーフィールド)
                                    Dim fldname_keysub As String = readtbl.Columns(cntjj).ColumnName.Trim
                                    Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub

                                    '登録値取得
                                    Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                                    .Vari_Memo = fldvalue

                                    '固定値
                                    .Vari_Memo_no = (cntjj).ToString
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

                                            '挿入処理
                                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                            If normalflg Then
                                                If list_chkduplicate.Contains(tmp_keymain) = False Then
                                                    list_chkduplicate.Add(tmp_keymain)
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

            ' ''' <summary>
            ' ''' 【中間ファイル→変数】
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="syorikomok"></param>
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

            '    Dim model_cvitem As New Njc.Model.Jisyadata_memo_Model          '移行値格納用モデル初期化
            '    Dim keycol As Integer = 1                                       'キー列
            '    Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

            '    '************************
            '    '作業準備
            '    '************************

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If Not rtn Then
            '        Return rtn
            '    End If

            '    'テーブル名/フィールド名セット
            '    Dim tblname_base As String = "jisyadata"
            '    Dim tblname As String = "jisyadata_memo"
            '    Dim fldnamegrp As String = "jisya_no,memo_no,memo,history"

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
            '        Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '        'キーヘッダー名取得
            '        Dim fldname_keymain As String = headervalue(startrow - 1, keycol)

            '        For cntii = 1 To rowcnt

            '            '中断処理
            '            Application.DoEvents()
            '            If CancelFlg Then
            '                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            '                Return rtn
            '            End If

            '            '行取得
            '            Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

            '            'キー値取得
            '            Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))
            '            .Vari_Jisya_no = tmp_keymain

            '            '備考カウント初期化
            '            Dim cvitemcnt As Integer = 0

            '            '移行値取得
            '            For cntjj = 2 To columncnt

            '                'サブキー取得
            '                Dim tmp_keysub As String = (cntjj - 1).ToString

            '                '全キー取得
            '                Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

            '                'ログ出力用データ格納(サブキーフィールド)
            '                Dim fldname_keysub As String = headervalue(startrow - 1, cntjj)
            '                Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub

            '                Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            '                Dim fldvalue As String = ""
            '                If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            '                End If
            '                .Vari_Memo = fldvalue

            '                '固定値
            '                .Vari_Memo_no = (cntjj - 1).ToString
            '                .Vari_History = DefHistory

            '                'メモにデータが存在する場合に書込処理を行う
            '                If .Vari_Memo <> "" Then

            '                    '20160928 メモ関連の移行件数表示修正 -add
            '                    tmp_cvrowcnt = tmp_cvrowcnt + 1

            '                    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            '                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '                    'データチェック
            '                    Dim skipflg As Boolean = False
            '                    Dim hash_cvitem As New Hashtable
            '                    Dim hash_log As New Hashtable
            '                    skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

            '                    '書込処理
            '                    If Not skipflg Then

            '                        '挿入処理
            '                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

            '                        '挿入処理が正常終了したレコードのキーを重複チェック用に格納
            '                        If normalflg Then
            '                            list_chkduplicate.Add(fldvalue_key)
            '                            cvitemcnt = cvitemcnt + 1
            '                        End If

            '                    End If

            '                    'ログ出力メッセージ整形
            '                    Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

            '                    'ログ出力
            '                    If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
            '                        Dim tmp_cnt As Integer = 0
            '                        '挿入
            '                        For Each logvalue In sortlist_log
            '                            Dim tmp_sql_insert As String = logvalue.Value
            '                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
            '                        Next
            '                        '初期化
            '                        tmp_logcnt = 0
            '                        sortlist_log.Clear()
            '                    End If

            '                End If

            '            Next

            '            '--------------------------------------------------------------------
            '            '移行した備考が1データ以上ある場合移行したレコードの数を更新する
            '            '--------------------------------------------------------------------
            '            '20160928 メモ関連の移行件数表示修正 -chg sta
            '            'tmp_cvcnt = tmp_cvcnt + 1
            '            If cvitemcnt > 0 Then
            '                tmp_cvcnt = tmp_cvcnt + 1
            '            End If
            '            '20160928 メモ関連の移行件数表示修正 -chg end
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
            '    '20160928 メモ関連の移行件数表示修正 -chg sta
            '    'midrowcnt = rowcnt
            '    midrowcnt = tmp_cvrowcnt
            '    '20160928 メモ関連の移行件数表示修正 -chg end
            '    '移行件数を取得
            '    cvrowcnt = tmp_cvcnt

            '    '調整件数を取得
            '    conditioncnt = tmp_condcnt

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            '    '返却
            '    Return rtn

            'End Function

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

                Dim tmp_sql As String = " SELECT jisya_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "振込依頼人情報"

    Public Class M_fb_sgfirai_Repository

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

                Dim model_cvitem As New Njc.Model.M_fb_sgfirai_Model            '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列

                '************************
                '作業準備
                '************************

                '2016.04.26 メインの方へも反映させる修正 -add sta
                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Jisyakoza()
                Call SetRelItemToObject.Set_RelData_FBInfo()        '20160609 紐付設定値取得に伴う修正 -add
                '2016.04.26 メインの方へも反映させる修正 -add end

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
                Dim tblname As String = "m_fb_sgfirai"
                '20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg sta
                'Dim fldnamegrp As String = "sgfirai_no,sgfirai_name,jisya_no,jisya_kozano,sgfirainin_code," & _
                '                           "sgfirainin_kana,biko_kihon,sgfzenfmt_no,file_sosin,crlf," & _
                '                           "changewo_flg,biko_data,keisandefault_flg,futancyousei_flg,biko_tesu," & _
                '                           "history,sgfdata_createflg,useflg"
                Dim fldnamegrp As String = "sgfirai_no,sgfirai_name,jisya_no,jisya_kozano,sgfirainin_code," & _
                                           "sgfirainin_kana,biko,sgfzenfmt_no,file_sosin,changewo_flg," & _
                                           "futancyousei_flg,history,sgfdata_createflg,useflg"
                '20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg end

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

                    'キーヘッダー名取得
                    Dim fldname_key As String = readtbl.Columns(keycol - 1).ColumnName.Trim

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
                        Dim tmp_kinyuno As String = ""
                        Dim tmp_kinyutenno As String = ""
                        Dim tmp_kozasyu As String = ""      '20160609 紐付取得用に追加 -add
                        Dim tmp_kozabango As String = ""
                        Dim tmp_jisyano As String = ""      '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
                        Dim tmp_jisyakozano As String = ""  '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add

                        'キー値取得
                        Dim fldvalue_key As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol - 1)).Trim

                        'ログ出力用
                        Dim str_logkey As String = fldname_key & " = " & fldvalue_key

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "振込依頼人No"
                                    .Vari_Sgfirai_no = fldvalue
                                    .Vari_Sgfzenfmt_no = fldvalue   '20160609 紐付設定値取得に伴う修正 -add
                                Case "振込依頼人名"
                                    .Vari_Sgfirai_name = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -chg sta
                                    'Case "自社支店No."
                                    '    .Vari_Jisya_no = fldvalue.Trim
                                    'Case "自社口座No"
                                    '    .Vari_Jisya_kozano = fldvalue.Trim
                                Case "口座種別" '20160609 紐付取得用に追加 -add
                                    tmp_kozasyu = fldvalue
                                Case "金融機関No"
                                    tmp_kinyuno = fldvalue
                                Case "金融機関支店No"
                                    tmp_kinyutenno = fldvalue
                                Case "口座番号"
                                    tmp_kozabango = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -chg end
                                Case "振込依頼人コード"
                                    .Vari_Sgfirainin_code = fldvalue
                                Case "振込依頼人カナ"
                                    .Vari_Sgfirainin_kana = fldvalue
                                Case "備考(基本情報)"
                                    .Vari_Biko = fldvalue             '20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg (Vari_Biko_kihon → Vari_Biko)
                                    '20160609 紐付設定値取得に伴う修正 -del sta
                                    'Case "総合振込全銀フォーマットNo"
                                    '    .Vari_Sgfzenfmt_no = fldvalue
                                    '20160609 紐付設定値取得に伴う修正 -del end
                                Case "送信ファイルパス"
                                    .Vari_File_sosin = fldvalue
                                    '20160519 EXEUpdateに伴う修正 振込依頼人情報 -del sta
                                    'Case "改行有無(0:改行しない 1:改行する)"
                                    '    .Vari_Crlf = fldvalue
                                    '20160519 EXEUpdateに伴う修正 振込依頼人情報 -del end
                                Case "ヲ変換フラグ(0:変換しない 1:変換する)"
                                    .Vari_Changewo_flg = fldvalue
                                    '20160519 EXEUpdateに伴う修正 振込依頼人情報 -del sta
                                    'Case "備考(データ)"
                                    '    .Vari_Biko_data = fldvalue
                                    'Case "計算時の初期値フラグ(0:初期値としない 1:初期値とする)"
                                    '    .Vari_Keisandefault_flg = fldvalue
                                    '20160519 EXEUpdateに伴う修正 振込依頼人情報 -del end
                                Case "負担調整フラグ(0:調整しない 1:調整する)"
                                    .Vari_Futancyousei_flg = fldvalue
                                    '20160519 EXEUpdateに伴う修正 振込依頼人情報 -del sta
                                    'Case "備考(手数料)"
                                    '    .Vari_Biko_tesu = fldvalue
                                    '20160519 EXEUpdateに伴う修正 振込依頼人情報 -del end
                                Case "総合振込データ生成フラグ"
                                    .Vari_Sgfdata_createflg = fldvalue
                                    '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
                                Case "自社支店No"  '20160929 汎用→既存コピー処理改善対応 自社支店No. → 自社支店No -chg
                                    tmp_jisyano = fldvalue
                                Case "自社口座No"
                                    tmp_jisyakozano = fldvalue
                                    '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end
                            End Select

                        Next
                        '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg sta
                        ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                        ' ''20160609 紐付取得用に追加 -chg sta
                        '' ''2016.04.26 メインの方へも反映させる修正 -add sta
                        '' ''自社No/自社口座No取得のため口座情報を格納する
                        ' ''.Vari_Jisya_no = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                        ' ''.Vari_Jisya_kozano = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                        '' ''2016.04.26 メインの方へも反映させる修正 -add end
                        ''Dim tmp_jisyainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
                        ''.Vari_Jisya_no = tmp_jisyainfo
                        ''.Vari_Jisya_kozano = tmp_jisyainfo
                        ' ''20160609 紐付取得用に追加 -chg end
                        'Dim tmp_kozamoto As String = "振込依頼人マスタ"
                        'Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Sgfirai_no
                        '.Vari_Jisya_no = tmp_jisyainfo
                        '.Vari_Jisya_kozano = tmp_jisyainfo
                        ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                        Select Case CNVNO
                            Case ConvertTypes._既存ユーザ用
                                Dim tmp_kozamoto As String = "振込依頼人マスタ"
                                Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Sgfirai_no
                                .Vari_Jisya_no = tmp_jisyainfo
                                .Vari_Jisya_kozano = tmp_jisyainfo
                            Case ConvertTypes._汎用
                                .Vari_Sgfzenfmt_no = ""
                                '20160928 自社口座取得方法の修正 -chg sta
                                '.Vari_Jisya_no = tmp_jisyano
                                '.Vari_Jisya_kozano = tmp_jisyakozano
                                .Vari_Jisya_no = tmp_jisyano & "-" & tmp_jisyakozano
                                .Vari_Jisya_kozano = tmp_jisyano & "-" & tmp_jisyakozano
                                '20160928 自社口座取得方法の修正 -chg end
                        End Select
                        '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg end
                        '固定値
                        .Vari_Useflg = 1
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20160928 自社口座取得方法の修正 -add sta
                        tmp_hash("jisya_no") = tmp_jisyano
                        tmp_hash("jisya_kozano") = tmp_jisyakozano
                        '20160928 自社口座取得方法の修正 -add end

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

            ' ''' <summary>
            ' ''' 【中間ファイル→変数】
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="syorikomok"></param>
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

            '    Dim model_cvitem As New Njc.Model.M_fb_sgfiraitesu_Model        '移行値格納用モデル初期化
            '    Dim keycol_main As Integer = 1                                  'メインキー列
            '    Dim keycol_sub As Integer = 2                                   'サブキー列

            '    '************************
            '    '作業準備
            '    '************************

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If Not rtn Then
            '        Return rtn
            '    End If

            '    'テーブル名/フィールド名セット
            '    Dim tblname_base As String = "m_fb_sgfirai"
            '    Dim tblname As String = "m_fb_sgfiraitesu"
            '    Dim fldnamegrp As String = "sgfirai_no,rec_no,from_gak,to_gak,doukoudouten_densingak," & _
            '                               "doukoudouten_bunsyogak,doukoutaten_densingak,doukoutaten_bunsyogak,takou_densingak,takou_bunsyogak"

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
            '        Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '        'キーヘッダー名取得
            '        Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
            '        Dim fldname_keysub As String = headervalue(startrow - 1, keycol_sub)

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
            '            Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

            '            'キー値取得
            '            Dim fldvalue_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
            '            Dim fldvalue_keysub As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub))
            '            Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

            '            'ログ出力用
            '            Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

            '            '移行値取得
            '            For cntjj = 1 To columncnt

            '                Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            '                Dim fldvalue As String = ""
            '                If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            '                End If

            '                Select Case fldname
            '                    Case "振込依頼人No"
            '                        .Vari_Sgfirai_no = fldvalue
            '                    Case "レコードNo"
            '                        .Vari_Rec_no = fldvalue
            '                    Case "金額From"
            '                        .Vari_From_gak = fldvalue
            '                    Case "金額To"
            '                        .Vari_To_gak = fldvalue
            '                    Case "同行同支店手数料(電信扱)"
            '                        .Vari_Doukoudouten_densingak = fldvalue
            '                    Case "同行同支店手数料(文書扱)"
            '                        .Vari_Doukoudouten_bunsyogak = fldvalue
            '                    Case "同行他支店手数料(電信扱)"
            '                        .Vari_Doukoutaten_densingak = fldvalue
            '                    Case "同行他支店手数料(文書扱)"
            '                        .Vari_Doukoutaten_bunsyogak = fldvalue
            '                    Case "他行手数料(電信扱)"
            '                        .Vari_Takou_densingak = fldvalue
            '                    Case "他行手数料(文書扱)"
            '                        .Vari_Takou_bunsyogak = fldvalue
            '                End Select

            '            Next

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


                Dim tmp_sql As String = " SELECT jisya_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)


            End Sub

        End Class

    End Class

#End Region

#Region "振込手数料情報"       '20160517 振込手数料情報の新規作成

    Public Class M_fb_sgfiraitesu_Repository

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

                Dim model_cvitem As New Njc.Model.M_fb_sgfiraitesu_Model        '移行値格納用モデル初期化
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
                Dim tblname_base As String = "m_fb_sgfirai"
                Dim tblname As String = "m_fb_sgfiraitesu"
                Dim fldnamegrp As String = "sgfirai_no,rec_no,from_gak,to_gak,doukoudouten_densingak," & _
                                           "doukoudouten_bunsyogak,doukoutaten_densingak,doukoutaten_bunsyogak,takou_densingak,takou_bunsyogak"

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
                                Case "振込依頼人No"
                                    .Vari_Sgfirai_no = fldvalue
                                Case "レコードNo"
                                    .Vari_Rec_no = fldvalue
                                Case "金額From"
                                    .Vari_From_gak = fldvalue
                                Case "金額To"
                                    .Vari_To_gak = fldvalue
                                Case "同行同支店手数料(電信扱)"
                                    .Vari_Doukoudouten_densingak = fldvalue
                                Case "同行同支店手数料(文書扱)"
                                    .Vari_Doukoudouten_bunsyogak = fldvalue
                                Case "同行他支店手数料(電信扱)"
                                    .Vari_Doukoutaten_densingak = fldvalue
                                Case "同行他支店手数料(文書扱)"
                                    .Vari_Doukoutaten_bunsyogak = fldvalue
                                Case "他行手数料(電信扱)"
                                    .Vari_Takou_densingak = fldvalue
                                Case "他行手数料(文書扱)"
                                    .Vari_Takou_bunsyogak = fldvalue
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
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT sgfirai_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,sgfirai_no) + '-' + CONVERT(varchar,rec_no) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "口座振替情報"

    Public Class M_fb_fkaejyoho_Repository

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

                Dim model_cvitem As New Njc.Model.M_fb_fkaejyoho_Model          '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列

                '************************
                '作業準備
                '************************

                '2016.04.26 メインの方へも反映させる修正 -add sta
                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Jisyakoza()
                Call SetRelItemToObject.Set_RelData_FBInfo()        '20160609 紐付設定値取得に伴う修正 -add
                '2016.04.26 メインの方へも反映させる修正 -add end

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
                Dim tblname As String = "m_fb_fkaejyoho"
                '20161012 革命10アップデートに伴う修正 「saifkae_use」を追加 -add
                '20160519 EXEUpdateに伴う修正 口座振替情報 -chg sta
                'Dim fldnamegrp As String = "fkae_no,fkae_name,fkae_kana,servicetype,fkae_fb_fkomiraino," & _
                '                           "fkae_fb_fkomiraikana,kamei_no,jlease_nyukinkbn,jisya_no,fkae_fb_nkinukekozano," & _
                '                           "hikiotosibi,tesu_gak,biko,fkae_fb_orgfmtno,datasort," & _
                '                           "keiyakusyano_syuturyoku,crlf,wo_mojihenkan,file_sosin,file_jyusin," & _
                '                           "fdsakusei,seikyu_tani,seikyu_taino,seikyu_tukitani,tesu_nyukinkanri," & _
                '                           "tesu_kurikosi,syogorule,yutyobango,history,useflg"
                Dim fldnamegrp As String = "fkae_no,fkae_name,fkae_kana,fkae_fb_fkomiraino,fkae_fb_fkomiraikana," & _
                                           "kamei_no,jlease_nyukinkbn,jisya_no,fkae_fb_nkinukekozano,hikiotosibi," & _
                                           "tesu_gak,biko,fkae_fb_orgfmtno,datasort,keiyakusyano_syuturyoku," & _
                                           "crlf,wo_mojihenkan,file_sosin,file_jyusin,seikyu_tani," & _
                                           "seikyu_taino,seikyu_tukitani,tesu_nyukinkanri,tesu_kurikosi,syogorule," & _
                                           "yutyobango,history,useflg,multi_flg,yutyo_appendcode," & _
                                           "yutyo_usefkomcode,resultfile_notuse,saifkae_use"
                '20160519 EXEUpdateに伴う修正 口座振替情報 -chg end

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

                    'キーヘッダー名取得
                    Dim fldname_key As String = readtbl.Columns(keycol - 1).ColumnName.Trim

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
                        Dim tmp_kinyuno As String = ""
                        Dim tmp_kinyutenno As String = ""
                        Dim tmp_kozasyu As String = ""      '20160609 紐付取得用に追加 -add
                        Dim tmp_kozabango As String = ""
                        Dim tmp_jisyano As String = ""      '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
                        Dim tmp_jisyakozano As String = ""  '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add

                        'キー値取得
                        Dim fldvalue_key As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol - 1)).Trim

                        'ログ出力用
                        Dim str_logkey As String = fldname_key & " = " & fldvalue_key

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "振替情報No"                          '20160929 汎用→既存コピー処理改善対応 振替情報No. → 振替情報No -chg
                                    .Vari_Fkae_no = fldvalue
                                    .Vari_Fkae_fb_orgfmtno = fldvalue       '20160609 紐付設定値取得に伴う修正 -add
                                Case "振替情報名称"
                                    .Vari_Fkae_name = fldvalue
                                Case "振替情報カナ名称"
                                    .Vari_Fkae_kana = fldvalue
                                    '20160519 EXEUpdateに伴う修正 口座振替情報 -del sta
                                    'Case "サービスタイプ　1:オリコ　2:ジャックス　3:アプラス"
                                    '    .Vari_Servicetype = fldvalue
                                    '20160519 EXEUpdateに伴う修正 口座振替情報 -del end
                                Case "振込依頼人"
                                    .Vari_Fkae_fb_fkomiraino = fldvalue
                                Case "振込依頼人カナ名"
                                    .Vari_Fkae_fb_fkomiraikana = fldvalue
                                Case "加盟店No 依頼人番号と共用"          '20160929 汎用→既存コピー処理改善対応 加盟店No.　依頼人番号と共用 → 加盟店No 依頼人番号と共用 -chg
                                    .Vari_Kamei_no = fldvalue
                                Case "ジェイリース　入金区分　1:変更なし・・・・"
                                    .Vari_Jlease_nyukinkbn = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -chg sta
                                    'Case "自社支店No"
                                    '    .Vari_Jisya_no = fldvalue
                                    'Case "振替先口座No."
                                    '    .Vari_Fkae_fb_nkinukekozano = fldvalue
                                Case "金融機関No"
                                    tmp_kinyuno = fldvalue
                                Case "金融機関支店No"
                                    tmp_kinyutenno = fldvalue
                                Case "口座種別" '20160609 紐付取得用に追加 -add
                                    tmp_kozasyu = fldvalue
                                Case "口座番号"
                                    tmp_kozabango = fldvalue
                                    '2016.04.26 メインの方へも反映させる修正 -chg end
                                Case "引落日"
                                    .Vari_Hikiotosibi = fldvalue
                                Case "手数料　請求額"
                                    .Vari_Tesu_gak = fldvalue
                                Case "備考"
                                    .Vari_Biko = fldvalue
                                    '20160609 紐付設定値取得に伴う修正 -del sta
                                    'Case "オリジナルフォーマットNo"
                                    '    .Vari_Fkae_fb_orgfmtno = fldvalue
                                    '20160609 紐付設定値取得に伴う修正 -del end
                                Case "データ並び替え　0:物件番号と部屋番号　1:契約者番号"
                                    .Vari_Datasort = fldvalue
                                Case "契約者Noを契約者番号に出力する  0:出力しない　1:出力する"    '20160929 汎用→既存コピー処理改善対応 契約者No.を契約者番号に出力する  0:出力しない　1:出力する → 契約者Noを契約者番号に出力する  0:出力しない　1:出力する -chg
                                    .Vari_Keiyakusyano_syuturyoku = fldvalue
                                Case "改行コード(CRLF)出力　0:なし　1:あり"
                                    .Vari_Crlf = fldvalue
                                Case "「ｦ」→「ｵ」変換　0:しない　1:する"
                                    .Vari_Wo_mojihenkan = fldvalue
                                Case "送信ファイルパス名"
                                    .Vari_File_sosin = fldvalue
                                Case "受信ファイルパス名"
                                    .Vari_File_jyusin = fldvalue
                                    '20160519 EXEUpdateに伴う修正 口座振替情報 -del sta
                                    'Case "2枚ＦＤ作成手順を使用する　0:しない　1:する"
                                    '    .Vari_Fdsakusei = fldvalue
                                    '20160519 EXEUpdateに伴う修正 口座振替情報 -del end
                                Case "請求のまとめ方　0:請求先単位　1:契約単位"
                                    .Vari_Seikyu_tani = fldvalue
                                Case "滞納分請求　0:しない　1:する"
                                    .Vari_Seikyu_taino = fldvalue
                                Case "滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける"
                                    .Vari_Seikyu_tukitani = fldvalue
                                Case "手数料の請求入金管理をする　0:しない　1:する"
                                    .Vari_Tesu_nyukinkanri = fldvalue
                                Case "未入金の手数料は次回請求に加える　0:加えない　1:加える"
                                    .Vari_Tesu_kurikosi = fldvalue
                                Case "照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合"
                                    .Vari_Syogorule = fldvalue
                                Case "ゆうちょ銀行金融機関番号指定　不使用の場合はNULL"
                                    .Vari_Yutyobango = fldvalue
                                    '20160519 EXEUpdateに伴う修正 口座振替情報 -add sta
                                Case "マルチヘッダー形式フラグ"
                                    .Vari_Yutyobango = fldvalue
                                Case "ゆうちょ銀行付加コード"
                                    .Vari_Yutyobango = fldvalue
                                Case "振替でのゆうちょ銀行コードに振込用変換を利用するフラグ"
                                    .Vari_Yutyobango = fldvalue
                                Case "振替入金処理時に受信ファイルを読み込まないフラグ"
                                    .Vari_Yutyobango = fldvalue
                                    '20160519 EXEUpdateに伴う修正 口座振替情報 -add end
                                    '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
                                Case "自社支店No"
                                    tmp_jisyano = fldvalue
                                Case "振替先口座No"                 '20160929 汎用→既存コピー処理改善対応 振替先口座No. → 振替先口座No -chg
                                    tmp_jisyakozano = fldvalue
                                    '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end
                                Case "再振替対応利用フラグ"   '20161012 革命10アップデートに伴う修正 -add
                                    .Vari_Saifkae_use = fldvalue
                            End Select

                        Next
                        '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg sta
                        ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                        ' ''20160609 紐付取得用に追加 -chg sta
                        '' ''2016.04.26 メインの方へも反映させる修正 -add sta
                        '' ''自社No/自社口座No取得のため口座情報を格納する
                        ' ''.Vari_Jisya_no = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                        ' ''.Vari_Fkae_fb_nkinukekozano = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                        '' ''2016.04.26 メインの方へも反映させる修正 -add end
                        ''Dim tmp_jisyainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
                        ''.Vari_Jisya_no = tmp_jisyainfo
                        ''.Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
                        ' ''20160609 紐付取得用に追加 -chg end
                        'Dim tmp_kozamoto As String = "口座振替マスタ"
                        'Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Fkae_no
                        '.Vari_Jisya_no = tmp_jisyainfo
                        '.Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
                        ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                        Select Case CNVNO
                            Case ConvertTypes._既存ユーザ用
                                Dim tmp_kozamoto As String = "口座振替マスタ"
                                Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Fkae_no
                                .Vari_Jisya_no = tmp_jisyainfo
                                .Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
                            Case ConvertTypes._汎用
                                .Vari_Fkae_fb_orgfmtno = ""
                                '20160928 自社口座取得方法の修正 -chg sta
                                '.Vari_Jisya_no = tmp_jisyano
                                '.Vari_Fkae_fb_nkinukekozano = tmp_jisyakozano
                                .Vari_Jisya_no = tmp_jisyano & "-" & tmp_jisyakozano
                                .Vari_Fkae_fb_nkinukekozano = tmp_jisyano & "-" & tmp_jisyakozano
                                '20160928 自社口座取得方法の修正 -chg end
                        End Select
                        '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg end
                        '固定値
                        .Vari_Useflg = 1
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20160928 自社口座取得方法の修正 -add sta
                        tmp_hash("jisya_no") = tmp_jisyano
                        tmp_hash("fkae_fb_nkinukekozano") = tmp_jisyakozano
                        '20160928 自社口座取得方法の修正 -add end

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

            ' ''' <summary>
            ' ''' 【中間ファイル→変数】
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="syorikomok"></param>
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

            '    Dim model_cvitem As New Njc.Model.M_fb_fkaejyoho_Model          '移行値格納用モデル初期化
            '    Dim keycol As Integer = 1                                       'キー列

            '    '************************
            '    '作業準備
            '    '************************

            '    '2016.04.26 メインの方へも反映させる修正 -add sta
            '    '紐付けデータ取得
            '    Call SetRelItemToObject.Set_RelData_Jisyakoza()
            '    Call SetRelItemToObject.Set_RelData_FBInfo()        '20160609 紐付設定値取得に伴う修正 -add
            '    '2016.04.26 メインの方へも反映させる修正 -add end

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If Not rtn Then
            '        Return rtn
            '    End If

            '    'テーブル名/フィールド名セット
            '    Dim tblname As String = "m_fb_fkaejyoho"
            '    '20160519 EXEUpdateに伴う修正 口座振替情報 -chg sta
            '    'Dim fldnamegrp As String = "fkae_no,fkae_name,fkae_kana,servicetype,fkae_fb_fkomiraino," & _
            '    '                           "fkae_fb_fkomiraikana,kamei_no,jlease_nyukinkbn,jisya_no,fkae_fb_nkinukekozano," & _
            '    '                           "hikiotosibi,tesu_gak,biko,fkae_fb_orgfmtno,datasort," & _
            '    '                           "keiyakusyano_syuturyoku,crlf,wo_mojihenkan,file_sosin,file_jyusin," & _
            '    '                           "fdsakusei,seikyu_tani,seikyu_taino,seikyu_tukitani,tesu_nyukinkanri," & _
            '    '                           "tesu_kurikosi,syogorule,yutyobango,history,useflg"
            '    Dim fldnamegrp As String = "fkae_no,fkae_name,fkae_kana,fkae_fb_fkomiraino,fkae_fb_fkomiraikana," & _
            '                               "kamei_no,jlease_nyukinkbn,jisya_no,fkae_fb_nkinukekozano,hikiotosibi," & _
            '                               "tesu_gak,biko,fkae_fb_orgfmtno,datasort,keiyakusyano_syuturyoku," & _
            '                               "crlf,wo_mojihenkan,file_sosin,file_jyusin,seikyu_tani," & _
            '                               "seikyu_taino,seikyu_tukitani,tesu_nyukinkanri,tesu_kurikosi,syogorule," & _
            '                               "yutyobango,history,useflg,multi_flg,yutyo_appendcode," & _
            '                               "yutyo_usefkomcode,resultfile_notuse"
            '    '20160519 EXEUpdateに伴う修正 口座振替情報 -chg end

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

            '        '---------------
            '        'ヘッダー処理
            '        '---------------
            '        'ヘッダー行取得
            '        Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '        'キーヘッダー名取得
            '        Dim fldname_key As String = headervalue(startrow - 1, keycol)

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
            '            Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

            '            'キー値取得
            '            Dim fldvalue_key As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))

            '            'ログ出力用
            '            Dim str_logkey As String = fldname_key & " = " & fldvalue_key

            '            '2016.04.26 メインの方へも反映させる修正 -add sta
            '            '作業用変数
            '            Dim tmp_kinyuno As String = ""
            '            Dim tmp_kinyutenno As String = ""
            '            Dim tmp_kozasyu As String = ""      '20160609 紐付取得用に追加 -add
            '            Dim tmp_kozabango As String = ""
            '            Dim tmp_jisyano As String = ""      '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
            '            Dim tmp_jisyakozano As String = ""  '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
            '            '2016.04.26 メインの方へも反映させる修正 -add end

            '            '移行値取得
            '            For cntjj = 1 To columncnt

            '                Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            '                Dim fldvalue As String = ""
            '                If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            '                End If

            '                Select Case fldname
            '                    Case "振替情報No"                          '20160929 汎用→既存コピー処理改善対応 振替情報No. → 振替情報No -chg
            '                        .Vari_Fkae_no = fldvalue
            '                        .Vari_Fkae_fb_orgfmtno = fldvalue       '20160609 紐付設定値取得に伴う修正 -add
            '                    Case "振替情報名称"
            '                        .Vari_Fkae_name = fldvalue
            '                    Case "振替情報カナ名称"
            '                        .Vari_Fkae_kana = fldvalue
            '                        '20160519 EXEUpdateに伴う修正 口座振替情報 -del sta
            '                        'Case "サービスタイプ　1:オリコ　2:ジャックス　3:アプラス"
            '                        '    .Vari_Servicetype = fldvalue
            '                        '20160519 EXEUpdateに伴う修正 口座振替情報 -del end
            '                    Case "振込依頼人"
            '                        .Vari_Fkae_fb_fkomiraino = fldvalue
            '                    Case "振込依頼人カナ名"
            '                        .Vari_Fkae_fb_fkomiraikana = fldvalue
            '                    Case "加盟店No 依頼人番号と共用"          '20160929 汎用→既存コピー処理改善対応 加盟店No.　依頼人番号と共用 → 加盟店No 依頼人番号と共用 -chg
            '                        .Vari_Kamei_no = fldvalue
            '                    Case "ジェイリース　入金区分　1:変更なし・・・・"
            '                        .Vari_Jlease_nyukinkbn = fldvalue
            '                        '2016.04.26 メインの方へも反映させる修正 -chg sta
            '                        'Case "自社支店No"
            '                        '    .Vari_Jisya_no = fldvalue
            '                        'Case "振替先口座No."
            '                        '    .Vari_Fkae_fb_nkinukekozano = fldvalue
            '                    Case "金融機関No"
            '                        tmp_kinyuno = fldvalue
            '                    Case "金融機関支店No"
            '                        tmp_kinyutenno = fldvalue
            '                    Case "口座種別" '20160609 紐付取得用に追加 -add
            '                        tmp_kozasyu = fldvalue
            '                    Case "口座番号"
            '                        tmp_kozabango = fldvalue
            '                        '2016.04.26 メインの方へも反映させる修正 -chg end
            '                    Case "引落日"
            '                        .Vari_Hikiotosibi = fldvalue
            '                    Case "手数料　請求額"
            '                        .Vari_Tesu_gak = fldvalue
            '                    Case "備考"
            '                        .Vari_Biko = fldvalue
            '                        '20160609 紐付設定値取得に伴う修正 -del sta
            '                        'Case "オリジナルフォーマットNo"
            '                        '    .Vari_Fkae_fb_orgfmtno = fldvalue
            '                        '20160609 紐付設定値取得に伴う修正 -del end
            '                    Case "データ並び替え　0:物件番号と部屋番号　1:契約者番号"
            '                        .Vari_Datasort = fldvalue
            '                    Case "契約者Noを契約者番号に出力する  0:出力しない　1:出力する"    '20160929 汎用→既存コピー処理改善対応 契約者No.を契約者番号に出力する  0:出力しない　1:出力する → 契約者Noを契約者番号に出力する  0:出力しない　1:出力する -chg
            '                        .Vari_Keiyakusyano_syuturyoku = fldvalue
            '                    Case "改行コード(CRLF)出力　0:なし　1:あり"
            '                        .Vari_Crlf = fldvalue
            '                    Case "「ｦ」→「ｵ」変換　0:しない　1:する"
            '                        .Vari_Wo_mojihenkan = fldvalue
            '                    Case "送信ファイルパス名"
            '                        .Vari_File_sosin = fldvalue
            '                    Case "受信ファイルパス名"
            '                        .Vari_File_jyusin = fldvalue
            '                        '20160519 EXEUpdateに伴う修正 口座振替情報 -del sta
            '                        'Case "2枚ＦＤ作成手順を使用する　0:しない　1:する"
            '                        '    .Vari_Fdsakusei = fldvalue
            '                        '20160519 EXEUpdateに伴う修正 口座振替情報 -del end
            '                    Case "請求のまとめ方　0:請求先単位　1:契約単位"
            '                        .Vari_Seikyu_tani = fldvalue
            '                    Case "滞納分請求　0:しない　1:する"
            '                        .Vari_Seikyu_taino = fldvalue
            '                    Case "滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける"
            '                        .Vari_Seikyu_tukitani = fldvalue
            '                    Case "手数料の請求入金管理をする　0:しない　1:する"
            '                        .Vari_Tesu_nyukinkanri = fldvalue
            '                    Case "未入金の手数料は次回請求に加える　0:加えない　1:加える"
            '                        .Vari_Tesu_kurikosi = fldvalue
            '                    Case "照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合"
            '                        .Vari_Syogorule = fldvalue
            '                    Case "ゆうちょ銀行金融機関番号指定　不使用の場合はNULL"
            '                        .Vari_Yutyobango = fldvalue
            '                        '20160519 EXEUpdateに伴う修正 口座振替情報 -add sta
            '                    Case "マルチヘッダー形式フラグ"
            '                        .Vari_Yutyobango = fldvalue
            '                    Case "ゆうちょ銀行付加コード"
            '                        .Vari_Yutyobango = fldvalue
            '                    Case "振替でのゆうちょ銀行コードに振込用変換を利用するフラグ"
            '                        .Vari_Yutyobango = fldvalue
            '                    Case "振替入金処理時に受信ファイルを読み込まないフラグ"
            '                        .Vari_Yutyobango = fldvalue
            '                        '20160519 EXEUpdateに伴う修正 口座振替情報 -add end
            '                        '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
            '                    Case "自社支店No"
            '                        tmp_jisyano = fldvalue
            '                    Case "振替先口座No"                 '20160929 汎用→既存コピー処理改善対応 振替先口座No. → 振替先口座No -chg
            '                        tmp_jisyakozano = fldvalue
            '                        '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end
            '                End Select

            '            Next
            '            '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg sta
            '            ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
            '            ' ''20160609 紐付取得用に追加 -chg sta
            '            '' ''2016.04.26 メインの方へも反映させる修正 -add sta
            '            '' ''自社No/自社口座No取得のため口座情報を格納する
            '            ' ''.Vari_Jisya_no = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
            '            ' ''.Vari_Fkae_fb_nkinukekozano = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
            '            '' ''2016.04.26 メインの方へも反映させる修正 -add end
            '            ''Dim tmp_jisyainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
            '            ''.Vari_Jisya_no = tmp_jisyainfo
            '            ''.Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
            '            ' ''20160609 紐付取得用に追加 -chg end
            '            'Dim tmp_kozamoto As String = "口座振替マスタ"
            '            'Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Fkae_no
            '            '.Vari_Jisya_no = tmp_jisyainfo
            '            '.Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
            '            ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
            '            Select Case CNVNO
            '                Case ConvertTypes._既存ユーザ用
            '                    Dim tmp_kozamoto As String = "口座振替マスタ"
            '                    Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Fkae_no
            '                    .Vari_Jisya_no = tmp_jisyainfo
            '                    .Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
            '                Case ConvertTypes._汎用
            '                    .Vari_Fkae_fb_orgfmtno = ""
            '                    '20160928 自社口座取得方法の修正 -chg sta
            '                    '.Vari_Jisya_no = tmp_jisyano
            '                    '.Vari_Fkae_fb_nkinukekozano = tmp_jisyakozano
            '                    .Vari_Jisya_no = tmp_jisyano & "-" & tmp_jisyakozano
            '                    .Vari_Fkae_fb_nkinukekozano = tmp_jisyano & "-" & tmp_jisyakozano
            '                    '20160928 自社口座取得方法の修正 -chg end
            '            End Select
            '            '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg end
            '            '固定値
            '            .Vari_Useflg = 1
            '            .Vari_History = DefHistory

            '            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            '            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '            'データチェック
            '            Dim skipflg As Boolean = False
            '            Dim hash_cvitem As New Hashtable
            '            Dim hash_log As New Hashtable
            '            skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

            '            '20160928 自社口座取得方法の修正 -add sta
            '            tmp_hash("jisya_no") = tmp_jisyano
            '            tmp_hash("fkae_fb_nkinukekozano") = tmp_jisyakozano
            '            '20160928 自社口座取得方法の修正 -add end

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


                Dim tmp_sql As String = " SELECT jisya_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)


            End Sub

        End Class

    End Class

#End Region

#Region "入出金取得情報"       '20160517 入出金取得情報の新規作成

    Public Class M_fb_nskinsetting_Repository

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

                Dim model_cvitem As New Njc.Model.M_fb_nskinsetting_Model       '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'メインキー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '20160609 紐付設定値取得に伴う修正 -chg sta
                ''FBフォーマット紐付け情報を取得
                'Call SetRelItemToObject.Set_RelData_FBFmt()
                Call SetRelItemToObject.Set_RelData_FBInfo()
                '20160609 紐付設定値取得に伴う修正 -chg end

                'Excelファイル初期設定
                '20160609 紐付設定値取得に伴う修正 -chg sta
                '中間ファイルから読み込むように修正
                ''入出金取得情報は紐付ファイルから読込む
                'Dim readfiledirpath As String = RelationDirPath             '紐付けファイル格納先
                'Dim readfilename As String = MIDFILE_RELNAME                '紐付けファイル名
                'Dim readsheetname As String = "入出金取得情報マスタ"        '紐付けシート名
                'rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, readfiledirpath, readfilename, readsheetname)
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)
                '20160609 紐付設定値取得に伴う修正 -chg end

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_fb_nskinsetting"
                '20160519 EXEUpdateに伴う修正 入出金取得情報 -del(crlfを削除して成形)
                Dim fldnamegrp As String = "ns_no,ns_name,ns_kana,jisya_no,orgfmt_no," & _
                                           "file_jyusin,data_syubetu,biko,history,useflg," & _
                                           "yatin_kozano,file_sosin,yokinbunkatusetting,duplicate_check"

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
                        Dim tmp_fmtno As String = ""  '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
                        Dim tmp_jisyano As String = ""  '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "入出金取得No"
                                    .Vari_Ns_no = fldvalue
                                    .Vari_Orgfmt_no = fldvalue   '20160609 紐付設定値取得に伴う修正 -add
                                    .Vari_Jisya_no = fldvalue       '20160609 紐付設定値取得に伴う修正 -add
                                Case "設定名称"
                                    .Vari_Ns_name = fldvalue
                                Case "設定カナ名称"
                                    .Vari_Ns_kana = fldvalue
                                    '20160609 紐付設定値取得に伴う修正 -del sta
                                    'Case "フォーマット定義名"
                                    '    .Vari_Orgfmt_no = fldvalue
                                    '20160609 紐付設定値取得に伴う修正 -del end
                                Case "受信ファイル名（フルパス）"
                                    .Vari_File_jyusin = fldvalue
                                Case "取込データ種別"
                                    .Vari_Data_syubetu = fldvalue
                                    '20160519 EXEUpdateに伴う修正 入出金取得情報 -del sta
                                    'Case "改行コード有無"
                                    '    .Vari_Crlf = fldvalues
                                    '20160519 EXEUpdateに伴う修正 入出金取得情報 -del end
                                Case "備考"
                                    .Vari_Biko = fldvalue
                                Case "家賃口座No"
                                    .Vari_Yatin_kozano = fldvalue
                                Case "送信ファイル名（フルパス）"
                                    .Vari_File_sosin = fldvalue
                                Case "預金分割設定"
                                    .Vari_Yokinbunkatusetting = fldvalue
                                Case "重複チェック"
                                    .Vari_Duplicate_check = fldvalue
                                    '20160609 紐付設定値取得に伴う修正 -del sta
                                    'Case "自社No"
                                    '    .Vari_Jisya_no = fldvalue
                                    '20160609 紐付設定値取得に伴う修正 -del end
                                    '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
                                Case "自社支店No"
                                    tmp_jisyano = fldvalue
                                Case "フォーマット定義名"
                                    tmp_fmtno = fldvalue
                                    '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end
                            End Select

                        Next

                        '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
                        If CNVNO = ConvertTypes._汎用 Then
                            .Vari_Orgfmt_no = tmp_fmtno
                            .Vari_Jisya_no = tmp_jisyano
                        End If
                        '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end

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

                Dim tmp_sql As String = " SELECT ns_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "家賃入金口座情報"       '20160517 家賃入金口座情報の新規作成

    Public Class M_yatinkoza_Repository

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

                Dim model_cvitem As New Njc.Model.M_yatinkoza_Model             '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列

                '************************
                '作業準備
                '************************

                '20160609 紐付取得用に追加 -add sta
                '紐付けデータ取得
                Call SetRelItemToObject.Set_RelData_Jisyakoza()
                '20160609 紐付取得用に追加 -add end

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
                Dim tblname As String = "m_yatinkoza"
                Dim fldnamegrp As String = "yatin_kozano,yatin_kozaname,yatin_kozanamesjis,yatin_kozakana,sqsaki_kbn," & _
                                           "jisya_no,jisya_kozano,ow_no,ow_kozano,spc_useflg," & _
                                           "spc_kozasiteikbn,spc_syumokukbn,spc_ninibango,spc_kanyubango,spc_kobetuflg," & _
                                           "spc_kobetukbn,spc_kobetuareano,spc_kobetutikuno,spc_kobetutel,spc_ansyobango," & _
                                           "spc_servicecode,history,useflg,yatin_kozabiko,headertemplate," & _
                                           "trallertemplate"

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

                '20160624 家賃入金口座情報 重複ログ出力制御修正 -add
                Dim chg_totalcnt As Integer = rowcnt    '重複データ件数調整用(カウントしないため)

                '************************
                '処理開始
                '************************
                With model_cvitem

                    'キーヘッダー名取得
                    Dim fldname_key As String = readtbl.Columns(keycol - 1).ColumnName.Trim

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
                        Dim tmp_kinyuno As String = ""
                        Dim tmp_kinyutenno As String = ""
                        Dim tmp_kozasyu As String = ""
                        Dim tmp_kozabango As String = ""
                        Dim tmp_jisyano As String = ""      '20160928 自社口座取得方法の修正 -add
                        Dim tmp_jisyakozano As String = ""  '20160928 自社口座取得方法の修正 -add
                        Dim tmp_owno As String = ""      '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                        Dim tmp_owkozano As String = ""  '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add

                        'キー値取得
                        Dim fldvalue_key As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol - 1)).Trim

                        'ログ出力用
                        Dim str_logkey As String = fldname_key & " = " & fldvalue_key

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "家賃入金受付口座No"
                                    .Vari_Yatin_kozano = fldvalue
                                Case "受取人名"
                                    .Vari_Yatin_kozaname = fldvalue
                                    '20160609 紐付取得用に追加 -add sta
                                Case "金融機関No"
                                    tmp_kinyuno = fldvalue
                                Case "金融機関支店No"
                                    tmp_kinyutenno = fldvalue
                                Case "口座種別" '20160609 紐付取得用に追加 -add
                                    tmp_kozasyu = fldvalue
                                Case "口座番号"
                                    tmp_kozabango = fldvalue
                                    '20160609 紐付取得用に追加 -add end
                                Case "受取人名Unicode"
                                    .Vari_Yatin_kozanamesjis = fldvalue
                                Case "受取人カナ"
                                    .Vari_Yatin_kozakana = fldvalue
                                Case "請求先区分"
                                    .Vari_Sqsaki_kbn = fldvalue
                                Case "自社支店No"
                                    '20160928 自社口座取得方法の修正 -chg sta
                                    '.Vari_Jisya_no = fldvalue
                                    tmp_jisyano = fldvalue
                                    '20160928 自社口座取得方法の修正 -chg end
                                Case "自社口座No"
                                    '20160928 自社口座取得方法の修正 -chg sta
                                    '.Vari_Jisya_kozano = fldvalue
                                    tmp_jisyakozano = fldvalue
                                    '20160928 自社口座取得方法の修正 -chg end
                                Case "家主No"
                                    '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg sta
                                    '.Vari_Ow_no = fldvalue
                                    tmp_owno = fldvalue
                                    '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg end
                                Case "家主口座No"
                                    '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg sta
                                    '.Vari_Ow_kozano = fldvalue
                                    tmp_owkozano = fldvalue
                                    '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg end
                                Case "ANSER-SPC有無"
                                    .Vari_Spc_useflg = fldvalue
                                Case "ANSER-SPC口座指定方式"
                                    .Vari_Spc_kozasiteikbn = fldvalue
                                Case "ANSER-SPC種目付加方法"
                                    .Vari_Spc_syumokukbn = fldvalue
                                Case "ANSER-SPC任意番号"
                                    .Vari_Spc_ninibango = fldvalue
                                Case "ANSER-SPC加入者番号"
                                    .Vari_Spc_kanyubango = fldvalue
                                Case "ANSER-SPC接続先個別設定フラグ"
                                    .Vari_Spc_kobetuflg = fldvalue
                                Case "ANSER-SPC接続方法"
                                    .Vari_Spc_kobetukbn = fldvalue
                                Case "ANSER-SPCエリアNo"
                                    .Vari_Spc_kobetuareano = fldvalue
                                Case "ANSER-SPC地区No"
                                    .Vari_Spc_kobetutikuno = fldvalue
                                Case "ANSER-SPCTEL"
                                    .Vari_Spc_kobetutel = fldvalue
                                Case "ANSER-SPC照会用暗証番号"
                                    .Vari_Spc_ansyobango = fldvalue
                                Case "ANSER-SPCサービスコード"
                                    .Vari_Spc_servicecode = fldvalue
                                Case "家賃入金口座備考"
                                    .Vari_Yatin_kozabiko = fldvalue
                                Case "ヘッダーテンプレート"
                                    .Vari_Headertemplate = fldvalue
                                Case "トレーラーテンプレート"
                                    .Vari_Trallertemplate = fldvalue
                            End Select

                        Next
                        '20160928 自社口座取得方法の修正 -chg sta
                        ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                        ' ''20160609 紐付取得用に追加 -add sta
                        ''Dim tmp_jisyainfo As String = ""
                        ''If .Vari_Ow_no = "" OrElse .Vari_Ow_kozano = "" Then
                        ''    tmp_jisyainfo = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
                        ''End If
                        ''.Vari_Jisya_no = tmp_jisyainfo
                        ''.Vari_Jisya_kozano = tmp_jisyainfo
                        ' ''20160609 紐付取得用に追加 -add end
                        'Dim tmp_kozamoto As String = "振込先口座マスタ"
                        'Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Yatin_kozano
                        '.Vari_Jisya_no = tmp_jisyainfo
                        '.Vari_Jisya_kozano = tmp_jisyainfo
                        ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                        Select Case CNVNO
                            Case ConvertTypes._既存ユーザ用
                                Dim tmp_kozamoto As String = "振込先口座マスタ"
                                Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Yatin_kozano
                                .Vari_Jisya_no = tmp_jisyainfo
                                .Vari_Jisya_kozano = tmp_jisyainfo
                                .Vari_Ow_no = tmp_owno          '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                                .Vari_Ow_kozano = tmp_owkozano  '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                            Case ConvertTypes._汎用
                                .Vari_Jisya_no = tmp_jisyano & "-" & tmp_jisyakozano
                                .Vari_Jisya_kozano = tmp_jisyano & "-" & tmp_jisyakozano
                                .Vari_Ow_no = tmp_owno & "-" & tmp_owkozano         '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                                .Vari_Ow_kozano = tmp_owno & "-" & tmp_owkozano     '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                        End Select
                        '20160928 自社口座取得方法の修正 -chg end
                        '固定値
                        .Vari_Useflg = 1
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20160928 自社口座取得方法の修正 -add sta
                        tmp_hash("jisya_no") = tmp_jisyano
                        tmp_hash("jisya_kozano") = tmp_jisyakozano
                        '20160928 自社口座取得方法の修正 -add end

                        '20160930 家賃入金口座移行制御処理の追加 -add sta
                        If CNVNO = ConvertTypes._汎用 Then
                            skipflg = Not (Me.Chk_ExistJisyaOwkoza(tblname, hash_cvitem, hash_log))
                        End If
                        '20160930 家賃入金口座移行制御処理の追加 -add end

                        '書込処理
                        If Not skipflg Then

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1
                            End If

                            '20160624 家賃入金口座情報 重複ログ出力制御修正 -add sta
                        Else
                            If hash_log.Count = 1 Then
                                Dim logclearflg As Boolean = False
                                For Each erritem In hash_log
                                    Dim tmp_str As String = erritem.Value
                                    Dim tmp_errnaiyo() As String = tmp_str.Split("-")
                                    If tmp_errnaiyo(0) = LOG_NAIYO_ERR_OVERLAP Then
                                        chg_totalcnt = chg_totalcnt - 1
                                        logclearflg = True
                                    End If
                                Next
                                If logclearflg Then
                                    hash_log.Clear()
                                End If
                            End If
                            '20160624 家賃入金口座情報 重複ログ出力制御修正 -add end
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
                '20160624 家賃入金口座情報 重複ログ出力制御修正 -chg sta
                'midrowcnt = rowcnt
                midrowcnt = chg_totalcnt
                '20160624 家賃入金口座情報 重複ログ出力制御修正 -chg end

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

            ' ''' <summary>
            ' ''' 【中間ファイル→変数】
            ' ''' </summary>
            ' ''' <param name="sqlcnnv10"></param>
            ' ''' <param name="syorikomok"></param>
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

            '    Dim model_cvitem As New Njc.Model.M_yatinkoza_Model             '移行値格納用モデル初期化
            '    Dim keycol As Integer = 1                                       'キー列

            '    '************************
            '    '作業準備
            '    '************************

            '    '20160609 紐付取得用に追加 -add sta
            '    '紐付けデータ取得
            '    Call SetRelItemToObject.Set_RelData_Jisyakoza()
            '    '20160609 紐付取得用に追加 -add end

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If Not rtn Then
            '        Return rtn
            '    End If

            '    'テーブル名/フィールド名セット
            '    Dim tblname As String = "m_yatinkoza"
            '    Dim fldnamegrp As String = "yatin_kozano,yatin_kozaname,yatin_kozanamesjis,yatin_kozakana,sqsaki_kbn," & _
            '                               "jisya_no,jisya_kozano,ow_no,ow_kozano,spc_useflg," & _
            '                               "spc_kozasiteikbn,spc_syumokukbn,spc_ninibango,spc_kanyubango,spc_kobetuflg," & _
            '                               "spc_kobetukbn,spc_kobetuareano,spc_kobetutikuno,spc_kobetutel,spc_ansyobango," & _
            '                               "spc_servicecode,history,useflg,yatin_kozabiko,headertemplate," & _
            '                               "trallertemplate"

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

            '    '20160624 家賃入金口座情報 重複ログ出力制御修正 -add
            '    Dim chg_totalcnt As Integer = rowcnt    '重複データ件数調整用(カウントしないため)

            '    '************************
            '    '処理開始
            '    '************************
            '    With model_cvitem

            '        '---------------
            '        'ヘッダー処理
            '        '---------------
            '        'ヘッダー行取得
            '        Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '        'キーヘッダー名取得
            '        Dim fldname_key As String = headervalue(startrow - 1, keycol)

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
            '            Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

            '            'キー値取得
            '            Dim fldvalue_key As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))

            '            'ログ出力用
            '            Dim str_logkey As String = fldname_key & " = " & fldvalue_key

            '            '20160609 紐付取得用に追加 -add sta
            '            '作業用変数
            '            Dim tmp_kinyuno As String = ""
            '            Dim tmp_kinyutenno As String = ""
            '            Dim tmp_kozasyu As String = ""
            '            Dim tmp_kozabango As String = ""
            '            Dim tmp_jisyano As String = ""      '20160928 自社口座取得方法の修正 -add
            '            Dim tmp_jisyakozano As String = ""  '20160928 自社口座取得方法の修正 -add
            '            Dim tmp_owno As String = ""      '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            '            Dim tmp_owkozano As String = ""  '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            '            '20160609 紐付取得用に追加 -add end

            '            '移行値取得
            '            For cntjj = 1 To columncnt

            '                Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            '                Dim fldvalue As String = ""
            '                If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            '                End If

            '                Select Case fldname
            '                    Case "家賃入金受付口座No"
            '                        .Vari_Yatin_kozano = fldvalue
            '                    Case "受取人名"
            '                        .Vari_Yatin_kozaname = fldvalue
            '                        '20160609 紐付取得用に追加 -add sta
            '                    Case "金融機関No"
            '                        tmp_kinyuno = fldvalue
            '                    Case "金融機関支店No"
            '                        tmp_kinyutenno = fldvalue
            '                    Case "口座種別" '20160609 紐付取得用に追加 -add
            '                        tmp_kozasyu = fldvalue
            '                    Case "口座番号"
            '                        tmp_kozabango = fldvalue
            '                        '20160609 紐付取得用に追加 -add end
            '                    Case "受取人名Unicode"
            '                        .Vari_Yatin_kozanamesjis = fldvalue
            '                    Case "受取人カナ"
            '                        .Vari_Yatin_kozakana = fldvalue
            '                    Case "請求先区分"
            '                        .Vari_Sqsaki_kbn = fldvalue
            '                    Case "自社支店No"
            '                        '20160928 自社口座取得方法の修正 -chg sta
            '                        '.Vari_Jisya_no = fldvalue
            '                        tmp_jisyano = fldvalue
            '                        '20160928 自社口座取得方法の修正 -chg end
            '                    Case "自社口座No"
            '                        '20160928 自社口座取得方法の修正 -chg sta
            '                        '.Vari_Jisya_kozano = fldvalue
            '                        tmp_jisyakozano = fldvalue
            '                        '20160928 自社口座取得方法の修正 -chg end
            '                    Case "家主No"
            '                        '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg sta
            '                        '.Vari_Ow_no = fldvalue
            '                        tmp_owno = fldvalue
            '                        '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg end
            '                    Case "家主口座No"
            '                        '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg sta
            '                        '.Vari_Ow_kozano = fldvalue
            '                        tmp_owkozano = fldvalue
            '                        '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg end
            '                    Case "ANSER-SPC有無"
            '                        .Vari_Spc_useflg = fldvalue
            '                    Case "ANSER-SPC口座指定方式"
            '                        .Vari_Spc_kozasiteikbn = fldvalue
            '                    Case "ANSER-SPC種目付加方法"
            '                        .Vari_Spc_syumokukbn = fldvalue
            '                    Case "ANSER-SPC任意番号"
            '                        .Vari_Spc_ninibango = fldvalue
            '                    Case "ANSER-SPC加入者番号"
            '                        .Vari_Spc_kanyubango = fldvalue
            '                    Case "ANSER-SPC接続先個別設定フラグ"
            '                        .Vari_Spc_kobetuflg = fldvalue
            '                    Case "ANSER-SPC接続方法"
            '                        .Vari_Spc_kobetukbn = fldvalue
            '                    Case "ANSER-SPCエリアNo"
            '                        .Vari_Spc_kobetuareano = fldvalue
            '                    Case "ANSER-SPC地区No"
            '                        .Vari_Spc_kobetutikuno = fldvalue
            '                    Case "ANSER-SPCTEL"
            '                        .Vari_Spc_kobetutel = fldvalue
            '                    Case "ANSER-SPC照会用暗証番号"
            '                        .Vari_Spc_ansyobango = fldvalue
            '                    Case "ANSER-SPCサービスコード"
            '                        .Vari_Spc_servicecode = fldvalue
            '                    Case "家賃入金口座備考"
            '                        .Vari_Yatin_kozabiko = fldvalue
            '                    Case "ヘッダーテンプレート"
            '                        .Vari_Headertemplate = fldvalue
            '                    Case "トレーラーテンプレート"
            '                        .Vari_Trallertemplate = fldvalue
            '                End Select

            '            Next
            '            '20160928 自社口座取得方法の修正 -chg sta
            '            ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
            '            ' ''20160609 紐付取得用に追加 -add sta
            '            ''Dim tmp_jisyainfo As String = ""
            '            ''If .Vari_Ow_no = "" OrElse .Vari_Ow_kozano = "" Then
            '            ''    tmp_jisyainfo = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
            '            ''End If
            '            ''.Vari_Jisya_no = tmp_jisyainfo
            '            ''.Vari_Jisya_kozano = tmp_jisyainfo
            '            ' ''20160609 紐付取得用に追加 -add end
            '            'Dim tmp_kozamoto As String = "振込先口座マスタ"
            '            'Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Yatin_kozano
            '            '.Vari_Jisya_no = tmp_jisyainfo
            '            '.Vari_Jisya_kozano = tmp_jisyainfo
            '            ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
            '            Select Case CNVNO
            '                Case ConvertTypes._既存ユーザ用
            '                    Dim tmp_kozamoto As String = "振込先口座マスタ"
            '                    Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Yatin_kozano
            '                    .Vari_Jisya_no = tmp_jisyainfo
            '                    .Vari_Jisya_kozano = tmp_jisyainfo
            '                    .Vari_Ow_no = tmp_owno          '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            '                    .Vari_Ow_kozano = tmp_owkozano  '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            '                Case ConvertTypes._汎用
            '                    .Vari_Jisya_no = tmp_jisyano & "-" & tmp_jisyakozano
            '                    .Vari_Jisya_kozano = tmp_jisyano & "-" & tmp_jisyakozano
            '                    .Vari_Ow_no = tmp_owno & "-" & tmp_owkozano         '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            '                    .Vari_Ow_kozano = tmp_owno & "-" & tmp_owkozano     '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            '            End Select
            '            '20160928 自社口座取得方法の修正 -chg end
            '            '固定値
            '            .Vari_Useflg = 1
            '            .Vari_History = DefHistory

            '            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            '            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '            'データチェック
            '            Dim skipflg As Boolean = False
            '            Dim hash_cvitem As New Hashtable
            '            Dim hash_log As New Hashtable
            '            skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

            '            '20160928 自社口座取得方法の修正 -add sta
            '            tmp_hash("jisya_no") = tmp_jisyano
            '            tmp_hash("jisya_kozano") = tmp_jisyakozano
            '            '20160928 自社口座取得方法の修正 -add end

            '            '20160930 家賃入金口座移行制御処理の追加 -add sta
            '            If CNVNO = ConvertTypes._汎用 Then
            '                skipflg = Not (Me.Chk_ExistJisyaOwkoza(tblname, hash_cvitem, hash_log))
            '            End If
            '            '20160930 家賃入金口座移行制御処理の追加 -add end

            '            '書込処理
            '            If Not skipflg Then

            '                '挿入処理
            '                Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

            '                '挿入処理が正常終了したレコードのキーを重複チェック用に格納
            '                If normalflg Then
            '                    list_chkduplicate.Add(fldvalue_key)
            '                    tmp_cvcnt = tmp_cvcnt + 1
            '                End If

            '                '20160624 家賃入金口座情報 重複ログ出力制御修正 -add sta
            '            Else
            '                If hash_log.Count = 1 Then
            '                    Dim logclearflg As Boolean = False
            '                    For Each erritem In hash_log
            '                        Dim tmp_str As String = erritem.Value
            '                        Dim tmp_errnaiyo() As String = tmp_str.Split("-")
            '                        If tmp_errnaiyo(0) = LOG_NAIYO_ERR_OVERLAP Then
            '                            chg_totalcnt = chg_totalcnt - 1
            '                            logclearflg = True
            '                        End If
            '                    Next
            '                    If logclearflg Then
            '                        hash_log.Clear()
            '                    End If
            '                End If
            '                '20160624 家賃入金口座情報 重複ログ出力制御修正 -add end
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
            '    '20160624 家賃入金口座情報 重複ログ出力制御修正 -chg sta
            '    'midrowcnt = rowcnt
            '    midrowcnt = chg_totalcnt
            '    '20160624 家賃入金口座情報 重複ログ出力制御修正 -chg end

            '    '移行件数を取得
            '    cvrowcnt = tmp_cvcnt

            '    '調整件数を取得
            '    conditioncnt = tmp_condcnt

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            '    '返却
            '    Return rtn

            'End Function

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT yatin_kozano FROM " & tblname
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

                Dim tmp_sql As String = " SELECT yatin_kozano FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 自社口座、家主口座有無チェック '20160930 家賃入金口座移行制御処理の追加
            ''' 無い場合は移行自体を行わない
            ''' </summary>
            ''' <param name="tblname"></param>
            ''' <param name="hash_cvitem"></param>
            ''' <param name="hash_log"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Chk_ExistJisyaOwkoza(ByVal tblname As String, ByRef hash_cvitem As Hashtable, ByRef hash_log As Hashtable) As Boolean

                Dim rtn As Boolean = True

                Dim tmp_sqsakikbn As String = hash_cvitem("sqsaki_kbn")
                Dim tmp_jisyano As String = hash_cvitem("jisya_no")
                Dim tmp_jisyakozano As String = hash_cvitem("jisya_kozano")
                Dim tmp_owno As String = hash_cvitem("ow_no")
                Dim tmp_owkozano As String = hash_cvitem("ow_kozano")
                Dim tmp_errstr As String = ""
                Dim tmp_errfld As String = ""

                '設定された値をチェックする
                If tmp_jisyano <> "" And tmp_jisyakozano <> "" And tmp_owno <> "" And tmp_owkozano <> "" Then
                    '自社、家主共に設定されている場合は自社を優先
                    'ログ出力
                    Dim log_key As String = tblname & "-" & "jisya_kozano"
                    Dim tmp_hubi As String = "自社口座情報、家主口座情報両方にデータが設定されているため自社口座情報を優先して移行します。"
                    Dim errstr As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & tmp_hubi & "-" & LOG_TAISYO_DEFAULT
                    hash_log.Clear()
                    hash_log.Add(log_key, errstr)
                    hash_cvitem("sqsaki_kbn") = "900"
                    hash_cvitem("ow_no") = ""
                    hash_cvitem("ow_kozano") = ""
                    Return rtn
                ElseIf tmp_jisyano <> "" And tmp_jisyakozano <> "" Then
                    '自社のみ設定されている場合は自社を移行
                    hash_cvitem("sqsaki_kbn") = "900"
                    Return rtn
                ElseIf tmp_owno <> "" And tmp_owkozano <> "" Then
                    '家主のみ設定されている場合は家主を移行
                    hash_cvitem("sqsaki_kbn") = "200"
                    Return rtn
                Else
                    '上記3つ以外は全て不完全な口座情報なので移行しない
                    Dim log_key As String = tblname & "-" & "jisya_kozano"
                    Dim tmp_hubi As String = "自社口座情報、または家主口座情報が不完全であるため移行できません。"
                    Dim errstr As String = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & tmp_hubi & "-" & LOG_TAISYO_DEFAULT
                    hash_log.Clear()
                    hash_log.Add(log_key, errstr)
                    rtn = False
                    Return rtn
                End If

                Return rtn

            End Function

        End Class

    End Class

#End Region

#Region "ANSERエリア情報"

    Public Class M_spcarea_Repository

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

                Dim model_cvitem As New Njc.Model.M_spcarea_Model                  '移行値格納用モデル初期化
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
                Dim tblname As String = "m_spcarea"
                Dim fldnamegrp As String = "spc_areano,spc_areaname,biko_spcarea,history"

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

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "SPCエリアNo"
                                    .Vari_Spc_areano = fldvalue
                                Case "SPCエリア名称"
                                    .Vari_Spc_areaname = fldvalue
                                Case "備考"
                                    .Vari_Biko_spcarea = fldvalue
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

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_BaseKey(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

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

                Dim tmp_sql As String = " SELECT spc_areano FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "ANSERアクセスポイント情報"

    Public Class M_spcaccesspoint_Repository

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

                Dim model_cvitem As New Njc.Model.M_spcaccesspoint_Model        '移行値格納用モデル初期化
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
                Dim tblname As String = "m_spcaccesspoint"
                Dim fldnamegrp As String = "spc_tikuno,spc_areano,spc_tikuname,biko,history"

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

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "SPC地区No"
                                    .Vari_Spc_tikuno = fldvalue
                                Case "SPCエリアNo"
                                    .Vari_Spc_areano = fldvalue
                                Case "SPC地区名称"
                                    .Vari_Spc_tikuname = fldvalue
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

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_BaseKey(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

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

                Dim tmp_sql As String = " SELECT spc_tikuno FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "ANSER接続情報"

    Public Class M_spcsetuzoku_Repository

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

                Dim model_cvitem As New Njc.Model.M_spcsetuzoku_Model           '移行値格納用モデル初期化
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
                Dim tblname As String = "m_spcsetuzoku"
                Dim fldnamegrp As String = "spc_setuzokuno,spc_name,spc_kaisensyu,spc_devtype,spc_devname," & _
                                           "spc_setuzoku_hoho,spc_areano,spc_tikuno,spc_tel,spc_gaisen," & _
                                           "spc_retry_kaisu,spc_retry_kankaku,spc_servicecode,spc_biko,history," & _
                                           "spc_crlf"

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

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "接続方法No"
                                    .Vari_Spc_setuzokuno = fldvalue
                                Case "設定名"
                                    .Vari_Spc_name = fldvalue
                                Case "回線種別(1:電話回線 2:ISDN)"
                                    .Vari_Spc_kaisensyu = fldvalue
                                Case "開発タイプ"
                                    .Vari_Spc_devtype = fldvalue
                                Case "開発名"
                                    .Vari_Spc_devname = fldvalue
                                Case "接続方法(1:アクセスポイント 2:TEL)"
                                    .Vari_Spc_setuzoku_hoho = fldvalue
                                Case "エリアNo"
                                    .Vari_Spc_areano = fldvalue
                                Case "地区No"
                                    .Vari_Spc_tikuno = fldvalue
                                Case "電話番号"
                                    .Vari_Spc_tel = fldvalue
                                Case "外線No"
                                    .Vari_Spc_gaisen = fldvalue
                                Case "リトライ回数"
                                    .Vari_Spc_retry_kaisu = fldvalue
                                Case "リトライの間隔"
                                    .Vari_Spc_retry_kankaku = fldvalue
                                Case "利用するサービスコード"
                                    .Vari_Spc_servicecode = fldvalue
                                Case "備考"
                                    .Vari_Spc_biko = fldvalue
                                Case "改行有無"
                                    .Vari_Spc_crlf = fldvalue
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

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_BaseKey(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

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

                Dim tmp_sql As String = " SELECT spc_setuzokuno FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

End Namespace


