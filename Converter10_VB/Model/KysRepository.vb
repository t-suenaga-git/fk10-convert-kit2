Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "契約者情報"

    Public Class Kysdata_Repository

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

                Dim model_cvitem As New Njc.Model.Kysdata_Model                 '移行値格納用モデル初期化
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
                Dim tblname As String = "kysdata"
                '20160519 EXEUpdateに伴う修正 契約者基本情報 -chg sta
                'Dim fldnamegrp As String = "kys_no,kys_name,kys_namesjis,kys_kana,kojinhojin_flg," & _
                '                           "keisyo,post_code,addr1,addr2,tel1," & _
                '                           "tel2,fax,mobiletel1,mobiletel2,mail," & _
                '                           "mobilemail,yusentel_kbn,yusenmail_kbn,birthday,nensyu," & _
                '                           "biko_kihon,gender,honseki,kinmu_name,kinmu_namesjis," & _
                '                           "kinmu_kana,kinmu_postcode,kinmu_addr1,kinmu_addr2,kinmu_tel1," & _
                '                           "kinmu_tel2,kinmu_fax,kinmu_gyosyu,kinmu_busyo,kinmu_nyuryokuym," & _
                '                           "kinmu_nyusyaym,nyukyomae_nyuryokuym,nyukyomae_postcode,nyukyomae_addr1,nyukyomae_addr2," & _
                '                           "nyukyomae_tel1,nyukyomae_tel2,nyukyomae_fax,url,gyosyu," & _
                '                           "daihyo_name,daihyo_namesjis,daihyo_kana,daihyo_yakusyoku,daihyo_atenaflg," & _
                '                           "tanto_name,tanto_namesjis,tanto_kana,tanto_busyo,tanto_yakusyoku," & _
                '                           "tanto_atenaflg,sihonkin,jugyosu,nyuryoku_ym,torihikisaki," & _
                '                           "renraku_name,renraku_namesjis,renraku_kana,renraku_keisyo,renraku_postcode," & _
                '                           "renraku_addr1,renraku_addr2,renraku_tel1,renraku_tel2,renraku_fax," & _
                '                           "renraku_mobiletel1,renraku_mobiletel2,renraku_yusentelkbn,renraku_aidagara,biko_renraku," & _
                '                           "sofu_kbn,sofu_name,sofu_namesjis,sofu_kana,sofu_keisyo," & _
                '                           "sofu_postcode,sofu_addr1,sofu_addr2,sofu_tel1,sofu_tel2," & _
                '                           "sofu_fax,biko_sofu,sofu_hakkofurikomi,sofu_hakkofurikae,sofu_hakkotokusoku," & _
                '                           "furikomituti_bktaniflg,furikaetuti_bktaniflg,furikomi_kasouseflg,hosyo_multipleflg,furikae_hosyoumu," & _
                '                           "useflg,history,rowid,cvpay_umu"
                Dim fldnamegrp As String = "kys_no,kys_name,kys_namesjis,kys_kana,kojinhojin_flg," & _
                                           "keisyo,post_code,addr1,addr2,tel1," & _
                                           "tel2,fax,mobiletel1,mobiletel2,mail," & _
                                           "mobilemail,yusentel_kbn,yusenmail_kbn,birthday,nensyu," & _
                                           "biko_kihon,gender,honseki,kinmu_name,kinmu_namesjis," & _
                                           "kinmu_kana,kinmu_postcode,kinmu_addr1,kinmu_addr2,kinmu_tel1," & _
                                           "kinmu_tel2,kinmu_fax,kinmu_gyosyu,kinmu_busyo,kinmu_nyuryokuym," & _
                                           "kinmu_nyusyaym,nyukyomae_nyuryokuym,nyukyomae_postcode,nyukyomae_addr1,nyukyomae_addr2," & _
                                           "nyukyomae_tel1,nyukyomae_tel2,nyukyomae_fax,url,gyosyu," & _
                                           "daihyo_name,daihyo_namesjis,daihyo_kana,daihyo_yakusyoku,tanto_name," & _
                                           "tanto_namesjis,tanto_kana,tanto_busyo,tanto_yakusyoku,tanto_atenaflg," & _
                                           "sihonkin,jugyosu,nyuryoku_ym,torihikisaki,renraku_name," & _
                                           "renraku_namesjis,renraku_kana,renraku_keisyo,renraku_postcode,renraku_addr1," & _
                                           "renraku_addr2,renraku_tel1,renraku_tel2,renraku_fax,renraku_mobiletel1," & _
                                           "renraku_mobiletel2,renraku_yusentelkbn,renraku_aidagara,biko_renraku,sofu_kbn," & _
                                           "sofu_name,sofu_namesjis,sofu_kana,sofu_keisyo,sofu_postcode," & _
                                           "sofu_addr1,sofu_addr2,sofu_tel1,sofu_tel2,sofu_fax," & _
                                           "biko_sofu,sofu_hakkofurikomi,sofu_hakkofurikae,sofu_hakkotokusoku,furikomi_kasouseflg," & _
                                           "hosyo_multipleflg,furikae_hosyoumu,useflg,history,rowid," & _
                                           "cvpay_umu,daihyo_atenaflg"
                '20160519 EXEUpdateに伴う修正 契約者基本情報 -chg end

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
                        Dim tmp_tantobusyo As String = ""
                        Dim tmp_yakusyoku As String = ""

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
                                Case "契約者No"
                                    .Vari_Kys_no = fldvalue.Trim
                                Case "契約者名"
                                    .Vari_Kys_name = fldvalue.Trim
                                Case "契約者名SJIS"
                                    .Vari_Kys_namesjis = fldvalue.Trim
                                Case "契約者カナ"
                                    .Vari_Kys_kana = fldvalue.Trim
                                Case "個人法人フラグ"
                                    .Vari_Kojinhojin_flg = fldvalue.Trim
                                Case "宛名敬称"
                                    .Vari_Keisyo = fldvalue.Trim
                                Case "郵便番号"
                                    .Vari_Post_code = fldvalue.Trim
                                Case "住所１"
                                    .Vari_Addr1 = fldvalue.Trim
                                Case "住所２"
                                    .Vari_Addr2 = fldvalue.Trim
                                Case "TEL１"
                                    .Vari_Tel1 = fldvalue.Trim
                                Case "TEL２"
                                    .Vari_Tel2 = fldvalue.Trim
                                Case "FAX"
                                    .Vari_Fax = fldvalue.Trim
                                Case "携帯１"
                                    .Vari_Mobiletel1 = fldvalue.Trim
                                Case "携帯２"
                                    .Vari_Mobiletel2 = fldvalue.Trim
                                Case "メールアドレス"
                                    .Vari_Mail = fldvalue.Trim
                                Case "携帯メールアドレス"
                                    .Vari_Mobilemail = fldvalue.Trim
                                Case "優先設定(電話番号)"
                                    .Vari_Yusentel_kbn = fldvalue.Trim
                                Case "優先設定(メールアドレス)"
                                    .Vari_Yusenmail_kbn = fldvalue.Trim
                                Case "生年月日・設立日"
                                    .Vari_Birthday = fldvalue.Trim
                                Case "年収・年商"
                                    .Vari_Nensyu = EtcMethod.Get_RoundingValue(fldvalue.Trim)
                                Case "備考(基本情報)"
                                    .Vari_Biko_kihon = fldvalue.Trim
                                Case "性別"
                                    .Vari_Gender = fldvalue.Trim
                                Case "本籍地"
                                    .Vari_Honseki = fldvalue.Trim
                                Case "勤務先名"
                                    .Vari_Kinmu_name = fldvalue.Trim
                                Case "勤務先名SJIS"
                                    .Vari_Kinmu_namesjis = fldvalue.Trim
                                Case "勤務先カナ"
                                    .Vari_Kinmu_kana = fldvalue.Trim
                                Case "勤務先郵便番号"
                                    .Vari_Kinmu_postcode = fldvalue.Trim
                                Case "勤務先住所１"
                                    .Vari_Kinmu_addr1 = fldvalue.Trim
                                Case "勤務先住所２"
                                    .Vari_Kinmu_addr2 = fldvalue.Trim
                                Case "勤務先TEL１"
                                    .Vari_Kinmu_tel1 = fldvalue.Trim
                                Case "勤務先TEL２"
                                    .Vari_Kinmu_tel2 = fldvalue.Trim
                                Case "勤務先FAX"
                                    .Vari_Kinmu_fax = fldvalue.Trim
                                Case "勤務先業種"
                                    .Vari_Kinmu_gyosyu = fldvalue.Trim
                                Case "勤務先部署"
                                    .Vari_Kinmu_busyo = fldvalue.Trim
                                Case "勤務先情報記入年月"
                                    .Vari_Kinmu_nyuryokuym = fldvalue.Trim
                                Case "勤務先入社年月"
                                    .Vari_Kinmu_nyusyaym = fldvalue.Trim
                                Case "入居前連絡先情報記入年月"
                                    .Vari_Nyukyomae_nyuryokuym = fldvalue.Trim
                                Case "入居前連絡先郵便番号"
                                    .Vari_Nyukyomae_postcode = fldvalue.Trim
                                Case "入居前連絡先住所１"
                                    .Vari_Nyukyomae_addr1 = fldvalue.Trim
                                Case "入居前連絡先住所２"
                                    .Vari_Nyukyomae_addr2 = fldvalue.Trim
                                Case "入居前連絡先TEL１"
                                    .Vari_Nyukyomae_tel1 = fldvalue.Trim
                                Case "入居前連絡先TEL２"
                                    .Vari_Nyukyomae_tel2 = fldvalue.Trim
                                Case "入居前連絡先FAX"
                                    .Vari_Nyukyomae_fax = fldvalue.Trim
                                Case "Webアドレス(URL)"
                                    .Vari_Url = fldvalue.Trim
                                Case "業種"
                                    .Vari_Gyosyu = fldvalue.Trim
                                Case "代表者名"
                                    .Vari_Daihyo_name = fldvalue.Trim
                                Case "代表者名SJIS"
                                    .Vari_Daihyo_namesjis = fldvalue.Trim
                                Case "代表者カナ"
                                    .Vari_Daihyo_kana = fldvalue.Trim
                                Case "代表者役職"
                                    .Vari_Daihyo_yakusyoku = fldvalue.Trim
                                Case "代表者を宛先に含める"
                                    .Vari_Daihyo_atenaflg = fldvalue.Trim
                                Case "担当者名"
                                    .Vari_Tanto_name = fldvalue.Trim
                                Case "担当者名SJIS"
                                    .Vari_Tanto_namesjis = fldvalue.Trim
                                Case "担当者カナ"
                                    .Vari_Tanto_kana = fldvalue.Trim
                                Case "担当者部署"
                                    tmp_tantobusyo = fldvalue.Trim
                                Case "担当者役職"
                                    tmp_yakusyoku = fldvalue.Trim
                                Case "担当者を宛先に含める"
                                    .Vari_Tanto_atenaflg = fldvalue.Trim
                                Case "資本金"
                                    .Vari_Sihonkin = EtcMethod.Get_RoundingValue(fldvalue.Trim)
                                Case "従業員数"
                                    .Vari_Jugyosu = fldvalue.Trim
                                Case "記入年月"
                                    .Vari_Nyuryoku_ym = fldvalue.Trim
                                Case "主要取引先"
                                    .Vari_Torihikisaki = fldvalue.Trim
                                Case "連絡先名"
                                    .Vari_Renraku_name = fldvalue.Trim
                                Case "連絡先名SJIS"
                                    .Vari_Renraku_namesjis = fldvalue.Trim
                                Case "連絡先カナ"
                                    .Vari_Renraku_kana = fldvalue.Trim
                                Case "連絡先宛名敬称"
                                    .Vari_Renraku_keisyo = fldvalue.Trim
                                Case "連絡先郵便番号"
                                    .Vari_Renraku_postcode = fldvalue.Trim
                                Case "連絡先住所１"
                                    .Vari_Renraku_addr1 = fldvalue.Trim
                                Case "連絡先住所２"
                                    .Vari_Renraku_addr2 = fldvalue.Trim
                                Case "連絡先TEL１"
                                    .Vari_Renraku_tel1 = fldvalue.Trim
                                Case "連絡先TEL２"
                                    .Vari_Renraku_tel2 = fldvalue.Trim
                                Case "連絡先FAX"
                                    .Vari_Renraku_fax = fldvalue.Trim
                                Case "連絡先携帯１"
                                    .Vari_Renraku_mobiletel1 = fldvalue.Trim
                                Case "連絡先携帯２"
                                    .Vari_Renraku_mobiletel2 = fldvalue.Trim
                                Case "連絡先優先設定(電話番号)"
                                    .Vari_Renraku_yusentelkbn = fldvalue.Trim
                                Case "連絡先間柄"
                                    .Vari_Renraku_aidagara = fldvalue.Trim
                                Case "連絡先備考"
                                    .Vari_Biko_renraku = fldvalue.Trim
                                Case "書類送付先区分"
                                    .Vari_Sofu_kbn = fldvalue.Trim
                                Case "書類送付先名"
                                    .Vari_Sofu_name = fldvalue.Trim
                                Case "書類送付先名SJIS"
                                    .Vari_Sofu_namesjis = fldvalue.Trim
                                Case "書類送付先カナ"
                                    .Vari_Sofu_kana = fldvalue.Trim
                                Case "書類送付先宛名敬称"
                                    .Vari_Sofu_keisyo = fldvalue.Trim
                                Case "書類送付先郵便番号"
                                    .Vari_Sofu_postcode = fldvalue.Trim
                                Case "書類送付先住所１"
                                    .Vari_Sofu_addr1 = fldvalue.Trim
                                Case "書類送付先住所２"
                                    .Vari_Sofu_addr2 = fldvalue.Trim
                                Case "書類送付先TEL１"
                                    .Vari_Sofu_tel1 = fldvalue.Trim
                                Case "書類送付先TEL２"
                                    .Vari_Sofu_tel2 = fldvalue.Trim
                                Case "書類送付先FAX"
                                    .Vari_Sofu_fax = fldvalue.Trim
                                Case "書類送付先備考"
                                    .Vari_Biko_sofu = fldvalue.Trim
                                Case "振込通知書発行の可否"
                                    .Vari_Sofu_hakkofurikomi = fldvalue.Trim
                                Case "振替通知書発行の可否"
                                    .Vari_Sofu_hakkofurikae = fldvalue.Trim
                                Case "督促状発行の可否"
                                    .Vari_Sofu_hakkotokusoku = fldvalue.Trim
                                Case "口座振込仮想口座使用フラグ"
                                    .Vari_Furikomi_kasouseflg = fldvalue.Trim
                                Case "保証人複数フラグ"
                                    .Vari_Hosyo_multipleflg = fldvalue.Trim
                                Case "コンビニ収納サービス利用有無"
                                    .Vari_Cvpay_umu = fldvalue.Trim
                            End Select

                        Next

                        '担当部署役職を結合して格納
                        .Vari_Tanto_busyo = (tmp_tantobusyo & " " & tmp_yakusyoku).Trim

                        '固定値
                        .Vari_Furikae_hosyoumu = "2"    '京王カスタマイズフィールド
                        .Vari_Useflg = 1
                        .Vari_History = DefHistory
                        .Vari_Rowid = Guid.NewGuid.ToString

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim errstr As String = ""
                        Dim hash_log As New Hashtable                                                           'kakaka ↓親マスタチェック("list_basekeydata")が不要な場合は、わかり易いようにした方がよいかと(ここでは親チェックは不要とわかるように)
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If skipflg = False Then                                                                     'kakaka 京王様ソースでは、途中仕様変更などでNotを使用してましたが、出来るだけTrueFalseで記載下さい(1つ前のコードで取得値のNotをSkipFlgにセットし、そのSkipFlgのNotを条件とする…、、、間違い防止)

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

                Dim tmp_sql As String = " SELECT kys_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "契約者口座情報"

    Public Class Kysdata_koza_Repository

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

                Dim model_cvitem As New Njc.Model.Kysdata_koza_Model            '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                Call Me.Get_RelData_KozaSyubetu()
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
                Dim tblname_base As String = "kysdata"
                Dim tblname As String = "kysdata_koza"
                '20161012 革命10アップデートに伴う修正 -add(最後尾にkoza_yucyoflgを追加)
                '20160519 EXEUpdateに伴う修正 契約者口座情報 -add(最後尾にkoza_printkbnを追加)
                Dim fldnamegrp As String = "kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno," & _
                                           "koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1," & _
                                           "yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango," & _
                                           "biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn," & _
                                           "sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn," & _
                                           "koza_yucyoflg"

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
                                Case "契約者No"
                                    .Vari_Kys_no = fldvalue.Trim
                                Case "契約者口座No"
                                    .Vari_Kys_kozano = fldvalue.Trim
                                    .Vari_Koza_kbn = fldvalue.Trim
                                Case "金融期間No"
                                    .Vari_Kinyu_no = fldvalue.Trim
                                Case "金融機関支店No"
                                    .Vari_Kinyu_tenno = fldvalue.Trim
                                Case "口座種別"
                                    .Vari_Koza_syubetu = fldvalue.Trim
                                Case "口座番号"
                                    .Vari_Koza_bango = fldvalue.Trim
                                Case "口座名義"
                                    .Vari_Koza_meigi = fldvalue.Trim
                                Case "口座名義カナ"
                                    .Vari_Koza_meigikana = fldvalue.Trim
                                Case "ゆうちょ記号１"
                                    .Vari_Yucyokoza_kigo1 = fldvalue.Trim
                                Case "ゆうちょ記号２"
                                    .Vari_Yucyokoza_kigo2 = fldvalue.Trim
                                Case "ゆうちょ口座番号"
                                    .Vari_Yucyokoza_bango = fldvalue.Trim
                                Case "口座振替No"
                                    .Vari_Fkae_no = fldvalue.Trim
                                Case "口座振替手数料"
                                    .Vari_Fkae_tesugak = fldvalue.Trim
                                Case "口座振替契約者番号"
                                    .Vari_Fkae_kysbango = fldvalue.Trim
                                Case "口座備考"
                                    .Vari_Biko_koza = fldvalue.Trim
                                Case "口座振込備考"
                                    .Vari_Biko_furikomi = fldvalue.Trim
                                Case "総合振込区分"
                                    .Vari_Sgfirai_kbn = fldvalue.Trim
                                Case "振込依頼人No"
                                    .Vari_Sgfirai_no = fldvalue.Trim
                                Case "振込手数料負担区分"
                                    .Vari_Sgfirai_tesufutankbn = fldvalue.Trim
                                Case "振込手数料計算区分"
                                    .Vari_Sgfirai_tesukeisankbn = fldvalue.Trim
                                Case "振込手数料固定額1"
                                    .Vari_Sgfirai_tesukotei1gak = fldvalue.Trim
                                Case "振込手数料固定額2"
                                    .Vari_Sgfirai_tesukotei2gak = fldvalue.Trim
                                Case "備考(総合振込情報)"
                                    .Vari_Biko_sgfirai = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約者口座情報 -add sta
                                Case "印刷口座区分"
                                    .Vari_Koza_printkbn = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 契約者口座情報 -add end
                                Case "金融機関区分(ゆうちょ銀行フラグ)"
                                    .Vari_Koza_yucyoflg = fldvalue.Trim
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

                '20170116 契約者口座情報ダミーレコード作成処理の追加 -add sta
                If CNVNO = ConvertTypes._汎用 Then
                    Try
                        Call Me.Set_KysDummyKoza(sqlcnnv10)
                    Catch ex As Exception
                        '----- ログ出力 -----
                        Dim tmptmpcnt As Integer = 0
                        Dim tmptmpstr As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(41, sheetname, "データ調整", "データ調整中にエラーが発生しました。"), False)
                        DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, tmptmpcnt)
                    End Try
                End If
                '20170116 契約者口座情報ダミーレコード作成処理の追加 -add end

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
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT kys_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,kys_no) + '-' + CONVERT(varchar,kys_kozano) FROM " & tblname
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

            ''' <summary>
            ''' 契約者口座情報へダミーレコードを作成する
            ''' </summary>
            ''' <param name="sqlcnnv10">賃貸10DB接続用オブジェクト</param>
            ''' <remarks>
            ''' 20170116 契約者口座情報ダミーレコード作成処理の追加 新規追加
            ''' 　詳細はメソッド内へ記載
            ''' </remarks>
            Public Sub Set_KysDummyKoza(ByVal sqlcnnv10 As SqlConnection)

                '****************************************************************************
                '　中間ファイルへの登録作業上、契約者基本情報と契約者口座情報が別になっているため
                '　契約者情報は存在するが契約者口座情報は存在しない状態が発生する
                '　画面上では契約者情報を作成した段階で契約者基本情報(kysdata)と契約者口座情報(kysdata_koza)が
                '　3レコード自動で作成されるため整合性が取れない状態が発生する
                '****************************************************************************

                Dim tmp_sql As String = ""
                Dim tmpcnt As Integer = 0

                '①契約者口座No = 1 のデータを作成する
                tmp_sql = tmp_sql & " INSERT INTO kysdata_koza "
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 kys_no "
                tmp_sql = tmp_sql & " 	,1 AS kys_kozano "
                tmp_sql = tmp_sql & " 	,1 AS koza_kbn "
                tmp_sql = tmp_sql & " 	,NULL AS kinyu_no "
                tmp_sql = tmp_sql & " 	,NULL AS kinyu_tenno "
                tmp_sql = tmp_sql & " 	,1 AS koza_syubetu "
                tmp_sql = tmp_sql & " 	,NULL AS koza_bango "
                tmp_sql = tmp_sql & " 	,kys_name AS koza_meigi "
                tmp_sql = tmp_sql & " 	,kys_kana AS koza_meigikana "
                tmp_sql = tmp_sql & " 	,'' AS yucyokoza_kigo1 "
                tmp_sql = tmp_sql & " 	,'' AS yucyokoza_kigo2 "
                tmp_sql = tmp_sql & " 	,'' AS yucyokoza_bango "
                tmp_sql = tmp_sql & " 	,NULL AS fkae_no "
                tmp_sql = tmp_sql & " 	,NULL AS fkae_tesugak "
                tmp_sql = tmp_sql & " 	,'' AS fkae_kysbango "
                tmp_sql = tmp_sql & " 	,'' AS biko_koza "
                tmp_sql = tmp_sql & " 	,'' AS biko_furikomi "
                tmp_sql = tmp_sql & " 	,1 AS sgfirai_kbn "
                tmp_sql = tmp_sql & " 	,NULL AS sgfirai_no "
                tmp_sql = tmp_sql & " 	,2 AS sgfirai_tesufutankbn "
                tmp_sql = tmp_sql & " 	,1 AS sgfirai_tesukeisankbn "
                tmp_sql = tmp_sql & " 	,NULL AS sgfirai_tesukotei1gak "
                tmp_sql = tmp_sql & " 	,NULL AS sgfirai_tesukotei2gak "
                tmp_sql = tmp_sql & " 	,'' AS biko_sgfirai "
                tmp_sql = tmp_sql & " 	,1 AS koza_printkbn "
                tmp_sql = tmp_sql & " 	,0 AS koza_yucyoflg "
                tmp_sql = tmp_sql & " FROM kysdata WHERE kys_no NOT IN (SELECT kys_no FROM kysdata_koza WHERE kys_kozano = 1) "
                tmp_sql = tmp_sql & " UNION "
                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 kys_no "
                tmp_sql = tmp_sql & " 	,1 AS kys_kozano "
                tmp_sql = tmp_sql & " 	,1 AS koza_kbn "
                tmp_sql = tmp_sql & " 	,NULL AS kinyu_no "
                tmp_sql = tmp_sql & " 	,NULL AS kinyu_tenno "
                tmp_sql = tmp_sql & " 	,1 AS koza_syubetu "
                tmp_sql = tmp_sql & " 	,NULL AS koza_bango "
                tmp_sql = tmp_sql & " 	,kys_name AS koza_meigi "
                tmp_sql = tmp_sql & " 	,kys_kana AS koza_meigikana "
                tmp_sql = tmp_sql & " 	,'' AS yucyokoza_kigo1 "
                tmp_sql = tmp_sql & " 	,'' AS yucyokoza_kigo2 "
                tmp_sql = tmp_sql & " 	,'' AS yucyokoza_bango "
                tmp_sql = tmp_sql & " 	,NULL AS fkae_no "
                tmp_sql = tmp_sql & " 	,NULL AS fkae_tesugak "
                tmp_sql = tmp_sql & " 	,'' AS fkae_kysbango "
                tmp_sql = tmp_sql & " 	,'' AS biko_koza "
                tmp_sql = tmp_sql & " 	,'' AS biko_furikomi "
                tmp_sql = tmp_sql & " 	,1 AS sgfirai_kbn "
                tmp_sql = tmp_sql & " 	,NULL AS sgfirai_no "
                tmp_sql = tmp_sql & " 	,2 AS sgfirai_tesufutankbn "
                tmp_sql = tmp_sql & " 	,1 AS sgfirai_tesukeisankbn "
                tmp_sql = tmp_sql & " 	,NULL AS sgfirai_tesukotei1gak "
                tmp_sql = tmp_sql & " 	,NULL AS sgfirai_tesukotei2gak "
                tmp_sql = tmp_sql & " 	,'' AS biko_sgfirai "
                tmp_sql = tmp_sql & " 	,1 AS koza_printkbn "
                tmp_sql = tmp_sql & " 	,0 AS koza_yucyoflg "
                tmp_sql = tmp_sql & " FROM kysdata WHERE kys_no NOT IN (SELECT kys_no FROM kysdata_koza) "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmpcnt)
                tmp_sql = ""

                '②契約者口座No = 2、3 のデータを作成する
                '※過去データ調整で対応したクエリを一部修正して流用する
                tmp_sql = tmp_sql & " BEGIN "
                tmp_sql = tmp_sql & " 	/*1.変数宣言*/ "
                tmp_sql = tmp_sql & " 	DECLARE @kysno INT				/*kys_no*/ "
                tmp_sql = tmp_sql & " 	DECLARE @maxkozano INT			/*[最大口座No]*/ "
                tmp_sql = tmp_sql & " 	DECLARE @maxkozacnt INT			/*[最大口座数]*/ "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 	/*2.カーソル宣言*/ "
                tmp_sql = tmp_sql & " 	DECLARE KysInfo CURSOR FOR "
                tmp_sql = tmp_sql & " 		/*最大値抽出クエリ文 sta*/ "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 VW.kys_no "
                tmp_sql = tmp_sql & " 			,[最大口座No] "
                tmp_sql = tmp_sql & " 			,[最大口座数] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 kys_no "
                tmp_sql = tmp_sql & " 				,MAX(kys_kozano) AS [最大口座No] "
                tmp_sql = tmp_sql & " 			FROM kysdata_koza "
                tmp_sql = tmp_sql & " 			GROUP BY kys_no "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		LEFT JOIN "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT "
                tmp_sql = tmp_sql & " 					 kys_no "
                tmp_sql = tmp_sql & " 					/*,kys_kozano*/ "
                tmp_sql = tmp_sql & " 					,MAX([口座数抽出用]) AS [最大口座数] "
                tmp_sql = tmp_sql & " 				FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 kys_no "
                tmp_sql = tmp_sql & " 						/*,kys_kozano*/ "
                tmp_sql = tmp_sql & " 						,ROW_NUMBER()OVER(PARTITION BY kys_no ORDER BY kys_kozano) AS [口座数抽出用] "
                tmp_sql = tmp_sql & " 					FROM kysdata_koza "
                tmp_sql = tmp_sql & " 				) AS VW1 "
                tmp_sql = tmp_sql & " 				GROUP BY kys_no "
                tmp_sql = tmp_sql & " 			) AS VW3 "
                tmp_sql = tmp_sql & " 		ON VW.kys_no = VW3.kys_no "
                tmp_sql = tmp_sql & " 		ORDER BY kys_no "
                tmp_sql = tmp_sql & " 		/*最大値抽出クエリ文 end*/ "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 	/*3.初期化関連処理*/ "
                tmp_sql = tmp_sql & " 	SET NOCOUNT ON				/*コメントON*/ "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 	/*4.メイン処理*/ "
                tmp_sql = tmp_sql & " 	OPEN KysInfo "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 		FETCH NEXT FROM KysInfo "
                tmp_sql = tmp_sql & " 		INTO @kysno,@maxkozano,@maxkozacnt "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 		/*INSERT処理*/ "
                tmp_sql = tmp_sql & " 		WHILE @@FETCH_STATUS = 0 "
                tmp_sql = tmp_sql & " 		BEGIN "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 			/*最大口座No = 1 かつ 最大口座数 = 1 の場合は口座No = 2、3を作成*/ "
                tmp_sql = tmp_sql & " 			IF @maxkozano = 1 AND @maxkozacnt = 1 "
                tmp_sql = tmp_sql & " 				BEGIN "
                tmp_sql = tmp_sql & " 					INSERT INTO kysdata_koza (kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno,koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn,koza_yucyoflg) VALUES (@kysno,2,2,NULL,NULL,1,NULL,'','','','','',NULL,0.00,'','','',1,NULL,2,1,NULL,NULL,'',1,0) "
                tmp_sql = tmp_sql & " 					INSERT INTO kysdata_koza (kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno,koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn,koza_yucyoflg) VALUES (@kysno,3,3,NULL,NULL,1,NULL,'','','','','',NULL,0.00,'','','',NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,0) "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & " 			/*最大口座No = 2 かつ 最大口座数 = 2 の場合は口座No = 3を作成*/ "
                tmp_sql = tmp_sql & " 			IF @maxkozano = 2 AND @maxkozacnt = 2 "
                tmp_sql = tmp_sql & " 				BEGIN "
                tmp_sql = tmp_sql & " 					INSERT INTO kysdata_koza (kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno,koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn,koza_yucyoflg) VALUES (@kysno,3,3,NULL,NULL,1,NULL,'','','','','',NULL,0.00,'','','',NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,0) "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & " 			/*最大口座No = 3 かつ 最大口座数 = 2 の場合は口座No = 2を作成*/ "
                tmp_sql = tmp_sql & " 			IF @maxkozano = 3 AND @maxkozacnt = 2 "
                tmp_sql = tmp_sql & " 				BEGIN "
                tmp_sql = tmp_sql & " 					INSERT INTO kysdata_koza (kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno,koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn,koza_yucyoflg) VALUES (@kysno,2,2,NULL,NULL,1,NULL,'','','','','',NULL,0.00,'','','',1,NULL,2,1,NULL,NULL,'',1,0) "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 			FETCH NEXT FROM KysInfo "
                tmp_sql = tmp_sql & " 			INTO @kysno,@maxkozano,@maxkozacnt "
                tmp_sql = tmp_sql & " 		END "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 	CLOSE KysInfo "
                tmp_sql = tmp_sql & " 	 "
                tmp_sql = tmp_sql & " 	/*5.解放*/ "
                tmp_sql = tmp_sql & " 	DEALLOCATE KysInfo "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " END "
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmpcnt)
                tmp_sql = ""

            End Sub

        End Class

    End Class

#End Region

#Region "契約者照合用カナ情報"

    Public Class M_fb_fkomsyogo_kys_Repository

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

                Dim model_cvitem As New Njc.Model.M_fb_fkomsyogo_Model          '移行値格納用モデル初期化
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
                Dim tblname_base As String = "kysdata"
                Dim tblname As String = "m_fb_fkomsyogo"
                Dim fldnamegrp As String = "sqsaki_kbn,sqsaki_no,sqsaki_kozano,syogo_no,syogo_priority," & _
                                           "fkom_syogomoji"

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

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol - 1)).Trim
                        .Vari_Sqsaki_no = tmp_keymain

                        '照合用カナカウント初期化
                        Dim cvitemcnt As Integer = 0
                        Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

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
                            .Vari_Fkom_syogomoji = fldvalue

                            '固定値
                            .Vari_Sqsaki_kbn = 100
                            .Vari_Sqsaki_kozano = 1
                            '20161125 契約者照合用カナ情報の移行処理修正 -chg sta
                            '.Vari_Syogo_no = (cntjj - 1).ToString
                            .Vari_Syogo_no = (cntjj).ToString
                            '20161125 契約者照合用カナ情報の移行処理修正 -chg end
                            .Vari_Syogo_priority = 1

                            '照合用カナデータが存在する場合に書込処理を行う
                            If .Vari_Fkom_syogomoji <> "" Then

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
                '20160928 メモ関連の移行件数表示修正 -chg sta
                'midrowcnt = rowcnt
                midrowcnt = totalrowcnt
                '20160928 メモ関連の移行件数表示修正 -chg end
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
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT kys_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT sqsaki_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "契約者保証人情報"

    Public Class Kysdata_hosyo_Repository

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

                Dim model_cvitem As New Njc.Model.Kysdata_hosyo_Model           '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列

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
                Dim tblname_base As String = "kysdata"
                Dim tblname As String = "kysdata_hosyo"
                Dim fldnamegrp As String = "kys_no,hosyo_no,hosyo_name,hosyo_namesjis,hosyo_kana," & _
                                           "keisyo,post_code,addr1,addr2,tel1," & _
                                           "tel2,fax,mobiletel1,mobiletel2,yusentel_kbn," & _
                                           "birthday,nensyu,biko_hosyo,aidagara,kinmu_name," & _
                                           "kinmu_namesjis,kinmu_kana,kinmu_postcode,kinmu_addr1,kinmu_addr2," & _
                                           "kinmu_tel1,kinmu_tel2,kinmu_fax,kinmu_gyosyu,kinmu_busyo," & _
                                           "kinmu_nyuryokuym,kinmu_nyusyaym,kinmu_taisyaym,biko_kinmu"

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
                                Case "契約者No"
                                    .Vari_Kys_no = fldvalue.Trim
                                Case "保証人No"
                                    .Vari_Hosyo_no = fldvalue.Trim
                                Case "保証人名"
                                    .Vari_Hosyo_name = fldvalue.Trim
                                Case "保証人名SJIS"
                                    .Vari_Hosyo_namesjis = fldvalue.Trim
                                Case "保証人カナ"
                                    .Vari_Hosyo_kana = fldvalue.Trim
                                Case "宛名敬称"
                                    .Vari_Keisyo = fldvalue.Trim
                                Case "郵便番号"
                                    .Vari_Post_code = fldvalue.Trim
                                Case "住所１"
                                    .Vari_Addr1 = fldvalue.Trim
                                Case "住所２"
                                    .Vari_Addr2 = fldvalue.Trim
                                Case "TEL１"
                                    .Vari_Tel1 = fldvalue.Trim
                                Case "TEL２"
                                    .Vari_Tel2 = fldvalue.Trim
                                Case "FAX"
                                    .Vari_Fax = fldvalue.Trim
                                Case "携帯１"
                                    .Vari_Mobiletel1 = fldvalue.Trim
                                Case "携帯２"
                                    .Vari_Mobiletel2 = fldvalue.Trim
                                Case "優先設定(電話番号)"
                                    .Vari_Yusentel_kbn = fldvalue.Trim
                                Case "誕生日・設立日"
                                    .Vari_Birthday = fldvalue.Trim
                                Case "年収・年商"
                                    '2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化 -chg sta
                                    '↓↓↓旧srcコメントアウト↓↓↓
                                    ''2016.03.31 変換時の型落ち対応 -chg sta
                                    ''Dim tmpnensyustr As String = fldvalue.Trim
                                    ''Dim tmpnensyudbl As Double = Double.Parse(tmpnensyustr)                 'kakaka4 0322_1300 変換時の型落ちを考慮。
                                    ' ''四捨五入して取得
                                    ''.Vari_Nensyu = Math.Round(tmpnensyudbl, MidpointRounding.AwayFromZero).ToString
                                    'Dim tmpnensyustr As String = fldvalue.Trim
                                    'Dim tmpnensyudbl As Double = 0
                                    'If Double.TryParse(tmpnensyustr, tmpnensyudbl) Then
                                    '    '四捨五入して取得
                                    '    .Vari_Nensyu = Math.Round(tmpnensyudbl, MidpointRounding.AwayFromZero).ToString
                                    'Else
                                    '    '変換できない場合は値をそのまま格納 (ログ出力させる)
                                    '    .Vari_Nensyu = tmpnensyustr
                                    'End If
                                    ''2016.03.31 変換時の型落ち対応 -chg end
                                    '↑↑↑旧srcコメントアウト↑↑↑
                                    .Vari_Nensyu = EtcMethod.Get_RoundingValue(fldvalue.Trim)
                                    '2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化 -chg end
                                Case "備考"
                                    .Vari_Biko_hosyo = fldvalue.Trim
                                Case "間柄"
                                    .Vari_Aidagara = fldvalue.Trim
                                Case "勤務先名"
                                    .Vari_Kinmu_name = fldvalue.Trim
                                Case "勤務先名SJIS"
                                    .Vari_Kinmu_namesjis = fldvalue.Trim
                                Case "勤務先カナ"
                                    .Vari_Kinmu_kana = fldvalue.Trim
                                Case "勤務先郵便番号"
                                    .Vari_Kinmu_postcode = fldvalue.Trim
                                Case "勤務先住所１"
                                    .Vari_Kinmu_addr1 = fldvalue.Trim
                                Case "勤務先住所２"
                                    .Vari_Kinmu_addr2 = fldvalue.Trim
                                Case "勤務先TEL１"
                                    .Vari_Kinmu_tel1 = fldvalue.Trim
                                Case "勤務先TEL２"
                                    .Vari_Kinmu_tel2 = fldvalue.Trim
                                Case "勤務先FAX"
                                    .Vari_Kinmu_fax = fldvalue.Trim
                                Case "勤務先業種"
                                    .Vari_Kinmu_gyosyu = fldvalue.Trim
                                Case "勤務先部署"
                                    .Vari_Kinmu_busyo = fldvalue.Trim
                                Case "勤務先情報記入年月"
                                    .Vari_Kinmu_nyuryokuym = fldvalue.Trim
                                Case "勤務先入社年月"
                                    .Vari_Kinmu_nyusyaym = fldvalue.Trim
                                Case "勤務先退社年月"
                                    .Vari_Kinmu_taisyaym = fldvalue.Trim
                                Case "勤務先備考"
                                    .Vari_Biko_kinmu = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_Keisyo = "様"
                        '20160530 ログ修正 -del
                        '.Vari_Yusentel_kbn = -1

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
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT kys_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,kys_no) + '-' + CONVERT(varchar,hosyo_no) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "契約者メモ情報"

    Public Class Kysdata_memo_Repository

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

                Dim model_cvitem As New Njc.Model.Kysdata_memo_Model            '移行値格納用モデル初期化
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
                Dim tblname_base As String = "kysdata"
                Dim tblname As String = "kysdata_memo"
                Dim fldnamegrp As String = "kys_no,memo_no,memo,history"

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
                        .Vari_Kys_no = tmp_keymain

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
                                Dim log_key As String = "kysdata_memo-kys_no"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(tmp_keymain) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "kysdata_memo-kys_no"
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

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT kys_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT kys_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

End Namespace


