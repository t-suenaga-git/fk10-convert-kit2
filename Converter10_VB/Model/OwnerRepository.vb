Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "家主情報"

    Public Class Owdata_Repository

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

                Dim model_cvitem As New Njc.Model.Owdata_Model                  '移行値格納用モデル初期化
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
                Dim tblname As String = "owdata"
                '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add (kaikei_jouhou,addr1,addr2を追加 ※kaikei_jouhouがもともと入っていなかったため同時に追加する)
                Dim fldnamegrp As String = "ow_no,kojinhojin_flg,ow_name,ow_namesjis,ow_kana," & _
                                           "keisyo,post_code,addr_kenno,addr_sino,addr_cyo," & _
                                           "addr_cyome,addr_cyomeptn,addr_banti,addr_etc,toukiaddr1," & _
                                           "toukiaddr2,tel1,tel2,fax,mobiletel1," & _
                                           "mobiletel2,mail,mobilemail,yusentel_kbn,yusenmail_kbn," & _
                                           "birthday,nensyu,biko_kihon,gender,honseki," & _
                                           "kinmu_name,kinmu_namesjis,kinmu_kana,kinmu_postcode,kinmu_addr1," & _
                                           "kinmu_addr2,kinmu_tel1,kinmu_tel2,kinmu_fax,kinmu_gyosyu," & _
                                           "kinmu_busyo,kinmu_nyuryokuym,kinmu_nyusyaym,url,gyosyu," & _
                                           "daihyo_name,daihyo_namesjis,daihyo_kana,daihyo_yakusyoku,daihyo_atenaflg," & _
                                           "tanto_name,tanto_namesjis,tanto_kana,tanto_busyo,tanto_atenaflg," & _
                                           "sihonkin,jugyosu,nyuryoku_ym,torihikisaki,renraku_name," & _
                                           "renraku_namesjis,renraku_kana,renraku_keisyo,renraku_postcode,renraku_addr1," & _
                                           "renraku_addr2,renraku_tel1,renraku_tel2,renraku_fax,renraku_mobiletel1," & _
                                           "renraku_mobiletel2,renraku_yusentelkbn,renraku_aidagara,biko_renraku,sofu_kbn," & _
                                           "sofu_name,sofu_namesjis,sofu_kana,sofu_keisyo,sofu_postcode," & _
                                           "sofu_addr1,sofu_addr2,sofu_tel1,sofu_tel2,sofu_fax," & _
                                           "biko_sofu,event_tantono,useflg,findkeyword,history," & _
                                           "rowid,kaikei_jouhou,addr1,addr2"

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
                                Case "家主No"
                                    .Vari_Ow_no = fldvalue.Trim
                                Case "個人法人フラグ"
                                    .Vari_Kojinhojin_flg = fldvalue.Trim
                                Case "家主名"
                                    .Vari_Ow_name = fldvalue.Trim
                                Case "家主名(SJIS)"
                                    .Vari_Ow_namesjis = fldvalue.Trim
                                Case "家主カナ"
                                    .Vari_Ow_kana = fldvalue.Trim
                                Case "宛名敬称"
                                    .Vari_Keisyo = fldvalue.Trim
                                Case "郵便番号"
                                    .Vari_Post_code = fldvalue.Trim
                                    '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -chg sta
                                    'Case "住所1"
                                    '    tmp_address1 = fldvalue.Trim
                                    'Case "住所2"
                                    '    tmp_address2 = fldvalue.Trim
                                Case "住所1"
                                    .Vari_Addr1 = fldvalue.Trim
                                Case "住所2"
                                    .Vari_Addr2 = fldvalue.Trim
                                    '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -chg end
                                Case "登記住所１"
                                    .Vari_Toukiaddr1 = fldvalue.Trim
                                Case "登記住所２"
                                    .Vari_Toukiaddr2 = fldvalue.Trim
                                Case "TEL1"
                                    .Vari_Tel1 = fldvalue.Trim
                                Case "TEL2"
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
                                Case "優先電話設定"
                                    .Vari_Yusentel_kbn = fldvalue.Trim
                                Case "優先メール設定"
                                    .Vari_Yusenmail_kbn = fldvalue.Trim
                                Case "誕生日・設立日"
                                    .Vari_Birthday = fldvalue.Trim
                                Case "年収・年商"
                                    .Vari_Nensyu = fldvalue.Trim
                                Case "備考"
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
                                Case "URL"
                                    .Vari_Url = fldvalue.Trim
                                Case "業種"
                                    .Vari_Gyosyu = fldvalue.Trim
                                Case "代表者名"
                                    .Vari_Daihyo_name = fldvalue.Trim
                                Case "代表者名(SJIS)"
                                    .Vari_Daihyo_namesjis = fldvalue.Trim
                                Case "代表者カナ"
                                    .Vari_Daihyo_kana = fldvalue.Trim
                                Case "代表者役職"
                                    .Vari_Daihyo_yakusyoku = fldvalue.Trim
                                Case "代表者を宛先に含める"
                                    .Vari_Daihyo_atenaflg = fldvalue.Trim
                                Case "担当者名"
                                    .Vari_Tanto_name = fldvalue.Trim
                                Case "担当者名(SJIS)"
                                    .Vari_Tanto_namesjis = fldvalue.Trim
                                Case "担当者カナ"
                                    .Vari_Tanto_kana = fldvalue.Trim
                                Case "担当者部署"
                                    .Vari_Tanto_busyo = fldvalue.Trim
                                Case "担当者を宛先に含める"
                                    .Vari_Tanto_atenaflg = fldvalue.Trim
                                Case "資本金"
                                    .Vari_Sihonkin = fldvalue.Trim
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
                                Case "送付先名"
                                    .Vari_Sofu_name = fldvalue.Trim
                                Case "送付先名(SJIS)"
                                    .Vari_Sofu_namesjis = fldvalue.Trim
                                Case "送付先カナ"
                                    .Vari_Sofu_kana = fldvalue.Trim
                                Case "送付先宛名敬称"
                                    .Vari_Sofu_keisyo = fldvalue.Trim
                                Case "送付先郵便番号"
                                    .Vari_Sofu_postcode = fldvalue.Trim
                                Case "送付先住所１"
                                    .Vari_Sofu_addr1 = fldvalue.Trim
                                Case "送付先住所２"
                                    .Vari_Sofu_addr2 = fldvalue.Trim
                                Case "送付先TEL１"
                                    .Vari_Sofu_tel1 = fldvalue.Trim
                                Case "送付先TEL2"
                                    .Vari_Sofu_tel2 = fldvalue.Trim
                                Case "送付先FAX"
                                    .Vari_Sofu_fax = fldvalue.Trim
                                Case "送付先備考"
                                    .Vari_Biko_sofu = fldvalue.Trim
                                Case "イベント担当者No"
                                    .Vari_Event_tantono = fldvalue.Trim
                                Case "絞り込みキーワード"
                                    .Vari_Findkeyword = fldvalue.Trim
                                Case "会計情報"
                                    .Vari_Kaikei_jouhou = fldvalue.Trim
                            End Select

                        Next

                        '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del sta
                        ''住所変換/格納
                        'Dim tmp_address As String = tmp_address1 & " " & tmp_address2
                        'Dim hash_address As New Hashtable
                        'hash_address = AddressChange.Get_Address(tmp_address, sqlcnnv10)
                        '.Vari_Addr_kenno = hash_address("ken_no")
                        '.Vari_Addr_sino = hash_address("si_no")
                        '.Vari_Addr_cyo = hash_address("mati")
                        '.Vari_Addr_cyome = hash_address("cyome")
                        '.Vari_Addr_cyomeptn = hash_address("chomeptn")
                        '.Vari_Addr_banti = hash_address("banti")
                        '.Vari_Addr_etc = hash_address("etc")
                        '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del end

                        '20160706 家主個人法人区分の取得方法変更 -add sta
                        '個人法人区分が個人に設定されている場合は名称から判別する(法人に設定されているデータはそのまま移行)
                        If .Vari_Kojinhojin_flg = "1" Then
                            .Vari_Kojinhojin_flg = Me.Get_KojinHojinvalue(.Vari_Ow_name)
                        End If
                        '20160706 家主個人法人区分の取得方法変更 -add end

                        '固定値
                        .Vari_Useflg = 1
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

            '    Dim model_cvitem As New Njc.Model.Owdata_Model                  '移行値格納用モデル初期化
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
            '    Dim tblname As String = "owdata"
            '    '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add (kaikei_jouhou,addr1,addr2を追加 ※kaikei_jouhouがもともと入っていなかったため同時に追加する)
            '    Dim fldnamegrp As String = "ow_no,kojinhojin_flg,ow_name,ow_namesjis,ow_kana," & _
            '                               "keisyo,post_code,addr_kenno,addr_sino,addr_cyo," & _
            '                               "addr_cyome,addr_cyomeptn,addr_banti,addr_etc,toukiaddr1," & _
            '                               "toukiaddr2,tel1,tel2,fax,mobiletel1," & _
            '                               "mobiletel2,mail,mobilemail,yusentel_kbn,yusenmail_kbn," & _
            '                               "birthday,nensyu,biko_kihon,gender,honseki," & _
            '                               "kinmu_name,kinmu_namesjis,kinmu_kana,kinmu_postcode,kinmu_addr1," & _
            '                               "kinmu_addr2,kinmu_tel1,kinmu_tel2,kinmu_fax,kinmu_gyosyu," & _
            '                               "kinmu_busyo,kinmu_nyuryokuym,kinmu_nyusyaym,url,gyosyu," & _
            '                               "daihyo_name,daihyo_namesjis,daihyo_kana,daihyo_yakusyoku,daihyo_atenaflg," & _
            '                               "tanto_name,tanto_namesjis,tanto_kana,tanto_busyo,tanto_atenaflg," & _
            '                               "sihonkin,jugyosu,nyuryoku_ym,torihikisaki,renraku_name," & _
            '                               "renraku_namesjis,renraku_kana,renraku_keisyo,renraku_postcode,renraku_addr1," & _
            '                               "renraku_addr2,renraku_tel1,renraku_tel2,renraku_fax,renraku_mobiletel1," & _
            '                               "renraku_mobiletel2,renraku_yusentelkbn,renraku_aidagara,biko_renraku,sofu_kbn," & _
            '                               "sofu_name,sofu_namesjis,sofu_kana,sofu_keisyo,sofu_postcode," & _
            '                               "sofu_addr1,sofu_addr2,sofu_tel1,sofu_tel2,sofu_fax," & _
            '                               "biko_sofu,event_tantono,useflg,findkeyword,history," & _
            '                               "rowid,kaikei_jouhou,addr1,addr2"

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

            '            '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del sta
            '            ''作業用変数
            '            'Dim tmp_address1 As String = ""
            '            'Dim tmp_address2 As String = ""
            '            '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del end

            '            '移行値取得
            '            For cntjj = 1 To columncnt

            '                Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            '                Dim fldvalue As String = ""
            '                If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            '                End If

            '                Select Case fldname
            '                    Case "家主No"
            '                        .Vari_Ow_no = fldvalue.Trim
            '                    Case "個人法人フラグ"
            '                        .Vari_Kojinhojin_flg = fldvalue.Trim
            '                    Case "家主名"
            '                        .Vari_Ow_name = fldvalue.Trim
            '                    Case "家主名(SJIS)"
            '                        .Vari_Ow_namesjis = fldvalue.Trim
            '                    Case "家主カナ"
            '                        .Vari_Ow_kana = fldvalue.Trim
            '                    Case "宛名敬称"
            '                        .Vari_Keisyo = fldvalue.Trim
            '                    Case "郵便番号"
            '                        .Vari_Post_code = fldvalue.Trim
            '                        '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -chg sta
            '                        'Case "住所1"
            '                        '    tmp_address1 = fldvalue.Trim
            '                        'Case "住所2"
            '                        '    tmp_address2 = fldvalue.Trim
            '                    Case "住所1"
            '                        .Vari_Addr1 = fldvalue.Trim
            '                    Case "住所2"
            '                        .Vari_Addr2 = fldvalue.Trim
            '                        '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -chg end
            '                    Case "登記住所１"
            '                        .Vari_Toukiaddr1 = fldvalue.Trim
            '                    Case "登記住所２"
            '                        .Vari_Toukiaddr2 = fldvalue.Trim
            '                    Case "TEL1"
            '                        .Vari_Tel1 = fldvalue.Trim
            '                    Case "TEL2"
            '                        .Vari_Tel2 = fldvalue.Trim
            '                    Case "FAX"
            '                        .Vari_Fax = fldvalue.Trim
            '                    Case "携帯１"
            '                        .Vari_Mobiletel1 = fldvalue.Trim
            '                    Case "携帯２"
            '                        .Vari_Mobiletel2 = fldvalue.Trim
            '                    Case "メールアドレス"
            '                        .Vari_Mail = fldvalue.Trim
            '                    Case "携帯メールアドレス"
            '                        .Vari_Mobilemail = fldvalue.Trim
            '                    Case "優先電話設定"
            '                        .Vari_Yusentel_kbn = fldvalue.Trim
            '                    Case "優先メール設定"
            '                        .Vari_Yusenmail_kbn = fldvalue.Trim
            '                    Case "誕生日・設立日"
            '                        .Vari_Birthday = fldvalue.Trim
            '                    Case "年収・年商"
            '                        .Vari_Nensyu = fldvalue.Trim
            '                    Case "備考"
            '                        .Vari_Biko_kihon = fldvalue.Trim
            '                    Case "性別"
            '                        .Vari_Gender = fldvalue.Trim
            '                    Case "本籍地"
            '                        .Vari_Honseki = fldvalue.Trim
            '                    Case "勤務先名"
            '                        .Vari_Kinmu_name = fldvalue.Trim
            '                    Case "勤務先名SJIS"
            '                        .Vari_Kinmu_namesjis = fldvalue.Trim
            '                    Case "勤務先カナ"
            '                        .Vari_Kinmu_kana = fldvalue.Trim
            '                    Case "勤務先郵便番号"
            '                        .Vari_Kinmu_postcode = fldvalue.Trim
            '                    Case "勤務先住所１"
            '                        .Vari_Kinmu_addr1 = fldvalue.Trim
            '                    Case "勤務先住所２"
            '                        .Vari_Kinmu_addr2 = fldvalue.Trim
            '                    Case "勤務先TEL１"
            '                        .Vari_Kinmu_tel1 = fldvalue.Trim
            '                    Case "勤務先TEL２"
            '                        .Vari_Kinmu_tel2 = fldvalue.Trim
            '                    Case "勤務先FAX"
            '                        .Vari_Kinmu_fax = fldvalue.Trim
            '                    Case "勤務先業種"
            '                        .Vari_Kinmu_gyosyu = fldvalue.Trim
            '                    Case "勤務先部署"
            '                        .Vari_Kinmu_busyo = fldvalue.Trim
            '                    Case "勤務先情報記入年月"
            '                        .Vari_Kinmu_nyuryokuym = fldvalue.Trim
            '                    Case "勤務先入社年月"
            '                        .Vari_Kinmu_nyusyaym = fldvalue.Trim
            '                    Case "URL"
            '                        .Vari_Url = fldvalue.Trim
            '                    Case "業種"
            '                        .Vari_Gyosyu = fldvalue.Trim
            '                    Case "代表者名"
            '                        .Vari_Daihyo_name = fldvalue.Trim
            '                    Case "代表者名(SJIS)"
            '                        .Vari_Daihyo_namesjis = fldvalue.Trim
            '                    Case "代表者カナ"
            '                        .Vari_Daihyo_kana = fldvalue.Trim
            '                    Case "代表者役職"
            '                        .Vari_Daihyo_yakusyoku = fldvalue.Trim
            '                    Case "代表者を宛先に含める"
            '                        .Vari_Daihyo_atenaflg = fldvalue.Trim
            '                    Case "担当者名"
            '                        .Vari_Tanto_name = fldvalue.Trim
            '                    Case "担当者名(SJIS)"
            '                        .Vari_Tanto_namesjis = fldvalue.Trim
            '                    Case "担当者カナ"
            '                        .Vari_Tanto_kana = fldvalue.Trim
            '                    Case "担当者部署"
            '                        .Vari_Tanto_busyo = fldvalue.Trim
            '                    Case "担当者を宛先に含める"
            '                        .Vari_Tanto_atenaflg = fldvalue.Trim
            '                    Case "資本金"
            '                        .Vari_Sihonkin = fldvalue.Trim
            '                    Case "従業員数"
            '                        .Vari_Jugyosu = fldvalue.Trim
            '                    Case "記入年月"
            '                        .Vari_Nyuryoku_ym = fldvalue.Trim
            '                    Case "主要取引先"
            '                        .Vari_Torihikisaki = fldvalue.Trim
            '                    Case "連絡先名"
            '                        .Vari_Renraku_name = fldvalue.Trim
            '                    Case "連絡先名SJIS"
            '                        .Vari_Renraku_namesjis = fldvalue.Trim
            '                    Case "連絡先カナ"
            '                        .Vari_Renraku_kana = fldvalue.Trim
            '                    Case "連絡先宛名敬称"
            '                        .Vari_Renraku_keisyo = fldvalue.Trim
            '                    Case "連絡先郵便番号"
            '                        .Vari_Renraku_postcode = fldvalue.Trim
            '                    Case "連絡先住所１"
            '                        .Vari_Renraku_addr1 = fldvalue.Trim
            '                    Case "連絡先住所２"
            '                        .Vari_Renraku_addr2 = fldvalue.Trim
            '                    Case "連絡先TEL１"
            '                        .Vari_Renraku_tel1 = fldvalue.Trim
            '                    Case "連絡先TEL２"
            '                        .Vari_Renraku_tel2 = fldvalue.Trim
            '                    Case "連絡先FAX"
            '                        .Vari_Renraku_fax = fldvalue.Trim
            '                    Case "連絡先携帯１"
            '                        .Vari_Renraku_mobiletel1 = fldvalue.Trim
            '                    Case "連絡先携帯２"
            '                        .Vari_Renraku_mobiletel2 = fldvalue.Trim
            '                    Case "連絡先優先設定(電話番号)"
            '                        .Vari_Renraku_yusentelkbn = fldvalue.Trim
            '                    Case "連絡先間柄"
            '                        .Vari_Renraku_aidagara = fldvalue.Trim
            '                    Case "連絡先備考"
            '                        .Vari_Biko_renraku = fldvalue.Trim
            '                    Case "書類送付先区分"
            '                        .Vari_Sofu_kbn = fldvalue.Trim
            '                    Case "送付先名"
            '                        .Vari_Sofu_name = fldvalue.Trim
            '                    Case "送付先名(SJIS)"
            '                        .Vari_Sofu_namesjis = fldvalue.Trim
            '                    Case "送付先カナ"
            '                        .Vari_Sofu_kana = fldvalue.Trim
            '                    Case "送付先宛名敬称"
            '                        .Vari_Sofu_keisyo = fldvalue.Trim
            '                    Case "送付先郵便番号"
            '                        .Vari_Sofu_postcode = fldvalue.Trim
            '                    Case "送付先住所１"
            '                        .Vari_Sofu_addr1 = fldvalue.Trim
            '                    Case "送付先住所２"
            '                        .Vari_Sofu_addr2 = fldvalue.Trim
            '                    Case "送付先TEL１"
            '                        .Vari_Sofu_tel1 = fldvalue.Trim
            '                    Case "送付先TEL2"
            '                        .Vari_Sofu_tel2 = fldvalue.Trim
            '                    Case "送付先FAX"
            '                        .Vari_Sofu_fax = fldvalue.Trim
            '                    Case "送付先備考"
            '                        .Vari_Biko_sofu = fldvalue.Trim
            '                    Case "イベント担当者No"
            '                        .Vari_Event_tantono = fldvalue.Trim
            '                    Case "絞り込みキーワード"
            '                        .Vari_Findkeyword = fldvalue.Trim
            '                    Case "会計情報"
            '                        .Vari_Kaikei_jouhou = fldvalue.Trim
            '                End Select

            '            Next

            '            '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del sta
            '            ''住所変換/格納
            '            'Dim tmp_address As String = tmp_address1 & " " & tmp_address2
            '            'Dim hash_address As New Hashtable
            '            'hash_address = AddressChange.Get_Address(tmp_address, sqlcnnv10)
            '            '.Vari_Addr_kenno = hash_address("ken_no")
            '            '.Vari_Addr_sino = hash_address("si_no")
            '            '.Vari_Addr_cyo = hash_address("mati")
            '            '.Vari_Addr_cyome = hash_address("cyome")
            '            '.Vari_Addr_cyomeptn = hash_address("chomeptn")
            '            '.Vari_Addr_banti = hash_address("banti")
            '            '.Vari_Addr_etc = hash_address("etc")
            '            '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del end

            '            '20160706 家主個人法人区分の取得方法変更 -add sta
            '            '個人法人区分が個人に設定されている場合は名称から判別する(法人に設定されているデータはそのまま移行)
            '            If .Vari_Kojinhojin_flg = "1" Then
            '                .Vari_Kojinhojin_flg = Me.Get_KojinHojinvalue(.Vari_Ow_name)
            '            End If
            '            '20160706 家主個人法人区分の取得方法変更 -add end

            '            '固定値
            '            .Vari_Useflg = 1
            '            .Vari_History = DefHistory
            '            .Vari_Rowid = Guid.NewGuid.ToString

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

                Dim tmp_sql As String = " SELECT ow_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 家主の名称に特定の文字列が含まれている場合は個人法人区分を法人にして移行する '20160706 家主個人法人区分の取得方法変更
            ''' </summary>
            ''' <param name="owname"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_KojinHojinvalue(ByVal owname As String) As Integer

                Dim rtn_int As Integer = 1
                Dim list_hojinname As New List(Of String) From _
                    {"㈱", "(株)", "株式", "㈲", "(有)", "有限", "(資)", "合資", "(名)", "合名", "法人", "ホーム", "サービス", "工務店", "土地", "建物", "産業", "不動産", "企画"}

                For Each hojinname In list_hojinname

                    If owname <> owname.Replace(hojinname, "") Then
                        rtn_int = 2
                        Return rtn_int
                    End If

                Next

                Return rtn_int

            End Function

        End Class

    End Class

#End Region

#Region "家主口座情報"

    Public Class Owdata_koza_Repository

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

                Dim model_cvitem As New Njc.Model.Owdata_koza_Model             '移行値格納用モデル初期化
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
                Dim tblname_base As String = "owdata"
                Dim tblname As String = "owdata_koza"
                Dim fldnamegrp As String = "ow_no,ow_kozano,kinyu_no,kinyu_tenno,koza_syubetu," & _
                                           "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," & _
                                           "yucyokoza_bango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no," & _
                                           "sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai," & _
                                           "history,rowid,koza_printkbn"

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
                                Case "家主No"
                                    .Vari_Ow_no = fldvalue.Trim
                                Case "口座No"
                                    .Vari_Ow_kozano = fldvalue.Trim
                                Case "金融機関No"
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
                                Case "ゆうちょ口座記号１"
                                    .Vari_Yucyokoza_kigo1 = fldvalue.Trim
                                Case "ゆうちょ口座記号２"
                                    .Vari_Yucyokoza_kigo2 = fldvalue.Trim
                                Case "ゆうちょ口座番号"
                                    .Vari_Yucyokoza_bango = fldvalue.Trim
                                Case "口座備考"
                                    .Vari_Biko_koza = fldvalue.Trim
                                Case "振込情報備考"
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
                                Case "総合振込情報備考"
                                    .Vari_Biko_sgfirai = fldvalue.Trim
                                Case "家主向け帳票の表示"
                                    .Vari_Koza_printkbn = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory
                        .Vari_Rowid = Guid.NewGuid.ToString     'GUID

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

            '    Dim model_cvitem As New Njc.Model.Owdata_koza_Model             '移行値格納用モデル初期化
            '    Dim keycol_main As Integer = 1                                  'メインキー列
            '    Dim keycol_sub As Integer = 2                                   'サブキー列

            '    '************************
            '    '作業準備
            '    '************************

            '    '紐付けデータ取得
            '    Call Me.Get_RelData_KozaSyubetu()

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
            '    If Not rtn Then
            '        Return rtn
            '    End If

            '    'テーブル名/フィールド名セット
            '    Dim tblname_base As String = "owdata"
            '    Dim tblname As String = "owdata_koza"
            '    Dim fldnamegrp As String = "ow_no,ow_kozano,kinyu_no,kinyu_tenno,koza_syubetu," & _
            '                               "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," & _
            '                               "yucyokoza_bango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no," & _
            '                               "sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai," & _
            '                               "history,rowid,koza_printkbn"

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
            '                    Case "家主No"
            '                        .Vari_Ow_no = fldvalue.Trim
            '                    Case "口座No"
            '                        .Vari_Ow_kozano = fldvalue.Trim
            '                    Case "金融機関No"
            '                        .Vari_Kinyu_no = fldvalue.Trim
            '                    Case "金融機関支店No"
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
            '                    Case "口座備考"
            '                        .Vari_Biko_koza = fldvalue.Trim
            '                    Case "振込情報備考"
            '                        .Vari_Biko_furikomi = fldvalue.Trim
            '                    Case "総合振込区分"
            '                        .Vari_Sgfirai_kbn = fldvalue.Trim
            '                    Case "振込依頼人No"
            '                        .Vari_Sgfirai_no = fldvalue.Trim
            '                    Case "振込手数料負担区分"
            '                        .Vari_Sgfirai_tesufutankbn = fldvalue.Trim
            '                    Case "振込手数料計算区分"
            '                        .Vari_Sgfirai_tesukeisankbn = fldvalue.Trim
            '                    Case "振込手数料固定額1"
            '                        .Vari_Sgfirai_tesukotei1gak = fldvalue.Trim
            '                    Case "振込手数料固定額2"
            '                        .Vari_Sgfirai_tesukotei2gak = fldvalue.Trim
            '                    Case "総合振込情報備考"
            '                        .Vari_Biko_sgfirai = fldvalue.Trim
            '                    Case "家主向け帳票の表示"
            '                        .Vari_Koza_printkbn = fldvalue.Trim
            '                End Select

            '            Next

            '            '固定値
            '            .Vari_History = DefHistory
            '            .Vari_Rowid = Guid.NewGuid.ToString     'GUID

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

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT ow_no FROM " & tblname
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

#Region "家主イベント情報"

    Public Class Owdata_event_Repository

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

                Dim model_cvitem As New Njc.Model.Owdata_event_Model            '移行値格納用モデル初期化
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
                Dim tblname_base As String = "owdata"
                Dim tblname As String = "owdata_event"
                Dim fldnamegrp As String = "ow_no,event_kbn,event_cnt,event_ymd,event_data"

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
                    Dim fldname_keymain As String = headervalue(startrow - 1, keycol)

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
                        .Vari_Ow_no = tmp_keymain

                        '-------------------
                        'イベントカウント初期化
                        '-------------------
                        Dim cvitemcnt As Integer = 0

                        '-------------
                        '列単位処理
                        '-------------
                        For cntjj = 2 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            'サブキー取得
                            Dim tmp_keysub As String = ""
                            Select Case fldname
                                Case "年賀状"
                                    tmp_keysub = 1
                                Case "暑中お見舞い"
                                    tmp_keysub = 2
                                Case "誕生日"
                                    tmp_keysub = 3
                                Case "お歳暮"
                                    tmp_keysub = 4
                                Case "お中元"
                                    tmp_keysub = 5
                            End Select

                            'イベントデータ取得
                            .Vari_Event_kbn = tmp_keysub
                            .Vari_Event_data = fldvalue.Trim

                            '全キー取得
                            Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

                            'ログ出力用データ格納(サブキーフィールド)
                            Dim fldname_keysub As String = fldname
                            Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub & " = " & fldvalue

                            '固定値
                            .Vari_Event_cnt = 1
                            .Vari_Event_ymd = Nothing

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

                        Next

                        '--------------------------------------------------------------------
                        '移行した備考が1データ以上ある場合移行したレコードの数を更新する
                        '--------------------------------------------------------------------
                        tmp_cvcnt = tmp_cvcnt + 1

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

                Dim tmp_sql As String = " SELECT ow_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT ow_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "家主メモ情報"

    Public Class Owdata_memo_Repository

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

                Dim model_cvitem As New Njc.Model.Owdata_memo_Model             '移行値格納用モデル初期化
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
                Dim tblname_base As String = "owdata"
                Dim tblname As String = "owdata_memo"
                Dim fldnamegrp As String = "ow_no,memo_no,memo,history"

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
                Dim cvtaisyocnt As Integer = 0  '20161007 メモ関連のログ出力修正 -add

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
                        Dim cvflg As Boolean = True         '20161007 メモ関連のログ出力修正 -add
                        Dim keychkflg As Boolean = True     '20161007 メモ関連のログ出力修正 -add

                        'キー値取得
                        Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol - 1)).Trim
                        .Vari_Ow_no = tmp_keymain

                        '備考カウント初期化
                        Dim cvitemcnt As Integer = 0
                        Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

                        '20161007 メモ関連のログ出力修正 -chg sta
                        ''移行値取得
                        'For cntjj = 1 To readtbl.Columns.Count - 1

                        '    'サブキー取得
                        '    Dim tmp_keysub As String = (cntjj).ToString

                        '    '全キー取得
                        '    Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

                        '    'ログ出力用データ格納(サブキーフィールド)
                        '    Dim fldname_keysub As String = readtbl.Columns(cntjj).ColumnName.Trim
                        '    Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub

                        '    '登録値取得
                        '    Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                        '    .Vari_Memo = fldvalue

                        '    '固定値
                        '    .Vari_Memo_no = (cntjj).ToString
                        '    .Vari_History = DefHistory

                        '    'メモにデータが存在する場合に書込処理を行う
                        '    If .Vari_Memo <> "" Then

                        '        '20160928 メモ関連の移行件数表示修正 -add
                        '        tmp_cvrowcnt = tmp_cvrowcnt + 1

                        '        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        '        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        '        'データチェック
                        '        Dim skipflg As Boolean = False
                        '        Dim hash_cvitem As New Hashtable
                        '        Dim hash_log As New Hashtable
                        '        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '        '書込処理
                        '        If Not skipflg Then

                        '            '挿入処理
                        '            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                        '            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        '            If normalflg Then
                        '                list_chkduplicate.Add(fldvalue_key)
                        '                cvitemcnt = cvitemcnt + 1
                        '            End If

                        '        End If

                        '        'ログ出力メッセージ整形
                        '        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        '        'ログ出力
                        '        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                        '            Dim tmp_cnt As Integer = 0
                        '            '挿入
                        '            For Each logvalue In sortlist_log
                        '                Dim tmp_sql_insert As String = logvalue.Value
                        '                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                        '            Next
                        '            '初期化
                        '            tmp_logcnt = 0
                        '            sortlist_log.Clear()
                        '        End If

                        '    End If

                        'Next

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
                                Dim log_key As String = "owdata_memo-ow_no"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(tmp_keymain) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "owdata_memo-ow_no"
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
                        '20161007 メモ関連のログ出力修正 -chg end

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
                '20161007 メモ関連のログ出力修正 -chg sta
                ''20160928 メモ関連の移行件数表示修正 -chg sta
                ''midrowcnt = rowcnt
                'midrowcnt = totalrowcnt
                ''20160928 メモ関連の移行件数表示修正 -chg end
                midrowcnt = cvtaisyocnt
                '20161007 メモ関連のログ出力修正 -chg end

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

                Dim tmp_sql As String = " SELECT ow_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT ow_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

End Namespace


