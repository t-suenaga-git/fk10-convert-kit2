Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "物件基本情報"

    Public Class Bkdata_Repository

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
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Bkdata_Model                  '移行値格納用モデル初期化
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
                Dim tblname As String = "bkdata"
                '20160519 EXEUpdateに伴う修正 物件基本情報 -add (最後尾にlastupdateを追加)
                Dim fldnamegrp As String = "bk_guid,bk_no,bk_deleteflg,delete_guid,delete_day," & _
                                           "delete_cnt,bk_name,bk_deletename,bk_kana,tatemono_sikibetu," & _
                                           "post_code,addr_kenno,addr_sino,addr_cyo,addr_cyome," & _
                                           "addr_cyomeptn,addr_banti,addr_etc,history,rowid," & _
                                           "bk_namesjis,bk_gaibuno,delete_cause,lastupdate"

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
                        '20170511 汎用_物件住所の移行不備対応 -del sta
                        'Dim tmp_addresscyotiiki As String = ""
                        'Dim tmp_addresscyobanti1 As String = ""
                        'Dim tmp_addresscyobanti2 As String = ""
                        'Dim tmp_addressother As String = ""
                        '20170511 汎用_物件住所の移行不備対応 -del end

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
                                Case "物件NO"
                                    .Vari_Bk_no = fldvalue.Trim
                                    'Case "物件削除フラグ"
                                    '    .Vari_Bk_deleteflg = fldvalue.Trim
                                    'Case "削除日"
                                    '    .Vari_Delete_day = fldvalue.Trim
                                    'Case "削除復旧回数"
                                    '    .Vari_Delete_cnt = fldvalue.Trim
                                Case "物件名"
                                    .Vari_Bk_name = fldvalue.Trim
                                    'Case "物件名(削除時の退避用)"
                                    '    .Vari_Bk_deletename = fldvalue.Trim
                                Case "物件カナ"
                                    .Vari_Bk_kana = fldvalue.Trim
                                Case "建物識別コード"
                                    .Vari_Tatemono_sikibetu = fldvalue.Trim
                                Case "郵便番号"
                                    .Vari_Post_code = fldvalue.Trim
                                Case "都道府県コード"
                                    .Vari_Addr_kenno = fldvalue.Trim
                                Case "市区町村コード"
                                    .Vari_Addr_sino = fldvalue.Trim
                                Case "町地域"
                                    '20160622 住所分割処理対応 -chg sta
                                    '町地域はそのまま移行する(丁番地以降が分割対象)
                                    ''.Vari_Addr_cyo = fldvalue.Trim
                                    'tmp_addresscyotiiki = fldvalue.Trim
                                    .Vari_Addr_cyo = fldvalue.Trim
                                    '20160622 住所分割処理対応 -chg end
                                Case "丁番地"
                                    '20170511 汎用_物件住所の移行不備対応 -chg sta
                                    ''.Vari_Addr_cyome = fldvalue.Trim
                                    'tmp_addresscyobanti1 = fldvalue.Trim
                                    .Vari_Addr_cyome = fldvalue.Trim
                                    '20170511 汎用_物件住所の移行不備対応 -chg end
                                Case "丁番地選択"
                                    '.Vari_Addr_cyomeptn = fldvalue.Trim
                                Case "街区番号地番"
                                    '20170511 汎用_物件住所の移行不備対応 -chg sta
                                    ''.Vari_Addr_banti = fldvalue.Trim
                                    'tmp_addresscyobanti2 = fldvalue.Trim
                                    .Vari_Addr_banti = fldvalue.Trim
                                    '20170511 汎用_物件住所の移行不備対応 -chg end
                                Case "その他"
                                    '20170511 汎用_物件住所の移行不備対応 -chg sta
                                    ''.Vari_Addr_etc = fldvalue.Trim
                                    'tmp_addressother = fldvalue.Trim
                                    .Vari_Addr_etc = fldvalue.Trim
                                    '20170511 汎用_物件住所の移行不備対応 -chg end
                                Case "物件名(SJIS)"
                                    .Vari_Bk_namesjis = fldvalue.Trim
                                Case "物件外部No"
                                    .Vari_Bk_gaibuno = fldvalue.Trim
                                    'Case "削除理由"
                                    '    .Vari_Delete_cause = fldvalue.Trim
                            End Select

                        Next

                        '20170511 汎用_物件住所の移行不備対応 -chg sta
                        ''住所変換/格納
                        ''20160622 住所分割処理対応 -chg sta
                        ''Dim tmp_address As String = tmp_addresscyotiiki & " " & tmp_addresscyobanti1 & tmp_addresscyobanti2 & " " & tmp_addressother
                        ''Dim hash_address As New Hashtable
                        ''hash_address = AddressChange.Get_CyoBanti(tmp_address)
                        ''.Vari_Addr_cyo = hash_address("mati")
                        ''.Vari_Addr_cyome = hash_address("cyome")
                        ''.Vari_Addr_cyomeptn = hash_address("chomeptn")
                        ''.Vari_Addr_banti = hash_address("banti")
                        ''.Vari_Addr_etc = hash_address("etc")
                        'Dim tmp_address As String = tmp_addresscyobanti1 & tmp_addresscyobanti2 & tmp_addressother
                        'Dim hash_address As New Hashtable
                        'hash_address = AddressChange.Get_CyoBanti(tmp_address)
                        '.Vari_Addr_cyome = hash_address("cyome")
                        '.Vari_Addr_cyomeptn = hash_address("chomeptn")
                        '.Vari_Addr_banti = hash_address("banti")
                        '.Vari_Addr_etc = hash_address("etc")
                        ''20160622 住所分割処理対応 -chg end
                        '丁番地の値の有無で処理を分岐する(あれば「丁目」を設定)
                        If .Vari_Addr_cyome <> "" Then
                            .Vari_Addr_cyomeptn = "1"
                        Else
                            .Vari_Addr_cyomeptn = "-1"
                        End If
                        '20170511 汎用_物件住所の移行不備対応 -chg end

                        '固定値
                        .Vari_Bk_guid = Guid.NewGuid.ToString
                        .Vari_Bk_deleteflg = 0
                        .Vari_Delete_guid = Nothing
                        .Vari_Delete_day = ""
                        .Vari_Delete_cnt = ""
                        .Vari_Bk_deletename = ""
                        .Vari_Delete_cause = ""
                        .Vari_History = DefHistory
                        .Vari_Rowid = Guid.NewGuid.ToString

                        '20160519 EXEUpdateに伴う修正 物件基本情報 -add sta
                        'とりあえずNULLにしておく (必要であればここで値を設定)
                        .Vari_Lastupdate = Nothing
                        '20160519 EXEUpdateに伴う修正 物件基本情報 -add end

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '20160704 物件住所から郵便番号を読み込む処理の追加 -add sta
                        If skipflg = False And .Vari_Post_code = "" Then
                            skipflg = Not (Me.Get_PostCode(sqlcnnv10, tblname, hash_cvitem, hash_log))
                        End If
                        '20160704 物件住所から郵便番号を読み込む処理の追加 -add end

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
                '物件基本情報テーブル外部No一括更新
                Dim tmpcnt As Integer = 0
                Dim gaibunoupdatesql As String = Me.Get_UseQry_Update()
                DBExec.Exec_NonQuery(sqlcnnv10, gaibunoupdatesql, tmpcnt)

                '外部No管理テーブル一括更新
                Dim obj_bkdatagaibuno As New Njc.Repository.Bkdata_gaibuno_Repository.SubConv
                rtn = obj_bkdatagaibuno.Set_Bkdata_Gaibuno(sqlcnnv10)

                '返却
                Return rtn

            End Function

            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' [bkdata].[bk_gaibuno]の一括更新クエリ
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Update() As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " UPDATE bkdata SET "
                tmp_sql = tmp_sql & " 	bk_gaibuno = GAIBUNO "
                tmp_sql = tmp_sql & " FROM bkdata AS BK1 "
                tmp_sql = tmp_sql & " LEFT JOIN "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 ROW_NUMBER()OVER(ORDER BY bk_no) AS GAIBUNO "
                tmp_sql = tmp_sql & " 			,bk_guid "
                tmp_sql = tmp_sql & " 			,bk_no "
                tmp_sql = tmp_sql & " 		FROM bkdata "
                tmp_sql = tmp_sql & " 	) AS BK2 "
                tmp_sql = tmp_sql & " ON BK1.bk_guid = BK2.bk_guid "
                tmp_sql = tmp_sql & " ; "

                Return tmp_sql

            End Function

            ''' <summary>
            ''' 郵便番号が存在しない場合、住所から取得する '20160704 物件住所から郵便番号を読み込む処理の追加
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_PostCode(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_cvitem As Hashtable, ByRef hash_log As Hashtable) As Boolean

                Dim rtn As Boolean = True
                Dim tmp_kenno As String = hash_cvitem("addr_kenno")
                Dim tmp_sino As String = hash_cvitem("addr_sino")
                '20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -add sta
                '県No、市Noが存在しない場合の処理追加に伴い移動
                Dim errstr As String = ""
                Dim log_key As String = tblname & "-" & "post_code"
                '20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -add end

                If tmp_kenno <> "" And tmp_sino <> "" Then

                    Dim postcodevalu As String = ""
                    Dim tmp_sql As String = ""
                    tmp_sql = tmp_sql & " SELECT TOP 1 post_code FROM post_code "
                    tmp_sql = tmp_sql & " WHERE "
                    tmp_sql = tmp_sql & " ken_no = " & tmp_kenno
                    tmp_sql = tmp_sql & " AND "
                    tmp_sql = tmp_sql & " si_no = " & tmp_sino

                    postcodevalu = DBExec.Exec_Scalar(tmp_sql, sqlcnnv10)

                    '20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -del sta
                    '県No、市Noが存在しない場合の処理追加に伴い移動
                    'Dim errstr As String = ""
                    'Dim log_key As String = tblname & "-" & "post_code"
                    '20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -del end
                    If postcodevalu = "" Then
                        errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                        hash_log.Clear()
                        hash_log.Add(log_key, errstr)
                        rtn = False
                    Else
                        hash_cvitem("post_code") = postcodevalu
                        errstr = LOG_NAIYO_GET_POSTCODE & "-" & LOG_HUBI_GET_POSTCODE & "-" & LOG_TAISYO_GET_POSTCODE
                        hash_log.Add(log_key, errstr)
                    End If

                    '20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -add sta
                    '県No、市Noが存在しない場合は移行対象外とする。(ログを出力)
                Else
                    errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                    hash_log.Clear()
                    hash_log.Add(log_key, errstr)
                    rtn = False
                    '20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -add end
                End If

                Return rtn

            End Function

        End Class

    End Class

#End Region

#Region "物件外部No情報"

    '物件基本情報作成直後に呼び出して処理する

    Public Class Bkdata_gaibuno_Repository

        Public Class SubConv

            ''' <summary>
            ''' 物件基本情報(bkdata)から外部Noを取得して外部No管理テーブル(bkdata_gaibuno)へ挿入する
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Bkdata_Gaibuno(ByVal sqlcnnv10 As SqlConnection) As Boolean

                Dim rtn As Boolean = True
                Dim hash_gaibuno As New Hashtable
                Dim tmp_hash As New Hashtable
                Dim model_cvitem As New Njc.Model.Bkdata_gaibuno_Model
                Dim normalflg As Boolean = True
                Dim tmp_cnt As Integer = 0
                Dim list_bkgaibuno As New List(Of String)

                '対象テーブル初期化
                Dim tmp_sql_delete As String = " DELETE FROM bkdata_gaibuno "
                rtn = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_delete, tmp_cnt)

                'テーブル名/フィールド名セット
                Dim tblname As String = "bkdata_gaibuno"
                Dim fldnamegrp As String = "current_no,history"

                'bkdataから外部No最大値を抽出
                Dim tmp_sql_select As String = " SELECT MAX(bk_gaibuno) FROM bkdata "

                '変数格納
                With model_cvitem

                    .Vari_Current_no = DBExec.Exec_Scalar(tmp_sql_select, sqlcnnv10)
                    .Vari_History = DefHistory

                End With

                'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                '挿入処理
                Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, normalflg)

                Return rtn

            End Function

        End Class

    End Class

#End Region

#Region "物件詳細情報"

    Public Class Bkdata_detail_Repository

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

                Dim model_cvitem As New Njc.Model.Bkdata_detail_Model           '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列

                '************************
                '作業準備
                '************************

                '紐付けデータ取得
                '2016.04.06 紐付データ取得処理を外出し -chg sta
                'Call Me.Get_RelData_Bkrui()
                'Call Me.Get_RelData_Kozo()
                Call SetRelItemToObject.Set_RelData_Bkrui()
                Call SetRelItemToObject.Set_RelData_Kozo()
                '2016.04.06 紐付データ取得処理を外出し -chg end

                '2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -add
                Call SetRelItemToObject.Set_RelData_Yototiki()
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
                Dim tblname As String = "bkdata_detail"
                '20160519 EXEUpdateに伴う修正 物件詳細情報 -chg sta
                'Dim fldnamegrp As String = "bk_guid,bk_ruinokbn,gps_wgsido,gps_wgskeido,gps_tokyoido," & _
                '                           "gps_tokyokeido,bk_dentetuflg,syunko_ymd,kaidate,tika," & _
                '                           "elevator_flg,elevator_number,rooftop_flg,kozo_nokbn,kozo_yanekbn," & _
                '                           "moto_gy_fudono,biko_kihon,kosu_total,men_nobeyuka,men_nobeyukatubo," & _
                '                           "men_sikiti,men_sikititubo,men_parking,men_parkingtubo,men_yukatoki," & _
                '                           "men_yukatokitubo,men_sikititoki,men_sikititokitubo,default_sqtuki,keiyakuyou_sime," & _
                '                           "yatin_jisansaki,yatin_kozano,yatin_kykozaflg,yatin_kykozano,kozo_taikakbn," & _
                '                           "gy_sekono,gy_hosyuno,kanri_hosiki,kanri_gyname,kanri_gytanto," & _
                '                           "kanri_gytel,kanrinin_gyomukeitai,kanrinin_tel,jisya_no,kotu_sonota1," & _
                '                           "kotu_sonota1kyori,kotu_sonota2,kotu_sonota2kyori,denki_gy_lifeno,water_gy_lifeno," & _
                '                           "gas_gy_lifeno,haisui_gy_lifeno,toyu_gy_lifeno,lifeline1_gy_lineno,lifeline2_gy_lineno," & _
                '                           "lifeline3_gy_lineno,kensingyomu_umu,kensinorder_kbn,parenthendo_useflg,parking_kanriflg," & _
                '                           "parking_car,parking_bike,parking_bicycle,parking_bicyclefreeflg,kubunsyo," & _
                '                           "soymd_basis,isiwata_kirokukbn,isiwata_syokai1flg,isiwata_syokai2flg,isiwata_syokai3flg," & _
                '                           "isiwata_syokai4flg,isiwata_gy_sekono,isiwata_cyosaymd,isiwata_cyosakikankbn,isiwata_cyosahani," & _
                '                           "ishiwata_useflg,isiwata_usearea,isiwata_biko,taisin_sindanflg,taisin_syokai1flg," & _
                '                           "taisin_syokai2flg,taisin_syokai3flg,taisin_syorui1flg,taisin_syorui2flg,taisin_syorui3flg," & _
                '                           "taisin_biko,horei_dosyatiiki,horei_ruikbn,horei_naiyo,sikiti_riyoruikbn," & _
                '                           "sikiti_kystartymd,sikiti_kyendymd,sikiti_biko,nyukyoritu_startymd,history," & _
                '                           "hyothergyfudo_flg,gomi_hosoku,toki_ymd,syo_kenriflg,syo_kenrikbn," & _
                '                           "other_kenriflg,men_kentiku,men_kentikutubo,men_kentikutoki,men_kentikutokitubo," & _
                '                           "bk_logonuser_no,kanri_gyfax,homeelevator_flg,horei_dosyatokubetutiiki,horei_zoseitakutitiiki," & _
                '                           "horei_tunamitiiki,horei_dosyatiikibiko,horei_dosyatokubetutiikibiko,horei_zoseitakutitiikibiko,horei_tunamitiikibiko," & _
                '                           "svbunrui_no,krbunrui_no,kanrinin_name,kozo_other,kadoti_flg," & _
                '                           "cityplan,yototiki,kanrinin_namesjis,kanri_gynamesjis,kanri_gytantosjis," & _
                '                           "parking_caraki,parking_bikeaki,parking_bicycleaki,syogaku_name,syogaku_kyori," & _
                '                           "cyugaku_name,cyugaku_kyori,bk_area_no"
                '20160829 革命10バージョンアップに伴う修正 rooftop_flg を削除 -del 
                Dim fldnamegrp As String = "bk_guid,bk_ruinokbn,gps_wgsido,gps_wgskeido,gps_tokyoido," & _
                                           "gps_tokyokeido,bk_dentetuflg,syunko_ymd,kaidate,tika," & _
                                           "elevator_flg,elevator_number,kozo_nokbn,kozo_yanekbn," & _
                                           "moto_gy_fudono,biko_kihon,kosu_total,men_nobeyuka,men_nobeyukatubo," & _
                                           "men_sikiti,men_sikititubo,men_parking,men_parkingtubo,men_yukatoki," & _
                                           "men_yukatokitubo,men_sikititoki,men_sikititokitubo,default_sqtuki,keiyakuyou_sime," & _
                                           "yatin_jisansaki,yatin_kozano,yatin_kykozaflg,yatin_kykozano,kozo_taikakbn," & _
                                           "gy_sekono,gy_hosyuno,kanri_hosiki,kanri_gyname,kanri_gytanto," & _
                                           "kanri_gytel,kanrinin_gyomukeitai,kanrinin_tel,jisya_no,kotu_sonota1," & _
                                           "kotu_sonota1kyori,kotu_sonota2,kotu_sonota2kyori,denki_gy_lifeno,water_gy_lifeno," & _
                                           "gas_gy_lifeno,haisui_gy_lifeno,toyu_gy_lifeno,lifeline1_gy_lineno,lifeline2_gy_lineno," & _
                                           "lifeline3_gy_lineno,kensingyomu_umu,kensinorder_kbn,parenthendo_useflg,parking_kanriflg," & _
                                           "parking_car,parking_bike,parking_bicycle,parking_bicyclefreeflg,kubunsyo," & _
                                           "isiwata_kirokukbn,isiwata_syokai1flg,isiwata_syokai2flg,isiwata_syokai3flg,isiwata_syokai4flg," & _
                                           "isiwata_gy_sekono,isiwata_cyosaymd,isiwata_cyosakikankbn,isiwata_cyosahani,ishiwata_useflg," & _
                                           "isiwata_usearea,isiwata_biko,taisin_sindanflg,taisin_syokai1flg,taisin_syokai2flg," & _
                                           "taisin_syokai3flg,taisin_syorui1flg,taisin_syorui2flg,taisin_syorui3flg,taisin_biko," & _
                                           "horei_dosyatiiki,horei_ruikbn,horei_naiyo,sikiti_riyoruikbn,sikiti_kystartymd," & _
                                           "sikiti_kyendymd,sikiti_biko,nyukyoritu_startymd,history,hyothergyfudo_flg," & _
                                           "gomi_hosoku,toki_ymd,syo_kenriflg,syo_kenrikbn,other_kenriflg," & _
                                           "men_kentiku,men_kentikutubo,men_kentikutoki,men_kentikutokitubo,bk_logonuser_no," & _
                                           "kanri_gyfax,homeelevator_flg,horei_dosyatokubetutiiki,horei_zoseitakutitiiki,horei_tunamitiiki," & _
                                           "horei_dosyatiikibiko,horei_dosyatokubetutiikibiko,horei_zoseitakutitiikibiko,horei_tunamitiikibiko,svbunrui_no," & _
                                           "krbunrui_no,kanrinin_name,kozo_other,kadoti_flg,cityplan," & _
                                           "yototiki,kanrinin_namesjis,kanri_gynamesjis,kanri_gytantosjis,parking_caraki," & _
                                           "parking_bikeaki,parking_bicycleaki,syogaku_name,syogaku_kyori,cyugaku_name," & _
                                           "cyugaku_kyori,bk_area_no,emergencyelevator_flg"
                '20160519 EXEUpdateに伴う修正 物件詳細情報 -chg end

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                        Dim tmp_tosiyoto1 As String = ""    '都市計画/用途地域作業用変数生成
                        Dim tmp_tosiyoto2 As String = ""    '都市計画/用途地域作業用変数生成

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
                                Case "物件分類"
                                    .Vari_Bk_ruinokbn = fldvalue.Trim
                                Case "緯度"
                                    .Vari_Gps_wgsido = fldvalue.Trim
                                Case "経度"
                                    .Vari_Gps_wgskeido = fldvalue.Trim
                                Case "緯度(日本測地系)"
                                    .Vari_Gps_tokyoido = fldvalue.Trim
                                Case "経度(日本測地系)"
                                    .Vari_Gps_tokyokeido = fldvalue.Trim
                                Case "電鉄物件フラグ"
                                    .Vari_Bk_dentetuflg = fldvalue.Trim
                                Case "竣工日"
                                    .Vari_Syunko_ymd = fldvalue.Trim
                                Case "地上階建て"
                                    .Vari_Kaidate = fldvalue.Trim
                                Case "地下階"
                                    .Vari_Tika = fldvalue.Trim
                                Case "エレベータフラグ"
                                    .Vari_Elevator_flg = fldvalue.Trim
                                Case "エレベータ数"
                                    .Vari_Elevator_number = fldvalue.Trim
                                    '20160829 革命10バージョンアップに伴う修正 -del sta
                                    'Case "屋上フラグ"
                                    '    .Vari_Rooftop_flg = fldvalue.Trim
                                    '20160829 革命10バージョンアップに伴う修正 -del end
                                Case "建物構造(基本)-構造"
                                    .Vari_Kozo_nokbn = fldvalue.Trim
                                Case "建物構造(基本)-屋根構造"
                                    .Vari_Kozo_yanekbn = fldvalue.Trim
                                Case "情報元業者NO"
                                    .Vari_Moto_gy_fudono = fldvalue.Trim
                                Case "備考(基本情報)"
                                    .Vari_Biko_kihon = fldvalue.Trim
                                Case "総戸数"
                                    '20161125 総戸数の半角変換処理を追加 -chg sta
                                    '.Vari_Kosu_total = fldvalue.Trim
                                    .Vari_Kosu_total = StrConv(fldvalue, VbStrConv.Narrow)
                                    '20161125 総戸数の半角変換処理を追加 -chg end
                                Case "物件面積-延床面積"
                                    .Vari_Men_nobeyuka = fldvalue.Trim
                                Case "物件面積-延床面積坪数"
                                    .Vari_Men_nobeyukatubo = fldvalue.Trim
                                Case "物件面積-敷地面積"
                                    .Vari_Men_sikiti = fldvalue.Trim
                                Case "物件面積-敷地面積坪数"
                                    .Vari_Men_sikititubo = fldvalue.Trim
                                Case "物件面積-駐車面積"
                                    .Vari_Men_parking = fldvalue.Trim
                                Case "物件面積-駐車面積坪数"
                                    .Vari_Men_parkingtubo = fldvalue.Trim
                                Case "物件面積-延床面積(登記)"
                                    .Vari_Men_yukatoki = fldvalue.Trim
                                Case "物件面積-延床面積坪数(登記)"
                                    .Vari_Men_yukatokitubo = fldvalue.Trim
                                Case "物件面積-敷地面積(公簿)"
                                    .Vari_Men_sikititoki = fldvalue.Trim
                                Case "物件面積-敷地面積坪数(公簿)"
                                    .Vari_Men_sikititokitubo = fldvalue.Trim
                                Case "入金口座-請求月初期値"
                                    .Vari_Default_sqtuki = fldvalue.Trim
                                Case "契約書用入金締切日"
                                    .Vari_Keiyakuyou_sime = fldvalue.Trim
                                Case "家賃持参先"
                                    .Vari_Yatin_jisansaki = fldvalue.Trim
                                Case "入金口座-家賃入金口座No"
                                    .Vari_Yatin_kozano = fldvalue.Trim
                                Case "入金口座-契約金用入金口座の有無"
                                    .Vari_Yatin_kykozaflg = fldvalue.Trim
                                Case "入金口座-契約金用入金口座No"
                                    .Vari_Yatin_kykozano = fldvalue.Trim
                                Case "耐火構造区分"
                                    .Vari_Kozo_taikakbn = fldvalue.Trim
                                Case "施工会社No"
                                    .Vari_Gy_sekono = fldvalue.Trim
                                Case "保守業者No"
                                    .Vari_Gy_hosyuno = fldvalue.Trim
                                Case "管理形態-管理方式"
                                    .Vari_Kanri_hosiki = fldvalue.Trim
                                Case "管理形態-管理業者名"
                                    .Vari_Kanri_gyname = fldvalue.Trim
                                Case "管理形態-管理業者担当者"
                                    .Vari_Kanri_gytanto = fldvalue.Trim
                                Case "管理形態-管理業者電話番号"
                                    .Vari_Kanri_gytel = fldvalue.Trim
                                Case "管理形態-管理業者業務形態"
                                    .Vari_Kanrinin_gyomukeitai = fldvalue.Trim
                                Case "管理形態-管理人電話番号"
                                    .Vari_Kanrinin_tel = fldvalue.Trim
                                Case "支店No"
                                    .Vari_Jisya_no = fldvalue.Trim
                                Case "その他交通１-その他交通"
                                    .Vari_Kotu_sonota1 = fldvalue.Trim
                                Case "その他交通１-距離"
                                    .Vari_Kotu_sonota1kyori = fldvalue.Trim
                                Case "その他交通２-その他交通"
                                    .Vari_Kotu_sonota2 = fldvalue.Trim
                                Case "その他交通２-距離"
                                    .Vari_Kotu_sonota2kyori = fldvalue.Trim
                                Case "ライフライン-電気-公共機関No"
                                    .Vari_Denki_gy_lifeno = fldvalue.Trim
                                Case "ライフライン-上水-公共機関No"
                                    .Vari_Water_gy_lifeno = fldvalue.Trim
                                Case "ライフライン-ガス-公共機関No"
                                    .Vari_Gas_gy_lifeno = fldvalue.Trim
                                Case "ライフライン-排水-公共機関No"
                                    .Vari_Haisui_gy_lifeno = fldvalue.Trim
                                Case "ライフライン-灯油-公共機関No"
                                    .Vari_Toyu_gy_lifeno = fldvalue.Trim
                                Case "ライフライン-その他１-公共機関No"
                                    .Vari_Lifeline1_gy_lineno = fldvalue.Trim
                                Case "ライフライン-その他２-公共機関No"
                                    .Vari_Lifeline2_gy_lineno = fldvalue.Trim
                                Case "ライフライン-その他３-公共機関No"
                                    .Vari_Lifeline3_gy_lineno = fldvalue.Trim
                                Case "検針業務の有無"
                                    .Vari_Kensingyomu_umu = fldvalue.Trim
                                Case "検針登録の並び順"
                                    .Vari_Kensinorder_kbn = fldvalue.Trim
                                Case "親子メーター変動費の使用有無"
                                    .Vari_Parenthendo_useflg = fldvalue.Trim
                                Case "駐車場-付随駐車場有無"
                                    .Vari_Parking_kanriflg = fldvalue.Trim
                                Case "駐車場-自動車台数"
                                    .Vari_Parking_car = fldvalue.Trim
                                Case "駐車場-バイク台数"
                                    .Vari_Parking_bike = fldvalue.Trim
                                Case "駐車場-自転車台数"
                                    .Vari_Parking_bicycle = fldvalue.Trim
                                Case "駐車場-自転車利用自由フラグ"
                                    .Vari_Parking_bicyclefreeflg = fldvalue.Trim
                                Case "所有者-一棟・所有区分"
                                    .Vari_Kubunsyo = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 物件詳細情報 -del sta
                                    'Case "該当月/送金月ベース決定"
                                    '    .Vari_Soymd_basis = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 物件詳細情報 -del end
                                Case "石綿使用調査-調査の有無"
                                    .Vari_Isiwata_kirokukbn = fldvalue.Trim
                                Case "石綿使用調査-調査結果の問合せ先-所有者"
                                    .Vari_Isiwata_syokai1flg = fldvalue.Trim
                                Case "石綿使用調査-調査結果の問合せ先-管理組合"
                                    .Vari_Isiwata_syokai2flg = fldvalue.Trim
                                Case "石綿使用調査-調査結果の問合せ先-管理業者"
                                    .Vari_Isiwata_syokai3flg = fldvalue.Trim
                                Case "石綿使用調査-調査結果の問合せ先-施工業者"
                                    .Vari_Isiwata_syokai4flg = fldvalue.Trim
                                Case "石綿使用調査-調査年月日"
                                    .Vari_Isiwata_cyosaymd = fldvalue.Trim
                                Case "石綿使用調査-実施機関"
                                    .Vari_Isiwata_cyosakikankbn = fldvalue.Trim
                                Case "石綿使用調査-調査範囲"
                                    .Vari_Isiwata_cyosahani = fldvalue.Trim
                                Case "石綿使用調査-使用有無"
                                    .Vari_Ishiwata_useflg = fldvalue.Trim
                                Case "石綿使用調査-使用箇所"
                                    .Vari_Isiwata_usearea = fldvalue.Trim
                                Case "石綿使用調査-備考"
                                    .Vari_Isiwata_biko = fldvalue.Trim
                                Case "耐震診断-診断有無"
                                    .Vari_Taisin_sindanflg = fldvalue.Trim
                                Case "耐震診断-診断記録の問合せ先-所有者"
                                    .Vari_Taisin_syokai1flg = fldvalue.Trim
                                Case "耐震診断-診断記録の問合せ先-管理組合"
                                    .Vari_Taisin_syokai2flg = fldvalue.Trim
                                Case "耐震診断-診断記録の問合せ先-管理業者"
                                    .Vari_Taisin_syokai3flg = fldvalue.Trim
                                Case "耐震診断-耐震基準適合証明書の写し"
                                    .Vari_Taisin_syorui1flg = fldvalue.Trim
                                Case "耐震診断-住宅性能評価所の写し"
                                    .Vari_Taisin_syorui2flg = fldvalue.Trim
                                Case "耐震診断-耐震診断結果の写し"
                                    .Vari_Taisin_syorui3flg = fldvalue.Trim
                                Case "耐震診断-備考"
                                    .Vari_Taisin_biko = fldvalue.Trim
                                Case "法令-土砂災害警戒地域内外"
                                    .Vari_Horei_dosyatiiki = fldvalue.Trim
                                Case "法令-法令分類"
                                    .Vari_Horei_ruikbn = fldvalue.Trim
                                Case "法令-法令内容"
                                    .Vari_Horei_naiyo = fldvalue.Trim
                                Case "敷地利用-敷地利用種類"
                                    .Vari_Sikiti_riyoruikbn = fldvalue.Trim
                                Case "敷地利用-契約期間開始"
                                    .Vari_Sikiti_kystartymd = fldvalue.Trim
                                Case "敷地利用-契約期間終了"
                                    '20160706 敷地利用-契約期間終了の文字列日付変換処理の追加 -chg sta
                                    '.Vari_Sikiti_kyendymd = fldvalue.Trim
                                    .Vari_Sikiti_kyendymd = Me.Get_ChgDate(fldvalue.Trim)
                                    '20160706 敷地利用-契約期間終了の文字列日付変換処理の追加 -chg end
                                Case "敷地利用-備考"
                                    .Vari_Sikiti_biko = fldvalue.Trim
                                Case "入居率一覧対象開始日"
                                    .Vari_Nyukyoritu_startymd = fldvalue.Trim
                                Case "部屋毎に業者が異なる場合フラグ"
                                    .Vari_Hyothergyfudo_flg = fldvalue.Trim
                                Case "ゴミ出しに関する補足情報"
                                    .Vari_Gomi_hosoku = fldvalue.Trim
                                Case "登記情報の日付"
                                    .Vari_Toki_ymd = fldvalue.Trim
                                Case "所有権にかかる権利有無"
                                    .Vari_Syo_kenriflg = fldvalue.Trim
                                Case "所有権にかかる権利の種類"
                                    .Vari_Syo_kenrikbn = fldvalue.Trim
                                Case "所有権以外の権利有無"
                                    .Vari_Other_kenriflg = fldvalue.Trim
                                Case "物件面積-建築面積"
                                    .Vari_Men_kentiku = fldvalue.Trim
                                Case "物件面積-建築面積坪数"
                                    .Vari_Men_kentikutubo = fldvalue.Trim
                                Case "物件面積-建築面積(登記)"
                                    .Vari_Men_kentikutoki = fldvalue.Trim
                                Case "物件面積-建築面積坪数(登記)"
                                    .Vari_Men_kentikutokitubo = fldvalue.Trim
                                Case "自社担当者No"
                                    .Vari_Bk_logonuser_no = fldvalue.Trim
                                Case "管理形態-管理業者FAX"
                                    .Vari_Kanri_gyfax = fldvalue.Trim
                                Case "エレベータフラグ(家)"
                                    .Vari_Homeelevator_flg = fldvalue.Trim
                                Case "法令-土砂災害特別警戒区域内外"
                                    .Vari_Horei_dosyatokubetutiiki = fldvalue.Trim
                                Case "法令-造成宅地防災区域内外"
                                    .Vari_Horei_zoseitakutitiiki = fldvalue.Trim
                                Case "法令-津波災害警戒区域内外"
                                    .Vari_Horei_tunamitiiki = fldvalue.Trim
                                Case "法令-土砂災害警戒地域備考"
                                    .Vari_Horei_dosyatiikibiko = fldvalue.Trim
                                Case "法令-土砂災害特別警戒区域備考"
                                    .Vari_Horei_dosyatokubetutiikibiko = fldvalue.Trim
                                Case "法令-造成宅地防災区域備考"
                                    .Vari_Horei_zoseitakutitiikibiko = fldvalue.Trim
                                Case "法令-津波災害警戒区域備考"
                                    .Vari_Horei_tunamitiikibiko = fldvalue.Trim
                                Case "管理形態-管理人名"
                                    .Vari_Kanrinin_name = fldvalue.Trim
                                Case "その他建物構造"
                                    .Vari_Kozo_other = fldvalue.Trim
                                Case "角地フラグ"
                                    .Vari_Kadoti_flg = fldvalue.Trim
                                    '2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -chg sta
                                    'Case "都市計画・用途地域1"
                                    '    '.Vari_Cityplan = fldvalue.Trim
                                    '    tmp_tosiyoto1 = fldvalue.Trim
                                    'Case "都市計画・用途地域2"
                                    '    '.Vari_Yototiki = fldvalue.Trim
                                    '    tmp_tosiyoto2 = fldvalue.Trim
                                Case "都市計画"
                                    tmp_tosiyoto1 = fldvalue.Trim
                                Case "用途地域"
                                    tmp_tosiyoto2 = fldvalue.Trim
                                    '2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -chg end
                                Case "管理形態-管理人名(SJIS)"
                                    .Vari_Kanrinin_namesjis = fldvalue.Trim
                                Case "管理形態-管理業者名(SJIS)"
                                    .Vari_Kanri_gynamesjis = fldvalue.Trim
                                Case "管理形態-管理業者担当者名(SJIS)"
                                    .Vari_Kanri_gytantosjis = fldvalue.Trim
                                Case "駐車場空き台数"
                                    .Vari_Parking_caraki = fldvalue.Trim
                                Case "バイク置き場空き台数"
                                    .Vari_Parking_bikeaki = fldvalue.Trim
                                Case "駐輪場空き台数"
                                    .Vari_Parking_bicycleaki = fldvalue.Trim
                                Case "小学校区"
                                    .Vari_Syogaku_name = fldvalue.Trim
                                Case "小学校距離"
                                    .Vari_Syogaku_kyori = fldvalue.Trim
                                Case "中学校区"
                                    .Vari_Cyugaku_name = fldvalue.Trim
                                Case "中学校距離"
                                    .Vari_Cyugaku_kyori = fldvalue.Trim
                                Case "エリア"
                                    'エリアマスタと照合
                                    .Vari_Bk_area_no = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 物件詳細情報 -add sta
                                Case "非常用エレベーターフラグ"
                                    .Vari_Emergencyelevator_flg = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 物件詳細情報 -add end
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory
                        .Vari_Krbunrui_no = -1
                        .Vari_Svbunrui_no = -1
                        .Vari_Isiwata_gy_sekono = Nothing

                        '2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -chg sta
                        ''都市計画/用途地域照合
                        'Dim tmp_tosi As String = ""
                        'Dim tmp_yoto As String = ""

                        'Call Me.Get_TosiYotoValue(tmp_tosiyoto1, tmp_tosiyoto2, tmp_tosi, tmp_yoto)

                        '.Vari_Cityplan = tmp_tosi
                        '.Vari_Yototiki = tmp_yoto

                        Dim tmp_tosi As String = ""
                        Dim tmp_yoto As String = ""

                        '都市計画・用途地域を判別
                        Call Me.Set_Tosiyotokbn(tmp_tosiyoto1, tmp_tosiyoto2, tmp_tosi, tmp_yoto)

                        .Vari_Cityplan = tmp_tosi
                        .Vari_Yototiki = tmp_yoto
                        '2016.04.06 都市計画・用途地域を紐データから取得するように修正 -chg end

                        '20161128 物件面積関連情報の追加 -add sta
                        If CNVNO = ConvertTypes._汎用 Then
                            .Vari_Men_nobeyukatubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_nobeyuka)), RoundingTypes.Kirisute, 2)
                            .Vari_Men_sikititubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_sikiti)), RoundingTypes.Kirisute, 2)
                            .Vari_Men_parkingtubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_parking)), RoundingTypes.Kirisute, 2)
                            .Vari_Men_yukatokitubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_yukatoki)), RoundingTypes.Kirisute, 2)
                            .Vari_Men_sikititokitubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_sikititoki)), RoundingTypes.Kirisute, 2)
                            .Vari_Men_kentikutubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_kentiku)), RoundingTypes.Kirisute, 2)
                            .Vari_Men_kentikutokitubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(.Vari_Men_kentikutoki)), RoundingTypes.Kirisute, 2)
                        End If
                        '20161128 物件面積関連情報の追加 -add end

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
                '返却
                Return rtn

            End Function

            ''' <summary>
            ''' 物件分類仮テーブルからハッシュテーブルを生成
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_HashBkRui(ByVal sqlcnnv10 As SqlConnection) As Hashtable

                Dim rtn_hash As New Hashtable
                Dim readtbl As New DataTable
                Dim reccnt As Integer
                Dim flg As Boolean
                Dim fldname As String = ""
                Dim fldvalue As String = ""
                Dim tmp_v7ruino As String = ""
                Dim tmp_v10ruino As String = ""

                '物件分類仮テーブル取得
                Dim tmp_sql As String = "SELECT * FROM tmp_bkrui_mst"
                reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, flg)

                For cntii As Integer = 0 To readtbl.Rows.Count - 1
                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                        '項目名取得
                        fldname = readtbl.Columns(cntjj).ColumnName

                        '登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

                        '変数セット
                        Select Case fldname
                            Case "移行元物件分類№"
                                tmp_v7ruino = fldvalue
                            Case "賃貸革命10物件分類№"
                                tmp_v10ruino = fldvalue
                        End Select

                    Next

                    'ハッシュテーブルセット
                    rtn_hash.Add(tmp_v7ruino, tmp_v10ruino)

                Next

                Return rtn_hash

            End Function

            ''' <summary>
            ''' エリアマスタからハッシュテーブルを生成(名称ではなく番号を設定するようになったため)
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_HashArea(ByVal sqlcnnv10 As SqlConnection) As Hashtable

                Dim rtn_hash As New Hashtable
                Dim readtbl As New DataTable
                Dim reccnt As Integer
                Dim flg As Boolean
                Dim fldname As String = ""
                Dim fldvalue As String = ""
                Dim tmp_areano As String = ""
                Dim tmp_areaname As String = ""

                'エリアマスタ取得
                Dim tmp_sql As String = "SELECT * FROM m_area"
                reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, flg)

                For cntii As Integer = 0 To readtbl.Rows.Count - 1
                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                        '項目名取得
                        fldname = readtbl.Columns(cntjj).ColumnName

                        '登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

                        '変数セット
                        Select Case fldname
                            Case "area_no"
                                tmp_areano = fldvalue
                            Case "area_name"
                                tmp_areaname = fldvalue
                        End Select

                    Next

                    'ハッシュテーブルセット
                    rtn_hash.Add(tmp_areaname, tmp_areano)

                Next

                Return rtn_hash

            End Function

            ''' <summary>
            ''' 構造仮マスタからハッシュテーブルを生成(番号を設定するため)
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_HashKozo(ByVal sqlcnnv10 As SqlConnection) As Hashtable

                Dim rtn_hash As New Hashtable
                Dim readtbl As New DataTable
                Dim reccnt As Integer
                Dim flg As Boolean
                Dim fldname As String = ""
                Dim fldvalue As String = ""
                Dim tmp_v7kozoname As String = ""
                Dim tmp_v10kozono As String = ""

                '構造マスタ仮テーブル取得
                Dim tmp_sql As String = "SELECT * FROM tmp_kozo_mst"
                reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, flg)

                For cntii As Integer = 0 To readtbl.Rows.Count - 1
                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                        '項目名取得
                        fldname = readtbl.Columns(cntjj).ColumnName

                        '登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

                        '変数セット
                        Select Case fldname
                            Case "移行元構造名称"
                                tmp_v7kozoname = fldvalue
                            Case "賃貸革命10構造№"
                                tmp_v10kozono = fldvalue
                        End Select

                    Next

                    'ハッシュテーブルセット
                    rtn_hash.Add(tmp_v7kozoname, tmp_v10kozono)

                Next

                Return rtn_hash

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

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

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

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

            '2016.04.06 紐付データ取得処理を外出し -del sta
            ' ''' <summary>
            ' ''' 物件分類紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_Bkrui()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "物件分類マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldbkruiname As String = ""
            '        Dim tmp_newbkruino As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元物件分類名称"
            '                    tmp_oldbkruiname = fldvalue.Trim
            '                Case "賃貸革命10物件分類No"
            '                    tmp_newbkruino = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newbkruino <> "" Then
            '            Hash_Rel_Bkrui.Add(tmp_oldbkruiname, tmp_newbkruino)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub

            ' ''' <summary>
            ' ''' 構造マスタ紐付情報取得
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Public Sub Get_RelData_Kozo()

            '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            '    Dim startrow As Integer                                         '書込開始行
            '    Dim columncnt As Integer                                        '列数
            '    Dim maxrowcnt As Integer                                        '既存データの行数
            '    Dim rowcnt As Integer                                           '書込行数
            '    Dim relsheetname As String = "構造マスタ"
            '    Dim rtn As Boolean = True

            '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            '    'Excelファイル初期設定
            '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            '    For cntii = 0 To rowcnt - 1

            '        '作業用変数作成
            '        Dim tmp_oldbkozoname As String = ""
            '        Dim tmp_newkozono As String = ""

            '        '紐付設定値取得
            '        For cntjj = 1 To columncnt

            '            'ヘッダー格納
            '            Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            '            '移行値格納
            '            Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            '            Select Case fldname
            '                Case "移行元構造名称"
            '                    tmp_oldbkozoname = fldvalue.Trim
            '                Case "賃貸革命10構造No"
            '                    tmp_newkozono = fldvalue.Trim
            '            End Select

            '        Next

            '        'ハッシュテーブル格納
            '        If tmp_newkozono <> "" Then
            '            Hash_Rel_Kozo.Add(tmp_oldbkozoname, tmp_newkozono)
            '        End If

            '    Next

            '    'Excelファイル終了設定
            '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            'End Sub
            '2016.04.06 紐付データ取得処理を外出し -del end

            '2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -del sta
            ' ''' <summary>
            ' ''' 用途地域の文字列を10の都市計画/用途地域と照合
            ' ''' </summary>
            ' ''' <param name="tosiyoto1"></param>
            ' ''' <param name="tosiyoto2"></param>
            ' ''' <param name="tosi"></param>
            ' ''' <param name="yoto"></param>
            ' ''' <remarks></remarks>
            'Public Sub Get_TosiYotoValue(ByVal tosiyoto1 As String, ByVal tosiyoto2 As String, ByRef tosi As String, ByRef yoto As String)

            '    Dim tmp_tosiyotokbn1 As Integer
            '    Dim tmp_tosiyotono1 As String = ""
            '    Dim tmp_tosiyotokbn2 As Integer
            '    Dim tmp_tosiyotono2 As String = ""

            '    TosiYotoConv.Get_TosiYoto(tosiyoto1, tmp_tosiyotokbn1, tmp_tosiyotono1)
            '    TosiYotoConv.Get_TosiYoto(tosiyoto2, tmp_tosiyotokbn2, tmp_tosiyotono2)

            '    Select Case True
            '        Case tmp_tosiyotokbn1 = 0 And tmp_tosiyotokbn2 = 0  '1個目が空白、2個目が空白
            '            tosi = -1
            '            yoto = -1
            '        Case tmp_tosiyotokbn1 = 0 And tmp_tosiyotokbn2 = 1  '1個目が空白、2個目が都市計画
            '            tosi = tmp_tosiyotono2
            '            yoto = -1
            '        Case tmp_tosiyotokbn1 = 0 And tmp_tosiyotokbn2 = 2  '1個目が空白、2個目が用途地域
            '            tosi = -1
            '            yoto = tmp_tosiyotono2
            '        Case tmp_tosiyotokbn1 = 0 And tmp_tosiyotokbn2 = 3  '1個目が空白、2個目が該当データ無し
            '            tosi = -1
            '            yoto = 99
            '        Case tmp_tosiyotokbn1 = 1 And tmp_tosiyotokbn2 = 0  '1個目が都市計画、2個目が空白
            '            tosi = tmp_tosiyotono1
            '            yoto = -1
            '        Case tmp_tosiyotokbn1 = 1 And tmp_tosiyotokbn2 = 1  '1個目が都市計画、2個目が都市計画
            '            tosi = tmp_tosiyotono1
            '            yoto = -1
            '        Case tmp_tosiyotokbn1 = 1 And tmp_tosiyotokbn2 = 2  '1個目が都市計画、2個目が用途地域
            '            tosi = tmp_tosiyotono1
            '            yoto = tmp_tosiyotono2
            '        Case tmp_tosiyotokbn1 = 1 And tmp_tosiyotokbn2 = 3  '1個目が都市計画、2個目が該当データ無し
            '            tosi = tmp_tosiyotono1
            '            yoto = 99
            '        Case tmp_tosiyotokbn1 = 2 And tmp_tosiyotokbn2 = 0  '1個目が用途地域、2個目が空白
            '            tosi = -1
            '            yoto = tmp_tosiyotono1
            '        Case tmp_tosiyotokbn1 = 2 And tmp_tosiyotokbn2 = 1  '1個目が用途地域、2個目が都市計画
            '            tosi = tmp_tosiyotono2
            '            yoto = tmp_tosiyotono1
            '        Case tmp_tosiyotokbn1 = 2 And tmp_tosiyotokbn2 = 2  '1個目が用途地域、2個目が用途地域
            '            tosi = -1
            '            yoto = tmp_tosiyotono1
            '        Case tmp_tosiyotokbn1 = 2 And tmp_tosiyotokbn2 = 3  '1個目が用途地域、2個目が該当データ無し
            '            tosi = 99
            '            yoto = tmp_tosiyotono1
            '        Case tmp_tosiyotokbn1 = 3 And tmp_tosiyotokbn2 = 0  '1個目が該当データ無し、2個目が空白
            '            tosi = 99
            '            yoto = -1
            '        Case tmp_tosiyotokbn1 = 3 And tmp_tosiyotokbn2 = 1  '1個目が該当データ無し、2個目が都市計画
            '            tosi = tmp_tosiyotono2
            '            yoto = 99
            '        Case tmp_tosiyotokbn1 = 3 And tmp_tosiyotokbn2 = 2  '1個目が該当データ無し、2個目が用途地域
            '            tosi = 99
            '            yoto = tmp_tosiyotono2
            '        Case tmp_tosiyotokbn1 = 3 And tmp_tosiyotokbn2 = 3  '1個目が該当データ無し、2個目が該当データ無し
            '            tosi = 99
            '            yoto = 99
            '    End Select

            'End Sub
            '2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -del end

            ''' <summary>
            ''' 都市計画・用途地域を判別する
            ''' </summary>
            ''' <param name="tosiyoto1"></param>
            ''' <param name="tosiyoto2"></param>
            ''' <param name="tosi"></param>
            ''' <param name="yoto"></param>
            ''' <remarks></remarks>
            Public Sub Set_Tosiyotokbn(ByVal tosiyoto1 As String, ByVal tosiyoto2 As String, ByRef tosi As String, ByRef yoto As String)

                '両方データが存在しない場合は処理を抜ける
                If tosiyoto1 = "" And tosiyoto2 = "" Then
                    Exit Sub
                End If

                Dim tmp_kbn1 As String = ""
                Dim tmp_kbn2 As String = ""

                '1つ目の都市用途区分を取得
                If Hash_Rel_Tosiyotokbn.Contains(tosiyoto1) Then
                    tmp_kbn1 = Hash_Rel_Tosiyotokbn(tosiyoto1)
                End If

                '2つ目の都市用途区分を取得
                If Hash_Rel_Tosiyotokbn.Contains(tosiyoto2) Then
                    tmp_kbn2 = Hash_Rel_Tosiyotokbn(tosiyoto2)
                End If

                Select Case True
                    Case tmp_kbn1 = "" And tmp_kbn2 = ""
                        tosi = ""
                        yoto = ""
                    Case tmp_kbn1 = "" And tmp_kbn2 = "都市計画"
                        tosi = tosiyoto2
                        yoto = ""
                    Case tmp_kbn1 = "" And tmp_kbn2 = "用途地域"
                        tosi = ""
                        yoto = tosiyoto2
                    Case tmp_kbn1 = "都市計画" And tmp_kbn2 = ""
                        tosi = tosiyoto1
                        yoto = ""
                    Case tmp_kbn1 = "都市計画" And tmp_kbn2 = "都市計画"
                        tosi = tosiyoto1
                        yoto = "都市計画の重複"    '重複した場合は移行不可のため判別用の文字列を入れておく
                    Case tmp_kbn1 = "都市計画" And tmp_kbn2 = "用途地域"
                        tosi = tosiyoto1
                        yoto = tosiyoto2
                    Case tmp_kbn1 = "用途地域" And tmp_kbn2 = ""
                        tosi = ""
                        yoto = tosiyoto1
                    Case tmp_kbn1 = "用途地域" And tmp_kbn2 = "都市計画"
                        tosi = tosiyoto2
                        yoto = tosiyoto1
                    Case tmp_kbn1 = "用途地域" And tmp_kbn2 = "用途地域"
                        tosi = "用途地域の重複"    '重複した場合は移行不可のため判別用の文字列を入れておく
                        yoto = tosiyoto1
                End Select


            End Sub

            ''' <summary>
            ''' 敷地利用-契約期間終了の文字列を日付型の文字列へ変換する処理 '20160706 敷地利用-契約期間終了の文字列日付変換処理の追加
            ''' 「平成○○年●●月まで」の形のみ変換対象とする(年号は明治、大正、昭和、平成)
            ''' </summary>
            ''' <param name="value"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Private Function Get_ChgDate(ByVal value As String) As String

                Dim rtn_str As String = ""
                Dim matchflg As Boolean = False
                Dim list_nengo As New List(Of String) From {"明治", "大正", "昭和", "平成"}
                Dim tmp_year As String = "年"
                Dim tmp_num As String = "[1-9]?[0-9]"
                Dim tmp_month As String = "月"
                Dim tmp_kotei As String = "まで"
                Dim chk_nengo As String = ""
                Dim tmp_value As String = ""

                '空文字チェック
                If value = "" Then
                    Return rtn_str
                End If

                '半角変換と空白文字列除去
                tmp_value = StrConv(value, VbStrConv.Narrow)
                tmp_value = tmp_value.Replace(" ", "")

                '形式チェック
                For Each nengo In list_nengo
                    Dim tmp_syogo As String = nengo & tmp_num & tmp_year & tmp_num & tmp_month & tmp_kotei
                    If System.Text.RegularExpressions.Regex.IsMatch(tmp_value, tmp_syogo) Then
                        chk_nengo = nengo
                        matchflg = True
                        Exit For
                    End If
                Next

                '形式が標準以外の場合は処理を抜ける
                If matchflg = False Then
                    '20160808 連動検証中に発見した不具合修正 不適合時は空文字を返す -del
                    'rtn_str = value
                    Return rtn_str
                End If

                '形式が標準の場合は数値チェックを行う
                Dim heisei_maxyear As Integer = Int32.Parse(Int32.Parse(Now.Year) - 1988)
                '20160711 レビュー後指摘対応 敷地利用処理の最大年数修正 -chg sta
                'Dim hash_nengomaxyear As New Hashtable From _
                '    {{"明治", 44}, {"大正", 14}, {"昭和", 63}, {"平成", heisei_maxyear}}
                Dim hash_nengomaxyear As New Hashtable From _
                    {{"明治", 45}, {"大正", 15}, {"昭和", 64}, {"平成", heisei_maxyear}}
                '20160711 レビュー後指摘対応 敷地利用処理の最大年数修正 -chg end
                Dim hash_nengotoseirekiaddyear As New Hashtable From _
                    {{"明治", 1867}, {"大正", 1911}, {"昭和", 1925}, {"平成", 1988}}
                Dim tmp_str() As String = tmp_value.Split("年")

                '年のチェック
                Dim chk_year As Integer = Int32.Parse(tmp_str(0).Replace(chk_nengo, ""))
                Dim max_year As Integer = hash_nengomaxyear(chk_nengo)

                If max_year < chk_year Then
                    '20160808 連動検証中に発見した不具合修正 不適合時は空文字を返す -del
                    'rtn_str = value
                    Return rtn_str
                End If

                '月のチェック
                '20160711 レビュー後指摘対応 敷地利用処理の変数表記 -chg sta
                'Dim chk_month As Integer = Int32.Parse(tmp_str(1).Replace("月まで", ""))
                Dim chk_month As Integer = Int32.Parse(tmp_str(1).Replace(tmp_month & tmp_kotei, ""))
                '20160711 レビュー後指摘対応 敷地利用処理の変数表記 -chg end
                Dim max_month As Integer = 12

                If max_month < chk_month Then
                    '20160808 連動検証中に発見した不具合修正 不適合時は空文字を返す -del
                    'rtn_str = value
                    Return rtn_str
                End If

                '全てのチェックに問題が無い場合は西暦(月末)に変換して返す
                Dim chg_seirekiyear As String = (chk_year + hash_nengotoseirekiaddyear(chk_nengo)).ToString
                Dim chg_month As String = IIf(chk_month.ToString.Length = 1, ("0" & chk_month).ToString, chk_month.ToString)
                Dim tmp_lastday As String = Date.DaysInMonth(chk_year, chk_month)

                Dim chg_date As String = chg_seirekiyear & "/" & chg_month & "/" & tmp_lastday

                rtn_str = chg_date

                Return rtn_str

            End Function

        End Class

    End Class

#End Region

#Region "物件所有者情報"

    Public Class Bkdata_syo_Repository

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
                Dim model_cvitem As New Njc.Model.Bkdata_syo_Model              '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_syo"
                Dim fldnamegrp As String = "bk_guid,kn_no,sorule_guid,kasi1_ow_no,syo1_ow_no," & _
                                           "kasi2_ow_no,syo2_ow_no,syo_startymd,syo_endymd,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件No"
                                    .Vari_Bk_guid = fldvalue.Trim
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,kn_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件ゴミ情報"

    Public Class Bkdata_dust_Repository

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
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Bkdata_dust_Model             '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_dust"
                Dim fldnamegrp As String = "bk_guid,dust_no,dust_rui,dust_youbi,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
                                Case "ゴミ情報No"
                                    .Vari_Dust_no = fldvalue.Trim
                                Case "ゴミ分類"
                                    .Vari_Dust_rui = fldvalue.Trim
                                Case "ゴミ出し曜日"
                                    .Vari_Dust_youbi = fldvalue.Trim
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

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,dust_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件変動費親メーター情報(汎用ツールのみ)"

    Public Class Bkdata_hendo_Repository

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
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Bkdata_hendo_Model            '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_hendo"
                Dim fldnamegrp As String = "bk_guid,bkhendo_guid,rec_no,hendo_kbn,meter_name," & _
                                           "nkin_no,child_meter,biko,useflg,history," & _
                                           "sq_mmkbn,tani,tanisjis"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
                                Case "行No"
                                    .Vari_Rec_no = fldvalue.Trim
                                Case "メーター分類 (水道/ガス/電気/灯油/その他)"
                                    .Vari_Hendo_kbn = fldvalue.Trim
                                Case "メーター名(部屋no・フロア名など)"
                                    .Vari_Meter_name = fldvalue.Trim
                                Case "変動費入金項目"
                                    .Vari_Nkin_no = fldvalue.Trim
                                Case "子メーター検針"
                                    .Vari_Child_meter = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko = fldvalue.Trim
                                Case "使用フラグ"
                                    .Vari_Useflg = fldvalue.Trim
                                Case "請求月"
                                    .Vari_Sq_mmkbn = fldvalue.Trim
                                Case "単位"
                                    .Vari_Tani = fldvalue.Trim
                                Case "単位(SJIS)"
                                    .Vari_Tanisjis = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_Bkhendo_guid = Guid.NewGuid.ToString
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

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,rec_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件鍵情報"

    Public Class Bkdata_kagi_Repository

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

                Dim model_cvitem As New Njc.Model.Bkdata_kagi_Model             '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub As Integer = 2                                   'サブキー列

                Dim cnt_cvcnt As Integer = 0                                    '20160531 鍵情報移行処理の修正 共用鍵の件数取得用 -add

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
                'テーブル名/フィールド名セット
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_kagi"
                Dim fldnamegrp As String = "bk_guid,kagi_no,honsu,biko,history," & _
                                           "hokan,gyshare"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件No"
                                    .Vari_Bk_guid = fldvalue
                                Case "鍵No"      '20160613 鍵情報の取得処理修正 -chg 鍵タイトル名→鍵No
                                    .Vari_Kagi_no = fldvalue
                                Case "鍵本数"
                                    .Vari_Honsu = fldvalue
                                Case "備考"
                                    .Vari_Biko = fldvalue
                                Case "保管場所"
                                    .Vari_Hokan = fldvalue
                                Case "業者間での情報共有"
                                    .Vari_Gyshare = fldvalue
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
                        ' ''    hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

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

                        ''If Hash_KagiTitleKyoyo.Contains(.Vari_Kagi_no) Then

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
                        ''        hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

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
                        '        hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

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

                                End If

                            Case ConvertTypes._汎用

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

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,kagi_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件権利情報"

    Public Class Bkdata_kenri_Repository

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
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Bkdata_kenri_Model            '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_kenri"
                Dim fldnamegrp As String = "bk_guid,kenri_no,other_kenrirui,other_kenribiko,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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

                        '2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add sta
                        '作業用変数
                        Dim tmp_kenriumu As String = ""
                        '2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add end

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
                                Case "権利情報No"
                                    .Vari_Kenri_no = fldvalue
                                    '2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add sta
                                Case "所有権以外の権利の有無"
                                    tmp_kenriumu = fldvalue
                                    '2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add end
                                Case "所有権以外の権利の種類"
                                    .Vari_Other_kenrirui = fldvalue
                                Case "内容"
                                    .Vari_Other_kenribiko = fldvalue
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
                            hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

                            '挿入処理
                            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                            '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                            If normalflg Then
                                list_chkduplicate.Add(fldvalue_key)
                                tmp_cvcnt = tmp_cvcnt + 1

                                '2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add sta
                                '更新処理 (所有権以外の権利有無のフラグはbkdata_detailに存在するため)
                                Dim tmpcnt As Integer = 0
                                Dim kenriflgupdatesql As String = Me.Get_UseQry_Update(hash_cvitem("bk_guid"), tmp_kenriumu)
                                DBExec.Exec_NonQuery(sqlcnnv10, kenriflgupdatesql, tmpcnt)
                                '2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add end

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

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            ''' <summary>
            ''' 既存データの取得
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <param name="list_existdata"></param>
            ''' <remarks></remarks>
            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,kenri_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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
            ''' [bkdata_detail].[other_kenriflg]の一括更新クエリ 2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Update(ByVal bkguid As String, ByVal updatevalue As String) As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " UPDATE bkdata_detail SET "
                tmp_sql = tmp_sql & " 	other_kenriflg = " & updatevalue
                tmp_sql = tmp_sql & " WHERE bk_guid = '" & bkguid & "' "
                tmp_sql = tmp_sql & " ; "

                Return tmp_sql

            End Function

        End Class

    End Class

#End Region

#Region "物件交通情報"

    Public Class Bkdata_kotu_Repository

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

                Dim model_cvitem As New Njc.Model.Bkdata_kotu_Model             '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_kotu"
                '20160519 EXEUpdateに伴う修正 物件交通情報 -del(kotu_syudanを削除して成形)
                Dim fldnamegrp As String = "bk_guid,ensen_cnt,ensen_no,eki_no,firsttrain_flg," & _
                                           "kyori,toho_min,car_min,bus_min,bus_company," & _
                                           "bus_station,bus_tohomin,bus_kyori,history,setting_kotutype," & _
                                           "bus_companytoeki,bus_stationtoeki,bus_kyoritoeki,bus_tohomintoeki,car_kyori"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
                                Case "沿線番号"
                                    .Vari_Ensen_cnt = fldvalue.Trim
                                Case "沿線No"
                                    .Vari_Ensen_no = fldvalue.Trim
                                Case "駅No"
                                    .Vari_Eki_no = fldvalue.Trim
                                Case "始発フラグ"
                                    .Vari_Firsttrain_flg = fldvalue.Trim
                                Case "距離"
                                    .Vari_Kyori = fldvalue.Trim
                                Case "徒歩"
                                    .Vari_Toho_min = fldvalue.Trim
                                Case "車"
                                    .Vari_Car_min = fldvalue.Trim
                                Case "バス"
                                    .Vari_Bus_min = fldvalue.Trim
                                Case "バス会社"
                                    .Vari_Bus_company = fldvalue.Trim
                                Case "バス停"
                                    .Vari_Bus_station = fldvalue.Trim
                                Case "バス停徒歩"
                                    .Vari_Bus_tohomin = fldvalue.Trim
                                Case "バス停までの距離"
                                    .Vari_Bus_kyori = fldvalue.Trim
                                Case "設定する交通のタイプ"
                                    .Vari_Setting_kotutype = fldvalue.Trim
                                Case "バス系統・路線名(駅までバスを利用)"
                                    .Vari_Bus_companytoeki = fldvalue.Trim
                                Case "バス停名(駅までバスを利用)"
                                    .Vari_Bus_stationtoeki = fldvalue.Trim
                                Case "バス停(駅までバスを利用)までの道路距離"
                                    .Vari_Bus_kyoritoeki = fldvalue.Trim
                                Case "バス停(駅までバスを利用)までの徒歩(分)"
                                    .Vari_Bus_tohomintoeki = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 物件交通情報 -del sta
                                    'Case "交通手段"
                                    '    .Vari_Kotu_syudan = fldvalue.Trim
                                    '20160519 EXEUpdateに伴う修正 物件交通情報 -del end
                                Case "車距離"
                                    .Vari_Car_kyori = fldvalue.Trim
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,ensen_cnt) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件近隣駐車場情報(汎用ツールのみ)"

    Public Class Bkdata_parkingother_Repository

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

                Dim model_cvitem As New Njc.Model.Bkdata_parkingother_Model     '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_parkingother"
                Dim fldnamegrp As String = "bk_guid,parking_no,naiyo,kyori,gak," & _
                                           "zeikbn"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
                                Case "駐車場No"
                                    .Vari_Parking_no = fldvalue.Trim
                                Case "駐車場名"
                                    .Vari_Naiyo = fldvalue.Trim
                                Case "距離(m)"
                                    .Vari_Kyori = fldvalue.Trim
                                Case "料金(月額)"
                                    .Vari_Gak = fldvalue.Trim
                                Case "税区分"
                                    .Vari_Zeikbn = fldvalue.Trim
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,parking_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件参照ファイル情報(汎用ツールのみ)"

    Public Class Bkdata_relfile_Repository

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
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Bkdata_relfile_Model          '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_relfile"
                Dim fldnamegrp As String = "bk_guid,file_no,fullpath,addtime,biko," & _
                                           "history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
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

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,file_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件接道情報"

    Public Class Bkdata_setudo_Repository

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
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Bkdata_setudo_Model           '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_setudo"
                Dim fldnamegrp As String = "bk_guid,setudo_cnt,setudo_muki,setudo_roadpattern,setudo_haba," & _
                                           "setudo_setudokyori,setudo_pointselectflg,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
                                Case "接道No"
                                    .Vari_Setudo_cnt = fldvalue.Trim
                                Case "方角"
                                    .Vari_Setudo_muki = fldvalue.Trim
                                Case "道路種"
                                    .Vari_Setudo_roadpattern = fldvalue.Trim
                                Case "幅員"
                                    .Vari_Setudo_haba = fldvalue.Trim
                                Case "接道距離"
                                    .Vari_Setudo_setudokyori = fldvalue.Trim
                                Case "位置指定道路"
                                    .Vari_Setudo_pointselectflg = fldvalue.Trim
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

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,setudo_cnt) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件周辺情報"

    Public Class Bkdata_syuhen_Repository

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
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Bkdata_syuhen_Model           '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_syuhen"
                Dim fldnamegrp As String = "syuhen_guid,bk_guid,syuhen_cnt,sisetu_kbn,sisetu_name," & _
                                           "sisetu_kyori,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
                                Case "周辺No"
                                    .Vari_Syuhen_cnt = fldvalue.Trim
                                Case "周辺施設分類"
                                    .Vari_Sisetu_kbn = fldvalue.Trim
                                Case "周辺施設名"
                                    .Vari_Sisetu_name = fldvalue.Trim
                                Case "周辺施設距離"
                                    '2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化 -chg sta
                                    '↓↓↓旧srcコメントアウト↓↓↓
                                    ''2016.04.01 変換時の型落ち対応 -chg sta
                                    ''Dim tmpsyuhenkyoristr As String = fldvalue.Trim
                                    ''Dim tmpsyuhenkyoridbl As Double = Double.Parse(tmpsyuhenkyoristr)       'kakaka4 0322_1300 型落ち考慮
                                    ' ''四捨五入して取得                               'kakaka2 四捨五入は汎用や他社用の為に用意？
                                    ''.Vari_Sisetu_kyori = Math.Round(tmpsyuhenkyoridbl, MidpointRounding.AwayFromZero).ToString
                                    'Dim tmpsyuhenkyoristr As String = fldvalue.Trim
                                    'Dim tmpsyuhenkyoridbl As Double = 0
                                    'If Double.TryParse(tmpsyuhenkyoristr, tmpsyuhenkyoridbl) Then
                                    '    '四捨五入して取得
                                    '    .Vari_Sisetu_kyori = Math.Round(tmpsyuhenkyoridbl, MidpointRounding.AwayFromZero).ToString
                                    'Else
                                    '    '変換できない場合は値をそのまま格納 (ログ出力させる)
                                    '    .Vari_Sisetu_kyori = tmpsyuhenkyoristr
                                    'End If
                                    ''2016.04.01 変換時の型落ち対応 -chg end
                                    '↑↑↑旧srcコメントアウト↑↑↑
                                    .Vari_Sisetu_kyori = EtcMethod.Get_RoundingValue(fldvalue.Trim)
                                    '2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化 -chg end
                            End Select

                        Next

                        '固定値
                        .Vari_Syuhen_guid = Guid.NewGuid.ToString
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

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,syuhen_cnt) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件修繕維持管理連絡先情報"

    Public Class Bkdata_szeniji_Repository

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
                Dim hash_guid As New Hashtable                                  'guid格納用ハッシュテーブル
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.Bkdata_szeniji_Model          '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_szeniji"
                Dim fldnamegrp As String = "bk_guid,syuzenijikanri_no,syuzenijikanri_kasyo,syuzenijikanri_taisyokbn,syuzenijikanri_gyno," & _
                                           "syuzenijikanri_simei,syuzenijikanri_address,syuzenijikanri_tel,syuzenijikanri_jisyano," & _
                                           "syuzenijikanri_kasino,syuzenijikanri_syono,syuzenijikanri_simeiu"


                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                                Case "物件NO"
                                    .Vari_Bk_guid = fldvalue.Trim
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

                Dim tmp_sql As String = " SELECT bk_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

            End Sub

            Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

            End Function

            Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,syuzenijikanri_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "物件メモ情報"

    Public Class Bkdata_memo_Repository

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

                Dim model_cvitem As New Njc.Model.Bkdata_memo_Model             '移行値格納用モデル初期化
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
                Dim tblname_base As String = "bkdata"
                Dim tblname As String = "bkdata_memo"
                Dim fldnamegrp As String = "bk_guid,memo_no,memo,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                'guidを取得
                Call Me.Get_Guid(sqlcnnv10, tblname_base, hash_guid)

                '親マスタ取得
                Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

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
                        .Vari_Bk_guid = tmp_keymain

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
                                Dim log_key As String = "bkdata_memo-bk_guid"
                                Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
                                hash_keylog.Clear()
                                hash_keylog.Add(log_key, log_value)
                            End If

                            '重複チェック
                            If keychkflg And list_chkduplicate.Contains(tmp_keymain) Then
                                keychkflg = False
                                'ログ出力
                                Dim log_key As String = "bkdata_memo-bk_guid"
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

                                            'キーをguidへ変換
                                            hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,memo_no) FROM " & tblname & " LEFT JOIN bkdata ON " & tblname & ".bk_guid = bkdata.bk_guid "
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

        End Class

    End Class

#End Region

#Region "×物件KeyValueテーブル"

    '汎用ツールのみ
    '汎用は後で対応することとなったため今は対応しない

#End Region

#Region "×物件画像情報"

    '画像は別途行うためここではコンバートしない

#End Region

#Region "×物件周辺画像情報"

    '画像は別途行うためここではコンバートしない

#End Region

#Region "×物件bkdata_youtube情報"

    '使用箇所が不明のため保留

#End Region

End Namespace


