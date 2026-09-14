Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common
Imports System.Data.OleDb

Namespace Njc.Repository

#Region "物件分類マスタ"

    '20160525 紐付項目の挿入処理の追加 -del sta
    'Public Class M_bk_rui_Repository
    '    '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    '    'Public Class SubConv
    '    '    Implements IConv

    '    '    ''' <summary>
    '    '    ''' 【変数→DB(INSERT)】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <param name="model_bk_rui"></param>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Cnv_Db(ByVal sqlcnnv10 As SqlConnection, ByVal model_cvitem As Object, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Cnv_Db

    '    '        Dim reptbl As String                                '置換TBL名
    '    '        Dim repfld As String                                '置換項目名
    '    '        Dim repprm As String                                '置換パラメータ
    '    '        Dim rowcnt As Integer                               '書込行数
    '    '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    '    '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    '    '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    '    '        '書込行数取得
    '    '        rowcnt = model_cvitem.ItemCnt

    '    '        'プログレスバー初期化
    '    '        Dim pgbtotalcnt As Integer = rowcnt
    '    '        Call obj_pgb.pgbInitPart(pgbtotalcnt)

    '    '        'テーブル/フィールド名取得
    '    '        reptbl = "m_bk_rui"
    '    '        repfld = "bk_ruino,bk_ruiname,bk_ruibiko,bk_ruiuseflg,sincyoku_keiyakukbn," & _
    '    '                 "sincyoku_kosinkbn,sincyoku_kaiyakukbn,history,homemate_ruikbn,bk_endofmonthflg," & _
    '    '                 "ikkatukariage_siwakekbn"

    '    '        With model_cvitem

    '    '            For cntii = 0 To rowcnt - 1

    '    '                '中断処理
    '    '                Application.DoEvents()
    '    '                If CancelFlg Then
    '    '                    rtn = False
    '    '                    Return rtn
    '    '                End If

    '    '                If .Vari_tmp_skipflg(cntii) = False Then

    '    '                    'ハッシュテーブル作成
    '    '                    Dim hash_cvitem As New Hashtable
    '    '                    'hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(repfld, model_cvitem, cntii)

    '    '                    '書込処理
    '    '                    'Call CVDBInsert.CVitem_Insert(sqlcnnv10, reptbl, repfld, hash_cvitem)

    '    '                End If

    '    '                'プログレスバー更新/進捗率表示
    '    '                Dim pgbcnt As Integer = cntii + 1
    '    '                Call obj_pgb.pgbsettingPart(pgbcnt)
    '    '                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)

    '    '            Next

    '    '        End With

    '    '        'Update実行モジュールへ遷移
    '    '        Call Cnv_Db_Sub(sqlcnnv10)

    '    '        '返却
    '    '        Return rtn

    '    '    End Function

    '    '    ''' <summary>
    '    '    ''' 【移行先DB初期化】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <remarks></remarks>
    '    '    Public Sub Initialize_Table(sqlcnnv10 As SqlConnection) Implements IConv.Initialize_Table

    '    '        '初期化対象外

    '    '    End Sub

    '    '    ''' <summary>
    '    '    ''' 【仮テーブル→変数】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <param name="syorikomok"></param>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Set_Vari2(ByVal sheetname As String, ByRef model_cvitem As Object, ByVal sqlcnnv10 As SqlConnection) As Boolean Implements IConv.Set_Vari2

    '    '        Dim reccnt As Integer                               '抽出レコード件数
    '    '        Dim fldname As String                               '該当項目名
    '    '        Dim fldvalue As String                              '該当値
    '    '        Dim readtbl As New DataTable                        'テーブル格納用
    '    '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    '    '        Dim pgbcnt As Integer = 1                           'プログレスバー更新用カウンター
    '    '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    '    '        Dim logvalue As String                              'ログ出力用
    '    '        Dim tblname As String = "tmp_bkrui_mst"             '対象テーブル名
    '    '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    '    '        '移行元情報取得
    '    '        reccnt = DBExec.Exec_DataTable(Me.Get_UseQry(), sqlcnnv10, readtbl, rtn)

    '    '        '移行元情報取得時にエラーが生じた場合は処理を抜ける
    '    '        If Not rtn Then
    '    '            Return rtn
    '    '        End If

    '    '        '初期化
    '    '        model_cvitem = New Njc.Model.M_bk_rui_Model(reccnt - 1)
    '    '        Dim tmp_duplicatechk As New Njc.Model.CommonModel(reccnt - 1)

    '    '        'プログレスバー初期化
    '    '        Dim pgbtotalcnt As Integer = (readtbl.Rows.Count) * (readtbl.Columns.Count)
    '    '        Call obj_pgb.pgbInitPart(pgbtotalcnt)

    '    '        '==========================
    '    '        'データ取得　→　変数格納
    '    '        '==========================
    '    '        With model_cvitem

    '    '            For cntii As Integer = 0 To readtbl.Rows.Count - 1

    '    '                'DoEvents
    '    '                Application.DoEvents()
    '    '                If CancelFlg Then
    '    '                    rtn = False
    '    '                    Return rtn
    '    '                End If

    '    '                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

    '    '                    '項目名取得
    '    '                    fldname = readtbl.Columns(cntjj).ColumnName

    '    '                    '登録値取得
    '    '                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

    '    '                    'ログ用
    '    '                    logvalue = "賃貸革命10物件分類№ = " & Typ.ToStr(readtbl.Rows(cntii).Item(3)).Trim & _
    '    '                               " " & _
    '    '                               "移行元物件分類名称 = " & Typ.ToStr(readtbl.Rows(cntii).Item(1)).Trim

    '    '                    '各項目値→変数格納
    '    '                    Select Case fldname
    '    '                        Case "賃貸革命10物件分類№"
    '    '                            .Vari_Bk_ruino(cntii) = DataChk.Chk_DataInt(fldvalue, 1, 9999)
    '    '                        Case "移行元物件分類名称"
    '    '                            .Vari_Bk_ruiname(cntii) = DataChk.Chk_DataString(fldvalue, 50)
    '    '                        Case "移行元物件分類備考"
    '    '                            .Vari_Bk_ruibiko(cntii) = DataChk.Chk_DataString(fldvalue, 50)
    '    '                    End Select

    '    '                    '固定値設定
    '    '                    .Vari_UseFlg(cntii) = 1
    '    '                    .Vari_SincyokuKeiyakuKbn(cntii) = 0
    '    '                    .Vari_SincyokuKosinKbn(cntii) = 0
    '    '                    .Vari_SincyokuKaiyakuKbn(cntii) = 0
    '    '                    .Vari_HomemateRuiKbn(cntii) = 0
    '    '                    .Vari_History(cntii) = DefHistory
    '    '                    .Vari_tmp_skipflg(cntii) = 0

    '    '                    'プログレスバー更新/進捗率表示
    '    '                    Call obj_pgb.pgbsettingPart(pgbcnt)
    '    '                    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
    '    '                    pgbcnt = pgbcnt + 1

    '    '                Next

    '    '            Next

    '    '        End With

    '    '        '返却
    '    '        Return rtn

    '    '    End Function

    '    '    ''' <summary>
    '    '    ''' 【変数→DB(UPDATE)】
    '    '    ''' </summary>
    '    '    ''' <remarks></remarks>
    '    '    Public Sub Cnv_Db_Sub(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection)

    '    '        'コマンドセット
    '    '        Dim tmpcomtxt As String = Get_UseQry_Update()
    '    '        Dim tmpsqlcom As New SqlCommand(tmpcomtxt, sqlcnnv10)
    '    '        Dim rowsAffected As Integer = tmpsqlcom.ExecuteNonQuery()

    '    '        '終了処理
    '    '        tmpsqlcom.Dispose()

    '    '    End Sub

    '    '    ''' <summary>
    '    '    ''' 【抽出クエリ】2015.07.17 sol レビュー後修正_レビュー№312
    '    '    ''' </summary>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Get_UseQry(ByRef sortkeycnt As Integer) As String Implements IConv.Get_UseQry

    '    '        Dim tmp_sql As String = ""

    '    '        tmp_sql = tmp_sql & " SELECT "
    '    '        tmp_sql = tmp_sql & " 	* "
    '    '        tmp_sql = tmp_sql & " FROM tmp_bkrui_mst "
    '    '        tmp_sql = tmp_sql & " WHERE [賃貸革命10物件分類№] NOT IN (SELECT bk_ruino FROM m_bk_rui) "
    '    '        tmp_sql = tmp_sql & " AND   [賃貸革命10物件分類№] <> 0 "
    '    '        tmp_sql = tmp_sql & " ORDER BY [移行元物件分類№] "

    '    '        Return tmp_sql

    '    '    End Function

    '    '    ''' <summary>
    '    '    ''' 【更新クエリ】
    '    '    ''' </summary>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Get_UseQry_Update() As String

    '    '        Dim tmp_sql As String = ""

    '    '        tmp_sql = tmp_sql & " UPDATE m_bk_rui "
    '    '        tmp_sql = tmp_sql & " 	SET bk_ruibiko = TMP.移行元物件分類備考 "
    '    '        tmp_sql = tmp_sql & " FROM tmp_bkrui_mst AS TMP "
    '    '        tmp_sql = tmp_sql & " LEFT JOIN m_bk_rui AS MBKR "
    '    '        tmp_sql = tmp_sql & " ON TMP.[賃貸革命10物件分類№] = MBKR.bk_ruino; "

    '    '        Return tmp_sql

    '    '    End Function

    '    '    Public Function Get_BaseKey() As Object Implements IConv.Get_BaseKey

    '    '    End Function

    '    '                Public Function Set_Vari( _
    '    'ByVal sqlcnnv10 As SqlConnection, _
    '    'ByVal filename As String, ByVal sheetname As String, _
    '    'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    '    ') As Boolean Implements IConv.Set_Vari

    '    '    End Function

    '    '                Public Function Set_Vari( _
    '    'ByVal sqlcnnv10 As SqlConnection, _
    '    'ByVal filename As String, ByVal sheetname As String, _
    '    'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    '    ') As Boolean Implements IConv.Set_Vari

    '    '    End Function

    '    '    Public Function Write_IntermediateFile(sqlcnnv7 As SqlConnection, filename As String, sheetname As String, model_cvitem As Object) As Boolean Implements IConv.Write_IntermediateFile

    '    '    End Function

    '    'End Class
    '    '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    'End Class
    '20160525 紐付項目の挿入処理の追加 -del end

#End Region

#Region "バス交通マスタ"

    Public Class M_buskotu_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】20161004 既存中間→DB書込処理修正
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
                '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値
                Dim model_cvitem As New Njc.Model.M_buskotu_Model               '移行値格納用モデル初期化
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
                'プログレスバー初期化
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_buskotu"
                Dim fldnamegrp As String = "buskotu_no,buskotu_name,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '読込処理
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
                                Case "バス交通No"
                                    .Vari_Buskotu_no = fldvalue.Trim
                                Case "バス会社名"
                                    .Vari_Buskotu_name = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim errstr As String = ""
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If skipflg = False Then

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
            ''' 【中間ファイル→変数】旧処理のバックアップ
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="filename"></param>
            ''' <param name="sheetname"></param>
            ''' <param name="midrowcnt"></param>
            ''' <param name="cvrowcnt"></param>
            ''' <param name="conditioncnt"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Set_Vari_back(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean

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

                Dim model_cvitem As New Njc.Model.M_buskotu_Model               '移行値格納用モデル初期化
                Dim keycol As Integer = 1                                       'キー列

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定                          'kakaka バス交通マスタ以外も共通メソッドで纏められそうな場合はまとめて下さい。
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Return rtn
                End If

                'プログレスバー初期化
                '2016.02.22 プログレスバーの表示修正 -chg sta
                'Dim pgbtotalcnt As Integer = rowcnt
                'Call obj_pgb.pgbInitPart(pgbtotalcnt)
                Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
                Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値
                '2016.03.28 プログレスバーの表示修正(再) -del
                'Dim pgbsimpleflg As Boolean = False     '実件数で表示するか否かのフラグ (True:基準値で割った件数 False:実件数)

                If rowcnt <= pgbbasecnt Then
                    pgbtotalcnt = rowcnt
                Else
                    pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
                    '2016.03.28 プログレスバーの表示修正(再) -del
                    'pgbsimpleflg = True
                End If
                Call obj_pgb.pgbInitPart(pgbtotalcnt)
                '2016.02.22 プログレスバーの表示修正 -chg end

                'テーブル名/フィールド名セット
                Dim tblname As String = "m_buskotu"
                Dim fldnamegrp As String = "buskotu_no,buskotu_name,history"

                '追加コンバート時の重複チェック用に既存データのキーを取得
                If Not InitDBFlg Then
                    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
                End If

                '************************
                '処理開始
                '************************

                With model_cvitem

                    '---------------
                    'ヘッダー処理
                    '---------------
                    'ヘッダー行取得
                    Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value    '2016.02.22 移行項目取得方法修正

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
                        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value   '2016.02.22 移行項目取得方法修正

                        'キー値取得
                        Dim fldvalue_key As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))

                        'ログ出力用
                        Dim str_logkey As String = fldname_key & " = " & fldvalue_key

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim  '2016.02.22 移行項目取得方法修正
                            Dim fldvalue As String = ""                                             '2016.02.22 移行項目取得方法修正
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "バス交通No"
                                    .Vari_Buskotu_no = fldvalue.Trim
                                Case "バス会社名"
                                    .Vari_Buskotu_name = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_History = DefHistory

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
                        '2016.03.28 プログレスバーの表示修正(再) -chg sta
                        '↓↓↓旧srcコメントアウト↓↓↓
                        ''2016.02.22 プログレスバーの表示修正 -chg sta
                        ''Dim pgbcnt As Integer = cntii + 1
                        ''Call obj_pgb.pgbsettingPart(pgbcnt)                         'kakaka レスポンス改善の為、以下参考(他の情報(ブロック)も同様)
                        ''Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)            'kakaka 進捗間隔(進み方)を次の通りに変更 ⇒ 実件数/100<=1…実件数、実件数/100>1…実件数/100の切り上げ(実件数が1015の場合、10.15なので11回)

                        ''件数取得
                        'Dim pgbcnt As Integer = 0
                        'If pgbsimpleflg = False Then        '100件以下の場合は実件数を取得                                                  'kakaka4 0322_2050 どっちでもいいですが、rowcnt<=pgbbasecnt でもよかったような…
                        '    pgbcnt = cntii + 1
                        'ElseIf rowcnt <> cntii + 1 Then     '100件より大きい、かつ最終レコードに達していない場合は100件毎に値を取得         'kakaka4 0322_2050 これだと比較値+1ではない場合って意味では？上から順に条件を考えた場合でも本来の条件の意味が違うような…(rowcnt>pgbbasecnt AND rowcnt<=cntii )
                        '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                        'ElseIf rowcnt = cntii + 1 Then      '100件より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得    'kakaka 0322_20504 これだと比較値+1と同じって意味では？(rowcnt>pgbbasecnt AND rowcnt>cntii )
                        '    pgbcnt = pgbtotalcnt
                        'End If

                        ''表示
                        'If pgbcnt <> 0 Then
                        '    Call obj_pgb.pgbsettingPart(pgbcnt)
                        '    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                        'End If
                        ''2016.02.22 プログレスバーの表示修正 -chg end
                        '↑↑↑旧srcコメントアウト↑↑↑

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

                        '2016.03.28 プログレスバーの表示修正(再) -chg end
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

                Dim tmp_sql As String = " SELECT buskotu_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "バス停マスタ"

    Public Class M_buskotu_stop_Repository

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
                '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用   
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値
                Dim model_cvitem As New Njc.Model.M_buskotu_stop_Model          '移行値格納用モデル初期化
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
                Dim tblname_base As String = "m_buskotu"
                Dim tblname As String = "m_buskotu_stop"
                Dim fldnamegrp As String = "buskotu_no,sortorder,keito_name,busstop_name"

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
                '読込処理
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
                                Case "バス交通No"
                                    .Vari_Buskotu_no = fldvalue.Trim
                                Case "バス停No"
                                    .Vari_Sortorder = fldvalue.Trim
                                Case "系統名"
                                    .Vari_Keito_name = fldvalue.Trim
                                Case "バス停名"
                                    .Vari_Busstop_name = fldvalue.Trim
                            End Select

                        Next

                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                        'データチェック
                        Dim skipflg As Boolean = False
                        Dim hash_cvitem As New Hashtable
                        Dim errstr As String = ""
                        Dim hash_log As New Hashtable
                        skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                        '書込処理
                        If skipflg = False Then

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

            '    Dim model_cvitem As New Njc.Model.M_buskotu_stop_Model          '移行値格納用モデル初期化
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
            '    Dim tblname_base As String = "m_buskotu"
            '    Dim tblname As String = "m_buskotu_stop"
            '    Dim fldnamegrp As String = "buskotu_no,sortorder,keito_name,busstop_name"

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
            '                    Case "バス交通No"
            '                        .Vari_Buskotu_no = fldvalue.Trim
            '                    Case "バス停No"
            '                        .Vari_Sortorder = fldvalue.Trim
            '                    Case "系統名"
            '                        .Vari_Keito_name = fldvalue.Trim
            '                    Case "バス停名"
            '                        .Vari_Busstop_name = fldvalue.Trim
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

            ''' <summary>
            ''' 親マスタ取得→リスト格納
            ''' </summary>
            ''' <param name="sqlcnnv10"></param>
            ''' <param name="tblname"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

                Dim tmp_sql As String = " SELECT buskotu_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,buskotu_no) + '-' + CONVERT(varchar,sortorder) FROM  " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "鍵タイトルマスタ"

    Public Class M_kagi_title_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【中間ファイル→変数】 2016.04.06 鍵情報の取得修正 (メソッド内をほぼ全て修正)
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

                Dim model_cvitem As New Njc.Model.M_kagi_title_Model            '移行値格納用モデル初期化
                '20160613 鍵情報の取得処理修正 -chg sta
                'Dim keycol As Integer = 2                                       'メインキー列 (共用専用で振り分けることでNoが変化するため名称をキーにしておく)
                Dim keycol As Integer = 1                                       'メインキー列 (Noをキーに変更)
                '20160613 鍵情報の取得処理修正 -chg end

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定 (紐付ファイルの読込)
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, sheetname)

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
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
                                Case "共用チェック"
                                    .Vari_Kagi_kbn = IIf(fldvalue = "True", 1, 2)
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

                '鍵No一括成形
                Dim tmpcnt As Integer = 0
                Dim kaginoupdatesql As String = Me.Get_UseQry_Update()
                DBExec.Exec_NonQuery(sqlcnnv10, kaginoupdatesql, tmpcnt)

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

                '20160516 鍵タイトルマスタの重複チェック修正 -chg sta
                'Dim tmp_sql As String = " SELECT CONVERT(varchar,kagi_kbn) + '-' + CONVERT(varchar,kagi_no) FROM  " & tblname
                Dim tmp_sql As String = " SELECT kagi_name FROM  " & tblname
                '20160516 鍵タイトルマスタの重複チェック修正 -chg end
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

            ''' <summary>
            ''' 鍵Noの一括更新クエリ 2016.04.06 鍵情報の取得修正 (新規追加)
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function Get_UseQry_Update() As String

                Dim tmp_sql As String = ""

                tmp_sql = tmp_sql & " UPDATE m_kagi_title SET "
                tmp_sql = tmp_sql & " 	m_kagi_title.kagi_no = MKT.[連番] "
                tmp_sql = tmp_sql & " FROM m_kagi_title "
                tmp_sql = tmp_sql & " LEFT JOIN "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 kagi_kbn "
                tmp_sql = tmp_sql & " 			,kagi_no "
                tmp_sql = tmp_sql & " 			,ROW_NUMBER()OVER(PARTITION BY kagi_kbn ORDER BY kagi_kbn,kagi_no) AS [連番] "
                tmp_sql = tmp_sql & " 		FROM m_kagi_title "
                tmp_sql = tmp_sql & " 	)  AS MKT "
                tmp_sql = tmp_sql & " ON  m_kagi_title.kagi_kbn = MKT.kagi_kbn "
                tmp_sql = tmp_sql & " AND m_kagi_title.kagi_no = MKT.kagi_no "
                tmp_sql = tmp_sql & " ; "

                Return tmp_sql

            End Function

        End Class

    End Class

#End Region

#Region "箇所クレーム分類マスタ"

    Public Class M_claim_rui_Repository

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

                Dim model_cvitem As New Njc.Model.M_claim_rui_Model       '移行値格納用モデル初期化
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
                Dim tblname As String = "m_claim_rui"
                Dim fldnamegrp As String = "claim_ruikbn,claim_ruino,claim_ruisortorder,claim_name,claim_useflg," & _
                                           "history"

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
                                Case "箇所・クレーム分類区分"
                                    .Vari_Claim_ruikbn = fldvalue.Trim
                                Case "分類No"
                                    .Vari_Claim_ruino = fldvalue.Trim
                                    .Vari_Claim_ruisortorder = fldvalue.Trim
                                Case "名称"
                                    .Vari_Claim_name = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_Claim_useflg = 1
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
                                list_chkduplicate.Add(fldvalue_key)                             'kakaka メモ：Cnv_DbでINSERT失敗の場合、ここで重複データを貯めこむ
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

                Dim tmp_sql As String = " SELECT claim_ruino FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "特約マスタ"

    Public Class M_tokuyaku_Repository

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
                '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用  
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値
                Dim model_cvitem As New Njc.Model.M_tokuyaku_Model     '移行値格納用モデル初期化
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
                Dim tblname As String = "m_tokuyaku"
                Dim fldnamegrp As String = "tokuyaku_grpno,tokuyaku_no,tokuyaku_title,tokuyaku_template,history"

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
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

                            '各項目値→変数格納
                            Select Case fldname
                                Case "特約区分"
                                    .Vari_Tokuyaku_grpno = fldvalue
                                Case "特約No"
                                    .Vari_Tokuyaku_no = fldvalue
                                Case "特約タイトル"
                                    .Vari_Tokuyaku_title = fldvalue
                                Case "特約詳細"
                                    .Vari_Tokuyaku_template = fldvalue
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

            '    Dim model_cvitem As New Njc.Model.M_tokuyaku_Model     '移行値格納用モデル初期化
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
            '    Dim tblname As String = "m_tokuyaku"
            '    Dim fldnamegrp As String = "tokuyaku_grpno,tokuyaku_no,tokuyaku_title,tokuyaku_template,history"

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
            '                    '20160706 特約マスタの文字列移行処理修正 -chg sta
            '                    'fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString
            '                    '20160706 特約マスタの文字列移行処理修正 -chg end
            '                End If

            '                Select Case fldname
            '                    Case "特約区分"
            '                        .Vari_Tokuyaku_grpno = fldvalue
            '                    Case "特約No"
            '                        .Vari_Tokuyaku_no = fldvalue
            '                    Case "特約タイトル"
            '                        .Vari_Tokuyaku_title = fldvalue
            '                    Case "特約詳細"
            '                        .Vari_Tokuyaku_template = fldvalue
            '                End Select

            '            Next

            '            '固定値
            '            .Vari_History = DefHistory

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

                Dim tmp_sql As String = " SELECT tokuyaku_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "契約分類マスタ"

    '20160525 紐付項目の挿入処理の追加 -del sta
    'Public Class M_ky_rui_Repository

    '    Public Class SubConv
    '        Implements IConv

    '        ''' <summary>
    '        ''' 【中間ファイル→変数】
    '        ''' </summary>
    '        ''' <param name="sqlcnnv10"></param>
    '        ''' <param name="syorikomok"></param>
    '        ''' <returns></returns>
    '        ''' <remarks></remarks>
    '        Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

    '            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
    '            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
    '            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
    '            Dim startrow As Integer                                         '書込開始行
    '            Dim columncnt As Integer                                        '列数
    '            Dim maxrowcnt As Integer                                        '既存データの行数
    '            Dim rowcnt As Integer                                           '書込行数
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
    '            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
    '            Dim rtn As Boolean = True                                       '戻り値

    '            Dim chgcolor As New Njc.Common.Translate_Color                  '色変換用
    '            Dim model_cvitem As New Njc.Model.M_ky_rui_Model                '移行値格納用モデル初期化
    '            Dim keycol As Integer = 1                                       'キー列

    '            '************************
    '            '作業準備
    '            '************************

    '            'Excelファイル初期設定
    '            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    '            'Excelファイル設定時にエラーが生じた際は処理を抜ける
    '            If Not rtn Then
    '                Return rtn
    '            End If

    '            'テーブル名/フィールド名セット
    '            Dim tblname As String = "m_ky_rui"
    '            Dim fldnamegrp As String = "ky_ruino,ky_ruiname,ky_ruibiko,ky_ruitutimm,ky_ruicolor," & _
    '                                       "teisyaku_flg,history,useflg"

    '            '追加コンバート時の重複チェック用に既存データのキーを取得
    '            If Not InitDBFlg Then
    '                Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
    '            End If

    '            'プログレスバー初期化
    '            Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
    '            Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

    '            If rowcnt <= pgbbasecnt Then
    '                pgbtotalcnt = rowcnt
    '            Else
    '                pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
    '            End If
    '            Call obj_pgb.pgbInitPart(pgbtotalcnt)

    '            '************************
    '            '処理開始
    '            '************************
    '            With model_cvitem

    '                '---------------
    '                'ヘッダー処理
    '                '---------------
    '                'ヘッダー行取得
    '                Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    '                'キーヘッダー名取得
    '                Dim fldname_key As String = headervalue(startrow - 1, keycol)

    '                '---------------
    '                'データ部処理
    '                '---------------
    '                For cntii = 1 To rowcnt

    '                    '中断処理
    '                    Application.DoEvents()
    '                    If CancelFlg Then
    '                        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
    '                        Return rtn
    '                    End If

    '                    '行取得
    '                    Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

    '                    'キー値取得
    '                    Dim fldvalue_key As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))

    '                    'ログ出力用
    '                    Dim str_logkey As String = fldname_key & " = " & fldvalue_key

    '                    '移行値取得
    '                    For cntjj = 1 To columncnt

    '                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    '                        Dim fldvalue As String = ""
    '                        If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
    '                            fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
    '                        End If

    '                        Select Case fldname
    '                            Case "契約分類No"
    '                                .Vari_Ky_ruino = fldvalue.Trim
    '                            Case "契約分類名"
    '                                .Vari_Ky_ruiname = fldvalue.Trim
    '                            Case "備考"
    '                                .Vari_Ky_ruibiko = fldvalue.Trim
    '                            Case "更新・解約通知期間(ヶ月)"
    '                                .Vari_Ky_ruitutimm = fldvalue.Trim
    '                            Case "契約分類色"
    '                                .Vari_Ky_ruicolor = chgcolor.Chg_ColorValue(IIf(fldvalue.Trim = "", 0, fldvalue.Trim))
    '                            Case "定期借地借家権契約扱い有無"
    '                                .Vari_Teisyaku_flg = fldvalue.Trim
    '                        End Select

    '                    Next

    '                    '固定値
    '                    .Vari_Useflg = 1
    '                    .Vari_History = DefHistory
    '                    '
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
    '                            tmp_cvcnt = tmp_cvcnt + 1
    '                        End If

    '                    End If

    '                    '-------------------
    '                    'ログ出力
    '                    '-------------------

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

    '                    '-------------------------------
    '                    'プログレスバー更新/進捗率表示
    '                    '-------------------------------
    '                    '件数取得
    '                    Dim pgbcnt As Integer = 0
    '                    If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
    '                        pgbcnt = cntii
    '                    ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
    '                        Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
    '                    ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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
    '            midrowcnt = rowcnt

    '            '移行件数を取得
    '            cvrowcnt = tmp_cvcnt

    '            '調整件数を取得
    '            conditioncnt = tmp_condcnt

    '            'Excelファイル終了設定
    '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    '            '返却
    '            Return rtn

    '        End Function

    '        Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

    '        End Function

    '        Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

    '        End Sub

    '        Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

    '        End Function

    '        ''' <summary>
    '        ''' 既存データのキーを取得してリストへ格納
    '        ''' </summary>
    '        ''' <param name="sqlcnnv10"></param>
    '        ''' <param name="tblname"></param>
    '        ''' <param name="list_existdata"></param>
    '        ''' <remarks></remarks>
    '        Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

    '            Dim tmp_sql As String = " SELECT ky_ruino FROM " & tblname
    '            Dim flg As Boolean = True

    '            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

    '        End Sub

    '    End Class

    'End Class
    '20160525 紐付項目の挿入処理の追加 -del end

#End Region

#Region "保険種類マスタ"

    Public Class M_hoken_rui_Repository

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
                '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用 
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値

                Dim model_cvitem As New Njc.Model.M_hoken_rui_Model             '移行値格納用モデル初期化
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
                Dim tblname As String = "m_hoken_rui"
                Dim fldnamegrp As String = "hoken_ruino,hoken_ruiname,hoken_ruikana,biko_kihon,history," & _
                                           "useflg"

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
                                Case "保険種類No"
                                    .Vari_Hoken_ruino = fldvalue.Trim
                                Case "保険種類名"
                                    .Vari_Hoken_ruiname = fldvalue.Trim
                                Case "保険種類カナ"
                                    .Vari_Hoken_ruikana = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko_kihon = fldvalue.Trim
                            End Select

                        Next

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

            '    Dim model_cvitem As New Njc.Model.M_hoken_rui_Model             '移行値格納用モデル初期化
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
            '    Dim tblname As String = "m_hoken_rui"
            '    Dim fldnamegrp As String = "hoken_ruino,hoken_ruiname,hoken_ruikana,biko_kihon,history," & _
            '                               "useflg"

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
            '                    Case "保険種類No"
            '                        .Vari_Hoken_ruino = fldvalue.Trim
            '                    Case "保険種類名"
            '                        .Vari_Hoken_ruiname = fldvalue.Trim
            '                    Case "保険種類カナ"
            '                        .Vari_Hoken_ruikana = fldvalue.Trim
            '                    Case "備考"
            '                        .Vari_Biko_kihon = fldvalue.Trim
            '                End Select

            '            Next

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

                Dim tmp_sql As String = " SELECT hoken_ruino FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "入金区分マスタ"

    '20160525 紐付項目の挿入処理の追加 -del sta
    'Public Class M_nkbn_Repository
    '    '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    '    'Public Class SubConv
    '    '    Implements IConv

    '    '    ''' <summary>
    '    '    ''' 【変数→V10DB】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <param name="model_nkbn"></param>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Cnv_Db(ByVal sqlcnnv10 As SqlConnection, ByVal model_cvitem As Object, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Cnv_Db

    '    '        Dim reptbl As String                                '置換TBL名
    '    '        Dim repfld As String                                '置換項目名
    '    '        Dim repprm As String                                '置換パラメータ
    '    '        Dim rowcnt As Integer                               '書込行数
    '    '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    '    '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    '    '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    '    '        '書込行数取得
    '    '        rowcnt = model_cvitem.ItemCnt

    '    '        'プログレスバー初期化
    '    '        Dim pgbtotalcnt As Integer = rowcnt
    '    '        Call obj_pgb.pgbInitPart(rowcnt)

    '    '        'テーブル/フィールド名取得
    '    '        reptbl = "m_nkbn"
    '    '        repfld = "nkbn_no,nkbn_order,nkbn_name,nkbn_shortname,nkbn_zokusei," & _
    '    '                 "history"

    '    '        With model_cvitem

    '    '            For cntii = 0 To rowcnt - 1

    '    '                '中断処理
    '    '                Application.DoEvents()
    '    '                If CancelFlg Then
    '    '                    rtn = False
    '    '                    Return rtn
    '    '                End If

    '    '                If .Vari_tmp_skipflg(cntii) = False Then

    '    '                    'ハッシュテーブル作成
    '    '                    Dim hash_cvitem As New Hashtable
    '    '                    'hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(repfld, model_cvitem, cntii)

    '    '                    '書込処理
    '    '                    'Call CVDBInsert.CVitem_Insert(sqlcnnv10, reptbl, repfld, hash_cvitem)

    '    '                End If

    '    '                'プログレスバー更新/進捗率表示
    '    '                Dim pgbcnt As Integer = cntii + 1
    '    '                Call obj_pgb.pgbsettingPart(pgbcnt)
    '    '                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)

    '    '            Next

    '    '        End With

    '    '        '返却
    '    '        Return rtn

    '    '    End Function

    '    '    ''' <summary>
    '    '    ''' 【移行先DB初期化】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <remarks></remarks>
    '    '    Public Sub Initialize_Table(sqlcnnv10 As SqlConnection) Implements IConv.Initialize_Table

    '    '        '初期化対象外

    '    '    End Sub

    '    '    ''' <summary>
    '    '    ''' 【仮テーブル→変数】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <param name="syorikomok"></param>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Set_Vari2(ByVal sheetname As String, ByRef model_cvitem As Object, ByVal sqlcnnv10 As SqlConnection) As Boolean Implements IConv.Set_Vari2

    '    '        Dim reccnt As Integer                               '抽出レコード件数
    '    '        Dim fldname As String                               '該当項目名
    '    '        Dim fldvalue As String                              '該当値
    '    '        Dim readtbl As New DataTable                        'テーブル格納用
    '    '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    '    '        Dim pgbcnt As Integer = 1                           'プログレスバー更新用カウンター
    '    '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    '    '        Dim logvalue As String                              'ログ出力用
    '    '        Dim tblname As String = "tmp_nkbn_mst"              '対象テーブル名
    '    '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    '    '        '移行元情報取得
    '    '        reccnt = DBExec.Exec_DataTable(Me.Get_UseQry(), sqlcnnv10, readtbl, rtn)

    '    '        '移行元情報取得時にエラーが生じた場合は処理を抜ける
    '    '        If Not rtn Then
    '    '            Return rtn
    '    '        End If

    '    '        '初期化
    '    '        model_cvitem = New Njc.Model.M_nkbn_Model(reccnt - 1)
    '    '        Dim tmp_duplicatechk As New Njc.Model.CommonModel(reccnt - 1)

    '    '        'プログレスバー初期化
    '    '        Dim pgbtotalcnt As Integer = (readtbl.Rows.Count) * (readtbl.Columns.Count)
    '    '        Call obj_pgb.pgbInitPart(pgbtotalcnt)

    '    '        '==========================
    '    '        'データ取得　→　変数格納
    '    '        '==========================
    '    '        With model_cvitem

    '    '            For cntii As Integer = 0 To readtbl.Rows.Count - 1

    '    '                'DoEvents
    '    '                Application.DoEvents()
    '    '                If CancelFlg Then
    '    '                    rtn = False
    '    '                    Return rtn
    '    '                End If

    '    '                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

    '    '                    '項目名取得
    '    '                    fldname = readtbl.Columns(cntjj).ColumnName

    '    '                    '登録値取得
    '    '                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

    '    '                    'ログ用
    '    '                    logvalue = "賃貸革命10入金区分№ = " & Typ.ToStr(readtbl.Rows(cntii).Item(2)).Trim & _
    '    '                               " " & _
    '    '                               "移行元入金区分名称 = " & Typ.ToStr(readtbl.Rows(cntii).Item(1)).Trim

    '    '                    '各項目値→変数格納
    '    '                    Select Case fldname
    '    '                        Case "賃貸革命10入金区分№"
    '    '                            .Vari_Nkbn_no(cntii) = DataChk.Chk_DataInt(fldvalue, 1, 20)
    '    '                            .Vari_Nkbn_order(cntii) = DataChk.Chk_DataInt(fldvalue, 1, 20)
    '    '                        Case "移行元入金区分名称"
    '    '                            .Vari_Nkbn_name(cntii) = DataChk.Chk_DataString(fldvalue, 50)
    '    '                        Case "賃貸革命10入金区分属性№"
    '    '                            .Vari_Nkbn_zokusei(cntii) = fldvalue
    '    '                    End Select

    '    '                    '固定値設定
    '    '                    .Vari_Nkbn_shortname(cntii) = ""
    '    '                    .Vari_History(cntii) = DefHistory
    '    '                    .Vari_tmp_skipflg(cntii) = 0

    '    '                    'プログレスバー更新/進捗率表示
    '    '                    Call obj_pgb.pgbsettingPart(pgbcnt)
    '    '                    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
    '    '                    pgbcnt = pgbcnt + 1

    '    '                Next

    '    '            Next

    '    '        End With

    '    '        '返却
    '    '        Return rtn

    '    '    End Function

    '    '    ''' <summary>
    '    '    ''' 【抽出クエリ】2015.07.17 sol レビュー後修正_レビュー№312
    '    '    ''' </summary>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Get_UseQry(ByRef sortkeycnt As Integer) As String Implements IConv.Get_UseQry

    '    '        Dim tmp_sql As String = ""

    '    '        tmp_sql = tmp_sql & " SELECT "
    '    '        tmp_sql = tmp_sql & " 	* "
    '    '        tmp_sql = tmp_sql & " FROM tmp_nkbn_mst "
    '    '        tmp_sql = tmp_sql & " WHERE [賃貸革命10入金区分№] NOT IN (SELECT nkbn_no FROM m_nkbn) "
    '    '        tmp_sql = tmp_sql & " AND   [賃貸革命10入金区分№] <> 0 "
    '    '        tmp_sql = tmp_sql & " ORDER BY [移行元入金区分№] "

    '    '        Return tmp_sql

    '    '    End Function

    '    '    ''' <summary>
    '    '    ''' 【フィールド名と移行値を紐付けるハッシュテーブル作成】
    '    '    ''' </summary>
    '    '    ''' <param name="model_cvitem"></param>
    '    '    ''' <param name="cntii"></param>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Get_Hash_CVItem(ByVal model_cvitem As Object, ByVal cntii As Integer) As Hashtable Implements IConv.Get_Hash_CVItem

    '    '        Dim rtn_hash As New Hashtable

    '    '        With model_cvitem

    '    '            rtn_hash.Add("nkbn_no", .Vari_Nkbn_no(cntii))
    '    '            rtn_hash.Add("nkbn_order", .Vari_Nkbn_order(cntii))
    '    '            rtn_hash.Add("nkbn_name", .Vari_Nkbn_name(cntii))
    '    '            rtn_hash.Add("nkbn_shortname", .Vari_Nkbn_shortname(cntii))
    '    '            rtn_hash.Add("nkbn_zokusei", .Vari_Nkbn_zokusei(cntii))
    '    '            rtn_hash.Add("history", .Vari_History(cntii))

    '    '        End With

    '    '        Return rtn_hash

    '    '    End Function

    '    '    Public Function Set_Vari1(sqlcnnv7 As SqlConnection, ByRef model_cvitem As Object) As Boolean Implements IConv.Set_Vari1

    '    '    End Function

    '    '    Public Function Set_Vari21(sheetname As String, ByRef model_cvitem As Object) As Boolean Implements IConv.Set_Vari2

    '    '    End Function

    '    '    Public Function Write_IntermediateFile(sheetname As String, model_cvitem As Object) As Boolean Implements IConv.Write_IntermediateFile

    '    '    End Function

    '    '    Public Function Get_BaseKey() As Object Implements IConv.Get_BaseKey

    '    '    End Function

    '    '                Public Function Set_Vari( _
    '    'ByVal sqlcnnv10 As SqlConnection, _
    '    'ByVal filename As String, ByVal sheetname As String, _
    '    'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    '    ') As Boolean Implements IConv.Set_Vari

    '    '    End Function

    '    '                Public Function Set_Vari( _
    '    'ByVal sqlcnnv10 As SqlConnection, _
    '    'ByVal filename As String, ByVal sheetname As String, _
    '    'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    '    ') As Boolean Implements IConv.Set_Vari

    '    '    End Function

    '    '    Public Function Write_IntermediateFile(sqlcnnv7 As SqlConnection, filename As String, sheetname As String, model_cvitem As Object) As Boolean Implements IConv.Write_IntermediateFile

    '    '    End Function
    '    'End Class
    '    '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    'End Class
    '20160525 紐付項目の挿入処理の追加 -del end

#End Region

#Region "学校区マスタ"

    Public Class M_koku_add_Repository

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
                '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値
                Dim model_cvitem As New Njc.Model.M_koku_add_Model              '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                  'サブキー1列
                Dim keycol_sub2 As Integer = 4                                  'サブキー3列
                Dim keycol_sub3 As Integer = 5                                  'サブキー4列
                Dim keycol_sub4 As Integer = 7                                  'サブキー5列
                Dim keycol_sub5 As Integer = 8                                  'サブキー6列

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
                Dim tblname As String = "m_koku_add"
                Dim fldnamegrp As String = "ken_no,si_no,no,add_cyo,add_banti," & _
                                           "syogaku_name,cyugaku_name,history,add_cyome"

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
                        Dim tmp_addresscyotiiki As String = ""
                        Dim tmp_addresscyobanti As String = ""

                        ''キー値取得
                        'Dim fldvalue_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                        'Dim fldvalue_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub - 1)).Trim
                        'Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

                        ''ログ出力用
                        'Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

                        'データ取得
                        For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                            '項目名取得
                            fldname = readtbl.Columns(cntjj).ColumnName.Trim

                            '登録値取得
                            fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                            '各項目値→変数格納
                            Select Case fldname
                                Case "県No"
                                    .Vari_Ken_no = fldvalue
                                Case "市No"
                                    .Vari_Si_no = fldvalue
                                Case "学校区No"
                                    .Vari_No = fldvalue
                                Case "町"
                                    '町地域はそのまま移行する(丁番地以降が分割対象)
                                    .Vari_Add_cyo = fldvalue.Trim
                                Case "丁目"
                                    tmp_addresscyobanti = fldvalue
                                Case "番地"
                                    .Vari_Add_banti = fldvalue
                                Case "小学校名"
                                    .Vari_Syogaku_name = fldvalue
                                Case "中学校名"
                                    .Vari_Cyugaku_name = fldvalue
                            End Select

                        Next

                        '20161009 改善対応：学校区丁目移行不具合 -chg sta
                        ''20160622 住所分割処理対応 -chg sta
                        ' ''2016.04.06 学校区マスタ移行内容修正 -add sta
                        ' ''住所変換/格納
                        ''Dim tmp_address As String = tmp_addresscyotiiki & " " & tmp_addresscyobanti
                        ''Dim hash_address As New Hashtable
                        ''hash_address = AddressChange.Get_CyoBanti(tmp_address)
                        ''.Vari_Add_cyo = hash_address("mati")
                        ''.Vari_Add_cyome = hash_address("cyome")
                        ' ''2016.04.06 学校区マスタ移行内容修正 -add end
                        ''住所変換/格納
                        'Dim tmp_address As String = tmp_addresscyobanti
                        'Dim hash_address As New Hashtable
                        'hash_address = AddressChange.Get_CyoBanti(tmp_address)
                        '.Vari_Add_cyome = hash_address("cyome")
                        ''20160622 住所分割処理対応 -chg end

                        '※Get_CyoBantiメソッドでは連結された住所または「丁目」が付加されているデータでないと正しく移行できない為、強制で「丁目」を付加。
                        Dim tmp_address As String = Replace(Replace(Replace(tmp_addresscyobanti, "丁目", ""), "丁", ""), "町目", "")
                        tmp_address = IIf(IsNumeric(tmp_address), tmp_address & "丁目", tmp_addresscyobanti)
                        Dim hash_address As New Hashtable
                        hash_address = AddressChange.Get_CyoBanti(tmp_address)
                        .Vari_Add_cyome = hash_address("cyome")
                        '20161009 改善対応：学校区丁目移行不具合 -chg end

                        '固定値
                        .Vari_History = DefHistory

                        '2016.04.26 メインの方へも反映させる修正 -add sta
                        '町、丁目、学校名もキーとしているが、空でも登録可能であるため半角スペースを入れておく
                        If .Vari_Add_cyo = "" Then
                            .Vari_Add_cyo = " "
                        End If
                        If .Vari_Add_cyome = "" Then
                            .Vari_Add_cyome = " "
                        End If
                        If .Vari_Syogaku_name = "" Then
                            .Vari_Syogaku_name = " "
                        End If
                        If .Vari_Cyugaku_name = "" Then
                            .Vari_Cyugaku_name = " "
                        End If
                        '2016.04.26 メインの方へも反映させる修正 -add end

                        '2016.04.06 学校区マスタ移行内容修正 -add
                        '住所分割後の値で重複判断を行うためここでキーを取得する
                        Dim fldvalue_key As String = .Vari_Ken_no & "-" & _
                                                     .Vari_Si_no & "-" & _
                                                     .Vari_Add_cyo & "-" & _
                                                     .Vari_Add_cyome & "-" & _
                                                     .Vari_Syogaku_name & "-" & _
                                                     .Vari_Cyugaku_name

                        '2016.04.06 学校区マスタ移行内容修正 -del
                        'キー取得後にログ出力用文字列を作成する
                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & .Vari_Ken_no & "、" & _
                                                   fldname_keysub1 & " = " & .Vari_Si_no & "、" & _
                                                   fldname_keysub2 & " = " & .Vari_Add_cyo & "、" & _
                                                   fldname_keysub3 & " = " & .Vari_Add_cyome & "、" & _
                                                   fldname_keysub4 & " = " & .Vari_Syogaku_name & "、" & _
                                                   fldname_keysub5 & " = " & .Vari_Cyugaku_name

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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,ken_no) + '-' + CONVERT(varchar,si_no) + '-' + CONVERT(varchar,no) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "エリアマスタ"

    Public Class M_area_Repository

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
                '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
                Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
                Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
                Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
                Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
                Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
                Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
                Dim normalflg As Boolean                                        'INSERT正常終了フラグ
                Dim rtn As Boolean = True                                       '戻り値
                Dim model_cvitem As New Njc.Model.M_area_Model                  '移行値格納用モデル初期化
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
                Dim tblname As String = "m_area"
                Dim fldnamegrp As String = "area_no,area_sortorder,area_name,area_useflg,history"

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
                                Case "エリアNo"
                                    .Vari_Area_no = fldvalue.Trim
                                Case "エリア名"
                                    .Vari_Area_name = fldvalue.Trim
                            End Select

                        Next

                        '固定値
                        .Vari_Area_sortorder = Nothing
                        .Vari_Area_useflg = 1
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

            '    Dim model_cvitem As New Njc.Model.M_area_Model                  '移行値格納用モデル初期化
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
            '    Dim tblname As String = "m_area"
            '    Dim fldnamegrp As String = "area_no,area_sortorder,area_name,area_useflg,history"

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
            '                    Case "エリアNo"
            '                        .Vari_Area_no = fldvalue.Trim
            '                    Case "エリア名"
            '                        .Vari_Area_name = fldvalue.Trim
            '                End Select

            '            Next

            '            '固定値
            '            .Vari_Area_sortorder = Nothing
            '            .Vari_Area_useflg = 1
            '            .Vari_History = DefHistory

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

                Dim tmp_sql As String = " SELECT area_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "入金項目マスタ"

    '20160525 紐付項目の挿入処理の追加 -del sta
    'Public Class M_nkin_Repository

    '    '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    '    'Public Class SubConv
    '    '    Implements IConv

    '    '    ''' <summary>
    '    '    ''' 【仮テーブル→V10DB】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <param name="model_nkin"></param>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Cnv_Db(ByVal sqlcnnv10 As SqlConnection, ByVal model_cvitem As Object, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Cnv_Db

    '    '        Dim reptbl As String                                '置換TBL名
    '    '        Dim repfld As String                                '置換項目名
    '    '        Dim repprm As String                                '置換パラメータ
    '    '        Dim rowcnt As Integer                               '書込行数
    '    '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    '    '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    '    '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    '    '        '書込行数取得
    '    '        rowcnt = model_cvitem.ItemCnt

    '    '        'プログレスバー初期化
    '    '        Dim pgbtotalcnt As Integer = rowcnt
    '    '        Call obj_pgb.pgbInitPart(rowcnt)

    '    '        'テーブル/フィールド名取得
    '    '        reptbl = "m_nkin"
    '    '        repfld = "nkin_no,nkin_name,nkin_printname,nkin_ruino,nkin_zkseino," & _
    '    '                 "nkin_useflg,keiyaku_nkinno,keiyakuhiki_nkinno,kosin_nkinno,kaiyaku_nkinno," & _
    '    '                 "kaiyakuhiki_nkinno,ryosyu_flg,biko_nkin,history"

    '    '        With model_cvitem

    '    '            For cntii = 0 To rowcnt - 1

    '    '                '中断処理
    '    '                Application.DoEvents()
    '    '                If CancelFlg Then
    '    '                    rtn = False
    '    '                    Return rtn
    '    '                End If

    '    '                If .Vari_tmp_skipflg(cntii) = False Then

    '    '                    'ハッシュテーブル作成
    '    '                    Dim hash_cvitem As New Hashtable
    '    '                    'hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(repfld, model_cvitem, cntii)

    '    '                    '書込処理
    '    '                    'Call CVDBInsert.CVitem_Insert(sqlcnnv10, reptbl, repfld, hash_cvitem)

    '    '                End If

    '    '                'プログレスバー更新/進捗率表示
    '    '                Dim pgbcnt As Integer = cntii + 1
    '    '                Call obj_pgb.pgbsettingPart(pgbcnt)
    '    '                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)

    '    '            Next

    '    '        End With

    '    '        '返却
    '    '        Return rtn

    '    '    End Function

    '    '    ''' <summary>
    '    '    ''' 【移行先DB初期化】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <remarks></remarks>
    '    '    Public Sub Initialize_Table(sqlcnnv10 As SqlConnection) Implements IConv.Initialize_Table

    '    '        '初期化対象外

    '    '    End Sub

    '    '    ''' <summary>
    '    '    ''' 【仮テーブル→変数】
    '    '    ''' </summary>
    '    '    ''' <param name="sqlcnnv10"></param>
    '    '    ''' <param name="syorikomok"></param>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Overridable Function Set_Vari2(ByVal sheetname As String, ByRef model_cvitem As Object, ByVal sqlcnnv10 As SqlConnection) As Boolean Implements IConv.Set_Vari2

    '    '        Dim reccnt As Integer                               '抽出レコード件数
    '    '        Dim fldname As String                               '該当項目名
    '    '        Dim fldvalue As String                              '該当値
    '    '        Dim readtbl As New DataTable                        'テーブル格納用
    '    '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    '    '        Dim pgbcnt As Integer = 1                           'プログレスバー更新用カウンター
    '    '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    '    '        Dim logvalue As String                              'ログ出力用
    '    '        Dim tblname As String = "tmp_nkin_mst"              '対象テーブル名
    '    '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    '    '        '移行元情報取得
    '    '        reccnt = DBExec.Exec_DataTable(Me.Get_UseQry(), sqlcnnv10, readtbl, rtn)

    '    '        '移行元情報取得時にエラーが生じた場合は処理を抜ける
    '    '        If Not rtn Then
    '    '            Return rtn
    '    '        End If

    '    '        '初期化
    '    '        model_cvitem = New Njc.Model.M_nkin_Model(reccnt - 1)
    '    '        Dim tmp_duplicatechk As New Njc.Model.CommonModel(reccnt - 1)

    '    '        'プログレスバー初期化
    '    '        Dim pgbtotalcnt As Integer = (readtbl.Rows.Count) * (readtbl.Columns.Count)
    '    '        Call obj_pgb.pgbInitPart(pgbtotalcnt)

    '    '        '==========================
    '    '        'データ取得　→　変数格納
    '    '        '==========================
    '    '        With model_cvitem

    '    '            For cntii As Integer = 0 To readtbl.Rows.Count - 1

    '    '                'DoEvents
    '    '                Application.DoEvents()
    '    '                If CancelFlg Then
    '    '                    rtn = False
    '    '                    Return rtn
    '    '                End If

    '    '                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

    '    '                    '項目名取得
    '    '                    fldname = readtbl.Columns(cntjj).ColumnName

    '    '                    '登録値取得
    '    '                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

    '    '                    'ログ用
    '    '                    logvalue = "賃貸革命10入金項目№ = " & Typ.ToStr(readtbl.Rows(cntii).Item(3)).Trim & _
    '    '                               " " & _
    '    '                               "移行元入金項目名称 = " & Typ.ToStr(readtbl.Rows(cntii).Item(1)).Trim

    '    '                    '各項目値→変数格納
    '    '                    Select Case fldname
    '    '                        Case "賃貸革命10入金項目№"
    '    '                            .Vari_Nkin_no(cntii) = DataChk.Chk_DataInt(fldvalue, 1000, 9999)
    '    '                        Case "移行元入金項目名称"
    '    '                            .Vari_Nkin_name(cntii) = DataChk.Chk_DataString(fldvalue, 100)
    '    '                        Case "入金項目区分"
    '    '                            .Vari_Nkin_ruino(cntii) = fldvalue     '内部№
    '    '                        Case "賃貸革命10入金項目属性№"
    '    '                            .Vari_Nkin_zkseino(cntii) = fldvalue   '内部№
    '    '                    End Select

    '    '                    '固定値設定
    '    '                    .Vari_Nkin_printname(cntii) = ""
    '    '                    .Vari_Nkin_useflg(cntii) = 1
    '    '                    .Vari_Keiyaku_nkinno(cntii) = 0
    '    '                    .Vari_Keiyakuhiki_nkinno(cntii) = 0
    '    '                    .Vari_Kosin_nkinno(cntii) = 0
    '    '                    .Vari_Kaiyaku_nkinno(cntii) = 0
    '    '                    .Vari_Kaiyakuhiki_nkinno(cntii) = 0
    '    '                    .Vari_Ryosyu_flg(cntii) = 2
    '    '                    .Vari_Biko_nkin(cntii) = ""
    '    '                    .Vari_History(cntii) = DefHistory
    '    '                    .Vari_tmp_skipflg(cntii) = 0

    '    '                    'プログレスバー更新/進捗率表示
    '    '                    Call obj_pgb.pgbsettingPart(pgbcnt)
    '    '                    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
    '    '                    pgbcnt = pgbcnt + 1

    '    '                Next

    '    '            Next

    '    '        End With

    '    '        '返却
    '    '        Return rtn

    '    '    End Function

    '    '    ''' <summary>
    '    '    ''' 【抽出クエリ】2015.07.17 sol レビュー後修正_レビュー№312
    '    '    ''' </summary>
    '    '    ''' <returns></returns>
    '    '    ''' <remarks></remarks>
    '    '    Public Function Get_UseQry(ByRef sortkeycnt As Integer) As String Implements IConv.Get_UseQry

    '    '        Dim tmp_sql As String = ""

    '    '        tmp_sql = tmp_sql & " SELECT "
    '    '        tmp_sql = tmp_sql & " 	* "
    '    '        tmp_sql = tmp_sql & " FROM tmp_nkin_mst "
    '    '        tmp_sql = tmp_sql & " WHERE [賃貸革命10入金項目№] NOT IN (SELECT nkin_no FROM m_nkin) "
    '    '        tmp_sql = tmp_sql & " AND   [賃貸革命10入金項目№] <> 0 "
    '    '        tmp_sql = tmp_sql & " ORDER BY [移行元入金項目№] "

    '    '        Return tmp_sql

    '    '    End Function

    '    '    Public Function Get_BaseKey() As Object Implements IConv.Get_BaseKey

    '    '    End Function

    '    '                Public Function Set_Vari( _
    '    'ByVal sqlcnnv10 As SqlConnection, _
    '    'ByVal filename As String, ByVal sheetname As String, _
    '    'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    '    ') As Boolean Implements IConv.Set_Vari

    '    '    End Function

    '    '                Public Function Set_Vari( _
    '    'ByVal sqlcnnv10 As SqlConnection, _
    '    'ByVal filename As String, ByVal sheetname As String, _
    '    'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    '    ') As Boolean Implements IConv.Set_Vari

    '    '    End Function

    '    '    Public Function Write_IntermediateFile(sqlcnnv7 As SqlConnection, filename As String, sheetname As String, model_cvitem As Object) As Boolean Implements IConv.Write_IntermediateFile

    '    '    End Function

    '    'End Class
    '    '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -end

    'End Class
    '20160525 紐付項目の挿入処理の追加 -del end

#End Region

#Region "変動費マスタ"

    Public Class M_hendorule_Repository

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

                Dim model_cvitem As New Njc.Model.M_hendorule_Model             '移行値格納用モデル初期化
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
                Dim tblname As String = "m_hendorule"
                Dim fldnamegrp As String = "rule_no,rule_name,tani,tanisjis,ryokin_cnt," & _
                                           "hendo_rui,hasuusyori_sel,hasuusyosu_ptn,koukei,biko_kihon," & _
                                           "keijo_rui,zei_kbnhayami,zei_kbntanka,history,useflg"

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
                                Case "ルールNo"
                                    .Vari_Rule_no = fldvalue.Trim
                                Case "ルール名"
                                    .Vari_Rule_name = fldvalue.Trim
                                Case "その他単位"
                                    .Vari_Tani = fldvalue.Trim
                                Case "その他単位SJIS"
                                    .Vari_Tanisjis = fldvalue.Trim
                                Case "料金数"
                                    .Vari_Ryokin_cnt = fldvalue.Trim
                                Case "分類"
                                    .Vari_Hendo_rui = fldvalue.Trim
                                Case "端数処理"
                                    .Vari_Hasuusyori_sel = fldvalue.Trim
                                Case "端数処理パターン"
                                    .Vari_Hasuusyosu_ptn = fldvalue.Trim
                                Case "口径"
                                    .Vari_Koukei = fldvalue.Trim
                                Case "備考"
                                    .Vari_Biko_kihon = fldvalue.Trim
                                Case "計上分類"
                                    .Vari_Keijo_rui = fldvalue.Trim
                                Case "税区分早見"
                                    .Vari_Zei_kbnhayami = fldvalue.Trim
                                Case "税区分単価"
                                    .Vari_Zei_kbntanka = fldvalue.Trim
                            End Select

                        Next

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

                Dim tmp_sql As String = " SELECT rule_no FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "変動費一覧マスタ"

    Public Class M_hendorule_itiran_Repository

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

                Dim model_cvitem As New Njc.Model.M_hendorule_itiran_Model      '移行値格納用モデル初期化
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
                Dim tblname_base As String = "m_hendorule"
                Dim tblname As String = "m_hendorule_itiran"
                Dim fldnamegrp As String = "rule_no,item_no,siyoryo,ryokin1,ryokin2," & _
                                           "ryokin3,ryokin4,ryokin5"

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

                    '---------------
                    'ヘッダー処理
                    '---------------
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
                        Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim tmp_keysub As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub))
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub & " = " & tmp_keysub

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "変動費ルールNo"
                                    .Vari_Rule_no = fldvalue.Trim
                                Case "行No"
                                    .Vari_Item_no = fldvalue.Trim
                                Case "使用量"
                                    .Vari_Siyoryo = fldvalue.Trim
                                Case "料金１"
                                    .Vari_Ryokin1 = fldvalue.Trim
                                Case "料金２"
                                    .Vari_Ryokin2 = fldvalue.Trim
                                Case "料金３"
                                    .Vari_Ryokin3 = fldvalue.Trim
                                Case "料金４"
                                    .Vari_Ryokin4 = fldvalue.Trim
                                Case "料金５"
                                    .Vari_Ryokin5 = fldvalue.Trim
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

                Dim tmp_sql As String = " SELECT rule_no FROM " & tblname
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,buskotu_no) + '-' + CONVERT(varchar,sortorder) FROM  " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "備考タイトルマスタ" '20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -add

    Public Class M_memo_Repository

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

                Dim model_cvitem As New Njc.Model.M_memo_Model                  '移行値格納用モデル初期化
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
                Dim tblname As String = "m_memo"
                Dim fldnamegrp As String = "memo_kbn,memo_no,memo_sortorder,memo_name,memo_useflg," & _
                                           "memo_usehojoitems,history"

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
                                Case "備考区分"
                                    .Vari_Memo_kbn = fldvalue
                                Case "備考No"
                                    .Vari_Memo_no = fldvalue
                                Case "備考タイトル"
                                    .Vari_Memo_name = fldvalue
                                Case "使用有無"
                                    .Vari_Memo_useflg = fldvalue
                                Case "補助アイテム使用有無"
                                    .Vari_Memo_usehojoitems = fldvalue
                            End Select

                        Next

                        '固定値
                        .Vari_Memo_sortorder = 0
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,memo_kbn) + '-' + CONVERT(varchar,memo_no) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "備考入力補助リストマスタ"  '20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -add

    Public Class M_memo_lst_Repository

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

                Dim model_cvitem As New Njc.Model.M_memo_lst_Model              '移行値格納用モデル初期化
                Dim keycol_main As Integer = 1                                  'メインキー列
                Dim keycol_sub1 As Integer = 2                                   'サブ1キー列
                Dim keycol_sub2 As Integer = 3                                   'サブ2キー列

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
                Dim tblname_base As String = "m_memo"
                Dim tblname As String = "m_memo_lst"
                Dim fldnamegrp As String = "memo_kbn,memo_no,memo_lstno,memo_lstname,history"

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
                        Dim fldvalue_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
                        Dim fldvalue_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
                        Dim fldvalue_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
                        Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub1 & "-" & fldvalue_keysub2

                        'ログ出力用
                        Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub1 & " = " & fldvalue_keysub1 & "、" & fldname_keysub2 & " = " & fldvalue_keysub2

                        '移行値取得
                        For cntjj = 1 To columncnt

                            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                            Dim fldvalue As String = ""
                            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                            End If

                            Select Case fldname
                                Case "備考区分"
                                    .Vari_Memo_kbn = fldvalue
                                Case "備考No"
                                    .Vari_Memo_no = fldvalue
                                Case "備考入力補助リストNo"
                                    .Vari_Memo_lstno = fldvalue
                                Case "備考入力補助リスト内容"
                                    .Vari_Memo_lstname = fldvalue
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,memo_kbn) + '-' + CONVERT(varchar,memo_no) FROM " & tblname
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,memo_kbn) + '-' + CONVERT(varchar,memo_no) + '-' + CONVERT(varchar,memo_lstno) FROM  " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

#Region "画像タイトルマスタ" '20160720 連動情報構築

    Public Class M_gazo_title_Repository

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

                Dim model_cvitem As New Njc.Model.M_gazo_title_Model            '移行値格納用モデル初期化
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
                Dim tblname As String = "m_gazo_title"
                Dim fldnamegrp As String = "gazo_kbn,gazo_no,gazo_name,history"

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
                                Case "画像区分"
                                    .Vari_Gazo_kbn = fldvalue
                                Case "画像No"
                                    .Vari_Gazo_no = fldvalue
                                Case "画像名"
                                    .Vari_Gazo_name = fldvalue
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

                Dim tmp_sql As String = " SELECT CONVERT(varchar,memo_kbn) + '-' + CONVERT(varchar,memo_no) FROM " & tblname
                Dim flg As Boolean = True

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

            End Sub

        End Class

    End Class

#End Region

End Namespace


