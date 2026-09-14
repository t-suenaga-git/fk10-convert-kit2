Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "既存用中間ファイルのヘッダーを革命10のヘッダーに変換して値を紐付ける"

    Public Class Chk_middata_Repository

        ''' <summary>
        ''' 家主イベント情報の変換とチェック処理
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="filename"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_owdata_event_mid(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim tmp_condcnt As Integer                                      '調整件数格納

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
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

                        'ログ出力メッセージ整形
                        Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                        'ログ出力
                        If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                            Dim tmp_cnt As Integer = 0
                            '挿入
                            For Each logvalue In sortlist_log
                                Dim tmp_sql_insert As String = logvalue.Value
                                tmp_sql_insert = tmp_sql_insert.Replace(LOG_TMP_TABLENAME, LOG_TMP_MIDCHKTABLENAME)
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                            Next
                            '初期化
                            tmp_logcnt = 0
                            sortlist_log.Clear()
                        End If

                    Next

                Next

            End With

            '************************
            '終了処理
            '************************

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 家主メモ情報の変換とチェック処理
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="filename"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Chk_owdata_memo_mid(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim tmp_condcnt As Integer                                      '調整件数格納

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
            Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
            Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
            Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

            Dim tmp_hash As New Hashtable                                   '作業用ハッシュテーブル
            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            Dim rtn As Boolean = True                                       '戻り値

            Dim model_cvitem As New Njc.Model.Owdata_memo_Model            '移行値格納用モデル初期化
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
            Dim tblname As String = "owdata_memo"
            Dim fldnamegrp As String = "ow_no,memo_no,memo,history"

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

                    '備考カウント初期化
                    Dim cvitemcnt As Integer = 0

                    '移行値取得
                    For cntjj = 2 To columncnt

                        'サブキー取得
                        Dim tmp_keysub As String = (cntjj - 1).ToString

                        '全キー取得
                        Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

                        'ログ出力用データ格納(サブキーフィールド)
                        Dim fldname_keysub As String = headervalue(startrow - 1, cntjj)
                        Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub

                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                        Dim fldvalue As String = ""
                        If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                            fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                        End If
                        .Vari_Memo = fldvalue

                        '固定値
                        .Vari_Memo_no = (cntjj - 1).ToString
                        .Vari_History = DefHistory

                        'メモにデータが存在する場合に書込処理を行う
                        If .Vari_Memo <> "" Then

                            'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                            tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                            'データチェック
                            Dim skipflg As Boolean = False
                            Dim hash_cvitem As New Hashtable
                            Dim hash_log As New Hashtable
                            skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

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

                Next

            End With

            '************************
            '終了処理
            '************************

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            '返却
            Return rtn

        End Function

    End Class

#End Region

End Namespace
