using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Converter10.Njc.N3Lib.Utys;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Repository
{

    #region 自社口座情報(汎用用)

    public class Jisyadata_koza_BaseMid_Repository
    {

        public class SubConv : IConv
        {
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -add sta
            /// <summary>
            /// 【中間ファイル→変数】 
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="filename"></param>
            /// <param name="sheetname"></param>
            /// <param name="midrowcnt"></param>
            /// <param name="cvrowcnt"></param>
            /// <param name="conditioncnt"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納
                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Jisyadata_koza_BaseMid_Model();  // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                Get_RelData_KozaSyubetu();

                // データ取得
                string tmptblname = CommonModule.PRE_TBL_NAME + sheetname;
                string tmp_sql = " SELECT * FROM " + tmptblname;
                var readtbl = new DataTable();
                int rowcnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, ref rtn);

                // オープン処理失敗時は処理を抜ける
                if (rtn == false)
                {
                    return rtn;
                }

                // テーブル名/フィールド名セット
                string tblname_base = "jisyadata";
                string tblname = "jisyadata_koza";
                string fldnamegrp = "jisya_no,jisya_kozano,kinyu_no,kinyu_tenno,koza_syubetu," + "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," + "yucyokoza_bango,biko_koza,history,useflg";


                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, tblname_base, ref list_basekeydata);

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // プログレスバー初期化
                int pgbtotalcnt = 0;          // プログレスバー総件数初期化
                int pgbbasecnt = 100;         // 実件数で表示するか否かの基準値

                if (rowcnt <= pgbbasecnt)
                {
                    pgbtotalcnt = rowcnt;
                }
                else
                {
                    pgbtotalcnt = (int)Math.Round(Math.Ceiling(rowcnt / (double)pgbbasecnt));
                }
                obj_pgb.pgbInitPart(pgbtotalcnt);

                // 'ヘッダー格納用
                // Dim fldname_keymain As String = ""
                // Dim fldname_keysub As String = ""

                // ************************
                // 処理開始
                // ************************

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol_main - 1].ColumnName.Trim();
                string fldname_keysub = readtbl.Columns[keycol_sub - 1].ColumnName.Trim();

                for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
                {

                    // 中断処理
                    Application.DoEvents();
                    if (CommonModule.CancelFlg)
                    {
                        return rtn;
                    }

                    // 作業用変数
                    string fldname = "";
                    string fldvalue = "";

                    // キー値取得
                    string fldvalue_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string fldvalue_keysub = Typ.ToStr(readtbl.Rows[cntii][keycol_sub - 1]).Trim();
                    string fldvalue_key = fldvalue_keymain + "-" + fldvalue_keysub;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + fldvalue_keymain + "、" + fldname_keysub + " = " + fldvalue_keysub;

                    // データ取得
                    for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                    {

                        // 項目名取得
                        fldname = readtbl.Columns[cntjj].ColumnName.Trim();

                        // 登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();

                        switch (fldname ?? "")
                        {
                            case "自社・支店No":
                                {
                                    model_cvitem.Vari_Jisya_no = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店口座No":
                                {
                                    model_cvitem.Vari_Jisya_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関No":
                                {
                                    model_cvitem.Vari_Kinyu_no = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関店No":
                                {
                                    model_cvitem.Vari_Kinyu_tenno = fldvalue.Trim();
                                    break;
                                }
                            case "口座種別":
                                {
                                    model_cvitem.Vari_Koza_syubetu = fldvalue.Trim();
                                    break;
                                }
                            case "口座番号":
                                {
                                    model_cvitem.Vari_Koza_bango = fldvalue.Trim();
                                    break;
                                }
                            case "口座名義":
                                {
                                    model_cvitem.Vari_Koza_meigi = fldvalue.Trim();
                                    break;
                                }
                            case "口座名義カナ":
                                {
                                    model_cvitem.Vari_Koza_meigikana = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ口座記号１":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo1 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ口座記号２":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo2 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ口座番号":
                                {
                                    model_cvitem.Vari_Yucyokoza_bango = fldvalue.Trim();
                                    break;
                                }
                            case "備考(口座情報)":
                                {
                                    model_cvitem.Vari_Biko_koza = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Useflg = 1.ToString();

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 20161125 自社口座情報で仮口座作成時にゆうちょが存在する場合は設定しないようにする処理を追加 -add sta
                    // ゆうちょ関連情報にデータが存在する場合は設定しておいた仮の金融機関とログを削除する
                    string tmp_yucyoinfo = model_cvitem.Vari_Yucyokoza_kigo1 + model_cvitem.Vari_Yucyokoza_kigo2 + model_cvitem.Vari_Yucyokoza_bango;
                    if (!string.IsNullOrEmpty(tmp_yucyoinfo))
                    {
                        // ゆうちょ銀行にデータがある場合はそちらを優先する
                        // 金融機関、支店Noを削除
                        hash_cvitem["kinyu_no"] = "";
                        hash_cvitem["kinyu_tenno"] = "";

                        // ログを削除
                        hash_log.Remove("jisyadata_koza-kinyu_no");
                        hash_log.Remove("jisyadata_koza-kinyu_tenno");
                    }
                    // 20161125 自社口座情報で仮口座作成時にゆうちょが存在する場合は設定しないようにする処理を追加 -add end

                    // 書込処理
                    if (!skipflg)
                    {

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                        }

                    }

                    // -------------------
                    // ログ出力
                    // -------------------

                    // ログ出力メッセージ整形
                    LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, ref tmp_logcnt, ref sortlist_log);

                    // ログ出力
                    if (sortlist_log.Count >= CommonModule.Log_OutputCnt | sortlist_log.Count != 0 & cntii == rowcnt - 1)
                    {
                        int tmp_cnt = 0;
                        // 挿入
                        foreach (var logvalue in sortlist_log)
                        {
                            string tmp_sql_insert = logvalue.Value;
                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_cnt);
                        }
                        // 初期化
                        tmp_logcnt = 0;
                        sortlist_log.Clear();
                    }

                    // -------------------------------
                    // プログレスバー更新/進捗率表示
                    // -------------------------------
                    // 件数取得
                    int pgbcnt = 0;
                    if (rowcnt <= pgbbasecnt)        // 基準件数(100件)以下の場合は実件数を取得
                    {
                        pgbcnt = cntii + 1;
                    }
                    else if (rowcnt > cntii + 2)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 2)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                    {
                        pgbcnt = pgbtotalcnt;
                    }

                    // 表示
                    if (pgbcnt != 0)
                    {
                        obj_pgb.pgbsettingPart(pgbcnt);
                        obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt);
                    }


                }

                // ************************
                // 終了処理
                // ************************

                // 中間ファイル件数を取得
                midrowcnt = rowcnt;

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 返却
                return rtn;

            }
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -add end

            // 20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -del sta
            // ''' <summary>
            // ''' 【中間ファイル→変数】
            // ''' </summary>
            // ''' <param name="sqlcnnv10"></param>
            // ''' <param name="filename"></param>
            // ''' <param name="sheetname"></param>
            // ''' <param name="midrowcnt"></param>
            // ''' <param name="cvrowcnt"></param>
            // ''' <param name="conditioncnt"></param>
            // ''' <returns></returns>
            // ''' <remarks></remarks>
            // Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim tmp_cvcnt As Integer                                        '移行件数格納
            // Dim tmp_condcnt As Integer                                      '調整件数格納

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            // Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            // Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
            // Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
            // Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
            // Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
            // Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用

            // Dim tmp_hash As New SafeDictionary<string, string>                                   '作業用ハッシュテーブル
            // Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            // Dim rtn As Boolean = True                                       '戻り値

            // Dim model_cvitem As New Njc.Model.Jisyadata_koza_BaseMid_Model  '移行値格納用モデル初期化
            // Dim keycol_main As Integer = 1                                  'メインキー列
            // Dim keycol_sub As Integer = 2                                   'サブキー列

            // '************************
            // '作業準備
            // '************************

            // '紐付けデータ取得
            // Call Me.Get_RelData_KozaSyubetu()

            // 'Excelファイル初期設定
            // '汎用コンバートの自社口座は汎用用中間ファイルから取得する
            // Dim readfilename As String = CV_FROM_MIDDLE            '汎用用中間ファイル名
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, BaseMidDirPath, readfilename, sheetname)

            // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
            // If Not rtn Then
            // Return rtn
            // End If

            // 'テーブル名/フィールド名セット
            // Dim tblname_base As String = "jisyadata"
            // Dim tblname As String = "jisyadata_koza"
            // Dim fldnamegrp As String = "jisya_no,jisya_kozano,kinyu_no,kinyu_tenno,koza_syubetu," & _
            // "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," & _
            // "yucyokoza_bango,biko_koza,history,useflg"

            // '親マスタ取得
            // Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)

            // '追加コンバート時の重複チェック用に既存データのキーを取得
            // If Not InitDBFlg Then
            // Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
            // End If

            // 'プログレスバー初期化
            // Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
            // Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

            // If rowcnt <= pgbbasecnt Then
            // pgbtotalcnt = rowcnt
            // Else
            // pgbtotalcnt = Math.Ceiling(rowcnt / pgbbasecnt)
            // End If
            // Call obj_pgb.pgbInitPart(pgbtotalcnt)

            // '************************
            // '処理開始
            // '************************
            // With model_cvitem

            // 'ヘッダー行取得
            // Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
            // Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 2), wsheet.Cells(headerrow, columncnt)).Value

            // 'キーヘッダー名取得
            // Dim fldname_keymain As String = headervalue(1, keycol_main)
            // Dim fldname_keysub As String = headervalue(1, keycol_sub)

            // '---------------
            // 'データ部処理
            // '---------------
            // For cntii = 1 To rowcnt

            // '中断処理
            // Application.DoEvents()
            // If CancelFlg Then
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            // Return rtn
            // End If

            // '行取得
            // Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 2), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

            // 'キー値取得
            // Dim fldvalue_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(1, keycol_main))
            // Dim fldvalue_keysub As String = EtcMethod.Get_ShapeStr(datarowvalue(1, keycol_sub))
            // Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

            // 'ログ出力用
            // Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

            // '移行値取得
            // For cntjj = 1 To columncnt - 1

            // Dim fldname As String = headervalue(1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "自社・支店No"
            // .Vari_Jisya_no = fldvalue.Trim
            // Case "自社・支店口座No"
            // .Vari_Jisya_kozano = fldvalue.Trim
            // Case "金融機関No"
            // .Vari_Kinyu_no = fldvalue.Trim
            // Case "金融機関店No"
            // .Vari_Kinyu_tenno = fldvalue.Trim
            // Case "口座種別"
            // .Vari_Koza_syubetu = fldvalue.Trim
            // Case "口座番号"
            // .Vari_Koza_bango = fldvalue.Trim
            // Case "口座名義"
            // .Vari_Koza_meigi = fldvalue.Trim
            // Case "口座名義カナ"
            // .Vari_Koza_meigikana = fldvalue.Trim
            // Case "ゆうちょ口座記号１"
            // .Vari_Yucyokoza_kigo1 = fldvalue.Trim
            // Case "ゆうちょ口座記号２"
            // .Vari_Yucyokoza_kigo2 = fldvalue.Trim
            // Case "ゆうちょ口座番号"
            // .Vari_Yucyokoza_bango = fldvalue.Trim
            // Case "備考(口座情報)"
            // .Vari_Biko_koza = fldvalue.Trim
            // End Select

            // Next

            // '固定値
            // .Vari_History = DefHistory
            // .Vari_Useflg = 1

            // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            // 'データチェック
            // Dim skipflg As Boolean = False
            // Dim hash_cvitem As New SafeDictionary<string, string>
            // Dim hash_log As New SafeDictionary<string, string>
            // skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

            // '書込処理
            // If Not skipflg Then

            // '挿入処理
            // Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

            // '挿入処理が正常終了したレコードのキーを重複チェック用に格納
            // If normalflg Then
            // list_chkduplicate.Add(fldvalue_key)
            // tmp_cvcnt = tmp_cvcnt + 1
            // End If

            // End If

            // '-------------------
            // 'ログ出力
            // '-------------------

            // 'ログ出力メッセージ整形
            // Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

            // 'ログ出力
            // If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
            // Dim tmp_cnt As Integer = 0
            // '挿入
            // For Each logvalue In sortlist_log
            // Dim tmp_sql_insert As String = logvalue.Value
            // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
            // Next
            // '初期化
            // tmp_logcnt = 0
            // sortlist_log.Clear()
            // End If

            // '-------------------------------
            // 'プログレスバー更新/進捗率表示
            // '-------------------------------
            // '件数取得
            // Dim pgbcnt As Integer = 0
            // If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
            // pgbcnt = cntii
            // ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
            // Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
            // ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
            // pgbcnt = pgbtotalcnt
            // End If

            // '表示
            // If pgbcnt <> 0 Then
            // Call obj_pgb.pgbsettingPart(pgbcnt)
            // Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
            // End If

            // Next

            // End With

            // '************************
            // '終了処理
            // '************************

            // '中間ファイル件数を取得
            // midrowcnt = rowcnt

            // '移行件数を取得
            // cvrowcnt = tmp_cvcnt

            // '調整件数を取得
            // conditioncnt = tmp_condcnt

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // '返却
            // Return rtn

            // End Function
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -del end

            /// <summary>
            /// 親マスタ取得→リスト格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT jisya_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_basedata);

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public string Get_UseQry(ref string sortstr)
            {
                return default;

            }

            /// <summary>
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT CONVERT(varchar,ow_no) + '-' + CONVERT(varchar,ow_kozano) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 口座種別紐付情報取得
            /// </summary>
            /// <remarks></remarks>
            public void Get_RelData_KozaSyubetu()
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                string relsheetname = "口座種別マスタ";
                bool rtn = true;

                var excelfile = new ExcelFileManager();                // Excelファイル操作用

                // 既存データ有無確認(口座種別は他の項目でも参照するため)
                if (CommonModule.Hash_Rel_Kozasyubetu.Count != 0)
                {
                    return;
                }

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

                for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
                {

                    // 作業用変数作成
                    string tmp_oldkozasyubetuname = "";
                    string tmp_newkozasyubetuno = "";

                    // 紐付設定値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        // ヘッダー格納
                        string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                        // 移行値格納
                        string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                        switch (fldname ?? "")
                        {
                            case "移行元口座種別名称":
                                {
                                    tmp_oldkozasyubetuname = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸革命10口座種別No":
                                {
                                    tmp_newkozasyubetuno = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // ハッシュテーブル格納
                    if (!string.IsNullOrEmpty(tmp_newkozasyubetuno))
                    {
                        CommonModule.Hash_Rel_Kozasyubetu.Add(tmp_oldkozasyubetuname, tmp_newkozasyubetuno);
                    }

                }

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            }

        }

    }

    #endregion

    #region 部屋設備情報(汎用用)

    public class Hydata_setubilst_BaseMid_Repository
    {

        public class SubConv : IConv
        {
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -add sta
            /// <summary>
            /// 【中間ファイル→変数】 
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納
                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var hash_hyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_setubilst_BaseMid_Model();    // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                  // サブ1キー列

                // ************************
                // 作業準備
                // ************************

                // 革命10から設備マスタを取得
                var hash_setubimst = new SafeDictionary<string, string>();
                EtcMethod.Set_SetubiMst_BaseMId(sqlcnnv10, ref hash_setubimst);

                // データ取得
                string tmptblname = CommonModule.PRE_TBL_NAME + sheetname;
                string tmp_sql = " SELECT * FROM " + tmptblname;
                var readtbl = new DataTable();
                int rowcnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, ref rtn);

                // オープン処理失敗時は処理を抜ける
                if (rtn == false)
                {
                    return rtn;
                }

                // guid取得
                Set_HyGuid(sqlcnnv10, ref hash_hyguid);

                // テーブル名/フィールド名セット
                string tblname = "hydata_setubilst";
                string fldnamegrp = "hy_guid,komok_guid,override_kbn,komok_name,disp1name," + "disp2name,disp3name,disp1iconguid,disp2iconguid,disp3iconguid," + "history";


                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_hyguid, ref list_basekeydata);

                // プログレスバー初期化
                int pgbtotalcnt = 0;          // プログレスバー総件数初期化
                int pgbbasecnt = 100;         // 実件数で表示するか否かの基準値

                if (rowcnt <= pgbbasecnt)
                {
                    pgbtotalcnt = rowcnt;
                }
                else
                {
                    pgbtotalcnt = (int)Math.Round(Math.Ceiling(rowcnt / (double)pgbbasecnt));
                }
                obj_pgb.pgbInitPart(pgbtotalcnt);

                // 移行対象件数取得用
                int cvtaisyocnt = 0;

                // ************************
                // 処理開始
                // ************************

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol_main - 1].ColumnName.Trim();
                string fldname_keysub = readtbl.Columns[keycol_sub - 1].ColumnName.Trim();

                for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
                {

                    // 中断処理
                    Application.DoEvents();
                    if (CommonModule.CancelFlg)
                    {
                        return rtn;
                    }

                    // 作業用変数
                    var hash_log = new SafeDictionary<string, string>();
                    var hash_cvitem = new SafeDictionary<string, string>();
                    int setubicvcnt = 0;

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub = Typ.ToStr(readtbl.Rows[cntii][keycol_sub - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub + " = " + tmp_keysub;

                    // 移行判別用
                    bool cvflg = true;
                    bool keychkflg = true;

                    // データ有無チェック
                    string tmp_fldvalueumuchk = "";
                    for (int cntjj = 2, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                        tmp_fldvalueumuchk = tmp_fldvalueumuchk + Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();
                    if (string.IsNullOrEmpty(tmp_fldvalueumuchk))
                    {
                        cvflg = false;
                    }

                    if (cvflg)
                    {

                        // 移行対象件数カウント
                        cvtaisyocnt = cvtaisyocnt + 1;

                        // 親データ有無チェック
                        if (list_basekeydata.Contains(fldvalue_key) == false)
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "hydata_setubilst-hy_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_log.Clear();
                            hash_log.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(fldvalue_key))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "hydata_setubilst-hy_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP;
                            hash_log.Clear();
                            hash_log.Add(log_key, log_value);
                        }

                        // データ取得、移行処理
                        if (keychkflg)
                        {

                            // 作業用変数
                            string fldname = "";
                            string fldvalue = "";

                            // 移行値取得
                            for (int cntjj = 2, loopTo2 = readtbl.Columns.Count - 1; cntjj <= loopTo2; cntjj++)
                            {

                                // 項目名取得
                                fldname = readtbl.Columns[cntjj].ColumnName.Trim();

                                // 登録値取得
                                fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();

                                string value_total = fldname.Replace("-", CommonModule.STR_SPLIT_1) + CommonModule.STR_SPLIT_1 + fldvalue;

                                // 20170116 部屋設備情報未設定項目のログ出力制御修正 -chg sta
                                // 項目が存在するかで処理を分岐するように修正する
                                // '設備マスターと一致した場合
                                // If hash_setubimst.Contains(value_total) Then

                                // '項目guidを取得
                                // .Vari_Komok_guid = hash_setubimst(value_total)

                                // '固定値
                                // .Vari_Override_kbn = 0
                                // .Vari_Komok_name = ""
                                // .Vari_Disp1name = ""
                                // .Vari_Disp2name = ""
                                // .Vari_Disp3name = ""
                                // .Vari_Disp1iconguid = "00000000-0000-0000-0000-000000000000"
                                // .Vari_Disp2iconguid = "00000000-0000-0000-0000-000000000000"
                                // .Vari_Disp3iconguid = "00000000-0000-0000-0000-000000000000"
                                // .Vari_History = DefHistory

                                // '部屋guid変換
                                // .Vari_Hy_guid = hash_hyguid(fldvalue_key)

                                // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                                // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                                // '挿入処理
                                // Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, normalflg)

                                // '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                // If normalflg Then
                                // If list_chkduplicate.Contains(fldvalue_key) = False Then
                                // list_chkduplicate.Add(fldvalue_key)
                                // End If
                                // setubicvcnt = setubicvcnt + 1
                                // End If

                                // Else

                                // 'マスタ不一致ログ出力
                                // Dim log_key As String = "hydata_setubilst-komok_guid"
                                // Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & "設備" & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                                // Dim hash_komklog As New SafeDictionary<string, string> From {{log_key, log_value}}

                                // 'ログ出力メッセージ整形
                                // tmp_hash.Clear()
                                // hash_cvitem.Clear()
                                // Call LogSetting.Set_Log_Value_KomkErr(hash_komklog, tmp_hash, hash_cvitem, str_logkey & "、" & fldname, tblname, tmp_logcnt, sortlist_log)

                                // 'ログ出力
                                // If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                                // Dim tmp_cnt As Integer = 0
                                // '挿入
                                // For Each logvalue In sortlist_log
                                // Dim tmp_sql_insert As String = logvalue.Value
                                // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                                // Next
                                // '初期化
                                // tmp_logcnt = 0
                                // sortlist_log.Clear()
                                // End If

                                // End If
                                if (!string.IsNullOrEmpty(fldvalue))
                                {

                                    // 設備マスターと一致した場合
                                    if (hash_setubimst.ContainsKey(value_total))
                                    {

                                        // 項目guidを取得
                                        model_cvitem.Vari_Komok_guid = Conversions.ToString(hash_setubimst[value_total]);

                                        // 固定値
                                        model_cvitem.Vari_Override_kbn = 0.ToString();
                                        model_cvitem.Vari_Komok_name = "";
                                        model_cvitem.Vari_Disp1name = "";
                                        model_cvitem.Vari_Disp2name = "";
                                        model_cvitem.Vari_Disp3name = "";
                                        model_cvitem.Vari_Disp1iconguid = "00000000-0000-0000-0000-000000000000";
                                        model_cvitem.Vari_Disp2iconguid = "00000000-0000-0000-0000-000000000000";
                                        model_cvitem.Vari_Disp3iconguid = "00000000-0000-0000-0000-000000000000";
                                        model_cvitem.Vari_History = CommonModule.DefHistory;

                                        // 部屋guid変換
                                        model_cvitem.Vari_Hy_guid = Conversions.ToString(hash_hyguid[fldvalue_key]);

                                        // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                                        // 挿入処理
                                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, ref normalflg);

                                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                        if (normalflg)
                                        {
                                            if (list_chkduplicate.Contains(fldvalue_key) == false)
                                            {
                                                list_chkduplicate.Add(fldvalue_key);
                                            }
                                            setubicvcnt = setubicvcnt + 1;
                                        }
                                    }

                                    else
                                    {

                                        // マスタ不一致ログ出力
                                        string log_key = "hydata_setubilst-komok_guid";
                                        string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + "設備" + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                                        var hash_komklog = new SafeDictionary<string, string>() { { log_key, log_value } };

                                        // ログ出力メッセージ整形
                                        tmp_hash.Clear();
                                        hash_cvitem.Clear();
                                        LogSetting.Set_Log_Value_KomkErr(hash_komklog, tmp_hash, hash_cvitem, str_logkey + "、" + fldname, tblname, ref tmp_logcnt, ref sortlist_log);

                                        // ログ出力
                                        if (sortlist_log.Count >= CommonModule.Log_OutputCnt | sortlist_log.Count != 0 & cntii == rowcnt - 1)
                                        {
                                            int tmp_cnt = 0;
                                            // 挿入
                                            foreach (var logvalue in sortlist_log)
                                            {
                                                string tmp_sql_insert = logvalue.Value;
                                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_cnt);
                                            }
                                            // 初期化
                                            tmp_logcnt = 0;
                                            sortlist_log.Clear();
                                        }

                                    }

                                }
                                // 20170116 部屋設備情報未設定項目のログ出力制御修正 -chg end
                            }
                        }

                        else
                        {

                            // 親データ無し、重複チェックエラーログ出力
                            // ログ出力メッセージ整形
                            tmp_hash.Clear();
                            hash_cvitem.Clear();
                            LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, ref tmp_logcnt, ref sortlist_log);

                            // ログ出力
                            if (sortlist_log.Count >= CommonModule.Log_OutputCnt | sortlist_log.Count != 0 & cntii == rowcnt - 1)
                            {
                                int tmp_cnt = 0;
                                // 挿入
                                foreach (var logvalue in sortlist_log)
                                {
                                    string tmp_sql_insert = logvalue.Value;
                                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_cnt);
                                }
                                // 初期化
                                tmp_logcnt = 0;
                                sortlist_log.Clear();
                            }

                        }

                    }

                    // -------------------------------
                    // 件数取得
                    // -------------------------------
                    if (setubicvcnt > 0)
                    {
                        tmp_cvcnt = tmp_cvcnt + 1;
                    }

                    // -------------------------------
                    // プログレスバー更新/進捗率表示
                    // -------------------------------
                    // 件数取得
                    int pgbcnt = 0;
                    if (rowcnt <= pgbbasecnt)        // 基準件数(100件)以下の場合は実件数を取得
                    {
                        pgbcnt = cntii + 1;
                    }
                    else if (rowcnt > cntii + 2)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 2)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                    {
                        pgbcnt = pgbtotalcnt;
                    }

                    // 表示
                    if (pgbcnt != 0)
                    {
                        obj_pgb.pgbsettingPart(pgbcnt);
                        obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt);
                    }


                }

                // ************************
                // 終了処理
                // ************************

                // 中間ファイル件数を取得
                midrowcnt = cvtaisyocnt;

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 返却
                return rtn;

            }
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -add end

            // 20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -del sta
            // ''' <summary>
            // ''' 【中間ファイル→変数】 
            // ''' </summary>
            // ''' <param name="sqlcnnv10"></param>
            // ''' <returns></returns>
            // ''' <remarks></remarks>
            // Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

            // Dim tmp_cvcnt As Integer                                        '移行件数格納
            // Dim tmp_condcnt As Integer                                      '調整件数格納
            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            // Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            // Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
            // Dim list_chkduplicate As New List(Of String)                    '重複チェック用リスト
            // Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
            // Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
            // Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
            // Dim tmp_hash As New SafeDictionary<string, string>                                   '作業用ハッシュテーブル
            // Dim hash_hyguid As New SafeDictionary<string, string>                            'キー/guid格納用ハッシュテーブル
            // Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            // Dim rtn As Boolean = True                                       '戻り値

            // Dim model_cvitem As New Njc.Model.Hydata_setubilst_BaseMid_Model    '移行値格納用モデル初期化
            // Dim keycol_main As Integer = 1                                  'メインキー列
            // Dim keycol_sub As Integer = 2                                  'サブ1キー列

            // '************************
            // '作業準備
            // '************************

            // '革命10から設備マスタを取得
            // Dim hash_setubimst As New SafeDictionary<string, string>
            // Call EtcMethod.Set_SetubiMst_BaseMId(sqlcnnv10, hash_setubimst)

            // 'Excelファイル初期設定                  
            // Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
            // Dim readtbl As New DataTable()
            // Dim con_read As New OleDbConnection()

            // 'オープン処理
            // Dim readfilename As String = CV_FROM_MIDDLE            '汎用用中間ファイル名
            // rtn = excelfile.ExcelFile_ReadOpen(BaseMidDirPath, readfilename, tmp_sql, readtbl, con_read)

            // 'オープン処理失敗時は処理を抜ける
            // If rtn = False Then
            // excelfile.ExcelFile_ReadClose(con_read)
            // Return rtn
            // End If

            // '行数取得
            // Dim totalrowcnt As Integer = readtbl.Rows.Count     '全件数
            // Dim rowcnt As Integer = readtbl.Rows.Count - 10     'データ部件数

            // 'guid取得
            // Call Me.Set_HyGuid(sqlcnnv10, hash_hyguid)

            // 'テーブル名/フィールド名セット
            // Dim tblname As String = "hydata_setubilst"
            // Dim fldnamegrp As String = "hy_guid,komok_guid,override_kbn,komok_name,disp1name," & _
            // "disp2name,disp3name,disp1iconguid,disp2iconguid,disp3iconguid," & _
            // "history"

            // '追加コンバート時の重複チェック用に既存データのキーを取得
            // 'If Not InitDBFlg Then
            // '    Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
            // 'End If

            // '親マスタ取得
            // Call EtcMethod.Set_HashKeyToList(hash_hyguid, list_basekeydata)

            // 'プログレスバー初期化
            // Dim pgbtotalcnt As Integer = 0          'プログレスバー総件数初期化
            // Dim pgbbasecnt As Integer = 100         '実件数で表示するか否かの基準値

            // If totalrowcnt <= pgbbasecnt Then
            // pgbtotalcnt = totalrowcnt
            // Else
            // pgbtotalcnt = Math.Ceiling(totalrowcnt / pgbbasecnt)
            // End If
            // Call obj_pgb.pgbInitPart(pgbtotalcnt)

            // 'ヘッダー格納用
            // Dim fldname_keymain As String = ""
            // Dim fldname_keysub As String = ""

            // '移行対象件数取得用
            // Dim cvtaisyocnt As Integer = 0

            // '************************
            // '処理開始
            // '************************
            // With model_cvitem

            // '-------------------------------------------------------------------
            // 'データ部処理(汎用中間ファイルからはヘッダーを取得できないので行Noから取得する)
            // '-------------------------------------------------------------------
            // For cntii = 0 To totalrowcnt - 1

            // '中断処理
            // Application.DoEvents()
            // If CancelFlg Then
            // Call excelfile.ExcelFile_ReadClose(con_read)
            // Return rtn
            // End If

            // 'キー格納用
            // Dim fldvalue_key As String = ""
            // Dim str_logkey As String = ""
            // Dim str_taihilogkey As String = ""
            // Dim str_taihilog As String = ""
            // Dim setubicvcnt As Integer = 0

            // Dim hash_log As New SafeDictionary<string, string>
            // Dim hash_cvitem As New SafeDictionary<string, string>

            // '移行判別用
            // Dim cvflg As Boolean = True
            // Dim keychkflg As Boolean = True

            // If cntii = 0 Then

            // 'ヘッダー取得
            // fldname_keymain = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main)).Trim
            // fldname_keysub = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub)).Trim

            // ElseIf cntii >= BASEMIDFILE_READWRITE_ROW - 2 Then

            // 'キー値取得
            // Dim tmp_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main)).Trim
            // Dim tmp_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub)).Trim
            // fldvalue_key = tmp_keymain & "-" & tmp_keysub
            // '.Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub

            // 'ログ出力用
            // str_logkey = fldname_keymain & " = " & tmp_keymain & "、" & _
            // fldname_keysub & " = " & tmp_keysub

            // 'データ有無チェック
            // Dim tmp_fldvalueumuchk As String = ""
            // For cntjj = 3 To readtbl.Columns.Count - 1
            // tmp_fldvalueumuchk = tmp_fldvalueumuchk & Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
            // Next
            // If tmp_fldvalueumuchk = "" Then
            // cvflg = False
            // End If

            // If cvflg Then

            // '移行対象件数カウント
            // cvtaisyocnt = cvtaisyocnt + 1

            // '親データ有無チェック
            // If list_basekeydata.Contains(fldvalue_key) = False Then
            // keychkflg = False
            // 'ログ出力
            // Dim log_key As String = "hydata_setubilst-hy_guid"
            // Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE
            // hash_log.Clear()
            // hash_log.Add(log_key, log_value)
            // End If

            // '重複チェック
            // If keychkflg And list_chkduplicate.Contains(fldvalue_key) Then
            // keychkflg = False
            // 'ログ出力
            // Dim log_key As String = "hydata_setubilst-hy_guid"
            // Dim log_value As String = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP
            // hash_log.Clear()
            // hash_log.Add(log_key, log_value)
            // End If

            // 'データ取得、移行処理
            // If keychkflg Then

            // '作業用変数
            // Dim tmp_fldname As String = ""
            // Dim fldname_syogo As String = ""
            // Dim fldname_log As String = ""
            // Dim fldvalue As String = ""

            // '移行値取得
            // For cntjj = 0 To readtbl.Columns.Count - 1

            // '項目名取得
            // tmp_fldname = Typ.ToStr(readtbl.Rows(0).Item(cntjj)).Trim
            // fldname_log = tmp_fldname.Replace(vbLf, "-")
            // fldname_syogo = tmp_fldname.Replace(vbLf, STR_SPLIT_1)

            // '登録値取得
            // fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

            // If fldname_syogo <> "項目名" And fldname_syogo <> "物件NO" And fldname_syogo <> "部屋NO" And fldvalue <> "" Then

            // Dim value_total As String = fldname_syogo & STR_SPLIT_1 & fldvalue

            // '設備マスターと一致した場合
            // If hash_setubimst.Contains(value_total) Then

            // '項目guidを取得
            // .Vari_Komok_guid = hash_setubimst(value_total)

            // '固定値
            // .Vari_Override_kbn = 0
            // .Vari_Komok_name = ""
            // .Vari_Disp1name = ""
            // .Vari_Disp2name = ""
            // .Vari_Disp3name = ""
            // .Vari_Disp1iconguid = "00000000-0000-0000-0000-000000000000"
            // .Vari_Disp2iconguid = "00000000-0000-0000-0000-000000000000"
            // .Vari_Disp3iconguid = "00000000-0000-0000-0000-000000000000"
            // .Vari_History = DefHistory

            // '部屋guid変換
            // .Vari_Hy_guid = hash_hyguid(fldvalue_key)

            // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            // '挿入処理
            // Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, normalflg)

            // '挿入処理が正常終了したレコードのキーを重複チェック用に格納
            // If normalflg Then
            // If list_chkduplicate.Contains(fldvalue_key) = False Then
            // list_chkduplicate.Add(fldvalue_key)
            // End If
            // setubicvcnt = setubicvcnt + 1
            // End If

            // Else

            // 'マスタ不一致ログ出力
            // Dim log_key As String = "hydata_setubilst-komok_guid"
            // Dim log_value As String = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & "設備" & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            // Dim hash_komklog As New SafeDictionary<string, string> From {{log_key, log_value}}

            // 'ログ出力メッセージ整形
            // tmp_hash.Clear()
            // hash_cvitem.Clear()
            // Call LogSetting.Set_Log_Value_KomkErr(hash_komklog, tmp_hash, hash_cvitem, str_logkey & "、" & fldname_log, tblname, tmp_logcnt, sortlist_log)

            // 'ログ出力
            // If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = totalrowcnt - 1) Then
            // Dim tmp_cnt As Integer = 0
            // '挿入
            // For Each logvalue In sortlist_log
            // Dim tmp_sql_insert As String = logvalue.Value
            // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
            // Next
            // '初期化
            // tmp_logcnt = 0
            // sortlist_log.Clear()
            // End If

            // End If

            // End If

            // Next

            // Else

            // '親データ無し、重複チェックエラーログ出力
            // 'ログ出力メッセージ整形
            // tmp_hash.Clear()
            // hash_cvitem.Clear()
            // Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

            // 'ログ出力
            // If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = totalrowcnt - 1) Then
            // Dim tmp_cnt As Integer = 0
            // '挿入
            // For Each logvalue In sortlist_log
            // Dim tmp_sql_insert As String = logvalue.Value
            // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
            // Next
            // '初期化
            // tmp_logcnt = 0
            // sortlist_log.Clear()
            // End If

            // End If

            // End If

            // End If

            // '-------------------------------
            // '件数取得
            // '-------------------------------
            // If setubicvcnt > 0 Then
            // tmp_cvcnt = tmp_cvcnt + 1
            // End If

            // '-------------------------------
            // 'プログレスバー更新/進捗率表示
            // '-------------------------------
            // '件数取得
            // 'Dim pgbcnt As Integer = 0
            // 'If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
            // '    pgbcnt = cntii
            // 'ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
            // '    Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
            // 'ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
            // '    pgbcnt = pgbtotalcnt
            // 'End If

            // Dim pgbcnt As Integer = 0
            // If totalrowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
            // pgbcnt = cntii + 1
            // ElseIf totalrowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
            // Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, pgbcnt)
            // ElseIf totalrowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
            // pgbcnt = pgbtotalcnt
            // End If

            // '表示
            // If pgbcnt <> 0 Then
            // Call obj_pgb.pgbsettingPart(pgbcnt)
            // Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
            // End If

            // Next

            // End With

            // '************************
            // '終了処理
            // '************************

            // '中間ファイル件数を取得
            // midrowcnt = cvtaisyocnt

            // '移行件数を取得
            // cvrowcnt = tmp_cvcnt

            // '調整件数を取得
            // conditioncnt = tmp_condcnt

            // 'クローズ処理
            // Call excelfile.ExcelFile_ReadClose(con_read)

            // '返却
            // Return rtn

            // End Function
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_汎用専用処理 -del end

            /// <summary>
            /// 親マスタ取得→リスト格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_basedata"></param>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT キー FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_basedata);

            }

            public string Get_UseQry(ref string sortstr)
            {
                return default;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,szen_no) FROM szendata_szen AS SZ ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 部屋guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_HyGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 CONVERT(VARCHAR,bk_no) + '-' + CONVERT(VARCHAR,hy_no) ";
                tmp_sql = tmp_sql + " 	,hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

        }

    }

    #endregion
    // 20161028 物件/部屋鍵取得方法修正 -add sta
    #region 鍵タイトルマスタ(汎用用)

    public class M_kagi_title_BaseMid_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】 
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="filename"></param>
            /// <param name="sheetname"></param>
            /// <param name="midrowcnt"></param>
            /// <param name="cvrowcnt"></param>
            /// <param name="conditioncnt"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                bool normalflg;                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値
                int tmpcnt = 0;                                       // クエリ実行用作業変数

                var model_cvitem = new Model.M_kagi_title_BaseMid_Model();            // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

                // ************************
                // 作業準備
                // ************************

                // データ取得
                string tmptblname = CommonModule.PRE_TBL_NAME + sheetname;
                string tmp_sql = " SELECT * FROM " + tmptblname;
                var readtbl = new DataTable();
                int rowcnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, ref rtn);

                // オープン処理失敗時は処理を抜ける
                if (rtn == false)
                {
                    return rtn;
                }

                // テーブル名/フィールド名セット
                string tblname = "m_kagi_title";
                string fldnamegrp = "kagi_kbn,kagi_no,kagi_name,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // プログレスバー初期化
                int pgbtotalcnt = 0;          // プログレスバー総件数初期化
                int pgbbasecnt = 100;         // 実件数で表示するか否かの基準値

                if (rowcnt <= pgbbasecnt)
                {
                    pgbtotalcnt = rowcnt;
                }
                else
                {
                    pgbtotalcnt = (int)Math.Round(Math.Ceiling(rowcnt / (double)pgbbasecnt));
                }
                obj_pgb.pgbInitPart(pgbtotalcnt);

                // 鍵タイトルマスタをデフォルト値で一括更新する
                if (rowcnt > 0)
                {
                    string tmp_sql_kagititle_def = " UPDATE m_kagi_title SET kagi_name = '鍵' + CONVERT(VARCHAR,kagi_no) WHERE kagi_kbn IN (1,2) ";
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_kagititle_def, ref tmpcnt);
                }

                // ************************
                // 処理開始
                // ************************

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol_main - 1].ColumnName.Trim();
                string fldname_keysub = readtbl.Columns[keycol_sub - 1].ColumnName.Trim();

                for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
                {

                    // 中断処理
                    Application.DoEvents();
                    if (CommonModule.CancelFlg)
                    {
                        return rtn;
                    }

                    // 作業用変数
                    string fldname = "";
                    string fldvalue = "";

                    // キー値取得
                    string fldvalue_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string fldvalue_keysub = Typ.ToStr(readtbl.Rows[cntii][keycol_sub - 1]).Trim();
                    string fldvalue_key = fldvalue_keymain + "-" + fldvalue_keysub;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + fldvalue_keymain + "、" + fldname_keysub + " = " + fldvalue_keysub;
                    // 鍵区分はプログラム内部で使用する項目のためログ出力用に成形する
                    str_logkey = str_logkey.Replace("鍵区分 = 1", "物件鍵タイトルマスタ");
                    str_logkey = str_logkey.Replace("鍵区分 = 2", "部屋鍵タイトルマスタ");

                    // データ取得
                    for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                    {

                        // 項目名取得
                        fldname = readtbl.Columns[cntjj].ColumnName.Trim();

                        // 登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();

                        switch (fldname ?? "")
                        {
                            case "鍵区分":
                                {
                                    model_cvitem.Vari_Kagi_kbn = fldvalue;
                                    break;
                                }
                            case "鍵No":
                                {
                                    model_cvitem.Vari_Kagi_no = fldvalue;
                                    break;
                                }
                            case "鍵名称":
                                {
                                    model_cvitem.Vari_Kagi_name = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (skipflg == false)
                    {

                        // 更新処理
                        string tmp_sql_kagititle = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(" UPDATE m_kagi_title SET kagi_name = '", hash_cvitem["kagi_name"]), "' WHERE kagi_kbn = "), hash_cvitem["kagi_kbn"]), " AND kagi_no = "), hash_cvitem["kagi_no"]));
                        normalflg = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_kagititle, ref tmpcnt);

                        // 更新処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                        }

                    }

                    // -------------------
                    // ログ出力
                    // -------------------

                    // ログ出力メッセージ整形
                    LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, ref tmp_logcnt, ref sortlist_log);

                    // ログ出力
                    if (sortlist_log.Count >= CommonModule.Log_OutputCnt | sortlist_log.Count != 0 & cntii == rowcnt - 1)
                    {
                        int tmp_cnt = 0;
                        // 挿入
                        foreach (var logvalue in sortlist_log)
                        {
                            string tmp_sql_insert = logvalue.Value;
                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_cnt);
                        }
                        // 初期化
                        tmp_logcnt = 0;
                        sortlist_log.Clear();
                    }

                    // -------------------------------
                    // プログレスバー更新/進捗率表示
                    // -------------------------------
                    // 件数取得
                    int pgbcnt = 0;
                    if (rowcnt <= pgbbasecnt)        // 基準件数(100件)以下の場合は実件数を取得
                    {
                        pgbcnt = cntii + 1;
                    }
                    else if (rowcnt > cntii + 2)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 2)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                    {
                        pgbcnt = pgbtotalcnt;
                    }

                    // 表示
                    if (pgbcnt != 0)
                    {
                        obj_pgb.pgbsettingPart(pgbcnt);
                        obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt);
                    }


                }

                // ************************
                // 終了処理
                // ************************

                // 中間ファイル件数を取得
                midrowcnt = rowcnt;

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 返却
                return rtn;

            }

            /// <summary>
            /// 親マスタ取得→リスト格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT jisya_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_basedata);

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public string Get_UseQry(ref string sortstr)
            {
                return default;

            }

            /// <summary>
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT CONVERT(varchar,kagi_kbn) + '-' + CONVERT(varchar,kagi_no) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion
    // 20161028 物件/部屋鍵取得方法修正 -add end
}