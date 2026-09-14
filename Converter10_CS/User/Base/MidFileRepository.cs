using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Repository
{

    #region 既存用中間ファイルのヘッダーを革命10のヘッダーに変換して値を紐付ける

    public class Chk_middata_Repository
    {

        /// <summary>
        /// 家主イベント情報の変換とチェック処理
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="filename"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_owdata_event_mid(SqlConnection sqlcnnv10, string filename, string sheetname)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            var tmp_condcnt = default(int);                                      // 調整件数格納

            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
            var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
            var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
            int tmp_logcnt = 0;                                   // ログ出力時のソート用

            var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
            bool normalflg;                                        // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値

            var model_cvitem = new Model.Owdata_event_Model();            // 移行値格納用モデル初期化
            int keycol = 1;                                       // キー列

            // ************************
            // 作業準備
            // ************************

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (!rtn)
            {
                return rtn;
            }

            // テーブル名/フィールド名セット
            string tblname_base = "owdata";
            string tblname = "owdata_event";
            string fldnamegrp = "ow_no,event_kbn,event_cnt,event_ymd,event_data";

            // ************************
            // 処理開始
            // ************************

            // ヘッダー行取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

            // キーヘッダー名取得
            string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol));

            for (int cntii = 1, loopTo = rowcnt; cntii <= loopTo; cntii++)
            {

                // 中断処理
                Application.DoEvents();
                if (CommonModule.CancelFlg)
                {
                    excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                    return rtn;
                }

                // 行取得
                var datarowvalue = (object[,])((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[cntii + startrow - 1, 1], wsheet.Cells[cntii + startrow - 1, columncnt]]).Value;

                // キー値取得
                string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol]);
                model_cvitem.Vari_Ow_no = tmp_keymain;

                // -------------------
                // イベントカウント初期化
                // -------------------
                int cvitemcnt = 0;

                // -------------
                // 列単位処理
                // -------------
                for (int cntjj = 2, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                    string fldvalue = "";
                    if (datarowvalue[startrow - 1, cntjj] is not null)
                    {
                        fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                    }

                    // サブキー取得
                    string tmp_keysub = "";
                    switch (fldname ?? "")
                    {
                        case "年賀状":
                            {
                                tmp_keysub = 1.ToString();
                                break;
                            }
                        case "暑中お見舞い":
                            {
                                tmp_keysub = 2.ToString();
                                break;
                            }
                        case "誕生日":
                            {
                                tmp_keysub = 3.ToString();
                                break;
                            }
                        case "お歳暮":
                            {
                                tmp_keysub = 4.ToString();
                                break;
                            }
                        case "お中元":
                            {
                                tmp_keysub = 5.ToString();
                                break;
                            }
                    }

                    // イベントデータ取得
                    model_cvitem.Vari_Event_kbn = tmp_keysub;
                    model_cvitem.Vari_Event_data = fldvalue.Trim();

                    // 全キー取得
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub;

                    // ログ出力用データ格納(サブキーフィールド)
                    string fldname_keysub = fldname;
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub + " = " + fldvalue;

                    // 固定値
                    model_cvitem.Vari_Event_cnt = 1.ToString();
                    model_cvitem.Vari_Event_ymd = null;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

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
                            tmp_sql_insert = tmp_sql_insert.Replace(CommonModule.LOG_TMP_TABLENAME, CommonModule.LOG_TMP_MIDCHKTABLENAME);
                            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_cnt);
                        }
                        // 初期化
                        tmp_logcnt = 0;
                        sortlist_log.Clear();
                    }

                }


            }

            // ************************
            // 終了処理
            // ************************

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            // 返却
            return rtn;

        }

        /// <summary>
        /// 家主メモ情報の変換とチェック処理
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="filename"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_owdata_memo_mid(SqlConnection sqlcnnv10, string filename, string sheetname)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            var tmp_condcnt = default(int);                                      // 調整件数格納

            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
            var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
            var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
            int tmp_logcnt = 0;                                   // ログ出力時のソート用

            var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
            bool normalflg;                                        // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値

            var model_cvitem = new Model.Owdata_memo_Model();            // 移行値格納用モデル初期化
            int keycol = 1;                                       // キー列

            // ************************
            // 作業準備
            // ************************

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (!rtn)
            {
                return rtn;
            }

            // テーブル名/フィールド名セット
            string tblname_base = "owdata";
            string tblname = "owdata_memo";
            string fldnamegrp = "ow_no,memo_no,memo,history";

            // ************************
            // 処理開始
            // ************************

            // ヘッダー行取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

            // キーヘッダー名取得
            string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol));

            for (int cntii = 1, loopTo = rowcnt; cntii <= loopTo; cntii++)
            {

                // 中断処理
                Application.DoEvents();
                if (CommonModule.CancelFlg)
                {
                    excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                    return rtn;
                }

                // 行取得
                var datarowvalue = (object[,])((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[cntii + startrow - 1, 1], wsheet.Cells[cntii + startrow - 1, columncnt]]).Value;

                // キー値取得
                string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol]);
                model_cvitem.Vari_Ow_no = tmp_keymain;

                // 備考カウント初期化
                int cvitemcnt = 0;

                // 移行値取得
                for (int cntjj = 2, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // サブキー取得
                    string tmp_keysub = (cntjj - 1).ToString();

                    // 全キー取得
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub;

                    // ログ出力用データ格納(サブキーフィールド)
                    string fldname_keysub = Conversions.ToString(headervalue(startrow - 1, cntjj));
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub;

                    string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                    string fldvalue = "";
                    if (datarowvalue[startrow - 1, cntjj] is not null)
                    {
                        fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                    }
                    model_cvitem.Vari_Memo = fldvalue;

                    // 固定値
                    model_cvitem.Vari_Memo_no = (cntjj - 1).ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // メモにデータが存在する場合に書込処理を行う
                    if (!string.IsNullOrEmpty(model_cvitem.Vari_Memo))
                    {

                        // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                        // データチェック
                        bool skipflg = false;
                        var hash_cvitem = new SafeDictionary<string, string>();
                        var hash_log = new SafeDictionary<string, string>();
                        skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

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

                    }

                }


            }

            // ************************
            // 終了処理
            // ************************

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            // 返却
            return rtn;

        }

    }

    #endregion

}