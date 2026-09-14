using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Repository
{

    #region 送信設定基本情報

    public class M_sendsetting_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.M_sendsetting_Model();           // 移行値格納用モデル初期化
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
                string tblname = "m_sendsetting";
                string fldnamegrp = "setting_guid,setting_sortorder,setting_name,keisai_siten,datalink_id," + "adconfirm_limitday,history,basedate,basedate_flg,emailaddr1," + "emailaddr2,emailaddr3,emailaddr4,emailaddr5";


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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol));

                // ---------------
                // データ部処理
                // ---------------
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
                    string fldvalue_key = tmp_keymain;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain;

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "送信設定順":
                                {
                                    model_cvitem.Vari_Setting_sortorder = fldvalue;
                                    break;
                                }
                            case "送信設定名":
                                {
                                    model_cvitem.Vari_Setting_name = fldvalue;
                                    break;
                                }
                            case "掲載支店":
                                {
                                    model_cvitem.Vari_Keisai_siten = fldvalue;
                                    break;
                                }
                            case "連動ID":
                                {
                                    model_cvitem.Vari_Datalink_id = fldvalue;
                                    break;
                                }
                            case "広告確認からの確認期間":
                                {
                                    model_cvitem.Vari_Adconfirm_limitday = fldvalue;
                                    break;
                                }
                            case "基準日":
                                {
                                    model_cvitem.Vari_Basedate = fldvalue;
                                    break;
                                }
                            case "基準日区分":
                                {
                                    model_cvitem.Vari_Basedate_flg = fldvalue;
                                    break;
                                }
                            case "Eメールアドレス1":
                                {
                                    model_cvitem.Vari_Emailaddr1 = fldvalue;
                                    break;
                                }
                            case "Eメールアドレス2":
                                {
                                    model_cvitem.Vari_Emailaddr2 = fldvalue;
                                    break;
                                }
                            case "Eメールアドレス3":
                                {
                                    model_cvitem.Vari_Emailaddr3 = fldvalue;
                                    break;
                                }
                            case "Eメールアドレス4":
                                {
                                    model_cvitem.Vari_Emailaddr4 = fldvalue;
                                    break;
                                }
                            case "Eメールアドレス5":
                                {
                                    model_cvitem.Vari_Emailaddr5 = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Setting_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                // Dim tmp_sql As String = ""
                // Dim flg As Boolean = True
                // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

                // キーをプログラム内で作成しているため重複不可情報を考慮する必要あり

            }

        }

    }

    #endregion

    #region 送信設定自社web情報

    public class M_site_sendsetting_jisyaweb_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_settingguid = new SafeDictionary<string, string>();                           // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.M_site_sendsetting_Model();      // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();

                // 革命10の入金項目マスタを取得
                var hash_nkinkomkmst = new SafeDictionary<string, string>();
                string tmp_sql = "SELECT nkin_no,nkin_name FROM m_nkin";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_nkinkomkmst);

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // 送信設定guid取得
                Set_SettingGuid(sqlcnnv10, ref hash_settingguid);

                // テーブル名/フィールド名セット
                string tblname = "m_site_sendsetting";
                string fldnamegrp = "sitesetting_guid,setting_guid,site_no,site_id,site_sendumu," + "site_data,site_password";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_settingguid, ref list_basekeydata);

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1;

                    // 作業用変数
                    var hash_cvxmlitem = new SafeDictionary<string, string>();

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "送信設定順":
                                {
                                    model_cvitem.Vari_Setting_guid = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }
                            case "ポータルサイトID":
                                {
                                    model_cvitem.Vari_Site_id = fldvalue;
                                    break;
                                }
                            case "サイト別送信有無":
                                {
                                    model_cvitem.Vari_Site_sendumu = fldvalue;
                                    break;
                                }
                            case "ポータルサイトパスワード":
                                {
                                    model_cvitem.Vari_Site_password = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(fldvalue), " ", fldvalue));
                                    break;
                                }

                            default:
                                {
                                    if (!string.IsNullOrEmpty(fldname))
                                    {
                                        hash_cvxmlitem.Add(fldname, fldvalue);
                                    }

                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Sitesetting_guid = Guid.NewGuid().ToString();

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["setting_guid"] = hash_settingguid[Strings.StrConv(Conversions.ToString(hash_cvitem["setting_guid"]), VbStrConv.Narrow)];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // xmlの移行
                            Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem, hash_nkinkomkmst);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // ポータルサイトパスワード一括更新 (半角スペース " " を設定していた分を空文字へ変換)
                int tmpcnt = 0;
                string passwordupdateqry = " UPDATE " + tblname + " SET site_password = '' WHERE site_no = 10 AND site_password = ' ' ";
                DBExec.Exec_NonQuery(sqlcnnv10, passwordupdateqry, ref tmpcnt);

                // 共通設定情報へ個別設定No1を反映
                // 20160908 個別処理を共通処理へコピーする処理の制御修正 -chg sta
                // Dim tmp_sortstr As String = ""  '使用しない
                // Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                // DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                if (cvrowcnt != 0)
                {
                    string tmp_sortstr = "";  // 使用しない
                    string kobetuinfotocommonqry = Get_UseQry(ref tmp_sortstr);
                    DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, ref tmpcnt);
                }
                // 20160908 個別処理を共通処理へコピーする処理の制御修正 -chg end
                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 共通設定情報更新クエリの作成
            /// </summary>
            /// <param name="sortstr"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " UPDATE m_site_sendsetting SET ";
                tmp_sql = tmp_sql + " 	 site_id = (SELECT site_id FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 10) ";
                tmp_sql = tmp_sql + " 	,site_sendumu = (SELECT site_sendumu FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 10) ";
                tmp_sql = tmp_sql + " 	,site_data = (SELECT site_data FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 10) ";
                tmp_sql = tmp_sql + " 	,site_password = (SELECT site_password FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 10) ";
                tmp_sql = tmp_sql + " WHERE setting_guid = ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0 ";
                tmp_sql = tmp_sql + " ) ";
                tmp_sql = tmp_sql + " AND site_no = 10 ";

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) ";
                tmp_sql = tmp_sql + " FROM hydata_kokoku AS HYKOKOKU ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 送信設定guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_SettingGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = " SELECT setting_sortorder,setting_guid FROM m_sendsetting ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 送信設定のxml項目の移行
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_cvxmlitem"></param>
            /// <remarks></remarks>
            public void Set_XmlItem(SqlConnection sqlcnnv10, SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_xmlitem, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                string sitesettingguid = Conversions.ToString(hash_cvitem["sitesetting_guid"]);   // 更新対象選択用

                // 「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                string xmlupdateqry_xmlpre = "";
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre + "<WmpSendSettingModel xmlns:i=" + "\"" + "http://www.w3.org/2001/XMLSchema-instance" + "\"" + ">";
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre + "<__identity xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/System" + "\"" + " i:nil=" + "\"" + "true" + "\"" + " />";
                string xmlupdateqry_xmlpost = "";
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost + "<ProductModelList />";
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost + "</WmpSendSettingModel>";

                // グループ毎にXml文字列を作成
                // 画像
                string tmp_xml_gazo = Get_XmlStr_Gazo(hash_xmlitem);
                // 入金項目
                string tmp_xml_nkin = Get_XmlStr_Nkinkomk(hash_xmlitem, hash_nkinkomkmst);
                // その他
                string tmp_xml_other = Get_XmlStr_Other(hash_cvitem, hash_xmlitem);

                // Xml文字列の結合
                string xmlupdateqry_main = xmlupdateqry_xmlpre + tmp_xml_gazo + tmp_xml_nkin + tmp_xml_other + xmlupdateqry_xmlpost;

                // 更新用クエリの作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " DECLARE @tmp_xml XML ";
                tmp_sql = tmp_sql + " SET @tmp_xml = CAST('" + xmlupdateqry_main + "' AS XML) ";
                tmp_sql = tmp_sql + " UPDATE m_site_sendsetting SET site_data = @tmp_xml ";
                tmp_sql = tmp_sql + " WHERE sitesetting_guid = '" + sitesettingguid + "' ";

                // 更新
                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

            /// <summary>
            /// 画像に関するXml文字列の作成
            /// </summary>
            /// <param name="hash"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Gazo(SafeDictionary<string, string> hash)
            {

                string rtn_str = "";
                string xml_pre = "<GazoSettingList xmlns:a=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">";
                string xml_post = "</GazoSettingList>";

                int gazocnt = 30;
                var tmp_gazono = new string[gazocnt + 1];
                var tmp_gazosyuV7 = new string[gazocnt + 1];
                var tmp_gazonoV7 = new string[gazocnt + 1];
                var tmp_gazosyurendo = new string[gazocnt + 1];

                // 画像に関する情報を配列へ格納
                foreach (var item in hash)
                {

                    string fldname = Conversions.ToString(item.Key);
                    string value = Conversions.ToString(item.Value);

                    switch (true)
                    {
                        case object _ when (fldname.Replace("画像No", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("画像No", ""));
                                tmp_gazono[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("革命側の画像種別", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("革命側の画像種別", ""));
                                tmp_gazosyuV7[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("革命側の画像タイトルNo", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("革命側の画像タイトルNo", ""));
                                tmp_gazonoV7[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("連動側の画像種別", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("連動側の画像種別", ""));
                                tmp_gazosyurendo[gazono] = value;
                                break;
                            }
                    }

                }

                // 文字列へ成形
                var xml_gazoitem = new System.Text.StringBuilder();
                for (int cntii = 1, loopTo = gazocnt; cntii <= loopTo; cntii++)
                {
                    if (tmp_gazosyurendo[cntii] != "0" & tmp_gazosyuV7[cntii] != "0" & tmp_gazonoV7[cntii] != "0")
                    {
                        xml_gazoitem.Append("<a:PortalGazoSettingModel>");
                        xml_gazoitem.Append("<a:FK8LocationKbn>" + tmp_gazosyuV7[cntii] + "</a:FK8LocationKbn>");
                        xml_gazoitem.Append("<a:FK8LocationNo>" + tmp_gazonoV7[cntii] + "</a:FK8LocationNo>");
                        xml_gazoitem.Append("<a:No>" + tmp_gazono[cntii] + "</a:No>");
                        xml_gazoitem.Append("<a:PortalGazoSyubetu>" + tmp_gazosyurendo[cntii] + "</a:PortalGazoSyubetu>");
                        xml_gazoitem.Append("</a:PortalGazoSettingModel>");
                    }
                }

                // 結合
                string xml_gazo = xml_pre + xml_gazoitem.ToString() + xml_post;
                rtn_str = xml_gazo;

                return rtn_str;

            }

            /// <summary>
            /// 入金項目に関するXml文字列の作成
            /// </summary>
            /// <param name="hash"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Nkinkomk(SafeDictionary<string, string> hash_xmlitem, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                string rtn_str = "";
                string xml_pre = "<NkinKomokDict xmlns:a=" + "\"" + "http://schemas.microsoft.com/2003/10/Serialization/Arrays" + "\"" + ">";
                string xml_post = "</NkinKomokDict>";
                int nkincnt = 11;
                int cntindex = 1;
                var xml_nkinkomk = new string[nkincnt + 1];
                var hash_nkinkomkitem = new SafeDictionary<string, string>();

                // 入金項目を移行用に成形する
                var hash_cvxmlitem = new SafeDictionary<string, string>();
                hash_cvxmlitem = Chg_NkinKomk(hash_xmlitem, hash_nkinkomkmst);

                // 賃料
                hash_nkinkomkitem.Add("item_key", "1");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_賃料"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_賃料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_賃料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 管理費共益費
                hash_nkinkomkitem.Add("item_key", "2");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_共益費/管理費"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_共益費管理費");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_共益費/管理費No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 敷金
                hash_nkinkomkitem.Add("item_key", "6");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_敷金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["敷金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_敷金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_敷金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 礼金
                hash_nkinkomkitem.Add("item_key", "5");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_礼金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["礼金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_礼金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_礼金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 保証金
                hash_nkinkomkitem.Add("item_key", "7");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_保証金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["保証金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_保証金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_保証金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 償却金
                hash_nkinkomkitem.Add("item_key", "9");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_償却金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["償却金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_償却金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_償却金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 敷引(解約引き)
                hash_nkinkomkitem.Add("item_key", "12");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_敷引(解約引き)"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["敷引(解約引き)表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_敷引金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_敷引(解約引き)No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 更新料
                hash_nkinkomkitem.Add("item_key", "14");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_更新料"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["更新料表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_更新料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_更新料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 仲介手数料
                hash_nkinkomkitem.Add("item_key", "15");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_仲介手数料"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["仲介手数料表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_仲介手数料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_仲介手数料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // その他一時金
                hash_nkinkomkitem.Add("item_key", "18");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_その他一時金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["その他一時金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_その他一時金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_その他一時金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();


                // その他月額費用
                hash_nkinkomkitem.Add("item_key", "17");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_その他月額費用"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_その他費用");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_その他月額費用No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 結合
                string xml_total = "";
                string xml_nkintotal = "";
                for (int cntii = 1, loopTo = nkincnt; cntii <= loopTo; cntii++)
                    xml_nkintotal = xml_nkintotal + xml_nkinkomk[cntii];
                xml_total = xml_pre + xml_nkintotal + xml_post;

                rtn_str = xml_total;

                return rtn_str;

            }

            /// <summary>
            /// 入金項目の編集
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_xml"></param>
            /// <param name="hash_nkinkomkmst"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, string> Chg_NkinKomk(SafeDictionary<string, string> hash_xml, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                var rtn_hash = new SafeDictionary<string, string>();
                var tmp_hash = new SafeDictionary<string, string>();

                int tmp_nkincnt = 16;
                var tmp_nkinkomkname = new string[tmp_nkincnt + 1];
                var tmp_nkinkomkkbn = new string[tmp_nkincnt + 1];
                var tmp_fldname = new string[tmp_nkincnt + 1];
                // -----------------------
                // 入金項目取得
                // -----------------------
                foreach (var item in hash_xml)
                {

                    string midfldname = Conversions.ToString(item.Key);
                    string value = Conversions.ToString(item.Value);

                    switch (midfldname ?? "")
                    {
                        case "入金項目_賃料1":
                            {
                                tmp_nkinkomkname[1] = value;
                                tmp_fldname[1] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料1":
                            {
                                tmp_nkinkomkkbn[1] = value;
                                break;
                            }
                        case "入金項目_賃料2":
                            {
                                tmp_nkinkomkname[2] = value;
                                tmp_fldname[2] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料2":
                            {
                                tmp_nkinkomkkbn[2] = value;
                                break;
                            }
                        case "入金項目_賃料3":
                            {
                                tmp_nkinkomkname[3] = value;
                                tmp_fldname[3] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料3":
                            {
                                tmp_nkinkomkkbn[3] = value;
                                break;
                            }
                        case "入金項目_共益費/管理費1":
                            {
                                tmp_nkinkomkname[4] = value;
                                tmp_fldname[4] = midfldname;
                                break;
                            }
                        case "入金項目区分_共益費/管理費1":
                            {
                                tmp_nkinkomkkbn[4] = value;
                                break;
                            }
                        case "入金項目_共益費/管理費2":
                            {
                                tmp_nkinkomkname[5] = value;
                                tmp_fldname[5] = midfldname;
                                break;
                            }
                        case "入金項目区分_共益費/管理費2":
                            {
                                tmp_nkinkomkkbn[5] = value;
                                break;
                            }
                        case "入金項目_共益費/管理費3":
                            {
                                tmp_nkinkomkname[6] = value;
                                tmp_fldname[6] = midfldname;
                                break;
                            }
                        case "入金項目区分_共益費/管理費3":
                            {
                                tmp_nkinkomkkbn[6] = value;
                                break;
                            }
                        case "入金項目_敷金":
                            {
                                tmp_nkinkomkname[7] = value;
                                tmp_fldname[7] = midfldname;
                                break;
                            }
                        case "入金項目区分_敷金":
                            {
                                tmp_nkinkomkkbn[7] = value;
                                break;
                            }
                        case "入金項目_礼金":
                            {
                                tmp_nkinkomkname[8] = value;
                                tmp_fldname[8] = midfldname;
                                break;
                            }
                        case "入金項目区分_礼金":
                            {
                                tmp_nkinkomkkbn[8] = value;
                                break;
                            }
                        case "入金項目_保証金":
                            {
                                tmp_nkinkomkname[9] = value;
                                tmp_fldname[9] = midfldname;
                                break;
                            }
                        case "入金項目区分_保証金":
                            {
                                tmp_nkinkomkkbn[9] = value;
                                break;
                            }
                        case "入金項目_償却金":
                            {
                                tmp_nkinkomkname[10] = value;
                                tmp_fldname[10] = midfldname;
                                break;
                            }
                        case "入金項目区分_償却金":
                            {
                                tmp_nkinkomkkbn[10] = value;
                                break;
                            }
                        case "入金項目_敷引(解約引き)":
                            {
                                tmp_nkinkomkname[11] = value;
                                tmp_fldname[11] = midfldname;
                                break;
                            }
                        case "入金項目区分_敷引(解約引き)":
                            {
                                tmp_nkinkomkkbn[11] = value;
                                break;
                            }
                        case "入金項目_更新料":
                            {
                                tmp_nkinkomkname[12] = value;
                                tmp_fldname[12] = midfldname;
                                break;
                            }
                        case "入金項目区分_更新料":
                            {
                                tmp_nkinkomkkbn[12] = value;
                                break;
                            }
                        case "入金項目_仲介手数料":
                            {
                                tmp_nkinkomkname[13] = value;
                                tmp_fldname[13] = midfldname;
                                break;
                            }
                        case "入金項目区分_仲介手数料":
                            {
                                tmp_nkinkomkkbn[13] = value;
                                break;
                            }
                        case "入金項目_その他一時金":
                            {
                                tmp_nkinkomkname[14] = value;
                                tmp_fldname[14] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他一時金":
                            {
                                tmp_nkinkomkkbn[14] = value;
                                break;
                            }
                        case "入金項目_その他月額費用1":
                            {
                                tmp_nkinkomkname[15] = value;
                                tmp_fldname[15] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他月額費用1":
                            {
                                tmp_nkinkomkkbn[15] = value;
                                break;
                            }
                        case "入金項目_その他月額費用2":
                            {
                                tmp_nkinkomkname[16] = value;
                                tmp_fldname[16] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他月額費用2":
                            {
                                tmp_nkinkomkkbn[16] = value;
                                break;
                            }

                        default:
                            {
                                tmp_hash.Add(midfldname, value);
                                break;
                            }
                    }

                }

                // -----------------------
                // 入金項目をXml移行用に成形
                // -----------------------
                var tmp_nkinno = new string[tmp_nkincnt + 1];
                var tmp_dispnkinname = new string[tmp_nkincnt + 1];
                for (int cntii = 1, loopTo = tmp_nkincnt; cntii <= loopTo; cntii++)
                {

                    string tmp_relnkinstr = tmp_nkinkomkname[cntii] + "-" + EtcMethod.Get_Nkinruiname(ref tmp_nkinkomkkbn[cntii]);
                    tmp_nkinno[cntii] = "";
                    tmp_dispnkinname[cntii] = "";

                    if (CommonModule.Hash_Rel_Nkinkomk.ContainsKey(tmp_relnkinstr))
                    {
                        tmp_nkinno[cntii] = Conversions.ToString(CommonModule.Hash_Rel_Nkinkomk[tmp_relnkinstr]);
                        tmp_dispnkinname[cntii] = Conversions.ToString(Operators.ConcatenateObject(tmp_nkinno[cntii] + ":", hash_nkinkomkmst[tmp_nkinno[cntii]]));
                    }

                }

                // -----------------------
                // 入金項目の再格納
                // -----------------------
                // 結合する項目
                Set_NkinkomkAddToHash(1, 3, tmp_nkinno, tmp_dispnkinname, "入金項目_賃料", ref tmp_hash);
                Set_NkinkomkAddToHash(4, 6, tmp_nkinno, tmp_dispnkinname, "入金項目_共益費/管理費", ref tmp_hash);
                Set_NkinkomkAddToHash(15, 16, tmp_nkinno, tmp_dispnkinname, "入金項目_その他月額費用", ref tmp_hash);

                // 結合不要の項目
                for (int cntii = 7; cntii <= 14; cntii++)
                {
                    tmp_hash.Add(tmp_fldname[cntii], tmp_dispnkinname[cntii]);
                    tmp_hash.Add(tmp_fldname[cntii] + "No", tmp_nkinno[cntii]);
                }

                rtn_hash = tmp_hash;

                return rtn_hash;

            }

            /// <summary>
            /// 入金項目を結合して1つのデータにまとめる
            /// まとめたデータをハッシュテーブルへ格納する
            /// </summary>
            /// <param name="cntsta"></param>
            /// <param name="cntend"></param>
            /// <param name="tmp_nkinno"></param>
            /// <param name="tmp_dispnkinname"></param>
            /// <param name="hashkey"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_NkinkomkAddToHash(int cntsta, int cntend, string[] tmp_nkinno, string[] tmp_dispnkinname, string hashkey, ref SafeDictionary<string, string> hash)
            {

                string nkinno = "";
                string dispname = "";
                for (int cntii = cntsta, loopTo = cntend; cntii <= loopTo; cntii++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(tmp_nkinno[cntii], "", false)))
                    {
                        nkinno = Conversions.ToString(Operators.ConcatenateObject(nkinno + ",", tmp_nkinno[cntii]));
                        dispname = Conversions.ToString(Operators.ConcatenateObject(dispname + " / ", tmp_dispnkinname[cntii]));
                    }
                }
                if (!string.IsNullOrEmpty(nkinno))
                {
                    nkinno = nkinno.Remove(0, 1);
                }
                if (!string.IsNullOrEmpty(dispname))
                {
                    dispname = dispname.Remove(0, 3);
                }
                hash.Add(hashkey, dispname);
                hash.Add(hashkey + "No", nkinno);

            }

            /// <summary>
            /// 各入金項目毎にXml文字列作成
            /// </summary>
            /// <param name="hash_nkin"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Chg_NkinkomkToXml(SafeDictionary<string, string> hash_nkin)
            {

                string rtn_str = "";

                // 置換用/ハッシュテーブル値取得用文字列
                string item_key = "item_key";
                string item_disp = "item_disp";
                string item_tani = "item_tani";
                string item_komk = "item_komk";
                string item_nkinno = "item_nkinno";

                // Xmlのベースを作成
                var xml_nkinkomkbase = new System.Text.StringBuilder();
                xml_nkinkomkbase.Append("<a:KeyValueOfintPortalNkinKomokLjh4bohd>");
                xml_nkinkomkbase.Append("<a:Key>" + item_key + "</a:Key>");
                xml_nkinkomkbase.Append("<a:Value>");
                xml_nkinkomkbase.Append("<DisplayString>" + item_disp + "</DisplayString>");
                xml_nkinkomkbase.Append("<EnKagetu>" + item_tani + "</EnKagetu>");
                xml_nkinkomkbase.Append("<NkinKomok>" + item_komk + "</NkinKomok>");
                xml_nkinkomkbase.Append("<NkinNoList>");
                xml_nkinkomkbase.Append(item_nkinno);
                xml_nkinkomkbase.Append("</NkinNoList>");
                xml_nkinkomkbase.Append("</a:Value>");
                xml_nkinkomkbase.Append("</a:KeyValueOfintPortalNkinKomokLjh4bohd>");

                // 各要素で置換してい成形
                string xml_nkinkomk = xml_nkinkomkbase.ToString();
                xml_nkinkomk = xml_nkinkomk.Replace(item_key, Conversions.ToString(hash_nkin[item_key]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_disp, Conversions.ToString(hash_nkin[item_disp]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_tani, Conversions.ToString(hash_nkin[item_tani]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_komk, Conversions.ToString(hash_nkin[item_komk]));
                string tmp_nkinnogrp_tinryo = Conversions.ToString(hash_nkin[item_nkinno]);
                string[] tmp_nkinno_tinryo = tmp_nkinnogrp_tinryo.Split(',');
                string nkinno_tinryo = "";
                for (int cntii = 0, loopTo = Information.UBound(tmp_nkinno_tinryo); cntii <= loopTo; cntii++)
                {
                    if (!string.IsNullOrEmpty(tmp_nkinno_tinryo[cntii]))
                    {
                        nkinno_tinryo = nkinno_tinryo + "<a:int>" + tmp_nkinno_tinryo[cntii] + "</a:int>";
                    }
                }
                xml_nkinkomk = xml_nkinkomk.Replace(item_nkinno, nkinno_tinryo);

                rtn_str = xml_nkinkomk;

                return rtn_str;

            }

            /// <summary>
            /// その他項目に関するXml文字列の作成
            /// </summary>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_xmlitem"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Other(SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_xmlitem)
            {

                string rtn_str = "";
                int othercnt = 19;
                var xml_other = new string[othercnt + 1];
                int cntindex = 1;

                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<Password>", hash_cvitem["site_password"]), "</Password>"));
                cntindex = cntindex + 1;           // パスワード
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<SiteID>", hash_cvitem["site_id"]), "</SiteID>"));
                cntindex = cntindex + 1;                       // ポータルサイトID
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<SiteSendUmu>", Interaction.IIf(Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(hash_cvitem["site_sendumu"], "1", false)), "true", "false")), "</SiteSendUmu>"));
                cntindex = cntindex + 1;           // サイト別送信有無
                xml_other[cntindex] = "<UseCommonSetting>" + "false" + "</UseCommonSetting>";
                cntindex = cntindex + 1;                         // 共通設定使用有無
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BantiDispSetting>", hash_xmlitem["番地以降の表示"]), "</BantiDispSetting>"));
                cntindex = cntindex + 1;   // 番地以降の表示
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BkHyDispSetting>", hash_xmlitem["物件部屋の表示"]), "</BkHyDispSetting>"));
                cntindex = cntindex + 1;   // 物件部屋の表示
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<EmailAddr1>", hash_xmlitem["Eメールアドレス1"]), "</EmailAddr1>"));
                cntindex = cntindex + 1;          // Eメールアドレス1
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<EmailAddr2>", hash_xmlitem["Eメールアドレス2"]), "</EmailAddr2>"));
                cntindex = cntindex + 1;          // Eメールアドレス2
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<EmailAddr3>", hash_xmlitem["Eメールアドレス3"]), "</EmailAddr3>"));
                cntindex = cntindex + 1;          // Eメールアドレス3
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<EmailAddr4>", hash_xmlitem["Eメールアドレス4"]), "</EmailAddr4>"));
                cntindex = cntindex + 1;         // Eメールアドレス4
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<EmailAddr5>", hash_xmlitem["Eメールアドレス5"]), "</EmailAddr5>"));
                cntindex = cntindex + 1;         // Eメールアドレス5

                // 固定分
                var tmp_strb = new System.Text.StringBuilder();
                tmp_strb.Append("<GazoSettingList>");
                tmp_strb.Append("<WmpGazoSettingModel>");
                tmp_strb.Append("<FK8LocationKbn xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</FK8LocationKbn>");
                tmp_strb.Append("<FK8LocationNo xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</FK8LocationNo>");
                tmp_strb.Append("<No xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</No>");
                tmp_strb.Append("<PortalGazoSyubetu xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</PortalGazoSyubetu>");
                tmp_strb.Append("</WmpGazoSettingModel>");
                tmp_strb.Append("<WmpGazoSettingModel>");
                tmp_strb.Append("<FK8LocationKbn xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">2</FK8LocationKbn>");
                tmp_strb.Append("<FK8LocationNo xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</FK8LocationNo>");
                tmp_strb.Append("<No xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">16</No>");
                tmp_strb.Append("<PortalGazoSyubetu xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">5</PortalGazoSyubetu>");
                tmp_strb.Append("</WmpGazoSettingModel>");
                tmp_strb.Append("</GazoSettingList>");
                xml_other[cntindex] = tmp_strb.ToString();
                cntindex = cntindex + 1;

                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<KagiInfo1>", hash_xmlitem["鍵情報1"]), "</KagiInfo1>"));
                cntindex = cntindex + 1;                  // 鍵情報1
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<KagiInfo2>", hash_xmlitem["鍵情報2"]), "</KagiInfo2>"));
                cntindex = cntindex + 1;                  // 鍵情報2
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<KagiInfo3>", hash_xmlitem["鍵情報3"]), "</KagiInfo3>"));
                cntindex = cntindex + 1;                  // 鍵情報3
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<KurasapoIgaiMapDispSetting>", hash_xmlitem["くらさぽ以外への地図表示"]), "</KurasapoIgaiMapDispSetting>"));
                cntindex = cntindex + 1;   // くらさぽ以外への地図表示
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<KurasapoMapDispSetting>", hash_xmlitem["くらさぽへの地図表示"]), "</KurasapoMapDispSetting>"));
                cntindex = cntindex + 1;              // くらさぽへの地図表示
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<MovieDispSetting>", hash_xmlitem["動画の表示"]), "</MovieDispSetting>"));
                cntindex = cntindex + 1;                       // 動画の表示

                // 動画の表示順序
                string tmp_replacedoga = "item_doga";
                string tmp_xml_gazojyunjo = "<MovieYusenList xmlns:a=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.N3Lib.Utys" + "\"" + ">" + tmp_replacedoga + "</MovieYusenList>";

                int dogacnt = 4;
                var xml_doga = new string[dogacnt + 1];
                var tmp_dogastrb = new System.Text.StringBuilder();
                for (int cntii = 1, loopTo = dogacnt; cntii <= loopTo; cntii++)
                {
                    string dogavalue = "";
                    switch (cntii)
                    {
                        case 1:
                            {
                                dogavalue = "物件動画１";
                                break;
                            }
                        case 2:
                            {
                                dogavalue = "物件動画２";
                                break;
                            }
                        case 3:
                            {
                                dogavalue = "部屋動画１";
                                break;
                            }
                        case 4:
                            {
                                dogavalue = "部屋動画２";
                                break;
                            }
                    }
                    tmp_dogastrb.Append("<a:DataKeyValue>");
                    tmp_dogastrb.Append("<__identity xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/System" + "\"" + " i:nil=" + "\"" + "true" + "\"" + " />");
                    tmp_dogastrb.Append("<a:_alwaysVisible>true</a:_alwaysVisible>");
                    tmp_dogastrb.Append("<a:_key>" + cntii.ToString() + "</a:_key>");
                    tmp_dogastrb.Append("<a:_keyDescription i:nil=" + "\"" + "true" + "\"" + " />");
                    tmp_dogastrb.Append("<a:_obj i:nil=" + "\"" + "true" + "\"" + " />");
                    tmp_dogastrb.Append("<a:_value>" + dogavalue + "</a:_value>");
                    tmp_dogastrb.Append("</a:DataKeyValue>");
                    xml_doga[cntii] = tmp_dogastrb.ToString();
                    tmp_dogastrb.Clear();
                }
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(hash_xmlitem["動画の表示"], "2", false)))
                {
                    tmp_xml_gazojyunjo = tmp_xml_gazojyunjo.Replace(tmp_replacedoga, xml_doga[int.Parse(Conversions.ToString(hash_xmlitem["動画の優先項目1"]))] + xml_doga[int.Parse(Conversions.ToString(hash_xmlitem["動画の優先項目2"]))] + xml_doga[int.Parse(Conversions.ToString(hash_xmlitem["動画の優先項目3"]))] + xml_doga[int.Parse(Conversions.ToString(hash_xmlitem["動画の優先項目4"]))]);
                }
                else
                {
                    tmp_xml_gazojyunjo = tmp_xml_gazojyunjo.Replace(tmp_replacedoga, xml_doga[1] + xml_doga[2] + xml_doga[3] + xml_doga[4]);
                }
                xml_other[cntindex] = tmp_xml_gazojyunjo;

                // 結合
                string tmp_str = "";
                for (int cntii = 1, loopTo1 = othercnt; cntii <= loopTo1; cntii++)
                    tmp_str = tmp_str + xml_other[cntii];

                rtn_str = tmp_str;

                return rtn_str;

            }

        }

    }

    #endregion

    #region 送信設定HOMES情報

    public class M_site_sendsetting_homes_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_settingguid = new SafeDictionary<string, string>();                           // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.M_site_sendsetting_Model();      // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();

                // 革命10の入金項目マスタを取得
                var hash_nkinkomkmst = new SafeDictionary<string, string>();
                string tmp_sql = "SELECT nkin_no,nkin_name FROM m_nkin";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_nkinkomkmst);

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // 送信設定guid取得
                Set_SettingGuid(sqlcnnv10, ref hash_settingguid);

                // テーブル名/フィールド名セット
                string tblname = "m_site_sendsetting";
                string fldnamegrp = "sitesetting_guid,setting_guid,site_no,site_id,site_sendumu," + "site_data,site_password";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_settingguid, ref list_basekeydata);

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1;

                    // 作業用変数
                    var hash_cvxmlitem = new SafeDictionary<string, string>();

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "送信設定順":
                                {
                                    model_cvitem.Vari_Setting_guid = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }
                            case "ポータルサイトID":
                                {
                                    model_cvitem.Vari_Site_id = fldvalue;
                                    break;
                                }
                            case "サイト別送信有無":
                                {
                                    model_cvitem.Vari_Site_sendumu = fldvalue;
                                    break;
                                }
                            case "ポータルサイトパスワード":
                                {
                                    model_cvitem.Vari_Site_password = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(fldvalue), " ", fldvalue));
                                    break;
                                }

                            default:
                                {
                                    if (!string.IsNullOrEmpty(fldname))
                                    {
                                        hash_cvxmlitem.Add(fldname, fldvalue);
                                    }

                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Sitesetting_guid = Guid.NewGuid().ToString();

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // FTPユーザーID、パスワード有無チェック
                    Chk_FTPData(sqlcnnv10, tblname, ref hash_cvitem, ref hash_cvxmlitem);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["setting_guid"] = hash_settingguid[Strings.StrConv(Conversions.ToString(hash_cvitem["setting_guid"]), VbStrConv.Narrow)];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // xmlの移行
                            Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem, hash_nkinkomkmst);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // ポータルサイトパスワード一括更新 (半角スペース " " を設定していた分を空文字へ変換)
                int tmpcnt = 0;
                string passwordupdateqry = " UPDATE " + tblname + " SET site_password = '' WHERE site_no = 20 AND site_password = ' ' ";
                DBExec.Exec_NonQuery(sqlcnnv10, passwordupdateqry, ref tmpcnt);

                // 共通設定情報へ個別設定No1を反映
                // 20160908 個別処理を共通処理へコピーする処理の制御修正 -chg sta
                // Dim tmp_sortstr As String = ""  '使用しない
                // Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                // DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                if (cvrowcnt != 0)
                {
                    string tmp_sortstr = "";  // 使用しない
                    string kobetuinfotocommonqry = Get_UseQry(ref tmp_sortstr);
                    DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, ref tmpcnt);
                }
                // 20160908 個別処理を共通処理へコピーする処理の制御修正 -chg end

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 共通設定情報更新クエリの作成
            /// </summary>
            /// <param name="sortstr"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " UPDATE m_site_sendsetting SET ";
                tmp_sql = tmp_sql + " 	 site_id = (SELECT site_id FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 20) ";
                tmp_sql = tmp_sql + " 	,site_sendumu = (SELECT site_sendumu FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 20) ";
                tmp_sql = tmp_sql + " 	,site_data = (SELECT site_data FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 20) ";
                tmp_sql = tmp_sql + " 	,site_password = (SELECT site_password FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 20) ";
                tmp_sql = tmp_sql + " WHERE setting_guid = ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0 ";
                tmp_sql = tmp_sql + " ) ";
                tmp_sql = tmp_sql + " AND site_no = 20 ";

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) ";
                tmp_sql = tmp_sql + " FROM hydata_kokoku AS HYKOKOKU ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 送信設定guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_SettingGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = " SELECT setting_sortorder,setting_guid FROM m_sendsetting ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 送信設定のxml項目の移行
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_cvxmlitem"></param>
            /// <remarks></remarks>
            public void Set_XmlItem(SqlConnection sqlcnnv10, SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_xmlitem, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                string sitesettingguid = Conversions.ToString(hash_cvitem["sitesetting_guid"]);   // 更新対象選択用

                // 「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                string xmlupdateqry_xmlpre = "";
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre + "<HomesSendSettingModel xmlns:i=" + "\"" + "http://www.w3.org/2001/XMLSchema-instance" + "\"" + ">";
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre + "<__identity xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/System" + "\"" + " i:nil=" + "\"" + "true" + "\"" + " />";
                string xmlupdateqry_xmlpost = "";
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost + "</HomesSendSettingModel>";

                // グループ毎にXml文字列を作成
                // 画像
                string tmp_xml_gazo = Get_XmlStr_Gazo(hash_xmlitem);
                // 入金項目
                string tmp_xml_nkin = Get_XmlStr_Nkinkomk(hash_xmlitem, hash_nkinkomkmst);
                // その他
                string tmp_xml_other = Get_XmlStr_Other(hash_cvitem, hash_xmlitem);

                // Xml文字列の結合
                string xmlupdateqry_main = xmlupdateqry_xmlpre + tmp_xml_gazo + tmp_xml_nkin + tmp_xml_other + xmlupdateqry_xmlpost;
                // Dim xmlupdateqry_main As String = xmlupdateqry_xmlpre & tmp_xml_nkin & xmlupdateqry_xmlpost

                // 更新用クエリの作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " DECLARE @tmp_xml XML ";
                tmp_sql = tmp_sql + " SET @tmp_xml = CAST('" + xmlupdateqry_main + "' AS XML) ";
                tmp_sql = tmp_sql + " UPDATE m_site_sendsetting SET site_data = @tmp_xml ";
                tmp_sql = tmp_sql + " WHERE sitesetting_guid = '" + sitesettingguid + "' ";

                // 更新
                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

            /// <summary>
            /// 画像に関するXml文字列の作成
            /// </summary>
            /// <param name="hash"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Gazo(SafeDictionary<string, string> hash)
            {

                string rtn_str = "";
                string xml_pre = "<GazoSettingList xmlns:a=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">";
                string xml_post = "</GazoSettingList>";

                int gazocnt = 30;
                var tmp_gazono = new string[gazocnt + 1];
                var tmp_gazosyuV7 = new string[gazocnt + 1];
                var tmp_gazonoV7 = new string[gazocnt + 1];
                var tmp_gazosyurendo = new string[gazocnt + 1];

                // 画像に関する情報を配列へ格納
                foreach (var item in hash)
                {

                    string fldname = Conversions.ToString(item.Key);
                    string value = Conversions.ToString(item.Value);

                    switch (true)
                    {
                        case object _ when (fldname.Replace("画像No", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("画像No", ""));
                                tmp_gazono[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("革命側の画像種別", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("革命側の画像種別", ""));
                                tmp_gazosyuV7[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("革命側の画像タイトルNo", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("革命側の画像タイトルNo", ""));
                                tmp_gazonoV7[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("連動側の画像種別", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("連動側の画像種別", ""));
                                tmp_gazosyurendo[gazono] = value;
                                break;
                            }
                    }

                }

                // 文字列へ成形
                var xml_gazoitem = new System.Text.StringBuilder();
                for (int cntii = 1, loopTo = gazocnt; cntii <= loopTo; cntii++)
                {
                    if (tmp_gazosyurendo[cntii] != "0" & tmp_gazosyuV7[cntii] != "0" & tmp_gazonoV7[cntii] != "0")
                    {
                        xml_gazoitem.Append("<a:PortalGazoSettingModel>");
                        xml_gazoitem.Append("<a:FK8LocationKbn>" + tmp_gazosyuV7[cntii] + "</a:FK8LocationKbn>");
                        xml_gazoitem.Append("<a:FK8LocationNo>" + tmp_gazonoV7[cntii] + "</a:FK8LocationNo>");
                        xml_gazoitem.Append("<a:No>" + tmp_gazono[cntii] + "</a:No>");
                        xml_gazoitem.Append("<a:PortalGazoSyubetu>" + tmp_gazosyurendo[cntii] + "</a:PortalGazoSyubetu>");
                        xml_gazoitem.Append("</a:PortalGazoSettingModel>");
                    }
                }

                // 結合
                string xml_gazo = xml_pre + xml_gazoitem.ToString() + xml_post;
                rtn_str = xml_gazo;

                return rtn_str;

            }

            /// <summary>
            /// 入金項目に関するXml文字列の作成
            /// </summary>
            /// <param name="hash"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Nkinkomk(SafeDictionary<string, string> hash_xmlitem, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                string rtn_str = "";
                string xml_pre = "<NkinKomokDict xmlns:a=" + "\"" + "http://schemas.microsoft.com/2003/10/Serialization/Arrays" + "\"" + ">";
                string xml_post = "</NkinKomokDict>";
                int nkincnt = 13;
                int cntindex = 1;
                var xml_nkinkomk = new string[nkincnt + 1];
                var hash_nkinkomkitem = new SafeDictionary<string, string>();

                // 入金項目を移行用に成形する
                var hash_cvxmlitem = new SafeDictionary<string, string>();
                hash_cvxmlitem = Chg_NkinKomk(hash_xmlitem, hash_nkinkomkmst);

                // 賃料
                hash_nkinkomkitem.Add("item_key", "1");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_賃料"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_賃料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_賃料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 管理費/共益費
                hash_nkinkomkitem.Add("item_key", "2");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_管理費/共益費"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_共益費管理費");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_管理費/共益費No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 敷金
                hash_nkinkomkitem.Add("item_key", "6");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_敷金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["敷金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_敷金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_敷金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 礼金
                hash_nkinkomkitem.Add("item_key", "5");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_礼金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["礼金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_礼金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_礼金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 保証金
                hash_nkinkomkitem.Add("item_key", "7");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_保証金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["保証金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_保証金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_保証金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 権利金
                hash_nkinkomkitem.Add("item_key", "8");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_権利金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["権利金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_権利金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_権利金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 償却/敷引金
                hash_nkinkomkitem.Add("item_key", "10");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_償却/敷引金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["償却/敷引金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_償却敷引金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_償却/敷引金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 更新料
                hash_nkinkomkitem.Add("item_key", "14");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_更新料"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["更新料表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_更新料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_更新料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 造作譲渡金
                hash_nkinkomkitem.Add("item_key", "16");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_造作譲渡金"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_造作譲渡金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_造作譲渡金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 仲介手数料
                hash_nkinkomkitem.Add("item_key", "15");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_仲介手数料"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_仲介手数料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_仲介手数料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 鍵交換費用
                hash_nkinkomkitem.Add("item_key", "20");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_鍵交換費用"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_鍵交換代等");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_鍵交換費用No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 室内清掃費用
                hash_nkinkomkitem.Add("item_key", "23");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_室内清掃費用"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_室内清掃費");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_室内清掃費用No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // その他月額費用
                hash_nkinkomkitem.Add("item_key", "17");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_その他月額費用"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_その他費用");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_その他月額費用No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 結合
                string xml_total = "";
                string xml_nkintotal = "";
                for (int cntii = 1, loopTo = nkincnt; cntii <= loopTo; cntii++)
                    xml_nkintotal = xml_nkintotal + xml_nkinkomk[cntii];
                xml_total = xml_pre + xml_nkintotal + xml_post;

                rtn_str = xml_total;

                return rtn_str;

            }

            /// <summary>
            /// 入金項目の編集
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_xml"></param>
            /// <param name="hash_nkinkomkmst"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, string> Chg_NkinKomk(SafeDictionary<string, string> hash_xml, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                var rtn_hash = new SafeDictionary<string, string>();
                var tmp_hash = new SafeDictionary<string, string>();

                int tmp_nkincnt = 19;
                var tmp_nkinkomkname = new string[tmp_nkincnt + 1];
                var tmp_nkinkomkkbn = new string[tmp_nkincnt + 1];
                var tmp_fldname = new string[tmp_nkincnt + 1];
                // -----------------------
                // 入金項目取得
                // -----------------------
                foreach (var item in hash_xml)
                {

                    string midfldname = Conversions.ToString(item.Key);
                    string value = Conversions.ToString(item.Value);

                    switch (midfldname ?? "")
                    {
                        case "入金項目_賃料1":
                            {
                                tmp_nkinkomkname[1] = value;
                                tmp_fldname[1] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料1":
                            {
                                tmp_nkinkomkkbn[1] = value;
                                break;
                            }
                        case "入金項目_賃料2":
                            {
                                tmp_nkinkomkname[2] = value;
                                tmp_fldname[2] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料2":
                            {
                                tmp_nkinkomkkbn[2] = value;
                                break;
                            }
                        case "入金項目_賃料3":
                            {
                                tmp_nkinkomkname[3] = value;
                                tmp_fldname[3] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料3":
                            {
                                tmp_nkinkomkkbn[3] = value;
                                break;
                            }
                        case "入金項目_管理費/共益費1":
                            {
                                tmp_nkinkomkname[4] = value;
                                tmp_fldname[4] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費/共益費1":
                            {
                                tmp_nkinkomkkbn[4] = value;
                                break;
                            }
                        case "入金項目_管理費/共益費2":
                            {
                                tmp_nkinkomkname[5] = value;
                                tmp_fldname[5] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費/共益費2":
                            {
                                tmp_nkinkomkkbn[5] = value;
                                break;
                            }
                        case "入金項目_管理費/共益費3":
                            {
                                tmp_nkinkomkname[6] = value;
                                tmp_fldname[6] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費/共益費3":
                            {
                                tmp_nkinkomkkbn[6] = value;
                                break;
                            }
                        case "入金項目_敷金":
                            {
                                tmp_nkinkomkname[7] = value;
                                tmp_fldname[7] = midfldname;
                                break;
                            }
                        case "入金項目区分_敷金":
                            {
                                tmp_nkinkomkkbn[7] = value;
                                break;
                            }
                        case "入金項目_礼金":
                            {
                                tmp_nkinkomkname[8] = value;
                                tmp_fldname[8] = midfldname;
                                break;
                            }
                        case "入金項目区分_礼金":
                            {
                                tmp_nkinkomkkbn[8] = value;
                                break;
                            }
                        case "入金項目_保証金":
                            {
                                tmp_nkinkomkname[9] = value;
                                tmp_fldname[9] = midfldname;
                                break;
                            }
                        case "入金項目区分_保証金":
                            {
                                tmp_nkinkomkkbn[9] = value;
                                break;
                            }
                        case "入金項目_権利金":
                            {
                                tmp_nkinkomkname[10] = value;
                                tmp_fldname[10] = midfldname;
                                break;
                            }
                        case "入金項目区分_権利金":
                            {
                                tmp_nkinkomkkbn[10] = value;
                                break;
                            }
                        case "入金項目_償却/敷引金":
                            {
                                tmp_nkinkomkname[11] = value;
                                tmp_fldname[11] = midfldname;
                                break;
                            }
                        case "入金項目区分_償却/敷引金":
                            {
                                tmp_nkinkomkkbn[11] = value;
                                break;
                            }
                        case "入金項目_更新料":
                            {
                                tmp_nkinkomkname[12] = value;
                                tmp_fldname[12] = midfldname;
                                break;
                            }
                        case "入金項目区分_更新料":
                            {
                                tmp_nkinkomkkbn[12] = value;
                                break;
                            }
                        case "入金項目_造作譲渡金":
                            {
                                tmp_nkinkomkname[13] = value;
                                tmp_fldname[13] = midfldname;
                                break;
                            }
                        case "入金項目区分_造作譲渡金":
                            {
                                tmp_nkinkomkkbn[13] = value;
                                break;
                            }
                        case "入金項目_仲介手数料":
                            {
                                tmp_nkinkomkname[14] = value;
                                tmp_fldname[14] = midfldname;
                                break;
                            }
                        case "入金項目区分_仲介手数料":
                            {
                                tmp_nkinkomkkbn[14] = value;
                                break;
                            }
                        case "入金項目_鍵交換費用":
                            {
                                tmp_nkinkomkname[15] = value;
                                tmp_fldname[15] = midfldname;
                                break;
                            }
                        case "入金項目区分_鍵交換費用":
                            {
                                tmp_nkinkomkkbn[15] = value;
                                break;
                            }
                        case "入金項目_室内清掃費用":
                            {
                                tmp_nkinkomkname[16] = value;
                                tmp_fldname[16] = midfldname;
                                break;
                            }
                        case "入金項目区分_室内清掃費用":
                            {
                                tmp_nkinkomkkbn[16] = value;
                                break;
                            }
                        case "入金項目_その他月額費用1":
                            {
                                tmp_nkinkomkname[17] = value;
                                tmp_fldname[17] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他費用1":
                            {
                                tmp_nkinkomkkbn[17] = value;
                                break;
                            }
                        case "入金項目_その他月額費用2":
                            {
                                tmp_nkinkomkname[18] = value;
                                tmp_fldname[18] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他費用2":
                            {
                                tmp_nkinkomkkbn[18] = value;
                                break;
                            }
                        case "入金項目_その他月額費用3":
                            {
                                tmp_nkinkomkname[19] = value;
                                tmp_fldname[19] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他費用3":
                            {
                                tmp_nkinkomkkbn[19] = value;
                                break;
                            }

                        default:
                            {
                                tmp_hash.Add(midfldname, value);
                                break;
                            }
                    }

                }

                // -----------------------
                // 入金項目をXml移行用に成形
                // -----------------------
                var tmp_nkinno = new string[tmp_nkincnt + 1];
                var tmp_dispnkinname = new string[tmp_nkincnt + 1];
                for (int cntii = 1, loopTo = tmp_nkincnt; cntii <= loopTo; cntii++)
                {

                    string tmp_relnkinstr = tmp_nkinkomkname[cntii] + "-" + EtcMethod.Get_Nkinruiname(ref tmp_nkinkomkkbn[cntii]);
                    tmp_nkinno[cntii] = "";
                    tmp_dispnkinname[cntii] = "";

                    if (CommonModule.Hash_Rel_Nkinkomk.ContainsKey(tmp_relnkinstr))
                    {
                        tmp_nkinno[cntii] = Conversions.ToString(CommonModule.Hash_Rel_Nkinkomk[tmp_relnkinstr]);
                        tmp_dispnkinname[cntii] = Conversions.ToString(Operators.ConcatenateObject(tmp_nkinno[cntii] + ":", hash_nkinkomkmst[tmp_nkinno[cntii]]));
                    }

                }

                // -----------------------
                // 入金項目の再格納
                // -----------------------
                // 結合する項目
                Set_NkinkomkAddToHash(1, 3, tmp_nkinno, tmp_dispnkinname, "入金項目_賃料", ref tmp_hash);
                Set_NkinkomkAddToHash(4, 6, tmp_nkinno, tmp_dispnkinname, "入金項目_管理費/共益費", ref tmp_hash);
                Set_NkinkomkAddToHash(17, 19, tmp_nkinno, tmp_dispnkinname, "入金項目_その他月額費用", ref tmp_hash);

                // 結合不要の項目
                for (int cntii = 7; cntii <= 16; cntii++)
                {
                    tmp_hash.Add(tmp_fldname[cntii], tmp_dispnkinname[cntii]);
                    tmp_hash.Add(tmp_fldname[cntii] + "No", tmp_nkinno[cntii]);
                }

                rtn_hash = tmp_hash;

                return rtn_hash;

            }

            /// <summary>
            /// 入金項目を結合して1つのデータにまとめる
            /// まとめたデータをハッシュテーブルへ格納する
            /// </summary>
            /// <param name="cntsta"></param>
            /// <param name="cntend"></param>
            /// <param name="tmp_nkinno"></param>
            /// <param name="tmp_dispnkinname"></param>
            /// <param name="hashkey"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_NkinkomkAddToHash(int cntsta, int cntend, string[] tmp_nkinno, string[] tmp_dispnkinname, string hashkey, ref SafeDictionary<string, string> hash)
            {

                string nkinno = "";
                string dispname = "";
                for (int cntii = cntsta, loopTo = cntend; cntii <= loopTo; cntii++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(tmp_nkinno[cntii], "", false)))
                    {
                        nkinno = Conversions.ToString(Operators.ConcatenateObject(nkinno + ",", tmp_nkinno[cntii]));
                        dispname = Conversions.ToString(Operators.ConcatenateObject(dispname + " / ", tmp_dispnkinname[cntii]));
                    }
                }
                if (!string.IsNullOrEmpty(nkinno))
                {
                    nkinno = nkinno.Remove(0, 1);
                }
                if (!string.IsNullOrEmpty(dispname))
                {
                    dispname = dispname.Remove(0, 3);
                }
                hash.Add(hashkey, dispname);
                hash.Add(hashkey + "No", nkinno);

            }

            /// <summary>
            /// 各入金項目毎にXml文字列作成
            /// </summary>
            /// <param name="hash_nkin"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Chg_NkinkomkToXml(SafeDictionary<string, string> hash_nkin)
            {

                string rtn_str = "";

                // 置換用/ハッシュテーブル値取得用文字列
                string item_key = "item_key";
                string item_disp = "item_disp";
                string item_tani = "item_tani";
                string item_komk = "item_komk";
                string item_nkinno = "item_nkinno";

                // Xmlのベースを作成
                var xml_nkinkomkbase = new System.Text.StringBuilder();
                xml_nkinkomkbase.Append("<a:KeyValueOfintPortalNkinKomokLjh4bohd>");
                xml_nkinkomkbase.Append("<a:Key>" + item_key + "</a:Key>");
                xml_nkinkomkbase.Append("<a:Value>");
                xml_nkinkomkbase.Append("<DisplayString>" + item_disp + "</DisplayString>");
                xml_nkinkomkbase.Append("<EnKagetu>" + item_tani + "</EnKagetu>");
                xml_nkinkomkbase.Append("<NkinKomok>" + item_komk + "</NkinKomok>");
                xml_nkinkomkbase.Append("<NkinNoList>");
                xml_nkinkomkbase.Append(item_nkinno);
                xml_nkinkomkbase.Append("</NkinNoList>");
                xml_nkinkomkbase.Append("</a:Value>");
                xml_nkinkomkbase.Append("</a:KeyValueOfintPortalNkinKomokLjh4bohd>");

                // 各要素で置換してい成形
                string xml_nkinkomk = xml_nkinkomkbase.ToString();
                xml_nkinkomk = xml_nkinkomk.Replace(item_key, Conversions.ToString(hash_nkin[item_key]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_disp, Conversions.ToString(hash_nkin[item_disp]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_tani, Conversions.ToString(hash_nkin[item_tani]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_komk, Conversions.ToString(hash_nkin[item_komk]));
                string tmp_nkinnogrp_tinryo = Conversions.ToString(hash_nkin[item_nkinno]);
                string[] tmp_nkinno_tinryo = tmp_nkinnogrp_tinryo.Split(',');
                string nkinno_tinryo = "";
                for (int cntii = 0, loopTo = Information.UBound(tmp_nkinno_tinryo); cntii <= loopTo; cntii++)
                {
                    if (!string.IsNullOrEmpty(tmp_nkinno_tinryo[cntii]))
                    {
                        nkinno_tinryo = nkinno_tinryo + "<a:int>" + tmp_nkinno_tinryo[cntii] + "</a:int>";
                    }
                }
                xml_nkinkomk = xml_nkinkomk.Replace(item_nkinno, nkinno_tinryo);

                rtn_str = xml_nkinkomk;

                return rtn_str;

            }

            /// <summary>
            /// その他項目に関するXml文字列の作成
            /// </summary>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_xmlitem"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Other(SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_xmlitem)
            {

                string rtn_str = "";
                int othercnt = 100;
                var xml_other = new string[othercnt + 1];
                int cntindex = 1;

                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<Password>", hash_cvitem["site_password"]), "</Password>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<SiteID>", hash_cvitem["site_id"]), "</SiteID>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<SiteSendUmu>", Interaction.IIf(Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(hash_cvitem["site_sendumu"], "1", false)), "true", "false")), "</SiteSendUmu>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = "<UseCommonSetting>" + "false" + "</UseCommonSetting>";
                cntindex = cntindex + 1;
                xml_other[cntindex] = "<AssignedSyuhenKbn>" + "15" + "</AssignedSyuhenKbn>";
                cntindex = cntindex + 1;   // 要確認
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BantiDispSetting>", hash_xmlitem["番地以降の表示"]), "</BantiDispSetting>"));
                cntindex = cntindex + 1;

                // グループ設定 sta
                string fldnamegrp = "所属グループ設定";
                string grpname = "grpname";
                string xml_grp = "";
                var tmp_strbgrp = new System.Text.StringBuilder();
                for (int cntgrp = 1; cntgrp <= 10; cntgrp++)
                {
                    tmp_strbgrp.Append("    <HomesSendSettingModel.BelongGroupModel>");
                    tmp_strbgrp.Append("      <__identity xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/System" + "\"" + " i:nil=" + "\"" + "true" + "\"" + " />");
                    tmp_strbgrp.Append("      <Id>" + grpname + "</Id>");
                    tmp_strbgrp.Append("      <IdNo>ID" + cntgrp.ToString() + "</IdNo>");
                    tmp_strbgrp.Append("    </HomesSendSettingModel.BelongGroupModel>");
                    string tmp_grpstr = tmp_strbgrp.ToString();
                    tmp_grpstr = tmp_grpstr.Replace(grpname, Conversions.ToString(hash_xmlitem[fldnamegrp + cntgrp.ToString()]));
                    xml_grp = xml_grp + tmp_grpstr;
                    tmp_strbgrp.Clear();
                }
                // 結合
                xml_other[cntindex] = "<BelongGroupList>" + xml_grp + "</BelongGroupList>";
                cntindex = cntindex + 1;
                // グループ設定 end

                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BkHyDispSetting>", hash_xmlitem["物件部屋の表示"]), "</BkHyDispSetting>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<FtpPassword>", hash_xmlitem["FTPパスワード"]), "</FtpPassword>"));
                cntindex = cntindex + 1;  // 要追加
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<FtpUserId>", hash_xmlitem["FTPユーザーID"]), "</FtpUserId>"));
                cntindex = cntindex + 1;       // 要追加

                // 固定分
                var tmp_strb = new System.Text.StringBuilder();
                tmp_strb.Append("<GazoSettingList>");
                tmp_strb.Append("<HomesGazoSettingModel>");
                tmp_strb.Append("<FK8LocationKbn xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</FK8LocationKbn>");
                tmp_strb.Append("<FK8LocationNo xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</FK8LocationNo>");
                tmp_strb.Append("<No xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</No>");
                tmp_strb.Append("<PortalGazoSyubetu xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</PortalGazoSyubetu>");
                tmp_strb.Append("</HomesGazoSettingModel>");
                tmp_strb.Append("<HomesGazoSettingModel>");
                tmp_strb.Append("<FK8LocationKbn xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">2</FK8LocationKbn>");
                tmp_strb.Append("<FK8LocationNo xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">1</FK8LocationNo>");
                tmp_strb.Append("<No xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">16</No>");
                tmp_strb.Append("<PortalGazoSyubetu xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">5</PortalGazoSyubetu>");
                tmp_strb.Append("</HomesGazoSettingModel>");
                tmp_strb.Append("</GazoSettingList>");
                xml_other[cntindex] = tmp_strb.ToString();
                cntindex = cntindex + 1;

                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<IsAsteriskOsusumePointScore>", hash_xmlitem["特別広告ポイント数"]), "</IsAsteriskOsusumePointScore>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<IsAsteriskPanorama>", hash_xmlitem["パノラマセット"]), "</IsAsteriskPanorama>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<IsAsteriskStaffComment>", hash_xmlitem["スタッフコメント"]), "</IsAsteriskStaffComment>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<KeisaiKakuninBi>", hash_xmlitem["掲載確認日"]), "</KeisaiKakuninBi>"));
                cntindex = cntindex + 1;

                // 結合
                string tmp_str = "";
                for (int cntii = 1, loopTo = othercnt; cntii <= loopTo; cntii++)
                {
                    if (xml_other[cntii] is null)
                    {
                        break;
                    }
                    tmp_str = tmp_str + xml_other[cntii];
                }

                rtn_str = tmp_str;

                return rtn_str;

            }

            /// <summary>
            /// FTPユーザーID、パスワードの有無チェック
            /// </summary>
            /// <param name="tblname"></param>
            /// <param name="hash_xmlitem"></param>
            /// <param name="hash_log"></param>
            /// <remarks></remarks>
            public void Chk_FTPData(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_cvitem, ref SafeDictionary<string, string> hash_xmlitem)
            {

                string ftpuserid = Conversions.ToString(hash_xmlitem["FTPユーザーID"]);
                string ftppassword = Conversions.ToString(hash_xmlitem["FTPパスワード"]);
                string log_key = tblname + "-" + "site_data";
                string errstr = "";

                // ログ出力用の項目設定
                string syorikomok = "送信設定HOMES情報";
                string taisyodata = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("送信設定順 = ", hash_cvitem["setting_guid"]), "、"), "サイトNo = "), hash_cvitem["site_no"]));

                // FTPユーザーIDチェック
                if (string.IsNullOrEmpty(ftpuserid))
                {

                    // 任意の文字列を格納
                    string tmp_ftpuserid = "ftpuserid";

                    // 再格納
                    hash_xmlitem["FTPユーザーID"] = tmp_ftpuserid;

                    // ログ出力
                    string syoriitem = "FTPユーザーID";
                    string syorikekka = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED;
                    string hubigein = CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED_ADDSTR;
                    string taisyo = CommonModule.LOG_TAISYO_DEFAULT;
                    string beforechgvalue = ftpuserid;
                    string afterchgvalue = tmp_ftpuserid;

                    // ログ挿入用へ加工
                    string logitem = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname);
                    // 挿入用クエリへ加工
                    string log_sql = LogSetting.Get_LogTblInsertQry(logitem, false);
                    // ログ挿入処理
                    int tmp_cnt = 0;
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, ref tmp_cnt);

                }

                // FTPパスワードチェック
                if (string.IsNullOrEmpty(ftppassword))
                {

                    // 任意の文字列を格納
                    string tmp_ftppassword = "ftppassword";

                    // 再格納
                    hash_xmlitem["FTPパスワード"] = tmp_ftppassword;

                    // ログ出力
                    string syoriitem = "FTPパスワード";
                    string syorikekka = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED;
                    string hubigein = CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED_ADDSTR;
                    string taisyo = CommonModule.LOG_TAISYO_DEFAULT;
                    string beforechgvalue = ftppassword;
                    string afterchgvalue = tmp_ftppassword;

                    // ログ挿入用へ加工
                    string logitem = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname);
                    // 挿入用クエリへ加工
                    string log_sql = LogSetting.Get_LogTblInsertQry(logitem, false);
                    // ログ挿入処理
                    int tmp_cnt = 0;
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, ref tmp_cnt);

                }

            }

        }

    }

    #endregion

    #region 送信設定athome情報

    public class M_site_sendsetting_athome_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_settingguid = new SafeDictionary<string, string>();                           // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.M_site_sendsetting_Model();      // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();

                // 革命10の入金項目マスタを取得
                var hash_nkinkomkmst = new SafeDictionary<string, string>();
                string tmp_sql = "SELECT nkin_no,nkin_name FROM m_nkin";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_nkinkomkmst);

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // 送信設定guid取得
                Set_SettingGuid(sqlcnnv10, ref hash_settingguid);

                // テーブル名/フィールド名セット
                string tblname = "m_site_sendsetting";
                string fldnamegrp = "sitesetting_guid,setting_guid,site_no,site_id,site_sendumu," + "site_data,site_password";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_settingguid, ref list_basekeydata);

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1;

                    // 作業用変数
                    var hash_cvxmlitem = new SafeDictionary<string, string>();

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "送信設定順":
                                {
                                    model_cvitem.Vari_Setting_guid = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }
                            case "ポータルサイトID":
                                {
                                    model_cvitem.Vari_Site_id = fldvalue;
                                    break;
                                }
                            case "サイト別送信有無":
                                {
                                    model_cvitem.Vari_Site_sendumu = fldvalue;
                                    break;
                                }
                            case "ポータルサイトパスワード":
                                {
                                    model_cvitem.Vari_Site_password = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(fldvalue), " ", fldvalue));
                                    break;
                                }

                            default:
                                {
                                    if (!string.IsNullOrEmpty(fldname))
                                    {
                                        hash_cvxmlitem.Add(fldname, fldvalue);
                                    }

                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Sitesetting_guid = Guid.NewGuid().ToString();

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 会社支店コード桁数チェック
                    Chk_Tenpocode(sqlcnnv10, tblname, ref hash_cvitem, ref hash_cvxmlitem);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["setting_guid"] = hash_settingguid[Strings.StrConv(Conversions.ToString(hash_cvitem["setting_guid"]), VbStrConv.Narrow)];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // xmlの移行
                            Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem, hash_nkinkomkmst);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // ポータルサイトパスワード一括更新 (半角スペース " " を設定していた分を空文字へ変換)
                int tmpcnt = 0;
                string passwordupdateqry = " UPDATE " + tblname + " SET site_password = '' WHERE site_no = 30 AND site_password = ' ' ";
                DBExec.Exec_NonQuery(sqlcnnv10, passwordupdateqry, ref tmpcnt);

                // 共通設定情報へ個別設定No1を反映
                // 20160908 個別処理を共通処理へコピーする処理の制御修正 -chg sta
                // Dim tmp_sortstr As String = ""  '使用しない
                // Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                // DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                if (cvrowcnt != 0)
                {
                    string tmp_sortstr = "";  // 使用しない
                    string kobetuinfotocommonqry = Get_UseQry(ref tmp_sortstr);
                    DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, ref tmpcnt);
                }
                // 20160908 個別処理を共通処理へコピーする処理の制御修正 -chg end
                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 共通設定情報更新クエリの作成
            /// </summary>
            /// <param name="sortstr"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " UPDATE m_site_sendsetting SET ";
                tmp_sql = tmp_sql + " 	 site_id = (SELECT site_id FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 30) ";
                tmp_sql = tmp_sql + " 	,site_sendumu = (SELECT site_sendumu FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 30) ";
                tmp_sql = tmp_sql + " 	,site_data = (SELECT site_data FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 30) ";
                tmp_sql = tmp_sql + " 	,site_password = (SELECT site_password FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 30) ";
                tmp_sql = tmp_sql + " WHERE setting_guid = ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0 ";
                tmp_sql = tmp_sql + " ) ";
                tmp_sql = tmp_sql + " AND site_no = 30 ";

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) ";
                tmp_sql = tmp_sql + " FROM hydata_kokoku AS HYKOKOKU ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 送信設定guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_SettingGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = " SELECT setting_sortorder,setting_guid FROM m_sendsetting ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 送信設定のxml項目の移行
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_cvxmlitem"></param>
            /// <remarks></remarks>
            public void Set_XmlItem(SqlConnection sqlcnnv10, SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_xmlitem, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                string sitesettingguid = Conversions.ToString(hash_cvitem["sitesetting_guid"]);   // 更新対象選択用

                // 「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                string xmlupdateqry_xmlpre = "";
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre + "<AthomeSendSettingModel xmlns:i=" + "\"" + "http://www.w3.org/2001/XMLSchema-instance" + "\"" + ">";
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre + "<__identity xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/System" + "\"" + " i:nil=" + "\"" + "true" + "\"" + " />";
                string xmlupdateqry_xmlpost = "";
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost + "</AthomeSendSettingModel>";

                // グループ毎にXml文字列を作成
                // 画像
                string tmp_xml_gazo = Get_XmlStr_Gazo(hash_xmlitem);
                // 入金項目
                string tmp_xml_nkin = Get_XmlStr_Nkinkomk(hash_xmlitem, hash_nkinkomkmst);
                // その他
                string tmp_xml_other = Get_XmlStr_Other(hash_cvitem, hash_xmlitem);

                // Xml文字列の結合
                string xmlupdateqry_main = xmlupdateqry_xmlpre + tmp_xml_gazo + tmp_xml_nkin + tmp_xml_other + xmlupdateqry_xmlpost;

                // 更新用クエリの作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " DECLARE @tmp_xml XML ";
                tmp_sql = tmp_sql + " SET @tmp_xml = CAST('" + xmlupdateqry_main + "' AS XML) ";
                tmp_sql = tmp_sql + " UPDATE m_site_sendsetting SET site_data = @tmp_xml ";
                tmp_sql = tmp_sql + " WHERE sitesetting_guid = '" + sitesettingguid + "' ";

                // 更新
                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

            /// <summary>
            /// 画像に関するXml文字列の作成
            /// </summary>
            /// <param name="hash"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Gazo(SafeDictionary<string, string> hash)
            {

                string rtn_str = "";
                string xml_pre = "<GazoSettingList xmlns:a=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">";
                string xml_post = "</GazoSettingList>";

                int gazocnt = 16;
                var tmp_gazono = new string[gazocnt + 1];
                var tmp_gazosyuV7 = new string[gazocnt + 1];
                var tmp_gazonoV7 = new string[gazocnt + 1];
                var tmp_gazosyurendo = new string[gazocnt + 1];

                // 画像に関する情報を配列へ格納
                foreach (var item in hash)
                {

                    string fldname = Conversions.ToString(item.Key);
                    string value = Conversions.ToString(item.Value);

                    switch (true)
                    {
                        case object _ when (fldname.Replace("画像No", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("画像No", ""));
                                tmp_gazono[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("革命側の画像種別", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("革命側の画像種別", ""));
                                tmp_gazosyuV7[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("革命側の画像タイトルNo", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("革命側の画像タイトルNo", ""));
                                tmp_gazonoV7[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("連動側の画像種別", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("連動側の画像種別", ""));
                                tmp_gazosyurendo[gazono] = value;
                                break;
                            }
                    }

                }

                // 文字列へ成形
                var xml_gazoitem = new System.Text.StringBuilder();
                for (int cntii = 1, loopTo = gazocnt; cntii <= loopTo; cntii++)
                {
                    if (tmp_gazosyurendo[cntii] != "0" & tmp_gazosyuV7[cntii] != "0" & tmp_gazonoV7[cntii] != "0")
                    {
                        xml_gazoitem.Append("<a:PortalGazoSettingModel>");
                        xml_gazoitem.Append("<a:FK8LocationKbn>" + tmp_gazosyuV7[cntii] + "</a:FK8LocationKbn>");
                        xml_gazoitem.Append("<a:FK8LocationNo>" + tmp_gazonoV7[cntii] + "</a:FK8LocationNo>");
                        xml_gazoitem.Append("<a:No>" + tmp_gazono[cntii] + "</a:No>");
                        xml_gazoitem.Append("<a:PortalGazoSyubetu>" + tmp_gazosyurendo[cntii] + "</a:PortalGazoSyubetu>");
                        xml_gazoitem.Append("</a:PortalGazoSettingModel>");
                    }
                }

                // 固定部分の追加
                var xml_gazokotei = new System.Text.StringBuilder();
                int koteicnt = 8;
                for (int cntjj = 1, loopTo1 = koteicnt; cntjj <= loopTo1; cntjj++)
                {
                    xml_gazokotei.Append("<a:PortalGazoSettingModel>");
                    xml_gazokotei.Append("<a:FK8LocationKbn>" + "10" + "</a:FK8LocationKbn>");
                    xml_gazokotei.Append("<a:FK8LocationNo>" + cntjj.ToString() + "</a:FK8LocationNo>");
                    xml_gazokotei.Append("<a:No>" + "10" + cntjj.ToString() + "</a:No>");
                    xml_gazokotei.Append("<a:PortalGazoSyubetu>" + "-1" + "</a:PortalGazoSyubetu>");
                    xml_gazokotei.Append("</a:PortalGazoSettingModel>");
                }

                // 結合
                string xml_gazo = xml_pre + xml_gazoitem.ToString() + xml_gazokotei.ToString() + xml_post;
                rtn_str = xml_gazo;

                return rtn_str;

            }

            /// <summary>
            /// 入金項目に関するXml文字列の作成
            /// </summary>
            /// <param name="hash"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Nkinkomk(SafeDictionary<string, string> hash_xmlitem, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                string rtn_str = "";
                string xml_pre = "<NkinKomokDict xmlns:a=" + "\"" + "http://schemas.microsoft.com/2003/10/Serialization/Arrays" + "\"" + ">";
                string xml_post = "</NkinKomokDict>";
                int nkincnt = 16;
                int cntindex = 1;
                var xml_nkinkomk = new string[nkincnt + 1];
                var hash_nkinkomkitem = new SafeDictionary<string, string>();

                // 入金項目を移行用に成形する
                var hash_cvxmlitem = new SafeDictionary<string, string>();
                hash_cvxmlitem = Chg_NkinKomk(hash_xmlitem, hash_nkinkomkmst);

                // 賃料
                hash_nkinkomkitem.Add("item_key", "1");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_賃料"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_賃料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_賃料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 共益費
                hash_nkinkomkitem.Add("item_key", "3");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_共益費"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_共益費");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_共益費No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 管理費
                hash_nkinkomkitem.Add("item_key", "4");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_管理費"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_管理費");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_管理費No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 敷金
                hash_nkinkomkitem.Add("item_key", "6");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_敷金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["敷金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_敷金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_敷金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 礼金
                hash_nkinkomkitem.Add("item_key", "5");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_礼金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["礼金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_礼金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_礼金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 保証金
                hash_nkinkomkitem.Add("item_key", "7");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_保証金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["保証金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_保証金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_保証金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 敷引金
                hash_nkinkomkitem.Add("item_key", "12");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_敷引金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["敷引金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_敷引金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_敷引金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 保証金償却
                hash_nkinkomkitem.Add("item_key", "9");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_保証金償却"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["保証金償却表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_償却金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_保証金償却No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 更新料
                hash_nkinkomkitem.Add("item_key", "14");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_更新料"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["更新料表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_更新料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_更新料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 仲介手数料
                hash_nkinkomkitem.Add("item_key", "15");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_仲介手数料"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["仲介手数料表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_仲介手数料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_仲介手数料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 雑費
                hash_nkinkomkitem.Add("item_key", "19");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_雑費"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_雑費");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_雑費No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 鍵交換代等
                hash_nkinkomkitem.Add("item_key", "20");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_鍵交換代等"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_鍵交換代等");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_鍵交換代等No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // その他一時金
                hash_nkinkomkitem.Add("item_key", "18");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_その他一時金"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_その他一時金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_その他一時金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // その他月額費用
                hash_nkinkomkitem.Add("item_key", "17");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_その他月額費用"]);
                hash_nkinkomkitem.Add("item_tani", "_円");
                hash_nkinkomkitem.Add("item_komk", "_その他費用");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_その他月額費用No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 駐車場敷金
                hash_nkinkomkitem.Add("item_key", "21");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_駐車場敷金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["駐車場敷金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_駐車場敷金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_駐車場敷金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 駐車場礼金
                hash_nkinkomkitem.Add("item_key", "22");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_駐車場礼金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["駐車場礼金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_駐車場礼金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_駐車場礼金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 結合
                string xml_total = "";
                string xml_nkintotal = "";
                for (int cntii = 1, loopTo = nkincnt; cntii <= loopTo; cntii++)
                    xml_nkintotal = xml_nkintotal + xml_nkinkomk[cntii];
                xml_total = xml_pre + xml_nkintotal + xml_post;

                rtn_str = xml_total;

                return rtn_str;

            }

            /// <summary>
            /// 入金項目の編集
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_xml"></param>
            /// <param name="hash_nkinkomkmst"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, string> Chg_NkinKomk(SafeDictionary<string, string> hash_xml, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                var rtn_hash = new SafeDictionary<string, string>();
                var tmp_hash = new SafeDictionary<string, string>();

                int tmp_nkincnt = 32;
                var tmp_nkinkomkname = new string[tmp_nkincnt + 1];
                var tmp_nkinkomkkbn = new string[tmp_nkincnt + 1];
                var tmp_fldname = new string[tmp_nkincnt + 1];
                // -----------------------
                // 入金項目取得
                // -----------------------
                foreach (var item in hash_xml)
                {

                    string midfldname = Conversions.ToString(item.Key);
                    string value = Conversions.ToString(item.Value);

                    switch (midfldname ?? "")
                    {
                        case "入金項目_賃料1":
                            {
                                tmp_nkinkomkname[1] = value;
                                tmp_fldname[1] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料1":
                            {
                                tmp_nkinkomkkbn[1] = value;
                                break;
                            }
                        case "入金項目_賃料2":
                            {
                                tmp_nkinkomkname[2] = value;
                                tmp_fldname[2] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料2":
                            {
                                tmp_nkinkomkkbn[2] = value;
                                break;
                            }
                        case "入金項目_賃料3":
                            {
                                tmp_nkinkomkname[3] = value;
                                tmp_fldname[3] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料3":
                            {
                                tmp_nkinkomkkbn[3] = value;
                                break;
                            }
                        case "入金項目_共益費1":
                            {
                                tmp_nkinkomkname[4] = value;
                                tmp_fldname[4] = midfldname;
                                break;
                            }
                        case "入金項目区分_共益費1":
                            {
                                tmp_nkinkomkkbn[4] = value;
                                break;
                            }
                        case "入金項目_共益費2":
                            {
                                tmp_nkinkomkname[5] = value;
                                tmp_fldname[5] = midfldname;
                                break;
                            }
                        case "入金項目区分_共益費2":
                            {
                                tmp_nkinkomkkbn[5] = value;
                                break;
                            }
                        case "入金項目_共益費3":
                            {
                                tmp_nkinkomkname[6] = value;
                                tmp_fldname[6] = midfldname;
                                break;
                            }
                        case "入金項目区分_共益費3":
                            {
                                tmp_nkinkomkkbn[6] = value;
                                break;
                            }
                        case "入金項目_管理費1":
                            {
                                tmp_nkinkomkname[7] = value;
                                tmp_fldname[7] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費1":
                            {
                                tmp_nkinkomkkbn[7] = value;
                                break;
                            }
                        case "入金項目_管理費2":
                            {
                                tmp_nkinkomkname[8] = value;
                                tmp_fldname[8] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費2":
                            {
                                tmp_nkinkomkkbn[8] = value;
                                break;
                            }
                        case "入金項目_管理費3":
                            {
                                tmp_nkinkomkname[9] = value;
                                tmp_fldname[9] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費3":
                            {
                                tmp_nkinkomkkbn[9] = value;
                                break;
                            }
                        case "入金項目_敷金":
                            {
                                tmp_nkinkomkname[10] = value;
                                tmp_fldname[10] = midfldname;
                                break;
                            }
                        case "入金項目区分_敷金":
                            {
                                tmp_nkinkomkkbn[10] = value;
                                break;
                            }
                        case "入金項目_礼金":
                            {
                                tmp_nkinkomkname[11] = value;
                                tmp_fldname[11] = midfldname;
                                break;
                            }
                        case "入金項目区分_礼金":
                            {
                                tmp_nkinkomkkbn[11] = value;
                                break;
                            }
                        case "入金項目_保証金":
                            {
                                tmp_nkinkomkname[12] = value;
                                tmp_fldname[12] = midfldname;
                                break;
                            }
                        case "入金項目区分_保証金":
                            {
                                tmp_nkinkomkkbn[12] = value;
                                break;
                            }
                        case "入金項目_保証金償却":
                            {
                                tmp_nkinkomkname[13] = value;
                                tmp_fldname[13] = midfldname;
                                break;
                            }
                        case "入金項目区分_保証金償却":
                            {
                                tmp_nkinkomkkbn[13] = value;
                                break;
                            }
                        case "入金項目_敷引金":
                            {
                                tmp_nkinkomkname[14] = value;
                                tmp_fldname[14] = midfldname;
                                break;
                            }
                        case "入金項目区分_敷引金":
                            {
                                tmp_nkinkomkkbn[14] = value;
                                break;
                            }
                        case "入金項目_更新料":
                            {
                                tmp_nkinkomkname[15] = value;
                                tmp_fldname[15] = midfldname;
                                break;
                            }
                        case "入金項目区分_更新料":
                            {
                                tmp_nkinkomkkbn[15] = value;
                                break;
                            }
                        case "入金項目_仲介手数料":
                            {
                                tmp_nkinkomkname[16] = value;
                                tmp_fldname[16] = midfldname;
                                break;
                            }
                        case "入金項目区分_仲介手数料":
                            {
                                tmp_nkinkomkkbn[16] = value;
                                break;
                            }
                        case "入金項目_鍵交換代等":
                            {
                                tmp_nkinkomkname[17] = value;
                                tmp_fldname[17] = midfldname;
                                break;
                            }
                        case "入金項目区分_鍵交換代等":
                            {
                                tmp_nkinkomkkbn[17] = value;
                                break;
                            }
                        case "入金項目_雑費1":
                            {
                                tmp_nkinkomkname[18] = value;
                                tmp_fldname[18] = midfldname;
                                break;
                            }
                        case "入金項目区分_雑費1":
                            {
                                tmp_nkinkomkkbn[18] = value;
                                break;
                            }
                        case "入金項目_雑費2":
                            {
                                tmp_nkinkomkname[19] = value;
                                tmp_fldname[19] = midfldname;
                                break;
                            }
                        case "入金項目区分_雑費2":
                            {
                                tmp_nkinkomkkbn[19] = value;
                                break;
                            }
                        case "入金項目_雑費3":
                            {
                                tmp_nkinkomkname[20] = value;
                                tmp_fldname[20] = midfldname;
                                break;
                            }
                        case "入金項目区分_雑費3":
                            {
                                tmp_nkinkomkkbn[20] = value;
                                break;
                            }
                        case "入金項目_駐車場敷金":
                            {
                                tmp_nkinkomkname[21] = value;
                                tmp_fldname[21] = midfldname;
                                break;
                            }
                        case "入金項目区分_駐車場敷金":
                            {
                                tmp_nkinkomkkbn[21] = value;
                                break;
                            }
                        case "入金項目_駐車場礼金":
                            {
                                tmp_nkinkomkname[22] = value;
                                tmp_fldname[22] = midfldname;
                                break;
                            }
                        case "入金項目区分_駐車場礼金":
                            {
                                tmp_nkinkomkkbn[22] = value;
                                break;
                            }
                        case "入金項目_その他月額費用1":
                            {
                                tmp_nkinkomkname[23] = value;
                                tmp_fldname[23] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他月額費用1":
                            {
                                tmp_nkinkomkkbn[23] = value;
                                break;
                            }
                        case "入金項目_その他月額費用2":
                            {
                                tmp_nkinkomkname[24] = value;
                                tmp_fldname[24] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他月額費用2":
                            {
                                tmp_nkinkomkkbn[24] = value;
                                break;
                            }
                        case "入金項目_その他月額費用3":
                            {
                                tmp_nkinkomkname[25] = value;
                                tmp_fldname[25] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他月額費用3":
                            {
                                tmp_nkinkomkkbn[25] = value;
                                break;
                            }
                        case "入金項目_その他月額費用4":
                            {
                                tmp_nkinkomkname[26] = value;
                                tmp_fldname[26] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他月額費用4":
                            {
                                tmp_nkinkomkkbn[26] = value;
                                break;
                            }
                        case "入金項目_その他月額費用5":
                            {
                                tmp_nkinkomkname[27] = value;
                                tmp_fldname[27] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他月額費用5":
                            {
                                tmp_nkinkomkkbn[27] = value;
                                break;
                            }
                        case "入金項目_その他一時金1":
                            {
                                tmp_nkinkomkname[28] = value;
                                tmp_fldname[28] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他一時金1":
                            {
                                tmp_nkinkomkkbn[28] = value;
                                break;
                            }
                        case "入金項目_その他一時金2":
                            {
                                tmp_nkinkomkname[29] = value;
                                tmp_fldname[29] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他一時金2":
                            {
                                tmp_nkinkomkkbn[29] = value;
                                break;
                            }
                        case "入金項目_その他一時金3":
                            {
                                tmp_nkinkomkname[30] = value;
                                tmp_fldname[30] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他一時金3":
                            {
                                tmp_nkinkomkkbn[30] = value;
                                break;
                            }
                        case "入金項目_その他一時金4":
                            {
                                tmp_nkinkomkname[31] = value;
                                tmp_fldname[31] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他一時金4":
                            {
                                tmp_nkinkomkkbn[31] = value;
                                break;
                            }
                        case "入金項目_その他一時金5":
                            {
                                tmp_nkinkomkname[32] = value;
                                tmp_fldname[32] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他一時金5":
                            {
                                tmp_nkinkomkkbn[32] = value;
                                break;
                            }

                        default:
                            {
                                tmp_hash.Add(midfldname, value);
                                break;
                            }
                    }

                }

                // -----------------------
                // 入金項目をXml移行用に成形
                // -----------------------
                var tmp_nkinno = new string[tmp_nkincnt + 1];
                var tmp_dispnkinname = new string[tmp_nkincnt + 1];
                for (int cntii = 1, loopTo = tmp_nkincnt; cntii <= loopTo; cntii++)
                {

                    string tmp_relnkinstr = tmp_nkinkomkname[cntii] + "-" + EtcMethod.Get_Nkinruiname(ref tmp_nkinkomkkbn[cntii]);
                    tmp_nkinno[cntii] = "";
                    tmp_dispnkinname[cntii] = "";

                    if (CommonModule.Hash_Rel_Nkinkomk.ContainsKey(tmp_relnkinstr))
                    {
                        tmp_nkinno[cntii] = Conversions.ToString(CommonModule.Hash_Rel_Nkinkomk[tmp_relnkinstr]);
                        tmp_dispnkinname[cntii] = Conversions.ToString(Operators.ConcatenateObject(tmp_nkinno[cntii] + ":", hash_nkinkomkmst[tmp_nkinno[cntii]]));
                    }

                }

                // -----------------------
                // 入金項目の再格納
                // -----------------------
                // 結合する項目
                Set_NkinkomkAddToHash(1, 3, tmp_nkinno, tmp_dispnkinname, "入金項目_賃料", ref tmp_hash);
                Set_NkinkomkAddToHash(4, 6, tmp_nkinno, tmp_dispnkinname, "入金項目_共益費", ref tmp_hash);
                Set_NkinkomkAddToHash(7, 9, tmp_nkinno, tmp_dispnkinname, "入金項目_管理費", ref tmp_hash);
                Set_NkinkomkAddToHash(18, 20, tmp_nkinno, tmp_dispnkinname, "入金項目_雑費", ref tmp_hash);
                Set_NkinkomkAddToHash(23, 27, tmp_nkinno, tmp_dispnkinname, "入金項目_その他月額費用", ref tmp_hash);
                Set_NkinkomkAddToHash(28, 32, tmp_nkinno, tmp_dispnkinname, "入金項目_その他一時金", ref tmp_hash);

                // 結合不要の項目
                for (int cntii = 10; cntii <= 22; cntii++)
                {
                    tmp_hash.Add(tmp_fldname[cntii], tmp_dispnkinname[cntii]);
                    tmp_hash.Add(tmp_fldname[cntii] + "No", tmp_nkinno[cntii]);
                }

                rtn_hash = tmp_hash;

                return rtn_hash;

            }

            /// <summary>
            /// 入金項目を結合して1つのデータにまとめる
            /// まとめたデータをハッシュテーブルへ格納する
            /// </summary>
            /// <param name="cntsta"></param>
            /// <param name="cntend"></param>
            /// <param name="tmp_nkinno"></param>
            /// <param name="tmp_dispnkinname"></param>
            /// <param name="hashkey"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_NkinkomkAddToHash(int cntsta, int cntend, string[] tmp_nkinno, string[] tmp_dispnkinname, string hashkey, ref SafeDictionary<string, string> hash)
            {

                string nkinno = "";
                string dispname = "";
                for (int cntii = cntsta, loopTo = cntend; cntii <= loopTo; cntii++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(tmp_nkinno[cntii], "", false)))
                    {
                        nkinno = Conversions.ToString(Operators.ConcatenateObject(nkinno + ",", tmp_nkinno[cntii]));
                        dispname = Conversions.ToString(Operators.ConcatenateObject(dispname + " / ", tmp_dispnkinname[cntii]));
                    }
                }
                if (!string.IsNullOrEmpty(nkinno))
                {
                    nkinno = nkinno.Remove(0, 1);
                }
                if (!string.IsNullOrEmpty(dispname))
                {
                    dispname = dispname.Remove(0, 3);
                }
                hash.Add(hashkey, dispname);
                hash.Add(hashkey + "No", nkinno);

            }

            /// <summary>
            /// 各入金項目毎にXml文字列作成
            /// </summary>
            /// <param name="hash_nkin"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Chg_NkinkomkToXml(SafeDictionary<string, string> hash_nkin)
            {

                string rtn_str = "";

                // 置換用/ハッシュテーブル値取得用文字列
                string item_key = "item_key";
                string item_disp = "item_disp";
                string item_tani = "item_tani";
                string item_komk = "item_komk";
                string item_nkinno = "item_nkinno";

                // Xmlのベースを作成
                var xml_nkinkomkbase = new System.Text.StringBuilder();
                xml_nkinkomkbase.Append("<a:KeyValueOfintPortalNkinKomokLjh4bohd>");
                xml_nkinkomkbase.Append("<a:Key>" + item_key + "</a:Key>");
                xml_nkinkomkbase.Append("<a:Value>");
                xml_nkinkomkbase.Append("<DisplayString>" + item_disp + "</DisplayString>");
                xml_nkinkomkbase.Append("<EnKagetu>" + item_tani + "</EnKagetu>");
                xml_nkinkomkbase.Append("<NkinKomok>" + item_komk + "</NkinKomok>");
                xml_nkinkomkbase.Append("<NkinNoList>");
                xml_nkinkomkbase.Append(item_nkinno);
                xml_nkinkomkbase.Append("</NkinNoList>");
                xml_nkinkomkbase.Append("</a:Value>");
                xml_nkinkomkbase.Append("</a:KeyValueOfintPortalNkinKomokLjh4bohd>");

                // 各要素で置換してい成形
                string xml_nkinkomk = xml_nkinkomkbase.ToString();
                xml_nkinkomk = xml_nkinkomk.Replace(item_key, Conversions.ToString(hash_nkin[item_key]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_disp, Conversions.ToString(hash_nkin[item_disp]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_tani, Conversions.ToString(hash_nkin[item_tani]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_komk, Conversions.ToString(hash_nkin[item_komk]));
                string tmp_nkinnogrp_tinryo = Conversions.ToString(hash_nkin[item_nkinno]);
                string[] tmp_nkinno_tinryo = tmp_nkinnogrp_tinryo.Split(',');
                string nkinno_tinryo = "";
                for (int cntii = 0, loopTo = Information.UBound(tmp_nkinno_tinryo); cntii <= loopTo; cntii++)
                {
                    if (!string.IsNullOrEmpty(tmp_nkinno_tinryo[cntii]))
                    {
                        nkinno_tinryo = nkinno_tinryo + "<a:int>" + tmp_nkinno_tinryo[cntii] + "</a:int>";
                    }
                }
                xml_nkinkomk = xml_nkinkomk.Replace(item_nkinno, nkinno_tinryo);

                rtn_str = xml_nkinkomk;

                return rtn_str;

            }

            /// <summary>
            /// その他項目に関するXml文字列の作成
            /// </summary>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_xmlitem"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Other(SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_xmlitem)
            {

                string rtn_str = "";
                int othercnt = 100;
                var xml_other = new string[othercnt + 1];
                int cntindex = 1;

                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<Password>", hash_cvitem["site_password"]), "</Password>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<SiteID>", hash_cvitem["site_id"]), "</SiteID>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<SiteSendUmu>", Interaction.IIf(Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(hash_cvitem["site_sendumu"], "1", false)), "true", "false")), "</SiteSendUmu>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = "<UseCommonSetting>" + "false" + "</UseCommonSetting>";
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BantiDispSetting>", hash_xmlitem["番地以降の表示"]), "</BantiDispSetting>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BkHyDispSetting>", hash_xmlitem["物件名・部屋Noの表示(一般ユーザー向け)"]), "</BkHyDispSetting>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BkHyDispSettingKaiin>", hash_xmlitem["物件名・部屋Noの表示(不動産会社向け)"]), "</BkHyDispSettingKaiin>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<KeisaiKakuninBi>", hash_xmlitem["掲載確認日"]), "</KeisaiKakuninBi>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<MapDispSetting>", hash_xmlitem["地図表示"]), "</MapDispSetting>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<TenpoCode>", hash_xmlitem["店舗ID"]), "</TenpoCode>"));
                cntindex = cntindex + 1;

                // 結合
                string tmp_str = "";
                for (int cntii = 1, loopTo = othercnt; cntii <= loopTo; cntii++)
                {
                    if (xml_other[cntii] is null)
                    {
                        break;
                    }
                    tmp_str = tmp_str + xml_other[cntii];
                }

                rtn_str = tmp_str;

                return rtn_str;

            }

            /// <summary>
            /// 店舗コードの桁数(4桁)チェック
            /// V7の店舗IDはathomeIDと店舗コード(下4桁)が結合された状態になっている
            /// 抽出時に下4桁を店舗ID、残りをathomeIDとして取得している
            /// </summary>
            /// <param name="tblname"></param>
            /// <param name="hash_xmlimte"></param>
            /// <param name="hash_log"></param>
            /// <remarks></remarks>
            public void Chk_Tenpocode(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_cvitem, ref SafeDictionary<string, string> hash_xmlitem)
            {

                string errstr = "";
                string log_key = tblname + "-" + "site_data";
                string athomeid = Conversions.ToString(hash_cvitem["site_id"]);
                string tenpocode = Conversions.ToString(hash_xmlitem["店舗ID"]);

                // ログ出力用の項目設定
                string syorikomok = "送信設定athome情報";
                string taisyodata = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("送信設定順 = ", hash_cvitem["setting_guid"]), "、"), "サイトNo = "), hash_cvitem["site_no"]));

                // 桁数チェック
                if (tenpocode.Length < 4)
                {

                    // 4桁へ成形
                    string tmp_tenpocode = tenpocode;
                    tmp_tenpocode = "0000" + tmp_tenpocode;
                    tmp_tenpocode = tmp_tenpocode.Substring(tmp_tenpocode.Length - 4);

                    // 再格納
                    hash_xmlitem["店舗ID"] = tmp_tenpocode;

                    // ログ出力
                    string syoriitem = "店舗ID";
                    string syorikekka = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE_STR;
                    string hubigein = CommonModule.LOG_HUBI_OUTOFRANGE_STR_SHORTAGE;
                    string taisyo = CommonModule.LOG_TAISYO_OUTOOFRANGE_STR;
                    string beforechgvalue = tenpocode;
                    string afterchgvalue = tmp_tenpocode;

                    // ログ挿入用へ加工
                    string logitem = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname);
                    // 挿入用クエリへ加工
                    string log_sql = LogSetting.Get_LogTblInsertQry(logitem, false);
                    // ログ挿入処理
                    int tmp_cnt = 0;
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, ref tmp_cnt);

                }

                // -----------------------------------------------
                // athomeIDが空の場合の対応
                // ※tenpocodeが4桁以下の場合は空文字になっている
                // ※念の為空文字チェックは行っておく
                // -----------------------------------------------
                if (string.IsNullOrEmpty(athomeid))
                {

                    // 任意の文字列を格納
                    string tmp_athomeid = "athomeid";

                    // 再格納
                    hash_cvitem["site_id"] = tmp_athomeid;

                    // ログ出力
                    string syoriitem = "athomeID";
                    string syorikekka = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE_STR;
                    string hubigein = CommonModule.LOG_HUBI_OUTOFRANGE_STR_SHORTAGE;
                    string taisyo = CommonModule.LOG_TAISYO_OUTOOFRANGE_STR;
                    string beforechgvalue = athomeid;
                    string afterchgvalue = tmp_athomeid;

                    // ログ挿入用へ加工
                    string logitem = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname);
                    // 挿入用クエリへ加工
                    string log_sql = LogSetting.Get_LogTblInsertQry(logitem, false);
                    // ログ挿入処理
                    int tmp_cnt = 0;
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, ref tmp_cnt);

                }

            }

        }

    }

    #endregion

    #region 送信設定SUUMO情報

    public class M_site_sendsetting_suumo_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_settingguid = new SafeDictionary<string, string>();                           // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.M_site_sendsetting_Model();      // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();

                // 革命10の入金項目マスタを取得
                var hash_nkinkomkmst = new SafeDictionary<string, string>();
                string tmp_sql = "SELECT nkin_no,nkin_name FROM m_nkin";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_nkinkomkmst);

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // 送信設定guid取得
                Set_SettingGuid(sqlcnnv10, ref hash_settingguid);

                // テーブル名/フィールド名セット
                string tblname = "m_site_sendsetting";
                string fldnamegrp = "sitesetting_guid,setting_guid,site_no,site_id,site_sendumu," + "site_data,site_password";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_settingguid, ref list_basekeydata);

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1;

                    // 作業用変数
                    var hash_cvxmlitem = new SafeDictionary<string, string>();

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "送信設定順":
                                {
                                    model_cvitem.Vari_Setting_guid = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }
                            case "ポータルサイトID":
                                {
                                    model_cvitem.Vari_Site_id = fldvalue;
                                    break;
                                }
                            case "サイト別送信有無":
                                {
                                    model_cvitem.Vari_Site_sendumu = fldvalue;
                                    break;
                                }
                            case "ポータルサイトパスワード":
                                {
                                    model_cvitem.Vari_Site_password = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(fldvalue), " ", fldvalue));
                                    break;
                                }

                            default:
                                {
                                    if (!string.IsNullOrEmpty(fldname))
                                    {
                                        hash_cvxmlitem.Add(fldname, fldvalue);
                                    }

                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Sitesetting_guid = Guid.NewGuid().ToString();

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 会社支店コード桁数チェック
                    Chk_Kaisyasitencode(sqlcnnv10, tblname, ref hash_cvitem);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["setting_guid"] = hash_settingguid[Strings.StrConv(Conversions.ToString(hash_cvitem["setting_guid"]), VbStrConv.Narrow)];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // xmlの移行
                            Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem, hash_nkinkomkmst);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // ポータルサイトパスワード一括更新 (半角スペース " " を設定していた分を空文字へ変換)
                int tmpcnt = 0;
                string passwordupdateqry = " UPDATE " + tblname + " SET site_password = '' WHERE site_no = 40 AND site_password = ' ' ";
                DBExec.Exec_NonQuery(sqlcnnv10, passwordupdateqry, ref tmpcnt);

                // 共通設定情報へ個別設定No1を反映
                // 20160908 個別処理を共通処理へコピーする処理の制御修正 -chg sta
                // Dim tmp_sortstr As String = ""  '使用しない
                // Dim kobetuinfotocommonqry As String = Me.Get_UseQry(tmp_sortstr)
                // DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, tmpcnt)
                if (cvrowcnt != 0)
                {
                    string tmp_sortstr = "";  // 使用しない
                    string kobetuinfotocommonqry = Get_UseQry(ref tmp_sortstr);
                    DBExec.Exec_NonQuery(sqlcnnv10, kobetuinfotocommonqry, ref tmpcnt);
                }
                // 20160908 個別処理を共通処理へコピーする処理の制御修正 -chg end

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 共通設定情報更新クエリの作成
            /// </summary>
            /// <param name="sortstr"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " UPDATE m_site_sendsetting SET ";
                tmp_sql = tmp_sql + " 	 site_id = (SELECT site_id FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 40) ";
                tmp_sql = tmp_sql + " 	,site_sendumu = (SELECT site_sendumu FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 40) ";
                tmp_sql = tmp_sql + " 	,site_data = (SELECT site_data FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 40) ";
                tmp_sql = tmp_sql + " 	,site_password = (SELECT site_password FROM m_site_sendsetting WHERE setting_guid =(SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 1) AND site_no = 40) ";
                tmp_sql = tmp_sql + " WHERE setting_guid = ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0 ";
                tmp_sql = tmp_sql + " ) ";
                tmp_sql = tmp_sql + " AND site_no = 40 ";

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) ";
                tmp_sql = tmp_sql + " FROM hydata_kokoku AS HYKOKOKU ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 送信設定guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_SettingGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = " SELECT setting_sortorder,setting_guid FROM m_sendsetting ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 送信設定のxml項目の移行
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_cvxmlitem"></param>
            /// <remarks></remarks>
            public void Set_XmlItem(SqlConnection sqlcnnv10, SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_xmlitem, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                string sitesettingguid = Conversions.ToString(hash_cvitem["sitesetting_guid"]);   // 更新対象選択用

                // 「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                string xmlupdateqry_xmlpre = "";
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre + "<SuumoSendSettingModel xmlns:i=" + "\"" + "http://www.w3.org/2001/XMLSchema-instance" + "\"" + ">";
                xmlupdateqry_xmlpre = xmlupdateqry_xmlpre + "<__identity xmlns=" + "\"" + "http://schemas.datacontract.org/2004/07/System" + "\"" + " i:nil=" + "\"" + "true" + "\"" + " />";
                string xmlupdateqry_xmlpost = "";
                xmlupdateqry_xmlpost = xmlupdateqry_xmlpost + "</SuumoSendSettingModel>";

                // グループ毎にXml文字列を作成
                // 画像
                string tmp_xml_gazo = Get_XmlStr_Gazo(hash_xmlitem);
                // 入金項目
                string tmp_xml_nkin = Get_XmlStr_Nkinkomk(hash_xmlitem, hash_nkinkomkmst);
                // その他
                string tmp_xml_other = Get_XmlStr_Other(hash_cvitem, hash_xmlitem);

                // Xml文字列の結合
                string xmlupdateqry_main = xmlupdateqry_xmlpre + tmp_xml_gazo + tmp_xml_nkin + tmp_xml_other + xmlupdateqry_xmlpost;

                // 更新用クエリの作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " DECLARE @tmp_xml XML ";
                tmp_sql = tmp_sql + " SET @tmp_xml = CAST('" + xmlupdateqry_main + "' AS XML) ";
                tmp_sql = tmp_sql + " UPDATE m_site_sendsetting SET site_data = @tmp_xml ";
                tmp_sql = tmp_sql + " WHERE sitesetting_guid = '" + sitesettingguid + "' ";

                // 更新
                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

            /// <summary>
            /// 画像に関するXml文字列の作成
            /// </summary>
            /// <param name="hash"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Gazo(SafeDictionary<string, string> hash)
            {

                string rtn_str = "";
                string xml_pre = "<GazoSettingList xmlns:a=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.FK8Rendo.Core.Model" + "\"" + ">";
                string xml_post = "</GazoSettingList>";

                int gazocnt = 17;
                var tmp_gazono = new string[gazocnt + 1];
                var tmp_gazosyuV7 = new string[gazocnt + 1];
                var tmp_gazonoV7 = new string[gazocnt + 1];
                var tmp_gazosyurendo = new string[gazocnt + 1];

                // 画像に関する情報を配列へ格納
                foreach (var item in hash)
                {

                    string fldname = Conversions.ToString(item.Key);
                    string value = Conversions.ToString(item.Value);

                    switch (true)
                    {
                        case object _ when (fldname.Replace("画像No", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("画像No", ""));
                                tmp_gazono[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("革命側の画像種別", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("革命側の画像種別", ""));
                                tmp_gazosyuV7[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("革命側の画像タイトルNo", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("革命側の画像タイトルNo", ""));
                                tmp_gazonoV7[gazono] = value;
                                break;
                            }
                        case object _ when (fldname.Replace("連動側の画像種別", "") ?? "") != (fldname ?? ""):
                            {
                                int gazono = int.Parse(fldname.Replace("連動側の画像種別", ""));
                                tmp_gazosyurendo[gazono] = value;
                                break;
                            }
                    }

                }

                // 文字列へ成形
                var xml_gazoitem = new System.Text.StringBuilder();
                for (int cntii = 1, loopTo = gazocnt; cntii <= loopTo; cntii++)
                {
                    if (tmp_gazosyurendo[cntii] != "0" & tmp_gazosyuV7[cntii] != "0" & tmp_gazonoV7[cntii] != "0")
                    {
                        xml_gazoitem.Append("<a:PortalGazoSettingModel>");
                        xml_gazoitem.Append("<a:FK8LocationKbn>" + tmp_gazosyuV7[cntii] + "</a:FK8LocationKbn>");
                        xml_gazoitem.Append("<a:FK8LocationNo>" + tmp_gazonoV7[cntii] + "</a:FK8LocationNo>");
                        xml_gazoitem.Append("<a:No>" + tmp_gazono[cntii] + "</a:No>");
                        xml_gazoitem.Append("<a:PortalGazoSyubetu>" + tmp_gazosyurendo[cntii] + "</a:PortalGazoSyubetu>");
                        xml_gazoitem.Append("</a:PortalGazoSettingModel>");
                    }
                }

                // 結合
                string xml_gazo = xml_pre + xml_gazoitem.ToString() + xml_post;
                rtn_str = xml_gazo;

                return rtn_str;

            }

            /// <summary>
            /// 入金項目に関するXml文字列の作成
            /// </summary>
            /// <param name="hash"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Nkinkomk(SafeDictionary<string, string> hash_xmlitem, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                string rtn_str = "";
                string xml_pre = "<NkinKomokDict xmlns:a=" + "\"" + "http://schemas.microsoft.com/2003/10/Serialization/Arrays" + "\"" + ">";
                string xml_post = "</NkinKomokDict>";
                int nkincnt = 10;
                int cntindex = 1;
                var xml_nkinkomk = new string[nkincnt + 1];
                var hash_nkinkomkitem = new SafeDictionary<string, string>();

                // 入金項目を移行用に成形する
                var hash_cvxmlitem = new SafeDictionary<string, string>();
                hash_cvxmlitem = Chg_NkinKomk(hash_xmlitem, hash_nkinkomkmst);

                // 賃料
                hash_nkinkomkitem.Add("item_key", "1");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_賃料"]);
                hash_nkinkomkitem.Add("item_tani", "_万円");
                hash_nkinkomkitem.Add("item_komk", "_賃料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_賃料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 管理費
                hash_nkinkomkitem.Add("item_key", "4");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_管理費"]);
                hash_nkinkomkitem.Add("item_tani", "_万円");
                hash_nkinkomkitem.Add("item_komk", "_管理費");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_管理費No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 礼金
                hash_nkinkomkitem.Add("item_key", "5");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_礼金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["礼金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_礼金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_礼金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 敷金
                hash_nkinkomkitem.Add("item_key", "6");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_敷金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["敷金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_敷金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_敷金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 償却金
                hash_nkinkomkitem.Add("item_key", "9");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_償却金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["償却金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_償却金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_償却金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 保証金
                hash_nkinkomkitem.Add("item_key", "7");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_保証金"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["保証金表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_保証金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_保証金No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 敷引
                hash_nkinkomkitem.Add("item_key", "12");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_敷引"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["敷引表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_敷引金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_敷引No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 仲介手数料
                hash_nkinkomkitem.Add("item_key", "15");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_仲介手数料"]);
                hash_nkinkomkitem.Add("item_tani", hash_cvxmlitem["仲介手数料表示単位"]);
                hash_nkinkomkitem.Add("item_komk", "_仲介手数料");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_仲介手数料No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // ほか初期費用
                hash_nkinkomkitem.Add("item_key", "18");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_ほか初期費用"]);
                hash_nkinkomkitem.Add("item_tani", "_万円");
                hash_nkinkomkitem.Add("item_komk", "_その他一時金");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_ほか初期費用No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // その他諸費用
                hash_nkinkomkitem.Add("item_key", "17");
                hash_nkinkomkitem.Add("item_disp", hash_cvxmlitem["入金項目_その他諸費用"]);
                hash_nkinkomkitem.Add("item_tani", "_万円");
                hash_nkinkomkitem.Add("item_komk", "_その他費用");
                hash_nkinkomkitem.Add("item_nkinno", hash_cvxmlitem["入金項目_その他諸費用No"]);
                xml_nkinkomk[cntindex] = Chg_NkinkomkToXml(hash_nkinkomkitem);
                cntindex = cntindex + 1;
                hash_nkinkomkitem.Clear();

                // 結合
                string xml_total = "";
                string xml_nkintotal = "";
                for (int cntii = 1, loopTo = nkincnt; cntii <= loopTo; cntii++)
                    xml_nkintotal = xml_nkintotal + xml_nkinkomk[cntii];
                xml_total = xml_pre + xml_nkintotal + xml_post;

                rtn_str = xml_total;

                return rtn_str;

            }

            /// <summary>
            /// 入金項目の編集
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_xml"></param>
            /// <param name="hash_nkinkomkmst"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, string> Chg_NkinKomk(SafeDictionary<string, string> hash_xml, SafeDictionary<string, string> hash_nkinkomkmst)
            {

                var rtn_hash = new SafeDictionary<string, string>();
                var tmp_hash = new SafeDictionary<string, string>();

                int tmp_nkincnt = 14;
                var tmp_nkinkomkname = new string[tmp_nkincnt + 1];
                var tmp_nkinkomkkbn = new string[tmp_nkincnt + 1];
                var tmp_fldname = new string[tmp_nkincnt + 1];
                // -----------------------
                // 入金項目取得
                // -----------------------
                foreach (var item in hash_xml)
                {

                    string midfldname = Conversions.ToString(item.Key);
                    string value = Conversions.ToString(item.Value);

                    switch (midfldname ?? "")
                    {
                        case "入金項目_賃料1":
                            {
                                tmp_nkinkomkname[1] = value;
                                tmp_fldname[1] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料1":
                            {
                                tmp_nkinkomkkbn[1] = value;
                                break;
                            }
                        case "入金項目_賃料2":
                            {
                                tmp_nkinkomkname[2] = value;
                                tmp_fldname[2] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料2":
                            {
                                tmp_nkinkomkkbn[2] = value;
                                break;
                            }
                        case "入金項目_賃料3":
                            {
                                tmp_nkinkomkname[3] = value;
                                tmp_fldname[3] = midfldname;
                                break;
                            }
                        case "入金項目区分_賃料3":
                            {
                                tmp_nkinkomkkbn[3] = value;
                                break;
                            }
                        case "入金項目_管理費1":
                            {
                                tmp_nkinkomkname[4] = value;
                                tmp_fldname[4] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費1":
                            {
                                tmp_nkinkomkkbn[4] = value;
                                break;
                            }
                        case "入金項目_管理費2":
                            {
                                tmp_nkinkomkname[5] = value;
                                tmp_fldname[5] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費2":
                            {
                                tmp_nkinkomkkbn[5] = value;
                                break;
                            }
                        case "入金項目_管理費3":
                            {
                                tmp_nkinkomkname[6] = value;
                                tmp_fldname[6] = midfldname;
                                break;
                            }
                        case "入金項目区分_管理費3":
                            {
                                tmp_nkinkomkkbn[6] = value;
                                break;
                            }
                        case "入金項目_敷金":
                            {
                                tmp_nkinkomkname[7] = value;
                                tmp_fldname[7] = midfldname;
                                break;
                            }
                        case "入金項目区分_敷金":
                            {
                                tmp_nkinkomkkbn[7] = value;
                                break;
                            }
                        case "入金項目_礼金":
                            {
                                tmp_nkinkomkname[8] = value;
                                tmp_fldname[8] = midfldname;
                                break;
                            }
                        case "入金項目区分_礼金":
                            {
                                tmp_nkinkomkkbn[8] = value;
                                break;
                            }
                        case "入金項目_保証金":
                            {
                                tmp_nkinkomkname[9] = value;
                                tmp_fldname[9] = midfldname;
                                break;
                            }
                        case "入金項目区分_保証金":
                            {
                                tmp_nkinkomkkbn[9] = value;
                                break;
                            }
                        case "入金項目_償却金":
                            {
                                tmp_nkinkomkname[10] = value;
                                tmp_fldname[10] = midfldname;
                                break;
                            }
                        case "入金項目区分_償却金":
                            {
                                tmp_nkinkomkkbn[10] = value;
                                break;
                            }
                        case "入金項目_敷引":
                            {
                                tmp_nkinkomkname[11] = value;
                                tmp_fldname[11] = midfldname;
                                break;
                            }
                        case "入金項目区分_敷引":
                            {
                                tmp_nkinkomkkbn[11] = value;
                                break;
                            }
                        case "入金項目_仲介手数料":
                            {
                                tmp_nkinkomkname[12] = value;
                                tmp_fldname[12] = midfldname;
                                break;
                            }
                        case "入金項目区分_仲介手数料":
                            {
                                tmp_nkinkomkkbn[12] = value;
                                break;
                            }
                        case "入金項目_その他諸費用":
                            {
                                tmp_nkinkomkname[13] = value;
                                tmp_fldname[13] = midfldname;
                                break;
                            }
                        case "入金項目区分_その他諸費用":
                            {
                                tmp_nkinkomkkbn[13] = value;
                                break;
                            }
                        case "入金項目_ほか初期費用":
                            {
                                tmp_nkinkomkname[14] = value;
                                tmp_fldname[14] = midfldname;
                                break;
                            }
                        case "入金項目区分_ほか初期費用":
                            {
                                tmp_nkinkomkkbn[14] = value;
                                break;
                            }

                        default:
                            {
                                tmp_hash.Add(midfldname, value);
                                break;
                            }
                    }

                }

                // -----------------------
                // 入金項目をXml移行用に成形
                // -----------------------
                var tmp_nkinno = new string[tmp_nkincnt + 1];
                var tmp_dispnkinname = new string[tmp_nkincnt + 1];
                for (int cntii = 1, loopTo = tmp_nkincnt; cntii <= loopTo; cntii++)
                {

                    string tmp_relnkinstr = tmp_nkinkomkname[cntii] + "-" + EtcMethod.Get_Nkinruiname(ref tmp_nkinkomkkbn[cntii]);
                    tmp_nkinno[cntii] = "";
                    tmp_dispnkinname[cntii] = "";

                    if (CommonModule.Hash_Rel_Nkinkomk.ContainsKey(tmp_relnkinstr))
                    {
                        tmp_nkinno[cntii] = Conversions.ToString(CommonModule.Hash_Rel_Nkinkomk[tmp_relnkinstr]);
                        tmp_dispnkinname[cntii] = Conversions.ToString(Operators.ConcatenateObject(tmp_nkinno[cntii] + ":", hash_nkinkomkmst[tmp_nkinno[cntii]]));
                    }

                }

                // -----------------------
                // 入金項目の再格納
                // -----------------------
                // 結合する項目
                Set_NkinkomkAddToHash(1, 3, tmp_nkinno, tmp_dispnkinname, "入金項目_賃料", ref tmp_hash);
                Set_NkinkomkAddToHash(4, 6, tmp_nkinno, tmp_dispnkinname, "入金項目_管理費", ref tmp_hash);

                // 結合不要の項目
                for (int cntii = 7; cntii <= 14; cntii++)
                {
                    tmp_hash.Add(tmp_fldname[cntii], tmp_dispnkinname[cntii]);
                    tmp_hash.Add(tmp_fldname[cntii] + "No", tmp_nkinno[cntii]);
                }

                rtn_hash = tmp_hash;

                return rtn_hash;

            }

            /// <summary>
            /// 入金項目を結合して1つのデータにまとめる
            /// まとめたデータをハッシュテーブルへ格納する
            /// </summary>
            /// <param name="cntsta"></param>
            /// <param name="cntend"></param>
            /// <param name="tmp_nkinno"></param>
            /// <param name="tmp_dispnkinname"></param>
            /// <param name="hashkey"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_NkinkomkAddToHash(int cntsta, int cntend, string[] tmp_nkinno, string[] tmp_dispnkinname, string hashkey, ref SafeDictionary<string, string> hash)
            {

                string nkinno = "";
                string dispname = "";
                for (int cntii = cntsta, loopTo = cntend; cntii <= loopTo; cntii++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(tmp_nkinno[cntii], "", false)))
                    {
                        nkinno = Conversions.ToString(Operators.ConcatenateObject(nkinno + ",", tmp_nkinno[cntii]));
                        dispname = Conversions.ToString(Operators.ConcatenateObject(dispname + " / ", tmp_dispnkinname[cntii]));
                    }
                }
                if (!string.IsNullOrEmpty(nkinno))
                {
                    nkinno = nkinno.Remove(0, 1);
                }
                if (!string.IsNullOrEmpty(dispname))
                {
                    dispname = dispname.Remove(0, 3);
                }
                hash.Add(hashkey, dispname);
                hash.Add(hashkey + "No", nkinno);

            }

            /// <summary>
            /// 各入金項目毎にXml文字列作成
            /// </summary>
            /// <param name="hash_nkin"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Chg_NkinkomkToXml(SafeDictionary<string, string> hash_nkin)
            {

                string rtn_str = "";

                // 置換用/ハッシュテーブル値取得用文字列
                string item_key = "item_key";
                string item_disp = "item_disp";
                string item_tani = "item_tani";
                string item_komk = "item_komk";
                string item_nkinno = "item_nkinno";

                // Xmlのベースを作成
                var xml_nkinkomkbase = new System.Text.StringBuilder();
                xml_nkinkomkbase.Append("<a:KeyValueOfintPortalNkinKomokLjh4bohd>");
                xml_nkinkomkbase.Append("<a:Key>" + item_key + "</a:Key>");
                xml_nkinkomkbase.Append("<a:Value>");
                xml_nkinkomkbase.Append("<DisplayString>" + item_disp + "</DisplayString>");
                xml_nkinkomkbase.Append("<EnKagetu>" + item_tani + "</EnKagetu>");
                xml_nkinkomkbase.Append("<NkinKomok>" + item_komk + "</NkinKomok>");
                xml_nkinkomkbase.Append("<NkinNoList>");
                xml_nkinkomkbase.Append(item_nkinno);
                xml_nkinkomkbase.Append("</NkinNoList>");
                xml_nkinkomkbase.Append("</a:Value>");
                xml_nkinkomkbase.Append("</a:KeyValueOfintPortalNkinKomokLjh4bohd>");

                // 各要素で置換してい成形
                string xml_nkinkomk = xml_nkinkomkbase.ToString();
                xml_nkinkomk = xml_nkinkomk.Replace(item_key, Conversions.ToString(hash_nkin[item_key]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_disp, Conversions.ToString(hash_nkin[item_disp]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_tani, Conversions.ToString(hash_nkin[item_tani]));
                xml_nkinkomk = xml_nkinkomk.Replace(item_komk, Conversions.ToString(hash_nkin[item_komk]));
                string tmp_nkinnogrp_tinryo = Conversions.ToString(hash_nkin[item_nkinno]);
                string[] tmp_nkinno_tinryo = tmp_nkinnogrp_tinryo.Split(',');
                string nkinno_tinryo = "";
                for (int cntii = 0, loopTo = Information.UBound(tmp_nkinno_tinryo); cntii <= loopTo; cntii++)
                {
                    if (!string.IsNullOrEmpty(tmp_nkinno_tinryo[cntii]))
                    {
                        nkinno_tinryo = nkinno_tinryo + "<a:int>" + tmp_nkinno_tinryo[cntii] + "</a:int>";
                    }
                }
                xml_nkinkomk = xml_nkinkomk.Replace(item_nkinno, nkinno_tinryo);

                rtn_str = xml_nkinkomk;

                return rtn_str;

            }

            /// <summary>
            /// その他項目に関するXml文字列の作成
            /// </summary>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_xmlitem"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlStr_Other(SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_xmlitem)
            {

                string rtn_str = "";
                int othercnt = 100;
                var xml_other = new string[othercnt + 1];
                int cntindex = 1;

                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<Password>", hash_cvitem["site_password"]), "</Password>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<SiteID>", hash_cvitem["site_id"]), "</SiteID>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<SiteSendUmu>", Interaction.IIf(Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(hash_cvitem["site_sendumu"], "1", false)), "true", "false")), "</SiteSendUmu>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = "<UseCommonSetting>" + "false" + "</UseCommonSetting>";
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BantiDispSetting>", hash_xmlitem["番地以降の表示(会社間図面および会社間流通向け)"]), "</BantiDispSetting>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BkHyDispSetting>", hash_xmlitem["物件名・部屋Noの表示(一般ユーザー向け)"]), "</BkHyDispSetting>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<BkHyDispSettingKaisya>", hash_xmlitem["物件名・部屋Noの表示(会社間図面および会社間流通向け)"]), "</BkHyDispSettingKaisya>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<FtpPassword>", hash_xmlitem["FTPパスワード"]), "</FtpPassword>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<FtpUserId>", hash_xmlitem["FTPユーザーID"]), "</FtpUserId>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<IsKaiyakuYoteiAsKusitu>", hash_xmlitem["空室状況設定"]), "</IsKaiyakuYoteiAsKusitu>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<KeisaiKakuninBi>", hash_xmlitem["元付確認日(先物物件のみ)"]), "</KeisaiKakuninBi>"));
                cntindex = cntindex + 1;
                xml_other[cntindex] = "<KusituJokyoSetting>" + "false" + "</KusituJokyoSetting>";
                cntindex = cntindex + 1;   // 要確認
                xml_other[cntindex] = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("<MapDispSetting>", hash_xmlitem["地図表示"]), "</MapDispSetting>"));
                cntindex = cntindex + 1;

                // 結合
                string tmp_str = "";
                for (int cntii = 1, loopTo = othercnt; cntii <= loopTo; cntii++)
                {
                    if (xml_other[cntii] is null)
                    {
                        break;
                    }
                    tmp_str = tmp_str + xml_other[cntii];
                }

                rtn_str = tmp_str;

                return rtn_str;

            }

            /// <summary>
            /// 会社支店コードの桁数(9桁)チェック
            /// </summary>
            /// <param name="tblname"></param>
            /// <param name="hash_xmlimte"></param>
            /// <param name="hash_log"></param>
            /// <remarks></remarks>
            public void Chk_Kaisyasitencode(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_cvitem)
            {

                string errstr = "";
                string log_key = tblname + "-" + "site_data";
                string kaisyasitencode = Conversions.ToString(hash_cvitem["site_id"]);

                // 桁数チェック(V7画面上で最大9桁の制御が設けられているため9桁に満たないかチェックする)
                if (kaisyasitencode.Length < 9)
                {

                    // 9桁へ成形
                    string tmp_kaisyasitencode = kaisyasitencode;
                    tmp_kaisyasitencode = "000000000" + tmp_kaisyasitencode;
                    tmp_kaisyasitencode = tmp_kaisyasitencode.Substring(tmp_kaisyasitencode.Length - 9);

                    // 再格納
                    hash_cvitem["site_id"] = tmp_kaisyasitencode;

                    // ログ出力(※xmlデータのログは直接出力する)
                    string syorikomok = "送信設定SUUMO情報";
                    string syoriitem = "会社支店コード";
                    string syorikekka = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE_STR;
                    string taisyodata = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("送信設定順 = ", hash_cvitem["setting_guid"]), "、"), "サイトNo = "), hash_cvitem["site_no"]));
                    string hubigein = CommonModule.LOG_HUBI_OUTOFRANGE_STR_SHORTAGE;
                    string taisyo = CommonModule.LOG_TAISYO_OUTOOFRANGE_STR;
                    string beforechgvalue = kaisyasitencode;
                    string afterchgvalue = tmp_kaisyasitencode;

                    // ログ挿入用へ加工
                    string logitem = LogSetting.Set_LogValue(42, syorikomok, syoriitem, syorikekka, taisyodata, hubigein, taisyo, beforechgvalue, afterchgvalue, tblname);
                    // 挿入用クエリへ加工
                    string log_sql = LogSetting.Get_LogTblInsertQry(logitem, false);
                    // ログ挿入処理
                    int tmp_cnt = 0;
                    DBExec.Exec_NonQuery(sqlcnnv10, log_sql, ref tmp_cnt);

                }

            }

        }

    }

    #endregion

    #region 広告補足自社web情報

    public class Hydata_kokoku_jisyaweb_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_hyguid = new SafeDictionary<string, string>();                                // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_kokoku_Model();           // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列

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

                // 部屋guid取得
                Set_HyGuid(sqlcnnv10, ref hash_hyguid);

                // テーブル名/フィールド名セット
                string tblname = "hydata_kokoku";
                string fldnamegrp = "hy_guid,site_no,item";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string tmp_keysub2 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub2]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;


                    // 作業用変数
                    string tmp_bkno = "";
                    string tmp_hyno = "";
                    var hash_cvxmlitem = new SafeDictionary<string, string>();

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "物件No":
                                {
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }

                            default:
                                {
                                    if (!string.IsNullOrEmpty(fldname))
                                    {
                                        hash_cvxmlitem.Add(fldname, fldvalue);
                                    }

                                    break;
                                }
                        }

                    }

                    // 部屋guid取得用
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["hy_guid"] = hash_hyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["hy_guid"]), VbStrConv.Narrow)];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // xmlの移行
                            Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) ";
                tmp_sql = tmp_sql + " FROM hydata_kokoku AS HYKOKOKU ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

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
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql = tmp_sql + " 	,HY.hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 広告補足のxml項目の移行
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_cvxmlitem"></param>
            /// <remarks></remarks>
            public void Set_XmlItem(SqlConnection sqlcnnv10, SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_cvxmlitem)
            {

                string hyguid = Conversions.ToString(hash_cvitem["hy_guid"]);   // 更新対象選択用
                string siteno = Conversions.ToString(hash_cvitem["site_no"]);   // 更新対象選択用
                var hash_xmlindex = new SafeDictionary<string, int>();              // xml要素の順番決定用

                // xml要素の順番設定
                hash_xmlindex = Get_Hash_XmlIndex();

                // 「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                string xmlupdateqry_xmlpre = "<WmpKokokuModel xmlns:i=" + "\"" + "http://www.w3.org/2001/XMLSchema-instance" + "\"" + ">";
                string xmlupdateqry_xmlpost = "</WmpKokokuModel>";
                string xmlitem = Get_XmlValue(hash_cvxmlitem, hash_xmlindex);
                string xmlupdateqry_main = xmlupdateqry_xmlpre + xmlitem + xmlupdateqry_xmlpost;

                // 更新用クエリの作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " DECLARE @tmp_xml XML ";
                tmp_sql = tmp_sql + " SET @tmp_xml = CAST('" + xmlupdateqry_main + "' AS XML) ";
                tmp_sql = tmp_sql + " UPDATE hydata_kokoku SET item = @tmp_xml ";
                tmp_sql = tmp_sql + " WHERE hy_guid = '" + hyguid + "' ";
                tmp_sql = tmp_sql + " AND site_no = " + siteno;

                // 更新
                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

            /// <summary>
            /// xml要素の順番決定
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, int> Get_Hash_XmlIndex()
            {

                var rtn_hash = new SafeDictionary<string, int>();

                rtn_hash.Add("AdditionalSalesPoint", 1);
                rtn_hash.Add("Biko", 2);
                rtn_hash.Add("CelingHeight", 3);
                rtn_hash.Add("ChinryoNegotiable", 4);
                rtn_hash.Add("RaisedFloor", 5);
                rtn_hash.Add("TantoComment", 6);
                rtn_hash.Add("Tokki", 7);

                return rtn_hash;

            }

            /// <summary>
            /// xml移行用の文字列の生成
            /// </summary>
            /// <param name="hash_cvxmlitem"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlValue(SafeDictionary<string, string> hash_cvxmlitem, SafeDictionary<string, int> hash_xmlindex)
            {

                var rtn_str = new System.Text.StringBuilder();
                var tmp_slist = new SortedList<int, string>();

                foreach (var midvalue in hash_cvxmlitem)
                {

                    string midfldname = Conversions.ToString(midvalue.Key);
                    string cvvalue = Conversions.ToString(midvalue.Value);
                    string tmp_xmlelement = "";

                    switch (midfldname ?? "")
                    {
                        case "セールスポイント補足":
                            {
                                tmp_xmlelement = "AdditionalSalesPoint";
                                break;
                            }
                        case "担当者コメント":
                            {
                                tmp_xmlelement = "TantoComment";
                                break;
                            }
                        case "特記事項":
                            {
                                tmp_xmlelement = "Tokki";
                                break;
                            }
                        case "広告備考":
                            {
                                tmp_xmlelement = "Biko";
                                break;
                            }
                        case "天井高":
                            {
                                tmp_xmlelement = "CelingHeight";
                                break;
                            }
                        case "OAフロアの高さ":
                            {
                                tmp_xmlelement = "RaisedFloor";
                                break;
                            }
                        case "賃料応相談":
                            {
                                tmp_xmlelement = "ChinryoNegotiable";
                                break;
                            }
                    }

                    string xmlelement = "<" + tmp_xmlelement + ">" + cvvalue + "</" + tmp_xmlelement + ">";
                    int keyno = hash_xmlindex[tmp_xmlelement];
                    tmp_slist.Add(keyno, xmlelement);

                }

                foreach (var item in tmp_slist)
                {
                    string xmlvalue = item.Value;
                    rtn_str.Append(xmlvalue);
                }

                return rtn_str.ToString();

            }

        }

    }

    #endregion

    #region 広告補足HOMES情報

    public class Hydata_kokoku_homes_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_hyguid = new SafeDictionary<string, string>();                                // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_kokoku_Model();           // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列

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

                // 部屋guid取得
                Set_HyGuid(sqlcnnv10, ref hash_hyguid);

                // テーブル名/フィールド名セット
                string tblname = "hydata_kokoku";
                string fldnamegrp = "hy_guid,site_no,item";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string tmp_keysub2 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub2]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;


                    // 作業用変数
                    string tmp_bkno = "";
                    string tmp_hyno = "";
                    var hash_cvxmlitem = new SafeDictionary<string, string>();

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "物件No":
                                {
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }

                            default:
                                {
                                    if (!string.IsNullOrEmpty(fldname))
                                    {
                                        hash_cvxmlitem.Add(fldname, fldvalue);
                                    }

                                    break;
                                }
                        }

                    }

                    // 部屋guid取得用
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["hy_guid"] = hash_hyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["hy_guid"]), VbStrConv.Narrow)];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // xmlの移行
                            Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) ";
                tmp_sql = tmp_sql + " FROM hydata_kokoku AS HYKOKOKU ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

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
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql = tmp_sql + " 	,HY.hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 広告補足のxml項目の移行
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_cvxmlitem"></param>
            /// <remarks></remarks>
            public void Set_XmlItem(SqlConnection sqlcnnv10, SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_cvxmlitem)
            {

                string hyguid = Conversions.ToString(hash_cvitem["hy_guid"]);   // 更新対象選択用
                string siteno = Conversions.ToString(hash_cvitem["site_no"]);   // 更新対象選択用
                var hash_xmlindex = new SafeDictionary<string, int>();              // xml要素の順番決定用

                // xml要素の順番設定
                hash_xmlindex = Get_Hash_XmlIndex();

                // 「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                string xmlupdateqry_xmlpre = "<HomesKokokuModel xmlns:i=" + "\"" + "http://www.w3.org/2001/XMLSchema-instance" + "\"" + ">";
                string xmlupdateqry_xmlpost = "</HomesKokokuModel>";
                string xmlitem = Get_XmlValue(hash_cvxmlitem, hash_xmlindex);
                string xmlupdateqry_main = xmlupdateqry_xmlpre + xmlitem + xmlupdateqry_xmlpost;

                // 更新用クエリの作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " DECLARE @tmp_xml XML ";
                tmp_sql = tmp_sql + " SET @tmp_xml = CAST('" + xmlupdateqry_main + "' AS XML) ";
                tmp_sql = tmp_sql + " UPDATE hydata_kokoku SET item = @tmp_xml ";
                tmp_sql = tmp_sql + " WHERE hy_guid = '" + hyguid + "' ";
                tmp_sql = tmp_sql + " AND site_no = " + siteno;

                // 更新
                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

            /// <summary>
            /// xml要素の順番決定
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, int> Get_Hash_XmlIndex()
            {

                var rtn_hash = new SafeDictionary<string, int>();

                rtn_hash.Add("BaikaiKeiyakuDate", 1);
                rtn_hash.Add("Etc", 2);
                rtn_hash.Add("EtcOemA", 3);
                rtn_hash.Add("EtcOemB", 4);
                rtn_hash.Add("GazoNotDownload", 5);
                rtn_hash.Add("GenteiKokai", 6);
                rtn_hash.Add("IsAkiyaBank", 7);
                rtn_hash.Add("KagiBiko", 8);
                rtn_hash.Add("KagiHokan", 9);
                rtn_hash.Add("Memo", 10);
                rtn_hash.Add("MonthlyEnable", 11);
                rtn_hash.Add("MustContractParking", 12);
                rtn_hash.Add("NaisoCustomize", 13);
                rtn_hash.Add("NaisoCustomizeJoken", 14);
                rtn_hash.Add("NaisoCustomizeNaiyo", 15);
                rtn_hash.Add("OsusumePointScore", 16);
                rtn_hash.Add("ReformBiko", 17);
                rtn_hash.Add("ReformDate", 18);
                rtn_hash.Add("ReformPartBathroom", 19);
                rtn_hash.Add("ReformPartCloth", 20);
                rtn_hash.Add("ReformPartEtc", 21);
                rtn_hash.Add("ReformPartFloor", 22);
                rtn_hash.Add("ReformPartInteriorDesign", 23);
                rtn_hash.Add("ReformPartKitchen", 24);
                rtn_hash.Add("ReformPartOutside", 25);
                rtn_hash.Add("ReformPartToilet", 26);
                rtn_hash.Add("ReformPartWashRoom", 27);
                rtn_hash.Add("RenovationDate", 28);
                rtn_hash.Add("RenovationMemo", 29);
                rtn_hash.Add("SalesStaff", 30);
                rtn_hash.Add("StaffComment", 31);
                rtn_hash.Add("StaffCommentKbn", 32);
                rtn_hash.Add("StaffCommentSettingMethod", 33);
                rtn_hash.Add("StaffCommentTemplateNo", 34);
                rtn_hash.Add("Tasyatorikomi", 35);
                rtn_hash.Add("TokucyouA", 36);
                rtn_hash.Add("TokucyouB", 37);
                rtn_hash.Add("TokuyutinBiko", 38);
                rtn_hash.Add("TokuyutinEnable", 39);
                rtn_hash.Add("TokuyutinJyogen", 40);
                rtn_hash.Add("TokuyutinJyousyoritu", 41);
                rtn_hash.Add("TokuyutinKagen", 42);
                rtn_hash.Add("TokuyutinRyokinHendo", 43);
                rtn_hash.Add("TokuyutinYatinHojyoNensu", 44);
                rtn_hash.Add("Url", 45);
                rtn_hash.Add("UrlKind", 46);
                rtn_hash.Add("YotakusakiKaisya", 47);

                return rtn_hash;

            }

            /// <summary>
            /// xml移行用の文字列の生成
            /// </summary>
            /// <param name="hash_cvxmlitem"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlValue(SafeDictionary<string, string> hash_cvxmlitem, SafeDictionary<string, int> hash_xmlindex)
            {

                var rtn_str = new System.Text.StringBuilder();
                var tmp_slist = new SortedList<int, string>();

                foreach (var midvalue in hash_cvxmlitem)
                {

                    string midfldname = Conversions.ToString(midvalue.Key);
                    string cvvalue = Conversions.ToString(midvalue.Value);
                    string tmp_xmlelement = "";

                    switch (midfldname ?? "")
                    {
                        case "物件の特徴_自社HPに表示":
                            {
                                tmp_xmlelement = "TokucyouA";
                                break;
                            }
                        case "物件の特徴_他社HPに表示":
                            {
                                tmp_xmlelement = "TokucyouB";
                                break;
                            }
                        case "広告備考":
                            {
                                tmp_xmlelement = "Etc";
                                break;
                            }
                        case "広告備考_自社HPに表示":
                            {
                                tmp_xmlelement = "EtcOemA";
                                break;
                            }
                        case "広告備考_他社HPに表示":
                            {
                                tmp_xmlelement = "EtcOemB";
                                break;
                            }
                        case "リンク先区分":
                            {
                                tmp_xmlelement = "UrlKind";
                                break;
                            }
                        case "リンク先URL":
                            {
                                tmp_xmlelement = "Url";
                                break;
                            }
                        case "社内用メモ":
                            {
                                tmp_xmlelement = "Memo";
                                break;
                            }
                        case "リフォーム実施年月":
                            {
                                tmp_xmlelement = "ReformDate";
                                break;
                            }
                        case "リフォームその他箇所":
                            {
                                tmp_xmlelement = "ReformPartEtc";
                                break;
                            }
                        case "リフォーム備考":
                            {
                                tmp_xmlelement = "ReformBiko";
                                break;
                            }
                        case "リノベーション施工完了年月":
                            {
                                tmp_xmlelement = "RenovationDate";
                                break;
                            }
                        case "リノベーション内容":
                            {
                                tmp_xmlelement = "RenovationMemo";
                                break;
                            }
                        case "カスタマイズ可否":
                            {
                                tmp_xmlelement = "NaisoCustomize";
                                break;
                            }
                        case "カスタマイズ内容":
                            {
                                tmp_xmlelement = "NaisoCustomizeNaiyo";
                                break;
                            }
                        case "カスタマイズ条件":
                            {
                                tmp_xmlelement = "NaisoCustomizeJoken";
                                break;
                            }
                        case "特定優遇賃貸住宅チェック":
                            {
                                tmp_xmlelement = "TokuyutinEnable";
                                break;
                            }
                        case "特優賃下限":
                            {
                                tmp_xmlelement = "TokuyutinKagen";
                                break;
                            }
                        case "特優賃上限":
                            {
                                tmp_xmlelement = "TokuyutinJyogen";
                                break;
                            }
                        case "特優賃料金変動区分":
                            {
                                tmp_xmlelement = "TokuyutinRyokinHendo";
                                break;
                            }
                        case "特優賃料金上昇率":
                            {
                                tmp_xmlelement = "TokuyutinJyousyoritu";
                                break;
                            }
                        case "特優賃家賃補助年数":
                            {
                                tmp_xmlelement = "TokuyutinYatinHojyoNensu";
                                break;
                            }
                        case "特優賃備考":
                            {
                                tmp_xmlelement = "TokuyutinBiko";
                                break;
                            }
                        case "鍵保管場所":
                            {
                                tmp_xmlelement = "KagiHokan";
                                break;
                            }
                        case "預託先会社名":
                            {
                                tmp_xmlelement = "YotakusakiKaisya";
                                break;
                            }
                        case "鍵備考":
                            {
                                tmp_xmlelement = "KagiBiko";
                                break;
                            }
                        case "限定公開":
                            {
                                tmp_xmlelement = "GenteiKokai";
                                break;
                            }
                        case "画像ダウンロード不可":
                            {
                                tmp_xmlelement = "GazoNotDownload";
                                break;
                            }
                        case "営業スタッフ名":
                            {
                                tmp_xmlelement = "SalesStaff";
                                break;
                            }
                        case "営業スタッフコメント設定方法":
                            {
                                tmp_xmlelement = "StaffCommentSettingMethod";
                                break;
                            }
                        case "コメントテンプレート":
                            {
                                tmp_xmlelement = "StaffCommentTemplateNo";
                                break;
                            }
                        case "コメント種別":
                            {
                                tmp_xmlelement = "StaffCommentKbn";
                                break;
                            }
                        case "コメント":
                            {
                                tmp_xmlelement = "StaffComment";
                                break;
                            }
                        case "マンスリー可":
                            {
                                tmp_xmlelement = "MonthlyEnable";
                                break;
                            }
                        case "駐車場契約必須":
                            {
                                tmp_xmlelement = "MustContractParking";
                                break;
                            }
                        case "他社取込":
                            {
                                tmp_xmlelement = "Tasyatorikomi";
                                break;
                            }
                        case "媒介契約年月日":
                            {
                                tmp_xmlelement = "BaikaiKeiyakuDate";
                                break;
                            }
                        case "特別広告ポイント数":
                            {
                                tmp_xmlelement = "OsusumePointScore";
                                break;
                            }
                        case "空き家バンク登録物件":
                            {
                                tmp_xmlelement = "IsAkiyaBank";
                                break;
                            }
                    }

                    string xmlelement = "<" + tmp_xmlelement + ">" + cvvalue + "</" + tmp_xmlelement + ">";
                    int keyno = hash_xmlindex[tmp_xmlelement];
                    tmp_slist.Add(keyno, xmlelement);

                }

                foreach (var item in tmp_slist)
                {
                    string xmlvalue = item.Value;
                    rtn_str.Append(xmlvalue);
                }

                return rtn_str.ToString();

            }

        }

    }

    #endregion

    #region 広告補足athome情報

    public class Hydata_kokoku_athome_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_hyguid = new SafeDictionary<string, string>();                                // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_kokoku_Model();           // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列

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

                // 部屋guid取得
                Set_HyGuid(sqlcnnv10, ref hash_hyguid);

                // テーブル名/フィールド名セット
                string tblname = "hydata_kokoku";
                string fldnamegrp = "hy_guid,site_no,item";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string tmp_keysub2 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub2]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;


                    // 作業用変数
                    string tmp_bkno = "";
                    string tmp_hyno = "";
                    var hash_cvxmlitem = new SafeDictionary<string, string>();

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "物件No":
                                {
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }

                            default:
                                {
                                    if (!string.IsNullOrEmpty(fldname))
                                    {
                                        hash_cvxmlitem.Add(fldname, fldvalue);
                                    }

                                    break;
                                }
                        }

                    }

                    // 部屋guid取得用
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["hy_guid"] = hash_hyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["hy_guid"]), VbStrConv.Narrow)];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // xmlの移行
                            Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) ";
                tmp_sql = tmp_sql + " FROM hydata_kokoku AS HYKOKOKU ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

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
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql = tmp_sql + " 	,HY.hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 広告補足のxml項目の移行
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_cvxmlitem"></param>
            /// <remarks></remarks>
            public void Set_XmlItem(SqlConnection sqlcnnv10, SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_cvxmlitem)
            {

                string hyguid = Conversions.ToString(hash_cvitem["hy_guid"]);   // 更新対象選択用
                string siteno = Conversions.ToString(hash_cvitem["site_no"]);   // 更新対象選択用
                var hash_xmlindex = new SafeDictionary<string, int>();              // xml要素の順番決定用

                // xml要素の順番設定
                hash_xmlindex = Get_Hash_XmlIndex();

                // 「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                string xmlupdateqry_xmlpre = "<AthomeKokokuModel xmlns:i=" + "\"" + "http://www.w3.org/2001/XMLSchema-instance" + "\"" + ">";
                string xmlupdateqry_xmlpost = "</AthomeKokokuModel>";
                string xmlitem = Get_XmlValue(hash_cvxmlitem, hash_xmlindex);
                string xmlupdateqry_main = xmlupdateqry_xmlpre + xmlitem + xmlupdateqry_xmlpost;

                // 更新用クエリの作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " DECLARE @tmp_xml XML ";
                tmp_sql = tmp_sql + " SET @tmp_xml = CAST('" + xmlupdateqry_main + "' AS XML) ";
                tmp_sql = tmp_sql + " UPDATE hydata_kokoku SET item = @tmp_xml ";
                tmp_sql = tmp_sql + " WHERE hy_guid = '" + hyguid + "' ";
                tmp_sql = tmp_sql + " AND site_no = " + siteno;

                // 更新
                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

            /// <summary>
            /// xml要素の順番決定
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, int> Get_Hash_XmlIndex()
            {

                var rtn_hash = new SafeDictionary<string, int>();

                rtn_hash.Add("BkNameEnabled", 1);
                rtn_hash.Add("CanDiscountInitialCost", 2);
                rtn_hash.Add("CanDiscountTinryo", 3);
                rtn_hash.Add("Etc", 4);
                rtn_hash.Add("EtcCardPayment", 5);
                rtn_hash.Add("EtcCardPaymentUmu", 6);
                rtn_hash.Add("EtcPet", 7);
                rtn_hash.Add("EtcReform", 8);
                rtn_hash.Add("EtcRenovation", 9);
                rtn_hash.Add("EtcSyokyakuJoken", 10);
                rtn_hash.Add("ExteriorReformKbn", 11);
                rtn_hash.Add("HyNoEnabled", 12);
                rtn_hash.Add("InteriorReformKbn", 13);
                rtn_hash.Add("KagiGentiTaio", 14);
                rtn_hash.Add("KosinRyoKbn", 15);
                rtn_hash.Add("NearBikeOkibaRyokin", 16);
                rtn_hash.Add("NearBikeOkibaTax", 17);
                rtn_hash.Add("NearTyurinjoRyokin", 18);
                rtn_hash.Add("NearTyurinjoTax", 19);
                rtn_hash.Add("PanoramaContentsId", 20);
                rtn_hash.Add("ProComment", 21);
                rtn_hash.Add("ReformUmu", 22);
                rtn_hash.Add("ReformYm", 23);
                rtn_hash.Add("RenovationUmu", 24);
                rtn_hash.Add("RenovationYm", 25);
                rtn_hash.Add("StaffID", 26);
                rtn_hash.Add("YahooEnabled", 27);

                return rtn_hash;

            }

            /// <summary>
            /// xml移行用の文字列の生成
            /// </summary>
            /// <param name="hash_cvxmlitem"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlValue(SafeDictionary<string, string> hash_cvxmlitem, SafeDictionary<string, int> hash_xmlindex)
            {

                var rtn_str = new System.Text.StringBuilder();
                var tmp_slist = new SortedList<int, string>();
                string xml_null = "i:nil=" + "\"" + "true" + "\"";

                foreach (var midvalue in hash_cvxmlitem)
                {

                    string midfldname = Conversions.ToString(midvalue.Key);
                    string cvvalue = Conversions.ToString(midvalue.Value);
                    string tmp_xmlelement = "";

                    switch (midfldname ?? "")
                    {
                        case "プロのコメント":
                            {
                                tmp_xmlelement = "ProComment";
                                break;
                            }
                        case "プロのコメントスタッフID":
                            {
                                tmp_xmlelement = "StaffID";
                                break;
                            }
                        case "広告備考":
                            {
                                tmp_xmlelement = "Etc";
                                break;
                            }
                        case "その他ペット可内容":
                            {
                                tmp_xmlelement = "EtcPet";
                                break;
                            }
                        case "更新料区分":
                            {
                                tmp_xmlelement = "KosinRyoKbn";
                                break;
                            }
                        case "保証金その他償却条件内容":
                            {
                                tmp_xmlelement = "EtcSyokyakuJoken";
                                break;
                            }
                        case "その他クレジット決済可":
                            {
                                tmp_xmlelement = "EtcCardPaymentUmu";
                                break;
                            }
                        case "その他クレジット決済可能条件等":
                            {
                                tmp_xmlelement = "EtcCardPayment";
                                break;
                            }
                        case "リフォーム有無":
                            {
                                tmp_xmlelement = "ReformUmu";
                                break;
                            }
                        case "リフォーム年月":
                            {
                                tmp_xmlelement = "ReformYm";
                                break;
                            }
                        case "対象(内装関連)":
                            {
                                tmp_xmlelement = "InteriorReformKbn";
                                break;
                            }
                        case "対象(外装関連)":
                            {
                                tmp_xmlelement = "ExteriorReformKbn";
                                break;
                            }
                        case "リフォーム内容":
                            {
                                tmp_xmlelement = "EtcReform";
                                break;
                            }
                        case "リノベーション有無":
                            {
                                tmp_xmlelement = "RenovationUmu";
                                break;
                            }
                        case "リノベーション実施年月":
                            {
                                tmp_xmlelement = "RenovationYm";
                                break;
                            }
                        case "リノベーション内容":
                            {
                                tmp_xmlelement = "EtcRenovation";
                                break;
                            }
                        case "建物名表示":
                            {
                                tmp_xmlelement = "BkNameEnabled";
                                break;
                            }
                        case "部屋番号表示":
                            {
                                tmp_xmlelement = "HyNoEnabled";
                                break;
                            }
                        case "鍵現地対応":
                            {
                                tmp_xmlelement = "KagiGentiTaio";
                                break;
                            }
                        case "賃料値下げ可":
                            {
                                tmp_xmlelement = "CanDiscountTinryo";
                                break;
                            }
                        case "初期費用値下げ可":
                            {
                                tmp_xmlelement = "CanDiscountInitialCost";
                                break;
                            }
                        case "パノラマコンテンツID":
                            {
                                tmp_xmlelement = "PanoramaContentsId";
                                break;
                            }
                        case "近隣駐輪場料金":
                            {
                                tmp_xmlelement = "NearTyurinjoRyokin";
                                break;
                            }
                        case "近隣駐輪場料金_税区分":
                            {
                                tmp_xmlelement = "NearTyurinjoTax";
                                break;
                            }
                        case "近隣バイク置き場料金":
                            {
                                tmp_xmlelement = "NearBikeOkibaRyokin";
                                break;
                            }
                        case "近隣バイク置き場料金_税区分":
                            {
                                tmp_xmlelement = "NearBikeOkibaTax";
                                break;
                            }
                        case "Yahoo公開不可":
                            {
                                tmp_xmlelement = "YahooEnabled";
                                break;
                            }
                    }

                    string xmlelement = "<" + tmp_xmlelement + ">" + cvvalue + "</" + tmp_xmlelement + ">";
                    if ((tmp_xmlelement == "NearTyurinjoRyokin" | tmp_xmlelement == "NearBikeOkibaRyokin") & string.IsNullOrEmpty(cvvalue))
                    {
                        xmlelement = xmlelement.Replace("<" + tmp_xmlelement + ">", "<" + tmp_xmlelement + " " + xml_null + ">");
                    }
                    int keyno = hash_xmlindex[tmp_xmlelement];
                    tmp_slist.Add(keyno, xmlelement);

                }

                foreach (var item in tmp_slist)
                {
                    string xmlvalue = item.Value;
                    rtn_str.Append(xmlvalue);
                }

                return rtn_str.ToString();

            }

        }

    }

    #endregion

    #region 広告補足SUUMO情報

    public class Hydata_kokoku_suumo_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_hyguid = new SafeDictionary<string, string>();                                // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_kokoku_Model();           // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列

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

                // 部屋guid取得
                Set_HyGuid(sqlcnnv10, ref hash_hyguid);

                // テーブル名/フィールド名セット
                string tblname = "hydata_kokoku";
                string fldnamegrp = "hy_guid,site_no,item";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string tmp_keysub2 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub2]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;


                    // 作業用変数
                    string tmp_bkno = "";
                    string tmp_hyno = "";
                    var hash_cvxmlitem = new SafeDictionary<string, string>();

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "物件No":
                                {
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }

                            default:
                                {
                                    if (!string.IsNullOrEmpty(fldname))
                                    {
                                        hash_cvxmlitem.Add(fldname, fldvalue);
                                    }

                                    break;
                                }
                        }

                    }

                    // 部屋guid取得用
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["hy_guid"] = hash_hyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["hy_guid"]), VbStrConv.Narrow)];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // xmlの移行
                            Set_XmlItem(sqlcnnv10, hash_cvitem, hash_cvxmlitem);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) ";
                tmp_sql = tmp_sql + " FROM hydata_kokoku AS HYKOKOKU ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYKOKOKU.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

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
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql = tmp_sql + " 	,HY.hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 広告補足のxml項目の移行
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_cvxmlitem"></param>
            /// <remarks></remarks>
            public void Set_XmlItem(SqlConnection sqlcnnv10, SafeDictionary<string, string> hash_cvitem, SafeDictionary<string, string> hash_cvxmlitem)
            {

                string hyguid = Conversions.ToString(hash_cvitem["hy_guid"]);   // 更新対象選択用
                string siteno = Conversions.ToString(hash_cvitem["site_no"]);   // 更新対象選択用
                var hash_xmlindex = new SafeDictionary<string, int>();              // xml要素の順番決定用

                // xml要素の順番設定
                hash_xmlindex = Get_Hash_XmlIndex();

                // 「中間ファイルフィールド名 - 移行値」のハッシュテーブルからxml移行用の文字列を取得
                string xmlupdateqry_xmlpre = "<SuumoKokokuModel xmlns:i=" + "\"" + "http://www.w3.org/2001/XMLSchema-instance" + "\"" + ">";
                string xmlupdateqry_xmlpost = "</SuumoKokokuModel>";
                string xmlitem = Get_XmlValue(hash_cvxmlitem, hash_xmlindex);
                string xmlupdateqry_main = xmlupdateqry_xmlpre + xmlitem + xmlupdateqry_xmlpost;

                // 更新用クエリの作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " DECLARE @tmp_xml XML ";
                tmp_sql = tmp_sql + " SET @tmp_xml = CAST('" + xmlupdateqry_main + "' AS XML) ";
                tmp_sql = tmp_sql + " UPDATE hydata_kokoku SET item = @tmp_xml ";
                tmp_sql = tmp_sql + " WHERE hy_guid = '" + hyguid + "' ";
                tmp_sql = tmp_sql + " AND site_no = " + siteno;

                // 更新
                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

            /// <summary>
            /// xml要素の順番決定
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, int> Get_Hash_XmlIndex()
            {

                var rtn_hash = new SafeDictionary<string, int>();
                string xml_namesp = "xmlns:a=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.N3Lib.Utys" + "\"";

                rtn_hash.Add("AddressEnabled", 1);
                rtn_hash.Add("AllowCopySearchForKaisyakan", 2);
                rtn_hash.Add("BkNameEnabled", 3);
                rtn_hash.Add("EtcCarParking", 4);
                rtn_hash.Add("EtcReform", 5);
                rtn_hash.Add("EtcTokuyutin", 6);
                rtn_hash.Add("FreeCommentForKaisyaKan", 7);
                rtn_hash.Add("FreeCommentForNet", 8);
                rtn_hash.Add("HyNoEnabled", 9);
                rtn_hash.Add("JosyoRitu", 10);
                rtn_hash.Add("JutakuSienSite", 11);
                rtn_hash.Add("KengakuYoyakuKino", 12);
                rtn_hash.Add("KisyaKanriCode1", 13);
                rtn_hash.Add("KisyaKanriCode2", 14);
                rtn_hash.Add("MainCatch1", 15);
                rtn_hash.Add("MainCatch2", 16);
                rtn_hash.Add("NetCatch", 17);
                rtn_hash.Add("NyukyoSyaFutanFrom", 18);
                rtn_hash.Add("NyukyosyaFutanTo", 19);
                rtn_hash.Add("OsusumePickUp1", 20);
                rtn_hash.Add("OsusumePickUp2", 21);
                rtn_hash.Add("OsusumePickUp3", 22);
                rtn_hash.Add("OsusumePickUpRui1", 23);
                rtn_hash.Add("OsusumePickUpRui2", 24);
                rtn_hash.Add("OsusumePickUpRui3", 25);
                rtn_hash.Add("ReformJiki", 26);
                rtn_hash.Add("ReformKasyo", 27);
                rtn_hash.Add("RyokinHendoKbn", 28);
                rtn_hash.Add("SelectedKaguKadenList" + " " + xml_namesp, 29);
                rtn_hash.Add("SelectedNyukyoJokenList" + " " + xml_namesp, 30);
                rtn_hash.Add("SelectedReformList" + " " + xml_namesp, 31);
                rtn_hash.Add("SelectedReformNaiyoList" + " " + xml_namesp, 32);
                rtn_hash.Add("SelectedSecurityList" + " " + xml_namesp, 33);
                rtn_hash.Add("SelectedSitunaiMadoriList" + " " + xml_namesp, 34);
                rtn_hash.Add("SetubiKankyo1", 35);
                rtn_hash.Add("SetubiKankyo1F", 36);
                rtn_hash.Add("SetubiKankyo2", 37);
                rtn_hash.Add("SetubiKankyoKyori1", 38);
                rtn_hash.Add("SetubiKankyoKyori2", 39);
                rtn_hash.Add("SetubiKankyoRinsetu1", 40);
                rtn_hash.Add("SetubiKankyoRinsetu2", 41);
                rtn_hash.Add("SetubiKankyoRinsetuTani1", 42);
                rtn_hash.Add("SetubiKankyoRinsetuTani2", 43);
                rtn_hash.Add("SikikinTumimasiGaku", 44);
                rtn_hash.Add("SikikinTumimasiGakuTani", 45);
                rtn_hash.Add("SikikinTumimasiJoken", 46);
                rtn_hash.Add("SubCatch1", 47);
                rtn_hash.Add("SubCatch10", 48);
                rtn_hash.Add("SubCatch2", 49);
                rtn_hash.Add("SubCatch3", 50);
                rtn_hash.Add("SubCatch4", 51);
                rtn_hash.Add("SubCatch5", 52);
                rtn_hash.Add("SubCatch6", 53);
                rtn_hash.Add("SubCatch7", 54);
                rtn_hash.Add("SubCatch8", 55);
                rtn_hash.Add("SubCatch9", 56);
                rtn_hash.Add("SuumoGazoYusen", 57);
                rtn_hash.Add("TokuyutinUmu", 58);
                rtn_hash.Add("YatinHojoNensu", 59);
                rtn_hash.Add("ZumenPattern", 60);

                return rtn_hash;

            }

            /// <summary>
            /// xml移行用の文字列の生成
            /// </summary>
            /// <param name="hash_cvxmlitem"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_XmlValue(SafeDictionary<string, string> hash_cvxmlitem, SafeDictionary<string, int> hash_xmlindex)
            {

                var rtn_str = new System.Text.StringBuilder();
                var tmp_slist = new SortedList<int, string>();

                foreach (var midvalue in hash_cvxmlitem)
                {

                    string midfldname = Conversions.ToString(midvalue.Key);
                    string cvvalue = Conversions.ToString(midvalue.Value);
                    string tmp_xmlelement = "";
                    string xml_namesp = "xmlns:a=" + "\"" + "http://schemas.datacontract.org/2004/07/Njc.N3Lib.Utys" + "\"";
                    string xml_null = "i:nil=" + "\"" + "true" + "\"";

                    switch (midfldname ?? "")
                    {
                        case "物件の特徴_ネット用キャッチ":
                            {
                                tmp_xmlelement = "NetCatch";
                                break;
                            }
                        case "物件の特徴_フリーコメント":
                            {
                                tmp_xmlelement = "FreeCommentForNet";
                                break;
                            }
                        case "敷金積増条件":
                            {
                                tmp_xmlelement = "SikikinTumimasiJoken";
                                break;
                            }
                        case "敷金積増後総額":
                            {
                                tmp_xmlelement = "SikikinTumimasiGaku";
                                break;
                            }
                        case "敷金積増後総額_単位":
                            {
                                tmp_xmlelement = "SikikinTumimasiGakuTani";
                                break;
                            }
                        case "駐車場備考":
                            {
                                tmp_xmlelement = "EtcCarParking";
                                break;
                            }
                        case "特優賃有無":
                            {
                                tmp_xmlelement = "TokuyutinUmu";
                                break;
                            }
                        case "特優賃入居者負担額_下限":
                            {
                                tmp_xmlelement = "NyukyoSyaFutanFrom";
                                break;
                            }
                        case "特優賃入居者負担額_上限":
                            {
                                tmp_xmlelement = "NyukyosyaFutanTo";
                                break;
                            }
                        case "特優賃料金変動区分":
                            {
                                tmp_xmlelement = "RyokinHendoKbn";
                                break;
                            }
                        case "特優賃料金変動区分上昇率":
                            {
                                tmp_xmlelement = "JosyoRitu";
                                break;
                            }
                        case "特優賃家賃補助年数":
                            {
                                tmp_xmlelement = "YatinHojoNensu";
                                break;
                            }
                        case "特優賃補足":
                            {
                                tmp_xmlelement = "EtcTokuyutin";
                                break;
                            }
                        case "リフォーム時期":
                            {
                                tmp_xmlelement = "ReformJiki";
                                break;
                            }
                        case "リフォーム箇所":
                            {
                                tmp_xmlelement = "ReformKasyo";
                                break;
                            }
                        case "リフォーム補足":
                            {
                                tmp_xmlelement = "EtcReform";
                                break;
                            }
                        case "設備環境1":
                            {
                                tmp_xmlelement = "SetubiKankyo1";
                                break;
                            }
                        case "設備環境1_距離":
                            {
                                tmp_xmlelement = "SetubiKankyoKyori1";
                                break;
                            }
                        case "設備環境2":
                            {
                                tmp_xmlelement = "SetubiKankyo2";
                                break;
                            }
                        case "設備環境2_距離":
                            {
                                tmp_xmlelement = "SetubiKankyoKyori2";
                                break;
                            }
                        case "周辺環境隣接1":
                            {
                                tmp_xmlelement = "SetubiKankyoRinsetu1";
                                break;
                            }
                        case "周辺環境隣接1_単位":
                            {
                                tmp_xmlelement = "SetubiKankyoRinsetuTani1";
                                break;
                            }
                        case "周辺環境隣接2":
                            {
                                tmp_xmlelement = "SetubiKankyoRinsetu2";
                                break;
                            }
                        case "周辺環境隣接2_単位":
                            {
                                tmp_xmlelement = "SetubiKankyoRinsetuTani2";
                                break;
                            }
                        case "設備環境1F":
                            {
                                tmp_xmlelement = "SetubiKankyo1F";
                                break;
                            }
                        case "セキュリティ":
                            {
                                tmp_xmlelement = "SelectedSecurityList" + " " + xml_namesp;
                                break;
                            }
                        case "室内間取":
                            {
                                tmp_xmlelement = "SelectedSitunaiMadoriList" + " " + xml_namesp;
                                break;
                            }
                        case "家具・家電":
                            {
                                tmp_xmlelement = "SelectedKaguKadenList" + " " + xml_namesp;
                                break;
                            }
                        case "リフォーム":
                            {
                                tmp_xmlelement = "SelectedReformList" + " " + xml_namesp;
                                break;
                            }
                        case "リフォーム(内容)":
                            {
                                tmp_xmlelement = "SelectedReformNaiyoList" + " " + xml_namesp;
                                break;
                            }
                        case "入居条件":
                            {
                                tmp_xmlelement = "SelectedNyukyoJokenList" + " " + xml_namesp;
                                break;
                            }
                        case "SUUMO以外の災害時住宅支援サイトへの掲載":
                            {
                                tmp_xmlelement = "JutakuSienSite";
                                break;
                            }
                        case "おすすめピックアップピクト指定1":
                            {
                                tmp_xmlelement = "OsusumePickUpRui1";
                                break;
                            }
                        case "特徴1":
                            {
                                tmp_xmlelement = "OsusumePickUp1";
                                break;
                            }
                        case "おすすめピックアップピクト指定2":
                            {
                                tmp_xmlelement = "OsusumePickUpRui2";
                                break;
                            }
                        case "特徴2":
                            {
                                tmp_xmlelement = "OsusumePickUp2";
                                break;
                            }
                        case "おすすめピックアップピクト指定3":
                            {
                                tmp_xmlelement = "OsusumePickUpRui3";
                                break;
                            }
                        case "特徴3":
                            {
                                tmp_xmlelement = "OsusumePickUp3";
                                break;
                            }
                        case "物件名公開":
                            {
                                tmp_xmlelement = "BkNameEnabled";
                                break;
                            }
                        case "部屋番号公開":
                            {
                                tmp_xmlelement = "HyNoEnabled";
                                break;
                            }
                        case "詳細住所公開":
                            {
                                tmp_xmlelement = "AddressEnabled";
                                break;
                            }
                        case "会社間物件検索コピー許可":
                            {
                                tmp_xmlelement = "AllowCopySearchForKaisyakan";
                                break;
                            }
                        case "図面パターン":
                            {
                                tmp_xmlelement = "ZumenPattern";
                                break;
                            }
                        case "メインキャッチ1":
                            {
                                tmp_xmlelement = "MainCatch1";
                                break;
                            }
                        case "メインキャッチ2":
                            {
                                tmp_xmlelement = "MainCatch2";
                                break;
                            }
                        case "サブキャッチ1":
                            {
                                tmp_xmlelement = "SubCatch1";
                                break;
                            }
                        case "サブキャッチ2":
                            {
                                tmp_xmlelement = "SubCatch2";
                                break;
                            }
                        case "サブキャッチ3":
                            {
                                tmp_xmlelement = "SubCatch3";
                                break;
                            }
                        case "サブキャッチ4":
                            {
                                tmp_xmlelement = "SubCatch4";
                                break;
                            }
                        case "サブキャッチ5":
                            {
                                tmp_xmlelement = "SubCatch5";
                                break;
                            }
                        case "サブキャッチ6":
                            {
                                tmp_xmlelement = "SubCatch6";
                                break;
                            }
                        case "サブキャッチ7":
                            {
                                tmp_xmlelement = "SubCatch7";
                                break;
                            }
                        case "サブキャッチ8":
                            {
                                tmp_xmlelement = "SubCatch8";
                                break;
                            }
                        case "サブキャッチ9":
                            {
                                tmp_xmlelement = "SubCatch9";
                                break;
                            }
                        case "サブキャッチ10":
                            {
                                tmp_xmlelement = "SubCatch10";
                                break;
                            }
                        case "会社間用補足フリーコメント":
                            {
                                tmp_xmlelement = "FreeCommentForKaisyaKan";
                                break;
                            }
                        case "SUUMO内優先画像":
                            {
                                tmp_xmlelement = "SuumoGazoYusen";
                                break;
                            }
                        case "貴社管理コード1":
                            {
                                tmp_xmlelement = "KisyaKanriCode1";
                                break;
                            }
                        case "貴社管理コード2":
                            {
                                tmp_xmlelement = "KisyaKanriCode2";
                                break;
                            }
                        case "見学予約機能を利用する":
                            {
                                tmp_xmlelement = "KengakuYoyakuKino";
                                break;
                            }
                    }

                    string xmlelement = "<" + tmp_xmlelement + ">" + cvvalue + "</" + tmp_xmlelement.Replace(xml_namesp, "") + ">";
                    if ((tmp_xmlelement == "JosyoRitu" | tmp_xmlelement == "NyukyoSyaFutanFrom" | tmp_xmlelement == "NyukyosyaFutanTo" | tmp_xmlelement == "SetubiKankyoKyori1" | tmp_xmlelement == "SetubiKankyoKyori2" | tmp_xmlelement == "SikikinTumimasiGaku" | tmp_xmlelement == "YatinHojoNensu") & string.IsNullOrEmpty(cvvalue))






                    {
                        xmlelement = xmlelement.Replace("<" + tmp_xmlelement + ">", "<" + tmp_xmlelement + " " + xml_null + ">");
                    }
                    int keyno = hash_xmlindex[tmp_xmlelement];
                    tmp_slist.Add(keyno, xmlelement);

                }

                foreach (var item in tmp_slist)
                {
                    string xmlvalue = item.Value;
                    rtn_str.Append(xmlvalue);
                }

                return rtn_str.ToString();

            }

        }

    }

    #endregion

    #region ポータル連動部屋分類情報

    public class M_hy_ruisite_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.M_hy_ruisite_Model();            // 移行値格納用モデル初期化
                int keycol_main = 2;                                  // メインキー列
                int keycol_sub1 = 3;                                  // サブ1キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Hyrui();

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // テーブル名/フィールド名セット
                string tblname = "m_hy_ruisite";
                string fldnamegrp = "hy_ruino,site_no,hy_ruisiteitemdata";

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1;

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "部屋分類名":
                                {
                                    model_cvitem.Vari_Hy_ruino = fldvalue;
                                    break;
                                }
                            case "サイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }
                            case "サイト用部屋分類No":
                                {
                                    model_cvitem.Vari_Hy_ruisiteitemdata = fldvalue;
                                    break;
                                }
                        }

                    }

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_jisyano) ";
                tmp_sql = tmp_sql + " FROM hydata_sosin AS HYSOSIN ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYSOSIN.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 部屋分類取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_Hyrui(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash_hyrui)
            {

                string tmp_sql = "SELECT hy_ruiname,hy_ruino FROM m_hy_rui ORDER BY hy_ruino";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_hyrui);

            }

        }

    }

    #endregion

    #region 部屋毎送信情報

    public class Hydata_sosin_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_hyguid = new SafeDictionary<string, string>();                                // キー/guid格納用ハッシュテーブル
                var hash_komkguid = new SafeDictionary<string, string>();

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_sosin_Model();            // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列

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

                // guid取得
                Set_HyGuid(sqlcnnv10, ref hash_hyguid);          // 部屋guid
                Set_KomokGuid(sqlcnnv10, ref hash_komkguid);     // 項目guid

                // テーブル名/フィールド名セット
                string tblname = "hydata_sosin";
                string fldnamegrp = "hy_guid,site_no,site_jisyano,select_no,komok_guid," + "history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));
                string fldname_keysub3 = Conversions.ToString(headervalue(startrow - 1, keycol_sub3));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string tmp_keysub2 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub2]);
                    string tmp_keysub3 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub3]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



                    // 作業用変数
                    string tmp_bkno = "";
                    string tmp_hyno = "";
                    string tmp_siteno = "";
                    string tmp_sitejisyano = "";
                    string tmp_komok = "";

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "物件No":
                                {
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "ポータルサイトNo":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    tmp_siteno = fldvalue;
                                    break;
                                }
                            case "自社サービスNo":
                                {
                                    model_cvitem.Vari_Site_jisyano = fldvalue;
                                    tmp_sitejisyano = fldvalue;
                                    break;
                                }
                            case "選択No":
                                {
                                    model_cvitem.Vari_Select_no = fldvalue;
                                    break;
                                }
                            case "項目値":
                                {
                                    tmp_komok = fldvalue;
                                    break;
                                }
                        }

                    }

                    // guid取得用
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;                                   // 部屋guid
                    model_cvitem.Vari_Komok_guid = tmp_siteno + "-" + tmp_sitejisyano + "-" + tmp_komok;     // 項目guid

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
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["hy_guid"] = hash_hyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["hy_guid"]), VbStrConv.Narrow)];

                        // 項目をguidへ変換
                        hash_cvitem["komok_guid"] = hash_komkguid[Strings.StrConv(Conversions.ToString(hash_cvitem["komok_guid"]), VbStrConv.Narrow)];

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_jisyano) ";
                tmp_sql = tmp_sql + " FROM hydata_sosin AS HYSOSIN ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYSOSIN.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

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
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql = tmp_sql + " 	,HY.hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 項目guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_KomokGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,site_no) + '-' + CONVERT(varchar,site_jisyano)  + '-' + CONVERT(varchar,komok_value) AS キー ";
                tmp_sql = tmp_sql + " 	,komok_guid ";
                tmp_sql = tmp_sql + " FROM m_site_komok ";
                tmp_sql = tmp_sql + " WHERE site_no IN (10,20,30,40) ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

        }

    }

    #endregion

    #region BtoBグループ設定情報

    public class Hydata_btobgroup_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_hyguid = new SafeDictionary<string, string>();                                // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_btobgroup_Model();        // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列

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

                // 部屋guid取得
                Set_HyGuid(sqlcnnv10, ref hash_hyguid);

                // テーブル名/フィールド名セット
                string tblname = "hydata_btobgroup";
                string fldnamegrp = "hy_guid,group_no,display_kbn,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string tmp_keysub2 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub2]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;


                    // 作業用変数
                    string tmp_bkno = "";
                    string tmp_hyno = "";

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "物件No":
                                {
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "BtoBプラグイングループNo":
                                {
                                    model_cvitem.Vari_Group_no = fldvalue;
                                    break;
                                }
                            case "グループ公開有無":
                                {
                                    model_cvitem.Vari_Display_kbn = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 部屋guid取得用
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;

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
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["hy_guid"] = hash_hyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["hy_guid"]), VbStrConv.Narrow)];

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,group_no) ";
                tmp_sql = tmp_sql + " FROM hydata_btobgroup AS HYBB ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYBB.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

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
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql = tmp_sql + " 	,HY.hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

        }

    }

    #endregion

    #region 地図表示詳細設定情報

    public class Hydata_mapdisp_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納

                var excelfile = new ExcelFileManager();                // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト(一棟所有)
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_hyguid = new SafeDictionary<string, string>();                                // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_mapdisp_Model();          // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列

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

                // 部屋guid取得
                Set_HyGuid(sqlcnnv10, ref hash_hyguid);

                // テーブル名/フィールド名セット
                string tblname = "hydata_mapdisp";
                string fldnamegrp = "hy_guid,site_no,site_jisyano,map_dispflg,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));
                string fldname_keysub3 = Conversions.ToString(headervalue(startrow - 1, keycol_sub3));

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string tmp_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string tmp_keysub2 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub2]);
                    string tmp_keysub3 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub3]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



                    // 作業用変数
                    string tmp_bkno = "";
                    string tmp_hyno = "";

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "物件No":
                                {
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "外部サイトNo（FK）":
                                {
                                    model_cvitem.Vari_Site_no = fldvalue;
                                    break;
                                }
                            case "自社サイトNo":
                                {
                                    model_cvitem.Vari_Site_jisyano = fldvalue;
                                    break;
                                }
                            case "地図上に表示フラグ":
                                {
                                    model_cvitem.Vari_Map_dispflg = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 部屋guid取得用
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;

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
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["hy_guid"] = hash_hyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["hy_guid"]), VbStrConv.Narrow)];

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
                        pgbcnt = cntii;
                    }
                    else if (rowcnt > cntii + 1)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, ref pgbcnt);
                    }
                    else if (rowcnt <= cntii + 1)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
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

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

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

            /// <summary>
            /// 既存データ取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,BK.bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,HY.hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,site_jisyano) ";
                tmp_sql = tmp_sql + " FROM hydata_mapdisp AS HYMAP ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON HYMAP.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";

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
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql = tmp_sql + " 	,HY.hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

        }

    }

    #endregion

}