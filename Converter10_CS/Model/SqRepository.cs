using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Converter10.Njc.N3Lib.Utys;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Repository
{

    #region 過剰金情報

    public class Azukanri_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="syorikomok"></param>
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Azukanri_Model();                // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                // Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                string viewnamebktobkguid = CommonModule.PRE_VIEW_NAME + "物件キー情報";
                string viewnamebkhytohyguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebktobkguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "azukanri";
                string fldnamegrp = "azu_guid,azu_ymd,nkin_no,dispnkin_name,nkbn_no," + "azu_gak,stakeholder_kbn,stakeholder_no,sui_guid,tanto_no," + "bk_guid,hy_guid,ky_guid,ky_recno,tuki_kbn," + "ky_nkin_no,nkin_recno,kynkin_guid,yoteiazu_flg,printazuryo_guid," + "siwake_flg,biko,history,tujo_siwakegak";




                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebktobkguid, ref hash_bkguid);
                Get_Guid(sqlcnnv10, viewnamebkhytohyguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, ref list_basekeydata);

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

                // 出納情報移行用モデル
                var model_suicvitem = new Model.Suikanri_Model();

                // 20160624 請求関連情報 親データ無しログ出力制御修正 -add
                int chg_totalcnt = rowcnt;        // 親データ無しのデータ件数調整用(カウントしないため)

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
                // Dim fldname_keysub4 As String = headervalue(startrow - 1, keycol_sub4)

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
                    // Dim tmp_keysub4 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub4))
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;
                    model_cvitem.Vari_Bk_guid = tmp_keymain;
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



                    // 作業用変数
                    string tmp_nkinname = "";
                    string tmp_nkinkbn = "";

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
                                    break;
                                }
                            // キー取得済み
                            case "部屋No":
                                {
                                    break;
                                }
                            // キー取得済み
                            case "契約No":
                                {
                                    break;
                                }
                            // キー取得済み
                            case "契約レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "預り金処理日":
                                {
                                    model_cvitem.Vari_Azu_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目名":
                                {
                                    tmp_nkinname = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目区分":
                                {
                                    var a = fldvalue.Trim();
                                    tmp_nkinkbn = EtcMethod.Get_Nkinruiname(ref a);
                                    break;
                                }
                            case "表示用入金項目名":
                                {
                                    model_cvitem.Vari_Dispnkin_name = fldvalue.Trim();
                                    break;
                                }
                            case "入金区分":
                                {
                                    model_cvitem.Vari_Nkbn_no = fldvalue.Trim();
                                    break;
                                }
                            case "預り額":
                                {
                                    model_cvitem.Vari_Azu_gak = fldvalue.Trim();
                                    break;
                                }
                            case "利害関係者区分":
                                {
                                    model_cvitem.Vari_Stakeholder_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "利害関係者No":
                                {
                                    model_cvitem.Vari_Stakeholder_no = fldvalue.Trim();
                                    break;
                                }
                            case "担当者No":
                                {
                                    model_cvitem.Vari_Tanto_no = fldvalue.Trim();
                                    break;
                                }
                            case "月区分":
                                {
                                    model_cvitem.Vari_Tuki_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "契約入金情報の入金項目No":
                                {
                                    model_cvitem.Vari_Ky_nkin_no = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目レコードNo":
                                {
                                    model_cvitem.Vari_Nkin_recno = fldvalue.Trim();
                                    break;
                                }
                            case "預り予定フラグ":
                                {
                                    model_cvitem.Vari_Yoteiazu_flg = fldvalue.Trim();
                                    break;
                                }
                            case "仕訳フラグ":
                                {
                                    model_cvitem.Vari_Siwake_flg = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "通常仕訳額":
                                {
                                    model_cvitem.Vari_Tujo_siwakegak = fldvalue.Trim();
                                    break;
                                }
                            case "家賃入金口座No":
                                {
                                    model_suicvitem.Vari_Yatin_fkom_kozano = fldvalue.Trim();      // 出納情報移行用
                                    break;
                                }
                        }

                    }

                    // 入金項目格納
                    model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + tmp_nkinkbn;

                    // 固定値
                    model_cvitem.Vari_Azu_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Sui_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Kynkin_guid = "";
                    model_cvitem.Vari_Printazuryo_guid = "";
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
                        hash_cvitem["bk_guid"] = hash_bkguid[hash_cvitem["bk_guid"]];
                        hash_cvitem["hy_guid"] = hash_bkhyguid[hash_cvitem["hy_guid"]];
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                            // 出納情報へ反映
                            bool suitblinitflg = Conversions.ToBoolean(Interaction.IIf(tmp_cvcnt == 1, true, false));
                            Set_Suikanri(sqlcnnv10, fldnamegrp, hash_cvitem, model_suicvitem, suitblinitflg);
                        }
                    }

                    // 20160624 請求関連情報 親データ無しログ出力制御修正 -add sta
                    else if (hash_log.Count == 1)
                    {
                        bool logclearflg = false;
                        foreach (var erritem in hash_log)
                        {
                            string tmp_str = Conversions.ToString(erritem.Value);
                            string[] tmp_errnaiyo = tmp_str.Split('-');
                            if ((tmp_errnaiyo[0] ?? "") == CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE)
                            {
                                chg_totalcnt = chg_totalcnt - 1;
                                logclearflg = true;
                            }
                        }
                        if (logclearflg)
                        {
                            hash_log.Clear();
                        }
                        // 20160624 請求関連情報 親データ無しログ出力制御修正 -add end

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
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = chg_totalcnt;
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg end

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebktobkguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN kydata ON " + tblname + ".ky_guid = kydata.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// guid取得→ハッシュテーブル格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="hash_guid"></param>
            /// <remarks></remarks>
            public void Get_Guid(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_guid)
            {

                string tmp_sql = "SELECT * FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);

            }

            /// <summary>
            /// キー/guidをセットにしたVIEWの作成
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                // VIEW初期化
                Drop_TmpView(sqlcnnv10, viewname);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT bk_no AS キー,bk_guid FROM bkdata ";
                            break;
                        }
                    case "物件部屋キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,hy_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋契約キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                            break;
                        }
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 作業用VIEWの削除
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="viewname"></param>
            /// <remarks></remarks>
            public void Drop_TmpView(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

            }

            /// <summary>
            /// 過剰金情報を出納情報へ反映
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="azfldnamegrp"></param>
            /// <param name="hash_azcvitem"></param>
            /// <param name="addkomk"></param>
            /// <remarks></remarks>
            public void Set_Suikanri(SqlConnection sqlcnnv10, string azfldnamegrp, SafeDictionary<string, string> hash_azcvitem, Model.Suikanri_Model model_suicvitem, bool suitblinitflg)
            {

                bool normalflg = true;
                var hash_suicvitem = new SafeDictionary<string, string>();

                string suitblname = "suikanri";
                string suifldnamegrp = "sui_guid,sui_ymd,sui_kbn,stakeholder_kbn,stakeholder_no," + "nkbn_no,gak,tanto_no,jisya_no,yatin_fkom_kozano," + "fkae_rirekino,ns_no,ns_bunkatu_no,biko,history," + "cvpay_no,kokyaku_bango,fkae_no";



                // テーブル初期化処理
                if (suitblinitflg)
                {
                    // 初期化処理実行
                    int rowcnt = 0;
                    string qrywhere = " stakeholder_kbn = 100 ";
                    string tmp_sqldelete = DBQuery.Qry_DelInfo(suitblname, qrywhere);
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldelete, ref rowcnt);
                }

                // 過剰金情報から出納情報を取得
                foreach (var azitem in hash_azcvitem)
                {

                    string azfldname = Conversions.ToString(azitem.Key);
                    string azvalue = Conversions.ToString(azitem.Value);

                    switch (azfldname ?? "")
                    {
                        case "azu_ymd":
                            {
                                hash_suicvitem.Add("sui_ymd", azvalue);
                                break;
                            }
                        case "nkbn_no":
                            {
                                hash_suicvitem.Add("nkbn_no", azvalue);
                                break;
                            }
                        case "azu_gak":
                            {
                                hash_suicvitem.Add("gak", azvalue);
                                break;
                            }
                        case "stakeholder_no":
                            {
                                hash_suicvitem.Add("stakeholder_no", azvalue);
                                break;
                            }
                        case "sui_guid":
                            {
                                hash_suicvitem.Add("sui_guid", azvalue);
                                break;
                            }
                        case "history":
                            {
                                hash_suicvitem.Add("history", azvalue);
                                break;
                            }
                    }

                }

                // 過剰金情報に存在しない出納情報を取得
                hash_suicvitem.Add("yatin_fkom_kozano", model_suicvitem.Vari_Yatin_fkom_kozano);

                // 固定値
                hash_suicvitem.Add("sui_kbn", "1");
                hash_suicvitem.Add("stakeholder_kbn", "100");
                hash_suicvitem.Add("tanto_no", "");
                hash_suicvitem.Add("jisya_no", "");
                hash_suicvitem.Add("fkae_rirekino", "");
                hash_suicvitem.Add("ns_no", "");
                hash_suicvitem.Add("ns_bunkatu_no", "");
                hash_suicvitem.Add("biko", "");
                hash_suicvitem.Add("cvpay_no", "");
                hash_suicvitem.Add("kokyaku_bango", "");
                hash_suicvitem.Add("fkae_no", "");

                // 挿入処理
                CVDBInsert.Cnv_Db(sqlcnnv10, suitblname, suifldnamegrp, hash_suicvitem, ref normalflg);

            }

        }

    }

    #endregion

    #region 出納情報 (過剰金移行)

    // 過剰金移行と同時に作成するためここでは処理をしない

    // Public Class Suikanri_Repository

    // Public Class SubConv
    // Implements IConv

    // ''' <summary>
    // ''' 【中間ファイル→変数】
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="syorikomok"></param>
    // ''' <returns></returns>
    // ''' <remarks></remarks>
    // Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

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
    // Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト(一棟所有)
    // Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
    // Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
    // Dim hash_bkguid As New SafeDictionary<string, string>                            'キー/guid格納用ハッシュテーブル
    // Dim hash_bkhyguid As New SafeDictionary<string, string>                            'キー/guid格納用ハッシュテーブル
    // Dim hash_bkhykyguid As New SafeDictionary<string, string>                            'キー/guid格納用ハッシュテーブル

    // Dim tmp_hash As New SafeDictionary<string, string>                                   '作業用ハッシュテーブル
    // Dim normalflg As Boolean                                        'INSERT正常終了フラグ
    // Dim rtn As Boolean = True                                       '戻り値

    // Dim model_cvitem As New Njc.Model.Suikanri_Model                '移行値格納用モデル初期化
    // Dim keycol_main As Integer = 1                                  'メインキー列
    // Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
    // Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
    // Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
    // Dim viewnamebkhykytosuiguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

    // '************************
    // '作業準備
    // '************************

    // '紐付けデータ取得
    // Call SetRelItemToObject.Set_RelData_Nkinkbn()

    // 'Excelファイル初期設定
    // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
    // If Not rtn Then
    // Return rtn
    // End If

    // 'キー取得用のVIEWを作成
    // Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytosuiguid)

    // 'テーブル名/フィールド名セット
    // Dim tblname As String = "suikanri"
    // Dim fldnamegrp As String = "sui_guid,sui_ymd,sui_kbn,stakeholder_kbn,stakeholder_no," & _
    // "nkbn_no,gak,tanto_no,jisya_no,yatin_fkom_kozano," & _
    // "fkae_rirekino,ns_no,ns_bunkatu_no,biko,history," & _
    // "cvpay_no,kokyaku_bango,fkae_no"

    // '追加コンバート時の重複チェック用に既存データのキーを取得
    // If Not InitDBFlg Then
    // Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
    // End If

    // 'guidを取得
    // Call Me.Get_Guid(sqlcnnv10, viewnamebktobkguid, hash_bkguid)
    // Call Me.Get_Guid(sqlcnnv10, viewnamebkhytohyguid, hash_bkhyguid)
    // Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

    // '親マスタ取得 (過剰金は物件Noがあれば登録可能)
    // Call Get_BaseKey(sqlcnnv10, viewnamebktobkguid, list_basekeydata)

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
    // Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    // 'キーヘッダー名取得
    // Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
    // Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
    // Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)
    // Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)
    // 'Dim fldname_keysub4 As String = headervalue(startrow - 1, keycol_sub4)

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
    // Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

    // 'キー値取得
    // Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
    // Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
    // Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
    // Dim tmp_keysub3 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub3))
    // 'Dim tmp_keysub4 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub4))
    // Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
    // .Vari_Bk_guid = tmp_keymain
    // .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1
    // .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

    // 'ログ出力用
    // Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
    // fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
    // fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
    // fldname_keysub3 & " = " & tmp_keysub3

    // '移行値取得
    // For cntjj = 1 To columncnt

    // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    // Dim fldvalue As String = ""
    // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
    // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
    // End If

    // Select Case fldname
    // Case "物件No"
    // 'キー取得済み
    // Case "部屋No"
    // 'キー取得済み
    // Case "契約No"
    // 'キー取得済み
    // Case "契約レコードNo"
    // .Vari_Ky_recno = fldvalue.Trim
    // Case "預り金処理日"
    // .Vari_Azu_ymd = fldvalue.Trim
    // Case "入金項目名"
    // .Vari_Nkin_no = fldvalue.Trim & "-" & EtcMethod.Get_Nkinruiname("6")
    // Case "表示用入金項目名"
    // .Vari_Dispnkin_name = fldvalue.Trim
    // Case "入金区分"
    // .Vari_Nkbn_no = fldvalue.Trim
    // Case "預り額"
    // .Vari_Azu_gak = fldvalue.Trim
    // Case "利害関係者区分"
    // .Vari_Stakeholder_kbn = fldvalue.Trim
    // Case "利害関係者No"
    // .Vari_Stakeholder_no = fldvalue.Trim
    // Case "担当者No"
    // .Vari_Tanto_no = fldvalue.Trim
    // Case "月区分"
    // .Vari_Tuki_kbn = fldvalue.Trim
    // Case "契約入金情報の入金項目No"
    // .Vari_Ky_nkin_no = fldvalue.Trim
    // Case "入金項目レコードNo"
    // .Vari_Nkin_recno = fldvalue.Trim
    // Case "預り予定フラグ"
    // .Vari_Yoteiazu_flg = fldvalue.Trim
    // Case "仕訳フラグ"
    // .Vari_Siwake_flg = fldvalue.Trim
    // Case "備考"
    // .Vari_Biko = fldvalue.Trim
    // Case "通常仕訳額"
    // .Vari_Tujo_siwakegak = fldvalue.Trim
    // End Select

    // Next

    // '固定値
    // .Vari_Azu_guid = Guid.NewGuid.ToString
    // .Vari_Sui_guid = Guid.NewGuid.ToString
    // .Vari_Kynkin_guid = ""
    // .Vari_Printazuryo_guid = ""
    // .Vari_History = DefHistory

    // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
    // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

    // 'データチェック
    // Dim skipflg As Boolean = False
    // Dim hash_cvitem As New SafeDictionary<string, string>
    // Dim hash_log As New SafeDictionary<string, string>
    // skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

    // '書込処理
    // If Not skipflg Then

    // 'キーをguidへ変換
    // hash_cvitem("bk_guid") = hash_bkguid.Item(hash_cvitem.Item("bk_guid"))
    // hash_cvitem("hy_guid") = hash_bkhyguid.Item(hash_cvitem.Item("hy_guid"))
    // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))

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

    // '作業用VIEWの削除
    // Call Me.Drop_TmpView(sqlcnnv10, viewnamebktobkguid)
    // Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid)
    // Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)

    // 'Excelファイル終了設定
    // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    // '返却
    // Return rtn

    // End Function

    // ''' <summary>
    // ''' 親マスタ取得→リスト格納
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="tblname"></param>
    // ''' <param name="list_basedata"></param>
    // ''' <remarks></remarks>
    // Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

    // Dim tmp_sql As String = " SELECT キー FROM " & tblname
    // Dim flg As Boolean = True

    // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

    // End Sub

    // Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

    // End Function

    // Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As SafeDictionary<string, string>, ByRef hash_chkafter As SafeDictionary<string, string>, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

    // End Function

    // Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

    // Dim tmp_sql As String = ""
    // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
    // tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
    // tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
    // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
    // Dim flg As Boolean = True

    // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

    // End Sub

    // ''' <summary>
    // ''' guid取得→ハッシュテーブル格納
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="tblname"></param>
    // ''' <param name="hash_guid"></param>
    // ''' <remarks></remarks>
    // Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As SafeDictionary<string, string>)

    // Dim tmp_sql As String = "SELECT * FROM " & tblname
    // Dim flg As Boolean = True

    // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)

    // End Sub

    // ''' <summary>
    // ''' キー/guidをセットにしたVIEWの作成
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <remarks></remarks>
    // Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

    // Dim tmpcnt As Integer = 0

    // 'VIEW初期化
    // Call Me.Drop_TmpView(sqlcnnv10, viewname)

    // 'VIEW作成
    // Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

    // tmp_sql_create = tmp_sql_create & " SELECT "
    // tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
    // tmp_sql_create = tmp_sql_create & " 	,ky_guid "
    // tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
    // tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
    // tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "

    // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

    // End Sub

    // ''' <summary>
    // ''' 作業用VIEWの削除
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="viewname"></param>
    // ''' <remarks></remarks>
    // Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

    // Dim tmpcnt As Integer = 0

    // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
    // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

    // End Sub

    // End Class

    // End Class

    #endregion

    #region 運用開始時未収滞納金情報

    public class Unyotainodata_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="syorikomok"></param>
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Unyotainodata_Model();                // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                // Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                string viewnamebktobkguid = CommonModule.PRE_VIEW_NAME + "物件キー情報";
                string viewnamebkhytohyguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebktobkguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "unyotainodata";
                string fldnamegrp = "unyotaino_guid,bk_guid,hy_guid,ky_guid,ky_recno," + "gt_ym,sq_simeymd,nkin_no,nkbn_yotei,yotei_yatin_kozano," + "sqsaki_no,sq_gak,sq_zeikbn,sq_zeigak,so_gak," + "so_zeikbn,so_zeigak,sosumi_flg,kanrigak_rit,sosaki_no," + "sokoza_no,so_rit,soymd_gaitoukbn,soymd_simekbn,soymd_sokintukikbn," + "soymd_sokinsimekbn,zei_rit,sqbuild_flg,tanto_no,biko," + "history,kanri_taxflg,kanritesu_zeiumukbn,kanri_utizeiflg";






                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebktobkguid, ref hash_bkguid);
                Get_Guid(sqlcnnv10, viewnamebkhytohyguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, ref list_basekeydata);

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

                // 20160624 請求関連情報 親データ無しログ出力制御修正 -add
                int chg_totalcnt = rowcnt;        // 親データ無しのデータ件数調整用(カウントしないため)

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
                    model_cvitem.Vari_Bk_guid = tmp_keymain;
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



                    // 作業用変数
                    string tmp_nkinname = "";
                    string tmp_nkinkbn = "";

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
                                    break;
                                }
                            // .Vari_Bk_guid = fldvalue.Trim
                            // 取得済み
                            case "部屋No":
                                {
                                    break;
                                }
                            // .Vari_Hy_guid = fldvalue.Trim
                            // 取得済み
                            case "契約No":
                                {
                                    break;
                                }
                            // .Vari_Ky_guid = fldvalue.Trim
                            // 取得済み
                            case "契約レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "該当年月":
                                {
                                    model_cvitem.Vari_Gt_ym = fldvalue.Trim();
                                    break;
                                }
                            case "請求締め日":
                                {
                                    model_cvitem.Vari_Sq_simeymd = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目名":
                                {
                                    tmp_nkinname = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目区分":
                                {
                                    var a = fldvalue.Trim();
                                    tmp_nkinkbn = EtcMethod.Get_Nkinruiname(ref a);
                                    break;
                                }
                            case "(入金予定)入金区分":
                                {
                                    model_cvitem.Vari_Nkbn_yotei = fldvalue.Trim();
                                    break;
                                }
                            case "(入金予定)家賃入金口座":
                                {
                                    model_cvitem.Vari_Yotei_yatin_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "請求先No":
                                {
                                    model_cvitem.Vari_Sqsaki_no = fldvalue.Trim();
                                    break;
                                }
                            case "請求額":
                                {
                                    model_cvitem.Vari_Sq_gak = fldvalue.Trim();
                                    break;
                                }
                            case "請求税区分":
                                {
                                    model_cvitem.Vari_Sq_zeikbn = fldvalue.Trim();
                                    break;
                                }
                            case "請求税額":
                                {
                                    model_cvitem.Vari_Sq_zeigak = fldvalue.Trim();
                                    break;
                                }
                            case "送金額":
                                {
                                    model_cvitem.Vari_So_gak = fldvalue.Trim();
                                    break;
                                }
                            case "送金税区分":
                                {
                                    model_cvitem.Vari_So_zeikbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金税額":
                                {
                                    model_cvitem.Vari_So_zeigak = fldvalue.Trim();
                                    break;
                                }
                            case "既に送金済みフラグ":
                                {
                                    model_cvitem.Vari_Sosumi_flg = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料率":
                                {
                                    model_cvitem.Vari_Kanrigak_rit = fldvalue.Trim();
                                    break;
                                }
                            case "送金先家主No":
                                {
                                    model_cvitem.Vari_Sosaki_no = fldvalue.Trim();
                                    break;
                                }
                            case "送金先家主口座No":
                                {
                                    model_cvitem.Vari_Sokoza_no = fldvalue.Trim();
                                    break;
                                }
                            case "送金率":
                                {
                                    model_cvitem.Vari_So_rit = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法 - 該当年月":
                                {
                                    model_cvitem.Vari_Soymd_gaitoukbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法 - 締日":
                                {
                                    model_cvitem.Vari_Soymd_simekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法 - 送金月":
                                {
                                    model_cvitem.Vari_Soymd_sokintukikbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法 - 送金締日":
                                {
                                    model_cvitem.Vari_Soymd_sokinsimekbn = fldvalue.Trim();
                                    break;
                                }
                            case "適用税率":
                                {
                                    model_cvitem.Vari_Zei_rit = fldvalue.Trim();
                                    break;
                                }
                            case "請求データ生成フラグ":
                                {
                                    model_cvitem.Vari_Sqbuild_flg = fldvalue.Trim();
                                    break;
                                }
                            case "担当者No":
                                {
                                    model_cvitem.Vari_Tanto_no = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料税フラグ":
                                {
                                    model_cvitem.Vari_Kanri_taxflg = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料税込フラグ":
                                {
                                    model_cvitem.Vari_Kanritesu_zeiumukbn = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料率内税フラグ":
                                {
                                    model_cvitem.Vari_Kanri_utizeiflg = fldvalue.Trim();
                                    break;
                                }

                        }

                    }

                    // 入金項目格納
                    model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + tmp_nkinkbn;

                    // 固定値
                    model_cvitem.Vari_Unyotaino_guid = Guid.NewGuid().ToString();
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
                        // Dim tmp_bkno As String = hash_cvitem.Item("bk_guid")
                        // Dim tmp_hyno As String = hash_cvitem.Item("hy_guid")
                        // Dim tmp_kyno As String = hash_cvitem.Item("ky_guid")
                        // hash_cvitem("bk_guid") = hash_bkguid.Item(tmp_bkno)
                        // hash_cvitem("hy_guid") = hash_bkhyguid.Item(tmp_bkno & "-" & tmp_hyno)
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(tmp_bkno & "-" & tmp_hyno & "-" & tmp_kyno)
                        hash_cvitem["bk_guid"] = hash_bkguid[hash_cvitem["bk_guid"]];
                        hash_cvitem["hy_guid"] = hash_bkhyguid[hash_cvitem["hy_guid"]];
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                        }
                    }

                    // 20160624 請求関連情報 親データ無しログ出力制御修正 -add sta
                    else if (hash_log.Count == 1)
                    {
                        bool logclearflg = false;
                        foreach (var erritem in hash_log)
                        {
                            string tmp_str = Conversions.ToString(erritem.Value);
                            string[] tmp_errnaiyo = tmp_str.Split('-');
                            if ((tmp_errnaiyo[0] ?? "") == CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE)
                            {
                                chg_totalcnt = chg_totalcnt - 1;
                                logclearflg = true;
                            }
                        }
                        if (logclearflg)
                        {
                            hash_log.Clear();
                        }
                        // 20160624 請求関連情報 親データ無しログ出力制御修正 -add end

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
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = chg_totalcnt;
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg end

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebktobkguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN kydata ON " + tblname + ".ky_guid = kydata.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// guid取得→ハッシュテーブル格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="hash_guid"></param>
            /// <remarks></remarks>
            public void Get_Guid(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_guid)
            {

                string tmp_sql = "SELECT * FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);

            }

            /// <summary>
            /// キー/guidをセットにしたVIEWの作成
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                // VIEW初期化
                Drop_TmpView(sqlcnnv10, viewname);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT bk_no AS キー,bk_guid FROM bkdata ";
                            break;
                        }
                    case "物件部屋キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,hy_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋契約キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                            break;
                        }
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 作業用VIEWの削除
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="viewname"></param>
            /// <remarks></remarks>
            public void Drop_TmpView(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 請求情報

    public class Sqdata_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="syorikomok"></param>
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                // 20160530 送金データ調査での不足箇所修正 -add
                var hash_bkhykykynkinguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Sqdata_Model();                // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                // Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
                string viewnamebktobkguid = CommonModule.PRE_VIEW_NAME + "物件キー情報";
                string viewnamebkhytohyguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";
                // 20160530 送金データ調査での不足箇所修正 -add
                string viewnamebkhykytokynkinguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約入金項目キー情報";

                int tmp_socnt = 0;

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebktobkguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);
                // 20160530 送金データ調査での不足箇所修正 -add
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokynkinguid);

                // テーブル名/フィールド名セット
                string tblname = "sqdata";
                // 20160519 EXEUpdateに伴う修正 請求情報 -add (最後尾にshdirect_flg、koteishrule_guid、kojorule_taisyokbnを追加)
                string fldnamegrp = "sqmei_guid,gt_ym,nkin_no,dispnkin_name,nkbn_yotei," + "sq_gak,sq_zeikbn,sq_zeigak,sq_simeymd,multisqsaki_flg," + "sqsaki_kbn,sqsaki_no,yotei_yatin_kozano,fkae_rirekino,fkae_rirekimeino," + "fkae_sqtaisyoflg,fkae_sqselectflg,fkae_sqflg,ky_guid,ky_recno," + "tuki_kbn,nkin_recno,nkin_kbn,kynkin_guid,nextky_flg," + "hendosq_guid,hendo_kbn,kensin_ymd,bk_guid,hy_guid," + "so_gak,so_zeikbn,so_zeigak,etcsq_flg,etcsq_sorit," + "etcsq_soymd,bunkatu_guid,sh_kbn,shsaki_kbn,shsaki_no," + "shsaki_kozano,ikkatu_nkin_bango,jisya_no,tanto_no,kjsqkmk_guid," + "koteirule_guid,kojorule_flg,sq_hakkoyoteiymd,sqsort_no,tateazu_flg," + "tateazu_guid,sikikinzuiji_guid,unyotaino_guid,sokaisyu_guid,szen_kbn," + "szen_no,szen_sqno,zei_rit,edit_flg,lock_flg," + "biko,history,tatekae_siwakegak,tatekae_siwakezeigak,sq_torikomiymd," + "cvpay_no,kokyaku_bango,cvpay_sqflg,etcsq_kanritesukbn,misyu_siwake_flg," + "zatusyu_kbn,shdirect_flg,koteishrule_guid,kojorule_taisyokbn";














                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebktobkguid, ref hash_bkguid);
                Get_Guid(sqlcnnv10, viewnamebkhytohyguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);
                // 20160530 送金データ調査での不足箇所修正 -add
                Get_Guid(sqlcnnv10, viewnamebkhykytokynkinguid, ref hash_bkhykykynkinguid);

                // 親マスタ取得
                // 20160601 請求データ取得処理の修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Get_BaseKey(sqlcnnv10, viewnamebktobkguid, ref list_basekeydata);
                // 20160601 請求データ取得処理の修正 -chg end

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

                // 送金情報移行用モデル
                var model_sqsocvitem = new Model.Sqdata_so_Model();

                // 20160601 請求データ取得処理の修正 -add sta
                var model_sqso_jisyacvitem = new Model.Sqdata_so_Model();
                var model_sqkesi_jisyacvitem = new Model.Sqdata_sqkesi_Model();
                // 20160601 請求データ取得処理の修正 -add end

                // 20160530 送金データ調査での不足箇所修正 -add
                // 送金管理情報移行用モデル
                var model_sqsokanricvitem = new Model.Sokanri_Model();

                // 20160601 請求データ取得処理の修正 -add sta
                // sqdata_so、sokanri、sqdata_sqkesiの初期化 (azukanriは過剰金情報移行時に初期化済み)
                int tmp_rowcnt = 0;
                var list_sqtable = new List<string>() { "sqdata_so", "sqdata_sqkesi", "sokanri" };
                foreach (var inittablename in list_sqtable)
                {
                    string tmp_sqldelete = DBQuery.Qry_DelInfo(inittablename);
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldelete, ref tmp_rowcnt);
                }
                // 20160601 請求データ取得処理の修正 -add end

                // 20160624 請求関連情報 親データ無しログ出力制御修正 -add
                int chg_totalcnt = rowcnt;        // 親データ無しのデータ件数調整用(カウントしないため)

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
                    model_cvitem.Vari_Bk_guid = tmp_keymain;
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



                    // 作業用変数
                    string tmp_nkinname = "";
                    string tmp_nkinkbn = "";
                    // 20160601 請求データ取得処理の修正 -add sta
                    string tmp_sono = "";         // 送金ルール送金先情報取得用
                    string tmp_sokozano = "";     // 送金ルール送金先情報取得用
                    string sqcvkbn = "";          // 請求データの移行処理を判別する作業用変数
                                                  // 20160601 請求データ取得処理の修正 -add end

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
                                    break;
                                }
                            // .Vari_Bk_guid = fldvalue.Trim
                            // 取得済み
                            case "部屋No":
                                {
                                    break;
                                }
                            // .Vari_Hy_guid = fldvalue.Trim
                            // 取得済み
                            case "契約No":
                                {
                                    break;
                                }
                            // .Vari_Ky_guid = fldvalue.Trim
                            // 取得済み
                            case "契約レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "該当年月":
                                {
                                    model_cvitem.Vari_Gt_ym = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目名":
                                {
                                    tmp_nkinname = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目区分":
                                {
                                    model_cvitem.Vari_Nkin_kbn = fldvalue.Trim();
                                    break;
                                }
                            // 20160601 請求データ取得処理の修正 -del sta
                            // Dim tmp_str As String = IIf(fldvalue.Trim = "", "6", fldvalue.Trim) 'その他請求は入金区分が空なので「6」をセットして紐付ける
                            // tmp_nkinkbn = EtcMethod.Get_Nkinruiname(tmp_str)
                            // 20160601 請求データ取得処理の修正 -del end
                            case "入金項目表示用名称":
                                {
                                    model_cvitem.Vari_Dispnkin_name = fldvalue.Trim();
                                    break;
                                }
                            case "(入金予定)入金区分":
                                {
                                    model_cvitem.Vari_Nkbn_yotei = fldvalue.Trim();
                                    break;
                                }
                            case "請求額":
                                {
                                    model_cvitem.Vari_Sq_gak = fldvalue.Trim();
                                    // 20160601 請求データ取得処理の修正 -add sta
                                    model_sqkesi_jisyacvitem.Vari_Kesi_gak = fldvalue.Trim();
                                    model_sqsocvitem.Vari_Calc_sqgak = fldvalue.Trim();
                                    break;
                                }
                            // 20160601 請求データ取得処理の修正 -add end
                            case "請求税区分":
                                {
                                    model_cvitem.Vari_Sq_zeikbn = fldvalue.Trim();
                                    break;
                                }
                            case "請求税額":
                                {
                                    model_cvitem.Vari_Sq_zeigak = fldvalue.Trim();
                                    // 20160601 請求データ取得処理の修正 -add
                                    model_sqkesi_jisyacvitem.Vari_Kesi_zeigak = fldvalue.Trim();
                                    break;
                                }
                            case "請求締切日":
                                {
                                    model_cvitem.Vari_Sq_simeymd = fldvalue.Trim();
                                    break;
                                }
                            case "複数請求先フラグ":
                                {
                                    model_cvitem.Vari_Multisqsaki_flg = fldvalue.Trim();
                                    break;
                                }
                            case "請求先区分":
                                {
                                    model_cvitem.Vari_Sqsaki_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "請求先No":
                                {
                                    model_cvitem.Vari_Sqsaki_no = fldvalue.Trim();
                                    break;
                                }
                            case "(入金予定)家賃振込先口座No":
                                {
                                    model_cvitem.Vari_Yotei_yatin_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "口座振替履歴No":
                                {
                                    model_cvitem.Vari_Fkae_rirekino = fldvalue.Trim();
                                    break;
                                }
                            case "口座振替履歴明細No":
                                {
                                    model_cvitem.Vari_Fkae_rirekimeino = fldvalue.Trim();
                                    break;
                                }
                            case "振替請求対象フラグ 1:請求中　2:未請求":
                                {
                                    model_cvitem.Vari_Fkae_sqtaisyoflg = fldvalue.Trim();
                                    break;
                                }
                            case "振替請求対象フラグ　1:対象　2:対象外":
                                {
                                    model_cvitem.Vari_Fkae_sqselectflg = fldvalue.Trim();
                                    break;
                                }
                            case "振替請求中フラグ":
                                {
                                    model_cvitem.Vari_Fkae_sqflg = fldvalue.Trim();
                                    break;
                                }
                            case "月区分":
                                {
                                    model_cvitem.Vari_Tuki_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目明細No":
                                {
                                    model_cvitem.Vari_Nkin_recno = fldvalue.Trim();
                                    break;
                                }
                            case "次回契約情報フラグ":
                                {
                                    model_cvitem.Vari_Nextky_flg = fldvalue.Trim();
                                    break;
                                }
                            case "変動費請求GUID":
                                {
                                    model_cvitem.Vari_Hendosq_guid = fldvalue.Trim();
                                    break;
                                }
                            case "変動費区分":
                                {
                                    model_cvitem.Vari_Hendo_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "変動費検針日":
                                {
                                    model_cvitem.Vari_Kensin_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "送金額":
                                {
                                    model_cvitem.Vari_So_gak = fldvalue.Trim();
                                    // 20160601 請求データ取得処理の修正 -add sta
                                    model_sqsocvitem.Vari_So_gak = fldvalue.Trim();
                                    model_sqso_jisyacvitem.Vari_So_gak = fldvalue.Trim();
                                    break;
                                }
                            // 20160601 請求データ取得処理の修正 -add end
                            case "送金税区分":
                                {
                                    model_cvitem.Vari_So_zeikbn = fldvalue.Trim();
                                    // 20160601 請求データ取得処理の修正 -add sta
                                    model_sqsocvitem.Vari_So_zeikbn = fldvalue.Trim();
                                    model_sqso_jisyacvitem.Vari_So_zeikbn = fldvalue.Trim();
                                    break;
                                }
                            // 20160601 請求データ取得処理の修正 -add end
                            case "送金税額":
                                {
                                    model_cvitem.Vari_So_zeigak = fldvalue.Trim();
                                    // 20160601 請求データ取得処理の修正 -add sta
                                    model_sqsocvitem.Vari_So_zeigak = fldvalue.Trim();
                                    model_sqso_jisyacvitem.Vari_So_zeigak = fldvalue.Trim();
                                    break;
                                }
                            // 20160601 請求データ取得処理の修正 -add end
                            case "その他請求フラグ":
                                {
                                    model_cvitem.Vari_Etcsq_flg = fldvalue.Trim();
                                    break;
                                }
                            case "その他請求の送金率":
                                {
                                    model_cvitem.Vari_Etcsq_sorit = fldvalue.Trim();
                                    break;
                                }
                            case "その他請求の指定送金日":
                                {
                                    model_cvitem.Vari_Etcsq_soymd = fldvalue.Trim();
                                    break;
                                }
                            case "分割GUID":
                                {
                                    model_cvitem.Vari_Bunkatu_guid = fldvalue.Trim();
                                    break;
                                }
                            case "支払・返金区分":
                                {
                                    model_cvitem.Vari_Sh_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "支払・返金先区分":
                                {
                                    model_cvitem.Vari_Shsaki_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "支払・返金先No":
                                {
                                    model_cvitem.Vari_Shsaki_no = fldvalue.Trim();
                                    break;
                                }
                            case "支払・返金先口座No":
                                {
                                    model_cvitem.Vari_Shsaki_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)":
                                {
                                    model_cvitem.Vari_Ikkatu_nkin_bango = fldvalue.Trim();
                                    break;
                                }
                            case "支店No":
                                {
                                    model_cvitem.Vari_Jisya_no = fldvalue.Trim();
                                    break;
                                }
                            case "担当者No":
                                {
                                    model_cvitem.Vari_Tanto_no = fldvalue.Trim();
                                    break;
                                }
                            case "控除請求項目GUID":
                                {
                                    model_cvitem.Vari_Kjsqkmk_guid = fldvalue.Trim();
                                    break;
                                }
                            case "固定控除ルールGUID(請求)":
                                {
                                    model_cvitem.Vari_Koteirule_guid = fldvalue.Trim();
                                    break;
                                }
                            case "控除ルールフラグ":
                                {
                                    model_cvitem.Vari_Kojorule_flg = fldvalue.Trim();
                                    break;
                                }
                            case "請求書発行予定日":
                                {
                                    model_cvitem.Vari_Sq_hakkoyoteiymd = fldvalue.Trim();
                                    break;
                                }
                            case "請求ソートNo":
                                {
                                    model_cvitem.Vari_Sqsort_no = fldvalue.Trim();
                                    break;
                                }
                            case "立替・預りフラグ":
                                {
                                    model_cvitem.Vari_Tateazu_flg = fldvalue.Trim();
                                    break;
                                }
                            case "立替・預り設定時の預り予定GUID（預り金データ）":
                                {
                                    model_cvitem.Vari_Tateazu_guid = fldvalue.Trim();
                                    break;
                                }
                            case "敷金保証金随時処理GUID":
                                {
                                    model_cvitem.Vari_Sikikinzuiji_guid = fldvalue.Trim();
                                    break;
                                }
                            case "運用開始時の未納・滞納金GUID":
                                {
                                    model_cvitem.Vari_Unyotaino_guid = fldvalue.Trim();
                                    break;
                                }
                            case "送金回収GUID":
                                {
                                    model_cvitem.Vari_Sokaisyu_guid = fldvalue.Trim();
                                    break;
                                }
                            case "修繕区分":
                                {
                                    model_cvitem.Vari_Szen_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "随時修繕No":
                                {
                                    model_cvitem.Vari_Szen_no = fldvalue.Trim();
                                    break;
                                }
                            case "修繕請求No":
                                {
                                    model_cvitem.Vari_Szen_sqno = fldvalue.Trim();
                                    break;
                                }
                            case "適用税率":
                                {
                                    model_cvitem.Vari_Zei_rit = fldvalue.Trim();
                                    break;
                                }
                            case "編集フラグ":
                                {
                                    model_cvitem.Vari_Edit_flg = fldvalue.Trim();
                                    break;
                                }
                            case "更新ロックフラグ":
                                {
                                    model_cvitem.Vari_Lock_flg = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "立替仕訳額":
                                {
                                    model_cvitem.Vari_Tatekae_siwakegak = fldvalue.Trim();
                                    break;
                                }
                            case "立替仕訳税額":
                                {
                                    model_cvitem.Vari_Tatekae_siwakezeigak = fldvalue.Trim();
                                    break;
                                }
                            case "取込日（インポート実施日）":
                                {
                                    model_cvitem.Vari_Sq_torikomiymd = fldvalue.Trim();
                                    break;
                                }
                            case "コンビニ収納サービスNo":
                                {
                                    model_cvitem.Vari_Cvpay_no = fldvalue.Trim();
                                    break;
                                }
                            case "顧客番号":
                                {
                                    model_cvitem.Vari_Kokyaku_bango = fldvalue.Trim();
                                    break;
                                }
                            case "コンビニ収納請求フラグ":
                                {
                                    model_cvitem.Vari_Cvpay_sqflg = fldvalue.Trim();
                                    break;
                                }
                            case "その他請求管理手数料区分":
                                {
                                    model_cvitem.Vari_Etcsq_kanritesukbn = fldvalue.Trim();
                                    break;
                                }
                            case "未収仕訳フラグ":
                                {
                                    model_cvitem.Vari_Misyu_siwake_flg = fldvalue.Trim();
                                    break;
                                }
                            case "雑収入区分":
                                {
                                    model_cvitem.Vari_Zatusyu_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金確定日":
                                {
                                    model_sqsocvitem.Vari_So_yoteiymd = fldvalue.Trim();
                                    model_sqso_jisyacvitem.Vari_So_yoteiymd = fldvalue.Trim();
                                    model_sqkesi_jisyacvitem.Vari_Kesi_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "滞納保証フラグ":
                                {
                                    model_sqsocvitem.Vari_Hosyo_flg = fldvalue.Trim();     // 送金情報移行用
                                    break;
                                }
                            // 20160530 送金データ調査での不足箇所修正 -add sta
                            case "支払返金先No":
                                {
                                    model_sqsokanricvitem.Vari_Sosaki_no = fldvalue.Trim();   // 送金管理情報移行用
                                    tmp_sono = fldvalue.Trim();    // 20160601 請求データ取得処理の修正 -add
                                    model_cvitem.Vari_Shsaki_no = fldvalue.Trim(); // 20160601 請求データ取得処理の修正 -add
                                    break;
                                }
                            case "支払返金先口座No":
                                {
                                    model_sqsokanricvitem.Vari_Sokoza_no = fldvalue.Trim();     // 送金管理情報移行用
                                    tmp_sokozano = fldvalue.Trim();    // 20160601 請求データ取得処理の修正 -add
                                    model_cvitem.Vari_Shsaki_kozano = fldvalue.Trim(); // 20160601 請求データ取得処理の修正 -add
                                    break;
                                }
                            // 20160530 送金データ調査での不足箇所修正 -add end

                            // 20160601 請求データ取得処理の修正 -add sta
                            case "請求区分":
                                {
                                    sqcvkbn = fldvalue.Trim();
                                    break;
                                }
                            case "支払返金区分":
                                {
                                    model_cvitem.Vari_Sh_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "支払返金先区分":
                                {
                                    model_cvitem.Vari_Shsaki_kbn = fldvalue.Trim();
                                    break;
                                }
                                // 20160601 請求データ取得処理の修正 -add end
                        }

                    }

                    // 入金項目格納
                    // 20160601 請求データ取得処理の修正 -chg sta
                    // .Vari_Nkin_no = tmp_nkinname & "-" & tmp_nkinkbn
                    switch (sqcvkbn ?? "")
                    {
                        case "その他請求":
                        case "送金確定_その他請求":
                            {
                                string argtukikbn = "6";
                                tmp_nkinkbn = EtcMethod.Get_Nkinruiname(ref argtukikbn);
                                break;
                            }
                        case "控除支払の請求":
                        case "自社支払":
                            {
                                string argtukikbn1 = "9";
                                tmp_nkinkbn = EtcMethod.Get_Nkinruiname(ref argtukikbn1);
                                break;
                            }

                        default:
                            {
                                string argtukikbn2 = model_cvitem.Vari_Nkin_kbn;
                                tmp_nkinkbn = EtcMethod.Get_Nkinruiname(ref argtukikbn2);
                                model_cvitem.Vari_Nkin_kbn = argtukikbn2;
                                break;
                            }
                    }
                    model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + tmp_nkinkbn;
                    // 20160601 請求データ取得処理の修正 -chg end

                    // 固定値
                    // 20160601 請求データ取得処理の修正 -chg sta
                    // .Vari_Sqmei_guid = Guid.NewGuid.ToString
                    string tmp_sqmeiguid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Sqmei_guid = tmp_sqmeiguid;
                    model_sqsocvitem.Vari_Sqmei_guid = tmp_sqmeiguid;
                    model_sqso_jisyacvitem.Vari_Sqmei_guid = tmp_sqmeiguid;
                    model_sqkesi_jisyacvitem.Vari_Sqmei_guid = tmp_sqmeiguid;
                    model_cvitem.Vari_Lock_flg = 1.ToString();  // ロックフラグをプログラム内で立てておく(抽出クエリでは受領データの更新ロックフラグを元に抽出している)
                                                                // 20160601 請求データ取得処理の修正 -chg end
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
                        hash_cvitem["bk_guid"] = hash_bkguid[hash_cvitem["bk_guid"]];
                        hash_cvitem["hy_guid"] = hash_bkhyguid[hash_cvitem["hy_guid"]];
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160530 送金データ調査での不足箇所修正 -add sta
                        // 契約入金項目guid取得用のキーを生成
                        // 変換した入金項目Noを取得
                        string tmp_nkinno = Conversions.ToString(hash_cvitem["nkin_no"]);
                        string tmp_key_kynkin_guid = model_cvitem.Vari_Ky_guid + "-" + model_cvitem.Vari_Ky_recno + "-" + tmp_nkinno;
                        hash_cvitem["kynkin_guid"] = hash_bkhykykynkinguid[tmp_key_kynkin_guid];
                        // 20160530 送金データ調査での不足箇所修正 -add end

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                            // 20160601 請求データ取得処理の修正 -chg sta
                            // '送金情報へ反映
                            // If model_sqsocvitem.Vari_So_yoteiymd <> "" Then
                            // '20160530 送金データ調査での不足箇所修正 -add
                            // Dim hash_sqso As New SafeDictionary<string, string>      '移行する送金データから送金管理データを作成するための格納用
                            // tmp_socnt = tmp_socnt + 1
                            // Dim sotblinitflg As Boolean = IIf(tmp_socnt = 1, True, False)
                            // Call Me.Set_SqdataSo(sqlcnnv10, fldnamegrp, hash_cvitem, model_sqsocvitem, hash_sqso, sotblinitflg)
                            // '20160530 送金データ調査での不足箇所修正 -add
                            // Call Me.Set_SqdataSoKanri(sqlcnnv10, fldnamegrp, hash_cvitem, hash_sqso, model_sqsokanricvitem, sotblinitflg)
                            // End If

                            // sqdata以外に作成するテーブルに関して、送金確定日の有無で判別していた処理を新規に設けた「請求区分」から判別するように修正
                            switch (sqcvkbn ?? "")     // sqdataまで作成(ここで終わり)
                            {
                                case "部分入金":
                                case "控除支払の請求":
                                case "その他請求":
                                    {
                                        break;
                                    }
                                // 処理無し
                                case "送金確定_その他請求":
                                case "送金確定_その他請求以外":     // sqdata_so、sokanriを作成
                                    {
                                        string tmp_bkguid = Conversions.ToString(hash_cvitem["bk_guid"]);
                                        Set_SqdataSo(sqlcnnv10, model_sqsocvitem, tmp_bkguid, tmp_sono, tmp_sokozano);           // sqdata_so作成モジュール
                                        break;
                                    }
                                case "自社支払":
                                    {
                                        // sqdata_sqkesi、sqdata_so、azukanriを作成
                                        Set_SqdataSo(sqlcnnv10, model_sqso_jisyacvitem);             // sqdata_so作成モジュール
                                        Set_Sqdatakesi(sqlcnnv10, model_sqkesi_jisyacvitem);         // sqdata_sqkesi作成モジュール (azukanriを含む)
                                        break;
                                    }
                            }
                            // 20160601 請求データ取得処理の修正 -chg end
                        }
                    }

                    // 20160624 請求関連情報 親データ無しログ出力制御修正 -add sta
                    else if (hash_log.Count == 1)
                    {
                        bool logclearflg = false;
                        foreach (var erritem in hash_log)
                        {
                            string tmp_str = Conversions.ToString(erritem.Value);
                            string[] tmp_errnaiyo = tmp_str.Split('-');
                            if ((tmp_errnaiyo[0] ?? "") == CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE)
                            {
                                chg_totalcnt = chg_totalcnt - 1;
                                logclearflg = true;
                            }
                        }
                        if (logclearflg)
                        {
                            hash_log.Clear();
                        }
                        // 20160624 請求関連情報 親データ無しログ出力制御修正 -add end

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

                // 20160601 請求データ取得処理の修正 -add sta
                // 送金管理情報の作成
                Set_SqdataSoKanri(sqlcnnv10);        // sokanri作成モジュール
                // 20160601 請求データ取得処理の修正 -add end

                // ************************
                // 終了処理
                // ************************

                // 中間ファイル件数を取得
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = chg_totalcnt;
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg end

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebktobkguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
                // 20160530 送金データ調査での不足箇所修正 -add
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN kydata ON " + tblname + ".ky_guid = kydata.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// guid取得→ハッシュテーブル格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="hash_guid"></param>
            /// <remarks></remarks>
            public void Get_Guid(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_guid)
            {

                string tmp_sql = "SELECT * FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);

            }

            /// <summary>
            /// キー/guidをセットにしたVIEWの作成
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                // VIEW初期化
                Drop_TmpView(sqlcnnv10, viewname);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT bk_no AS キー,bk_guid FROM bkdata ";
                            break;
                        }
                    case "物件部屋キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,hy_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋契約キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    // 20160530 送金データ調査での不足箇所修正 -add sta
                    case "物件部屋契約入金項目キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) + '-' +  ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,KYN.ky_recno)  + '-' + CONVERT(varchar,KYN.nkin_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,nkin_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM kydata_nkin AS KYN ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN kydata AS KY ON KYN.ky_guid = KY.ky_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                            tmp_sql_create = tmp_sql_create + " WHERE KYN.tuki_kbn = 2 ";
                            break;
                        }
                        // 20160530 送金データ調査での不足箇所修正 -add end
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 作業用VIEWの削除
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="viewname"></param>
            /// <remarks></remarks>
            public void Drop_TmpView(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

            }

            // 20160601 請求データ取得処理の修正 -del sta
            // ''' <summary>
            // ''' 請求情報のうち送金確定されているデータを送金予定情報へ反映
            // ''' </summary>
            // ''' <param name="sqlcnnv10"></param>
            // ''' <param name="azfldnamegrp"></param>
            // ''' <param name="hash_sqcvitem"></param>
            // ''' <param name="model_sqsocvitem"></param>
            // ''' <remarks></remarks>
            // Public Sub Set_SqdataSo(ByVal sqlcnnv10 As SqlConnection, ByVal azfldnamegrp As String, ByVal hash_sqcvitem As SafeDictionary<string, string>, ByVal model_sqsocvitem As Object, ByRef hash_sqsocvitem As SafeDictionary<string, string>, ByVal sqsotblinitflg As Boolean)

            // Dim normalflg As Boolean = True
            // '20160530 送金データ調査での不足箇所修正 -del
            // '引数へ移動
            // 'Dim hash_sqsocvitem As New SafeDictionary<string, string>

            // '物件guidと送金ルール情報を紐付けた値の格納用
            // Dim hash_bktosoruleguid As New SafeDictionary<string, string>
            // Dim hash_bktosoruleno As New SafeDictionary<string, string>
            // Dim hash_bktosorulesoguid As New SafeDictionary<string, string>
            // Dim hash_bktosorulesono As New SafeDictionary<string, string>

            // Dim sqsotblname As String = "sqdata_so"
            // Dim sqsofldnamegrp As String = "sqmei_guid,so_recno,so_gak,so_zeikbn,so_zeigak," & _
            // "edit_flg,lock_flg,hosyo_flg,so_yoteiymd,so_kakuymd," & _
            // "sokotik_guid,sorule_guid,sorule_no,sosaki_sorule_guid,sosaki_sorule_no," & _
            // "sokin_rit,calc_sqgak,sokaisyu_kbn,sokaisyu_guid,siwake_flg," & _
            // "history"

            // 'テーブル初期化処理
            // If sqsotblinitflg Then
            // '初期化処理実行
            // Dim rowcnt As Integer = 0
            // 'Dim qrywhere As String = " stakeholder_kbn = 100 "
            // 'Dim tmp_sqldelete As String = DBQuery.Qry_DelInfo(sqsotblname, qrywhere)
            // Dim tmp_sqldelete As String = DBQuery.Qry_DelInfo(sqsotblname)
            // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldelete, rowcnt)
            // End If

            // '物件guidと送金ルール情報を紐付けた値を取得
            // Dim soymd As String = model_sqsocvitem.Vari_So_yoteiymd     '送金ルール取得用に送金確定日を取得
            // soymd = soymd.Replace("-", "")
            // soymd = soymd.Replace(" 00:00:00.000", "")                  '成形 (念の為)
            // hash_bktosoruleguid = Me.Get_BkToSoruleInfo(sqlcnnv10, soymd, 1)
            // hash_bktosoruleno = Me.Get_BkToSoruleInfo(sqlcnnv10, soymd, 2)
            // hash_bktosorulesoguid = Me.Get_BkToSoruleInfo(sqlcnnv10, soymd, 3)
            // hash_bktosorulesono = Me.Get_BkToSoruleInfo(sqlcnnv10, soymd, 4)

            // '作業用変数
            // Dim tmp_sqgak As String = ""
            // Dim tmp_sogak As String = ""
            // Dim tmp_etcsqflg As String = ""
            // Dim tmp_etcsqsorit As String = ""
            // Dim tmp_bkguid As String = ""

            // '請求情報から送金予定情報へ
            // For Each sqitem In hash_sqcvitem

            // Dim sqfldname As String = sqitem.Key
            // Dim sqvalue As String = sqitem.Value

            // Select Case sqfldname
            // Case "bk_guid"
            // tmp_bkguid = sqvalue

            // Case "sqmei_guid"
            // hash_sqsocvitem.Add("sqmei_guid", sqvalue)
            // Case "so_gak"
            // hash_sqsocvitem.Add("so_gak", sqvalue)
            // hash_sqsocvitem.Add("calc_sqgak", sqvalue)
            // tmp_sogak = sqvalue
            // Case "so_zeikbn"
            // hash_sqsocvitem.Add("so_zeikbn", sqvalue)
            // Case "so_zeigak"
            // hash_sqsocvitem.Add("so_zeigak", sqvalue)
            // Case "sq_gak"
            // tmp_sqgak = sqvalue         '送金率算出用
            // Case "etcsq_flg"
            // tmp_etcsqflg = sqvalue      '送金率算出用
            // Case "etcsq_sorit"
            // tmp_etcsqsorit = sqvalue    '送金率算出用
            // Case "history"
            // hash_sqsocvitem.Add("history", sqvalue)
            // End Select

            // Next

            // '送金率を取得
            // If tmp_sqgak = "0" Then
            // hash_sqsocvitem.Add("sokin_rit", "0")
            // ElseIf tmp_etcsqflg = "1" Then
            // hash_sqsocvitem.Add("sokin_rit", tmp_etcsqsorit)
            // ElseIf tmp_etcsqflg = "0" Then
            // Dim tmp_db As Double = 0
            // Dim tmp_dbsqgak As Double = IIf(Double.TryParse(tmp_sqgak, tmp_db) = False, tmp_db, Double.Parse(tmp_sqgak))
            // Dim tmp_dbsogak As Double = IIf(Double.TryParse(tmp_sogak, tmp_db) = False, tmp_db, Double.Parse(tmp_sogak))
            // hash_sqsocvitem.Add("sokin_rit", (tmp_dbsogak / tmp_dbsqgak * 100).ToString)
            // End If

            // '予め取得しておいた送金情報を設定
            // hash_sqsocvitem.Add("so_yoteiymd", model_sqsocvitem.Vari_So_yoteiymd)   'V7送金確定日→10金予定日
            // hash_sqsocvitem.Add("hosyo_flg", model_sqsocvitem.Vari_Hosyo_flg)       '滞納保証フラグ

            // '送金ルールguid、送金ルールNoを設定
            // hash_sqsocvitem.Add("sorule_guid", hash_bktosoruleguid.Item(tmp_bkguid))
            // hash_sqsocvitem.Add("sorule_no", hash_bktosoruleno.Item(tmp_bkguid))
            // hash_sqsocvitem.Add("sosaki_sorule_guid", hash_bktosorulesoguid.Item(tmp_bkguid))
            // hash_sqsocvitem.Add("sosaki_sorule_no", hash_bktosorulesono.Item(tmp_bkguid))

            // '固定値
            // hash_sqsocvitem.Add("so_recno", "1")
            // hash_sqsocvitem.Add("edit_flg", "")
            // hash_sqsocvitem.Add("lock_flg", "1")
            // hash_sqsocvitem.Add("so_kakuymd", "")
            // '20160530 送金データ調査での不足箇所修正 -chg sta
            // 'hash_sqsocvitem.Add("sokotik_guid", Guid.NewGuid.ToString)
            // hash_sqsocvitem.Add("sokotik_guid", "")
            // '20160530 送金データ調査での不足箇所修正 -chg end
            // hash_sqsocvitem.Add("sokaisyu_kbn", "")
            // hash_sqsocvitem.Add("sokaisyu_guid", "")
            // hash_sqsocvitem.Add("siwake_flg", "")

            // '挿入処理
            // Call CVDBInsert.Cnv_Db(sqlcnnv10, sqsotblname, sqsofldnamegrp, hash_sqsocvitem, normalflg)

            // End Sub

            // ''' <summary>
            // ''' 物件guidと送金ルール情報を紐付けたハッシュテーブルを作成
            // ''' </summary>
            // ''' <param name="sqlcnnv10"></param>
            // ''' <param name="typeno"></param>
            // ''' <remarks></remarks>
            // Private Function Get_BkToSoruleInfo(ByVal sqlcnnv10 As SqlConnection, ByVal soymd As String, ByVal typeno As Integer) As SafeDictionary<string, string>

            // Dim rtn_hash As New SafeDictionary<string, string>
            // Dim tmp_sql As String = ""

            // Select Case typeno
            // Case 1
            // tmp_sql = " SELECT bk_guid,sorule_guid FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid "
            // tmp_sql = tmp_sql & " WHERE sorule_startymd <= '" & soymd & "' AND ISNULL(sorule_endymd,'21001231') >=  '" & soymd & "'"
            // Case 2
            // tmp_sql = "SELECT bk_guid,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid"
            // tmp_sql = tmp_sql & " WHERE sorule_startymd <= '" & soymd & "' AND ISNULL(sorule_endymd,'21001231') >=  '" & soymd & "'"
            // Case 3
            // tmp_sql = tmp_sql & " SELECT bk_guid,SOSAKI.sorule_guid FROM "
            // tmp_sql = tmp_sql & " ( "
            // tmp_sql = tmp_sql & " 	SELECT bk_guid,sorule_guid,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid "
            // tmp_sql = tmp_sql & " 	WHERE sorule_startymd <= '" & soymd & "' AND ISNULL(sorule_endymd,'21001231') >=  '" & soymd & "'"
            // tmp_sql = tmp_sql & " ) AS SO "
            // tmp_sql = tmp_sql & " LEFT JOIN "
            // tmp_sql = tmp_sql & " 	( "
            // tmp_sql = tmp_sql & " 		SELECT DISTINCT sorule_guid,sorule_no FROM sorule_sosaki "
            // tmp_sql = tmp_sql & " 	) AS SOSAKI "
            // tmp_sql = tmp_sql & " ON  SO.sorule_guid = SOSAKI.sorule_guid "
            // tmp_sql = tmp_sql & " AND SO.sorule_no = SOSAKI.sorule_no "
            // Case 4
            // tmp_sql = tmp_sql & " SELECT bk_guid,SOSAKI.sorule_no FROM "
            // tmp_sql = tmp_sql & " ( "
            // tmp_sql = tmp_sql & " 	SELECT bk_guid,sorule_guid,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid "
            // tmp_sql = tmp_sql & " 	WHERE sorule_startymd <= '" & soymd & "' AND ISNULL(sorule_endymd,'21001231') >=  '" & soymd & "'"
            // tmp_sql = tmp_sql & " ) AS SO "
            // tmp_sql = tmp_sql & " LEFT JOIN "
            // tmp_sql = tmp_sql & " 	( "
            // tmp_sql = tmp_sql & " 		SELECT DISTINCT sorule_guid,sorule_no FROM sorule_sosaki "
            // tmp_sql = tmp_sql & " 	) AS SOSAKI "
            // tmp_sql = tmp_sql & " ON  SO.sorule_guid = SOSAKI.sorule_guid "
            // tmp_sql = tmp_sql & " AND SO.sorule_no = SOSAKI.sorule_no "
            // End Select

            // DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, rtn_hash)

            // Return rtn_hash

            // End Function
            // 20160601 請求データ取得処理の修正 -del end

            /// <summary>
            /// 請求情報のうち送金確定されているデータを送金予定情報へ反映 '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="model_sqsocvitem"></param>
            /// <remarks></remarks>
            public void Set_SqdataSo(SqlConnection sqlcnnv10, Model.Sqdata_so_Model model_sqsocvitem, string bkguid, string sono, string sokozano)
            {

                bool normalflg = true;

                string tblname = "sqdata_so";
                string fldnamegrp = "sqmei_guid,so_recno,so_gak,so_zeikbn,so_zeigak," + "edit_flg,lock_flg,hosyo_flg,so_yoteiymd,so_kakuymd," + "sokotik_guid,sorule_guid,sorule_no,sosaki_sorule_guid,sosaki_sorule_no," + "sokin_rit,calc_sqgak,sokaisyu_kbn,sokaisyu_guid,siwake_flg," + "history";





                // 送金ルールguidの取得
                string tmp_soruleguid = "";
                string tmp_soruleno = "";
                string tmp_sorulesosakiguid = "";
                string tmp_sorulesosakino = "";

                Set_SoruleInfo(sqlcnnv10, bkguid, ref tmp_soruleguid, ref tmp_soruleno);
                Set_SoruleSosakiInfo(sqlcnnv10, bkguid, sono, sokozano, ref tmp_sorulesosakiguid, ref tmp_sorulesosakino);

                // モデルへ値をセット
                model_sqsocvitem.Vari_Sorule_guid = tmp_soruleguid;
                model_sqsocvitem.Vari_Sorule_no = tmp_soruleno;
                model_sqsocvitem.Vari_Sosaki_sorule_guid = tmp_sorulesosakiguid;
                model_sqsocvitem.Vari_Sosaki_sorule_no = tmp_sorulesosakino;

                // 送金率を取得
                // 20160603 ユーザーデータ検証による修正 -chg sta
                // If .Vari_Calc_sqgak = "0" Then
                // .Vari_Sokin_rit = 0
                // Else
                // Dim tmp_db As Double = 0
                // Dim tmp_dbsqgak As Double = IIf(Double.TryParse(.Vari_Calc_sqgak, tmp_db) = False, tmp_db, Double.Parse(.Vari_Calc_sqgak))
                // Dim tmp_dbsogak As Double = IIf(Double.TryParse(.Vari_So_gak, tmp_db) = False, tmp_db, Double.Parse(.Vari_So_gak))
                // .Vari_Sokin_rit = ((tmp_dbsogak / tmp_dbsqgak) * 100).ToString
                // End If
                double tmp_db_sq = 0d;
                double tmp_dbsqgak = Conversions.ToDouble(Interaction.IIf(double.TryParse(model_sqsocvitem.Vari_Calc_sqgak, out tmp_db_sq) == false, tmp_db_sq, double.Parse(model_sqsocvitem.Vari_Calc_sqgak)));
                if (tmp_dbsqgak == 0d)
                {
                    model_sqsocvitem.Vari_Sokin_rit = 0.ToString();
                }
                else
                {
                    double tmp_db_so = 0d;
                    double tmp_dbsogak = Conversions.ToDouble(Interaction.IIf(double.TryParse(model_sqsocvitem.Vari_So_gak, out tmp_db_so) == false, tmp_db_so, double.Parse(model_sqsocvitem.Vari_So_gak)));
                    model_sqsocvitem.Vari_Sokin_rit = (tmp_dbsogak / tmp_dbsqgak * 100d).ToString();
                }
                // 20160603 ユーザーデータ検証による修正 -chg end

                // 固定値
                model_sqsocvitem.Vari_So_recno = 1.ToString();
                model_sqsocvitem.Vari_Edit_flg = 1.ToString();
                model_sqsocvitem.Vari_Lock_flg = 1.ToString();
                model_sqsocvitem.Vari_So_kakuymd = "";
                model_sqsocvitem.Vari_Sokotik_guid = "";
                model_sqsocvitem.Vari_Sokaisyu_kbn = "";
                model_sqsocvitem.Vari_Sokaisyu_guid = "";
                model_sqsocvitem.Vari_Siwake_flg = "";

                model_sqsocvitem.Vari_History = CommonModule.DefHistory;

                // 移行テーブルのフィールド名と移行値を紐付けたハッシュテーブルを作成
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_sqsocvitem);

                // 挿入処理
                CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

            }

            /// <summary>
            /// 請求情報のうち送金確定されているデータを送金予定情報へ反映(自社支払データ) '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            public void Set_SqdataSo(SqlConnection sqlcnnv10, Model.Sqdata_so_Model model_sqso_jisyacvitem)
            {

                bool normalflg = true;

                string tblname = "sqdata_so";
                string fldnamegrp = "sqmei_guid,so_recno,so_gak,so_zeikbn,so_zeigak," + "edit_flg,lock_flg,hosyo_flg,so_yoteiymd,so_kakuymd," + "sokotik_guid,sorule_guid,sorule_no,sosaki_sorule_guid,sosaki_sorule_no," + "sokin_rit,calc_sqgak,sokaisyu_kbn,sokaisyu_guid,siwake_flg," + "history";





                // 固定値セット
                model_sqso_jisyacvitem.Vari_So_recno = 1.ToString();
                model_sqso_jisyacvitem.Vari_Edit_flg = "";
                model_sqso_jisyacvitem.Vari_Lock_flg = 1.ToString();
                model_sqso_jisyacvitem.Vari_Hosyo_flg = 0.ToString();
                model_sqso_jisyacvitem.Vari_So_kakuymd = "";
                model_sqso_jisyacvitem.Vari_Sorule_guid = "";
                model_sqso_jisyacvitem.Vari_Sorule_no = "";
                model_sqso_jisyacvitem.Vari_Sosaki_sorule_guid = "";
                model_sqso_jisyacvitem.Vari_Sosaki_sorule_no = "";
                model_sqso_jisyacvitem.Vari_Sokin_rit = "";
                model_sqso_jisyacvitem.Vari_Calc_sqgak = "";
                model_sqso_jisyacvitem.Vari_Sokotik_guid = "";
                model_sqso_jisyacvitem.Vari_Sokaisyu_kbn = "";
                model_sqso_jisyacvitem.Vari_Sokaisyu_guid = "";
                model_sqso_jisyacvitem.Vari_Siwake_flg = "";

                model_sqso_jisyacvitem.Vari_History = CommonModule.DefHistory;

                // 移行テーブルのフィールド名と移行値を紐付けたハッシュテーブルを作成
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_sqso_jisyacvitem);

                // 挿入処理
                CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

            }

            /// <summary>
            /// 請求情報のうち送金確定されているデータを送金管理情報へ反映 '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            public void Set_SqdataSoKanri(SqlConnection sqlcnnv10)
            {

                bool normalflg = true;

                // sqdata_soを元に送金額を集約した仮テーブルの作成
                Set_SqsoTmpTable(sqlcnnv10);

                // sokanriへ挿入
                Set_Sokanri(sqlcnnv10);


            }

            /// <summary>
            /// 請求情報のうち自社支払データを請求消込情報へ反映 '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_sqcvitem"></param>
            /// <param name="model_sqsocvitem"></param>
            /// <param name="hash_sqsocvitem"></param>
            /// <remarks></remarks>
            public void Set_Sqdatakesi(SqlConnection sqlcnnv10, Model.Sqdata_sqkesi_Model model_sqkesicvitem)
            {

                bool normalflg = true;

                string tblname = "sqdata_sqkesi";
                string fldnamegrp = "sqmei_guid,sqkesi_recno,kesi_ymd,nkbn_jitu,kesi_gak," + "kesi_zeigak,sui_guid,hikiazu_guid,idoazu_guid,printazuryo_guid," + "siwake_flg,tanto_no,history";


                string tmp_azuguid = "";


                // 固定値セット
                model_sqkesicvitem.Vari_Sqkesi_recno = 1.ToString();
                model_sqkesicvitem.Vari_Nkbn_jitu = "";
                model_sqkesicvitem.Vari_Sui_guid = "";
                tmp_azuguid = Guid.NewGuid().ToString();
                model_sqkesicvitem.Vari_Hikiazu_guid = tmp_azuguid;
                model_sqkesicvitem.Vari_Idoazu_guid = "";
                model_sqkesicvitem.Vari_Printazuryo_guid = "";
                model_sqkesicvitem.Vari_Siwake_flg = "";
                model_sqkesicvitem.Vari_Tanto_no = "99001";

                model_sqkesicvitem.Vari_History = CommonModule.DefHistory;

                // 移行テーブルのフィールド名と移行値を紐付けたハッシュテーブルを作成
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_sqkesicvitem);

                // 挿入処理
                CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                // 預り金情報への移行
                Set_Azukanri(sqlcnnv10, tmp_azuguid);             // azukanri作成モジュール

            }

            /// <summary>
            /// 請求情報のうち自社支払データを預り管理テーブルへ反映 '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash_sqcvitem"></param>
            /// <param name="model_sqsocvitem"></param>
            /// <param name="hash_sqsocvitem"></param>
            /// <remarks></remarks>
            public void Set_Azukanri(SqlConnection sqlcnnv10, string azuguid)
            {

                bool normalflg = true;

                var model_azucvitem = new Model.Azukanri_Model();

                string tblname = "azukanri";
                string fldnamegrp = "azu_guid,azu_ymd,nkin_no,dispnkin_name,nkbn_no," + "azu_gak,stakeholder_kbn,stakeholder_no,sui_guid,tanto_no," + "bk_guid,hy_guid,ky_guid,ky_recno,tuki_kbn," + "ky_nkin_no,nkin_recno,kynkin_guid,yoteiazu_flg,printazuryo_guid," + "siwake_flg,biko,history,tujo_siwakegak";




                string tmp_azuguid = "";


                // 固定値セット
                model_azucvitem.Vari_Azu_guid = azuguid;
                model_azucvitem.Vari_Azu_ymd = "";
                model_azucvitem.Vari_Nkin_no = "";
                model_azucvitem.Vari_Dispnkin_name = "";
                model_azucvitem.Vari_Nkbn_no = "";
                model_azucvitem.Vari_Azu_gak = "";
                model_azucvitem.Vari_Stakeholder_kbn = 900.ToString();
                model_azucvitem.Vari_Stakeholder_no = "";
                model_azucvitem.Vari_Sui_guid = "";
                model_azucvitem.Vari_Tanto_no = "";
                model_azucvitem.Vari_Bk_guid = "";
                model_azucvitem.Vari_Hy_guid = "";
                model_azucvitem.Vari_Ky_guid = "";
                model_azucvitem.Vari_Ky_recno = "";
                model_azucvitem.Vari_Tuki_kbn = "";
                model_azucvitem.Vari_Ky_nkin_no = "";
                model_azucvitem.Vari_Nkin_recno = "";
                model_azucvitem.Vari_Kynkin_guid = "";
                model_azucvitem.Vari_Yoteiazu_flg = "";
                model_azucvitem.Vari_Printazuryo_guid = "";
                model_azucvitem.Vari_Siwake_flg = "";
                model_azucvitem.Vari_Biko = "※※　システム自動作成　※※";
                model_azucvitem.Vari_History = CommonModule.DefHistory;

                model_azucvitem.Vari_Tujo_siwakegak = "";

                // 移行テーブルのフィールド名と移行値を紐付けたハッシュテーブルを作成
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_azucvitem);

                // 挿入処理
                CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

            }

            /// <summary>
            /// bk_guidからsorule_guidを取得 '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="bkguid"></param>
            /// <param name="soruleguid"></param>
            /// <param name="soruleno"></param>
            /// <remarks></remarks>
            private void Set_SoruleInfo(SqlConnection sqlcnnv10, string bkguid, ref string soruleguid, ref string soruleno)
            {

                string tmp_sql_soruleguid = " SELECT CONVERT(varchar(50),sorule_guid) FROM sorule WHERE relation_guid = '" + bkguid + "'";
                soruleguid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql_soruleguid, ref sqlcnnv10));

                string tmp_sql_soruleno = " SELECT CONVERT(varchar(50),sorule_no) FROM sorule WHERE relation_guid = '" + bkguid + "'";
                soruleno = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql_soruleno, ref sqlcnnv10));

            }

            /// <summary>
            /// bk_guid、送金先No、送金先口座Noからsorule_guidを取得 '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="bkguid"></param>
            /// <param name="soruleguid"></param>
            /// <param name="soruleno"></param>
            /// <remarks></remarks>
            private void Set_SoruleSosakiInfo(SqlConnection sqlcnnv10, string bkguid, string sono, string sokozano, ref string soruleguid, ref string soruleno)
            {

                string tmp_sql_soruleguid = "";
                tmp_sql_soruleguid = tmp_sql_soruleguid + " SELECT CONVERT(varchar(50),SOSAKI.sorule_guid) FROM sorule_sosaki AS SOSAKI ";
                tmp_sql_soruleguid = tmp_sql_soruleguid + " LEFT JOIN sorule AS SO ON SOSAKI.sorule_guid = SO.sorule_guid ";
                tmp_sql_soruleguid = tmp_sql_soruleguid + " WHERE relation_guid = '" + bkguid + "'";
                tmp_sql_soruleguid = tmp_sql_soruleguid + " AND   so_ow_no = " + sono;
                tmp_sql_soruleguid = tmp_sql_soruleguid + " AND   so_ow_kozano = " + sokozano;
                soruleguid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql_soruleguid, ref sqlcnnv10));

                string tmp_sql_soruleno = "";
                tmp_sql_soruleno = tmp_sql_soruleno + " SELECT CONVERT(varchar(50),SOSAKI.sorule_no) FROM sorule_sosaki AS SOSAKI ";
                tmp_sql_soruleno = tmp_sql_soruleno + " LEFT JOIN sorule AS SO ON SOSAKI.sorule_guid = SO.sorule_guid ";
                tmp_sql_soruleguid = tmp_sql_soruleguid + " WHERE relation_guid = '" + bkguid + "'";
                tmp_sql_soruleguid = tmp_sql_soruleguid + " AND   so_ow_no = " + sono;
                tmp_sql_soruleguid = tmp_sql_soruleguid + " AND   so_ow_kozano = " + sokozano;
                soruleno = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql_soruleno, ref sqlcnnv10));

            }

            /// <summary>
            /// sqdata_soを元に送金額を集約した仮テーブルの作成 '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            private void Set_SqsoTmpTable(SqlConnection sqlcnnv10)
            {

                int tmp_rowcnt = 0;

                // 仮テーブルの初期化
                string tmp_sqldrop = DBQuery.Qry_DropInfo("tmp_sqso", true);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldrop, ref tmp_rowcnt);

                // 仮テーブルの作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + " CREATE TABLE tmp_sqso ";
                tmp_sql_create = tmp_sql_create + " 	( ";
                tmp_sql_create = tmp_sql_create + " 		 so_gak money ";
                tmp_sql_create = tmp_sql_create + " 		,so_yoteiymd datetime ";
                tmp_sql_create = tmp_sql_create + " 		,sorule_guid uniqueidentifier ";
                tmp_sql_create = tmp_sql_create + " 		,sorule_no int ";
                tmp_sql_create = tmp_sql_create + " 		,sosaki_sorule_guid uniqueidentifier ";
                tmp_sql_create = tmp_sql_create + " 		,sosaki_sorule_no int ";
                tmp_sql_create = tmp_sql_create + " 		,sokotik_guid uniqueidentifier	 ";
                tmp_sql_create = tmp_sql_create + " 	) ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmp_rowcnt);

                // 仮テーブルへデータ挿入
                string tmp_sql_insert = "";
                tmp_sql_insert = tmp_sql_insert + " INSERT INTO tmp_sqso ";
                tmp_sql_insert = tmp_sql_insert + " 	SELECT ";
                tmp_sql_insert = tmp_sql_insert + " 		 SUM(so_gak) + SUM(so_zeigak) AS so_gak ";
                tmp_sql_insert = tmp_sql_insert + " 		,so_yoteiymd ";
                tmp_sql_insert = tmp_sql_insert + " 		,sorule_guid ";
                tmp_sql_insert = tmp_sql_insert + " 		,sorule_no ";
                tmp_sql_insert = tmp_sql_insert + " 		,sosaki_sorule_guid ";
                tmp_sql_insert = tmp_sql_insert + " 		,sosaki_sorule_no ";
                tmp_sql_insert = tmp_sql_insert + " 		,NEWID() AS sokotik_guid ";
                tmp_sql_insert = tmp_sql_insert + " 	FROM sqdata_so ";
                tmp_sql_insert = tmp_sql_insert + " 	GROUP BY ";
                tmp_sql_insert = tmp_sql_insert + " 		 so_yoteiymd ";
                tmp_sql_insert = tmp_sql_insert + " 		,sorule_guid ";
                tmp_sql_insert = tmp_sql_insert + " 		,sorule_no ";
                tmp_sql_insert = tmp_sql_insert + " 		,sosaki_sorule_guid ";
                tmp_sql_insert = tmp_sql_insert + " 		,sosaki_sorule_no ";
                tmp_sql_insert = tmp_sql_insert + " 	HAVING sorule_guid IS NOT NULL ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_rowcnt);

            }

            /// <summary>
            /// sokanriの作成 '20160601 請求データ取得処理の修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            private void Set_Sokanri(SqlConnection sqlcnnv10)
            {

                int tmp_rowcnt = 0;

                // sqdata_soおよび仮テーブルからsokanriを作成
                string tmp_sql = "";
                tmp_sql = tmp_sql + " INSERT INTO sokanri ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	* ";
                tmp_sql = tmp_sql + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 NEWID() AS sokanri_guid ";
                tmp_sql = tmp_sql + " 		,SQSO.sorule_guid AS sorule_guid ";
                tmp_sql = tmp_sql + " 		,SQSO.sorule_no AS sorule_no ";
                tmp_sql = tmp_sql + " 		,bk_guid AS bk_guid ";
                tmp_sql = tmp_sql + " 		,NULL AS hy_guid ";
                tmp_sql = tmp_sql + " 		,so_yoteiymd AS so_ymd ";
                tmp_sql = tmp_sql + " 		,200 AS sosaki_kbn ";
                tmp_sql = tmp_sql + " 		,SOSAKI.so_ow_no AS sosaki_no ";
                tmp_sql = tmp_sql + " 		,SOSAKI.so_ow_kozano AS sokoza_no ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN sosaki_koteiflg = 0 THEN so_gak * so_rit / 100 ";
                tmp_sql = tmp_sql + " 			WHEN sosaki_koteiflg = 1 THEN ";
                tmp_sql = tmp_sql + " 				CASE ";
                tmp_sql = tmp_sql + " 					WHEN ISNULL(so_fixsogak,0) <> 0 THEN so_fixsogak ";
                tmp_sql = tmp_sql + " 					WHEN ISNULL(so_fixsogak,0) = 0 THEN so_gak - so_fixsogak_total	/*合計-固定額合計*/ ";
                tmp_sql = tmp_sql + " 				END ";
                tmp_sql = tmp_sql + " 		 END AS so_gak ";
                tmp_sql = tmp_sql + " 		,sokotik_guid AS sokotik_guid ";
                tmp_sql = tmp_sql + " 		,0 AS karikaku_flg ";
                tmp_sql = tmp_sql + " 		,0 AS sgf_no ";
                tmp_sql = tmp_sql + " 		,0 AS error_kbn ";
                tmp_sql = tmp_sql + " 		,'' AS biko ";
                tmp_sql = tmp_sql + " 		,'' AS history ";
                tmp_sql = tmp_sql + " 		,NULL AS karikaku_guid ";
                tmp_sql = tmp_sql + " 		,0 AS shdirect_flg ";
                tmp_sql = tmp_sql + " 	FROM tmp_sqso AS SQSO		/*作成しておいた仮テーブル*/ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 SOBASE.bk_guid ";
                tmp_sql = tmp_sql + " 			,SOSAKI1.sorule_guid ";
                tmp_sql = tmp_sql + " 			,SOSAKI1.sorule_no ";
                tmp_sql = tmp_sql + " 			,SOSAKI1.sosaki_recno ";
                tmp_sql = tmp_sql + " 			,SOSAKI2.so_ow_no ";
                tmp_sql = tmp_sql + " 			,SOSAKI1.so_ow_kozano ";
                tmp_sql = tmp_sql + " 			,SOBASE.sosaki_koteiflg ";
                tmp_sql = tmp_sql + " 			,SOSAKI1.so_rit ";
                tmp_sql = tmp_sql + " 			,SOSAKI1.so_fixsogak ";
                tmp_sql = tmp_sql + " 		FROM sorule_sosaki AS SOSAKI1 ";
                tmp_sql = tmp_sql + " 		LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT sorule_guid,sorule_no,sosaki_koteiflg,bk_guid FROM sorule AS SO ";
                tmp_sql = tmp_sql + " 			LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " 		) AS SOBASE ";
                tmp_sql = tmp_sql + " 		ON  SOSAKI1.sorule_guid = SOBASE.sorule_guid ";
                tmp_sql = tmp_sql + " 		AND SOSAKI1.sorule_no = SOBASE.sorule_no ";
                tmp_sql = tmp_sql + " 		LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT  ";
                tmp_sql = tmp_sql + " 				 sorule_guid ";
                tmp_sql = tmp_sql + " 				,sorule_no ";
                tmp_sql = tmp_sql + " 				,so_ow_no ";
                tmp_sql = tmp_sql + " 			FROM sorule_sosaki ";
                tmp_sql = tmp_sql + " 			WHERE sosaki_recno = 1 ";
                tmp_sql = tmp_sql + " 		) AS SOSAKI2 ";
                tmp_sql = tmp_sql + " 		ON  SOSAKI1.sorule_guid = SOSAKI2.sorule_guid ";
                tmp_sql = tmp_sql + " 		AND SOSAKI1.sorule_no = SOSAKI2.sorule_no ";
                tmp_sql = tmp_sql + " 		WHERE SOSAKI1.so_ow_kozano IS NOT NULL ";
                tmp_sql = tmp_sql + " 	) AS SOSAKI ";
                tmp_sql = tmp_sql + " 	ON  SQSO.sosaki_sorule_guid = SOSAKI.sorule_guid ";
                tmp_sql = tmp_sql + " 	AND SQSO.sorule_no = SOSAKI.sorule_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					SOSAKI1.sorule_guid,SOSAKI1.sorule_no,SOSAKI2.so_ow_no,SUM(SOSAKI1.so_fixsogak) AS so_fixsogak_total ";
                tmp_sql = tmp_sql + " 				FROM sorule_sosaki AS SOSAKI1 ";
                tmp_sql = tmp_sql + " 				LEFT JOIN ";
                tmp_sql = tmp_sql + " 					( ";
                tmp_sql = tmp_sql + " 						SELECT * FROM sorule_sosaki WHERE sosaki_recno = 1 ";
                tmp_sql = tmp_sql + " 					) AS SOSAKI2 ";
                tmp_sql = tmp_sql + " 				ON  SOSAKI1.sorule_guid = SOSAKI2.sorule_guid ";
                tmp_sql = tmp_sql + " 				AND SOSAKI1.sorule_no = SOSAKI2.sorule_no ";
                tmp_sql = tmp_sql + " 				GROUP BY SOSAKI1.sorule_guid,SOSAKI1.sorule_no,SOSAKI2.so_ow_no ";
                tmp_sql = tmp_sql + " 			) AS VW ";
                tmp_sql = tmp_sql + " 		) AS GETKOTEIGAK ";
                tmp_sql = tmp_sql + " 	ON  SQSO.sosaki_sorule_guid = GETKOTEIGAK.sorule_guid ";
                tmp_sql = tmp_sql + " 	AND SQSO.sorule_no = GETKOTEIGAK.sorule_no ";
                tmp_sql = tmp_sql + " ) AS VW ";

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_rowcnt);

                // 履歴を作成
                string tmp_replacestr = "<?xml version=" + "\"" + "1.0" + "\"" + " encoding=" + "\"" + "utf-16" + "\"" + "?>";
                string tmp_sql_update = " UPDATE sokanri SET history = '" + CommonModule.DefHistory.Replace(tmp_replacestr, "") + "'";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_update, ref tmp_rowcnt);

                // 仮テーブルの削除
                string tmp_sqldrop = DBQuery.Qry_DropInfo("tmp_sqso", true);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldrop, ref tmp_rowcnt);

            }

        }

    }

    #endregion

    #region 送金確定情報

    // 請求情報移行と同時に作成するためここでは処理をしない

    // Public Class Sqdata_so_Repository

    // Public Class SubConv
    // Implements IConv

    // ''' <summary>
    // ''' 【中間ファイル→変数】
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="syorikomok"></param>
    // ''' <returns></returns>
    // ''' <remarks></remarks>
    // Public Function Set_Vari(ByVal sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

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
    // Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト(一棟所有)
    // Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
    // Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
    // Dim hash_bkguid As New SafeDictionary<string, string>                            'キー/guid格納用ハッシュテーブル
    // Dim hash_bkhyguid As New SafeDictionary<string, string>                            'キー/guid格納用ハッシュテーブル
    // Dim hash_bkhykyguid As New SafeDictionary<string, string>                            'キー/guid格納用ハッシュテーブル

    // Dim tmp_hash As New SafeDictionary<string, string>                                   '作業用ハッシュテーブル
    // Dim normalflg As Boolean                                        'INSERT正常終了フラグ
    // Dim rtn As Boolean = True                                       '戻り値

    // Dim model_cvitem As New Njc.Model.Unyotainodata_Model                '移行値格納用モデル初期化
    // Dim keycol_main As Integer = 1                                  'メインキー列
    // Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
    // Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
    // Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
    // 'Dim keycol_sub4 As Integer = 5                                  'サブ4キー列
    // Dim viewnamebktobkguid As String = PRE_VIEW_NAME & "物件キー情報"
    // Dim viewnamebkhytohyguid As String = PRE_VIEW_NAME & "物件部屋キー情報"
    // Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

    // '************************
    // '作業準備
    // '************************

    // '紐付けデータ取得
    // Call SetRelItemToObject.Set_RelData_Nkinkomk()
    // Call SetRelItemToObject.Set_RelData_Nkinkbn()

    // 'Excelファイル初期設定
    // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
    // If Not rtn Then
    // Return rtn
    // End If

    // 'キー取得用のVIEWを作成
    // Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebktobkguid)
    // Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyguid)
    // Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)

    // 'テーブル名/フィールド名セット
    // Dim tblname As String = "unyotainodata"
    // Dim fldnamegrp As String = "unyotaino_guid,bk_guid,hy_guid,ky_guid,ky_recno," & _
    // "gt_ym,sq_simeymd,nkin_no,nkbn_yotei,yotei_yatin_kozano," & _
    // "sqsaki_no,sq_gak,sq_zeikbn,sq_zeigak,so_gak," & _
    // "so_zeikbn,so_zeigak,sosumi_flg,kanrigak_rit,sosaki_no," & _
    // "sokoza_no,so_rit,soymd_gaitoukbn,soymd_simekbn,soymd_sokintukikbn," & _
    // "soymd_sokinsimekbn,zei_rit,sqbuild_flg,tanto_no,biko," & _
    // "history,kanri_taxflg,kanritesu_zeiumukbn,kanri_utizeiflg"

    // '追加コンバート時の重複チェック用に既存データのキーを取得
    // If Not InitDBFlg Then
    // Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
    // End If

    // 'guidを取得
    // Call Me.Get_Guid(sqlcnnv10, viewnamebktobkguid, hash_bkguid)
    // Call Me.Get_Guid(sqlcnnv10, viewnamebkhytohyguid, hash_bkhyguid)
    // Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)

    // '親マスタ取得
    // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)

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
    // Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    // 'キーヘッダー名取得
    // Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
    // Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
    // Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)
    // Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)

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
    // Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

    // 'キー値取得
    // Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
    // Dim tmp_keysub1 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub1))
    // Dim tmp_keysub2 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub2))
    // Dim tmp_keysub3 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub3))
    // Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
    // .Vari_Bk_guid = tmp_keymain
    // .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1
    // .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2

    // 'ログ出力用
    // Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
    // fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
    // fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
    // fldname_keysub3 & " = " & tmp_keysub3

    // '作業用変数
    // Dim tmp_nkinname As String = ""
    // Dim tmp_nkinkbn As String = ""

    // '移行値取得
    // For cntjj = 1 To columncnt

    // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    // Dim fldvalue As String = ""
    // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
    // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
    // End If

    // Select Case fldname
    // Case "物件No"
    // '.Vari_Bk_guid = fldvalue.Trim
    // '取得済み
    // Case "部屋No"
    // '.Vari_Hy_guid = fldvalue.Trim
    // '取得済み
    // Case "契約No"
    // '.Vari_Ky_guid = fldvalue.Trim
    // '取得済み
    // Case "契約レコードNo"
    // .Vari_Ky_recno = fldvalue.Trim
    // Case "該当年月"
    // .Vari_Gt_ym = fldvalue.Trim
    // Case "請求締め日"
    // .Vari_Sq_simeymd = fldvalue.Trim
    // Case "入金項目名"
    // tmp_nkinname = fldvalue.Trim
    // Case "入金項目区分"
    // tmp_nkinkbn = EtcMethod.Get_Nkinruiname(fldvalue.Trim)
    // Case "(入金予定)入金区分"
    // .Vari_Nkbn_yotei = fldvalue.Trim
    // Case "(入金予定)家賃入金口座"
    // .Vari_Yotei_yatin_kozano = fldvalue.Trim
    // Case "請求先No"
    // .Vari_Sqsaki_no = fldvalue.Trim
    // Case "請求額"
    // .Vari_Sq_gak = fldvalue.Trim
    // Case "請求税区分"
    // .Vari_Sq_zeikbn = fldvalue.Trim
    // Case "請求税額"
    // .Vari_Sq_zeigak = fldvalue.Trim
    // Case "送金額"
    // .Vari_So_gak = fldvalue.Trim
    // Case "送金税区分"
    // .Vari_So_zeikbn = fldvalue.Trim
    // Case "送金税額"
    // .Vari_So_zeigak = fldvalue.Trim
    // Case "既に送金済みフラグ"
    // .Vari_Sosumi_flg = fldvalue.Trim
    // Case "管理手数料率"
    // .Vari_Kanrigak_rit = fldvalue.Trim
    // Case "送金先家主No"
    // .Vari_Sosaki_no = fldvalue.Trim
    // Case "送金先家主口座No"
    // .Vari_Sokoza_no = fldvalue.Trim
    // Case "送金率"
    // .Vari_So_rit = fldvalue.Trim
    // Case "送金日決定方法 - 該当年月"
    // .Vari_Soymd_gaitoukbn = fldvalue.Trim
    // Case "送金日決定方法 - 締日"
    // .Vari_Soymd_simekbn = fldvalue.Trim
    // Case "送金日決定方法 - 送金月"
    // .Vari_Soymd_sokintukikbn = fldvalue.Trim
    // Case "送金日決定方法 - 送金締日"
    // .Vari_Soymd_sokinsimekbn = fldvalue.Trim
    // Case "適用税率"
    // .Vari_Zei_rit = fldvalue.Trim
    // Case "請求データ生成フラグ"
    // .Vari_Sqbuild_flg = fldvalue.Trim
    // Case "担当者No"
    // .Vari_Tanto_no = fldvalue.Trim
    // Case "備考"
    // .Vari_Biko = fldvalue.Trim
    // Case "管理手数料税フラグ"
    // .Vari_Kanri_taxflg = fldvalue.Trim
    // Case "管理手数料税込フラグ"
    // .Vari_Kanritesu_zeiumukbn = fldvalue.Trim
    // Case "管理手数料率内税フラグ"
    // .Vari_Kanri_utizeiflg = fldvalue.Trim

    // End Select

    // Next

    // '入金項目格納
    // .Vari_Nkin_no = tmp_nkinname & "-" & tmp_nkinkbn

    // '固定値
    // .Vari_Unyotaino_guid = Guid.NewGuid.ToString
    // .Vari_History = DefHistory

    // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
    // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

    // 'データチェック
    // Dim skipflg As Boolean = False
    // Dim hash_cvitem As New SafeDictionary<string, string>
    // Dim hash_log As New SafeDictionary<string, string>
    // skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

    // '書込処理
    // If Not skipflg Then

    // 'キーをguidへ変換
    // 'Dim tmp_bkno As String = hash_cvitem.Item("bk_guid")
    // 'Dim tmp_hyno As String = hash_cvitem.Item("hy_guid")
    // 'Dim tmp_kyno As String = hash_cvitem.Item("ky_guid")
    // 'hash_cvitem("bk_guid") = hash_bkguid.Item(tmp_bkno)
    // 'hash_cvitem("hy_guid") = hash_bkhyguid.Item(tmp_bkno & "-" & tmp_hyno)
    // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(tmp_bkno & "-" & tmp_hyno & "-" & tmp_kyno)
    // hash_cvitem("bk_guid") = hash_bkguid.Item(hash_cvitem.Item("bk_guid"))
    // hash_cvitem("hy_guid") = hash_bkhyguid.Item(hash_cvitem.Item("hy_guid"))
    // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))

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

    // '作業用VIEWの削除
    // Call Me.Drop_TmpView(sqlcnnv10, viewnamebktobkguid)
    // Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid)
    // Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)

    // 'Excelファイル終了設定
    // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

    // '返却
    // Return rtn

    // End Function

    // ''' <summary>
    // ''' 親マスタ取得→リスト格納
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="tblname"></param>
    // ''' <param name="list_basedata"></param>
    // ''' <remarks></remarks>
    // Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

    // Dim tmp_sql As String = " SELECT キー FROM " & tblname
    // Dim flg As Boolean = True

    // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_basedata)

    // End Sub

    // Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

    // End Function

    // Public Function Chk_Data(tblname As String, keyvalue As String, list As List(Of String), hash_chkbefore As SafeDictionary<string, string>, ByRef hash_chkafter As SafeDictionary<string, string>, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

    // End Function

    // Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

    // Dim tmp_sql As String = ""
    // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
    // tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
    // tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
    // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
    // Dim flg As Boolean = True

    // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

    // End Sub

    // ''' <summary>
    // ''' guid取得→ハッシュテーブル格納
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="tblname"></param>
    // ''' <param name="hash_guid"></param>
    // ''' <remarks></remarks>
    // Public Sub Get_Guid(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef hash_guid As SafeDictionary<string, string>)

    // Dim tmp_sql As String = "SELECT * FROM " & tblname
    // Dim flg As Boolean = True

    // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)

    // End Sub

    // ''' <summary>
    // ''' キー/guidをセットにしたVIEWの作成
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <remarks></remarks>
    // Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

    // Dim tmpcnt As Integer = 0

    // 'VIEW初期化
    // Call Me.Drop_TmpView(sqlcnnv10, viewname)

    // Dim tmp_viewmainname As String = viewname.Replace(PRE_VIEW_NAME, "")

    // 'VIEW作成
    // Dim tmp_sql_create As String = PRE_VIEW_QRY & viewname & POST_VIEW_QRY

    // Select Case tmp_viewmainname
    // Case "物件キー情報"
    // tmp_sql_create = tmp_sql_create & " SELECT bk_no AS キー,bk_guid FROM bkdata "
    // Case "物件部屋キー情報"
    // tmp_sql_create = tmp_sql_create & " SELECT "
    // tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
    // tmp_sql_create = tmp_sql_create & " 	,hy_guid "
    // tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
    // tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
    // Case "物件部屋契約キー情報"
    // tmp_sql_create = tmp_sql_create & " SELECT "
    // tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
    // tmp_sql_create = tmp_sql_create & " 	,ky_guid "
    // tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
    // tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
    // tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "
    // End Select

    // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

    // End Sub

    // ''' <summary>
    // ''' 作業用VIEWの削除
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="viewname"></param>
    // ''' <remarks></remarks>
    // Public Sub Drop_TmpView(ByVal sqlcnnv10 As SqlConnection, ByVal viewname As String)

    // Dim tmpcnt As Integer = 0

    // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
    // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

    // End Sub

    // End Class

    // End Class

    #endregion

    #region 変動費検針情報

    public class Hendodata_meisai_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="syorikomok"></param>
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhyhendoguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hendodata_meisai_Model();        // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebktobkguid = CommonModule.PRE_VIEW_NAME + "物件キー情報";
                string viewnamebkhytohyguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhytohyhendoguid = CommonModule.PRE_VIEW_NAME + "物件部屋変動費キー情報";
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                var list_hendosqguid = new List<string>();                     // 請求情報初期化用の変動請求guidを格納
                int tmp_sqcnt = 0;
                int tmp_hdbasecnt = 0;

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebktobkguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyhendoguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "hendodata_meisai";
                string fldnamegrp = "bk_guid,hendo_kbn,kensin_ymd,gt_ym,kensin_ptn," + "hendo_guid,hy_guid,hendosq_guid,ky_guid,ky_recno," + "sqsaki_no,nkin_no,nkbn_no,sqsime_ymd,zenkai_ymd," + "zenkai_metervalue,zenkai_siyoryo,zenkai_sqgak,zenkai_sqzeigak,kensin_metervalue," + "kensin_siyoryo,kensin_ryokin1,kensin_ryokin2,kensin_ryokin3,kensin_ryokin4," + "kensin_ryokin5,kensin_sqgak,kensin_sqzeigak,kensin_biko,kensin_soyoteiymd," + "kensin_sorit,kensin_sogak,kensin_sozeigak,zenkaihendosq_guid";






                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebktobkguid, ref hash_bkguid);
                Get_Guid(sqlcnnv10, viewnamebkhytohyguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhytohyhendoguid, ref hash_bkhyhendoguid);
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, viewnamebkhytohyguid, ref list_basekeydata);

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

                // 初期化用に変動請求guidを取得
                Get_HendosqGuid(sqlcnnv10, ref list_hendosqguid);

                // 他の移行項目用モデル
                var model_hdbasecvitem = new Model.Hendodata_Model();      // 変動検針データ(基本)移行用モデル
                var model_sqcvitem = new Model.Sqdata_Model();            // 請求情報移行用モデル

                // 20160624 請求関連情報 親データ無しログ出力制御修正 -add
                int chg_totalcnt = rowcnt;        // 親データ無しのデータ件数調整用(カウントしないため)

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
                string fldname_keysub4 = Conversions.ToString(headervalue(startrow - 1, keycol_sub4));

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
                    string tmp_keysub4 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub4]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4;
                    model_cvitem.Vari_Bk_guid = tmp_keymain;
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;
                    model_cvitem.Vari_Hendo_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub4;  // 部屋各戸メーター情報から変動費guidを取得するためキーを格納しておく
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




                    // 作業用変数
                    string tmp_nkinname = "";
                    // 20160530 ログ修正 -chg sta
                    // Dim tmp_nkinkbn As String = ""
                    string argtukikbn = "5";
                    string tmp_nkinkbn = EtcMethod.Get_Nkinruiname(ref argtukikbn);
                    // 20160530 ログ修正 -chg end
                    string tmp_olddata = "";
                    string tmp_newdata = "";

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        // ************************変動費検針情報の取得 (変動費検針データと請求データから取得)************************
                        switch (fldname ?? "")
                        {

                            // 変動費データから取得
                            case "物件No":
                                {
                                    break;
                                }
                            // .Vari_Bk_guid = fldvalue
                            // 取得済み
                            case "部屋No":
                                {
                                    break;
                                }
                            // .Vari_Hy_guid = fldvalue
                            // 取得済み
                            case "契約No":
                                {
                                    break;
                                }
                            // .Vari_Ky_guid = fldvalue
                            // 取得済み
                            case "契約レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue;
                                    model_sqcvitem.Vari_Ky_recno = fldvalue;
                                    break;
                                }
                            case "変動区分":
                                {
                                    break;
                                }
                            // 20160603 ユーザーデータ検証による修正 -del sta
                            // 部屋、契約変動費情報と同様に紐付設定値から変動区分(メーター分類)を取得する
                            // 入金項目名から取得するためここでは取得しない
                            // .Vari_Hendo_kbn = fldvalue
                            // 20160603 ユーザーデータ検証による修正 -del end
                            case "検針日":
                                {
                                    model_cvitem.Vari_Kensin_ymd = fldvalue;
                                    break;
                                }
                            case "該当年月":
                                {
                                    model_cvitem.Vari_Gt_ym = fldvalue;
                                    break;
                                }
                            case "検針パターン":
                                {
                                    model_cvitem.Vari_Kensin_ptn = fldvalue;
                                    break;
                                }
                            case "前回値":
                                {
                                    model_cvitem.Vari_Zenkai_metervalue = fldvalue;
                                    break;
                                }
                            case "今回値":
                                {
                                    model_cvitem.Vari_Kensin_metervalue = fldvalue;
                                    break;
                                }
                            case "今回使用量":
                                {
                                    model_cvitem.Vari_Kensin_siyoryo = fldvalue;
                                    break;
                                }
                            case "検針料金1":
                                {
                                    model_cvitem.Vari_Kensin_ryokin1 = fldvalue;
                                    break;
                                }
                            case "検針料金2":
                                {
                                    model_cvitem.Vari_Kensin_ryokin2 = fldvalue;
                                    break;
                                }
                            case "検針料金3":
                                {
                                    model_cvitem.Vari_Kensin_ryokin3 = fldvalue;
                                    break;
                                }
                            case "検針料金4":
                                {
                                    model_cvitem.Vari_Kensin_ryokin4 = fldvalue;
                                    break;
                                }
                            case "検針料金5":
                                {
                                    model_cvitem.Vari_Kensin_ryokin5 = fldvalue;
                                    break;
                                }
                            case "変動費検針情報備考":
                                {
                                    model_cvitem.Vari_Kensin_biko = fldvalue;
                                    break;
                                }
                            case "最新データ取得用":     // 最新最古データ判別用 (「1」の場合最新)
                                {
                                    tmp_olddata = fldvalue;
                                    break;
                                }
                            case "最古データ取得用":     // 最新最古データ判別用 (「1」の場合最古)
                                {
                                    tmp_newdata = fldvalue;
                                    break;
                                }

                            // 請求データから取得
                            case "請求先No":
                                {
                                    model_cvitem.Vari_Sqsaki_no = fldvalue;
                                    break;
                                }
                            case "入金項目名":
                                {
                                    tmp_nkinname = fldvalue;
                                    // 20160530 ログ修正 -chg sta
                                    // Case "入金項目区分"
                                    // .Vari_Nkbn_no = fldvalue
                                    // tmp_nkinkbn = EtcMethod.Get_Nkinruiname("5")
                                    string argtukikbn1 = "5";
                                    model_cvitem.Vari_Hendo_kbn = fldvalue + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn1);
                                    break;
                                }
                            case "(入金予定)入金区分":
                                {
                                    model_cvitem.Vari_Nkbn_no = fldvalue;
                                    break;
                                }
                            // 20160530 ログ修正 -chg end
                            case "請求締切日":
                                {
                                    model_cvitem.Vari_Sqsime_ymd = fldvalue;
                                    break;
                                }
                            case "請求額":
                                {
                                    model_cvitem.Vari_Kensin_sqgak = fldvalue;
                                    break;
                                }
                            case "請求税額":
                                {
                                    model_cvitem.Vari_Kensin_sqzeigak = fldvalue;
                                    break;
                                }
                            case "送金予定日":
                                {
                                    model_cvitem.Vari_Kensin_soyoteiymd = fldvalue;
                                    break;
                                }
                            case "送金率":
                                {
                                    model_cvitem.Vari_Kensin_sorit = fldvalue;
                                    break;
                                }
                            case "送金額":
                                {
                                    model_cvitem.Vari_Kensin_sogak = fldvalue;
                                    break;
                                }
                            case "送金税額":
                                {
                                    model_cvitem.Vari_Kensin_sozeigak = fldvalue;
                                    break;
                                }
                        }

                        // ************************変動費検針情報 (基本) の取得************************
                        switch (fldname ?? "")
                        {
                            // 20160603 ユーザーデータ検証による修正 -del sta
                            // Case "変動区分"
                            // 紐付設定値から取得した変動費区分を後で設定するためここではコメントアウト
                            // model_hdbasecvitem.Vari_Hendo_kbn = fldvalue
                            // 20160603 ユーザーデータ検証による修正 -del end
                            case "検針日":
                                {
                                    model_hdbasecvitem.Vari_Kensin_ymd = fldvalue;
                                    break;
                                }
                            case "該当年月":
                                {
                                    model_hdbasecvitem.Vari_Gt_ym = fldvalue;
                                    break;
                                }
                        }

                        // ************************請求情報の取得************************
                        switch (fldname ?? "")
                        {
                            case "契約レコードNo":
                                {
                                    model_sqcvitem.Vari_Ky_recno = fldvalue;
                                    break;
                                }
                            case "検針日":
                                {
                                    model_sqcvitem.Vari_Kensin_ymd = fldvalue;
                                    break;
                                }
                            case "該当年月":
                                {
                                    model_sqcvitem.Vari_Gt_ym = fldvalue;
                                    break;
                                }
                            case "入金項目名":
                                {
                                    tmp_nkinname = fldvalue;
                                    break;
                                }
                            case "入金項目区分":
                                {
                                    model_sqcvitem.Vari_Nkin_kbn = fldvalue;
                                    break;
                                }
                            // 20160530 ログ修正 -del
                            // tmp_nkinkbn = EtcMethod.Get_Nkinruiname("5")
                            case "入金項目表示用名称":
                                {
                                    model_sqcvitem.Vari_Dispnkin_name = fldvalue;
                                    break;
                                }
                            case "(入金予定)入金区分":
                                {
                                    model_sqcvitem.Vari_Nkbn_yotei = fldvalue;
                                    break;
                                }
                            case "請求額":
                                {
                                    model_sqcvitem.Vari_Sq_gak = fldvalue;
                                    break;
                                }
                            case "請求税区分":
                                {
                                    model_sqcvitem.Vari_Sq_zeikbn = fldvalue;
                                    break;
                                }
                            case "請求税額":
                                {
                                    model_sqcvitem.Vari_Sq_zeigak = fldvalue;
                                    break;
                                }
                            case "請求締切日":
                                {
                                    model_sqcvitem.Vari_Sq_simeymd = fldvalue;
                                    break;
                                }
                            case "複数請求先フラグ":
                                {
                                    model_sqcvitem.Vari_Multisqsaki_flg = fldvalue;
                                    break;
                                }
                            case "請求先区分":
                                {
                                    model_sqcvitem.Vari_Sqsaki_kbn = fldvalue;
                                    break;
                                }
                            case "請求先No":
                                {
                                    model_sqcvitem.Vari_Sqsaki_no = fldvalue;
                                    break;
                                }
                            case "(入金予定)家賃振込先口座No":
                                {
                                    model_sqcvitem.Vari_Yotei_yatin_kozano = fldvalue;
                                    break;
                                }
                            case "口座振替履歴No":
                                {
                                    model_sqcvitem.Vari_Fkae_rirekino = fldvalue;
                                    break;
                                }
                            case "口座振替履歴明細No":
                                {
                                    model_sqcvitem.Vari_Fkae_rirekimeino = fldvalue;
                                    break;
                                }
                            case "振替請求対象フラグ 1:請求中　2:未請求":
                                {
                                    model_sqcvitem.Vari_Fkae_sqtaisyoflg = fldvalue;
                                    break;
                                }
                            case "振替請求対象フラグ　1:対象　2:対象外":
                                {
                                    model_sqcvitem.Vari_Fkae_sqselectflg = fldvalue;
                                    break;
                                }
                            case "振替請求中フラグ":
                                {
                                    model_sqcvitem.Vari_Fkae_sqflg = fldvalue;
                                    break;
                                }
                            case "月区分":
                                {
                                    model_sqcvitem.Vari_Tuki_kbn = fldvalue;
                                    break;
                                }
                            case "入金項目明細No":
                                {
                                    model_sqcvitem.Vari_Nkin_recno = fldvalue;
                                    break;
                                }
                            case "次回契約情報フラグ":
                                {
                                    model_sqcvitem.Vari_Nextky_flg = fldvalue;
                                    break;
                                }
                            case "変動費請求GUID":
                                {
                                    model_sqcvitem.Vari_Hendosq_guid = fldvalue;
                                    break;
                                }
                            // 20160603 ユーザーデータ検証による修正 -del sta
                            // Case "変動費区分"
                            // 紐付設定値から取得した変動費区分を後で設定するためここではコメントアウト
                            // model_sqcvitem.Vari_Hendo_kbn = fldvalue
                            // 20160603 ユーザーデータ検証による修正 -del end
                            case "変動費検針日":
                                {
                                    model_sqcvitem.Vari_Kensin_ymd = fldvalue;
                                    break;
                                }
                            case "送金額":
                                {
                                    model_sqcvitem.Vari_So_gak = fldvalue;
                                    break;
                                }
                            case "送金税区分":
                                {
                                    model_sqcvitem.Vari_So_zeikbn = fldvalue;
                                    break;
                                }
                            case "送金税額":
                                {
                                    model_sqcvitem.Vari_So_zeigak = fldvalue;
                                    break;
                                }
                            case "その他請求フラグ":
                                {
                                    model_sqcvitem.Vari_Etcsq_flg = fldvalue;
                                    break;
                                }
                            case "その他請求の送金率":
                                {
                                    model_sqcvitem.Vari_Etcsq_sorit = fldvalue;
                                    break;
                                }
                            case "その他請求の指定送金日":
                                {
                                    model_sqcvitem.Vari_Etcsq_soymd = fldvalue;
                                    break;
                                }
                            case "分割GUID":
                                {
                                    model_sqcvitem.Vari_Bunkatu_guid = fldvalue;
                                    break;
                                }
                            case "支払・返金区分":
                                {
                                    model_sqcvitem.Vari_Sh_kbn = fldvalue;
                                    break;
                                }
                            case "支払・返金先区分":
                                {
                                    model_sqcvitem.Vari_Shsaki_kbn = fldvalue;
                                    break;
                                }
                            case "支払・返金先No":
                                {
                                    model_sqcvitem.Vari_Shsaki_no = fldvalue;
                                    break;
                                }
                            case "支払・返金先口座No":
                                {
                                    model_sqcvitem.Vari_Shsaki_kozano = fldvalue;
                                    break;
                                }
                            case "一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)":
                                {
                                    model_sqcvitem.Vari_Ikkatu_nkin_bango = fldvalue;
                                    break;
                                }
                            case "支店No":
                                {
                                    model_sqcvitem.Vari_Jisya_no = fldvalue;
                                    break;
                                }
                            case "担当者No":
                                {
                                    model_sqcvitem.Vari_Tanto_no = fldvalue;
                                    break;
                                }
                            case "控除請求項目GUID":
                                {
                                    model_sqcvitem.Vari_Kjsqkmk_guid = fldvalue;
                                    break;
                                }
                            case "固定控除ルールGUID(請求)":
                                {
                                    model_sqcvitem.Vari_Koteirule_guid = fldvalue;
                                    break;
                                }
                            case "控除ルールフラグ":
                                {
                                    model_sqcvitem.Vari_Kojorule_flg = fldvalue;
                                    break;
                                }
                            case "請求書発行予定日":
                                {
                                    model_sqcvitem.Vari_Sq_hakkoyoteiymd = fldvalue;
                                    break;
                                }
                            case "請求ソートNo":
                                {
                                    model_sqcvitem.Vari_Sqsort_no = fldvalue;
                                    break;
                                }
                            case "立替・預りフラグ":
                                {
                                    model_sqcvitem.Vari_Tateazu_flg = fldvalue;
                                    break;
                                }
                            case "立替・預り設定時の預り予定GUID（預り金データ）":
                                {
                                    model_sqcvitem.Vari_Tateazu_guid = fldvalue;
                                    break;
                                }
                            case "敷金保証金随時処理GUID":
                                {
                                    model_sqcvitem.Vari_Sikikinzuiji_guid = fldvalue;
                                    break;
                                }
                            case "運用開始時の未納・滞納金GUID":
                                {
                                    model_sqcvitem.Vari_Unyotaino_guid = fldvalue;
                                    break;
                                }
                            case "送金回収GUID":
                                {
                                    model_sqcvitem.Vari_Sokaisyu_guid = fldvalue;
                                    break;
                                }
                            case "修繕区分":
                                {
                                    model_sqcvitem.Vari_Szen_kbn = fldvalue;
                                    break;
                                }
                            case "随時修繕No":
                                {
                                    model_sqcvitem.Vari_Szen_no = fldvalue;
                                    break;
                                }
                            case "修繕請求No":
                                {
                                    model_sqcvitem.Vari_Szen_sqno = fldvalue;
                                    break;
                                }
                            case "適用税率":
                                {
                                    model_sqcvitem.Vari_Zei_rit = fldvalue;
                                    break;
                                }
                            case "編集フラグ":
                                {
                                    model_sqcvitem.Vari_Edit_flg = fldvalue;
                                    break;
                                }
                            case "更新ロックフラグ":
                                {
                                    model_sqcvitem.Vari_Lock_flg = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_sqcvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            case "立替仕訳額":
                                {
                                    model_sqcvitem.Vari_Tatekae_siwakegak = fldvalue;
                                    break;
                                }
                            case "立替仕訳税額":
                                {
                                    model_sqcvitem.Vari_Tatekae_siwakezeigak = fldvalue;
                                    break;
                                }
                            case "取込日（インポート実施日）":
                                {
                                    model_sqcvitem.Vari_Sq_torikomiymd = fldvalue;
                                    break;
                                }
                            case "コンビニ収納サービスNo":
                                {
                                    model_sqcvitem.Vari_Cvpay_no = fldvalue;
                                    break;
                                }
                            case "顧客番号":
                                {
                                    model_sqcvitem.Vari_Kokyaku_bango = fldvalue;
                                    break;
                                }
                            case "コンビニ収納請求フラグ":
                                {
                                    model_sqcvitem.Vari_Cvpay_sqflg = fldvalue;
                                    break;
                                }
                            case "その他請求管理手数料区分":
                                {
                                    model_sqcvitem.Vari_Etcsq_kanritesukbn = fldvalue;
                                    break;
                                }
                            case "未収仕訳フラグ":
                                {
                                    model_sqcvitem.Vari_Misyu_siwake_flg = fldvalue;
                                    break;
                                }
                            case "雑収入区分":
                                {
                                    model_sqcvitem.Vari_Zatusyu_kbn = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 入金項目格納
                    model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + tmp_nkinkbn;                    // 変動費検針情報用
                    model_sqcvitem.Vari_Nkin_no = tmp_nkinname + "-" + tmp_nkinkbn;      // 請求情報用

                    // 固定値(変動費検針)
                    model_cvitem.Vari_Hendosq_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Zenkaihendosq_guid = "";   // 移行後に一括更新する
                    model_cvitem.Vari_Zenkai_sqgak = "";         // 移行後に一括更新する
                    model_cvitem.Vari_Zenkai_sqzeigak = "";      // 移行後に一括更新する

                    // 固定値(請求情報)
                    model_sqcvitem.Vari_Hendosq_guid = model_cvitem.Vari_Hendosq_guid;
                    model_sqcvitem.Vari_Sqmei_guid = Guid.NewGuid().ToString();
                    model_sqcvitem.Vari_History = CommonModule.DefHistory;

                    // 固定値(変動費検針 (基本))
                    model_hdbasecvitem.Vari_History = CommonModule.DefHistory;

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
                        hash_cvitem["bk_guid"] = hash_bkguid[hash_cvitem["bk_guid"]];
                        hash_cvitem["hy_guid"] = hash_bkhyguid[hash_cvitem["hy_guid"]];
                        // 20160603 ユーザーデータ検証による修正 -del
                        // hash_cvitem("hendo_guid") = hash_bkhyhendoguid.Item(hash_cvitem.Item("hendo_guid"))
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];

                        // 20160603 ユーザーデータ検証による修正 -add sta
                        // 移行用に変換された物件guid、入金項目Noおよび該当年月から固定控除ルールのguidを取得
                        string tmp_hyguid = Conversions.ToString(hash_cvitem["hy_guid"]);
                        string tmp_hendokbn = Conversions.ToString(hash_cvitem["hendo_kbn"]);
                        string tmp_nkinno = Conversions.ToString(hash_cvitem["nkin_no"]);
                        string tmp_hendoguid = Get_HendoGuid(sqlcnnv10, tmp_hyguid, tmp_hendokbn, tmp_nkinno);
                        hash_cvitem["hendo_guid"] = tmp_hendoguid;
                        // 20160603 ユーザーデータ検証による修正 -add end

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // 変動費検針情報 (基本) へ反映
                            // 変換された各値を取得 (変動費検針情報 (基本))
                            // 20160603 ユーザーデータ検証による修正 -add sta
                            // 紐付設定値から取得した変動区分を取得
                            string tmp_chghendkbn = Conversions.ToString(hash_cvitem["hendo_kbn"]);
                            model_hdbasecvitem.Vari_Hendo_kbn = tmp_chghendkbn;
                            // 20160603 ユーザーデータ検証による修正 -add end
                            model_hdbasecvitem.Vari_Bk_guid = Conversions.ToString(hash_cvitem["bk_guid"]);
                            tmp_hdbasecnt = tmp_hdbasecnt + 1;
                            bool hdbasetblinitflg = Conversions.ToBoolean(Interaction.IIf(tmp_hdbasecnt == 1, true, false));
                            Set_Hdbasedata(sqlcnnv10, model_hdbasecvitem, hdbasetblinitflg);

                            // 請求情報へ反映
                            tmp_sqcnt = tmp_sqcnt + 1;
                            bool sqtblinitflg = Conversions.ToBoolean(Interaction.IIf(tmp_sqcnt == 1, true, false));
                            // 変換された各値を取得 (請求情報)
                            // 20160603 ユーザーデータ検証による修正 -add sta
                            // 紐付設定値から取得した変動区分を取得
                            model_sqcvitem.Vari_Hendo_kbn = tmp_chghendkbn;
                            // 20160603 ユーザーデータ検証による修正 -add end
                            model_sqcvitem.Vari_Bk_guid = Conversions.ToString(hash_cvitem["bk_guid"]);
                            model_sqcvitem.Vari_Hy_guid = Conversions.ToString(hash_cvitem["hy_guid"]);
                            model_sqcvitem.Vari_Ky_guid = Conversions.ToString(hash_cvitem["ky_guid"]);
                            model_sqcvitem.Vari_Nkbn_yotei = Conversions.ToString(hash_cvitem["nkbn_no"]);
                            model_sqcvitem.Vari_Nkin_no = Conversions.ToString(hash_cvitem["nkin_no"]);
                            Set_Sqdata(sqlcnnv10, model_sqcvitem, sqtblinitflg, list_hendosqguid);

                        }
                    }

                    // 20160624 請求関連情報 親データ無しログ出力制御修正 -add sta
                    else if (hash_log.Count == 1)
                    {
                        bool logclearflg = false;
                        foreach (var erritem in hash_log)
                        {
                            string tmp_str = Conversions.ToString(erritem.Value);
                            string[] tmp_errnaiyo = tmp_str.Split('-');
                            if ((tmp_errnaiyo[0] ?? "") == CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE)
                            {
                                chg_totalcnt = chg_totalcnt - 1;
                                logclearflg = true;
                            }
                        }
                        if (logclearflg)
                        {
                            hash_log.Clear();
                        }
                        // 20160624 請求関連情報 親データ無しログ出力制御修正 -add end

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
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = chg_totalcnt;
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg end

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebktobkguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhytohyhendoguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 前回分の検針値等一括更新
                int tmpcnt = 0;
                string zenkaiinfoupdatesql = Get_UseQry_Update();
                DBExec.Exec_NonQuery(sqlcnnv10, zenkaiinfoupdatesql, ref tmpcnt);

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN kydata ON " + tblname + ".ky_guid = kydata.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// guid取得→ハッシュテーブル格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="hash_guid"></param>
            /// <remarks></remarks>
            public void Get_Guid(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_guid)
            {

                string tmp_sql = "SELECT * FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);

            }

            /// <summary>
            /// 変動費請求guid取得→リストオブジェクトへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="list_guid"></param>
            /// <remarks></remarks>
            public void Get_HendosqGuid(SqlConnection sqlcnnv10, ref List<string> list_guid)
            {

                string tmp_sql = " SELECT CONVERT(VARCHAR(50),hendosq_guid) FROM sqdata WHERE hendosq_guid IS NOT NULL ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_guid);

            }

            /// <summary>
            /// キー/guidをセットにしたVIEWの作成
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                // VIEW初期化
                Drop_TmpView(sqlcnnv10, viewname);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT bk_no AS キー,bk_guid FROM bkdata ";
                            break;
                        }
                    case "物件部屋キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,hy_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋変動費キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,HYHENDO.hendo_kbn) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,hyhendo_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata_hendo AS HYHENDO ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON HYHENDO.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋契約キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                            break;
                        }
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 作業用VIEWの削除
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="viewname"></param>
            /// <remarks></remarks>
            public void Drop_TmpView(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

            }

            /// <summary>
            /// 変動費検針情報を請求情報へ反映
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hendofldnamegrp"></param>
            /// <param name="hash_hendocvitem"></param>
            /// <param name="model_sqcvitem"></param>
            /// <remarks></remarks>
            public void Set_Sqdata(SqlConnection sqlcnnv10, object model_sqcvitem, bool sqtblinitflg, List<string> list_hendosqguid)
            {

                bool normalflg = true;
                var hash_sqcvitem = new SafeDictionary<string, string>();

                string sqtblname = "sqdata";
                string sqfldnamegrp = "sqmei_guid,gt_ym,nkin_no,dispnkin_name,nkbn_yotei," + "sq_gak,sq_zeikbn,sq_zeigak,sq_simeymd,multisqsaki_flg," + "sqsaki_kbn,sqsaki_no,yotei_yatin_kozano,fkae_rirekino,fkae_rirekimeino," + "fkae_sqtaisyoflg,fkae_sqselectflg,fkae_sqflg,ky_guid,ky_recno," + "tuki_kbn,nkin_recno,nkin_kbn,kynkin_guid,nextky_flg," + "hendosq_guid,hendo_kbn,kensin_ymd,bk_guid,hy_guid," + "so_gak,so_zeikbn,so_zeigak,etcsq_flg,etcsq_sorit," + "etcsq_soymd,bunkatu_guid,sh_kbn,shsaki_kbn,shsaki_no," + "shsaki_kozano,ikkatu_nkin_bango,jisya_no,tanto_no,kjsqkmk_guid," + "koteirule_guid,kojorule_flg,sq_hakkoyoteiymd,sqsort_no,tateazu_flg," + "tateazu_guid,sikikinzuiji_guid,unyotaino_guid,sokaisyu_guid,szen_kbn," + "szen_no,szen_sqno,zei_rit,edit_flg,lock_flg," + "biko,history,tatekae_siwakegak,tatekae_siwakezeigak,sq_torikomiymd," + "cvpay_no,kokyaku_bango,cvpay_sqflg,etcsq_kanritesukbn,misyu_siwake_flg," + "zatusyu_kbn";














                // テーブル初期化処理
                if (sqtblinitflg)
                {
                    // 取得しておいた変動請求guidをクエリ用に変換
                    string deleteguid = "";
                    foreach (var hendosqguid in list_hendosqguid)
                        deleteguid = deleteguid + "," + "'" + hendosqguid + "'";
                    if (!string.IsNullOrEmpty(deleteguid))
                    {
                        deleteguid = deleteguid.Remove(0, 1);
                    }
                    int rowcnt = 0;
                    string tmp_sqldelete = DBQuery.Qry_DelInfo(sqtblname, " hendosq_guid IN (" + deleteguid + ")");
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldelete, ref rowcnt);
                }

                // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                hash_sqcvitem = GetHashFldToValue.Get_Hash_fldvalue(sqfldnamegrp, model_sqcvitem);

                // 挿入処理
                CVDBInsert.Cnv_Db(sqlcnnv10, sqtblname, sqfldnamegrp, hash_sqcvitem, ref normalflg);

            }

            /// <summary>
            /// 変動費検針情報 (基本) へ反映
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="model_hdbasecvitem"></param>
            /// <param name="hdbasetblinitflg"></param>
            /// <remarks></remarks>
            public void Set_Hdbasedata(SqlConnection sqlcnnv10, object model_hdbasecvitem, bool hdbasetblinitflg)
            {

                bool normalflg = true;
                var hash_sqcvitem = new SafeDictionary<string, string>();

                string hdbasetblname = "hendodata";
                string hdbasefldnamegrp = "bk_guid,hendo_kbn,kensin_ymd,gt_ym,history";

                // テーブル初期化処理
                if (hdbasetblinitflg)
                {
                    int rowcnt = 0;
                    string tmp_sqldelete = DBQuery.Qry_DelInfo(hdbasetblname);
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldelete, ref rowcnt);
                }

                // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                hash_sqcvitem = GetHashFldToValue.Get_Hash_fldvalue(hdbasefldnamegrp, model_hdbasecvitem);

                // 挿入処理
                CVDBInsert.Cnv_Db(sqlcnnv10, hdbasetblname, hdbasefldnamegrp, hash_sqcvitem, ref normalflg);

            }

            /// <summary>
            /// 物件guidと送金ルール情報を紐付けたハッシュテーブルを作成
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="typeno"></param>
            /// <remarks></remarks>
            private SafeDictionary<string, string> Get_BkToSoruleInfo(SqlConnection sqlcnnv10, string soymd, int typeno)
            {

                var rtn_hash = new SafeDictionary<string, string>();
                string tmp_sql = "";

                switch (typeno)
                {
                    case 1:
                        {
                            tmp_sql = " SELECT bk_guid,sorule_guid FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid ";
                            tmp_sql = tmp_sql + " WHERE sorule_startymd <= '" + soymd + "' AND ISNULL(sorule_endymd,'21001231') >=  '" + soymd + "'";
                            break;
                        }
                    case 2:
                        {
                            tmp_sql = "SELECT bk_guid,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid";
                            tmp_sql = tmp_sql + " WHERE sorule_startymd <= '" + soymd + "' AND ISNULL(sorule_endymd,'21001231') >=  '" + soymd + "'";
                            break;
                        }
                    case 3:
                        {
                            tmp_sql = tmp_sql + " SELECT bk_guid,SOSAKI.sorule_guid FROM ";
                            tmp_sql = tmp_sql + " ( ";
                            tmp_sql = tmp_sql + " 	SELECT bk_guid,sorule_guid,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid ";
                            tmp_sql = tmp_sql + " 	WHERE sorule_startymd <= '" + soymd + "' AND ISNULL(sorule_endymd,'21001231') >=  '" + soymd + "'";
                            tmp_sql = tmp_sql + " ) AS SO ";
                            tmp_sql = tmp_sql + " LEFT JOIN ";
                            tmp_sql = tmp_sql + " 	( ";
                            tmp_sql = tmp_sql + " 		SELECT DISTINCT sorule_guid,sorule_no FROM sorule_sosaki ";
                            tmp_sql = tmp_sql + " 	) AS SOSAKI ";
                            tmp_sql = tmp_sql + " ON  SO.sorule_guid = SOSAKI.sorule_guid ";
                            tmp_sql = tmp_sql + " AND SO.sorule_no = SOSAKI.sorule_no ";
                            break;
                        }
                    case 4:
                        {
                            tmp_sql = tmp_sql + " SELECT bk_guid,SOSAKI.sorule_no FROM ";
                            tmp_sql = tmp_sql + " ( ";
                            tmp_sql = tmp_sql + " 	SELECT bk_guid,sorule_guid,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid ";
                            tmp_sql = tmp_sql + " 	WHERE sorule_startymd <= '" + soymd + "' AND ISNULL(sorule_endymd,'21001231') >=  '" + soymd + "'";
                            tmp_sql = tmp_sql + " ) AS SO ";
                            tmp_sql = tmp_sql + " LEFT JOIN ";
                            tmp_sql = tmp_sql + " 	( ";
                            tmp_sql = tmp_sql + " 		SELECT DISTINCT sorule_guid,sorule_no FROM sorule_sosaki ";
                            tmp_sql = tmp_sql + " 	) AS SOSAKI ";
                            tmp_sql = tmp_sql + " ON  SO.sorule_guid = SOSAKI.sorule_guid ";
                            tmp_sql = tmp_sql + " AND SO.sorule_no = SOSAKI.sorule_no ";
                            break;
                        }
                }

                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref rtn_hash);

                return rtn_hash;

            }

            /// <summary>
            /// 前回検針値等の一括更新クエリ
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE hendodata_meisai SET ";
                tmp_sql = tmp_sql + " 	 zenkai_ymd = HENDOMEI2.kensin_ymd ";
                tmp_sql = tmp_sql + " 	,zenkai_siyoryo = HENDOMEI2.kensin_siyoryo ";
                tmp_sql = tmp_sql + " 	,zenkai_sqgak = HENDOMEI2.kensin_sqgak ";
                tmp_sql = tmp_sql + " 	,zenkai_sqzeigak = HENDOMEI2.kensin_sqzeigak ";
                tmp_sql = tmp_sql + " 	,zenkaihendosq_guid = HENDOMEI2.hendosq_guid ";
                tmp_sql = tmp_sql + " FROM hendodata_meisai AS HENDOMEI ";
                tmp_sql = tmp_sql + " /*SELECT * FROM hendodata_meisai AS HENDOMEI*/ ";
                tmp_sql = tmp_sql + " LEFT JOIN ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 ROW_NUMBER()OVER(PARTITION BY bk_guid,hendo_kbn,kensin_ptn,hendo_guid,hy_guid ORDER BY kensin_ymd) -1 AS 結合用 ";
                tmp_sql = tmp_sql + " 		,* ";
                tmp_sql = tmp_sql + " 	FROM hendodata_meisai ";
                tmp_sql = tmp_sql + " ) AS HENDOMEI1 ";
                tmp_sql = tmp_sql + " ON HENDOMEI.hendosq_guid = HENDOMEI1.hendosq_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 ROW_NUMBER()OVER(PARTITION BY bk_guid,hendo_kbn,kensin_ptn,hendo_guid,hy_guid ORDER BY kensin_ymd) AS 結合用 ";
                tmp_sql = tmp_sql + " 		,* ";
                tmp_sql = tmp_sql + " 	FROM hendodata_meisai ";
                tmp_sql = tmp_sql + " ) AS HENDOMEI2 ";
                tmp_sql = tmp_sql + " ON  HENDOMEI1.bk_guid = HENDOMEI2.bk_guid ";
                tmp_sql = tmp_sql + " AND HENDOMEI1.hendo_kbn = HENDOMEI2.hendo_kbn ";
                tmp_sql = tmp_sql + " AND HENDOMEI1.kensin_ptn = HENDOMEI2.kensin_ptn ";
                tmp_sql = tmp_sql + " AND HENDOMEI1.hendo_guid = HENDOMEI2.hendo_guid ";
                tmp_sql = tmp_sql + " AND HENDOMEI1.hy_guid = HENDOMEI2.hy_guid ";
                tmp_sql = tmp_sql + " AND HENDOMEI1.結合用 = HENDOMEI2.結合用 ";
                tmp_sql = tmp_sql + " ; ";

                return tmp_sql;

            }

            /// <summary>
            /// 部屋guid、変動費区分、入金項目No、部屋変動guidを取得 '20160603 ユーザーデータ検証による修正
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hyguid"></param>
            /// <param name="hendokbn"></param>
            /// <param name="nkinno"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            private string Get_HendoGuid(SqlConnection sqlcnnv10, string hyguid, string hendokbn, string nkinno)
            {

                string rtn_guid = "";
                string tmp_sql = "";

                // 抽出条件が不足している場合は処理を抜ける
                if (string.IsNullOrEmpty(hyguid.Trim()) | string.IsNullOrEmpty(hendokbn.Trim()) | string.IsNullOrEmpty(nkinno.Trim()))
                {
                    return rtn_guid;
                }

                tmp_sql = tmp_sql + " SELECT CONVERT(VARCHAR(50),hyhendo_guid) FROM hydata_hendo ";
                tmp_sql = tmp_sql + " WHERE hy_guid = '" + hyguid + "'";
                tmp_sql = tmp_sql + " AND   hendo_kbn = " + hendokbn;
                tmp_sql = tmp_sql + " AND   nkin_no = " + nkinno;

                rtn_guid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));

                return rtn_guid;

            }


        }

    }

    #endregion

    #region 家主固定控除情報

    public class Koteirule_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="syorikomok"></param>
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Koteirule_Model();               // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 3;                                   // サブキー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                // 'Excelファイル初期設定                  
                // Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                // Dim readtbl As New DataTable()
                // Dim con_read As New OleDbConnection()

                // 'オープン処理
                // rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                // 'オープン処理失敗時は処理を抜ける
                // If rtn = False Then
                // excelfile.ExcelFile_ReadClose(con_read)
                // Return rtn
                // End If

                // '行数取得
                // Dim rowcnt As Integer = readtbl.Rows.Count

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
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname_base = "bkdata";
                string tblname = "koteirule";
                string fldnamegrp = "koteirule_guid,bk_guid,nkin_sortorder,nkin_no,nkin_name," + "hy_guid,sosai_flg,gak,zei_kbn,zei_gak," + "nkbn_yotei,yatin_kozano,gtstart_ym,gtend_ym,tateazu_flg," + "sq_mmkbn,sq_ptn,sq_interval,sq_nen,sq_tuki," + "zei_rit,biko,history,sq_simekbn,ryosyu_flg";




                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, tblname_base, ref hash_guid);

                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, tblname_base, ref list_basekeydata);

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

                // 20160624 請求関連情報 親データ無しログ出力制御修正 -add
                int chg_totalcnt = rowcnt;        // 親データ無しのデータ件数調整用(カウントしないため)

                // ************************
                // 処理開始
                // ************************

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol_main - 1].ColumnName.Trim();
                string fldname_keysub = readtbl.Columns[keycol_sub - 1].ColumnName.Trim();

                // ---------------
                // データ部処理
                // ---------------
                for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
                {

                    // 中断処理
                    Application.DoEvents();
                    if (CommonModule.CancelFlg)
                    {
                        // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                        // excelfile.ExcelFile_ReadClose(con_read)
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

                    // 作業用変数
                    string tmp_sqkankaku = "";

                    // データ取得
                    for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                    {

                        // 項目名取得
                        fldname = readtbl.Columns[cntjj].ColumnName.Trim();

                        // 登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();

                        // 各項目値→変数格納
                        switch (fldname ?? "")
                        {
                            case "物件No":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue;
                                    break;
                                }
                            case "入金項目ソートNo":
                                {
                                    model_cvitem.Vari_Nkin_sortorder = fldvalue;
                                    break;
                                }
                            case "入金項目名":
                                {
                                    string argtukikbn = "9";
                                    model_cvitem.Vari_Nkin_no = fldvalue + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn);
                                    model_cvitem.Vari_Nkin_name = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    model_cvitem.Vari_Hy_guid = fldvalue;
                                    break;
                                }
                            case "相殺フラグ":
                                {
                                    model_cvitem.Vari_Sosai_flg = fldvalue;
                                    break;
                                }
                            case "額":
                                {
                                    model_cvitem.Vari_Gak = fldvalue;
                                    break;
                                }
                            case "税区分":
                                {
                                    model_cvitem.Vari_Zei_kbn = fldvalue;
                                    break;
                                }
                            case "税額":
                                {
                                    model_cvitem.Vari_Zei_gak = fldvalue;
                                    break;
                                }
                            case "入金区分":
                                {
                                    model_cvitem.Vari_Nkbn_yotei = fldvalue;
                                    break;
                                }
                            case "家賃入金口座No":
                                {
                                    model_cvitem.Vari_Yatin_kozano = fldvalue;
                                    break;
                                }
                            case "適用開始該当年月":
                                {
                                    model_cvitem.Vari_Gtstart_ym = fldvalue;
                                    break;
                                }
                            case "適用終了該当年月":
                                {
                                    model_cvitem.Vari_Gtend_ym = fldvalue;
                                    break;
                                }
                            case "立替・預りフラグ":
                                {
                                    model_cvitem.Vari_Tateazu_flg = fldvalue;
                                    break;
                                }
                            case "控除請求月区分":
                                {
                                    model_cvitem.Vari_Sq_mmkbn = fldvalue;
                                    break;
                                }
                            case "控除請求発生パターン":
                                {
                                    model_cvitem.Vari_Sq_ptn = fldvalue;
                                    break;
                                }
                            case "控除請求発生間隔":
                                {
                                    tmp_sqkankaku = fldvalue;
                                    break;
                                }
                            case "控除請求発生月指定年区分":
                                {
                                    model_cvitem.Vari_Sq_nen = fldvalue;
                                    break;
                                }
                            case "控除請求発生月指定月区分":
                                {
                                    model_cvitem.Vari_Sq_tuki = fldvalue;
                                    break;
                                }
                            case "適用税率":
                                {
                                    model_cvitem.Vari_Zei_rit = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            case "請求締区分":
                                {
                                    model_cvitem.Vari_Sq_simekbn = fldvalue;
                                    break;
                                }
                            case "領収フラグ":
                                {
                                    model_cvitem.Vari_Ryosyu_flg = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 変動費で固定項目の請求間隔、請求発生年、発生月を取得するために必要なデータを格納しておく
                    // 必要な情報
                    // 移行項目(部屋と契約入金項目情報で移行内容が異なるため)、請求パターン、請求月、請求開始月、請求間隔
                    model_cvitem.Vari_Sq_ptn = sheetname + "-" + model_cvitem.Vari_Sq_mmkbn + "-" + model_cvitem.Vari_Gtstart_ym + "-" + tmp_sqkankaku;
                    model_cvitem.Vari_Sq_interval = sheetname + "-" + model_cvitem.Vari_Sq_mmkbn + "-" + model_cvitem.Vari_Gtstart_ym + "-" + tmp_sqkankaku;
                    model_cvitem.Vari_Sq_nen = sheetname + "-" + model_cvitem.Vari_Sq_mmkbn + "-" + model_cvitem.Vari_Gtstart_ym + "-" + tmp_sqkankaku;
                    model_cvitem.Vari_Sq_tuki = sheetname + "-" + model_cvitem.Vari_Sq_mmkbn + "-" + model_cvitem.Vari_Gtstart_ym + "-" + tmp_sqkankaku;

                    // 固定値
                    model_cvitem.Vari_Koteirule_guid = Guid.NewGuid().ToString();
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
                        hash_cvitem["bk_guid"] = hash_guid[hash_cvitem["bk_guid"]];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                        }
                    }

                    // 20160624 請求関連情報 親データ無しログ出力制御修正 -add sta
                    else if (hash_log.Count == 1)
                    {
                        bool logclearflg = false;
                        foreach (var erritem in hash_log)
                        {
                            string tmp_str = Conversions.ToString(erritem.Value);
                            string[] tmp_errnaiyo = tmp_str.Split('-');
                            if ((tmp_errnaiyo[0] ?? "") == CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE)
                            {
                                chg_totalcnt = chg_totalcnt - 1;
                                logclearflg = true;
                            }
                        }
                        if (logclearflg)
                        {
                            hash_log.Clear();
                        }
                        // 20160624 請求関連情報 親データ無しログ出力制御修正 -add end

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
                    // Dim pgbcnt As Integer = 0
                    // If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
                    // pgbcnt = cntii
                    // ElseIf rowcnt > cntii + 1 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    // Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                    // ElseIf rowcnt <= cntii + 1 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                    // pgbcnt = pgbtotalcnt
                    // End If

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
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = chg_totalcnt;
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg end

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                // 'クローズ処理
                // excelfile.ExcelFile_ReadClose(con_read)
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                // 返却
                return rtn;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

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

                string tmp_sql = " SELECT bk_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_basedata);

            }

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT bk_no FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            public string Get_UseQry(ref string sortstr)
            {
                return default;

            }

            /// <summary>
            /// guid取得→ハッシュテーブル格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="hash_guid"></param>
            /// <remarks></remarks>
            public void Get_Guid(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_guid)
            {

                string tmp_sql = "SELECT bk_no,bk_guid FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);

            }

        }

    }

    #endregion

    #region 家主請求控除情報

    public class Kjdata_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="syorikomok"></param>
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kjdata_Model();                  // 移行値格納用モデル初期化
                // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                // Dim keycol_main As Integer = 1                                  'メインキー列
                // Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
                // Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
                // Dim keycol_sub3 As Integer = 4                                  'サブ3キー列
                // Dim keycol_sub4 As Integer = 8                                  'サブ4キー列
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 3;                                  // サブ1キー列
                int keycol_sub2 = 6;                                  // サブ2キー列
                // 20160601 家主請求控除情報の取得方法を変更 -chg end

                string viewnamebktobkguid = CommonModule.PRE_VIEW_NAME + "物件キー情報";
                // 20160601 家主請求控除情報の取得方法を変更 -del sta
                // Dim viewnamebkhytohyguid As String = PRE_VIEW_NAME & "物件部屋キー情報"
                // Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"
                // 20160601 家主請求控除情報の取得方法を変更 -del end

                // 20160601 家主請求控除情報の取得方法を変更 -del sta
                // '物件guidと送金ルール情報を紐付けた値の格納用
                // Dim hash_bktosoruleguid As New SafeDictionary<string, string>
                // Dim hash_bktosoruleno As New SafeDictionary<string, string>
                // Dim hash_bktosorulesoguid As New SafeDictionary<string, string>
                // Dim hash_bktosorulesono As New SafeDictionary<string, string>
                // 20160601 家主請求控除情報の取得方法を変更 -del end

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebktobkguid);
                // 20160601 家主請求控除情報の取得方法を変更 -del sta
                // Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyguid)
                // Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid)
                // 20160601 家主請求控除情報の取得方法を変更 -del end

                // テーブル名/フィールド名セット
                string tblname = "kjdata";
                // 20160519 EXEUpdateに伴う修正 家主請求控除情報 -add(最後尾にkojorule_taisyokbnを追加)
                string fldnamegrp = "kj_guid,kjkmk_guid,bk_guid,hy_guid,kj_yoteiymd," + "kjsort_no,gt_ym,nkin_no,dispnkin_name,kj_gak," + "kj_zeikbn,kj_zeigak,lock_flg,ryosyu_flg,tateazu_flg," + "tateazu_guid,bunkatu_guid,koteirule_guid,kojorule_flg,so_kakuymd," + "sokotik_guid,sorule_guid,sorule_no,sosaki_sorule_guid,sosaki_sorule_no," + "sosaki_no,sokoza_no,sikikinzuiji_guid,ky_guid,ky_recno," + "sokaisyu_guid,kanritesu_kbn,kanritesu_cyosyukbn,kanrigak_rit,szen_kbn," + "szen_no,szen_sqno,siwake_flg,zei_rit,biko," + "kj_torikomiymd,history,kubungrp_guid,kojorule_taisyokbn";








                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebktobkguid, ref hash_bkguid);
                // 20160601 家主請求控除情報の取得方法を変更 -del sta
                // Call Me.Get_Guid(sqlcnnv10, viewnamebkhytohyguid, hash_bkhyguid)
                // Call Me.Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, hash_bkhykyguid)
                // 20160601 家主請求控除情報の取得方法を変更 -del end

                // 親マスタ取得
                // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                Get_BaseKey(sqlcnnv10, viewnamebktobkguid, ref list_basekeydata);
                // 20160601 家主請求控除情報の取得方法を変更 -chg end

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

                // 20160624 請求関連情報 親データ無しログ出力制御修正 -add
                int chg_totalcnt = rowcnt;        // 親データ無しのデータ件数調整用(カウントしないため)

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得

                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));
                // 20160601 家主請求控除情報の取得方法を変更 -del sta
                // Dim fldname_keysub3 As String = headervalue(startrow - 1, keycol_sub3)
                // Dim fldname_keysub4 As String = headervalue(startrow - 1, keycol_sub4)
                // 20160601 家主請求控除情報の取得方法を変更 -del end

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
                    // 20160601 家主請求控除情報の取得方法を変更 -del sta
                    // Dim tmp_keysub3 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub3))
                    // Dim tmp_keysub4 As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub4))
                    // 20160601 家主請求控除情報の取得方法を変更 -del end

                    // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                    // Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2 & "-" & tmp_keysub3
                    // .Vari_Bk_guid = tmp_keymain
                    // .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1
                    // .Vari_Ky_guid = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
                    string fldvalue_key = tmp_keymain;
                    // 20160601 家主請求控除情報の取得方法を変更 -chg end

                    // ログ出力用
                    // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                    // Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & _
                    // fldname_keysub1 & " = " & tmp_keysub1 & "、" & _
                    // fldname_keysub2 & " = " & tmp_keysub2 & "、" & _
                    // fldname_keysub3 & " = " & tmp_keysub3 & "、" & _
                    // fldname_keysub4 & " = " & tmp_keysub4
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

                    // 20160601 家主請求控除情報の取得方法を変更 -chg end

                    // 作業用変数
                    // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                    // Dim tmp_kystaymd As String = ""
                    // Dim tmp_kyendymd As String = ""
                    string tmp_bkkanrino = "";
                    string tmp_sono = "";
                    string tmp_sokozano = "";
                    // 20160601 家主請求控除情報の取得方法を変更 -chg end

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
                            // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                            // Case "物件No"
                            // '.Vari_Bk_guid = fldvalue.Trim
                            // '取得済み
                            // Case "部屋No"
                            // '.Vari_Hy_guid = fldvalue.Trim
                            // '取得済み
                            // Case "契約No"
                            // '.Vari_Ky_guid = fldvalue.Trim
                            // '取得済み
                            // Case "契約レコードNo"
                            // .Vari_Ky_recno = fldvalue.Trim
                            // Case "控除予定日"
                            // .Vari_Kj_yoteiymd = fldvalue
                            // Case "項目ソートNo"
                            // .Vari_Kjsort_no = fldvalue
                            // Case "該当年月"
                            // .Vari_Gt_ym = fldvalue
                            // Case "入金項目名"
                            // .Vari_Nkin_no = fldvalue & "-" & EtcMethod.Get_Nkinruiname("9")
                            // .Vari_Dispnkin_name = fldvalue
                            // Case "控除額"
                            // .Vari_Kj_gak = fldvalue
                            // Case "控除税区分"
                            // .Vari_Kj_zeikbn = fldvalue
                            // Case "控除税額"
                            // .Vari_Kj_zeigak = fldvalue
                            // Case "更新ロックフラグ"
                            // .Vari_Lock_flg = fldvalue
                            // Case "領収扱いフラグ"
                            // .Vari_Ryosyu_flg = fldvalue
                            // Case "立替・預りフラグ"
                            // .Vari_Tateazu_flg = fldvalue
                            // Case "送金（控除）確定日"
                            // .Vari_So_kakuymd = fldvalue
                            // Case "(控除先指定時）送金先No"
                            // .Vari_Sosaki_no = fldvalue
                            // Case "(控除先指定時）送金先口座No"
                            // .Vari_Sokoza_no = fldvalue
                            // Case "管理手数料区分"
                            // .Vari_Kanritesu_kbn = fldvalue
                            // Case "管理手数料請求条件"
                            // .Vari_Kanritesu_cyosyukbn = fldvalue
                            // Case "管理手数料率"
                            // .Vari_Kanrigak_rit = fldvalue
                            // Case "修繕区分"
                            // .Vari_Szen_kbn = fldvalue
                            // Case "随時修繕No"
                            // .Vari_Szen_no = fldvalue
                            // Case "修繕請求No"
                            // .Vari_Szen_sqno = fldvalue
                            // Case "仕訳フラグ"
                            // .Vari_Siwake_flg = fldvalue
                            // Case "適用税率"
                            // .Vari_Zei_rit = fldvalue
                            // Case "備考"
                            // .Vari_Biko = fldvalue
                            // Case "控除データ取込日"
                            // .Vari_Kj_torikomiymd = fldvalue
                            // Case "契約開始日"
                            // tmp_kystaymd = fldvalue
                            // Case "契約終了日"
                            // tmp_kyendymd = fldvalue

                            case "物件No":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
                            case "部屋No":
                                {
                                    model_cvitem.Vari_Hy_guid = fldvalue.Trim();
                                    break;
                                }
                            case "契約No":
                                {
                                    model_cvitem.Vari_Ky_guid = fldvalue.Trim();
                                    break;
                                }
                            case "契約レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "控除予定日":
                                {
                                    model_cvitem.Vari_Kj_yoteiymd = fldvalue;
                                    break;
                                }
                            case "項目ソートNo":
                                {
                                    model_cvitem.Vari_Kjsort_no = fldvalue;
                                    break;
                                }
                            case "該当年月":
                                {
                                    model_cvitem.Vari_Gt_ym = fldvalue;
                                    break;
                                }
                            case "入金項目名":
                                {
                                    string argtukikbn = "9";
                                    model_cvitem.Vari_Nkin_no = fldvalue + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn);
                                    break;
                                }
                            case "表示用入金項目名":
                                {
                                    model_cvitem.Vari_Dispnkin_name = fldvalue;
                                    break;
                                }
                            case "控除額":
                                {
                                    model_cvitem.Vari_Kj_gak = fldvalue;
                                    break;
                                }
                            case "控除税区分":
                                {
                                    model_cvitem.Vari_Kj_zeikbn = fldvalue;
                                    break;
                                }
                            case "控除税額":
                                {
                                    model_cvitem.Vari_Kj_zeigak = fldvalue;
                                    break;
                                }
                            case "更新ロックフラグ":
                                {
                                    model_cvitem.Vari_Lock_flg = fldvalue;
                                    break;
                                }
                            case "領収扱いフラグ":
                                {
                                    model_cvitem.Vari_Ryosyu_flg = fldvalue;
                                    break;
                                }
                            case "立替・預りフラグ":
                                {
                                    model_cvitem.Vari_Tateazu_flg = fldvalue;
                                    break;
                                }
                            case "送金（控除）確定日":
                                {
                                    model_cvitem.Vari_So_kakuymd = fldvalue;
                                    break;
                                }
                            case "(控除先指定時）送金先No":
                                {
                                    model_cvitem.Vari_Sosaki_no = fldvalue;
                                    break;
                                }
                            case "(控除先指定時）送金先口座No":
                                {
                                    model_cvitem.Vari_Sokoza_no = fldvalue;
                                    break;
                                }
                            case "管理手数料区分":
                                {
                                    model_cvitem.Vari_Kanritesu_kbn = fldvalue;
                                    break;
                                }
                            case "管理手数料請求条件":
                                {
                                    model_cvitem.Vari_Kanritesu_cyosyukbn = fldvalue;
                                    break;
                                }
                            case "管理手数料率":
                                {
                                    model_cvitem.Vari_Kanrigak_rit = fldvalue;
                                    break;
                                }
                            case "修繕区分":
                                {
                                    model_cvitem.Vari_Szen_kbn = fldvalue;
                                    break;
                                }
                            case "随時修繕No":
                                {
                                    model_cvitem.Vari_Szen_no = fldvalue;
                                    break;
                                }
                            case "修繕請求No":
                                {
                                    model_cvitem.Vari_Szen_sqno = fldvalue;
                                    break;
                                }
                            case "仕訳フラグ":
                                {
                                    model_cvitem.Vari_Siwake_flg = fldvalue;
                                    break;
                                }
                            case "適用税率":
                                {
                                    model_cvitem.Vari_Zei_rit = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            case "控除データ取込日":
                                {
                                    model_cvitem.Vari_Kj_torikomiymd = fldvalue;
                                    break;
                                }
                            case "物件管理No":
                                {
                                    model_cvitem.Vari_Sorule_no = fldvalue;
                                    model_cvitem.Vari_Sosaki_sorule_no = fldvalue;
                                    tmp_bkkanrino = fldvalue;
                                    break;
                                }
                            case "送金先No":
                                {
                                    tmp_sono = fldvalue;
                                    break;
                                }
                            case "送金先口座No":
                                {
                                    tmp_sokozano = fldvalue;
                                    break;
                                }
                                // 20160601 家主請求控除情報の取得方法を変更 -chg end
                        }

                    }

                    // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                    // '物件Noと送金ルール情報を紐付けた値を取得
                    // Dim ky_endymd As String = tmp_kyendymd                '送金ルール取得用に契約終了日を取得
                    // ky_endymd = ky_endymd.Replace("-", "")
                    // ky_endymd = ky_endymd.Replace(" 00:00:00.000", "")                  '成形 (念の為)
                    // hash_bktosoruleguid = Me.Get_BkToSoruleInfo(sqlcnnv10, ky_endymd, 1)
                    // hash_bktosoruleno = Me.Get_BkToSoruleInfo(sqlcnnv10, ky_endymd, 2)
                    // hash_bktosorulesoguid = Me.Get_BkToSoruleInfo(sqlcnnv10, ky_endymd, 3)
                    // hash_bktosorulesono = Me.Get_BkToSoruleInfo(sqlcnnv10, ky_endymd, 4)

                    // '送金ルールから取得
                    // .Vari_Sorule_guid = IIf(hash_bktosoruleguid.Item(.Vari_Bk_guid) Is Nothing, "", hash_bktosoruleguid.Item(.Vari_Bk_guid))
                    // .Vari_Sorule_no = IIf(hash_bktosoruleno.Item(.Vari_Bk_guid) Is Nothing, "", hash_bktosoruleno.Item(.Vari_Bk_guid))
                    // .Vari_Sosaki_sorule_guid = IIf(hash_bktosorulesoguid.Item(.Vari_Bk_guid) Is Nothing, "", hash_bktosorulesoguid.Item(.Vari_Bk_guid))
                    // .Vari_Sosaki_sorule_no = IIf(hash_bktosorulesono.Item(.Vari_Bk_guid) Is Nothing, "", hash_bktosorulesono.Item(.Vari_Bk_guid))

                    // 物件No、送金管理Noから送金ルールguidを取得
                    model_cvitem.Vari_Sorule_guid = Get_SoruleGuid(sqlcnnv10, model_cvitem.Vari_Bk_guid, tmp_bkkanrino);

                    // 物件No、送金管理No、送金先口座Noから送金ルールguidを取得
                    model_cvitem.Vari_Sosaki_sorule_guid = Get_SoruleSosakiGuid(sqlcnnv10, model_cvitem.Vari_Bk_guid, tmp_bkkanrino, tmp_sono, tmp_sokozano);

                    // 20160601 家主請求控除情報の取得方法を変更 -chg end
                    // 固定値
                    model_cvitem.Vari_Kj_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Kjkmk_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Tateazu_guid = "";
                    model_cvitem.Vari_Bunkatu_guid = Guid.NewGuid().ToString();
                    // 20160601 家主請求控除情報の取得方法を変更 -del
                    // .Vari_Koteirule_guid = ""
                    model_cvitem.Vari_Kojorule_flg = 0.ToString();
                    // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                    // .Vari_Sokotik_guid = Guid.NewGuid.ToString
                    model_cvitem.Vari_Sokotik_guid = "";
                    // 20160601 家主請求控除情報の取得方法を変更 -chg end
                    model_cvitem.Vari_Sikikinzuiji_guid = "";
                    model_cvitem.Vari_Sokaisyu_guid = "";
                    model_cvitem.Vari_Kubungrp_guid = "";
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Kojorule_taisyokbn = "";   // 20160601 家主請求控除情報の取得方法を変更 -add

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
                        hash_cvitem["bk_guid"] = hash_bkguid[hash_cvitem["bk_guid"]];
                        // 20160601 家主請求控除情報の取得方法を変更 -del sta
                        // hash_cvitem("hy_guid") = hash_bkhyguid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // 20160601 家主請求控除情報の取得方法を変更 -del end

                        // 20160601 家主請求控除情報の取得方法を変更 -add sta
                        // 移行用に変換された物件guid、入金項目Noおよび該当年月から固定控除ルールのguidを取得
                        string tmp_bkguid = Conversions.ToString(hash_cvitem["bk_guid"]);
                        string tmp_nkinno = Conversions.ToString(hash_cvitem["nkin_no"]);
                        string tmp_gtym = Conversions.ToString(hash_cvitem["kj_yoteiymd"]);
                        string tmp_kjguid = Get_KoteiruleGuid(sqlcnnv10, tmp_bkguid, tmp_nkinno, tmp_gtym);
                        hash_cvitem["koteirule_guid"] = tmp_kjguid;
                        // 20160601 家主請求控除情報の取得方法を変更 -add end

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                        }
                    }

                    // 20160624 請求関連情報 親データ無しログ出力制御修正 -add sta
                    else if (hash_log.Count == 1)
                    {
                        bool logclearflg = false;
                        foreach (var erritem in hash_log)
                        {
                            string tmp_str = Conversions.ToString(erritem.Value);
                            string[] tmp_errnaiyo = tmp_str.Split('-');
                            if ((tmp_errnaiyo[0] ?? "") == CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE)
                            {
                                chg_totalcnt = chg_totalcnt - 1;
                                logclearflg = true;
                            }
                        }
                        if (logclearflg)
                        {
                            hash_log.Clear();
                        }
                        // 20160624 請求関連情報 親データ無しログ出力制御修正 -add end

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
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = chg_totalcnt;
                // 20160624 請求関連情報 親データ無しログ出力制御修正 -chg end

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebktobkguid);
                // 20160601 家主請求控除情報の取得方法を変更 -del sta
                // Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid)
                // Call Me.Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid)
                // 20160601 家主請求控除情報の取得方法を変更 -del end

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN kydata ON " + tblname + ".ky_guid = kydata.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// guid取得→ハッシュテーブル格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="hash_guid"></param>
            /// <remarks></remarks>
            public void Get_Guid(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_guid)
            {

                string tmp_sql = "SELECT * FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);

            }

            /// <summary>
            /// キー/guidをセットにしたVIEWの作成
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <remarks></remarks>
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                // VIEW初期化
                Drop_TmpView(sqlcnnv10, viewname);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT bk_no AS キー,bk_guid FROM bkdata ";
                            break;
                        }
                        // 20160601 家主請求控除情報の取得方法を変更 -del sta
                        // Case "物件部屋キー情報"
                        // tmp_sql_create = tmp_sql_create & " SELECT "
                        // tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
                        // tmp_sql_create = tmp_sql_create & " 	,hy_guid "
                        // tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
                        // tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
                        // Case "物件部屋契約キー情報"
                        // tmp_sql_create = tmp_sql_create & " SELECT "
                        // tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー "
                        // tmp_sql_create = tmp_sql_create & " 	,ky_guid "
                        // tmp_sql_create = tmp_sql_create & " FROM kydata AS KY "
                        // tmp_sql_create = tmp_sql_create & " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid "
                        // tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid "
                        // 20160601 家主請求控除情報の取得方法を変更 -del end
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 作業用VIEWの削除
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="viewname"></param>
            /// <remarks></remarks>
            public void Drop_TmpView(SqlConnection sqlcnnv10, string viewname)
            {

                int tmpcnt = 0;

                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

            }

            // 20160601 家主請求控除情報の取得方法を変更 -del sta
            // ''' <summary>
            // ''' 物件Noと送金ルール情報を紐付けたハッシュテーブルを作成
            // ''' </summary>
            // ''' <param name="sqlcnnv10"></param>
            // ''' <param name="typeno"></param>
            // ''' <remarks></remarks>
            // Private Function Get_BkToSoruleInfo(ByVal sqlcnnv10 As SqlConnection, ByVal soymd As String, ByVal typeno As Integer) As SafeDictionary<string, string>

            // Dim rtn_hash As New SafeDictionary<string, string>
            // Dim tmp_sql As String = ""

            // Select Case typeno
            // Case 1
            // tmp_sql = " SELECT bk_no,sorule_guid FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid "
            // tmp_sql = tmp_sql & " WHERE sorule_startymd <= '" & soymd & "' AND ISNULL(sorule_endymd,'21001231') >=  '" & soymd & "'"
            // Case 2
            // tmp_sql = "SELECT bk_no,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid"
            // tmp_sql = tmp_sql & " WHERE sorule_startymd <= '" & soymd & "' AND ISNULL(sorule_endymd,'21001231') >=  '" & soymd & "'"
            // Case 3
            // tmp_sql = tmp_sql & " SELECT bk_no,SOSAKI.sorule_guid FROM "
            // tmp_sql = tmp_sql & " ( "
            // tmp_sql = tmp_sql & " 	SELECT bk_no,sorule_guid,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid "
            // tmp_sql = tmp_sql & " 	WHERE sorule_startymd <= '" & soymd & "' AND ISNULL(sorule_endymd,'21001231') >=  '" & soymd & "'"
            // tmp_sql = tmp_sql & " ) AS SO "
            // tmp_sql = tmp_sql & " LEFT JOIN "
            // tmp_sql = tmp_sql & " 	( "
            // tmp_sql = tmp_sql & " 		SELECT DISTINCT sorule_guid,sorule_no FROM sorule_sosaki "
            // tmp_sql = tmp_sql & " 	) AS SOSAKI "
            // tmp_sql = tmp_sql & " ON  SO.sorule_guid = SOSAKI.sorule_guid "
            // tmp_sql = tmp_sql & " AND SO.sorule_no = SOSAKI.sorule_no "
            // Case 4
            // tmp_sql = tmp_sql & " SELECT bk_no,SOSAKI.sorule_no FROM "
            // tmp_sql = tmp_sql & " ( "
            // tmp_sql = tmp_sql & " 	SELECT bk_no,sorule_guid,sorule_no FROM sorule AS SO LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid "
            // tmp_sql = tmp_sql & " 	WHERE sorule_startymd <= '" & soymd & "' AND ISNULL(sorule_endymd,'21001231') >=  '" & soymd & "'"
            // tmp_sql = tmp_sql & " ) AS SO "
            // tmp_sql = tmp_sql & " LEFT JOIN "
            // tmp_sql = tmp_sql & " 	( "
            // tmp_sql = tmp_sql & " 		SELECT DISTINCT sorule_guid,sorule_no FROM sorule_sosaki "
            // tmp_sql = tmp_sql & " 	) AS SOSAKI "
            // tmp_sql = tmp_sql & " ON  SO.sorule_guid = SOSAKI.sorule_guid "
            // tmp_sql = tmp_sql & " AND SO.sorule_no = SOSAKI.sorule_no "
            // End Select

            // DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, rtn_hash)

            // Return rtn_hash

            // End Function
            // 20160601 家主請求控除情報の取得方法を変更 -del end

            /// <summary>
            /// 物件No、物件管理Noから送金ルールguidを取得 '20160601 家主請求控除情報の取得方法を変更
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="bkno"></param>
            /// <param name="knno"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            private string Get_SoruleGuid(SqlConnection sqlcnnv10, string bkno, string knno)
            {

                string rtn_guid = "";
                string tmp_sql = "";

                tmp_sql = tmp_sql + " SELECT CONVERT(VARCHAR(50),sorule_guid) FROM sorule AS SO ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " WHERE bk_no = " + bkno;
                tmp_sql = tmp_sql + " AND   sorule_no = " + knno;

                rtn_guid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));

                return rtn_guid;

            }

            /// <summary>
            /// 物件No、物件管理No、送金先No、送金先口座Noから送金ルールguidを取得 '20160601 家主請求控除情報の取得方法を変更
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="bkno"></param>
            /// <param name="knno"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            private string Get_SoruleSosakiGuid(SqlConnection sqlcnnv10, string bkno, string knno, string sono, string sokozano)
            {

                string rtn_guid = "";
                string tmp_sql = "";

                tmp_sql = tmp_sql + " SELECT CONVERT(VARCHAR(50),SOSAKI.sorule_guid) FROM sorule_sosaki AS SOSAKI ";
                tmp_sql = tmp_sql + " LEFT JOIN sorule AS SO ON SOSAKI.sorule_guid = SO.sorule_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " WHERE bk_no = " + bkno;
                tmp_sql = tmp_sql + " AND   SOSAKI.sorule_no = " + knno;
                tmp_sql = tmp_sql + " AND   so_ow_no = " + sono;
                tmp_sql = tmp_sql + " AND   so_ow_kozano = " + sokozano;

                rtn_guid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));

                return rtn_guid;

            }

            /// <summary>
            /// 物件guid、入金項目No、控除予定日から固定控除ルールguidを取得 '20160601 家主請求控除情報の取得方法を変更
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="bkguid"></param>
            /// <param name="nkinno"></param>
            /// <param name="kjym"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            private string Get_KoteiruleGuid(SqlConnection sqlcnnv10, string bkguid, string nkinno, string kjym)
            {

                string rtn_guid = "";
                string tmp_sql = "";

                // 20160603 固定控除ルール取得方法の修正 -add sta
                // 抽出条件が不足している場合は処理を抜ける
                if (string.IsNullOrEmpty(bkguid.Trim()) | string.IsNullOrEmpty(nkinno.Trim()) | string.IsNullOrEmpty(kjym))
                {
                    return rtn_guid;
                }
                // 20160603 固定控除ルール取得方法の修正 -add end

                tmp_sql = tmp_sql + " SELECT CONVERT(VARCHAR(50),koteirule_guid) FROM koteirule ";
                tmp_sql = tmp_sql + " WHERE bk_guid = '" + bkguid + "'";
                tmp_sql = tmp_sql + " AND   nkin_no = " + nkinno;
                tmp_sql = tmp_sql + " AND   gtstart_ym <= '" + kjym + "' AND gtend_ym >= '" + kjym + "' ";

                rtn_guid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));

                return rtn_guid;

            }

        }

    }

    #endregion

}