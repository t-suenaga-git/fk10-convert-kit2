using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Repository
{

    #region クレーム基本情報

    public class Claimdata_Repository
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
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Claimdata_Model();               // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列
                var hash_bkguid = new SafeDictionary<string, string>();                                // キー/guid格納用ハッシュテーブル
                var hash_bkhyguid = new SafeDictionary<string, string>();                              // キー/guid格納用ハッシュテーブル
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                string viewnamebktobkguid = CommonModule.PRE_VIEW_NAME + "物件キー情報";
                string viewnamebkhytohyguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************

                // 10クレーム分類マスタから分類Noと分類名を紐付けたデータを取得
                EtcMethod.Set_ClaimBruiMst(sqlcnnv10, 1);
                EtcMethod.Set_ClaimBruiMst(sqlcnnv10, 2);

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
                string tblname = "claimdata";
                string fldnamegrp = "claim_no,status,title,uke_ymd,uke_logonuser_no," + "uke_name,uke_tel1,uke_tel2,renraku_timestart,renraku_timeend," + "kinkyu_kbn,taio_limitymd,kasyo_ruino,claim_ruino,uke_report," + "taio_logonuser_no,emailbiko,bk_guid,hy_guid,ky_guid," + "ky_recno,reform_guid,taiosaki_kbn,taiosaki_gy_no,taiosaki_tantoname," + "taiosaki_tel,taio_finishymd,taio_report,hutan1_kbn,hutan1_gaketc," + "hutan2_kbn,hutan2_gaketc,hutan3_kbn,hutan3_gaketc,hutanbiko," + "jisya_no,jisya_name,history,rowid,uke_hm," + "taio_finishhm,uke_namesjis,taiosaki_tantonamesjis,ow_no";








                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebktobkguid, ref hash_bkguid);
                Get_Guid(sqlcnnv10, viewnamebkhytohyguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

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

                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_key = Conversions.ToString(headervalue[startrow - 1, keycol]);

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
                    string fldvalue_key = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol]);

                    // ログ出力用
                    string str_logkey = fldname_key + " = " + fldvalue_key;

                    // 作業用変数
                    string tmp_bkno = "";
                    string tmp_hyno = "";
                    string tmp_kyno = "";

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue[startrow - 1, cntjj].ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "クレームNo":
                                {
                                    model_cvitem.Vari_Claim_no = fldvalue;
                                    break;
                                }
                            case "対応状況":
                                {
                                    model_cvitem.Vari_Status = fldvalue;
                                    break;
                                }
                            case "タイトル":
                                {
                                    model_cvitem.Vari_Title = fldvalue;
                                    break;
                                }
                            case "受付日時":
                                {
                                    model_cvitem.Vari_Uke_ymd = fldvalue;
                                    break;
                                }
                            case "受付担当者":
                                {
                                    model_cvitem.Vari_Uke_logonuser_no = fldvalue;
                                    break;
                                }
                            case "連絡者名":
                                {
                                    model_cvitem.Vari_Uke_name = fldvalue;
                                    break;
                                }
                            case "連絡者TEL1":
                                {
                                    model_cvitem.Vari_Uke_tel1 = fldvalue;
                                    break;
                                }
                            case "連絡者TEL2":
                                {
                                    model_cvitem.Vari_Uke_tel2 = fldvalue;
                                    break;
                                }
                            case "連絡可能時間（開始）":
                                {
                                    model_cvitem.Vari_Renraku_timestart = fldvalue;
                                    break;
                                }
                            case "連絡可能時間（終了）":
                                {
                                    model_cvitem.Vari_Renraku_timeend = fldvalue;
                                    break;
                                }
                            case "緊急度":
                                {
                                    model_cvitem.Vari_Kinkyu_kbn = fldvalue;
                                    break;
                                }
                            case "対応期限":
                                {
                                    model_cvitem.Vari_Taio_limitymd = fldvalue;
                                    break;
                                }
                            case "箇所分類No":
                                {
                                    model_cvitem.Vari_Kasyo_ruino = fldvalue;
                                    break;
                                }
                            case "クレーム分類No":
                                {
                                    model_cvitem.Vari_Claim_ruino = fldvalue;
                                    break;
                                }
                            case "内容":
                                {
                                    model_cvitem.Vari_Uke_report = fldvalue;
                                    break;
                                }
                            case "対応担当者":
                                {
                                    model_cvitem.Vari_Taio_logonuser_no = fldvalue;
                                    break;
                                }
                            case "電子メール備考":
                                {
                                    model_cvitem.Vari_Emailbiko = fldvalue;
                                    break;
                                }
                            case "物件No":
                                {
                                    // .Vari_Bk_guid = fldvalue
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    // .Vari_Hy_guid = fldvalue
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "契約No":
                                {
                                    // .Vari_Ky_guid = fldvalue
                                    tmp_kyno = fldvalue;
                                    break;
                                }
                            case "契約レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue;
                                    break;
                                }
                            case "対応先区分":
                                {
                                    model_cvitem.Vari_Taiosaki_kbn = fldvalue;
                                    break;
                                }
                            case "依頼業者No":
                                {
                                    model_cvitem.Vari_Taiosaki_gy_no = fldvalue;
                                    break;
                                }
                            case "業者担当者":
                                {
                                    model_cvitem.Vari_Taiosaki_tantoname = fldvalue;
                                    break;
                                }
                            case "業者担当者TEL":
                                {
                                    model_cvitem.Vari_Taiosaki_tel = fldvalue;
                                    break;
                                }
                            case "完了日時":
                                {
                                    model_cvitem.Vari_Taio_finishymd = fldvalue;
                                    break;
                                }
                            case "結果入力":
                                {
                                    model_cvitem.Vari_Taio_report = fldvalue;
                                    break;
                                }
                            case "負担者1":
                                {
                                    model_cvitem.Vari_Hutan1_kbn = fldvalue;
                                    break;
                                }
                            case "負担者1金額等":
                                {
                                    model_cvitem.Vari_Hutan1_gaketc = fldvalue;
                                    break;
                                }
                            case "負担者2":
                                {
                                    model_cvitem.Vari_Hutan2_kbn = fldvalue;
                                    break;
                                }
                            case "負担者2金額等":
                                {
                                    model_cvitem.Vari_Hutan2_gaketc = fldvalue;
                                    break;
                                }
                            case "負担者3":
                                {
                                    model_cvitem.Vari_Hutan3_kbn = fldvalue;
                                    break;
                                }
                            case "負担者3金額等":
                                {
                                    model_cvitem.Vari_Hutan3_gaketc = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Hutanbiko = fldvalue;
                                    break;
                                }
                            case "自社・支店No":
                                {
                                    model_cvitem.Vari_Jisya_no = fldvalue;
                                    break;
                                }
                            case "自社・支店名":
                                {
                                    model_cvitem.Vari_Jisya_name = fldvalue;
                                    break;
                                }
                            case "受付時刻":
                                {
                                    model_cvitem.Vari_Uke_hm = fldvalue;
                                    break;
                                }
                            case "対応終了時刻":
                                {
                                    model_cvitem.Vari_Taio_finishhm = fldvalue;
                                    break;
                                }
                            case "連絡者名SJIS":
                                {
                                    model_cvitem.Vari_Uke_namesjis = fldvalue;
                                    break;
                                }
                            case "対応担当者名SJIS":
                                {
                                    model_cvitem.Vari_Taiosaki_tantonamesjis = fldvalue;
                                    break;
                                }
                            case "家主No":
                                {
                                    model_cvitem.Vari_Ow_no = fldvalue;
                                    break;
                                }
                        }

                    }

                    // guid取得用に成形
                    model_cvitem.Vari_Bk_guid = tmp_bkno;
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;
                    model_cvitem.Vari_Ky_guid = tmp_bkno + "-" + tmp_hyno + "-" + tmp_kyno;

                    // 固定値
                    model_cvitem.Vari_Reform_guid = "00000000-0000-0000-0000-000000000000";  // 修繕関連は構築中のため現時点でダミーを設定しておく
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebktobkguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhytohyguid);
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 返却
                return rtn;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

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

                string tmp_sql = " SELECT claim_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

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
                            tmp_sql_create = tmp_sql_create + " SELECT * FROM ";
                            tmp_sql_create = tmp_sql_create + " ( ";
                            tmp_sql_create = tmp_sql_create + " 	SELECT ";
                            tmp_sql_create = tmp_sql_create + " 		 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 		,ky_guid ";
                            tmp_sql_create = tmp_sql_create + " 	FROM kydata AS KY ";
                            tmp_sql_create = tmp_sql_create + " 	LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " 	LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                            tmp_sql_create = tmp_sql_create + " ) AS VW ";
                            tmp_sql_create = tmp_sql_create + " WHERE [キー] IS NOT NULL ";
                            break;
                        }
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

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

    #region クレーム対応履歴情報

    public class Claim_taio_Repository
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
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Claim_taio_Model();              // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

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
                string tblname_base = "claimdata";
                string tblname = "claim_taio";
                string fldnamegrp = "claim_no,taio_no,taio_ymd,taio_logonuser_no,taio_report," + "history,taio_hm";

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub = Conversions.ToString(headervalue[startrow - 1, keycol_sub]);

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
                    string fldvalue_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string fldvalue_keysub = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub]);
                    string fldvalue_key = fldvalue_keymain + "-" + fldvalue_keysub;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + fldvalue_keymain + "、" + fldname_keysub + " = " + fldvalue_keysub;

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue[startrow - 1, cntjj].ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "クレームNo":
                                {
                                    model_cvitem.Vari_Claim_no = fldvalue;
                                    break;
                                }
                            case "対応履歴No":
                                {
                                    model_cvitem.Vari_Taio_no = fldvalue;
                                    break;
                                }
                            case "対応日時":
                                {
                                    model_cvitem.Vari_Taio_ymd = fldvalue;
                                    break;
                                }
                            case "対応担当者":
                                {
                                    model_cvitem.Vari_Taio_logonuser_no = fldvalue;
                                    break;
                                }
                            case "内容":
                                {
                                    model_cvitem.Vari_Taio_report = fldvalue;
                                    break;
                                }
                            case "対応時刻":
                                {
                                    model_cvitem.Vari_Taio_hm = fldvalue;
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
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT claim_no FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,claim_no) + '-' + CONVERT(varchar,taio_no) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region クレーム関連ファイル情報

    public class Claimdata_relfile_Repository
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
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Claimdata_relfile_Model();       // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

                // ************************
                // 作業準備
                // ************************

                // 20160525 クレーム関連ファイルの画像判別処理実装 -add sta
                // 画像ファイルの拡張子リストを作成
                EtcMethod.Set_ImportableImageFileAttributesList();
                // 20160525 クレーム関連ファイルの画像判別処理実装 -add end

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // テーブル名/フィールド名セット
                string tblname_base = "claimdata";
                string tblname = "claimdata_relfile";
                string fldnamegrp = "claim_no,file_no,fullpath,addtime,biko," + "history";

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

                // ************************
                // 処理開始
                // ************************

                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub = Conversions.ToString(headervalue[startrow - 1, keycol_sub]);

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
                    string fldvalue_keymain = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_main]);
                    string fldvalue_keysub = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub]);
                    string fldvalue_key = fldvalue_keymain + "-" + fldvalue_keysub;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + fldvalue_keymain + "、" + fldname_keysub + " = " + fldvalue_keysub;

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue[startrow - 1, cntjj].ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "クレームNo":
                                {
                                    model_cvitem.Vari_Claim_no = fldvalue;
                                    break;
                                }
                            case "ファイルNo":
                                {
                                    model_cvitem.Vari_File_no = fldvalue;
                                    break;
                                }
                            case "ファイルパス":
                                {
                                    model_cvitem.Vari_Fullpath = fldvalue;
                                    break;
                                }
                            case "追加日":
                                {
                                    model_cvitem.Vari_Addtime = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
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
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT claim_no FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,claim_no) + '-' + CONVERT(varchar,file_no) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

}