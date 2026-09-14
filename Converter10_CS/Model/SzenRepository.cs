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

    #region 修繕基本情報

    public class Szendata_Repository
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Szendata_Model();                // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                // Dim viewnamebkhykytokyguid As String = PRE_VIEW_NAME & "物件部屋契約キー情報"

                // ************************
                // 作業準備
                // ************************

                // 入金区分紐付情報取得
                SetRelItemToObject.Set_RelData_Nkinkbn();

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // guid取得
                Set_BkGuid(sqlcnnv10, ref hash_bkguid);
                Set_HyGuid(sqlcnnv10, ref hash_bkhyguid);
                Set_KyGuid(sqlcnnv10, ref hash_bkhykyguid);

                // テーブル名/フィールド名セット
                string tblname = "szendata";
                string fldnamegrp = "bk_guid,szen_no,hy_useflg,hy_guid,ky_useflg," + "ky_guid,szen_name,szen_ukeymd,szen_endymd,tanto_no," + "jisya_no,koji_yoteistartymd,koji_yoteiendymd,koji_basyo,koji_gaiyo," + "koji_kagi,koji_tatiai,biko,kys_futanrit,ow_futanrit," + "rowid,history,jisya_futanrit,kys_sqflg,kys_no," + "kys_gtym,kys_sqsimeymd,kys_nkbn,kys_sorit,kys_yatinkozano," + "kys_biko,ow_sqflg,ow_no,ow_kaisyukbn,ow_kojoymd," + "ow_kojogtym,ow_sosaki_soruleguid,ow_sosaki_sorule_no,ow_sokozano,ow_sqymd," + "ow_sqgtym,ow_sqsqsimeymd,ow_sqnkbn,ow_sqyatinkozano,ow_sqsakino," + "ow_biko";









                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_bkguid, ref list_basekeydata);

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
                    string tmp_bkno = "";
                    string tmp_hyno = "";
                    string tmp_kyno = "";

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
                                    model_cvitem.Vari_Bk_guid = fldvalue;
                                    tmp_bkno = fldvalue;
                                    break;
                                }
                            case "修繕No":
                                {
                                    model_cvitem.Vari_Szen_no = fldvalue;
                                    break;
                                }
                            case "部屋使用フラグ":
                                {
                                    model_cvitem.Vari_Hy_useflg = fldvalue;
                                    break;
                                }
                            case "部屋No":
                                {
                                    tmp_hyno = fldvalue;
                                    break;
                                }
                            case "契約使用フラグ":
                                {
                                    model_cvitem.Vari_Ky_useflg = fldvalue;
                                    break;
                                }
                            case "契約No":
                                {
                                    tmp_kyno = fldvalue;
                                    break;
                                }
                            case "修繕名":
                                {
                                    model_cvitem.Vari_Szen_name = fldvalue;
                                    break;
                                }
                            case "修繕受付日":
                                {
                                    model_cvitem.Vari_Szen_ukeymd = fldvalue;
                                    break;
                                }
                            case "修繕終了日":
                                {
                                    model_cvitem.Vari_Szen_endymd = fldvalue;
                                    break;
                                }
                            case "修繕担当者":
                                {
                                    model_cvitem.Vari_Tanto_no = fldvalue;
                                    break;
                                }
                            case "自社支店No":
                                {
                                    model_cvitem.Vari_Jisya_no = fldvalue;
                                    break;
                                }
                            case "工事開始予定日":
                                {
                                    model_cvitem.Vari_Koji_yoteistartymd = fldvalue;
                                    break;
                                }
                            case "工事終了予定日":
                                {
                                    model_cvitem.Vari_Koji_yoteiendymd = fldvalue;
                                    break;
                                }
                            case "工事場所":
                                {
                                    model_cvitem.Vari_Koji_basyo = fldvalue;
                                    break;
                                }
                            case "工事概要":
                                {
                                    model_cvitem.Vari_Koji_gaiyo = fldvalue;
                                    break;
                                }
                            case "鍵":
                                {
                                    model_cvitem.Vari_Koji_kagi = fldvalue;
                                    break;
                                }
                            case "立会者":
                                {
                                    model_cvitem.Vari_Koji_tatiai = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            case "契約者負担率":
                                {
                                    model_cvitem.Vari_Kys_futanrit = fldvalue;
                                    break;
                                }
                            case "家主負担率":
                                {
                                    model_cvitem.Vari_Ow_futanrit = fldvalue;
                                    break;
                                }
                            case "自社負担率":
                                {
                                    model_cvitem.Vari_Jisya_futanrit = fldvalue;
                                    break;
                                }
                            case "契約者請求作成フラグ":
                                {
                                    model_cvitem.Vari_Kys_sqflg = fldvalue;
                                    break;
                                }
                            case "契約者請求先No":
                                {
                                    model_cvitem.Vari_Kys_no = fldvalue;
                                    break;
                                }
                            case "契約者該当年月":
                                {
                                    model_cvitem.Vari_Kys_gtym = fldvalue;
                                    break;
                                }
                            case "契約者請求締日":
                                {
                                    model_cvitem.Vari_Kys_sqsimeymd = fldvalue;
                                    break;
                                }
                            case "契約者入金区分":
                                {
                                    model_cvitem.Vari_Kys_nkbn = fldvalue;
                                    break;
                                }
                            case "契約者送金率":
                                {
                                    model_cvitem.Vari_Kys_sorit = fldvalue;
                                    break;
                                }
                            case "契約者振込先口座No":
                                {
                                    model_cvitem.Vari_Kys_yatinkozano = fldvalue;
                                    break;
                                }
                            case "契約者請求備考":
                                {
                                    model_cvitem.Vari_Kys_biko = fldvalue;
                                    break;
                                }
                            case "家主請求作成フラグ":
                                {
                                    model_cvitem.Vari_Ow_sqflg = fldvalue;
                                    break;
                                }
                            case "家主控除先No":
                                {
                                    model_cvitem.Vari_Ow_no = fldvalue;
                                    break;
                                }
                            case "家主控除請求区分":
                                {
                                    model_cvitem.Vari_Ow_kaisyukbn = fldvalue;
                                    break;
                                }
                            case "家主控除予定日":
                                {
                                    model_cvitem.Vari_Ow_kojoymd = fldvalue;
                                    break;
                                }
                            case "家主控除該当年月":
                                {
                                    model_cvitem.Vari_Ow_kojogtym = fldvalue;
                                    break;
                                }
                            case "家主控除先送金ルールNo":
                                {
                                    model_cvitem.Vari_Ow_sosaki_sorule_no = fldvalue;
                                    break;
                                }
                            case "家主控除先口座No":
                                {
                                    model_cvitem.Vari_Ow_sokozano = fldvalue;
                                    break;
                                }
                            case "家主請求書発行予定日":
                                {
                                    model_cvitem.Vari_Ow_sqymd = fldvalue;
                                    break;
                                }
                            case "家主請求該当年月":
                                {
                                    model_cvitem.Vari_Ow_sqgtym = fldvalue;
                                    break;
                                }
                            case "家主請求締日":
                                {
                                    model_cvitem.Vari_Ow_sqsqsimeymd = fldvalue;
                                    break;
                                }
                            case "家主請求入金区分":
                                {
                                    model_cvitem.Vari_Ow_sqnkbn = fldvalue;
                                    break;
                                }
                            case "家主請求振込先口座No":
                                {
                                    model_cvitem.Vari_Ow_sqyatinkozano = fldvalue;
                                    break;
                                }
                            case "家主請求先No":
                                {
                                    model_cvitem.Vari_Ow_sqsakino = fldvalue;
                                    break;
                                }
                            case "家主控除請求備考":
                                {
                                    model_cvitem.Vari_Ow_biko = fldvalue;
                                    break;
                                }
                        }

                    }

                    // guid取得用に成形
                    model_cvitem.Vari_Hy_guid = tmp_bkno + "-" + tmp_hyno;
                    model_cvitem.Vari_Ky_guid = tmp_bkno + "-" + tmp_hyno + "-" + tmp_kyno;

                    // 固定値
                    model_cvitem.Vari_Ow_sosaki_soruleguid = "";
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 20160629 修繕検証後修正 -add sta
                    if (skipflg == false)
                    {
                        Chk_Futanritu(tblname, ref hash_cvitem, ref hash_log);
                    }
                    // 20160629 修繕検証後修正 -add end

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["bk_guid"] = hash_bkguid[Strings.StrConv(Conversions.ToString(hash_cvitem["bk_guid"]), VbStrConv.Narrow)];
                        if (!string.IsNullOrEmpty(tmp_hyno))
                        {
                            hash_cvitem["hy_guid"] = hash_bkhyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["hy_guid"]), VbStrConv.Narrow)];
                        }
                        else
                        {
                            hash_cvitem["hy_guid"] = "";
                        }
                        if (!string.IsNullOrEmpty(tmp_kyno))
                        {
                            hash_cvitem["ky_guid"] = hash_bkhykyguid[Strings.StrConv(Conversions.ToString(hash_cvitem["ky_guid"]), VbStrConv.Narrow)];
                        }
                        else
                        {
                            hash_cvitem["ky_guid"] = "";
                        }

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,szen_no) FROM szendata AS SZ ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            // 20160621 del sta
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

            // '2016.04.26 メインの方へも反映させる修正 -chg sta
            // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
            // Dim tmp_hash As New SafeDictionary<string, string>
            // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

            // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
            // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
            // '2016.04.26 メインの方へも反映させる修正 -chg end

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
            // 20160621 del end

            /// <summary>
            /// 物件guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_BkGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "SELECT bk_no,bk_guid FROM bkdata ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

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
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) ";
                tmp_sql = tmp_sql + " 	,hy_guid ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 契約guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_KyGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql = tmp_sql + " 	,ky_guid ";
                tmp_sql = tmp_sql + " FROM kydata AS KY ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 負担率の総和チェック '20160629 修繕検証後修正
            /// </summary>
            /// <param name="tblname"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_log"></param>
            /// <remarks></remarks>
            public void Chk_Futanritu(string tblname, ref SafeDictionary<string, string> hash_cvitem, ref SafeDictionary<string, string> hash_log)
            {

                double futan_kys = 0d;
                double futan_ow = 0d;
                double futan_jisya = 0d;
                double futan_total = 0d;

                string tmp_futan_kys = Conversions.ToString(Interaction.IIf(hash_cvitem["kys_futanrit"] is null, "0", hash_cvitem["kys_futanrit"]));
                string tmp_futan_ow = Conversions.ToString(Interaction.IIf(hash_cvitem["ow_futanrit"] is null, "0", hash_cvitem["ow_futanrit"]));
                string tmp_futan_jisya = Conversions.ToString(Interaction.IIf(hash_cvitem["jisya_futanrit"] is null, "0", hash_cvitem["jisya_futanrit"]));

                if (string.IsNullOrEmpty(tmp_futan_kys.Trim()) & string.IsNullOrEmpty(tmp_futan_ow.Trim()) & string.IsNullOrEmpty(tmp_futan_jisya.Trim()))
                {
                    return;
                }

                // ※データチェックで数値確認済み
                futan_kys = double.Parse(tmp_futan_kys);
                futan_ow = double.Parse(tmp_futan_ow);
                futan_jisya = double.Parse(tmp_futan_jisya);
                futan_total = futan_kys + futan_ow + futan_jisya;

                if (futan_total > 100d)
                {

                    // 負担率の総和が100を超えた場合は0に変換して移行する
                    hash_cvitem["kys_futanrit"] = "0";
                    hash_cvitem["ow_futanrit"] = "0";
                    hash_cvitem["jisya_futanrit"] = "0";

                    // ログ出力
                    string log_key = tblname + "-" + "kys_futanrit";
                    string errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                    if (hash_log.ContainsKey(log_key))
                    {
                        hash_log[log_key] = errstr;
                    }
                    else
                    {
                        hash_log.Add(log_key, errstr);
                    }

                }

            }

        }

    }

    #endregion

    #region 修繕見積情報

    public class Szendata_szen_Repository
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Szendata_szen_Model();           // 移行値格納用モデル初期化
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

                // guid取得
                Set_BkGuid(sqlcnnv10, ref hash_bkguid);

                // テーブル名/フィールド名セット
                string tblname = "szendata_szen";
                string fldnamegrp = "bk_guid,szen_no,szen_mituno,mitu_title,mitu_bango," + "mitu_ymd,seiyaku_flg,seiyaku_ymd,futan_kbn,kys_futanrit," + "ow_futanrit,jisya_futanrit,zei_kbn,gokei_zeikbn,gokei_zeirit," + "kys_gokeizeigak,ow_gokeizeigak,jisya_gokeizeigak,hasu_futankbn";



                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_bkguid, ref list_basekeydata);

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
                                    model_cvitem.Vari_Bk_guid = fldvalue;
                                    break;
                                }
                            case "修繕No":
                                {
                                    model_cvitem.Vari_Szen_no = fldvalue;
                                    break;
                                }
                            case "見積No":
                                {
                                    model_cvitem.Vari_Szen_mituno = fldvalue;
                                    break;
                                }
                            case "見積タイトル":
                                {
                                    model_cvitem.Vari_Mitu_title = fldvalue;
                                    break;
                                }
                            case "見積番号":
                                {
                                    model_cvitem.Vari_Mitu_bango = fldvalue;
                                    break;
                                }
                            case "見積日":
                                {
                                    model_cvitem.Vari_Mitu_ymd = fldvalue;
                                    break;
                                }
                            case "成約フラグ":
                                {
                                    model_cvitem.Vari_Seiyaku_flg = fldvalue;
                                    break;
                                }
                            case "成約日":
                                {
                                    model_cvitem.Vari_Seiyaku_ymd = fldvalue;
                                    break;
                                }
                            case "負担区分(全体・個別)":
                                {
                                    model_cvitem.Vari_Futan_kbn = fldvalue;
                                    break;
                                }
                            case "契約者負担率(全体用)":
                                {
                                    model_cvitem.Vari_Kys_futanrit = fldvalue;
                                    break;
                                }
                            case "家主負担率(全体用)":
                                {
                                    model_cvitem.Vari_Ow_futanrit = fldvalue;
                                    break;
                                }
                            case "自社負担率(全体用)":
                                {
                                    model_cvitem.Vari_Jisya_futanrit = fldvalue;
                                    break;
                                }
                            case "税適用区分(全体・個別)":
                                {
                                    model_cvitem.Vari_Zei_kbn = fldvalue;
                                    break;
                                }
                            case "税区分(全体用)":
                                {
                                    model_cvitem.Vari_Gokei_zeikbn = fldvalue;
                                    break;
                                }
                            case "適用税率":
                                {
                                    model_cvitem.Vari_Gokei_zeirit = fldvalue;
                                    break;
                                }
                            case "契約者税額(全体用)":
                                {
                                    model_cvitem.Vari_Kys_gokeizeigak = fldvalue;
                                    break;
                                }
                            case "家主税額(全体用)":
                                {
                                    model_cvitem.Vari_Ow_gokeizeigak = fldvalue;
                                    break;
                                }
                            case "自社税額(全体用)":
                                {
                                    model_cvitem.Vari_Jisya_gokeizeigak = fldvalue;
                                    break;
                                }
                            case "端数負担者区分":
                                {
                                    model_cvitem.Vari_Hasu_futankbn = fldvalue;
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

                    // 20160629 修繕検証後修正 -add sta
                    if (skipflg == false)
                    {
                        Chk_Futanritu(tblname, ref hash_cvitem, ref hash_log);
                    }
                    // 20160629 修繕検証後修正 -add end

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["bk_guid"] = hash_bkguid[Strings.StrConv(Conversions.ToString(hash_cvitem["bk_guid"]), VbStrConv.Narrow)];

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,szen_no) FROM szendata_szen AS SZ ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 物件guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_BkGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "SELECT bk_no,bk_guid FROM bkdata ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 負担率の総和チェック '20160629 修繕検証後修正
            /// </summary>
            /// <param name="tblname"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_log"></param>
            /// <remarks></remarks>
            public void Chk_Futanritu(string tblname, ref SafeDictionary<string, string> hash_cvitem, ref SafeDictionary<string, string> hash_log)
            {

                double futan_kys = 0d;
                double futan_ow = 0d;
                double futan_jisya = 0d;
                double futan_total = 0d;

                string tmp_futan_kys = Conversions.ToString(Interaction.IIf(hash_cvitem["kys_futanrit"] is null, "0", hash_cvitem["kys_futanrit"]));
                string tmp_futan_ow = Conversions.ToString(Interaction.IIf(hash_cvitem["ow_futanrit"] is null, "0", hash_cvitem["ow_futanrit"]));
                string tmp_futan_jisya = Conversions.ToString(Interaction.IIf(hash_cvitem["jisya_futanrit"] is null, "0", hash_cvitem["jisya_futanrit"]));

                if (string.IsNullOrEmpty(tmp_futan_kys.Trim()) & string.IsNullOrEmpty(tmp_futan_ow.Trim()) & string.IsNullOrEmpty(tmp_futan_jisya.Trim()))
                {
                    return;
                }

                // ※データチェックで数値確認済み
                futan_kys = double.Parse(tmp_futan_kys);
                futan_ow = double.Parse(tmp_futan_ow);
                futan_jisya = double.Parse(tmp_futan_jisya);
                futan_total = futan_kys + futan_ow + futan_jisya;

                if (futan_total > 100d)
                {

                    // 負担率の総和が100を超えた場合は0に変換して移行する
                    hash_cvitem["kys_futanrit"] = "0";
                    hash_cvitem["ow_futanrit"] = "0";
                    hash_cvitem["jisya_futanrit"] = "0";

                    // ログ出力
                    string log_key = tblname + "-" + "kys_futanrit";
                    string errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                    if (hash_log.ContainsKey(log_key))
                    {
                        hash_log[log_key] = errstr;
                    }
                    else
                    {
                        hash_log.Add(log_key, errstr);
                    }

                }

            }

        }

    }

    #endregion

    #region 修繕見積詳細情報

    public class Szendata_szenmeisai_Repository
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Szendata_szenmeisai_Model();     // 移行値格納用モデル初期化
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
                Set_BkGuid(sqlcnnv10, ref hash_bkguid);

                // テーブル名/フィールド名セット
                string tblname = "szendata_szenmeisai";
                string fldnamegrp = "bk_guid,szen_no,szen_mituno,szen_meisaino,szen_name," + "tekiyo,mitu_suryo,mitu_tani,mitu_tanka,mitu_zeikbn," + "mitu_zeigak,kys_futanrit,ow_futanrit,jisya_futanrit,szen_gyno," + "jikko_suryo,jikko_tani,jikko_tanka,jikko_zeikbn,jikko_zeigak";



                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_bkguid, ref list_basekeydata);

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
                                    model_cvitem.Vari_Bk_guid = fldvalue;
                                    break;
                                }
                            case "修繕No":
                                {
                                    model_cvitem.Vari_Szen_no = fldvalue;
                                    break;
                                }
                            case "見積No":
                                {
                                    model_cvitem.Vari_Szen_mituno = fldvalue;
                                    break;
                                }
                            case "見積明細No":
                                {
                                    model_cvitem.Vari_Szen_meisaino = fldvalue;
                                    break;
                                }
                            case "修繕項目名":
                                {
                                    model_cvitem.Vari_Szen_name = fldvalue;
                                    break;
                                }
                            case "摘要":
                                {
                                    model_cvitem.Vari_Tekiyo = fldvalue;
                                    break;
                                }
                            case "見積数量":
                                {
                                    model_cvitem.Vari_Mitu_suryo = fldvalue;
                                    break;
                                }
                            case "見積単位":
                                {
                                    model_cvitem.Vari_Mitu_tani = fldvalue;
                                    break;
                                }
                            case "見積単価":
                                {
                                    model_cvitem.Vari_Mitu_tanka = fldvalue;
                                    break;
                                }
                            case "見積税区分":
                                {
                                    model_cvitem.Vari_Mitu_zeikbn = fldvalue;
                                    break;
                                }
                            case "見積税額(税入力用)":
                                {
                                    model_cvitem.Vari_Mitu_zeigak = fldvalue;
                                    break;
                                }
                            case "契約者負担率(個別用)":
                                {
                                    model_cvitem.Vari_Kys_futanrit = fldvalue;
                                    break;
                                }
                            case "家主負担率(個別用)":
                                {
                                    model_cvitem.Vari_Ow_futanrit = fldvalue;
                                    break;
                                }
                            case "自社負担率(個別用)":
                                {
                                    model_cvitem.Vari_Jisya_futanrit = fldvalue;
                                    break;
                                }
                            case "発注業者No":
                                {
                                    model_cvitem.Vari_Szen_gyno = fldvalue;
                                    break;
                                }
                            case "実行数量":
                                {
                                    model_cvitem.Vari_Jikko_suryo = fldvalue;
                                    break;
                                }
                            case "実行単位":
                                {
                                    model_cvitem.Vari_Jikko_tani = fldvalue;
                                    break;
                                }
                            case "実行単価":
                                {
                                    model_cvitem.Vari_Jikko_tanka = fldvalue;
                                    break;
                                }
                            case "実行税区分":
                                {
                                    model_cvitem.Vari_Jikko_zeikbn = fldvalue;
                                    break;
                                }
                            case "実行税額(税入力用)":
                                {
                                    model_cvitem.Vari_Jikko_zeigak = fldvalue;
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

                    // 20160629 修繕検証後修正 -add sta
                    if (skipflg == false)
                    {
                        skipflg = !Chk_Futanritu(tblname, ref hash_cvitem, ref hash_log);
                    }
                    // 20160629 修繕検証後修正 -add end

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        hash_cvitem["bk_guid"] = hash_bkguid[Strings.StrConv(Conversions.ToString(hash_cvitem["bk_guid"]), VbStrConv.Narrow)];

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,bk_no) + '-' + ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,szen_no) + '-' + ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,szen_mituno) + '-' + ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,szen_meisaino) ";
                tmp_sql = tmp_sql + " FROM szendata_szenmeisai AS SZ ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid ";

                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 物件guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_BkGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "SELECT bk_no,bk_guid FROM bkdata ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

            /// <summary>
            /// 負担率の総和チェック '20160629 修繕検証後修正
            /// </summary>
            /// <param name="tblname"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_log"></param>
            /// <remarks></remarks>
            public bool Chk_Futanritu(string tblname, ref SafeDictionary<string, string> hash_cvitem, ref SafeDictionary<string, string> hash_log)
            {

                bool rtn = true;

                double futan_kys = 0d;
                double futan_ow = 0d;
                double futan_jisya = 0d;
                double futan_total = 0d;

                string tmp_futan_kys = Conversions.ToString(Interaction.IIf(hash_cvitem["kys_futanrit"] is null, "0", hash_cvitem["kys_futanrit"]));
                string tmp_futan_ow = Conversions.ToString(Interaction.IIf(hash_cvitem["ow_futanrit"] is null, "0", hash_cvitem["ow_futanrit"]));
                string tmp_futan_jisya = Conversions.ToString(Interaction.IIf(hash_cvitem["jisya_futanrit"] is null, "0", hash_cvitem["jisya_futanrit"]));

                if (string.IsNullOrEmpty(tmp_futan_kys.Trim()) & string.IsNullOrEmpty(tmp_futan_ow.Trim()) & string.IsNullOrEmpty(tmp_futan_jisya.Trim()))
                {
                    return rtn;
                }

                // ※データチェックで数値確認済み
                futan_kys = double.Parse(tmp_futan_kys);
                futan_ow = double.Parse(tmp_futan_ow);
                futan_jisya = double.Parse(tmp_futan_jisya);
                futan_total = futan_kys + futan_ow + futan_jisya;

                if (futan_total > 100d)
                {

                    // 負担率の総和が100を超えた場合は0に変換して移行する
                    hash_cvitem["kys_futanrit"] = "0";
                    hash_cvitem["ow_futanrit"] = "0";
                    hash_cvitem["jisya_futanrit"] = "0";

                    // ログ出力
                    string log_key = tblname + "-" + "kys_futanrit";
                    string errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                    if (hash_log.ContainsKey(log_key))
                    {
                        hash_log[log_key] = errstr;
                    }
                    else
                    {
                        hash_log.Add(log_key, errstr);
                    }

                    rtn = false;

                }

                return rtn;

            }

        }

    }

    #endregion

    #region 修繕クレーム関連付け情報

    public class Szendata_claim_Repository
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Szendata_claim_Model();          // 移行値格納用モデル初期化
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

                // guid取得
                Set_BkGuid(sqlcnnv10, ref hash_bkguid);

                // テーブル名/フィールド名セット
                string tblname = "szendata_claim";
                string fldnamegrp = "bk_guid,szen_no,claim_no,claimdata_sortorder,szendata_sortorder," + "history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_bkguid, ref list_basekeydata);

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
                                    model_cvitem.Vari_Bk_guid = fldvalue;
                                    break;
                                }
                            case "修繕No":
                                {
                                    model_cvitem.Vari_Szen_no = fldvalue;
                                    break;
                                }
                            case "クレームNo":
                                {
                                    model_cvitem.Vari_Claim_no = fldvalue;
                                    break;
                                }
                            case "クレーム画面での並び順":
                                {
                                    model_cvitem.Vari_Claimdata_sortorder = fldvalue;
                                    break;
                                }
                            case "修繕情報登録画面での並び順":
                                {
                                    model_cvitem.Vari_Szendata_sortorder = fldvalue;
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

                        // キーをguidへ変換
                        hash_cvitem["bk_guid"] = hash_bkguid[Strings.StrConv(Conversions.ToString(hash_cvitem["bk_guid"]), VbStrConv.Narrow)];

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,bk_no) + '-' + ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,szen_no) + '-' + ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,claim_no) ";
                tmp_sql = tmp_sql + " FROM szendata_claim AS SZ ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid ";

                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 物件guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_BkGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "SELECT bk_no,bk_guid FROM bkdata ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

        }

    }

    #endregion

    #region 修繕関連ファイル情報

    public class Szendata_relfile_Repository
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Szendata_relfile_Model();        // 移行値格納用モデル初期化
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

                // guid取得
                Set_BkGuid(sqlcnnv10, ref hash_bkguid);

                // テーブル名/フィールド名セット
                string tblname = "szendata_relfile";
                string fldnamegrp = "bk_guid,szen_no,file_no,fullpath,addtime," + "biko,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_bkguid, ref list_basekeydata);

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
                                    model_cvitem.Vari_Bk_guid = fldvalue;
                                    break;
                                }
                            case "修繕No":
                                {
                                    model_cvitem.Vari_Szen_no = fldvalue;
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

                        // キーをguidへ変換
                        hash_cvitem["bk_guid"] = hash_bkguid[Strings.StrConv(Conversions.ToString(hash_cvitem["bk_guid"]), VbStrConv.Narrow)];

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,szen_no) FROM szendata_szen AS SZ ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 物件guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_BkGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "SELECT bk_no,bk_guid FROM bkdata ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

        }

    }

    #endregion

    #region 修繕メモ情報

    public class Szendata_memo_Repository
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
                var hash_bkguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Szendata_memo_Model();           // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列
                int tmp_cvrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

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
                Set_BkGuid(sqlcnnv10, ref hash_bkguid);

                // テーブル名/フィールド名セット
                string tblname = "szendata_memo";
                string fldnamegrp = "bk_guid,szen_no,memo_no,memo,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_bkguid, ref list_basekeydata);

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
                string fldname_keysub = Conversions.ToString(headervalue(startrow - 1, keycol_sub));

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
                    string tmp_keysub = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub]);
                    model_cvitem.Vari_Bk_guid = tmp_keymain;
                    model_cvitem.Vari_Szen_no = tmp_keysub;

                    // 備考カウント初期化
                    int cvitemcnt = 0;

                    // 移行値取得
                    for (int cntjj = 3, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        // サブ2キー取得
                        string tmp_keysub2 = (cntjj - 2).ToString();

                        // 全キー取得
                        string fldvalue_key = tmp_keymain + "-" + tmp_keysub + "-" + tmp_keysub2;

                        // ログ出力用データ格納(サブキーフィールド)
                        string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, cntjj));
                        string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub + " = " + tmp_keysub + "、" + fldname_keysub2;

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }
                        model_cvitem.Vari_Memo = fldvalue;

                        // 固定値
                        model_cvitem.Vari_Memo_no = (cntjj - 2).ToString();
                        model_cvitem.Vari_History = CommonModule.DefHistory;

                        // メモにデータが存在する場合に書込処理を行う
                        if (!string.IsNullOrEmpty(model_cvitem.Vari_Memo))
                        {

                            // 20160928 メモ関連の移行件数表示修正 -add
                            tmp_cvrowcnt = tmp_cvrowcnt + 1;

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

                                // 挿入処理
                                CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                                // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                if (normalflg)
                                {
                                    list_chkduplicate.Add(fldvalue_key);
                                    cvitemcnt = cvitemcnt + 1;
                                }

                            }

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

                    // --------------------------------------------------------------------
                    // 移行した備考が1データ以上ある場合移行したレコードの数を更新する
                    // --------------------------------------------------------------------
                    // 20160928 メモ関連の移行件数表示修正 -chg sta
                    // tmp_cvcnt = tmp_cvcnt + 1
                    if (cvitemcnt > 0)
                    {
                        tmp_cvcnt = tmp_cvcnt + 1;
                    }
                    // 20160928 メモ関連の移行件数表示修正 -chg end
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
                // 20160928 メモ関連の移行件数表示修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = tmp_cvrowcnt;
                // 20160928 メモ関連の移行件数表示修正 -chg end

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

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,bk_no) + '-' + ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,szen_no) + '-' + ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,memo_no) ";
                tmp_sql = tmp_sql + " FROM szendata_memo AS SZ ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON SZ.bk_guid = BK.bk_guid ";

                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 物件guidの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Set_BkGuid(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
            {

                string tmp_sql = "SELECT bk_no,bk_guid FROM bkdata ";
                DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

            }

        }

    }

    #endregion

}