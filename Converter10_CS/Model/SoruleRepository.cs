using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Converter10.Njc.N3Lib.Utys;

namespace Converter10.Njc.Repository
{

    #region 送金ルール基本情報

    public class Sorule_Repository
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
                var hash_bkhyguid = new SafeDictionary<string, string>();                              // bkguid格納用ハッシュテーブル
                var hash_bkhysoguid = new SafeDictionary<string, string>();                            // bkguid/soruleguid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Sorule_Model();                  // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewnamebkhytoguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhytosoruleguid = CommonModule.PRE_VIEW_NAME + "物件部屋送金キー情報";

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

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytoguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytosoruleguid);

                // テーブル名/フィールド名セット
                string tblname = "sorule";
                string fldnamegrp = "sorule_guid,sorule_no,relation_guid,kubunsyo,sorule_startymd," + "sorule_endymd,kanri_keitaikbn,ikkatu_kanriflg,cyukai_kyflg,cyukai_koflg," + "cyukai_kaiflg,soymd1_gaitoukbn,soymd1_simekbn,soymd1_sokintukikbn,soymd1_sokinsimekbn," + "soymd2_gaitoukbn,soymd2_simekbn,soymd2_sokintukikbn,soymd2_sokinsimekbn,soymd3_gaitoukbn," + "soymd3_simekbn,soymd3_sokintukikbn,soymd3_sokinsimekbn,soymd4_gaitoukbn,soymd4_simekbn," + "soymd4_sokintukikbn,soymd4_sokinsimekbn,soymd5_gaitoukbn,soymd5_simekbn,soymd5_sokintukikbn," + "soymd5_sokinsimekbn,sosaki_multikbn,sosaki_koteiflg,sosaki_koteisu,sosaki_anbunflg," + "sosaki_hasuuketorisaki,sosaki_hasuadjustmentflg,ikkatu_kbn,ikkatu_bkgak,ikkatu_menseki," + "ikkatu_mensekimonth,kanritesu_kbn,kanritesu_cyosyukbn,kanri_reigaiky,kanri_reigaikai," + "hyteigaku_kbn,hyteigaku_kanrigak,hyteigaku_hiwariflg,bkteigaku_kanrigak,kanri_taxflg," + "biko_basic,sh_daihyohyflg,sh_kozabetuflg,sh_kozabetuListNokbn,hyikkatu_nkin_no1kbn," + "hyikkatu_nkin_no2kbn,hyikkatu_nkin_no3kbn,hyikkatu_nkin_no4kbn,hyikkatu_nkin_no5kbn,kanritesu_zeiumukbn," + "history,rowid,no_soruleflg,cyukai_zuijisokin,ikkatu_calckbn," + "ikkatu_taxflg,kanri_utizeiflg,soymd_basis";













                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhytoguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhytosoruleguid, ref hash_bkhysoguid);

                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, viewnamebkhytoguid, ref list_basekeydata);

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

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol_main - 1].ColumnName.Trim();
                string fldname_keysub1 = readtbl.Columns[keycol_sub1 - 1].ColumnName.Trim();
                string fldname_keysub2 = readtbl.Columns[keycol_sub2 - 1].ColumnName.Trim();

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    model_cvitem.Vari_Sorule_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    model_cvitem.Vari_Relation_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

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
                            case "送金ルール管理No":
                                {
                                    model_cvitem.Vari_Sorule_no = fldvalue.Trim();
                                    break;
                                }
                            case "一所有形態区分-棟/区分":
                                {
                                    model_cvitem.Vari_Kubunsyo = fldvalue.Trim();
                                    break;
                                }
                            case "送金ルール適用開始日":
                                {
                                    model_cvitem.Vari_Sorule_startymd = fldvalue.Trim();
                                    break;
                                }
                            case "送金ルール適用終了日":
                                {
                                    model_cvitem.Vari_Sorule_endymd = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態":
                                {
                                    model_cvitem.Vari_Kanri_keitaikbn = fldvalue.Trim();
                                    break;
                                }
                            case "一括借上 一部管理":
                                {
                                    model_cvitem.Vari_Ikkatu_kanriflg = fldvalue.Trim();
                                    break;
                                }
                            case "仲介物件 - 新規契約業務":
                                {
                                    model_cvitem.Vari_Cyukai_kyflg = fldvalue.Trim();
                                    break;
                                }
                            case "仲介物件 - 契約更新業務":
                                {
                                    model_cvitem.Vari_Cyukai_koflg = fldvalue.Trim();
                                    break;
                                }
                            case "仲介物件 - 解約業務":
                                {
                                    model_cvitem.Vari_Cyukai_kaiflg = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法1 - 該当年月":
                                {
                                    model_cvitem.Vari_Soymd1_gaitoukbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法1 - 締日":
                                {
                                    model_cvitem.Vari_Soymd1_simekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法1 - 送金月":
                                {
                                    model_cvitem.Vari_Soymd1_sokintukikbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法1 - 送金締日":
                                {
                                    model_cvitem.Vari_Soymd1_sokinsimekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法2 - 該当年月":
                                {
                                    model_cvitem.Vari_Soymd2_gaitoukbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法2 - 締日":
                                {
                                    model_cvitem.Vari_Soymd2_simekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法2 - 送金月":
                                {
                                    model_cvitem.Vari_Soymd2_sokintukikbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法2 - 送金締日":
                                {
                                    model_cvitem.Vari_Soymd2_sokinsimekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法3 - 該当年月":
                                {
                                    model_cvitem.Vari_Soymd3_gaitoukbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法3 - 締日":
                                {
                                    model_cvitem.Vari_Soymd3_simekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法3 - 送金月":
                                {
                                    model_cvitem.Vari_Soymd3_sokintukikbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法3 - 送金締日":
                                {
                                    model_cvitem.Vari_Soymd3_sokinsimekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法4 - 該当年月":
                                {
                                    model_cvitem.Vari_Soymd4_gaitoukbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法4 - 締日":
                                {
                                    model_cvitem.Vari_Soymd4_simekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法4 - 送金月":
                                {
                                    model_cvitem.Vari_Soymd4_sokintukikbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法4 - 送金締日":
                                {
                                    model_cvitem.Vari_Soymd4_sokinsimekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法5 - 該当年月":
                                {
                                    model_cvitem.Vari_Soymd5_gaitoukbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法5 - 締日":
                                {
                                    model_cvitem.Vari_Soymd5_simekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法5 - 送金月":
                                {
                                    model_cvitem.Vari_Soymd5_sokintukikbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金日決定方法5 - 送金締日":
                                {
                                    model_cvitem.Vari_Soymd5_sokinsimekbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金先単独/複数指定":
                                {
                                    model_cvitem.Vari_Sosaki_multikbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金固定額使用フラグ":
                                {
                                    model_cvitem.Vari_Sosaki_koteiflg = fldvalue.Trim();
                                    break;
                                }
                            case "送金固定数":
                                {
                                    model_cvitem.Vari_Sosaki_koteisu = fldvalue.Trim();
                                    break;
                                }
                            case "均等案分フラグ":
                                {
                                    model_cvitem.Vari_Sosaki_anbunflg = fldvalue.Trim();
                                    break;
                                }
                            case "端数受取先":
                                {
                                    model_cvitem.Vari_Sosaki_hasuuketorisaki = fldvalue.Trim();
                                    break;
                                }
                            case "送金額案分端数調整フラグ":
                                {
                                    model_cvitem.Vari_Sosaki_hasuadjustmentflg = fldvalue.Trim();
                                    break;
                                }
                            case "一括借上 物件毎/部屋毎":
                                {
                                    model_cvitem.Vari_Ikkatu_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "一括借上 物件毎設定額":
                                {
                                    model_cvitem.Vari_Ikkatu_bkgak = fldvalue.Trim();
                                    break;
                                }
                            case "一括借上 免責期間の設定フラグ":
                                {
                                    model_cvitem.Vari_Ikkatu_menseki = fldvalue.Trim();
                                    break;
                                }
                            case "一括借上 免責期間(解約翌月から何カ月)":
                                {
                                    model_cvitem.Vari_Ikkatu_mensekimonth = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料区分":
                                {
                                    model_cvitem.Vari_Kanritesu_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料徴収区分":
                                {
                                    model_cvitem.Vari_Kanritesu_cyosyukbn = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料 例外：契約金は対象外とする":
                                {
                                    model_cvitem.Vari_Kanri_reigaiky = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料 例外：解約金は対象外とする":
                                {
                                    model_cvitem.Vari_Kanri_reigaikai = fldvalue.Trim();
                                    break;
                                }
                            case "部屋毎定額 - 全部屋一律管理手数料":
                                {
                                    model_cvitem.Vari_Hyteigaku_kanrigak = fldvalue.Trim();
                                    break;
                                }
                            case "部屋毎定額 - 部屋毎日割りフラグ":
                                {
                                    model_cvitem.Vari_Hyteigaku_hiwariflg = fldvalue.Trim();
                                    break;
                                }
                            case "物件毎定額 - 物件管理手数料":
                                {
                                    model_cvitem.Vari_Bkteigaku_kanrigak = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料 消費税適用フラグ":
                                {
                                    model_cvitem.Vari_Kanri_taxflg = fldvalue.Trim();
                                    break;
                                }
                            case "備考 - 基本情報":
                                {
                                    model_cvitem.Vari_Biko_basic = fldvalue.Trim();
                                    break;
                                }
                            case "支払明細書関連 - 同時契約の場合は代表する部屋を表示":
                                {
                                    model_cvitem.Vari_Sh_daihyohyflg = fldvalue.Trim();
                                    break;
                                }
                            case "支払明細書関連 - 口座毎に支払明細書を作成1":
                                {
                                    model_cvitem.Vari_Sh_kozabetuflg = fldvalue.Trim();
                                    break;
                                }
                            case "支払明細書関連 - 口座毎に支払明細書を作成2":
                                {
                                    model_cvitem.Vari_Sh_kozabetuListNokbn = fldvalue.Trim();
                                    break;
                                }
                            case "部屋毎一括借上額 - 入金項目No1":
                                {
                                    model_cvitem.Vari_Hyikkatu_nkin_no1kbn = fldvalue.Trim();
                                    break;
                                }
                            case "部屋毎一括借上額 - 入金項目No2":
                                {
                                    model_cvitem.Vari_Hyikkatu_nkin_no2kbn = fldvalue.Trim();
                                    break;
                                }
                            case "部屋毎一括借上額 - 入金項目No3":
                                {
                                    model_cvitem.Vari_Hyikkatu_nkin_no3kbn = fldvalue.Trim();
                                    break;
                                }
                            case "部屋毎一括借上額 - 入金項目No4":
                                {
                                    model_cvitem.Vari_Hyikkatu_nkin_no4kbn = fldvalue.Trim();
                                    break;
                                }
                            case "部屋毎一括借上額 - 入金項目No5":
                                {
                                    model_cvitem.Vari_Hyikkatu_nkin_no5kbn = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料計算基準区分":
                                {
                                    model_cvitem.Vari_Kanritesu_zeiumukbn = fldvalue.Trim();
                                    break;
                                }
                            case "満額入金-送金フラグ":
                                {
                                    model_cvitem.Vari_No_soruleflg = fldvalue.Trim();
                                    break;
                                }
                            case "仲介物件 - 随時送金":
                                {
                                    model_cvitem.Vari_Cyukai_zuijisokin = fldvalue.Trim();
                                    break;
                                }
                            case "一括借上 部屋毎一括借上額の計算方法":
                                {
                                    model_cvitem.Vari_Ikkatu_calckbn = fldvalue.Trim();
                                    break;
                                }
                            case "一括借上 消費税適用フラグ":
                                {
                                    model_cvitem.Vari_Ikkatu_taxflg = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Kanri_utizeiflg = 0.ToString();
                    model_cvitem.Vari_Soymd_basis = 0.ToString();

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
                        hash_cvitem["sorule_guid"] = hash_bkhysoguid[hash_cvitem["sorule_guid"]];
                        hash_cvitem["relation_guid"] = hash_bkhyguid[hash_cvitem["relation_guid"]];

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

                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		CONVERT(varchar,bk_no) + '-' + '' + '-' + CONVERT(varchar,sorule_no) AS キー ";
                tmp_sql = tmp_sql + " 	FROM " + tblname;
                tmp_sql = tmp_sql + " 	LEFT JOIN bkdata ON " + tblname + ".relation_guid = bkdata.bk_guid ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,sorule_no) ";
                tmp_sql = tmp_sql + " 	FROM " + tblname;
                tmp_sql = tmp_sql + " 	LEFT JOIN hydata ON " + tblname + ".relation_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
                tmp_sql = tmp_sql + " ) AS VW ";
                tmp_sql = tmp_sql + " WHERE キー IS NOT NULL";


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
                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件部屋キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,bk_no)  + '-' AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,bk_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM bkdata ";
                            tmp_sql_create = tmp_sql_create + " UNION ";
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋送金キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + '-' + CONVERT(varchar,BKS.kn_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,sorule_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM bkdata_syo AS BKS ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON BKS.bk_guid = BK.bk_guid ";
                            tmp_sql_create = tmp_sql_create + " UNION ";
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no)  + '-' + CONVERT(varchar,HYS.kn_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,sorule_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata_syo AS HYS ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON HYS.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 送金ルール送金先情報

    public class Sorule_sosaki_Repository
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
                var hash_bkhyguid = new SafeDictionary<string, string>();                                // bkguid格納用ハッシュテーブル
                var hash_bkhysoguid = new SafeDictionary<string, string>();                              // bkguid/soruleguid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Sorule_sosaki_Model();           // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                string viewnamebkhytoguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhytosoruleguid = CommonModule.PRE_VIEW_NAME + "物件部屋送金キー情報";

                // ************************
                // 作業準備
                // ************************
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
                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytoguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytosoruleguid);

                // テーブル名/フィールド名セット
                string tblname = "sorule_sosaki";
                string fldnamegrp = "sorule_guid,sorule_no,sosaki_recno,so_ow_no,so_ow_kozano," + "so_rit,so_fixsogak";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhytoguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhytosoruleguid, ref hash_bkhysoguid);

                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, viewnamebkhytosoruleguid, ref list_basekeydata);

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

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol_main - 1].ColumnName.Trim();
                string fldname_keysub1 = readtbl.Columns[keycol_sub1 - 1].ColumnName.Trim();
                string fldname_keysub2 = readtbl.Columns[keycol_sub2 - 1].ColumnName.Trim();
                string fldname_keysub3 = readtbl.Columns[keycol_sub3 - 1].ColumnName.Trim();

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
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string tmp_keysub3 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub3 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;
                    model_cvitem.Vari_Sorule_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

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
                            case "送金ルール管理No":
                                {
                                    model_cvitem.Vari_Sorule_no = fldvalue.Trim();
                                    break;
                                }
                            case "送金先行No":
                                {
                                    model_cvitem.Vari_Sosaki_recno = fldvalue.Trim();
                                    break;
                                }
                            case "送金先家主NO":
                                {
                                    model_cvitem.Vari_So_ow_no = fldvalue.Trim();
                                    break;
                                }
                            case "送金先口座NO":
                                {
                                    model_cvitem.Vari_So_ow_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "送金率":
                                {
                                    model_cvitem.Vari_So_rit = fldvalue.Trim();
                                    break;
                                }
                            case "固定送金額":
                                {
                                    model_cvitem.Vari_So_fixsogak = fldvalue.Trim();
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

                        // キーをguidへ変換
                        hash_cvitem["sorule_guid"] = hash_bkhysoguid[hash_cvitem["sorule_guid"]];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                        }
                        else    // 20161014 重複エラーのログ出力処理を追加 -add
                        {
                            string log_key = "sorule_sosaki-sorule_guid" + "/" + "sorule_sosaki-sorule_no" + "/" + "sorule_sosaki-sosaki_recno";
                            string errstr = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP;
                            hash_log.Clear();
                            hash_log.Add(log_key, errstr);
                            tmp_hash["sorule_guid"] = "";
                            hash_cvitem["sorule_guid"] = "";
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

                // 20161108_2 送金ルール送金先情報の移行制御処理修正 -add sta
                // 移行したデータのうち、管理形態が自社物件、送金保留以外の場合は送金先と送金先Noが必須なのでデータ調整を行う
                // ※ここで簡易的な例外処理を入れておく(20161108 例外処理追加)
                if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                {
                    try
                    {
                        Chk_SoruleSosaki_Sono(sqlcnnv10, tblname);
                        Chk_SoruleSosaki_SoKozano(sqlcnnv10, tblname);
                        Chk_SoruleBase(sqlcnnv10, tblname);
                        Set_SoruleSosakiDummyRecord(sqlcnnv10);
                    }
                    catch (Exception ex)
                    {
                        // ----- ログ出力 -----
                        int tmptmpcnt = 0;
                        string tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(41, sheetname, "データ調整", "データ調整中にエラーが発生しました。"), false);
                        DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref tmptmpcnt);
                    }
                }
                // 20161108_2 送金ルール送金先情報の移行制御処理修正 -add end

                // 20161108_2 送金ルール送金先情報の移行制御処理修正 -del sta
                // ↑に統合
                // '20161011 送金ルール送金先有無による送金ルール基本情報の移行制御 -add sta
                // '送金ルール基本情報と送金ルール送金先情報を照合し移行制御を行う
                // If CNVNO = ConvertTypes._汎用 Then
                // Call Me.Chk_SoruleBase(sqlcnnv10, tblname)
                // End If
                // '20161011 送金ルール送金先有無による送金ルール基本情報の移行制御 -add end
                // 20161108_2 送金ルール送金先情報の移行制御処理修正 -del end
                // ************************
                // 終了処理
                // ************************

                // 中間ファイル件数を取得
                midrowcnt = rowcnt;

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

                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT CONVERT(varchar,bk_no) + '-' + '' + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " + tblname;
                tmp_sql = tmp_sql + " 	LEFT JOIN sorule ON " + tblname + ".sorule_guid = sorule.sorule_guid ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bkdata ON sorule.relation_guid = bkdata.bk_guid ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " + tblname;
                tmp_sql = tmp_sql + " 	LEFT JOIN sorule ON " + tblname + ".sorule_guid = sorule.sorule_guid ";
                tmp_sql = tmp_sql + " 	LEFT JOIN hydata ON sorule.relation_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
                tmp_sql = tmp_sql + " ) AS VW ";
                tmp_sql = tmp_sql + " WHERE キー IS NOT NULL ";


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
                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件部屋キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,bk_no)  + '-' AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,bk_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM bkdata ";
                            tmp_sql_create = tmp_sql_create + " UNION ";
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋送金キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + '-' + CONVERT(varchar,BKS.kn_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,sorule_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM bkdata_syo AS BKS ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON BKS.bk_guid = BK.bk_guid ";
                            tmp_sql_create = tmp_sql_create + " UNION ";
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no)  + '-' + CONVERT(varchar,HYS.kn_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,sorule_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata_syo AS HYS ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON HYS.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 送金先情報で送金先が存在しないデータを抽出、削除
            /// </summary>
            /// <param name="sqlcnnv10">DB接続用オブジェクト</param>
            /// <param name="tblname">送金先情報のテーブル名</param>
            /// <remarks>20161108_2 送金ルール送金先情報の移行制御処理修正 新規追加</remarks>
            public void Chk_SoruleSosaki_Sono(SqlConnection sqlcnnv10, string tblname)
            {

                var list_err = new List<string>();

                // ----------------
                // 不正データ抽出
                // ----------------
                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) AS [ログ対象データ] ";
                tmp_sql = tmp_sql + " 		FROM sorule_sosaki AS SOS ";
                tmp_sql = tmp_sql + " 		LEFT JOIN sorule AS SO ON SOS.sorule_guid = SO.sorule_guid ";
                tmp_sql = tmp_sql + " 		LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " 		WHERE SOS.so_ow_no IS NULL AND (SO.kanri_keitaikbn NOT IN (3,4,5) OR (SO.kanri_keitaikbn = 3 AND (SO.cyukai_kyflg = 1 OR SO.cyukai_koflg = 1 OR SO.cyukai_kaiflg = 1))) ";
                tmp_sql = tmp_sql + " 	) AS ITTO ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) + '、' + '部屋No = ' + CONVERT(VARCHAR(MAX),hy_no) AS [ログ対象データ] ";
                tmp_sql = tmp_sql + " 		FROM sorule_sosaki AS SOS ";
                tmp_sql = tmp_sql + " 		LEFT JOIN sorule AS SO ON SOS.sorule_guid = SO.sorule_guid ";
                tmp_sql = tmp_sql + " 		LEFT JOIN hydata AS HY ON SO.relation_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " 		LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " 		WHERE SOS.so_ow_no IS NULL AND (SO.kanri_keitaikbn NOT IN (3,4,5) OR (SO.kanri_keitaikbn = 3 AND (SO.cyukai_kyflg = 1 OR SO.cyukai_koflg = 1 OR SO.cyukai_kaiflg = 1))) ";
                tmp_sql = tmp_sql + " 	) AS KBN ";
                tmp_sql = tmp_sql + " ) AS VW ";
                tmp_sql = tmp_sql + " WHERE [ログ対象データ] IS NOT NULL ";
                DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_err);

                // ----------------------------------------------
                // 不正データが存在する場合、ログ出力→削除処理
                // ----------------------------------------------
                if (list_err.Count != 0)
                {

                    // 作業用変数
                    var tmp_hash = new SafeDictionary<string, string>();       // ログ作成用の仮オブジェクト(引数として使用するだけ)
                    int tmp_logcnt = 0;       // ログ作成用の仮変数(引数として使用するだけ)
                    int tmp_cnt = 0;          // 仮変数(引数として使用するだけ)
                    var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                    var tmp_loghash = new SafeDictionary<string, string>();
                    string log_key = tblname + "-" + "so_ow_no";
                    string tmp_hubi = "物件管理情報の送金先が設定されていないため移行できません。";
                    string errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + tmp_hubi + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                    tmp_loghash.Add(log_key, errstr);

                    // ログ出力
                    foreach (var str_logkey in list_err)
                    {

                        // ログ出力メッセージ整形
                        LogSetting.Set_Log_Value_KomkErr(tmp_loghash, tmp_hash, tmp_hash, str_logkey, tblname, ref tmp_logcnt, ref sortlist_log);

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

                    // 不正データ削除
                    string tmp_sql_delete = "";
                    tmp_sql_delete = tmp_sql_delete + " DELETE FROM sorule_sosaki ";
                    tmp_sql_delete = tmp_sql_delete + " WHERE so_ow_no IS NULL ";
                    tmp_sql_delete = tmp_sql_delete + " AND   sorule_guid IN ";
                    tmp_sql_delete = tmp_sql_delete + " (SELECT sorule_guid FROM sorule WHERE kanri_keitaikbn NOT IN (3,4,5) OR (kanri_keitaikbn = 3 AND (cyukai_kyflg = 1 OR cyukai_koflg = 1 OR cyukai_kaiflg = 1))) ";
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_delete, ref tmp_cnt);

                }

            }

            /// <summary>
            /// 送金先情報で送金先口座が存在しないデータを抽出、削除
            /// 送金先が存在しないデータは調整済み
            /// </summary>
            /// <param name="sqlcnnv10">DB接続用オブジェクト</param>
            /// <param name="tblname">送金先情報のテーブル名</param>
            /// <remarks>20161108_2 送金ルール送金先情報の移行制御処理修正 新規追加</remarks>
            public void Chk_SoruleSosaki_SoKozano(SqlConnection sqlcnnv10, string tblname)
            {

                var list_err = new List<string>();

                // ----------------
                // 不正データ抽出
                // ----------------
                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) AS [ログ対象データ] ";
                tmp_sql = tmp_sql + " 		FROM sorule_sosaki AS SOS ";
                tmp_sql = tmp_sql + " 		LEFT JOIN sorule AS SO ON SOS.sorule_guid = SO.sorule_guid ";
                tmp_sql = tmp_sql + " 		LEFT JOIN bkdata AS BK ON SO.relation_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " 		WHERE SOS.so_ow_no IS NOT NULL AND SOS.so_ow_kozano IS NULL AND (SO.kanri_keitaikbn NOT IN (3,4,5) OR (SO.kanri_keitaikbn = 3 AND (SO.cyukai_kyflg = 1 OR SO.cyukai_koflg = 1 OR SO.cyukai_kaiflg = 1))) ";
                tmp_sql = tmp_sql + " 	) AS ITTO ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) + '、' + '部屋No = ' + CONVERT(VARCHAR(MAX),hy_no) AS [ログ対象データ] ";
                tmp_sql = tmp_sql + " 		FROM sorule_sosaki AS SOS ";
                tmp_sql = tmp_sql + " 		LEFT JOIN sorule AS SO ON SOS.sorule_guid = SO.sorule_guid ";
                tmp_sql = tmp_sql + " 		LEFT JOIN hydata AS HY ON SO.relation_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " 		LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " 		WHERE SOS.so_ow_no IS NOT NULL AND SOS.so_ow_kozano IS NULL AND (SO.kanri_keitaikbn NOT IN (3,4,5) OR (SO.kanri_keitaikbn = 3 AND (SO.cyukai_kyflg = 1 OR SO.cyukai_koflg = 1 OR SO.cyukai_kaiflg = 1))) ";
                tmp_sql = tmp_sql + " 	) AS KBN ";
                tmp_sql = tmp_sql + " ) AS VW ";
                tmp_sql = tmp_sql + " WHERE [ログ対象データ] IS NOT NULL ";
                DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_err);

                // ----------------------------------------------
                // 不正データが存在する場合、ログ出力→削除処理
                // ----------------------------------------------
                if (list_err.Count != 0)
                {

                    // 作業用変数
                    var tmp_hash = new SafeDictionary<string, string>();       // ログ作成用の仮オブジェクト(引数として使用するだけ)
                    int tmp_logcnt = 0;       // ログ作成用の仮変数(引数として使用するだけ)
                    int tmp_cnt = 0;          // 仮変数(引数として使用するだけ)
                    var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                    var tmp_loghash = new SafeDictionary<string, string>();
                    string log_key = tblname + "-" + "so_ow_kozano";
                    string tmp_hubi = "物件管理情報の送金先口座が設定されていないため移行できません。";
                    string errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + tmp_hubi + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                    tmp_loghash.Add(log_key, errstr);

                    // ログ出力
                    foreach (var str_logkey in list_err)
                    {

                        // ログ出力メッセージ整形
                        LogSetting.Set_Log_Value_KomkErr(tmp_loghash, tmp_hash, tmp_hash, str_logkey, tblname, ref tmp_logcnt, ref sortlist_log);

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

                    // 不正データ削除
                    string tmp_sql_delete = "";
                    tmp_sql_delete = tmp_sql_delete + " DELETE FROM sorule_sosaki ";
                    tmp_sql_delete = tmp_sql_delete + " WHERE so_ow_no IS NOT NULL AND so_ow_kozano IS NULL ";
                    tmp_sql_delete = tmp_sql_delete + " AND   sorule_guid IN ";
                    tmp_sql_delete = tmp_sql_delete + " (SELECT sorule_guid FROM sorule WHERE kanri_keitaikbn NOT IN (3,4,5) OR (kanri_keitaikbn = 3 AND (cyukai_kyflg = 1 OR cyukai_koflg = 1 OR cyukai_kaiflg = 1))) ";
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_delete, ref tmp_cnt);

                }

            }

            /// <summary>
            /// 送金ルール基本情報が存在するが送金先が存在しないデータのチェック '20161011 送金ルール送金先有無による送金ルール基本情報の移行制御 -add
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <remarks></remarks>
            public void Chk_SoruleBase(SqlConnection sqlcnnv10, string tblname)
            {

                var list_err = new List<string>();

                // ----------------
                // 不正データ抽出
                // ----------------

                // 送金ルール基本情報が存在するが送金先が存在しないデータを抽出(ログ出力用に成形済み)
                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) AS [ログ対象データ] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*20161107 自社物件の送金ルール作成方法修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*SELECT * FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kubunsyo = 1*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kubunsyo = 1 AND kanri_keitaikbn <> 4 ";
                tmp_sql = tmp_sql + " 		/*20161107 自社物件の送金ルール作成方法修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	) AS ITTO ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bkdata AS BK ";
                tmp_sql = tmp_sql + " 	ON ITTO.relation_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		'物件No = ' + CONVERT(VARCHAR(MAX),bk_no) + '、' + '部屋No = ' + CONVERT(VARCHAR(MAX),hy_no) AS [ログ対象データ] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*20161107 自社物件の送金ルール作成方法修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*SELECT * FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kubunsyo = 2*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kubunsyo = 2 AND kanri_keitaikbn <> 4 ";
                tmp_sql = tmp_sql + " 		/*20161107 自社物件の送金ルール作成方法修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	) AS KBN ";
                tmp_sql = tmp_sql + " 	LEFT JOIN hydata AS HY ON KBN.relation_guid = HY.hy_guid ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " ) AS VW ";
                DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_err);

                // ----------------------------------------------
                // 不正データが存在する場合、ログ出力→削除処理
                // ----------------------------------------------
                if (list_err.Count != 0)
                {

                    // 作業用変数
                    var tmp_hash = new SafeDictionary<string, string>();       // ログ作成用の仮オブジェクト(引数として使用するだけ)
                    int tmp_logcnt = 0;       // ログ作成用の仮変数(引数として使用するだけ)
                    int tmp_cnt = 0;          // 仮変数(引数として使用するだけ)
                    var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                    var tmp_loghash = new SafeDictionary<string, string>();
                    string log_key = tblname + "-" + "so_ow_no";
                    string tmp_hubi = "物件管理情報の送金先が設定されていないため移行できません。";
                    string errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + tmp_hubi + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                    tmp_loghash.Add(log_key, errstr);

                    // ログ出力
                    foreach (var str_logkey in list_err)
                    {

                        // ログ出力メッセージ整形
                        LogSetting.Set_Log_Value_KomkErr(tmp_loghash, tmp_hash, tmp_hash, str_logkey, tblname, ref tmp_logcnt, ref sortlist_log);

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

                    // 不正データ削除
                    // 20161107 自社物件の送金ルール作成方法修正 -chg sta
                    // Dim tmp_sql_ittodelete As String = " DELETE FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) "
                    string tmp_sql_ittodelete = " DELETE FROM sorule WHERE sorule_guid NOT IN (SELECT sorule_guid FROM sorule_sosaki) AND kanri_keitaikbn <> 4 ";
                    // 20161107 自社物件の送金ルール作成方法修正 -chg end
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_ittodelete, ref tmp_cnt);

                }

            }

            /// <summary>
            /// 送金先情報のダミーレコード作成処理
            /// 画面上で送金ルールを作成した際に送金先情報が自動で3レコード作成されるため、それに合わせる
            /// </summary>
            /// <param name="sqlcnnv10">DB接続用オブジェクト</param>
            /// <remarks>20161108_2 送金ルール送金先情報のダミーレコード作成処理の追加 新規追加</remarks>
            public void Set_SoruleSosakiDummyRecord(SqlConnection sqlcnnv10)
            {

                // ダミーレコード作成用クエリ
                string tmp_sql = "";
                tmp_sql = tmp_sql + " INSERT INTO sorule_sosaki ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 			DUMMYREC.* ";
                tmp_sql = tmp_sql + " 		,NULL AS so_ow_no ";
                tmp_sql = tmp_sql + " 		,NULL AS so_ow_kozano ";
                tmp_sql = tmp_sql + " 		,NULL AS so_rit ";
                tmp_sql = tmp_sql + " 		,NULL AS so_fixsogak ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT DISTINCT ";
                tmp_sql = tmp_sql + " 				sorule_guid ";
                tmp_sql = tmp_sql + " 			,1 AS sorule_no ";
                tmp_sql = tmp_sql + " 			,1 AS sosaki_recno ";
                tmp_sql = tmp_sql + " 		FROM sorule_sosaki ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT DISTINCT ";
                tmp_sql = tmp_sql + " 				sorule_guid ";
                tmp_sql = tmp_sql + " 			,1 AS sorule_no ";
                tmp_sql = tmp_sql + " 			,2 AS sosaki_recno ";
                tmp_sql = tmp_sql + " 		FROM sorule_sosaki ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT DISTINCT ";
                tmp_sql = tmp_sql + " 				sorule_guid ";
                tmp_sql = tmp_sql + " 			,1 AS sorule_no ";
                tmp_sql = tmp_sql + " 			,3 AS sosaki_recno ";
                tmp_sql = tmp_sql + " 		FROM sorule_sosaki ";
                tmp_sql = tmp_sql + " 	) AS DUMMYREC ";
                tmp_sql = tmp_sql + " 	WHERE NOT EXISTS ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM sorule_sosaki AS SOS ";
                tmp_sql = tmp_sql + " 			WHERE DUMMYREC.sorule_guid = SOS.sorule_guid ";
                tmp_sql = tmp_sql + " 			AND   DUMMYREC.sorule_no = SOS.sorule_no ";
                tmp_sql = tmp_sql + " 			AND   DUMMYREC.sosaki_recno = SOS.sosaki_recno ";
                tmp_sql = tmp_sql + " 		) ";

                int tmp_cnt = 0;
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);

            }

        }

    }

    #endregion

    #region 送金ルール入金項目情報

    public class Sorule_nk_cmrule_Repository
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
                var hash_bkhyguid = new SafeDictionary<string, string>();                              // bkguid格納用ハッシュテーブル
                var hash_bkhysoguid = new SafeDictionary<string, string>();                            // bkguid/soruleguid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Sorule_nk_cmrule_Model();        // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列                    
                string viewnamebkhytoguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhytosoruleguid = CommonModule.PRE_VIEW_NAME + "物件部屋送金キー情報";

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Nkinkomk();

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

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytoguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytosoruleguid);

                // テーブル名/フィールド名セット
                string tblname = "sorule_nk_cmrule";
                string fldnamegrp = "sorule_guid,sorule_no,taisyokbn,nkin_sortorder,nkin_no," + "sokin_rit,kanrigak_rit,hosyo_flg,so_no";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhytoguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhytosoruleguid, ref hash_bkhysoguid);

                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, viewnamebkhytosoruleguid, ref list_basekeydata);

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

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol_main - 1].ColumnName.Trim();
                string fldname_keysub1 = readtbl.Columns[keycol_sub1 - 1].ColumnName.Trim();
                string fldname_keysub2 = readtbl.Columns[keycol_sub2 - 1].ColumnName.Trim();
                string fldname_keysub3 = readtbl.Columns[keycol_sub3 - 1].ColumnName.Trim();
                string fldname_keysub4 = readtbl.Columns[keycol_sub4 - 1].ColumnName.Trim();

                // ---------------
                // データ部処理
                // ---------------
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
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string tmp_keysub3 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub3 - 1]).Trim();
                    string tmp_keysub4 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub4 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4;
                    model_cvitem.Vari_Sorule_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




                    // 作業用変数
                    string tmp_nkinname = "";

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
                            case "送金ルール管理No":
                                {
                                    model_cvitem.Vari_Sorule_no = fldvalue;
                                    break;
                                }
                            case "月々/契約時/更新時区分":
                                {
                                    model_cvitem.Vari_Taisyokbn = fldvalue;
                                    break;
                                }
                            case "行NO":
                                {
                                    model_cvitem.Vari_Nkin_sortorder = fldvalue;
                                    break;
                                }
                            case "入金項目名":
                                {
                                    tmp_nkinname = fldvalue;
                                    break;
                                }
                            case "送金率":
                                {
                                    model_cvitem.Vari_Sokin_rit = fldvalue;
                                    break;
                                }
                            case "管理手数料率":
                                {
                                    model_cvitem.Vari_Kanrigak_rit = fldvalue;
                                    break;
                                }
                            case "滞納保証有無":
                                {
                                    model_cvitem.Vari_Hosyo_flg = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 紐付用に成形
                    string localGet_Nkinruiname() { string argtukikbn = model_cvitem.Vari_Taisyokbn; var ret = EtcMethod.Get_Nkinruiname(ref argtukikbn); model_cvitem.Vari_Taisyokbn = argtukikbn; return ret; }

                    model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + localGet_Nkinruiname();

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
                        hash_cvitem["sorule_guid"] = hash_bkhysoguid[hash_cvitem["sorule_guid"]];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;
                        }
                        else    // 20161014 重複エラーのログ出力処理を追加 -add
                        {
                            string log_key = "sorule_nk_cmrule-sorule_guid" + "/" + "sorule_nk_cmrule-sorule_no" + "/" + "sorule_nk_cmrule-taisyokbn" + "/" + "sorule_nk_cmrule-nkin_sortorder" + "/" + "sorule_nk_cmrule-nkin_no";
                            string errstr = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP;
                            hash_log.Clear();
                            hash_log.Add(log_key, errstr);
                            tmp_hash["sorule_guid"] = "";
                            hash_cvitem["sorule_guid"] = "";
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

                // その他請求レコード一括挿入処理
                int tmpcnt = 0;
                string othernkkomksql = Get_UseQry_Insert();
                DBExec.Exec_NonQuery(sqlcnnv10, othernkkomksql, ref tmpcnt);

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

                // 入金項目の重複チェックは特殊なので考慮すること

                // Dim tmp_sql As String = ""

                // tmp_sql = tmp_sql & " SELECT * FROM "
                // tmp_sql = tmp_sql & " ( "
                // tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + '' + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                // tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                // tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON sorule.relation_guid = bkdata.bk_guid "
                // tmp_sql = tmp_sql & " 	UNION "
                // tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                // tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                // tmp_sql = tmp_sql & " 	LEFT JOIN hydata ON sorule.relation_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                // tmp_sql = tmp_sql & " ) AS VW "
                // tmp_sql = tmp_sql & " WHERE キー IS NOT NULL "


                // Dim flg As Boolean = True

                // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件部屋キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,bk_no)  + '-' AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,bk_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM bkdata ";
                            tmp_sql_create = tmp_sql_create + " UNION ";
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋送金キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + '-' + CONVERT(varchar,BKS.kn_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,sorule_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM bkdata_syo AS BKS ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON BKS.bk_guid = BK.bk_guid ";
                            tmp_sql_create = tmp_sql_create + " UNION ";
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no)  + '-' + CONVERT(varchar,HYS.kn_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,sorule_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata_syo AS HYS ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON HYS.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 送金ルールのその他請求レコード挿入クエリ '20160928 送金ルールその他請求取得処理修正
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Insert()
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " INSERT INTO sorule_nk_cmrule ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 sorule_guid ";
                tmp_sql = tmp_sql + " 	,sorule_no ";
                tmp_sql = tmp_sql + " 	,7 AS taisyokbn ";
                tmp_sql = tmp_sql + " 	,0 AS nkin_sortorder ";
                tmp_sql = tmp_sql + " 	,0 AS nkin_no ";
                tmp_sql = tmp_sql + " 	,0 AS sokin_rit ";
                tmp_sql = tmp_sql + " 	,0 AS kanrigak_rit ";
                tmp_sql = tmp_sql + " 	,hosyo_flg ";
                tmp_sql = tmp_sql + " 	,NULL AS so_no ";
                tmp_sql = tmp_sql + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT DISTINCT sorule_guid,sorule_no,hosyo_flg FROM sorule_nk_cmrule ";
                tmp_sql = tmp_sql + " ) AS VW ";
                return tmp_sql;


            }

            // 2016.04.06 入金項目読込処理の修正 -del sta
            // ''' <summary>
            // ''' 入金項目紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "入金項目マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // '既存データ有無確認(入金項目は他の項目でも参照するため)
            // If Hash_Rel_Nkinkomk.Count <> 0 Then
            // Exit Sub
            // End If

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldnkinname As String = ""
            // Dim tmp_newnkinno As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元入金項目名称"
            // tmp_oldnkinname = fldvalue.Trim
            // Case "賃貸革命10入金項目No"
            // tmp_newnkinno = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newnkinno <> "" Then
            // Hash_Rel_Nkinkomk.Add(tmp_oldnkinname, tmp_newnkinno)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 入金項目読込処理の修正 -del end

        }

    }

    #endregion

    #region 送金ルール控除項目情報

    public class Sorule_kojo_cmrule_Repository
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
                var hash_bkhyguid = new SafeDictionary<string, string>();                                // bkguid格納用ハッシュテーブル
                var hash_bkhysoguid = new SafeDictionary<string, string>();                              // bkguid/soruleguid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Sorule_kojo_cmrule_Model();      // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebkhytoguid = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";
                string viewnamebkhytosoruleguid = CommonModule.PRE_VIEW_NAME + "物件部屋送金キー情報";

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 2016.04.06 入金項目読込処理の修正 -chg sta
                // Call Me.Get_RelData()
                SetRelItemToObject.Set_RelData_Nkinkomk();
                // 2016.04.06 入金項目読込処理の修正 -chg end
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
                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytoguid);
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytosoruleguid);

                // テーブル名/フィールド名セット
                string tblname = "sorule_kojo_cmrule";
                string fldnamegrp = "sorule_guid,sorule_no,kojo_taisyokbn,nkin_sortorder,taisyonkin_no," + "sosai_flg,kojocalc_kbn,kojo_gak,kojo_gakzeikbn,kojonkin_no," + "kojo_rit,kojo_ritzeikbn,kojo_ritutizei,zei_rit,tateazu_flg," + "kojo_zeigak,sotaisyo_flg,sotaisyonkin_no";



                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhytoguid, ref hash_bkhyguid);
                Get_Guid(sqlcnnv10, viewnamebkhytosoruleguid, ref hash_bkhysoguid);

                // 親マスタ取得
                Get_BaseKey(sqlcnnv10, viewnamebkhytosoruleguid, ref list_basekeydata);

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

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol_main - 1].ColumnName.Trim();
                string fldname_keysub1 = readtbl.Columns[keycol_sub1 - 1].ColumnName.Trim();
                string fldname_keysub2 = readtbl.Columns[keycol_sub2 - 1].ColumnName.Trim();
                string fldname_keysub3 = readtbl.Columns[keycol_sub3 - 1].ColumnName.Trim();
                string fldname_keysub4 = readtbl.Columns[keycol_sub4 - 1].ColumnName.Trim();

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
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string tmp_keysub3 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub3 - 1]).Trim();
                    string tmp_keysub4 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub4 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4;
                    model_cvitem.Vari_Sorule_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;





                    // 2016.04.06 入金項目読込処理の修正 -add
                    // 作業用変数
                    string tmp_nkinname = "";

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
                            case "送金ルール管理No":
                                {
                                    model_cvitem.Vari_Sorule_no = fldvalue.Trim();
                                    break;
                                }
                            case "契約時/更新時区分":
                                {
                                    model_cvitem.Vari_Kojo_taisyokbn = fldvalue.Trim();
                                    break;
                                }
                            case "行NO":
                                {
                                    model_cvitem.Vari_Nkin_sortorder = fldvalue.Trim();
                                    break;
                                }
                            case "控除入金項目名":
                                {
                                    // 2016.04.06 入金項目読込処理の修正 -chg sta
                                    // .Vari_Taisyonkin_no = fldvalue.Trim
                                    tmp_nkinname = fldvalue;
                                    break;
                                }
                            // 2016.04.06 入金項目読込処理の修正 -chg end
                            case "相殺予定フラグ":
                                {
                                    model_cvitem.Vari_Sosai_flg = fldvalue.Trim();
                                    break;
                                }
                            case "控除額基準":
                                {
                                    model_cvitem.Vari_Kojocalc_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "控除額":
                                {
                                    model_cvitem.Vari_Kojo_gak = fldvalue.Trim();
                                    break;
                                }
                            case "控除額税区分":
                                {
                                    model_cvitem.Vari_Kojo_gakzeikbn = fldvalue.Trim();
                                    break;
                                }
                            case "控除対象入金項目名":
                                {
                                    model_cvitem.Vari_Kojonkin_no = fldvalue.Trim();
                                    break;
                                }
                            case "控除率":
                                {
                                    model_cvitem.Vari_Kojo_rit = fldvalue.Trim();
                                    break;
                                }
                            case "控除率税区分":
                                {
                                    model_cvitem.Vari_Kojo_ritzeikbn = fldvalue.Trim();
                                    break;
                                }
                            case "控除率税有無":
                                {
                                    model_cvitem.Vari_Kojo_ritutizei = fldvalue.Trim();
                                    break;
                                }
                            case "適用税率":
                                {
                                    model_cvitem.Vari_Zei_rit = fldvalue.Trim();
                                    break;
                                }
                            case "立替回収または預り金":
                                {
                                    model_cvitem.Vari_Tateazu_flg = fldvalue.Trim();
                                    break;
                                }
                            case "控除税額":
                                {
                                    model_cvitem.Vari_Kojo_zeigak = fldvalue.Trim();
                                    break;
                                }
                            case "送金対象フラグ":
                                {
                                    model_cvitem.Vari_Sotaisyo_flg = fldvalue.Trim();
                                    break;
                                }
                            case "送金対象入金項目名":
                                {
                                    model_cvitem.Vari_Sotaisyonkin_no = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 2016.04.06 入金項目読込処理の修正 -add sta
                    // 紐付用に成形
                    string argtukikbn = "9";
                    model_cvitem.Vari_Taisyonkin_no = tmp_nkinname + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn);
                    // 2016.04.06 入金項目読込処理の修正 -add end

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
                        hash_cvitem["sorule_guid"] = hash_bkhysoguid[hash_cvitem["sorule_guid"]];

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
                midrowcnt = rowcnt;

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

                // 入金項目の重複チェックは特殊なので考慮すること

                // Dim tmp_sql As String = ""

                // tmp_sql = tmp_sql & " SELECT * FROM "
                // tmp_sql = tmp_sql & " ( "
                // tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + '' + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                // tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                // tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON sorule.relation_guid = bkdata.bk_guid "
                // tmp_sql = tmp_sql & " 	UNION "
                // tmp_sql = tmp_sql & " 	SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,sorule.sorule_no) + '-' + CONVERT(varchar,sosaki_recno) AS キー FROM " & tblname
                // tmp_sql = tmp_sql & " 	LEFT JOIN sorule ON " & tblname & ".sorule_guid = sorule.sorule_guid "
                // tmp_sql = tmp_sql & " 	LEFT JOIN hydata ON sorule.relation_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " 	LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                // tmp_sql = tmp_sql & " ) AS VW "
                // tmp_sql = tmp_sql & " WHERE キー IS NOT NULL "


                // Dim flg As Boolean = True

                // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

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
                string tmp_sql_drop = DBQuery.Qry_DropInfo(viewname, false);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                string tmp_viewmainname = viewname.Replace(CommonModule.PRE_VIEW_NAME, "");

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                switch (tmp_viewmainname ?? "")
                {
                    case "物件部屋キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,bk_no)  + '-' AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,bk_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM bkdata ";
                            tmp_sql_create = tmp_sql_create + " UNION ";
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                    case "物件部屋送金キー情報":
                        {
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + '-' + CONVERT(varchar,BKS.kn_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,sorule_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM bkdata_syo AS BKS ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON BKS.bk_guid = BK.bk_guid ";
                            tmp_sql_create = tmp_sql_create + " UNION ";
                            tmp_sql_create = tmp_sql_create + " SELECT ";
                            tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no)  + '-' + CONVERT(varchar,HYS.kn_no) AS キー ";
                            tmp_sql_create = tmp_sql_create + " 	,sorule_guid ";
                            tmp_sql_create = tmp_sql_create + " FROM hydata_syo AS HYS ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON HYS.hy_guid = HY.hy_guid ";
                            tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                            break;
                        }
                }

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            // 2016.04.06 入金項目読込処理の修正 -del sta
            // ''' <summary>
            // ''' 入金項目紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "入金項目マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // '既存データ有無確認(入金項目は他の項目でも参照するため)
            // If Hash_Rel_Nkinkomk.Count <> 0 Then
            // Exit Sub
            // End If

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldnkinname As String = ""
            // Dim tmp_newnkinno As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元入金項目名称"
            // tmp_oldnkinname = fldvalue.Trim
            // Case "賃貸革命10入金項目No"
            // tmp_newnkinno = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newnkinno <> "" Then
            // Hash_Rel_Nkinkomk.Add(tmp_oldnkinname, tmp_newnkinno)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 入金項目読込処理の修正 -del end

        }

    }

    #endregion

}