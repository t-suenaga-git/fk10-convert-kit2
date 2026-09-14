using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Converter10.Njc.N3Lib.Utys;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Repository
{

    #region 契約基本情報

    public class Kydata_Repository
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
                var hash_bkguid = new SafeDictionary<string, string>();                                // bkguid格納用ハッシュテーブル
                var hash_hyguid = new SafeDictionary<string, string>();                                // bkguid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_Model();                  // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = CommonModule.PRE_VIEW_NAME + "物件部屋キー情報";

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
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewname);

                // テーブル名/フィールド名セット
                string tblname = "kydata";
                string fldnamegrp = "ky_guid,bk_guid,hy_guid,ky_no,ky_deleteflg," + "delete_guid,delete_day,delete_cnt,syokai_kyymd,status," + "cancelriyu,status_ymd,cyukai_gy_fudono,syunin_logonuser_no,tetuke_gak1," + "tetuke_ymd1,tetuke_biko1,kaiyaku_flg,history,rowid," + "movefrom_kyguid,moveto_kyguid,svbunrui_no,krbunrui_no,kaiyaku_uketukekbn," + "kaiyaku_months,kaiyaku_days,kaiyaku_day,cyukai_tantoname,cyukai_tantonamesjis";





                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname, ref hash_bkguid, 0);  // bk_guid取得
                Get_Guid(sqlcnnv10, viewname, ref hash_hyguid, 1);  // hy_guid取得

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_hyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                        // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                        // excelfile.ExcelFile_ReadClose(con_read)
                        return rtn;
                    }

                    // 作業用変数
                    string fldname = "";
                    string fldvalue = "";
                    string tmp_kystartymd = "";
                    string tmp_kyendymd = "";

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    model_cvitem.Vari_Bk_guid = tmp_keymain + "-" + tmp_keysub1;
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

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
                            case "契約No":
                                {
                                    model_cvitem.Vari_Ky_no = fldvalue;
                                    break;
                                }
                            case "初回契約日":
                                {
                                    model_cvitem.Vari_Syokai_kyymd = fldvalue;
                                    break;
                                }
                            case "契約状況(ステータス)":
                                {
                                    model_cvitem.Vari_Status = fldvalue;
                                    break;
                                }
                            case "キャンセル理由":
                                {
                                    model_cvitem.Vari_Cancelriyu = fldvalue;
                                    break;
                                }
                            case "ステータス変更日":
                                {
                                    model_cvitem.Vari_Status_ymd = fldvalue;
                                    break;
                                }
                            case "仲介業者No":
                                {
                                    model_cvitem.Vari_Cyukai_gy_fudono = fldvalue;
                                    break;
                                }
                            case "取引主任者No":
                                {
                                    model_cvitem.Vari_Syunin_logonuser_no = fldvalue;
                                    break;
                                }
                            case "手付預り額①":
                                {
                                    model_cvitem.Vari_Tetuke_gak1 = fldvalue;
                                    break;
                                }
                            case "手付預り日①":
                                {
                                    model_cvitem.Vari_Tetuke_ymd1 = fldvalue;
                                    break;
                                }
                            case "手付預り備考①":
                                {
                                    model_cvitem.Vari_Tetuke_biko1 = fldvalue;
                                    break;
                                }
                            case "解約フラグ":
                                {
                                    model_cvitem.Vari_Kaiyaku_flg = fldvalue;
                                    break;
                                }
                            case "サービス分類":
                                {
                                    model_cvitem.Vari_Svbunrui_no = fldvalue;
                                    break;
                                }
                            case "会計グループ分類":
                                {
                                    model_cvitem.Vari_Krbunrui_no = fldvalue;
                                    break;
                                }
                            case "解約受付区分":
                                {
                                    model_cvitem.Vari_Kaiyaku_uketukekbn = fldvalue;
                                    break;
                                }
                            case "解約受付月数":
                                {
                                    model_cvitem.Vari_Kaiyaku_months = fldvalue;
                                    break;
                                }
                            case "解約受付日数":
                                {
                                    model_cvitem.Vari_Kaiyaku_days = fldvalue;
                                    break;
                                }
                            case "解約受付日にち":
                                {
                                    model_cvitem.Vari_Kaiyaku_day = fldvalue;
                                    break;
                                }
                            case "仲介担当者名":
                                {
                                    model_cvitem.Vari_Cyukai_tantoname = fldvalue;
                                    break;
                                }
                            case "仲介担当者名SJIS":
                                {
                                    model_cvitem.Vari_Cyukai_tantonamesjis = fldvalue;
                                    break;
                                }
                            // 2016.04.06 仮契約情報の移行制御処理を追加 -add sta
                            case "契約開始日":
                                {
                                    tmp_kystartymd = fldvalue;
                                    break;
                                }
                            case "契約終了日":
                                {
                                    tmp_kyendymd = fldvalue;
                                    break;
                                }
                                // 2016.04.06 仮契約情報の移行制御処理を追加 -add end
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Ky_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Ky_deleteflg = 0.ToString();
                    model_cvitem.Vari_Delete_guid = "";
                    model_cvitem.Vari_Delete_day = "";
                    model_cvitem.Vari_Delete_cnt = "";
                    model_cvitem.Vari_Movefrom_kyguid = "";
                    model_cvitem.Vari_Moveto_kyguid = "";
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();

                    // 20170530 契約状況確定日項目追加対応 -add sta
                    // 契約状況とステータス変更日を見て値の有無を確認・設定する
                    if (model_cvitem.Vari_Status == "2" || model_cvitem.Vari_Status == "3")
                    {
                        if (string.IsNullOrEmpty(model_cvitem.Vari_Status_ymd))
                        {
                            // 空の場合は契約開始日を設定する
                            model_cvitem.Vari_Status_ymd = tmp_kystartymd;
                        }
                    }
                    // 20170530 契約状況確定日項目追加対応 -add end

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 2016.04.06 仮契約情報の移行制御処理を追加 -add sta
                    // *****************************************************************************************
                    // ※制御自体は可能だが、ログの出力に個別に対応が必要
                    // ※契約基本情報のみにチェックを入れた場合はテーブル名とフィールド名が出力されない
                    // (契約履歴情報のテーブル名、フィールド名をログに出力するため)
                    // ※ただし、革命上では契約基本情報と契約履歴情報は同時に作成される。そのため
                    // ・どちらかのみの移行を不可にする等の制御を設ける (必ずセットで移行する)
                    // ・移行項目の統合で自動で両方移行するようにする (統合する予定なのでおそらくこれになる)
                    // *****************************************************************************************
                    if (string.IsNullOrEmpty(tmp_kystartymd))
                    {
                        hash_log.Add("kydata_kihon-kystart_ymd", CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED);
                        skipflg = true;
                    }
                    else if (string.IsNullOrEmpty(tmp_kyendymd))
                    {
                        hash_log.Add("kydata_kihon-kyend_ymd", CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED);
                        skipflg = true;
                    }
                    // 2016.04.06 仮契約情報の移行制御処理を追加 -add end

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("bk_guid") = hash_bkguid.Item(hash_cvitem.Item("bk_guid"))
                        // 'hash_cvitem("hy_guid") = hash_hyguid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("bk_guid") = hash_bkguid.Item(StrConv(hash_cvitem.Item("bk_guid"), VbStrConv.Narrow))
                        // hash_cvitem("hy_guid") = hash_hyguid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["bk_guid"] = hash_bkguid[hash_cvitem["bk_guid"]];
                        hash_cvitem["hy_guid"] = hash_hyguid[hash_cvitem["hy_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewname);

                // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
                // 'クローズ処理
                // excelfile.ExcelFile_ReadClose(con_read)
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
                // 20160929 汎用CV時のステータス変更日の一括更新処理 -add sta
                // 仮契約の場合はステータス変更日NULLに一括更新
                if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                {
                    int tmpcnt = 0;
                    string stymdupdatesql = " UPDATE kydata SET status_ymd = NULL WHERE [status] = 1 ";
                    DBExec.Exec_NonQuery(sqlcnnv10, stymdupdatesql, ref tmpcnt);
                }
                // 20160929 汎用CV時のステータス変更日の一括更新処理 -add end
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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
            public void Get_Guid(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_guid, int bkhytype)
            {

                string fldname = "";

                switch (bkhytype)
                {
                    case 0:
                        {
                            fldname = " キー,bk_guid ";
                            break;
                        }
                    case 1:
                        {
                            fldname = " キー,hy_guid ";
                            break;
                        }
                }

                string tmp_sql = " SELECT " + fldname + " FROM " + tblname;

                bool flg = true;
                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end
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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.bk_guid ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
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

    #region 契約履歴情報

    public class Kydata_kihon_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                              // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_kihon_Model();       // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 2016.04.06 紐付データ取得処理を外出し -chg sta
                // Call Me.Get_RelData_NkinKbn()
                SetRelItemToObject.Set_RelData_Nkinkbn();
                // 2016.04.06 紐付データ取得処理を外出し -chg end

                // 2016.04.06 契約分類を紐付データを元に移行 -add
                SetRelItemToObject.Set_RelData_Kyrui();
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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_kihon";
                // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add (torihiki_tesumoto,torihiki_tesukyaku,torihiki_tesugak,torihiki_tesuzeikbn,torihiki_tesuzeigak)を追加
                // 20160519 EXEUpdateに伴う修正 契約履歴情報 -chg sta
                // Dim fldnamegrp As String = "ky_guid,ky_recno,ko_no,henko_no,sqdata_guid," & _
                // "ky_bango,ky_ymd,kystart_ymd,kyend_ymd,henko_ymd," & _
                // "tuti_ymd,print_ymd,kanri_gy_fudono,gy_hosyono,hosyo_naiyo," & _
                // "kokyaku_bango,kyrui_no,siyo_mokuteki,kosin_umu,yatin_kbn," & _
                // "fkae_startym,fkae_willstartflg,yatin_kozakbn,yatin_kozano,maitukiyatin_kozano," & _
                // "yokugetu_uketoriflg,biko,tougetu_sagakuflg,nextky_startymd,nextky_endymd," & _
                // "kosin_hiwariflg,sokojorule_kbn,soyotei_ymdflg,soyotei_ymd,kanritesu_flg," & _
                // "kanritesu_gak,nextnkinset_kbn,sqdata_startymd,sqdata_endymd,history," & _
                // "biko2,hoken_biko,nkinsime_ymd,yatin_jisansaki,hoken_kikan," & _
                // "hoken_gak,confirmky_sekininsya,confirmky_ymd,confirmky_print,confirmkai_sekininsya," & _
                // "confirmkai_ymd,confirmkai_print,hikiuke_name,hikiuke_addr,hikiuke_tel," & _
                // "kohokennkin_ymd,kokanryo_ymd,kokanryotuti_ymd,kotuti_ymd,kosaisoku_ymd," & _
                // "kosyoruiuke_ymd,kokairenraku_umu,kokairenraku_ymd,kokairenraku_logonuser_no,kokairenraku_biko," & _
                // "maitukisq_umu,hikiuke_namesjis,kofkae_ymd,yokugetu_uketorimonth,ky_logonuser_no," & _
                // "sq_logonuser_no"
                string fldnamegrp = "ky_guid,ky_recno,ko_no,henko_no,sqdata_guid," + "ky_bango,ky_ymd,kystart_ymd,kyend_ymd,henko_ymd," + "tuti_ymd,print_ymd,kanri_gy_fudono,gy_hosyono,hosyo_naiyo," + "kokyaku_bango,kyrui_no,siyo_mokuteki,kosin_umu,yatin_kbn," + "fkae_startym,fkae_willstartflg,yatin_kozakbn,yatin_kozano,maitukiyatin_kozano," + "yokugetu_uketoriflg,biko,tougetu_sagakuflg,nextky_startymd,nextky_endymd," + "kosin_hiwariflg,sokojorule_kbn,soyotei_ymdflg,soyotei_ymd,kanritesu_flg," + "kanritesu_gak,nextnkinset_kbn,sqdata_startymd,sqdata_endymd,history," + "biko2,hoken_biko,nkinsime_ymd,yatin_jisansaki,hoken_kikan," + "hoken_gak,hikiuke_name,hikiuke_addr,hikiuke_tel,kohokennkin_ymd," + "kokanryo_ymd,kokanryotuti_ymd,kotuti_ymd,kosaisoku_ymd,kosyoruiuke_ymd," + "maitukisq_umu,hikiuke_namesjis,kofkae_ymd,yokugetu_uketorimonth,ky_logonuser_no," + "sq_logonuser_no,torihiki_tesumoto,torihiki_tesukyaku,torihiki_tesugak,torihiki_tesuzeikbn," + "torihiki_tesuzeigak";












                // 20160519 EXEUpdateに伴う修正 契約履歴情報 -chg end
                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "更新No":
                                {
                                    model_cvitem.Vari_Ko_no = fldvalue.Trim();
                                    break;
                                }
                            case "改定No":
                                {
                                    model_cvitem.Vari_Henko_no = fldvalue.Trim();
                                    break;
                                }
                            case "契約番号":
                                {
                                    model_cvitem.Vari_Ky_bango = fldvalue.Trim();
                                    break;
                                }
                            case "契約日":
                                {
                                    model_cvitem.Vari_Ky_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "契約開始日":
                                {
                                    model_cvitem.Vari_Kystart_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "契約終了日":
                                {
                                    model_cvitem.Vari_Kyend_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "条件変更日":
                                {
                                    model_cvitem.Vari_Henko_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "契約更新通知日":
                                {
                                    model_cvitem.Vari_Tuti_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "通知書印刷日":
                                {
                                    model_cvitem.Vari_Print_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "管理業者No":
                                {
                                    model_cvitem.Vari_Kanri_gy_fudono = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸保証業者No":
                                {
                                    model_cvitem.Vari_Gy_hosyono = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸保証内容":
                                {
                                    model_cvitem.Vari_Hosyo_naiyo = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸保証顧客番号":
                                {
                                    model_cvitem.Vari_Kokyaku_bango = fldvalue.Trim();
                                    break;
                                }
                            case "契約分類名":   // 2016.04.06 契約分類を紐付データを元に移行 「契約分類No」→「契約分類名」へ変更
                                {
                                    model_cvitem.Vari_Kyrui_no = fldvalue.Trim();
                                    break;
                                }
                            case "使用目的":
                                {
                                    model_cvitem.Vari_Siyo_mokuteki = fldvalue.Trim();
                                    break;
                                }
                            case "契約更新業務有無":
                                {
                                    model_cvitem.Vari_Kosin_umu = fldvalue.Trim();
                                    break;
                                }
                            case "家賃入金区分":
                                {
                                    model_cvitem.Vari_Yatin_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "口座振替開始日":
                                {
                                    model_cvitem.Vari_Fkae_startym = fldvalue.Trim();
                                    break;
                                }
                            case "口座振替開始待ちフラグ":
                                {
                                    model_cvitem.Vari_Fkae_willstartflg = fldvalue.Trim();
                                    break;
                                }
                            case "家賃入金口座区分":
                                {
                                    model_cvitem.Vari_Yatin_kozakbn = fldvalue.Trim();
                                    break;
                                }
                            case "契約一時金入金口座No":
                                {
                                    model_cvitem.Vari_Yatin_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "毎月分入金口座No":
                                {
                                    model_cvitem.Vari_Maitukiyatin_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "翌月分受取り有無":
                                {
                                    model_cvitem.Vari_Yokugetu_uketoriflg = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "当月分の差額受取り有無":
                                {
                                    model_cvitem.Vari_Tougetu_sagakuflg = fldvalue.Trim();
                                    break;
                                }
                            case "次回契約開始日":
                                {
                                    model_cvitem.Vari_Nextky_startymd = fldvalue.Trim();
                                    break;
                                }
                            case "次回契約終了日":
                                {
                                    model_cvitem.Vari_Nextky_endymd = fldvalue.Trim();
                                    break;
                                }
                            case "更新時日割り有無":
                                {
                                    model_cvitem.Vari_Kosin_hiwariflg = fldvalue.Trim();
                                    break;
                                }
                            case "送金控除ルール適用有無":
                                {
                                    model_cvitem.Vari_Sokojorule_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "送金予定日使用フラグ":
                                {
                                    model_cvitem.Vari_Soyotei_ymdflg = fldvalue.Trim();
                                    break;
                                }
                            case "送金予定日":
                                {
                                    model_cvitem.Vari_Soyotei_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "部屋固定管理手数料フラグ":
                                {
                                    model_cvitem.Vari_Kanritesu_flg = fldvalue.Trim();
                                    break;
                                }
                            case "部屋固定管理手数料額":
                                {
                                    model_cvitem.Vari_Kanritesu_gak = fldvalue.Trim();
                                    break;
                                }
                            case "次回更新設定区分":
                                {
                                    model_cvitem.Vari_Nextnkinset_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "請求データ作成開始日":
                                {
                                    model_cvitem.Vari_Sqdata_startymd = fldvalue.Trim();
                                    break;
                                }
                            case "請求データ作成終了日":
                                {
                                    model_cvitem.Vari_Sqdata_endymd = fldvalue.Trim();
                                    break;
                                }
                            case "備考2(基本)":
                                {
                                    model_cvitem.Vari_Biko2 = fldvalue.Trim();
                                    break;
                                }
                            case "備考(保険)":
                                {
                                    model_cvitem.Vari_Hoken_biko = fldvalue.Trim();
                                    break;
                                }
                            case "入金締め日":
                                {
                                    model_cvitem.Vari_Nkinsime_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "家賃持参先":
                                {
                                    model_cvitem.Vari_Yatin_jisansaki = fldvalue.Trim();
                                    break;
                                }
                            case "保険期間":
                                {
                                    model_cvitem.Vari_Hoken_kikan = fldvalue.Trim();
                                    break;
                                }
                            case "保険額":
                                {
                                    model_cvitem.Vari_Hoken_gak = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 契約履歴情報 -del sta
                            // Case "確認事項(契約)責任者"
                            // .Vari_Confirmky_sekininsya = fldvalue.Trim
                            // Case "確認事項(契約)内容確認日"
                            // .Vari_Confirmky_ymd = fldvalue.Trim
                            // Case "確認事項(契約)印刷時"
                            // .Vari_Confirmky_print = fldvalue.Trim
                            // Case "確認事項(解約)責任者"
                            // .Vari_Confirmkai_sekininsya = fldvalue.Trim
                            // Case "確認事項(解約)内容確認日"
                            // .Vari_Confirmkai_ymd = fldvalue.Trim
                            // Case "確認事項(解約)印刷時"
                            // .Vari_Confirmkai_print = fldvalue.Trim
                            // 20160519 EXEUpdateに伴う修正 契約履歴情報 -del end
                            case "身元引受人名":
                                {
                                    model_cvitem.Vari_Hikiuke_name = fldvalue.Trim();
                                    break;
                                }
                            case "身元引受人住所":
                                {
                                    model_cvitem.Vari_Hikiuke_addr = fldvalue.Trim();
                                    break;
                                }
                            case "身元引受人連絡先":
                                {
                                    model_cvitem.Vari_Hikiuke_tel = fldvalue.Trim();
                                    break;
                                }
                            case "保険料入金日":
                                {
                                    model_cvitem.Vari_Kohokennkin_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "更新完了日":
                                {
                                    model_cvitem.Vari_Kokanryo_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "更新完了通知日":
                                {
                                    model_cvitem.Vari_Kokanryotuti_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "更新通知日":
                                {
                                    model_cvitem.Vari_Kotuti_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "催促実施日":
                                {
                                    model_cvitem.Vari_Kosaisoku_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "書類返送受取日":
                                {
                                    model_cvitem.Vari_Kosyoruiuke_ymd = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 契約履歴情報 -del sta
                            // Case "更新時の解約検討連絡有無"
                            // .Vari_Kokairenraku_umu = fldvalue.Trim
                            // Case "更新時の解約検討連絡受付日"
                            // .Vari_Kokairenraku_ymd = fldvalue.Trim
                            // Case "更新時の解約検討連絡受付担当者"
                            // .Vari_Kokairenraku_logonuser_no = fldvalue.Trim
                            // Case "更新時の解約検討連絡備考"
                            // .Vari_Kokairenraku_biko = fldvalue.Trim
                            // 20160519 EXEUpdateに伴う修正 契約履歴情報 -del end
                            case "毎月分請求有無":
                                {
                                    model_cvitem.Vari_Maitukisq_umu = fldvalue.Trim();
                                    break;
                                }
                            case "身元引受人名SJIS":
                                {
                                    model_cvitem.Vari_Hikiuke_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "口座振替日":
                                {
                                    model_cvitem.Vari_Kofkae_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "翌月分受取り月":
                                {
                                    model_cvitem.Vari_Yokugetu_uketorimonth = fldvalue.Trim();
                                    break;
                                }
                            case "契約担当者No":
                                {
                                    model_cvitem.Vari_Ky_logonuser_no = fldvalue.Trim();
                                    break;
                                }
                            case "請求担当者No":
                                {
                                    model_cvitem.Vari_Sq_logonuser_no = fldvalue.Trim();
                                    break;
                                }
                            // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add sta
                            case "配分割合_元付":
                                {
                                    model_cvitem.Vari_Torihiki_tesumoto = fldvalue;
                                    break;
                                }
                            case "配分割合_客付":
                                {
                                    model_cvitem.Vari_Torihiki_tesukyaku = fldvalue;
                                    break;
                                }
                            case "客付会社の手数料額":
                                {
                                    model_cvitem.Vari_Torihiki_tesugak = fldvalue;
                                    break;
                                }
                            case "客付会社の手数料税区分":
                                {
                                    model_cvitem.Vari_Torihiki_tesuzeikbn = fldvalue;
                                    break;
                                }
                            case "客付会社の手数料税額":
                                {
                                    model_cvitem.Vari_Torihiki_tesuzeigak = fldvalue;
                                    break;
                                }
                                // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add end
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Sqdata_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // 20170525 契約情報が自動更新の場合の次回契約期間を空にする修正対応 -add sta
                    // 自動更新の場合は次回契約期間を空にする
                    if (model_cvitem.Vari_Kosin_umu == "2")
                    {
                        model_cvitem.Vari_Nextky_startymd = "";
                        model_cvitem.Vari_Nextky_endymd = "";
                    }
                    // 20170525 契約情報が自動更新の場合の次回契約期間を空にする修正対応 -add end

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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) FROM " + tblname;
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 入金区分紐付情報取得
            /// </summary>
            /// <remarks></remarks>
            public void Get_RelData_NkinKbn()
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                string relsheetname = "入金区分マスタ";
                bool rtn = true;

                var excelfile = new ExcelFileManager();                // Excelファイル操作用

                // 既存データ有無確認(入金項目は他の項目でも参照するため)
                if (CommonModule.Hash_Rel_Nkinkbn.Count != 0)
                {
                    return;
                }

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

                for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
                {

                    // 作業用変数作成
                    string tmp_oldnkinkbnname = "";
                    string tmp_newnkinkbnno = "";

                    // 紐付設定値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        // ヘッダー格納
                        string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                        // 移行値格納
                        string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                        switch (fldname ?? "")
                        {
                            case "移行元入金区分名称":
                                {
                                    tmp_oldnkinkbnname = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸革命10入金区分No":
                                {
                                    tmp_newnkinkbnno = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // ハッシュテーブル格納
                    if (!string.IsNullOrEmpty(tmp_newnkinkbnno))
                    {
                        CommonModule.Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno);
                    }

                }

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

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

    #region 契約契約者情報

    public class Kydata_kys_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_kys_Model();         // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_kys";
                string fldnamegrp = "ky_guid,ky_recno,kys_cnt,kys_no,nyukyo_flg," + "history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "並び順No":
                                {
                                    model_cvitem.Vari_Kys_cnt = fldvalue.Trim();
                                    break;
                                }
                            case "契約者No":
                                {
                                    model_cvitem.Vari_Kys_no = fldvalue.Trim();
                                    break;
                                }
                            case "入居フラグ":
                                {
                                    model_cvitem.Vari_Nyukyo_flg = fldvalue.Trim();
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

                    // 20161130 契約者入居フラグの移行処理修正_再 -del sta
                    // '20161128 契約者入居フラグの移行処理修正 -add sta
                    // '入居フラグが0の場合は値をそのまま移行する
                    // If tmp_hash("nyukyo_flg") = "0" Then
                    // hash_cvitem("nyukyo_flg") = "0"
                    // End If
                    // '20161128 契約者入居フラグの移行処理修正 -add end
                    // 20161130 契約者入居フラグの移行処理修正_再 -del end

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 20170116 契約情報の契約者情報ダミーレコード作成処理の追加 -add sta
                if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                {
                    try
                    {
                        Set_KysDummyRecord(sqlcnnv10);
                    }
                    catch (Exception ex)
                    {
                        // ----- ログ出力 -----
                        int tmptmpcnt = 0;
                        string tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(41, sheetname, "データ調整", "データ調整中にエラーが発生しました。"), false);
                        DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref tmptmpcnt);
                    }
                }
                // 20170116 契約情報の契約者情報ダミーレコード作成処理の追加 -add end

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
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

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
            /// 契約情報の契約者情報へダミーレコードを作成する
            /// </summary>
            /// <param name="sqlcnnv10">賃貸10DB接続用オブジェクト</param>
            /// <remarks>
            /// 20170116 契約情報の契約者情報ダミーレコード作成処理の追加 新規追加
            /// 　契約情報登録時に契約者No1～3まで自動でレコードが作成されるためこれの対応を行う
            /// 　※契約者No = 2、3が未設定でも作成される
            /// 　TSから頂いたクエリをそのまま流用する
            /// </remarks>
            public void Set_KysDummyRecord(SqlConnection sqlcnnv10)
            {

                string tmp_sql = "";
                int tmpcnt = 0;

                // ①kys_cnt=1しかないデータに対してkys_cnt=2を作る
                tmp_sql = tmp_sql + " INSERT INTO  ";
                tmp_sql = tmp_sql + " 	kydata_kys ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		 ky_guid ";
                tmp_sql = tmp_sql + " 		,ky_recno ";
                tmp_sql = tmp_sql + " 		,kys_cnt ";
                tmp_sql = tmp_sql + " 		,kys_no ";
                tmp_sql = tmp_sql + " 		,nyukyo_flg ";
                tmp_sql = tmp_sql + " 		,history ";
                tmp_sql = tmp_sql + " 	) ";
                tmp_sql = tmp_sql + " select  ";
                tmp_sql = tmp_sql + " 	 ky_guid ";
                tmp_sql = tmp_sql + " 	,ky_recno ";
                tmp_sql = tmp_sql + " 	,2 kys_cnt ";
                tmp_sql = tmp_sql + " 	,null ";
                tmp_sql = tmp_sql + " 	,0 ";
                tmp_sql = tmp_sql + " 	,history ";
                tmp_sql = tmp_sql + " from kydata_kys ";
                tmp_sql = tmp_sql + " WHERE ky_guid not in (select distinct ky_guid from kydata_kys where kys_cnt in (2,3)) ";
                tmp_sql = tmp_sql + " AND kys_cnt = 1 ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmpcnt);
                tmp_sql = "";

                // ②kys_cnt=3がないデータに対してkys_cnt=3を作る
                // ※過去データ調整で対応したクエリを一部修正して流用する
                tmp_sql = tmp_sql + " INSERT INTO  ";
                tmp_sql = tmp_sql + " 	kydata_kys ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		 ky_guid ";
                tmp_sql = tmp_sql + " 		,ky_recno ";
                tmp_sql = tmp_sql + " 		,kys_cnt ";
                tmp_sql = tmp_sql + " 		,kys_no ";
                tmp_sql = tmp_sql + " 		,nyukyo_flg ";
                tmp_sql = tmp_sql + " 		,history ";
                tmp_sql = tmp_sql + " 	) ";
                tmp_sql = tmp_sql + " select  ";
                tmp_sql = tmp_sql + " 	 ky_guid ";
                tmp_sql = tmp_sql + " 	,ky_recno ";
                tmp_sql = tmp_sql + " 	,3 kys_cnt ";
                tmp_sql = tmp_sql + " 	,null ";
                tmp_sql = tmp_sql + " 	,0 ";
                tmp_sql = tmp_sql + " 	,history ";
                tmp_sql = tmp_sql + " from kydata_kys ";
                tmp_sql = tmp_sql + " WHERE ky_guid not in (select distinct ky_guid from kydata_kys where kys_cnt in (3)) ";
                tmp_sql = tmp_sql + " AND kys_cnt = 1 ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmpcnt);
                tmp_sql = "";

            }

        }

    }

    #endregion

    #region 契約入居者情報

    public class Kydata_nyukyo_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_nyukyo_Model();      // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_nyukyo";
                string fldnamegrp = "ky_guid,ky_recno,nyukyosya_cnt,kys_no,name," + "nameu,kana,gender,aidagara,birthday," + "kinmusaki,tel,mobiletel,biko,history";


                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "入居者No":
                                {
                                    model_cvitem.Vari_Nyukyosya_cnt = fldvalue.Trim();
                                    break;
                                }
                            case "契約者No":
                                {
                                    model_cvitem.Vari_Kys_no = fldvalue.Trim();
                                    break;
                                }
                            case "氏名":
                                {
                                    // "氏名"はsjisのフィールドと思われるが[name]へ移行しても10に反映されないため[nameu]へ移行する
                                    model_cvitem.Vari_Nameu = fldvalue.Trim();
                                    break;
                                }
                            case "氏名Unicode":
                                {
                                    break;
                                }
                            // "氏名"の値を[nameu]へ移行するためここでは何もしない
                            case "カナ":
                                {
                                    model_cvitem.Vari_Kana = fldvalue.Trim();
                                    break;
                                }
                            case "性別":
                                {
                                    model_cvitem.Vari_Gender = fldvalue.Trim();
                                    break;
                                }
                            case "続柄":
                                {
                                    model_cvitem.Vari_Aidagara = fldvalue.Trim();
                                    break;
                                }
                            case "生年月日":
                                {
                                    model_cvitem.Vari_Birthday = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先":
                                {
                                    model_cvitem.Vari_Kinmusaki = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先":
                                {
                                    model_cvitem.Vari_Tel = fldvalue.Trim();
                                    break;
                                }
                            case "携帯電話番号":
                                {
                                    model_cvitem.Vari_Mobiletel = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,nyukyosya_cnt) FROM " + tblname;
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

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

    #region 契約保証人情報

    public class Kydata_hosyonin_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_hosyonin_Model();    // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_hosyonin";
                string fldnamegrp = "ky_guid,ky_recno,hosyonin_cnt,kys_no,hosyonin_no," + "history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "並び順":
                                {
                                    model_cvitem.Vari_Hosyonin_cnt = fldvalue.Trim();
                                    break;
                                }
                            case "契約者No":
                                {
                                    model_cvitem.Vari_Kys_no = fldvalue.Trim();
                                    break;
                                }
                            case "保証人No":
                                {
                                    model_cvitem.Vari_Hosyonin_no = fldvalue.Trim();
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,hosyonin_cnt) FROM " + tblname;
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

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

    #region 契約車情報

    public class Kydata_car_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_car_Model();         // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

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
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_car";
                string fldnamegrp = "ky_guid,ky_recno,car_cnt,carmaker,carname," + "carcolor,carnumber,biko,history,car_kukaku";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "車情報No":
                                {
                                    model_cvitem.Vari_Car_cnt = fldvalue.Trim();
                                    break;
                                }
                            case "メーカー":
                                {
                                    model_cvitem.Vari_Carmaker = fldvalue.Trim();
                                    break;
                                }
                            case "車名":
                                {
                                    model_cvitem.Vari_Carname = fldvalue.Trim();
                                    break;
                                }
                            case "車色":
                                {
                                    model_cvitem.Vari_Carcolor = fldvalue.Trim();
                                    break;
                                }
                            case "ナンバー":
                                {
                                    model_cvitem.Vari_Carnumber = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "駐車区画":
                                {
                                    model_cvitem.Vari_Car_kukaku = fldvalue.Trim();
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,car_cnt) FROM " + tblname;
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end
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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

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

    #region 契約保険情報 (10でも履歴管理されていない)

    public class Kydata_hoken_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_hoken_Model();       // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_hoken";
                string fldnamegrp = "ky_guid,hoken_no,gy_hokenno,ky_ymd,kystart_ymd," + "kyend_ymd,hoken_gak,mankituti_flg,biko,history," + "syoken_bango";


                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



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
                            case "保険種類No":
                                {
                                    model_cvitem.Vari_Hoken_no = fldvalue.Trim();
                                    break;
                                }
                            case "保険業者No":
                                {
                                    model_cvitem.Vari_Gy_hokenno = fldvalue.Trim();
                                    break;
                                }
                            case "契約日":
                                {
                                    model_cvitem.Vari_Ky_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "適用開始年月日":
                                {
                                    model_cvitem.Vari_Kystart_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "適用終了年月日":
                                {
                                    model_cvitem.Vari_Kyend_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "保険金額":
                                {
                                    model_cvitem.Vari_Hoken_gak = fldvalue.Trim();
                                    break;
                                }
                            case "満期案内通知有無":
                                {
                                    model_cvitem.Vari_Mankituti_flg = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "証券番号":
                                {
                                    model_cvitem.Vari_Syoken_bango = fldvalue.Trim();
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,hoken_no) FROM " + tblname;
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

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

    #region 契約特約事項情報

    public class Kydata_tokuyaku_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_tokuyaku_Model();    // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";
                int totalrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_tokuyaku";
                string fldnamegrp = "ky_guid,ky_recno,tokuyaku_grpno,naiyo,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                string fldname_keysub1 = readtbl.Columns[keycol_sub1 - 1].ColumnName.Trim();
                string fldname_keysub2 = readtbl.Columns[keycol_sub2 - 1].ColumnName.Trim();
                string fldname_keysub3 = readtbl.Columns[keycol_sub3 - 1].ColumnName.Trim();
                string fldname_keysub4 = readtbl.Columns[keycol_sub4 - 1].ColumnName.Trim();

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

                    // 移行判別用
                    bool cvflg = true;
                    bool keychkflg = true;

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string tmp_keysub3 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub3 - 1]).Trim();
                    string tmp_keysub4 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub4 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4;
                    string tmp_keyoya = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    string tmp_keydup = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4;

                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    model_cvitem.Vari_Ky_recno = tmp_keysub3;
                    model_cvitem.Vari_Tokuyaku_grpno = tmp_keysub4;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




                    // 特約カウント初期化
                    int cvitemcnt = 0;
                    int tmp_cvrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                    // データ有無チェック
                    string tmp_fldvalueumuchk = "";
                    for (int cntjj = 5, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                        tmp_fldvalueumuchk = tmp_fldvalueumuchk + Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();
                    if (string.IsNullOrEmpty(tmp_fldvalueumuchk))
                    {
                        cvflg = false;
                    }

                    if (cvflg)
                    {

                        // 移行対象件数カウント
                        cvtaisyocnt = cvtaisyocnt + 1;

                        var hash_keylog = new SafeDictionary<string, string>();
                        string log_keyout = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1;
                        // 親データ有無チェック
                        if (list_basekeydata.Contains(tmp_keyoya) == false)
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "kydata_tokuyaku-ky_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keydup))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "kydata_tokuyaku-ky_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        if (keychkflg)
                        {

                            // 作業用変数作成
                            string tmp_tokuyaku = "";

                            // 移行値取得
                            for (int cntjj = 5, loopTo2 = readtbl.Columns.Count - 1; cntjj <= loopTo2; cntjj++)
                            {

                                // サブキー取得
                                string tmp_keysub5 = (cntjj - 4).ToString();

                                // ログ出力用データ格納(サブキーフィールド)
                                string fldname_keysub = readtbl.Columns[cntjj].ColumnName.Trim();

                                // 登録値取得
                                string fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();
                                if (!string.IsNullOrEmpty(fldvalue))
                                {
                                    tmp_tokuyaku = tmp_tokuyaku + CommonModule.LINE_BREAK + fldvalue;
                                }

                            }

                            // 2016.04.26 メインの方へも反映させる修正 -chg sta
                            // 成形
                            // .Vari_Naiyo = tmp_tokuyaku.Remove(0, LINE_BREAK.Length)
                            if (!string.IsNullOrEmpty(tmp_tokuyaku))
                            {
                                model_cvitem.Vari_Naiyo = tmp_tokuyaku.Remove(0, CommonModule.LINE_BREAK.Length);
                            }
                            // 2016.04.26 メインの方へも反映させる修正 -chg end

                            // 固定値
                            model_cvitem.Vari_History = CommonModule.DefHistory;

                            // 特約にデータが存在する場合に書込処理を行う
                            if (!string.IsNullOrEmpty(model_cvitem.Vari_Naiyo))
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
                                    // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                                    // '2016.04.26 メインの方へも反映させる修正 -chg sta
                                    // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                                    // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                                    // '2016.04.26 メインの方へも反映させる修正 -chg end
                                    hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                                    // 20160819 部屋の半角変換処理によるエラー修正 -chg end

                                    // 挿入処理
                                    CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                                    // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                    if (normalflg)
                                    {
                                        if (list_chkduplicate.Contains(tmp_keydup) == false)
                                        {
                                            list_chkduplicate.Add(tmp_keydup);
                                        }
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

                        else
                        {

                            // 親データ無し、重複チェックエラーログ出力
                            // ログ出力メッセージ整形
                            tmp_hash.Clear();
                            LogSetting.Set_Log_Value_KomkErr(hash_keylog, tmp_hash, tmp_hash, log_keyout, tblname, ref tmp_logcnt, ref sortlist_log);  // ※引数のtmp_hashは使用しないが仮で入れておく

                            // ログ出力
                            if (sortlist_log.Count >= CommonModule.Log_OutputCnt | sortlist_log.Count != 0 & cntii == totalrowcnt - 1)
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
                    if (tmp_cvrowcnt > 0)
                    {
                        totalrowcnt = totalrowcnt + 1;
                    }
                    if (cvitemcnt > 0)
                    {
                        tmp_cvcnt = tmp_cvcnt + 1;
                    }
                    // 20160928 メモ関連の移行件数表示修正 -chg end
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
                midrowcnt = cvtaisyocnt;

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,tokuyaku_grpno) FROM " + tblname;
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

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

    #region 契約メモ情報

    public class Kydata_memo_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_memo_Model();        // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";
                int totalrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_memo";
                string fldnamegrp = "ky_guid,ky_recno,memo_no,memo,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                string fldname_keysub1 = readtbl.Columns[keycol_sub1 - 1].ColumnName.Trim();
                string fldname_keysub2 = readtbl.Columns[keycol_sub2 - 1].ColumnName.Trim();
                string fldname_keysub3 = readtbl.Columns[keycol_sub3 - 1].ColumnName.Trim();

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

                    // 移行判別用
                    bool cvflg = true;
                    bool keychkflg = true;

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string tmp_keysub3 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub3 - 1]).Trim();
                    string tmp_keyoya = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    string tmp_keydup = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    model_cvitem.Vari_Ky_recno = tmp_keysub3;

                    // 備考カウント初期化
                    int cvitemcnt = 0;
                    int tmp_cvrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                    // データ有無チェック
                    string tmp_fldvalueumuchk = "";
                    for (int cntjj = 4, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                        tmp_fldvalueumuchk = tmp_fldvalueumuchk + Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();
                    if (string.IsNullOrEmpty(tmp_fldvalueumuchk))
                    {
                        cvflg = false;
                    }

                    if (cvflg)
                    {

                        // 移行対象件数カウント
                        cvtaisyocnt = cvtaisyocnt + 1;

                        var hash_keylog = new SafeDictionary<string, string>();
                        string log_keyout = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1;
                        // 親データ有無チェック
                        if (list_basekeydata.Contains(tmp_keyoya) == false)
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "kydata_memo-ky_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keydup))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "kydata_memo-ky_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        if (keychkflg)
                        {

                            // 移行値取得
                            for (int cntjj = 4, loopTo2 = readtbl.Columns.Count - 1; cntjj <= loopTo2; cntjj++)
                            {

                                // サブキー取得
                                string tmp_keysub4 = (cntjj - 3).ToString();

                                // 全キー取得
                                string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4;

                                // ログ出力用データ格納(サブキーフィールド)
                                string fldname_keysub4 = readtbl.Columns[cntjj].ColumnName.Trim();
                                string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



                                // 登録値取得
                                string fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();
                                model_cvitem.Vari_Memo = fldvalue;

                                // 固定値
                                model_cvitem.Vari_Memo_no = (cntjj - 3).ToString();
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
                                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

                                        // 挿入処理
                                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                        if (normalflg)
                                        {
                                            if (list_chkduplicate.Contains(tmp_keydup) == false)
                                            {
                                                list_chkduplicate.Add(tmp_keydup);
                                            }
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
                        }

                        else
                        {

                            // 親データ無し、重複チェックエラーログ出力
                            // ログ出力メッセージ整形
                            tmp_hash.Clear();
                            LogSetting.Set_Log_Value_KomkErr(hash_keylog, tmp_hash, tmp_hash, log_keyout, tblname, ref tmp_logcnt, ref sortlist_log);  // ※引数のtmp_hashは使用しないが仮で入れておく

                            // ログ出力
                            if (sortlist_log.Count >= CommonModule.Log_OutputCnt | sortlist_log.Count != 0 & cntii == totalrowcnt - 1)
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
                    if (tmp_cvrowcnt > 0)
                    {
                        totalrowcnt = totalrowcnt + 1;
                    }
                    if (cvitemcnt > 0)
                    {
                        tmp_cvcnt = tmp_cvcnt + 1;
                    }
                    // 20160928 メモ関連の移行件数表示修正 -chg end
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
                midrowcnt = cvtaisyocnt;

                // 移行件数を取得
                cvrowcnt = tmp_cvcnt;

                // 調整件数を取得
                conditioncnt = tmp_condcnt;

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,memo_no) FROM " + tblname;
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

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

    #region 契約入金項目情報

    public class Kydata_nkin_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_nkin_Model();        // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                int keycol_sub5 = 6;                                  // サブ5キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 2016.04.06 入金項目読込処理の修正 -chg sta
                // Call Me.Get_RelData_NkinKomk()
                // Call Me.Get_RelData_NkinKbn()
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();
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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_nkin";
                string fldnamegrp = "ky_guid,ky_recno,tuki_kbn,nkin_no,nkin_recno," + "nkin_sortorder,nkin_kbn,sq_gak,sq_zeikbn,sq_zeigak," + "calc_kbn,calc_nkinno,calc_monthcnt,sqsaki_no,nkbn_yotei," + "sq_mmkbn,frstart_ymd,frend_ymd,frsq_gak,sqstart_ymd," + "sq_ptn,sq_interval,sq_nen,sq_tuki,zei_rit," + "biko,nkin_guid,history,fr_kbn";





                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                string fldname_keysub5 = readtbl.Columns[keycol_sub5 - 1].ColumnName.Trim();

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
                    string tmp_nkinname = "";         // 2016.04.06 入金項目読込処理の修正 -add
                    string tmp_nkinname_kijyun = "";  // 20160530 ログ修正 -add
                    string tmp_sqkankaku = "";        // 2016.04.26 メインの方へも反映させる修正 -add

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string tmp_keysub3 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub3 - 1]).Trim();
                    string tmp_keysub4 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub4 - 1]).Trim();
                    string tmp_keysub5 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub5 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4 + "-" + tmp_keysub5;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4 + "、" + fldname_keysub5 + " = " + tmp_keysub5;





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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue;
                                    break;
                                }
                            case "月区分":
                                {
                                    model_cvitem.Vari_Tuki_kbn = fldvalue;
                                    break;
                                }
                            case "入金項目名":
                                {
                                    // 2016.04.06 入金項目読込処理の修正 -chg sta
                                    // .Vari_Nkin_no = fldvalue
                                    tmp_nkinname = fldvalue;
                                    break;
                                }
                            // 2016.04.06 入金項目読込処理の修正 -chg end
                            case "入金項目レコードNo":
                                {
                                    model_cvitem.Vari_Nkin_recno = fldvalue;
                                    break;
                                }
                            case "並び順No":
                                {
                                    model_cvitem.Vari_Nkin_sortorder = fldvalue;
                                    break;
                                }
                            case "入金項目区分":
                                {
                                    model_cvitem.Vari_Nkin_kbn = fldvalue;
                                    break;
                                }
                            case "請求額":
                                {
                                    model_cvitem.Vari_Sq_gak = fldvalue;
                                    break;
                                }
                            case "税区分":
                                {
                                    model_cvitem.Vari_Sq_zeikbn = fldvalue;
                                    break;
                                }
                            case "請求税額":
                                {
                                    model_cvitem.Vari_Sq_zeigak = fldvalue;
                                    break;
                                }
                            case "算出区分":
                                {
                                    model_cvitem.Vari_Calc_kbn = fldvalue;
                                    break;
                                }
                            case "算出基準入金項目名":
                                {
                                    // 20160530 ログ修正 -chg sta
                                    // .Vari_Calc_nkinno = fldvalue
                                    tmp_nkinname_kijyun = fldvalue;
                                    break;
                                }
                            // 20160530 ログ修正 -chg end
                            case "算出ヶ月":
                                {
                                    model_cvitem.Vari_Calc_monthcnt = fldvalue;
                                    break;
                                }
                            case "請求先No":
                                {
                                    model_cvitem.Vari_Sqsaki_no = fldvalue;
                                    break;
                                }
                            case "入金方法":
                                {
                                    model_cvitem.Vari_Nkbn_yotei = fldvalue;
                                    break;
                                }
                            case "請求月区分":
                                {
                                    model_cvitem.Vari_Sq_mmkbn = fldvalue;
                                    break;
                                }
                            case "フリーレント適用開始日":
                                {
                                    model_cvitem.Vari_Frstart_ymd = fldvalue;
                                    break;
                                }
                            case "フリーレント適用終了日":
                                {
                                    model_cvitem.Vari_Frend_ymd = fldvalue;
                                    break;
                                }
                            case "フリーレント終了月請求額":
                                {
                                    model_cvitem.Vari_Frsq_gak = fldvalue;
                                    break;
                                }
                            case "請求開始月":
                                {
                                    model_cvitem.Vari_Sqstart_ymd = fldvalue;
                                    break;
                                }
                            // 2016.04.26 メインの方へも反映させる修正 -del sta
                            // 必要な情報をまとめて格納するためここではコメントアウト
                            // Case "請求パターン"
                            // .Vari_Sq_ptn = fldvalue
                            // 2016.04.26 メインの方へも反映させる修正 -del end
                            case "請求間隔":
                                {
                                    // 2016.04.26 メインの方へも反映させる修正 -chg sta
                                    // 請求間隔は変換して移行するため作業用変数へ格納しておく
                                    // .Vari_Sq_interval = fldvalue
                                    tmp_sqkankaku = fldvalue;
                                    break;
                                }
                            // 2016.04.26 メインの方へも反映させる修正 -chg end

                            // 2016.04.26 メインの方へも反映させる修正 -del sta
                            // 必要な情報をまとめて格納するためここではコメントアウト
                            // Case "請求発生年"
                            // .Vari_Sq_nen = fldvalue
                            // Case "請求発生月"
                            // .Vari_Sq_tuki = fldvalue
                            // 2016.04.26 メインの方へも反映させる修正 -del end
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
                            case "フリーレント適用区分":
                                {
                                    model_cvitem.Vari_Fr_kbn = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 2016.04.26 メインの方へも反映させる修正 -chg sta
                    // '2016.04.06 入金項目読込処理の修正 -add sta
                    // '紐付用に成形
                    // .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname(.Vari_Nkin_kbn)
                    // '2016.04.06 入金項目読込処理の修正 -add end

                    // 紐付用に成形
                    if (model_cvitem.Vari_Tuki_kbn == "2" & model_cvitem.Vari_Nkin_kbn == "4")       // 解約時の通常月入金項目は通常月の入金項目へ紐付ける
                    {
                        string argtukikbn = "1";
                        model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn);
                    }
                    else
                    {
                        string localGet_Nkinruiname() { string argtukikbn = model_cvitem.Vari_Nkin_kbn; var ret = EtcMethod.Get_Nkinruiname(ref argtukikbn); model_cvitem.Vari_Nkin_kbn = argtukikbn; return ret; }

                        model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + localGet_Nkinruiname();
                    }
                    // 2016.04.26 メインの方へも反映させる修正 -chg end

                    // 20160530 ログ修正 -add sta
                    if (!string.IsNullOrEmpty(tmp_nkinname_kijyun))
                    {
                        string argtukikbn1 = "1";
                        model_cvitem.Vari_Calc_nkinno = tmp_nkinname_kijyun + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn1);
                    }
                    else
                    {
                        model_cvitem.Vari_Calc_nkinno = "";
                    }
                    // 20160530 ログ修正 -add end

                    // 2016.04.26 メインの方へも反映させる修正 -add sta
                    // 変動費で固定項目の請求間隔、請求発生年、発生月を取得するために必要なデータを格納しておく
                    // 必要な情報
                    // 移行項目(部屋と契約入金項目情報で移行内容が異なるため)、請求パターン、請求月、請求開始月、請求間隔
                    model_cvitem.Vari_Sq_ptn = sheetname + "-" + model_cvitem.Vari_Sq_mmkbn + "-" + model_cvitem.Vari_Sqstart_ymd + "-" + tmp_sqkankaku;
                    model_cvitem.Vari_Sq_interval = sheetname + "-" + model_cvitem.Vari_Sq_mmkbn + "-" + model_cvitem.Vari_Sqstart_ymd + "-" + tmp_sqkankaku;
                    model_cvitem.Vari_Sq_nen = sheetname + "-" + model_cvitem.Vari_Sq_mmkbn + "-" + model_cvitem.Vari_Sqstart_ymd + "-" + tmp_sqkankaku;
                    model_cvitem.Vari_Sq_tuki = sheetname + "-" + model_cvitem.Vari_Sq_mmkbn + "-" + model_cvitem.Vari_Sqstart_ymd + "-" + tmp_sqkankaku;
                    // 2016.04.26 メインの方へも反映させる修正 -add end

                    // 固定値
                    model_cvitem.Vari_Nkin_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // 20160617 マイナス金額移行処理修正 -add sta
                    // マイナス金額を予めチェックしておく
                    Chk_MinusSqgak(ref tmp_hash);
                    // 20160617 マイナス金額移行処理修正 -add end

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                // tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            // 2016.04.06 入金項目読込処理の修正 -del sta
            // ''' <summary>
            // ''' 入金項目紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKomk()

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

            // ''' <summary>
            // ''' 入金区分紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKbn()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "入金区分マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // '既存データ有無確認(入金項目は他の項目でも参照するため)
            // If Hash_Rel_Nkinkbn.Count <> 0 Then
            // Exit Sub
            // End If

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldnkinkbnname As String = ""
            // Dim tmp_newnkinkbnno As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元入金区分名称"
            // tmp_oldnkinkbnname = fldvalue.Trim
            // Case "賃貸革命10入金区分No"
            // tmp_newnkinkbnno = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newnkinkbnno <> "" Then
            // Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 入金項目読込処理の修正 -del end

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
            /// マイナス請求の補正 '20160617 マイナス金額移行処理修正 -add
            /// 革命10にはマイナス金額を設定することができないためV7のマイナス金額を補正する
            /// 「敷金差額」「保証金差額」「敷金戻し」「保証金戻し」の属性に紐付いていることが条件
            /// </summary>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Chk_MinusSqgak(ref SafeDictionary<string, string> hash)
            {

                string chg_sqgak = "";
                string tmp_sqgak = Conversions.ToString(hash["sq_gak"]);
                string tmp_nkinname = Conversions.ToString(hash["nkin_no"]);
                string tmp_nkinzkseino = "";
                // 20161012 敷金、保証金差額のマイナス値移行対応 -chg sta
                // Dim list_taisyozkseino As New List(Of String) From {"400", "420"}       '「敷金戻し」「保証金戻し」の属性No
                var list_taisyozkseino = new List<string>() { "320", "330", "400", "420" };       // 「敷金差額」「保証金差額」「敷金戻し」「保証金戻し」の属性No
                // 20161012 敷金、保証金差額のマイナス値移行対応 -chg end
                // マイナスの値か判別→それ以外は処理を抜ける
                double tmp_dbl = 0d;
                if (double.TryParse(tmp_sqgak, out tmp_dbl) == false)
                {
                    return;
                }
                else if (tmp_dbl >= 0d)
                {
                    return;
                }

                // 入金項目属性Noとの照合→金額補正
                if (CommonModule.Hash_Rel_NkinkomkZksei.ContainsKey(tmp_nkinname))
                {
                    tmp_nkinzkseino = Conversions.ToString(CommonModule.Hash_Rel_NkinkomkZksei[tmp_nkinname]);

                    if (list_taisyozkseino.Contains(tmp_nkinzkseino))
                    {

                        // 「敷金戻し」「保証金戻し」の属性Noに該当する場合は絶対値を取得して再格納(マイナスをプラスへ変更)
                        chg_sqgak = Math.Abs(tmp_dbl).ToString();
                        hash["sq_gak"] = chg_sqgak.ToString();

                    }

                }

            }

        }

    }

    #endregion

    #region 契約次回入金項目情報

    public class Kydata_nkin_nx_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_nkin_nx_Model();     // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                int keycol_sub5 = 6;                                  // サブ5キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 2016.04.06 入金項目読込処理の修正 -chg sta
                // Call Me.Get_RelData_NkinKomk()
                // Call Me.Get_RelData_NkinKbn()
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();
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
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_nkin_nx";
                string fldnamegrp = "ky_guid,ky_recno,tuki_kbn,nkin_no,nkin_recno," + "nkin_sortorder,nkin_kbn,sq_gak,sq_zeikbn,sq_zeigak," + "calc_kbn,calc_nkinno,calc_monthcnt,sqsaki_no,nkbn_yotei," + "sq_mmkbn,sqstart_ymd,sq_ptn,sq_interval,sq_nen," + "sq_tuki,zei_rit,biko,history";




                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                string fldname_keysub5 = readtbl.Columns[keycol_sub5 - 1].ColumnName.Trim();

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
                    string tmp_nkinname = "";         // 2016.04.06 入金項目読込処理の修正 -add
                    string tmp_nkinname_kijyun = "";  // 20160530 ログ修正 -add

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string tmp_keysub3 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub3 - 1]).Trim();
                    string tmp_keysub4 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub4 - 1]).Trim();
                    string tmp_keysub5 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub5 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4 + "-" + tmp_keysub5;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4 + "、" + fldname_keysub5 + " = " + tmp_keysub5;





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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue;
                                    break;
                                }
                            case "月区分":
                                {
                                    model_cvitem.Vari_Tuki_kbn = fldvalue;
                                    break;
                                }
                            case "入金項目名":
                                {
                                    // 2016.04.06 入金項目読込処理の修正 -chg sta
                                    // .Vari_Nkin_no = fldvalue
                                    tmp_nkinname = fldvalue;
                                    break;
                                }
                            // 2016.04.06 入金項目読込処理の修正 -chg end
                            case "入金項目レコードNo":
                                {
                                    model_cvitem.Vari_Nkin_recno = fldvalue;
                                    break;
                                }
                            case "並び順No":
                                {
                                    model_cvitem.Vari_Nkin_sortorder = fldvalue;
                                    break;
                                }
                            case "入金項目区分":
                                {
                                    model_cvitem.Vari_Nkin_kbn = fldvalue;
                                    break;
                                }
                            case "請求額":
                                {
                                    model_cvitem.Vari_Sq_gak = fldvalue;
                                    break;
                                }
                            case "税区分":
                                {
                                    model_cvitem.Vari_Sq_zeikbn = fldvalue;
                                    break;
                                }
                            case "請求税額":
                                {
                                    model_cvitem.Vari_Sq_zeigak = fldvalue;
                                    break;
                                }
                            case "算出区分":
                                {
                                    model_cvitem.Vari_Calc_kbn = fldvalue;
                                    break;
                                }
                            case "算出基準入金項目名":
                                {
                                    // 20160530 ログ修正 -chg sta
                                    // .Vari_Calc_nkinno = fldvalue
                                    tmp_nkinname_kijyun = fldvalue;
                                    break;
                                }
                            // 20160530 ログ修正 -chg end
                            case "算出ヶ月":
                                {
                                    model_cvitem.Vari_Calc_monthcnt = fldvalue;
                                    break;
                                }
                            case "請求先No":
                                {
                                    model_cvitem.Vari_Sqsaki_no = fldvalue;
                                    break;
                                }
                            case "入金方法":
                                {
                                    model_cvitem.Vari_Nkbn_yotei = fldvalue;
                                    break;
                                }
                            case "請求月区分":
                                {
                                    model_cvitem.Vari_Sq_mmkbn = fldvalue;
                                    break;
                                }
                            case "請求開始月":
                                {
                                    model_cvitem.Vari_Sqstart_ymd = fldvalue;
                                    break;
                                }
                            case "請求パターン":
                                {
                                    model_cvitem.Vari_Sq_ptn = fldvalue;
                                    break;
                                }
                            case "請求発生間隔":
                                {
                                    model_cvitem.Vari_Sq_interval = fldvalue;
                                    break;
                                }
                            case "請求対象年":
                                {
                                    model_cvitem.Vari_Sq_nen = fldvalue;
                                    break;
                                }
                            case "請求対象月":
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
                        }

                    }

                    // 2016.04.06 入金項目読込処理の修正 -add sta
                    // 紐付用に成形
                    string localGet_Nkinruiname() { string argtukikbn1 = model_cvitem.Vari_Nkin_kbn; var ret = EtcMethod.Get_Nkinruiname(ref argtukikbn1); model_cvitem.Vari_Nkin_kbn = argtukikbn1; return ret; }

                    model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + localGet_Nkinruiname();
                    // 2016.04.06 入金項目読込処理の修正 -add end

                    // 20160530 ログ修正 -add sta
                    if (!string.IsNullOrEmpty(tmp_nkinname_kijyun))
                    {
                        string argtukikbn = "1";
                        model_cvitem.Vari_Calc_nkinno = tmp_nkinname_kijyun + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn);
                    }
                    else
                    {
                        model_cvitem.Vari_Calc_nkinno = "";
                    }
                    // 20160530 ログ修正 -add end

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // 20161012 敷金、保証金差額のマイナス値移行対応 -add sta
                    // マイナス金額を予めチェックしておく
                    Chk_MinusSqgak(ref tmp_hash);
                    // 20161012 敷金、保証金差額のマイナス値移行対応 -add end

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 作業用VIEWの削除
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);
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
                // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                // tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            // 2016.04.06 入金項目読込処理の修正 -del sta
            // ''' <summary>
            // ''' 入金項目紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKomk()

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

            // ''' <summary>
            // ''' 入金区分紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKbn()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "入金区分マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // '既存データ有無確認(入金項目は他の項目でも参照するため)
            // If Hash_Rel_Nkinkbn.Count <> 0 Then
            // Exit Sub
            // End If

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldnkinkbnname As String = ""
            // Dim tmp_newnkinkbnno As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元入金区分名称"
            // tmp_oldnkinkbnname = fldvalue.Trim
            // Case "賃貸革命10入金区分No"
            // tmp_newnkinkbnno = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newnkinkbnno <> "" Then
            // Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 入金項目読込処理の修正 -del end

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
            /// マイナス請求の補正 '20161012 敷金、保証金差額のマイナス値移行対応 -add
            /// 革命10にはマイナス金額を設定することができないためV7のマイナス金額を補正する
            /// 「敷金差額」「保証金差額」「敷金戻し」「保証金戻し」の属性に紐付いていることが条件
            /// </summary>
            /// <param name="hash"></param>
            /// <remarks></remarks>
            public void Chk_MinusSqgak(ref SafeDictionary<string, string> hash)
            {

                string chg_sqgak = "";
                string tmp_sqgak = Conversions.ToString(hash["sq_gak"]);
                string tmp_nkinname = Conversions.ToString(hash["nkin_no"]);
                string tmp_nkinzkseino = "";
                var list_taisyozkseino = new List<string>() { "320", "330", "400", "420" };       // 「敷金差額」「保証金差額」「敷金戻し」「保証金戻し」の属性No

                // マイナスの値か判別→それ以外は処理を抜ける
                double tmp_dbl = 0d;
                if (double.TryParse(tmp_sqgak, out tmp_dbl) == false)
                {
                    return;
                }
                else if (tmp_dbl >= 0d)
                {
                    return;
                }

                // 入金項目属性Noとの照合→金額補正
                if (CommonModule.Hash_Rel_NkinkomkZksei.ContainsKey(tmp_nkinname))
                {
                    tmp_nkinzkseino = Conversions.ToString(CommonModule.Hash_Rel_NkinkomkZksei[tmp_nkinname]);

                    if (list_taisyozkseino.Contains(tmp_nkinzkseino))
                    {

                        // 「敷金戻し」「保証金戻し」の属性Noに該当する場合は絶対値を取得して再格納(マイナスをプラスへ変更)
                        chg_sqgak = Math.Abs(tmp_dbl).ToString();
                        hash["sq_gak"] = chg_sqgak.ToString();

                    }

                }

            }

        }

    }

    #endregion

    #region 契約変動費各戸メーター情報

    public class Kydata_hendo_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル
                // 20160603 ユーザーデータ検証による修正 -del
                // Dim hash_bkhyhendoguid As New SafeDictionary<string, string>                            'キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_hendo_Model();       // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";
                // 20160603 ユーザーデータ検証による修正 -del
                // Dim viewnamebkhytohyhendoguid As String = PRE_VIEW_NAME & "物件部屋変動費キー情報"

                // ************************
                // 作業準備
                // ************************



                // 入金項目の重複チェックは特殊なので考慮すること

                // 紐付けデータ取得
                // 2016.04.06 入金項目読込処理の修正 -chg sta
                // Call Me.Get_RelData_NkinKomk()
                // Call Me.Get_RelData_NkinKbn()
                // Call Me.Get_RelData_Hendometer()
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();
                // 2016.04.06 入金項目読込処理の修正 -chg end

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);
                // 20160603 ユーザーデータ検証による修正 -del
                // Call Me.Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhytohyhendoguid)

                // テーブル名/フィールド名セット
                string tblname = "kydata_hendo";
                string fldnamegrp = "ky_guid,ky_recno,hyhendo_guid,hendo_sortorder,useflg," + "hendo_kbn,meter_name,nkin_no,hendorule_no,hendorule_biko," + "sqsaki_no,nkbn_yotei,biko,history,sq_mmkbn";


                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);
                // 20160603 ユーザーデータ検証による修正 -del
                // Call Me.Get_Guid(sqlcnnv10, viewnamebkhytohyhendoguid, hash_bkhyhendoguid)

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);
                string fldname_keysub3 = Conversions.ToString(headervalue[startrow - 1, keycol_sub3]);
                string fldname_keysub4 = Conversions.ToString(headervalue[startrow - 1, keycol_sub4]);

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
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "変動費No":
                                {
                                    model_cvitem.Vari_Hendo_sortorder = fldvalue.Trim();
                                    break;
                                }
                            case "請求対象フラグ":
                                {
                                    model_cvitem.Vari_Useflg = fldvalue.Trim();
                                    break;
                                }
                            case "メーター分類":
                                {
                                    break;
                                }
                            // .Vari_Hendo_kbn = fldvalue.Trim
                            case "メーター名":
                                {
                                    model_cvitem.Vari_Meter_name = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目名":
                                {
                                    // 2016.04.06 入金項目読込処理の修正 -chg sta
                                    // .Vari_Nkin_no = fldvalue.Trim
                                    // .Vari_Hendo_kbn = fldvalue.Trim     '紐付けたメーター分類を取得するため入金項目名を格納しておく
                                    string argtukikbn = "5";
                                    model_cvitem.Vari_Nkin_no = fldvalue + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn);
                                    string argtukikbn1 = "5";
                                    model_cvitem.Vari_Hendo_kbn = fldvalue + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn1);     // 紐付けたメーター分類を取得するため入金項目名を格納しておく
                                    break;
                                }
                            // 2016.04.06 入金項目読込処理の修正 -chg end
                            case "変動費請求ルールNo":
                                {
                                    model_cvitem.Vari_Hendorule_no = fldvalue.Trim();
                                    break;
                                }
                            case "変動費請求ルール備考":
                                {
                                    model_cvitem.Vari_Hendorule_biko = fldvalue.Trim();
                                    break;
                                }
                            case "請求先No":
                                {
                                    model_cvitem.Vari_Sqsaki_no = fldvalue.Trim();
                                    break;
                                }
                            case "入金方法":
                                {
                                    model_cvitem.Vari_Nkbn_yotei = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "請求月":
                                {
                                    model_cvitem.Vari_Sq_mmkbn = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Hyhendo_guid = Guid.NewGuid().ToString();
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

                        // 20160603 ユーザーデータ検証による修正 -del sta
                        // '変動区分を取得
                        // Dim hendo_guidkey As String = tmp_keymain & "-" & tmp_keysub1 & "-" & hash_cvitem.Item("hendo_kbn")
                        // 20160603 ユーザーデータ検証による修正 -del end

                        // キーをguidへ変換
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // 'hash_cvitem("hyhendo_guid") = hash_bkhyhendoguid.Item(hendo_guidkey)        '20160603 ユーザーデータ検証による修正 -del
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

                        // 20160603 ユーザーデータ検証による修正 -add sta
                        string kyguid = Conversions.ToString(hash_cvitem["ky_guid"]);
                        string hendokbn = Conversions.ToString(hash_cvitem["hendo_kbn"]);
                        string nkinno = Conversions.ToString(hash_cvitem["nkin_no"]);
                        string tmp_hyhendoguid = Get_HyHendoguid(sqlcnnv10, kyguid, hendokbn, nkinno);
                        hash_cvitem["hyhendo_guid"] = tmp_hyhendoguid;
                        // 20160603 ユーザーデータ検証による修正 -add end

                        // 挿入処理
                        // 20160819 部屋変動費guidが存在しない場合の処理を追加 -chg sta
                        // Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)
                        if (string.IsNullOrEmpty(tmp_hyhendoguid))
                        {
                            normalflg = false;
                            // ログ出力
                            string log_key = tblname + "-" + "hyhendo_guid";
                            string errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED;
                            hash_log.Clear();
                            hash_log.Add(log_key, errstr);
                        }
                        else
                        {
                            CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);
                        }
                        // 20160819 部屋変動費guidが存在しない場合の処理を追加 -chg end

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
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 2016.04.06 物件情報の変動費検針有無の一括更新処理追加 -add sta
                // 変動費が存在する物件の検針業務有無フラグを一括更新
                int tmpcnt = 0;
                string kensinumuupdatesql = Get_UseQry_Update();
                DBExec.Exec_NonQuery(sqlcnnv10, kensinumuupdatesql, ref tmpcnt);
                // 2016.04.06 物件情報の変動費検針有無の一括更新処理追加 -add end

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
                // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                // tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

            // 2016.04.06 入金項目読込処理の修正 -del sta
            // ''' <summary>
            // ''' 入金項目紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKomk()

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

            // ''' <summary>
            // ''' 入金区分紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKbn()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "入金区分マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // '既存データ有無確認(入金項目は他の項目でも参照するため)
            // If Hash_Rel_Nkinkbn.Count <> 0 Then
            // Exit Sub
            // End If

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldnkinkbnname As String = ""
            // Dim tmp_newnkinkbnno As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元入金区分名称"
            // tmp_oldnkinkbnname = fldvalue.Trim
            // Case "賃貸革命10入金区分No"
            // tmp_newnkinkbnno = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newnkinkbnno <> "" Then
            // Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub

            // ''' <summary>
            // ''' 各戸メーター分類紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_Hendometer()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "各戸メーター分類"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // '既存データ有無確認(入金項目は他の項目でも参照するため)
            // If Hash_Rel_Hendometer.Count <> 0 Then
            // Exit Sub
            // End If

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldnkinname As String = ""
            // Dim tmp_newmeterno As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元入金項目名称"
            // tmp_oldnkinname = fldvalue.Trim
            // Case "賃貸革命10メーター分類No"
            // tmp_newmeterno = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newmeterno <> "" Then
            // Hash_Rel_Hendometer.Add(tmp_oldnkinname, tmp_newmeterno)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 入金項目読込処理の修正 -del end

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
            /// [bkdata_detail].[kensingyomu_umu]の一括更新クエリ  2016.04.06 物件情報の変動費検針有無の一括更新処理追加
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE bkdata_detail SET ";
                tmp_sql = tmp_sql + " 	kensingyomu_umu = 1 ";
                tmp_sql = tmp_sql + " WHERE bk_guid IN ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT bk_guid FROM kydata AS KY ";
                tmp_sql = tmp_sql + " 	WHERE ky_guid IN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ky_guid FROM kydata_hendo ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " ); ";

                return tmp_sql;

            }

            /// <summary>
            /// 部屋変動費情報から変動費guidを取得 '20160603 ユーザーデータ検証による修正 新規追加
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hyguid"></param>
            /// <param name="hendokbn"></param>
            /// <param name="nkinno"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_HyHendoguid(SqlConnection sqlcnnv10, string kyguid, string hendokbn, string nkinno)
            {

                string rtn_guid = "";
                string tmp_sql = "";

                // 20160616 契約変動費情報の抽出条件不足時の処理の追加 -add sta
                // 抽出条件が不足している場合は処理を抜ける
                if (string.IsNullOrEmpty(kyguid.Trim()) | string.IsNullOrEmpty(hendokbn.Trim()) | string.IsNullOrEmpty(nkinno.Trim()))
                {
                    return rtn_guid;
                }
                // 20160616 契約変動費情報の抽出条件不足時の処理の追加 -add end

                tmp_sql = tmp_sql + " SELECT CONVERT(VARCHAR(50),hyhendo_guid) FROM hydata_hendo ";
                tmp_sql = tmp_sql + " WHERE hy_guid =  ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT hy_guid FROM kydata WHERE ky_guid = '" + kyguid + "' ";
                tmp_sql = tmp_sql + " 	) ";
                tmp_sql = tmp_sql + " AND   hendo_kbn = " + hendokbn;
                tmp_sql = tmp_sql + " AND   nkin_no = " + nkinno;
                rtn_guid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));

                return rtn_guid;

            }

        }

    }

    #endregion

    #region 契約控除ルール情報

    public class Kydata_kojorule_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_kojorule_Model();    // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                int keycol_sub5 = 6;                                  // サブ5キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************

                // 入金項目の重複チェックは特殊なので考慮すること

                // 紐付けデータ取得
                // 2016.04.06 入金項目読込処理の修正 -chg sta
                // Call Me.Get_RelData_NkinKomk()
                // Call Me.Get_RelData_NkinKbn()
                SetRelItemToObject.Set_RelData_Nkinkomk();
                SetRelItemToObject.Set_RelData_Nkinkbn();
                // 2016.04.06 入金項目読込処理の修正 -chg end

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_kojorule";
                string fldnamegrp = "ky_guid,ky_recno,taisyokbn,nkin_no,nkin_sortorder," + "sosai_flg,calc_kbn,kojo_gak,kojo_gakzeikbn,kojonkin_no," + "kojo_rit,kojo_ritzeikbn,kojo_ritutizei,zei_rit,tateazu_flg," + "kojo_zeigak,sotaisyo_flg,sotaisyonkin_no";



                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);
                string fldname_keysub3 = Conversions.ToString(headervalue[startrow - 1, keycol_sub3]);
                string fldname_keysub4 = Conversions.ToString(headervalue[startrow - 1, keycol_sub4]);
                string fldname_keysub5 = Conversions.ToString(headervalue[startrow - 1, keycol_sub5]);

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
                    string tmp_keysub5 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub5]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4 + "-" + tmp_keysub5;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4 + "、" + fldname_keysub5 + " = " + tmp_keysub5;





                    // 2016.04.06 入金項目読込処理の修正 -add
                    // 作業用変数
                    string tmp_nkinname = "";

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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "対象区分":
                                {
                                    model_cvitem.Vari_Taisyokbn = fldvalue.Trim();
                                    break;
                                }
                            case "控除入金項目名":
                                {
                                    // 2016.04.06 入金項目読込処理の修正 -chg sta
                                    // .Vari_Nkin_no = fldvalue.Trim
                                    tmp_nkinname = fldvalue;
                                    break;
                                }
                            // 2016.04.06 入金項目読込処理の修正 -chg end
                            case "表示順":
                                {
                                    model_cvitem.Vari_Nkin_sortorder = fldvalue.Trim();
                                    break;
                                }
                            case "相殺予定フラグ":
                                {
                                    model_cvitem.Vari_Sosai_flg = fldvalue.Trim();
                                    break;
                                }
                            case "控除額算出基準":
                                {
                                    model_cvitem.Vari_Calc_kbn = fldvalue.Trim();
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
                            case "控除率内税":
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
                    model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn);
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // 入金項目の重複チェックは特殊なので考慮すること

                // Dim tmp_sql As String = ""
                // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                // tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            // 2016.04.06 入金項目読込処理の修正 -del sta
            // ''' <summary>
            // ''' 入金項目紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKomk()

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

            // ''' <summary>
            // ''' 入金区分紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKbn()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "入金区分マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // '既存データ有無確認(入金項目は他の項目でも参照するため)
            // If Hash_Rel_Nkinkbn.Count <> 0 Then
            // Exit Sub
            // End If

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldnkinkbnname As String = ""
            // Dim tmp_newnkinkbnno As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元入金区分名称"
            // tmp_oldnkinkbnname = fldvalue.Trim
            // Case "賃貸革命10入金区分No"
            // tmp_newnkinkbnno = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newnkinkbnno <> "" Then
            // Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 入金項目読込処理の修正 -del end

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

    #region 契約送金ルール情報

    public class Kydata_sorule_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_sorule_Model();      // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                int keycol_sub5 = 6;                                  // サブ5キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************



                // 入金項目の重複チェックは特殊なので考慮すること


                // 紐付けデータ取得
                // 2016.04.06 入金項目読込処理の修正 -chg sta
                // Call Me.Get_RelData_NkinKomk()
                SetRelItemToObject.Set_RelData_Nkinkomk();
                // 2016.04.06 入金項目読込処理の修正 -chg end

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_sorule";
                string fldnamegrp = "ky_guid,ky_recno,taisyokbn,nkin_no,nkin_sortorder," + "sokin_rit,kanrigak_rit,hosyo_flg";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);
                string fldname_keysub3 = Conversions.ToString(headervalue[startrow - 1, keycol_sub3]);
                string fldname_keysub4 = Conversions.ToString(headervalue[startrow - 1, keycol_sub4]);
                string fldname_keysub5 = Conversions.ToString(headervalue[startrow - 1, keycol_sub5]);

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
                    string tmp_keysub5 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub5]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4 + "-" + tmp_keysub5;
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4 + "、" + fldname_keysub5 + " = " + tmp_keysub5;





                    // 2016.04.06 入金項目読込処理の修正 -add
                    // 作業用変数
                    string tmp_nkinname = "";

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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "対象区分":
                                {
                                    model_cvitem.Vari_Taisyokbn = fldvalue.Trim();
                                    break;
                                }
                            case "入金項目名":
                                {
                                    // 2016.04.06 入金項目読込処理の修正 -chg sta
                                    // .Vari_Nkin_no = fldvalue.Trim
                                    tmp_nkinname = fldvalue;
                                    break;
                                }
                            // 2016.04.06 入金項目読込処理の修正 -chg end
                            case "表示順":
                                {
                                    model_cvitem.Vari_Nkin_sortorder = fldvalue.Trim();
                                    break;
                                }
                            case "送金率":
                                {
                                    model_cvitem.Vari_Sokin_rit = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料率":
                                {
                                    model_cvitem.Vari_Kanrigak_rit = fldvalue.Trim();
                                    break;
                                }
                            case "滞納保証有無":
                                {
                                    model_cvitem.Vari_Hosyo_flg = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 2016.04.26 メインの方へも反映させる修正 -chg sta
                    // '2016.04.06 入金項目読込処理の修正 -add sta
                    // '紐付用に成形
                    // .Vari_Nkin_no = tmp_nkinname & "-" & EtcMethod.Get_Nkinruiname(.Vari_Taisyokbn)
                    // '2016.04.06 入金項目読込処理の修正 -add end

                    // 紐付用に成形
                    switch (model_cvitem.Vari_Taisyokbn ?? "")
                    {
                        case "4":
                            {
                                string argtukikbn = "1";
                                model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn);     // 解約時の通常月入金項目は通常月の入金項目へ紐付ける
                                break;
                            }
                        case "6":
                            {
                                string argtukikbn1 = "4";
                                model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + EtcMethod.Get_Nkinruiname(ref argtukikbn1);     // 解約費の入金項目は解約時の入金項目へ紐付ける
                                break;
                            }

                        default:
                            {
                                string localGet_Nkinruiname() { string argtukikbn2 = model_cvitem.Vari_Taisyokbn; var ret = EtcMethod.Get_Nkinruiname(ref argtukikbn2); model_cvitem.Vari_Taisyokbn = argtukikbn2; return ret; }

                                model_cvitem.Vari_Nkin_no = tmp_nkinname + "-" + localGet_Nkinruiname();
                                break;
                            }
                    }
                    // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

                // 契約送金ルール区分一括更新
                int tmpcnt = 0;
                string gaibunoupdatesql = Get_UseQry_Update();
                DBExec.Exec_NonQuery(sqlcnnv10, gaibunoupdatesql, ref tmpcnt);

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
                // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) + '-' + CONVERT(varchar,kys_cnt) FROM " & tblname
                // tmp_sql = tmp_sql & " LEFT JOIN kydata ON " & tblname & ".ky_guid = kydata.ky_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN hydata ON kydata.hy_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON kydata.bk_guid = bkdata.bk_guid "
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            // 2016.04.06 入金項目読込処理の修正 -del sta
            // ''' <summary>
            // ''' 入金項目紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_NkinKomk()

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
            /// [kydata_kihon].[sokojorule_kbn]の一括更新クエリ
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE kydata_kihon SET ";
                tmp_sql = tmp_sql + " 	sokojorule_kbn = 2 ";
                tmp_sql = tmp_sql + " ; ";

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 契約解約情報

    public class Kydata_kai_Repository
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
                // Dim hash_bkhyguid As New SafeDictionary<string, string>                                'bkguid格納用ハッシュテーブル
                var hash_bkhykyguid = new SafeDictionary<string, string>();                              // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_kai_Model();         // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                string viewnamebkhykytokyguid = CommonModule.PRE_VIEW_NAME + "物件部屋契約キー情報";

                // ************************
                // 作業準備
                // ************************

                // 20160629 修繕検証後修正 入金区分の紐付情報取得 -add
                SetRelItemToObject.Set_RelData_Nkinkbn();

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10, viewnamebkhykytokyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_kai";
                // 20160620 EXEUpdateに伴う修正2 -chg sta
                // '20160519 EXEUpdateに伴う修正 契約解約情報 -chg sta
                // 'Dim fldnamegrp As String = "ky_guid,ky_recno,kai_ymd,seisan_ymd,seisan_completeymd," & _
                // '                           "uketuke_ymd,riyu,tatiai_ymd,uketuke_logonuser_no,tatiai_logonuser_no," & _
                // '                           "seisan_logonuser_no,szen_logonuser_no,szen_jisyano,szen_endymd,szen_kojiyoteistartymd," & _
                // '                           "szen_kojiyoteiendymd,szen_kojibasyo,szen_kojigaiyo,next_name,next_namesjis," & _
                // '                           "next_postcode,next_addr1,next_addr2,next_tel1,yatin_kozano," & _
                // '                           "biko,biko_tatiai,seisan_henkinymd,seisan_sqymd,kanritesu_flg," & _
                // '                           "kanritesu_gak,rowid,history,so_yoteiymd,tatiai," & _
                // '                           "tatiai_yoteiymd,bosyu_jokenymd,ow_logonuser_no,bosyujoken,confirmkai_sekininsya," & _
                // '                           "confirmkai_ymd,confirmkai_print,szen_jisyafutanumuflg,szen_jisyafutangak,szen_jisyafutanzeikbn," & _
                // '                           "szen_jisyafutanzeigak,szen_bikojisyafutan,szen_sonotafutanumuflg,szen_sonotafutangak,szen_sonotafutanzeikbn," & _
                // '                           "szen_sonotafutanzeigak,szen_sonotafutanname,szen_bikosonotafutan,tatiai_yoteitime,next_keisyo"
                // Dim fldnamegrp As String = "ky_guid,ky_recno,kai_ymd,seisan_ymd,seisan_completeymd," & _
                // "uketuke_ymd,riyu,tatiai_ymd,uketuke_logonuser_no,tatiai_logonuser_no," & _
                // "seisan_logonuser_no,szen_logonuser_no,szen_jisyano,szen_endymd,szen_kojiyoteistartymd," & _
                // "szen_kojiyoteiendymd,szen_kojibasyo,szen_kojigaiyo,next_name,next_namesjis," & _
                // "next_postcode,next_addr1,next_addr2,next_tel1,yatin_kozano," & _
                // "biko,biko_tatiai,seisan_henkinymd,seisan_sqymd,kanritesu_flg," & _
                // "kanritesu_gak,rowid,history,so_yoteiymd,tatiai," & _
                // "tatiai_yoteiymd,bosyu_jokenymd,ow_logonuser_no,bosyujoken,szen_jisyafutanumuflg," & _
                // "szen_jisyafutangak,szen_jisyafutanzeikbn,szen_jisyafutanzeigak,szen_bikojisyafutan,szen_sonotafutanumuflg," & _
                // "szen_sonotafutangak,szen_sonotafutanzeikbn,szen_sonotafutanzeigak,szen_sonotafutanname,szen_bikosonotafutan," & _
                // "tatiai_yoteitime,next_keisyo,nkin_kaiyakutukinoprintflg"
                // '20160519 EXEUpdateに伴う修正 契約解約情報 -chg end
                // 20160829 革命10バージョンアップに伴う修正 szen_owkojoymd,szen_owsqymd を追加 -add
                string fldnamegrp = "ky_guid,ky_recno,kai_ymd,seisan_ymd,seisan_completeymd," + "uketuke_ymd,riyu,tatiai_ymd,uketuke_logonuser_no,tatiai_logonuser_no," + "seisan_logonuser_no,szen_logonuser_no,szen_jisyano,szen_endymd,szen_kojiyoteistartymd," + "szen_kojiyoteiendymd,szen_kojibasyo,szen_kojigaiyo,next_name,next_namesjis," + "next_postcode,next_addr1,next_addr2,next_tel1,yatin_kozano," + "biko,biko_tatiai,seisan_henkinymd,seisan_sqymd,kanritesu_flg," + "kanritesu_gak,rowid,history,so_yoteiymd,tatiai," + "tatiai_yoteiymd,bosyu_jokenymd,ow_logonuser_no,bosyujoken,tatiai_yoteitime," + "next_keisyo,nkin_kaiyakutukinoprintflg,szen_kysno,szen_kysnkbn,szen_kyssorit," + "szen_kyssogak,szen_owno,szen_owkaisyukbn,szen_owsosakisoruleguid,szen_owsosakisoruleno," + "szen_owsokozano,szen_owsqsimeymd,szen_ownkbn,szen_owyatinkozano,szen_owsqsakino," + "szen_owkojoymd,szen_owsqymd";










                // 20160620 EXEUpdateに伴う修正2 -chg end

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewnamebkhykytokyguid, ref hash_bkhykyguid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewnamebkhykytokyguid, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

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
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);
                string fldname_keysub3 = Conversions.ToString(headervalue[startrow - 1, keycol_sub3]);

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
                    model_cvitem.Vari_Ky_guid = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;



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
                            case "契約管理レコードNo":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue.Trim();
                                    break;
                                }
                            case "解約日":
                                {
                                    model_cvitem.Vari_Kai_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "解約精算費用決定日":
                                {
                                    model_cvitem.Vari_Seisan_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "解約精算業務完了日":
                                {
                                    model_cvitem.Vari_Seisan_completeymd = fldvalue.Trim();
                                    break;
                                }
                            case "解約受付日":
                                {
                                    model_cvitem.Vari_Uketuke_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "解約理由":
                                {
                                    model_cvitem.Vari_Riyu = fldvalue.Trim();
                                    break;
                                }
                            case "退去日":
                                {
                                    model_cvitem.Vari_Tatiai_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "受付担当者No":
                                {
                                    model_cvitem.Vari_Uketuke_logonuser_no = fldvalue.Trim();
                                    break;
                                }
                            case "立会担当者No":
                                {
                                    model_cvitem.Vari_Tatiai_logonuser_no = fldvalue.Trim();
                                    break;
                                }
                            case "精算担当者No":
                                {
                                    model_cvitem.Vari_Seisan_logonuser_no = fldvalue.Trim();
                                    break;
                                }
                            case "修繕担当者No":
                                {
                                    model_cvitem.Vari_Szen_logonuser_no = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店No":
                                {
                                    model_cvitem.Vari_Szen_jisyano = fldvalue.Trim();
                                    break;
                                }
                            case "修繕完了日":
                                {
                                    model_cvitem.Vari_Szen_endymd = fldvalue.Trim();
                                    break;
                                }
                            case "工事予定期間開始日":
                                {
                                    model_cvitem.Vari_Szen_kojiyoteistartymd = fldvalue.Trim();
                                    break;
                                }
                            case "工事予定期間終了日":
                                {
                                    model_cvitem.Vari_Szen_kojiyoteiendymd = fldvalue.Trim();
                                    break;
                                }
                            case "工事場所":
                                {
                                    model_cvitem.Vari_Szen_kojibasyo = fldvalue.Trim();
                                    break;
                                }
                            case "工事概要":
                                {
                                    model_cvitem.Vari_Szen_kojigaiyo = fldvalue.Trim();
                                    break;
                                }
                            case "退去後宛名":
                                {
                                    model_cvitem.Vari_Next_name = fldvalue.Trim();
                                    break;
                                }
                            case "退去後宛名Shift-jis":
                                {
                                    model_cvitem.Vari_Next_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "退去後郵便番号":
                                {
                                    model_cvitem.Vari_Next_postcode = fldvalue.Trim();
                                    break;
                                }
                            case "退去後住所①":
                                {
                                    model_cvitem.Vari_Next_addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "退去後住所②":
                                {
                                    model_cvitem.Vari_Next_addr2 = fldvalue.Trim();
                                    break;
                                }
                            case "退去後電話番号①":
                                {
                                    model_cvitem.Vari_Next_tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "不足時入金口座":
                                {
                                    model_cvitem.Vari_Yatin_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "備考(解約精算)":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "備考(立会)":
                                {
                                    model_cvitem.Vari_Biko_tatiai = fldvalue.Trim();
                                    break;
                                }
                            case "返金予定日":
                                {
                                    model_cvitem.Vari_Seisan_henkinymd = fldvalue.Trim();
                                    break;
                                }
                            case "請求締め日":
                                {
                                    model_cvitem.Vari_Seisan_sqymd = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料フラグ":
                                {
                                    model_cvitem.Vari_Kanritesu_flg = fldvalue.Trim();
                                    break;
                                }
                            case "管理手数料額":
                                {
                                    model_cvitem.Vari_Kanritesu_gak = fldvalue.Trim();
                                    break;
                                }
                            case "送金予定日":
                                {
                                    model_cvitem.Vari_So_yoteiymd = fldvalue.Trim();
                                    break;
                                }
                            case "立会い":
                                {
                                    model_cvitem.Vari_Tatiai = fldvalue.Trim();
                                    break;
                                }
                            case "退去予定日":
                                {
                                    model_cvitem.Vari_Tatiai_yoteiymd = fldvalue.Trim();
                                    break;
                                }
                            case "募集条件確認日":
                                {
                                    model_cvitem.Vari_Bosyu_jokenymd = fldvalue.Trim();
                                    break;
                                }
                            case "家主連絡担当者No":
                                {
                                    model_cvitem.Vari_Ow_logonuser_no = fldvalue.Trim();
                                    break;
                                }
                            case "募集条件内容":
                                {
                                    model_cvitem.Vari_Bosyujoken = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 契約解約情報 -del sta
                            // Case "確認事項入力責任者"
                            // .Vari_Confirmkai_sekininsya = fldvalue.Trim
                            // Case "確認事項確認日"
                            // .Vari_Confirmkai_ymd = fldvalue.Trim
                            // Case "印刷時の確認事項"
                            // .Vari_Confirmkai_print = fldvalue.Trim
                            // 20160519 EXEUpdateに伴う修正 契約解約情報 -del end
                            // 20160620 EXEUpdateに伴う修正2 -del sta
                            // Case "修繕自社負担フラグ"
                            // .Vari_Szen_jisyafutanumuflg = fldvalue.Trim
                            // Case "修繕自社負担額"
                            // .Vari_Szen_jisyafutangak = fldvalue.Trim
                            // Case "修繕自社負担税区分"
                            // .Vari_Szen_jisyafutanzeikbn = fldvalue.Trim
                            // Case "修繕自社負担税額"
                            // .Vari_Szen_jisyafutanzeigak = fldvalue.Trim
                            // Case "修繕自社負担備考"
                            // .Vari_Szen_bikojisyafutan = fldvalue.Trim
                            // Case "修繕その他負担フラグ"
                            // .Vari_Szen_sonotafutanumuflg = fldvalue.Trim
                            // Case "修繕その他負担額"
                            // .Vari_Szen_sonotafutangak = fldvalue.Trim
                            // Case "修繕その他負担税区分"
                            // .Vari_Szen_sonotafutanzeikbn = fldvalue.Trim
                            // Case "修繕その他負担税額"
                            // .Vari_Szen_sonotafutanzeigak = fldvalue.Trim
                            // Case "修繕その他負担者"
                            // .Vari_Szen_sonotafutanname = fldvalue.Trim
                            // Case "修繕その他負担備考"
                            // .Vari_Szen_bikosonotafutan = fldvalue.Trim
                            // 20160620 EXEUpdateに伴う修正2 -del end
                            case "退去予定時刻":
                                {
                                    model_cvitem.Vari_Tatiai_yoteitime = fldvalue.Trim();
                                    break;
                                }
                            case "退去後敬称":
                                {
                                    model_cvitem.Vari_Next_keisyo = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 契約解約情報 -add sta
                            case "解約月賃料は扱わないフラグ":
                                {
                                    model_cvitem.Vari_Nkin_kaiyakutukinoprintflg = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 契約解約情報 -add end
                            // 20160620 EXEUpdateに伴う修正2 -add sta
                            case "契約者修繕請求先No":
                                {
                                    model_cvitem.Vari_Szen_kysno = fldvalue;
                                    break;
                                }
                            case "契約者修繕入金区分":
                                {
                                    model_cvitem.Vari_Szen_kysnkbn = fldvalue;
                                    break;
                                }
                            case "契約者修繕送金率":
                                {
                                    model_cvitem.Vari_Szen_kyssorit = fldvalue;
                                    break;
                                }
                            case "契約者修繕送金額":
                                {
                                    model_cvitem.Vari_Szen_kyssogak = fldvalue;
                                    break;
                                }
                            case "家主修繕控除先No":
                                {
                                    model_cvitem.Vari_Szen_owno = fldvalue;
                                    break;
                                }
                            case "家主修繕控除請求区分":
                                {
                                    model_cvitem.Vari_Szen_owkaisyukbn = fldvalue;
                                    break;
                                }
                            case "家主修繕控除先送金ルールGuid":
                                {
                                    model_cvitem.Vari_Szen_owsosakisoruleguid = fldvalue;
                                    break;
                                }
                            case "家主修繕控除先送金ルールNo":
                                {
                                    model_cvitem.Vari_Szen_owsosakisoruleno = fldvalue;
                                    break;
                                }
                            case "家主修繕控除先口座No":
                                {
                                    model_cvitem.Vari_Szen_owsokozano = fldvalue;
                                    break;
                                }
                            case "家主修繕請求締日":
                                {
                                    model_cvitem.Vari_Szen_owsqsimeymd = fldvalue;
                                    break;
                                }
                            case "家主修繕請求入金区分":
                                {
                                    model_cvitem.Vari_Szen_ownkbn = fldvalue;
                                    break;
                                }
                            case "家主修繕請求振込先口座No":
                                {
                                    model_cvitem.Vari_Szen_owyatinkozano = fldvalue;
                                    break;
                                }
                            case "家主修繕請求先No":
                                {
                                    model_cvitem.Vari_Szen_owsqsakino = fldvalue;
                                    break;
                                }
                            // 20160620 EXEUpdateに伴う修正2 -add end
                            // 20160829 革命10バージョンアップに伴う修正 -add sta
                            case "家主修繕控除予定日":
                                {
                                    model_cvitem.Vari_Szen_owkojoymd = fldvalue;
                                    break;
                                }
                            case "家主修繕請求書発行予定日":
                                {
                                    model_cvitem.Vari_Szen_owsqymd = fldvalue;
                                    break;
                                }
                                // 20160829 革命10バージョンアップに伴う修正 -add end
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("ky_guid") = hash_bkhykyguid.Item(hash_cvitem.Item("ky_guid"))
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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
                Drop_TmpView(sqlcnnv10, viewnamebkhykytokyguid);

                // 契約基本情報の解約フラグを一括更新
                int tmpcnt = 0;
                string kaiflgupdatesql = Get_UseQry_Update();
                DBExec.Exec_NonQuery(sqlcnnv10, kaiflgupdatesql, ref tmpcnt);
                // 20160829 革命10バージョンアップに伴う修正 -add sta
                // 控除予定日の一括更新
                string kojodayupdatesql = Get_UseQry_Kojo();
                DBExec.Exec_NonQuery(sqlcnnv10, kojodayupdatesql, ref tmpcnt);

                // 請求書発行予定日の一括更新
                string sqdayupdatesql = Get_UseQry_Sq();
                DBExec.Exec_NonQuery(sqlcnnv10, sqdayupdatesql, ref tmpcnt);
                // 20160829 革命10バージョンアップに伴う修正 -add end
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
            /// 控除予定日の一括更新クエリ 20160829 革命10バージョンアップに伴う修正
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Kojo()
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " UPDATE kydata_kai ";
                tmp_sql = tmp_sql + " 	SET szen_owkojoymd = LEFT(CONVERT(VARCHAR,DATEADD(MONTH,SO.soymd1_sokintukikbn - 1,GETDATE()),112),6) + CONVERT(VARCHAR,SO.soymd1_sokinsimekbn) ";
                tmp_sql = tmp_sql + " FROM kydata_kai AS KYK ";
                tmp_sql = tmp_sql + " LEFT JOIN kydata AS KY ON KYK.ky_guid = KY.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN sorule AS SO ON KY.bk_guid = SO.relation_guid ";
                tmp_sql = tmp_sql + " WHERE szen_owkaisyukbn = 1 ";
                tmp_sql = tmp_sql + " AND   szen_owkojoymd IS NULL ";

                return tmp_sql;

            }

            /// <summary>
            /// 請求書発行予定日の一括更新クエリ 20160829 革命10バージョンアップに伴う修正
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Sq()
            {

                string tmp_sql = "";
                tmp_sql = tmp_sql + " UPDATE kydata_kai ";
                tmp_sql = tmp_sql + " 	SET szen_owsqymd = LEFT(CONVERT(VARCHAR,DATEADD(MONTH,SO.soymd1_sokintukikbn - 1,GETDATE()),112),6) + CONVERT(VARCHAR,SO.soymd1_sokinsimekbn) ";
                tmp_sql = tmp_sql + " FROM kydata_kai AS KYK ";
                tmp_sql = tmp_sql + " LEFT JOIN kydata AS KY ON KYK.ky_guid = KY.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN sorule AS SO ON KY.bk_guid = SO.relation_guid ";
                tmp_sql = tmp_sql + " WHERE szen_owkaisyukbn = 2 ";
                tmp_sql = tmp_sql + " AND   szen_owsqymd IS NULL ";

                return tmp_sql;

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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,ky_no) + '-' + CONVERT(varchar,ky_recno) FROM " + tblname;
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

                // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // 'flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, hash_guid)
                // Dim tmp_hash As New SafeDictionary<string, string>
                // flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, sqlcnnv10, tmp_hash)

                // '取得した部屋Noを半角変換して照合用に統一するrtn_hash
                // hash_guid = EtcMethod.Get_HashKeyChg(tmp_hash)
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);
                // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                // VIEW作成
                string tmp_sql_create = CommonModule.PRE_VIEW_QRY + viewname + CommonModule.POST_VIEW_QRY;

                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) + '-' + CONVERT(varchar,KY.ky_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,ky_guid ";
                tmp_sql_create = tmp_sql_create + " FROM kydata AS KY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";

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
            /// [kydata].[kaiyaku_flg]の一括更新クエリ
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE kydata SET ";
                tmp_sql = tmp_sql + " 	kaiyaku_flg = KYK.ky_recno ";
                tmp_sql = tmp_sql + " FROM kydata_kai AS KYK ";
                tmp_sql = tmp_sql + " LEFT JOIN kydata AS KY ON KYK.ky_guid = KY.ky_guid ";

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 契約修繕見積情報

    public class Kydata_kaiszen_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_kaiszen_Model();          // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列

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
                Set_KyGuid(sqlcnnv10, ref hash_bkhykyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_kaiszen";
                string fldnamegrp = "ky_guid,ky_recno,szen_mituno,mitu_title,mitu_bango," + "mitu_ymd,seiyaku_flg,seiyaku_ymd,futan_kbn,kys_futanrit," + "ow_futanrit,jisya_futanrit,zei_kbn,gokei_zeikbn,gokei_zeirit," + "kys_gokeizeigak,ow_gokeizeigak,jisya_gokeizeigak,hasu_futankbn";



                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);

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
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);
                string fldname_keysub3 = Conversions.ToString(headervalue[startrow - 1, keycol_sub3]);
                string fldname_keysub4 = Conversions.ToString(headervalue[startrow - 1, keycol_sub4]);

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

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4;




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
                            case "契約No":
                                {
                                    tmp_kyno = fldvalue;
                                    break;
                                }
                            case "更新No":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue;
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

                    // guid取得用に成形
                    model_cvitem.Vari_Ky_guid = tmp_bkno + "-" + tmp_hyno + "-" + tmp_kyno;

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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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
                tmp_sql = tmp_sql + " 	CONVERT(varchar,bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,ky_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,ky_recno) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,szen_mituno) ";
                tmp_sql = tmp_sql + " FROM kydata_kaiszen AS KYSZEN ";
                tmp_sql = tmp_sql + " LEFT JOIN kydata AS KY ON KYSZEN.ky_guid = KY.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

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

    #region 契約修繕見積詳細情報

    public class Kydata_kaiszenmeisai_Repository
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
                var hash_bkhykyguid = new SafeDictionary<string, string>();                            // キー/guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Kydata_kaiszenmeisai_Model();    // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                int keycol_sub4 = 5;                                  // サブ4キー列
                int keycol_sub5 = 6;                                  // サブ5キー列

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
                Set_KyGuid(sqlcnnv10, ref hash_bkhykyguid);

                // テーブル名/フィールド名セット
                string tblname = "kydata_kaiszenmeisai";
                string fldnamegrp = "ky_guid,ky_recno,szen_mituno,szen_meisaino,szen_name," + "tekiyo,mitu_suryo,mitu_tani,mitu_tanka,mitu_zeikbn," + "mitu_zeigak,kys_futanrit,ow_futanrit,jisya_futanrit,szen_gyno," + "jikko_suryo,jikko_tani,jikko_tanka,jikko_zeikbn,jikko_zeigak";



                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 親マスタ取得
                EtcMethod.Set_HashKeyToList(hash_bkhykyguid, ref list_basekeydata);

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
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);
                string fldname_keysub3 = Conversions.ToString(headervalue[startrow - 1, keycol_sub3]);
                string fldname_keysub4 = Conversions.ToString(headervalue[startrow - 1, keycol_sub4]);
                string fldname_keysub5 = Conversions.ToString(headervalue[startrow - 1, keycol_sub5]);

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
                    string tmp_keysub5 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub5]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2 + "-" + tmp_keysub3 + "-" + tmp_keysub4 + "-" + tmp_keysub5;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3 + "、" + fldname_keysub4 + " = " + tmp_keysub4 + "、" + fldname_keysub5 + " = " + tmp_keysub5;





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
                            case "契約No":
                                {
                                    tmp_kyno = fldvalue;
                                    break;
                                }
                            case "更新No":
                                {
                                    model_cvitem.Vari_Ky_recno = fldvalue;
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

                    // guid取得用に成形
                    model_cvitem.Vari_Ky_guid = tmp_bkno + "-" + tmp_hyno + "-" + tmp_kyno;

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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // hash_cvitem("ky_guid") = hash_bkhykyguid.Item(StrConv(hash_cvitem.Item("ky_guid"), VbStrConv.Narrow))
                        hash_cvitem["ky_guid"] = hash_bkhykyguid[hash_cvitem["ky_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

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
                tmp_sql = tmp_sql + " 	CONVERT(varchar,bk_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,hy_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,ky_no) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,ky_recno) + '-' +  ";
                tmp_sql = tmp_sql + " 	CONVERT(varchar,szen_mituno) ";
                tmp_sql = tmp_sql + " FROM kydata_kaiszen AS KYSZEN ";
                tmp_sql = tmp_sql + " LEFT JOIN kydata AS KY ON KYSZEN.ky_guid = KY.ky_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON KY.bk_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN hydata AS HY ON KY.hy_guid = HY.hy_guid ";
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

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





    // 20160531 鍵情報移行処理の修正 -del sta
    // #Region "契約鍵情報"

    // Public Class Kydata_kagi_Repository

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
    // Dim list_basekeydata As New List(Of String)                     '親マスタ有無チェック用リスト
    // Dim sortlist_log As New SortedList(Of Integer, String)          'ログ格納用オブジェクト
    // Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
    // Dim hash_guid As New SafeDictionary<string, string>                                  'guid格納用ハッシュテーブル

    // Dim tmp_hash As New SafeDictionary<string, string>                                   '作業用ハッシュテーブル
    // Dim normalflg As Boolean                                        'INSERT正常終了フラグ
    // Dim rtn As Boolean = True                                       '戻り値

    // Dim model_cvitem As New Njc.Model.Hydata_kagi_Model             '移行値格納用モデル初期化
    // Dim keycol_main As Integer = 1                                  'メインキー列
    // Dim keycol_sub1 As Integer = 2                                  'サブ1キー列
    // Dim keycol_sub2 As Integer = 3                                  'サブ2キー列
    // Dim viewname As String = "物件部屋キー情報"

    // '************************
    // '作業準備
    // '************************

    // '10鍵タイトルマスタから鍵Noとタイトル名を紐付けたデータを取得
    // Call EtcMethod.Set_KagiTitleMst(sqlcnnv10, 1)
    // Call EtcMethod.Set_KagiTitleMst(sqlcnnv10, 2)

    // 'Excelファイル初期設定
    // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

    // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
    // If Not rtn Then
    // Return rtn
    // End If

    // 'キー取得用のVIEWを作成
    // Call Me.Create_View_KeyAndGuid(sqlcnnv10)

    // 'テーブル名/フィールド名セット
    // Dim viewname_base As String = PRE_VIEW_NAME & viewname
    // Dim tblname As String = "hydata_kagi"
    // Dim fldnamegrp As String = "hy_guid,kagi_no,honsu,biko,history," & _
    // "hokan,gyshare"

    // '追加コンバート時の重複チェック用に既存データのキーを取得
    // If Not InitDBFlg Then
    // Call Me.Get_ExistData(sqlcnnv10, tblname, list_chkduplicate)
    // End If

    // 'guidを取得
    // Call Me.Get_Guid(sqlcnnv10, viewname_base, hash_guid)

    // '親マスタ取得
    // '2016.04.26 メインの方へも反映させる修正 -chg sta
    // 'Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
    // Call EtcMethod.Set_HashKeyToList(hash_guid, list_basekeydata)
    // '2016.04.26 メインの方へも反映させる修正 -chg end

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

    // '---------------
    // 'ヘッダー処理
    // '---------------
    // 'ヘッダー行取得
    // Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

    // 'キーヘッダー名取得
    // Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
    // Dim fldname_keysub1 As String = headervalue(startrow - 1, keycol_sub1)
    // Dim fldname_keysub2 As String = headervalue(startrow - 1, keycol_sub2)

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
    // Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub1 & "-" & tmp_keysub2
    // .Vari_Hy_guid = tmp_keymain & "-" & tmp_keysub1

    // 'ログ出力用
    // Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2

    // '移行値取得
    // For cntjj = 1 To columncnt

    // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
    // Dim fldvalue As String = ""
    // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
    // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
    // End If

    // Select Case fldname
    // Case "鍵タイトル名"
    // .Vari_Kagi_no = fldvalue.Trim
    // Case "鍵本数"
    // .Vari_Honsu = fldvalue.Trim
    // Case "備考"
    // .Vari_Biko = fldvalue.Trim
    // Case "保管場所"
    // .Vari_Hokan = fldvalue.Trim
    // Case "業者間での情報共有"
    // .Vari_Gyshare = fldvalue.Trim
    // End Select

    // Next

    // '固定値
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
    // hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))

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
    // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,kagi_no) FROM " & tblname
    // tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
    // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
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
    // Public Sub Create_View_KeyAndGuid(ByVal sqlcnnv10 As SqlConnection)

    // Dim viewname As String = "物件部屋キー情報"
    // Dim tmpcnt As Integer = 0

    // 'VIEW初期化
    // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
    // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmpcnt)

    // 'VIEW作成
    // Dim tmp_sql_create As String = ""
    // tmp_sql_create = tmp_sql_create & PRE_VIEW_QRY & PRE_VIEW_NAME & viewname & POST_VIEW_QRY
    // tmp_sql_create = tmp_sql_create & " SELECT "
    // tmp_sql_create = tmp_sql_create & " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー "
    // tmp_sql_create = tmp_sql_create & " 	,HY.hy_guid "
    // tmp_sql_create = tmp_sql_create & " FROM hydata AS HY "
    // tmp_sql_create = tmp_sql_create & " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid "
    // DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmpcnt)

    // End Sub

    // End Class

    // End Class

    // #End Region
    // 20160531 鍵情報移行処理の修正 -del end

}