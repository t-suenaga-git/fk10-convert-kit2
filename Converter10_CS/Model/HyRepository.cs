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

    #region 部屋基本情報

    public class Hydata_Repository
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

                var model_cvitem = new Model.Hydata_Model();                  // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 20160720 連動情報構築 -chg sta
                // Call Me.Get_RelData()
                SetRelItemToObject.Set_RelData_Hyrui();
                // 20160720 連動情報構築 -chg end
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
                string tblname = "hydata";
                // 20161012 革命10アップデートに伴う修正 -add (最後尾にwmp_idを追加)
                // 20160519 EXEUpdateに伴う修正 部屋基本情報 -add (最後尾にsort_hy_no、lastupdateを追加)
                string fldnamegrp = "hy_guid,bk_guid,hy_no,hy_gaibuno,hy_deleteflg," + "delete_guid,delete_day,delete_cnt,hy_ruinokbn,madori_cnt," + "madori_typekbn,madori_biko,men_senyujitu,men_senyujitutubo,men_yuka," + "men_yukatubo,men_senyutouki,men_senyutoukitubo,men_toki,men_tokitubo," + "men_balcony,men_balconytubo,men_tempo,men_tempotubo,men_jutaku," + "men_jutakutubo,history,rowid,delete_cause,kaiyaku_uketukekbn," + "kaiyaku_months,kaiyaku_days,kaiyaku_day,sort_hy_no,lastupdate," + "wmp_id";







                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, tblname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, tblname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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
                string fldname_keysub = readtbl.Columns[keycol_sub - 1].ColumnName.Trim();

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
                    string tmp_madori = "";

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

                        // 各項目値→変数格納
                        switch (fldname ?? "")
                        {
                            case "物件NO":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
                            case "部屋NO":
                                {
                                    model_cvitem.Vari_Hy_no = fldvalue.Trim();
                                    break;
                                }
                            case "部屋分類名":
                                {
                                    model_cvitem.Vari_Hy_ruinokbn = fldvalue.Trim();
                                    break;
                                }
                            case "間取り":
                                {
                                    tmp_madori = fldvalue.Trim();
                                    break;
                                }
                            case "間取り備考":
                                {
                                    model_cvitem.Vari_Madori_biko = fldvalue.Trim();
                                    break;
                                }
                            case "専有実面積(m2)":
                                {
                                    model_cvitem.Vari_Men_senyujitu = fldvalue.Trim();
                                    break;
                                }
                            case "専有実面積(坪)":
                                {
                                    model_cvitem.Vari_Men_senyujitutubo = fldvalue.Trim();
                                    break;
                                }
                            case "床面積(m2)":
                                {
                                    model_cvitem.Vari_Men_yuka = fldvalue.Trim();
                                    break;
                                }
                            case "床面積(坪)":
                                {
                                    model_cvitem.Vari_Men_yukatubo = fldvalue.Trim();
                                    break;
                                }
                            case "専有登記面積(m2)":
                                {
                                    model_cvitem.Vari_Men_senyutouki = fldvalue.Trim();
                                    break;
                                }
                            case "専有登記面積(坪)":
                                {
                                    model_cvitem.Vari_Men_senyutoukitubo = fldvalue.Trim();
                                    break;
                                }
                            case "登記延床面積(m2)":
                                {
                                    model_cvitem.Vari_Men_toki = fldvalue.Trim();
                                    break;
                                }
                            case "登記延床面積(坪)":
                                {
                                    model_cvitem.Vari_Men_tokitubo = fldvalue.Trim();
                                    break;
                                }
                            case "バルコニー面積(m2)":
                                {
                                    model_cvitem.Vari_Men_balcony = fldvalue.Trim();
                                    break;
                                }
                            case "バルコニー面積(坪)":
                                {
                                    model_cvitem.Vari_Men_balconytubo = fldvalue.Trim();
                                    break;
                                }
                            case "店舗付き住宅店舗部分面積(m2)":
                                {
                                    model_cvitem.Vari_Men_tempo = fldvalue.Trim();
                                    model_cvitem.Vari_Men_tempotubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(fldvalue.Trim())), Typ.RoundingTypes.Kirisute, 2).ToString();
                                    break;
                                }
                            case "店舗付き住宅店舗部分面積(坪)":
                                {
                                    break;
                                }
                            // m2を変換
                            case "店舗付き住宅住宅部分面積(m2)":
                                {
                                    model_cvitem.Vari_Men_jutaku = fldvalue.Trim();
                                    model_cvitem.Vari_Men_jutakutubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(fldvalue.Trim())), Typ.RoundingTypes.Kirisute, 2).ToString();
                                    break;
                                }
                            case "店舗付き住宅住宅部分面積(坪)":
                                {
                                    break;
                                }
                            // m2を変換
                            case "解約受付区分":
                                {
                                    model_cvitem.Vari_Kaiyaku_uketukekbn = fldvalue.Trim();
                                    break;
                                }
                            case "解約受付月数":
                                {
                                    model_cvitem.Vari_Kaiyaku_months = fldvalue.Trim();
                                    break;
                                }
                            case "解約受付日数":
                                {
                                    model_cvitem.Vari_Kaiyaku_days = fldvalue.Trim();
                                    break;
                                }
                            case "解約受付日にち":
                                {
                                    model_cvitem.Vari_Kaiyaku_day = fldvalue.Trim();
                                    break;
                                }
                            case "自社Web内部キー採番ID":    // 20161012 革命10アップデートに伴う修正 -add
                                {
                                    model_cvitem.Vari_Wmp_id = fldvalue.Trim();
                                    break;
                                }
                        }

                    }
                    // 20161004 汎用コンバート時の平米→坪変換処理の追加 -add sta
                    // 汎用の場合は平米から坪を算出する
                    if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                    {
                        model_cvitem.Vari_Men_senyujitutubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_senyujitu)), Typ.RoundingTypes.Kirisute, 2).ToString();
                        model_cvitem.Vari_Men_yukatubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_yuka)), Typ.RoundingTypes.Kirisute, 2).ToString();
                        model_cvitem.Vari_Men_balconytubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_balcony)), Typ.RoundingTypes.Kirisute, 2).ToString();
                    }
                    // 20161004 汎用コンバート時の平米→坪変換処理の追加 -add end
                    // 間取変換
                    string tmp_madoricnt = "";
                    string tmp_madorinaiyo = "";

                    MadoriConv.SplitMadori(tmp_madori, ref tmp_madoricnt, ref tmp_madorinaiyo);

                    model_cvitem.Vari_Madori_cnt = tmp_madoricnt;
                    model_cvitem.Vari_Madori_typekbn = tmp_madorinaiyo;

                    // 固定値
                    model_cvitem.Vari_Hy_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Hy_gaibuno = 0.ToString();    // 後で一括更新する
                    model_cvitem.Vari_Hy_deleteflg = 0.ToString();
                    model_cvitem.Vari_Delete_guid = null;
                    model_cvitem.Vari_Delete_day = null;
                    model_cvitem.Vari_Delete_cnt = null;
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Delete_cause = null;
                    // 20160519 EXEUpdateに伴う修正 部屋基本情報 -add sta
                    // とりあえずNULLにしておく (必要であればここで値を設定)
                    model_cvitem.Vari_Sort_hy_no = null;
                    model_cvitem.Vari_Lastupdate = null;
                    // 20160519 EXEUpdateに伴う修正 部屋基本情報 -add end

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
                // 部屋基本情報テーブル外部No一括更新
                int tmpcnt = 0;
                string gaibunoupdatesql = Get_UseQry_Update();
                DBExec.Exec_NonQuery(sqlcnnv10, gaibunoupdatesql, ref tmpcnt);

                // 外部No管理テーブル一括更新
                var obj_hydatagaibuno = new Hydata_gaibuno_Repository.SubConv();
                rtn = obj_hydatagaibuno.Set_Hydata_Gaibuno(sqlcnnv10);

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

                string tmp_sql = " SELECT bk_no FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

                string tmp_sql = "SELECT bk_no,bk_guid FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash_guid);

            }

            /// <summary>
            /// [hydata].[hy_gaibuno]の一括更新クエリ
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE hydata SET ";
                tmp_sql = tmp_sql + " 	hy_gaibuno = HYGAIBUNO.GAIBUNO ";
                tmp_sql = tmp_sql + " FROM hydata AS HY ";
                tmp_sql = tmp_sql + " LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 ROW_NUMBER()OVER(PARTITION BY bk_guid ORDER BY bk_guid,hy_no) AS GAIBUNO ";
                tmp_sql = tmp_sql + " 			,bk_guid ";
                tmp_sql = tmp_sql + " 			,hy_guid ";
                tmp_sql = tmp_sql + " 		FROM hydata ";
                tmp_sql = tmp_sql + " 	) AS HYGAIBUNO ";
                tmp_sql = tmp_sql + " ON  HY.bk_guid = HYGAIBUNO.bk_guid ";
                tmp_sql = tmp_sql + " AND HY.hy_guid = HYGAIBUNO.hy_guid ";
                tmp_sql = tmp_sql + " ; ";

                return tmp_sql;

            }
            // 20160720 連動情報構築 -del sta
            // ''' <summary>
            // ''' 部屋分類紐付情報取得
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
            // Dim relsheetname As String = "部屋分類マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldhyruiname As String = ""
            // Dim tmp_newhyruino As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元部屋分類名称"
            // tmp_oldhyruiname = fldvalue.Trim
            // Case "賃貸革命10部屋分類No"
            // tmp_newhyruino = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newhyruino <> "" Then
            // Hash_Rel_Hyrui.Add(tmp_oldhyruiname, tmp_newhyruino)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            // End Sub
            // 20160720 連動情報構築 -del end
        }

    }

    #endregion

    #region 部屋外部No管理情報

    // 部屋基本情報作成直後に呼び出して処理する

    public class Hydata_gaibuno_Repository
    {

        public class SubConv
        {

            /// <summary>
            /// 部屋基本情報(hydata)から外部Noを取得して外部No管理テーブル(hydata_gaibuno)へ挿入する
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Hydata_Gaibuno(SqlConnection sqlcnnv10)
            {

                bool rtn = true;
                var hash_gaibuno = new SafeDictionary<string, string>();
                var tmp_hash = new SafeDictionary<string, string>();
                var model_cvitem = new Model.Hydata_gaibuno_Model();
                bool normalflg = true;
                int tmp_cnt = 0;

                // 対象テーブル初期化
                string tmp_sql_delete = " DELETE FROM hydata_gaibuno ";
                rtn = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_delete, ref tmp_cnt);

                // テーブル名/フィールド名セット
                string tblname = "hydata_gaibuno";
                string fldnamegrp = "bk_guid,current_no,history";

                // hydataから外部No最大値を集約して必要なデータを抽出
                string tmp_sql_select = Get_UseQry_Select();

                // 抽出結果をハッシュテーブルへ格納
                DBExec.Exec_DataReader_Col_Hash(tmp_sql_select, ref sqlcnnv10, ref hash_gaibuno);

                // データ取得→挿入処理

                foreach (var gaibunodata in hash_gaibuno)
                {

                    model_cvitem.Vari_Bk_guid = Conversions.ToString(gaibunodata.Key);
                    model_cvitem.Vari_Current_no = Conversions.ToString(gaibunodata.Value);
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // 挿入処理
                    CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, ref normalflg);


                }

                return rtn;

            }

            /// <summary>
            /// 最大外部Noを持つ物件情報取得クエリ作成
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Select()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 bk_guid ";
                tmp_sql = tmp_sql + " 	,hy_gaibuno ";
                tmp_sql = tmp_sql + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_guid ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY bk_guid ORDER BY hy_gaibuno DESC) AS 抽出用連番 ";
                tmp_sql = tmp_sql + " 		,hy_gaibuno ";
                tmp_sql = tmp_sql + " 	FROM hydata ";
                tmp_sql = tmp_sql + " ) AS VW ";
                tmp_sql = tmp_sql + " WHERE 抽出用連番 = 1 ";

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 部屋詳細情報

    public class Hydata_detail_Repository
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

                var model_cvitem = new Model.Hydata_detail_Model();           // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列
                string viewname = "物件部屋キー情報";

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 2016.04.06 紐付データ取得処理を外出し -chg sta
                // Call Me.Get_RelData()
                SetRelItemToObject.Set_RelData_Toritaiyo();
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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_detail";
                // 20160519 EXEUpdateに伴う修正 部屋詳細情報 -chg sta
                // Dim fldnamegrp As String = "hy_guid,hygyomu_setflg,syozai_kaisu1,syozai_kaisu2,syozai_kaisu3," & _
                // "syozai_tikaflg1,syozai_tikaflg2,syozai_tikaflg3,mukikbn,kadoheya," & _
                // "saikokbn,balcony_mukikbn,nyukyo_jokyokbn,nyukyo_jokyomemo,nyukyo_syunflg," & _
                // "nyukyo_ym,nyukyo_syunkbn,nyukyo_joken,kakunin_ymd,bosyu_startflg," & _
                // "freerent_flg,freerent_month,freerent_detaill,hoken_kbn,hoken_kikan," & _
                // "hoken_gak,hoken_biko,torihiki_taiyokbn,torihiki_kyakutuke,torihiki_tesumoto," & _
                // "torihiki_tesukyaku,torihiki_futankasi,torihiki_futankari,torihiki_kyakutukecomment,torihiki_gykokokukatudokbn," & _
                // "moto_gy_fudono,koukoku_ryoukbn,koukoku_jogengak,koukoku_jokennaiyo,biko," & _
                // "addr_replaceflg,addr_replacecyome,addr_replacecyomeptn,addr_replacebanti,addr_replaceetc," & _
                // "hosyo_gyno,hosyo_naiyo,kyrui_nokbn,keiyaku_kikan,keiyaku_kijitu," & _
                // "parking_biko,parking_bikebiko,parking_cyurinbiko,shared_salespoint,history," & _
                // "keisai_bkflg,keisai_hyflg,keisai_bantiflg,keisai_mapflg,hosyo_kbn," & _
                // "koukoku_jogengakkbn,koukoku_jogenrit,koukoku_jogentaxkbn,commonsalespoint_useflg,floors_flg," & _
                // "jisya_no,jisya_tanto,confirmky_sekininsya,confirmky_ymd,confirmky_print," & _
                // "confirmkai_sekininsya,confirmkai_ymd,confirmkai_print,nextnkin_kosindefault,toki_ymd," & _
                // "syo_kenriflg,syo_kenrikbn,other_kenriflg,btob_groupkbn,jisyaweb_osusumebk," & _
                // "kaiyaku_ym,taikyo_ym,svbunrui_no,krbunrui_no,siyo_mokuteki," & _
                // "nyukyo_sintikukbn,nyukyo_jikikbn,torihiki_jisyakbn,keiyaku_kikankbn"
                // 20160829 革命10バージョンアップに伴う修正 saikokbnを削除 -del
                string fldnamegrp = "hy_guid,hygyomu_setflg,syozai_kaisu1,syozai_kaisu2,syozai_kaisu3," + "syozai_tikaflg1,syozai_tikaflg2,syozai_tikaflg3,mukikbn,kadoheya," + "balcony_mukikbn,nyukyo_jokyokbn,nyukyo_jokyomemo,nyukyo_syunflg," + "nyukyo_ym,nyukyo_syunkbn,nyukyo_joken,kakunin_ymd,bosyu_startflg," + "freerent_flg,freerent_month,freerent_detaill,hoken_kbn,hoken_kikan," + "hoken_gak,hoken_biko,torihiki_taiyokbn,torihiki_kyakutuke,torihiki_tesumoto," + "torihiki_tesukyaku,torihiki_futankasi,torihiki_futankari,torihiki_kyakutukecomment,torihiki_gykokokukatudokbn," + "moto_gy_fudono,koukoku_ryoukbn,koukoku_jogengak,koukoku_jokennaiyo,biko," + "addr_replaceflg,addr_replacecyome,addr_replacecyomeptn,addr_replacebanti,addr_replaceetc," + "hosyo_gyno,hosyo_naiyo,kyrui_nokbn,keiyaku_kikan,keiyaku_kijitu," + "parking_biko,parking_bikebiko,parking_cyurinbiko,shared_salespoint,history," + "keisai_bkflg,keisai_hyflg,keisai_bantiflg,keisai_mapflg,hosyo_kbn," + "koukoku_jogengakkbn,koukoku_jogenrit,koukoku_jogentaxkbn,commonsalespoint_useflg,floors_flg," + "jisya_no,jisya_tanto,nextnkin_kosindefault,toki_ymd,syo_kenriflg," + "syo_kenrikbn,other_kenriflg,btob_groupkbn,jisyaweb_osusumebk,kaiyaku_ym," + "taikyo_ym,svbunrui_no,krbunrui_no,siyo_mokuteki,nyukyo_sintikukbn," + "nyukyo_jikikbn,torihiki_jisyakbn,keiyaku_kikankbn";















                // 20160519 EXEUpdateに伴う修正 部屋詳細情報 -chg end

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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
                string fldname_keysub = readtbl.Columns[keycol_sub - 1].ColumnName.Trim();

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
                    model_cvitem.Vari_Hy_guid = fldvalue_key;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + fldvalue_keymain + "、" + fldname_keysub + " = " + fldvalue_keysub;

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
                            case "部屋業務期間有無":
                                {
                                    model_cvitem.Vari_Hygyomu_setflg = fldvalue.Trim();
                                    break;
                                }
                            case "所在階１":
                                {
                                    model_cvitem.Vari_Syozai_kaisu1 = fldvalue.Trim();
                                    break;
                                }
                            case "所在階２":
                                {
                                    model_cvitem.Vari_Syozai_kaisu2 = fldvalue.Trim();
                                    break;
                                }
                            case "所在階３":
                                {
                                    model_cvitem.Vari_Syozai_kaisu3 = fldvalue.Trim();
                                    break;
                                }
                            case "地下フラグ１":
                                {
                                    model_cvitem.Vari_Syozai_tikaflg1 = fldvalue.Trim();
                                    break;
                                }
                            case "地下フラグ２":
                                {
                                    model_cvitem.Vari_Syozai_tikaflg2 = fldvalue.Trim();
                                    break;
                                }
                            case "地下フラグ３":
                                {
                                    model_cvitem.Vari_Syozai_tikaflg3 = fldvalue.Trim();
                                    break;
                                }
                            case "向き":
                                {
                                    model_cvitem.Vari_Mukikbn = fldvalue.Trim();
                                    break;
                                }
                            case "角部屋フラグ":
                                {
                                    model_cvitem.Vari_Kadoheya = fldvalue.Trim();
                                    break;
                                }
                            // 20160829 革命10バージョンアップに伴う修正 -del sta
                            // Case "採光"
                            // .Vari_Saikokbn = fldvalue.Trim
                            // 20160829 革命10バージョンアップに伴う修正 -del end
                            case "バルコニー方向":
                                {
                                    model_cvitem.Vari_Balcony_mukikbn = fldvalue.Trim();
                                    break;
                                }
                            case "状況":
                                {
                                    model_cvitem.Vari_Nyukyo_jokyokbn = fldvalue.Trim();
                                    break;
                                }
                            case "入居状況備考":
                                {
                                    model_cvitem.Vari_Nyukyo_jokyomemo = fldvalue.Trim();
                                    break;
                                }
                            case "入居可能日を時期で指定":
                                {
                                    model_cvitem.Vari_Nyukyo_syunflg = fldvalue.Trim();
                                    break;
                                }
                            case "入居可能時期(年月)":
                                {
                                    model_cvitem.Vari_Nyukyo_ym = fldvalue.Trim();
                                    break;
                                }
                            case "入居可能時期(上旬・中旬・下旬)":
                                {
                                    model_cvitem.Vari_Nyukyo_syunkbn = fldvalue.Trim();
                                    break;
                                }
                            case "入居条件":
                                {
                                    model_cvitem.Vari_Nyukyo_joken = fldvalue.Trim();
                                    break;
                                }
                            case "広告内容確認日":
                                {
                                    model_cvitem.Vari_Kakunin_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "募集可能種別(しない・する)":
                                {
                                    model_cvitem.Vari_Bosyu_startflg = fldvalue.Trim();
                                    break;
                                }
                            case "フリーレント有無":
                                {
                                    model_cvitem.Vari_Freerent_flg = fldvalue.Trim();
                                    break;
                                }
                            case "フリーレントカ月":
                                {
                                    model_cvitem.Vari_Freerent_month = fldvalue.Trim();
                                    break;
                                }
                            case "フリーレント詳細":
                                {
                                    model_cvitem.Vari_Freerent_detaill = fldvalue.Trim();
                                    break;
                                }
                            case "保険-保険の利用":
                                {
                                    model_cvitem.Vari_Hoken_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "保険-保険期間":
                                {
                                    model_cvitem.Vari_Hoken_kikan = fldvalue.Trim();
                                    break;
                                }
                            case "保険-保険料":
                                {
                                    model_cvitem.Vari_Hoken_gak = fldvalue.Trim();
                                    break;
                                }
                            case "保険-保険の備考":
                                {
                                    model_cvitem.Vari_Hoken_biko = fldvalue.Trim();
                                    break;
                                }
                            case "取引形態種別":
                                {
                                    model_cvitem.Vari_Torihiki_taiyokbn = fldvalue.Trim();
                                    break;
                                }
                            case "客付け状態可否":
                                {
                                    model_cvitem.Vari_Torihiki_kyakutuke = fldvalue.Trim();
                                    break;
                                }
                            case "手数料負担割合貸主(%)":
                                {
                                    model_cvitem.Vari_Torihiki_futankasi = fldvalue.Trim();
                                    break;
                                }
                            case "手数料負担割合借主(%)":
                                {
                                    model_cvitem.Vari_Torihiki_futankari = fldvalue.Trim();
                                    break;
                                }
                            case "手数料配分元付(%)":
                                {
                                    model_cvitem.Vari_Torihiki_tesumoto = fldvalue.Trim();
                                    break;
                                }
                            case "手数料配分先物(%)":
                                {
                                    model_cvitem.Vari_Torihiki_tesukyaku = fldvalue.Trim();
                                    break;
                                }
                            case "客付会社への物件コメント":
                                {
                                    model_cvitem.Vari_Torihiki_kyakutukecomment = fldvalue.Trim();
                                    break;
                                }
                            case "業者間広告広告活動種別":
                                {
                                    model_cvitem.Vari_Torihiki_gykokokukatudokbn = fldvalue.Trim();
                                    break;
                                }
                            case "情報元業者No":
                                {
                                    model_cvitem.Vari_Moto_gy_fudono = fldvalue.Trim();
                                    break;
                                }
                            case "広告料有無":
                                {
                                    model_cvitem.Vari_Koukoku_ryoukbn = fldvalue.Trim();
                                    break;
                                }
                            case "広告料上限額":
                                {
                                    model_cvitem.Vari_Koukoku_jogengak = fldvalue.Trim();
                                    break;
                                }
                            case "広告料条件内容":
                                {
                                    model_cvitem.Vari_Koukoku_jokennaiyo = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "部屋住所の変更フラグ":
                                {
                                    model_cvitem.Vari_Addr_replaceflg = fldvalue.Trim();
                                    break;
                                }
                            case "部屋住所(丁目)":
                                {
                                    model_cvitem.Vari_Addr_replacecyome = fldvalue.Trim();
                                    break;
                                }
                            case "部屋住所(丁目区分)":
                                {
                                    model_cvitem.Vari_Addr_replacecyomeptn = fldvalue.Trim();
                                    break;
                                }
                            case "部屋住所(町地域)":
                                {
                                    model_cvitem.Vari_Addr_replacebanti = fldvalue.Trim();
                                    break;
                                }
                            case "部屋住所(その他)":
                                {
                                    model_cvitem.Vari_Addr_replaceetc = fldvalue.Trim();
                                    break;
                                }
                            case "保証会社":
                                {
                                    model_cvitem.Vari_Hosyo_gyno = fldvalue.Trim();
                                    break;
                                }
                            case "保証内容":
                                {
                                    model_cvitem.Vari_Hosyo_naiyo = fldvalue.Trim();
                                    break;
                                }
                            case "契約の種類":
                                {
                                    model_cvitem.Vari_Kyrui_nokbn = fldvalue.Trim();
                                    break;
                                }
                            case "定期借家契約(期間)":
                                {
                                    model_cvitem.Vari_Keiyaku_kikan = fldvalue.Trim();
                                    break;
                                }
                            case "定期借家契約(期日)":
                                {
                                    model_cvitem.Vari_Keiyaku_kijitu = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場備考":
                                {
                                    model_cvitem.Vari_Parking_biko = fldvalue.Trim();
                                    break;
                                }
                            case "バイク駐車場備考":
                                {
                                    model_cvitem.Vari_Parking_bikebiko = fldvalue.Trim();
                                    break;
                                }
                            case "駐輪場備考":
                                {
                                    model_cvitem.Vari_Parking_cyurinbiko = fldvalue.Trim();
                                    break;
                                }
                            case "共通セールスポイント":
                                {
                                    model_cvitem.Vari_Shared_salespoint = fldvalue.Trim();
                                    break;
                                }
                            case "不動産検索サイト掲載設定-物件名":
                                {
                                    model_cvitem.Vari_Keisai_bkflg = fldvalue.Trim();
                                    break;
                                }
                            case "不動産検索サイト掲載設定-部屋NO":
                                {
                                    model_cvitem.Vari_Keisai_hyflg = fldvalue.Trim();
                                    break;
                                }
                            case "不動産検索サイト掲載設定-丁番地以下":
                                {
                                    model_cvitem.Vari_Keisai_bantiflg = fldvalue.Trim();
                                    break;
                                }
                            case "不動産検索サイト掲載設定-地図上":
                                {
                                    model_cvitem.Vari_Keisai_mapflg = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸保証利用区分":
                                {
                                    model_cvitem.Vari_Hosyo_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "広告料上限額区分":
                                {
                                    model_cvitem.Vari_Koukoku_jogengakkbn = fldvalue.Trim();
                                    break;
                                }
                            case "広告料上限率":
                                {
                                    model_cvitem.Vari_Koukoku_jogenrit = fldvalue.Trim();
                                    break;
                                }
                            case "広告料上限額税区分":
                                {
                                    model_cvitem.Vari_Koukoku_jogentaxkbn = fldvalue.Trim();
                                    break;
                                }
                            case "共通セールスポイント使用フラグ":
                                {
                                    model_cvitem.Vari_Commonsalespoint_useflg = fldvalue.Trim();
                                    break;
                                }
                            case "複数階有りのフラグ":
                                {
                                    model_cvitem.Vari_Floors_flg = fldvalue.Trim();
                                    break;
                                }
                            case "支店NO":
                                {
                                    model_cvitem.Vari_Jisya_no = fldvalue.Trim();
                                    break;
                                }
                            case "自社担当者No":
                                {
                                    model_cvitem.Vari_Jisya_tanto = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 部屋詳細情報 -del sta
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
                            // 20160519 EXEUpdateに伴う修正 部屋詳細情報 -del end
                            case "次回更新時の初期値":
                                {
                                    model_cvitem.Vari_Nextnkin_kosindefault = fldvalue.Trim();
                                    break;
                                }
                            case "登記情報の日付":
                                {
                                    model_cvitem.Vari_Toki_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "所有権にかかる権利有無":
                                {
                                    model_cvitem.Vari_Syo_kenriflg = fldvalue.Trim();
                                    break;
                                }
                            case "所有権にかかる権利の種類":
                                {
                                    model_cvitem.Vari_Syo_kenrikbn = fldvalue.Trim();
                                    break;
                                }
                            case "所有権以外の権利有無":
                                {
                                    model_cvitem.Vari_Other_kenriflg = fldvalue.Trim();
                                    break;
                                }
                            case "BtoBプラグイングループ設定区分":
                                {
                                    model_cvitem.Vari_Btob_groupkbn = fldvalue.Trim();
                                    break;
                                }
                            case "自社Webオススメ物件表示":
                                {
                                    model_cvitem.Vari_Jisyaweb_osusumebk = fldvalue.Trim();
                                    break;
                                }
                            case "解約日":
                                {
                                    model_cvitem.Vari_Kaiyaku_ym = fldvalue.Trim();
                                    break;
                                }
                            case "退去日":
                                {
                                    model_cvitem.Vari_Taikyo_ym = fldvalue.Trim();
                                    break;
                                }
                            case "使用目的":
                                {
                                    model_cvitem.Vari_Siyo_mokuteki = fldvalue.Trim();
                                    break;
                                }
                            case "新築区分":
                                {
                                    model_cvitem.Vari_Nyukyo_sintikukbn = fldvalue.Trim();
                                    break;
                                }
                            case "入居時期区分":
                                {
                                    model_cvitem.Vari_Nyukyo_jikikbn = fldvalue.Trim();
                                    break;
                                }
                            case "取引自社区分":
                                {
                                    model_cvitem.Vari_Torihiki_jisyakbn = fldvalue.Trim();
                                    break;
                                }
                            case "契約期間区分":
                                {
                                    model_cvitem.Vari_Keiyaku_kikankbn = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Svbunrui_no = 0.ToString();   // 京王カスタマイズ分
                    model_cvitem.Vari_Krbunrui_no = 0.ToString();   // 使用箇所不明なためデフォルト値を設定
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 20161012 部屋所在階チェック処理の追加 -add
                    Chk_Kaisu(sqlcnnv10, tblname, ref hash_cvitem, ref hash_log);

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 部屋所在階のチェック '20161012 部屋所在階チェック処理の追加
            /// 例)物件階数=5階の場合で部屋階数=6階はあり得ない
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_log"></param>
            /// <remarks></remarks>
            public void Chk_Kaisu(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_cvitem, ref SafeDictionary<string, string> hash_log)
            {

                string[] tmp_key = hash_cvitem["hy_guid"].ToString().Split('-');
                string bkno = tmp_key[0];
                // Dim kaisu1 As Integer = (IIf(hash_cvitem("syozai_kaisu1") = "", 0, Int32.Parse(hash_cvitem("syozai_kaisu1"))))
                // Dim kaisu2 As Integer = (IIf(hash_cvitem("syozai_kaisu2") = "", 0, Int32.Parse(hash_cvitem("syozai_kaisu2"))))
                // Dim kaisu3 As Integer = (IIf(hash_cvitem("syozai_kaisu3") = "", 0, Int32.Parse(hash_cvitem("syozai_kaisu3"))))
                string tmp_kaisu1 = Conversions.ToString(Interaction.IIf(hash_cvitem["syozai_kaisu1"] is null, "", hash_cvitem["syozai_kaisu1"]));
                string tmp_kaisu2 = Conversions.ToString(Interaction.IIf(hash_cvitem["syozai_kaisu2"] is null, "", hash_cvitem["syozai_kaisu2"]));
                string tmp_kaisu3 = Conversions.ToString(Interaction.IIf(hash_cvitem["syozai_kaisu3"] is null, "", hash_cvitem["syozai_kaisu3"]));
                int tikaflg1 = Conversions.ToInteger(Interaction.IIf(int.Parse(Conversions.ToString(hash_cvitem["syozai_tikaflg1"])) == 1, -1, 1));
                int tikaflg2 = Conversions.ToInteger(Interaction.IIf(int.Parse(Conversions.ToString(hash_cvitem["syozai_tikaflg2"])) == 1, -1, 1));
                int tikaflg3 = Conversions.ToInteger(Interaction.IIf(int.Parse(Conversions.ToString(hash_cvitem["syozai_tikaflg3"])) == 1, -1, 1));
                int kaisu1 = 0;
                int kaisu2 = 0;
                int kaisu3 = 0;
                int bkkai = 0;
                string tmp_bkkai = "";
                string tmp_hubi = "物件情報の階建てを超えているため移行できません。";     // ログ出力用

                // 階数調整
                if (!string.IsNullOrEmpty(tmp_kaisu1))
                {
                    kaisu1 = int.Parse(tmp_kaisu1) * tikaflg1;
                }
                if (!string.IsNullOrEmpty(tmp_kaisu2))
                {
                    kaisu2 = int.Parse(tmp_kaisu2) * tikaflg2;
                }
                if (!string.IsNullOrEmpty(tmp_kaisu3))
                {
                    kaisu3 = int.Parse(tmp_kaisu3) * tikaflg3;
                }

                // 物件No有無/無効データチェック(念の為)
                if (string.IsNullOrEmpty(bkno) || Typ.ToInt(bkno) == 0)
                {
                    return;
                }

                // 物件情報の階建てを取得
                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT ISNULL(kaidate,0) FROM bkdata_detail AS BKD ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata AS BK ON BKD.bk_guid = BK.bk_guid ";
                tmp_sql = tmp_sql + " WHERE bk_no = " + bkno;
                if (DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10) is not null)
                {
                    bkkai = int.Parse(Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10)));
                }

                // 階建てが設定されていない場合は処理を抜ける(未設定の場合は部屋所在階を自由に設定できるため)
                if (bkkai == 0)
                {
                    return;
                }

                // 所在階1チェック
                if (kaisu1 != 0)
                {
                    if (bkkai < kaisu1)
                    {
                        // ログ
                        string log_key = tblname + "-" + "syozai_kaisu1";
                        string errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + tmp_hubi + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                        hash_log.Add(log_key, errstr);
                        // 値の再格納
                        hash_cvitem["syozai_kaisu1"] = "";
                        hash_cvitem["syozai_tikaflg1"] = "0";
                    }
                }

                // 所在階2チェック
                if (kaisu2 != 0)
                {
                    if (bkkai < kaisu2)
                    {
                        // ログ
                        string log_key = tblname + "-" + "syozai_kaisu2";
                        string errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + tmp_hubi + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                        hash_log.Add(log_key, errstr);
                        // 値の再格納
                        hash_cvitem["syozai_kaisu2"] = "";
                        hash_cvitem["syozai_tikaflg2"] = "0";
                    }
                }

                // 所在階3チェック
                if (kaisu3 != 0)
                {
                    if (bkkai < kaisu3)
                    {
                        // ログ
                        string log_key = tblname + "-" + "syozai_kaisu3";
                        string errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + tmp_hubi + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                        hash_log.Add(log_key, errstr);
                        // 値の再格納
                        hash_cvitem["syozai_kaisu3"] = "";
                        hash_cvitem["syozai_tikaflg3"] = "0";
                    }
                }

            }

            // 2016.04.06 紐付データ取得処理を外出し -del sta
            // ''' <summary>
            // ''' 取引態様マスタ紐付情報取得
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
            // Dim relsheetname As String = "取引態様マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldtoritaiyoname As String = ""
            // Dim tmp_newtoritaiyono As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元取引態様名称"
            // tmp_oldtoritaiyoname = fldvalue.Trim
            // Case "賃貸革命10取引態様No"
            // tmp_newtoritaiyono = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newtoritaiyono <> "" Then
            // Hash_Rel_Toritaiyo.Add(tmp_oldtoritaiyoname, tmp_newtoritaiyono)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 紐付データ取得処理を外出し -del end

        }

    }

    #endregion

    #region 部屋所有者情報

    public class Hydata_syo_Repository
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

                var model_cvitem = new Model.Hydata_syo_Model();              // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = "物件部屋キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_syo";
                string fldnamegrp = "hy_guid,kn_no,sorule_guid,kasi1_ow_no,syo1_ow_no," + "kasi2_ow_no,syo2_ow_no,syo_startymd,syo_endymd,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
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
                            case "管理No":
                                {
                                    model_cvitem.Vari_Kn_no = fldvalue.Trim();
                                    break;
                                }
                            case "貸主１No":
                                {
                                    model_cvitem.Vari_Kasi1_ow_no = fldvalue.Trim();
                                    break;
                                }
                            case "所有者１No":
                                {
                                    model_cvitem.Vari_Syo1_ow_no = fldvalue.Trim();
                                    break;
                                }
                            case "貸主２No":
                                {
                                    model_cvitem.Vari_Kasi2_ow_no = fldvalue.Trim();
                                    break;
                                }
                            case "所有者２No":
                                {
                                    model_cvitem.Vari_Syo2_ow_no = fldvalue.Trim();
                                    break;
                                }
                            case "所有期間開始":
                                {
                                    model_cvitem.Vari_Syo_startymd = fldvalue.Trim();
                                    break;
                                }
                            case "所有期間終了":
                                {
                                    model_cvitem.Vari_Syo_endymd = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Sorule_guid = Guid.NewGuid().ToString();
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
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,kn_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋駐車場情報

    public class Hydata_parking_Repository
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

                var model_cvitem = new Model.Hydata_parking_Model();          // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = "物件部屋キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_parking";
                string fldnamegrp = "hy_guid,parking_kbn,parking_akisu,parking_status,parking_gakkbn," + "parking_gak,parking_gakzei,history,parking_tintaisu";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
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
                            case "駐車場区分":
                                {
                                    model_cvitem.Vari_Parking_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場の空き数(手動設定用)":
                                {
                                    model_cvitem.Vari_Parking_akisu = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場の空き有無(手動設定用)":
                                {
                                    model_cvitem.Vari_Parking_status = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場料金区分(手動設定用)":
                                {
                                    model_cvitem.Vari_Parking_gakkbn = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場料金(手動設定用)":
                                {
                                    model_cvitem.Vari_Parking_gak = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場料金税区分(手動設定用)":
                                {
                                    model_cvitem.Vari_Parking_gakzei = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場の賃貸可能数":
                                {
                                    model_cvitem.Vari_Parking_tintaisu = fldvalue.Trim();
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
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parking_kbn) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋特約情報

    public class Hydata_tokuyaku_Repository
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

                var model_cvitem = new Model.Hydata_tokuyaku_Model();         // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = "物件部屋キー情報";
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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_tokuyaku";
                string fldnamegrp = "hy_guid,tokuyaku_grpno,naiyo,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    string tmp_keyoya = tmp_keymain + "-" + tmp_keysub1;
                    string tmp_keydup = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;
                    model_cvitem.Vari_Tokuyaku_grpno = tmp_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;


                    // 特約カウント初期化
                    int cvitemcnt = 0;
                    int tmp_cvrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                    // データ有無チェック
                    string tmp_fldvalueumuchk = "";
                    for (int cntjj = 3, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
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
                            string log_key = "hydata_tokuyaku-hy_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keydup))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "hydata_tokuyaku-hy_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        if (keychkflg)
                        {

                            // 作業用変数作成
                            string tmp_tokuyaku = "";

                            // 移行値取得
                            for (int cntjj = 3, loopTo2 = readtbl.Columns.Count - 1; cntjj <= loopTo2; cntjj++)
                            {

                                // サブキー取得
                                string tmp_keysub3 = (cntjj - 2).ToString();

                                // ログ出力用データ格納(サブキーフィールド)
                                string fldname_keysub3 = readtbl.Columns[cntjj].ColumnName.Trim();

                                // 登録値取得
                                string fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();
                                if (!string.IsNullOrEmpty(fldvalue))
                                {
                                    tmp_tokuyaku = tmp_tokuyaku + CommonModule.LINE_BREAK + fldvalue;
                                }

                            }

                            // 成形
                            // 2016.04.26 メインの方へも反映させる修正 -chg sta
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
                                    // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                                    // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                                    // '2016.04.26 メインの方へも反映させる修正 -chg end
                                    hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,tokuyaku_grpno) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋鍵情報

    public class Hydata_kagi_Repository
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

                var model_cvitem = new Model.Hydata_kagi_Model();             // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = "物件部屋キー情報";

                int cnt_cvcnt = 0;                                    // 20160531 鍵情報移行処理の修正 専用鍵の件数取得用 -add

                // ************************
                // 作業準備
                // ************************

                // 20160613 鍵情報の取得処理修正 -chg sta
                // '10鍵タイトルマスタから鍵Noとタイトル名を紐付けたデータを取得
                // Call EtcMethod.Set_KagiTitleMst(sqlcnnv10, 1)
                // Call EtcMethod.Set_KagiTitleMst(sqlcnnv10, 2)

                // V7鍵Noと10鍵No(鍵区分を含む)を紐付けたハッシュテーブルを作成
                SetRelItemToObject.Set_RelData_KagiInfo(sqlcnnv10);
                // 20160613 鍵情報の取得処理修正 -chg end
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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_kagi";
                string fldnamegrp = "hy_guid,kagi_no,honsu,biko,history," + "hokan,gyshare";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
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
                            case "鍵No":      // 20160613 鍵情報の取得処理修正 -chg sta 鍵タイトル名→鍵No
                                {
                                    model_cvitem.Vari_Kagi_no = fldvalue.Trim();
                                    break;
                                }
                            case "鍵本数":
                                {
                                    model_cvitem.Vari_Honsu = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "保管場所":
                                {
                                    model_cvitem.Vari_Hokan = fldvalue.Trim();
                                    break;
                                }
                            case "業者間での情報共有":
                                {
                                    model_cvitem.Vari_Gyshare = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    // 20160927 汎用CV時の鍵情報の取得処理修正 -chg sta
                    // '20160613 鍵情報の取得処理修正 -chg sta
                    // ''20160531 鍵情報移行処理の修正 -chg sta
                    // ''共用鍵の場合に移行処理を行う
                    // ' ''フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    // ''tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                    // ' ''データチェック
                    // ''Dim skipflg As Boolean = False
                    // ''Dim hash_cvitem As New SafeDictionary<string, string>
                    // ''Dim hash_log As New SafeDictionary<string, string>
                    // ''skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                    // ' ''書込処理
                    // ''If Not skipflg Then

                    // ''    'キーをguidへ変換
                    // ''    '2016.04.26 メインの方へも反映させる修正 -chg sta
                    // ''    'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                    // ''    hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                    // ''    '2016.04.26 メインの方へも反映させる修正 -chg end

                    // ''    '挿入処理
                    // ''    Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                    // ''    '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                    // ''    If normalflg Then
                    // ''        list_chkduplicate.Add(fldvalue_key)
                    // ''        tmp_cvcnt = tmp_cvcnt + 1
                    // ''    End If

                    // ''End If

                    // ' ''-------------------
                    // ' ''ログ出力
                    // ' ''-------------------

                    // ' ''ログ出力メッセージ整形
                    // ''Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                    // ' ''ログ出力
                    // ''If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                    // ''    Dim tmp_cnt As Integer = 0
                    // ''    '挿入
                    // ''    For Each logvalue In sortlist_log
                    // ''        Dim tmp_sql_insert As String = logvalue.Value
                    // ''        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                    // ''    Next
                    // ''    '初期化
                    // ''    tmp_logcnt = 0
                    // ''    sortlist_log.Clear()
                    // ''End If

                    // 'If Hash_KagiTitleSenyo.Contains(.Vari_Kagi_no) Then

                    // '    '件数カウント
                    // '    cnt_cvcnt = cnt_cvcnt + 1

                    // '    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    // '    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                    // '    'データチェック
                    // '    Dim skipflg As Boolean = False
                    // '    Dim hash_cvitem As New SafeDictionary<string, string>
                    // '    Dim hash_log As New SafeDictionary<string, string>
                    // '    skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                    // '    '書込処理
                    // '    If Not skipflg Then

                    // '        'キーをguidへ変換
                    // '        '2016.04.26 メインの方へも反映させる修正 -chg sta
                    // '        'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                    // '        hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                    // '        '2016.04.26 メインの方へも反映させる修正 -chg end

                    // '        '挿入処理
                    // '        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                    // '        '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                    // '        If normalflg Then
                    // '            list_chkduplicate.Add(fldvalue_key)
                    // '            tmp_cvcnt = tmp_cvcnt + 1
                    // '        End If

                    // '    End If

                    // '    '-------------------
                    // '    'ログ出力
                    // '    '-------------------

                    // '    'ログ出力メッセージ整形
                    // '    Call LogSetting.Set_Log_Value_KomkErr(hash_log, tmp_hash, hash_cvitem, str_logkey, tblname, tmp_logcnt, sortlist_log)

                    // '    'ログ出力
                    // '    If sortlist_log.Count >= Log_OutputCnt Or (sortlist_log.Count <> 0 And cntii = rowcnt - 1) Then
                    // '        Dim tmp_cnt As Integer = 0
                    // '        '挿入
                    // '        For Each logvalue In sortlist_log
                    // '            Dim tmp_sql_insert As String = logvalue.Value
                    // '            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
                    // '        Next
                    // '        '初期化
                    // '        tmp_logcnt = 0
                    // '        sortlist_log.Clear()
                    // '    End If

                    // 'End If
                    // ''20160531 鍵情報移行処理の修正 -chg end

                    // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                    // 'データチェック
                    // Dim skipflg As Boolean = False
                    // Dim hash_cvitem As New SafeDictionary<string, string>
                    // Dim hash_log As New SafeDictionary<string, string>
                    // skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

                    // If hash_cvitem.Item("kagi_no") <> "" Then

                    // '件数カウント
                    // cnt_cvcnt = cnt_cvcnt + 1

                    // '書込処理
                    // If Not skipflg Then

                    // 'キーをguidへ変換
                    // '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                    // ''2016.04.26 メインの方へも反映させる修正 -chg sta
                    // ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                    // 'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                    // ''2016.04.26 メインの方へも反映させる修正 -chg end
                    // hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                    // '20160819 部屋の半角変換処理によるエラー修正 -chg end

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

                    // End If
                    // '20160613 鍵情報の取得処理修正 -chg end

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    switch (CommonModule.CNVNO)
                    {

                        case (int)CommonModule.ConvertTypes._汎用:
                            {

                                // 書込処理
                                if (!skipflg)
                                {

                                    // キーをguidへ変換
                                    hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];

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

                                break;
                            }

                    }
                    // 20160927 汎用CV時の鍵情報の取得処理修正 -chg end
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
                // 20160927 汎用CV時の鍵情報の取得処理修正 -chg sta
                // '20160531 鍵情報移行処理の修正 -chg sta
                // '共用鍵の場合に中間ファイルの件数をカウントする
                // '中間ファイル件数を取得
                // 'midrowcnt = rowcnt
                // midrowcnt = cnt_cvcnt
                // '20160531 鍵情報移行処理の修正 -chg end
                switch (CommonModule.CNVNO)
                {
                    case (int)CommonModule.ConvertTypes._汎用:
                        {
                            midrowcnt = rowcnt;
                            break;
                        }
                }
                // 20160927 汎用CV時の鍵情報の取得処理修正 -chg end
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,kagi_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋面積情報

    public class Hydata_othermenseki_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_othermenseki_Model();     // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = "物件部屋キー情報";

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

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_othermenseki";
                string fldnamegrp = "hy_guid,mensekitype,no,mensekikbn,name," + "meter,tubo,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);

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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

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
                            case "面積タイプ":
                                {
                                    model_cvitem.Vari_Mensekitype = fldvalue.Trim();
                                    break;
                                }
                            case "面積No":
                                {
                                    model_cvitem.Vari_No = fldvalue.Trim();
                                    break;
                                }
                            case "面積区分":
                                {
                                    model_cvitem.Vari_Mensekikbn = fldvalue.Trim();
                                    break;
                                }
                            case "区画名":
                                {
                                    model_cvitem.Vari_Name = fldvalue.Trim();
                                    break;
                                }
                            case "面積(m2)":
                                {
                                    model_cvitem.Vari_Meter = fldvalue.Trim();
                                    break;
                                }
                            case "坪数":
                                {
                                    model_cvitem.Vari_Tubo = fldvalue.Trim();
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
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,mensekitype)  + '-' + CONVERT(varchar,no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋間取内訳情報

    public class Hydata_madoriutiwake_Repository
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

                var model_cvitem = new Model.Hydata_madoriutiwake_Model();    // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列
                string viewname = "物件部屋キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_madoriutiwake";
                string fldnamegrp = "hy_guid,no,hykbn,jo,history," + "syozaikai";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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
                    model_cvitem.Vari_Hy_guid = fldvalue_key;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + fldvalue_keymain + "、" + fldname_keysub + " = " + fldvalue_keysub;

                    // 間取内訳カウント初期化
                    int cvitemcnt = 0;

                    // 作業用変数作成
                    string tmp_madori = "";
                    string[] tmp_madorino = null;
                    string[] tmp_madorikbn = null;
                    string[] tmp_madorijo = null;

                    // 格納
                    // 移行値取得
                    // 20160915 間取内訳情報の汎用/既存処理の分岐 -chg sta
                    // For cntjj = 1 To columncnt

                    // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                    // Dim fldvalue As String = ""
                    // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                    // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                    // End If

                    // If fldname = "既存間取内訳データ" Then

                    // '間取の連結文字列を取得
                    // tmp_madori = fldvalue.Trim

                    // '間取変換
                    // Call MadoriConv.SplitMadori(tmp_madori, tmp_madorino, tmp_madorikbn, tmp_madorijo)

                    // '変数格納→挿入処理
                    // For cntkk = 0 To UBound(tmp_madorino)

                    // .Vari_No = tmp_madorino(cntkk)
                    // .Vari_Hykbn = tmp_madorikbn(cntkk)
                    // .Vari_Jo = tmp_madorijo(cntkk)

                    // '固定値
                    // .Vari_History = DefHistory
                    // .Vari_Syozaikai = ""

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
                    // '20160819 部屋の半角変換処理によるエラー修正 -chg sta
                    // ''2016.04.26 メインの方へも反映させる修正 -chg sta
                    // ''hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                    // 'hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                    // ''2016.04.26 メインの方へも反映させる修正 -chg end
                    // hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                    // '20160819 部屋の半角変換処理によるエラー修正 -chg end

                    // '挿入処理
                    // Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                    // '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                    // If normalflg Then
                    // list_chkduplicate.Add(fldvalue_key)
                    // cvitemcnt = cvitemcnt + 1
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

                    // Next

                    // End If

                    // Next

                    // '--------------------------------------------------------------------
                    // '移行した間取内訳が1データ以上ある場合移行したレコードの数を更新する
                    // '--------------------------------------------------------------------
                    // If cvitemcnt > 0 Then
                    // tmp_cvcnt = tmp_cvcnt + 1
                    // End If

                    switch (CommonModule.CNVNO)
                    {

                        case (int)CommonModule.ConvertTypes._汎用:
                            {

                                // データ取得
                                for (int cntjj = 0, loopTo3 = readtbl.Columns.Count - 1; cntjj <= loopTo3; cntjj++)
                                {

                                    // 項目名取得
                                    fldname = readtbl.Columns[cntjj].ColumnName.Trim();

                                    // 登録値取得
                                    fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();

                                    switch (fldname ?? "")
                                    {
                                        case "間取り内訳No":
                                            {
                                                model_cvitem.Vari_No = fldvalue.Trim();
                                                break;
                                            }
                                        case "間取り内訳区分":
                                            {
                                                model_cvitem.Vari_Hykbn = fldvalue.Trim();
                                                break;
                                            }
                                        case "畳数":
                                            {
                                                model_cvitem.Vari_Jo = fldvalue.Trim();
                                                break;
                                            }
                                        case "所在階":
                                            {
                                                model_cvitem.Vari_Syozaikai = fldvalue.Trim();
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
                                    hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];

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

                                break;
                            }

                    }
                    // 20160915 間取内訳情報の汎用/既存処理の分岐 -chg end

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

                string tmp_sql = "";
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 間取内訳行Noの一括更新クエリ
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE hydata_madoriutiwake SET ";
                tmp_sql = tmp_sql + " 	no = MADORINO.連番 ";
                tmp_sql = tmp_sql + " FROM hydata_madoriutiwake AS HYMADORI ";
                tmp_sql = tmp_sql + " LEFT JOIN ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 hy_guid ";
                tmp_sql = tmp_sql + " 		,no ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY hy_guid ORDER BY no) AS 連番 ";
                tmp_sql = tmp_sql + " 	FROM hydata_madoriutiwake ";
                tmp_sql = tmp_sql + " ) AS MADORINO ";
                tmp_sql = tmp_sql + " ON  HYMADORI.hy_guid = MADORINO.hy_guid ";
                tmp_sql = tmp_sql + " AND HYMADORI.no = MADORINO.no ";
                tmp_sql = tmp_sql + " ; ";

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 部屋修繕維持管理連絡先情報

    public class Hydata_szeniji_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_szeniji_Model();          // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = "物件部屋キー情報";

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

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_szeniji";
                string fldnamegrp = "hy_guid,syuzenijikanri_no,syuzenijikanri_kasyo,syuzenijikanri_taisyokbn,syuzenijikanri_gyno," + "syuzenijikanri_simei,syuzenijikanri_address,syuzenijikanri_tel,syuzenijikanri_jisyano,syuzenijikanri_kasino," + "syuzenijikanri_syono,syuzenijikanri_simeiu";


                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);

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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

                    // 作業用変数 
                    string tmp_kasyoname = "";    // 20160829 箇所未登録データにデフォルト値を設定 -add

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
                            case "修繕及び維持管理No":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_no = fldvalue.Trim();
                                    break;
                                }
                            case "修繕及び維持管理の箇所":
                                {
                                    // 20160829 箇所未登録データにデフォルト値を設定 -chg sta
                                    // .Vari_Syuzenijikanri_kasyo = fldvalue.Trim
                                    tmp_kasyoname = fldvalue.Trim();
                                    break;
                                }
                            // 20160829 箇所未登録データにデフォルト値を設定 -chg end
                            case "対象区分":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_taisyokbn = fldvalue.Trim();
                                    break;
                                }
                            case "業者no":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_gyno = fldvalue.Trim();
                                    break;
                                }
                            case "氏名(商号または名称)(SJIS)":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_simei = fldvalue.Trim();
                                    break;
                                }
                            case "住所(主たる事務所の所在地)":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_address = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先電話番号":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_tel = fldvalue.Trim();
                                    break;
                                }
                            case "自社no":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_jisyano = fldvalue.Trim();
                                    break;
                                }
                            case "貸主No":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_kasino = fldvalue.Trim();
                                    break;
                                }
                            case "所有者No":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_syono = fldvalue.Trim();
                                    break;
                                }
                            case "氏名(商号または名称)":
                                {
                                    model_cvitem.Vari_Syuzenijikanri_simeiu = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 20160829 箇所未登録データにデフォルト値を設定 -add
                    model_cvitem.Vari_Syuzenijikanri_kasyo = tmp_kasyoname + model_cvitem.Vari_Syuzenijikanri_no;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 20160829 箇所未登録データにデフォルト値を設定 -add sta
                    if (string.IsNullOrEmpty(tmp_kasyoname))
                    {
                        tmp_hash["syuzenijikanri_kasyo"] = "";
                    }
                    // 20160829 箇所未登録データにデフォルト値を設定 -add end

                    // 書込処理
                    if (!skipflg)
                    {

                        // キーをguidへ変換
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,syuzenijikanri_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋設備情報

    public class Hydata_setubilst_Repository
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
                var list_chkduplicateguid = new List<string>();                // 重複チェック用リスト (移行後の重複チェック用)   '2016.04.26 メインの方へも反映させる修正
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_setubilst_Model();        // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = "物件部屋キー情報";

                int cnt_cvcnt = 0;                                    // 20160603 ユーザーデータ検証による修正 -add

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 2016.04.06 部屋設備情報の取得処理修正 -chg sta
                // Call Me.Get_RelData()
                SetRelItemToObject.Set_RelData_Setubi_V7To10No();
                EtcMethod.Set_SetubiMst(sqlcnnv10);
                EtcMethod.Set_SetubiMst_UserMake(sqlcnnv10); // 20160829 設備の新規挿入処理を追加 -add
                Set_RelData_Setubi_V7To10Guid();
                // 2016.04.06 部屋設備情報の取得処理修正 -chg end

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_setubilst";
                string fldnamegrp = "hy_guid,komok_guid,override_kbn,komok_name,disp1name," + "disp2name,disp3name,disp1iconguid,disp2iconguid,disp3iconguid," + "history";


                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicateguid);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);

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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

                    // 作業用変数
                    string tmp_setubikomkname = "";
                    string tmp_setubinaiyo = "";

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
                            case "設備項目名":
                                {
                                    tmp_setubikomkname = fldvalue.Trim();
                                    break;
                                }
                            case "設備内容":
                                {
                                    tmp_setubinaiyo = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 項目guidへ一時格納
                    model_cvitem.Vari_Komok_guid = tmp_setubikomkname + "-" + tmp_setubinaiyo;

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

                        // 2016.04.26 メインの方へも反映させる修正 -chg sta
                        // guidへ変換した値での重複チェックを行う
                        // ↓↓↓旧srcコメントアウト↓↓↓
                        // 'キーをguidへ変換
                        // hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))

                        // '挿入処理
                        // Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                        // '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        // If normalflg Then
                        // list_chkduplicate.Add(fldvalue_key)
                        // tmp_cvcnt = tmp_cvcnt + 1
                        // End If
                        // ↑↑↑旧srcコメントアウト↑↑↑

                        // キーをguidへ変換
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

                        // guidでの重複チェック
                        string tmp_hyguid = Conversions.ToString(hash_cvitem["hy_guid"]);
                        string tmp_komkguid = Conversions.ToString(hash_cvitem["komok_guid"]);

                        // 20160829 エレベーター移行対応 -chg sta
                        // '20160603 ユーザーデータ検証による修正 -add sta
                        // '紐付けられた値のみ移行件数をカウントする
                        // If tmp_komkguid <> "" Then
                        // cnt_cvcnt = cnt_cvcnt + 1
                        // End If
                        // '20160603 ユーザーデータ検証による修正 -add end

                        // If list_chkduplicateguid.Contains(tmp_hyguid & "-" & tmp_komkguid) = False Then

                        // '挿入処理
                        // Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

                        // '挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        // If normalflg Then
                        // list_chkduplicate.Add(fldvalue_key)
                        // list_chkduplicateguid.Add(tmp_hyguid & "-" & tmp_komkguid)
                        // tmp_cvcnt = tmp_cvcnt + 1
                        // End If

                        // Else

                        // 'ログ出力
                        // hash_log.Clear()    '重複エラーを優先する
                        // Dim tmp_tblfldvalue As String = tblname & "-" & "komok_guid"
                        // hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)

                        // End If
                        // '2016.04.26 メインの方へも反映させる修正 -chg end

                        // エレーベータの値をセット
                        var list_elv = new List<string>() { "1", "2", "3" };
                        if (list_elv.Contains(tmp_komkguid))
                        {
                            // 物件情報のエレベーターを更新する
                            Set_Bkdatadetail_ElvFlg(sqlcnnv10, tmp_hyguid, tmp_komkguid);
                        }
                        else
                        {
                            // 紐付けられた値のみ移行件数をカウントする
                            if (!string.IsNullOrEmpty(tmp_komkguid))
                            {
                                cnt_cvcnt = cnt_cvcnt + 1;
                            }

                            if (list_chkduplicateguid.Contains(tmp_hyguid + "-" + tmp_komkguid) == false)
                            {

                                // 挿入処理
                                CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                                // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                if (normalflg)
                                {
                                    list_chkduplicate.Add(fldvalue_key);
                                    list_chkduplicateguid.Add(tmp_hyguid + "-" + tmp_komkguid);
                                    tmp_cvcnt = tmp_cvcnt + 1;
                                }
                            }

                            else
                            {

                                // ログ出力
                                hash_log.Clear();    // 重複エラーを優先する
                                string tmp_tblfldvalue = tblname + "-" + "komok_guid";
                                hash_log.Add(tmp_tblfldvalue, CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP);

                            }

                        }
                        // 20160829 エレベーター移行対応 -chg end
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

                // 20160603 ユーザーデータ検証による修正 -chg sta
                // '中間ファイル件数を取得
                // midrowcnt = rowcnt
                // 移行件数を取得
                midrowcnt = cnt_cvcnt;
                // 20160603 ユーザーデータ検証による修正 -chg end

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
                // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parking_kbn) FROM " & tblname
                // tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
                // Dim flg As Boolean = True

                // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

                // 2016.04.26 メインの方へも反映させる修正 -add sta
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	CONVERT(VARCHAR(200),hy_guid) + '-' + CONVERT(VARCHAR(200),komok_guid) ";
                tmp_sql = tmp_sql + " FROM hydata_setubilst ";
                bool flg = true;
                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);
                // 2016.04.26 メインの方へも反映させる修正 -add end


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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            // 2016.04.06 部屋設備情報の取得処理修正 -del sta
            // ''' <summary>
            // ''' 設備紐付情報取得
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
            // Dim relsheetname As String = "設備マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_setubigrpname As String = ""
            // Dim tmp_setubiname As String = ""
            // Dim tmp_setubiguid As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元設備グループ名称"
            // tmp_setubigrpname = fldvalue.Trim
            // Case "移行元設備名称"
            // tmp_setubiname = fldvalue.Trim
            // Case "賃貸革命10項目GUID"
            // tmp_setubiguid = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_setubiguid <> "" Then
            // Hash_Rel_Setubi.Add(tmp_setubigrpname & "-" & tmp_setubiname, tmp_setubiguid)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 部屋設備情報の取得処理修正 -del end

            /// <summary>
            /// V7設備データと10設備項目guidを紐付 2016.04.06 部屋設備情報の取得処理修正 -add sta
            /// </summary>
            /// <remarks></remarks>
            public void Set_RelData_Setubi_V7To10Guid()
            {

                // 初期化
                CommonModule.Hash_Rel_Setubi.Clear();

                foreach (var setubiitem in CommonModule.Hash_SetubiMid)
                {

                    string tmp_V7setubi = Conversions.ToString(setubiitem.Key);
                    string tmp_10setubino = Conversions.ToString(setubiitem.Value);

                    if (CommonModule.Hash_SetubiMst.ContainsKey(tmp_10setubino))
                    {
                        string tmp_10setubiguid = Conversions.ToString(CommonModule.Hash_SetubiMst[tmp_10setubino]);
                        CommonModule.Hash_Rel_Setubi.Add(tmp_V7setubi, tmp_10setubiguid);
                    }
                    else if (string.IsNullOrEmpty(tmp_10setubino))     // 20160829 設備の新規挿入処理を追加 -add
                    {
                        if (CommonModule.Hash_SetubiMst_UserMake.ContainsKey(tmp_V7setubi))
                        {
                            string tmp_10setubiguid_usermake = Conversions.ToString(CommonModule.Hash_SetubiMst_UserMake[tmp_V7setubi]);
                            CommonModule.Hash_Rel_Setubi.Add(tmp_V7setubi, tmp_10setubiguid_usermake);
                        }
                    }

                }

            }
            // 2016.04.06 部屋設備情報の取得処理修正 -add end

            /// <summary>
            /// 物件詳細情報のエレベーターフラグを更新 '20160829 エレベーター移行対応
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="hyguid"></param>
            /// <remarks></remarks>
            public void Set_Bkdatadetail_ElvFlg(SqlConnection sqlcnnv10, string hyguid, string elvvalue)
            {

                // 部屋guidが存在しない場合は処理を抜ける
                if (string.IsNullOrEmpty(hyguid))
                {
                    return;
                }

                // bkguidを取得
                string tmp_sql = " SELECT CONVERT(VARCHAR(MAX),bk_guid) FROM hydata WHERE hy_guid = '" + hyguid + "'";
                string tmp_bkguid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));

                // 更新クエリを作成
                string tmp_updatefld = "";
                switch (elvvalue ?? "")
                {
                    case "1":
                        {
                            tmp_updatefld = "elevator_flg";
                            break;
                        }
                    case "2":
                        {
                            tmp_updatefld = "homeelevator_flg";
                            break;
                        }
                    case "3":
                        {
                            tmp_updatefld = "emergencyelevator_flg";
                            break;
                        }
                }

                // 更新クエリ作成
                string tmp_sqlupdate = "";
                tmp_sqlupdate = tmp_sqlupdate + " UPDATE bkdata_detail SET " + tmp_updatefld + " = 1 ";
                tmp_sqlupdate = tmp_sqlupdate + " WHERE bk_guid = '" + tmp_bkguid + "'";

                // 実行
                int tmpcnt = 0;
                bool flg = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqlupdate, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋入金項目情報

    public class Hydata_nkin_Repository
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

                var model_cvitem = new Model.Hydata_nkin_Model();                 // 移行値格納用モデル初期化
                int keycol_main = 1;                                      // メインキー列
                int keycol_sub1 = 2;                                      // サブ1キー列
                int keycol_sub2 = 3;                                      // サブ2キー列
                int keycol_sub3 = 4;                                      // サブ3キー列
                string viewname = "物件部屋キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_nkin";
                string fldnamegrp = "hy_guid,tuki_kbn,nkin_no,nkin_recno,nkin_kbn," + "sq_gak,sq_zeikbn,calc_kbn,calc_nkinno,calc_monthcnt," + "sqsaki_no,sq_mmkbn,sqstart_ymd,sq_ptn,sq_interval," + "sq_nen,sq_tuki,biko,history";



                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    // 20161125 ログに入金項目名を追加 -chg sta
                    // Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub1 & " = " & tmp_keysub1 & "、" & fldname_keysub2 & " = " & tmp_keysub2
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2 + "、" + fldname_keysub3 + " = " + tmp_keysub3;
                    // 20161125 ログに入金項目名を追加 -chg end

                    // 2016.04.06 入金項目読込処理の修正 -add
                    // 作業用変数
                    string tmp_nkinname = "";

                    // 20160530 ログ修正 -add
                    string tmp_nkinname_kijyun = "";


                    // 2016.04.26 メインの方へも反映させる修正 -add
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
                            case "月区分":
                                {
                                    model_cvitem.Vari_Tuki_kbn = fldvalue;
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
                            case "入金項目行No":
                                {
                                    model_cvitem.Vari_Nkin_recno = fldvalue;
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
                            // 2016.04.26 メインの方へも反映させる修正 -del sta
                            // 必要な情報をまとめて格納するためここではコメントアウト
                            // Case "請求パターン"
                            // .Vari_Sq_ptn = fldvalue
                            // 2016.04.26 メインの方へも反映させる修正 -del end
                            case "固定公共料金で使用":
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
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 2016.04.06 入金項目読込処理の修正 -add sta
                    // 紐付用に成形
                    string localGet_Nkinruiname() { string argtukikbn = model_cvitem.Vari_Nkin_kbn; var ret = EtcMethod.Get_Nkinruiname(ref argtukikbn); model_cvitem.Vari_Nkin_kbn = argtukikbn; return ret; }

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
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parking_kbn) FROM " & tblname
                // tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            // 2016.04.06 入金項目読込処理の修正 -del sta
            /// <summary>
            /// 入金項目紐付情報取得
            /// </summary>
            /// <remarks></remarks>
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

    #region 部屋変動費各戸メーター情報

    public class Hydata_hendo_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_hendo_Model();            // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                string viewname = "物件部屋キー情報";

                // ************************
                // 作業準備
                // ************************


                // 入金項目の重複チェックは特殊なので考慮すること






                // 紐付けデータ取得
                // 2016.04.06 入金項目読込処理の修正 -chg sta
                // Call Me.Get_RelData_Nkinkomk()
                // Call Me.Get_RelData_Hendometer()
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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_hendo";
                string fldnamegrp = "hy_guid,hyhendo_guid,rec_no,hendo_kbn,meter_name," + "nkin_no,biko,hendorule_no,hendorule_biko,history," + "sq_mmkbn,useflg";


                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);

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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

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
                            case "行No":
                                {
                                    model_cvitem.Vari_Rec_no = fldvalue;
                                    break;
                                }
                            case "メーター分類":
                                {
                                    break;
                                }
                            // .Vari_Hendo_kbn = fldvalue.Trim
                            case "メーター名":
                                {
                                    model_cvitem.Vari_Meter_name = fldvalue;
                                    break;
                                }
                            case "入金項目名":      // 2016.04.01 他の移行項目との統一のためヘッダー名を変更 変動費入金項目 → 入金項目名
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
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            case "変動費請求ルールNo":
                                {
                                    model_cvitem.Vari_Hendorule_no = fldvalue;
                                    break;
                                }
                            case "変動ルール備考":
                                {
                                    model_cvitem.Vari_Hendorule_biko = fldvalue;
                                    break;
                                }
                            case "請求月区分":
                                {
                                    model_cvitem.Vari_Sq_mmkbn = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Hyhendo_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Useflg = 1.ToString();        // 変動費検針データに必要なため1をセット

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
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                // tmp_sql = tmp_sql & " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parking_kbn) FROM " & tblname
                // tmp_sql = tmp_sql & " LEFT JOIN hydata ON " & tblname & ".hy_guid = hydata.hy_guid "
                // tmp_sql = tmp_sql & " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid "
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

            /// <summary>
            /// 入金項目紐付情報取得
            /// </summary>
            /// <remarks></remarks>
            public void Get_RelData_Nkinkomk()
            {

                Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
                var startrow = default(int);                                         // 書込開始行
                var columncnt = default(int);                                        // 列数
                var maxrowcnt = default(int);                                        // 既存データの行数
                var rowcnt = default(int);                                           // 書込行数
                string relsheetname = "入金項目マスタ";
                bool rtn = true;

                var excelfile = new ExcelFileManager();                // Excelファイル操作用

                // 既存データ有無確認(入金項目は他の項目でも参照するため)
                if (CommonModule.Hash_Rel_Nkinkomk.Count != 0)
                {
                    return;
                }

                // Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

                for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
                {

                    // 作業用変数作成
                    string tmp_oldnkinname = "";
                    string tmp_newnkinno = "";
                    string tmp_nkinkbn = "";
                    string tmp_hendometerno = "";

                    // 紐付設定値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        // ヘッダー格納
                        string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                        // 移行値格納
                        string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                        switch (fldname ?? "")
                        {
                            case "移行元入金項目名称":
                                {
                                    tmp_oldnkinname = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸革命10入金項目No":
                                {
                                    tmp_newnkinno = fldvalue.Trim();
                                    break;
                                }
                            // 2016.04.06 入金項目読込処理の修正 -add sta
                            case "移行元入金項目区分":
                                {
                                    tmp_nkinkbn = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸革命10変動費メーター分類No":
                                {
                                    tmp_hendometerno = fldvalue.Trim();
                                    break;
                                }
                                // 2016.04.06 入金項目読込処理の修正 -add end
                        }

                    }

                    // 入金項目をハッシュテーブル格納
                    if (!string.IsNullOrEmpty(tmp_newnkinno))
                    {
                        CommonModule.Hash_Rel_Nkinkomk.Add(tmp_oldnkinname, tmp_newnkinno);
                    }

                    // 2016.04.06 入金項目読込処理の修正 -add sta
                    // 変動費メーター分類をハッシュテーブル格納
                    if (tmp_nkinkbn == "5.随時変動" & !string.IsNullOrEmpty(tmp_hendometerno))
                    {
                        CommonModule.Hash_Rel_Hendometer.Add(tmp_oldnkinname, tmp_hendometerno);
                    }
                    // 2016.04.06 入金項目読込処理の修正 -add end

                }

                // Excelファイル終了設定
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            }

            // 2016.04.06 入金項目読込処理の修正 -del sta
            /// <summary>
            /// 各戸メーター分類紐付情報取得
            /// </summary>
            /// <remarks></remarks>
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
                tmp_sql = tmp_sql + " 	SELECT bk_guid FROM hydata AS HY ";
                tmp_sql = tmp_sql + " 	WHERE hy_guid IN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT hy_guid FROM hydata_hendo AS HYH ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " ); ";

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 部屋メモ情報

    public class Hydata_memo_Repository
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

                var model_cvitem = new Model.Hydata_memo_Model();             // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列
                string viewname = "物件部屋キー情報";
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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_memo";
                string fldnamegrp = "hy_guid,memo_no,memo,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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
                string fldname_keysub1 = readtbl.Columns[keycol_sub - 1].ColumnName.Trim();

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
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub - 1]).Trim();
                    string tmp_keytotal = tmp_keymain + "-" + tmp_keysub1;
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // 備考カウント初期化
                    int cvitemcnt = 0;
                    int tmp_cvrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

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

                        var hash_keylog = new SafeDictionary<string, string>();
                        string log_keyout = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1;
                        // 親データ有無チェック
                        if (list_basekeydata.Contains(tmp_keytotal) == false)
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "hydata_memo-hy_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keytotal))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "hydata_memo-hy_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        if (keychkflg)
                        {

                            // 移行値取得
                            for (int cntjj = 2, loopTo2 = readtbl.Columns.Count - 1; cntjj <= loopTo2; cntjj++)
                            {

                                // サブキー取得
                                string tmp_keysub2 = (cntjj - 1).ToString();

                                // 全キー取得
                                string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;

                                // ログ出力用データ格納(サブキーフィールド)
                                string fldname_keysub2 = readtbl.Columns[cntjj].ColumnName.Trim();
                                string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2;

                                // 登録値取得
                                string fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();
                                model_cvitem.Vari_Memo = fldvalue;

                                // 固定値
                                model_cvitem.Vari_Memo_no = (cntjj - 1).ToString();
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
                                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
                                        // 20160819 部屋の半角変換処理によるエラー修正 -chg end

                                        // 挿入処理
                                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                        if (normalflg)
                                        {
                                            if (list_chkduplicate.Contains(tmp_keytotal) == false)
                                            {
                                                list_chkduplicate.Add(tmp_keytotal);
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,memo_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋共通セールスポイント情報

    public class Hydata_commonsalespointparts_Repository
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

                var model_cvitem = new Model.Hydata_commonsalespointparts_Model();    // 移行値格納用モデル初期化
                int keycol_main = 1;                                          // メインキー列
                int keycol_sub1 = 2;                                          // サブ1キー列
                int keycol_sub2 = 3;                                          // サブ2キー列
                string viewname = "物件部屋キー情報";

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
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_commonsalespointparts";
                string fldnamegrp = "hy_guid,parts_no,salespoint_parts";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol_main - 1]).Trim();
                    string tmp_keysub1 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub1 - 1]).Trim();
                    string tmp_keysub2 = Typ.ToStr(readtbl.Rows[cntii][keycol_sub2 - 1]).Trim();
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub1 + "-" + tmp_keysub2;
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
                            case "パーツNO":
                                {
                                    model_cvitem.Vari_Parts_no = fldvalue.Trim();
                                    break;
                                }
                            case "セールスポイントパーツ":
                                {
                                    model_cvitem.Vari_Salespoint_parts = fldvalue.Trim();
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,parts_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋契約解約確認事項情報

    public class Hydata_confirm_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Hydata_confirm_Model();          // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブ1キー列
                int keycol_sub2 = 3;                                  // サブ2キー列
                int keycol_sub3 = 4;                                  // サブ3キー列
                string viewname = "物件部屋キー情報";

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

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_confirm";
                string fldnamegrp = "hy_guid,confirm_kbn,confirm_no,naiyo";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                // ---------------
                // ヘッダー処理
                // ---------------
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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

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
                            case "確認事項区分":
                                {
                                    model_cvitem.Vari_Confirm_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "確認事項No":
                                {
                                    model_cvitem.Vari_Confirm_no = fldvalue.Trim();
                                    break;
                                }
                            case "内容":
                                {
                                    model_cvitem.Vari_Naiyo = fldvalue.Trim();
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,confirm_kbn)  + '-' + CONVERT(varchar,confirm_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋権利情報

    public class Hydata_kenri_Repository
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

                Microsoft.Office.Interop.Excel.Application appli = null;                                // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                                   // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                                 // Excelオブジェクト
                var startrow = default(int);                                                 // 書込開始行
                var columncnt = default(int);                                                // 列数
                var maxrowcnt = default(int);                                                // 既存データの行数
                var rowcnt = default(int);                                                   // 書込行数
                var tmp_cvcnt = default(int);                                                // 移行件数格納
                var tmp_condcnt = default(int);                                              // 調整件数格納

                var excelfile = new ExcelFileManager();                        // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                        // プログレスバー設定用
                var obj_com = new CommonRepository();                          // 共通処理用
                var list_chkduplicate = new List<string>();                            // 重複チェック用リスト
                var list_basekeydata = new List<string>();                             // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();                  // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                           // ログ出力時のソート用
                var hash_guid = new SafeDictionary<string, string>();                                          // guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                           // 作業用ハッシュテーブル
                var normalflg = default(bool);                                                // INSERT正常終了フラグ
                bool rtn = true;                                               // 戻り値

                var model_cvitem = new Model.Hydata_kenri_Model();                    // 移行値格納用モデル初期化
                int keycol_main = 1;                                          // メインキー列
                int keycol_sub1 = 2;                                          // サブ1キー列
                int keycol_sub2 = 3;                                          // サブ2キー列
                string viewname = "物件部屋キー情報";

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

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_kenri";
                string fldnamegrp = "hy_guid,kenri_no,other_kenrirui,other_kenribiko,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);

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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

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
                            case "権利情報No":
                                {
                                    model_cvitem.Vari_Kenri_no = fldvalue.Trim();
                                    break;
                                }
                            case "所有権以外の権利の種類":
                                {
                                    model_cvitem.Vari_Other_kenrirui = fldvalue.Trim();
                                    break;
                                }
                            case "所有権以外の権利備考":
                                {
                                    model_cvitem.Vari_Other_kenribiko = fldvalue.Trim();
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
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,kenri_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋参照ファイル情報

    public class Hydata_relfile_Repository
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

                Microsoft.Office.Interop.Excel.Application appli = null;                                // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                                   // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                                 // Excelオブジェクト
                var startrow = default(int);                                                 // 書込開始行
                var columncnt = default(int);                                                // 列数
                var maxrowcnt = default(int);                                                // 既存データの行数
                var rowcnt = default(int);                                                   // 書込行数
                var tmp_cvcnt = default(int);                                                // 移行件数格納
                var tmp_condcnt = default(int);                                              // 調整件数格納

                var excelfile = new ExcelFileManager();                        // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                        // プログレスバー設定用
                var obj_com = new CommonRepository();                          // 共通処理用
                var list_chkduplicate = new List<string>();                            // 重複チェック用リスト
                var list_basekeydata = new List<string>();                             // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();                  // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                           // ログ出力時のソート用
                var hash_guid = new SafeDictionary<string, string>();                                          // guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                           // 作業用ハッシュテーブル
                var normalflg = default(bool);                                                // INSERT正常終了フラグ
                bool rtn = true;                                               // 戻り値

                var model_cvitem = new Model.Hydata_relfile_Model();                  // 移行値格納用モデル初期化
                int keycol_main = 1;                                          // メインキー列
                int keycol_sub1 = 2;                                          // サブ1キー列
                int keycol_sub2 = 3;                                          // サブ2キー列
                string viewname = "物件部屋キー情報";

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

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_relfile";
                string fldnamegrp = "hy_guid,file_no,fullpath,addtime,biko," + "history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);

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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

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
                            case "ファイルNo":
                                {
                                    model_cvitem.Vari_File_no = fldvalue.Trim();
                                    break;
                                }
                            case "フルパス":
                                {
                                    model_cvitem.Vari_Fullpath = fldvalue.Trim();
                                    break;
                                }
                            case "追加時間":
                                {
                                    model_cvitem.Vari_Addtime = fldvalue.Trim();
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
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,file_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion

    #region 部屋原状回復目安単価情報

    public class Hydata_szen_Repository
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

                Microsoft.Office.Interop.Excel.Application appli = null;                                // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Workbook wbook = null;                                   // Excelオブジェクト
                Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                                 // Excelオブジェクト
                var startrow = default(int);                                                 // 書込開始行
                var columncnt = default(int);                                                // 列数
                var maxrowcnt = default(int);                                                // 既存データの行数
                var rowcnt = default(int);                                                   // 書込行数
                var tmp_cvcnt = default(int);                                                // 移行件数格納
                var tmp_condcnt = default(int);                                              // 調整件数格納

                var excelfile = new ExcelFileManager();                        // Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                        // プログレスバー設定用
                var obj_com = new CommonRepository();                          // 共通処理用
                var list_chkduplicate = new List<string>();                            // 重複チェック用リスト
                var list_basekeydata = new List<string>();                             // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();                  // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                           // ログ出力時のソート用
                var hash_guid = new SafeDictionary<string, string>();                                          // guid格納用ハッシュテーブル

                var tmp_hash = new SafeDictionary<string, string>();                                           // 作業用ハッシュテーブル
                var normalflg = default(bool);                                                // INSERT正常終了フラグ
                bool rtn = true;                                               // 戻り値

                var model_cvitem = new Model.Hydata_szen_Model();                     // 移行値格納用モデル初期化
                int keycol_main = 1;                                          // メインキー列
                int keycol_sub1 = 2;                                          // サブ1キー列
                int keycol_sub2 = 3;                                          // サブ2キー列
                string viewname = "物件部屋キー情報";

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

                // キー取得用のVIEWを作成
                Create_View_KeyAndGuid(sqlcnnv10);

                // テーブル名/フィールド名セット
                string viewname_base = CommonModule.PRE_VIEW_NAME + viewname;
                string tblname = "hydata_szen";
                string fldnamegrp = "hy_guid,szen_no,szen_name,ryo,unit_name," + "tanka";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // guidを取得
                Get_Guid(sqlcnnv10, viewname_base, ref hash_guid);

                // 親マスタ取得
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // Call Get_BaseKey(sqlcnnv10, viewname_base, list_basekeydata)
                EtcMethod.Set_HashKeyToList(hash_guid, ref list_basekeydata);
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

                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue[startrow - 1, keycol_main]);
                string fldname_keysub1 = Conversions.ToString(headervalue[startrow - 1, keycol_sub1]);
                string fldname_keysub2 = Conversions.ToString(headervalue[startrow - 1, keycol_sub2]);

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
                    model_cvitem.Vari_Hy_guid = tmp_keymain + "-" + tmp_keysub1;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub1 + " = " + tmp_keysub1 + "、" + fldname_keysub2 + " = " + tmp_keysub2;

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
                            case "修繕No":
                                {
                                    model_cvitem.Vari_Szen_no = fldvalue.Trim();
                                    break;
                                }
                            case "修繕項目名":
                                {
                                    model_cvitem.Vari_Szen_name = fldvalue.Trim();
                                    break;
                                }
                            case "数量":
                                {
                                    model_cvitem.Vari_Ryo = fldvalue.Trim();
                                    break;
                                }
                            case "単位":
                                {
                                    model_cvitem.Vari_Unit_name = fldvalue.Trim();
                                    break;
                                }
                            case "単価":
                                {
                                    model_cvitem.Vari_Tanka = fldvalue.Trim();
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
                        // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
                        // '2016.04.26 メインの方へも反映させる修正 -chg sta
                        // 'hash_cvitem("hy_guid") = hash_guid.Item(hash_cvitem.Item("hy_guid"))
                        // hash_cvitem("hy_guid") = hash_guid.Item(StrConv(hash_cvitem.Item("hy_guid"), VbStrConv.Narrow))
                        // '2016.04.26 メインの方へも反映させる修正 -chg end
                        hash_cvitem["hy_guid"] = hash_guid[hash_cvitem["hy_guid"]];
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
                tmp_sql = tmp_sql + " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,hy_no) + '-' + CONVERT(varchar,szen_no) FROM " + tblname;
                tmp_sql = tmp_sql + " LEFT JOIN hydata ON " + tblname + ".hy_guid = hydata.hy_guid ";
                tmp_sql = tmp_sql + " LEFT JOIN bkdata ON hydata.bk_guid = bkdata.bk_guid ";
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
            public void Create_View_KeyAndGuid(SqlConnection sqlcnnv10)
            {

                string viewname = "物件部屋キー情報";
                int tmpcnt = 0;

                // VIEW初期化
                // 20160531 部屋仮VIEWの初期化処理修正 -chg sta
                // Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(viewname, False)
                string tmp_sql_drop = DBQuery.Qry_DropInfo(CommonModule.PRE_VIEW_NAME + viewname, false);
                // 20160531 部屋仮VIEWの初期化処理修正 -chg end
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmpcnt);

                // VIEW作成
                string tmp_sql_create = "";
                tmp_sql_create = tmp_sql_create + CommonModule.PRE_VIEW_QRY + CommonModule.PRE_VIEW_NAME + viewname + CommonModule.POST_VIEW_QRY;
                tmp_sql_create = tmp_sql_create + " SELECT ";
                tmp_sql_create = tmp_sql_create + " 	 CONVERT(varchar,BK.bk_no) + '-' + CONVERT(varchar,HY.hy_no) AS キー ";
                tmp_sql_create = tmp_sql_create + " 	,HY.hy_guid ";
                tmp_sql_create = tmp_sql_create + " FROM hydata AS HY ";
                tmp_sql_create = tmp_sql_create + " LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmpcnt);

            }

        }

    }

    #endregion



    #region 部屋所有者情報

    #endregion

    #region 部屋画像情報

    #endregion

    #region 部屋業務期間情報

    #endregion

    #region 部屋BtoBグループ情報

    #endregion

    #region 部屋KeyValue情報

    #endregion

    #region 部屋広告補足情報(作り直し中)

    #endregion

    #region 部屋hydata_kys_atenaprint

    #endregion

    #region 部屋ポータル最終送信履歴情報

    #endregion

    #region 部屋セールスポイント情報

    #endregion

    #region 部屋連動サイト毎地図表示情報

    #endregion

    #region 部屋hydata_panorama情報

    #endregion

    #region 部屋支店情報

    #endregion

    #region 部屋hydata_sosin情報

    #endregion

    #region 部屋hydata_youtube情報

    #endregion

    #region 部屋親メーター関連情報

    #endregion

}