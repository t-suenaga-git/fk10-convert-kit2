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

    #region 物件基本情報

    public class Bkdata_Repository
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
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Bkdata_Model();                  // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

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
                // テーブル名/フィールド名セット
                string tblname = "bkdata";
                // 20160519 EXEUpdateに伴う修正 物件基本情報 -add (最後尾にlastupdateを追加)
                string fldnamegrp = "bk_guid,bk_no,bk_deleteflg,delete_guid,delete_day," + "delete_cnt,bk_name,bk_deletename,bk_kana,tatemono_sikibetu," + "post_code,addr_kenno,addr_sino,addr_cyo,addr_cyome," + "addr_cyomeptn,addr_banti,addr_etc,history,rowid," + "bk_namesjis,bk_gaibuno,delete_cause,lastupdate";




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

                // キーヘッダー名取得
                string fldname_key = readtbl.Columns[keycol - 1].ColumnName.Trim();

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
                    // 20170511 汎用_物件住所の移行不備対応 -del sta
                    // Dim tmp_addresscyotiiki As String = ""
                    // Dim tmp_addresscyobanti1 As String = ""
                    // Dim tmp_addresscyobanti2 As String = ""
                    // Dim tmp_addressother As String = ""
                    // 20170511 汎用_物件住所の移行不備対応 -del end

                    // キー値取得
                    string fldvalue_key = Typ.ToStr(readtbl.Rows[cntii][keycol - 1]).Trim();

                    // ログ出力用
                    string str_logkey = fldname_key + " = " + fldvalue_key;

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
                                    model_cvitem.Vari_Bk_no = fldvalue.Trim();
                                    break;
                                }
                            // Case "物件削除フラグ"
                            // .Vari_Bk_deleteflg = fldvalue.Trim
                            // Case "削除日"
                            // .Vari_Delete_day = fldvalue.Trim
                            // Case "削除復旧回数"
                            // .Vari_Delete_cnt = fldvalue.Trim
                            case "物件名":
                                {
                                    model_cvitem.Vari_Bk_name = fldvalue.Trim();
                                    break;
                                }
                            // Case "物件名(削除時の退避用)"
                            // .Vari_Bk_deletename = fldvalue.Trim
                            case "物件カナ":
                                {
                                    model_cvitem.Vari_Bk_kana = fldvalue.Trim();
                                    break;
                                }
                            case "建物識別コード":
                                {
                                    model_cvitem.Vari_Tatemono_sikibetu = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
                            case "都道府県コード":
                                {
                                    model_cvitem.Vari_Addr_kenno = fldvalue.Trim();
                                    break;
                                }
                            case "市区町村コード":
                                {
                                    model_cvitem.Vari_Addr_sino = fldvalue.Trim();
                                    break;
                                }
                            case "町地域":
                                {
                                    // 20160622 住所分割処理対応 -chg sta
                                    // 町地域はそのまま移行する(丁番地以降が分割対象)
                                    // '.Vari_Addr_cyo = fldvalue.Trim
                                    // tmp_addresscyotiiki = fldvalue.Trim
                                    model_cvitem.Vari_Addr_cyo = fldvalue.Trim();
                                    break;
                                }
                            // 20160622 住所分割処理対応 -chg end
                            case "丁番地":
                                {
                                    // 20170511 汎用_物件住所の移行不備対応 -chg sta
                                    // '.Vari_Addr_cyome = fldvalue.Trim
                                    // tmp_addresscyobanti1 = fldvalue.Trim
                                    model_cvitem.Vari_Addr_cyome = fldvalue.Trim();
                                    break;
                                }
                            // 20170511 汎用_物件住所の移行不備対応 -chg end
                            case "丁番地選択":
                                {
                                    break;
                                }
                            // .Vari_Addr_cyomeptn = fldvalue.Trim
                            case "街区番号地番":
                                {
                                    // 20170511 汎用_物件住所の移行不備対応 -chg sta
                                    // '.Vari_Addr_banti = fldvalue.Trim
                                    // tmp_addresscyobanti2 = fldvalue.Trim
                                    model_cvitem.Vari_Addr_banti = fldvalue.Trim();
                                    break;
                                }
                            // 20170511 汎用_物件住所の移行不備対応 -chg end
                            case "その他":
                                {
                                    // 20170511 汎用_物件住所の移行不備対応 -chg sta
                                    // '.Vari_Addr_etc = fldvalue.Trim
                                    // tmp_addressother = fldvalue.Trim
                                    model_cvitem.Vari_Addr_etc = fldvalue.Trim();
                                    break;
                                }
                            // 20170511 汎用_物件住所の移行不備対応 -chg end
                            case "物件名(SJIS)":
                                {
                                    model_cvitem.Vari_Bk_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "物件外部No":
                                {
                                    model_cvitem.Vari_Bk_gaibuno = fldvalue.Trim();
                                    break;
                                }
                                // Case "削除理由"
                                // .Vari_Delete_cause = fldvalue.Trim
                        }

                    }

                    // 20170511 汎用_物件住所の移行不備対応 -chg sta
                    // '住所変換/格納
                    // '20160622 住所分割処理対応 -chg sta
                    // 'Dim tmp_address As String = tmp_addresscyotiiki & " " & tmp_addresscyobanti1 & tmp_addresscyobanti2 & " " & tmp_addressother
                    // 'Dim hash_address As New SafeDictionary<string, string>
                    // 'hash_address = AddressChange.Get_CyoBanti(tmp_address)
                    // '.Vari_Addr_cyo = hash_address("mati")
                    // '.Vari_Addr_cyome = hash_address("cyome")
                    // '.Vari_Addr_cyomeptn = hash_address("chomeptn")
                    // '.Vari_Addr_banti = hash_address("banti")
                    // '.Vari_Addr_etc = hash_address("etc")
                    // Dim tmp_address As String = tmp_addresscyobanti1 & tmp_addresscyobanti2 & tmp_addressother
                    // Dim hash_address As New SafeDictionary<string, string>
                    // hash_address = AddressChange.Get_CyoBanti(tmp_address)
                    // .Vari_Addr_cyome = hash_address("cyome")
                    // .Vari_Addr_cyomeptn = hash_address("chomeptn")
                    // .Vari_Addr_banti = hash_address("banti")
                    // .Vari_Addr_etc = hash_address("etc")
                    // '20160622 住所分割処理対応 -chg end
                    // 丁番地の値の有無で処理を分岐する(あれば「丁目」を設定)
                    if (!string.IsNullOrEmpty(model_cvitem.Vari_Addr_cyome))
                    {
                        model_cvitem.Vari_Addr_cyomeptn = "1";
                    }
                    else
                    {
                        model_cvitem.Vari_Addr_cyomeptn = "-1";
                    }
                    // 20170511 汎用_物件住所の移行不備対応 -chg end

                    // 固定値
                    model_cvitem.Vari_Bk_guid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Bk_deleteflg = 0.ToString();
                    model_cvitem.Vari_Delete_guid = null;
                    model_cvitem.Vari_Delete_day = "";
                    model_cvitem.Vari_Delete_cnt = "";
                    model_cvitem.Vari_Bk_deletename = "";
                    model_cvitem.Vari_Delete_cause = "";
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();

                    // 20160519 EXEUpdateに伴う修正 物件基本情報 -add sta
                    // とりあえずNULLにしておく (必要であればここで値を設定)
                    model_cvitem.Vari_Lastupdate = null;
                    // 20160519 EXEUpdateに伴う修正 物件基本情報 -add end

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 20160704 物件住所から郵便番号を読み込む処理の追加 -add sta
                    if (skipflg == false & string.IsNullOrEmpty(model_cvitem.Vari_Post_code))
                    {
                        skipflg = !Get_PostCode(sqlcnnv10, tblname, ref hash_cvitem, ref hash_log);
                    }
                    // 20160704 物件住所から郵便番号を読み込む処理の追加 -add end

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
                // 物件基本情報テーブル外部No一括更新
                int tmpcnt = 0;
                string gaibunoupdatesql = Get_UseQry_Update();
                DBExec.Exec_NonQuery(sqlcnnv10, gaibunoupdatesql, ref tmpcnt);

                // 外部No管理テーブル一括更新
                var obj_bkdatagaibuno = new Bkdata_gaibuno_Repository.SubConv();
                rtn = obj_bkdatagaibuno.Set_Bkdata_Gaibuno(sqlcnnv10);

                // 返却
                return rtn;

            }

            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

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

                string tmp_sql = " SELECT bk_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// [bkdata].[bk_gaibuno]の一括更新クエリ
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE bkdata SET ";
                tmp_sql = tmp_sql + " 	bk_gaibuno = GAIBUNO ";
                tmp_sql = tmp_sql + " FROM bkdata AS BK1 ";
                tmp_sql = tmp_sql + " LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 ROW_NUMBER()OVER(ORDER BY bk_no) AS GAIBUNO ";
                tmp_sql = tmp_sql + " 			,bk_guid ";
                tmp_sql = tmp_sql + " 			,bk_no ";
                tmp_sql = tmp_sql + " 		FROM bkdata ";
                tmp_sql = tmp_sql + " 	) AS BK2 ";
                tmp_sql = tmp_sql + " ON BK1.bk_guid = BK2.bk_guid ";
                tmp_sql = tmp_sql + " ; ";

                return tmp_sql;

            }

            /// <summary>
            /// 郵便番号が存在しない場合、住所から取得する '20160704 物件住所から郵便番号を読み込む処理の追加
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Get_PostCode(SqlConnection sqlcnnv10, string tblname, ref SafeDictionary<string, string> hash_cvitem, ref SafeDictionary<string, string> hash_log)
            {

                bool rtn = true;
                string tmp_kenno = Conversions.ToString(hash_cvitem["addr_kenno"]);
                string tmp_sino = Conversions.ToString(hash_cvitem["addr_sino"]);
                // 20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -add sta
                // 県No、市Noが存在しない場合の処理追加に伴い移動
                string errstr = "";
                string log_key = tblname + "-" + "post_code";
                // 20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -add end

                if (!string.IsNullOrEmpty(tmp_kenno) & !string.IsNullOrEmpty(tmp_sino))
                {

                    string postcodevalu = "";
                    string tmp_sql = "";
                    tmp_sql = tmp_sql + " SELECT TOP 1 post_code FROM post_code ";
                    tmp_sql = tmp_sql + " WHERE ";
                    tmp_sql = tmp_sql + " ken_no = " + tmp_kenno;
                    tmp_sql = tmp_sql + " AND ";
                    tmp_sql = tmp_sql + " si_no = " + tmp_sino;

                    postcodevalu = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));

                    // 20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -del sta
                    // 県No、市Noが存在しない場合の処理追加に伴い移動
                    // Dim errstr As String = ""
                    // Dim log_key As String = tblname & "-" & "post_code"
                    // 20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -del end
                    if (string.IsNullOrEmpty(postcodevalu))
                    {
                        errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED;
                        hash_log.Clear();
                        hash_log.Add(log_key, errstr);
                        rtn = false;
                    }
                    else
                    {
                        hash_cvitem["post_code"] = postcodevalu;
                        errstr = CommonModule.LOG_NAIYO_GET_POSTCODE + "-" + CommonModule.LOG_HUBI_GET_POSTCODE + "-" + CommonModule.LOG_TAISYO_GET_POSTCODE;
                        hash_log.Add(log_key, errstr);
                    }
                }

                // 20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -add sta
                // 県No、市Noが存在しない場合は移行対象外とする。(ログを出力)
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED;
                    hash_log.Clear();
                    hash_log.Add(log_key, errstr);
                    rtn = false;
                    // 20160711 レビュー後指摘対応 物件住所_県No市Noが存在しない場合の処理を追加 -add end
                }

                return rtn;

            }

        }

    }

    #endregion

    #region 物件外部No情報

    // 物件基本情報作成直後に呼び出して処理する

    public class Bkdata_gaibuno_Repository
    {

        public class SubConv
        {

            /// <summary>
            /// 物件基本情報(bkdata)から外部Noを取得して外部No管理テーブル(bkdata_gaibuno)へ挿入する
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Bkdata_Gaibuno(SqlConnection sqlcnnv10)
            {

                bool rtn = true;
                var hash_gaibuno = new SafeDictionary<string, string>();
                var tmp_hash = new SafeDictionary<string, string>();
                var model_cvitem = new Model.Bkdata_gaibuno_Model();
                bool normalflg = true;
                int tmp_cnt = 0;
                var list_bkgaibuno = new List<string>();

                // 対象テーブル初期化
                string tmp_sql_delete = " DELETE FROM bkdata_gaibuno ";
                rtn = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_delete, ref tmp_cnt);

                // テーブル名/フィールド名セット
                string tblname = "bkdata_gaibuno";
                string fldnamegrp = "current_no,history";

                // bkdataから外部No最大値を抽出
                string tmp_sql_select = " SELECT MAX(bk_gaibuno) FROM bkdata ";

                // 変数格納

                model_cvitem.Vari_Current_no = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql_select, ref sqlcnnv10));

                model_cvitem.Vari_History = CommonModule.DefHistory;

                // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                // 挿入処理
                CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, tmp_hash, ref normalflg);

                return rtn;

            }

        }

    }

    #endregion

    #region 物件詳細情報

    public class Bkdata_detail_Repository
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

                var model_cvitem = new Model.Bkdata_detail_Model();           // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 2016.04.06 紐付データ取得処理を外出し -chg sta
                // Call Me.Get_RelData_Bkrui()
                // Call Me.Get_RelData_Kozo()
                SetRelItemToObject.Set_RelData_Bkrui();
                SetRelItemToObject.Set_RelData_Kozo();
                // 2016.04.06 紐付データ取得処理を外出し -chg end

                // 2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -add
                SetRelItemToObject.Set_RelData_Yototiki();
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
                string tblname = "bkdata_detail";
                // 20160519 EXEUpdateに伴う修正 物件詳細情報 -chg sta
                // Dim fldnamegrp As String = "bk_guid,bk_ruinokbn,gps_wgsido,gps_wgskeido,gps_tokyoido," & _
                // "gps_tokyokeido,bk_dentetuflg,syunko_ymd,kaidate,tika," & _
                // "elevator_flg,elevator_number,rooftop_flg,kozo_nokbn,kozo_yanekbn," & _
                // "moto_gy_fudono,biko_kihon,kosu_total,men_nobeyuka,men_nobeyukatubo," & _
                // "men_sikiti,men_sikititubo,men_parking,men_parkingtubo,men_yukatoki," & _
                // "men_yukatokitubo,men_sikititoki,men_sikititokitubo,default_sqtuki,keiyakuyou_sime," & _
                // "yatin_jisansaki,yatin_kozano,yatin_kykozaflg,yatin_kykozano,kozo_taikakbn," & _
                // "gy_sekono,gy_hosyuno,kanri_hosiki,kanri_gyname,kanri_gytanto," & _
                // "kanri_gytel,kanrinin_gyomukeitai,kanrinin_tel,jisya_no,kotu_sonota1," & _
                // "kotu_sonota1kyori,kotu_sonota2,kotu_sonota2kyori,denki_gy_lifeno,water_gy_lifeno," & _
                // "gas_gy_lifeno,haisui_gy_lifeno,toyu_gy_lifeno,lifeline1_gy_lineno,lifeline2_gy_lineno," & _
                // "lifeline3_gy_lineno,kensingyomu_umu,kensinorder_kbn,parenthendo_useflg,parking_kanriflg," & _
                // "parking_car,parking_bike,parking_bicycle,parking_bicyclefreeflg,kubunsyo," & _
                // "soymd_basis,isiwata_kirokukbn,isiwata_syokai1flg,isiwata_syokai2flg,isiwata_syokai3flg," & _
                // "isiwata_syokai4flg,isiwata_gy_sekono,isiwata_cyosaymd,isiwata_cyosakikankbn,isiwata_cyosahani," & _
                // "ishiwata_useflg,isiwata_usearea,isiwata_biko,taisin_sindanflg,taisin_syokai1flg," & _
                // "taisin_syokai2flg,taisin_syokai3flg,taisin_syorui1flg,taisin_syorui2flg,taisin_syorui3flg," & _
                // "taisin_biko,horei_dosyatiiki,horei_ruikbn,horei_naiyo,sikiti_riyoruikbn," & _
                // "sikiti_kystartymd,sikiti_kyendymd,sikiti_biko,nyukyoritu_startymd,history," & _
                // "hyothergyfudo_flg,gomi_hosoku,toki_ymd,syo_kenriflg,syo_kenrikbn," & _
                // "other_kenriflg,men_kentiku,men_kentikutubo,men_kentikutoki,men_kentikutokitubo," & _
                // "bk_logonuser_no,kanri_gyfax,homeelevator_flg,horei_dosyatokubetutiiki,horei_zoseitakutitiiki," & _
                // "horei_tunamitiiki,horei_dosyatiikibiko,horei_dosyatokubetutiikibiko,horei_zoseitakutitiikibiko,horei_tunamitiikibiko," & _
                // "svbunrui_no,krbunrui_no,kanrinin_name,kozo_other,kadoti_flg," & _
                // "cityplan,yototiki,kanrinin_namesjis,kanri_gynamesjis,kanri_gytantosjis," & _
                // "parking_caraki,parking_bikeaki,parking_bicycleaki,syogaku_name,syogaku_kyori," & _
                // "cyugaku_name,cyugaku_kyori,bk_area_no"
                // 20160829 革命10バージョンアップに伴う修正 rooftop_flg を削除 -del 
                string fldnamegrp = "bk_guid,bk_ruinokbn,gps_wgsido,gps_wgskeido,gps_tokyoido," + "gps_tokyokeido,bk_dentetuflg,syunko_ymd,kaidate,tika," + "elevator_flg,elevator_number,kozo_nokbn,kozo_yanekbn," + "moto_gy_fudono,biko_kihon,kosu_total,men_nobeyuka,men_nobeyukatubo," + "men_sikiti,men_sikititubo,men_parking,men_parkingtubo,men_yukatoki," + "men_yukatokitubo,men_sikititoki,men_sikititokitubo,default_sqtuki,keiyakuyou_sime," + "yatin_jisansaki,yatin_kozano,yatin_kykozaflg,yatin_kykozano,kozo_taikakbn," + "gy_sekono,gy_hosyuno,kanri_hosiki,kanri_gyname,kanri_gytanto," + "kanri_gytel,kanrinin_gyomukeitai,kanrinin_tel,jisya_no,kotu_sonota1," + "kotu_sonota1kyori,kotu_sonota2,kotu_sonota2kyori,denki_gy_lifeno,water_gy_lifeno," + "gas_gy_lifeno,haisui_gy_lifeno,toyu_gy_lifeno,lifeline1_gy_lineno,lifeline2_gy_lineno," + "lifeline3_gy_lineno,kensingyomu_umu,kensinorder_kbn,parenthendo_useflg,parking_kanriflg," + "parking_car,parking_bike,parking_bicycle,parking_bicyclefreeflg,kubunsyo," + "isiwata_kirokukbn,isiwata_syokai1flg,isiwata_syokai2flg,isiwata_syokai3flg,isiwata_syokai4flg," + "isiwata_gy_sekono,isiwata_cyosaymd,isiwata_cyosakikankbn,isiwata_cyosahani,ishiwata_useflg," + "isiwata_usearea,isiwata_biko,taisin_sindanflg,taisin_syokai1flg,taisin_syokai2flg," + "taisin_syokai3flg,taisin_syorui1flg,taisin_syorui2flg,taisin_syorui3flg,taisin_biko," + "horei_dosyatiiki,horei_ruikbn,horei_naiyo,sikiti_riyoruikbn,sikiti_kystartymd," + "sikiti_kyendymd,sikiti_biko,nyukyoritu_startymd,history,hyothergyfudo_flg," + "gomi_hosoku,toki_ymd,syo_kenriflg,syo_kenrikbn,other_kenriflg," + "men_kentiku,men_kentikutubo,men_kentikutoki,men_kentikutokitubo,bk_logonuser_no," + "kanri_gyfax,homeelevator_flg,horei_dosyatokubetutiiki,horei_zoseitakutitiiki,horei_tunamitiiki," + "horei_dosyatiikibiko,horei_dosyatokubetutiikibiko,horei_zoseitakutitiikibiko,horei_tunamitiikibiko,svbunrui_no," + "krbunrui_no,kanrinin_name,kozo_other,kadoti_flg,cityplan," + "yototiki,kanrinin_namesjis,kanri_gynamesjis,kanri_gytantosjis,parking_caraki," + "parking_bikeaki,parking_bicycleaki,syogaku_name,syogaku_kyori,cyugaku_name," + "cyugaku_kyori,bk_area_no,emergencyelevator_flg";

























                // 20160519 EXEUpdateに伴う修正 物件詳細情報 -chg end

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

                // ************************
                // 処理開始
                // ************************

                // キーヘッダー名取得
                string fldname_key = readtbl.Columns[keycol - 1].ColumnName.Trim();

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
                    string tmp_tosiyoto1 = "";    // 都市計画/用途地域作業用変数生成
                    string tmp_tosiyoto2 = "";    // 都市計画/用途地域作業用変数生成

                    // キー値取得
                    string fldvalue_key = Typ.ToStr(readtbl.Rows[cntii][keycol - 1]).Trim();

                    // ログ出力用
                    string str_logkey = fldname_key + " = " + fldvalue_key;

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
                            case "物件分類":
                                {
                                    model_cvitem.Vari_Bk_ruinokbn = fldvalue.Trim();
                                    break;
                                }
                            case "緯度":
                                {
                                    model_cvitem.Vari_Gps_wgsido = fldvalue.Trim();
                                    break;
                                }
                            case "経度":
                                {
                                    model_cvitem.Vari_Gps_wgskeido = fldvalue.Trim();
                                    break;
                                }
                            case "緯度(日本測地系)":
                                {
                                    model_cvitem.Vari_Gps_tokyoido = fldvalue.Trim();
                                    break;
                                }
                            case "経度(日本測地系)":
                                {
                                    model_cvitem.Vari_Gps_tokyokeido = fldvalue.Trim();
                                    break;
                                }
                            case "電鉄物件フラグ":
                                {
                                    model_cvitem.Vari_Bk_dentetuflg = fldvalue.Trim();
                                    break;
                                }
                            case "竣工日":
                                {
                                    model_cvitem.Vari_Syunko_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "地上階建て":
                                {
                                    model_cvitem.Vari_Kaidate = fldvalue.Trim();
                                    break;
                                }
                            case "地下階":
                                {
                                    model_cvitem.Vari_Tika = fldvalue.Trim();
                                    break;
                                }
                            case "エレベータフラグ":
                                {
                                    model_cvitem.Vari_Elevator_flg = fldvalue.Trim();
                                    break;
                                }
                            case "エレベータ数":
                                {
                                    model_cvitem.Vari_Elevator_number = fldvalue.Trim();
                                    break;
                                }
                            // 20160829 革命10バージョンアップに伴う修正 -del sta
                            // Case "屋上フラグ"
                            // .Vari_Rooftop_flg = fldvalue.Trim
                            // 20160829 革命10バージョンアップに伴う修正 -del end
                            case "建物構造(基本)-構造":
                                {
                                    model_cvitem.Vari_Kozo_nokbn = fldvalue.Trim();
                                    break;
                                }
                            case "建物構造(基本)-屋根構造":
                                {
                                    model_cvitem.Vari_Kozo_yanekbn = fldvalue.Trim();
                                    break;
                                }
                            case "情報元業者NO":
                                {
                                    model_cvitem.Vari_Moto_gy_fudono = fldvalue.Trim();
                                    break;
                                }
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                            case "総戸数":
                                {
                                    // 20161125 総戸数の半角変換処理を追加 -chg sta
                                    // .Vari_Kosu_total = fldvalue.Trim
                                    model_cvitem.Vari_Kosu_total = Strings.StrConv(fldvalue, VbStrConv.Narrow);
                                    break;
                                }
                            // 20161125 総戸数の半角変換処理を追加 -chg end
                            case "物件面積-延床面積":
                                {
                                    model_cvitem.Vari_Men_nobeyuka = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-延床面積坪数":
                                {
                                    model_cvitem.Vari_Men_nobeyukatubo = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-敷地面積":
                                {
                                    model_cvitem.Vari_Men_sikiti = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-敷地面積坪数":
                                {
                                    model_cvitem.Vari_Men_sikititubo = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-駐車面積":
                                {
                                    model_cvitem.Vari_Men_parking = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-駐車面積坪数":
                                {
                                    model_cvitem.Vari_Men_parkingtubo = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-延床面積(登記)":
                                {
                                    model_cvitem.Vari_Men_yukatoki = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-延床面積坪数(登記)":
                                {
                                    model_cvitem.Vari_Men_yukatokitubo = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-敷地面積(公簿)":
                                {
                                    model_cvitem.Vari_Men_sikititoki = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-敷地面積坪数(公簿)":
                                {
                                    model_cvitem.Vari_Men_sikititokitubo = fldvalue.Trim();
                                    break;
                                }
                            case "入金口座-請求月初期値":
                                {
                                    model_cvitem.Vari_Default_sqtuki = fldvalue.Trim();
                                    break;
                                }
                            case "契約書用入金締切日":
                                {
                                    model_cvitem.Vari_Keiyakuyou_sime = fldvalue.Trim();
                                    break;
                                }
                            case "家賃持参先":
                                {
                                    model_cvitem.Vari_Yatin_jisansaki = fldvalue.Trim();
                                    break;
                                }
                            case "入金口座-家賃入金口座No":
                                {
                                    model_cvitem.Vari_Yatin_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "入金口座-契約金用入金口座の有無":
                                {
                                    model_cvitem.Vari_Yatin_kykozaflg = fldvalue.Trim();
                                    break;
                                }
                            case "入金口座-契約金用入金口座No":
                                {
                                    model_cvitem.Vari_Yatin_kykozano = fldvalue.Trim();
                                    break;
                                }
                            case "耐火構造区分":
                                {
                                    model_cvitem.Vari_Kozo_taikakbn = fldvalue.Trim();
                                    break;
                                }
                            case "施工会社No":
                                {
                                    model_cvitem.Vari_Gy_sekono = fldvalue.Trim();
                                    break;
                                }
                            case "保守業者No":
                                {
                                    model_cvitem.Vari_Gy_hosyuno = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理方式":
                                {
                                    model_cvitem.Vari_Kanri_hosiki = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理業者名":
                                {
                                    model_cvitem.Vari_Kanri_gyname = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理業者担当者":
                                {
                                    model_cvitem.Vari_Kanri_gytanto = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理業者電話番号":
                                {
                                    model_cvitem.Vari_Kanri_gytel = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理業者業務形態":
                                {
                                    model_cvitem.Vari_Kanrinin_gyomukeitai = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理人電話番号":
                                {
                                    model_cvitem.Vari_Kanrinin_tel = fldvalue.Trim();
                                    break;
                                }
                            case "支店No":
                                {
                                    model_cvitem.Vari_Jisya_no = fldvalue.Trim();
                                    break;
                                }
                            case "その他交通１-その他交通":
                                {
                                    model_cvitem.Vari_Kotu_sonota1 = fldvalue.Trim();
                                    break;
                                }
                            case "その他交通１-距離":
                                {
                                    model_cvitem.Vari_Kotu_sonota1kyori = fldvalue.Trim();
                                    break;
                                }
                            case "その他交通２-その他交通":
                                {
                                    model_cvitem.Vari_Kotu_sonota2 = fldvalue.Trim();
                                    break;
                                }
                            case "その他交通２-距離":
                                {
                                    model_cvitem.Vari_Kotu_sonota2kyori = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン-電気-公共機関No":
                                {
                                    model_cvitem.Vari_Denki_gy_lifeno = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン-上水-公共機関No":
                                {
                                    model_cvitem.Vari_Water_gy_lifeno = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン-ガス-公共機関No":
                                {
                                    model_cvitem.Vari_Gas_gy_lifeno = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン-排水-公共機関No":
                                {
                                    model_cvitem.Vari_Haisui_gy_lifeno = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン-灯油-公共機関No":
                                {
                                    model_cvitem.Vari_Toyu_gy_lifeno = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン-その他１-公共機関No":
                                {
                                    model_cvitem.Vari_Lifeline1_gy_lineno = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン-その他２-公共機関No":
                                {
                                    model_cvitem.Vari_Lifeline2_gy_lineno = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン-その他３-公共機関No":
                                {
                                    model_cvitem.Vari_Lifeline3_gy_lineno = fldvalue.Trim();
                                    break;
                                }
                            case "検針業務の有無":
                                {
                                    model_cvitem.Vari_Kensingyomu_umu = fldvalue.Trim();
                                    break;
                                }
                            case "検針登録の並び順":
                                {
                                    model_cvitem.Vari_Kensinorder_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "親子メーター変動費の使用有無":
                                {
                                    model_cvitem.Vari_Parenthendo_useflg = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場-付随駐車場有無":
                                {
                                    model_cvitem.Vari_Parking_kanriflg = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場-自動車台数":
                                {
                                    model_cvitem.Vari_Parking_car = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場-バイク台数":
                                {
                                    model_cvitem.Vari_Parking_bike = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場-自転車台数":
                                {
                                    model_cvitem.Vari_Parking_bicycle = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場-自転車利用自由フラグ":
                                {
                                    model_cvitem.Vari_Parking_bicyclefreeflg = fldvalue.Trim();
                                    break;
                                }
                            case "所有者-一棟・所有区分":
                                {
                                    model_cvitem.Vari_Kubunsyo = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 物件詳細情報 -del sta
                            // Case "該当月/送金月ベース決定"
                            // .Vari_Soymd_basis = fldvalue.Trim
                            // 20160519 EXEUpdateに伴う修正 物件詳細情報 -del end
                            case "石綿使用調査-調査の有無":
                                {
                                    model_cvitem.Vari_Isiwata_kirokukbn = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-調査結果の問合せ先-所有者":
                                {
                                    model_cvitem.Vari_Isiwata_syokai1flg = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-調査結果の問合せ先-管理組合":
                                {
                                    model_cvitem.Vari_Isiwata_syokai2flg = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-調査結果の問合せ先-管理業者":
                                {
                                    model_cvitem.Vari_Isiwata_syokai3flg = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-調査結果の問合せ先-施工業者":
                                {
                                    model_cvitem.Vari_Isiwata_syokai4flg = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-調査年月日":
                                {
                                    model_cvitem.Vari_Isiwata_cyosaymd = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-実施機関":
                                {
                                    model_cvitem.Vari_Isiwata_cyosakikankbn = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-調査範囲":
                                {
                                    model_cvitem.Vari_Isiwata_cyosahani = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-使用有無":
                                {
                                    model_cvitem.Vari_Ishiwata_useflg = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-使用箇所":
                                {
                                    model_cvitem.Vari_Isiwata_usearea = fldvalue.Trim();
                                    break;
                                }
                            case "石綿使用調査-備考":
                                {
                                    model_cvitem.Vari_Isiwata_biko = fldvalue.Trim();
                                    break;
                                }
                            case "耐震診断-診断有無":
                                {
                                    model_cvitem.Vari_Taisin_sindanflg = fldvalue.Trim();
                                    break;
                                }
                            case "耐震診断-診断記録の問合せ先-所有者":
                                {
                                    model_cvitem.Vari_Taisin_syokai1flg = fldvalue.Trim();
                                    break;
                                }
                            case "耐震診断-診断記録の問合せ先-管理組合":
                                {
                                    model_cvitem.Vari_Taisin_syokai2flg = fldvalue.Trim();
                                    break;
                                }
                            case "耐震診断-診断記録の問合せ先-管理業者":
                                {
                                    model_cvitem.Vari_Taisin_syokai3flg = fldvalue.Trim();
                                    break;
                                }
                            case "耐震診断-耐震基準適合証明書の写し":
                                {
                                    model_cvitem.Vari_Taisin_syorui1flg = fldvalue.Trim();
                                    break;
                                }
                            case "耐震診断-住宅性能評価所の写し":
                                {
                                    model_cvitem.Vari_Taisin_syorui2flg = fldvalue.Trim();
                                    break;
                                }
                            case "耐震診断-耐震診断結果の写し":
                                {
                                    model_cvitem.Vari_Taisin_syorui3flg = fldvalue.Trim();
                                    break;
                                }
                            case "耐震診断-備考":
                                {
                                    model_cvitem.Vari_Taisin_biko = fldvalue.Trim();
                                    break;
                                }
                            case "法令-土砂災害警戒地域内外":
                                {
                                    model_cvitem.Vari_Horei_dosyatiiki = fldvalue.Trim();
                                    break;
                                }
                            case "法令-法令分類":
                                {
                                    model_cvitem.Vari_Horei_ruikbn = fldvalue.Trim();
                                    break;
                                }
                            case "法令-法令内容":
                                {
                                    model_cvitem.Vari_Horei_naiyo = fldvalue.Trim();
                                    break;
                                }
                            case "敷地利用-敷地利用種類":
                                {
                                    model_cvitem.Vari_Sikiti_riyoruikbn = fldvalue.Trim();
                                    break;
                                }
                            case "敷地利用-契約期間開始":
                                {
                                    model_cvitem.Vari_Sikiti_kystartymd = fldvalue.Trim();
                                    break;
                                }
                            case "敷地利用-契約期間終了":
                                {
                                    // 20160706 敷地利用-契約期間終了の文字列日付変換処理の追加 -chg sta
                                    // .Vari_Sikiti_kyendymd = fldvalue.Trim
                                    model_cvitem.Vari_Sikiti_kyendymd = Get_ChgDate(fldvalue.Trim());
                                    break;
                                }
                            // 20160706 敷地利用-契約期間終了の文字列日付変換処理の追加 -chg end
                            case "敷地利用-備考":
                                {
                                    model_cvitem.Vari_Sikiti_biko = fldvalue.Trim();
                                    break;
                                }
                            case "入居率一覧対象開始日":
                                {
                                    model_cvitem.Vari_Nyukyoritu_startymd = fldvalue.Trim();
                                    break;
                                }
                            case "部屋毎に業者が異なる場合フラグ":
                                {
                                    model_cvitem.Vari_Hyothergyfudo_flg = fldvalue.Trim();
                                    break;
                                }
                            case "ゴミ出しに関する補足情報":
                                {
                                    model_cvitem.Vari_Gomi_hosoku = fldvalue.Trim();
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
                            case "物件面積-建築面積":
                                {
                                    model_cvitem.Vari_Men_kentiku = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-建築面積坪数":
                                {
                                    model_cvitem.Vari_Men_kentikutubo = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-建築面積(登記)":
                                {
                                    model_cvitem.Vari_Men_kentikutoki = fldvalue.Trim();
                                    break;
                                }
                            case "物件面積-建築面積坪数(登記)":
                                {
                                    model_cvitem.Vari_Men_kentikutokitubo = fldvalue.Trim();
                                    break;
                                }
                            case "自社担当者No":
                                {
                                    model_cvitem.Vari_Bk_logonuser_no = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理業者FAX":
                                {
                                    model_cvitem.Vari_Kanri_gyfax = fldvalue.Trim();
                                    break;
                                }
                            case "エレベータフラグ(家)":
                                {
                                    model_cvitem.Vari_Homeelevator_flg = fldvalue.Trim();
                                    break;
                                }
                            case "法令-土砂災害特別警戒区域内外":
                                {
                                    model_cvitem.Vari_Horei_dosyatokubetutiiki = fldvalue.Trim();
                                    break;
                                }
                            case "法令-造成宅地防災区域内外":
                                {
                                    model_cvitem.Vari_Horei_zoseitakutitiiki = fldvalue.Trim();
                                    break;
                                }
                            case "法令-津波災害警戒区域内外":
                                {
                                    model_cvitem.Vari_Horei_tunamitiiki = fldvalue.Trim();
                                    break;
                                }
                            case "法令-土砂災害警戒地域備考":
                                {
                                    model_cvitem.Vari_Horei_dosyatiikibiko = fldvalue.Trim();
                                    break;
                                }
                            case "法令-土砂災害特別警戒区域備考":
                                {
                                    model_cvitem.Vari_Horei_dosyatokubetutiikibiko = fldvalue.Trim();
                                    break;
                                }
                            case "法令-造成宅地防災区域備考":
                                {
                                    model_cvitem.Vari_Horei_zoseitakutitiikibiko = fldvalue.Trim();
                                    break;
                                }
                            case "法令-津波災害警戒区域備考":
                                {
                                    model_cvitem.Vari_Horei_tunamitiikibiko = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理人名":
                                {
                                    model_cvitem.Vari_Kanrinin_name = fldvalue.Trim();
                                    break;
                                }
                            case "その他建物構造":
                                {
                                    model_cvitem.Vari_Kozo_other = fldvalue.Trim();
                                    break;
                                }
                            case "角地フラグ":
                                {
                                    model_cvitem.Vari_Kadoti_flg = fldvalue.Trim();
                                    break;
                                }
                            // 2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -chg sta
                            // Case "都市計画・用途地域1"
                            // '.Vari_Cityplan = fldvalue.Trim
                            // tmp_tosiyoto1 = fldvalue.Trim
                            // Case "都市計画・用途地域2"
                            // '.Vari_Yototiki = fldvalue.Trim
                            // tmp_tosiyoto2 = fldvalue.Trim
                            case "都市計画":
                                {
                                    tmp_tosiyoto1 = fldvalue.Trim();
                                    break;
                                }
                            case "用途地域":
                                {
                                    tmp_tosiyoto2 = fldvalue.Trim();
                                    break;
                                }
                            // 2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -chg end
                            case "管理形態-管理人名(SJIS)":
                                {
                                    model_cvitem.Vari_Kanrinin_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理業者名(SJIS)":
                                {
                                    model_cvitem.Vari_Kanri_gynamesjis = fldvalue.Trim();
                                    break;
                                }
                            case "管理形態-管理業者担当者名(SJIS)":
                                {
                                    model_cvitem.Vari_Kanri_gytantosjis = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場空き台数":
                                {
                                    model_cvitem.Vari_Parking_caraki = fldvalue.Trim();
                                    break;
                                }
                            case "バイク置き場空き台数":
                                {
                                    model_cvitem.Vari_Parking_bikeaki = fldvalue.Trim();
                                    break;
                                }
                            case "駐輪場空き台数":
                                {
                                    model_cvitem.Vari_Parking_bicycleaki = fldvalue.Trim();
                                    break;
                                }
                            case "小学校区":
                                {
                                    model_cvitem.Vari_Syogaku_name = fldvalue.Trim();
                                    break;
                                }
                            case "小学校距離":
                                {
                                    model_cvitem.Vari_Syogaku_kyori = fldvalue.Trim();
                                    break;
                                }
                            case "中学校区":
                                {
                                    model_cvitem.Vari_Cyugaku_name = fldvalue.Trim();
                                    break;
                                }
                            case "中学校距離":
                                {
                                    model_cvitem.Vari_Cyugaku_kyori = fldvalue.Trim();
                                    break;
                                }
                            case "エリア":
                                {
                                    // エリアマスタと照合
                                    model_cvitem.Vari_Bk_area_no = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 物件詳細情報 -add sta
                            case "非常用エレベーターフラグ":
                                {
                                    model_cvitem.Vari_Emergencyelevator_flg = fldvalue.Trim();
                                    break;
                                }
                                // 20160519 EXEUpdateに伴う修正 物件詳細情報 -add end
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Krbunrui_no = (-1).ToString();
                    model_cvitem.Vari_Svbunrui_no = (-1).ToString();
                    model_cvitem.Vari_Isiwata_gy_sekono = null;

                    // 2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -chg sta
                    // '都市計画/用途地域照合
                    // Dim tmp_tosi As String = ""
                    // Dim tmp_yoto As String = ""

                    // Call Me.Get_TosiYotoValue(tmp_tosiyoto1, tmp_tosiyoto2, tmp_tosi, tmp_yoto)

                    // .Vari_Cityplan = tmp_tosi
                    // .Vari_Yototiki = tmp_yoto

                    string tmp_tosi = "";
                    string tmp_yoto = "";

                    // 都市計画・用途地域を判別
                    Set_Tosiyotokbn(tmp_tosiyoto1, tmp_tosiyoto2, ref tmp_tosi, ref tmp_yoto);

                    model_cvitem.Vari_Cityplan = tmp_tosi;
                    model_cvitem.Vari_Yototiki = tmp_yoto;
                    // 2016.04.06 都市計画・用途地域を紐データから取得するように修正 -chg end

                    // 20161128 物件面積関連情報の追加 -add sta
                    if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                    {
                        model_cvitem.Vari_Men_nobeyukatubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_nobeyuka)), Typ.RoundingTypes.Kirisute, 2).ToString();
                        model_cvitem.Vari_Men_sikititubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_sikiti)), Typ.RoundingTypes.Kirisute, 2).ToString();
                        model_cvitem.Vari_Men_parkingtubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_parking)), Typ.RoundingTypes.Kirisute, 2).ToString();
                        model_cvitem.Vari_Men_yukatokitubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_yukatoki)), Typ.RoundingTypes.Kirisute, 2).ToString();
                        model_cvitem.Vari_Men_sikititokitubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_sikititoki)), Typ.RoundingTypes.Kirisute, 2).ToString();
                        model_cvitem.Vari_Men_kentikutubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_kentiku)), Typ.RoundingTypes.Kirisute, 2).ToString();
                        model_cvitem.Vari_Men_kentikutokitubo = Typ.ToDecimal(Typ.ToTubo(Typ.ToDouble(model_cvitem.Vari_Men_kentikutoki)), Typ.RoundingTypes.Kirisute, 2).ToString();
                    }
                    // 20161128 物件面積関連情報の追加 -add end

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

                        // 20260914 紐付マスタ(構造マスタ)での解決後、建物構造(基本)-構造(kozo_nokbn)が
                        // 「その他」(コード99)になった場合、その他建物構造には紐付け前の元の値
                        // (pre_物件情報.kozo_nokbnの生値)をそのまま設定する。それ以外はクリアする
                        if (hash_cvitem["kozo_nokbn"] == "99")
                        {
                            hash_cvitem["kozo_other"] = model_cvitem.Vari_Kozo_nokbn;
                        }
                        else
                        {
                            hash_cvitem["kozo_other"] = "";
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
            /// 物件分類仮テーブルからハッシュテーブルを生成
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, string> Get_HashBkRui(SqlConnection sqlcnnv10)
            {

                var rtn_hash = new SafeDictionary<string, string>();
                var readtbl = new DataTable();
                int reccnt;
                var flg = default(bool);
                string fldname = "";
                string fldvalue = "";
                string tmp_v7ruino = "";
                string tmp_v10ruino = "";

                // 物件分類仮テーブル取得
                string tmp_sql = "SELECT * FROM tmp_bkrui_mst";
                reccnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, ref flg);

                for (int cntii = 0, loopTo = readtbl.Rows.Count - 1; cntii <= loopTo; cntii++)
                {
                    for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                    {

                        // 項目名取得
                        fldname = readtbl.Columns[cntjj].ColumnName;

                        // 登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]);

                        // 変数セット
                        switch (fldname ?? "")
                        {
                            case "移行元物件分類№":
                                {
                                    tmp_v7ruino = fldvalue;
                                    break;
                                }
                            case "賃貸革命10物件分類№":
                                {
                                    tmp_v10ruino = fldvalue;
                                    break;
                                }
                        }

                    }

                    // ハッシュテーブルセット
                    rtn_hash.Add(tmp_v7ruino, tmp_v10ruino);

                }

                return rtn_hash;

            }

            /// <summary>
            /// エリアマスタからハッシュテーブルを生成(名称ではなく番号を設定するようになったため)
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, string> Get_HashArea(SqlConnection sqlcnnv10)
            {

                var rtn_hash = new SafeDictionary<string, string>();
                var readtbl = new DataTable();
                int reccnt;
                var flg = default(bool);
                string fldname = "";
                string fldvalue = "";
                string tmp_areano = "";
                string tmp_areaname = "";

                // エリアマスタ取得
                string tmp_sql = "SELECT * FROM m_area";
                reccnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, ref flg);

                for (int cntii = 0, loopTo = readtbl.Rows.Count - 1; cntii <= loopTo; cntii++)
                {
                    for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                    {

                        // 項目名取得
                        fldname = readtbl.Columns[cntjj].ColumnName;

                        // 登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]);

                        // 変数セット
                        switch (fldname ?? "")
                        {
                            case "area_no":
                                {
                                    tmp_areano = fldvalue;
                                    break;
                                }
                            case "area_name":
                                {
                                    tmp_areaname = fldvalue;
                                    break;
                                }
                        }

                    }

                    // ハッシュテーブルセット
                    rtn_hash.Add(tmp_areaname, tmp_areano);

                }

                return rtn_hash;

            }

            /// <summary>
            /// 構造仮マスタからハッシュテーブルを生成(番号を設定するため)
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public SafeDictionary<string, string> Get_HashKozo(SqlConnection sqlcnnv10)
            {

                var rtn_hash = new SafeDictionary<string, string>();
                var readtbl = new DataTable();
                int reccnt;
                var flg = default(bool);
                string fldname = "";
                string fldvalue = "";
                string tmp_v7kozoname = "";
                string tmp_v10kozono = "";

                // 構造マスタ仮テーブル取得
                string tmp_sql = "SELECT * FROM tmp_kozo_mst";
                reccnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, ref flg);

                for (int cntii = 0, loopTo = readtbl.Rows.Count - 1; cntii <= loopTo; cntii++)
                {
                    for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                    {

                        // 項目名取得
                        fldname = readtbl.Columns[cntjj].ColumnName;

                        // 登録値取得
                        fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]);

                        // 変数セット
                        switch (fldname ?? "")
                        {
                            case "移行元構造名称":
                                {
                                    tmp_v7kozoname = fldvalue;
                                    break;
                                }
                            case "賃貸革命10構造№":
                                {
                                    tmp_v10kozono = fldvalue;
                                    break;
                                }
                        }

                    }

                    // ハッシュテーブルセット
                    rtn_hash.Add(tmp_v7kozoname, tmp_v10kozono);

                }

                return rtn_hash;

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

            // 2016.04.06 紐付データ取得処理を外出し -del sta
            // ''' <summary>
            // ''' 物件分類紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_Bkrui()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "物件分類マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldbkruiname As String = ""
            // Dim tmp_newbkruino As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元物件分類名称"
            // tmp_oldbkruiname = fldvalue.Trim
            // Case "賃貸革命10物件分類No"
            // tmp_newbkruino = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newbkruino <> "" Then
            // Hash_Rel_Bkrui.Add(tmp_oldbkruiname, tmp_newbkruino)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub

            // ''' <summary>
            // ''' 構造マスタ紐付情報取得
            // ''' </summary>
            // ''' <remarks></remarks>
            // Public Sub Get_RelData_Kozo()

            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim startrow As Integer                                         '書込開始行
            // Dim columncnt As Integer                                        '列数
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim rowcnt As Integer                                           '書込行数
            // Dim relsheetname As String = "構造マスタ"
            // Dim rtn As Boolean = True

            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

            // For cntii = 0 To rowcnt - 1

            // '作業用変数作成
            // Dim tmp_oldbkozoname As String = ""
            // Dim tmp_newkozono As String = ""

            // '紐付設定値取得
            // For cntjj = 1 To columncnt

            // 'ヘッダー格納
            // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

            // '移行値格納
            // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

            // Select Case fldname
            // Case "移行元構造名称"
            // tmp_oldbkozoname = fldvalue.Trim
            // Case "賃貸革命10構造No"
            // tmp_newkozono = fldvalue.Trim
            // End Select

            // Next

            // 'ハッシュテーブル格納
            // If tmp_newkozono <> "" Then
            // Hash_Rel_Kozo.Add(tmp_oldbkozoname, tmp_newkozono)
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            // End Sub
            // 2016.04.06 紐付データ取得処理を外出し -del end

            // 2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -del sta
            // ''' <summary>
            // ''' 用途地域の文字列を10の都市計画/用途地域と照合
            // ''' </summary>
            // ''' <param name="tosiyoto1"></param>
            // ''' <param name="tosiyoto2"></param>
            // ''' <param name="tosi"></param>
            // ''' <param name="yoto"></param>
            // ''' <remarks></remarks>
            // Public Sub Get_TosiYotoValue(ByVal tosiyoto1 As String, ByVal tosiyoto2 As String, ByRef tosi As String, ByRef yoto As String)

            // Dim tmp_tosiyotokbn1 As Integer
            // Dim tmp_tosiyotono1 As String = ""
            // Dim tmp_tosiyotokbn2 As Integer
            // Dim tmp_tosiyotono2 As String = ""

            // TosiYotoConv.Get_TosiYoto(tosiyoto1, tmp_tosiyotokbn1, tmp_tosiyotono1)
            // TosiYotoConv.Get_TosiYoto(tosiyoto2, tmp_tosiyotokbn2, tmp_tosiyotono2)

            // Select Case True
            // Case tmp_tosiyotokbn1 = 0 And tmp_tosiyotokbn2 = 0  '1個目が空白、2個目が空白
            // tosi = -1
            // yoto = -1
            // Case tmp_tosiyotokbn1 = 0 And tmp_tosiyotokbn2 = 1  '1個目が空白、2個目が都市計画
            // tosi = tmp_tosiyotono2
            // yoto = -1
            // Case tmp_tosiyotokbn1 = 0 And tmp_tosiyotokbn2 = 2  '1個目が空白、2個目が用途地域
            // tosi = -1
            // yoto = tmp_tosiyotono2
            // Case tmp_tosiyotokbn1 = 0 And tmp_tosiyotokbn2 = 3  '1個目が空白、2個目が該当データ無し
            // tosi = -1
            // yoto = 99
            // Case tmp_tosiyotokbn1 = 1 And tmp_tosiyotokbn2 = 0  '1個目が都市計画、2個目が空白
            // tosi = tmp_tosiyotono1
            // yoto = -1
            // Case tmp_tosiyotokbn1 = 1 And tmp_tosiyotokbn2 = 1  '1個目が都市計画、2個目が都市計画
            // tosi = tmp_tosiyotono1
            // yoto = -1
            // Case tmp_tosiyotokbn1 = 1 And tmp_tosiyotokbn2 = 2  '1個目が都市計画、2個目が用途地域
            // tosi = tmp_tosiyotono1
            // yoto = tmp_tosiyotono2
            // Case tmp_tosiyotokbn1 = 1 And tmp_tosiyotokbn2 = 3  '1個目が都市計画、2個目が該当データ無し
            // tosi = tmp_tosiyotono1
            // yoto = 99
            // Case tmp_tosiyotokbn1 = 2 And tmp_tosiyotokbn2 = 0  '1個目が用途地域、2個目が空白
            // tosi = -1
            // yoto = tmp_tosiyotono1
            // Case tmp_tosiyotokbn1 = 2 And tmp_tosiyotokbn2 = 1  '1個目が用途地域、2個目が都市計画
            // tosi = tmp_tosiyotono2
            // yoto = tmp_tosiyotono1
            // Case tmp_tosiyotokbn1 = 2 And tmp_tosiyotokbn2 = 2  '1個目が用途地域、2個目が用途地域
            // tosi = -1
            // yoto = tmp_tosiyotono1
            // Case tmp_tosiyotokbn1 = 2 And tmp_tosiyotokbn2 = 3  '1個目が用途地域、2個目が該当データ無し
            // tosi = 99
            // yoto = tmp_tosiyotono1
            // Case tmp_tosiyotokbn1 = 3 And tmp_tosiyotokbn2 = 0  '1個目が該当データ無し、2個目が空白
            // tosi = 99
            // yoto = -1
            // Case tmp_tosiyotokbn1 = 3 And tmp_tosiyotokbn2 = 1  '1個目が該当データ無し、2個目が都市計画
            // tosi = tmp_tosiyotono2
            // yoto = 99
            // Case tmp_tosiyotokbn1 = 3 And tmp_tosiyotokbn2 = 2  '1個目が該当データ無し、2個目が用途地域
            // tosi = 99
            // yoto = tmp_tosiyotono2
            // Case tmp_tosiyotokbn1 = 3 And tmp_tosiyotokbn2 = 3  '1個目が該当データ無し、2個目が該当データ無し
            // tosi = 99
            // yoto = 99
            // End Select

            // End Sub
            // 2016.04.06 都市計画・用途地域を紐付データから取得するように修正 -del end

            /// <summary>
            /// 都市計画・用途地域を判別する
            /// </summary>
            /// <param name="tosiyoto1"></param>
            /// <param name="tosiyoto2"></param>
            /// <param name="tosi"></param>
            /// <param name="yoto"></param>
            /// <remarks></remarks>
            public void Set_Tosiyotokbn(string tosiyoto1, string tosiyoto2, ref string tosi, ref string yoto)
            {

                // 両方データが存在しない場合は処理を抜ける
                if (string.IsNullOrEmpty(tosiyoto1) & string.IsNullOrEmpty(tosiyoto2))
                {
                    return;
                }

                string tmp_kbn1 = "";
                string tmp_kbn2 = "";

                // 1つ目の都市用途区分を取得
                if (CommonModule.Hash_Rel_Tosiyotokbn.ContainsKey(tosiyoto1))
                {
                    tmp_kbn1 = Conversions.ToString(CommonModule.Hash_Rel_Tosiyotokbn[tosiyoto1]);
                }

                // 2つ目の都市用途区分を取得
                if (CommonModule.Hash_Rel_Tosiyotokbn.ContainsKey(tosiyoto2))
                {
                    tmp_kbn2 = Conversions.ToString(CommonModule.Hash_Rel_Tosiyotokbn[tosiyoto2]);
                }

                switch (true)
                {
                    case object _ when string.IsNullOrEmpty(tmp_kbn1) & string.IsNullOrEmpty(tmp_kbn2):
                        {
                            tosi = "";
                            yoto = "";
                            break;
                        }
                    case object _ when string.IsNullOrEmpty(tmp_kbn1) & tmp_kbn2 == "都市計画":
                        {
                            tosi = tosiyoto2;
                            yoto = "";
                            break;
                        }
                    case object _ when string.IsNullOrEmpty(tmp_kbn1) & tmp_kbn2 == "用途地域":
                        {
                            tosi = "";
                            yoto = tosiyoto2;
                            break;
                        }
                    case object _ when tmp_kbn1 == "都市計画" & string.IsNullOrEmpty(tmp_kbn2):
                        {
                            tosi = tosiyoto1;
                            yoto = "";
                            break;
                        }
                    case object _ when tmp_kbn1 == "都市計画" & tmp_kbn2 == "都市計画":
                        {
                            tosi = tosiyoto1;
                            yoto = "都市計画の重複";    // 重複した場合は移行不可のため判別用の文字列を入れておく
                            break;
                        }
                    case object _ when tmp_kbn1 == "都市計画" & tmp_kbn2 == "用途地域":
                        {
                            tosi = tosiyoto1;
                            yoto = tosiyoto2;
                            break;
                        }
                    case object _ when tmp_kbn1 == "用途地域" & string.IsNullOrEmpty(tmp_kbn2):
                        {
                            tosi = "";
                            yoto = tosiyoto1;
                            break;
                        }
                    case object _ when tmp_kbn1 == "用途地域" & tmp_kbn2 == "都市計画":
                        {
                            tosi = tosiyoto2;
                            yoto = tosiyoto1;
                            break;
                        }
                    case object _ when tmp_kbn1 == "用途地域" & tmp_kbn2 == "用途地域":
                        {
                            tosi = "用途地域の重複";    // 重複した場合は移行不可のため判別用の文字列を入れておく
                            yoto = tosiyoto1;
                            break;
                        }
                }


            }

            /// <summary>
            /// 敷地利用-契約期間終了の文字列を日付型の文字列へ変換する処理 '20160706 敷地利用-契約期間終了の文字列日付変換処理の追加
            /// 「平成○○年●●月まで」の形のみ変換対象とする(年号は明治、大正、昭和、平成)
            /// </summary>
            /// <param name="value"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            private string Get_ChgDate(string value)
            {

                string rtn_str = "";
                bool matchflg = false;
                var list_nengo = new List<string>() { "明治", "大正", "昭和", "平成" };
                string tmp_year = "年";
                string tmp_num = "[1-9]?[0-9]";
                string tmp_month = "月";
                string tmp_kotei = "まで";
                string chk_nengo = "";
                string tmp_value = "";

                // 空文字チェック
                if (string.IsNullOrEmpty(value))
                {
                    return rtn_str;
                }

                // 半角変換と空白文字列除去
                tmp_value = Strings.StrConv(value, VbStrConv.Narrow);
                tmp_value = tmp_value.Replace(" ", "");

                // 形式チェック
                foreach (var nengo in list_nengo)
                {
                    string tmp_syogo = nengo + tmp_num + tmp_year + tmp_num + tmp_month + tmp_kotei;
                    if (System.Text.RegularExpressions.Regex.IsMatch(tmp_value, tmp_syogo))
                    {
                        chk_nengo = nengo;
                        matchflg = true;
                        break;
                    }
                }

                // 形式が標準以外の場合は処理を抜ける
                if (matchflg == false)
                {
                    // 20160808 連動検証中に発見した不具合修正 不適合時は空文字を返す -del
                    // rtn_str = value
                    return rtn_str;
                }

                // 形式が標準の場合は数値チェックを行う
                int heisei_maxyear = int.Parse((int.Parse(DateTime.Now.Year.ToString()) - 1988).ToString());
                // 20160711 レビュー後指摘対応 敷地利用処理の最大年数修正 -chg sta
                // Dim hash_nengomaxyear As New SafeDictionary<string, string> From _
                // {{"明治", 44}, {"大正", 14}, {"昭和", 63}, {"平成", heisei_maxyear}}
                var hash_nengomaxyear = new SafeDictionary<string, int>() { { "明治", 45 }, { "大正", 15 }, { "昭和", 64 }, { "平成", heisei_maxyear } };
                // 20160711 レビュー後指摘対応 敷地利用処理の最大年数修正 -chg end
                var hash_nengotoseirekiaddyear = new SafeDictionary<string, int>() { { "明治", 1867 }, { "大正", 1911 }, { "昭和", 1925 }, { "平成", 1988 } };
                string[] tmp_str = tmp_value.Split('年');

                // 年のチェック
                int chk_year = int.Parse(tmp_str[0].Replace(chk_nengo, ""));
                int max_year = Conversions.ToInteger(hash_nengomaxyear[chk_nengo]);

                if (max_year < chk_year)
                {
                    // 20160808 連動検証中に発見した不具合修正 不適合時は空文字を返す -del
                    // rtn_str = value
                    return rtn_str;
                }

                // 月のチェック
                // 20160711 レビュー後指摘対応 敷地利用処理の変数表記 -chg sta
                // Dim chk_month As Integer = Int32.Parse(tmp_str(1).Replace("月まで", ""))
                int chk_month = int.Parse(tmp_str[1].Replace(tmp_month + tmp_kotei, ""));
                // 20160711 レビュー後指摘対応 敷地利用処理の変数表記 -chg end
                int max_month = 12;

                if (max_month < chk_month)
                {
                    // 20160808 連動検証中に発見した不具合修正 不適合時は空文字を返す -del
                    // rtn_str = value
                    return rtn_str;
                }

                // 全てのチェックに問題が無い場合は西暦(月末)に変換して返す
                string chg_seirekiyear = Operators.AddObject(chk_year, hash_nengotoseirekiaddyear[chk_nengo]).ToString();
                string chg_month = Conversions.ToString(Interaction.IIf(chk_month.ToString().Length == 1, ("0" + chk_month).ToString(), chk_month.ToString()));
                string tmp_lastday = DateTime.DaysInMonth(chk_year, chk_month).ToString();

                string chg_date = chg_seirekiyear + "/" + chg_month + "/" + tmp_lastday;

                rtn_str = chg_date;

                return rtn_str;

            }

        }

    }

    #endregion

    #region 物件所有者情報

    public class Bkdata_syo_Repository
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
                var model_cvitem = new Model.Bkdata_syo_Model();              // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

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
                // テーブル名/フィールド名セット
                string tblname_base = "bkdata";
                string tblname = "bkdata_syo";
                string fldnamegrp = "bk_guid,kn_no,sorule_guid,kasi1_ow_no,syo1_ow_no," + "kasi2_ow_no,syo2_ow_no,syo_startymd,syo_endymd,history";

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
                            case "物件No":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,kn_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件ゴミ情報

    public class Bkdata_dust_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Bkdata_dust_Model();             // 移行値格納用モデル初期化
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
                string tblname_base = "bkdata";
                string tblname = "bkdata_dust";
                string fldnamegrp = "bk_guid,dust_no,dust_rui,dust_youbi,history";

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
                            case "物件NO":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
                            case "ゴミ情報No":
                                {
                                    model_cvitem.Vari_Dust_no = fldvalue.Trim();
                                    break;
                                }
                            case "ゴミ分類":
                                {
                                    model_cvitem.Vari_Dust_rui = fldvalue.Trim();
                                    break;
                                }
                            case "ゴミ出し曜日":
                                {
                                    model_cvitem.Vari_Dust_youbi = fldvalue.Trim();
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,dust_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件変動費親メーター情報(汎用ツールのみ)

    public class Bkdata_hendo_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Bkdata_hendo_Model();            // 移行値格納用モデル初期化
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
                string tblname_base = "bkdata";
                string tblname = "bkdata_hendo";
                string fldnamegrp = "bk_guid,bkhendo_guid,rec_no,hendo_kbn,meter_name," + "nkin_no,child_meter,biko,useflg,history," + "sq_mmkbn,tani,tanisjis";


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
                            case "物件NO":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
                            case "行No":
                                {
                                    model_cvitem.Vari_Rec_no = fldvalue.Trim();
                                    break;
                                }
                            case "メーター分類 (水道/ガス/電気/灯油/その他)":
                                {
                                    model_cvitem.Vari_Hendo_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "メーター名(部屋no・フロア名など)":
                                {
                                    model_cvitem.Vari_Meter_name = fldvalue.Trim();
                                    break;
                                }
                            case "変動費入金項目":
                                {
                                    model_cvitem.Vari_Nkin_no = fldvalue.Trim();
                                    break;
                                }
                            case "子メーター検針":
                                {
                                    model_cvitem.Vari_Child_meter = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue.Trim();
                                    break;
                                }
                            case "使用フラグ":
                                {
                                    model_cvitem.Vari_Useflg = fldvalue.Trim();
                                    break;
                                }
                            case "請求月":
                                {
                                    model_cvitem.Vari_Sq_mmkbn = fldvalue.Trim();
                                    break;
                                }
                            case "単位":
                                {
                                    model_cvitem.Vari_Tani = fldvalue.Trim();
                                    break;
                                }
                            case "単位(SJIS)":
                                {
                                    model_cvitem.Vari_Tanisjis = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Bkhendo_guid = Guid.NewGuid().ToString();
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,rec_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件鍵情報

    public class Bkdata_kagi_Repository
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

                var model_cvitem = new Model.Bkdata_kagi_Model();             // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

                int cnt_cvcnt = 0;                                    // 20160531 鍵情報移行処理の修正 共用鍵の件数取得用 -add

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
                // テーブル名/フィールド名セット
                string tblname_base = "bkdata";
                string tblname = "bkdata_kagi";
                string fldnamegrp = "bk_guid,kagi_no,honsu,biko,history," + "hokan,gyshare";

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
                            case "物件No":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue;
                                    break;
                                }
                            case "鍵No":      // 20160613 鍵情報の取得処理修正 -chg 鍵タイトル名→鍵No
                                {
                                    model_cvitem.Vari_Kagi_no = fldvalue;
                                    break;
                                }
                            case "鍵本数":
                                {
                                    model_cvitem.Vari_Honsu = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            case "保管場所":
                                {
                                    model_cvitem.Vari_Hokan = fldvalue;
                                    break;
                                }
                            case "業者間での情報共有":
                                {
                                    model_cvitem.Vari_Gyshare = fldvalue;
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
                    // ''    hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

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

                    // 'If Hash_KagiTitleKyoyo.Contains(.Vari_Kagi_no) Then

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
                    // '        hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

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
                    // hash_cvitem("bk_guid") = hash_guid.Item(hash_cvitem.Item("bk_guid"))

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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,kagi_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件権利情報

    public class Bkdata_kenri_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Bkdata_kenri_Model();            // 移行値格納用モデル初期化
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
                string tblname_base = "bkdata";
                string tblname = "bkdata_kenri";
                string fldnamegrp = "bk_guid,kenri_no,other_kenrirui,other_kenribiko,history";

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

                    // 2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add sta
                    // 作業用変数
                    string tmp_kenriumu = "";
                    // 2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add end

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
                                    model_cvitem.Vari_Bk_guid = fldvalue;
                                    break;
                                }
                            case "権利情報No":
                                {
                                    model_cvitem.Vari_Kenri_no = fldvalue;
                                    break;
                                }
                            // 2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add sta
                            case "所有権以外の権利の有無":
                                {
                                    tmp_kenriumu = fldvalue;
                                    break;
                                }
                            // 2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add end
                            case "所有権以外の権利の種類":
                                {
                                    model_cvitem.Vari_Other_kenrirui = fldvalue;
                                    break;
                                }
                            case "内容":
                                {
                                    model_cvitem.Vari_Other_kenribiko = fldvalue;
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
                        hash_cvitem["bk_guid"] = hash_guid[hash_cvitem["bk_guid"]];

                        // 挿入処理
                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                        if (normalflg)
                        {
                            list_chkduplicate.Add(fldvalue_key);
                            tmp_cvcnt = tmp_cvcnt + 1;

                            // 2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add sta
                            // 更新処理 (所有権以外の権利有無のフラグはbkdata_detailに存在するため)
                            int tmpcnt = 0;
                            string kenriflgupdatesql = Get_UseQry_Update(Conversions.ToString(hash_cvitem["bk_guid"]), tmp_kenriumu);
                            DBExec.Exec_NonQuery(sqlcnnv10, kenriflgupdatesql, ref tmpcnt);
                            // 2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add end

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

            /// <summary>
            /// 既存データの取得
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,kenri_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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
            /// [bkdata_detail].[other_kenriflg]の一括更新クエリ 2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update(string bkguid, string updatevalue)
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE bkdata_detail SET ";
                tmp_sql = tmp_sql + " 	other_kenriflg = " + updatevalue;
                tmp_sql = tmp_sql + " WHERE bk_guid = '" + bkguid + "' ";
                tmp_sql = tmp_sql + " ; ";

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 物件交通情報

    public class Bkdata_kotu_Repository
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

                var model_cvitem = new Model.Bkdata_kotu_Model();             // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

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
                // テーブル名/フィールド名セット
                string tblname_base = "bkdata";
                string tblname = "bkdata_kotu";
                // 20160519 EXEUpdateに伴う修正 物件交通情報 -del(kotu_syudanを削除して成形)
                string fldnamegrp = "bk_guid,ensen_cnt,ensen_no,eki_no,firsttrain_flg," + "kyori,toho_min,car_min,bus_min,bus_company," + "bus_station,bus_tohomin,bus_kyori,history,setting_kotutype," + "bus_companytoeki,bus_stationtoeki,bus_kyoritoeki,bus_tohomintoeki,car_kyori";



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
                            case "沿線番号":
                                {
                                    model_cvitem.Vari_Ensen_cnt = fldvalue.Trim();
                                    break;
                                }
                            case "沿線No":
                                {
                                    model_cvitem.Vari_Ensen_no = fldvalue.Trim();
                                    break;
                                }
                            case "駅No":
                                {
                                    model_cvitem.Vari_Eki_no = fldvalue.Trim();
                                    break;
                                }
                            case "始発フラグ":
                                {
                                    model_cvitem.Vari_Firsttrain_flg = fldvalue.Trim();
                                    break;
                                }
                            case "距離":
                                {
                                    model_cvitem.Vari_Kyori = fldvalue.Trim();
                                    break;
                                }
                            case "徒歩":
                                {
                                    model_cvitem.Vari_Toho_min = fldvalue.Trim();
                                    break;
                                }
                            case "車":
                                {
                                    model_cvitem.Vari_Car_min = fldvalue.Trim();
                                    break;
                                }
                            case "バス":
                                {
                                    model_cvitem.Vari_Bus_min = fldvalue.Trim();
                                    break;
                                }
                            case "バス会社":
                                {
                                    model_cvitem.Vari_Bus_company = fldvalue.Trim();
                                    break;
                                }
                            case "バス停":
                                {
                                    model_cvitem.Vari_Bus_station = fldvalue.Trim();
                                    break;
                                }
                            case "バス停徒歩":
                                {
                                    model_cvitem.Vari_Bus_tohomin = fldvalue.Trim();
                                    break;
                                }
                            case "バス停までの距離":
                                {
                                    model_cvitem.Vari_Bus_kyori = fldvalue.Trim();
                                    break;
                                }
                            case "設定する交通のタイプ":
                                {
                                    model_cvitem.Vari_Setting_kotutype = fldvalue.Trim();
                                    break;
                                }
                            case "バス系統・路線名(駅までバスを利用)":
                                {
                                    model_cvitem.Vari_Bus_companytoeki = fldvalue.Trim();
                                    break;
                                }
                            case "バス停名(駅までバスを利用)":
                                {
                                    model_cvitem.Vari_Bus_stationtoeki = fldvalue.Trim();
                                    break;
                                }
                            case "バス停(駅までバスを利用)までの道路距離":
                                {
                                    model_cvitem.Vari_Bus_kyoritoeki = fldvalue.Trim();
                                    break;
                                }
                            case "バス停(駅までバスを利用)までの徒歩(分)":
                                {
                                    model_cvitem.Vari_Bus_tohomintoeki = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 物件交通情報 -del sta
                            // Case "交通手段"
                            // .Vari_Kotu_syudan = fldvalue.Trim
                            // 20160519 EXEUpdateに伴う修正 物件交通情報 -del end
                            case "車距離":
                                {
                                    model_cvitem.Vari_Car_kyori = fldvalue.Trim();
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,ensen_cnt) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件近隣駐車場情報(汎用ツールのみ)

    public class Bkdata_parkingother_Repository
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

                var model_cvitem = new Model.Bkdata_parkingother_Model();     // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

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
                // テーブル名/フィールド名セット
                string tblname_base = "bkdata";
                string tblname = "bkdata_parkingother";
                string fldnamegrp = "bk_guid,parking_no,naiyo,kyori,gak," + "zeikbn";

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
                            case "駐車場No":
                                {
                                    model_cvitem.Vari_Parking_no = fldvalue.Trim();
                                    break;
                                }
                            case "駐車場名":
                                {
                                    model_cvitem.Vari_Naiyo = fldvalue.Trim();
                                    break;
                                }
                            case "距離(m)":
                                {
                                    model_cvitem.Vari_Kyori = fldvalue.Trim();
                                    break;
                                }
                            case "料金(月額)":
                                {
                                    model_cvitem.Vari_Gak = fldvalue.Trim();
                                    break;
                                }
                            case "税区分":
                                {
                                    model_cvitem.Vari_Zeikbn = fldvalue.Trim();
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,parking_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件参照ファイル情報(汎用ツールのみ)

    public class Bkdata_relfile_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Bkdata_relfile_Model();          // 移行値格納用モデル初期化
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
                string tblname_base = "bkdata";
                string tblname = "bkdata_relfile";
                string fldnamegrp = "bk_guid,file_no,fullpath,addtime,biko," + "history";

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
                            case "物件NO":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,file_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件接道情報

    public class Bkdata_setudo_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Bkdata_setudo_Model();           // 移行値格納用モデル初期化
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
                string tblname_base = "bkdata";
                string tblname = "bkdata_setudo";
                string fldnamegrp = "bk_guid,setudo_cnt,setudo_muki,setudo_roadpattern,setudo_haba," + "setudo_setudokyori,setudo_pointselectflg,history";

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
                            case "物件NO":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
                            case "接道No":
                                {
                                    model_cvitem.Vari_Setudo_cnt = fldvalue.Trim();
                                    break;
                                }
                            case "方角":
                                {
                                    model_cvitem.Vari_Setudo_muki = fldvalue.Trim();
                                    break;
                                }
                            case "道路種":
                                {
                                    model_cvitem.Vari_Setudo_roadpattern = fldvalue.Trim();
                                    break;
                                }
                            case "幅員":
                                {
                                    model_cvitem.Vari_Setudo_haba = fldvalue.Trim();
                                    break;
                                }
                            case "接道距離":
                                {
                                    model_cvitem.Vari_Setudo_setudokyori = fldvalue.Trim();
                                    break;
                                }
                            case "位置指定道路":
                                {
                                    model_cvitem.Vari_Setudo_pointselectflg = fldvalue.Trim();
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,setudo_cnt) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件周辺情報

    public class Bkdata_syuhen_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Bkdata_syuhen_Model();           // 移行値格納用モデル初期化
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
                string tblname_base = "bkdata";
                string tblname = "bkdata_syuhen";
                string fldnamegrp = "syuhen_guid,bk_guid,syuhen_cnt,sisetu_kbn,sisetu_name," + "sisetu_kyori,history";

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
                            case "物件NO":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
                            case "周辺No":
                                {
                                    model_cvitem.Vari_Syuhen_cnt = fldvalue.Trim();
                                    break;
                                }
                            case "周辺施設分類":
                                {
                                    model_cvitem.Vari_Sisetu_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "周辺施設名":
                                {
                                    model_cvitem.Vari_Sisetu_name = fldvalue.Trim();
                                    break;
                                }
                            case "周辺施設距離":
                                {
                                    // 2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化 -chg sta
                                    // ↓↓↓旧srcコメントアウト↓↓↓
                                    // '2016.04.01 変換時の型落ち対応 -chg sta
                                    // 'Dim tmpsyuhenkyoristr As String = fldvalue.Trim
                                    // 'Dim tmpsyuhenkyoridbl As Double = Double.Parse(tmpsyuhenkyoristr)       'kakaka4 0322_1300 型落ち考慮
                                    // ''四捨五入して取得                               'kakaka2 四捨五入は汎用や他社用の為に用意？
                                    // '.Vari_Sisetu_kyori = Math.Round(tmpsyuhenkyoridbl, MidpointRounding.AwayFromZero).ToString
                                    // Dim tmpsyuhenkyoristr As String = fldvalue.Trim
                                    // Dim tmpsyuhenkyoridbl As Double = 0
                                    // If Double.TryParse(tmpsyuhenkyoristr, tmpsyuhenkyoridbl) Then
                                    // '四捨五入して取得
                                    // .Vari_Sisetu_kyori = Math.Round(tmpsyuhenkyoridbl, MidpointRounding.AwayFromZero).ToString
                                    // Else
                                    // '変換できない場合は値をそのまま格納 (ログ出力させる)
                                    // .Vari_Sisetu_kyori = tmpsyuhenkyoristr
                                    // End If
                                    // '2016.04.01 変換時の型落ち対応 -chg end
                                    // ↑↑↑旧srcコメントアウト↑↑↑
                                    model_cvitem.Vari_Sisetu_kyori = Conversions.ToString(EtcMethod.Get_RoundingValue(fldvalue.Trim()));
                                    break;
                                }
                                // 2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化 -chg end
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Syuhen_guid = Guid.NewGuid().ToString();
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,syuhen_cnt) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件修繕維持管理連絡先情報

    public class Bkdata_szeniji_Repository
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
                var hash_guid = new SafeDictionary<string, string>();                                  // guid格納用ハッシュテーブル
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用

                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.Bkdata_szeniji_Model();          // 移行値格納用モデル初期化
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
                string tblname_base = "bkdata";
                string tblname = "bkdata_szeniji";
                string fldnamegrp = "bk_guid,syuzenijikanri_no,syuzenijikanri_kasyo,syuzenijikanri_taisyokbn,syuzenijikanri_gyno," + "syuzenijikanri_simei,syuzenijikanri_address,syuzenijikanri_tel,syuzenijikanri_jisyano," + "syuzenijikanri_kasino,syuzenijikanri_syono,syuzenijikanri_simeiu";



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
                            case "物件NO":
                                {
                                    model_cvitem.Vari_Bk_guid = fldvalue.Trim();
                                    break;
                                }
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,syuzenijikanri_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region 物件メモ情報

    public class Bkdata_memo_Repository
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

                var model_cvitem = new Model.Bkdata_memo_Model();             // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列
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
                // テーブル名/フィールド名セット
                string tblname_base = "bkdata";
                string tblname = "bkdata_memo";
                string fldnamegrp = "bk_guid,memo_no,memo,history";

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

                // 移行対象件数取得用
                int cvtaisyocnt = 0;

                // ************************
                // 処理開始
                // ************************

                // キーヘッダー名取得
                string fldname_keymain = readtbl.Columns[keycol - 1].ColumnName.Trim();

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
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol - 1]).Trim();
                    model_cvitem.Vari_Bk_guid = tmp_keymain;

                    // 備考カウント初期化
                    int cvitemcnt = 0;
                    int tmp_cvrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                    // データ有無チェック
                    string tmp_fldvalueumuchk = "";
                    for (int cntjj = 1, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
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
                        string log_keyout = fldname_keymain + " = " + tmp_keymain;
                        // 親データ有無チェック
                        if (list_basekeydata.Contains(tmp_keymain) == false)
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "bkdata_memo-bk_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keymain))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "bkdata_memo-bk_guid";
                            string log_value = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        if (keychkflg)
                        {

                            // 移行値取得
                            for (int cntjj = 1, loopTo2 = readtbl.Columns.Count - 1; cntjj <= loopTo2; cntjj++)
                            {

                                // サブキー取得
                                string tmp_keysub = cntjj.ToString();

                                // 全キー取得
                                string fldvalue_key = tmp_keymain + "-" + tmp_keysub;

                                // ログ出力用データ格納(サブキーフィールド)
                                string fldname_keysub = readtbl.Columns[cntjj].ColumnName.Trim();
                                string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub;

                                // 登録値取得
                                string fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();
                                model_cvitem.Vari_Memo = fldvalue;

                                // 固定値
                                model_cvitem.Vari_Memo_no = cntjj.ToString();
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
                                        hash_cvitem["bk_guid"] = hash_guid[hash_cvitem["bk_guid"]];

                                        // 挿入処理
                                        CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);

                                        // 挿入処理が正常終了したレコードのキーを重複チェック用に格納
                                        if (normalflg)
                                        {
                                            if (list_chkduplicate.Contains(tmp_keymain) == false)
                                            {
                                                list_chkduplicate.Add(tmp_keymain);
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

                string tmp_sql = " SELECT CONVERT(varchar,bk_no) + '-' + CONVERT(varchar,memo_no) FROM " + tblname + " LEFT JOIN bkdata ON " + tblname + ".bk_guid = bkdata.bk_guid ";
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

        }

    }

    #endregion

    #region ×物件KeyValueテーブル

    // 汎用ツールのみ
    // 汎用は後で対応することとなったため今は対応しない

    #endregion

    #region ×物件画像情報

    // 画像は別途行うためここではコンバートしない

    #endregion

    #region ×物件周辺画像情報

    // 画像は別途行うためここではコンバートしない

    #endregion

    #region ×物件bkdata_youtube情報

    // 使用箇所が不明のため保留

    #endregion

}