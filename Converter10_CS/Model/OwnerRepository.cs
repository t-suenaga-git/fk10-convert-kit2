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

    #region 家主情報

    public class Owdata_Repository
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

                var model_cvitem = new Model.Owdata_Model();                  // 移行値格納用モデル初期化
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
                string tblname = "owdata";
                // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add (kaikei_jouhou,addr1,addr2を追加 ※kaikei_jouhouがもともと入っていなかったため同時に追加する)
                string fldnamegrp = "ow_no,kojinhojin_flg,ow_name,ow_namesjis,ow_kana," + "keisyo,post_code,addr_kenno,addr_sino,addr_cyo," + "addr_cyome,addr_cyomeptn,addr_banti,addr_etc,toukiaddr1," + "toukiaddr2,tel1,tel2,fax,mobiletel1," + "mobiletel2,mail,mobilemail,yusentel_kbn,yusenmail_kbn," + "birthday,nensyu,biko_kihon,gender,honseki," + "kinmu_name,kinmu_namesjis,kinmu_kana,kinmu_postcode,kinmu_addr1," + "kinmu_addr2,kinmu_tel1,kinmu_tel2,kinmu_fax,kinmu_gyosyu," + "kinmu_busyo,kinmu_nyuryokuym,kinmu_nyusyaym,url,gyosyu," + "daihyo_name,daihyo_namesjis,daihyo_kana,daihyo_yakusyoku,daihyo_atenaflg," + "tanto_name,tanto_namesjis,tanto_kana,tanto_busyo,tanto_atenaflg," + "sihonkin,jugyosu,nyuryoku_ym,torihikisaki,renraku_name," + "renraku_namesjis,renraku_kana,renraku_keisyo,renraku_postcode,renraku_addr1," + "renraku_addr2,renraku_tel1,renraku_tel2,renraku_fax,renraku_mobiletel1," + "renraku_mobiletel2,renraku_yusentelkbn,renraku_aidagara,biko_renraku,sofu_kbn," + "sofu_name,sofu_namesjis,sofu_kana,sofu_keisyo,sofu_postcode," + "sofu_addr1,sofu_addr2,sofu_tel1,sofu_tel2,sofu_fax," + "biko_sofu,event_tantono,useflg,findkeyword,history," + "rowid,kaikei_jouhou,addr1,addr2";


















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
                            case "家主No":
                                {
                                    model_cvitem.Vari_Ow_no = fldvalue.Trim();
                                    break;
                                }
                            case "個人法人フラグ":
                                {
                                    model_cvitem.Vari_Kojinhojin_flg = fldvalue.Trim();
                                    break;
                                }
                            case "家主名":
                                {
                                    model_cvitem.Vari_Ow_name = fldvalue.Trim();
                                    break;
                                }
                            case "家主名(SJIS)":
                                {
                                    model_cvitem.Vari_Ow_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "家主カナ":
                                {
                                    model_cvitem.Vari_Ow_kana = fldvalue.Trim();
                                    break;
                                }
                            case "宛名敬称":
                                {
                                    model_cvitem.Vari_Keisyo = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
                            // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -chg sta
                            // Case "住所1"
                            // tmp_address1 = fldvalue.Trim
                            // Case "住所2"
                            // tmp_address2 = fldvalue.Trim
                            case "住所1":
                                {
                                    model_cvitem.Vari_Addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "住所2":
                                {
                                    model_cvitem.Vari_Addr2 = fldvalue.Trim();
                                    break;
                                }
                            // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -chg end
                            case "登記住所１":
                                {
                                    model_cvitem.Vari_Toukiaddr1 = fldvalue.Trim();
                                    break;
                                }
                            case "登記住所２":
                                {
                                    model_cvitem.Vari_Toukiaddr2 = fldvalue.Trim();
                                    break;
                                }
                            case "TEL1":
                                {
                                    model_cvitem.Vari_Tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "TEL2":
                                {
                                    model_cvitem.Vari_Tel2 = fldvalue.Trim();
                                    break;
                                }
                            case "FAX":
                                {
                                    model_cvitem.Vari_Fax = fldvalue.Trim();
                                    break;
                                }
                            case "携帯１":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯２":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "メールアドレス":
                                {
                                    model_cvitem.Vari_Mail = fldvalue.Trim();
                                    break;
                                }
                            case "携帯メールアドレス":
                                {
                                    model_cvitem.Vari_Mobilemail = fldvalue.Trim();
                                    break;
                                }
                            case "優先電話設定":
                                {
                                    model_cvitem.Vari_Yusentel_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "優先メール設定":
                                {
                                    model_cvitem.Vari_Yusenmail_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "誕生日・設立日":
                                {
                                    model_cvitem.Vari_Birthday = fldvalue.Trim();
                                    break;
                                }
                            case "年収・年商":
                                {
                                    model_cvitem.Vari_Nensyu = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                            case "性別":
                                {
                                    model_cvitem.Vari_Gender = fldvalue.Trim();
                                    break;
                                }
                            case "本籍地":
                                {
                                    model_cvitem.Vari_Honseki = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先名":
                                {
                                    model_cvitem.Vari_Kinmu_name = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先名SJIS":
                                {
                                    model_cvitem.Vari_Kinmu_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先カナ":
                                {
                                    model_cvitem.Vari_Kinmu_kana = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先郵便番号":
                                {
                                    model_cvitem.Vari_Kinmu_postcode = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先住所１":
                                {
                                    model_cvitem.Vari_Kinmu_addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先住所２":
                                {
                                    model_cvitem.Vari_Kinmu_addr2 = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先TEL１":
                                {
                                    model_cvitem.Vari_Kinmu_tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先TEL２":
                                {
                                    model_cvitem.Vari_Kinmu_tel2 = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先FAX":
                                {
                                    model_cvitem.Vari_Kinmu_fax = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先業種":
                                {
                                    model_cvitem.Vari_Kinmu_gyosyu = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先部署":
                                {
                                    model_cvitem.Vari_Kinmu_busyo = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先情報記入年月":
                                {
                                    model_cvitem.Vari_Kinmu_nyuryokuym = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先入社年月":
                                {
                                    model_cvitem.Vari_Kinmu_nyusyaym = fldvalue.Trim();
                                    break;
                                }
                            case "URL":
                                {
                                    model_cvitem.Vari_Url = fldvalue.Trim();
                                    break;
                                }
                            case "業種":
                                {
                                    model_cvitem.Vari_Gyosyu = fldvalue.Trim();
                                    break;
                                }
                            case "代表者名":
                                {
                                    model_cvitem.Vari_Daihyo_name = fldvalue.Trim();
                                    break;
                                }
                            case "代表者名(SJIS)":
                                {
                                    model_cvitem.Vari_Daihyo_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "代表者カナ":
                                {
                                    model_cvitem.Vari_Daihyo_kana = fldvalue.Trim();
                                    break;
                                }
                            case "代表者役職":
                                {
                                    model_cvitem.Vari_Daihyo_yakusyoku = fldvalue.Trim();
                                    break;
                                }
                            case "代表者を宛先に含める":
                                {
                                    model_cvitem.Vari_Daihyo_atenaflg = fldvalue.Trim();
                                    break;
                                }
                            case "担当者名":
                                {
                                    model_cvitem.Vari_Tanto_name = fldvalue.Trim();
                                    break;
                                }
                            case "担当者名(SJIS)":
                                {
                                    model_cvitem.Vari_Tanto_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "担当者カナ":
                                {
                                    model_cvitem.Vari_Tanto_kana = fldvalue.Trim();
                                    break;
                                }
                            case "担当者部署":
                                {
                                    model_cvitem.Vari_Tanto_busyo = fldvalue.Trim();
                                    break;
                                }
                            case "担当者を宛先に含める":
                                {
                                    model_cvitem.Vari_Tanto_atenaflg = fldvalue.Trim();
                                    break;
                                }
                            case "資本金":
                                {
                                    model_cvitem.Vari_Sihonkin = fldvalue.Trim();
                                    break;
                                }
                            case "従業員数":
                                {
                                    model_cvitem.Vari_Jugyosu = fldvalue.Trim();
                                    break;
                                }
                            case "記入年月":
                                {
                                    model_cvitem.Vari_Nyuryoku_ym = fldvalue.Trim();
                                    break;
                                }
                            case "主要取引先":
                                {
                                    model_cvitem.Vari_Torihikisaki = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先名":
                                {
                                    model_cvitem.Vari_Renraku_name = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先名SJIS":
                                {
                                    model_cvitem.Vari_Renraku_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先カナ":
                                {
                                    model_cvitem.Vari_Renraku_kana = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先宛名敬称":
                                {
                                    model_cvitem.Vari_Renraku_keisyo = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先郵便番号":
                                {
                                    model_cvitem.Vari_Renraku_postcode = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先住所１":
                                {
                                    model_cvitem.Vari_Renraku_addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先住所２":
                                {
                                    model_cvitem.Vari_Renraku_addr2 = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先TEL１":
                                {
                                    model_cvitem.Vari_Renraku_tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先TEL２":
                                {
                                    model_cvitem.Vari_Renraku_tel2 = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先FAX":
                                {
                                    model_cvitem.Vari_Renraku_fax = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先携帯１":
                                {
                                    model_cvitem.Vari_Renraku_mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先携帯２":
                                {
                                    model_cvitem.Vari_Renraku_mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先優先設定(電話番号)":
                                {
                                    model_cvitem.Vari_Renraku_yusentelkbn = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先間柄":
                                {
                                    model_cvitem.Vari_Renraku_aidagara = fldvalue.Trim();
                                    break;
                                }
                            case "連絡先備考":
                                {
                                    model_cvitem.Vari_Biko_renraku = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先区分":
                                {
                                    model_cvitem.Vari_Sofu_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "送付先名":
                                {
                                    model_cvitem.Vari_Sofu_name = fldvalue.Trim();
                                    break;
                                }
                            case "送付先名(SJIS)":
                                {
                                    model_cvitem.Vari_Sofu_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "送付先カナ":
                                {
                                    model_cvitem.Vari_Sofu_kana = fldvalue.Trim();
                                    break;
                                }
                            case "送付先宛名敬称":
                                {
                                    model_cvitem.Vari_Sofu_keisyo = fldvalue.Trim();
                                    break;
                                }
                            case "送付先郵便番号":
                                {
                                    model_cvitem.Vari_Sofu_postcode = fldvalue.Trim();
                                    break;
                                }
                            case "送付先住所１":
                                {
                                    model_cvitem.Vari_Sofu_addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "送付先住所２":
                                {
                                    model_cvitem.Vari_Sofu_addr2 = fldvalue.Trim();
                                    break;
                                }
                            case "送付先TEL１":
                                {
                                    model_cvitem.Vari_Sofu_tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "送付先TEL2":
                                {
                                    model_cvitem.Vari_Sofu_tel2 = fldvalue.Trim();
                                    break;
                                }
                            case "送付先FAX":
                                {
                                    model_cvitem.Vari_Sofu_fax = fldvalue.Trim();
                                    break;
                                }
                            case "送付先備考":
                                {
                                    model_cvitem.Vari_Biko_sofu = fldvalue.Trim();
                                    break;
                                }
                            case "イベント担当者No":
                                {
                                    model_cvitem.Vari_Event_tantono = fldvalue.Trim();
                                    break;
                                }
                            case "絞り込みキーワード":
                                {
                                    model_cvitem.Vari_Findkeyword = fldvalue.Trim();
                                    break;
                                }
                            case "会計情報":
                                {
                                    model_cvitem.Vari_Kaikei_jouhou = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del sta
                    // '住所変換/格納
                    // Dim tmp_address As String = tmp_address1 & " " & tmp_address2
                    // Dim hash_address As New SafeDictionary<string, string>
                    // hash_address = AddressChange.Get_Address(tmp_address, sqlcnnv10)
                    // .Vari_Addr_kenno = hash_address("ken_no")
                    // .Vari_Addr_sino = hash_address("si_no")
                    // .Vari_Addr_cyo = hash_address("mati")
                    // .Vari_Addr_cyome = hash_address("cyome")
                    // .Vari_Addr_cyomeptn = hash_address("chomeptn")
                    // .Vari_Addr_banti = hash_address("banti")
                    // .Vari_Addr_etc = hash_address("etc")
                    // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del end

                    // 20160706 家主個人法人区分の取得方法変更 -add sta
                    // 個人法人区分が個人に設定されている場合は名称から判別する(法人に設定されているデータはそのまま移行)
                    if (model_cvitem.Vari_Kojinhojin_flg == "1")
                    {
                        model_cvitem.Vari_Kojinhojin_flg = Get_KojinHojinvalue(model_cvitem.Vari_Ow_name).ToString();
                    }
                    // 20160706 家主個人法人区分の取得方法変更 -add end

                    // 固定値
                    model_cvitem.Vari_Useflg = 1.ToString();
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

            // ''' <summary>
            // ''' 【中間ファイル→変数】
            // ''' </summary>
            // ''' <param name="sqlcnnv10"></param>
            // ''' <param name="syorikomok"></param>
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

            // Dim model_cvitem As New Njc.Model.Owdata_Model                  '移行値格納用モデル初期化
            // Dim keycol As Integer = 1                                       'キー列

            // '************************
            // '作業準備
            // '************************

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
            // If Not rtn Then
            // Return rtn
            // End If

            // 'テーブル名/フィールド名セット
            // Dim tblname As String = "owdata"
            // '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add (kaikei_jouhou,addr1,addr2を追加 ※kaikei_jouhouがもともと入っていなかったため同時に追加する)
            // Dim fldnamegrp As String = "ow_no,kojinhojin_flg,ow_name,ow_namesjis,ow_kana," & _
            // "keisyo,post_code,addr_kenno,addr_sino,addr_cyo," & _
            // "addr_cyome,addr_cyomeptn,addr_banti,addr_etc,toukiaddr1," & _
            // "toukiaddr2,tel1,tel2,fax,mobiletel1," & _
            // "mobiletel2,mail,mobilemail,yusentel_kbn,yusenmail_kbn," & _
            // "birthday,nensyu,biko_kihon,gender,honseki," & _
            // "kinmu_name,kinmu_namesjis,kinmu_kana,kinmu_postcode,kinmu_addr1," & _
            // "kinmu_addr2,kinmu_tel1,kinmu_tel2,kinmu_fax,kinmu_gyosyu," & _
            // "kinmu_busyo,kinmu_nyuryokuym,kinmu_nyusyaym,url,gyosyu," & _
            // "daihyo_name,daihyo_namesjis,daihyo_kana,daihyo_yakusyoku,daihyo_atenaflg," & _
            // "tanto_name,tanto_namesjis,tanto_kana,tanto_busyo,tanto_atenaflg," & _
            // "sihonkin,jugyosu,nyuryoku_ym,torihikisaki,renraku_name," & _
            // "renraku_namesjis,renraku_kana,renraku_keisyo,renraku_postcode,renraku_addr1," & _
            // "renraku_addr2,renraku_tel1,renraku_tel2,renraku_fax,renraku_mobiletel1," & _
            // "renraku_mobiletel2,renraku_yusentelkbn,renraku_aidagara,biko_renraku,sofu_kbn," & _
            // "sofu_name,sofu_namesjis,sofu_kana,sofu_keisyo,sofu_postcode," & _
            // "sofu_addr1,sofu_addr2,sofu_tel1,sofu_tel2,sofu_fax," & _
            // "biko_sofu,event_tantono,useflg,findkeyword,history," & _
            // "rowid,kaikei_jouhou,addr1,addr2"

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

            // '---------------
            // 'ヘッダー処理
            // '---------------
            // 'ヘッダー行取得
            // Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            // 'キーヘッダー名取得
            // Dim fldname_key As String = headervalue(startrow - 1, keycol)

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
            // Dim fldvalue_key As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))

            // 'ログ出力用
            // Dim str_logkey As String = fldname_key & " = " & fldvalue_key

            // '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del sta
            // ''作業用変数
            // 'Dim tmp_address1 As String = ""
            // 'Dim tmp_address2 As String = ""
            // '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del end

            // '移行値取得
            // For cntjj = 1 To columncnt

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "家主No"
            // .Vari_Ow_no = fldvalue.Trim
            // Case "個人法人フラグ"
            // .Vari_Kojinhojin_flg = fldvalue.Trim
            // Case "家主名"
            // .Vari_Ow_name = fldvalue.Trim
            // Case "家主名(SJIS)"
            // .Vari_Ow_namesjis = fldvalue.Trim
            // Case "家主カナ"
            // .Vari_Ow_kana = fldvalue.Trim
            // Case "宛名敬称"
            // .Vari_Keisyo = fldvalue.Trim
            // Case "郵便番号"
            // .Vari_Post_code = fldvalue.Trim
            // '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -chg sta
            // 'Case "住所1"
            // '    tmp_address1 = fldvalue.Trim
            // 'Case "住所2"
            // '    tmp_address2 = fldvalue.Trim
            // Case "住所1"
            // .Vari_Addr1 = fldvalue.Trim
            // Case "住所2"
            // .Vari_Addr2 = fldvalue.Trim
            // '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -chg end
            // Case "登記住所１"
            // .Vari_Toukiaddr1 = fldvalue.Trim
            // Case "登記住所２"
            // .Vari_Toukiaddr2 = fldvalue.Trim
            // Case "TEL1"
            // .Vari_Tel1 = fldvalue.Trim
            // Case "TEL2"
            // .Vari_Tel2 = fldvalue.Trim
            // Case "FAX"
            // .Vari_Fax = fldvalue.Trim
            // Case "携帯１"
            // .Vari_Mobiletel1 = fldvalue.Trim
            // Case "携帯２"
            // .Vari_Mobiletel2 = fldvalue.Trim
            // Case "メールアドレス"
            // .Vari_Mail = fldvalue.Trim
            // Case "携帯メールアドレス"
            // .Vari_Mobilemail = fldvalue.Trim
            // Case "優先電話設定"
            // .Vari_Yusentel_kbn = fldvalue.Trim
            // Case "優先メール設定"
            // .Vari_Yusenmail_kbn = fldvalue.Trim
            // Case "誕生日・設立日"
            // .Vari_Birthday = fldvalue.Trim
            // Case "年収・年商"
            // .Vari_Nensyu = fldvalue.Trim
            // Case "備考"
            // .Vari_Biko_kihon = fldvalue.Trim
            // Case "性別"
            // .Vari_Gender = fldvalue.Trim
            // Case "本籍地"
            // .Vari_Honseki = fldvalue.Trim
            // Case "勤務先名"
            // .Vari_Kinmu_name = fldvalue.Trim
            // Case "勤務先名SJIS"
            // .Vari_Kinmu_namesjis = fldvalue.Trim
            // Case "勤務先カナ"
            // .Vari_Kinmu_kana = fldvalue.Trim
            // Case "勤務先郵便番号"
            // .Vari_Kinmu_postcode = fldvalue.Trim
            // Case "勤務先住所１"
            // .Vari_Kinmu_addr1 = fldvalue.Trim
            // Case "勤務先住所２"
            // .Vari_Kinmu_addr2 = fldvalue.Trim
            // Case "勤務先TEL１"
            // .Vari_Kinmu_tel1 = fldvalue.Trim
            // Case "勤務先TEL２"
            // .Vari_Kinmu_tel2 = fldvalue.Trim
            // Case "勤務先FAX"
            // .Vari_Kinmu_fax = fldvalue.Trim
            // Case "勤務先業種"
            // .Vari_Kinmu_gyosyu = fldvalue.Trim
            // Case "勤務先部署"
            // .Vari_Kinmu_busyo = fldvalue.Trim
            // Case "勤務先情報記入年月"
            // .Vari_Kinmu_nyuryokuym = fldvalue.Trim
            // Case "勤務先入社年月"
            // .Vari_Kinmu_nyusyaym = fldvalue.Trim
            // Case "URL"
            // .Vari_Url = fldvalue.Trim
            // Case "業種"
            // .Vari_Gyosyu = fldvalue.Trim
            // Case "代表者名"
            // .Vari_Daihyo_name = fldvalue.Trim
            // Case "代表者名(SJIS)"
            // .Vari_Daihyo_namesjis = fldvalue.Trim
            // Case "代表者カナ"
            // .Vari_Daihyo_kana = fldvalue.Trim
            // Case "代表者役職"
            // .Vari_Daihyo_yakusyoku = fldvalue.Trim
            // Case "代表者を宛先に含める"
            // .Vari_Daihyo_atenaflg = fldvalue.Trim
            // Case "担当者名"
            // .Vari_Tanto_name = fldvalue.Trim
            // Case "担当者名(SJIS)"
            // .Vari_Tanto_namesjis = fldvalue.Trim
            // Case "担当者カナ"
            // .Vari_Tanto_kana = fldvalue.Trim
            // Case "担当者部署"
            // .Vari_Tanto_busyo = fldvalue.Trim
            // Case "担当者を宛先に含める"
            // .Vari_Tanto_atenaflg = fldvalue.Trim
            // Case "資本金"
            // .Vari_Sihonkin = fldvalue.Trim
            // Case "従業員数"
            // .Vari_Jugyosu = fldvalue.Trim
            // Case "記入年月"
            // .Vari_Nyuryoku_ym = fldvalue.Trim
            // Case "主要取引先"
            // .Vari_Torihikisaki = fldvalue.Trim
            // Case "連絡先名"
            // .Vari_Renraku_name = fldvalue.Trim
            // Case "連絡先名SJIS"
            // .Vari_Renraku_namesjis = fldvalue.Trim
            // Case "連絡先カナ"
            // .Vari_Renraku_kana = fldvalue.Trim
            // Case "連絡先宛名敬称"
            // .Vari_Renraku_keisyo = fldvalue.Trim
            // Case "連絡先郵便番号"
            // .Vari_Renraku_postcode = fldvalue.Trim
            // Case "連絡先住所１"
            // .Vari_Renraku_addr1 = fldvalue.Trim
            // Case "連絡先住所２"
            // .Vari_Renraku_addr2 = fldvalue.Trim
            // Case "連絡先TEL１"
            // .Vari_Renraku_tel1 = fldvalue.Trim
            // Case "連絡先TEL２"
            // .Vari_Renraku_tel2 = fldvalue.Trim
            // Case "連絡先FAX"
            // .Vari_Renraku_fax = fldvalue.Trim
            // Case "連絡先携帯１"
            // .Vari_Renraku_mobiletel1 = fldvalue.Trim
            // Case "連絡先携帯２"
            // .Vari_Renraku_mobiletel2 = fldvalue.Trim
            // Case "連絡先優先設定(電話番号)"
            // .Vari_Renraku_yusentelkbn = fldvalue.Trim
            // Case "連絡先間柄"
            // .Vari_Renraku_aidagara = fldvalue.Trim
            // Case "連絡先備考"
            // .Vari_Biko_renraku = fldvalue.Trim
            // Case "書類送付先区分"
            // .Vari_Sofu_kbn = fldvalue.Trim
            // Case "送付先名"
            // .Vari_Sofu_name = fldvalue.Trim
            // Case "送付先名(SJIS)"
            // .Vari_Sofu_namesjis = fldvalue.Trim
            // Case "送付先カナ"
            // .Vari_Sofu_kana = fldvalue.Trim
            // Case "送付先宛名敬称"
            // .Vari_Sofu_keisyo = fldvalue.Trim
            // Case "送付先郵便番号"
            // .Vari_Sofu_postcode = fldvalue.Trim
            // Case "送付先住所１"
            // .Vari_Sofu_addr1 = fldvalue.Trim
            // Case "送付先住所２"
            // .Vari_Sofu_addr2 = fldvalue.Trim
            // Case "送付先TEL１"
            // .Vari_Sofu_tel1 = fldvalue.Trim
            // Case "送付先TEL2"
            // .Vari_Sofu_tel2 = fldvalue.Trim
            // Case "送付先FAX"
            // .Vari_Sofu_fax = fldvalue.Trim
            // Case "送付先備考"
            // .Vari_Biko_sofu = fldvalue.Trim
            // Case "イベント担当者No"
            // .Vari_Event_tantono = fldvalue.Trim
            // Case "絞り込みキーワード"
            // .Vari_Findkeyword = fldvalue.Trim
            // Case "会計情報"
            // .Vari_Kaikei_jouhou = fldvalue.Trim
            // End Select

            // Next

            // '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del sta
            // ''住所変換/格納
            // 'Dim tmp_address As String = tmp_address1 & " " & tmp_address2
            // 'Dim hash_address As New SafeDictionary<string, string>
            // 'hash_address = AddressChange.Get_Address(tmp_address, sqlcnnv10)
            // '.Vari_Addr_kenno = hash_address("ken_no")
            // '.Vari_Addr_sino = hash_address("si_no")
            // '.Vari_Addr_cyo = hash_address("mati")
            // '.Vari_Addr_cyome = hash_address("cyome")
            // '.Vari_Addr_cyomeptn = hash_address("chomeptn")
            // '.Vari_Addr_banti = hash_address("banti")
            // '.Vari_Addr_etc = hash_address("etc")
            // '20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -del end

            // '20160706 家主個人法人区分の取得方法変更 -add sta
            // '個人法人区分が個人に設定されている場合は名称から判別する(法人に設定されているデータはそのまま移行)
            // If .Vari_Kojinhojin_flg = "1" Then
            // .Vari_Kojinhojin_flg = Me.Get_KojinHojinvalue(.Vari_Ow_name)
            // End If
            // '20160706 家主個人法人区分の取得方法変更 -add end

            // '固定値
            // .Vari_Useflg = 1
            // .Vari_History = DefHistory
            // .Vari_Rowid = Guid.NewGuid.ToString

            // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            // 'データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
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

                string tmp_sql = " SELECT ow_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 家主の名称に特定の文字列が含まれている場合は個人法人区分を法人にして移行する '20160706 家主個人法人区分の取得方法変更
            /// </summary>
            /// <param name="owname"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public int Get_KojinHojinvalue(string owname)
            {

                int rtn_int = 1;
                var list_hojinname = new List<string>() { "㈱", "(株)", "株式", "㈲", "(有)", "有限", "(資)", "合資", "(名)", "合名", "法人", "ホーム", "サービス", "工務店", "土地", "建物", "産業", "不動産", "企画" };

                foreach (var hojinname in list_hojinname)
                {

                    if ((owname ?? "") != (owname.Replace(hojinname, "") ?? ""))
                    {
                        rtn_int = 2;
                        return rtn_int;
                    }

                }

                return rtn_int;

            }

        }

    }

    #endregion

    #region 家主口座情報

    public class Owdata_koza_Repository
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

                var model_cvitem = new Model.Owdata_koza_Model();             // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub = 2;                                   // サブキー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                Get_RelData_KozaSyubetu();

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
                string tblname_base = "owdata";
                string tblname = "owdata_koza";
                string fldnamegrp = "ow_no,ow_kozano,kinyu_no,kinyu_tenno,koza_syubetu," + "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," + "yucyokoza_bango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no," + "sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai," + "history,rowid,koza_printkbn";




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
                            case "家主No":
                                {
                                    model_cvitem.Vari_Ow_no = fldvalue.Trim();
                                    break;
                                }
                            case "口座No":
                                {
                                    model_cvitem.Vari_Ow_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関No":
                                {
                                    model_cvitem.Vari_Kinyu_no = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関支店No":
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
                            case "口座備考":
                                {
                                    model_cvitem.Vari_Biko_koza = fldvalue.Trim();
                                    break;
                                }
                            case "振込情報備考":
                                {
                                    model_cvitem.Vari_Biko_furikomi = fldvalue.Trim();
                                    break;
                                }
                            case "総合振込区分":
                                {
                                    model_cvitem.Vari_Sgfirai_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "振込依頼人No":
                                {
                                    model_cvitem.Vari_Sgfirai_no = fldvalue.Trim();
                                    break;
                                }
                            case "振込手数料負担区分":
                                {
                                    model_cvitem.Vari_Sgfirai_tesufutankbn = fldvalue.Trim();
                                    break;
                                }
                            case "振込手数料計算区分":
                                {
                                    model_cvitem.Vari_Sgfirai_tesukeisankbn = fldvalue.Trim();
                                    break;
                                }
                            case "振込手数料固定額1":
                                {
                                    model_cvitem.Vari_Sgfirai_tesukotei1gak = fldvalue.Trim();
                                    break;
                                }
                            case "振込手数料固定額2":
                                {
                                    model_cvitem.Vari_Sgfirai_tesukotei2gak = fldvalue.Trim();
                                    break;
                                }
                            case "総合振込情報備考":
                                {
                                    model_cvitem.Vari_Biko_sgfirai = fldvalue.Trim();
                                    break;
                                }
                            case "家主向け帳票の表示":
                                {
                                    model_cvitem.Vari_Koza_printkbn = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();     // GUID

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

            // ''' <summary>
            // ''' 【中間ファイル→変数】
            // ''' </summary>
            // ''' <param name="sqlcnnv10"></param>
            // ''' <param name="syorikomok"></param>
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

            // Dim model_cvitem As New Njc.Model.Owdata_koza_Model             '移行値格納用モデル初期化
            // Dim keycol_main As Integer = 1                                  'メインキー列
            // Dim keycol_sub As Integer = 2                                   'サブキー列

            // '************************
            // '作業準備
            // '************************

            // '紐付けデータ取得
            // Call Me.Get_RelData_KozaSyubetu()

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
            // If Not rtn Then
            // Return rtn
            // End If

            // 'テーブル名/フィールド名セット
            // Dim tblname_base As String = "owdata"
            // Dim tblname As String = "owdata_koza"
            // Dim fldnamegrp As String = "ow_no,ow_kozano,kinyu_no,kinyu_tenno,koza_syubetu," & _
            // "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," & _
            // "yucyokoza_bango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no," & _
            // "sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai," & _
            // "history,rowid,koza_printkbn"

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
            // Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            // 'キーヘッダー名取得
            // Dim fldname_keymain As String = headervalue(startrow - 1, keycol_main)
            // Dim fldname_keysub As String = headervalue(startrow - 1, keycol_sub)

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
            // Dim fldvalue_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_main))
            // Dim fldvalue_keysub As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol_sub))
            // Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

            // 'ログ出力用
            // Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

            // '移行値取得
            // For cntjj = 1 To columncnt

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "家主No"
            // .Vari_Ow_no = fldvalue.Trim
            // Case "口座No"
            // .Vari_Ow_kozano = fldvalue.Trim
            // Case "金融機関No"
            // .Vari_Kinyu_no = fldvalue.Trim
            // Case "金融機関支店No"
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
            // Case "口座備考"
            // .Vari_Biko_koza = fldvalue.Trim
            // Case "振込情報備考"
            // .Vari_Biko_furikomi = fldvalue.Trim
            // Case "総合振込区分"
            // .Vari_Sgfirai_kbn = fldvalue.Trim
            // Case "振込依頼人No"
            // .Vari_Sgfirai_no = fldvalue.Trim
            // Case "振込手数料負担区分"
            // .Vari_Sgfirai_tesufutankbn = fldvalue.Trim
            // Case "振込手数料計算区分"
            // .Vari_Sgfirai_tesukeisankbn = fldvalue.Trim
            // Case "振込手数料固定額1"
            // .Vari_Sgfirai_tesukotei1gak = fldvalue.Trim
            // Case "振込手数料固定額2"
            // .Vari_Sgfirai_tesukotei2gak = fldvalue.Trim
            // Case "総合振込情報備考"
            // .Vari_Biko_sgfirai = fldvalue.Trim
            // Case "家主向け帳票の表示"
            // .Vari_Koza_printkbn = fldvalue.Trim
            // End Select

            // Next

            // '固定値
            // .Vari_History = DefHistory
            // .Vari_Rowid = Guid.NewGuid.ToString     'GUID

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

            /// <summary>
            /// 親マスタ取得→リスト格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT ow_no FROM " + tblname;
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

    #region 家主イベント情報

    public class Owdata_event_Repository
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

                        // 書込処理
                        if (!skipflg)
                        {

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

                    // --------------------------------------------------------------------
                    // 移行した備考が1データ以上ある場合移行したレコードの数を更新する
                    // --------------------------------------------------------------------
                    tmp_cvcnt = tmp_cvcnt + 1;

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

                string tmp_sql = " SELECT ow_no FROM " + tblname;
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

                string tmp_sql = " SELECT ow_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 家主メモ情報

    public class Owdata_memo_Repository
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

                var model_cvitem = new Model.Owdata_memo_Model();             // 移行値格納用モデル初期化
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
                string tblname_base = "owdata";
                string tblname = "owdata_memo";
                string fldnamegrp = "ow_no,memo_no,memo,history";

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

                // 移行対象件数取得用
                int cvtaisyocnt = 0;  // 20161007 メモ関連のログ出力修正 -add

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
                    bool cvflg = true;         // 20161007 メモ関連のログ出力修正 -add
                    bool keychkflg = true;     // 20161007 メモ関連のログ出力修正 -add

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol - 1]).Trim();
                    model_cvitem.Vari_Ow_no = tmp_keymain;

                    // 備考カウント初期化
                    int cvitemcnt = 0;
                    int tmp_cvrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                    // 20161007 メモ関連のログ出力修正 -chg sta
                    // '移行値取得
                    // For cntjj = 1 To readtbl.Columns.Count - 1

                    // 'サブキー取得
                    // Dim tmp_keysub As String = (cntjj).ToString

                    // '全キー取得
                    // Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

                    // 'ログ出力用データ格納(サブキーフィールド)
                    // Dim fldname_keysub As String = readtbl.Columns(cntjj).ColumnName.Trim
                    // Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub

                    // '登録値取得
                    // Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
                    // .Vari_Memo = fldvalue

                    // '固定値
                    // .Vari_Memo_no = (cntjj).ToString
                    // .Vari_History = DefHistory

                    // 'メモにデータが存在する場合に書込処理を行う
                    // If .Vari_Memo <> "" Then

                    // '20160928 メモ関連の移行件数表示修正 -add
                    // tmp_cvrowcnt = tmp_cvrowcnt + 1

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
                    // cvitemcnt = cvitemcnt + 1
                    // End If

                    // End If

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

                    // Next

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
                            string log_key = "owdata_memo-ow_no";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keymain))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "owdata_memo-ow_no";
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
                    // 20161007 メモ関連のログ出力修正 -chg end

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
                // 20161007 メモ関連のログ出力修正 -chg sta
                // '20160928 メモ関連の移行件数表示修正 -chg sta
                // 'midrowcnt = rowcnt
                // midrowcnt = totalrowcnt
                // '20160928 メモ関連の移行件数表示修正 -chg end
                midrowcnt = cvtaisyocnt;
                // 20161007 メモ関連のログ出力修正 -chg end

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
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT ow_no FROM " + tblname;
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

                string tmp_sql = " SELECT ow_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

}