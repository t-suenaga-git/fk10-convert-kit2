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

    #region 契約者情報

    public class Kysdata_Repository
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

                var model_cvitem = new Model.Kysdata_Model();                 // 移行値格納用モデル初期化
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
                string tblname = "kysdata";
                // 20160519 EXEUpdateに伴う修正 契約者基本情報 -chg sta
                // Dim fldnamegrp As String = "kys_no,kys_name,kys_namesjis,kys_kana,kojinhojin_flg," & _
                // "keisyo,post_code,addr1,addr2,tel1," & _
                // "tel2,fax,mobiletel1,mobiletel2,mail," & _
                // "mobilemail,yusentel_kbn,yusenmail_kbn,birthday,nensyu," & _
                // "biko_kihon,gender,honseki,kinmu_name,kinmu_namesjis," & _
                // "kinmu_kana,kinmu_postcode,kinmu_addr1,kinmu_addr2,kinmu_tel1," & _
                // "kinmu_tel2,kinmu_fax,kinmu_gyosyu,kinmu_busyo,kinmu_nyuryokuym," & _
                // "kinmu_nyusyaym,nyukyomae_nyuryokuym,nyukyomae_postcode,nyukyomae_addr1,nyukyomae_addr2," & _
                // "nyukyomae_tel1,nyukyomae_tel2,nyukyomae_fax,url,gyosyu," & _
                // "daihyo_name,daihyo_namesjis,daihyo_kana,daihyo_yakusyoku,daihyo_atenaflg," & _
                // "tanto_name,tanto_namesjis,tanto_kana,tanto_busyo,tanto_yakusyoku," & _
                // "tanto_atenaflg,sihonkin,jugyosu,nyuryoku_ym,torihikisaki," & _
                // "renraku_name,renraku_namesjis,renraku_kana,renraku_keisyo,renraku_postcode," & _
                // "renraku_addr1,renraku_addr2,renraku_tel1,renraku_tel2,renraku_fax," & _
                // "renraku_mobiletel1,renraku_mobiletel2,renraku_yusentelkbn,renraku_aidagara,biko_renraku," & _
                // "sofu_kbn,sofu_name,sofu_namesjis,sofu_kana,sofu_keisyo," & _
                // "sofu_postcode,sofu_addr1,sofu_addr2,sofu_tel1,sofu_tel2," & _
                // "sofu_fax,biko_sofu,sofu_hakkofurikomi,sofu_hakkofurikae,sofu_hakkotokusoku," & _
                // "furikomituti_bktaniflg,furikaetuti_bktaniflg,furikomi_kasouseflg,hosyo_multipleflg,furikae_hosyoumu," & _
                // "useflg,history,rowid,cvpay_umu"
                string fldnamegrp = "kys_no,kys_name,kys_namesjis,kys_kana,kojinhojin_flg," + "keisyo,post_code,addr1,addr2,tel1," + "tel2,fax,mobiletel1,mobiletel2,mail," + "mobilemail,yusentel_kbn,yusenmail_kbn,birthday,nensyu," + "biko_kihon,gender,honseki,kinmu_name,kinmu_namesjis," + "kinmu_kana,kinmu_postcode,kinmu_addr1,kinmu_addr2,kinmu_tel1," + "kinmu_tel2,kinmu_fax,kinmu_gyosyu,kinmu_busyo,kinmu_nyuryokuym," + "kinmu_nyusyaym,nyukyomae_nyuryokuym,nyukyomae_postcode,nyukyomae_addr1,nyukyomae_addr2," + "nyukyomae_tel1,nyukyomae_tel2,nyukyomae_fax,url,gyosyu," + "daihyo_name,daihyo_namesjis,daihyo_kana,daihyo_yakusyoku,tanto_name," + "tanto_namesjis,tanto_kana,tanto_busyo,tanto_yakusyoku,tanto_atenaflg," + "sihonkin,jugyosu,nyuryoku_ym,torihikisaki,renraku_name," + "renraku_namesjis,renraku_kana,renraku_keisyo,renraku_postcode,renraku_addr1," + "renraku_addr2,renraku_tel1,renraku_tel2,renraku_fax,renraku_mobiletel1," + "renraku_mobiletel2,renraku_yusentelkbn,renraku_aidagara,biko_renraku,sofu_kbn," + "sofu_name,sofu_namesjis,sofu_kana,sofu_keisyo,sofu_postcode," + "sofu_addr1,sofu_addr2,sofu_tel1,sofu_tel2,sofu_fax," + "biko_sofu,sofu_hakkofurikomi,sofu_hakkofurikae,sofu_hakkotokusoku,furikomi_kasouseflg," + "hosyo_multipleflg,furikae_hosyoumu,useflg,history,rowid," + "cvpay_umu,daihyo_atenaflg";


















                // 20160519 EXEUpdateに伴う修正 契約者基本情報 -chg end

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
                    string tmp_tantobusyo = "";
                    string tmp_yakusyoku = "";

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
                            case "契約者No":
                                {
                                    model_cvitem.Vari_Kys_no = fldvalue.Trim();
                                    break;
                                }
                            case "契約者名":
                                {
                                    model_cvitem.Vari_Kys_name = fldvalue.Trim();
                                    break;
                                }
                            case "契約者名SJIS":
                                {
                                    model_cvitem.Vari_Kys_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "契約者カナ":
                                {
                                    model_cvitem.Vari_Kys_kana = fldvalue.Trim();
                                    break;
                                }
                            case "個人法人フラグ":
                                {
                                    model_cvitem.Vari_Kojinhojin_flg = fldvalue.Trim();
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
                            case "住所１":
                                {
                                    model_cvitem.Vari_Addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "住所２":
                                {
                                    model_cvitem.Vari_Addr2 = fldvalue.Trim();
                                    break;
                                }
                            case "TEL１":
                                {
                                    model_cvitem.Vari_Tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "TEL２":
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
                            case "優先設定(電話番号)":
                                {
                                    model_cvitem.Vari_Yusentel_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "優先設定(メールアドレス)":
                                {
                                    model_cvitem.Vari_Yusenmail_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "生年月日・設立日":
                                {
                                    model_cvitem.Vari_Birthday = fldvalue.Trim();
                                    break;
                                }
                            case "年収・年商":
                                {
                                    model_cvitem.Vari_Nensyu = Conversions.ToString(EtcMethod.Get_RoundingValue(fldvalue.Trim()));
                                    break;
                                }
                            case "備考(基本情報)":
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
                            case "入居前連絡先情報記入年月":
                                {
                                    model_cvitem.Vari_Nyukyomae_nyuryokuym = fldvalue.Trim();
                                    break;
                                }
                            case "入居前連絡先郵便番号":
                                {
                                    model_cvitem.Vari_Nyukyomae_postcode = fldvalue.Trim();
                                    break;
                                }
                            case "入居前連絡先住所１":
                                {
                                    model_cvitem.Vari_Nyukyomae_addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "入居前連絡先住所２":
                                {
                                    model_cvitem.Vari_Nyukyomae_addr2 = fldvalue.Trim();
                                    break;
                                }
                            case "入居前連絡先TEL１":
                                {
                                    model_cvitem.Vari_Nyukyomae_tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "入居前連絡先TEL２":
                                {
                                    model_cvitem.Vari_Nyukyomae_tel2 = fldvalue.Trim();
                                    break;
                                }
                            case "入居前連絡先FAX":
                                {
                                    model_cvitem.Vari_Nyukyomae_fax = fldvalue.Trim();
                                    break;
                                }
                            case "Webアドレス(URL)":
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
                            case "代表者名SJIS":
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
                            case "担当者名SJIS":
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
                                    tmp_tantobusyo = fldvalue.Trim();
                                    break;
                                }
                            case "担当者役職":
                                {
                                    tmp_yakusyoku = fldvalue.Trim();
                                    break;
                                }
                            case "担当者を宛先に含める":
                                {
                                    model_cvitem.Vari_Tanto_atenaflg = fldvalue.Trim();
                                    break;
                                }
                            case "資本金":
                                {
                                    model_cvitem.Vari_Sihonkin = Conversions.ToString(EtcMethod.Get_RoundingValue(fldvalue.Trim()));
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
                            case "書類送付先名":
                                {
                                    model_cvitem.Vari_Sofu_name = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先名SJIS":
                                {
                                    model_cvitem.Vari_Sofu_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先カナ":
                                {
                                    model_cvitem.Vari_Sofu_kana = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先宛名敬称":
                                {
                                    model_cvitem.Vari_Sofu_keisyo = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先郵便番号":
                                {
                                    model_cvitem.Vari_Sofu_postcode = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先住所１":
                                {
                                    model_cvitem.Vari_Sofu_addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先住所２":
                                {
                                    model_cvitem.Vari_Sofu_addr2 = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先TEL１":
                                {
                                    model_cvitem.Vari_Sofu_tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先TEL２":
                                {
                                    model_cvitem.Vari_Sofu_tel2 = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先FAX":
                                {
                                    model_cvitem.Vari_Sofu_fax = fldvalue.Trim();
                                    break;
                                }
                            case "書類送付先備考":
                                {
                                    model_cvitem.Vari_Biko_sofu = fldvalue.Trim();
                                    break;
                                }
                            case "振込通知書発行の可否":
                                {
                                    model_cvitem.Vari_Sofu_hakkofurikomi = fldvalue.Trim();
                                    break;
                                }
                            case "振替通知書発行の可否":
                                {
                                    model_cvitem.Vari_Sofu_hakkofurikae = fldvalue.Trim();
                                    break;
                                }
                            case "督促状発行の可否":
                                {
                                    model_cvitem.Vari_Sofu_hakkotokusoku = fldvalue.Trim();
                                    break;
                                }
                            case "口座振込仮想口座使用フラグ":
                                {
                                    model_cvitem.Vari_Furikomi_kasouseflg = fldvalue.Trim();
                                    break;
                                }
                            case "保証人複数フラグ":
                                {
                                    model_cvitem.Vari_Hosyo_multipleflg = fldvalue.Trim();
                                    break;
                                }
                            case "コンビニ収納サービス利用有無":
                                {
                                    model_cvitem.Vari_Cvpay_umu = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 担当部署役職を結合して格納
                    model_cvitem.Vari_Tanto_busyo = (tmp_tantobusyo + " " + tmp_yakusyoku).Trim();

                    // 固定値
                    model_cvitem.Vari_Furikae_hosyoumu = "2";    // 京王カスタマイズフィールド
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    string errstr = "";
                    var hash_log = new SafeDictionary<string, string>();                                                           // kakaka ↓親マスタチェック("list_basekeydata")が不要な場合は、わかり易いようにした方がよいかと(ここでは親チェックは不要とわかるように)
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (skipflg == false)                                                                     // kakaka 京王様ソースでは、途中仕様変更などでNotを使用してましたが、出来るだけTrueFalseで記載下さい(1つ前のコードで取得値のNotをSkipFlgにセットし、そのSkipFlgのNotを条件とする…、、、間違い防止)
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

                string tmp_sql = " SELECT kys_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 契約者口座情報

    public class Kysdata_koza_Repository
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

                var model_cvitem = new Model.Kysdata_koza_Model();            // 移行値格納用モデル初期化
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
                string tblname_base = "kysdata";
                string tblname = "kysdata_koza";
                // 20161012 革命10アップデートに伴う修正 -add(最後尾にkoza_yucyoflgを追加)
                // 20160519 EXEUpdateに伴う修正 契約者口座情報 -add(最後尾にkoza_printkbnを追加)
                string fldnamegrp = "kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno," + "koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1," + "yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango," + "biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn," + "sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn," + "koza_yucyoflg";





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
                            case "契約者No":
                                {
                                    model_cvitem.Vari_Kys_no = fldvalue.Trim();
                                    break;
                                }
                            case "契約者口座No":
                                {
                                    model_cvitem.Vari_Kys_kozano = fldvalue.Trim();
                                    model_cvitem.Vari_Koza_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "金融期間No":
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
                            case "ゆうちょ記号１":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo1 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ記号２":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo2 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ口座番号":
                                {
                                    model_cvitem.Vari_Yucyokoza_bango = fldvalue.Trim();
                                    break;
                                }
                            case "口座振替No":
                                {
                                    model_cvitem.Vari_Fkae_no = fldvalue.Trim();
                                    break;
                                }
                            case "口座振替手数料":
                                {
                                    model_cvitem.Vari_Fkae_tesugak = fldvalue.Trim();
                                    break;
                                }
                            case "口座振替契約者番号":
                                {
                                    model_cvitem.Vari_Fkae_kysbango = fldvalue.Trim();
                                    break;
                                }
                            case "口座備考":
                                {
                                    model_cvitem.Vari_Biko_koza = fldvalue.Trim();
                                    break;
                                }
                            case "口座振込備考":
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
                            case "備考(総合振込情報)":
                                {
                                    model_cvitem.Vari_Biko_sgfirai = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 契約者口座情報 -add sta
                            case "印刷口座区分":
                                {
                                    model_cvitem.Vari_Koza_printkbn = fldvalue.Trim();
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 契約者口座情報 -add end
                            case "金融機関区分(ゆうちょ銀行フラグ)":
                                {
                                    model_cvitem.Vari_Koza_yucyoflg = fldvalue.Trim();
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

                // 20170116 契約者口座情報ダミーレコード作成処理の追加 -add sta
                if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                {
                    try
                    {
                        Set_KysDummyKoza(sqlcnnv10);
                    }
                    catch (Exception ex)
                    {
                        // ----- ログ出力 -----
                        int tmptmpcnt = 0;
                        string tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(41, sheetname, "データ調整", "データ調整中にエラーが発生しました。"), false);
                        DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref tmptmpcnt);
                    }
                }
                // 20170116 契約者口座情報ダミーレコード作成処理の追加 -add end

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
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT kys_no FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,kys_no) + '-' + CONVERT(varchar,kys_kozano) FROM " + tblname;
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

            /// <summary>
            /// 契約者口座情報へダミーレコードを作成する
            /// </summary>
            /// <param name="sqlcnnv10">賃貸10DB接続用オブジェクト</param>
            /// <remarks>
            /// 20170116 契約者口座情報ダミーレコード作成処理の追加 新規追加
            /// 　詳細はメソッド内へ記載
            /// </remarks>
            public void Set_KysDummyKoza(SqlConnection sqlcnnv10)
            {

                // ****************************************************************************
                // 中間ファイルへの登録作業上、契約者基本情報と契約者口座情報が別になっているため
                // 契約者情報は存在するが契約者口座情報は存在しない状態が発生する
                // 画面上では契約者情報を作成した段階で契約者基本情報(kysdata)と契約者口座情報(kysdata_koza)が
                // 3レコード自動で作成されるため整合性が取れない状態が発生する
                // ****************************************************************************

                string tmp_sql = "";
                int tmpcnt = 0;

                // ①契約者口座No = 1 のデータを作成する
                tmp_sql = tmp_sql + " INSERT INTO kysdata_koza ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 kys_no ";
                tmp_sql = tmp_sql + " 	,1 AS kys_kozano ";
                tmp_sql = tmp_sql + " 	,1 AS koza_kbn ";
                tmp_sql = tmp_sql + " 	,NULL AS kinyu_no ";
                tmp_sql = tmp_sql + " 	,NULL AS kinyu_tenno ";
                tmp_sql = tmp_sql + " 	,1 AS koza_syubetu ";
                tmp_sql = tmp_sql + " 	,NULL AS koza_bango ";
                tmp_sql = tmp_sql + " 	,kys_name AS koza_meigi ";
                tmp_sql = tmp_sql + " 	,kys_kana AS koza_meigikana ";
                tmp_sql = tmp_sql + " 	,'' AS yucyokoza_kigo1 ";
                tmp_sql = tmp_sql + " 	,'' AS yucyokoza_kigo2 ";
                tmp_sql = tmp_sql + " 	,'' AS yucyokoza_bango ";
                tmp_sql = tmp_sql + " 	,NULL AS fkae_no ";
                tmp_sql = tmp_sql + " 	,NULL AS fkae_tesugak ";
                tmp_sql = tmp_sql + " 	,'' AS fkae_kysbango ";
                tmp_sql = tmp_sql + " 	,'' AS biko_koza ";
                tmp_sql = tmp_sql + " 	,'' AS biko_furikomi ";
                tmp_sql = tmp_sql + " 	,1 AS sgfirai_kbn ";
                tmp_sql = tmp_sql + " 	,NULL AS sgfirai_no ";
                tmp_sql = tmp_sql + " 	,2 AS sgfirai_tesufutankbn ";
                tmp_sql = tmp_sql + " 	,1 AS sgfirai_tesukeisankbn ";
                tmp_sql = tmp_sql + " 	,NULL AS sgfirai_tesukotei1gak ";
                tmp_sql = tmp_sql + " 	,NULL AS sgfirai_tesukotei2gak ";
                tmp_sql = tmp_sql + " 	,'' AS biko_sgfirai ";
                tmp_sql = tmp_sql + " 	,1 AS koza_printkbn ";
                tmp_sql = tmp_sql + " 	,0 AS koza_yucyoflg ";
                tmp_sql = tmp_sql + " FROM kysdata WHERE kys_no NOT IN (SELECT kys_no FROM kysdata_koza WHERE kys_kozano = 1) ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 kys_no ";
                tmp_sql = tmp_sql + " 	,1 AS kys_kozano ";
                tmp_sql = tmp_sql + " 	,1 AS koza_kbn ";
                tmp_sql = tmp_sql + " 	,NULL AS kinyu_no ";
                tmp_sql = tmp_sql + " 	,NULL AS kinyu_tenno ";
                tmp_sql = tmp_sql + " 	,1 AS koza_syubetu ";
                tmp_sql = tmp_sql + " 	,NULL AS koza_bango ";
                tmp_sql = tmp_sql + " 	,kys_name AS koza_meigi ";
                tmp_sql = tmp_sql + " 	,kys_kana AS koza_meigikana ";
                tmp_sql = tmp_sql + " 	,'' AS yucyokoza_kigo1 ";
                tmp_sql = tmp_sql + " 	,'' AS yucyokoza_kigo2 ";
                tmp_sql = tmp_sql + " 	,'' AS yucyokoza_bango ";
                tmp_sql = tmp_sql + " 	,NULL AS fkae_no ";
                tmp_sql = tmp_sql + " 	,NULL AS fkae_tesugak ";
                tmp_sql = tmp_sql + " 	,'' AS fkae_kysbango ";
                tmp_sql = tmp_sql + " 	,'' AS biko_koza ";
                tmp_sql = tmp_sql + " 	,'' AS biko_furikomi ";
                tmp_sql = tmp_sql + " 	,1 AS sgfirai_kbn ";
                tmp_sql = tmp_sql + " 	,NULL AS sgfirai_no ";
                tmp_sql = tmp_sql + " 	,2 AS sgfirai_tesufutankbn ";
                tmp_sql = tmp_sql + " 	,1 AS sgfirai_tesukeisankbn ";
                tmp_sql = tmp_sql + " 	,NULL AS sgfirai_tesukotei1gak ";
                tmp_sql = tmp_sql + " 	,NULL AS sgfirai_tesukotei2gak ";
                tmp_sql = tmp_sql + " 	,'' AS biko_sgfirai ";
                tmp_sql = tmp_sql + " 	,1 AS koza_printkbn ";
                tmp_sql = tmp_sql + " 	,0 AS koza_yucyoflg ";
                tmp_sql = tmp_sql + " FROM kysdata WHERE kys_no NOT IN (SELECT kys_no FROM kysdata_koza) ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmpcnt);
                tmp_sql = "";

                // ②契約者口座No = 2、3 のデータを作成する
                // ※過去データ調整で対応したクエリを一部修正して流用する
                tmp_sql = tmp_sql + " BEGIN ";
                tmp_sql = tmp_sql + " 	/*1.変数宣言*/ ";
                tmp_sql = tmp_sql + " 	DECLARE @kysno INT				/*kys_no*/ ";
                tmp_sql = tmp_sql + " 	DECLARE @maxkozano INT			/*[最大口座No]*/ ";
                tmp_sql = tmp_sql + " 	DECLARE @maxkozacnt INT			/*[最大口座数]*/ ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " 	/*2.カーソル宣言*/ ";
                tmp_sql = tmp_sql + " 	DECLARE KysInfo CURSOR FOR ";
                tmp_sql = tmp_sql + " 		/*最大値抽出クエリ文 sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 VW.kys_no ";
                tmp_sql = tmp_sql + " 			,[最大口座No] ";
                tmp_sql = tmp_sql + " 			,[最大口座数] ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 kys_no ";
                tmp_sql = tmp_sql + " 				,MAX(kys_kozano) AS [最大口座No] ";
                tmp_sql = tmp_sql + " 			FROM kysdata_koza ";
                tmp_sql = tmp_sql + " 			GROUP BY kys_no ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		LEFT JOIN ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 kys_no ";
                tmp_sql = tmp_sql + " 					/*,kys_kozano*/ ";
                tmp_sql = tmp_sql + " 					,MAX([口座数抽出用]) AS [最大口座数] ";
                tmp_sql = tmp_sql + " 				FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 kys_no ";
                tmp_sql = tmp_sql + " 						/*,kys_kozano*/ ";
                tmp_sql = tmp_sql + " 						,ROW_NUMBER()OVER(PARTITION BY kys_no ORDER BY kys_kozano) AS [口座数抽出用] ";
                tmp_sql = tmp_sql + " 					FROM kysdata_koza ";
                tmp_sql = tmp_sql + " 				) AS VW1 ";
                tmp_sql = tmp_sql + " 				GROUP BY kys_no ";
                tmp_sql = tmp_sql + " 			) AS VW3 ";
                tmp_sql = tmp_sql + " 		ON VW.kys_no = VW3.kys_no ";
                tmp_sql = tmp_sql + " 		ORDER BY kys_no ";
                tmp_sql = tmp_sql + " 		/*最大値抽出クエリ文 end*/ ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " 	/*3.初期化関連処理*/ ";
                tmp_sql = tmp_sql + " 	SET NOCOUNT ON				/*コメントON*/ ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " 	/*4.メイン処理*/ ";
                tmp_sql = tmp_sql + " 	OPEN KysInfo ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " 		FETCH NEXT FROM KysInfo ";
                tmp_sql = tmp_sql + " 		INTO @kysno,@maxkozano,@maxkozacnt ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " 		/*INSERT処理*/ ";
                tmp_sql = tmp_sql + " 		WHILE @@FETCH_STATUS = 0 ";
                tmp_sql = tmp_sql + " 		BEGIN ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " 			/*最大口座No = 1 かつ 最大口座数 = 1 の場合は口座No = 2、3を作成*/ ";
                tmp_sql = tmp_sql + " 			IF @maxkozano = 1 AND @maxkozacnt = 1 ";
                tmp_sql = tmp_sql + " 				BEGIN ";
                tmp_sql = tmp_sql + " 					INSERT INTO kysdata_koza (kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno,koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn,koza_yucyoflg) VALUES (@kysno,2,2,NULL,NULL,1,NULL,'','','','','',NULL,0.00,'','','',1,NULL,2,1,NULL,NULL,'',1,0) ";
                tmp_sql = tmp_sql + " 					INSERT INTO kysdata_koza (kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno,koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn,koza_yucyoflg) VALUES (@kysno,3,3,NULL,NULL,1,NULL,'','','','','',NULL,0.00,'','','',NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,0) ";
                tmp_sql = tmp_sql + " 			END ";
                tmp_sql = tmp_sql + " 			/*最大口座No = 2 かつ 最大口座数 = 2 の場合は口座No = 3を作成*/ ";
                tmp_sql = tmp_sql + " 			IF @maxkozano = 2 AND @maxkozacnt = 2 ";
                tmp_sql = tmp_sql + " 				BEGIN ";
                tmp_sql = tmp_sql + " 					INSERT INTO kysdata_koza (kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno,koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn,koza_yucyoflg) VALUES (@kysno,3,3,NULL,NULL,1,NULL,'','','','','',NULL,0.00,'','','',NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,0) ";
                tmp_sql = tmp_sql + " 			END ";
                tmp_sql = tmp_sql + " 			/*最大口座No = 3 かつ 最大口座数 = 2 の場合は口座No = 2を作成*/ ";
                tmp_sql = tmp_sql + " 			IF @maxkozano = 3 AND @maxkozacnt = 2 ";
                tmp_sql = tmp_sql + " 				BEGIN ";
                tmp_sql = tmp_sql + " 					INSERT INTO kysdata_koza (kys_no,kys_kozano,koza_kbn,kinyu_no,kinyu_tenno,koza_syubetu,koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2,yucyokoza_bango,fkae_no,fkae_tesugak,fkae_kysbango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai,koza_printkbn,koza_yucyoflg) VALUES (@kysno,2,2,NULL,NULL,1,NULL,'','','','','',NULL,0.00,'','','',1,NULL,2,1,NULL,NULL,'',1,0) ";
                tmp_sql = tmp_sql + " 			END ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " 			FETCH NEXT FROM KysInfo ";
                tmp_sql = tmp_sql + " 			INTO @kysno,@maxkozano,@maxkozacnt ";
                tmp_sql = tmp_sql + " 		END ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " 	CLOSE KysInfo ";
                tmp_sql = tmp_sql + " 	 ";
                tmp_sql = tmp_sql + " 	/*5.解放*/ ";
                tmp_sql = tmp_sql + " 	DEALLOCATE KysInfo ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " END ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmpcnt);
                tmp_sql = "";

            }

        }

    }

    #endregion

    #region 契約者照合用カナ情報

    public class M_fb_fkomsyogo_kys_Repository
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

                var model_cvitem = new Model.M_fb_fkomsyogo_Model();          // 移行値格納用モデル初期化
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
                string tblname_base = "kysdata";
                string tblname = "m_fb_fkomsyogo";
                string fldnamegrp = "sqsaki_kbn,sqsaki_no,sqsaki_kozano,syogo_no,syogo_priority," + "fkom_syogomoji";

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

                    // キー値取得
                    string tmp_keymain = Typ.ToStr(readtbl.Rows[cntii][keycol - 1]).Trim();
                    model_cvitem.Vari_Sqsaki_no = tmp_keymain;

                    // 照合用カナカウント初期化
                    int cvitemcnt = 0;
                    int tmp_cvrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
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
                        model_cvitem.Vari_Fkom_syogomoji = fldvalue;

                        // 固定値
                        model_cvitem.Vari_Sqsaki_kbn = 100.ToString();
                        model_cvitem.Vari_Sqsaki_kozano = 1.ToString();
                        // 20161125 契約者照合用カナ情報の移行処理修正 -chg sta
                        // .Vari_Syogo_no = (cntjj - 1).ToString
                        model_cvitem.Vari_Syogo_no = cntjj.ToString();
                        // 20161125 契約者照合用カナ情報の移行処理修正 -chg end
                        model_cvitem.Vari_Syogo_priority = 1.ToString();

                        // 照合用カナデータが存在する場合に書込処理を行う
                        if (!string.IsNullOrEmpty(model_cvitem.Vari_Fkom_syogomoji))
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
                // 20160928 メモ関連の移行件数表示修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = totalrowcnt;
                // 20160928 メモ関連の移行件数表示修正 -chg end
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

                string tmp_sql = " SELECT kys_no FROM " + tblname;
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

                string tmp_sql = " SELECT sqsaki_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 契約者保証人情報

    public class Kysdata_hosyo_Repository
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

                var model_cvitem = new Model.Kysdata_hosyo_Model();           // 移行値格納用モデル初期化
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
                string tblname_base = "kysdata";
                string tblname = "kysdata_hosyo";
                string fldnamegrp = "kys_no,hosyo_no,hosyo_name,hosyo_namesjis,hosyo_kana," + "keisyo,post_code,addr1,addr2,tel1," + "tel2,fax,mobiletel1,mobiletel2,yusentel_kbn," + "birthday,nensyu,biko_hosyo,aidagara,kinmu_name," + "kinmu_namesjis,kinmu_kana,kinmu_postcode,kinmu_addr1,kinmu_addr2," + "kinmu_tel1,kinmu_tel2,kinmu_fax,kinmu_gyosyu,kinmu_busyo," + "kinmu_nyuryokuym,kinmu_nyusyaym,kinmu_taisyaym,biko_kinmu";






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
                            case "契約者No":
                                {
                                    model_cvitem.Vari_Kys_no = fldvalue.Trim();
                                    break;
                                }
                            case "保証人No":
                                {
                                    model_cvitem.Vari_Hosyo_no = fldvalue.Trim();
                                    break;
                                }
                            case "保証人名":
                                {
                                    model_cvitem.Vari_Hosyo_name = fldvalue.Trim();
                                    break;
                                }
                            case "保証人名SJIS":
                                {
                                    model_cvitem.Vari_Hosyo_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "保証人カナ":
                                {
                                    model_cvitem.Vari_Hosyo_kana = fldvalue.Trim();
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
                            case "住所１":
                                {
                                    model_cvitem.Vari_Addr1 = fldvalue.Trim();
                                    break;
                                }
                            case "住所２":
                                {
                                    model_cvitem.Vari_Addr2 = fldvalue.Trim();
                                    break;
                                }
                            case "TEL１":
                                {
                                    model_cvitem.Vari_Tel1 = fldvalue.Trim();
                                    break;
                                }
                            case "TEL２":
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
                            case "優先設定(電話番号)":
                                {
                                    model_cvitem.Vari_Yusentel_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "誕生日・設立日":
                                {
                                    model_cvitem.Vari_Birthday = fldvalue.Trim();
                                    break;
                                }
                            case "年収・年商":
                                {
                                    // 2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化 -chg sta
                                    // ↓↓↓旧srcコメントアウト↓↓↓
                                    // '2016.03.31 変換時の型落ち対応 -chg sta
                                    // 'Dim tmpnensyustr As String = fldvalue.Trim
                                    // 'Dim tmpnensyudbl As Double = Double.Parse(tmpnensyustr)                 'kakaka4 0322_1300 変換時の型落ちを考慮。
                                    // ''四捨五入して取得
                                    // '.Vari_Nensyu = Math.Round(tmpnensyudbl, MidpointRounding.AwayFromZero).ToString
                                    // Dim tmpnensyustr As String = fldvalue.Trim
                                    // Dim tmpnensyudbl As Double = 0
                                    // If Double.TryParse(tmpnensyustr, tmpnensyudbl) Then
                                    // '四捨五入して取得
                                    // .Vari_Nensyu = Math.Round(tmpnensyudbl, MidpointRounding.AwayFromZero).ToString
                                    // Else
                                    // '変換できない場合は値をそのまま格納 (ログ出力させる)
                                    // .Vari_Nensyu = tmpnensyustr
                                    // End If
                                    // '2016.03.31 変換時の型落ち対応 -chg end
                                    // ↑↑↑旧srcコメントアウト↑↑↑
                                    model_cvitem.Vari_Nensyu = Conversions.ToString(EtcMethod.Get_RoundingValue(fldvalue.Trim()));
                                    break;
                                }
                            // 2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化 -chg end
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko_hosyo = fldvalue.Trim();
                                    break;
                                }
                            case "間柄":
                                {
                                    model_cvitem.Vari_Aidagara = fldvalue.Trim();
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
                            case "勤務先退社年月":
                                {
                                    model_cvitem.Vari_Kinmu_taisyaym = fldvalue.Trim();
                                    break;
                                }
                            case "勤務先備考":
                                {
                                    model_cvitem.Vari_Biko_kinmu = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Keisyo = "様";
                    // 20160530 ログ修正 -del
                    // .Vari_Yusentel_kbn = -1

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

            /// <summary>
            /// 親マスタ取得→リスト格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT kys_no FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,kys_no) + '-' + CONVERT(varchar,hosyo_no) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 契約者メモ情報

    public class Kysdata_memo_Repository
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

                var model_cvitem = new Model.Kysdata_memo_Model();            // 移行値格納用モデル初期化
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
                string tblname_base = "kysdata";
                string tblname = "kysdata_memo";
                string fldnamegrp = "kys_no,memo_no,memo,history";

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
                    model_cvitem.Vari_Kys_no = tmp_keymain;

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
                            string log_key = "kysdata_memo-kys_no";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keymain))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "kysdata_memo-kys_no";
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
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT kys_no FROM " + tblname;
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

                string tmp_sql = " SELECT kys_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

}