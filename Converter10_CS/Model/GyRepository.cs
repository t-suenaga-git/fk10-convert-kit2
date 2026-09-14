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

    #region 仲介業者情報

    public class Gydata_fudo_Repository
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

                var model_cvitem = new Model.Gydata_fudo_Model();             // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname = "gydata_fudo";
                string fldnamegrp = "gy_fudono,gy_fudoname,gy_fudonamesjis,gy_fudokana,post_code," + "addr_kenno,addr_sino,addr_cyo,addr_cyome,addr_cyomeptn," + "addr_banti,addr_etc,tel1,tel2,fax," + "mobiletel1,mobiletel2,daihyo_yakusyoku,daihyo_name,tanto_busyoyakusyoku," + "tanto_name,tanto_namesjis,tanto_kana,mail,mobilemail," + "url,yusentel_kbn,yusenmail_kbn,biko_kihon,sofu_hakkofurikomi," + "menkyo_bango,menkyo_ymd,syunin_bango,syunin_name,kanrikyokai_bango," + "keieikanrisi_bango,keieikanrisi_name,kanrigy_bango,useflg,history," + "rowid,daihyo_namesjis,syunin_namesjis,tanto_autoinputflg,keieikanrisi_namesjis";








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
                    string tmp_address1 = "";
                    string tmp_address2 = "";

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
                            case "業者No":
                                {
                                    model_cvitem.Vari_Gy_fudono = fldvalue.Trim();
                                    break;
                                }
                            case "業者名":
                                {
                                    model_cvitem.Vari_Gy_fudoname = fldvalue.Trim();
                                    break;
                                }
                            case "業者名(SJIS)":
                                {
                                    model_cvitem.Vari_Gy_fudonamesjis = fldvalue.Trim();
                                    break;
                                }
                            case "業者名カナ":
                                {
                                    model_cvitem.Vari_Gy_fudokana = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
                            case "住所1":
                                {
                                    tmp_address1 = fldvalue.Trim();
                                    break;
                                }
                            case "住所2":
                                {
                                    tmp_address2 = fldvalue.Trim();
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
                            case "携帯1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯2":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "代表者役職":
                                {
                                    model_cvitem.Vari_Daihyo_yakusyoku = fldvalue.Trim();
                                    break;
                                }
                            case "代表者名":
                                {
                                    model_cvitem.Vari_Daihyo_name = fldvalue.Trim();
                                    break;
                                }
                            case "担当者部署役職":
                                {
                                    model_cvitem.Vari_Tanto_busyoyakusyoku = fldvalue.Trim();
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
                            case "Webアドレス(URL)":
                                {
                                    model_cvitem.Vari_Url = fldvalue.Trim();
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
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                            case "口座振込通知書発行の許可":
                                {
                                    model_cvitem.Vari_Sofu_hakkofurikomi = fldvalue.Trim();
                                    break;
                                }
                            case "免許証番号":
                                {
                                    model_cvitem.Vari_Menkyo_bango = fldvalue.Trim();
                                    break;
                                }
                            case "免許年月日":
                                {
                                    model_cvitem.Vari_Menkyo_ymd = fldvalue.Trim();
                                    break;
                                }
                            case "取引主任者登録番号":
                                {
                                    model_cvitem.Vari_Syunin_bango = fldvalue.Trim();
                                    break;
                                }
                            case "取引主任者名":
                                {
                                    model_cvitem.Vari_Syunin_name = fldvalue.Trim();
                                    break;
                                }
                            case "全国賃貸不動産管理業協会会員番号":
                                {
                                    model_cvitem.Vari_Kanrikyokai_bango = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸不動産経営管理士登録番号":
                                {
                                    model_cvitem.Vari_Keieikanrisi_bango = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸不動産経営管理士名":
                                {
                                    model_cvitem.Vari_Keieikanrisi_name = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸住宅管理業者登録番号(仮)":
                                {
                                    model_cvitem.Vari_Kanrigy_bango = fldvalue.Trim();
                                    break;
                                }
                            case "代表者名(SJIS)":
                                {
                                    model_cvitem.Vari_Daihyo_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "取引主任者名(SJIS)":
                                {
                                    model_cvitem.Vari_Syunin_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "仲介業者契約担当者自動入力フラグ":
                                {
                                    model_cvitem.Vari_Tanto_autoinputflg = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸不動産経営管理士名(SJIS)":
                                {
                                    model_cvitem.Vari_Keieikanrisi_namesjis = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 住所変換/格納
                    string tmp_address = tmp_address1 + " " + tmp_address2;
                    var hash_address = new SafeDictionary<string, string>();
                    // 20160622 住所分割処理対応 -chg sta
                    // hash_address = AddressChange.Get_Address(tmp_address, sqlcnnv10)
                    hash_address = AddressChange_old.Get_Address(tmp_address, ref sqlcnnv10);
                    // 20160622 住所分割処理対応 -chg end
                    model_cvitem.Vari_Addr_kenno = Conversions.ToString(hash_address["ken_no"]);
                    model_cvitem.Vari_Addr_sino = Conversions.ToString(hash_address["si_no"]);
                    model_cvitem.Vari_Addr_cyo = Conversions.ToString(hash_address["mati"]);
                    model_cvitem.Vari_Addr_cyome = Conversions.ToString(hash_address["cyome"]);
                    model_cvitem.Vari_Addr_cyomeptn = Conversions.ToString(hash_address["chomeptn"]);
                    model_cvitem.Vari_Addr_banti = Conversions.ToString(hash_address["banti"]);
                    model_cvitem.Vari_Addr_etc = Conversions.ToString(hash_address["etc"]);

                    // 固定値
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();     // guid
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

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

            // Dim model_cvitem As New Njc.Model.Gydata_fudo_Model             '移行値格納用モデル初期化
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
            // Dim tblname As String = "gydata_fudo"
            // Dim fldnamegrp As String = "gy_fudono,gy_fudoname,gy_fudonamesjis,gy_fudokana,post_code," & _
            // "addr_kenno,addr_sino,addr_cyo,addr_cyome,addr_cyomeptn," & _
            // "addr_banti,addr_etc,tel1,tel2,fax," & _
            // "mobiletel1,mobiletel2,daihyo_yakusyoku,daihyo_name,tanto_busyoyakusyoku," & _
            // "tanto_name,tanto_namesjis,tanto_kana,mail,mobilemail," & _
            // "url,yusentel_kbn,yusenmail_kbn,biko_kihon,sofu_hakkofurikomi," & _
            // "menkyo_bango,menkyo_ymd,syunin_bango,syunin_name,kanrikyokai_bango," & _
            // "keieikanrisi_bango,keieikanrisi_name,kanrigy_bango,useflg,history," & _
            // "rowid,daihyo_namesjis,syunin_namesjis,tanto_autoinputflg,keieikanrisi_namesjis"

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

            // '作業用変数
            // Dim tmp_address1 As String = ""
            // Dim tmp_address2 As String = ""

            // '移行値取得
            // For cntjj = 1 To columncnt

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "業者No"
            // .Vari_Gy_fudono = fldvalue.Trim
            // Case "業者名"
            // .Vari_Gy_fudoname = fldvalue.Trim
            // Case "業者名(SJIS)"
            // .Vari_Gy_fudonamesjis = fldvalue.Trim
            // Case "業者名カナ"
            // .Vari_Gy_fudokana = fldvalue.Trim
            // Case "郵便番号"
            // .Vari_Post_code = fldvalue.Trim
            // Case "住所1"
            // tmp_address1 = fldvalue.Trim
            // Case "住所2"
            // tmp_address2 = fldvalue.Trim
            // Case "TEL1"
            // .Vari_Tel1 = fldvalue.Trim
            // Case "TEL2"
            // .Vari_Tel2 = fldvalue.Trim
            // Case "FAX"
            // .Vari_Fax = fldvalue.Trim
            // Case "携帯1"
            // .Vari_Mobiletel1 = fldvalue.Trim
            // Case "携帯2"
            // .Vari_Mobiletel2 = fldvalue.Trim
            // Case "代表者役職"
            // .Vari_Daihyo_yakusyoku = fldvalue.Trim
            // Case "代表者名"
            // .Vari_Daihyo_name = fldvalue.Trim
            // Case "担当者部署役職"
            // .Vari_Tanto_busyoyakusyoku = fldvalue.Trim
            // Case "担当者名"
            // .Vari_Tanto_name = fldvalue.Trim
            // Case "担当者名(SJIS)"
            // .Vari_Tanto_namesjis = fldvalue.Trim
            // Case "担当者カナ"
            // .Vari_Tanto_kana = fldvalue.Trim
            // Case "メールアドレス"
            // .Vari_Mail = fldvalue.Trim
            // Case "携帯メールアドレス"
            // .Vari_Mobilemail = fldvalue.Trim
            // Case "Webアドレス(URL)"
            // .Vari_Url = fldvalue.Trim
            // Case "優先設定(電話番号)"
            // .Vari_Yusentel_kbn = fldvalue.Trim
            // Case "優先設定(メールアドレス)"
            // .Vari_Yusenmail_kbn = fldvalue.Trim
            // Case "備考(基本情報)"
            // .Vari_Biko_kihon = fldvalue.Trim
            // Case "口座振込通知書発行の許可"
            // .Vari_Sofu_hakkofurikomi = fldvalue.Trim
            // Case "免許証番号"
            // .Vari_Menkyo_bango = fldvalue.Trim
            // Case "免許年月日"
            // .Vari_Menkyo_ymd = fldvalue.Trim
            // Case "取引主任者登録番号"
            // .Vari_Syunin_bango = fldvalue.Trim
            // Case "取引主任者名"
            // .Vari_Syunin_name = fldvalue.Trim
            // Case "全国賃貸不動産管理業協会会員番号"
            // .Vari_Kanrikyokai_bango = fldvalue.Trim
            // Case "賃貸不動産経営管理士登録番号"
            // .Vari_Keieikanrisi_bango = fldvalue.Trim
            // Case "賃貸不動産経営管理士名"
            // .Vari_Keieikanrisi_name = fldvalue.Trim
            // Case "賃貸住宅管理業者登録番号(仮)"
            // .Vari_Kanrigy_bango = fldvalue.Trim
            // Case "代表者名(SJIS)"
            // .Vari_Daihyo_namesjis = fldvalue.Trim
            // Case "取引主任者名(SJIS)"
            // .Vari_Syunin_namesjis = fldvalue.Trim
            // Case "仲介業者契約担当者自動入力フラグ"
            // .Vari_Tanto_autoinputflg = fldvalue.Trim
            // Case "賃貸不動産経営管理士名(SJIS)"
            // .Vari_Keieikanrisi_namesjis = fldvalue.Trim
            // End Select

            // Next

            // '住所変換/格納
            // Dim tmp_address As String = tmp_address1 & " " & tmp_address2
            // Dim hash_address As New SafeDictionary<string, string>
            // '20160622 住所分割処理対応 -chg sta
            // 'hash_address = AddressChange.Get_Address(tmp_address, sqlcnnv10)
            // hash_address = AddressChange_old.Get_Address(tmp_address, sqlcnnv10)
            // '20160622 住所分割処理対応 -chg end
            // .Vari_Addr_kenno = hash_address("ken_no")
            // .Vari_Addr_sino = hash_address("si_no")
            // .Vari_Addr_cyo = hash_address("mati")
            // .Vari_Addr_cyome = hash_address("cyome")
            // .Vari_Addr_cyomeptn = hash_address("chomeptn")
            // .Vari_Addr_banti = hash_address("banti")
            // .Vari_Addr_etc = hash_address("etc")

            // '固定値
            // .Vari_Rowid = Guid.NewGuid.ToString     'guid
            // .Vari_Useflg = 1
            // .Vari_History = DefHistory

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

            /// <summary>
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_fudono FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 仲介業者口座情報

    public class Gydata_fudokoza_Repository
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

                var model_cvitem = new Model.Gydata_fudokoza_Model();         // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Get_RelData_KozaSyubetu();
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname_base = "gydata_fudo";
                string tblname = "gydata_fudokoza";
                string fldnamegrp = "gy_fudono,gy_fudokozano,kinyu_no,kinyu_tenno,koza_syubetu," + "koza_bango,koza_meigi,koza_meigikana,biko_koza,biko_furikomi," + "sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak," + "sgfirai_tesukotei2gak,biko_sgfirai,history,yucyokoza_kigo1,yucyokoza_kigo2," + "yucyokoza_bango";




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
                            case "業者No":
                                {
                                    model_cvitem.Vari_Gy_fudono = fldvalue.Trim();
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
                            case "備考(口座情報)":
                                {
                                    model_cvitem.Vari_Biko_koza = fldvalue.Trim();
                                    break;
                                }
                            case "備考(口座振込情報)":
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
                            case "ゆうちょ銀行記号1":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo1 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ銀行記号2":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo2 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ銀行番号":
                                {
                                    model_cvitem.Vari_Yucyokoza_bango = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Gy_fudokozano = 1.ToString();
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

            // Dim model_cvitem As New Njc.Model.Gydata_fudokoza_Model         '移行値格納用モデル初期化
            // Dim keycol As Integer = 1                                       'キー列

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
            // Dim tblname_base As String = "gydata_fudo"
            // Dim tblname As String = "gydata_fudokoza"
            // Dim fldnamegrp As String = "gy_fudono,gy_fudokozano,kinyu_no,kinyu_tenno,koza_syubetu," & _
            // "koza_bango,koza_meigi,koza_meigikana,biko_koza,biko_furikomi," & _
            // "sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak," & _
            // "sgfirai_tesukotei2gak,biko_sgfirai,history,yucyokoza_kigo1,yucyokoza_kigo2," & _
            // "yucyokoza_bango"

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

            // '移行値取得
            // For cntjj = 1 To columncnt

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "業者No"
            // .Vari_Gy_fudono = fldvalue.Trim
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
            // Case "備考(口座情報)"
            // .Vari_Biko_koza = fldvalue.Trim
            // Case "備考(口座振込情報)"
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
            // Case "備考(総合振込情報)"
            // .Vari_Biko_sgfirai = fldvalue.Trim
            // Case "ゆうちょ銀行記号1"
            // .Vari_Yucyokoza_kigo1 = fldvalue.Trim
            // Case "ゆうちょ銀行記号2"
            // .Vari_Yucyokoza_kigo2 = fldvalue.Trim
            // Case "ゆうちょ銀行番号"
            // .Vari_Yucyokoza_bango = fldvalue.Trim
            // End Select

            // Next

            // '固定値
            // .Vari_Gy_fudokozano = 1
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
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT gy_fudono FROM " + tblname;
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
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_fudono FROM " + tblname;
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

    #region 仲介業者メモ情報

    public class Gydata_fudomemo_Repository
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

                var model_cvitem = new Model.Gydata_fudomemo_Model();         // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列
                int totalrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname_base = "gydata_fudo";
                string tblname = "gydata_fudomemo";
                string fldnamegrp = "gy_fudono,memo_no,memo,history";

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
                    model_cvitem.Vari_Gy_fudono = tmp_keymain;

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
                            string log_key = "gydata_fudomemo-gy_fudono";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keymain))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "gydata_fudomemo-gy_fudono";
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

                string tmp_sql = " SELECT gy_fudono FROM " + tblname;
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
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_fudono FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 保険業者基本情報

    public class Gydata_hoken_Repository
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

                var model_cvitem = new Model.Gydata_hoken_Model();            // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname = "gydata_hoken";
                string fldnamegrp = "gy_hokenno,gy_hokenname,gy_hokennamesjis,gy_hokenkana,post_code," + "addr1,addr2,tel1,tel2,fax," + "mobiletel1,mobiletel2,tanto_busyoyakusyoku,tanto_name,tanto_namesjis," + "tanto_kana,mail,mobilemail,url,yusentel_kbn," + "yusenmail_kbn,biko_kihon,useflg,history,rowid";




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
                            case "業者No":
                                {
                                    model_cvitem.Vari_Gy_hokenno = fldvalue.Trim();
                                    break;
                                }
                            case "業者名":
                                {
                                    model_cvitem.Vari_Gy_hokenname = fldvalue.Trim();
                                    break;
                                }
                            case "業者名(SJIS)":
                                {
                                    model_cvitem.Vari_Gy_hokennamesjis = fldvalue.Trim();
                                    break;
                                }
                            case "業者名カナ":
                                {
                                    model_cvitem.Vari_Gy_hokenkana = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
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
                            case "携帯1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯2":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "担当者部署役職":
                                {
                                    model_cvitem.Vari_Tanto_busyoyakusyoku = fldvalue.Trim();
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
                            case "担当者名カナ":
                                {
                                    model_cvitem.Vari_Tanto_kana = fldvalue.Trim();
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
                            case "Webアドレス(URL)":
                                {
                                    model_cvitem.Vari_Url = fldvalue.Trim();
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
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

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

            // Dim model_cvitem As New Njc.Model.Gydata_hoken_Model            '移行値格納用モデル初期化
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
            // Dim tblname As String = "gydata_hoken"
            // Dim fldnamegrp As String = "gy_hokenno,gy_hokenname,gy_hokennamesjis,gy_hokenkana,post_code," & _
            // "addr1,addr2,tel1,tel2,fax," & _
            // "mobiletel1,mobiletel2,tanto_busyoyakusyoku,tanto_name,tanto_namesjis," & _
            // "tanto_kana,mail,mobilemail,url,yusentel_kbn," & _
            // "yusenmail_kbn,biko_kihon,useflg,history,rowid"

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

            // '移行値取得
            // For cntjj = 1 To columncnt

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "業者No"
            // .Vari_Gy_hokenno = fldvalue.Trim
            // Case "業者名"
            // .Vari_Gy_hokenname = fldvalue.Trim
            // Case "業者名(SJIS)"
            // .Vari_Gy_hokennamesjis = fldvalue.Trim
            // Case "業者名カナ"
            // .Vari_Gy_hokenkana = fldvalue.Trim
            // Case "郵便番号"
            // .Vari_Post_code = fldvalue.Trim
            // Case "住所1"
            // .Vari_Addr1 = fldvalue.Trim
            // Case "住所2"
            // .Vari_Addr2 = fldvalue.Trim
            // Case "TEL1"
            // .Vari_Tel1 = fldvalue.Trim
            // Case "TEL2"
            // .Vari_Tel2 = fldvalue.Trim
            // Case "FAX"
            // .Vari_Fax = fldvalue.Trim
            // Case "携帯1"
            // .Vari_Mobiletel1 = fldvalue.Trim
            // Case "携帯2"
            // .Vari_Mobiletel2 = fldvalue.Trim
            // Case "担当者部署役職"
            // .Vari_Tanto_busyoyakusyoku = fldvalue.Trim
            // Case "担当者名"
            // .Vari_Tanto_name = fldvalue.Trim
            // Case "担当者名(SJIS)"
            // .Vari_Tanto_namesjis = fldvalue.Trim
            // Case "担当者名カナ"
            // .Vari_Tanto_kana = fldvalue.Trim
            // Case "メールアドレス"
            // .Vari_Mail = fldvalue.Trim
            // Case "携帯メールアドレス"
            // .Vari_Mobilemail = fldvalue.Trim
            // Case "Webアドレス(URL)"
            // .Vari_Url = fldvalue.Trim
            // Case "優先設定(電話番号)"
            // .Vari_Yusentel_kbn = fldvalue.Trim
            // Case "優先設定(メールアドレス)"
            // .Vari_Yusenmail_kbn = fldvalue.Trim
            // Case "備考(基本情報)"
            // .Vari_Biko_kihon = fldvalue.Trim
            // End Select

            // Next

            // '固定値
            // .Vari_Rowid = Guid.NewGuid.ToString
            // .Vari_Useflg = 1
            // .Vari_History = DefHistory

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

            /// <summary>
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_hokenno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 保険業者口座情報

    public class Gydata_hokenkoza_Repository
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

                var model_cvitem = new Model.Gydata_hokenkoza_Model();        // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Get_RelData_KozaSyubetu();
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname_base = "gydata_hoken";
                string tblname = "gydata_hokenkoza";
                string fldnamegrp = "gy_hokenno,gy_hokenkozano,kinyu_no,kinyu_tenno,koza_syubetu," + "koza_bango,koza_meigi,koza_meigikana,biko_koza,biko_furikomi," + "sgfirai_kbn,sgfirai_no,sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak," + "sgfirai_tesukotei2gak,biko_sgfirai,history";



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
                            case "業者No":
                                {
                                    model_cvitem.Vari_Gy_hokenno = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関No":
                                {
                                    model_cvitem.Vari_Kinyu_no = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関店No":
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
                            case "備考(口座情報)":
                                {
                                    model_cvitem.Vari_Biko_koza = fldvalue.Trim();
                                    break;
                                }
                            case "備考(口座振込情報)":
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
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Gy_hokenkozano = 1.ToString();
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

                string tmp_sql = " SELECT gy_hokenno FROM " + tblname;
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
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_hokenno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 保険業者メモ情報

    public class Gydata_hokenmemo_Repository
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

                var model_cvitem = new Model.Gydata_hokenmemo_Model();        // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列
                int totalrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname_base = "gydata_hoken";
                string tblname = "gydata_hokenmemo";
                string fldnamegrp = "gy_hokenno,memo_no,memo,history";

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
                    model_cvitem.Vari_Gy_hokenno = tmp_keymain;

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
                            string log_key = "gydata_hokenmemo-gy_hokenno";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keymain))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "gydata_hokenmemo-gy_hokenno";
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

                string tmp_sql = " SELECT gy_hokenno FROM " + tblname;
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
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_hokenno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 家賃保証業者情報

    public class Gydata_hosyo_Repository
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

                var model_cvitem = new Model.Gydata_hosyo_Model();            // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname = "gydata_hosyo";
                string fldnamegrp = "gy_hosyono,gy_hosyoname,gy_hosyonamesjis,gy_hosyokana,post_code," + "addr1,addr2,tel1,tel2,fax," + "mobiletel1,mobiletel2,tanto_busyoyakusyoku,tanto_name,tanto_namesjis," + "tanto_kana,mail,mobilemail,url,yusentel_kbn," + "yusenmail_kbn,biko_kihon,useflg,history,rowid," + "rendo_namekbn";





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
                            case "業者No":
                                {
                                    model_cvitem.Vari_Gy_hosyono = fldvalue.Trim();
                                    break;
                                }
                            case "業者名":
                                {
                                    model_cvitem.Vari_Gy_hosyoname = fldvalue.Trim();
                                    break;
                                }
                            case "業者名(shift-jis)":
                                {
                                    model_cvitem.Vari_Gy_hosyonamesjis = fldvalue.Trim();
                                    break;
                                }
                            case "業者名カナ":
                                {
                                    model_cvitem.Vari_Gy_hosyokana = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
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
                            case "携帯1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯2":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "担当者部署役職":
                                {
                                    model_cvitem.Vari_Tanto_busyoyakusyoku = fldvalue.Trim();
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
                            case "担当者名カナ":
                                {
                                    model_cvitem.Vari_Tanto_kana = fldvalue.Trim();
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
                            case "Webアドレス(URL)":
                                {
                                    model_cvitem.Vari_Url = fldvalue.Trim();
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
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                            case "連動用家賃保証会社区分":
                                {
                                    model_cvitem.Vari_Rendo_namekbn = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

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

            /// <summary>
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_hosyono FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 家賃保証会社口座情報

    public class gydata_hosyokoza_Repository
    {

        // 革命10に過去にあったが現在(2015/10/23)は無くなっている

    }

    #endregion

    #region 家賃保証業者メモ情報

    public class gydata_hosyomemo_Repository
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

                var model_cvitem = new Model.Gydata_hosyomemo_Model();        // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列
                int totalrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname_base = "gydata_hosyo";
                string tblname = "gydata_hosyomemo";
                string fldnamegrp = "gy_hosyono,memo_no,memo,history";

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
                    model_cvitem.Vari_Gy_hosyono = tmp_keymain;

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
                            string log_key = "gydata_hosyomemo-gy_hosyono";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keymain))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "gydata_hosyomemo-gy_hosyono";
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

                string tmp_sql = " SELECT gy_hosyono FROM " + tblname;
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
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_hosyono FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 修繕業者情報

    public class Gydata_szen_Repository
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

                var model_cvitem = new Model.Gydata_szen_Model();             // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname = "gydata_szen";
                string fldnamegrp = "gy_szenno,gy_szenname,gy_szennamesjis,gy_szenkana,post_code," + "addr1,addr2,tel1,tel2,fax," + "mobiletel1,mobiletel2,daihyo_yakusyoku,daihyo_name,tanto_busyoyakusyoku," + "tanto_name,tanto_namesjis,tanto_kana,mail,mobilemail," + "url,yusentel_kbn,yusenmail_kbn,biko_kihon,useflg," + "history,rowid,daihyo_namesjis";





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
                            case "業者No":
                                {
                                    model_cvitem.Vari_Gy_szenno = fldvalue.Trim();
                                    break;
                                }
                            case "業者名":
                                {
                                    model_cvitem.Vari_Gy_szenname = fldvalue.Trim();
                                    break;
                                }
                            case "業者名(SJIS)":
                                {
                                    model_cvitem.Vari_Gy_szennamesjis = fldvalue.Trim();
                                    break;
                                }
                            case "業者名カナ":
                                {
                                    model_cvitem.Vari_Gy_szenkana = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
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
                            case "携帯1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯2":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "代表者役職":
                                {
                                    model_cvitem.Vari_Daihyo_yakusyoku = fldvalue.Trim();
                                    break;
                                }
                            case "代表者名":
                                {
                                    model_cvitem.Vari_Daihyo_name = fldvalue.Trim();
                                    break;
                                }
                            case "担当者部署役職":
                                {
                                    model_cvitem.Vari_Tanto_busyoyakusyoku = fldvalue.Trim();
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
                            case "担当者名カナ":
                                {
                                    model_cvitem.Vari_Tanto_kana = fldvalue.Trim();
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
                            case "Webアドレス(URL)":
                                {
                                    model_cvitem.Vari_Url = fldvalue.Trim();
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
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                            case "代表者名(SJIS)":
                                {
                                    model_cvitem.Vari_Daihyo_namesjis = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();     // guid
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

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

            /// <summary>
            /// 既存データのキーを取得してリストへ格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <param name="list_existdata"></param>
            /// <remarks></remarks>
            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

                string tmp_sql = " SELECT gy_szenno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 修繕業者口座情報

    public class Gydata_szenkoza_Repository
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

                var model_cvitem = new Model.Gydata_szenkoza_Model();         // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                SetRelItemToObject.Get_RelData_KozaSyubetu();

                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname_base = "gydata_szen";
                string tblname = "gydata_szenkoza";
                string fldnamegrp = "gy_szenno,gy_szenkozano,kinyu_no,kinyu_tenno,koza_syubetu," + "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," + "yucyokoza_bango,biko_koza,biko_furikomi,sgfirai_kbn,sgfirai_no," + "sgfirai_tesufutankbn,sgfirai_tesukeisankbn,sgfirai_tesukotei1gak,sgfirai_tesukotei2gak,biko_sgfirai," + "history";




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
                            case "業者No":
                                {
                                    model_cvitem.Vari_Gy_szenno = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関No":
                                {
                                    model_cvitem.Vari_Kinyu_no = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関店No":
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
                            case "ゆうちょ銀行記号1":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo1 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ銀行記号2":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo2 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ銀行番号":
                                {
                                    model_cvitem.Vari_Yucyokoza_bango = fldvalue.Trim();
                                    break;
                                }
                            case "備考(口座情報)":
                                {
                                    model_cvitem.Vari_Biko_koza = fldvalue.Trim();
                                    break;
                                }
                            case "備考(口座振込情報)":
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
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Gy_szenkozano = 1.ToString();
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

                string tmp_sql = " SELECT gy_szenno FROM " + tblname;
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

                string tmp_sql = " SELECT gy_szenno FROM " + tblname;
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

    #region 修繕業者メモ情報

    public class Gydata_szenmemo_Repository
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

                var model_cvitem = new Model.Gydata_szenmemo_Model();         // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列
                int totalrowcnt = 0;                                 // 20160928 メモ関連の移行件数表示修正 -add

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname_base = "gydata_szen";
                string tblname = "gydata_szenmemo";
                string fldnamegrp = "gy_szenno,memo_no,memo,history";

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
                    model_cvitem.Vari_Gy_szenno = tmp_keymain;

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
                            string log_key = "gydata_szenmemo-gy_szenno";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keymain))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "gydata_szenmemo-gy_szenno";
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

                string tmp_sql = " SELECT gy_szenno FROM " + tblname;
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

                string tmp_sql = " SELECT gy_szenno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region ライフライン業者情報

    public class Gydata_lifeline_Repository
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

                var model_cvitem = new Model.Gydata_lifeline_Model();         // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname = "gydata_lifeline";
                string fldnamegrp = "gy_lifelineno,gy_lifelinename,gy_lifelinenamesjis,gy_lifelinekana,eigyosyo," + "eigyosyosjis,post_code,addr1,addr2,tel1," + "tel2,fax,mobiletel1,mobiletel2,yusentel_kbn," + "biko_kihon,useflg,history,rowid,lifeline_denkiflg," + "lifeline_josuidoflg,lifeline_gasflg,lifeline_toyuflg,lifeline_other1flg,lifeline_haisuiflg";




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
                            case "ライフライン業者No":
                                {
                                    model_cvitem.Vari_Gy_lifelineno = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン業者名":
                                {
                                    model_cvitem.Vari_Gy_lifelinename = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン業者名(Shift-jis)":
                                {
                                    model_cvitem.Vari_Gy_lifelinenamesjis = fldvalue.Trim();
                                    break;
                                }
                            case "ライフライン業者名カナ":
                                {
                                    model_cvitem.Vari_Gy_lifelinekana = fldvalue.Trim();
                                    break;
                                }
                            case "営業所":
                                {
                                    model_cvitem.Vari_Eigyosyo = fldvalue.Trim();
                                    break;
                                }
                            case "営業所(Shift-jis)":
                                {
                                    model_cvitem.Vari_Eigyosyosjis = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
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
                            case "携帯1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯2":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "優先設定(電話番号)":
                                {
                                    model_cvitem.Vari_Yusentel_kbn = fldvalue.Trim();
                                    break;
                                }
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                            case "電気":
                                {
                                    model_cvitem.Vari_Lifeline_denkiflg = fldvalue.Trim();
                                    break;
                                }
                            case "上水道":
                                {
                                    model_cvitem.Vari_Lifeline_josuidoflg = fldvalue.Trim();
                                    break;
                                }
                            case "ガス":
                                {
                                    model_cvitem.Vari_Lifeline_gasflg = fldvalue.Trim();
                                    break;
                                }
                            case "灯油":
                                {
                                    model_cvitem.Vari_Lifeline_toyuflg = fldvalue.Trim();
                                    break;
                                }
                            case "その他":
                                {
                                    model_cvitem.Vari_Lifeline_other1flg = fldvalue.Trim();
                                    break;
                                }
                            case "排水":
                                {
                                    model_cvitem.Vari_Lifeline_haisuiflg = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Useflg = 1.ToString();
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

                string tmp_sql = " SELECT gy_lifelineno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 施工会社情報

    public class Gydata_seko_Repository
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

                var model_cvitem = new Model.Gydata_seko_Model();             // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname = "gydata_seko";
                string fldnamegrp = "gy_sekono,gy_sekoname,gy_sekonamesjis,gy_sekokana,eigyosyo," + "eigyosyosjis,post_code,addr1,addr2,tel1," + "tel2,fax,mobiletel1,mobiletel2,tanto_busyoyakusyoku," + "tanto_name,tanto_namesjis,tanto_kana,mail,mobilemail," + "url,yusentel_kbn,yusenmail_kbn,biko_kihon,useflg," + "history,rowid";





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
                            case "施工業者No":
                                {
                                    model_cvitem.Vari_Gy_sekono = fldvalue.Trim();
                                    break;
                                }
                            case "施工業者名":
                                {
                                    model_cvitem.Vari_Gy_sekoname = fldvalue.Trim();
                                    break;
                                }
                            case "施工業者名(Shift-jis)":
                                {
                                    model_cvitem.Vari_Gy_sekonamesjis = fldvalue.Trim();
                                    break;
                                }
                            case "施工業者名カナ":
                                {
                                    model_cvitem.Vari_Gy_sekokana = fldvalue.Trim();
                                    break;
                                }
                            case "営業所":
                                {
                                    model_cvitem.Vari_Eigyosyo = fldvalue.Trim();
                                    break;
                                }
                            case "営業所(Shift-jis)":
                                {
                                    model_cvitem.Vari_Eigyosyosjis = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
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
                            case "携帯1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯2":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "担当者部署役職":
                                {
                                    model_cvitem.Vari_Tanto_busyoyakusyoku = fldvalue.Trim();
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
                            case "Webアドレス(URL)":
                                {
                                    model_cvitem.Vari_Url = fldvalue.Trim();
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
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();

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

                string tmp_sql = " SELECT gy_sekono FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 施設保守業者情報

    public class Gydata_sisetu_Repository
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

                var model_cvitem = new Model.Gydata_sisetu_Model();           // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
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
                // '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
                // テーブル名/フィールド名セット
                string tblname = "gydata_sisetu";
                string fldnamegrp = "gy_sisetuno,gy_sisetuname,gy_sisetunamesjis,gy_sisetukana,eigyosyo," + "eigyosyosjis,post_code,addr1,addr2,tel1," + "tel2,fax,mobiletel1,mobiletel2,tanto_busyoyakusyoku," + "tanto_name,tanto_namesjis,tanto_kana,mail,mobilemail," + "url,yusentel_kbn,yusenmail_kbn,biko_kihon,useflg," + "history,rowid";





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
                            case "施設保守業者No":
                                {
                                    model_cvitem.Vari_Gy_sisetuno = fldvalue.Trim();
                                    break;
                                }
                            case "施設保守業者名":
                                {
                                    model_cvitem.Vari_Gy_sisetuname = fldvalue.Trim();
                                    break;
                                }
                            case "施設保守業者名(Shift-jis)":
                                {
                                    model_cvitem.Vari_Gy_sisetunamesjis = fldvalue.Trim();
                                    break;
                                }
                            case "施設保守業者名カナ":
                                {
                                    model_cvitem.Vari_Gy_sisetukana = fldvalue.Trim();
                                    break;
                                }
                            case "営業所":
                                {
                                    model_cvitem.Vari_Eigyosyo = fldvalue.Trim();
                                    break;
                                }
                            case "営業所(Shift-jis)":
                                {
                                    model_cvitem.Vari_Eigyosyosjis = fldvalue.Trim();
                                    break;
                                }
                            case "郵便番号":
                                {
                                    model_cvitem.Vari_Post_code = fldvalue.Trim();
                                    break;
                                }
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
                            case "携帯1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯2":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "担当者部署役職":
                                {
                                    model_cvitem.Vari_Tanto_busyoyakusyoku = fldvalue.Trim();
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
                            case "Webアドレス(URL)":
                                {
                                    model_cvitem.Vari_Url = fldvalue.Trim();
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
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Rowid = Guid.NewGuid().ToString();     // guid

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

                string tmp_sql = " SELECT gy_sisetuno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

}