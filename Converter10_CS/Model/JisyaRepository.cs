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

    #region 自社情報

    public class Jisyadata_Repository
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

                var model_cvitem = new Model.Jisyadata_Model();               // 移行値格納用モデル初期化
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
                string tblname = "jisyadata";
                string fldnamegrp = "jisya_no,jisya_name,jisya_namesjis,jisya_kana,jisya_name2," + "jisya_name2sjis,jisya_name3,jisya_name3sjis,daihyo_yakusyoku,daihyo_name," + "post_code,addr_kenno,addr_sino,addr_cyo,addr_cyome," + "addr_cyomeptn,addr_banti,addr_etc,tel1,tel2," + "mobiletel1,mobiletel2,fax,mail,mobilemail," + "url,biko_kihon,menkyo_bango,menkyo_ymd,syunin_bango," + "syunin_name,kanrikyokai_bango,keieikanrisi_bango,keieikanrisi_name,kanrigy_bango," + "kanyu_dantai1,kanyu_dantai2,kanyu_dantai3,kanyu_dantai4,kanyu_dantai5," + "kanyu_dantai6,kanyu_dantai7,kanyu_dantai8,kanyu_dantai9,kanyu_dantai10," + "history,useflg,daihyo_namesjis,syunin_namesjis,keieikanrisi_namesjis";









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
                            case "自社・支店No":
                                {
                                    model_cvitem.Vari_Jisya_no = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店名":
                                {
                                    model_cvitem.Vari_Jisya_name = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店名（SJIS）":
                                {
                                    model_cvitem.Vari_Jisya_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店カナ":
                                {
                                    model_cvitem.Vari_Jisya_kana = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店名２":
                                {
                                    model_cvitem.Vari_Jisya_name2 = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店名２（SJIS）":
                                {
                                    model_cvitem.Vari_Jisya_name2sjis = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店名３":
                                {
                                    model_cvitem.Vari_Jisya_name3 = fldvalue.Trim();
                                    break;
                                }
                            case "自社・支店名３（SJIS）":
                                {
                                    model_cvitem.Vari_Jisya_name3sjis = fldvalue.Trim();
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
                            case "携帯TEL1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue.Trim();
                                    break;
                                }
                            case "携帯TEL2":
                                {
                                    model_cvitem.Vari_Mobiletel2 = fldvalue.Trim();
                                    break;
                                }
                            case "FAX":
                                {
                                    model_cvitem.Vari_Fax = fldvalue.Trim();
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
                            case "Webアドレス":
                                {
                                    model_cvitem.Vari_Url = fldvalue.Trim();
                                    break;
                                }
                            case "備考（基本情報）":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
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
                            case "取引主任者登録番号（代表）":
                                {
                                    model_cvitem.Vari_Syunin_bango = fldvalue.Trim();
                                    break;
                                }
                            case "取引主任者名（代表）":
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
                            case "賃貸住宅管理業者登録番号":
                                {
                                    model_cvitem.Vari_Kanrigy_bango = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体１":
                                {
                                    model_cvitem.Vari_Kanyu_dantai1 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体２":
                                {
                                    model_cvitem.Vari_Kanyu_dantai2 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体３":
                                {
                                    model_cvitem.Vari_Kanyu_dantai3 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体４":
                                {
                                    model_cvitem.Vari_Kanyu_dantai4 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体５":
                                {
                                    model_cvitem.Vari_Kanyu_dantai5 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体６":
                                {
                                    model_cvitem.Vari_Kanyu_dantai6 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体７":
                                {
                                    model_cvitem.Vari_Kanyu_dantai7 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体８":
                                {
                                    model_cvitem.Vari_Kanyu_dantai8 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体９":
                                {
                                    model_cvitem.Vari_Kanyu_dantai9 = fldvalue.Trim();
                                    break;
                                }
                            case "加入団体１０":
                                {
                                    model_cvitem.Vari_Kanyu_dantai10 = fldvalue.Trim();
                                    break;
                                }
                            case "代表者名（SJIS）":
                                {
                                    model_cvitem.Vari_Daihyo_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "取引主任者名（代表）（SJIS）":
                                {
                                    model_cvitem.Vari_Syunin_namesjis = fldvalue.Trim();
                                    break;
                                }
                            case "賃貸不動産経営管理士名（SJIS）":
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

            // Dim model_cvitem As New Njc.Model.Jisyadata_Model               '移行値格納用モデル初期化
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
            // Dim tblname As String = "jisyadata"
            // Dim fldnamegrp As String = "jisya_no,jisya_name,jisya_namesjis,jisya_kana,jisya_name2," & _
            // "jisya_name2sjis,jisya_name3,jisya_name3sjis,daihyo_yakusyoku,daihyo_name," & _
            // "post_code,addr_kenno,addr_sino,addr_cyo,addr_cyome," & _
            // "addr_cyomeptn,addr_banti,addr_etc,tel1,tel2," & _
            // "mobiletel1,mobiletel2,fax,mail,mobilemail," & _
            // "url,biko_kihon,menkyo_bango,menkyo_ymd,syunin_bango," & _
            // "syunin_name,kanrikyokai_bango,keieikanrisi_bango,keieikanrisi_name,kanrigy_bango," & _
            // "kanyu_dantai1,kanyu_dantai2,kanyu_dantai3,kanyu_dantai4,kanyu_dantai5," & _
            // "kanyu_dantai6,kanyu_dantai7,kanyu_dantai8,kanyu_dantai9,kanyu_dantai10," & _
            // "history,useflg,daihyo_namesjis,syunin_namesjis,keieikanrisi_namesjis"

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
            // Case "自社・支店No"
            // .Vari_Jisya_no = fldvalue.Trim
            // Case "自社・支店名"
            // .Vari_Jisya_name = fldvalue.Trim
            // Case "自社・支店名（SJIS）"
            // .Vari_Jisya_namesjis = fldvalue.Trim
            // Case "自社・支店カナ"
            // .Vari_Jisya_kana = fldvalue.Trim
            // Case "自社・支店名２"
            // .Vari_Jisya_name2 = fldvalue.Trim
            // Case "自社・支店名２（SJIS）"
            // .Vari_Jisya_name2sjis = fldvalue.Trim
            // Case "自社・支店名３"
            // .Vari_Jisya_name3 = fldvalue.Trim
            // Case "自社・支店名３（SJIS）"
            // .Vari_Jisya_name3sjis = fldvalue.Trim
            // Case "代表者役職"
            // .Vari_Daihyo_yakusyoku = fldvalue.Trim
            // Case "代表者名"
            // .Vari_Daihyo_name = fldvalue.Trim
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
            // Case "携帯TEL1"
            // .Vari_Mobiletel1 = fldvalue.Trim
            // Case "携帯TEL2"
            // .Vari_Mobiletel2 = fldvalue.Trim
            // Case "FAX"
            // .Vari_Fax = fldvalue.Trim
            // Case "メールアドレス"
            // .Vari_Mail = fldvalue.Trim
            // Case "携帯メールアドレス"
            // .Vari_Mobilemail = fldvalue.Trim
            // Case "Webアドレス"
            // .Vari_Url = fldvalue.Trim
            // Case "備考（基本情報）"
            // .Vari_Biko_kihon = fldvalue.Trim
            // Case "免許証番号"
            // .Vari_Menkyo_bango = fldvalue.Trim
            // Case "免許年月日"
            // .Vari_Menkyo_ymd = fldvalue.Trim
            // Case "取引主任者登録番号（代表）"
            // .Vari_Syunin_bango = fldvalue.Trim
            // Case "取引主任者名（代表）"
            // .Vari_Syunin_name = fldvalue.Trim
            // Case "全国賃貸不動産管理業協会会員番号"
            // .Vari_Kanrikyokai_bango = fldvalue.Trim
            // Case "賃貸不動産経営管理士登録番号"
            // .Vari_Keieikanrisi_bango = fldvalue.Trim
            // Case "賃貸不動産経営管理士名"
            // .Vari_Keieikanrisi_name = fldvalue.Trim
            // Case "賃貸住宅管理業者登録番号"
            // .Vari_Kanrigy_bango = fldvalue.Trim
            // Case "加入団体１"
            // .Vari_Kanyu_dantai1 = fldvalue.Trim
            // Case "加入団体２"
            // .Vari_Kanyu_dantai2 = fldvalue.Trim
            // Case "加入団体３"
            // .Vari_Kanyu_dantai3 = fldvalue.Trim
            // Case "加入団体４"
            // .Vari_Kanyu_dantai4 = fldvalue.Trim
            // Case "加入団体５"
            // .Vari_Kanyu_dantai5 = fldvalue.Trim
            // Case "加入団体６"
            // .Vari_Kanyu_dantai6 = fldvalue.Trim
            // Case "加入団体７"
            // .Vari_Kanyu_dantai7 = fldvalue.Trim
            // Case "加入団体８"
            // .Vari_Kanyu_dantai8 = fldvalue.Trim
            // Case "加入団体９"
            // .Vari_Kanyu_dantai9 = fldvalue.Trim
            // Case "加入団体１０"
            // .Vari_Kanyu_dantai10 = fldvalue.Trim
            // Case "代表者名（SJIS）"
            // .Vari_Daihyo_namesjis = fldvalue.Trim
            // Case "取引主任者名（代表）（SJIS）"
            // .Vari_Syunin_namesjis = fldvalue.Trim
            // Case "賃貸不動産経営管理士名（SJIS）"
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


                string tmp_sql = " SELECT jisya_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);


            }

        }

    }

    #endregion

    #region 自社口座情報

    // V7には自社口座情報が存在しない
    // 家賃入金口座マスタに登録されている口座情報から判別
    // 判別は紐付けツールで行う
    // 紐付け材料は名称、口座名義、カナ等

    public class Jisyadata_koza_Repository
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

                var model_cvitem = new Model.Jisyadata_koza_Model();          // 移行値格納用モデル初期化
                int keycol_main = 15;                                  // メインキー列
                int keycol_sub = 17;                                   // サブキー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                Get_RelData_KozaSyubetu();

                // Excelファイル初期設定
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // 自社口座は紐付けファイルから取得する
                // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)
                string readfiledirpath = CommonModule.RelationDirPath;         // 紐付けファイル格納先
                string readfilename = CommonModule.MIDFILE_RELNAME;            // 紐付けファイル名
                string readsheetname = "自社口座マスタ";          // 紐付けシート名
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, readfiledirpath, readfilename, readsheetname);
                // 2016.04.26 メインの方へも反映させる修正 -chg end

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // テーブル名/フィールド名セット
                string tblname_base = "jisyadata";
                string tblname = "jisyadata_koza";
                string fldnamegrp = "jisya_no,jisya_kozano,kinyu_no,kinyu_tenno,koza_syubetu," + "koza_bango,koza_meigi,koza_meigikana,yucyokoza_kigo1,yucyokoza_kigo2," + "yucyokoza_bango,biko_koza,history,useflg";


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
                            case "自社No":
                                {
                                    model_cvitem.Vari_Jisya_no = fldvalue.Trim();
                                    break;
                                }
                            case "自社口座No":
                                {
                                    model_cvitem.Vari_Jisya_kozano = fldvalue.Trim();
                                    break;
                                }
                            case "金融機関No":
                                {
                                    model_cvitem.Vari_Kinyu_no = fldvalue.Trim();
                                    break;
                                }
                            case "支店No":
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
                            // 20160516 FB構築に伴う自社口座情報の修正 -add sta
                            case "口座名義カナ":
                                {
                                    model_cvitem.Vari_Koza_meigikana = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ記号1":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo1 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ記号2":
                                {
                                    model_cvitem.Vari_Yucyokoza_kigo2 = fldvalue.Trim();
                                    break;
                                }
                            case "ゆうちょ口座番号":
                                {
                                    model_cvitem.Vari_Yucyokoza_bango = fldvalue.Trim();
                                    break;
                                }
                            case "備考(口座情報)":
                                {
                                    model_cvitem.Vari_Biko_koza = fldvalue.Trim();
                                    break;
                                }
                                // 20160516 FB構築に伴う自社口座情報の修正 -add end
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Useflg = 1.ToString();

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

                string tmp_sql = " SELECT jisya_no FROM " + tblname;
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

    #region 自社担当者情報

    public class Profile_logonuser_Repository
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

                var model_cvitem = new Model.Profile_logonuser_Model();       // 移行値格納用モデル初期化
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
                string tblname = "profile_logonuser";
                string fldnamegrp = "logonuser_no,logonuser_name,logonuser_namesjis,logonuser_kana,logongroup_no," + "tel1,tel2,mobiletel1,fax1,mailaddress1," + "mobilemailaddress1,takkenmenkyo,biko,denylogon,noassign," + "password,behaviourdefine_base64,colordefine_base64,n3spreadcolordefine_base64,menudesigndefine_base64," + "smtpdefine_base64,history,sendtargetsiteid,windowsuser_name,hysearchdefine_base64," + "csvsetting_base64,dontshowfeedbackagain";





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
                            case "賃貸革命ログオンユーザNo":
                                {
                                    model_cvitem.Vari_Logonuser_no = fldvalue;
                                    break;
                                }
                            case "賃貸革命ログオンユーザ名":
                                {
                                    model_cvitem.Vari_Logonuser_name = fldvalue;
                                    break;
                                }
                            case "賃貸革命ログオンユーザ名(SJIS)":
                                {
                                    model_cvitem.Vari_Logonuser_namesjis = fldvalue;
                                    break;
                                }
                            case "賃貸革命ログオンユーザ名カナ":
                                {
                                    model_cvitem.Vari_Logonuser_kana = fldvalue;
                                    break;
                                }
                            case "賃貸革命ログオングループNo":
                                {
                                    // 20160928 自社担当者移行処理の修正 -chg sta
                                    // .Vari_Logongroup_no = fldvalue
                                    model_cvitem.Vari_Logongroup_no = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(fldvalue), "99001", fldvalue));
                                    break;
                                }
                            // 20160928 自社担当者移行処理の修正 -chg end
                            case "TEL1":
                                {
                                    model_cvitem.Vari_Tel1 = fldvalue;
                                    break;
                                }
                            case "TEL2":
                                {
                                    model_cvitem.Vari_Tel2 = fldvalue;
                                    break;
                                }
                            case "携帯電話1":
                                {
                                    model_cvitem.Vari_Mobiletel1 = fldvalue;
                                    break;
                                }
                            case "FAX1":
                                {
                                    model_cvitem.Vari_Fax1 = fldvalue;
                                    break;
                                }
                            case "メールアドレス1":
                                {
                                    model_cvitem.Vari_Mailaddress1 = fldvalue;
                                    break;
                                }
                            case "携帯メールアドレス1":
                                {
                                    model_cvitem.Vari_Mobilemailaddress1 = fldvalue;
                                    break;
                                }
                            case "宅建免許番号":
                                {
                                    model_cvitem.Vari_Takkenmenkyo = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            case "担当者ログオン禁止フラグ":
                                {
                                    model_cvitem.Vari_Denylogon = fldvalue;
                                    break;
                                }
                            case "使用しないフラグ":
                                {
                                    model_cvitem.Vari_Noassign = fldvalue;
                                    break;
                                }
                            case "ログオンパスワード":
                                {
                                    model_cvitem.Vari_Password = fldvalue;
                                    break;
                                }
                            case "windowsログオンユーザー名":
                                {
                                    model_cvitem.Vari_Windowsuser_name = fldvalue;
                                    break;
                                }
                            case "意見画面表示フラグ":
                                {
                                    model_cvitem.Vari_Dontshowfeedbackagain = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Behaviourdefine_base64 = "H4sIAAAAAAAEAI1WW2/aMBR+n7T/gHgvIbTd1IpSQRgr4tKutNL6VLnJIXhx7MyXptmvnzFJSUxI+wbnu5zvOMdR+tdvMWm9AheY0au22+m2rwdfv/RHsEGvmCk+hjWm0NIsKi7fBL5qb6RMLh0nTdNOetphPHR63a7r/F7MV/4GYnSCqZCI+tB+VwUfq9q6a6vV9xCdMF+JB46owFKHGmU/qAQ+g2wguYK+00g5avKYjFlKPcUF+8DLZlYsR4wHwD1GGPc2iIZgGR3iRm5jQx6+DE7c3tlFt/fd7Tu1eNH4l8JSjyh8lEAlug0Ugl1nj1HJGTF+e00NVsimIWUcpjRR8h7+Kswh2OvqwEJ4r/SxTc4r2Uq1gqZ/Dzkg3ZSgRJROzgaMYIEwQUHAQQgPCPEI9qMHtq0CX0nE5WCNiNAGHxNzP6F3ZM5YpJJtKtc922qrxRLzNgG6450XvKK0G0gfIbzJGyC5XU/PUq0ZnjWc4fZO+05N3fCHSrIFYIHwKsVS3w1eDFqDGMVUTED/HRIyRhIVbKuaPwRCIJgwHh/v0sg5klCzl0xPII5H3VPyzHp5/egnK93d9+CHkNHcISFSfUdusJCMZy2nWp7rZ7RgAV5jCJ7ioID1ItAFUDWNUQgjJODbmQ2ZLVmiWK8kS3wkIdT2zy+R2S2LURHe43AjVziAiaL+9u0xcPeaQzB/FXB/e//WmMe31JARVR5hAtMwvxPNnJ1NxOKFEnLG/qEVEPDlXB+LPdn4boJePaJPyIRf/vE7+dUXneWpQbulyBbdshpFDU6jqGpUJh/4dN3nGaIczwBLhOuDz7KGbrOs2q1MPvD5XDfNmopHfREzvTz5cziOW/InJDFtyGvwamRLUmf4meAeaWjrkWrPMtnyWWUswu+o6dN3rK+AwX9Db0j/LggAAA==";
                    model_cvitem.Vari_Colordefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E9zzs8rKcrPcc7PyS9ySU3LzEtVAKrPK7aqKM60VcooKSmw0tcvLy/XKzfWyy9K1zcyMDDUj/D1CU7OSM1N1M3MKy5JzEtOVYLrSiGsS0lB3w4AaXGCjpcAAAA=";
                    model_cvitem.Vari_N3spreadcolordefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zMw4uKEpNTAnNS0kt8snMS3XOz8kvcklNAzIVgDrziq0qijNtlTJKSgqs9PXLy8v1yo318ovS9Y0MDAz1I3x9gpMzUnMTdTPziksS85JTleC6UgjrUlLQtwMAoBcTOKEAAAA=";
                    model_cvitem.Vari_Menudesigndefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zTc0rdUktzkzPc0lNy8xLVQCqziu2qijOtFXKKCkpsNLXLy8v1ys31ssvStc3MjAw1I/w9QlOzkjNTdTNzCsuScxLTlWC60ohrEtJQd8OAJAcjj6VAAAA";
                    model_cvitem.Vari_Smtpdefine_base64 = "H4sIAAAAAAAEAIWRz2vCMBTH7wP/h5C7fa1jF4mRqRdhboU62DXU17WgeSXJFv3vTSPtSi+75fvj8XmPiPX1cma/aGxDesWzJOVrOXsSxeGY77BqNLJQ0HZ5tc2K1861SwDvfeKfEzLfsEjTDL4Ob0VZ40XNG22d0iXyYer0/xQPQMYiskATdmHw5+RknFy8CBhEjLakNZbueGtR7ki7rTIoYOzG2uuPq6N6p+4pYDAGwKedAJW1nsyp9/Y2p3aDFRnsYlmpsw2oqR27+cfkgmDEnbMsFdCLPhmTu2wMftz7+AF5B/390ZmlAQAA";
                    model_cvitem.Vari_Sendtargetsiteid = "";
                    model_cvitem.Vari_Hysearchdefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zqAxOTSxKzggAKsjPS8xxzs9LySwBKlUAassrtqoozrRVyigpKbDS1y8vL9crN9bLL0rXNzIwMNSP8PUJTs5IzU3UzcwrLknMS05VgutKIaxLSUHfDgDu3pVrngAAAA==";
                    model_cvitem.Vari_Csvsetting_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E9zDg4LTi0pycxLVwCqyyu2qijOtFXKKCkpsNLXLy8v1ys31ssvStc3MjAw1I/w9QlOzkjNTdTNzCsuScxLTlWC60ohrEtJQd8OAIz7H0GPAAAA";
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

            // Dim model_cvitem As New Njc.Model.Profile_logonuser_Model       '移行値格納用モデル初期化
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
            // Dim tblname As String = "profile_logonuser"
            // Dim fldnamegrp As String = "logonuser_no,logonuser_name,logonuser_namesjis,logonuser_kana,logongroup_no," & _
            // "tel1,tel2,mobiletel1,fax1,mailaddress1," & _
            // "mobilemailaddress1,takkenmenkyo,biko,denylogon,noassign," & _
            // "password,behaviourdefine_base64,colordefine_base64,n3spreadcolordefine_base64,menudesigndefine_base64," & _
            // "smtpdefine_base64,history,sendtargetsiteid,windowsuser_name,hysearchdefine_base64," & _
            // "csvsetting_base64,dontshowfeedbackagain"

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

            // '移行値取得
            // For cntjj = 1 To columncnt

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "賃貸革命ログオンユーザNo"
            // .Vari_Logonuser_no = fldvalue
            // Case "賃貸革命ログオンユーザ名"
            // .Vari_Logonuser_name = fldvalue
            // Case "賃貸革命ログオンユーザ名(SJIS)"
            // .Vari_Logonuser_namesjis = fldvalue
            // Case "賃貸革命ログオンユーザ名カナ"
            // .Vari_Logonuser_kana = fldvalue
            // Case "賃貸革命ログオングループNo"
            // '20160928 自社担当者移行処理の修正 -chg sta
            // '.Vari_Logongroup_no = fldvalue
            // .Vari_Logongroup_no = IIf(fldvalue = "", "99001", fldvalue)
            // '20160928 自社担当者移行処理の修正 -chg end
            // Case "TEL1"
            // .Vari_Tel1 = fldvalue
            // Case "TEL2"
            // .Vari_Tel2 = fldvalue
            // Case "携帯電話1"
            // .Vari_Mobiletel1 = fldvalue
            // Case "FAX1"
            // .Vari_Fax1 = fldvalue
            // Case "メールアドレス1"
            // .Vari_Mailaddress1 = fldvalue
            // Case "携帯メールアドレス1"
            // .Vari_Mobilemailaddress1 = fldvalue
            // Case "宅建免許番号"
            // .Vari_Takkenmenkyo = fldvalue
            // Case "備考"
            // .Vari_Biko = fldvalue
            // Case "担当者ログオン禁止フラグ"
            // .Vari_Denylogon = fldvalue
            // Case "使用しないフラグ"
            // .Vari_Noassign = fldvalue
            // Case "ログオンパスワード"
            // .Vari_Password = fldvalue
            // Case "windowsログオンユーザー名"
            // .Vari_Windowsuser_name = fldvalue
            // Case "意見画面表示フラグ"
            // .Vari_Dontshowfeedbackagain = fldvalue
            // End Select

            // Next

            // '固定値
            // .Vari_Behaviourdefine_base64 = "H4sIAAAAAAAEAI1WW2/aMBR+n7T/gHgvIbTd1IpSQRgr4tKutNL6VLnJIXhx7MyXptmvnzFJSUxI+wbnu5zvOMdR+tdvMWm9AheY0au22+m2rwdfv/RHsEGvmCk+hjWm0NIsKi7fBL5qb6RMLh0nTdNOetphPHR63a7r/F7MV/4GYnSCqZCI+tB+VwUfq9q6a6vV9xCdMF+JB46owFKHGmU/qAQ+g2wguYK+00g5avKYjFlKPcUF+8DLZlYsR4wHwD1GGPc2iIZgGR3iRm5jQx6+DE7c3tlFt/fd7Tu1eNH4l8JSjyh8lEAlug0Ugl1nj1HJGTF+e00NVsimIWUcpjRR8h7+Kswh2OvqwEJ4r/SxTc4r2Uq1gqZ/Dzkg3ZSgRJROzgaMYIEwQUHAQQgPCPEI9qMHtq0CX0nE5WCNiNAGHxNzP6F3ZM5YpJJtKtc922qrxRLzNgG6450XvKK0G0gfIbzJGyC5XU/PUq0ZnjWc4fZO+05N3fCHSrIFYIHwKsVS3w1eDFqDGMVUTED/HRIyRhIVbKuaPwRCIJgwHh/v0sg5klCzl0xPII5H3VPyzHp5/egnK93d9+CHkNHcISFSfUdusJCMZy2nWp7rZ7RgAV5jCJ7ioID1ItAFUDWNUQgjJODbmQ2ZLVmiWK8kS3wkIdT2zy+R2S2LURHe43AjVziAiaL+9u0xcPeaQzB/FXB/e//WmMe31JARVR5hAtMwvxPNnJ1NxOKFEnLG/qEVEPDlXB+LPdn4boJePaJPyIRf/vE7+dUXneWpQbulyBbdshpFDU6jqGpUJh/4dN3nGaIczwBLhOuDz7KGbrOs2q1MPvD5XDfNmopHfREzvTz5cziOW/InJDFtyGvwamRLUmf4meAeaWjrkWrPMtnyWWUswu+o6dN3rK+AwX9Db0j/LggAAA=="
            // .Vari_Colordefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E9zzs8rKcrPcc7PyS9ySU3LzEtVAKrPK7aqKM60VcooKSmw0tcvLy/XKzfWyy9K1zcyMDDUj/D1CU7OSM1N1M3MKy5JzEtOVYLrSiGsS0lB3w4AaXGCjpcAAAA="
            // .Vari_N3spreadcolordefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zMw4uKEpNTAnNS0kt8snMS3XOz8kvcklNAzIVgDrziq0qijNtlTJKSgqs9PXLy8v1yo318ovS9Y0MDAz1I3x9gpMzUnMTdTPziksS85JTleC6UgjrUlLQtwMAoBcTOKEAAAA="
            // .Vari_Menudesigndefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zTc0rdUktzkzPc0lNy8xLVQCqziu2qijOtFXKKCkpsNLXLy8v1ys31ssvStc3MjAw1I/w9QlOzkjNTdTNzCsuScxLTlWC60ohrEtJQd8OAJAcjj6VAAAA"
            // .Vari_Smtpdefine_base64 = "H4sIAAAAAAAEAIWRz2vCMBTH7wP/h5C7fa1jF4mRqRdhboU62DXU17WgeSXJFv3vTSPtSi+75fvj8XmPiPX1cma/aGxDesWzJOVrOXsSxeGY77BqNLJQ0HZ5tc2K1861SwDvfeKfEzLfsEjTDL4Ob0VZ40XNG22d0iXyYer0/xQPQMYiskATdmHw5+RknFy8CBhEjLakNZbueGtR7ki7rTIoYOzG2uuPq6N6p+4pYDAGwKedAJW1nsyp9/Y2p3aDFRnsYlmpsw2oqR27+cfkgmDEnbMsFdCLPhmTu2wMftz7+AF5B/390ZmlAQAA"
            // .Vari_Sendtargetsiteid = ""
            // .Vari_Hysearchdefine_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E/zqAxOTSxKzggAKsjPS8xxzs9LySwBKlUAassrtqoozrRVyigpKbDS1y8vL9crN9bLL0rXNzIwMNSP8PUJTs5IzU3UzcwrLknMS05VgutKIaxLSUHfDgDu3pVrngAAAA=="
            // .Vari_Csvsetting_base64 = "H4sIAAAAAAAEALOxr8jNUShLLSrOzM+zVTLUM1Cyt+PlsnEsKkqs9E9zDg4LTi0pycxLVwCqyyu2qijOtFXKKCkpsNLXLy8v1ys31ssvStc3MjAw1I/w9QlOzkjNTdTNzCsuScxLTlWC60ohrEtJQd8OAIz7H0GPAAAA"
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


                string tmp_sql = " SELECT logonuser_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);


            }

        }

    }

    #endregion

    #region 自社メモ情報

    public class Jisyadata_memo_Repository
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

                var model_cvitem = new Model.Jisyadata_memo_Model();          // 移行値格納用モデル初期化
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
                string tblname_base = "jisyadata";
                string tblname = "jisyadata_memo";
                string fldnamegrp = "jisya_no,memo_no,memo,history";

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
                    model_cvitem.Vari_Jisya_no = tmp_keymain;

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
                            string log_key = "jisyadata_memo-jisya_no";
                            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE;
                            hash_keylog.Clear();
                            hash_keylog.Add(log_key, log_value);
                        }

                        // 重複チェック
                        if (keychkflg & list_chkduplicate.Contains(tmp_keymain))
                        {
                            keychkflg = false;
                            // ログ出力
                            string log_key = "jisyadata_memo-jisya_no";
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

            // Dim model_cvitem As New Njc.Model.Jisyadata_memo_Model          '移行値格納用モデル初期化
            // Dim keycol As Integer = 1                                       'キー列
            // Dim tmp_cvrowcnt As Integer = 0                                 '20160928 メモ関連の移行件数表示修正 -add

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
            // Dim tblname_base As String = "jisyadata"
            // Dim tblname As String = "jisyadata_memo"
            // Dim fldnamegrp As String = "jisya_no,memo_no,memo,history"

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
            // Dim fldname_keymain As String = headervalue(startrow - 1, keycol)

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
            // Dim tmp_keymain As String = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol))
            // .Vari_Jisya_no = tmp_keymain

            // '備考カウント初期化
            // Dim cvitemcnt As Integer = 0

            // '移行値取得
            // For cntjj = 2 To columncnt

            // 'サブキー取得
            // Dim tmp_keysub As String = (cntjj - 1).ToString

            // '全キー取得
            // Dim fldvalue_key As String = tmp_keymain & "-" & tmp_keysub

            // 'ログ出力用データ格納(サブキーフィールド)
            // Dim fldname_keysub As String = headervalue(startrow - 1, cntjj)
            // Dim str_logkey As String = fldname_keymain & " = " & tmp_keymain & "、" & fldname_keysub

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If
            // .Vari_Memo = fldvalue

            // '固定値
            // .Vari_Memo_no = (cntjj - 1).ToString
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

            // '--------------------------------------------------------------------
            // '移行した備考が1データ以上ある場合移行したレコードの数を更新する
            // '--------------------------------------------------------------------
            // '20160928 メモ関連の移行件数表示修正 -chg sta
            // 'tmp_cvcnt = tmp_cvcnt + 1
            // If cvitemcnt > 0 Then
            // tmp_cvcnt = tmp_cvcnt + 1
            // End If
            // '20160928 メモ関連の移行件数表示修正 -chg end
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
            // '20160928 メモ関連の移行件数表示修正 -chg sta
            // 'midrowcnt = rowcnt
            // midrowcnt = tmp_cvrowcnt
            // '20160928 メモ関連の移行件数表示修正 -chg end
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

                string tmp_sql = " SELECT jisya_no FROM " + tblname;
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

                string tmp_sql = " SELECT jisya_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 振込依頼人情報

    public class M_fb_sgfirai_Repository
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

                var model_cvitem = new Model.M_fb_sgfirai_Model();            // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************

                // 2016.04.26 メインの方へも反映させる修正 -add sta
                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Jisyakoza();
                SetRelItemToObject.Set_RelData_FBInfo();        // 20160609 紐付設定値取得に伴う修正 -add
                // 2016.04.26 メインの方へも反映させる修正 -add end

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
                string tblname = "m_fb_sgfirai";
                // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg sta
                // Dim fldnamegrp As String = "sgfirai_no,sgfirai_name,jisya_no,jisya_kozano,sgfirainin_code," & _
                // "sgfirainin_kana,biko_kihon,sgfzenfmt_no,file_sosin,crlf," & _
                // "changewo_flg,biko_data,keisandefault_flg,futancyousei_flg,biko_tesu," & _
                // "history,sgfdata_createflg,useflg"
                string fldnamegrp = "sgfirai_no,sgfirai_name,jisya_no,jisya_kozano,sgfirainin_code," + "sgfirainin_kana,biko,sgfzenfmt_no,file_sosin,changewo_flg," + "futancyousei_flg,history,sgfdata_createflg,useflg";

                // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg end

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
                    string tmp_kinyuno = "";
                    string tmp_kinyutenno = "";
                    string tmp_kozasyu = "";      // 20160609 紐付取得用に追加 -add
                    string tmp_kozabango = "";
                    string tmp_jisyano = "";      // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
                    string tmp_jisyakozano = "";  // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add

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
                            case "振込依頼人No":
                                {
                                    model_cvitem.Vari_Sgfirai_no = fldvalue;
                                    model_cvitem.Vari_Sgfzenfmt_no = fldvalue;   // 20160609 紐付設定値取得に伴う修正 -add
                                    break;
                                }
                            case "振込依頼人名":
                                {
                                    model_cvitem.Vari_Sgfirai_name = fldvalue;
                                    break;
                                }
                            // 2016.04.26 メインの方へも反映させる修正 -chg sta
                            // Case "自社支店No."
                            // .Vari_Jisya_no = fldvalue.Trim
                            // Case "自社口座No"
                            // .Vari_Jisya_kozano = fldvalue.Trim
                            case "口座種別": // 20160609 紐付取得用に追加 -add
                                {
                                    tmp_kozasyu = fldvalue;
                                    break;
                                }
                            case "金融機関No":
                                {
                                    tmp_kinyuno = fldvalue;
                                    break;
                                }
                            case "金融機関支店No":
                                {
                                    tmp_kinyutenno = fldvalue;
                                    break;
                                }
                            case "口座番号":
                                {
                                    tmp_kozabango = fldvalue;
                                    break;
                                }
                            // 2016.04.26 メインの方へも反映させる修正 -chg end
                            case "振込依頼人コード":
                                {
                                    model_cvitem.Vari_Sgfirainin_code = fldvalue;
                                    break;
                                }
                            case "振込依頼人カナ":
                                {
                                    model_cvitem.Vari_Sgfirainin_kana = fldvalue;
                                    break;
                                }
                            case "備考(基本情報)":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;             // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -chg (Vari_Biko_kihon → Vari_Biko)
                                    break;
                                }
                            // 20160609 紐付設定値取得に伴う修正 -del sta
                            // Case "総合振込全銀フォーマットNo"
                            // .Vari_Sgfzenfmt_no = fldvalue
                            // 20160609 紐付設定値取得に伴う修正 -del end
                            case "送信ファイルパス":
                                {
                                    model_cvitem.Vari_File_sosin = fldvalue;
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -del sta
                            // Case "改行有無(0:改行しない 1:改行する)"
                            // .Vari_Crlf = fldvalue
                            // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -del end
                            case "ヲ変換フラグ(0:変換しない 1:変換する)":
                                {
                                    model_cvitem.Vari_Changewo_flg = fldvalue;
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -del sta
                            // Case "備考(データ)"
                            // .Vari_Biko_data = fldvalue
                            // Case "計算時の初期値フラグ(0:初期値としない 1:初期値とする)"
                            // .Vari_Keisandefault_flg = fldvalue
                            // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -del end
                            case "負担調整フラグ(0:調整しない 1:調整する)":
                                {
                                    model_cvitem.Vari_Futancyousei_flg = fldvalue;
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -del sta
                            // Case "備考(手数料)"
                            // .Vari_Biko_tesu = fldvalue
                            // 20160519 EXEUpdateに伴う修正 振込依頼人情報 -del end
                            case "総合振込データ生成フラグ":
                                {
                                    model_cvitem.Vari_Sgfdata_createflg = fldvalue;
                                    break;
                                }
                            // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
                            case "自社支店No":  // 20160929 汎用→既存コピー処理改善対応 自社支店No. → 自社支店No -chg
                                {
                                    tmp_jisyano = fldvalue;
                                    break;
                                }
                            case "自社口座No":
                                {
                                    tmp_jisyakozano = fldvalue;
                                    break;
                                }
                                // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end
                        }

                    }
                    // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg sta
                    // '20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                    // ''20160609 紐付取得用に追加 -chg sta
                    // ' ''2016.04.26 メインの方へも反映させる修正 -add sta
                    // ' ''自社No/自社口座No取得のため口座情報を格納する
                    // ''.Vari_Jisya_no = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                    // ''.Vari_Jisya_kozano = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                    // ' ''2016.04.26 メインの方へも反映させる修正 -add end
                    // 'Dim tmp_jisyainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
                    // '.Vari_Jisya_no = tmp_jisyainfo
                    // '.Vari_Jisya_kozano = tmp_jisyainfo
                    // ''20160609 紐付取得用に追加 -chg end
                    // Dim tmp_kozamoto As String = "振込依頼人マスタ"
                    // Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Sgfirai_no
                    // .Vari_Jisya_no = tmp_jisyainfo
                    // .Vari_Jisya_kozano = tmp_jisyainfo
                    // '20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                    switch (CommonModule.CNVNO)
                    {
                        case (int)CommonModule.ConvertTypes._汎用:
                            {
                                model_cvitem.Vari_Sgfzenfmt_no = "";
                                // 20160928 自社口座取得方法の修正 -chg sta
                                // .Vari_Jisya_no = tmp_jisyano
                                // .Vari_Jisya_kozano = tmp_jisyakozano
                                model_cvitem.Vari_Jisya_no = tmp_jisyano + "-" + tmp_jisyakozano;
                                model_cvitem.Vari_Jisya_kozano = tmp_jisyano + "-" + tmp_jisyakozano;
                                break;
                            }
                            // 20160928 自社口座取得方法の修正 -chg end
                    }
                    // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg end
                    // 固定値
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 20160928 自社口座取得方法の修正 -add sta
                    tmp_hash["jisya_no"] = tmp_jisyano;
                    tmp_hash["jisya_kozano"] = tmp_jisyakozano;
                    // 20160928 自社口座取得方法の修正 -add end

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

            // Dim model_cvitem As New Njc.Model.M_fb_sgfiraitesu_Model        '移行値格納用モデル初期化
            // Dim keycol_main As Integer = 1                                  'メインキー列
            // Dim keycol_sub As Integer = 2                                   'サブキー列

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
            // Dim tblname_base As String = "m_fb_sgfirai"
            // Dim tblname As String = "m_fb_sgfiraitesu"
            // Dim fldnamegrp As String = "sgfirai_no,rec_no,from_gak,to_gak,doukoudouten_densingak," & _
            // "doukoudouten_bunsyogak,doukoutaten_densingak,doukoutaten_bunsyogak,takou_densingak,takou_bunsyogak"

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
            // Case "振込依頼人No"
            // .Vari_Sgfirai_no = fldvalue
            // Case "レコードNo"
            // .Vari_Rec_no = fldvalue
            // Case "金額From"
            // .Vari_From_gak = fldvalue
            // Case "金額To"
            // .Vari_To_gak = fldvalue
            // Case "同行同支店手数料(電信扱)"
            // .Vari_Doukoudouten_densingak = fldvalue
            // Case "同行同支店手数料(文書扱)"
            // .Vari_Doukoudouten_bunsyogak = fldvalue
            // Case "同行他支店手数料(電信扱)"
            // .Vari_Doukoutaten_densingak = fldvalue
            // Case "同行他支店手数料(文書扱)"
            // .Vari_Doukoutaten_bunsyogak = fldvalue
            // Case "他行手数料(電信扱)"
            // .Vari_Takou_densingak = fldvalue
            // Case "他行手数料(文書扱)"
            // .Vari_Takou_bunsyogak = fldvalue
            // End Select

            // Next

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


                string tmp_sql = " SELECT jisya_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);


            }

        }

    }

    #endregion

    #region 振込手数料情報

    public class M_fb_sgfiraitesu_Repository
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

                var model_cvitem = new Model.M_fb_sgfiraitesu_Model();        // 移行値格納用モデル初期化
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
                string tblname_base = "m_fb_sgfirai";
                string tblname = "m_fb_sgfiraitesu";
                string fldnamegrp = "sgfirai_no,rec_no,from_gak,to_gak,doukoudouten_densingak," + "doukoudouten_bunsyogak,doukoutaten_densingak,doukoutaten_bunsyogak,takou_densingak,takou_bunsyogak";

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
                            case "振込依頼人No":
                                {
                                    model_cvitem.Vari_Sgfirai_no = fldvalue;
                                    break;
                                }
                            case "レコードNo":
                                {
                                    model_cvitem.Vari_Rec_no = fldvalue;
                                    break;
                                }
                            case "金額From":
                                {
                                    model_cvitem.Vari_From_gak = fldvalue;
                                    break;
                                }
                            case "金額To":
                                {
                                    model_cvitem.Vari_To_gak = fldvalue;
                                    break;
                                }
                            case "同行同支店手数料(電信扱)":
                                {
                                    model_cvitem.Vari_Doukoudouten_densingak = fldvalue;
                                    break;
                                }
                            case "同行同支店手数料(文書扱)":
                                {
                                    model_cvitem.Vari_Doukoudouten_bunsyogak = fldvalue;
                                    break;
                                }
                            case "同行他支店手数料(電信扱)":
                                {
                                    model_cvitem.Vari_Doukoutaten_densingak = fldvalue;
                                    break;
                                }
                            case "同行他支店手数料(文書扱)":
                                {
                                    model_cvitem.Vari_Doukoutaten_bunsyogak = fldvalue;
                                    break;
                                }
                            case "他行手数料(電信扱)":
                                {
                                    model_cvitem.Vari_Takou_densingak = fldvalue;
                                    break;
                                }
                            case "他行手数料(文書扱)":
                                {
                                    model_cvitem.Vari_Takou_bunsyogak = fldvalue;
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

                string tmp_sql = " SELECT sgfirai_no FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,sgfirai_no) + '-' + CONVERT(varchar,rec_no) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 口座振替情報

    public class M_fb_fkaejyoho_Repository
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

                var model_cvitem = new Model.M_fb_fkaejyoho_Model();          // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************

                // 2016.04.26 メインの方へも反映させる修正 -add sta
                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Jisyakoza();
                SetRelItemToObject.Set_RelData_FBInfo();        // 20160609 紐付設定値取得に伴う修正 -add
                // 2016.04.26 メインの方へも反映させる修正 -add end

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
                string tblname = "m_fb_fkaejyoho";
                // 20161012 革命10アップデートに伴う修正 「saifkae_use」を追加 -add
                // 20160519 EXEUpdateに伴う修正 口座振替情報 -chg sta
                // Dim fldnamegrp As String = "fkae_no,fkae_name,fkae_kana,servicetype,fkae_fb_fkomiraino," & _
                // "fkae_fb_fkomiraikana,kamei_no,jlease_nyukinkbn,jisya_no,fkae_fb_nkinukekozano," & _
                // "hikiotosibi,tesu_gak,biko,fkae_fb_orgfmtno,datasort," & _
                // "keiyakusyano_syuturyoku,crlf,wo_mojihenkan,file_sosin,file_jyusin," & _
                // "fdsakusei,seikyu_tani,seikyu_taino,seikyu_tukitani,tesu_nyukinkanri," & _
                // "tesu_kurikosi,syogorule,yutyobango,history,useflg"
                string fldnamegrp = "fkae_no,fkae_name,fkae_kana,fkae_fb_fkomiraino,fkae_fb_fkomiraikana," + "kamei_no,jlease_nyukinkbn,jisya_no,fkae_fb_nkinukekozano,hikiotosibi," + "tesu_gak,biko,fkae_fb_orgfmtno,datasort,keiyakusyano_syuturyoku," + "crlf,wo_mojihenkan,file_sosin,file_jyusin,seikyu_tani," + "seikyu_taino,seikyu_tukitani,tesu_nyukinkanri,tesu_kurikosi,syogorule," + "yutyobango,history,useflg,multi_flg,yutyo_appendcode," + "yutyo_usefkomcode,resultfile_notuse,saifkae_use";





                // 20160519 EXEUpdateに伴う修正 口座振替情報 -chg end

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
                    string tmp_kinyuno = "";
                    string tmp_kinyutenno = "";
                    string tmp_kozasyu = "";      // 20160609 紐付取得用に追加 -add
                    string tmp_kozabango = "";
                    string tmp_jisyano = "";      // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
                    string tmp_jisyakozano = "";  // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add

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
                            case "振替情報No":                          // 20160929 汎用→既存コピー処理改善対応 振替情報No. → 振替情報No -chg
                                {
                                    model_cvitem.Vari_Fkae_no = fldvalue;
                                    model_cvitem.Vari_Fkae_fb_orgfmtno = fldvalue;       // 20160609 紐付設定値取得に伴う修正 -add
                                    break;
                                }
                            case "振替情報名称":
                                {
                                    model_cvitem.Vari_Fkae_name = fldvalue;
                                    break;
                                }
                            case "振替情報カナ名称":
                                {
                                    model_cvitem.Vari_Fkae_kana = fldvalue;
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 口座振替情報 -del sta
                            // Case "サービスタイプ　1:オリコ　2:ジャックス　3:アプラス"
                            // .Vari_Servicetype = fldvalue
                            // 20160519 EXEUpdateに伴う修正 口座振替情報 -del end
                            case "振込依頼人":
                                {
                                    model_cvitem.Vari_Fkae_fb_fkomiraino = fldvalue;
                                    break;
                                }
                            case "振込依頼人カナ名":
                                {
                                    model_cvitem.Vari_Fkae_fb_fkomiraikana = fldvalue;
                                    break;
                                }
                            case "加盟店No 依頼人番号と共用":          // 20160929 汎用→既存コピー処理改善対応 加盟店No.　依頼人番号と共用 → 加盟店No 依頼人番号と共用 -chg
                                {
                                    model_cvitem.Vari_Kamei_no = fldvalue;
                                    break;
                                }
                            case "ジェイリース　入金区分　1:変更なし・・・・":
                                {
                                    model_cvitem.Vari_Jlease_nyukinkbn = fldvalue;
                                    break;
                                }
                            // 2016.04.26 メインの方へも反映させる修正 -chg sta
                            // Case "自社支店No"
                            // .Vari_Jisya_no = fldvalue
                            // Case "振替先口座No."
                            // .Vari_Fkae_fb_nkinukekozano = fldvalue
                            case "金融機関No":
                                {
                                    tmp_kinyuno = fldvalue;
                                    break;
                                }
                            case "金融機関支店No":
                                {
                                    tmp_kinyutenno = fldvalue;
                                    break;
                                }
                            case "口座種別": // 20160609 紐付取得用に追加 -add
                                {
                                    tmp_kozasyu = fldvalue;
                                    break;
                                }
                            case "口座番号":
                                {
                                    tmp_kozabango = fldvalue;
                                    break;
                                }
                            // 2016.04.26 メインの方へも反映させる修正 -chg end
                            case "引落日":
                                {
                                    model_cvitem.Vari_Hikiotosibi = fldvalue;
                                    break;
                                }
                            case "手数料　請求額":
                                {
                                    model_cvitem.Vari_Tesu_gak = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            // 20160609 紐付設定値取得に伴う修正 -del sta
                            // Case "オリジナルフォーマットNo"
                            // .Vari_Fkae_fb_orgfmtno = fldvalue
                            // 20160609 紐付設定値取得に伴う修正 -del end
                            case "データ並び替え　0:物件番号と部屋番号　1:契約者番号":
                                {
                                    model_cvitem.Vari_Datasort = fldvalue;
                                    break;
                                }
                            case "契約者Noを契約者番号に出力する  0:出力しない　1:出力する":    // 20160929 汎用→既存コピー処理改善対応 契約者No.を契約者番号に出力する  0:出力しない　1:出力する → 契約者Noを契約者番号に出力する  0:出力しない　1:出力する -chg
                                {
                                    model_cvitem.Vari_Keiyakusyano_syuturyoku = fldvalue;
                                    break;
                                }
                            case "改行コード(CRLF)出力　0:なし　1:あり":
                                {
                                    model_cvitem.Vari_Crlf = fldvalue;
                                    break;
                                }
                            case "「ｦ」→「ｵ」変換　0:しない　1:する":
                                {
                                    model_cvitem.Vari_Wo_mojihenkan = fldvalue;
                                    break;
                                }
                            case "送信ファイルパス名":
                                {
                                    model_cvitem.Vari_File_sosin = fldvalue;
                                    break;
                                }
                            case "受信ファイルパス名":
                                {
                                    model_cvitem.Vari_File_jyusin = fldvalue;
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 口座振替情報 -del sta
                            // Case "2枚ＦＤ作成手順を使用する　0:しない　1:する"
                            // .Vari_Fdsakusei = fldvalue
                            // 20160519 EXEUpdateに伴う修正 口座振替情報 -del end
                            case "請求のまとめ方　0:請求先単位　1:契約単位":
                                {
                                    model_cvitem.Vari_Seikyu_tani = fldvalue;
                                    break;
                                }
                            case "滞納分請求　0:しない　1:する":
                                {
                                    model_cvitem.Vari_Seikyu_taino = fldvalue;
                                    break;
                                }
                            case "滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける":
                                {
                                    model_cvitem.Vari_Seikyu_tukitani = fldvalue;
                                    break;
                                }
                            case "手数料の請求入金管理をする　0:しない　1:する":
                                {
                                    model_cvitem.Vari_Tesu_nyukinkanri = fldvalue;
                                    break;
                                }
                            case "未入金の手数料は次回請求に加える　0:加えない　1:加える":
                                {
                                    model_cvitem.Vari_Tesu_kurikosi = fldvalue;
                                    break;
                                }
                            case "照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合":
                                {
                                    model_cvitem.Vari_Syogorule = fldvalue;
                                    break;
                                }
                            case "ゆうちょ銀行金融機関番号指定　不使用の場合はNULL":
                                {
                                    model_cvitem.Vari_Yutyobango = fldvalue;
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 口座振替情報 -add sta
                            case "マルチヘッダー形式フラグ":
                                {
                                    model_cvitem.Vari_Yutyobango = fldvalue;
                                    break;
                                }
                            case "ゆうちょ銀行付加コード":
                                {
                                    model_cvitem.Vari_Yutyobango = fldvalue;
                                    break;
                                }
                            case "振替でのゆうちょ銀行コードに振込用変換を利用するフラグ":
                                {
                                    model_cvitem.Vari_Yutyobango = fldvalue;
                                    break;
                                }
                            case "振替入金処理時に受信ファイルを読み込まないフラグ":
                                {
                                    model_cvitem.Vari_Yutyobango = fldvalue;
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 口座振替情報 -add end
                            // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
                            case "自社支店No":
                                {
                                    tmp_jisyano = fldvalue;
                                    break;
                                }
                            case "振替先口座No":                 // 20160929 汎用→既存コピー処理改善対応 振替先口座No. → 振替先口座No -chg
                                {
                                    tmp_jisyakozano = fldvalue;
                                    break;
                                }
                            // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end
                            case "再振替対応利用フラグ":   // 20161012 革命10アップデートに伴う修正 -add
                                {
                                    model_cvitem.Vari_Saifkae_use = fldvalue;
                                    break;
                                }
                        }

                    }
                    // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg sta
                    // '20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                    // ''20160609 紐付取得用に追加 -chg sta
                    // ' ''2016.04.26 メインの方へも反映させる修正 -add sta
                    // ' ''自社No/自社口座No取得のため口座情報を格納する
                    // ''.Vari_Jisya_no = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                    // ''.Vari_Fkae_fb_nkinukekozano = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                    // ' ''2016.04.26 メインの方へも反映させる修正 -add end
                    // 'Dim tmp_jisyainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
                    // '.Vari_Jisya_no = tmp_jisyainfo
                    // '.Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
                    // ''20160609 紐付取得用に追加 -chg end
                    // Dim tmp_kozamoto As String = "口座振替マスタ"
                    // Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Fkae_no
                    // .Vari_Jisya_no = tmp_jisyainfo
                    // .Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
                    // '20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                    switch (CommonModule.CNVNO)
                    {
                        case (int)CommonModule.ConvertTypes._汎用:
                            {
                                model_cvitem.Vari_Fkae_fb_orgfmtno = "";
                                // 20160928 自社口座取得方法の修正 -chg sta
                                // .Vari_Jisya_no = tmp_jisyano
                                // .Vari_Fkae_fb_nkinukekozano = tmp_jisyakozano
                                model_cvitem.Vari_Jisya_no = tmp_jisyano + "-" + tmp_jisyakozano;
                                model_cvitem.Vari_Fkae_fb_nkinukekozano = tmp_jisyano + "-" + tmp_jisyakozano;
                                break;
                            }
                            // 20160928 自社口座取得方法の修正 -chg end
                    }
                    // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg end
                    // 固定値
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 20160928 自社口座取得方法の修正 -add sta
                    tmp_hash["jisya_no"] = tmp_jisyano;
                    tmp_hash["fkae_fb_nkinukekozano"] = tmp_jisyakozano;
                    // 20160928 自社口座取得方法の修正 -add end

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

            // Dim model_cvitem As New Njc.Model.M_fb_fkaejyoho_Model          '移行値格納用モデル初期化
            // Dim keycol As Integer = 1                                       'キー列

            // '************************
            // '作業準備
            // '************************

            // '2016.04.26 メインの方へも反映させる修正 -add sta
            // '紐付けデータ取得
            // Call SetRelItemToObject.Set_RelData_Jisyakoza()
            // Call SetRelItemToObject.Set_RelData_FBInfo()        '20160609 紐付設定値取得に伴う修正 -add
            // '2016.04.26 メインの方へも反映させる修正 -add end

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
            // If Not rtn Then
            // Return rtn
            // End If

            // 'テーブル名/フィールド名セット
            // Dim tblname As String = "m_fb_fkaejyoho"
            // '20160519 EXEUpdateに伴う修正 口座振替情報 -chg sta
            // 'Dim fldnamegrp As String = "fkae_no,fkae_name,fkae_kana,servicetype,fkae_fb_fkomiraino," & _
            // '                           "fkae_fb_fkomiraikana,kamei_no,jlease_nyukinkbn,jisya_no,fkae_fb_nkinukekozano," & _
            // '                           "hikiotosibi,tesu_gak,biko,fkae_fb_orgfmtno,datasort," & _
            // '                           "keiyakusyano_syuturyoku,crlf,wo_mojihenkan,file_sosin,file_jyusin," & _
            // '                           "fdsakusei,seikyu_tani,seikyu_taino,seikyu_tukitani,tesu_nyukinkanri," & _
            // '                           "tesu_kurikosi,syogorule,yutyobango,history,useflg"
            // Dim fldnamegrp As String = "fkae_no,fkae_name,fkae_kana,fkae_fb_fkomiraino,fkae_fb_fkomiraikana," & _
            // "kamei_no,jlease_nyukinkbn,jisya_no,fkae_fb_nkinukekozano,hikiotosibi," & _
            // "tesu_gak,biko,fkae_fb_orgfmtno,datasort,keiyakusyano_syuturyoku," & _
            // "crlf,wo_mojihenkan,file_sosin,file_jyusin,seikyu_tani," & _
            // "seikyu_taino,seikyu_tukitani,tesu_nyukinkanri,tesu_kurikosi,syogorule," & _
            // "yutyobango,history,useflg,multi_flg,yutyo_appendcode," & _
            // "yutyo_usefkomcode,resultfile_notuse"
            // '20160519 EXEUpdateに伴う修正 口座振替情報 -chg end

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

            // '2016.04.26 メインの方へも反映させる修正 -add sta
            // '作業用変数
            // Dim tmp_kinyuno As String = ""
            // Dim tmp_kinyutenno As String = ""
            // Dim tmp_kozasyu As String = ""      '20160609 紐付取得用に追加 -add
            // Dim tmp_kozabango As String = ""
            // Dim tmp_jisyano As String = ""      '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
            // Dim tmp_jisyakozano As String = ""  '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
            // '2016.04.26 メインの方へも反映させる修正 -add end

            // '移行値取得
            // For cntjj = 1 To columncnt

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "振替情報No"                          '20160929 汎用→既存コピー処理改善対応 振替情報No. → 振替情報No -chg
            // .Vari_Fkae_no = fldvalue
            // .Vari_Fkae_fb_orgfmtno = fldvalue       '20160609 紐付設定値取得に伴う修正 -add
            // Case "振替情報名称"
            // .Vari_Fkae_name = fldvalue
            // Case "振替情報カナ名称"
            // .Vari_Fkae_kana = fldvalue
            // '20160519 EXEUpdateに伴う修正 口座振替情報 -del sta
            // 'Case "サービスタイプ　1:オリコ　2:ジャックス　3:アプラス"
            // '    .Vari_Servicetype = fldvalue
            // '20160519 EXEUpdateに伴う修正 口座振替情報 -del end
            // Case "振込依頼人"
            // .Vari_Fkae_fb_fkomiraino = fldvalue
            // Case "振込依頼人カナ名"
            // .Vari_Fkae_fb_fkomiraikana = fldvalue
            // Case "加盟店No 依頼人番号と共用"          '20160929 汎用→既存コピー処理改善対応 加盟店No.　依頼人番号と共用 → 加盟店No 依頼人番号と共用 -chg
            // .Vari_Kamei_no = fldvalue
            // Case "ジェイリース　入金区分　1:変更なし・・・・"
            // .Vari_Jlease_nyukinkbn = fldvalue
            // '2016.04.26 メインの方へも反映させる修正 -chg sta
            // 'Case "自社支店No"
            // '    .Vari_Jisya_no = fldvalue
            // 'Case "振替先口座No."
            // '    .Vari_Fkae_fb_nkinukekozano = fldvalue
            // Case "金融機関No"
            // tmp_kinyuno = fldvalue
            // Case "金融機関支店No"
            // tmp_kinyutenno = fldvalue
            // Case "口座種別" '20160609 紐付取得用に追加 -add
            // tmp_kozasyu = fldvalue
            // Case "口座番号"
            // tmp_kozabango = fldvalue
            // '2016.04.26 メインの方へも反映させる修正 -chg end
            // Case "引落日"
            // .Vari_Hikiotosibi = fldvalue
            // Case "手数料　請求額"
            // .Vari_Tesu_gak = fldvalue
            // Case "備考"
            // .Vari_Biko = fldvalue
            // '20160609 紐付設定値取得に伴う修正 -del sta
            // 'Case "オリジナルフォーマットNo"
            // '    .Vari_Fkae_fb_orgfmtno = fldvalue
            // '20160609 紐付設定値取得に伴う修正 -del end
            // Case "データ並び替え　0:物件番号と部屋番号　1:契約者番号"
            // .Vari_Datasort = fldvalue
            // Case "契約者Noを契約者番号に出力する  0:出力しない　1:出力する"    '20160929 汎用→既存コピー処理改善対応 契約者No.を契約者番号に出力する  0:出力しない　1:出力する → 契約者Noを契約者番号に出力する  0:出力しない　1:出力する -chg
            // .Vari_Keiyakusyano_syuturyoku = fldvalue
            // Case "改行コード(CRLF)出力　0:なし　1:あり"
            // .Vari_Crlf = fldvalue
            // Case "「ｦ」→「ｵ」変換　0:しない　1:する"
            // .Vari_Wo_mojihenkan = fldvalue
            // Case "送信ファイルパス名"
            // .Vari_File_sosin = fldvalue
            // Case "受信ファイルパス名"
            // .Vari_File_jyusin = fldvalue
            // '20160519 EXEUpdateに伴う修正 口座振替情報 -del sta
            // 'Case "2枚ＦＤ作成手順を使用する　0:しない　1:する"
            // '    .Vari_Fdsakusei = fldvalue
            // '20160519 EXEUpdateに伴う修正 口座振替情報 -del end
            // Case "請求のまとめ方　0:請求先単位　1:契約単位"
            // .Vari_Seikyu_tani = fldvalue
            // Case "滞納分請求　0:しない　1:する"
            // .Vari_Seikyu_taino = fldvalue
            // Case "滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける"
            // .Vari_Seikyu_tukitani = fldvalue
            // Case "手数料の請求入金管理をする　0:しない　1:する"
            // .Vari_Tesu_nyukinkanri = fldvalue
            // Case "未入金の手数料は次回請求に加える　0:加えない　1:加える"
            // .Vari_Tesu_kurikosi = fldvalue
            // Case "照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合"
            // .Vari_Syogorule = fldvalue
            // Case "ゆうちょ銀行金融機関番号指定　不使用の場合はNULL"
            // .Vari_Yutyobango = fldvalue
            // '20160519 EXEUpdateに伴う修正 口座振替情報 -add sta
            // Case "マルチヘッダー形式フラグ"
            // .Vari_Yutyobango = fldvalue
            // Case "ゆうちょ銀行付加コード"
            // .Vari_Yutyobango = fldvalue
            // Case "振替でのゆうちょ銀行コードに振込用変換を利用するフラグ"
            // .Vari_Yutyobango = fldvalue
            // Case "振替入金処理時に受信ファイルを読み込まないフラグ"
            // .Vari_Yutyobango = fldvalue
            // '20160519 EXEUpdateに伴う修正 口座振替情報 -add end
            // '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
            // Case "自社支店No"
            // tmp_jisyano = fldvalue
            // Case "振替先口座No"                 '20160929 汎用→既存コピー処理改善対応 振替先口座No. → 振替先口座No -chg
            // tmp_jisyakozano = fldvalue
            // '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end
            // End Select

            // Next
            // '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg sta
            // ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
            // ' ''20160609 紐付取得用に追加 -chg sta
            // '' ''2016.04.26 メインの方へも反映させる修正 -add sta
            // '' ''自社No/自社口座No取得のため口座情報を格納する
            // ' ''.Vari_Jisya_no = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
            // ' ''.Vari_Fkae_fb_nkinukekozano = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
            // '' ''2016.04.26 メインの方へも反映させる修正 -add end
            // ''Dim tmp_jisyainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
            // ''.Vari_Jisya_no = tmp_jisyainfo
            // ''.Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
            // ' ''20160609 紐付取得用に追加 -chg end
            // 'Dim tmp_kozamoto As String = "口座振替マスタ"
            // 'Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Fkae_no
            // '.Vari_Jisya_no = tmp_jisyainfo
            // '.Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
            // ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
            // Select Case CNVNO
            // Case ConvertTypes._既存ユーザ用
            // Dim tmp_kozamoto As String = "口座振替マスタ"
            // Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Fkae_no
            // .Vari_Jisya_no = tmp_jisyainfo
            // .Vari_Fkae_fb_nkinukekozano = tmp_jisyainfo
            // Case ConvertTypes._汎用
            // .Vari_Fkae_fb_orgfmtno = ""
            // '20160928 自社口座取得方法の修正 -chg sta
            // '.Vari_Jisya_no = tmp_jisyano
            // '.Vari_Fkae_fb_nkinukekozano = tmp_jisyakozano
            // .Vari_Jisya_no = tmp_jisyano & "-" & tmp_jisyakozano
            // .Vari_Fkae_fb_nkinukekozano = tmp_jisyano & "-" & tmp_jisyakozano
            // '20160928 自社口座取得方法の修正 -chg end
            // End Select
            // '20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg end
            // '固定値
            // .Vari_Useflg = 1
            // .Vari_History = DefHistory

            // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            // 'データチェック
            // Dim skipflg As Boolean = False
            // Dim hash_cvitem As New SafeDictionary<string, string>
            // Dim hash_log As New SafeDictionary<string, string>
            // skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

            // '20160928 自社口座取得方法の修正 -add sta
            // tmp_hash("jisya_no") = tmp_jisyano
            // tmp_hash("fkae_fb_nkinukekozano") = tmp_jisyakozano
            // '20160928 自社口座取得方法の修正 -add end

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


                string tmp_sql = " SELECT jisya_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);


            }

        }

    }

    #endregion

    #region 入出金取得情報

    public class M_fb_nskinsetting_Repository
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

                var model_cvitem = new Model.M_fb_nskinsetting_Model();       // 移行値格納用モデル初期化
                int keycol = 1;                                       // メインキー列

                // ************************
                // 作業準備
                // ************************

                // 紐付けデータ取得
                // 20160609 紐付設定値取得に伴う修正 -chg sta
                // 'FBフォーマット紐付け情報を取得
                // Call SetRelItemToObject.Set_RelData_FBFmt()
                SetRelItemToObject.Set_RelData_FBInfo();
                // 20160609 紐付設定値取得に伴う修正 -chg end

                // Excelファイル初期設定
                // 20160609 紐付設定値取得に伴う修正 -chg sta
                // 中間ファイルから読み込むように修正
                // '入出金取得情報は紐付ファイルから読込む
                // Dim readfiledirpath As String = RelationDirPath             '紐付けファイル格納先
                // Dim readfilename As String = MIDFILE_RELNAME                '紐付けファイル名
                // Dim readsheetname As String = "入出金取得情報マスタ"        '紐付けシート名
                // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, readfiledirpath, readfilename, readsheetname)
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);
                // 20160609 紐付設定値取得に伴う修正 -chg end

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // テーブル名/フィールド名セット
                string tblname = "m_fb_nskinsetting";
                // 20160519 EXEUpdateに伴う修正 入出金取得情報 -del(crlfを削除して成形)
                string fldnamegrp = "ns_no,ns_name,ns_kana,jisya_no,orgfmt_no," + "file_jyusin,data_syubetu,biko,history,useflg," + "yatin_kozano,file_sosin,yokinbunkatusetting,duplicate_check";


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
                    var datarowvalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[cntii + startrow - 1, 1], wsheet.Cells[cntii + startrow - 1, columncnt]]).Value;

                    // キー値取得
                    string fldvalue_key = EtcMethod.Get_ShapeStr(datarowvalue(startrow - 1, keycol));

                    // ログ出力用
                    string str_logkey = fldname_key + " = " + fldvalue_key;

                    // 作業用変数
                    string tmp_fmtno = "";  // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add
                    string tmp_jisyano = "";  // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add

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
                            case "入出金取得No":
                                {
                                    model_cvitem.Vari_Ns_no = fldvalue;
                                    model_cvitem.Vari_Orgfmt_no = fldvalue;   // 20160609 紐付設定値取得に伴う修正 -add
                                    model_cvitem.Vari_Jisya_no = fldvalue;       // 20160609 紐付設定値取得に伴う修正 -add
                                    break;
                                }
                            case "設定名称":
                                {
                                    model_cvitem.Vari_Ns_name = fldvalue;
                                    break;
                                }
                            case "設定カナ名称":
                                {
                                    model_cvitem.Vari_Ns_kana = fldvalue;
                                    break;
                                }
                            // 20160609 紐付設定値取得に伴う修正 -del sta
                            // Case "フォーマット定義名"
                            // .Vari_Orgfmt_no = fldvalue
                            // 20160609 紐付設定値取得に伴う修正 -del end
                            case "受信ファイル名（フルパス）":
                                {
                                    model_cvitem.Vari_File_jyusin = fldvalue;
                                    break;
                                }
                            case "取込データ種別":
                                {
                                    model_cvitem.Vari_Data_syubetu = fldvalue;
                                    break;
                                }
                            // 20160519 EXEUpdateに伴う修正 入出金取得情報 -del sta
                            // Case "改行コード有無"
                            // .Vari_Crlf = fldvalues
                            // 20160519 EXEUpdateに伴う修正 入出金取得情報 -del end
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko = fldvalue;
                                    break;
                                }
                            case "家賃口座No":
                                {
                                    model_cvitem.Vari_Yatin_kozano = fldvalue;
                                    break;
                                }
                            case "送信ファイル名（フルパス）":
                                {
                                    model_cvitem.Vari_File_sosin = fldvalue;
                                    break;
                                }
                            case "預金分割設定":
                                {
                                    model_cvitem.Vari_Yokinbunkatusetting = fldvalue;
                                    break;
                                }
                            case "重複チェック":
                                {
                                    model_cvitem.Vari_Duplicate_check = fldvalue;
                                    break;
                                }
                            // 20160609 紐付設定値取得に伴う修正 -del sta
                            // Case "自社No"
                            // .Vari_Jisya_no = fldvalue
                            // 20160609 紐付設定値取得に伴う修正 -del end
                            // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
                            case "自社支店No":
                                {
                                    tmp_jisyano = fldvalue;
                                    break;
                                }
                            case "フォーマット定義名":
                                {
                                    tmp_fmtno = fldvalue;
                                    break;
                                }
                                // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end
                        }

                    }

                    // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add sta
                    if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                    {
                        model_cvitem.Vari_Orgfmt_no = tmp_fmtno;
                        model_cvitem.Vari_Jisya_no = tmp_jisyano;
                    }
                    // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -add end

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;
                    model_cvitem.Vari_Useflg = 1.ToString();

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

                string tmp_sql = " SELECT jisya_no FROM " + tblname;
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

                string tmp_sql = " SELECT ns_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 家賃入金口座情報

    public class M_yatinkoza_Repository
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

                var model_cvitem = new Model.M_yatinkoza_Model();             // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************

                // 20160609 紐付取得用に追加 -add sta
                // 紐付けデータ取得
                SetRelItemToObject.Set_RelData_Jisyakoza();
                // 20160609 紐付取得用に追加 -add end

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
                string tblname = "m_yatinkoza";
                string fldnamegrp = "yatin_kozano,yatin_kozaname,yatin_kozanamesjis,yatin_kozakana,sqsaki_kbn," + "jisya_no,jisya_kozano,ow_no,ow_kozano,spc_useflg," + "spc_kozasiteikbn,spc_syumokukbn,spc_ninibango,spc_kanyubango,spc_kobetuflg," + "spc_kobetukbn,spc_kobetuareano,spc_kobetutikuno,spc_kobetutel,spc_ansyobango," + "spc_servicecode,history,useflg,yatin_kozabiko,headertemplate," + "trallertemplate";





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

                // 20160624 家賃入金口座情報 重複ログ出力制御修正 -add
                int chg_totalcnt = rowcnt;    // 重複データ件数調整用(カウントしないため)

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
                    string tmp_kinyuno = "";
                    string tmp_kinyutenno = "";
                    string tmp_kozasyu = "";
                    string tmp_kozabango = "";
                    string tmp_jisyano = "";      // 20160928 自社口座取得方法の修正 -add
                    string tmp_jisyakozano = "";  // 20160928 自社口座取得方法の修正 -add
                    string tmp_owno = "";      // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                    string tmp_owkozano = "";  // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add

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
                            case "家賃入金受付口座No":
                                {
                                    model_cvitem.Vari_Yatin_kozano = fldvalue;
                                    break;
                                }
                            case "受取人名":
                                {
                                    model_cvitem.Vari_Yatin_kozaname = fldvalue;
                                    break;
                                }
                            // 20160609 紐付取得用に追加 -add sta
                            case "金融機関No":
                                {
                                    tmp_kinyuno = fldvalue;
                                    break;
                                }
                            case "金融機関支店No":
                                {
                                    tmp_kinyutenno = fldvalue;
                                    break;
                                }
                            case "口座種別": // 20160609 紐付取得用に追加 -add
                                {
                                    tmp_kozasyu = fldvalue;
                                    break;
                                }
                            case "口座番号":
                                {
                                    tmp_kozabango = fldvalue;
                                    break;
                                }
                            // 20160609 紐付取得用に追加 -add end
                            case "受取人名Unicode":
                                {
                                    model_cvitem.Vari_Yatin_kozanamesjis = fldvalue;
                                    break;
                                }
                            case "受取人カナ":
                                {
                                    model_cvitem.Vari_Yatin_kozakana = fldvalue;
                                    break;
                                }
                            case "請求先区分":
                                {
                                    model_cvitem.Vari_Sqsaki_kbn = fldvalue;
                                    break;
                                }
                            case "自社支店No":
                                {
                                    // 20160928 自社口座取得方法の修正 -chg sta
                                    // .Vari_Jisya_no = fldvalue
                                    tmp_jisyano = fldvalue;
                                    break;
                                }
                            // 20160928 自社口座取得方法の修正 -chg end
                            case "自社口座No":
                                {
                                    // 20160928 自社口座取得方法の修正 -chg sta
                                    // .Vari_Jisya_kozano = fldvalue
                                    tmp_jisyakozano = fldvalue;
                                    break;
                                }
                            // 20160928 自社口座取得方法の修正 -chg end
                            case "家主No":
                                {
                                    // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg sta
                                    // .Vari_Ow_no = fldvalue
                                    tmp_owno = fldvalue;
                                    break;
                                }
                            // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg end
                            case "家主口座No":
                                {
                                    // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg sta
                                    // .Vari_Ow_kozano = fldvalue
                                    tmp_owkozano = fldvalue;
                                    break;
                                }
                            // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg end
                            case "ANSER-SPC有無":
                                {
                                    model_cvitem.Vari_Spc_useflg = fldvalue;
                                    break;
                                }
                            case "ANSER-SPC口座指定方式":
                                {
                                    model_cvitem.Vari_Spc_kozasiteikbn = fldvalue;
                                    break;
                                }
                            case "ANSER-SPC種目付加方法":
                                {
                                    model_cvitem.Vari_Spc_syumokukbn = fldvalue;
                                    break;
                                }
                            case "ANSER-SPC任意番号":
                                {
                                    model_cvitem.Vari_Spc_ninibango = fldvalue;
                                    break;
                                }
                            case "ANSER-SPC加入者番号":
                                {
                                    model_cvitem.Vari_Spc_kanyubango = fldvalue;
                                    break;
                                }
                            case "ANSER-SPC接続先個別設定フラグ":
                                {
                                    model_cvitem.Vari_Spc_kobetuflg = fldvalue;
                                    break;
                                }
                            case "ANSER-SPC接続方法":
                                {
                                    model_cvitem.Vari_Spc_kobetukbn = fldvalue;
                                    break;
                                }
                            case "ANSER-SPCエリアNo":
                                {
                                    model_cvitem.Vari_Spc_kobetuareano = fldvalue;
                                    break;
                                }
                            case "ANSER-SPC地区No":
                                {
                                    model_cvitem.Vari_Spc_kobetutikuno = fldvalue;
                                    break;
                                }
                            case "ANSER-SPCTEL":
                                {
                                    model_cvitem.Vari_Spc_kobetutel = fldvalue;
                                    break;
                                }
                            case "ANSER-SPC照会用暗証番号":
                                {
                                    model_cvitem.Vari_Spc_ansyobango = fldvalue;
                                    break;
                                }
                            case "ANSER-SPCサービスコード":
                                {
                                    model_cvitem.Vari_Spc_servicecode = fldvalue;
                                    break;
                                }
                            case "家賃入金口座備考":
                                {
                                    model_cvitem.Vari_Yatin_kozabiko = fldvalue;
                                    break;
                                }
                            case "ヘッダーテンプレート":
                                {
                                    model_cvitem.Vari_Headertemplate = fldvalue;
                                    break;
                                }
                            case "トレーラーテンプレート":
                                {
                                    model_cvitem.Vari_Trallertemplate = fldvalue;
                                    break;
                                }
                        }

                    }
                    // 20160928 自社口座取得方法の修正 -chg sta
                    // '20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                    // ''20160609 紐付取得用に追加 -add sta
                    // 'Dim tmp_jisyainfo As String = ""
                    // 'If .Vari_Ow_no = "" OrElse .Vari_Ow_kozano = "" Then
                    // '    tmp_jisyainfo = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
                    // 'End If
                    // '.Vari_Jisya_no = tmp_jisyainfo
                    // '.Vari_Jisya_kozano = tmp_jisyainfo
                    // ''20160609 紐付取得用に追加 -add end
                    // Dim tmp_kozamoto As String = "振込先口座マスタ"
                    // Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Yatin_kozano
                    // .Vari_Jisya_no = tmp_jisyainfo
                    // .Vari_Jisya_kozano = tmp_jisyainfo
                    // '20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                    switch (CommonModule.CNVNO)
                    {
                        case (int)CommonModule.ConvertTypes._汎用:
                            {
                                model_cvitem.Vari_Jisya_no = tmp_jisyano + "-" + tmp_jisyakozano;
                                model_cvitem.Vari_Jisya_kozano = tmp_jisyano + "-" + tmp_jisyakozano;
                                model_cvitem.Vari_Ow_no = tmp_owno + "-" + tmp_owkozano;         // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                                model_cvitem.Vari_Ow_kozano = tmp_owno + "-" + tmp_owkozano;     // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                                break;
                            }
                    }
                    // 20160928 自社口座取得方法の修正 -chg end
                    // 固定値
                    model_cvitem.Vari_Useflg = 1.ToString();
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 20160928 自社口座取得方法の修正 -add sta
                    tmp_hash["jisya_no"] = tmp_jisyano;
                    tmp_hash["jisya_kozano"] = tmp_jisyakozano;
                    // 20160928 自社口座取得方法の修正 -add end

                    // 20160930 家賃入金口座移行制御処理の追加 -add sta
                    if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                    {
                        skipflg = !Chk_ExistJisyaOwkoza(tblname, ref hash_cvitem, ref hash_log);
                    }
                    // 20160930 家賃入金口座移行制御処理の追加 -add end

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

                    // 20160624 家賃入金口座情報 重複ログ出力制御修正 -add sta
                    else if (hash_log.Count == 1)
                    {
                        bool logclearflg = false;
                        foreach (var erritem in hash_log)
                        {
                            string tmp_str = Conversions.ToString(erritem.Value);
                            string[] tmp_errnaiyo = tmp_str.Split('-');
                            if ((tmp_errnaiyo[0] ?? "") == CommonModule.LOG_NAIYO_ERR_OVERLAP)
                            {
                                chg_totalcnt = chg_totalcnt - 1;
                                logclearflg = true;
                            }
                        }
                        if (logclearflg)
                        {
                            hash_log.Clear();
                        }
                        // 20160624 家賃入金口座情報 重複ログ出力制御修正 -add end
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
                // 20160624 家賃入金口座情報 重複ログ出力制御修正 -chg sta
                // midrowcnt = rowcnt
                midrowcnt = chg_totalcnt;
                // 20160624 家賃入金口座情報 重複ログ出力制御修正 -chg end

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

            // Dim model_cvitem As New Njc.Model.M_yatinkoza_Model             '移行値格納用モデル初期化
            // Dim keycol As Integer = 1                                       'キー列

            // '************************
            // '作業準備
            // '************************

            // '20160609 紐付取得用に追加 -add sta
            // '紐付けデータ取得
            // Call SetRelItemToObject.Set_RelData_Jisyakoza()
            // '20160609 紐付取得用に追加 -add end

            // 'Excelファイル初期設定
            // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

            // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
            // If Not rtn Then
            // Return rtn
            // End If

            // 'テーブル名/フィールド名セット
            // Dim tblname As String = "m_yatinkoza"
            // Dim fldnamegrp As String = "yatin_kozano,yatin_kozaname,yatin_kozanamesjis,yatin_kozakana,sqsaki_kbn," & _
            // "jisya_no,jisya_kozano,ow_no,ow_kozano,spc_useflg," & _
            // "spc_kozasiteikbn,spc_syumokukbn,spc_ninibango,spc_kanyubango,spc_kobetuflg," & _
            // "spc_kobetukbn,spc_kobetuareano,spc_kobetutikuno,spc_kobetutel,spc_ansyobango," & _
            // "spc_servicecode,history,useflg,yatin_kozabiko,headertemplate," & _
            // "trallertemplate"

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

            // '20160624 家賃入金口座情報 重複ログ出力制御修正 -add
            // Dim chg_totalcnt As Integer = rowcnt    '重複データ件数調整用(カウントしないため)

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

            // '20160609 紐付取得用に追加 -add sta
            // '作業用変数
            // Dim tmp_kinyuno As String = ""
            // Dim tmp_kinyutenno As String = ""
            // Dim tmp_kozasyu As String = ""
            // Dim tmp_kozabango As String = ""
            // Dim tmp_jisyano As String = ""      '20160928 自社口座取得方法の修正 -add
            // Dim tmp_jisyakozano As String = ""  '20160928 自社口座取得方法の修正 -add
            // Dim tmp_owno As String = ""      '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            // Dim tmp_owkozano As String = ""  '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            // '20160609 紐付取得用に追加 -add end

            // '移行値取得
            // For cntjj = 1 To columncnt

            // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
            // Dim fldvalue As String = ""
            // If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // End If

            // Select Case fldname
            // Case "家賃入金受付口座No"
            // .Vari_Yatin_kozano = fldvalue
            // Case "受取人名"
            // .Vari_Yatin_kozaname = fldvalue
            // '20160609 紐付取得用に追加 -add sta
            // Case "金融機関No"
            // tmp_kinyuno = fldvalue
            // Case "金融機関支店No"
            // tmp_kinyutenno = fldvalue
            // Case "口座種別" '20160609 紐付取得用に追加 -add
            // tmp_kozasyu = fldvalue
            // Case "口座番号"
            // tmp_kozabango = fldvalue
            // '20160609 紐付取得用に追加 -add end
            // Case "受取人名Unicode"
            // .Vari_Yatin_kozanamesjis = fldvalue
            // Case "受取人カナ"
            // .Vari_Yatin_kozakana = fldvalue
            // Case "請求先区分"
            // .Vari_Sqsaki_kbn = fldvalue
            // Case "自社支店No"
            // '20160928 自社口座取得方法の修正 -chg sta
            // '.Vari_Jisya_no = fldvalue
            // tmp_jisyano = fldvalue
            // '20160928 自社口座取得方法の修正 -chg end
            // Case "自社口座No"
            // '20160928 自社口座取得方法の修正 -chg sta
            // '.Vari_Jisya_kozano = fldvalue
            // tmp_jisyakozano = fldvalue
            // '20160928 自社口座取得方法の修正 -chg end
            // Case "家主No"
            // '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg sta
            // '.Vari_Ow_no = fldvalue
            // tmp_owno = fldvalue
            // '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg end
            // Case "家主口座No"
            // '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg sta
            // '.Vari_Ow_kozano = fldvalue
            // tmp_owkozano = fldvalue
            // '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -chg end
            // Case "ANSER-SPC有無"
            // .Vari_Spc_useflg = fldvalue
            // Case "ANSER-SPC口座指定方式"
            // .Vari_Spc_kozasiteikbn = fldvalue
            // Case "ANSER-SPC種目付加方法"
            // .Vari_Spc_syumokukbn = fldvalue
            // Case "ANSER-SPC任意番号"
            // .Vari_Spc_ninibango = fldvalue
            // Case "ANSER-SPC加入者番号"
            // .Vari_Spc_kanyubango = fldvalue
            // Case "ANSER-SPC接続先個別設定フラグ"
            // .Vari_Spc_kobetuflg = fldvalue
            // Case "ANSER-SPC接続方法"
            // .Vari_Spc_kobetukbn = fldvalue
            // Case "ANSER-SPCエリアNo"
            // .Vari_Spc_kobetuareano = fldvalue
            // Case "ANSER-SPC地区No"
            // .Vari_Spc_kobetutikuno = fldvalue
            // Case "ANSER-SPCTEL"
            // .Vari_Spc_kobetutel = fldvalue
            // Case "ANSER-SPC照会用暗証番号"
            // .Vari_Spc_ansyobango = fldvalue
            // Case "ANSER-SPCサービスコード"
            // .Vari_Spc_servicecode = fldvalue
            // Case "家賃入金口座備考"
            // .Vari_Yatin_kozabiko = fldvalue
            // Case "ヘッダーテンプレート"
            // .Vari_Headertemplate = fldvalue
            // Case "トレーラーテンプレート"
            // .Vari_Trallertemplate = fldvalue
            // End Select

            // Next
            // '20160928 自社口座取得方法の修正 -chg sta
            // ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
            // ' ''20160609 紐付取得用に追加 -add sta
            // ''Dim tmp_jisyainfo As String = ""
            // ''If .Vari_Ow_no = "" OrElse .Vari_Ow_kozano = "" Then
            // ''    tmp_jisyainfo = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
            // ''End If
            // ''.Vari_Jisya_no = tmp_jisyainfo
            // ''.Vari_Jisya_kozano = tmp_jisyainfo
            // ' ''20160609 紐付取得用に追加 -add end
            // 'Dim tmp_kozamoto As String = "振込先口座マスタ"
            // 'Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Yatin_kozano
            // '.Vari_Jisya_no = tmp_jisyainfo
            // '.Vari_Jisya_kozano = tmp_jisyainfo
            // ''20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
            // Select Case CNVNO
            // Case ConvertTypes._既存ユーザ用
            // Dim tmp_kozamoto As String = "振込先口座マスタ"
            // Dim tmp_jisyainfo As String = tmp_kozamoto & "-" & .Vari_Yatin_kozano
            // .Vari_Jisya_no = tmp_jisyainfo
            // .Vari_Jisya_kozano = tmp_jisyainfo
            // .Vari_Ow_no = tmp_owno          '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            // .Vari_Ow_kozano = tmp_owkozano  '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            // Case ConvertTypes._汎用
            // .Vari_Jisya_no = tmp_jisyano & "-" & tmp_jisyakozano
            // .Vari_Jisya_kozano = tmp_jisyano & "-" & tmp_jisyakozano
            // .Vari_Ow_no = tmp_owno & "-" & tmp_owkozano         '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            // .Vari_Ow_kozano = tmp_owno & "-" & tmp_owkozano     '20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
            // End Select
            // '20160928 自社口座取得方法の修正 -chg end
            // '固定値
            // .Vari_Useflg = 1
            // .Vari_History = DefHistory

            // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
            // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            // 'データチェック
            // Dim skipflg As Boolean = False
            // Dim hash_cvitem As New SafeDictionary<string, string>
            // Dim hash_log As New SafeDictionary<string, string>
            // skipflg = Not (DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, hash_cvitem, tmp_condcnt, hash_log))

            // '20160928 自社口座取得方法の修正 -add sta
            // tmp_hash("jisya_no") = tmp_jisyano
            // tmp_hash("jisya_kozano") = tmp_jisyakozano
            // '20160928 自社口座取得方法の修正 -add end

            // '20160930 家賃入金口座移行制御処理の追加 -add sta
            // If CNVNO = ConvertTypes._汎用 Then
            // skipflg = Not (Me.Chk_ExistJisyaOwkoza(tblname, hash_cvitem, hash_log))
            // End If
            // '20160930 家賃入金口座移行制御処理の追加 -add end

            // '書込処理
            // If Not skipflg Then

            // '挿入処理
            // Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)

            // '挿入処理が正常終了したレコードのキーを重複チェック用に格納
            // If normalflg Then
            // list_chkduplicate.Add(fldvalue_key)
            // tmp_cvcnt = tmp_cvcnt + 1
            // End If

            // '20160624 家賃入金口座情報 重複ログ出力制御修正 -add sta
            // Else
            // If hash_log.Count = 1 Then
            // Dim logclearflg As Boolean = False
            // For Each erritem In hash_log
            // Dim tmp_str As String = erritem.Value
            // Dim tmp_errnaiyo() As String = tmp_str.Split("-")
            // If tmp_errnaiyo(0) = LOG_NAIYO_ERR_OVERLAP Then
            // chg_totalcnt = chg_totalcnt - 1
            // logclearflg = True
            // End If
            // Next
            // If logclearflg Then
            // hash_log.Clear()
            // End If
            // End If
            // '20160624 家賃入金口座情報 重複ログ出力制御修正 -add end
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
            // '20160624 家賃入金口座情報 重複ログ出力制御修正 -chg sta
            // 'midrowcnt = rowcnt
            // midrowcnt = chg_totalcnt
            // '20160624 家賃入金口座情報 重複ログ出力制御修正 -chg end

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

                string tmp_sql = " SELECT yatin_kozano FROM " + tblname;
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

                string tmp_sql = " SELECT yatin_kozano FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 自社口座、家主口座有無チェック '20160930 家賃入金口座移行制御処理の追加
            /// 無い場合は移行自体を行わない
            /// </summary>
            /// <param name="tblname"></param>
            /// <param name="hash_cvitem"></param>
            /// <param name="hash_log"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Chk_ExistJisyaOwkoza(string tblname, ref SafeDictionary<string, string> hash_cvitem, ref SafeDictionary<string, string> hash_log)
            {

                bool rtn = true;

                string tmp_sqsakikbn = Conversions.ToString(hash_cvitem["sqsaki_kbn"]);
                string tmp_jisyano = Conversions.ToString(hash_cvitem["jisya_no"]);
                string tmp_jisyakozano = Conversions.ToString(hash_cvitem["jisya_kozano"]);
                string tmp_owno = Conversions.ToString(hash_cvitem["ow_no"]);
                string tmp_owkozano = Conversions.ToString(hash_cvitem["ow_kozano"]);
                string tmp_errstr = "";
                string tmp_errfld = "";

                // 設定された値をチェックする
                if (!string.IsNullOrEmpty(tmp_jisyano) & !string.IsNullOrEmpty(tmp_jisyakozano) & !string.IsNullOrEmpty(tmp_owno) & !string.IsNullOrEmpty(tmp_owkozano))
                {
                    // 自社、家主共に設定されている場合は自社を優先
                    // ログ出力
                    string log_key = tblname + "-" + "jisya_kozano";
                    string tmp_hubi = "自社口座情報、家主口座情報両方にデータが設定されているため自社口座情報を優先して移行します。";
                    string errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + tmp_hubi + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                    hash_log.Clear();
                    hash_log.Add(log_key, errstr);
                    hash_cvitem["sqsaki_kbn"] = "900";
                    hash_cvitem["ow_no"] = "";
                    hash_cvitem["ow_kozano"] = "";
                    return rtn;
                }
                else if (!string.IsNullOrEmpty(tmp_jisyano) & !string.IsNullOrEmpty(tmp_jisyakozano))
                {
                    // 自社のみ設定されている場合は自社を移行
                    hash_cvitem["sqsaki_kbn"] = "900";
                    return rtn;
                }
                else if (!string.IsNullOrEmpty(tmp_owno) & !string.IsNullOrEmpty(tmp_owkozano))
                {
                    // 家主のみ設定されている場合は家主を移行
                    hash_cvitem["sqsaki_kbn"] = "200";
                    return rtn;
                }
                else
                {
                    // 上記3つ以外は全て不完全な口座情報なので移行しない
                    string log_key = tblname + "-" + "jisya_kozano";
                    string tmp_hubi = "自社口座情報、または家主口座情報が不完全であるため移行できません。";
                    string errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + tmp_hubi + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                    hash_log.Clear();
                    hash_log.Add(log_key, errstr);
                    rtn = false;
                    return rtn;
                }

                return rtn;

            }

        }

    }

    #endregion

    #region ANSERエリア情報

    public class M_spcarea_Repository
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

                var model_cvitem = new Model.M_spcarea_Model();                  // 移行値格納用モデル初期化
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
                string tblname = "m_spcarea";
                string fldnamegrp = "spc_areano,spc_areaname,biko_spcarea,history";

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
                            case "SPCエリアNo":
                                {
                                    model_cvitem.Vari_Spc_areano = fldvalue;
                                    break;
                                }
                            case "SPCエリア名称":
                                {
                                    model_cvitem.Vari_Spc_areaname = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko_spcarea = fldvalue;
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

                string tmp_sql = " SELECT spc_areano FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region ANSERアクセスポイント情報

    public class M_spcaccesspoint_Repository
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

                var model_cvitem = new Model.M_spcaccesspoint_Model();        // 移行値格納用モデル初期化
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
                string tblname = "m_spcaccesspoint";
                string fldnamegrp = "spc_tikuno,spc_areano,spc_tikuname,biko,history";

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
                            case "SPC地区No":
                                {
                                    model_cvitem.Vari_Spc_tikuno = fldvalue;
                                    break;
                                }
                            case "SPCエリアNo":
                                {
                                    model_cvitem.Vari_Spc_areano = fldvalue;
                                    break;
                                }
                            case "SPC地区名称":
                                {
                                    model_cvitem.Vari_Spc_tikuname = fldvalue;
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

                string tmp_sql = " SELECT spc_tikuno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region ANSER接続情報

    public class M_spcsetuzoku_Repository
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

                var model_cvitem = new Model.M_spcsetuzoku_Model();           // 移行値格納用モデル初期化
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
                string tblname = "m_spcsetuzoku";
                string fldnamegrp = "spc_setuzokuno,spc_name,spc_kaisensyu,spc_devtype,spc_devname," + "spc_setuzoku_hoho,spc_areano,spc_tikuno,spc_tel,spc_gaisen," + "spc_retry_kaisu,spc_retry_kankaku,spc_servicecode,spc_biko,history," + "spc_crlf";



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
                            case "接続方法No":
                                {
                                    model_cvitem.Vari_Spc_setuzokuno = fldvalue;
                                    break;
                                }
                            case "設定名":
                                {
                                    model_cvitem.Vari_Spc_name = fldvalue;
                                    break;
                                }
                            case "回線種別(1:電話回線 2:ISDN)":
                                {
                                    model_cvitem.Vari_Spc_kaisensyu = fldvalue;
                                    break;
                                }
                            case "開発タイプ":
                                {
                                    model_cvitem.Vari_Spc_devtype = fldvalue;
                                    break;
                                }
                            case "開発名":
                                {
                                    model_cvitem.Vari_Spc_devname = fldvalue;
                                    break;
                                }
                            case "接続方法(1:アクセスポイント 2:TEL)":
                                {
                                    model_cvitem.Vari_Spc_setuzoku_hoho = fldvalue;
                                    break;
                                }
                            case "エリアNo":
                                {
                                    model_cvitem.Vari_Spc_areano = fldvalue;
                                    break;
                                }
                            case "地区No":
                                {
                                    model_cvitem.Vari_Spc_tikuno = fldvalue;
                                    break;
                                }
                            case "電話番号":
                                {
                                    model_cvitem.Vari_Spc_tel = fldvalue;
                                    break;
                                }
                            case "外線No":
                                {
                                    model_cvitem.Vari_Spc_gaisen = fldvalue;
                                    break;
                                }
                            case "リトライ回数":
                                {
                                    model_cvitem.Vari_Spc_retry_kaisu = fldvalue;
                                    break;
                                }
                            case "リトライの間隔":
                                {
                                    model_cvitem.Vari_Spc_retry_kankaku = fldvalue;
                                    break;
                                }
                            case "利用するサービスコード":
                                {
                                    model_cvitem.Vari_Spc_servicecode = fldvalue;
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Spc_biko = fldvalue;
                                    break;
                                }
                            case "改行有無":
                                {
                                    model_cvitem.Vari_Spc_crlf = fldvalue;
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

                string tmp_sql = " SELECT spc_setuzokuno FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

}