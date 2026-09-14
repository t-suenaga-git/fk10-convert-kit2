using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Core.Metadata.Edm;
using System.Data.Entity.Infrastructure;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Converter10.EntityFramework;
using Converter10.Njc.Common;
using Converter10.Njc.N3Lib.Utys;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.Win32;             // 2016.02.22 V7DB接続情報取得処理追加

namespace Converter10.Njc.Frm
{

    public partial class MainFrm
    {

        #region 宣言

        // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        // 20160705 確認メモ
        // 使用する前に初期化を行っているかを再確認する
        // (CV後に再度CVをそのまま行った時に問題ない)
        // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

        // [接続情報]
        private System.Data.SqlClient.SqlConnection sqlcnnv7;     // V7接続情報
        private System.Data.SqlClient.SqlConnection sqlcnnv10;    // 10接続情報
        private Model.DefSQLConnection fstmodelv7 = new Model.DefSQLConnection();        // V7接続情報格納モデル
        private Model.DefSQLConnection fstmodelv10 = new Model.DefSQLConnection();       // 10接続情報格納モデル
        private DBConnection cnnv7 = new DBConnection();                // V7接続処理用クラス
        private DBConnection cnnv10 = new DBConnection();               // 10接続処理用クラス

        // [各タブ生成オブジェクト]
        private TabPageManager tabPageManager;         // タブページ表示用(メイン画面タブ)
        private TabPageManager tabPageManagerCvitem;   // タブページ表示用(項目選択タブ)
        private TabPageManager tabPageManagerJizen;    // タブページ表示用(事前作業タブ)
        private TabPageManager tabPageManagerJigo;     // タブページ表示用(事後作業タブ)
        private TabPageManager tabPageManagerHojyo;    // タブページ表示用(補助機能タブ)

        // [その他]
        private LogSetting logset = new LogSetting();                 // ログ出力用
        private Model.CVItemInfo model_cvitem = new Model.CVItemInfo();            // 移行項目名称格納用
        private CommonRepository obj_com = new CommonRepository();          // 共通処理用
        private SafeDictionary<string, string> hash_deltblqrywhere = new SafeDictionary<string, string>();                // テーブル名と削除対象を紐付値格納用リスト

        // [処理件数関連]
        private int pgbtotalcnt;                              // プログレスバー全件数
        private int pgbpartcnt;                               // プログレスバー個別件数
        private bool beforemidfilechkflg = false;              // 中間ファイルチェック前準備完了フラグ

        // [汎用コンバートキット用]
        private List<CheckBox> list_baseCVchkitem = new List<CheckBox>();         // 20160516 汎用のみ表示するチェックボックスのチェックをOFFにする処理を追加 -add
        private List<CheckBox> list_existCVchkitem = new List<CheckBox>();        // 20160905 汎用の場合の移行対象外項目(既存のみ移行対象)選定の修正 -add

        // [パス]
        private string dcv_exepath = System.Reflection.Assembly.GetExecutingAssembly().Location;                  // データコンバーター実行ファイル格納パス]
        private string dcv_exedir;                                        // データコンバーターディレクトリパス

        // [コンバーター対象項目選択タブ名]
        private const string DC_SELTAB_K01 = "tabPageKizon110";
        private const string DC_SELTAB_K02 = "tabPageKizon120";
        private const string DC_SELTAB_K03 = "tabPageKizon130";
        private const string DC_SELTAB_B01 = "tabPageBase110";
        private const string DC_SELTAB_B02 = "tabPageBase120";
        private const string DC_SELTAB_B03 = "tabPageBase130";
        private const string DC_SELTAB_B04 = "tabPageBase140";
        private const string DC_SELTAB_B05 = "tabPageBase150";
        private const string DC_SELTAB_B06 = "tabPageBase160";
        private const string DC_SELTAB_B07 = "tabPageBase170";
        private const string DC_SELTAB_B08 = "tabPageBase180";
        private const string DC_SELTAB_B09 = "tabPageBase190";
        private const string DC_SELTAB_B10 = "tabPageBase200";
        private const string DC_SELTAB_B11 = "tabPageBase210";                                                    // 20160720 連動情報構築 -add
        private const string DC_SELTAB_H01 = "tabPageHanyo110";
        private const string DC_SELTAB_H02 = "tabPageHanyo120";
        private const string DC_SELTAB_H03 = "tabPageHanyo130";
        // 20160719 V7オプション選択による表示設定処理追加_不足分対応 -add sta
        // 画面起動時にチェックボックスイベントに入るため起動時は入らないようにするフラグ
        // 画面起動イベント(FrmMain_Load)の最後にFalseに変更して保持しておく
        private bool frmloadflg_kysq = true;                   // 画面起動時判別用フラグ_契約請求関連チェックボックス制御用(True:画面起動時 False:左以外)
        private bool frmloadflg_clsz = true;                   // 画面起動時判別用フラグ_クレーム修繕関連チェックボックス制御用(True:画面起動時 False:左以外)
        // 20160719 V7オプション選択による表示設定処理追加_不足分対応 -add end
        // 20160720 紐付設定ファイル出力機能の追加 -add
        private string relitemstr = "";       // 紐付呼出時の項目格納用
        // 20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add                                       
        private int relexeerr = 0;       // 紐付ツール終了時の終了コード格納用(0:正常終了 1:キャンセル 99:異常終了)

        public MainFrm()
        {
            dcv_exedir = Path.GetDirectoryName(dcv_exepath);
            InitializeComponent();
        }

        #endregion

        #region ※汎用コンバート用処理

        /// <summary>
        /// 汎用/既存用中間ファイルの照合と中間ファイル作成 '20160812 汎用コンバート対応
        /// </summary>
        /// <param name="list_cv"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_BaseMidFileToExistingMidFile(List<string> list_cv, ref string errstr)
        {

            bool rtn = true;
            int tmp_cnt = 0;
            string tmptmpstr = "";                                      // 20161009 ログ出力処理追加 -add
            var tmptmpcnt = default(long);                                             // 20161009 ログ出力処理追加 -add
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
            var list_basecvtotal = new List<string>();

            // ----- 全体用プログレスバー初期化 -----
            var obj_pgb = new ProgressBarManager();
            int pgbcnt = 0;
            int pgbtotalcnt = 7;                                    // 7工程で進捗表示を行う                            


            // ---------------------------------------------------------------
            // 汎用用中間ファイルと既存用中間ファイルの照合用仮テーブル作成
            // ---------------------------------------------------------------
            // 仮テーブル初期化
            string midfileinfo_dropqry = DBQuery.Qry_DropInfo(MidFileInfoModule.MIDFILEINFO_DBNAME, true);
            DBExec.Exec_NonQuery(sqlcnnv10, midfileinfo_dropqry, ref tmp_cnt);

            // 仮テーブル新規作成
            string midfileinfo_createqry = MidFileInfoModule.Get_MidFileInfo_CreateQry();
            DBExec.Exec_NonQuery(sqlcnnv10, midfileinfo_createqry, ref tmp_cnt);

            // 挿入
            bool normalflg = true;
            string midfileinfo_insertqry = MidFileInfoModule.Get_MidFileInfo_InsertQry();
            if (!string.IsNullOrEmpty(midfileinfo_insertqry))
            {
                normalflg = DBExec.Exec_NonQuery(sqlcnnv10, midfileinfo_insertqry, ref tmp_cnt);
            }
            if (normalflg == false)
            {
                errstr = CommonModule.MSG_ERR_MAKE_CVDB;  // 20160913_2 エラー時の処理を追加する修正 -add
                rtn = false;
                return rtn;
            }


            // 20160905 中間ファイルコピー処理改善 -chg sta
            // '-----------------------------------------------------------------------
            // '作成した照合用仮テーブルからデータを取得してオブジェクトへ格納
            // '-----------------------------------------------------------------------
            // Call Me.Set_MidInfoToObj()
            // -----------------------------------------------------------------------
            // 作成した照合用仮テーブルからヘッダーに関する情報を取得してオブジェクトへ格納
            // -----------------------------------------------------------------------
            Set_MidHeaderInfoToObj();
            // 20160905 中間ファイルコピー処理改善 -chg end

            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
            // '-----------------------------------------------------------------------
            // ' 既存中間ファイル初期化 '20161007 既存中間ファイル初期化処理の追加 -add
            // '-----------------------------------------------------------------------
            // If CancelFlg = False And MidChkCancelFlg = False Then
            // normalflg = Me.Init_ExistMidFile(errstr, list_cv)
            // If normalflg = False Then
            // rtn = normalflg
            // Return rtn
            // End If
            // End If
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

            // -----------------------------------------------------------------------
            // 仮テーブルから紐付情報を取得して既存用中間ファイルを作成する
            // →処理が遅いため要速度改善
            // -----------------------------------------------------------------------
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
            // -----------------------------------------------------------------------
            // 中間仮テーブル作成
            // -----------------------------------------------------------------------
            // 移行する全項目をオブジェクトへ格納
            list_basecvtotal = Get_BaseCvItem();

            // 仮テーブル作成処理
            foreach (var basecvitem in list_basecvtotal)
            {

                string[] tmp_str = basecvitem.Split('-');
                string tmptblname = CommonModule.PRE_TBL_NAME + tmp_str[1];

                // 初期化
                string tmp_sql_drop = DBQuery.Qry_DropInfo(tmptblname, true);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmp_cnt);

                // 作成
                string tmp_replacetblname = "replacetmptblname";
                string tmp_sql_create = MidTableModule.MakeMidTable(basecvitem.Replace("-", CommonModule.STR_SPLIT_1), tmp_replacetblname);
                tmp_sql_create = tmp_sql_create.Replace(tmp_replacetblname, tmptblname);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmp_cnt);

            }
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

            // **************************************************
            // 通常コピー
            // **************************************************
            if (CommonModule.CancelFlg == false & CommonModule.MidChkCancelFlg == false)               // 20160927 キャンセル処理を追加 CancelFlg = False の条件追加 -add   '20160929 データチェックのキャンセル処理を追加 And MidChkCancelFlg = False の条件も追加 -chg
            {
                lblTotalSituation.Text = CommonModule.SITUATION_MID_READ + " (1/4)";       // ----- 全体用プログレスバーラベル表示 -----
                lblCheckSituation.Text = CommonModule.SITUATION_MID_READ + " (1/4)";       // ----- 中間用プログレスバーラベル表示 -----
                Refresh();
                normalflg = Set_BaseMidFile_Copy(ref errstr, list_cv);
                if (normalflg == false)
                {
                    rtn = normalflg;
                    return rtn;
                }
            }
            if (CommonModule.MidChkCancelFlg)
            {
                return rtn;                                                      // 中断時移行の処理をスキップ
            }

            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
            // 通常コピーへの統合に伴いコメントアウト
            // '**************************************************
            // ' 変換コピー
            // '**************************************************
            // If CancelFlg = False And MidChkCancelFlg = False Then   '20160927 キャンセル処理を追加 CancelFlg = False の条件追加 -add   '20160929 データチェックのキャンセル処理を追加 And MidChkCancelFlg = False の条件も追加 -chg
            // Me.lblTotalSituation.Text = SITUATION_MID_READ & " (2/5)"       '----- 全体用プログレスバーラベル表示 -----
            // Me.lblCheckSituation.Text = SITUATION_MID_READ & " (2/5)"       '----- 中間用プログレスバーラベル表示 -----
            // Me.Refresh()
            // normalflg = Me.Set_BaseMidFile_ChgCopyMain(errstr, list_cv)
            // If normalflg = False Then
            // rtn = normalflg
            // Return rtn
            // End If
            // End If

            // If MidChkCancelFlg Then
            // Return rtn                                                      '中断時移行の処理をスキップ
            // End If
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end           
            // **************************************************
            // 自動設定処理 契約No等に「1」を自動設定                            '20160920 プログラム自動設定箇所の追加 -add
            // **************************************************
            if (CommonModule.CancelFlg == false & CommonModule.MidChkCancelFlg == false)   // 20160927 キャンセル処理を追加 CancelFlg = False の条件追加 -add   '20160929 データチェックのキャンセル処理を追加 And MidChkCancelFlg = False の条件も追加 -chg
            {
                lblTotalSituation.Text = CommonModule.SITUATION_DATACONDITION + " (2/4)";  // ----- 全体用プログレスバーラベル表示 -----
                lblCheckSituation.Text = CommonModule.SITUATION_DATACONDITION + " (2/4)";  // ----- 中間用プログレスバーラベル表示 -----
                Refresh();

                // ----- ログ出力 -----                                           '20161009 ログ出力処理追加 -add
                Console.WriteLine(Conversions.ToString(DateTime.Now) + "調整①：管理Noを「1」に設定");
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(902, "調整①"), false);
                int argrowcnt = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt);
                tmptmpcnt = argrowcnt;

                normalflg = Set_ExistMidFile_DefCopyMain(ref errstr);
                if (normalflg == false)
                {
                    rtn = normalflg;
                    return rtn;
                }
            }
            if (CommonModule.MidChkCancelFlg)
            {
                return rtn;                                                      // 中断時移行の処理をスキップ
            }

            // ******************************************************
            // ファイル名集約                                                    '20161011 汎用→既存書込時の不要処理除去修正 -add
            // ******************************************************
            // ----- ログ出力 -----                                               '20161009 ログ出力処理追加 -add
            Console.WriteLine(Conversions.ToString(DateTime.Now) + "調整②：ファイル名集約");
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(902, "調整②"), false);
            int argrowcnt1 = (int)tmptmpcnt;
            DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt1);
            tmptmpcnt = argrowcnt1;

            var list_filename = new List<string>();
            list_filename = Set_FileNameIntensive(list_cv);

            // ******************************************************
            // 自動設定処理 所有形態に一棟/区分の設定値を自動設定                 '20160920 プログラム自動設定箇所の追加 -add
            // ******************************************************
            if (CommonModule.CancelFlg == false & CommonModule.MidChkCancelFlg == false)
            {
                // 20161011 汎用→既存書込時の不要処理除去修正 -chg sta
                // normalflg = Me.Set_ExistMidFile_SyoyuCopyMain(errstr)
                lblTotalSituation.Text = CommonModule.SITUATION_DATACONDITION + " (3/4)";  // ----- 全体用プログレスバーラベル表示 -----
                lblCheckSituation.Text = CommonModule.SITUATION_DATACONDITION + " (3/4)";  // ----- 中間用プログレスバーラベル表示 -----
                Refresh();

                // ----- ログ出力 -----                                           '20161009 ログ出力処理追加 -add
                Console.WriteLine(Conversions.ToString(DateTime.Now) + "調整③：所有形態自動設定");
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(902, "調整③"), false);
                int argrowcnt2 = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt2);
                tmptmpcnt = argrowcnt2;

                normalflg = Set_ExistMidFile_SyoyuCopyMain(ref errstr, list_filename);
                // 20161011 汎用→既存書込時の不要処理除去修正 -chg end
                if (normalflg == false)
                {
                    rtn = normalflg;
                    return rtn;
                }
            }
            if (CommonModule.MidChkCancelFlg)
            {
                return rtn;                                                      // 中断時移行の処理をスキップ
            }

            // ******************************************************
            // 適用開始年月自動設定
            // ******************************************************
            // 20161004 適用開始年月の自動設定 -add sta
            if (CommonModule.CancelFlg == false & CommonModule.MidChkCancelFlg == false)
            {
                // 20161011 汎用→既存書込時の不要処理除去修正 -chg sta
                // normalflg = Me.Set_ExistMidFile_SoruleTekiyoYmdCopyMain(errstr)
                lblTotalSituation.Text = CommonModule.SITUATION_DATACONDITION + " (4/4)";  // ----- 全体用プログレスバーラベル表示 -----
                lblCheckSituation.Text = CommonModule.SITUATION_DATACONDITION + " (4/4)";  // ----- 中間用プログレスバーラベル表示 -----
                Refresh();

                // ----- ログ出力 -----                                           '20161009 ログ出力処理追加 -add
                Console.WriteLine(Conversions.ToString(DateTime.Now) + "調整④：適用開始年月自動設定");
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(902, "調整④"), false);
                int argrowcnt3 = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt3);
                tmptmpcnt = argrowcnt3;

                normalflg = Set_ExistMidFile_SoruleTekiyoYmdCopyMain(ref errstr, list_filename);
                // 20161011 汎用→既存書込時の不要処理除去修正 -chg end
                if (normalflg == false)
                {
                    rtn = normalflg;
                    return rtn;
                }
            }
            // 20161004 適用開始年月の自動設定 -add end
            if (CommonModule.MidChkCancelFlg)
            {
                return rtn;                                                      // 中断時移行の処理をスキップ
            }

            // 返却
            return rtn;

        }

        /// <summary>
        /// 既存用中間ファイルのヘッダー名をオブジェクトへ格納 '20160812 汎用コンバート対応
        /// '20160905 中間ファイルコピー処理改善 引数を追加
        /// </summary>
        /// <remarks></remarks>
        private void Set_MidHeaderInfoToObj()
        {

            var readtbl = new DataTable();
            int reccnt = 0;
            string fldname = "";
            string fldvalue = "";
            string tmp_既存中間シート名 = "";
            string tmp_既存中間フィールド名 = "";
            string tmp_汎用中間ファイル名 = "";
            string tmp_汎用中間シート名 = "";
            string tmp_汎用中間フィールド名_大分類 = "";
            string tmp_汎用中間フィールド名_小分類 = "";


            // --------------------------------------------------
            // 中間ファイル情報取得
            // --------------------------------------------------
            // 20160905 中間ファイルコピー処理改善 -chg sta
            // Dim tmp_sql As String = " SELECT DISTINCT * FROM " & MIDFILEINFO_DBNAME
            string tmp_sql = "";
            tmp_sql = tmp_sql + " SELECT DISTINCT * FROM " + MidFileInfoModule.MIDFILEINFO_DBNAME;
            // 20160905 中間ファイルコピー処理改善 -chg end
            bool argnormalflg = true;
            reccnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, normalflg: ref argnormalflg);

            // --------------------------------------------------
            // 初期化
            // --------------------------------------------------
            CommonModule.List_Existmidheader.Clear();
            CommonModule.List_Basemidheader.Clear();

            // --------------------------------------------------
            // 読込開始
            // --------------------------------------------------
            for (int cntii = 0, loopTo = readtbl.Rows.Count - 1; cntii <= loopTo; cntii++)
            {

                // 中断処理
                Application.DoEvents();
                if (CommonModule.CancelFlg | CommonModule.MidChkCancelFlg)
                {
                    return;
                }

                for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                {

                    // 項目名取得
                    fldname = readtbl.Columns[cntjj].ColumnName;

                    // 登録値取得
                    fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]);

                    // 各項目値→変数格納
                    switch (fldname ?? "")
                    {
                        case "既存中間シート名":
                            {
                                tmp_既存中間シート名 = fldvalue.Trim();
                                break;
                            }
                        case "既存中間フィールド名":
                            {
                                tmp_既存中間フィールド名 = fldvalue.Trim();
                                break;
                            }
                        case "汎用中間シート名":
                            {
                                tmp_汎用中間シート名 = fldvalue.Trim();
                                break;
                            }
                        case "汎用中間フィールド名_大分類":
                            {
                                tmp_汎用中間フィールド名_大分類 = fldvalue.Trim();
                                break;
                            }
                        case "汎用中間フィールド名_小分類":
                            {
                                tmp_汎用中間フィールド名_小分類 = fldvalue.Trim();
                                break;
                            }
                    }

                }

                // --------------------------------------------------
                // 既存用中間ファイルと紐付けて各オブジェクトへ格納
                // --------------------------------------------------
                string tmp_existmidheader = tmp_既存中間シート名 + CommonModule.STR_SPLIT_1 + tmp_既存中間フィールド名;
                if ((tmp_existmidheader ?? "") == CommonModule.STR_SPLIT_1)
                {
                    tmp_existmidheader = "";
                }

                // 既存用中間ファイルのヘッダー
                if (!string.IsNullOrEmpty(tmp_existmidheader))
                {
                    CommonModule.List_Existmidheader.Add(tmp_existmidheader);
                }

                // --------------------------------------------------
                // 汎用用中間ファイルと紐付けて各オブジェクトへ格納
                // --------------------------------------------------
                string tmp_basemidheader = "";
                if (!string.IsNullOrEmpty(tmp_汎用中間フィールド名_小分類))
                {
                    // 20160913_2 部屋設備移行処理の追加 -chg sta
                    // tmp_basemidheader = tmp_汎用中間シート名 & STR_SPLIT_1 & tmp_汎用中間フィールド名_小分類
                    tmp_basemidheader = tmp_汎用中間シート名 + CommonModule.STR_SPLIT_1 + tmp_汎用中間フィールド名_大分類 + CommonModule.STR_SPLIT_1 + tmp_汎用中間フィールド名_小分類;
                }
                // 20160913_2 部屋設備移行処理の追加 -chg end
                else
                {
                    tmp_basemidheader = tmp_汎用中間シート名 + CommonModule.STR_SPLIT_1 + tmp_汎用中間フィールド名_大分類;
                }
                if ((tmp_basemidheader ?? "") == CommonModule.STR_SPLIT_1)
                {
                    tmp_basemidheader = "";
                }

                // 汎用用中間ファイルのヘッダー
                if (!string.IsNullOrEmpty(tmp_basemidheader) & CommonModule.List_Basemidheader.Contains(tmp_basemidheader) == false)
                {
                    CommonModule.List_Basemidheader.Add(tmp_basemidheader);
                }

            }

        }

        /// <summary>
        /// 既存中間ファイル初期化処理 '20161007 既存中間ファイル初期化処理の追加 -add
        /// </summary>
        /// <param name="errstr"></param>
        /// <param name="list_cv"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        //private bool Init_ExistMidFile(ref string errstr, List<string> list_cv)
        //{

        //    Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
        //    Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
        //    Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
        //    var excelfile = new ExcelFileManager();                // Excelファイル操作用
        //    var list_filename = new List<string>();


        //    // ファイル名を集約
        //    foreach (var cvitem in list_cv)
        //    {
        //        string[] tmp_str = cvitem.Split('-');
        //        string filename = tmp_str[0];
        //        if (list_filename.Contains(filename) == false)
        //        {
        //            list_filename.Add(filename);
        //        }
        //    }

        //    foreach (var initfile in list_filename)
        //    {

        //        // 中断処理
        //        Application.DoEvents();
        //        if (CommonModule.CancelFlg | CommonModule.MidChkCancelFlg)
        //        {
        //            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
        //            return false;
        //        }

        //        // ファイルパス設定
        //        string filepath = EtcMethod.Set_Path(CommonModule.MiddleDirPath, initfile + ".xlsx");

        //        // 初期化処理
        //        try
        //        {
        //            // ファイル準備
        //            appli = (Microsoft.Office.Interop.Excel.Application)Interaction.CreateObject("Excel.Application");
        //            appli.Visible = false;
        //            wbook = appli.Workbooks.Open(filepath);

        //            for (int cntii = 1, loopTo = wbook.Worksheets.Count; cntii <= loopTo; cntii++)
        //            {
        //                // シートを取得
        //                wsheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets[cntii];
        //                // 最大行を取得
        //                int maxrowcnt = wsheet.UsedRange.Rows.Count;

        //                if (maxrowcnt > 1)
        //                {
        //                    // 既存データ範囲を取得
        //                    string dataarea = CommonModule.EXISTMIDFILE_READWRITE_ROW.ToString() + ":" + maxrowcnt.ToString();
        //                    wsheet.Rows[dataarea].Delete();
        //                }
        //            }
        //        }

        //        catch (Exception ex)
        //        {
        //            errstr = CommonModule.MSG_ERR_EXISTMIDFILEINIT;
        //            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
        //            return false;
        //        }

        //        // 終了処理
        //        excelfile.Set_ExcelFile_WriteClose(ref appli, ref wbook, ref wsheet);

        //    }

        //    return true;

        //}

        /// <summary>
        /// 汎用用中間ファイル→既存用仮テーブル通常処理 '20160929 汎用→既存コピー処理改善対応
        /// </summary>
        /// <param name="errstr"></param>
        /// <param name="list_cv"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_BaseMidFile_Copy(ref string errstr, List<string> list_cv)
        {

            bool rtn = true;

            string tmpprovider = "Microsoft.ACE.OLEDB.12.0; ";                        // EXCEL2007以上(xlsx)
            string tmpextend = "Excel 8.0;HDR=YES;";

            string tmptmpstr;                                                         // 20161009 ログ出力処理追加 -add
            var tmptmpcnt = default(long);                                                           // 20161009 ログ出力処理追加 -add

            // ----- 全体用・中間用プログレスバー更新 -----
            var obj_pgb = new ProgressBarManager();
            int pgbcnt = 0;
            int pgbtotalcnt_total = list_cv.Count;
            obj_pgb.pgbInitTotal(pgbtotalcnt_total);
            obj_pgb.pgbInitChkTotal(pgbtotalcnt_total);
            pgbtotalcnt = pgbtotalcnt_total;

            // 20161017 Excel接続処理の外出し修正 -add sta
            // ----- 汎用用中間ファイルへの接続・オープン処理 -----
            MainFrmHelper.init();
            // 20161017 Excel接続処理の外出し修正 -add end

            foreach (var cvitem_tmp in list_cv)
            {
                var cvitem = cvitem_tmp;

                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
                // 変換コピー処理を統合
                // 自社口座、部屋設備はそのまま仮テーブルを作成する
                // 既存処理は「Case Else」の中へそのまま移動
                switch (cvitem ?? "")
                {
                    case "部屋情報-部屋駐車場情報":
                    case "部屋情報-部屋特約情報":
                    case "契約情報-契約契約者情報":
                    case "契約情報-契約保証人情報":
                    case "契約情報-契約特約事項情報":
                        {
                            // 変換コピー
                            rtn = Set_BaseMidFile_ChgCopyMain(cvitem, ref errstr);
                            if (rtn == false)
                            {
                                return rtn;
                            }

                            break;
                        }
                    // 20161025 部屋設備取得方法の修正 -chg sta
                    // Case "自社情報-自社口座情報", "部屋情報-部屋設備情報"
                    // '汎用用仮テーブル作成
                    // rtn = Me.Set_BaseMidFile_NormalCopyMain(cvitem, con_read, cmd_read, errstr)
                    // If rtn = False Then
                    // con_read.Close()
                    // Return rtn
                    // End If
                    case "自社情報-自社口座情報":
                        {
                            // 汎用用仮テーブル作成
                            rtn = Set_BaseMidFile_JisyaCopyMain(cvitem, ref errstr);
                            if (rtn == false)
                            {
                                return rtn;
                            }

                            break;
                        }
                    case "部屋情報-部屋設備情報":
                        {
                            // 汎用用仮テーブル作成
                            rtn = Set_BaseMidFile_SetubiCopyMain(cvitem, ref errstr);
                            if (rtn == false)
                            {
                                return rtn;
                            }

                            break;
                        }
                    // 20161025 部屋設備取得方法の修正 -chg end
                    // 20161028 物件/部屋鍵取得方法修正 -add sta
                    case "各マスタ情報-鍵タイトルマスタ":
                        {
                            // 汎用用仮テーブル作成
                            rtn = Set_BaseMidFile_KagiTitleCopyMain(cvitem, ref errstr);
                            if (rtn == false)
                            {
                                return rtn;
                                // 20161028 物件/部屋鍵取得方法修正 -add end
                            }

                            break;
                        }

                    default:
                        {
                            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end
                            // 20161017 DoEvent処理の移動 -del sta
                            // '中断処理
                            // Application.DoEvents()
                            // If CancelFlg Or MidChkCancelFlg Then
                            // Return rtn
                            // End If
                            // 20161017 DoEvent処理の移動 -del end
                            // コピーフラグ
                            bool copyflg = true;

                            // --------------------------------------------------
                            // 既存データ関連
                            // --------------------------------------------------
                            // コンバート対象項目から既存ファイル名、シート名を取得
                            cvitem = cvitem.Replace("-", CommonModule.STR_SPLIT_1);
                            string[] tmp_existmidstr = Strings.Split(cvitem, CommonModule.STR_SPLIT_1);
                            string existfilename = tmp_existmidstr[0];
                            string existsheetname = tmp_existmidstr[1];
                            string existmidpath = "";                                                     // 既存中間ファイルパス

                            // 既存中間ファイルパスを取得
                            existmidpath = EtcMethod.Set_Path(CommonModule.MiddleDirPath, existfilename + ".xlsx");

                            // 照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
                            Set_MidDataInfoToObj(existfilename, existsheetname, false);

                            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
                            int tmp_cnt = 0;
                            string tmptblname = CommonModule.PRE_TBL_NAME + existsheetname;            // シート名は仮テーブルのテーブル名の一部にも使用する
                                                                                                       // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

                            // --------------------------------------------------
                            // 汎用データ関連
                            // --------------------------------------------------
                            // 移行項目からオープンする汎用用中間ファイル名を取得
                            string basemidfsname = "";                                                    // 汎用中間ファイルのファイル名-シート名格納用
                            string basefilename = "";                                                     // 汎用中間ファイル名
                            string basesheetname = "";                                                    // 汎用中間シート名
                                                                                                          // 20161017 Excel接続処理の外出し修正 -del
                                                                                                          // Dim basemidpath As String = ""                                                      '汎用中間ファイルパス

                            basemidfsname = Conversions.ToString(CommonModule.Hash_ExistMidToBaseMid_FS[cvitem]);
                            if (!string.IsNullOrEmpty(basemidfsname))
                            {
                                string[] tmp_basemidstr = Strings.Split(basemidfsname, CommonModule.STR_SPLIT_1);
                                basefilename = tmp_basemidstr[0];
                                basesheetname = tmp_basemidstr[1];
                            }
                            // 20161017 Excel接続処理の外出し修正 -del
                            // basemidpath = EtcMethod.Set_Path(BaseMidDirPath, basefilename & ".xlsx")        '汎用中間ファイルパスを取得
                            else
                            {
                                copyflg = false;
                            }

                            // --------------------------------------------------
                            // メイン処理
                            // --------------------------------------------------

                            if (copyflg)
                            {
                                // 20161017 Excel接続処理の外出し修正 -del sta
                                // Dim con_read As New OleDbConnection()
                                // Dim cmd_read As New OleDbCommand()
                                // 20161017 Excel接続処理の外出し修正 -del end
                                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                                // Dim con_write As New OleDbConnection()
                                // Dim cmd_write As New OleDbCommand()
                                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                                try
                                {
                                    // 20161017 Excel接続処理の外出し修正 -del sta
                                    // '--------------------------------------------------
                                    // ' ①EXCEL接続 (ファイル読込用)
                                    // '--------------------------------------------------
                                    // '接続文字列生成
                                    // con_read.ConnectionString = _
                                    // "Provider=" & tmpprovider & _
                                    // "Data Source=" & basemidpath & ";" & _
                                    // "Extended Properties=" & """" & tmpextend & """"
                                    // '接続設定
                                    // cmd_read.Connection = con_read
                                    // 20161017 Excel接続処理の外出し修正 -del end
                                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                                    // '--------------------------------------------------
                                    // ' ②EXCEL接続 (ファイル書込用)
                                    // '--------------------------------------------------
                                    // '接続文字列生成
                                    // con_write.ConnectionString = _
                                    // "Provider=" & tmpprovider & _
                                    // "Data Source=" & existmidpath & ";" & _
                                    // "Extended Properties=" & """" & tmpextend & """"
                                    // '接続設定
                                    // cmd_write.Connection = con_write

                                    // '--------------------------------------------------
                                    // ' 読込→書込
                                    // '--------------------------------------------------
                                    // '接続オープン処理
                                    // '20161017 Excel接続処理の外出し修正 -del
                                    // 'con_read.Open()
                                    // con_write.Open()
                                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                                    // ----- ログ出力 -----                                           '20161009 ログ出力処理追加 -add
                                    Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDCOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + basesheetname + ")");
                                    tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(901, CommonModule.LOG_SYORIKOMK_MIDFILE, basesheetname), false);
                                    int argrowcnt = (int)tmptmpcnt;
                                    DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt);
                                    tmptmpcnt = argrowcnt;

                                    var dbSetName = MainFrmHelper.dic汎用中間ファイル[basesheetname].First().Value.Preテーブル.TableName;
                                    if (string.IsNullOrEmpty(dbSetName)) break;
                                    if (MainFrmHelper.CheckTableExists(MainFrmHelper.db, dbSetName) == false) break;
                                    var dbSet = MainFrmHelper.GetDbSet(MainFrmHelper.db, dbSetName);

                                    int readrowcnt = 1;

                                    // コピー項目表示
                                    lblCVItem.Text = existfilename + " ： " + existsheetname;

                                    // ----- 個別プログレスバー初期化 -----
                                    int pgbtotalcnt_part = 0;                             // プログレスバー総件数初期化
                                    int pgbbasecnt = 100;                                 // 実件数で表示するか否かの基準値
                                    if (dbSet.Count <= pgbbasecnt)
                                    {
                                        pgbtotalcnt_part = dbSet.Count;
                                    }
                                    else
                                    {
                                        pgbtotalcnt_part = (int)Math.Round(Math.Ceiling(dbSet.Count / (double)pgbbasecnt));
                                    }

                                    // 1行ずつ取得
                                    obj_pgb.pgbInitPart(pgbtotalcnt_part);
                                    foreach (var entity in dbSet)
                                    {

                                        // 20161017 DoEvent処理の移動 -add sta
                                        // 中断処理
                                        Application.DoEvents();
                                        if (CommonModule.CancelFlg | CommonModule.MidChkCancelFlg)
                                        {
                                            return rtn;
                                        }
                                        // 20161017 DoEvent処理の移動 -add end

                                        var tmp_hash = new SafeDictionary<string, string>();

                                        // 汎用ヘッダーから既存ヘッダーを取得し、実データを紐付けてオブジェクトへ格納する
                                        foreach (var existmid in CommonModule.Hash_ExistMidToBaseMid_SH)
                                        {
                                            // 既存中間ファイルと汎用中間ファイルの紐付状況を取得
                                            string existmidshname_fromobj = Conversions.ToString(existmid.Key);
                                            string basemidshname_fromobj = Conversions.ToString(existmid.Value);
                                            if (!string.IsNullOrEmpty(basemidshname_fromobj))
                                            {
                                                string[] tmp_str = Strings.Split(basemidshname_fromobj, CommonModule.STR_SPLIT_1);
                                                var tmp_basesheetname = tmp_str[0];
                                                var tmp_baseheader = tmp_str[1];
                                                var entityColName = MainFrmHelper.dic汎用中間ファイル[tmp_basesheetname][tmp_baseheader].Preテーブル.ColumnName;
                                                if (!string.IsNullOrEmpty(entityColName))
                                                {
                                                    string tmp_basevalue = MainFrmHelper.GetPropertyValue(entity, entityColName);
                                                    tmp_str = Strings.Split(existmidshname_fromobj, CommonModule.STR_SPLIT_1);
                                                    tmp_hash.Add(tmp_str[1], tmp_basevalue);
                                                }
                                            }
                                        }

                                        // 挿入用に編集
                                        string taisyo = "";
                                        string value = "";

                                        // 20161205 初回契約日を契約日から取得するように修正 -add sta
                                        // 契約日が存在しない場合があるためチェックを行う
                                        // 存在しない場合は契約開始日をセットする
                                        if ((cvitem ?? "") == "契約情報" + CommonModule.STR_SPLIT_1 + "契約基本情報")
                                        {
                                            // 初回契約日チェック
                                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(tmp_hash["初回契約日"], "", false)))
                                            {
                                                tmp_hash["初回契約日"] = tmp_hash["契約開始日"];
                                            }
                                        }
                                        if ((cvitem ?? "") == "契約情報" + CommonModule.STR_SPLIT_1 + "契約履歴情報")
                                        {
                                            // 契約日チェック
                                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(tmp_hash["契約日"], "", false)))
                                            {
                                                tmp_hash["契約日"] = tmp_hash["契約開始日"];
                                            }
                                        }
                                        // 20161205 初回契約日を契約日から取得するように修正 -add end

                                        foreach (var item in tmp_hash)
                                        {
                                            string tmp_taisyo = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("[", item.Key), "]"));
                                            // 20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg sta
                                            // Dim tmp_value As String = "'" & item.Value & "'"
                                            string tmp_value = Conversions.ToString(item.Value);
                                            if (tmp_value is null)
                                            {
                                                tmp_value = "";
                                            }
                                            tmp_value = "'" + tmp_value.Replace("'", "’") + "'";
                                            // 20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg end
                                            taisyo = taisyo + "," + tmp_taisyo;
                                            value = value + "," + tmp_value;
                                        }
                                        taisyo = "(" + taisyo.Remove(0, 1) + ")";
                                        value = "(" + value.Remove(0, 1) + ")";

                                        // 挿入
                                        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                        // qry = ""
                                        // qry += " INSERT INTO [" & existsheetname & "$] "
                                        // qry += taisyo
                                        // qry += " VALUES "
                                        // qry += value
                                        // cmd_write.CommandText = qry
                                        // cmd_write.ExecuteNonQuery()
                                        string tmp_qry_insert = "";
                                        tmp_qry_insert += " INSERT INTO " + tmptblname;
                                        tmp_qry_insert += taisyo;
                                        tmp_qry_insert += " VALUES ";
                                        tmp_qry_insert += value;
                                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_qry_insert, ref tmp_cnt);
                                        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end

                                        // ----- プログレスバー更新/進捗率表示 -----
                                        int tmp_pgbcnt = 0;
                                        if (dbSet.Count <= pgbbasecnt)             // 基準件数(100件)以下の場合は実件数を取得
                                        {
                                            tmp_pgbcnt = readrowcnt;
                                        }
                                        else if (dbSet.Count > readrowcnt)          // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                                        {
                                            EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, ref tmp_pgbcnt);
                                        }
                                        else if (dbSet.Count <= readrowcnt)         // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                                        {
                                            tmp_pgbcnt = pgbtotalcnt_part;
                                        }

                                        // 表示
                                        if (tmp_pgbcnt != 0)
                                        {
                                            obj_pgb.pgbsettingPart(tmp_pgbcnt);
                                            obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part);
                                        }
                                        // 20161017 進捗表示処理による速度低下の修正 -del
                                        // Me.Refresh()
                                        readrowcnt = readrowcnt + 1;
                                    }
                                }

                                catch (Exception ex)
                                {

                                    // エラー処理
                                    errstr = CommonModule.MSG_ERR_BASEMIDFILEREAD;
                                    rtn = false;

                                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                                    // con_write.Close()

                                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
                                    return rtn;

                                }

                                // クローズ処理
                                // 20161017 Excel接続処理の外出し修正 -del
                                // con_read.Close()
                                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                                // con_write.Close()
                            }

                            break;
                        }

                }

                // ----- 全体用プログレスバー更新 -----
                pgbcnt = pgbcnt + 1;
                obj_pgb.pgbsettingTotal(pgbcnt);
                obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt, true);
                // ----- 中間用プログレスバー更新 -----
                obj_pgb.pgbsettingChkTotal(pgbcnt);
                obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt, true);
                Refresh();

            }

            // 20161017 Excel接続処理の外出し修正 -add

            return rtn;

        }
        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
        /// <summary>
        /// 汎用用中間ファイル→既存用仮テーブル個別処理 
        /// </summary>
        /// <param name="cvitem"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_BaseMidFile_ChgCopyMain(string cvitem, ref string errstr)
        {

            bool rtn = true;
            var obj_rep = new object();
            string tmptmpstr = "";                                // 20161009 ログ出力処理追加 -add
            var tmptmpcnt = default(long);                                       // 20161009 ログ出力処理追加 -add
            string basemidfilesheetname = "";

            // ----- 仮テーブル名取得 -----
            string[] tmp_str = cvitem.Split('-');
            string existsheetname = tmp_str[1];
            string tmptblname = CommonModule.PRE_TBL_NAME + existsheetname;            // シート名は仮テーブルのテーブル名の一部にも使用する

            // ----- 各項目毎に移行用のオブジェクトを生成 -----
            switch (cvitem ?? "")
            {
                case "部屋情報-部屋駐車場情報":
                    {
                        obj_rep = new Repository.Hy_cyusyajo_mid_Repository();
                        lblCVItem.Text = cvitem.Replace("-", " ： ");
                        basemidfilesheetname = "部屋情報";
                        break;
                    }
                case "部屋情報-部屋特約情報":
                    {
                        obj_rep = new Repository.Hy_tokuyaku_mid_Repository();
                        lblCVItem.Text = cvitem.Replace("-", " ： ");
                        basemidfilesheetname = "部屋情報";
                        break;
                    }
                case "契約情報-契約契約者情報":
                    {
                        obj_rep = new Repository.Ky_kys_mid_Repository();
                        lblCVItem.Text = cvitem.Replace("-", " ： ");
                        basemidfilesheetname = "契約契約者保証人情報";
                        break;
                    }
                case "契約情報-契約保証人情報":
                    {
                        obj_rep = new Repository.Ky_hosyonin_mid_Repository();
                        lblCVItem.Text = cvitem.Replace("-", " ： ");
                        basemidfilesheetname = "契約契約者保証人情報";
                        break;
                    }
                case "契約情報-契約特約事項情報":
                    {
                        obj_rep = new Repository.Ky_tokuyaku_mid_Repository();
                        lblCVItem.Text = cvitem.Replace("-", " ： ");
                        basemidfilesheetname = "契約特約およびメモ情報";
                        break;
                    }
            }

            // ----- ログ出力 -----                       '20161009 ログ出力処理追加 -add
            Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDONECOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + cvitem + ")");
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(903, CommonModule.LOG_SYORIKOMK_MIDFILE, cvitem), false);
            int argrowcnt = (int)tmptmpcnt;
            DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt);
            tmptmpcnt = argrowcnt;

            // ----- データ取得 -----
            var dbSetName = MainFrmHelper.dic汎用中間ファイル[basemidfilesheetname].First().Value.Preテーブル.TableName;
            if (MainFrmHelper.CheckTableExists(MainFrmHelper.db, dbSetName) == false) return true;
            var dbSet = MainFrmHelper.GetDbSet(MainFrmHelper.db, dbSetName);

            // ----- 項目毎に個別処理開始 -----
            rtn = Conversions.ToBoolean(((dynamic)obj_rep).Read_BaseMidFile(sqlcnnv10, MainFrmHelper.dic汎用中間ファイル[basemidfilesheetname], dbSet, tmptblname));

            // ----- コピー処理失敗時は処理を抜ける -----
            if (rtn == false)
            {

                errstr = CommonModule.MSG_ERR_BASEMIDFILEREAD;

                // ----- ログ出力 -----                   '20161009 ログ出力処理追加 -add
                Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_ERR_MIDONECOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + cvitem + ")");
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, CommonModule.LOG_SYORIKOMK_MIDFILE, cvitem), false);
                int argrowcnt1 = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt1);
                tmptmpcnt = argrowcnt1;
                return rtn;

            }

            // ----- 返却 -----
            return rtn;

        }
        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
        /// <summary>
        /// 汎用用中間ファイル→既存用仮テーブル個別処理_自社 
        /// 自社口座情報は汎用中間ファイルの形でそのまま仮テーブルを作成する
        /// </summary>
        /// <param name="cvitem"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_BaseMidFile_JisyaCopyMain(string cvitem, ref string errstr)
        {

            bool rtn = true;
            var obj_rep = new object();
            string tmptmpstr = "";                                // 20161009 ログ出力処理追加 -add
            var tmptmpcnt = default(long);                                       // 20161009 ログ出力処理追加 -add

            string basesheetname = "";
            var obj_pgb = new ProgressBarManager();

            // ----- 仮テーブル名取得 -----
            string[] tmp_str = cvitem.Split('-');
            string existsheetname = tmp_str[1];
            int tmp_cnt = 0;
            string tmptblname = CommonModule.PRE_TBL_NAME + existsheetname;            // シート名は仮テーブルのテーブル名の一部にも使用する
            // 20161025 部屋設備取得方法の修正 -add sta
            basesheetname = tmp_str[1];
            lblCVItem.Text = cvitem.Replace("-", " ： ");
            // 20161025 部屋設備取得方法の修正 -add end

            // 20161025 部屋設備取得方法の修正 -del sta
            // '----- 各項目毎に移行用のオブジェクトを生成 -----
            // Select Case cvitem
            // Case "自社情報-自社口座情報"
            // basesheetname = "自社口座情報"
            // Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
            // Case "部屋情報-部屋設備情報"
            // basesheetname = "部屋設備情報"
            // Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
            // End Select
            // 20161025 部屋設備取得方法の修正 -del end
            // ----- ログ出力 -----             
            Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDONECOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + cvitem + ")");
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(903, CommonModule.LOG_SYORIKOMK_MIDFILE, cvitem), false);
            int argrowcnt = (int)tmptmpcnt;
            DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt);
            tmptmpcnt = argrowcnt;

            // ----- メイン処理 -----
            try
            {

                // ----- ログ出力 -----                                          
                Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDCOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + basesheetname + ")");
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(901, CommonModule.LOG_SYORIKOMK_MIDFILE, basesheetname), false);
                int argrowcnt1 = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt1);
                tmptmpcnt = argrowcnt1;

                var dic = MainFrmHelper.dic汎用中間ファイル[basesheetname];
                var dbSetName = dic.First().Value.Preテーブル.TableName;
                if (MainFrmHelper.CheckTableExists(MainFrmHelper.db, dbSetName) == false) return true;
                var dbSet = MainFrmHelper.GetDbSet(MainFrmHelper.db, dbSetName);

                int readrowcnt = 1;

                // ----- 個別プログレスバー初期化 -----
                int pgbtotalcnt_part = 0;                             // プログレスバー総件数初期化
                int pgbbasecnt = 100;                                 // 実件数で表示するか否かの基準値
                if (dbSet.Count <= pgbbasecnt)
                {
                    pgbtotalcnt_part = dbSet.Count;
                }
                else
                {
                    pgbtotalcnt_part = (int)Math.Round(Math.Ceiling(dbSet.Count / (double)pgbbasecnt));
                }

                // 1行ずつ取得
                obj_pgb.pgbInitPart(pgbtotalcnt_part);
                foreach (var row in dbSet)
                {

                    // 中断処理
                    Application.DoEvents();
                    if (CommonModule.CancelFlg | CommonModule.MidChkCancelFlg)
                    {
                        return rtn;
                    }

                    var tmp_hash = new SafeDictionary<string, string>();

                    // 行データを取得して作業用オブジェクトへ格納
                    string insertvalue = "";
                    foreach (var tmp_baseheader in dic.Keys)
                    {
                        var entityColName = dic[tmp_baseheader].Preテーブル.ColumnName;
                        if (!string.IsNullOrEmpty(entityColName))
                        {
                            string tmp_basevalue = MainFrmHelper.GetPropertyValue(row, entityColName);
                            // 20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg sta
                            // insertvalue = insertvalue & "," & "'" & row(cntjj).ToString & "'"
                            insertvalue = insertvalue + "," + "'" + tmp_basevalue.ToString().Replace("'", "’") + "'";
                            // 20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg end
                        }
                    }

                    // 挿入用に成形
                    insertvalue = insertvalue.Remove(0, 1);

                    // 挿入処理
                    string tmp_sql_insert = " INSERT INTO " + tmptblname + " VALUES(" + insertvalue + ")";
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_cnt);

                    // ----- プログレスバー更新/進捗率表示 -----
                    int tmp_pgbcnt = 0;
                    if (dbSet.Count <= pgbbasecnt)             // 基準件数(100件)以下の場合は実件数を取得
                    {
                        tmp_pgbcnt = readrowcnt;
                    }
                    else if (dbSet.Count > readrowcnt)          // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, ref tmp_pgbcnt);
                    }
                    else if (dbSet.Count <= readrowcnt)         // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                    {
                        tmp_pgbcnt = pgbtotalcnt_part;
                    }

                    // 表示
                    if (tmp_pgbcnt != 0)
                    {
                        obj_pgb.pgbsettingPart(tmp_pgbcnt);
                        obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part);
                    }

                    readrowcnt = readrowcnt + 1;

                }
            }

            catch (Exception ex)
            {

                // エラー処理
                rtn = false;
                errstr = CommonModule.MSG_ERR_BASEMIDFILEREAD;

                // ----- ログ出力 -----              
                Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_ERR_MIDONECOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + cvitem + ")");
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, CommonModule.LOG_SYORIKOMK_MIDFILE, cvitem), false);
                int argrowcnt2 = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt2);
                tmptmpcnt = argrowcnt2;

                return rtn;

            }

            return rtn;

        }
        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

        // 20161025 部屋設備取得方法の修正 -add sta
        /// <summary>
        /// 汎用用中間ファイル→既存用仮テーブル個別処理_設備
        /// </summary>
        /// <param name="cvitem"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_BaseMidFile_SetubiCopyMain(string cvitem, ref string errstr)
        {

            var rowcnt = default(int);                                           // 書込行数
            var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
            var obj_com = new CommonRepository();                  // 共通処理用
            bool rtn = true;                                       // 戻り値
            int tmp_cnt = 0;
            string tmptmpstr = "";
            var tmptmpcnt = default(long);

            // ************************
            // 作業準備
            // ************************

            // 移行項目、仮テーブル名セット
            string[] tmp_str = cvitem.Split('-');
            string basesheetname = tmp_str[1];
            lblCVItem.Text = cvitem.Replace("-", " ： ");
            string tmptblname = CommonModule.PRE_TBL_NAME + basesheetname;            // シート名は仮テーブルのテーブル名の一部にも使用する

            var dic = MainFrmHelper.dic汎用中間ファイル[basesheetname];
            var dbSetName = dic.First().Value.Preテーブル.TableName;
            if (MainFrmHelper.CheckTableExists(MainFrmHelper.db, dbSetName) == false) return true;
            var dbSet = MainFrmHelper.GetDbSet(MainFrmHelper.db, dbSetName);
            rowcnt = dbSet.Count;

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

            // ----- ログ出力 -----                          
            Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDCOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + basesheetname + ")");
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(901, CommonModule.LOG_SYORIKOMK_MIDFILE, basesheetname), false);
            int argrowcnt1 = (int)tmptmpcnt;
            DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt1);
            tmptmpcnt = argrowcnt1;

            {
                ref var withBlock = ref model_cvitem;

                // ---------------
                // データ部処理
                // ---------------
                int cntii = 1;
                foreach (var row in dbSet)
                {

                    // 中断処理
                    Application.DoEvents();
                    if (CommonModule.CancelFlg)
                    {
                        return rtn;
                    }

                    // 挿入用に成形
                    string insertvalue = "";
                    foreach (var key in dic.Keys)
                    {
                        var entityColName = dic[key].Preテーブル.ColumnName;
                        if (!string.IsNullOrEmpty(entityColName))
                        {
                            string value = MainFrmHelper.GetPropertyValue(row, entityColName);
                            // 20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg sta
                            // insertvalue = insertvalue & "," & "'" & datarowvalue(1, cntjj) & "'"
                            insertvalue = insertvalue + "," + "'" + value.ToString().Replace("'", "’") + "'";
                        }
                        // 20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg end
                    }
                    if (!string.IsNullOrEmpty(insertvalue.Trim()))
                    {
                        insertvalue = insertvalue.Remove(0, 1);
                    }

                    // 挿入処理
                    string tmp_sql_insert = " INSERT INTO " + tmptblname + " VALUES(" + insertvalue + ")";
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_cnt);

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
                    cntii++;
                }

            }

            // ************************
            // 終了処理
            // ************************
            // 20161027 部屋設備ログ出力箇所修正 -del sta
            // '----- ログ出力 -----               
            // Console.WriteLine(Now & " " & LOG_SYORIKOMK_ERR_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
            // tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, LOG_SYORIKOMK_MIDFILE, cvitem), False)
            // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)
            // 20161027 部屋設備ログ出力箇所修正 -del end

            // 返却
            return rtn;

        }
        // 20161025 部屋設備取得方法の修正 -add end

        // 20161028 物件/部屋鍵取得方法修正 -add sta
        /// <summary>
        /// 汎用用中間ファイル→既存用仮テーブル個別処理_鍵タイトル
        /// </summary>
        /// <param name="cvitem"></param>
        /// <param name="con_read"></param>
        /// <param name="cmd_read"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_BaseMidFile_KagiTitleCopyMain(string cvitem, ref string errstr)
        {

            bool rtn = true;
            var obj_rep = new object();
            string tmptmpstr = "";
            var tmptmpcnt = default(long);

            string[] basesheetname = new string[] { "", "物件鍵タイトルマスタ", "部屋鍵タイトルマスタ" };   // 要素数も使用するため1から値をセットする
            var obj_pgb = new ProgressBarManager();

            // ----- 仮テーブル名取得 -----
            string[] tmp_str = cvitem.Split('-');
            string existsheetname = tmp_str[1];
            int tmp_cnt = 0;
            string tmptblname = CommonModule.PRE_TBL_NAME + existsheetname;            // シート名は仮テーブルのテーブル名の一部にも使用する

            // ----- ログ出力 -----             
            Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDONECOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + cvitem + ")");
            tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(903, CommonModule.LOG_SYORIKOMK_MIDFILE, cvitem), false);
            int argrowcnt = (int)tmptmpcnt;
            DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt);
            tmptmpcnt = argrowcnt;

            // ----- メイン処理 -----
            try
            {

                // ----- ログ出力 -----                                          
                Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDCOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + existsheetname + ")");
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(901, CommonModule.LOG_SYORIKOMK_MIDFILE, existsheetname), false);
                int argrowcnt1 = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt1);
                tmptmpcnt = argrowcnt1;

                for (int cntii = 1, loopTo = Information.UBound(basesheetname); cntii <= loopTo; cntii++)
                {

                    // 移行項目表示
                    lblCVItem.Text = basesheetname[cntii];

                    var dic = MainFrmHelper.dic汎用中間ファイル[basesheetname[cntii]];
                    var dbSetName = dic.First().Value.Preテーブル.TableName;
                    if (MainFrmHelper.CheckTableExists(MainFrmHelper.db, dbSetName) == false) continue;
                    var dbSet = MainFrmHelper.GetDbSet(MainFrmHelper.db, dbSetName);

                    int readrowcnt = 1;

                    // ----- 個別プログレスバー初期化 -----
                    int pgbtotalcnt_part = 0;                             // プログレスバー総件数初期化
                    int pgbbasecnt = 100;                                 // 実件数で表示するか否かの基準値
                    if (dbSet.Count <= pgbbasecnt)
                    {
                        pgbtotalcnt_part = dbSet.Count;
                    }
                    else
                    {
                        pgbtotalcnt_part = (int)Math.Round(Math.Ceiling(dbSet.Count / (double)pgbbasecnt));
                    }

                    // 1行ずつ取得
                    obj_pgb.pgbInitPart(pgbtotalcnt_part);
                    foreach (var row in dbSet)
                    {

                        // 中断処理
                        Application.DoEvents();
                        if (CommonModule.CancelFlg | CommonModule.MidChkCancelFlg)
                        {
                            return rtn;
                        }

                        var tmp_hash = new SafeDictionary<string, string>();

                        // 行データを取得して作業用オブジェクトへ格納
                        string insertvalue = "";
                        foreach (var key in dic.Keys)
                        {
                            var entityColName = dic[key].Preテーブル.ColumnName;
                            if (!string.IsNullOrEmpty(entityColName))
                            {
                                string value = MainFrmHelper.GetPropertyValue(row, entityColName);
                                insertvalue = insertvalue + "," + "'" + value.ToString().Replace("'", "’") + "'";
                                // 20161125 仮テーブル作成時に「'」を全角変換する処理を追加 -chg end
                            }
                        }

                        // 挿入用に成形
                        insertvalue = insertvalue.Remove(0, 1);
                        insertvalue = "'" + cntii.ToString() + "'" + "," + insertvalue;

                        // 挿入処理
                        string tmp_sql_insert = " INSERT INTO " + tmptblname + " VALUES(" + insertvalue + ")";
                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_cnt);

                        // ----- プログレスバー更新/進捗率表示 -----
                        int tmp_pgbcnt = 0;
                        if (dbSet.Count <= pgbbasecnt)             // 基準件数(100件)以下の場合は実件数を取得
                        {
                            tmp_pgbcnt = readrowcnt;
                        }
                        else if (dbSet.Count > readrowcnt)          // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                        {
                            EtcMethod.Get_MultipleFlg(readrowcnt, pgbbasecnt, ref tmp_pgbcnt);
                        }
                        else if (dbSet.Count <= readrowcnt)         // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                        {
                            tmp_pgbcnt = pgbtotalcnt_part;
                        }

                        // 表示
                        if (tmp_pgbcnt != 0)
                        {
                            obj_pgb.pgbsettingPart(tmp_pgbcnt);
                            obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part);
                        }

                        readrowcnt = readrowcnt + 1;

                    }

                }
            }

            catch (Exception ex)
            {

                // エラー処理
                rtn = false;
                errstr = CommonModule.MSG_ERR_BASEMIDFILEREAD;

                // ----- ログ出力 -----              
                Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_ERR_MIDONECOPY + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + cvitem + ")");
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, CommonModule.LOG_SYORIKOMK_MIDFILE, cvitem), false);
                int argrowcnt2 = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt2);
                tmptmpcnt = argrowcnt2;

                return rtn;

            }

            return rtn;

        }
        // 20161028 物件/部屋鍵取得方法修正 -add end

        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
        // 通常コピーと個別コピーの統合に伴い、もともとあった処理をコメントアウト
        // ''' <summary>
        // ''' 汎用用中間ファイル→既存用中間ファイル個別処理 '20160929 汎用→既存コピー処理改善対応
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <param name="list_cv"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidFile_ChgCopyMain(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

        // Dim rtn As Boolean = True
        // Dim obj_rep As New Object
        // Dim model_basemiditem As New Object
        // Dim basemidfilename As String = ""
        // Dim basemidfilesheetname As String = ""
        // Dim existmidfilename As String = ""
        // Dim existmidfilesheetname As String = ""

        // Dim tmptmpstr As String = ""                                '20161009 ログ出力処理追加 -add
        // Dim tmptmpcnt As Long                                       '20161009 ログ出力処理追加 -add

        // '----- 全体用・中間用プログレスバー更新 -----
        // Dim obj_pgb As New ProgressBarManager
        // Dim pgbcnt As Integer = 0
        // Dim pgbtotalcnt_total As Integer = list_cv.Count
        // Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
        // Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)


        // For Each item In list_cv

        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then
        // Return rtn
        // End If

        // Dim cvitem As String = item

        // 'コピー項目表示
        // Me.lblCVItem.Text = ""

        // '----- デバッグ用処理 ----- sta
        // 'Dim list As New List(Of String) From {"部屋情報-部屋特約情報"}
        // 'If list.Contains(cvitem) = False Then
        // '    cvitem = ""
        // 'End If
        // '----- デバッグ用処理 ----- end

        // Dim normalflg As Boolean = True
        // Dim qry As String = Nothing
        // Dim copyflg As Boolean = True

        // Select Case cvitem
        // Case "部屋情報-部屋駐車場情報"
        // obj_rep = New Njc.Repository.Hy_cyusyajo_mid_Repository
        // pgbcnt = pgbcnt + 1
        // Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        // Case "部屋情報-部屋特約情報"
        // obj_rep = New Njc.Repository.Hy_tokuyaku_mid_Repository
        // pgbcnt = pgbcnt + 1
        // Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        // Case "契約情報-契約契約者情報"
        // obj_rep = New Njc.Repository.Ky_kys_mid_Repository
        // pgbcnt = pgbcnt + 1
        // Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        // Case "契約情報-契約保証人情報"
        // obj_rep = New Njc.Repository.Ky_hosyonin_mid_Repository
        // pgbcnt = pgbcnt + 1
        // Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        // Case "契約情報-契約特約事項情報"
        // obj_rep = New Njc.Repository.Ky_tokuyaku_mid_Repository
        // pgbcnt = pgbcnt + 1
        // Me.lblCVItem.Text = cvitem.Replace("-", " ： ")
        // Case Else
        // copyflg = False
        // End Select

        // '----- ログ出力 -----                       '20161009 ログ出力処理追加 -add
        // Console.WriteLine(Now & " " & LOG_SYORIKOMK_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
        // tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(903, LOG_SYORIKOMK_MIDFILE, cvitem), False)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)

        // '項目毎に個別処理開始
        // If copyflg Then
        // '挿入処理
        // rtn = obj_rep.Read_BaseMidFile()
        // End If

        // 'コピー処理失敗時は処理を抜ける
        // If rtn = False Then
        // errstr = MSG_ERR_BASEMIDFILEREAD

        // '----- ログ出力 -----                   '20161009 ログ出力処理追加 -add
        // Console.WriteLine(Now & " " & LOG_SYORIKOMK_ERR_MIDONECOPY & "：" & LOG_SYORIKOMK_MIDFILE & "(" & cvitem & ")")
        // tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(904, LOG_SYORIKOMK_MIDFILE, cvitem), False)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmptmpstr, tmptmpcnt)
        // Return rtn
        // End If

        // '----- 全体用プログレスバー更新 -----
        // Call obj_pgb.pgbsettingTotal(pgbcnt)
        // Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        // '----- 中間用プログレスバー更新 -----
        // Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        // Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        // Me.Refresh()
        // Next

        // If pgbcnt < pgbtotalcnt_total Then
        // '----- 全体用プログレスバー更新(最終調整) -----
        // pgbcnt = pgbtotalcnt_total
        // Call obj_pgb.pgbsettingTotal(pgbcnt)
        // Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        // '----- 中間用プログレスバー更新(最終調整) -----
        // Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        // Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        // Me.Refresh()
        // End If

        // '返却
        // Return rtn

        // End Function
        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

        /// <summary>
        /// 既存用中間ファイルのヘッダー名をオブジェクトへ格納 '20160812 汎用コンバート対応
        /// '20160905 中間ファイルコピー処理改善 引数を追加
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="sheetname"></param>
        /// <remarks></remarks>
        private void Set_MidDataInfoToObj(string filename, string sheetname, bool getbasemidflg)
        {

            var readtbl = new DataTable();
            int reccnt = 0;
            string fldname = "";                                              // 20160928 レビュー指摘修正 変数宣言時に初期化処理を行う
            string fldvalue = "";
            string tmp_既存中間ファイル名 = "";
            string tmp_既存中間シート名 = "";
            string tmp_既存中間フィールド名 = "";
            string tmp_汎用中間ファイル名 = "";
            string tmp_汎用中間シート名 = "";
            string tmp_汎用中間フィールド名_大分類 = "";
            string tmp_汎用中間フィールド名_小分類 = "";
            string tmp_コピーフラグ = "";
            string tmp_データ型 = "";
            string tmp_CV用最小値 = "";
            string tmp_CV用最大値 = "";
            string tmp_デフォルト値 = "";
            string tmp_CV用キーNo = "";
            string tmp_CV用必須項目フラグ = "";
            string tmp_有無参照 = "";
            string tmp_移行対象フラグ = "";                                   // 20160926 選定した移行項目をプログラムへ反映する修正 -add


            // --------------------------------------------------
            // 中間ファイル情報取得
            // --------------------------------------------------
            // 20160905 中間ファイルコピー処理改善 -chg sta
            // Dim tmp_sql As String = " SELECT DISTINCT * FROM " & MIDFILEINFO_DBNAME
            string tmp_sql = "";
            string tmp_taisyofname = "";
            string tmp_taisyosname = "";
            string tmp_cvflg = "移行対象フラグ";                              // 20160929 汎用→既存コピー処理改善対応 -add
            if (getbasemidflg)
            {
                tmp_taisyofname = "汎用中間ファイル名";
                tmp_taisyosname = "汎用中間シート名";
            }
            else
            {
                tmp_taisyofname = "既存中間ファイル名";
                tmp_taisyosname = "既存中間シート名";
            }
            tmp_sql = tmp_sql + " SELECT DISTINCT * FROM " + MidFileInfoModule.MIDFILEINFO_DBNAME;
            tmp_sql = tmp_sql + " WHERE RTRIM(" + tmp_taisyosname + ") = '" + sheetname + "' ";
            tmp_sql = tmp_sql + " AND RTRIM(" + tmp_cvflg + ") = '●' ";             // 20160929 汎用→既存コピー処理改善対応 -add
            // 20160905 中間ファイルコピー処理改善 -chg end
            bool argnormalflg = true;
            reccnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, normalflg: ref argnormalflg);

            // --------------------------------------------------
            // 初期化
            // --------------------------------------------------
            // List_Existmidheader.Clear()                                             '20160812 汎用コンバート対応 -del
            CommonModule.Hash_ExistMiddatatype.Clear();
            CommonModule.Hash_ExistMiddatamin.Clear();
            CommonModule.Hash_ExistMiddatamax.Clear();
            CommonModule.Hash_ExistMiddatadef.Clear();
            CommonModule.Hash_ExistMiddatakey.Clear();
            CommonModule.List_ExistMiddatareq.Clear();
            CommonModule.Hash_ExistMiddataref.Clear();
            CommonModule.List_ExistMiddatacv.Clear();                                             // 20160926 選定した移行項目をプログラムへ反映する修正(開発用) -add
            // List_Basemidheader.Clear()                                              '20160812 汎用コンバート対応 -del
            CommonModule.Hash_BaseMiddatatype.Clear();
            CommonModule.Hash_BaseMiddatamin.Clear();
            CommonModule.Hash_BaseMiddatamax.Clear();
            CommonModule.Hash_BaseMiddatadef.Clear();
            CommonModule.Hash_BaseMiddatakey.Clear();
            CommonModule.List_BaseMiddatareq.Clear();
            CommonModule.Hash_BaseMiddataref.Clear();
            CommonModule.List_BaseMiddatacv.Clear();                                              // 20160926 選定した移行項目をプログラムへ反映する修正 -add
            CommonModule.Hash_BaseMidToExistMid_FS.Clear();                                       // 20160905 中間ファイルコピー処理改善 -add
            CommonModule.Hash_ExistMidToBaseMid_FS.Clear();                                       // 20160905 中間ファイルコピー処理改善 -add
            CommonModule.Hash_BaseMidToExistMid_SH.Clear();                                       // 20160905 中間ファイルコピー処理改善 -add
            CommonModule.Hash_ExistMidToBaseMid_SH.Clear();                                       // 20160905 中間ファイルコピー処理改善 -add

            // --------------------------------------------------
            // 読込開始
            // --------------------------------------------------
            for (int cntii = 0, loopTo = readtbl.Rows.Count - 1; cntii <= loopTo; cntii++)
            {

                // 中断処理
                Application.DoEvents();
                if (CommonModule.CancelFlg | CommonModule.MidChkCancelFlg)
                {
                    return;
                }

                for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                {

                    // 項目名取得
                    fldname = readtbl.Columns[cntjj].ColumnName;

                    // 登録値取得
                    fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]);

                    // 各項目値→変数格納
                    switch (fldname ?? "")
                    {
                        case "既存中間ファイル名":
                            {
                                tmp_既存中間ファイル名 = fldvalue.Trim();
                                break;
                            }
                        case "既存中間シート名":
                            {
                                tmp_既存中間シート名 = fldvalue.Trim();
                                break;
                            }
                        case "既存中間フィールド名":
                            {
                                tmp_既存中間フィールド名 = fldvalue.Trim();
                                break;
                            }
                        case "汎用中間ファイル名":
                            {
                                tmp_汎用中間ファイル名 = fldvalue.Trim();
                                break;
                            }
                        case "汎用中間シート名":
                            {
                                tmp_汎用中間シート名 = fldvalue.Trim();
                                break;
                            }
                        case "汎用中間フィールド名_大分類":
                            {
                                tmp_汎用中間フィールド名_大分類 = fldvalue.Trim();
                                break;
                            }
                        case "汎用中間フィールド名_小分類":
                            {
                                tmp_汎用中間フィールド名_小分類 = fldvalue.Trim();
                                break;
                            }
                        case "コピーフラグ":
                            {
                                tmp_コピーフラグ = fldvalue.Trim();
                                break;
                            }
                        case "データ型":
                            {
                                tmp_データ型 = fldvalue.Trim();
                                break;
                            }
                        case "CV用最小値":
                            {
                                tmp_CV用最小値 = fldvalue.Trim();
                                break;
                            }
                        case "CV用最大値":
                            {
                                tmp_CV用最大値 = fldvalue.Trim();
                                break;
                            }
                        case "デフォルト値":
                            {
                                tmp_デフォルト値 = fldvalue.Trim();
                                break;
                            }
                        case "CV用キーNo":
                            {
                                tmp_CV用キーNo = fldvalue.Trim();
                                break;
                            }
                        case "CV用必須項目フラグ":
                            {
                                tmp_CV用必須項目フラグ = fldvalue.Trim();
                                break;
                            }
                        case "有無参照":
                            {
                                tmp_有無参照 = fldvalue.Trim();
                                break;
                            }
                        case "移行対象フラグ":
                            {
                                tmp_移行対象フラグ = fldvalue.Trim();
                                break;
                            }
                    }
                }

                // --------------------------------------------------
                // 汎用用と既存用の紐付(ファイル名-シート名)                         '20160905 中間ファイルコピー処理改善 -add
                // --------------------------------------------------
                string tmp_basemiddatatotal = tmp_汎用中間ファイル名 + CommonModule.STR_SPLIT_1 + tmp_汎用中間シート名;
                string tmp_existmiddatatotal = tmp_既存中間ファイル名 + CommonModule.STR_SPLIT_1 + tmp_既存中間シート名;
                if ((tmp_basemiddatatotal ?? "") == CommonModule.STR_SPLIT_1)
                {
                    tmp_basemiddatatotal = "";
                }
                if ((tmp_existmiddatatotal ?? "") == CommonModule.STR_SPLIT_1)
                {
                    tmp_existmiddatatotal = "";
                }

                if (!string.IsNullOrEmpty(tmp_basemiddatatotal) & !string.IsNullOrEmpty(tmp_existmiddatatotal) & CommonModule.Hash_BaseMidToExistMid_FS.ContainsKey(tmp_basemiddatatotal) == false)
                {
                    CommonModule.Hash_BaseMidToExistMid_FS.Add(tmp_basemiddatatotal, tmp_existmiddatatotal);
                }
                if (!string.IsNullOrEmpty(tmp_existmiddatatotal) & !string.IsNullOrEmpty(tmp_basemiddatatotal) & CommonModule.Hash_ExistMidToBaseMid_FS.ContainsKey(tmp_existmiddatatotal) == false)
                {
                    CommonModule.Hash_ExistMidToBaseMid_FS.Add(tmp_existmiddatatotal, tmp_basemiddatatotal);
                }

                // --------------------------------------------------
                // 既存用中間ファイルと紐付けて各オブジェクトへ格納
                // --------------------------------------------------
                string tmp_existmidheader = tmp_既存中間シート名 + CommonModule.STR_SPLIT_1 + tmp_既存中間フィールド名;
                if ((tmp_existmidheader ?? "") == CommonModule.STR_SPLIT_1)
                {
                    tmp_existmidheader = "";
                }
                // 20160905 中間ファイルコピー処理改善 -del sta
                // '既存用中間ファイルのヘッダー
                // If tmp_existmidheader <> "" Then
                // List_Existmidheader.Add(tmp_existmidheader)
                // End If
                // 20160905 中間ファイルコピー処理改善 -del end
                // 既存用中間ファイルデータ型
                if (!string.IsNullOrEmpty(tmp_existmidheader))
                {
                    CommonModule.Hash_ExistMiddatatype.Add(tmp_existmidheader, tmp_データ型);
                }

                // 既存用中間ファイル最小値
                if (!string.IsNullOrEmpty(tmp_existmidheader))
                {
                    CommonModule.Hash_ExistMiddatamin.Add(tmp_existmidheader, tmp_CV用最小値);
                }

                // 既存用中間ファイル最大値
                if (!string.IsNullOrEmpty(tmp_existmidheader))
                {
                    CommonModule.Hash_ExistMiddatamax.Add(tmp_existmidheader, tmp_CV用最大値);
                }

                // 既存用中間ファイルデフォルト値
                if (!string.IsNullOrEmpty(tmp_existmidheader) & !string.IsNullOrEmpty(tmp_デフォルト値))
                {
                    CommonModule.Hash_ExistMiddatadef.Add(tmp_existmidheader, tmp_デフォルト値);
                }

                // 既存用中間ファイルキーNo
                if (!string.IsNullOrEmpty(tmp_existmidheader) & !string.IsNullOrEmpty(tmp_CV用キーNo))
                {
                    CommonModule.Hash_ExistMiddatakey.Add(tmp_existmidheader, tmp_CV用キーNo);
                }

                // 既存用中間ファイル必須項目フラグ
                if (!string.IsNullOrEmpty(tmp_existmidheader) & !string.IsNullOrEmpty(tmp_CV用必須項目フラグ))
                {
                    CommonModule.List_ExistMiddatareq.Add(tmp_existmidheader);
                }

                // 既存用中間ファイル有無参照
                if (!string.IsNullOrEmpty(tmp_existmidheader) & !string.IsNullOrEmpty(tmp_有無参照))
                {
                    CommonModule.Hash_ExistMiddataref.Add(tmp_existmidheader, tmp_有無参照);
                }

                // 移行対象フラグ                                                    '20160926 選定した移行項目をプログラムへ反映する修正(開発用) -add
                if (!string.IsNullOrEmpty(tmp_existmidheader) & tmp_移行対象フラグ == "●")
                {
                    CommonModule.List_ExistMiddatacv.Add(tmp_existmidheader);
                }

                // --------------------------------------------------
                // 汎用用中間ファイルと紐付けて各オブジェクトへ格納
                // --------------------------------------------------
                string tmp_basemidheader = "";
                if (!string.IsNullOrEmpty(tmp_汎用中間フィールド名_小分類))
                {
                    // 20160913_2 部屋設備移行処理の追加 -chg sta
                    // tmp_basemidheader = tmp_汎用中間シート名 & STR_SPLIT_1 & tmp_汎用中間フィールド名_小分類
                    tmp_basemidheader = tmp_汎用中間シート名 + CommonModule.STR_SPLIT_1 + tmp_汎用中間フィールド名_大分類 + CommonModule.STR_SPLIT_1 + tmp_汎用中間フィールド名_小分類;
                }
                // 20160913_2 部屋設備移行処理の追加 -chg end
                else
                {
                    tmp_basemidheader = tmp_汎用中間シート名 + CommonModule.STR_SPLIT_1 + tmp_汎用中間フィールド名_大分類;
                }
                if ((tmp_basemidheader ?? "") == CommonModule.STR_SPLIT_1)
                {
                    tmp_basemidheader = "";
                }
                // 20160905 中間ファイルコピー処理改善 -del sta
                // '汎用用中間ファイルのヘッダー
                // If tmp_basemidheader <> "" And List_Basemidheader.Contains(tmp_basemidheader) = False Then
                // List_Basemidheader.Add(tmp_basemidheader)
                // End If
                // 20160905 中間ファイルコピー処理改善 -del end
                // 汎用用中間ファイルデータ型
                if (!string.IsNullOrEmpty(tmp_basemidheader) & CommonModule.Hash_BaseMiddatatype.ContainsKey(tmp_basemidheader) == false)
                {
                    CommonModule.Hash_BaseMiddatatype.Add(tmp_basemidheader, tmp_データ型);
                }

                // 汎用用中間ファイル最小値
                if (!string.IsNullOrEmpty(tmp_basemidheader) & CommonModule.Hash_BaseMiddatamin.ContainsKey(tmp_basemidheader) == false)
                {
                    CommonModule.Hash_BaseMiddatamin.Add(tmp_basemidheader, tmp_CV用最小値);
                }

                // 汎用用中間ファイル最大値
                if (!string.IsNullOrEmpty(tmp_basemidheader) & CommonModule.Hash_BaseMiddatamax.ContainsKey(tmp_basemidheader) == false)
                {
                    CommonModule.Hash_BaseMiddatamax.Add(tmp_basemidheader, tmp_CV用最大値);
                }

                // 汎用用中間ファイルデフォルト値
                if (!string.IsNullOrEmpty(tmp_basemidheader) & !string.IsNullOrEmpty(tmp_デフォルト値) & CommonModule.Hash_BaseMiddatadef.ContainsKey(tmp_basemidheader) == false)
                {
                    CommonModule.Hash_BaseMiddatadef.Add(tmp_basemidheader, tmp_デフォルト値);
                }

                // 汎用用中間ファイルキーNo
                if (!string.IsNullOrEmpty(tmp_basemidheader) & !string.IsNullOrEmpty(tmp_CV用キーNo) & CommonModule.Hash_BaseMiddatakey.ContainsKey(tmp_basemidheader) == false)
                {
                    CommonModule.Hash_BaseMiddatakey.Add(tmp_basemidheader, tmp_CV用キーNo);
                }

                // 汎用用中間ファイル必須項目フラグ
                if (!string.IsNullOrEmpty(tmp_basemidheader) & !string.IsNullOrEmpty(tmp_CV用必須項目フラグ) & CommonModule.List_BaseMiddatareq.Contains(tmp_basemidheader) == false)
                {
                    CommonModule.List_BaseMiddatareq.Add(tmp_basemidheader);
                }

                // 汎用用中間ファイル有無参照
                if (!string.IsNullOrEmpty(tmp_basemidheader) & !string.IsNullOrEmpty(tmp_有無参照) & CommonModule.Hash_BaseMiddataref.ContainsKey(tmp_basemidheader) == false)
                {
                    CommonModule.Hash_BaseMiddataref.Add(tmp_basemidheader, tmp_有無参照);
                }

                // 移行対象フラグ    '20160926 選定した移行項目をプログラムへ反映する修正 -add
                if (!string.IsNullOrEmpty(tmp_basemidheader) & tmp_移行対象フラグ == "●")
                {
                    CommonModule.List_BaseMiddatacv.Add(tmp_basemidheader);
                }

                // --------------------------------------------------
                // 汎用用と既存用の紐付(シート名-ヘッダー名)                         '20160905 中間ファイルコピー処理改善 -add
                // --------------------------------------------------
                if (!string.IsNullOrEmpty(tmp_existmidheader) & CommonModule.Hash_ExistMidToBaseMid_SH.ContainsKey(tmp_existmidheader) == false)
                {
                    CommonModule.Hash_ExistMidToBaseMid_SH.Add(tmp_existmidheader, tmp_basemidheader);
                }
                if (!string.IsNullOrEmpty(tmp_basemidheader) & CommonModule.Hash_BaseMidToExistMid_SH.ContainsKey(tmp_basemidheader) == false)
                {
                    CommonModule.Hash_BaseMidToExistMid_SH.Add(tmp_basemidheader, tmp_existmidheader);
                }

            }

        }

        /// <summary>
        /// 契約No、管理Noに「1」を自動設定する '20160929 汎用→既存コピー処理改善対応 -add
        /// </summary>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_ExistMidFile_DefCopyMain(ref string errstr)
        {

            bool rtn = true;
            var list_file = new List<string>() { "契約情報", "送金ルール情報", "物件情報", "部屋情報" };
            string qry = null;
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
            // Dim da As New OleDbDataAdapter()
            // Dim ds As DataSet = New DataSet()
            // Dim dt As New DataTable()
            // Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
            // Dim tmpextend As String = "Excel 8.0;HDR=YES;"

            // Dim con_read As New OleDbConnection()
            // Dim cmd_read As New OleDbCommand()
            // Dim con_write As New OleDbConnection()
            // Dim cmd_write As New OleDbCommand()
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
            int tmp_cnt = 0;

            // ----- 全体用・中間用プログレスバー更新 -----
            var obj_pgb = new ProgressBarManager();
            int pgbcnt = 0;
            int pgbtotalcnt_total = list_file.Count;
            obj_pgb.pgbInitTotal(pgbtotalcnt_total);
            obj_pgb.pgbInitChkTotal(pgbtotalcnt_total);
            int pgbcnt_part = 0;

            // 更新クエリ
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            // Dim tmpqry_ky As String()
            // qry = " UPDATE [契約基本情報$] SET [契約No] = 1"
            // qry = qry & "@#@" & " UPDATE [契約履歴情報$] SET [契約No] = 1,[契約管理レコードNo] = 1,[更新No] = 1,[改定No] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約契約者情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約入居者情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約保証人情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約車情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約保険情報$] SET [契約No] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約特約事項情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約メモ情報$] SET [契約NO] = 1,[契約管理レコードNo] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約入金項目情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            // qry = qry & "@#@" & " UPDATE [契約次回入金項目情報$] SET [契約No] = 1,[契約管理レコードNo] = 1 "
            // tmpqry_ky = Split(qry, "@#@")
            // Dim tmpqry_so As String()
            // qry = " UPDATE [送金ルール基本情報$] SET [送金ルール管理No] = 1 "
            // qry = qry & "@#@" & " UPDATE [送金ルール送金先情報$] SET [送金ルール管理No] = 1 "
            // qry = qry & "@#@" & " UPDATE [送金ルール入金項目情報$] SET [送金ルール管理No] = 1 "
            // qry = qry & "@#@" & " UPDATE [送金ルール控除項目情報$] SET [送金ルール管理No] = 1 "
            // tmpqry_so = Split(qry, "@#@")
            // Dim tmpqry_bk As String
            // tmpqry_bk = " UPDATE [物件所有者情報$] SET [管理No] = 1 "
            // Dim tmpqry_hy As String
            // tmpqry_hy = " UPDATE [部屋所有者情報$] SET [管理No] = 1 "
            string[] tmpqry_ky;
            qry = " UPDATE [CVTBL_契約基本情報] SET [契約No] = 1";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約履歴情報] SET [契約No] = 1,[契約管理レコードNo] = 1,[更新No] = 1,[改定No] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約契約者情報] SET [契約No] = 1,[契約管理レコードNo] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約入居者情報] SET [契約No] = 1,[契約管理レコードNo] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約保証人情報] SET [契約No] = 1,[契約管理レコードNo] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約車情報] SET [契約No] = 1,[契約管理レコードNo] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約保険情報] SET [契約No] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約特約事項情報] SET [契約No] = 1,[契約管理レコードNo] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約メモ情報] SET [契約NO] = 1,[契約管理レコードNo] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約入金項目情報] SET [契約No] = 1,[契約管理レコードNo] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_契約次回入金項目情報] SET [契約No] = 1,[契約管理レコードNo] = 1 ";
            tmpqry_ky = Strings.Split(qry, "@#@");
            string[] tmpqry_so;
            qry = " UPDATE [CVTBL_送金ルール基本情報] SET [送金ルール管理No] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_送金ルール送金先情報] SET [送金ルール管理No] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_送金ルール入金項目情報] SET [送金ルール管理No] = 1 ";
            qry = qry + "@#@" + " UPDATE [CVTBL_送金ルール控除項目情報] SET [送金ルール管理No] = 1 ";
            tmpqry_so = Strings.Split(qry, "@#@");
            string tmpqry_bk;
            tmpqry_bk = " UPDATE [CVTBL_物件所有者情報] SET [管理No] = 1 ";
            string tmpqry_hy;
            tmpqry_hy = " UPDATE [CVTBL_部屋所有者情報] SET [管理No] = 1 ";
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
            // 件数取得
            int pgbtotalcnt_part = tmpqry_ky.Count() + tmpqry_so.Count() + 1 + 1;
            obj_pgb.pgbInitPart(pgbtotalcnt_part);


            try
            {
                foreach (var filename in list_file)
                {

                    lblCVItem.Text = filename;                                            // コピー項目表示
                    string existmidpath = EtcMethod.Set_Path(CommonModule.MiddleDirPath, filename + ".xlsx");

                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                    // '--------------------------------------------------
                    // ' EXCEL接続 (ファイル書込用)
                    // '--------------------------------------------------
                    // '接続文字列生成
                    // con_write.ConnectionString = _
                    // "Provider=" & tmpprovider & _
                    // "Data Source=" & existmidpath & ";" & _
                    // "Extended Properties=" & """" & tmpextend & """"
                    // '接続設定
                    // cmd_write.Connection = con_write

                    // '--------------------------------------------------
                    // ' 読込→書込
                    // '--------------------------------------------------
                    // '接続オープン処理
                    // con_write.Open()
                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                    // クエリ実行
                    switch (filename ?? "")
                    {
                        case "契約情報":
                            {
                                for (int tmpcnt = 0, loopTo = Information.UBound(tmpqry_ky); tmpcnt <= loopTo; tmpcnt++)
                                {
                                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                    // cmd_write.CommandText = tmpqry_ky(tmpcnt)
                                    // cmd_write.ExecuteNonQuery()
                                    string tmp_qry = tmpqry_ky[tmpcnt];
                                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_qry, ref tmp_cnt);
                                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                                    // ----- 個別用プログレスバー更新 -----
                                    pgbcnt_part = pgbcnt_part + 1;
                                    obj_pgb.pgbsettingPart(pgbcnt_part);
                                    obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part);
                                    // 20161017 進捗表示処理による速度低下の修正 -del
                                    // Me.Refresh()
                                }

                                break;
                            }
                        case "送金ルール情報":
                            {
                                for (int tmpcnt = 0, loopTo1 = Information.UBound(tmpqry_so); tmpcnt <= loopTo1; tmpcnt++)
                                {
                                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                    // cmd_write.CommandText = tmpqry_so(tmpcnt)
                                    // cmd_write.ExecuteNonQuery()
                                    string tmp_qry = tmpqry_so[tmpcnt];
                                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_qry, ref tmp_cnt);
                                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                                    // ----- 個別用プログレスバー更新 -----
                                    pgbcnt_part = pgbcnt_part + 1;
                                    obj_pgb.pgbsettingPart(pgbcnt_part);
                                    obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part);
                                    // 20161017 進捗表示処理による速度低下の修正 -del
                                    // Me.Refresh()
                                }

                                break;
                            }
                        case "物件情報":
                            {
                                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                // cmd_write.CommandText = tmpqry_bk
                                // cmd_write.ExecuteNonQuery()
                                string tmp_qry = tmpqry_bk;
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_qry, ref tmp_cnt);
                                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                                // ----- 個別用プログレスバー更新 -----
                                pgbcnt_part = pgbcnt_part + 1;
                                obj_pgb.pgbsettingPart(pgbcnt_part);
                                obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part);
                                break;
                            }
                        // 20161017 進捗表示処理による速度低下の修正 -del
                        // Me.Refresh()
                        case "部屋情報":
                            {
                                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                                // cmd_write.CommandText = tmpqry_hy
                                // cmd_write.ExecuteNonQuery()
                                string tmp_qry = tmpqry_hy;
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_qry, ref tmp_cnt);
                                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                                // ----- 個別用プログレスバー更新 -----
                                pgbcnt_part = pgbcnt_part + 1;
                                obj_pgb.pgbsettingPart(pgbcnt_part);
                                // 20161017 進捗表示処理による速度低下の修正 -del
                                // Me.Refresh()
                                obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part);
                                break;
                            }

                        default:
                            {
                                break;
                            }

                    }

                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                    // '接続クローズ処理
                    // con_write.Close()
                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                    // ----- 全体用・中間用プログレスバー更新 -----
                    pgbcnt = pgbcnt + 1;
                    obj_pgb.pgbsettingTotal(pgbcnt);
                    obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, true);
                    obj_pgb.pgbsettingChkTotal(pgbcnt);
                    obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, true);
                    Refresh();
                }
            }

            catch (Exception ex)
            {

                // エラー処理
                errstr = CommonModule.MSG_ERR_BASEMIDFILEREAD;
                rtn = false;
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                // 'クローズ処理
                // con_write.Close()
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
            }

            // 20161026 移行項目画面表示修正 -add
            lblCVItem.Text = "";

            // 返却
            return rtn;

        }

        /// <summary>
        /// 一棟区分を所有情報から判断して自動設定する '20160929 汎用→既存コピー処理改善対応 -add
        /// '20161011 汎用→既存書込時の不要処理除去修正 引数に「list_filename」を追加
        /// </summary>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_ExistMidFile_SyoyuCopyMain(ref string errstr, List<string> list_filename)
        {

            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            // Dim rtn As Boolean = True
            // '20161011 一棟/区分一括調整処理改善 -chg sta
            // 'Dim list_ittobkno As New List(Of String)    
            // 'Dim list_kbnbkno As New List(Of String)         
            // Dim solist_kbnbkno As New SortedList(Of Integer, String)
            // '20161011 一棟/区分一括調整処理改善 -chg end

            // '20161011 汎用→既存書込時の不要処理除去修正 -add sta
            // '物件情報、部屋情報どちらも移行対象ではない場合は処理を抜ける
            // If list_filename.Contains("物件情報") = False And list_filename.Contains("部屋情報") = False Then
            // Return rtn
            // End If
            // '20161011 汎用→既存書込時の不要処理除去修正 -add end

            // '物件No取得
            // '20161011 一棟/区分一括調整処理改善 -chg sta
            // '最初に全て一棟に設定した後で区分のみ更新するように修正するため一棟用のオブジェクトを削除
            // 'rtn = Set_ExistMidFile_Syoyubkno(list_ittobkno, list_kbnbkno, errstr)
            // rtn = Set_ExistMidFile_Syoyubkno(solist_kbnbkno, errstr)
            // '20161011 一棟/区分一括調整処理改善 -chg end
            // If rtn Then
            // '取得した物件Noから一棟/区分を設定
            // '20161011 一棟/区分一括調整処理改善 -chg sta
            // '最初に全て一棟に設定した後で区分のみ更新するように修正するため一棟用のオブジェクトを削除
            // 'rtn = Set_ExistMidFile_Syoyukbn(list_ittobkno, list_kbnbkno, errstr)
            // rtn = Set_ExistMidFile_Syoyukbn(solist_kbnbkno, errstr)
            // '20161011 一棟/区分一括調整処理改善 -chg end
            // End If

            // Return rtn

            bool rtn = true;
            string tmp_sql = "";
            int tmp_cnt = 0;

            // 実行有無チェック
            if (list_filename.Contains("物件情報") == false & list_filename.Contains("部屋情報") == false)
            {
                return rtn;
            }

            try
            {

                // ----------------------------------
                // あらかじめ一棟所有へ更新しておく
                // ----------------------------------
                // 物件詳細情報
                tmp_sql = " UPDATE CVTBL_物件詳細情報 SET [所有者-一棟・所有区分] = 1 ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
                tmp_sql = "";

                // 送金ルール基本情報
                tmp_sql = " UPDATE CVTBL_送金ルール基本情報 SET [一所有形態区分-棟/区分] = 1 ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
                tmp_sql = "";

                // ----------------------------------
                // 区分所有を更新
                // ----------------------------------
                // 物件詳細情報
                tmp_sql = tmp_sql + " UPDATE CVTBL_物件詳細情報 SET [所有者-一棟・所有区分] = 2 ";
                tmp_sql = tmp_sql + " WHERE [物件NO] IN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT [物件No] FROM CVTBL_部屋所有者情報 ";
                tmp_sql = tmp_sql + " 	) ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
                tmp_sql = "";

                // 送金ルール基本情報
                tmp_sql = tmp_sql + " UPDATE CVTBL_送金ルール基本情報 SET [一所有形態区分-棟/区分] = 2 ";
                tmp_sql = tmp_sql + " WHERE [物件No] IN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT [物件No] FROM CVTBL_部屋所有者情報 ";
                tmp_sql = tmp_sql + " 	) ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
                tmp_sql = "";
            }

            catch (Exception ex)
            {

                errstr = CommonModule.MSG_ERR_BASEMIDFILEREAD;
                rtn = false;

            }

            return rtn;
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
        }

        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
        // ''' <summary>
        // ''' 中間ファイルの物件/部屋所有者情報から物件Noを取得してオブジェクトへ格納 '20160929 汎用→既存コピー処理改善対応
        // ''' '20161011 一棟/区分一括調整処理改善
        // ''' 引数「ByRef list_ittobkno As List(Of String)」を削除
        // ''' 引数「ByRef list_kbnbkno As List(Of String)」を「ByRef solist_kbnbkno As SortedList(Of Integer, String)」へ変更
        // ''' →最初に全て一棟に設定した後で区分のみ更新するように修正する
        // ''' </summary>
        // ''' <param name="solist_kbnbkno"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_ExistMidFile_Syoyubkno(ByRef solist_kbnbkno As SortedList(Of Integer, String), ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True
        // Dim da As New OleDbDataAdapter()
        // Dim ds As DataSet = New DataSet()
        // Dim dt As New DataTable()
        // Dim dr As OleDbDataReader
        // Dim qry As String = Nothing

        // Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
        // Dim tmpextend As String = "Excel 8.0;HDR=YES;"
        // '20161011 一棟/区分一括調整処理改善 -chg sta
        // 'Dim list_filename As New List(Of String) From {"物件情報", "部屋情報"}
        // Dim filename As String = "部屋情報"
        // Dim sheetname As String = "部屋所有者情報"
        // '20161011 一棟/区分一括調整処理改善 -chg end


        // '20161011 一棟/区分一括調整処理改善 -chg sta
        // ''物件/部屋所有者情報から対象物件Noを取得
        // 'For Each filename In list_filename

        // '    '中断処理
        // '    Application.DoEvents()
        // '    If CancelFlg Or MidChkCancelFlg Then
        // '        Return rtn
        // '    End If

        // '    'ファイルパス設定
        // '    Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")

        // '    '------------------
        // '    'メイン処理
        // '    '------------------

        // '    Dim con_read As New OleDbConnection()
        // '    Dim cmd_read As New OleDbCommand()
        // '    Dim con_write As New OleDbConnection()
        // '    Dim cmd_write As New OleDbCommand()

        // '    Try

        // '        '================================================================================
        // '        ' ●接続設定●
        // '        '================================================================================
        // '        '--------------------------------------------------
        // '        ' ①EXCEL接続 (ファイル読込用)
        // '        '--------------------------------------------------

        // '        '接続文字列生成
        // '        con_read.ConnectionString = _
        // '            "Provider=" & tmpprovider & _
        // '            "Data Source=" & existmidpath_read & ";" & _
        // '            "Extended Properties=" & """" & tmpextend & """"

        // '        '接続設定
        // '        cmd_read.Connection = con_read

        // '        '--------------------------------------------------
        // '        ' 読込
        // '        '--------------------------------------------------
        // '        Dim sheetname As String = ""
        // '        Select Case filename
        // '            Case "物件情報"
        // '                sheetname = "物件所有者情報"
        // '            Case "部屋情報"
        // '                sheetname = "部屋所有者情報"
        // '        End Select

        // '        '接続オープン処理
        // '        con_read.Open()
        // '        '20161011 一棟/区分一括調整処理改善 -chg sta
        // '        'qry = "SELECT [物件No] FROM [" & sheetname & "$] "            '20161009 不具合調査：クエリが複雑すぎる。DISTINCTできない？部屋の場合に同内容の大量のデータを格納することになる
        // '        qry = "SELECT DISTINCT [物件No] FROM [" & sheetname & "$] "
        // '        '20161011 一棟/区分一括調整処理改善 -chg end
        // '        cmd_read = con_read.CreateCommand
        // '        cmd_read.CommandText = qry
        // '        dt = New DataTable
        // '        dr = cmd_read.ExecuteReader
        // '        dt.Load(dr)

        // '        Dim colcnt As Integer = dt.Columns.Count
        // '        Dim readrowcnt As Integer = 1
        // '        Dim solist_header As New SortedList(Of Integer, String)
        // '        Dim tmp_list As New List(Of String)

        // '        '1行ずつ取得
        // '        Dim row As DataRow
        // '        For Each row In dt.Rows
        // '            tmp_list.Add(row(0).ToString)
        // '        Next

        // '        Select Case filename
        // '            Case "物件情報"
        // '                list_ittobkno = tmp_list
        // '            Case "部屋情報"
        // '                list_kbnbkno = tmp_list
        // '        End Select

        // '    Catch ex As Exception

        // '        'エラー処理
        // '        errstr = MSG_ERR_BASEMIDFILEREAD
        // '        rtn = False

        // '        'クローズ処理
        // '        con_read.Close()
        // '        con_write.Close()

        // '    End Try

        // '    'クローズ処理
        // '    con_read.Close()
        // '    con_write.Close()

        // 'Next

        // '中断処理
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then
        // Return rtn
        // End If

        // 'ファイルパス設定
        // Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")

        // '--------------------------------------------------
        // ' メイン処理
        // '--------------------------------------------------
        // Dim con_read As New OleDbConnection()
        // Dim cmd_read As New OleDbCommand()
        // Dim con_write As New OleDbConnection()
        // Dim cmd_write As New OleDbCommand()

        // Try
        // '--------------------------------------------------
        // ' ①EXCEL接続 (ファイル読込用)
        // '--------------------------------------------------
        // '接続文字列生成
        // con_read.ConnectionString = _
        // "Provider=" & tmpprovider & _
        // "Data Source=" & existmidpath_read & ";" & _
        // "Extended Properties=" & """" & tmpextend & """"
        // '接続設定
        // cmd_read.Connection = con_read

        // '--------------------------------------------------
        // ' 物件/部屋所有者情報から対象物件Noを取得
        // '--------------------------------------------------
        // con_read.Open()
        // '20161011 一棟/区分一括調整処理改善 -chg sta
        // 'qry = "SELECT [物件No] FROM [" & sheetname & "$] "                   '20161009 不具合調査：クエリが複雑すぎる為エラー(in連結の許容範囲オーバー)
        // qry = "SELECT DISTINCT [物件No] FROM [" & sheetname & "$] "
        // '20161011 一棟/区分一括調整処理改善 -chg end
        // cmd_read = con_read.CreateCommand
        // cmd_read.CommandText = qry
        // dt = New DataTable
        // dr = cmd_read.ExecuteReader
        // dt.Load(dr)

        // Dim colcnt As Integer = dt.Columns.Count
        // Dim readrowcnt As Integer = 1
        // Dim solist_header As New SortedList(Of Integer, String)
        // Dim tmp_list As New List(Of String)

        // '1行ずつ取得
        // Dim row As DataRow
        // Dim cnt As Integer = 0
        // Dim cnttotal As Integer = 0
        // Dim cnt_kugiri As Integer = 10000
        // Dim tmp_str As String = ""

        // For Each row In dt.Rows

        // 'カウンター更新
        // cnt = cnt + 1
        // cnttotal = cnttotal + 1

        // 'sql用の文字列へ変換
        // tmp_str = tmp_str & "," & "'" & row(0).ToString & "'"

        // If cnttotal = dt.Rows.Count Or cnt = cnt_kugiri Then
        // 'sql用文字列成形
        // If tmp_str <> "" Then
        // tmp_str = tmp_str.Remove(0, 1)
        // End If

        // 'オブジェクトへ格納
        // solist_kbnbkno.Add(Math.Ceiling(cnttotal / cnt_kugiri), tmp_str)

        // 'カウンター、作業用文字列初期化
        // cnt = 0
        // tmp_str = ""
        // End If
        // Next

        // Catch ex As Exception

        // 'エラー処理
        // errstr = MSG_ERR_BASEMIDFILEREAD
        // rtn = False
        // End Try

        // 'クローズ処理
        // con_read.Close()
        // con_write.Close()
        // '20161011 一棟/区分一括調整処理改善 -chg end

        // '返却
        // Return rtn

        // End Function

        // ''' <summary>
        // ''' 所有者情報から取得した物件Noを元に一棟・区分を自動設定 '20160929 汎用→既存コピー処理改善対応
        // ''' 物件詳細情報-所有者-一棟・所有区分 と 送金ルール基本情報-一所有形態区分-棟/区分 の2フィールドへ設定する
        // ''' '20161011 一棟/区分一括調整処理改善
        // ''' 引数「ByVal list_ittobkno As List(Of String)」を削除
        // ''' 引数「ByVal list_kbnbkno As List(Of String)」を「ByVal solist_kbnbkno As SortedList(Of Integer, String)」へ変更
        // ''' →最初に全て一棟に設定した後で区分のみ更新するように修正する
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <param name="solist_kbnbkno"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_ExistMidFile_Syoyukbn(ByVal solist_kbnbkno As SortedList(Of Integer, String), ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True
        // Dim da As New OleDbDataAdapter()
        // Dim ds As DataSet = New DataSet()
        // Dim dt As New DataTable()
        // Dim dr As OleDbDataReader
        // Dim qry As String = Nothing

        // Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
        // Dim tmpextend As String = "Excel 8.0;HDR=YES;"

        // Dim filename As String() = New String() {"物件情報", "送金ルール情報"}
        // Dim sheetname As String() = New String() {"物件詳細情報", "送金ルール基本情報"}
        // Dim bknofldname As String() = New String() {"物件NO", "物件No"}
        // Dim copyfldname As String() = New String() {"所有者-一棟・所有区分", "一所有形態区分-棟/区分"}

        // '----- 全体用・中間用プログレスバー初期化 -----
        // Dim obj_pgb As New ProgressBarManager
        // Dim pgbcnt As Integer = 0
        // Dim pgbtotalcnt_total As Integer = 2
        // Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
        // Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)


        // '20161011 一棟/区分一括調整処理改善 -del sta
        // '更新対象を取得する
        // 'Dim sql_ittobkno As String = ""
        // 'Dim sql_kbnbkno As String = ""
        // 'If list_ittobkno.Count <> 0 Then                         '20161009 不具合調査：クエリが複雑すぎる。分割して処理するなどを考える
        // '    For Each ittobkno In list_ittobkno
        // '        sql_ittobkno = sql_ittobkno & "," & "'" & ittobkno & "'"
        // '    Next
        // '    sql_ittobkno = sql_ittobkno.Remove(0, 1)
        // 'End If
        // 'If list_kbnbkno.Count <> 0 Then                         '20161009 不具合調査：クエリが複雑すぎる。分割して処理するなどを考える
        // '    For Each kbnbkno In list_kbnbkno
        // '        sql_kbnbkno = sql_kbnbkno & "," & "'" & kbnbkno & "'"
        // '    Next
        // '    sql_kbnbkno = sql_kbnbkno.Remove(0, 1)
        // 'End If
        // '20161011 一棟/区分一括調整処理改善 -del end

        // '--------------------------------------------------
        // ' 物件/部屋所有者情報から対象物件Noを取得
        // '--------------------------------------------------
        // For cntii = 0 To 1

        // '中断処理
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then
        // Return rtn
        // End If

        // 'ファイルパス設定
        // Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename(cntii) & ".xlsx")

        // '--------------------------------------------------
        // ' メイン処理
        // '--------------------------------------------------
        // Dim con_write As New OleDbConnection()
        // Dim cmd_write As New OleDbCommand()

        // Try
        // '--------------------------------------------------
        // ' EXCEL接続 (ファイル書込用)
        // '--------------------------------------------------
        // '接続文字列生成
        // con_write.ConnectionString = _
        // "Provider=" & tmpprovider & _
        // "Data Source=" & existmidpath_read & ";" & _
        // "Extended Properties=" & """" & tmpextend & """"
        // '接続設定
        // cmd_write.Connection = con_write

        // '--------------------------------------------------
        // ' 読込
        // '--------------------------------------------------
        // '接続オープン処理
        // con_write.Open()

        // '更新
        // '一度全てのデータを「1」に設定しておく
        // qry = ""
        // '20161011 一棟/区分一括調整処理改善 -chg sta
        // 'qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = 1 "
        // qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = '1' "
        // '20161011 一棟/区分一括調整処理改善 -chg end
        // cmd_write.CommandText = qry
        // cmd_write.ExecuteNonQuery()
        // '20161011 一棟/区分一括調整処理改善 -chg sta
        // ''一棟所有物件の更新
        // 'If sql_ittobkno <> "" Then
        // '    qry = ""
        // '    qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = 1 "
        // '    qry += " WHERE " & bknofldname(cntii) & " IN ( " & sql_ittobkno & ")"
        // '    cmd_write.CommandText = qry
        // '    cmd_write.ExecuteNonQuery()
        // 'End If

        // ''区分所有物件の更新
        // 'If sql_kbnbkno <> "" Then
        // '    qry = ""
        // '    qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = 2 "
        // '    qry += " WHERE " & bknofldname(cntii) & " IN ( " & sql_kbnbkno & ")"
        // '    cmd_write.CommandText = qry
        // '    cmd_write.ExecuteNonQuery()
        // 'End If

        // '区分所有物件の更新
        // '更新対象が存在しない場合は処理を抜ける
        // If solist_kbnbkno.Count = 0 Then
        // Exit For
        // End If

        // '----- 個別用プログレスバー更新 -----
        // Dim pgbcnt_part As Integer = 0
        // Dim pgbtotalcnt_part As Integer = solist_kbnbkno.Count
        // Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

        // For Each kbnbkno In solist_kbnbkno
        // Dim tmp_sqltaisyo As String = kbnbkno.Value
        // qry = ""
        // qry += " UPDATE [" & sheetname(cntii) & "$] SET [" & copyfldname(cntii) & "] = 2 "
        // qry += " WHERE " & bknofldname(cntii) & " IN ( " & tmp_sqltaisyo & ")"
        // cmd_write.CommandText = qry
        // cmd_write.ExecuteNonQuery()

        // '----- 個別用プログレスバー更新 -----
        // pgbcnt_part = pgbcnt_part + 1
        // Call obj_pgb.pgbsettingPart(pgbcnt_part)
        // Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
        // '20161017 進捗表示処理による速度低下の修正 -del
        // 'Me.Refresh()
        // Next
        // '20161011 一棟/区分一括調整処理改善 -chg end

        // Catch ex As Exception

        // 'エラー処理
        // errstr = MSG_ERR_BASEMIDFILEREAD
        // rtn = False
        // End Try

        // 'クローズ処理
        // con_write.Close()

        // '----- 全体用・中間用プログレスバー更新 -----
        // pgbcnt = pgbcnt + 1
        // Call obj_pgb.pgbsettingTotal(pgbcnt)
        // Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        // Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        // Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        // Me.Refresh()
        // Next

        // '返却
        // Return rtn

        // End Function
        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

        /// <summary>
        /// 物件部屋所有情報の所有開始終了年月を送金ルールの適用開始終了年月にコピーする '20161004 適用開始年月の自動設定 -add
        /// '20161011 汎用→既存書込時の不要処理除去修正 引数に「list_filename」を追加
        /// </summary>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_ExistMidFile_SoruleTekiyoYmdCopyMain(ref string errstr, List<string> list_filename)
        {

            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
            // Dim rtn As Boolean = True
            // Dim hash_ittosyoymd As New SafeDictionary<string, string>                '一棟所有の所有開始終了年月格納
            // Dim hash_kbnsyoymd As New SafeDictionary<string, string>                 '区分所有の所有開始終了年月格納

            // '20161011 汎用→既存書込時の不要処理除去修正 -add sta
            // '物件情報、部屋情報どちらも移行対象ではない場合は処理を抜ける
            // If list_filename.Contains("物件情報") = False And list_filename.Contains("部屋情報") = False Then
            // Return rtn
            // End If
            // '20161011 汎用→既存書込時の不要処理除去修正 -add end

            // '物件Noと所有開始終了年月を取得
            // rtn = Set_ExistMidFile_SyoyuYmdbkno(hash_ittosyoymd, hash_kbnsyoymd, errstr)

            // If rtn Then
            // '取得した情報から自動設定
            // rtn = Set_ExistMidFile_SyoyuYmd(hash_ittosyoymd, hash_kbnsyoymd, errstr)
            // Else
            // Return rtn
            // End If

            // Return rtn

            bool rtn = true;
            string tmp_sql = "";
            int tmp_cnt = 0;

            // 実行有無チェック
            if (list_filename.Contains("物件情報") == false & list_filename.Contains("部屋情報") == false)
            {
                return rtn;
            }

            try
            {

                // 一棟所有を更新
                tmp_sql = tmp_sql + " UPDATE CVTBL_送金ルール基本情報 SET ";
                tmp_sql = tmp_sql + " 	 [送金ルール適用開始日] = RTRIM(LTRIM(BKSYO.[所有期間開始])) ";
                tmp_sql = tmp_sql + " 	,[送金ルール適用終了日] = RTRIM(LTRIM(BKSYO.[所有期間終了])) ";
                tmp_sql = tmp_sql + " FROM CVTBL_送金ルール基本情報 AS SO ";
                tmp_sql = tmp_sql + " LEFT JOIN CVTBL_物件所有者情報 AS BKSYO ";
                tmp_sql = tmp_sql + " ON RTRIM(LTRIM(SO.[物件No])) = RTRIM(LTRIM(BKSYO.[物件No])) ";
                tmp_sql = tmp_sql + " WHERE BKSYO.[物件No] IS NOT NULL ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
                tmp_sql = "";

                // 区分所有を更新
                tmp_sql = tmp_sql + " UPDATE CVTBL_送金ルール基本情報 SET ";
                tmp_sql = tmp_sql + " 	 [送金ルール適用開始日] = RTRIM(LTRIM(HYSYO.[所有期間開始])) ";
                tmp_sql = tmp_sql + " 	,[送金ルール適用終了日] = RTRIM(LTRIM(HYSYO.[所有期間終了])) ";
                tmp_sql = tmp_sql + " FROM CVTBL_送金ルール基本情報 AS SO ";
                tmp_sql = tmp_sql + " LEFT JOIN CVTBL_部屋所有者情報 AS HYSYO ";
                tmp_sql = tmp_sql + " ON  RTRIM(LTRIM(SO.[物件No])) = RTRIM(LTRIM(HYSYO.[物件No])) ";
                tmp_sql = tmp_sql + " AND RTRIM(LTRIM(SO.[部屋No])) = RTRIM(LTRIM(HYSYO.[部屋No])) ";
                tmp_sql = tmp_sql + " WHERE HYSYO.[物件No] IS NOT NULL ";
                tmp_sql = tmp_sql + "  ";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
                tmp_sql = "";
            }

            catch (Exception ex)
            {

                errstr = CommonModule.MSG_ERR_BASEMIDFILEREAD;
                rtn = false;

            }

            return rtn;
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
        }

        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
        // ''' <summary>
        // ''' 中間ファイルの物件/部屋所有者情報から物件Noと所有開始終了年月を紐付けて取得してオブジェクトへ格納 '20161004 適用開始年月の自動設定
        // ''' </summary>
        // ''' <param name="hash_ittosyoymd"></param>
        // ''' <param name="hash_kbnsyoymd"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_ExistMidFile_SyoyuYmdbkno(ByRef hash_ittosyoymd As SafeDictionary<string, string>, ByRef hash_kbnsyoymd As SafeDictionary<string, string>, ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True
        // Dim da As New OleDbDataAdapter()
        // Dim ds As DataSet = New DataSet()
        // Dim dt As New DataTable()
        // Dim dr As OleDbDataReader
        // Dim qry As String = Nothing

        // Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
        // Dim tmpextend As String = "Excel 8.0;HDR=YES;"
        // Dim list_filename As New List(Of String) From {"物件情報", "部屋情報"}


        // '--------------------------------------------------
        // ' 物件/部屋所有者情報から対象物件Noを取得
        // '--------------------------------------------------
        // For Each filename In list_filename

        // '中断処理
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then
        // Return rtn
        // End If

        // 'ファイルパス設定
        // Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")

        // '--------------------------------------------------
        // ' メイン処理
        // '--------------------------------------------------
        // Dim con_read As New OleDbConnection()
        // Dim cmd_read As New OleDbCommand()
        // Dim con_write As New OleDbConnection()
        // Dim cmd_write As New OleDbCommand()

        // Try
        // '--------------------------------------------------
        // ' ①EXCEL接続 (ファイル読込用)
        // '--------------------------------------------------
        // '接続文字列生成
        // con_read.ConnectionString = _
        // "Provider=" & tmpprovider & _
        // "Data Source=" & existmidpath_read & ";" & _
        // "Extended Properties=" & """" & tmpextend & """"
        // '接続設定
        // cmd_read.Connection = con_read

        // '--------------------------------------------------
        // ' 読込
        // '--------------------------------------------------
        // '接続オープン処理
        // con_read.Open()

        // Dim sheetname As String = ""
        // Select Case filename
        // Case "物件情報"
        // sheetname = "物件所有者情報"
        // qry = "SELECT [物件No],[所有期間開始],[所有期間終了] FROM [" & sheetname & "$] "
        // Case "部屋情報"
        // sheetname = "部屋所有者情報"
        // qry = "SELECT [物件No],[部屋No],[所有期間開始],[所有期間終了] FROM [" & sheetname & "$] "
        // End Select

        // cmd_read = con_read.CreateCommand
        // cmd_read.CommandText = qry
        // dt = New DataTable
        // dr = cmd_read.ExecuteReader
        // dt.Load(dr)

        // '1行ずつ取得
        // Dim row As DataRow
        // For Each row In dt.Rows
        // '作業用変数
        // Dim tmp_bkno As String = ""
        // Dim tmp_hyno As String = ""
        // Dim tmp_syosta As String = ""
        // Dim tmp_syoend As String = ""

        // For cntjj = 0 To dt.Columns.Count - 1
        // Dim fldname As String = dt.Columns(cntjj).ColumnName
        // Dim fldvalue As String = row(cntjj).ToString.Trim
        // Select Case fldname
        // Case "物件No"
        // tmp_bkno = fldvalue
        // Case "部屋No"
        // tmp_hyno = fldvalue
        // Case "所有期間開始"
        // tmp_syosta = fldvalue
        // Case "所有期間終了"
        // tmp_syoend = fldvalue
        // End Select
        // Next

        // 'オブジェクトへ格納
        // Dim syostaendymd As String = tmp_syosta & STR_SPLIT_1 & tmp_syoend  'オブジェクト格納用

        // Select Case filename
        // Case "物件情報"
        // '一棟所有の所有期間開始年月
        // If tmp_bkno <> "" And hash_ittosyoymd.Contains(tmp_bkno) = False Then
        // hash_ittosyoymd.Add(tmp_bkno, syostaendymd)
        // End If
        // Case "部屋情報"
        // '区分所有の所有期間開始年月
        // Dim tmp_key As String = tmp_bkno & "-" & tmp_hyno
        // If tmp_key <> "" And hash_kbnsyoymd.Contains(tmp_key) = False Then
        // hash_kbnsyoymd.Add(tmp_key, syostaendymd)
        // End If
        // End Select

        // Next

        // Catch ex As Exception

        // 'エラー処理
        // errstr = MSG_ERR_BASEMIDFILEREAD
        // rtn = False
        // End Try

        // 'クローズ処理
        // con_read.Close()
        // con_write.Close()
        // Next

        // '返却
        // Return rtn

        // End Function

        // ''' <summary>
        // ''' 所有者情報から取得した物件No、適用開始終了年月を自動設定 '20161004 適用開始年月の自動設定
        // ''' </summary>
        // ''' <param name="hash_ittosyoymd"></param>
        // ''' <param name="hash_kbnsyoymd"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_ExistMidFile_SyoyuYmd(ByRef hash_ittosyoymd As SafeDictionary<string, string>, ByRef hash_kbnsyoymd As SafeDictionary<string, string>, ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True
        // Dim da As New OleDbDataAdapter()
        // Dim ds As DataSet = New DataSet()
        // Dim dt As New DataTable()
        // Dim qry As String = Nothing

        // Dim tmpprovider As String = "Microsoft.ACE.OLEDB.12.0; "                        'EXCEL2007以上(xlsx)
        // Dim tmpextend As String = "Excel 8.0;HDR=YES;"

        // Dim filename As String = "送金ルール情報"
        // Dim sheetname As String = "送金ルール基本情報"
        // Dim fldname_staymd As String = "送金ルール適用開始日"
        // Dim fldname_endymd As String = "送金ルール適用終了日"

        // '----- 全体用・中間用プログレスバー更新 -----
        // Dim obj_pgb As New ProgressBarManager
        // Dim pgbcnt As Integer = 0
        // Dim pgbtotalcnt_total As Integer = 2
        // Call obj_pgb.pgbInitTotal(pgbtotalcnt_total)
        // Call obj_pgb.pgbInitChkTotal(pgbtotalcnt_total)

        // 'ファイルパス設定
        // Dim existmidpath_read As String = EtcMethod.Set_Path(MiddleDirPath, filename & ".xlsx")


        // '--------------------------------------------------
        // ' EXCEL接続 (ファイル書込用)
        // '--------------------------------------------------
        // Dim con_write As New OleDbConnection()
        // Dim cmd_write As New OleDbCommand()
        // '接続文字列生成
        // con_write.ConnectionString = _
        // "Provider=" & tmpprovider & _
        // "Data Source=" & existmidpath_read & ";" & _
        // "Extended Properties=" & """" & tmpextend & """"
        // '接続設定
        // cmd_write.Connection = con_write
        // '接続オープン処理
        // con_write.Open()

        // '--------------------------------------------------
        // ' 一棟所有分を更新
        // '--------------------------------------------------
        // '----- 個別用プログレスバー初期化 -----
        // Dim pgbcnt_part As Integer = 0
        // Dim pgbtotalcnt_part As Integer = hash_ittosyoymd.Count
        // Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

        // For Each ittoitem In hash_ittosyoymd

        // '中断処理
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then
        // Return rtn
        // End If

        // Dim tmp_bkno As String = ittoitem.Key
        // Dim tmp_syoymd() As String = Split(ittoitem.Value.ToString, STR_SPLIT_1)
        // Dim tmp_staymd As String = tmp_syoymd(0)
        // Dim tmp_endymd As String = tmp_syoymd(1)

        // Try
        // '一棟所有物件の更新
        // If tmp_bkno <> "" Then
        // qry = ""
        // qry += " UPDATE [" & sheetname & "$] SET [" & fldname_staymd & "] = '" & tmp_staymd & "'" & "," & "[" & fldname_endymd & "] = '" & tmp_endymd & "'"
        // qry += " WHERE [物件No] = " & "'" & tmp_bkno & "'"
        // cmd_write.CommandText = qry
        // cmd_write.ExecuteNonQuery()
        // End If

        // Catch ex As Exception

        // 'エラー処理
        // errstr = MSG_ERR_BASEMIDFILEREAD
        // rtn = False
        // 'クローズ処理
        // con_write.Close()
        // '返却
        // Return rtn
        // End Try

        // '----- 個別用プログレスバー更新 -----
        // pgbcnt_part = pgbcnt_part + 1
        // Call obj_pgb.pgbsettingPart(pgbcnt_part)
        // Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
        // '20161017 進捗表示処理による速度低下の修正 -del
        // 'Me.Refresh()
        // Next

        // '----- 全体用・中間用プログレスバー更新 -----
        // pgbcnt = pgbcnt + 1
        // Call obj_pgb.pgbsettingTotal(pgbcnt)
        // Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        // Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        // Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        // Me.Refresh()

        // '--------------------------------------------------
        // ' 区分所有分を更新
        // '--------------------------------------------------
        // '----- 個別用プログレスバー初期化 -----
        // pgbcnt_part = 0
        // '20161018 プログレスバーカウントエラー修正 -chg sta
        // 'pgbtotalcnt_part = hash_ittosyoymd.Count
        // pgbtotalcnt_part = hash_kbnsyoymd.Count
        // '20161018 プログレスバーカウントエラー修正 -chg end
        // Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

        // For Each kbnitem In hash_kbnsyoymd

        // '中断処理
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then
        // Return rtn
        // End If

        // Dim tmp_bkhyno As String = kbnitem.Key
        // Dim tmp_syoymd() As String = Split(kbnitem.Value.ToString, STR_SPLIT_1)
        // Dim tmp_staymd As String = tmp_syoymd(0)
        // Dim tmp_endymd As String = tmp_syoymd(1)

        // Try
        // '一棟所有物件の更新
        // If tmp_bkhyno <> "-" Then
        // qry = ""
        // qry += " UPDATE [" & sheetname & "$] SET [" & fldname_staymd & "] = '" & tmp_staymd & "'" & "," & "[" & fldname_endymd & "] = '" & tmp_endymd & "'"
        // qry += " WHERE [物件No] + '-' + [部屋No] = " & "'" & tmp_bkhyno & "'"
        // cmd_write.CommandText = qry
        // cmd_write.ExecuteNonQuery()
        // End If

        // Catch ex As Exception

        // 'エラー処理
        // errstr = MSG_ERR_BASEMIDFILEREAD
        // rtn = False
        // 'クローズ処理
        // con_write.Close()
        // '返却
        // Return rtn
        // End Try

        // '----- 個別用プログレスバー更新 -----
        // pgbcnt_part = pgbcnt_part + 1
        // Call obj_pgb.pgbsettingPart(pgbcnt_part)
        // Call obj_com.ProgressOutPut(pgbcnt_part, pgbtotalcnt_part)
        // '20161017 進捗表示処理による速度低下の修正 -del
        // 'Me.Refresh()
        // Next

        // 'クローズ処理
        // con_write.Close()

        // '----- 全体用・中間用プログレスバー更新 -----
        // pgbcnt = pgbcnt + 1
        // Call obj_pgb.pgbsettingTotal(pgbcnt)
        // Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, True)
        // Call obj_pgb.pgbsettingChkTotal(pgbcnt)
        // Call obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, True)
        // Me.Refresh()

        // '返却
        // Return rtn

        // End Function
        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

        /// <summary>
        /// 中間ファイルから件数を取得して表示する 20160905 汎用コンバートの件数表示修正
        /// </summary>
        /// <remarks></remarks>
        private bool Set_BaseCVItemCnt()
        {

            var hash_lbltosheet = new SafeDictionary<string, string>();
            var hash_lbltocnt = new SafeDictionary<string, string>();
            // 20161012 中間ファイル件数表示速度改善 -del sta
            // Dim rtn As Boolean = True
            // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            // Dim maxrowcnt As Integer                                        '既存データの行数
            // Dim startrow As Integer = 0
            // Dim matchflg As Boolean = False
            // Dim flg As Boolean = True
            // Dim list_cvitem As New List(Of String) From _
            // {"バス交通マスタ", "学校区マスタ", "エリアマスタ", "保険種類マスタ", "特約マスタ", "変動費設定内容", _
            // "自社情報", "家主情報", "仲介・管理業者情報", "修繕業者情報", "ライフライン業者情報", "保険業者情報", "家賃保証業者情報", "施設保守業者情報", "施工業者情報", _
            // "物件情報", "部屋情報", "部屋設備情報", "契約者情報", "契約情報", "請求情報_未収", "請求情報_預り"}
            // 20161012 中間ファイル件数表示速度改善 -del sta

            // 20161012 中間ファイル件数表示速度改善 -add sta
            bool normalflg = true;                                 // 20161012 中間ファイル件数表示速度改善 -add
            var list_cvitem = new List<string>() { "バス交通マスタ", "学校区マスタ", "エリアマスタ", "保険種類マスタ", "特約マスタ", "自社情報", "家主情報", "仲介業者情報", "修繕業者情報", "ライフライン業者情報", "保険業者情報", "家賃保証業者情報", "施設保守業者情報", "施工業者情報", "物件情報", "部屋情報", "部屋設備情報", "契約者情報", "契約情報", "家主固定控除情報" };


            // 20161012 中間ファイル件数表示速度改善 -add end

            // 20161012 中間ファイル件数表示速度改善 -chg sta
            // flg = EtcMethod.Chk_FileExist(basemidfilepath)

            // If flg Then
            // middirpath = Path.GetDirectoryName(basemidfilepath)
            // Else
            // MsgResult = MessageBox.Show(MSG_ERR_MIDFILENOTEXIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            // Exit Sub
            // End If

            // 20161012 中間ファイル件数表示速度改善 -chg end

            // ----- 中間用プログレスバー初期化 -----
            var obj_pgb = new ProgressBarManager();
            int pgbcnt = 0;
            int pgbtotalcnt = list_cvitem.Count;
            obj_pgb.pgbInitChkTotal(pgbtotalcnt);
            lblCheckSituation.Text = CommonModule.SITUATION_EXTRACTION;
            lblPgbCheck.Text = "0 %";
            pgbCheck.Value = 0;
            pnlPrgChk.Visible = true;
            tabCtrlMain.Enabled = false;
            Refresh();    // 20161017 画面リフレッシュ機能の追加 -add

            // 20161012 中間ファイル件数表示速度改善 -add sta
            // オープン処理
            MainFrmHelper.init();

            // オープン処理失敗時は処理を抜ける
            if (normalflg == false)
            {
                // 20161104 中間ファイル読込エラー時の処理対応 -chg sta
                // MsgResult = MessageBox.Show(MSG_ERR_MIDFILENOTEXIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                // Exit Sub
                Close();
                return false;                    // 20161108 レビュー結果：SUBからFunctionへ変更した？
                // 20161108_2 レビュー結果戻り修正
                // すみません、SUBからFunctionへ変更しました。ヘッダー部へ記載しておりませんでした。
                // 20161104 中間ファイル読込エラー時の処理対応 -chg end
            }
            // 20161012 中間ファイル件数表示速度改善 -add end

            foreach (var basecvitem in list_cvitem)
            {

                // 20161012 中間ファイル件数表示速度改善 -chg sta
                // Dim tmp_lbl As New Label
                // Dim tmp_sheetname As String = ""

                // Select Case basecvitem
                // Case "バス交通マスタ" : hash_lbltosheet.Add(Me.lblKiMstBusCnt, "バス交通マスタ")
                // Case "学校区マスタ" : hash_lbltosheet.Add(Me.lblKiMstSchoolCnt, "学校区マスタ")
                // Case "エリアマスタ" : hash_lbltosheet.Add(Me.lblKiMstAreaCnt, "エリアマスタ")
                // Case "保険種類マスタ" : hash_lbltosheet.Add(Me.lblKiMstHokenruiCnt, "保険種類マスタ")
                // Case "特約マスタ" : hash_lbltosheet.Add(Me.lblKiMstTokuyakuCnt, "特約マスタ")
                // 'Case "変動費設定内容" : hash_lbltosheet.Add(Me.lblKiMstHendoCnt, "変動費設定内容")    '20160926 選定した移行項目をプログラムへ反映する修正 -del
                // Case "自社情報" : hash_lbltosheet.Add(Me.lblKiJisyaBaseCnt, "自社情報")
                // Case "家主情報" : hash_lbltosheet.Add(Me.lblKiOwBaseCnt, "家主情報")
                // Case "仲介・管理業者情報" : hash_lbltosheet.Add(Me.lblKiGyCyukaiBaseCnt, "仲介業者情報")
                // Case "修繕業者情報" : hash_lbltosheet.Add(Me.lblKiGySyuzenBaseCnt, "修繕業者情報")
                // Case "ライフライン業者情報" : hash_lbltosheet.Add(Me.lblKiGyLifelineBaseCnt, "ライフライン業者情報")
                // Case "保険業者情報" : hash_lbltosheet.Add(Me.lblKiGyHokenBaseCnt, "保険業者情報")
                // Case "家賃保証業者情報" : hash_lbltosheet.Add(Me.lblKiGyYatinhosyoBaseCnt, "家賃保証業者情報")
                // Case "施設保守業者情報" : hash_lbltosheet.Add(Me.lblKiGySisetuBaseCnt, "施設保守業者情報")
                // Case "施工業者情報" : hash_lbltosheet.Add(Me.lblKiGySekoBaseCnt, "施工業者情報")
                // Case "物件情報" : hash_lbltosheet.Add(Me.lblKiBkBaseCnt, "物件情報")
                // Case "部屋情報" : hash_lbltosheet.Add(Me.lblKiHyBaseCnt, "部屋情報")
                // Case "部屋設備情報" : hash_lbltosheet.Add(Me.lblKiHySetubiCnt, "部屋設備情報")
                // Case "契約者情報" : hash_lbltosheet.Add(Me.lblKiKysBaseCnt, "契約者情報")
                // Case "契約情報" : hash_lbltosheet.Add(Me.lblKiKyBaseCnt, "契約情報")
                // 'Case "請求情報_未収" : hash_lbltosheet.Add(Me.lblKiSqMiBaseCnt, "運用開始滞納金情報")  '20160926 選定した移行項目をプログラムへ反映する修正 -del
                // 'Case "請求情報_預り" : hash_lbltosheet.Add(Me.lblKiSqAzBaseCnt, "預り金情報")          '20160926 選定した移行項目をプログラムへ反映する修正 -del
                // End Select

                var tmp_lbl = new Label();
                switch (basecvitem ?? "")
                {
                    case "バス交通マスタ":
                        {
                            tmp_lbl = lblKiMstBusCnt;
                            break;
                        }
                    case "学校区マスタ":
                        {
                            tmp_lbl = lblKiMstSchoolCnt;
                            break;
                        }
                    case "エリアマスタ":
                        {
                            tmp_lbl = lblKiMstAreaCnt;
                            break;
                        }
                    case "保険種類マスタ":
                        {
                            tmp_lbl = lblKiMstHokenruiCnt;
                            break;
                        }
                    case "特約マスタ":
                        {
                            tmp_lbl = lblKiMstTokuyakuCnt;
                            break;
                        }
                    case "自社情報":
                        {
                            tmp_lbl = lblKiJisyaBaseCnt;
                            break;
                        }
                    case "家主情報":
                        {
                            tmp_lbl = lblKiOwBaseCnt;
                            break;
                        }
                    case "仲介業者情報":
                        {
                            tmp_lbl = lblKiGyCyukaiBaseCnt;
                            break;
                        }
                    case "修繕業者情報":
                        {
                            tmp_lbl = lblKiGySyuzenBaseCnt;
                            break;
                        }
                    case "ライフライン業者情報":
                        {
                            tmp_lbl = lblKiGyLifelineBaseCnt;
                            break;
                        }
                    case "保険業者情報":
                        {
                            tmp_lbl = lblKiGyHokenBaseCnt;
                            break;
                        }
                    case "家賃保証業者情報":
                        {
                            tmp_lbl = lblKiGyYatinhosyoBaseCnt;
                            break;
                        }
                    case "施設保守業者情報":
                        {
                            tmp_lbl = lblKiGySisetuBaseCnt;
                            break;
                        }
                    case "施工業者情報":
                        {
                            tmp_lbl = lblKiGySekoBaseCnt;
                            break;
                        }
                    case "物件情報":
                        {
                            tmp_lbl = lblKiBkBaseCnt;
                            break;
                        }
                    case "部屋情報":
                        {
                            tmp_lbl = lblKiHyBaseCnt;
                            break;
                        }
                    case "部屋設備情報":
                        {
                            tmp_lbl = lblKiHySetubiCnt;
                            break;
                        }
                    case "契約者情報":
                        {
                            tmp_lbl = lblKiKysBaseCnt;
                            break;
                        }
                    case "契約情報":
                        {
                            tmp_lbl = lblKiKyBaseCnt;
                            break;
                        }
                    case "家主固定控除情報":
                        {
                            tmp_lbl = lblKiSqOwKojoBaseCnt;
                            break;
                        }
                }

                // データを取得しオブジェクトへ格納
                List<object> dbSet = new List<object>();
                if (MainFrmHelper.dic汎用中間ファイル.ContainsKey(basecvitem))
                {
                    var dbSetName = MainFrmHelper.dic汎用中間ファイル[basecvitem].First().Value.Preテーブル.TableName;
                    if (!string.IsNullOrEmpty(dbSetName) && MainFrmHelper.CheckTableExists(MainFrmHelper.db, dbSetName))
                    {
                        dbSet = MainFrmHelper.GetDbSet(MainFrmHelper.db, dbSetName);
                    }
                }

                // 行数取得
                int rowcnt = dbSet.Count;

                // 件数表示
                if (rowcnt != 0)
                {
                    tmp_lbl.Text = int.Parse(rowcnt.ToString()).ToString("#,0") + " 件";
                }
                else
                {
                    tmp_lbl.Text = "なし";
                }

                // 20161012 中間ファイル件数表示速度改善 -chg end

                // ----- 中間用プログレスバー更新 -----
                pgbcnt = pgbcnt + 1;
                obj_pgb.pgbsettingChkTotal(pgbcnt);
                obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt, true);
                Refresh();
            }

            // クローズ処理 '20161012 中間ファイル件数表示速度改善 -add

            // 20161012 中間ファイル件数表示速度改善 -del sta
            // 'Excelファイル初期設定                          
            // rtn = excelfile.Set_ExcelFile_ReadOpen_BaseMidCnt(appli, wbook, wsheet, startrow, maxrowcnt, middirpath, CV_FROM_MIDDLE, hash_lbltosheet, hash_lbltocnt)

            // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
            // If rtn = False Then
            // Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
            // MsgResult = MessageBox.Show(MSG_ERR_MIDFILENOTEXIST, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            // Exit Sub
            // End If

            // '件数表示
            // For Each baseitem In hash_lbltocnt
            // Dim tmp_lbl As Label = baseitem.Key
            // Dim cnt As Integer = baseitem.Value
            // If cnt <> 0 Then
            // tmp_lbl.Text = Int32.Parse(cnt).ToString("#,0") & " 件"
            // Else
            // tmp_lbl.Text = "なし"
            // End If

            // Next

            // 'Excelファイル終了設定
            // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
            // 20161012 中間ファイル件数表示速度改善 -del end

            // ----- 中間用プログレスバー表示設定 -----
            pnlPrgChk.Visible = false;
            tabCtrlMain.Enabled = true;

            // 20161104 中間ファイル読込エラー時の処理対応 -add
            return true;

        }

        /// <summary>
        /// 既存用→汎用用中間ファイルへコピーする処理(開発用) 20160912 既存用→汎用用中間ファイルへのコピー処理
        /// (開発用処理)
        /// </summary>
        /// <param name="errstr"></param>
        /// <param name="list_cv"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_ExistMidFile_Copy(ref string errstr, List<string> list_cv)
        {

            bool rtn = true;
            bool normalflg = true;
            var list_existmidfileinfo = new List<string>();


            foreach (var cvitem_tmp in list_cv)
            {
                var cvitem = cvitem_tmp;
                // コンバート対象項目からファイル名とシート名を取得
                cvitem = cvitem.Replace("-", CommonModule.STR_SPLIT_1);
                string[] tmp_existmidstr = Strings.Split(cvitem, CommonModule.STR_SPLIT_1);
                string existfilename = tmp_existmidstr[0];
                string existsheetname = tmp_existmidstr[1];

                // 照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
                Set_MidDataInfoToObj(existfilename, existsheetname, false);

                // 移行項目からオープンする汎用用中間ファイル名を取得
                string basemidfsname = "";

                if (CommonModule.Hash_ExistMidToBaseMid_FS.ContainsKey(cvitem))
                {
                    basemidfsname = Conversions.ToString(CommonModule.Hash_ExistMidToBaseMid_FS[cvitem]);

                    // ----- デバッグ用処理 ----- sta
                    // Dim list As New List(Of String) From {"物件情報@#@物件詳細情報"}
                    // If list.Contains(cvitem) = False Then
                    // basemidfsname = ""
                    // End If
                    // ----- デバッグ用処理 ----- end

                    if (!string.IsNullOrEmpty(basemidfsname))
                    {
                        string[] tmp_basemidstr = Strings.Split(basemidfsname, CommonModule.STR_SPLIT_1);
                        string basefilename = tmp_basemidstr[0];
                        string basesheetname = tmp_basemidstr[1];

                        // 既存用中間ファイルからデータ取得
                        var hash_middata = new SafeDictionary<string, string[,]>();
                        int writerowcnt = 0;
                        normalflg = Set_ExistMidDataToObj(existfilename, existsheetname, ref hash_middata, ref writerowcnt, basesheetname, ref errstr);

                        // 汎用用中間ファイルへコピー
                        if (normalflg)
                        {

                            // 開くファイル名とシート名を元に初期化フラグを設定(一度開いたシート名をオブジェクトへ退避させておく)
                            bool iniflg = false;
                            if (list_existmidfileinfo.Contains(basefilename + CommonModule.STR_SPLIT_1 + basesheetname) == false)
                            {
                                list_existmidfileinfo.Add(basefilename + CommonModule.STR_SPLIT_1 + basesheetname);
                                iniflg = true;
                            }

                            // コピー処理
                            Set_ExistMidDataToBaseMidFile(basefilename, basesheetname, hash_middata, writerowcnt, iniflg);
                        }
                    }
                }
            }

            // 返却
            return rtn;

        }

        /// <summary>
        /// 既存用中間ファイルからデータを取得してオブジェクトへ格納 20160912 既存用→汎用用中間ファイルへのコピー処理
        /// (開発用処理)
        /// </summary>
        /// <param name="existmidfilename"></param>
        /// <param name="existmidsheetname"></param>
        /// <param name="middata"></param>
        /// <param name="hash_middata"></param>
        /// <param name="writerowcnt"></param>
        /// <param name="basesheetname"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_ExistMidDataToObj(string existmidfilename, string existmidsheetname, ref SafeDictionary<string, string[,]> hash_middata, ref int writerowcnt, string basesheetname, ref string errstr)





        {

            bool rtn = true;
            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            bool matchflg = false;
            string tmp_middirpath = txtExistMidToBaseMid.Text;

            // Excelファイル初期設定                          
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, tmp_middirpath, existmidfilename, existmidsheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (rtn == false)
            {
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // データが存在しない場合は処理を抜ける
            if (rowcnt == 0)
            {
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // ヘッダー取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

            // 対象ヘッダー検索
            for (int cntjj = 1, loopTo = columncnt; cntjj <= loopTo; cntjj++)
            {
                string fldname = headervalue[startrow - 1, cntjj].ToString().Trim();
                string keysfname = existmidsheetname + CommonModule.STR_SPLIT_1 + fldname;

                // 20160926 選定した移行項目をプログラムへ反映する修正(開発用) -chg sta
                // '既存用中間ファイルヘッダー存在有無確認
                // If List_Existmidheader.Contains(keysfname) Then

                // '紐付く汎用中間ファイルヘッダー有無確認
                // If Hash_ExistMidToBaseMid_SH.Contains(keysfname) Then
                // Dim tmp_existstr() As String = Split(Hash_ExistMidToBaseMid_SH(keysfname).ToString, STR_SPLIT_1)
                // If basesheetname = tmp_existstr(0) Then
                // middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
                // hash_middata.Add(keysfname, middata)
                // End If
                // End If

                // End If

                // 移行対象フラグ確認
                if (CommonModule.List_ExistMiddatacv.Contains(keysfname))
                {
                    // 既存用中間ファイルヘッダー存在有無確認
                    if (CommonModule.List_Existmidheader.Contains(keysfname))
                    {
                        // 紐付く汎用中間ファイルヘッダー有無確認
                        if (CommonModule.Hash_ExistMidToBaseMid_SH.ContainsKey(keysfname))
                        {
                            string[] tmp_existstr = Strings.Split(CommonModule.Hash_ExistMidToBaseMid_SH[keysfname].ToString(), CommonModule.STR_SPLIT_1);
                            if ((basesheetname ?? "") == (tmp_existstr[0] ?? ""))
                            {
                                var middata = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow, cntjj], wsheet.Cells[maxrowcnt, cntjj]]).Value;
                                hash_middata.Add(keysfname, middata);
                            }
                        }
                    }
                }
                // 20160926 選定した移行項目をプログラムへ反映する修正(開発用) -chg end
            }

            // 書込行数を返却
            writerowcnt = rowcnt;
            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
            // 返却
            return rtn;

        }

        /// <summary>
        /// 取得した既存用中間ファイルデータを汎用用中間ファイルへコピー '20160912 既存用→汎用用中間ファイルへのコピー処理
        /// (開発用処理)
        /// </summary>
        /// <param name="basemidfilename"></param>
        /// <param name="basemidsheetname"></param>
        /// <param name="hash_existmiddata"></param>
        /// <param name="writerowcnt"></param>
        /// <param name="iniflg"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Set_ExistMidDataToBaseMidFile(string basemidfilename, string basemidsheetname, SafeDictionary<string, string[,]> hash_existmiddata, int writerowcnt, bool iniflg)



        {

            bool rtn = true;
            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            int headerrow = 2;
            string tmp_middirpath = txtExistMidToBaseMid.Text;


            // **************************************************
            // 作業準備
            // **************************************************
            // Excelファイル初期設定(汎用用中間ファイルの初期化有無で処理を分岐させる)
            if (iniflg)
            {
                rtn = excelfile.Set_ExcelFile_WriteOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, tmp_middirpath, basemidfilename, basemidsheetname);
            }
            else
            {
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, tmp_middirpath, basemidfilename, basemidsheetname);
            }

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (rtn == false)
            {
                return rtn;
            }

            // 既存用中間ファイルが空の場合は処理を抜ける
            if (hash_existmiddata.Count == 0)
            {
                excelfile.Set_ExcelFile_WriteClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // 列数を退避
            int writecol = columncnt;

            // ヘッダー取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[headerrow, 1], wsheet.Cells[headerrow, columncnt]]).Value;

            // 既存用中間ファイルから取得したデータを元に貼付け先を取得
            foreach (var existmiddata in hash_existmiddata)
            {

                string existmidfldname = Conversions.ToString(existmiddata.Key);
                var existmidcoldata = existmiddata.Value;
                bool fldmatchflg = false;

                // 照合→コピー処理
                if (CommonModule.Hash_ExistMidToBaseMid_SH.ContainsKey(existmidfldname))
                {

                    string[] basemiddata = Strings.Split(CommonModule.Hash_ExistMidToBaseMid_SH[existmidfldname].ToString(), CommonModule.STR_SPLIT_1);
                    string basemidfldname = basemiddata[1];

                    // 20160915 開発用_貼付け処理改善 -chg sta
                    // '取得した列名を元にコピー処理
                    // '一致するフィールドが存在する場合
                    // For cntjj = 1 To columncnt
                    // Dim fldname As String = headervalue(1, cntjj).ToString.Trim
                    // If fldname = basemidfldname Then
                    // wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(startrow + writerowcnt - 1, cntjj)).Value = existmidcoldata
                    // fldmatchflg = True
                    // Exit For
                    // End If
                    // Next

                    // '一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
                    // If fldmatchflg = False Then
                    // writecol = writecol + 1
                    // wsheet.Cells(headerrow, writecol).Value = existmidfldname
                    // wsheet.Range(wsheet.Cells(startrow, writecol), wsheet.Cells(startrow + writerowcnt - 1, writecol)).Value = existmidcoldata
                    // End If
                    if (basemidfldname != "共有No")
                    {
                        // 取得した列名を元にコピー処理
                        // 一致するフィールドが存在する場合
                        for (int cntjj = 1, loopTo = columncnt; cntjj <= loopTo; cntjj++)
                        {
                            string fldname = headervalue[1, cntjj].ToString().Trim();
                            if ((fldname ?? "") == (basemidfldname ?? ""))
                            {
                                wsheet.get_Range(wsheet.Cells[startrow, cntjj], wsheet.Cells[startrow + writerowcnt - 1, cntjj]).set_Value(value: existmidcoldata);
                                fldmatchflg = true;
                                break;
                            }
                        }

                        // 20160926 選定した移行項目をプログラムへ反映する修正(開発用) -del sta
                        // 一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
                        // If fldmatchflg = False Then
                        // writecol = writecol + 1
                        // wsheet.Cells(headerrow, writecol).Value = existmidfldname
                        // wsheet.Range(wsheet.Cells(startrow, writecol), wsheet.Cells(startrow + writerowcnt - 1, writecol)).Value = existmidcoldata
                        // End If
                        // 20160926 選定した移行項目をプログラムへ反映する修正(開発用) -del end
                    }
                    // 20160915 開発用_貼付け処理改善 -chg end
                }
            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_WriteClose(ref appli, ref wbook, ref wsheet);
            // 返却
            return rtn;

        }

        /// <summary>
        /// 移行項目のファイル名を集約する
        /// </summary>
        /// <param name="list_cv"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private List<string> Set_FileNameIntensive(List<string> list_cv)
        {

            var rtn_list = new List<string>();

            // データが存在しない場合は処理を抜ける(念の為)
            if (list_cv.Count == 0)
            {
                return rtn_list;
            }

            // ファイル名を集約する
            foreach (var cvitem in list_cv)
            {
                string[] tmp_str = cvitem.ToString().Split('-');
                if (rtn_list.Contains(tmp_str[0]) == false)
                {
                    rtn_list.Add(tmp_str[0]);
                }
            }

            return rtn_list;

        }

        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add sta
        /// <summary>
        /// 汎用で移行対象となっている全項目名をオブジェクトへ格納 
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        private List<string> Get_BaseCvItem()
        {

            var rtn_list = new List<string>();

            rtn_list.Add("各マスタ情報-バス交通マスタ");
            rtn_list.Add("各マスタ情報-バス停マスタ");
            rtn_list.Add("各マスタ情報-学校区マスタ");
            rtn_list.Add("各マスタ情報-エリアマスタ");
            rtn_list.Add("各マスタ情報-保険種類マスタ");
            rtn_list.Add("各マスタ情報-特約マスタ");
            // 20161028 物件/部屋鍵取得方法修正 -add
            rtn_list.Add("各マスタ情報-鍵タイトルマスタ");
            rtn_list.Add("業者情報-仲介業者基本情報");
            rtn_list.Add("業者情報-仲介業者口座情報");
            rtn_list.Add("業者情報-仲介業者メモ情報");
            rtn_list.Add("業者情報-保険業者基本情報");
            rtn_list.Add("業者情報-保険業者口座情報");
            rtn_list.Add("業者情報-保険業者メモ情報");
            rtn_list.Add("業者情報-家賃保証業者基本情報");
            rtn_list.Add("業者情報-家賃保証業者メモ情報");
            rtn_list.Add("業者情報-修繕業者基本情報");
            rtn_list.Add("業者情報-修繕業者口座情報");
            rtn_list.Add("業者情報-修繕業者メモ情報");
            rtn_list.Add("業者情報-ライフライン業者情報");
            rtn_list.Add("業者情報-施工業者情報");
            rtn_list.Add("業者情報-施設保守業者情報");
            rtn_list.Add("自社情報-自社基本情報");
            rtn_list.Add("自社情報-自社口座情報");
            rtn_list.Add("自社情報-自社担当者情報");
            rtn_list.Add("自社情報-自社メモ情報");
            rtn_list.Add("自社情報-振込依頼人情報");
            rtn_list.Add("自社情報-口座振替情報");
            rtn_list.Add("自社情報-家賃入金口座情報");
            rtn_list.Add("家主情報-家主基本情報");
            rtn_list.Add("家主情報-家主口座情報");
            rtn_list.Add("家主情報-家主メモ情報");
            rtn_list.Add("契約者情報-契約者基本情報");
            rtn_list.Add("契約者情報-契約者口座情報");
            rtn_list.Add("契約者情報-契約者メモ情報");
            rtn_list.Add("契約者情報-契約者照合用カナ情報");
            rtn_list.Add("契約者情報-契約者保証人情報");
            rtn_list.Add("物件情報-物件基本情報");
            rtn_list.Add("物件情報-物件詳細情報");
            rtn_list.Add("物件情報-物件所有者情報");
            rtn_list.Add("物件情報-物件交通情報");
            rtn_list.Add("物件情報-物件メモ情報");
            rtn_list.Add("物件情報-物件鍵情報");
            rtn_list.Add("物件情報-物件近隣駐車場情報");
            rtn_list.Add("部屋情報-部屋基本情報");
            rtn_list.Add("部屋情報-部屋詳細情報");
            rtn_list.Add("部屋情報-部屋所有者情報");
            rtn_list.Add("部屋情報-部屋駐車場情報");
            rtn_list.Add("部屋情報-部屋特約情報");
            rtn_list.Add("部屋情報-部屋鍵情報");
            rtn_list.Add("部屋情報-部屋間取内訳情報");
            rtn_list.Add("部屋情報-部屋設備情報");
            rtn_list.Add("部屋情報-部屋入金項目情報");
            rtn_list.Add("部屋情報-部屋メモ情報");
            rtn_list.Add("部屋情報-部屋共通セールスポイント情報");
            rtn_list.Add("送金ルール情報-送金ルール基本情報");
            rtn_list.Add("送金ルール情報-送金ルール送金先情報");
            rtn_list.Add("送金ルール情報-送金ルール入金項目情報");
            rtn_list.Add("送金ルール情報-送金ルール控除項目情報");
            rtn_list.Add("契約情報-契約基本情報");
            rtn_list.Add("契約情報-契約履歴情報");
            rtn_list.Add("契約情報-契約車情報");
            rtn_list.Add("契約情報-契約契約者情報");
            rtn_list.Add("契約情報-契約保証人情報");
            rtn_list.Add("契約情報-契約入居者情報");
            rtn_list.Add("契約情報-契約特約事項情報");
            rtn_list.Add("契約情報-契約保険情報");
            rtn_list.Add("契約情報-契約メモ情報");
            rtn_list.Add("契約情報-契約入金項目情報");
            rtn_list.Add("契約情報-契約次回入金項目情報");
            rtn_list.Add("請求情報-家主固定控除情報");

            return rtn_list;

        }
        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add end

        // 20160929 汎用→既存コピー処理改善対応 -del sta
        // ''' <summary>
        // ''' 一棟区分を所有情報から判断して自動設定する '20160920 プログラム自動設定箇所の追加 -add
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidFile_SyoyuCopyMain(ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True
        // Dim list_ittobkno As New List(Of String)
        // Dim list_kbnbkno As New List(Of String)

        // '汎用用中間ファイルからデータを取得する
        // rtn = Set_BaseMidFile_Syoyubkno(list_ittobkno, list_kbnbkno)

        // If rtn = False Then
        // errstr = MSG_ERR_BASEMIDFILEREAD
        // Return rtn
        // Else
        // '取得したデータを元に一棟区分を設定する
        // rtn = Set_ExistMidFile_Syoyubkno(list_ittobkno, list_kbnbkno)

        // If rtn = False Then
        // errstr = MSG_ERR_EXISTMIDFILEWRITE
        // Return rtn
        // End If
        // End If

        // Return rtn

        // End Function

        // ''' <summary>
        // ''' 汎用用中間ファイルの物件/部屋所有者情報から物件Noを取得してオブジェクトへ格納 '20160920 プログラム自動設定箇所の追加
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <param name="list_ittobkno"></param>
        // ''' <param name="list_kbnbkno"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidFile_Syoyubkno(ByRef list_ittobkno As List(Of String), ByRef list_kbnbkno As List(Of String)) As Boolean

        // Dim rtn As Boolean = True
        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim obj_ittobkno As Object
        // Dim obj_kbnbkno As Object
        // Dim list_sheetname As New List(Of String) From {"物件所有者情報", "部屋所有者情報"}
        // '20160927 汎用用中間ファイル格納先変更時の不具合修正 -chg sta
        // 'Dim filepath As String = MiddleDirPath & "\" & CV_FROM_MIDDLE & ".xlsx"
        // Dim filepath As String = BaseMidDirPath & "\" & CV_FROM_MIDDLE & ".xlsx"
        // '20160927 汎用用中間ファイル格納先変更時の不具合修正 -chg end
        // Try

        // 'ファイル準備
        // appli = CreateObject("Excel.Application")
        // appli.Visible = False
        // wbook = appli.Workbooks.Open(filepath)

        // For Each sheetname In list_sheetname

        // '中断処理 '20160927 キャンセル処理を追加 -add
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // Return rtn
        // End If

        // 'シート選択
        // wsheet = wbook.Worksheets(sheetname)

        // '既存データの確認(ヘッダを含めた最大行の取得)
        // Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count
        // 'ヘッダの列数取得
        // Dim columncnt As Integer = wsheet.UsedRange.Columns.Count
        // '行数取得(全行 - 読込開始行 + 1)
        // Dim rowcnt As Integer = maxrowcnt - BASEMIDFILE_READWRITE_ROW + 1

        // '作業用変数
        // Dim tmp_obj As Object = Nothing

        // 'データが存在する場合はコピー処理を行う
        // If rowcnt <> 0 Then

        // 'ヘッダー取得
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(BASEMIDFILE_HEADER_ROW, 1), wsheet.Cells(BASEMIDFILE_HEADER_ROW, columncnt)).Value

        // 'ヘッダー照合
        // Dim tasiyoheader As String = "物件No"
        // Dim taisyohaederindex As Integer = 0

        // For cntjj = 1 To columncnt
        // Dim fldname As String = headervalue(1, cntjj).ToString.Trim
        // If fldname = tasiyoheader Then
        // taisyohaederindex = cntjj
        // tmp_obj = wsheet.Range(wsheet.Cells(BASEMIDFILE_READWRITE_ROW, taisyohaederindex), wsheet.Cells(maxrowcnt, taisyohaederindex)).Value
        // Exit For
        // End If
        // Next

        // End If

        // '取得した物件Noをリストオブジェクトへ格納
        // If tmp_obj IsNot Nothing Then

        // Dim tmp_list As New List(Of String)
        // For cntii = 1 To tmp_obj.Length
        // tmp_list.Add(tmp_obj(cntii, 1))
        // Next

        // Select Case sheetname
        // Case "物件所有者情報"
        // list_ittobkno = tmp_list
        // Case "部屋所有者情報"
        // list_kbnbkno = tmp_list
        // End Select

        // End If

        // Next

        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        // Catch ex As Exception

        // rtn = False
        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        // End Try

        // Return rtn

        // End Function

        // ''' <summary>
        // ''' 所有者情報から取得した物件Noを元に一棟・区分を自動設定 '20160920 プログラム自動設定箇所の追加
        // ''' 物件詳細情報-所有者-一棟・所有区分 と 送金ルール基本情報-一所有形態区分-棟/区分 の2フィールドへ設定する
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <param name="list_ittobkno"></param>
        // ''' <param name="list_kbnbkno"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_ExistMidFile_Syoyubkno(ByVal list_ittobkno As List(Of String), ByVal list_kbnbkno As List(Of String)) As Boolean

        // Dim rtn As Boolean = True
        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim startrow As Integer                                         '書込開始行
        // Dim columncnt As Integer                                        '列数
        // Dim maxrowcnt As Integer                                        '既存データの行数
        // Dim rowcnt As Integer                                           '書込行数
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim matchflg As Boolean = False

        // Dim filename As String() = New String() {"物件情報", "送金ルール情報"}
        // Dim sheetname As String() = New String() {"物件詳細情報", "送金ルール基本情報"}
        // Dim bknofldname As String() = New String() {"物件NO", "物件No"}
        // Dim copyfldname As String() = New String() {"所有者-一棟・所有区分", "一所有形態区分-棟/区分"}

        // For cntfile = 0 To UBound(filename)

        // 'Excelファイル初期設定                          
        // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename(cntfile), sheetname(cntfile))

        // '中断処理 '20160927 キャンセル処理を追加 -add
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // Return rtn
        // End If

        // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
        // If rtn = False Then
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // Return rtn
        // End If

        // 'データが存在しない場合は処理を抜ける
        // If rowcnt = 0 Then
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // Return rtn
        // End If

        // 'ヘッダー取得
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, 1), wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, columncnt)).Value

        // '作業用変数
        // Dim copycolindex As Integer = 0
        // Dim obj_bknocol As Object = Nothing

        // '対象ヘッダー検索
        // For cntjj = 1 To columncnt - 1

        // Dim fldname As String = headervalue(1, cntjj).ToString.Trim

        // '物件No列と一棟区分設定列から必要な情報を取得
        // Select Case fldname
        // Case bknofldname(cntfile)
        // obj_bknocol = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        // Case copyfldname(cntfile)
        // copycolindex = cntjj
        // End Select

        // Next

        // '取得した物件Noから一棟区分を設定
        // For cntkk = 1 To obj_bknocol.Length

        // Dim tmp_bkno As String = obj_bknocol(cntkk, 1)
        // If list_ittobkno.Contains(tmp_bkno) Then
        // wsheet.Cells(cntkk + EXISTMIDFILE_READWRITE_ROW - 1, copycolindex).Value = "1"
        // ElseIf list_kbnbkno.Contains(tmp_bkno) Then
        // wsheet.Cells(cntkk + EXISTMIDFILE_READWRITE_ROW - 1, copycolindex).Value = "2"
        // Else
        // wsheet.Cells(cntkk + EXISTMIDFILE_READWRITE_ROW - 1, copycolindex).Value = "1"
        // End If

        // Next

        // '終了処理
        // Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        // Next

        // Return rtn

        // End Function
        // 20160929 汎用→既存コピー処理改善対応 -del end

        // 20160929 汎用→既存コピー処理改善対応 -del sta
        // ''' <summary>
        // ''' 契約No、管理Noに「1」を自動設定する '20160920 プログラム自動設定箇所の追加 -add
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidFile_DefCopyMain(ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True
        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

        // Dim list_file As New List(Of String) From {"契約情報", "送金ルール情報", "物件情報", "部屋情報"}   '20160921 プログラム自動設定箇所の追加2 -add 
        // Dim list_sheet_ky As New List(Of String) From _
        // {"契約メモ情報", "契約基本情報", "契約契約者情報", "契約次回入金項目情報", "契約車情報", "契約特約事項情報", "契約入居者情報", _
        // "契約入金項目情報", "契約変動費各戸メーター情報", "契約保険情報", "契約保証人情報", "契約履歴情報"}
        // Dim list_sheet_so As New List(Of String) From _
        // {"送金ルール基本情報", "送金ルール送金先情報", "送金ルール入金項目情報", "送金ルール控除項目情報"}
        // Dim list_fldname_ky As New List(Of String) From _
        // {"契約NO", "契約No", "契約管理レコードNo", "更新No", "改定No"}
        // Dim fldname_so As String = "送金ルール管理No"

        // '20160921 プログラム自動設定箇所の追加2 -add sta
        // Dim sheetname_bksyo As String = "物件所有者情報"
        // Dim sheetname_hysyo As String = "部屋所有者情報"
        // Dim fldname_comsyo As String = "管理No"
        // '20160921 プログラム自動設定箇所の追加2 -add end

        // Try

        // For Each filename In list_file

        // '中断処理 '20160927 キャンセル処理を追加 -add
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        // Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
        // Return rtn
        // End If

        // Dim filepath As String = MiddleDirPath & "\" & filename & ".xlsx"

        // 'ファイル準備
        // appli = CreateObject("Excel.Application")
        // appli.Visible = False
        // wbook = appli.Workbooks.Open(filepath)

        // Select Case filename
        // Case "契約情報"

        // For Each sheetname In list_sheet_ky

        // 'シート選択
        // wsheet = wbook.Worksheets(sheetname)

        // '既存データの確認(ヘッダを含めた最大行の取得)
        // Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count
        // 'ヘッダの列数取得
        // Dim columncnt As Integer = wsheet.UsedRange.Columns.Count
        // '行数取得(全行 - 読込開始行 + 1)
        // Dim rowcnt As Integer = maxrowcnt - EXISTMIDFILE_READWRITE_ROW + 1

        // 'データが存在する場合はコピー処理を行う
        // If rowcnt <> 0 Then

        // 'ヘッダー取得
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, 1), wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, columncnt)).Value

        // 'ヘッダー照合
        // For Each header In list_fldname_ky

        // Dim tasiyoheader As String = header
        // Dim taisyohaederindex As Integer = 0

        // For cntjj = 1 To columncnt
        // Dim fldname As String = headervalue(1, cntjj).ToString.Trim
        // If fldname = tasiyoheader Then

        // taisyohaederindex = cntjj

        // '自動設定
        // wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW, taisyohaederindex), wsheet.Cells(maxrowcnt, taisyohaederindex)).Value = "1"

        // Exit For

        // End If
        // Next

        // Next

        // End If

        // Next

        // Case "送金ルール情報"

        // For Each sheetname In list_sheet_so

        // 'シート選択
        // wsheet = wbook.Worksheets(sheetname)

        // '既存データの確認(ヘッダを含めた最大行の取得)
        // Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count
        // 'ヘッダの列数取得
        // Dim columncnt As Integer = wsheet.UsedRange.Columns.Count
        // '行数取得(全行 - 読込開始行 + 1)
        // Dim rowcnt As Integer = maxrowcnt - EXISTMIDFILE_READWRITE_ROW + 1

        // 'データが存在する場合はコピー処理を行う
        // If rowcnt <> 0 Then

        // 'ヘッダー取得
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, 1), wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, columncnt)).Value

        // 'ヘッダー照合
        // Dim tasiyoheader As String = fldname_so
        // Dim taisyohaederindex As Integer = 0

        // For cntjj = 1 To columncnt
        // Dim fldname As String = headervalue(1, cntjj).ToString.Trim
        // If fldname = tasiyoheader Then
        // taisyohaederindex = cntjj
        // Exit For
        // End If
        // Next

        // '自動設定
        // wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW, taisyohaederindex), wsheet.Cells(maxrowcnt, taisyohaederindex)).Value = "1"

        // End If

        // Next

        // Case "物件情報", "部屋情報"     '20160921 プログラム自動設定箇所の追加2 -add

        // Dim tmp_sheetname As String = ""
        // Select Case filename
        // Case "物件情報"
        // tmp_sheetname = sheetname_bksyo
        // Case "部屋情報"
        // tmp_sheetname = sheetname_hysyo
        // End Select

        // 'シート選択
        // wsheet = wbook.Worksheets(tmp_sheetname)

        // '既存データの確認(ヘッダを含めた最大行の取得)
        // Dim maxrowcnt As Integer = wsheet.UsedRange.Rows.Count
        // 'ヘッダの列数取得
        // Dim columncnt As Integer = wsheet.UsedRange.Columns.Count
        // '行数取得(全行 - 読込開始行 + 1)
        // Dim rowcnt As Integer = maxrowcnt - EXISTMIDFILE_READWRITE_ROW + 1

        // 'データが存在する場合はコピー処理を行う
        // If rowcnt <> 0 Then

        // 'ヘッダー取得
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, 1), wsheet.Cells(EXISTMIDFILE_READWRITE_ROW - 1, columncnt)).Value

        // 'ヘッダー照合
        // Dim tasiyoheader As String = fldname_comsyo
        // Dim taisyohaederindex As Integer = 0

        // For cntjj = 1 To columncnt
        // Dim fldname As String = headervalue(1, cntjj).ToString.Trim
        // If fldname = tasiyoheader Then
        // taisyohaederindex = cntjj
        // Exit For
        // End If
        // Next

        // '自動設定
        // wsheet.Range(wsheet.Cells(EXISTMIDFILE_READWRITE_ROW, taisyohaederindex), wsheet.Cells(maxrowcnt, taisyohaederindex)).Value = "1"

        // End If

        // End Select

        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        // Next

        // Catch ex As Exception

        // rtn = False
        // errstr = MSG_ERR_EXISTMIDFILEWRITE
        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        // End Try

        // Return rtn

        // End Function
        // 20160929 汎用→既存コピー処理改善対応 -del end

        // 20160929 汎用→既存コピー処理改善対応 -del sta
        // ''' <summary>
        // ''' 汎用用中間ファイル→既存用中間ファイル個別処理 '20160812 汎用コンバート対応
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <param name="list_cv"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidFile_ChgCopyMain(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

        // Dim rtn As Boolean = True
        // Dim obj_rep As New Object
        // Dim model_basemiditem As New Object
        // Dim basemidfilename As String = ""
        // Dim basemidfilesheetname As String = ""
        // Dim existmidfilename As String = ""
        // Dim existmidfilesheetname As String = ""

        // For Each item In list_cv

        // '中断処理 '20160927 キャンセル処理を追加 -add
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        // Return rtn
        // End If

        // Dim cvitem As String = item

        // 'test
        // 'Dim list As New List(Of String) From {"部屋情報-部屋特約情報"}
        // 'If list.Contains(cvitem) = False Then
        // '    cvitem = ""
        // 'End If

        // Dim normalflg As Boolean = True     '20160913_2 エラー時の処理を追加する修正 -add

        // Select Case cvitem
        // Case "部屋情報-部屋駐車場情報"
        // obj_rep = New Njc.Repository.Hy_cyusyajo_mid_Repository
        // basemidfilename = "中間ファイル"
        // basemidfilesheetname = "部屋情報"
        // existmidfilename = "部屋情報"
        // existmidfilesheetname = "部屋駐車場情報"
        // Case "部屋情報-部屋特約情報"
        // obj_rep = New Njc.Repository.Hy_tokuyaku_mid_Repository
        // basemidfilename = "中間ファイル"
        // basemidfilesheetname = "部屋情報"
        // existmidfilename = "部屋情報"
        // existmidfilesheetname = "部屋特約情報"
        // Case "部屋情報-部屋契約解約確認事項情報"
        // obj_rep = New Njc.Repository.Hy_kykaikakuninjiko_mid_Repository
        // basemidfilename = "中間ファイル"
        // basemidfilesheetname = "部屋情報"
        // existmidfilename = "部屋情報"
        // existmidfilesheetname = "部屋契約解約確認事項情報"
        // Case "契約情報-契約契約者情報"
        // obj_rep = New Njc.Repository.Ky_kys_mid_Repository
        // basemidfilename = "中間ファイル"
        // basemidfilesheetname = "契約契約者保証人情報"
        // existmidfilename = "契約情報"
        // existmidfilesheetname = "契約契約者情報"
        // Case "契約情報-契約保証人情報"
        // obj_rep = New Njc.Repository.Ky_hosyonin_mid_Repository
        // basemidfilename = "中間ファイル"
        // basemidfilesheetname = "契約契約者保証人情報"
        // existmidfilename = "契約情報"
        // existmidfilesheetname = "契約保証人情報"

        // Case "契約情報-契約特約事項情報"
        // obj_rep = New Njc.Repository.Ky_tokuyaku_mid_Repository
        // basemidfilename = "中間ファイル"
        // basemidfilesheetname = "契約特約およびメモ情報"
        // existmidfilename = "契約情報"
        // existmidfilesheetname = "契約特約事項情報"

        // Case Else
        // obj_rep = Nothing
        // End Select

        // If obj_rep IsNot Nothing Then
        // '20160913_2 エラー時の処理を追加する修正 -chg sta
        // ''汎用用中間ファイル読込
        // 'obj_rep.Read_BaseMidFile(basemidfilename, basemidfilesheetname, model_basemiditem)

        // ''既存用中間ファイル書込
        // 'obj_rep.Set_ExistMidFile(existmidfilename, existmidfilesheetname, model_basemiditem)

        // '汎用用中間ファイル読込
        // normalflg = obj_rep.Read_BaseMidFile(basemidfilename, basemidfilesheetname, model_basemiditem)

        // If normalflg Then

        // '既存用中間ファイル書込
        // normalflg = obj_rep.Set_ExistMidFile(existmidfilename, existmidfilesheetname, model_basemiditem)

        // If normalflg = False Then
        // errstr = MSG_ERR_EXISTMIDFILEWRITE
        // rtn = False
        // Return rtn
        // End If
        // Else
        // errstr = MSG_ERR_BASEMIDFILEREAD
        // rtn = False
        // Return rtn
        // End If
        // '20160913_2 エラー時の処理を追加する修正 -chg end
        // End If

        // Next

        // Return rtn

        // End Function
        // 20160929 汎用→既存コピー処理改善対応 -del end

        // 20160929 汎用→既存コピー処理改善対応 -del sta
        // ''' <summary>
        // ''' 汎用用中間ファイルからデータを取得してオブジェクトへ格納 20160905 中間ファイルコピー処理改善
        // ''' 20160913_2 エラー時の処理を追加する修正 引数「errstr」を削除
        // ''' </summary>
        // ''' <param name="midfilename"></param>
        // ''' <param name="midsheetname"></param>
        // ''' <param name="middata"></param>
        // ''' <param name="hash_middata"></param>
        // ''' <param name="writerowcnt"></param>
        // ''' <param name="existsheetname"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidDataToObj(ByVal midfilename As String, ByVal midsheetname As String, ByRef middata As Object, ByRef hash_middata As SafeDictionary<string, string>, ByRef writerowcnt As Integer, _
        // ByVal existsheetname As String) As Boolean

        // Dim rtn As Boolean = True
        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim startrow As Integer                                         '書込開始行
        // Dim columncnt As Integer                                        '列数
        // Dim maxrowcnt As Integer                                        '既存データの行数
        // Dim rowcnt As Integer                                           '書込行数
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim matchflg As Boolean = False
        // Dim headerrow As Integer = 2                                    '20160912 中間ファイル作成に伴うプログラム修正 -add

        // 'Excelファイル初期設定
        // '20160927 汎用用中間ファイル格納先変更時の不具合修正 -chg sta
        // 'rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, midfilename, midsheetname)
        // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, BaseMidDirPath, midfilename, midsheetname)
        // '20160927 汎用用中間ファイル格納先変更時の不具合修正 -chg end
        // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
        // If rtn = False Then
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // Return rtn
        // End If

        // 'データが存在しない場合は処理を抜ける
        // If rowcnt = 0 Then
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160905 Excel終了処理の修正 -add
        // Return rtn
        // End If

        // 'ヘッダー取得
        // '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        // 'Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(headerrow, 2), wsheet.Cells(headerrow, columncnt)).Value
        // '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        // '対象ヘッダー検索
        // For cntjj = 1 To columncnt - 1

        // '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        // 'Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim

        // 'Dim keysfname As String = midsheetname & "-" & fldname

        // ''中間ファイルヘッダー存在有無確認
        // 'If List_Basemidheader.Contains(keysfname) Then

        // '    '紐付く既存中間ファイルヘッダー有無確認
        // '    If Hash_BaseMidToExistMid_SH.Contains(keysfname) Then
        // '        Dim tmp_existstr() As String = Hash_BaseMidToExistMid_SH(keysfname).ToString.Split("-")
        // '        If existsheetname = tmp_existstr(0) Then
        // '            middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        // '            hash_middata.Add(keysfname, middata)
        // '        End If
        // '    End If

        // 'End If
        // Dim fldname As String = headervalue(1, cntjj).ToString.Trim

        // If fldname <> "項目名" Then

        // Dim keysfname As String = midsheetname & STR_SPLIT_1 & fldname            '20160913 「@#@」を共通変数にして下さい(他の箇所も同様)

        // '20160926 選定した移行項目をプログラムへ反映する修正 -chg sta
        // ''中間ファイルヘッダー存在有無確認
        // 'If List_Basemidheader.Contains(keysfname) Then

        // '    '紐付く既存中間ファイルヘッダー有無確認
        // '    If Hash_BaseMidToExistMid_SH.Contains(keysfname) Then
        // '        Dim tmp_existstr() As String = Split(Hash_BaseMidToExistMid_SH(keysfname).ToString, STR_SPLIT_1)
        // '        If existsheetname = tmp_existstr(0) Then
        // '            '20160915 コピー処理のずれを修正 -chg sta
        // '            'middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        // '            middata = wsheet.Range(wsheet.Cells(startrow, cntjj + 1), wsheet.Cells(maxrowcnt, cntjj + 1)).Value
        // '            '20160915 コピー処理のずれを修正 -chg end
        // '            hash_middata.Add(keysfname, middata)
        // '        End If
        // '    End If

        // 'End If

        // '移行対象フラグ確認
        // If List_BaseMiddatacv.Contains(keysfname) Then

        // '中間ファイルヘッダー存在有無確認
        // If List_Basemidheader.Contains(keysfname) Then

        // '紐付く既存中間ファイルヘッダー有無確認
        // If Hash_BaseMidToExistMid_SH.Contains(keysfname) Then
        // Dim tmp_existstr() As String = Split(Hash_BaseMidToExistMid_SH(keysfname).ToString, STR_SPLIT_1)
        // If existsheetname = tmp_existstr(0) Then
        // '20160915 コピー処理のずれを修正 -chg sta
        // 'middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        // middata = wsheet.Range(wsheet.Cells(startrow, cntjj + 1), wsheet.Cells(maxrowcnt, cntjj + 1)).Value
        // '20160915 コピー処理のずれを修正 -chg end
        // hash_middata.Add(keysfname, middata)
        // End If
        // End If

        // End If

        // End If
        // '20160926 選定した移行項目をプログラムへ反映する修正 -chg end
        // End If
        // '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        // Next

        // '書込行数を返却
        // '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        // 'writerowcnt = maxrowcnt
        // writerowcnt = rowcnt
        // '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        // Return rtn

        // End Function

        // ''' <summary>
        // ''' 取得した汎用用中間ファイルデータを既存用中間ファイルへコピー '20160905 中間ファイルコピー処理改善
        // ''' </summary>
        // ''' <param name="existmidfilename"></param>
        // ''' <param name="existmidsheetname"></param>
        // ''' <param name="hash_basemiddata"></param>
        // ''' <param name="writerowcnt"></param>
        // ''' <param name="iniflg"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidDataToExistMidFile(ByVal existmidfilename As String, ByVal existmidsheetname As String, ByVal hash_basemiddata As SafeDictionary<string, string>, ByVal writerowcnt As Integer, ByVal iniflg As Boolean) As Boolean

        // '変数説明
        // 'basemidfsname		    	… 画面上(UIチェックボックス)の移行項目からオープンする汎用用中間ファイル名を取得
        // 'List_BaseMiddatacv		    … 移行項目選定対応：汎用用に選定した項目を格納
        // 'Hash_BaseMidToExistMid_FS	… 汎用と既存のファイルとシート結合データ(中間ファイル.物件情報-物件情報.物件基本情報)
        // 'Hash_ExistMidToBaseMid_FS	… 既存と汎用のファイルとシート結合データ(物件情報.物件基本情報-中間ファイル.物件情報)
        // 'Hash_BaseMidToExistMid_SH	… 汎用と既存のシートと項目名の結合データ(物件情報.物件No-物件基本情報.物件No)
        // '                               → 汎用用中間ファイルデータを作業用のオブジェクトへ格納 
        // 'Hash_ExistMidToBaseMid_SH	… 既存と汎用のシートと項目名の結合データ(物件基本情報.物件No-物件情報.物件No)
        // '                               → オブジェクトへ格納した汎用用中間ファイルデータを既存U用中間ファイルへ貼り付けるためのハッシュ

        // Dim rtn As Boolean = True
        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim startrow As Integer                                         '書込開始行
        // Dim columncnt As Integer                                        '列数
        // Dim maxrowcnt As Integer                                        '既存データの行数
        // Dim rowcnt As Integer                                           '書込行数
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim fldmatchflg As Boolean = False

        // '************************
        // '作業準備
        // '************************

        // 'Excelファイル初期設定(既存用中間ファイルの初期化有無で処理を分岐させる)
        // If iniflg Then
        // rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, existmidfilename, existmidsheetname)
        // Else
        // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, existmidfilename, existmidsheetname)
        // End If

        // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
        // If rtn = False Then                             '20160913 どこそこで「rtn」を返しているが、エラーの際、ソース上階層に移った時に何のエラーかが不明(エラーの際の処理が何もされていない？)
        // Return rtn
        // End If
        // '20160905 中間ファイルが空の場合の共通コピー処理修正 -add sta
        // If hash_basemiddata.Count = 0 Then
        // Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)
        // Return rtn
        // End If
        // '20160905 中間ファイルが空の場合の共通コピー処理修正 -add end

        // '列数を退避 '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -add
        // Dim writecol As Integer = columncnt

        // 'ヘッダー取得
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value
        // '20160926 選定した移行項目をプログラムへ反映する修正 -del sta
        // '貼付処理を修正するためコメントアウト
        // ''汎用用中間ファイルから取得したデータを元に貼付け先を取得
        // 'For Each basemiddata In hash_basemiddata

        // '    Dim basemidfldname As String = basemiddata.Key
        // '    Dim basemidcoldata As Object = basemiddata.Value

        // '    '照合→コピー処理
        // '    If Hash_BaseMidToExistMid_SH.Contains(basemidfldname) Then

        // '        Dim existmiddata() As String = Split(Hash_BaseMidToExistMid_SH(basemidfldname).ToString, STR_SPLIT_1)
        // '        Dim existmidfldname As String = existmiddata(1)

        // '        '取得した列名を元にコピー処理
        // '        '一致するフィールドが存在する場合
        // '        For cntjj = 1 To columncnt
        // '            Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
        // '            If fldname = existmidfldname Then
        // '                '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        // '                'wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(writerowcnt, cntjj)).Value = basemidcoldata
        // '                wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(startrow + writerowcnt - 1, cntjj)).Value = basemidcoldata     '20160913 ここで書き出さずに、オブジェクト配列とかに格納して最後に纏めて登録できない？
        // '                '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        // '                fldmatchflg = True
        // '                Exit For
        // '            End If
        // '        Next

        // '        '一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
        // '        If fldmatchflg = False Then
        // '            '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg sta
        // '            'wsheet.Cells(startrow - 1, columncnt + 1).Value = existmidfldname
        // '            'wsheet.Range(wsheet.Cells(startrow, columncnt + 1), wsheet.Cells(writerowcnt, columncnt + 1)).Value = basemidcoldata
        // '            writecol = writecol + 1
        // '            wsheet.Cells(startrow - 1, writecol).Value = existmidfldname
        // '            wsheet.Range(wsheet.Cells(startrow, writecol), wsheet.Cells(startrow + writerowcnt - 1, writecol)).Value = basemidcoldata
        // '            '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg end
        // '        End If

        // '    End If

        // 'Next
        // '20160926 選定した移行項目をプログラムへ反映する修正 -del end

        // '20160926 選定した移行項目をプログラムへ反映する修正 -add sta
        // For Each existmid In Hash_ExistMidToBaseMid_SH

        // '既存中間ファイルと汎用中間ファイルの紐付状況を取得
        // Dim existmidshname As String = existmid.Key
        // Dim basemidshname As String = existmid.Value

        // '汎用中間ファイルから取得したデータを照合し、一致する列データを取得する
        // If basemidshname <> "" Then

        // Dim basemidfldname As String = ""
        // Dim basemidcoldata As Object = Nothing

        // For Each basemiddata In hash_basemiddata

        // basemidfldname = basemiddata.Key
        // basemidcoldata = basemiddata.Value

        // If basemidfldname = basemidshname Then

        // '一致した列データが存在する場合は貼付け処理を行う
        // Dim existmiddata() As String = Split(existmidshname, STR_SPLIT_1)
        // Dim existmidfldname As String = existmiddata(1)

        // For cntjj = 1 To columncnt
        // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
        // If fldname = existmidfldname Then
        // wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(startrow + writerowcnt - 1, cntjj)).Value = basemidcoldata
        // fldmatchflg = True
        // Exit For
        // End If
        // Next

        // '一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
        // If fldmatchflg Then
        // Exit For
        // Else
        // writecol = writecol + 1
        // wsheet.Cells(startrow - 1, writecol).Value = existmidfldname
        // wsheet.Range(wsheet.Cells(startrow, writecol), wsheet.Cells(startrow + writerowcnt - 1, writecol)).Value = basemidcoldata
        // End If

        // End If

        // Next

        // End If

        // Next
        // '20160926 選定した移行項目をプログラムへ反映する修正 -add end

        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        // Return rtn

        // End Function
        // 20160929 汎用→既存コピー処理改善対応 -del end

        // 20160905 中間ファイルコピー処理改善 -del sta
        // ''' <summary>
        // ''' 汎用用中間ファイル→既存用中間ファイル通常処理 '20160812 汎用コンバート対応
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidFile_Copy(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

        // Dim rtn As Boolean = True
        // Dim readtbl As New DataTable
        // Dim reccnt As Integer
        // Dim fldname As String
        // Dim fldvalue As String
        // Dim tmp_basemidfilename As String = ""
        // Dim tmp_basemidfilesheetname As String = ""
        // Dim tmp_basemidfilefldnamemain As String = ""
        // Dim tmp_basemidfilefldnamesub As String = ""
        // Dim tmp_existmidfilename As String = ""
        // Dim tmp_existmidfilesheetname As String = ""
        // Dim tmp_existmidfilefldname As String = ""
        // Dim tmp_copyflg As String = ""
        // Dim list_existmidfileinfo As New List(Of String)
        // Dim basemidfilematchflg As Boolean = True

        // '中間ファイル紐付情報取得
        // Dim tmp_sql As String = " SELECT DISTINCT * FROM tmp_midfileinfo ORDER BY [汎用中間ファイル名],[汎用中間シート名] "
        // reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl)

        // '読込開始
        // For cntii As Integer = 0 To readtbl.Rows.Count - 1

        // '中断処理
        // Application.DoEvents()
        // If CancelFlg Then
        // Exit Function
        // End If

        // '既存用中間ファイル初期化フラグ
        // Dim iniflg As Boolean = True

        // For cntjj As Integer = 0 To readtbl.Columns.Count - 1

        // '項目名取得
        // fldname = readtbl.Columns(cntjj).ColumnName

        // '登録値取得
        // fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj))

        // '各項目値→変数格納
        // Select Case fldname
        // Case "既存中間ファイル名"
        // tmp_existmidfilename = fldvalue.Trim
        // Case "既存中間シート名"
        // tmp_existmidfilesheetname = fldvalue.Trim
        // Case "既存中間フィールド名"
        // tmp_existmidfilefldname = fldvalue.Trim
        // Case "汎用中間ファイル名"
        // tmp_basemidfilename = fldvalue.Trim
        // Case "汎用中間シート名"
        // tmp_basemidfilesheetname = fldvalue.Trim
        // Case "汎用中間フィールド名_大分類"
        // tmp_basemidfilefldnamemain = fldvalue.Trim
        // Case "汎用中間フィールド名_小分類"
        // tmp_basemidfilefldnamesub = fldvalue.Trim
        // Case "コピーフラグ"
        // tmp_copyflg = fldvalue.Trim
        // End Select

        // Next

        // '処理対象の場合はコピー処理を実行
        // Dim cvtaisyo As String = tmp_existmidfilename & "-" & tmp_existmidfilesheetname

        // If list_cv.Contains(cvtaisyo) Then

        // 'そのままコピーできるデータの処理
        // If tmp_copyflg = "●" Then

        // '開くファイル名とシート名を元に初期化フラグを設定(一度開いたシート名をオブジェクトへ退避させておく)
        // Dim existmidfileinfo As String = tmp_existmidfilename & "-" & tmp_existmidfilesheetname
        // If list_existmidfileinfo.Contains(existmidfileinfo) = False Then
        // iniflg = True
        // list_existmidfileinfo.Add(existmidfileinfo)
        // Else
        // iniflg = False
        // End If

        // '汎用用中間ファイルからデータ取得
        // Dim middata As New Object
        // basemidfilematchflg = Me.Set_BaseMidDataToObj(tmp_basemidfilename, tmp_basemidfilesheetname, tmp_basemidfilefldnamemain, middata, errstr)

        // '汎用用中間ファイルのヘッダーが一致しなかった場合(編集されている場合)は処理を抜ける
        // If basemidfilematchflg = False Then
        // errstr = "中間ファイルの形式が変更された可能性があるためコンバート処理を続行することができません。"
        // rtn = basemidfilematchflg
        // Return rtn
        // End If

        // '既存用中間ファイルへコピー
        // Call Me.Set_BaseMidDataToExistMidFile(tmp_existmidfilename, tmp_existmidfilesheetname, tmp_existmidfilefldname, middata, iniflg)

        // End If

        // End If

        // Next

        // Return rtn

        // End Function
        // 20160905 中間ファイルコピー処理改善 -del end

        // 20160905 中間ファイルコピー処理改善 -del sta
        // ''' <summary>
        // ''' 汎用用中間ファイルからデータを取得してオブジェクトへ格納 '20160812 汎用コンバート対応
        // ''' </summary>
        // ''' <param name="midfilename"></param>
        // ''' <param name="midsheetname"></param>
        // ''' <param name="midfldname"></param>
        // ''' <param name="middata"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidDataToObj(ByVal midfilename As String, ByVal midsheetname As String, ByVal midfldname As String, ByRef middata As Object, ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True
        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim startrow As Integer                                         '書込開始行
        // Dim columncnt As Integer                                        '列数
        // Dim maxrowcnt As Integer                                        '既存データの行数
        // Dim rowcnt As Integer                                           '書込行数
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim matchflg As Boolean = False

        // 'Excelファイル初期設定                          
        // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, midfilename, midsheetname)

        // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
        // If rtn = False Then
        // Return rtn
        // End If

        // 'ヘッダー取得
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

        // '対象ヘッダー検索
        // For cntjj = 1 To columncnt
        // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
        // If fldname = midfldname Then
        // middata = wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells(maxrowcnt, cntjj)).Value
        // matchflg = True
        // Exit For
        // End If
        // Next

        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        // '照合フラグをセット
        // If matchflg = False Then
        // rtn = matchflg
        // End If

        // Return rtn

        // End Function
        // 20160905 中間ファイルコピー処理改善 -del end

        // 20160905 中間ファイルコピー処理改善 -del sta
        // ''' <summary>
        // ''' 取得した汎用用中間ファイルデータを既存用中間ファイルへコピー '20160812 汎用コンバート対応
        // ''' </summary>
        // ''' <param name="existmidfilename"></param>
        // ''' <param name="existmidsheetname"></param>
        // ''' <param name="existmidfldname"></param>
        // ''' <param name="middata"></param>
        // ''' <param name="iniflg"></param>
        // ''' <param name="basemidfilename"></param>
        // ''' <param name="basemidsheetname"></param>
        // ''' <param name="basemidfldname"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidDataToExistMidFile(ByVal existmidfilename As String, ByVal existmidsheetname As String, ByVal existmidfldname As String, ByRef middata As Object, ByVal iniflg As Boolean) As Boolean

        // Dim rtn As Boolean = True
        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim startrow As Integer                                         '書込開始行
        // Dim columncnt As Integer                                        '列数
        // Dim maxrowcnt As Integer                                        '既存データの行数
        // Dim rowcnt As Integer                                           '書込行数
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim fldmatchflg As Boolean = False

        // '************************
        // '作業準備
        // '************************

        // 'Excelファイル初期設定(既存用中間ファイルの初期化有無で処理を分岐させる)
        // If iniflg Then
        // rtn = excelfile.Set_ExcelFile_WriteOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, MiddleDirPath, existmidfilename, existmidsheetname)
        // Else
        // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, existmidfilename, existmidsheetname)
        // End If

        // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
        // If rtn = False Then
        // Return rtn
        // End If

        // 'ヘッダー取得
        // Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

        // 'コピー件数設定
        // Dim writerowcnt As Integer = middata.Length

        // '対象ヘッダー検索→一致した列にコピー
        // For cntjj = 1 To columncnt
        // Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
        // If fldname = existmidfldname Then
        // wsheet.Range(wsheet.Cells(startrow, cntjj), wsheet.Cells((startrow - 1) + writerowcnt, cntjj)).Value = middata
        // fldmatchflg = True
        // Exit For
        // End If
        // Next

        // '一致するフィールドが存在しない(既存用中間ファイルが編集されている)場合は新規に列を作成する
        // If fldmatchflg = False Then
        // wsheet.Cells(startrow - 1, columncnt + 1).Value = existmidfldname
        // wsheet.Range(wsheet.Cells(startrow, columncnt + 1), wsheet.Cells((startrow - 1) + writerowcnt, columncnt + 1)).Value = middata
        // End If

        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_WriteClose(appli, wbook, wsheet)

        // Return rtn

        // End Function
        // 20160905 中間ファイルコピー処理改善 -del end

        // 20160929 汎用→既存コピー処理改善対応 -del sta
        // ''' <summary>
        // ''' 汎用用中間ファイル→既存用中間ファイル通常処理 '20160905 中間ファイルコピー処理改善
        // ''' </summary>
        // ''' <param name="errstr"></param>
        // ''' <param name="list_cv"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Set_BaseMidFile_Copy(ByRef errstr As String, ByVal list_cv As List(Of String)) As Boolean

        // Dim rtn As Boolean = True
        // Dim normalflg As Boolean = True
        // Dim list_existmidfileinfo As New List(Of String)

        // For Each cvitem In list_cv

        // ''そのままコピー可判別用
        // 'Dim copyflg As Boolean = False

        // '中断処理 '20160927 キャンセル処理を追加 -add
        // Application.DoEvents()
        // If CancelFlg Or MidChkCancelFlg Then   '20160929 データチェックのキャンセル処理を追加 Or MidChkCancelFlg の条件も追加 -chg
        // Return rtn
        // End If

        // 'コンバート対象項目からファイル名とシート名を取得
        // cvitem = cvitem.Replace("-", STR_SPLIT_1)
        // Dim tmp_existmidstr() As String = Split(cvitem, STR_SPLIT_1)
        // Dim existfilename As String = tmp_existmidstr(0)
        // Dim existsheetname As String = tmp_existmidstr(1)

        // '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
        // Call Me.Set_MidDataInfoToObj(existfilename, existsheetname, False)

        // '移行項目からオープンする汎用用中間ファイル名を取得
        // Dim basemidfsname As String = ""
        // If Hash_ExistMidToBaseMid_FS.Contains(cvitem) Then

        // basemidfsname = Hash_ExistMidToBaseMid_FS(cvitem)
        // 'copyflg = True

        // 'test
        // 'Dim list As New List(Of String) From {"物件情報@#@物件交通情報"}
        // 'If list.Contains(cvitem) = False Then
        // '    basemidfsname = ""
        // 'End If

        // If basemidfsname <> "" Then

        // Dim tmp_basemidstr() As String = Split(basemidfsname, STR_SPLIT_1)
        // Dim basefilename As String = tmp_basemidstr(0)
        // Dim basesheetname As String = tmp_basemidstr(1)

        // '汎用用中間ファイルからデータ取得
        // Dim middata As New Object
        // Dim hash_middata As New SafeDictionary<string, string>
        // Dim writerowcnt As Integer = 0
        // '20160913_2 エラー時の処理を追加する修正 -chg sta
        // 'normalflg = Me.Set_BaseMidDataToObj(basefilename, basesheetname, middata, hash_middata, writerowcnt, existsheetname, errstr)
        // normalflg = Me.Set_BaseMidDataToObj(basefilename, basesheetname, middata, hash_middata, writerowcnt, existsheetname)
        // '20160913_2 エラー時の処理を追加する修正 -chg end
        // '既存用中間ファイルへコピー
        // If normalflg Then

        // '開くファイル名とシート名を元に初期化フラグを設定(一度開いたシート名をオブジェクトへ退避させておく)
        // Dim iniflg As Boolean = False
        // If list_existmidfileinfo.Contains(cvitem) = False Then
        // list_existmidfileinfo.Add(cvitem)
        // iniflg = True
        // End If

        // 'コピー処理
        // '20160913_2 エラー時の処理を追加する修正 -chg sta
        // 'Call Me.Set_BaseMidDataToExistMidFile(existfilename, existsheetname, hash_middata, writerowcnt, iniflg)
        // normalflg = Me.Set_BaseMidDataToExistMidFile(existfilename, existsheetname, hash_middata, writerowcnt, iniflg)
        // If normalflg = False Then
        // errstr = MSG_ERR_EXISTMIDFILEWRITE
        // rtn = False
        // Return rtn
        // End If
        // '20160913_2 エラー時の処理を追加する修正 -chg end
        // Else    '20160913_2 エラー時の処理を追加する修正 -add
        // errstr = MSG_ERR_BASEMIDFILEREAD
        // rtn = False
        // Return rtn

        // End If

        // End If

        // End If

        // Next

        // Return rtn

        // End Function
        // 20160929 汎用→既存コピー処理改善対応 -del end

        /// <summary>
        /// ダミー用の金融機関9999を作成する処理
        /// </summary>
        /// <returns>終了判定 True:正常終了 False:異常終了</returns>
        /// <remarks>
        /// 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 新規追加
        /// </remarks>
        private bool Set_DummyBankData()
        {

            string xmlupdatestr = "<?xml version=" + "\"" + "1.0" + "\"" + " encoding=" + "\"" + "utf-16" + "\"" + "?>";  // 挿入用に成形するための文字列(除去用)
            int tmp_cnt = 0;  // クエリ実行用の作業用変数
            string tmp_sql = "";
            bool rtn = true;

            // ダミー銀行有無確認
            tmp_sql = " SELECT COUNT(*) FROM m_kinyu WHERE kinyu_no = 9999 ";
            string tmp_kinyucnt = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));
            tmp_sql = "";

            // ダミー銀行挿入
            if (tmp_kinyucnt == "0")
            {
                tmp_sql = " INSERT INTO m_kinyu VALUES(9999,'ダミー銀行','','',1,NULL," + "'" + CommonModule.DefHistory.Replace(xmlupdatestr, "") + "'" + ",1,NULL) ";
                rtn = DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
                tmp_sql = "";
                if (rtn == false)
                {
                    return rtn;
                }
            }

            return rtn;

        }

        #endregion

        #region 初期設定処理

        /// <summary>
        /// コマンドライン値取得・条件設定
        /// </summary>
        /// <remarks>
        /// ・コンバートタイプ判別
        /// 　第1引数…コンバートタイプ (0.汎用, 1.既存ユーザ用(V7), 2.他社システム用, 左記以外.既存ユーザ用(DEF))
        /// 　第2引数…_"njc_dev".開発用機能表示
        /// 　第3引数…ユーザ名またはシステム名 [呼出起動時または他社システム用の場合に使用]
        /// 　※引数なし…P1.既存ユーザ用＋P2.開発用処理なし
        /// 
        /// ・パラメータ渡し例
        /// 　1 njc_dev  … 既存ユーザ用＋開発用機能表示
        /// 　0          … 汎用＋開発用機能なし
        /// 　2 0 nissin … 他社システム用＋開発用機能なし＋日新不動産様用
        /// </remarks>
        private void IniCmdline()
        {

            string[] cmdline = Environment.GetCommandLineArgs();
            int cmdlinecnt = Information.UBound(cmdline);

            for (int cntii = 0, loopTo = Information.UBound(cmdline); cntii <= loopTo; cntii++)
            {
                switch (cntii)
                {
                    case 0:
                        {
                            break;
                        }
                    // なし
                    case 1:
                        {
                            break;
                        }
                    // コンバートタイプ取得
                    // 20160926 コンバーター起動時パラメータの仕様変更修正 -del sta
                    // Dim tmp_cnvnostr As String = cmdline(cntii).ToString
                    // Dim tmp_cnvnoint As Integer = 0
                    // If (Int32.TryParse(tmp_cnvnostr, tmp_cnvnoint)) = False Then
                    // tmp_cnvnoint = 1
                    // End If
                    // CNVNO = tmp_cnvnoint
                    // 20160926 コンバーター起動時パラメータの仕様変更修正 -del end
                    case 2:
                        {
                            // 開発用機能表示
                            if ((cmdline[cntii] ?? "") == CommonModule.NJC_DEV)
                            {
                                CommonModule.Dev_CVFlg = true;
                            }

                            break;
                        }
                }
            }

            // 20160926 コンバーター起動時パラメータの仕様変更修正 -add sta
            // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
            // 確認メモ
            // →既存・汎用・他社の各々で固定値をハードコーディングする
            // CNVNO = ConvertTypes._既存ユーザ用
            CommonModule.CNVNO = (int)CommonModule.ConvertTypes._汎用;
            // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
            // 20160926 コンバーター起動時パラメータの仕様変更修正 -add end

        }

        /// <summary>
        /// メイン.開始タブ設定(1)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMStart()
        {

            lblHFirstLabel.Visible = false;
            lblHFirstLabel.Visible = false;
            grpHFirstNaiyo.Visible = false;
            lblTitleH.Visible = false;
            lblHFirstLabel.Location = new System.Drawing.Point(30, 20);
            grpHFirstNaiyo.Location = new System.Drawing.Point(65, 115);
            lblTitleH.Location = new System.Drawing.Point(33, 18);

            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        lblHFirstLabel.Visible = true;
                        grpHFirstNaiyo.Visible = true;
                        grpHFirstNaiyo.Visible = true;
                        lblTitleH.Visible = true;
                        break;
                    }
            }

        }

        /// <summary>
        /// メイン.接続設定タブ設定(2)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMSession()
        {

            // --------------------------------------------------
            // 内容説明ラベル
            // --------------------------------------------------
            lblHSessionDescription1.Visible = false;
            lblHSessionDescription1.Location = new System.Drawing.Point(30, 20);
            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        lblHSessionDescription1.Visible = true;
                        break;
                    }
            }


            // --------------------------------------------------
            // コントロール初期設定
            // --------------------------------------------------
            // 初期接続情報読込ボタン制御
            btnDefConInfoRead.Visible = false;

            // 接続値取得(V7.レジストリ, 10.XML設定ファイル)
            if (Set_DBExistConInfo() == false)
            {
                Set_10DBConInfo();
            }
            else
            {
                btnDefConInfoRead.Visible = true;
            }

            // 接続値セット
            Set_DefConInfo_To_Control(ref fstmodelv10);

            // パスワード入力ボックスマスク設定
            txtV10Pass.PasswordChar = '*';

            // ※ネットワークライブラリはとりあえず非表示
            lblNetworklib.Visible = false;
            // Me.txtV7Networklib.Visible = False
            // Me.txtV10Networklib.Visible = False

            // (開発用) タイムアウト値 
            txtTimeOut.MaxLength = 3;
            grpTimeOut.Visible = CommonModule.Dev_CVFlg;

            // 20160711 ネットワークライブラリ対応 -add sta
            // ネットワークライブラリ
            string[] networklibrary = new string[] { "名前付きパイプ", "共有メモリ", "TCP/IP" };
            cmbV10Networklib.Items.Add("");
            cmbV10Networklib.Items.Add(networklibrary[0]);
            cmbV10Networklib.Items.Add(networklibrary[1]);
            cmbV10Networklib.Items.Add(networklibrary[2]);

            // ※構築中のため非表示
            cmbV10Networklib.Visible = false;
            // 20160711 ネットワークライブラリ対応 -add end
        }

        /// <summary>
        /// メイン.初期設定タブ設定(3)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMSyoki()
        {

            // 作業者名
            txtRecUser.Text = CommonModule.DEF_REC_USER;


            // --------------------------------------------------
            // パス設定
            // --------------------------------------------------
            // コンバーターログファイルパス
            string tmp_deflogdirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_MAINLOG_NAME);
            txtLogDirPath.Text = tmp_deflogdirpath;


            // --------------------------------------------------
            // コントロール初期設定
            // --------------------------------------------------

            // コンバーターログファイル
            txtLogDirPath.BackColor = Color.White;
            txtLogDirPath.ReadOnly = true;
            txtLogDirPath.ForeColor = Color.Gray;

            // 中間ファイル関連
            txtMidDirLogPath.BackColor = Color.White;
            txtMidDirLogPath.ReadOnly = true;
            txtMidDirLogPath.ForeColor = Color.Gray;

            // (開発用)
            btnDevTabChange.Visible = CommonModule.Dev_CVFlg;

            // 紐付設定ファイル関連
            txtRelationDirPath.BackColor = Color.White;
            txtRelationDirPath.ReadOnly = true;
            txtRelationDirPath.ForeColor = Color.Gray;

        }

        /// <summary>
        /// メイン.コンバート作業選択タブ(4)
        /// </summary>
        /// <remarks>
        /// ・コンバート作業選択
        /// </remarks>
        private void IniFormtabPageMMenu()
        {

            // --------------------------------------------------
            // コントロール初期設定
            // --------------------------------------------------
            btnMenuJizen.Enabled = true;
            btnMenuDatacv.Enabled = false;
            btnMenuJigo.Enabled = true;
            btnNext.Enabled = false;

            lblHMenuDatacv.Location = new Point((Size)lblKMenuDatacv.Location);
            lblHMenuDatacv.Visible = false;
            lblKMenuDatacv.Visible = false;
            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        // 20161018 UI最終確認での修正 -add sta
                        lblHMenuJizen.Location = new Point((Size)lblKMenuJizen.Location);
                        lblKMenuJizen.Visible = false;
                        lblHMenuJizen.Visible = true;
                        // 20161018 UI最終確認での修正 -add end
                        lblHMenuDatacv.Visible = true;
                        break;
                    }

            }

        }

        /// <summary>
        /// メイン.作業選択.事前調整作業タブ(5)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMJizen()
        {

            // タブ隠し(TAB3)
            lblHidden3.Width = 885;
            lblHidden3.Height = 55;


            // --------------------------------------------------
            // パス設定
            // --------------------------------------------------
            // 事前作業での出力リスト格納パス
            string tmp_defjizenlistdirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_LIST_NAME);
            txtJizenListPath.Text = tmp_defjizenlistdirpath;


            // --------------------------------------------------
            // 中間ファイル関連設定
            // --------------------------------------------------
            // 中間ファイルパス
            string tmp_defmiddirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_MID_NAME);
            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        tmp_defmiddirpath = EtcMethod.Set_Path(tmp_defmiddirpath, CommonModule.FILE_HMIDD_NAME);
                        break;
                    }

                default:
                    {
                        break;
                    }

            }
            txtMidDirPath.Text = tmp_defmiddirpath;
            txtMidDirPath.BackColor = Color.White;
            txtMidDirPath.ReadOnly = true;
            txtMidDirPath.ForeColor = Color.Gray;
            grpMiddleFile.Visible = false;
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用 | CommonModule.Dev_CVFlg)
            {
                grpMiddleFile.Visible = true;
            }
            else
            {
                // '※位置調整
                // Me.lblHidden2.Location = New Point(6, 185)
                // Me.lblLine2.Location = New Point(31, 225)
                // Me.tabCtrlCVItem.Location = New Point(30, 200)
            }


            // --------------------------------------------------
            // コントロール初期設定
            // --------------------------------------------------
            // リスト格納先テキストボックス設定
            txtJizenListPath.BackColor = Color.White;
            txtJizenListPath.ReadOnly = true;
            txtJizenListPath.ForeColor = Color.Gray;

            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        pnlJizenListPath.Visible = false;
                        // 20161018 UI最終確認での修正 -add sta
                        Label355.Visible = false;
                        // Me.lblCautionDescription.Text = ""
                        // 20161018 UI最終確認での修正 -add end
                        lblJizenDescription1.Text = "移行元データとなる中間ファイルを事前に作成する必要があります。" + Constants.vbCrLf + "(完了したらチェックボックスをONにして下さい。ONにした場合のみデータコンバートを行うことができます)";
                        break;
                    }

            }

        }

        /// <summary>
        /// メイン.作業選択.事後調整作業タブ(6)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMJigo()
        {

            // タブ隠し(TAB4)
            lblHidden4.Width = 883;
            lblHidden4.Height = 26;
        }

        /// <summary>
        /// (仮)メイン.必須項目一括設定タブ(14)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMIkkatu()
        {

        }
        /// <summary>
        /// (仮)メイン.事前調整作業タブ(汎用コンバートキット用)(15)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMHanyoJizen()
        {

        }

        /// <summary>
        /// コンバーター.はじめにタブ(8)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageDCHajimeni()
        {

            // 20160928 同意内容修正 -chg sta
            // Me.btnDoui.Enabled = True
            // '20160725 汎用CVK対応 -add sta
            // pnlDcFstCmt09.Visible = True
            // lblHDatacvHajimeniLabel.Visible = False
            // lblKDatacvHajimeniLabel.Visible = False
            // Select Case CNVNO
            // Case ConvertTypes._汎用
            // lblHDatacvHajimeniLabel.Visible = True
            // '20160921 同意内容修正 -chg sta
            // 'pnlDcFstCmt09.Location = New Size(pnlDcFstCmt08.Location)
            // 'pnlDcFstCmt08.Visible = False
            // pnlDcFstCmtK01.Visible = False
            // pnlDcFstCmtK02.Visible = False
            // '20160921 同意内容修正 -chg end
            // Case Else
            // lblKDatacvHajimeniLabel.Visible = True
            // '20160921 同意内容修正 -chg sta
            // pnlDcFstCmt07.Visible = False
            // pnlDcFstCmt10.Visible = False
            // '20160921 同意内容修正 -chg end
            // End Select
            // '20160725 汎用CVK対応 -add end
            btnDoui.Enabled = true;
            lblHDatacvHajimeniLabel.Visible = true;
            // 20160928 同意内容修正 -chg end

        }

        /// <summary>
        /// コンバーター. 移行項目選択タブ(9)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageDCSelect()
        {

            // --------------------------------------------------
            // 中間ファイル関連設定
            // --------------------------------------------------
            // 中間ファイルログパス
            string tmp_pathstr = "";
            tmp_pathstr = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_MID_NAME);
            string tmp_defmidlogdirpath = EtcMethod.Set_Path(tmp_pathstr, CommonModule.DIR_MIDLOG_NAME);
            txtMidDirLogPath.Text = tmp_defmidlogdirpath;
            // 中間ファイル関連
            txtMidDirLogPath.BackColor = Color.White;
            txtMidDirLogPath.ReadOnly = true;
            txtMidDirLogPath.ForeColor = Color.Gray;
            grpMiddleFile.Visible = false;
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用 | CommonModule.Dev_CVFlg)
            {
                grpMiddleFile.Visible = true;
            }
            else
            {
                // '※位置調整
                // Me.lblHidden2.Location = New Point(6, 185)
                // Me.lblLine2.Location = New Point(31, 225)
                // Me.tabCtrlCVItem.Location = New Point(30, 200)
            }

            // 20160912 既存用→汎用用中間ファイルへのコピー処理 -add sta
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用 & CommonModule.Dev_CVFlg)
            {
                txtExistMidToBaseMid.Text = Path.GetDirectoryName(txtMidDirPath.Text);
            }
            // 20160912 既存用→汎用用中間ファイルへのコピー処理 -add end

            // --------------------------------------------------
            // コントロール初期設定
            // --------------------------------------------------
            // 全移行対象ブロック名取得(グループ・チェックボックス名取得)
            Set_CVItemName();

            // 全チェックボックスON(既存ユーザ用)
            switch (CommonModule.CNVNO)
            {

                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        // 20160801 タイトルマスタ統合処理 -chg sta
                        // Me.pnlKiMstBikotitle.Visible = False
                        pnlKiMstTitle.Visible = false;
                        // 20160801 タイトルマスタ統合処理 -chg end
                        pnlKiSyskanriBase.Visible = false;
                        grpKizonKagi.Visible = false;
                        // '20160905 汎用コンバートの請求情報修正 -del
                        // Me.pnlKiSq.Visible = False
                        pnlKiClaim.Visible = false;
                        pnlSzen.Visible = false;
                        pnlRendo.Visible = false;
                        Label359.Visible = false;
                        pnlKiOw.Location = (Point)new Size(pnlKiJisya.Location);
                        pnlKiJisya.Location = (Point)new Size(pnlKiSyskanriBase.Location);
                        pnlKiMstKasyoClaimrui.Visible = false;                                        // 20160905 汎用の場合の移行対象外チェックボックスを非表示にする -add
                                                                                                      // 20160725 汎用CVK対応 -add end
                        pnlKiMstHendo.Visible = false;                                              // 20160926 選定した移行項目をプログラムへ反映する修正 -add
                                                                                                    // 20161012 家主固定控除件数表示処理の追加 -add sta
                        Label288.Text = "家主固定控除";
                        Label290.Visible = false;
                        Label289.Visible = false;
                        lblKiSqMiBaseCnt.Visible = false;
                        Label160.Visible = false;
                        Label162.Visible = false;
                        lblKiSqAzBaseCnt.Visible = false;
                        Label308.Location = new Point((Size)Label290.Location);  // 「( 参考件数 = 」
                                                                                 // 20161017 家主固定控除の件数表示の修正 -chg sta
                                                                                 // Me.Label311.Location = New Point(Me.Label289.Location)  '「)」
                                                                                 // Me.lblKiSqOwKojoBaseCnt.Location = New Point(Me.lblKiSqMiBaseCnt.Location)
                                                                                 // 契約情報のLocationを引用する
                        Label311.Location = new Point((Size)Label285.Location);  // 「)」
                        lblKiSqOwKojoBaseCnt.Location = new Point((Size)lblKiKyBaseCnt.Location);
                        // 20161017 家主固定控除の件数表示の修正 -chg end
                        Label308.Visible = true;
                        Label311.Visible = true;
                        lblKiSqOwKojoBaseCnt.Visible = true;
                        // 20161012 家主固定控除件数表示処理の追加 -add end
                        // 20161018 UI最終確認での修正 -add
                        Label284.Text = "契約, 入居者";
                        break;
                    }

                default:
                    {
                        break;
                    }

            }

            // 既存用→汎用用中間ファイルへコピーする処理(開発用) -add
            grpExistMidToBaseMid.Visible = CommonModule.Dev_CVFlg;

            // タブ隠し(TAB2)
            lblHidden2.Width = 885;
            lblHidden2.Height = 28;
            lblKiMstBusCnt.Text = "なし";
            lblKiMstSchoolCnt.Text = "なし";
            lblKiMstAreaCnt.Text = "なし";
            lblKiMstHokenruiCnt.Text = "なし";
            lblKiMstTokuyakuCnt.Text = "なし";
            lblKiMstKasyoClaimruiCnt.Text = "なし";
            lblKiMstHendoCnt.Text = "なし";
            lblKiMstKagititleCnt.Text = "なし";
            lblKiMstBikotitleCnt.Text = "なし";
            lblKiSyskanriBaseCnt.Text = "なし";
            lblKiJisyaBaseCnt.Text = "なし";
            lblKiOwBaseCnt.Text = "なし";
            lblKiGyCyukaiBaseCnt.Text = "なし";
            lblKiGySyuzenBaseCnt.Text = "なし";
            lblKiGyLifelineBaseCnt.Text = "なし";
            lblKiGyHokenBaseCnt.Text = "なし";
            lblKiGyYatinhosyoBaseCnt.Text = "なし";
            lblKiGySisetuBaseCnt.Text = "なし";
            lblKiGySekoBaseCnt.Text = "なし";
            lblKiBkBaseCnt.Text = "なし";
            lblKiHyBaseCnt.Text = "なし";
            lblKiHySetubiCnt.Text = "なし";
            lblKiKysBaseCnt.Text = "なし";
            lblKiKyBaseCnt.Text = "なし";
            lblKiSqMiBaseCnt.Text = "なし";
            lblKiSqAzBaseCnt.Text = "なし";
            lblKiClaimBaseCnt.Text = "なし";
            lblKiSzenBaseCnt.Text = "なし";
            lblKiRendoBaseCnt.Text = "なし";
            // 20161108_2 レビュー結果戻り修正時に気付いた箇所の修正 -add
            // 家主請求控除の件数表示文字列初期化処理を追加
            lblKiSqOwKojoBaseCnt.Text = "なし";

            // --------------------------------------------------
            // グループボックス内処理順番設定
            // (ベースタブ→グループ→チェックボックスの順序を指定)
            // --------------------------------------------------
            // 各基本情報(自社,家主,契約者)
            Set_ControlIndex(tabPageBase130);
            // クレーム修繕情報
            Set_ControlIndex(tabPageBase190);


            // --------------------------------------------------
            // チェックボックス処理順番設定
            // --------------------------------------------------
            // 条件(開発用、コンバートタイプ)によるチェックボックスの表示設定
            Set_ChkBoxVisible();

            // マスタ系
            Set_ControlIndex(grpMst);
            Set_ChkBoxValue(grpMst, false);

            // 業者情報
            Set_ControlIndex(grpGy);
            Set_ChkBoxValue(grpGy, false);

            // 自社情報
            Set_ControlIndex(grpJisya);
            Set_ChkBoxValue(grpJisya, false);

            // 家主情報
            Set_ControlIndex(grpOw);
            Set_ChkBoxValue(grpOw, false);

            // 契約者情報
            Set_ControlIndex(grpKys);
            Set_ChkBoxValue(grpKys, false);

            // 物件情報
            Set_ControlIndex(grpBk);
            Set_ChkBoxValue(grpBk, false);

            // 部屋情報
            Set_ControlIndex(grpHy);
            Set_ChkBoxValue(grpHy, false);

            // 送金ルール
            Set_ControlIndex(grpSorule);
            Set_ChkBoxValue(grpSorule, false);

            // 契約情報
            Set_ControlIndex(grpKy);
            Set_ChkBoxValue(grpKy, false);

            // 請求情報
            Set_ControlIndex(grpSq);
            Set_ChkBoxValue(grpSq, false);

            // クレーム情報
            Set_ControlIndex(grpClaim);
            Set_ChkBoxValue(grpClaim, false);

            // 修繕情報
            Set_ControlIndex(grpSzen);
            Set_ChkBoxValue(grpSzen, false);

            // 初期設定
            Set_ControlIndex(grpSyskanri);
            Set_ChkBoxValue(grpSyskanri, false);

            // 物件データ連動情報  
            Set_ControlIndex(grpRendo);
            Set_ChkBoxValue(grpRendo, false);

            // (開発用) 子項目チェック制御
            if (chkChildItemControl.Checked)
            {
                Set_ChkBoxCondition(false);
            }

            // 移行対象外項目をオブジェクトへ格納(既存、汎用で処理を分岐)
            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        Set_ExistCVItemToList();
                        break;
                    }
            }

            // 連動ID入力テキストボックスの制御
            txtRendoID.MaxLength = 100;

            // 20161012 家主固定控除件数表示処理の追加 -del sta
            // '20160905 汎用コンバートの請求情報修正 -add
            // Me.Label288.Text = "導入時未収金, 預り金"
            // 20161012 家主固定控除件数表示処理の追加 -del end

        }

        /// <summary>
        /// コンバーター.処理実行タブ(10)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageDCJikko()
        {

            // -----------------------------
            // プログレスバー初期設定
            // -----------------------------
            // 20161017 進捗表示ラベル初期表示修正 -chg sta
            // Me.lblTotalSituation.Text = "コンバート処理準備中 ..."
            lblTotalSituation.Text = CommonModule.SITUATION_CV_BFRUN;
            // 20161017 進捗表示ラベル初期表示修正 -chg end
            lblCVItem.Text = "";
            txtTotalSituation.ReadOnly = true;
            txtPartialSituation.ReadOnly = true;
            txtTotalSituation.BackColor = Color.White;
            txtPartialSituation.BackColor = Color.White;
            txtTotalSituation.ScrollBars = ScrollBars.Both;
            txtTotalSituation.WordWrap = false;
            txtPartialSituation.ScrollBars = ScrollBars.Both;
            txtPartialSituation.WordWrap = false;

        }

        /// <summary>
        /// コンバーター.終了タブ(11)(12)(13)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageDCEnd()
        {
            // --------------------------------------------------
            // 終了画面の不要コントロール制御
            // --------------------------------------------------
            // キャンセル画面
            GroupBox22.Visible = false;
            GroupBox21.Visible = false;

            // エラー画面
            GroupBox20.Visible = false;
            GroupBox23.Visible = false;
            GroupBox19.Visible = false;

        }

        /// <summary>
        /// (隠)メイン.開発用タブ(0)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMDev()
        {

            // --------------------------------------------------
            // 紐付設定ファイル関連パス
            // --------------------------------------------------
            string tmp_pathstr;
            tmp_pathstr = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_RELEXEDIR_NAME);
            string tmp_defreldirpath = EtcMethod.Set_Path(tmp_pathstr, CommonModule.DIR_REL_NAME);
            txtRelationDirPath.Text = tmp_defreldirpath;


            // --------------------------------------------------
            // 各チェックボックス初期設定
            // --------------------------------------------------
            chkCVStart.Checked = true;
            chkLogTblDrop.Checked = true;
            chkV7ViewDrop.Checked = true;
            chkMiddleFile.Checked = true;
            chkOverWrite.Checked = false;
            chkRelTblDrop.Checked = true;
            chkMidNotStop.Checked = true;
            chkRelation.Checked = true;

            // (汎用)
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
            {
                chkMiddleFile.Checked = false;
            }

            // (テスト用)                                                            
            chkV7ViewDrop.Checked = false;
            chkLogTblDrop.Checked = false;
            chkRelTblDrop.Checked = false;

        }

        /// <summary>
        /// メイン.初期表示フォーム設定(共通画面)
        /// </summary>
        /// <remarks></remarks>
        private void IniFormtabPageMain()
        {

            // --------------------------------------------------
            // ラベル(タイトル)設定
            // --------------------------------------------------
            Text = "";
            switch (CommonModule.CNVNO)
            {

                default:
                    {
                        Text = Text + CommonModule.CV_HANYO_TITLE;
                        Text = Text + $" ver{Assembly.GetExecutingAssembly().GetName().Version} ";
                        Text = Text + ":" + CommonModule.CV_FROM_MIDDLE + " → " + CommonModule.CV_TO_NAME;
                        break;
                    }
            }


            // --------------------------------------------------
            // コントロール初期設定
            // --------------------------------------------------
            // サイズ設定
            // ----- 要修正 ----- sta
            // ※サイズ可変対応
            FormBorderStyle = FormBorderStyle.FixedSingle;
            // ----- 要修正 ----- end

            // フォームコントロールボックス非可視設定
            ControlBox = false;

            // タブ隠し(TAB1)                                                  
            lblHidden1.Width = 950;
            lblHidden1.Height = 55;

            // ボタンの活性制御
            Chg_BtnKariStatus("mstart");

            // タブ非表示設定("1.開始タブ"以外は非表示)
            int tmpcntii = 0;
            tabPageManager = new TabPageManager(tabCtrlMain);
            int tmpcntjj = tabCtrlMain.TabPages.Count - 1;
            for (int cntii = tmpcntii, loopTo = tmpcntjj; cntii <= loopTo; cntii++)
                tabPageManager.ChangeTabPageVisible(cntii, false);
            tabPageManager.ChangeTabPageVisible(1, true);

            // (開発用) 開発用タブON/OFFボタン表示設定
            btnDevTabChange.Visible = CommonModule.Dev_CVFlg;

            // 中間用プログレスバー非表示(必要時のみ表示)
            pnlPrgChk.Visible = false;


            // --------------------------------------------------
            // タブオブジェクト生成
            // -------------------------------------------------- 
            // 共通設定値
            tmpcntii = 1;
            // 作業選択.事前作業
            tabPageManagerJizen = new TabPageManager(tabCtrlJizen);
            tmpcntjj = tabCtrlJizen.TabPages.Count - 1;
            for (int cntii = tmpcntii, loopTo1 = tmpcntjj; cntii <= loopTo1; cntii++)
                tabPageManagerJizen.ChangeTabPageVisible(cntii, false);
            // 作業選択.コンバーター.選択項目
            tabPageManagerCvitem = new TabPageManager(tabCtrlCVItem);
            tmpcntjj = tabCtrlCVItem.TabPages.Count - 1;
            for (int cntii = tmpcntii, loopTo2 = tmpcntjj; cntii <= loopTo2; cntii++)
                tabPageManagerCvitem.ChangeTabPageVisible(cntii, false);
            // 作業選択.事後作業
            tabPageManagerJigo = new TabPageManager(tabCtrlJigo);
            tmpcntjj = tabCtrlJigo.TabPages.Count - 1;
            for (int cntii = tmpcntii, loopTo3 = tmpcntjj; cntii <= loopTo3; cntii++)
                tabPageManagerJigo.ChangeTabPageVisible(cntii, false);


            // --------------------------------------------------
            // 画面左進捗・セットフォーカス
            // --------------------------------------------------
            Chg_LeftProgress("mstart");
            Set_Forcus("next");

        }

        /// <summary>
        /// 画面左進捗表示設定
        /// </summary>
        /// <param name="chgstr">状態文字列</param>
        /// <param name="chgtype">状態タイプ…0.メイン画面, 1.作業選択画面 [op]</param>
        /// <remarks>
        /// ・デフォルトタイプ値=0.メイン画面
        /// </remarks>
        private void Chg_LeftProgress(string chgstr, int chgtype = 0)
        {

            switch (chgtype)
            {

                case 1:
                    {
                        // --------------------------------------------------
                        // コンバーター画面左進捗表示
                        // --------------------------------------------------
                        // 画面左進捗表示
                        pnlMenuDatacv.Visible = true;                // コンバーター画面左進捗
                                                                     // 背景初期化
                        pnlPrgDCHajimeni.BackColor = Color.LightBlue;
                        pnlPrgDCSelect.BackColor = Color.LightBlue;
                        pnlPrgDCJikko.BackColor = Color.LightBlue;
                        pnlPrgDCEnd.BackColor = Color.LightBlue;
                        pnlPrgDCFileWrite.BackColor = Color.LightBlue;
                        pnlPrgDCRelation.BackColor = Color.LightBlue;
                        pnlPrgDCConv.BackColor = Color.LightBlue;
                        // アイコン(青)表示
                        picIcoDCHajimeni1.Visible = true;
                        picIcoDCSelect1.Visible = true;
                        picIcoDCJikko1.Visible = true;
                        picIcoDCEnd1.Visible = true;
                        picIcoDCFileWrite1.Visible = true;
                        picIcoDCRelation1.Visible = true;
                        picIcoDCConv1.Visible = true;
                        // アイコン(赤)表示
                        picIcoDCHajimeni2.Visible = false;
                        picIcoDCSelect2.Visible = false;
                        picIcoDCJikko2.Visible = false;
                        picIcoDCEnd2.Visible = false;
                        picIcoDCFileWrite2.Visible = false;
                        picIcoDCRelation2.Visible = false;
                        picIcoDCConv2.Visible = false;
                        // 矢印(青)表示
                        picArwDCSelect1.Visible = true;
                        picArwDCJikko1.Visible = true;
                        picArwDCEnd1.Visible = true;
                        // 矢印(赤)表示
                        picArwDCSelect2.Visible = false;
                        picArwDCJikko2.Visible = false;
                        picArwDCEnd2.Visible = false;
                        // 矢印(赤)前面表示
                        picArwDCSelect2.BringToFront();
                        picArwDCJikko2.BringToFront();
                        picArwDCEnd2.BringToFront();
                        // 背景前面表示
                        pnlPrgDCHajimeni.BringToFront();
                        pnlPrgDCSelect.BringToFront();
                        pnlPrgDCJikko.BringToFront();
                        pnlPrgDCEnd.BringToFront();
                        pnlPrgDCFileWrite.BringToFront();
                        pnlPrgDCRelation.BringToFront();
                        pnlPrgDCConv.BringToFront();
                        // カレント箇所状態変更
                        switch (chgstr ?? "")
                        {
                            case "dchajimeni":
                                {
                                    pnlPrgDCHajimeni.BackColor = Color.LightPink;
                                    picIcoDCHajimeni2.Visible = true;
                                    break;
                                }
                            case "dcselect1":
                            case "dcselect2":
                                {
                                    pnlPrgDCSelect.BackColor = Color.LightPink;
                                    picIcoDCSelect2.Visible = true;
                                    picArwDCSelect2.Visible = true;
                                    break;
                                }
                            case "dcjikko1":
                            case "dcjikko2":
                                {
                                    pnlPrgDCJikko.BackColor = Color.LightPink;
                                    picIcoDCJikko2.Visible = true;
                                    picArwDCJikko2.Visible = true;
                                    break;
                                }
                            case "dcend":
                                {
                                    pnlPrgDCEnd.BackColor = Color.LightPink;
                                    picIcoDCEnd2.Visible = true;
                                    picArwDCEnd2.Visible = true;
                                    break;
                                }
                            case "dcfilewrite":
                                {
                                    pnlPrgDCJikko.BackColor = Color.LightPink;
                                    picIcoDCJikko2.Visible = true;
                                    picArwDCJikko2.Visible = true;
                                    pnlPrgDCFileWrite.BackColor = Color.LightPink;
                                    picIcoDCFileWrite2.Visible = true;
                                    picArwDCJikko2.Visible = true;
                                    break;
                                }
                            case "dcrelation":
                                {
                                    pnlPrgDCJikko.BackColor = Color.LightPink;
                                    picIcoDCJikko2.Visible = true;
                                    picArwDCJikko2.Visible = true;
                                    pnlPrgDCRelation.BackColor = Color.LightPink;
                                    picIcoDCRelation2.Visible = true;
                                    picArwDCJikko2.Visible = true;
                                    break;
                                }
                            case "dcconv":
                                {
                                    pnlPrgDCJikko.BackColor = Color.LightPink;
                                    picIcoDCJikko2.Visible = true;
                                    picArwDCJikko2.Visible = true;
                                    pnlPrgDCConv.BackColor = Color.LightPink;
                                    picIcoDCConv2.Visible = true;
                                    picArwDCJikko2.Visible = true;
                                    break;
                                }
                        }

                        break;
                    }
            }

        }

        /// <summary>
        /// ハッシュテーブル初期化
        /// </summary>
        /// <remarks></remarks>
        public void Ini_HashTable()
        {

            CommonModule.Hash_TblName_JpToAlpha = new SafeDictionary<string, string>();
            CommonModule.Hash_TblName_AlphaToJp = new SafeDictionary<string, string>();
            CommonModule.Hash_FldName_JpToAlpha = new SafeDictionary<string, string>();
            CommonModule.Hash_FldName_AlphaToJp = new SafeDictionary<string, string>();
            CommonModule.Hash_FiledTypeAlpha = new SafeDictionary<string, string>();
            CommonModule.Hash_FiledSizeAlpha = new SafeDictionary<string, string>();
            CommonModule.Hash_FiledTypeJp = new SafeDictionary<string, string>();
            CommonModule.Hash_FiledSizeJp = new SafeDictionary<string, string>();
            CommonModule.Hash_DefaultValue = new SafeDictionary<string, string>();
            CommonModule.Hash_FiledKey = new SafeDictionary<string, string>();
            CommonModule.Hash_Min_Code = new SafeDictionary<string, string>();
            CommonModule.Hash_Max_Code = new SafeDictionary<string, string>();
            CommonModule.Hash_Mst_ReferenceAlpha = new SafeDictionary<string, string>();
            CommonModule.Hash_MidKey_FieldAlpha = new SafeDictionary<string, string>();
            CommonModule.Hash_MidKey_FieldJp = new SafeDictionary<string, string>();
            CommonModule.Hash_FldName_JpToAlpha_Mid = new SafeDictionary<string, string>();

            CommonModule.Hash_Kinyu = new SafeDictionary<string, string>();
            CommonModule.Hash_KinyuTen = new SafeDictionary<string, string>();
            CommonModule.Hash_KagiTitleKyoyo = new SafeDictionary<string, string>();
            CommonModule.Hash_KagiTitleSenyo = new SafeDictionary<string, string>();
            CommonModule.Hash_KagiTitle_V7to10 = new SafeDictionary<string, string>();
            CommonModule.Hash_SetubiMid = new SafeDictionary<string, string>();
            CommonModule.Hash_SetubiMst = new SafeDictionary<string, string>();
            CommonModule.Hash_SetubiMst_UserMake = new SafeDictionary<string, string>(); 		// 20160829 設備の新規挿入処理を追加

            CommonModule.Hash_ClaimBruiKasyo = new SafeDictionary<string, string>();
            CommonModule.Hash_ClaimBruiClaim = new SafeDictionary<string, string>();

            // 20160711 オブジェクトインスタンス化の修正 -add sta
            // 紐付データ取得用のハッシュテーブルも初期化する
            CommonModule.Hash_Rel_Bkrui = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Kozo = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Hyrui = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Toritaiyo = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Kozasyubetu = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Nkinkomk = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_NkinkomkZksei = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Nkinkbn = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Hendometer = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Setubi = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Kyrui = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Kyrui_Teikisyakuyakbn = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Tosiyotokbn = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_Tosiyotono = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_JisyaKoza = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_FBInfo_KozaFurikaeFmt = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_FBInfo_FuriIraiFmt = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_FBInfo_NsSettingFmt = new SafeDictionary<string, string>();
            CommonModule.Hash_Rel_FBInfo_NsSettingJisya = new SafeDictionary<string, string>();
            // 20160711 オブジェクトインスタンス化の修正 -add end

        }

        /// <summary>
        /// コンバート関連フォルダ自動生成処理 '20161018 フォルダ自動生成箇所の修正 -add
        /// </summary>
        /// <remarks></remarks>
        private bool IniDirSet()
        {

            // --------------------------------------------------
            // フォルダ自動生成処理
            // --------------------------------------------------
            string tmp_middirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_MID_NAME);             // コンバーター実行ファイル格納フォルダ
            string tmp_reldirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_RELEXEDIR_NAME);       // 紐付ツール実行ファイル格納フォルダ
            string tmp_gazodirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_GAZOEXEDIR_NAME);     // 画像CVツール実行ファイル格納フォルダ
            var list_makedir = new List<string>();

            // コンバーターフォルダ同階層
            list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_MAINLOG_NAME));
            list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_INI_NAME));
            list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_LIST_NAME));
            // 中間ファイルフォルダ
            if (EtcMethod.Chk_DirExist(tmp_middirpath))
            {
                list_makedir.Add(EtcMethod.Set_Path(tmp_middirpath, CommonModule.DIR_TEMP_CSV));
                list_makedir.Add(EtcMethod.Set_Path(tmp_middirpath, CommonModule.DIR_MIDLOG_NAME));
            }
            // 紐付ツールフォルダ
            if (EtcMethod.Chk_DirExist(tmp_reldirpath))
            {
                list_makedir.Add(EtcMethod.Set_Path(tmp_reldirpath, CommonModule.DIR_RELLOG_NAME));
            }
            // 画像CVツールフォルダ
            // 20161108_2 レビュー結果戻り修正 -chg sta
            // If EtcMethod.Chk_DirExist(tmp_gazodirpath) Then                                         '20161108 レビュー結果：汎用CVKの場合でも作成する？
            // list_makedir.Add(EtcMethod.Set_Path(tmp_gazodirpath, DIR_INI_NAME))
            // End If
            // 20161108_2 レビュー結果戻り修正 -chg end
            // 自動生成処理
            // 20161104 フォルダ自動作成処理の修正 -chg sta
            // For Each chkdir In list_makedir
            // If System.IO.Directory.Exists(chkdir) = False Then
            // System.IO.Directory.CreateDirectory(chkdir)
            // End If
            // Next

            try
            {

                foreach (var chkdir in list_makedir)
                {
                    if (Directory.Exists(chkdir) == false)
                    {
                        Directory.CreateDirectory(chkdir);
                    }
                }
            }

            // アクセスが拒否された場合
            // Catch ex As UnauthorizedAccessException
            // MsgResult = MessageBox.Show("フォルダ作成時にアクセスが拒否されました。" & vbCrLf & _
            // "コンバーターのインストール先にアクセス権限を与えるか、インストール先を変更して再度セットアップを行って下さい。", _
            // "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            // Return False

            // 無効なフォルダパスが指定された場合
            catch (DirectoryNotFoundException ex)
            {
                CommonModule.MsgResult = MessageBox.Show("無効なフォルダパスが指定されました。" + Constants.vbCrLf + "関連フォルダが削除された可能性があります。再度セットアップを行ってから、本プログラムを実行して下さい。", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return false;
            }

            // その他(アクセスが拒否された場合を含む)
            catch (Exception ex)
            {
                CommonModule.MsgResult = MessageBox.Show("セットアップフォルダのアクセスが拒否されました。" + Constants.vbCrLf + "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return false;

            }

            return true;
            // 20161104 フォルダ自動作成処理の修正 -chg end
        }

        #endregion

        #region イベント処理

        /// <summary>
        /// イベント処理：メイン画面ロード
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void FrmMain_Load(object sender, EventArgs e)
        {

            // --------------------------------------------------
            // 共通初期設定
            // --------------------------------------------------
            // コマンドライン値取得・条件設定
            IniCmdline();
            // 20161018 フォルダ自動生成箇所の修正 -add sta
            // フォルダ自動生成処理
            // 20161104 フォルダ自動作成処理の修正 -chg sta
            // Call Me.IniDirSet()
            if (IniDirSet() == false)
            {
                Close();
                return;
            }
            // 20161104 フォルダ自動作成処理の修正 -chg end
            // 20161018 フォルダ自動生成箇所の修正 -add end
            // --------------------------------------------------
            // メイン画面＋各タブの初期設定
            // --------------------------------------------------
            // メインタブ設定
            IniFormtabPageMStart();               // 1.開始タブ
            IniFormtabPageMSession();             // 2.DB接続タブ
            IniFormtabPageMSyoki();               // 3.初期設定タブ
            IniFormtabPageMMenu();                // 4.作業選択タブ

            // メイン.作業選択ボタン設定
            IniFormtabPageMJizen();               // 5.事前調整作業タブ                      
            IniFormtabPageMJigo();                // 6.事後調整作業タブ                     

            // コンバータータブ設定
            IniFormtabPageDCHajimeni();           // 8.はじめにタブ
            IniFormtabPageDCSelect();             // 9.対象項目選択タブ                      
            IniFormtabPageDCJikko();              // 10.処理実行タブ
            IniFormtabPageDCEnd();                // 11～13.終了タブ

            // 開発用・仮用意タブ設定
            IniFormtabPageMIkkatu();              // 14.(仮)必須項目一括設定タブ
            IniFormtabPageMHanyoJizen();          // 15.(仮)事前調整作業タブ(汎用CVK用)
            IniFormtabPageMDev();                 // 0.開発用タブ

            // メイン画面設定
            IniFormtabPageMain();

            // 20160719 V7オプション選択による表示設定処理追加_不足分対応 -add sta
            // ※画面起動時にチェックボックスイベントに入るため起動時は入らないようにするフラグ
            frmloadflg_kysq = false;
            frmloadflg_clsz = false;
            // 20160719 V7オプション選択による表示設定処理追加_不足分対応 -add end

        }

        /// <summary>
        /// イベント処理：作業選択.事前作業ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnMenuJizen_Click(object sender, EventArgs e)
        {

            tabPageManager.ChangeTabPageVisible(5, true);
            tabPageManager.ChangeTabPageVisible(4, false);

            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        tabPageManagerJizen.ChangeTabPageVisible(0, true);
                        Chg_BtnKariStatus("jizenh");
                        Set_Forcus("jizen_tyukan");
                        break;
                    }
            }

        }

        /// <summary>
        /// イベント：作業選択.データコンバートボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnMenuDatacv_Click(object sender, EventArgs e)
        {

            tabPageManager.ChangeTabPageVisible(8, true);
            tabPageManager.ChangeTabPageVisible(4, false);
            Chg_BtnKariStatus("dchajimeni");
            Chg_LeftProgress("dchajimeni", 1);
            Set_Forcus("dcdoui");

        }

        /// <summary>
        /// イベント：作業選択.事後作業ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnMenuJigo_Click(object sender, EventArgs e)
        {

            tabPageManager.ChangeTabPageVisible(6, true);
            tabPageManager.ChangeTabPageVisible(4, false);

            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        tabPageManagerJigo.ChangeTabPageVisible(1, true);
                        tabPageManagerJigo.ChangeTabPageVisible(0, false);
                        Chg_BtnKariStatus("jigoh");
                        Set_Forcus("next");
                        break;
                    }
            }

        }

        /// <summary>
        /// イベント処理：戻るボタンクリック　
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnBack_Click(object sender, EventArgs e)
        {

            string seltab = tabCtrlMain.SelectedTab.Name;

            switch (seltab ?? "")
            {

                // ================================================== メインタブ ================================================== -sta
                case "tabPageStart":         // なし
                    {
                        break;
                    }

                case "tabPageSession":       // 戻る [2.接続設定 -> 1.初期設定]
                    {
                        tabPageManager.ChangeTabPageVisible(1, true);
                        tabPageManager.ChangeTabPageVisible(2, false);
                        Chg_BtnKariStatus("mstart");
                        Chg_LeftProgress("mstart");
                        Set_Forcus("next");
                        break;
                    }

                case "tabPageSyoki":         // 戻る [3.初期設定 -> 2.開始画面]
                    {
                        tabPageManager.ChangeTabPageVisible(2, true);
                        tabPageManager.ChangeTabPageVisible(3, false);
                        Chg_BtnKariStatus("msession");
                        Chg_LeftProgress("msession");
                        Set_Forcus("prev");
                        break;
                    }

                case "tabPageMenu":          // 戻る [4.作業選択 -> 3.接続設定]
                    {
                        tabPageManager.ChangeTabPageVisible(3, true);
                        tabPageManager.ChangeTabPageVisible(4, false);
                        Chg_BtnKariStatus("msyoki");
                        Chg_LeftProgress("msyoki");
                        Set_Forcus("prev");
                        break;
                    }

                // ================================================== メインタブ ================================================== -end

                // ================================================== 作業選択タブ ================================================== -sta

                case "tabPageSelect":        // 戻る [9.項目選択 -> 8.はじめに]
                    {

                        switch (tabCtrlCVItem.SelectedTab.Name ?? "")
                        {
                            case DC_SELTAB_K01:
                                {
                                    tabPageManager.ChangeTabPageVisible(8, true);
                                    tabPageManager.ChangeTabPageVisible(9, false);
                                    Chg_BtnKariStatus("dchajimeni");
                                    Chg_LeftProgress("dchajimeni", 1);
                                    Set_Forcus("prev");
                                    if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                                    {
                                        pnlRefresh.Visible = false;                       // 件数再読込ボタン表示設定
                                    }

                                    break;
                                }

                            case DC_SELTAB_K02:
                                {
                                    tabPageManagerCvitem.ChangeTabPageVisible(0, true);
                                    tabPageManagerCvitem.ChangeTabPageVisible(1, false);
                                    Chg_BtnKariStatus(1.ToString());
                                    Chg_LeftProgress("dcselect1", 1);
                                    Set_Forcus("prev");
                                    break;
                                }

                            case DC_SELTAB_K03:
                                {
                                    tabPageManagerCvitem.ChangeTabPageVisible(1, true);
                                    tabPageManagerCvitem.ChangeTabPageVisible(2, false);
                                    Chg_BtnKariStatus("dcselect2");
                                    Set_Forcus("prev");
                                    break;
                                }

                            default:
                                {
                                    break;
                                }
                                // ※上記以外のタブは表示させない
                        }

                        break;
                    }

                case "tabPageJikko":         // なし
                    {
                        break;
                    }


                case "tabPageEndOK":         // なし
                    {
                        break;
                    }

                // ================================================== 作業選択タブ ================================================== -end

                // ================================================== コンバータータブ ================================================== -sta
                case "tabPageJigo":          // 戻る [6.事後作業 -> ]
                    {

                        switch (tabCtrlJigo.SelectedTab.Name ?? "")
                        {
                            case "tabPageJigo1":
                                {
                                    tabPageManager.ChangeTabPageVisible(4, true);
                                    tabPageManager.ChangeTabPageVisible(6, false);
                                    pnlRefresh.Visible = false;                           // 件数再読込ボタン表示設定
                                    Chg_BtnKariStatus("mmenu");
                                    Chg_LeftProgress("mmenu");
                                    Set_Forcus("prev");
                                    break;
                                }

                            default:
                                {
                                    break;
                                }

                        }

                        break;
                    }

                default:
                    {
                        break;
                    }

            }
            // ================================================== コンバータータブ ================================================== -end

        }

        /// <summary>
        /// イベント処理：次へボタンクリック　
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnNext_Click(object sender, EventArgs e)
        {

            string seltab = tabCtrlMain.SelectedTab.Name;

            switch (seltab ?? "")
            {

                // ================================================== メインタブ ================================================== -sta
                case "tabPageStart":     // 次へ [1.メイン.開始 -> 2.メイン.接続設定]
                    {
                        tabPageManager.ChangeTabPageVisible(2, true);
                        tabPageManager.ChangeTabPageVisible(1, false);
                        Chg_BtnKariStatus("msession");
                        Chg_LeftProgress("msession");
                        Set_Forcus("next");
                        if (btnNext.Enabled)
                        {
                            Set_Forcus("next");
                        }
                        else
                        {
                            Set_Forcus("testset");
                        }

                        break;
                    }

                case "tabPageSession":   // 次へ [2.メイン.接続設定 -> 3.メイン.初期設定]
                    {
                        tabPageManager.ChangeTabPageVisible(3, true);
                        tabPageManager.ChangeTabPageVisible(2, false);
                        Chg_BtnKariStatus("msyoki");
                        Chg_LeftProgress("msyoki");
                        Set_Forcus("next");
                        break;
                    }

                case "tabPageSyoki":     // 次へ [3.メイン.初期設定 -> 4.メイン.作業選択]
                    {
                        tabPageManager.ChangeTabPageVisible(4, true);
                        tabPageManager.ChangeTabPageVisible(3, false);
                        pnlRefresh.Visible = false;                                                   // 件数再読込ボタン表示設定
                        Chg_BtnKariStatus("mmenu");
                        Chg_LeftProgress("mmenu");
                        // ※事前作業の全チェックボックスがONの場合のみ、コンバーター処理ボタンを活性化
                        if (Chk_JizenChkitem(tabCtrlJizen.Controls) == false)
                        {
                            btnNext.Enabled = false;
                        }
                        Set_Forcus("menu_jizen");
                        break;
                    }

                case "tabPageMenu":      // なし [4.メイン.作業選択 -> なし]
                    {
                        break;
                    }

                // ================================================== メインタブ ================================================== -end

                // ================================================== 作業選択タブ ================================================== -sta
                case "tabPageJizen":     // 次へ [5.事前作業 -> タブ選択 -> 4.作業選択]
                    {
                        switch (tabCtrlJizen.SelectedTab.Name ?? "")
                        {
                            case "tabPageHJizen1":
                                {
                                    bool jizenallchk = Chk_JizenChkitem(tabCtrlJizen.Controls);
                                    tabPageManagerJizen.ChangeTabPageVisible(0, true);
                                    if (jizenallchk == false)
                                    {
                                        // 未チェックありの場合のメッセージを先に出力
                                        if ((int)MessageBox.Show(CommonModule.MSG_CATION_JIZENCHK, "注意", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == (int)Constants.vbYes)
                                        {
                                            return;
                                        }
                                    }
                                    tabPageManager.ChangeTabPageVisible(4, true);
                                    tabPageManager.ChangeTabPageVisible(5, false);

                                    Chg_BtnKariStatus("mmenu");
                                    if (jizenallchk == false)
                                    {
                                        btnMenuDatacv.Enabled = false;
                                        Set_Forcus("menu_jizen");
                                    }
                                    else
                                    {
                                        btnMenuDatacv.Enabled = true;
                                        Set_Forcus("menu_datacv");
                                    }

                                    break;
                                }

                            default:
                                {
                                    break;
                                }

                        }

                        break;
                    }

                case "tabPageJigo":      // 次へ [6.事後作業 -> タブ選択 -> 4.作業選択] 
                    {
                        switch (tabCtrlJigo.SelectedTab.Name ?? "")
                        {
                            case "tabPageJigo1":
                                {
                                    tabPageManagerJigo.ChangeTabPageVisible(1, true);
                                    tabPageManagerJigo.ChangeTabPageVisible(0, false);
                                    Chg_BtnKariStatus("jigo2");
                                    Set_Forcus("next");
                                    break;
                                }

                            default:
                                {
                                    break;
                                }

                        }

                        break;
                    }

                // ================================================== 作業選択タブ ================================================== -end

                // ================================================== コンバータータブ ================================================== -sta
                case "tabPageHajimeni":  // 次へ [4-1.コンバーター.はじめに -> 4-2.対象項目選択]
                    {
                        tabPageManager.ChangeTabPageVisible(9, true);
                        tabPageManager.ChangeTabPageVisible(8, false);
                        tabPageManagerCvitem.ChangeTabPageVisible(0, true);
                        tabPageManagerCvitem.ChangeTabPageVisible(1, false);
                        tabPageManagerCvitem.ChangeTabPageVisible(2, false);
                        Chg_BtnKariStatus("dcselect1");
                        Chg_LeftProgress("dcselect1", 1);
                        pnlMenuDatacv.Refresh();       // kekeke
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    pnlRefresh.Visible = true;                                            // 件数再読込ボタン表示設定
                                                                                                          // 20161104 中間ファイル読込エラー時の処理対応 -chg sta
                                                                                                          // Call Me.Set_BaseCVItemCnt()
                                    if (Set_BaseCVItemCnt() == false)
                                    {
                                        return;
                                    }

                                    break;
                                }
                                // 20161104 中間ファイル読込エラー時の処理対応 -chg end
                        }

                        // コンバート実績保持
                        Set_CVJisseki();

                        Set_Forcus("next");
                        break;
                    }


                case "tabPageSelect":    // 次へ＋実行 [4-2.コンバーター.対象項目選択 -> タブ選択 -> 処理実行]
                    {
                        CommonModule.CancelFlg = false;
                        CommonModule.MidChkCancelFlg = false;                                         // チェックボタンからの中断フラグを初期化

                        switch (tabCtrlCVItem.SelectedTab.Name ?? "")
                        {
                            case DC_SELTAB_K01:
                                {
                                    tabPageManagerCvitem.ChangeTabPageVisible(1, true);
                                    tabPageManagerCvitem.ChangeTabPageVisible(0, false);
                                    Chg_BtnKariStatus("dcselect2");
                                    Set_Forcus("next");
                                    break;
                                }

                            case DC_SELTAB_K02:
                                {
                                    tabPageManagerCvitem.ChangeTabPageVisible(2, true);
                                    tabPageManagerCvitem.ChangeTabPageVisible(1, false);
                                    Chg_BtnKariStatus("dcselect3");
                                    Set_Forcus("next");
                                    break;
                                }

                            case DC_SELTAB_K03:
                                {
                                    var list_cvitem = new List<string>();                  // 移行項目取得
                                    if (!Chk_CVStartFlg(ref list_cvitem))                 // 実行確認メッセージ
                                    {
                                        return;
                                    }

                                    tabPageManager.ChangeTabPageVisible(10, true);
                                    tabPageManager.ChangeTabPageVisible(9, false);
                                    tabPageManager.ChangeTabPageVisible(1, false);           // 開発タブ非表示            
                                    btnDevTabChange.Enabled = false;                         // 開発用ボタン非活性状態設定
                                    pnlRefresh.Visible = false;                           // 件数再読込ボタン表示設定
                                    Chg_BtnKariStatus("dcjikko1");                      // ボタン状態変更
                                    Chg_LeftProgress("dcjikko1", 1);                    // 画面左進捗画面設定
                                    tabPageJikko.Refresh();                               // コンバートタブ画面初期化
                                    pnlMenuDatacv.Refresh();       // kekeke

                                    // --------------------------------------------------
                                    // メイン処理
                                    // --------------------------------------------------
                                    if (chkCVStart.Checked)                              // (開発用) コンバート処理有無
                                    {
                                        Cnv_Main(list_cvitem);                       // ◎コンバートメイン処理
                                    }
                                    // 20161104 紐付設定値保存用ファイルの読込エラー時の対応 -chg sta
                                    // If CancelFlg Then
                                    // tabPageManager.ChangeTabPageVisible(10, False)
                                    // tabPageManager.ChangeTabPageVisible(13, True)
                                    // Call Chg_BtnKariStatus("dcend")                     'ボタン状態変更
                                    // Call Chg_LeftProgress("dcend", 1)                   '画面左進捗画面設定
                                    // Set_Forcus("next")
                                    // Else        '続行
                                    // Call Chg_BtnKariStatus("dcjikko2")                  'ボタン状態変更
                                    // Set_Forcus("next")                                  'セットフォーカス
                                    // End If
                                    if (CommonModule.CancelFlg)
                                    {
                                        if (relexeerr == 99)
                                        {
                                            // 紐付ツールでエラーが発生した場合の処理
                                            tabPageManager.ChangeTabPageVisible(10, false);
                                            tabPageManager.ChangeTabPageVisible(12, true);
                                            relexeerr = 0;
                                        }
                                        else
                                        {
                                            tabPageManager.ChangeTabPageVisible(10, false);
                                            tabPageManager.ChangeTabPageVisible(13, true);
                                        }
                                        Chg_BtnKariStatus("dcend");                     // ボタン状態変更
                                        Chg_LeftProgress("dcend", 1);                   // 画面左進捗画面設定
                                        Set_Forcus("next");
                                    }
                                    else        // 続行
                                    {
                                        Chg_BtnKariStatus("dcjikko2");                  // ボタン状態変更
                                        Set_Forcus("next");
                                    }                                  // セットフォーカス
                                                                       // 20161104 紐付設定値保存用ファイルの読込エラー時の対応 -chg end
                                    pnlMenuDatacv.Refresh();       // kekeke

                                    bool logdropflg = chkLogTblDrop.Checked;    // (開発用) ログテーブル削除フラグ
                                    Set_LogFile(logdropflg);                         // ログ出力
                                    break;
                                }

                            default:
                                {
                                    break;
                                }
                                // ※上記以外のタブは表示させない
                        }

                        break;
                    }

                case "tabPageJikko":     // 次へ [10.実行処理 -> 終了]
                    {

                        tabPageManager.ChangeTabPageVisible(10, false);
                        if (CommonModule.CancelFlg)
                        {
                            // キャンセル終了時
                            tabPageManager.ChangeTabPageVisible(13, true);
                        }
                        else if (txtPartialSituation.Text.IndexOf("×") <= 0)
                        {
                            // 正常終了時
                            tabPageManager.ChangeTabPageVisible(11, true);
                        }
                        else
                        {
                            // 異常終了時
                            tabPageManager.ChangeTabPageVisible(12, true);
                        }
                        Chg_BtnKariStatus("dcend");
                        Chg_LeftProgress("dcend", 1);
                        Set_Forcus("next");

                        // --------------------------------------------------
                        // ハッシュ初期化 
                        // --------------------------------------------------
                        Ini_HashTable();
                        break;
                    }

                case "tabPageEndOK":
                case "tabPageEndCancel":
                case "tabPageEndError":          // ログを開く [終了, キャンセル, エラー]
                    {

                        Process.Start(CommonModule.LogFilePath);                   // ログファイル表示

                        // --------------------------------------------------
                        // ハッシュ初期化 
                        // --------------------------------------------------
                        // ================================================== コンバータータブ ================================================== -end

                        Ini_HashTable();
                        break;
                    }

                default:
                    {
                        break;
                    }

            }

        }

        /// <summary>
        /// イベント処理：終了ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>
        /// ・コントロールのキャプションから処理を判定
        /// </remarks>
        private void btnEnd_Click(object sender, EventArgs e)
        {

            bool cnncloseflg = false;
            bool frmcloseflg = false;


            // ==================================================
            // 終了・中断制御
            // ==================================================
            switch (Strings.Replace(btnEnd.Text, " ", "") ?? "")
            {
                case "中止":
                    {
                        CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_STOP_A, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                        if (CommonModule.MsgResult == DialogResult.Yes)
                        {
                            int tmpcntii = 2;
                            int tmpcntjj = tabCtrlMain.TabPages.Count - 1;
                            for (int cntii = tmpcntii, loopTo = tmpcntjj; cntii <= loopTo; cntii++)
                                tabPageManager.ChangeTabPageVisible(cntii, false);
                            tabPageManager.ChangeTabPageVisible(4, true);
                            pnlRefresh.Visible = false;                               // 件数再読込ボタン表示設定
                            Chg_BtnKariStatus("mmenu");
                            Chg_LeftProgress("mmenu");
                            Set_Forcus("end");
                        }

                        break;
                    }

                case "終了":
                    {
                        Chg_LeftProgress("mend");

                        string seltab = tabCtrlMain.SelectedTab.Name;
                        CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_END_CV, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                        if (CommonModule.MsgResult == DialogResult.Yes)
                        {
                            cnncloseflg = true;
                            frmcloseflg = true;
                        }
                        else
                        {
                            switch (seltab ?? "")
                            {
                                case "tabPageStart":
                                    {
                                        Chg_LeftProgress("mstart");
                                        break;
                                    }
                                case "tabPageSession":
                                    {
                                        Chg_LeftProgress("msession");
                                        break;
                                    }
                                case "tabPageSyoki":
                                    {
                                        Chg_LeftProgress("msyoki");
                                        break;
                                    }
                                case "tabPageMenu":
                                    {
                                        Chg_LeftProgress("mmenu");
                                        break;
                                    }
                            }
                            Set_Forcus("menu_hojyo");
                        }

                        break;
                    }

                case "キャンセル":
                    {
                        CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_CANCEL_A, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                        if (CommonModule.MsgResult == DialogResult.Yes)
                        {
                            // 処理中断
                            CommonModule.CancelFlg = true;
                        }
                        else
                        {
                            // 処理再開
                        }

                        break;
                    }

                case "作業選択へ":
                    {
                        // コンバート実績保持
                        string seltab = tabCtrlMain.SelectedTab.Name;
                        if (seltab == "tabPageEndOK")
                        {
                            CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_END_CV_TORIREKI, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2);
                            if (CommonModule.MsgResult == DialogResult.Yes)
                            {
                                MakeFile_CVJisseki();
                            }
                            else
                            {
                                // ファイル削除
                                DeleteFile_CVJisseki();
                            }
                        }

                        tabPageManager.ChangeTabPageVisible(4, true);
                        tabPageManager.ChangeTabPageVisible(11, false);
                        pnlRefresh.Visible = false;                                   // 件数再読込ボタン表示設定
                        Chg_BtnKariStatus("mmenu");
                        Chg_LeftProgress("mmenu");
                        if (seltab == "tabPageEndOK")
                        {
                            switch (CommonModule.CNVNO)
                            {
                                case (int)CommonModule.ConvertTypes._汎用:
                                    {
                                        Set_Forcus("menu_jigo");
                                        break;
                                    }
                            }
                        }
                        else
                        {
                            Set_Forcus("menu_datacv");
                        }

                        break;
                    }

                default:
                    {
                        break;
                    }

            }


            // ==================================================
            // 終了処理 (クローズ処理)
            // ==================================================
            if (cnncloseflg)
            {
                cnnv7.CnnClose(sqlcnnv7);
                cnnv10.CnnClose(sqlcnnv10);
            }

            if (frmcloseflg)
            {
                Close();
            }

        }

        /// <summary>
        /// イベント処理：フォルダ参照ボタン  '20160912 既存用→汎用用中間ファイルへのコピー処理 btnExistMidToBaseMidDirSerach を追加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnDirSeach_Click(object sender, EventArgs e)


        {

            var dirdialog = new FolderBrowserDialog();

            // 上部に表示する説明の文字列設定
            dirdialog.Description = "フォルダを指定してください。";

            // ルートフォルダ設定(デフォルトでDesktop)
            dirdialog.RootFolder = Environment.SpecialFolder.Desktop;

            // 最初に選択するフォルダを指定する
            switch (true)
            {
                case object _ when ReferenceEquals(sender, btnLogDirSeach):
                    {
                        if (!string.IsNullOrEmpty(txtLogDirPath.Text))
                        {
                            dirdialog.SelectedPath = txtLogDirPath.Text;
                        }
                        else
                        {
                            dirdialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                        }

                        break;
                    }
                case object _ when ReferenceEquals(sender, btnRelDirSeach):
                    {
                        if (!string.IsNullOrEmpty(txtRelationDirPath.Text))
                        {
                            dirdialog.SelectedPath = txtRelationDirPath.Text;
                        }
                        else
                        {
                            dirdialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                        }

                        break;
                    }
                case object _ when ReferenceEquals(sender, btnMidDirLogSeach):
                    {
                        if (!string.IsNullOrEmpty(txtMidDirLogPath.Text))
                        {
                            dirdialog.SelectedPath = txtMidDirLogPath.Text;
                        }
                        else
                        {
                            dirdialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                        }

                        break;
                    }
                case object _ when ReferenceEquals(sender, btnJizenListDirSeach):
                    {
                        if (!string.IsNullOrEmpty(txtJizenListPath.Text))
                        {
                            dirdialog.SelectedPath = txtJizenListPath.Text;
                        }
                        else
                        {
                            dirdialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                        }

                        break;
                    }
                case object _ when ReferenceEquals(sender, btnExistMidToBaseMidDirSerach):         // 20160912 既存用→汎用用中間ファイルへのコピー処理 -add
                    {
                        if (!string.IsNullOrEmpty(txtExistMidToBaseMid.Text))
                        {
                            dirdialog.SelectedPath = txtExistMidToBaseMid.Text;
                        }
                        else
                        {
                            dirdialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                        }

                        break;
                    }
            }

            // フォルダの新規作成可否設定(True:作成可)
            dirdialog.ShowNewFolderButton = true;

            // ダイアログ表示
            if (dirdialog.ShowDialog(this) == DialogResult.OK)
            {
                switch (true)
                {
                    case object _ when ReferenceEquals(sender, btnLogDirSeach):
                        {
                            txtLogDirPath.Text = dirdialog.SelectedPath;
                            break;
                        }
                    case object _ when ReferenceEquals(sender, btnRelDirSeach):
                        {
                            txtRelationDirPath.Text = dirdialog.SelectedPath;
                            break;
                        }
                    case object _ when ReferenceEquals(sender, btnMidDirLogSeach):
                        {
                            txtMidDirLogPath.Text = dirdialog.SelectedPath;
                            break;
                        }
                    case object _ when ReferenceEquals(sender, btnJizenListDirSeach):
                        {
                            txtJizenListPath.Text = dirdialog.SelectedPath;
                            break;
                        }
                    case object _ when ReferenceEquals(sender, btnExistMidToBaseMidDirSerach):     // 20160912 既存用→汎用用中間ファイルへのコピー処理 -add
                        {
                            txtExistMidToBaseMid.Text = dirdialog.SelectedPath;
                            break;
                        }
                }
            }

        }

        /// <summary>
        /// イベント処理：同意ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnDoui_Click(object sender, EventArgs e)
        {

            btnNext.Enabled = true;
            btnDoui.Enabled = false;
            Set_Forcus("next");

        }

        /// <summary>
        /// イベント処理：コンバート履歴処理(画面上のチェックボックス着色・チェックON/OFFの初期化) '20160711 コンバート履歴処理の追加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnRekiClear_Click(object sender, EventArgs e)
        {

            // 確認メッセージ
            CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_DEL_CV_JISSEKI, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (CommonModule.MsgResult == DialogResult.No)
            {
                return;
            }
            else
            {
                // チェックボックスのチェックON、着色初期化
                // 項目名とチェックボックスを紐付けたハッシュテーブルを生成
                var hash_cvitemtochkbox = new SafeDictionary<string, CheckBox>();
                hash_cvitemtochkbox = Get_Hash_CVChkitemAll();
                foreach (var chkbox in hash_cvitemtochkbox)
                {
                    var tmp_chkbox = new CheckBox();
                    tmp_chkbox = (CheckBox)chkbox.Value;
                    tmp_chkbox.ForeColor = Color.Black;
                    // 20161018 UI最終確認での修正 -chg sta
                    // tmp_chkbox.Checked = True
                    switch (CommonModule.CNVNO)
                    {
                        case (int)CommonModule.ConvertTypes._汎用:
                            {
                                tmp_chkbox.Checked = false;
                                break;
                            }
                    }
                    // 20161018 UI最終確認での修正 -chg end
                }

                // 実績が記録されたファイルの削除
                DeleteFile_CVJisseki();

                // 終了メッセージ
                CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_DEL_CV_JISSEKIEND, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                // ボタンの活性制御
                btnRekiClear.Enabled = false;

            }

            Set_Forcus("next");

        }

        /// <summary>
        /// イベント処理：件数表示リフレッシュボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnRefresh_Click(object sender, EventArgs e)
        {

            // 選択されているタブ名を取得
            string seltab = tabCtrlMain.SelectedTab.Name;

            // 選択されているタブ毎に処理を分岐
            switch (seltab ?? "")
            {
                case "tabPageSelect":
                    {
                        if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)              // ※汎用時のみ表示するボタンだが念の為
                        {
                            Set_BaseCVItemCnt();
                        }

                        break;
                    }
            }

        }

        private void chkJizenSo_CheckedChanged(object sender, EventArgs e)
        {
            // Set_Forcus("next")
        }

        private void chkHJizenTyukan_CheckedChanged(object sender, EventArgs e)
        {
            Set_Forcus("next");
        }

        private void chkHJizenGazoKeisiki_CheckedChanged(object sender, EventArgs e)
        {
            // Set_Forcus("next")
        }

        /// <summary>
        /// イベント処理：接続テストボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>
        /// ・接続テストが成功した場合のみ、次の処理へ進めるようにする
        /// </remarks>
        private void btnConnectTest_Click(object sender, EventArgs e)
        {

            bool flg_authentV10 = optV10Authent1.Checked;
            string condb = "";


            // コントロール設定値を接続情報へ格納
            Set_Control_To_ConModel(ref fstmodelv10);

            bool flg = true;

            // V10接続処理
            if (string.IsNullOrEmpty(fstmodelv10.ServerName) || string.IsNullOrEmpty(fstmodelv10.InitialCatalog) || cnnv10.CnnSession(fstmodelv10, ref sqlcnnv10, flg_authentV10) == false)
            {
                if (string.IsNullOrEmpty(condb))
                {
                    condb = CommonModule.CV_TO_NAME + "DB";
                }
                else
                {
                    condb = condb + "、" + CommonModule.CV_TO_NAME + "DB";
                }
                flg = false;
            }

            // 接続結果
            if (flg)
            {
                //一旦ここでApplicationDbContextのインスタンス生成を行う
                if (string.IsNullOrWhiteSpace(pre_table_connection_string.Text))
                {
                    MainFrmHelper.db = new ApplicationDbContext(sqlcnnv10.ConnectionString);
                }
                else
                {
                    MainFrmHelper.db = new ApplicationDbContext(pre_table_connection_string.Text);
                }

                // 成功時
                CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_CNN_SUCCESS_A, "成功", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                btnNext.Enabled = true;
                btnConnectTest.Enabled = false;
                grp10ConnectInfo.Enabled = false;
                txtTimeOut.Enabled = false;
                btnDefConInfoRead.Enabled = false;

                // 接続情報を保存
                // '20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg sta
                // Call Me.MakeFile_DBConInfo(fstmodelv7, 0)
                // Call Me.MakeFile_DBConInfo(fstmodelv10, 1)
                if (MakeFile_DBConInfo(fstmodelv10) == false)
                {
                    Close();
                    return;
                }
                // 20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg end
                Set_Forcus("next");
            }
            else
            {
                // 失敗時
                CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_CNN_FAILURE_A + Constants.vbCrLf + "接続先：" + condb, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cnnv7.CnnClose(sqlcnnv7);
                cnnv10.CnnClose(sqlcnnv10);
            }

        }

        /// <summary>
        /// イベント処理：移行項目全選択
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>
        /// ・各タブの最左上部のチェックボックスを基準にする
        /// </remarks>
        private void btnAllChk_Click(object sender, EventArgs e)
        {

            Set_ChkboxOnOff(tabCtrlCVItem.SelectedTab.Name, true);
            Set_OptSelect();

        }

        /// <summary>
        /// イベント処理：オプションボタン設定値変更
        /// </summary>
        /// <remarks></remarks>
        private void opt_CheckedChanged(object sender, EventArgs e)

        {

            switch (true)
            {
                case object _ when ReferenceEquals(sender, optV10Authent1):
                    {
                        txtV10User.Enabled = Conversions.ToBoolean(((RadioButton)sender).Checked);
                        txtV10Pass.Enabled = Conversions.ToBoolean(((RadioButton)sender).Checked);
                        break;
                    }
            }

        }

        /// <summary>
        /// イベント処理：中間ファイルチェックボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnMidFileCheck_Click(object sender, EventArgs e)
        {

            bool normalflg = true;
            bool btnflg = true;
            string errstr = "";
            string btntxt = Strings.Trim(Conversions.ToString(((Button)sender).Text));


            switch (Strings.Trim(btntxt) ?? "")
            {
                case "キャンセル":
                    {
                        CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_CANCEL_B, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                        if (CommonModule.MsgResult == DialogResult.Yes)
                        {
                            CommonModule.MidChkCancelFlg = true;
                            Chg_BtnKariStatus("midcheck");
                        }

                        break;
                    }

                case "チェック":
                    {

                        CommonModule.MidChkCancelFlg = false;                                                     // 中断フラグを初期化
                        CommonModule.CancelFlg = false;                                                           // 実行ボタンの中断フラグを初期化
                        var list_cv = new List<string>();                                          // 移行項目取得

                        if (!Chk_CVStartFlg(ref list_cv, btnflg))                                 // 実行確認メッセージ→実行有無
                        {
                            return;
                        }

                        btnMidFileCheck.Image = My.Resources.Resources.x;
                        Chg_BtnKariStatus("midstop");

                        // ----- 中間用プログレスバー初期化 -----
                        if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                        {
                            lblCheckSituation.Text = CommonModule.SITUATION_MID_READ;
                            lblPgbCheck.Text = "0 %";
                            pgbCheck.Value = 0;
                            pnlPrgChk.Visible = true;
                            tabCtrlCVItem.Enabled = false;
                            btnMidDirLogSeach.Enabled = false;
                            pnlRefresh.Enabled = false;
                            Refresh();
                        }

                        // --------------------------------------------------
                        // ハッシュテーブル初期化 
                        // --------------------------------------------------
                        Ini_HashTable();

                        // --------------------------------------------------
                        // 中間ファイル作成 
                        // --------------------------------------------------
                        // 汎用用中間ファイルから既存用中間ファイルへのコピー処理
                        normalflg = Set_BaseMidFileToExistingMidFile(list_cv, ref errstr);
                        if (normalflg == false)
                        {
                            CommonModule.MsgResult = MessageBox.Show(errstr, "注意", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                            Chg_BtnKariStatus("midcheck");
                            btnMidFileCheck.Image = My.Resources.Resources.機能;
                            return;
                        }

                        // --------------------------------------------------
                        // 中間ファイルチェック
                        // --------------------------------------------------
                        // 20161009 改善対応：中断処理が中断しない -chg sta
                        // Call Me.Before_DBWrite(list_cv)                                             '実行前準備 
                        // normalflg = Me.Chk_MidFileChk_Main(list_cv)                                 'チェック   
                        if (CommonModule.MidChkCancelFlg == false)
                        {
                            Before_DBWrite(list_cv);                                         // 実行前準備 
                            normalflg = Chk_MidFileChk_Main(list_cv);                             // チェック                  
                        }
                        // 20161009 改善対応：中断処理が中断しない -chg end

                        // --------------------------------------------------
                        // エラー時対処 (20160222)
                        // --------------------------------------------------
                        string tmp_endmsg = "";
                        if (CommonModule.MidChkCancelFlg)                                                     // 中断時
                        {
                            tmp_endmsg = CommonModule.MSG_STOP_MIDCHK;
                        }
                        else if (normalflg == false)                                               // エラー時
                        {
                            tmp_endmsg = CommonModule.MSG_ERR_CHK_MID;
                        }
                        else                                                                        // 通常
                        {
                            tmp_endmsg = CommonModule.MSG_END_MIDCHK;
                        }
                        CommonModule.MsgResult = MessageBox.Show(tmp_endmsg, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                        // 20161009 改善対応：中断処理が中断しない -add sta
                        // System.Diagnostics.Process.Start(MiddleLogFilePath)                         'ログファイル表示
                        if (CommonModule.MidChkCancelFlg == false)
                        {
                            Process.Start(CommonModule.MiddleLogFilePath);                     // ログファイル表示
                        }
                        // 20161009 改善対応：中断処理が中断しない -add end
                        Chg_BtnKariStatus("midcheck");

                        pnlPrgChk.Visible = false;                                                // 中間用プログレスバー表示設定
                        tabCtrlCVItem.Enabled = true;
                        btnMidDirLogSeach.Enabled = true;
                        pnlRefresh.Enabled = true;
                        btnMidFileCheck.Image = My.Resources.Resources.機能;          // ボタンのアイコンセット
                        break;
                    }

                default:
                    {
                        break;
                    }

            }

        }

        /// <summary>
        /// イベント処理：接続情報の初期値読込
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnDefConInfoRead_Click(object sender, EventArgs e)
        {

            // デフォルト値読込
            Set_10DBConInfo();

            // コントロールへセット
            Set_DefConInfo_To_Control(ref fstmodelv10);

            // セットフォーカス
            Set_Forcus("testset");

        }

        /// <summary>
        /// イベント処理：チェックボックスチェンジ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>・既存システム用-項目1</remarks>
        private void chkKiMstBus_CheckedChanged(object sender, EventArgs e)
        {
            chkMstBus.Checked = chkKiMstBus.Checked;
            chkMstBusKotu.Checked = chkKiMstBus.Checked;
        }

        private void chkKiMstSchool_CheckedChanged(object sender, EventArgs e)
        {
            chkMstSchool.Checked = chkKiMstSchool.Checked;
        }

        private void chkKiMstArea_CheckedChanged(object sender, EventArgs e)
        {
            chkMstArea.Checked = chkKiMstArea.Checked;
        }

        private void chkKiMstHokenrui_CheckedChanged(object sender, EventArgs e)
        {
            chkMstHokenrui.Checked = chkKiMstHokenrui.Checked;
        }

        // 20160516 契約分類削除に伴うコメントアウト -del sta
        // Private Sub chkKiMstKeiyakurui_CheckedChanged(sender As Object, e As EventArgs)
        // chkMstKeiyakurui.Checked = chkKiMstKeiyakurui.Checked
        // End Sub
        // 20160516 契約分類削除に伴うコメントアウト -del end

        private void chkKiMstTokuyaku_CheckedChanged(object sender, EventArgs e)
        {
            chkMstTokuyaku.Checked = chkKiMstTokuyaku.Checked;
        }

        private void chkKiMstKasyoClaimrui_CheckedChanged(object sender, EventArgs e)
        {
            chkMstKasyoClaimrui.Checked = chkKiMstKasyoClaimrui.Checked;
        }

        private void chkKiMstHendo_CheckedChanged(object sender, EventArgs e)
        {
            chkMstHendo.Checked = chkKiMstHendo.Checked;
            chkMstHendoitiran.Checked = chkKiMstHendo.Checked;
        }

        private void chkKiMstKagititle_CheckedChanged(object sender, EventArgs e)
        {
            chkMstKagititle.Checked = chkKiMstTitle.Checked;
            // 20160801 タイトルマスタ統合処理 -add sta
            chkMstBikotitle.Checked = chkKiMstTitle.Checked;
            chkMstBikolst.Checked = chkKiMstTitle.Checked;
            chkMstGazotitle.Checked = chkKiMstTitle.Checked;
            // 20160801 タイトルマスタ統合処理 -add end
        }

        // 20160801 タイトルマスタ統合処理 -del sta
        // Private Sub chkKiMstBikotitle_CheckedChanged(sender As Object, e As EventArgs)
        // chkMstBikotitle.Checked = chkKiMstBikotitle.Checked
        // chkMstBikolst.Checked = chkKiMstBikotitle.Checked
        // End Sub
        // 20160801 タイトルマスタ統合処理 -del end

        /// <summary>
        /// イベント処理：チェックボックスチェンジ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void chkKiJisyaBase_CheckedChanged(object sender, EventArgs e)
        {
            Set_ChkBoxValue(grpJisya, chkKiJisyaBase.Checked, 1);
        }

        private void chkKiOwBase_CheckedChanged(object sender, EventArgs e)
        {
            Set_ChkBoxValue(grpOw, chkKiOwBase.Checked, 1);
        }

        private void chkKiBkBase_CheckedChanged(object sender, EventArgs e)
        {
            Set_ChkBoxValue(grpBk, chkKiBkBase.Checked, 1);
        }

        private void chkKiHyBase_CheckedChanged(object sender, EventArgs e)
        {
            Set_ChkBoxValue(grpHy, chkKiHyBase.Checked, 1);
            chkHySetubi.Checked = chkKiHySetubi.Checked;
            Set_ChkBoxValue(grpSorule, chkKiHyBase.Checked, 1);
            chkKiHySetubi.Enabled = chkKiHyBase.Checked;
            chkHySetubi.Enabled = chkKiHyBase.Checked;
            if (chkKiHyBase.Checked == false)
            {
                chkKiHySetubi.Checked = chkKiHyBase.Checked;
            }
        }

        private void chkKiHySetubi_CheckedChanged(object sender, EventArgs e)
        {
            chkHySetubi.Checked = chkKiHySetubi.Checked;
        }

        private void chkKiKysBase_CheckedChanged(object sender, EventArgs e)
        {
            Set_ChkBoxValue(grpKys, chkKiKysBase.Checked, 1);
        }

        private void chkKiKyBase_CheckedChanged(object sender, EventArgs e)
        {
            Set_ChkBoxValue(grpKy, chkKiKyBase.Checked, 1);
        }

        private void optKiKagi_CheckedChanged(object sender, EventArgs e)
        {
            optHyKagi.Checked = optKiHyKagi.Checked;
            optKyKagi.Checked = optKiKyKagi.Checked;
        }

        private void chkKiGyCyukaiBase_CheckedChanged(object sender, EventArgs e)
        {
            chkGyCyukaiBase.Checked = chkKiGyCyukaiBase.Checked;
            chkGyCyukaiKoza.Checked = chkKiGyCyukaiBase.Checked;
            chkGyCyukaiMemo.Checked = chkKiGyCyukaiBase.Checked;
        }

        private void chkKiGySyuzenBase_CheckedChanged(object sender, EventArgs e)
        {
            chkGySyuzenBase.Checked = chkKiGySyuzenBase.Checked;
            chkGySyuzenKoza.Checked = chkKiGySyuzenBase.Checked;
            chkGySyuzenMemo.Checked = chkKiGySyuzenBase.Checked;
        }

        private void chkKiGyLifelineBase_CheckedChanged(object sender, EventArgs e)
        {
            chkGyLifelineBase.Checked = chkKiGyLifelineBase.Checked;
        }

        private void chkKiGyHokenBase_CheckedChanged(object sender, EventArgs e)
        {
            chkGyHokenBase.Checked = chkKiGyHokenBase.Checked;
            chkGyHokenKoza.Checked = chkKiGyHokenBase.Checked;
            chkGyHokenMemo.Checked = chkKiGyHokenBase.Checked;
        }

        private void chkKiGyYatinhosyoBase_CheckedChanged(object sender, EventArgs e)
        {
            chkGyYatinhosyoBase.Checked = chkKiGyYatinhosyoBase.Checked;
            chkGyYatinhosyoMemo.Checked = chkKiGyYatinhosyoBase.Checked;
        }

        private void chkKiGySisetuBase_CheckedChanged(object sender, EventArgs e)
        {
            chkGySisetuBase.Checked = chkKiGySisetuBase.Checked;
        }

        private void chkKiGySekoBase_CheckedChanged(object sender, EventArgs e)
        {
            chkGySekoBase.Checked = chkKiGySekoBase.Checked;
        }

        private void chkKiSqBase_CheckedChanged(object sender, EventArgs e)                 // 20160519 請求情報追加
        {
            Set_ChkBoxValue(grpSq, chkKiSqBase.Checked, 1);
        }

        private void chkKiClaimBase_CheckedChanged(object sender, EventArgs e)           // 20160524 クレーム情報移行処理実装
        {
            Set_ChkBoxValue(grpClaim, chkKiClaimBase.Checked, 1);
        }

        private void chkKiSyskanriBase_CheckedChanged(object sender, EventArgs e)     // 20160616 初期設定情報移行処理実装
        {
            Set_ChkBoxValue(grpSyskanri, chkKiSyskanriBase.Checked, 1);
        }

        private void chkKiSzenBase_CheckedChanged(object sender, EventArgs e)             // 20160621 修繕関連移行処理追加
        {
            Set_ChkBoxValue(grpSzen, chkKiSzenBase.Checked, 1);
        }

        private void chkKiRendoBase_CheckedChanged(object sender, EventArgs e)           // 20160720 連動情報構築
        {
            Set_ChkBoxValue(grpRendo, chkKiRendoBase.Checked, 1);
        }

        /// <summary>
        /// イベント処理：既存用→汎用用中間ファイルへのコピー処理(開発用処理)　　　　　　
        /// 20160912 既存用→汎用用中間ファイルへのコピー処理
        /// (開発用処理)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void btnExistMidToBaseMid_Click(object sender, EventArgs e)
        {

            int tmp_cnt = 0;
            var list_cv = new List<string>();                                          // 移行項目取得
            list_cv = Get_ListCVChkitem(true);

            // 移行項目チェック
            if (list_cv.Count == 0)
            {
                MessageBox.Show("移行項目を選択して下さい。");
                return;
            }

            // 格納先フォルダチェック
            if (EtcMethod.Chk_DirExist(txtExistMidToBaseMid.Text) == false)
            {
                MessageBox.Show("中間ファイル格納先フォルダが存在しません。再度選択して下さい。");
                return;
            }

            CommonModule.MsgResult = MessageBox.Show("既存→汎用中間ファイルへのコピー処理を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (CommonModule.MsgResult == DialogResult.No)
            {
                return;
            }

            // ---------------------------------------------------------------
            // 汎用用中間ファイルと既存用中間ファイルの照合用仮テーブル作成
            // ---------------------------------------------------------------
            // 仮テーブル初期化
            string midfileinfo_dropqry = DBQuery.Qry_DropInfo(MidFileInfoModule.MIDFILEINFO_DBNAME, true);
            DBExec.Exec_NonQuery(sqlcnnv10, midfileinfo_dropqry, ref tmp_cnt);

            // 仮テーブル新規作成
            string midfileinfo_createqry = MidFileInfoModule.Get_MidFileInfo_CreateQry();
            DBExec.Exec_NonQuery(sqlcnnv10, midfileinfo_createqry, ref tmp_cnt);

            // 挿入
            bool normalflg = true;
            string midfileinfo_insertqry = MidFileInfoModule.Get_MidFileInfo_InsertQry();
            if (!string.IsNullOrEmpty(midfileinfo_insertqry))
            {
                normalflg = DBExec.Exec_NonQuery(sqlcnnv10, midfileinfo_insertqry, ref tmp_cnt);
            }
            if (normalflg == false)
            {
                MessageBox.Show("仮テーブル作成中にエラーが発生しました。");
                return;
            }

            // -----------------------------------------------------------------------
            // 作成した照合用仮テーブルからヘッダーに関する情報を取得してオブジェクトへ格納
            // -----------------------------------------------------------------------
            Set_MidHeaderInfoToObj();

            // -----------------------------------------------------------------------
            // 仮テーブルから紐付情報を取得して汎用用中間ファイルを作成する
            // -----------------------------------------------------------------------
            string errstr = "";
            normalflg = Set_ExistMidFile_Copy(ref errstr, list_cv);
            if (normalflg)
            {
                MessageBox.Show("既存用→汎用用中間ファイルへのコピー処理が完了しました。");
            }
            else
            {
                MessageBox.Show("コピー処理中にエラーが発生しました。");
            }

        }

        #endregion

        #region 接続情報設定

        /// <summary>
        /// 接続情報：初期値→コントロール
        /// </summary>
        /// <param name="cnnv7"></param>
        /// <param name="cnnv10"></param>
        /// <remarks></remarks>
        public void Set_DefConInfo_To_Control(ref Model.DefSQLConnection cnnv10)
        {
            // 10
            txtV10Server.Text = cnnv10.ServerName;
            txtV10Catalog.Text = cnnv10.InitialCatalog;
            // Me.txtV10Networklib.Text = cnnv10.NetworkLibrary
            txtV10User.Text = cnnv10.User;
            txtV10Pass.Text = cnnv10.Pass;

            // 共通
            txtTimeOut.Text = cnnv10.TimeOut;

        }

        /// <summary>
        /// 接続情報：コントロール→モデル
        /// </summary>
        /// <param name="model"></param>
        /// <param name="flg"></param>
        /// <remarks></remarks>
        public void Set_Control_To_ConModel(ref Model.DefSQLConnection model)
        {
            // 10
            model.ServerName = txtV10Server.Text;
            model.InitialCatalog = txtV10Catalog.Text;
            // model.NetworkLibrary = Me.txtV10Networklib.Text
            model.User = txtV10User.Text;
            model.Pass = txtV10Pass.Text;

            // 共通
            model.TimeOut = txtTimeOut.Text;

        }

        /// <summary>
        /// DB接続情報取得 (2回目以降起動時)  
        /// </summary>
        /// <remarks></remarks>
        private bool Set_DBExistConInfo()
        {

            bool rtn = true;

            // 保管されている接続情報ファイルの読込
            string exepath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string exedir = Path.GetDirectoryName(exepath);
            string coninfodirpath = EtcMethod.Set_Path(exedir, CommonModule.DIR_INI_NAME);
            string coninfofilepathV7 = EtcMethod.Set_Path(coninfodirpath, CommonModule.FILE_V7_CONNAME);
            string coninfofilepath10 = EtcMethod.Set_Path(coninfodirpath, CommonModule.FILE_10_CONNAME);

            // どちらかのファイルが存在しない(初回起動またはファイル名が編集されている)場合
            if (EtcMethod.Chk_FileExist(coninfofilepathV7) == false || EtcMethod.Chk_FileExist(coninfofilepath10) == false)
            {
                rtn = false;
                return rtn;
            }

            var xmlreader = default(System.Xml.XmlReader);
            string item = "";
            string data = "";
            string svname = "";
            string dbname = "";
            string username = "";
            string password = "";
            string encryptedpassword = "";
            string passkey = "tRwmj5U4";
            Model.DefSQLConnection conmodel = null;

            for (int cntii = 0; cntii <= 1; cntii++)
            {

                switch (cntii)
                {
                    case 0:
                        {
                            xmlreader = System.Xml.XmlReader.Create(coninfofilepathV7);
                            conmodel = fstmodelv7;
                            break;
                        }
                    case 1:
                        {
                            xmlreader = System.Xml.XmlReader.Create(coninfofilepath10);
                            conmodel = fstmodelv10;
                            break;
                        }
                }

                while (xmlreader.Read())
                {

                    if (xmlreader.NodeType == System.Xml.XmlNodeType.Element)
                    {

                        // データ取得
                        item = xmlreader.LocalName;
                        data = xmlreader.ReadString();

                        // それぞれの要素で分岐
                        switch (item ?? "")
                        {
                            case "ServerName":
                                {
                                    svname = data;
                                    break;
                                }
                            case "InitialCatalog":
                                {
                                    dbname = data;
                                    break;
                                }
                            case "UserID":
                                {
                                    username = data;
                                    break;
                                }
                            case "Password":
                                {
                                    break;
                                }
                            // password = data
                            case "EncryptedPassword":
                                {
                                    encryptedpassword = data;
                                    break;
                                }
                        }

                    }

                }

                // パスワード複合化
                password = Decrypt(encryptedpassword, passkey);

                conmodel.ServerName = svname;
                conmodel.InitialCatalog = dbname;
                conmodel.User = username;
                conmodel.Pass = password;

                switch (cntii)
                {
                    case 0:
                        {
                            fstmodelv7 = (Model.DefSQLConnection)conmodel;
                            break;
                        }
                    case 1:
                        {
                            fstmodelv10 = (Model.DefSQLConnection)conmodel;
                            break;
                        }
                }

                xmlreader.Close();

            }

            return rtn;

        }

        /// <summary>
        /// DB接続情報作成 (2回目以降起動時用に退避させておく)
        /// </summary>
        /// <remarks>
        /// ・接続情報パスワード保存処理
        /// </remarks>
        private bool MakeFile_DBConInfo(Model.DefSQLConnection conmodel)
        {

            string svname = Conversions.ToString(conmodel.ServerName);
            string catalog = Conversions.ToString(conmodel.InitialCatalog);
            string user = Conversions.ToString(conmodel.User);
            string pass = Conversions.ToString(conmodel.Pass);

            // 接続情報保存ファイル格納先取得
            string exepath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string exedir = Path.GetDirectoryName(exepath);
            string coninfodirpath = EtcMethod.Set_Path(exedir, CommonModule.DIR_INI_NAME);
            string coninfofilepath = "";

            coninfofilepath = EtcMethod.Set_Path(coninfodirpath, CommonModule.FILE_10_CONNAME);

            // Windows認証時にUserにデフォルト値を入れて保存する処理を追加
            // ※空のままだとXML作成の際に改行が混入するためこれを防ぐ
            if (string.IsNullOrEmpty(user.Trim()))
            {
                user = "sa";
            }

            string passkey = "tRwmj5U4";
            string encryptedpassword = Encrypt(pass, passkey);

            string strXml = "<?xml version='1.0'?>" + "<coninfo>" + "<!--サーバー名、カタログ名、ユーザー名、パスワード-->" + "<ServerName>" + svname + "</ServerName>" + "<InitialCatalog>" + catalog + "</InitialCatalog>" + "<UserID>" + user + "</UserID>" + "<Password />" + "<EncryptedPassword>" + encryptedpassword + "</EncryptedPassword>" + "</coninfo>";








            var xmlDoc = new System.Xml.XmlDocument();

            // 文字列からDOMドキュメントを生成
            xmlDoc.LoadXml(strXml);

            // 20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg sta
            // Try
            // '作成したDOMドキュメントをファイルに保存
            // xmlDoc.Save(coninfofilepath)
            // Catch ex As System.Xml.XmlException
            // 'XMLによる例外をキャッチ
            // Console.WriteLine(ex.Message)
            // Catch ex As Exception
            // 'その他の例外をキャッチ
            // Console.WriteLine(ex.Message)
            // End Try

            try
            {
                // 作成したDOMドキュメントをファイルに保存
                xmlDoc.Save(coninfofilepath);
            }

            // アクセス拒否
            // Catch ex As UnauthorizedAccessException
            // MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
            // "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
            // "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            // Return False
            // その他の例外(アクセス拒否を含む)
            catch (Exception ex)
            {
                CommonModule.MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" + Constants.vbCrLf + "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return false;

            }

            return true;
            // 20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg end
        }

        /// <summary>
        /// 10DB接続情報取得 
        /// </summary>
        /// <remarks></remarks>
        private void Set_10DBConInfo()
        {

            string xmlfilepath = "";

            // XMLファイルパスを取得
            string tmp_filepath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            tmp_filepath = EtcMethod.Set_Path(tmp_filepath, CommonModule.COMPANY_NAME);
            tmp_filepath = EtcMethod.Set_Path(tmp_filepath, CommonModule.PRODUCT_NAME);
            tmp_filepath = EtcMethod.Set_Path(tmp_filepath, string.Format("{0}.sql.Main.xml", CommonModule.PRODUCT_NAME));
            xmlfilepath = tmp_filepath;

            // ファイル有無チェック
            if (EtcMethod.Chk_FileExist(xmlfilepath) == false)
            {
                return;
            }

            // ファイル読込
            var xmlreader = System.Xml.XmlReader.Create(xmlfilepath);
            string item = "";
            string data = "";
            string svname = "";
            string dbname = "";
            string username = "";
            string password = "";
            string encryptedpassword = "";
            string passkey = "tRwmj5U4";
            string tmp_pass = "";

            while (xmlreader.Read())
            {
                if (xmlreader.NodeType == System.Xml.XmlNodeType.Element)
                {

                    // データ取得
                    item = xmlreader.LocalName;
                    data = xmlreader.ReadString();

                    // それぞれの要素で分岐
                    switch (item ?? "")
                    {
                        case "ServerName":
                            {
                                svname = data;
                                break;
                            }
                        case "InitialCatalog":
                            {
                                dbname = data;
                                break;
                            }
                        case "UserID":
                            {
                                username = data;
                                break;
                            }
                        case "EncryptedPassword":
                            {
                                encryptedpassword = data;
                                break;
                            }
                    }
                }
            }

            // パスワード複合化
            password = Decrypt(encryptedpassword, passkey);

            // モデルへ格納
            {
                ref var withBlock = ref fstmodelv10;
                withBlock.ServerName = svname;
                withBlock.InitialCatalog = dbname;
                withBlock.User = username;
                withBlock.Pass = password;
            }

            xmlreader.Close();

        }

        /// <summary>
        /// DESで暗号化します。<br/> 
        /// </summary>
        public string Encrypt(string target, string key)
        {
            byte[] targetBytes = System.Text.Encoding.UTF8.GetBytes(target);

            var des = new System.Security.Cryptography.DESCryptoServiceProvider();
            byte[] keyBytes = System.Text.Encoding.UTF8.GetBytes(key);
            des.Key = resizeBytesArray(keyBytes, des.Key.Length);
            des.IV = resizeBytesArray(keyBytes, des.IV.Length);

            var outStream = new MemoryStream();
            var desencrypt = des.CreateEncryptor();
            var cryptStream = new System.Security.Cryptography.CryptoStream(outStream, desencrypt, System.Security.Cryptography.CryptoStreamMode.Write);

            cryptStream.Write(targetBytes, 0, targetBytes.Length);
            cryptStream.FlushFinalBlock();
            byte[] outBytes = outStream.ToArray();

            cryptStream.Close();
            outStream.Close();

            return Convert.ToBase64String(outBytes);
        }

        /// <summary>
        /// DESで復号化します。<br/> 
        /// </summary>
        public string Decrypt(string encryptedTarget, string key)
        {
            byte[] encryptedBytes = Convert.FromBase64String(encryptedTarget);

            var des = new System.Security.Cryptography.DESCryptoServiceProvider();
            byte[] keyBytes = System.Text.Encoding.UTF8.GetBytes(key);
            des.Key = resizeBytesArray(keyBytes, des.Key.Length);
            des.IV = resizeBytesArray(keyBytes, des.IV.Length);

            var inStream = new MemoryStream(encryptedBytes);
            var desdecrypt = des.CreateDecryptor();
            var cryptStream = new System.Security.Cryptography.CryptoStream(inStream, desdecrypt, System.Security.Cryptography.CryptoStreamMode.Read);

            var outStream = new StreamReader(cryptStream, System.Text.Encoding.UTF8);
            string result = outStream.ReadToEnd();

            outStream.Close();
            cryptStream.Close();
            inStream.Close();

            return result;
        }

        private byte[] resizeBytesArray(byte[] bytes, int newSize)
        {
            var newBytes = new byte[newSize];
            if (bytes.Length <= newSize)
            {
                int i;
                var loopTo = bytes.Length - 1;
                for (i = 0; i <= loopTo; i++)
                    newBytes[i] = bytes[i];
            }
            else
            {
                int pos = 0;
                int i;
                var loopTo1 = bytes.Length - 1;
                for (i = 0; i <= loopTo1; i++)
                {
                    newBytes[pos] = (byte)(newBytes[pos] ^ bytes[i]);
                    pos += 1;
                    if (pos >= newBytes.Length)
                    {
                        pos = 0;
                    }
                }
            }
            return newBytes;
        }

        #endregion

        #region コンバート実行処理(既存ユーザ用, 汎用共通)

        /// <summary>
        /// コンバート処理 
        /// </summary>
        /// <remarks></remarks>
        private void Cnv_Main(List<string> list_cv)
        {

            int rowcnt = 0;                                       // 確認メモ：Exec_NonQueryの第3引数の為用意(実際は未使用)
            var obj_pgb = new ProgressBarManager();
            string errstr = "";
            bool normalflg = true;

            // --------------------------------------------------
            // ログ状況初期設定・出力
            // --------------------------------------------------
            // 初期設定
            Set_LogInit(true);
            Set_LogInit(false);
            // ログ出力
            string tmp_sql_sta = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, CommonModule.LOG_SYORIKOMK_CV_STA), false);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_sta, ref rowcnt);
            // 状況出力
            // 20161018 UI最終確認での修正 -add sta
            // 進捗状況出力内容初期化
            txtTotalSituation.Text = "";
            txtPartialSituation.Text = "";
            // 20161018 UI最終確認での修正 -add end
            Set_Situation(CommonModule.SITUATION_CV_STA, 0);
            // 20161017 進捗表示ラベル初期表示修正 -add sta
            lblTotalSituation.Text = CommonModule.SITUATION_CV_BFRUN;  // 全体進捗ラベル初期化
            // 20161104 2回目実行時の画面初期化処理修正 -add
            lblCVItem.Text = "";                          // 個別進捗ラベル初期化
            obj_com.ProgressOutPut(0, 1, true);         // 全体進捗進捗率初期化
            obj_com.ProgressOutPut(0, 1);               // 個別進捗進捗率初期化
            // 20161017 進捗表示ラベル初期表示修正 -add end
            // --------------------------------------------------
            // 設定値取得
            // --------------------------------------------------             
            bool midfileflg = chkMiddleFile.Checked;            // 中間ファイル作成フラグ(開発用)
            bool himoflg = chkRelation.Checked;                 // 紐付作業実行フラグ
            string tmp_cntstr = txtLogOutputCnt.Text.Trim();         // ログ出力件数設定値取得
            int tmp_cntint = 0;                                   // ログ出力件数(作業用)
            if (!int.TryParse(tmp_cntstr, out tmp_cntint))
            {
                tmp_cntint = 1;
            }
            CommonModule.Log_OutputCnt = tmp_cntint;


            // --------------------------------------------------
            // ハッシュテーブル初期化 '20160711 オブジェクトインスタンス化の修正
            // --------------------------------------------------
            Ini_HashTable();


            // ================================================== メイン処理 ================================================== s
            // --------------------------------------------------
            // 中間ファイル処理
            // --------------------------------------------------
            Chg_LeftProgress("dcfilewrite", 1);
            pnlMenuDatacv.Refresh();       // kekeke

            // ----- 全体用・個別用プログレスバー初期化 -----
            obj_pgb.pgbInitPart(0);
            obj_pgb.pgbInitTotal(0);
            tabPageJikko.Refresh();
            pgbtotalcnt = list_cv.Count;
            obj_pgb.pgbInitTotal(pgbtotalcnt);

            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
            {

                // 汎用用中間ファイルから既存用中間ファイルへのコピー処理
                normalflg = Set_BaseMidFileToExistingMidFile(list_cv, ref errstr);
                if (normalflg == false)
                {
                    CommonModule.MsgResult = MessageBox.Show(errstr, "注意", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    goto skiplabel;
                }
            }



            // --------------------------------------------------
            // DB書込実行前準備 (作業用ファイルの仮TBL作成)
            // --------------------------------------------------                               
            if (CommonModule.CancelFlg == false)
            {
                normalflg = Before_DBWrite(list_cv);
                if (normalflg == false)
                {
                    CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_ERR_MAKE_CVDB, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
            }


            // --------------------------------------------------
            // 中間ファイルチェック
            // --------------------------------------------------
            if (CommonModule.CancelFlg == false & CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
            {
                normalflg = Chk_MidFileChk_Main(list_cv);
                if (normalflg == false)
                {
                    CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_ERR_CHK_MID, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
            }


            // --------------------------------------------------
            // 紐付設定画面呼出
            // --------------------------------------------------
            Chg_LeftProgress("dcrelation", 1);
            pnlMenuDatacv.Refresh();       // kekeke

            if (CommonModule.CancelFlg == false & himoflg)
            {
                // 選択された移行項目に関連する紐付情報を起動させるために移行項目のチェック値を引数として渡す
                string cvitemtorel = Set_CVItemChkToCmdLine(list_cv);
                // 20160720 紐付設定ファイル出力機能の追加 -add
                relitemstr = cvitemtorel.Trim();
                // 移行項目と関連する紐付項目が存在する場合は紐付ツールを起動
                if (!string.IsNullOrEmpty(cvitemtorel.Trim()))
                {
                    // 20160707 本体と紐付ツールの中断を同期させる処理の追加 -chg sta
                    // Call Me.Rel_ExeCall(cvitemtorel)

                    // ----- 全体用・個別用プログレスバー初期化 -----
                    obj_pgb.pgbInitPart(0);
                    obj_pgb.pgbInitTotal(0);
                    obj_com.ProgressOutPut(0, pgbtotalcnt);
                    obj_com.ProgressOutPut(0, pgbtotalcnt, true);
                    tabPageJikko.Refresh();
                    pgbtotalcnt = 1;
                    obj_pgb.pgbInitPart(pgbtotalcnt);
                    obj_pgb.pgbInitTotal(pgbtotalcnt);
                    lblTotalSituation.Text = CommonModule.SITUATION_REL_RUN;

                    // ◎紐付設定画面呼出
                    CommonModule.CancelFlg = Rel_ExeCall(cvitemtorel);
                    // 20160707 本体と紐付ツールの中断を同期させる処理の追加 -chg end
                }
            }


            // --------------------------------------------------
            // DB書込処理
            // --------------------------------------------------
            Chg_LeftProgress("dcconv", 1);
            pnlMenuDatacv.Refresh();       // kekeke
            if (CommonModule.CancelFlg == false)
            {

                // ----- 全体用・個別用プログレスバー初期化 -----
                obj_pgb.pgbInitPart(0);
                obj_pgb.pgbInitTotal(0);
                obj_com.ProgressOutPut(0, pgbtotalcnt);
                obj_com.ProgressOutPut(0, pgbtotalcnt, true);
                tabPageJikko.Refresh();
                pgbtotalcnt = list_cv.Count;
                obj_pgb.pgbInitTotal(pgbtotalcnt);
                lblTotalSituation.Text = CommonModule.SITUATION_CV_RUN;

                // ◎DB書き込み処理
                Read_MidFile_WriteDB(list_cv);
            }


            // 20160929 データ調整用メソッドの作成 -add sta
            // --------------------------------------------------
            // データ調整 
            // --------------------------------------------------
            if (CommonModule.CancelFlg == false)
            {
                // ◎データ調整処理
                DataCond();
            }
        // 20160929 データ調整用メソッドの作成 -add end

        // ================================================== メイン処理 ================================================== e


        // --------------------------------------------------
        // 終了処理
        // --------------------------------------------------
        skiplabel:
            ;


            // ログ出力
            string tmp_sql_end = "";
            if (CommonModule.CancelFlg)
            {
                tmp_sql_end = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, CommonModule.LOG_SYORIKOMK_CV_STOP), false);
            }
            else
            {
                tmp_sql_end = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, CommonModule.LOG_SYORIKOMK_CV_END), false);
            }
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_end, ref rowcnt);

            // 状況出力
            string tmp_situationstr = "";
            if (CommonModule.CancelFlg)
            {
                tmp_situationstr = CommonModule.SITUATION_STOP;
            }
            else
            {
                tmp_situationstr = CommonModule.SITUATION_CV_END;
            }

            Set_Situation(tmp_situationstr, 0);
            lblTotalSituation.Text = tmp_situationstr;

        }

        /// <summary>
        /// DB書込処理 
        /// </summary>
        /// <param name="list_cv"></param>
        /// <remarks></remarks>
        private void Read_MidFile_WriteDB(List<string> list_cv)
        {

            int pgbcnt = 0;
            string tmp_totalsituation = "";
            string tmp_partsituation = "";
            var obj_pgb = new ProgressBarManager();
            int rowcnt = 0;


            // --------------------------------------------------
            // ログ/状況出力
            // --------------------------------------------------
            // ログ出力
            string tmp_sql_sta = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, CommonModule.LOG_SYORIKOMK_DBCVSTA), false);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_sta, ref rowcnt);

            // 状況出力
            tmp_partsituation = CommonModule.SITUATION_WRITESTA;
            Set_Situation(tmp_partsituation, 1);


            // --------------------------------------------------
            // 設定値取得
            // --------------------------------------------------
            CommonModule.InitDBFlg = true;                 // 新規コンバートフラグ
            CommonModule.OverWriteDBFlg = chkOverWrite.Checked;        // 上書きコンバートフラグ


            // --------------------------------------------------
            // コンバート実行前準備
            // --------------------------------------------------
            Setting_BeforeCV();


            // --------------------------------------------------
            // 紐付項目の挿入処理実行
            // --------------------------------------------------
            bool relcvnormalflg = Set_RelItem();
            if (relcvnormalflg == false)
            {
                CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_ERR_RELCV, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            // 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -add sta
            // --------------------------------------------------
            // ダミー用の口座情報作成
            // --------------------------------------------------
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
            {
                if (Set_DummyBankData() == false)
                {
                    CommonModule.MsgResult = MessageBox.Show(CommonModule.MSG_ERR_DUMMYBANKCV, "確認", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
            }
            // 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -add end

            // ==================================================
            // メイン処理 (処理開始)
            // ==================================================
            foreach (var cvitem in list_cv)
            {

                // 中断処理
                Application.DoEvents();
                if (CommonModule.CancelFlg)
                {
                    break;
                }

                // グループボックス/チェックボックス名取得
                string[] tmp_arry = cvitem.Split('-');
                string syorigrp = tmp_arry[0];
                string syoriitem = tmp_arry[1];

                // 処理Repositoryの生成 (DB登録用)
                var obj_rep = Get_ObjRep_DB(syoriitem);

                // ----- デバッグ用処理 ----- sta
                // Dim list As New List(Of String) From {"送信設定athome情報"}
                // Dim list As New List(Of String) From {"送信設定自社web情報", "送信設定HOMES情報", "送信設定athome情報", "送信設定SUUMO情報"}
                // Dim list As New List(Of String) From {"送金ルール基本情報", "送金ルール送金先情報", "送金ルール入金項目情報", "送金ルール控除項目情報"}
                // 'Dim list As New List(Of String) From {"広告補足自社web情報", "広告補足HOMES情報", "広告補足athome情報", "広告補足SUUMO情報"}
                // Dim list As New List(Of String) From {"送金ルール控除項目情報"}
                // If list.Contains(syoriitem) = False Then
                // obj_rep = Nothing
                // End If
                // ----- デバッグ用処理 ----- end

                // 処理開始
                if (obj_rep is not null)
                {
                    // 状況出力
                    tmp_totalsituation = syoriitem + CommonModule.INDENT_1 + CommonModule.SITUATION_WRITESTA;
                    Set_Situation(tmp_totalsituation, 0);
                    lblCVItem.Text = syorigrp + " ： " + syoriitem;

                    // テーブル初期化
                    if (CommonModule.InitDBFlg)
                    {
                        string sql_where = Conversions.ToString(Interaction.IIf(Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(hash_deltblqrywhere[syoriitem], null, false)), "", hash_deltblqrywhere[syoriitem]));
                        // 20160720 連動情報構築 -chg sta
                        // If syorigrp <> "初期設定" Then
                        // Me.Initialize_Table(Hash_TblName_JpToAlpha(syoriitem), sql_where)
                        // End If
                        switch (syorigrp ?? "")
                        {
                            case "初期設定":
                                {
                                    break;
                                }
                            // 初期化しない
                            case "物件データ連動情報":
                                {
                                    string tmp_chgsyoriitem = syoriitem;
                                    if (syoriitem == "広告補足自社web情報" | syoriitem == "広告補足HOMES情報" | syoriitem == "広告補足athome情報" | syoriitem == "広告補足SUUMO情報")
                                    {
                                        tmp_chgsyoriitem = "広告補足情報";
                                    }
                                    else if (syoriitem == "送信設定自社web情報" | syoriitem == "送信設定HOMES情報" | syoriitem == "送信設定athome情報" | syoriitem == "送信設定SUUMO情報")
                                    {
                                        tmp_chgsyoriitem = "送信設定情報";
                                    }
                                    Initialize_Table(Conversions.ToString(CommonModule.Hash_TblName_JpToAlpha[tmp_chgsyoriitem]), sql_where);
                                    break;
                                }
                            // 20161028 物件/部屋鍵取得方法修正 -add sta
                            case "各マスタ情報":
                                {
                                    if (syoriitem == "鍵タイトルマスタ")
                                    {
                                        switch (CommonModule.CNVNO)
                                        {
                                            case (int)CommonModule.ConvertTypes._汎用:
                                                {
                                                    break;
                                                }
                                            // 初期化しない
                                        }
                                    }
                                    else
                                    {
                                        Initialize_Table(Conversions.ToString(CommonModule.Hash_TblName_JpToAlpha[syoriitem]), sql_where);
                                        // 20161028 物件/部屋鍵取得方法修正 -add end
                                    }

                                    break;
                                }

                            default:
                                {
                                    Initialize_Table(Conversions.ToString(CommonModule.Hash_TblName_JpToAlpha[syoriitem]), sql_where);
                                    break;
                                }
                        }
                        // 20160720 連動情報構築 -chg end
                    }

                    // 中間ファイル読込→DB書込          
                    int midrowcnt = 0;                // 中間ファイル件数                                                        
                    int cvrowcnt = 0;                 // コンバート件数
                    int conditioncnt = 0;             // 調整件数 (文字列切り捨て等)
                    string tmp_str = "";
                    bool normalflg = true;             // 正常処理判定

                    // ----- 要チェック ----- sta
                    // →必要？
                    // DB書込前の個別処理
                    switch (syorigrp ?? "")
                    {
                        case "契約情報":
                            {
                                break;
                            }
                            // 革命10用の契約行No、更新No、契約改定Noを作成した仮テーブル作成
                    }
                    // ----- 要チェック ----- end

                    // 20161124 物件部屋所有者情報有無による所有者情報のデータ調整処理を追加 -add sta
                    if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                    {
                        if (syoriitem == "送金ルール基本情報")
                        {
                            int tmp_cnt = 0;

                            // ---------------------------------------------------------------------
                            // 汎用コンバートで送金ルール基本情報を移行する際は
                            // 物件部屋所有者情報を成形(データ調整)してから行う              '20161124 レビュー結果：sorule_guidを使う為
                            // ※送金ルール作成には所有者情報にあるsorule_guidを使用するため
                            // ---------------------------------------------------------------------

                            // 物件・部屋所有者情報が重複して存在する場合、区分所有を優先して移行するため物件所有者情報を削除する
                            string sql_bksyodelete = DataCondModule.Get_UseQry_BkSyoDelete();
                            DBExec.Exec_NonQuery(sqlcnnv10, sql_bksyodelete, ref tmp_cnt);

                            // 物件・部屋基本情報が存在し、所有者情報のレコードが存在しないデータに対して仮レコードを作成する
                            var sql_bkhysyoinsert = DataCondModule.Get_List_UseQry_BkSyoInsert();
                            foreach (var tmp_sql in sql_bkhysyoinsert)
                                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
                        }
                    }
                    // 20161124 物件部屋所有者情報有無による所有者情報のデータ調整処理を追加 -add end

                    // メイン処理
                    normalflg = Conversions.ToBoolean(obj_rep.Set_Vari(sqlcnnv10, syorigrp, syoriitem, ref midrowcnt, ref cvrowcnt, ref conditioncnt));

                    // 状況出力用情報格納                                          
                    if (normalflg)
                    {
                        tmp_str = tmp_str + Constants.vbCrLf + CommonModule.INDENT_2 + CommonModule.SITUATION_CVITEMTOTALCNT + CommonModule.INDENT_1 + midrowcnt.ToString();
                        tmp_str = tmp_str + Constants.vbCrLf + CommonModule.INDENT_2 + CommonModule.SITUATION_CVITEMCVCNT + CommonModule.INDENT_1 + cvrowcnt.ToString();
                        tmp_str = tmp_str + Constants.vbCrLf + CommonModule.INDENT_2 + CommonModule.SITUATION_CVITEMNOTCVCNT + CommonModule.INDENT_1 + (midrowcnt - cvrowcnt).ToString();
                        tmp_str = tmp_str + Constants.vbCrLf + CommonModule.INDENT_2 + CommonModule.SITUATION_CVITEMCONDCNT + CommonModule.INDENT_1 + conditioncnt.ToString();
                    }

                    // ログ出力
                    string totalcntstr = CommonModule.SITUATION_CVITEMTOTALCNT.Trim() + " " + midrowcnt.ToString();
                    string cvcntstr = CommonModule.SITUATION_CVITEMCVCNT.Trim() + " " + cvrowcnt.ToString();
                    string notcvcntstr = CommonModule.SITUATION_CVITEMNOTCVCNT.Trim() + " " + (midrowcnt - cvrowcnt).ToString();
                    string condcntstr = CommonModule.SITUATION_CVITEMCONDCNT.Trim() + " " + conditioncnt.ToString();
                    string endkbnstr = Chg_FlgToStr(normalflg, 2);
                    string outstr = endkbnstr + " " + totalcntstr + "/" + cvcntstr + "/" + notcvcntstr + "/" + condcntstr;
                    string tmp_sql_itemsta = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(41, syoriitem, "-", outstr), false);
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_itemsta, ref rowcnt);

                    // 状況出力
                    tmp_partsituation = CommonModule.INDENT_1 + Chg_FlgToStr(normalflg, 1) + syoriitem + tmp_str;
                    tmp_totalsituation = syoriitem + CommonModule.INDENT_1 + CommonModule.SITUATION_WRITEEND;
                    Set_Situation(tmp_totalsituation, 0);
                    Set_Situation(tmp_partsituation, 1);
                }

                // ----- 全体用プログレスバー更新 -----
                pgbcnt = pgbcnt + 1;
                obj_pgb.pgbsettingTotal(pgbcnt);      // 全体進捗
                obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt, true);
                // 20161018 個別進捗プログレスバーの表示修正 -add sta
                // ----- 個別用プログレスバー更新 -----
                int pgbmax = pgbpartial.Maximum;
                if (pgbmax == 0)
                {
                    pgbmax = 1;
                    obj_pgb.pgbInitPart(pgbmax);
                    obj_pgb.pgbsettingPart(pgbmax);
                    obj_com.ProgressOutPut(pgbmax, pgbmax);
                }
                // 20161018 個別進捗プログレスバーの表示修正 -add end
            }


            // --------------------------------------------------
            // ログ/状況出力
            // --------------------------------------------------
            // ログ出力
            string tmp_sql_end = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, CommonModule.LOG_SYORIKOMK_DBCVEND), false);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_end, ref rowcnt);

            // 状況出力
            tmp_partsituation = CommonModule.SITUATION_WRITEEND;
            Set_Situation(tmp_partsituation, 1);

        }

        /// <summary>
        /// DB書込処理前準備
        /// </summary>
        /// <remarks></remarks>
        private void Setting_BeforeCV()
        {

            int rowcnt = 0;


            // --------------------------------------------------
            // 革命10DBテーブル名を取得して削除対象
            // (DELETE構文のWHERE句)をハッシュテーブルへ格納
            // --------------------------------------------------
            // 契約者照合用カナ情報のように契約者以外に家主や業者など他の情報も含まれている場合は、初期化時に削除対象を指定する必要がある
            // 契約者照合用カナを移行する際は、登録されているデータのうち、区分=契約者を削除する必要がある
            hash_deltblqrywhere = new SafeDictionary<string, string>();
            Set_DelTablename_qrywhere(ref hash_deltblqrywhere);


            // --------------------------------------------------
            // 金融機関・支店情報を取得して
            // ハッシュテーブルへ格納(マスタ参照用)
            // --------------------------------------------------
            string tmp_sql_getkinyu = DBQuery.Qry_GetKinyuInfo();
            GetBankInfo.Get_HashKinyu(sqlcnnv10, tmp_sql_getkinyu);


            // --------------------------------------------------                 
            // 実行者設定                                                   
            // 初期設定画面で設定した作業者名を履歴フィールド[history]へ登録
            // 未設定時はデフォルト値を設定 ("CONVUSER"を履歴フィールドへ登録)
            // --------------------------------------------------
            string tmp_executor = txtRecUser.Text;
            if (string.IsNullOrEmpty(tmp_executor))
            {
                tmp_executor = CommonModule.DEF_REC_USER;
            }
            CommonModule.RecUser = tmp_executor;


            // --------------------------------------------------
            // 履歴設定
            // --------------------------------------------------
            var xml = new HistorySetting();
            xml.Set_HistoryData(CommonModule.RecUser);
            string str_xml = xml.ToXmlString();
            CommonModule.DefHistory = str_xml;

        }

        /// <summary>
        /// テーブル名(日本語/削除対象)紐付けハッシュテーブル作成処理
        /// </summary>
        /// <param name="hash"></param>
        /// <remarks></remarks>
        private void Set_DelTablename_qrywhere(ref SafeDictionary<string, string> hash)
        {

            // ※固定値
            hash.Add("契約者照合用カナ情報", " sqsaki_kbn = 100 ");
            hash.Add("自社担当者情報", " logonuser_no NOT IN (99001,99999) ");
            hash.Add("備考タイトルマスタ", " memo_kbn IN (1,2,4,6,12) ");
            hash.Add("備考入力補助リストマスタ", " memo_kbn IN (12) ");
            hash.Add("画像タイトルマスタ", " gazo_kbn IN (1,2,3,8) ");        // 20160720 連動情報構築 -add
            hash.Add("送信設定自社web情報", " setting_guid <> (SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0) AND site_no = 10 ");                // 20160720 連動情報構築 -add
            hash.Add("送信設定HOMES情報", " setting_guid <> (SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0) AND site_no = 20 ");                  // 20160720 連動情報構築 -add
            hash.Add("送信設定athome情報", " setting_guid <> (SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0) AND site_no = 30 ");                 // 20160720 連動情報構築 -add
            hash.Add("送信設定SUUMO情報", " setting_guid <> (SELECT setting_guid FROM m_sendsetting WHERE setting_sortorder = 0) AND site_no = 40 ");                  // 20160720 連動情報構築 -add
            hash.Add("広告補足自社web情報", " site_no = 10 ");                // 20160720 連動情報構築 -add
            hash.Add("広告補足HOMES情報", " site_no = 20 ");                  // 20160720 連動情報構築 -add
            hash.Add("広告補足athome情報", " site_no = 30 ");                 // 20160720 連動情報構築 -add
            hash.Add("広告補足SUUMO情報", " site_no = 40 ");                  // 20160720 連動情報構築 -add
            hash.Add("送信設定基本情報", " setting_sortorder <> 0 ");         // 20160720 連動情報構築 -add

        }

        /// <summary>
        /// テーブル初期化処理
        /// </summary>
        /// <param name="tblname"></param>
        /// <param name="qrywhere"></param>
        /// <remarks></remarks>
        private void Initialize_Table(string tblname, string qrywhere)
        {

            int rowcnt = 0;

            // 初期化処理実行
            string tmp_sqldelete = DBQuery.Qry_DelInfo(tblname, qrywhere);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sqldelete, ref rowcnt);

        }

        /// <summary>
        /// DB書込Repositoryの生成
        /// </summary>
        /// <param name="chkboxtext"></param>
        /// <remarks></remarks>
        private IConv Get_ObjRep_DB(string chkboxtext)
        {

            IConv obj_rep = null;

            switch (chkboxtext ?? "")
            {

                // 各マスタ情報
                case var @case when @case == (model_cvitem.MstBus ?? ""):
                    {
                        obj_rep = new Repository.M_buskotu_Repository.SubConv();                                    // バス交通マスタ
                        break;
                    }
                case var case1 when case1 == (model_cvitem.MstBusKotu ?? ""):
                    {
                        obj_rep = new Repository.M_buskotu_stop_Repository.SubConv();                           // バス停マスタ
                        break;
                    }
                // 20161028 物件/部屋鍵取得方法修正 -chg sta
                // Case model_cvitem.MstKagititle : obj_rep = New Njc.Repository.M_kagi_title_Repository.SubConv                           '鍵タイトルマスタ
                case var case2 when case2 == (model_cvitem.MstKagititle ?? ""):
                    {
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    obj_rep = new Repository.M_kagi_title_BaseMid_Repository.SubConv();                                      // 鍵タイトルマスタ(汎用用)
                                    break;
                                }
                        }

                        break;
                    }
                // 20161028 物件/部屋鍵取得方法修正 -chg end
                case var case3 when case3 == (model_cvitem.MstKasyoClaimrui ?? ""):
                    {
                        obj_rep = new Repository.M_claim_rui_Repository.SubConv();                        // 箇所クレーム分類マスタ
                        break;
                    }
                case var case4 when case4 == (model_cvitem.MstTokuyaku ?? ""):
                    {
                        obj_rep = new Repository.M_tokuyaku_Repository.SubConv();                              // 特約マスタ
                        break;
                    }
                // Case model_cvitem.MstKeiyakurui : obj_rep = New Njc.Repository.M_ky_rui_Repository.SubConv                              '契約分類マスタ      
                case var case5 when case5 == (model_cvitem.MstHokenrui ?? ""):
                    {
                        obj_rep = new Repository.M_hoken_rui_Repository.SubConv();                             // 保険種類マスタ
                        break;
                    }
                case var case6 when case6 == (model_cvitem.MstSchool ?? ""):
                    {
                        obj_rep = new Repository.M_koku_add_Repository.SubConv();                                // 学校区マスタ
                        break;
                    }
                case var case7 when case7 == (model_cvitem.MstArea ?? ""):
                    {
                        obj_rep = new Repository.M_area_Repository.SubConv();                                      // エリアマスタ
                        break;
                    }
                case var case8 when case8 == (model_cvitem.MstHendo ?? ""):
                    {
                        obj_rep = new Repository.M_hendorule_Repository.SubConv();                                // 変動費マスタ
                        break;
                    }
                case var case9 when case9 == (model_cvitem.MstHendoitiran ?? ""):
                    {
                        obj_rep = new Repository.M_hendorule_itiran_Repository.SubConv();                   // 変動費一覧マスタ
                        break;
                    }
                case var case10 when case10 == (model_cvitem.MstBikotitle ?? ""):
                    {
                        obj_rep = new Repository.M_memo_Repository.SubConv();                                 // 備考タイトルマスタ      
                        break;
                    }
                case var case11 when case11 == (model_cvitem.MstBikolst ?? ""):
                    {
                        obj_rep = new Repository.M_memo_lst_Repository.SubConv();                               // 備考入力補助リストマスタ
                        break;
                    }
                case var case12 when case12 == (model_cvitem.MstGazotitle ?? ""):
                    {
                        obj_rep = new Repository.M_gazo_title_Repository.SubConv();                           // 画像タイトルマスタ
                        break;
                    }
                // 仮マスタ読込先を"仮テーブルから中間ファイル"へ変更するためコメントアウト
                // Case model_cvitem.MstNkinkomok : obj_rep = New Njc.Repository.M_nkin_Repository.SubConv                                '入金項目マスタ(仮テーブル読込)
                // Case model_cvitem.MstBkbrui : obj_rep = New Njc.Repository.M_bk_rui_Repository.SubConv                                 '物件分類マスタ(仮テーブル読込)
                // Case model_cvitem.MstNkinkbn : obj_rep = New Njc.Repository.M_nkbn_Repository.SubConv                                  '入金区分マスタ(仮テーブル読込)
                case var case13 when case13 == (model_cvitem.MstKozasyubetu ?? ""):    // 口座種別マスタ
                    {
                        break;
                    }
                case var case14 when case14 == (model_cvitem.MstKozo ?? ""):           // 物件構造マスタ
                    {
                        break;
                    }
                case var case15 when case15 == (model_cvitem.MstTorihikitaiyo ?? ""):  // 取引態様マスタ
                    {
                        break;
                    }
                case var case16 when case16 == (model_cvitem.MstHyrui ?? ""):          // 部屋分類マスタ
                    {
                        break;
                    }
                case var case17 when case17 == (model_cvitem.MstSetubi ?? ""):         // 設備マスタ
                    {
                        break;
                    }

                // 業者情報
                case var case18 when case18 == (model_cvitem.GyCyukaiBase ?? ""):
                    {
                        obj_rep = new Repository.Gydata_fudo_Repository.SubConv();                            // 仲介・管理業者マスタ
                        break;
                    }
                case var case19 when case19 == (model_cvitem.GyCyukaiKoza ?? ""):
                    {
                        obj_rep = new Repository.Gydata_fudokoza_Repository.SubConv();                        // 仲介・管理業者マスタ口座
                        break;
                    }
                case var case20 when case20 == (model_cvitem.GyCyukaiMemo ?? ""):
                    {
                        obj_rep = new Repository.Gydata_fudomemo_Repository.SubConv();                        // 仲介・管理業者マスタメモ
                        break;
                    }
                case var case21 when case21 == (model_cvitem.GyHokenBase ?? ""):
                    {
                        obj_rep = new Repository.Gydata_hoken_Repository.SubConv();                            // 保険業者マスタ
                        break;
                    }
                case var case22 when case22 == (model_cvitem.GyHokenKoza ?? ""):
                    {
                        obj_rep = new Repository.Gydata_hokenkoza_Repository.SubConv();                        // 保険業者マスタ口座
                        break;
                    }
                case var case23 when case23 == (model_cvitem.GyHokenMemo ?? ""):
                    {
                        obj_rep = new Repository.Gydata_hokenmemo_Repository.SubConv();                        // 保険業者マスタメモ
                        break;
                    }
                case var case24 when case24 == (model_cvitem.GyYatinhosyoBase ?? ""):
                    {
                        obj_rep = new Repository.Gydata_hosyo_Repository.SubConv();                       // 家賃保証業者マスタ
                        break;
                    }
                // Case model_cvitem.GyYatinhosyoKoza : obj_rep = New Njc.Repository.gydata_hosyokoza_Repository.SubConv                  '家賃保証業者マスタ口座
                case var case25 when case25 == (model_cvitem.GyYatinhosyoMemo ?? ""):
                    {
                        obj_rep = new Repository.gydata_hosyomemo_Repository.SubConv();                   // 家賃保証業者マスタメモ
                        break;
                    }
                case var case26 when case26 == (model_cvitem.GySyuzenBase ?? ""):
                    {
                        obj_rep = new Repository.Gydata_szen_Repository.SubConv();                            // 修繕業者マスタ
                        break;
                    }
                case var case27 when case27 == (model_cvitem.GySyuzenKoza ?? ""):
                    {
                        obj_rep = new Repository.Gydata_szenkoza_Repository.SubConv();                        // 修繕業者マスタ口座
                        break;
                    }
                case var case28 when case28 == (model_cvitem.GySyuzenMemo ?? ""):
                    {
                        obj_rep = new Repository.Gydata_szenmemo_Repository.SubConv();                        // 修繕業者マスタメモ
                        break;
                    }
                case var case29 when case29 == (model_cvitem.GyLifelineBase ?? ""):
                    {
                        obj_rep = new Repository.Gydata_lifeline_Repository.SubConv();                      // ライフラインマスタ
                        break;
                    }
                case var case30 when case30 == (model_cvitem.GySekoBase ?? ""):
                    {
                        obj_rep = new Repository.Gydata_seko_Repository.SubConv();                              // 施工会社マスタ
                        break;
                    }
                case var case31 when case31 == (model_cvitem.GySisetuBase ?? ""):
                    {
                        obj_rep = new Repository.Gydata_sisetu_Repository.SubConv();                          // 施設保守会社マスタ
                        break;
                    }

                // 自社情報
                case var case32 when case32 == (model_cvitem.JisyaBase ?? ""):
                    {
                        obj_rep = new Repository.Jisyadata_Repository.SubConv();                                 // 自社基本情報
                        break;
                    }
                // 20160913_2 自社口座移行処理の追加 -chg sta
                // Case model_cvitem.JisyaKoza : obj_rep = New Njc.Repository.Jisyadata_koza_Repository.SubConv                            '自社口座情報
                case var case33 when case33 == (model_cvitem.JisyaKoza ?? ""):
                    {
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    obj_rep = new Repository.Jisyadata_koza_BaseMid_Repository.SubConv();                                      // 自社口座情報(汎用用)
                                    break;
                                }
                        }

                        break;
                    }
                // 20160913_2 自社口座移行処理の追加 -chg end
                case var case34 when case34 == (model_cvitem.JisyaTanto ?? ""):
                    {
                        obj_rep = new Repository.Profile_logonuser_Repository.SubConv();                        // 自社担当者情報       
                        break;
                    }
                case var case35 when case35 == (model_cvitem.JisyaMemo ?? ""):
                    {
                        obj_rep = new Repository.Jisyadata_memo_Repository.SubConv();                            // 自社メモ情報
                        break;
                    }

                // 口座関連情報
                case var case36 when case36 == (model_cvitem.FBFuriirai ?? ""):
                    {
                        obj_rep = new Repository.M_fb_sgfirai_Repository.SubConv();                             // 振込依頼人情報
                        break;
                    }
                case var case37 when case37 == (model_cvitem.FBFuritesuryo ?? ""):
                    {
                        obj_rep = new Repository.M_fb_sgfiraitesu_Repository.SubConv();                      // 振込手数料情報
                        break;
                    }
                case var case38 when case38 == (model_cvitem.FBKozafurikae ?? ""):
                    {
                        obj_rep = new Repository.M_fb_fkaejyoho_Repository.SubConv();                        // 口座振替情報
                        break;
                    }
                case var case39 when case39 == (model_cvitem.FBNsSyutoku ?? ""):
                    {
                        obj_rep = new Repository.M_fb_nskinsetting_Repository.SubConv();                       // 入出金取得情報
                        break;
                    }
                case var case40 when case40 == (model_cvitem.MstYatinKoza ?? ""):
                    {
                        obj_rep = new Repository.M_yatinkoza_Repository.SubConv();                            // 家賃入金口座情報
                        break;
                    }
                case var case41 when case41 == (model_cvitem.MstANSERArea ?? ""):
                    {
                        obj_rep = new Repository.M_spcarea_Repository.SubConv();                              // ANSERエリア情報                 
                        break;
                    }
                case var case42 when case42 == (model_cvitem.MstANSERAccpoint ?? ""):
                    {
                        obj_rep = new Repository.M_spcaccesspoint_Repository.SubConv();                   // ANSERアクセスポイント情報       
                        break;
                    }
                case var case43 when case43 == (model_cvitem.FBANSERSetuzoku ?? ""):
                    {
                        obj_rep = new Repository.M_spcsetuzoku_Repository.SubConv();                       // ANSER接続情報                   
                        break;
                    }

                // 家主情報
                case var case44 when case44 == (model_cvitem.OwBase ?? ""):
                    {
                        obj_rep = new Repository.Owdata_Repository.SubConv();                                       // 家主情報(基本)
                        break;
                    }
                case var case45 when case45 == (model_cvitem.OwKoza ?? ""):
                    {
                        obj_rep = new Repository.Owdata_koza_Repository.SubConv();                                  // 家主情報(口座)
                        break;
                    }
                case var case46 when case46 == (model_cvitem.OwEvent ?? ""):
                    {
                        obj_rep = new Repository.Owdata_event_Repository.SubConv();                                // 家主情報(イベント)
                        break;
                    }
                case var case47 when case47 == (model_cvitem.OwMemo ?? ""):
                    {
                        obj_rep = new Repository.Owdata_memo_Repository.SubConv();                                  // 家主情報(メモ)
                        break;
                    }

                // 契約者情報
                case var case48 when case48 == (model_cvitem.KysBase ?? ""):
                    {
                        obj_rep = new Repository.Kysdata_Repository.SubConv();                                     // 契約者情報(基本)
                        break;
                    }
                case var case49 when case49 == (model_cvitem.KysKoza ?? ""):
                    {
                        obj_rep = new Repository.Kysdata_koza_Repository.SubConv();                                // 契約者情報(口座)
                        break;
                    }
                case var case50 when case50 == (model_cvitem.KysMemo ?? ""):
                    {
                        obj_rep = new Repository.Kysdata_memo_Repository.SubConv();                                // 契約者情報(メモ)
                        break;
                    }
                case var case51 when case51 == (model_cvitem.KysSyogoKana ?? ""):
                    {
                        obj_rep = new Repository.M_fb_fkomsyogo_kys_Repository.SubConv();                     // 契約者情報(照合用カナ)
                        break;
                    }
                case var case52 when case52 == (model_cvitem.KysHosyonin ?? ""):
                    {
                        obj_rep = new Repository.Kysdata_hosyo_Repository.SubConv();                           // 契約者情報(保証人)
                        break;
                    }

                // 物件情報
                case var case53 when case53 == (model_cvitem.BkBase ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_Repository.SubConv();                                       // 物件基本情報
                        break;
                    }
                case var case54 when case54 == (model_cvitem.BkSyosai ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_detail_Repository.SubConv();                              // 物件詳細情報
                        break;
                    }
                case var case55 when case55 == (model_cvitem.Bksyo ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_syo_Repository.SubConv();                                    // 物件所有者情報
                        break;
                    }
                case var case56 when case56 == (model_cvitem.BkGomi ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_dust_Repository.SubConv();                                  // 物件ゴミ情報
                        break;
                    }
                case var case57 when case57 == (model_cvitem.BkKenri ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_kenri_Repository.SubConv();                                // 物件権利情報
                        break;
                    }
                case var case58 when case58 == (model_cvitem.BkKotu ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_kotu_Repository.SubConv();                                  // 物件交通情報
                        break;
                    }
                case var case59 when case59 == (model_cvitem.BkSetudo ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_setudo_Repository.SubConv();                              // 物件接道情報
                        break;
                    }
                case var case60 when case60 == (model_cvitem.BkSyuhen ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_syuhen_Repository.SubConv();                              // 物件周辺情報
                        break;
                    }
                case var case61 when case61 == (model_cvitem.BkSzeniji ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_szeniji_Repository.SubConv();                            // 物件修繕維持管理連絡先情報
                        break;
                    }
                case var case62 when case62 == (model_cvitem.BkMemo ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_memo_Repository.SubConv();                                  // 物件メモ情報
                        break;
                    }
                case var case63 when case63 == (model_cvitem.BkKagi ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_kagi_Repository.SubConv();                                  // 物件鍵情報(汎用ツールのみ)
                        break;
                    }
                case var case64 when case64 == (model_cvitem.BkKinrincyusyajo ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_parkingother_Repository.SubConv();                // 物件近隣駐車場情報(汎用ツールのみ)
                        break;
                    }
                case var case65 when case65 == (model_cvitem.BkSansyofile ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_relfile_Repository.SubConv();                         // 物件参照ファイル情報(汎用ツールのみ)
                        break;
                    }
                case var case66 when case66 == (model_cvitem.BkHendo ?? ""):
                    {
                        obj_rep = new Repository.Bkdata_hendo_Repository.SubConv();                                // 物件変動費親メーター情報(汎用ツールのみ)
                        break;
                    }

                // 部屋情報
                case var case67 when case67 == (model_cvitem.HyBase ?? ""):
                    {
                        obj_rep = new Repository.Hydata_Repository.SubConv();                                       // 部屋基本情報
                        break;
                    }
                case var case68 when case68 == (model_cvitem.HySyosai ?? ""):
                    {
                        obj_rep = new Repository.Hydata_detail_Repository.SubConv();                              // 部屋詳細情報
                        break;
                    }
                case var case69 when case69 == (model_cvitem.Hysyo ?? ""):
                    {
                        obj_rep = new Repository.Hydata_syo_Repository.SubConv();                                    // 部屋所有者情報
                        break;
                    }
                case var case70 when case70 == (model_cvitem.HyParking ?? ""):
                    {
                        obj_rep = new Repository.Hydata_parking_Repository.SubConv();                            // 部屋駐車場情報
                        break;
                    }
                case var case71 when case71 == (model_cvitem.HyTokuyaku ?? ""):
                    {
                        obj_rep = new Repository.Hydata_tokuyaku_Repository.SubConv();                          // 部屋特約情報
                        break;
                    }
                case var case72 when case72 == (model_cvitem.HyKagi ?? ""):
                    {
                        obj_rep = new Repository.Hydata_kagi_Repository.SubConv();                                  // 部屋鍵情報
                        break;
                    }
                case var case73 when case73 == (model_cvitem.HyMenseki ?? ""):
                    {
                        obj_rep = new Repository.Hydata_othermenseki_Repository.SubConv();                       // 部屋面積情報
                        break;
                    }
                case var case74 when case74 == (model_cvitem.HyMadoriutiwake ?? ""):
                    {
                        obj_rep = new Repository.Hydata_madoriutiwake_Repository.SubConv();                // 部屋間取内訳情報
                        break;
                    }
                case var case75 when case75 == (model_cvitem.HySzeniji ?? ""):
                    {
                        obj_rep = new Repository.Hydata_szeniji_Repository.SubConv();                            // 部屋修繕維持管理連絡先情報
                        break;
                    }
                // 20160913_2 部屋設備移行処理の追加 -chg sta
                // Case model_cvitem.HySetubi : obj_rep = New Njc.Repository.Hydata_setubilst_Repository.SubConv                           '部屋設備情報
                case var case76 when case76 == (model_cvitem.HySetubi ?? ""):
                    {
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    obj_rep = new Repository.Hydata_setubilst_BaseMid_Repository.SubConv();                                    // 部屋設備情報(汎用用)
                                    break;
                                }
                        }

                        break;
                    }
                // 20160913_2 部屋設備移行処理の追加 -chg end
                case var case77 when case77 == (model_cvitem.HyNkinkomk ?? ""):
                    {
                        obj_rep = new Repository.Hydata_nkin_Repository.SubConv();                              // 部屋入金項目情報
                        break;
                    }
                case var case78 when case78 == (model_cvitem.HyHendo ?? ""):
                    {
                        obj_rep = new Repository.Hydata_hendo_Repository.SubConv();                                // 部屋変動費各戸メーター情報
                        break;
                    }
                case var case79 when case79 == (model_cvitem.HyMemo ?? ""):
                    {
                        obj_rep = new Repository.Hydata_memo_Repository.SubConv();                                  // 部屋メモ情報
                        break;
                    }
                case var case80 when case80 == (model_cvitem.HyCommonsalespoint ?? ""):
                    {
                        obj_rep = new Repository.Hydata_commonsalespointparts_Repository.SubConv();     // 部屋共通セールスポイント情報(汎用ツールのみ)
                        break;
                    }
                case var case81 when case81 == (model_cvitem.HyConfirm ?? ""):
                    {
                        obj_rep = new Repository.Hydata_confirm_Repository.SubConv();                            // 部屋契約解約確認事項情報(汎用ツールのみ)
                        break;
                    }
                case var case82 when case82 == (model_cvitem.HyKenri ?? ""):
                    {
                        obj_rep = new Repository.Hydata_kenri_Repository.SubConv();                                // 部屋権利情報(汎用ツールのみ)
                        break;
                    }
                case var case83 when case83 == (model_cvitem.HySansyofile ?? ""):
                    {
                        obj_rep = new Repository.Hydata_relfile_Repository.SubConv();                         // 部屋参照ファイル情報(汎用ツールのみ)
                        break;
                    }
                case var case84 when case84 == (model_cvitem.HyGenjotanka ?? ""):
                    {
                        obj_rep = new Repository.Hydata_szen_Repository.SubConv();                            // 部屋原状回復目安単価情報(汎用ツールのみ)
                        break;
                    }

                // 送金ルール情報
                case var case85 when case85 == (model_cvitem.SoruleBase ?? ""):
                    {
                        obj_rep = new Repository.Sorule_Repository.SubConv();                                   // 送金ルール基本情報
                        break;
                    }
                case var case86 when case86 == (model_cvitem.SoruleSosaki ?? ""):
                    {
                        obj_rep = new Repository.Sorule_sosaki_Repository.SubConv();                          // 送金ルール送金先情報
                        break;
                    }
                case var case87 when case87 == (model_cvitem.SoruleNkin ?? ""):
                    {
                        obj_rep = new Repository.Sorule_nk_cmrule_Repository.SubConv();                         // 送金ルール入金項目情報
                        break;
                    }
                case var case88 when case88 == (model_cvitem.SoruleKojo ?? ""):
                    {
                        obj_rep = new Repository.Sorule_kojo_cmrule_Repository.SubConv();                       // 送金ルール控除項目情報
                        break;
                    }

                // 契約情報
                case var case89 when case89 == (model_cvitem.KyBase ?? ""):
                    {
                        obj_rep = new Repository.Kydata_Repository.SubConv();                                       // 契約基本情報
                        break;
                    }
                case var case90 when case90 == (model_cvitem.KyRireki ?? ""):
                    {
                        obj_rep = new Repository.Kydata_kihon_Repository.SubConv();                               // 契約履歴情報
                        break;
                    }
                case var case91 when case91 == (model_cvitem.KyKys ?? ""):
                    {
                        obj_rep = new Repository.Kydata_kys_Repository.SubConv();                                    // 契約契約者情報
                        break;
                    }
                case var case92 when case92 == (model_cvitem.KyNyukyo ?? ""):
                    {
                        obj_rep = new Repository.Kydata_nyukyo_Repository.SubConv();                              // 契約入居者情報
                        break;
                    }
                case var case93 when case93 == (model_cvitem.KyHosyonin ?? ""):
                    {
                        obj_rep = new Repository.Kydata_hosyonin_Repository.SubConv();                          // 契約保証人情報
                        break;
                    }
                case var case94 when case94 == (model_cvitem.KyCar ?? ""):
                    {
                        obj_rep = new Repository.Kydata_car_Repository.SubConv();                                    // 契約車情報
                        break;
                    }
                case var case95 when case95 == (model_cvitem.KyHoken ?? ""):
                    {
                        obj_rep = new Repository.Kydata_hoken_Repository.SubConv();                                // 契約保険情報
                        break;
                    }
                case var case96 when case96 == (model_cvitem.KyTokuyaku ?? ""):
                    {
                        obj_rep = new Repository.Kydata_tokuyaku_Repository.SubConv();                          // 契約特約事項情報
                        break;
                    }
                case var case97 when case97 == (model_cvitem.KyMemo ?? ""):
                    {
                        obj_rep = new Repository.Kydata_memo_Repository.SubConv();                                  // 契約メモ情報
                        break;
                    }
                case var case98 when case98 == (model_cvitem.KyNkinkomk ?? ""):
                    {
                        obj_rep = new Repository.Kydata_nkin_Repository.SubConv();                              // 契約入金項目情報
                        break;
                    }
                case var case99 when case99 == (model_cvitem.KyNkinkomkNx ?? ""):
                    {
                        obj_rep = new Repository.Kydata_nkin_nx_Repository.SubConv();                         // 契約次回入金項目情報
                        break;
                    }
                case var case100 when case100 == (model_cvitem.KyHendo ?? ""):
                    {
                        obj_rep = new Repository.Kydata_hendo_Repository.SubConv();                                // 契約変動費各戸メーター情報
                        break;
                    }
                case var case101 when case101 == (model_cvitem.KyKojoRule ?? ""):
                    {
                        obj_rep = new Repository.Kydata_kojorule_Repository.SubConv();                          // 契約控除ルール情報
                        break;
                    }
                case var case102 when case102 == (model_cvitem.KySorule ?? ""):
                    {
                        obj_rep = new Repository.Kydata_sorule_Repository.SubConv();                              // 契約送金ルール情報
                        break;
                    }
                case var case103 when case103 == (model_cvitem.KyKai ?? ""):
                    {
                        obj_rep = new Repository.Kydata_kai_Repository.SubConv();                                    // 契約解約情報
                        break;
                    }
                // 20160531 鍵情報移行処理の修正 -del sta
                // '2016.04.06 契約鍵情報の移行処理追加 -add
                // Case model_cvitem.KyKagi : obj_rep = New Njc.Repository.Kydata_kagi_Repository.SubConv                                  '契約鍵情報
                // 20160531 鍵情報移行処理の修正 -del end
                case var case104 when case104 == (model_cvitem.KySzen ?? ""):
                    {
                        obj_rep = new Repository.Kydata_kaiszen_Repository.SubConv();                               // 契約修繕見積情報       
                        break;
                    }
                case var case105 when case105 == (model_cvitem.KySzenmeisai ?? ""):
                    {
                        obj_rep = new Repository.Kydata_kaiszenmeisai_Repository.SubConv();                   // 契約修繕見積詳細情報   
                        break;
                    }

                // 請求情報
                case var case106 when case106 == (model_cvitem.SqKajyo ?? ""):
                    {
                        obj_rep = new Repository.Azukanri_Repository.SubConv();                                    // 過剰金情報
                        break;
                    }
                case var case107 when case107 == (model_cvitem.SqUnyotaino ?? ""):
                    {
                        obj_rep = new Repository.Unyotainodata_Repository.SubConv();                           // 運用開始時未収滞納金情報
                        break;
                    }
                case var case108 when case108 == (model_cvitem.SqSq ?? ""):
                    {
                        obj_rep = new Repository.Sqdata_Repository.SubConv();                                         // 請求情報
                        break;
                    }
                case var case109 when case109 == (model_cvitem.SqHendokensin ?? ""):
                    {
                        obj_rep = new Repository.Hendodata_meisai_Repository.SubConv();                      // 変動費検針情報
                        break;
                    }
                case var case110 when case110 == (model_cvitem.SqKoteiKojo ?? ""):
                    {
                        obj_rep = new Repository.Koteirule_Repository.SubConv();                               // 家主固定控除情報
                        break;
                    }
                case var case111 when case111 == (model_cvitem.SqSqKojo ?? ""):
                    {
                        obj_rep = new Repository.Kjdata_Repository.SubConv();                                     // 家主請求控除情報
                        break;
                    }

                // クレーム情報
                case var case112 when case112 == (model_cvitem.ClaimBase ?? ""):
                    {
                        obj_rep = new Repository.Claimdata_Repository.SubConv();                                 // クレーム基本情報
                        break;
                    }
                case var case113 when case113 == (model_cvitem.ClaimTaiorireki ?? ""):
                    {
                        obj_rep = new Repository.Claim_taio_Repository.SubConv();                          // クレーム対応履歴情報
                        break;
                    }
                case var case114 when case114 == (model_cvitem.ClaimRelfile ?? ""):
                    {
                        obj_rep = new Repository.Claimdata_relfile_Repository.SubConv();                      // クレーム関連ファイル情報
                        break;
                    }

                // 修繕関連
                case var case115 when case115 == (model_cvitem.SzenBase ?? ""):
                    {
                        obj_rep = new Repository.Szendata_Repository.SubConv();                                   // 修繕基本情報
                        break;
                    }
                case var case116 when case116 == (model_cvitem.SzenSzen ?? ""):
                    {
                        obj_rep = new Repository.Szendata_szen_Repository.SubConv();                              // 修繕見積情報
                        break;
                    }
                case var case117 when case117 == (model_cvitem.SzenSzenmeisai ?? ""):
                    {
                        obj_rep = new Repository.Szendata_szenmeisai_Repository.SubConv();                  // 修繕見積詳細情報
                        break;
                    }
                case var case118 when case118 == (model_cvitem.SzenClaim ?? ""):
                    {
                        obj_rep = new Repository.Szendata_claim_Repository.SubConv();                            // 修繕クレーム関連付け情報
                        break;
                    }
                case var case119 when case119 == (model_cvitem.SzenRelfile ?? ""):
                    {
                        obj_rep = new Repository.Szendata_relfile_Repository.SubConv();                            // 修繕関連ファイル情報       '20160627 修繕関連ファイル移行修正
                        break;
                    }
                case var case120 when case120 == (model_cvitem.SzenMemo ?? ""):
                    {
                        obj_rep = new Repository.Szendata_memo_Repository.SubConv();                              // 修繕メモ情報
                        break;
                    }

                // 初期設定情報
                case var case121 when case121 == (model_cvitem.SyskanriBase ?? ""):
                    {
                        obj_rep = new Repository.Profile_fk_base_Repository.SubConv();                        // 初期設定情報
                        break;
                    }
                case var case122 when case122 == (model_cvitem.SyskanriZei ?? ""):
                    {
                        obj_rep = new Repository.Profile_fk_zei_Repository.SubConv();                          // 税編集情報
                        break;
                    }
                case var case123 when case123 == (model_cvitem.SyskanriHenkanmoji ?? ""):
                    {
                        obj_rep = new Repository.Profile_fk_henkanmoji_Repository.SubConv();            // 変換文字情報
                        break;
                    }
                case var case124 when case124 == (model_cvitem.SyskanriNkinkomkmerge ?? ""):
                    {
                        obj_rep = new Repository.Profile_fk_nkinkomkmerge_Repository.SubConv();      // 入金項目集約情報
                        break;
                    }

                // 物件データ連動情報
                case var case125 when case125 == (model_cvitem.RendoSosinBase ?? ""):
                    {
                        obj_rep = new Repository.M_sendsetting_Repository.SubConv();                        // 送信設定基本情報
                        break;
                    }
                case var case126 when case126 == (model_cvitem.RendoSosinJisyaweb ?? ""):
                    {
                        obj_rep = new Repository.M_site_sendsetting_jisyaweb_Repository.SubConv();      // 送信設定自社web情報
                        break;
                    }
                case var case127 when case127 == (model_cvitem.RendoSosinHomes ?? ""):
                    {
                        obj_rep = new Repository.M_site_sendsetting_homes_Repository.SubConv();            // 送信設定HOMES情報
                        break;
                    }
                case var case128 when case128 == (model_cvitem.RendoSosinAthome ?? ""):
                    {
                        obj_rep = new Repository.M_site_sendsetting_athome_Repository.SubConv();          // 送信設定athome情報
                        break;
                    }
                case var case129 when case129 == (model_cvitem.RendoSosinSuumo ?? ""):
                    {
                        obj_rep = new Repository.M_site_sendsetting_suumo_Repository.SubConv();            // 送信設定SUUMO情報
                        break;
                    }

                // ポータル連動情報
                case var case130 when case130 == (model_cvitem.RendoKokokuJisyaweb ?? ""):
                    {
                        obj_rep = new Repository.Hydata_kokoku_jisyaweb_Repository.SubConv();          // 広告補足自社web情報
                        break;
                    }
                case var case131 when case131 == (model_cvitem.RendoKokokuHomes ?? ""):
                    {
                        obj_rep = new Repository.Hydata_kokoku_homes_Repository.SubConv();                // 広告補足HOMES情報
                        break;
                    }
                case var case132 when case132 == (model_cvitem.RendoKokokuAthome ?? ""):
                    {
                        obj_rep = new Repository.Hydata_kokoku_athome_Repository.SubConv();              // 広告補足athome情報
                        break;
                    }
                case var case133 when case133 == (model_cvitem.RendoKokokuSuumo ?? ""):
                    {
                        obj_rep = new Repository.Hydata_kokoku_suumo_Repository.SubConv();                // 広告補足SUUMO情報
                        break;
                    }
                case var case134 when case134 == (model_cvitem.RendoHyrui ?? ""):
                    {
                        obj_rep = new Repository.M_hy_ruisite_Repository.SubConv();                             // ポータル連動部屋分類情報
                        break;
                    }
                case var case135 when case135 == (model_cvitem.RendoHysosin ?? ""):
                    {
                        obj_rep = new Repository.Hydata_sosin_Repository.SubConv();                           // 部屋毎送信情報
                        break;
                    }
                case var case136 when case136 == (model_cvitem.RendoBtoBgroup ?? ""):
                    {
                        obj_rep = new Repository.Hydata_btobgroup_Repository.SubConv();                     // BtoBグループ設定情報
                        break;
                    }
                case var case137 when case137 == (model_cvitem.RendoMapdisp ?? ""):
                    {
                        obj_rep = new Repository.Hydata_mapdisp_Repository.SubConv();                         // 地図表示詳細設定情報
                        break;
                    }

            }

            return obj_rep;

        }

        #endregion

        #region 中間ファイルチェック処理

        /// <summary>
        /// 中間ファイルチェック実行前準備
        /// </summary>
        /// <param name="list_cv"></param>
        /// <returns></returns>
        /// <remarks>
        /// ①コンバート用のDB情報を格納する仮テーブル作成
        /// ②DBの基本情報と①を紐付けた値をハッシュテーブルへ格納
        /// ※一度処理を行えば良いのでフラグを持たせて処理の実行有無を分岐させる
        /// </remarks>
        private bool Before_DBWrite(List<string> list_cv)
        {

            int rowcnt = 0;
            bool rtn = true;


            // ----------------------------------------------------------------------
            // 革命10DBの情報に対してコンバート用のDB情報を付加する
            // ----------------------------------------------------------------------
            // ①付加する情報の仮テーブルを作成(自作仕様)
            // 初期化(DROP)
            string cvdbinfo_dropqry = DBQuery.Qry_DropInfo(CVDBInfoModule.CVDBINFO_DBNAME, true);
            DBExec.Exec_NonQuery(sqlcnnv10, cvdbinfo_dropqry, ref rowcnt);

            // 仮テーブル作成(CREATE)
            string cvdbinfo_createqry = CVDBInfoModule.Get_CVDBInfo_CreateQry();
            DBExec.Exec_NonQuery(sqlcnnv10, cvdbinfo_createqry, ref rowcnt);

            // ②仮テーブルへコンバート用のDB情報を挿入
            // 2016.02.22 エラー時の対処4 -add
            bool normalflg = true;
            foreach (var item in list_cv)
            {
                string cvitem = item;
                string cvdbinfo_insertqry = CVDBInfoModule.Get_CVDBInfo_InsertQry(item);
                if (!string.IsNullOrEmpty(cvdbinfo_insertqry))
                {
                    normalflg = DBExec.Exec_NonQuery(sqlcnnv10, cvdbinfo_insertqry, ref rowcnt);
                }
                if (normalflg == false)
                {
                    rtn = false;
                    return rtn;
                }
            }


            // ----------------------------------------------------------------------
            // 革命10DBフィールド情報(型/サイズ)を取得してハッシュテーブルへ格納
            // ----------------------------------------------------------------------
            string tmp_sql_getfldinfo = CVDBInfoModule.Qry_GetTableInfo();
            normalflg = GetFieldInfo.Get_HashFieldInfo(sqlcnnv10, tmp_sql_getfldinfo);
            if (normalflg == false)
            {
                rtn = false;
                return rtn;
            }

            return rtn;

        }

        /// <summary>
        /// 中間ファイルチェック処理 (DBを元に型・サイズチェックを行う)
        /// </summary>
        /// <param name="list_chkitem"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Chk_MidFileChk_Main(List<string> list_chkitem)
        {

            bool rtn = true;
            // 20161007 実行ボタン押下→DB書込開始までの進捗表示処理追加 -add sta
            var obj_pgb = new ProgressBarManager();


            // --------------------------------------------------
            // ログ初期設定
            // --------------------------------------------------
            obj_pgb.pgbInitPart(pgbtotalcnt);                               // ----- 個別用プログレスバー更新 -----
            lblTotalSituation.Text = CommonModule.SITUATION_MID_CHKBEFORE;
            Set_LogInit(true);                                           // ログファイル出力


            // --------------------------------------------------
            // 中間ファイルチェック処理
            // --------------------------------------------------
            lblTotalSituation.Text = CommonModule.SITUATION_MID_CHK;
            bool normalflg = Get_MidFile_Data(list_chkitem);


            // --------------------------------------------------
            // 終了処理
            // --------------------------------------------------
            obj_pgb.pgbInitPart(pgbtotalcnt);                               // ----- 個別用プログレスバー更新 -----
            lblTotalSituation.Text = CommonModule.SITUATION_MID_CHKAFTER;

            if (EtcMethod.Chk_FileExist(CommonModule.MiddleLogFilePath) == false)          // ログファイル出力
            {
                // ログファイル(CSV)のデータ部出力
                string tmp_sql = LogSetting.Get_LogTblSelectQry(true);
                bool logoutflg = true;
                string errstr = "";
                logoutflg = FileMethod.TblView_Output_CSV(sqlcnnv10, CommonModule.MiddleLogFilePath, "", "", ref errstr, tmp_sql, true);
                // ヘッダーを加えて加工
                if (logoutflg)
                {
                    // 20161104 ログ出力内容修正(開発用以外はテーブル名を表示しない) -chg sta
                    // Call Me.Set_LogHeader(MiddleLogFilePath, LOG_HEADER_TOTAL)
                    string tmp_header = CommonModule.LOG_HEADER_TOTAL;
                    if (CommonModule.Dev_CVFlg == false)
                    {
                        tmp_header = tmp_header.Replace(",対象TBL名", "");
                    }
                    Set_LogHeader(CommonModule.MiddleLogFilePath, tmp_header);
                }
                // 20161104 ログ出力内容修正(開発用以外はテーブル名を表示しない) -chg end
                else
                {
                    CommonModule.MsgResult = MessageBox.Show(errstr, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }

            // 返却
            rtn = normalflg;
            return rtn;

        }

        /// <summary>
        /// 中間ファイルチェック実行処理→ログ出力 '20161006 中間ファイルチェック処理の速度改善対応 -add
        /// </summary>
        /// <param name="midchklist"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool Get_MidFile_Data(List<string> midchklist)
        {
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
            // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用   
            var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
            var obj_com = new CommonRepository();                  // 共通処理用
            int tmp_logcnt = 0;                                   // ログ出力時のソート用
            bool normalflg = true;                                 // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値

            // ----- 全体用・中間用プログレスバー更新 -----
            int pgbcnt = 0;
            int pgbtotalcnt_total = midchklist.Count;
            obj_pgb.pgbInitTotal(pgbtotalcnt_total);
            obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, true);
            obj_pgb.pgbInitChkTotal(pgbtotalcnt_total);
            obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, true);


            // ログ出力
            CommonModule.Log_OutputCnt = int.Parse(txtLogOutputCnt.Text);            // ログ出力件数
            string tmp_sql_sta = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, CommonModule.LOG_SYORIKOMK_MIDCHKSTA), true);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_sta, ref tmp_logcnt);

            // 中間ファイルチェック処理
            var tmptmpcnt = default(long);
            foreach (var filesheetname in midchklist)
            {

                string[] tmp_str = filesheetname.Split('-');
                string filename = tmp_str[0];
                string sheetname = tmp_str[1];
                var headervalue = new object();
                var fldvaluegrp = new object();
                var list_chkduplicate = new List<string>();
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -add
                string tmptblname = CommonModule.PRE_TBL_NAME + sheetname;

                // ----- デバッグ用処理 ----- sta
                // If filesheetname <> "自社情報-自社口座情報" Then
                // GoTo chkskiplbl
                // End If
                // ----- デバッグ用処理 ----- end

                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                // '汎用用中間ファイルを直接チェックする処理
                // Select Case filesheetname
                // Case "自社情報-自社口座情報", "部屋情報-部屋設備情報"
                // filename = "中間ファイル"
                // rtn = Me.Get_BaseMidFile_Data(filename, sheetname)
                // If rtn = False Then
                // Return rtn
                // Else
                // GoTo chkskiplbl
                // End If
                // Case "物件情報-物件変動費親メーター情報", "物件情報-物件近隣駐車場情報", "物件情報-(未使用の契約者データは移行しない)", "物件情報-(未使用の家主データは移行しない)", _
                // "部屋情報-部屋面積情報", "部屋情報-部屋権利情報", _
                // "契約情報-契約送金ルール情報", "契約情報-契約控除ルール情報", "契約情報-契約解約情報", "契約情報-契約修繕見積情報", _
                // "契約情報-契約修繕見積詳細情報", "契約情報-契約変動費親メーター情報", "契約情報-契約解約確認事項情報", "契約情報-契約同時契約情報", _
                // "契約情報-契約原状回復目安単価情報", "契約情報-契約敷金保証金随時処理情報", "契約情報-契約関連ファイル情報", "契約情報-契約空室待ち情報"
                // GoTo chkskiplbl
                // End Select

                // 'Excelファイル初期設定                  
                // Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
                // Dim readtbl As New DataTable()
                // Dim con_read As New OleDbConnection()
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

                // ----- ログ出力 -----                                       '20161009 ログ出力処理追加 -add
                Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDCHECK + "：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + filesheetname + ")");
                string tmptmpstr;
                tmptmpstr = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(905, CommonModule.LOG_SYORIKOMK_MIDFILE, filesheetname), false);
                int argrowcnt = (int)tmptmpcnt;
                DBExec.Exec_NonQuery(sqlcnnv10, tmptmpstr, ref argrowcnt);
                tmptmpcnt = argrowcnt;

                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                // 'オープン処理
                // rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

                // 'オープン処理失敗時は処理を抜ける
                // If rtn = False Then
                // Call excelfile.ExcelFile_ReadClose(con_read)
                // Return rtn
                // End If

                // データ取得
                string tmp_sql = " SELECT * FROM " + tmptblname;
                var readtbl = new DataTable();
                int rowcnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, ref rtn);

                // オープン処理失敗時は処理を抜ける
                if (rtn == false)
                {
                    return rtn;
                }
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end

                // ----- 個別用プログレスバー初期化 -----
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                // Dim rowcnt As Integer = readtbl.Rows.Count                  '行数取得   
                lblCVItem.Text = filename + " ： " + sheetname;           // チェック項目表示
                int pgbtotalcnt_part = 0;                         // プログレスバー総件数初期化
                int pgbbasecnt = 100;                             // 実件数で表示するか否かの基準値
                if (rowcnt <= pgbbasecnt)
                {
                    pgbtotalcnt_part = rowcnt;
                }
                else
                {
                    pgbtotalcnt_part = (int)Math.Round(Math.Ceiling(rowcnt / (double)pgbbasecnt));
                }
                obj_pgb.pgbInitPart(pgbtotalcnt_part);

                // 照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                // Call Me.Set_MidDataInfoToObj(filename, sheetname, False)
                switch (filesheetname ?? "")
                {
                    case "自社情報-自社口座情報":
                    case "部屋情報-部屋設備情報":
                        {
                            Set_MidDataInfoToObj(null, sheetname, true);
                            break;
                        }

                    default:
                        {
                            Set_MidDataInfoToObj(filename, sheetname, false);
                            break;
                        }
                }
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                // 'ヘッダーチェック
                // Dim list_errheader As New List(Of String)
                // normalflg = Me.Chk_MidFileHeader(sheetname, readtbl, list_errheader)

                // 'ヘッダーエラー時の終了処理
                // If normalflg = False Then

                // For Each errheader In list_errheader
                // 'ログ用に文字列を編集してクエリへ加工
                // Dim logvalue As String = ""
                // Dim logvalue_insertqry As String = ""
                // Dim tmp_cnt As Integer = 0

                // 'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
                // If errheader <> "" Then
                // logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
                // Else
                // logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
                // End If

                // 'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
                // logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
                // DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
                // Next

                // 'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
                // Call excelfile.ExcelFile_ReadClose(con_read)
                // rtn = normalflg
                // Return rtn

                // End If
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                // データチェック
                for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
                {

                    // 中断処理
                    Application.DoEvents();
                    if (CommonModule.MidChkCancelFlg | CommonModule.CancelFlg)
                    {
                        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                        // excelfile.ExcelFile_ReadClose(con_read)
                        string tmp_sql_cancel = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, CommonModule.LOG_SYORIKOMK_MIDCHKSTOP), true);
                        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_cancel, ref tmp_logcnt);
                        return rtn;
                    }

                    // 作業用変数
                    var hash_errdata = new SafeDictionary<string, string>();
                    var midkeytblfldname = new string[11];
                    var midkeyvalue = new string[11];
                    var tmp_hash = new SafeDictionary<string, string>();

                    // データ取得
                    for (int cntjj = 0, loopTo1 = readtbl.Columns.Count - 1; cntjj <= loopTo1; cntjj++)
                    {

                        // 項目名取得
                        string fldname = readtbl.Columns[cntjj].ColumnName.Trim();

                        // 登録値取得
                        string fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]).Trim();

                        // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                        tmp_hash.Add(fldname, fldvalue);

                    }
                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg sta
                    // 'ログ出力用のキーフィールドと値を取得
                    // For Each miditem In tmp_hash
                    // Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
                    // If Hash_ExistMiddatakey.Contains(get_key) Then
                    // Dim tmp_nostr As String = Hash_ExistMiddatakey.Item(get_key)
                    // Dim tmp_keyno As Integer = 0
                    // If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
                    // midkeytblfldname(tmp_keyno) = miditem.Key
                    // midkeyvalue(tmp_keyno) = miditem.Value
                    // End If
                    // End If
                    // Next

                    // Dim tmp_keyfldnamejp As String
                    // Dim tmp_keyvalue As String = ""
                    // Dim tmp_taisyodata As String = ""
                    // For cntkk = 1 To UBound(midkeytblfldname)
                    // If midkeytblfldname(cntkk) <> "" Then
                    // tmp_keyfldnamejp = midkeytblfldname(cntkk)
                    // tmp_keyvalue = midkeyvalue(cntkk)
                    // tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
                    // End If
                    // Next
                    // tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

                    // 'データチェック
                    // Dim tmp_flg As Boolean = True
                    // '20160927 ログの内容が不正になっているため修正 -add sta
                    // If filename = "送金ルール情報" Then
                    // tmp_hash("物件No") = tmp_hash("物件No") & "-" & tmp_hash("部屋No")
                    // End If
                    // '20160927 ログの内容が不正になっているため修正 -add end
                    // tmp_flg = Not (DataChk.Chk_DataPerItem_Mid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

                    string tmp_taisyodata = "";

                    switch (filesheetname ?? "")
                    {

                        case "自社情報-自社口座情報":
                        case "部屋情報-部屋設備情報":
                            {

                                // ログ出力用のキーフィールドと値を取得
                                foreach (var miditem in tmp_hash)
                                {
                                    string get_key = Conversions.ToString(Operators.ConcatenateObject(sheetname + CommonModule.STR_SPLIT_1, miditem.Key));
                                    if (CommonModule.Hash_BaseMiddatakey.ContainsKey(get_key))
                                    {
                                        string tmp_nostr = Conversions.ToString(CommonModule.Hash_BaseMiddatakey[get_key]);
                                        int tmp_keyno = 0;
                                        if (int.TryParse(tmp_nostr, out tmp_keyno))
                                        {
                                            midkeytblfldname[tmp_keyno] = Conversions.ToString(miditem.Key);
                                            midkeyvalue[tmp_keyno] = Conversions.ToString(miditem.Value);
                                        }
                                    }
                                }

                                string tmp_keyfldnamejp;
                                string tmp_keyvalue = "";
                                for (int cntkk = 1, loopTo2 = Information.UBound(midkeytblfldname); cntkk <= loopTo2; cntkk++)
                                {
                                    if (!string.IsNullOrEmpty(midkeytblfldname[cntkk]))
                                    {
                                        tmp_keyfldnamejp = midkeytblfldname[cntkk];
                                        tmp_keyvalue = midkeyvalue[cntkk];
                                        tmp_taisyodata = tmp_taisyodata + "、" + tmp_keyfldnamejp + " = " + tmp_keyvalue;
                                    }
                                }
                                tmp_taisyodata = tmp_taisyodata.Remove(0, 1);

                                // データチェック
                                bool tmp_flg = true;
                                bool localChk_DataPerItem_BaseMid() { string[] argmidkeytblfldname = midkeytblfldname; string[] argmidkeyvalue = midkeyvalue; var ret = DataChk.Chk_DataPerItem_BaseMid(sheetname, ref tmp_hash, list_chkduplicate, ref argmidkeytblfldname, ref argmidkeyvalue, ref hash_errdata); midkeytblfldname = (string[])argmidkeytblfldname; midkeyvalue = (string[])argmidkeyvalue; return ret; }

                                tmp_flg = !localChk_DataPerItem_BaseMid();
                                break;
                            }

                        default:
                            {

                                // ログ出力用のキーフィールドと値を取得
                                foreach (var miditem in tmp_hash)
                                {
                                    string get_key = Conversions.ToString(Operators.ConcatenateObject(sheetname + CommonModule.STR_SPLIT_1, miditem.Key));
                                    if (CommonModule.Hash_ExistMiddatakey.ContainsKey(get_key))
                                    {
                                        string tmp_nostr = Conversions.ToString(CommonModule.Hash_ExistMiddatakey[get_key]);
                                        int tmp_keyno = 0;
                                        if (int.TryParse(tmp_nostr, out tmp_keyno))
                                        {
                                            midkeytblfldname[tmp_keyno] = Conversions.ToString(miditem.Key);
                                            midkeyvalue[tmp_keyno] = Conversions.ToString(miditem.Value);
                                        }
                                    }
                                }

                                string tmp_keyfldnamejp;
                                string tmp_keyvalue = "";
                                for (int cntkk = 1, loopTo3 = Information.UBound(midkeytblfldname); cntkk <= loopTo3; cntkk++)
                                {
                                    if (!string.IsNullOrEmpty(midkeytblfldname[cntkk]))
                                    {
                                        tmp_keyfldnamejp = midkeytblfldname[cntkk];
                                        tmp_keyvalue = midkeyvalue[cntkk];
                                        tmp_taisyodata = tmp_taisyodata + "、" + tmp_keyfldnamejp + " = " + tmp_keyvalue;
                                    }
                                }
                                tmp_taisyodata = tmp_taisyodata.Remove(0, 1);

                                // 20161028 物件/部屋鍵取得方法修正 -add sta
                                // 鍵タイトルマスタの「鍵区分」はプログラム内部で使用する項目であるためログ出力用に成形する
                                if (sheetname == "鍵タイトルマスタ")
                                {
                                    tmp_taisyodata = tmp_taisyodata.Replace("鍵区分 = 1", "物件鍵タイトルマスタ");
                                    tmp_taisyodata = tmp_taisyodata.Replace("鍵区分 = 2", "部屋鍵タイトルマスタ");
                                }
                                // 20161028 物件/部屋鍵取得方法修正 -add end

                                // データチェック
                                bool tmp_flg = true;   // 実際に移行処理が実行された場合のレコードの移行可/不可  True…移行不可 False…移行可
                                if (filename == "送金ルール情報")
                                {
                                    tmp_hash["物件No"] = tmp_hash["物件No"] + "-" + tmp_hash["部屋No"];
                                }
                                bool localChk_DataPerItem_Mid() { string[] argmidkeytblfldname1 = midkeytblfldname; string[] argmidkeyvalue1 = midkeyvalue; var ret = DataChk.Chk_DataPerItem_Mid(sheetname, ref tmp_hash, list_chkduplicate, ref argmidkeytblfldname1, ref argmidkeyvalue1, ref hash_errdata); midkeytblfldname = (string[])argmidkeytblfldname1; midkeyvalue = (string[])argmidkeyvalue1; return ret; }

                                tmp_flg = !localChk_DataPerItem_Mid();

                                // 20161108_2 送金ルール送金先情報の移行制御処理修正 -add sta
                                // 送金先情報は管理形態を元に必須かどうかを判別する
                                // 自社物件、送金保留     … 必須ではない
                                // 自社物件、送金保留以外 … 必須
                                if (sheetname == "送金ルール送金先情報")
                                {
                                    if (tmp_flg == false)
                                    {
                                        Chk_SoruleSosaki(ref tmp_hash, ref hash_errdata);
                                    }
                                }
                                // 20161108_2 送金ルール送金先情報の移行制御処理修正 -add end

                                // 20161125 レビュー指摘事項対応 -chg sta
                                // '20161124 物件部屋所有者情報が両方存在する場合のチェック機能の追加 -add sta
                                // If sheetname = "物件所有者情報" Then
                                // If tmp_flg = False Then                                 '20161124 レビュー結果：所有者情報のチェックはあってもいいのでは？
                                // Call Me.Chk_BkHySyo(tmp_hash, hash_errdata)
                                // End If
                                // End If
                                // '20161124 物件部屋所有者情報が両方存在する場合のチェック機能の追加 -add end
                                if (sheetname == "物件所有者情報")
                                {
                                    Chk_BkHySyo(ref tmp_hash, ref hash_errdata);
                                }

                                break;
                            }
                            // 20161125 レビュー指摘事項対応 -chg end

                    }
                    // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -chg end

                    // ログ出力処理
                    var list_log = new List<string>();
                    foreach (var errvalue in hash_errdata)
                    {

                        // 20160927 ログの内容が不正になっているため修正 -chg sta
                        // Dim tmp_tblfldname As String = errvalue.Key
                        // Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
                        // Dim fldname As String = tmp_fldname(1)
                        string tmp_tblfldname = Conversions.ToString(errvalue.Key);
                        string[] tmp_sfname = Strings.Split(tmp_tblfldname, CommonModule.STR_SPLIT_2);
                        string fldname = "";

                        for (int cntjj = 0, loopTo4 = Information.UBound(tmp_sfname); cntjj <= loopTo4; cntjj++)
                        {
                            string[] tmp_sfsplietname = Strings.Split(tmp_sfname[cntjj], CommonModule.STR_SPLIT_1);
                            fldname = fldname + "、" + tmp_sfsplietname[1];
                        }
                        if (!string.IsNullOrEmpty(fldname))
                        {
                            fldname = fldname.Remove(0, 1);
                        }
                        // Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
                        // Dim fldname As String = tmp_fldname(1)
                        // 20160927 ログの内容が不正になっているため修正 -chg end
                        string[] errkomk = (string[])errvalue.Value.Split('-');
                        string logvalue = "";
                        string logvalue_insertqry = "";
                        logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk[0], tmp_taisyodata, errkomk[1], errkomk[2]);
                        logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, true);
                        list_log.Add(logvalue_insertqry);

                        // ログ出力
                        if (list_log.Count >= CommonModule.Log_OutputCnt | list_log.Count != 0 & cntii == rowcnt)
                        {
                            int tmp_cnt = 0;                      // ※Exec_NonQueryの第3引数の為用意
                            // 挿入
                            foreach (var loggrp in list_log)
                                DBExec.Exec_NonQuery(sqlcnnv10, loggrp, ref tmp_cnt);
                            // 初期化
                            list_log.Clear();
                        }

                    }

                    // ----- 個別用プログレスバー更新/進捗率表示 -----
                    int tmp_pgbcnt = 0;
                    if (rowcnt <= pgbbasecnt)        // 基準件数(100件)以下の場合は実件数を取得
                    {
                        tmp_pgbcnt = cntii + 1;
                    }
                    else if (rowcnt > cntii + 2)      // 基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
                    {
                        EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, ref tmp_pgbcnt);
                    }
                    else if (rowcnt <= cntii + 2)     // 基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
                    {
                        tmp_pgbcnt = pgbtotalcnt_part;
                    }
                    if (tmp_pgbcnt != 0)
                    {
                        obj_pgb.pgbsettingPart(tmp_pgbcnt);
                        obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part);
                    }
                    // 20161017 進捗表示処理による速度低下の修正 -del
                    // Me.Refresh()

                }
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
            // 'クローズ処理
            // Call excelfile.ExcelFile_ReadClose(con_read)
            // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
            chkskiplbl:
                ;

                // ----- 全体用・中間用プログレスバー更新 -----           
                pgbcnt = pgbcnt + 1;
                obj_pgb.pgbsettingTotal(pgbcnt);
                obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt_total, true);
                obj_pgb.pgbsettingChkTotal(pgbcnt);
                obj_com.ProgressChkOutPut(pgbcnt, pgbtotalcnt_total, true);
                Refresh();
                Console.WriteLine(Conversions.ToString(DateTime.Now) + " " + CommonModule.LOG_SYORIKOMK_MIDCHECK + "済：" + CommonModule.LOG_SYORIKOMK_MIDFILE + "(" + filesheetname + ")");                    // 20161009 ログ出力処理追加 -add
            }

            // チェック項目表示文字列初期化
            lblCVItem.Text = "";


            // --------------------------------------------------
            // ログ出力
            // --------------------------------------------------
            string tmp_sql_end = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, CommonModule.LOG_SYORIKOMK_MIDCHKEND), true);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_end, ref tmp_logcnt);

            return rtn;

        }

        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
        // ''' <summary>
        // ''' 汎用用中間ファイルチェック処理 '20161006 中間ファイルチェック処理の速度改善対応 -add
        // ''' </summary>
        // ''' <param name="filename"></param>
        // ''' <param name="sheetname"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Get_BaseMidFile_Data(ByVal filename As String, ByVal sheetname As String) As Boolean

        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
        // Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用
        // Dim tmp_logcnt As Integer = 0                                   'ログ出力時のソート用
        // Dim normalflg As Boolean = True                                 'INSERT正常終了フラグ
        // Dim rtn As Boolean = True                                       '戻り値

        // 'Excelファイル初期設定                  
        // Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
        // Dim readtbl As New DataTable()
        // Dim con_read As New OleDbConnection()


        // 'オープン処理
        // rtn = excelfile.ExcelFile_ReadOpen(MiddleDirPath, filename, tmp_sql, readtbl, con_read)

        // 'オープン処理失敗時は処理を抜ける
        // If rtn = False Then
        // Call excelfile.ExcelFile_ReadClose(con_read)
        // Return rtn
        // End If

        // '行数取得
        // Dim rowcnt As Integer = readtbl.Rows.Count

        // '----- 個別用プログレスバー更新 -----
        // Me.lblCVItem.Text = filename & " ： " & sheetname                'チェック項目表示
        // Dim pgbtotalcnt_part As Integer = 0                              'プログレスバー総件数初期化
        // Dim pgbbasecnt As Integer = 100                                  '実件数で表示するか否かの基準値
        // If rowcnt <= pgbbasecnt Then
        // pgbtotalcnt_part = rowcnt
        // Else
        // pgbtotalcnt_part = Math.Ceiling(rowcnt / pgbbasecnt)
        // End If
        // Call obj_pgb.pgbInitPart(pgbtotalcnt_part)

        // '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
        // Call Me.Set_MidDataInfoToObj(filename, sheetname, True)

        // '格納用変数初期化
        // Dim list_errheader As New List(Of String)
        // Dim list_chkduplicate As New List(Of String)
        // Dim list_header As New SortedList(Of Integer, String)           '自社口座情報ヘッダー格納用

        // 'データ取得(汎用中間ファイルには1行目にヘッダーが存在しないためデータとして取得し、行Noから判断する)
        // For cntii = 0 To rowcnt - 1

        // '中断処理
        // Application.DoEvents()
        // If MidChkCancelFlg Or CancelFlg Then
        // Call excelfile.ExcelFile_ReadClose(con_read)
        // Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
        // Return rtn
        // End If

        // Select Case sheetname
        // Case "自社口座情報"

        // 'データ格納用変数
        // Dim tmp_hash As New SafeDictionary<string, string>

        // If cntii = 0 Then

        // 'ヘッダー取得
        // For cntjj As Integer = 0 To readtbl.Columns.Count - 1
        // Dim fldname As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        // '20161007 汎用中間ファイルにゴミが残っている場合のヘッダーチェック処理修正 -chg sta
        // 'list_header.Add(cntjj, fldname)
        // If fldname <> "" Then
        // list_header.Add(cntjj, fldname)
        // End If
        // '20161007 汎用中間ファイルにゴミが残っている場合のヘッダーチェック処理修正 -chg end
        // Next

        // 'ヘッダーチェック
        // normalflg = Me.Chk_BaseMidFileHeader_Jisya(sheetname, list_header, list_errheader)
        // If normalflg = False Then
        // Exit For
        // End If

        // ElseIf cntii >= BASEMIDFILE_READWRITE_ROW - 2 Then

        // 'データ取得→チェック
        // For cntjj As Integer = 0 To readtbl.Columns.Count - 1
        // Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        // tmp_hash.Add(list_header(cntjj), fldvalue)
        // Next

        // '不要なフィールドを削除
        // tmp_hash.Remove("項目名")

        // Dim hash_errdata As New SafeDictionary<string, string>
        // Dim midkeytblfldname(10) As String
        // Dim midkeyvalue(10) As String

        // 'ログ出力用のキーフィールドと値を取得
        // For Each miditem In tmp_hash
        // Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        // If Hash_BaseMiddatakey.Contains(get_key) Then
        // Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
        // Dim tmp_keyno As Integer = 0
        // If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        // midkeytblfldname(tmp_keyno) = miditem.Key
        // midkeyvalue(tmp_keyno) = miditem.Value
        // End If
        // End If
        // Next

        // Dim tmp_keyfldnamejp As String
        // Dim tmp_keyvalue As String = ""
        // Dim tmp_taisyodata As String = ""
        // For cntkk = 1 To UBound(midkeytblfldname)
        // If midkeytblfldname(cntkk) <> "" Then
        // tmp_keyfldnamejp = midkeytblfldname(cntkk)
        // tmp_keyvalue = midkeyvalue(cntkk)
        // tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        // End If
        // Next
        // tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        // 'データチェック
        // Dim tmp_flg As Boolean = True
        // tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        // 'ログ出力処理
        // Dim list_log As New List(Of String)
        // For Each errvalue In hash_errdata

        // Dim tmp_tblfldname As String = errvalue.Key
        // Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        // Dim fldname As String = tmp_fldname(1)
        // Dim errkomk() As String = errvalue.Value.Split("-")
        // Dim logvalue As String = ""
        // Dim logvalue_insertqry As String = ""
        // logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        // logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // list_log.Add(logvalue_insertqry)

        // 'ログ出力
        // If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        // Dim tmp_cnt As Integer = 0              '※Exec_NonQueryの第3引数の為用意
        // '挿入
        // For Each loggrp In list_log
        // DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        // Next
        // '初期化
        // list_log.Clear()
        // End If

        // Next

        // End If

        // Case "部屋設備情報"

        // 'データ格納用変数
        // Dim tmp_hash As New SafeDictionary<string, string>

        // If cntii = 0 Then

        // 'グループヘッダー取得
        // For cntjj As Integer = 0 To readtbl.Columns.Count - 1
        // Dim fldname As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        // '20161007 汎用中間ファイルにゴミが残っている場合のヘッダーチェック処理修正 -chg sta
        // 'fldname = fldname.Replace(vbLf, STR_SPLIT_1)
        // 'list_header.Add(cntjj, fldname)
        // If fldname <> "" Then
        // fldname = fldname.Replace(vbLf, STR_SPLIT_1)
        // list_header.Add(cntjj, fldname)
        // End If
        // '20161007 汎用中間ファイルにゴミが残っている場合のヘッダーチェック処理修正 -chg end
        // Next

        // 'ヘッダーチェック
        // normalflg = Me.Chk_BaseMidFileHeader_Jisya(sheetname, list_header, list_errheader)
        // If normalflg = False Then
        // Exit For
        // End If

        // ElseIf cntii >= BASEMIDFILE_READWRITE_ROW - 2 Then

        // 'データ取得→チェック
        // For cntjj As Integer = 0 To readtbl.Columns.Count - 1
        // Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        // tmp_hash.Add(list_header(cntjj), fldvalue)
        // Next

        // '不要なフィールドを削除
        // tmp_hash.Remove("項目名")

        // Dim hash_errdata As New SafeDictionary<string, string>
        // Dim midkeytblfldname(10) As String
        // Dim midkeyvalue(10) As String

        // 'ログ出力用のキーフィールドと値を取得
        // For Each miditem In tmp_hash
        // Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        // If Hash_BaseMiddatakey.Contains(get_key) Then
        // Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
        // Dim tmp_keyno As Integer = 0
        // If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        // midkeytblfldname(tmp_keyno) = miditem.Key
        // midkeyvalue(tmp_keyno) = miditem.Value
        // End If
        // End If
        // Next

        // Dim tmp_keyfldnamejp As String
        // Dim tmp_keyvalue As String = ""
        // Dim tmp_taisyodata As String = ""
        // For cntkk = 1 To UBound(midkeytblfldname)
        // If midkeytblfldname(cntkk) <> "" Then
        // tmp_keyfldnamejp = midkeytblfldname(cntkk)
        // tmp_keyvalue = midkeyvalue(cntkk)
        // tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        // End If
        // Next
        // tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        // 'データチェック
        // Dim tmp_flg As Boolean = True
        // tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        // 'ログ出力処理
        // Dim list_log As New List(Of String)
        // For Each errvalue In hash_errdata

        // Dim tmp_tblfldname As String = errvalue.Key
        // Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        // Dim fldname As String = ""
        // If UBound(tmp_fldname) = 1 Then
        // fldname = tmp_fldname(1)
        // ElseIf UBound(tmp_fldname) = 2 Then
        // fldname = tmp_fldname(1) & "-" & tmp_fldname(2)
        // End If

        // Dim errkomk() As String = errvalue.Value.Split("-")
        // Dim logvalue As String = ""
        // Dim logvalue_insertqry As String = ""
        // logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        // logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // list_log.Add(logvalue_insertqry)

        // 'ログ出力
        // If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        // Dim tmp_cnt As Integer = 0              '※Exec_NonQueryの第3引数の為用意
        // '挿入
        // For Each loggrp In list_log
        // DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        // Next
        // '初期化
        // list_log.Clear()
        // End If

        // Next

        // End If

        // End Select


        // '----- 個人用プログレスバー更新/進捗率表示 -----
        // Dim tmp_pgbcnt As Integer = 0
        // If rowcnt <= pgbbasecnt Then        '基準件数(100件)以下の場合は実件数を取得
        // tmp_pgbcnt = cntii + 1
        // ElseIf rowcnt > cntii + 2 Then      '基準件数(100件)より大きい、かつ最終レコードに達していない場合は100件毎に値を取得
        // Call EtcMethod.Get_MultipleFlg(cntii + 2, pgbbasecnt, tmp_pgbcnt)
        // ElseIf rowcnt <= cntii + 2 Then     '基準件数(100件)より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得
        // tmp_pgbcnt = pgbtotalcnt_part
        // End If
        // If tmp_pgbcnt <> 0 Then
        // Call obj_pgb.pgbsettingPart(tmp_pgbcnt)
        // Call obj_com.ProgressOutPut(tmp_pgbcnt, pgbtotalcnt_part)
        // End If
        // '20161017 進捗表示処理による速度低下の修正 -del
        // 'Me.Refresh()

        // Next

        // 'ヘッダーエラー時の終了処理
        // If normalflg = False Then
        // For Each errheader In list_errheader
        // 'ログ用に文字列を編集してクエリへ加工
        // Dim logvalue As String = ""
        // Dim logvalue_insertqry As String = ""
        // Dim tmp_cnt As Integer = 0

        // 'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
        // If errheader <> "" Then
        // logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
        // Else
        // logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
        // End If

        // 'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
        // logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
        // Next

        // 'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
        // Call excelfile.ExcelFile_ReadClose(con_read)
        // rtn = normalflg
        // Return rtn

        // End If

        // 'クローズ処理
        // Call excelfile.ExcelFile_ReadClose(con_read)

        // '返却
        // Return rtn

        // End Function
        // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end

        // 20161006 中間ファイルチェック処理の速度改善対応 -del sta
        // ''' <summary>
        // ''' 中間ファイルチェック実行処理→ログ出力
        // ''' </summary>
        // ''' <param name="midchklist"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Get_MidFile_Data(midchklist As List(Of String)) As Boolean

        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim startrow As Integer                                         '書込開始行
        // Dim columncnt As Integer                                        '列数
        // Dim maxrowcnt As Integer                                        '既存データの行数
        // Dim rowcnt As Integer                                           '書込行数
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim tmp_logcnt As Integer                                       '作業用

        // Dim normalflg As Boolean                                        '正常終了フラグ
        // Dim rtn As Boolean = True                                       '戻り値

        // 'ログ出力
        // Log_OutputCnt = Int32.Parse(Me.txtLogOutputCnt.Text)            'ログ出力件数
        // Dim tmp_sql_sta As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTA), True)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_sta, tmp_logcnt)

        // '中間ファイルチェック処理
        // For Each filesheetname In midchklist

        // Dim tmp_str() As String = filesheetname.Split("-")
        // Dim filename As String = tmp_str(0)
        // Dim sheetname As String = tmp_str(1)
        // Dim headervalue As New Object
        // Dim fldvaluegrp As New Object
        // Dim list_chkduplicate As New List(Of String)    '20160812 汎用コンバート対応 -add

        // '汎用用中間ファイルを直接チェックする処理
        // '20160812 汎用コンバート対応 -add sta
        // Select Case filesheetname
        // Case "自社情報-自社口座情報", "部屋情報-部屋設備情報"
        // filename = "中間ファイル"
        // rtn = Me.Get_BaseMidFile_Data(filename, sheetname)
        // If rtn = False Then
        // Return rtn
        // Else
        // GoTo chkskiplbl
        // End If
        // Case "物件情報-物件変動費親メーター情報", "物件情報-物件近隣駐車場情報", "物件情報-(未使用の契約者データは移行しない)", "物件情報-(未使用の家主データは移行しない)", _
        // "部屋情報-部屋面積情報", "部屋情報-部屋権利情報", _
        // "契約情報-契約送金ルール情報", "契約情報-契約控除ルール情報", "契約情報-契約解約情報", "契約情報-契約修繕見積情報", _
        // "契約情報-契約修繕見積詳細情報", "契約情報-契約変動費親メーター情報", "契約情報-契約解約確認事項情報", "契約情報-契約同時契約情報", _
        // "契約情報-契約原状回復目安単価情報", "契約情報-契約敷金保証金随時処理情報", "契約情報-契約関連ファイル情報", "契約情報-契約空室待ち情報"
        // GoTo chkskiplbl         '不要なチェックボックスは削除すること
        // End Select
        // '20160812 汎用コンバート対応 -add end    

        // 'Excelファイル初期設定
        // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

        // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
        // If Not rtn Then
        // Return rtn
        // End If
        // '20160905 中間ファイルコピー処理改善 -add sta
        // '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
        // Call Me.Set_MidDataInfoToObj(filename, sheetname, False)
        // '20160905 中間ファイルコピー処理改善 -add end
        // For cntii = 0 To rowcnt

        // '中断処理
        // Application.DoEvents()
        // If MidChkCancelFlg Or CancelFlg Then '20160929 データチェックのキャンセル処理を追加 Or CancelFlg の条件も追加 -chg
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
        // Return rtn
        // End If

        // '格納用変数初期化
        // Dim list_errheader As New List(Of String)
        // Dim hash_errdata As New SafeDictionary<string, string>

        // 'データ取得
        // Dim rowvalue As Object
        // rowvalue = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

        // If cntii = 0 Then
        // 'ヘッダー格納用配列へコピー
        // headervalue = rowvalue.Clone

        // 'ヘッダーチェック
        // normalflg = Me.Chk_MidFileHeader(sheetname, headervalue, list_errheader)

        // 'エラー時の終了処理
        // If normalflg = False Then
        // For Each errheader In list_errheader
        // 'ログ用に文字列を編集してクエリへ加工
        // Dim logvalue As String = ""
        // Dim logvalue_insertqry As String = ""
        // Dim tmp_cnt As Integer = 0

        // 'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
        // If errheader <> "" Then
        // logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
        // Else
        // logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
        // End If

        // 'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
        // logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
        // Next

        // 'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // rtn = normalflg
        // Return rtn
        // End If
        // Else
        // '20160812 汎用コンバート対応 -chg sta
        // '                        Dim midkeytblfldname(10) As String
        // '                        Dim midkeyvalue(10) As String

        // '                        'ヘッダー格納用配列へコピー
        // '                        fldvaluegrp = rowvalue.Clone

        // '                        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
        // '                        Dim tmp_hash As New SafeDictionary<string, string>
        // '                        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(sheetname, headervalue, fldvaluegrp)

        // '                        'データチェック
        // '                        For Each fldvalue In tmp_hash

        // '                            Dim hashkey As String = fldvalue.Key
        // '                            Dim value As String = fldvalue.Value
        // '                            Dim errstr As String = ""

        // '                            'キーフィールドの判別
        // '                            If Hash_MidKey_FieldJp.Contains(hashkey) Then
        // '                                Dim tmp_nostr As String = Hash_MidKey_FieldJp.Item(hashkey)
        // '                                Dim tmp_keyno As Integer = 0
        // '                                If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        // '                                    midkeytblfldname(tmp_keyno) = hashkey
        // '                                    midkeyvalue(tmp_keyno) = value
        // '                                Else
        // '                                    errstr = LOG_NAIYO_ERR_DATACHK & "-" & LOG_HUBI_ERR_DATACHK_C & "-" & LOG_TAISYO_DATACHK_A
        // '                                    GoTo chkskiplabel
        // '                                End If
        // '                            End If

        // '                            'データチェック処理
        // '                            Call DataChk.Call_ChkMethodMidTotal(hashkey, value, errstr)
        // 'chkskiplabel:
        // '                            If errstr <> "" Then
        // '                                Dim fldname() As String = hashkey.Split("-")
        // '                                hash_errdata.Add(fldname(1), errstr)
        // '                            End If

        // '                        Next

        // '                        Dim tmp_keyfldname() As String
        // '                        Dim tmp_keyvalue As String = ""
        // '                        Dim tmp_taisyodata As String = ""
        // '                        Dim list_log As New List(Of String)

        // '                        '退避しておいたキーフィールドと値を取得
        // '                        For cntkk = 1 To UBound(midkeytblfldname)
        // '                            If midkeytblfldname(cntkk) <> "" Then
        // '                                tmp_keyfldname = midkeytblfldname(cntkk).Split("-")
        // '                                tmp_keyvalue = tmp_hash.Item(midkeytblfldname(cntkk))
        // '                                tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldname(1) & " = " & tmp_keyvalue
        // '                            End If
        // '                        Next
        // '                        tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        // '                        'ログ出力処理
        // '                        For Each errvalue In hash_errdata

        // '                            Dim fldname As String = errvalue.Key
        // '                            Dim errkomk() As String = errvalue.Value.Split("-")
        // '                            Dim logvalue As String = ""
        // '                            Dim logvalue_insertqry As String = ""
        // '                            logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        // '                            logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // '                            list_log.Add(logvalue_insertqry)

        // '                            'ログ出力
        // '                            If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        // '                                Dim tmp_cnt As Integer = 0      '※Exec_NonQueryの第3引数の為用意
        // '                                '挿入
        // '                                For Each loggrp In list_log
        // '                                    DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        // '                                Next
        // '                                '初期化
        // '                                list_log.Clear()
        // '                            End If

        // '                        Next
        // Dim midkeytblfldname(10) As String
        // Dim midkeyvalue(10) As String

        // 'ヘッダー格納用配列へコピー
        // fldvaluegrp = rowvalue.Clone

        // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
        // Dim tmp_hash As New SafeDictionary<string, string>
        // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue(sheetname, headervalue, fldvaluegrp)

        // 'ログ出力用のキーフィールドと値を取得
        // For Each miditem In tmp_hash
        // Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        // If Hash_ExistMiddatakey.Contains(get_key) Then
        // Dim tmp_nostr As String = Hash_ExistMiddatakey.Item(get_key)
        // Dim tmp_keyno As Integer = 0
        // If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        // midkeytblfldname(tmp_keyno) = miditem.Key
        // midkeyvalue(tmp_keyno) = miditem.Value
        // End If
        // End If
        // Next

        // Dim tmp_keyfldnamejp As String
        // Dim tmp_keyvalue As String = ""
        // Dim tmp_taisyodata As String = ""
        // For cntkk = 1 To UBound(midkeytblfldname)
        // If midkeytblfldname(cntkk) <> "" Then
        // tmp_keyfldnamejp = midkeytblfldname(cntkk)
        // tmp_keyvalue = midkeyvalue(cntkk)
        // tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        // End If
        // Next
        // tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        // 'データチェック
        // Dim tmp_flg As Boolean = True
        // '20160927 ログの内容が不正になっているため修正 -add sta
        // If filename = "送金ルール情報" Then
        // tmp_hash("物件No") = tmp_hash("物件No") & "-" & tmp_hash("部屋No")
        // End If
        // '20160927 ログの内容が不正になっているため修正 -add end
        // tmp_flg = Not (DataChk.Chk_DataPerItem_Mid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        // 'ログ出力処理
        // Dim list_log As New List(Of String)
        // For Each errvalue In hash_errdata

        // '20160927 ログの内容が不正になっているため修正 -chg sta
        // 'Dim tmp_tblfldname As String = errvalue.Key
        // 'Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        // 'Dim fldname As String = tmp_fldname(1)
        // Dim tmp_tblfldname As String = errvalue.Key
        // Dim tmp_sfname() As String = Split(tmp_tblfldname, STR_SPLIT_2)
        // Dim fldname As String = ""

        // For cntjj = 0 To UBound(tmp_sfname)
        // Dim tmp_sfsplietname() As String = Split(tmp_sfname(cntjj), STR_SPLIT_1)
        // fldname = fldname & "、" & tmp_sfsplietname(1)
        // Next
        // If fldname <> "" Then
        // fldname = fldname.Remove(0, 1)
        // End If
        // 'Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        // 'Dim fldname As String = tmp_fldname(1)
        // '20160927 ログの内容が不正になっているため修正 -chg end
        // Dim errkomk() As String = errvalue.Value.Split("-")
        // Dim logvalue As String = ""
        // Dim logvalue_insertqry As String = ""
        // logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        // logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // list_log.Add(logvalue_insertqry)

        // 'ログ出力
        // If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        // Dim tmp_cnt As Integer = 0      '※Exec_NonQueryの第3引数の為用意
        // '挿入
        // For Each loggrp In list_log
        // DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        // Next
        // '初期化
        // list_log.Clear()
        // End If

        // Next
        // '20160812 汎用コンバート対応 -chg end
        // End If

        // Next

        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        // chkskiplbl:     '20160812 汎用コンバート対応 -add

        // Next

        // '--------------------------------------------------
        // 'ログ出力
        // '--------------------------------------------------
        // Dim tmp_sql_end As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKEND), True)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_end, tmp_logcnt)

        // '20160907 EXCELが開いたままになっている不具合 -coment：汎用用の中間ファイルが活きているので閉じておく
        // '(別エクセルファイルを開いているとき、裏で活きたままになる)
        // Return rtn

        // End Function
        // 20161006 中間ファイルチェック処理の速度改善対応 -del end

        // 20161006 中間ファイルチェック処理の速度改善対応 -del sta
        // ''' <summary>
        // ''' 汎用用中間ファイルチェック処理 20160812 汎用コンバート対応
        // ''' </summary>
        // ''' <param name="filename"></param>
        // ''' <param name="sheetname"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Get_BaseMidFile_Data(ByVal filename As String, ByVal sheetname As String) As Boolean

        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim startrow As Integer                                         '書込開始行
        // Dim columncnt As Integer                                        '列数
        // Dim maxrowcnt As Integer                                        '既存データの行数
        // Dim rowcnt As Integer                                           '書込行数
        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        // Dim tmp_logcnt As Integer                                       '作業用

        // Dim normalflg As Boolean                                        '正常終了フラグ
        // Dim rtn As Boolean = True                                       '戻り値
        // Dim headervalue As New Object
        // Dim fldvaluegrp As New Object
        // Dim list_chkduplicate As New List(Of String)
        // '20160913_2 不要処理の削除 -del sta
        // ''ヘッダー行不定のためここで一時的に設定しておく
        // 'If sheetname = "部屋設備情報" Then                    '20160913 これなんでしたっけ？
        // '    startrow = 3
        // 'End If
        // '20160913_2 不要処理の削除 -del end
        // '20160907 EXCELが開いたままになっている不具合 -coment：オープン処理内でエラーの場合の処理(ログ出力)がない？
        // 'Excelファイル初期設定
        // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MiddleDirPath, filename, sheetname)

        // 'Excelファイル設定時にエラーが生じた際は処理を抜ける
        // If Not rtn Then
        // Return rtn
        // End If
        // '20160905 中間ファイルコピー処理改善 -add sta
        // '照合用テーブルから既存用中間ファイルの必要な情報を取得してオブジェクトへ格納
        // Call Me.Set_MidDataInfoToObj(filename, sheetname, True)
        // '20160905 中間ファイルコピー処理改善 -add end
        // '20160913_2 汎用用中間ファイルチェック処理の全体的な修正 -chg sta
        // '↓↓↓旧srcコメントアウト↓↓↓
        // 'For cntii = 0 To rowcnt

        // '    '中断処理
        // '    Application.DoEvents()
        // '    If MidChkCancelFlg Then

        // '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // '        Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
        // '        DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
        // '        Return rtn
        // '    End If

        // '    '格納用変数初期化
        // '    Dim list_errheader As New List(Of String)
        // '    Dim hash_errdata As New SafeDictionary<string, string>

        // '    'データ取得
        // '    Dim rowvalue As Object
        // '    '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg sta
        // '    'rowvalue = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value
        // '    If cntii = 0 Then                                                       '20160913 ★cnt=0の時の処理について別途教えて下さい。★
        // '        Dim headerrow As Integer = 2
        // '        rowvalue = wsheet.Range(wsheet.Cells(headerrow, 2), wsheet.Cells(headerrow, columncnt)).Value
        // '    Else
        // '        rowvalue = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 2), wsheet.Cells(cntii + startrow - 1, columncnt)).Value
        // '    End If
        // '    '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg end
        // '    If cntii = 0 Then                                               '20160913 ↑と↓の処理は纏められない？
        // '        'ヘッダー格納用配列へコピー
        // '        headervalue = rowvalue.Clone

        // '        'ヘッダーチェック
        // '        normalflg = Me.Chk_BaseMidFileHeader(sheetname, headervalue, list_errheader)

        // '        'エラー時の終了処理
        // '        If normalflg = False Then
        // '            For Each errheader In list_errheader
        // '                'ログ用に文字列を編集してクエリへ加工
        // '                Dim logvalue As String = ""
        // '                Dim logvalue_insertqry As String = ""
        // '                Dim tmp_cnt As Integer = 0

        // '                'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
        // '                If errheader <> "" Then
        // '                    logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
        // '                Else
        // '                    logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
        // '                End If

        // '                'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
        // '                logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // '                DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
        // '            Next

        // '            'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
        // '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // '            rtn = normalflg
        // '            Return rtn
        // '        End If
        // '    Else

        // '        Dim midkeytblfldname(10) As String
        // '        Dim midkeyvalue(10) As String

        // '        'ヘッダー格納用配列へコピー
        // '        fldvaluegrp = rowvalue.Clone

        // '        'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
        // '        Dim tmp_hash As New SafeDictionary<string, string>
        // '        tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue(sheetname, headervalue, fldvaluegrp)

        // '        'ログ出力用のキーフィールドと値を取得
        // '        For Each miditem In tmp_hash
        // '            Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        // '            If Hash_BaseMiddatakey.Contains(get_key) Then
        // '                Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
        // '                Dim tmp_keyno As Integer = 0
        // '                If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        // '                    midkeytblfldname(tmp_keyno) = miditem.Key
        // '                    midkeyvalue(tmp_keyno) = miditem.Value
        // '                End If
        // '            End If
        // '        Next

        // '        Dim tmp_keyfldnamejp As String
        // '        Dim tmp_keyvalue As String = ""
        // '        Dim tmp_taisyodata As String = ""
        // '        For cntkk = 1 To UBound(midkeytblfldname)
        // '            If midkeytblfldname(cntkk) <> "" Then
        // '                tmp_keyfldnamejp = midkeytblfldname(cntkk)
        // '                tmp_keyvalue = midkeyvalue(cntkk)
        // '                tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        // '            End If
        // '        Next
        // '        tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        // '        'データチェック
        // '        Dim tmp_flg As Boolean = True
        // '        tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        // '        'ログ出力処理
        // '        Dim list_log As New List(Of String)
        // '        For Each errvalue In hash_errdata

        // '            Dim tmp_tblfldname As String = errvalue.Key
        // '            Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        // '            Dim fldname As String = tmp_fldname(1)
        // '            Dim errkomk() As String = errvalue.Value.Split("-")
        // '            Dim logvalue As String = ""
        // '            Dim logvalue_insertqry As String = ""
        // '            logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        // '            logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // '            list_log.Add(logvalue_insertqry)

        // '            'ログ出力
        // '            If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        // '                Dim tmp_cnt As Integer = 0      '※Exec_NonQueryの第3引数の為用意
        // '                '挿入
        // '                For Each loggrp In list_log
        // '                    DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        // '                Next
        // '                '初期化
        // '                list_log.Clear()
        // '            End If

        // '        Next
        // '    End If

        // 'Next
        // '↑↑↑旧srcコメントアウト↑↑↑

        // '格納用変数初期化
        // Dim list_errheader As New List(Of String)
        // Dim hash_errdata As New SafeDictionary<string, string>

        // 'ヘッダー取得
        // Select Case sheetname
        // Case "自社口座情報"

        // Dim headerrow As Integer = BASEMIDFILE_HEADER_ROW
        // headervalue = wsheet.Range(wsheet.Cells(headerrow, 2), wsheet.Cells(headerrow, columncnt)).Value

        // 'ヘッダーチェック
        // normalflg = Me.Chk_BaseMidFileHeader(sheetname, headervalue, list_errheader)

        // Case "部屋設備情報"

        // Dim headerrow_grp As Integer = BASEMIDFILE_HEADER_ROW_SETUBIGRP
        // Dim headerrow_setubi As Integer = BASEMIDFILE_HEADER_ROW
        // Dim headervalue_grp As Object = wsheet.Range(wsheet.Cells(headerrow_grp, 2), wsheet.Cells(headerrow_grp, columncnt)).Value
        // Dim headervalue_setubi As Object = wsheet.Range(wsheet.Cells(headerrow_setubi, 2), wsheet.Cells(headerrow_setubi, columncnt)).Value
        // headervalue = EtcMethod.Set_Header(headervalue_grp, headervalue_setubi)

        // 'ヘッダーチェック
        // normalflg = Me.Chk_BaseMidFileHeader_Setubi(sheetname, headervalue, list_errheader)

        // End Select

        // 'エラー時の終了処理
        // If normalflg = False Then
        // For Each errheader In list_errheader
        // 'ログ用に文字列を編集してクエリへ加工
        // Dim logvalue As String = ""
        // Dim logvalue_insertqry As String = ""
        // Dim tmp_cnt As Integer = 0

        // 'ヘッダーが空の場合、ヘッダーが削除されたかヘッダー外にデータが書き込まれているためログの内容を分岐
        // If errheader <> "" Then
        // logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERCHK, errheader, LOG_HUBI_MIDHEADERCHK, LOG_TAISYO_MIDHEADERCHK)
        // Else
        // logvalue = LogSetting.Set_LogValue(101, sheetname, "-", LOG_NAIYO_ERR_MIDHEADERRANGECHK, errheader, LOG_HUBI_MIDHEADERRANGECHK, LOG_TAISYO_MIDHEADERRANGECHK)
        // End If

        // 'テーブルへ挿入(ヘッダー不正件数は多くないと考えられるため1件ずつ実行する)
        // logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, logvalue_insertqry, tmp_cnt)
        // Next

        // 'ヘッダーが不正の時点でコンバート不可であるため処理を抜ける(ヘッダー範囲外にデータが登録されている場合も同様)
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // rtn = normalflg
        // Return rtn
        // End If

        // 'データチェック
        // For cntii = 1 To rowcnt

        // '中断処理
        // Application.DoEvents()
        // If MidChkCancelFlg Or CancelFlg Then '20160929 データチェックのキャンセル処理を追加 Or CancelFlg の条件も追加 -chg
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
        // Dim tmp_sql_cancel As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, LOG_SYORIKOMK_MIDCHKSTOP), True)
        // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_cancel, tmp_logcnt)
        // Return rtn
        // End If

        // 'データ取得
        // Dim rowvalue As Object
        // rowvalue = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 2), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

        // Dim midkeytblfldname(10) As String
        // Dim midkeyvalue(10) As String

        // 'ヘッダー格納用配列へコピー
        // fldvaluegrp = rowvalue.Clone

        // 'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
        // Dim tmp_hash As New SafeDictionary<string, string>
        // '20160913_2 部屋設備移行処理の追加 -chg sta
        // 'tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue(sheetname, headervalue, fldvaluegrp)
        // Select Case sheetname
        // Case "自社口座情報"
        // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue(sheetname, headervalue, fldvaluegrp)
        // Case "部屋設備情報"
        // tmp_hash = GetHashFldToValue.Get_Hash_fldvalue_MidheaderToValue_Setubi(sheetname, headervalue, fldvaluegrp)
        // End Select
        // '20160913_2 部屋設備移行処理の追加 -chg end

        // 'ログ出力用のキーフィールドと値を取得
        // For Each miditem In tmp_hash
        // Dim get_key As String = sheetname & STR_SPLIT_1 & miditem.Key
        // If Hash_BaseMiddatakey.Contains(get_key) Then
        // Dim tmp_nostr As String = Hash_BaseMiddatakey.Item(get_key)
        // Dim tmp_keyno As Integer = 0
        // If (Int32.TryParse(tmp_nostr, tmp_keyno)) Then
        // midkeytblfldname(tmp_keyno) = miditem.Key
        // midkeyvalue(tmp_keyno) = miditem.Value
        // End If
        // End If
        // Next

        // Dim tmp_keyfldnamejp As String
        // Dim tmp_keyvalue As String = ""
        // Dim tmp_taisyodata As String = ""
        // For cntkk = 1 To UBound(midkeytblfldname)
        // If midkeytblfldname(cntkk) <> "" Then
        // tmp_keyfldnamejp = midkeytblfldname(cntkk)
        // tmp_keyvalue = midkeyvalue(cntkk)
        // tmp_taisyodata = tmp_taisyodata & "、" & tmp_keyfldnamejp & " = " & tmp_keyvalue
        // End If
        // Next
        // tmp_taisyodata = tmp_taisyodata.Remove(0, 1)

        // 'データチェック
        // Dim tmp_flg As Boolean = True
        // tmp_flg = Not (DataChk.Chk_DataPerItem_BaseMid(sheetname, tmp_hash, list_chkduplicate, midkeytblfldname, midkeyvalue, hash_errdata))

        // 'ログ出力処理
        // Dim list_log As New List(Of String)
        // For Each errvalue In hash_errdata

        // Dim tmp_tblfldname As String = errvalue.Key
        // Dim tmp_fldname() As String = Split(tmp_tblfldname, STR_SPLIT_1)
        // Dim fldname As String = tmp_fldname(1)
        // Dim errkomk() As String = errvalue.Value.Split("-")
        // Dim logvalue As String = ""
        // Dim logvalue_insertqry As String = ""
        // logvalue = LogSetting.Set_LogValue(102, sheetname, fldname, errkomk(0), tmp_taisyodata, errkomk(1), errkomk(2))
        // logvalue_insertqry = LogSetting.Get_LogTblInsertQry(logvalue, True)
        // list_log.Add(logvalue_insertqry)

        // 'ログ出力
        // If list_log.Count >= Log_OutputCnt Or (list_log.Count <> 0 And cntii = rowcnt) Then
        // Dim tmp_cnt As Integer = 0      '※Exec_NonQueryの第3引数の為用意
        // '挿入
        // For Each loggrp In list_log
        // DBExec.Exec_NonQuery(sqlcnnv10, loggrp, tmp_cnt)
        // Next
        // '初期化
        // list_log.Clear()
        // End If

        // Next

        // Next
        // '20160913_2 汎用用中間ファイルチェック処理の全体的な修正 -chg end

        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        // Return rtn

        // End Function
        // 20161006 中間ファイルチェック処理の速度改善対応 -del end

        // ''' <summary>
        // ''' 個別でチェックを行う処理
        // ''' </summary>
        // ''' <param name="filename"></param>
        // ''' <param name="sheetname"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Chk_MidFile_Chg(ByVal filename As String, ByVal sheetname As String)

        // Dim rtn As Boolean = False
        // Dim filesheetname As String = filename & "-" & sheetname

        // Select Case filesheetname

        // Case "家主情報-家主イベント情報"
        // Call Njc.Repository.Chk_middata_Repository.Chk_owdata_event_mid(sqlcnnv10, filename, sheetname) : Return rtn
        // Case "家主情報-家主メモ情報"
        // Call Njc.Repository.Chk_middata_Repository.Chk_owdata_memo_mid(sqlcnnv10, filename, sheetname) : Return rtn




        // Case Else
        // rtn = True
        // End Select

        // Return rtn

        // End Function

        // 20161006 中間ファイルチェック処理の速度改善対応 -del sta
        // ''' <summary>
        // ''' 既存用中間ファイルヘッダーチェック
        // ''' </summary>
        // ''' <param name="tblnamejp"></param>
        // ''' <param name="headervalue"></param>
        // ''' <param name="list_errheader"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Chk_MidFileHeader(ByVal tblnamejp As String, ByVal headervalue As Object, ByRef list_errheader As List(Of String)) As Boolean

        // Dim rtn As Boolean = True

        // For cntii = 1 To headervalue.Length

        // Dim tmp_header As String = tblnamejp & STR_SPLIT_1 & headervalue(1, cntii)

        // '20160812 汎用コンバート対応 -chg sta
        // '汎用-既存照合用仮テーブルから取得したものと比較するように修正
        // 'If Not List_FldNameJp.Contains(tmp_header) Then
        // '    list_errheader.Add(headervalue(1, cntii))
        // 'End If
        // If List_Existmidheader.Contains(tmp_header) = False Then
        // list_errheader.Add(headervalue(1, cntii))
        // End If
        // '20160812 汎用コンバート対応 -chg end

        // Next

        // If list_errheader.Count <> 0 Then
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20161006 中間ファイルチェック処理の速度改善対応 -del end

        /// <summary>
        /// 既存用中間ファイルヘッダーチェック '20161006 中間ファイルチェック処理の速度改善対応 -add
        /// </summary>
        /// <param name="tblnamejp"></param>
        /// <param name="dt"></param>
        /// <param name="list_errheader"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        //private bool Chk_MidFileHeader(string tblnamejp, DataTable dt, ref List<string> list_errheader)
        //{

        //    bool rtn = true;

        //    for (int cntjj = 0, loopTo = dt.Columns.Count - 1; cntjj <= loopTo; cntjj++)
        //    {
        //        string tmp_header = tblnamejp + CommonModule.STR_SPLIT_1 + dt.Columns[cntjj].ColumnName.Trim();
        //        if (CommonModule.List_Existmidheader.Contains(tmp_header) == false)
        //        {
        //            list_errheader.Add(dt.Columns[cntjj].ColumnName.Trim());
        //        }
        //    }

        //    if (list_errheader.Count != 0)
        //    {
        //        rtn = false;
        //    }

        //    return rtn;

        //}

        /// <summary>
        /// 汎用用中間ファイル自社口座情報ヘッダーチェック '20161006 中間ファイルチェック処理の速度改善対応 -add
        /// </summary>
        /// <param name="tblnamejp"></param>
        /// <param name="solist_header"></param>
        /// <param name="list_errheader"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        //private bool Chk_BaseMidFileHeader_Jisya(string tblnamejp, SortedList<int, string> solist_header, ref List<string> list_errheader)
        //{

        //    bool rtn = true;

        //    for (int cntii = 0, loopTo = solist_header.Count - 1; cntii <= loopTo; cntii++)
        //    {
        //        string tmp_header = tblnamejp + CommonModule.STR_SPLIT_1 + solist_header[cntii];
        //        if (solist_header[cntii] != "項目名")
        //        {
        //            if (CommonModule.List_Basemidheader.Contains(tmp_header) == false)
        //            {
        //                list_errheader.Add(solist_header[cntii]);
        //            }
        //        }
        //    }

        //    if (list_errheader.Count != 0)
        //    {
        //        rtn = false;
        //    }

        //    return rtn;

        //}

        // 20161006 中間ファイルチェック処理の速度改善対応 -del sta
        // ''' <summary>
        // ''' 汎用用中間ファイルヘッダーチェック
        // ''' </summary>
        // ''' <param name="tblnamejp"></param>
        // ''' <param name="headervalue"></param>
        // ''' <param name="list_errheader"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Private Function Chk_BaseMidFileHeader(ByVal tblnamejp As String, ByVal headervalue As Object, ByRef list_errheader As List(Of String)) As Boolean

        // Dim rtn As Boolean = True

        // For cntii = 1 To headervalue.Length

        // Dim tmp_header As String = tblnamejp & STR_SPLIT_1 & headervalue(1, cntii)

        // '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
        // 'If List_Basemidheader.Contains(tmp_header) = False Then
        // '    list_errheader.Add(headervalue(1, cntii))
        // 'End If
        // If headervalue(1, cntii) <> "項目名" Then
        // If List_Basemidheader.Contains(tmp_header) = False Then
        // list_errheader.Add(headervalue(1, cntii))
        // End If
        // End If
        // '20160912 中間ファイル作成に伴うプログラム修正 -chg end
        // Next

        // If list_errheader.Count <> 0 Then
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20161006 中間ファイルチェック処理の速度改善対応 -del end

        /// <summary>
        /// 汎用用中間ファイルヘッダーチェック(設備用) 20160913_2 部屋設備移行処理の追加
        /// </summary>
        /// <param name="tblnamejp"></param>
        /// <param name="headervalue"></param>
        /// <param name="list_errheader"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        //private bool Chk_BaseMidFileHeader_Setubi(string tblnamejp, string[,] headervalue, ref List<string> list_errheader)
        //{

        //    bool rtn = true;

        //    for (int cntii = 1, loopTo = Conversions.ToInteger(Operators.SubtractObject(headervalue.Length, 1)); cntii <= loopTo; cntii++)
        //    {

        //        string tmp_header = Conversions.ToString(Operators.ConcatenateObject(tblnamejp + CommonModule.STR_SPLIT_1, headervalue[0, cntii]));

        //        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(headervalue[0, cntii], "項目名", false)))
        //        {
        //            if (CommonModule.List_Basemidheader.Contains(tmp_header) == false)
        //            {
        //                list_errheader.Add(Conversions.ToString(headervalue[0, cntii]));
        //            }
        //        }

        //    }

        //    if (list_errheader.Count != 0)
        //    {
        //        rtn = false;
        //    }

        //    return rtn;

        //}

        /// <summary>
        /// 送金ルール送金先情報の送金先No、送金先口座Noの有無を確認する
        /// 管理形態が自社物件、送金保留の場合は必須ではない(無くても良い)
        /// それ以外は必須
        /// </summary>
        /// <param name="hash_miditem">中間ファイルのヘッダーと値を紐付けたオブジェクト</param>
        /// <param name="hash_log">エラー時のログ格納用オブジェクト</param>
        /// <returns>True…移行可 False…移行不可 ※戻り値は移行可/不可だがここではチェックのみなので実際の移行可/不可はコンバート処理時に判別する</returns>
        /// <remarks>20161108_2 送金ルール送金先情報の移行制御処理修正 新規追加</remarks>
        private bool Chk_SoruleSosaki(ref SafeDictionary<string, string> hash_miditem, ref SafeDictionary<string, string> hash_log)
        {

            // 送金先情報から必要なデータを取得
            string[] tmp_bkno = hash_miditem["物件No"].ToString().Split('-');     // チェック用に(「物件No-部屋No」の形で保持しているため分割)
            string sosakibkno = tmp_bkno[0];
            string sosakihyno = Conversions.ToString(hash_miditem["部屋No"]);
            string sosakiknno = Conversions.ToString(hash_miditem["送金ルール管理No"]);
            string sosakisono = Conversions.ToString(hash_miditem["送金先家主NO"]);
            string sosakisokozano = Conversions.ToString(hash_miditem["送金先口座NO"]);

            // 送金ルール情報から管理形態を取得
            string tmp_sql = "";
            tmp_sql = tmp_sql + " SELECT 管理形態, [仲介物件 - 新規契約業務], [仲介物件 - 契約更新業務], [仲介物件 - 解約業務] FROM CVTBL_送金ルール基本情報 ";
            tmp_sql = tmp_sql + " WHERE [物件No] = '" + sosakibkno + "'";
            tmp_sql = tmp_sql + " AND   [部屋No] = '" + sosakihyno + "'";
            tmp_sql = tmp_sql + " AND   [送金ルール管理No] = '" + sosakiknno + "'";
            SafeDictionary<string, object> hashtbl = [];
            DBExec.Exec_DataReader(tmp_sql, ref sqlcnnv10, ref hashtbl);
            string soknkei = hashtbl["管理形態"].ToString();
            int.TryParse(hashtbl["仲介物件 - 新規契約業務"].ToString(), out int chukaiShinki);
            int.TryParse(hashtbl["仲介物件 - 契約更新業務"].ToString(), out int chukaiKoshin);
            int.TryParse(hashtbl["仲介物件 - 解約業務"].ToString(), out int chukaiKaiyaku);

            // 管理形態による送金先No、送金先口座Noの必須チェック
            string log_key = "";
            string log_value = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED;

            if ((soknkei == "3" && chukaiShinki == 0 && chukaiKoshin == 0 && chukaiKaiyaku == 0) | soknkei == "4" | soknkei == "5")
            {
            }
            // 自社物件、送金保留の物件は必須ではないためチェック処理を行わない
            // 自社物件、送金保留以外は必須チェックを行う
            // 仲介物件かつ新規契約業務、契約更新業務、解約業務でない場合も必須ではないためチェック処理を行わない
            else if (string.IsNullOrEmpty(sosakisono))
            {
                // ログ出力
                hash_log.Clear();
                log_key = "送金ルール送金先情報" + CommonModule.STR_SPLIT_1 + "送金先家主NO";
                hash_log.Add(log_key, log_value);
                return false;
            }
            else if (string.IsNullOrEmpty(sosakisokozano))
            {
                // ログ出力
                hash_log.Clear();
                log_key = "送金ルール送金先情報" + CommonModule.STR_SPLIT_1 + "送金先口座NO";
                hash_log.Add(log_key, log_value);
                return false;
            }

            return true;

        }


        /// <summary>
        /// 物件・部屋所有者情報重複チェック
        /// </summary>
        /// <param name="hash_miditem">物件所有者情報1レコードを格納したオブジェクト</param>
        /// <param name="hash_log">ログ格納用オブジェクト</param>
        /// <remarks>
        /// 20161124 物件部屋所有者情報が両方存在する場合のチェック機能の追加 新規追加
        /// 物件所有者情報チェック時に部屋所有者情報と照合して重複した場合にログを出力する
        /// </remarks>
        private void Chk_BkHySyo(ref SafeDictionary<string, string> hash_miditem, ref SafeDictionary<string, string> hash_log)
        {

            // 物件所有者情報から必要な情報を取得
            string tmp_bkno = Conversions.ToString(hash_miditem["物件No"]);
            string tmp_knno = Conversions.ToString(hash_miditem["管理No"]);

            // 部屋所有者情報に存在するか確認
            string tmp_sql = "";
            tmp_sql = tmp_sql + " SELECT COUNT(*) FROM CVTBL_部屋所有者情報 ";
            tmp_sql = tmp_sql + " WHERE [物件No] = '" + tmp_bkno + "'";
            tmp_sql = tmp_sql + " AND   [管理No] = '" + tmp_knno + "'";
            string hysyocnt = Conversions.ToString(DBExec.Exec_Scalar(tmp_sql, ref sqlcnnv10));

            // データが存在する場合は区分所有として移行される内容をログに出力する
            // ※既に物件Noに関するエラーがある場合はキーのエラーなのでそちらの内容を優先する
            if (hysyocnt != "0")
            {

                // 20161124 レビュー結果：文言修正「log_err=物件所有者情報と部屋所有者情報へ重複してデータが登録されています。」
                // 20161124 レビュー結果：文言修正「log_taisyo=区分所有として移行されます。(部屋所有者情報優先)」
                // 20161124 レビュー結果：文言修正「一棟所有の場合は物件所有者情報のみ、区分所有の場合は部屋所有者情報のみに登録を行って下さい。」

                string log_key = "物件所有者情報" + CommonModule.STR_SPLIT_1 + "物件No";
                if (hash_log.ContainsKey(log_key) == false)
                {
                    // 20161125 レビュー指摘事項対応 -chg sta
                    // Dim log_err As String = "物件所有者情報と部屋所有者情報が重複しています。"
                    // Dim log_taisyo As String = "部屋所有者情報が優先され区分所有として移行されます。" & _
                    // "一棟所有物件として移行する場合は部屋所有者情報を削除して下さい。"
                    string log_err = "物件所有者情報と部屋所有者情報へ重複してデータが登録されています。";
                    string log_taisyo = "区分所有として移行されます。(部屋所有者情報優先)" + "一棟所有の場合は物件所有者情報のみ、区分所有の場合は部屋所有者情報のみに登録を行って下さい。";
                    // 20161125 レビュー指摘事項対応 -chg end
                    string log_value = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + log_err + "-" + log_taisyo;
                    hash_log.Add(log_key, log_value);
                }

            }

        }

        #endregion

        #region 画面制御：共通処理

        /// <summary>
        /// 全移行対象ブロック名取得
        /// </summary>
        /// <remarks>
        /// ・全移行ブロック(物件情報、修繕情報など)を取得
        /// →必要な移行ブロックは全てコンバーターの対象項目タブにチェックボックスで用意する
        /// ・グループボックス、チェックボックスの名称を取得
        /// ※上記より、対処項目タブ内のチェックボックス名を変更する場合、関連する内部処理も修正する
        /// </remarks>
        private void Set_CVItemName()
        {

            // マスタ系
            model_cvitem.MstBkbrui = chkMstBkrui.Text;
            model_cvitem.MstBusKotu = chkMstBusKotu.Text;
            model_cvitem.MstBus = chkMstBus.Text;
            model_cvitem.MstHyrui = chkMstHyrui.Text;
            model_cvitem.MstKagititle = chkMstKagititle.Text;
            model_cvitem.MstKasyoClaimrui = chkMstKasyoClaimrui.Text;
            model_cvitem.MstTokuyaku = chkMstTokuyaku.Text;
            // 20160516 契約分類削除に伴うコメントアウト -del
            // model_cvitem.MstKeiyakurui = Me.chkMstKeiyakurui.Text
            model_cvitem.MstHokenrui = chkMstHokenrui.Text;
            model_cvitem.MstNkinkbn = chkMstNkinkbn.Text;
            model_cvitem.MstSchool = chkMstSchool.Text;
            model_cvitem.MstArea = chkMstArea.Text;
            model_cvitem.MstKozasyubetu = chkMstKozasyubetu.Text;
            model_cvitem.MstKozo = chkMstKozo.Text;
            model_cvitem.MstTorihikitaiyo = chkMstTorihikitaiyo.Text;
            model_cvitem.MstNkinkomok = chkMstNkinkomok.Text;
            model_cvitem.MstSetubi = chkMstSetubi.Text;
            model_cvitem.MstHendo = chkMstHendo.Text;
            model_cvitem.MstHendoitiran = chkMstHendoitiran.Text;
            model_cvitem.MstBikotitle = chkMstBikotitle.Text;
            model_cvitem.MstBikolst = chkMstBikolst.Text;
            model_cvitem.MstGazotitle = chkMstGazotitle.Text;                     // 20160801 タイトルマスタ統合処理 -add

            // 業者
            model_cvitem.GyCyukaiBase = chkGyCyukaiBase.Text;
            model_cvitem.GyCyukaiKoza = chkGyCyukaiKoza.Text;
            model_cvitem.GyCyukaiMemo = chkGyCyukaiMemo.Text;                     // 汎用分(V7には無い)
            model_cvitem.GyHokenBase = chkGyHokenBase.Text;
            model_cvitem.GyHokenKoza = chkGyHokenKoza.Text;
            model_cvitem.GyHokenMemo = chkGyHokenMemo.Text;                       // 汎用分(V7には無い)
            model_cvitem.GyYatinhosyoBase = chkGyYatinhosyoBase.Text;
            model_cvitem.GyYatinhosyoKoza = chkGyYatinhosyoKoza.Text;             // 革命10に過去にあったが現在(2015/10/23)は無くなっている
            model_cvitem.GyYatinhosyoMemo = chkGyYatinhosyoMemo.Text;             // 汎用分(V7には無い)
            model_cvitem.GySyuzenBase = chkGySyuzenBase.Text;
            model_cvitem.GySyuzenKoza = chkGySyuzenKoza.Text;
            model_cvitem.GySyuzenMemo = chkGySyuzenMemo.Text;                     // 汎用分(V7には無い)
            model_cvitem.GyLifelineBase = chkGyLifelineBase.Text;
            model_cvitem.GySekoBase = chkGySekoBase.Text;
            model_cvitem.GySisetuBase = chkGySisetuBase.Text;

            // 自社情報
            model_cvitem.JisyaBase = chkJisyaBase.Text;
            model_cvitem.JisyaKoza = chkJisyaKoza.Text;
            model_cvitem.JisyaTanto = chkJisyaTanto.Text;
            model_cvitem.JisyaMemo = chkJisyaMemo.Text;
            model_cvitem.FBFuriirai = chkFBFuriirai.Text;
            model_cvitem.FBFuritesuryo = chkFBFuritesuryo.Text;
            model_cvitem.FBKozafurikae = chkFBKozafurikae.Text;
            model_cvitem.FBNsSyutoku = chkFBNsSyutoku.Text;
            model_cvitem.MstYatinKoza = chkMstYatinKoza.Text;
            model_cvitem.MstANSERArea = chkMstANSERArea.Text;
            model_cvitem.MstANSERAccpoint = chkMstANSERAccpoint.Text;
            model_cvitem.FBANSERSetuzoku = chkFBANSERSetuzoku.Text;

            // 家主情報
            model_cvitem.OwBase = chkOwBase.Text;
            model_cvitem.OwKoza = chkOwKoza.Text;
            model_cvitem.OwEvent = chkOwEvent.Text;
            model_cvitem.OwMemo = chkOwMemo.Text;

            // 契約者情報
            model_cvitem.KysBase = chkKysBase.Text;
            model_cvitem.KysKoza = chkKysKoza.Text;
            model_cvitem.KysMemo = chkKysMemo.Text;
            model_cvitem.KysSyogoKana = chkKysSyogoKana.Text;
            model_cvitem.KysHosyonin = chkKysHosyonin.Text;

            // 物件情報
            model_cvitem.BkBase = chkBkBase.Text;
            model_cvitem.BkSyosai = chkBkSyosai.Text;
            model_cvitem.Bksyo = chkBkSyo.Text;
            model_cvitem.BkGomi = chkBkGomi.Text;
            model_cvitem.BkKenri = chkBkKenri.Text;
            model_cvitem.BkKotu = chkBkKotu.Text;
            model_cvitem.BkSetudo = chkBkSetudo.Text;
            model_cvitem.BkSyuhen = chkBkSyuhen.Text;
            model_cvitem.BkSzeniji = chkBkSzeniji.Text;
            model_cvitem.BkMemo = chkBkMemo.Text;
            model_cvitem.BkKagi = chkBkKagi.Text;                                 // 汎用分(V7には無い)   
            model_cvitem.BkHendo = chkBkHendo.Text;                               // 汎用分(V7には無い)
            model_cvitem.BkKinrincyusyajo = chkBkKinrincyusyajo.Text;             // 汎用分(V7には無い)
            model_cvitem.BkSansyofile = chkBkSansyofile.Text;                     // 汎用分(V7には無い)

            // 部屋情報
            model_cvitem.HyBase = chkHyBase.Text;
            model_cvitem.HySyosai = chkHySyosai.Text;
            model_cvitem.Hysyo = chkHySyo.Text;
            model_cvitem.HyParking = chkHyParking.Text;
            model_cvitem.HyTokuyaku = chkHyTokuyaku.Text;
            model_cvitem.HyKagi = chkHyKagi.Text;
            model_cvitem.HyMadoriutiwake = chkHyMadoriutiwake.Text;
            model_cvitem.HyMenseki = chkHyMenseki.Text;
            model_cvitem.HySetubi = chkHySetubi.Text;
            model_cvitem.HyNkinkomk = chkHyNkinkomk.Text;
            model_cvitem.HyHendo = chkHyHendo.Text;
            model_cvitem.HyMemo = chkHyMemo.Text;
            model_cvitem.HySzeniji = chkHySzeniji.Text;
            model_cvitem.HyCommonsalespoint = chkHyCommonsalespoint.Text;         // 汎用分(V7には無い)
            model_cvitem.HyConfirm = chkHyConfirm.Text;                           // 汎用分(V7には無い)
            model_cvitem.HyKenri = chkHyKenri.Text;                               // 汎用分(V7には無い)
            model_cvitem.HySansyofile = chkHySansyofile.Text;                     // 汎用分(V7には無い)
            model_cvitem.HyGenjotanka = chkHyGenjotanka.Text;                     // 汎用分(V7には無い)

            // 送金ルール
            model_cvitem.SoruleBase = chkSoruleBase.Text;
            model_cvitem.SoruleSosaki = chkSoruleSosaki.Text;
            model_cvitem.SoruleNkin = chkSoruleNkin.Text;
            model_cvitem.SoruleKojo = chkSoruleKojo.Text;

            // 契約情報
            model_cvitem.KyBase = chkKyBase.Text;
            model_cvitem.KyRireki = chkKyRireki.Text;
            model_cvitem.KyCar = chkKyCar.Text;
            model_cvitem.KyKys = chkKyKys.Text;
            model_cvitem.KyHosyonin = chkKyHosyonin.Text;
            model_cvitem.KyNyukyo = chkKyNyukyo.Text;
            model_cvitem.KyTokuyaku = chkKyTokuyaku.Text;
            model_cvitem.KyHoken = chkKyHoken.Text;
            model_cvitem.KyMemo = chkKyMemo.Text;
            model_cvitem.KyNkinkomk = chkKyNkinkomk.Text;
            model_cvitem.KyNkinkomkNx = chkKyNkinkomkNx.Text;
            model_cvitem.KyHendo = chkKyHendo.Text;
            model_cvitem.KyKojoRule = chkKyKojoRule.Text;
            model_cvitem.KySorule = chkKySorule.Text;
            model_cvitem.KyKai = chkKyKai.Text;
            // 20160531 鍵情報移行処理の修正 -del sta
            // '2016.04.06 契約鍵情報の移行処理追加 -add
            // model_cvitem.KyKagi = Me.chkKyKagi.Text
            // 20160531 鍵情報移行処理の修正 -del end
            model_cvitem.KySzen = chkKySzen.Text;
            model_cvitem.KySzenmeisai = chkKySzenmeisai.Text;

            // 請求情報
            model_cvitem.SqKajyo = chkSqKajyo.Text;
            model_cvitem.SqUnyotaino = chkSqUnyotaino.Text;
            model_cvitem.SqSq = chkSqSq.Text;
            model_cvitem.SqHendokensin = chkSqHendokensin.Text;
            model_cvitem.SqKoteiKojo = chkSqKoteiKojo.Text;
            model_cvitem.SqSqKojo = chkSqSqKojo.Text;

            // クレーム情報
            model_cvitem.ClaimBase = chkClaimBase.Text;
            model_cvitem.ClaimTaiorireki = chkClaimTaiorireki.Text;
            model_cvitem.ClaimRelfile = chkClaimRelfile.Text;
            model_cvitem.SzenBase = chkSzenBase.Text;
            model_cvitem.SzenSzen = chkSzenSzen.Text;
            model_cvitem.SzenSzenmeisai = chkSzenSzenmeisai.Text;
            model_cvitem.SzenClaim = chkSzenClaim.Text;
            model_cvitem.SzenRelfile = chkSzenRelfile.Text;                   	// 20160627 修繕関連ファイル移行修正 -add
            model_cvitem.SzenMemo = chkSzenMemo.Text;

            // 初期設定情報
            model_cvitem.SyskanriBase = chkSyskanriBase.Text;
            model_cvitem.SyskanriZei = chkSyskanriZei.Text;
            model_cvitem.SyskanriHenkanmoji = chkSyskanriHenkanmoji.Text;
            model_cvitem.SyskanriNkinkomkmerge = chkSyskanriNkinkomkmerge.Text;

            // 物件データ連動情報 													'20160720 連動情報構築 -add
            model_cvitem.RendoSosinBase = chkRendoSosinBase.Text;
            model_cvitem.RendoSosinJisyaweb = chkRendoSosinJisyaweb.Text;
            model_cvitem.RendoSosinHomes = chkRendoSosinHomes.Text;
            model_cvitem.RendoSosinAthome = chkRendoSosinAthome.Text;
            model_cvitem.RendoSosinSuumo = chkRendoSosinSuumo.Text;
            model_cvitem.RendoMapdisp = chkRendoMapdisp.Text;
            model_cvitem.RendoBtoBgroup = chkRendoBtoBgroup.Text;
            model_cvitem.RendoHysosin = chkRendoHysosin.Text;
            model_cvitem.RendoHyrui = chkRendoHyrui.Text;
            model_cvitem.RendoKokokuSuumo = chkRendoKokokuSuumo.Text;
            model_cvitem.RendoKokokuAthome = chkRendoKokokuAthome.Text;
            model_cvitem.RendoKokokuHomes = chkRendoKokokuHomes.Text;
            model_cvitem.RendoKokokuJisyaweb = chkRendoKokokuJisyaweb.Text;

        }

        /// <summary>
        /// チェックONのチェックボックス数を取得
        /// </summary>
        /// <param name="container"></param>
        /// <param name="int"></param>
        /// <remarks></remarks>
        //private void Set_ChkBoxCount(Control container, ref int @int)
        //{

        //    foreach (Control item in container.Controls)
        //    {
        //        if (item.GetType().Equals(typeof(CheckBox)))
        //        {
        //            if (((CheckBox)item).Checked)
        //            {
        //                @int = @int + 1;
        //            }
        //        }
        //    }

        //}

        /// <summary>
        /// チェック項目取得→リスト(オブジェクト)へ格納 '20160719 中間ファイルチェック処理時のエラー対応 引数に中間ファイルチェックフラグを追加(Optional)
        /// </summary>
        /// <returns></returns>
        /// <remarks>
        /// ・仕様上、全タブを表示する
        /// </remarks>
        private List<string> Get_ListCVChkitem(bool midchkflg = false)
        {

            var rtn_list = new List<string>();
            string tmp_grpname = "";
            string tmp_chkname = "";
            // 20160720 連動情報構築 -add リストオブジェクトへDC_SELTAB_B11を追加
            // 20160707 作業用DB作成時エラー修正 -add
            var list_tabpagename = new List<string>() { DC_SELTAB_B01, DC_SELTAB_B02, DC_SELTAB_B03, DC_SELTAB_B04, DC_SELTAB_B05, DC_SELTAB_B06, DC_SELTAB_B07, DC_SELTAB_B08, DC_SELTAB_B09, DC_SELTAB_B10, DC_SELTAB_B11 };

            // 20160707 作業用DB作成時エラー修正 -chg sta
            // もともと作成していたテーブル毎のチェックボックスを見に行く必要があるためtabPageBase110～200の表示処理を追加
            // tabPageKizon110～130のチェックボックスは見ていないのでコメントアウト(残しておくと実行ボタン押下時に一瞬画面が変化する)
            // tabPageManagerCvitem.ChangeTabPageVisible(0, True)
            // tabPageManagerCvitem.ChangeTabPageVisible(1, True)
            // tabPageManagerCvitem.ChangeTabPageVisible(2, True)
            tabPageManagerCvitem.ChangeTabPageVisible(3, true);
            tabPageManagerCvitem.ChangeTabPageVisible(4, true);
            tabPageManagerCvitem.ChangeTabPageVisible(5, true);
            tabPageManagerCvitem.ChangeTabPageVisible(6, true);
            tabPageManagerCvitem.ChangeTabPageVisible(7, true);
            tabPageManagerCvitem.ChangeTabPageVisible(8, true);
            tabPageManagerCvitem.ChangeTabPageVisible(9, true);
            tabPageManagerCvitem.ChangeTabPageVisible(10, true);
            tabPageManagerCvitem.ChangeTabPageVisible(11, true);
            tabPageManagerCvitem.ChangeTabPageVisible(12, true);
            tabPageManagerCvitem.ChangeTabPageVisible(13, true); // 20160720 連動情報構築 -add
            // 20160707 作業用DB作成時エラー修正 -chg end

            // タブコントロール内
            foreach (Control item in tabCtrlCVItem.Controls)
            {
                if (item.GetType().Equals(typeof(TabPage)))
                {
                    TabPage container_tab = (TabPage)item;

                    if (list_tabpagename.Contains(container_tab.Name.Trim()))
                    {
                        // タブページ内
                        foreach (Control item_sub1 in container_tab.Controls)
                        {
                            if (item_sub1.GetType().Equals(typeof(GroupBox)))
                            {

                                GroupBox container_grp = (GroupBox)item_sub1;

                                // 移行グループ名取得
                                tmp_grpname = container_grp.Text;

                                // グループボックス内
                                foreach (Control item_sub2 in container_grp.Controls)
                                {
                                    if (item_sub2.GetType().Equals(typeof(CheckBox)))
                                    {
                                        CheckBox control_chk = (CheckBox)item_sub2;
                                        switch (CommonModule.CNVNO)
                                        {

                                            case (int)CommonModule.ConvertTypes._汎用:
                                                {
                                                    if (list_existCVchkitem.Contains(control_chk) == false)
                                                    {
                                                        if (control_chk.Checked)
                                                        {
                                                            tmp_chkname = control_chk.Text;
                                                            // リスト作成
                                                            rtn_list.Add(Strings.Trim(tmp_grpname) + "-" + Strings.Trim(tmp_chkname));
                                                            Console.WriteLine(control_chk.Text);
                                                        }
                                                    }

                                                    break;
                                                }
                                        }
                                    }
                                }

                            }
                        }
                    }

                }
            }

            // ※表示状態を戻す
            if (midchkflg)
            {
                string seltab = tabCtrlCVItem.SelectedTab.Name;
                int seltabindex = 0;
                int seltabindexhid1 = 0;
                int seltabindexhid2 = 0;
                switch (seltab ?? "")
                {
                    case "tabPageKizon110":
                        {
                            seltabindex = 0;
                            seltabindexhid1 = 1;
                            seltabindexhid2 = 2;
                            break;
                        }
                    case "tabPageKizon120":
                        {
                            seltabindex = 1;
                            seltabindexhid1 = 0;
                            seltabindexhid2 = 2;
                            break;
                        }
                    case "tabPageKizon130":
                        {
                            seltabindex = 2;
                            seltabindexhid1 = 0;
                            seltabindexhid2 = 1;
                            break;
                        }
                }
                tabPageManagerCvitem.ChangeTabPageVisible(seltabindex, true);
                tabPageManagerCvitem.ChangeTabPageVisible(seltabindexhid1, false);
                tabPageManagerCvitem.ChangeTabPageVisible(seltabindexhid2, false);
            }
            else
            {
                tabPageManagerCvitem.ChangeTabPageVisible(0, false);
                tabPageManagerCvitem.ChangeTabPageVisible(1, false);
                tabPageManagerCvitem.ChangeTabPageVisible(2, true);
            }
            tabPageManagerCvitem.ChangeTabPageVisible(3, false);
            tabPageManagerCvitem.ChangeTabPageVisible(4, false);
            tabPageManagerCvitem.ChangeTabPageVisible(5, false);
            tabPageManagerCvitem.ChangeTabPageVisible(6, false);
            tabPageManagerCvitem.ChangeTabPageVisible(7, false);
            tabPageManagerCvitem.ChangeTabPageVisible(8, false);
            tabPageManagerCvitem.ChangeTabPageVisible(9, false);
            tabPageManagerCvitem.ChangeTabPageVisible(10, false);
            tabPageManagerCvitem.ChangeTabPageVisible(11, false);
            tabPageManagerCvitem.ChangeTabPageVisible(12, false);



            Console.WriteLine("dddd");
            foreach (var ee in rtn_list)
                Console.WriteLine(ee);
            Console.WriteLine("huju");

            return rtn_list;

        }

        /// <summary>
        /// ボタン状態変更処理
        /// </summary>
        /// <param name="status">設定条件文字列</param>
        /// <remarks></remarks>
        private void Chg_BtnKariStatus(string status)
        {

            btnBack.Visible = true;
            btnNext.Visible = true;
            btnEnd.Visible = true;
            btnBack.Enabled = true;
            btnNext.Enabled = true;
            btnEnd.Enabled = true;

            switch (status ?? "")
            {

                // --- メイン画面 ---
                case "mstart":
                    {
                        btnBack.Text = "";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   終  了";
                        btnBack.Visible = false;
                        break;
                    }
                case "msession":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   終  了";
                        if (btnConnectTest.Enabled)
                        {
                            btnNext.Enabled = false;
                        }

                        break;
                    }
                case "msyoki":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   終  了";
                        break;
                    }
                case "mmenu":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   終  了";
                        btnNext.Enabled = false;
                        break;
                    }

                // --- 作業選択 ---
                case "jizen1":
                    {
                        btnBack.Text = "";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = false;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }
                case "jizen2":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = true;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }
                case "jizen3":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = true;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }
                case "jizen4":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "作業選択へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = true;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }

                case "jigo1":
                    {
                        btnBack.Text = "";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = false;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }
                case "jigo2":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "作業選択へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = true;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }

                case "hojyo1":
                    {
                        btnBack.Text = "";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = false;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }
                case "hojyo2":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "作業選択へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = true;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }

                // --- 作業選択(汎用) ---
                case "jizenh":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "作業選択へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = false;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }
                case "jigoh":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "作業選択へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = false;
                        btnNext.Visible = true;
                        btnEnd.Visible = false;
                        break;
                    }

                // --- コンバート画面 ---
                case "dchajimeni":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = false;
                        btnNext.Enabled = !btnDoui.Enabled;
                        break;
                    }
                case "dcselect1":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        break;
                    }
                case "dcselect2":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        break;
                    }
                case "dcselect3":
                    {
                        // ※ボタン押下イベント(btnNext_Click)にて、ボタン名が「実行」の場合にメイン処理を実行(ボタン名変更不可)
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   実  行";
                        btnEnd.Text = "   中  止";
                        break;
                    }
                case "dcjikko1":
                    {
                        // ※ボタン押下イベント(btnEnd_Click)にて、ボタン名が「キャンセル」の場合に終了処理を実行(ボタン名変更不可)
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   実行中";
                        btnEnd.Text = "キャンセル";
                        btnBack.Enabled = false;
                        btnNext.Enabled = false;
                        break;
                    }
                case "dcjikko2":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   中  止";
                        btnBack.Visible = false;
                        btnEnd.Visible = false;
                        break;
                    }
                case "dcend":
                    {
                        btnBack.Text = "";
                        btnNext.Text = " ログを開く";
                        btnEnd.Text = "作業選択へ";
                        btnBack.Visible = false;
                        break;
                    }

                // --- 中間ファイルチェック ---
                case "midstop":
                    {
                        btnMidFileCheck.Text = " キャンセル";
                        btnBack.Enabled = false;
                        btnNext.Enabled = false;
                        btnEnd.Enabled = false;
                        btnAllChk.Enabled = false;
                        break;
                    }
                case "midcheck":
                    {
                        btnMidFileCheck.Text = " チェック";
                        btnBack.Enabled = true;
                        btnNext.Enabled = true;
                        btnEnd.Enabled = true;
                        btnAllChk.Enabled = true;
                        break;
                    }

                // --- 99.その他 ---
                case "etc":
                    {
                        btnBack.Text = "   戻  る";
                        btnNext.Text = "   次  へ";
                        btnEnd.Text = "   終  了";
                        btnBack.Visible = false;
                        btnNext.Visible = false;
                        btnEnd.Visible = false;
                        break;
                    }

                default:
                    {
                        break;
                    }

            }

        }

        /// <summary>
        /// セットフォーカス
        /// </summary>
        /// <param name="ctrlmoji">セット用文字列</param>
        /// <remarks>
        /// ・イベントやコントロール遷移時にフォーカスをあてる
        /// </remarks>
        private void Set_Forcus(string ctrlmoji)
        {

            switch (ctrlmoji ?? "")
            {

                // --- メイン画面 ---
                case "prev":                             // [戻る]ボタンにセット
                    {
                        btnBack.Focus();
                        break;
                    }
                case "next":                             // [次へ]ボタンにセット
                    {
                        btnNext.Focus();
                        break;
                    }
                case "end":                              // [終了]ボタンにセット]
                    {
                        btnEnd.Focus();
                        break;
                    }

                // --- 接続設定 ---
                case "testset":                          // 接続設定.[接続テスト]ボタンへセット
                    {
                        btnConnectTest.Focus();
                        break;
                    }

                // --- 作業選択 ---
                case "menu_jizen":
                    {
                        btnMenuJizen.Focus();                // 作業選択.[事前作業]ボタンへセット
                        break;
                    }
                case "menu_datacv":                      // 作業選択.[データコンバート]ボタンへセット
                    {
                        btnMenuDatacv.Focus();
                        break;
                    }
                case "menu_jigo":                        // 作業選択.[事後作業]ボタンへセット
                    {
                        btnMenuJigo.Focus();
                        break;
                    }

                // -- 事前作業4 --
                // --- 事前作業(汎用)
                case "jizen_tyukan":
                    {
                        chkHJizenTyukan.Focus();
                        break;
                    }

                // --- コンバーター ---
                case "dcdoui":                           // はじめに.[同意]ボタンにセット
                    {
                        btnDoui.Focus();
                        break;
                    }

                default:
                    {
                        break;
                    }

            }

        }

        /// <summary>
        /// 移行項目から紐付項目をセット → コマンドライン用に成形
        /// </summary>
        /// <param name="list_cvitem"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private string Set_CVItemChkToCmdLine(List<string> list_cvitem)
        {

            string tmp_bkrui = "";
            string tmp_hyrui = "";
            string tmp_nkinkbn = "";
            string tmp_toritaiyo = "";
            string tmp_nkinkomk = "";
            string tmp_kozo = "";
            string tmp_kozasyu = "";
            string tmp_setubi = "";
            string tmp_kagi = "";
            string tmp_jisyakoza = "";
            string tmp_bktosiyoto = "";
            string tmp_kyrui = "";
            string tmp_sohoken = "";
            string tmp_gazo = "";
            string tmp_fbfmt = "";
            string rtn_str = "";


            foreach (var item in list_cvitem)
            {

                string[] tmp_item = item.Split('-');
                string tmp_grpname = tmp_item[0];
                string tmp_komkname = tmp_item[1];

                // 20161012 紐付ツール呼出起動時の項目過不足修正 -chg sta
                // グループ名から判別可能なものを取得
                // Select Case tmp_grpname
                // Case "物件情報"
                // tmp_bkrui = "bkruichk"
                // tmp_kozo = "kozochk"
                // tmp_kagi = "kagichk"
                // tmp_bktosiyoto = "tosiyotochk"
                // tmp_gazo = "gazochk"
                // Case "部屋情報"
                // tmp_hyrui = "hyruichk"
                // tmp_kagi = "kagichk"
                // tmp_kyrui = "kyruichk"
                // tmp_nkinkbn = "nkinkbnchk"
                // tmp_toritaiyo = "toritaiyochk"
                // tmp_nkinkomk = "nkinkomkchk"
                // tmp_setubi = "setubichk"
                // tmp_sohoken = "sohokenchk"
                // tmp_gazo = "gazochk"
                // Case "契約情報"
                // tmp_kagi = "kagichk"
                // tmp_kyrui = "kyruichk"
                // tmp_nkinkbn = "nkinkbnchk"
                // tmp_nkinkomk = "nkinkomkchk"
                // tmp_sohoken = "sohokenchk"
                // Case "送金ルール情報"
                // tmp_nkinkbn = "nkinkbnchk"
                // tmp_nkinkomk = "nkinkomkchk"
                // tmp_sohoken = "sohokenchk"
                // Case Else
                // 'グループ名から判別できないものを項目名から取得
                // Select Case True
                // '※メインの方へも反映させる
                // Case tmp_komkname = "自社口座情報" Or tmp_komkname = "振込依頼人情報" Or tmp_komkname = "口座振替情報"
                // tmp_jisyakoza = "jisyakozachk"
                // tmp_nkinkbn = "nkinkbnchk"
                // tmp_kozasyu = "kozasyuchk"
                // tmp_fbfmt = "fbfmtchk"
                // Case tmp_komkname <> tmp_komkname.Replace("口座情報", "")
                // tmp_nkinkbn = "nkinkbnchk"
                // tmp_kozasyu = "kozasyuchk"
                // Case tmp_komkname <> tmp_komkname.Replace("入金項目情報", "")
                // tmp_nkinkomk = "nkinkomkchk"
                // Case tmp_komkname <> tmp_komkname.Replace("変動費", "")
                // tmp_nkinkomk = "nkinkomkchk"
                // Case tmp_komkname = "鍵タイトルマスタ"
                // tmp_kagi = "kagichk"
                // Case Else

                // End Select
                // End Select
                switch (true)
                {
                    // ※メインの方へも反映させる
                    case object _ when tmp_komkname == "自社口座情報" | tmp_komkname == "振込依頼人情報" | tmp_komkname == "口座振替情報":
                        {
                            tmp_jisyakoza = "jisyakozachk";
                            // '20161025 紐付不要項目の除去 -del
                            // tmp_nkinkbn = "nkinkbnchk"
                            tmp_kozasyu = "kozasyuchk";
                            tmp_fbfmt = "fbfmtchk";
                            break;
                        }
                    case object _ when (tmp_komkname ?? "") != (tmp_komkname.Replace("口座情報", "") ?? ""):
                        {
                            // '20161025 紐付不要項目の除去 -del
                            // tmp_nkinkbn = "nkinkbnchk"
                            tmp_kozasyu = "kozasyuchk";
                            break;
                        }
                    case object _ when (tmp_komkname ?? "") != (tmp_komkname.Replace("入金項目情報", "") ?? ""):
                        {
                            tmp_nkinkomk = "nkinkomkchk";
                            break;
                        }
                    case object _ when (tmp_komkname ?? "") != (tmp_komkname.Replace("変動費", "") ?? ""):
                        {
                            tmp_nkinkomk = "nkinkomkchk";
                            break;
                        }
                    case object _ when tmp_komkname == "鍵タイトルマスタ":
                        {
                            tmp_kagi = "kagichk";
                            break;
                        }
                    case object _ when tmp_komkname == "部屋設備情報":
                        {
                            tmp_setubi = "setubichk";
                            break;
                        }

                    default:
                        {
                            switch (tmp_grpname ?? "")
                            {
                                case "物件情報":
                                    {
                                        tmp_bkrui = "bkruichk";
                                        tmp_kozo = "kozochk";
                                        tmp_kagi = "kagichk";
                                        tmp_bktosiyoto = "tosiyotochk";
                                        tmp_gazo = "gazochk";
                                        break;
                                    }
                                case "部屋情報":
                                    {
                                        tmp_hyrui = "hyruichk";
                                        tmp_kagi = "kagichk";
                                        tmp_kyrui = "kyruichk";
                                        tmp_nkinkbn = "nkinkbnchk";
                                        tmp_toritaiyo = "toritaiyochk";
                                        tmp_nkinkomk = "nkinkomkchk";
                                        tmp_sohoken = "sohokenchk";
                                        tmp_gazo = "gazochk";
                                        break;
                                    }
                                case "契約情報":
                                    {
                                        tmp_kagi = "kagichk";
                                        tmp_kyrui = "kyruichk";
                                        tmp_nkinkbn = "nkinkbnchk";
                                        tmp_nkinkomk = "nkinkomkchk";
                                        tmp_sohoken = "sohokenchk";
                                        break;
                                    }
                                case "送金ルール情報":
                                    {
                                        tmp_nkinkbn = "nkinkbnchk";
                                        tmp_nkinkomk = "nkinkomkchk";
                                        tmp_sohoken = "sohokenchk";
                                        break;
                                    }
                                case "請求情報":
                                    {
                                        tmp_nkinkbn = "nkinkbnchk";
                                        tmp_nkinkomk = "nkinkomkchk";
                                        break;
                                    }
                            }

                            break;
                        }
                }
                // 20161012 紐付ツール呼出起動時の項目過不足修正 -chg end
            }
            // 20160905 汎用紐付対応 汎用では表示しない紐付項目の修正_本体 -add sta
            // 汎用では紐付を行わない項目を除去
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
            {
                tmp_setubi = "";
                tmp_jisyakoza = "";
                tmp_gazo = "";
                tmp_fbfmt = "";
                tmp_bktosiyoto = "";  // 20161012 汎用時に紐付対象外となる項目非表示修正 -add
                tmp_kagi = "";        // 20161025 紐付不要項目の除去 -add
            }
            // 20160905 汎用紐付対応 汎用では表示しない紐付項目の修正_本体 -add end
            // コマンドライン用に半角スペースで結合
            rtn_str = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(tmp_bkrui, Interaction.IIf(string.IsNullOrEmpty(tmp_bkrui), "", "-")), tmp_hyrui), Interaction.IIf(string.IsNullOrEmpty(tmp_hyrui), "", "-")), tmp_nkinkbn), Interaction.IIf(string.IsNullOrEmpty(tmp_nkinkbn), "", "-")), tmp_toritaiyo), Interaction.IIf(string.IsNullOrEmpty(tmp_toritaiyo), "", "-")), tmp_nkinkomk), Interaction.IIf(string.IsNullOrEmpty(tmp_nkinkomk), "", "-")), tmp_kozo), Interaction.IIf(string.IsNullOrEmpty(tmp_kozo), "", "-")), tmp_kozasyu), Interaction.IIf(string.IsNullOrEmpty(tmp_kozasyu), "", "-")), tmp_setubi), Interaction.IIf(string.IsNullOrEmpty(tmp_setubi), "", "-")), tmp_kagi), Interaction.IIf(string.IsNullOrEmpty(tmp_kagi), "", "-")), tmp_jisyakoza), Interaction.IIf(string.IsNullOrEmpty(tmp_jisyakoza), "", "-")), tmp_bktosiyoto), Interaction.IIf(string.IsNullOrEmpty(tmp_bktosiyoto), "", "-")), tmp_kyrui), Interaction.IIf(string.IsNullOrEmpty(tmp_kyrui), "", "-")), tmp_gazo), Interaction.IIf(string.IsNullOrEmpty(tmp_gazo), "", "-")), tmp_fbfmt), Interaction.IIf(string.IsNullOrEmpty(tmp_fbfmt), "", "-")), tmp_sohoken));














            // 20160829 紐付ツール起動位置修正_本体→紐付 -add sta
            // ※強制でコンバートPG本体と同位置にセット
            if (!string.IsNullOrEmpty(rtn_str))
            {
                int tmp_strcnt = rtn_str.Length;
                if (rtn_str.Substring(tmp_strcnt - 1) == "-")
                {
                    rtn_str = rtn_str.Remove(tmp_strcnt - 1);
                }
            }
            // 20160829 紐付ツール起動位置修正_本体→紐付 -add end

            return rtn_str;

        }

        /// <summary>
        /// 項目選択1～3の全チェックボックス名を取得→リスト(オブジェクト)へ格納 '20160707 各抽出件数の出力処理追加
        /// </summary>
        /// <returns></returns>
        /// <remarks>
        /// </remarks>
        private List<string> Get_ListCVChkitemAll(bool chkonoffflg = false)
        {

            var rtn_list = new List<string>();
            string tmp_grpname = "";
            string tmp_chkname = "";

            tabPageManagerCvitem.ChangeTabPageVisible(0, true);
            tabPageManagerCvitem.ChangeTabPageVisible(1, true);
            tabPageManagerCvitem.ChangeTabPageVisible(2, true);

            // タブコントロール内
            foreach (Control item in tabCtrlCVItem.Controls)
            {
                if (item.GetType().Equals(typeof(TabPage)))
                {
                    TabPage container_tab = (TabPage)item;

                    // タブページ内
                    foreach (Control item_sub1 in container_tab.Controls)
                    {
                        if (item_sub1.GetType().Equals(typeof(GroupBox)))
                        {

                            GroupBox container_grp = (GroupBox)item_sub1;
                            // 移行グループ名取得
                            tmp_grpname = container_grp.Text;

                            // グループボックス内
                            foreach (Control item_sub2 in container_grp.Controls)
                            {
                                if (item_sub2.GetType().Equals(typeof(Panel)))
                                {

                                    Panel control_panel = (Panel)item_sub2;

                                    foreach (Control item_sub3 in control_panel.Controls)
                                    {

                                        if (item_sub3.GetType().Equals(typeof(CheckBox)))
                                        {

                                            CheckBox control_chk = (CheckBox)item_sub3;
                                            // 20161014 着色処理の修正 -chg sta
                                            // If CNVNO = ConvertTypes._既存ユーザ用 And list_baseCVchkitem.Contains(control_chk) = False Then

                                            // If chkonoffflg = False Then

                                            // '移行項目取得
                                            // tmp_chkname = control_chk.Text

                                            // 'リスト作成
                                            // rtn_list.Add(tmp_chkname.Trim)

                                            // ElseIf control_chk.Checked Then

                                            // '移行項目取得
                                            // tmp_chkname = control_chk.Text

                                            // 'リスト作成
                                            // rtn_list.Add(tmp_chkname.Trim)

                                            // End If

                                            // End If
                                            // コンバート実績あり項目の着色設定
                                            switch (CommonModule.CNVNO)
                                            {
                                                case (int)CommonModule.ConvertTypes._汎用:
                                                    {
                                                        if (list_existCVchkitem.Contains(control_chk) == false)
                                                        {
                                                            if (chkonoffflg == false)
                                                            {

                                                                // 移行項目取得
                                                                tmp_chkname = control_chk.Text;

                                                                // リスト作成
                                                                rtn_list.Add(tmp_chkname.Trim());
                                                            }

                                                            else if (control_chk.Checked)
                                                            {

                                                                // 移行項目取得
                                                                tmp_chkname = control_chk.Text;

                                                                // リスト作成
                                                                rtn_list.Add(tmp_chkname.Trim());

                                                            }
                                                        }

                                                        break;
                                                    }
                                            }
                                            // 20161014 着色処理の修正 -chg end
                                        }

                                    }

                                }
                            }

                        }
                    }

                }
            }

            // ※表示状態を戻す
            tabPageManagerCvitem.ChangeTabPageVisible(0, true);
            tabPageManagerCvitem.ChangeTabPageVisible(1, false);
            tabPageManagerCvitem.ChangeTabPageVisible(2, false);

            return rtn_list;

        }

        /// <summary>
        /// 項目選択1～3の全チェックボックスをキャプションとセットで取得→ハッシュテーブルへ(オブジェクト)へ格納 '20160707 コンバート実績保持の処理追加
        /// </summary>
        /// <returns></returns>
        /// <remarks>
        /// </remarks>
        private SafeDictionary<string, CheckBox> Get_Hash_CVChkitemAll(bool chkonoffflg = false)
        {

            var rtn_hash = new SafeDictionary<string, CheckBox>();
            string tmp_grpname = "";
            string tmp_chkname = "";

            tabPageManagerCvitem.ChangeTabPageVisible(0, true);
            tabPageManagerCvitem.ChangeTabPageVisible(1, true);
            tabPageManagerCvitem.ChangeTabPageVisible(2, true);

            // タブコントロール内
            foreach (Control item in tabCtrlCVItem.Controls)
            {
                if (item.GetType().Equals(typeof(TabPage)))
                {
                    TabPage container_tab = (TabPage)item;

                    // タブページ内
                    foreach (Control item_sub1 in container_tab.Controls)
                    {
                        if (item_sub1.GetType().Equals(typeof(GroupBox)))
                        {

                            GroupBox container_grp = (GroupBox)item_sub1;
                            // 移行グループ名取得
                            tmp_grpname = container_grp.Text;

                            // グループボックス内
                            foreach (Control item_sub2 in container_grp.Controls)
                            {
                                if (item_sub2.GetType().Equals(typeof(Panel)))
                                {

                                    Panel control_panel = (Panel)item_sub2;

                                    foreach (Control item_sub3 in control_panel.Controls)
                                    {
                                        if (item_sub3.GetType().Equals(typeof(CheckBox)))
                                        {

                                            CheckBox control_chk = (CheckBox)item_sub3;
                                            // 20161014 着色処理の修正 -chg sta
                                            // If CNVNO = ConvertTypes._既存ユーザ用 And list_baseCVchkitem.Contains(control_chk) = False Then

                                            // If chkonoffflg = False Then

                                            // '移行項目取得
                                            // tmp_chkname = control_chk.Text

                                            // 'ハッシュテーブル作成
                                            // If rtn_hash.Contains(tmp_chkname) = False Then
                                            // rtn_hash.Add(tmp_chkname, control_chk)
                                            // End If

                                            // ElseIf control_chk.Checked Then

                                            // '移行項目取得
                                            // tmp_chkname = control_chk.Text

                                            // 'ハッシュテーブル作成
                                            // If rtn_hash.Contains(tmp_chkname) = False Then
                                            // rtn_hash.Add(tmp_chkname, control_chk)
                                            // End If

                                            // End If

                                            // End If
                                            // コンバート実績あり項目の着色設定
                                            switch (CommonModule.CNVNO)
                                            {
                                                case (int)CommonModule.ConvertTypes._汎用:
                                                    {
                                                        if (list_existCVchkitem.Contains(control_chk) == false)
                                                        {
                                                            if (chkonoffflg == false)
                                                            {

                                                                // 移行項目取得
                                                                tmp_chkname = control_chk.Text;

                                                                // ハッシュテーブル作成
                                                                if (rtn_hash.ContainsKey(tmp_chkname) == false)
                                                                {
                                                                    rtn_hash.Add(tmp_chkname, control_chk);
                                                                }
                                                            }

                                                            else if (control_chk.Checked)
                                                            {

                                                                // 移行項目取得
                                                                tmp_chkname = control_chk.Text;

                                                                // ハッシュテーブル作成
                                                                if (rtn_hash.ContainsKey(tmp_chkname) == false)
                                                                {
                                                                    rtn_hash.Add(tmp_chkname, control_chk);
                                                                }

                                                            }
                                                        }

                                                        break;
                                                    }
                                            }
                                            // 20161014 着色処理の修正 -chg end
                                        }
                                    }

                                }
                            }

                        }
                    }

                }
            }

            // ※表示状態を戻す
            tabPageManagerCvitem.ChangeTabPageVisible(0, true);
            tabPageManagerCvitem.ChangeTabPageVisible(1, false);
            tabPageManagerCvitem.ChangeTabPageVisible(2, false);

            return rtn_hash;

        }

        #endregion

        #region 画面制御：処理状況画面処理

        /// <summary>
        /// 処理状況をテキストボックスへ出力(画面表示)<br/>
        /// </summary>
        private void Set_Situation(string value, int @type)
        {

            string tmp_str;

            switch (type)
            {
                case 0:  // 処理状況
                    {
                        tmp_str = txtTotalSituation.Text;
                        value = Conversions.ToString(DateTime.Now) + "  " + value + Constants.vbCrLf;
                        txtTotalSituation.Text = tmp_str + value;
                        break;
                    }
                case 1:  // 移行完了項目
                    {
                        tmp_str = txtPartialSituation.Text;
                        value = value + Constants.vbCrLf;
                        txtPartialSituation.Text = tmp_str + value;
                        break;
                    }
            }

            // スクロールを下に移動
            txtTotalSituation.SelectionStart = txtTotalSituation.Text.Length;
            txtTotalSituation.ScrollToCaret();

            txtPartialSituation.SelectionStart = txtPartialSituation.Text.Length;
            txtPartialSituation.ScrollToCaret();

        }

        /// <summary>
        /// 論理値を文字列へ変換
        /// </summary>
        /// <param name="flg">True.正常, False.異常</param>         
        /// <param name="typeno">1.個別進捗出力用, 2.全体進捗出力用</param>
        /// <returns></returns>
        /// <remarks></remarks>
        private string Chg_FlgToStr(bool flg, int typeno)
        {

            string rtn_str = "";

            switch (typeno)
            {

                case 1:
                    {
                        if (flg)
                        {
                            rtn_str = CommonModule.SITUATION_CVITEMSUCCESS;
                        }
                        else
                        {
                            rtn_str = CommonModule.SITUATION_CVITEMFAILURE;
                        }

                        break;
                    }
                case 2:
                    {
                        if (flg)
                        {
                            rtn_str = CommonModule.LOG_NAIYO_NORMALEND;
                        }
                        else
                        {
                            rtn_str = CommonModule.LOG_NAIYO_NOTNORMALEND;
                        }

                        break;
                    }
            }

            return rtn_str;

        }

        /// <summary>
        /// 移行項目の数を元にメッセージを表示　　　'20160912 既存用→汎用用中間ファイルへのコピー処理 開発用引数(existmidtobasemidflg)を追加
        /// </summary>
        /// <param name="list"></param>
        /// <param name="midchkbtnflg"></param>
        /// <returns>True.実行可, False.実行不可</returns>
        /// <remarks></remarks>
        private bool Chk_CVStartFlg(ref List<string> list, bool midchkbtnflg = false)
        {

            string tmp_msg = "";
            var tmp_list = new List<string>();

            // チェックON項目を格納したリストを取得 (グループ-チェックボックス名)
            tmp_list = Get_ListCVChkitem(midchkbtnflg);

            // 移行項目チェック
            if (tmp_list.Count == 0)
            {
                tmp_msg = tmp_msg + Constants.vbCrLf + CommonModule.MSG_ERR_CHK_CVITEM;
            }

            // 紐付ファイル格納先フォルダ有無確認
            string tmp_reldir = txtRelationDirPath.Text;
            if (EtcMethod.Chk_DirExist(tmp_reldir))
            {
                CommonModule.RelationDirPath = tmp_reldir;
            }
            else
            {
                tmp_msg = tmp_msg + Constants.vbCrLf + CommonModule.MSG_ERR_DIR_REL;
            }

            // ログファイル格納先フォルダ有無確認
            string tmp_logdir = txtLogDirPath.Text;
            if (EtcMethod.Chk_DirExist(tmp_logdir))
            {
                string logfilename = CommonModule.LOG_TMP_TABLENAME + "_" + Strings.Replace(Strings.Replace(Strings.Replace(string.Format(Conversions.ToString(DateTime.Now)), ":", ""), "/", ""), " ", "") + ".csv";
                CommonModule.LogFilePath = EtcMethod.Set_Path(tmp_logdir, logfilename);
            }
            else
            {
                tmp_msg = tmp_msg + Constants.vbCrLf + CommonModule.MSG_ERR_DIR_LOG;
            }

            // 中間ファイル格納先フォルダ有無確認
            string tmp_existmiddir = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_MID_NAME);    // 既存用中間ファイル格納先

            // 既存用中間ファイル格納先有無確認
            if (EtcMethod.Chk_DirExist(tmp_existmiddir))
            {
                CommonModule.MiddleDirPath = tmp_existmiddir;
            }
            else
            {
                tmp_msg = tmp_msg + Constants.vbCrLf + CommonModule.MSG_ERR_DIR_MID;
            }

            // 汎用用中間ファイル有無確認
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
            {
                if (EtcMethod.Chk_DirExist(tmp_existmiddir))
                {
                    CommonModule.BaseMidDirPath = Path.GetDirectoryName(tmp_existmiddir);
                    CommonModule.CV_FROM_MIDDLE = "preテーブル";
                }
                else
                {
                    tmp_msg = tmp_msg + Constants.vbCrLf + CommonModule.MSG_ERR_DIR_MID;
                }
            }

            // 中間ファイルログ格納先フォルダ有無確認
            string tmp_midlogdir = txtMidDirLogPath.Text;
            if (EtcMethod.Chk_DirExist(tmp_midlogdir))
            {
                string midlogfilename = CommonModule.LOG_TMP_MIDCHKTABLENAME + "_" + Strings.Replace(Strings.Replace(Strings.Replace(string.Format(Conversions.ToString(DateTime.Now)), ":", ""), "/", ""), " ", "") + ".csv";
                CommonModule.MiddleLogFilePath = EtcMethod.Set_Path(tmp_midlogdir, midlogfilename);
            }
            else
            {
                tmp_msg = tmp_msg + Constants.vbCrLf + CommonModule.MSG_ERR_DIR_MIDLOG;
            }


            // --------------------------------------------------
            // ファイルオープンチェック処理
            // --------------------------------------------------
            if (string.IsNullOrEmpty(tmp_msg))
            {
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del
                // Dim list_chkzumi As New List(Of String) 
                string openfilename = "";
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del sta
                // 既存用中間ファイルのオープン確認
                // For Each filesheetname In tmp_list
                // 'ファイル名、ファイルパスを取得
                // Dim tmp_str() As String = filesheetname.Split("-")
                // Dim filename As String = tmp_str(0) & ".xlsx"
                // Dim filepath As String = EtcMethod.Set_Path(MiddleDirPath, filename)

                // 'ファイルオープンチェック
                // If list_chkzumi.Contains(filepath) = False Then
                // Dim fileopenflg As Boolean = EtcMethod.Chk_FileOpen(filepath)
                // If fileopenflg = False Then
                // openfilename = openfilename & vbCrLf & INDENT_0 & "・" & filename
                // End If
                // list_chkzumi.Add(filepath)
                // End If
                // Next
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_メインフォーム -del end
                // 汎用用中間ファイルのオープン確認

                // 開かれているファイルが存在する場合
                if (!string.IsNullOrEmpty(openfilename))
                {
                    tmp_msg = tmp_msg + Constants.vbCrLf + CommonModule.MSG_ERR_FILEOPEN;
                    tmp_msg = tmp_msg + Constants.vbCrLf + "格納先：" + CommonModule.MiddleDirPath;
                    tmp_msg = tmp_msg + openfilename;
                }
            }

            // 20161018 フォルダ自動生成箇所の修正 -del sta
            // '--------------------------------------------------
            // ' フォルダ自動生成処理
            // '--------------------------------------------------
            // Dim tmp_middirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_MID_NAME)             'コンバーター実行ファイル格納フォルダ
            // Dim tmp_reldirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_RELEXEDIR_NAME)       '紐付ツール実行ファイル格納フォルダ
            // Dim tmp_gazodirpath As String = EtcMethod.Set_Path(dcv_exedir, DIR_GAZOEXEDIR_NAME)     '画像CVツール実行ファイル格納フォルダ
            // Dim list_makedir As New List(Of String)

            // 'コンバーターフォルダ同階層
            // list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_MAINLOG_NAME))
            // list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_INI_NAME))
            // list_makedir.Add(EtcMethod.Set_Path(dcv_exedir, DIR_LIST_NAME))
            // '中間ファイルフォルダ
            // If EtcMethod.Chk_DirExist(tmp_middirpath) Then
            // list_makedir.Add(EtcMethod.Set_Path(tmp_middirpath, DIR_TEMP_CSV))
            // list_makedir.Add(EtcMethod.Set_Path(tmp_middirpath, DIR_MIDLOG_NAME))
            // End If
            // '紐付ツールフォルダ
            // If EtcMethod.Chk_DirExist(tmp_reldirpath) Then
            // list_makedir.Add(EtcMethod.Set_Path(tmp_reldirpath, DIR_RELLOG_NAME))
            // End If
            // '画像CVツールフォルダ
            // If EtcMethod.Chk_DirExist(tmp_gazodirpath) Then
            // list_makedir.Add(EtcMethod.Set_Path(tmp_gazodirpath, DIR_INI_NAME))
            // End If

            // '自動生成処理 
            // For Each chkdir In list_makedir
            // If System.IO.Directory.Exists(chkdir) = False Then
            // System.IO.Directory.CreateDirectory(chkdir)
            // End If
            // Next
            // 20161018 フォルダ自動生成箇所の修正 -del end

            // --------------------------------------------------
            // 運用開始年月の追加処理
            // --------------------------------------------------
            string tmp_unyoymd = DateTime.Now.ToString("yyyy/MM");
            tmp_unyoymd = tmp_unyoymd.Replace("/", "").Trim();
            if (string.IsNullOrEmpty(tmp_unyoymd))
            {
                tmp_msg = tmp_msg + Constants.vbCrLf + CommonModule.MSG_ERR_UNYOYMD;
            }
            else
            {
                CommonModule.UnyoYMD = tmp_unyoymd + "01";
            }


            // --------------------------------------------------
            // 中間ファイルチェック or コンバート実行 ダイアログ出力
            // --------------------------------------------------
            // メッセージ表示
            if (!string.IsNullOrEmpty(tmp_msg))
            {
                tmp_msg = tmp_msg.Remove(0, 1);
                CommonModule.MsgResult = MessageBox.Show(tmp_msg, "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            // 実行確認
            string tmp_msgsta = "";

            if (midchkbtnflg)
            {
                tmp_msgsta = CommonModule.MSG_STA_C;
            }
            else
            {
                tmp_msgsta = CommonModule.MSG_STA_A;
                // 20161014 改善対応：コンバート処理前に中間ファイルチェックを促す -add sta
                if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
                {
                    tmp_msgsta = tmp_msgsta + Constants.vbCrLf + CommonModule.MSG_STA_D;
                }
                // 20161014 改善対応：コンバート処理前に中間ファイルチェックを促す -add end
            }
            CommonModule.MsgResult = MessageBox.Show(tmp_msgsta, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

            if (CommonModule.MsgResult == DialogResult.No)
            {
                Set_Forcus("next");
                return false;
            }

            // 実行時は移行項目を格納
            list = tmp_list;

            // 20161028 物件/部屋鍵取得方法修正 -add sta
            // 汎用CVの場合、鍵タイトルをプログラム内で自動実行するためにオブジェクトへ鍵タイトルを追加する
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
            {
                list.Add("各マスタ情報-鍵タイトルマスタ");
            }
            // 20161028 物件/部屋鍵取得方法修正 -add end

            return true;

        }

        #endregion

        #region 画面制御：チェックボックス処理・ボタン状態処理

        /// <summary>
        /// 条件(開発用、コンバートタイプ)によるチェックボックスの表示設定
        /// </summary>
        /// <remarks>
        /// ・条件はコマンドラインより取得
        /// </remarks>
        private void Set_ChkBoxVisible()
        {

            if (!CommonModule.Dev_CVFlg)
            {

                switch (CommonModule.CNVNO)
                {
                    case (int)CommonModule.ConvertTypes._汎用:
                        {
                            break;
                        }




                }

            }

        }

        /// <summary>
        /// 関連コントロール内の全チェックボックスON/OFF
        /// </summary>
        /// <param name="seltabname">コントロール名</param>
        /// <param name="setnotflg">True.[setbln]の逆値をセット, False.[setbln]をそのままセット [op]</param>
        /// <param name="setbln">True or False [op]</param>
        /// <remarks>
        /// 例：setnotflg=True, setbln=False の場合、全てTrueでセットする　
        /// </remarks>
        public void Set_ChkboxOnOff(string seltabname, bool setnotflg = false, bool setbln = false)

        {

            bool chk_checked;
            int chkflg = 0;
            chkflg = Conversions.ToInteger(Interaction.IIf(setnotflg, 0, 1));

            switch (seltabname ?? "")
            {
                case DC_SELTAB_K01:                       // 項目選択1
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkKiMstBus.Checked, setbln));
                        Set_ChkBoxValue(grpKizon1, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_K02:                       // 項目選択2
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkKiJisyaBase.Checked, setbln));
                        Set_ChkBoxValue(grpKizon5, chk_checked, chkflg);
                        Set_ChkBoxValue(grpKizon3, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_K03:                       // 項目選択3
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkKiBkBase.Checked, setbln));
                        Set_ChkBoxValue(grpKizon2, chk_checked, chkflg);
                        break;
                    }

                case DC_SELTAB_B01:                      // 各マスタ情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkMstBus.Checked, setbln));
                        Set_ChkBoxValue(grpMst, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B02:                      // 業者情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkGyCyukaiBase.Checked, setbln));
                        Set_ChkBoxValue(grpGy, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B03:                      // 各基本情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkJisyaBase.Checked, setbln));
                        Set_ChkBoxValue(grpJisya, chk_checked, chkflg);
                        Set_ChkBoxValue(grpOw, chk_checked, chkflg);
                        Set_ChkBoxValue(grpKys, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B04:                      // 物件情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkBkBase.Checked, setbln));
                        Set_ChkBoxValue(grpBk, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B05:                      // 部屋情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkHyBase.Checked, setbln));
                        Set_ChkBoxValue(grpHy, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B06:                      // 送金ルール
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkSoruleBase.Checked, setbln));
                        Set_ChkBoxValue(grpSorule, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B07:                      // 契約情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkKyBase.Checked, setbln));
                        Set_ChkBoxValue(grpKy, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B08:                      // 未収金・過剰金情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkSqKajyo.Checked, setbln));
                        Set_ChkBoxValue(grpSq, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B09:                      // クレーム修繕情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkClaimBase.Checked, setbln));
                        Set_ChkBoxValue(grpClaim, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_B10:                      // 初期設定情報
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkSyskanriBase.Checked, setbln));
                        Set_ChkBoxValue(grpSyskanri, chk_checked, chkflg);
                        break;
                    }
                case var @case when @case == DC_SELTAB_B10:                      // 物件データ連動情報  '20160720 連動情報構築
                    {
                        chk_checked = Conversions.ToBoolean(Interaction.IIf(setnotflg, !chkRendoSosinBase.Checked, setbln));
                        Set_ChkBoxValue(grpRendo, chk_checked, chkflg);
                        break;
                    }
                case DC_SELTAB_H01:                      // 没選択1(汎用)
                    {
                        break;
                    }
                // chk_checked = IIf(setnotflg, Not Me.chkKiMstBus.Checked, setbln)
                // Call Me.Set_ChkBoxValue(Me.grpKizon1, chk_checked, chkflg)
                case var case1 when case1 == DC_SELTAB_H01:                      // 没選択2(汎用)
                    {
                        break;
                    }
                // chk_checked = IIf(setnotflg, Not Me.chkKiMstBus.Checked, setbln)
                // Call Me.Set_ChkBoxValue(Me.grpKizon1, chk_checked, chkflg)
                case DC_SELTAB_H03:                      // 没選択3(汎用)
                    {
                        break;
                    }
                // chk_checked = IIf(setnotflg, Not Me.chkKiMstBus.Checked, setbln)
                // Call Me.Set_ChkBoxValue(Me.grpKizon1, chk_checked, chkflg)

                default:
                    {
                        break;
                    }

            }

        }

        /// <summary>
        /// 鍵選択制御(ラジオボタン選択値をチェックボックスへ反映)
        /// </summary>
        /// <remarks></remarks>
        public void Set_OptSelect()
        {

            optHyKagi.Checked = optKiHyKagi.Checked;
            optKyKagi.Checked = optKiKyKagi.Checked;
            // 20160531 鍵情報移行処理の修正 -del sta
            // chkHyKagi.Enabled = optKiHyKagi.Checked
            // chkKyKagi.Enabled = Not optKiHyKagi.Checked
            // 20160531 鍵情報移行処理の修正 -del end

        }

        /// <summary>
        /// コンテナへの格納順序指定
        /// </summary>
        /// <param name="container"></param>
        /// <remarks></remarks>
        private void Set_ControlIndex(Control container)
        {

            var controls = new List<Control>();

            switch (container.Name ?? "")
            {

                // 契約者情報より先に自社情報がコンバートされるように順序を指定する
                case DC_SELTAB_B03:
                    {
                        // 20161026 自社情報と家主情報のコンバート順序入れ替え -chg sta
                        // controls.Add(Me.grpJisya)
                        // controls.Add(Me.grpOw)
                        switch (CommonModule.CNVNO)
                        {
                            case (int)CommonModule.ConvertTypes._汎用:
                                {
                                    controls.Add(grpOw);
                                    controls.Add(grpJisya);
                                    break;
                                }
                        }
                        // 20161026 自社情報と家主情報のコンバート順序入れ替え -chg end
                        controls.Add(grpKys);
                        break;
                    }
                case DC_SELTAB_B09:
                    {
                        controls.Add(grpClaim);
                        controls.Add(grpSzen);
                        break;
                    }

                case "grpMst":
                    {
                        controls.Add(chkMstBus);
                        controls.Add(chkMstBusKotu);
                        controls.Add(chkMstSchool);
                        controls.Add(chkMstArea);
                        controls.Add(chkMstKagititle);
                        controls.Add(chkMstHokenrui);
                        // controls.Add(Me.chkMstKeiyakurui)
                        controls.Add(chkMstTokuyaku);
                        controls.Add(chkMstKasyoClaimrui);
                        controls.Add(chkMstHendo);
                        controls.Add(chkMstHendoitiran);
                        controls.Add(chkMstBikotitle);
                        controls.Add(chkMstBikolst);
                        controls.Add(chkMstGazotitle);                // 20160801 タイトルマスタ統合処理 -add
                        break;
                    }
                case "grpGy":
                    {
                        controls.Add(chkGyCyukaiBase);
                        controls.Add(chkGyCyukaiKoza);
                        controls.Add(chkGyCyukaiMemo);
                        controls.Add(chkGyHokenBase);
                        controls.Add(chkGyHokenKoza);
                        controls.Add(chkGyHokenMemo);
                        controls.Add(chkGyYatinhosyoBase);
                        // controls.Add(Me.chkGyYatinhosyoKoza)
                        controls.Add(chkGyYatinhosyoMemo);
                        controls.Add(chkGySyuzenBase);
                        controls.Add(chkGySyuzenKoza);
                        controls.Add(chkGySyuzenMemo);
                        controls.Add(chkGyLifelineBase);
                        controls.Add(chkGySekoBase);
                        controls.Add(chkGySisetuBase);
                        break;
                    }
                case "grpJisya":
                    {
                        controls.Add(chkJisyaBase);
                        controls.Add(chkJisyaKoza);
                        controls.Add(chkJisyaTanto);
                        controls.Add(chkJisyaMemo);
                        controls.Add(chkFBFuriirai);
                        controls.Add(chkFBFuritesuryo);
                        controls.Add(chkFBKozafurikae);
                        controls.Add(chkFBNsSyutoku);
                        controls.Add(chkMstYatinKoza);
                        controls.Add(chkMstANSERArea);
                        controls.Add(chkMstANSERAccpoint);
                        controls.Add(chkFBANSERSetuzoku);
                        break;
                    }
                case "grpOw":
                    {
                        controls.Add(chkOwBase);
                        controls.Add(chkOwKoza);
                        controls.Add(chkOwEvent);
                        controls.Add(chkOwMemo);
                        break;
                    }
                case "grpKys":
                    {
                        controls.Add(chkKysBase);
                        controls.Add(chkKysKoza);
                        controls.Add(chkKysMemo);
                        controls.Add(chkKysSyogoKana);
                        controls.Add(chkKysHosyonin);
                        break;
                    }
                case "grpBk":
                    {
                        controls.Add(chkBkBase);
                        controls.Add(chkBkSyosai);
                        controls.Add(chkBkSyo);
                        controls.Add(chkBkGomi);
                        controls.Add(chkBkKenri);
                        controls.Add(chkBkKotu);
                        controls.Add(chkBkSetudo);
                        controls.Add(chkBkSyuhen);
                        controls.Add(chkBkSzeniji);
                        controls.Add(chkBkMemo);
                        controls.Add(chkBkKagi);                      // 汎用分(V7には無い)
                        controls.Add(chkBkHendo);                     // 汎用分(V7には無い)
                        controls.Add(chkBkKinrincyusyajo);            // 汎用分(V7には無い)
                        controls.Add(chkBkSansyofile);                // 汎用分(V7には無い)
                        break;
                    }
                case "grpHy":
                    {
                        controls.Add(chkHyBase);
                        controls.Add(chkHySyosai);
                        controls.Add(chkHySyo);
                        controls.Add(chkHyParking);
                        controls.Add(chkHyTokuyaku);
                        controls.Add(chkHyKagi);
                        controls.Add(chkHyMadoriutiwake);
                        controls.Add(chkHyMenseki);
                        controls.Add(chkHySetubi);
                        controls.Add(chkHyNkinkomk);
                        controls.Add(chkHyHendo);
                        controls.Add(chkHyMemo);
                        controls.Add(chkHyCommonsalespoint);
                        controls.Add(chkHyConfirm);
                        controls.Add(chkHyKenri);
                        controls.Add(chkHySansyofile);
                        controls.Add(chkHyGenjotanka);
                        controls.Add(chkHySzeniji);
                        break;
                    }
                case "grpSorule":
                    {
                        controls.Add(chkSoruleBase);
                        controls.Add(chkSoruleSosaki);
                        controls.Add(chkSoruleNkin);
                        controls.Add(chkSoruleKojo);
                        break;
                    }
                case "grpKy":
                    {
                        controls.Add(chkKyBase);
                        controls.Add(chkKyRireki);
                        controls.Add(chkKyCar);
                        controls.Add(chkKyKys);
                        controls.Add(chkKyHosyonin);
                        controls.Add(chkKyNyukyo);
                        controls.Add(chkKyTokuyaku);
                        controls.Add(chkKyHoken);
                        controls.Add(chkKyMemo);
                        controls.Add(chkKyNkinkomk);
                        controls.Add(chkKyNkinkomkNx);
                        controls.Add(chkKyHendo);
                        controls.Add(chkKyKojoRule);
                        controls.Add(chkKySorule);
                        controls.Add(chkKyKai);
                        controls.Add(chkKySzen);
                        controls.Add(chkKySzenmeisai);
                        break;
                    }
                case "grpSq":
                    {
                        controls.Add(chkSqKajyo);
                        controls.Add(chkSqUnyotaino);
                        controls.Add(chkSqSq);
                        controls.Add(chkSqHendokensin);
                        controls.Add(chkSqKoteiKojo);
                        controls.Add(chkSqSqKojo);
                        break;
                    }
                case "grpClaim":
                    {
                        controls.Add(chkClaimBase);
                        controls.Add(chkClaimTaiorireki);
                        controls.Add(chkClaimRelfile);
                        break;
                    }
                case "grpSzen":
                    {
                        controls.Add(chkSzenBase);
                        controls.Add(chkSzenSzen);
                        controls.Add(chkSzenSzenmeisai);
                        controls.Add(chkSzenClaim);
                        controls.Add(chkSzenRelfile);               // 20160627 修繕関連ファイル移行修正 -add
                        controls.Add(chkSzenMemo);
                        break;
                    }
                case "grpSyskanri":
                    {
                        controls.Add(chkSyskanriBase);
                        controls.Add(chkSyskanriZei);
                        controls.Add(chkSyskanriHenkanmoji);
                        controls.Add(chkSyskanriNkinkomkmerge);
                        break;
                    }
                case "grpRendo": 									// 20160720 連動情報構築 -add
                    {
                        controls.Add(chkRendoSosinBase);
                        controls.Add(chkRendoSosinJisyaweb);
                        controls.Add(chkRendoSosinHomes);
                        controls.Add(chkRendoSosinAthome);
                        controls.Add(chkRendoSosinSuumo);
                        controls.Add(chkRendoMapdisp);
                        controls.Add(chkRendoBtoBgroup);
                        controls.Add(chkRendoHysosin);
                        controls.Add(chkRendoHyrui);
                        controls.Add(chkRendoKokokuSuumo);
                        controls.Add(chkRendoKokokuAthome);
                        controls.Add(chkRendoKokokuHomes);
                        controls.Add(chkRendoKokokuJisyaweb);
                        break;
                    }
            }

            for (int i = 0, loopTo = controls.Count - 1; i <= loopTo; i++)
                container.Controls.SetChildIndex(controls[i], i);

        }

        /// <summary>
        /// チェックボックスのON/OFF一括設定(コンテナ毎)
        /// </summary>
        /// <param name="container">コンテナ名</param>
        /// <param name="value">設定値</param>
        /// <param name="chkflg">強制チェックONOFFフラグ…1.強制, 1以外.通常</param>
        /// <remarks></remarks>
        private void Set_ChkBoxValue(Control container, bool value, int chkflg = 0)
        {

            var list_tab = new List<string>() { "grpKizon1", "grpKizon2", "grpKizon3", "grpKizon5" };

            if (list_tab.Contains(container.Name))
            {
                foreach (Control item in container.Controls)
                {
                    if (item.GetType().Equals(typeof(Panel)))
                    {
                        Panel tmp_panel = (Panel)item;
                        foreach (Control item_inpanel in tmp_panel.Controls)
                        {
                            if (item_inpanel.GetType().Equals(typeof(CheckBox)))
                            {
                                CheckBox tmp_chkbox = (CheckBox)item_inpanel;
                                if (chkflg == 1 || tmp_chkbox.Visible)
                                {
                                    if (tmp_chkbox.Enabled)
                                    {
                                        tmp_chkbox.Checked = value;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (Control item in container.Controls)
                {
                    if (item.GetType().Equals(typeof(CheckBox)))
                    {
                        CheckBox tmp_chkbox = (CheckBox)item;
                        if (chkflg == 1 || tmp_chkbox.Visible)
                        {
                            if (tmp_chkbox.Enabled)
                            {
                                tmp_chkbox.Checked = value;
                            }
                        }
                    }
                }
            }

        }

        /// <summary>
        /// チェックボックスの全チェック(事前作業画面)
        /// </summary>
        /// <param name="ctrl">チェック対象コントロール</param>
        /// <returns>True.全チェック, False.チェックOFFあり</returns>
        /// <remarks>
        /// ・"Get_ListCVChkitem"メソッドを参考に作成
        /// ・Me.tabCtrlJizen.Controls用
        /// ・仕様上、全タブを表示する
        /// </remarks>
        private bool Chk_JizenChkitem(object ctrl)
        {

            bool rtn = true;
            string tmp_grpname = "";
            string tmp_chkname = "";


            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        tabPageManagerJizen.ChangeTabPageVisible(0, true);
                        break;
                    }
            }


            // タブコントロール内
            foreach (Control item in (IEnumerable)ctrl)
            {
                if (item.GetType().Equals(typeof(TabPage)))
                {
                    TabPage container_tab = (TabPage)item;

                    // タブページ内
                    foreach (Control item_sub1 in container_tab.Controls)
                    {
                        if (item_sub1.GetType().Equals(typeof(GroupBox)))
                        {
                            GroupBox container_grp = (GroupBox)item_sub1;

                            // 移行グループ名取得
                            tmp_grpname = container_grp.Text;

                            // グループボックス内
                            foreach (Control item_sub2 in container_grp.Controls)
                            {
                                if (item_sub2.GetType().Equals(typeof(Panel)))
                                {
                                    Panel container_panel = (Panel)item_sub2;
                                    foreach (Control item_sub3 in container_panel.Controls)
                                    {
                                        if (item_sub3.GetType().Equals(typeof(CheckBox)))
                                        {
                                            CheckBox control_chk = (CheckBox)item_sub3;
                                            if (control_chk.Enabled & control_chk.Checked == false)
                                            {
                                                return false;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return rtn;

        }

        /// <summary>
        /// 子項目チェックボックスの活性制御 
        /// </summary>
        /// <param name="flg"></param>
        /// <remarks></remarks>
        private void Set_ChkBoxCondition(bool flg)
        {

            if (true)
            {

                // 各マスタ情報
                chkMstBusKotu.Enabled = flg;
                chkMstHendoitiran.Enabled = flg;
                chkMstBikolst.Enabled = flg;
                // 業者情報
                chkGyCyukaiKoza.Enabled = flg;
                chkGyCyukaiMemo.Enabled = flg;
                chkGyHokenKoza.Enabled = flg;
                chkGyHokenMemo.Enabled = flg;
                chkGyYatinhosyoKoza.Enabled = flg;
                chkGyYatinhosyoMemo.Enabled = flg;
                chkGySyuzenKoza.Enabled = flg;
                chkGySyuzenMemo.Enabled = flg;
                // 自社情報
                chkJisyaKoza.Enabled = flg;
                chkJisyaTanto.Enabled = flg;
                chkJisyaMemo.Enabled = flg;
                chkFBFuriirai.Enabled = flg;
                chkFBKozafurikae.Enabled = flg;
                // 家主情報
                chkOwKoza.Enabled = flg;
                chkOwEvent.Enabled = flg;
                chkOwMemo.Enabled = flg;
                // 契約者情報
                chkKysKoza.Enabled = flg;
                chkKysMemo.Enabled = flg;
                chkKysSyogoKana.Enabled = flg;
                chkKysHosyonin.Enabled = flg;
                // 物件情報
                chkBkSyosai.Enabled = flg;
                chkBkSyo.Enabled = flg;
                chkBkGomi.Enabled = flg;
                chkBkKenri.Enabled = flg;
                chkBkKotu.Enabled = flg;
                chkBkSetudo.Enabled = flg;
                chkBkSyuhen.Enabled = flg;
                chkBkSzeniji.Enabled = flg;
                chkBkMemo.Enabled = flg;
                chkBkKagi.Enabled = flg;
                chkBkHendo.Enabled = flg;
                chkBkKinrincyusyajo.Enabled = flg;
                chkBkSansyofile.Enabled = flg;
                // 部屋情報
                chkHySyosai.Enabled = flg;
                chkHySyo.Enabled = flg;
                chkHyParking.Enabled = flg;
                chkHyTokuyaku.Enabled = flg;
                chkHyKagi.Enabled = flg;
                chkHyMadoriutiwake.Enabled = flg;
                chkHyMenseki.Enabled = flg;
                chkHySetubi.Enabled = flg;
                chkHyNkinkomk.Enabled = flg;
                chkHyHendo.Enabled = flg;
                chkHyMemo.Enabled = flg;
                chkHyCommonsalespoint.Enabled = flg;
                chkHyConfirm.Enabled = flg;
                chkHyKenri.Enabled = flg;
                chkHySansyofile.Enabled = flg;
                chkHyGenjotanka.Enabled = flg;
                chkHySzeniji.Enabled = flg;
                // 送金ルール情報
                chkSoruleSosaki.Enabled = flg;
                chkSoruleNkin.Enabled = flg;
                chkSoruleKojo.Enabled = flg;
                // 契約情報
                chkKyRireki.Enabled = flg;
                chkKyCar.Enabled = flg;
                chkKyKys.Enabled = flg;
                chkKyHosyonin.Enabled = flg;
                chkKyNyukyo.Enabled = flg;
                chkKyTokuyaku.Enabled = flg;
                chkKyHoken.Enabled = flg;
                chkKyMemo.Enabled = flg;
                chkKyNkinkomk.Enabled = flg;
                chkKyNkinkomkNx.Enabled = flg;
                chkKyHendo.Enabled = flg;
                chkKyKojoRule.Enabled = flg;
                chkKySorule.Enabled = flg;
                chkKyKai.Enabled = flg;
                // 20160531 鍵情報移行処理の修正 -del sta
                // '2016.04.06 契約鍵情報の移行処理追加 -add
                // Me.chkKyKagi.Enabled = flg
                // 20160531 鍵情報移行処理の修正 -del end

            }

        }

        /// <summary>
        /// 既存コンバーターのみ表示するチェックボックスのチェックをOFFにする処理 '20160905 汎用の場合の移行対象外項目(既存のみ移行対象)選定の修正 -add
        /// </summary>
        /// <remarks>
        /// ・(汎用)
        /// </remarks>
        private void Set_ExistCVItemToList()
        {

            // 各基本情報_自社                                        '20160915 汎用移行対象からANSERを除外 -add
            list_existCVchkitem.Add(chkFBNsSyutoku);
            list_existCVchkitem.Add(chkMstANSERAccpoint);
            list_existCVchkitem.Add(chkMstANSERArea);
            list_existCVchkitem.Add(chkFBANSERSetuzoku);
            list_existCVchkitem.Add(chkFBNsSyutoku);          	// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkFBFuritesuryo);        	// 20160926 選定した移行項目をプログラムへ反映する修正 -add

            // 各基本情報_家主   										'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkOwEvent);

            // 各マスタ情報      										'20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkMstHendo);
            list_existCVchkitem.Add(chkMstHendoitiran);
            list_existCVchkitem.Add(chkMstKagititle);

            // 物件情報
            list_existCVchkitem.Add(chkBkHendo);
            // list_existCVchkitem.Add(Me.chkBkKinrincyusyajo)        '20160926 選定した移行項目をプログラムへ反映する修正 -del
            list_existCVchkitem.Add(chkOpOw);
            list_existCVchkitem.Add(chkOpKys);
            list_existCVchkitem.Add(chkBkGomi);       			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkBkKenri);      			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkBkSetudo);     			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkBkSyuhen);     			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkBkSzeniji);    			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkBkSansyofile); 			// 20160926 選定した移行項目をプログラムへ反映する修正 -add

            // 部屋情報
            list_existCVchkitem.Add(chkHyMenseki);
            list_existCVchkitem.Add(chkHyKenri);
            list_existCVchkitem.Add(chkHyConfirm);    			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkHySzeniji);    			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkHySansyofile); 			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkHyGenjotanka); 			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkHyHendo);      			// 20160926 選定した移行項目をプログラムへ反映する修正 -add

            // 契約情報
            list_existCVchkitem.Add(chkKySorule);
            list_existCVchkitem.Add(chkKyKojoRule);
            list_existCVchkitem.Add(chkKyKai);
            list_existCVchkitem.Add(chkKySzen);
            list_existCVchkitem.Add(chkKySzenmeisai);
            list_existCVchkitem.Add(CheckBox14);                  // 契約変動費親メーター情報
            list_existCVchkitem.Add(CheckBox24);                  // 契約解約確認事項情報
            list_existCVchkitem.Add(CheckBox23);                  // 契約同時契約情報
            list_existCVchkitem.Add(CheckBox18);                  // 契約原状回復目安単価情報
            list_existCVchkitem.Add(CheckBox16);                  // 契約敷金保証金随時処理情報
            list_existCVchkitem.Add(CheckBox15);                  // 契約関連ファイル情報
            list_existCVchkitem.Add(CheckBox5);                   // 契約空室待ち情報
            list_existCVchkitem.Add(chkKyHendo);					// 20160926 選定した移行項目をプログラムへ反映する修正 -add

            // 請求情報
            list_existCVchkitem.Add(chkSqSq);
            list_existCVchkitem.Add(chkSqHendokensin);
            list_existCVchkitem.Add(chkSqSqKojo);
            list_existCVchkitem.Add(chkSqKajyo);      			// 20160926 選定した移行項目をプログラムへ反映する修正 -add
            list_existCVchkitem.Add(chkSqUnyotaino);  			// 20160926 選定した移行項目をプログラムへ反映する修正 -add

        }

        #endregion

        #region 紐付設定画面関連処理

        /// <summary>
        /// 紐付項目書込処理
        /// </summary>
        /// <returns></returns>
        /// <remarks>
        /// ・紐付項目チェックの機能追加時に移行対象項目を設定する
        /// ・現時点では全項目を移行するようにしておく
        /// </remarks>
        private bool Set_RelItem()
        {

            bool rtn = true;

            // 紐付項目取得
            var list_relitem = new List<string>() { "物件分類マスタ", "部屋分類マスタ", "入金区分マスタ", "契約分類マスタ", "入金項目マスタ", "設備マスタ" };        // 20160829 設備の新規挿入処理を追加 "設備マスタ"を追加

            foreach (var relitem in list_relitem)
            {

                // 正常終了フラグ
                bool normalflg = true;

                // 処理Repositoryの生成
                object obj_rep = null;
                obj_rep = Get_ObjRep_RelItem(relitem);

                if (obj_rep is not null)
                {
                    normalflg = Conversions.ToBoolean(((dynamic)obj_rep).Set_RelItem(sqlcnnv10, relitem));
                }

                if (normalflg == false)
                {
                    rtn = false;
                    return rtn;
                }

            }

            return rtn;

        }

        /// <summary>
        /// 紐付項目書込Repositoryの生成 
        /// </summary>
        /// <param name="relitem"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private object Get_ObjRep_RelItem(string relitem)
        {

            object obj_rep = null;

            switch (relitem ?? "")
            {
                case "物件分類マスタ":
                    {
                        obj_rep = new Repository.M_bk_rui_Repository();
                        break;
                    }
                case "部屋分類マスタ":
                    {
                        obj_rep = new Repository.M_hy_rui_Repository();
                        break;
                    }
                case "入金区分マスタ":
                    {
                        obj_rep = new Repository.M_nkbn_Repository();
                        break;
                    }
                case "契約分類マスタ":
                    {
                        obj_rep = new Repository.M_ky_rui_Repository();
                        break;
                    }
                case "入金項目マスタ":
                    {
                        obj_rep = new Repository.M_nkin_Repository();
                        break;
                    }
                case "設備マスタ":
                    {
                        obj_rep = new Repository.M_setubi_rel_Repository();    // 20160829 設備の新規挿入処理を追加 -add
                        break;
                    }

            }

            return obj_rep;

        }

        /// <summary>
        /// 紐付画面呼出処理　'20160707 本体と紐付ツールの中断を同期させる処理の追加 Sub → Function へ変更
        /// </summary>
        /// <remarks></remarks>
        private bool Rel_ExeCall(string cvitemtorel)
        {

            int rowcnt = 0;
            bool rtn_cancelflg = false;        									// 20160707 本体と紐付ツールの中断を同期させる処理の追加 -add 中断フラグ返却用

            // ログ/状況出力
            string tmp_sql_relsta = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(1, CommonModule.LOG_SYORIKOMK_RELSTA), false);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_relsta, ref rowcnt);
            Set_Situation(CommonModule.SITUATION_RELSTA, 0);

            // 紐付ツールパス取得(Exe実行用)
            string exepath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string exedir = Path.GetDirectoryName(exepath);
            string tmp_strpath = EtcMethod.Set_Path(exedir, CommonModule.DIR_RELEXEDIR_NAME);
            string relexepath = EtcMethod.Set_Path(tmp_strpath, CommonModule.DIR_RELEXE_NAME);


            // --------------------------------------------------
            // 紐付設定画面呼出処理
            // --------------------------------------------------
            string tmpstre = "0" + " " + CommonModule.CNVNO + " ";

            // V7接続情報
            string coninfov7 = "True,,,,";




            // 10接続情報
            string coninfov10 = optV10Authent1.Checked + "," + fstmodelv10.ServerName.Trim() + "," + fstmodelv10.InitialCatalog.Trim() + "," + fstmodelv10.User.Trim() + "," + fstmodelv10.Pass.Trim() + "," + fstmodelv10.TimeOut;





            // 紐付設定画面PG強制表示位置セット
            string tmp_relloc = Left.ToString() + "," + Top.ToString();

            // 紐付設定画面へ渡す連結パラメータ (接続情報 + 紐付けファイル格納先 + ログファイル格納先)
            string cmd_fk8db_cnv = "\"" + coninfov7 + "\"" + " " + "\"" + coninfov10 + "\"" + " " + "\"" + CommonModule.RelationDirPath + "\"" + " " + "\"" + CommonModule.MiddleDirPath + "\"" + " " + cvitemtorel;




            cmd_fk8db_cnv = tmpstre + cmd_fk8db_cnv;
            cmd_fk8db_cnv = cmd_fk8db_cnv + " " + tmp_relloc;                            	// 20160829 紐付ツール起動位置修正_本体→紐付 -add
            cmd_fk8db_cnv = cmd_fk8db_cnv + " " + "\"" + CommonModule.BaseMidDirPath + "\"";          	// 20161004 自社口座の口座種別取得処理の修正 -add

            // ◎紐付設定画面起動(連結パラメータ渡し)
            var obj_ExeStart = Process.Start(relexepath, cmd_fk8db_cnv);
            obj_ExeStart.WaitForExit();

            // 20160707 本体と紐付ツールの中断を同期させる処理の追加 -chg sta
            // 'ログ/状況出力
            // Dim tmp_sql_relend As String = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, LOG_SYORIKOMK_RELEND), False)
            // DBExec.Exec_NonQuery(Me.sqlcnnv10, tmp_sql_relend, rowcnt)
            // Call Me.Set_Situation(SITUATION_RELEND, 0)

            // 10本体アクティブ化
            Activate();                                                                   // 20160829_紐付完了後の本体をアクティブにする修正 -add

            // ----- 全体用・個別用プログレスバー更新 -----
            var obj_pgb = new ProgressBarManager();
            int pgbcnt = 1;
            obj_pgb.pgbsettingPart(pgbcnt);
            obj_com.ProgressOutPut(pgbcnt, pgbcnt);
            obj_pgb.pgbsettingTotal(pgbcnt);
            obj_com.ProgressOutPut(pgbcnt, pgbcnt, true);
            tabPageJikko.Refresh();

            // 紐付設定PGの終了戻り値(紐付設定画面で固定値を設定)を取得 (0.正常終了, 1.中断/キャンセル終了, 99.異常終了)        '20161108 レビュー結果：「紐付設定画面にて固定値設定」のような内容の記載をして下さい 20161108_2 レビュー結果戻り修正 コメントを追記しました。
            int exitcode = obj_ExeStart.ExitCode;
            string logstr = "";
            switch (exitcode)
            {
                case 0:                                                                              // 20161108 レビュー結果：エラーの場合もキャンセルフラグを立てている？そうであれば、他のキャンセルフラグの箇所は問題ない？
                    {
                        // → 20161108_2 レビュー結果戻り修正
                        // 紐付ツール移行の処理をスキップするためにエラーの場合でもキャンセルフラグを立てました。
                        // エラー終了かキャンセル終了かは戻り値「relexeerr」で判別するように修正したので問題ないと思います。
                        // ※エラー終了とキャンセル終了で検証しました。
                        logstr = CommonModule.LOG_SYORIKOMK_RELEND;
                        rtn_cancelflg = false;
                        // 20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add
                        relexeerr = exitcode;
                        break;
                    }
                case 1:
                    {
                        logstr = CommonModule.LOG_SYORIKOMK_RELCANCEL;
                        rtn_cancelflg = true;
                        // 20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add
                        relexeerr = exitcode;
                        break;
                    }
                // 20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add sta
                case 99:
                    {
                        logstr = CommonModule.LOG_SYORIKOMK_RELCANCEL;
                        rtn_cancelflg = true;
                        relexeerr = exitcode;
                        break;
                    }
                    // 20161104 紐付設定値保存用ファイルの読込エラー時の対応 -add end
            }

            // ログ/状況出力
            string tmp_sql_relend = LogSetting.Get_LogTblInsertQry(LogSetting.Set_LogValue(9, logstr), false);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_relend, ref rowcnt);
            Set_Situation(CommonModule.SITUATION_RELEND, 0);

            // 返却
            return rtn_cancelflg;
            // 20160707 本体と紐付ツールの中断を同期させる処理の追加 -chg end

        }

        #endregion

        #region ログ出力処理

        /// <summary>
        /// CSVファイル出力時にヘッダを挿入する       
        /// </summary>
        /// <remarks></remarks>
        private void Set_LogHeader(string filepath, string headerstr)
        {

            var strtype = System.Text.Encoding.GetEncoding("shift_jis");     // 文字コード設定
            var strread = new StreamReader(filepath, strtype);                                      // ログファイルを開く
            string tmp_path = Path.GetTempFileName();                                         // 仮ファイル作成
            var strwrite = new StreamWriter(tmp_path, false, strtype);                              // 仮ファイルを開く

            // 仮ファイルへヘッダー書込み
            strwrite.WriteLine(headerstr);

            // 内容読込
            while (strread.Peek() > -1)
            {
                string line = strread.ReadLine();
                strwrite.WriteLine(line);
            }

            // 終了処理
            strread.Close();
            strwrite.Close();

            // 仮ファイルと入れ替え
            File.Copy(tmp_path, filepath, true);
            File.Delete(tmp_path);

        }

        /// <summary>
        /// ログ出力初期設定 
        /// </summary>
        /// <remarks></remarks>
        private void Set_LogInit(bool midchkflg)
        {

            int rowcnt = 0;
            string taisyoname = "";

            if (midchkflg)
            {
                taisyoname = CommonModule.LOG_TMP_MIDCHKTABLENAME;
            }
            else
            {
                taisyoname = CommonModule.LOG_TMP_TABLENAME;
            }

            // ログ格納用テーブル初期化(DROP)
            string tmp_sql_drop = DBQuery.Qry_DropInfo(taisyoname, true);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref rowcnt);

            // ログ格納用テーブル生成
            string tmp_sql_create = LogSetting.Get_LogTblCreateQry(midchkflg);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref rowcnt);

        }

        /// <summary>
        /// ログ出力設定 
        /// </summary>
        /// <param name="logdropflg"></param>
        /// <remarks></remarks>
        private void Set_LogFile(bool logdropflg)
        {

            int rowcnt = 0;

            // ログファイル出力
            if (!EtcMethod.Chk_FileExist(CommonModule.LogFilePath))
            {
                // ログファイル(CSV)のデータ部出力
                string tmp_sql = LogSetting.Get_LogTblSelectQry(false);
                bool rtn = true;
                string errstr = "";
                rtn = FileMethod.TblView_Output_CSV(sqlcnnv10, CommonModule.LogFilePath, "", "", ref errstr, tmp_sql, true);

                // ヘッダーを加えて加工
                // 20161104 ログ出力内容修正(開発用以外はテーブル名を表示しない) -chg sta
                // Call Me.Set_LogHeader(LogFilePath, LOG_HEADER_TOTAL)
                string tmp_header = CommonModule.LOG_HEADER_TOTAL;
                if (CommonModule.Dev_CVFlg == false)
                {
                    tmp_header = tmp_header.Replace(",対象TBL名", "");
                }
                Set_LogHeader(CommonModule.LogFilePath, tmp_header);
                // 20161104 ログ出力内容修正(開発用以外はテーブル名を表示しない) -chg end
            }

            // ログ格納用テーブル削除(DROP)
            if (logdropflg)
            {
                // 中間ファイルチェックログ
                string tmp_sql_midchklogdrop = DBQuery.Qry_DropInfo(CommonModule.LOG_TMP_MIDCHKTABLENAME, true);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_midchklogdrop, ref rowcnt);
                // コンバートログ
                string tmp_sql_logdrop = DBQuery.Qry_DropInfo(CommonModule.LOG_TMP_TABLENAME, true);
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_logdrop, ref rowcnt);
            }

        }

        #endregion

        #region 開発用処理

        /// <summary>
        /// 開発用タブ表示ON/OFF
        /// </summary>
        /// <param name="sender">Object</param>
        /// <param name="e">Event</param>
        /// <remarks>
        /// ・開発用タブ表示時に押下…前回表示タブを表示
        /// ・上記以外時に押下…開発用タブを表示
        ///   ※UI(タブINDEX)を変更した場合、ここを修正する
        /// </remarks>
        private void btnDevTabChange_Click(object sender, EventArgs e)
        {

            Chg_TabSelect();

        }

        /// <summary>
        /// 開発用タブ表示制御
        /// </summary>
        /// <remarks></remarks>
        public void Chg_TabSelect()
        {
            string seltab = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(tabCtrlMain.SelectedTab.Name), "", tabCtrlMain.SelectedTab.Name));
            int tmptaihitab;
            int devtab = 0;               // 開発用タブINDEX

            // ①現開いているタブINDEX値を退避
            switch (seltab ?? "")
            {
                case "tabPageStart":                 // メイン.開始TAB
                    {
                        tmptaihitab = 1;
                        break;
                    }
                case "tabPageSession":               // メイン.DB接続設定TAB
                    {
                        tmptaihitab = 2;
                        break;
                    }
                case "tabPageSyoki":                 // メイン.初期設定TAB
                    {
                        tmptaihitab = 3;
                        break;
                    }
                case "tabPageMenu":                  // メイン.作業選択TAB
                    {
                        tmptaihitab = 4;
                        break;
                    }
                case "tabPageJizen":                 // メイン.事前作業TAB
                    {
                        tmptaihitab = 5;
                        break;
                    }
                case "tabPageJigo":                  // メイン.事後作業TAB
                    {
                        tmptaihitab = 6;
                        break;
                    }
                case "tabPageHojyo":                 // メイン.補助機能TAB
                    {
                        tmptaihitab = 7;
                        break;
                    }
                case "tabPageHajimeni":              // コンバーター.はじめにTAB
                    {
                        tmptaihitab = 8;
                        break;
                    }
                case "tabPageSelect":                // コンバーター.項目選択TAB
                    {
                        tmptaihitab = 9;
                        break;
                    }
                case "tabPageJikko":                 // コンバーター.コンバート実行TAB
                    {
                        tmptaihitab = 10;
                        break;
                    }
                case "tabPageEndOK":                 // コンバーター.終了(正常)TAB
                    {
                        tmptaihitab = 11;
                        break;
                    }
                case "tabPageEndError":              // コンバーター.終了(エラー)TAB
                    {
                        tmptaihitab = 12;
                        break;
                    }
                case "tabPageEndCancel":             // コンバーター.終了(中断)TAB
                    {
                        tmptaihitab = 13;
                        break;
                    }
                case "tabPageIkkatu":                // (仮)コンバーター.一括設定TAB
                    {
                        tmptaihitab = 14;
                        break;
                    }
                case "tabPageHanyoJizen":            // (仮)コンバーター.汎用CVK用TAB
                    {
                        tmptaihitab = 15;
                        break;
                    }

                default:
                    {
                        tmptaihitab = devtab;
                        break;
                    }
            }

            // ②現開いているタブ毎の制御
            switch (tmptaihitab)
            {
                case var @case when @case == devtab:
                    {
                        // ※開発用タブが開かれている場合は前回のタブを表示
                        switch (CommonModule.taihitab)
                        {
                            case 1:
                                {
                                    tabCtrlMain.SelectedTab = tabPageStart;
                                    break;
                                }
                            case 2:
                                {
                                    tabCtrlMain.SelectedTab = tabPageSession;
                                    break;
                                }
                            case 3:
                                {
                                    tabCtrlMain.SelectedTab = tabPageSyoki;
                                    break;
                                }
                            case 4:
                                {
                                    tabCtrlMain.SelectedTab = tabPageMenu;
                                    break;
                                }
                            case 5:
                                {
                                    tabCtrlMain.SelectedTab = tabPageJizen;
                                    break;
                                }
                            case 6:
                                {
                                    tabCtrlMain.SelectedTab = tabPageJigo;
                                    break;
                                }
                            case 7:
                                {
                                    tabCtrlMain.SelectedTab = tabPageHojyo;
                                    break;
                                }
                            case 8:
                                {
                                    tabCtrlMain.SelectedTab = tabPageHajimeni;
                                    break;
                                }
                            case 9:
                                {
                                    tabCtrlMain.SelectedTab = tabPageSelect;
                                    break;
                                }
                            case 10:
                                {
                                    tabCtrlMain.SelectedTab = tabPageJikko;
                                    break;
                                }
                            case 11:
                                {
                                    tabCtrlMain.SelectedTab = tabPageEndOK;
                                    break;
                                }
                            case 12:
                                {
                                    tabCtrlMain.SelectedTab = tabPageEndError;
                                    break;
                                }
                            case 13:
                                {
                                    tabCtrlMain.SelectedTab = tabPageEndCancel;
                                    break;
                                }
                            case 14:
                                {
                                    tabCtrlMain.SelectedTab = tabPageIkkatu;
                                    break;
                                }
                            case 15:
                                {
                                    tabCtrlMain.SelectedTab = tabPageHanyoJizen;
                                    break;
                                }

                            default:
                                {
                                    break;
                                }

                        }
                        tabPageManager.ChangeTabPageVisible(CommonModule.taihitab, true);
                        tabPageManager.ChangeTabPageVisible(devtab, false);
                        Chg_DevBtn(true);
                        break;
                    }

                default:
                    {
                        // ※現開発用タブ以外が開かれている場合は開発用タブを表示
                        tabCtrlMain.SelectedTab = tabPageDev;
                        tabPageManager.ChangeTabPageVisible(devtab, true);
                        tabPageManager.ChangeTabPageVisible(tmptaihitab, false);
                        Chg_DevBtn(false);
                        break;
                    }
            }
            CommonModule.taihitab = tmptaihitab;

        }

        /// <summary>
        /// 開発用タブ表示状態のボタン制御(表示)
        /// </summary>
        /// <param name="condi">状態…True.表示、False.非表示</param>
        /// <remarks></remarks>
        public void Chg_DevBtn(bool condi)
        {

            // 「戻る」ボタンへ元の情報格納
            btnBack.Enabled = condi;
            // 「次へ」ボタンへ元の情報格納
            btnNext.Enabled = condi;
            // 「中止」ボタンへ元の情報格納
            btnEnd.Enabled = condi;

        }

        #endregion

        #region コンバート後データ調整処理

        /// <summary>
        /// データ調整を行うメソッド 20160929 データ調整用メソッドの作成 -add
        /// </summary>
        /// <remarks></remarks>
        private void DataCond()
        {

            int tmp_cnt = 0;

            // 契約者口座情報更新処理
            var list_kyscondqry = DataCondModule.Get_List_Qry_KysKozaCond();
            foreach (var tmp_sql in list_kyscondqry)
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
            // 自社口座情報のゆうちょ→銀行変換処理
            DataCondModule.Get_List_Qry_JisyaKozaCond(sqlcnnv10);
            // 適用税率更新処理
            DataCondModule.Set_ZeiMst(sqlcnnv10);
            // メモタイトル一括更新処理(汎用CVのみ)
            if (CommonModule.CNVNO == (int)CommonModule.ConvertTypes._汎用)
            {
                var list_bikotitlecondqry = DataCondModule.Get_List_Qry_BikoTitleCond();
                foreach (var tmp_sql in list_bikotitlecondqry)
                    DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, ref tmp_cnt);
            }
            // 契約情報のデータ調整
            DataCondModule.Set_KydataCond(sqlcnnv10);
            // 契約者口座情報のデータ調整
            DataCondModule.Set_KysdataKozaCond(sqlcnnv10);
            // 当月分の毎月入金項目データ調整
            DataCondModule.Set_KydataNkinCond(sqlcnnv10);
            // 自社web連動用テーブル・フィールドのデータ調整
            DataCondModule.Set_WmpTable(sqlcnnv10);
        }

        #endregion

        #region コンバート実績を元にした処理

        /// <summary>
        /// コンバート実績をXMLファイルへ保持しておく処理 '20160707 コンバート実績保持の処理追加 -add
        /// </summary>
        /// <remarks></remarks>
        private void MakeFile_CVJisseki()
        {

            // 移行項目取得
            var list_cvitem = new List<string>();
            list_cvitem = Get_ListCVChkitemAll(true);

            // ※コンバート後なので0件は無いが念の為
            if (list_cvitem.Count == 0)
            {
                return;
            }

            // コンバート実績保存ファイル格納先取得
            string cvjissekiinfodirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_INI_NAME);
            string cvjissekifilepath = "";
            cvjissekifilepath = EtcMethod.Set_Path(cvjissekiinfodirpath, CommonModule.FILE_CVJISSEKINAME);

            // xml作成
            // 最初の固定文字列
            string strxml_pre = "<?xml version='1.0'?>" + "<cvjisseki>" + "<!--コンバート実績-->";


            // 最後の固定文字列
            string strxml_post = "</cvjisseki>";

            // 移行項目の文字列
            string strxml_main = "";
            string tmp_str = "";
            int cntii = 1;
            foreach (var cvitem in list_cvitem)
            {
                tmp_str = tmp_str + "<cvitem" + cntii.ToString() + ">" + cvitem + "</cvitem" + cntii.ToString() + ">";
                cntii = cntii + 1;
            }
            strxml_main = tmp_str;

            // 統合
            string strxml_total = strxml_pre + strxml_main + strxml_post;

            var xmlDoc = new System.Xml.XmlDocument();

            // 文字列からDOMドキュメントを生成
            xmlDoc.LoadXml(strxml_total);
            // 20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg sta
            // Try
            // '作成したDOMドキュメントをファイルに保存
            // xmlDoc.Save(cvjissekifilepath)
            // Catch ex As System.Xml.XmlException
            // 'XMLによる例外をキャッチ
            // Console.WriteLine(ex.Message)
            // Catch ex As Exception
            // 'その他の例外をキャッチ
            // Console.WriteLine(ex.Message)
            // End Try
            try
            {
                // 作成したDOMドキュメントをファイルに保存
                xmlDoc.Save(cvjissekifilepath);
            }

            // アクセス拒否
            // Catch ex As UnauthorizedAccessException
            // MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
            // "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
            // "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            // Return False
            // その他の例外(アクセス拒否を含む)
            catch (Exception ex)               // 20161108 レビュー結果：「セットアップフォルダ内～拒否されたため、コンバート実績を～できませんでした。」のように修正
            {
                // 20161108_2 レビュー結果戻り修正 -chg sta
                // MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                // "コンバート実績を保持することができませんでした。" & vbCrLf & _
                // "コンバート実績の保持を有効にするためにはセットアップ先をアクセス権限のあるフォルダに変更して、" & _
                // "再度セットアップを行ってから、本プログラムを実行して下さい。" & vbCrLf & _
                // "※コンバート履歴の保持のみであるためコンバートデータに影響はありません。", _
                // "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                CommonModule.MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されため、" + Constants.vbCrLf + "コンバート実績を保持することができませんでした。" + Constants.vbCrLf + "コンバート実績の保持を有効にするためにはセットアップ先をアクセス権限のあるフォルダに変更して、" + "再度セットアップを行ってから、本プログラムを実行して下さい。" + Constants.vbCrLf + "※コンバートデータへの影響は無いためこのまま作業を続行しても問題はありません。", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);




                // 20161108_2 レビュー結果戻り修正 -chg end
            }                             // 20161108 レビュー結果：「※コンバート履歴～」は除去またはもう少しわかり易い内容で。
            // 20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg end
        }

        /// <summary>
        /// コンバート実績の削除(実績を残したxmlファイルを削除) '20160707 コンバート実績保持の処理追加 -add
        /// </summary>
        /// <remarks></remarks>
        private void DeleteFile_CVJisseki()
        {

            // 保管されているコンバート実績ファイルの読込
            string cvjissekiinfodirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_INI_NAME);
            string cvjissekifilepath = "";
            cvjissekifilepath = EtcMethod.Set_Path(cvjissekiinfodirpath, CommonModule.FILE_CVJISSEKINAME);

            // ファイルが存在しない(コンバート実績が存在しないまたはファイル名が編集されている)場合は処理を抜ける
            if (EtcMethod.Chk_FileExist(cvjissekifilepath) == false)
            {
                return;
            }

            // ファイル削除
            // 20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg sta
            // File.Delete(cvjissekifilepath)
            try
            {
                // 作成したDOMドキュメントをファイルに保存
                File.Delete(cvjissekifilepath);
            }

            // アクセス拒否
            // Catch ex As UnauthorizedAccessException
            // MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
            // "セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", _
            // "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            // Return False
            // その他の例外(アクセス拒否を含む)
            catch (Exception ex)           // 20161108 レビュー結果：MakeFile_CVJissekiと同様
            {
                // 20161108_2 レビュー結果戻り修正 -chg sta
                // MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されました。" & vbCrLf & _
                // "コンバート実績を削除することができませんでした。" & vbCrLf & _
                // "コンバート実績の削除を有効にするためにはセットアップ先をアクセス権限のあるフォルダに変更して、" & _
                // "再度セットアップを行ってから、本プログラムを実行して下さい。" & vbCrLf & _
                // "※コンバート履歴の削除のみであるためコンバートデータに影響はありません。", _
                // "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                CommonModule.MsgResult = MessageBox.Show("セットアップフォルダ内のファイルへのアクセスが拒否されたため、" + Constants.vbCrLf + "コンバート実績を削除することができませんでした。" + Constants.vbCrLf + "コンバート実績の削除を有効にするためにはセットアップ先をアクセス権限のあるフォルダに変更して、" + "再度セットアップを行ってから、本プログラムを実行して下さい。" + Constants.vbCrLf + "※コンバートデータへの影響は無いためこのまま作業を続行しても問題はありません。", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);




                // 20161108_2 レビュー結果戻り修正 -chg end
            }
            // 20161104 iniフォルダ内ファイル書込み処理時のエラー対応 -chg end

        }

        /// <summary>
        /// コンバート実績読込処理 '20160707 コンバート実績保持の処理追加
        /// </summary>
        /// <remarks></remarks>
        private void Set_CVJisseki()
        {

            // ボタンの活性制御
            btnRekiClear.Enabled = true;                                 // 20160711 コンバート履歴処理の追加 -add

            // 項目名とチェックボックスを紐付けたハッシュテーブルを生成
            var hash_cvitemtochkbox = new SafeDictionary<string, CheckBox>();
            hash_cvitemtochkbox = Get_Hash_CVChkitemAll();

            // チェックボックスのチェックON、着色初期化
            foreach (var chkbox in hash_cvitemtochkbox)
            {
                var tmp_chkbox = new CheckBox();
                tmp_chkbox = (CheckBox)chkbox.Value;
                tmp_chkbox.ForeColor = Color.Black;
                // 20161014 着色処理の修正 -chg sta
                // tmp_chkbox.Checked = True
                switch (CommonModule.CNVNO)
                {
                    case (int)CommonModule.ConvertTypes._汎用:
                        {
                            tmp_chkbox.Checked = false;
                            break;
                        }
                }
                // 20161014 着色処理の修正 -chg end
            }

            // 保管されているコンバート実績ファイルの読込
            string cvjissekiinfodirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_INI_NAME);
            string cvjissekifilepath = "";
            cvjissekifilepath = EtcMethod.Set_Path(cvjissekiinfodirpath, CommonModule.FILE_CVJISSEKINAME);

            // ファイルが存在しない(コンバート実績が存在しないまたはファイル名が編集されている)場合は処理を抜ける
            if (EtcMethod.Chk_FileExist(cvjissekifilepath) == false)
            {
                // ボタンの活性制御
                btnRekiClear.Enabled = false;                            // 20160711 コンバート履歴処理の追加 -add
                return;
            }

            // xmlファイル読込処理
            System.Xml.XmlReader xmlreader;
            string item = "";
            string data = "";
            string itemprename = "cvitem";
            var list_cvjissekiitem = new List<string>();
            int itemcnt = 1;
            xmlreader = System.Xml.XmlReader.Create(cvjissekifilepath);

            while (xmlreader.Read())
            {
                if (xmlreader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    // データ取得
                    item = xmlreader.LocalName;
                    data = xmlreader.ReadString();

                    if ((item ?? "") == (itemprename + itemcnt.ToString() ?? ""))
                    {
                        list_cvjissekiitem.Add(data);
                        itemcnt = itemcnt + 1;
                    }
                }
            }

            xmlreader.Close();

            if (list_cvjissekiitem.Count == 0)                        // ※コンバート実績が残っている段階で0は無いが念の為
            {
                // ボタンの活性制御
                btnRekiClear.Enabled = false;                            // 20160711 コンバート履歴処理の追加 -add
                return;
            }

            // コンバート実績に残っている項目と照合し、一致したチェックボックスのチェックOFFと着色
            foreach (var cvitem in list_cvjissekiitem)
            {
                var chkbox = new CheckBox();
                if (hash_cvitemtochkbox.ContainsKey(cvitem))
                {
                    chkbox = (CheckBox)hash_cvitemtochkbox[cvitem];
                    chkbox.Checked = false;
                    chkbox.ForeColor = Color.Green;
                }
            }

        }

        #endregion

    }

}