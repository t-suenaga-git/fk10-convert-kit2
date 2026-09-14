using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Converter10.Njc.N3Lib.Utys;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Common
{

    #region 共通処理クラス

    /// <summary>
    /// 共通処理
    /// </summary>
    /// <remarks></remarks>
    public class CommonRepository
    {

        /// <summary>
        /// 進捗率をラベルに表示
        /// </summary>
        /// <param name="pgbcnt"></param>
        /// <param name="pgbtotalcnt"></param>
        /// <remarks></remarks>
        public void ProgressOutPut(int pgbcnt, int pgbtotalcnt, bool totalflg = false)
        {

            {
                var withBlock = My.MyProject.Forms.MainFrm;

                if (totalflg)
                {
                    withBlock.lblPgbTotal.Text = Conversion.Int(pgbcnt / (double)pgbtotalcnt * 100d).ToString() + " %";
                }
                else
                {
                    withBlock.lblPgbPartial.Text = Conversion.Int(pgbcnt / (double)pgbtotalcnt * 100d).ToString() + " %";
                }

            }

        }

        // 20161009 改善対応：中間ファイルチェック時の進捗表示 -add sta
        // ※とりあえずメイン進捗とは別に用意し、既存U用コンバーター対応時に修正する
        /// <summary>
        /// 進捗率をラベルに表示
        /// </summary>
        /// <param name="pgbcnt"></param>
        /// <param name="pgbtotalcnt"></param>
        /// <remarks></remarks>
        public void ProgressChkOutPut(int pgbcnt, int pgbtotalcnt, bool totalflg = false)
        {

            {
                var withBlock = My.MyProject.Forms.MainFrm;

                if (totalflg)
                {
                    withBlock.lblPgbCheck.Text = Conversion.Int(pgbcnt / (double)pgbtotalcnt * 100d).ToString() + " %";
                }
                else
                {
                    // .lblPgbChkPartial.Text = Int(((pgbcnt / pgbtotalcnt) * 100)).ToString & " %"
                }

            }

        }
        // 20161009 改善対応：中間ファイルチェック時の進捗表示 -add end

    }

    #endregion

    #region タブページ表示設定クラス

    /// <summary>
    /// タブページ表示処理
    /// </summary>
    /// <remarks></remarks>
    public class TabPageManager
    {

        // --------------------------------------------------
        // 変数
        // --------------------------------------------------
        private TabPageInfo[] _tabPageInfos = null;
        private TabControl _tabControl = null;


        /// <summary>
        /// TabPage情報セット
        /// </summary>
        /// <remarks></remarks>
        private class TabPageInfo
        {
            public TabPage objTabPage;
            public bool blnVisible;

            /// <summary>
            /// インスタンス化
            /// </summary>
            /// <param name="page"></param>
            /// <param name="visi"></param>
            /// <remarks></remarks>
            public TabPageInfo(TabPage page, bool visi)
            {
                objTabPage = page;
                blnVisible = visi;
            }
        }

        /// <summary>
        /// TabPageManagerクラスのインスタンス化
        /// </summary>
        /// <param name="crl">基となるTabControlオブジェクト</param>
        /// <remarks></remarks>
        public TabPageManager(TabControl crl)
        {
            _tabControl = crl;
            _tabPageInfos = new TabPageInfo[_tabControl.TabPages.Count];

            int ii;
            var loopTo = _tabControl.TabPages.Count - 1;
            for (ii = 0; ii <= loopTo; ii++)
                _tabPageInfos[ii] = new TabPageInfo(_tabControl.TabPages[ii], true);
        }

        /// <summary>
        /// TabPageの表示・非表示設定
        /// </summary>
        /// <param name="index">変更するTabPageのIndex番号</param>
        /// <param name="visi">TABページ表示…TRUE.表示 FALSE.非表示</param>
        /// <remarks></remarks>
        public void ChangeTabPageVisible(int index, bool visi)
        {
            if (_tabPageInfos[index].blnVisible == visi)
            {
                return;
            }
            _tabPageInfos[index].blnVisible = visi;
            _tabControl.SuspendLayout();
            _tabControl.TabPages.Clear();

            int ii;
            var loopTo = _tabPageInfos.Length - 1;
            for (ii = 0; ii <= loopTo; ii++)
            {
                if (_tabPageInfos[ii].blnVisible)
                {
                    _tabControl.TabPages.Add(_tabPageInfos[ii].objTabPage);
                }
            }
            _tabControl.ResumeLayout();
        }

    }

    #endregion

    #region 接続処理関連クラス

    public class DBConnection
    {

        private bool _rtn;

        /// <summary>
        /// 接続処理 (OPEN)
        /// </summary>
        /// <param name="cnninfo">接続情報</param>
        /// <param name="sqlcnn">接続 [out]</param>
        /// <param name="flg_authent">認証方法flg</param>
        /// <returns>接続状態<br/> True.成功<br/> False.失敗<br/></returns>
        /// <remarks>・接続情報</remarks>
        public bool CnnSession(Model.DefSQLConnection cnninfo, ref SqlConnection sqlcnn, bool flg_authent)


        {

            var cnnset = new SqlConnection();

            _rtn = true;

            try
            {

                if (flg_authent)
                {

                    // SQLServer認証
                    cnnset.ConnectionString = "Persist Security Info=True" + ";Data Source = " + cnninfo.ServerName + ";Initial Catalog = " + cnninfo.InitialCatalog + ";User ID = " + cnninfo.User + ";Password = " + cnninfo.Pass + ";Connection Timeout = " + cnninfo.TimeOut;





                }
                else
                {

                    // Windows認証
                    cnnset.ConnectionString = "Persist Security Info=True" + ";Data Source = " + cnninfo.ServerName + ";Initial Catalog = " + cnninfo.InitialCatalog + ";Integrated Security = SSPI" + ";Connection Timeout = " + cnninfo.TimeOut;





                }


                cnnset.Open();
                Console.WriteLine("{0}の{1}に接続しました", cnninfo.ServerName, cnninfo.InitialCatalog);
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error! {0}", ex.Message);
                _rtn = false;
            }
            finally
            {
                sqlcnn = cnnset;
                CommonModule.LimitTimeOut = Conversions.ToInteger(cnninfo.TimeOut);

            }
            return _rtn;

        }

        /// <summary>
        /// 接続処理(CLOSE:リソース解放)
        /// </summary>
        /// <param name="sqlcnn">接続情報</param>
        public void CnnClose(SqlConnection sqlcnn)
        {

            if (sqlcnn is not null)
            {
                if (CnnCheck(sqlcnn) == false)
                {
                    sqlcnn.Close();
                }
                sqlcnn.Dispose();
            }

        }

        /// <summary>
        /// 接続状態をチェック
        /// </summary>
        /// <param name="sqlcnn">接続情報</param>
        /// <returns>接続状態：<br/> True.接続<br/> False.非接続<br/></returns>
        public bool CnnCheck(SqlConnection sqlcnn)
        {

            _rtn = true;
            if (sqlcnn.State != ConnectionState.Closed)
            {
                _rtn = false;
            }
            return _rtn;

        }

    }

    #endregion

    #region インターフェース

    public interface IConv
    {

        // 抽出クエリ(VIEW作成用)
        string Get_UseQry(ref string sortstr);

        // 中間ファイル読込→変数→DB書込
        bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt);




        // 親マスタ有無確認用オブジェクト作成
        void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata);

        // 追加コンバート時の既存データのキーを取得してオブジェクトへ格納
        void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata);

        // データチェック(特別にチェックが必要な場合)
        bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt);

    }

    #endregion

    #region プログレスバー表示設定クラス

    public class ProgressBarManager
    {

        /// <summary>
        /// 個別進捗プログレスバー初期設定
        /// </summary>
        /// <param name="int"></param>
        /// <remarks></remarks>
        public void pgbInitPart(int @int)
        {

            {
                var withBlock = My.MyProject.Forms.MainFrm.pgbpartial;
                withBlock.Minimum = 0;
                withBlock.Value = 0;
                withBlock.Maximum = @int;
            }

        }

        /// <summary>
        /// 全体進捗プログレスバー初期設定
        /// </summary>
        /// <param name="int"></param>
        /// <remarks></remarks>
        public void pgbInitTotal(int @int)
        {

            {
                var withBlock = My.MyProject.Forms.MainFrm.pgbTotal;
                withBlock.Minimum = 0;
                withBlock.Value = 0;
                withBlock.Maximum = @int;
            }

        }

        /// <summary>
        /// 個別進捗プログレスバー更新
        /// </summary>
        /// <param name="int"></param>
        /// <remarks></remarks>
        public void pgbsettingPart(int @int)
        {

            {
                var withBlock = My.MyProject.Forms.MainFrm.pgbpartial;
                if (@int < withBlock.Maximum)
                {
                    withBlock.Value = @int + 1;
                    withBlock.Value = @int;
                }
                else
                {
                    withBlock.Maximum = withBlock.Maximum + 1;
                    withBlock.Value = @int + 1;
                    withBlock.Value = @int;
                    withBlock.Maximum = withBlock.Maximum - 1;
                }
            }

        }

        /// <summary>
        /// 全体進捗プログレスバー更新
        /// </summary>
        /// <param name="int"></param>
        /// <remarks></remarks>
        public void pgbsettingTotal(int @int)
        {

            {
                var withBlock = My.MyProject.Forms.MainFrm.pgbTotal;
                if (@int < withBlock.Maximum)
                {
                    withBlock.Value = @int + 1;
                    withBlock.Value = @int;
                }
                else
                {
                    withBlock.Maximum = withBlock.Maximum + 1;
                    withBlock.Value = @int + 1;
                    withBlock.Value = @int;
                    withBlock.Maximum = withBlock.Maximum - 1;
                }
            }

        }

        // 20161009 改善対応：中間ファイルチェック時の進捗表示 -add sta
        // ※とりあえずメイン進捗とは別に用意し、既存ユーザ用コンバーター対応時に修正する
        /// <summary>
        /// 全体進捗プログレスバー初期設定
        /// </summary>
        /// <param name="int"></param>
        /// <remarks></remarks>
        public void pgbInitChkTotal(int @int)
        {

            {
                var withBlock = My.MyProject.Forms.MainFrm.pgbCheck;
                withBlock.Minimum = 0;
                withBlock.Value = 0;
                withBlock.Maximum = @int;
            }

        }

        /// <summary>
        /// 全体進捗プログレスバー更新
        /// </summary>
        /// <param name="int"></param>
        /// <remarks></remarks>
        public void pgbsettingChkTotal(int @int)
        {
            Debug.Print(My.MyProject.Forms.MainFrm.pgbCheck.Maximum.ToString());
            {
                var withBlock = My.MyProject.Forms.MainFrm.pgbCheck;
                if (@int < withBlock.Maximum)
                {
                    withBlock.Value = @int + 1;
                    withBlock.Value = @int;
                }
                else
                {
                    withBlock.Maximum = withBlock.Maximum + 1;
                    withBlock.Value = @int + 1;
                    withBlock.Value = @int;
                    withBlock.Maximum = withBlock.Maximum - 1;
                }
            }

        }
        // 20161009 改善対応：中間ファイルチェック時の進捗表示 -add end
    }

    #endregion

    #region データチェッククラス(ログ出力文字列を設定)

    public class DataChk
    {

        /// <summary>
        /// 日付チェック
        /// </summary>
        /// <param name="chkstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataDate(string chkstr, ref string chkstrafter, string min_str, string max_str, string defvalue, ref string errstr)
        {

            bool rtn = true;
            DateTime chgdate;

            // テスト用                                   'kakaka 日付チェックの内容(min_str～max_strの範囲)を仕様書に記載しておく
            // 2016.03.28 日付の最大最小値の移動 -del sta
            // min_str = "1900/01/01"                      'kakaka4 CommonModule.vb へ移動(ここでもいいが、後で編集が予想されるものは1か所に纏めたい為)
            // max_str = "2100/12/31"
            // 2016.03.28 日付の最大最小値の移動 -del end

            // 2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            // Dim tmp_date_min As DateTime = DateTime.Parse(min_str)          'kakaka4 0322_1300 念の為、変換時の型落ちを考慮。
            // Dim tmp_date_max As DateTime = DateTime.Parse(max_str)          'kakaka4 0322_1300 念の為、変換時の型落ちを考慮。
            DateTime tmp_date_min;
            DateTime tmp_date_max;
            if (DateTime.TryParse(min_str, out tmp_date_min) == false)
            {
                tmp_date_min = Conversions.ToDate(CommonModule.DEF_MIN_YMD);
            }
            if (DateTime.TryParse(max_str, out tmp_date_max) == false)
            {
                tmp_date_max = Conversions.ToDate(CommonModule.DEF_MAX_YMD);
            }
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg end

            string def = defvalue;

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                // chkstrafter = def                       'kakaka 所有期間、契約期間など、チェックした期間に関連する期間がある場合はデフォルト設定は不可
                // If def <> "" Then
                // errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                // rtn = False
                // End If
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                // 20160929 デフォルト値設定処理の追加 -add sta
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                }
                else
                {
                    chkstrafter = chkstr;
                }
                // 20160929 デフォルト値設定処理の追加 -add end
                return rtn;
            }

            // 日付変換チェック
            if (!DateTime.TryParse(chkstr, out chgdate))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 範囲チェック
            if (!(tmp_date_min <= chgdate && chgdate <= tmp_date_max))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 20170524 竣工日がyyyy/mm/dd形式以外の日付の場合の移行処理対応 -chg sta
            // '正常値はそのまま格納
            // chkstrafter = chkstr
            // チェックで問題ない場合は変換した日付を返却する
            chkstrafter = chgdate.ToString();
            // 20170524 竣工日がyyyy/mm/dd形式以外の日付の場合の移行処理対応 -chg end

            // 判定結果を返却
            return rtn;

        }

        /// <summary>
        /// 時間チェック
        /// </summary>
        /// <param name="chkstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataTime(string chkstr, ref string chkstrafter, ref string errstr)
        {

            bool rtn = true;

            // 構築中

            return rtn;

        }

        /// <summary>
        /// 数値チェック(整数)
        /// </summary>
        /// <param name="chkstr"></param>
        /// <param name="max"></param>
        /// <param name="min"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataNumeric(string chkstr, ref string chkstrafter, string min_str, string max_str, string defvalue, ref string errstr)
        {

            long lng = 0L;
            bool rtn = true;
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            // Dim min As Long = Int64.Parse(min_str)                      'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            // Dim max As Long = Int64.Parse(max_str)                      'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            long min = 0L;
            long max = 0L;
            if (long.TryParse(min_str, out min) == false || long.TryParse(max_str, out max) == false)
            {
                errstr = CommonModule.LOG_NAIYO_ERR_DATACHK + "-" + CommonModule.LOG_HUBI_ERR_DATACHK_A + "-" + CommonModule.LOG_TAISYO_DATACHK_A;
                rtn = false;
                return rtn;
            }
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg end

            string def = defvalue;

            // 20161130 契約者入居フラグの移行処理修正_再 -chg sta
            // 空文字チェックで「0」が含まれていたため空文字と「0」でチェックを分ける
            // '空文字チェック
            // If chkstr = "" Or chkstr = "0" Then     '20160603 ユーザーデータ検証による修正 条件に「0」を追加
            // '20160829 送金予定日デフォルト値設定 -chg sta
            // 'chkstrafter = chkstr                '20160603 ユーザーデータ検証による修正 条件に「0」の追加に伴いチェック対象をそのまま返却する処理を追加
            // If def <> "" Then
            // chkstrafter = def
            // Else
            // chkstrafter = chkstr
            // End If
            // '20160829 送金予定日デフォルト値設定 -chg end
            // '2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
            // 'chkstrafter = def
            // 'If def <> "" Then
            // '    errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT        'kakaka 中間ファイルチェックの場合でもデフォルト値に設定と表示される(以下チェック、他メソッドも同様)
            // '    rtn = False
            // 'End If
            // '2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
            // Return rtn
            // End If

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                if (!string.IsNullOrEmpty(def))
                {
                    // デフォルト値がある場合は設定する
                    chkstrafter = def;
                }
                return rtn;
            }

            // 「0」チェック
            // --------------------------------- ↓「0」チェックの内容↓ ---------------------------------
            // 最小値が「0」以下の場合は「0」を値として移行する
            // 最小値が「0」より大きい場合は値として意味をなさないので空データで移行する
            // ※物件情報に登録されている業者情報No等は画面上で未設定でもDB上で「0」として登録されている
            // 業者情報No等は1から始まる値なので値は空データとして移行する必要がある
            // --------------------------------- ↑「0」チェックの内容 ↑ ---------------------------------
            if (chkstr == "0")
            {
                if (min <= 0L)
                {
                    // 最小値が「0」以下の場合はそのまま移行
                    chkstrafter = chkstr;
                    return rtn;
                }
                else if (min > 0L)
                {
                    // 最小値が「0」より大きい場合は空データを移行
                    chkstrafter = "";
                    return rtn;
                }
            }
            // 20161130 契約者入居フラグの移行処理修正_再 -chg end

            // 数値変換チェック
            if (!long.TryParse(chkstr, out lng))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 範囲チェック
            if (!(min <= lng && lng <= max))             // kakaka Or→OrElse、And→AndAlsoに変更してください。(他のチェックのところも)
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 正常値はそのまま格納
            chkstrafter = chkstr;

            // 判定結果を返却
            return rtn;

        }

        /// <summary>
        /// 数値チェック(Decimal)
        /// </summary>
        /// <param name="chkstr"></param>
        /// <param name="max"></param>
        /// <param name="min"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataDecimal(string chkstr, ref string chkstrafter, string min_str, string max_str, string defvalue, ref string errstr)
        {

            decimal dec = 0m;
            bool rtn = true;
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            // Dim min As Decimal = Decimal.Parse(min_str)             'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            // Dim max As Decimal = Decimal.Parse(max_str)             'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            decimal min = 0m;
            decimal max = 0m;
            if (decimal.TryParse(min_str, out min) == false || decimal.TryParse(max_str, out max) == false)
            {
                errstr = CommonModule.LOG_NAIYO_ERR_DATACHK + "-" + CommonModule.LOG_HUBI_ERR_DATACHK_A + "-" + CommonModule.LOG_TAISYO_DATACHK_A;
                rtn = false;
                return rtn;
            }
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg end

            string def = defvalue;

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                // chkstrafter = def
                // If def <> "" Then
                // errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                // rtn = False
                // End If
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
                // 20160929 デフォルト値設定処理の追加 -add sta
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                }
                else
                {
                    chkstrafter = chkstr;
                }
                // 20160929 デフォルト値設定処理の追加 -add end
                return rtn;
            }

            // 数値変換チェック
            if (!decimal.TryParse(chkstr, out dec))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 範囲チェック
            if (!(min <= dec && dec <= max))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 正常値はそのまま格納
            chkstrafter = chkstr;

            // 判定結果を返却
            return rtn;

        }

        /// <summary>
        /// 数値チェック(小数)
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataDouble(string chkstr, ref string chkstrafter, string min_str, string max_str, string defvalue, ref string errstr)
        {

            double dbl = 0d;
            bool rtn = true;
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            // Dim min As Double = Double.Parse(min_str)                   'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            // Dim max As Double = Double.Parse(max_str)                   'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            double min = 0d;
            double max = 0d;
            if (double.TryParse(min_str, out min) == false || double.TryParse(max_str, out max) == false)
            {
                errstr = CommonModule.LOG_NAIYO_ERR_DATACHK + "-" + CommonModule.LOG_HUBI_ERR_DATACHK_A + "-" + CommonModule.LOG_TAISYO_DATACHK_A;
                rtn = false;
                return rtn;
            }
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg end

            string def = defvalue;

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                // chkstrafter = def
                // If def <> "" Then
                // errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                // rtn = False
                // End If
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
                // 20160929 デフォルト値設定処理の追加 -add sta
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                }
                else
                {
                    chkstrafter = chkstr;
                }
                // 20160929 デフォルト値設定処理の追加 -add end
                return rtn;
            }

            // 数値変換チェック
            if (!double.TryParse(chkstr, out dbl))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 範囲チェック
            if (!(min <= dbl && dbl <= max))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 正常値はそのまま格納
            chkstrafter = chkstr;

            // 判定結果を返却
            return rtn;

        }

        /// <summary>
        /// 数値チェック(通貨)
        /// </summary>
        /// <param name="chkstr"></param>
        /// <param name="max"></param>
        /// <param name="min"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMoney(string chkstr, ref string chkstrafter, string min_str, string max_str, string defvalue, ref string errstr)
        {

            decimal dec = 0m;
            bool rtn = true;
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            // Dim min As Decimal = Decimal.Parse(min_str)                 'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            // Dim max As Decimal = Decimal.Parse(max_str)                 'kakaka4 0322_1300 数値チェック前に型落ちする可能性あり。変換時の型落ちを考慮。
            decimal min = 0m;
            decimal max = 0m;
            if (decimal.TryParse(min_str, out min) == false || decimal.TryParse(max_str, out max) == false)
            {
                errstr = CommonModule.LOG_NAIYO_ERR_DATACHK + "-" + CommonModule.LOG_HUBI_ERR_DATACHK_A + "-" + CommonModule.LOG_TAISYO_DATACHK_A;
                rtn = false;
                return rtn;
            }
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg end

            string def = defvalue;
            string tmp_str = "";

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                // chkstrafter = def
                // If def <> "" Then
                // errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                // rtn = False
                // End If
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
                // 20160929 デフォルト値設定処理の追加 -add sta
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                }
                else
                {
                    chkstrafter = chkstr;
                }
                // 20160929 デフォルト値設定処理の追加 -add end
                return rtn;
            }

            // 半角変換
            // 20160531 不正金額の移行制御修正 -chg sta
            // tmp_str = EtcMethod.Get_LenB(EtcMethod.Chg_NumNarrow(chkstr))
            tmp_str = EtcMethod.Chg_NumNarrow(chkstr);
            // 20160531 不正金額の移行制御修正 -chg end

            // カンマチェック
            if (Conversions.ToBoolean(Strings.InStr(tmp_str, ",")))
            {
                tmp_str = Strings.Replace(tmp_str, ",", "");
            }

            // 数値変換チェック
            if (!decimal.TryParse(tmp_str, out dec))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def                                               'kakaka moneyのデフォルト値は0円？
                // '20160531 不正金額の移行制御修正 -add
                // errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_MISMATCH & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;                                               // kakaka moneyのデフォルト値は0円？
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 範囲チェック
            if (!(min <= dec && dec <= max))
            {
                // 20161014 ログ出力内容修正 -chg sta
                // chkstrafter = def
                // '20160531 不正金額の移行制御修正 -add
                // errstr = LOG_NAIYO_ERR_OUTOFRANGE & "-" & LOG_HUBI_OUTOFRANGE & "-" & LOG_TAISYO_DEFAULT
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                else
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE + "-" + CommonModule.LOG_HUBI_OUTOFRANGE_NODEF + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                }
                // 20161014 ログ出力内容修正 -chg end
                rtn = false;
                return rtn;
            }

            // 正常値はそのまま格納
            chkstrafter = chkstr;

            // 判定結果を返却
            return rtn;

        }

        /// <summary>
        /// 論理値チェック
        /// </summary>
        /// <param name="chkstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataBit(string chkstr, ref string chkstrafter, string defvalue, ref string errstr)
        {

            int @int = 0;
            bool rtn = true;
            string def = defvalue;

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del sta
                // chkstrafter = def
                // If def <> "" Then
                // errstr = LOG_NAIYO_ERR_NOTEXISTDATA & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
                // rtn = False
                // End If
                // 2016.02.22 受領データが空の場合はログを出力しないように修正 -del end
                // 20160929 デフォルト値設定処理の追加 -add sta
                if (!string.IsNullOrEmpty(def))
                {
                    chkstrafter = def;
                }
                else
                {
                    chkstrafter = chkstr;
                }
                // 20160929 デフォルト値設定処理の追加 -add end
                return rtn;
            }

            // 数値変換チェック
            if (!int.TryParse(chkstr, out @int))
            {
                chkstrafter = def;
                rtn = false;
                return rtn;
            }

            // 範囲チェック
            if (@int >= 2)
            {
                chkstrafter = def;
                rtn = false;
                return rtn;
            }

            // 正常値はそのまま格納
            chkstrafter = chkstr;

            // 判定結果を返却
            return rtn;

        }

        /// <summary>
        /// 文字列チェック
        /// </summary>
        /// <param name="chkstr"></param>
        /// <param name="max"></param>
        /// <param name="strtype"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataString(string chkstr, ref string chkstrafter, int max, int strtype, ref string errstr)
        {

            bool rtn = true;

            // サイズチェック
            if (max == -1)
            {
                chkstrafter = chkstr;
                return rtn;
            }

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                chkstrafter = chkstr;
                return rtn;
            }

            // バイト数取得                           'kakaka この辺の条件を表など一覧にして仕様書に記載しておく
            int cnt_byte = EtcMethod.Get_LenB(chkstr);

            // 許容範囲チェック
            if (cnt_byte > max)
            {
                if (strtype == 0)
                {
                    // SJIS用
                    chkstrafter = Typ.ToStrSafe(chkstr, max);                     // kakaka メモ：文字列調整
                }
                else
                {
                    // Unicode用
                    chkstrafter = Typ.ToStrSafeUnicode(chkstr, max);
                }
            }
            else
            {
                chkstrafter = chkstr;
            }

            if ((chkstr ?? "") != (chkstrafter ?? ""))
            {
                errstr = CommonModule.LOG_NAIYO_ERR_OUTOFRANGE_STR + "-" + CommonModule.LOG_HUBI_OUTOFRANGE_STR + "-" + CommonModule.LOG_TAISYO_OUTOOFRANGE_STR;
                rtn = false;
            }

            // 外字チェック
            // 外字チェック処理
            // 変換処理をして作業用変数へ格納(未実装)                       'kakaka 10の導入先がV7と同じはずなので、外字はクリアされると思うが、汎用を考えて入れておく
            // 2016.04.04 外字チェック処理追加対応 -add sta
            // 許容範囲内にした状態でチェックを行う(切断された部分に外字が含まれている場合は変換する必要がないため)

            string tmp_gaijichkstr = chkstrafter;
            int cnt_str = chkstrafter.Length;

            // 1文字ずつ抽出して照合
            for (int cntii = 1, loopTo = cnt_str; cntii <= loopTo; cntii++)
            {

                string tmp_chkstr = Strings.Mid(tmp_gaijichkstr, cntii, 1);

                // アプリで処理できない文字
                if (char.GetUnicodeCategory(Conversions.ToChar(tmp_chkstr)) == System.Globalization.UnicodeCategory.OtherNotAssigned)
                {
                    tmp_gaijichkstr = tmp_gaijichkstr.Replace(tmp_chkstr, "■");     // 外字を置換
                }

                // サロゲートペア
                if (char.IsHighSurrogate(Conversions.ToChar(tmp_chkstr)) == true | char.IsLowSurrogate(Conversions.ToChar(tmp_chkstr)) == true)
                {
                    tmp_gaijichkstr = Strings.Left(tmp_gaijichkstr, cntii - 1) + "■" + Strings.Mid(tmp_gaijichkstr, cntii + 2, cnt_str - cntii);     // 外字を置換
                }

                // Unicode文字
                if (char.IsHighSurrogate(Conversions.ToChar(tmp_chkstr)) == false & char.IsLowSurrogate(Conversions.ToChar(tmp_chkstr)) == false)
                {
                    byte[] bytes = null;
                    byte[] encBytes = null;
                    string chgAfterStr = "";
                    var _sjisEncoding = System.Text.Encoding.GetEncoding("Shift_JIS");

                    // 文字列をBytes配列に変換 (unicode -> bytes()unicode)
                    bytes = _sjisEncoding.GetBytes(tmp_chkstr);
                    // Bytes配列の中身をShift_JISに変換 (bytes()unicode -> bytes()Shift_JIS)
                    encBytes = System.Text.Encoding.Convert(_sjisEncoding, System.Text.Encoding.Unicode, bytes);
                    // Bytes配列を文字列に変換 (bytes()Shift_JIS -> unicode)
                    chgAfterStr = System.Text.Encoding.Unicode.GetString(encBytes);

                    if (tmp_chkstr.ToString() != "?" & chgAfterStr == "?")
                    {
                        tmp_gaijichkstr = tmp_gaijichkstr.Replace(tmp_chkstr, "■");     // 外字を置換
                    }

                    bytes.Initialize();
                    encBytes.Initialize();
                    chgAfterStr = "";
                }

            }

            if ((chkstrafter ?? "") != (tmp_gaijichkstr ?? ""))
            {
                errstr = CommonModule.LOG_NAIYO_ERR_GAIJI + "-" + CommonModule.LOG_HUBI_GAIJI + "-" + CommonModule.LOG_TAISYO_GAIJI;
                rtn = false;
                chkstrafter = tmp_gaijichkstr;
            }
            // 2016.04.04 外字チェック処理追加対応 -add end

            return rtn;

        }

        /// <summary>
        /// 口座名義カナチェック 20160706 半角カナの移行処理修正
        /// </summary>
        /// <param name="chkstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_Kana(string chkstr, ref string chkstrafter, ref string errstr)
        {

            // 20160829 口座名義カナチェック機能の追加 -chg sta
            // Dim rtn As Boolean = True
            // Dim tmp_str As String = chkstr

            // '構築中

            // '空文字チェック
            // If chkstr = "" Then
            // Return rtn
            // End If

            // 'カタカナ変換
            // tmp_str = StrConv(tmp_str, VbStrConv.Katakana)

            // '半角変換
            // tmp_str = StrConv(tmp_str, VbStrConv.Narrow)

            // '半角カナチェック

            // ↓↓↓ここから↓↓↓
            bool rtn = true;
            string tmp_str = "";

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                return rtn;
            }

            // 口座名義用に変換
            var kana_obj = new DataChk();
            tmp_str = kana_obj.ToStrKozaKana(chkstr);

            // 口座名義用に変換できなかった場合
            // If normalflg = False Then
            // errstr = LOG_NAIYO_ERR_MISMATCH & "-" & LOG_HUBI_ERR_KANASTR & "-" & LOG_TAISYO_DEFAULT
            // rtn = normalflg
            // Return rtn
            // End If

            // 口座名義用に変換可、かつ変換されている場合
            if ((chkstr ?? "") != (tmp_str ?? ""))
            {
                errstr = CommonModule.LOG_NAIYO_CHG_STR + "-" + CommonModule.LOG_HUBI_CHG_KANASTR + "-" + "";
                rtn = false;
            }

            // 戻り値を格納
            // 20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 -chg sta
            // 大文字変換処理を行う
            // chkstrafter = tmp_str
            chkstrafter = kana_obj.Get_KozaKanaUpper(tmp_str);
            // 20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 -chg end
            // 20160829 口座名義カナチェック機能の追加 -chg end

            return rtn;

        }

        /// <summary>
        /// 口座名義カナ用に変換する(革命10から引用) '20160829 口座名義カナチェック機能の追加 -add
        /// </summary>
        /// <param name="value">対象文字列</param>
        /// <returns>変換後の文字列</returns>
        /// <remarks>
        /// 20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正
        /// 　大文字変換処理の削除に伴い引数「Optional wohenkan As Boolean」を削除
        /// </remarks>
        public string ToStrKozaKana(object value)
        {
            string result = Typ.ToStr(value);

            // ひらがな->カタカナ
            // 全角->半角
            // 小文字 -> 大文字
            result = Strings.StrConv(result, VbStrConv.Katakana | VbStrConv.Narrow | VbStrConv.Uppercase);

            // 20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 -del sta
            // ログ出力制御のためここでは変換しない
            // 'カナのチェック
            // Dim replacefrom As String = _kozakanareplacefrom
            // Dim replaceto As String = _kozakanareplaceto
            // If wohenkan = False Then
            // replacefrom = _kozakanareplacefrom2
            // replaceto = _kozakanareplaceto2
            // End If
            // For thisstep = 0 To replacefrom.Length - 1
            // Dim checkfrom = replacefrom(thisstep)
            // If (0 <= result.IndexOf(checkfrom)) Then
            // Dim checkto = replaceto(thisstep)
            // result = result.Replace(checkfrom, checkto)
            // End If
            // Next
            // 20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 -del end

            // 使用できない文字のチェック
            for (int thisstep = 0, loopTo = _kozakanarareplaceclear.Length - 1; thisstep <= loopTo; thisstep++)
            {
                char checkclear = _kozakanarareplaceclear[thisstep];
                if (0 <= result.IndexOf(checkclear))
                {
                    result = result.Replace(Conversions.ToString(checkclear), string.Empty);
                }
            }

            // 2bytes文字が残っていたら除去
            byte[] bytes = _sjisEncoding.GetBytes(result);
            if (bytes.Length != result.Length)
            {
                string target = result;
                var resultsb = new System.Text.StringBuilder();
                foreach (var @this in target)
                {
                    if (_sjisEncoding.GetBytes(Conversions.ToString(@this)).Length == 1)
                    {
                        resultsb.Append(@this);
                    }
                }
                result = resultsb.ToString();
            }

            return result;
        }

        /// <summary>
        /// 指定文字列を大文字へ変換する処理
        /// </summary>
        /// <param name="value">対象文字列</param>
        /// <param name="wohenkan">「ｦ」→「ｵ」変換フラグ True:変換する False:変換しない</param>
        /// <returns>大文字変換後の文字列</returns>
        /// <remarks>
        /// 20161209 口座名義カナの小文字→大文字変換のログを出力しないように修正 新規追加
        /// 　大文字へ変換した場合の変換ログは不要なためこの処理だけ別メソッドにして個別で処理する
        /// 　処理内容は革命10本体から引用
        /// </remarks>
        public string Get_KozaKanaUpper(string value, bool wohenkan = true)
        {

            string result = Typ.ToStr(value);

            string replacefrom = _kozakanareplacefrom;
            string replaceto = _kozakanareplaceto;
            if (wohenkan == false)
            {
                replacefrom = _kozakanareplacefrom2;
                replaceto = _kozakanareplaceto2;
            }
            for (int thisstep = 0, loopTo = replacefrom.Length - 1; thisstep <= loopTo; thisstep++)
            {
                char checkfrom = replacefrom[thisstep];
                if (0 <= result.IndexOf(checkfrom))
                {
                    char checkto = replaceto[thisstep];
                    result = result.Replace(checkfrom, checkto);
                }
            }

            return result;

        }

        // 20160829 口座名義カナチェック機能の追加 革命10から引用 -add sta
        private readonly System.Text.Encoding _sjisEncoding = System.Text.Encoding.GetEncoding("Shift_JIS");
        private string _kozakanareplacefrom = "ｧｨｩｪｫｬｭｮｯｰ･ヵヶｦ";
        private string _kozakanareplaceto = "ｱｲｳｴｵﾔﾕﾖﾂ-.ｶｹｵ";
        private string _kozakanarareplaceclear = "@`!\"#$%&'*+:;[{<|=]}>^~?_｡､";
        private string _kozakanareplacefrom2 = "ｧｨｩｪｫｬｭｮｯｰ･ヵヶ";
        private string _kozakanareplaceto2 = "ｱｲｳｴｵﾔﾕﾖﾂ-.ｶｹ";
        // 20160829 口座名義カナチェック機能の追加 革命10から引用 -add end

        // 20161005 自社、家主口座デフォルト値設定処理追加 -del sta
        // 家主口座チェック処理と統合するためコメントアウト
        // ''' <summary>
        // ''' 自社口座有無の確認 '20160829 自社口座にデフォルト値を設定する処理を追加 -add
        // ''' </summary>
        // ''' <param name="chkstr"></param>
        // ''' <param name="chkstrafter"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_Jisyakoza(ByVal chkstr As String, ByRef chkstrafter As String, ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True

        // '空の場合は「1：みずほ銀行」をセット
        // If chkstr = "" Or Chk_DataNumeric(chkstr, chkstrafter, "1", "9999", "", errstr) = False Then     '20160928 自社口座取得方法の修正 条件追加
        // errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA & "-" & LOG_TAISYO_DEFAULT
        // chkstrafter = "1"
        // rtn = False
        // Return rtn
        // End If

        // chkstrafter = chkstr

        // Return rtn

        // End Function
        // 20161005 自社、家主口座デフォルト値設定処理追加 -del end

        /// <summary>
        /// 自社、家主口座有無の確認 '20161005 自社、家主口座デフォルト値設定処理追加 -add
        /// </summary>
        /// <param name="chkstr"></param>
        /// <param name="chkstrafter"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_JisyaOwkoza(string chkstr, ref string chkstrafter, string erritem, ref string errstr)
        {

            bool rtn = true;

            // 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正(コメント追記) -add sta
            // 空の場合は「9999:ダミー口座」をセット(DB書込み前に予め作成したデータ)
            // この処理は家賃入金口座情報作成時に、元となる自社口座情報に金融機関情報が存在しないと保存できない
            // 状態になるのを回避するために仮で口座情報を設定する
            // 金融機関Noを設定することで保存可能となるため金融機関Noが空の場合は仮口座を設定する
            // 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正(コメント追記) -add end

            if (string.IsNullOrEmpty(chkstr) | Chk_DataNumeric(chkstr, ref chkstrafter, "1", "9999", "", ref errstr) == false)
            {
                // 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -chg sta
                // Dim tmp_str As String = erritem & "の金融機関情報が存在しないためデフォルト値「1」を設定します。"
                // errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & tmp_str & "-" & LOG_TAISYO_DEFAULT
                // chkstrafter = "1"
                string tmp_str = erritem + "の金融機関情報が存在しないためデフォルト値「9999」を設定します。";
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + tmp_str + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                chkstrafter = "9999";
                // 20161125 自社口座未設定情報へ仮データを登録する金融機関を9999へ変更する修正 -chg end
                rtn = false;
                return rtn;
            }

            chkstrafter = chkstr;

            return rtn;

        }

        /// <summary>
        /// 箇所分類チェック 20160829 箇所未登録データにデフォルト値を設定 -add
        /// 「箇所分類」+「行No」の形でデータをチェックする
        /// データが数値の場合は箇所分類が存在しない場合はデフォルト値をセットする
        /// </summary>
        /// <param name="chkstr"></param>
        /// <param name="chkstrafter"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool ChK_KasyoBrui(string chkstr, ref string chkstrafter, ref string errstr)
        {

            int tmp_int = 0;

            // 数値変換チェック
            if (int.TryParse(chkstr, out tmp_int) == false)
            {
                return true;
            }

            // 数値のみの場合は箇所分類が未設定なのでデフォルト値をセット
            errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA + "-" + CommonModule.LOG_TAISYO_DEFAULT;
            chkstrafter = "箇所未登録" + tmp_int.ToString();
            return false;

        }

        /// <summary>
        /// URL、メールアドレスをデータベース保存用に変換 20160829 メールアドレス、URLの正規化処理を追加
        /// </summary>
        /// <param name="chkstr"></param>
        /// <param name="chkstrafter"></param>
        /// <param name="length"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_URL(string chkstr, ref string chkstrafter, string length_str, ref string errstr)
        {

            bool rtn = true;

            // ※構築中

            string tmp_str = "";
            int length_int = 0;
            int tmp_int = 0;

            // 空文字チェック
            if (string.IsNullOrEmpty(chkstr))
            {
                return rtn;
            }

            // '最大サイズを数値変換
            // If Int32.TryParse(length_str, tmp_int) Then
            // length_int = tmp_int
            // End If

            // '半角変換、半角変換不可文字は半角スペースに置換
            // Dim kana_obj As New DataChk
            // tmp_str = kana_obj.ToStrHankakuKana(chkstr)

            // 'データベース保存用に変換
            // chkstrafter = ToStrSafeUrlEncode(tmp_str, length_int)

            // '変換後のチェック
            // If chkstr <> chkstrafter Then
            // errstr = LOG_NAIYO_CHG_STR & "-" & LOG_HUBI_CHG_URLSTR & "-" & LOG_TAISYO_DEFAULT
            // chkstrafter = chkstrafter
            // rtn = False
            // Return rtn
            // End If

            chkstrafter = chkstr;

            return rtn;

        }

        /// <summary>
        /// 半角カナに変換する(革命10から引用) '20160829 メールアドレス、URLの正規化処理を追加 -add
        /// </summary>
        /// <param name="value"></param>
        /// <param name="flg"></param>
        /// <param name="wohenkan"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public string ToStrHankakuKana(object value)
        {
            string result = Typ.ToStr(value);

            // ひらがな->カタカナ
            // 全角->半角
            result = Strings.StrConv(result, VbStrConv.Katakana | VbStrConv.Narrow);

            // 2bytes文字が残っていたら除去
            byte[] bytes = _sjisEncoding.GetBytes(result);
            if (bytes.Length != result.Length)
            {
                string target = result;
                var resultsb = new System.Text.StringBuilder();
                foreach (var @this in target)
                {
                    if (_sjisEncoding.GetBytes(Conversions.ToString(@this)).Length == 1)
                    {
                        resultsb.Append(@this);
                    }
                    else
                    {
                        resultsb.Append(" ");
                    }
                }
                result = resultsb.ToString();
            }

            return result;
        }

        /// <summary>
        /// 半角英数と一部記号以外の文字を^22のような問題のない文字に置換します。<br/>
        /// データベース保存用です(Web用には使えません。)<br/>
        /// 20160829 メールアドレス、URLの正規化処理を追加
        /// System.Web.HttpUtilityを有効にするために参照の追加で「System.Web」を追加
        /// </summary>
        public static string ToStrSafeUrlEncode(object value, int length)
        {
            string result = Typ.ToStr(value);
            if (Typ.IsStrMissing(value) == false)
            {
                result = System.Web.HttpUtility.UrlEncode(Typ.SubstrByte(value, 0, length));
                result = result.Replace("%", "^");
            }
            return result;
        }

        /// <summary>
        /// 各マスタ有無チェック '20160530 ログ修正
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist(SqlConnection sqlcnnv10, string sql_select, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            // マスタからコードを取得
            var tmp_list = new List<string>();
            bool flg = true;
            bool rtn = true;

            // データが存在しない場合は処理を抜ける
            if (string.IsNullOrEmpty(value) | value == "0")
            {
                return rtn;
            }

            flg = DBExec.Exec_DataReader_Col_List(sql_select, ref sqlcnnv10, ref tmp_list);

            if (tmp_list.Contains(value))
            {
                chkvalue = value;
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_RELEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        // 20160530 ログ修正 -del sta
        // ''' <summary>
        // ''' 都道府県マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_Todofuken(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // 'マスタからコードを取得
        // Dim tmp_list As New List(Of String)
        // Dim tmp_sql As String = " SELECT ken_no FROM m_ken"
        // Dim flg As Boolean = True
        // Dim rtn As Boolean = True

        // '20160530 ログ修正 -add sta
        // If value = "" Then
        // Return rtn
        // End If
        // '20160530 ログ修正 -add end

        // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        // If tmp_list.Contains(value) Then
        // chkvalue = value
        // Else
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_RELEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = ""
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20160530 ログ修正 -del end

        // 20160530 ログ修正 -del sta
        // ''' <summary>
        // ''' 市区町村マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_Sikucyoson(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // 'マスタからコードを取得
        // Dim tmp_list As New List(Of String)
        // Dim tmp_sql As String = " SELECT si_no FROM m_si "
        // Dim flg As Boolean = True
        // Dim rtn As Boolean = True

        // '20160530 ログ修正 -add sta
        // If value = "" Then
        // Return rtn
        // End If
        // '20160530 ログ修正 -add end

        // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        // If tmp_list.Contains(value) Then
        // chkvalue = value
        // Else
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = ""
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20160530 ログ修正 -del end

        /// <summary>
        /// 沿線マスタ有無チェック
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_Ensen(SqlConnection sqlcnnv10, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            // マスタからコードを取得
            var tmp_list = new List<string>();
            string tmp_sql = " SELECT ensen_no FROM m_ensen ";
            bool flg = true;
            bool rtn = true;

            // 20160915 沿線マスタログ修正 -add sta
            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }
            // 20160915 沿線マスタログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

            if (tmp_list.Contains(value))
            {
                chkvalue = value;
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        /// <summary>
        /// 駅マスタ有無チェック
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_Eki(SqlConnection sqlcnnv10, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            // マスタからコードを取得
            var tmp_list = new List<string>();
            string tmp_sql = " SELECT eki_no FROM m_ensen_eki ";
            bool flg = true;
            bool rtn = true;

            // 20160915 沿線マスタログ修正 -add sta
            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }
            // 20160915 沿線マスタログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

            if (tmp_list.Contains(value))
            {
                chkvalue = value;
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        // 金融機関

        /// <summary>
        /// 口座振替マスタ有無チェック '2016.04.26 メインの方へも反映させる修正 -chg sta
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_KozaFkae(SqlConnection sqlcnnv10, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            // マスタからコードを取得
            var tmp_list = new List<string>();
            string tmp_sql = " SELECT fkae_no FROM m_fb_fkaejyoho ";
            bool flg = true;
            bool rtn = true;

            // 20160530 ログ修正 -add sta
            if (string.IsNullOrEmpty(value) | value == "0")
            {
                return rtn;
            }
            // 20160530 ログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

            if (tmp_list.Contains(value))
            {
                chkvalue = value;
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_RELEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        // 20160530 ログ修正 -del sta
        // ''' <summary>
        // ''' 仲介業者マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_GyCyukai(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // 'コードを取得
        // Dim tmp_list As New List(Of String)
        // Dim tmp_sql As String = " SELECT gy_fudono FROM gydata_fudo "
        // Dim flg As Boolean = True
        // Dim rtn As Boolean = True

        // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        // If tmp_list.Contains(value) Then
        // chkvalue = value
        // Else
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = ""
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20160530 ログ修正 -del end

        // 20160530 ログ修正 -del sta
        // ''' <summary>
        // ''' 施工業者マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_GySeko(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // 'コードを取得
        // Dim tmp_list As New List(Of String)
        // Dim tmp_sql As String = " SELECT gy_sekono FROM gydata_seko "
        // Dim flg As Boolean = True
        // Dim rtn As Boolean = True

        // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        // If tmp_list.Contains(value) Then
        // chkvalue = value
        // Else
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = ""
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20160530 ログ修正 -del end

        // 20160530 ログ修正 -del sta
        // ''' <summary>
        // ''' 保守業者マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_GyHosyu(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // 'コードを取得
        // Dim tmp_list As New List(Of String)
        // Dim tmp_sql As String = " SELECT gy_sisetuno FROM gydata_sisetu "
        // Dim flg As Boolean = True
        // Dim rtn As Boolean = True

        // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        // If tmp_list.Contains(value) Then
        // chkvalue = value
        // Else
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = ""
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20160530 ログ修正 -del end

        // 20160530 ログ修正 -del sta
        // ''' <summary>
        // ''' 家賃保証業者マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_GyHosyo(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // 'コードを取得
        // Dim tmp_list As New List(Of String)
        // Dim tmp_sql As String = " SELECT gy_hosyono FROM gydata_hosyo "
        // Dim flg As Boolean = True
        // Dim rtn As Boolean = True

        // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        // If tmp_list.Contains(value) Then
        // chkvalue = value
        // Else
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = ""
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20160530 ログ修正 -del end

        /// <summary>
        /// ライフライン業者マスタ有無チェック
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_GyLifeline(SqlConnection sqlcnnv10, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            // ライフライン業者の種別から該当するレコードを指定する
            string tmp_where = "";
            string lifelinekbn = erritem.Replace("ライフライン", "");

            switch (lifelinekbn ?? "")
            {
                case "(電気)":
                    {
                        tmp_where = " WHERE lifeline_denkiflg = 1 ";
                        break;
                    }
                case "(上水)":
                    {
                        tmp_where = " WHERE lifeline_josuidoflg = 1 ";
                        break;
                    }
                case "(ガス)":
                    {
                        tmp_where = " WHERE lifeline_gasflg = 1 ";
                        break;
                    }
                case "(排水)":
                    {
                        tmp_where = " WHERE lifeline_haisuiflg = 1 ";
                        break;
                    }
                case "(灯油)":
                    {
                        tmp_where = " WHERE lifeline_toyuflg = 1 ";
                        break;
                    }
                case "(その他1)":
                case "(その他2)":
                case "(その他3)":
                    {
                        tmp_where = " WHERE lifeline_other1flg = 1 ";
                        break;
                    }
            }

            // コードを取得
            var tmp_list = new List<string>();
            string tmp_sql = " SELECT gy_lifelineno FROM gydata_lifeline " + tmp_where;
            bool flg = true;
            bool rtn = true;

            // 20160530 ログ修正 -add sta
            if (string.IsNullOrEmpty(value) | value == "0")
            {
                return rtn;
            }
            // 20160530 ログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

            if (tmp_list.Contains(value))
            {
                chkvalue = value;
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        // 20160530 ログ修正 -del sta
        // ''' <summary>
        // ''' 自社マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_Jisya(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // 'コードを取得
        // Dim tmp_list As New List(Of String)
        // Dim tmp_sql As String = " SELECT jisya_no FROM jisyadata "
        // Dim flg As Boolean = True
        // Dim rtn As Boolean = True

        // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        // If tmp_list.Contains(value) Then
        // chkvalue = value
        // Else
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = ""
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20160530 ログ修正 -del end

        // 20160530 ログ修正 -del sta
        // ''' <summary>
        // ''' 家賃入金口座マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_YatinKoza(ByVal sqlcnnv10 As SqlConnection, ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // 'コードを取得
        // Dim tmp_list As New List(Of String)
        // Dim tmp_sql As String = " SELECT yatin_kozano FROM m_yatinkoza "
        // Dim flg As Boolean = True
        // Dim rtn As Boolean = True

        // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, tmp_list)

        // If tmp_list.Contains(value) Then
        // chkvalue = value
        // Else
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = ""
        // rtn = False
        // End If

        // Return rtn

        // End Function
        // 20160530 ログ修正 -del end

        /// <summary>
        /// エリアマスタ有無チェック
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_Area(SqlConnection sqlcnnv10, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            // コードを取得
            var tmp_list = new List<string>();
            string tmp_sql = " SELECT area_no FROM m_area WHERE area_name = '" + value + "'";
            bool flg = true;
            bool rtn = true;

            // 20160530 ログ修正 -add sta
            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }
            // 20160530 ログ修正 -add end

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

            if (tmp_list.Count != 0)
            {
                chkvalue = tmp_list[0];
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        // 2016.04.06 都市計画・用途地域を紐データから取得するように修正 -del sta
        // ''' <summary>
        // '''都市開発/用途地域マスタ有無チェック
        // ''' </summary>
        // ''' <param name="sqlcnnv10"></param>
        // ''' <param name="value"></param>
        // ''' <param name="chkvalue"></param>
        // ''' <param name="errstr"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Shared Function Chk_DataMstExist_TosiYoto(ByVal value As String, ByRef chkvalue As String, ByVal erritem As String, ByRef errstr As String) As Boolean

        // Dim rtn As Boolean = True

        // If value = "99" Then
        // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
        // chkvalue = "-1"
        // rtn = False
        // Else
        // chkvalue = value
        // End If

        // Return rtn

        // End Function
        // 2016.04.06 都市計画・用途地域を紐データから取得するように修正 -del end

        /// <summary>
        /// 紐付け項目有無チェック (都市計画・用途地域用) 2016.04.06 都市計画・用途地域を紐データから取得するように修正
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_RelItem_Tosiyoto(SafeDictionary<string, string> hash_rel, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            bool rtn = true;

            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }

            switch (value ?? "")
            {
                case "都市計画の重複":  // 用途地域に入れておいた文字列
                    {
                        errstr = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + "2つ目の都市計画データは用途地域へ移行できません" + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                        chkvalue = "";
                        rtn = false;
                        break;
                    }
                case "用途地域の重複":  // 都市計画に入れておいた文字列
                    {
                        errstr = CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + "2つ目の用途地域データは都市計画へ移行できません" + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                        chkvalue = "";
                        rtn = false;
                        break;
                    }

                default:
                    {
                        if (hash_rel.ContainsKey(value))
                        {
                            chkvalue = Conversions.ToString(hash_rel[value]);
                        }
                        else
                        {
                            errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                            chkvalue = "";
                            rtn = false;
                        }

                        break;
                    }
            }

            return rtn;

        }

        /// <summary>
        /// 紐付け項目有無チェック
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks>
        /// 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応
        /// 　引数にデフォルト値「defvalue」を追加
        /// </remarks>
        public static bool Chk_DataMstExist_RelItem(SafeDictionary<string, string> hash_rel, string value, ref string chkvalue, string erritem, ref string errstr, string defvalue = "")
        {

            bool rtn = true;

            // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -add sta
            if (erritem == "物件分類" || erritem == "部屋分類")
            {
                if (string.IsNullOrEmpty(value))
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED;
                    chkvalue = defvalue;
                    rtn = false;
                    return rtn;
                }
            }
            // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -add end

            // 20160530 ログ修正 -add sta
            if (string.IsNullOrEmpty(value) | value == "0")
            {
                return rtn;
            }
            // 20160530 ログ修正 -add end

            // ダミーレコードとして挿入する必要がある場合は「挿入用」として抽出しておいたデータをそのまま移行する
            // (送金ルール入金項目のその他請求等)
            // 2016.04.26 メインの方へも反映させる修正 -chg sta
            // If value = "挿入用" Then
            // chkvalue = "0"
            // Return rtn
            // End If
            if ((value ?? "") != (value.Replace("挿入用", "") ?? ""))
            {
                chkvalue = "0";
                return rtn;
            }
            // 2016.04.26 メインの方へも反映させる修正 -chg end

            if (hash_rel.ContainsKey(value))
            {
                chkvalue = Conversions.ToString(hash_rel[value]);
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        /// <summary>
        /// 紐付け項目(物件部屋鍵)有無チェック '20160530 ログ修正 -add sta
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_RelItem_kagi(string value, ref string chkvalue, string erritem, ref string errstr)
        {

            bool rtn = true;

            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }

            // 20160927 汎用CV時の鍵情報の取得処理修正 -chg sta
            // '20160613 鍵情報の取得処理修正 -chg sta
            // 'Select Case erritem
            // '    Case "共用鍵タイトル"
            // '        If Hash_KagiTitleKyoyo.Contains(value) Then
            // '            chkvalue = Hash_KagiTitleKyoyo.Item(value)
            // '        ElseIf Hash_KagiTitleSenyo.Contains(value) Then
            // '            '20160603 ユーザーデータ検証による修正 -chg sta
            // '            'errstr = LOG_NAIYO_ERR_SENYOKAGI & "-" & LOG_HUBI_SENYOKAGI & "-" & LOG_TAISYO_SENYOKAGI
            // '            errstr = ""
            // '            '20160603 ユーザーデータ検証による修正 -chg end
            // '        Else    '対象データが共用鍵かつ専用鍵タイトルにデータが含まれていない場合はエラーログ出力
            // '            errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            // '            chkvalue = ""
            // '            rtn = False
            // '        End If
            // '    Case "専用鍵タイトル"
            // '        If Hash_KagiTitleSenyo.Contains(value) Then
            // '            chkvalue = Hash_KagiTitleSenyo.Item(value)
            // '        ElseIf Hash_KagiTitleKyoyo.Contains(value) Then
            // '            '20160603 ユーザーデータ検証による修正 -chg sta
            // '            'errstr = LOG_NAIYO_ERR_KYOYOKAGI & "-" & LOG_HUBI_KYOYOKAGI & "-" & LOG_TAISYO_KYOYOKAGI
            // '            errstr = ""
            // '            '20160603 ユーザーデータ検証による修正 -chg end
            // '        Else    '対象データが専用鍵かつ共用鍵タイトルにデータが含まれていない場合はエラーログ出力
            // '            errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            // '            chkvalue = ""
            // '            rtn = False
            // '        End If
            // 'End Select

            // If Hash_KagiTitle_V7to10.Contains(value) Then

            // Dim tmp_value As String = Hash_KagiTitle_V7to10.Item(value)
            // Dim tmp_str() As String = tmp_value.Split("-")
            // Dim tmp_kbn As String = tmp_str(0)
            // Dim tmp_kagino As String = tmp_str(1)

            // Select Case erritem

            // Case "共用鍵タイトル"
            // If tmp_kbn = "1" Then
            // chkvalue = tmp_kagino
            // Return rtn
            // End If
            // Case "専用鍵タイトル"
            // If tmp_kbn = "2" Then
            // chkvalue = tmp_kagino
            // Return rtn
            // End If
            // End Select

            // End If
            // '20160613 鍵情報の取得処理修正 -chg end
            switch (CommonModule.CNVNO)
            {

                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        break;
                    }
                // 20161028 物件/部屋鍵取得方法修正 -del sta
                // 汎用CVの場合はあらかじめデフォルト「鍵1」～「鍵50」を設定しているため鍵タイトルの有無の確認は行わないようにする
                // Dim list_kagino As New List(Of String)

                // Select Case erritem
                // Case "共用鍵タイトル"
                // Dim list_bkkagino As New List(Of String) From {"1", "2", "3"}
                // list_kagino = list_bkkagino
                // Case "専用鍵タイトル"
                // Dim list_hykagino As New List(Of String) From {"1", "2", "3", "4", "5", "6"}
                // list_kagino = list_hykagino
                // End Select

                // 'マスタと照合
                // If list_kagino.Contains(value) Then
                // chkvalue = value
                // Else
                // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
                // chkvalue = ""
                // rtn = False
                // End If
                // 20161028 物件/部屋鍵取得方法修正 -del end
            }
            // 20160927 汎用CV時の鍵情報の取得処理修正 -chg end

            return rtn;

        }

        /// <summary>
        /// 紐付け項目(設備)有無チェック '20160530 ログ修正
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_RelItem_setubi(SafeDictionary<string, string> hash_rel, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            bool rtn = true;

            if (string.IsNullOrEmpty(value) | value == "0")
            {
                return rtn;
            }

            // 20160913_2 部屋設備移行処理の追加 -chg sta
            // 設備で紐付けられていないデータはOKとしログを出力しない
            // If hash_rel.Contains(value) Then
            // chkvalue = hash_rel.Item(value)
            // Else
            // 'errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            // 'chkvalue = ""
            // 'rtn = False
            // End If
            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        chkvalue = value;
                        break;
                    }
            }
            // 20160913_2 部屋設備移行処理の追加 -chg end

            return rtn;

        }

        /// <summary>
        /// 間取文字列有無チェック
        /// </summary>
        /// <param name="chk_m"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_MadoriStr(string value, ref string chkvalue, string erritem, ref string errstr)
        {

            bool rtn = true;

            if (!string.IsNullOrEmpty(value))
            {

                switch (value ?? "")
                {
                    case "R":
                        {
                            chkvalue = 1.ToString();
                            break;
                        }
                    case "K":
                        {
                            chkvalue = 2.ToString();
                            break;
                        }
                    case "DK":
                        {
                            chkvalue = 3.ToString();
                            break;
                        }
                    case "LDK":
                        {
                            chkvalue = 4.ToString();
                            break;
                        }
                    case "SK":
                        {
                            chkvalue = 5.ToString();
                            break;
                        }
                    case "SDK":
                        {
                            chkvalue = 6.ToString();
                            break;
                        }
                    case "SLDK":
                        {
                            chkvalue = 7.ToString();
                            break;
                        }
                    case "LK":
                        {
                            chkvalue = 8.ToString();
                            break;
                        }
                    case "SLK":
                        {
                            chkvalue = 9.ToString();
                            break;
                        }
                    case "SR":
                        {
                            chkvalue = 10.ToString();
                            break;
                        }

                    default:
                        {
                            errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                            chkvalue = (-1).ToString();
                            rtn = false;
                            break;
                        }
                }

            }

            return rtn;

        }

        /// <summary>
        /// 間取文字列有無チェック(内訳)
        /// </summary>
        /// <param name="chk_m"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_MadoriUtiwakeStr(string value, ref string chkvalue, string erritem, ref string errstr)
        {

            bool rtn = true;

            // 20160915 間取内訳情報の汎用/既存処理の分岐 -chg sta
            // If value <> "" Then

            // Select Case value
            // Case "和"
            // chkvalue = 1
            // Case "洋"
            // chkvalue = 2
            // Case "K"
            // chkvalue = 8
            // Case "L"
            // chkvalue = 9
            // Case "DK"
            // chkvalue = 10
            // Case "LD"
            // chkvalue = 11
            // Case "LDK"
            // chkvalue = 12
            // Case "ロフト"
            // chkvalue = 3
            // Case "S"
            // chkvalue = 4
            // Case "書斎"
            // chkvalue = 5
            // Case "サンルーム"
            // chkvalue = 6
            // Case "グルニエ"
            // chkvalue = 7
            // Case Else
            // '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg sta
            // 'errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            // If value.Replace("(対応不要)", "") <> value Then
            // errstr = LOG_NAIYO_ERR_TAIONASI & "-" & LOG_HUBI_TAIONASI & "-" & LOG_TAISYO_TAIONASI
            // Else
            // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            // End If
            // '20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg end
            // chkvalue = ""
            // rtn = False
            // End Select

            // End If

            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }

            switch (CommonModule.CNVNO)
            {
                case (int)CommonModule.ConvertTypes._汎用:
                    {
                        chkvalue = value;
                        break;
                    }
            }
            // 20160915 間取内訳情報の汎用/既存処理の分岐 -chg end
            return rtn;

        }

        /// <summary>
        /// 契約分類マスタ有無チェック
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_Kybrui(SqlConnection sqlcnnv10, string value, ref string chkvalue, string erritem, ref string errstr)
        {

            // コードを取得
            var tmp_list = new List<string>();
            string tmp_sql = " SELECT ky_ruino FROM m_ky_rui ";
            bool flg = true;
            bool rtn = true;

            flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

            if (tmp_list.Contains(value))
            {
                chkvalue = value;
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        /// <summary>
        /// 参照ファイル有無チェック
        /// </summary>
        /// <param name="chk_m"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataExist_FilePath(string value, ref string chkvalue, ref string errstr)
        {

            bool rtn = true;

            // ファイル有無確認
            if (File.Exists(value))
            {
                chkvalue = value;
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTFILEDATA + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_FILEEXIST + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        /// <summary>
        /// 請求発生詳細情報を設定する '2016.04.26 メインの方へも反映させる修正
        /// </summary>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="chkkbn"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_Data_SqHasseiItem(string value, ref string chkvalue, string chkkbn, ref string errstr)
        {

            bool rtn = true;

            // 連結しているデータを分割
            string[] tmp_str = value.Split('-');

            // 各変数へ格納
            string tmp_cvitem = tmp_str[0];       // 移行対象項目
            string tmp_sqtuki = tmp_str[1];       // 請求月 (前月、当月等)
            string tmp_sqkaisiymd = tmp_str[2];   // 請求開始年月
            string tmp_sqkankaku = tmp_str[3];    // 請求間隔

            // 返却用変数
            string sqptn = "1";       // デフォルト値を設定しておく
            string sqinterval = "1";  // デフォルト値を設定しておく
            string sqnenyear = "";
            string sqtukiyear = "";

            switch (tmp_cvitem ?? "")
            {
                case "部屋入金項目情報":
                    {
                        // -------------------------
                        // 請求間隔の判別
                        // -------------------------
                        if (string.IsNullOrEmpty(tmp_sqkankaku) || tmp_sqkankaku == "0" || tmp_sqkankaku == "1")
                        {
                            return rtn;
                        }
                        else
                        {
                            rtn = false;
                            return rtn;
                        }
                    }

                case "契約入金項目情報":
                case "家主固定控除情報":
                    {
                        // -------------------------
                        // 各設定値のチェック (設定値が全て揃っていないと移行できないため)
                        // -------------------------
                        if (string.IsNullOrEmpty(tmp_sqtuki) || string.IsNullOrEmpty(tmp_sqkaisiymd) || string.IsNullOrEmpty(tmp_sqkankaku))
                        {
                            return rtn;
                        }

                        // -------------------------
                        // 請求月の取得
                        // -------------------------
                        int tmp_chgsqtuki = 0;    // 請求月
                        int tmp_addtuki = 0;
                        if (string.IsNullOrEmpty(tmp_sqtuki) || int.TryParse(tmp_sqtuki, out tmp_chgsqtuki) == false)
                        {
                            rtn = false;
                            return rtn;
                        }
                        else
                        {
                            // 変換
                            switch (tmp_chgsqtuki)
                            {
                                case 1:
                                    {
                                        tmp_addtuki = -2;
                                        break;
                                    }
                                case 2:
                                    {
                                        tmp_addtuki = -1;
                                        break;
                                    }
                                case 3:
                                    {
                                        tmp_addtuki = 0;
                                        break;
                                    }
                                case 4:
                                    {
                                        tmp_addtuki = 1;
                                        break;
                                    }
                                case 5:
                                    {
                                        tmp_addtuki = 2;
                                        break;
                                    }
                            }
                        }

                        // -------------------------
                        // 請求開始年月の月部分を取得
                        // -------------------------
                        DateTime tmp_chgsqkaisiymd;
                        int tmp_sqkaisituki = 0;  // 請求開始月
                        if (string.IsNullOrEmpty(tmp_sqkaisiymd) || DateTime.TryParse(tmp_sqkaisiymd, out tmp_chgsqkaisiymd) == false)
                        {
                            rtn = false;
                            return rtn;
                        }
                        else
                        {
                            tmp_sqkaisituki = int.Parse(tmp_chgsqkaisiymd.Month.ToString());
                        }

                        // -------------------------
                        // 請求開始月に対する請求月の取得 (請求間隔が2ヶ月の場合の奇数偶数請求月設定用)
                        // -------------------------
                        int tmp_sqtaisyotuki = tmp_sqkaisituki + tmp_addtuki;     // 請求月を取得 (該当月に対する請求月)
                        if (tmp_sqtaisyotuki <= 0)
                        {
                            tmp_sqtaisyotuki = tmp_sqtaisyotuki + 12;    // 負の請求月の調整 (前月請求で1月の場合は12月に調整)
                        }

                        // -------------------------
                        // 請求間隔判別
                        // -------------------------
                        int tmp_chgsqkankaku = 0;

                        // 空または数値以外の場合は処理を抜ける
                        if (string.IsNullOrEmpty(tmp_sqkankaku) || int.TryParse(tmp_sqkankaku, out tmp_chgsqkankaku) == false)
                        {
                            rtn = false;
                            return rtn;
                        }

                        // 請求間隔が13ヶ月以上または12を割った場合に余りが生じる場合は処理を抜ける
                        if (tmp_chgsqkankaku >= 13 || 12 % tmp_chgsqkankaku != 0)
                        {
                            rtn = false;
                            return rtn;
                        }

                        // 請求間隔による処理の分岐
                        if (tmp_chgsqkankaku == 2)
                        {
                            // 2ヶ月の場合は奇数月か偶数月か判別して処理を抜ける
                            if (tmp_sqtaisyotuki % 2 != 0)
                            {
                                // 奇数月請求
                                sqinterval = "2";
                            }
                            else if (tmp_sqtaisyotuki % 2 == 0)
                            {
                                // 偶数月請求
                                sqinterval = "3";
                            }
                        }
                        else if (tmp_chgsqkankaku >= 3)
                        {
                            // 3ヶ月以上の場合は詳細な設定にチェック
                            sqptn = "2";

                        }

                        // 年あたりの請求回数を取得
                        int tmp_sqcnt = (int)Math.Round(12d / tmp_chgsqkankaku);
                        var tmp_listsqtukigrp = new List<int>();
                        for (int cntii = 1, loopTo = tmp_sqcnt; cntii <= loopTo; cntii++)
                        {
                            if (tmp_chgsqkankaku == 1)
                            {
                                tmp_listsqtukigrp.Add(cntii);
                            }
                            else
                            {
                                tmp_listsqtukigrp.Add((tmp_chgsqkankaku * (cntii - 1) + tmp_sqkaisituki) % 12);
                            }
                        }

                        Set_Sqnen(tmp_listsqtukigrp, tmp_sqkaisituki, tmp_addtuki, ref sqnenyear);
                        Set_Sqtuki(tmp_listsqtukigrp, tmp_sqkaisituki, tmp_addtuki, ref sqtukiyear);
                        break;
                    }

            }

            // チェック項目で返す値を設定
            switch (chkkbn ?? "")
            {
                case "請求パターン":
                    {
                        chkvalue = sqptn;
                        break;
                    }
                case "請求間隔":
                    {
                        chkvalue = sqinterval;
                        break;
                    }
                case "請求発生年":
                    {
                        chkvalue = sqnenyear;
                        break;
                    }
                case "請求発生月":
                    {
                        chkvalue = sqtukiyear;
                        break;
                    }
            }

            return rtn;

        }

        /// <summary>
        /// 請求発生年の設定
        /// </summary>
        /// <param name="sqtukigrp"></param>
        /// <param name="sqkaisituki"></param>
        /// <param name="addtuki"></param>
        /// <param name="sqnenyear"></param>
        /// <remarks></remarks>
        public static void Set_Sqnen(object sqtukigrp, int sqkaisituki, int addtuki, ref string sqnenyear)
        {

            // 請求年の設定
            var tmp_sqnenyear = new string[13];
            foreach (var sqtuki in (IEnumerable)sqtukigrp)
            {

                for (int cntii = 1; cntii <= 12; cntii++)

                    // 20160603 ユーザーデータ検証による修正 -chg sta
                    // 全ての該当月に対して「当年」を設定するように修正
                    // If cntii < sqkaisituki Then
                    // tmp_sqnenyear(cntii) = "3"
                    // Else
                    // tmp_sqnenyear(cntii) = "2"
                    // End If
                    // 20160603 ユーザーデータ検証による修正 -chg end

                    tmp_sqnenyear[cntii] = "2";

            }

            // 移行用に成形
            string tmp_strsqnen = "";
            string tmp_strsqtuki = "";

            for (int cntii = 1; cntii <= 12; cntii++)
                tmp_strsqnen = tmp_strsqnen + "," + tmp_sqnenyear[cntii];

            sqnenyear = tmp_strsqnen.Remove(0, 1);

        }

        /// <summary>
        /// 請求発生月の設定
        /// </summary>
        /// <param name="sqtukigrp"></param>
        /// <param name="sqkaisituki"></param>
        /// <param name="addtuki"></param>
        /// <param name="sqtukiyear"></param>
        /// <remarks></remarks>
        public static void Set_Sqtuki(object sqtukigrp, int sqkaisituki, int addtuki, ref string sqtukiyear)
        {

            // 請求月の設定
            var tmp_sqnentuki = new string[13];
            foreach (var sqtuki in (IEnumerable)sqtukigrp)
            {

                int tmp_sqtaisyotuki = Conversions.ToInteger(Operators.AddObject(sqtuki, addtuki));     // 請求月を取得 (該当月に対する請求月)
                if (tmp_sqtaisyotuki <= 0)
                {
                    tmp_sqtaisyotuki = tmp_sqtaisyotuki + 12;    // 負の請求月の調整 (前月請求で1月の場合は12月に調整)
                }

                tmp_sqnentuki[Conversions.ToInteger(sqtuki)] = tmp_sqtaisyotuki.ToString();

            }

            for (int cntii = 1; cntii <= 12; cntii++)
            {

                if (tmp_sqnentuki[cntii] is null)
                {
                    tmp_sqnentuki[cntii] = "13";
                }

            }

            // 移行用に成形
            string tmp_strsqnen = "";
            string tmp_strsqtuki = "";

            for (int cntii = 1; cntii <= 12; cntii++)
                tmp_strsqtuki = tmp_strsqtuki + "," + tmp_sqnentuki[cntii];

            sqtukiyear = tmp_strsqtuki.Remove(0, 1);

        }

        /// <summary>
        /// 自社口座マスタ有無チェック
        /// 引数 sqlcnnv10 を追加 '20160928 自社口座取得方法の修正
        /// </summary>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="chkkbn"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_Data_FBKoza(SqlConnection sqlcnnv10, string value, ref string chkvalue, string chkkbn, ref string errstr)
        {

            bool rtn = true;
            string jisyano = "";
            string jisyakozano = "";
            string erritem = "自社口座";

            // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg sta
            // '20160530 ログ修正 -add sta
            // If value = "" Then
            // Return rtn
            // End If
            // '20160530 ログ修正 -add end

            // If Hash_Rel_JisyaKoza.Contains(value) = False Then
            // errstr = LOG_NAIYO_ERR_NOTEXISTMSTDATA & "-" & erritem & LOG_HUBI_NOTEXISTDATA_MSTEXIST & "-" & LOG_TAISYO_NOTEXISTDATA_TODOFUKEN
            // chkvalue = ""
            // rtn = False
            // Return rtn
            // Else
            // Dim tmp_jisyainfo As String = Hash_Rel_JisyaKoza.Item(value)
            // Dim tmp_item() As String = tmp_jisyainfo.Split("-")
            // jisyano = tmp_item(0)
            // jisyakozano = tmp_item(1)
            // End If

            // Select Case chkkbn
            // Case "自社_口座用"
            // chkvalue = jisyano
            // Case "自社口座_口座用"
            // chkvalue = jisyakozano
            // End Select

            // 空データチェック
            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }

            // 汎用/既存で処理を分岐
            switch (CommonModule.CNVNO)
            {


                case (int)CommonModule.ConvertTypes._汎用:
                    {

                        // 汎用は直接設定された値から取得
                        // 20160928 自社口座取得方法の修正 -chg sta
                        // chkvalue = value

                        // データが存在しない場合は処理を抜ける
                        if (value == "-")
                        {
                            return rtn;
                        }

                        // 自社No、自社口座No分割
                        string[] tmp_str = value.Split('-');
                        string tmp_jisyano = tmp_str[0];
                        string tmp_jisyakozano = tmp_str[1];
                        string chkstrafter = "";
                        string tmp_sql = "";
                        var tmp_list = new List<string>();
                        string taisyovalu = "";

                        switch (chkkbn ?? "")
                        {
                            case "自社_口座用":
                                {
                                    // 自社Noが無い場合は処理を抜ける
                                    if (string.IsNullOrEmpty(tmp_jisyano))
                                    {
                                        chkvalue = "";
                                        return rtn;
                                    }
                                    else if (Chk_DataNumeric(tmp_jisyano, ref chkstrafter, "1", "999999999", "", ref errstr) == false)
                                    {
                                        chkvalue = "";
                                        rtn = false;
                                        return rtn;
                                    }
                                    tmp_sql = " SELECT jisya_no FROM jisyadata ";
                                    taisyovalu = tmp_jisyano;

                                    DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

                                    if (tmp_list.Contains(taisyovalu))
                                    {
                                        chkvalue = taisyovalu;
                                    }
                                    else
                                    {
                                        errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                                        chkvalue = "";
                                        rtn = false;
                                    }

                                    break;
                                }

                            case "自社口座_口座用":
                                {
                                    // 自社口座Noチェック
                                    if (string.IsNullOrEmpty(tmp_jisyakozano))
                                    {
                                        chkvalue = "";
                                        return rtn;
                                    }
                                    else if (Chk_DataNumeric(tmp_jisyakozano, ref chkstrafter, "1", "999", "", ref errstr) == false)
                                    {
                                        chkvalue = "";
                                        rtn = false;
                                        return rtn;
                                    }
                                    tmp_sql = " SELECT CONVERT(VARCHAR,jisya_no) + '-' + CONVERT(VARCHAR,jisya_kozano) FROM jisyadata_koza ";
                                    taisyovalu = value;

                                    DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

                                    if (tmp_list.Contains(taisyovalu))
                                    {
                                        chkvalue = tmp_jisyakozano;
                                    }
                                    else
                                    {
                                        errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                                        chkvalue = "";
                                        rtn = false;
                                    }

                                    break;
                                }

                        }

                        break;
                    }
                    // 20160928 自社口座取得方法の修正 -chg end
            }

            // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 -chg end
            return rtn;

        }

        /// <summary>
        /// 家賃入金口座に紐付く家主口座マスタ有無チェック 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
        /// </summary>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="chkkbn"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_Data_YatinOwKoza(SqlConnection sqlcnnv10, string value, ref string chkvalue, string chkkbn, ref string errstr)
        {

            bool rtn = true;

            // 空データチェック
            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }

            // 汎用/既存で処理を分岐
            switch (CommonModule.CNVNO)
            {

                case (int)CommonModule.ConvertTypes._汎用:
                    {

                        // データが存在しない場合は処理を抜ける([ow_no]-[ow_kozano]の形)
                        if (value == "-")
                        {
                            return rtn;
                        }

                        // 家主No、家主口座No分割
                        string[] tmp_str = value.Split('-');
                        string tmp_owno = tmp_str[0];
                        string tmp_owkozano = tmp_str[1];
                        string chkstrafter = "";
                        string tmp_sql = "";
                        var tmp_list = new List<string>();
                        string taisyovalu = "";

                        switch (chkkbn ?? "")
                        {
                            case "家主_口座用":
                                {
                                    // 家主Noが無い場合は処理を抜ける
                                    if (string.IsNullOrEmpty(tmp_owno))
                                    {
                                        chkvalue = "";
                                        return rtn;
                                    }
                                    else if (Chk_DataNumeric(tmp_owno, ref chkstrafter, "1", "999999999", "", ref errstr) == false)
                                    {
                                        chkvalue = "";
                                        rtn = false;
                                        return rtn;
                                    }

                                    tmp_sql = " SELECT ow_no FROM owdata ";
                                    taisyovalu = tmp_owno;

                                    DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

                                    if (tmp_list.Contains(taisyovalu))
                                    {
                                        chkvalue = taisyovalu;
                                    }
                                    else
                                    {
                                        errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + "家主" + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                                        chkvalue = "";
                                        rtn = false;
                                    }

                                    break;
                                }

                            case "家主口座_口座用":
                                {
                                    // 家主口座Noチェック
                                    if (string.IsNullOrEmpty(tmp_owkozano))
                                    {
                                        chkvalue = "";
                                        return rtn;
                                    }
                                    else if (Chk_DataNumeric(tmp_owkozano, ref chkstrafter, "1", "999999999", "", ref errstr) == false)
                                    {
                                        chkvalue = "";
                                        rtn = false;
                                        return rtn;
                                    }
                                    tmp_sql = " SELECT CONVERT(VARCHAR,ow_no) + '-' + CONVERT(VARCHAR,ow_kozano) FROM owdata_koza ";
                                    taisyovalu = value;

                                    DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref tmp_list);

                                    if (tmp_list.Contains(taisyovalu))
                                    {
                                        chkvalue = tmp_owkozano;
                                    }
                                    else
                                    {
                                        errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + "家主口座" + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                                        chkvalue = "";
                                        rtn = false;
                                    }

                                    break;
                                }

                        }

                        break;
                    }

            }

            return rtn;

        }





        /// <summary>
        /// 画像判別 '20160525 クレーム関連ファイルの画像判別処理実装
        /// </summary>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_PictureFile(string value, ref string chkvalue, ref string errstr)
        {

            bool rtn = true;

            // 20160627 クレーム関連ファイル移行修正 -chg sta
            // 画像は移行せずファイルパスをそのまま移行するように修正するためコメントアウト
            // For Each pickaku In List_ImportableImageFileAttributes
            // If Strings.StrConv(Path.GetExtension(value), VbStrConv.Narrow Or VbStrConv.Lowercase, 0) = pickaku Then
            // errstr = LOG_NAIYO_ERR_RELFILEPIC & "-" & LOG_HUBI_RELFILEPIC & "-" & LOG_TAISYO_RELFILEPIC
            // chkvalue = ""
            // rtn = False
            // Exit For
            // Else
            // chkvalue = value
            // End If
            // Next
            chkvalue = value;
            // 20160627 クレーム関連ファイル移行修正 -chg end

            return rtn;

        }

        /// <summary>
        /// 紐付け項目(FB関連)有無チェック '20160609 紐付設定値取得に伴う修正
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="errstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataMstExist_RelItem_FB(string value, ref string chkvalue, string erritem, ref string errstr)
        {

            bool rtn = true;

            if (string.IsNullOrEmpty(value))
            {
                return rtn;
            }

            Dictionary<string, string> obj = null;

            switch (erritem ?? "")
            {
                case "FB関連_口座振替":
                    {
                        obj = CommonModule.Hash_Rel_FBInfo_KozaFurikaeFmt;
                        break;
                    }
                case "FB関連_総合振込":
                    {
                        obj = CommonModule.Hash_Rel_FBInfo_FuriIraiFmt;
                        break;
                    }
                case "FB関連_入出金フォーマット":
                    {
                        obj = CommonModule.Hash_Rel_FBInfo_NsSettingFmt;
                        break;
                    }
                case "FB関連_入出金自社No":
                    {
                        obj = CommonModule.Hash_Rel_FBInfo_NsSettingJisya;
                        break;
                    }
            }

            if (Conversions.ToBoolean(obj.ContainsKey(value)))
            {
                chkvalue = Conversions.ToString(obj[value]);
            }
            else
            {
                errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTMSTDATA + "-" + erritem + CommonModule.LOG_HUBI_NOTEXISTDATA_MSTEXIST + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_TODOFUKEN;
                chkvalue = "";
                rtn = false;
            }

            return rtn;

        }

        /// <summary>
        /// 全体データチェック→チェックメソッド呼出
        /// </summary>
        /// <param name="tblfldname"></param>
        /// <param name="value"></param>
        /// <param name="errstr"></param>
        /// <remarks></remarks>
        public static void Call_ChkMethodMidTotal(string tblfldname, string value, ref string errstr)
        {

            string tmp_type = Conversions.ToString(CommonModule.Hash_FiledTypeJp[tblfldname]);
            var min = new object();
            var max = new object();
            string datemin = "";
            string datemax = "";

            string chkaftervalue = "";
            string defvalue = "";
            bool normalflg = true;                     // kakaka 以下の条件を表など一覧にして仕様書に記載しておく

            // 2016.03.28 型落ちを考慮した処理へ修正 -chg sta
            // Dim tmp_size As Integer = Int32.Parse(Hash_FiledSizeJp(tblfldname))         'kakaka4 0322_1300 変換時の型落ちを考慮。
            int tmp_size = 0;
            if (int.TryParse(Conversions.ToString(CommonModule.Hash_FiledSizeJp[tblfldname]), out tmp_size) == false)
            {
                errstr = CommonModule.LOG_NAIYO_ERR_DATACHK + "-" + CommonModule.LOG_HUBI_ERR_DATACHK_B + "-" + CommonModule.LOG_TAISYO_DATACHK_B;
                normalflg = false;
                return;
            }
            // 2016.03.28 型落ちを考慮した処理へ修正 -chg end

            switch (tmp_type ?? "")                                // kakaka 要確認：資料を見てチェックする(金丸)
            {
                case "int":
                case "smallint":
                case "bigint":
                case "tinyint":
                    {
                        Chg_SizeToValue(tmp_size, ref min, ref max);
                        normalflg = Chk_DataNumeric(value, ref chkaftervalue, min.ToString(), max.ToString(), defvalue, ref errstr);
                        break;
                    }
                case "bit":
                    {
                        normalflg = Chk_DataBit(value, ref chkaftervalue, defvalue, ref errstr);
                        break;
                    }
                case "char":
                case "varchar":
                case "text":
                    {
                        normalflg = Chk_DataString(value, ref chkaftervalue, tmp_size, 0, ref errstr);
                        break;
                    }
                case "nchar":
                case "nvarchar":
                case "ntext":
                case "sysname":
                    {
                        normalflg = Chk_DataString(value, ref chkaftervalue, tmp_size, 1, ref errstr);
                        break;
                    }
                case "datetime":
                case "smalldatetime":
                case "date":
                case "datetime2":
                case "datetimeoffset":
                    {
                        normalflg = Chk_DataDate(value, ref chkaftervalue, datemin, datemax, defvalue, ref errstr);
                        break;
                    }
                case "time":                                     // kakaka 注意：未実装
                    {
                        normalflg = Chk_DataTime(value, ref chkaftervalue, ref errstr);
                        break;
                    }
                case "money":
                    {
                        Chg_SizeToValue(tmp_size, ref min, ref max);
                        normalflg = Chk_DataMoney(value, ref chkaftervalue, min.ToString(), max.ToString(), defvalue, ref errstr);
                        break;
                    }
                case "decimal":
                case "numeric":
                    {
                        Chg_SizeToValue(tmp_size, ref min, ref max);
                        normalflg = Chk_DataDecimal(value, ref chkaftervalue, min.ToString(), max.ToString(), defvalue, ref errstr);
                        break;
                    }
                case "float":
                case "real":
                    {
                        normalflg = Chk_DataDouble(value, ref chkaftervalue, min.ToString(), max.ToString(), defvalue, ref errstr);
                        break;
                    }
                case "varbinary":
                    {
                        break;
                    }

                case "xml":
                case "uniqueidentifier":
                    {
                        break;
                    }
                    // プログラム内で生成しているためチェック不要
            }

            // hash_err.Add(tblfldname, )



        }

        /// <summary>
        /// 個別データチェック→チェックメソッド呼出
        /// </summary>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="type"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <param name="defvalue"></param>
        /// <param name="normalflg"></param>
        /// <param name="errstr"></param>
        /// <remarks></remarks>
        public static void Call_ChkMethodPerItem(string value, ref string chkvalue, string @type, string min, string max, string defvalue, ref bool normalflg, ref string errstr)
        {

            // 初期化                            'kakaka メモ：型チェック(ソースレビューでは全て必要なチェックがあることを前提とするので、確認しません)
            chkvalue = "";
            errstr = "";
            normalflg = true;

            switch (type ?? "")
            {
                case "int":
                case "smallint":
                case "bigint":
                case "tinyint":
                    {
                        normalflg = Chk_DataNumeric(value, ref chkvalue, min, max, defvalue, ref errstr);
                        break;
                    }
                case "bit":
                    {
                        normalflg = Chk_DataBit(value, ref chkvalue, defvalue, ref errstr);
                        break;
                    }
                case "char":
                case "varchar":
                case "text":
                    {
                        normalflg = Chk_DataString(value, ref chkvalue, Conversions.ToInteger(max), 0, ref errstr);
                        break;
                    }
                case "nchar":
                case "nvarchar":
                case "ntext":
                case "sysname":
                    {
                        normalflg = Chk_DataString(value, ref chkvalue, Conversions.ToInteger(max), 1, ref errstr);
                        break;
                    }
                case "datetime":
                case "smalldatetime":
                case "date":
                case "datetime2":
                case "datetimeoffset":
                    {
                        normalflg = Chk_DataDate(value, ref chkvalue, min, max, defvalue, ref errstr);
                        break;
                    }
                case "time":
                    {
                        normalflg = Chk_DataTime(value, ref chkvalue, ref errstr);
                        break;
                    }
                case "money":
                    {
                        normalflg = Chk_DataMoney(value, ref chkvalue, min, max, defvalue, ref errstr);
                        break;
                    }
                case "decimal":
                case "numeric":
                    {
                        normalflg = Chk_DataDecimal(value, ref chkvalue, min, max, defvalue, ref errstr);
                        break;
                    }
                case "float":
                case "real":
                    {
                        normalflg = Chk_DataDouble(value, ref chkvalue, min, max, defvalue, ref errstr);
                        break;
                    }
                case "varbinary":
                    {
                        break;
                    }

                case "xml":
                case "uniqueidentifier":
                case "任意文字":  // 20160927 ログの内容が不正になっているため修正 "任意文字"を追加
                    {
                        // プログラム内で生成しているためチェックせずそのまま返却
                        chkvalue = value;
                        break;
                    }

            }

        }

        /// <summary>
        /// 有無参照データチェックメソッド呼出
        /// '20160829 メールアドレス、URLの正規化処理を追加 Optionalで引数を追加
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="value"></param>
        /// <param name="chkvalue"></param>
        /// <param name="chkkbn"></param>
        /// <param name="normalflg"></param>
        /// <param name="errstr"></param>
        /// <remarks>
        /// 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応
        /// 　引数にデフォルト値「defvalue」を追加
        /// </remarks>
        public static void Call_ChkMethodPerItem_MstExist(SqlConnection sqlcnnv10, string value, ref string chkvalue, string chkkbn, ref bool normalflg, ref string errstr, string maxsize = "", string defvalue = "")
        {

            // 初期化
            chkvalue = "";
            errstr = "";
            normalflg = true;
            // 20160530 ログ修正 -add
            string tmp_sql = "";

            switch (chkkbn ?? "")
            {
                case "都道府県":
                    {
                        // 20160530 ログ修正 -chg sta
                        // normalflg = Chk_DataMstExist_Todofuken(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        tmp_sql = " SELECT ken_no FROM m_ken ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -chg end
                case "市区町村":
                    {
                        // 20160530 ログ修正 -chg sta
                        // normalflg = Chk_DataMstExist_Sikucyoson(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        tmp_sql = " SELECT si_no FROM m_si ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -chg end
                case "金融機関":
                    {
                        // 構築中
                        chkvalue = value;
                        break;
                    }
                case "金融機関支店":
                    {
                        // 構築中
                        chkvalue = value;
                        break;
                    }
                case "仲介業者":
                    {
                        // 20160530 ログ修正 -chg sta
                        // normalflg = Chk_DataMstExist_GyCyukai(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        tmp_sql = " SELECT gy_fudono FROM gydata_fudo ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -chg end
                case "ライフライン(電気)":
                case "ライフライン(上水)":
                case "ライフライン(ガス)":
                case "ライフライン(排水)":
                case "ライフライン(灯油)":
                case "ライフライン(その他1)":
                case "ライフライン(その他2)":
                case "ライフライン(その他3)":
                    {
                        normalflg = Chk_DataMstExist_GyLifeline(sqlcnnv10, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "家賃入金口座":
                    {
                        // 20160530 ログ修正 -chg sta
                        // normalflg = Chk_DataMstExist_YatinKoza(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        tmp_sql = " SELECT yatin_kozano FROM m_yatinkoza ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -chg end
                case "施工業者":
                    {
                        // 20160530 ログ修正 -chg sta
                        // normalflg = Chk_DataMstExist_GySeko(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        tmp_sql = " SELECT gy_sekono FROM gydata_seko ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -chg end
                case "保守業者":
                    {
                        // 20160530 ログ修正 -chg sta
                        // normalflg = Chk_DataMstExist_GyHosyu(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        tmp_sql = " SELECT gy_sisetuno FROM gydata_sisetu ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -chg end
                case "家賃保証業者":
                    {
                        // 20160530 ログ修正 -chg sta
                        // normalflg = Chk_DataMstExist_GyHosyo(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        tmp_sql = " SELECT gy_hosyono FROM gydata_hosyo ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -chg end

                // 20160530 ログ修正 -add sta
                case "修繕業者":
                    {
                        tmp_sql = " SELECT gy_szenno FROM gydata_szen ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -add end
                case "自社":
                    {
                        // 20160530 ログ修正 -chg sta
                        // normalflg = Chk_DataMstExist_Jisya(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        tmp_sql = " SELECT jisya_no FROM jisyadata ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160530 ログ修正 -chg end
                case "自社口座":
                    {
                        // 構築中
                        chkvalue = value;
                        break;
                    }
                case "家主":
                    {
                        tmp_sql = " SELECT ow_no FROM owdata ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "家主口座":
                    {
                        // 構築中
                        chkvalue = value;
                        break;
                    }
                case "契約者":
                    {
                        // 構築中
                        chkvalue = value;
                        break;
                    }
                // 2016.04.06 都市計画・用途地域を紐データから取得するように修正 -chg sta
                // Case "都市計画・用途地域"
                // normalflg = Chk_DataMstExist_TosiYoto(value, chkvalue, chkkbn, errstr)
                case "都市計画":
                case "用途地域":
                    {
                        normalflg = Chk_DataMstExist_RelItem_Tosiyoto(CommonModule.Hash_Rel_Tosiyotono, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 2016.04.06 都市計画・用途地域を紐データから取得するように修正 -chg end
                case "エリア":
                    {
                        normalflg = Chk_DataMstExist_Area(sqlcnnv10, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "沿線":
                    {
                        normalflg = Chk_DataMstExist_Ensen(sqlcnnv10, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "駅":
                    {
                        normalflg = Chk_DataMstExist_Eki(sqlcnnv10, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "間取":
                    {
                        normalflg = Chk_DataMstExist_MadoriStr(value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "間取内訳":
                    {
                        normalflg = Chk_DataMstExist_MadoriUtiwakeStr(value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "契約分類":
                    {
                        // 2016.04.06 契約分類を紐付データを元に移行 -chg sta
                        // normalflg = Chk_DataMstExist_Kybrui(sqlcnnv10, value, chkvalue, chkkbn, errstr)
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Kyrui, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 2016.04.06 契約分類を紐付データを元に移行 -chg end
                case "物件分類":
                    {
                        // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg sta
                        // normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Bkrui, value, chkvalue, chkkbn, errstr)
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Bkrui, value, ref chkvalue, chkkbn, ref errstr, defvalue);
                        break;
                    }
                // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg end
                case "構造":
                    {
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Kozo, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "部屋分類":
                    {
                        // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg sta
                        // normalflg = Chk_DataMstExist_RelItem(Hash_Rel_Hyrui, value, chkvalue, chkkbn, errstr)
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Hyrui, value, ref chkvalue, chkkbn, ref errstr, defvalue);
                        break;
                    }
                // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg end
                case "パス":
                    {
                        normalflg = Chk_DataExist_FilePath(value, ref chkvalue, ref errstr);
                        break;
                    }
                case "設備":
                    {
                        normalflg = Chk_DataMstExist_RelItem_setubi(CommonModule.Hash_Rel_Setubi, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "取引態様":
                    {
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Toritaiyo, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "口座種別":
                    {
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Kozasyubetu, value, ref chkvalue, chkkbn, ref errstr);
                        // 20161004 口座種別のデフォルト値設定処理追加 -add sta
                        if (string.IsNullOrEmpty(chkvalue))
                        {
                            chkvalue = "1";
                        }

                        break;
                    }
                // 20161004 口座種別のデフォルト値設定処理追加 -add end
                case "入金項目":
                    {
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Nkinkomk, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "入金区分":
                    {
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Nkinkbn, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "変動費マスタ":
                    {
                        tmp_sql = " SELECT rule_no FROM m_hendorule ";
                        normalflg = Chk_DataMstExist(sqlcnnv10, tmp_sql, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "メーター分類":   // 要紐付け作成(入金項目名→メーター分類)   ※仮で紐付けファイルだけ用意して検証
                    {
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_Rel_Hendometer, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160613 鍵情報の取得処理修正 -del sta
                // '2016.04.06 物件鍵情報の移行処理修正 -add sta
                // Case "共用鍵タイトル"
                // '20160530 ログ修正 -chg sta
                // 'normalflg = Chk_DataMstExist_RelItem(Hash_KagiTitleKyoyo, value, chkvalue, chkkbn, errstr)
                // normalflg = Chk_DataMstExist_RelItem_kagi(value, chkvalue, chkkbn, errstr)
                // '20160530 ログ修正 -chg end
                // Case "専用鍵タイトル"
                // '20160530 ログ修正 -chg sta
                // 'normalflg = Chk_DataMstExist_RelItem(Hash_KagiTitleSenyo, value, chkvalue, chkkbn, errstr)
                // normalflg = Chk_DataMstExist_RelItem_kagi(value, chkvalue, chkkbn, errstr)
                // '20160530 ログ修正 -chg end
                // '2016.04.06 物件鍵情報の移行処理修正 -add end
                // '2016.04.26 メインの方へも反映させる修正 -add sta
                // 20160613 鍵情報の取得処理修正 -del end


                // 20160613 鍵情報の取得処理修正 -add sta
                case "共用鍵タイトル":
                case "専用鍵タイトル":
                    {
                        normalflg = Chk_DataMstExist_RelItem_kagi(value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160613 鍵情報の取得処理修正 -add end

                case "請求パターン":
                case "請求間隔":
                case "請求発生年":
                case "請求発生月":
                    {
                        normalflg = Chk_Data_SqHasseiItem(value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 2016.04.26 メインの方へも反映させる修正 -add end


                // 2016.04.26 メインの方へも反映させる修正 -add sta
                case "自社_口座用":
                case "自社口座_口座用":
                    {
                        normalflg = Chk_Data_FBKoza(sqlcnnv10, value, ref chkvalue, chkkbn, ref errstr);    // 20160915 各口座情報の自社口座およびフォーマットNo設定方法修正 引数 sqlcnnv10 を追加
                        break;
                    }
                // 2016.04.26 メインの方へも反映させる修正 -add end

                case "家主_口座用":
                case "家主口座_口座用":    // 20161005 家賃入金口座に紐付く家主口座有無チェック処理追加 -add
                    {
                        normalflg = Chk_Data_YatinOwKoza(sqlcnnv10, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }


                // 2016.04.26 メインの方へも反映させる修正 -add sta
                case "口座振替":
                    {
                        normalflg = Chk_DataMstExist_KozaFkae(sqlcnnv10, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 2016.04.26 メインの方へも反映させる修正 -add end

                // 20160609 紐付設定値取得に伴う修正 -del sta
                // '20160517 入出金取得情報の新規作成 -add sta
                // Case "FBフォーマット"
                // normalflg = Chk_DataMstExist_RelItem(Hash_Rel_FBFmt, value, chkvalue, chkkbn, errstr)
                // '20160517 入出金取得情報の新規作成 -add end
                // 20160609 紐付設定値取得に伴う修正 -del end

                // 20160609 紐付設定値取得に伴う修正 -add sta
                case "FB関連_口座振替":
                case "FB関連_総合振込":
                case "FB関連_入出金フォーマット":
                case "FB関連_入出金自社No":
                    {
                        normalflg = Chk_DataMstExist_RelItem_FB(value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160609 紐付設定値取得に伴う修正 -add end

                // 20160524 クレーム情報移行処理実装　-add sta
                case "箇所分類":
                    {
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_ClaimBruiKasyo, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                case "クレーム分類":
                    {
                        normalflg = Chk_DataMstExist_RelItem(CommonModule.Hash_ClaimBruiClaim, value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20160524 クレーム情報移行処理実装　-add end

                // 20160525 クレーム関連ファイルの画像判別処理実装 -add sta
                case "画像判別":
                    {
                        normalflg = Chk_PictureFile(value, ref chkvalue, ref errstr);
                        break;
                    }
                // 20160525 クレーム関連ファイルの画像判別処理実装 -add end

                // 20160530 ログ修正 -add sta
                case "担当者":
                    {
                        // 構築中
                        chkvalue = value;
                        break;
                    }
                // 20160530 ログ修正 -add end
                case "口座名義カナ":     // 20160829 口座名義カナチェック機能の追加 -add
                    {
                        normalflg = Chk_Kana(value, ref chkvalue, ref errstr);
                        break;
                    }
                case "金融機関_自社口座":
                case "金融機関_家主口座":             // 20160829 自社口座にデフォルト値を設定する処理を追加 -add '20161005 自社、家主口座デフォルト値設定処理追加 "金融機関_家主口座" を追加 -add
                    {
                        // 20161005 自社、家主口座デフォルト値設定処理追加 -chg sta
                        // normalflg = Chk_Jisyakoza(value, chkvalue, errstr)
                        normalflg = Chk_JisyaOwkoza(value, ref chkvalue, chkkbn, ref errstr);
                        break;
                    }
                // 20161005 自社、家主口座デフォルト値設定処理追加 -chg end
                case "修繕維持箇所分類":             // 20160829 箇所未登録データにデフォルト値を設定 -add
                    {
                        normalflg = ChK_KasyoBrui(value, ref chkvalue, ref errstr);
                        break;
                    }
                case "URL":  // 20160829 メールアドレス、URLの正規化処理を追加
                    {
                        normalflg = Chk_URL(value, ref chkvalue, maxsize, ref errstr);
                        break;
                    }

            }

        }

        /// <summary>
        /// 個別データチェックメイン処理(移行処理からの第一階層)
        /// </summary>
        /// <param name="tblname"></param>
        /// <param name="keyvalue"></param>
        /// <param name="list"></param>
        /// <param name="hash_chkbefore"></param>
        /// <param name="hash_chkafter"></param>
        /// <param name="conditioncnt"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataPerItem(SqlConnection sqlcnnv10, string tblname, List<string> list_chkduplicate, List<string> list_basekeydata, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt, ref SafeDictionary<string, string> hash_log)
        {
            // kakaka ↑"keyvalue"は読み込んだセル値が渡されるが、どこで使用する？
            // kakaka4 0322_1435 chg s 
            // Dim rtn As Boolean = True
            bool rtn = false;
            // kakaka4 0322_1435 chg e
            bool condflg = false;
            string tmp_mainkey = "";
            var midkeytblfldname = new string[11];
            var midkeyvalue = new string[11];

            // 2016.04.26 メインの方へも反映させる修正 -add
            bool chk_duplicateflg = true;                          // 重複チェック実行フラグ

            // 20160928 親有無チェック実行フラグの設定処理追加 -add sta
            bool parentchkflg = true;                              // 親有無チェック実行フラグ

            // ------------------------
            // 各移行値チェック
            // ------------------------

            foreach (var hashvalue in hash_chkbefore)
            {

                string tmp_key = Conversions.ToString(hashvalue.Key);                           // 移行項目名 (アルファベット)                                              'kakaka 説明を入れてください。
                string get_key = tblname + "-" + tmp_key;                 // テーブル名-フィールド名 (アルファベット)
                string tmp_value = Conversions.ToString(hashvalue.Value);                       // 移行値
                string tmp_type = Conversions.ToString(CommonModule.Hash_FiledTypeAlpha[get_key]);      // データ型
                string tmp_size_min = Conversions.ToString(CommonModule.Hash_Min_Code[get_key]);        // 最小値 (画面上)
                string tmp_size_max = Conversions.ToString(CommonModule.Hash_Max_Code[get_key]);        // 最大値 (画面上)
                string tmp_defvalue = Conversions.ToString(CommonModule.Hash_DefaultValue[get_key]);    // デフォルト値
                string tmp_chkvalue = "";                                 // データチェック後の値の格納用
                string errstr = "";                                       // エラー文字列格納用 (ログ出力)

                bool normalflg = true;
                bool keyflg = false;
                bool requiredflg = false;

                // 2016.04.26 メインの方へも反映させる修正 -add sta
                // 重複チェック実行フラグの設定
                // 一部の移行項目は重複チェックできないため項目毎に処理を分ける
                string tmp_tblnamejp = Conversions.ToString(CommonModule.Hash_TblName_AlphaToJp[tblname]);
                switch (tmp_tblnamejp ?? "")
                {
                    case "変動費検針情報":
                    case "家主固定控除情報":
                    case "家主請求控除情報":
                    case "その他請求情報":
                        {
                            chk_duplicateflg = false;
                            break;
                        }

                    default:
                        {
                            chk_duplicateflg = true;
                            break;
                        }
                }
                // 2016.04.26 メインの方へも反映させる修正 -add end

                // 20160928 親有無チェック実行フラグの設定処理追加 -add sta
                switch (tmp_tblnamejp ?? "")
                {
                    case "物件基本情報":
                    case "クレーム基本情報":
                    case "契約者基本情報":
                    case "家主基本情報":
                    case "仲介業者基本情報":
                    case "保険業者基本情報":
                    case "家賃保証業者基本情報":
                    case "修繕業者基本情報":
                    case "ライフライン業者情報":
                    case "施工業者情報":
                    case "施設保守業者情報":
                    case "自社基本情報":
                    case "自社担当者情報":
                    case "振込依頼人情報":
                    case "口座振替情報":
                    case "入出金取得情報":
                    case "家賃入金口座情報":
                    case "ANSERエリア情報":
                    case "ANSERアクセスポイント情報":
                    case "ANSER接続情報":
                    case "送信設定基本情報":
                    case "ポータル連動部屋分類情報":
                    case "バス交通マスタ":
                    case "鍵タイトルマスタ":
                    case "特約マスタ":
                    case "クレーム分類設定内容":
                    case "契約分類マスタ":
                    case "保険種類マスタ":
                    case "学校区マスタ":
                    case "エリアマスタ":
                    case "変動費設定内容":
                    case "備考タイトルマスタ":
                    case "画像タイトルマスタ":
                        {
                            parentchkflg = false;
                            break;
                        }

                    default:
                        {
                            parentchkflg = true;
                            break;
                        }
                }

                // 20160928 親有無チェック実行フラグの設定処理追加 -add end

                // キーフィールドの判別
                if (CommonModule.Hash_MidKey_FieldAlpha.ContainsKey(get_key))
                {

                    // 2016.02.22 型落ちを考慮した処理へ修正 -chg sta
                    // Dim tmp_keyno As Integer = Int32.Parse(Hash_MidKey_FieldAlpha.Item(get_key))        'kakaka 型落ちは大丈夫？
                    // midkeytblfldname(tmp_keyno) = get_key
                    // midkeyvalue(tmp_keyno) = tmp_value
                    // keyflg = True
                    string tmp_nostr = Conversions.ToString(CommonModule.Hash_MidKey_FieldAlpha[get_key]);
                    int tmp_keyno = 0;
                    if (int.TryParse(tmp_nostr, out tmp_keyno))
                    {
                        midkeytblfldname[tmp_keyno] = get_key;
                        midkeyvalue[tmp_keyno] = tmp_value;
                        keyflg = true;
                    }
                    // 2016.02.22 型落ちを考慮した処理へ修正 -chg end

                }

                // キーではないが登録する際に必須となる項目の判別(契約情報の契約開始日等)                 'kakaka メモ：実際の革命にて登録を行って確認した結果。型による一定のチェックとは別。
                if (CommonModule.List_Required_Field.Contains(get_key))                                           // kakaka メモ：作業ファイルへ予め用意。それがセットされたもの
                {
                    requiredflg = true;
                }

                // データチェック                                                                        'kakaka メモ：型(共通)チェック
                Call_ChkMethodPerItem(tmp_value, ref tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, ref normalflg, ref errstr);      // kakaka4 ここのログ内容は出力されない場合がある？

                // 参照マスタ有無チェック                                                                'kakaka わかり易い文言にして下さい(参照マスタチェック(設定元マスタがあるかどうか))
                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // チェック項目名称の外出し (後で使用するため)
                // If Hash_Mst_ReferenceAlpha.Contains(get_key) Then
                // Dim chkkbn As String = Hash_Mst_ReferenceAlpha.Item(get_key)
                // Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr)                          'kakaka4 ここのログ内容は出力されない場合がある？
                // End If
                string chkkbn = "";
                if (CommonModule.Hash_Mst_ReferenceAlpha.ContainsKey(get_key))
                {
                    chkkbn = Conversions.ToString(CommonModule.Hash_Mst_ReferenceAlpha[get_key]);
                    // 20160531 不正金額の移行制御修正 -chg sta
                    // Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr)                          'kakaka4 ここのログ内容は出力されない場合がある？
                    if (chkkbn == "金額")
                    {
                    }
                    // 金額の場合はチェックで問題があった場合に後で移行制御を行うためここでは個別データチェックを行わない
                    else if (chkkbn == "URL")     // 20160829 メールアドレス、URLの正規化処理を追加
                    {
                        Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, ref tmp_chkvalue, chkkbn, ref normalflg, ref errstr, tmp_size_max);
                    }
                    else
                    {
                        // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg sta
                        // Call DataChk.Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, tmp_chkvalue, chkkbn, normalflg, errstr)                          'kakaka4 ここのログ内容は出力されない場合がある？
                        Call_ChkMethodPerItem_MstExist(sqlcnnv10, tmp_value, ref tmp_chkvalue, chkkbn, ref normalflg, ref errstr, "", tmp_defvalue);
                        // 20170524 物件および部屋分類が空の場合にデフォルト値を設定して移行する対応 -chg end
                    }                          // kakaka4 ここのログ内容は出力されない場合がある？
                    // 20160531 不正金額の移行制御修正 -chg end
                }
                // 2016.04.26 メインの方へも反映させる修正 -chg end

                // 2016.02.22 エラー処理をまとめる -chg sta
                // ↓↓↓旧srcコメントアウト↓↓↓
                // 'チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける                    'kakaka メモ：上記までの処理の結果で不備の場合、処理をしない
                // If (keyflg AndAlso Not normalflg) OrElse (keyflg AndAlso tmp_chkvalue = "") Then                   'kakaka 可能な場合、Notではなく、flg=Falseを使用下さい(可読性の為)
                // errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
                // hash_log.Add(get_key, errstr)                                                      'kakaka エラーの場合の処理を纏められたらまとめて下さい(ここと、この下)
                // rtn = False
                // Return rtn
                // End If

                // 'チェックデータが必須項目かつ空の場合処理を抜ける
                // If requiredflg AndAlso tmp_chkvalue = "" Then
                // errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                // hash_log.Add(get_key, errstr)
                // rtn = False
                // Return rtn
                // End If
                // ↑↑↑旧srcコメントアウト↑↑↑

                // kakaka4 0322_1435 chg s
                // If (keyflg AndAlso normalflg = False) OrElse (keyflg AndAlso tmp_chkvalue = "") Then            'チェックデータがキーかつエラー、またはキーかつ空の場合は移行不可とする
                // errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
                // rtn = False
                // ElseIf requiredflg AndAlso tmp_chkvalue = "" Then                                           'チェックデータが必須項目かつ空の場合は移行不可とする
                // errstr = LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED & "-" & LOG_HUBI_NOTEXISTDATA_REQUIRED & "-" & LOG_TAISYO_NOTEXISTDATA_REQUIRED
                // rtn = False
                // End If
                // If rtn = False AndAlso errstr <> "" Then
                // hash_log.Add(get_key, errstr)
                // Return rtn
                // End If

                // チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける                    'kakaka メモ：上記までの処理の結果で不備の場合、処理をしない
                if (keyflg && normalflg == false || keyflg && string.IsNullOrEmpty(tmp_chkvalue))                   // kakaka 可能な場合、Notではなく、flg=Falseを使用下さい(可読性の為)
                {
                    // 20161014 ログ出力内容修正 -chg sta
                    // キーエラーを優先してログに出力する(データ型やサイズに関することに触れいているため)
                    // '20160530 ログ修正 -del ログの詳細な内容が表示されないためコメントアウト
                    // 'errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
                    hash_log.Clear();
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH_KEY + "-" + CommonModule.LOG_HUBI_MISMATCH_KEY + "-" + CommonModule.LOG_TAISYO_MISMATCH_KEY + "(空行の場合は調整不要です)";
                    // 20161014 ログ出力内容修正 -chg end
                    hash_log.Add(get_key, errstr);                                                      // kakaka エラーの場合の処理を纏められたらまとめて下さい(ここと、この下)
                    return rtn;
                }

                // チェックデータが必須項目かつ空の場合処理を抜ける
                if (requiredflg && string.IsNullOrEmpty(tmp_chkvalue))
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED;
                    hash_log.Add(get_key, errstr);
                    return rtn;
                }
                // kakaka4 0322_1435 chg e
                // 2016.02.22 エラー処理をまとめる -chg end

                // 2016.04.26 メインの方へも反映させる修正 -add sta
                // キー、必須項目以外の項目でデータの移行を制御する特殊な場合 (増えた場合は外出しすること)
                switch (chkkbn ?? "")
                {
                    case "請求パターン":
                    case "請求間隔":
                    case "請求発生年":
                    case "請求発生月":
                        {
                            if (normalflg == false)
                            {
                                errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH_NOTDEFAULT + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                                hash_log.Add(get_key, errstr);
                                return rtn;
                            }

                            break;
                        }
                    // 20160525 クレーム関連ファイルの画像判別処理実装 -add sta
                    case "画像判別":
                        {
                            if (normalflg == false)
                            {
                                hash_log.Add(get_key, errstr);
                                return rtn;
                            }

                            break;
                        }
                    // 20160525 クレーム関連ファイルの画像判別処理実装 -add end
                    // 20160531 不正金額の移行制御修正 -add sta
                    case "金額":
                        {
                            if (normalflg == false)
                            {
                                hash_log.Add(get_key, errstr);
                                return rtn;
                            }

                            break;
                        }
                        // 20160531 不正金額の移行制御修正 -add end
                }
                // 2016.04.26 メインの方へも反映させる修正 -add end

                // 調整された値が存在する場合の処理
                if (normalflg == false)                                                                // kakaka 可能な場合、Notではなく、flg=Falseを使用下さい(可読性の為)
                {
                    // 調整フラグをONにする
                    condflg = true;
                    // エラー内容をオブジェクトへ格納
                    hash_log.Add(get_key, errstr);
                }

                // チェックした値をチェック済み格納用ハッシュテーブルへ格納
                hash_chkafter.Add(tmp_key, tmp_chkvalue);

            }

            // ------------------------
            // 重複チェック
            // ------------------------

            string tmp_tblfldvalue = "";
            string tmp_keyvalue = "";

            // 20160530 ログ修正 -add sta
            // テーブル名(日本語)を取得しておく
            string tblnamejp = Conversions.ToString(CommonModule.Hash_TblName_AlphaToJp[tblname]);
            // 20160530 ログ修正 -add end

            // 退避しておいたキーフィールドと値を取得
            for (int cntii = 1, loopTo = Information.UBound(midkeyvalue); cntii <= loopTo; cntii++)
            {
                if (!string.IsNullOrEmpty(midkeyvalue[cntii]))
                {
                    tmp_tblfldvalue = tmp_tblfldvalue + "/" + midkeytblfldname[cntii];
                    tmp_keyvalue = tmp_keyvalue + "-" + midkeyvalue[cntii];
                }
            }
            tmp_tblfldvalue = tmp_tblfldvalue.Remove(0, 1);                                          // kakaka メモ：先頭文字除去
            tmp_keyvalue = tmp_keyvalue.Remove(0, 1);

            // 照合
            // 2016.04.26 メインの方へも反映させる修正 -chg sta
            // 重複チェック実行有無の条件追加
            // If list_chkduplicate.Contains(tmp_keyvalue) Then                                        'kakaka メモ：先で貯めこんでいた重複データを比較
            // hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
            // hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)
            // 'kakaka4 0322_1435  del
            // 'rtn = False
            // Return rtn
            // End If
            if (chk_duplicateflg)
            {
                if (list_chkduplicate.Contains(tmp_keyvalue))                                        // kakaka メモ：先で貯めこんでいた重複データを比較
                {
                    // 20160530 ログ修正 -chg sta
                    // 重複エラーログを出力する項目を条件に追加
                    // hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
                    // hash_log.Add(tmp_tblfldvalue, LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & LOG_TAISYO_OVERLAP)
                    if (tblnamejp == "学校区マスタ")
                    {
                    }
                    // 重複ログを出力しない
                    else
                    {
                        hash_log.Clear();    // 重複/親マスタ有無のエラーを優先する
                        hash_log.Add(tmp_tblfldvalue, CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP);
                    }
                    // 20160530 ログ修正 -chg end
                    return rtn;
                }
            }
            // 2016.04.26 メインの方へも反映させる修正 -chg end

            // ------------------------
            // 親マスタ有無チェック
            // ------------------------
            // 2016.04.26 メインの方へも反映させる修正 -chg sta
            // '照合
            // If list_basekeydata.Count <> 0 Then
            // If Not list_basekeydata.Contains(midkeyvalue(1)) Then
            // hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
            // hash_log.Add(midkeytblfldname(1), LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE)
            // 'kakaka4 del
            // 'rtn = False
            // Return rtn
            // End If
            // End If

            // 20161208 部屋情報照合の際の全角半角統一処理の削除漏れ修正 -chg sta
            // '照合用に文字を編集する処理を追加 (全角半角対応)
            // '20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -chg sta
            // 'Dim tmp_chgkeyvalu As String = StrConv(midkeyvalue(1), VbStrConv.Narrow)
            // Dim tmp_chgkeyvalu As String = ""
            // If tblnamejp = "備考入力補助リストマスタ" Then
            // tmp_chgkeyvalu = StrConv(midkeyvalue(1) & "-" & midkeyvalue(2), VbStrConv.Narrow)
            // Else
            // tmp_chgkeyvalu = StrConv(midkeyvalue(1), VbStrConv.Narrow)
            // End If
            // '20160621 修繕関連移行処理追加(備考タイトルマスタを含める) -chg end
            // If parentchkflg Then     '20160928 親有無チェック実行フラグの設定処理追加 条件変更 list_basekeydata.Count <> 0 → parentchkflg -chg
            // If Not list_basekeydata.Contains(tmp_chgkeyvalu) Then
            // hash_log.Clear()    '重複/親マスタ有無のエラーを優先する
            // hash_log.Add(midkeytblfldname(1), LOG_NAIYO_ERR_NOTEXISTDATA_BASE & "-" & LOG_HUBI_NOTEXISTDATA_BASE & "-" & LOG_TAISYO_NOTEXISTDATA_BASE)
            // 'kakaka4 del
            // 'rtn = False
            // Return rtn
            // End If
            // End If
            // '2016.04.26 メインの方へも反映させる修正 -chg end

            string tmp_chgkeyvalu = "";
            if (tblnamejp == "備考入力補助リストマスタ")
            {
                tmp_chgkeyvalu = midkeyvalue[1] + "-" + midkeyvalue[2];
            }
            else
            {
                tmp_chgkeyvalu = midkeyvalue[1];
            }
            if (parentchkflg)
            {
                if (!list_basekeydata.Contains(tmp_chgkeyvalu))
                {
                    hash_log.Clear();    // 重複/親マスタ有無のエラーを優先する
                    hash_log.Add(midkeytblfldname[1], CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_BASE + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_BASE);
                    return rtn;
                }
            }
            // 20161208 部屋情報照合の際の全角半角統一処理の削除漏れ修正 -chg end

            // --------------------------------------------
            // 調整フラグがTrueの場合、調整件数を追加する
            // --------------------------------------------

            if (condflg)
            {
                conditioncnt = conditioncnt + 1;
            }

            // kakaka4 0322_1435 add
            rtn = true;
            return rtn;

        }

        /// <summary>
        /// 数値型フィールドから取得したサイズから最大最小値を作成
        /// </summary>
        /// <param name="size"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <remarks></remarks>
        public static void Chg_SizeToValue(int size, ref object min, ref object max)
        {

            min = Math.Pow(2d, size * 8 - 1) * -1;
            max = Math.Pow(2d, size * 8 - 1) - 1d;

        }

        /// <summary>
        /// 中間ファイルデータチェックメイン処理 '20160812 汎用コンバート対応
        /// </summary>
        /// <param name="tblname"></param>
        /// <param name="keyvalue"></param>
        /// <param name="list"></param>
        /// <param name="hash_chkbefore"></param>
        /// <param name="hash_chkafter"></param>
        /// <param name="conditioncnt"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataPerItem_Mid(string tblname, ref SafeDictionary<string, string> hash_miditem, List<string> list_chkduplicate, ref string[] midkeytblfldname, ref string[] midkeyvalue, ref SafeDictionary<string, string> hash_log)
        {

            bool rtn = false;
            bool condflg = false;
            string tmp_mainkey = "";

            bool chk_duplicateflg = true;                          // 重複チェック実行フラグ(現時点で未使用)

            // ------------------------
            // 各移行値チェック
            // ------------------------

            foreach (var hashvalue in hash_miditem)
            {

                string tmp_key = Conversions.ToString(hashvalue.Key);                           // 移行項目名 (アルファベット)         
                // 20160913_2 「@#@」を共通変数に変更 -chg sta
                // Dim get_key As String = tblname & "@#@" & tmp_key                 'テーブル名-フィールド名 (アルファベット)
                string get_key = tblname + CommonModule.STR_SPLIT_1 + tmp_key;                 // テーブル名-フィールド名 (アルファベット)
                // 20160913_2 「@#@」を共通変数に変更 -chg end
                string tmp_value = Conversions.ToString(hashvalue.Value);                       // 移行値
                string tmp_type = Conversions.ToString(CommonModule.Hash_ExistMiddatatype[get_key]);    // データ型
                string tmp_size_min = Conversions.ToString(CommonModule.Hash_ExistMiddatamin[get_key]); // 最小値 (画面上)
                string tmp_size_max = Conversions.ToString(CommonModule.Hash_ExistMiddatamax[get_key]); // 最大値 (画面上)
                string tmp_defvalue = Conversions.ToString(CommonModule.Hash_ExistMiddatadef[get_key]); // デフォルト値
                string tmp_chkvalue = "";                                 // データチェック後の値の格納用
                string errstr = "";                                       // エラー文字列格納用 (ログ出力)

                bool normalflg = true;
                bool keyflg = false;
                bool requiredflg = false;

                // 重複チェック実行フラグの設定(既存用の処理を流用)
                // 一部の移行項目は重複チェックできないため項目毎に処理を分ける
                // 20160927 ログの内容が不正になっているため修正 -del
                // Dim tmp_tblnamejp As String = Hash_TblName_AlphaToJp.Item(tblname)
                switch (tblname ?? "")   // 20160927 ログの内容が不正になっているため修正 tmp_tblnamejp → tblname -chg
                {
                    case "変動費検針情報":
                    case "家主固定控除情報":
                    case "家主請求控除情報":
                    case "その他請求情報":
                        {
                            chk_duplicateflg = false;
                            break;
                        }

                    default:
                        {
                            chk_duplicateflg = true;
                            break;
                        }
                }

                // キーフィールドの判別
                if (CommonModule.Hash_ExistMiddatakey.ContainsKey(get_key))
                {

                    string tmp_nostr = Conversions.ToString(CommonModule.Hash_ExistMiddatakey[get_key]);
                    int tmp_keyno = 0;
                    if (int.TryParse(tmp_nostr, out tmp_keyno))
                    {
                        midkeytblfldname[tmp_keyno] = get_key;
                        midkeyvalue[tmp_keyno] = tmp_value;
                        keyflg = true;
                    }

                }

                // キーではないが登録する際に必須となる項目の判別(契約情報の契約開始日等)    
                if (CommonModule.List_ExistMiddatareq.Contains(get_key))
                {
                    requiredflg = true;
                }
                // 20160927 ログの内容が不正になっているため修正 -chg sta
                // '参照マスタ有無チェックは移行データが必要のためここでは行わない                                                    
                // Dim chkkbn As String = ""
                // If Hash_ExistMiddataref.Contains(get_key) Then
                // chkkbn = Hash_Mst_ReferenceAlpha.Item(get_key)
                // Else
                // 'データチェック
                // Call DataChk.Call_ChkMethodPerItem(tmp_value, tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, normalflg, errstr)
                // End If
                string chkkbn = "";
                if (CommonModule.Hash_ExistMiddataref.ContainsKey(get_key))
                {
                    chkkbn = Conversions.ToString(CommonModule.Hash_ExistMiddataref[get_key]);
                }
                // データチェック
                // 20160927 ログの内容が不正になっているため修正 -add sta
                if ((tblname == "送金ルール基本情報" | tblname == "送金ルール送金先情報" | tblname == "送金ルール入金項目情報" | tblname == "送金ルール控除項目情報") & tmp_key == "物件No")
                {
                    string[] tmp_str = tmp_value.Split('-');
                    tmp_value = tmp_str[0];
                }
                // 20160927 ログの内容が不正になっているため修正 -add end
                Call_ChkMethodPerItem(tmp_value, ref tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, ref normalflg, ref errstr);
                // 20160927 ログの内容が不正になっているため修正 -chg end
                // チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける      
                if (keyflg && normalflg == false || keyflg && string.IsNullOrEmpty(tmp_chkvalue))
                {
                    // 20161028 ログ出力内容の修正 -chg sta
                    // キーがエラーになっている場合はその内容を優先する
                    // If errstr = "" Then
                    // errstr = LOG_NAIYO_ERR_MISMATCH_KEY & "-" & LOG_HUBI_MISMATCH_KEY & "-" & LOG_TAISYO_MISMATCH_KEY
                    // End If
                    errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH_KEY + "-" + CommonModule.LOG_HUBI_MISMATCH_KEY + "-" + CommonModule.LOG_TAISYO_MISMATCH_KEY + "(空行の場合は調整不要です)";
                    // 20161028 ログ出力内容の修正 -chg end
                    // 20160927 ログの内容が不正になっているため修正 キーが既にエラーになっている場合はその内容を優先する -add
                    hash_log.Clear();
                    hash_log.Add(get_key, errstr);
                    return rtn;
                }

                // チェックデータが必須項目かつ空の場合処理を抜ける
                if (requiredflg && string.IsNullOrEmpty(tmp_chkvalue))
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED;
                    // 20160927 ログの内容が不正になっているため修正 必須項目エラーの内容を優先する -add
                    hash_log.Clear();
                    hash_log.Add(get_key, errstr);
                    return rtn;
                }

                // キー、必須項目以外の項目でデータの移行を制御する特殊な場合 (既存用の処理を流用)
                switch (chkkbn ?? "")
                {
                    case "請求パターン":
                    case "請求間隔":
                    case "請求発生年":
                    case "請求発生月":
                        {
                            if (normalflg == false)
                            {
                                errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH + "-" + CommonModule.LOG_HUBI_MISMATCH_NOTDEFAULT + "-" + CommonModule.LOG_TAISYO_DEFAULT;
                                hash_log.Add(get_key, errstr);
                                return rtn;
                            }

                            break;
                        }
                    case "画像判別":
                        {
                            if (normalflg == false)
                            {
                                hash_log.Add(get_key, errstr);
                                return rtn;
                            }

                            break;
                        }
                    case "金額":
                        {
                            if (normalflg == false)
                            {
                                hash_log.Add(get_key, errstr);
                                return rtn;
                            }

                            break;
                        }
                }

                // 調整された値が存在する場合の処理
                if (normalflg == false)
                {
                    // 調整フラグをONにする
                    condflg = true;
                    // エラー内容をオブジェクトへ格納
                    hash_log.Add(get_key, errstr);
                }

            }

            // ------------------------
            // 重複チェック
            // ------------------------
            string tmp_tblfldvalue = "";
            string tmp_keyvalue = "";

            // 退避しておいたキーフィールドと値を取得
            for (int cntii = 1, loopTo = Information.UBound((Array)midkeyvalue); cntii <= loopTo; cntii++)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(midkeyvalue[cntii], "", false)))
                {
                    // 20160927 ログの内容が不正になっているため修正 -chg sta
                    // tmp_tblfldvalue = tmp_tblfldvalue & "/" & midkeytblfldname(cntii)
                    tmp_tblfldvalue = Conversions.ToString(Operators.ConcatenateObject(tmp_tblfldvalue + CommonModule.STR_SPLIT_2, midkeytblfldname[cntii]));
                    // 20160927 ログの内容が不正になっているため修正 -chg end
                    tmp_keyvalue = Conversions.ToString(Operators.ConcatenateObject(tmp_keyvalue + "-", midkeyvalue[cntii]));
                }
            }
            tmp_tblfldvalue = tmp_tblfldvalue.Remove(0, 1);
            tmp_keyvalue = tmp_keyvalue.Remove(0, 1);

            // 照合
            if (chk_duplicateflg)
            {
                if (list_chkduplicate.Contains(tmp_keyvalue))
                {
                    if (tblname == "学校区マスタ")
                    {
                    }
                    // 重複ログを出力しない
                    else
                    {
                        hash_log.Clear();    // 重複/親マスタ有無のエラーを優先する
                        hash_log.Add(tmp_tblfldvalue, CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP);
                    }
                    return rtn;
                }
                list_chkduplicate.Add(tmp_keyvalue);
            }

            rtn = true;
            return rtn;

        }

        /// <summary>
        /// 汎用中間ファイルデータチェックメイン処理 '20160812 汎用コンバート対応
        /// </summary>
        /// <param name="tblname"></param>
        /// <param name="keyvalue"></param>
        /// <param name="list"></param>
        /// <param name="hash_chkbefore"></param>
        /// <param name="hash_chkafter"></param>
        /// <param name="conditioncnt"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DataPerItem_BaseMid(string tblname, ref SafeDictionary<string, string> hash_miditem, List<string> list_chkduplicate, ref string[] midkeytblfldname, ref string[] midkeyvalue, ref SafeDictionary<string, string> hash_log)
        {

            bool rtn = false;
            bool condflg = false;
            string tmp_mainkey = "";

            bool chk_duplicateflg = true;                          // 重複チェック実行フラグ(現時点で未使用)

            // ------------------------
            // 各移行値チェック
            // ------------------------

            foreach (var hashvalue in hash_miditem)
            {

                string tmp_key = Conversions.ToString(hashvalue.Key);                           // 移行項目名 (アルファベット)     
                // 20160913_2 「@#@」を共通変数に変更 -chg sta
                // Dim get_key As String = tblname & "@#@" & tmp_key                 'テーブル名-フィールド名 (アルファベット)
                string get_key = tblname + CommonModule.STR_SPLIT_1 + tmp_key;                 // テーブル名-フィールド名 (アルファベット)
                // 20160913_2 「@#@」を共通変数に変更 -chg end
                string tmp_value = Conversions.ToString(hashvalue.Value);                       // 移行値
                string tmp_type = Conversions.ToString(CommonModule.Hash_BaseMiddatatype[get_key]);     // データ型
                // 20161006 部屋設備チェック処理の修正 -chg sta
                // Dim tmp_size_min As String = Hash_BaseMiddatamin.Item(get_key)  '最小値 (画面上)
                string tmp_size_min = Conversions.ToString(Interaction.IIf(Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(CommonModule.Hash_BaseMiddatamin[get_key], "未入力", false)), "0", CommonModule.Hash_BaseMiddatamin[get_key]));  // 最小値 (画面上)
                // 20161006 部屋設備チェック処理の修正 -chg end
                string tmp_size_max = Conversions.ToString(CommonModule.Hash_BaseMiddatamax[get_key]);  // 最大値 (画面上)
                string tmp_defvalue = Conversions.ToString(CommonModule.Hash_BaseMiddatadef[get_key]);  // デフォルト値
                string tmp_chkvalue = "";                                 // データチェック後の値の格納用
                string errstr = "";                                       // エラー文字列格納用 (ログ出力)

                bool normalflg = true;
                bool keyflg = false;
                bool requiredflg = false;

                // キーフィールドの判別
                if (CommonModule.Hash_BaseMiddatakey.ContainsKey(get_key))
                {

                    string tmp_nostr = Conversions.ToString(CommonModule.Hash_BaseMiddatakey[get_key]);
                    int tmp_keyno = 0;
                    if (int.TryParse(tmp_nostr, out tmp_keyno))
                    {
                        midkeytblfldname[tmp_keyno] = get_key;
                        midkeyvalue[tmp_keyno] = tmp_value;
                        keyflg = true;
                    }

                }

                // キーではないが登録する際に必須となる項目の判別(契約情報の契約開始日等)                 
                if (CommonModule.List_ExistMiddatareq.Contains(get_key))
                {
                    requiredflg = true;
                }

                // 参照マスタ有無チェックは移行データが必要のためここでは行わない                                                    
                string chkkbn = "";
                if (CommonModule.Hash_BaseMiddataref.ContainsKey(get_key))
                {
                    chkkbn = Conversions.ToString(CommonModule.Hash_Mst_ReferenceAlpha[get_key]);
                }
                else
                {
                    // データチェック                                                                       
                    Call_ChkMethodPerItem(tmp_value, ref tmp_chkvalue, tmp_type, tmp_size_min, tmp_size_max, tmp_defvalue, ref normalflg, ref errstr);
                }

                // チェックデータがキーかつエラー、またはキーかつ空の場合処理を抜ける 
                if (keyflg && normalflg == false || keyflg && string.IsNullOrEmpty(tmp_chkvalue))
                {
                    if (string.IsNullOrEmpty(errstr))
                    {
                        errstr = CommonModule.LOG_NAIYO_ERR_MISMATCH_KEY + "-" + CommonModule.LOG_HUBI_MISMATCH_KEY + "-" + CommonModule.LOG_TAISYO_MISMATCH_KEY + "(空行の場合は調整不要です)";
                    }
                    // 20160927 ログの内容が不正になっているため修正 キーが既にエラーになっている場合はその内容を優先する -add
                    hash_log.Clear();
                    hash_log.Add(get_key, errstr);
                    return rtn;
                }

                // チェックデータが必須項目かつ空の場合処理を抜ける
                if (requiredflg && string.IsNullOrEmpty(tmp_chkvalue))
                {
                    errstr = CommonModule.LOG_NAIYO_ERR_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_HUBI_NOTEXISTDATA_REQUIRED + "-" + CommonModule.LOG_TAISYO_NOTEXISTDATA_REQUIRED;
                    // 20160927 ログの内容が不正になっているため修正 必須項目エラーの内容を優先する -add
                    hash_log.Clear();
                    hash_log.Add(get_key, errstr);
                    return rtn;
                }

                // 調整された値が存在する場合の処理
                if (normalflg == false)
                {
                    // 調整フラグをONにする
                    condflg = true;
                    // エラー内容をオブジェクトへ格納
                    if (hash_log.ContainsKey(get_key) == false)
                    {
                        hash_log.Add(get_key, errstr);
                    }
                }

            }

            // ------------------------
            // 重複チェック
            // ------------------------
            string tmp_tblfldvalue = "";
            string tmp_keyvalue = "";

            // 退避しておいたキーフィールドと値を取得
            for (int cntii = 1, loopTo = Information.UBound((Array)midkeyvalue); cntii <= loopTo; cntii++)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(midkeyvalue[cntii], "", false)))
                {
                    tmp_tblfldvalue = Conversions.ToString(Operators.ConcatenateObject(tmp_tblfldvalue + "/", midkeytblfldname[cntii]));
                    tmp_keyvalue = Conversions.ToString(Operators.ConcatenateObject(tmp_keyvalue + "-", midkeyvalue[cntii]));
                }
            }
            tmp_tblfldvalue = tmp_tblfldvalue.Remove(0, 1);
            tmp_keyvalue = tmp_keyvalue.Remove(0, 1);

            // 照合
            if (chk_duplicateflg)
            {
                if (list_chkduplicate.Contains(tmp_keyvalue))
                {
                    if (tblname == "学校区マスタ")
                    {
                    }
                    // 重複ログを出力しない
                    else
                    {
                        hash_log.Clear();    // 重複/親マスタ有無のエラーを優先する
                        hash_log.Add(tmp_tblfldvalue, CommonModule.LOG_NAIYO_ERR_OVERLAP + "-" + CommonModule.LOG_HUBI_OVERLAP + "-" + CommonModule.LOG_TAISYO_OVERLAP);
                    }
                    return rtn;
                }
                list_chkduplicate.Add(tmp_keyvalue);
            }

            rtn = true;
            return rtn;

        }

    }

    #endregion

    #region 配列設定クラス

    /// <summary>
    /// 【配列設定クラス】
    /// </summary>
    /// <remarks></remarks>
    public class SetArrayData
    {

        private string[] myStrings;
        public SetArrayData(int capacity)
        {
            // コンストラクタで配列を作成
            myStrings = new string[capacity + 1];
        }

        // 既定のプロパティを宣言
        public string this[int index]
        {
            get
            {
                return myStrings[index];
            }
            set
            {
                myStrings[index] = value;
            }
        }

    }

    #endregion

    #region ファイル処理クラス

    /// <summary>
    /// 【ファイル処理クラス】
    /// </summary>
    /// <remarks></remarks>
    public class FileMethod
    {

        /// <summary>
        /// ファイル書込み処理
        /// </summary>
        /// <param name="str">出力文字列</param>
        /// <param name="filepath">ファイルパス(例：C:\Temp\test.txt)</param>
        /// <param name="append">ファイル追記 Ture.追記(Def) False.上書き</param>
        /// <remarks>
        /// ・指定書込ファイルが存在しない場合、新規ファイル作成<br/>
        /// ・ShiftJisで書込み<br/>
        /// ・書込みファイルパスがNULLの場合はログ出力を行わない<br/>
        /// </remarks>
        public static void FileWriteAppend(string str, string filepath, bool append = true)
        {

            if (filepath is not null)
            {
                if (!string.IsNullOrEmpty(filepath.Trim()))
                {
                    var sw = new StreamWriter(filepath, append, System.Text.Encoding.GetEncoding("shift_jis"));
                    sw.Write(str + sw.NewLine);
                    sw.Close();
                }
            }
        }

        /// <summary>
        /// Table/ViewからCSVファイル出力
        /// Optional はログを出力する際に用いる 2016.02.22 エラー時の対処2(エラー内容を格納する引数追加)
        /// </summary>
        /// <param name="sqlcnn"></param>
        /// <param name="csvfilepath"></param>
        /// <param name="taisyoname"></param>
        /// <param name="sortstr"></param>
        /// <param name="sql"></param>
        /// <param name="logflg">True:中間ファイル出力時  False:ログファイル出力時</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool TblView_Output_CSV(SqlConnection sqlcnn, string csvfilepath, string taisyoname, string sortstr, ref string errstr, string sql = "", bool logflg = false)
        {
            // kakaka4 0322_1745 変数logflgがTrueは中間ファイル出力時、Falseはログファイル出力時の内容で処理ということを明記しておく。
            bool rtn = true;

            // 抽出クエリ作成
            string tmp_str = "";
            string tmp_sql = "";
            string tmp_order = "";

            tmp_order = " ORDER BY " + sortstr;
            if (string.IsNullOrEmpty(sql))
            {
                tmp_sql = "SELECT * FROM " + taisyoname + tmp_order;
            }
            else
            {
                tmp_sql = sql;
            }

            // ファイル出力用コマンド文字列作成
            string tmp_cmdstr = "";
            string[] tmp_constr = sqlcnn.ConnectionString.Split(';');
            string servername = " -S " + tmp_constr[1].Remove(0, Strings.InStr(tmp_constr[1], "=") + 1);
            string catalogname = " -d " + tmp_constr[2].Remove(0, Strings.InStr(tmp_constr[2], "=") + 1);
            string username = " -U " + tmp_constr[3].Remove(0, Strings.InStr(tmp_constr[3], "=") + 1);
            string password = " -P " + tmp_constr[4].Remove(0, Strings.InStr(tmp_constr[4], "=") + 1);
            string strtype = " -c";
            if (logflg)
            {
                strtype = strtype + " -t,";  // ログファイルを出力する際はカンマ区切りにする
            }
            string outtype = " queryout ";
            // 20160829 CSV出力エラー修正 -chg sta
            // tmp_cmdstr = """" & tmp_sql & """" & outtype & csvfilepath & servername & catalogname & username & password & strtype
            tmp_cmdstr = "\"" + tmp_sql + "\"" + outtype + "\"" + csvfilepath + "\"" + servername + catalogname + username + password + strtype;
            // 20160829 CSV出力エラー修正 -chg end
            // 20160516 Windows認証でログファイルが出力されない現象の対応 -add sta
            if (My.MyProject.Forms.MainFrm.optV10Authent2.Checked)
            {
                tmp_cmdstr = tmp_cmdstr.Replace(username + password, " -T");
            }
            // 20160516 Windows認証でログファイルが出力されない現象の対応 -add end

            // ファイル出力処理
            try
            {
                var proc = new Process();
                proc.StartInfo.FileName = "bcp";
                proc.StartInfo.Arguments = tmp_cmdstr;

                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.UseShellExecute = false;
                proc.Start();

                proc.WaitForExit();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                // 2016.02.22 エラー時の対処2 -add sta
                if (logflg)
                {
                    errstr = CommonModule.LOG_HUBI_MAKEMIDFILE + Constants.vbCrLf + "エラー内容：" + ex.Message;
                }
                else
                {
                    errstr = CommonModule.LOG_HUBI_MAKELOGFILE + Constants.vbCrLf + "エラー内容：" + ex.Message;
                }
                // 2016.02.22 エラー時の対処2 -add end
                rtn = false;
            }

            return rtn;

        }

        /// <summary>
        /// CSV→中間ファイル書込
        /// </summary>
        /// <param name="csvfilepath"></param>
        /// <param name="midfilepath"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static object CSV_To_Excel(string csvfilepath, string midfilepath, string sheetname, int colcnt, ref int rowcnt)
        {

            bool rtn = true;

            Microsoft.Office.Interop.Excel.Application exlapp;
            Microsoft.Office.Interop.Excel.Workbooks exlwbook;
            Microsoft.Office.Interop.Excel.Workbook exlwbookmid;
            Microsoft.Office.Interop.Excel.Worksheet exlwsheetmid;

            // ファイル準備
            exlapp = (Microsoft.Office.Interop.Excel.Application)Interaction.CreateObject("Excel.Application");
            exlapp.Visible = false;

            // CSV展開ファイルブック
            exlwbook = exlapp.Workbooks;

            // ----------------------
            // 中間ファイル設定
            // ----------------------
            // オブジェクト生成
            exlwbookmid = exlapp.Workbooks.Open(midfilepath);
            exlwsheetmid = (Microsoft.Office.Interop.Excel.Worksheet)exlapp.Worksheets[sheetname];

            // 書込開始行の設定
            int startrow = CommonModule.EXISTMIDFILE_READWRITE_ROW;

            // 既存データの確認(最大行の取得)
            int maxrowcnt = exlwsheetmid.UsedRange.Rows.Count;

            // 既存データ範囲を取得
            string dataarea = startrow + ":" + maxrowcnt;

            // 初期化
            if (maxrowcnt >= startrow)
            {
                exlwsheetmid.Rows[dataarea].Delete();
            }

            // ----------------------
            // CSV展開用ファイル設定
            // ----------------------
            object arrays;
            var arrayItems = new object[colcnt];

            for (int cntii = 0, loopTo = colcnt - 1; cntii <= loopTo; cntii++)
                arrayItems[cntii] = new object[] { cntii + 1, 2 };
            arrays = arrayItems;

            // '拡張子チェック (※出力CSVファイルの拡張子は「.tmp」で出力している)
            // '拡張子がCSVだとFieldInfoが使えないので一時ファイルにコピーして開く
            // Dim filePath As String = csvfilepath
            // If (String.Equals(System.IO.Path.GetExtension(csvfilepath), ".csv", StringComparison.OrdinalIgnoreCase)) Then
            // filePath = System.IO.Path.GetTempFileName()
            // System.IO.File.Copy(csvfilepath, filePath, True)
            // End If

            // ----------------------------
            // CSV展開→中間ファイル書込
            // ----------------------------
            try
            {
                // CSVファイル展開
                exlwbook.OpenText(csvfilepath, FieldInfo: arrays, Tab: true, DataType: 1);

                // 最終セルの位置を取得
                int maxrow = Conversions.ToInteger(exlapp.Worksheets[1].Range(CommonModule.MIDFILE_READWRITE_CELLSTA).SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).row);
                int maxcol = Conversions.ToInteger(exlapp.Worksheets[1].Range(CommonModule.MIDFILE_READWRITE_CELLSTA).SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column);

                // 中間ファイルへ書込
                string writecellsta = CommonModule.EXISTMIDFILE_READWRITE_COL + CommonModule.EXISTMIDFILE_READWRITE_ROW.ToString();
                exlapp.get_Range(exlapp.Cells[1, 1], exlapp.Cells[maxrow, maxcol]).Copy(exlwsheetmid.get_Range(writecellsta));

                // 書込行数取得
                rowcnt = exlwsheetmid.UsedRange.Rows.Count - (CommonModule.EXISTMIDFILE_READWRITE_ROW - 1);

                // 保存
                exlwbookmid.Save();
            }

            catch (Exception ex)
            {
                // エラー処理
                Debug.WriteLine(ex.Message);
                rtn = false;
            }
            finally
            {
                exlwbookmid.Close();
                exlwbook.Close();
                exlapp.Quit();
                Marshal.ReleaseComObject(exlwsheetmid);
                // 20161014 DisconnectedContextエラー修正 -chg sta
                // System.Runtime.InteropServices.Marshal.ReleaseComObject(exlwbookmid)
                // System.Runtime.InteropServices.Marshal.ReleaseComObject(exlwbook)
                Marshal.ReleaseComObject(exlwbookmid);
                Marshal.ReleaseComObject(exlwbook);
                // 20161014 DisconnectedContextエラー修正 -chg end
                Marshal.ReleaseComObject(exlapp);
                exlwsheetmid = null;
                exlwbookmid = null;
                exlwbook = null;
                exlapp = null;
            }

            return rtn;

        }

    }

    #endregion

    #region DB処理関連クラス

    public class DBExec
    {

        /// <summary>
        /// DBより全レコードを取得して返却 [DataTable]
        /// </summary>
        /// <param name="sqlstr">取得クエリ文</param>
        /// <param name="sqlcnn">接続情報</param>
        /// <param name="readtbl">読込TBL情報 [out]</param>
        /// <returns>レコード件数<br/></returns>
        /// <remarks>
        /// ・sqlcnnSH→sqlcnn<br/>
        /// ・発行クエリより抽出した全データをDataTableへ全格納
        /// </remarks>
        public static int Exec_DataTable(string sqlstr, ref SqlConnection sqlcnn, ref DataTable readtbl, [Optional, DefaultParameterValue(true)] ref bool normalflg)




        {

            var adater = new SqlDataAdapter();
            DateTime time_sta;
            // ★★★ CommonRepositry.Exec_DataTableを参考
            try
            {
                readtbl.Clear();                                                     // 蓄積データクリア
                adater.SelectCommand = new SqlCommand(sqlstr, sqlcnn);
                adater.SelectCommand.CommandTimeout = CommonModule.LimitTimeOut;                  // タイムアウト設定
                time_sta = DateTime.Now;                                                      // 開始時間
                adater.Fill(readtbl);                                                // データ取得
                Debug.Print(Typ.ToStr(DateTime.Now - DateAndTime.TimeOfDay));                             // 経過時間
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                normalflg = false;
            }

            return readtbl.Rows.Count;

        }

        /// <summary>
        /// DBより1データを取得して返却 [ExecuteScalar]
        /// </summary>
        /// <param name="sqlstr">取得クエリ文</param>
        /// <param name="sqlcnn">接続情報</param>
        /// <returns>取得値<br/></returns>
        /// <remarks>
        /// ・sqlcnnSH→sqlcnn<br/>
        /// ・前提：発行クエリ(sqlstr)は1データ抽出クエリとします
        /// </remarks>
        public static object Exec_Scalar(string sqlstr, ref SqlConnection sqlcnn)


        {
            // ★★★ CommonRepository.Exec_Scalarを参考
            var sqlcom = new SqlCommand(sqlstr, sqlcnn);
            var reccnt = Interaction.IIf(sqlcom.ExecuteScalar() is DBNull, 0, sqlcom.ExecuteScalar());
            sqlcom.Dispose();
            return reccnt;

        }

        /// <summary>
        /// DBより1レコードを取得してSafeDictionary<string, string>へ格納 [ExecuteReader]
        /// </summary>
        /// <param name="sqlstr">抽出クエリ文</param>
        /// <param name="sqlcnn">接続情報</param>
        /// <param name="hashtbl">返却抽出データ</param>
        /// <returns>抽出レコード結果：True.あり False.なし</returns>
        /// <remarks>
        /// ・前提：発行クエリ(sqlstr)は1レコード抽出クエリとします<br/>
        /// </remarks>
        public static bool Exec_DataReader(string sqlstr, ref SqlConnection sqlcnn, ref SafeDictionary<string, object> hashtbl)



        {
            // ★★★ CommonRepository.Exec_DataReaderを参考
            var sqlcom = new SqlCommand(sqlstr, sqlcnn);
            var sqldrd = sqlcom.ExecuteReader(CommandBehavior.SingleRow);
            bool recumuflg;                                // True.レコードあり

            sqldrd.Read();
            recumuflg = sqldrd.HasRows;
            if (recumuflg)
            {
                hashtbl = new SafeDictionary<string, object>();
                for (int cnt = 0, loopTo = sqldrd.FieldCount - 1; cnt <= loopTo; cnt++)
                    hashtbl.Add(sqldrd.GetName(cnt), sqldrd.GetValue(cnt));
            }

            sqlcom.Dispose();
            sqldrd.Close();
            return recumuflg;

        }

        /// <summary>
        /// DBより全レコードを取得して1列目のデータをリストへ格納
        /// 前提：抽出したテーブルの最初の1列のみ取得する
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <param name="sqlcnn"></param>
        /// <param name="list"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Exec_DataReader_Col_List(string sqlstr, ref SqlConnection sqlcnn, ref List<string> list)



        {

            var sqlcom = new SqlCommand(sqlstr, sqlcnn);
            var sqldrd = sqlcom.ExecuteReader();
            bool recumuflg;                                // True.レコードあり

            recumuflg = sqldrd.HasRows;
            if (recumuflg)
            {
                list = new List<string>();
                while (sqldrd.Read())
                    list.Add(Conversions.ToString(sqldrd.GetValue(0)));
            }

            sqlcom.Dispose();
            sqldrd.Close();
            return recumuflg;

        }

        /// <summary>
        /// DBより全レコードを取得してハッシュテーブルへ格納
        /// 前提：2列取得し、1列目の値をハッシュテーブルのキーとする
        /// </summary>
        /// <param name="sqlstr"></param>
        /// <param name="sqlcnn"></param>
        /// <param name="hash"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Exec_DataReader_Col_Hash(string sqlstr, ref SqlConnection sqlcnn, ref SafeDictionary<string, string> hash)



        {

            var sqlcom = new SqlCommand(sqlstr, sqlcnn);
            var sqldrd = sqlcom.ExecuteReader();
            bool recumuflg;                                // True.レコードあり

            recumuflg = sqldrd.HasRows;
            if (recumuflg)
            {
                hash = new SafeDictionary<string, string>();
                while (sqldrd.Read())
                {
                    string valuekey = sqldrd.GetValue(0).ToString();
                    string value = sqldrd.GetValue(1).ToString();
                    // 20160527 指摘事項対応 -chg sta
                    // hash.Add(valuekey, value)
                    if (hash.ContainsKey(valuekey) == false)
                    {
                        hash.Add(valuekey, value);
                    }
                    // 20160527 指摘事項対応 -chg end
                }
            }

            sqlcom.Dispose();
            sqldrd.Close();
            return recumuflg;

        }

        /// <summary>
        /// クエリ実行
        /// </summary>
        /// <param name="sqlcnnv10">接続情報</param>
        /// <param name="qry">クエリ文</param>
        /// <param name="rowcnt">処理件数</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Exec_NonQuery(SqlConnection sqlcnnv10, string qry, ref int rowcnt)

        {
            bool rtn = true;
            try
            {
                var sqlcom = new SqlCommand(qry, sqlcnnv10);
                rowcnt = sqlcom.ExecuteNonQuery();
                sqlcom.Dispose();
            }
            catch (Exception ex)
            {
                rtn = false;
            }
            return rtn;

        }

    }

    #endregion

    #region 個別処理クラス

    /// <summary>
    /// 【個別処理クラス】
    /// </summary>
    /// <remarks></remarks>
    public class EtcMethod
    {

        /// <summary>
        /// DB登録用パラメータ文字列作成
        /// </summary>
        /// <param name="maxcnt">パラメータ最大値</param>
        /// <param name="addcnt">パラメータ開始番号</param>
        /// <param name="skipstr">
        /// スキップ番号(","で連結)<br/>
        /// 例："5,12,18"
        /// →Para5,Para12,Para18は作成しない<br/>
        /// </param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Get_ParaNum(int maxcnt, int addcnt = 0, string skipstr = "")



        {

            var sbobj = new System.Text.StringBuilder();      // StringBuilderオブジェクト生成
            int sumcnt;                           // 作業用カウント値
            var skipcnt = default(string[]);                         // 分割SKIPカウント値
            bool skipflg = false;                  // スキップFLG：True.Skip
            // ★★★ CommonRepository.Get_ParaNumを参考
            skipstr = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(skipstr), "", skipstr));
            if (!string.IsNullOrEmpty(skipstr.Trim()))
            {
                skipcnt = skipstr.Split(',');
            }

            addcnt = Conversions.ToInteger(Interaction.IIf(addcnt == 0, 0, addcnt - 1));
            for (int cnti = 1, loopTo = maxcnt; cnti <= loopTo; cnti++)
            {
                sumcnt = cnti + addcnt;

                if (skipcnt is not null)
                {
                    for (int cntj = 0, loopTo1 = Information.UBound(skipcnt); cntj <= loopTo1; cntj++)
                    {
                        if (cnti == Conversions.ToInteger(skipcnt[cntj]))
                        {
                            skipflg = true;
                            break;
                        }
                    }
                }
                if (skipflg == false)
                {
                    sbobj.Append("@para");
                    sbobj.Append(RightStrMethod("000" + sumcnt, 3));
                    if (cnti <= maxcnt - 1)
                    {
                        sbobj.Append(",");
                    }
                }
                skipflg = false;
            }
            return sbobj.ToString();

        }

        /// <summary>
        /// 文字列抽出(RIGHTメソッド)
        /// </summary>
        /// <param name="str">対象文字列<br/></param>
        /// <param name="lenght">抽出文字数<br/></param>
        /// <returns>抽出文字列</returns>
        /// <remarks></remarks>
        public static string RightStrMethod(string str, int lenght)
        {
            // ★★★ RightStrMethodを参考
            string rtn;

            rtn = str;
            if (lenght <= str.Length)
            {
                rtn = str.Substring(str.Length - lenght);
            }
            return rtn;

        }

        /// <summary>
        /// 検索文字カウント数
        /// </summary>
        /// <param name="str">検索元文字列</param>
        /// <param name="chr">検索文字</param>
        /// <returns>検索文字カウント数</returns>
        /// <remarks></remarks>
        public static int CountChar(string str, char chr)
        {

            return str.Length - str.Replace(chr.ToString(), "").Length;

        }

        /// <summary>
        /// 対象文字列のバイト数取得
        /// </summary>
        /// <param name="target">対象文字列</param>
        /// <returns>半角1バイト、全角2バイトでカウントされたバイト数</returns>
        /// <remarks></remarks>
        public static int Get_LenB(string target)
        {

            return System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(target);

        }

        /// <summary>
        /// 英数字半角変換(一部記号を含む)
        /// </summary>
        /// <param name="str">対象文字列</param>
        /// <returns>変換後文字列</returns>
        /// <remarks>
        /// ・変換文字列：０-９Ａ-Ｚａ-ｚ：，－　<br/>
        /// </remarks>
        public static string Chg_NumNarrow(string str)
        {

            var reg = new Regex("[０-９Ａ-Ｚａ-ｚ：，－　]+");
            string output = reg.Replace(str, RepNarrow);
            return output;

        }

        /// <summary>
        /// 半角文字列変換
        /// </summary>
        /// <param name="mat">一致文字列</param>
        /// <returns>変換後文字列</returns>
        /// <remarks></remarks>
        public static string RepNarrow(Match mat)
        {

            return Strings.StrConv(mat.Value, VbStrConv.Narrow, 0);

        }

        /// <summary>
        /// 値が基準値の倍数か判断  2016.02.22 プログレスバーの表示修正
        /// </summary>
        /// <param name="value"></param>
        /// <param name="base"></param>
        /// <param name="chgvalue"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static void Get_MultipleFlg(int value, int @base, [Optional, DefaultParameterValue(0)] ref int chgvalue)
        {

            int tmp_result = 0;

            // 値が0の時は処理を抜ける
            if (value == 0)
            {
                return;
            }

            // 基準値で割った値が整数か否か判別
            string tmp_value = (value / (double)@base).ToString();
            bool flg = int.TryParse(tmp_value, out tmp_result);

            // 基準値の倍数の場合かつ割った値が必要な場合は返す
            if (flg)
            {
                chgvalue = tmp_result;
            }

        }

        /// <summary>
        /// 四捨五入処理 2016.04.04 変換時の型落ち対応(再) 四捨五入処理の共通化
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static object Get_RoundingValue(string value)
        {

            double tmp_dbl = 0d;
            string rtn_str = "";

            if (double.TryParse(value, out tmp_dbl))
            {
                // 四捨五入して取得
                rtn_str = Math.Round(tmp_dbl, MidpointRounding.AwayFromZero).ToString();
            }
            else
            {
                // 変換できない場合は値をそのまま格納 (ログ出力させる)
                rtn_str = value;
            }

            return rtn_str;

        }

        /// <summary>
        /// 入金項目区分から区分名を取得 2016.04.06 入金項目読込処理の修正
        /// </summary>
        /// <param name="tukikbn"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Get_Nkinruiname(ref string tukikbn)
        {

            string rtn_string = "";

            switch (tukikbn ?? "")
            {
                case "1":
                    {
                        rtn_string = tukikbn + "." + "通常月";
                        break;
                    }
                case "2":
                    {
                        rtn_string = tukikbn + "." + "契約時";
                        break;
                    }
                case "3":
                    {
                        rtn_string = tukikbn + "." + "更新時";
                        break;
                    }
                case "4":
                    {
                        rtn_string = tukikbn + "." + "解約時";
                        break;
                    }
                case "5":
                    {
                        rtn_string = tukikbn + "." + "随時変動";
                        break;
                    }
                case "6":
                    {
                        rtn_string = tukikbn + "." + "その他";
                        break;
                    }
                case "7":
                    {
                        rtn_string = tukikbn + "." + "修繕";
                        break;
                    }
                case "8":
                    {
                        rtn_string = tukikbn + "." + "家主送金";
                        break;
                    }
                case "9":
                    {
                        rtn_string = tukikbn + "." + "家主控除";
                        break;
                    }
            }

            return rtn_string;

        }

        /// <summary>
        /// 10設備マスタから項目guidを取得してハッシュテーブルへ格納(setubi_grpsortorder-setubi_sortorder-komok_sortorder,komk_guid) 2016.04.06 部屋設備情報の取得処理修正
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <remarks></remarks>
        public static void Set_SetubiMst(SqlConnection sqlcnnv10)
        {

            string tmp_sql = "";

            // 抽出クエリの作成
            tmp_sql = tmp_sql + " SELECT ";
            tmp_sql = tmp_sql + " 	 CONVERT(varchar,setubi_grpsortorder) + '-' + ";
            tmp_sql = tmp_sql + " 	 CONVERT(varchar,setubi_sortorder) + '-' + ";
            tmp_sql = tmp_sql + " 	 CONVERT(varchar,komok_sortorder) AS キー ";
            tmp_sql = tmp_sql + " 	,komok_guid AS [項目guid] ";
            tmp_sql = tmp_sql + " /* ";
            tmp_sql = tmp_sql + " 	,SE.setubi_guid ";
            tmp_sql = tmp_sql + " 	,komok_guid  ";
            tmp_sql = tmp_sql + " 	,setubi_grpsortorder ";
            tmp_sql = tmp_sql + " 	,setubi_grpname  ";
            tmp_sql = tmp_sql + " 	,setubi_sortorder ";
            tmp_sql = tmp_sql + " 	,setubi_name ";
            tmp_sql = tmp_sql + " 	,komok_sortorder ";
            tmp_sql = tmp_sql + " 	,komok_name ";
            tmp_sql = tmp_sql + " */ ";
            tmp_sql = tmp_sql + " FROM m_setubi_grp AS SEG ";
            tmp_sql = tmp_sql + " LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid ";
            tmp_sql = tmp_sql + " LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid ";
            tmp_sql = tmp_sql + " ORDER BY setubi_grpsortorder,setubi_sortorder,komok_sortorder ";


            // 設備のソートNoが重複しているデータが存在するためとりあえず重複を除去したクエリを作成しておく
            tmp_sql = "";
            tmp_sql = tmp_sql + " /*20160829 エレベーター移行対応 chg sta*/ ";
            tmp_sql = tmp_sql + " /* ";
            tmp_sql = tmp_sql + " SELECT [キー],[項目guid] FROM ";
            tmp_sql = tmp_sql + " ( ";
            tmp_sql = tmp_sql + "     SELECT ";
            tmp_sql = tmp_sql + "             ROW_NUMBER()OVER(PARTITION BY [キー] ORDER BY [キー]) AS [抽出用] ";
            tmp_sql = tmp_sql + "         ,* ";
            tmp_sql = tmp_sql + "     FROM ";
            tmp_sql = tmp_sql + "     ( ";
            tmp_sql = tmp_sql + "         SELECT ";
            tmp_sql = tmp_sql + "             	CONVERT(varchar,setubi_grpsortorder) + '-' + ";
            tmp_sql = tmp_sql + "             	CONVERT(varchar,setubi_sortorder) + '-' + ";
            tmp_sql = tmp_sql + "             	CONVERT(varchar,komok_sortorder) AS [キー] ";
            tmp_sql = tmp_sql + "             ,komok_guid AS [項目guid] ";
            tmp_sql = tmp_sql + "         /* ";
            tmp_sql = tmp_sql + "             ,SE.setubi_guid ";
            tmp_sql = tmp_sql + "             ,komok_guid  ";
            tmp_sql = tmp_sql + "             ,setubi_grpsortorder ";
            tmp_sql = tmp_sql + "             ,setubi_grpname  ";
            tmp_sql = tmp_sql + "             ,setubi_sortorder ";
            tmp_sql = tmp_sql + "             ,setubi_name ";
            tmp_sql = tmp_sql + "             ,komok_sortorder ";
            tmp_sql = tmp_sql + "             ,komok_name ";
            tmp_sql = tmp_sql + "         */ ";
            tmp_sql = tmp_sql + "         FROM m_setubi_grp AS SEG ";
            tmp_sql = tmp_sql + "         LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid ";
            tmp_sql = tmp_sql + "         LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid ";
            tmp_sql = tmp_sql + "     ) AS VW1 ";
            tmp_sql = tmp_sql + " ) AS VW2 ";
            tmp_sql = tmp_sql + " WHERE [抽出用] = 1 ";
            tmp_sql = tmp_sql + " */ ";
            tmp_sql = tmp_sql + " SELECT * FROM ";
            tmp_sql = tmp_sql + " ( ";
            tmp_sql = tmp_sql + " 	SELECT [キー],CONVERT(varchar(MAX),[項目guid]) AS [項目guid] FROM ";
            tmp_sql = tmp_sql + " 	( ";
            tmp_sql = tmp_sql + " 		SELECT ";
            tmp_sql = tmp_sql + " 				ROW_NUMBER()OVER(PARTITION BY [キー] ORDER BY [キー]) AS [抽出用] ";
            tmp_sql = tmp_sql + " 			,* ";
            tmp_sql = tmp_sql + " 		FROM ";
            tmp_sql = tmp_sql + " 		( ";
            tmp_sql = tmp_sql + " 			SELECT ";
            tmp_sql = tmp_sql + " 					CONVERT(varchar,setubi_grpsortorder) + '-' + ";
            tmp_sql = tmp_sql + " 					CONVERT(varchar,setubi_sortorder) + '-' + ";
            tmp_sql = tmp_sql + " 					CONVERT(varchar,komok_sortorder) AS [キー] ";
            tmp_sql = tmp_sql + " 				,komok_guid AS [項目guid] ";
            tmp_sql = tmp_sql + " 			/* ";
            tmp_sql = tmp_sql + " 				,SE.setubi_guid ";
            tmp_sql = tmp_sql + " 				,komok_guid  ";
            tmp_sql = tmp_sql + " 				,setubi_grpsortorder ";
            tmp_sql = tmp_sql + " 				,setubi_grpname  ";
            tmp_sql = tmp_sql + " 				,setubi_sortorder ";
            tmp_sql = tmp_sql + " 				,setubi_name ";
            tmp_sql = tmp_sql + " 				,komok_sortorder ";
            tmp_sql = tmp_sql + " 				,komok_name ";
            tmp_sql = tmp_sql + " 			*/ ";
            tmp_sql = tmp_sql + " 			FROM m_setubi_grp AS SEG ";
            tmp_sql = tmp_sql + " 			LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid ";
            tmp_sql = tmp_sql + " 			LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid ";
            tmp_sql = tmp_sql + " 		) AS VW1 ";
            tmp_sql = tmp_sql + " 	) AS VW2 ";
            tmp_sql = tmp_sql + " 	WHERE [抽出用] = 1 ";
            tmp_sql = tmp_sql + " ) AS VW3 ";
            tmp_sql = tmp_sql + " UNION SELECT '999-1' AS [キー],'1' AS [項目guid] ";
            tmp_sql = tmp_sql + " UNION SELECT '999-2' AS [キー],'2' AS [項目guid] ";
            tmp_sql = tmp_sql + " UNION SELECT '999-3' AS [キー],'3' AS [項目guid] ";
            tmp_sql = tmp_sql + " /*20160829 エレベーター移行対応 chg end*/ ";

            // 初期化
            CommonModule.Hash_SetubiMst.Clear();

            // キーNoと項目guidを紐付けたハッシュテーブルを作成
            DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref CommonModule.Hash_SetubiMst);

        }

        /// <summary>
        /// 10設備マスタから項目guidを取得してハッシュテーブルへ格納(ユーザー作成設備分) '20160829 設備の新規挿入処理を追加 -add
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <remarks></remarks>
        public static void Set_SetubiMst_UserMake(SqlConnection sqlcnnv10)
        {

            string tmp_sql = "";

            // 抽出クエリの作成
            tmp_sql = tmp_sql + " SELECT * FROM ";
            tmp_sql = tmp_sql + " ( ";
            tmp_sql = tmp_sql + "     SELECT [キー],CONVERT(varchar(MAX),[項目guid]) AS [項目guid] FROM ";
            tmp_sql = tmp_sql + "     ( ";
            tmp_sql = tmp_sql + "         SELECT ";
            tmp_sql = tmp_sql + "             	ROW_NUMBER()OVER(PARTITION BY [キー] ORDER BY [キー]) AS [抽出用] ";
            tmp_sql = tmp_sql + "             ,* ";
            tmp_sql = tmp_sql + "         FROM ";
            tmp_sql = tmp_sql + "         ( ";
            tmp_sql = tmp_sql + "             SELECT ";
            tmp_sql = tmp_sql + "             	 CONVERT(varchar,setubi_name) + '-' + ";
            tmp_sql = tmp_sql + "             	 CONVERT(varchar,komok_name) AS [キー] ";
            tmp_sql = tmp_sql + "             	,komok_guid AS [項目guid] ";
            tmp_sql = tmp_sql + "             FROM m_setubi_grp AS SEG ";
            tmp_sql = tmp_sql + "             LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid ";
            tmp_sql = tmp_sql + "             LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid ";
            tmp_sql = tmp_sql + " 			WHERE setubi_grpname = 'ユーザー作成設備' ";
            tmp_sql = tmp_sql + "         ) AS VW1 ";
            tmp_sql = tmp_sql + "     ) AS VW2 ";
            tmp_sql = tmp_sql + "     WHERE [抽出用] = 1 ";
            tmp_sql = tmp_sql + " ) AS VW3 ";

            // 初期化
            CommonModule.Hash_SetubiMst_UserMake.Clear();

            // キーNoと項目guidを紐付けたハッシュテーブルを作成
            DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref CommonModule.Hash_SetubiMst_UserMake);

        }

        /// <summary>
        /// 10鍵タイトルマスタから鍵Noとタイトル名を取得してハッシュテーブルへ格納 2016.04.06 物件鍵情報の移行処理修正
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <remarks></remarks>
        public static void Set_KagiTitleMst(SqlConnection sqlcnnv10, int kagikbn)
        {

            string tmp_sql = "";

            // 抽出クエリの作成
            tmp_sql = tmp_sql + " SELECT kagi_name,kagi_no FROM m_kagi_title ";
            tmp_sql = tmp_sql + " WHERE kagi_kbn = " + kagikbn.ToString();
            tmp_sql = tmp_sql + " ORDER BY kagi_no ";

            switch (kagikbn)
            {
                case 1:
                    {
                        // 初期化
                        CommonModule.Hash_KagiTitleKyoyo.Clear();
                        // キーNoと項目guidを紐付けたハッシュテーブルを作成
                        DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref CommonModule.Hash_KagiTitleKyoyo);
                        break;
                    }
                case 2:
                    {
                        // 初期化
                        CommonModule.Hash_KagiTitleSenyo.Clear();
                        // キーNoと項目guidを紐付けたハッシュテーブルを作成
                        DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref CommonModule.Hash_KagiTitleSenyo);
                        break;
                    }
            }

        }

        /// <summary>
        /// 10クレーム分類マスタから分類Noと分類名を取得してハッシュテーブルへ格納 '20160524 クレーム情報移行処理実装
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <remarks></remarks>
        public static void Set_ClaimBruiMst(SqlConnection sqlcnnv10, int claimkbn)
        {

            string tmp_sql = "";

            // 抽出クエリの作成
            tmp_sql = tmp_sql + " SELECT claim_name,claim_ruino FROM m_claim_rui ";
            tmp_sql = tmp_sql + " WHERE claim_ruikbn = " + claimkbn.ToString();
            tmp_sql = tmp_sql + " ORDER BY claim_ruino ";

            switch (claimkbn)
            {
                case 1:
                    {
                        // 初期化
                        CommonModule.Hash_ClaimBruiKasyo.Clear();
                        // キーNoと項目guidを紐付けたハッシュテーブルを作成
                        DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref CommonModule.Hash_ClaimBruiKasyo);
                        break;
                    }
                case 2:
                    {
                        // 初期化
                        CommonModule.Hash_ClaimBruiClaim.Clear();
                        // キーNoと項目guidを紐付けたハッシュテーブルを作成
                        DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref CommonModule.Hash_ClaimBruiClaim);
                        break;
                    }
            }

        }

        /// <summary>
        /// Trim処理を行い文字列として返却する 20160413 ユーザーデータテストでの修正
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Get_ShapeStr(object value)
        {

            string rtn_str = "";
            string tmp_str = Conversions.ToString(Interaction.IIf(value is null, "", value));

            tmp_str = tmp_str.ToString().Trim();

            rtn_str = tmp_str;

            return rtn_str;

        }

        /// <summary>
        /// 抽出クエリの抽出項目に各文字列置換関数を付加して返却
        /// </summary>
        /// <param name="fldnamegrp">抽出対象をカンマ区切りで連結した文字列</param>
        /// <param name="deleteflg">改行除去判別フラグ True:除去 False:改行用置換文字列に変換</param>
        /// <param name="commaflg">カンマ変換フラグ True:全角変換 False:何もしない</param>
        /// <returns>置換関数を付加した抽出対象の連結文字列</returns>
        /// <remarks>
        /// 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加
        /// 　引数に Optional commaflg As Boolean = False を追加
        /// 　CSV出力時の抽出クエリはカンマを全角変換するため追加したフラグで条件分岐を行う
        /// </remarks>
        public static string Get_FldnameAddReplaceMethod(string fldnamegrp, bool deleteflg, bool commaflg = false)
        {

            string rtn_str = "";
            string tmp_chgfldname = "";
            string[] tmp_fldname = fldnamegrp.Split(',');
            string replacestr = "";

            // 置換文字列設定 (引数のフラグで制御)
            if (deleteflg)
            {
                replacestr = "";
            }
            else
            {
                replacestr = CommonModule.LINE_BREAK;
            }

            for (int cntii = 0, loopTo = Information.UBound(tmp_fldname); cntii <= loopTo; cntii++)
            {
                // 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg sta
                // ↓↓↓旧srcコメントアウト↓↓↓
                // Dim tmp_str As String = tmp_fldname(cntii)
                // 'tmp_str = "REPLACE(" & tmp_str & ", CHAR(13) + CHAR(10), '" & replacestr & "') AS " & tmp_str  '改行が残存している箇所が確認できたため変更
                // '2016.04.26 メインの方へも反映させる修正 -chg sta
                // '改行に加えてタブも確認できたため除去する
                // 'tmp_str = "REPLACE(REPLACE(" & tmp_str & ", CHAR(13),'" & replacestr & "'), CHAR(10),'') AS " & tmp_str
                // tmp_str = "REPLACE(REPLACE(REPLACE(" & tmp_str & ", CHAR(13),'" & replacestr & "'), CHAR(10),''), CHAR(9),'') AS " & tmp_str
                // '2016.04.26 メインの方へも反映させる修正 -chg end
                // tmp_chgfldname = tmp_chgfldname & "," & tmp_str
                // ↑↑↑旧srcコメントアウト↑↑↑

                // ※コメントが大きくなったのでまとめてコメントアウトしています。
                string tmp_str = tmp_fldname[cntii];
                if (commaflg)
                {
                    tmp_str = "NULLIF(RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(" + tmp_str + ", CHAR(13),'" + replacestr + "'), CHAR(10),''), CHAR(9),''),'''','’'),',','，')),SPACE(0))";
                }
                else
                {
                    tmp_str = "NULLIF(RTRIM(REPLACE(REPLACE(REPLACE(REPLACE(" + tmp_str + ", CHAR(13),'" + replacestr + "'), CHAR(10),''), CHAR(9),''),'''','’')),SPACE(0))";
                }
                tmp_str = tmp_str + " AS " + tmp_fldname[cntii];
                tmp_chgfldname = tmp_chgfldname + "," + tmp_str;
                // 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg end
            }

            tmp_chgfldname = tmp_chgfldname.Remove(0, 1);

            rtn_str = tmp_chgfldname;

            return rtn_str;

        }

        /// <summary>
        /// ハッシュテーブルのキーを半角変換して返す '2016.04.26 メインの方へも反映させる修正
        /// </summary>
        /// <param name="hash"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static SafeDictionary<string, string> Get_HashKeyChg(SafeDictionary<string, string> hash)
        {

            var rtn_hash = new SafeDictionary<string, string>();

            // 20160819 部屋の半角変換処理によるエラー修正 -chg sta
            // For Each item In hash

            // Dim tmp_key As String = item.Key
            // Dim tmp_value As String = item.Value

            // tmp_key = StrConv(tmp_key, VbStrConv.Narrow)

            // rtn_hash.Add(tmp_key, tmp_value)

            // Next
            rtn_hash = hash;
            // 20160819 部屋の半角変換処理によるエラー修正 -chg end

            return rtn_hash;

        }

        /// <summary>
        /// ハッシュテーブルのキーを取得してリストオブジェクトへ格納 '2016.04.26 メインの方へも反映させる修正
        /// </summary>
        /// <param name="hash"></param>
        /// <param name="list_basedata"></param>
        /// <remarks></remarks>
        public static void Set_HashKeyToList(SafeDictionary<string, string> hash, ref List<string> list_basedata)
        {

            foreach (var item in hash)
            {

                string tmp_key = Conversions.ToString(item.Key);
                list_basedata.Add(tmp_key);

            }

        }

        /// <summary>
        /// 画像ファイルの拡張子のリスト作成 '20160525 クレーム関連ファイルの画像判別処理実装
        /// </summary>
        /// <remarks></remarks>
        public static void Set_ImportableImageFileAttributesList()
        {

            // 画像ファイルの拡張子のリスト作成
            CommonModule.List_ImportableImageFileAttributes.Add(".jpg");
            CommonModule.List_ImportableImageFileAttributes.Add(".jpeg");
            CommonModule.List_ImportableImageFileAttributes.Add(".png");
            CommonModule.List_ImportableImageFileAttributes.Add(".gif");
            CommonModule.List_ImportableImageFileAttributes.Add(".tif");
            CommonModule.List_ImportableImageFileAttributes.Add(".tiff");
            CommonModule.List_ImportableImageFileAttributes.Add(".wmf");
            CommonModule.List_ImportableImageFileAttributes.Add(".bmp");

        }

        // 20160707 共通処理へ移動 -add sta
        // 移動に伴い Private → Public Shared へ変更
        /// <summary>
        /// ファイル有無チェック
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_FileExist(string filepath)
        {

            if (!File.Exists(filepath))
            {
                return false;
            }
            return true;

        }

        /// <summary>
        /// フォルダ有無チェック 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_DirExist(string dirpath)
        {

            if (!Directory.Exists(dirpath))
            {
                return false;
            }
            return true;

        }

        /// <summary>
        /// ファイルオープンチェック '20160825 ファイルが開かれているかチェックする機能を追加 -add
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Chk_FileOpen(string filepath)
        {

            string filename = Path.GetFileName(filepath);
            string filekaku = Path.GetExtension(filepath);
            string dirpath = filepath.Replace(filename, "");
            string chgfilename = Set_Path(dirpath, "temp" + filekaku);

            // 念の為
            if (Chk_FileExist(filepath) == false)
            {

            }

            try
            {
                // ファイル名を変更して、使用中かチェックする
                File.Move(filepath, chgfilename);
                // ファイル名を元に戻す
                File.Move(chgfilename, filepath);
                // ファイル名の変更が成功したので、使用中ではない
                return true;
            }
            catch (Exception ex)
            {
                // ファイル名が変更できないので、使用中とする
                return false;
            }


            return true;

        }

        /// <summary>
        /// ファイル/フォルダのフルパスを設定  
        /// </summary>
        /// <param name="path"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Set_Path(string path, string name)
        {

            string rtn_path = "";

            if (RightStrMethod(path, 1) == @"\")
            {
                rtn_path = path + name;
            }
            else
            {
                rtn_path = path + @"\" + name;
            }

            return rtn_path;

        }
        // 20160707 共通処理へ移動 -add end

        /// <summary>
        /// 親ヘッダーと子ヘッダーを結合してオブジェクトへ格納 20160913_2 部屋設備移行処理の追加
        /// </summary>
        /// <param name="objmain"></param>
        /// <param name="objsub"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static object Set_Header(string[,] objmain, string[,] objsub)
        {

            var rtn_arry = new string[1, Conversions.ToInteger(objmain.Length + 1)];

            for (int cntii = 1, loopTo = Conversions.ToInteger(objmain.Length); cntii <= loopTo; cntii++)
            {

                string tmp_main = Conversions.ToString(Interaction.IIf(objmain[1, cntii] is null, "", objmain[1, cntii]));
                string tmp_sub = Conversions.ToString(Interaction.IIf(objsub[1, cntii] is null, "", objsub[1, cntii]));

                string tmp_total = "";
                if (!string.IsNullOrEmpty(tmp_main) & !string.IsNullOrEmpty(tmp_sub))
                {
                    tmp_total = tmp_main + CommonModule.STR_SPLIT_1 + tmp_sub;
                }
                else if (!string.IsNullOrEmpty(tmp_main) & string.IsNullOrEmpty(tmp_sub))
                {
                    tmp_total = tmp_main;
                }
                else if (string.IsNullOrEmpty(tmp_main) & !string.IsNullOrEmpty(tmp_sub))
                {
                    tmp_total = tmp_sub;
                }
                else if (string.IsNullOrEmpty(tmp_main) & string.IsNullOrEmpty(tmp_sub))
                {
                    tmp_total = "";
                }

                rtn_arry[0, Conversions.ToInteger(cntii)] = tmp_total;

            }

            return rtn_arry;

        }

        /// <summary>
        /// 汎用中間ファイルのヘッダー取得処理 '20161006 中間ファイルチェック処理の速度改善対応 -add
        /// グループ名と設備項目名を
        /// </summary>
        /// <param name="solist_grp"></param>
        /// <param name="solist_komk"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static SortedList<int, string> Set_SetubiHeader(SortedList<int, string> solist_grp, ref SortedList<int, string> solist_komk)
        {

            var rtn_solist = new SortedList<int, string>();

            for (int cntii = 0, loopTo = solist_grp.Count - 1; cntii <= loopTo; cntii++)
            {

                string tmp_main = Conversions.ToString(Interaction.IIf(solist_grp[cntii] is null, "", solist_grp[cntii]));
                string tmp_sub = Conversions.ToString(Interaction.IIf(solist_komk[cntii] is null, "", solist_komk[cntii]));

                string tmp_total = "";
                if (!string.IsNullOrEmpty(tmp_main) & !string.IsNullOrEmpty(tmp_sub))
                {
                    tmp_total = tmp_main + CommonModule.STR_SPLIT_1 + tmp_sub;
                }
                else if (!string.IsNullOrEmpty(tmp_main) & string.IsNullOrEmpty(tmp_sub))
                {
                    tmp_total = tmp_main;
                }
                else if (string.IsNullOrEmpty(tmp_main) & !string.IsNullOrEmpty(tmp_sub))
                {
                    tmp_total = tmp_sub;
                }
                else if (string.IsNullOrEmpty(tmp_main) & string.IsNullOrEmpty(tmp_sub))
                {
                    tmp_total = "";
                }

                rtn_solist.Add(cntii, tmp_total);

            }

            return rtn_solist;

        }




        /// <summary>
        /// 10設備マスタから項目guidを取得してハッシュテーブルへ格納(汎用コンバート用) '20160913_2 部屋設備移行処理の追加
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <remarks></remarks>
        public static void Set_SetubiMst_BaseMId(SqlConnection sqlcnnv10, ref SafeDictionary<string, string> hash)
        {

            string tmp_sql = "";
            tmp_sql = tmp_sql + " SELECT  ";
            tmp_sql = tmp_sql + " 	 SEG.setubi_grpname + '@#@' + setubi_name + '@#@' + CONVERT(VARCHAR,komok_sortorder) ";
            tmp_sql = tmp_sql + " 	,CONVERT(VARCHAR(MAX),komok_guid) ";
            tmp_sql = tmp_sql + " FROM m_setubi_grp AS SEG ";
            tmp_sql = tmp_sql + " LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid ";
            tmp_sql = tmp_sql + " LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid ";

            // 初期化
            hash.Clear();

            // キーNoと項目guidを紐付けたハッシュテーブルを作成
            DBExec.Exec_DataReader_Col_Hash(tmp_sql, ref sqlcnnv10, ref hash);

        }

    }

    #endregion

    #region 共通クエリクラス

    /// <summary>
    /// 【共通クエリクラス】
    /// </summary>
    /// <remarks></remarks>
    public class DBQuery
    {

        /// <summary>
        /// 削除(DELETE)クエリ
        /// </summary>
        /// <param name="reptbl">FROM句(例：m_bk_rui)</param>
        /// <param name="repwhere">WHERE句 [optional]</param>
        /// <returns>返却クエリ</returns>
        /// <remarks>
        /// ・DELETE FROM [reptbl] WHERE [repwhere]
        /// </remarks>
        public static string Qry_DelInfo(string reptbl, string repwhere = "")


        {
            // ★★★ CommonRepository.Get_UseQueryを参考
            string qry_rtn;                                   // 返却クエリ文
            string qry_tblname = "";                          // FROM句
            string qry_where = "";                            // WHERE句

            qry_tblname = " DELETE FROM " + reptbl;
            qry_where = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(repwhere.Trim()), "", " WHERE " + repwhere));
            qry_rtn = qry_tblname + qry_where;
            return qry_rtn;

        }

        /// <summary>
        /// テーブル/ビュー削除(DROP)クエリ
        /// </summary>
        /// <param name="targetname"></param>
        /// <param name="tableflg"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Qry_DropInfo(string targetname, bool tableflg)
        {

            string droptarget = "";
            string qry_rtn = "";

            // 削除対象選択
            if (tableflg)
            {
                droptarget = "TABLE";
            }
            else
            {
                droptarget = "VIEW";
            }

            qry_rtn = qry_rtn + "  IF EXISTS ";
            qry_rtn = qry_rtn + " 	( ";
            qry_rtn = qry_rtn + " 		SELECT * FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[" + targetname + "]') ";
            qry_rtn = qry_rtn + " 		AND OBJECTPROPERTY(id, N'Is" + droptarget + "') = 1 ";
            qry_rtn = qry_rtn + " 	) ";
            qry_rtn = qry_rtn + " DROP " + droptarget + " [dbo].[" + targetname + "]; ";

            return qry_rtn;

        }

        /// <summary>
        /// ストアド削除(DROP)クエリ '20160829 口座名義カナチェック機能の追加
        /// </summary>
        /// <param name="targetname"></param>
        /// <param name="tableflg"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Qry_DropFnInfo(string targetname)
        {

            string qry_rtn = "";
            qry_rtn = qry_rtn + " IF OBJECT_ID (N'[dbo].[" + targetname + "]', N'FN') IS NOT NULL ";
            qry_rtn = qry_rtn + " DROP FUNCTION [dbo].[" + targetname + "]; ";
            return qry_rtn;

        }

        /// <summary>
        /// 挿入クエリ(パラメータで記載)
        /// </summary>
        /// <param name="tblname"></param>
        /// <param name="fldnamegrp"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Qry_Insert_ParameterSet(string tblname, string fldnamegrp, ref string[] para)
        {

            string tmp_para = "";

            // 取得したフィールド名(カンマで連結)を分割して配列へ格納
            string[] tmp_str = fldnamegrp.Split(',');

            // テーブル名と各フィールドを「-」で連結してカンマ区切りで連結した文字列に変換
            for (int cntii = 0, loopTo = Information.UBound(tmp_str); cntii <= loopTo; cntii++)
                // tmp_para = tmp_para & ",@" & tblname & "-" & fldname(cntii)
                tmp_para = tmp_para + ",@" + tmp_str[cntii];
            tmp_para = tmp_para.Remove(0, 1);

            // '取得したテーブル/フィールド名("tbl-fld")を分割して配列へ格納
            // tblfldname = tmp_para.Split(",")

            // クエリ作成
            string tmp_sql = "INSERT INTO " + tblname + " (" + fldnamegrp + ") VALUES (" + tmp_para + ")";

            return tmp_sql;

        }

        /// <summary>
        /// テーブル列数取得クエリ
        /// </summary>
        /// <param name="tblname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Qry_GetColCount(string tblname)
        {

            string qry_rtn = "";

            qry_rtn = qry_rtn + " SELECT ";
            qry_rtn = qry_rtn + "	COUNT(syscolumns.name) ";
            qry_rtn = qry_rtn + " FROM ";
            qry_rtn = qry_rtn + "	syscolumns ";
            qry_rtn = qry_rtn + " INNER JOIN sysobjects ON ";
            qry_rtn = qry_rtn + "	sysobjects.id = syscolumns.id ";
            qry_rtn = qry_rtn + " WHERE ";
            qry_rtn = qry_rtn + "	sysobjects.name = '";
            qry_rtn = qry_rtn + tblname;
            qry_rtn = qry_rtn + "'";

            return qry_rtn;

        }

        /// <summary>
        /// テーブル列数取得クエリ
        /// </summary>
        /// <param name="tblname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Qry_GetRowCount(string tblname)
        {

            string qry_rtn = "";

            qry_rtn = qry_rtn + " SELECT COUNT(*) FROM ";
            qry_rtn = qry_rtn + tblname;

            return qry_rtn;

        }

        /// <summary>
        /// 革命10DBフィールド情報取得クエリ
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        // Public Shared Function Qry_GetTableInfo() As String

        // CVDBInfoModuleへ移動


        // Dim tmp_sql As String = ""

        // '速度改善のためクエリの結合を減らす

        // 'tmp_sql = tmp_sql & " SELECT "
        // 'tmp_sql = tmp_sql & " 	 VW.* "
        // 'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用種別,'') AS CV用種別 "
        // 'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用項目名,'') AS CV用項目名 "
        // 'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用最小値,'') AS CV用最小値 "
        // 'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用最大値,'') AS CV用最大値 "
        // 'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用キーフラグ,'') AS CV用キーフラグ "
        // 'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用必須項目フラグ,'') AS CV用必須項目フラグ "
        // 'tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV備考,'') AS CV備考 "
        // 'tmp_sql = tmp_sql & " FROM "
        // 'tmp_sql = tmp_sql & " ( "
        // 'tmp_sql = tmp_sql & " 	SELECT "
        // 'tmp_sql = tmp_sql & " 		 col.column_id AS [行No] "
        // 'tmp_sql = tmp_sql & " 		,CASE "
        // 'tmp_sql = tmp_sql & " 			WHEN index_column_id IS NULL THEN '' "
        // 'tmp_sql = tmp_sql & " 			ELSE '*' "
        // 'tmp_sql = tmp_sql & " 		 END AS [主キー] "
        // 'tmp_sql = tmp_sql & " 		,jpnname.value AS [種別] "
        // 'tmp_sql = tmp_sql & " 		,obj.NAME AS [TBL名] "
        // 'tmp_sql = tmp_sql & " 		,fldjpname.value AS [項目名] "
        // 'tmp_sql = tmp_sql & " 		,col.NAME AS [フィールド名] "
        // 'tmp_sql = tmp_sql & " 		,type_name(col.user_type_id) AS [属性] "
        // 'tmp_sql = tmp_sql & " 		,col.max_length AS [サイズ] "
        // 'tmp_sql = tmp_sql & " 		,CASE is_nullable "
        // 'tmp_sql = tmp_sql & " 			WHEN '1' THEN '○' "
        // 'tmp_sql = tmp_sql & " 			ELSE '' "
        // 'tmp_sql = tmp_sql & " 		 END AS [Null許容] "
        // 'tmp_sql = tmp_sql & " 		,fldcomment.value AS [説明] "
        // 'tmp_sql = tmp_sql & " 		,obj.id "
        // 'tmp_sql = tmp_sql & " 		,creater.value AS creator "
        // 'tmp_sql = tmp_sql & " 		,checker.value AS checker "
        // 'tmp_sql = tmp_sql & " 		,comment.value  AS [TBL説明] "
        // 'tmp_sql = tmp_sql & " 	FROM sys.sysobjects AS obj "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS jpnname ON obj.id = jpnname.major_id "
        // 'tmp_sql = tmp_sql & " 		AND jpnname.NAME = 'MS_Description' "
        // 'tmp_sql = tmp_sql & " 		AND jpnname.minor_id = 0 "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS creater ON obj.id = creater.major_id "
        // 'tmp_sql = tmp_sql & " 		AND creater.NAME = 'creator' "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS checker ON obj.id = checker.major_id "
        // 'tmp_sql = tmp_sql & " 		AND checker.NAME = 'checker' "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS comment ON obj.id = comment.major_id "
        // 'tmp_sql = tmp_sql & " 		AND comment.NAME = 'comment' "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.columns AS col ON obj.id = col.object_id "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS fldjpname ON col.object_id = fldjpname.major_id "
        // 'tmp_sql = tmp_sql & " 		AND col.column_id = fldjpname.minor_id "
        // 'tmp_sql = tmp_sql & " 		AND fldjpname.NAME = 'Jpfieldname' "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS fldcomment ON col.object_id = fldcomment.major_id "
        // 'tmp_sql = tmp_sql & " 		AND col.column_id = fldcomment.minor_id "
        // 'tmp_sql = tmp_sql & " 		AND fldcomment.NAME = 'MS_Description' "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.indexes AS I ON col.object_id = I.object_id "
        // 'tmp_sql = tmp_sql & " 		AND I.is_primary_key = 1 "
        // 'tmp_sql = tmp_sql & " 	LEFT JOIN sys.index_columns AS pkey ON I.object_id = pkey.object_id "
        // 'tmp_sql = tmp_sql & " 		AND col.column_id = pkey.column_id "
        // 'tmp_sql = tmp_sql & " 		AND I.index_id = pkey.index_id "
        // 'tmp_sql = tmp_sql & " 	WHERE obj.xtype = 'U' "
        // 'tmp_sql = tmp_sql & " ) AS VW "
        // 'tmp_sql = tmp_sql & " LEFT JOIN cv_dbinfo AS CVDB "
        // 'tmp_sql = tmp_sql & " ON  VW.TBL名 = CVDB.紐付用TBL名 "
        // 'tmp_sql = tmp_sql & " AND VW.フィールド名 = CVDB.紐付用フィールド名 "
        // 'tmp_sql = tmp_sql & " ORDER BY [TBL名],id,[行No] "

        // tmp_sql = tmp_sql & " SELECT "
        // tmp_sql = tmp_sql & " 	 VW.* "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用種別,'') AS CV用種別 "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用項目名,'') AS CV用項目名 "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用最小値,'') AS CV用最小値 "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用最大値,'') AS CV用最大値 "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.デフォルト値,'') AS デフォルト値 "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用キーフラグ,'') AS CV用キーフラグ "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV用必須項目フラグ,'') AS CV用必須項目フラグ "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.革命マスタ参照区分,'') AS 革命マスタ参照区分 "
        // tmp_sql = tmp_sql & " 	,ISNULL(CVDB.CV備考,'') AS CV備考 "
        // tmp_sql = tmp_sql & " FROM "
        // tmp_sql = tmp_sql & " ( "
        // tmp_sql = tmp_sql & " 	SELECT "
        // tmp_sql = tmp_sql & " 		 col.column_id AS [行No] "
        // tmp_sql = tmp_sql & " 		,CASE "
        // tmp_sql = tmp_sql & " 			WHEN index_column_id IS NULL THEN '' "
        // tmp_sql = tmp_sql & " 			ELSE '*' "
        // tmp_sql = tmp_sql & " 		 END AS [主キー] "
        // tmp_sql = tmp_sql & " 		,jpnname.value AS [種別] "
        // tmp_sql = tmp_sql & " 		,obj.NAME AS [TBL名] "
        // tmp_sql = tmp_sql & " 		,fldjpname.value AS [項目名] "
        // tmp_sql = tmp_sql & " 		,col.NAME AS [フィールド名] "
        // tmp_sql = tmp_sql & " 		,type_name(col.user_type_id) AS [属性] "
        // tmp_sql = tmp_sql & " 		,col.max_length AS [サイズ] "
        // tmp_sql = tmp_sql & " 		,CASE is_nullable "
        // tmp_sql = tmp_sql & " 			WHEN '1' THEN '○' "
        // tmp_sql = tmp_sql & " 			ELSE '' "
        // tmp_sql = tmp_sql & " 		 END AS [Null許容] "
        // tmp_sql = tmp_sql & " 		 ,obj.id "
        // tmp_sql = tmp_sql & " 		 /* "
        // tmp_sql = tmp_sql & " 		,fldcomment.value AS [説明] "
        // tmp_sql = tmp_sql & " 		,creater.value AS creator "
        // tmp_sql = tmp_sql & " 		,checker.value AS checker "
        // tmp_sql = tmp_sql & " 		,comment.value  AS [TBL説明] "
        // tmp_sql = tmp_sql & " 		*/ "
        // tmp_sql = tmp_sql & " 	FROM sys.sysobjects AS obj "
        // tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS jpnname ON obj.id = jpnname.major_id "
        // tmp_sql = tmp_sql & " 		AND jpnname.NAME = 'MS_Description' "
        // tmp_sql = tmp_sql & " 		AND jpnname.minor_id = 0 "
        // tmp_sql = tmp_sql & " 	LEFT JOIN sys.columns AS col ON obj.id = col.object_id "
        // tmp_sql = tmp_sql & " 	LEFT JOIN sys.extended_properties AS fldjpname ON col.object_id = fldjpname.major_id "
        // tmp_sql = tmp_sql & " 		AND col.column_id = fldjpname.minor_id "
        // tmp_sql = tmp_sql & " 		AND fldjpname.NAME = 'Jpfieldname' "
        // tmp_sql = tmp_sql & " 	LEFT JOIN sys.indexes AS I ON col.object_id = I.object_id "
        // tmp_sql = tmp_sql & " 		AND I.is_primary_key = 1 "
        // tmp_sql = tmp_sql & " 	LEFT JOIN sys.index_columns AS pkey ON I.object_id = pkey.object_id "
        // tmp_sql = tmp_sql & " 		AND col.column_id = pkey.column_id "
        // tmp_sql = tmp_sql & " 		AND I.index_id = pkey.index_id "
        // tmp_sql = tmp_sql & " 	WHERE obj.xtype = 'U' "
        // tmp_sql = tmp_sql & " ) AS VW "
        // tmp_sql = tmp_sql & " LEFT JOIN cv_dbinfo AS CVDB "
        // tmp_sql = tmp_sql & " ON  VW.TBL名 = CVDB.紐付用TBL名 "
        // tmp_sql = tmp_sql & " AND VW.フィールド名 = CVDB.紐付用フィールド名 "
        // tmp_sql = tmp_sql & " ORDER BY [TBL名],id,[行No] "

        // Return tmp_sql

        // End Function

        /// <summary>
        /// 金融機関情報取得クエリ
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Qry_GetKinyuInfo()
        {

            string tmp_sql = "";

            tmp_sql = tmp_sql + " SELECT ";
            tmp_sql = tmp_sql + " 	 KINYU.kinyu_no ";
            tmp_sql = tmp_sql + " 	,KINYU.kinyu_name ";
            tmp_sql = tmp_sql + " 	,TEN.kinyu_tenno ";
            tmp_sql = tmp_sql + " 	,TEN.kinyu_tenname ";
            tmp_sql = tmp_sql + " FROM m_kinyu_ten AS TEN ";
            tmp_sql = tmp_sql + " LEFT JOIN m_kinyu AS KINYU ON TEN.kinyu_no = KINYU.kinyu_no ";
            tmp_sql = tmp_sql + " ORDER BY KINYU.kinyu_no,TEN.kinyu_tenno ";

            return tmp_sql;

        }

    }

    #endregion

    #region 履歴関連クラス

    public class HistorySetting
    {

        public string Created { get; set; }
        public string CreatedHostName { get; set; }
        public string CreatedOSUserName { get; set; }
        public string CreatedAppUserName { get; set; }
        public string Updated { get; set; }
        public string UpdatedHostName { get; set; }
        public string UpdatedOSUserName { get; set; }
        public string UpdatedAppUserName { get; set; }

        /// <summary>
        /// 履歴出力内容をセット
        /// </summary>
        /// <param name="recuser"></param>
        /// <remarks></remarks>
        public void Set_HistoryData(string recuser)
        {

            Created = Conversions.ToString(DateTime.Now);
            CreatedHostName = recuser;
            CreatedOSUserName = CreatedHostName;
            CreatedAppUserName = "標準ユーザ";
            Updated = Created;
            UpdatedHostName = CreatedHostName;
            UpdatedOSUserName = CreatedOSUserName;
            UpdatedAppUserName = CreatedAppUserName;

        }

        /// <summary>
        /// タグ付きデータを作成
        /// </summary>
        /// <remarks></remarks>
        public string ToXmlString()
        {

            var sb = new System.Text.StringBuilder();
            var xw = System.Xml.XmlWriter.Create(sb);
            xw.WriteStartDocument();
            xw.WriteStartElement("history");

            // 登録日
            if (Typ.IsDateNotNull(Created) == true)
            {
                xw.WriteStartElement("c");
                xw.WriteString(Created);
                xw.WriteEndElement();
            }

            // 登録者
            if (!string.IsNullOrEmpty(CreatedHostName))
            {
                xw.WriteStartElement("chost");
                xw.WriteString(CreatedHostName);
                xw.WriteEndElement();
            }

            // 登録者OSユーザ名
            if (!string.IsNullOrEmpty(CreatedOSUserName))
            {
                xw.WriteStartElement("cosuser");
                xw.WriteString(CreatedOSUserName);
                xw.WriteEndElement();
            }

            // 権限
            if (!string.IsNullOrEmpty(CreatedAppUserName))
            {
                xw.WriteStartElement("cappuser");
                xw.WriteString(CreatedAppUserName);
                xw.WriteEndElement();
            }

            // 更新日
            if (Typ.IsDateNotNull(Updated) == true)
            {
                xw.WriteStartElement("u");
                xw.WriteString(Updated);
                xw.WriteEndElement();
            }

            // 更新者名
            if (!string.IsNullOrEmpty(UpdatedHostName))
            {
                xw.WriteStartElement("uhost");
                xw.WriteString(UpdatedHostName);
                xw.WriteEndElement();
            }

            // 更新者OSユーザ名
            if (!string.IsNullOrEmpty(UpdatedOSUserName))
            {
                xw.WriteStartElement("uosuser");
                xw.WriteString(UpdatedOSUserName);
                xw.WriteEndElement();
            }

            // 更新者権限
            if (!string.IsNullOrEmpty(UpdatedAppUserName))
            {
                xw.WriteStartElement("uappuser");
                xw.WriteString(UpdatedAppUserName);
                xw.WriteEndElement();
            }

            xw.WriteEndElement();
            xw.Flush();
            xw.Close();

            string str_xml = sb.ToString();

            return str_xml;

        }

    }

    #endregion

    #region 色変換クラス

    public class Translate_Color
    {

        /// <summary>
        /// RGB → ARGB 変換
        /// </summary>
        /// <param name="int"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public int Chg_ColorValue(int @int)
        {

            var color = new Color();

            // RGB → ARGB 変換
            color = ColorTranslator.FromOle(@int);

            // 変換値取得
            int chgvalue = color.ToArgb();

            return chgvalue;

        }

    }

    #endregion

    #region Excelファイル設定クラス

    public class ExcelFileManager
    {
        // 20160812 汎用コンバート対応 -del sta
        // ''' <summary>
        // ''' 書込時のExcelファイルオープン処理
        // ''' </summary>
        // ''' <param name="appli"></param>
        // ''' <param name="wbook"></param>
        // ''' <param name="wsheet"></param>
        // ''' <param name="startrow"></param>
        // ''' <param name="columncnt"></param>
        // ''' <param name="maxrowcnt"></param>
        // ''' <param name="dataarea"></param>
        // ''' <param name="sheetname"></param>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        // Public Function Set_ExcelFile_WriteOpen( _
        // ByRef appli As Excel.Application, _
        // ByRef wbook As Excel.Workbook, _
        // ByRef wsheet As Excel.Worksheet, _
        // ByRef startrow As Integer, _
        // ByRef columncnt As Integer, _
        // ByRef maxrowcnt As Integer, _
        // ByRef dataarea As String, _
        // ByVal filename As String, _
        // ByVal sheetname As String _
        // ) As Boolean

        // Dim rtn As Boolean = True
        // Dim midfilepath As String = MiddleDirPath & "\" & filename & ".xlsx"

        // Try
        // 'ファイル準備
        // appli = CreateObject("Excel.Application")
        // appli.Visible = False
        // wbook = appli.Workbooks.Open(midfilepath)
        // wsheet = wbook.Worksheets(sheetname)

        // '書込開始行の設定
        // startrow = MIDFILE_READWRITE_ROW

        // '既存データの確認(最大行の取得)
        // maxrowcnt = wsheet.UsedRange.Rows.Count

        // '既存データ範囲を取得
        // dataarea = startrow & ":" & maxrowcnt

        // '初期化
        // If maxrowcnt >= startrow Then
        // wsheet.Rows(dataarea).Delete()
        // End If

        // 'ヘッダの列数取得
        // columncnt = wsheet.UsedRange.Columns.Count

        // Catch ex As Exception
        // 'エラー処理を行うこと
        // Debug.WriteLine(ex.Message)
        // rtn = False
        // End Try

        // Return rtn

        // End Function
        // 20160812 汎用コンバート対応 -del end
        /// <summary>
        /// 書込時のExcelファイルオープン処理 '20160812 汎用コンバート対応
        /// </summary>
        /// <param name="appli"></param>
        /// <param name="wbook"></param>
        /// <param name="wsheet"></param>
        /// <param name="startrow"></param>
        /// <param name="columncnt"></param>
        /// <param name="maxrowcnt"></param>
        /// <param name="dataarea"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_ExcelFile_WriteOpen(ref Microsoft.Office.Interop.Excel.Application appli, ref Microsoft.Office.Interop.Excel.Workbook wbook, ref Microsoft.Office.Interop.Excel.Worksheet wsheet, ref int startrow, ref int columncnt, ref int maxrowcnt, string dirpath, string filename, string sheetname)









        {
            return default;

            // 20170601 未使用箇所のコメントアウト -del sta
            // Dim rtn As Boolean = True
            // '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg sta
            // 'Dim midfilepath As String = MiddleDirPath & "\" & filename & ".xlsx"
            // Dim midfilepath As String = dirpath & "\" & filename & ".xlsx"
            // '20160912 既存用→汎用用中間ファイルへのコピー処理構築中に発見した不具合修正 -chg end
            // Try
            // 'ファイル準備
            // appli = CreateObject("Excel.Application")
            // appli.Visible = False
            // wbook = appli.Workbooks.Open(midfilepath)
            // wsheet = wbook.Worksheets(sheetname)

            // '書込開始行の設定
            // '20160912 中間ファイル作成に伴うプログラム修正 -chg sta
            // 'startrow = EXISTMIDFILE_READWRITE_ROW
            // If filename = CV_FROM_MIDDLE Then
            // startrow = BASEMIDFILE_READWRITE_ROW
            // Else
            // startrow = EXISTMIDFILE_READWRITE_ROW
            // End If
            // '20160912 中間ファイル作成に伴うプログラム修正 -chg end
            // '既存データの確認(最大行の取得)
            // maxrowcnt = wsheet.UsedRange.Rows.Count

            // '既存データ範囲を取得
            // Dim dataarea As String = startrow & ":" & maxrowcnt

            // '初期化
            // If maxrowcnt >= startrow Then
            // wsheet.Rows(dataarea).Delete()
            // End If

            // 'ヘッダの列数取得
            // columncnt = wsheet.UsedRange.Columns.Count

            // Catch ex As Exception
            // Debug.WriteLine(ex.Message)
            // rtn = False
            // End Try

            // Return rtn
            // 20170601 未使用箇所のコメントアウト -del end
        }

        /// <summary>
        /// 読込時のExcelファイルオープン処理
        /// </summary>
        /// <param name="appli"></param>
        /// <param name="wbook"></param>
        /// <param name="wsheet"></param>
        /// <param name="startrow"></param>
        /// <param name="columncnt"></param>
        /// <param name="maxrowcnt"></param>
        /// <param name="rowcnt"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_ExcelFile_ReadOpen(ref Microsoft.Office.Interop.Excel.Application appli, ref Microsoft.Office.Interop.Excel.Workbook wbook, ref Microsoft.Office.Interop.Excel.Worksheet wsheet, ref int startrow, ref int columncnt, ref int maxrowcnt, ref int rowcnt, string dirpath, string filename, string sheetname)










        {

            bool rtn = true;
            string filepath = dirpath + @"\" + filename + ".xlsx";

            try
            {
                // ファイル準備
                appli = (Microsoft.Office.Interop.Excel.Application)Interaction.CreateObject("Excel.Application");
                appli.Visible = false;
                wbook = appli.Workbooks.Open(filepath);
                wsheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets[sheetname];

                // 読込開始行の設定
                // 20160912 中間ファイル作成に伴うプログラム修正 -chg sta
                // If startrow = 0 Then
                // startrow = MIDFILE_READWRITE_ROW
                // End If
                if ((filename ?? "") == (CommonModule.CV_FROM_MIDDLE ?? ""))                       // 20160913 CNVNOで分岐ではなくファイル名で分岐したのは？(他の箇所も同様)
                {
                    startrow = CommonModule.BASEMIDFILE_READWRITE_ROW;
                }
                else
                {
                    startrow = CommonModule.EXISTMIDFILE_READWRITE_ROW;
                }
                // 20160912 中間ファイル作成に伴うプログラム修正 -chg end
                // 既存データの確認(ヘッダを含めた最大行の取得)
                maxrowcnt = wsheet.UsedRange.Rows.Count;

                // ヘッダの列数取得
                columncnt = wsheet.UsedRange.Columns.Count;

                // 行数取得(全行 - 読込開始行 + 1)
                rowcnt = maxrowcnt - startrow + 1;
            }

            catch (Exception ex)
            {
                // 20160907 EXCELが開いたままになっている不具合 -coment：エラーの場合の処理(ログ出力)がない？
                // エラー処理を行うこと
                Debug.WriteLine(ex.Message);
                rtn = false;
            }

            return rtn;


        }

        /// <summary>
        /// 中間ファイルの件数を取得する処理 20160905 汎用コンバートの件数表示修正
        /// </summary>
        /// <param name="appli"></param>
        /// <param name="wbook"></param>
        /// <param name="wsheet"></param>
        /// <param name="startrow"></param>
        /// <param name="columncnt"></param>
        /// <param name="maxrowcnt"></param>
        /// <param name="rowcnt"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_ExcelFile_ReadOpen_BaseMidCnt(ref Microsoft.Office.Interop.Excel.Application appli, ref Microsoft.Office.Interop.Excel.Workbook wbook, ref Microsoft.Office.Interop.Excel.Worksheet wsheet, ref int startrow, ref int maxrowcnt, string dirpath, string filename, SafeDictionary<Label, string> hash_sheet, ref SafeDictionary<Label, int> hash_cnt)









        {

            bool rtn = true;
            string filepath = dirpath + @"\" + filename + ".xlsx";

            try
            {
                // ファイル準備
                appli = (Microsoft.Office.Interop.Excel.Application)Interaction.CreateObject("Excel.Application");
                appli.Visible = false;
                wbook = appli.Workbooks.Open(filepath);

                foreach (var baseitem in hash_sheet)
                {

                    Label tmp_lbl = (Label)baseitem.Key;
                    string tmp_sheetname = Conversions.ToString(baseitem.Value);

                    // シート選択
                    wsheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets[tmp_sheetname];

                    // 読込開始行の設定
                    // 20160912 中間ファイル作成に伴うプログラム修正 -chg sta
                    // If tmp_sheetname = "部屋設備情報" Then
                    // startrow = 3
                    // Else
                    // startrow = EXISTMIDFILE_READWRITE_ROW
                    // End If
                    startrow = CommonModule.BASEMIDFILE_READWRITE_ROW;
                    // 20160912 中間ファイル作成に伴うプログラム修正 -chg end
                    // 既存データの確認(ヘッダを含めた最大行の取得)
                    maxrowcnt = wsheet.UsedRange.Rows.Count;

                    // 行数取得(全行 - 読込開始行 + 1)
                    int rowcnt = maxrowcnt - startrow + 1;

                    // ハッシュテーブルへ再格納
                    hash_cnt.Add(tmp_lbl, rowcnt);

                }
            }

            catch (Exception ex)
            {
                // 20160907 EXCELが開いたままになっている不具合 -coment：エラーの場合の処理(ログ出力)がない？
                // エラー処理を行うこと
                Debug.WriteLine(ex.Message);
                rtn = false;
            }

            return rtn;

        }

        /// <summary>
        /// Excelファイルクローズ処理(読込時)
        /// </summary>
        /// <param name="appli"></param>
        /// <param name="wbook"></param>
        /// <param name="wsheet"></param>
        /// <remarks></remarks>
        public void Set_ExcelFile_ReadClose(ref Microsoft.Office.Interop.Excel.Application appli, ref Microsoft.Office.Interop.Excel.Workbook wbook, ref Microsoft.Office.Interop.Excel.Worksheet wsheet)
        {

            try
            {
                wbook.Close();
                appli.Quit();
                Marshal.ReleaseComObject(appli);      // Excelオブジェクトの開放(ロック明示的解放)
                wsheet = null;
                wbook = null;
                appli = null;
            }
            catch (Exception ex)
            {
                // 20160907 EXCELが開いたままになっている不具合 -coment：エラーの場合の処理(ログ出力)がない？
                Debug.WriteLine(ex.Message);
            }

        }

        /// <summary>
        /// Excelファイルクローズ処理(書込時)
        /// </summary>
        /// <param name="appli"></param>
        /// <param name="wbook"></param>
        /// <param name="wsheet"></param>
        /// <remarks></remarks>
        public void Set_ExcelFile_WriteClose(ref Microsoft.Office.Interop.Excel.Application appli, ref Microsoft.Office.Interop.Excel.Workbook wbook, ref Microsoft.Office.Interop.Excel.Worksheet wsheet)
        {

            try
            {
                wbook.Save();
                wbook.Close();
                appli.Quit();
                // 20161014 DisconnectedContextエラー修正 -chg sta
                // System.Runtime.InteropServices.Marshal.ReleaseComObject(wbook)
                Marshal.ReleaseComObject(wbook);
                // 20161014 DisconnectedContextエラー修正 -chg end
                Marshal.ReleaseComObject(appli);
                wsheet = null;
                wbook = null;
                appli = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

        }

        /// <summary>
        /// Excel読込時のオープン処理 '20161004 既存中間→DB書込処理修正 -add
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <param name="dt"></param>
        /// <param name="con_read"></param>
        /// <param name="cmd_read"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool ExcelFile_ReadOpen(string filepath, string filename, string sql, ref DataTable dt, ref OleDbConnection con_read)
        {

            OleDbDataReader dr;
            string tmpprovider = "Microsoft.ACE.OLEDB.12.0; ";    // EXCEL2007以上(xlsx)
            string tmpextend = "Excel 8.0;HDR=YES;";
            var cmd_read = new OleDbCommand();
            string openfile = EtcMethod.Set_Path(filepath, filename + ".xlsx");

            try
            {

                // 接続文字列生成
                con_read.ConnectionString = "Provider=" + tmpprovider + "Data Source=" + openfile + ";" + "Extended Properties=" + "\"" + tmpextend + "\"";



                // 接続設定
                cmd_read.Connection = con_read;

                // 接続オープン処理
                con_read.Open();

                cmd_read = con_read.CreateCommand();
                cmd_read.CommandText = sql;
                dt = new DataTable();
                dr = cmd_read.ExecuteReader();
                dt.Load(dr);
            }

            catch (Exception ex)
            {

                return false;

            }

            return true;

        }

        /// <summary>
        /// Excel読込時のオープン処理 20161012 中間ファイル件数表示速度改善
        /// データテーブル等のオブジェクトは返さずにそのまま開くだけ
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <param name="dt"></param>
        /// <param name="con_read"></param>
        /// <param name="cmd_read"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool ExcelFile_ReadOpenOnly(string filepath, string filename, ref OleDbConnection con_read)
        {


            OleDbDataReader dr;
            string tmpprovider = "Microsoft.ACE.OLEDB.12.0; ";    // EXCEL2007以上(xlsx)
            string tmpextend = "Excel 8.0;HDR=YES;";
            var cmd_read = new OleDbCommand();
            string openfile = EtcMethod.Set_Path(filepath, filename + ".xlsx");

            try
            {

                // 接続文字列生成
                con_read.ConnectionString = "Provider=" + tmpprovider + "Data Source=" + openfile + ";" + "Extended Properties=" + "\"" + tmpextend + "\"";



                // 接続設定
                cmd_read.Connection = con_read;

                // 接続オープン処理
                con_read.Open();
            }

            // 20161104 中間ファイル読込エラー時の処理対応 -chg sta
            // Catch ex As Exception

            // Return False

            // アクセスが拒否された場合
            // Catch ex As UnauthorizedAccessException
            // MsgResult = MessageBox.Show("フォルダ作成時にアクセスが拒否されました。" & vbCrLf & _
            // "コンバーターのインストール先にアクセス権限を与えるか、インストール先を変更して再度セットアップを行って下さい。", _
            // "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            // Return False

            // 無効なフォルダパスが指定された場合
            catch (FileNotFoundException ex)
            {
                CommonModule.MsgResult = MessageBox.Show("無効なファイルパスが指定されました。" + Constants.vbCrLf + "関連ファイルが削除された可能性があります。再度セットアップを行ってから、本プログラムを実行して下さい。", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return false;
            }

            // その他(アクセスが拒否された場合を含む)
            catch (Exception ex)
            {
                CommonModule.MsgResult = MessageBox.Show(filename + "へのアクセスが拒否されました。" + Constants.vbCrLf + filename + "の読み取り専用チェックを外すか、セットアップ先をアクセス権限のあるフォルダに変更して、再度セットアップを行ってから、本プログラムを実行して下さい。", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return false;
                // 20161104 中間ファイル読込エラー時の処理対応 -chg end
            }

            return true;

        }

        /// <summary>
        /// Excel読込時のクローズ処理 '20161004 既存中間→DB書込処理修正 -add
        /// </summary>
        /// <param name="con_read"></param>
        /// <remarks></remarks>
        public void ExcelFile_ReadClose(ref OleDbConnection con_read)
        {
            con_read.Close();
        }

    }

    #endregion

    #region 住所変換関連クラス

    public class AddressChange
    {

        /// <summary>
        /// 住所分割＋チェック処理
        /// </summary>
        /// <param name="chkstr">指定住所</param>
        /// <param name="sqlcnn">接続情報</param>
        /// <returns>分割住所情報</returns>
        /// <remarks>
        /// ・都道府県市区町村丁番地その他情報分割
        /// ・例：chkstr = "宮崎県都城市"
        /// </remarks>
        public static SafeDictionary<string, string> Get_Address(string chkstr, ref SqlConnection sqlcnn)


        {

            var readtbl = new DataTable();                                    // DataTableオブジェクト生成
            string fldname;                                           // 項目名
            string fldvalu;                                           // 項目値
            var rtnhash = new SafeDictionary<string, string>();                                    // 返却用抽出データ格納変数
            var tmphash = new SafeDictionary<string, string>();                                    // 作業用抽出データ格納変数

            // 都道府県市情報抽出クエリ
            string tmp_sql = "";
            tmp_sql = tmp_sql + " SELECT (SELECT ken_name FROM m_ken WHERE ken_no =  m_si.ken_no) AS ken_name,* ";
            tmp_sql = tmp_sql + " FROM m_si ";
            tmp_sql = tmp_sql + " WHERE '";
            tmp_sql = tmp_sql + chkstr;
            tmp_sql = tmp_sql + "' LIKE ((SELECT ken_name FROM m_ken WHERE ken_no =  m_si.ken_no) + si_name + '%') ";

            // 件数取得
            int reccnt = 0;
            bool argnormalflg = true;
            reccnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnn, ref readtbl, normalflg: ref argnormalflg);

            // 20160622 住所分割処理対応 -chg sta
            // 既存処理は都道府県がヒットしない場合全てその他へ移行していたが分割処理自体は行うように修正
            // If reccnt <> 0 Then        '抽出件数あり
            // '抽出データありの場合(都道府県市情報あり)
            // For cntii = 0 To readtbl.Columns.Count - 1
            // '項目名取得
            // fldname = readtbl.Columns(cntii).ColumnName
            // '登録値取得
            // fldvalu = N3Lib.Utys.Typ.ToStr(readtbl.Rows(0).Item(cntii))
            // 'データ格納
            // rtnhash.Add(fldname, fldvalu)
            // Next
            // tmphash = Get_CyoBanti(chkstr.Replace(rtnhash("ken_name"), "").Replace(rtnhash("si_name"), ""))
            // rtnhash.Add("mati", tmphash("mati"))
            // rtnhash.Add("cyome", tmphash("cyome"))
            // rtnhash.Add("banti", tmphash("banti"))
            // rtnhash.Add("etc", tmphash("etc"))
            // rtnhash.Add("chomeptn", tmphash("chomeptn"))
            // Else
            // '抽出データなしの場合
            // rtnhash.Add("etc", chkstr)
            // 'ログ出力処理

            // End If
            string tmp_addresstotal = "";
            string tmp_mati = "";
            string tmp_address = "";

            if (reccnt != 0)        // 抽出件数あり
            {
                // 抽出データありの場合(都道府県市情報あり)
                for (int cntii = 0, loopTo = readtbl.Columns.Count - 1; cntii <= loopTo; cntii++)
                {
                    // 項目名取得
                    fldname = readtbl.Columns[cntii].ColumnName;
                    // 登録値取得
                    fldvalu = Typ.ToStr(readtbl.Rows[0][cntii]);
                    // データ格納
                    rtnhash.Add(fldname, fldvalu);
                }

                tmp_addresstotal = chkstr.Replace(Conversions.ToString(rtnhash["ken_name"]), "").Replace(Conversions.ToString(rtnhash["si_name"]), "");
            }
            else
            {
                tmp_addresstotal = chkstr;
            }

            tmp_mati = Get_Cyotiiki(tmp_addresstotal);
            if (!string.IsNullOrEmpty(tmp_mati))
            {
                tmp_address = tmp_addresstotal.Replace(tmp_mati, "");
            }
            else
            {
                tmp_address = tmp_addresstotal;
            }
            tmphash = Get_CyoBanti(tmp_address);
            rtnhash.Add("mati", tmp_mati);
            rtnhash.Add("cyome", tmphash["cyome"]);
            rtnhash.Add("banti", tmphash["banti"]);
            rtnhash.Add("etc", tmphash["etc"]);
            rtnhash.Add("chomeptn", tmphash["chomeptn"]);
            // 20160622 住所分割処理対応 -chg end

            return rtnhash;

        }

        // 20160622 住所分割処理対応 -del sta
        // ''' <summary>
        // ''' 町村丁番地その他情報取得
        // ''' </summary>
        // ''' <param name="chkstr">都道府県市区除去済み住所データ</param>
        // ''' <returns>町丁番地その他情報</returns>
        // ''' <remarks>
        // ''' ・全角文字を半角文字へ変換して処理<br/>
        // ''' 
        // ''' ※例(テスト用) 
        // '''     chkstr = "宮崎県都城市妻ヶ丘1-2"
        // '''     chkstr = "鹿児島県鹿児島市吉野町7815"
        // '''     chkstr = "熊本県熊本区１２丁目３番地４５６　1aiueo34"
        // '''     chkstr = "熊本県熊本区１-4-9　1aiueo34"
        // '''     chkstr = "北海道古宇郡神恵内村aaa 12-33a"
        // '''     chkstr = 2345
        // '''     chkstr = "Flat D, 73/F, Block 2, The Arch,"
        // '''     chkstr = "テスト3丁目-2-3"
        // ''' </remarks>
        // Public Shared Function Get_CyoBanti(ByVal chkstr As String) As SafeDictionary<string, string>

        // Dim idxban As Integer                                                   '検索文字位置
        // Dim tmpAddrMati As String = ""                                          '丁番地以前格納
        // Dim tmpAddrKari As String = ""                                          '丁番地以降格納
        // Dim tmpAddrChome As String = ""                                         '丁目
        // Dim tmpAddrBan As String = ""                                           '番地
        // Dim tmpAddrEtc As String = ""                                           'その他
        // Dim tmpChomePtn As Integer = -1                                         '丁目パターン(固定)
        // Dim chkflg As Boolean = False                                           '番地チェック：True.該当あり,False.該当なし

        // '全角文字を半角へ変換
        // chkstr = StrConv(chkstr, VbStrConv.Narrow)

        // '番地チェック
        // For chkLen As Integer = 1 To Len(chkstr)
        // '1文字ずつチェック
        // Dim chklenstr As String = Mid(chkstr, chkLen, 1)

        // If IsNumeric(chklenstr) Then
        // '数値の場合(丁番地以降)
        // If chkLen > 1 Then
        // tmpAddrMati = N3Lib.Utys.Substr(chkstr, 0, chkLen - 1)
        // tmpAddrKari = chkstr.Replace(tmpAddrMati, "")
        // Else
        // tmpAddrKari = chkstr
        // End If

        // '町丁番地(配列(0))
        // Dim tmpAddrTemp1 As Array = tmpAddrKari.Replace("　", " ").Split(" ")
        // 'その他(配列(1)～)
        // If UBound(tmpAddrTemp1) > 0 Then
        // For idxban = 1 To UBound(tmpAddrTemp1)
        // tmpAddrEtc = Trim(tmpAddrEtc & " " & tmpAddrTemp1(idxban))
        // Next
        // End If

        // '"丁目"位置 
        // idxban = IIf(tmpAddrTemp1(0).IndexOf("丁目") = 0, tmpAddrTemp1(0).IndexOf("丁"), tmpAddrTemp1(0).IndexOf("丁目"))
        // If idxban <> -1 Then
        // '"丁目"連結文字列チェック
        // tmpAddrChome = N3Lib.Utys.Substr(tmpAddrTemp1(0), 0, idxban)
        // tmpAddrBan = tmpAddrTemp1(0).ToString.Replace(tmpAddrChome.ToString & "丁目", "")
        // tmpAddrBan = tmpAddrBan.ToString.Replace(tmpAddrChome.ToString & "丁", "")
        // tmpAddrBan = tmpAddrBan.TrimStart("-")
        // tmpChomePtn = 1
        // Else
        // '"-"連結文字列チェック
        // Dim tmpAddrTemp2 As Array = tmpAddrTemp1(0).ToString.Split("-")
        // For cnt As Integer = 0 To UBound(tmpAddrTemp2)
        // Select Case cnt
        // Case 0
        // tmpAddrChome = tmpAddrTemp2(cnt)
        // Case Else
        // tmpAddrBan = tmpAddrBan & "-" & tmpAddrTemp2(cnt)
        // tmpAddrBan = tmpAddrBan.TrimStart("-")
        // End Select
        // Next
        // End If

        // '処理を抜ける
        // chkflg = True
        // Exit For
        // End If
        // Next

        // '番地チェック該当なしの場合
        // If chkflg = False Then
        // tmpAddrMati = chkstr
        // End If

        // '[丁目など]設定
        // If tmpAddrChome.Trim <> "" Then
        // tmpChomePtn = 1
        // Else
        // tmpChomePtn = -1
        // End If

        // '返却値
        // Dim rtnhash As New SafeDictionary<string, string>
        // rtnhash.Add("mati", tmpAddrMati)
        // rtnhash.Add("cyome", tmpAddrChome)
        // rtnhash.Add("banti", tmpAddrBan)
        // rtnhash.Add("etc", tmpAddrEtc)
        // rtnhash.Add("chomeptn", tmpChomePtn)
        // Return rtnhash

        // End Function
        // 20160622 住所分割処理対応 -del end

        /// <summary>
        /// 文字列が連結された状態の住所から町地域のみ抽出する '20160622 住所分割処理対応
        /// </summary>
        /// <param name="chkstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Get_Cyotiiki(string chkstr)
        {

            string rtn_str = "";
            string tmpAddrMati = "";

            if (string.IsNullOrEmpty(chkstr.Trim()))
            {
                return rtn_str;
            }

            for (int cntii = 1, loopTo = Strings.Len(chkstr); cntii <= loopTo; cntii++)
            {

                // 1文字ずつチェック
                string chklenstr = Strings.Mid(chkstr, cntii, 1);

                if (Typ.IsNumeric(chklenstr))
                {

                    if (cntii > 1)
                    {
                        tmpAddrMati = Typ.Substr(chkstr, 0, cntii - 1);
                        rtn_str = tmpAddrMati;
                        return rtn_str;
                    }
                    else
                    {
                        // 1文字目が数値の場合は町地域が無く丁目から始まっていると見なす(町地域が = "" とする)
                        return rtn_str;
                    }

                }

            }

            return rtn_str;

        }

        /// <summary>
        /// 丁番地分割処理(革命V7から引用して一部編集) '20160622 住所分割処理対応
        /// </summary>
        /// <param name="strtotal"></param>
        /// <param name="strcyo"></param>
        /// <param name="strbanti"></param>
        /// <param name="stretc"></param>
        /// <remarks></remarks>
        public static SafeDictionary<string, string> Get_CyoBanti(string strtotal)
        {

            var rtn_hash = new SafeDictionary<string, string>();
            string strTmpOrg = "";
            string strTmp1 = "";
            string strTmp2 = "";
            string strTmp3 = "";
            string strTmp4 = "";

            int i1 = 0;
            int i2 = 0;
            int i3 = 0;

            int a1 = 0;
            int a2 = 0;
            int a3 = 0;
            int a4 = 0;

            int e1 = 0;
            int e2 = 0;

            bool rtn = true;

            if (string.IsNullOrEmpty(strtotal))
            {
                return rtn_hash;
            }

            strTmpOrg = Strings.StrConv(strtotal, Constants.vbNarrow);

            strTmpOrg = Strings.Replace(strTmpOrg, "町目", "丁目");

            strTmpOrg = Strings.Replace(strTmpOrg, "一", "1");
            strTmpOrg = Strings.Replace(strTmpOrg, "二", "2");
            strTmpOrg = Strings.Replace(strTmpOrg, "三", "3");
            strTmpOrg = Strings.Replace(strTmpOrg, "四", "4");
            strTmpOrg = Strings.Replace(strTmpOrg, "五", "5");
            strTmpOrg = Strings.Replace(strTmpOrg, "六", "6");
            strTmpOrg = Strings.Replace(strTmpOrg, "七", "7");
            strTmpOrg = Strings.Replace(strTmpOrg, "八", "8");
            strTmpOrg = Strings.Replace(strTmpOrg, "九", "9");

            strTmpOrg = Strings.Replace(strTmpOrg, "ｰ", "-");
            strTmpOrg = Strings.Replace(strTmpOrg, "ー", "-");
            strTmpOrg = Strings.Replace(strTmpOrg, "-", "-");
            strTmpOrg = Strings.Replace(strTmpOrg, "－", "-");
            strTmpOrg = Strings.Replace(strTmpOrg, "―", "-");
            strTmpOrg = Strings.Replace(strTmpOrg, "‐", "-");

            // 不正な文字が入っていないかチェック
            for (int cntii = 1, loopTo = Strings.Len(strTmpOrg); cntii <= loopTo; cntii++)
            {
                strTmp1 = Strings.Mid(strTmpOrg, cntii, 1);

                bool exitFor = false;
                switch (strTmp1 ?? "")
                {
                    case var @case when @case == "":
                    case " ":
                    case "-":
                    case "丁":
                    case "目":
                    case "番":
                    case "地":
                    case "割":
                    case "号":
                        {
                            break;
                        }
                    case "0":
                    case "1":
                    case "2":
                    case "3":
                    case "4":
                    case "5":
                    case "6":
                    case "7":
                    case "8":
                    case "9":
                        {
                            break;
                        }

                    default:
                        {
                            strTmp4 = Strings.Mid(strTmpOrg, cntii); // その他住所

                            if (cntii > 1)
                            {
                                strTmpOrg = strTmpOrg.Substring(0, cntii - 1);
                                exitFor = true;
                                break;
                            }

                            break;
                        }
                }

                if (exitFor)
                {
                    break;
                }
            }

            strTmp1 = "";
            strTmp2 = "";
            strTmp3 = "";

            // 「丁目」「丁」「地割」の判別
            string tmp_cyostr = "";
            string strcyoptn = "";

            if ((strTmpOrg ?? "") != (strTmpOrg.Replace("丁目", "") ?? ""))        // 「丁目」が含まれるか判定
            {
                tmp_cyostr = "丁目";
                strcyoptn = "1";
            }
            else if ((strTmpOrg ?? "") != (strTmpOrg.Replace("丁", "") ?? ""))      // 「丁」が含まれるか判定
            {
                tmp_cyostr = "丁";
                strcyoptn = "2";
            }
            else if ((strTmpOrg ?? "") != (strTmpOrg.Replace("地割", "") ?? ""))    // 「地割」が含まれるか判定
            {
                tmp_cyostr = "地割";
                strcyoptn = "3";
            }
            else if ((strTmpOrg ?? "") != (strTmpOrg.Replace("-", "") ?? ""))       // 「-」が含まれるか判定
            {
                tmp_cyostr = "-";
                strcyoptn = "1";
            }
            else
            {
                tmp_cyostr = "丁目";
                strcyoptn = "-1";
            }

            a1 = Strings.InStr(1, strTmpOrg, tmp_cyostr);
            a2 = Strings.InStr(1, strTmpOrg, "-");
            a3 = Strings.InStr(1, strTmpOrg, "番地");
            a4 = Strings.InStr(1, strTmpOrg, "番");

            // 「丁目」「丁」「地割」が存在する
            // "-" が "番地" or "番" の前に存在する

            if (a1 != 0 | a2 != 0 & a3 != 0 & a4 != 0 & a2 < a3 & a2 < a4 | a2 != 0 & a3 == 0 & a4 == 0)
            {
                if (a1 != 0)
                {
                    strTmp1 = strTmpOrg.Substring(0, a1 - 1);
                    strTmpOrg = Strings.Mid(strTmpOrg, a1 + Strings.Len(tmp_cyostr));
                }
                else if (a2 != 0)
                {
                    strTmp1 = strTmpOrg.Substring(0, a2 - 1);
                    strTmpOrg = Strings.Mid(strTmpOrg, a2 + Strings.Len("-"));
                }

                i1 = Strings.InStr(1, strTmpOrg, "番地");
                i3 = Strings.InStr(1, strTmpOrg, "番");
                i2 = Strings.InStr(1, strTmpOrg, "-");

                if (i1 != 0 | i2 != 0 | i3 != 0)
                {
                    // 20160829 住所分割処理修正 -chg sta
                    // If i1 <> 0 Then
                    // strTmp2 = strTmpOrg.Substring(0, i1 - 1)
                    // strTmpOrg = Mid(strTmpOrg, i1 + Len("番地"))
                    // ElseIf i2 <> 0 Then
                    // strTmp2 = strTmpOrg.Substring(0, i2 - 1)
                    // strTmpOrg = Mid(strTmpOrg, i2 + Len("-"))
                    // ElseIf i3 <> 0 Then
                    // strTmp2 = strTmpOrg.Substring(0, i3 - 1)
                    // strTmpOrg = Mid(strTmpOrg, i3 + Len("番"))
                    // End If
                    // strTmp3 = strTmpOrg
                    strTmp2 = strTmpOrg;
                    // 20160829 住所分割処理修正 -chg end
                    rtn = true;
                }
                else
                {
                    if (a1 == 0)
                    {
                        strTmp2 = strTmp1;
                        strTmp3 = strTmpOrg;
                        strTmp1 = "";
                    }
                    else
                    {
                        strTmp2 = strTmpOrg;
                    }
                    rtn = true;
                }
            }

            else
            {
                i1 = Strings.InStr(1, strTmpOrg, "番地");
                i3 = Strings.InStr(1, strTmpOrg, "番");
                i2 = Strings.InStr(1, strTmpOrg, "-");

                if (i1 != 0 | i2 != 0 | i3 != 0)
                {

                    if (i1 != 0)
                    {
                        strTmp2 = strTmpOrg.Substring(0, i1 - 1);
                        strTmp3 = Strings.Mid(strTmpOrg, i1 + Strings.Len("番地"));
                    }
                    else if (i2 != 0)
                    {
                        strTmp2 = strTmpOrg.Substring(0, i2 - 1);
                        strTmp3 = Strings.Mid(strTmpOrg, i2 + Strings.Len("-"));
                    }
                    else if (i3 != 0)
                    {
                        strTmp2 = strTmpOrg.Substring(0, i3 - 1);
                        strTmp3 = Strings.Mid(strTmpOrg, i3 + Strings.Len("番"));
                    }

                    rtn = true;
                }
                else
                {
                    strTmp4 = strTmpOrg;
                    rtn = true;
                }
            }

            // 例外
            // If strTmp3 = "" And Val(strTmp1) > 50 Then
            // strTmp3 = strTmp2
            // strTmp2 = strTmp1
            // strTmp1 = ""
            // End If

            strTmp1 = Strings.Trim(Strings.Replace(strTmp1, tmp_cyostr, ""));
            strTmp1 = Strings.Trim(Strings.Replace(strTmp1, "番", ""));
            strTmp1 = Strings.Trim(Strings.Replace(strTmp1, "号", ""));
            strTmp1 = Strings.Trim(Strings.Replace(strTmp1, "-", ""));
            // 20160829 住所分割処理修正 -del sta
            // strTmp2 = Trim(Replace(strTmp2, "番", ""))
            // strTmp2 = Trim(Replace(strTmp2, tmp_cyostr, ""))
            // strTmp2 = Trim(Replace(strTmp2, "号", ""))
            // strTmp2 = Trim(Replace(strTmp2, "-", ""))
            // 20160829 住所分割処理修正 -del end
            string strcyo = strTmp1;
            string strbanti = strTmp2;
            // 20160829 住所分割処理修正 -chg sta
            // Dim stretc As String = (strTmp3 & " " & strTmp4).Trim
            string stretc = strTmp4.Trim();
            // 20160829 住所分割処理修正 -chg end
            rtn_hash.Add("cyome", strcyo);
            rtn_hash.Add("banti", strbanti);
            rtn_hash.Add("etc", stretc);
            rtn_hash.Add("chomeptn", strcyoptn);

            return rtn_hash;

        }

    }

    public class AddressChange_old
    {

        /// <summary>
        /// 住所分割＋チェック処理
        /// </summary>
        /// <param name="chkstr">指定住所</param>
        /// <param name="sqlcnn">接続情報</param>
        /// <returns>分割住所情報</returns>
        /// <remarks>
        /// ・都道府県市区町村丁番地その他情報分割
        /// ・例：chkstr = "宮崎県都城市"
        /// </remarks>
        public static SafeDictionary<string, string> Get_Address(string chkstr, ref SqlConnection sqlcnn)


        {

            var readtbl = new DataTable();                                    // DataTableオブジェクト生成
            string fldname;                                           // 項目名
            string fldvalu;                                           // 項目値
            var rtnhash = new SafeDictionary<string, string>();                                    // 返却用抽出データ格納変数
            var tmphash = new SafeDictionary<string, object>();                                    // 作業用抽出データ格納変数

            // 都道府県市情報抽出クエリ
            string tmp_sql = "";
            tmp_sql = tmp_sql + " SELECT (SELECT ken_name FROM m_ken WHERE ken_no =  m_si.ken_no) AS ken_name,* ";
            tmp_sql = tmp_sql + " FROM m_si ";
            tmp_sql = tmp_sql + " WHERE '";
            tmp_sql = tmp_sql + chkstr;
            tmp_sql = tmp_sql + "' LIKE ((SELECT ken_name FROM m_ken WHERE ken_no =  m_si.ken_no) + si_name + '%') ";

            // 件数取得
            int reccnt = 0;
            bool argnormalflg = true;
            reccnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnn, ref readtbl, normalflg: ref argnormalflg);

            if (reccnt != 0)        // 抽出件数あり
            {
                // 抽出データありの場合(都道府県市情報あり)
                for (int cntii = 0, loopTo = readtbl.Columns.Count - 1; cntii <= loopTo; cntii++)
                {
                    // 項目名取得
                    fldname = readtbl.Columns[cntii].ColumnName;
                    // 登録値取得
                    fldvalu = Typ.ToStr(readtbl.Rows[0][cntii]);
                    // データ格納
                    rtnhash.Add(fldname, fldvalu);
                }
                tmphash = Get_CyoBanti(chkstr.Replace(Conversions.ToString(rtnhash["ken_name"]), "").Replace(Conversions.ToString(rtnhash["si_name"]), ""));
                rtnhash.Add("mati", (string)tmphash["mati"]);
                rtnhash.Add("cyome", (string)tmphash["cyome"]);
                rtnhash.Add("banti", (string)tmphash["banti"]);
                rtnhash.Add("etc", (string)tmphash["etc"]);
                rtnhash.Add("chomeptn", tmphash["chomeptn"].ToString());
            }
            else
            {
                // 抽出データなしの場合
                rtnhash.Add("etc", chkstr);
                // ログ出力処理

            }

            return rtnhash;

        }

        /// <summary>
        /// 町村丁番地その他情報取得
        /// </summary>
        /// <param name="chkstr">都道府県市区除去済み住所データ</param>
        /// <returns>町丁番地その他情報</returns>
        /// <remarks>
        /// ・全角文字を半角文字へ変換して処理<br/>
        /// 
        /// ※例(テスト用) 
        ///     chkstr = "宮崎県都城市妻ヶ丘1-2"
        ///     chkstr = "鹿児島県鹿児島市吉野町7815"
        ///     chkstr = "熊本県熊本区１２丁目３番地４５６　1aiueo34"
        ///     chkstr = "熊本県熊本区１-4-9　1aiueo34"
        ///     chkstr = "北海道古宇郡神恵内村aaa 12-33a"
        ///     chkstr = 2345
        ///     chkstr = "Flat D, 73/F, Block 2, The Arch,"
        ///     chkstr = "テスト3丁目-2-3"
        /// </remarks>
        public static SafeDictionary<string, object> Get_CyoBanti(string chkstr)
        {

            int idxban;                                                   // 検索文字位置
            string tmpAddrMati = "";                                          // 丁番地以前格納
            string tmpAddrKari = "";                                          // 丁番地以降格納
            string tmpAddrChome = "";                                         // 丁目
            string tmpAddrBan = "";                                           // 番地
            string tmpAddrEtc = "";                                           // その他
            int tmpChomePtn = -1;                                         // 丁目パターン(固定)
            bool chkflg = false;                                           // 番地チェック：True.該当あり,False.該当なし

            // 全角文字を半角へ変換
            chkstr = Strings.StrConv(chkstr, VbStrConv.Narrow);

            // 番地チェック
            for (int chkLen = 1, loopTo = Strings.Len(chkstr); chkLen <= loopTo; chkLen++)
            {
                // 1文字ずつチェック
                string chklenstr = Strings.Mid(chkstr, chkLen, 1);

                if (Typ.IsNumeric(chklenstr))
                {
                    // 数値の場合(丁番地以降)
                    if (chkLen > 1)
                    {
                        tmpAddrMati = Typ.Substr(chkstr, 0, chkLen - 1);
                        tmpAddrKari = chkstr.Replace(tmpAddrMati, "");
                    }
                    else
                    {
                        tmpAddrKari = chkstr;
                    }

                    // 町丁番地(配列(0))
                    var tmpAddrTemp1 = tmpAddrKari.Replace("　", " ").Split(' ');
                    // その他(配列(1)～)
                    if (Information.UBound(tmpAddrTemp1) > 0)
                    {
                        var loopTo1 = Information.UBound(tmpAddrTemp1);
                        for (idxban = 1; idxban <= loopTo1; idxban++)
                            tmpAddrEtc = Strings.Trim(Conversions.ToString(Operators.ConcatenateObject(tmpAddrEtc + " ", tmpAddrTemp1[idxban])));
                    }

                    // "丁目"位置 
                    idxban = Conversions.ToInteger(Interaction.IIf(Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(tmpAddrTemp1[0].IndexOf("丁目"), 0, false)), tmpAddrTemp1[0].IndexOf("丁"), tmpAddrTemp1[0].IndexOf("丁目")));
                    if (idxban != -1)
                    {
                        // "丁目"連結文字列チェック
                        tmpAddrChome = Typ.Substr(tmpAddrTemp1[0], 0, idxban);
                        tmpAddrBan = tmpAddrTemp1[0].ToString().Replace(tmpAddrChome.ToString() + "丁目", "");
                        tmpAddrBan = tmpAddrBan.ToString().Replace(tmpAddrChome.ToString() + "丁", "");
                        tmpAddrBan = tmpAddrBan.TrimStart('-');
                        tmpChomePtn = 1;
                    }
                    else
                    {
                        // "-"連結文字列チェック
                        var tmpAddrTemp2 = tmpAddrTemp1[0].ToString().Split('-');
                        for (int cnt = 0, loopTo2 = Information.UBound(tmpAddrTemp2); cnt <= loopTo2; cnt++)
                        {
                            switch (cnt)
                            {
                                case 0:
                                    {
                                        tmpAddrChome = Conversions.ToString(tmpAddrTemp2[cnt]);
                                        break;
                                    }

                                default:
                                    {
                                        tmpAddrBan = Conversions.ToString(Operators.ConcatenateObject(tmpAddrBan + "-", tmpAddrTemp2[cnt]));
                                        tmpAddrBan = tmpAddrBan.TrimStart('-');
                                        break;
                                    }
                            }
                        }
                    }

                    // 処理を抜ける
                    chkflg = true;
                    break;
                }
            }

            // 番地チェック該当なしの場合
            if (chkflg == false)
            {
                tmpAddrMati = chkstr;
            }

            // [丁目など]設定
            if (!string.IsNullOrEmpty(tmpAddrChome.Trim()))
            {
                tmpChomePtn = 1;
            }
            else
            {
                tmpChomePtn = -1;
            }

            // 返却値
            var rtnhash = new SafeDictionary<string, object>();
            rtnhash.Add("mati", tmpAddrMati);
            rtnhash.Add("cyome", tmpAddrChome);
            rtnhash.Add("banti", tmpAddrBan);
            rtnhash.Add("etc", tmpAddrEtc);
            rtnhash.Add("chomeptn", tmpChomePtn);
            return rtnhash;

        }

    }


    #endregion

    #region 革命10DBフィールドタイプ格納用ハッシュテーブル作成クラス

    public class GetFieldInfo
    {

        /// <summary>
        /// 革命10DBフィールド情報(サイズ/型)を取得してハッシュテーブルへ格納 接続先不正によるエラー時の処理追加 (Sub→Functionに変更し、正常終了の判別を行う)
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="strsql"></param>
        /// <param name="hash_type"></param>
        /// <param name="hash_size"></param>
        /// <remarks></remarks>
        public static bool Get_HashFieldInfo(SqlConnection sqlcnnv10, string strsql)
        {

            var readtbl = new DataTable();
            int reccnt;
            string fldname;
            string fldvalue;
            string tmp_tblname = "";
            string tmp_tblname_cv = "";
            string tmp_fldname = "";
            string tmp_fldname_cv = "";
            string tmp_type = "";
            string tmp_size = "";
            string tmp_size_cv_max = "";
            string tmp_size_cv_min = "";
            string tmp_defaultvalue = "";
            string tmp_keyno = "";
            string tmp_keyflg = "";
            string tmp_cvkeyno = "";
            string tmp_requiredflg = "";
            string tmp_mstreference = "";
            // 2016.03.28 接続先不正によるエラー時の処理追加 -add
            bool rtn = true;

            // --------------------------------
            // 革命10フィールド情報取得
            // --------------------------------
            bool argnormalflg = true;
            reccnt = DBExec.Exec_DataTable(strsql, ref sqlcnnv10, ref readtbl, normalflg: ref argnormalflg);

            // 2016.03.28 接続先不正によるエラー時の処理追加 -add sta
            // DB情報取得件数→接続情報が不正の可能性があるため処理を抜ける
            if (reccnt == 0)
            {
                rtn = false;
                return rtn;
            }
            // 2016.03.28 接続先不正によるエラー時の処理追加 -add end

            // 2016.02.22 中間ファイルヘッダーエラー修正が生じる現象の対応 -add sta
            // --------------------------------
            // 初期化
            // --------------------------------
            CommonModule.Hash_FiledTypeAlpha.Clear();
            CommonModule.Hash_FiledTypeJp.Clear();
            CommonModule.Hash_FiledSizeAlpha.Clear();
            CommonModule.Hash_FiledSizeJp.Clear();
            CommonModule.Hash_Min_Code.Clear();
            CommonModule.Hash_Max_Code.Clear();
            CommonModule.Hash_TblName_AlphaToJp.Clear();
            CommonModule.Hash_TblName_JpToAlpha.Clear();
            CommonModule.Hash_FldName_AlphaToJp.Clear();
            CommonModule.Hash_FldName_JpToAlpha.Clear();
            CommonModule.Hash_DefaultValue.Clear();
            CommonModule.Hash_MidKey_FieldAlpha.Clear();
            CommonModule.Hash_MidKey_FieldJp.Clear();
            CommonModule.Hash_Mst_ReferenceAlpha.Clear();
            CommonModule.Hash_FldName_JpToAlpha_Mid.Clear();  // 20160812 汎用コンバート対応 -add
            // 2016.02.22 中間ファイルヘッダーエラー修正が生じる現象の対応 -add end

            // --------------------------------
            // 読込開始
            // --------------------------------
            for (int cntii = 0, loopTo = readtbl.Rows.Count - 1; cntii <= loopTo; cntii++)
            {

                // 中断処理
                Application.DoEvents();
                if (CommonModule.CancelFlg)
                {
                    return default;
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
                        case "行No":
                            {
                                tmp_keyno = fldvalue.Trim();
                                break;
                            }
                        case "主キー":
                            {
                                tmp_keyflg = fldvalue.Trim();
                                break;
                            }
                        case "TBL名":
                            {
                                tmp_tblname = fldvalue.Trim();
                                break;
                            }
                        case "フィールド名":
                            {
                                tmp_fldname = fldvalue.Trim();
                                break;
                            }
                        case "属性":
                            {
                                tmp_type = fldvalue.Trim();
                                break;
                            }
                        case "サイズ":
                            {
                                tmp_size = fldvalue.Trim();
                                break;
                            }
                        case "CV用種別":
                            {
                                tmp_tblname_cv = fldvalue.Trim();
                                break;
                            }
                        case "CV用項目名":
                            {
                                tmp_fldname_cv = fldvalue.Trim();
                                break;
                            }
                        case "CV用最小値":
                            {
                                tmp_size_cv_min = fldvalue.Trim();
                                break;
                            }
                        case "CV用最大値":
                            {
                                tmp_size_cv_max = fldvalue.Trim();
                                break;
                            }
                        case "デフォルト値":
                            {
                                tmp_defaultvalue = fldvalue.Trim();
                                break;
                            }
                        case "CV用キーNo":
                            {
                                tmp_cvkeyno = fldvalue.Trim();
                                break;
                            }
                        case "CV用必須項目フラグ":
                            {
                                tmp_requiredflg = fldvalue.Trim();
                                break;
                            }
                        case "有無確認区分":
                            {
                                tmp_mstreference = fldvalue.Trim();
                                break;
                            }
                    }

                }

                // --------------------------------
                // 各オブジェクトへ格納
                // --------------------------------

                // 革命のフィールド名とフィールドタイプ取得
                CommonModule.Hash_FiledTypeAlpha.Add(tmp_tblname + "-" + tmp_fldname, tmp_type);

                // コンバート用フィールド名とフィールドタイプ取得
                if (!string.IsNullOrEmpty(tmp_tblname_cv))
                {
                    CommonModule.Hash_FiledTypeJp.Add(tmp_tblname_cv + "-" + tmp_fldname_cv, tmp_type);
                }

                // 革命のフィールド名とフィールドサイズ取得
                CommonModule.Hash_FiledSizeAlpha.Add(tmp_tblname + "-" + tmp_fldname, tmp_size);

                // コンバート用フィールド名とフィールドサイズ取得
                if (!string.IsNullOrEmpty(tmp_tblname_cv))
                {
                    CommonModule.Hash_FiledSizeJp.Add(tmp_tblname_cv + "-" + tmp_fldname_cv, tmp_size);       // kakaka メモ：Alpha,Jp各、サイズ等の情報を付加
                }

                // 20160812 汎用コンバート対応 -add sta
                // コンバート用フィールド名(日本語)と革命のフィールド名(アルファベット)取得
                if (!string.IsNullOrEmpty(tmp_tblname_cv))
                {
                    CommonModule.Hash_FldName_JpToAlpha_Mid.Add(tmp_tblname_cv + "-" + tmp_fldname_cv, tmp_fldname);
                }
                // 20160812 汎用コンバート対応 -add end

                // 革命上の最小値取得
                if (!string.IsNullOrEmpty(tmp_size_cv_min))
                {
                    CommonModule.Hash_Min_Code.Add(tmp_tblname + "-" + tmp_fldname, tmp_size_cv_min);
                }

                // 革命上の最大値取得
                if (!string.IsNullOrEmpty(tmp_size_cv_max))
                {
                    CommonModule.Hash_Max_Code.Add(tmp_tblname + "-" + tmp_fldname, tmp_size_cv_max);
                }

                // 革命のテーブル名(アルファベット)とコンバート用テーブル名(日本語)取得
                if (!string.IsNullOrEmpty(tmp_tblname) & !CommonModule.Hash_TblName_AlphaToJp.ContainsKey(tmp_tblname))
                {
                    CommonModule.Hash_TblName_AlphaToJp.Add(tmp_tblname, tmp_tblname_cv);
                }

                // コンバート用テーブル名(日本語)と革命のテーブル名(アルファベット)取得
                if (!string.IsNullOrEmpty(tmp_tblname_cv) & !CommonModule.Hash_TblName_JpToAlpha.ContainsKey(tmp_tblname_cv))
                {
                    CommonModule.Hash_TblName_JpToAlpha.Add(tmp_tblname_cv, tmp_tblname);
                }

                // 革命のフィールド名(アルファベット)とコンバート用フィールド名(日本語)取得
                if (!string.IsNullOrEmpty(tmp_tblname_cv))
                {
                    CommonModule.Hash_FldName_AlphaToJp.Add(tmp_tblname + "-" + tmp_fldname, tmp_fldname_cv);
                }

                // コンバート用フィールド名(日本語)と革命のフィールド名(アルファベット)取得
                if (!string.IsNullOrEmpty(tmp_tblname_cv))
                {
                    CommonModule.Hash_FldName_JpToAlpha.Add(tmp_tblname_cv + "-" + tmp_fldname_cv, tmp_tblname + "-" + tmp_fldname);
                }

                // 革命のフィールド名とデフォルト値取得
                if (!string.IsNullOrEmpty(tmp_defaultvalue))
                {
                    CommonModule.Hash_DefaultValue.Add(tmp_tblname + "-" + tmp_fldname, tmp_defaultvalue);
                }

                // 移行元のキーNoを革命のフィールド名で取得
                if (!string.IsNullOrEmpty(tmp_cvkeyno))
                {
                    CommonModule.Hash_MidKey_FieldAlpha.Add(tmp_tblname + "-" + tmp_fldname, tmp_cvkeyno);
                }

                // 移行元のキーNoをコンバート用フィールド名で取得
                if (!string.IsNullOrEmpty(tmp_cvkeyno))
                {
                    CommonModule.Hash_MidKey_FieldJp.Add(tmp_tblname_cv + "-" + tmp_fldname_cv, tmp_cvkeyno);
                }

                // キーではないが必須になっている項目を取得(物件名称等)
                if (!string.IsNullOrEmpty(tmp_requiredflg))
                {
                    CommonModule.List_Required_Field.Add(tmp_tblname + "-" + tmp_fldname);
                }

                // コンバート用フィールド名(日本語)を取得
                if (!string.IsNullOrEmpty(tmp_tblname_cv))
                {
                    CommonModule.List_FldNameJp.Add(tmp_tblname_cv + "-" + tmp_fldname_cv);
                }

                // 有無確認区分を取得
                if (!string.IsNullOrEmpty(tmp_mstreference))
                {
                    CommonModule.Hash_Mst_ReferenceAlpha.Add(tmp_tblname + "-" + tmp_fldname, tmp_mstreference);
                }

            }

            // 2016.03.28 接続先不正によるエラー時の処理追加 -add
            return rtn;

        }

    }

    #endregion

    #region 金融機関マスタ格納用ハッシュテーブル作成クラス

    public class GetBankInfo
    {

        /// <summary>
        /// 金融機関/金融機関支店情報を取得してハッシュテーブルへ格納
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="tmp_sql"></param>
        /// <remarks></remarks>
        public static void Get_HashKinyu(SqlConnection sqlcnnv10, string tmp_sql)
        {

            var readtbl = new DataTable();
            int reccnt;
            string fldname;
            string fldvalue;
            string tmp_kinyuno = "";
            string tmp_kinyuname = "";
            string tmp_tenno = "";
            string tmp_tenname = "";

            // 金融機関情報取得
            bool argnormalflg = true;
            reccnt = DBExec.Exec_DataTable(tmp_sql, ref sqlcnnv10, ref readtbl, normalflg: ref argnormalflg);

            // 読込開始
            for (int cntii = 0, loopTo = readtbl.Rows.Count - 1; cntii <= loopTo; cntii++)
            {

                // 中断処理
                Application.DoEvents();
                if (CommonModule.CancelFlg)
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
                        case "kinyu_no":
                            {
                                tmp_kinyuno = fldvalue;
                                break;
                            }
                        case "kinyu_name":
                            {
                                tmp_kinyuname = fldvalue;
                                break;
                            }
                        case "kinyu_tenno":
                            {
                                tmp_tenno = fldvalue;
                                break;
                            }
                        case "kinyu_tenname":
                            {
                                tmp_tenname = fldvalue;
                                break;
                            }
                    }

                }

                // ハッシュテーブルへ格納
                if (!CommonModule.Hash_Kinyu.ContainsKey(tmp_kinyuno))
                {
                    CommonModule.Hash_Kinyu.Add(tmp_kinyuno, tmp_kinyuname);
                }
                CommonModule.Hash_KinyuTen.Add(tmp_kinyuno + "-" + tmp_tenno, tmp_kinyuname + "-" + tmp_tenname);

            }

        }

    }

    #endregion

    #region フィールド/移行値紐付けハッシュテーブル作成クラス

    public class GetHashFldToValue
    {

        /// <summary>
        /// フィールド名と移行値を紐付けてハッシュテーブルへ格納
        /// </summary>
        /// <param name="fldnamegrp"></param>
        /// <param name="model_cvitem"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static SafeDictionary<string, string> Get_Hash_fldvalue(string fldnamegrp, object model_cvitem)
        {

            var rtn_hash = new SafeDictionary<string, string>();
            var modeltype = model_cvitem.GetType();
            string[] fldname = fldnamegrp.Split(',');

            for (int cntii = 0, loopTo = Information.UBound(fldname); cntii <= loopTo; cntii++)
            {

                string tmp_variname = "Vari_" + fldname[cntii].Substring(0, 1).ToUpper() + fldname[cntii].Substring(1);
                var propertyinfo = modeltype.GetProperty(tmp_variname);
                var targetproperty = propertyinfo.GetValue(model_cvitem, (object[])null);
                string targetvalue = Conversions.ToString(targetproperty);

                rtn_hash.Add(fldname[cntii], targetvalue);

            }

            return rtn_hash;

        }

        /// <summary>
        /// フィールド名と移行値を紐付けてハッシュテーブルへ格納
        /// 備考等、複数行挿入する場合
        /// </summary>
        /// <param name="fldnamegrp"></param>
        /// <param name="model_cvitem"></param>
        /// <param name="rowindex"></param>
        /// <param name="colindex"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static SafeDictionary<string, string> Get_Hash_fldvalue(string fldnamegrp, object model_cvitem, int rowindex, int colindex)
        {

            var rtn_hash = new SafeDictionary<string, string>();
            var modeltype = model_cvitem.GetType();
            string[] fldname = fldnamegrp.Split(',');

            for (int cntii = 0, loopTo = Information.UBound(fldname); cntii <= loopTo; cntii++)
            {

                string tmp_variname = "Vari_" + fldname[cntii].Substring(0, 1).ToUpper() + fldname[cntii].Substring(1) + colindex.ToString();
                var propertyinfo = modeltype.GetProperty(tmp_variname);
                string[] targetproperty = (string[])propertyinfo.GetValue(model_cvitem, (object[])null);
                string targetvalue = Conversions.ToString(targetproperty[rowindex]);

                rtn_hash.Add(fldname[cntii], targetvalue);

            }

            return rtn_hash;

        }

        /// <summary>
        /// フィールド名と移行値を紐付けてハッシュテーブルへ格納(配列同士を紐付け)
        /// </summary>
        /// <param name="fldname"></param>
        /// <param name="fldvalue"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static SafeDictionary<string, string> Get_Hash_fldvalue(string tblname, string[,] fldname, string[,] fldvalue)
        {

            var rtn_hash = new SafeDictionary<string, string>();

            for (int cntii = 0, loopTo = Information.UBound((Array)fldname); cntii <= loopTo; cntii++)

                // 2016.02.22 セル取得方法修正 -chg sta
                // rtn_hash.Add(tblname & "-" & fldname(cntii), fldvalue(cntii))
                // 2016.02.22 セル取得方法修正 -chg end
                rtn_hash.Add(tblname + "-" + fldname[1, cntii + 1], fldvalue[1, cntii + 1]);

            return rtn_hash;

        }

        /// <summary>
        /// Xpathと移行値を紐付けてハッシュテーブルへ格納 '20160616 初期設定情報移行処理実装
        /// </summary>
        /// <param name="fldnamegrp"></param>
        /// <param name="model_cvitem"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static SafeDictionary<string, string> Get_Hash_fldvalue(SafeDictionary<string, string> hash_varitopath, object model_cvitem)
        {

            var rtn_hash = new SafeDictionary<string, string>();
            var modeltype = model_cvitem.GetType();
            // Dim fldname() As String = fldnamegrp.Split(",")

            // modeltype.GetProperty()

            // For cntii = 0 To UBound(fldname)

            // Dim tmp_variname As String = "Vari_" & fldname(cntii).Substring(0, 1).ToUpper & fldname(cntii).Substring(1)
            // Dim propertyinfo As Object = modeltype.GetProperty(tmp_variname)
            // Dim targetproperty As Object = propertyinfo.getValue(model_cvitem, Nothing)
            // Dim targetvalue As String = targetproperty

            // rtn_hash.Add(fldname(cntii), targetvalue)

            // Next

            return rtn_hash;

        }

        /// <summary>
        /// 既存用中間ファイルのフィールド名と移行値を紐付けてハッシュテーブルへ格納 '20160812 汎用コンバート対応 -add
        /// </summary>
        /// <param name="fldname"></param>
        /// <param name="fldvalue"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static SafeDictionary<string, string> Get_Hash_fldvalue_MidheaderToValue(string tblname, string[,] fldname, string[,] fldvalue)
        {

            var rtn_hash = new SafeDictionary<string, string>();

            for (int cntii = 0, loopTo = fldname.Length - 1; cntii <= loopTo; cntii++)
            {
                string midfldname = fldname[1, cntii + 1];
                rtn_hash.Add(midfldname, fldvalue[1, cntii + 1]);
            }

            return rtn_hash;

        }

        /// <summary>
        /// 汎用用中間ファイルのフィールド名と移行値を紐付けてハッシュテーブルへ格納(部屋設備用) '20160913_2 部屋設備移行処理の追加 -add
        /// </summary>
        /// <param name="fldname"></param>
        /// <param name="fldvalue"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static SafeDictionary<string, string> Get_Hash_fldvalue_MidheaderToValue_Setubi(string tblname, string[,] fldname, string[,] fldvalue)
        {

            var rtn_hash = new SafeDictionary<string, string>();

            for (int cntii = 1, loopTo = Conversions.ToInteger(Operators.SubtractObject(fldname.Length, 1)); cntii <= loopTo; cntii++)
            {
                string midfldname = Conversions.ToString(fldname[0, cntii]);
                rtn_hash.Add(midfldname, fldvalue[1, cntii]);
            }

            return rtn_hash;

        }

    }

    #endregion

    #region ログ出力設定クラス

    public class LogSetting
    {

        /// <summary>
        /// ログ出力テーブルCREATEクエリ生成
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Get_LogTblCreateQry(bool midchkflg)
        {

            string tmp_sql = "";
            string taisyotbl = "";

            if (midchkflg)
            {
                taisyotbl = CommonModule.LOG_TMP_MIDCHKTABLENAME;
            }
            else
            {
                taisyotbl = CommonModule.LOG_TMP_TABLENAME;
            }

            tmp_sql = tmp_sql + " CREATE TABLE " + taisyotbl;
            tmp_sql = tmp_sql + " 	( ";
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER1 + " NCHAR(50), ";
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER2 + " NCHAR(50), ";
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER3 + " NCHAR(50), ";
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER4 + " NCHAR(50), ";
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER5 + " NCHAR(200), ";               // 2016.03.28 許容文字数を拡張
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER6 + " NCHAR(100), ";                // kakaka 50で足りる？足りればOK
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER7 + " NCHAR(100), ";
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER8 + " NCHAR(200), ";                // kakaka 50で足りる？足りればOK
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER9 + " NCHAR(200), ";                // kakaka 50で足りる？足りればOK
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER10 + " NCHAR(200), ";
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER11 + " NCHAR(200), ";
            tmp_sql = tmp_sql + CommonModule.LOG_HEADER12 + " NCHAR(50), ";
            tmp_sql = tmp_sql + "  ";
            tmp_sql = tmp_sql + " 	); ";

            return tmp_sql;

        }

        /// <summary>
        /// ログ出力テーブルINSERTクエリ生成
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Get_LogTblInsertQry(string logvalue, bool midchkflg)
        {

            string tmp_sql = "";
            string taisyotbl = "";

            if (midchkflg)
            {
                taisyotbl = CommonModule.LOG_TMP_MIDCHKTABLENAME;
            }
            else
            {
                taisyotbl = CommonModule.LOG_TMP_TABLENAME;
            }

            tmp_sql = " INSERT INTO " + taisyotbl + " ( " + CommonModule.LOG_HEADER_TOTAL + ") VALUES (";
            tmp_sql = tmp_sql + logvalue;
            tmp_sql = tmp_sql + ");";

            return tmp_sql;

        }

        /// <summary>
        /// ログ出力テーブルSELECTクエリ作成
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Get_LogTblSelectQry(bool midchkflg)
        {

            string tmp_sql = "";
            string taisyotbl = "";

            if (midchkflg)
            {
                taisyotbl = CommonModule.LOG_TMP_MIDCHKTABLENAME;
            }
            else
            {
                taisyotbl = CommonModule.LOG_TMP_TABLENAME;
            }

            tmp_sql = tmp_sql + " SELECT ";
            tmp_sql = tmp_sql + "     RTRIM(" + CommonModule.LOG_HEADER1 + ") AS " + CommonModule.LOG_HEADER1;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER2 + ") AS " + CommonModule.LOG_HEADER2;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER3 + ") AS " + CommonModule.LOG_HEADER3;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER4 + ") AS " + CommonModule.LOG_HEADER4;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER5 + ") AS " + CommonModule.LOG_HEADER5;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER6 + ") AS " + CommonModule.LOG_HEADER6;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER7 + ") AS " + CommonModule.LOG_HEADER7;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER8 + ") AS " + CommonModule.LOG_HEADER8;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER9 + ") AS " + CommonModule.LOG_HEADER9;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER10 + ") AS " + CommonModule.LOG_HEADER10;
            tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER11 + ") AS " + CommonModule.LOG_HEADER11;
            if (CommonModule.Dev_CVFlg)                                                                       // kakaka メモ：TBL名は通常非表示にする
            {
                tmp_sql = tmp_sql + "    ,RTRIM(" + CommonModule.LOG_HEADER12 + ") AS " + CommonModule.LOG_HEADER12;
            }
            tmp_sql = tmp_sql + " FROM " + taisyotbl;
            tmp_sql = tmp_sql + " ORDER BY " + CommonModule.LOG_HEADER1;

            return tmp_sql;

        }

        /// <summary>
        /// ログに出力する値を設定
        /// </summary>
        /// <param name="typeno"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Set_LogValue(int typeno, string syorikomok = "-", string syoriitem = "-", string syorikekka = "-", string taisyodata = "-", string hubigein = "-", string taisyo = "-", string beforechgvalue = "-", string afterchgvalue = "-", string tblname = "-")








        {

            int itemcnt = 12;
            var tmp_logitem = new string[itemcnt];
            string tmp_str = "";
            string rtn_logvalue = "";
            // 20161108_2 レビュー結果戻り修正時に気付いた箇所の修正 -chg sta
            // 余計な「/」が入っていたため除去
            // 日付と時間の間に空白があるとCSVファイルをExcelで開いた際に日付の表示形式が
            // 不正になるので「_」を入れる
            // '20161025 ログ出力日時をミリ秒まで拡張する -chg sta
            // 'tmp_logitem(0) = String.Format(Now)
            // tmp_logitem(0) = DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss fff")
            // '20161025 ログ出力日時をミリ秒まで拡張する -chg end
            tmp_logitem[0] = DateTime.Now.ToString("yyyy/MM/dd_HH:mm:ss fff");
            // 20161108_2 レビュー結果戻り修正時に気付いた箇所の修正 -chg end
            tmp_logitem[3] = syorikomok;
            tmp_logitem[4] = syoriitem;
            tmp_logitem[5] = syorikekka;
            // 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg sta
            // tmp_logitem(6) = taisyodata
            tmp_logitem[6] = taisyodata.Replace(",", "，").Replace("'", "’");
            // 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg end
            tmp_logitem[7] = hubigein;
            tmp_logitem[8] = taisyo;
            // 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg sta
            // tmp_logitem(9) = beforechgvalue
            // tmp_logitem(10) = afterchgvalue
            tmp_logitem[9] = beforechgvalue.Replace(",", "，").Replace("'", "’");
            tmp_logitem[10] = afterchgvalue.Replace(",", "，").Replace("'", "’");
            // 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg end
            tmp_logitem[11] = tblname;

            switch (typeno)
            {
                case 1:  // 処理開始
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYU_STA;
                        break;
                    }
                case 9:  // 処理終了
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYU_END;
                        break;
                    }

                case 11: // 移行元読込時
                    {
                        tmp_logitem[1] = "";
                        tmp_logitem[2] = "";
                        break;
                    }

                case 21: // 中間ファイル作成時
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYU_MIDMAKE;
                        break;
                    }

                case 31: // 中間ファイル読込時
                    {
                        tmp_logitem[1] = "";
                        tmp_logitem[2] = "";
                        break;
                    }

                case 41: // DB書込時エラー
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYU_DBWRITE;
                        break;
                    }

                case 42: // DB書込時エラー
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_TYUUI;
                        tmp_logitem[2] = CommonModule.LOG_SYU_DBWRITE;
                        break;
                    }

                case 99: // 中断
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYU_STOP;
                        break;
                    }

                case 101:    // 中間ファイルヘッダーチェック
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_KEIKOKU;
                        tmp_logitem[2] = CommonModule.LOG_SYU_MIDHEADERCHK;
                        break;
                    }
                case 102:    // 中間ファイルデータチェック
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_TYUUI;
                        tmp_logitem[2] = CommonModule.LOG_SYU_MIDDATACHK;
                        break;
                    }

                case 901:    // 汎用→既存U用中間ファイルコピー                   '20161009 ログ出力処理追加 -add
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYORIKOMK_MIDCOPY;
                        break;
                    }
                case 905:    // 汎用→既存U用中間ファイルチェック                 '20161009 ログ出力処理追加 -add
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYORIKOMK_MIDCHECK;
                        break;
                    }
                case 902:    // 汎用→既存U用中間ファイルデータ調整               '20161009 ログ出力処理追加 -add
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYORIKOMK_MIDCONDI;
                        break;
                    }
                case 903:    // 汎用→既存U用中間ファイルコピー(個別)             '20161009 ログ出力処理追加 -add
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYORIKOMK_MIDONECOPY;
                        break;
                    }
                case 904:    // 汎用→既存U用中間ファイルコピーエラー             '20161009 ログ出力処理追加 -add
                    {
                        tmp_logitem[1] = CommonModule.LOG_RUI_RIREKI;
                        tmp_logitem[2] = CommonModule.LOG_SYORIKOMK_ERR_MIDONECOPY;
                        break;
                    }

                default:
                    {
                        break;
                    }

            }

            for (int cntii = 0, loopTo = itemcnt - 1; cntii <= loopTo; cntii++)


                tmp_str = tmp_str + ",'" + tmp_logitem[cntii] + "'";

            rtn_logvalue = tmp_str.Remove(0, 1);

            return rtn_logvalue;

        }

        /// <summary>
        /// 移行時に生じたログ内容を各変数へ格納　　　　　　　　　　　　　　　'kakaka 手順・フローなどは"remarks"に記載下さい。(他の箇所についても同様)
        /// 別メソッド(同一クラス内)にて成形                                  'kakaka 注意：現時点ではいいが、各タグ(paramなど)の内容も記載下さい。(他の箇所についても同様)
        /// 挿入用クエリへ加工
        /// ソートリストへ格納
        /// </summary>
        /// <param name="hash_log"></param>
        /// <param name="hash_beforechk"></param>
        /// <param name="hash_afterchk"></param>
        /// <param name="taisyodata"></param>
        /// <param name="tblname"></param>
        /// <param name="logcnt"></param>
        /// <param name="sortlist"></param>
        /// <remarks></remarks>
        public static void Set_Log_Value_KomkErr(SafeDictionary<string, string> hash_log, SafeDictionary<string, string> hash_beforechk, SafeDictionary<string, string> hash_afterchk, string taisyodata, string tblname, ref int logcnt, ref SortedList<int, string> sortlist)

        {

            foreach (var logvalue in hash_log)
            {

                string log_key = Conversions.ToString(logvalue.Key);
                string log_value = Conversions.ToString(logvalue.Value);
                string log_totalstr = "";

                if (!string.IsNullOrEmpty(log_value.Trim())) // 20160530 ログ修正 ログが空でない場合の条件追加
                {

                    // 処理項目を分割
                    string[] tmp_syori = log_key.Split('/');

                    // 日本語テーブル名を取得して「処理項目_大分類」へ格納
                    string tmp_tblfldname = tmp_syori[0];                 // ※テーブル名は共通なので配列のインデックスは任意の値
                    string[] tmp_tblname = tmp_tblfldname.Split('-');
                    string syorikomk = Conversions.ToString(CommonModule.Hash_TblName_AlphaToJp[tmp_tblname[0]]);

                    // 日本語フィールド名を取得して「処理項目_小分類」へ格納
                    string tmp_syoriitem = "";
                    for (int cntii = 0, loopTo = Information.UBound(tmp_syori); cntii <= loopTo; cntii++)
                    {
                        string tmp_str = Conversions.ToString(CommonModule.Hash_FldName_AlphaToJp[tmp_syori[cntii]]);
                        tmp_syoriitem = tmp_syoriitem + "、" + tmp_str;
                    }
                    tmp_syoriitem = tmp_syoriitem.Remove(0, 1);
                    string syoriitem = tmp_syoriitem;

                    // 「-」で連結されたログ内容を分割
                    string[] tmp_log = log_value.Split('-');

                    // ログの中身を各変数へ格納
                    string syorikekka = tmp_log[0];
                    string hubinaiyo = tmp_log[1];
                    string taisyo = tmp_log[2];

                    // 調整前後の値を変数へ格納
                    // 20160829 口座名義カナチェック機能の追加 -chg sta
                    // Dim tmp_before As String = hash_beforechk(tmp_tblfldname(1))
                    // Dim tmp_after As String = hash_afterchk(tmp_tblfldname(1))
                    string tmp_before = Conversions.ToString(Interaction.IIf(hash_beforechk[tmp_tblname[1]] is null, "", hash_beforechk[tmp_tblname[1]]));
                    string tmp_after = Conversions.ToString(Interaction.IIf(hash_afterchk[tmp_tblname[1]] is null, "", hash_afterchk[tmp_tblname[1]]));
                    // 20160829 口座名義カナチェック機能の追加 -chg end
                    string tmp_empty = "-";
                    string beforechgvalue = "";
                    string afterchgvalue = "";
                    // 20160829 口座名義カナチェック機能の追加 -chg sta
                    // If hash_afterchk.Count = 0 Then         '移行されなかった場合(チェック後ハッシュテーブルにデータが存在しない)は値を設定しない
                    // beforechgvalue = tmp_empty
                    // afterchgvalue = tmp_empty
                    // ElseIf tmp_before = tmp_after Then      '調整前後で値が同じ場合は値を設定しない
                    // beforechgvalue = tmp_empty
                    // afterchgvalue = tmp_empty
                    // ElseIf tmp_before <> tmp_after Then     '調整前後で値が異なる場合は値を設定する(出力時を考慮しカンマを全角に変換しておく)
                    // beforechgvalue = tmp_before.Replace(",", "，")
                    // afterchgvalue = tmp_after.Replace(",", "，")
                    // End If
                    if ((tmp_before ?? "") == (tmp_after ?? ""))          // 調整前後で値が同じ場合は値を設定しない
                    {
                        beforechgvalue = tmp_empty;
                        afterchgvalue = tmp_empty;
                    }
                    else if ((tmp_before ?? "") != (tmp_after ?? ""))     // 調整前後で値が異なる場合は値を設定する(出力時を考慮しカンマを全角に変換しておく)
                    {
                        // 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg sta
                        // 最後にまとめて変換処理を行うためここではそのまま設定するように修正
                        // beforechgvalue = tmp_before.Replace(",", "，")
                        // afterchgvalue = tmp_after.Replace(",", "，")
                        beforechgvalue = tmp_before;
                        afterchgvalue = tmp_after;
                        // 20161209 CSVファイル出力時のカンマ、シングルクォーテーションを全角に変換する処理を追加 -chg end
                    }
                    // 20160829 口座名義カナチェック機能の追加 -chg end
                    // 最終的なログ文字列群を取得
                    log_totalstr = Set_LogValue(42, syorikomk, syoriitem, syorikekka, taisyodata, hubinaiyo, taisyo, beforechgvalue, afterchgvalue, tblname);

                    // 挿入用クエリへ加工
                    string log_sql = Get_LogTblInsertQry(log_totalstr, false);

                    // ログカウントを更新してリストへ格納
                    logcnt = logcnt + 1;
                    sortlist.Add(logcnt, log_sql);

                }

            }

        }

    }

    #endregion

    #region DB書込処理クラス

    public class CVDBInsert
    {

        /// <summary>
        /// 革命10DBへのINSERT処理実行
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="tblname"></param>
        /// <param name="fldname"></param>
        /// <param name="hash_cvitem"></param>
        /// <remarks></remarks>
        public static void Cnv_Db(SqlConnection sqlcnnv10, string tblname, string fldname, SafeDictionary<string, string> hash_cvitem, ref bool flg)
        {

            // 初期値設定
            flg = true;

            // パラメータ生成
            string repprm = Get_Parameter(fldname);

            // フィールド名を分割して配列へ格納
            string[] tmp_fld = fldname.Split(',');

            // クエリ作成
            string tmpcomtxt = "INSERT INTO " + tblname + " (" + fldname + ") VALUES (" + repprm + ")";
            var tmpsqlcom = new SqlCommand(tmpcomtxt, sqlcnnv10);

            // DB挿入値取得
            for (int cntjj = 0, loopTo = Information.UBound(tmp_fld); cntjj <= loopTo; cntjj++)
            {
                string cv_value = Conversions.ToString(hash_cvitem[tmp_fld[cntjj]]);
                SqlComAdd(tmpsqlcom, Interaction.IIf(string.IsNullOrEmpty(cv_value), DBNull.Value, cv_value), tblname + "-" + tmp_fld[cntjj], "@" + tmp_fld[cntjj]);
            }

            // 挿入処理
            try
            {
                int rowsAffected = tmpsqlcom.ExecuteNonQuery();
            }
            catch (Exception ex)                                               // kakaka INSERT失敗を重複と判断、呼び出し先にてnormalflgへセットして重複データのカウントを取っている？↓
            {
                flg = false;                                                     // kakaka 重複エラーではない場合も重複エラーとカウントされてる？
            }
            finally
            {
                // 終了処理
                tmpsqlcom.Dispose();
            }

        }

        /// <summary>
        /// 革命10フィールドタイプ設定
        /// </summary>
        /// <param name="tmpsqlcom"></param>
        /// <param name="objval"></param>
        /// <param name="sqlprm"></param>
        /// <remarks></remarks>
        public static void SqlComAdd(SqlCommand tmpsqlcom, object objval, string tblfldname, string sqlprm)
        {

            var spltyp = default(int);
            string tmp_type = Conversions.ToString(CommonModule.Hash_FiledTypeAlpha[tblfldname]);

            switch (tmp_type ?? "")
            {
                case "int":
                    {
                        spltyp = (int)SqlDbType.Int;
                        break;
                    }
                case "smallint":
                    {
                        spltyp = (int)SqlDbType.SmallInt;
                        break;
                    }
                case "varchar":
                    {
                        spltyp = (int)SqlDbType.VarChar;
                        break;
                    }
                case "xml":
                    {
                        spltyp = (int)SqlDbType.Xml;
                        break;
                    }
                case "uniqueidentifier":
                    {
                        spltyp = (int)SqlDbType.UniqueIdentifier;
                        break;
                    }
                case "datetime":
                    {
                        spltyp = (int)SqlDbType.DateTime;
                        break;
                    }
                case "money":
                    {
                        spltyp = (int)SqlDbType.Money;
                        break;
                    }
                case "nvarchar":
                    {
                        spltyp = (int)SqlDbType.NVarChar;
                        break;
                    }
                case "float":
                    {
                        spltyp = (int)SqlDbType.Float;
                        break;
                    }
                case "varbinary":
                    {
                        spltyp = (int)SqlDbType.VarBinary;
                        break;
                    }
                case "nchar":
                    {
                        spltyp = (int)SqlDbType.NChar;
                        break;
                    }
                case "decimal":
                    {
                        spltyp = (int)SqlDbType.Decimal;
                        break;
                    }
                case "date":
                    {
                        spltyp = (int)SqlDbType.Date;
                        break;
                    }
                // Case "sysname"
                // spltyp = SqlDbType.sysname
                case "bit":
                    {
                        spltyp = (int)SqlDbType.Bit;
                        break;
                    }
                case "bigint":
                    {
                        spltyp = (int)SqlDbType.BigInt;
                        break;
                    }
            }

            tmpsqlcom.Parameters.AddWithValue(sqlprm, spltyp);
            tmpsqlcom.Parameters[sqlprm].Value = objval;

        }

        /// <summary>
        /// カンマで連結されたフィールド名群の各フィールド名に「@」を付加
        /// </summary>
        /// <param name="fldnamegrp"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Get_Parameter(string fldnamegrp)
        {

            string rtn_str = "";
            string tmp_str = "";
            string[] tmp_parastr = fldnamegrp.Split(',');

            for (int cntii = 0, loopTo = Information.UBound(tmp_parastr); cntii <= loopTo; cntii++)
                tmp_str = tmp_str + ",@" + tmp_parastr[cntii];

            rtn_str = tmp_str.Remove(0, 1);

            return rtn_str;

        }

    }

    #endregion

    #region 部屋間取関連クラス

    public class MadoriConv
    {

        /// <summary>
        /// 間取り分割処理
        /// </summary>
        /// <param name="chkmadori"></param>
        /// <returns></returns>
        /// <remarks>
        /// ・1データのみ対象<br/>
        ///   例：「2LDK」…「2」と「LDK」に分割
        ///       「3DK+4LDK」…「3」と「DK+4LDK」に分割
        ///       「LDK」…「」と「LDK」に分割
        ///       「345」…「345」と「」に分割
        /// </remarks>
        public static void SplitMadori(string madoristr, ref string madoricnt, ref string madorinaiyo)
        {

            bool numflg = false;                                           // 全数値FLG…True.末尾
            string madori_num = "";
            string madori_nai = "";
            string tmpmadori = "";                                            // 作業用：文字列編集用
            string tmpstr = "";                                               // 作業用：結合用
            string tmpchr = "";                                               // 作業用：1文字チェック

            tmpmadori = madoristr.ToString().Trim();

            if (!string.IsNullOrEmpty(tmpmadori))
            {

                tmpmadori = Strings.StrConv(tmpmadori, VbStrConv.Narrow).ToUpper();          // 半角大文字変換

                int len = tmpmadori.Length;                               // チェック文字数

                for (int cnt = 0, loopTo = len - 1; cnt <= loopTo; cnt++)
                {
                    // 1文字ずつチェック
                    tmpchr = tmpmadori.Substring(cnt, 1);
                    if (Typ.IsNumeric(tmpchr))
                    {
                        tmpstr = tmpstr + tmpchr;
                    }
                    else
                    {
                        break;
                    }
                    if (cnt == len - 1)
                    {
                        // 全て数値の場合
                        numflg = true;
                    }
                }

                madori_num = tmpstr;

                if (!string.IsNullOrEmpty(madori_num))
                {
                    madori_nai = Conversions.ToString(Interaction.IIf(numflg == true, "", tmpmadori.Replace(madori_num, "")));
                }
                else
                {
                    madori_nai = tmpmadori;
                }

            }

            madoricnt = madori_num;
            madorinaiyo = madori_nai;

        }

        /// <summary>
        /// 間取り分割処理(内訳)
        /// </summary>
        /// <param name="madoristr"></param>
        /// <param name="madorino"></param>
        /// <param name="madorikbn"></param>
        /// <param name="madorijo"></param>
        /// <remarks></remarks>
        public static void SplitMadori(string madoristr, ref string[] madorino, ref string[] madorikbn, ref string[] madorijo)
        {

            // 空文字チェック
            if (string.IsNullOrEmpty(madoristr))
            {
                return;
            }

            // 半角スペースで連結された文字列を分割
            string[] tmp_madorigrp = madoristr.Split(' ');

            madorino = new string[Information.UBound(tmp_madorigrp) + 1];
            madorikbn = new string[Information.UBound(tmp_madorigrp) + 1];
            madorijo = new string[Information.UBound(tmp_madorigrp) + 1];

            for (int cntii = 0, loopTo = Information.UBound(tmp_madorigrp); cntii <= loopTo; cntii++)
            {
                // 20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg sta
                // '20160822 部屋間取内訳情報文字列分割処理修正 -chg sta
                // '空文字の場合の処理を入れる
                // ''間取型*畳数の形になっているデータをさらに分割する
                // 'Dim tmp_madoripart() As String = tmp_madorigrp(cntii).Split("*")

                // ''各配列へ格納
                // 'madorino(cntii) = cntii + 1
                // 'madorikbn(cntii) = tmp_madoripart(0)
                // 'madorijo(cntii) = tmp_madoripart(1)

                // '間取No格納(歯抜で登録されたデータは詰めずにそのまま移行する)
                // madorino(cntii) = cntii + 1

                // 'データが存在する場合は分割して格納する
                // If tmp_madorigrp(cntii) <> "" Then

                // '間取型*畳数の形になっているデータをさらに分割する
                // Dim tmp_madoripart() As String = tmp_madorigrp(cntii).Split("*")

                // '各配列へ格納
                // madorikbn(cntii) = tmp_madoripart(0)
                // madorijo(cntii) = tmp_madoripart(1)

                // End If
                // '20160822 部屋間取内訳情報文字列分割処理修正 -chg end


                // 間取No格納(歯抜で登録されたデータは詰めずにそのまま移行する)
                madorino[cntii] = (cntii + 1).ToString();

                // データが存在する場合は分割して格納する
                if (!string.IsNullOrEmpty(tmp_madorigrp[cntii]))
                {

                    // 「*」が含まれているかチェック
                    if ((tmp_madorigrp[cntii].Replace("*", "") ?? "") == (tmp_madorigrp[cntii] ?? ""))
                    {
                        // 「*」が含まれていない場合は「対応不要」を付加して格納
                        madorikbn[cntii] = tmp_madorigrp[cntii] + "(対応不要)";
                        madorijo[cntii] = "";
                    }
                    else
                    {
                        // 間取型*畳数の形になっている場合はさらに分割する
                        string[] tmp_madoripart = tmp_madorigrp[cntii].Split('*');

                        // 各配列へ格納
                        madorikbn[cntii] = tmp_madoripart[0];
                        madorijo[cntii] = tmp_madoripart[1];
                    }

                }
                // 20160825 部屋間取内訳情報文字列分割処理修正 「*」が存在しないデータの対応 -chg end
            }

        }

    }

    #endregion

    #region 都市計画/用途地域関連クラス

    public class TosiYotoConv
    {

        public static void Get_TosiYoto(string tosiyoto, ref int tosiyotokbn, ref string tosiyotono)
        {

            switch (tosiyoto ?? "")
            {
                case "第一種低層住居専用地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "1";
                        break;
                    }
                case "第二種低層住居専用地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "2";
                        break;
                    }
                case "第一種中高層住居専用地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "3";
                        break;
                    }
                case "第二種中高層住居専用地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "4";
                        break;
                    }
                case "第一種住居地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "5";
                        break;
                    }
                case "第二種住居地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "6";
                        break;
                    }
                case "準住居地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "7";
                        break;
                    }
                case "近隣商業地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "8";
                        break;
                    }
                case "商業地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "9";
                        break;
                    }
                case "準工業地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "10";
                        break;
                    }
                case "工業地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "11";
                        break;
                    }
                case "工業専用地域":
                    {
                        tosiyotokbn = 2;
                        tosiyotono = "12";
                        break;
                    }
                case "指定なし":
                    {
                        tosiyotono = "13";
                        tosiyotokbn = 2;
                        break;
                    }
                case "都市計画区域外":
                    {
                        tosiyotokbn = 1;
                        tosiyotono = "7";
                        break;
                    }

                case var @case when @case == "":
                    {
                        tosiyotokbn = 0;
                        tosiyotono = "-1";
                        break;
                    }

                default:
                    {
                        tosiyotokbn = 3;
                        tosiyotono = "99";
                        break;
                    }
            }

        }

    }

    #endregion

    #region 紐付項目取得関連クラス

    public class SetRelItemToObject
    {

        /// <summary>
        /// 入金項目紐付情報取得
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Nkinkomk()
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
                string tmp_newnkinzkseino = "";

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
                        // 20160617 マイナス金額移行処理修正 -add sta
                        case "賃貸革命10入金項目属性No":
                            {
                                tmp_newnkinzkseino = fldvalue.Trim();
                                break;
                            }
                            // 20160617 マイナス金額移行処理修正 -add end
                    }

                }

                // 入金項目をハッシュテーブル格納
                if (!string.IsNullOrEmpty(tmp_newnkinno))
                {
                    // 20160527 指摘事項対応 -chg sta
                    // Hash_Rel_Nkinkomk.Add(tmp_oldnkinname & "-" & tmp_nkinkbn, tmp_newnkinno)
                    if (CommonModule.Hash_Rel_Nkinkomk.ContainsKey(tmp_oldnkinname + "-" + tmp_nkinkbn) == false)
                    {
                        CommonModule.Hash_Rel_Nkinkomk.Add(tmp_oldnkinname + "-" + tmp_nkinkbn, tmp_newnkinno);
                    }
                    // 20160527 指摘事項対応 -chg end
                }

                // 2016.04.06 入金項目読込処理の修正 -add sta
                // 変動費メーター分類をハッシュテーブル格納
                if (tmp_nkinkbn == "5.随時変動" & !string.IsNullOrEmpty(tmp_hendometerno))
                {
                    CommonModule.Hash_Rel_Hendometer.Add(tmp_oldnkinname + "-" + tmp_nkinkbn, tmp_hendometerno);
                }
                // 2016.04.06 入金項目読込処理の修正 -add end

                // 20160617 マイナス金額移行処理修正 -add sta
                if (CommonModule.Hash_Rel_NkinkomkZksei.ContainsKey(tmp_oldnkinname + "-" + tmp_nkinkbn) == false)
                {
                    CommonModule.Hash_Rel_NkinkomkZksei.Add(tmp_oldnkinname + "-" + tmp_nkinkbn, tmp_newnkinzkseino);
                }
                // 20160617 マイナス金額移行処理修正 -add end

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 入金区分紐付情報取得
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Nkinkbn()
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
                    // 20160527 指摘事項対応 -chg sta
                    // Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno)
                    if (CommonModule.Hash_Rel_Nkinkbn.ContainsKey(tmp_oldnkinkbnname) == false)
                    {
                        CommonModule.Hash_Rel_Nkinkbn.Add(tmp_oldnkinkbnname, tmp_newnkinkbnno);
                    }
                    // 20160527 指摘事項対応 -chg end
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 設備紐付情報取得 (V7のデータと10の設備Noを紐付)
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Setubi_V7To10No()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "設備マスタ";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_setubigrpname = "";
                string tmp_setubiname = "";
                // 2016.04.06 部屋設備情報の取得処理修正 -chg sta
                // Dim tmp_setubiguid As String = ""
                string tmp_setubigrpno = "";
                string tmp_setubino = "";
                string tmp_setubikomkno = "";
                // 2016.04.06 部屋設備情報の取得処理修正 -chg end

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "移行元設備グループ名称":
                            {
                                tmp_setubigrpname = fldvalue.Trim();
                                break;
                            }
                        case "移行元設備名称":
                            {
                                tmp_setubiname = fldvalue.Trim();
                                break;
                            }
                        // 2016.04.06 部屋設備情報の取得処理修正 -chg sta
                        // Case "賃貸革命10項目GUID"
                        // tmp_setubiguid = fldvalue.Trim
                        case "賃貸革命10設備グループNo":
                            {
                                tmp_setubigrpno = fldvalue.Trim();
                                break;
                            }
                        case "賃貸革命10設備No":
                            {
                                tmp_setubino = fldvalue.Trim();
                                break;
                            }
                        case "賃貸革命10項目No":
                            {
                                tmp_setubikomkno = fldvalue.Trim();
                                break;
                            }
                            // 2016.04.06 部屋設備情報の取得処理修正 -chg end
                    }

                }

                // ハッシュテーブル格納
                // 2016.04.06 部屋設備情報の取得処理修正 -chg sta
                // If tmp_setubiguid <> "" Then
                // Hash_Rel_Setubi.Add(tmp_setubigrpname & "-" & tmp_setubiname, tmp_setubiguid)
                // End If
                string tmp_V7setubikey = tmp_setubigrpname + "-" + tmp_setubiname;
                // 20160829 エレベーター移行対応 -chg sta
                // Dim tmp_10setubikey As String = tmp_setubigrpno & "-" & tmp_setubino & "-" & tmp_setubikomkno
                string tmp_10setubikey = "";

                if (tmp_setubigrpno != "999" & (string.IsNullOrEmpty(tmp_setubigrpno) | string.IsNullOrEmpty(tmp_setubino) | string.IsNullOrEmpty(tmp_setubikomkno)))
                {
                    tmp_10setubikey = "";
                }
                else if (tmp_setubigrpno == "999" & !string.IsNullOrEmpty(tmp_setubino))
                {
                    tmp_10setubikey = tmp_setubigrpno + "-" + tmp_setubino;
                }
                else
                {
                    tmp_10setubikey = tmp_setubigrpno + "-" + tmp_setubino + "-" + tmp_setubikomkno;
                }
                // 20160829 エレベーター移行対応 -chg end
                // 20160829 エレベーター移行対応 -chg sta
                // If Hash_SetubiMid.Contains(tmp_V7setubikey) = False Then
                // Hash_SetubiMid.Add(tmp_V7setubikey, tmp_10setubikey)
                // End If
                if (CommonModule.Hash_SetubiMid.ContainsKey(tmp_V7setubikey) == false | tmp_setubigrpno == "999")
                {
                    CommonModule.Hash_SetubiMid.Add(tmp_V7setubikey, tmp_10setubikey);
                }
                // 20160829 エレベーター移行対応 -chg end
                // 2016.04.06 部屋設備情報の取得処理修正 -chg end
            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 契約分類紐付情報取得 2016.04.06 契約分類を紐付データを元に移行
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Kyrui()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "契約分類マスタ";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // 初期化
            CommonModule.Hash_Rel_Kyrui.Clear();
            CommonModule.Hash_Rel_Kyrui_Teikisyakuyakbn.Clear();

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_oldkyruiname = "";
                string tmp_newkyruino = "";
                string tmp_newkyruiteisyakukbn = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "移行元契約分類名称":
                            {
                                tmp_oldkyruiname = fldvalue.Trim();
                                break;
                            }
                        case "賃貸革命10契約分類No":
                            {
                                tmp_newkyruino = fldvalue.Trim();
                                break;
                            }
                        case "定期借家として扱う":
                            {
                                tmp_newkyruiteisyakukbn = Conversions.ToString(Interaction.IIf(fldvalue == "True", 1, 0));
                                break;
                            }
                    }

                }

                // ハッシュテーブル格納(契約分類データ)
                if (!string.IsNullOrEmpty(tmp_newkyruino))
                {
                    if (CommonModule.Hash_Rel_Kyrui.ContainsKey(tmp_oldkyruiname) == false)
                    {
                        CommonModule.Hash_Rel_Kyrui.Add(tmp_oldkyruiname, tmp_newkyruino);
                    }
                }

                // ハッシュテーブル格納(契約分類の定期借家区分)
                if (!string.IsNullOrEmpty(tmp_newkyruino))
                {
                    // 20160527 指摘事項対応 -chg sta
                    // Hash_Rel_Kyrui_Teikisyakuyakbn.Add(tmp_oldkyruiname, tmp_newkyruiteisyakukbn)
                    if (CommonModule.Hash_Rel_Kyrui_Teikisyakuyakbn.ContainsKey(tmp_oldkyruiname) == false)
                    {
                        CommonModule.Hash_Rel_Kyrui_Teikisyakuyakbn.Add(tmp_oldkyruiname, tmp_newkyruiteisyakukbn);
                    }
                    // 20160527 指摘事項対応 -chg end
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 取引態様マスタ紐付情報取得
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Toritaiyo()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "取引態様マスタ";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_oldtoritaiyoname = "";
                string tmp_newtoritaiyono = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "移行元取引態様名称":
                            {
                                tmp_oldtoritaiyoname = fldvalue.Trim();
                                break;
                            }
                        case "賃貸革命10取引態様No":
                            {
                                tmp_newtoritaiyono = fldvalue.Trim();
                                break;
                            }
                    }

                }

                // ハッシュテーブル格納
                if (!string.IsNullOrEmpty(tmp_newtoritaiyono))
                {
                    // 20160527 指摘事項対応 -chg sta
                    // Hash_Rel_Toritaiyo.Add(tmp_oldtoritaiyoname, tmp_newtoritaiyono)
                    if (CommonModule.Hash_Rel_Toritaiyo.ContainsKey(tmp_oldtoritaiyoname) == false)
                    {
                        CommonModule.Hash_Rel_Toritaiyo.Add(tmp_oldtoritaiyoname, tmp_newtoritaiyono);
                    }
                    // 20160527 指摘事項対応 -chg end
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 物件分類紐付情報取得
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Bkrui()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "物件分類マスタ";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_oldbkruiname = "";
                string tmp_newbkruino = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "移行元物件分類名称":
                            {
                                tmp_oldbkruiname = fldvalue.Trim();
                                break;
                            }
                        case "賃貸革命10物件分類No":
                            {
                                tmp_newbkruino = fldvalue.Trim();
                                break;
                            }
                    }

                }

                // ハッシュテーブル格納
                if (!string.IsNullOrEmpty(tmp_newbkruino))
                {
                    // 20160527 指摘事項対応 -chg sta
                    // Hash_Rel_Bkrui.Add(tmp_oldbkruiname, tmp_newbkruino)
                    if (CommonModule.Hash_Rel_Bkrui.ContainsKey(tmp_oldbkruiname) == false)
                    {
                        CommonModule.Hash_Rel_Bkrui.Add(tmp_oldbkruiname, tmp_newbkruino);
                    }
                    // 20160527 指摘事項対応 -chg end
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 部屋分類紐付情報取得
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Hyrui()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "部屋分類マスタ";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // 既存データ有無確認 20160829 革命10バージョンアップに伴う修正 -add
            if (CommonModule.Hash_Rel_Hyrui.Count != 0)
            {
                return;
            }

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_oldhyruiname = "";
                string tmp_newhyruino = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "移行元部屋分類名称":
                            {
                                tmp_oldhyruiname = fldvalue.Trim();
                                break;
                            }
                        case "賃貸革命10部屋分類No":
                            {
                                tmp_newhyruino = fldvalue.Trim();
                                break;
                            }
                    }

                }

                // ハッシュテーブル格納
                if (!string.IsNullOrEmpty(tmp_newhyruino))
                {
                    // 20160829 革命10バージョンアップに伴う修正 -chg sta
                    // Hash_Rel_Hyrui.Add(tmp_oldhyruiname, tmp_newhyruino)
                    if (CommonModule.Hash_Rel_Hyrui.ContainsKey(tmp_oldhyruiname) == false)
                    {
                        CommonModule.Hash_Rel_Hyrui.Add(tmp_oldhyruiname, tmp_newhyruino);
                    }
                    // 20160829 革命10バージョンアップに伴う修正 -chg end
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 構造マスタ紐付情報取得
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Kozo()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "構造マスタ";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_oldbkozoname = "";
                string tmp_newkozono = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "移行元構造名称":
                            {
                                tmp_oldbkozoname = fldvalue.Trim();
                                break;
                            }
                        case "賃貸革命10構造No":
                            {
                                tmp_newkozono = fldvalue.Trim();
                                break;
                            }
                    }

                }

                // ハッシュテーブル格納
                if (!string.IsNullOrEmpty(tmp_newkozono))
                {
                    // 20160527 指摘事項対応 -chg sta
                    // Hash_Rel_Kozo.Add(tmp_oldbkozoname, tmp_newkozono)
                    if (CommonModule.Hash_Rel_Kozo.ContainsKey(tmp_oldbkozoname) == false)
                    {
                        CommonModule.Hash_Rel_Kozo.Add(tmp_oldbkozoname, tmp_newkozono);
                    }
                    // 20160527 指摘事項対応 -chg end
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 用途地域・都市計画マスタ紐付情報取得 2016.04.06 都市計画・用途地域を紐付データから取得するように修正
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Yototiki()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            // 20161014 都市計画用途地域ファイルオープンエラー修正 -chg sta
            // Dim relsheetname As String = "用途地域・都市計画マスタ"
            string relsheetname = "都市計画・用途地域マスタ";
            // 20161014 都市計画用途地域ファイルオープンエラー修正 -chg end
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // 初期化
            CommonModule.Hash_Rel_Tosiyotokbn.Clear();
            CommonModule.Hash_Rel_Tosiyotono.Clear();

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_oldyotoname = "";
                string tmp_newyototikikbn = "";
                string tmp_newyototikino = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "移行元用途地域・都市計画名":
                            {
                                tmp_oldyotoname = fldvalue.Trim();
                                break;
                            }
                        case "移行先用途地域・都市計画区分":
                            {
                                tmp_newyototikikbn = fldvalue.Trim();
                                break;
                            }
                        case "移行先用途地域・都市計画No":
                            {
                                tmp_newyototikino = fldvalue.Trim();
                                break;
                            }
                    }

                }

                // ハッシュテーブル格納
                if (!string.IsNullOrEmpty(tmp_newyototikikbn))
                {
                    // 20160527 指摘事項対応 -chg sta
                    // Hash_Rel_Tosiyotokbn.Add(tmp_oldyotoname, tmp_newyototikikbn)
                    if (CommonModule.Hash_Rel_Tosiyotokbn.ContainsKey(tmp_oldyotoname) == false)
                    {
                        CommonModule.Hash_Rel_Tosiyotokbn.Add(tmp_oldyotoname, tmp_newyototikikbn);
                    }
                    // 20160527 指摘事項対応 -chg end
                }

                if (!string.IsNullOrEmpty(tmp_newyototikino))
                {
                    // 20160527 指摘事項対応 -chg sta
                    // Hash_Rel_Tosiyotono.Add(tmp_oldyotoname, tmp_newyototikino)
                    if (CommonModule.Hash_Rel_Tosiyotono.ContainsKey(tmp_oldyotoname) == false)
                    {
                        CommonModule.Hash_Rel_Tosiyotono.Add(tmp_oldyotoname, tmp_newyototikino);
                    }
                    // 20160527 指摘事項対応 -chg end
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 自社口座紐付情報取得
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_Jisyakoza()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "自社口座マスタ";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // 既存データ有無確認
            if (CommonModule.Hash_Rel_JisyaKoza.Count != 0)
            {
                return;
            }

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                // 20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                // Dim tmp_kinyuno As String = ""
                // Dim tmp_kinyutenno As String = ""
                // Dim tmp_kozasyu As String = ""   '20160609 紐付取得用に追加
                // Dim tmp_kozabango As String = ""
                string tmp_kozasyutokumoto = "";
                string tmp_kozasyutokumotono = "";
                // 20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                string tmp_jisyano = "";
                string tmp_jisyakozano = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        // 20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                        // Case "金融機関No"
                        // tmp_kinyuno = fldvalue.Trim
                        // Case "支店No"
                        // tmp_kinyutenno = fldvalue.Trim
                        // Case "口座種別" '20160609 紐付取得用に追加
                        // tmp_kozasyu = fldvalue.Trim
                        // Case "口座番号"
                        // tmp_kozabango = fldvalue.Trim
                        // Case "自社No"
                        // tmp_jisyano = fldvalue.Trim
                        // Case "自社口座No"
                        // tmp_jisyakozano = fldvalue.Trim
                        case "口座取得元":
                            {
                                tmp_kozasyutokumoto = fldvalue.Trim();
                                break;
                            }
                        case "各口座設定No":
                            {
                                tmp_kozasyutokumotono = fldvalue.Trim();
                                break;
                            }
                        case "自社No":
                            {
                                tmp_jisyano = fldvalue.Trim();
                                break;
                            }
                        case "自社口座No":
                            {
                                tmp_jisyakozano = fldvalue.Trim();
                                break;
                            }
                            // 20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                    }

                }

                // 紐付け用に値を成形
                // 20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                // '20160609 紐付取得用に追加 -chg sta
                // 'Dim tmp_kozainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozabango
                // Dim tmp_kozainfo As String = tmp_kinyuno & "-" & tmp_kinyutenno & "-" & tmp_kozasyu & "-" & tmp_kozabango
                // '20160609 紐付取得用に追加 -chg end
                string tmp_kozainfo = tmp_kozasyutokumoto + "-" + tmp_kozasyutokumotono;
                // 20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
                string tmp_jisyainfo = tmp_jisyano + "-" + tmp_jisyakozano;

                // ハッシュテーブル格納
                if (!string.IsNullOrEmpty(tmp_jisyano) && !string.IsNullOrEmpty(tmp_jisyakozano))
                {
                    if (CommonModule.Hash_Rel_JisyaKoza.ContainsKey(tmp_kozainfo) == false)
                    {
                        CommonModule.Hash_Rel_JisyaKoza.Add(tmp_kozainfo, tmp_jisyainfo);
                    }
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// FB関連紐付情報取得 '20160609 紐付設定値取得に伴う修正
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_FBInfo()
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "FBフォーマット割付";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // オブジェクト初期化
            CommonModule.Hash_Rel_FBInfo_KozaFurikaeFmt.Clear();
            CommonModule.Hash_Rel_FBInfo_FuriIraiFmt.Clear();
            CommonModule.Hash_Rel_FBInfo_NsSettingFmt.Clear();
            CommonModule.Hash_Rel_FBInfo_NsSettingJisya.Clear();

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_V7fmtsyubetu = "";
                string tmp_V7kozano = "";
                string tmp_10fmtno = "";
                string tmp_jisyano = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "フォーマット種別":
                            {
                                tmp_V7fmtsyubetu = fldvalue.Trim();
                                break;
                            }
                        case "各口座No":
                            {
                                tmp_V7kozano = fldvalue.Trim();
                                break;
                            }
                        case "フォーマットNo":
                            {
                                tmp_10fmtno = fldvalue.Trim();
                                break;
                            }
                        case "自社No":
                            {
                                tmp_jisyano = fldvalue.Trim();
                                break;
                            }
                    }

                }

                switch (tmp_V7fmtsyubetu ?? "")
                {
                    case "口座振替":
                        {
                            if (CommonModule.Hash_Rel_FBInfo_KozaFurikaeFmt.ContainsKey(tmp_V7kozano) == false)
                            {
                                CommonModule.Hash_Rel_FBInfo_KozaFurikaeFmt.Add(tmp_V7kozano, tmp_10fmtno);
                            }

                            break;
                        }
                    case "総合振込":
                        {
                            if (CommonModule.Hash_Rel_FBInfo_FuriIraiFmt.ContainsKey(tmp_V7kozano) == false)
                            {
                                CommonModule.Hash_Rel_FBInfo_FuriIraiFmt.Add(tmp_V7kozano, tmp_10fmtno);
                            }

                            break;
                        }
                    case "入出金":
                        {
                            if (CommonModule.Hash_Rel_FBInfo_NsSettingFmt.ContainsKey(tmp_V7kozano) == false)
                            {
                                CommonModule.Hash_Rel_FBInfo_NsSettingFmt.Add(tmp_V7kozano, tmp_10fmtno);
                            }
                            if (CommonModule.Hash_Rel_FBInfo_NsSettingJisya.ContainsKey(tmp_V7kozano) == false)
                            {
                                CommonModule.Hash_Rel_FBInfo_NsSettingJisya.Add(tmp_V7kozano, tmp_jisyano);
                            }

                            break;
                        }
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        /// <summary>
        /// 鍵タイトルマスタ情報取得 '20160613 鍵情報の取得処理修正
        /// </summary>
        /// <remarks></remarks>
        public static void Set_RelData_KagiInfo(SqlConnection sqlcnnv10)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            string relsheetname = "鍵タイトルマスタ";
            bool rtn = true;

            var excelfile = new ExcelFileManager();                // Excelファイル操作用

            // オブジェクト初期化
            CommonModule.Hash_KagiTitle_V7to10.Clear();

            // 鍵タイトル情報仮テーブル
            string tmp_kagiinfo_tblname = "tmp_kagiinfo";
            int tmp_rowcnt = 0;

            // 鍵タイトル情報仮テーブル初期化
            string tmp_sql_drop = DBQuery.Qry_DropInfo(tmp_kagiinfo_tblname, true);
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmp_rowcnt);

            // 鍵タイトル情報仮テーブル作成
            string tmp_sql_create = "";
            tmp_sql_create = tmp_sql_create + " CREATE TABLE " + tmp_kagiinfo_tblname;
            tmp_sql_create = tmp_sql_create + " 	( ";
            tmp_sql_create = tmp_sql_create + " 		 oldkagino INT ";
            tmp_sql_create = tmp_sql_create + " 		,newkagikbn INT ";
            tmp_sql_create = tmp_sql_create + " 	) ";
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, ref tmp_rowcnt);

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, relsheetname);

            for (int cntii = 0, loopTo = rowcnt - 1; cntii <= loopTo; cntii++)
            {

                // 作業用変数作成
                string tmp_V7kagino = "";
                string tmp_10kagisyubetu = "";

                // 紐付設定値取得
                for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                {

                    // ヘッダー格納
                    string fldname = Typ.ToStr(wsheet.Cells[startrow - 1, cntjj].Value);

                    // 移行値格納
                    string fldvalue = Typ.ToStr(wsheet.Cells[cntii + startrow, cntjj].Value);

                    switch (fldname ?? "")
                    {
                        case "鍵No":
                            {
                                tmp_V7kagino = fldvalue.Trim();
                                break;
                            }
                        case "共用チェック":
                            {
                                tmp_10kagisyubetu = Conversions.ToString(Interaction.IIf(fldvalue.Trim() == "True", 1, 2));
                                break;
                            }
                    }

                }

                // 仮テーブルへ挿入
                string tmp_insertitem = tmp_V7kagino + "," + tmp_10kagisyubetu;
                string tmp_sql_insert = " INSERT INTO " + tmp_kagiinfo_tblname + " VALUES(" + tmp_insertitem + ");";
                DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, ref tmp_rowcnt);

            }

            // 仮テーブルからV7鍵Noと10鍵No(鍵区分を含む)を紐付けたハッシュテーブルを作成する
            string tmp_sql_select = "";
            tmp_sql_select = tmp_sql_select + " SELECT ";
            tmp_sql_select = tmp_sql_select + " 	 oldkagino ";
            tmp_sql_select = tmp_sql_select + " 	,CONVERT(VARCHAR,newkagikbn) + '-' + CONVERT(VARCHAR,ROW_NUMBER()OVER(PARTITION BY newkagikbn ORDER BY oldkagino)) AS newkagino ";
            tmp_sql_select = tmp_sql_select + " FROM tmp_kagiinfo ";
            DBExec.Exec_DataReader_Col_Hash(tmp_sql_select, ref sqlcnnv10, ref CommonModule.Hash_KagiTitle_V7to10);

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            // 鍵タイトル情報仮テーブル削除
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, ref tmp_rowcnt);

        }

        /// <summary>
        /// 口座種別紐付情報取得
        /// </summary>
        /// <remarks></remarks>
        public static void Get_RelData_KozaSyubetu()
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
                    if (CommonModule.Hash_Rel_Kozasyubetu.ContainsKey(tmp_oldkozasyubetuname) == false)
                    {
                        CommonModule.Hash_Rel_Kozasyubetu.Add(tmp_oldkozasyubetuname, tmp_newkozasyubetuno);
                    }
                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

        }

        // 20160609 紐付設定値取得に伴う修正 -del sta
        // ''' <summary>
        // ''' FBフォーマット紐付情報取得 '20160517 入出金取得情報の新規作成
        // ''' </summary>
        // ''' <remarks></remarks>
        // Public Shared Sub Set_RelData_FBFmt()

        // Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        // Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        // Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        // Dim startrow As Integer                                         '書込開始行
        // Dim columncnt As Integer                                        '列数
        // Dim maxrowcnt As Integer                                        '既存データの行数
        // Dim rowcnt As Integer                                           '書込行数
        // Dim relsheetname As String = "FBフォーマットマスタ"
        // Dim rtn As Boolean = True

        // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

        // 'Excelファイル初期設定
        // rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, relsheetname)

        // For cntii = 0 To rowcnt - 1

        // '作業用変数作成
        // Dim tmp_oldfmtname As String = ""
        // Dim tmp_newfmtno As String = ""

        // '紐付設定値取得
        // For cntjj = 1 To columncnt

        // 'ヘッダー格納
        // Dim fldname As String = ToStr(wsheet.Cells(startrow - 1, cntjj).Value)

        // '移行値格納
        // Dim fldvalue As String = ToStr(wsheet.Cells(cntii + startrow, cntjj).Value)

        // Select Case fldname
        // Case "移行元フォーマット名称"
        // tmp_oldfmtname = fldvalue.Trim
        // Case "賃貸革命10フォーマットNo"
        // tmp_newfmtno = fldvalue.Trim
        // End Select

        // Next

        // 'ハッシュテーブル格納
        // If tmp_newfmtno <> "" Then
        // '20160527 指摘事項対応 -chg sta
        // 'Hash_Rel_FBFmt.Add(tmp_oldfmtname, tmp_newfmtno)
        // If Hash_Rel_FBFmt.Contains(tmp_oldfmtname) = False Then
        // Hash_Rel_FBFmt.Add(tmp_oldfmtname, tmp_newfmtno)
        // End If
        // '20160527 指摘事項対応 -chg end
        // End If

        // Next

        // 'Excelファイル終了設定
        // Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        // End Sub
        // 20160609 紐付設定値取得に伴う修正 -del end

    }

    #endregion


}