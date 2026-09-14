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

    #region 物件分類マスタ

    // 20160525 紐付項目の挿入処理の追加 -del sta
    // Public Class M_bk_rui_Repository
    // '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    // 'Public Class SubConv
    // '    Implements IConv

    // '    ''' <summary>
    // '    ''' 【変数→DB(INSERT)】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <param name="model_bk_rui"></param>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Cnv_Db(ByVal sqlcnnv10 As SqlConnection, ByVal model_cvitem As Object, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Cnv_Db

    // '        Dim reptbl As String                                '置換TBL名
    // '        Dim repfld As String                                '置換項目名
    // '        Dim repprm As String                                '置換パラメータ
    // '        Dim rowcnt As Integer                               '書込行数
    // '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    // '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    // '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    // '        '書込行数取得
    // '        rowcnt = model_cvitem.ItemCnt

    // '        'プログレスバー初期化
    // '        Dim pgbtotalcnt As Integer = rowcnt
    // '        Call obj_pgb.pgbInitPart(pgbtotalcnt)

    // '        'テーブル/フィールド名取得
    // '        reptbl = "m_bk_rui"
    // '        repfld = "bk_ruino,bk_ruiname,bk_ruibiko,bk_ruiuseflg,sincyoku_keiyakukbn," & _
    // '                 "sincyoku_kosinkbn,sincyoku_kaiyakukbn,history,homemate_ruikbn,bk_endofmonthflg," & _
    // '                 "ikkatukariage_siwakekbn"

    // '        With model_cvitem

    // '            For cntii = 0 To rowcnt - 1

    // '                '中断処理
    // '                Application.DoEvents()
    // '                If CancelFlg Then
    // '                    rtn = False
    // '                    Return rtn
    // '                End If

    // '                If .Vari_tmp_skipflg(cntii) = False Then

    // '                    'ハッシュテーブル作成
    // '                    Dim hash_cvitem As New SafeDictionary<string, string>
    // '                    'hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(repfld, model_cvitem, cntii)

    // '                    '書込処理
    // '                    'Call CVDBInsert.CVitem_Insert(sqlcnnv10, reptbl, repfld, hash_cvitem)

    // '                End If

    // '                'プログレスバー更新/進捗率表示
    // '                Dim pgbcnt As Integer = cntii + 1
    // '                Call obj_pgb.pgbsettingPart(pgbcnt)
    // '                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)

    // '            Next

    // '        End With

    // '        'Update実行モジュールへ遷移
    // '        Call Cnv_Db_Sub(sqlcnnv10)

    // '        '返却
    // '        Return rtn

    // '    End Function

    // '    ''' <summary>
    // '    ''' 【移行先DB初期化】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <remarks></remarks>
    // '    Public Sub Initialize_Table(sqlcnnv10 As SqlConnection) Implements IConv.Initialize_Table

    // '        '初期化対象外

    // '    End Sub

    // '    ''' <summary>
    // '    ''' 【仮テーブル→変数】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <param name="syorikomok"></param>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Set_Vari2(ByVal sheetname As String, ByRef model_cvitem As Object, ByVal sqlcnnv10 As SqlConnection) As Boolean Implements IConv.Set_Vari2

    // '        Dim reccnt As Integer                               '抽出レコード件数
    // '        Dim fldname As String                               '該当項目名
    // '        Dim fldvalue As String                              '該当値
    // '        Dim readtbl As New DataTable                        'テーブル格納用
    // '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    // '        Dim pgbcnt As Integer = 1                           'プログレスバー更新用カウンター
    // '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    // '        Dim logvalue As String                              'ログ出力用
    // '        Dim tblname As String = "tmp_bkrui_mst"             '対象テーブル名
    // '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    // '        '移行元情報取得
    // '        reccnt = DBExec.Exec_DataTable(Me.Get_UseQry(), sqlcnnv10, readtbl, rtn)

    // '        '移行元情報取得時にエラーが生じた場合は処理を抜ける
    // '        If Not rtn Then
    // '            Return rtn
    // '        End If

    // '        '初期化
    // '        model_cvitem = New Njc.Model.M_bk_rui_Model(reccnt - 1)
    // '        Dim tmp_duplicatechk As New Njc.Model.CommonModel(reccnt - 1)

    // '        'プログレスバー初期化
    // '        Dim pgbtotalcnt As Integer = (readtbl.Rows.Count) * (readtbl.Columns.Count)
    // '        Call obj_pgb.pgbInitPart(pgbtotalcnt)

    // '        '==========================
    // '        'データ取得　→　変数格納
    // '        '==========================
    // '        With model_cvitem

    // '            For cntii As Integer = 0 To readtbl.Rows.Count - 1

    // '                'DoEvents
    // '                Application.DoEvents()
    // '                If CancelFlg Then
    // '                    rtn = False
    // '                    Return rtn
    // '                End If

    // '                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

    // '                    '項目名取得
    // '                    fldname = readtbl.Columns(cntjj).ColumnName

    // '                    '登録値取得
    // '                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

    // '                    'ログ用
    // '                    logvalue = "賃貸革命10物件分類№ = " & Typ.ToStr(readtbl.Rows(cntii).Item(3)).Trim & _
    // '                               " " & _
    // '                               "移行元物件分類名称 = " & Typ.ToStr(readtbl.Rows(cntii).Item(1)).Trim

    // '                    '各項目値→変数格納
    // '                    Select Case fldname
    // '                        Case "賃貸革命10物件分類№"
    // '                            .Vari_Bk_ruino(cntii) = DataChk.Chk_DataInt(fldvalue, 1, 9999)
    // '                        Case "移行元物件分類名称"
    // '                            .Vari_Bk_ruiname(cntii) = DataChk.Chk_DataString(fldvalue, 50)
    // '                        Case "移行元物件分類備考"
    // '                            .Vari_Bk_ruibiko(cntii) = DataChk.Chk_DataString(fldvalue, 50)
    // '                    End Select

    // '                    '固定値設定
    // '                    .Vari_UseFlg(cntii) = 1
    // '                    .Vari_SincyokuKeiyakuKbn(cntii) = 0
    // '                    .Vari_SincyokuKosinKbn(cntii) = 0
    // '                    .Vari_SincyokuKaiyakuKbn(cntii) = 0
    // '                    .Vari_HomemateRuiKbn(cntii) = 0
    // '                    .Vari_History(cntii) = DefHistory
    // '                    .Vari_tmp_skipflg(cntii) = 0

    // '                    'プログレスバー更新/進捗率表示
    // '                    Call obj_pgb.pgbsettingPart(pgbcnt)
    // '                    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
    // '                    pgbcnt = pgbcnt + 1

    // '                Next

    // '            Next

    // '        End With

    // '        '返却
    // '        Return rtn

    // '    End Function

    // '    ''' <summary>
    // '    ''' 【変数→DB(UPDATE)】
    // '    ''' </summary>
    // '    ''' <remarks></remarks>
    // '    Public Sub Cnv_Db_Sub(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection)

    // '        'コマンドセット
    // '        Dim tmpcomtxt As String = Get_UseQry_Update()
    // '        Dim tmpsqlcom As New SqlCommand(tmpcomtxt, sqlcnnv10)
    // '        Dim rowsAffected As Integer = tmpsqlcom.ExecuteNonQuery()

    // '        '終了処理
    // '        tmpsqlcom.Dispose()

    // '    End Sub

    // '    ''' <summary>
    // '    ''' 【抽出クエリ】2015.07.17 sol レビュー後修正_レビュー№312
    // '    ''' </summary>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Get_UseQry(ByRef sortkeycnt As Integer) As String Implements IConv.Get_UseQry

    // '        Dim tmp_sql As String = ""

    // '        tmp_sql = tmp_sql & " SELECT "
    // '        tmp_sql = tmp_sql & " 	* "
    // '        tmp_sql = tmp_sql & " FROM tmp_bkrui_mst "
    // '        tmp_sql = tmp_sql & " WHERE [賃貸革命10物件分類№] NOT IN (SELECT bk_ruino FROM m_bk_rui) "
    // '        tmp_sql = tmp_sql & " AND   [賃貸革命10物件分類№] <> 0 "
    // '        tmp_sql = tmp_sql & " ORDER BY [移行元物件分類№] "

    // '        Return tmp_sql

    // '    End Function

    // '    ''' <summary>
    // '    ''' 【更新クエリ】
    // '    ''' </summary>
    // '    ''' <remarks></remarks>
    // '    Public Function Get_UseQry_Update() As String

    // '        Dim tmp_sql As String = ""

    // '        tmp_sql = tmp_sql & " UPDATE m_bk_rui "
    // '        tmp_sql = tmp_sql & " 	SET bk_ruibiko = TMP.移行元物件分類備考 "
    // '        tmp_sql = tmp_sql & " FROM tmp_bkrui_mst AS TMP "
    // '        tmp_sql = tmp_sql & " LEFT JOIN m_bk_rui AS MBKR "
    // '        tmp_sql = tmp_sql & " ON TMP.[賃貸革命10物件分類№] = MBKR.bk_ruino; "

    // '        Return tmp_sql

    // '    End Function

    // '    Public Function Get_BaseKey() As Object Implements IConv.Get_BaseKey

    // '    End Function

    // '                Public Function Set_Vari( _
    // 'ByVal sqlcnnv10 As SqlConnection, _
    // 'ByVal filename As String, ByVal sheetname As String, _
    // 'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    // ') As Boolean Implements IConv.Set_Vari

    // '    End Function

    // '                Public Function Set_Vari( _
    // 'ByVal sqlcnnv10 As SqlConnection, _
    // 'ByVal filename As String, ByVal sheetname As String, _
    // 'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    // ') As Boolean Implements IConv.Set_Vari

    // '    End Function

    // '    Public Function Write_IntermediateFile(sqlcnnv7 As SqlConnection, filename As String, sheetname As String, model_cvitem As Object) As Boolean Implements IConv.Write_IntermediateFile

    // '    End Function

    // 'End Class
    // '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    // End Class
    // 20160525 紐付項目の挿入処理の追加 -del end

    #endregion

    #region バス交通マスタ

    public class M_buskotu_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】20161004 既存中間→DB書込処理修正
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="filename"></param>
            /// <param name="sheetname"></param>
            /// <param name="midrowcnt"></param>
            /// <param name="cvrowcnt"></param>
            /// <param name="conditioncnt"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {

                var tmp_cvcnt = default(int);                                        // 移行件数格納
                var tmp_condcnt = default(int);                                      // 調整件数格納
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値
                var model_cvitem = new Model.M_buskotu_Model();               // 移行値格納用モデル初期化
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

                // テーブル名/フィールド名セット
                string tblname = "m_buskotu";
                string fldnamegrp = "buskotu_no,buskotu_name,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // 読込処理

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
                            case "バス交通No":
                                {
                                    model_cvitem.Vari_Buskotu_no = fldvalue.Trim();
                                    break;
                                }
                            case "バス会社名":
                                {
                                    model_cvitem.Vari_Buskotu_name = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック (移行項目が親なので親マスタチェックは不要 list_basekeydata はダミーで入れておく)
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    string errstr = "";
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (skipflg == false)
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
            /// 【中間ファイル→変数】旧処理のバックアップ
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="filename"></param>
            /// <param name="sheetname"></param>
            /// <param name="midrowcnt"></param>
            /// <param name="cvrowcnt"></param>
            /// <param name="conditioncnt"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public bool Set_Vari_back(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
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

                var model_cvitem = new Model.M_buskotu_Model();               // 移行値格納用モデル初期化
                int keycol = 1;                                       // キー列

                // ************************
                // 作業準備
                // ************************

                // Excelファイル初期設定                          'kakaka バス交通マスタ以外も共通メソッドで纏められそうな場合はまとめて下さい。
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.MiddleDirPath, filename, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // プログレスバー初期化
                // 2016.02.22 プログレスバーの表示修正 -chg sta
                // Dim pgbtotalcnt As Integer = rowcnt
                // Call obj_pgb.pgbInitPart(pgbtotalcnt)
                int pgbtotalcnt = 0;          // プログレスバー総件数初期化
                int pgbbasecnt = 100;         // 実件数で表示するか否かの基準値
                // 2016.03.28 プログレスバーの表示修正(再) -del
                // Dim pgbsimpleflg As Boolean = False     '実件数で表示するか否かのフラグ (True:基準値で割った件数 False:実件数)

                if (rowcnt <= pgbbasecnt)
                {
                    pgbtotalcnt = rowcnt;
                }
                else
                {
                    pgbtotalcnt = (int)Math.Round(Math.Ceiling(rowcnt / (double)pgbbasecnt));
                    // 2016.03.28 プログレスバーの表示修正(再) -del
                    // pgbsimpleflg = True
                }
                obj_pgb.pgbInitPart(pgbtotalcnt);
                // 2016.02.22 プログレスバーの表示修正 -chg end

                // テーブル名/フィールド名セット
                string tblname = "m_buskotu";
                string fldnamegrp = "buskotu_no,buskotu_name,history";

                // 追加コンバート時の重複チェック用に既存データのキーを取得
                if (!CommonModule.InitDBFlg)
                {
                    Get_ExistData(sqlcnnv10, tblname, ref list_chkduplicate);
                }

                // ************************
                // 処理開始
                // ************************


                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;    // 2016.02.22 移行項目取得方法修正

                // キーヘッダー名取得
                string fldname_key = Conversions.ToString(headervalue(startrow - 1, keycol));

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
                    var datarowvalue = (object[,])((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[cntii + startrow - 1, 1], wsheet.Cells[cntii + startrow - 1, columncnt]]).Value;   // 2016.02.22 移行項目取得方法修正

                    // キー値取得
                    string fldvalue_key = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol]);

                    // ログ出力用
                    string str_logkey = fldname_key + " = " + fldvalue_key;

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();  // 2016.02.22 移行項目取得方法修正
                        string fldvalue = "";                                             // 2016.02.22 移行項目取得方法修正
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "バス交通No":
                                {
                                    model_cvitem.Vari_Buskotu_no = fldvalue.Trim();
                                    break;
                                }
                            case "バス会社名":
                                {
                                    model_cvitem.Vari_Buskotu_name = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;

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
                    // 2016.03.28 プログレスバーの表示修正(再) -chg sta
                    // ↓↓↓旧srcコメントアウト↓↓↓
                    // '2016.02.22 プログレスバーの表示修正 -chg sta
                    // 'Dim pgbcnt As Integer = cntii + 1
                    // 'Call obj_pgb.pgbsettingPart(pgbcnt)                         'kakaka レスポンス改善の為、以下参考(他の情報(ブロック)も同様)
                    // 'Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)            'kakaka 進捗間隔(進み方)を次の通りに変更 ⇒ 実件数/100<=1…実件数、実件数/100>1…実件数/100の切り上げ(実件数が1015の場合、10.15なので11回)

                    // '件数取得
                    // Dim pgbcnt As Integer = 0
                    // If pgbsimpleflg = False Then        '100件以下の場合は実件数を取得                                                  'kakaka4 0322_2050 どっちでもいいですが、rowcnt<=pgbbasecnt でもよかったような…
                    // pgbcnt = cntii + 1
                    // ElseIf rowcnt <> cntii + 1 Then     '100件より大きい、かつ最終レコードに達していない場合は100件毎に値を取得         'kakaka4 0322_2050 これだと比較値+1ではない場合って意味では？上から順に条件を考えた場合でも本来の条件の意味が違うような…(rowcnt>pgbbasecnt AND rowcnt<=cntii )
                    // Call EtcMethod.Get_MultipleFlg(cntii + 1, pgbbasecnt, pgbcnt)
                    // ElseIf rowcnt = cntii + 1 Then      '100件より大きい、かつ最終レコードに達した場合はプログレスバーの総件数を取得    'kakaka 0322_20504 これだと比較値+1と同じって意味では？(rowcnt>pgbbasecnt AND rowcnt>cntii )
                    // pgbcnt = pgbtotalcnt
                    // End If

                    // '表示
                    // If pgbcnt <> 0 Then
                    // Call obj_pgb.pgbsettingPart(pgbcnt)
                    // Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
                    // End If
                    // '2016.02.22 プログレスバーの表示修正 -chg end
                    // ↑↑↑旧srcコメントアウト↑↑↑

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

                    // 2016.03.28 プログレスバーの表示修正(再) -chg end

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

                string tmp_sql = " SELECT buskotu_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region バス停マスタ

    public class M_buskotu_stop_Repository
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
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用   
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値
                var model_cvitem = new Model.M_buskotu_stop_Model();          // 移行値格納用モデル初期化
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
                string tblname_base = "m_buskotu";
                string tblname = "m_buskotu_stop";
                string fldnamegrp = "buskotu_no,sortorder,keito_name,busstop_name";

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
                // 読込処理

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
                            case "バス交通No":
                                {
                                    model_cvitem.Vari_Buskotu_no = fldvalue.Trim();
                                    break;
                                }
                            case "バス停No":
                                {
                                    model_cvitem.Vari_Sortorder = fldvalue.Trim();
                                    break;
                                }
                            case "系統名":
                                {
                                    model_cvitem.Vari_Keito_name = fldvalue.Trim();
                                    break;
                                }
                            case "バス停名":
                                {
                                    model_cvitem.Vari_Busstop_name = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    tmp_hash = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                    // データチェック
                    bool skipflg = false;
                    var hash_cvitem = new SafeDictionary<string, string>();
                    string errstr = "";
                    var hash_log = new SafeDictionary<string, string>();
                    skipflg = !DataChk.Chk_DataPerItem(sqlcnnv10, tblname, list_chkduplicate, list_basekeydata, tmp_hash, ref hash_cvitem, ref tmp_condcnt, ref hash_log);

                    // 書込処理
                    if (skipflg == false)
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

            // Dim model_cvitem As New Njc.Model.M_buskotu_stop_Model          '移行値格納用モデル初期化
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
            // Dim tblname_base As String = "m_buskotu"
            // Dim tblname As String = "m_buskotu_stop"
            // Dim fldnamegrp As String = "buskotu_no,sortorder,keito_name,busstop_name"

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
            // Case "バス交通No"
            // .Vari_Buskotu_no = fldvalue.Trim
            // Case "バス停No"
            // .Vari_Sortorder = fldvalue.Trim
            // Case "系統名"
            // .Vari_Keito_name = fldvalue.Trim
            // Case "バス停名"
            // .Vari_Busstop_name = fldvalue.Trim
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

            /// <summary>
            /// 親マスタ取得→リスト格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT buskotu_no FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,buskotu_no) + '-' + CONVERT(varchar,sortorder) FROM  " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 鍵タイトルマスタ

    public class M_kagi_title_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【中間ファイル→変数】 2016.04.06 鍵情報の取得修正 (メソッド内をほぼ全て修正)
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

                var model_cvitem = new Model.M_kagi_title_Model();            // 移行値格納用モデル初期化
                // 20160613 鍵情報の取得処理修正 -chg sta
                // Dim keycol As Integer = 2                                       'メインキー列 (共用専用で振り分けることでNoが変化するため名称をキーにしておく)
                int keycol = 1;                                       // メインキー列 (Noをキーに変更)
                // 20160613 鍵情報の取得処理修正 -chg end

                // ************************
                // 作業準備
                // ************************

                // Excelファイル初期設定 (紐付ファイルの読込)
                rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, sheetname);

                // Excelファイル設定時にエラーが生じた際は処理を抜ける
                if (!rtn)
                {
                    return rtn;
                }

                // テーブル名/フィールド名セット
                string tblname = "m_kagi_title";
                string fldnamegrp = "kagi_kbn,kagi_no,kagi_name,history";

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
                string fldname_key = Conversions.ToString(headervalue(startrow - 1, keycol));

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

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "共用チェック":
                                {
                                    model_cvitem.Vari_Kagi_kbn = Conversions.ToString(Interaction.IIf(fldvalue == "True", 1, 2));
                                    break;
                                }
                            case "鍵No":
                                {
                                    model_cvitem.Vari_Kagi_no = fldvalue;
                                    break;
                                }
                            case "鍵名称":
                                {
                                    model_cvitem.Vari_Kagi_name = fldvalue;
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

                // 鍵No一括成形
                int tmpcnt = 0;
                string kaginoupdatesql = Get_UseQry_Update();
                DBExec.Exec_NonQuery(sqlcnnv10, kaginoupdatesql, ref tmpcnt);

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

                // 20160516 鍵タイトルマスタの重複チェック修正 -chg sta
                // Dim tmp_sql As String = " SELECT CONVERT(varchar,kagi_kbn) + '-' + CONVERT(varchar,kagi_no) FROM  " & tblname
                string tmp_sql = " SELECT kagi_name FROM  " + tblname;
                // 20160516 鍵タイトルマスタの重複チェック修正 -chg end
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

            /// <summary>
            /// 鍵Noの一括更新クエリ 2016.04.06 鍵情報の取得修正 (新規追加)
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_Update()
            {

                string tmp_sql = "";

                tmp_sql = tmp_sql + " UPDATE m_kagi_title SET ";
                tmp_sql = tmp_sql + " 	m_kagi_title.kagi_no = MKT.[連番] ";
                tmp_sql = tmp_sql + " FROM m_kagi_title ";
                tmp_sql = tmp_sql + " LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 kagi_kbn ";
                tmp_sql = tmp_sql + " 			,kagi_no ";
                tmp_sql = tmp_sql + " 			,ROW_NUMBER()OVER(PARTITION BY kagi_kbn ORDER BY kagi_kbn,kagi_no) AS [連番] ";
                tmp_sql = tmp_sql + " 		FROM m_kagi_title ";
                tmp_sql = tmp_sql + " 	)  AS MKT ";
                tmp_sql = tmp_sql + " ON  m_kagi_title.kagi_kbn = MKT.kagi_kbn ";
                tmp_sql = tmp_sql + " AND m_kagi_title.kagi_no = MKT.kagi_no ";
                tmp_sql = tmp_sql + " ; ";

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 箇所クレーム分類マスタ

    public class M_claim_rui_Repository
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

                var model_cvitem = new Model.M_claim_rui_Model();       // 移行値格納用モデル初期化
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
                string tblname = "m_claim_rui";
                string fldnamegrp = "claim_ruikbn,claim_ruino,claim_ruisortorder,claim_name,claim_useflg," + "history";

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
                string fldname_key = Conversions.ToString(headervalue(startrow - 1, keycol));

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

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "箇所・クレーム分類区分":
                                {
                                    model_cvitem.Vari_Claim_ruikbn = fldvalue.Trim();
                                    break;
                                }
                            case "分類No":
                                {
                                    model_cvitem.Vari_Claim_ruino = fldvalue.Trim();
                                    model_cvitem.Vari_Claim_ruisortorder = fldvalue.Trim();
                                    break;
                                }
                            case "名称":
                                {
                                    model_cvitem.Vari_Claim_name = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Claim_useflg = 1.ToString();
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
                            list_chkduplicate.Add(fldvalue_key);                             // kakaka メモ：Cnv_DbでINSERT失敗の場合、ここで重複データを貯めこむ
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

                string tmp_sql = " SELECT claim_ruino FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 特約マスタ

    public class M_tokuyaku_Repository
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
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用  
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値
                var model_cvitem = new Model.M_tokuyaku_Model();     // 移行値格納用モデル初期化
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
                string tblname = "m_tokuyaku";
                string fldnamegrp = "tokuyaku_grpno,tokuyaku_no,tokuyaku_title,tokuyaku_template,history";

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
                        fldvalue = Typ.ToStr(readtbl.Rows[cntii][cntjj]);

                        // 各項目値→変数格納
                        switch (fldname ?? "")
                        {
                            case "特約区分":
                                {
                                    model_cvitem.Vari_Tokuyaku_grpno = fldvalue;
                                    break;
                                }
                            case "特約No":
                                {
                                    model_cvitem.Vari_Tokuyaku_no = fldvalue;
                                    break;
                                }
                            case "特約タイトル":
                                {
                                    model_cvitem.Vari_Tokuyaku_title = fldvalue;
                                    break;
                                }
                            case "特約詳細":
                                {
                                    model_cvitem.Vari_Tokuyaku_template = fldvalue;
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

            // Dim model_cvitem As New Njc.Model.M_tokuyaku_Model     '移行値格納用モデル初期化
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
            // Dim tblname As String = "m_tokuyaku"
            // Dim fldnamegrp As String = "tokuyaku_grpno,tokuyaku_no,tokuyaku_title,tokuyaku_template,history"

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
            // '20160706 特約マスタの文字列移行処理修正 -chg sta
            // 'fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
            // fldvalue = datarowvalue(startrow - 1, cntjj).ToString
            // '20160706 特約マスタの文字列移行処理修正 -chg end
            // End If

            // Select Case fldname
            // Case "特約区分"
            // .Vari_Tokuyaku_grpno = fldvalue
            // Case "特約No"
            // .Vari_Tokuyaku_no = fldvalue
            // Case "特約タイトル"
            // .Vari_Tokuyaku_title = fldvalue
            // Case "特約詳細"
            // .Vari_Tokuyaku_template = fldvalue
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

                string tmp_sql = " SELECT tokuyaku_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 契約分類マスタ

    // 20160525 紐付項目の挿入処理の追加 -del sta
    // Public Class M_ky_rui_Repository

    // Public Class SubConv
    // Implements IConv

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

    // Dim chgcolor As New Njc.Common.Translate_Color                  '色変換用
    // Dim model_cvitem As New Njc.Model.M_ky_rui_Model                '移行値格納用モデル初期化
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
    // Dim tblname As String = "m_ky_rui"
    // Dim fldnamegrp As String = "ky_ruino,ky_ruiname,ky_ruibiko,ky_ruitutimm,ky_ruicolor," & _
    // "teisyaku_flg,history,useflg"

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
    // Case "契約分類No"
    // .Vari_Ky_ruino = fldvalue.Trim
    // Case "契約分類名"
    // .Vari_Ky_ruiname = fldvalue.Trim
    // Case "備考"
    // .Vari_Ky_ruibiko = fldvalue.Trim
    // Case "更新・解約通知期間(ヶ月)"
    // .Vari_Ky_ruitutimm = fldvalue.Trim
    // Case "契約分類色"
    // .Vari_Ky_ruicolor = chgcolor.Chg_ColorValue(IIf(fldvalue.Trim = "", 0, fldvalue.Trim))
    // Case "定期借地借家権契約扱い有無"
    // .Vari_Teisyaku_flg = fldvalue.Trim
    // End Select

    // Next

    // '固定値
    // .Vari_Useflg = 1
    // .Vari_History = DefHistory
    // '
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

    // Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As SafeDictionary<string, string>, ByRef hash_chkafter As SafeDictionary<string, string>, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

    // End Function

    // Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

    // End Sub

    // Public Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

    // End Function

    // ''' <summary>
    // ''' 既存データのキーを取得してリストへ格納
    // ''' </summary>
    // ''' <param name="sqlcnnv10"></param>
    // ''' <param name="tblname"></param>
    // ''' <param name="list_existdata"></param>
    // ''' <remarks></remarks>
    // Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

    // Dim tmp_sql As String = " SELECT ky_ruino FROM " & tblname
    // Dim flg As Boolean = True

    // flg = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)

    // End Sub

    // End Class

    // End Class
    // 20160525 紐付項目の挿入処理の追加 -del end

    #endregion

    #region 保険種類マスタ

    public class M_hoken_rui_Repository
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
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用 
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値

                var model_cvitem = new Model.M_hoken_rui_Model();             // 移行値格納用モデル初期化
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
                string tblname = "m_hoken_rui";
                string fldnamegrp = "hoken_ruino,hoken_ruiname,hoken_ruikana,biko_kihon,history," + "useflg";

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
                            case "保険種類No":
                                {
                                    model_cvitem.Vari_Hoken_ruino = fldvalue.Trim();
                                    break;
                                }
                            case "保険種類名":
                                {
                                    model_cvitem.Vari_Hoken_ruiname = fldvalue.Trim();
                                    break;
                                }
                            case "保険種類カナ":
                                {
                                    model_cvitem.Vari_Hoken_ruikana = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

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

            // Dim model_cvitem As New Njc.Model.M_hoken_rui_Model             '移行値格納用モデル初期化
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
            // Dim tblname As String = "m_hoken_rui"
            // Dim fldnamegrp As String = "hoken_ruino,hoken_ruiname,hoken_ruikana,biko_kihon,history," & _
            // "useflg"

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
            // Case "保険種類No"
            // .Vari_Hoken_ruino = fldvalue.Trim
            // Case "保険種類名"
            // .Vari_Hoken_ruiname = fldvalue.Trim
            // Case "保険種類カナ"
            // .Vari_Hoken_ruikana = fldvalue.Trim
            // Case "備考"
            // .Vari_Biko_kihon = fldvalue.Trim
            // End Select

            // Next

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

                string tmp_sql = " SELECT hoken_ruino FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 入金区分マスタ

    // 20160525 紐付項目の挿入処理の追加 -del sta
    // Public Class M_nkbn_Repository
    // '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    // 'Public Class SubConv
    // '    Implements IConv

    // '    ''' <summary>
    // '    ''' 【変数→V10DB】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <param name="model_nkbn"></param>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Cnv_Db(ByVal sqlcnnv10 As SqlConnection, ByVal model_cvitem As Object, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Cnv_Db

    // '        Dim reptbl As String                                '置換TBL名
    // '        Dim repfld As String                                '置換項目名
    // '        Dim repprm As String                                '置換パラメータ
    // '        Dim rowcnt As Integer                               '書込行数
    // '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    // '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    // '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    // '        '書込行数取得
    // '        rowcnt = model_cvitem.ItemCnt

    // '        'プログレスバー初期化
    // '        Dim pgbtotalcnt As Integer = rowcnt
    // '        Call obj_pgb.pgbInitPart(rowcnt)

    // '        'テーブル/フィールド名取得
    // '        reptbl = "m_nkbn"
    // '        repfld = "nkbn_no,nkbn_order,nkbn_name,nkbn_shortname,nkbn_zokusei," & _
    // '                 "history"

    // '        With model_cvitem

    // '            For cntii = 0 To rowcnt - 1

    // '                '中断処理
    // '                Application.DoEvents()
    // '                If CancelFlg Then
    // '                    rtn = False
    // '                    Return rtn
    // '                End If

    // '                If .Vari_tmp_skipflg(cntii) = False Then

    // '                    'ハッシュテーブル作成
    // '                    Dim hash_cvitem As New SafeDictionary<string, string>
    // '                    'hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(repfld, model_cvitem, cntii)

    // '                    '書込処理
    // '                    'Call CVDBInsert.CVitem_Insert(sqlcnnv10, reptbl, repfld, hash_cvitem)

    // '                End If

    // '                'プログレスバー更新/進捗率表示
    // '                Dim pgbcnt As Integer = cntii + 1
    // '                Call obj_pgb.pgbsettingPart(pgbcnt)
    // '                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)

    // '            Next

    // '        End With

    // '        '返却
    // '        Return rtn

    // '    End Function

    // '    ''' <summary>
    // '    ''' 【移行先DB初期化】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <remarks></remarks>
    // '    Public Sub Initialize_Table(sqlcnnv10 As SqlConnection) Implements IConv.Initialize_Table

    // '        '初期化対象外

    // '    End Sub

    // '    ''' <summary>
    // '    ''' 【仮テーブル→変数】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <param name="syorikomok"></param>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Set_Vari2(ByVal sheetname As String, ByRef model_cvitem As Object, ByVal sqlcnnv10 As SqlConnection) As Boolean Implements IConv.Set_Vari2

    // '        Dim reccnt As Integer                               '抽出レコード件数
    // '        Dim fldname As String                               '該当項目名
    // '        Dim fldvalue As String                              '該当値
    // '        Dim readtbl As New DataTable                        'テーブル格納用
    // '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    // '        Dim pgbcnt As Integer = 1                           'プログレスバー更新用カウンター
    // '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    // '        Dim logvalue As String                              'ログ出力用
    // '        Dim tblname As String = "tmp_nkbn_mst"              '対象テーブル名
    // '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    // '        '移行元情報取得
    // '        reccnt = DBExec.Exec_DataTable(Me.Get_UseQry(), sqlcnnv10, readtbl, rtn)

    // '        '移行元情報取得時にエラーが生じた場合は処理を抜ける
    // '        If Not rtn Then
    // '            Return rtn
    // '        End If

    // '        '初期化
    // '        model_cvitem = New Njc.Model.M_nkbn_Model(reccnt - 1)
    // '        Dim tmp_duplicatechk As New Njc.Model.CommonModel(reccnt - 1)

    // '        'プログレスバー初期化
    // '        Dim pgbtotalcnt As Integer = (readtbl.Rows.Count) * (readtbl.Columns.Count)
    // '        Call obj_pgb.pgbInitPart(pgbtotalcnt)

    // '        '==========================
    // '        'データ取得　→　変数格納
    // '        '==========================
    // '        With model_cvitem

    // '            For cntii As Integer = 0 To readtbl.Rows.Count - 1

    // '                'DoEvents
    // '                Application.DoEvents()
    // '                If CancelFlg Then
    // '                    rtn = False
    // '                    Return rtn
    // '                End If

    // '                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

    // '                    '項目名取得
    // '                    fldname = readtbl.Columns(cntjj).ColumnName

    // '                    '登録値取得
    // '                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

    // '                    'ログ用
    // '                    logvalue = "賃貸革命10入金区分№ = " & Typ.ToStr(readtbl.Rows(cntii).Item(2)).Trim & _
    // '                               " " & _
    // '                               "移行元入金区分名称 = " & Typ.ToStr(readtbl.Rows(cntii).Item(1)).Trim

    // '                    '各項目値→変数格納
    // '                    Select Case fldname
    // '                        Case "賃貸革命10入金区分№"
    // '                            .Vari_Nkbn_no(cntii) = DataChk.Chk_DataInt(fldvalue, 1, 20)
    // '                            .Vari_Nkbn_order(cntii) = DataChk.Chk_DataInt(fldvalue, 1, 20)
    // '                        Case "移行元入金区分名称"
    // '                            .Vari_Nkbn_name(cntii) = DataChk.Chk_DataString(fldvalue, 50)
    // '                        Case "賃貸革命10入金区分属性№"
    // '                            .Vari_Nkbn_zokusei(cntii) = fldvalue
    // '                    End Select

    // '                    '固定値設定
    // '                    .Vari_Nkbn_shortname(cntii) = ""
    // '                    .Vari_History(cntii) = DefHistory
    // '                    .Vari_tmp_skipflg(cntii) = 0

    // '                    'プログレスバー更新/進捗率表示
    // '                    Call obj_pgb.pgbsettingPart(pgbcnt)
    // '                    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
    // '                    pgbcnt = pgbcnt + 1

    // '                Next

    // '            Next

    // '        End With

    // '        '返却
    // '        Return rtn

    // '    End Function

    // '    ''' <summary>
    // '    ''' 【抽出クエリ】2015.07.17 sol レビュー後修正_レビュー№312
    // '    ''' </summary>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Get_UseQry(ByRef sortkeycnt As Integer) As String Implements IConv.Get_UseQry

    // '        Dim tmp_sql As String = ""

    // '        tmp_sql = tmp_sql & " SELECT "
    // '        tmp_sql = tmp_sql & " 	* "
    // '        tmp_sql = tmp_sql & " FROM tmp_nkbn_mst "
    // '        tmp_sql = tmp_sql & " WHERE [賃貸革命10入金区分№] NOT IN (SELECT nkbn_no FROM m_nkbn) "
    // '        tmp_sql = tmp_sql & " AND   [賃貸革命10入金区分№] <> 0 "
    // '        tmp_sql = tmp_sql & " ORDER BY [移行元入金区分№] "

    // '        Return tmp_sql

    // '    End Function

    // '    ''' <summary>
    // '    ''' 【フィールド名と移行値を紐付けるハッシュテーブル作成】
    // '    ''' </summary>
    // '    ''' <param name="model_cvitem"></param>
    // '    ''' <param name="cntii"></param>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Get_Hash_CVItem(ByVal model_cvitem As Object, ByVal cntii As Integer) As SafeDictionary<string, string> Implements IConv.Get_Hash_CVItem

    // '        Dim rtn_hash As New SafeDictionary<string, string>

    // '        With model_cvitem

    // '            rtn_hash.Add("nkbn_no", .Vari_Nkbn_no(cntii))
    // '            rtn_hash.Add("nkbn_order", .Vari_Nkbn_order(cntii))
    // '            rtn_hash.Add("nkbn_name", .Vari_Nkbn_name(cntii))
    // '            rtn_hash.Add("nkbn_shortname", .Vari_Nkbn_shortname(cntii))
    // '            rtn_hash.Add("nkbn_zokusei", .Vari_Nkbn_zokusei(cntii))
    // '            rtn_hash.Add("history", .Vari_History(cntii))

    // '        End With

    // '        Return rtn_hash

    // '    End Function

    // '    Public Function Set_Vari1(sqlcnnv7 As SqlConnection, ByRef model_cvitem As Object) As Boolean Implements IConv.Set_Vari1

    // '    End Function

    // '    Public Function Set_Vari21(sheetname As String, ByRef model_cvitem As Object) As Boolean Implements IConv.Set_Vari2

    // '    End Function

    // '    Public Function Write_IntermediateFile(sheetname As String, model_cvitem As Object) As Boolean Implements IConv.Write_IntermediateFile

    // '    End Function

    // '    Public Function Get_BaseKey() As Object Implements IConv.Get_BaseKey

    // '    End Function

    // '                Public Function Set_Vari( _
    // 'ByVal sqlcnnv10 As SqlConnection, _
    // 'ByVal filename As String, ByVal sheetname As String, _
    // 'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    // ') As Boolean Implements IConv.Set_Vari

    // '    End Function

    // '                Public Function Set_Vari( _
    // 'ByVal sqlcnnv10 As SqlConnection, _
    // 'ByVal filename As String, ByVal sheetname As String, _
    // 'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    // ') As Boolean Implements IConv.Set_Vari

    // '    End Function

    // '    Public Function Write_IntermediateFile(sqlcnnv7 As SqlConnection, filename As String, sheetname As String, model_cvitem As Object) As Boolean Implements IConv.Write_IntermediateFile

    // '    End Function
    // 'End Class
    // '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    // End Class
    // 20160525 紐付項目の挿入処理の追加 -del end

    #endregion

    #region 学校区マスタ

    public class M_koku_add_Repository
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
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値
                var model_cvitem = new Model.M_koku_add_Model();              // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                  // サブキー1列
                int keycol_sub2 = 4;                                  // サブキー3列
                int keycol_sub3 = 5;                                  // サブキー4列
                int keycol_sub4 = 7;                                  // サブキー5列
                int keycol_sub5 = 8;                                  // サブキー6列

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
                string tblname = "m_koku_add";
                string fldnamegrp = "ken_no,si_no,no,add_cyo,add_banti," + "syogaku_name,cyugaku_name,history,add_cyome";

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
                    string tmp_addresscyotiiki = "";
                    string tmp_addresscyobanti = "";

                    // 'キー値取得
                    // Dim fldvalue_keymain As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_main - 1)).Trim
                    // Dim fldvalue_keysub As String = Typ.ToStr(readtbl.Rows(cntii).Item(keycol_sub - 1)).Trim
                    // Dim fldvalue_key As String = fldvalue_keymain & "-" & fldvalue_keysub

                    // 'ログ出力用
                    // Dim str_logkey As String = fldname_keymain & " = " & fldvalue_keymain & "、" & fldname_keysub & " = " & fldvalue_keysub

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
                            case "県No":
                                {
                                    model_cvitem.Vari_Ken_no = fldvalue;
                                    break;
                                }
                            case "市No":
                                {
                                    model_cvitem.Vari_Si_no = fldvalue;
                                    break;
                                }
                            case "学校区No":
                                {
                                    model_cvitem.Vari_No = fldvalue;
                                    break;
                                }
                            case "町":
                                {
                                    // 町地域はそのまま移行する(丁番地以降が分割対象)
                                    model_cvitem.Vari_Add_cyo = fldvalue.Trim();
                                    break;
                                }
                            case "丁目":
                                {
                                    tmp_addresscyobanti = fldvalue;
                                    break;
                                }
                            case "番地":
                                {
                                    model_cvitem.Vari_Add_banti = fldvalue;
                                    break;
                                }
                            case "小学校名":
                                {
                                    model_cvitem.Vari_Syogaku_name = fldvalue;
                                    break;
                                }
                            case "中学校名":
                                {
                                    model_cvitem.Vari_Cyugaku_name = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 20161009 改善対応：学校区丁目移行不具合 -chg sta
                    // '20160622 住所分割処理対応 -chg sta
                    // ''2016.04.06 学校区マスタ移行内容修正 -add sta
                    // ''住所変換/格納
                    // 'Dim tmp_address As String = tmp_addresscyotiiki & " " & tmp_addresscyobanti
                    // 'Dim hash_address As New SafeDictionary<string, string>
                    // 'hash_address = AddressChange.Get_CyoBanti(tmp_address)
                    // '.Vari_Add_cyo = hash_address("mati")
                    // '.Vari_Add_cyome = hash_address("cyome")
                    // ''2016.04.06 学校区マスタ移行内容修正 -add end
                    // '住所変換/格納
                    // Dim tmp_address As String = tmp_addresscyobanti
                    // Dim hash_address As New SafeDictionary<string, string>
                    // hash_address = AddressChange.Get_CyoBanti(tmp_address)
                    // .Vari_Add_cyome = hash_address("cyome")
                    // '20160622 住所分割処理対応 -chg end

                    // ※Get_CyoBantiメソッドでは連結された住所または「丁目」が付加されているデータでないと正しく移行できない為、強制で「丁目」を付加。
                    string tmp_address = Strings.Replace(Strings.Replace(Strings.Replace(tmp_addresscyobanti, "丁目", ""), "丁", ""), "町目", "");
                    tmp_address = Conversions.ToString(Interaction.IIf(Typ.IsNumeric(tmp_address), tmp_address + "丁目", tmp_addresscyobanti));
                    var hash_address = new SafeDictionary<string, string>();
                    hash_address = AddressChange.Get_CyoBanti(tmp_address);
                    model_cvitem.Vari_Add_cyome = Conversions.ToString(hash_address["cyome"]);
                    // 20161009 改善対応：学校区丁目移行不具合 -chg end

                    // 固定値
                    model_cvitem.Vari_History = CommonModule.DefHistory;

                    // 2016.04.26 メインの方へも反映させる修正 -add sta
                    // 町、丁目、学校名もキーとしているが、空でも登録可能であるため半角スペースを入れておく
                    if (string.IsNullOrEmpty(model_cvitem.Vari_Add_cyo))
                    {
                        model_cvitem.Vari_Add_cyo = " ";
                    }
                    if (string.IsNullOrEmpty(model_cvitem.Vari_Add_cyome))
                    {
                        model_cvitem.Vari_Add_cyome = " ";
                    }
                    if (string.IsNullOrEmpty(model_cvitem.Vari_Syogaku_name))
                    {
                        model_cvitem.Vari_Syogaku_name = " ";
                    }
                    if (string.IsNullOrEmpty(model_cvitem.Vari_Cyugaku_name))
                    {
                        model_cvitem.Vari_Cyugaku_name = " ";
                    }
                    // 2016.04.26 メインの方へも反映させる修正 -add end

                    // 2016.04.06 学校区マスタ移行内容修正 -add
                    // 住所分割後の値で重複判断を行うためここでキーを取得する
                    string fldvalue_key = model_cvitem.Vari_Ken_no + "-" + model_cvitem.Vari_Si_no + "-" + model_cvitem.Vari_Add_cyo + "-" + model_cvitem.Vari_Add_cyome + "-" + model_cvitem.Vari_Syogaku_name + "-" + model_cvitem.Vari_Cyugaku_name;





                    // 2016.04.06 学校区マスタ移行内容修正 -del
                    // キー取得後にログ出力用文字列を作成する
                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + model_cvitem.Vari_Ken_no + "、" + fldname_keysub1 + " = " + model_cvitem.Vari_Si_no + "、" + fldname_keysub2 + " = " + model_cvitem.Vari_Add_cyo + "、" + fldname_keysub3 + " = " + model_cvitem.Vari_Add_cyome + "、" + fldname_keysub4 + " = " + model_cvitem.Vari_Syogaku_name + "、" + fldname_keysub5 + " = " + model_cvitem.Vari_Cyugaku_name;





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

                string tmp_sql = " SELECT CONVERT(varchar,ken_no) + '-' + CONVERT(varchar,si_no) + '-' + CONVERT(varchar,no) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region エリアマスタ

    public class M_area_Repository
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
                // 20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del
                // Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
                var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
                var obj_com = new CommonRepository();                  // 共通処理用
                var list_chkduplicate = new List<string>();                    // 重複チェック用リスト
                var list_basekeydata = new List<string>();                     // 親マスタ有無チェック用リスト
                var sortlist_log = new SortedList<int, string>();          // ログ格納用オブジェクト
                int tmp_logcnt = 0;                                   // ログ出力時のソート用
                var tmp_hash = new SafeDictionary<string, string>();                                   // 作業用ハッシュテーブル
                var normalflg = default(bool);                                        // INSERT正常終了フラグ
                bool rtn = true;                                       // 戻り値
                var model_cvitem = new Model.M_area_Model();                  // 移行値格納用モデル初期化
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
                string tblname = "m_area";
                string fldnamegrp = "area_no,area_sortorder,area_name,area_useflg,history";

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
                            case "エリアNo":
                                {
                                    model_cvitem.Vari_Area_no = fldvalue.Trim();
                                    break;
                                }
                            case "エリア名":
                                {
                                    model_cvitem.Vari_Area_name = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Area_sortorder = null;
                    model_cvitem.Vari_Area_useflg = 1.ToString();
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

            // Dim model_cvitem As New Njc.Model.M_area_Model                  '移行値格納用モデル初期化
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
            // Dim tblname As String = "m_area"
            // Dim fldnamegrp As String = "area_no,area_sortorder,area_name,area_useflg,history"

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
            // Case "エリアNo"
            // .Vari_Area_no = fldvalue.Trim
            // Case "エリア名"
            // .Vari_Area_name = fldvalue.Trim
            // End Select

            // Next

            // '固定値
            // .Vari_Area_sortorder = Nothing
            // .Vari_Area_useflg = 1
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

                string tmp_sql = " SELECT area_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 入金項目マスタ

    // 20160525 紐付項目の挿入処理の追加 -del sta
    // Public Class M_nkin_Repository

    // '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -sta
    // 'Public Class SubConv
    // '    Implements IConv

    // '    ''' <summary>
    // '    ''' 【仮テーブル→V10DB】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <param name="model_nkin"></param>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Cnv_Db(ByVal sqlcnnv10 As SqlConnection, ByVal model_cvitem As Object, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Cnv_Db

    // '        Dim reptbl As String                                '置換TBL名
    // '        Dim repfld As String                                '置換項目名
    // '        Dim repprm As String                                '置換パラメータ
    // '        Dim rowcnt As Integer                               '書込行数
    // '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    // '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    // '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    // '        '書込行数取得
    // '        rowcnt = model_cvitem.ItemCnt

    // '        'プログレスバー初期化
    // '        Dim pgbtotalcnt As Integer = rowcnt
    // '        Call obj_pgb.pgbInitPart(rowcnt)

    // '        'テーブル/フィールド名取得
    // '        reptbl = "m_nkin"
    // '        repfld = "nkin_no,nkin_name,nkin_printname,nkin_ruino,nkin_zkseino," & _
    // '                 "nkin_useflg,keiyaku_nkinno,keiyakuhiki_nkinno,kosin_nkinno,kaiyaku_nkinno," & _
    // '                 "kaiyakuhiki_nkinno,ryosyu_flg,biko_nkin,history"

    // '        With model_cvitem

    // '            For cntii = 0 To rowcnt - 1

    // '                '中断処理
    // '                Application.DoEvents()
    // '                If CancelFlg Then
    // '                    rtn = False
    // '                    Return rtn
    // '                End If

    // '                If .Vari_tmp_skipflg(cntii) = False Then

    // '                    'ハッシュテーブル作成
    // '                    Dim hash_cvitem As New SafeDictionary<string, string>
    // '                    'hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(repfld, model_cvitem, cntii)

    // '                    '書込処理
    // '                    'Call CVDBInsert.CVitem_Insert(sqlcnnv10, reptbl, repfld, hash_cvitem)

    // '                End If

    // '                'プログレスバー更新/進捗率表示
    // '                Dim pgbcnt As Integer = cntii + 1
    // '                Call obj_pgb.pgbsettingPart(pgbcnt)
    // '                Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)

    // '            Next

    // '        End With

    // '        '返却
    // '        Return rtn

    // '    End Function

    // '    ''' <summary>
    // '    ''' 【移行先DB初期化】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <remarks></remarks>
    // '    Public Sub Initialize_Table(sqlcnnv10 As SqlConnection) Implements IConv.Initialize_Table

    // '        '初期化対象外

    // '    End Sub

    // '    ''' <summary>
    // '    ''' 【仮テーブル→変数】
    // '    ''' </summary>
    // '    ''' <param name="sqlcnnv10"></param>
    // '    ''' <param name="syorikomok"></param>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Overridable Function Set_Vari2(ByVal sheetname As String, ByRef model_cvitem As Object, ByVal sqlcnnv10 As SqlConnection) As Boolean Implements IConv.Set_Vari2

    // '        Dim reccnt As Integer                               '抽出レコード件数
    // '        Dim fldname As String                               '該当項目名
    // '        Dim fldvalue As String                              '該当値
    // '        Dim readtbl As New DataTable                        'テーブル格納用
    // '        Dim obj_pgb As New Njc.Common.ProgressBarManager    'プログレスバー設定用
    // '        Dim pgbcnt As Integer = 1                           'プログレスバー更新用カウンター
    // '        Dim obj_com As New Njc.Common.CommonRepository      '共通処理用
    // '        Dim logvalue As String                              'ログ出力用
    // '        Dim tblname As String = "tmp_nkin_mst"              '対象テーブル名
    // '        Dim rtn As Boolean = True                           '戻り値(正常/異常終了判定)

    // '        '移行元情報取得
    // '        reccnt = DBExec.Exec_DataTable(Me.Get_UseQry(), sqlcnnv10, readtbl, rtn)

    // '        '移行元情報取得時にエラーが生じた場合は処理を抜ける
    // '        If Not rtn Then
    // '            Return rtn
    // '        End If

    // '        '初期化
    // '        model_cvitem = New Njc.Model.M_nkin_Model(reccnt - 1)
    // '        Dim tmp_duplicatechk As New Njc.Model.CommonModel(reccnt - 1)

    // '        'プログレスバー初期化
    // '        Dim pgbtotalcnt As Integer = (readtbl.Rows.Count) * (readtbl.Columns.Count)
    // '        Call obj_pgb.pgbInitPart(pgbtotalcnt)

    // '        '==========================
    // '        'データ取得　→　変数格納
    // '        '==========================
    // '        With model_cvitem

    // '            For cntii As Integer = 0 To readtbl.Rows.Count - 1

    // '                'DoEvents
    // '                Application.DoEvents()
    // '                If CancelFlg Then
    // '                    rtn = False
    // '                    Return rtn
    // '                End If

    // '                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

    // '                    '項目名取得
    // '                    fldname = readtbl.Columns(cntjj).ColumnName

    // '                    '登録値取得
    // '                    fldvalue = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

    // '                    'ログ用
    // '                    logvalue = "賃貸革命10入金項目№ = " & Typ.ToStr(readtbl.Rows(cntii).Item(3)).Trim & _
    // '                               " " & _
    // '                               "移行元入金項目名称 = " & Typ.ToStr(readtbl.Rows(cntii).Item(1)).Trim

    // '                    '各項目値→変数格納
    // '                    Select Case fldname
    // '                        Case "賃貸革命10入金項目№"
    // '                            .Vari_Nkin_no(cntii) = DataChk.Chk_DataInt(fldvalue, 1000, 9999)
    // '                        Case "移行元入金項目名称"
    // '                            .Vari_Nkin_name(cntii) = DataChk.Chk_DataString(fldvalue, 100)
    // '                        Case "入金項目区分"
    // '                            .Vari_Nkin_ruino(cntii) = fldvalue     '内部№
    // '                        Case "賃貸革命10入金項目属性№"
    // '                            .Vari_Nkin_zkseino(cntii) = fldvalue   '内部№
    // '                    End Select

    // '                    '固定値設定
    // '                    .Vari_Nkin_printname(cntii) = ""
    // '                    .Vari_Nkin_useflg(cntii) = 1
    // '                    .Vari_Keiyaku_nkinno(cntii) = 0
    // '                    .Vari_Keiyakuhiki_nkinno(cntii) = 0
    // '                    .Vari_Kosin_nkinno(cntii) = 0
    // '                    .Vari_Kaiyaku_nkinno(cntii) = 0
    // '                    .Vari_Kaiyakuhiki_nkinno(cntii) = 0
    // '                    .Vari_Ryosyu_flg(cntii) = 2
    // '                    .Vari_Biko_nkin(cntii) = ""
    // '                    .Vari_History(cntii) = DefHistory
    // '                    .Vari_tmp_skipflg(cntii) = 0

    // '                    'プログレスバー更新/進捗率表示
    // '                    Call obj_pgb.pgbsettingPart(pgbcnt)
    // '                    Call obj_com.ProgressOutPut(pgbcnt, pgbtotalcnt)
    // '                    pgbcnt = pgbcnt + 1

    // '                Next

    // '            Next

    // '        End With

    // '        '返却
    // '        Return rtn

    // '    End Function

    // '    ''' <summary>
    // '    ''' 【抽出クエリ】2015.07.17 sol レビュー後修正_レビュー№312
    // '    ''' </summary>
    // '    ''' <returns></returns>
    // '    ''' <remarks></remarks>
    // '    Public Function Get_UseQry(ByRef sortkeycnt As Integer) As String Implements IConv.Get_UseQry

    // '        Dim tmp_sql As String = ""

    // '        tmp_sql = tmp_sql & " SELECT "
    // '        tmp_sql = tmp_sql & " 	* "
    // '        tmp_sql = tmp_sql & " FROM tmp_nkin_mst "
    // '        tmp_sql = tmp_sql & " WHERE [賃貸革命10入金項目№] NOT IN (SELECT nkin_no FROM m_nkin) "
    // '        tmp_sql = tmp_sql & " AND   [賃貸革命10入金項目№] <> 0 "
    // '        tmp_sql = tmp_sql & " ORDER BY [移行元入金項目№] "

    // '        Return tmp_sql

    // '    End Function

    // '    Public Function Get_BaseKey() As Object Implements IConv.Get_BaseKey

    // '    End Function

    // '                Public Function Set_Vari( _
    // 'ByVal sqlcnnv10 As SqlConnection, _
    // 'ByVal filename As String, ByVal sheetname As String, _
    // 'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    // ') As Boolean Implements IConv.Set_Vari

    // '    End Function

    // '                Public Function Set_Vari( _
    // 'ByVal sqlcnnv10 As SqlConnection, _
    // 'ByVal filename As String, ByVal sheetname As String, _
    // 'ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer _
    // ') As Boolean Implements IConv.Set_Vari

    // '    End Function

    // '    Public Function Write_IntermediateFile(sqlcnnv7 As SqlConnection, filename As String, sheetname As String, model_cvitem As Object) As Boolean Implements IConv.Write_IntermediateFile

    // '    End Function

    // 'End Class
    // '仮マスタを中間ファイルから読み込むように修正するためとりあえずコメントアウトしておく -end

    // End Class
    // 20160525 紐付項目の挿入処理の追加 -del end

    #endregion

    #region 変動費マスタ

    public class M_hendorule_Repository
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

                var model_cvitem = new Model.M_hendorule_Model();             // 移行値格納用モデル初期化
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
                string tblname = "m_hendorule";
                string fldnamegrp = "rule_no,rule_name,tani,tanisjis,ryokin_cnt," + "hendo_rui,hasuusyori_sel,hasuusyosu_ptn,koukei,biko_kihon," + "keijo_rui,zei_kbnhayami,zei_kbntanka,history,useflg";


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
                string fldname_key = Conversions.ToString(headervalue(startrow - 1, keycol));

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

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "ルールNo":
                                {
                                    model_cvitem.Vari_Rule_no = fldvalue.Trim();
                                    break;
                                }
                            case "ルール名":
                                {
                                    model_cvitem.Vari_Rule_name = fldvalue.Trim();
                                    break;
                                }
                            case "その他単位":
                                {
                                    model_cvitem.Vari_Tani = fldvalue.Trim();
                                    break;
                                }
                            case "その他単位SJIS":
                                {
                                    model_cvitem.Vari_Tanisjis = fldvalue.Trim();
                                    break;
                                }
                            case "料金数":
                                {
                                    model_cvitem.Vari_Ryokin_cnt = fldvalue.Trim();
                                    break;
                                }
                            case "分類":
                                {
                                    model_cvitem.Vari_Hendo_rui = fldvalue.Trim();
                                    break;
                                }
                            case "端数処理":
                                {
                                    model_cvitem.Vari_Hasuusyori_sel = fldvalue.Trim();
                                    break;
                                }
                            case "端数処理パターン":
                                {
                                    model_cvitem.Vari_Hasuusyosu_ptn = fldvalue.Trim();
                                    break;
                                }
                            case "口径":
                                {
                                    model_cvitem.Vari_Koukei = fldvalue.Trim();
                                    break;
                                }
                            case "備考":
                                {
                                    model_cvitem.Vari_Biko_kihon = fldvalue.Trim();
                                    break;
                                }
                            case "計上分類":
                                {
                                    model_cvitem.Vari_Keijo_rui = fldvalue.Trim();
                                    break;
                                }
                            case "税区分早見":
                                {
                                    model_cvitem.Vari_Zei_kbnhayami = fldvalue.Trim();
                                    break;
                                }
                            case "税区分単価":
                                {
                                    model_cvitem.Vari_Zei_kbntanka = fldvalue.Trim();
                                    break;
                                }
                        }

                    }

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

                string tmp_sql = " SELECT rule_no FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 変動費一覧マスタ

    public class M_hendorule_itiran_Repository
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

                var model_cvitem = new Model.M_hendorule_itiran_Model();      // 移行値格納用モデル初期化
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
                string tblname_base = "m_hendorule";
                string tblname = "m_hendorule_itiran";
                string fldnamegrp = "rule_no,item_no,siyoryo,ryokin1,ryokin2," + "ryokin3,ryokin4,ryokin5";

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


                // ---------------
                // ヘッダー処理
                // ---------------
                // ヘッダー行取得
                var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

                // キーヘッダー名取得
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub = Conversions.ToString(headervalue(startrow - 1, keycol_sub));

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
                    string tmp_keysub = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub]);
                    string fldvalue_key = tmp_keymain + "-" + tmp_keysub;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + tmp_keymain + "、" + fldname_keysub + " = " + tmp_keysub;

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "変動費ルールNo":
                                {
                                    model_cvitem.Vari_Rule_no = fldvalue.Trim();
                                    break;
                                }
                            case "行No":
                                {
                                    model_cvitem.Vari_Item_no = fldvalue.Trim();
                                    break;
                                }
                            case "使用量":
                                {
                                    model_cvitem.Vari_Siyoryo = fldvalue.Trim();
                                    break;
                                }
                            case "料金１":
                                {
                                    model_cvitem.Vari_Ryokin1 = fldvalue.Trim();
                                    break;
                                }
                            case "料金２":
                                {
                                    model_cvitem.Vari_Ryokin2 = fldvalue.Trim();
                                    break;
                                }
                            case "料金３":
                                {
                                    model_cvitem.Vari_Ryokin3 = fldvalue.Trim();
                                    break;
                                }
                            case "料金４":
                                {
                                    model_cvitem.Vari_Ryokin4 = fldvalue.Trim();
                                    break;
                                }
                            case "料金５":
                                {
                                    model_cvitem.Vari_Ryokin5 = fldvalue.Trim();
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

                string tmp_sql = " SELECT rule_no FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,buskotu_no) + '-' + CONVERT(varchar,sortorder) FROM  " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 備考タイトルマスタ

    public class M_memo_Repository
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

                var model_cvitem = new Model.M_memo_Model();                  // 移行値格納用モデル初期化
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
                string tblname = "m_memo";
                string fldnamegrp = "memo_kbn,memo_no,memo_sortorder,memo_name,memo_useflg," + "memo_usehojoitems,history";

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
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub = Conversions.ToString(headervalue(startrow - 1, keycol_sub));

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

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "備考区分":
                                {
                                    model_cvitem.Vari_Memo_kbn = fldvalue;
                                    break;
                                }
                            case "備考No":
                                {
                                    model_cvitem.Vari_Memo_no = fldvalue;
                                    break;
                                }
                            case "備考タイトル":
                                {
                                    model_cvitem.Vari_Memo_name = fldvalue;
                                    break;
                                }
                            case "使用有無":
                                {
                                    model_cvitem.Vari_Memo_useflg = fldvalue;
                                    break;
                                }
                            case "補助アイテム使用有無":
                                {
                                    model_cvitem.Vari_Memo_usehojoitems = fldvalue;
                                    break;
                                }
                        }

                    }

                    // 固定値
                    model_cvitem.Vari_Memo_sortorder = 0.ToString();
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

                string tmp_sql = " SELECT CONVERT(varchar,memo_kbn) + '-' + CONVERT(varchar,memo_no) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 備考入力補助リストマスタ

    public class M_memo_lst_Repository
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

                var model_cvitem = new Model.M_memo_lst_Model();              // 移行値格納用モデル初期化
                int keycol_main = 1;                                  // メインキー列
                int keycol_sub1 = 2;                                   // サブ1キー列
                int keycol_sub2 = 3;                                   // サブ2キー列

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
                string tblname_base = "m_memo";
                string tblname = "m_memo_lst";
                string fldnamegrp = "memo_kbn,memo_no,memo_lstno,memo_lstname,history";

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
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub1 = Conversions.ToString(headervalue(startrow - 1, keycol_sub1));
                string fldname_keysub2 = Conversions.ToString(headervalue(startrow - 1, keycol_sub2));

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
                    string fldvalue_keysub1 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub1]);
                    string fldvalue_keysub2 = EtcMethod.Get_ShapeStr(datarowvalue[startrow - 1, keycol_sub2]);
                    string fldvalue_key = fldvalue_keymain + "-" + fldvalue_keysub1 + "-" + fldvalue_keysub2;

                    // ログ出力用
                    string str_logkey = fldname_keymain + " = " + fldvalue_keymain + "、" + fldname_keysub1 + " = " + fldvalue_keysub1 + "、" + fldname_keysub2 + " = " + fldvalue_keysub2;

                    // 移行値取得
                    for (int cntjj = 1, loopTo1 = columncnt; cntjj <= loopTo1; cntjj++)
                    {

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "備考区分":
                                {
                                    model_cvitem.Vari_Memo_kbn = fldvalue;
                                    break;
                                }
                            case "備考No":
                                {
                                    model_cvitem.Vari_Memo_no = fldvalue;
                                    break;
                                }
                            case "備考入力補助リストNo":
                                {
                                    model_cvitem.Vari_Memo_lstno = fldvalue;
                                    break;
                                }
                            case "備考入力補助リスト内容":
                                {
                                    model_cvitem.Vari_Memo_lstname = fldvalue;
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

            /// <summary>
            /// 親マスタ取得→リスト格納
            /// </summary>
            /// <param name="sqlcnnv10"></param>
            /// <param name="tblname"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

                string tmp_sql = " SELECT CONVERT(varchar,memo_kbn) + '-' + CONVERT(varchar,memo_no) FROM " + tblname;
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

                string tmp_sql = " SELECT CONVERT(varchar,memo_kbn) + '-' + CONVERT(varchar,memo_no) + '-' + CONVERT(varchar,memo_lstno) FROM  " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

    #region 画像タイトルマスタ

    public class M_gazo_title_Repository
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

                var model_cvitem = new Model.M_gazo_title_Model();            // 移行値格納用モデル初期化
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
                string tblname = "m_gazo_title";
                string fldnamegrp = "gazo_kbn,gazo_no,gazo_name,history";

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
                string fldname_keymain = Conversions.ToString(headervalue(startrow - 1, keycol_main));
                string fldname_keysub = Conversions.ToString(headervalue(startrow - 1, keycol_sub));

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

                        string fldname = headervalue(startrow - 1, cntjj).ToString().Trim();
                        string fldvalue = "";
                        if (datarowvalue[startrow - 1, cntjj] is not null)
                        {
                            fldvalue = datarowvalue[startrow - 1, cntjj].ToString().Trim();
                        }

                        switch (fldname ?? "")
                        {
                            case "画像区分":
                                {
                                    model_cvitem.Vari_Gazo_kbn = fldvalue;
                                    break;
                                }
                            case "画像No":
                                {
                                    model_cvitem.Vari_Gazo_no = fldvalue;
                                    break;
                                }
                            case "画像名":
                                {
                                    model_cvitem.Vari_Gazo_name = fldvalue;
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

                string tmp_sql = " SELECT CONVERT(varchar,memo_kbn) + '-' + CONVERT(varchar,memo_no) FROM " + tblname;
                bool flg = true;

                flg = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);

            }

        }

    }

    #endregion

}