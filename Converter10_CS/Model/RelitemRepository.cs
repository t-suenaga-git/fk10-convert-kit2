using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Repository
{

    #region 物件分類マスタ

    public class M_bk_rui_Repository
    {

        /// <summary>
        /// 物件分類マスタ移行処理
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_RelItem(SqlConnection sqlcnnv10, string sheetname)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            int tmp_cvcnt;                                        // 移行件数格納
            int tmp_condcnt;                                      // 調整件数格納

            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
            var obj_com = new CommonRepository();                  // 共通処理用

            var normalflg = default(bool);                                        // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値
            var model_cvitem = new Model.M_bk_rui_Model();                  // 移行値格納用モデル初期化

            // 移行対象のテーブル名、フィールド名を取得
            string tblname = "m_bk_rui";
            string fldnamegrp = "bk_ruino,bk_ruiname,bk_ruibiko,bk_ruiuseflg,sincyoku_keiyakukbn," + "sincyoku_kosinkbn,sincyoku_kaiyakukbn,history,homemate_ruikbn,bk_endofmonthflg," + "ikkatukariage_siwakekbn";


            // 既存データ取得
            var list_existdata = new List<string>();
            string tmp_sql = " SELECT bk_ruino FROM " + tblname;
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);
            if (rtn == false)
            {
                return rtn;
            }

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, sheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (!rtn)
            {
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // ----------------------------
            // 紐付データ取得
            // ----------------------------

            // ヘッダー行取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;


            // データ部処理
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
                        case "賃貸革命10物件分類No":
                            {
                                model_cvitem.Vari_Bk_ruino = fldvalue;
                                break;
                            }
                        case "賃貸革命10物件分類名称":
                            {
                                model_cvitem.Vari_Bk_ruiname = fldvalue;
                                break;
                            }
                    }

                }

                // 固定値
                model_cvitem.Vari_Bk_ruibiko = "";
                model_cvitem.Vari_Bk_ruiuseflg = 1.ToString();
                model_cvitem.Vari_Sincyoku_keiyakukbn = 0.ToString();
                model_cvitem.Vari_Sincyoku_kosinkbn = 0.ToString();
                model_cvitem.Vari_Sincyoku_kaiyakukbn = 0.ToString();
                model_cvitem.Vari_Homemate_ruikbn = "";
                model_cvitem.Vari_Bk_endofmonthflg = "";
                model_cvitem.Vari_Ikkatukariage_siwakekbn = "";
                model_cvitem.Vari_History = CommonModule.DefHistory;

                // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                // 挿入処理
                if (!string.IsNullOrEmpty(model_cvitem.Vari_Bk_ruino) && list_existdata.Contains(model_cvitem.Vari_Bk_ruino) == false)
                {
                    CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);
                }


            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            return rtn;

        }

    }

    #endregion

    #region 部屋分類マスタ

    public class M_hy_rui_Repository
    {

        /// <summary>
        /// 部屋分類マスタ移行処理
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_RelItem(SqlConnection sqlcnnv10, string sheetname)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            int tmp_cvcnt;                                        // 移行件数格納
            int tmp_condcnt;                                      // 調整件数格納

            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
            var obj_com = new CommonRepository();                  // 共通処理用

            var normalflg = default(bool);                                        // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値
            var model_cvitem = new Model.M_hy_rui_Model();                  // 移行値格納用モデル初期化

            // 移行対象のテーブル名、フィールド名を取得
            string tblname = "m_hy_rui";
            // hy_ruisortorderは自動で挿入されるため対象外
            string fldnamegrp = "hy_ruino,hy_ruiname,hy_ruisyubetu,hy_ruiuseflg,hy_carportflg," + "history";

            // 既存データ取得
            var list_existdata = new List<string>();
            string tmp_sql = " SELECT hy_ruino FROM " + tblname;
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);
            if (rtn == false)
            {
                return rtn;
            }

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, sheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (!rtn)
            {
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // ----------------------------
            // 紐付データ取得
            // ----------------------------

            // ヘッダー行取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;


            // データ部処理
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
                        case "賃貸革命10部屋分類No":
                            {
                                model_cvitem.Vari_Hy_ruino = fldvalue;
                                break;
                            }
                        // .Vari_Hy_ruisortorder = fldvalue    'hy_ruisortorderは自動で挿入されるため対象外
                        case "賃貸革命10部屋分類名称":
                            {
                                model_cvitem.Vari_Hy_ruiname = fldvalue;
                                break;
                            }
                    }

                }

                // 固定値
                model_cvitem.Vari_Hy_ruisyubetu = 1.ToString();         // 紐付けに追加する必要あり
                model_cvitem.Vari_Hy_ruiuseflg = 1.ToString();
                model_cvitem.Vari_Hy_carportflg = 1.ToString();         // 紐付けに追加する必要あり
                model_cvitem.Vari_History = CommonModule.DefHistory;

                // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                // 挿入処理
                if (!string.IsNullOrEmpty(model_cvitem.Vari_Hy_ruino) && list_existdata.Contains(model_cvitem.Vari_Hy_ruino) == false)
                {
                    CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);
                }


            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            return rtn;

        }

    }

    #endregion

    #region 入金区分マスタ

    public class M_nkbn_Repository
    {

        /// <summary>
        /// 入金区分マスタ移行処理
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_RelItem(SqlConnection sqlcnnv10, string sheetname)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            int tmp_cvcnt;                                        // 移行件数格納
            int tmp_condcnt;                                      // 調整件数格納

            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
            var obj_com = new CommonRepository();                  // 共通処理用

            var normalflg = default(bool);                                        // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値
            var model_cvitem = new Model.M_nkbn_Model();                  // 移行値格納用モデル初期化

            // 移行対象のテーブル名、フィールド名を取得
            string tblname = "m_nkbn";
            string fldnamegrp = "nkbn_no,nkbn_order,nkbn_name,nkbn_shortname,nkbn_zokusei," + "history";

            // 既存データ取得
            var list_existdata = new List<string>();
            string tmp_sql = " SELECT nkbn_no FROM " + tblname;
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);
            if (rtn == false)
            {
                return rtn;
            }

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, sheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (!rtn)
            {
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // ----------------------------
            // 紐付データ取得
            // ----------------------------

            // ヘッダー行取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;


            // データ部処理
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
                        case "賃貸革命10入金区分No":
                            {
                                model_cvitem.Vari_Nkbn_no = fldvalue;
                                model_cvitem.Vari_Nkbn_order = fldvalue;
                                break;
                            }
                        case "賃貸革命10入金区分名称":
                            {
                                model_cvitem.Vari_Nkbn_name = fldvalue;
                                break;
                            }
                        case "賃貸革命10入金区分属性No":
                            {
                                model_cvitem.Vari_Nkbn_zokusei = fldvalue;
                                break;
                            }
                    }

                }

                // 固定値
                model_cvitem.Vari_Nkbn_shortname = "";
                model_cvitem.Vari_History = CommonModule.DefHistory;

                // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                // 挿入処理
                if (!string.IsNullOrEmpty(model_cvitem.Vari_Nkbn_no) && list_existdata.Contains(model_cvitem.Vari_Nkbn_no) == false)
                {
                    CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);
                }


            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            return rtn;

        }

    }

    #endregion

    #region 契約分類マスタ

    public class M_ky_rui_Repository
    {

        /// <summary>
        /// 契約分類マスタ移行処理
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_RelItem(SqlConnection sqlcnnv10, string sheetname)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            int tmp_cvcnt;                                        // 移行件数格納
            int tmp_condcnt;                                      // 調整件数格納

            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
            var obj_com = new CommonRepository();                  // 共通処理用

            var normalflg = default(bool);                                        // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値
            var model_cvitem = new Model.M_ky_rui_Model();                // 移行値格納用モデル初期化

            // 移行対象のテーブル名、フィールド名を取得
            string tblname = "m_ky_rui";
            string fldnamegrp = "ky_ruino,ky_ruiname,ky_ruibiko,ky_ruitutimm,ky_ruicolor," + "teisyaku_flg,history,useflg";

            // 既存データ取得
            var list_existdata = new List<string>();
            string tmp_sql = " SELECT ky_ruino FROM " + tblname;
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);
            if (rtn == false)
            {
                return rtn;
            }

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, sheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (!rtn)
            {
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // ----------------------------
            // 紐付データ取得
            // ----------------------------

            // ヘッダー行取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;


            // データ部処理
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

                // 作業用変数
                string tmp_syakuyakbn = "";

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
                        case "賃貸革命10契約分類No":
                            {
                                model_cvitem.Vari_Ky_ruino = fldvalue;
                                break;
                            }
                        case "賃貸革命10契約分類名称":
                            {
                                model_cvitem.Vari_Ky_ruiname = fldvalue;
                                break;
                            }
                        case "定期借家として扱う":
                            {
                                tmp_syakuyakbn = Conversions.ToString(Interaction.IIf(fldvalue == "True", 1, 2));
                                break;
                            }
                    }

                }

                // 定期借家区分の判定
                if (string.IsNullOrEmpty(model_cvitem.Vari_Ky_ruino) | string.IsNullOrEmpty(model_cvitem.Vari_Ky_ruino))
                {
                    model_cvitem.Vari_Teisyaku_flg = 2.ToString();
                }
                else
                {
                    model_cvitem.Vari_Teisyaku_flg = tmp_syakuyakbn;
                }

                // 固定値
                model_cvitem.Vari_Ky_ruibiko = "";
                model_cvitem.Vari_Ky_ruitutimm = "";
                model_cvitem.Vari_Ky_ruicolor = "-16777216";
                model_cvitem.Vari_Useflg = 1.ToString();
                model_cvitem.Vari_History = CommonModule.DefHistory;

                // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                // 挿入処理
                if (!string.IsNullOrEmpty(model_cvitem.Vari_Ky_ruino) && list_existdata.Contains(model_cvitem.Vari_Ky_ruino) == false)
                {
                    CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);
                }


            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            return rtn;

        }

    }

    #endregion

    #region 入金項目マスタ

    public class M_nkin_Repository
    {

        /// <summary>
        /// 入金項目マスタ移行処理
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_RelItem(SqlConnection sqlcnnv10, string sheetname)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            int tmp_cvcnt;                                        // 移行件数格納
            int tmp_condcnt;                                      // 調整件数格納

            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
            var obj_com = new CommonRepository();                  // 共通処理用

            var normalflg = default(bool);                                        // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値
            var model_cvitem = new Model.M_nkin_Model();                  // 移行値格納用モデル初期化

            // 移行対象のテーブル名、フィールド名を取得
            string tblname = "m_nkin";
            string fldnamegrp = "nkin_no,nkin_name,nkin_printname,nkin_ruino,nkin_zkseino," + "nkin_useflg,keiyaku_nkinno,keiyakuhiki_nkinno,kosin_nkinno,kaiyaku_nkinno," + "kaiyakuhiki_nkinno,ryosyu_flg,biko_nkin,history";


            // 既存データ取得
            var list_existdata = new List<string>();
            string tmp_sql = " SELECT nkin_no FROM " + tblname;
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, ref sqlcnnv10, ref list_existdata);
            if (rtn == false)
            {
                return rtn;
            }

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, sheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (!rtn)
            {
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // ----------------------------
            // 紐付データ取得
            // ----------------------------

            // ヘッダー行取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;


            // データ部処理
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
                        case "賃貸革命10入金項目No":
                            {
                                model_cvitem.Vari_Nkin_no = fldvalue;
                                break;
                            }
                        case "賃貸革命10入金項目名称":
                            {
                                model_cvitem.Vari_Nkin_name = fldvalue;
                                break;
                            }
                        case "移行元入金項目区分":
                            {
                                model_cvitem.Vari_Nkin_ruino = Strings.Left(fldvalue, 1);
                                break;
                            }
                        case "賃貸革命10入金項目属性No":
                            {
                                model_cvitem.Vari_Nkin_zkseino = fldvalue;
                                break;
                            }
                    }

                }

                // 固定値
                model_cvitem.Vari_Nkin_printname = "";
                model_cvitem.Vari_Nkin_useflg = 1.ToString();
                model_cvitem.Vari_Keiyaku_nkinno = 0.ToString();
                model_cvitem.Vari_Keiyakuhiki_nkinno = 0.ToString();
                model_cvitem.Vari_Kosin_nkinno = "";
                model_cvitem.Vari_Kaiyaku_nkinno = "";
                model_cvitem.Vari_Kaiyakuhiki_nkinno = 0.ToString();
                model_cvitem.Vari_Ryosyu_flg = 2.ToString();
                model_cvitem.Vari_Biko_nkin = "";
                model_cvitem.Vari_History = CommonModule.DefHistory;

                // フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                var hash_cvitem = new SafeDictionary<string, string>();
                hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

                // 挿入処理
                if (!string.IsNullOrEmpty(model_cvitem.Vari_Nkin_no) && list_existdata.Contains(model_cvitem.Vari_Nkin_no) == false)
                {
                    CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, ref normalflg);
                }


            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            return rtn;

        }

    }

    #endregion

    #region 設備マスタ

    public class M_setubi_rel_Repository
    {

        /// <summary>
        /// 設備マスタ移行処理
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Set_RelItem(SqlConnection sqlcnnv10, string sheetname)
        {

            Microsoft.Office.Interop.Excel.Application appli = null;                        // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Workbook wbook = null;                           // Excelオブジェクト
            Microsoft.Office.Interop.Excel.Worksheet wsheet = null;                         // Excelオブジェクト
            var startrow = default(int);                                         // 書込開始行
            var columncnt = default(int);                                        // 列数
            var maxrowcnt = default(int);                                        // 既存データの行数
            var rowcnt = default(int);                                           // 書込行数
            int tmp_cvcnt;                                        // 移行件数格納
            int tmp_condcnt;                                      // 調整件数格納

            var excelfile = new ExcelFileManager();                // Excelファイル操作用
            var obj_pgb = new ProgressBarManager();                // プログレスバー設定用
            var obj_com = new CommonRepository();                  // 共通処理用

            bool normalflg;                                        // INSERT正常終了フラグ
            bool rtn = true;                                       // 戻り値

            // Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(ref appli, ref wbook, ref wsheet, ref startrow, ref columncnt, ref maxrowcnt, ref rowcnt, CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME, sheetname);

            // Excelファイル設定時にエラーが生じた際は処理を抜ける
            if (!rtn)
            {
                excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);
                return rtn;
            }

            // 革命10へ設備グループマスタを移行する用意
            bool grpinsertflg = true;
            string tmp_grpguid = "";
            string tmp_grpno = "";

            // ----------------------------
            // 紐付データ取得
            // ----------------------------

            // ヘッダー行取得
            var headervalue = ((Microsoft.Office.Interop.Excel.Range)wsheet.Range[wsheet.Cells[startrow - 1, 1], wsheet.Cells[startrow - 1, columncnt]]).Value;

            // データ部処理
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

                // 作業用変数
                string tmp_setubiname = "";
                string tmp_komkname = "";
                string tmp_grpname10 = "";
                string tmp_komokno10 = "";

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
                        case "移行元設備グループ名称":
                            {
                                tmp_setubiname = fldvalue;
                                break;
                            }
                        case "移行元設備名称":
                            {
                                tmp_komkname = fldvalue;
                                break;
                            }
                        case "賃貸革命10設備グループNo":
                            {
                                tmp_grpname10 = fldvalue;
                                break;
                            }
                        case "賃貸革命10設備グループ名称":
                            {
                                break;
                            }

                        case "賃貸革命10設備No":
                            {
                                break;
                            }

                        case "賃貸革命10設備名称":
                            {
                                break;
                            }

                        case "賃貸革命10項目No":
                            {
                                tmp_komokno10 = fldvalue;
                                break;
                            }
                        case "賃貸革命10項目名称":
                            {
                                break;
                            }

                    }

                }

                // 新規移行対象確認
                // ※グループNo = 999はエレベーターなので項目は存在しない
                if (tmp_grpname10 != "999" & string.IsNullOrEmpty(tmp_komokno10))
                {

                    // 新規設備作成開始
                    // 設備グループ作成
                    if (grpinsertflg)
                    {
                        Set_SetubiGrp(sqlcnnv10, ref tmp_grpguid);
                        grpinsertflg = false;
                    }

                    // 設備作成
                    string setubiguid = "";
                    Set_Setubi(sqlcnnv10, tmp_grpguid, tmp_setubiname, ref setubiguid);

                    // 設備項目作成
                    Set_Setubi_Komk(sqlcnnv10, setubiguid, tmp_komkname);

                }

            }

            // Excelファイル終了設定
            excelfile.Set_ExcelFile_ReadClose(ref appli, ref wbook, ref wsheet);

            return rtn;

        }

        /// <summary>
        /// 設備グループマスタの作成(最初に1回のみ実行する)
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="grpguid"></param>
        /// <param name="grpno"></param>
        /// <remarks></remarks>
        public void Set_SetubiGrp(SqlConnection sqlcnnv10, ref string grpguid)
        {

            bool normalflg = true;
            var model_cvitem = new Model.M_setubi_rel_grp_Model();
            string tblname = "m_setubi_grp";
            string fldnamegrp = "setubi_grpguid,setubi_grpsortorder,setubi_grpname,setubi_grpuseflg,setubi_grpsyskbn," + "history";

            // 設備グループマスタ有無確認
            string tmp_sqlexist = " SELECT CONVERT(VARCHAR(MAX),setubi_grpguid) FROM m_setubi_grp WHERE setubi_grpname = 'ユーザー作成設備' ";
            string exist_setubigrpguid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sqlexist, ref sqlcnnv10));

            // 存在する場合は戻り値用に設備guidを格納して処理を抜ける
            if (!string.IsNullOrEmpty(exist_setubigrpguid))
            {
                grpguid = exist_setubigrpguid;
                return;
            }

            // 設備グループマスタ作成準備
            string new_grpguid = Guid.NewGuid().ToString();
            string tmp_sqlgetgrpno = " SELECT CONVERT(VARCHAR,MAX(setubi_grpsortorder) + 1)  FROM m_setubi_grp ";
            string tmp_grpno = Conversions.ToString(DBExec.Exec_Scalar(tmp_sqlgetgrpno, ref sqlcnnv10));

            // モデルへ値を格納
            string defint = "1";
            model_cvitem.Vari_Setubi_grpguid = new_grpguid;
            model_cvitem.Vari_Setubi_grpsortorder = tmp_grpno;
            model_cvitem.Vari_Setubi_grpname = "ユーザー作成設備";
            model_cvitem.Vari_Setubi_grpuseflg = defint;
            model_cvitem.Vari_Setubi_grpsyskbn = defint;
            model_cvitem.Vari_History = CommonModule.DefHistory;

            // ハッシュテーブルへ格納
            var hash_setubigrpitem = new SafeDictionary<string, string>();
            hash_setubigrpitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

            // 挿入処理実行
            CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_setubigrpitem, ref normalflg);

            // 戻り値用に値を格納
            grpguid = new_grpguid;

        }

        /// <summary>
        /// 設備マスタの作成(同一名称のものは同一と見なし、統合する)
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="grpguid"></param>
        /// <param name="grpno"></param>
        /// <remarks></remarks>
        public void Set_Setubi(SqlConnection sqlcnnv10, string grpguid, string setubiname, ref string setubiguid)
        {

            bool normalflg = true;
            var model_cvitem = new Model.M_setubi_rel_Model();
            string tblname = "m_setubi";
            string fldnamegrp = "setubi_guid,setubi_grpguid,setubi_sortorder,setubi_code,setubi_name," + "setubi_useflg,setubi_hyuseflg,setubi_disp1name,setubi_disp2name,setubi_disp3name," + "history,jyuyojiko_flg";



            // 設備マスタ有無確認
            string tmp_sqlexist = "";
            tmp_sqlexist = tmp_sqlexist + " SELECT CONVERT(VARCHAR(MAX),setubi_guid) FROM m_setubi ";
            tmp_sqlexist = tmp_sqlexist + " WHERE setubi_grpguid = '" + grpguid + "' ";
            tmp_sqlexist = tmp_sqlexist + " AND setubi_name = '" + setubiname + "' ";
            string exist_setubiguid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sqlexist, ref sqlcnnv10));

            // 存在する場合は戻り値用に設備guidを格納して処理を抜ける
            if (!string.IsNullOrEmpty(exist_setubiguid))
            {
                setubiguid = exist_setubiguid;
                return;
            }

            // 設備マスタの新規作成
            string new_setubiguid = Guid.NewGuid().ToString();
            string tmp_sqlgetsetubino = "";
            tmp_sqlgetsetubino = tmp_sqlgetsetubino + " SELECT CONVERT(VARCHAR,MAX(setubi_sortorder) + 1) FROM m_setubi ";
            tmp_sqlgetsetubino = tmp_sqlgetsetubino + " WHERE setubi_grpguid = '" + grpguid + "' ";
            string tmp_setubino = Conversions.ToString(DBExec.Exec_Scalar(tmp_sqlgetsetubino, ref sqlcnnv10));
            if (tmp_setubino == "0")
            {
                tmp_setubino = 1.ToString();
            }

            // 挿入クエリ作成
            string defint1 = "1";
            string defint2 = "0";
            model_cvitem.Vari_Setubi_guid = new_setubiguid;
            model_cvitem.Vari_Setubi_grpguid = grpguid;
            model_cvitem.Vari_Setubi_sortorder = tmp_setubino;
            model_cvitem.Vari_Setubi_code = "";
            model_cvitem.Vari_Setubi_name = setubiname;
            model_cvitem.Vari_Setubi_useflg = defint1;
            model_cvitem.Vari_Setubi_hyuseflg = defint1;
            model_cvitem.Vari_Setubi_disp1name = "";
            model_cvitem.Vari_Setubi_disp2name = "";
            model_cvitem.Vari_Setubi_disp3name = "";
            model_cvitem.Vari_History = CommonModule.DefHistory;
            model_cvitem.Vari_Jyuyojiko_flg = defint2;

            // ハッシュテーブルへ格納
            var hash_setubiitem = new SafeDictionary<string, string>();
            hash_setubiitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

            // 挿入処理実行
            CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_setubiitem, ref normalflg);

            // 戻り値用に値を格納
            setubiguid = new_setubiguid;

        }

        /// <summary>
        /// 設備項目マスタの作成(同一名称のものは同一と見なす)
        /// </summary>
        /// <param name="sqlcnnv10"></param>
        /// <param name="grpguid"></param>
        /// <param name="grpno"></param>
        /// <remarks></remarks>
        public void Set_Setubi_Komk(SqlConnection sqlcnnv10, string setubiguid, string komkname)
        {

            bool normalflg = true;
            var model_cvitem = new Model.M_setubi_rel_komk_Model();
            string tblname = "m_setubi_lst";
            string fldnamegrp = "komok_guid,setubi_guid,komok_sortorder,komok_name,komok_code1," + "komok_code2,komok_code3,default_flg,disp1name,disp2name," + "disp3name,disp1iconguid,disp2iconguid,disp3iconguid,history";


            // 設備項目マスタ有無確認(紐付時に重複は除去しているが念の為)
            string tmp_sqlexist = "";
            tmp_sqlexist = tmp_sqlexist + " SELECT CONVERT(VARCHAR(MAX),komok_guid) FROM m_setubi_lst ";
            tmp_sqlexist = tmp_sqlexist + " WHERE setubi_guid = '" + setubiguid + "' ";
            tmp_sqlexist = tmp_sqlexist + " AND komok_name = '" + komkname + "' ";
            string exist_komkguid = Conversions.ToString(DBExec.Exec_Scalar(tmp_sqlexist, ref sqlcnnv10));

            // 存在する場合は処理を抜ける
            if (!string.IsNullOrEmpty(exist_komkguid))
            {
                return;
            }

            // 設備項目マスタの新規作成
            string new_komkguid = Guid.NewGuid().ToString();
            string tmp_sqlkomkno = "";
            tmp_sqlkomkno = tmp_sqlkomkno + " SELECT CONVERT(VARCHAR,MAX(komok_sortorder) + 1) FROM m_setubi_lst ";
            tmp_sqlkomkno = tmp_sqlkomkno + " WHERE setubi_guid = '" + setubiguid + "' ";
            string tmp_komkno = Conversions.ToString(DBExec.Exec_Scalar(tmp_sqlkomkno, ref sqlcnnv10));
            if (tmp_komkno == "0")
            {
                tmp_komkno = 1.ToString();
            }

            // 挿入クエリ作成
            string defint = "0";
            string defstr = "00000000-0000-0000-0000-000000000000";
            model_cvitem.Vari_Komok_guid = new_komkguid;
            model_cvitem.Vari_Setubi_guid = setubiguid;
            model_cvitem.Vari_Komok_sortorder = tmp_komkno;
            model_cvitem.Vari_Komok_name = komkname;
            model_cvitem.Vari_Komok_code1 = "";
            model_cvitem.Vari_Komok_code2 = "";
            model_cvitem.Vari_Komok_code3 = "";
            model_cvitem.Vari_Default_flg = defint;
            model_cvitem.Vari_Disp1name = "";
            model_cvitem.Vari_Disp2name = "";
            model_cvitem.Vari_Disp3name = "";
            model_cvitem.Vari_Disp1iconguid = defstr;
            model_cvitem.Vari_Disp2iconguid = defstr;
            model_cvitem.Vari_Disp3iconguid = defstr;
            model_cvitem.Vari_History = CommonModule.DefHistory;

            // ハッシュテーブルへ格納
            var hash_setubikomkitem = new SafeDictionary<string, string>();
            hash_setubikomkitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem);

            // 挿入処理実行
            CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_setubikomkitem, ref normalflg);

        }

    }

    #endregion

}