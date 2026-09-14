using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using Converter10.Njc.Common;

namespace Converter10.Njc.Repository
{

    #region 過剰金情報

    public class Sq_kajo_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【抽出クエリ】
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";

                sortstr = "[物件No],[部屋No],[契約No],[契約レコードNo]";

                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	/******************************************************************************** ";
                tmp_sql = tmp_sql + " 	過剰金データの抽出 ";
                tmp_sql = tmp_sql + " 	  過剰金の場合、sq_meisaiの履歴管理No(ko_no)が存在しないためky_kosinkaiから取得する ";
                tmp_sql = tmp_sql + " 	  【取得方法】 ";
                tmp_sql = tmp_sql + " 		契約開始終了日の範囲内に預かり金入金日 (sq_meisai.gt_ym)が存在するか ";
                tmp_sql = tmp_sql + " 		または請求先Noが一致するかで判定し、抽出/結合してko_noを取得 ";
                tmp_sql = tmp_sql + " 	********************************************************************************/ ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 KAJYO.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,KAJYO.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,KAJYO.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,KYK.ko_no AS [契約レコードNo] ";
                tmp_sql = tmp_sql + " 		,gt_ym AS [預り金処理日] ";
                tmp_sql = tmp_sql + " 		,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 		,6 AS [入金項目区分]		/*過剰入金の場合、入金項目区分にデータが存在しないため「その他」で抽出する*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [表示用入金項目名] ";
                tmp_sql = tmp_sql + " 		,nkbn_name AS [入金区分] ";
                tmp_sql = tmp_sql + " 		,sq_gak AS [預り額] ";
                tmp_sql = tmp_sql + " 		,100 AS [利害関係者区分] ";
                tmp_sql = tmp_sql + " 		,KAJYO.kys_no AS [利害関係者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [月区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約入金情報の入金項目No] ";
                tmp_sql = tmp_sql + " 		,'' AS [入金項目レコードNo] ";
                tmp_sql = tmp_sql + " 		,'' AS [預り予定フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [仕訳フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [通常仕訳額] ";
                tmp_sql = tmp_sql + " 		,KAJYO.fkom_no AS [家賃入金口座No]	/*出納データ移行用*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*過剰金データの抽出*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM sq_meisai AS SQ WHERE SQ.nkin_no = 6020 ";
                tmp_sql = tmp_sql + " 	) AS KAJYO ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 	ON  KAJYO.bk_no = KYK.bk_no ";
                tmp_sql = tmp_sql + " 	AND KAJYO.hy_no = KYK.hy_no ";
                tmp_sql = tmp_sql + " 	AND KAJYO.ky_no = KYK.ky_no ";
                tmp_sql = tmp_sql + " 	AND (KAJYO.gt_ym BETWEEN KYK.ky_start_ymd AND KYK.ky_end_ymd AND KAJYO.kys_no = KYK.kys_no) ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MNKIN ON KAJYO.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkbn AS MNKBN ON KAJYO.nkbn_jitu = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " ) AS VW ";

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

            }

            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {
                return default;

            }

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

            }

        }

    }

    #endregion

    #region 出納情報 (過剰金データを移行)

    // 過剰金データ抽出時に出納データを抽出→過剰金移行時に出納データも移行するためここでは抽出しない

    // Public Class Sq_kajo_suito_Repository

    // Public Class SubConv
    // Implements IConv

    // ''' <summary>
    // ''' 【抽出クエリ】
    // ''' </summary>
    // ''' <returns></returns>
    // ''' <remarks></remarks>
    // Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

    // Dim tmp_sql As String = ""

    // sortstr = "[物件No],[部屋No],[契約No],[契約レコードNo]"

    // tmp_sql = tmp_sql & " SELECT "
    // tmp_sql = tmp_sql & " 	 SQ.bk_no AS [物件No] "
    // tmp_sql = tmp_sql & " 	,SQ.hy_no AS [部屋No] "
    // tmp_sql = tmp_sql & " 	,SQ.ky_no AS [契約No] "
    // tmp_sql = tmp_sql & " 	,KYK.ko_no AS [契約レコードNo] "
    // tmp_sql = tmp_sql & " 	,gt_ym AS [出納日] "
    // tmp_sql = tmp_sql & " 	,1 AS [出納区分(1:入金/2:出金)] "
    // tmp_sql = tmp_sql & " 	,100 AS [利害関係者区分(100:契約者/200:家主/3:不動産業者...)] "
    // tmp_sql = tmp_sql & " 	,SQ.kys_no AS [利害関係者No] "
    // tmp_sql = tmp_sql & " 	,nkbn_name AS [入金区分] "
    // tmp_sql = tmp_sql & " 	,sq_gak AS [金額] "
    // tmp_sql = tmp_sql & " 	,'' AS [担当者No] "
    // tmp_sql = tmp_sql & " 	,'' AS [自社支店No] "
    // tmp_sql = tmp_sql & " 	,SQ.fkom_no AS [家賃振込口座No] "
    // tmp_sql = tmp_sql & " 	,fkae_sqno AS [振替履歴No] "
    // tmp_sql = tmp_sql & " 	,'' AS [入出金No] "
    // tmp_sql = tmp_sql & " 	,'' AS [分割No] "
    // tmp_sql = tmp_sql & " 	,'' AS [備考] "
    // tmp_sql = tmp_sql & " 	,'' AS [コンビニ収納No] "
    // tmp_sql = tmp_sql & " 	,'' AS [顧客番号] "
    // tmp_sql = tmp_sql & " 	,fkae_no AS [口座振替No] "
    // tmp_sql = tmp_sql & " FROM sq_meisai AS SQ "
    // tmp_sql = tmp_sql & " LEFT JOIN ky_kosinkai AS KYK "
    // tmp_sql = tmp_sql & " ON  SQ.bk_no = KYK.bk_no "
    // tmp_sql = tmp_sql & " AND SQ.hy_no = KYK.hy_no "
    // tmp_sql = tmp_sql & " AND SQ.ky_no = KYK.ky_no "
    // tmp_sql = tmp_sql & " AND (SQ.gt_ym BETWEEN KYK.ky_start_ymd AND KYK.ky_end_ymd OR SQ.kys_no = KYK.kys_no) "
    // tmp_sql = tmp_sql & " LEFT JOIN kys_mst AS KYS ON SQ.kys_no = KYS.kys_no "
    // tmp_sql = tmp_sql & " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_jitu = MNKBN.nkbn_no "
    // tmp_sql = tmp_sql & " WHERE nkin_ymd IS NULL "
    // tmp_sql = tmp_sql & " AND   SQ.nkin_no = 6020 "

    // Return tmp_sql

    // End Function

    // Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As SafeDictionary<string, string>, ByRef hash_chkafter As SafeDictionary<string, string>, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

    // End Function

    // Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

    // End Sub

    // Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

    // End Function

    // Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

    // End Sub

    // End Class

    // End Class

    #endregion

    #region 運用開始時未収滞納金情報

    public class Sq_unyomisyutaino_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【抽出クエリ】
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";

                sortstr = "[物件No],[部屋No],[契約No],[契約レコードNo]";

                tmp_sql = tmp_sql + " /*************************************************************************************************** ";
                tmp_sql = tmp_sql + " ・以下の条件を満たす場合、運用開始時未納データとして移行する(unyotainodata へ移行) ";
                tmp_sql = tmp_sql + "   ・未入金であること ";
                tmp_sql = tmp_sql + "   ・「該当月が運用開始対象年月以降 かつ 該当月が適用開始月以降」以外であること ";
                tmp_sql = tmp_sql + "   ";
                tmp_sql = tmp_sql + " ・送金実績有無の2パターンで構成する ";
                tmp_sql = tmp_sql + "   ①過去送金実績が無い場合 ";
                tmp_sql = tmp_sql + "     →  条件に一致したものを移行 ";
                tmp_sql = tmp_sql + "   ②過去送金実績が有る場合 ";
                tmp_sql = tmp_sql + "     →  unyotainodata の sosumi_flg = 1 / so_rit = 0 / so_gak = 0 にして移行 ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + "   ※送金実績がない = 送金確定日が空 ";
                tmp_sql = tmp_sql + " ***************************************************************************************************/ ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " /************************** ";
                tmp_sql = tmp_sql + " ①過去送金実績が無い場合 ";
                tmp_sql = tmp_sql + " **************************/ ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 SQ.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 	,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 	,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 	,ko_no AS [契約レコードNo] ";
                tmp_sql = tmp_sql + " 	,gt_ym AS [該当年月] ";
                tmp_sql = tmp_sql + " 	,sq_simeymd AS [請求締め日] ";
                tmp_sql = tmp_sql + " 	/*,nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 	,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 	/*20160915 運用開始滞納金取得時の入金区分取得方法修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/*,SQ.nkin_kbn AS [入金項目区分]*/ ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN SQ.nkin_kbn = 7 OR SQ.nkin_kbn = 8 OR SQ.nkin_kbn = 9 THEN 9 ";
                tmp_sql = tmp_sql + " 		ELSE SQ.nkin_kbn ";
                tmp_sql = tmp_sql + " 	 END AS [入金項目区分] ";
                tmp_sql = tmp_sql + " 	/*20160915 運用開始滞納金取得時の入金区分取得方法修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	/*,nkbn_yotei AS [(入金予定)入金区分]*/ ";
                tmp_sql = tmp_sql + " 	,nkbn_name  AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 	,SQ.fkom_no AS [(入金予定)家賃入金口座] ";
                tmp_sql = tmp_sql + " 	,SQ.kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 	,sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(sq_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [請求税区分] ";
                tmp_sql = tmp_sql + " 	,sq_zeigak AS [請求税額] ";
                tmp_sql = tmp_sql + " 	,so_gak AS [送金額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [送金税区分]	 ";
                tmp_sql = tmp_sql + " 	,so_zeigak AS [送金税額] ";
                tmp_sql = tmp_sql + " 	,2 AS [既に送金済みフラグ] ";
                tmp_sql = tmp_sql + " 	,kanri_rit AS [管理手数料率] ";
                tmp_sql = tmp_sql + " 	,BKS.so_no AS [送金先家主No] ";
                tmp_sql = tmp_sql + " 	,BKS.sokoza_no AS [送金先家主口座No] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(sq_gak,0) <> 0 THEN (so_gak / sq_gak) * 100 ";
                tmp_sql = tmp_sql + " 		ELSE 0 ";
                tmp_sql = tmp_sql + " 	 END AS [送金率] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = -2 THEN 1 ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = -1 THEN 2 ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = 0 THEN 3 ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = 1 THEN 4 ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = 2 THEN 5 ";
                tmp_sql = tmp_sql + " 		ELSE 2 ";
                tmp_sql = tmp_sql + " 	 END AS [送金日決定方法 - 該当年月] ";
                tmp_sql = tmp_sql + " 	,sime1 AS [送金日決定方法 - 締日] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 2 THEN 3 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 3 THEN 4 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 4 THEN 5 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 5 THEN 6 ";
                tmp_sql = tmp_sql + " 		ELSE 2 ";
                tmp_sql = tmp_sql + " 	 END AS [送金日決定方法 - 送金月] ";
                tmp_sql = tmp_sql + " 	,so_dd1 AS [送金日決定方法 - 送金締日]	 ";
                tmp_sql = tmp_sql + " 	,'' AS [適用税率] ";
                tmp_sql = tmp_sql + " 	,0 AS [請求データ生成フラグ]	/*不明項目のためとりあえず0を設定しておく*/ ";
                tmp_sql = tmp_sql + " 	,'' AS [担当者No] ";
                tmp_sql = tmp_sql + " 	,'' AS [備考] ";
                tmp_sql = tmp_sql + " 	,kanri_zeikbn AS [管理手数料税フラグ] ";
                tmp_sql = tmp_sql + " 	,0 AS [管理手数料税込フラグ] ";
                tmp_sql = tmp_sql + " 	,0 AS [管理手数料率内税フラグ] ";
                tmp_sql = tmp_sql + " /*20160914 未収滞納金抽出クエリの修正 chg sta*/ ";
                tmp_sql = tmp_sql + " /* ";
                tmp_sql = tmp_sql + " FROM sq_meisai AS SQ ";
                tmp_sql = tmp_sql + " LEFT JOIN bk_kanri AS BKK ON SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,so_no,sokoza_no FROM bk_sokin WHERE daihyo = 1) AS BKS ON SQ.bk_no = BKS.bk_no ";
                tmp_sql = tmp_sql + " WHERE nkin_ymd IS NULL ";
                tmp_sql = tmp_sql + " AND   SQ.nkin_no <> 6020 ";
                tmp_sql = tmp_sql + " AND   so_kakuymd IS NULL			/*送金実績が無い*/ ";
                tmp_sql = tmp_sql + " AND   NOT ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/*/ ";
                tmp_sql = tmp_sql + " 			SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/			 ";
                tmp_sql = tmp_sql + " 			/*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg end*/ ";
                tmp_sql = tmp_sql + " 			AND ";
                tmp_sql = tmp_sql + " 			EXISTS					/*該当年月が適用開始年月以降*/ ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT * FROM bk_kanri AS BKK ";
                tmp_sql = tmp_sql + " 				WHERE SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 				AND   SQ.gt_ym >= BKK.start_ym				 ";
                tmp_sql = tmp_sql + " 			) ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " */ ";
                tmp_sql = tmp_sql + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 TMPSQ.* ";
                tmp_sql = tmp_sql + " 		,BKK.kn_no ";
                tmp_sql = tmp_sql + " 		,BKK.start_ym ";
                tmp_sql = tmp_sql + " 		,BKK.end_ym ";
                tmp_sql = tmp_sql + " 		,BKK.sime_kbn1 ";
                tmp_sql = tmp_sql + " 		,BKK.sime1 ";
                tmp_sql = tmp_sql + " 		,BKK.som_kbn1 ";
                tmp_sql = tmp_sql + " 		,BKK.so_dd1 ";
                tmp_sql = tmp_sql + " 	FROM sq_meisai AS TMPSQ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bk_kanri AS BKK ON TMPSQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 	WHERE TMPSQ.sq_simeymd BETWEEN BKK.start_ym AND BKK.end_ym ";
                tmp_sql = tmp_sql + " ) AS SQ ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,kn_no,so_no,sokoza_no FROM bk_sokin WHERE daihyo = 1) AS BKS ";
                tmp_sql = tmp_sql + " ON SQ.bk_no = BKS.bk_no AND SQ.kn_no = BKS.kn_no ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " WHERE nkin_ymd IS NULL ";
                tmp_sql = tmp_sql + " AND   SQ.nkin_no <> 6020 ";
                tmp_sql = tmp_sql + " AND   so_kakuymd IS NULL			/*送金実績が無い*/ ";
                tmp_sql = tmp_sql + " AND   NOT ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/*/ ";
                tmp_sql = tmp_sql + " 			SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/			 ";
                tmp_sql = tmp_sql + " 			/*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg end*/ ";
                tmp_sql = tmp_sql + " 			AND ";
                tmp_sql = tmp_sql + " 			EXISTS					/*該当年月が適用開始年月以降*/ ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT * FROM bk_kanri AS BKK ";
                tmp_sql = tmp_sql + " 				WHERE SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 				AND   SQ.gt_ym >= BKK.start_ym				 ";
                tmp_sql = tmp_sql + " 			) ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " /*20160914 未収滞納金抽出クエリの修正 chg end*/ ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " /************************** ";
                tmp_sql = tmp_sql + " ①過去送金実績が有る場合 ";
                tmp_sql = tmp_sql + " **************************/ ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 SQ.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 	,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 	,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 	,ko_no AS [契約レコードNo] ";
                tmp_sql = tmp_sql + " 	,gt_ym AS [該当年月] ";
                tmp_sql = tmp_sql + " 	,sq_simeymd AS [請求締め日] ";
                tmp_sql = tmp_sql + " 	/*,nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 	,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 	/*20160915 運用開始滞納金取得時の入金区分取得方法修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/*,SQ.nkin_kbn AS [入金項目区分]*/ ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN SQ.nkin_kbn = 7 OR SQ.nkin_kbn = 8 OR SQ.nkin_kbn = 9 THEN 9 ";
                tmp_sql = tmp_sql + " 		ELSE SQ.nkin_kbn ";
                tmp_sql = tmp_sql + " 	 END AS [入金項目区分] ";
                tmp_sql = tmp_sql + " 	/*20160915 運用開始滞納金取得時の入金区分取得方法修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	/*,nkbn_yotei AS [(入金予定)入金区分]*/ ";
                tmp_sql = tmp_sql + " 	,nkbn_name  AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 	,SQ.fkom_no AS [(入金予定)家賃入金口座] ";
                tmp_sql = tmp_sql + " 	,SQ.kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 	,sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(sq_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [請求税区分] ";
                tmp_sql = tmp_sql + " 	,sq_zeigak AS [請求税額] ";
                tmp_sql = tmp_sql + " 	,0 AS [送金額]					/*送金済みを示す*/ ";
                tmp_sql = tmp_sql + " 	,1 AS [送金税区分]	 ";
                tmp_sql = tmp_sql + " 	,0 AS [送金税額]				/*送金済みを示す*/ ";
                tmp_sql = tmp_sql + " 	,1 AS [既に送金済みフラグ]		/*送金済みを示す*/ ";
                tmp_sql = tmp_sql + " 	,kanri_rit AS [管理手数料率] ";
                tmp_sql = tmp_sql + " 	,BKS.so_no AS [送金先家主No] ";
                tmp_sql = tmp_sql + " 	,BKS.sokoza_no AS [送金先家主口座No] ";
                tmp_sql = tmp_sql + " 	,0 AS [送金率]					/*送金済みを示す*/ ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = -2 THEN 1 ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = -1 THEN 2 ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = 0 THEN 3 ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = 1 THEN 4 ";
                tmp_sql = tmp_sql + " 		WHEN sime_kbn1 = 2 THEN 5 ";
                tmp_sql = tmp_sql + " 		ELSE -1 ";
                tmp_sql = tmp_sql + " 	 END AS [送金日決定方法 - 該当年月] ";
                tmp_sql = tmp_sql + " 	,sime1 AS [送金日決定方法 - 締日] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 2 THEN 3 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 3 THEN 4 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 4 THEN 5 ";
                tmp_sql = tmp_sql + " 		WHEN som_kbn1 = 5 THEN 6 ";
                tmp_sql = tmp_sql + " 		ELSE -1 ";
                tmp_sql = tmp_sql + " 	 END AS [送金日決定方法 - 送金月] ";
                tmp_sql = tmp_sql + " 	,so_dd1 AS [送金日決定方法 - 送金締日]	 ";
                tmp_sql = tmp_sql + " 	,'' AS [適用税率] ";
                tmp_sql = tmp_sql + " 	,0 AS [請求データ生成フラグ]	/*不明項目のためとりあえず0を設定しておく*/ ";
                tmp_sql = tmp_sql + " 	,'' AS [担当者No] ";
                tmp_sql = tmp_sql + " 	/*20160608 対応依頼分 -chg sta*/ ";
                tmp_sql = tmp_sql + " 	/*,'' AS [備考]*/ ";
                tmp_sql = tmp_sql + " 	,'※送金済みデータ' AS [備考] ";
                tmp_sql = tmp_sql + " 	/*20160608 対応依頼分 -chg end*/ ";
                tmp_sql = tmp_sql + " 	,kanri_zeikbn AS [管理手数料税フラグ] ";
                tmp_sql = tmp_sql + " 	,0 AS [管理手数料税込フラグ] ";
                tmp_sql = tmp_sql + " 	,0 AS [管理手数料率内税フラグ] ";
                tmp_sql = tmp_sql + " /*20160914 未収滞納金抽出クエリの修正 chg sta*/ ";
                tmp_sql = tmp_sql + " /* ";
                tmp_sql = tmp_sql + " FROM sq_meisai AS SQ ";
                tmp_sql = tmp_sql + " LEFT JOIN bk_kanri AS BKK ON SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,so_no,sokoza_no FROM bk_sokin WHERE daihyo = 1) AS BKS ON SQ.bk_no = BKS.bk_no ";
                tmp_sql = tmp_sql + " WHERE nkin_ymd IS NULL ";
                tmp_sql = tmp_sql + " AND   SQ.nkin_no <> 6020 ";
                tmp_sql = tmp_sql + " AND   so_kakuymd IS NOT NULL		/*送金実績が有る*/ ";
                tmp_sql = tmp_sql + " AND   NOT ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/*/ ";
                tmp_sql = tmp_sql + " 			SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/			 ";
                tmp_sql = tmp_sql + " 			/*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg end*/ ";
                tmp_sql = tmp_sql + " 			AND ";
                tmp_sql = tmp_sql + " 			EXISTS					/*該当年月が適用開始年月以降*/ ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT * FROM bk_kanri AS BKK ";
                tmp_sql = tmp_sql + " 				WHERE SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 				AND   SQ.gt_ym >= BKK.start_ym				 ";
                tmp_sql = tmp_sql + " 			) ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " */ ";
                tmp_sql = tmp_sql + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 TMPSQ.* ";
                tmp_sql = tmp_sql + " 		,BKK.kn_no ";
                tmp_sql = tmp_sql + " 		,BKK.start_ym ";
                tmp_sql = tmp_sql + " 		,BKK.end_ym ";
                tmp_sql = tmp_sql + " 		,BKK.sime_kbn1 ";
                tmp_sql = tmp_sql + " 		,BKK.sime1 ";
                tmp_sql = tmp_sql + " 		,BKK.som_kbn1 ";
                tmp_sql = tmp_sql + " 		,BKK.so_dd1 ";
                tmp_sql = tmp_sql + " 	FROM sq_meisai AS TMPSQ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bk_kanri AS BKK ON TMPSQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 	WHERE TMPSQ.sq_simeymd BETWEEN BKK.start_ym AND BKK.end_ym ";
                tmp_sql = tmp_sql + " ) AS SQ ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,kn_no,so_no,sokoza_no FROM bk_sokin WHERE daihyo = 1) AS BKS ";
                tmp_sql = tmp_sql + " ON SQ.bk_no = BKS.bk_no AND SQ.kn_no = BKS.kn_no ";
                tmp_sql = tmp_sql + "  ";
                tmp_sql = tmp_sql + " WHERE nkin_ymd IS NULL ";
                tmp_sql = tmp_sql + " AND   SQ.nkin_no <> 6020 ";
                tmp_sql = tmp_sql + " AND   so_kakuymd IS NOT NULL		/*送金実績が有る*/ ";
                tmp_sql = tmp_sql + " AND   NOT ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/*/ ";
                tmp_sql = tmp_sql + " 			SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/			 ";
                tmp_sql = tmp_sql + " 			/*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg end*/ ";
                tmp_sql = tmp_sql + " 			AND ";
                tmp_sql = tmp_sql + " 			EXISTS					/*該当年月が適用開始年月以降*/ ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT * FROM bk_kanri AS BKK ";
                tmp_sql = tmp_sql + " 				WHERE SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 				AND   SQ.gt_ym >= BKK.start_ym				 ";
                tmp_sql = tmp_sql + " 			) ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " /*20160914 未収滞納金抽出クエリの修正 chg end*/ ";

                // 20160525 運用開始年月の追加 -add
                tmp_sql = tmp_sql.Replace("運用開始年月置換用文字列", CommonModule.UnyoYMD);

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

            }

            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {
                return default;

            }

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

            }

        }

    }

    #endregion

    #region 請求情報

    public class Sq_sq_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【抽出クエリ】
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";

                sortstr = "[物件No],[部屋No],[契約No],[契約レコードNo]";

                // 20160601 請求データの抽出クエリ修正 -chg sta
                // tmp_sql = tmp_sql & " /****************************************************************************** "
                // tmp_sql = tmp_sql & " その他請求の移行 "
                // tmp_sql = tmp_sql & "   ・その他請求項目以外は請求データ構築で作成される (運用開始年月以降) "
                // tmp_sql = tmp_sql & "   ・その他請求項目は作成されないため移行する必要がある (sqdataへ移行) "
                // tmp_sql = tmp_sql & "   ・移行条件 "
                // tmp_sql = tmp_sql & "       ・その他請求項目 "
                // tmp_sql = tmp_sql & "         ※データ区分 = 6 のその他請求を抽出対象とするが "
                // tmp_sql = tmp_sql & "           過剰金 (6020) は移行先が異なるため "
                // tmp_sql = tmp_sql & "       ・コンバート対象データの該当月が "
                // tmp_sql = tmp_sql & "         「運用開始年月」以降、かつ該当物件の「送金ルール開始年月」以降 "
                // tmp_sql = tmp_sql & " ******************************************************************************/ "
                // tmp_sql = tmp_sql & " SELECT "
                // tmp_sql = tmp_sql & " 	 SQ.bk_no AS [物件No] "
                // tmp_sql = tmp_sql & " 	,SQ.hy_no AS [部屋No] "
                // tmp_sql = tmp_sql & " 	,SQ.ky_no AS [契約No] "
                // tmp_sql = tmp_sql & " 	,SQ.ko_no AS [契約レコードNo]	 "
                // tmp_sql = tmp_sql & " 	,gt_ym AS [該当年月] "
                // tmp_sql = tmp_sql & " 	/*,nkin_no AS [入金項目No]*/ "
                // tmp_sql = tmp_sql & " 	,nkin_name AS [入金項目名] "
                // tmp_sql = tmp_sql & " 	,NULL AS [入金項目区分]		/*10で請求データ作成時にNULLになることを確認*/	 "
                // tmp_sql = tmp_sql & " 	,'' AS [入金項目表示用名称] "
                // tmp_sql = tmp_sql & " 	/*,nkbn_yotei AS [(入金予定)入金区分]*/ "
                // tmp_sql = tmp_sql & " 	,nkbn_name  AS [(入金予定)入金区分] "
                // tmp_sql = tmp_sql & " 	,SQ.sq_gak AS [請求額] "
                // tmp_sql = tmp_sql & " 	,CASE "
                // tmp_sql = tmp_sql & " 		WHEN ISNULL(sq_zeigak,0) <> 0 THEN 2 "
                // tmp_sql = tmp_sql & " 		ELSE 1 "
                // tmp_sql = tmp_sql & " 	 END AS [請求税区分]	 "
                // tmp_sql = tmp_sql & " 	,sq_zeigak AS [請求税額] "
                // tmp_sql = tmp_sql & " 	,sq_simeymd AS [請求締切日] "
                // tmp_sql = tmp_sql & " 	,'' AS [複数請求先フラグ] "
                // tmp_sql = tmp_sql & " 	,100 AS [請求先区分]	/*契約者固定*/ "
                // tmp_sql = tmp_sql & " 	,SQ.kys_no AS [請求先No] "
                // tmp_sql = tmp_sql & " 	,SQ.fkom_no AS [(入金予定)家賃振込先口座No] "
                // tmp_sql = tmp_sql & " 	,fkae_sqno AS [口座振替履歴No]	/*要確認*/ "
                // tmp_sql = tmp_sql & " 	,'' AS [口座振替履歴明細No] "
                // tmp_sql = tmp_sql & " 	/*20160530 送金データ調査での不足箇所修正 chg sta*/ "
                // tmp_sql = tmp_sql & " 	/* "
                // tmp_sql = tmp_sql & " 	,fkae_sqflg AS [振替請求対象フラグ 1:請求中　2:未請求]	/*要確認*/ "
                // tmp_sql = tmp_sql & " 	,fkae_sqtaiflg AS [振替請求対象フラグ　1:対象　2:対象外]	/*要確認*/ "
                // tmp_sql = tmp_sql & " 	,fkae_sqflg AS [振替請求中フラグ]	/*要確認*/ "
                // tmp_sql = tmp_sql & " 	*/ "
                // tmp_sql = tmp_sql & " 	,'' AS [振替請求対象フラグ 1:請求中　2:未請求] "
                // tmp_sql = tmp_sql & " 	,1 AS [振替請求対象フラグ　1:対象　2:対象外] "
                // tmp_sql = tmp_sql & " 	,0 AS [振替請求中フラグ] "
                // tmp_sql = tmp_sql & " 	/*20160530 送金データ調査での不足箇所修正 chg end*/ "
                // tmp_sql = tmp_sql & " 	,NULL AS [月区分]			/*10で請求データ作成時にNULLになることを確認*/ "
                // tmp_sql = tmp_sql & " 	,NULL AS [入金項目明細No]	/*10で請求データ作成時にNULLになることを確認*/ "
                // tmp_sql = tmp_sql & " 	/*,'' AS [契約入金項目GUID]	契約入金データから取得する*/ "
                // tmp_sql = tmp_sql & " 	,NULL AS [次回契約情報フラグ] "
                // tmp_sql = tmp_sql & " 	,NULL AS [変動費請求GUID] "
                // tmp_sql = tmp_sql & " 	,NULL AS [変動費区分] "
                // tmp_sql = tmp_sql & " 	,NULL AS [変動費検針日] "
                // tmp_sql = tmp_sql & " 	,SQ.so_gak AS [送金額] "
                // tmp_sql = tmp_sql & " 	,CASE "
                // tmp_sql = tmp_sql & " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 2 "
                // tmp_sql = tmp_sql & " 		ELSE 1 "
                // tmp_sql = tmp_sql & " 	 END AS [送金税区分]	 "
                // tmp_sql = tmp_sql & " 	,so_zeigak AS [送金税額] "
                // tmp_sql = tmp_sql & " 	,1 AS [その他請求フラグ] "
                // tmp_sql = tmp_sql & " 	,CASE "
                // tmp_sql = tmp_sql & " 		WHEN ISNULL(SQ.sq_gak,0) <> 0 THEN (SQ.so_gak / SQ.sq_gak) * 100 "
                // tmp_sql = tmp_sql & " 		ELSE 0 "
                // tmp_sql = tmp_sql & " 	 END AS [その他請求の送金率]	 "
                // tmp_sql = tmp_sql & " 	,so_yoteiymd AS [その他請求の指定送金日] "
                // tmp_sql = tmp_sql & " 	,'' AS [分割GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [支払・返金区分] "
                // tmp_sql = tmp_sql & " 	,'' AS [支払・返金先区分] "
                // tmp_sql = tmp_sql & " 	/*'20160530 送金データ調査での不足箇所修正 chg sta*/ "
                // tmp_sql = tmp_sql & " 	/* "
                // tmp_sql = tmp_sql & " 	,'' AS [支払・返金先No] "
                // tmp_sql = tmp_sql & " 	,'' AS [支払・返金先口座No] "
                // tmp_sql = tmp_sql & " 	*/ "
                // tmp_sql = tmp_sql & " 	,so_no AS [支払・返金先No] "
                // tmp_sql = tmp_sql & " 	,sokoza_no AS [支払・返金先口座No] "
                // tmp_sql = tmp_sql & " 	/*'20160530 送金データ調査での不足箇所修正 chg end*/ "
                // tmp_sql = tmp_sql & " 	,'' AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] "
                // tmp_sql = tmp_sql & " 	,'' AS [支店No] "
                // tmp_sql = tmp_sql & " 	,'' AS [担当者No] "
                // tmp_sql = tmp_sql & " 	,'' AS [控除請求項目GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [固定控除ルールGUID(請求)] "
                // tmp_sql = tmp_sql & " 	,'' AS [控除ルールフラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [請求書発行予定日] "
                // tmp_sql = tmp_sql & " 	,'' AS [請求ソートNo] "
                // tmp_sql = tmp_sql & " 	,'' AS [立替・預りフラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [立替・預り設定時の預り予定GUID（預り金データ）] "
                // tmp_sql = tmp_sql & " 	,'' AS [敷金保証金随時処理GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [運用開始時の未納・滞納金GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [送金回収GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [修繕区分] "
                // tmp_sql = tmp_sql & " 	,'' AS [随時修繕No] "
                // tmp_sql = tmp_sql & " 	,'' AS [修繕請求No] "
                // tmp_sql = tmp_sql & " 	,'' AS [適用税率] "
                // tmp_sql = tmp_sql & " 	,'' AS [編集フラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [更新ロックフラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [備考] "
                // tmp_sql = tmp_sql & " 	,'' AS [立替仕訳額] "
                // tmp_sql = tmp_sql & " 	,'' AS [立替仕訳税額] "
                // tmp_sql = tmp_sql & " 	,'' AS [取込日（インポート実施日）] "
                // tmp_sql = tmp_sql & " 	,'' AS [コンビニ収納サービスNo] "
                // tmp_sql = tmp_sql & " 	,'' AS [顧客番号] "
                // tmp_sql = tmp_sql & " 	,'' AS [コンビニ収納請求フラグ] "
                // tmp_sql = tmp_sql & " 	,2 AS [その他請求管理手数料区分] "
                // tmp_sql = tmp_sql & " 	,'' AS [未収仕訳フラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [雑収入区分] "
                // tmp_sql = tmp_sql & " 	,so_kakuymd AS [送金確定日]		/*10送金情報作成用*/ "
                // tmp_sql = tmp_sql & " 	,tnh_umu AS [滞納保証フラグ]	/*10送金情報作成用*/ "
                // tmp_sql = tmp_sql & " FROM sq_meisai AS SQ "
                // tmp_sql = tmp_sql & " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no "
                // tmp_sql = tmp_sql & " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no "
                // tmp_sql = tmp_sql & " LEFT JOIN (SELECT bk_no,start_ym,end_ym,tnh_umu FROM bk_kanri) AS BKK ON SQ.bk_no = BKK.bk_no "
                // tmp_sql = tmp_sql & " WHERE nkin_ymd IS NULL "
                // tmp_sql = tmp_sql & " AND   data_kbn = 6 "
                // tmp_sql = tmp_sql & " AND   SQ.nkin_no <> 6020		/*入金項目No:6020 は過剰金*/ "
                // tmp_sql = tmp_sql & " AND   SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/ "
                // tmp_sql = tmp_sql & " AND   EXISTS					/*該当年月が適用開始年月以降*/ "
                // tmp_sql = tmp_sql & " 		( "
                // tmp_sql = tmp_sql & " 			SELECT * FROM bk_kanri AS BKK "
                // tmp_sql = tmp_sql & " 			WHERE SQ.bk_no = BKK.bk_no "
                // tmp_sql = tmp_sql & " 			AND   SQ.gt_ym >= BKK.start_ym "
                // tmp_sql = tmp_sql & " 		) "
                // tmp_sql = tmp_sql & " AND   (so_kakuymd IS NULL OR so_kakuymd >= start_ym AND so_kakuymd <= end_ym) "
                // tmp_sql = tmp_sql & " /****************************************************************************** "
                // tmp_sql = tmp_sql & " 部分入金の移行 "
                // tmp_sql = tmp_sql & "   ・部分入金されている入金項目のうち、未入金部分はは請求データ構築で "
                // tmp_sql = tmp_sql & "     作成されないため移行する必要がある "
                // tmp_sql = tmp_sql & "   ・移行条件 "
                // tmp_sql = tmp_sql & "       ・その他請求項目以外 (データ区分 <> 6) "
                // tmp_sql = tmp_sql & "       ・コンバート対象データの該当月が "
                // tmp_sql = tmp_sql & "         「運用開始年月」以降、かつ該当物件の「送金ルール開始年月」以降 "
                // tmp_sql = tmp_sql & " ******************************************************************************/ "
                // tmp_sql = tmp_sql & " UNION "
                // tmp_sql = tmp_sql & " SELECT "
                // tmp_sql = tmp_sql & " 	 SQ.bk_no AS [物件No] "
                // tmp_sql = tmp_sql & " 	,SQ.hy_no AS [部屋No] "
                // tmp_sql = tmp_sql & " 	,SQ.ky_no AS [契約No] "
                // tmp_sql = tmp_sql & " 	,SQ.ko_no AS [契約レコードNo]	 "
                // tmp_sql = tmp_sql & " 	,gt_ym AS [該当年月] "
                // tmp_sql = tmp_sql & " 	/*,nkin_no AS [入金項目No]*/ "
                // tmp_sql = tmp_sql & " 	,nkin_name AS [入金項目名] "
                // tmp_sql = tmp_sql & " 	,SQ.nkin_kbn AS [入金項目区分]	 "
                // tmp_sql = tmp_sql & " 	,'' AS [入金項目表示用名称] "
                // tmp_sql = tmp_sql & " 	/*,nkbn_yotei AS [(入金予定)入金区分]*/ "
                // tmp_sql = tmp_sql & " 	,nkbn_name  AS [(入金予定)入金区分] "
                // tmp_sql = tmp_sql & " 	,SQ.sq_gak AS [請求額] "
                // tmp_sql = tmp_sql & " 	,CASE "
                // tmp_sql = tmp_sql & " 		WHEN ISNULL(sq_zeigak,0) <> 0 THEN 2 "
                // tmp_sql = tmp_sql & " 		ELSE 1 "
                // tmp_sql = tmp_sql & " 	 END AS [請求税区分]	 "
                // tmp_sql = tmp_sql & " 	,sq_zeigak AS [請求税額] "
                // tmp_sql = tmp_sql & " 	,sq_simeymd AS [請求締切日] "
                // tmp_sql = tmp_sql & " 	,'' AS [複数請求先フラグ] "
                // tmp_sql = tmp_sql & " 	,100 AS [請求先区分]	/*契約者固定*/ "
                // tmp_sql = tmp_sql & " 	,SQ.kys_no AS [請求先No] "
                // tmp_sql = tmp_sql & " 	,SQ.fkom_no AS [(入金予定)家賃振込先口座No] "
                // tmp_sql = tmp_sql & " 	,fkae_sqno AS [口座振替履歴No]	/*要確認*/ "
                // tmp_sql = tmp_sql & " 	,'' AS [口座振替履歴明細No] "
                // tmp_sql = tmp_sql & " 	/*20160530 送金データ調査での不足箇所修正 chg sta*/ "
                // tmp_sql = tmp_sql & " 	/* "
                // tmp_sql = tmp_sql & " 	,fkae_sqflg AS [振替請求対象フラグ 1:請求中　2:未請求]	/*要確認*/ "
                // tmp_sql = tmp_sql & " 	,fkae_sqtaiflg AS [振替請求対象フラグ　1:対象　2:対象外]	/*要確認*/ "
                // tmp_sql = tmp_sql & " 	,fkae_sqflg AS [振替請求中フラグ]	/*要確認*/ "
                // tmp_sql = tmp_sql & " 	*/ "
                // tmp_sql = tmp_sql & " 	,'' AS [振替請求対象フラグ 1:請求中　2:未請求] "
                // tmp_sql = tmp_sql & " 	,1 AS [振替請求対象フラグ　1:対象　2:対象外] "
                // tmp_sql = tmp_sql & " 	,0 AS [振替請求中フラグ] "
                // tmp_sql = tmp_sql & " 	/*20160530 送金データ調査での不足箇所修正 chg end*/ "
                // tmp_sql = tmp_sql & " 	,CASE "
                // tmp_sql = tmp_sql & " 		WHEN data_kbn = 1 THEN 2 "
                // tmp_sql = tmp_sql & " 		WHEN data_kbn >= 2 THEN 1 "
                // tmp_sql = tmp_sql & " 	 END AS [月区分] "
                // tmp_sql = tmp_sql & " 	,1 AS [入金項目明細No] "
                // tmp_sql = tmp_sql & " 	/*,'' AS [契約入金項目GUID]	契約入金データから取得する*/ "
                // tmp_sql = tmp_sql & " 	,'' AS [次回契約情報フラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [変動費請求GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [変動費区分] "
                // tmp_sql = tmp_sql & " 	,'' AS [変動費検針日] "
                // tmp_sql = tmp_sql & " 	,SQ.so_gak AS [送金額] "
                // tmp_sql = tmp_sql & " 	,CASE "
                // tmp_sql = tmp_sql & " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 2 "
                // tmp_sql = tmp_sql & " 		ELSE 1 "
                // tmp_sql = tmp_sql & " 	 END AS [送金税区分]	 "
                // tmp_sql = tmp_sql & " 	,so_zeigak AS [送金税額] "
                // tmp_sql = tmp_sql & " 	,0 AS [その他請求フラグ] "
                // tmp_sql = tmp_sql & " 	,0 AS [その他請求の送金率]	 "
                // tmp_sql = tmp_sql & " 	,NULL AS [その他請求の指定送金日] "
                // tmp_sql = tmp_sql & " 	,'' AS [分割GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [支払・返金区分] "
                // tmp_sql = tmp_sql & " 	,'' AS [支払・返金先区分] "
                // tmp_sql = tmp_sql & " 	/*'20160530 送金データ調査での不足箇所修正 chg sta*/ "
                // tmp_sql = tmp_sql & " 	/* "
                // tmp_sql = tmp_sql & " 	,'' AS [支払・返金先No] "
                // tmp_sql = tmp_sql & " 	,'' AS [支払・返金先口座No] "
                // tmp_sql = tmp_sql & " 	*/ "
                // tmp_sql = tmp_sql & " 	,so_no AS [支払・返金先No] "
                // tmp_sql = tmp_sql & " 	,sokoza_no AS [支払・返金先口座No] "
                // tmp_sql = tmp_sql & " 	/*'20160530 送金データ調査での不足箇所修正 chg end*/ "
                // tmp_sql = tmp_sql & " 	,'' AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] "
                // tmp_sql = tmp_sql & " 	,'' AS [支店No] "
                // tmp_sql = tmp_sql & " 	,'' AS [担当者No] "
                // tmp_sql = tmp_sql & " 	,'' AS [控除請求項目GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [固定控除ルールGUID(請求)] "
                // tmp_sql = tmp_sql & " 	,'' AS [控除ルールフラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [請求書発行予定日] "
                // tmp_sql = tmp_sql & " 	,'' AS [請求ソートNo] "
                // tmp_sql = tmp_sql & " 	,'' AS [立替・預りフラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [立替・預り設定時の預り予定GUID（預り金データ）] "
                // tmp_sql = tmp_sql & " 	,'' AS [敷金保証金随時処理GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [運用開始時の未納・滞納金GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [送金回収GUID] "
                // tmp_sql = tmp_sql & " 	,'' AS [修繕区分] "
                // tmp_sql = tmp_sql & " 	,'' AS [随時修繕No] "
                // tmp_sql = tmp_sql & " 	,'' AS [修繕請求No] "
                // tmp_sql = tmp_sql & " 	,'' AS [NULL] "
                // tmp_sql = tmp_sql & " 	,'' AS [編集フラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [更新ロックフラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [備考] "
                // tmp_sql = tmp_sql & " 	,'' AS [立替仕訳額] "
                // tmp_sql = tmp_sql & " 	,'' AS [立替仕訳税額] "
                // tmp_sql = tmp_sql & " 	,'' AS [取込日（インポート実施日）] "
                // tmp_sql = tmp_sql & " 	,'' AS [コンビニ収納サービスNo] "
                // tmp_sql = tmp_sql & " 	,'' AS [顧客番号] "
                // tmp_sql = tmp_sql & " 	,'' AS [コンビニ収納請求フラグ] "
                // tmp_sql = tmp_sql & " 	,NULL AS [その他請求管理手数料区分] "
                // tmp_sql = tmp_sql & " 	,'' AS [未収仕訳フラグ] "
                // tmp_sql = tmp_sql & " 	,'' AS [雑収入区分] "
                // tmp_sql = tmp_sql & " 	,so_kakuymd AS [送金確定日]		/*10送金情報作成用*/ "
                // tmp_sql = tmp_sql & " 	,tnh_umu AS [滞納保証フラグ]	/*10送金情報作成用*/	 "
                // tmp_sql = tmp_sql & " FROM sq_meisai AS SQ "
                // tmp_sql = tmp_sql & " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no "
                // tmp_sql = tmp_sql & " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no "
                // tmp_sql = tmp_sql & " LEFT JOIN (SELECT bk_no,start_ym,end_ym,tnh_umu FROM bk_kanri) AS BKK ON SQ.bk_no = BKK.bk_no "
                // tmp_sql = tmp_sql & " WHERE rec_kbn <> 0 "
                // tmp_sql = tmp_sql & " AND   nkin_ymd IS NULL "
                // tmp_sql = tmp_sql & " AND   data_kbn <> 6 "
                // tmp_sql = tmp_sql & " AND   SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/ "
                // tmp_sql = tmp_sql & " AND   EXISTS					/*該当年月が適用開始年月以降*/ "
                // tmp_sql = tmp_sql & " 		( "
                // tmp_sql = tmp_sql & " 			SELECT * FROM bk_kanri AS BKK "
                // tmp_sql = tmp_sql & " 			WHERE SQ.bk_no = BKK.bk_no "
                // tmp_sql = tmp_sql & " 			AND   SQ.gt_ym >= BKK.start_ym "
                // tmp_sql = tmp_sql & " 		) "
                // tmp_sql = tmp_sql & " AND   (so_kakuymd IS NULL OR so_kakuymd >= start_ym AND so_kakuymd <= end_ym) "

                tmp_sql = tmp_sql + " /****************************************************************************** ";
                tmp_sql = tmp_sql + " その他請求の移行 ";
                tmp_sql = tmp_sql + "   ・その他請求項目以外は請求データ構築で作成される (運用開始年月以降) ";
                tmp_sql = tmp_sql + "   ・その他請求項目は作成されないため移行する必要がある (sqdataへ移行) ";
                tmp_sql = tmp_sql + "   ・移行条件 ";
                tmp_sql = tmp_sql + "       ・その他請求項目 ";
                tmp_sql = tmp_sql + "         ※データ区分 = 6 のその他請求を抽出対象とするが ";
                tmp_sql = tmp_sql + "           過剰金 (6020) は移行先が異なるため抽出しない ";
                tmp_sql = tmp_sql + "       ・コンバート対象データの該当月が ";
                tmp_sql = tmp_sql + "         「運用開始年月」以降、かつ該当物件の「送金ルール開始年月」以降 ";
                tmp_sql = tmp_sql + " ******************************************************************************/ ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 SQ.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 	,SQ.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 	,SQ.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/*,SQ.ko_no AS [契約レコードNo]*/ ";
                tmp_sql = tmp_sql + " 	,KYK.[契約レコードNo] ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	,gt_ym AS [該当年月] ";
                tmp_sql = tmp_sql + " 	/*,nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 	,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目区分]		/*10で請求データ作成時にNULLになることを確認*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目表示用名称] ";
                tmp_sql = tmp_sql + " 	/*,nkbn_yotei AS [(入金予定)入金区分]*/ ";
                tmp_sql = tmp_sql + " 	,nkbn_name  AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 	,SQ.sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(sq_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [請求税区分]	 ";
                tmp_sql = tmp_sql + " 	,sq_zeigak AS [請求税額] ";
                tmp_sql = tmp_sql + " 	,sq_simeymd AS [請求締切日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [複数請求先フラグ] ";
                tmp_sql = tmp_sql + " 	,100 AS [請求先区分]	/*契約者固定*/ ";
                tmp_sql = tmp_sql + " 	,SQ.kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 	,SQ.fkom_no AS [(入金予定)家賃振込先口座No] ";
                tmp_sql = tmp_sql + " 	,fkae_sqno AS [口座振替履歴No]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [口座振替履歴明細No] ";
                tmp_sql = tmp_sql + " 	/*20160530 送金データ調査での不足箇所修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/* ";
                tmp_sql = tmp_sql + " 	,fkae_sqflg AS [振替請求対象フラグ 1:請求中　2:未請求]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,fkae_sqtaiflg AS [振替請求対象フラグ　1:対象　2:対象外]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,fkae_sqflg AS [振替請求中フラグ]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [振替請求対象フラグ 1:請求中　2:未請求] ";
                tmp_sql = tmp_sql + " 	,1 AS [振替請求対象フラグ　1:対象　2:対象外] ";
                tmp_sql = tmp_sql + " 	,0 AS [振替請求中フラグ] ";
                tmp_sql = tmp_sql + " 	/*20160530 送金データ調査での不足箇所修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [月区分]			/*10で請求データ作成時にNULLになることを確認*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目明細No]	/*10で請求データ作成時にNULLになることを確認*/ ";
                tmp_sql = tmp_sql + " 	/*,NULL AS [契約入金項目GUID]	契約入金データから取得する*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [次回契約情報フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費請求GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費検針日] ";
                tmp_sql = tmp_sql + " 	,SQ.so_gak AS [送金額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [送金税区分]	 ";
                tmp_sql = tmp_sql + " 	,so_zeigak AS [送金税額] ";
                tmp_sql = tmp_sql + " 	,1 AS [その他請求フラグ] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(SQ.sq_gak,0) <> 0 THEN (SQ.so_gak / SQ.sq_gak) * 100 ";
                tmp_sql = tmp_sql + " 		ELSE 0 ";
                tmp_sql = tmp_sql + " 	 END AS [その他請求の送金率]	 ";
                tmp_sql = tmp_sql + " 	,so_yoteiymd AS [その他請求の指定送金日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [分割GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先区分] ";
                tmp_sql = tmp_sql + " 	/*'20160530 送金データ調査での不足箇所修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/* ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先口座No] ";
                tmp_sql = tmp_sql + " 	*/ ";
                tmp_sql = tmp_sql + " 	,so_no AS [支払・返金先No] ";
                tmp_sql = tmp_sql + " 	,sokoza_no AS [支払・返金先口座No] ";
                tmp_sql = tmp_sql + " 	/*'20160530 送金データ調査での不足箇所修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支店No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [担当者No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除請求項目GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [固定控除ルールGUID(請求)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除ルールフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求書発行予定日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求ソートNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預り設定時の預り予定GUID（預り金データ）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [敷金保証金随時処理GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [運用開始時の未納・滞納金GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金回収GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [随時修繕No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕請求No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [適用税率] ";
                tmp_sql = tmp_sql + " 	,NULL AS [編集フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [更新ロックフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [備考] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳税額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [取込日（インポート実施日）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納サービスNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [顧客番号] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納請求フラグ] ";
                tmp_sql = tmp_sql + " 	,2 AS [その他請求管理手数料区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [未収仕訳フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [雑収入区分] ";
                tmp_sql = tmp_sql + " 	,so_kakuymd AS [送金確定日]		/*10送金情報作成用*/ ";
                tmp_sql = tmp_sql + " 	,tnh_umu AS [滞納保証フラグ]	/*10送金情報作成用*/ ";
                tmp_sql = tmp_sql + " 	,'その他請求' AS [請求区分]		/*作業用*/ ";
                tmp_sql = tmp_sql + " FROM sq_meisai AS SQ ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,start_ym,end_ym,tnh_umu FROM bk_kanri) AS BKK ON SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " /*20160603 ユーザーデータ検証による修正 add sta*/ ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,hy_no,ky_no,ko_no,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約レコードNo] FROM ky_kosinkai) AS KYK ";
                tmp_sql = tmp_sql + " ON  SQ.bk_no = KYK.bk_no AND SQ.hy_no = KYK.hy_no AND SQ.ky_no = KYK.ky_no AND SQ.ko_no = KYK.ko_no ";
                tmp_sql = tmp_sql + " /*20160603 ユーザーデータ検証による修正 add end*/ ";
                tmp_sql = tmp_sql + " WHERE nkin_ymd IS NULL ";
                tmp_sql = tmp_sql + " AND   data_kbn = 6 ";
                tmp_sql = tmp_sql + " AND   SQ.nkin_no <> 6020		/*入金項目No:6020 は過剰金*/ ";
                tmp_sql = tmp_sql + " /*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg sta*/ ";
                tmp_sql = tmp_sql + " /*AND   SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/*/ ";
                tmp_sql = tmp_sql + " AND   SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/ ";
                tmp_sql = tmp_sql + " /*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg end*/ ";
                tmp_sql = tmp_sql + " AND   EXISTS					/*該当年月が適用開始年月以降*/ ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM bk_kanri AS BKK ";
                tmp_sql = tmp_sql + " 			WHERE SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 			AND   SQ.gt_ym >= BKK.start_ym ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " /*20160603 ユーザーデータ検証による修正 add sta*/ ";
                tmp_sql = tmp_sql + " /*送金確定されているデータは別で抽出するため重複回避のため送金確定日が未設定のデータを抽出する*/ ";
                tmp_sql = tmp_sql + " AND   so_kakuymd IS NULL ";
                tmp_sql = tmp_sql + " /*20160603 ユーザーデータ検証による修正 add end*/ ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " /****************************************************************************** ";
                tmp_sql = tmp_sql + " 部分入金の移行 ";
                tmp_sql = tmp_sql + "   ・部分入金されている入金項目のうち、未入金部分は請求データ構築で ";
                tmp_sql = tmp_sql + "     作成されないため移行する必要がある ";
                tmp_sql = tmp_sql + "   ・移行条件 ";
                tmp_sql = tmp_sql + "       ・その他請求項目以外 (データ区分 <> 6) ";
                tmp_sql = tmp_sql + "       ・コンバート対象データの該当月が ";
                tmp_sql = tmp_sql + "         「運用開始年月」以降、かつ該当物件の「送金ルール開始年月」以降 ";
                tmp_sql = tmp_sql + " ******************************************************************************/ ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 SQ.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 	,SQ.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 	,SQ.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/*,SQ.ko_no AS [契約レコードNo]*/ ";
                tmp_sql = tmp_sql + " 	,KYK.[契約レコードNo] ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	,gt_ym AS [該当年月] ";
                tmp_sql = tmp_sql + " 	/*,nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 	,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 	,SQ.nkin_kbn AS [入金項目区分]	 ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目表示用名称] ";
                tmp_sql = tmp_sql + " 	/*,nkbn_yotei AS [(入金予定)入金区分]*/ ";
                tmp_sql = tmp_sql + " 	,nkbn_name  AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 	,SQ.sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(sq_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [請求税区分]	 ";
                tmp_sql = tmp_sql + " 	,sq_zeigak AS [請求税額] ";
                tmp_sql = tmp_sql + " 	,sq_simeymd AS [請求締切日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [複数請求先フラグ] ";
                tmp_sql = tmp_sql + " 	,100 AS [請求先区分]	/*契約者固定*/ ";
                tmp_sql = tmp_sql + " 	,SQ.kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 	,SQ.fkom_no AS [(入金予定)家賃振込先口座No] ";
                tmp_sql = tmp_sql + " 	,fkae_sqno AS [口座振替履歴No]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [口座振替履歴明細No] ";
                tmp_sql = tmp_sql + " 	/*20160530 送金データ調査での不足箇所修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/* ";
                tmp_sql = tmp_sql + " 	,fkae_sqflg AS [振替請求対象フラグ 1:請求中　2:未請求]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,fkae_sqtaiflg AS [振替請求対象フラグ　1:対象　2:対象外]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,fkae_sqflg AS [振替請求中フラグ]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [振替請求対象フラグ 1:請求中　2:未請求] ";
                tmp_sql = tmp_sql + " 	,1 AS [振替請求対象フラグ　1:対象　2:対象外] ";
                tmp_sql = tmp_sql + " 	,0 AS [振替請求中フラグ] ";
                tmp_sql = tmp_sql + " 	/*20160530 送金データ調査での不足箇所修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN data_kbn = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		WHEN data_kbn >= 2 THEN 1 ";
                tmp_sql = tmp_sql + " 	 END AS [月区分] ";
                tmp_sql = tmp_sql + " 	,1 AS [入金項目明細No] ";
                tmp_sql = tmp_sql + " 	/*,NULL AS [契約入金項目GUID]	契約入金データから取得する*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [次回契約情報フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費請求GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費検針日] ";
                tmp_sql = tmp_sql + " 	,SQ.so_gak AS [送金額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [送金税区分]	 ";
                tmp_sql = tmp_sql + " 	,so_zeigak AS [送金税額] ";
                tmp_sql = tmp_sql + " 	,0 AS [その他請求フラグ] ";
                tmp_sql = tmp_sql + " 	,0 AS [その他請求の送金率]	 ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求の指定送金日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [分割GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先区分] ";
                tmp_sql = tmp_sql + " 	/*'20160530 送金データ調査での不足箇所修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/* ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先口座No] ";
                tmp_sql = tmp_sql + " 	*/ ";
                tmp_sql = tmp_sql + " 	,so_no AS [支払・返金先No] ";
                tmp_sql = tmp_sql + " 	,sokoza_no AS [支払・返金先口座No] ";
                tmp_sql = tmp_sql + " 	/*'20160530 送金データ調査での不足箇所修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支店No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [担当者No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除請求項目GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [固定控除ルールGUID(請求)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除ルールフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求書発行予定日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求ソートNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預り設定時の預り予定GUID（預り金データ）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [敷金保証金随時処理GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [運用開始時の未納・滞納金GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金回収GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [随時修繕No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕請求No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [NULL] ";
                tmp_sql = tmp_sql + " 	,NULL AS [編集フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [更新ロックフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [備考] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳税額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [取込日（インポート実施日）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納サービスNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [顧客番号] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納請求フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求管理手数料区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [未収仕訳フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [雑収入区分] ";
                tmp_sql = tmp_sql + " 	,so_kakuymd AS [送金確定日]		/*10送金情報作成用*/ ";
                tmp_sql = tmp_sql + " 	,tnh_umu AS [滞納保証フラグ]	/*10送金情報作成用*/ ";
                tmp_sql = tmp_sql + " 	,'部分入金' AS [請求区分]		/*作業用*/ ";
                tmp_sql = tmp_sql + " FROM sq_meisai AS SQ ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,start_ym,end_ym,tnh_umu FROM bk_kanri) AS BKK ON SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " /*20160603 ユーザーデータ検証による修正 add sta*/ ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,hy_no,ky_no,ko_no,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約レコードNo] FROM ky_kosinkai) AS KYK ";
                tmp_sql = tmp_sql + " ON  SQ.bk_no = KYK.bk_no AND SQ.hy_no = KYK.hy_no AND SQ.ky_no = KYK.ky_no AND SQ.ko_no = KYK.ko_no ";
                tmp_sql = tmp_sql + " /*20160603 ユーザーデータ検証による修正 add end*/ ";
                tmp_sql = tmp_sql + " WHERE rec_kbn <> 0 ";
                tmp_sql = tmp_sql + " AND   nkin_ymd IS NULL ";
                tmp_sql = tmp_sql + " AND   data_kbn <> 6 ";
                tmp_sql = tmp_sql + " /*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg sta*/ ";
                tmp_sql = tmp_sql + " /*AND   SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/*/ ";
                tmp_sql = tmp_sql + " AND   SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/ ";
                tmp_sql = tmp_sql + " /*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg end*/ ";
                tmp_sql = tmp_sql + " AND   EXISTS					/*該当年月が適用開始年月以降*/ ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM bk_kanri AS BKK ";
                tmp_sql = tmp_sql + " 			WHERE SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 			AND   SQ.gt_ym >= BKK.start_ym ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " AND   so_kakuymd IS NULL ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " /****************************************************************************** ";
                tmp_sql = tmp_sql + " 未入金かつ送金確定されているデータの移行 ";
                tmp_sql = tmp_sql + "   送金確定されているデータを抽出する ";
                tmp_sql = tmp_sql + "   ・移行条件 ";
                tmp_sql = tmp_sql + "       ・未入金 ";
                tmp_sql = tmp_sql + "       ・部分入金以外 (部分入金で入金されていないデータは抽出済み) ";
                tmp_sql = tmp_sql + "       ・その他請求項目とその他請求以外の項目で分けて抽出し、結合する ";
                tmp_sql = tmp_sql + "         「運用開始年月」以降、かつ該当物件の「送金ルール開始年月」以降 ";
                tmp_sql = tmp_sql + " ******************************************************************************/ ";
                tmp_sql = tmp_sql + " /*************************** ";
                tmp_sql = tmp_sql + " その他請求以外の送金確定 ";
                tmp_sql = tmp_sql + " ***************************/ ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 SQ.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 	,SQ.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 	,SQ.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/*,SQ.ko_no AS [契約レコードNo]*/ ";
                tmp_sql = tmp_sql + " 	,KYK.[契約レコードNo] ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 chg end*/	 ";
                tmp_sql = tmp_sql + " 	,gt_ym AS [該当年月] ";
                tmp_sql = tmp_sql + " 	/*,nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 	,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 	,SQ.nkin_kbn AS [入金項目区分]	 ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目表示用名称] ";
                tmp_sql = tmp_sql + " 	/*,nkbn_yotei AS [(入金予定)入金区分]*/ ";
                tmp_sql = tmp_sql + " 	,nkbn_name  AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 	,SQ.sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(sq_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [請求税区分]	 ";
                tmp_sql = tmp_sql + " 	,sq_zeigak AS [請求税額] ";
                tmp_sql = tmp_sql + " 	,sq_simeymd AS [請求締切日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [複数請求先フラグ] ";
                tmp_sql = tmp_sql + " 	,100 AS [請求先区分]	/*契約者固定*/ ";
                tmp_sql = tmp_sql + " 	,SQ.kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 	,SQ.fkom_no AS [(入金予定)家賃振込先口座No] ";
                tmp_sql = tmp_sql + " 	,fkae_sqno AS [口座振替履歴No]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [口座振替履歴明細No] ";
                tmp_sql = tmp_sql + " 	/*20160530 送金データ調査での不足箇所修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/* ";
                tmp_sql = tmp_sql + " 	,fkae_sqflg AS [振替請求対象フラグ 1:請求中　2:未請求]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,fkae_sqtaiflg AS [振替請求対象フラグ　1:対象　2:対象外]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,fkae_sqflg AS [振替請求中フラグ]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [振替請求対象フラグ 1:請求中　2:未請求] ";
                tmp_sql = tmp_sql + " 	,1 AS [振替請求対象フラグ　1:対象　2:対象外] ";
                tmp_sql = tmp_sql + " 	,0 AS [振替請求中フラグ] ";
                tmp_sql = tmp_sql + " 	/*20160530 送金データ調査での不足箇所修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN data_kbn = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		WHEN data_kbn >= 2 THEN 1 ";
                tmp_sql = tmp_sql + " 	 END AS [月区分] ";
                tmp_sql = tmp_sql + " 	,1 AS [入金項目明細No] ";
                tmp_sql = tmp_sql + " 	/*,NULL AS [契約入金項目GUID]	契約入金データから取得する*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [次回契約情報フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費請求GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費検針日] ";
                tmp_sql = tmp_sql + " 	,SQ.so_gak AS [送金額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [送金税区分]	 ";
                tmp_sql = tmp_sql + " 	,so_zeigak AS [送金税額] ";
                tmp_sql = tmp_sql + " 	,0 AS [その他請求フラグ] ";
                tmp_sql = tmp_sql + " 	,0 AS [その他請求の送金率]	 ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求の指定送金日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [分割GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先区分] ";
                tmp_sql = tmp_sql + " 	/*'20160530 送金データ調査での不足箇所修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 	/* ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先口座No] ";
                tmp_sql = tmp_sql + " 	*/ ";
                tmp_sql = tmp_sql + " 	,so_no AS [支払・返金先No] ";
                tmp_sql = tmp_sql + " 	,sokoza_no AS [支払・返金先口座No] ";
                tmp_sql = tmp_sql + " 	/*'20160530 送金データ調査での不足箇所修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支店No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [担当者No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除請求項目GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [固定控除ルールGUID(請求)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除ルールフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求書発行予定日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求ソートNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預り設定時の預り予定GUID（預り金データ）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [敷金保証金随時処理GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [運用開始時の未納・滞納金GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金回収GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [随時修繕No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕請求No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [NULL] ";
                tmp_sql = tmp_sql + " 	,NULL AS [編集フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [更新ロックフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [備考] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳税額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [取込日（インポート実施日）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納サービスNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [顧客番号] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納請求フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求管理手数料区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [未収仕訳フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [雑収入区分] ";
                tmp_sql = tmp_sql + " 	,so_kakuymd AS [送金確定日]		/*10送金情報作成用*/ ";
                tmp_sql = tmp_sql + " 	,tnh_umu AS [滞納保証フラグ]	/*10送金情報作成用*/ ";
                tmp_sql = tmp_sql + " 	,'送金確定_その他請求以外' AS [請求区分]		/*作業用*/ ";
                tmp_sql = tmp_sql + " FROM sq_meisai AS SQ ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,start_ym,end_ym,tnh_umu FROM bk_kanri) AS BKK ON SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " /*20160603 ユーザーデータ検証による修正 add sta*/ ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,hy_no,ky_no,ko_no,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約レコードNo] FROM ky_kosinkai) AS KYK ";
                tmp_sql = tmp_sql + " ON  SQ.bk_no = KYK.bk_no AND SQ.hy_no = KYK.hy_no AND SQ.ky_no = KYK.ky_no AND SQ.ko_no = KYK.ko_no ";
                tmp_sql = tmp_sql + " /*20160603 ユーザーデータ検証による修正 add end*/ ";
                tmp_sql = tmp_sql + " WHERE nkin_ymd IS NULL ";
                tmp_sql = tmp_sql + " AND   data_kbn <> 6 ";
                tmp_sql = tmp_sql + " /*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg sta*/ ";
                tmp_sql = tmp_sql + " /*AND   SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/*/ ";
                tmp_sql = tmp_sql + " AND   SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/ ";
                tmp_sql = tmp_sql + " /*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg end*/ ";
                tmp_sql = tmp_sql + " AND   EXISTS					/*該当年月が適用開始年月以降*/ ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM bk_kanri AS BKK ";
                tmp_sql = tmp_sql + " 			WHERE SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 			AND   SQ.gt_ym >= BKK.start_ym ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " AND   so_kakuymd IS NOT NULL ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " /*********************** ";
                tmp_sql = tmp_sql + " その他請求の送金確定 20160603 ユーザーデータ検証による修正 add ";
                tmp_sql = tmp_sql + " ***********************/ ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 SQ.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 	,SQ.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 	,SQ.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 	,KYK.[契約レコードNo] ";
                tmp_sql = tmp_sql + " 	,gt_ym AS [該当年月] ";
                tmp_sql = tmp_sql + " 	/*,nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 	,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目区分]		/*10で請求データ作成時にNULLになることを確認*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目表示用名称] ";
                tmp_sql = tmp_sql + " 	/*,nkbn_yotei AS [(入金予定)入金区分]*/ ";
                tmp_sql = tmp_sql + " 	,nkbn_name  AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 	,SQ.sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(sq_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [請求税区分]	 ";
                tmp_sql = tmp_sql + " 	,sq_zeigak AS [請求税額] ";
                tmp_sql = tmp_sql + " 	,sq_simeymd AS [請求締切日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [複数請求先フラグ] ";
                tmp_sql = tmp_sql + " 	,100 AS [請求先区分]	/*契約者固定*/ ";
                tmp_sql = tmp_sql + " 	,SQ.kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 	,SQ.fkom_no AS [(入金予定)家賃振込先口座No] ";
                tmp_sql = tmp_sql + " 	,fkae_sqno AS [口座振替履歴No]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [口座振替履歴明細No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [振替請求対象フラグ 1:請求中　2:未請求] ";
                tmp_sql = tmp_sql + " 	,1 AS [振替請求対象フラグ　1:対象　2:対象外] ";
                tmp_sql = tmp_sql + " 	,0 AS [振替請求中フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [月区分]			/*10で請求データ作成時にNULLになることを確認*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目明細No]	/*10で請求データ作成時にNULLになることを確認*/ ";
                tmp_sql = tmp_sql + " 	/*,NULL AS [契約入金項目GUID]	契約入金データから取得する*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [次回契約情報フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費請求GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費検針日] ";
                tmp_sql = tmp_sql + " 	,SQ.so_gak AS [送金額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		ELSE 1 ";
                tmp_sql = tmp_sql + " 	 END AS [送金税区分]	 ";
                tmp_sql = tmp_sql + " 	,so_zeigak AS [送金税額] ";
                tmp_sql = tmp_sql + " 	,1 AS [その他請求フラグ] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(SQ.sq_gak,0) <> 0 THEN (SQ.so_gak / SQ.sq_gak) * 100 ";
                tmp_sql = tmp_sql + " 		ELSE 0 ";
                tmp_sql = tmp_sql + " 	 END AS [その他請求の送金率]	 ";
                tmp_sql = tmp_sql + " 	,so_yoteiymd AS [その他請求の指定送金日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [分割GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先区分] ";
                tmp_sql = tmp_sql + " 	,so_no AS [支払・返金先No] ";
                tmp_sql = tmp_sql + " 	,sokoza_no AS [支払・返金先口座No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支店No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [担当者No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除請求項目GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [固定控除ルールGUID(請求)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除ルールフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求書発行予定日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求ソートNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預り設定時の預り予定GUID（預り金データ）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [敷金保証金随時処理GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [運用開始時の未納・滞納金GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金回収GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [随時修繕No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕請求No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [適用税率] ";
                tmp_sql = tmp_sql + " 	,NULL AS [編集フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [更新ロックフラグ] ";
                tmp_sql = tmp_sql + " 	/*20160608 対応依頼分 -chg sta*/ ";
                tmp_sql = tmp_sql + " 	/*,NULL AS [備考]*/ ";
                tmp_sql = tmp_sql + " 	,'※送金済みデータ' AS [備考] ";
                tmp_sql = tmp_sql + " 	/*20160608 対応依頼分 -chg end*/ ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳税額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [取込日（インポート実施日）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納サービスNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [顧客番号] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納請求フラグ] ";
                tmp_sql = tmp_sql + " 	,2 AS [その他請求管理手数料区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [未収仕訳フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [雑収入区分] ";
                tmp_sql = tmp_sql + " 	,so_kakuymd AS [送金確定日]		/*10送金情報作成用*/ ";
                tmp_sql = tmp_sql + " 	,tnh_umu AS [滞納保証フラグ]	/*10送金情報作成用*/ ";
                tmp_sql = tmp_sql + " 	,'送金確定_その他請求' AS [請求区分]		/*作業用*/ ";
                tmp_sql = tmp_sql + " FROM sq_meisai AS SQ ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,start_ym,end_ym,tnh_umu FROM bk_kanri) AS BKK ON SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " LEFT JOIN (SELECT bk_no,hy_no,ky_no,ko_no,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約レコードNo] FROM ky_kosinkai) AS KYK ";
                tmp_sql = tmp_sql + " ON  SQ.bk_no = KYK.bk_no AND SQ.hy_no = KYK.hy_no AND SQ.ky_no = KYK.ky_no AND SQ.ko_no = KYK.ko_no ";
                tmp_sql = tmp_sql + " WHERE nkin_ymd IS NULL ";
                tmp_sql = tmp_sql + " AND   data_kbn = 6 ";
                tmp_sql = tmp_sql + " AND   SQ.nkin_no <> 6020		/*入金項目No:6020 は過剰金*/ ";
                tmp_sql = tmp_sql + " /*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg sta*/ ";
                tmp_sql = tmp_sql + " /*AND   SQ.gt_ym >= '運用開始年月置換用文字列'	/*該当年月が運用開始年月以降*/*/ ";
                tmp_sql = tmp_sql + " AND   SQ.sq_simeymd >= '運用開始年月置換用文字列'	/*請求締年月が運用開始年月以降*/ ";
                tmp_sql = tmp_sql + " /*20160610 運用開始年月を基準にした判別用年月を該当年月から請求締年月へ変更 chg end*/ ";
                tmp_sql = tmp_sql + " AND   EXISTS					/*該当年月が適用開始年月以降*/ ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM bk_kanri AS BKK ";
                tmp_sql = tmp_sql + " 			WHERE SQ.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 			AND   SQ.gt_ym >= BKK.start_ym ";
                tmp_sql = tmp_sql + " 		) ";
                tmp_sql = tmp_sql + " AND   so_kakuymd IS NOT NULL ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " /***************************************************************** ";
                tmp_sql = tmp_sql + " 控除支払登録データから請求のフラグが立っているデータを抽出 ";
                tmp_sql = tmp_sql + " *****************************************************************/ ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [部屋No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [契約No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [契約レコードNo]	 ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(gt_ym,'') <> '' THEN gt_ym ";
                tmp_sql = tmp_sql + " 		ELSE DATEADD(DAY,(DAY(so_ymd) - 1 ) * (-1),so_ymd) ";
                tmp_sql = tmp_sql + " 	 END AS [該当年月] ";
                tmp_sql = tmp_sql + " 	/*,nkin_no AS [入金項目名]*/ ";
                tmp_sql = tmp_sql + " 	,MN.nkin_name AS [入金項目名]	 ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目区分] ";
                tmp_sql = tmp_sql + " 	,SHK.nkin_name AS [入金項目表示用名称] ";
                tmp_sql = tmp_sql + " 	,'振込' AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 	,kj_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN kj_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 		WHEN kj_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 	 END AS [請求税区分]	 ";
                tmp_sql = tmp_sql + " 	,kj_zeigak AS [請求税額] ";
                tmp_sql = tmp_sql + " 	,DATEADD(DAY,-1,DATEADD(DAY,(DAY(so_ymd) - 1 ) * (-1),so_ymd)) AS [請求締切日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [複数請求先フラグ] ";
                tmp_sql = tmp_sql + " 	,200 AS [請求先区分] ";
                tmp_sql = tmp_sql + " 	,so_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [(入金予定)家賃振込先口座No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [口座振替履歴No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [口座振替履歴明細No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [振替請求対象フラグ 1:請求中　2:未請求] ";
                tmp_sql = tmp_sql + " 	,1 AS [振替請求対象フラグ　1:対象　2:対象外] ";
                tmp_sql = tmp_sql + " 	,0 AS [振替請求中フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [月区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目明細No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [次回契約情報フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費請求GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費検針日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金税区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金税額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求の送金率] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求の指定送金日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [分割GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払・返金先区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払返金先No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支払返金先口座No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] ";
                tmp_sql = tmp_sql + " 	,NULL AS [支店No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [担当者No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除請求項目GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [固定控除ルールGUID(請求)] ";
                tmp_sql = tmp_sql + " 	,0 AS [控除ルールフラグ] ";
                tmp_sql = tmp_sql + " 	,so_ymd AS [請求書発行予定日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求ソートNo] ";
                tmp_sql = tmp_sql + " 	,2 AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預り設定時の預り予定GUID（預り金データ）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [敷金保証金随時処理GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [運用開始時の未納・滞納金GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金回収GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [随時修繕No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕請求No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [適用税率] ";
                tmp_sql = tmp_sql + " 	,NULL AS [編集フラグ] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN inpu_umu = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		WHEN inpu_umu = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 	 END AS [更新ロックフラグ]	 ";
                tmp_sql = tmp_sql + " 	,biko AS [備考] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳税額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [取込日（インポート実施日）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納サービスNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [顧客番号] ";
                tmp_sql = tmp_sql + " 	,0 AS [コンビニ収納請求フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求管理手数料区分] ";
                tmp_sql = tmp_sql + " 	,0 AS [未収仕訳フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [雑収入区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金確定日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [滞納保証フラグ] ";
                tmp_sql = tmp_sql + " 	,'控除支払の請求' AS [請求区分]		/*作業用*/ ";
                tmp_sql = tmp_sql + " 	FROM sh_kojo AS SHK ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MN ON SHK.nkin_no = MN.nkin_no ";
                tmp_sql = tmp_sql + " WHERE SHK.sq_umu = 1 ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " /***************************************************************** ";
                tmp_sql = tmp_sql + " 控除支払登録データから支払に金額が設定されているデータの抽出 ";
                tmp_sql = tmp_sql + " →自社支払いへ移行する ";
                tmp_sql = tmp_sql + " *****************************************************************/ ";
                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [部屋No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [契約No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [契約レコードNo]	 ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN ISNULL(gt_ym,'') <> '' THEN gt_ym ";
                tmp_sql = tmp_sql + " 		ELSE DATEADD(DAY,(DAY(so_ymd) - 1 ) * (-1),so_ymd) ";
                tmp_sql = tmp_sql + " 	 END AS [該当年月] ";
                tmp_sql = tmp_sql + " 	/*,nkin_no AS [入金項目名]*/ ";
                tmp_sql = tmp_sql + " 	,MN.nkin_name AS [入金項目名]	 ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目表示用名称] ";
                tmp_sql = tmp_sql + " 	,NULL AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 	,siharai_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN siharai_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 		WHEN siharai_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 	 END AS [請求税区分]	 ";
                tmp_sql = tmp_sql + " 	,siharai_zeigak AS [請求税額]	 ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求締切日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [複数請求先フラグ] ";
                tmp_sql = tmp_sql + " 	,900 AS [請求先区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求先No] ";
                tmp_sql = tmp_sql + " 	,''AS [(入金予定)家賃振込先口座No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [口座振替履歴No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [口座振替履歴明細No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [振替請求対象フラグ 1:請求中　2:未請求] ";
                tmp_sql = tmp_sql + " 	,1 AS [振替請求対象フラグ　1:対象　2:対象外] ";
                tmp_sql = tmp_sql + " 	,0 AS [振替請求中フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [月区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [入金項目明細No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [次回契約情報フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費請求GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [変動費検針日] ";
                tmp_sql = tmp_sql + " 	,siharai_gak AS [送金額] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN siharai_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 		WHEN siharai_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 	 END AS [送金税区分]	 ";
                tmp_sql = tmp_sql + " 	,siharai_zeigak AS [送金税額]	 ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求の送金率] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求の指定送金日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [分割GUID] ";
                tmp_sql = tmp_sql + " 	,1 AS [支払・返金区分] ";
                tmp_sql = tmp_sql + " 	,200 AS [支払・返金先区分] ";
                tmp_sql = tmp_sql + " 	,so_no AS [支払返金先No] ";
                tmp_sql = tmp_sql + " 	,sokoza_no AS [支払返金先口座No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] ";
                tmp_sql = tmp_sql + " 	,1 AS [支店No] ";
                tmp_sql = tmp_sql + " 	,99001 AS [担当者No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [控除請求項目GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [固定控除ルールGUID(請求)] ";
                tmp_sql = tmp_sql + " 	,0 AS [控除ルールフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求書発行予定日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [請求ソートNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替・預り設定時の預り予定GUID（預り金データ）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [敷金保証金随時処理GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [運用開始時の未納・滞納金GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [送金回収GUID] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕区分] ";
                tmp_sql = tmp_sql + " 	,NULL AS [随時修繕No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [修繕請求No] ";
                tmp_sql = tmp_sql + " 	,NULL AS [適用税率] ";
                tmp_sql = tmp_sql + " 	,NULL AS [編集フラグ] ";
                tmp_sql = tmp_sql + " 	,CASE ";
                tmp_sql = tmp_sql + " 		WHEN inpu_umu = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 		WHEN inpu_umu = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 	 END AS [更新ロックフラグ]	 ";
                tmp_sql = tmp_sql + " 	,biko AS [備考] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [立替仕訳税額] ";
                tmp_sql = tmp_sql + " 	,NULL AS [取込日（インポート実施日）] ";
                tmp_sql = tmp_sql + " 	,NULL AS [コンビニ収納サービスNo] ";
                tmp_sql = tmp_sql + " 	,NULL AS [顧客番号] ";
                tmp_sql = tmp_sql + " 	,0 AS [コンビニ収納請求フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [その他請求管理手数料区分] ";
                tmp_sql = tmp_sql + " 	,0 AS [未収仕訳フラグ] ";
                tmp_sql = tmp_sql + " 	,NULL AS [雑収入区分] ";
                tmp_sql = tmp_sql + " 	,so_ymd AS [送金確定日] ";
                tmp_sql = tmp_sql + " 	,NULL AS [滞納保証フラグ] ";
                tmp_sql = tmp_sql + " 	,'自社支払' AS [請求区分]		/*作業用*/ ";
                tmp_sql = tmp_sql + " 	FROM sh_kojo AS SHK ";
                tmp_sql = tmp_sql + " LEFT JOIN m_nkin AS MN ON SHK.nkin_no = MN.nkin_no ";
                tmp_sql = tmp_sql + " WHERE siharai_gak <> 0 AND SHK.nkin_no <> 9910 ";
                // 20160601 請求データの抽出クエリ修正 -chg end

                // 20160525 運用開始年月の追加 -add
                tmp_sql = tmp_sql.Replace("運用開始年月置換用文字列", CommonModule.UnyoYMD);

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

            }

            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {
                return default;

            }

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

            }

        }

    }

    #endregion

    #region 送金確定情報

    // 請求データ抽出時に送金確定データを抽出→請求データ移行時に送金確定データも移行するためここでは抽出しない

    // Public Class Sq_sokin_Repository

    // Public Class SubConv
    // Implements IConv

    // ''' <summary>
    // ''' 【抽出クエリ】
    // ''' </summary>
    // ''' <returns></returns>
    // ''' <remarks></remarks>
    // Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

    // Dim tmp_sql As String = ""

    // sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[送金レコードNo]"

    // tmp_sql = tmp_sql & " /****************************************************************************** "
    // tmp_sql = tmp_sql & " 送金予定情報の移行 "
    // tmp_sql = tmp_sql & "   移行対象となる請求データ作成時に抽出したデータのうち、送金実績があるデータを "
    // tmp_sql = tmp_sql & "   抽出し送金予定データへ移行する "
    // tmp_sql = tmp_sql & " ******************************************************************************/ "
    // tmp_sql = tmp_sql & "  "
    // tmp_sql = tmp_sql & " /********************************** "
    // tmp_sql = tmp_sql & " その他請求 "
    // tmp_sql = tmp_sql & " **********************************/ "
    // tmp_sql = tmp_sql & " SELECT "
    // tmp_sql = tmp_sql & " 	  SQ.bk_no AS [物件No] "
    // tmp_sql = tmp_sql & " 	 ,hy_no AS [部屋No] "
    // tmp_sql = tmp_sql & " 	 ,ky_no AS [契約No] "
    // tmp_sql = tmp_sql & " 	 ,ko_no AS [契約管理レコードNo] "
    // tmp_sql = tmp_sql & " 	,1 AS [送金レコードNo] "
    // tmp_sql = tmp_sql & " 	,so_gak AS [送金額] "
    // tmp_sql = tmp_sql & " 	,CASE "
    // tmp_sql = tmp_sql & " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 1 "
    // tmp_sql = tmp_sql & " 		ELSE 0 "
    // tmp_sql = tmp_sql & " 	 END AS [送金税フラグ] "
    // tmp_sql = tmp_sql & " 	,so_zeigak AS [送金税額] "
    // tmp_sql = tmp_sql & " 	,'' AS [編集フラグ] "
    // tmp_sql = tmp_sql & " 	,1 AS [更新ロックフラグ] "
    // tmp_sql = tmp_sql & " 	,tnh_umu AS [滞納保証フラグ] "
    // tmp_sql = tmp_sql & " 	,so_yoteiymd AS [送金予定日] "
    // tmp_sql = tmp_sql & " 	,CASE "
    // tmp_sql = tmp_sql & " 		WHEN ISNULL(sq_gak,0) <> 0 THEN (so_gak / sq_gak) * 100 "
    // tmp_sql = tmp_sql & " 		ELSE 0 "
    // tmp_sql = tmp_sql & " 	 END AS [送金率] "
    // tmp_sql = tmp_sql & " 	,so_gak AS [送金額計算対象請求額] "
    // tmp_sql = tmp_sql & " 	,'' AS [送金回収区分] "
    // tmp_sql = tmp_sql & " 	,'' AS [仕訳フラグ] "
    // tmp_sql = tmp_sql & " FROM sq_meisai AS SQ "
    // tmp_sql = tmp_sql & " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no "
    // tmp_sql = tmp_sql & " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no "
    // tmp_sql = tmp_sql & " LEFT JOIN bk_kanri AS BKK ON SQ.bk_no = BKK.bk_no "
    // tmp_sql = tmp_sql & " WHERE nkin_ymd IS NULL "
    // tmp_sql = tmp_sql & " AND   so_kakuymd IS NOT NULL "
    // tmp_sql = tmp_sql & " AND   data_kbn = 6 "
    // tmp_sql = tmp_sql & " AND   SQ.nkin_no <> 6020		/*入金項目No:6020 は過剰金*/ "
    // tmp_sql = tmp_sql & " AND   SQ.so_kakuymd IS NOT NULL "
    // tmp_sql = tmp_sql & " AND   SQ.gt_ym >= '20160401'	/*該当年月が運用開始年月以降*/ "
    // tmp_sql = tmp_sql & " AND   EXISTS					/*該当年月が適用開始年月以降*/ "
    // tmp_sql = tmp_sql & " 		( "
    // tmp_sql = tmp_sql & " 			SELECT * FROM bk_kanri AS BKK "
    // tmp_sql = tmp_sql & " 			WHERE SQ.bk_no = BKK.bk_no "
    // tmp_sql = tmp_sql & " 			AND   SQ.gt_ym >= BKK.start_ym "
    // tmp_sql = tmp_sql & " 		) "
    // tmp_sql = tmp_sql & "  "
    // tmp_sql = tmp_sql & " /********************************** "
    // tmp_sql = tmp_sql & " 部分入金 "
    // tmp_sql = tmp_sql & " **********************************/ "
    // tmp_sql = tmp_sql & " UNION "
    // tmp_sql = tmp_sql & " SELECT "
    // tmp_sql = tmp_sql & " 	  SQ.bk_no AS [物件No] "
    // tmp_sql = tmp_sql & " 	 ,hy_no AS [部屋No] "
    // tmp_sql = tmp_sql & " 	 ,ky_no AS [契約No] "
    // tmp_sql = tmp_sql & " 	 ,ko_no AS [契約管理レコードNo] "
    // tmp_sql = tmp_sql & " 	,1 AS [送金レコードNo] "
    // tmp_sql = tmp_sql & " 	,so_gak AS [送金額] "
    // tmp_sql = tmp_sql & " 	,CASE "
    // tmp_sql = tmp_sql & " 		WHEN ISNULL(so_zeigak,0) <> 0 THEN 1 "
    // tmp_sql = tmp_sql & " 		ELSE 0 "
    // tmp_sql = tmp_sql & " 	 END AS [送金税フラグ] "
    // tmp_sql = tmp_sql & " 	,so_zeigak AS [送金税額] "
    // tmp_sql = tmp_sql & " 	,'' AS [編集フラグ] "
    // tmp_sql = tmp_sql & " 	,1 AS [更新ロックフラグ] "
    // tmp_sql = tmp_sql & " 	,tnh_umu AS [滞納保証フラグ] "
    // tmp_sql = tmp_sql & " 	,so_yoteiymd AS [送金予定日] "
    // tmp_sql = tmp_sql & " 	,CASE "
    // tmp_sql = tmp_sql & " 		WHEN ISNULL(sq_gak,0) <> 0 THEN (so_gak / sq_gak) * 100 "
    // tmp_sql = tmp_sql & " 		ELSE 0 "
    // tmp_sql = tmp_sql & " 	 END AS [送金率] "
    // tmp_sql = tmp_sql & " 	,so_gak AS [送金額計算対象請求額] "
    // tmp_sql = tmp_sql & " 	,'' AS [送金回収区分] "
    // tmp_sql = tmp_sql & " 	,'' AS [仕訳フラグ] "
    // tmp_sql = tmp_sql & " FROM sq_meisai AS SQ "
    // tmp_sql = tmp_sql & " LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no "
    // tmp_sql = tmp_sql & " LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no "
    // tmp_sql = tmp_sql & " LEFT JOIN bk_kanri AS BKK ON SQ.bk_no = BKK.bk_no "
    // tmp_sql = tmp_sql & " WHERE rec_kbn <> 0 "
    // tmp_sql = tmp_sql & " AND   nkin_ymd IS NULL "
    // tmp_sql = tmp_sql & " AND   so_kakuymd IS NOT NULL "
    // tmp_sql = tmp_sql & " AND   data_kbn <> 6 "
    // tmp_sql = tmp_sql & " AND   SQ.gt_ym >= '20160401'	/*該当年月が運用開始年月以降*/ "
    // tmp_sql = tmp_sql & " AND   EXISTS					/*該当年月が適用開始年月以降*/ "
    // tmp_sql = tmp_sql & " 		( "
    // tmp_sql = tmp_sql & " 			SELECT * FROM bk_kanri AS BKK "
    // tmp_sql = tmp_sql & " 			WHERE SQ.bk_no = BKK.bk_no "
    // tmp_sql = tmp_sql & " 			AND   SQ.gt_ym >= BKK.start_ym "
    // tmp_sql = tmp_sql & " 		) "

    // Return tmp_sql

    // End Function

    // Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As SafeDictionary<string, string>, ByRef hash_chkafter As SafeDictionary<string, string>, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

    // End Function

    // Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

    // End Sub

    // Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

    // End Function

    // Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

    // End Sub

    // End Class

    // End Class

    #endregion

    #region 変動費検針情報

    public class Sq_hendokensin_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【抽出クエリ】
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";

                // 20160603 ユーザーデータ検証による修正 -chg sta
                // sortstr = "[物件No],[部屋No],[契約No],[契約レコードNo],[変動区分],[最新データ取得用] DESC"
                sortstr = "[物件No],[部屋No],[契約No],[契約レコードNo],[最新データ取得用] DESC";
                // 20160603 ユーザーデータ検証による修正 -chg end

                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		/*20160603 ユーザーデータ検証による修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/* HENDO_TOTAL.**/ ";
                tmp_sql = tmp_sql + " 		 [物件No] ";
                tmp_sql = tmp_sql + " 		,[部屋No] ";
                tmp_sql = tmp_sql + " 		,[契約No] ";
                tmp_sql = tmp_sql + " 		,[契約レコードNo] ";
                tmp_sql = tmp_sql + " 		,[検針日] ";
                tmp_sql = tmp_sql + " 		,[該当年月] ";
                tmp_sql = tmp_sql + " 		,[最新データ取得用] ";
                tmp_sql = tmp_sql + " 		,[最古データ取得用] ";
                tmp_sql = tmp_sql + " 		,[検針パターン] ";
                tmp_sql = tmp_sql + " 		,[前回値] ";
                tmp_sql = tmp_sql + " 		,[今回値] ";
                tmp_sql = tmp_sql + " 		,[今回使用量] ";
                tmp_sql = tmp_sql + " 		,[検針料金1] ";
                tmp_sql = tmp_sql + " 		,[検針料金2] ";
                tmp_sql = tmp_sql + " 		,[検針料金3] ";
                tmp_sql = tmp_sql + " 		,[検針料金4] ";
                tmp_sql = tmp_sql + " 		,[検針料金5] ";
                tmp_sql = tmp_sql + " 		,[変動費検針情報備考] ";
                tmp_sql = tmp_sql + " 		/*20160603 ユーザーデータ検証による修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 		,9 AS [入金項目区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [入金項目表示用名称] ";
                tmp_sql = tmp_sql + " 		,nkbn_name AS [(入金予定)入金区分] ";
                tmp_sql = tmp_sql + " 		,SQ.sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(sq_zeigak,0) <> 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 0 ";
                tmp_sql = tmp_sql + " 		 END AS [請求税区分] ";
                tmp_sql = tmp_sql + " 		,sq_zeigak AS [請求税額] ";
                tmp_sql = tmp_sql + " 		,sq_simeymd AS [請求締切日] ";
                tmp_sql = tmp_sql + " 		,'' AS [複数請求先フラグ] ";
                tmp_sql = tmp_sql + " 		,100 AS [請求先区分]	/*契約者固定*/ ";
                tmp_sql = tmp_sql + " 		,SQ.kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 		,SQ.fkom_no AS [(入金予定)家賃振込先口座No] ";
                tmp_sql = tmp_sql + " 		,fkae_sqno AS [口座振替履歴No]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替履歴明細No] ";
                tmp_sql = tmp_sql + " 		,fkae_sqflg AS [振替請求対象フラグ 1:請求中　2:未請求]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 		,fkae_sqtaiflg AS [振替請求対象フラグ　1:対象　2:対象外]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 		,fkae_sqflg AS [振替請求中フラグ]	/*要確認*/ ";
                tmp_sql = tmp_sql + " 		,NULL AS [月区分]			/*10で請求データ作成時にNULLになることを確認*/ ";
                tmp_sql = tmp_sql + " 		,NULL AS [入金項目明細No]	/*10で請求データ作成時にNULLになることを確認*/ ";
                tmp_sql = tmp_sql + " 		/*,'' AS [契約入金項目GUID]	契約入金データから取得する*/ ";
                tmp_sql = tmp_sql + " 		,NULL AS [次回契約情報フラグ] ";
                tmp_sql = tmp_sql + " 		/*,NULL AS [変動費請求GUID]*/	/*プログラム内で作成*/ ";
                tmp_sql = tmp_sql + " 		/*20160603 ユーザーデータ検証による修正 -del*/ ";
                tmp_sql = tmp_sql + " 		/*,1 AS [変動費区分]*/ ";
                tmp_sql = tmp_sql + " 		,so_yoteiymd AS [送金予定日] ";
                tmp_sql = tmp_sql + " 		,SQ.so_gak AS [送金額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(so_zeigak,0) <> 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 0 ";
                tmp_sql = tmp_sql + " 		 END AS [送金税区分] ";
                tmp_sql = tmp_sql + " 		,so_zeigak AS [送金税額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(SQ.sq_gak,0) <> 0 THEN (SQ.so_gak / SQ.sq_gak) * 100 ";
                tmp_sql = tmp_sql + " 			ELSE 0 ";
                tmp_sql = tmp_sql + " 		 END AS [送金率] /*変動費検針データ用*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [その他請求フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [その他請求の送金率]	 ";
                tmp_sql = tmp_sql + " 		,'' AS [その他請求の指定送金日] ";
                tmp_sql = tmp_sql + " 		,'' AS [分割GUID] ";
                tmp_sql = tmp_sql + " 		,'' AS [支払・返金区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [支払・返金先区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [支払・返金先No] ";
                tmp_sql = tmp_sql + " 		,'' AS [支払・返金先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [一括借上額の部屋毎明細の入金項目番号(sorule_hyikkatugakのnkin_bango)] ";
                tmp_sql = tmp_sql + " 		,'' AS [支店No] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [控除請求項目GUID] ";
                tmp_sql = tmp_sql + " 		,'' AS [固定控除ルールGUID(請求)] ";
                tmp_sql = tmp_sql + " 		,'' AS [控除ルールフラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [請求書発行予定日] ";
                tmp_sql = tmp_sql + " 		,'' AS [請求ソートNo] ";
                tmp_sql = tmp_sql + " 		,'' AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [立替・預り設定時の預り予定GUID（預り金データ）] ";
                tmp_sql = tmp_sql + " 		,'' AS [敷金保証金随時処理GUID] ";
                tmp_sql = tmp_sql + " 		,'' AS [運用開始時の未納・滞納金GUID] ";
                tmp_sql = tmp_sql + " 		,'' AS [送金回収GUID] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [随時修繕No] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕請求No] ";
                tmp_sql = tmp_sql + " 		,'' AS [適用税率] ";
                tmp_sql = tmp_sql + " 		,'' AS [編集フラグ] ";
                tmp_sql = tmp_sql + " 		,1 AS [更新ロックフラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [立替仕訳額] ";
                tmp_sql = tmp_sql + " 		,'' AS [立替仕訳税額] ";
                tmp_sql = tmp_sql + " 		,'' AS [取込日（インポート実施日）] ";
                tmp_sql = tmp_sql + " 		,'' AS [コンビニ収納サービスNo] ";
                tmp_sql = tmp_sql + " 		,'' AS [顧客番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [コンビニ収納請求フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [その他請求管理手数料区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [未収仕訳フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [雑収入区分]	 ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM sq_meisai ";
                tmp_sql = tmp_sql + " 		WHERE data_kbn = 5 ";
                tmp_sql = tmp_sql + " 		AND   koteihendo_kbn = 2 ";
                tmp_sql = tmp_sql + " 	) AS SQ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MNKIN ON SQ.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkbn AS MNKBN ON SQ.nkbn_yotei = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " 	/*契約請求情報から変動費ルールNoを取得するため結合 sta*/ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN (SELECT * FROM ky_sqdata WHERE koteihendo_kbn = 2 AND tuki_kbn = 2) AS KYSQ ";
                tmp_sql = tmp_sql + " 	ON  SQ.bk_no = KYSQ.bk_no ";
                tmp_sql = tmp_sql + " 	AND SQ.hy_no = KYSQ.hy_no ";
                tmp_sql = tmp_sql + " 	AND SQ.ky_no = KYSQ.ky_no ";
                tmp_sql = tmp_sql + " 	AND SQ.ko_no = KYSQ.ko_no ";
                tmp_sql = tmp_sql + " 	AND SQ.nkin_no = KYSQ.nkin_no ";
                tmp_sql = tmp_sql + " 	/*契約請求情報から変動費ルールNoを取得するため結合 end*/ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 				,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 				,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 				,ko_no AS [契約レコードNo] ";
                tmp_sql = tmp_sql + " 				,hendo_rui AS [変動区分] ";
                tmp_sql = tmp_sql + " 				,kensin_ymd AS [検針日] ";
                tmp_sql = tmp_sql + " 				,gt_ym AS [該当年月] ";
                tmp_sql = tmp_sql + " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no,hendo_rui,hendo_no ORDER BY gt_ym DESC) AS [最新データ取得用] ";
                tmp_sql = tmp_sql + " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no,hendo_rui,hendo_no ORDER BY gt_ym) AS [最古データ取得用] ";
                tmp_sql = tmp_sql + " 				,1 AS [検針パターン]			/*デフォルトで各戸メーターを設定*/ ";
                tmp_sql = tmp_sql + " 				/*,'' AS [変動guid]*/ ";
                tmp_sql = tmp_sql + " 				/*,'' AS [変動請求guid]*/ ";
                tmp_sql = tmp_sql + " 				/*,kys_no AS [請求先No]*/		/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				/*,nkin_no AS [入金項目No]*/	/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				/*,nkbn_yotei AS [入金区分]*/	/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				/*,sq_simeymd AS [請求締切日]*/		/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				/*,kensin_ymd AS [前回検針日]*/		/*移行後に更新*/ ";
                tmp_sql = tmp_sql + " 				,zen_kensin AS [前回値] ";
                tmp_sql = tmp_sql + " 				/*,siyoryo AS [前回使用量]*/		/*移行後に更新*/ ";
                tmp_sql = tmp_sql + " 				/*,sq_zeigak AS [前回請求額]*/		/*移行後に更新*/ ";
                tmp_sql = tmp_sql + " 				,kon_kensin AS [今回値] ";
                tmp_sql = tmp_sql + " 				,siyoryo AS [今回使用量] ";
                tmp_sql = tmp_sql + " 				,ryokin1 AS [検針料金1] ";
                tmp_sql = tmp_sql + " 				,ryokin2 AS [検針料金2] ";
                tmp_sql = tmp_sql + " 				,ryokin3 AS [検針料金3] ";
                tmp_sql = tmp_sql + " 				,ryokin4 AS [検針料金4] ";
                tmp_sql = tmp_sql + " 				,ryokin5 AS [検針料金5] ";
                tmp_sql = tmp_sql + " 				/*,sq_gak AS [請求額]*/				/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				/*,sq_zeigak AS [請求税額]*/		/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				,biko AS [変動費検針情報備考] ";
                tmp_sql = tmp_sql + " 				/*,so_yoteiymd AS [送金予定日]*/	/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				/*,CASE ";
                tmp_sql = tmp_sql + " 					WHEN ISNULL(sq_gak,0) <> 0 THEN (so_gak / sq_gak) * 100		/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 					ELSE 0 ";
                tmp_sql = tmp_sql + " 				 END AS [送金率] ";
                tmp_sql = tmp_sql + " 				*/ ";
                tmp_sql = tmp_sql + " 				/*,so_gak AS [送金額]*/				/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				/*,so_zeigak AS [送金税額]*/		/*請求データから取得*/ ";
                tmp_sql = tmp_sql + " 				/*,'' AS [前回変動請求guid]*/		/*移行後に更新*/ ";
                tmp_sql = tmp_sql + " 				,hendo_no AS [結合用変動No] ";
                tmp_sql = tmp_sql + " 				/*20160603 ユーザーデータ検証による修正 add*/ ";
                tmp_sql = tmp_sql + " 				,nkin_no AS [結合用入金項目No]				 ";
                tmp_sql = tmp_sql + " 			FROM sq_hendomei ";
                tmp_sql = tmp_sql + " 		) AS HENDO_TOTAL ";
                tmp_sql = tmp_sql + " 	ON  SQ.bk_no = HENDO_TOTAL.物件No ";
                tmp_sql = tmp_sql + " 	AND SQ.hy_no = HENDO_TOTAL.部屋No ";
                tmp_sql = tmp_sql + " 	AND SQ.ky_no = HENDO_TOTAL.契約No ";
                tmp_sql = tmp_sql + " 	AND SQ.ko_no = HENDO_TOTAL.契約レコードNo ";
                tmp_sql = tmp_sql + " 	AND SQ.gt_ym = HENDO_TOTAL.該当年月 ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 add*/ ";
                tmp_sql = tmp_sql + " 	AND SQ.nkin_no = HENDO_TOTAL.結合用入金項目No ";
                tmp_sql = tmp_sql + " 	AND KYSQ.hendo_no = HENDO_TOTAL.結合用変動No ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 add sta*/ ";
                tmp_sql = tmp_sql + " 	/*請求情報には存在するが変動費明細情報に存在しない場合があるため除外*/ ";
                tmp_sql = tmp_sql + " 	WHERE HENDO_TOTAL.物件No IS NOT NULL ";
                tmp_sql = tmp_sql + " 	/*20160603 ユーザーデータ検証による修正 add end*/ ";
                tmp_sql = tmp_sql + " ) AS VW ";

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

            }

            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {
                return default;

            }

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

            }

        }

    }

    #endregion

    #region 家主固定控除情報

    public class Sq_koteikojo_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【抽出クエリ】
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";

                sortstr = "[物件No],[入金項目ソートNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[入金項目ソートNo],[入金項目名],[部屋No],[相殺フラグ],[額],[税区分],[税額],[入金区分],[家賃入金口座No],[適用開始該当年月],[適用終了該当年月],[立替・預りフラグ],[控除請求月区分],[控除請求発生パターン],[控除請求発生間隔],[控除請求発生月指定年区分],[控除請求発生月指定月区分],[適用税率],[備考],[請求締区分],[領収フラグ]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 BKKMEI.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		/*20160603 ユーザーデータ検証による修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,ROW_NUMBER()OVER(PARTITION BY BKKMEI.bk_no,BKKMEI.kjkn_no ORDER BY BKKMEI.nkin_no) AS [入金項目ソートNo]*/ ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY BKKMEI.bk_no ORDER BY BKKMEI.kjkn_no,BKKMEI.nkin_no) AS [入金項目ソートNo] ";
                tmp_sql = tmp_sql + " 		/*20160603 ユーザーデータ検証による修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		/*,BKKMEI.nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 		,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 		,'' AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,1 AS [相殺フラグ] ";
                tmp_sql = tmp_sql + " 		,kj_gak AS [額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kj_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN kj_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		 END AS [税区分] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kj_zeiumu = 0 THEN 0 ";
                tmp_sql = tmp_sql + " 			WHEN kj_zeiumu = 1 THEN kj_gak * 0.08 ";
                tmp_sql = tmp_sql + " 		 END AS [税額] ";
                tmp_sql = tmp_sql + " 		,'振込' AS [入金区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [家賃入金口座No] ";
                tmp_sql = tmp_sql + " 		,start_ym AS [適用開始該当年月] ";
                tmp_sql = tmp_sql + " 		/*20160603 ユーザーデータ検証による修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,end_ym AS [適用終了該当年月]*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN end_ym = '99991201' THEN '21001231' ";
                tmp_sql = tmp_sql + " 			ELSE end_ym ";
                tmp_sql = tmp_sql + " 		 END AS [適用終了該当年月] ";
                tmp_sql = tmp_sql + " 		/*20160603 ユーザーデータ検証による修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		,2 AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 		,3 AS [控除請求月区分]			/*V7には請求月を設定する箇所が無いのでデフォルトで(3:当月)を設定しておく*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 3 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 4 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 5 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 6 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 7 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 8 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 9 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 10 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 11 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm = 12 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN hasei_mm >= 13 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [控除請求発生パターン] ";
                tmp_sql = tmp_sql + " 		,hasei_mm AS [控除請求発生間隔] ";
                tmp_sql = tmp_sql + " 		,'' AS [控除請求発生月指定年区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [控除請求発生月指定月区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [適用税率] ";
                tmp_sql = tmp_sql + " 		,BKKMEI.biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [請求締区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [領収フラグ] ";
                tmp_sql = tmp_sql + " 	FROM bk_kojyomei AS BKKMEI ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bk_kojyo AS BKK ";
                tmp_sql = tmp_sql + " 	ON  BKKMEI.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 	AND BKKMEI.kjkn_no = BKK.kjkn_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN ON BKKMEI.nkin_no = MN.nkin_no ";
                tmp_sql = tmp_sql + " ) AS VW ";

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

            }

            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {
                return default;

            }

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

            }

        }

    }

    #endregion

    #region 家主請求控除情報

    public class Sq_sqkojo_Repository
    {

        public class SubConv : IConv
        {

            /// <summary>
            /// 【抽出クエリ】
            /// </summary>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry(ref string sortstr)
            {

                string tmp_sql = "";

                // 20160601 家主請求控除情報の取得方法を変更 -chg sta
                // 契約入金項目情報(ky_sqdata) → 控除支払情報(sh_kojo)へ変更
                // sortstr = "[物件No],[部屋No],[契約No],[契約レコードNo],[項目ソートNo]"

                // tmp_sql = tmp_sql & " SELECT * FROM "
                // tmp_sql = tmp_sql & " ( "
                // tmp_sql = tmp_sql & " 	SELECT DISTINCT "
                // tmp_sql = tmp_sql & " 		 KYKOJO.bk_no AS [物件No] "
                // tmp_sql = tmp_sql & " 		,KYKOJO.hy_no AS [部屋No] "
                // tmp_sql = tmp_sql & " 		,KYKOJO.ky_no AS [契約No] "
                // tmp_sql = tmp_sql & " 		,[契約レコードNo]	 "
                // tmp_sql = tmp_sql & " 		,so_ymd AS [控除予定日] "
                // tmp_sql = tmp_sql & " 		,naibu_no AS [項目ソートNo] "
                // tmp_sql = tmp_sql & " 		,DATEADD(DAY,(DAY(so_ymd) - 1 ) * (-1),so_ymd) AS [該当年月] "
                // tmp_sql = tmp_sql & " 		/*,KYKOJO.nkin_no AS [入金項目No]*/ "
                // tmp_sql = tmp_sql & " 		,nkin_name AS [入金項目表示用名称] "
                // tmp_sql = tmp_sql & " 		,so_gak AS [控除額] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN so_zeiumu = 0 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN so_zeiumu = 1 THEN 2 "
                // tmp_sql = tmp_sql & " 		 END AS [控除税区分] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN so_zeiumu = 0 THEN 0 "
                // tmp_sql = tmp_sql & " 			WHEN so_zeiumu = 1 THEN so_gak * 0.08 "
                // tmp_sql = tmp_sql & " 		 END AS [控除税額]	 "
                // tmp_sql = tmp_sql & " 		,1 AS [更新ロックフラグ]		/*デフォルト値 1:チェックON を設定しておく*/ "
                // tmp_sql = tmp_sql & " 		,2 AS [領収扱いフラグ] "
                // tmp_sql = tmp_sql & " 		,2 AS [立替・預りフラグ] "
                // tmp_sql = tmp_sql & " 		/* "
                // tmp_sql = tmp_sql & " 		,'' AS [預り予定GUID] "
                // tmp_sql = tmp_sql & " 		,'' AS [分割グループGUID] "
                // tmp_sql = tmp_sql & " 		,'' AS [固定控除ルールGUID] "
                // tmp_sql = tmp_sql & " 		,'' AS [控除ルールフラグ] "
                // tmp_sql = tmp_sql & " 		*/ "
                // tmp_sql = tmp_sql & " 		,'' AS [送金（控除）確定日] "
                // tmp_sql = tmp_sql & " 		/* "
                // tmp_sql = tmp_sql & " 		,'' AS [送金データ構築GUID] "
                // tmp_sql = tmp_sql & " 		,'' AS [送金ルールGUID] "
                // tmp_sql = tmp_sql & " 		,'' AS [送金ルールNo] "
                // tmp_sql = tmp_sql & " 		,'' AS [(送金先)送金ルールGUID] "
                // tmp_sql = tmp_sql & " 		,'' AS [(送金先)送金ルールNo] "
                // tmp_sql = tmp_sql & " 		*/ "
                // tmp_sql = tmp_sql & " 		,'' AS [(控除先指定時）送金先No] "
                // tmp_sql = tmp_sql & " 		,'' AS [(控除先指定時）送金先口座No] "
                // tmp_sql = tmp_sql & " 		/*,'' AS [敷金保証金随時処理GUID]*/ "
                // tmp_sql = tmp_sql & " 		/* "
                // tmp_sql = tmp_sql & " 		,'' AS [送金回収GUID] "
                // tmp_sql = tmp_sql & " 		*/ "
                // tmp_sql = tmp_sql & " 		,'' AS [管理手数料区分] "
                // tmp_sql = tmp_sql & " 		,'' AS [管理手数料請求条件] "
                // tmp_sql = tmp_sql & " 		,'' AS [管理手数料率] "
                // tmp_sql = tmp_sql & " 		,'' AS [修繕区分] "
                // tmp_sql = tmp_sql & " 		,'' AS [随時修繕No] "
                // tmp_sql = tmp_sql & " 		,'' AS [修繕請求No] "
                // tmp_sql = tmp_sql & " 		,'' AS [仕訳フラグ] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN so_zeiumu = 0 THEN NULL "
                // tmp_sql = tmp_sql & " 			WHEN so_zeiumu = 1 THEN 8 "
                // tmp_sql = tmp_sql & " 		 END AS [適用税率]	 "
                // tmp_sql = tmp_sql & " 		,biko AS [備考] "
                // tmp_sql = tmp_sql & " 		,'' AS [控除データ取込日] "
                // tmp_sql = tmp_sql & " 		/* "
                // tmp_sql = tmp_sql & " 		,'' AS [区分グループguid] "
                // tmp_sql = tmp_sql & " 		*/ "
                // tmp_sql = tmp_sql & " 		,ky_start_ymd AS [契約開始日]	/*送金ルール情報取得用*/ "
                // tmp_sql = tmp_sql & " 		,ky_end_ymd AS [契約終了日]		/*送金ルール情報取得用*/ "
                // tmp_sql = tmp_sql & " 	FROM "
                // tmp_sql = tmp_sql & " 	( "
                // tmp_sql = tmp_sql & " 		SELECT * FROM ky_sqdata "
                // tmp_sql = tmp_sql & " 		WHERE nkin_kbn = 7 "
                // tmp_sql = tmp_sql & " 	) AS KYKOJO "
                // tmp_sql = tmp_sql & " 	LEFT JOIN "
                // tmp_sql = tmp_sql & " 		( "
                // tmp_sql = tmp_sql & " 			SELECT bk_no,hy_no,ky_no,ko_no,ky_start_ymd,ky_end_ymd "
                // tmp_sql = tmp_sql & " 				  ,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約レコードNo] "
                // tmp_sql = tmp_sql & " 			FROM ky_kosinkai AS KYK "
                // tmp_sql = tmp_sql & " 		) AS KYRECNO "
                // tmp_sql = tmp_sql & " 	ON  KYKOJO.bk_no = KYRECNO.bk_no "
                // tmp_sql = tmp_sql & " 	AND KYKOJO.hy_no = KYRECNO.hy_no "
                // tmp_sql = tmp_sql & " 	AND KYKOJO.ky_no = KYRECNO.ky_no "
                // tmp_sql = tmp_sql & " 	AND KYKOJO.ko_no = KYRECNO.ko_no "
                // tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin AS MN ON KYKOJO.nkin_no = MN.nkin_no "
                // tmp_sql = tmp_sql & " 	WHERE so_gak <> 0 "
                // tmp_sql = tmp_sql & " ) AS VW "

                sortstr = "[物件No],[部屋No],[控除予定日],[項目ソートNo]";

                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 SHK.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,'' AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,so_ymd AS [控除予定日] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY so_ymd,SHK.bk_no ORDER BY mei_no) AS [項目ソートNo] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(gt_ym,'') <> '' THEN gt_ym ";
                tmp_sql = tmp_sql + " 			ELSE DATEADD(DAY,(DAY(so_ymd) - 1 ) * (-1),so_ymd) ";
                tmp_sql = tmp_sql + " 		 END AS [該当年月] ";
                tmp_sql = tmp_sql + " 		/*,nkin_no AS [入金項目名]*/ ";
                tmp_sql = tmp_sql + " 		,MN.nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 		,SHK.nkin_name AS [表示用入金項目名] ";
                tmp_sql = tmp_sql + " 		,kj_gak AS [控除額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kj_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN kj_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		 END AS [控除税区分] ";
                tmp_sql = tmp_sql + " 		,kj_zeigak AS [控除税額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN inpu_umu = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN inpu_umu = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 		 END AS [更新ロックフラグ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN SHK.ryo_umu = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN SHK.ryo_umu = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 		 END AS [領収扱いフラグ] ";
                tmp_sql = tmp_sql + " 		,2 AS [立替・預りフラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [送金（控除）確定日] ";
                tmp_sql = tmp_sql + " 		,'' AS [(控除先指定時）送金先No] ";
                tmp_sql = tmp_sql + " 		,'' AS [(控除先指定時）送金先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約No] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約レコードNo] ";
                tmp_sql = tmp_sql + " 		,kn_tesu_kbn AS [管理手数料区分] ";
                tmp_sql = tmp_sql + " 		,1 AS [管理手数料請求条件] ";
                tmp_sql = tmp_sql + " 		,kanritesu_rit1 AS [管理手数料率] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [随時修繕No] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕請求No] ";
                tmp_sql = tmp_sql + " 		,'' AS [仕訳フラグ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kj_zeiumu = 0 THEN NULL ";
                tmp_sql = tmp_sql + " 			WHEN kj_zeiumu = 1 THEN 8 ";
                tmp_sql = tmp_sql + " 		 END AS [適用税率]	 ";
                tmp_sql = tmp_sql + " 		,SHK.biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [控除データ取込日] ";
                tmp_sql = tmp_sql + " 		,kn_no AS [物件管理No]			/*送金ルール取得用*/ ";
                tmp_sql = tmp_sql + " 		,so_no AS [送金先No]			/*送金ルール取得用*/ ";
                tmp_sql = tmp_sql + " 		,sokoza_no AS [送金先口座No]	/*送金ルール取得用*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM sh_kojo ";
                tmp_sql = tmp_sql + " 		WHERE sq_umu = 0 ";
                tmp_sql = tmp_sql + " 		AND   nkin_no <> 9910	/*一括借上額は除外*/ ";
                tmp_sql = tmp_sql + " 	) AS SHK ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bk_kanri AS BKK ON SHK.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN ON SHK.nkin_no = MN.nkin_no ";
                tmp_sql = tmp_sql + " 	WHERE so_ymd BETWEEN BKK.start_ym AND BKK.end_ym	/*控除予定日が物件管理情報の適用開始期間内のデータを取得*/ ";
                tmp_sql = tmp_sql + " ) AS VW ";
                // 20160601 家主請求控除情報の取得方法を変更 -chg end

                return tmp_sql;

            }

            public bool Chk_Data(string tblname, string keyvalue, List<string> list, SafeDictionary<string, string> hash_chkbefore, ref SafeDictionary<string, string> hash_chkafter, ref int conditioncnt)
            {
                return default;

            }

            public void Get_BaseKey(SqlConnection sqlcnnv10, string tblname, ref List<string> list_basedata)
            {

            }

            public bool Set_Vari(SqlConnection sqlcnnv10, string filename, string sheetname, ref int midrowcnt, ref int cvrowcnt, ref int conditioncnt)
            {
                return default;

            }

            public void Get_ExistData(SqlConnection sqlcnnv10, string tblname, ref List<string> list_existdata)
            {

            }

        }

    }

    #endregion

}