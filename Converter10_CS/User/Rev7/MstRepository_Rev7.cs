using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using Converter10.Njc.Common;

namespace Converter10.Njc.Repository
{

    #region バス交通マスタ(汎用ツールのみだがバス停マスタの親になるため作成する)

    public class M_bus_V10_Repository
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

                sortstr = "[バス交通No]";

                tmp_sql = tmp_sql + " SELECT ";
                tmp_sql = tmp_sql + " 	 1 AS [バス交通No] ";
                tmp_sql = tmp_sql + " 	 ,'バス会社名' AS [バス会社名] ";

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

    #region バス停マスタ

    public class M_bus_Repository
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

                sortstr = "[バス交通No],[バス停No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[バス交通No],[バス停No],[系統名],[バス停名]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 1 AS [バス交通No] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(ORDER BY bus_no) AS [バス停No] ";
                tmp_sql = tmp_sql + " 		,brui_name AS [系統名] ";
                tmp_sql = tmp_sql + " 		,bus_name AS [バス停名] ";
                tmp_sql = tmp_sql + " 	FROM m_bus ";
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

    #region 鍵タイトルマスタ

    public class M_biko4_Repository
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

                sortstr = "[鍵No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[鍵No],[鍵名称]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 biko_no AS [鍵No] ";
                tmp_sql = tmp_sql + " 		,biko_titl AS [鍵名称] ";
                tmp_sql = tmp_sql + " 	FROM m_biko ";
                tmp_sql = tmp_sql + " 	WHERE biko_kbn = 4 ";
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

    #region 箇所クレーム分類マスタ

    public class Claim_data_Rev7_Repository
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

                sortstr = "[箇所・クレーム分類区分],[分類No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[箇所・クレーム分類区分],[分類No],[名称]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	/*V7ハードコーディング箇所 'kakaka2 ハードコーディングではなく[claim_data]から取得しています。sta*/ ";
                tmp_sql = tmp_sql + " 	SELECT 1 AS [箇所・クレーム分類区分],1 AS [分類No],'建物(専用部)' AS [名称] UNION ";
                tmp_sql = tmp_sql + " 	SELECT 1 AS [箇所・クレーム分類区分],2 AS [分類No],'建物(共用部)' AS [名称] UNION ";
                tmp_sql = tmp_sql + " 	SELECT 1 AS [箇所・クレーム分類区分],3 AS [分類No],'外構' AS [名称] UNION ";
                tmp_sql = tmp_sql + " 	SELECT 1 AS [箇所・クレーム分類区分],4 AS [分類No],'駐車場' AS [名称] UNION ";
                tmp_sql = tmp_sql + " 	SELECT 2 AS [箇所・クレーム分類区分],1 AS [分類No],'クレーム' AS [名称] UNION ";
                tmp_sql = tmp_sql + " 	SELECT 2 AS [箇所・クレーム分類区分],2 AS [分類No],'要望' AS [名称] UNION ";
                tmp_sql = tmp_sql + " 	SELECT 2 AS [箇所・クレーム分類区分],3 AS [分類No],'その他' AS [名称] UNION ";
                tmp_sql = tmp_sql + " 	/*V7ハードコーディング箇所 end*/ ";
                tmp_sql = tmp_sql + " 	/*実データ取得 sta*/ ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 1 AS [箇所・クレーム分類区分] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER() OVER(ORDER BY kasho_rui) + 4 AS [分類No] ";
                tmp_sql = tmp_sql + " 		,kasho_rui AS [名称] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT DISTINCT kasho_rui FROM claim_data ";
                tmp_sql = tmp_sql + " 			WHERE ISNULL(kasho_rui,'') <> '' ";
                tmp_sql = tmp_sql + " 		) AS KASYO ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 2 AS [箇所・クレーム分類区分] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER() OVER(ORDER BY claim_rui) + 3 AS [分類No] ";
                tmp_sql = tmp_sql + " 		,claim_rui AS [名称] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT DISTINCT claim_rui FROM claim_data ";
                tmp_sql = tmp_sql + " 			WHERE ISNULL(claim_rui,'') <> '' ";
                tmp_sql = tmp_sql + " 		) AS CLAIM ";
                tmp_sql = tmp_sql + " 	/*実データ取得 end*/ ";
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

    #region 特約マスタ

    public class M_tokuyaku_Rev7_Repository
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

                // kakaka2 特約・原状回復特約・修繕特約について見直してください。マスタ作成はV7の各マスタから、部屋・契約情報の特約データは以下のクエリより取得。(理由：たまたま使用していない特約マスタの任意値を、現移行方針だと使用できなくなる為)

                string tmp_sql = "";

                sortstr = "[特約区分],[特約No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[特約区分],[特約No],[特約タイトル],[特約詳細]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 1 AS [特約区分] ";
                tmp_sql = tmp_sql + " 		,tokuyaku_no AS [特約No] ";
                tmp_sql = tmp_sql + " 		,tokuyaku_name AS [特約タイトル] ";
                tmp_sql = tmp_sql + " 		,TOKUAYKU.内容1 + TOKUAYKU.内容2 + TOKUAYKU.内容3 + TOKUAYKU.内容4 + TOKUAYKU.内容5 + TOKUAYKU.内容6 + TOKUAYKU.内容7 + TOKUAYKU.内容8 + TOKUAYKU.内容9 + TOKUAYKU.内容10 +  ";
                tmp_sql = tmp_sql + " 		 TOKUAYKU.内容11 + TOKUAYKU.内容12 + TOKUAYKU.内容13 + TOKUAYKU.内容14 + TOKUAYKU.内容15 + TOKUAYKU.内容16 + TOKUAYKU.内容17 + TOKUAYKU.内容18 + TOKUAYKU.内容19 + TOKUAYKU.内容20 +  ";
                tmp_sql = tmp_sql + " 		 TOKUAYKU.内容21 + TOKUAYKU.内容22 + TOKUAYKU.内容23 + TOKUAYKU.内容24 + TOKUAYKU.内容25 + TOKUAYKU.内容26 + TOKUAYKU.内容27 + TOKUAYKU.内容28 + TOKUAYKU.内容29 + TOKUAYKU.内容30 +  ";
                tmp_sql = tmp_sql + " 		 TOKUAYKU.内容31 + TOKUAYKU.内容32 + TOKUAYKU.内容33 + TOKUAYKU.内容34 + TOKUAYKU.内容35 + TOKUAYKU.内容36 + TOKUAYKU.内容37 + TOKUAYKU.内容38 + TOKUAYKU.内容39 + TOKUAYKU.内容40 +  ";
                tmp_sql = tmp_sql + " 		 TOKUAYKU.内容41 + TOKUAYKU.内容42 + TOKUAYKU.内容43 + TOKUAYKU.内容44 + TOKUAYKU.内容45 + TOKUAYKU.内容46 + TOKUAYKU.内容47 + TOKUAYKU.内容48 + TOKUAYKU.内容49 + TOKUAYKU.内容50 ";
                tmp_sql = tmp_sql + " 		 AS [特約詳細] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 [tokuyaku_kbn] ";
                tmp_sql = tmp_sql + " 			,[tokuyaku_no] ";
                tmp_sql = tmp_sql + " 			,[tokuyaku_name] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 1 THEN [特約内容] ELSE '' END) AS 内容1 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 2 THEN [特約内容] ELSE '' END) AS 内容2 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 3 THEN [特約内容] ELSE '' END) AS 内容3 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 4 THEN [特約内容] ELSE '' END) AS 内容4 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 5 THEN [特約内容] ELSE '' END) AS 内容5 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 6 THEN [特約内容] ELSE '' END) AS 内容6 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 7 THEN [特約内容] ELSE '' END) AS 内容7 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 8 THEN [特約内容] ELSE '' END) AS 内容8 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 9 THEN [特約内容] ELSE '' END) AS 内容9 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 10 THEN [特約内容] ELSE '' END) AS 内容10 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 11 THEN [特約内容] ELSE '' END) AS 内容11 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 12 THEN [特約内容] ELSE '' END) AS 内容12 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 13 THEN [特約内容] ELSE '' END) AS 内容13 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 14 THEN [特約内容] ELSE '' END) AS 内容14 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 15 THEN [特約内容] ELSE '' END) AS 内容15 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 16 THEN [特約内容] ELSE '' END) AS 内容16 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 17 THEN [特約内容] ELSE '' END) AS 内容17 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 18 THEN [特約内容] ELSE '' END) AS 内容18 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 19 THEN [特約内容] ELSE '' END) AS 内容19 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 20 THEN [特約内容] ELSE '' END) AS 内容20 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 21 THEN [特約内容] ELSE '' END) AS 内容21 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 22 THEN [特約内容] ELSE '' END) AS 内容22 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 23 THEN [特約内容] ELSE '' END) AS 内容23 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 24 THEN [特約内容] ELSE '' END) AS 内容24 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 25 THEN [特約内容] ELSE '' END) AS 内容25 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 26 THEN [特約内容] ELSE '' END) AS 内容26 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 27 THEN [特約内容] ELSE '' END) AS 内容27 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 28 THEN [特約内容] ELSE '' END) AS 内容28 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 29 THEN [特約内容] ELSE '' END) AS 内容29 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 30 THEN [特約内容] ELSE '' END) AS 内容30 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 31 THEN [特約内容] ELSE '' END) AS 内容31 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 32 THEN [特約内容] ELSE '' END) AS 内容32 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 33 THEN [特約内容] ELSE '' END) AS 内容33 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 34 THEN [特約内容] ELSE '' END) AS 内容34 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 35 THEN [特約内容] ELSE '' END) AS 内容35 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 36 THEN [特約内容] ELSE '' END) AS 内容36 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 37 THEN [特約内容] ELSE '' END) AS 内容37 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 38 THEN [特約内容] ELSE '' END) AS 内容38 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 39 THEN [特約内容] ELSE '' END) AS 内容39 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 40 THEN [特約内容] ELSE '' END) AS 内容40 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 41 THEN [特約内容] ELSE '' END) AS 内容41 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 42 THEN [特約内容] ELSE '' END) AS 内容42 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 43 THEN [特約内容] ELSE '' END) AS 内容43 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 44 THEN [特約内容] ELSE '' END) AS 内容44 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 45 THEN [特約内容] ELSE '' END) AS 内容45 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 46 THEN [特約内容] ELSE '' END) AS 内容46 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 47 THEN [特約内容] ELSE '' END) AS 内容47 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 48 THEN [特約内容] ELSE '' END) AS 内容48 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 49 THEN [特約内容] ELSE '' END) AS 内容49 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行№] = 50 THEN [特約内容] ELSE '' END) AS 内容50 ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 MT.tokuyaku_kbn ";
                tmp_sql = tmp_sql + " 				,MT.tokuyaku_no ";
                tmp_sql = tmp_sql + " 				,MT.tokuyaku_name ";
                tmp_sql = tmp_sql + " 				,MTMEI.tokuyaku_name AS [特約行№] ";
                tmp_sql = tmp_sql + " 				,MTMEI.tokuyaku_kana AS [特約内容] ";
                tmp_sql = tmp_sql + " 			FROM m_tokuyaku AS MT ";
                tmp_sql = tmp_sql + " 			LEFT JOIN m_tokuyakumei AS MTMEI ";
                tmp_sql = tmp_sql + " 			ON  MT.tokuyaku_kbn = MTMEI.tokuyaku_kbn ";
                tmp_sql = tmp_sql + " 			AND MT.tokuyaku_no = MTMEI.tokuyaku_no ";
                tmp_sql = tmp_sql + " 			WHERE MT.tokuyaku_kbn = 1 ";
                tmp_sql = tmp_sql + " 		) AS TOTAL ";
                tmp_sql = tmp_sql + " 		GROUP BY [tokuyaku_kbn],[tokuyaku_no],[tokuyaku_name] ";
                tmp_sql = tmp_sql + " 	) AS TOKUAYKU ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 2 AS [特約区分] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(ORDER BY tokuyaku_brui) AS [特約No] ";
                tmp_sql = tmp_sql + " 		,titl_mj + ' - ' + tokuyaku_name AS [特約タイトル] ";
                tmp_sql = tmp_sql + " 		,GENJO_TAIKYO.内容1 + GENJO_TAIKYO.内容2 + GENJO_TAIKYO.内容3 + GENJO_TAIKYO.内容4 + GENJO_TAIKYO.内容5 + GENJO_TAIKYO.内容6 + GENJO_TAIKYO.内容7 + GENJO_TAIKYO.内容8 + GENJO_TAIKYO.内容9 + GENJO_TAIKYO.内容10 +  ";
                tmp_sql = tmp_sql + " 		 GENJO_TAIKYO.内容11 + GENJO_TAIKYO.内容12 + GENJO_TAIKYO.内容13 + GENJO_TAIKYO.内容14 + GENJO_TAIKYO.内容15 + GENJO_TAIKYO.内容16 + GENJO_TAIKYO.内容17 + GENJO_TAIKYO.内容18 + GENJO_TAIKYO.内容19 + GENJO_TAIKYO.内容20 +  ";
                tmp_sql = tmp_sql + " 		 GENJO_TAIKYO.内容21 + GENJO_TAIKYO.内容22 + GENJO_TAIKYO.内容23 + GENJO_TAIKYO.内容24 + GENJO_TAIKYO.内容25 + GENJO_TAIKYO.内容26 + GENJO_TAIKYO.内容27 + GENJO_TAIKYO.内容28 + GENJO_TAIKYO.内容29 + GENJO_TAIKYO.内容30 +  ";
                tmp_sql = tmp_sql + " 		 GENJO_TAIKYO.内容31 + GENJO_TAIKYO.内容32 + GENJO_TAIKYO.内容33 + GENJO_TAIKYO.内容34 + GENJO_TAIKYO.内容35 + GENJO_TAIKYO.内容36 + GENJO_TAIKYO.内容37 + GENJO_TAIKYO.内容38 + GENJO_TAIKYO.内容39 + GENJO_TAIKYO.内容40 +  ";
                tmp_sql = tmp_sql + " 		 GENJO_TAIKYO.内容41 + GENJO_TAIKYO.内容42 + GENJO_TAIKYO.内容43 + GENJO_TAIKYO.内容44 + GENJO_TAIKYO.内容45 + GENJO_TAIKYO.内容46 + GENJO_TAIKYO.内容47 + GENJO_TAIKYO.内容48 + GENJO_TAIKYO.内容49 + GENJO_TAIKYO.内容50 ";
                tmp_sql = tmp_sql + " 		 AS [特約詳細] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 tokuyaku_kbn ";
                tmp_sql = tmp_sql + " 			,tokuyaku_brui ";
                tmp_sql = tmp_sql + " 			,titl_mj ";
                tmp_sql = tmp_sql + " 			,tokuyaku_no ";
                tmp_sql = tmp_sql + " 			,tokuyaku_name ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 1 THEN [原状特約内容] ELSE '' END) AS 内容1 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 2 THEN [原状特約内容] ELSE '' END) AS 内容2 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 3 THEN [原状特約内容] ELSE '' END) AS 内容3 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 4 THEN [原状特約内容] ELSE '' END) AS 内容4 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 5 THEN [原状特約内容] ELSE '' END) AS 内容5 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 6 THEN [原状特約内容] ELSE '' END) AS 内容6 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 7 THEN [原状特約内容] ELSE '' END) AS 内容7 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 8 THEN [原状特約内容] ELSE '' END) AS 内容8 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 9 THEN [原状特約内容] ELSE '' END) AS 内容9 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 10 THEN [原状特約内容] ELSE '' END) AS 内容10 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 11 THEN [原状特約内容] ELSE '' END) AS 内容11 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 12 THEN [原状特約内容] ELSE '' END) AS 内容12 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 13 THEN [原状特約内容] ELSE '' END) AS 内容13 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 14 THEN [原状特約内容] ELSE '' END) AS 内容14 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 15 THEN [原状特約内容] ELSE '' END) AS 内容15 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 16 THEN [原状特約内容] ELSE '' END) AS 内容16 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 17 THEN [原状特約内容] ELSE '' END) AS 内容17 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 18 THEN [原状特約内容] ELSE '' END) AS 内容18 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 19 THEN [原状特約内容] ELSE '' END) AS 内容19 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 20 THEN [原状特約内容] ELSE '' END) AS 内容20 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 21 THEN [原状特約内容] ELSE '' END) AS 内容21 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 22 THEN [原状特約内容] ELSE '' END) AS 内容22 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 23 THEN [原状特約内容] ELSE '' END) AS 内容23 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 24 THEN [原状特約内容] ELSE '' END) AS 内容24 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 25 THEN [原状特約内容] ELSE '' END) AS 内容25 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 26 THEN [原状特約内容] ELSE '' END) AS 内容26 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 27 THEN [原状特約内容] ELSE '' END) AS 内容27 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 28 THEN [原状特約内容] ELSE '' END) AS 内容28 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 29 THEN [原状特約内容] ELSE '' END) AS 内容29 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 30 THEN [原状特約内容] ELSE '' END) AS 内容30 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 31 THEN [原状特約内容] ELSE '' END) AS 内容31 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 32 THEN [原状特約内容] ELSE '' END) AS 内容32 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 33 THEN [原状特約内容] ELSE '' END) AS 内容33 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 34 THEN [原状特約内容] ELSE '' END) AS 内容34 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 35 THEN [原状特約内容] ELSE '' END) AS 内容35 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 36 THEN [原状特約内容] ELSE '' END) AS 内容36 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 37 THEN [原状特約内容] ELSE '' END) AS 内容37 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 38 THEN [原状特約内容] ELSE '' END) AS 内容38 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 39 THEN [原状特約内容] ELSE '' END) AS 内容39 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 40 THEN [原状特約内容] ELSE '' END) AS 内容40 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 41 THEN [原状特約内容] ELSE '' END) AS 内容41 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 42 THEN [原状特約内容] ELSE '' END) AS 内容42 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 43 THEN [原状特約内容] ELSE '' END) AS 内容43 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 44 THEN [原状特約内容] ELSE '' END) AS 内容44 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 45 THEN [原状特約内容] ELSE '' END) AS 内容45 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 46 THEN [原状特約内容] ELSE '' END) AS 内容46 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 47 THEN [原状特約内容] ELSE '' END) AS 内容47 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 48 THEN [原状特約内容] ELSE '' END) AS 内容48 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 49 THEN [原状特約内容] ELSE '' END) AS 内容49 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 50 THEN [原状特約内容] ELSE '' END) AS 内容50 ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 GENJO.tokuyaku_kbn AS tokuyaku_kbn ";
                tmp_sql = tmp_sql + " 				,GENJO.tokuyaku_brui AS tokuyaku_brui ";
                tmp_sql = tmp_sql + " 				,BRUITTL.titl_mj AS titl_mj ";
                tmp_sql = tmp_sql + " 				,GENJO.tokuyaku_no AS tokuyaku_no ";
                tmp_sql = tmp_sql + " 				,GENJO.tokuyaku_name AS tokuyaku_name ";
                tmp_sql = tmp_sql + " 				,GENJOMEI.tokuyaku_name AS [原状特約行№] ";
                tmp_sql = tmp_sql + " 				,GENJOMEI.tokuyaku_kana AS [原状特約内容] ";
                tmp_sql = tmp_sql + " 			FROM m_tokuyaku_genjo AS GENJO ";
                tmp_sql = tmp_sql + " 			LEFT JOIN m_tokuyakumei_genjo AS GENJOMEI ";
                tmp_sql = tmp_sql + " 			ON  GENJO.tokuyaku_kbn = GENJOMEI.tokuyaku_kbn ";
                tmp_sql = tmp_sql + " 			AND GENJO.tokuyaku_brui = GENJOMEI.tokuyaku_brui ";
                tmp_sql = tmp_sql + " 			AND GENJO.tokuyaku_no = GENJOMEI.tokuyaku_no ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 biko_no ";
                tmp_sql = tmp_sql + " 						,titl_mj ";
                tmp_sql = tmp_sql + " 					FROM m_biko ";
                tmp_sql = tmp_sql + " 					WHERE biko_kbn = 10 ";
                tmp_sql = tmp_sql + " 				) AS BRUITTL ";
                tmp_sql = tmp_sql + " 			ON GENJO.tokuyaku_brui = BRUITTL.biko_no ";
                tmp_sql = tmp_sql + " 			WHERE GENJO.tokuyaku_kbn = 1 ";
                tmp_sql = tmp_sql + " 			/*2016.04.06 原状特約を分類毎に抽出する add sta*/ ";
                tmp_sql = tmp_sql + " 			AND   GENJO.tokuyaku_brui = 1 ";
                tmp_sql = tmp_sql + " 			/*2016.04.06 原状特約を分類毎に抽出する add end*/ ";
                tmp_sql = tmp_sql + " 		) AS TOTAL ";
                tmp_sql = tmp_sql + " 		GROUP BY tokuyaku_kbn,tokuyaku_brui,titl_mj,tokuyaku_no,tokuyaku_name ";
                tmp_sql = tmp_sql + " 	) AS GENJO_TAIKYO ";
                tmp_sql = tmp_sql + " 	/*2016.04.06 原状特約を分類毎に抽出する add sta*/ ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 3 AS [特約区分] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(ORDER BY tokuyaku_brui) AS [特約No] ";
                tmp_sql = tmp_sql + " 		,titl_mj + ' - ' + tokuyaku_name AS [特約タイトル] ";
                tmp_sql = tmp_sql + " 		,GENJO_JYUTAKU.内容1 + GENJO_JYUTAKU.内容2 + GENJO_JYUTAKU.内容3 + GENJO_JYUTAKU.内容4 + GENJO_JYUTAKU.内容5 + GENJO_JYUTAKU.内容6 + GENJO_JYUTAKU.内容7 + GENJO_JYUTAKU.内容8 + GENJO_JYUTAKU.内容9 + GENJO_JYUTAKU.内容10 +  ";
                tmp_sql = tmp_sql + " 		 GENJO_JYUTAKU.内容11 + GENJO_JYUTAKU.内容12 + GENJO_JYUTAKU.内容13 + GENJO_JYUTAKU.内容14 + GENJO_JYUTAKU.内容15 + GENJO_JYUTAKU.内容16 + GENJO_JYUTAKU.内容17 + GENJO_JYUTAKU.内容18 + GENJO_JYUTAKU.内容19 + GENJO_JYUTAKU.内容20 +  ";
                tmp_sql = tmp_sql + " 		 GENJO_JYUTAKU.内容21 + GENJO_JYUTAKU.内容22 + GENJO_JYUTAKU.内容23 + GENJO_JYUTAKU.内容24 + GENJO_JYUTAKU.内容25 + GENJO_JYUTAKU.内容26 + GENJO_JYUTAKU.内容27 + GENJO_JYUTAKU.内容28 + GENJO_JYUTAKU.内容29 + GENJO_JYUTAKU.内容30 +  ";
                tmp_sql = tmp_sql + " 		 GENJO_JYUTAKU.内容31 + GENJO_JYUTAKU.内容32 + GENJO_JYUTAKU.内容33 + GENJO_JYUTAKU.内容34 + GENJO_JYUTAKU.内容35 + GENJO_JYUTAKU.内容36 + GENJO_JYUTAKU.内容37 + GENJO_JYUTAKU.内容38 + GENJO_JYUTAKU.内容39 + GENJO_JYUTAKU.内容40 +  ";
                tmp_sql = tmp_sql + " 		 GENJO_JYUTAKU.内容41 + GENJO_JYUTAKU.内容42 + GENJO_JYUTAKU.内容43 + GENJO_JYUTAKU.内容44 + GENJO_JYUTAKU.内容45 + GENJO_JYUTAKU.内容46 + GENJO_JYUTAKU.内容47 + GENJO_JYUTAKU.内容48 + GENJO_JYUTAKU.内容49 + GENJO_JYUTAKU.内容50 ";
                tmp_sql = tmp_sql + " 		 AS [特約詳細] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 tokuyaku_kbn ";
                tmp_sql = tmp_sql + " 			,tokuyaku_brui ";
                tmp_sql = tmp_sql + " 			,titl_mj ";
                tmp_sql = tmp_sql + " 			,tokuyaku_no ";
                tmp_sql = tmp_sql + " 			,tokuyaku_name ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 1 THEN [原状特約内容] ELSE '' END) AS 内容1 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 2 THEN [原状特約内容] ELSE '' END) AS 内容2 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 3 THEN [原状特約内容] ELSE '' END) AS 内容3 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 4 THEN [原状特約内容] ELSE '' END) AS 内容4 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 5 THEN [原状特約内容] ELSE '' END) AS 内容5 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 6 THEN [原状特約内容] ELSE '' END) AS 内容6 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 7 THEN [原状特約内容] ELSE '' END) AS 内容7 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 8 THEN [原状特約内容] ELSE '' END) AS 内容8 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 9 THEN [原状特約内容] ELSE '' END) AS 内容9 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 10 THEN [原状特約内容] ELSE '' END) AS 内容10 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 11 THEN [原状特約内容] ELSE '' END) AS 内容11 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 12 THEN [原状特約内容] ELSE '' END) AS 内容12 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 13 THEN [原状特約内容] ELSE '' END) AS 内容13 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 14 THEN [原状特約内容] ELSE '' END) AS 内容14 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 15 THEN [原状特約内容] ELSE '' END) AS 内容15 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 16 THEN [原状特約内容] ELSE '' END) AS 内容16 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 17 THEN [原状特約内容] ELSE '' END) AS 内容17 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 18 THEN [原状特約内容] ELSE '' END) AS 内容18 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 19 THEN [原状特約内容] ELSE '' END) AS 内容19 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 20 THEN [原状特約内容] ELSE '' END) AS 内容20 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 21 THEN [原状特約内容] ELSE '' END) AS 内容21 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 22 THEN [原状特約内容] ELSE '' END) AS 内容22 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 23 THEN [原状特約内容] ELSE '' END) AS 内容23 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 24 THEN [原状特約内容] ELSE '' END) AS 内容24 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 25 THEN [原状特約内容] ELSE '' END) AS 内容25 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 26 THEN [原状特約内容] ELSE '' END) AS 内容26 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 27 THEN [原状特約内容] ELSE '' END) AS 内容27 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 28 THEN [原状特約内容] ELSE '' END) AS 内容28 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 29 THEN [原状特約内容] ELSE '' END) AS 内容29 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 30 THEN [原状特約内容] ELSE '' END) AS 内容30 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 31 THEN [原状特約内容] ELSE '' END) AS 内容31 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 32 THEN [原状特約内容] ELSE '' END) AS 内容32 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 33 THEN [原状特約内容] ELSE '' END) AS 内容33 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 34 THEN [原状特約内容] ELSE '' END) AS 内容34 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 35 THEN [原状特約内容] ELSE '' END) AS 内容35 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 36 THEN [原状特約内容] ELSE '' END) AS 内容36 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 37 THEN [原状特約内容] ELSE '' END) AS 内容37 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 38 THEN [原状特約内容] ELSE '' END) AS 内容38 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 39 THEN [原状特約内容] ELSE '' END) AS 内容39 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 40 THEN [原状特約内容] ELSE '' END) AS 内容40 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 41 THEN [原状特約内容] ELSE '' END) AS 内容41 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 42 THEN [原状特約内容] ELSE '' END) AS 内容42 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 43 THEN [原状特約内容] ELSE '' END) AS 内容43 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 44 THEN [原状特約内容] ELSE '' END) AS 内容44 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 45 THEN [原状特約内容] ELSE '' END) AS 内容45 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 46 THEN [原状特約内容] ELSE '' END) AS 内容46 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 47 THEN [原状特約内容] ELSE '' END) AS 内容47 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 48 THEN [原状特約内容] ELSE '' END) AS 内容48 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 49 THEN [原状特約内容] ELSE '' END) AS 内容49 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [原状特約行№] = 50 THEN [原状特約内容] ELSE '' END) AS 内容50 ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 GENJO.tokuyaku_kbn AS tokuyaku_kbn ";
                tmp_sql = tmp_sql + " 				,GENJO.tokuyaku_brui AS tokuyaku_brui ";
                tmp_sql = tmp_sql + " 				,BRUITTL.titl_mj AS titl_mj ";
                tmp_sql = tmp_sql + " 				,GENJO.tokuyaku_no AS tokuyaku_no ";
                tmp_sql = tmp_sql + " 				,GENJO.tokuyaku_name AS tokuyaku_name ";
                tmp_sql = tmp_sql + " 				,GENJOMEI.tokuyaku_name AS [原状特約行№] ";
                tmp_sql = tmp_sql + " 				,GENJOMEI.tokuyaku_kana AS [原状特約内容] ";
                tmp_sql = tmp_sql + " 			FROM m_tokuyaku_genjo AS GENJO ";
                tmp_sql = tmp_sql + " 			LEFT JOIN m_tokuyakumei_genjo AS GENJOMEI ";
                tmp_sql = tmp_sql + " 			ON  GENJO.tokuyaku_kbn = GENJOMEI.tokuyaku_kbn ";
                tmp_sql = tmp_sql + " 			AND GENJO.tokuyaku_brui = GENJOMEI.tokuyaku_brui ";
                tmp_sql = tmp_sql + " 			AND GENJO.tokuyaku_no = GENJOMEI.tokuyaku_no ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 biko_no ";
                tmp_sql = tmp_sql + " 						,titl_mj ";
                tmp_sql = tmp_sql + " 					FROM m_biko ";
                tmp_sql = tmp_sql + " 					WHERE biko_kbn = 10 ";
                tmp_sql = tmp_sql + " 				) AS BRUITTL ";
                tmp_sql = tmp_sql + " 			ON GENJO.tokuyaku_brui = BRUITTL.biko_no ";
                tmp_sql = tmp_sql + " 			WHERE GENJO.tokuyaku_kbn = 1 ";
                tmp_sql = tmp_sql + " 			/*2016.04.06 原状特約を分類毎に抽出する add sta*/ ";
                tmp_sql = tmp_sql + " 			AND   GENJO.tokuyaku_brui = 2 ";
                tmp_sql = tmp_sql + " 			/*2016.04.06 原状特約を分類毎に抽出する add end*/ ";
                tmp_sql = tmp_sql + " 		) AS TOTAL ";
                tmp_sql = tmp_sql + " 		GROUP BY tokuyaku_kbn,tokuyaku_brui,titl_mj,tokuyaku_no,tokuyaku_name ";
                tmp_sql = tmp_sql + " 	) AS GENJO_JYUTAKU	 ";
                tmp_sql = tmp_sql + " 	/*2016.04.06 原状特約を分類毎に抽出する add end*/ ";
                tmp_sql = tmp_sql + " ) AS TOKUYAKUTOTAL ";

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

    #region 契約分類マスタ

    public class M_keirui_Repository
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
                return default;

                // メモ
                // 色変換は中間ファイル→変数格納処理時に行う

                // 紐付項目へ変更

                // Dim tmp_sql As String = ""

                // sortstr = "[契約分類No]"

                // tmp_sql = tmp_sql & " SELECT "
                // tmp_sql = tmp_sql & " 	 [keirui_no] AS [契約分類No] "
                // tmp_sql = tmp_sql & " 	,[keirui_name] AS [契約分類名] "
                // tmp_sql = tmp_sql & " 	,[biko] AS [備考] "
                // tmp_sql = tmp_sql & " 	,[yobi_si1] AS [更新・解約通知期間(ヶ月)] "
                // tmp_sql = tmp_sql & " 	,[yobi_fl1] AS [契約分類色] "
                // tmp_sql = tmp_sql & " 	,CASE "
                // tmp_sql = tmp_sql & " 		WHEN [keirui_no] = 2 THEN 1 "                       'kakaka2 通常は2が定借だが、契約分類マスタは変更できる為、この条件ではよくない。
                // tmp_sql = tmp_sql & " 		ELSE 2 "
                // tmp_sql = tmp_sql & " 	 END AS [定期借地借家権契約扱い有無] "
                // tmp_sql = tmp_sql & " FROM m_keirui "

                // Return tmp_sql

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

    #region 保険種類マスタ

    public class M_hoken_rui_Rev7_Repository
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

                sortstr = "[保険種類No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[保険種類No],[保険種類名],[保険種類カナ],[備考]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 [hknrui_no] AS [保険種類No] ";
                tmp_sql = tmp_sql + " 		,[hknrui_name] AS [保険種類名] ";
                tmp_sql = tmp_sql + " 		,[hknrui_kana] AS [保険種類カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考] ";
                tmp_sql = tmp_sql + " 	FROM m_hoken_rui ";
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

    #region 学校区マスタ

    public class M_school_Repository
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

                sortstr = "[県No],[市No],[学校区No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[県No],[市No],[学校区No],[町],[丁目],[番地],[小学校名],[中学校名]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 [ken_no] AS [県No] ";
                tmp_sql = tmp_sql + " 		,[si_no] AS [市No] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY ken_no,si_no ORDER BY ken_no,si_no) AS [学校区No] ";
                tmp_sql = tmp_sql + " 		,[add_cyo] AS [町] ";
                tmp_sql = tmp_sql + " 		,[add_banti] AS [丁目] ";
                tmp_sql = tmp_sql + " 		/*2016.04.06 学校区マスタ移行内容修正 -chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,[add_etc] AS [番地]*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [番地] ";
                tmp_sql = tmp_sql + " 		/*2016.04.06 学校区マスタ移行内容修正 -chg end*/ ";
                tmp_sql = tmp_sql + " 		,[syogaku_name] AS [小学校名] ";
                tmp_sql = tmp_sql + " 		,[cyugaku_name] AS [中学校名] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT DISTINCT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 [ken_no] ";
                tmp_sql = tmp_sql + " 				,[si_no] ";
                tmp_sql = tmp_sql + " 				,[add_cyo] ";
                tmp_sql = tmp_sql + " 				,[add_banti] ";
                tmp_sql = tmp_sql + " 				,[add_etc] ";
                tmp_sql = tmp_sql + " 				,[syogaku_name] ";
                tmp_sql = tmp_sql + " 				,[cyugaku_name] ";
                tmp_sql = tmp_sql + " 			FROM bk_mst ";
                tmp_sql = tmp_sql + " 			WHERE ken_no IS NOT NULL AND (ISNULL(syogaku_name,'') <> '' OR ISNULL(cyugaku_name,'') <> '') ";
                tmp_sql = tmp_sql + " 		) AS VW1 ";
                tmp_sql = tmp_sql + " 	) AS VW2 ";
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

    #region エリアマスタ

    public class M_addkbn_Repository
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

                sortstr = "[エリアNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[エリアNo],[エリア名]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 [addkbn_no] AS [エリアNo] ";
                tmp_sql = tmp_sql + " 		,[addkbn_name] AS [エリア名] ";
                tmp_sql = tmp_sql + " 	FROM m_addkbn ";
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

    #region 変動費マスタ

    public class M_hendo_Repository
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

                sortstr = "[ルールNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[ルールNo],[ルール名],[その他単位],[その他単位SJIS],[料金数],[分類],[端数処理],[端数処理パターン],[口径],[備考],[計上分類],[税区分早見],[税区分単価]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 hendo_no AS [ルールNo] ";
                tmp_sql = tmp_sql + " 		,hendo_name AS [ルール名] ";
                tmp_sql = tmp_sql + " 		,tanni AS [その他単位] ";
                tmp_sql = tmp_sql + " 		,'' AS [その他単位SJIS] ";
                tmp_sql = tmp_sql + " 		,ryokin_su AS [料金数] ";
                tmp_sql = tmp_sql + " 		,hendo_rui AS [分類] ";
                tmp_sql = tmp_sql + " 		,hasuu AS [端数処理] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN yobi_si1 = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN yobi_si1 = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN yobi_si1 = 2 THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN yobi_si1 = 3 THEN 4 ";
                tmp_sql = tmp_sql + " 		 END AS [端数処理パターン] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN hendo_rui = 1 THEN yobi_si2 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [口径] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考] ";
                tmp_sql = tmp_sql + " 		,1 AS [計上分類] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN zei_umu = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN zei_umu = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [税区分早見] ";
                tmp_sql = tmp_sql + " 		,'' AS [税区分単価] ";
                tmp_sql = tmp_sql + " 	FROM m_hendo ";
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

    #region 変動費一覧マスタ

    public class M_hendo_mei_Repository
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

                sortstr = "[変動費ルールNo],[行No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[変動費ルールNo],[行No],[使用量],[料金１],[料金２],[料金３],[料金４],[料金５]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 hendo_no AS [変動費ルールNo] ";
                tmp_sql = tmp_sql + " 		,meisai_no AS [行No] ";
                tmp_sql = tmp_sql + " 		,siyoryo AS [使用量] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ryokin1 <> 0 THEN ryokin1 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [料金１] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ryokin2 <> 0 THEN ryokin2 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [料金２] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ryokin3 <> 0 THEN ryokin3 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [料金３] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ryokin4 <> 0 THEN ryokin4 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [料金４] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ryokin5 <> 0 THEN ryokin5 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [料金５] ";
                tmp_sql = tmp_sql + " 	FROM m_hendo_mei ";
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

    #region 備考タイトルマスタ

    public class M_biko_Repository
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

                sortstr = "[備考区分],[備考No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[備考区分],[備考No],[備考タイトル],[使用有無],[補助アイテム使用有無]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 CASE ";
                tmp_sql = tmp_sql + " 				WHEN biko_kbn = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 				WHEN biko_kbn = 2 THEN 6 ";
                tmp_sql = tmp_sql + " 				WHEN biko_kbn = 3 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN biko_kbn = 5 THEN 4 ";
                tmp_sql = tmp_sql + " 				ELSE NULL ";
                tmp_sql = tmp_sql + " 			 END AS [備考区分] ";
                tmp_sql = tmp_sql + " 			,biko_no AS [備考No] ";
                tmp_sql = tmp_sql + " 			,biko_titl AS [備考タイトル] ";
                tmp_sql = tmp_sql + " 			,1 AS [使用有無] ";
                tmp_sql = tmp_sql + " 			,0 AS [補助アイテム使用有無] ";
                tmp_sql = tmp_sql + " 		FROM m_biko ";
                tmp_sql = tmp_sql + " 		WHERE biko_kbn <> 4	/*鍵タイトルは別で移行済み*/ ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 12 AS [備考区分] ";
                tmp_sql = tmp_sql + " 			,REFBIKO.jyuyo_no AS [備考No] ";
                tmp_sql = tmp_sql + " 			,jyuyo_name AS [備考タイトル] ";
                tmp_sql = tmp_sql + " 			,use_umu AS [使用有無] ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN REFBIKOLST.jyuyo_kbn IS NULL THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE 1 ";
                tmp_sql = tmp_sql + " 			 END AS [補助アイテム使用有無] ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM m_jyuyo WHERE jyuyo_kbn = 4 ";
                tmp_sql = tmp_sql + " 		) AS REFBIKO ";
                tmp_sql = tmp_sql + " 		LEFT JOIN ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT * FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM m_jyuyo_lst ";
                tmp_sql = tmp_sql + " 					WHERE jyuyo_kbn = 4 ";
                tmp_sql = tmp_sql + " 					AND   jyuyo_lstno = 1 ";
                tmp_sql = tmp_sql + " 				) AS VW ";
                tmp_sql = tmp_sql + " 			) AS REFBIKOLST ";
                tmp_sql = tmp_sql + " 		ON  REFBIKO.jyuyo_kbn = REFBIKOLST.jyuyo_kbn ";
                tmp_sql = tmp_sql + " 		AND REFBIKO.jyuyo_no = REFBIKOLST.jyuyo_no ";
                tmp_sql = tmp_sql + " 	) AS TOTAL ";
                tmp_sql = tmp_sql + " 	WHERE [備考区分] IS NOT NULL ";
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

    #region 備考入力補助リストマスタ

    public class M_biko_lst_Repository
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

                sortstr = "[備考区分],[備考No],[備考入力補助リストNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[備考区分],[備考No],[備考入力補助リストNo],[備考入力補助リスト内容]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		  12 AS 備考区分 ";
                tmp_sql = tmp_sql + " 		 ,jyuyo_no AS [備考No] ";
                tmp_sql = tmp_sql + " 		 ,jyuyo_lstno AS [備考入力補助リストNo] ";
                tmp_sql = tmp_sql + " 		 ,jyuyo_lstname AS [備考入力補助リスト内容] ";
                tmp_sql = tmp_sql + " 	FROM m_jyuyo_lst ";
                tmp_sql = tmp_sql + " 	WHERE jyuyo_kbn = 4 ";
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

    #region 画像タイトルマスタ

    public class M_gazo_syoki_Repository
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

                sortstr = "[画像区分],[画像No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[画像区分],[画像No],[画像名]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 CASE ";
                tmp_sql = tmp_sql + " 				WHEN gazo_kbn = 1 THEN 1	/*物件*/ ";
                tmp_sql = tmp_sql + " 				WHEN gazo_kbn = 2 THEN 2	/*部屋*/ ";
                tmp_sql = tmp_sql + " 				WHEN gazo_kbn = 3 THEN 8	/*自社*/ ";
                tmp_sql = tmp_sql + " 				WHEN gazo_kbn = 5 THEN 3	/*契約*/ ";
                tmp_sql = tmp_sql + " 			 END AS [画像区分] ";
                tmp_sql = tmp_sql + " 			,gazo_no AS [画像No] ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN ISNULL(gazo_name,'') <> '' THEN gazo_name ";
                tmp_sql = tmp_sql + " 				WHEN gazo_kbn = 1 AND ISNULL(gazo_name,'') = '' THEN '物件画像' + CONVERT(VARCHAR,gazo_no) ";
                tmp_sql = tmp_sql + " 				WHEN gazo_kbn = 2 AND ISNULL(gazo_name,'') = '' THEN '部屋画像' + CONVERT(VARCHAR,gazo_no) ";
                tmp_sql = tmp_sql + " 				WHEN gazo_kbn = 3 AND ISNULL(gazo_name,'') = '' THEN '自社画像' + CONVERT(VARCHAR,gazo_no) ";
                tmp_sql = tmp_sql + " 				WHEN gazo_kbn = 5 AND ISNULL(gazo_name,'') = '' THEN '契約画像' + CONVERT(VARCHAR,gazo_no) ";
                tmp_sql = tmp_sql + " 			 END AS [画像名]	 ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 gazo_kbn ";
                tmp_sql = tmp_sql + " 				,gazo_no ";
                tmp_sql = tmp_sql + " 				,gazo_name ";
                tmp_sql = tmp_sql + " 			FROM m_gazo_syoki ";
                tmp_sql = tmp_sql + " 		) AS GAZOTTL ";
                tmp_sql = tmp_sql + " 		WHERE gazo_kbn IN (1,2,3,5) ";
                tmp_sql = tmp_sql + " 	) AS EXISTDATA ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],21 AS [画像No],'物件画像21' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],22 AS [画像No],'物件画像22' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],23 AS [画像No],'物件画像23' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],24 AS [画像No],'物件画像24' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],25 AS [画像No],'物件画像25' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],26 AS [画像No],'物件画像26' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],27 AS [画像No],'物件画像27' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],28 AS [画像No],'物件画像28' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],29 AS [画像No],'物件画像29' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],30 AS [画像No],'物件画像30' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],31 AS [画像No],'物件画像31' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],32 AS [画像No],'物件画像32' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],33 AS [画像No],'物件画像33' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],34 AS [画像No],'物件画像34' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],35 AS [画像No],'物件画像35' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],36 AS [画像No],'物件画像36' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],37 AS [画像No],'物件画像37' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],38 AS [画像No],'物件画像38' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],39 AS [画像No],'物件画像39' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],40 AS [画像No],'物件画像40' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],41 AS [画像No],'物件画像41' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],42 AS [画像No],'物件画像42' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],43 AS [画像No],'物件画像43' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],44 AS [画像No],'物件画像44' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],45 AS [画像No],'物件画像45' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],46 AS [画像No],'物件画像46' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],47 AS [画像No],'物件画像47' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],48 AS [画像No],'物件画像48' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],49 AS [画像No],'物件画像49' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 1 AS [画像区分],50 AS [画像No],'物件画像50' AS [画像名] ";
                tmp_sql = tmp_sql + " 	) AS BKDEF ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],21 AS [画像No],'部屋画像21' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],22 AS [画像No],'部屋画像22' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],23 AS [画像No],'部屋画像23' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],24 AS [画像No],'部屋画像24' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],25 AS [画像No],'部屋画像25' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],26 AS [画像No],'部屋画像26' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],27 AS [画像No],'部屋画像27' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],28 AS [画像No],'部屋画像28' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],29 AS [画像No],'部屋画像29' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],30 AS [画像No],'部屋画像30' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],31 AS [画像No],'部屋画像31' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],32 AS [画像No],'部屋画像32' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],33 AS [画像No],'部屋画像33' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],34 AS [画像No],'部屋画像34' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],35 AS [画像No],'部屋画像35' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],36 AS [画像No],'部屋画像36' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],37 AS [画像No],'部屋画像37' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],38 AS [画像No],'部屋画像38' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],39 AS [画像No],'部屋画像39' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],40 AS [画像No],'部屋画像40' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],41 AS [画像No],'部屋画像41' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],42 AS [画像No],'部屋画像42' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],43 AS [画像No],'部屋画像43' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],44 AS [画像No],'部屋画像44' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],45 AS [画像No],'部屋画像45' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],46 AS [画像No],'部屋画像46' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],47 AS [画像No],'部屋画像47' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],48 AS [画像No],'部屋画像48' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],49 AS [画像No],'部屋画像49' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 2 AS [画像区分],50 AS [画像No],'部屋画像50' AS [画像名] ";
                tmp_sql = tmp_sql + " 	) AS HYDEF ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],21 AS [画像No],'契約画像21' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],22 AS [画像No],'契約画像22' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],23 AS [画像No],'契約画像23' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],24 AS [画像No],'契約画像24' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],25 AS [画像No],'契約画像25' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],26 AS [画像No],'契約画像26' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],27 AS [画像No],'契約画像27' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],28 AS [画像No],'契約画像28' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],29 AS [画像No],'契約画像29' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],30 AS [画像No],'契約画像30' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],31 AS [画像No],'契約画像31' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],32 AS [画像No],'契約画像32' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],33 AS [画像No],'契約画像33' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],34 AS [画像No],'契約画像34' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],35 AS [画像No],'契約画像35' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],36 AS [画像No],'契約画像36' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],37 AS [画像No],'契約画像37' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],38 AS [画像No],'契約画像38' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],39 AS [画像No],'契約画像39' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],40 AS [画像No],'契約画像40' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],41 AS [画像No],'契約画像41' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],42 AS [画像No],'契約画像42' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],43 AS [画像No],'契約画像43' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],44 AS [画像No],'契約画像44' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],45 AS [画像No],'契約画像45' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],46 AS [画像No],'契約画像46' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],47 AS [画像No],'契約画像47' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],48 AS [画像No],'契約画像48' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],49 AS [画像No],'契約画像49' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 3 AS [画像区分],50 AS [画像No],'契約画像50' AS [画像名] ";
                tmp_sql = tmp_sql + " 	) AS KYDEF ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],11 AS [画像No],'自社画像11' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],12 AS [画像No],'自社画像12' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],13 AS [画像No],'自社画像13' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],14 AS [画像No],'自社画像14' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],15 AS [画像No],'自社画像15' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],16 AS [画像No],'自社画像16' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],17 AS [画像No],'自社画像17' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],18 AS [画像No],'自社画像18' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],19 AS [画像No],'自社画像19' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],20 AS [画像No],'自社画像20' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],21 AS [画像No],'自社画像21' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],22 AS [画像No],'自社画像22' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],23 AS [画像No],'自社画像23' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],24 AS [画像No],'自社画像24' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],25 AS [画像No],'自社画像25' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],26 AS [画像No],'自社画像26' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],27 AS [画像No],'自社画像27' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],28 AS [画像No],'自社画像28' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],29 AS [画像No],'自社画像29' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],30 AS [画像No],'自社画像30' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],31 AS [画像No],'自社画像31' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],32 AS [画像No],'自社画像32' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],33 AS [画像No],'自社画像33' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],34 AS [画像No],'自社画像34' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],35 AS [画像No],'自社画像35' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],36 AS [画像No],'自社画像36' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],37 AS [画像No],'自社画像37' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],38 AS [画像No],'自社画像38' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],39 AS [画像No],'自社画像39' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],40 AS [画像No],'自社画像40' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],41 AS [画像No],'自社画像41' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],42 AS [画像No],'自社画像42' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],43 AS [画像No],'自社画像43' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],44 AS [画像No],'自社画像44' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],45 AS [画像No],'自社画像45' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],46 AS [画像No],'自社画像46' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],47 AS [画像No],'自社画像47' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],48 AS [画像No],'自社画像48' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],49 AS [画像No],'自社画像49' AS [画像名] UNION ";
                tmp_sql = tmp_sql + " 		SELECT 8 AS [画像区分],50 AS [画像No],'自社画像50' AS [画像名] ";
                tmp_sql = tmp_sql + " 	) AS JISYADEF ";
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

}