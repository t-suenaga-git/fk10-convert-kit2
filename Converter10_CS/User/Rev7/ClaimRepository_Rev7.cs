using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using Converter10.Njc.Common;

namespace Converter10.Njc.Repository
{

    #region クレーム基本情報

    public class Claim_data_Base_Repository
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

                sortstr = "[クレームNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[クレームNo],[対応状況],[タイトル],[受付日時],[受付担当者],[連絡者名],[連絡者TEL1],[連絡者TEL2],[連絡可能時間（開始）],[連絡可能時間（終了）],[緊急度],[対応期限],[箇所分類No],[クレーム分類No],[内容],[対応担当者],[電子メール備考],[物件No],[部屋番号],[契約No],[契約レコードNo],[対応先区分],[依頼業者No],[業者担当者],[業者担当者TEL],[完了日時],[結果入力],[負担者1],[負担者1金額等],[負担者2],[負担者2金額等],[負担者3],[負担者3金額等],[備考],[自社・支店No],[自社・支店名],[受付時刻],[対応終了時刻],[連絡者名SJIS],[対応担当者名SJIS],[家主No]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 CL.id_no AS [クレームNo] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN syori_kbn = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN syori_kbn = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN syori_kbn = 2 THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN syori_kbn = 3 THEN 4 ";
                tmp_sql = tmp_sql + " 		 END AS [対応状況] ";
                tmp_sql = tmp_sql + " 		/*20160905 クレームタイトル取得内容修正_2 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		/*20160829 クレームタイトル取得内容修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,'受付No' + CONVERT(VARCHAR(50),CL.id_no) AS [タイトル]*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kasho_rui <> '' AND claim_rui <> '' THEN '箇所分類：' + kasho_rui + ' ' + 'クレーム分類：' + claim_rui ";
                tmp_sql = tmp_sql + " 			WHEN kasho_rui <> '' AND claim_rui = '' THEN '箇所分類：' + kasho_rui ";
                tmp_sql = tmp_sql + " 			WHEN kasho_rui = '' AND claim_rui <> '' THEN 'クレーム分類：' + claim_rui ";
                tmp_sql = tmp_sql + " 			WHEN kasho_rui = '' AND claim_rui = '' AND hy_no = '' THEN '物件No' + CONVERT(VARCHAR,CL.bk_no) ";
                tmp_sql = tmp_sql + " 			WHEN kasho_rui = '' AND claim_rui = '' AND hy_no <> '' THEN '物件No' + CONVERT(VARCHAR,CL.bk_no) + ' ' +'部屋No' + LTRIM(hy_no) ";
                tmp_sql = tmp_sql + " 		 END AS [タイトル] ";
                tmp_sql = tmp_sql + " 		/*20160829 クレームタイトル取得内容修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN claim_rui <> '' AND kasho_rui <> '' THEN claim_rui + '：' + kasho_rui ";
                tmp_sql = tmp_sql + " 			WHEN claim_rui = '' AND kasho_rui <> '' THEN '未設定' + '：' + kasho_rui ";
                tmp_sql = tmp_sql + " 			WHEN claim_rui <> '' AND kasho_rui = '' THEN claim_rui + '：' + '未設定' ";
                tmp_sql = tmp_sql + " 			WHEN claim_rui = '' AND kasho_rui = '' AND hy_no = '' THEN '物件No' + CONVERT(VARCHAR,CL.bk_no) ";
                tmp_sql = tmp_sql + " 			WHEN claim_rui = '' AND kasho_rui = '' AND hy_no <> '' THEN '物件No' + CONVERT(VARCHAR,CL.bk_no) + ' ' +'部屋No' + LTRIM(hy_no) ";
                tmp_sql = tmp_sql + " 		 END AS [タイトル] ";
                tmp_sql = tmp_sql + " 		/*20160905 クレームタイトル取得内容修正_2 chg end*/ ";
                tmp_sql = tmp_sql + " 		,uke_datetime AS [受付日時] ";
                tmp_sql = tmp_sql + " 		,uke_tanto_no AS [受付担当者] ";
                tmp_sql = tmp_sql + " 		,houkoku_name AS [連絡者名] ";
                tmp_sql = tmp_sql + " 		,houkoku_tel AS [連絡者TEL1] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡者TEL2] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡可能時間（開始）] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡可能時間（終了）] ";
                tmp_sql = tmp_sql + " 		,2 AS [緊急度] ";
                tmp_sql = tmp_sql + " 		,'' AS [対応期限] ";
                tmp_sql = tmp_sql + " 		,kasho_rui AS [箇所分類No] ";
                tmp_sql = tmp_sql + " 		,claim_rui AS [クレーム分類No] ";
                tmp_sql = tmp_sql + " 		,REPLACE(claim_report,'@CRLF@','$0D$0A') AS [内容] ";
                tmp_sql = tmp_sql + " 		,'' AS [対応担当者] ";
                tmp_sql = tmp_sql + " 		,'' AS [電子メール備考] ";
                tmp_sql = tmp_sql + " 		,CL.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋番号] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約レコードNo] ";
                tmp_sql = tmp_sql + " 		/*,'' AS [リフォームGuid]*/		/*修繕データと関連するため要考慮*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN result_taio = '業者依頼' THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN result_taio = '自社対応' THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN result_taio = '家主対応' THEN 3 ";
                tmp_sql = tmp_sql + " 			ELSE 4 ";
                tmp_sql = tmp_sql + " 		 END AS [対応先区分]			/*紐付の場合は要修正*/ ";
                tmp_sql = tmp_sql + " 		,result_gyo_no AS [依頼業者No] ";
                tmp_sql = tmp_sql + " 		,result_gyo_tanto AS [業者担当者] ";
                tmp_sql = tmp_sql + " 		,'' AS [業者担当者TEL] ";
                tmp_sql = tmp_sql + " 		,result_datetime AS [完了日時] ";
                tmp_sql = tmp_sql + " 			,REPLACE(result_report,'@CRLF@','$0D$0A') AS [結果入力] ";
                tmp_sql = tmp_sql + " 		,result_hutan_kbn AS [負担者1]	/*革命10と同一*/ ";
                tmp_sql = tmp_sql + " 		,result_gak_etc AS [負担者1金額等] ";
                tmp_sql = tmp_sql + " 		,result_hutan_kbn2 AS [負担者2]	/*革命10と同一*/ ";
                tmp_sql = tmp_sql + " 		,result_gak_etc2 AS [負担者2金額等] ";
                tmp_sql = tmp_sql + " 		,result_hutan_kbn3 AS [負担者3]	/*革命10と同一*/ ";
                tmp_sql = tmp_sql + " 		,result_gak_etc3 AS [負担者3金額等] ";
                tmp_sql = tmp_sql + " 		,result_biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [自社・支店No] ";
                tmp_sql = tmp_sql + " 		,'' AS [自社・支店名] ";
                tmp_sql = tmp_sql + " 		,RIGHT('0' + CONVERT(VARCHAR,DATEPART(HOUR,uke_datetime)),2) + ':' + RIGHT('0' + CONVERT(VARCHAR,DATEPART(MINUTE,uke_datetime)),2) AS [受付時刻] ";
                tmp_sql = tmp_sql + " 		,RIGHT('0' + CONVERT(VARCHAR,DATEPART(HOUR,result_datetime)),2) + ':' + RIGHT('0' + CONVERT(VARCHAR,DATEPART(MINUTE,result_datetime)),2) AS [対応終了時刻] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡者名SJIS] ";
                tmp_sql = tmp_sql + " 		,'' AS [対応担当者名SJIS] ";
                tmp_sql = tmp_sql + " 		,so_no_kasi AS [家主No] ";
                tmp_sql = tmp_sql + " 	FROM claim_data AS CL ";
                tmp_sql = tmp_sql + " 	LEFT JOIN claim_subdata AS CLSUB ON CL.id_no = CLSUB.id_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN (SELECT kanren_claim_id,reform_id FROM reform_data) AS REF ON CL.id_no = REF.kanren_claim_id ";
                tmp_sql = tmp_sql + " 	LEFT JOIN (SELECT bk_no,so_no_kasi,so_no_kasi2,so_no_syo,so_no_syo2 FROM bk_mst) AS BK ON CL.bk_no = BK.bk_no ";
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

    #region クレーム対応履歴情報

    public class Claim_data_taiorireki_Repository
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

                sortstr = "[クレームNo],[対応履歴No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[クレームNo],[対応履歴No],[対応日時],[対応担当者],[内容],[対応時刻]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 id_no AS [クレームNo] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY id_no ORDER BY sort_no) AS [対応履歴No] ";
                tmp_sql = tmp_sql + " 		,taisho1_date AS [対応日時] ";
                tmp_sql = tmp_sql + " 		,-1 AS [対応担当者] ";
                tmp_sql = tmp_sql + " 		,REPLACE(taisho1_report,'@CRLF@','$0D$0A') AS [内容]	 ";
                tmp_sql = tmp_sql + " 		,taisho1_time AS [対応時刻] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,1 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho1_date,taisho1_time,taisho1_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho1_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho1_time),'') + RTRIM(LTRIM(ISNULL(taisho1_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,2 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho2_date,taisho2_time,taisho2_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho2_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho2_time),'') + RTRIM(LTRIM(ISNULL(taisho2_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,3 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho3_date,taisho3_time,taisho3_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho3_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho3_time),'') + RTRIM(LTRIM(ISNULL(taisho3_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,4 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho4_date,taisho4_time,taisho4_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho4_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho4_time),'') + RTRIM(LTRIM(ISNULL(taisho4_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,5 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho5_date,taisho5_time,taisho5_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho5_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho5_time),'') + RTRIM(LTRIM(ISNULL(taisho5_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,6 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho6_date,taisho6_time,taisho6_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho6_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho6_time),'') + RTRIM(LTRIM(ISNULL(taisho6_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,7 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho7_date,taisho7_time,taisho7_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho7_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho7_time),'') + RTRIM(LTRIM(ISNULL(taisho7_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,8 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho8_date,taisho8_time,taisho8_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho8_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho8_time),'') + RTRIM(LTRIM(ISNULL(taisho8_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,9 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho9_date,taisho9_time,taisho9_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho9_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho9_time),'') + RTRIM(LTRIM(ISNULL(taisho9_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 id_no ";
                tmp_sql = tmp_sql + " 			,10 AS sort_no ";
                tmp_sql = tmp_sql + " 			,taisho10_date,taisho10_time,taisho10_report ";
                tmp_sql = tmp_sql + " 		FROM claim_data ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(CONVERT(VARCHAR(50),taisho10_date),'') + ISNULL(CONVERT(VARCHAR(50),taisho10_time),'') + RTRIM(LTRIM(ISNULL(taisho10_report,''))) <> '' ";
                tmp_sql = tmp_sql + " 	) AS CLTAIO ";
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

    #region クレーム関連ファイル情報

    public class Claim_data_relfile_Repository
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

                sortstr = "[クレームNo],[ファイルNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[クレームNo],[ファイルNo],[ファイルパス],[追加日],[備考]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 id_no AS [クレームNo] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY id_no ORDER BY sort_no) -1 AS [ファイルNo] ";
                tmp_sql = tmp_sql + " 		,rel_file1 AS [ファイルパス] ";
                tmp_sql = tmp_sql + " 		,CONVERT (date, GETDATE()) AS [追加日] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT id_no,1 AS [sort_no],rel_file1 FROM claim_data WHERE RTRIM(LTRIM(ISNULL(rel_file1,''))) <> '' UNION ";
                tmp_sql = tmp_sql + " 		SELECT id_no,2 AS [sort_no],rel_file2 FROM claim_data WHERE RTRIM(LTRIM(ISNULL(rel_file2,''))) <> '' UNION ";
                tmp_sql = tmp_sql + " 		SELECT id_no,3 AS [sort_no],rel_file3 FROM claim_data WHERE RTRIM(LTRIM(ISNULL(rel_file3,''))) <> '' UNION ";
                tmp_sql = tmp_sql + " 		SELECT id_no,4 AS [sort_no],rel_file4 FROM claim_data WHERE RTRIM(LTRIM(ISNULL(rel_file4,''))) <> '' UNION ";
                tmp_sql = tmp_sql + " 		SELECT id_no,5 AS [sort_no],rel_file5 FROM claim_data WHERE RTRIM(LTRIM(ISNULL(rel_file5,''))) <> '' ";
                tmp_sql = tmp_sql + " 	) AS RELTOTAL ";
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