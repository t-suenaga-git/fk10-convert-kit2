using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using Converter10.Njc.Common;

namespace Converter10.Njc.Repository
{

    #region 修繕基本情報

    public class Reform_data_Base_Repository
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

                sortstr = "[物件No],[修繕No],[部屋No],[契約No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[修繕No],[部屋使用フラグ],[部屋No],[契約使用フラグ],[契約No],[修繕名],[修繕受付日],[修繕終了日],[修繕担当者],[自社支店No],[工事開始予定日],[工事終了予定日],[工事場所],[工事概要],[鍵],[立会者],[備考],[契約者負担率],[家主負担率],[自社負担率],[契約者請求作成フラグ],[契約者請求先No],[契約者該当年月],[契約者請求締日],[契約者入金区分],[契約者送金率],[契約者振込先口座No],[契約者請求備考],[家主請求作成フラグ],[家主控除先No],[家主控除請求区分],[家主控除予定日],[家主控除該当年月],[家主控除先送金ルールNo],[家主控除先口座No],[家主請求書発行予定日],[家主請求該当年月],[家主請求締日],[家主請求入金区分],[家主請求振込先口座No],[家主請求先No],[家主控除請求備考]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 REFDATA.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY REFDATA.bk_no ORDER BY REFDATA.hy_composite,REFDATA.reform_id) AS [修繕No] ";
                tmp_sql = tmp_sql + " 		,NULL AS [部屋使用フラグ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.hy_composite = '999999999' THEN NULL ";
                tmp_sql = tmp_sql + " 			ELSE REFDATA.hy_composite ";
                tmp_sql = tmp_sql + " 		 END AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,NULL AS [契約使用フラグ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.ky_no = 0 THEN NULL ";
                tmp_sql = tmp_sql + " 			ELSE REFDATA.ky_no ";
                tmp_sql = tmp_sql + " 		 END AS [契約No] ";
                tmp_sql = tmp_sql + " 		,reform_name AS [修繕名] ";
                tmp_sql = tmp_sql + " 		,[start_date] AS [修繕受付日] ";
                tmp_sql = tmp_sql + " 		,NULL AS [修繕終了日] ";
                tmp_sql = tmp_sql + " 		,REFDATA.tanto_no AS [修繕担当者] ";
                tmp_sql = tmp_sql + " 		,'' AS [自社支店No] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN plan_start_date = '18991230' THEN NULL ";
                tmp_sql = tmp_sql + " 			ELSE plan_start_date ";
                tmp_sql = tmp_sql + " 		 END AS [工事開始予定日] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN plan_end_date = '18991230' THEN NULL ";
                tmp_sql = tmp_sql + " 			ELSE plan_end_date ";
                tmp_sql = tmp_sql + " 		 END AS [工事終了予定日]	   ";
                tmp_sql = tmp_sql + " 		,'' AS [工事場所] ";
                tmp_sql = tmp_sql + " 		,'' AS [工事概要] ";
                tmp_sql = tmp_sql + " 		,keyinfo AS [鍵] ";
                tmp_sql = tmp_sql + " 		,TANTO.tanto_name AS [立会者] ";
                tmp_sql = tmp_sql + " 		,REFDATA.biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,kihon_futan_rit_kari AS [契約者負担率] ";
                tmp_sql = tmp_sql + " 		,kihon_futan_rit_owner AS [家主負担率] ";
                tmp_sql = tmp_sql + " 		,kihon_futan_rit_jisha AS [自社負担率] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.hy_composite = '999999999' OR REFDATA.ky_no = 0 THEN 0 ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [契約者請求作成フラグ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.hy_composite = '999999999' OR REFDATA.ky_no = 0 THEN NULL ";
                tmp_sql = tmp_sql + " 			ELSE KYK.kys_no ";
                tmp_sql = tmp_sql + " 		 END AS [契約者請求先No]	 ";
                tmp_sql = tmp_sql + " 		,DATEADD(DAY,(DAY([start_date]) - 1 ) * (-1),[start_date]) AS [契約者該当年月] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN keijo_kbn = 0 THEN [start_date] ";
                tmp_sql = tmp_sql + " 			WHEN keijo_kbn = 1 THEN sq_date ";
                tmp_sql = tmp_sql + " 		 END AS [契約者請求締日] ";
                tmp_sql = tmp_sql + " 		,MNKBN.nkbn_name AS [契約者入金区分] ";
                tmp_sql = tmp_sql + " 		,so_rit_sogo AS [契約者送金率] ";
                tmp_sql = tmp_sql + " 		,BK.fkom_no AS [契約者振込先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約者請求備考] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN so_date = '18991230' THEN 0 ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [家主請求作成フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主控除先No] ";
                tmp_sql = tmp_sql + " 		,1 AS [家主控除請求区分] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN so_date = '18991230' THEN NULL ";
                tmp_sql = tmp_sql + " 			ELSE so_date ";
                tmp_sql = tmp_sql + " 		 END AS [家主控除予定日]	 ";
                tmp_sql = tmp_sql + " 		,DATEADD(DAY,(DAY([start_date]) - 1 ) * (-1),[start_date]) AS [家主控除該当年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主控除先送金ルールNo] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主控除先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主請求書発行予定日] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主請求該当年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主請求締日] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主請求入金区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主請求振込先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主請求先No] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主控除請求備考] ";
                tmp_sql = tmp_sql + " 	FROM reform_data AS REFDATA ";
                tmp_sql = tmp_sql + " 	LEFT JOIN bk_mst AS BK ON REFDATA.bk_no = BK.bk_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_tanto AS TANTO ON REFDATA.tanto_no = TANTO.tanto_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkbn AS MNKBN ON REFDATA.kari_kbn = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 	ON  REFDATA.bk_no = KYK.bk_no ";
                tmp_sql = tmp_sql + " 	AND REFDATA.hy_composite = KYK.hy_no ";
                tmp_sql = tmp_sql + " 	AND REFDATA.ky_no = KYK.ky_no ";
                tmp_sql = tmp_sql + " 	AND REFDATA.ko_no = KYK.ko_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/* ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 bk_no ";
                tmp_sql = tmp_sql + " 				,hy_composite ";
                tmp_sql = tmp_sql + " 				,reform_id ";
                tmp_sql = tmp_sql + " 				,SUM(ISNULL(total_kari,0)) AS total_kari ";
                tmp_sql = tmp_sql + " 				,SUM(ISNULL(kari_so_gak,0)) AS kari_so_gak ";
                tmp_sql = tmp_sql + " 				,CASE ";
                tmp_sql = tmp_sql + " 					WHEN SUM(ISNULL(total_kari,0)) = 0 THEN 0 ";
                tmp_sql = tmp_sql + " 					ELSE (SUM(ISNULL(kari_so_gak,0)) / SUM(ISNULL(total_kari,0))) * 100 ";
                tmp_sql = tmp_sql + " 				 END AS so_rit_sogo ";
                tmp_sql = tmp_sql + " 			FROM reform_item ";
                tmp_sql = tmp_sql + " 			GROUP BY bk_no,hy_composite,reform_id ";
                tmp_sql = tmp_sql + " 			*/ ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 bk_no ";
                tmp_sql = tmp_sql + " 				,hy_composite ";
                tmp_sql = tmp_sql + " 				,reform_id ";
                tmp_sql = tmp_sql + " 				,CASE ";
                tmp_sql = tmp_sql + " 					WHEN ISNULL(total_kari,0) = 0 THEN 0 ";
                tmp_sql = tmp_sql + " 					/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 					/*ELSE ROUND(((kari_so_gak + zei_gak_kari_so) / (total_kari + zei_gak_kari)) * 100,-1)*/ ";
                tmp_sql = tmp_sql + " 					ELSE ";
                tmp_sql = tmp_sql + " 						CASE ";
                tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND(((kari_so_gak + zei_gak_kari_so) / (total_kari + zei_gak_kari)) * 100,2,1)				/*切り捨て*/ ";
                tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND(((kari_so_gak + zei_gak_kari_so) / (total_kari + zei_gak_kari)) * 100,2)					/*四捨五入*/ ";
                tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING((((kari_so_gak + zei_gak_kari_so) / (total_kari + zei_gak_kari)) * 100) * 100) / 100	/*切り上げ*/ ";
                tmp_sql = tmp_sql + " 						END ";
                tmp_sql = tmp_sql + " 					/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 				 END AS so_rit_sogo ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 REFITEMSUM.bk_no ";
                tmp_sql = tmp_sql + " 				,REFITEMSUM.hy_composite ";
                tmp_sql = tmp_sql + " 				,REFITEMSUM.reform_id ";
                tmp_sql = tmp_sql + " 				,REFITEMSUM.total_kari ";
                tmp_sql = tmp_sql + " 				,CASE ";
                tmp_sql = tmp_sql + " 					WHEN REFITEMTAX.bk_no IS NULL THEN REFITEMSUM.zei_gak_kari ";
                tmp_sql = tmp_sql + " 					ELSE REFITEMTAX.zei_gak_kari ";
                tmp_sql = tmp_sql + " 				 END AS zei_gak_kari ";
                tmp_sql = tmp_sql + " 				 ,REFITEMSUM.kari_so_gak ";
                tmp_sql = tmp_sql + " 				,CASE ";
                tmp_sql = tmp_sql + " 					WHEN REFITEMTAX.bk_no IS NULL THEN 0 ";
                tmp_sql = tmp_sql + " 					ELSE REFITEMTAX.kari_so_zeigak ";
                tmp_sql = tmp_sql + " 				 END AS zei_gak_kari_so	 ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				/*合計の総和と送金額の総和を抽出 sta*/ ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_composite ";
                tmp_sql = tmp_sql + " 					,reform_id ";
                tmp_sql = tmp_sql + " 					,SUM(ISNULL(total_kari,0)) AS total_kari ";
                tmp_sql = tmp_sql + " 					,SUM(ISNULL(zei_gak_kari,0)) AS zei_gak_kari ";
                tmp_sql = tmp_sql + " 					,SUM(ISNULL(kari_so_gak,0)) AS kari_so_gak ";
                tmp_sql = tmp_sql + " 				FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM reform_item ";
                tmp_sql = tmp_sql + " 					WHERE nkin_no <> 8999 ";
                tmp_sql = tmp_sql + " 				) AS VW1 ";
                tmp_sql = tmp_sql + " 				GROUP BY bk_no,hy_composite,reform_id ";
                tmp_sql = tmp_sql + " 				/*合計の総和と送金額の総和を抽出 end*/ ";
                tmp_sql = tmp_sql + " 			) AS REFITEMSUM ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM reform_item AS REFITEM ";
                tmp_sql = tmp_sql + " 					WHERE nkin_no = 8999 ";
                tmp_sql = tmp_sql + " 				) AS REFITEMTAX ";
                tmp_sql = tmp_sql + " 			ON  REFITEMSUM.bk_no = REFITEMTAX.bk_no ";
                tmp_sql = tmp_sql + " 			AND REFITEMSUM.hy_composite = REFITEMTAX.hy_composite ";
                tmp_sql = tmp_sql + " 			AND REFITEMSUM.reform_id = REFITEMTAX.reform_id ";
                tmp_sql = tmp_sql + " 			) AS VW ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		) AS SORIT ";
                tmp_sql = tmp_sql + " 	ON  REFDATA.bk_no = SORIT.bk_no ";
                tmp_sql = tmp_sql + " 	AND REFDATA.hy_composite = SORIT.hy_composite ";
                tmp_sql = tmp_sql + " 	AND REFDATA.reform_id = SORIT.reform_id ";
                tmp_sql = tmp_sql + " 	WHERE REFDATA.reform_kbn = 1	/*リフォームを抽出*/ ";
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

    #region 修繕見積情報

    public class Reform_data_Mitumori_Repository
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

                sortstr = "[物件No],[修繕No],[見積No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[修繕No],[見積No],[見積タイトル],[見積番号],[見積日],[成約フラグ],[成約日],[負担区分(全体・個別)],[契約者負担率(全体用)],[家主負担率(全体用)],[自社負担率(全体用)],[税適用区分(全体・個別)],[税区分(全体用)],[適用税率],[契約者税額(全体用)],[家主税額(全体用)],[自社税額(全体用)],[端数負担者区分]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 REFDATA.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,REFDATA.[修繕No] ";
                tmp_sql = tmp_sql + " 		,1 AS [見積No] ";
                tmp_sql = tmp_sql + " 		,'' AS [見積タイトル] ";
                tmp_sql = tmp_sql + " 		,'' AS [見積番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [見積日] ";
                tmp_sql = tmp_sql + " 		/*20160629 修繕検証後修正 chg sta 成約フラグのチェックをONにする*/ ";
                tmp_sql = tmp_sql + " 		,1 AS [成約フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [成約日] ";
                tmp_sql = tmp_sql + " 		,1 AS [負担区分(全体・個別)] ";
                tmp_sql = tmp_sql + " 		,kihon_futan_rit_kari AS [契約者負担率(全体用)] ";
                tmp_sql = tmp_sql + " 		,kihon_futan_rit_owner AS [家主負担率(全体用)] ";
                tmp_sql = tmp_sql + " 		,kihon_futan_rit_jisha AS [自社負担率(全体用)] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN tax_kbn = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN tax_kbn = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 		 END AS [税適用区分(全体・個別)] ";
                tmp_sql = tmp_sql + " 		 ,CASE ";
                tmp_sql = tmp_sql + " 			WHEN tax = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN tax = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		 END AS [税区分(全体用)] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [start_date] BETWEEN '19890401' AND '19970331' THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN [start_date] BETWEEN '19970401' AND '20140331' THEN 5 ";
                tmp_sql = tmp_sql + " 			WHEN [start_date] BETWEEN '20140401' AND '99991231' THEN 8 ";
                tmp_sql = tmp_sql + " 			ELSE 8 ";
                tmp_sql = tmp_sql + " 		 END AS [適用税率] ";
                tmp_sql = tmp_sql + " 		,zei_gak_kari AS [契約者税額(全体用)] ";
                tmp_sql = tmp_sql + " 		,zei_gak_owner AS [家主税額(全体用)] ";
                tmp_sql = tmp_sql + " 		,zei_gak_jisha AS [自社税額(全体用)] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kihon_futan_hasuuke = 0 THEN 100 ";
                tmp_sql = tmp_sql + " 			WHEN kihon_futan_hasuuke = 1 THEN 200 ";
                tmp_sql = tmp_sql + " 			WHEN kihon_futan_hasuuke = 2 THEN 900 ";
                tmp_sql = tmp_sql + " 		 END AS [端数負担者区分] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No],* FROM reform_data ";
                tmp_sql = tmp_sql + " 		) AS REFDATA ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*合計金額に対して消費税が設定されているデータ sta*/ ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 bk_no ";
                tmp_sql = tmp_sql + " 				,hy_composite ";
                tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 del*/ ";
                tmp_sql = tmp_sql + " 				/*,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No]*/ ";
                tmp_sql = tmp_sql + " 				,reform_id ";
                tmp_sql = tmp_sql + " 				,zei_gak_kari ";
                tmp_sql = tmp_sql + " 				,zei_gak_owner ";
                tmp_sql = tmp_sql + " 				,zei_gak_jisha ";
                tmp_sql = tmp_sql + " 			FROM reform_item AS REFITEM ";
                tmp_sql = tmp_sql + " 			WHERE EXISTS ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM reform_data AS REFDATA ";
                tmp_sql = tmp_sql + " 					WHERE REFITEM.bk_no = REFDATA.bk_no ";
                tmp_sql = tmp_sql + " 					AND   REFITEM.hy_composite = REFDATA.hy_composite ";
                tmp_sql = tmp_sql + " 					AND   REFITEM.reform_id = REFDATA.reform_id ";
                tmp_sql = tmp_sql + " 					AND   REFDATA.tax = 1 ";
                tmp_sql = tmp_sql + " 					AND   REFDATA.tax_kbn = 0 ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 			AND nkin_no = 8999 ";
                tmp_sql = tmp_sql + " 			/*合計金額に対して消費税が設定されているデータ end*/ ";
                tmp_sql = tmp_sql + " 			/*項目毎に対して消費税が設定されているデータ sta*/ ";
                tmp_sql = tmp_sql + " 			UNION ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 bk_no ";
                tmp_sql = tmp_sql + " 				,hy_composite ";
                tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 del*/ ";
                tmp_sql = tmp_sql + " 				/*,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No]*/ ";
                tmp_sql = tmp_sql + " 				,reform_id ";
                tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 				/* ";
                tmp_sql = tmp_sql + " 				,zei_gak_kari ";
                tmp_sql = tmp_sql + " 				,zei_gak_owner ";
                tmp_sql = tmp_sql + " 				,zei_gak_jisha ";
                tmp_sql = tmp_sql + " 				*/ ";
                tmp_sql = tmp_sql + " 				,SUM(ISNULL(zei_gak_kari,0)) AS zei_gak_kari ";
                tmp_sql = tmp_sql + " 				,SUM(ISNULL(zei_gak_owner,0)) AS zei_gak_owner ";
                tmp_sql = tmp_sql + " 				,SUM(ISNULL(zei_gak_jisha,0)) AS zei_gak_jisha ";
                tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 			FROM reform_item AS REFITEM ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 add*/ ";
                tmp_sql = tmp_sql + " 			GROUP BY bk_no,hy_composite,reform_id ";
                tmp_sql = tmp_sql + " 			HAVING EXISTS	/*20160629 修繕検証後修正 chg (WHERE → HAVINGへ変更)*/ ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM reform_data AS REFDATA ";
                tmp_sql = tmp_sql + " 					WHERE REFITEM.bk_no = REFDATA.bk_no ";
                tmp_sql = tmp_sql + " 					AND   REFITEM.hy_composite = REFDATA.hy_composite ";
                tmp_sql = tmp_sql + " 					AND   REFITEM.reform_id = REFDATA.reform_id ";
                tmp_sql = tmp_sql + " 					AND   REFDATA.tax = 1 ";
                tmp_sql = tmp_sql + " 					AND   REFDATA.tax_kbn = 1 ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 			/*項目毎に対して消費税が設定されているデータ end*/ ";
                tmp_sql = tmp_sql + " 			/*消費税なしのデータ sta*/ ";
                tmp_sql = tmp_sql + " 			UNION ";
                tmp_sql = tmp_sql + " 			SELECT DISTINCT		/*20160629 修繕検証後修正 add (DISTINCTを追加)*/ ";
                tmp_sql = tmp_sql + " 				 bk_no ";
                tmp_sql = tmp_sql + " 				,hy_composite ";
                tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 del*/ ";
                tmp_sql = tmp_sql + " 				/*,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No]*/ ";
                tmp_sql = tmp_sql + " 				,reform_id ";
                tmp_sql = tmp_sql + " 				,NULL AS zei_gak_kari ";
                tmp_sql = tmp_sql + " 				,NULL AS zei_gak_owner ";
                tmp_sql = tmp_sql + " 				,NULL AS zei_gak_jisha ";
                tmp_sql = tmp_sql + " 			FROM reform_item AS REFITEM ";
                tmp_sql = tmp_sql + " 			WHERE EXISTS ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM reform_data AS REFDATA ";
                tmp_sql = tmp_sql + " 					WHERE REFITEM.bk_no = REFDATA.bk_no ";
                tmp_sql = tmp_sql + " 					AND   REFITEM.hy_composite = REFDATA.hy_composite ";
                tmp_sql = tmp_sql + " 					AND   REFITEM.reform_id = REFDATA.reform_id ";
                tmp_sql = tmp_sql + " 					AND   REFDATA.tax = 0 ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 			/*消費税なしのデータ end*/ ";
                tmp_sql = tmp_sql + " 		) AS ZEITOTAL ";
                tmp_sql = tmp_sql + " 	ON  REFDATA.bk_no = ZEITOTAL.bk_no ";
                tmp_sql = tmp_sql + " 	AND REFDATA.hy_composite = ZEITOTAL.hy_composite ";
                tmp_sql = tmp_sql + " 	AND REFDATA.reform_id = ZEITOTAL.reform_id ";
                tmp_sql = tmp_sql + " 	WHERE REFDATA.reform_kbn = 1	/*リフォームを抽出*/ ";
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

    #region 修繕見積詳細情報

    public class Reform_item_Repository
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

                sortstr = "[物件No],[修繕No],[見積No],[見積明細No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[修繕No],[見積No],[見積明細No],[修繕項目名],[摘要],[見積数量],[見積単位],[見積単価],[見積税区分],[見積税額(税入力用)],[契約者負担率(個別用)],[家主負担率(個別用)],[自社負担率(個別用)],[発注業者No],[実行数量],[実行単位],[実行単価],[実行税区分],[実行税額(税入力用)]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 REFITEM.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,[修繕No] ";
                tmp_sql = tmp_sql + " 		,1 AS [見積No] ";
                tmp_sql = tmp_sql + " 		,item_id + 1 AS [見積明細No] ";
                tmp_sql = tmp_sql + " 		/*,REFITEM.nkin_no AS [修繕項目名]*/ ";
                tmp_sql = tmp_sql + " 		,MN.nkin_name AS [修繕項目名] ";
                tmp_sql = tmp_sql + " 		,REFITEM.biko AS [摘要] ";
                tmp_sql = tmp_sql + " 		,amount_plan AS [見積数量] ";
                tmp_sql = tmp_sql + " 		,unit_name AS [見積単位] ";
                tmp_sql = tmp_sql + " 		,price_plan AS [見積単価] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 0 THEN 1						/*消費税無し*/ ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 1 AND tax_kbn = 0 THEN 1		/*消費税有りかつ合計額に設定*/ ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*WHEN REFDATA.tax = 1 AND tax_kbn = 1 THEN 2		/*消費税有りかつ項目毎に設定*/*/ ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 1 AND tax_kbn = 1 THEN		/*消費税有りかつ項目毎に設定*/ ";
                tmp_sql = tmp_sql + " 				CASE ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_plan = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_plan = 1 THEN 3 ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_plan = 2 THEN 2 ";
                tmp_sql = tmp_sql + " 				END ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		 END AS [見積税区分] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 0 THEN 0									/*消費税無し*/ ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 1 AND tax_kbn = 0 THEN 0					/*消費税有りかつ合計額に設定*/ ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*WHEN REFDATA.tax = 1 AND tax_kbn = 1 THEN zei_gak_plan		/*消費税有りかつ項目毎に設定*/*/ ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 1 AND tax_kbn = 1 THEN					/*消費税有りかつ項目毎に設定*/ ";
                tmp_sql = tmp_sql + " 				CASE ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_plan = 0 THEN NULL ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_plan = 1 THEN zei_gak_plan ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_plan = 2 THEN zei_gak_plan ";
                tmp_sql = tmp_sql + " 				END ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		 END AS [見積税額(税入力用)]	 ";
                tmp_sql = tmp_sql + " 		/*20160627 修繕追加分移行修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*負担率が存在しない場合の対応*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,futan_rit_kari AS [契約者負担率(個別用)] ";
                tmp_sql = tmp_sql + " 		,futan_rit_owner AS [家主負担率(個別用)] ";
                tmp_sql = tmp_sql + " 		,futan_rit_jisha AS [自社負担率(個別用)] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(futan_rit_kari,0) <> 0 THEN futan_rit_kari ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*ELSE ROUND((total_kari / total_plan) * 100,0)*/ ";
                tmp_sql = tmp_sql + " 			ELSE ";
                tmp_sql = tmp_sql + " 				CASE ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((total_kari / total_plan) * 100,2,1)				/*切り捨て*/ ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((total_kari / total_plan) * 100,2)				/*四捨五入*/ ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((total_kari / total_plan) * 100) * 100) / 100	/*切り上げ*/ ";
                tmp_sql = tmp_sql + " 				END ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		 END AS [契約者負担率(個別用)] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(futan_rit_owner,0) <> 0 THEN futan_rit_owner ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*ELSE ROUND((total_owner / total_plan) * 100,0)*/ ";
                tmp_sql = tmp_sql + " 			ELSE ";
                tmp_sql = tmp_sql + " 				CASE ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((total_owner / total_plan) * 100,2,1)				/*切り捨て*/ ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((total_owner / total_plan) * 100,2)				/*四捨五入*/ ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((total_owner / total_plan) * 100) * 100) / 100	/*切り上げ*/ ";
                tmp_sql = tmp_sql + " 				END ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		 END AS [家主負担率(個別用)] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(futan_rit_jisha,0) <> 0 THEN futan_rit_jisha ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*ELSE ROUND((total_jisha / total_plan) * 100,0)*/ ";
                tmp_sql = tmp_sql + " 			ELSE ";
                tmp_sql = tmp_sql + " 				CASE ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((total_jisha / total_plan) * 100,2,1)				/*切り捨て*/ ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((total_jisha / total_plan) * 100,2)				/*四捨五入*/ ";
                tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((total_jisha / total_plan) * 100) * 100) / 100	/*切り上げ*/ ";
                tmp_sql = tmp_sql + " 				END ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		 END AS [自社負担率(個別用)] ";
                tmp_sql = tmp_sql + " 		/*20160627 修繕追加分移行修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		,repare_no AS [発注業者No] ";
                tmp_sql = tmp_sql + " 		,amount_cost AS [実行数量] ";
                tmp_sql = tmp_sql + " 		,REFITEM.yobi_mj1 AS [実行単位] ";
                tmp_sql = tmp_sql + " 		,price_cost AS [実行単価] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 0 THEN 1						/*消費税無し*/ ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 1 AND tax_kbn = 0 THEN 1		/*消費税有りかつ合計額に設定*/ ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*WHEN REFDATA.tax = 1 AND tax_kbn = 1 THEN 2		/*消費税有りかつ項目毎に設定*/*/ ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 1 AND tax_kbn = 1 THEN		/*消費税有りかつ項目毎に設定*/ ";
                tmp_sql = tmp_sql + " 				CASE ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_cost = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_cost = 1 THEN 3 ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_cost = 2 THEN 2 ";
                tmp_sql = tmp_sql + " 				END			 ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		 END AS [実行税区分] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 0 THEN 0									/*消費税無し*/ ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 1 AND tax_kbn = 0 THEN 0					/*消費税有りかつ合計額に設定*/ ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 			/*WHEN REFDATA.tax = 1 AND tax_kbn = 1 THEN zei_gak_cost		/*消費税有りかつ項目毎に設定*/*/ ";
                tmp_sql = tmp_sql + " 			WHEN REFDATA.tax = 1 AND tax_kbn = 1 THEN					/*消費税有りかつ項目毎に設定*/ ";
                tmp_sql = tmp_sql + " 				CASE ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_cost = 0 THEN NULL ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_cost = 1 THEN zei_gak_cost ";
                tmp_sql = tmp_sql + " 					WHEN zei_kbn_cost = 2 THEN zei_gak_cost ";
                tmp_sql = tmp_sql + " 				END			 ";
                tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		 END AS [実行税額(税入力用)]		 ";
                tmp_sql = tmp_sql + " 	FROM reform_item AS REFITEM ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No],* FROM reform_data ";
                tmp_sql = tmp_sql + " 		) AS REFDATA ";
                tmp_sql = tmp_sql + " 	ON  REFITEM.bk_no = REFDATA.bk_no ";
                tmp_sql = tmp_sql + " 	AND REFITEM.hy_composite = REFDATA.hy_composite ";
                tmp_sql = tmp_sql + " 	AND REFITEM.reform_id = REFDATA.reform_id ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN ON REFITEM.nkin_no = MN.nkin_no ";
                tmp_sql = tmp_sql + " 	WHERE reform_kbn = 1			/*リフォームを抽出*/ ";
                tmp_sql = tmp_sql + " 	AND   REFITEM.nkin_no <> 8999	/*合計額に消費税が設定されている場合の消費税レコードは抽出対象外*/ ";
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

    #region 修繕クレーム関連付け情報

    public class Reform_data_claim_Repository
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

                sortstr = "[物件No],[修繕No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[修繕No],[クレームNo],[クレーム画面での並び順],[修繕情報登録画面での並び順]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 			,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No] ";
                tmp_sql = tmp_sql + " 			,kanren_claim_id AS [クレームNo] ";
                tmp_sql = tmp_sql + " 			,1 AS [クレーム画面での並び順] ";
                tmp_sql = tmp_sql + " 			,999999999 AS [修繕情報登録画面での並び順] ";
                tmp_sql = tmp_sql + " 		FROM reform_data ";
                tmp_sql = tmp_sql + " 		/*20160629 修繕検証後修正 del*/ ";
                tmp_sql = tmp_sql + " 		/*WHERE kanren_claim_id IS NOT NULL*/ ";
                tmp_sql = tmp_sql + " 	) AS TOTAL ";
                tmp_sql = tmp_sql + " 	/*20160629 修繕検証後修正 add*/ ";
                tmp_sql = tmp_sql + " 	WHERE クレームNo IS NOT NULL ";
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

    #region 修繕関連ファイル情報

    public class Reform_data_relfile_Repository
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

                sortstr = "[物件No],[修繕No],[ファイルNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[修繕No],[ファイルNo],[ファイルパス],[追加日],[備考]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,[修繕No] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY bk_no,[修繕No] ORDER BY sort_no) -1 AS [ファイルNo] ";
                tmp_sql = tmp_sql + " 		,rel_file1 AS [ファイルパス] ";
                tmp_sql = tmp_sql + " 		,CONVERT (date, GETDATE()) AS [追加日] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考]	 ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 REFDATABASE.bk_no ";
                tmp_sql = tmp_sql + " 				,[修繕No] ";
                tmp_sql = tmp_sql + " 				,reform_kbn ";
                tmp_sql = tmp_sql + " 				,1 AS [sort_no] ";
                tmp_sql = tmp_sql + " 				,rel_file1 ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_composite ";
                tmp_sql = tmp_sql + " 					,reform_id ";
                tmp_sql = tmp_sql + " 					,reform_kbn ";
                tmp_sql = tmp_sql + " 					,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No] ";
                tmp_sql = tmp_sql + " 				FROM reform_data ";
                tmp_sql = tmp_sql + " 			) AS REFDATABASE ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_composite ";
                tmp_sql = tmp_sql + " 						,reform_id ";
                tmp_sql = tmp_sql + " 						,rel_file1,rel_file2,rel_file3,rel_file4,rel_file5 ";
                tmp_sql = tmp_sql + " 					FROM reform_data		 ";
                tmp_sql = tmp_sql + " 				) AS REFDATARELFILE ";
                tmp_sql = tmp_sql + " 			ON  REFDATABASE.bk_no = REFDATARELFILE.bk_no ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.hy_composite = REFDATARELFILE.hy_composite ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.reform_id = REFDATARELFILE.reform_id ";
                tmp_sql = tmp_sql + " 			WHERE RTRIM(LTRIM(ISNULL(rel_file1,''))) <> '' ";
                tmp_sql = tmp_sql + " 		) AS RELFILE1 ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 REFDATABASE.bk_no ";
                tmp_sql = tmp_sql + " 				,[修繕No] ";
                tmp_sql = tmp_sql + " 				,reform_kbn ";
                tmp_sql = tmp_sql + " 				,2 AS [sort_no] ";
                tmp_sql = tmp_sql + " 				,rel_file2 ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_composite ";
                tmp_sql = tmp_sql + " 					,reform_id ";
                tmp_sql = tmp_sql + " 					,reform_kbn ";
                tmp_sql = tmp_sql + " 					,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No] ";
                tmp_sql = tmp_sql + " 				FROM reform_data ";
                tmp_sql = tmp_sql + " 			) AS REFDATABASE ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_composite ";
                tmp_sql = tmp_sql + " 						,reform_id ";
                tmp_sql = tmp_sql + " 						,rel_file1,rel_file2,rel_file3,rel_file4,rel_file5 ";
                tmp_sql = tmp_sql + " 					FROM reform_data		 ";
                tmp_sql = tmp_sql + " 				) AS REFDATARELFILE ";
                tmp_sql = tmp_sql + " 			ON  REFDATABASE.bk_no = REFDATARELFILE.bk_no ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.hy_composite = REFDATARELFILE.hy_composite ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.reform_id = REFDATARELFILE.reform_id ";
                tmp_sql = tmp_sql + " 			WHERE RTRIM(LTRIM(ISNULL(rel_file2,''))) <> '' ";
                tmp_sql = tmp_sql + " 		) AS RELFILE2 ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 REFDATABASE.bk_no ";
                tmp_sql = tmp_sql + " 				,[修繕No] ";
                tmp_sql = tmp_sql + " 				,reform_kbn ";
                tmp_sql = tmp_sql + " 				,3 AS [sort_no] ";
                tmp_sql = tmp_sql + " 				,rel_file3 ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_composite ";
                tmp_sql = tmp_sql + " 					,reform_id ";
                tmp_sql = tmp_sql + " 					,reform_kbn ";
                tmp_sql = tmp_sql + " 					,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No] ";
                tmp_sql = tmp_sql + " 				FROM reform_data ";
                tmp_sql = tmp_sql + " 			) AS REFDATABASE ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_composite ";
                tmp_sql = tmp_sql + " 						,reform_id ";
                tmp_sql = tmp_sql + " 						,rel_file1,rel_file2,rel_file3,rel_file4,rel_file5 ";
                tmp_sql = tmp_sql + " 					FROM reform_data		 ";
                tmp_sql = tmp_sql + " 				) AS REFDATARELFILE ";
                tmp_sql = tmp_sql + " 			ON  REFDATABASE.bk_no = REFDATARELFILE.bk_no ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.hy_composite = REFDATARELFILE.hy_composite ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.reform_id = REFDATARELFILE.reform_id ";
                tmp_sql = tmp_sql + " 			WHERE RTRIM(LTRIM(ISNULL(rel_file3,''))) <> '' ";
                tmp_sql = tmp_sql + " 		) AS RELFILE3 ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 REFDATABASE.bk_no ";
                tmp_sql = tmp_sql + " 				,[修繕No] ";
                tmp_sql = tmp_sql + " 				,reform_kbn ";
                tmp_sql = tmp_sql + " 				,4 AS [sort_no] ";
                tmp_sql = tmp_sql + " 				,rel_file4 ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_composite ";
                tmp_sql = tmp_sql + " 					,reform_id ";
                tmp_sql = tmp_sql + " 					,reform_kbn ";
                tmp_sql = tmp_sql + " 					,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No] ";
                tmp_sql = tmp_sql + " 				FROM reform_data ";
                tmp_sql = tmp_sql + " 			) AS REFDATABASE ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_composite ";
                tmp_sql = tmp_sql + " 						,reform_id ";
                tmp_sql = tmp_sql + " 						,rel_file1,rel_file2,rel_file3,rel_file4,rel_file5 ";
                tmp_sql = tmp_sql + " 					FROM reform_data		 ";
                tmp_sql = tmp_sql + " 				) AS REFDATARELFILE ";
                tmp_sql = tmp_sql + " 			ON  REFDATABASE.bk_no = REFDATARELFILE.bk_no ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.hy_composite = REFDATARELFILE.hy_composite ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.reform_id = REFDATARELFILE.reform_id ";
                tmp_sql = tmp_sql + " 			WHERE RTRIM(LTRIM(ISNULL(rel_file4,''))) <> '' ";
                tmp_sql = tmp_sql + " 		) AS RELFILE4 ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 REFDATABASE.bk_no ";
                tmp_sql = tmp_sql + " 				,[修繕No] ";
                tmp_sql = tmp_sql + " 				,reform_kbn ";
                tmp_sql = tmp_sql + " 				,5 AS [sort_no] ";
                tmp_sql = tmp_sql + " 				,rel_file5 ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_composite ";
                tmp_sql = tmp_sql + " 					,reform_id ";
                tmp_sql = tmp_sql + " 					,reform_kbn ";
                tmp_sql = tmp_sql + " 					,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No] ";
                tmp_sql = tmp_sql + " 				FROM reform_data ";
                tmp_sql = tmp_sql + " 			) AS REFDATABASE ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_composite ";
                tmp_sql = tmp_sql + " 						,reform_id ";
                tmp_sql = tmp_sql + " 						,rel_file1,rel_file2,rel_file3,rel_file4,rel_file5 ";
                tmp_sql = tmp_sql + " 					FROM reform_data		 ";
                tmp_sql = tmp_sql + " 				) AS REFDATARELFILE ";
                tmp_sql = tmp_sql + " 			ON  REFDATABASE.bk_no = REFDATARELFILE.bk_no ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.hy_composite = REFDATARELFILE.hy_composite ";
                tmp_sql = tmp_sql + " 			AND REFDATABASE.reform_id = REFDATARELFILE.reform_id ";
                tmp_sql = tmp_sql + " 			WHERE RTRIM(LTRIM(ISNULL(rel_file5,''))) <> '' ";
                tmp_sql = tmp_sql + " 		) AS RELFILE5 ";
                tmp_sql = tmp_sql + " 	) AS TOTAL ";
                tmp_sql = tmp_sql + " 	WHERE reform_kbn = 1 ";
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

    #region 修繕メモ情報

    public class Reform_jyuyo_Repository
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

                sortstr = "[物件No],[修繕No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[修繕No],[備考1],[備考2],[備考3],[備考4],[備考5],[備考6],[備考7],[備考8],[備考9],[備考10],[備考11],[備考12],[備考13],[備考14],[備考15],[備考16],[備考17],[備考18],[備考19],[備考20],[備考21],[備考22],[備考23],[備考24],[備考25],[備考26],[備考27],[備考28],[備考29],[備考30],[備考31],[備考32],[備考33],[備考34],[備考35],[備考36],[備考37],[備考38],[備考39],[備考40],[備考41],[備考42],[備考43],[備考44],[備考45],[備考46],[備考47],[備考48],[備考49],[備考50]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 REFJUYO.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,[修繕No] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 1 THEN jyuyo_lstname ELSE '' END) AS 備考1 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 2 THEN jyuyo_lstname ELSE '' END) AS 備考2 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 3 THEN jyuyo_lstname ELSE '' END) AS 備考3 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 4 THEN jyuyo_lstname ELSE '' END) AS 備考4 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 5 THEN jyuyo_lstname ELSE '' END) AS 備考5 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 6 THEN jyuyo_lstname ELSE '' END) AS 備考6 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 7 THEN jyuyo_lstname ELSE '' END) AS 備考7 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 8 THEN jyuyo_lstname ELSE '' END) AS 備考8 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 9 THEN jyuyo_lstname ELSE '' END) AS 備考9 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 10 THEN jyuyo_lstname ELSE '' END) AS 備考10 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 11 THEN jyuyo_lstname ELSE '' END) AS 備考11 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 12 THEN jyuyo_lstname ELSE '' END) AS 備考12 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 13 THEN jyuyo_lstname ELSE '' END) AS 備考13 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 14 THEN jyuyo_lstname ELSE '' END) AS 備考14 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 15 THEN jyuyo_lstname ELSE '' END) AS 備考15 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 16 THEN jyuyo_lstname ELSE '' END) AS 備考16 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 17 THEN jyuyo_lstname ELSE '' END) AS 備考17 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 18 THEN jyuyo_lstname ELSE '' END) AS 備考18 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 19 THEN jyuyo_lstname ELSE '' END) AS 備考19 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 20 THEN jyuyo_lstname ELSE '' END) AS 備考20 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 21 THEN jyuyo_lstname ELSE '' END) AS 備考21 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 22 THEN jyuyo_lstname ELSE '' END) AS 備考22 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 23 THEN jyuyo_lstname ELSE '' END) AS 備考23 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 24 THEN jyuyo_lstname ELSE '' END) AS 備考24 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 25 THEN jyuyo_lstname ELSE '' END) AS 備考25 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 26 THEN jyuyo_lstname ELSE '' END) AS 備考26 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 27 THEN jyuyo_lstname ELSE '' END) AS 備考27 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 28 THEN jyuyo_lstname ELSE '' END) AS 備考28 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 29 THEN jyuyo_lstname ELSE '' END) AS 備考29 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 30 THEN jyuyo_lstname ELSE '' END) AS 備考30 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 31 THEN jyuyo_lstname ELSE '' END) AS 備考31 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 32 THEN jyuyo_lstname ELSE '' END) AS 備考32 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 33 THEN jyuyo_lstname ELSE '' END) AS 備考33 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 34 THEN jyuyo_lstname ELSE '' END) AS 備考34 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 35 THEN jyuyo_lstname ELSE '' END) AS 備考35 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 36 THEN jyuyo_lstname ELSE '' END) AS 備考36 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 37 THEN jyuyo_lstname ELSE '' END) AS 備考37 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 38 THEN jyuyo_lstname ELSE '' END) AS 備考38 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 39 THEN jyuyo_lstname ELSE '' END) AS 備考39 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 40 THEN jyuyo_lstname ELSE '' END) AS 備考40 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 41 THEN jyuyo_lstname ELSE '' END) AS 備考41 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 42 THEN jyuyo_lstname ELSE '' END) AS 備考42 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 43 THEN jyuyo_lstname ELSE '' END) AS 備考43 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 44 THEN jyuyo_lstname ELSE '' END) AS 備考44 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 45 THEN jyuyo_lstname ELSE '' END) AS 備考45 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 46 THEN jyuyo_lstname ELSE '' END) AS 備考46 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 47 THEN jyuyo_lstname ELSE '' END) AS 備考47 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 48 THEN jyuyo_lstname ELSE '' END) AS 備考48 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 49 THEN jyuyo_lstname ELSE '' END) AS 備考49 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN jyuyo_no = 50 THEN jyuyo_lstname ELSE '' END) AS 備考50 ";
                tmp_sql = tmp_sql + " 	FROM reform_jyuyo AS REFJUYO ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_composite,reform_id) AS [修繕No],* FROM reform_data ";
                tmp_sql = tmp_sql + " 		) AS REFDATA ";
                tmp_sql = tmp_sql + " 	ON  REFJUYO.bk_no = REFDATA.bk_no ";
                tmp_sql = tmp_sql + " 	AND REFJUYO.hy_composite = REFDATA.hy_composite ";
                tmp_sql = tmp_sql + " 	AND REFJUYO.reform_id = REFDATA.reform_id ";
                tmp_sql = tmp_sql + " 	GROUP BY REFJUYO.bk_no,[修繕No] ";
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