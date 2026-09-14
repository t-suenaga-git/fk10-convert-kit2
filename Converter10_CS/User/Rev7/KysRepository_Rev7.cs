using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using Converter10.Njc.Common;

namespace Converter10.Njc.Repository
{

    #region 契約者情報

    public class Kys_mst_Repository
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

                sortstr = "[契約者No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[契約者No],[契約者名],[契約者名SJIS],[契約者カナ],[個人法人フラグ],[宛名敬称],[郵便番号],[住所１],[住所２],[TEL１],[TEL２],[FAX],[携帯１],[携帯２],[メールアドレス],[携帯メールアドレス],[優先設定(電話番号)],[優先設定(メールアドレス)],[生年月日・設立日],[年収・年商],[備考(基本情報)],[性別],[本籍地],[勤務先名],[勤務先名SJIS],[勤務先カナ],[勤務先郵便番号],[勤務先住所１],[勤務先住所２],[勤務先TEL１],[勤務先TEL２],[勤務先FAX],[勤務先業種],[勤務先部署],[勤務先情報記入年月],[勤務先入社年月],[入居前連絡先情報記入年月],[入居前連絡先郵便番号],[入居前連絡先住所１],[入居前連絡先住所２],[入居前連絡先TEL１],[入居前連絡先TEL２],[入居前連絡先FAX],[Webアドレス(URL)],[業種],[代表者名],[代表者名SJIS],[代表者カナ],[代表者役職],[代表者を宛先に含める],[担当者名],[担当者名SJIS],[担当者カナ],[担当者部署],[担当者役職],[担当者を宛先に含める],[資本金],[従業員数],[記入年月],[主要取引先],[連絡先名],[連絡先名SJIS],[連絡先カナ],[連絡先宛名敬称],[連絡先郵便番号],[連絡先住所１],[連絡先住所２],[連絡先TEL１],[連絡先TEL２],[連絡先FAX],[連絡先携帯１],[連絡先携帯２],[連絡先優先設定(電話番号)],[連絡先間柄],[連絡先備考],[書類送付先区分],[書類送付先名],[書類送付先名SJIS],[書類送付先カナ],[書類送付先宛名敬称],[書類送付先郵便番号],[書類送付先住所１],[書類送付先住所２],[書類送付先TEL１],[書類送付先TEL２],[書類送付先FAX],[書類送付先備考],[振込通知書発行の可否],[振替通知書発行の可否],[督促状発行の可否],[口座振込仮想口座使用フラグ],[保証人複数フラグ],[口座振替の保証有無],[コンビニ収納サービス利用有無]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,kys_name AS [契約者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約者名SJIS] ";
                tmp_sql = tmp_sql + " 		,kys_kana AS [契約者カナ] ";
                tmp_sql = tmp_sql + " 		,kys_rui AS [個人法人フラグ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kys_rui = 1 THEN '様' ";
                tmp_sql = tmp_sql + " 			WHEN kys_rui = 2 THEN '御中' ";
                tmp_sql = tmp_sql + " 			ELSE '様' ";
                tmp_sql = tmp_sql + " 		 END AS [宛名敬称] ";
                tmp_sql = tmp_sql + " 		,post AS [郵便番号] ";
                tmp_sql = tmp_sql + " 		,add1 AS [住所１] ";
                tmp_sql = tmp_sql + " 		,add2 AS [住所２] ";
                tmp_sql = tmp_sql + " 		,tel1 AS [TEL１] ";
                tmp_sql = tmp_sql + " 		,tel2 AS [TEL２] ";
                tmp_sql = tmp_sql + " 		,fax AS [FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯１] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯２] ";
                tmp_sql = tmp_sql + " 		,email AS [メールアドレス] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯メールアドレス] ";
                tmp_sql = tmp_sql + " 		,1 AS [優先設定(電話番号)] ";
                tmp_sql = tmp_sql + " 		,1 AS [優先設定(メールアドレス)] ";
                tmp_sql = tmp_sql + " 		,birthday AS [生年月日・設立日] ";
                tmp_sql = tmp_sql + " 		,nensyu AS [年収・年商] ";
                tmp_sql = tmp_sql + " 		,biko AS [備考(基本情報)] ";
                tmp_sql = tmp_sql + " 		,sex AS [性別] ";
                tmp_sql = tmp_sql + " 		,honseki AS [本籍地] ";
                tmp_sql = tmp_sql + " 		,kinmu_name AS [勤務先名] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先名SJIS] ";
                tmp_sql = tmp_sql + " 		,kinmu_kana AS [勤務先カナ] ";
                tmp_sql = tmp_sql + " 		,kinmu_post AS [勤務先郵便番号] ";
                tmp_sql = tmp_sql + " 		,kinmu_add1 AS [勤務先住所１] ";
                tmp_sql = tmp_sql + " 		,kinmu_add2 AS [勤務先住所２] ";
                tmp_sql = tmp_sql + " 		,kinmu_tel1 AS [勤務先TEL１] ";
                tmp_sql = tmp_sql + " 		,kinmu_tel2 AS [勤務先TEL２] ";
                tmp_sql = tmp_sql + " 		,kinmu_fax AS [勤務先FAX] ";
                tmp_sql = tmp_sql + " 		,kinmu_gy AS [勤務先業種] ";
                tmp_sql = tmp_sql + " 		,kinmu_busyo AS [勤務先部署] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先情報記入年月] ";
                tmp_sql = tmp_sql + " 		,kinmu_nsya AS [勤務先入社年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先情報記入年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先郵便番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先住所１] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先住所２] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先TEL１] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先TEL２] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先FAX] ";
                tmp_sql = tmp_sql + " 		,http AS [Webアドレス(URL)] ";
                tmp_sql = tmp_sql + " 		,gyosyu AS [業種] ";
                tmp_sql = tmp_sql + " 		,daihyo_nama AS [代表者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者名SJIS] ";
                tmp_sql = tmp_sql + " 		,daihyo_kana AS [代表者カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者役職] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者を宛先に含める] ";
                tmp_sql = tmp_sql + " 		,tanto_name AS [担当者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者名SJIS] ";
                tmp_sql = tmp_sql + " 		,tanto_kana AS [担当者カナ] ";
                tmp_sql = tmp_sql + " 		,busyo_name AS [担当者部署] ";
                tmp_sql = tmp_sql + " 		,yakusyoku AS [担当者役職] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者を宛先に含める] ";
                tmp_sql = tmp_sql + " 		,sihonkin * 100 AS [資本金] ";
                tmp_sql = tmp_sql + " 		,jyugyo_suu AS [従業員数] ";
                tmp_sql = tmp_sql + " 		,'' AS [記入年月] ";
                tmp_sql = tmp_sql + " 		,torihikisaki AS [主要取引先] ";
                tmp_sql = tmp_sql + " 		,ren_name AS [連絡先名] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先名SJIS] ";
                tmp_sql = tmp_sql + " 		,ren_kana AS [連絡先カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先宛名敬称] ";
                tmp_sql = tmp_sql + " 		,ren_post AS [連絡先郵便番号] ";
                tmp_sql = tmp_sql + " 		,ren_add1 AS [連絡先住所１] ";
                tmp_sql = tmp_sql + " 		,ren_add2 AS [連絡先住所２] ";
                tmp_sql = tmp_sql + " 		,ren_tel1 AS [連絡先TEL１] ";
                tmp_sql = tmp_sql + " 		,ren_tel2 AS [連絡先TEL２] ";
                tmp_sql = tmp_sql + " 		,ren_fax AS [連絡先FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先携帯１] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先携帯２] ";
                tmp_sql = tmp_sql + " 		,1 AS [連絡先優先設定(電話番号)] ";
                tmp_sql = tmp_sql + " 		,ren_aida AS [連絡先間柄] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先備考] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN syorui_kbn = 1 THEN 6 ";
                tmp_sql = tmp_sql + " 			WHEN syorui_kbn = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN syorui_kbn = 3 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN syorui_kbn = 4 THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN syorui_kbn = 5 THEN 5 ";
                tmp_sql = tmp_sql + " 			WHEN syorui_kbn = 6 THEN 4 ";
                tmp_sql = tmp_sql + " 		 END AS [書類送付先区分] ";
                tmp_sql = tmp_sql + " 		,syorui_name AS [書類送付先名] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先名SJIS] ";
                tmp_sql = tmp_sql + " 		,syorui_kana AS [書類送付先カナ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN keisyo_kbn = 1 THEN '様' ";
                tmp_sql = tmp_sql + " 			WHEN keisyo_kbn = 2 THEN '御中' ";
                tmp_sql = tmp_sql + " 		 END AS [書類送付先宛名敬称] ";
                tmp_sql = tmp_sql + " 		,syorui_post AS [書類送付先郵便番号] ";
                tmp_sql = tmp_sql + " 		,syorui_add1 AS [書類送付先住所１] ";
                tmp_sql = tmp_sql + " 		,syorui_add2 AS [書類送付先住所２] ";
                tmp_sql = tmp_sql + " 		,syorui_tel1 AS [書類送付先TEL１] ";
                tmp_sql = tmp_sql + " 		,syorui_tel2 AS [書類送付先TEL２] ";
                tmp_sql = tmp_sql + " 		,syorui_fax AS [書類送付先FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先備考] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN fkomfkae_tutiumu = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN fkomfkae_tutiumu = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 		 END AS [振込通知書発行の可否] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN fkomfkae_tutiumu = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN fkomfkae_tutiumu = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 		 END AS [振替通知書発行の可否] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN tokusoku_umu = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN tokusoku_umu = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 		 END AS [督促状発行の可否] ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約者基本情報 del sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,0 AS [物件単位振込通知フラグ] ";
                tmp_sql = tmp_sql + " 		,0 AS [物件単位振替通知フラグ] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約者基本情報 del end*/ ";
                tmp_sql = tmp_sql + " 		,0 AS [口座振込仮想口座使用フラグ] ";
                tmp_sql = tmp_sql + " 		,1 AS [保証人複数フラグ] ";
                tmp_sql = tmp_sql + " 		,2 AS [口座振替の保証有無] ";
                tmp_sql = tmp_sql + " 		,1 AS [コンビニ収納サービス利用有無] ";
                tmp_sql = tmp_sql + " 	FROM kys_mst ";
                tmp_sql = tmp_sql + " ) AS KYSMST ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 			)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 		,kari_name AS [契約者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約者名SJIS] ";
                tmp_sql = tmp_sql + " 		,kari_kana AS [契約者カナ] ";
                tmp_sql = tmp_sql + " 		,kari_rui AS [個人法人フラグ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_rui = 1 THEN '様' ";
                tmp_sql = tmp_sql + " 			WHEN kari_rui = 2 THEN '御中' ";
                tmp_sql = tmp_sql + " 			ELSE '様' ";
                tmp_sql = tmp_sql + " 		 END AS [宛名敬称] ";
                tmp_sql = tmp_sql + " 		,post AS [郵便番号] ";
                tmp_sql = tmp_sql + " 		,add1 AS [住所１] ";
                tmp_sql = tmp_sql + " 		,add2 AS [住所２] ";
                tmp_sql = tmp_sql + " 		,tel1 AS [TEL１] ";
                tmp_sql = tmp_sql + " 		,tel2 AS [TEL２] ";
                tmp_sql = tmp_sql + " 		,fax AS [FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯１] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯２] ";
                tmp_sql = tmp_sql + " 		,email AS [メールアドレス] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯メールアドレス] ";
                tmp_sql = tmp_sql + " 		,1 AS [優先設定(電話番号)] ";
                tmp_sql = tmp_sql + " 		,1 AS [優先設定(メールアドレス)] ";
                tmp_sql = tmp_sql + " 		,birthday AS [生年月日・設立日] ";
                tmp_sql = tmp_sql + " 		,nensyu AS [年収・年商] ";
                tmp_sql = tmp_sql + " 		,kari_biko AS [備考(基本情報)] ";
                tmp_sql = tmp_sql + " 		,sex AS [性別] ";
                tmp_sql = tmp_sql + " 		,honseki AS [本籍地] ";
                tmp_sql = tmp_sql + " 		,kinmu_name AS [勤務先名] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先名SJIS] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先カナ] ";
                tmp_sql = tmp_sql + " 		,kinmu_post AS [勤務先郵便番号] ";
                tmp_sql = tmp_sql + " 		,kinmu_add1 AS [勤務先住所１] ";
                tmp_sql = tmp_sql + " 		,kinmu_add2 AS [勤務先住所２] ";
                tmp_sql = tmp_sql + " 		,kinmu_tel1 AS [勤務先TEL１] ";
                tmp_sql = tmp_sql + " 		,kinmu_tel2 AS [勤務先TEL２] ";
                tmp_sql = tmp_sql + " 		,kinmu_fax AS [勤務先FAX] ";
                tmp_sql = tmp_sql + " 		,kinmu_gy AS [勤務先業種] ";
                tmp_sql = tmp_sql + " 		,kinmu_busyo AS [勤務先部署] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先情報記入年月] ";
                tmp_sql = tmp_sql + " 		,kinmu_nsya AS [勤務先入社年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先情報記入年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先郵便番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先住所１] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先住所２] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先TEL１] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先TEL２] ";
                tmp_sql = tmp_sql + " 		,'' AS [入居前連絡先FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [Webアドレス(URL)] ";
                tmp_sql = tmp_sql + " 		,gyosyu AS [業種] ";
                tmp_sql = tmp_sql + " 		,daihyo_name AS [代表者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者名SJIS] ";
                tmp_sql = tmp_sql + " 		,daihyo_kana AS [代表者カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者役職] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者を宛先に含める] ";
                tmp_sql = tmp_sql + " 		,tanto_name AS [担当者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者名SJIS] ";
                tmp_sql = tmp_sql + " 		,tanto_kana AS [担当者カナ] ";
                tmp_sql = tmp_sql + " 		,busyo_name AS [担当者部署] ";
                tmp_sql = tmp_sql + " 		,yakusyoku AS [担当者役職] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者宛名追加フラグ] ";
                tmp_sql = tmp_sql + " 		,sihonkin * 100 AS [資本金] ";
                tmp_sql = tmp_sql + " 		,jyugyo_suu AS [従業員数] ";
                tmp_sql = tmp_sql + " 		,'' AS [記入年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [主要取引先] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先名] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先名SJIS] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先宛名敬称] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先郵便番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先住所１] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先住所２] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先TEL１] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先TEL２] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先携帯１] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先携帯２] ";
                tmp_sql = tmp_sql + " 		,1 AS [連絡先優先設定(電話番号)] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先間柄] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先備考] ";
                tmp_sql = tmp_sql + " 		,2 AS [書類送付先区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先名] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先名SJIS] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先カナ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			 WHEN kari_rui = 1 THEN '様' ";
                tmp_sql = tmp_sql + " 			 WHEN kari_rui = 2 THEN '御中' ";
                tmp_sql = tmp_sql + " 		 END AS [書類送付先宛名敬称] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先郵便番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先住所１] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先住所２] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先TEL１] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先TEL２] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類送付先備考] ";
                tmp_sql = tmp_sql + " 		,2 AS [振込通知書発行の可否] ";
                tmp_sql = tmp_sql + " 		,2 AS [振替通知書発行の可否] ";
                tmp_sql = tmp_sql + " 		,2 AS [督促状発行の可否] ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約者基本情報 del sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,0 AS [物件単位振込通知フラグ] ";
                tmp_sql = tmp_sql + " 		,0 AS [物件単位振替通知フラグ] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約者基本情報 del end*/		 ";
                tmp_sql = tmp_sql + " 		,0 AS [口座振込仮想口座使用フラグ] ";
                tmp_sql = tmp_sql + " 		,1 AS [保証人複数フラグ] ";
                tmp_sql = tmp_sql + " 		,2 AS [口座振替の保証有無] ";
                tmp_sql = tmp_sql + " 		,1 AS [コンビニ収納サービス利用有無] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 	) AS VW ";
                tmp_sql = tmp_sql + " ) AS KARIKYS ";

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

    #region 契約者口座情報

    public class Kys_mst_koza_Repository
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

                sortstr = "[契約者No],[契約者口座No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[契約者No],[契約者口座No],[金融期間No],[金融機関支店No],[口座種別],[口座番号],[口座名義],[口座名義カナ],[ゆうちょ記号１],[ゆうちょ記号２],[ゆうちょ口座番号],[口座振替No],[口座振替手数料],[口座振替契約者番号],[口座備考],[口座振込備考],[総合振込区分],[振込依頼人No],[振込手数料負担区分],[振込手数料計算区分],[振込手数料固定額1],[振込手数料固定額2],[備考(総合振込情報)],[印刷口座区分],[金融機関区分(ゆうちょ銀行フラグ)]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,1 AS [契約者口座No] ";
                tmp_sql = tmp_sql + " 		,kinyu_no AS [金融期間No] ";
                tmp_sql = tmp_sql + " 		,ten_no AS [金融機関支店No] ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kosyu_no = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE kosyu_no ";
                tmp_sql = tmp_sql + " 		 END AS [口座種別] ";
                tmp_sql = tmp_sql + " 		 */ ";
                tmp_sql = tmp_sql + " 		,kosyu_name AS [口座種別] ";
                tmp_sql = tmp_sql + " 		,koza_no AS [口座番号] ";
                tmp_sql = tmp_sql + " 		,koza_meigi AS [口座名義] ";
                tmp_sql = tmp_sql + " 		,koza_kana AS [口座名義カナ] ";
                tmp_sql = tmp_sql + " 		,post_kigo AS [ゆうちょ記号１] ";
                tmp_sql = tmp_sql + " 		,post_n AS [ゆうちょ記号２] ";
                tmp_sql = tmp_sql + " 		,post_bango AS [ゆうちょ口座番号] ";
                tmp_sql = tmp_sql + " 		,fkae_no AS [口座振替No] ";
                tmp_sql = tmp_sql + " 		,fkae_tesu AS [口座振替手数料] ";
                tmp_sql = tmp_sql + " 		,kys_bango AS [口座振替契約者番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振込備考] ";
                tmp_sql = tmp_sql + " 		,1 AS [総合振込区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込依頼人No] ";
                tmp_sql = tmp_sql + " 		,2 AS [振込手数料負担区分] ";
                tmp_sql = tmp_sql + " 		,1 AS [振込手数料計算区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込手数料固定額1] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込手数料固定額2] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考(総合振込情報)] ";
                tmp_sql = tmp_sql + " 		,bankpost_kbn AS [印刷口座区分]		/*20160519 EXEUpdateに伴う修正 契約者口座情報 add*/ ";
                tmp_sql = tmp_sql + " 		/*20161012 革命10アップデートに伴う修正 add sta*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN bankpost_kbn = 1 THEN 0 ";
                tmp_sql = tmp_sql + " 			WHEN bankpost_kbn = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 0 ";
                tmp_sql = tmp_sql + " 		 END AS [金融機関区分(ゆうちょ銀行フラグ)] ";
                tmp_sql = tmp_sql + " 		/*20161012 革命10アップデートに伴う修正 add end*/ ";
                tmp_sql = tmp_sql + " 	FROM kys_mst AS KYS ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_kozasyu AS MK ON KYS.kosyu_no = MK.kosyu_no ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,2 AS [契約者口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融期間No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融機関支店No] ";
                tmp_sql = tmp_sql + " 		,'普通預金' AS [口座種別] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号１] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号２] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替No] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替手数料] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替契約者番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振込備考] ";
                tmp_sql = tmp_sql + " 		,1 AS [総合振込区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込依頼人No] ";
                tmp_sql = tmp_sql + " 		,2 AS [振込手数料負担区分] ";
                tmp_sql = tmp_sql + " 		,1 AS [振込手数料計算区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込手数料固定額1] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込手数料固定額2] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考(総合振込情報)] ";
                tmp_sql = tmp_sql + " 		,1 AS [印刷口座区分] ";
                tmp_sql = tmp_sql + " 		,0 AS [金融機関区分(ゆうちょ銀行フラグ)]	/*20161012 革命10アップデートに伴う修正 add*/ ";
                tmp_sql = tmp_sql + " 	FROM kys_mst ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,3 AS [契約者口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融期間No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融機関支店No] ";
                tmp_sql = tmp_sql + " 		,'普通預金' AS [口座種別] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号１] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号２] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替No] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替手数料] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替契約者番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振込備考] ";
                tmp_sql = tmp_sql + " 		,NULL AS [総合振込区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込依頼人No] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込手数料負担区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込手数料計算区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込手数料固定額1] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込手数料固定額2] ";
                tmp_sql = tmp_sql + " 		,NULL AS [備考(総合振込情報)] ";
                tmp_sql = tmp_sql + " 		,1 AS [印刷口座区分] ";
                tmp_sql = tmp_sql + " 		,0 AS [金融機関区分(ゆうちょ銀行フラグ)]	/*20161012 革命10アップデートに伴う修正 add*/ ";
                tmp_sql = tmp_sql + " 	FROM kys_mst ";
                tmp_sql = tmp_sql + " ) AS KYSKOZA ";
                tmp_sql = tmp_sql + " /*2016.04.06 仮契約者のダミー口座(3レコード作成) add sta*/ ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 			)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 		,1 AS [契約者口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融期間No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融機関支店No] ";
                tmp_sql = tmp_sql + " 		,'普通預金' AS [口座種別] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号１] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号２] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替No] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替手数料] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替契約者番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振込備考] ";
                tmp_sql = tmp_sql + " 		,1 AS [総合振込区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込依頼人No] ";
                tmp_sql = tmp_sql + " 		,2 AS [振込手数料負担区分] ";
                tmp_sql = tmp_sql + " 		,1 AS [振込手数料計算区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込手数料固定額1] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込手数料固定額2] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考(総合振込情報)] ";
                tmp_sql = tmp_sql + " 		,1 AS [印刷口座区分] ";
                tmp_sql = tmp_sql + " 		,0 AS [金融機関区分(ゆうちょ銀行フラグ)]	/*20161012 革命10アップデートに伴う修正 add*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 	) AS VW1 ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 			)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 		,2 AS [契約者口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融期間No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融機関支店No] ";
                tmp_sql = tmp_sql + " 		,'普通預金' AS [口座種別] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号１] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号２] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替No] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替手数料] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替契約者番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振込備考] ";
                tmp_sql = tmp_sql + " 		,1 AS [総合振込区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込依頼人No] ";
                tmp_sql = tmp_sql + " 		,2 AS [振込手数料負担区分] ";
                tmp_sql = tmp_sql + " 		,1 AS [振込手数料計算区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込手数料固定額1] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込手数料固定額2] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考(総合振込情報)] ";
                tmp_sql = tmp_sql + " 		,1 AS [印刷口座区分] ";
                tmp_sql = tmp_sql + " 		,0 AS [金融機関区分(ゆうちょ銀行フラグ)]	/*20161012 革命10アップデートに伴う修正 add*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 	) AS VW2 ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 			)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 		,3 AS [契約者口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融期間No] ";
                tmp_sql = tmp_sql + " 		,'' AS [金融機関支店No] ";
                tmp_sql = tmp_sql + " 		,'普通預金' AS [口座種別] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座名義カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号１] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ記号２] ";
                tmp_sql = tmp_sql + " 		,'' AS [ゆうちょ口座番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替No] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替手数料] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替契約者番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振込備考] ";
                tmp_sql = tmp_sql + " 		,NULL AS [総合振込区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込依頼人No] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込手数料負担区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込手数料計算区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込手数料固定額1] ";
                tmp_sql = tmp_sql + " 		,NULL AS [振込手数料固定額2] ";
                tmp_sql = tmp_sql + " 		,NULL AS [備考(総合振込情報)] ";
                tmp_sql = tmp_sql + " 		,1 AS [印刷口座区分] ";
                tmp_sql = tmp_sql + " 		,0 AS [金融機関区分(ゆうちょ銀行フラグ)]	/*20161012 革命10アップデートに伴う修正 add*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 		WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 	) AS VW3	 ";
                tmp_sql = tmp_sql + " ) AS KARIKYSKOZA ";
                tmp_sql = tmp_sql + " /*2016.04.06 仮契約者のダミー口座(3レコード作成) add end*/ ";

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

    #region 契約者照合用カナ情報

    public class Kys_mst_syogokana_Repository
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

                sortstr = "[契約者No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[契約者No],[照合用カナ1],[照合用カナ2],[照合用カナ3],[照合用カナ4],[照合用カナ5],[照合用カナ6],[照合用カナ7],[照合用カナ8],[照合用カナ9],[照合用カナ10]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 [kys_no] AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,[syogo_kana1] AS [照合用カナ1] ";
                tmp_sql = tmp_sql + " 		,[syogo_kana2] AS [照合用カナ2] ";
                tmp_sql = tmp_sql + " 		,[syogo_kana3] AS [照合用カナ3] ";
                tmp_sql = tmp_sql + " 		,'' AS [照合用カナ4] ";
                tmp_sql = tmp_sql + " 		,'' AS [照合用カナ5] ";
                tmp_sql = tmp_sql + " 		,'' AS [照合用カナ6] ";
                tmp_sql = tmp_sql + " 		,'' AS [照合用カナ7] ";
                tmp_sql = tmp_sql + " 		,'' AS [照合用カナ8] ";
                tmp_sql = tmp_sql + " 		,'' AS [照合用カナ9] ";
                tmp_sql = tmp_sql + " 		,'' AS [照合用カナ10] ";
                tmp_sql = tmp_sql + " 	FROM kys_mst ";
                tmp_sql = tmp_sql + " 	WHERE ([syogo_kana1] <> '' OR [syogo_kana2] <> '' OR [syogo_kana3] <> '') ";
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

    #region 契約者保証人情報

    public class Kys_mst_hosyonin_Repository
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

                sortstr = "[契約者No],[保証人No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[契約者No],[保証人No],[保証人名],[保証人名SJIS],[保証人カナ],[宛名敬称],[郵便番号],[住所１],[住所２],[TEL１],[TEL２],[FAX],[携帯１],[携帯２],[優先設定(電話番号)],[誕生日・設立日],[年収・年商],[備考],[間柄],[勤務先名],[勤務先名SJIS],[勤務先カナ],[勤務先郵便番号],[勤務先住所１],[勤務先住所２],[勤務先TEL１],[勤務先TEL２],[勤務先FAX],[勤務先業種],[勤務先部署],[勤務先情報記入年月],[勤務先入社年月],[勤務先退社年月],[勤務先備考]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 			,1 AS [保証人No] ";
                tmp_sql = tmp_sql + " 			,hs1_name AS [保証人名] ";
                tmp_sql = tmp_sql + " 			,'' AS [保証人名SJIS] ";
                tmp_sql = tmp_sql + " 			,hs1_kana AS [保証人カナ] ";
                tmp_sql = tmp_sql + " 			,'' AS [宛名敬称] ";
                tmp_sql = tmp_sql + " 			,hs1_post AS [郵便番号] ";
                tmp_sql = tmp_sql + " 			,hs1_add1 AS [住所１] ";
                tmp_sql = tmp_sql + " 			,hs1_add2 AS [住所２] ";
                tmp_sql = tmp_sql + " 			,hs1_tel1 AS [TEL１] ";
                tmp_sql = tmp_sql + " 			,hs1_tel2 AS [TEL２] ";
                tmp_sql = tmp_sql + " 			,hs1_fax AS [FAX] ";
                tmp_sql = tmp_sql + " 			,'' AS [携帯１] ";
                tmp_sql = tmp_sql + " 			,'' AS [携帯２] ";
                tmp_sql = tmp_sql + " 			,1 AS [優先設定(電話番号)] ";
                tmp_sql = tmp_sql + " 			,hs1_birthday AS [誕生日・設立日] ";
                tmp_sql = tmp_sql + " 			,hs1_nensyu AS [年収・年商] ";
                tmp_sql = tmp_sql + " 			,hs1_biko AS [備考] ";
                tmp_sql = tmp_sql + " 			,hs1_aida AS [間柄] ";
                tmp_sql = tmp_sql + " 			,hs1_kinmu_name AS [勤務先名] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先名SJIS] ";
                tmp_sql = tmp_sql + " 			,hs1_kinmu_kana AS [勤務先カナ] ";
                tmp_sql = tmp_sql + " 			,hs1_kinmu_post AS [勤務先郵便番号] ";
                tmp_sql = tmp_sql + " 			,hs1_kinmu_add1 AS [勤務先住所１] ";
                tmp_sql = tmp_sql + " 			,hs1_kinmu_add2 AS [勤務先住所２] ";
                tmp_sql = tmp_sql + " 			,hs1_kinmu_tel1 AS [勤務先TEL１] ";
                tmp_sql = tmp_sql + " 			,hs1_kinmu_tel2 AS [勤務先TEL２] ";
                tmp_sql = tmp_sql + " 			,hs1_kinmu_fax AS [勤務先FAX] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先業種] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先部署] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先情報記入年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先入社年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先退社年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先備考] ";
                tmp_sql = tmp_sql + " 		FROM kys_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 			,2 AS [保証人No] ";
                tmp_sql = tmp_sql + " 			,hs2_name AS [保証人名] ";
                tmp_sql = tmp_sql + " 			,'' AS [保証人名SJIS] ";
                tmp_sql = tmp_sql + " 			,hs2_kana AS [保証人カナ] ";
                tmp_sql = tmp_sql + " 			,'' AS [宛名敬称] ";
                tmp_sql = tmp_sql + " 			,hs2_post AS [郵便番号] ";
                tmp_sql = tmp_sql + " 			,hs2_add1 AS [住所１] ";
                tmp_sql = tmp_sql + " 			,hs2_add2 AS [住所２] ";
                tmp_sql = tmp_sql + " 			,hs2_tel1 AS [TEL１] ";
                tmp_sql = tmp_sql + " 			,hs2_tel2 AS [TEL２] ";
                tmp_sql = tmp_sql + " 			,hs2_fax AS [FAX] ";
                tmp_sql = tmp_sql + " 			,'' AS [携帯１] ";
                tmp_sql = tmp_sql + " 			,'' AS [携帯２] ";
                tmp_sql = tmp_sql + " 			,1 AS [優先設定(電話番号)] ";
                tmp_sql = tmp_sql + " 			,hs2_birthday AS [誕生日・設立日] ";
                tmp_sql = tmp_sql + " 			,hs2_nensyu AS [年収・年商] ";
                tmp_sql = tmp_sql + " 			,hs2_biko AS [備考] ";
                tmp_sql = tmp_sql + " 			,hs2_aida AS [間柄] ";
                tmp_sql = tmp_sql + " 			,hs2_kinmu_name AS [勤務先名] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先名SJIS] ";
                tmp_sql = tmp_sql + " 			,hs2_kinmu_kana AS [勤務先カナ] ";
                tmp_sql = tmp_sql + " 			,hs2_kinmu_post AS [勤務先郵便番号] ";
                tmp_sql = tmp_sql + " 			,hs2_kinmu_add1 AS [勤務先住所１] ";
                tmp_sql = tmp_sql + " 			,hs2_kinmu_add2 AS [勤務先住所２] ";
                tmp_sql = tmp_sql + " 			,hs2_kinmu_tel1 AS [勤務先TEL１] ";
                tmp_sql = tmp_sql + " 			,hs2_kinmu_tel2 AS [勤務先TEL２] ";
                tmp_sql = tmp_sql + " 			,hs2_kinmu_fax AS [勤務先FAX] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先業種] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先部署] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先情報記入年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先入社年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先退社年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先備考] ";
                tmp_sql = tmp_sql + " 		FROM kys_mst ";
                tmp_sql = tmp_sql + " 	) AS KYSHOSYO ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 				)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 			,1 AS [保証人No] ";
                tmp_sql = tmp_sql + " 			,hs1_name AS [保証人名] ";
                tmp_sql = tmp_sql + " 			,'' AS [保証人名SJIS] ";
                tmp_sql = tmp_sql + " 			,hs1_kana AS [保証人カナ] ";
                tmp_sql = tmp_sql + " 			,'' AS [宛名敬称] ";
                tmp_sql = tmp_sql + " 			,hs1_post AS [郵便番号] ";
                tmp_sql = tmp_sql + " 			,hs1_add1 AS [住所１] ";
                tmp_sql = tmp_sql + " 			,hs1_add2 AS [住所２] ";
                tmp_sql = tmp_sql + " 			,hs1_tel1 AS [TEL１] ";
                tmp_sql = tmp_sql + " 			,hs1_tel2 AS [TEL２] ";
                tmp_sql = tmp_sql + " 			,hs1_fax AS [FAX] ";
                tmp_sql = tmp_sql + " 			,'' AS [携帯１] ";
                tmp_sql = tmp_sql + " 			,'' AS [携帯２] ";
                tmp_sql = tmp_sql + " 			,1 AS [優先設定(電話番号)] ";
                tmp_sql = tmp_sql + " 			,hs1_birthday AS [誕生日・設立日] ";
                tmp_sql = tmp_sql + " 			,'' AS [年収・年商] ";
                tmp_sql = tmp_sql + " 			,hs1_biko AS [備考] ";
                tmp_sql = tmp_sql + " 			,hs1_aida AS [間柄] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先名] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先名SJIS] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先カナ] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先郵便番号] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先住所１] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先住所２] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先TEL１] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先TEL２] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先FAX] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先業種] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先部署] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先情報記入年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先入社年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先退社年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先備考] ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 			WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 				)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 			,2 AS [保証人No] ";
                tmp_sql = tmp_sql + " 			,hs2_name AS [保証人名] ";
                tmp_sql = tmp_sql + " 			,'' AS [保証人名SJIS] ";
                tmp_sql = tmp_sql + " 			,hs2_kana AS [保証人カナ] ";
                tmp_sql = tmp_sql + " 			,'' AS [宛名敬称] ";
                tmp_sql = tmp_sql + " 			,hs2_post AS [郵便番号] ";
                tmp_sql = tmp_sql + " 			,hs2_add1 AS [住所１] ";
                tmp_sql = tmp_sql + " 			,hs2_add2 AS [住所２] ";
                tmp_sql = tmp_sql + " 			,hs2_tel1 AS [TEL１] ";
                tmp_sql = tmp_sql + " 			,hs2_tel2 AS [TEL２] ";
                tmp_sql = tmp_sql + " 			,hs2_fax AS [FAX] ";
                tmp_sql = tmp_sql + " 			,'' AS [携帯１] ";
                tmp_sql = tmp_sql + " 			,'' AS [携帯２] ";
                tmp_sql = tmp_sql + " 			,1 AS [優先設定(電話番号)] ";
                tmp_sql = tmp_sql + " 			,hs2_birthday AS [誕生日・設立日] ";
                tmp_sql = tmp_sql + " 			,'' AS [年収・年商] ";
                tmp_sql = tmp_sql + " 			,hs2_biko AS [備考] ";
                tmp_sql = tmp_sql + " 			,hs2_aida AS [間柄] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先名] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先名SJIS] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先カナ] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先郵便番号] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先住所１] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先住所２] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先TEL１] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先TEL２] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先FAX] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先業種] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先部署] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先情報記入年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先入社年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先退社年月] ";
                tmp_sql = tmp_sql + " 			,'' AS [勤務先備考] ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 			WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 	) AS KARIKYSHOSYONIN ";
                tmp_sql = tmp_sql + " ) AS HOSYONINTOTAL ";
                tmp_sql = tmp_sql + " WHERE ISNULL([保証人名],'') <> '' ";

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

    #region 契約者メモ情報

    public class Kys_biko_Repository
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

                sortstr = "[契約者No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[契約者No],[メモ1],[メモ2],[メモ3],[メモ4],[メモ5],[メモ6],[メモ7],[メモ8],[メモ9],[メモ10],[メモ11],[メモ12],[メモ13],[メモ14],[メモ15],[メモ16],[メモ17],[メモ18],[メモ19],[メモ20],[メモ21],[メモ22],[メモ23],[メモ24],[メモ25],[メモ26],[メモ27],[メモ28],[メモ29],[メモ30],[メモ31],[メモ32],[メモ33],[メモ34],[メモ35],[メモ36],[メモ37],[メモ38],[メモ39],[メモ40],[メモ41],[メモ42],[メモ43],[メモ44],[メモ45],[メモ46],[メモ47],[メモ48],[メモ49],[メモ50]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 1 THEN biko ELSE '' END) AS メモ1 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 2 THEN biko ELSE '' END) AS メモ2 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 3 THEN biko ELSE '' END) AS メモ3 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 4 THEN biko ELSE '' END) AS メモ4 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 5 THEN biko ELSE '' END) AS メモ5 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 6 THEN biko ELSE '' END) AS メモ6 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 7 THEN biko ELSE '' END) AS メモ7 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 8 THEN biko ELSE '' END) AS メモ8 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 9 THEN biko ELSE '' END) AS メモ9 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 10 THEN biko ELSE '' END) AS メモ10 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 11 THEN biko ELSE '' END) AS メモ11 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 12 THEN biko ELSE '' END) AS メモ12 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 13 THEN biko ELSE '' END) AS メモ13 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 14 THEN biko ELSE '' END) AS メモ14 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 15 THEN biko ELSE '' END) AS メモ15 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 16 THEN biko ELSE '' END) AS メモ16 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 17 THEN biko ELSE '' END) AS メモ17 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 18 THEN biko ELSE '' END) AS メモ18 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 19 THEN biko ELSE '' END) AS メモ19 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 20 THEN biko ELSE '' END) AS メモ20 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 21 THEN biko ELSE '' END) AS メモ21 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 22 THEN biko ELSE '' END) AS メモ22 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 23 THEN biko ELSE '' END) AS メモ23 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 24 THEN biko ELSE '' END) AS メモ24 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 25 THEN biko ELSE '' END) AS メモ25 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 26 THEN biko ELSE '' END) AS メモ26 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 27 THEN biko ELSE '' END) AS メモ27 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 28 THEN biko ELSE '' END) AS メモ28 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 29 THEN biko ELSE '' END) AS メモ29 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 30 THEN biko ELSE '' END) AS メモ30 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 31 THEN biko ELSE '' END) AS メモ31 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 32 THEN biko ELSE '' END) AS メモ32 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 33 THEN biko ELSE '' END) AS メモ33 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 34 THEN biko ELSE '' END) AS メモ34 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 35 THEN biko ELSE '' END) AS メモ35 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 36 THEN biko ELSE '' END) AS メモ36 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 37 THEN biko ELSE '' END) AS メモ37 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 38 THEN biko ELSE '' END) AS メモ38 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 39 THEN biko ELSE '' END) AS メモ39 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 40 THEN biko ELSE '' END) AS メモ40 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 41 THEN biko ELSE '' END) AS メモ41 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 42 THEN biko ELSE '' END) AS メモ42 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 43 THEN biko ELSE '' END) AS メモ43 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 44 THEN biko ELSE '' END) AS メモ44 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 45 THEN biko ELSE '' END) AS メモ45 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 46 THEN biko ELSE '' END) AS メモ46 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 47 THEN biko ELSE '' END) AS メモ47 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 48 THEN biko ELSE '' END) AS メモ48 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 49 THEN biko ELSE '' END) AS メモ49 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN biko_no = 50 THEN biko ELSE '' END) AS メモ50 ";
                tmp_sql = tmp_sql + " 	FROM kys_biko ";
                tmp_sql = tmp_sql + " 	GROUP BY kys_no ";
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