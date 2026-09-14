using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using Converter10.Njc.Common;

namespace Converter10.Njc.Repository
{

    #region 家主基本情報

    public class M_yanu_so_Repository
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

                sortstr = "[家主No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[家主No],[個人法人フラグ],[家主名],[家主名(SJIS)],[家主カナ],[宛名敬称],[郵便番号],[住所1],[住所2],[登記住所１],[登記住所２],[TEL1],[TEL2],[FAX],[携帯１],[携帯２],[メールアドレス],[携帯メールアドレス],[優先電話設定],[優先メール設定],[誕生日・設立日],[年収・年商],[備考],[性別],[本籍地],[勤務先名],[勤務先名SJIS],[勤務先カナ],[勤務先郵便番号],[勤務先住所１],[勤務先住所２],[勤務先TEL１],[勤務先TEL２],[勤務先FAX],[勤務先業種],[勤務先部署],[勤務先情報記入年月],[勤務先入社年月],[URL],[業種],[代表者名],[代表者名(SJIS)],[代表者カナ],[代表者役職],[代表者を宛先に含める],[担当者名],[担当者名(SJIS)],[担当者カナ],[担当者部署],[担当者を宛先に含める],[資本金],[従業員数],[記入年月],[主要取引先],[連絡先名],[連絡先名SJIS],[連絡先カナ],[連絡先宛名敬称],[連絡先郵便番号],[連絡先住所１],[連絡先住所２],[連絡先TEL１],[連絡先TEL２],[連絡先FAX],[連絡先携帯１],[連絡先携帯２],[連絡先優先設定(電話番号)],[連絡先間柄],[連絡先備考],[書類送付先区分],[送付先名],[送付先名(SJIS)],[送付先カナ],[送付先宛名敬称],[送付先郵便番号],[送付先住所１],[送付先住所２],[送付先TEL１],[送付先TEL2],[送付先FAX],[送付先備考],[イベント担当者No],[絞り込みキーワード],[会計情報]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // 2016.04.25 ユーザーテスト用に抽出クエリを編集 -chg sta
                // 'tmp_sql = tmp_sql & " SELECT * FROM "
                // tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                // tmp_sql = tmp_sql & " ( "
                // tmp_sql = tmp_sql & " 	SELECT "
                // tmp_sql = tmp_sql & " 		 so_no AS [家主No] "
                // tmp_sql = tmp_sql & " 		,kojin_kbn AS [個人法人フラグ] "
                // tmp_sql = tmp_sql & " 		,so_name AS [家主名] "
                // tmp_sql = tmp_sql & " 		,'' AS [家主名(SJIS)] "
                // tmp_sql = tmp_sql & " 		,so_kana AS [家主カナ] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN kojin_kbn = 1 THEN '様' "
                // tmp_sql = tmp_sql & " 			WHEN kojin_kbn = 2 THEN '御中' "
                // tmp_sql = tmp_sql & " 			ELSE '様' "
                // tmp_sql = tmp_sql & " 		 END AS [宛名敬称] "
                // tmp_sql = tmp_sql & " 		,post AS [郵便番号] "
                // tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                // tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                // tmp_sql = tmp_sql & " 		,yobi_mj2 AS [登記住所１] "
                // tmp_sql = tmp_sql & " 		,yobi_mj3 AS [登記住所２] "
                // tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                // tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                // tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                // tmp_sql = tmp_sql & " 		,'' AS [携帯１] "
                // tmp_sql = tmp_sql & " 		,'' AS [携帯２] "
                // tmp_sql = tmp_sql & " 		,email AS [メールアドレス] "
                // tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス] "
                // tmp_sql = tmp_sql & " 		,1 AS [優先電話設定] "
                // tmp_sql = tmp_sql & " 		,1 AS [優先メール設定] "
                // tmp_sql = tmp_sql & " 		,birthday AS [誕生日・設立日] "
                // tmp_sql = tmp_sql & " 		,'' AS [年収・年商] "
                // tmp_sql = tmp_sql & " 		,biko AS [備考] "
                // tmp_sql = tmp_sql & " 		,1 AS [性別] "
                // tmp_sql = tmp_sql & " 		,'' AS [本籍地] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先名] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先名SJIS] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先カナ] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先郵便番号] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先住所１] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先住所２] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先TEL１] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先TEL２] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先FAX] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先業種] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先部署] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先情報記入年月] "
                // tmp_sql = tmp_sql & " 		,'' AS [勤務先入社年月] "
                // tmp_sql = tmp_sql & " 		,'' AS [URL] "
                // tmp_sql = tmp_sql & " 		,'' AS [業種] "
                // tmp_sql = tmp_sql & " 		,'' AS [代表者名] "
                // tmp_sql = tmp_sql & " 		,'' AS [代表者名(SJIS)] "
                // tmp_sql = tmp_sql & " 		,'' AS [代表者カナ] "
                // tmp_sql = tmp_sql & " 		,'' AS [代表者役職] "
                // tmp_sql = tmp_sql & " 		,0 AS [代表者を宛先に含める] "
                // tmp_sql = tmp_sql & " 		,'' AS [担当者名] "
                // tmp_sql = tmp_sql & " 		,'' AS [担当者名(SJIS)] "
                // tmp_sql = tmp_sql & " 		,'' AS [担当者カナ] "
                // tmp_sql = tmp_sql & " 		,'' AS [担当者部署] "
                // tmp_sql = tmp_sql & " 		,0 AS [担当者を宛先に含める] "
                // tmp_sql = tmp_sql & " 		,'' AS [資本金] "
                // tmp_sql = tmp_sql & " 		,'' AS [従業員数] "
                // tmp_sql = tmp_sql & " 		,'' AS [記入年月] "
                // tmp_sql = tmp_sql & " 		,'' AS [主要取引先] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先名] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先名SJIS] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先カナ] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先宛名敬称] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先郵便番号] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先住所１] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先住所２] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先TEL１] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先TEL２] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先FAX] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先携帯１] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先携帯２] "
                // tmp_sql = tmp_sql & " 		,1 AS [連絡先優先設定(電話番号)] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先間柄] "
                // tmp_sql = tmp_sql & " 		,'' AS [連絡先備考] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN ([so_name] + [post] + [add1] + [add2] + [tel1] + [tel2] + [fax]) = ([syorui_name] + [syorui_post] + [syorui_add1] + [syorui_add2] + [syorui_tel1] + [syorui_tel2] + [syorui_fax]) THEN 1 "
                // tmp_sql = tmp_sql & " 			ELSE 4 "
                // tmp_sql = tmp_sql & " 		 END AS [書類送付先区分] "
                // tmp_sql = tmp_sql & " 		,syorui_name AS [送付先名] "
                // tmp_sql = tmp_sql & " 		,'' AS [送付先名(SJIS)] "
                // tmp_sql = tmp_sql & " 		,'' AS [送付先カナ] "
                // tmp_sql = tmp_sql & " 		,'' AS [送付先宛名敬称] "
                // tmp_sql = tmp_sql & " 		,syorui_post AS [送付先郵便番号] "
                // tmp_sql = tmp_sql & " 		,syorui_add1 AS [送付先住所１] "
                // tmp_sql = tmp_sql & " 		,syorui_add2 AS [送付先住所２] "
                // tmp_sql = tmp_sql & " 		,syorui_tel1 AS [送付先TEL１] "
                // tmp_sql = tmp_sql & " 		,syorui_tel2 AS [送付先TEL2] "
                // tmp_sql = tmp_sql & " 		,syorui_fax AS [送付先FAX] "
                // tmp_sql = tmp_sql & " 		,'' AS [送付先備考] "
                // tmp_sql = tmp_sql & " 		,tanto_no AS [イベント担当者No] "
                // tmp_sql = tmp_sql & " 		,'' AS [絞り込みキーワード] "
                // tmp_sql = tmp_sql & " 		,'' AS [会計情報] "
                // tmp_sql = tmp_sql & " 	FROM m_yanu_so "
                // tmp_sql = tmp_sql & " ) AS VW "

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 so_no AS [家主No] ";
                tmp_sql = tmp_sql + " 		,kojin_kbn AS [個人法人フラグ] ";
                tmp_sql = tmp_sql + " 		/*2016.04.25 ユーザーテスト用に編集 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,so_name AS [家主名]*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN RTRIM(LTRIM(ISNULL(so_name,''))) = '' THEN '移行用テスト家主' + CONVERT(NVARCHAR(10),so_no) ";
                tmp_sql = tmp_sql + " 			ELSE so_name ";
                tmp_sql = tmp_sql + " 		 END AS [家主名] ";
                tmp_sql = tmp_sql + " 		/*2016.04.25 ユーザーテスト用に編集 chg end*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [家主名(SJIS)] ";
                tmp_sql = tmp_sql + " 		,so_kana AS [家主カナ] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kojin_kbn = 1 THEN '様' ";
                tmp_sql = tmp_sql + " 			WHEN kojin_kbn = 2 THEN '御中' ";
                tmp_sql = tmp_sql + " 			ELSE '様' ";
                tmp_sql = tmp_sql + " 		 END AS [宛名敬称] ";
                tmp_sql = tmp_sql + " 		,post AS [郵便番号] ";
                tmp_sql = tmp_sql + " 		,add1 AS [住所1] ";
                tmp_sql = tmp_sql + " 		,add2 AS [住所2] ";
                tmp_sql = tmp_sql + " 		,yobi_mj2 AS [登記住所１] ";
                tmp_sql = tmp_sql + " 		,yobi_mj3 AS [登記住所２] ";
                tmp_sql = tmp_sql + " 		,tel1 AS [TEL1] ";
                tmp_sql = tmp_sql + " 		,tel2 AS [TEL2] ";
                tmp_sql = tmp_sql + " 		,fax AS [FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯１] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯２] ";
                tmp_sql = tmp_sql + " 		,email AS [メールアドレス] ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯メールアドレス] ";
                tmp_sql = tmp_sql + " 		,1 AS [優先電話設定] ";
                tmp_sql = tmp_sql + " 		,1 AS [優先メール設定] ";
                tmp_sql = tmp_sql + " 		,birthday AS [誕生日・設立日] ";
                tmp_sql = tmp_sql + " 		,'' AS [年収・年商] ";
                tmp_sql = tmp_sql + " 		,biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,1 AS [性別] ";
                tmp_sql = tmp_sql + " 		,'' AS [本籍地] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先名] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先名SJIS] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先郵便番号] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先住所１] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先住所２] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先TEL１] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先TEL２] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先業種] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先部署] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先情報記入年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [勤務先入社年月] ";
                tmp_sql = tmp_sql + " 		,'' AS [URL] ";
                tmp_sql = tmp_sql + " 		,'' AS [業種] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者名(SJIS)] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [代表者役職] ";
                tmp_sql = tmp_sql + " 		,0 AS [代表者を宛先に含める] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者名(SJIS)] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [担当者部署] ";
                tmp_sql = tmp_sql + " 		,0 AS [担当者を宛先に含める] ";
                tmp_sql = tmp_sql + " 		,'' AS [資本金] ";
                tmp_sql = tmp_sql + " 		,'' AS [従業員数] ";
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
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ([so_name] + [post] + [add1] + [add2] + [tel1] + [tel2] + [fax]) = ([syorui_name] + [syorui_post] + [syorui_add1] + [syorui_add2] + [syorui_tel1] + [syorui_tel2] + [syorui_fax]) THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 4 ";
                tmp_sql = tmp_sql + " 		 END AS [書類送付先区分] ";
                tmp_sql = tmp_sql + " 		,syorui_name AS [送付先名] ";
                tmp_sql = tmp_sql + " 		,'' AS [送付先名(SJIS)] ";
                tmp_sql = tmp_sql + " 		,'' AS [送付先カナ] ";
                tmp_sql = tmp_sql + " 		,'' AS [送付先宛名敬称] ";
                tmp_sql = tmp_sql + " 		,syorui_post AS [送付先郵便番号] ";
                tmp_sql = tmp_sql + " 		,syorui_add1 AS [送付先住所１] ";
                tmp_sql = tmp_sql + " 		,syorui_add2 AS [送付先住所２] ";
                tmp_sql = tmp_sql + " 		,syorui_tel1 AS [送付先TEL１] ";
                tmp_sql = tmp_sql + " 		,syorui_tel2 AS [送付先TEL2] ";
                tmp_sql = tmp_sql + " 		,syorui_fax AS [送付先FAX] ";
                tmp_sql = tmp_sql + " 		,'' AS [送付先備考] ";
                tmp_sql = tmp_sql + " 		,tanto_no AS [イベント担当者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [絞り込みキーワード] ";
                tmp_sql = tmp_sql + " 		,'' AS [会計情報] ";
                tmp_sql = tmp_sql + " 	FROM m_yanu_so ";
                tmp_sql = tmp_sql + " ) AS VW ";
                // 2016.04.25 ユーザーテスト用に抽出クエリを編集 -chg end

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

    #region 家主口座情報

    public class M_yanu_so_koza_Repository
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

                sortstr = "[家主No],[口座No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[家主No],[口座No],[金融機関No],[金融機関支店No],[口座種別],[口座番号],[口座名義],[口座名義カナ],[ゆうちょ口座記号１],[ゆうちょ口座記号２],[ゆうちょ口座番号],[口座備考],[振込情報備考],[総合振込区分],[振込依頼人No],[振込手数料負担区分],[振込手数料計算区分],[振込手数料固定額1],[振込手数料固定額2],[総合振込情報備考],[家主向け帳票の表示]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 so_no AS [家主No] ";
                tmp_sql = tmp_sql + " 		,sokoza_no AS [口座No] ";
                tmp_sql = tmp_sql + " 		,kinyu_no AS [金融機関No] ";
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
                tmp_sql = tmp_sql + " 		,post_kigo AS [ゆうちょ口座記号１] ";
                tmp_sql = tmp_sql + " 		,post_n AS [ゆうちょ口座記号２] ";
                tmp_sql = tmp_sql + " 		,post_bango AS [ゆうちょ口座番号] ";
                tmp_sql = tmp_sql + " 		,YANUKOZA.biko AS [口座備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [振込情報備考] ";
                tmp_sql = tmp_sql + " 		,2 AS [総合振込区分] ";
                tmp_sql = tmp_sql + " 		,fkom_no AS [振込依頼人No] ";
                tmp_sql = tmp_sql + " 		,tesu_kbn AS [振込手数料負担区分] ";
                tmp_sql = tmp_sql + " 		,tesu_keikbn AS [振込手数料計算区分] ";
                tmp_sql = tmp_sql + " 		,ktesu1 AS [振込手数料固定額1] ";
                tmp_sql = tmp_sql + " 		,ktesu2 AS [振込手数料固定額2] ";
                tmp_sql = tmp_sql + " 		,'' AS [総合振込情報備考] ";
                tmp_sql = tmp_sql + " 		,bankpost_kbn AS [家主向け帳票の表示] ";
                tmp_sql = tmp_sql + " 	FROM m_yanu_so_koza AS YANUKOZA ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_kozasyu AS MK ON YANUKOZA.kosyu_no = MK.kosyu_no ";
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

    #region 家主イベント情報

    public class M_yanu_so_event_Repository
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

                sortstr = "[家主No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[家主No],[年賀状],[暑中お見舞い],[誕生日],[お歳暮],[お中元]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 [so_no] AS [家主No] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [event_nenga] = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN [event_nenga] = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 2 ";
                tmp_sql = tmp_sql + " 		 END AS [年賀状] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [event_syocyu] = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN [event_syocyu] = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 2 ";
                tmp_sql = tmp_sql + " 		 END AS [暑中お見舞い] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [event_birthday] = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN [event_birthday] = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 2 ";
                tmp_sql = tmp_sql + " 		 END AS [誕生日] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [event_oseibo] = 1 THEN 6 ";
                tmp_sql = tmp_sql + " 			WHEN [event_oseibo] BETWEEN 2 AND 6 THEN [event_oseibo] -1 ";
                tmp_sql = tmp_sql + " 			ELSE 6 ";
                tmp_sql = tmp_sql + " 		 END AS [お歳暮] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [event_ocyugen] = 1 THEN 6 ";
                tmp_sql = tmp_sql + " 			WHEN [event_ocyugen] BETWEEN 2 AND 6 THEN [event_ocyugen] -1 ";
                tmp_sql = tmp_sql + " 			ELSE 6 ";
                tmp_sql = tmp_sql + " 		 END AS [お中元] ";
                tmp_sql = tmp_sql + " 	FROM m_yanu_so ";
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

    #region 家主メモ情報

    public class M_yanu_so_biko_Repository
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

                sortstr = "[家主No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[家主No],[メモ1],[メモ2],[メモ3],[メモ4],[メモ5],[メモ6],[メモ7],[メモ8],[メモ9],[メモ10],[メモ11],[メモ12],[メモ13],[メモ14],[メモ15],[メモ16],[メモ17],[メモ18],[メモ19],[メモ20],[メモ21],[メモ22],[メモ23],[メモ24],[メモ25],[メモ26],[メモ27],[メモ28],[メモ29],[メモ30],[メモ31],[メモ32],[メモ33],[メモ34],[メモ35],[メモ36],[メモ37],[メモ38],[メモ39],[メモ40],[メモ41],[メモ42],[メモ43],[メモ44],[メモ45],[メモ46],[メモ47],[メモ48],[メモ49],[メモ50]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 so_no AS [家主No] ";
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
                tmp_sql = tmp_sql + " 	FROM m_yanu_so_biko ";
                tmp_sql = tmp_sql + " 	GROUP BY so_no ";
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