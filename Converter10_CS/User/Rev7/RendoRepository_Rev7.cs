using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using Converter10.Njc.Common;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Repository
{

    #region 送信設定基本

    public class Rendo_sosin_Base_Repository
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
                string rendoid = Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(My.MyProject.Forms.MainFrm.txtRendoID.Text), "kariID", My.MyProject.Forms.MainFrm.txtRendoID.Text));  // 20160825 連動ID入力箇所の新規追加 -add

                sortstr = "[送信設定順]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[送信設定順],[送信設定名],[掲載支店],[連動ID],[広告確認からの確認期間],[基準日],[基準日区分],[Eメールアドレス1],[Eメールアドレス2],[Eメールアドレス3],[Eメールアドレス4],[Eメールアドレス5]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		/*20160829 連動中間ファイル作成処理修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/* ROW_NUMBER()OVER(ORDER BY [No]) AS [送信設定順]*/ ";
                tmp_sql = tmp_sql + " 		 [No] AS [送信設定順] ";
                tmp_sql = tmp_sql + " 		/*20160829 連動中間ファイル作成処理修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		,Comment AS [送信設定名] ";
                tmp_sql = tmp_sql + " 		,0 AS [掲載支店]	/*取得方法を考慮する必要がある*/ ";
                tmp_sql = tmp_sql + " 		/*20160825 連動ID入力箇所の新規追加 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,830013 AS [連動ID]*/ ";
                tmp_sql = tmp_sql + " 		,'連動ID置換用文字列' AS [連動ID] ";
                tmp_sql = tmp_sql + " 		/*20160825 連動ID入力箇所の新規追加 chg end*/ ";
                tmp_sql = tmp_sql + " 		,90 AS [広告確認からの確認期間] ";
                tmp_sql = tmp_sql + " 		,NULL AS [基準日] ";
                tmp_sql = tmp_sql + " 		,1 AS [基準日区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [Eメールアドレス1] ";
                tmp_sql = tmp_sql + " 		,NULL AS [Eメールアドレス2] ";
                tmp_sql = tmp_sql + " 		,NULL AS [Eメールアドレス3] ";
                tmp_sql = tmp_sql + " 		,NULL AS [Eメールアドレス4] ";
                tmp_sql = tmp_sql + " 		,NULL AS [Eメールアドレス5] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*連動必須項目が存在するデータのみ抽出する sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM CVTBL_物件データ連動情報 AS VW ";
                tmp_sql = tmp_sql + " 		WHERE [Enabled] <> 0 ";
                tmp_sql = tmp_sql + " 		/*20160908 連動_送信設定情報抽出クエリ修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*EXISTS句を全てANDで結んでいたのでORへ修正*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		AND EXISTS ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1081002') AS SETTEI		/*自社WEB必須項目(ログインID)*/ ";
                tmp_sql = tmp_sql + " 					WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) ";
                tmp_sql = tmp_sql + " 					AND   ISNULL(SETTEI.set_value1,'') <> '' ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 		AND EXISTS ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001021') AS SETTEI		/*HOMES必須項目(会員番号)*/ ";
                tmp_sql = tmp_sql + " 					WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) ";
                tmp_sql = tmp_sql + " 					AND   ISNULL(SETTEI.set_value1,'') <> '' ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 		AND EXISTS ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1061002') AS SETTEI		/*athome必須項目(店舗ID)*/ ";
                tmp_sql = tmp_sql + " 					WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) ";
                tmp_sql = tmp_sql + " 					AND   ISNULL(SETTEI.set_value1,'') <> '' ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 		AND EXISTS ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001061') AS SETTEI		/*SUUMO必須項目(支店コード)*/ ";
                tmp_sql = tmp_sql + " 					WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) ";
                tmp_sql = tmp_sql + " 					AND   ISNULL(SETTEI.set_value1,'') <> '' ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 		/*連動必須項目が存在するデータのみ抽出する end*/ ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		AND ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				EXISTS ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1081002') AS SETTEI		/*自社WEB必須項目(ログインID)*/ ";
                tmp_sql = tmp_sql + " 							WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) ";
                tmp_sql = tmp_sql + " 							AND   ISNULL(SETTEI.set_value1,'') <> '' ";
                tmp_sql = tmp_sql + " 						) ";
                tmp_sql = tmp_sql + " 				OR EXISTS ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001021') AS SETTEI		/*HOMES必須項目(会員番号)*/ ";
                tmp_sql = tmp_sql + " 							WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) ";
                tmp_sql = tmp_sql + " 							AND   ISNULL(SETTEI.set_value1,'') <> '' ";
                tmp_sql = tmp_sql + " 						) ";
                tmp_sql = tmp_sql + " 				OR EXISTS ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1061002') AS SETTEI		/*athome必須項目(店舗ID)*/ ";
                tmp_sql = tmp_sql + " 							WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) ";
                tmp_sql = tmp_sql + " 							AND   ISNULL(SETTEI.set_value1,'') <> '' ";
                tmp_sql = tmp_sql + " 						) ";
                tmp_sql = tmp_sql + " 				OR EXISTS ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001061') AS SETTEI		/*SUUMO必須項目(支店コード)*/ ";
                tmp_sql = tmp_sql + " 							WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) ";
                tmp_sql = tmp_sql + " 							AND   ISNULL(SETTEI.set_value1,'') <> '' ";
                tmp_sql = tmp_sql + " 						) ";
                tmp_sql = tmp_sql + " 			) ";
                tmp_sql = tmp_sql + " 		/*20160908 連動_送信設定情報抽出クエリ修正 chg end*/ ";
                tmp_sql = tmp_sql + " 	) AS TMP ";
                tmp_sql = tmp_sql + " ) AS VW ";

                // 20160825 連動ID入力箇所の新規追加 -add
                tmp_sql = tmp_sql.Replace("連動ID置換用文字列", rendoid);

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

    #region 送信設定自社web

    public class Rendo_sosin_jisyaweb_Repository
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

                sortstr = "[送信設定順],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[送信設定順],[サイトNo],[ポータルサイトID],[サイト別送信有無],[ポータルサイトパスワード],[Eメールアドレス1],[Eメールアドレス2],[Eメールアドレス3],[Eメールアドレス4],[Eメールアドレス5],[物件部屋の表示],[番地以降の表示],[動画の表示],[動画の優先項目1],[動画の優先項目2],[動画の優先項目3],[動画の優先項目4],[くらさぽへの地図表示],[くらさぽ以外への地図表示],[敷金表示単位],[礼金表示単位],[保証金表示単位],[償却金表示単位],[敷引(解約引き)表示単位],[更新料表示単位],[仲介手数料表示単位],[その他一時金表示単位],[入金項目_賃料1],[入金項目区分_賃料1],[入金項目_賃料2],[入金項目区分_賃料2],[入金項目_賃料3],[入金項目区分_賃料3],[入金項目_共益費/管理費1],[入金項目区分_共益費/管理費1],[入金項目_共益費/管理費2],[入金項目区分_共益費/管理費2],[入金項目_共益費/管理費3],[入金項目区分_共益費/管理費3],[入金項目_敷金],[入金項目区分_敷金],[入金項目_礼金],[入金項目区分_礼金],[入金項目_保証金],[入金項目区分_保証金],[入金項目_償却金],[入金項目区分_償却金],[入金項目_敷引(解約引き)],[入金項目区分_敷引(解約引き)],[入金項目_更新料],[入金項目区分_更新料],[入金項目_仲介手数料],[入金項目区分_仲介手数料],[入金項目_その他一時金],[入金項目区分_その他一時金],[入金項目_その他月額費用1],[入金項目区分_その他月額費用1],[入金項目_その他月額費用2],[入金項目区分_その他月額費用2],[画像No1],[革命側の画像種別1],[革命側の画像タイトルNo1],[連動側の画像種別1],[画像No2],[革命側の画像種別2],[革命側の画像タイトルNo2],[連動側の画像種別2],[画像No3],[革命側の画像種別3],[革命側の画像タイトルNo3],[連動側の画像種別3],[画像No4],[革命側の画像種別4],[革命側の画像タイトルNo4],[連動側の画像種別4],[画像No5],[革命側の画像種別5],[革命側の画像タイトルNo5],[連動側の画像種別5],[画像No6],[革命側の画像種別6],[革命側の画像タイトルNo6],[連動側の画像種別6],[画像No7],[革命側の画像種別7],[革命側の画像タイトルNo7],[連動側の画像種別7],[画像No8],[革命側の画像種別8],[革命側の画像タイトルNo8],[連動側の画像種別8],[画像No9],[革命側の画像種別9],[革命側の画像タイトルNo9],[連動側の画像種別9],[画像No10],[革命側の画像種別10],[革命側の画像タイトルNo10],[連動側の画像種別10],[画像No11],[革命側の画像種別11],[革命側の画像タイトルNo11],[連動側の画像種別11],[画像No12],[革命側の画像種別12],[革命側の画像タイトルNo12],[連動側の画像種別12],[画像No13],[革命側の画像種別13],[革命側の画像タイトルNo13],[連動側の画像種別13],[画像No14],[革命側の画像種別14],[革命側の画像タイトルNo14],[連動側の画像種別14],[画像No15],[革命側の画像種別15],[革命側の画像タイトルNo15],[連動側の画像種別15],[画像No16],[革命側の画像種別16],[革命側の画像タイトルNo16],[連動側の画像種別16],[画像No17],[革命側の画像種別17],[革命側の画像タイトルNo17],[連動側の画像種別17],[画像No18],[革命側の画像種別18],[革命側の画像タイトルNo18],[連動側の画像種別18],[画像No19],[革命側の画像種別19],[革命側の画像タイトルNo19],[連動側の画像種別19],[画像No20],[革命側の画像種別20],[革命側の画像タイトルNo20],[連動側の画像種別20],[画像No21],[革命側の画像種別21],[革命側の画像タイトルNo21],[連動側の画像種別21],[画像No22],[革命側の画像種別22],[革命側の画像タイトルNo22],[連動側の画像種別22],[画像No23],[革命側の画像種別23],[革命側の画像タイトルNo23],[連動側の画像種別23],[画像No24],[革命側の画像種別24],[革命側の画像タイトルNo24],[連動側の画像種別24],[画像No25],[革命側の画像種別25],[革命側の画像タイトルNo25],[連動側の画像種別25],[画像No26],[革命側の画像種別26],[革命側の画像タイトルNo26],[連動側の画像種別26],[画像No27],[革命側の画像種別27],[革命側の画像タイトルNo27],[連動側の画像種別27],[画像No28],[革命側の画像種別28],[革命側の画像タイトルNo28],[連動側の画像種別28],[画像No29],[革命側の画像種別29],[革命側の画像タイトルNo29],[連動側の画像種別29],[画像No30],[革命側の画像種別30],[革命側の画像タイトルNo30],[連動側の画像種別30],[鍵情報1],[鍵情報2],[鍵情報3]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                for (int cntii = 0; cntii <= 9; cntii++)
                {

                    tmp_sql = tmp_sql + " SELECT * FROM ";
                    tmp_sql = tmp_sql + " ( ";
                    tmp_sql = tmp_sql + " 	SELECT ";
                    tmp_sql = tmp_sql + "       /*20160829 連動中間ファイル作成処理修正 chg sta*/ ";
                    tmp_sql = tmp_sql + "       /* 送信設定No置換文字列 AS [送信設定順_V7]*/ ";
                    tmp_sql = tmp_sql + "        送信設定No置換文字列 AS [送信設定順]  ";
                    tmp_sql = tmp_sql + "       /*20160829 連動中間ファイル作成処理修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 		,10 AS [サイトNo] ";
                    tmp_sql = tmp_sql + " 		/*******************************************フィールドへ移行するデータ sta*******************************************/ ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1081002') AS [ポータルサイトID] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01081001','11081001','21081001','31081001','41081001','51081001','61081001','71081001','81081001','91081001')) = 0 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01081001','11081001','21081001','31081001','41081001','51081001','61081001','71081001','81081001','91081001')) = 1 THEN 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [サイト別送信有無] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1081006') AS [ポータルサイトパスワード] ";
                    tmp_sql = tmp_sql + " 		,'' AS [Eメールアドレス1] ";
                    tmp_sql = tmp_sql + " 		,'' AS [Eメールアドレス2] ";
                    tmp_sql = tmp_sql + " 		,'' AS [Eメールアドレス3] ";
                    tmp_sql = tmp_sql + " 		,'' AS [Eメールアドレス4] ";
                    tmp_sql = tmp_sql + " 		,'' AS [Eメールアドレス5] ";
                    tmp_sql = tmp_sql + " 		/*******************************************フィールドへ移行するデータ end*******************************************/ ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		/************************************************xmlへ移行するデータ sta************************************************/ ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		/***********表示設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1081003') = 1 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1081003') = 2 THEN 3 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1081003') = 3 THEN 4 ";
                    tmp_sql = tmp_sql + " 		 END AS [物件部屋の表示] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1081004') = 1 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1081004') = 2 THEN 2 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1081004') = 3 THEN 3 ";
                    tmp_sql = tmp_sql + " 		 END AS [番地以降の表示]	 ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 		 /********************************************************************************** ";
                    tmp_sql = tmp_sql + " 		 動画の表示に関して ";
                    tmp_sql = tmp_sql + " 		  V7では送信しない選択が可能だが10では不可になっている。そのため、1つでも送信しない ";
                    tmp_sql = tmp_sql + " 		  データが存在する場合は全て表示しないとする。 ";
                    tmp_sql = tmp_sql + " 		  ※V7で1データ送信する設定の場合、全て表示しない設定にしないと優先順位の2つ目が ";
                    tmp_sql = tmp_sql + " 			送信されてしまう可能性があるため ";
                    tmp_sql = tmp_sql + " 		 **********************************************************************************/ ";
                    tmp_sql = tmp_sql + " 			WHEN ( ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081101') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081102') <> 0 AND ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081103') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081104') <> 0 ";
                    tmp_sql = tmp_sql + " 				 ) THEN 2 ";
                    tmp_sql = tmp_sql + " 			ELSE 1 ";
                    tmp_sql = tmp_sql + " 		 END AS [動画の表示] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN ( ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081101') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081102') <> 0 AND ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081103') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081104') <> 0 ";
                    tmp_sql = tmp_sql + " 				 ) THEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081101') ";
                    tmp_sql = tmp_sql + " 			ELSE NULL ";
                    tmp_sql = tmp_sql + " 		 END AS [動画の優先項目1] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN ( ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081101') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081102') <> 0 AND ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081103') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081104') <> 0 ";
                    tmp_sql = tmp_sql + " 				 ) THEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081102') ";
                    tmp_sql = tmp_sql + " 			ELSE NULL ";
                    tmp_sql = tmp_sql + " 		 END AS [動画の優先項目2] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN ( ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081101') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081102') <> 0 AND ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081103') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081104') <> 0 ";
                    tmp_sql = tmp_sql + " 				 ) THEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081103') ";
                    tmp_sql = tmp_sql + " 			ELSE NULL ";
                    tmp_sql = tmp_sql + " 		 END AS [動画の優先項目3] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN ( ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081101') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081102') <> 0 AND ";
                    tmp_sql = tmp_sql + " 					 (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081103') <> 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081104') <> 0 ";
                    tmp_sql = tmp_sql + " 				 ) THEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列7081104') ";
                    tmp_sql = tmp_sql + " 			ELSE NULL ";
                    tmp_sql = tmp_sql + " 		 END AS [動画の優先項目4]	 	 	 ";
                    tmp_sql = tmp_sql + " 		,2 AS [くらさぽへの地図表示]		/*要デフォルト値確認*/ ";
                    tmp_sql = tmp_sql + " 		,2 AS [くらさぽ以外への地図表示]	/*要デフォルト値確認*/ ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		/***********入金項目タブ***********/ ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [賃料表示単位]*/				/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [共益費/管理費表示単位]*/		/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081051') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081051') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [敷金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081052') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081052') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [礼金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081053') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081053') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [保証金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081054') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081054') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [償却金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081055') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081055') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [敷引(解約引き)表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081057') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081057') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [更新料表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081058') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081058') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [仲介手数料表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081056') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081056') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [その他一時金表示単位] ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [その他月額費用表示単位]*/				/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081001')) AS [入金項目_賃料1] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081002')) AS [入金項目_賃料2] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081003')) AS [入金項目_賃料3] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081004')) AS [入金項目_共益費/管理費1] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_共益費/管理費1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081005')) AS [入金項目_共益費/管理費2] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_共益費/管理費2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081006')) AS [入金項目_共益費/管理費3] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_共益費/管理費3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081009')) AS [入金項目_敷金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_敷金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081010')) AS [入金項目_礼金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_礼金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081011')) AS [入金項目_保証金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_保証金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081012')) AS [入金項目_償却金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_償却金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081013')) AS [入金項目_敷引(解約引き)] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_敷引(解約引き)] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081015')) AS [入金項目_更新料] ";
                    tmp_sql = tmp_sql + " 		,3 AS [入金項目区分_更新料] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081017')) AS [入金項目_仲介手数料] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_仲介手数料] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081014')) AS [入金項目_その他一時金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_その他一時金]	  ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081007')) AS [入金項目_その他月額費用1] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_その他月額費用1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081008')) AS [入金項目_その他月額費用2] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_その他月額費用2] ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		 /***********画像設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		,1 AS [画像No1] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081001'),1) AS [革命側の画像種別1] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081001'),2)) AS [革命側の画像タイトルNo1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081101') AS [連動側の画像種別1] ";
                    tmp_sql = tmp_sql + " 		,2 AS [画像No2] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081002'),1) AS [革命側の画像種別2] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081002'),2)) AS [革命側の画像タイトルNo2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081102') AS [連動側の画像種別2] ";
                    tmp_sql = tmp_sql + " 		,3 AS [画像No3] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081003'),1) AS [革命側の画像種別3] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081003'),2)) AS [革命側の画像タイトルNo3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081103') AS [連動側の画像種別3] ";
                    tmp_sql = tmp_sql + " 		,4 AS [画像No4] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081004'),1) AS [革命側の画像種別4] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081004'),2)) AS [革命側の画像タイトルNo4] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081104') AS [連動側の画像種別4] ";
                    tmp_sql = tmp_sql + " 		,5 AS [画像No5] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081005'),1) AS [革命側の画像種別5] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081005'),2)) AS [革命側の画像タイトルNo5] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081105') AS [連動側の画像種別5] ";
                    tmp_sql = tmp_sql + " 		,6 AS [画像No6] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081006'),1) AS [革命側の画像種別6] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081006'),2)) AS [革命側の画像タイトルNo6] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081106') AS [連動側の画像種別6] ";
                    tmp_sql = tmp_sql + " 		,7 AS [画像No7] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081007'),1) AS [革命側の画像種別7] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081007'),2)) AS [革命側の画像タイトルNo7] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081107') AS [連動側の画像種別7] ";
                    tmp_sql = tmp_sql + " 		,8 AS [画像No8] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081008'),1) AS [革命側の画像種別8] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081008'),2)) AS [革命側の画像タイトルNo8] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081108') AS [連動側の画像種別8] ";
                    tmp_sql = tmp_sql + " 		,9 AS [画像No9] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081009'),1) AS [革命側の画像種別9] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081009'),2)) AS [革命側の画像タイトルNo9] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081109') AS [連動側の画像種別9] ";
                    tmp_sql = tmp_sql + " 		,10 AS [画像No10] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081010'),1) AS [革命側の画像種別10] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081010'),2)) AS [革命側の画像タイトルNo10] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081110') AS [連動側の画像種別10] ";
                    tmp_sql = tmp_sql + " 		,11 AS [画像No11] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081011'),1) AS [革命側の画像種別11] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081011'),2)) AS [革命側の画像タイトルNo11] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081111') AS [連動側の画像種別11] ";
                    tmp_sql = tmp_sql + " 		,12 AS [画像No12] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081012'),1) AS [革命側の画像種別12] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081012'),2)) AS [革命側の画像タイトルNo12] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081112') AS [連動側の画像種別12] ";
                    tmp_sql = tmp_sql + " 		,13 AS [画像No13] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081013'),1) AS [革命側の画像種別13] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081013'),2)) AS [革命側の画像タイトルNo13] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081113') AS [連動側の画像種別13] ";
                    tmp_sql = tmp_sql + " 		,14 AS [画像No14] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081014'),1) AS [革命側の画像種別14] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081014'),2)) AS [革命側の画像タイトルNo14] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081114') AS [連動側の画像種別14] ";
                    tmp_sql = tmp_sql + " 		,15 AS [画像No15] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081015'),1) AS [革命側の画像種別15] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081015'),2)) AS [革命側の画像タイトルNo15] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081115') AS [連動側の画像種別15] ";
                    tmp_sql = tmp_sql + " 		,16 AS [画像No16] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081016'),1) AS [革命側の画像種別16] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081016'),2)) AS [革命側の画像タイトルNo16] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081116') AS [連動側の画像種別16] ";
                    tmp_sql = tmp_sql + " 		,17 AS [画像No17] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081017'),1) AS [革命側の画像種別17] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081017'),2)) AS [革命側の画像タイトルNo17] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081117') AS [連動側の画像種別17] ";
                    tmp_sql = tmp_sql + " 		,18 AS [画像No18] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081018'),1) AS [革命側の画像種別18] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081018'),2)) AS [革命側の画像タイトルNo18] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081118') AS [連動側の画像種別18] ";
                    tmp_sql = tmp_sql + " 		,19 AS [画像No19] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081019'),1) AS [革命側の画像種別19] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081019'),2)) AS [革命側の画像タイトルNo19] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081119') AS [連動側の画像種別19] ";
                    tmp_sql = tmp_sql + " 		,20 AS [画像No20] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081020'),1) AS [革命側の画像種別20] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081020'),2)) AS [革命側の画像タイトルNo20] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081120') AS [連動側の画像種別20] ";
                    tmp_sql = tmp_sql + " 		,21 AS [画像No21] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081021'),1) AS [革命側の画像種別21] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081021'),2)) AS [革命側の画像タイトルNo21] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081121') AS [連動側の画像種別21] ";
                    tmp_sql = tmp_sql + " 		,22 AS [画像No22] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081022'),1) AS [革命側の画像種別22] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081022'),2)) AS [革命側の画像タイトルNo22] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081122') AS [連動側の画像種別22] ";
                    tmp_sql = tmp_sql + " 		,23 AS [画像No23] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081023'),1) AS [革命側の画像種別23] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081023'),2)) AS [革命側の画像タイトルNo23] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081123') AS [連動側の画像種別23] ";
                    tmp_sql = tmp_sql + " 		,24 AS [画像No24] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081024'),1) AS [革命側の画像種別24] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081024'),2)) AS [革命側の画像タイトルNo24] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081124') AS [連動側の画像種別24] ";
                    tmp_sql = tmp_sql + " 		,25 AS [画像No25] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081025'),1) AS [革命側の画像種別25] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081025'),2)) AS [革命側の画像タイトルNo25] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081125') AS [連動側の画像種別25] ";
                    tmp_sql = tmp_sql + " 		,26 AS [画像No26] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081026'),1) AS [革命側の画像種別26] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081026'),2)) AS [革命側の画像タイトルNo26] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081126') AS [連動側の画像種別26] ";
                    tmp_sql = tmp_sql + " 		,27 AS [画像No27] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081027'),1) AS [革命側の画像種別27] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081027'),2)) AS [革命側の画像タイトルNo27] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081127') AS [連動側の画像種別27] ";
                    tmp_sql = tmp_sql + " 		,28 AS [画像No28] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081028'),1) AS [革命側の画像種別28] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081028'),2)) AS [革命側の画像タイトルNo28] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081128') AS [連動側の画像種別28] ";
                    tmp_sql = tmp_sql + " 		,29 AS [画像No29] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081029'),1) AS [革命側の画像種別29] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081029'),2)) AS [革命側の画像タイトルNo29] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081129') AS [連動側の画像種別29] ";
                    tmp_sql = tmp_sql + " 		,30 AS [画像No30] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081030'),1) AS [革命側の画像種別30] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081030'),2)) AS [革命側の画像タイトルNo30] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3081130') AS [連動側の画像種別30] ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		 /***********BB設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		 ,'-1' AS [鍵情報1] ";
                    tmp_sql = tmp_sql + " 		 ,'-1' AS [鍵情報2] ";
                    tmp_sql = tmp_sql + " 		 ,'-1' AS [鍵情報3] ";
                    tmp_sql = tmp_sql + " 		  ";
                    tmp_sql = tmp_sql + " 		 /***********その他タブ***********/ ";
                    tmp_sql = tmp_sql + " 		  ";
                    tmp_sql = tmp_sql + " 		 /***********優先設備タブ***********/ ";
                    tmp_sql = tmp_sql + " 		  ";
                    tmp_sql = tmp_sql + " 		 /***********連動項目設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		  ";
                    tmp_sql = tmp_sql + " 		/************************************************xmlへ移行するデータ end************************************************/ ";
                    tmp_sql = tmp_sql + " ) AS VW ";

                    tmp_sql = tmp_sql.Replace("送信設定No置換文字列", (cntii + 1).ToString());
                    tmp_sql = tmp_sql.Replace("V7送信設定値取得用置換文字列", cntii.ToString());

                    if (cntii < 9)
                    {
                        tmp_sql = tmp_sql + " UNION ";
                    }

                }

                // 20160908 連動_送信設定情報抽出クエリ修正 -add
                tmp_sql = " SELECT * FROM ( " + tmp_sql + " ) AS VW " + " WHERE [ポータルサイトID] <> '' ";

                // 20160829 連動中間ファイル作成処理修正 -del sta
                // Dim tmp_sql_base As String = ""
                // tmp_sql_base = tmp_sql_base & " SELECT [No] FROM CVTBL_物件データ連動情報 AS VW "
                // tmp_sql_base = tmp_sql_base & " WHERE [Enabled] <> 0 "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1081002') AS SETTEI		/*自社WEB必須項目(ログインID)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001021') AS SETTEI		/*HOMES必須項目(会員番号)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1061002') AS SETTEI		/*athome必須項目(店舗ID)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001061') AS SETTEI		/*SUUMO必須項目(支店コード)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "

                // 'tmp_sql = " SELECT * FROM ( " & tmp_sql & ") AS VW WHERE [送信設定順_V7] IN (SELECT [No] FROM CVTBL_物件データ連動情報 WHERE [Enabled] <> 0) "
                // tmp_sql = " SELECT * FROM ( " & tmp_sql & ") AS VW WHERE [送信設定順_V7] IN (" & tmp_sql_base & ") "
                // tmp_sql = " SELECT ROW_NUMBER()OVER(ORDER BY [送信設定順_V7]) AS [送信設定順],* FROM ( " & tmp_sql & ") AS VW "
                // tmp_sql = " SELECT " & tmp_chgmidheader & " FROM (" & tmp_sql & ") AS VW "
                // 20160829 連動中間ファイル作成処理修正 -del end

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

    #region 送信設定HOMES

    public class Rendo_sosin_homes_Repository
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

                sortstr = "[送信設定順],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[送信設定順],[サイトNo],[ポータルサイトID],[サイト別送信有無],[ポータルサイトパスワード],[FTPユーザーID],[FTPパスワード],[物件部屋の表示],[番地以降の表示],[掲載確認日],[特別広告ポイント数],[パノラマセット],[スタッフコメント],[所属グループ設定1],[所属グループ設定2],[所属グループ設定3],[所属グループ設定4],[所属グループ設定5],[所属グループ設定6],[所属グループ設定7],[所属グループ設定8],[所属グループ設定9],[所属グループ設定10],[敷金表示単位],[礼金表示単位],[保証金表示単位],[権利金表示単位],[償却/敷引金表示単位],[更新料表示単位],[入金項目_賃料1],[入金項目区分_賃料1],[入金項目_賃料2],[入金項目区分_賃料2],[入金項目_賃料3],[入金項目区分_賃料3],[入金項目_管理費/共益費1],[入金項目区分_管理費/共益費1],[入金項目_管理費/共益費2],[入金項目区分_管理費/共益費2],[入金項目_管理費/共益費3],[入金項目区分_管理費/共益費3],[入金項目_敷金],[入金項目区分_敷金],[入金項目_礼金],[入金項目区分_礼金],[入金項目_保証金],[入金項目区分_保証金],[入金項目_権利金],[入金項目区分_権利金],[入金項目_償却/敷引金],[入金項目区分_償却/敷引金],[入金項目_更新料],[入金項目区分_更新料],[入金項目_造作譲渡金],[入金項目区分_造作譲渡金],[入金項目_仲介手数料],[入金項目区分_仲介手数料],[入金項目_鍵交換費用],[入金項目区分_鍵交換費用],[入金項目_室内清掃費用],[入金項目区分_室内清掃費用],[入金項目_その他月額費用1],[入金項目区分_その他費用1],[入金項目_その他月額費用2],[入金項目区分_その他費用2],[入金項目_その他月額費用3],[入金項目区分_その他費用3],[画像No1],[革命側の画像種別1],[革命側の画像タイトルNo1],[連動側の画像種別1],[画像No2],[革命側の画像種別2],[革命側の画像タイトルNo2],[連動側の画像種別2],[画像No3],[革命側の画像種別3],[革命側の画像タイトルNo3],[連動側の画像種別3],[画像No4],[革命側の画像種別4],[革命側の画像タイトルNo4],[連動側の画像種別4],[画像No5],[革命側の画像種別5],[革命側の画像タイトルNo5],[連動側の画像種別5],[画像No6],[革命側の画像種別6],[革命側の画像タイトルNo6],[連動側の画像種別6],[画像No7],[革命側の画像種別7],[革命側の画像タイトルNo7],[連動側の画像種別7],[画像No8],[革命側の画像種別8],[革命側の画像タイトルNo8],[連動側の画像種別8],[画像No9],[革命側の画像種別9],[革命側の画像タイトルNo9],[連動側の画像種別9],[画像No10],[革命側の画像種別10],[革命側の画像タイトルNo10],[連動側の画像種別10],[画像No11],[革命側の画像種別11],[革命側の画像タイトルNo11],[連動側の画像種別11],[画像No12],[革命側の画像種別12],[革命側の画像タイトルNo12],[連動側の画像種別12],[画像No13],[革命側の画像種別13],[革命側の画像タイトルNo13],[連動側の画像種別13],[画像No14],[革命側の画像種別14],[革命側の画像タイトルNo14],[連動側の画像種別14],[画像No15],[革命側の画像種別15],[革命側の画像タイトルNo15],[連動側の画像種別15],[画像No16],[革命側の画像種別16],[革命側の画像タイトルNo16],[連動側の画像種別16],[画像No17],[革命側の画像種別17],[革命側の画像タイトルNo17],[連動側の画像種別17],[画像No18],[革命側の画像種別18],[革命側の画像タイトルNo18],[連動側の画像種別18],[画像No19],[革命側の画像種別19],[革命側の画像タイトルNo19],[連動側の画像種別19],[画像No20],[革命側の画像種別20],[革命側の画像タイトルNo20],[連動側の画像種別20],[画像No21],[革命側の画像種別21],[革命側の画像タイトルNo21],[連動側の画像種別21],[画像No22],[革命側の画像種別22],[革命側の画像タイトルNo22],[連動側の画像種別22],[画像No23],[革命側の画像種別23],[革命側の画像タイトルNo23],[連動側の画像種別23],[画像No24],[革命側の画像種別24],[革命側の画像タイトルNo24],[連動側の画像種別24],[画像No25],[革命側の画像種別25],[革命側の画像タイトルNo25],[連動側の画像種別25],[画像No26],[革命側の画像種別26],[革命側の画像タイトルNo26],[連動側の画像種別26],[画像No27],[革命側の画像種別27],[革命側の画像タイトルNo27],[連動側の画像種別27],[画像No28],[革命側の画像種別28],[革命側の画像タイトルNo28],[連動側の画像種別28],[画像No29],[革命側の画像種別29],[革命側の画像タイトルNo29],[連動側の画像種別29],[画像No30],[革命側の画像種別30],[革命側の画像タイトルNo30],[連動側の画像種別30]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                for (int cntii = 0; cntii <= 9; cntii++)
                {

                    tmp_sql = tmp_sql + " SELECT * FROM ";
                    tmp_sql = tmp_sql + " ( ";
                    tmp_sql = tmp_sql + " 	SELECT ";
                    tmp_sql = tmp_sql + "       /*20160829 連動中間ファイル作成処理修正 chg sta*/ ";
                    tmp_sql = tmp_sql + "       /* 送信設定No置換文字列 AS [送信設定順_V7]*/ ";
                    tmp_sql = tmp_sql + "        送信設定No置換文字列 AS [送信設定順]  ";
                    tmp_sql = tmp_sql + "       /*20160829 連動中間ファイル作成処理修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 		,20 AS [サイトNo] ";
                    tmp_sql = tmp_sql + " 		/*******************************************フィールドへ移行するデータ sta*******************************************/ ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001021') AS [ポータルサイトID] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01001020','11001020','21001020','31001020','41001020','51001020','61001020','71001020','81001020','91001020')) = 0 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01001020','11001020','21001020','31001020','41001020','51001020','61001020','71001020','81001020','91001020')) = 1 THEN 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [サイト別送信有無] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001022') AS [ポータルサイトパスワード] ";
                    tmp_sql = tmp_sql + " 		/*******************************************フィールドへ移行するデータ end*******************************************/ ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		/************************************************xmlへ移行するデータ sta************************************************/ ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		/***********表示設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001023') AS [FTPユーザーID] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001024') AS [FTPパスワード] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001101') = 1 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001101') = 2 THEN 3 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001101') = 3 THEN 4 ";
                    tmp_sql = tmp_sql + " 		 END AS [物件部屋の表示] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001102') = 1 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001102') = 2 THEN 2 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001102') = 3 THEN 3 ";
                    tmp_sql = tmp_sql + " 		 END AS [番地以降の表示] ";
                    tmp_sql = tmp_sql + " 		,1 AS [掲載確認日] ";
                    tmp_sql = tmp_sql + " 		,'false' AS [特別広告ポイント数] ";
                    tmp_sql = tmp_sql + " 		,'false' AS [パノラマセット] ";
                    tmp_sql = tmp_sql + " 		,'false' AS [スタッフコメント] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001111') AS [所属グループ設定1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001112') AS [所属グループ設定2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001113') AS [所属グループ設定3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001114') AS [所属グループ設定4] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001115') AS [所属グループ設定5] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001116') AS [所属グループ設定6] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001117') AS [所属グループ設定7] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001118') AS [所属グループ設定8] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001119') AS [所属グループ設定9] ";
                    tmp_sql = tmp_sql + " 		,'' AS [所属グループ設定10] ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		/***********入金項目タブ***********/ ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [賃料表示単位]*/				/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [管理費/共益費表示単位]*/		/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001151') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001151') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [敷金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001152') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001152') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [礼金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001153') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001153') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [保証金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001156') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001156') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [権利金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001154') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001154') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [償却/敷引金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001155') = 1 THEN '_円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001155') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_円' ";
                    tmp_sql = tmp_sql + " 		 END AS [更新料表示単位] ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [造作譲渡金]*/		/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [仲介手数料]*/		/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [鍵交換費用]*/		/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [室内清掃費用]*/	/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		/*,'_円' AS [その他費用]*/		/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001101')) AS [入金項目_賃料1] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001102')) AS [入金項目_賃料2] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001103')) AS [入金項目_賃料3] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001104')) AS [入金項目_管理費/共益費1] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_管理費/共益費1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001105')) AS [入金項目_管理費/共益費2] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_管理費/共益費2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001106')) AS [入金項目_管理費/共益費3] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_管理費/共益費3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001107')) AS [入金項目_敷金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_敷金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001108')) AS [入金項目_礼金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_礼金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001109')) AS [入金項目_保証金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_保証金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001116')) AS [入金項目_権利金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_権利金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001110')) AS [入金項目_償却/敷引金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_償却/敷引金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001111')) AS [入金項目_更新料] ";
                    tmp_sql = tmp_sql + " 		,3 AS [入金項目区分_更新料] ";
                    tmp_sql = tmp_sql + " 		/* ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5081017')) AS [入金項目_造作譲渡金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_造作譲渡金] ";
                    tmp_sql = tmp_sql + " 		*/ ";
                    tmp_sql = tmp_sql + " 		,NULL AS [入金項目_造作譲渡金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_造作譲渡金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001117')) AS [入金項目_仲介手数料] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_仲介手数料] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001118')) AS [入金項目_鍵交換費用] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_鍵交換費用] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001119')) AS [入金項目_室内清掃費用] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_室内清掃費用]		 ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001113')) AS [入金項目_その他月額費用1] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_その他費用1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001114')) AS [入金項目_その他月額費用2] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_その他費用2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001115')) AS [入金項目_その他月額費用3] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_その他費用3] ";
                    tmp_sql = tmp_sql + " 				 ";
                    tmp_sql = tmp_sql + " 		 /***********画像設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		,1 AS [画像No1] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009101'),1) AS [革命側の画像種別1] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009101'),2)) AS [革命側の画像タイトルNo1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009151') AS [連動側の画像種別1] ";
                    tmp_sql = tmp_sql + " 		,2 AS [画像No2] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009102'),1) AS [革命側の画像種別2] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009102'),2)) AS [革命側の画像タイトルNo2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009152') AS [連動側の画像種別2] ";
                    tmp_sql = tmp_sql + " 		,3 AS [画像No3] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009103'),1) AS [革命側の画像種別3] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009103'),2)) AS [革命側の画像タイトルNo3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009153') AS [連動側の画像種別3] ";
                    tmp_sql = tmp_sql + " 		,4 AS [画像No4] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009104'),1) AS [革命側の画像種別4] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009104'),2)) AS [革命側の画像タイトルNo4] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009154') AS [連動側の画像種別4] ";
                    tmp_sql = tmp_sql + " 		,5 AS [画像No5] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009105'),1) AS [革命側の画像種別5] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009105'),2)) AS [革命側の画像タイトルNo5] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009155') AS [連動側の画像種別5] ";
                    tmp_sql = tmp_sql + " 		,6 AS [画像No6] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009106'),1) AS [革命側の画像種別6] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009106'),2)) AS [革命側の画像タイトルNo6] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009156') AS [連動側の画像種別6] ";
                    tmp_sql = tmp_sql + " 		,7 AS [画像No7] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009107'),1) AS [革命側の画像種別7] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009107'),2)) AS [革命側の画像タイトルNo7] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009157') AS [連動側の画像種別7] ";
                    tmp_sql = tmp_sql + " 		,8 AS [画像No8] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009108'),1) AS [革命側の画像種別8] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009108'),2)) AS [革命側の画像タイトルNo8] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009158') AS [連動側の画像種別8] ";
                    tmp_sql = tmp_sql + " 		,9 AS [画像No9] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009109'),1) AS [革命側の画像種別9] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009109'),2)) AS [革命側の画像タイトルNo9] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009159') AS [連動側の画像種別9] ";
                    tmp_sql = tmp_sql + " 		,10 AS [画像No10] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009110'),1) AS [革命側の画像種別10] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009110'),2)) AS [革命側の画像タイトルNo10] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009160') AS [連動側の画像種別10] ";
                    tmp_sql = tmp_sql + " 		,11 AS [画像No11] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009111'),1) AS [革命側の画像種別11] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009111'),2)) AS [革命側の画像タイトルNo11] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009161') AS [連動側の画像種別11] ";
                    tmp_sql = tmp_sql + " 		,12 AS [画像No12] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009112'),1) AS [革命側の画像種別12] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009112'),2)) AS [革命側の画像タイトルNo12] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009162') AS [連動側の画像種別12] ";
                    tmp_sql = tmp_sql + " 		,13 AS [画像No13] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009113'),1) AS [革命側の画像種別13] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009113'),2)) AS [革命側の画像タイトルNo13] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009163') AS [連動側の画像種別13] ";
                    tmp_sql = tmp_sql + " 		,14 AS [画像No14] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009114'),1) AS [革命側の画像種別14] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009114'),2)) AS [革命側の画像タイトルNo14] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009164') AS [連動側の画像種別14] ";
                    tmp_sql = tmp_sql + " 		,15 AS [画像No15] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009115'),1) AS [革命側の画像種別15] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009115'),2)) AS [革命側の画像タイトルNo15] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009165') AS [連動側の画像種別15] ";
                    tmp_sql = tmp_sql + " 		,16 AS [画像No16] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009116'),1) AS [革命側の画像種別16] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009116'),2)) AS [革命側の画像タイトルNo16] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009166') AS [連動側の画像種別16] ";
                    tmp_sql = tmp_sql + " 		,17 AS [画像No17] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009117'),1) AS [革命側の画像種別17] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009117'),2)) AS [革命側の画像タイトルNo17] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009167') AS [連動側の画像種別17] ";
                    tmp_sql = tmp_sql + " 		,18 AS [画像No18] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009118'),1) AS [革命側の画像種別18] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009118'),2)) AS [革命側の画像タイトルNo18] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009168') AS [連動側の画像種別18] ";
                    tmp_sql = tmp_sql + " 		,19 AS [画像No19] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009119'),1) AS [革命側の画像種別19] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009119'),2)) AS [革命側の画像タイトルNo19] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009169') AS [連動側の画像種別19] ";
                    tmp_sql = tmp_sql + " 		,20 AS [画像No20] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009120'),1) AS [革命側の画像種別20] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009120'),2)) AS [革命側の画像タイトルNo20] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009170') AS [連動側の画像種別20] ";
                    tmp_sql = tmp_sql + " 		,21 AS [画像No21] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009121'),1) AS [革命側の画像種別21] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009121'),2)) AS [革命側の画像タイトルNo21] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009171') AS [連動側の画像種別21] ";
                    tmp_sql = tmp_sql + " 		,22 AS [画像No22] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009122'),1) AS [革命側の画像種別22] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009122'),2)) AS [革命側の画像タイトルNo22] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009172') AS [連動側の画像種別22] ";
                    tmp_sql = tmp_sql + " 		,23 AS [画像No23] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009123'),1) AS [革命側の画像種別23] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009123'),2)) AS [革命側の画像タイトルNo23] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009173') AS [連動側の画像種別23] ";
                    tmp_sql = tmp_sql + " 		,24 AS [画像No24] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009124'),1) AS [革命側の画像種別24] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009124'),2)) AS [革命側の画像タイトルNo24] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009174') AS [連動側の画像種別24] ";
                    tmp_sql = tmp_sql + " 		,25 AS [画像No25] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009125'),1) AS [革命側の画像種別25] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009125'),2)) AS [革命側の画像タイトルNo25] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009175') AS [連動側の画像種別25] ";
                    tmp_sql = tmp_sql + " 		,26 AS [画像No26] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009126'),1) AS [革命側の画像種別26] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009126'),2)) AS [革命側の画像タイトルNo26] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009176') AS [連動側の画像種別26] ";
                    tmp_sql = tmp_sql + " 		,27 AS [画像No27] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009127'),1) AS [革命側の画像種別27] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009127'),2)) AS [革命側の画像タイトルNo27] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009177') AS [連動側の画像種別27] ";
                    tmp_sql = tmp_sql + " 		,28 AS [画像No28] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009128'),1) AS [革命側の画像種別28] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009128'),2)) AS [革命側の画像タイトルNo28] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009178') AS [連動側の画像種別28] ";
                    tmp_sql = tmp_sql + " 		,29 AS [画像No29] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009129'),1) AS [革命側の画像種別29] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009129'),2)) AS [革命側の画像タイトルNo29] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009179') AS [連動側の画像種別29] ";
                    tmp_sql = tmp_sql + " 		,30 AS [画像No30] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009130'),1) AS [革命側の画像種別30] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009130'),2)) AS [革命側の画像タイトルNo30] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009180') AS [連動側の画像種別30] ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		 /***********優先設備タブ***********/ ";
                    tmp_sql = tmp_sql + " 		  ";
                    tmp_sql = tmp_sql + " 		 /***********連動項目設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		  ";
                    tmp_sql = tmp_sql + " 		/************************************************xmlへ移行するデータ end************************************************/ ";
                    tmp_sql = tmp_sql + " ) AS VW ";

                    tmp_sql = tmp_sql.Replace("送信設定No置換文字列", (cntii + 1).ToString());
                    tmp_sql = tmp_sql.Replace("V7送信設定値取得用置換文字列", cntii.ToString());

                    if (cntii < 9)
                    {
                        tmp_sql = tmp_sql + " UNION ";
                    }

                }

                // 20160908 連動_送信設定情報抽出クエリ修正 -add
                tmp_sql = " SELECT * FROM ( " + tmp_sql + " ) AS VW " + " WHERE [ポータルサイトID] <> '' ";

                // 20160829 連動中間ファイル作成処理修正 -del sta
                // Dim tmp_sql_base As String = ""
                // tmp_sql_base = tmp_sql_base & " SELECT [No] FROM CVTBL_物件データ連動情報 AS VW "
                // tmp_sql_base = tmp_sql_base & " WHERE [Enabled] <> 0 "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1081002') AS SETTEI		/*自社WEB必須項目(ログインID)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001021') AS SETTEI		/*HOMES必須項目(会員番号)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1061002') AS SETTEI		/*athome必須項目(店舗ID)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001061') AS SETTEI		/*SUUMO必須項目(支店コード)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "

                // 'tmp_sql = " SELECT * FROM ( " & tmp_sql & ") AS VW WHERE [送信設定順_V7] IN (SELECT [No] FROM CVTBL_物件データ連動情報 WHERE [Enabled] <> 0) "
                // tmp_sql = " SELECT * FROM ( " & tmp_sql & ") AS VW WHERE [送信設定順_V7] IN (" & tmp_sql_base & ") "
                // tmp_sql = " SELECT ROW_NUMBER()OVER(ORDER BY [送信設定順_V7]) AS [送信設定順],* FROM ( " & tmp_sql & ") AS VW "
                // tmp_sql = " SELECT " & tmp_chgmidheader & " FROM (" & tmp_sql & ") AS VW "
                // 20160829 連動中間ファイル作成処理修正 -del end
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

    #region 送信設定athome

    public class Rendo_sosin_athome_Repository
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

                sortstr = "[送信設定順],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[送信設定順],[サイトNo],[ポータルサイトID],[サイト別送信有無],[ポータルサイトパスワード],[店舗ID],[物件名・部屋Noの表示(一般ユーザー向け)],[物件名・部屋Noの表示(不動産会社向け)],[地図表示],[番地以降の表示],[掲載確認日],[敷金表示単位],[礼金表示単位],[保証金表示単位],[保証金償却表示単位],[敷引金表示単位],[更新料表示単位],[仲介手数料表示単位],[駐車場敷金表示単位],[駐車場礼金表示単位],[入金項目_賃料1],[入金項目区分_賃料1],[入金項目_賃料2],[入金項目区分_賃料2],[入金項目_賃料3],[入金項目区分_賃料3],[入金項目_共益費1],[入金項目区分_共益費1],[入金項目_共益費2],[入金項目区分_共益費2],[入金項目_共益費3],[入金項目区分_共益費3],[入金項目_管理費1],[入金項目区分_管理費1],[入金項目_管理費2],[入金項目区分_管理費2],[入金項目_管理費3],[入金項目区分_管理費3],[入金項目_敷金],[入金項目区分_敷金],[入金項目_礼金],[入金項目区分_礼金],[入金項目_保証金],[入金項目区分_保証金],[入金項目_保証金償却],[入金項目区分_保証金償却],[入金項目_敷引金],[入金項目区分_敷引金],[入金項目_更新料],[入金項目区分_更新料],[入金項目_仲介手数料],[入金項目区分_仲介手数料],[入金項目_鍵交換代等],[入金項目区分_鍵交換代等],[入金項目_雑費1],[入金項目区分_雑費1],[入金項目_雑費2],[入金項目区分_雑費2],[入金項目_雑費3],[入金項目区分_雑費3],[入金項目_駐車場敷金],[入金項目区分_駐車場敷金],[入金項目_駐車場礼金],[入金項目区分_駐車場礼金],[入金項目_その他月額費用1],[入金項目区分_その他月額費用1],[入金項目_その他月額費用2],[入金項目区分_その他月額費用2],[入金項目_その他月額費用3],[入金項目区分_その他月額費用3],[入金項目_その他月額費用4],[入金項目区分_その他月額費用4],[入金項目_その他月額費用5],[入金項目区分_その他月額費用5],[入金項目_その他一時金1],[入金項目区分_その他一時金1],[入金項目_その他一時金2],[入金項目区分_その他一時金2],[入金項目_その他一時金3],[入金項目区分_その他一時金3],[入金項目_その他一時金4],[入金項目区分_その他一時金4],[入金項目_その他一時金5],[入金項目区分_その他一時金5],[画像No1],[革命側の画像種別1],[革命側の画像タイトルNo1],[連動側の画像種別1],[画像No2],[革命側の画像種別2],[革命側の画像タイトルNo2],[連動側の画像種別2],[画像No3],[革命側の画像種別3],[革命側の画像タイトルNo3],[連動側の画像種別3],[画像No4],[革命側の画像種別4],[革命側の画像タイトルNo4],[連動側の画像種別4],[画像No5],[革命側の画像種別5],[革命側の画像タイトルNo5],[連動側の画像種別5],[画像No6],[革命側の画像種別6],[革命側の画像タイトルNo6],[連動側の画像種別6],[画像No7],[革命側の画像種別7],[革命側の画像タイトルNo7],[連動側の画像種別7],[画像No8],[革命側の画像種別8],[革命側の画像タイトルNo8],[連動側の画像種別8],[画像No9],[革命側の画像種別9],[革命側の画像タイトルNo9],[連動側の画像種別9],[画像No10],[革命側の画像種別10],[革命側の画像タイトルNo10],[連動側の画像種別10],[画像No11],[革命側の画像種別11],[革命側の画像タイトルNo11],[連動側の画像種別11],[画像No12],[革命側の画像種別12],[革命側の画像タイトルNo12],[連動側の画像種別12],[画像No13],[革命側の画像種別13],[革命側の画像タイトルNo13],[連動側の画像種別13],[画像No14],[革命側の画像種別14],[革命側の画像タイトルNo14],[連動側の画像種別14],[画像No15],[革命側の画像種別15],[革命側の画像タイトルNo15],[連動側の画像種別15],[画像No16],[革命側の画像種別16],[革命側の画像タイトルNo16],[連動側の画像種別16]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // 20160829 連動中間ファイル作成処理修正　-add sta
                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 NOTGAZO.[送信設定順] ";
                tmp_sql = tmp_sql + " 		,NOTGAZO.[サイトNo] ";
                tmp_sql = tmp_sql + " 		,[ポータルサイトID] ";
                tmp_sql = tmp_sql + " 		,[サイト別送信有無] ";
                tmp_sql = tmp_sql + " 		,[ポータルサイトパスワード] ";
                tmp_sql = tmp_sql + " 		,[店舗ID] ";
                tmp_sql = tmp_sql + " 		,[物件名・部屋Noの表示(一般ユーザー向け)] ";
                tmp_sql = tmp_sql + " 		,[物件名・部屋Noの表示(不動産会社向け)] ";
                tmp_sql = tmp_sql + " 		,[地図表示] ";
                tmp_sql = tmp_sql + " 		,[番地以降の表示] ";
                tmp_sql = tmp_sql + " 		,[掲載確認日] ";
                tmp_sql = tmp_sql + " 		,[敷金表示単位] ";
                tmp_sql = tmp_sql + " 		,[礼金表示単位] ";
                tmp_sql = tmp_sql + " 		,[保証金表示単位] ";
                tmp_sql = tmp_sql + " 		,[保証金償却表示単位] ";
                tmp_sql = tmp_sql + " 		,[敷引金表示単位] ";
                tmp_sql = tmp_sql + " 		,[更新料表示単位] ";
                tmp_sql = tmp_sql + " 		,[仲介手数料表示単位] ";
                tmp_sql = tmp_sql + " 		,[駐車場敷金表示単位] ";
                tmp_sql = tmp_sql + " 		,[駐車場礼金表示単位] ";
                tmp_sql = tmp_sql + " 		,[入金項目_賃料1] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_賃料1] ";
                tmp_sql = tmp_sql + " 		,[入金項目_賃料2] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_賃料2] ";
                tmp_sql = tmp_sql + " 		,[入金項目_賃料3] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_賃料3] ";
                tmp_sql = tmp_sql + " 		,[入金項目_共益費1] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_共益費1] ";
                tmp_sql = tmp_sql + " 		,[入金項目_共益費2] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_共益費2] ";
                tmp_sql = tmp_sql + " 		,[入金項目_共益費3] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_共益費3] ";
                tmp_sql = tmp_sql + " 		,[入金項目_管理費1] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_管理費1] ";
                tmp_sql = tmp_sql + " 		,[入金項目_管理費2] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_管理費2] ";
                tmp_sql = tmp_sql + " 		,[入金項目_管理費3] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_管理費3] ";
                tmp_sql = tmp_sql + " 		,[入金項目_敷金] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_敷金] ";
                tmp_sql = tmp_sql + " 		,[入金項目_礼金] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_礼金] ";
                tmp_sql = tmp_sql + " 		,[入金項目_保証金] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_保証金] ";
                tmp_sql = tmp_sql + " 		,[入金項目_保証金償却] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_保証金償却] ";
                tmp_sql = tmp_sql + " 		,[入金項目_敷引金] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_敷引金] ";
                tmp_sql = tmp_sql + " 		,[入金項目_更新料] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_更新料] ";
                tmp_sql = tmp_sql + " 		,[入金項目_仲介手数料] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_仲介手数料] ";
                tmp_sql = tmp_sql + " 		,[入金項目_鍵交換代等] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_鍵交換代等] ";
                tmp_sql = tmp_sql + " 		,[入金項目_雑費1] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_雑費1] ";
                tmp_sql = tmp_sql + " 		,[入金項目_雑費2] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_雑費2] ";
                tmp_sql = tmp_sql + " 		,[入金項目_雑費3] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_雑費3] ";
                tmp_sql = tmp_sql + " 		,[入金項目_駐車場敷金] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_駐車場敷金] ";
                tmp_sql = tmp_sql + " 		,[入金項目_駐車場礼金] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_駐車場礼金] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他月額費用1] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他月額費用1] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他月額費用2] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他月額費用2] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他月額費用3] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他月額費用3] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他月額費用4] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他月額費用4] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他月額費用5] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他月額費用5] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他一時金1] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他一時金1] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他一時金2] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他一時金2] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他一時金3] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他一時金3] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他一時金4] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他一時金4] ";
                tmp_sql = tmp_sql + " 		,[入金項目_その他一時金5] ";
                tmp_sql = tmp_sql + " 		,[入金項目区分_その他一時金5] ";
                tmp_sql = tmp_sql + " 		,[画像No1] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別1] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo1] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別1] ";
                tmp_sql = tmp_sql + " 		,[画像No2] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別2] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo2] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別2] ";
                tmp_sql = tmp_sql + " 		,[画像No3] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別3] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo3] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別3] ";
                tmp_sql = tmp_sql + " 		,[画像No4] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別4] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo4] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別4] ";
                tmp_sql = tmp_sql + " 		,[画像No5] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別5] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo5] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別5] ";
                tmp_sql = tmp_sql + " 		,[画像No6] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別6] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo6] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別6] ";
                tmp_sql = tmp_sql + " 		,[画像No7] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別7] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo7] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別7] ";
                tmp_sql = tmp_sql + " 		,[画像No8] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別8] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo8] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別8] ";
                tmp_sql = tmp_sql + " 		,[画像No9] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別9] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo9] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別9] ";
                tmp_sql = tmp_sql + " 		,[画像No10] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別10] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo10] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別10] ";
                tmp_sql = tmp_sql + " 		,[画像No11] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別11] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo11] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別11] ";
                tmp_sql = tmp_sql + " 		,[画像No12] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別12] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo12] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別12] ";
                tmp_sql = tmp_sql + " 		,[画像No13] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別13] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo13] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別13] ";
                tmp_sql = tmp_sql + " 		,[画像No14] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別14] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo14] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別14] ";
                tmp_sql = tmp_sql + " 		,[画像No15] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別15] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo15] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別15] ";
                tmp_sql = tmp_sql + " 		,[画像No16] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像種別16] ";
                tmp_sql = tmp_sql + " 		,[革命側の画像タイトルNo16] ";
                tmp_sql = tmp_sql + " 		,[連動側の画像種別16] ";
                tmp_sql = tmp_sql + " 	FROM CVVW_送信設定athome情報_画像以外 AS NOTGAZO ";
                tmp_sql = tmp_sql + " 	LEFT JOIN CVVW_送信設定athome情報_画像 AS ONLYGAZO ";
                tmp_sql = tmp_sql + " 	ON NOTGAZO.[送信設定順] = ONLYGAZO.[送信設定順] ";
                tmp_sql = tmp_sql + " 	AND NOTGAZO.[サイトNo] = ONLYGAZO.[サイトNo] ";
                tmp_sql = tmp_sql + " ) AS VW ";
                // 20160829 連動中間ファイル作成処理修正　-add end

                // 20160908 連動_送信設定情報抽出クエリ修正 -add
                tmp_sql = " SELECT * FROM ( " + tmp_sql + " ) AS VW " + " WHERE [ポータルサイトID] <> '' ";

                // 20160829 連動中間ファイル作成処理修正　-del sta
                // For cntii = 0 To 9

                // tmp_sql = tmp_sql & " SELECT * FROM "
                // tmp_sql = tmp_sql & " ( "
                // tmp_sql = tmp_sql & " 	SELECT "
                // tmp_sql = tmp_sql & "       /*20160829 連動中間ファイル作成処理修正 chg sta*/ "
                // tmp_sql = tmp_sql & "       /* 送信設定No置換文字列 AS [送信設定順_V7]*/ "
                // tmp_sql = tmp_sql & "        送信設定No置換文字列 AS [送信設定順]  "
                // tmp_sql = tmp_sql & "       /*20160829 連動中間ファイル作成処理修正 chg end*/ "
                // tmp_sql = tmp_sql & " 		,30 AS [サイトNo] "
                // tmp_sql = tmp_sql & " 		/*******************************************フィールドへ移行するデータ sta*******************************************/ "
                // tmp_sql = tmp_sql & " 		,(SELECT REPLACE(set_value1,RIGHT(set_value1,4),'') FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061002') AS [ポータルサイトID] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01061001','11061001','21061001','31061001','41061001','51061001','61061001','71061001','81061001','91061001')) = 0 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01061001','11061001','21061001','31061001','41061001','51061001','61061001','71061001','81061001','91061001')) = 1 THEN 0 "
                // tmp_sql = tmp_sql & " 		 END AS [サイト別送信有無] "
                // tmp_sql = tmp_sql & " 		,'' AS [ポータルサイトパスワード] "
                // tmp_sql = tmp_sql & " 		/*******************************************フィールドへ移行するデータ end*******************************************/ "
                // tmp_sql = tmp_sql & " 		 "
                // tmp_sql = tmp_sql & " 		/************************************************xmlへ移行するデータ sta************************************************/ "
                // tmp_sql = tmp_sql & " 		 "
                // tmp_sql = tmp_sql & " 		/***********表示設定タブ***********/ "
                // tmp_sql = tmp_sql & " 		,(SELECT RIGHT(set_value1,4) FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061002') AS [店舗ID] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 3 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 1 THEN 3 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 2 THEN 3 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 3 THEN 3 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 1 THEN 4 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 2 THEN 4 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061003') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061017') = 3 THEN 4 "
                // tmp_sql = tmp_sql & " 		 END AS [物件名・部屋Noの表示(一般ユーザー向け)] "
                // tmp_sql = tmp_sql & "  		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 3 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 1 THEN 3 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 2 THEN 3 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 2 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 3 THEN 3 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 1 THEN 4 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 2 THEN 4 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061016') = 3 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061018') = 3 THEN 4 "
                // tmp_sql = tmp_sql & " 		 END AS [物件名・部屋Noの表示(不動産会社向け)]		 "
                // tmp_sql = tmp_sql & " 		 ,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1061011') AS [地図表示] "
                // tmp_sql = tmp_sql & " 		,2 AS [番地以降の表示]		/*要デフォルト値確認*/ "
                // tmp_sql = tmp_sql & " 		,1 AS [掲載確認日] "
                // tmp_sql = tmp_sql & " 		 "
                // tmp_sql = tmp_sql & " 		/***********入金項目タブ***********/ "
                // tmp_sql = tmp_sql & " 		/*,'_円' AS [賃料表示単位]*/		/*固定値のためコメントアウト*/ "
                // tmp_sql = tmp_sql & " 		/*,'_円' AS [共益費表示単位]*/		/*固定値のためコメントアウト*/ "
                // tmp_sql = tmp_sql & " 		/*,'_円' AS [管理費表示単位]*/		/*固定値のためコメントアウト*/ "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061101') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061101') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [敷金表示単位] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061102') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061102') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [礼金表示単位] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061103') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061103') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [保証金表示単位] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061107') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061107') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [保証金償却表示単位] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061106') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061106') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [敷引金表示単位] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061108') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061108') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [更新料表示単位] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061109') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061109') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [仲介手数料表示単位] "
                // tmp_sql = tmp_sql & " 		/*,'_円' AS [鍵交換代等]*/		/*固定値のためコメントアウト*/ "
                // tmp_sql = tmp_sql & " 		/*,'_円' AS [雑費]*/			/*固定値のためコメントアウト*/ "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061104') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061104') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [駐車場敷金表示単位] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061105') = 1 THEN '_円' "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061105') = 2 THEN '_ヶ月' "
                // tmp_sql = tmp_sql & " 		    ELSE '_円' "
                // tmp_sql = tmp_sql & " 		 END AS [駐車場礼金表示単位]		 		 "
                // tmp_sql = tmp_sql & " 		/*,'_円' AS [その他一時金]*/	/*固定値のためコメントアウト*/ "
                // tmp_sql = tmp_sql & " 		/*,'_円' AS [その他月額費用]*/	/*固定値のためコメントアウト*/ "
                // tmp_sql = tmp_sql & " 		 "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061001')) AS [入金項目_賃料1] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_賃料1] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061002')) AS [入金項目_賃料2] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_賃料2] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061003')) AS [入金項目_賃料3] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_賃料3] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061004')) AS [入金項目_共益費1] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_共益費1] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061005')) AS [入金項目_共益費2] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_共益費2] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061006')) AS [入金項目_共益費3] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_共益費3] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061007')) AS [入金項目_管理費1] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_管理費1] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061008')) AS [入金項目_管理費2] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_管理費2] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061009')) AS [入金項目_管理費3] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_管理費3] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061013')) AS [入金項目_敷金] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_敷金] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061014')) AS [入金項目_礼金] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_礼金] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061015')) AS [入金項目_保証金] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_保証金] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061021')) AS [入金項目_保証金償却] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_保証金償却] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061020')) AS [入金項目_敷引金] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_敷引金] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061030')) AS [入金項目_更新料] "
                // tmp_sql = tmp_sql & " 		,3 AS [入金項目区分_更新料] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061034')) AS [入金項目_仲介手数料] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_仲介手数料] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061033')) AS [入金項目_鍵交換代等] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_鍵交換代等]		 "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061010')) AS [入金項目_雑費1] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_雑費1] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061011')) AS [入金項目_雑費2] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_雑費2] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061012')) AS [入金項目_雑費3] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_雑費3]		 "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061018')) AS [入金項目_駐車場敷金] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_駐車場敷金] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061019')) AS [入金項目_駐車場礼金] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_駐車場礼金] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061022')) AS [入金項目_その他月額費用1] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用1] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061023')) AS [入金項目_その他月額費用2] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用2] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061024')) AS [入金項目_その他月額費用3] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用3] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061031')) AS [入金項目_その他月額費用4] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用4] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061032')) AS [入金項目_その他月額費用5] "
                // tmp_sql = tmp_sql & " 		,1 AS [入金項目区分_その他月額費用5] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061025')) AS [入金項目_その他一時金1] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金1] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061026')) AS [入金項目_その他一時金2] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金2] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061027')) AS [入金項目_その他一時金3] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金3] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061028')) AS [入金項目_その他一時金4] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金4] "
                // tmp_sql = tmp_sql & " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5061029')) AS [入金項目_その他一時金5] "
                // tmp_sql = tmp_sql & " 		,2 AS [入金項目区分_その他一時金5]		 "
                // tmp_sql = tmp_sql & "  "
                // tmp_sql = tmp_sql & " 		 /***********画像設定タブ***********/ "
                // tmp_sql = tmp_sql & " 		,1 AS [画像No1] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061001'),1) AS [革命側の画像種別1] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061001'),2)) AS [革命側の画像タイトルNo1] "
                // tmp_sql = tmp_sql & " 		,1 AS [連動側の画像種別1]	/*固定*/ "
                // tmp_sql = tmp_sql & " 		,2 AS [画像No2] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061002'),1) AS [革命側の画像種別2] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061002'),2)) AS [革命側の画像タイトルNo2] "
                // tmp_sql = tmp_sql & " 		,3 AS [連動側の画像種別2]	/*固定*/ "
                // tmp_sql = tmp_sql & " 		,3 AS [画像No3] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AS [革命側の画像種別3] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) AS [革命側の画像タイトルNo3] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061003'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別3] "
                // tmp_sql = tmp_sql & " 		,4 AS [画像No4] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AS [革命側の画像種別4] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) AS [革命側の画像タイトルNo4] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061004'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別4] "
                // tmp_sql = tmp_sql & " 		,5 AS [画像No5] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AS [革命側の画像種別5] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) AS [革命側の画像タイトルNo5] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061005'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別5] "
                // tmp_sql = tmp_sql & " 		,6 AS [画像No6] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AS [革命側の画像種別6] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) AS [革命側の画像タイトルNo6] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061006'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別6]	 "
                // tmp_sql = tmp_sql & " 		,7 AS [画像No7] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AS [革命側の画像種別7] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) AS [革命側の画像タイトルNo7] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061007'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別7]		 "
                // tmp_sql = tmp_sql & " 		,8 AS [画像No8] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AS [革命側の画像種別8] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) AS [革命側の画像タイトルNo8] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061008'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別8] "
                // tmp_sql = tmp_sql & " 		,9 AS [画像No9] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AS [革命側の画像種別9] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) AS [革命側の画像タイトルNo9] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061009'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別9] "
                // tmp_sql = tmp_sql & " 		,10 AS [画像No10] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AS [革命側の画像種別10] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) AS [革命側の画像タイトルNo10] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061010'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別10] "
                // tmp_sql = tmp_sql & " 		,11 AS [画像No11] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AS [革命側の画像種別11] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) AS [革命側の画像タイトルNo11] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061011'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別11] "
                // tmp_sql = tmp_sql & " 		,12 AS [画像No12] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AS [革命側の画像種別12] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) AS [革命側の画像タイトルNo12] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061012'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別12] "
                // tmp_sql = tmp_sql & " 		,13 AS [画像No13] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AS [革命側の画像種別13] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) AS [革命側の画像タイトルNo13] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061013'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別13] "
                // tmp_sql = tmp_sql & " 		,14 AS [画像No14] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AS [革命側の画像種別14] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) AS [革命側の画像タイトルNo14] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061014'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別14] "
                // tmp_sql = tmp_sql & " 		,15 AS [画像No15] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AS [革命側の画像種別15] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) AS [革命側の画像タイトルNo15] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061015'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別15] "
                // tmp_sql = tmp_sql & " 		,16 AS [画像No16] "
                // tmp_sql = tmp_sql & " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AS [革命側の画像種別16] "
                // tmp_sql = tmp_sql & " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) AS [革命側の画像タイトルNo16] "
                // tmp_sql = tmp_sql & " 		,CASE "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 0 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 3 THEN 14 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 4 THEN 20 "
                // tmp_sql = tmp_sql & " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3061016'),2)) = 5 THEN 18 "
                // tmp_sql = tmp_sql & "     		ELSE 0 "
                // tmp_sql = tmp_sql & " 		 END AS [連動側の画像種別16] "
                // tmp_sql = tmp_sql & " 		 "
                // tmp_sql = tmp_sql & " 		 /***********連動項目設定タブ***********/ "
                // tmp_sql = tmp_sql & " 		  "
                // tmp_sql = tmp_sql & " 		/************************************************xmlへ移行するデータ end************************************************/ "
                // tmp_sql = tmp_sql & " ) AS VW "

                // tmp_sql = tmp_sql.Replace("送信設定No置換文字列", (cntii + 1).ToString)
                // tmp_sql = tmp_sql.Replace("V7送信設定値取得用置換文字列", cntii.ToString)

                // If cntii < 9 Then
                // tmp_sql = tmp_sql & " UNION "
                // End If

                // Next

                // Dim tmp_sql_base As String = ""
                // tmp_sql_base = tmp_sql_base & " SELECT [No] FROM CVTBL_物件データ連動情報 AS VW "
                // tmp_sql_base = tmp_sql_base & " WHERE [Enabled] <> 0 "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1081002') AS SETTEI		/*自社WEB必須項目(ログインID)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001021') AS SETTEI		/*HOMES必須項目(会員番号)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1061002') AS SETTEI		/*athome必須項目(店舗ID)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001061') AS SETTEI		/*SUUMO必須項目(支店コード)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "

                // 'tmp_sql = " SELECT * FROM ( " & tmp_sql & ") AS VW WHERE [送信設定順_V7] IN (SELECT [No] FROM CVTBL_物件データ連動情報 WHERE [Enabled] <> 0) "
                // tmp_sql = " SELECT * FROM ( " & tmp_sql & ") AS VW WHERE [送信設定順_V7] IN (" & tmp_sql_base & ") "
                // tmp_sql = " SELECT ROW_NUMBER()OVER(ORDER BY [送信設定順_V7]) AS [送信設定順],* FROM ( " & tmp_sql & ") AS VW "
                // tmp_sql = " SELECT " & tmp_chgmidheader & " FROM (" & tmp_sql & ") AS VW "
                // 20160829 連動中間ファイル作成処理修正　-del end
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

    #region 送信設定SUMMO

    public class Rendo_sosin_suumo_Repository
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

                sortstr = "[送信設定順],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[送信設定順],[サイトNo],[ポータルサイトID],[サイト別送信有無],[ポータルサイトパスワード],[FTPユーザーID],[FTPパスワード],[物件名・部屋Noの表示(一般ユーザー向け)],[物件名・部屋Noの表示(会社間図面および会社間流通向け)],[番地以降の表示(会社間図面および会社間流通向け)],[地図表示],[元付確認日(先物物件のみ)],[空室状況設定],[敷金表示単位],[礼金表示単位],[保証金表示単位],[償却金表示単位],[敷引表示単位],[仲介手数料表示単位],[入金項目_賃料1],[入金項目区分_賃料1],[入金項目_賃料2],[入金項目区分_賃料2],[入金項目_賃料3],[入金項目区分_賃料3],[入金項目_管理費1],[入金項目区分_管理費1],[入金項目_管理費2],[入金項目区分_管理費2],[入金項目_管理費3],[入金項目区分_管理費3],[入金項目_敷金],[入金項目区分_敷金],[入金項目_礼金],[入金項目区分_礼金],[入金項目_保証金],[入金項目区分_保証金],[入金項目_償却金],[入金項目区分_償却金],[入金項目_敷引],[入金項目区分_敷引],[入金項目_仲介手数料],[入金項目区分_仲介手数料],[入金項目_その他諸費用],[入金項目区分_その他諸費用],[入金項目_ほか初期費用],[入金項目区分_ほか初期費用],[画像No1],[革命側の画像種別1],[革命側の画像タイトルNo1],[連動側の画像種別1],[画像No2],[革命側の画像種別2],[革命側の画像タイトルNo2],[連動側の画像種別2],[画像No3],[革命側の画像種別3],[革命側の画像タイトルNo3],[連動側の画像種別3],[画像No4],[革命側の画像種別4],[革命側の画像タイトルNo4],[連動側の画像種別4],[画像No5],[革命側の画像種別5],[革命側の画像タイトルNo5],[連動側の画像種別5],[画像No6],[革命側の画像種別6],[革命側の画像タイトルNo6],[連動側の画像種別6],[画像No7],[革命側の画像種別7],[革命側の画像タイトルNo7],[連動側の画像種別7],[画像No8],[革命側の画像種別8],[革命側の画像タイトルNo8],[連動側の画像種別8],[画像No9],[革命側の画像種別9],[革命側の画像タイトルNo9],[連動側の画像種別9],[画像No10],[革命側の画像種別10],[革命側の画像タイトルNo10],[連動側の画像種別10],[画像No11],[革命側の画像種別11],[革命側の画像タイトルNo11],[連動側の画像種別11],[画像No12],[革命側の画像種別12],[革命側の画像タイトルNo12],[連動側の画像種別12],[画像No13],[革命側の画像種別13],[革命側の画像タイトルNo13],[連動側の画像種別13],[画像No14],[革命側の画像種別14],[革命側の画像タイトルNo14],[連動側の画像種別14],[画像No15],[革命側の画像種別15],[革命側の画像タイトルNo15],[連動側の画像種別15],[画像No16],[革命側の画像種別16],[革命側の画像タイトルNo16],[連動側の画像種別16],[画像No17],[革命側の画像種別17],[革命側の画像タイトルNo17],[連動側の画像種別17]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                for (int cntii = 0; cntii <= 9; cntii++)
                {

                    tmp_sql = tmp_sql + " SELECT * FROM ";
                    tmp_sql = tmp_sql + " ( ";
                    tmp_sql = tmp_sql + " 	SELECT ";
                    tmp_sql = tmp_sql + "       /*20160829 連動中間ファイル作成処理修正 chg sta*/ ";
                    tmp_sql = tmp_sql + "       /* 送信設定No置換文字列 AS [送信設定順_V7]*/ ";
                    tmp_sql = tmp_sql + "        送信設定No置換文字列 AS [送信設定順]  ";
                    tmp_sql = tmp_sql + "       /*20160829 連動中間ファイル作成処理修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 		,40 AS [サイトNo] ";
                    tmp_sql = tmp_sql + " 		/*******************************************フィールドへ移行するデータ sta*******************************************/ ";
                    tmp_sql = tmp_sql + " 		/*,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001061') AS [ポータルサイトID]*/ ";
                    tmp_sql = tmp_sql + "       ,(SELECT CONVERT(VARCHAR,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001060')) +  ";
                    tmp_sql = tmp_sql + " 	             CONVERT(VARCHAR,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001061'))) AS [ポータルサイトID] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01001059','11001059','21001059','31001059','41001059','51001059','61001059','71001059','81001059','91001059')) = 0 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT MIN(set_value1) AS [sosin_flg] FROM m_sosin_settei WHERE kbn_code IN ('01001059','11001059','21001059','31001059','41001059','51001059','61001059','71001059','81001059','91001059')) = 1 THEN 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [サイト別送信有無] ";
                    tmp_sql = tmp_sql + " 		,'' AS [ポータルサイトパスワード] ";
                    tmp_sql = tmp_sql + " 		/*******************************************フィールドへ移行するデータ end*******************************************/ ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		/************************************************xmlへ移行するデータ sta************************************************/ ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 		/***********表示設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001075') AS [FTPユーザーID] ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001076') AS [FTPパスワード] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001062') = 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001063') = 0 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001062') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001063') = 0 THEN 3 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001062') = 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001063') = 1 THEN 4 ";
                    tmp_sql = tmp_sql + " 		 END AS [物件名・部屋Noの表示(一般ユーザー向け)] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001062') = 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001063') = 0 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001062') = 1 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001063') = 0 THEN 3 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001062') = 0 AND (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001063') = 1 THEN 4 ";
                    tmp_sql = tmp_sql + " 		 END AS [物件名・部屋Noの表示(会社間図面および会社間流通向け)]		  ";
                    tmp_sql = tmp_sql + " 		,2 AS [番地以降の表示(会社間図面および会社間流通向け)]		/*要デフォルト値確認*/ ";
                    tmp_sql = tmp_sql + " 		,(SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001081') AS [地図表示] ";
                    tmp_sql = tmp_sql + " 		/********************************************** ";
                    tmp_sql = tmp_sql + " 		V7元付確認日を自動で更新するのチェック ";
                    tmp_sql = tmp_sql + " 		  OFF … 10の元付確認日→広告内容確認時点の日付		 ";
                    tmp_sql = tmp_sql + " 		  ON  … 10の元付確認日→送信時点の日付 ";
                    tmp_sql = tmp_sql + " 		**********************************************/ ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001069') = 0 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001069') = 1 THEN 2 ";
                    tmp_sql = tmp_sql + " 		 END AS [元付確認日(先物物件のみ)] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001071') = 0 THEN 'false' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列1001071') = 1 THEN 'true' ";
                    tmp_sql = tmp_sql + " 		 END AS [空室状況設定] ";
                    tmp_sql = tmp_sql + "  ";
                    tmp_sql = tmp_sql + " 		/***********入金項目タブ***********/ ";
                    tmp_sql = tmp_sql + " 		/*,'_万円' AS [賃料表示単位]*/				/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		/*,'_万円' AS [管理費表示単位]*/				/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009051') = 1 THEN '_万円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009051') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_万円' ";
                    tmp_sql = tmp_sql + " 		 END AS [敷金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009050') = 1 THEN '_万円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009050') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_万円' ";
                    tmp_sql = tmp_sql + " 		 END AS [礼金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009053') = 1 THEN '_万円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009053') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_万円' ";
                    tmp_sql = tmp_sql + " 		 END AS [保証金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009052') = 1 THEN '_万円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009052') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_万円' ";
                    tmp_sql = tmp_sql + " 		 END AS [償却金表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009056') = 1 THEN '_万円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009056') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_万円' ";
                    tmp_sql = tmp_sql + " 		 END AS [敷引表示単位] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009054') = 1 THEN '_万円' ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5009054') = 2 THEN '_ヶ月' ";
                    tmp_sql = tmp_sql + " 		    ELSE '_万円' ";
                    tmp_sql = tmp_sql + " 		 END AS [仲介手数料表示単位] ";
                    tmp_sql = tmp_sql + " 		/*,'_万円' AS [その他諸費用]*/		/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		/*,'_万円' AS [ほか初期費用]*/		/*固定値のためコメントアウト*/ ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001101')) AS [入金項目_賃料1] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001102')) AS [入金項目_賃料2] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001103')) AS [入金項目_賃料3] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_賃料3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001104')) AS [入金項目_管理費1] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_管理費1] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001105')) AS [入金項目_管理費2] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_管理費2] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001106')) AS [入金項目_管理費3] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_管理費3] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001107')) AS [入金項目_敷金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_敷金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001108')) AS [入金項目_礼金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_礼金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001109')) AS [入金項目_保証金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_保証金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001116')) AS [入金項目_償却金] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_償却金] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001110')) AS [入金項目_敷引] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_敷引] ";
                    tmp_sql = tmp_sql + " 		,(SELECT nkin_name FROM m_nkin WHERE nkin_no = (SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列5001117')) AS [入金項目_仲介手数料] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_仲介手数料] ";
                    tmp_sql = tmp_sql + " 		,NULL AS [入金項目_その他諸費用] ";
                    tmp_sql = tmp_sql + " 		,1 AS [入金項目区分_その他諸費用] ";
                    tmp_sql = tmp_sql + " 		,NULL AS [入金項目_ほか初期費用] ";
                    tmp_sql = tmp_sql + " 		,2 AS [入金項目区分_ほか初期費用] ";
                    tmp_sql = tmp_sql + " 		 ";
                    tmp_sql = tmp_sql + " 				 ";
                    tmp_sql = tmp_sql + " 		 /***********画像設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		,1 AS [画像No1] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009501'),1) AS [革命側の画像種別1] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009501'),2)) AS [革命側の画像タイトルNo1] ";
                    tmp_sql = tmp_sql + " 		,-1 AS [連動側の画像種別1] ";
                    tmp_sql = tmp_sql + " 		,2 AS [画像No2] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009503'),1) AS [革命側の画像種別2] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009503'),2)) AS [革命側の画像タイトルNo2] ";
                    tmp_sql = tmp_sql + " 		,-1 AS [連動側の画像種別2] ";
                    tmp_sql = tmp_sql + " 		,3 AS [画像No3] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009502'),1) AS [革命側の画像種別3] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009502'),2)) AS [革命側の画像タイトルNo3] ";
                    tmp_sql = tmp_sql + " 		,-1 AS [連動側の画像種別3] ";
                    tmp_sql = tmp_sql + " 		,4 AS [画像No4] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009504'),1) AS [革命側の画像種別4] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009504'),2)) AS [革命側の画像タイトルNo4] ";
                    tmp_sql = tmp_sql + " 		,-1 AS [連動側の画像種別4] ";
                    tmp_sql = tmp_sql + " 		,5 AS [画像No5] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009505'),1) AS [革命側の画像種別5] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009505'),2)) AS [革命側の画像タイトルNo5] ";
                    tmp_sql = tmp_sql + " 		,-1 AS [連動側の画像種別5] ";
                    tmp_sql = tmp_sql + " 		,6 AS [画像No6] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),1) AS [革命側の画像種別6] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),2)) AS [革命側の画像タイトルNo6] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009506'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別6] ";
                    tmp_sql = tmp_sql + " 		,7 AS [画像No7] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),1) AS [革命側の画像種別7] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),2)) AS [革命側の画像タイトルNo7] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009507'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別7] ";
                    tmp_sql = tmp_sql + " 		,8 AS [画像No8] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),1) AS [革命側の画像種別8] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),2)) AS [革命側の画像タイトルNo8] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009508'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別8] ";
                    tmp_sql = tmp_sql + " 		,9 AS [画像No9] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),1) AS [革命側の画像種別9] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),2)) AS [革命側の画像タイトルNo9] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009509'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別9] ";
                    tmp_sql = tmp_sql + " 		,10 AS [画像No10] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),1) AS [革命側の画像種別10] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),2)) AS [革命側の画像タイトルNo10] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009510'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別10] ";
                    tmp_sql = tmp_sql + " 		,11 AS [画像No11] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),1) AS [革命側の画像種別11] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),2)) AS [革命側の画像タイトルNo11] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009511'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別11] ";
                    tmp_sql = tmp_sql + " 		,12 AS [画像No12] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),1) AS [革命側の画像種別12] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),2)) AS [革命側の画像タイトルNo12] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009512'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別12] ";
                    tmp_sql = tmp_sql + " 		,13 AS [画像No13] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),1) AS [革命側の画像種別13] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),2)) AS [革命側の画像タイトルNo13] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009513'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別13] ";
                    tmp_sql = tmp_sql + " 		,14 AS [画像No14] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),1) AS [革命側の画像種別14] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),2)) AS [革命側の画像タイトルNo14] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009514'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別14] ";
                    tmp_sql = tmp_sql + " 		,15 AS [画像No15] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),1) AS [革命側の画像種別15] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),2)) AS [革命側の画像タイトルNo15] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009515'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別15] ";
                    tmp_sql = tmp_sql + " 		,16 AS [画像No16] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),1) AS [革命側の画像種別16] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),2)) AS [革命側の画像タイトルNo16] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009516'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別16] ";
                    tmp_sql = tmp_sql + " 		,17 AS [画像No17] ";
                    tmp_sql = tmp_sql + " 		,LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),1) AS [革命側の画像種別17] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),2)) AS [革命側の画像タイトルNo17] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),2)) = 0 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),2)) = 1 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),2)) = 2 THEN 20101 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),2)) = 3 THEN 30199 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),2)) = 4 THEN 999999 ";
                    tmp_sql = tmp_sql + " 			WHEN (SELECT gazo_category_brui FROM m_gazo_syoki WHERE gazo_kbn = LEFT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),1) AND gazo_no = RIGHT((SELECT set_value1 FROM m_sosin_settei WHERE kbn_code = 'V7送信設定値取得用置換文字列3009517'),2)) = 5 THEN 999999 ";
                    tmp_sql = tmp_sql + "     		ELSE 0 ";
                    tmp_sql = tmp_sql + " 		 END AS [連動側の画像種別17] ";
                    tmp_sql = tmp_sql + "  ";
                    tmp_sql = tmp_sql + " 		 /***********連動項目設定タブ***********/ ";
                    tmp_sql = tmp_sql + " 		  ";
                    tmp_sql = tmp_sql + " 		/************************************************xmlへ移行するデータ end************************************************/ ";
                    tmp_sql = tmp_sql + " ) AS VW ";

                    tmp_sql = tmp_sql.Replace("送信設定No置換文字列", (cntii + 1).ToString());
                    tmp_sql = tmp_sql.Replace("V7送信設定値取得用置換文字列", cntii.ToString());

                    if (cntii < 9)
                    {
                        tmp_sql = tmp_sql + " UNION ";
                    }

                }

                // 20160908 連動_送信設定情報抽出クエリ修正 -add
                tmp_sql = " SELECT * FROM ( " + tmp_sql + " ) AS VW " + " WHERE [ポータルサイトID] <> '' ";

                // 20160829 連動中間ファイル作成処理修正 -del sta
                // Dim tmp_sql_base As String = ""
                // tmp_sql_base = tmp_sql_base & " SELECT [No] FROM CVTBL_物件データ連動情報 AS VW "
                // tmp_sql_base = tmp_sql_base & " WHERE [Enabled] <> 0 "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1081002') AS SETTEI		/*自社WEB必須項目(ログインID)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001021') AS SETTEI		/*HOMES必須項目(会員番号)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1061002') AS SETTEI		/*athome必須項目(店舗ID)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "
                // tmp_sql_base = tmp_sql_base & " AND EXISTS "
                // tmp_sql_base = tmp_sql_base & " 		( "
                // tmp_sql_base = tmp_sql_base & " 			SELECT * FROM (SELECT * FROM m_sosin_settei WHERE RIGHT(kbn_code,7) = '1001061') AS SETTEI		/*SUUMO必須項目(支店コード)*/ "
                // tmp_sql_base = tmp_sql_base & " 			WHERE VW.[No]-1 = LEFT(SETTEI.kbn_code,1) "
                // tmp_sql_base = tmp_sql_base & " 			AND   ISNULL(SETTEI.set_value1,'') <> '' "
                // tmp_sql_base = tmp_sql_base & " 		) "

                // 'tmp_sql = " SELECT * FROM ( " & tmp_sql & ") AS VW WHERE [送信設定順_V7] IN (SELECT [No] FROM CVTBL_物件データ連動情報 WHERE [Enabled] <> 0) "
                // tmp_sql = " SELECT * FROM ( " & tmp_sql & ") AS VW WHERE [送信設定順_V7] IN (" & tmp_sql_base & ") "
                // tmp_sql = " SELECT ROW_NUMBER()OVER(ORDER BY [送信設定順_V7]) AS [送信設定順],* FROM ( " & tmp_sql & ") AS VW "
                // tmp_sql = " SELECT " & tmp_chgmidheader & " FROM (" & tmp_sql & ") AS VW "
                // 20160829 連動中間ファイル作成処理修正 -del end
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

    #region 広告補足自社web

    public class Rendo_kokoku_jisyaweb_Repository
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

                sortstr = "[物件No],[部屋No],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[サイトNo],[セールスポイント補足],[担当者コメント],[特記事項],[広告備考],[天井高],[OAフロアの高さ],[賃料応相談]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,10 AS [サイトNo] ";
                tmp_sql = tmp_sql + " 		/*物件の特徴*/ ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN d2_data_id = 123 THEN d2_item ELSE '' END) AS [セールスポイント補足] ";
                tmp_sql = tmp_sql + " 		,NULL AS [担当者コメント] ";
                tmp_sql = tmp_sql + " 		/*広告備考*/ ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN d2_data_id = 128 THEN d2_item ELSE '' END) AS [特記事項] ";
                tmp_sql = tmp_sql + " 		,NULL AS [広告備考] ";
                tmp_sql = tmp_sql + " 		/*テナント補足*/ ";
                tmp_sql = tmp_sql + " 		,NULL AS [天井高] ";
                tmp_sql = tmp_sql + " 		,NULL AS [OAフロアの高さ] ";
                tmp_sql = tmp_sql + " 		,'false' AS [賃料応相談] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM hy_d2 WHERE d2_data_kbn = 70 ";
                tmp_sql = tmp_sql + " 	) AS VW ";
                tmp_sql = tmp_sql + " 	GROUP BY bk_no,hy_no ";
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

    #region 広告補足HOMES

    public class Rendo_kokoku_homes_Repository
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

                sortstr = "[物件No],[部屋No],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[サイトNo],[物件の特徴_自社HPに表示],[物件の特徴_他社HPに表示],[広告備考],[広告備考_自社HPに表示],[広告備考_他社HPに表示],[リンク先区分],[リンク先URL],[社内用メモ],[リフォーム実施年月],[リフォームその他箇所],[リフォーム備考],[リノベーション施工完了年月],[リノベーション内容],[カスタマイズ可否],[カスタマイズ内容],[カスタマイズ条件],[特定優遇賃貸住宅チェック],[特優賃下限],[特優賃上限],[特優賃料金変動区分],[特優賃料金上昇率],[特優賃家賃補助年数],[特優賃備考],[鍵保管場所],[預託先会社名],[鍵備考],[限定公開],[画像ダウンロード不可],[営業スタッフ名],[営業スタッフコメント設定方法],[コメントテンプレート],[コメント種別],[コメント],[マンスリー可],[駐車場契約必須],[他社取込],[媒介契約年月日],[特別広告ポイント数],[空き家バンク登録物件]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 [物件No] ";
                tmp_sql = tmp_sql + " 		,[部屋No] ";
                tmp_sql = tmp_sql + " 		,20 AS [サイトNo] ";
                tmp_sql = tmp_sql + " 		,[物件の特徴_自社HPに表示] ";
                tmp_sql = tmp_sql + " 		,[物件の特徴_他社HPに表示] ";
                tmp_sql = tmp_sql + " 		,[広告備考] ";
                tmp_sql = tmp_sql + " 		,[広告備考_自社HPに表示] ";
                tmp_sql = tmp_sql + " 		,[広告備考_他社HPに表示] ";
                tmp_sql = tmp_sql + " 		,[リンク先区分] ";
                tmp_sql = tmp_sql + " 		,[リンク先URL] ";
                tmp_sql = tmp_sql + " 		,[社内用メモ] ";
                tmp_sql = tmp_sql + " 		,[リフォーム実施年月] ";
                tmp_sql = tmp_sql + " 		/*,[リフォーム箇所]*/ ";
                tmp_sql = tmp_sql + " 		/*,[リフォームその他箇所]*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [リフォーム箇所] = '' AND [リフォームその他箇所] = '' THEN ''							/*両方データ無し*/ ";
                tmp_sql = tmp_sql + " 			WHEN [リフォーム箇所] = '' AND [リフォームその他箇所] <> '' THEN [リフォームその他箇所]		/*[リフォームその他箇所]のみ存在する場合*/ ";
                tmp_sql = tmp_sql + " 			WHEN [リフォーム箇所] <> '' AND [リフォームその他箇所] = '' THEN [リフォーム箇所]			/*[リフォーム箇所]のみ存在する場合*/ ";
                tmp_sql = tmp_sql + " 			WHEN ([リフォーム箇所] <> '' AND [リフォームその他箇所] <> '') AND							/*両方存在し、結合時に半角50文字以下の場合*/ ";
                tmp_sql = tmp_sql + " 					DATALENGTH([リフォーム箇所] + '[' + [リフォームその他箇所] + ']' ) <= 50 THEN [リフォーム箇所] + '[' + [リフォームその他箇所] + ']' ";
                tmp_sql = tmp_sql + " 			WHEN ([リフォーム箇所] <> '' AND [リフォームその他箇所] <> '') AND							/*両方存在し、結合時に半角50文字を超える場合*/ ";
                tmp_sql = tmp_sql + " 					DATALENGTH([リフォーム箇所] + '[' + [リフォームその他箇所] + ']' ) > 50 THEN [リフォームその他箇所] ";
                tmp_sql = tmp_sql + " 		 END AS [リフォームその他箇所] ";
                tmp_sql = tmp_sql + " 		,[リフォーム備考] ";
                tmp_sql = tmp_sql + " 		,[リノベーション施工完了年月]	/*V7:実施日 10:完了日*/ ";
                tmp_sql = tmp_sql + " 		,[リノベーション内容] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [カスタマイズ可否] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [カスタマイズ可否] ";
                tmp_sql = tmp_sql + " 		 END AS [カスタマイズ可否] ";
                tmp_sql = tmp_sql + " 		,[カスタマイズ内容] ";
                tmp_sql = tmp_sql + " 		,[カスタマイズ条件] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ";
                tmp_sql = tmp_sql + " 				[特優賃下限] = '' AND [特優賃上限] = '' AND  ";
                tmp_sql = tmp_sql + " 				[特優賃料金変動区分] = '' AND [特優賃料金上昇率] = '' AND  ";
                tmp_sql = tmp_sql + " 				[特優賃家賃補助年数] = '' AND [特優賃備考] = '' ";
                tmp_sql = tmp_sql + " 			THEN 'false' ";
                tmp_sql = tmp_sql + " 			ELSE 'true' ";
                tmp_sql = tmp_sql + " 		 END AS [特定優遇賃貸住宅チェック] ";
                tmp_sql = tmp_sql + " 		,[特優賃下限] ";
                tmp_sql = tmp_sql + " 		,[特優賃上限] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [特優賃料金変動区分] = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN [特優賃料金変動区分] = 2 THEN 3 ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [特優賃料金変動区分] ";
                tmp_sql = tmp_sql + " 		,[特優賃料金上昇率] ";
                tmp_sql = tmp_sql + " 		,[特優賃家賃補助年数] ";
                tmp_sql = tmp_sql + " 		,[特優賃備考] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [鍵保管場所] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [鍵保管場所] ";
                tmp_sql = tmp_sql + " 		 END AS [鍵保管場所] ";
                tmp_sql = tmp_sql + " 		,[預託先会社名] ";
                tmp_sql = tmp_sql + " 		,[鍵備考] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [限定公開] = 1 THEN 'false' ";
                tmp_sql = tmp_sql + " 			WHEN [限定公開] = 2 THEN 'true' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [限定公開] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [画像ダウンロード不可] = 1 THEN 'false' ";
                tmp_sql = tmp_sql + " 			WHEN [画像ダウンロード不可] = 2 THEN 'true' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [画像ダウンロード不可]	 ";
                tmp_sql = tmp_sql + " 		,[営業スタッフ名] ";
                tmp_sql = tmp_sql + " 		,[営業スタッフコメント設定方法]	 ";
                tmp_sql = tmp_sql + " 		,[コメントテンプレート]	 ";
                tmp_sql = tmp_sql + " 		,[コメント種別] ";
                tmp_sql = tmp_sql + " 		,[コメント] ";
                tmp_sql = tmp_sql + " 		,[マンスリー可] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [駐車場契約必須] = '' THEN 'false' ";
                tmp_sql = tmp_sql + " 			WHEN [駐車場契約必須] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 		 END AS [駐車場契約必須] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [他社取込] = '' THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE [他社取込] ";
                tmp_sql = tmp_sql + " 		 END AS [他社取込] ";
                tmp_sql = tmp_sql + " 		,[媒介契約年月日] ";
                tmp_sql = tmp_sql + " 		,[特別広告ポイント数] ";
                tmp_sql = tmp_sql + " 		,[空き家バンク登録物件] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 			,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 			/*物件の特徴*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 132 THEN d2_item ELSE '' END) AS [物件の特徴_自社HPに表示] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 133 THEN d2_item ELSE '' END) AS [物件の特徴_他社HPに表示] ";
                tmp_sql = tmp_sql + " 			/*広告備考*/ ";
                tmp_sql = tmp_sql + " 			,'' AS [広告備考] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 135 THEN d2_item ELSE '' END) AS [広告備考_自社HPに表示] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 136 THEN d2_item ELSE '' END) AS [広告備考_他社HPに表示] ";
                tmp_sql = tmp_sql + " 			,0 AS [リンク先区分]	/*デフォルトで物件ページを設定*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 137 THEN d2_item ELSE '' END) AS [リンク先URL] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 138 THEN d2_item ELSE '' END) AS [社内用メモ] ";
                tmp_sql = tmp_sql + " 			/*リフォーム*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 267 THEN d2_item ELSE '' END) AS [リフォーム実施年月] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 268 THEN d2_item ELSE '' END) AS [リフォーム箇所]		/*要紐付？*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 269 THEN d2_item ELSE '' END) AS [リフォームその他箇所] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 270 THEN d2_item ELSE '' END) AS [リフォーム備考] ";
                tmp_sql = tmp_sql + " 			/*リノベーション*/ ";
                tmp_sql = tmp_sql + " 			/*,MAX(CASE WHEN d2_data_id = 271 THEN d2_item ELSE '' END) AS [リノベーション施工完了年月]*/ ";
                tmp_sql = tmp_sql + " 			,'' AS [リノベーション施工完了年月] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 272 THEN d2_item ELSE '' END) AS [リノベーション内容] ";
                tmp_sql = tmp_sql + " 			/*内装カスタマイズ*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 383 THEN d2_item ELSE '' END) AS [カスタマイズ可否] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 384 THEN d2_item ELSE '' END) AS [カスタマイズ内容] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 385 THEN d2_item ELSE '' END) AS [カスタマイズ条件] ";
                tmp_sql = tmp_sql + " 			/*特定優遇賃貸住宅*/ ";
                tmp_sql = tmp_sql + " 			/*,NULL AS [特定優遇賃貸住宅チェック]*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 262 THEN d2_item ELSE '' END) AS [特優賃下限] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 261 THEN d2_item ELSE '' END) AS [特優賃上限] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 263 THEN d2_item ELSE '' END) AS [特優賃料金変動区分] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 264 THEN d2_item ELSE '' END) AS [特優賃料金上昇率] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 265 THEN d2_item ELSE '' END) AS [特優賃家賃補助年数] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 266 THEN d2_item ELSE '' END) AS [特優賃備考] ";
                tmp_sql = tmp_sql + " 			/*HOMES PRO物件流通専用項目*/	 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 386 THEN d2_item ELSE '' END) AS [鍵保管場所] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 387 THEN d2_item ELSE '' END) AS [預託先会社名] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 388 THEN d2_item ELSE '' END) AS [鍵備考] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 389 THEN d2_item ELSE '' END) AS [限定公開] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 390 THEN d2_item ELSE '' END) AS [画像ダウンロード不可] ";
                tmp_sql = tmp_sql + " 			/*営業スタッフコメント*/ ";
                tmp_sql = tmp_sql + " 			,'' AS [営業スタッフ名] ";
                tmp_sql = tmp_sql + " 			,1 AS [営業スタッフコメント設定方法] ";
                tmp_sql = tmp_sql + " 			,-1 AS [コメントテンプレート] ";
                tmp_sql = tmp_sql + " 			,-1 AS [コメント種別] ";
                tmp_sql = tmp_sql + " 			,'' AS [コメント] ";
                tmp_sql = tmp_sql + " 			/*その他*/ ";
                tmp_sql = tmp_sql + " 			,'false' AS [マンスリー可] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 379 THEN d2_item ELSE '' END) AS [駐車場契約必須] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 200 THEN d2_item ELSE '' END) AS [他社取込] ";
                tmp_sql = tmp_sql + " 			,'' AS [媒介契約年月日] ";
                tmp_sql = tmp_sql + " 			,'' AS [特別広告ポイント数] ";
                tmp_sql = tmp_sql + " 			,'false' AS [空き家バンク登録物件] ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM hy_d2 WHERE d2_data_kbn = 50 ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		GROUP BY bk_no,hy_no ";
                tmp_sql = tmp_sql + " 	) AS VW ";
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

    #region 広告補足athome

    public class Rendo_kokoku_athome_Repository
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

                sortstr = "[物件No],[部屋No],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[サイトNo],[プロのコメント],[プロのコメントスタッフID],[広告備考],[その他ペット可内容],[更新料区分],[保証金その他償却条件内容],[その他クレジット決済可],[その他クレジット決済可能条件等],[リフォーム有無],[リフォーム年月],[対象(内装関連)],[対象(外装関連)],[リフォーム内容],[リノベーション有無],[リノベーション実施年月],[リノベーション内容],[建物名表示],[部屋番号表示],[鍵現地対応],[賃料値下げ可],[初期費用値下げ可],[パノラマコンテンツID],[近隣駐輪場料金],[近隣駐輪場料金_税区分],[近隣バイク置き場料金],[近隣バイク置き場料金_税区分],[Yahoo公開不可]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 [物件No] ";
                tmp_sql = tmp_sql + " 		,[部屋No] ";
                tmp_sql = tmp_sql + " 		,30 AS [サイトNo] ";
                tmp_sql = tmp_sql + " 		,[プロのコメント] ";
                tmp_sql = tmp_sql + " 		,[プロのコメントスタッフID] ";
                tmp_sql = tmp_sql + " 		,[広告備考] ";
                tmp_sql = tmp_sql + " 		,[その他ペット可内容] ";
                tmp_sql = tmp_sql + " 		,[更新料区分] ";
                tmp_sql = tmp_sql + " 		,[保証金その他償却条件内容] ";
                tmp_sql = tmp_sql + " 		,[その他クレジット決済可] ";
                tmp_sql = tmp_sql + " 		,[その他クレジット決済可能条件等] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [リフォーム有無] = 0 THEN 'false' ";
                tmp_sql = tmp_sql + " 			WHEN [リフォーム有無] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [リフォーム有無] ";
                tmp_sql = tmp_sql + " 		,[リフォーム年月] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [対象(内装関連)] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [対象(内装関連)] ";
                tmp_sql = tmp_sql + " 		 END AS [対象(内装関連)] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [対象(外装関連)] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [対象(外装関連)] ";
                tmp_sql = tmp_sql + " 		 END AS [対象(外装関連)]	 ";
                tmp_sql = tmp_sql + " 		,[リフォーム内容] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [リノベーション有無] = 0 THEN 'false' ";
                tmp_sql = tmp_sql + " 			WHEN [リノベーション有無] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [リノベーション有無] ";
                tmp_sql = tmp_sql + " 		,[リノベーション実施年月] ";
                tmp_sql = tmp_sql + " 		,[リノベーション内容] ";
                tmp_sql = tmp_sql + " 		,[建物名表示] ";
                tmp_sql = tmp_sql + " 		,[部屋番号表示] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [鍵現地対応] = '' THEN 'false' ";
                tmp_sql = tmp_sql + " 			WHEN [鍵現地対応] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 		END AS [鍵現地対応] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [賃料値下げ可] = '' THEN 'false' ";
                tmp_sql = tmp_sql + " 			WHEN [賃料値下げ可] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 		END AS [賃料値下げ可] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [初期費用値下げ可] = '' THEN 'false' ";
                tmp_sql = tmp_sql + " 			WHEN [初期費用値下げ可] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 		END AS [初期費用値下げ可] ";
                tmp_sql = tmp_sql + " 		,[パノラマコンテンツID] ";
                tmp_sql = tmp_sql + " 		,[近隣駐輪場料金] ";
                tmp_sql = tmp_sql + " 		,[近隣駐輪場料金_税区分] ";
                tmp_sql = tmp_sql + " 		,[近隣バイク置き場料金] ";
                tmp_sql = tmp_sql + " 		,[近隣バイク置き場料金_税区分] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [Yahoo公開不可] = 0 THEN 'true' ";
                tmp_sql = tmp_sql + " 			WHEN [Yahoo公開不可] = 1 THEN 'false' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [Yahoo公開不可] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 			,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 			/*物件の特徴*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 318 THEN d2_item ELSE '' END) AS [プロのコメント] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 319 THEN d2_item ELSE '' END) AS [プロのコメントスタッフID] ";
                tmp_sql = tmp_sql + " 			/*広告備考*/ ";
                tmp_sql = tmp_sql + " 			,'' AS [広告備考] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 405 THEN d2_item ELSE '' END) AS [その他ペット可内容] ";
                tmp_sql = tmp_sql + " 			/*入金項目補足*/ ";
                tmp_sql = tmp_sql + " 			,1 AS [更新料区分] ";
                tmp_sql = tmp_sql + " 			,'' AS [保証金その他償却条件内容] ";
                tmp_sql = tmp_sql + " 			/*設備補足*/ ";
                tmp_sql = tmp_sql + " 			,'false' AS [その他クレジット決済可] ";
                tmp_sql = tmp_sql + " 			,'' AS [その他クレジット決済可能条件等] ";
                tmp_sql = tmp_sql + " 			/*リフォーム*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 382 THEN d2_item ELSE '' END) AS [リフォーム有無] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 383 THEN d2_item ELSE '' END) AS [リフォーム年月] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 384 THEN d2_item ELSE '' END) AS [対象(内装関連)] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 385 THEN d2_item ELSE '' END) AS [対象(外装関連)] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 386 THEN d2_item ELSE '' END) AS [リフォーム内容] ";
                tmp_sql = tmp_sql + " 			/*リノベーション*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 387 THEN d2_item ELSE '' END) AS [リノベーション有無] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 388 THEN d2_item ELSE '' END) AS [リノベーション実施年月] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 389 THEN d2_item ELSE '' END) AS [リノベーション内容] ";
                tmp_sql = tmp_sql + " 			/*不動産会社向けサイト用設定*/ ";
                tmp_sql = tmp_sql + " 			,'false' AS [建物名表示] ";
                tmp_sql = tmp_sql + " 			,'false' AS [部屋番号表示] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 408 THEN d2_item ELSE '' END) AS [鍵現地対応] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 406 THEN d2_item ELSE '' END) AS [賃料値下げ可] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 407 THEN d2_item ELSE '' END) AS [初期費用値下げ可] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 323 THEN d2_item ELSE '' END) AS [パノラマコンテンツID] ";
                tmp_sql = tmp_sql + " 			/*付帯駐車場補足*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 395 THEN d2_item ELSE '' END) AS [近隣駐輪場料金] ";
                tmp_sql = tmp_sql + " 			,2 AS [近隣駐輪場料金_税区分] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 392 THEN d2_item ELSE '' END) AS [近隣バイク置き場料金] ";
                tmp_sql = tmp_sql + " 			,2 AS [近隣バイク置き場料金_税区分] ";
                tmp_sql = tmp_sql + " 			/*その他*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 269 THEN d2_item ELSE '' END) AS [Yahoo公開不可] ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM hy_d2 WHERE d2_data_kbn = 90 ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		GROUP BY bk_no,hy_no ";
                tmp_sql = tmp_sql + " 	) AS VW ";
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

    #region 広告補足SUUMO

    public class Rendo_kokoku_suumo_Repository
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

                sortstr = "[物件No],[部屋No],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[サイトNo],[物件の特徴_ネット用キャッチ],[物件の特徴_フリーコメント],[敷金積増条件],[敷金積増後総額],[敷金積増後総額_単位],[駐車場備考],[特優賃有無],[特優賃入居者負担額_下限],[特優賃入居者負担額_上限],[特優賃料金変動区分],[特優賃料金変動区分上昇率],[特優賃家賃補助年数],[特優賃補足],[リフォーム時期],[リフォーム箇所],[リフォーム補足],[設備環境1],[設備環境1_距離],[設備環境2],[設備環境2_距離],[周辺環境隣接1],[周辺環境隣接1_単位],[周辺環境隣接2],[周辺環境隣接2_単位],[設備環境1F],[セキュリティ],[室内間取],[家具・家電],[リフォーム],[リフォーム(内容)],[入居条件],[SUUMO以外の災害時住宅支援サイトへの掲載],[おすすめピックアップピクト指定1],[特徴1],[おすすめピックアップピクト指定2],[特徴2],[おすすめピックアップピクト指定3],[特徴3],[物件名公開],[部屋番号公開],[詳細住所公開],[会社間物件検索コピー許可],[図面パターン],[メインキャッチ1],[メインキャッチ2],[サブキャッチ1],[サブキャッチ2],[サブキャッチ3],[サブキャッチ4],[サブキャッチ5],[サブキャッチ6],[サブキャッチ7],[サブキャッチ8],[サブキャッチ9],[サブキャッチ10],[会社間用補足フリーコメント],[SUUMO内優先画像],[貴社管理コード1],[貴社管理コード2],[見学予約機能を利用する]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 [物件No] ";
                tmp_sql = tmp_sql + " 		,[部屋No] ";
                tmp_sql = tmp_sql + " 		,40 AS [サイトNo] ";
                tmp_sql = tmp_sql + " 		,[物件の特徴_ネット用キャッチ] ";
                tmp_sql = tmp_sql + " 		,[物件の特徴_フリーコメント] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [敷金積増条件] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [敷金積増条件] ";
                tmp_sql = tmp_sql + " 		 END AS [敷金積増条件] ";
                tmp_sql = tmp_sql + " 		,[敷金積増後総額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [敷金積増後総額_単位] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [敷金積増後総額_単位] ";
                tmp_sql = tmp_sql + " 		 END AS [敷金積増後総額_単位] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [駐車場備考] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [駐車場備考] ";
                tmp_sql = tmp_sql + " 		 END AS [駐車場備考] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ";
                tmp_sql = tmp_sql + " 				[特優賃入居者負担額_下限] = '' AND [特優賃入居者負担額_上限] = '' AND  ";
                tmp_sql = tmp_sql + " 				/*[特優賃料金変動区分] = '' AND*/ [特優賃料金変動区分上昇率] = '' AND  ";
                tmp_sql = tmp_sql + " 				[特優賃家賃補助年数] = '' AND [特優賃補足] = '' ";
                tmp_sql = tmp_sql + " 			THEN 'false' ";
                tmp_sql = tmp_sql + " 			ELSE 'true' ";
                tmp_sql = tmp_sql + " 		 END AS [特優賃有無] ";
                tmp_sql = tmp_sql + " 		,[特優賃入居者負担額_下限] ";
                tmp_sql = tmp_sql + " 		,[特優賃入居者負担額_上限] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [特優賃料金変動区分] = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 			WHEN [特優賃料金変動区分] = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN [特優賃料金変動区分] = 2 THEN 2 ";
                tmp_sql = tmp_sql + " 			ELSE -1 ";
                tmp_sql = tmp_sql + " 		 END AS [特優賃料金変動区分] ";
                tmp_sql = tmp_sql + " 		,[特優賃料金変動区分上昇率] ";
                tmp_sql = tmp_sql + " 		,[特優賃家賃補助年数] ";
                tmp_sql = tmp_sql + " 		,[特優賃補足] ";
                tmp_sql = tmp_sql + " 		,[リフォーム時期] ";
                tmp_sql = tmp_sql + " 		,[リフォーム箇所] ";
                tmp_sql = tmp_sql + " 		,[リフォーム補足] ";
                tmp_sql = tmp_sql + " 		,[設備環境1] ";
                tmp_sql = tmp_sql + " 		,[設備環境1_距離] ";
                tmp_sql = tmp_sql + " 		,[設備環境2] ";
                tmp_sql = tmp_sql + " 		,[設備環境2_距離] ";
                tmp_sql = tmp_sql + " 		,[周辺環境隣接1] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [周辺環境隣接1_単位] = '' THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE [周辺環境隣接1_単位] ";
                tmp_sql = tmp_sql + " 		 END AS [周辺環境隣接1_単位] ";
                tmp_sql = tmp_sql + " 		,[周辺環境隣接2] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [周辺環境隣接2_単位] = '' THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE [周辺環境隣接2_単位] ";
                tmp_sql = tmp_sql + " 		 END AS [周辺環境隣接2_単位] ";
                tmp_sql = tmp_sql + " 		,[設備環境1F] ";
                tmp_sql = tmp_sql + " 		,[セキュリティ] ";
                tmp_sql = tmp_sql + " 		,[室内間取] ";
                tmp_sql = tmp_sql + " 		,[家具・家電] ";
                tmp_sql = tmp_sql + " 		,[リフォーム] ";
                tmp_sql = tmp_sql + " 		,[リフォーム(内容)] ";
                tmp_sql = tmp_sql + " 		,[入居条件] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO以外の災害時住宅支援サイトへの掲載] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO以外の災害時住宅支援サイトへの掲載] = 0 THEN 'false' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [SUUMO以外の災害時住宅支援サイトへの掲載] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [特徴1] = 101 THEN 1 /**/ WHEN [特徴1] = 102 THEN 1 /**/ WHEN [特徴1] = 103 THEN 1 /**/ WHEN [特徴1] = 104 THEN 1 /**/ WHEN [特徴1] = 105 THEN 1 /**/ WHEN [特徴1] = 106 THEN 1 /**/ WHEN [特徴1] = 107 THEN 1 /**/ WHEN [特徴1] = 108 THEN 1 /**/ WHEN [特徴1] = 109 THEN 1 /**/ WHEN [特徴1] = 110 THEN 1 /**/ WHEN [特徴1] = 111 THEN 1 /**/ WHEN [特徴1] = 112 THEN 1 /**/ WHEN [特徴1] = 113 THEN 1 /**/ WHEN [特徴1] = 114 THEN 1 /**/ WHEN [特徴1] = 115 THEN 1 /**/ WHEN [特徴1] = 116 THEN 1 /**/ WHEN [特徴1] = 117 THEN 1 /**/ WHEN [特徴1] = 118 THEN 1 /**/ WHEN [特徴1] = 119 THEN 1 /**/ WHEN [特徴1] = 120 THEN 1 /**/ WHEN [特徴1] = 121 THEN 1 /**/ WHEN [特徴1] = 122 THEN 1 /**/ WHEN [特徴1] = 123 THEN 1 /**/ WHEN [特徴1] = 124 THEN 1 /**/ WHEN [特徴1] = 125 THEN 1 /**/ WHEN [特徴1] = 126 THEN 1 /**/ WHEN [特徴1] = 127 THEN 1 /**/ WHEN [特徴1] = 128 THEN 1 /**/ WHEN [特徴1] = 129 THEN 1 /**/ WHEN [特徴1] = 130 THEN 1 /**/ WHEN [特徴1] = 131 THEN 1 /**/ WHEN [特徴1] = 201 THEN 2 /**/ WHEN [特徴1] = 202 THEN 2 /**/ WHEN [特徴1] = 203 THEN 2 /**/ WHEN [特徴1] = 204 THEN 2 /**/ WHEN [特徴1] = 205 THEN 2 /**/ WHEN [特徴1] = 206 THEN 2 /**/ WHEN [特徴1] = 207 THEN 2 /**/ WHEN [特徴1] = 208 THEN 2 /**/ WHEN [特徴1] = 209 THEN 2 /**/ WHEN [特徴1] = 210 THEN 2 /**/ WHEN [特徴1] = 211 THEN 2 /**/ WHEN [特徴1] = 212 THEN 2 /**/ WHEN [特徴1] = 213 THEN 2 /**/ WHEN [特徴1] = 214 THEN 2 /**/ WHEN [特徴1] = 215 THEN 2 /**/ WHEN [特徴1] = 216 THEN 2 /**/ WHEN [特徴1] = 217 THEN 2 /**/ WHEN [特徴1] = 218 THEN 2 /**/ WHEN [特徴1] = 219 THEN 2 /**/ WHEN [特徴1] = 220 THEN 2 /**/ WHEN [特徴1] = 221 THEN 2 /**/ WHEN [特徴1] = 222 THEN 2 /**/ WHEN [特徴1] = 223 THEN 2 /**/ WHEN [特徴1] = 224 THEN 2 /**/ WHEN [特徴1] = 225 THEN 2 /**/ WHEN [特徴1] = 226 THEN 2 /**/ WHEN [特徴1] = 227 THEN 2 /**/ WHEN [特徴1] = 228 THEN 2 /**/ WHEN [特徴1] = 229 THEN 2 /**/ WHEN [特徴1] = 230 THEN 2 /**/ WHEN [特徴1] = 231 THEN 2 /**/ WHEN [特徴1] = 232 THEN 2 /**/ WHEN [特徴1] = 233 THEN 2 /**/ WHEN [特徴1] = 234 THEN 2 /**/ WHEN [特徴1] = 235 THEN 2 /**/ WHEN [特徴1] = 236 THEN 2 /**/ WHEN [特徴1] = 237 THEN 2 /**/ WHEN [特徴1] = 238 THEN 2 /**/ WHEN [特徴1] = 239 THEN 2 /**/ WHEN [特徴1] = 240 THEN 2 /**/ WHEN [特徴1] = 241 THEN 2 /**/ WHEN [特徴1] = 242 THEN 2 /**/ WHEN [特徴1] = 243 THEN 2 /**/ WHEN [特徴1] = 244 THEN 2 /**/ WHEN [特徴1] = 245 THEN 2 /**/ WHEN [特徴1] = 246 THEN 2 /**/ WHEN [特徴1] = 247 THEN 2 /**/ WHEN [特徴1] = 248 THEN 2 /**/ WHEN [特徴1] = 249 THEN 2 /**/ WHEN [特徴1] = 250 THEN 2 /**/ WHEN [特徴1] = 251 THEN 2 /**/ WHEN [特徴1] = 252 THEN 2 /**/ WHEN [特徴1] = 253 THEN 2 /**/ WHEN [特徴1] = 254 THEN 2 /**/ WHEN [特徴1] = 255 THEN 2 /**/ WHEN [特徴1] = 256 THEN 2 /**/ WHEN [特徴1] = 257 THEN 2 /**/ WHEN [特徴1] = 301 THEN 3 /**/ WHEN [特徴1] = 302 THEN 3 /**/ WHEN [特徴1] = 303 THEN 3 /**/ WHEN [特徴1] = 304 THEN 3 /**/ WHEN [特徴1] = 305 THEN 3 /**/ WHEN [特徴1] = 306 THEN 3 /**/ WHEN [特徴1] = 401 THEN 4 /**/ WHEN [特徴1] = 402 THEN 4 /**/ WHEN [特徴1] = 403 THEN 4 /**/ WHEN [特徴1] = 404 THEN 4 /**/ WHEN [特徴1] = 405 THEN 4 /**/ WHEN [特徴1] = 406 THEN 4 /**/ WHEN [特徴1] = 407 THEN 4 /**/ WHEN [特徴1] = 408 THEN 4 /**/ WHEN [特徴1] = 409 THEN 4 /**/ WHEN [特徴1] = 501 THEN 5 /**/ WHEN [特徴1] = 502 THEN 5 /**/ WHEN [特徴1] = 503 THEN 5 /**/ WHEN [特徴1] = 504 THEN 5 /**/ WHEN [特徴1] = 505 THEN 5 /**/ WHEN [特徴1] = 506 THEN 5 /**/ WHEN [特徴1] = 507 THEN 5 /**/ WHEN [特徴1] = 508 THEN 5 /**/ WHEN [特徴1] = 509 THEN 5 /**/ WHEN [特徴1] = 510 THEN 5 /**/ WHEN [特徴1] = 511 THEN 5 /**/ WHEN [特徴1] = 512 THEN 5 /**/ WHEN [特徴1] = 513 THEN 5 /**/ WHEN [特徴1] = 514 THEN 5 /**/ WHEN [特徴1] = 515 THEN 5 /**/ WHEN [特徴1] = 516 THEN 5 /**/ WHEN [特徴1] = 517 THEN 5 /**/ WHEN [特徴1] = 518 THEN 5 /**/ WHEN [特徴1] = 519 THEN 5 /**/ WHEN [特徴1] = 520 THEN 5 /**/ WHEN [特徴1] = 521 THEN 5 /**/ WHEN [特徴1] = 522 THEN 5 /**/ WHEN [特徴1] = 523 THEN 5 /**/ WHEN [特徴1] = 524 THEN 5 /**/ WHEN [特徴1] = 525 THEN 5 /**/ WHEN [特徴1] = 526 THEN 5 /**/ WHEN [特徴1] = 527 THEN 5 /**/ WHEN [特徴1] = 601 THEN 6 /**/ WHEN [特徴1] = 602 THEN 6 /**/ WHEN [特徴1] = 701 THEN 7 /**/ WHEN [特徴1] = 702 THEN 7 /**/ WHEN [特徴1] = 703 THEN 7 /**/ WHEN [特徴1] = 801 THEN 8 /**/ WHEN [特徴1] = 802 THEN 8 /**/ WHEN [特徴1] = 803 THEN 8 /**/ WHEN [特徴1] = 804 THEN 8 /**/ WHEN [特徴1] = 805 THEN 8 /**/ WHEN [特徴1] = 806 THEN 8 /**/ WHEN [特徴1] = 807 THEN 8 /**/ WHEN [特徴1] = 808 THEN 8 /**/ WHEN [特徴1] = 809 THEN 8 /**/ WHEN [特徴1] = 810 THEN 8 /**/ WHEN [特徴1] = 811 THEN 8 /**/ WHEN [特徴1] = 812 THEN 8 /**/ WHEN [特徴1] = 813 THEN 8 /**/ WHEN [特徴1] = 814 THEN 8 /**/ WHEN [特徴1] = 815 THEN 8 /**/ WHEN [特徴1] = 816 THEN 8 /**/ WHEN [特徴1] = 817 THEN 8 /**/ WHEN [特徴1] = 818 THEN 8 /**/ WHEN [特徴1] = 819 THEN 8 /**/ WHEN [特徴1] = 1001 THEN 10 /**/ WHEN [特徴1] = 1002 THEN 10 /**/ WHEN [特徴1] = 1003 THEN 10 /**/ WHEN [特徴1] = 1004 THEN 10 /**/ WHEN [特徴1] = 1005 THEN 10 /**/ WHEN [特徴1] = 1006 THEN 10 /**/ WHEN [特徴1] = 1007 THEN 10 /**/ WHEN [特徴1] = 1008 THEN 10 /**/ WHEN [特徴1] = 1009 THEN 10 /**/ WHEN [特徴1] = 1010 THEN 10 /**/ WHEN [特徴1] = 1011 THEN 10 /**/ WHEN [特徴1] = 1012 THEN 10 /**/ WHEN [特徴1] = 1013 THEN 10 /**/ WHEN [特徴1] = 1014 THEN 10 /**/ WHEN [特徴1] = 1015 THEN 10 /**/ WHEN [特徴1] = 1016 THEN 10 /**/ WHEN [特徴1] = 1017 THEN 10 /**/ WHEN [特徴1] = 1018 THEN 10 /**/ WHEN [特徴1] = 1019 THEN 10 /**/ WHEN [特徴1] = 1020 THEN 10 /**/ WHEN [特徴1] = 1101 THEN 11 /**/ WHEN [特徴1] = 1102 THEN 11 /**/ WHEN [特徴1] = 1103 THEN 11 /**/ WHEN [特徴1] = 1104 THEN 11 /**/ WHEN [特徴1] = 1105 THEN 11 /**/ WHEN [特徴1] = 1106 THEN 11 /**/ WHEN [特徴1] = 1107 THEN 11 /**/ WHEN [特徴1] = 1108 THEN 11 /**/ WHEN [特徴1] = 1201 THEN 12 /**/ WHEN [特徴1] = 1202 THEN 12 /**/ WHEN [特徴1] = 1203 THEN 12 /**/ WHEN [特徴1] = 1204 THEN 12 /**/ WHEN [特徴1] = 1205 THEN 12 /**/ WHEN [特徴1] = 1206 THEN 12 /**/ WHEN [特徴1] = 1207 THEN 12 /**/ WHEN [特徴1] = 1208 THEN 12 /**/ WHEN [特徴1] = 1209 THEN 12 /**/ WHEN [特徴1] = 1210 THEN 12 /**/ WHEN [特徴1] = 1211 THEN 12 /**/ WHEN [特徴1] = 1212 THEN 12 /**/ WHEN [特徴1] = 1213 THEN 12 /**/ WHEN [特徴1] = 1214 THEN 12 /**/ WHEN [特徴1] = 1215 THEN 12 /**/ WHEN [特徴1] = 1216 THEN 12 /**/ WHEN [特徴1] = 1217 THEN 12 /**/ WHEN [特徴1] = 1218 THEN 12 /**/ WHEN [特徴1] = 1301 THEN 13 /**/ WHEN [特徴1] = 1302 THEN 13 /**/ WHEN [特徴1] = 1303 THEN 13 /**/ WHEN [特徴1] = 1304 THEN 13 /**/ WHEN [特徴1] = 1305 THEN 13 /**/ WHEN [特徴1] = 1306 THEN 13 /**/ WHEN [特徴1] = 1307 THEN 13 /**/ WHEN [特徴1] = 1308 THEN 13 /**/ WHEN [特徴1] = 1309 THEN 13 /**/ WHEN [特徴1] = 1310 THEN 13 /**/ WHEN [特徴1] = 1311 THEN 13 /**/ WHEN [特徴1] = 1312 THEN 13 /**/ WHEN [特徴1] = 1313 THEN 13 /**/ WHEN [特徴1] = 1314 THEN 13 /**/ WHEN [特徴1] = 1315 THEN 13 /**/ WHEN [特徴1] = 1316 THEN 13 /**/ WHEN [特徴1] = 1317 THEN 13 /**/ WHEN [特徴1] = 1318 THEN 13 /**/ WHEN [特徴1] = 1319 THEN 13 /**/ WHEN [特徴1] = 1320 THEN 13 /**/ WHEN [特徴1] = 1321 THEN 13 /**/ WHEN [特徴1] = 1322 THEN 13 /**/ WHEN [特徴1] = 1323 THEN 13 /**/ WHEN [特徴1] = 1324 THEN 13 /**/ WHEN [特徴1] = 1325 THEN 13 /**/ WHEN [特徴1] = 1326 THEN 13 /**/ WHEN [特徴1] = 1327 THEN 13 /**/ WHEN [特徴1] = 1328 THEN 13 /**/ WHEN [特徴1] = 1329 THEN 13 /**/ WHEN [特徴1] = 1330 THEN 13 /**/ WHEN [特徴1] = 1331 THEN 13 /**/ WHEN [特徴1] = 1332 THEN 13 /**/ WHEN [特徴1] = 1333 THEN 13 /**/ WHEN [特徴1] = 1334 THEN 13 /**/ WHEN [特徴1] = 1335 THEN 13 /**/ WHEN [特徴1] = 1401 THEN 14 /**/ WHEN [特徴1] = 1402 THEN 14 /**/ WHEN [特徴1] = 1403 THEN 14 /**/ WHEN [特徴1] = 1404 THEN 14 /**/ WHEN [特徴1] = 1405 THEN 14 /**/ WHEN [特徴1] = 1406 THEN 14 /**/ WHEN [特徴1] = 1407 THEN 14 /**/ WHEN [特徴1] = 1408 THEN 14 /**/ WHEN [特徴1] = 1409 THEN 14 /**/ WHEN [特徴1] = 1410 THEN 14 /**/ WHEN [特徴1] = 1411 THEN 14 /**/ WHEN [特徴1] = 1412 THEN 14 /**/ WHEN [特徴1] = 1413 THEN 14 /**/ WHEN [特徴1] = 1414 THEN 14 /**/ WHEN [特徴1] = 1415 THEN 14 /**/ WHEN [特徴1] = 1416 THEN 14 /**/ WHEN [特徴1] = 1417 THEN 14 /**/ WHEN [特徴1] = 1418 THEN 14 /**/ WHEN [特徴1] = 1419 THEN 14 /**/ WHEN [特徴1] = 1420 THEN 14 /**/ WHEN [特徴1] = 1421 THEN 14 /**/ WHEN [特徴1] = 1422 THEN 14 /**/ WHEN [特徴1] = 1423 THEN 14 /**/ WHEN [特徴1] = 1424 THEN 14 /**/ WHEN [特徴1] = 1425 THEN 14 /**/ WHEN [特徴1] = 1426 THEN 14 /**/ WHEN [特徴1] = 1427 THEN 14 /**/ WHEN [特徴1] = 1428 THEN 14 /**/ WHEN [特徴1] = 1429 THEN 14 /**/ WHEN [特徴1] = 1430 THEN 14 /**/ WHEN [特徴1] = 1431 THEN 14 /**/ WHEN [特徴1] = 1432 THEN 14 /**/ WHEN [特徴1] = 1433 THEN 14 /**/ WHEN [特徴1] = 1434 THEN 14 /**/ WHEN [特徴1] = 1435 THEN 14 /**/ WHEN [特徴1] = 1436 THEN 14 /**/ WHEN [特徴1] = 1437 THEN 14 /**/ WHEN [特徴1] = 1501 THEN 15 /**/ WHEN [特徴1] = 1502 THEN 15 /**/ WHEN [特徴1] = 1503 THEN 15 /**/ WHEN [特徴1] = 1504 THEN 15 /**/ WHEN [特徴1] = 1505 THEN 15 /**/ WHEN [特徴1] = 1506 THEN 15 /**/ WHEN [特徴1] = 1507 THEN 15 /**/ WHEN [特徴1] = 1508 THEN 15 /**/ WHEN [特徴1] = 1509 THEN 15 /**/ WHEN [特徴1] = 1510 THEN 15 /**/ WHEN [特徴1] = 1511 THEN 15 /**/ WHEN [特徴1] = 1512 THEN 15 /**/ WHEN [特徴1] = 1513 THEN 15 /**/ WHEN [特徴1] = 1514 THEN 15 /**/ WHEN [特徴1] = 1515 THEN 15 /**/ WHEN [特徴1] = 1516 THEN 15 /**/ WHEN [特徴1] = 1517 THEN 15 /**/ WHEN [特徴1] = 1518 THEN 15 /**/ WHEN [特徴1] = 1601 THEN 16 /**/ WHEN [特徴1] = 1602 THEN 16 /**/ WHEN [特徴1] = 1603 THEN 16 /**/ WHEN [特徴1] = 1604 THEN 16 /**/ WHEN [特徴1] = 1605 THEN 16 /**/ WHEN [特徴1] = 1701 THEN 17 /**/ WHEN [特徴1] = 1702 THEN 17 /**/ WHEN [特徴1] = 1703 THEN 17 /**/ WHEN [特徴1] = 1704 THEN 17 /**/ WHEN [特徴1] = 1705 THEN 17 /**/ WHEN [特徴1] = 1706 THEN 17 /**/ WHEN [特徴1] = 1707 THEN 17 /**/ WHEN [特徴1] = 1708 THEN 17 /**/ WHEN [特徴1] = 1709 THEN 17 /**/ WHEN [特徴1] = 1710 THEN 17 /**/ WHEN [特徴1] = 1711 THEN 17 /**/ WHEN [特徴1] = 1801 THEN 18 /**/ WHEN [特徴1] = 1802 THEN 18 /**/ WHEN [特徴1] = 1803 THEN 18 /**/ WHEN [特徴1] = 1804 THEN 18 /**/ WHEN [特徴1] = 1805 THEN 18 /**/ WHEN [特徴1] = 1806 THEN 18 /**/ WHEN [特徴1] = 1807 THEN 18 /**/ WHEN [特徴1] = 1808 THEN 18 /**/ WHEN [特徴1] = 1809 THEN 18 /**/ WHEN [特徴1] = 1810 THEN 18 /**/ WHEN [特徴1] = 1811 THEN 18 /**/ WHEN [特徴1] = 1812 THEN 18 /**/ WHEN [特徴1] = 1901 THEN 19 /**/ WHEN [特徴1] = 1902 THEN 19 /**/ WHEN [特徴1] = 1903 THEN 19 /**/ WHEN [特徴1] = 1904 THEN 19 /**/ WHEN [特徴1] = 1905 THEN 19 /**/ WHEN [特徴1] = 1906 THEN 19 /**/ WHEN [特徴1] = 1907 THEN 19 /**/ WHEN [特徴1] = 2001 THEN 20 /**/ WHEN [特徴1] = 2002 THEN 20 /**/ WHEN [特徴1] = 2003 THEN 20 /**/ WHEN [特徴1] = 2004 THEN 20 /**/ WHEN [特徴1] = 2005 THEN 20 /**/ WHEN [特徴1] = 2006 THEN 20 /**/ WHEN [特徴1] = 2007 THEN 20 /**/ WHEN [特徴1] = 2008 THEN 20 /**/ WHEN [特徴1] = 2009 THEN 20 /**/ WHEN [特徴1] = 2010 THEN 20 /**/ WHEN [特徴1] = 2011 THEN 20 /**/ WHEN [特徴1] = 2012 THEN 20 /**/ WHEN [特徴1] = 2101 THEN 21 /**/ WHEN [特徴1] = 2102 THEN 21 /**/ WHEN [特徴1] = 2103 THEN 21 /**/ WHEN [特徴1] = 2104 THEN 21 /**/ WHEN [特徴1] = 2105 THEN 21 /**/ WHEN [特徴1] = 2106 THEN 21 /**/ WHEN [特徴1] = 2107 THEN 21 /**/ WHEN [特徴1] = 2108 THEN 21 /**/ WHEN [特徴1] = 2109 THEN 21 /**/ WHEN [特徴1] = 2110 THEN 21 /**/ WHEN [特徴1] = 2111 THEN 21 /**/ WHEN [特徴1] = 2112 THEN 21 /**/ WHEN [特徴1] = 2113 THEN 21 /**/ WHEN [特徴1] = 2114 THEN 21 /**/ WHEN [特徴1] = 2115 THEN 21 /**/ WHEN [特徴1] = 2116 THEN 21 /**/ WHEN [特徴1] = 2117 THEN 21 /**/ WHEN [特徴1] = 2118 THEN 21 /**/ WHEN [特徴1] = 2119 THEN 21 /**/ WHEN [特徴1] = 2120 THEN 21 /**/ WHEN [特徴1] = 2121 THEN 21 /**/ WHEN [特徴1] = 2122 THEN 21 /**/ WHEN [特徴1] = 2123 THEN 21 /**/ WHEN [特徴1] = 2124 THEN 21 /**/ WHEN [特徴1] = 2125 THEN 21 /**/ WHEN [特徴1] = 2126 THEN 21 /**/ WHEN [特徴1] = 2127 THEN 21 /**/ WHEN [特徴1] = 2128 THEN 21 /**/ WHEN [特徴1] = 2129 THEN 21 /**/ WHEN [特徴1] = 2130 THEN 21 /**/ WHEN [特徴1] = 2131 THEN 21 /**/ WHEN [特徴1] = 2132 THEN 21 /**/ WHEN [特徴1] = 2133 THEN 21 /**/ WHEN [特徴1] = 2201 THEN 22 /**/ WHEN [特徴1] = 2202 THEN 22 /**/ WHEN [特徴1] = 2203 THEN 22 /**/ WHEN [特徴1] = 2204 THEN 22 /**/ WHEN [特徴1] = 2205 THEN 22 /**/ WHEN [特徴1] = 2206 THEN 22 /**/ WHEN [特徴1] = 2207 THEN 22 /**/ WHEN [特徴1] = 2208 THEN 22 /**/ WHEN [特徴1] = 2209 THEN 22 /**/ WHEN [特徴1] = 2210 THEN 22 /**/ WHEN [特徴1] = 2211 THEN 22 /**/ WHEN [特徴1] = 2212 THEN 22 /**/ WHEN [特徴1] = 2213 THEN 22 /**/ WHEN [特徴1] = 2214 THEN 22 /**/ WHEN [特徴1] = 2215 THEN 22 /**/ WHEN [特徴1] = 2216 THEN 22 /**/ WHEN [特徴1] = 2217 THEN 22 /**/ WHEN [特徴1] = 2218 THEN 22 /**/ WHEN [特徴1] = 2219 THEN 22 /**/ WHEN [特徴1] = 2220 THEN 22 /**/ WHEN [特徴1] = 2221 THEN 22 /**/ WHEN [特徴1] = 2222 THEN 22 /**/ WHEN [特徴1] = 2223 THEN 22 /**/ WHEN [特徴1] = 2301 THEN 23 /**/ WHEN [特徴1] = 2302 THEN 23 /**/ WHEN [特徴1] = 2303 THEN 23 /**/ WHEN [特徴1] = 2304 THEN 23 /**/ WHEN [特徴1] = 2305 THEN 23 /**/ WHEN [特徴1] = 2401 THEN 24 /**/ WHEN [特徴1] = 2402 THEN 24 /**/ WHEN [特徴1] = 2403 THEN 24 /**/ WHEN [特徴1] = 2404 THEN 24 /**/ WHEN [特徴1] = 2405 THEN 24 /**/ WHEN [特徴1] = 2406 THEN 24 /**/ WHEN [特徴1] = 2407 THEN 24 /**/ WHEN [特徴1] = 2408 THEN 24 /**/ WHEN [特徴1] = 2409 THEN 24 /**/ WHEN [特徴1] = 2410 THEN 24 /**/ WHEN [特徴1] = 2411 THEN 24 /**/ WHEN [特徴1] = 2412 THEN 24 /**/ WHEN [特徴1] = 2413 THEN 24 /**/ WHEN [特徴1] = 2414 THEN 24 /**/ WHEN [特徴1] = 2415 THEN 24 /**/ WHEN [特徴1] = 2416 THEN 24 /**/ WHEN [特徴1] = 2417 THEN 24 /**/ WHEN [特徴1] = 2601 THEN 26 /**/ WHEN [特徴1] = 2602 THEN 26 /**/ WHEN [特徴1] = 2603 THEN 26 /**/ WHEN [特徴1] = 2604 THEN 26 /**/ WHEN [特徴1] = 2605 THEN 26 /**/ WHEN [特徴1] = 2606 THEN 26 /**/ WHEN [特徴1] = 2607 THEN 26 /**/ WHEN [特徴1] = 2608 THEN 26 /**/ WHEN [特徴1] = 2609 THEN 26 /**/ WHEN [特徴1] = 2701 THEN 27 /**/ WHEN [特徴1] = 2702 THEN 27 /**/ WHEN [特徴1] = 2703 THEN 27 /**/ WHEN [特徴1] = 2704 THEN 27 /**/ WHEN [特徴1] = 2705 THEN 27 /**/ WHEN [特徴1] = 2706 THEN 27 /**/ WHEN [特徴1] = 2707 THEN 27 /**/ WHEN [特徴1] = 2708 THEN 27 /**/ WHEN [特徴1] = 2709 THEN 27 /**/ WHEN [特徴1] = 2710 THEN 27 /**/ WHEN [特徴1] = 2711 THEN 27 /**/ WHEN [特徴1] = 2712 THEN 27 /**/ WHEN [特徴1] = 2713 THEN 27 /**/ WHEN [特徴1] = 2714 THEN 27 /**/ WHEN [特徴1] = 2715 THEN 27 /**/ WHEN [特徴1] = 2716 THEN 27 /**/ WHEN [特徴1] = 2717 THEN 27 /**/ WHEN [特徴1] = 2718 THEN 27 /**/ WHEN [特徴1] = 2719 THEN 27 /**/ WHEN [特徴1] = 2720 THEN 27 /**/ WHEN [特徴1] = 2721 THEN 27 /**/ WHEN [特徴1] = 2722 THEN 27 /**/ WHEN [特徴1] = 2723 THEN 27 /**/ WHEN [特徴1] = 2724 THEN 27 /**/ WHEN [特徴1] = 2725 THEN 27 /**/ WHEN [特徴1] = 2726 THEN 27 /**/ WHEN [特徴1] = 2727 THEN 27 /**/ WHEN [特徴1] = 2728 THEN 27 /**/ WHEN [特徴1] = 2729 THEN 27 /**/ WHEN [特徴1] = 2730 THEN 27 /**/ WHEN [特徴1] = 2731 THEN 27 /**/ WHEN [特徴1] = 2732 THEN 27 /**/ WHEN [特徴1] = 2733 THEN 27 /**/ WHEN [特徴1] = 2734 THEN 27 /**/ WHEN [特徴1] = 2735 THEN 27 /**/ WHEN [特徴1] = 2736 THEN 27 /**/ WHEN [特徴1] = 2737 THEN 27 /**/ WHEN [特徴1] = 2801 THEN 28 /**/ WHEN [特徴1] = 2802 THEN 28 /**/ WHEN [特徴1] = 2803 THEN 28 /**/ WHEN [特徴1] = 2804 THEN 28 /**/ WHEN [特徴1] = 2805 THEN 28 /**/ WHEN [特徴1] = 2806 THEN 28 /**/ WHEN [特徴1] = 2807 THEN 28 /**/ WHEN [特徴1] = 2808 THEN 28 /**/ WHEN [特徴1] = 2809 THEN 28 /**/ WHEN [特徴1] = 2810 THEN 28 /**/ WHEN [特徴1] = 2811 THEN 28 /**/ WHEN [特徴1] = 2812 THEN 28 /**/ WHEN [特徴1] = 2813 THEN 28 /**/ WHEN [特徴1] = 2814 THEN 28 /**/ WHEN [特徴1] = 2815 THEN 28 /**/ WHEN [特徴1] = 2816 THEN 28 /**/ WHEN [特徴1] = 2817 THEN 28 /**/ WHEN [特徴1] = 2818 THEN 28 /**/ WHEN [特徴1] = 2819 THEN 28 /**/ WHEN [特徴1] = 2820 THEN 28 /**/ WHEN [特徴1] = 2821 THEN 28 /**/ WHEN [特徴1] = 2822 THEN 28 /**/ WHEN [特徴1] = 2901 THEN 29 /**/ WHEN [特徴1] = 2902 THEN 29 /**/ WHEN [特徴1] = 2903 THEN 29 ";
                tmp_sql = tmp_sql + " 			ELSE -1 ";
                tmp_sql = tmp_sql + " 		 END AS [おすすめピックアップピクト指定1] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [特徴1] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [特徴1] ";
                tmp_sql = tmp_sql + " 		 END AS [特徴1] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [特徴2] = 101 THEN 1 /**/ WHEN [特徴2] = 102 THEN 1 /**/ WHEN [特徴2] = 103 THEN 1 /**/ WHEN [特徴2] = 104 THEN 1 /**/ WHEN [特徴2] = 105 THEN 1 /**/ WHEN [特徴2] = 106 THEN 1 /**/ WHEN [特徴2] = 107 THEN 1 /**/ WHEN [特徴2] = 108 THEN 1 /**/ WHEN [特徴2] = 109 THEN 1 /**/ WHEN [特徴2] = 110 THEN 1 /**/ WHEN [特徴2] = 111 THEN 1 /**/ WHEN [特徴2] = 112 THEN 1 /**/ WHEN [特徴2] = 113 THEN 1 /**/ WHEN [特徴2] = 114 THEN 1 /**/ WHEN [特徴2] = 115 THEN 1 /**/ WHEN [特徴2] = 116 THEN 1 /**/ WHEN [特徴2] = 117 THEN 1 /**/ WHEN [特徴2] = 118 THEN 1 /**/ WHEN [特徴2] = 119 THEN 1 /**/ WHEN [特徴2] = 120 THEN 1 /**/ WHEN [特徴2] = 121 THEN 1 /**/ WHEN [特徴2] = 122 THEN 1 /**/ WHEN [特徴2] = 123 THEN 1 /**/ WHEN [特徴2] = 124 THEN 1 /**/ WHEN [特徴2] = 125 THEN 1 /**/ WHEN [特徴2] = 126 THEN 1 /**/ WHEN [特徴2] = 127 THEN 1 /**/ WHEN [特徴2] = 128 THEN 1 /**/ WHEN [特徴2] = 129 THEN 1 /**/ WHEN [特徴2] = 130 THEN 1 /**/ WHEN [特徴2] = 131 THEN 1 /**/ WHEN [特徴2] = 201 THEN 2 /**/ WHEN [特徴2] = 202 THEN 2 /**/ WHEN [特徴2] = 203 THEN 2 /**/ WHEN [特徴2] = 204 THEN 2 /**/ WHEN [特徴2] = 205 THEN 2 /**/ WHEN [特徴2] = 206 THEN 2 /**/ WHEN [特徴2] = 207 THEN 2 /**/ WHEN [特徴2] = 208 THEN 2 /**/ WHEN [特徴2] = 209 THEN 2 /**/ WHEN [特徴2] = 210 THEN 2 /**/ WHEN [特徴2] = 211 THEN 2 /**/ WHEN [特徴2] = 212 THEN 2 /**/ WHEN [特徴2] = 213 THEN 2 /**/ WHEN [特徴2] = 214 THEN 2 /**/ WHEN [特徴2] = 215 THEN 2 /**/ WHEN [特徴2] = 216 THEN 2 /**/ WHEN [特徴2] = 217 THEN 2 /**/ WHEN [特徴2] = 218 THEN 2 /**/ WHEN [特徴2] = 219 THEN 2 /**/ WHEN [特徴2] = 220 THEN 2 /**/ WHEN [特徴2] = 221 THEN 2 /**/ WHEN [特徴2] = 222 THEN 2 /**/ WHEN [特徴2] = 223 THEN 2 /**/ WHEN [特徴2] = 224 THEN 2 /**/ WHEN [特徴2] = 225 THEN 2 /**/ WHEN [特徴2] = 226 THEN 2 /**/ WHEN [特徴2] = 227 THEN 2 /**/ WHEN [特徴2] = 228 THEN 2 /**/ WHEN [特徴2] = 229 THEN 2 /**/ WHEN [特徴2] = 230 THEN 2 /**/ WHEN [特徴2] = 231 THEN 2 /**/ WHEN [特徴2] = 232 THEN 2 /**/ WHEN [特徴2] = 233 THEN 2 /**/ WHEN [特徴2] = 234 THEN 2 /**/ WHEN [特徴2] = 235 THEN 2 /**/ WHEN [特徴2] = 236 THEN 2 /**/ WHEN [特徴2] = 237 THEN 2 /**/ WHEN [特徴2] = 238 THEN 2 /**/ WHEN [特徴2] = 239 THEN 2 /**/ WHEN [特徴2] = 240 THEN 2 /**/ WHEN [特徴2] = 241 THEN 2 /**/ WHEN [特徴2] = 242 THEN 2 /**/ WHEN [特徴2] = 243 THEN 2 /**/ WHEN [特徴2] = 244 THEN 2 /**/ WHEN [特徴2] = 245 THEN 2 /**/ WHEN [特徴2] = 246 THEN 2 /**/ WHEN [特徴2] = 247 THEN 2 /**/ WHEN [特徴2] = 248 THEN 2 /**/ WHEN [特徴2] = 249 THEN 2 /**/ WHEN [特徴2] = 250 THEN 2 /**/ WHEN [特徴2] = 251 THEN 2 /**/ WHEN [特徴2] = 252 THEN 2 /**/ WHEN [特徴2] = 253 THEN 2 /**/ WHEN [特徴2] = 254 THEN 2 /**/ WHEN [特徴2] = 255 THEN 2 /**/ WHEN [特徴2] = 256 THEN 2 /**/ WHEN [特徴2] = 257 THEN 2 /**/ WHEN [特徴2] = 301 THEN 3 /**/ WHEN [特徴2] = 302 THEN 3 /**/ WHEN [特徴2] = 303 THEN 3 /**/ WHEN [特徴2] = 304 THEN 3 /**/ WHEN [特徴2] = 305 THEN 3 /**/ WHEN [特徴2] = 306 THEN 3 /**/ WHEN [特徴2] = 401 THEN 4 /**/ WHEN [特徴2] = 402 THEN 4 /**/ WHEN [特徴2] = 403 THEN 4 /**/ WHEN [特徴2] = 404 THEN 4 /**/ WHEN [特徴2] = 405 THEN 4 /**/ WHEN [特徴2] = 406 THEN 4 /**/ WHEN [特徴2] = 407 THEN 4 /**/ WHEN [特徴2] = 408 THEN 4 /**/ WHEN [特徴2] = 409 THEN 4 /**/ WHEN [特徴2] = 501 THEN 5 /**/ WHEN [特徴2] = 502 THEN 5 /**/ WHEN [特徴2] = 503 THEN 5 /**/ WHEN [特徴2] = 504 THEN 5 /**/ WHEN [特徴2] = 505 THEN 5 /**/ WHEN [特徴2] = 506 THEN 5 /**/ WHEN [特徴2] = 507 THEN 5 /**/ WHEN [特徴2] = 508 THEN 5 /**/ WHEN [特徴2] = 509 THEN 5 /**/ WHEN [特徴2] = 510 THEN 5 /**/ WHEN [特徴2] = 511 THEN 5 /**/ WHEN [特徴2] = 512 THEN 5 /**/ WHEN [特徴2] = 513 THEN 5 /**/ WHEN [特徴2] = 514 THEN 5 /**/ WHEN [特徴2] = 515 THEN 5 /**/ WHEN [特徴2] = 516 THEN 5 /**/ WHEN [特徴2] = 517 THEN 5 /**/ WHEN [特徴2] = 518 THEN 5 /**/ WHEN [特徴2] = 519 THEN 5 /**/ WHEN [特徴2] = 520 THEN 5 /**/ WHEN [特徴2] = 521 THEN 5 /**/ WHEN [特徴2] = 522 THEN 5 /**/ WHEN [特徴2] = 523 THEN 5 /**/ WHEN [特徴2] = 524 THEN 5 /**/ WHEN [特徴2] = 525 THEN 5 /**/ WHEN [特徴2] = 526 THEN 5 /**/ WHEN [特徴2] = 527 THEN 5 /**/ WHEN [特徴2] = 601 THEN 6 /**/ WHEN [特徴2] = 602 THEN 6 /**/ WHEN [特徴2] = 701 THEN 7 /**/ WHEN [特徴2] = 702 THEN 7 /**/ WHEN [特徴2] = 703 THEN 7 /**/ WHEN [特徴2] = 801 THEN 8 /**/ WHEN [特徴2] = 802 THEN 8 /**/ WHEN [特徴2] = 803 THEN 8 /**/ WHEN [特徴2] = 804 THEN 8 /**/ WHEN [特徴2] = 805 THEN 8 /**/ WHEN [特徴2] = 806 THEN 8 /**/ WHEN [特徴2] = 807 THEN 8 /**/ WHEN [特徴2] = 808 THEN 8 /**/ WHEN [特徴2] = 809 THEN 8 /**/ WHEN [特徴2] = 810 THEN 8 /**/ WHEN [特徴2] = 811 THEN 8 /**/ WHEN [特徴2] = 812 THEN 8 /**/ WHEN [特徴2] = 813 THEN 8 /**/ WHEN [特徴2] = 814 THEN 8 /**/ WHEN [特徴2] = 815 THEN 8 /**/ WHEN [特徴2] = 816 THEN 8 /**/ WHEN [特徴2] = 817 THEN 8 /**/ WHEN [特徴2] = 818 THEN 8 /**/ WHEN [特徴2] = 819 THEN 8 /**/ WHEN [特徴2] = 1001 THEN 10 /**/ WHEN [特徴2] = 1002 THEN 10 /**/ WHEN [特徴2] = 1003 THEN 10 /**/ WHEN [特徴2] = 1004 THEN 10 /**/ WHEN [特徴2] = 1005 THEN 10 /**/ WHEN [特徴2] = 1006 THEN 10 /**/ WHEN [特徴2] = 1007 THEN 10 /**/ WHEN [特徴2] = 1008 THEN 10 /**/ WHEN [特徴2] = 1009 THEN 10 /**/ WHEN [特徴2] = 1010 THEN 10 /**/ WHEN [特徴2] = 1011 THEN 10 /**/ WHEN [特徴2] = 1012 THEN 10 /**/ WHEN [特徴2] = 1013 THEN 10 /**/ WHEN [特徴2] = 1014 THEN 10 /**/ WHEN [特徴2] = 1015 THEN 10 /**/ WHEN [特徴2] = 1016 THEN 10 /**/ WHEN [特徴2] = 1017 THEN 10 /**/ WHEN [特徴2] = 1018 THEN 10 /**/ WHEN [特徴2] = 1019 THEN 10 /**/ WHEN [特徴2] = 1020 THEN 10 /**/ WHEN [特徴2] = 1101 THEN 11 /**/ WHEN [特徴2] = 1102 THEN 11 /**/ WHEN [特徴2] = 1103 THEN 11 /**/ WHEN [特徴2] = 1104 THEN 11 /**/ WHEN [特徴2] = 1105 THEN 11 /**/ WHEN [特徴2] = 1106 THEN 11 /**/ WHEN [特徴2] = 1107 THEN 11 /**/ WHEN [特徴2] = 1108 THEN 11 /**/ WHEN [特徴2] = 1201 THEN 12 /**/ WHEN [特徴2] = 1202 THEN 12 /**/ WHEN [特徴2] = 1203 THEN 12 /**/ WHEN [特徴2] = 1204 THEN 12 /**/ WHEN [特徴2] = 1205 THEN 12 /**/ WHEN [特徴2] = 1206 THEN 12 /**/ WHEN [特徴2] = 1207 THEN 12 /**/ WHEN [特徴2] = 1208 THEN 12 /**/ WHEN [特徴2] = 1209 THEN 12 /**/ WHEN [特徴2] = 1210 THEN 12 /**/ WHEN [特徴2] = 1211 THEN 12 /**/ WHEN [特徴2] = 1212 THEN 12 /**/ WHEN [特徴2] = 1213 THEN 12 /**/ WHEN [特徴2] = 1214 THEN 12 /**/ WHEN [特徴2] = 1215 THEN 12 /**/ WHEN [特徴2] = 1216 THEN 12 /**/ WHEN [特徴2] = 1217 THEN 12 /**/ WHEN [特徴2] = 1218 THEN 12 /**/ WHEN [特徴2] = 1301 THEN 13 /**/ WHEN [特徴2] = 1302 THEN 13 /**/ WHEN [特徴2] = 1303 THEN 13 /**/ WHEN [特徴2] = 1304 THEN 13 /**/ WHEN [特徴2] = 1305 THEN 13 /**/ WHEN [特徴2] = 1306 THEN 13 /**/ WHEN [特徴2] = 1307 THEN 13 /**/ WHEN [特徴2] = 1308 THEN 13 /**/ WHEN [特徴2] = 1309 THEN 13 /**/ WHEN [特徴2] = 1310 THEN 13 /**/ WHEN [特徴2] = 1311 THEN 13 /**/ WHEN [特徴2] = 1312 THEN 13 /**/ WHEN [特徴2] = 1313 THEN 13 /**/ WHEN [特徴2] = 1314 THEN 13 /**/ WHEN [特徴2] = 1315 THEN 13 /**/ WHEN [特徴2] = 1316 THEN 13 /**/ WHEN [特徴2] = 1317 THEN 13 /**/ WHEN [特徴2] = 1318 THEN 13 /**/ WHEN [特徴2] = 1319 THEN 13 /**/ WHEN [特徴2] = 1320 THEN 13 /**/ WHEN [特徴2] = 1321 THEN 13 /**/ WHEN [特徴2] = 1322 THEN 13 /**/ WHEN [特徴2] = 1323 THEN 13 /**/ WHEN [特徴2] = 1324 THEN 13 /**/ WHEN [特徴2] = 1325 THEN 13 /**/ WHEN [特徴2] = 1326 THEN 13 /**/ WHEN [特徴2] = 1327 THEN 13 /**/ WHEN [特徴2] = 1328 THEN 13 /**/ WHEN [特徴2] = 1329 THEN 13 /**/ WHEN [特徴2] = 1330 THEN 13 /**/ WHEN [特徴2] = 1331 THEN 13 /**/ WHEN [特徴2] = 1332 THEN 13 /**/ WHEN [特徴2] = 1333 THEN 13 /**/ WHEN [特徴2] = 1334 THEN 13 /**/ WHEN [特徴2] = 1335 THEN 13 /**/ WHEN [特徴2] = 1401 THEN 14 /**/ WHEN [特徴2] = 1402 THEN 14 /**/ WHEN [特徴2] = 1403 THEN 14 /**/ WHEN [特徴2] = 1404 THEN 14 /**/ WHEN [特徴2] = 1405 THEN 14 /**/ WHEN [特徴2] = 1406 THEN 14 /**/ WHEN [特徴2] = 1407 THEN 14 /**/ WHEN [特徴2] = 1408 THEN 14 /**/ WHEN [特徴2] = 1409 THEN 14 /**/ WHEN [特徴2] = 1410 THEN 14 /**/ WHEN [特徴2] = 1411 THEN 14 /**/ WHEN [特徴2] = 1412 THEN 14 /**/ WHEN [特徴2] = 1413 THEN 14 /**/ WHEN [特徴2] = 1414 THEN 14 /**/ WHEN [特徴2] = 1415 THEN 14 /**/ WHEN [特徴2] = 1416 THEN 14 /**/ WHEN [特徴2] = 1417 THEN 14 /**/ WHEN [特徴2] = 1418 THEN 14 /**/ WHEN [特徴2] = 1419 THEN 14 /**/ WHEN [特徴2] = 1420 THEN 14 /**/ WHEN [特徴2] = 1421 THEN 14 /**/ WHEN [特徴2] = 1422 THEN 14 /**/ WHEN [特徴2] = 1423 THEN 14 /**/ WHEN [特徴2] = 1424 THEN 14 /**/ WHEN [特徴2] = 1425 THEN 14 /**/ WHEN [特徴2] = 1426 THEN 14 /**/ WHEN [特徴2] = 1427 THEN 14 /**/ WHEN [特徴2] = 1428 THEN 14 /**/ WHEN [特徴2] = 1429 THEN 14 /**/ WHEN [特徴2] = 1430 THEN 14 /**/ WHEN [特徴2] = 1431 THEN 14 /**/ WHEN [特徴2] = 1432 THEN 14 /**/ WHEN [特徴2] = 1433 THEN 14 /**/ WHEN [特徴2] = 1434 THEN 14 /**/ WHEN [特徴2] = 1435 THEN 14 /**/ WHEN [特徴2] = 1436 THEN 14 /**/ WHEN [特徴2] = 1437 THEN 14 /**/ WHEN [特徴2] = 1501 THEN 15 /**/ WHEN [特徴2] = 1502 THEN 15 /**/ WHEN [特徴2] = 1503 THEN 15 /**/ WHEN [特徴2] = 1504 THEN 15 /**/ WHEN [特徴2] = 1505 THEN 15 /**/ WHEN [特徴2] = 1506 THEN 15 /**/ WHEN [特徴2] = 1507 THEN 15 /**/ WHEN [特徴2] = 1508 THEN 15 /**/ WHEN [特徴2] = 1509 THEN 15 /**/ WHEN [特徴2] = 1510 THEN 15 /**/ WHEN [特徴2] = 1511 THEN 15 /**/ WHEN [特徴2] = 1512 THEN 15 /**/ WHEN [特徴2] = 1513 THEN 15 /**/ WHEN [特徴2] = 1514 THEN 15 /**/ WHEN [特徴2] = 1515 THEN 15 /**/ WHEN [特徴2] = 1516 THEN 15 /**/ WHEN [特徴2] = 1517 THEN 15 /**/ WHEN [特徴2] = 1518 THEN 15 /**/ WHEN [特徴2] = 1601 THEN 16 /**/ WHEN [特徴2] = 1602 THEN 16 /**/ WHEN [特徴2] = 1603 THEN 16 /**/ WHEN [特徴2] = 1604 THEN 16 /**/ WHEN [特徴2] = 1605 THEN 16 /**/ WHEN [特徴2] = 1701 THEN 17 /**/ WHEN [特徴2] = 1702 THEN 17 /**/ WHEN [特徴2] = 1703 THEN 17 /**/ WHEN [特徴2] = 1704 THEN 17 /**/ WHEN [特徴2] = 1705 THEN 17 /**/ WHEN [特徴2] = 1706 THEN 17 /**/ WHEN [特徴2] = 1707 THEN 17 /**/ WHEN [特徴2] = 1708 THEN 17 /**/ WHEN [特徴2] = 1709 THEN 17 /**/ WHEN [特徴2] = 1710 THEN 17 /**/ WHEN [特徴2] = 1711 THEN 17 /**/ WHEN [特徴2] = 1801 THEN 18 /**/ WHEN [特徴2] = 1802 THEN 18 /**/ WHEN [特徴2] = 1803 THEN 18 /**/ WHEN [特徴2] = 1804 THEN 18 /**/ WHEN [特徴2] = 1805 THEN 18 /**/ WHEN [特徴2] = 1806 THEN 18 /**/ WHEN [特徴2] = 1807 THEN 18 /**/ WHEN [特徴2] = 1808 THEN 18 /**/ WHEN [特徴2] = 1809 THEN 18 /**/ WHEN [特徴2] = 1810 THEN 18 /**/ WHEN [特徴2] = 1811 THEN 18 /**/ WHEN [特徴2] = 1812 THEN 18 /**/ WHEN [特徴2] = 1901 THEN 19 /**/ WHEN [特徴2] = 1902 THEN 19 /**/ WHEN [特徴2] = 1903 THEN 19 /**/ WHEN [特徴2] = 1904 THEN 19 /**/ WHEN [特徴2] = 1905 THEN 19 /**/ WHEN [特徴2] = 1906 THEN 19 /**/ WHEN [特徴2] = 1907 THEN 19 /**/ WHEN [特徴2] = 2001 THEN 20 /**/ WHEN [特徴2] = 2002 THEN 20 /**/ WHEN [特徴2] = 2003 THEN 20 /**/ WHEN [特徴2] = 2004 THEN 20 /**/ WHEN [特徴2] = 2005 THEN 20 /**/ WHEN [特徴2] = 2006 THEN 20 /**/ WHEN [特徴2] = 2007 THEN 20 /**/ WHEN [特徴2] = 2008 THEN 20 /**/ WHEN [特徴2] = 2009 THEN 20 /**/ WHEN [特徴2] = 2010 THEN 20 /**/ WHEN [特徴2] = 2011 THEN 20 /**/ WHEN [特徴2] = 2012 THEN 20 /**/ WHEN [特徴2] = 2101 THEN 21 /**/ WHEN [特徴2] = 2102 THEN 21 /**/ WHEN [特徴2] = 2103 THEN 21 /**/ WHEN [特徴2] = 2104 THEN 21 /**/ WHEN [特徴2] = 2105 THEN 21 /**/ WHEN [特徴2] = 2106 THEN 21 /**/ WHEN [特徴2] = 2107 THEN 21 /**/ WHEN [特徴2] = 2108 THEN 21 /**/ WHEN [特徴2] = 2109 THEN 21 /**/ WHEN [特徴2] = 2110 THEN 21 /**/ WHEN [特徴2] = 2111 THEN 21 /**/ WHEN [特徴2] = 2112 THEN 21 /**/ WHEN [特徴2] = 2113 THEN 21 /**/ WHEN [特徴2] = 2114 THEN 21 /**/ WHEN [特徴2] = 2115 THEN 21 /**/ WHEN [特徴2] = 2116 THEN 21 /**/ WHEN [特徴2] = 2117 THEN 21 /**/ WHEN [特徴2] = 2118 THEN 21 /**/ WHEN [特徴2] = 2119 THEN 21 /**/ WHEN [特徴2] = 2120 THEN 21 /**/ WHEN [特徴2] = 2121 THEN 21 /**/ WHEN [特徴2] = 2122 THEN 21 /**/ WHEN [特徴2] = 2123 THEN 21 /**/ WHEN [特徴2] = 2124 THEN 21 /**/ WHEN [特徴2] = 2125 THEN 21 /**/ WHEN [特徴2] = 2126 THEN 21 /**/ WHEN [特徴2] = 2127 THEN 21 /**/ WHEN [特徴2] = 2128 THEN 21 /**/ WHEN [特徴2] = 2129 THEN 21 /**/ WHEN [特徴2] = 2130 THEN 21 /**/ WHEN [特徴2] = 2131 THEN 21 /**/ WHEN [特徴2] = 2132 THEN 21 /**/ WHEN [特徴2] = 2133 THEN 21 /**/ WHEN [特徴2] = 2201 THEN 22 /**/ WHEN [特徴2] = 2202 THEN 22 /**/ WHEN [特徴2] = 2203 THEN 22 /**/ WHEN [特徴2] = 2204 THEN 22 /**/ WHEN [特徴2] = 2205 THEN 22 /**/ WHEN [特徴2] = 2206 THEN 22 /**/ WHEN [特徴2] = 2207 THEN 22 /**/ WHEN [特徴2] = 2208 THEN 22 /**/ WHEN [特徴2] = 2209 THEN 22 /**/ WHEN [特徴2] = 2210 THEN 22 /**/ WHEN [特徴2] = 2211 THEN 22 /**/ WHEN [特徴2] = 2212 THEN 22 /**/ WHEN [特徴2] = 2213 THEN 22 /**/ WHEN [特徴2] = 2214 THEN 22 /**/ WHEN [特徴2] = 2215 THEN 22 /**/ WHEN [特徴2] = 2216 THEN 22 /**/ WHEN [特徴2] = 2217 THEN 22 /**/ WHEN [特徴2] = 2218 THEN 22 /**/ WHEN [特徴2] = 2219 THEN 22 /**/ WHEN [特徴2] = 2220 THEN 22 /**/ WHEN [特徴2] = 2221 THEN 22 /**/ WHEN [特徴2] = 2222 THEN 22 /**/ WHEN [特徴2] = 2223 THEN 22 /**/ WHEN [特徴2] = 2301 THEN 23 /**/ WHEN [特徴2] = 2302 THEN 23 /**/ WHEN [特徴2] = 2303 THEN 23 /**/ WHEN [特徴2] = 2304 THEN 23 /**/ WHEN [特徴2] = 2305 THEN 23 /**/ WHEN [特徴2] = 2401 THEN 24 /**/ WHEN [特徴2] = 2402 THEN 24 /**/ WHEN [特徴2] = 2403 THEN 24 /**/ WHEN [特徴2] = 2404 THEN 24 /**/ WHEN [特徴2] = 2405 THEN 24 /**/ WHEN [特徴2] = 2406 THEN 24 /**/ WHEN [特徴2] = 2407 THEN 24 /**/ WHEN [特徴2] = 2408 THEN 24 /**/ WHEN [特徴2] = 2409 THEN 24 /**/ WHEN [特徴2] = 2410 THEN 24 /**/ WHEN [特徴2] = 2411 THEN 24 /**/ WHEN [特徴2] = 2412 THEN 24 /**/ WHEN [特徴2] = 2413 THEN 24 /**/ WHEN [特徴2] = 2414 THEN 24 /**/ WHEN [特徴2] = 2415 THEN 24 /**/ WHEN [特徴2] = 2416 THEN 24 /**/ WHEN [特徴2] = 2417 THEN 24 /**/ WHEN [特徴2] = 2601 THEN 26 /**/ WHEN [特徴2] = 2602 THEN 26 /**/ WHEN [特徴2] = 2603 THEN 26 /**/ WHEN [特徴2] = 2604 THEN 26 /**/ WHEN [特徴2] = 2605 THEN 26 /**/ WHEN [特徴2] = 2606 THEN 26 /**/ WHEN [特徴2] = 2607 THEN 26 /**/ WHEN [特徴2] = 2608 THEN 26 /**/ WHEN [特徴2] = 2609 THEN 26 /**/ WHEN [特徴2] = 2701 THEN 27 /**/ WHEN [特徴2] = 2702 THEN 27 /**/ WHEN [特徴2] = 2703 THEN 27 /**/ WHEN [特徴2] = 2704 THEN 27 /**/ WHEN [特徴2] = 2705 THEN 27 /**/ WHEN [特徴2] = 2706 THEN 27 /**/ WHEN [特徴2] = 2707 THEN 27 /**/ WHEN [特徴2] = 2708 THEN 27 /**/ WHEN [特徴2] = 2709 THEN 27 /**/ WHEN [特徴2] = 2710 THEN 27 /**/ WHEN [特徴2] = 2711 THEN 27 /**/ WHEN [特徴2] = 2712 THEN 27 /**/ WHEN [特徴2] = 2713 THEN 27 /**/ WHEN [特徴2] = 2714 THEN 27 /**/ WHEN [特徴2] = 2715 THEN 27 /**/ WHEN [特徴2] = 2716 THEN 27 /**/ WHEN [特徴2] = 2717 THEN 27 /**/ WHEN [特徴2] = 2718 THEN 27 /**/ WHEN [特徴2] = 2719 THEN 27 /**/ WHEN [特徴2] = 2720 THEN 27 /**/ WHEN [特徴2] = 2721 THEN 27 /**/ WHEN [特徴2] = 2722 THEN 27 /**/ WHEN [特徴2] = 2723 THEN 27 /**/ WHEN [特徴2] = 2724 THEN 27 /**/ WHEN [特徴2] = 2725 THEN 27 /**/ WHEN [特徴2] = 2726 THEN 27 /**/ WHEN [特徴2] = 2727 THEN 27 /**/ WHEN [特徴2] = 2728 THEN 27 /**/ WHEN [特徴2] = 2729 THEN 27 /**/ WHEN [特徴2] = 2730 THEN 27 /**/ WHEN [特徴2] = 2731 THEN 27 /**/ WHEN [特徴2] = 2732 THEN 27 /**/ WHEN [特徴2] = 2733 THEN 27 /**/ WHEN [特徴2] = 2734 THEN 27 /**/ WHEN [特徴2] = 2735 THEN 27 /**/ WHEN [特徴2] = 2736 THEN 27 /**/ WHEN [特徴2] = 2737 THEN 27 /**/ WHEN [特徴2] = 2801 THEN 28 /**/ WHEN [特徴2] = 2802 THEN 28 /**/ WHEN [特徴2] = 2803 THEN 28 /**/ WHEN [特徴2] = 2804 THEN 28 /**/ WHEN [特徴2] = 2805 THEN 28 /**/ WHEN [特徴2] = 2806 THEN 28 /**/ WHEN [特徴2] = 2807 THEN 28 /**/ WHEN [特徴2] = 2808 THEN 28 /**/ WHEN [特徴2] = 2809 THEN 28 /**/ WHEN [特徴2] = 2810 THEN 28 /**/ WHEN [特徴2] = 2811 THEN 28 /**/ WHEN [特徴2] = 2812 THEN 28 /**/ WHEN [特徴2] = 2813 THEN 28 /**/ WHEN [特徴2] = 2814 THEN 28 /**/ WHEN [特徴2] = 2815 THEN 28 /**/ WHEN [特徴2] = 2816 THEN 28 /**/ WHEN [特徴2] = 2817 THEN 28 /**/ WHEN [特徴2] = 2818 THEN 28 /**/ WHEN [特徴2] = 2819 THEN 28 /**/ WHEN [特徴2] = 2820 THEN 28 /**/ WHEN [特徴2] = 2821 THEN 28 /**/ WHEN [特徴2] = 2822 THEN 28 /**/ WHEN [特徴2] = 2901 THEN 29 /**/ WHEN [特徴2] = 2902 THEN 29 /**/ WHEN [特徴2] = 2903 THEN 29 ";
                tmp_sql = tmp_sql + " 			ELSE -1 ";
                tmp_sql = tmp_sql + " 		 END AS [おすすめピックアップピクト指定2] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [特徴2] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [特徴2] ";
                tmp_sql = tmp_sql + " 		 END AS [特徴2] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [特徴3] = 101 THEN 1 /**/ WHEN [特徴3] = 102 THEN 1 /**/ WHEN [特徴3] = 103 THEN 1 /**/ WHEN [特徴3] = 104 THEN 1 /**/ WHEN [特徴3] = 105 THEN 1 /**/ WHEN [特徴3] = 106 THEN 1 /**/ WHEN [特徴3] = 107 THEN 1 /**/ WHEN [特徴3] = 108 THEN 1 /**/ WHEN [特徴3] = 109 THEN 1 /**/ WHEN [特徴3] = 110 THEN 1 /**/ WHEN [特徴3] = 111 THEN 1 /**/ WHEN [特徴3] = 112 THEN 1 /**/ WHEN [特徴3] = 113 THEN 1 /**/ WHEN [特徴3] = 114 THEN 1 /**/ WHEN [特徴3] = 115 THEN 1 /**/ WHEN [特徴3] = 116 THEN 1 /**/ WHEN [特徴3] = 117 THEN 1 /**/ WHEN [特徴3] = 118 THEN 1 /**/ WHEN [特徴3] = 119 THEN 1 /**/ WHEN [特徴3] = 120 THEN 1 /**/ WHEN [特徴3] = 121 THEN 1 /**/ WHEN [特徴3] = 122 THEN 1 /**/ WHEN [特徴3] = 123 THEN 1 /**/ WHEN [特徴3] = 124 THEN 1 /**/ WHEN [特徴3] = 125 THEN 1 /**/ WHEN [特徴3] = 126 THEN 1 /**/ WHEN [特徴3] = 127 THEN 1 /**/ WHEN [特徴3] = 128 THEN 1 /**/ WHEN [特徴3] = 129 THEN 1 /**/ WHEN [特徴3] = 130 THEN 1 /**/ WHEN [特徴3] = 131 THEN 1 /**/ WHEN [特徴3] = 201 THEN 2 /**/ WHEN [特徴3] = 202 THEN 2 /**/ WHEN [特徴3] = 203 THEN 2 /**/ WHEN [特徴3] = 204 THEN 2 /**/ WHEN [特徴3] = 205 THEN 2 /**/ WHEN [特徴3] = 206 THEN 2 /**/ WHEN [特徴3] = 207 THEN 2 /**/ WHEN [特徴3] = 208 THEN 2 /**/ WHEN [特徴3] = 209 THEN 2 /**/ WHEN [特徴3] = 210 THEN 2 /**/ WHEN [特徴3] = 211 THEN 2 /**/ WHEN [特徴3] = 212 THEN 2 /**/ WHEN [特徴3] = 213 THEN 2 /**/ WHEN [特徴3] = 214 THEN 2 /**/ WHEN [特徴3] = 215 THEN 2 /**/ WHEN [特徴3] = 216 THEN 2 /**/ WHEN [特徴3] = 217 THEN 2 /**/ WHEN [特徴3] = 218 THEN 2 /**/ WHEN [特徴3] = 219 THEN 2 /**/ WHEN [特徴3] = 220 THEN 2 /**/ WHEN [特徴3] = 221 THEN 2 /**/ WHEN [特徴3] = 222 THEN 2 /**/ WHEN [特徴3] = 223 THEN 2 /**/ WHEN [特徴3] = 224 THEN 2 /**/ WHEN [特徴3] = 225 THEN 2 /**/ WHEN [特徴3] = 226 THEN 2 /**/ WHEN [特徴3] = 227 THEN 2 /**/ WHEN [特徴3] = 228 THEN 2 /**/ WHEN [特徴3] = 229 THEN 2 /**/ WHEN [特徴3] = 230 THEN 2 /**/ WHEN [特徴3] = 231 THEN 2 /**/ WHEN [特徴3] = 232 THEN 2 /**/ WHEN [特徴3] = 233 THEN 2 /**/ WHEN [特徴3] = 234 THEN 2 /**/ WHEN [特徴3] = 235 THEN 2 /**/ WHEN [特徴3] = 236 THEN 2 /**/ WHEN [特徴3] = 237 THEN 2 /**/ WHEN [特徴3] = 238 THEN 2 /**/ WHEN [特徴3] = 239 THEN 2 /**/ WHEN [特徴3] = 240 THEN 2 /**/ WHEN [特徴3] = 241 THEN 2 /**/ WHEN [特徴3] = 242 THEN 2 /**/ WHEN [特徴3] = 243 THEN 2 /**/ WHEN [特徴3] = 244 THEN 2 /**/ WHEN [特徴3] = 245 THEN 2 /**/ WHEN [特徴3] = 246 THEN 2 /**/ WHEN [特徴3] = 247 THEN 2 /**/ WHEN [特徴3] = 248 THEN 2 /**/ WHEN [特徴3] = 249 THEN 2 /**/ WHEN [特徴3] = 250 THEN 2 /**/ WHEN [特徴3] = 251 THEN 2 /**/ WHEN [特徴3] = 252 THEN 2 /**/ WHEN [特徴3] = 253 THEN 2 /**/ WHEN [特徴3] = 254 THEN 2 /**/ WHEN [特徴3] = 255 THEN 2 /**/ WHEN [特徴3] = 256 THEN 2 /**/ WHEN [特徴3] = 257 THEN 2 /**/ WHEN [特徴3] = 301 THEN 3 /**/ WHEN [特徴3] = 302 THEN 3 /**/ WHEN [特徴3] = 303 THEN 3 /**/ WHEN [特徴3] = 304 THEN 3 /**/ WHEN [特徴3] = 305 THEN 3 /**/ WHEN [特徴3] = 306 THEN 3 /**/ WHEN [特徴3] = 401 THEN 4 /**/ WHEN [特徴3] = 402 THEN 4 /**/ WHEN [特徴3] = 403 THEN 4 /**/ WHEN [特徴3] = 404 THEN 4 /**/ WHEN [特徴3] = 405 THEN 4 /**/ WHEN [特徴3] = 406 THEN 4 /**/ WHEN [特徴3] = 407 THEN 4 /**/ WHEN [特徴3] = 408 THEN 4 /**/ WHEN [特徴3] = 409 THEN 4 /**/ WHEN [特徴3] = 501 THEN 5 /**/ WHEN [特徴3] = 502 THEN 5 /**/ WHEN [特徴3] = 503 THEN 5 /**/ WHEN [特徴3] = 504 THEN 5 /**/ WHEN [特徴3] = 505 THEN 5 /**/ WHEN [特徴3] = 506 THEN 5 /**/ WHEN [特徴3] = 507 THEN 5 /**/ WHEN [特徴3] = 508 THEN 5 /**/ WHEN [特徴3] = 509 THEN 5 /**/ WHEN [特徴3] = 510 THEN 5 /**/ WHEN [特徴3] = 511 THEN 5 /**/ WHEN [特徴3] = 512 THEN 5 /**/ WHEN [特徴3] = 513 THEN 5 /**/ WHEN [特徴3] = 514 THEN 5 /**/ WHEN [特徴3] = 515 THEN 5 /**/ WHEN [特徴3] = 516 THEN 5 /**/ WHEN [特徴3] = 517 THEN 5 /**/ WHEN [特徴3] = 518 THEN 5 /**/ WHEN [特徴3] = 519 THEN 5 /**/ WHEN [特徴3] = 520 THEN 5 /**/ WHEN [特徴3] = 521 THEN 5 /**/ WHEN [特徴3] = 522 THEN 5 /**/ WHEN [特徴3] = 523 THEN 5 /**/ WHEN [特徴3] = 524 THEN 5 /**/ WHEN [特徴3] = 525 THEN 5 /**/ WHEN [特徴3] = 526 THEN 5 /**/ WHEN [特徴3] = 527 THEN 5 /**/ WHEN [特徴3] = 601 THEN 6 /**/ WHEN [特徴3] = 602 THEN 6 /**/ WHEN [特徴3] = 701 THEN 7 /**/ WHEN [特徴3] = 702 THEN 7 /**/ WHEN [特徴3] = 703 THEN 7 /**/ WHEN [特徴3] = 801 THEN 8 /**/ WHEN [特徴3] = 802 THEN 8 /**/ WHEN [特徴3] = 803 THEN 8 /**/ WHEN [特徴3] = 804 THEN 8 /**/ WHEN [特徴3] = 805 THEN 8 /**/ WHEN [特徴3] = 806 THEN 8 /**/ WHEN [特徴3] = 807 THEN 8 /**/ WHEN [特徴3] = 808 THEN 8 /**/ WHEN [特徴3] = 809 THEN 8 /**/ WHEN [特徴3] = 810 THEN 8 /**/ WHEN [特徴3] = 811 THEN 8 /**/ WHEN [特徴3] = 812 THEN 8 /**/ WHEN [特徴3] = 813 THEN 8 /**/ WHEN [特徴3] = 814 THEN 8 /**/ WHEN [特徴3] = 815 THEN 8 /**/ WHEN [特徴3] = 816 THEN 8 /**/ WHEN [特徴3] = 817 THEN 8 /**/ WHEN [特徴3] = 818 THEN 8 /**/ WHEN [特徴3] = 819 THEN 8 /**/ WHEN [特徴3] = 1001 THEN 10 /**/ WHEN [特徴3] = 1002 THEN 10 /**/ WHEN [特徴3] = 1003 THEN 10 /**/ WHEN [特徴3] = 1004 THEN 10 /**/ WHEN [特徴3] = 1005 THEN 10 /**/ WHEN [特徴3] = 1006 THEN 10 /**/ WHEN [特徴3] = 1007 THEN 10 /**/ WHEN [特徴3] = 1008 THEN 10 /**/ WHEN [特徴3] = 1009 THEN 10 /**/ WHEN [特徴3] = 1010 THEN 10 /**/ WHEN [特徴3] = 1011 THEN 10 /**/ WHEN [特徴3] = 1012 THEN 10 /**/ WHEN [特徴3] = 1013 THEN 10 /**/ WHEN [特徴3] = 1014 THEN 10 /**/ WHEN [特徴3] = 1015 THEN 10 /**/ WHEN [特徴3] = 1016 THEN 10 /**/ WHEN [特徴3] = 1017 THEN 10 /**/ WHEN [特徴3] = 1018 THEN 10 /**/ WHEN [特徴3] = 1019 THEN 10 /**/ WHEN [特徴3] = 1020 THEN 10 /**/ WHEN [特徴3] = 1101 THEN 11 /**/ WHEN [特徴3] = 1102 THEN 11 /**/ WHEN [特徴3] = 1103 THEN 11 /**/ WHEN [特徴3] = 1104 THEN 11 /**/ WHEN [特徴3] = 1105 THEN 11 /**/ WHEN [特徴3] = 1106 THEN 11 /**/ WHEN [特徴3] = 1107 THEN 11 /**/ WHEN [特徴3] = 1108 THEN 11 /**/ WHEN [特徴3] = 1201 THEN 12 /**/ WHEN [特徴3] = 1202 THEN 12 /**/ WHEN [特徴3] = 1203 THEN 12 /**/ WHEN [特徴3] = 1204 THEN 12 /**/ WHEN [特徴3] = 1205 THEN 12 /**/ WHEN [特徴3] = 1206 THEN 12 /**/ WHEN [特徴3] = 1207 THEN 12 /**/ WHEN [特徴3] = 1208 THEN 12 /**/ WHEN [特徴3] = 1209 THEN 12 /**/ WHEN [特徴3] = 1210 THEN 12 /**/ WHEN [特徴3] = 1211 THEN 12 /**/ WHEN [特徴3] = 1212 THEN 12 /**/ WHEN [特徴3] = 1213 THEN 12 /**/ WHEN [特徴3] = 1214 THEN 12 /**/ WHEN [特徴3] = 1215 THEN 12 /**/ WHEN [特徴3] = 1216 THEN 12 /**/ WHEN [特徴3] = 1217 THEN 12 /**/ WHEN [特徴3] = 1218 THEN 12 /**/ WHEN [特徴3] = 1301 THEN 13 /**/ WHEN [特徴3] = 1302 THEN 13 /**/ WHEN [特徴3] = 1303 THEN 13 /**/ WHEN [特徴3] = 1304 THEN 13 /**/ WHEN [特徴3] = 1305 THEN 13 /**/ WHEN [特徴3] = 1306 THEN 13 /**/ WHEN [特徴3] = 1307 THEN 13 /**/ WHEN [特徴3] = 1308 THEN 13 /**/ WHEN [特徴3] = 1309 THEN 13 /**/ WHEN [特徴3] = 1310 THEN 13 /**/ WHEN [特徴3] = 1311 THEN 13 /**/ WHEN [特徴3] = 1312 THEN 13 /**/ WHEN [特徴3] = 1313 THEN 13 /**/ WHEN [特徴3] = 1314 THEN 13 /**/ WHEN [特徴3] = 1315 THEN 13 /**/ WHEN [特徴3] = 1316 THEN 13 /**/ WHEN [特徴3] = 1317 THEN 13 /**/ WHEN [特徴3] = 1318 THEN 13 /**/ WHEN [特徴3] = 1319 THEN 13 /**/ WHEN [特徴3] = 1320 THEN 13 /**/ WHEN [特徴3] = 1321 THEN 13 /**/ WHEN [特徴3] = 1322 THEN 13 /**/ WHEN [特徴3] = 1323 THEN 13 /**/ WHEN [特徴3] = 1324 THEN 13 /**/ WHEN [特徴3] = 1325 THEN 13 /**/ WHEN [特徴3] = 1326 THEN 13 /**/ WHEN [特徴3] = 1327 THEN 13 /**/ WHEN [特徴3] = 1328 THEN 13 /**/ WHEN [特徴3] = 1329 THEN 13 /**/ WHEN [特徴3] = 1330 THEN 13 /**/ WHEN [特徴3] = 1331 THEN 13 /**/ WHEN [特徴3] = 1332 THEN 13 /**/ WHEN [特徴3] = 1333 THEN 13 /**/ WHEN [特徴3] = 1334 THEN 13 /**/ WHEN [特徴3] = 1335 THEN 13 /**/ WHEN [特徴3] = 1401 THEN 14 /**/ WHEN [特徴3] = 1402 THEN 14 /**/ WHEN [特徴3] = 1403 THEN 14 /**/ WHEN [特徴3] = 1404 THEN 14 /**/ WHEN [特徴3] = 1405 THEN 14 /**/ WHEN [特徴3] = 1406 THEN 14 /**/ WHEN [特徴3] = 1407 THEN 14 /**/ WHEN [特徴3] = 1408 THEN 14 /**/ WHEN [特徴3] = 1409 THEN 14 /**/ WHEN [特徴3] = 1410 THEN 14 /**/ WHEN [特徴3] = 1411 THEN 14 /**/ WHEN [特徴3] = 1412 THEN 14 /**/ WHEN [特徴3] = 1413 THEN 14 /**/ WHEN [特徴3] = 1414 THEN 14 /**/ WHEN [特徴3] = 1415 THEN 14 /**/ WHEN [特徴3] = 1416 THEN 14 /**/ WHEN [特徴3] = 1417 THEN 14 /**/ WHEN [特徴3] = 1418 THEN 14 /**/ WHEN [特徴3] = 1419 THEN 14 /**/ WHEN [特徴3] = 1420 THEN 14 /**/ WHEN [特徴3] = 1421 THEN 14 /**/ WHEN [特徴3] = 1422 THEN 14 /**/ WHEN [特徴3] = 1423 THEN 14 /**/ WHEN [特徴3] = 1424 THEN 14 /**/ WHEN [特徴3] = 1425 THEN 14 /**/ WHEN [特徴3] = 1426 THEN 14 /**/ WHEN [特徴3] = 1427 THEN 14 /**/ WHEN [特徴3] = 1428 THEN 14 /**/ WHEN [特徴3] = 1429 THEN 14 /**/ WHEN [特徴3] = 1430 THEN 14 /**/ WHEN [特徴3] = 1431 THEN 14 /**/ WHEN [特徴3] = 1432 THEN 14 /**/ WHEN [特徴3] = 1433 THEN 14 /**/ WHEN [特徴3] = 1434 THEN 14 /**/ WHEN [特徴3] = 1435 THEN 14 /**/ WHEN [特徴3] = 1436 THEN 14 /**/ WHEN [特徴3] = 1437 THEN 14 /**/ WHEN [特徴3] = 1501 THEN 15 /**/ WHEN [特徴3] = 1502 THEN 15 /**/ WHEN [特徴3] = 1503 THEN 15 /**/ WHEN [特徴3] = 1504 THEN 15 /**/ WHEN [特徴3] = 1505 THEN 15 /**/ WHEN [特徴3] = 1506 THEN 15 /**/ WHEN [特徴3] = 1507 THEN 15 /**/ WHEN [特徴3] = 1508 THEN 15 /**/ WHEN [特徴3] = 1509 THEN 15 /**/ WHEN [特徴3] = 1510 THEN 15 /**/ WHEN [特徴3] = 1511 THEN 15 /**/ WHEN [特徴3] = 1512 THEN 15 /**/ WHEN [特徴3] = 1513 THEN 15 /**/ WHEN [特徴3] = 1514 THEN 15 /**/ WHEN [特徴3] = 1515 THEN 15 /**/ WHEN [特徴3] = 1516 THEN 15 /**/ WHEN [特徴3] = 1517 THEN 15 /**/ WHEN [特徴3] = 1518 THEN 15 /**/ WHEN [特徴3] = 1601 THEN 16 /**/ WHEN [特徴3] = 1602 THEN 16 /**/ WHEN [特徴3] = 1603 THEN 16 /**/ WHEN [特徴3] = 1604 THEN 16 /**/ WHEN [特徴3] = 1605 THEN 16 /**/ WHEN [特徴3] = 1701 THEN 17 /**/ WHEN [特徴3] = 1702 THEN 17 /**/ WHEN [特徴3] = 1703 THEN 17 /**/ WHEN [特徴3] = 1704 THEN 17 /**/ WHEN [特徴3] = 1705 THEN 17 /**/ WHEN [特徴3] = 1706 THEN 17 /**/ WHEN [特徴3] = 1707 THEN 17 /**/ WHEN [特徴3] = 1708 THEN 17 /**/ WHEN [特徴3] = 1709 THEN 17 /**/ WHEN [特徴3] = 1710 THEN 17 /**/ WHEN [特徴3] = 1711 THEN 17 /**/ WHEN [特徴3] = 1801 THEN 18 /**/ WHEN [特徴3] = 1802 THEN 18 /**/ WHEN [特徴3] = 1803 THEN 18 /**/ WHEN [特徴3] = 1804 THEN 18 /**/ WHEN [特徴3] = 1805 THEN 18 /**/ WHEN [特徴3] = 1806 THEN 18 /**/ WHEN [特徴3] = 1807 THEN 18 /**/ WHEN [特徴3] = 1808 THEN 18 /**/ WHEN [特徴3] = 1809 THEN 18 /**/ WHEN [特徴3] = 1810 THEN 18 /**/ WHEN [特徴3] = 1811 THEN 18 /**/ WHEN [特徴3] = 1812 THEN 18 /**/ WHEN [特徴3] = 1901 THEN 19 /**/ WHEN [特徴3] = 1902 THEN 19 /**/ WHEN [特徴3] = 1903 THEN 19 /**/ WHEN [特徴3] = 1904 THEN 19 /**/ WHEN [特徴3] = 1905 THEN 19 /**/ WHEN [特徴3] = 1906 THEN 19 /**/ WHEN [特徴3] = 1907 THEN 19 /**/ WHEN [特徴3] = 2001 THEN 20 /**/ WHEN [特徴3] = 2002 THEN 20 /**/ WHEN [特徴3] = 2003 THEN 20 /**/ WHEN [特徴3] = 2004 THEN 20 /**/ WHEN [特徴3] = 2005 THEN 20 /**/ WHEN [特徴3] = 2006 THEN 20 /**/ WHEN [特徴3] = 2007 THEN 20 /**/ WHEN [特徴3] = 2008 THEN 20 /**/ WHEN [特徴3] = 2009 THEN 20 /**/ WHEN [特徴3] = 2010 THEN 20 /**/ WHEN [特徴3] = 2011 THEN 20 /**/ WHEN [特徴3] = 2012 THEN 20 /**/ WHEN [特徴3] = 2101 THEN 21 /**/ WHEN [特徴3] = 2102 THEN 21 /**/ WHEN [特徴3] = 2103 THEN 21 /**/ WHEN [特徴3] = 2104 THEN 21 /**/ WHEN [特徴3] = 2105 THEN 21 /**/ WHEN [特徴3] = 2106 THEN 21 /**/ WHEN [特徴3] = 2107 THEN 21 /**/ WHEN [特徴3] = 2108 THEN 21 /**/ WHEN [特徴3] = 2109 THEN 21 /**/ WHEN [特徴3] = 2110 THEN 21 /**/ WHEN [特徴3] = 2111 THEN 21 /**/ WHEN [特徴3] = 2112 THEN 21 /**/ WHEN [特徴3] = 2113 THEN 21 /**/ WHEN [特徴3] = 2114 THEN 21 /**/ WHEN [特徴3] = 2115 THEN 21 /**/ WHEN [特徴3] = 2116 THEN 21 /**/ WHEN [特徴3] = 2117 THEN 21 /**/ WHEN [特徴3] = 2118 THEN 21 /**/ WHEN [特徴3] = 2119 THEN 21 /**/ WHEN [特徴3] = 2120 THEN 21 /**/ WHEN [特徴3] = 2121 THEN 21 /**/ WHEN [特徴3] = 2122 THEN 21 /**/ WHEN [特徴3] = 2123 THEN 21 /**/ WHEN [特徴3] = 2124 THEN 21 /**/ WHEN [特徴3] = 2125 THEN 21 /**/ WHEN [特徴3] = 2126 THEN 21 /**/ WHEN [特徴3] = 2127 THEN 21 /**/ WHEN [特徴3] = 2128 THEN 21 /**/ WHEN [特徴3] = 2129 THEN 21 /**/ WHEN [特徴3] = 2130 THEN 21 /**/ WHEN [特徴3] = 2131 THEN 21 /**/ WHEN [特徴3] = 2132 THEN 21 /**/ WHEN [特徴3] = 2133 THEN 21 /**/ WHEN [特徴3] = 2201 THEN 22 /**/ WHEN [特徴3] = 2202 THEN 22 /**/ WHEN [特徴3] = 2203 THEN 22 /**/ WHEN [特徴3] = 2204 THEN 22 /**/ WHEN [特徴3] = 2205 THEN 22 /**/ WHEN [特徴3] = 2206 THEN 22 /**/ WHEN [特徴3] = 2207 THEN 22 /**/ WHEN [特徴3] = 2208 THEN 22 /**/ WHEN [特徴3] = 2209 THEN 22 /**/ WHEN [特徴3] = 2210 THEN 22 /**/ WHEN [特徴3] = 2211 THEN 22 /**/ WHEN [特徴3] = 2212 THEN 22 /**/ WHEN [特徴3] = 2213 THEN 22 /**/ WHEN [特徴3] = 2214 THEN 22 /**/ WHEN [特徴3] = 2215 THEN 22 /**/ WHEN [特徴3] = 2216 THEN 22 /**/ WHEN [特徴3] = 2217 THEN 22 /**/ WHEN [特徴3] = 2218 THEN 22 /**/ WHEN [特徴3] = 2219 THEN 22 /**/ WHEN [特徴3] = 2220 THEN 22 /**/ WHEN [特徴3] = 2221 THEN 22 /**/ WHEN [特徴3] = 2222 THEN 22 /**/ WHEN [特徴3] = 2223 THEN 22 /**/ WHEN [特徴3] = 2301 THEN 23 /**/ WHEN [特徴3] = 2302 THEN 23 /**/ WHEN [特徴3] = 2303 THEN 23 /**/ WHEN [特徴3] = 2304 THEN 23 /**/ WHEN [特徴3] = 2305 THEN 23 /**/ WHEN [特徴3] = 2401 THEN 24 /**/ WHEN [特徴3] = 2402 THEN 24 /**/ WHEN [特徴3] = 2403 THEN 24 /**/ WHEN [特徴3] = 2404 THEN 24 /**/ WHEN [特徴3] = 2405 THEN 24 /**/ WHEN [特徴3] = 2406 THEN 24 /**/ WHEN [特徴3] = 2407 THEN 24 /**/ WHEN [特徴3] = 2408 THEN 24 /**/ WHEN [特徴3] = 2409 THEN 24 /**/ WHEN [特徴3] = 2410 THEN 24 /**/ WHEN [特徴3] = 2411 THEN 24 /**/ WHEN [特徴3] = 2412 THEN 24 /**/ WHEN [特徴3] = 2413 THEN 24 /**/ WHEN [特徴3] = 2414 THEN 24 /**/ WHEN [特徴3] = 2415 THEN 24 /**/ WHEN [特徴3] = 2416 THEN 24 /**/ WHEN [特徴3] = 2417 THEN 24 /**/ WHEN [特徴3] = 2601 THEN 26 /**/ WHEN [特徴3] = 2602 THEN 26 /**/ WHEN [特徴3] = 2603 THEN 26 /**/ WHEN [特徴3] = 2604 THEN 26 /**/ WHEN [特徴3] = 2605 THEN 26 /**/ WHEN [特徴3] = 2606 THEN 26 /**/ WHEN [特徴3] = 2607 THEN 26 /**/ WHEN [特徴3] = 2608 THEN 26 /**/ WHEN [特徴3] = 2609 THEN 26 /**/ WHEN [特徴3] = 2701 THEN 27 /**/ WHEN [特徴3] = 2702 THEN 27 /**/ WHEN [特徴3] = 2703 THEN 27 /**/ WHEN [特徴3] = 2704 THEN 27 /**/ WHEN [特徴3] = 2705 THEN 27 /**/ WHEN [特徴3] = 2706 THEN 27 /**/ WHEN [特徴3] = 2707 THEN 27 /**/ WHEN [特徴3] = 2708 THEN 27 /**/ WHEN [特徴3] = 2709 THEN 27 /**/ WHEN [特徴3] = 2710 THEN 27 /**/ WHEN [特徴3] = 2711 THEN 27 /**/ WHEN [特徴3] = 2712 THEN 27 /**/ WHEN [特徴3] = 2713 THEN 27 /**/ WHEN [特徴3] = 2714 THEN 27 /**/ WHEN [特徴3] = 2715 THEN 27 /**/ WHEN [特徴3] = 2716 THEN 27 /**/ WHEN [特徴3] = 2717 THEN 27 /**/ WHEN [特徴3] = 2718 THEN 27 /**/ WHEN [特徴3] = 2719 THEN 27 /**/ WHEN [特徴3] = 2720 THEN 27 /**/ WHEN [特徴3] = 2721 THEN 27 /**/ WHEN [特徴3] = 2722 THEN 27 /**/ WHEN [特徴3] = 2723 THEN 27 /**/ WHEN [特徴3] = 2724 THEN 27 /**/ WHEN [特徴3] = 2725 THEN 27 /**/ WHEN [特徴3] = 2726 THEN 27 /**/ WHEN [特徴3] = 2727 THEN 27 /**/ WHEN [特徴3] = 2728 THEN 27 /**/ WHEN [特徴3] = 2729 THEN 27 /**/ WHEN [特徴3] = 2730 THEN 27 /**/ WHEN [特徴3] = 2731 THEN 27 /**/ WHEN [特徴3] = 2732 THEN 27 /**/ WHEN [特徴3] = 2733 THEN 27 /**/ WHEN [特徴3] = 2734 THEN 27 /**/ WHEN [特徴3] = 2735 THEN 27 /**/ WHEN [特徴3] = 2736 THEN 27 /**/ WHEN [特徴3] = 2737 THEN 27 /**/ WHEN [特徴3] = 2801 THEN 28 /**/ WHEN [特徴3] = 2802 THEN 28 /**/ WHEN [特徴3] = 2803 THEN 28 /**/ WHEN [特徴3] = 2804 THEN 28 /**/ WHEN [特徴3] = 2805 THEN 28 /**/ WHEN [特徴3] = 2806 THEN 28 /**/ WHEN [特徴3] = 2807 THEN 28 /**/ WHEN [特徴3] = 2808 THEN 28 /**/ WHEN [特徴3] = 2809 THEN 28 /**/ WHEN [特徴3] = 2810 THEN 28 /**/ WHEN [特徴3] = 2811 THEN 28 /**/ WHEN [特徴3] = 2812 THEN 28 /**/ WHEN [特徴3] = 2813 THEN 28 /**/ WHEN [特徴3] = 2814 THEN 28 /**/ WHEN [特徴3] = 2815 THEN 28 /**/ WHEN [特徴3] = 2816 THEN 28 /**/ WHEN [特徴3] = 2817 THEN 28 /**/ WHEN [特徴3] = 2818 THEN 28 /**/ WHEN [特徴3] = 2819 THEN 28 /**/ WHEN [特徴3] = 2820 THEN 28 /**/ WHEN [特徴3] = 2821 THEN 28 /**/ WHEN [特徴3] = 2822 THEN 28 /**/ WHEN [特徴3] = 2901 THEN 29 /**/ WHEN [特徴3] = 2902 THEN 29 /**/ WHEN [特徴3] = 2903 THEN 29 ";
                tmp_sql = tmp_sql + " 			ELSE -1 ";
                tmp_sql = tmp_sql + " 		 END AS [おすすめピックアップピクト指定3] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [特徴3] = '' THEN -1 ";
                tmp_sql = tmp_sql + " 			ELSE [特徴3] ";
                tmp_sql = tmp_sql + " 		 END AS [特徴3] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [物件名公開] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 			WHEN [物件名公開] = 0 THEN 'false' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [物件名公開] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [部屋番号公開] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 			WHEN [部屋番号公開] = 0 THEN 'false' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [部屋番号公開] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [詳細住所公開] = 1 THEN 'true' ";
                tmp_sql = tmp_sql + " 			WHEN [詳細住所公開] = 0 THEN 'false' ";
                tmp_sql = tmp_sql + " 			ELSE 'false' ";
                tmp_sql = tmp_sql + " 		 END AS [詳細住所公開] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [会社間物件検索コピー許可] = 1 THEN '1' ";
                tmp_sql = tmp_sql + " 			WHEN [会社間物件検索コピー許可] = 2 THEN '2' ";
                tmp_sql = tmp_sql + " 			ELSE '2' ";
                tmp_sql = tmp_sql + " 		 END AS [会社間物件検索コピー許可] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 2 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 17 THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 22 THEN 4 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 33 THEN 5 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 38 THEN 6 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 49 THEN 7 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 54 THEN 8 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 65 THEN 9 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 70 THEN 10 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 81 THEN 11 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 86 THEN 12 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 97 THEN 13 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 102 THEN 14 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 113 THEN 15 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 118 THEN 16 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 129 THEN 17 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 134 THEN 18 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 145 THEN 19 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 161 THEN 20 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 166 THEN 21 ";
                tmp_sql = tmp_sql + " 			WHEN [図面パターン] = 177 THEN 22 ";
                tmp_sql = tmp_sql + " 			ELSE -1 ";
                tmp_sql = tmp_sql + " 		 END AS [図面パターン] ";
                tmp_sql = tmp_sql + " 		,[メインキャッチ1] ";
                tmp_sql = tmp_sql + " 		,[メインキャッチ2] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ1] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ2] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ3] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ4] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ5] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ6] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ7] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ8] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ9] ";
                tmp_sql = tmp_sql + " 		,[サブキャッチ10] ";
                tmp_sql = tmp_sql + " 		,[会社間用補足フリーコメント] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '00' THEN -1 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '01' THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '02' THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '03' THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '04' THEN 4 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '05' THEN 5 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '06' THEN 6 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '10' THEN 7 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '11' THEN 8 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '12' THEN 9 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '13' THEN 10 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '14' THEN 11 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '15' THEN 12 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '16' THEN 13 ";
                tmp_sql = tmp_sql + " 			WHEN [SUUMO内優先画像] = '17' THEN 14 ";
                tmp_sql = tmp_sql + " 			ELSE -1 ";
                tmp_sql = tmp_sql + " 		 END AS [SUUMO内優先画像] ";
                tmp_sql = tmp_sql + " 		,[貴社管理コード1] ";
                tmp_sql = tmp_sql + " 		,[貴社管理コード2] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN [見学予約機能を利用する] = 1 THEN '1' ";
                tmp_sql = tmp_sql + " 			WHEN [見学予約機能を利用する] = 2 THEN '0' ";
                tmp_sql = tmp_sql + " 			ELSE '0' ";
                tmp_sql = tmp_sql + " 		 END AS [見学予約機能を利用する] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 			,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 229 THEN d2_item ELSE '' END) AS [物件の特徴_ネット用キャッチ] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 230 THEN d2_item ELSE '' END) AS [物件の特徴_フリーコメント] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 55 THEN d2_item ELSE '' END) AS [敷金積増条件] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 56 THEN d2_item ELSE '' END) AS [敷金積増後総額] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 57 THEN d2_item ELSE '' END) AS [敷金積増後総額_単位] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 74 THEN d2_item ELSE '' END) AS [駐車場備考] ";
                tmp_sql = tmp_sql + " 			/*,'' AS [特優賃有無]*/ ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 313 THEN d2_item ELSE '' END) AS [特優賃入居者負担額_下限] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 314 THEN d2_item ELSE '' END) AS [特優賃入居者負担額_上限] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 315 THEN d2_item ELSE '' END) AS [特優賃料金変動区分] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 316 THEN d2_item ELSE '' END) AS [特優賃料金変動区分上昇率] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 317 THEN d2_item ELSE '' END) AS [特優賃家賃補助年数] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 318 THEN d2_item ELSE '' END) AS [特優賃補足] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 307 THEN d2_item ELSE '' END) AS [リフォーム箇所] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 308 THEN d2_item ELSE '' END) AS [リフォーム時期] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 309 THEN d2_item ELSE '' END) AS [リフォーム補足] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 179 THEN d2_item ELSE '' END) AS [設備環境1] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 180 THEN d2_item ELSE '' END) AS [設備環境1_距離] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 247 THEN d2_item ELSE '' END) AS [設備環境2] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 248 THEN d2_item ELSE '' END) AS [設備環境2_距離] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 181 THEN d2_item ELSE '' END) AS [周辺環境隣接1] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 182 THEN d2_item ELSE '' END) AS [周辺環境隣接1_単位] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 303 THEN d2_item ELSE '' END) AS [周辺環境隣接2] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 304 THEN d2_item ELSE '' END) AS [周辺環境隣接2_単位] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 183 THEN d2_item ELSE '' END) AS [設備環境1F] ";
                tmp_sql = tmp_sql + " 			,'' AS [セキュリティ] ";
                tmp_sql = tmp_sql + " 			,'' AS [室内間取] ";
                tmp_sql = tmp_sql + " 			,'' AS [家具・家電] ";
                tmp_sql = tmp_sql + " 			,'' AS [リフォーム] ";
                tmp_sql = tmp_sql + " 			,'' AS [リフォーム(内容)] ";
                tmp_sql = tmp_sql + " 			,'' AS [入居条件] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 13001 THEN d2_item ELSE '' END) AS [SUUMO以外の災害時住宅支援サイトへの掲載] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 321 THEN d2_item ELSE '' END) AS [おすすめピックアップピクト指定1] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 1321 THEN d2_item ELSE '' END) AS [特徴1] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 322 THEN d2_item ELSE '' END) AS [おすすめピックアップピクト指定2] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 1322 THEN d2_item ELSE '' END) AS [特徴2] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 323 THEN d2_item ELSE '' END) AS [おすすめピックアップピクト指定3] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 1323 THEN d2_item ELSE '' END) AS [特徴3] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 7 THEN d2_item ELSE '' END) AS [物件名公開] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 11 THEN d2_item ELSE '' END) AS [部屋番号公開] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 32 THEN d2_item ELSE '' END) AS [詳細住所公開] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 336 THEN d2_item ELSE '' END) AS [会社間物件検索コピー許可] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 300 THEN d2_item ELSE '' END) AS [図面パターン] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 195 THEN d2_item ELSE '' END) AS [メインキャッチ1] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 196 THEN d2_item ELSE '' END) AS [メインキャッチ2] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 197 THEN d2_item ELSE '' END) AS [サブキャッチ1] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 198 THEN d2_item ELSE '' END) AS [サブキャッチ2] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 199 THEN d2_item ELSE '' END) AS [サブキャッチ3] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 200 THEN d2_item ELSE '' END) AS [サブキャッチ4] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 201 THEN d2_item ELSE '' END) AS [サブキャッチ5] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 202 THEN d2_item ELSE '' END) AS [サブキャッチ6] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 203 THEN d2_item ELSE '' END) AS [サブキャッチ7] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 204 THEN d2_item ELSE '' END) AS [サブキャッチ8] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 205 THEN d2_item ELSE '' END) AS [サブキャッチ9] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 206 THEN d2_item ELSE '' END) AS [サブキャッチ10] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 207 THEN d2_item ELSE '' END) AS [会社間用補足フリーコメント] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 302 THEN d2_item ELSE '' END) AS [SUUMO内優先画像] ";
                tmp_sql = tmp_sql + " 			,'' AS [貴社管理コード1] ";
                tmp_sql = tmp_sql + " 			,'' AS [貴社管理コード2] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN d2_data_id = 337 THEN d2_item ELSE '' END) AS [見学予約機能を利用する] ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT * FROM hy_d2 WHERE d2_data_kbn = 30 ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		GROUP BY bk_no,hy_no ";
                tmp_sql = tmp_sql + " 	) AS VW ";
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

    #region ポータル連動部屋分類情報

    public class Rendo_hyrui_Repository
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

                sortstr = "[部屋分類No],[サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[部屋分類No],[部屋分類名],[サイトNo],[サイト用部屋分類No]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 crui_no AS [部屋分類No] ";
                tmp_sql = tmp_sql + " 			,crui_name AS [部屋分類名] ";
                tmp_sql = tmp_sql + " 			,10 AS [サイトNo] ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN wmp_crui_syoki = 0 THEN '' ";
                tmp_sql = tmp_sql + " 				ELSE CONVERT(VARCHAR,wmp_crui_syoki) ";
                tmp_sql = tmp_sql + " 			 END AS [サイト用部屋分類No] ";
                tmp_sql = tmp_sql + " 		FROM m_crui ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 crui_no ";
                tmp_sql = tmp_sql + " 			,crui_name ";
                tmp_sql = tmp_sql + " 			,20 AS site_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN homes_crui_syoki = 0 THEN '' ";
                tmp_sql = tmp_sql + " 				WHEN homes_crui_syoki = 3209 OR homes_crui_syoki = 3211 OR homes_crui_syoki = 3212 THEN '3299' ";
                tmp_sql = tmp_sql + " 				ELSE CONVERT(VARCHAR,homes_crui_syoki) ";
                tmp_sql = tmp_sql + " 			 END AS [HOMES部屋分類No] ";
                tmp_sql = tmp_sql + " 		FROM m_crui ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 crui_no ";
                tmp_sql = tmp_sql + " 			,crui_name ";
                tmp_sql = tmp_sql + " 			,30 AS site_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN athome_crui_syoki = 0 THEN '' ";
                tmp_sql = tmp_sql + " 				ELSE CONVERT(VARCHAR,athome_crui_syoki) ";
                tmp_sql = tmp_sql + " 			 END AS [athome部屋分類No] ";
                tmp_sql = tmp_sql + " 		FROM m_crui ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 crui_no ";
                tmp_sql = tmp_sql + " 			,crui_name ";
                tmp_sql = tmp_sql + " 			,40 AS site_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN yobi_si4 = 0 THEN '' ";
                tmp_sql = tmp_sql + " 				ELSE RIGHT('0' + CONVERT(VARCHAR,yobi_si4),2) ";
                tmp_sql = tmp_sql + " 			 END AS [SUUMO部屋分類No] ";
                tmp_sql = tmp_sql + " 		FROM m_crui ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 	WHERE [サイト用部屋分類No] <> '' ";
                tmp_sql = tmp_sql + " 	AND [部屋分類No] IN (SELECT DISTINCT crui_no FROM hy_mst) ";
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

    #region 部屋毎送信情報

    public class Rendo_hysosin_Repository
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

                sortstr = "[物件No],[部屋No],[ポータルサイトNo],[自社サービスNo],[選択No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[ポータルサイトNo],[自社サービスNo],[選択No],[項目値]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,site_no AS [ポータルサイトNo] ";
                tmp_sql = tmp_sql + " 		,site_jisyano AS [自社サービスNo] ";
                tmp_sql = tmp_sql + " 		,select_no AS [選択No] ";
                tmp_sql = tmp_sql + " 		,komok_value AS [項目値]	 ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*自社WEB sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,1 AS site_jisyano	/*くらさぽ*/ ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN sland_umu = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN sland_umu = 1 THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,2 AS site_jisyano	/*WMP*/ ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN wmp_umu = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN wmp_umu = 1 THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,3 AS site_jisyano	/*MilMil部屋タッチ*/ ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN tentobk_umu = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN tentobk_umu = 1 THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,4 AS site_jisyano	/*BtoBプラグイン*/ ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN gysite_umu = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN gysite_umu = 1 THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,5 AS site_jisyano	/*営業支援システム*/ ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN wmpsien_umu = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN wmpsien_umu = 1 THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,6 AS site_jisyano	/*不動産BB*/ ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN fudosanbb_umu = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN fudosanbb_umu = 1 THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		/*自社WEB end*/ ";
                tmp_sql = tmp_sql + " 		/*HOMES sta*/ ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,20 AS site_no ";
                tmp_sql = tmp_sql + " 			,0 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN homes_kbn = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN homes_kbn = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 				WHEN homes_kbn = 2 THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		/*HOMES end*/ ";
                tmp_sql = tmp_sql + " 		/*athome sta*/ ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,30 AS site_no ";
                tmp_sql = tmp_sql + " 			,0 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN athome_umu = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN athome_umu = 1 THEN 0 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		/*athome end*/ ";
                tmp_sql = tmp_sql + " 		/*SUUMO sta*/ ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,40 AS site_no ";
                tmp_sql = tmp_sql + " 			,0 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,1 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN suumo_kbn = 0 THEN -1 ";
                tmp_sql = tmp_sql + " 				WHEN suumo_kbn = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN suumo_kbn = 2 THEN 3 ";
                tmp_sql = tmp_sql + " 				WHEN suumo_kbn = 3 THEN 5 ";
                tmp_sql = tmp_sql + " 				WHEN suumo_kbn = 4 THEN 1 ";
                tmp_sql = tmp_sql + " 				ELSE -1 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,40 AS site_no ";
                tmp_sql = tmp_sql + " 			,0 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,2 AS select_no ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN suumo_kbn = 1 THEN 5 ";
                tmp_sql = tmp_sql + " 				WHEN suumo_kbn = 2 THEN 5 ";
                tmp_sql = tmp_sql + " 			 END AS komok_value ";
                tmp_sql = tmp_sql + " 		FROM hy_mst ";
                tmp_sql = tmp_sql + " 		WHERE suumo_kbn IN (1,2) ";
                tmp_sql = tmp_sql + " 		/*SUUMO end*/ ";
                tmp_sql = tmp_sql + " 	) AS VW ";
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

    #region BtoBグループ設定情報

    public class Rendo_BtoBgroup_Repository
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

                sortstr = "[物件No],[部屋No],[BtoBプラグイングループNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[BtoBプラグイングループNo],[グループ公開有無]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,1 AS [BtoBプラグイングループNo] ";
                tmp_sql = tmp_sql + " 		,groupa_umu AS [グループ公開有無] ";
                tmp_sql = tmp_sql + " 	FROM hy_mst ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no ";
                tmp_sql = tmp_sql + " 		,hy_no ";
                tmp_sql = tmp_sql + " 		,2 AS [グループNo] ";
                tmp_sql = tmp_sql + " 		,groupb_umu ";
                tmp_sql = tmp_sql + " 	FROM hy_mst ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no ";
                tmp_sql = tmp_sql + " 		,hy_no ";
                tmp_sql = tmp_sql + " 		,3 AS [グループNo] ";
                tmp_sql = tmp_sql + " 		,groupc_umu ";
                tmp_sql = tmp_sql + " 	FROM hy_mst ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no ";
                tmp_sql = tmp_sql + " 		,hy_no ";
                tmp_sql = tmp_sql + " 		,4 AS [グループNo] ";
                tmp_sql = tmp_sql + " 		,groupd_umu ";
                tmp_sql = tmp_sql + " 	FROM hy_mst ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no ";
                tmp_sql = tmp_sql + " 		,hy_no ";
                tmp_sql = tmp_sql + " 		,5 AS [グループNo] ";
                tmp_sql = tmp_sql + " 		,groupe_umu ";
                tmp_sql = tmp_sql + " 	FROM hy_mst ";
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

    #region 地図表示詳細設定情報

    public class Rendo_mapdisp_Repository
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

                sortstr = "[物件No],[部屋No],[外部サイトNo（FK）],[自社サイトNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[外部サイトNo（FK）],[自社サイトNo],[地図上に表示フラグ]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 VW.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,site_no AS [外部サイトNo（FK）] ";
                tmp_sql = tmp_sql + " 		,site_jisyano AS [自社サイトNo] ";
                tmp_sql = tmp_sql + " 		,mapdispflg AS [地図上に表示フラグ] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,1 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN sland_map_open = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN sland_map_open = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			 END AS mapdispflg ";
                tmp_sql = tmp_sql + " 		FROM bk_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,2 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			 END AS mapdispflg ";
                tmp_sql = tmp_sql + " 		FROM bk_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,3 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			 END AS mapdispflg ";
                tmp_sql = tmp_sql + " 		FROM bk_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,4 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			 END AS mapdispflg ";
                tmp_sql = tmp_sql + " 		FROM bk_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,5 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			 END AS mapdispflg ";
                tmp_sql = tmp_sql + " 		FROM bk_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,10 AS site_no ";
                tmp_sql = tmp_sql + " 			,6 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN jisyaweb_map_open = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			 END AS mapdispflg ";
                tmp_sql = tmp_sql + " 		FROM bk_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,30 AS site_no ";
                tmp_sql = tmp_sql + " 			,0 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN athome_map_open = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN athome_map_open = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			 END AS mapdispflg ";
                tmp_sql = tmp_sql + " 		FROM bk_mst ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,40 AS site_no ";
                tmp_sql = tmp_sql + " 			,0 AS site_jisyano ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN suumo_map_open = 0 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN suumo_map_open = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			 END AS mapdispflg ";
                tmp_sql = tmp_sql + " 		FROM bk_mst ";
                tmp_sql = tmp_sql + " 	) AS VW ";
                tmp_sql = tmp_sql + " 	LEFT JOIN hy_mst AS HY ";
                tmp_sql = tmp_sql + " 	ON VW.bk_no = HY.bk_no ";
                tmp_sql = tmp_sql + " 	WHERE HY.hy_no IS NOT NULL ";
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