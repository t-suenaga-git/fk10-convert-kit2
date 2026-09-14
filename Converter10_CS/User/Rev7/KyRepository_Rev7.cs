using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using Converter10.Njc.Common;

namespace Converter10.Njc.Repository
{

    #region 契約基本情報

    public class Ky_kosinkai_Repository
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

                sortstr = "[物件No],[部屋No],[契約No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[初回契約日],[契約状況(ステータス)],[キャンセル理由],[ステータス変更日],[仲介業者No],[取引主任者No],[手付預り額①],[手付預り日①],[手付預り備考①],[解約フラグ],[サービス分類],[会計グループ分類],[解約受付区分],[解約受付月数],[解約受付日数],[解約受付日にち],[仲介担当者名],[仲介担当者名SJIS],[契約開始日],[契約終了日]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		/*V7の契約日は無くても登録可能なので無い場合は契約開始日を設定*/ ";
                tmp_sql = tmp_sql + " 		/*また仮契約の場合は仮契約日(V7必須項目)を設定する*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 AND ISNULL(ky_date,'') <> '' THEN ky_date ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 AND ISNULL(ky_date,'') = '' THEN ky_start_ymd ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 1 THEN kari_ymd ";
                tmp_sql = tmp_sql + " 		 END AS [初回契約日]	 ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 THEN 2	/*契約確定*/ ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 1 THEN 1	/*仮契約*/ ";
                tmp_sql = tmp_sql + " 		 END AS [契約状況(ステータス)] ";
                tmp_sql = tmp_sql + " 		,'' AS [キャンセル理由] ";
                tmp_sql = tmp_sql + " 		/*V7の契約日は無くても登録可能なので無い場合は契約開始日を設定*/ ";
                tmp_sql = tmp_sql + " 		/*また仮契約の場合は仮契約日(V7必須項目)を設定する*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 AND ISNULL(ky_date,'') <> '' THEN ky_date ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 AND ISNULL(ky_date,'') = '' THEN ky_start_ymd ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 1 THEN NULL ";
                tmp_sql = tmp_sql + " 		 END AS [ステータス変更日] ";
                tmp_sql = tmp_sql + " 		,gy_no AS [仲介業者No] ";
                tmp_sql = tmp_sql + " 		,tanto_no_syu AS [取引主任者No] ";
                tmp_sql = tmp_sql + " 		,tetuke AS [手付預り額①] ";
                tmp_sql = tmp_sql + " 		,tetuke_date1 AS [手付預り日①] ";
                tmp_sql = tmp_sql + " 		,'' AS [手付預り備考①] ";
                tmp_sql = tmp_sql + " 		/*解約フラグは10のkydata_kai.ky_recnoが設定されるためとりあえず999を設定しておく*/ ";
                tmp_sql = tmp_sql + " 		/*kydata_kai.ky_recnoはkydata_kihon.ky_recno + 1が設定されるのでkydata_kihon作成後に更新する*/ ";
                tmp_sql = tmp_sql + " 		,kaiflg AS [解約フラグ] ";
                tmp_sql = tmp_sql + " 		,-1 AS [サービス分類] ";
                tmp_sql = tmp_sql + " 		,'' AS [会計グループ分類] ";
                tmp_sql = tmp_sql + " 		,1 AS [解約受付区分] ";
                tmp_sql = tmp_sql + " 		,3 AS [解約受付月数] ";
                tmp_sql = tmp_sql + " 		,'' AS [解約受付日数] ";
                tmp_sql = tmp_sql + " 		,'' AS [解約受付日にち] ";
                tmp_sql = tmp_sql + " 		,'' AS [仲介担当者名] ";
                tmp_sql = tmp_sql + " 		,'' AS [仲介担当者名SJIS] ";
                tmp_sql = tmp_sql + " 		/*2016.04.06 仮契約情報の移行制御処理を追加 add sta*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 THEN ky_start_ymd ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 1 THEN kari_start_ymd ";
                tmp_sql = tmp_sql + " 		 END AS [契約開始日] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 THEN ky_end_ymd ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 1 THEN kari_end_ymd ";
                tmp_sql = tmp_sql + " 		 END AS [契約終了日]	 ";
                tmp_sql = tmp_sql + " 		/*2016.04.06 仮契約情報の移行制御処理を追加 add end*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*最古の契約情報へ解約されている契約を結合し解約フラグを取得したVIEWを作成 sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 KY.* ";
                tmp_sql = tmp_sql + " 			,ISNULL(KAI.ko_no,0) AS kaiflg ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*契約情報へ古い順に連番を作成 → 最古の契約を抽出 sta*/ ";
                tmp_sql = tmp_sql + " 			SELECT * FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY bk_no,hy_no,ky_no,ko_no) AS [抽出用] ";
                tmp_sql = tmp_sql + " 					,* ";
                tmp_sql = tmp_sql + " 				FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 			) AS VW1 ";
                tmp_sql = tmp_sql + " 			WHERE [抽出用] = 1 ";
                tmp_sql = tmp_sql + " 			/*契約情報へ古い順に連番を作成 → 最古の契約を抽出 end*/ ";
                tmp_sql = tmp_sql + " 		) AS KY ";
                tmp_sql = tmp_sql + " 		LEFT JOIN ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				/*最古の契約情報への結合用に解約情報のみ抽出 sta*/ ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_no ";
                tmp_sql = tmp_sql + " 					,ky_no ";
                tmp_sql = tmp_sql + " 					,ko_no ";
                tmp_sql = tmp_sql + " 				FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 				WHERE ko_no = 999 ";
                tmp_sql = tmp_sql + " 				/*最古の契約情報への結合用に解約情報のみ抽出 end*/ ";
                tmp_sql = tmp_sql + " 			) AS KAI ";
                tmp_sql = tmp_sql + " 		ON  KY.bk_no = KAI.bk_no ";
                tmp_sql = tmp_sql + " 		AND KY.hy_no = KAI.hy_no ";
                tmp_sql = tmp_sql + " 		AND KY.ky_no = KAI.ky_no ";
                tmp_sql = tmp_sql + " 		/*最古の契約情報へ解約されている契約を結合し解約フラグを取得したVIEWを作成 end*/ ";
                tmp_sql = tmp_sql + " 	) AS KYTOTAL ";
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

    #region 契約履歴情報

    public class Ky_kosinkai_rireki_Repository
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

                // -------------------------------------------------------
                // VIEW作成前に必要な情報を格納した仮テーブルを作成する
                // -------------------------------------------------------

                string tmp_sql = "";

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[更新No],[改定No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[更新No],[改定No],[契約番号],[契約日],[契約開始日],[契約終了日],[条件変更日],[契約更新通知日],[通知書印刷日],[管理業者No],[賃貸保証業者No],[賃貸保証内容],[賃貸保証顧客番号],[契約分類名],[使用目的],[契約更新業務有無],[家賃入金区分],[口座振替開始日],[口座振替開始待ちフラグ],[家賃入金口座区分],[契約一時金入金口座No],[毎月分入金口座No],[翌月分受取り有無],[備考],[当月分の差額受取り有無],[次回契約開始日],[次回契約終了日],[更新時日割り有無],[送金控除ルール適用有無],[送金予定日使用フラグ],[送金予定日],[部屋固定管理手数料フラグ],[部屋固定管理手数料額],[次回更新設定区分],[請求データ作成開始日],[請求データ作成終了日],[備考2(基本)],[備考(保険)],[入金締め日],[家賃持参先],[保険期間],[保険額],[身元引受人名],[身元引受人住所],[身元引受人連絡先],[保険料入金日],[更新完了日],[更新完了通知日],[更新通知日],[催促実施日],[書類返送受取日],[毎月分請求有無],[身元引受人名SJIS],[口座振替日],[翌月分受取り月],[契約担当者No],[請求担当者No],[配分割合_元付],[配分割合_客付],[客付会社の手数料額],[客付会社の手数料税区分],[客付会社の手数料税額]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	/********************************************************************************************************** ";
                tmp_sql = tmp_sql + " 	契約履歴情報抽出クエリに関して ";
                tmp_sql = tmp_sql + " 	①V7と10では履歴を管理するNoの仕様に差異があるためそのまま移行できない ";
                tmp_sql = tmp_sql + " 		V7  …  履歴管理はko_no,ko_kbnで行われている ";
                tmp_sql = tmp_sql + " 				・更新/改定に関わらずko_noはカウントアップ ";
                tmp_sql = tmp_sql + " 				・更新の場合は 1、改定の場合は 2がko_kbnへセットされる ";
                tmp_sql = tmp_sql + " 		10  …  履歴管理はko_no,henko_noで行われている ";
                tmp_sql = tmp_sql + " 				・更新されるとko_noはカウントアップ ";
                tmp_sql = tmp_sql + " 				  →  henko_noは1に初期化 ";
                tmp_sql = tmp_sql + " 				・改定されるとhenko_noはカウントアップ ";
                tmp_sql = tmp_sql + " 				  →  ko_noはそのまま ";
                tmp_sql = tmp_sql + " 	  →  ・革命10に合わせた更新Noと改定Noを作成した仮テーブルをあらかじめ作成しておき、 ";
                tmp_sql = tmp_sql + " 			V7のテーブルと結合してNoを取得する ";
                tmp_sql = tmp_sql + " 			  ※仮テーブル名「CVTBL_契約履歴情報」 ";
                tmp_sql = tmp_sql + " 	           ";
                tmp_sql = tmp_sql + " 	②契約日、契約開始日、契約終了日の仕様に差異があるためそのまま移行できない ";
                tmp_sql = tmp_sql + " 		V7  …  改定すると ";
                tmp_sql = tmp_sql + " 				・条件変更日が変更前の契約終了日になる ";
                tmp_sql = tmp_sql + " 				・条件変更時に変更している契約情報の契約終了日は変更可能 ";
                tmp_sql = tmp_sql + " 		10  …  改定すると ";
                tmp_sql = tmp_sql + " 				・更新するまで最初の契約開始/終了日が保持される (変更不可) ";
                tmp_sql = tmp_sql + " 	  →  ①で作成した更新Noで括り ";
                tmp_sql = tmp_sql + " 			・V7の最も古い契約日を契約日へ(無い場合は契約開始日) ";
                tmp_sql = tmp_sql + " 			・V7の最も古い契約開始日を契約開始日へ ";
                tmp_sql = tmp_sql + " 			・V7の最も新しい契約終了日を契約終了日へ ";
                tmp_sql = tmp_sql + " 	**********************************************************************************************************/ ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 KYK.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,KYK.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,KYK.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,KYK.ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		/*,KYK.ko_no AS [更新NoV7] ";
                tmp_sql = tmp_sql + " 		,KYK.ko_kbn AS [変更区分V7]*/ ";
                tmp_sql = tmp_sql + " 		,KYK.kono10 AS [更新No] ";
                tmp_sql = tmp_sql + " 		,KYK.henkono10 AS [改定No] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約番号] ";
                tmp_sql = tmp_sql + " 		/*仮契約の場合は仮契約日以外必須項目ではないため移行仕様を考慮すること*/ ";
                tmp_sql = tmp_sql + " 		/*仮契約の移行は現時点でそのまま移行する形にしておく*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 THEN KYYMD.OLD_ky_date ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 1 THEN yotei_ymd ";
                tmp_sql = tmp_sql + " 		 END AS [契約日] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 THEN KYSTARTYMD.OLD_ky_start_ymd ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 1 THEN kari_start_ymd ";
                tmp_sql = tmp_sql + " 		 END AS [契約開始日] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 0 THEN KYENDYMD.NEW_ky_end_ymd ";
                tmp_sql = tmp_sql + " 			WHEN kari_flg = 1 THEN kari_end_ymd ";
                tmp_sql = tmp_sql + " 		 END AS [契約終了日] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN KYK.henkono10 >= 2 THEN KYK.ky_start_ymd ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [条件変更日] ";
                tmp_sql = tmp_sql + " 		,tuti_end_ymd AS [契約更新通知日] ";
                tmp_sql = tmp_sql + " 		,'' AS [通知書印刷日] ";
                tmp_sql = tmp_sql + " 		,kanri_gy_no AS [管理業者No] ";
                tmp_sql = tmp_sql + " 		,hkngy_no AS [賃貸保証業者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [賃貸保証内容] ";
                tmp_sql = tmp_sql + " 		,'' AS [賃貸保証顧客番号] ";
                tmp_sql = tmp_sql + " 		/*2016.04.06 契約分類を紐付データを元に移行 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,keirui_no AS [契約分類No]*/ ";
                tmp_sql = tmp_sql + " 		,keirui_name AS [契約分類名] ";
                tmp_sql = tmp_sql + " 		/*2016.04.06 契約分類を紐付データを元に移行 chg end*/ ";
                tmp_sql = tmp_sql + " 		,mokuteki AS [使用目的] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN nx_kosin_kbn = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN nx_kosin_kbn = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 		 END AS [契約更新業務有無] ";
                tmp_sql = tmp_sql + " 		,nkbn_name AS [家賃入金区分] ";
                tmp_sql = tmp_sql + " 		,fkae_startym AS [口座振替開始日] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替開始待ちフラグ] ";
                tmp_sql = tmp_sql + " 		,1 AS [家賃入金口座区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約一時金入金口座No] ";
                tmp_sql = tmp_sql + " 		,fkom_no AS [毎月分入金口座No] ";
                tmp_sql = tmp_sql + " 		,rai_umu AS [翌月分受取り有無] ";
                tmp_sql = tmp_sql + " 		,KYK.biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [当月分の差額受取り有無] ";
                tmp_sql = tmp_sql + " 		/*条件変更を行うと条件変更日が契約終了日となるため次回更新契約開始日が変更される*/ ";
                tmp_sql = tmp_sql + " 		/*10では最初の契約開始・終了日が固定のため次回更新契約開始日も固定となる*/ ";
                tmp_sql = tmp_sql + " 		/*よって抽出した際の契約終了日 +1 日を次回更新契約開始日とする*/ ";
                tmp_sql = tmp_sql + " 		/*,ky_nx_start_ymd AS [次回契約開始日]*/ ";
                tmp_sql = tmp_sql + " 		,DATEADD(DAY,1,KYENDYMD.NEW_ky_end_ymd) AS [次回契約開始日] ";
                tmp_sql = tmp_sql + " 		/*次回契約終了日が不正になる場合があるため取得方法を変更 -chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*/*次回更新契約終了日はV7、10どちらも手動設定可能なのでそのまま移行する*/*/ ";
                tmp_sql = tmp_sql + " 		/*/*,ky_nx_end_ymd AS [次回契約終了日]*/*/ ";
                tmp_sql = tmp_sql + " 		/*↑の取得方法では10の仕様に合った状態で抽出できないので修正*/ ";
                tmp_sql = tmp_sql + " 		,NEWNX_ky_end_ymd AS [次回契約終了日] ";
                tmp_sql = tmp_sql + " 		/*次回契約終了日が不正になる場合があるため取得方法を変更 -chg end*/ ";
                tmp_sql = tmp_sql + " 		,hiwari_nx_flg AS [更新時日割り有無] ";
                tmp_sql = tmp_sql + " 		,1 AS [送金控除ルール適用有無] ";
                tmp_sql = tmp_sql + " 		,0 AS [送金予定日使用フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [送金予定日] ";
                tmp_sql = tmp_sql + " 		,0 AS [部屋固定管理手数料フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [部屋固定管理手数料額] ";
                tmp_sql = tmp_sql + " 		,0 AS [次回更新設定区分] ";
                tmp_sql = tmp_sql + " 		/*条件変更された場合、条件変更日が請求データ作成開始日へ設定される*/ ";
                tmp_sql = tmp_sql + " 		/*条件変更された場合でも契約終了日は固定のため契約終了日を移行する*/ ";
                tmp_sql = tmp_sql + " 		/*,teki_start_ymd AS [請求データ作成開始日]*/ ";
                tmp_sql = tmp_sql + " 		/*,teki_end_ymd AS [請求データ作成終了日]*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN KYK.henkono10 >= 2 THEN KYK.ky_start_ymd ";
                tmp_sql = tmp_sql + " 			ELSE KYSTARTYMD.OLD_ky_start_ymd ";
                tmp_sql = tmp_sql + " 		 END AS [請求データ作成開始日]	 ";
                tmp_sql = tmp_sql + " 		,KYENDYMD.NEW_ky_end_ymd AS [請求データ作成終了日] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考2(基本)] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考(保険)] ";
                tmp_sql = tmp_sql + " 		/*2016.04.06 入金締/家賃持参先を物件情報から取得 -chg sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,'' AS [入金締め日] ";
                tmp_sql = tmp_sql + " 		,'' AS [家賃持参先] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		,BK.sh_kigen AS [入金締め日] ";
                tmp_sql = tmp_sql + " 		,BK.jisansaki AS [家賃持参先]	 ";
                tmp_sql = tmp_sql + " 		/*2016.04.06 入金締/家賃持参先を物件情報から取得 -chg end*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [保険期間] ";
                tmp_sql = tmp_sql + " 		,'' AS [保険額] ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約履歴情報 del sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,'' AS [確認事項(契約)責任者] ";
                tmp_sql = tmp_sql + " 		,'' AS [確認事項(契約)内容確認日] ";
                tmp_sql = tmp_sql + " 		,'' AS [確認事項(契約)印刷時] ";
                tmp_sql = tmp_sql + " 		,'' AS [確認事項(解約)責任者] ";
                tmp_sql = tmp_sql + " 		,'' AS [確認事項(解約)内容確認日] ";
                tmp_sql = tmp_sql + " 		,'' AS [確認事項(解約)印刷時] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約履歴情報 del end*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [身元引受人名] ";
                tmp_sql = tmp_sql + " 		,'' AS [身元引受人住所] ";
                tmp_sql = tmp_sql + " 		,'' AS [身元引受人連絡先] ";
                tmp_sql = tmp_sql + " 		,'' AS [保険料入金日] ";
                tmp_sql = tmp_sql + " 		,'' AS [更新完了日] ";
                tmp_sql = tmp_sql + " 		,'' AS [更新完了通知日] ";
                tmp_sql = tmp_sql + " 		,'' AS [更新通知日] ";
                tmp_sql = tmp_sql + " 		,'' AS [催促実施日] ";
                tmp_sql = tmp_sql + " 		,'' AS [書類返送受取日] ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約履歴情報 del sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,'' AS [更新時の解約検討連絡有無] ";
                tmp_sql = tmp_sql + " 		,'' AS [更新時の解約検討連絡受付日] ";
                tmp_sql = tmp_sql + " 		,'' AS [更新時の解約検討連絡受付担当者] ";
                tmp_sql = tmp_sql + " 		,'' AS [更新時の解約検討連絡備考] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約履歴情報 del end*/ ";
                tmp_sql = tmp_sql + " 		,1 AS [毎月分請求有無] ";
                tmp_sql = tmp_sql + " 		,'' AS [身元引受人名SJIS] ";
                tmp_sql = tmp_sql + " 		,'' AS [口座振替日] ";
                tmp_sql = tmp_sql + " 		,'' AS [翌月分受取り月] ";
                tmp_sql = tmp_sql + " 		,tanto_no AS [契約担当者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [請求担当者No] ";
                tmp_sql = tmp_sql + " 		/*20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) add sta*/ ";
                tmp_sql = tmp_sql + " 		,NULL AS [配分割合_元付] ";
                tmp_sql = tmp_sql + " 		,NULL AS [配分割合_客付] ";
                tmp_sql = tmp_sql + " 		,NULL AS [客付会社の手数料額] ";
                tmp_sql = tmp_sql + " 		,NULL AS [客付会社の手数料税区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [客付会社の手数料税額]	 ";
                tmp_sql = tmp_sql + " 		/*20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) add end*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*V7テーブルへ10の更新No、改定Noを結合したテーブルを作成 sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 		LEFT JOIN CVTBL_契約情報 AS KYTMP ";
                tmp_sql = tmp_sql + " 		ON  KYK.bk_no = KYTMP.tmp_bkno ";
                tmp_sql = tmp_sql + " 		AND KYK.hy_no = KYTMP.tmp_hyno ";
                tmp_sql = tmp_sql + " 		AND KYK.ky_no = KYTMP.tmp_kyno ";
                tmp_sql = tmp_sql + " 		AND KYK.ko_no = KYTMP.tmp_kono ";
                tmp_sql = tmp_sql + " 		WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 		/*V7テーブルへ10の更新No、改定Noを結合したテーブルを作成 end*/ ";
                tmp_sql = tmp_sql + " 	) AS KYK ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*最も古い契約日を取得 (契約日が存在しない場合は契約開始日を設定) sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 * ";
                tmp_sql = tmp_sql + " 				,ky_date AS OLD_ky_date ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,kono10 ORDER BY henkono10) AS [最古契約日取得用] ";
                tmp_sql = tmp_sql + " 					,* ";
                tmp_sql = tmp_sql + " 				FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_no ";
                tmp_sql = tmp_sql + " 						,ky_no ";
                tmp_sql = tmp_sql + " 						,ko_no ";
                tmp_sql = tmp_sql + " 						,ko_kbn ";
                tmp_sql = tmp_sql + " 						,kono10 ";
                tmp_sql = tmp_sql + " 						,henkono10 ";
                tmp_sql = tmp_sql + " 						,CASE ";
                tmp_sql = tmp_sql + " 							WHEN ISNULL(ky_date,'') <> '' THEN ky_date ";
                tmp_sql = tmp_sql + " 							WHEN ISNULL(ky_date,'') = '' THEN ky_start_ymd ";
                tmp_sql = tmp_sql + " 						 END AS ky_date ";
                tmp_sql = tmp_sql + " 						/*,ky_start_ymd*/ ";
                tmp_sql = tmp_sql + " 						/*,ky_end_ymd*/ ";
                tmp_sql = tmp_sql + " 					FROM ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 					LEFT JOIN CVTBL_契約情報 AS KYTMP ";
                tmp_sql = tmp_sql + " 					ON  KYK.bk_no = KYTMP.tmp_bkno ";
                tmp_sql = tmp_sql + " 					AND KYK.hy_no = KYTMP.tmp_hyno ";
                tmp_sql = tmp_sql + " 					AND KYK.ky_no = KYTMP.tmp_kyno ";
                tmp_sql = tmp_sql + " 					AND KYK.ko_no = KYTMP.tmp_kono ";
                tmp_sql = tmp_sql + " 					WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 				) AS VW1 ";
                tmp_sql = tmp_sql + " 			) AS VW2 ";
                tmp_sql = tmp_sql + " 			WHERE [最古契約日取得用] = 1 ";
                tmp_sql = tmp_sql + " 		) AS VW3 ";
                tmp_sql = tmp_sql + " 		/*最も古い契約日を取得 (契約日が存在しない場合は契約開始日を設定) end*/ ";
                tmp_sql = tmp_sql + " 	) AS KYYMD ";
                tmp_sql = tmp_sql + " 	ON  KYK.bk_no = KYYMD.bk_no ";
                tmp_sql = tmp_sql + " 	AND KYK.hy_no = KYYMD.hy_no ";
                tmp_sql = tmp_sql + " 	AND KYK.ky_no = KYYMD.ky_no ";
                tmp_sql = tmp_sql + " 	AND KYK.kono10 = KYYMD.kono10 ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*最も古い契約開始日を取得 sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 * ";
                tmp_sql = tmp_sql + " 				,ky_start_ymd AS OLD_ky_start_ymd ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,kono10 ORDER BY henkono10) AS [最古契約開始日取得用] ";
                tmp_sql = tmp_sql + " 					,* ";
                tmp_sql = tmp_sql + " 				FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_no ";
                tmp_sql = tmp_sql + " 						,ky_no ";
                tmp_sql = tmp_sql + " 						,ko_no ";
                tmp_sql = tmp_sql + " 						,ko_kbn ";
                tmp_sql = tmp_sql + " 						,kono10 ";
                tmp_sql = tmp_sql + " 						,henkono10 ";
                tmp_sql = tmp_sql + " 						/*,ky_date*/ ";
                tmp_sql = tmp_sql + " 						,ky_start_ymd ";
                tmp_sql = tmp_sql + " 						/*,ky_end_ymd*/ ";
                tmp_sql = tmp_sql + " 					FROM ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 					LEFT JOIN CVTBL_契約情報 AS KYTMP ";
                tmp_sql = tmp_sql + " 					ON  KYK.bk_no = KYTMP.tmp_bkno ";
                tmp_sql = tmp_sql + " 					AND KYK.hy_no = KYTMP.tmp_hyno ";
                tmp_sql = tmp_sql + " 					AND KYK.ky_no = KYTMP.tmp_kyno ";
                tmp_sql = tmp_sql + " 					AND KYK.ko_no = KYTMP.tmp_kono ";
                tmp_sql = tmp_sql + " 					WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 				) AS VW1 ";
                tmp_sql = tmp_sql + " 			) AS VW2 ";
                tmp_sql = tmp_sql + " 			WHERE [最古契約開始日取得用] = 1 ";
                tmp_sql = tmp_sql + " 		) AS VW3 ";
                tmp_sql = tmp_sql + " 		/*最も古い契約開始日を取得 end*/ ";
                tmp_sql = tmp_sql + " 	) AS KYSTARTYMD ";
                tmp_sql = tmp_sql + " 	ON  KYK.bk_no = KYSTARTYMD.bk_no ";
                tmp_sql = tmp_sql + " 	AND KYK.hy_no = KYSTARTYMD.hy_no ";
                tmp_sql = tmp_sql + " 	AND KYK.ky_no = KYSTARTYMD.ky_no ";
                tmp_sql = tmp_sql + " 	AND KYK.kono10 = KYSTARTYMD.kono10 ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*最も新しい契約終了日を取得 sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 * ";
                tmp_sql = tmp_sql + " 				,ky_end_ymd AS NEW_ky_end_ymd ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,kono10 ORDER BY henkono10 DESC) AS [最新契約終了日取得用] ";
                tmp_sql = tmp_sql + " 					,* ";
                tmp_sql = tmp_sql + " 				FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_no ";
                tmp_sql = tmp_sql + " 						,ky_no ";
                tmp_sql = tmp_sql + " 						,ko_no ";
                tmp_sql = tmp_sql + " 						,ko_kbn ";
                tmp_sql = tmp_sql + " 						,kono10 ";
                tmp_sql = tmp_sql + " 						,henkono10 ";
                tmp_sql = tmp_sql + " 						/*,ky_date*/ ";
                tmp_sql = tmp_sql + " 						/*,ky_start_ymd*/ ";
                tmp_sql = tmp_sql + " 						,ky_end_ymd ";
                tmp_sql = tmp_sql + " 					FROM ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 					LEFT JOIN CVTBL_契約情報 AS KYTMP ";
                tmp_sql = tmp_sql + " 					ON  KYK.bk_no = KYTMP.tmp_bkno ";
                tmp_sql = tmp_sql + " 					AND KYK.hy_no = KYTMP.tmp_hyno ";
                tmp_sql = tmp_sql + " 					AND KYK.ky_no = KYTMP.tmp_kyno ";
                tmp_sql = tmp_sql + " 					AND KYK.ko_no = KYTMP.tmp_kono ";
                tmp_sql = tmp_sql + " 					WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 				) AS VW1 ";
                tmp_sql = tmp_sql + " 			) AS VW2 ";
                tmp_sql = tmp_sql + " 			WHERE [最新契約終了日取得用] = 1 ";
                tmp_sql = tmp_sql + " 		) AS VW3 ";
                tmp_sql = tmp_sql + " 		/*最も新しい契約終了日を取得 end*/ ";
                tmp_sql = tmp_sql + " 	) AS KYENDYMD ";
                tmp_sql = tmp_sql + " 	ON  KYK.bk_no = KYENDYMD.bk_no ";
                tmp_sql = tmp_sql + " 	AND KYK.hy_no = KYENDYMD.hy_no ";
                tmp_sql = tmp_sql + " 	AND KYK.ky_no = KYENDYMD.ky_no ";
                tmp_sql = tmp_sql + " 	AND KYK.kono10 = KYENDYMD.kono10 ";
                tmp_sql = tmp_sql + " 	/*次回契約終了日が不正になる場合があるため取得方法を変更 -add sta*/ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*最も新しい次回契約終了日を取得 sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 * ";
                tmp_sql = tmp_sql + " 				,ky_nx_end_ymd AS NEWNX_ky_end_ymd ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,kono10 ORDER BY henkono10 DESC) AS [最新次回契約終了日取得用] ";
                tmp_sql = tmp_sql + " 					,* ";
                tmp_sql = tmp_sql + " 				FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 bk_no ";
                tmp_sql = tmp_sql + " 						,hy_no ";
                tmp_sql = tmp_sql + " 						,ky_no ";
                tmp_sql = tmp_sql + " 						,ko_no ";
                tmp_sql = tmp_sql + " 						,ko_kbn ";
                tmp_sql = tmp_sql + " 						,kono10 ";
                tmp_sql = tmp_sql + " 						,henkono10 ";
                tmp_sql = tmp_sql + " 						/*,ky_date*/ ";
                tmp_sql = tmp_sql + " 						/*,ky_start_ymd*/ ";
                tmp_sql = tmp_sql + " 						,ky_nx_end_ymd ";
                tmp_sql = tmp_sql + " 					FROM ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 					LEFT JOIN CVTBL_契約情報 AS KYTMP ";
                tmp_sql = tmp_sql + " 					ON  KYK.bk_no = KYTMP.tmp_bkno ";
                tmp_sql = tmp_sql + " 					AND KYK.hy_no = KYTMP.tmp_hyno ";
                tmp_sql = tmp_sql + " 					AND KYK.ky_no = KYTMP.tmp_kyno ";
                tmp_sql = tmp_sql + " 					AND KYK.ko_no = KYTMP.tmp_kono ";
                tmp_sql = tmp_sql + " 					WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 				) AS VW1 ";
                tmp_sql = tmp_sql + " 			) AS VW2 ";
                tmp_sql = tmp_sql + " 			WHERE [最新次回契約終了日取得用] = 1 ";
                tmp_sql = tmp_sql + " 		) AS VW3 ";
                tmp_sql = tmp_sql + " 		/*最も新しい次回契約終了日を取得 end*/ ";
                tmp_sql = tmp_sql + " 	) AS KYNXENDYMD ";
                tmp_sql = tmp_sql + " 	ON  KYK.bk_no = KYNXENDYMD.bk_no ";
                tmp_sql = tmp_sql + " 	AND KYK.hy_no = KYNXENDYMD.hy_no ";
                tmp_sql = tmp_sql + " 	AND KYK.ky_no = KYNXENDYMD.ky_no ";
                tmp_sql = tmp_sql + " 	AND KYK.kono10 = KYNXENDYMD.kono10 ";
                tmp_sql = tmp_sql + " 	/*次回契約終了日が不正になる場合があるため取得方法を変更 -add end*/ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkbn AS MNKBN ON KYK.yatin_nkbn = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " 	/*2016.04.06 入金締/家賃持参先を物件情報から取得 -add*/ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN (SELECT bk_no,sh_kigen,jisansaki FROM bk_mst) AS BK ON KYK.bk_no = BK.bk_no ";
                tmp_sql = tmp_sql + " 	/*2016.04.06 契約分類を紐付データを元に移行 add*/ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_keirui AS KRUI ON KYK.keirui_no = KRUI.keirui_no ";
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

    #region 契約契約者情報

    public class Ky_kosinkai_kys_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[並び順No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[並び順No],[契約者No],[入居フラグ]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 KYSTOTAL.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,KYSTOTAL.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,KYSTOTAL.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,KYSTOTAL.ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,KYSTOTAL.sort_no AS [並び順No] ";
                tmp_sql = tmp_sql + " 		,KYSTOTAL.kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			/*間柄が「本人」に設定されていない場合でも契約者が入居者(家族)データに登録されている場合があるため名称一致で入居フラグを立てる*/ ";
                tmp_sql = tmp_sql + " 			/*WHEN RTRIM(LTRIM(KYK.kazoku_name)) = RTRIM(LTRIM(KYS.kys_name)) AND RTRIM(LTRIM(KYK.aida)) = '本人' THEN 1*/ ";
                tmp_sql = tmp_sql + " 			WHEN RTRIM(LTRIM(kazoku_name)) = RTRIM(LTRIM(kys_name)) THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE 0 ";
                tmp_sql = tmp_sql + " 		 END AS [入居フラグ] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 KYKYSTOTAL.bk_no ";
                tmp_sql = tmp_sql + " 			,KYKYSTOTAL.hy_no ";
                tmp_sql = tmp_sql + " 			,KYKYSTOTAL.ky_no ";
                tmp_sql = tmp_sql + " 			,ko_no ";
                tmp_sql = tmp_sql + " 			,sort_no ";
                tmp_sql = tmp_sql + " 			,kys_no ";
                tmp_sql = tmp_sql + " 			,kys_name ";
                tmp_sql = tmp_sql + " 			,kazoku_no ";
                tmp_sql = tmp_sql + " 			,kazoku_name ";
                tmp_sql = tmp_sql + " 			,aida ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*本契約の契約者を取得 sta*/ ";
                tmp_sql = tmp_sql + " 			SELECT KYKYS.*,kys_name FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT bk_no,hy_no,ky_no,ko_no,1 AS sort_no,kys_no FROM ky_kosinkai	WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 				UNION ";
                tmp_sql = tmp_sql + " 				SELECT bk_no,hy_no,ky_no,ko_no,2 AS sort_no,kys_no2 FROM ky_kosinkai WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 				UNION ";
                tmp_sql = tmp_sql + " 				SELECT bk_no,hy_no,ky_no,ko_no,3 AS sort_no,kys_no3 FROM ky_kosinkai WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 			) AS KYKYS ";
                tmp_sql = tmp_sql + " 			LEFT JOIN kys_mst AS KYSMST ON KYKYS.kys_no = KYSMST.kys_no ";
                tmp_sql = tmp_sql + " 			/*本契約の契約者を取得 end*/ ";
                tmp_sql = tmp_sql + " 			UNION ";
                tmp_sql = tmp_sql + " 			/*仮契約の契約者を取得 sta*/ ";
                tmp_sql = tmp_sql + " 			SELECT * FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no,hy_no,ky_no,ko_no,1 AS sort_no ";
                tmp_sql = tmp_sql + " 					,ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 						)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 					,kari_name ";
                tmp_sql = tmp_sql + " 				FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 				WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 				UNION ";
                tmp_sql = tmp_sql + " 				SELECT bk_no,hy_no,ky_no,ko_no,2 AS sort_no,NULL AS kys_no,'' AS kys_name FROM ky_kosinkai WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 				UNION ";
                tmp_sql = tmp_sql + " 				SELECT bk_no,hy_no,ky_no,ko_no,3 AS sort_no,NULL AS kys_no,'' AS kys_name FROM ky_kosinkai WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 			) AS KARIKYS ";
                tmp_sql = tmp_sql + " 			/*仮契約の契約者を取得 end*/ ";
                tmp_sql = tmp_sql + " 		) AS KYKYSTOTAL ";
                tmp_sql = tmp_sql + " 		LEFT JOIN ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_no ";
                tmp_sql = tmp_sql + " 					,ky_no ";
                tmp_sql = tmp_sql + " 					,kazoku_no ";
                tmp_sql = tmp_sql + " 					,kazoku_name ";
                tmp_sql = tmp_sql + " 					,aida ";
                tmp_sql = tmp_sql + " 				FROM ky_kazoku ";
                tmp_sql = tmp_sql + " 			) AS KYKAZOKU ";
                tmp_sql = tmp_sql + " 		ON  KYKYSTOTAL.bk_no = KYKAZOKU.bk_no ";
                tmp_sql = tmp_sql + " 		AND KYKYSTOTAL.hy_no = KYKAZOKU.hy_no ";
                tmp_sql = tmp_sql + " 		AND KYKYSTOTAL.ky_no = KYKAZOKU.ky_no ";
                tmp_sql = tmp_sql + " 		AND RTRIM(LTRIM(KYKYSTOTAL.kys_name)) = RTRIM(LTRIM(KYKAZOKU.kazoku_name)) ";
                tmp_sql = tmp_sql + " 		WHERE KYKYSTOTAL.ko_no <> 999 ";
                tmp_sql = tmp_sql + " 	) AS KYSTOTAL ";
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

    #region 契約入居者情報

    public class Ky_kazoku_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[入居者No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[入居者No],[契約者No],[氏名],[氏名Unicode],[カナ],[性別],[続柄],[生年月日],[勤務先],[連絡先],[携帯電話番号],[備考]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no ORDER BY sort_no) AS [入居者No] ";
                tmp_sql = tmp_sql + " 		,kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [氏名] ";
                tmp_sql = tmp_sql + " 		,kazoku_name AS [氏名Unicode] ";
                tmp_sql = tmp_sql + " 		,'' AS [カナ] ";
                tmp_sql = tmp_sql + " 		,sex AS [性別] ";
                tmp_sql = tmp_sql + " 		,aida AS [続柄] ";
                tmp_sql = tmp_sql + " 		,birthday AS [生年月日] ";
                tmp_sql = tmp_sql + " 		,kinmu_name AS [勤務先] ";
                tmp_sql = tmp_sql + " 		,'' AS [連絡先] ";
                tmp_sql = tmp_sql + " 		/*,renraku_tel AS [携帯電話番号]*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [携帯電話番号] ";
                tmp_sql = tmp_sql + " 		,biko AS [備考] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 KYKAZOKU.* ";
                tmp_sql = tmp_sql + " 				,CASE ";
                tmp_sql = tmp_sql + " 					WHEN sort_no IS NULL THEN 9999 + kazoku_no	/*契約者1,2,3の並びを維持するために9999を加算しておく*/ ";
                tmp_sql = tmp_sql + " 					ELSE sort_no ";
                tmp_sql = tmp_sql + " 				 END AS sort_no ";
                tmp_sql = tmp_sql + " 				,KYSMST.kys_no ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 KYKAZOKU.bk_no ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.hy_no ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.ky_no ";
                tmp_sql = tmp_sql + " 					,ko_no ";
                tmp_sql = tmp_sql + " 					,kazoku_no ";
                tmp_sql = tmp_sql + " 					,kazoku_name ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.sex ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.birthday ";
                tmp_sql = tmp_sql + " 					,aida ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.kinmu_name ";
                tmp_sql = tmp_sql + " 					,renraku_tel ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.biko ";
                tmp_sql = tmp_sql + " 				FROM ky_kazoku AS KYKAZOKU ";
                tmp_sql = tmp_sql + " 				LEFT JOIN ky_kosinkai AS KYKKONO ";
                tmp_sql = tmp_sql + " 					ON  KYKAZOKU.bk_no = KYKKONO.bk_no ";
                tmp_sql = tmp_sql + " 					AND KYKAZOKU.hy_no = KYKKONO.hy_no ";
                tmp_sql = tmp_sql + " 					AND KYKAZOKU.ky_no = KYKKONO.ky_no ";
                tmp_sql = tmp_sql + " 				WHERE KYKKONO.kari_flg = 0 ";
                tmp_sql = tmp_sql + " 			) AS KYKAZOKU ";
                tmp_sql = tmp_sql + " 			LEFT JOIN kys_mst AS KYSMST ";
                tmp_sql = tmp_sql + " 			ON RTRIM(LTRIM(KYKAZOKU.kazoku_name)) = RTRIM(LTRIM(KYSMST.kys_name)) ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,1 AS sort_no,kys_no FROM ky_kosinkai WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 					UNION ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,2 AS sort_no,kys_no2 FROM ky_kosinkai WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 					UNION ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,3 AS sort_no,kys_no3 FROM ky_kosinkai WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 				) AS KYKYSTOTAL ";
                tmp_sql = tmp_sql + " 			ON  KYKAZOKU.bk_no = KYKYSTOTAL.bk_no ";
                tmp_sql = tmp_sql + " 			AND KYKAZOKU.hy_no = KYKYSTOTAL.hy_no ";
                tmp_sql = tmp_sql + " 			AND KYKAZOKU.ky_no = KYKYSTOTAL.ky_no ";
                tmp_sql = tmp_sql + " 			AND KYKAZOKU.ko_no = KYKYSTOTAL.ko_no ";
                tmp_sql = tmp_sql + " 			AND KYKYSTOTAL.kys_no = KYSMST.kys_no ";
                tmp_sql = tmp_sql + " 		) AS KYKAZOKU ";
                tmp_sql = tmp_sql + " 		UNION ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 KYKAZOKU.* ";
                tmp_sql = tmp_sql + " 				,CASE ";
                tmp_sql = tmp_sql + " 					WHEN sort_no IS NULL THEN 9999 + kazoku_no	/*契約者1,2,3の並びを維持するために9999を加算しておく*/ ";
                tmp_sql = tmp_sql + " 					ELSE sort_no ";
                tmp_sql = tmp_sql + " 				 END AS sort_no ";
                tmp_sql = tmp_sql + " 				,KARIKYSMST.[仮契約者No] ";
                tmp_sql = tmp_sql + " 			FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 KYKAZOKU.bk_no ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.hy_no ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.ky_no ";
                tmp_sql = tmp_sql + " 					,ko_no ";
                tmp_sql = tmp_sql + " 					,kazoku_no ";
                tmp_sql = tmp_sql + " 					,kazoku_name ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.sex ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.birthday ";
                tmp_sql = tmp_sql + " 					,aida ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.kinmu_name ";
                tmp_sql = tmp_sql + " 					,renraku_tel ";
                tmp_sql = tmp_sql + " 					,KYKAZOKU.biko ";
                tmp_sql = tmp_sql + " 				FROM ky_kazoku AS KYKAZOKU ";
                tmp_sql = tmp_sql + " 				LEFT JOIN ky_kosinkai AS KYKKONO ";
                tmp_sql = tmp_sql + " 					ON  KYKAZOKU.bk_no = KYKKONO.bk_no ";
                tmp_sql = tmp_sql + " 					AND KYKAZOKU.hy_no = KYKKONO.hy_no ";
                tmp_sql = tmp_sql + " 					AND KYKAZOKU.ky_no = KYKKONO.ky_no ";
                tmp_sql = tmp_sql + " 				WHERE KYKKONO.kari_flg = 1 ";
                tmp_sql = tmp_sql + " 			) AS KYKAZOKU ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 							( ";
                tmp_sql = tmp_sql + " 								SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 							)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 						,kari_name ";
                tmp_sql = tmp_sql + " 					FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 					WHERE ISNULL(kari_name,'') <> ''	 ";
                tmp_sql = tmp_sql + " 				) AS KARIKYSMST ";
                tmp_sql = tmp_sql + " 			ON RTRIM(LTRIM(KYKAZOKU.kazoku_name)) = RTRIM(LTRIM(KARIKYSMST.kari_name)) ";
                tmp_sql = tmp_sql + " 			LEFT JOIN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,1 AS sort_no,kari_name FROM ky_kosinkai WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 					UNION ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,2 AS sort_no,'' AS kari_name FROM ky_kosinkai WHERE kari_flg = 1 ";
                tmp_sql = tmp_sql + " 					UNION ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,3 AS sort_no,'' AS kari_name FROM ky_kosinkai WHERE kari_flg = 1 ";
                tmp_sql = tmp_sql + " 				) AS KYKYSTOTAL ";
                tmp_sql = tmp_sql + " 			ON  KYKAZOKU.bk_no = KYKYSTOTAL.bk_no ";
                tmp_sql = tmp_sql + " 			AND KYKAZOKU.hy_no = KYKYSTOTAL.hy_no ";
                tmp_sql = tmp_sql + " 			AND KYKAZOKU.ky_no = KYKYSTOTAL.ky_no ";
                tmp_sql = tmp_sql + " 			AND KYKAZOKU.ko_no = KYKYSTOTAL.ko_no ";
                tmp_sql = tmp_sql + " 			AND KYKYSTOTAL.kari_name = KARIKYSMST.kari_name ";
                tmp_sql = tmp_sql + " 		) AS KARIKYKAZOKU ";
                tmp_sql = tmp_sql + " 	) AS KAZOKUTOTAL ";
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

    #region 契約保証人情報

    public class Ky_kosinkai_hosyonin_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[並び順]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[並び順],[契約者No],[保証人No]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,[並び順] ";
                tmp_sql = tmp_sql + " 		,kys_no AS [契約者No] ";
                tmp_sql = tmp_sql + " 		,hosyonin_no AS [保証人No] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*契約者No、保証人Noの昇順に連番を振る sta*/ ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 * ";
                tmp_sql = tmp_sql + " 			,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no ORDER BY sort_no,hosyonin_no) AS [並び順] ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			/*本契約の保証人情報を取得 sta*/ ";
                tmp_sql = tmp_sql + " 			SELECT * FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no ";
                tmp_sql = tmp_sql + " 					,hy_no ";
                tmp_sql = tmp_sql + " 					,ky_no ";
                tmp_sql = tmp_sql + " 					,ko_no ";
                tmp_sql = tmp_sql + " 					,sort_no ";
                tmp_sql = tmp_sql + " 					,KYKYS.kys_no ";
                tmp_sql = tmp_sql + " 					,hosyonin_no ";
                tmp_sql = tmp_sql + " 					,hs1_name ";
                tmp_sql = tmp_sql + " 				FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,1 AS sort_no,kys_no FROM ky_kosinkai	WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 					UNION ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,2 AS sort_no,kys_no2 FROM ky_kosinkai WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 					UNION ";
                tmp_sql = tmp_sql + " 					SELECT bk_no,hy_no,ky_no,ko_no,3 AS sort_no,kys_no3 FROM ky_kosinkai WHERE kari_flg = 0 ";
                tmp_sql = tmp_sql + " 				) AS KYKYS ";
                tmp_sql = tmp_sql + " 				LEFT JOIN ";
                tmp_sql = tmp_sql + " 					( ";
                tmp_sql = tmp_sql + " 						SELECT ";
                tmp_sql = tmp_sql + " 							 kys_no ";
                tmp_sql = tmp_sql + " 							,1 AS hosyonin_no ";
                tmp_sql = tmp_sql + " 							,hs1_name ";
                tmp_sql = tmp_sql + " 						FROM kys_mst ";
                tmp_sql = tmp_sql + " 						UNION ";
                tmp_sql = tmp_sql + " 						SELECT ";
                tmp_sql = tmp_sql + " 							 kys_no ";
                tmp_sql = tmp_sql + " 							,2 AS hosyonin_no ";
                tmp_sql = tmp_sql + " 							,hs2_name ";
                tmp_sql = tmp_sql + " 						FROM kys_mst ";
                tmp_sql = tmp_sql + " 					) AS KYSHOSYO ";
                tmp_sql = tmp_sql + " 				ON KYKYS.kys_no = KYSHOSYO.kys_no ";
                tmp_sql = tmp_sql + " 				WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 				AND   ISNULL(hs1_name,'') <> '' ";
                tmp_sql = tmp_sql + " 			) KYSHOSYONIN ";
                tmp_sql = tmp_sql + " 			/*本契約の保証人情報を取得 end*/ ";
                tmp_sql = tmp_sql + " 			UNION ";
                tmp_sql = tmp_sql + " 			/*仮契約の保証人情報を取得 sta*/ ";
                tmp_sql = tmp_sql + " 			SELECT * FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no,hy_no,ky_no,ko_no,1 AS sort_no ";
                tmp_sql = tmp_sql + " 					,ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 						)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 					,1 AS hosyonin_no ";
                tmp_sql = tmp_sql + " 					,hs1_name ";
                tmp_sql = tmp_sql + " 				FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 				WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 				AND ISNULL(hs1_name,'') <> '' ";
                tmp_sql = tmp_sql + " 				UNION ";
                tmp_sql = tmp_sql + " 				SELECT ";
                tmp_sql = tmp_sql + " 					 bk_no,hy_no,ky_no,ko_no,1 AS sort_no ";
                tmp_sql = tmp_sql + " 					,ROW_NUMBER()OVER(ORDER BY kari_name) + ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							SELECT MAX(kys_no) FROM kys_mst ";
                tmp_sql = tmp_sql + " 						)  AS [仮契約者No] ";
                tmp_sql = tmp_sql + " 					,2 AS hosyonin_no ";
                tmp_sql = tmp_sql + " 					,hs2_name ";
                tmp_sql = tmp_sql + " 				FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 				WHERE ISNULL(kari_name,'') <> '' ";
                tmp_sql = tmp_sql + " 				AND ISNULL(hs2_name,'') <> '' ";
                tmp_sql = tmp_sql + " 			) AS KARIKYSHOSYONIN ";
                tmp_sql = tmp_sql + " 			/*仮契約の保証人情報を取得 end*/ ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		/*契約者No、保証人Noの昇順に連番を振る end*/ ";
                tmp_sql = tmp_sql + " 	) AS TOTAL ";
                tmp_sql = tmp_sql + " 	WHERE [並び順] IN (1,2) ";
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

    #region 契約車情報

    public class Ky_car_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[車情報No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[車情報No],[メーカー],[車名],[車色],[ナンバー],[備考],[駐車区画]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	/*********************************************************** ";
                tmp_sql = tmp_sql + " 	契約車情報抽出 ";
                tmp_sql = tmp_sql + " 	  V7では履歴管理されていないが10では管理されているため ";
                tmp_sql = tmp_sql + " 	  最新情報を全履歴へ反映させるように抽出する ";
                tmp_sql = tmp_sql + " 	***********************************************************/ ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 KYCAR.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,KYCAR.hy_no AS [部屋No]  ";
                tmp_sql = tmp_sql + " 		,KYCAR.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,KYK.ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,car_no AS [車情報No] ";
                tmp_sql = tmp_sql + " 		,maker AS [メーカー] ";
                tmp_sql = tmp_sql + " 		,syasyu AS [車名] ";
                tmp_sql = tmp_sql + " 		,color AS [車色] ";
                tmp_sql = tmp_sql + " 		,bango AS [ナンバー] ";
                tmp_sql = tmp_sql + " 		,KYCAR.biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,cyuban AS [駐車区画] ";
                tmp_sql = tmp_sql + " 	FROM ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,ky_no ";
                tmp_sql = tmp_sql + " 			,car_no ";
                tmp_sql = tmp_sql + " 			,maker ";
                tmp_sql = tmp_sql + " 			,syasyu ";
                tmp_sql = tmp_sql + " 			,color ";
                tmp_sql = tmp_sql + " 			,bango ";
                tmp_sql = tmp_sql + " 			,biko ";
                tmp_sql = tmp_sql + " 			,cyuban ";
                tmp_sql = tmp_sql + " 		FROM ky_car ";
                tmp_sql = tmp_sql + " 	)  AS KYCAR ";
                tmp_sql = tmp_sql + " 	ON  KYK.bk_no = KYCAR.bk_no ";
                tmp_sql = tmp_sql + " 	AND KYK.hy_no = KYCAR.hy_no ";
                tmp_sql = tmp_sql + " 	AND KYK.ky_no = KYCAR.ky_no ";
                tmp_sql = tmp_sql + " 	WHERE KYK.ko_no <> 999 ";
                tmp_sql = tmp_sql + " 	AND   KYCAR.bk_no IS NOT NULL ";
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

    #region 契約保険情報

    public class ky_Hoken_Repository
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

                sortstr = "[物件No],[部屋No],[契約No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[保険種類No],[保険業者No],[契約日],[適用開始年月日],[適用終了年月日],[保険金額],[満期案内通知有無],[備考],[証券番号]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,hknrui_no AS [保険種類No] ";
                tmp_sql = tmp_sql + " 		,hkngy_no AS [保険業者No] ";
                tmp_sql = tmp_sql + " 		,ky_date AS [契約日] ";
                tmp_sql = tmp_sql + " 		,start_ymd AS [適用開始年月日] ";
                tmp_sql = tmp_sql + " 		,end_ymd AS [適用終了年月日] ";
                tmp_sql = tmp_sql + " 		,hoken_gak AS [保険金額] ";
                tmp_sql = tmp_sql + " 		,manki_umu AS [満期案内通知有無] ";
                tmp_sql = tmp_sql + " 		,biko AS [備考] ";
                tmp_sql = tmp_sql + " 		/*,hknsyoken_bango AS [証券番号] ※カスタマイズ*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [証券番号] ";
                tmp_sql = tmp_sql + " 	FROM ky_hoken ";
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

    #region 契約特約事項情報

    public class ky_tokuyaku_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[特約区分] ";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[特約区分],[内容1],[内容2],[内容3],[内容4],[内容5],[内容6],[内容7],[内容8],[内容9],[内容10],[内容11],[内容12],[内容13],[内容14],[内容15],[内容16],[内容17],[内容18],[内容19],[内容20],[内容21],[内容22],[内容23],[内容24],[内容25],[内容26],[内容27],[内容28],[内容29],[内容30],[内容31],[内容32],[内容33],[内容34],[内容35],[内容36],[内容37],[内容38],[内容39],[内容40],[内容41],[内容42],[内容43],[内容44],[内容45],[内容46],[内容47],[内容48],[内容49],[内容50],[内容51],[内容52],[内容53],[内容54],[内容55],[内容56],[内容57],[内容58],[内容59],[内容60],[内容61],[内容62],[内容63],[内容64],[内容65],[内容66],[内容67],[内容68],[内容69],[内容70],[内容71],[内容72],[内容73],[内容74],[内容75],[内容76],[内容77],[内容78],[内容79],[内容80],[内容81],[内容82],[内容83],[内容84],[内容85],[内容86],[内容87],[内容88],[内容89],[内容90],[内容91],[内容92],[内容93],[内容94],[内容95],[内容96],[内容97],[内容98],[内容99],[内容100],[内容101]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 			,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 			,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 			,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 			,1 AS [特約区分] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 1 THEN tokuyaku ELSE '' END) AS 内容1 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 2 THEN tokuyaku ELSE '' END) AS 内容2 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 3 THEN tokuyaku ELSE '' END) AS 内容3 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 4 THEN tokuyaku ELSE '' END) AS 内容4 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 5 THEN tokuyaku ELSE '' END) AS 内容5 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 6 THEN tokuyaku ELSE '' END) AS 内容6 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 7 THEN tokuyaku ELSE '' END) AS 内容7 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 8 THEN tokuyaku ELSE '' END) AS 内容8 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 9 THEN tokuyaku ELSE '' END) AS 内容9 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 10 THEN tokuyaku ELSE '' END) AS 内容10 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 11 THEN tokuyaku ELSE '' END) AS 内容11 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 12 THEN tokuyaku ELSE '' END) AS 内容12 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 13 THEN tokuyaku ELSE '' END) AS 内容13 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 14 THEN tokuyaku ELSE '' END) AS 内容14 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 15 THEN tokuyaku ELSE '' END) AS 内容15 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 16 THEN tokuyaku ELSE '' END) AS 内容16 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 17 THEN tokuyaku ELSE '' END) AS 内容17 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 18 THEN tokuyaku ELSE '' END) AS 内容18 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 19 THEN tokuyaku ELSE '' END) AS 内容19 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 20 THEN tokuyaku ELSE '' END) AS 内容20 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 21 THEN tokuyaku ELSE '' END) AS 内容21 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 22 THEN tokuyaku ELSE '' END) AS 内容22 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 23 THEN tokuyaku ELSE '' END) AS 内容23 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 24 THEN tokuyaku ELSE '' END) AS 内容24 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 25 THEN tokuyaku ELSE '' END) AS 内容25 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 26 THEN tokuyaku ELSE '' END) AS 内容26 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 27 THEN tokuyaku ELSE '' END) AS 内容27 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 28 THEN tokuyaku ELSE '' END) AS 内容28 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 29 THEN tokuyaku ELSE '' END) AS 内容29 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 30 THEN tokuyaku ELSE '' END) AS 内容30 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 31 THEN tokuyaku ELSE '' END) AS 内容31 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 32 THEN tokuyaku ELSE '' END) AS 内容32 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 33 THEN tokuyaku ELSE '' END) AS 内容33 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 34 THEN tokuyaku ELSE '' END) AS 内容34 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 35 THEN tokuyaku ELSE '' END) AS 内容35 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 36 THEN tokuyaku ELSE '' END) AS 内容36 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 37 THEN tokuyaku ELSE '' END) AS 内容37 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 38 THEN tokuyaku ELSE '' END) AS 内容38 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 39 THEN tokuyaku ELSE '' END) AS 内容39 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 40 THEN tokuyaku ELSE '' END) AS 内容40 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 41 THEN tokuyaku ELSE '' END) AS 内容41 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 42 THEN tokuyaku ELSE '' END) AS 内容42 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 43 THEN tokuyaku ELSE '' END) AS 内容43 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 44 THEN tokuyaku ELSE '' END) AS 内容44 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 45 THEN tokuyaku ELSE '' END) AS 内容45 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 46 THEN tokuyaku ELSE '' END) AS 内容46 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 47 THEN tokuyaku ELSE '' END) AS 内容47 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 48 THEN tokuyaku ELSE '' END) AS 内容48 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 49 THEN tokuyaku ELSE '' END) AS 内容49 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 50 THEN tokuyaku ELSE '' END) AS 内容50 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 51 THEN tokuyaku ELSE '' END) AS 内容51 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 52 THEN tokuyaku ELSE '' END) AS 内容52 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 53 THEN tokuyaku ELSE '' END) AS 内容53 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 54 THEN tokuyaku ELSE '' END) AS 内容54 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 55 THEN tokuyaku ELSE '' END) AS 内容55 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 56 THEN tokuyaku ELSE '' END) AS 内容56 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 57 THEN tokuyaku ELSE '' END) AS 内容57 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 58 THEN tokuyaku ELSE '' END) AS 内容58 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 59 THEN tokuyaku ELSE '' END) AS 内容59 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 60 THEN tokuyaku ELSE '' END) AS 内容60 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 61 THEN tokuyaku ELSE '' END) AS 内容61 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 62 THEN tokuyaku ELSE '' END) AS 内容62 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 63 THEN tokuyaku ELSE '' END) AS 内容63 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 64 THEN tokuyaku ELSE '' END) AS 内容64 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 65 THEN tokuyaku ELSE '' END) AS 内容65 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 66 THEN tokuyaku ELSE '' END) AS 内容66 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 67 THEN tokuyaku ELSE '' END) AS 内容67 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 68 THEN tokuyaku ELSE '' END) AS 内容68 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 69 THEN tokuyaku ELSE '' END) AS 内容69 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 70 THEN tokuyaku ELSE '' END) AS 内容70 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 71 THEN tokuyaku ELSE '' END) AS 内容71 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 72 THEN tokuyaku ELSE '' END) AS 内容72 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 73 THEN tokuyaku ELSE '' END) AS 内容73 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 74 THEN tokuyaku ELSE '' END) AS 内容74 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 75 THEN tokuyaku ELSE '' END) AS 内容75 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 76 THEN tokuyaku ELSE '' END) AS 内容76 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 77 THEN tokuyaku ELSE '' END) AS 内容77 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 78 THEN tokuyaku ELSE '' END) AS 内容78 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 79 THEN tokuyaku ELSE '' END) AS 内容79 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 80 THEN tokuyaku ELSE '' END) AS 内容80 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 81 THEN tokuyaku ELSE '' END) AS 内容81 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 82 THEN tokuyaku ELSE '' END) AS 内容82 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 83 THEN tokuyaku ELSE '' END) AS 内容83 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 84 THEN tokuyaku ELSE '' END) AS 内容84 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 85 THEN tokuyaku ELSE '' END) AS 内容85 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 86 THEN tokuyaku ELSE '' END) AS 内容86 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 87 THEN tokuyaku ELSE '' END) AS 内容87 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 88 THEN tokuyaku ELSE '' END) AS 内容88 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 89 THEN tokuyaku ELSE '' END) AS 内容89 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 90 THEN tokuyaku ELSE '' END) AS 内容90 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 91 THEN tokuyaku ELSE '' END) AS 内容91 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 92 THEN tokuyaku ELSE '' END) AS 内容92 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 93 THEN tokuyaku ELSE '' END) AS 内容93 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 94 THEN tokuyaku ELSE '' END) AS 内容94 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 95 THEN tokuyaku ELSE '' END) AS 内容95 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 96 THEN tokuyaku ELSE '' END) AS 内容96 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 97 THEN tokuyaku ELSE '' END) AS 内容97 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 98 THEN tokuyaku ELSE '' END) AS 内容98 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 99 THEN tokuyaku ELSE '' END) AS 内容99 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 100 THEN tokuyaku ELSE '' END) AS 内容100 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 101 THEN tokuyaku ELSE '' END) AS 内容101 ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 bk_no ";
                tmp_sql = tmp_sql + " 				,hy_no ";
                tmp_sql = tmp_sql + " 				,ky_no ";
                tmp_sql = tmp_sql + " 				,ko_no ";
                tmp_sql = tmp_sql + " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no ORDER BY bk_no,hy_no,ky_no,ko_no,mei_no) AS [特約行No] ";
                tmp_sql = tmp_sql + " 				,tokuyaku ";
                tmp_sql = tmp_sql + " 			FROM ky_tokuyaku ";
                tmp_sql = tmp_sql + " 			WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		GROUP BY bk_no,hy_no,ky_no,ko_no ";
                tmp_sql = tmp_sql + " 	) AS TOKUYAKU ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,ky_no ";
                tmp_sql = tmp_sql + " 			,ko_no ";
                tmp_sql = tmp_sql + " 			,2 AS [特約区分] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 1 THEN tokuyaku ELSE '' END) AS 内容1 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 2 THEN tokuyaku ELSE '' END) AS 内容2 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 3 THEN tokuyaku ELSE '' END) AS 内容3 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 4 THEN tokuyaku ELSE '' END) AS 内容4 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 5 THEN tokuyaku ELSE '' END) AS 内容5 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 6 THEN tokuyaku ELSE '' END) AS 内容6 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 7 THEN tokuyaku ELSE '' END) AS 内容7 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 8 THEN tokuyaku ELSE '' END) AS 内容8 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 9 THEN tokuyaku ELSE '' END) AS 内容9 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 10 THEN tokuyaku ELSE '' END) AS 内容10 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 11 THEN tokuyaku ELSE '' END) AS 内容11 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 12 THEN tokuyaku ELSE '' END) AS 内容12 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 13 THEN tokuyaku ELSE '' END) AS 内容13 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 14 THEN tokuyaku ELSE '' END) AS 内容14 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 15 THEN tokuyaku ELSE '' END) AS 内容15 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 16 THEN tokuyaku ELSE '' END) AS 内容16 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 17 THEN tokuyaku ELSE '' END) AS 内容17 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 18 THEN tokuyaku ELSE '' END) AS 内容18 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 19 THEN tokuyaku ELSE '' END) AS 内容19 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 20 THEN tokuyaku ELSE '' END) AS 内容20 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 21 THEN tokuyaku ELSE '' END) AS 内容21 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 22 THEN tokuyaku ELSE '' END) AS 内容22 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 23 THEN tokuyaku ELSE '' END) AS 内容23 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 24 THEN tokuyaku ELSE '' END) AS 内容24 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 25 THEN tokuyaku ELSE '' END) AS 内容25 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 26 THEN tokuyaku ELSE '' END) AS 内容26 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 27 THEN tokuyaku ELSE '' END) AS 内容27 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 28 THEN tokuyaku ELSE '' END) AS 内容28 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 29 THEN tokuyaku ELSE '' END) AS 内容29 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 30 THEN tokuyaku ELSE '' END) AS 内容30 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 31 THEN tokuyaku ELSE '' END) AS 内容31 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 32 THEN tokuyaku ELSE '' END) AS 内容32 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 33 THEN tokuyaku ELSE '' END) AS 内容33 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 34 THEN tokuyaku ELSE '' END) AS 内容34 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 35 THEN tokuyaku ELSE '' END) AS 内容35 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 36 THEN tokuyaku ELSE '' END) AS 内容36 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 37 THEN tokuyaku ELSE '' END) AS 内容37 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 38 THEN tokuyaku ELSE '' END) AS 内容38 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 39 THEN tokuyaku ELSE '' END) AS 内容39 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 40 THEN tokuyaku ELSE '' END) AS 内容40 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 41 THEN tokuyaku ELSE '' END) AS 内容41 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 42 THEN tokuyaku ELSE '' END) AS 内容42 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 43 THEN tokuyaku ELSE '' END) AS 内容43 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 44 THEN tokuyaku ELSE '' END) AS 内容44 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 45 THEN tokuyaku ELSE '' END) AS 内容45 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 46 THEN tokuyaku ELSE '' END) AS 内容46 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 47 THEN tokuyaku ELSE '' END) AS 内容47 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 48 THEN tokuyaku ELSE '' END) AS 内容48 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 49 THEN tokuyaku ELSE '' END) AS 内容49 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 50 THEN tokuyaku ELSE '' END) AS 内容50 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容51,'' AS 内容52,'' AS 内容53,'' AS 内容54,'' AS 内容55,'' AS 内容56,'' AS 内容57,'' AS 内容58,'' AS 内容59,'' AS 内容60 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容61,'' AS 内容62,'' AS 内容63,'' AS 内容64,'' AS 内容65,'' AS 内容66,'' AS 内容67,'' AS 内容68,'' AS 内容69,'' AS 内容70 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容71,'' AS 内容72,'' AS 内容73,'' AS 内容74,'' AS 内容75,'' AS 内容76,'' AS 内容77,'' AS 内容78,'' AS 内容79,'' AS 内容80 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容81,'' AS 内容82,'' AS 内容83,'' AS 内容84,'' AS 内容85,'' AS 内容86,'' AS 内容87,'' AS 内容88,'' AS 内容89,'' AS 内容90 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容91,'' AS 内容92,'' AS 内容93,'' AS 内容94,'' AS 内容95,'' AS 内容96,'' AS 内容97,'' AS 内容98,'' AS 内容99,'' AS 内容100 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容101			 ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 bk_no ";
                tmp_sql = tmp_sql + " 				,hy_no ";
                tmp_sql = tmp_sql + " 				,ky_no ";
                tmp_sql = tmp_sql + " 				,ko_no ";
                tmp_sql = tmp_sql + " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no ORDER BY bk_no,hy_no,ky_no,ko_no,mei_no) AS [特約行No] ";
                tmp_sql = tmp_sql + " 				,tokuyaku ";
                tmp_sql = tmp_sql + " 			FROM ky_genjo_tokuyaku ";
                tmp_sql = tmp_sql + " 			WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 			AND   brui_no = 1 ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		GROUP BY bk_no,hy_no,ky_no,ko_no ";
                tmp_sql = tmp_sql + " 	) AS GENJOTOKUYAKU_GENJO ";
                tmp_sql = tmp_sql + " 	UNION ";
                tmp_sql = tmp_sql + " 	SELECT * FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,ky_no ";
                tmp_sql = tmp_sql + " 			,ko_no ";
                tmp_sql = tmp_sql + " 			,3 AS [特約区分] ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 1 THEN tokuyaku ELSE '' END) AS 内容1 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 2 THEN tokuyaku ELSE '' END) AS 内容2 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 3 THEN tokuyaku ELSE '' END) AS 内容3 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 4 THEN tokuyaku ELSE '' END) AS 内容4 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 5 THEN tokuyaku ELSE '' END) AS 内容5 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 6 THEN tokuyaku ELSE '' END) AS 内容6 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 7 THEN tokuyaku ELSE '' END) AS 内容7 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 8 THEN tokuyaku ELSE '' END) AS 内容8 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 9 THEN tokuyaku ELSE '' END) AS 内容9 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 10 THEN tokuyaku ELSE '' END) AS 内容10 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 11 THEN tokuyaku ELSE '' END) AS 内容11 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 12 THEN tokuyaku ELSE '' END) AS 内容12 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 13 THEN tokuyaku ELSE '' END) AS 内容13 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 14 THEN tokuyaku ELSE '' END) AS 内容14 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 15 THEN tokuyaku ELSE '' END) AS 内容15 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 16 THEN tokuyaku ELSE '' END) AS 内容16 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 17 THEN tokuyaku ELSE '' END) AS 内容17 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 18 THEN tokuyaku ELSE '' END) AS 内容18 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 19 THEN tokuyaku ELSE '' END) AS 内容19 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 20 THEN tokuyaku ELSE '' END) AS 内容20 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 21 THEN tokuyaku ELSE '' END) AS 内容21 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 22 THEN tokuyaku ELSE '' END) AS 内容22 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 23 THEN tokuyaku ELSE '' END) AS 内容23 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 24 THEN tokuyaku ELSE '' END) AS 内容24 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 25 THEN tokuyaku ELSE '' END) AS 内容25 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 26 THEN tokuyaku ELSE '' END) AS 内容26 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 27 THEN tokuyaku ELSE '' END) AS 内容27 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 28 THEN tokuyaku ELSE '' END) AS 内容28 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 29 THEN tokuyaku ELSE '' END) AS 内容29 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 30 THEN tokuyaku ELSE '' END) AS 内容30 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 31 THEN tokuyaku ELSE '' END) AS 内容31 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 32 THEN tokuyaku ELSE '' END) AS 内容32 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 33 THEN tokuyaku ELSE '' END) AS 内容33 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 34 THEN tokuyaku ELSE '' END) AS 内容34 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 35 THEN tokuyaku ELSE '' END) AS 内容35 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 36 THEN tokuyaku ELSE '' END) AS 内容36 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 37 THEN tokuyaku ELSE '' END) AS 内容37 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 38 THEN tokuyaku ELSE '' END) AS 内容38 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 39 THEN tokuyaku ELSE '' END) AS 内容39 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 40 THEN tokuyaku ELSE '' END) AS 内容40 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 41 THEN tokuyaku ELSE '' END) AS 内容41 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 42 THEN tokuyaku ELSE '' END) AS 内容42 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 43 THEN tokuyaku ELSE '' END) AS 内容43 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 44 THEN tokuyaku ELSE '' END) AS 内容44 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 45 THEN tokuyaku ELSE '' END) AS 内容45 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 46 THEN tokuyaku ELSE '' END) AS 内容46 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 47 THEN tokuyaku ELSE '' END) AS 内容47 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 48 THEN tokuyaku ELSE '' END) AS 内容48 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 49 THEN tokuyaku ELSE '' END) AS 内容49 ";
                tmp_sql = tmp_sql + " 			,MAX(CASE WHEN [特約行No] = 50 THEN tokuyaku ELSE '' END) AS 内容50 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容51,'' AS 内容52,'' AS 内容53,'' AS 内容54,'' AS 内容55,'' AS 内容56,'' AS 内容57,'' AS 内容58,'' AS 内容59,'' AS 内容60 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容61,'' AS 内容62,'' AS 内容63,'' AS 内容64,'' AS 内容65,'' AS 内容66,'' AS 内容67,'' AS 内容68,'' AS 内容69,'' AS 内容70 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容71,'' AS 内容72,'' AS 内容73,'' AS 内容74,'' AS 内容75,'' AS 内容76,'' AS 内容77,'' AS 内容78,'' AS 内容79,'' AS 内容80 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容81,'' AS 内容82,'' AS 内容83,'' AS 内容84,'' AS 内容85,'' AS 内容86,'' AS 内容87,'' AS 内容88,'' AS 内容89,'' AS 内容90 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容91,'' AS 内容92,'' AS 内容93,'' AS 内容94,'' AS 内容95,'' AS 内容96,'' AS 内容97,'' AS 内容98,'' AS 内容99,'' AS 内容100 ";
                tmp_sql = tmp_sql + " 			,'' AS 内容101			 ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 bk_no ";
                tmp_sql = tmp_sql + " 				,hy_no ";
                tmp_sql = tmp_sql + " 				,ky_no ";
                tmp_sql = tmp_sql + " 				,ko_no				 ";
                tmp_sql = tmp_sql + " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no ORDER BY bk_no,hy_no,ky_no,ko_no,mei_no) AS [特約行No] ";
                tmp_sql = tmp_sql + " 				,tokuyaku ";
                tmp_sql = tmp_sql + " 			FROM ky_genjo_tokuyaku ";
                tmp_sql = tmp_sql + " 			WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 			AND   brui_no = 2 ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		GROUP BY bk_no,hy_no,ky_no,ko_no ";
                tmp_sql = tmp_sql + " 	) AS GENJOTOKUYAKU_NYUKYO ";
                tmp_sql = tmp_sql + " ) AS TOKUYAKUGENJOTOTAL ";
                tmp_sql = tmp_sql + " /* ";
                tmp_sql = tmp_sql + " 以下は重要事項の抽出 移行仕様確認中(保留) ";
                tmp_sql = tmp_sql + " LEFT JOIN ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,1 AS [特約区分] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 1 THEN jyuyo_lstname ELSE '' END) AS 内容1 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 2 THEN jyuyo_lstname ELSE '' END) AS 内容2 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 3 THEN jyuyo_lstname ELSE '' END) AS 内容3 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 4 THEN jyuyo_lstname ELSE '' END) AS 内容4 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 5 THEN jyuyo_lstname ELSE '' END) AS 内容5 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 6 THEN jyuyo_lstname ELSE '' END) AS 内容6 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 7 THEN jyuyo_lstname ELSE '' END) AS 内容7 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 8 THEN jyuyo_lstname ELSE '' END) AS 内容8 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 9 THEN jyuyo_lstname ELSE '' END) AS 内容9 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 10 THEN jyuyo_lstname ELSE '' END) AS 内容10 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 11 THEN jyuyo_lstname ELSE '' END) AS 内容11 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 12 THEN jyuyo_lstname ELSE '' END) AS 内容12 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 13 THEN jyuyo_lstname ELSE '' END) AS 内容13 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 14 THEN jyuyo_lstname ELSE '' END) AS 内容14 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 15 THEN jyuyo_lstname ELSE '' END) AS 内容15 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 16 THEN jyuyo_lstname ELSE '' END) AS 内容16 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 17 THEN jyuyo_lstname ELSE '' END) AS 内容17 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 18 THEN jyuyo_lstname ELSE '' END) AS 内容18 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 19 THEN jyuyo_lstname ELSE '' END) AS 内容19 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 20 THEN jyuyo_lstname ELSE '' END) AS 内容20 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 21 THEN jyuyo_lstname ELSE '' END) AS 内容21 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 22 THEN jyuyo_lstname ELSE '' END) AS 内容22 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 23 THEN jyuyo_lstname ELSE '' END) AS 内容23 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 24 THEN jyuyo_lstname ELSE '' END) AS 内容24 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 25 THEN jyuyo_lstname ELSE '' END) AS 内容25 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 26 THEN jyuyo_lstname ELSE '' END) AS 内容26 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 27 THEN jyuyo_lstname ELSE '' END) AS 内容27 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 28 THEN jyuyo_lstname ELSE '' END) AS 内容28 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 29 THEN jyuyo_lstname ELSE '' END) AS 内容29 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 30 THEN jyuyo_lstname ELSE '' END) AS 内容30 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 31 THEN jyuyo_lstname ELSE '' END) AS 内容31 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 32 THEN jyuyo_lstname ELSE '' END) AS 内容32 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 33 THEN jyuyo_lstname ELSE '' END) AS 内容33 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 34 THEN jyuyo_lstname ELSE '' END) AS 内容34 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 35 THEN jyuyo_lstname ELSE '' END) AS 内容35 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 36 THEN jyuyo_lstname ELSE '' END) AS 内容36 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 37 THEN jyuyo_lstname ELSE '' END) AS 内容37 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 38 THEN jyuyo_lstname ELSE '' END) AS 内容38 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 39 THEN jyuyo_lstname ELSE '' END) AS 内容39 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 40 THEN jyuyo_lstname ELSE '' END) AS 内容40 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 41 THEN jyuyo_lstname ELSE '' END) AS 内容41 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 42 THEN jyuyo_lstname ELSE '' END) AS 内容42 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 43 THEN jyuyo_lstname ELSE '' END) AS 内容43 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 44 THEN jyuyo_lstname ELSE '' END) AS 内容44 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 45 THEN jyuyo_lstname ELSE '' END) AS 内容45 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 46 THEN jyuyo_lstname ELSE '' END) AS 内容46 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 47 THEN jyuyo_lstname ELSE '' END) AS 内容47 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 48 THEN jyuyo_lstname ELSE '' END) AS 内容48 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 49 THEN jyuyo_lstname ELSE '' END) AS 内容49 ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [重要事項行No] = 50 THEN jyuyo_lstname ELSE '' END) AS 内容50 ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,ky_no ";
                tmp_sql = tmp_sql + " 			,ko_no ";
                tmp_sql = tmp_sql + " 			,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no ORDER BY bk_no,hy_no,ky_no,ko_no,jyuyo_no) AS [重要事項行No] ";
                tmp_sql = tmp_sql + " 			,jyuyo_lstname ";
                tmp_sql = tmp_sql + " 		FROM ky_jyuyo ";
                tmp_sql = tmp_sql + " 		WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 	) AS VW ";
                tmp_sql = tmp_sql + " 	GROUP BY bk_no,hy_no,ky_no,ko_no ";
                tmp_sql = tmp_sql + " ) AS JYUYOJIKO ";
                tmp_sql = tmp_sql + " ON  TOKUYAKUGENJOTOTAL.[物件No] = JYUYOJIKO.[物件No] ";
                tmp_sql = tmp_sql + " AND TOKUYAKUGENJOTOTAL.[部屋No] = JYUYOJIKO.[部屋No] ";
                tmp_sql = tmp_sql + " AND TOKUYAKUGENJOTOTAL.[契約No] = JYUYOJIKO.[契約No] ";
                tmp_sql = tmp_sql + " AND TOKUYAKUGENJOTOTAL.[契約管理レコードNo] = JYUYOJIKO.[契約管理レコードNo] ";
                tmp_sql = tmp_sql + " AND TOKUYAKUGENJOTOTAL.[特約区分] = JYUYOJIKO.[特約区分] ";
                tmp_sql = tmp_sql + " */ ";

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

    #region 契約メモ情報

    public class Ky_biko_Repository
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

                sortstr = "[物件NO],[部屋NO],[契約NO],[契約管理レコードNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件NO],[部屋NO],[契約NO],[契約管理レコードNo],[メモ1],[メモ2],[メモ3],[メモ4],[メモ5],[メモ6],[メモ7],[メモ8],[メモ9],[メモ10],[メモ11],[メモ12],[メモ13],[メモ14],[メモ15],[メモ16],[メモ17],[メモ18],[メモ19],[メモ20],[メモ21],[メモ22],[メモ23],[メモ24],[メモ25],[メモ26],[メモ27],[メモ28],[メモ29],[メモ30],[メモ31],[メモ32],[メモ33],[メモ34],[メモ35],[メモ36],[メモ37],[メモ38],[メモ39],[メモ40],[メモ41],[メモ42],[メモ43],[メモ44],[メモ45],[メモ46],[メモ47],[メモ48],[メモ49],[メモ50]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件NO] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋NO] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約NO] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 1 THEN [メモ] ELSE '' END) AS [メモ1] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 2 THEN [メモ] ELSE '' END) AS [メモ2] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 3 THEN [メモ] ELSE '' END) AS [メモ3] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 4 THEN [メモ] ELSE '' END) AS [メモ4] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 5 THEN [メモ] ELSE '' END) AS [メモ5] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 6 THEN [メモ] ELSE '' END) AS [メモ6] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 7 THEN [メモ] ELSE '' END) AS [メモ7] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 8 THEN [メモ] ELSE '' END) AS [メモ8] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 9 THEN [メモ] ELSE '' END) AS [メモ9] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 10 THEN [メモ] ELSE '' END) AS [メモ10] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 11 THEN [メモ] ELSE '' END) AS [メモ11] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 12 THEN [メモ] ELSE '' END) AS [メモ12] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 13 THEN [メモ] ELSE '' END) AS [メモ13] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 14 THEN [メモ] ELSE '' END) AS [メモ14] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 15 THEN [メモ] ELSE '' END) AS [メモ15] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 16 THEN [メモ] ELSE '' END) AS [メモ16] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 17 THEN [メモ] ELSE '' END) AS [メモ17] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 18 THEN [メモ] ELSE '' END) AS [メモ18] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 19 THEN [メモ] ELSE '' END) AS [メモ19] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 20 THEN [メモ] ELSE '' END) AS [メモ20] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 21 THEN [メモ] ELSE '' END) AS [メモ21] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 22 THEN [メモ] ELSE '' END) AS [メモ22] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 23 THEN [メモ] ELSE '' END) AS [メモ23] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 24 THEN [メモ] ELSE '' END) AS [メモ24] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 25 THEN [メモ] ELSE '' END) AS [メモ25] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 26 THEN [メモ] ELSE '' END) AS [メモ26] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 27 THEN [メモ] ELSE '' END) AS [メモ27] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 28 THEN [メモ] ELSE '' END) AS [メモ28] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 29 THEN [メモ] ELSE '' END) AS [メモ29] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 30 THEN [メモ] ELSE '' END) AS [メモ30] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 31 THEN [メモ] ELSE '' END) AS [メモ31] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 32 THEN [メモ] ELSE '' END) AS [メモ32] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 33 THEN [メモ] ELSE '' END) AS [メモ33] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 34 THEN [メモ] ELSE '' END) AS [メモ34] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 35 THEN [メモ] ELSE '' END) AS [メモ35] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 36 THEN [メモ] ELSE '' END) AS [メモ36] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 37 THEN [メモ] ELSE '' END) AS [メモ37] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 38 THEN [メモ] ELSE '' END) AS [メモ38] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 39 THEN [メモ] ELSE '' END) AS [メモ39] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 40 THEN [メモ] ELSE '' END) AS [メモ40] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 41 THEN [メモ] ELSE '' END) AS [メモ41] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 42 THEN [メモ] ELSE '' END) AS [メモ42] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 43 THEN [メモ] ELSE '' END) AS [メモ43] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 44 THEN [メモ] ELSE '' END) AS [メモ44] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 45 THEN [メモ] ELSE '' END) AS [メモ45] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 46 THEN [メモ] ELSE '' END) AS [メモ46] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 47 THEN [メモ] ELSE '' END) AS [メモ47] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 48 THEN [メモ] ELSE '' END) AS [メモ48] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 49 THEN [メモ] ELSE '' END) AS [メモ49] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 50 THEN [メモ] ELSE '' END) AS [メモ50] ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 51 THEN [メモ] ELSE '' END) AS [メモ51] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 52 THEN [メモ] ELSE '' END) AS [メモ52] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 53 THEN [メモ] ELSE '' END) AS [メモ53] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 54 THEN [メモ] ELSE '' END) AS [メモ54] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 55 THEN [メモ] ELSE '' END) AS [メモ55] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 56 THEN [メモ] ELSE '' END) AS [メモ56] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 57 THEN [メモ] ELSE '' END) AS [メモ57] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 58 THEN [メモ] ELSE '' END) AS [メモ58] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 59 THEN [メモ] ELSE '' END) AS [メモ59] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 60 THEN [メモ] ELSE '' END) AS [メモ60] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 61 THEN [メモ] ELSE '' END) AS [メモ61] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 62 THEN [メモ] ELSE '' END) AS [メモ62] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 63 THEN [メモ] ELSE '' END) AS [メモ63] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 64 THEN [メモ] ELSE '' END) AS [メモ64] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 65 THEN [メモ] ELSE '' END) AS [メモ65] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 66 THEN [メモ] ELSE '' END) AS [メモ66] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 67 THEN [メモ] ELSE '' END) AS [メモ67] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 68 THEN [メモ] ELSE '' END) AS [メモ68] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 69 THEN [メモ] ELSE '' END) AS [メモ69] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 70 THEN [メモ] ELSE '' END) AS [メモ70] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 71 THEN [メモ] ELSE '' END) AS [メモ71] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 72 THEN [メモ] ELSE '' END) AS [メモ72] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 73 THEN [メモ] ELSE '' END) AS [メモ73] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 74 THEN [メモ] ELSE '' END) AS [メモ74] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 75 THEN [メモ] ELSE '' END) AS [メモ75] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 76 THEN [メモ] ELSE '' END) AS [メモ76] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 77 THEN [メモ] ELSE '' END) AS [メモ77] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 78 THEN [メモ] ELSE '' END) AS [メモ78] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 79 THEN [メモ] ELSE '' END) AS [メモ79] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 80 THEN [メモ] ELSE '' END) AS [メモ80] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 81 THEN [メモ] ELSE '' END) AS [メモ81] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 82 THEN [メモ] ELSE '' END) AS [メモ82] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 83 THEN [メモ] ELSE '' END) AS [メモ83] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 84 THEN [メモ] ELSE '' END) AS [メモ84] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 85 THEN [メモ] ELSE '' END) AS [メモ85] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 86 THEN [メモ] ELSE '' END) AS [メモ86] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 87 THEN [メモ] ELSE '' END) AS [メモ87] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 88 THEN [メモ] ELSE '' END) AS [メモ88] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 89 THEN [メモ] ELSE '' END) AS [メモ89] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 90 THEN [メモ] ELSE '' END) AS [メモ90] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 91 THEN [メモ] ELSE '' END) AS [メモ91] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 92 THEN [メモ] ELSE '' END) AS [メモ92] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 93 THEN [メモ] ELSE '' END) AS [メモ93] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 94 THEN [メモ] ELSE '' END) AS [メモ94] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 95 THEN [メモ] ELSE '' END) AS [メモ95] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 96 THEN [メモ] ELSE '' END) AS [メモ96] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 97 THEN [メモ] ELSE '' END) AS [メモ97] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 98 THEN [メモ] ELSE '' END) AS [メモ98] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 99 THEN [メモ] ELSE '' END) AS [メモ99] ";
                tmp_sql = tmp_sql + " 		,MAX(CASE WHEN [行No] = 100 THEN [メモ] ELSE '' END) AS [メモ100] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,ky_no ";
                tmp_sql = tmp_sql + " 			,ko_no ";
                tmp_sql = tmp_sql + " 			,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,ko_no ORDER BY bk_no,hy_no,ky_no,ko_no,mei_no) AS [行No] ";
                tmp_sql = tmp_sql + " 			,biko AS [メモ] ";
                tmp_sql = tmp_sql + " 		FROM ky_biko ";
                tmp_sql = tmp_sql + " 		WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 	) AS VW ";
                tmp_sql = tmp_sql + " 	GROUP BY bk_no,hy_no,ky_no,ko_no ";
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

    #region 契約入金項目情報

    public class Ky_sqdata_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[月区分],[並び順No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[月区分],[入金項目名],[入金項目レコードNo],[並び順No],[入金項目区分],[請求額],[税区分],[請求税額],[算出区分],[算出基準入金項目名],[算出ヶ月],[請求先No],[入金方法],[請求月区分],[フリーレント適用開始日],[フリーレント適用終了日],[フリーレント終了月請求額],[請求開始月],[請求パターン],[請求間隔],[請求発生年],[請求発生月],[適用税率],[備考],[フリーレント適用区分]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 KYSQ1.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,KYSQ1.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,KYSQ1.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		/*,ko_no AS [契約管理レコードNo]*/ ";
                tmp_sql = tmp_sql + " 		,[契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		/*,tuki_kbn AS [月区分]*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			/*20160520 敷金戻し、保証金戻しの月区分を入金項目Noから判別して移行する add sta*/ ";
                tmp_sql = tmp_sql + " 			/*V7には入金項目属性が存在しないため戻しなのか解約費用なのか判別できないため*/ ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_no = 4010 OR KYSQ1.nkin_no = 4020 THEN 1 ";
                tmp_sql = tmp_sql + " 			/*20160520 敷金戻し、保証金戻しの月区分を入金項目Noから判別して移行する add end*/ ";
                tmp_sql = tmp_sql + " 			/*解約時の入金項目は変換が異なる sta*/ ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.ko_no = 999 AND tuki_kbn = 1 AND KYSQ1.nkin_kbn = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.ko_no = 999 AND tuki_kbn = 1 AND KYSQ1.nkin_kbn = 4 THEN 3 ";
                tmp_sql = tmp_sql + " 			/*解約時の入金項目は変換が異なる end*/ ";
                tmp_sql = tmp_sql + " 			ELSE tuki_kbn ";
                tmp_sql = tmp_sql + " 		 END AS [月区分] ";
                tmp_sql = tmp_sql + " 		/*,nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 		,MN1.nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 		,1 AS [入金項目レコードNo] ";
                tmp_sql = tmp_sql + " 		,naibu_no AS [並び順No] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			/*解約時の入金項目は変換が異なる sta*/ ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.ko_no = 999 AND KYSQ1.nkin_kbn = 1 THEN 4	/*解約時*/ ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.ko_no = 999 AND KYSQ1.nkin_kbn = 4 THEN 4	/*解約時*/ ";
                tmp_sql = tmp_sql + " 			/*解約時の入金項目は変換が異なる end*/ ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_kbn = 1 THEN 1	/*毎月*/ ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_kbn = 2 THEN 2	/*契約時*/ ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_kbn = 5 AND koteihendo_kbn = 1 THEN 1	/*毎月*/ ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [入金項目区分] ";
                tmp_sql = tmp_sql + " 		/*20160523 敷金の税区分修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*革命10では敷金に消費税の設定ができないため取得方法を変更*/ ";
                tmp_sql = tmp_sql + " 		/*敷金は2010で固定で設定する*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN sq_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN sq_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		 END AS [税区分] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_no = 2010 AND sq_zeiumu = 1 THEN sq_gak * 1.08 ";
                tmp_sql = tmp_sql + " 			ELSE sq_gak ";
                tmp_sql = tmp_sql + " 		 END AS [請求額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_no = 2010 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_no <> 2010 AND sq_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_no <> 2010 AND sq_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		 END AS [税区分]	 ";
                tmp_sql = tmp_sql + " 		/*20160523 敷金の税区分修正 chg end*/	 ";
                tmp_sql = tmp_sql + " 		,'' AS [請求税額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(tukisu,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [算出区分] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(tukisu,0) <> 0 THEN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT TOP 1 MN2.nkin_name FROM ky_sqdata AS KYSQ2	/*20160823 山川さん指摘対応 内部No重複によるエラー修正 TOP 1を追加*/ ";
                tmp_sql = tmp_sql + " 					LEFT JOIN  m_nkin AS MN2 ON KYSQ2.nkin_no = MN2.nkin_no ";
                tmp_sql = tmp_sql + " 					WHERE KYSQ1.bk_no = KYSQ2.bk_no ";
                tmp_sql = tmp_sql + " 					AND   KYSQ1.hy_no = KYSQ2.hy_no ";
                tmp_sql = tmp_sql + " 					AND   KYSQ1.ky_no = KYSQ2.ky_no ";
                tmp_sql = tmp_sql + " 					AND   KYSQ1.ko_no = KYSQ2.ko_no ";
                tmp_sql = tmp_sql + " 					AND   KYSQ2.tuki_kbn = 2 ";
                tmp_sql = tmp_sql + " 					AND   KYSQ2.nkin_kbn = 1 ";
                tmp_sql = tmp_sql + " 					AND   KYSQ2.naibu_no = 1 ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 		 END AS [算出基準入金項目名] ";
                tmp_sql = tmp_sql + " 		,tukisu AS [算出ヶ月] ";
                tmp_sql = tmp_sql + " 		,kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 		,nkbn_name AS [入金方法] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = -2 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = -1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 0 THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 1 THEN 4 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 2 THEN 5 ";
                tmp_sql = tmp_sql + " 		 END AS [請求月区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [フリーレント適用開始日] ";
                tmp_sql = tmp_sql + " 		,'' AS [フリーレント適用終了日] ";
                tmp_sql = tmp_sql + " 		,'' AS [フリーレント終了月請求額] ";
                tmp_sql = tmp_sql + " 		/*2016.04.26 メインの方へも反映させる修正 -chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*変動費で固定に設定されている入金項目は請求開始月を取得する*/ ";
                tmp_sql = tmp_sql + " 		/*,'' AS [請求開始月]*/ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_kbn = 5 AND koteihendo_kbn = 1 THEN zj_startym ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [請求開始月] ";
                tmp_sql = tmp_sql + " 		/*2016.04.26 メインの方へも反映させる修正 -chg end*/ ";
                tmp_sql = tmp_sql + " 		 ";
                tmp_sql = tmp_sql + " 		/*2016.04.26 メインの方へも反映させる修正 -chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*変動費で固定に設定されている入金項目は請求間隔から値を設定する (他は1を設定)*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN tuki_kbn = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [請求パターン] ";
                tmp_sql = tmp_sql + " 		 */ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_kbn = 5 AND koteihendo_kbn = 1 THEN ";
                tmp_sql = tmp_sql + " 				CASE WHEN zj_sqkan = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 3 THEN 2 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 4 THEN 2 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 5 THEN 1 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 6 THEN 2 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 7 THEN 1 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 8 THEN 1 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 9 THEN 1 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 10 THEN 1 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 11 THEN 1 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan = 12 THEN 2 ";
                tmp_sql = tmp_sql + " 					 WHEN zj_sqkan >= 13 THEN 1 END ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [請求パターン] ";
                tmp_sql = tmp_sql + " 		 /*2016.04.26 メインの方へも反映させる修正 -chg end*/ ";
                tmp_sql = tmp_sql + " 		 /*2016.04.26 メインの方へも反映させる修正 -chg sta*/ ";
                tmp_sql = tmp_sql + " 		 /*変動費で固定に設定されている入金項目は請求間隔を取得する (他は1を設定)*/ ";
                tmp_sql = tmp_sql + " 		 /* ";
                tmp_sql = tmp_sql + " 		,CASE  ";
                tmp_sql = tmp_sql + " 			WHEN tuki_kbn = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [請求間隔] ";
                tmp_sql = tmp_sql + " 		 */ ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN KYSQ1.nkin_kbn = 5 AND koteihendo_kbn = 1 THEN zj_sqkan ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [請求間隔] ";
                tmp_sql = tmp_sql + " 		 /*2016.04.26 メインの方へも反映させる修正 -chg end*/ ";
                tmp_sql = tmp_sql + " 		,NULL AS [請求発生年] ";
                tmp_sql = tmp_sql + " 		,NULL AS [請求発生月] ";
                tmp_sql = tmp_sql + " 		/*20161007 適用税率データ調整処理の追加 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*NULLで抽出し、コンバート後にデータ調整*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN sq_zeiumu = 1 THEN 8 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [適用税率] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		,NULL AS [適用税率] ";
                tmp_sql = tmp_sql + " 		/*20161007 適用税率データ調整処理の追加 chg end*/ ";
                tmp_sql = tmp_sql + " 		,biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,'' AS [フリーレント適用区分] ";
                tmp_sql = tmp_sql + " 	FROM ky_sqdata AS KYSQ1 ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN1 ON KYSQ1.nkin_no = MN1.nkin_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkbn AS MNKBN ON KYSQ1.nkbn = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " 	/*WHERE ko_no <> 999*/ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT bk_no,hy_no,ky_no,ko_no ";
                tmp_sql = tmp_sql + " 				  ,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 			FROM ky_kosinkai AS KYK ";
                tmp_sql = tmp_sql + " 		) AS KYRECNO ";
                tmp_sql = tmp_sql + " 	ON  KYSQ1.bk_no = KYRECNO.bk_no ";
                tmp_sql = tmp_sql + " 	AND KYSQ1.hy_no = KYRECNO.hy_no ";
                tmp_sql = tmp_sql + " 	AND KYSQ1.ky_no = KYRECNO.ky_no ";
                tmp_sql = tmp_sql + " 	AND KYSQ1.ko_no = KYRECNO.ko_no ";
                tmp_sql = tmp_sql + " ) AS KYSQTOTAL ";
                tmp_sql = tmp_sql + " WHERE [入金項目区分] IS NOT NULL ";
                tmp_sql = tmp_sql + " AND   [請求額] <> 0 ";
                tmp_sql = tmp_sql + " /*2016.04.26 メインの方へも反映させる修正 -add sta*/ ";
                tmp_sql = tmp_sql + " AND   [契約管理レコードNo] IS NOT NULL /*原因は不明だが契約入金項目のレコードが存在し契約情報のレコードが存在しないため抽出対象外としておく*/ ";
                tmp_sql = tmp_sql + " /*2016.04.26 メインの方へも反映させる修正 -add end*/ ";

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

    #region 契約次回入金項目情報

    public class Ky_sqdata_nx_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[月区分],[並び順No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[月区分],[入金項目名],[入金項目レコードNo],[並び順No],[入金項目区分],[請求額],[税区分],[請求税額],[算出区分],[算出基準入金項目名],[算出ヶ月],[請求先No],[入金方法],[請求月区分],[請求開始月],[請求パターン],[請求発生間隔],[請求対象年],[請求対象月],[適用税率],[備考]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,tuki_kbn AS [月区分] ";
                tmp_sql = tmp_sql + " 		/*,nkin_no AS [入金項目No]*/ ";
                tmp_sql = tmp_sql + " 		,MN1.nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 		,1 AS [入金項目レコードNo] ";
                tmp_sql = tmp_sql + " 		,naibu_no AS [並び順No] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN KYSQNX1.nkin_kbn = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN KYSQNX1.nkin_kbn = 3 THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN KYSQNX1.nkin_kbn = 5 AND koteihendo_kbn = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [入金項目区分] ";
                tmp_sql = tmp_sql + " 		,sq_gak AS [請求額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN sq_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN sq_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		 END AS [税区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [請求税額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(tukisu,0) <> 0 THEN 2 ";
                tmp_sql = tmp_sql + " 			ELSE 1 ";
                tmp_sql = tmp_sql + " 		 END AS [算出区分] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN ISNULL(tukisu,0) <> 0 THEN ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT TOP 1 MN2.nkin_name FROM ky_sqdata_nx AS KYSQNX2	/*20160823 山川さん指摘対応 内部No重複によるエラー修正 TOP 1を追加*/ ";
                tmp_sql = tmp_sql + " 					LEFT JOIN  m_nkin AS MN2 ON KYSQNX2.nkin_no = MN2.nkin_no ";
                tmp_sql = tmp_sql + " 					WHERE KYSQNX1.bk_no = KYSQNX2.bk_no ";
                tmp_sql = tmp_sql + " 					AND   KYSQNX1.hy_no = KYSQNX2.hy_no ";
                tmp_sql = tmp_sql + " 					AND   KYSQNX1.ky_no = KYSQNX2.ky_no ";
                tmp_sql = tmp_sql + " 					AND   KYSQNX1.ko_no = KYSQNX2.ko_no ";
                tmp_sql = tmp_sql + " 					AND   KYSQNX2.tuki_kbn = 2 ";
                tmp_sql = tmp_sql + " 					AND   KYSQNX2.nkin_kbn = 1 ";
                tmp_sql = tmp_sql + " 					AND   KYSQNX2.naibu_no = 1 ";
                tmp_sql = tmp_sql + " 				) ";
                tmp_sql = tmp_sql + " 		 END AS [算出基準入金項目名] ";
                tmp_sql = tmp_sql + " 		,tukisu AS [算出ヶ月] ";
                tmp_sql = tmp_sql + " 		,kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 		,nkbn_name AS [入金方法] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = -2 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = -1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 0 THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 1 THEN 4 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 2 THEN 5 ";
                tmp_sql = tmp_sql + " 		 END AS [請求月区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [請求開始月] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN tuki_kbn = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [請求パターン] ";
                tmp_sql = tmp_sql + " 		,CASE  ";
                tmp_sql = tmp_sql + " 			WHEN tuki_kbn = 2 THEN 1 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [請求発生間隔] ";
                tmp_sql = tmp_sql + " 		,NULL AS [請求対象年]	 ";
                tmp_sql = tmp_sql + " 		,NULL AS [請求対象月]	 ";
                tmp_sql = tmp_sql + " 		/*20161007 適用税率データ調整処理の追加 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*NULLで抽出し、コンバート後にデータ調整*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN sq_zeiumu = 1 THEN 8 ";
                tmp_sql = tmp_sql + " 			ELSE NULL ";
                tmp_sql = tmp_sql + " 		 END AS [適用税率] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		,NULL AS [適用税率] ";
                tmp_sql = tmp_sql + " 		/*20161007 適用税率データ調整処理の追加 chg end*/ ";
                tmp_sql = tmp_sql + " 		,biko AS [備考] ";
                tmp_sql = tmp_sql + " 	FROM ky_sqdata_nx AS KYSQNX1 ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN1 ON KYSQNX1.nkin_no = MN1.nkin_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkbn AS MNKBN ON KYSQNX1.nkbn = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " 	WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " ) AS KYSQTOTAL ";
                tmp_sql = tmp_sql + " WHERE [入金項目区分] IS NOT NULL ";
                tmp_sql = tmp_sql + " AND   NOT([月区分] = 1 AND [入金項目区分] = 1) ";
                tmp_sql = tmp_sql + " AND   [請求額] <> 0 ";

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

    #region 契約変動費各戸メーター情報

    public class Ky_sqdata_hendo_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[変動費No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[変動費No],[請求対象フラグ],[メーター分類],[メーター名],[入金項目名],[変動費請求ルールNo],[変動費請求ルール備考],[請求先No],[入金方法],[備考],[請求月]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,naibu_no AS [変動費No] ";
                tmp_sql = tmp_sql + " 		,1 AS [請求対象フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [メーター分類] ";
                tmp_sql = tmp_sql + " 		,'' AS [メーター名] ";
                tmp_sql = tmp_sql + " 		,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 		,hendo_no AS [変動費請求ルールNo] ";
                tmp_sql = tmp_sql + " 		,'' AS [変動費請求ルール備考] ";
                tmp_sql = tmp_sql + " 		,kys_no AS [請求先No] ";
                tmp_sql = tmp_sql + " 		,nkbn_name AS [入金方法] ";
                tmp_sql = tmp_sql + " 		,biko AS [備考] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = -2 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = -1 THEN 2 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 0 THEN 3 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 1 THEN 4 ";
                tmp_sql = tmp_sql + " 			WHEN sq_mm = 2 THEN 5 ";
                tmp_sql = tmp_sql + " 		 END AS [請求月]	 ";
                tmp_sql = tmp_sql + " 	FROM ky_sqdata AS KYSQ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN ON KYSQ.nkin_no = MN.nkin_no ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkbn AS MNKBN ON KYSQ.nkbn = MNKBN.nkbn_no ";
                tmp_sql = tmp_sql + " 	WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 	AND   koteihendo_kbn = 2 ";
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

    #region 契約控除ルール情報

    public class Ky_sqdata_kojo_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[対象区分],[表示順]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[対象区分],[控除入金項目名],[表示順],[相殺予定フラグ],[控除額算出基準],[控除額],[控除額税区分],[控除対象入金項目名],[控除率],[控除率税区分],[控除率内税],[適用税率],[立替回収または預り金],[控除税額],[送金対象フラグ],[送金対象入金項目名]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,2 AS [対象区分]	/*革命10の契約時*/ ";
                tmp_sql = tmp_sql + " 		,nkin_name AS [控除入金項目名] ";
                tmp_sql = tmp_sql + " 		,naibu_no AS [表示順] ";
                tmp_sql = tmp_sql + " 		,sosai_umu AS [相殺予定フラグ] ";
                tmp_sql = tmp_sql + " 		,2 AS [控除額算出基準] ";
                tmp_sql = tmp_sql + " 		,so_gak AS [控除額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN so_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN so_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		 END AS [控除額税区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [控除対象入金項目名] ";
                tmp_sql = tmp_sql + " 		,0 AS [控除率] ";
                tmp_sql = tmp_sql + " 		,2 AS [控除率税区分] ";
                tmp_sql = tmp_sql + " 		,0 AS [控除率内税] ";
                tmp_sql = tmp_sql + " 		,NULL AS [適用税率] ";
                tmp_sql = tmp_sql + " 		,0 AS [立替回収または預り金] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN so_zeiumu = 0 THEN 0 ";
                tmp_sql = tmp_sql + " 			WHEN so_zeiumu = 1 THEN so_gak * 0.08 ";
                tmp_sql = tmp_sql + " 		 END AS [控除税額] ";
                tmp_sql = tmp_sql + " 		,0 AS [送金対象フラグ] ";
                tmp_sql = tmp_sql + " 		,NULL AS [送金対象入金項目名] ";
                tmp_sql = tmp_sql + " 	FROM ky_sqdata AS KYSQ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN ON KYSQ.nkin_no = MN.nkin_no ";
                tmp_sql = tmp_sql + " 	WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 	AND   KYSQ.nkin_kbn = 7 ";
                tmp_sql = tmp_sql + " ) AS KYSQKOJO_KY ";
                tmp_sql = tmp_sql + " UNION ";
                tmp_sql = tmp_sql + " SELECT * FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,ko_no AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,3 AS [対象区分]	/*革命10の更新時*/ ";
                tmp_sql = tmp_sql + " 		,nkin_name AS [控除入金項目名] ";
                tmp_sql = tmp_sql + " 		,naibu_no AS [表示順] ";
                tmp_sql = tmp_sql + " 		,sosai_umu AS [相殺予定フラグ] ";
                tmp_sql = tmp_sql + " 		,2 AS [控除額算出基準] ";
                tmp_sql = tmp_sql + " 		,so_gak AS [控除額] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN so_zeiumu = 0 THEN 1 ";
                tmp_sql = tmp_sql + " 			WHEN so_zeiumu = 1 THEN 2 ";
                tmp_sql = tmp_sql + " 		 END AS [控除額税区分] ";
                tmp_sql = tmp_sql + " 		,NULL AS [控除対象入金項目名] ";
                tmp_sql = tmp_sql + " 		,0 AS [控除率] ";
                tmp_sql = tmp_sql + " 		,2 AS [控除率税区分] ";
                tmp_sql = tmp_sql + " 		,0 AS [控除率内税] ";
                tmp_sql = tmp_sql + " 		,NULL AS [適用税率] ";
                tmp_sql = tmp_sql + " 		,0 AS [立替回収または預り金] ";
                tmp_sql = tmp_sql + " 		,CASE ";
                tmp_sql = tmp_sql + " 			WHEN so_zeiumu = 0 THEN 0 ";
                tmp_sql = tmp_sql + " 			WHEN so_zeiumu = 1 THEN so_gak * 0.08 ";
                tmp_sql = tmp_sql + " 		 END AS [控除税額] ";
                tmp_sql = tmp_sql + " 		,0 AS [送金対象フラグ] ";
                tmp_sql = tmp_sql + " 		,NULL AS [送金対象入金項目名] ";
                tmp_sql = tmp_sql + " 	FROM ky_sqdata AS KYSQ ";
                tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN ON KYSQ.nkin_no = MN.nkin_no ";
                tmp_sql = tmp_sql + " 	WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 	AND   KYSQ.nkin_kbn = 7 ";
                tmp_sql = tmp_sql + " ) AS KYSQKOJO_KO ";

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

    #region 契約送金ルール情報

    public class Ky_sqdata_sorule_Repository
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

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[対象区分],[表示順]";

                // 2016.04.26 メインの方へも反映させる修正 -chg sta
                // tmp_sql = tmp_sql & " SELECT * FROM "
                // tmp_sql = tmp_sql & " ( "
                // tmp_sql = tmp_sql & " 	SELECT "
                // tmp_sql = tmp_sql & " 		 [物件No] "
                // tmp_sql = tmp_sql & " 		,[部屋No] "
                // tmp_sql = tmp_sql & " 		,[契約No] "
                // tmp_sql = tmp_sql & " 		,[契約管理レコードNo] "
                // tmp_sql = tmp_sql & " 		,[対象区分] "
                // tmp_sql = tmp_sql & " 		,[入金項目名] "
                // tmp_sql = tmp_sql & " 		,ROW_NUMBER()OVER(PARTITION BY [物件No],[部屋No],[契約No],[契約管理レコードNo],[対象区分] ORDER BY [内部No]) AS [表示順] "
                // tmp_sql = tmp_sql & " 		,[送金率] "
                // tmp_sql = tmp_sql & " 		,[管理手数料率] "
                // tmp_sql = tmp_sql & " 		,[滞納保証有無] "
                // tmp_sql = tmp_sql & " 	FROM "
                // tmp_sql = tmp_sql & " 	( "
                // tmp_sql = tmp_sql & " 		SELECT "
                // tmp_sql = tmp_sql & " 			 KYSQ.bk_no AS [物件No] "
                // tmp_sql = tmp_sql & " 			,KYSQ.hy_no AS [部屋No] "
                // tmp_sql = tmp_sql & " 			,KYSQ.ky_no AS [契約No] "
                // tmp_sql = tmp_sql & " 			,[契約管理レコードNo] "
                // tmp_sql = tmp_sql & " 			,CASE "
                // tmp_sql = tmp_sql & " 				WHEN KYSQ.ko_no = 999 AND KYSQ.nkin_kbn = 1 THEN 4	/*解約時*/ "
                // tmp_sql = tmp_sql & " 				WHEN KYSQ.ko_no = 999 AND KYSQ.nkin_kbn = 4 THEN 6	/*解約費*/ "
                // tmp_sql = tmp_sql & " 				WHEN KYSQ.nkin_kbn = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 				WHEN KYSQ.nkin_kbn = 2 THEN 2 "
                // tmp_sql = tmp_sql & " 				WHEN KYSQ.nkin_kbn = 3 THEN 3 "
                // tmp_sql = tmp_sql & " 				WHEN KYSQ.nkin_kbn = 5 AND koteihendo_kbn = 1 THEN 1 "
                // tmp_sql = tmp_sql & " 				WHEN KYSQ.nkin_kbn = 5 AND koteihendo_kbn = 2 THEN 5 "
                // tmp_sql = tmp_sql & " 				ELSE NULL "
                // tmp_sql = tmp_sql & " 			 END AS [対象区分] "
                // tmp_sql = tmp_sql & " 			,nkin_name AS [入金項目名] "
                // tmp_sql = tmp_sql & " 			,naibu_no AS [内部No] "
                // tmp_sql = tmp_sql & " 			,sokin_rit AS [送金率] "
                // tmp_sql = tmp_sql & " 			,[管理手数料率] "
                // tmp_sql = tmp_sql & " 			,tnh_umu AS [滞納保証有無] "
                // tmp_sql = tmp_sql & " 			/*V7で翌月分以降に入金項目を設定した際に当月分に内容が反映されるが管理手数料は反映されないため*/ "
                // tmp_sql = tmp_sql & " 			/*翌月分以降の入金項目を送金ルールへの移行対象とする*/ "
                // tmp_sql = tmp_sql & " 			,CASE "
                // tmp_sql = tmp_sql & " 				WHEN KYSQ.ko_no <> 999 AND KYSQ.nkin_kbn = 1 AND KYSQ.tuki_kbn = 1 THEN '対象外' "
                // tmp_sql = tmp_sql & " 				ELSE '対象' "
                // tmp_sql = tmp_sql & " 			 END AS [移行対象区分] "
                // tmp_sql = tmp_sql & " 		FROM ky_sqdata AS KYSQ "
                // tmp_sql = tmp_sql & " 		LEFT JOIN m_nkin AS MN ON KYSQ.nkin_no = MN.nkin_no "
                // tmp_sql = tmp_sql & " 		LEFT JOIN "
                // tmp_sql = tmp_sql & " 			( "
                // tmp_sql = tmp_sql & " 				/*管理手数料率をマスタ化 sta*/ "
                // tmp_sql = tmp_sql & " 				SELECT * FROM "
                // tmp_sql = tmp_sql & " 				( "
                // tmp_sql = tmp_sql & " 					SELECT bk_no,kn_no,1 AS [管理手数料率No],kanritesu_rit1 AS [管理手数料率] FROM bk_kanri "
                // tmp_sql = tmp_sql & " 					UNION "
                // tmp_sql = tmp_sql & " 					SELECT bk_no,kn_no,2 AS [管理手数料率No],kanritesu_rit2 AS [管理手数料率] FROM bk_kanri "
                // tmp_sql = tmp_sql & " 					UNION "
                // tmp_sql = tmp_sql & " 					SELECT bk_no,kn_no,3 AS [管理手数料率No],kanritesu_rit3 AS [管理手数料率] FROM bk_kanri "
                // tmp_sql = tmp_sql & " 					UNION "
                // tmp_sql = tmp_sql & " 					SELECT bk_no,kn_no,4 AS [管理手数料率No],kanritesu_rit4 AS [管理手数料率] FROM bk_kanri "
                // tmp_sql = tmp_sql & " 					UNION "
                // tmp_sql = tmp_sql & " 					SELECT bk_no,kn_no,5 AS [管理手数料率No],kanritesu_rit5 AS [管理手数料率] FROM bk_kanri "
                // tmp_sql = tmp_sql & " 				) AS VW "
                // tmp_sql = tmp_sql & " 				WHERE [管理手数料率] <> 0 "
                // tmp_sql = tmp_sql & " 				/*管理手数料率をマスタ化 end*/ "
                // tmp_sql = tmp_sql & " 			) AS KANRIRYOMST "
                // tmp_sql = tmp_sql & " 		ON  KYSQ.bk_no = KANRIRYOMST.bk_no "
                // tmp_sql = tmp_sql & " 		AND KYSQ.kn_ritno = KANRIRYOMST.[管理手数料率No] "
                // tmp_sql = tmp_sql & " 		LEFT JOIN bk_kanri AS BKK "
                // tmp_sql = tmp_sql & " 		ON KYSQ.bk_no = BKK.bk_no "
                // tmp_sql = tmp_sql & " 		/*WHERE ko_no <> 999*/ "
                // tmp_sql = tmp_sql & " 		LEFT JOIN "
                // tmp_sql = tmp_sql & " 			( "
                // tmp_sql = tmp_sql & " 				SELECT bk_no,hy_no,ky_no,ko_no "
                // tmp_sql = tmp_sql & " 					  ,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約管理レコードNo] "
                // tmp_sql = tmp_sql & " 				FROM ky_kosinkai AS KYK "
                // tmp_sql = tmp_sql & " 			) AS KYRECNO "
                // tmp_sql = tmp_sql & " 		ON  KYSQ.bk_no = KYRECNO.bk_no "
                // tmp_sql = tmp_sql & " 		AND KYSQ.hy_no = KYRECNO.hy_no "
                // tmp_sql = tmp_sql & " 		AND KYSQ.ky_no = KYRECNO.ky_no "
                // tmp_sql = tmp_sql & " 		AND KYSQ.ko_no = KYRECNO.ko_no	 "
                // tmp_sql = tmp_sql & " 	) AS KYSORULE "
                // tmp_sql = tmp_sql & " 	WHERE [対象区分] IS NOT NULL "
                // tmp_sql = tmp_sql & " 	AND   [移行対象区分] = '対象' "
                // tmp_sql = tmp_sql & " ) AS VW "

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[対象区分],[入金項目名],[表示順],[送金率],[管理手数料率],[滞納保証有無]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,[契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,[対象区分] ";
                tmp_sql = tmp_sql + " 		,[入金項目名] ";
                tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no,[契約管理レコードNo],[対象区分] ORDER BY [内部No]) AS [表示順] ";
                tmp_sql = tmp_sql + " 		/*20160613 送金率が移行されていない現象の対応 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,[送金率]*/ ";
                tmp_sql = tmp_sql + " 		,ISNULL([送金率],0) AS [送金率] ";
                tmp_sql = tmp_sql + " 		/*20160613 送金率が移行されていない現象の対応 chg end*/ ";
                tmp_sql = tmp_sql + " 		,[管理手数料率] ";
                tmp_sql = tmp_sql + " 		,tnh_umu AS [滞納保証有無] ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		SELECT ";
                tmp_sql = tmp_sql + " 			 bk_no ";
                tmp_sql = tmp_sql + " 			,hy_no ";
                tmp_sql = tmp_sql + " 			,ky_no ";
                tmp_sql = tmp_sql + " 			,[契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 			,CASE ";
                tmp_sql = tmp_sql + " 				WHEN ko_no = 999 AND nkin_kbn = 1 THEN 4	/*解約時*/ ";
                tmp_sql = tmp_sql + " 				WHEN ko_no = 999 AND nkin_kbn = 4 THEN 6	/*解約費*/ ";
                tmp_sql = tmp_sql + " 				WHEN nkin_kbn = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 				WHEN nkin_kbn = 2 THEN 2 ";
                tmp_sql = tmp_sql + " 				WHEN nkin_kbn = 3 THEN 3 ";
                tmp_sql = tmp_sql + " 				WHEN nkin_kbn = 5 AND koteihendo_kbn = 1 THEN 1 ";
                tmp_sql = tmp_sql + " 				WHEN nkin_kbn = 5 AND koteihendo_kbn = 2 THEN 5 ";
                tmp_sql = tmp_sql + " 				ELSE NULL ";
                tmp_sql = tmp_sql + " 			 END AS [対象区分] ";
                tmp_sql = tmp_sql + " 			,nkin_name AS [入金項目名] ";
                tmp_sql = tmp_sql + " 			,naibu_no AS [内部No] ";
                tmp_sql = tmp_sql + " 			,sokin_rit AS [送金率] ";
                tmp_sql = tmp_sql + " 			,[管理手数料率] ";
                tmp_sql = tmp_sql + " 			,tnh_umu ";
                tmp_sql = tmp_sql + " 		FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT VW2.*,BKK.tnh_umu FROM ";
                tmp_sql = tmp_sql + " 			( ";
                tmp_sql = tmp_sql + " 				SELECT * FROM ";
                tmp_sql = tmp_sql + " 				( ";
                tmp_sql = tmp_sql + " 					SELECT ";
                tmp_sql = tmp_sql + " 						 KYSQ.* ";
                tmp_sql = tmp_sql + " 						,KYK.ky_start_ymd ";
                tmp_sql = tmp_sql + " 						,KYK.ky_end_ymd ";
                tmp_sql = tmp_sql + " 						,KYK.契約管理レコードNo ";
                tmp_sql = tmp_sql + " 						,KANRIRYOMST.管理手数料率 ";
                tmp_sql = tmp_sql + " 						,KANRIRYOMST.start_ym ";
                tmp_sql = tmp_sql + " 						,KANRIRYOMST.end_ym ";
                tmp_sql = tmp_sql + " 					FROM ";
                tmp_sql = tmp_sql + " 					( ";
                tmp_sql = tmp_sql + " 						SELECT * FROM ky_sqdata ";
                tmp_sql = tmp_sql + " 						/*20160613 送金ルールで毎月以外移行されていない現象の対応 chg sta*/ ";
                tmp_sql = tmp_sql + " 						/*WHERE (ko_no <> 999 AND nkin_kbn = 1 AND tuki_kbn = 1)*/ ";
                tmp_sql = tmp_sql + " 						WHERE ko_no <> 999 ";
                tmp_sql = tmp_sql + " 						AND NOT (nkin_kbn = 1 AND tuki_kbn = 1) ";
                tmp_sql = tmp_sql + " 						/*20160613 送金ルールで毎月以外移行されていない現象の対応 chg end*/ ";
                tmp_sql = tmp_sql + " 					) AS KYSQ ";
                tmp_sql = tmp_sql + " 					LEFT JOIN ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							SELECT ";
                tmp_sql = tmp_sql + " 								 bk_no ";
                tmp_sql = tmp_sql + " 								,hy_no ";
                tmp_sql = tmp_sql + " 								,ky_no ";
                tmp_sql = tmp_sql + " 								,ko_no ";
                tmp_sql = tmp_sql + " 								,ky_start_ymd ";
                tmp_sql = tmp_sql + " 								,ky_end_ymd ";
                tmp_sql = tmp_sql + " 								,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 							FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 						) AS KYK ";
                tmp_sql = tmp_sql + " 					ON  KYSQ.bk_no = KYK.bk_no ";
                tmp_sql = tmp_sql + " 					AND KYSQ.hy_no = KYK.hy_no ";
                tmp_sql = tmp_sql + " 					AND KYSQ.ky_no = KYK.ky_no ";
                tmp_sql = tmp_sql + " 					AND KYSQ.ko_no = KYK.ko_no ";
                tmp_sql = tmp_sql + " 					LEFT JOIN ";
                tmp_sql = tmp_sql + " 						( ";
                tmp_sql = tmp_sql + " 							/*管理手数料率をマスタ化 sta*/ ";
                tmp_sql = tmp_sql + " 							SELECT * FROM ";
                tmp_sql = tmp_sql + " 							( ";
                tmp_sql = tmp_sql + " 								SELECT bk_no,kn_no,1 AS [管理手数料率No],kanritesu_rit1 AS [管理手数料率],start_ym,end_ym FROM bk_kanri ";
                tmp_sql = tmp_sql + " 								UNION ";
                tmp_sql = tmp_sql + " 								SELECT bk_no,kn_no,2 AS [管理手数料率No],kanritesu_rit2 AS [管理手数料率],start_ym,end_ym FROM bk_kanri ";
                tmp_sql = tmp_sql + " 								UNION ";
                tmp_sql = tmp_sql + " 								SELECT bk_no,kn_no,3 AS [管理手数料率No],kanritesu_rit3 AS [管理手数料率],start_ym,end_ym FROM bk_kanri ";
                tmp_sql = tmp_sql + " 								UNION ";
                tmp_sql = tmp_sql + " 								SELECT bk_no,kn_no,4 AS [管理手数料率No],kanritesu_rit4 AS [管理手数料率],start_ym,end_ym FROM bk_kanri ";
                tmp_sql = tmp_sql + " 								UNION ";
                tmp_sql = tmp_sql + " 								SELECT bk_no,kn_no,5 AS [管理手数料率No],kanritesu_rit5 AS [管理手数料率],start_ym,end_ym FROM bk_kanri ";
                tmp_sql = tmp_sql + " 							) AS VW ";
                tmp_sql = tmp_sql + " 							WHERE [管理手数料率] <> 0 ";
                tmp_sql = tmp_sql + " 							/*管理手数料率をマスタ化 end*/ ";
                tmp_sql = tmp_sql + " 						) AS KANRIRYOMST ";
                tmp_sql = tmp_sql + " 					ON  KYSQ.bk_no = KANRIRYOMST.bk_no ";
                tmp_sql = tmp_sql + " 					AND KYSQ.kn_ritno = KANRIRYOMST.[管理手数料率No] ";
                tmp_sql = tmp_sql + " 				) AS VW ";
                tmp_sql = tmp_sql + " 				WHERE (kn_ritno = 0 OR ky_end_ymd BETWEEN VW.start_ym AND VW.end_ym) ";
                tmp_sql = tmp_sql + " 			) AS VW2 ";
                tmp_sql = tmp_sql + " 			LEFT JOIN (SELECT bk_no,start_ym,end_ym,tnh_umu FROM bk_kanri) AS BKK ";
                tmp_sql = tmp_sql + " 			ON  VW2.bk_no = BKK.bk_no ";
                tmp_sql = tmp_sql + " 			WHERE ky_end_ymd BETWEEN BKK.start_ym AND BKK.end_ym ";
                tmp_sql = tmp_sql + " 		) AS VW3 ";
                tmp_sql = tmp_sql + " 		LEFT JOIN (SELECT nkin_no,nkin_name FROM m_nkin) AS MNKIN ON VW3.nkin_no = MNKIN.nkin_no ";
                tmp_sql = tmp_sql + " 	) AS VW4 ";
                tmp_sql = tmp_sql + " 	/*20160613 送金ルールで毎月以外移行されていない現象の対応 add*/ ";
                tmp_sql = tmp_sql + " 	WHERE [対象区分] IS NOT NULL ";
                tmp_sql = tmp_sql + " ) AS VW5 ";

                // 2016.04.26 メインの方へも反映させる修正 -chg end
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

    #region 契約解約情報

    public class Ky_kosinkai_kai_Repository
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

                // 20160629 修繕検証後修正 -add
                bool szenopflg = false;

                string tmp_sql = "";

                sortstr = "[物件No],[部屋No],[契約No],[契約管理レコードNo]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[契約管理レコードNo],[解約日],[解約精算費用決定日],[解約精算業務完了日],[解約受付日],[解約理由],[退去日],[受付担当者No],[立会担当者No],[精算担当者No],[修繕担当者No],[自社・支店No],[修繕完了日],[工事予定期間開始日],[工事予定期間終了日],[工事場所],[工事概要],[退去後宛名],[退去後宛名Shift-jis],[退去後郵便番号],[退去後住所①],[退去後住所②],[退去後電話番号①],[不足時入金口座],[備考(解約精算)],[備考(立会)],[返金予定日],[請求締め日],[管理手数料フラグ],[管理手数料額],[送金予定日],[立会い],[退去予定日],[募集条件確認日],[家主連絡担当者No],[募集条件内容],[退去予定時刻],[退去後敬称],[解約月賃料は扱わないフラグ],[契約者修繕請求先No],[契約者修繕入金区分],[契約者修繕送金率],[契約者修繕送金額],[家主修繕控除先No],[家主修繕控除請求区分],[家主修繕控除先送金ルールGuid],[家主修繕控除先送金ルールNo],[家主修繕控除先口座No],[家主修繕請求締日],[家主修繕請求入金区分],[家主修繕請求振込先口座No],[家主修繕請求先No],[家主修繕控除予定日],[家主修繕請求書発行予定日]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                tmp_sql = tmp_sql + " 	SELECT ";
                tmp_sql = tmp_sql + " 		 KYTOTAL.bk_no AS [物件No] ";
                tmp_sql = tmp_sql + " 		,KYTOTAL.hy_no AS [部屋No] ";
                tmp_sql = tmp_sql + " 		,KYTOTAL.ky_no AS [契約No] ";
                tmp_sql = tmp_sql + " 		,[革命10用解約レコードNo] AS [契約管理レコードNo] ";
                tmp_sql = tmp_sql + " 		,kai_ymd AS [解約日] ";
                tmp_sql = tmp_sql + " 		/*20160613 10解約確定日にV7の解約精算日を移行するように修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/*,kai_ymd AS [解約精算費用決定日]*/ ";
                tmp_sql = tmp_sql + " 		,seisan_yoteiymd AS [解約精算費用決定日] ";
                tmp_sql = tmp_sql + " 		/*20160613 10解約確定日にV7の解約精算日を移行するように修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		,seisan_yoteiymd AS [解約精算業務完了日] ";
                tmp_sql = tmp_sql + " 		,'' AS [解約受付日] ";
                tmp_sql = tmp_sql + " 		,taikyo_riyuu1 AS [解約理由] ";
                tmp_sql = tmp_sql + " 		,tai_yoteiymd AS [退去日] ";
                tmp_sql = tmp_sql + " 		,tachiai AS [受付担当者No] ";
                tmp_sql = tmp_sql + " 		,tachiai AS [立会担当者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [精算担当者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕担当者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [自社・支店No] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕完了日] ";
                tmp_sql = tmp_sql + " 		,'' AS [工事予定期間開始日] ";
                tmp_sql = tmp_sql + " 		,'' AS [工事予定期間終了日] ";
                tmp_sql = tmp_sql + " 		,'' AS [工事場所] ";
                tmp_sql = tmp_sql + " 		,'' AS [工事概要] ";
                tmp_sql = tmp_sql + " 		,tenk_name AS [退去後宛名] ";
                tmp_sql = tmp_sql + " 		,'' AS [退去後宛名Shift-jis] ";
                tmp_sql = tmp_sql + " 		,tenk_post AS [退去後郵便番号] ";
                tmp_sql = tmp_sql + " 		,tenk_add1 AS [退去後住所①] ";
                tmp_sql = tmp_sql + " 		,tenk_add2 AS [退去後住所②] ";
                tmp_sql = tmp_sql + " 		,tenk_tel1 AS [退去後電話番号①] ";
                tmp_sql = tmp_sql + " 		,fkom_no AS [不足時入金口座] ";
                tmp_sql = tmp_sql + " 		,'' AS [備考(解約精算)] ";
                tmp_sql = tmp_sql + " 		,kai_biko AS [備考(立会)] ";
                tmp_sql = tmp_sql + " 		,seisan_ymd AS [返金予定日] ";
                tmp_sql = tmp_sql + " 		,'' AS [請求締め日] ";
                tmp_sql = tmp_sql + " 		,0 AS [管理手数料フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [管理手数料額] ";
                tmp_sql = tmp_sql + " 		,seisan_soymd AS [送金予定日] ";
                tmp_sql = tmp_sql + " 		,'' AS [立会い] ";
                tmp_sql = tmp_sql + " 		,tai_yoteiymd AS [退去予定日] ";
                tmp_sql = tmp_sql + " 		,'' AS [募集条件確認日] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主連絡担当者No] ";
                tmp_sql = tmp_sql + " 		,'' AS [募集条件内容] ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約解約情報 del sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,'' AS [確認事項入力責任者] ";
                tmp_sql = tmp_sql + " 		,'' AS [確認事項確認日] ";
                tmp_sql = tmp_sql + " 		,'' AS [印刷時の確認事項] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約解約情報 del end*/ ";
                tmp_sql = tmp_sql + " 		/*20160620 EXEUpdateに伴う修正2 del sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		,2 AS [修繕自社負担フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕自社負担額] ";
                tmp_sql = tmp_sql + " 		,-1 AS [修繕自社負担税区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕自社負担税額] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕自社負担備考] ";
                tmp_sql = tmp_sql + " 		,2 AS [修繕その他負担フラグ] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕その他負担額] ";
                tmp_sql = tmp_sql + " 		,-1 AS [修繕その他負担税区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕その他負担税額] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕その他負担者] ";
                tmp_sql = tmp_sql + " 		,'' AS [修繕その他負担備考] ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		/*20160620 EXEUpdateに伴う修正2 del end*/ ";
                tmp_sql = tmp_sql + " 		,tai_yoteitimestr AS [退去予定時刻] ";
                tmp_sql = tmp_sql + " 		,'' AS [退去後敬称] ";
                tmp_sql = tmp_sql + " 		/*20160519 EXEUpdateに伴う修正 契約解約情報 add*/ ";
                tmp_sql = tmp_sql + " 		,2 AS [解約月賃料は扱わないフラグ] ";
                tmp_sql = tmp_sql + " 		 ";
                tmp_sql = tmp_sql + " 		 ";
                tmp_sql = tmp_sql + " 		/*20160629 修繕検証後修正 chg sta*/ ";
                tmp_sql = tmp_sql + " 		/* ";
                tmp_sql = tmp_sql + " 		/*20160620 EXEUpdateに伴う修正2 add sta*/ ";
                tmp_sql = tmp_sql + " 		,'' AS [契約者修繕請求先No] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約者修繕入金区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約者修繕送金率] ";
                tmp_sql = tmp_sql + " 		,'' AS [契約者修繕送金額] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除先No] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除請求区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除先送金ルールGuid] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除先送金ルールNo] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕請求締日] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕請求入金区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕請求振込先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕請求先No]		 ";
                tmp_sql = tmp_sql + " 		/*20160620 EXEUpdateに伴う修正2 add end*/ ";
                tmp_sql = tmp_sql + " 		*/ ";
                tmp_sql = tmp_sql + " 		,KAISZENDATA.kys_no AS [契約者修繕請求先No] ";
                tmp_sql = tmp_sql + " 		,KAISZENDATA.nkbn_name AS [契約者修繕入金区分] ";
                tmp_sql = tmp_sql + " 		,KAISZENDATA.so_rit AS [契約者修繕送金率] ";
                tmp_sql = tmp_sql + " 		,KAISZENDATA.so_gak AS [契約者修繕送金額] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除先No] ";
                tmp_sql = tmp_sql + " 		,1 AS [家主修繕控除請求区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除先送金ルールGuid] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除先送金ルールNo] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕控除先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕請求締日] ";
                tmp_sql = tmp_sql + " 		,-1 AS [家主修繕請求入金区分] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕請求振込先口座No] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕請求先No]	 ";
                tmp_sql = tmp_sql + " 		/*20160629 修繕検証後修正 chg end*/ ";
                tmp_sql = tmp_sql + " 		/*20160829 革命10バージョンアップに伴う修正 add sta*/ ";
                tmp_sql = tmp_sql + " 		/*家主修繕控除請求区分が固定で1(控除・別途請求区分で控除)なので控除予定日に値を設定*/ ";
                tmp_sql = tmp_sql + " 		,seisan_soymd AS [家主修繕控除予定日] ";
                tmp_sql = tmp_sql + " 		,'' AS [家主修繕請求書発行予定日] ";
                tmp_sql = tmp_sql + " 		/*20160829 革命10バージョンアップに伴う修正 add end*/ ";
                tmp_sql = tmp_sql + " 	FROM ";
                tmp_sql = tmp_sql + " 	( ";
                tmp_sql = tmp_sql + " 		/*更新Noへ連番を振り更新No = 999 に振られた連番の値を契約レコード管理Noとする sta*/ ";
                tmp_sql = tmp_sql + " 		/*抽出元の作成*/ ";
                tmp_sql = tmp_sql + " 		SELECT * FROM ";
                tmp_sql = tmp_sql + " 		( ";
                tmp_sql = tmp_sql + " 			SELECT ";
                tmp_sql = tmp_sql + " 				 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY bk_no,hy_no,ky_no,ko_no) AS [革命10用解約レコードNo] ";
                tmp_sql = tmp_sql + " 				,* ";
                tmp_sql = tmp_sql + " 			FROM ky_kosinkai ";
                tmp_sql = tmp_sql + " 		) AS VW ";
                tmp_sql = tmp_sql + " 		WHERE ko_no = 999 ";
                tmp_sql = tmp_sql + " 		/*更新Noへ連番を振り更新No = 999 に振られた連番の値を契約レコード管理Noとする end*/ ";
                tmp_sql = tmp_sql + " 	) AS KYTOTAL ";
                // 20160629 修繕検証後修正 -add
                tmp_sql = tmp_sql + Get_UseQry_SzenOp(szenopflg);
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

            /// <summary>
            /// 修繕OPの有無で取得元を分岐 '20160629 修繕検証後修正
            /// 修繕OP：有→リフォームデータから
            /// 修繕OP：無→契約請求データから
            /// </summary>
            /// <param name="opflg"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_SzenOp(bool opflg)
            {

                string tmp_sql = "";

                if (opflg)
                {
                    tmp_sql = tmp_sql + " 	LEFT JOIN ";
                    tmp_sql = tmp_sql + " 		( ";
                    tmp_sql = tmp_sql + " 			SELECT ";
                    tmp_sql = tmp_sql + " 				 SQSAKITOTAL.* ";
                    tmp_sql = tmp_sql + " 				,SOGAKTOTAL.so_gak ";
                    tmp_sql = tmp_sql + " 				,SOGAKTOTAL.so_rit ";
                    tmp_sql = tmp_sql + " 				,nkbn_name ";
                    tmp_sql = tmp_sql + " 			FROM ";
                    tmp_sql = tmp_sql + " 			( ";
                    tmp_sql = tmp_sql + " 				/*リフォームデータの請求先取得 sta*/ ";
                    tmp_sql = tmp_sql + " 				SELECT * FROM ";
                    tmp_sql = tmp_sql + " 				( ";
                    tmp_sql = tmp_sql + " 					SELECT ";
                    tmp_sql = tmp_sql + " 						 REFDATA.bk_no ";
                    tmp_sql = tmp_sql + " 						,REFDATA.hy_composite AS hy_no ";
                    tmp_sql = tmp_sql + " 						,REFDATA.ky_no ";
                    tmp_sql = tmp_sql + " 						,KYKBASE.[契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 						,KYKBASE.kys_no ";
                    tmp_sql = tmp_sql + " 						,REFDATA.kari_kbn ";
                    tmp_sql = tmp_sql + " 					FROM reform_data AS REFDATA ";
                    tmp_sql = tmp_sql + " 					LEFT JOIN ";
                    tmp_sql = tmp_sql + " 						( ";
                    tmp_sql = tmp_sql + " 							SELECT ";
                    tmp_sql = tmp_sql + " 								 * ";
                    tmp_sql = tmp_sql + " 								,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 							FROM ky_kosinkai AS KYK ";
                    tmp_sql = tmp_sql + " 						) AS KYKBASE ";
                    tmp_sql = tmp_sql + " 					ON  REFDATA.bk_no = KYKBASE.bk_no ";
                    tmp_sql = tmp_sql + " 					AND REFDATA.hy_composite = KYKBASE.hy_no ";
                    tmp_sql = tmp_sql + " 					AND REFDATA.ky_no = KYKBASE.ky_no ";
                    tmp_sql = tmp_sql + " 					AND REFDATA.ko_no = KYKBASE.ko_no ";
                    tmp_sql = tmp_sql + " 					LEFT JOIN ";
                    tmp_sql = tmp_sql + " 						( ";
                    tmp_sql = tmp_sql + " 							SELECT * FROM ";
                    tmp_sql = tmp_sql + " 							( ";
                    tmp_sql = tmp_sql + " 								SELECT ";
                    tmp_sql = tmp_sql + " 									 * ";
                    tmp_sql = tmp_sql + " 									,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY naibu_no) AS [請求先抽出用] ";
                    tmp_sql = tmp_sql + " 								FROM ky_sqdata ";
                    tmp_sql = tmp_sql + " 							) AS VW ";
                    tmp_sql = tmp_sql + " 							WHERE [請求先抽出用] = 1 ";
                    tmp_sql = tmp_sql + " 						) AS KYSQ ";
                    tmp_sql = tmp_sql + " 					ON  REFDATA.bk_no = KYSQ.bk_no ";
                    tmp_sql = tmp_sql + " 					AND REFDATA.hy_composite = KYSQ.hy_no ";
                    tmp_sql = tmp_sql + " 					AND REFDATA.ky_no = KYSQ.ky_no ";
                    tmp_sql = tmp_sql + " 					AND REFDATA.ko_no = KYSQ.ko_no ";
                    tmp_sql = tmp_sql + " 					WHERE REFDATA.reform_kbn = 0 ";
                    tmp_sql = tmp_sql + " 					AND   REFDATA.ko_no = 999 ";
                    tmp_sql = tmp_sql + " 				) AS REFSQSAKITOTAL ";
                    tmp_sql = tmp_sql + " 				/*リフォームデータの請求先取得 end*/ ";
                    tmp_sql = tmp_sql + " 			) AS SQSAKITOTAL ";
                    tmp_sql = tmp_sql + " 			LEFT JOIN m_nkbn AS MNKBN ON SQSAKITOTAL.kari_kbn = MNKBN.nkbn_no ";
                    tmp_sql = tmp_sql + " 			LEFT JOIN ";
                    tmp_sql = tmp_sql + " 				( ";
                    tmp_sql = tmp_sql + " 					SELECT * FROM ";
                    tmp_sql = tmp_sql + " 					( ";
                    tmp_sql = tmp_sql + " 						/*リフォームデータから送金率取得 sta*/ ";
                    tmp_sql = tmp_sql + " 						SELECT * FROM ";
                    tmp_sql = tmp_sql + " 						( ";
                    tmp_sql = tmp_sql + " 							SELECT ";
                    tmp_sql = tmp_sql + " 								 REFDATA.bk_no ";
                    tmp_sql = tmp_sql + " 								,REFDATA.hy_composite AS hy_no ";
                    tmp_sql = tmp_sql + " 								,REFDATA.ky_no ";
                    tmp_sql = tmp_sql + " 								,[契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 								,REFSOGAK.so_gak ";
                    tmp_sql = tmp_sql + " 								,REFSOGAK.so_rit_sogo AS so_rit ";
                    tmp_sql = tmp_sql + " 							FROM reform_data AS REFDATA ";
                    tmp_sql = tmp_sql + " 							LEFT JOIN ";
                    tmp_sql = tmp_sql + " 								( ";
                    tmp_sql = tmp_sql + " 									SELECT ";
                    tmp_sql = tmp_sql + " 										 bk_no ";
                    tmp_sql = tmp_sql + " 										,hy_composite ";
                    tmp_sql = tmp_sql + " 										,reform_id ";
                    tmp_sql = tmp_sql + " 										,kari_so_gak + zei_gak_kari_so AS so_gak ";
                    tmp_sql = tmp_sql + " 										,CASE ";
                    tmp_sql = tmp_sql + " 											WHEN ISNULL(total_kari,0) = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 											/*20160629 修繕検証後修正 chg sta*/ ";
                    tmp_sql = tmp_sql + " 											/*ELSE ROUND(((kari_so_gak + zei_gak_kari_so) / (total_kari + zei_gak_kari)) * 100,2)*/ ";
                    tmp_sql = tmp_sql + " 											ELSE ";
                    tmp_sql = tmp_sql + " 												CASE ";
                    tmp_sql = tmp_sql + " 													WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND(((kari_so_gak + zei_gak_kari_so) / (total_kari + zei_gak_kari)) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 													WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND(((kari_so_gak + zei_gak_kari_so) / (total_kari + zei_gak_kari)) * 100,2)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + "     												WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING((((kari_so_gak + zei_gak_kari_so) / (total_kari + zei_gak_kari)) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 												END ";
                    tmp_sql = tmp_sql + " 											/*20160629 修繕検証後修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 										 END AS so_rit_sogo ";
                    tmp_sql = tmp_sql + " 									FROM ";
                    tmp_sql = tmp_sql + " 									( ";
                    tmp_sql = tmp_sql + " 									SELECT ";
                    tmp_sql = tmp_sql + " 										 REFITEMSUM.bk_no ";
                    tmp_sql = tmp_sql + " 										,REFITEMSUM.hy_composite ";
                    tmp_sql = tmp_sql + " 										,REFITEMSUM.reform_id ";
                    tmp_sql = tmp_sql + " 										,REFITEMSUM.total_kari ";
                    tmp_sql = tmp_sql + " 										,CASE ";
                    tmp_sql = tmp_sql + " 											WHEN REFITEMTAX.bk_no IS NULL THEN REFITEMSUM.zei_gak_kari ";
                    tmp_sql = tmp_sql + " 											ELSE REFITEMTAX.zei_gak_kari ";
                    tmp_sql = tmp_sql + " 										 END AS zei_gak_kari ";
                    tmp_sql = tmp_sql + " 										 ,REFITEMSUM.kari_so_gak ";
                    tmp_sql = tmp_sql + " 										,CASE ";
                    tmp_sql = tmp_sql + " 											WHEN REFITEMTAX.bk_no IS NULL THEN 0 ";
                    tmp_sql = tmp_sql + " 											ELSE REFITEMTAX.kari_so_zeigak ";
                    tmp_sql = tmp_sql + " 										 END AS zei_gak_kari_so	 ";
                    tmp_sql = tmp_sql + " 									FROM ";
                    tmp_sql = tmp_sql + " 									( ";
                    tmp_sql = tmp_sql + " 										/*合計の総和と送金額の総和を抽出 sta*/ ";
                    tmp_sql = tmp_sql + " 										SELECT ";
                    tmp_sql = tmp_sql + " 											 bk_no ";
                    tmp_sql = tmp_sql + " 											,hy_composite ";
                    tmp_sql = tmp_sql + " 											,reform_id ";
                    tmp_sql = tmp_sql + " 											,SUM(ISNULL(total_kari,0)) AS total_kari ";
                    tmp_sql = tmp_sql + " 											,SUM(ISNULL(zei_gak_kari,0)) AS zei_gak_kari ";
                    tmp_sql = tmp_sql + " 											,SUM(ISNULL(kari_so_gak,0)) AS kari_so_gak ";
                    tmp_sql = tmp_sql + " 										FROM ";
                    tmp_sql = tmp_sql + " 										( ";
                    tmp_sql = tmp_sql + " 											SELECT * FROM reform_item ";
                    tmp_sql = tmp_sql + " 											WHERE nkin_no <> 8999 ";
                    tmp_sql = tmp_sql + " 										) AS VW1 ";
                    tmp_sql = tmp_sql + " 										GROUP BY bk_no,hy_composite,reform_id ";
                    tmp_sql = tmp_sql + " 										/*合計の総和と送金額の総和を抽出 end*/ ";
                    tmp_sql = tmp_sql + " 									) AS REFITEMSUM ";
                    tmp_sql = tmp_sql + " 									LEFT JOIN ";
                    tmp_sql = tmp_sql + " 										( ";
                    tmp_sql = tmp_sql + " 											SELECT * FROM reform_item AS REFITEM ";
                    tmp_sql = tmp_sql + " 											WHERE nkin_no = 8999 ";
                    tmp_sql = tmp_sql + " 										) AS REFITEMTAX ";
                    tmp_sql = tmp_sql + " 									ON  REFITEMSUM.bk_no = REFITEMTAX.bk_no ";
                    tmp_sql = tmp_sql + " 									AND REFITEMSUM.hy_composite = REFITEMTAX.hy_composite ";
                    tmp_sql = tmp_sql + " 									AND REFITEMSUM.reform_id = REFITEMTAX.reform_id ";
                    tmp_sql = tmp_sql + " 									) AS VW ";
                    tmp_sql = tmp_sql + " 								) AS REFSOGAK ";
                    tmp_sql = tmp_sql + " 							ON  REFDATA.bk_no = REFSOGAK.bk_no ";
                    tmp_sql = tmp_sql + " 							AND REFDATA.hy_composite = REFSOGAK.hy_composite ";
                    tmp_sql = tmp_sql + " 							AND REFDATA.reform_id = REFSOGAK.reform_id ";
                    tmp_sql = tmp_sql + " 							LEFT JOIN ";
                    tmp_sql = tmp_sql + " 								( ";
                    tmp_sql = tmp_sql + " 									SELECT ";
                    tmp_sql = tmp_sql + " 										 * ";
                    tmp_sql = tmp_sql + " 										,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 									FROM ky_kosinkai AS KYK ";
                    tmp_sql = tmp_sql + " 								) AS KYBASE ";
                    tmp_sql = tmp_sql + " 							ON  REFDATA.bk_no = KYBASE.bk_no ";
                    tmp_sql = tmp_sql + " 							AND REFDATA.hy_composite = KYBASE.hy_no ";
                    tmp_sql = tmp_sql + " 							AND REFDATA.ky_no = KYBASE.ky_no ";
                    tmp_sql = tmp_sql + " 							AND REFDATA.ko_no = KYBASE.ko_no ";
                    tmp_sql = tmp_sql + " 							WHERE REFDATA.reform_kbn = 0 ";
                    tmp_sql = tmp_sql + " 							AND   REFDATA.ko_no = 999 ";
                    tmp_sql = tmp_sql + " 						) AS REFSO ";
                    tmp_sql = tmp_sql + " 						/*リフォームデータから送金率取得 end*/ ";
                    tmp_sql = tmp_sql + " 					) AS VW2	 ";
                    tmp_sql = tmp_sql + " 				) AS SOGAKTOTAL ";
                    tmp_sql = tmp_sql + " 			ON  SQSAKITOTAL.bk_no = SOGAKTOTAL.bk_no ";
                    tmp_sql = tmp_sql + " 			AND SQSAKITOTAL.hy_no = SOGAKTOTAL.hy_no ";
                    tmp_sql = tmp_sql + " 			AND SQSAKITOTAL.ky_no = SOGAKTOTAL.ky_no ";
                    tmp_sql = tmp_sql + " 			AND SQSAKITOTAL.契約管理レコードNo = SOGAKTOTAL.契約管理レコードNo ";
                    tmp_sql = tmp_sql + " 		) AS KAISZENDATA ";
                    tmp_sql = tmp_sql + " 	ON  KYTOTAL.bk_no = KAISZENDATA.bk_no ";
                    tmp_sql = tmp_sql + " 	AND KYTOTAL.hy_no = KAISZENDATA.hy_no ";
                    tmp_sql = tmp_sql + " 	AND KYTOTAL.ky_no = KAISZENDATA.ky_no ";
                    tmp_sql = tmp_sql + " 	AND KYTOTAL.[革命10用解約レコードNo] = KAISZENDATA.契約管理レコードNo ";
                }
                else
                {
                    tmp_sql = tmp_sql + " 	LEFT JOIN ";
                    tmp_sql = tmp_sql + " 		( ";
                    tmp_sql = tmp_sql + " 			SELECT ";
                    tmp_sql = tmp_sql + " 				 SQSAKITOTAL.* ";
                    tmp_sql = tmp_sql + " 				,SOGAKTOTAL.so_gak ";
                    tmp_sql = tmp_sql + " 				,SOGAKTOTAL.so_rit ";
                    tmp_sql = tmp_sql + " 				,nkbn_name ";
                    tmp_sql = tmp_sql + " 			FROM ";
                    tmp_sql = tmp_sql + " 			( ";
                    tmp_sql = tmp_sql + " 				/*契約請求データから請求先取得 sta*/ ";
                    tmp_sql = tmp_sql + " 				SELECT * FROM ";
                    tmp_sql = tmp_sql + " 				( ";
                    tmp_sql = tmp_sql + " 					SELECT ";
                    tmp_sql = tmp_sql + " 						 bk_no ";
                    tmp_sql = tmp_sql + " 						,hy_no ";
                    tmp_sql = tmp_sql + " 						,ky_no ";
                    tmp_sql = tmp_sql + " 						,[契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 						,kys_no ";
                    tmp_sql = tmp_sql + " 						,nkbn ";
                    tmp_sql = tmp_sql + " 					FROM ";
                    tmp_sql = tmp_sql + " 					( ";
                    tmp_sql = tmp_sql + " 						SELECT ";
                    tmp_sql = tmp_sql + " 							 KYSQ.* ";
                    tmp_sql = tmp_sql + " 							,[契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 							,ROW_NUMBER()OVER(PARTITION BY KYSQ.bk_no,KYSQ.hy_no,KYSQ.ky_no ORDER BY KYSQ.naibu_no) AS [請求先抽出用] ";
                    tmp_sql = tmp_sql + " 						FROM ky_sqdata AS KYSQ ";
                    tmp_sql = tmp_sql + " 						LEFT JOIN ";
                    tmp_sql = tmp_sql + " 							( ";
                    tmp_sql = tmp_sql + " 								SELECT bk_no,hy_no,ky_no,ko_no ";
                    tmp_sql = tmp_sql + " 									  ,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 								FROM ky_kosinkai AS KYK ";
                    tmp_sql = tmp_sql + " 							) AS KYKBASE ";
                    tmp_sql = tmp_sql + " 						ON  KYSQ.bk_no = KYKBASE.bk_no ";
                    tmp_sql = tmp_sql + " 						AND KYSQ.hy_no = KYKBASE.hy_no ";
                    tmp_sql = tmp_sql + " 						AND KYSQ.ky_no = KYKBASE.ky_no ";
                    tmp_sql = tmp_sql + " 						AND KYSQ.ko_no = KYKBASE.ko_no ";
                    tmp_sql = tmp_sql + " 						WHERE KYSQ.ko_no = 999 ";
                    tmp_sql = tmp_sql + " 						AND   KYSQ.nkin_kbn = 8 ";
                    tmp_sql = tmp_sql + " 					) AS VW ";
                    tmp_sql = tmp_sql + " 					WHERE [請求先抽出用] = 1 ";
                    tmp_sql = tmp_sql + " 				) AS KYSQSAKI ";
                    tmp_sql = tmp_sql + " 				/*契約請求データから請求先取得 end*/ ";
                    tmp_sql = tmp_sql + " 			) AS SQSAKITOTAL ";
                    tmp_sql = tmp_sql + " 			LEFT JOIN m_nkbn AS MNKBN ON SQSAKITOTAL.nkbn = MNKBN.nkbn_no ";
                    tmp_sql = tmp_sql + " 			LEFT JOIN ";
                    tmp_sql = tmp_sql + " 				( ";
                    tmp_sql = tmp_sql + " 					SELECT * FROM ";
                    tmp_sql = tmp_sql + " 					( ";
                    tmp_sql = tmp_sql + " 						/*契約請求データから送金額・送金率取得 sta*/ ";
                    tmp_sql = tmp_sql + " 						SELECT ";
                    tmp_sql = tmp_sql + " 							 bk_no ";
                    tmp_sql = tmp_sql + " 							,hy_no ";
                    tmp_sql = tmp_sql + " 							,ky_no ";
                    tmp_sql = tmp_sql + " 							,[契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 							,SUM(ISNULL(so_gak,0)) AS so_gak ";
                    tmp_sql = tmp_sql + " 							,CASE ";
                    tmp_sql = tmp_sql + " 								WHEN SUM(ISNULL(so_gak,0)) = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 								/*20160629 修繕検証後修正 chg sta*/ ";
                    tmp_sql = tmp_sql + " 								/*ELSE (SUM(ISNULL(so_gak,0)) / SUM(ISNULL(kai_syuz_kin,0))) * 100*/ ";
                    tmp_sql = tmp_sql + " 								ELSE ";
                    tmp_sql = tmp_sql + " 									CASE ";
                    tmp_sql = tmp_sql + " 										WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((SUM(ISNULL(so_gak,0)) / SUM(ISNULL(kai_syuz_kin,0))) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 										WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((SUM(ISNULL(so_gak,0)) / SUM(ISNULL(kai_syuz_kin,0))) * 100,2)				/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 										WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((SUM(ISNULL(so_gak,0)) / SUM(ISNULL(kai_syuz_kin,0))) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 									END ";
                    tmp_sql = tmp_sql + " 								/*20160629 修繕検証後修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 							 END AS so_rit ";
                    tmp_sql = tmp_sql + " 						FROM ";
                    tmp_sql = tmp_sql + " 						( ";
                    tmp_sql = tmp_sql + " 							SELECT ";
                    tmp_sql = tmp_sql + " 								 KYSQ.* ";
                    tmp_sql = tmp_sql + " 								,[契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 							FROM ky_sqdata AS KYSQ ";
                    tmp_sql = tmp_sql + " 							LEFT JOIN ";
                    tmp_sql = tmp_sql + " 								( ";
                    tmp_sql = tmp_sql + " 									SELECT bk_no,hy_no,ky_no,ko_no ";
                    tmp_sql = tmp_sql + " 										  ,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY ko_no) AS [契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 									FROM ky_kosinkai AS KYK ";
                    tmp_sql = tmp_sql + " 								) AS KYKBASE ";
                    tmp_sql = tmp_sql + " 							ON  KYSQ.bk_no = KYKBASE.bk_no ";
                    tmp_sql = tmp_sql + " 							AND KYSQ.hy_no = KYKBASE.hy_no ";
                    tmp_sql = tmp_sql + " 							AND KYSQ.ky_no = KYKBASE.ky_no ";
                    tmp_sql = tmp_sql + " 							AND KYSQ.ko_no = KYKBASE.ko_no ";
                    tmp_sql = tmp_sql + " 							WHERE KYSQ.ko_no = 999 ";
                    tmp_sql = tmp_sql + " 							AND   KYSQ.nkin_kbn = 8 ";
                    tmp_sql = tmp_sql + " 						) AS VW1 ";
                    tmp_sql = tmp_sql + " 						GROUP BY bk_no,hy_no,ky_no,[契約管理レコードNo] ";
                    tmp_sql = tmp_sql + " 						/*契約請求データから送金額・送金率取得 end*/ ";
                    tmp_sql = tmp_sql + " 					) AS VW2	 ";
                    tmp_sql = tmp_sql + " 				) AS SOGAKTOTAL ";
                    tmp_sql = tmp_sql + " 			ON  SQSAKITOTAL.bk_no = SOGAKTOTAL.bk_no ";
                    tmp_sql = tmp_sql + " 			AND SQSAKITOTAL.hy_no = SOGAKTOTAL.hy_no ";
                    tmp_sql = tmp_sql + " 			AND SQSAKITOTAL.ky_no = SOGAKTOTAL.ky_no ";
                    tmp_sql = tmp_sql + " 			AND SQSAKITOTAL.契約管理レコードNo = SOGAKTOTAL.契約管理レコードNo ";
                    tmp_sql = tmp_sql + " 		) AS KAISZENDATA ";
                    tmp_sql = tmp_sql + " 	ON  KYTOTAL.bk_no = KAISZENDATA.bk_no ";
                    tmp_sql = tmp_sql + " 	AND KYTOTAL.hy_no = KAISZENDATA.hy_no ";
                    tmp_sql = tmp_sql + " 	AND KYTOTAL.ky_no = KAISZENDATA.ky_no ";
                    tmp_sql = tmp_sql + " 	AND KYTOTAL.[革命10用解約レコードNo] = KAISZENDATA.契約管理レコードNo ";
                }

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 契約修繕見積情報

    public class Reform_data_kaiszen_Repository
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

                // 20160629 修繕検証後修正 -add
                bool szenopflg = false;

                string tmp_sql = "";

                sortstr = "[物件No],[部屋No],[契約No],[更新No],[見積No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[更新No],[見積No],[見積タイトル],[見積番号],[見積日],[成約フラグ],[成約日],[負担区分(全体・個別)],[契約者負担率(全体用)],[家主負担率(全体用)],[自社負担率(全体用)],[税適用区分(全体・個別)],[税区分(全体用)],[適用税率],[契約者税額(全体用)],[家主税額(全体用)],[自社税額(全体用)],[端数負担者区分]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                // 20160629 修繕検証後修正 -add
                tmp_sql = tmp_sql + Get_UseQry_SzenOp(szenopflg);
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

            /// <summary>
            /// 修繕OPの有無で取得元を分岐 '20160629 修繕検証後修正
            /// 修繕OP：有→リフォームデータから
            /// 修繕OP：無→契約請求データから
            /// </summary>
            /// <param name="opflg"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_SzenOp(bool opflg)
            {

                string tmp_sql = "";

                if (opflg)
                {
                    tmp_sql = tmp_sql + " 	SELECT ";
                    tmp_sql = tmp_sql + " 		 REFDATA.bk_no AS [物件No] ";
                    tmp_sql = tmp_sql + " 		,REFDATA.hy_composite AS [部屋No] ";
                    tmp_sql = tmp_sql + " 		,REFDATA.ky_no AS [契約No] ";
                    tmp_sql = tmp_sql + " 		,[革命10用解約レコードNo] AS [更新No] ";
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
                    tmp_sql = tmp_sql + " 		 /*20160627 修繕追加分移行修正 整数型に変換する関数を入れる sta*/ ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,zei_gak_kari) AS [契約者税額(全体用)] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,zei_gak_owner) AS [家主税額(全体用)] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,zei_gak_jisha) AS [自社税額(全体用)] ";
                    tmp_sql = tmp_sql + " 		/*20160627 修繕追加分移行修正 整数型に変換する関数を入れる end*/ ";
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
                    tmp_sql = tmp_sql + " 				SELECT * FROM ";
                    tmp_sql = tmp_sql + " 				( ";
                    tmp_sql = tmp_sql + " 					SELECT ";
                    tmp_sql = tmp_sql + " 						 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY bk_no,hy_no,ky_no,ko_no) AS [革命10用解約レコードNo] ";
                    tmp_sql = tmp_sql + " 						,* ";
                    tmp_sql = tmp_sql + " 					FROM ky_kosinkai ";
                    tmp_sql = tmp_sql + " 				) AS VW ";
                    tmp_sql = tmp_sql + " 				WHERE ko_no = 999 ";
                    tmp_sql = tmp_sql + " 		) AS KYK ";
                    tmp_sql = tmp_sql + " 	ON  REFDATA.bk_no = KYK.bk_no ";
                    tmp_sql = tmp_sql + " 	AND REFDATA.hy_composite = KYK.hy_no ";
                    tmp_sql = tmp_sql + " 	AND REFDATA.ky_no = KYK.ky_no ";
                    tmp_sql = tmp_sql + " 	AND REFDATA.ko_no = KYK.ko_no ";
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
                    tmp_sql = tmp_sql + " 	WHERE REFDATA.reform_kbn = 0	/*原状回復*/ ";
                }
                else
                {
                    tmp_sql = tmp_sql + " 	SELECT ";
                    tmp_sql = tmp_sql + " 		 KYBASE.bk_no AS [物件No] ";
                    tmp_sql = tmp_sql + " 		,KYBASE.hy_no AS [部屋No] ";
                    tmp_sql = tmp_sql + " 		,KYBASE.ky_no AS [契約No] ";
                    tmp_sql = tmp_sql + " 		,[革命10用解約レコードNo] AS [更新No] ";
                    tmp_sql = tmp_sql + " 		,1 AS [見積No] ";
                    tmp_sql = tmp_sql + " 		,'' AS [見積タイトル] ";
                    tmp_sql = tmp_sql + " 		,'' AS [見積番号] ";
                    tmp_sql = tmp_sql + " 		,'' AS [見積日] ";
                    tmp_sql = tmp_sql + " 		/*20160629 修繕検証後修正 chg sta 成約フラグのチェックをONにする*/ ";
                    tmp_sql = tmp_sql + " 		,1 AS [成約フラグ] ";
                    tmp_sql = tmp_sql + " 		,'' AS [成約日] ";
                    tmp_sql = tmp_sql + " 		,1 AS [負担区分(全体・個別)] ";
                    tmp_sql = tmp_sql + " 		,NULL AS [契約者負担率(全体用)] ";
                    tmp_sql = tmp_sql + " 		,NULL AS [家主負担率(全体用)] ";
                    tmp_sql = tmp_sql + " 		,NULL AS [自社負担率(全体用)] ";
                    tmp_sql = tmp_sql + " 		,1 AS [税適用区分(全体・個別)] ";
                    tmp_sql = tmp_sql + " 		,2 AS [税区分(全体用)] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN 3 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN 5 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN 8 ";
                    tmp_sql = tmp_sql + " 			ELSE 8 ";
                    tmp_sql = tmp_sql + " 		 END AS [適用税率] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,kys_zeigak_total) AS [契約者税額(全体用)] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,ow_zeigak_total) AS [家主税額(全体用)] ";
                    tmp_sql = tmp_sql + " 		,0 AS [自社税額(全体用)] ";
                    tmp_sql = tmp_sql + " 		,100 AS [端数負担者区分] ";
                    tmp_sql = tmp_sql + " 	FROM ";
                    tmp_sql = tmp_sql + " 	( ";
                    tmp_sql = tmp_sql + " 		SELECT ";
                    tmp_sql = tmp_sql + " 			 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY bk_no,hy_no,ky_no,ko_no) AS [革命10用解約レコードNo] ";
                    tmp_sql = tmp_sql + " 			,* ";
                    tmp_sql = tmp_sql + " 		FROM ky_kosinkai ";
                    tmp_sql = tmp_sql + " 	) AS KYBASE ";
                    tmp_sql = tmp_sql + " 	LEFT JOIN ";
                    tmp_sql = tmp_sql + " 		( ";
                    tmp_sql = tmp_sql + " 			/*契約入金項目情報から契約者、家主の税額合計を抽出 sta*/ ";
                    tmp_sql = tmp_sql + " 			SELECT * FROM ";
                    tmp_sql = tmp_sql + " 			( ";
                    tmp_sql = tmp_sql + " 				SELECT ";
                    tmp_sql = tmp_sql + " 					 bk_no ";
                    tmp_sql = tmp_sql + " 					,hy_no ";
                    tmp_sql = tmp_sql + " 					,ky_no ";
                    tmp_sql = tmp_sql + " 					,ko_no ";
                    tmp_sql = tmp_sql + " 					,SUM(kys_zeigak) AS kys_zeigak_total ";
                    tmp_sql = tmp_sql + " 					,SUM(ow_zeigak) AS ow_zeigak_total ";
                    tmp_sql = tmp_sql + " 				FROM ";
                    tmp_sql = tmp_sql + " 				( ";
                    tmp_sql = tmp_sql + " 					SELECT ";
                    tmp_sql = tmp_sql + " 						 KYSQ.bk_no ";
                    tmp_sql = tmp_sql + " 						,KYSQ.hy_no ";
                    tmp_sql = tmp_sql + " 						,KYSQ.ky_no ";
                    tmp_sql = tmp_sql + " 						,KYSQ.ko_no ";
                    tmp_sql = tmp_sql + " 						,CASE ";
                    tmp_sql = tmp_sql + " 							WHEN ISNULL(kai_syuz_kin,0) = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 							WHEN kai_syuz_zeiumu = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 							WHEN kai_syuz_zeiumu = 1 THEN ";
                    tmp_sql = tmp_sql + " 								CASE ";
                    tmp_sql = tmp_sql + " 									WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ROUND(kai_syuz_kin * 0.03,0) ";
                    tmp_sql = tmp_sql + " 									WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ROUND(kai_syuz_kin * 0.05,0) ";
                    tmp_sql = tmp_sql + " 									WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ROUND(kai_syuz_kin * 0.08,0) ";
                    tmp_sql = tmp_sql + " 								END ";
                    tmp_sql = tmp_sql + " 						 END AS kys_zeigak ";
                    tmp_sql = tmp_sql + " 						,CASE ";
                    tmp_sql = tmp_sql + " 							WHEN ISNULL(kai_syuz_kin,0) = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 							WHEN kai_syuz_zeiumu = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 							WHEN kai_syuz_zeiumu = 1 THEN ";
                    tmp_sql = tmp_sql + " 								CASE ";
                    tmp_sql = tmp_sql + " 									WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ROUND(futanrit_so * 0.03/1.03,0) ";
                    tmp_sql = tmp_sql + " 									WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ROUND(futanrit_so * 0.05/1.05,0) ";
                    tmp_sql = tmp_sql + " 									WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ROUND(futanrit_so * 0.08/1.08,0) ";
                    tmp_sql = tmp_sql + " 								END ";
                    tmp_sql = tmp_sql + " 						 END AS ow_zeigak	 ";
                    tmp_sql = tmp_sql + " 					FROM ky_sqdata AS KYSQ ";
                    tmp_sql = tmp_sql + " 					LEFT JOIN ky_kosinkai AS KYK ";
                    tmp_sql = tmp_sql + " 					ON  KYSQ.bk_no = KYK.bk_no ";
                    tmp_sql = tmp_sql + " 					AND KYSQ.hy_no = KYK.hy_no ";
                    tmp_sql = tmp_sql + " 					AND KYSQ.ky_no = KYK.ky_no ";
                    tmp_sql = tmp_sql + " 					AND KYSQ.ko_no = KYK.ko_no ";
                    tmp_sql = tmp_sql + " 					WHERE KYSQ.ko_no = 999 ";
                    tmp_sql = tmp_sql + " 					AND   nkin_kbn = 8 ";
                    tmp_sql = tmp_sql + " 				) AS VW1 ";
                    tmp_sql = tmp_sql + " 				GROUP BY bk_no,hy_no,ky_no,ko_no ";
                    tmp_sql = tmp_sql + " 			) AS VW2 ";
                    tmp_sql = tmp_sql + " 			/*契約入金項目情報から契約者、家主の税額合計を抽出 end*/ ";
                    tmp_sql = tmp_sql + " 		) AS ZEIGAKTOTAL ";
                    tmp_sql = tmp_sql + " 	ON  KYBASE.bk_no = ZEIGAKTOTAL.bk_no ";
                    tmp_sql = tmp_sql + " 	AND KYBASE.hy_no = ZEIGAKTOTAL.hy_no ";
                    tmp_sql = tmp_sql + " 	AND KYBASE.ky_no = ZEIGAKTOTAL.ky_no ";
                    tmp_sql = tmp_sql + " 	AND KYBASE.ko_no = ZEIGAKTOTAL.ko_no ";
                    tmp_sql = tmp_sql + " 	WHERE ZEIGAKTOTAL.bk_no IS NOT NULL ";
                }

                return tmp_sql;

            }

        }

    }

    #endregion

    #region 契約修繕見積詳細情報

    public class Reform_item_kaiszen_Repository
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

                // 20160629 修繕検証後修正 -add
                bool szenopflg = false;

                string tmp_sql = "";

                sortstr = "[物件No],[部屋No],[契約No],[更新No],[見積No],[見積明細No]";

                // 抽出データの改行文字列を除去する暫定処理
                string tmp_midheader = "[物件No],[部屋No],[契約No],[更新No],[見積No],[見積明細No],[修繕項目名],[摘要],[見積数量],[見積単位],[見積単価],[見積税区分],[見積税額(税入力用)],[契約者負担率(個別用)],[家主負担率(個別用)],[自社負担率(個別用)],[発注業者No],[実行数量],[実行単位],[実行単価],[実行税区分],[実行税額(税入力用)]";
                string tmp_chgmidheader = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, true);
                // 抽出データの改行文字列を除去する暫定処理

                // tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql + " SELECT " + tmp_chgmidheader + " FROM ";
                tmp_sql = tmp_sql + " ( ";
                // 20160629 修繕検証後修正 -add
                tmp_sql = tmp_sql + Get_UseQry_SzenOp(szenopflg);
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

            /// <summary>
            /// 修繕OPの有無で取得元を分岐 '20160629 修繕検証後修正
            /// 修繕OP：有→リフォームデータから
            /// 修繕OP：無→契約請求データから
            /// </summary>
            /// <param name="opflg"></param>
            /// <returns></returns>
            /// <remarks></remarks>
            public string Get_UseQry_SzenOp(bool opflg)
            {

                string tmp_sql = "";

                if (opflg)
                {
                    tmp_sql = tmp_sql + " 	SELECT ";
                    tmp_sql = tmp_sql + " 		 REFITEM.bk_no AS [物件No] ";
                    tmp_sql = tmp_sql + " 		,REFDATA.hy_composite AS [部屋No] ";
                    tmp_sql = tmp_sql + " 		,REFDATA.ky_no AS [契約No] ";
                    tmp_sql = tmp_sql + " 		,[革命10用解約レコードNo] AS [更新No] ";
                    tmp_sql = tmp_sql + " 		/*,[修繕No]*/ ";
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
                    tmp_sql = tmp_sql + "     		/*20160819 除算エラー修正 add*/ ";
                    tmp_sql = tmp_sql + "     		WHEN ISNULL(total_plan,0) = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg sta*/ ";
                    tmp_sql = tmp_sql + " 			/*ELSE ROUND((total_kari / total_plan) * 100,0)*/ ";
                    tmp_sql = tmp_sql + " 			ELSE ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((total_kari / total_plan) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((total_kari / total_plan) * 100,2)				/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((total_kari / total_plan) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 			/*20160629 修繕検証後修正 chg end*/    			 ";
                    tmp_sql = tmp_sql + " 		 END AS [契約者負担率(個別用)] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN ISNULL(futan_rit_owner,0) <> 0 THEN futan_rit_owner ";
                    tmp_sql = tmp_sql + "     		/*20160819 除算エラー修正 add*/ ";
                    tmp_sql = tmp_sql + "     		WHEN ISNULL(total_plan,0) = 0 THEN 0 ";
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
                    tmp_sql = tmp_sql + "     		/*20160819 除算エラー修正 add*/ ";
                    tmp_sql = tmp_sql + "     		WHEN ISNULL(total_plan,0) = 0 THEN 0 ";
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
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,repare_no) AS [発注業者No] ";
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
                    tmp_sql = tmp_sql + " 	LEFT JOIN ";
                    tmp_sql = tmp_sql + " 		( ";
                    tmp_sql = tmp_sql + " 				SELECT * FROM ";
                    tmp_sql = tmp_sql + " 				( ";
                    tmp_sql = tmp_sql + " 					SELECT ";
                    tmp_sql = tmp_sql + " 						 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY bk_no,hy_no,ky_no,ko_no) AS [革命10用解約レコードNo] ";
                    tmp_sql = tmp_sql + " 						,* ";
                    tmp_sql = tmp_sql + " 					FROM ky_kosinkai ";
                    tmp_sql = tmp_sql + " 				) AS VW ";
                    tmp_sql = tmp_sql + " 				WHERE ko_no = 999 ";
                    tmp_sql = tmp_sql + " 		) AS KYK ";
                    tmp_sql = tmp_sql + " 	ON  REFDATA.bk_no = KYK.bk_no ";
                    tmp_sql = tmp_sql + " 	AND REFDATA.hy_composite = KYK.hy_no ";
                    tmp_sql = tmp_sql + " 	AND REFDATA.ky_no = KYK.ky_no ";
                    tmp_sql = tmp_sql + " 	AND REFDATA.ko_no = KYK.ko_no ";
                    tmp_sql = tmp_sql + " 	WHERE reform_kbn = 0			/*原状回復*/ ";
                    tmp_sql = tmp_sql + " 	AND   REFITEM.nkin_no <> 8999	/*合計額に消費税が設定されている場合の消費税レコードは抽出対象外*/ ";
                }
                else
                {
                    tmp_sql = tmp_sql + " 	SELECT ";
                    tmp_sql = tmp_sql + " 		 KYSQREF.bk_no AS [物件No] ";
                    tmp_sql = tmp_sql + " 		,KYSQREF.hy_no AS [部屋No] ";
                    tmp_sql = tmp_sql + " 		,KYSQREF.ky_no AS [契約No] ";
                    tmp_sql = tmp_sql + " 		,[革命10用解約レコードNo] AS [更新No] ";
                    tmp_sql = tmp_sql + " 		,1 AS [見積No] ";
                    tmp_sql = tmp_sql + " 		,ROW_NUMBER()OVER(PARTITION BY KYSQREF.bk_no,KYSQREF.hy_no,KYSQREF.ky_no,[革命10用解約レコードNo] ORDER BY KYSQREF.nkin_no) AS [見積明細No] ";
                    tmp_sql = tmp_sql + " 		,MN.nkin_name AS [修繕項目名] ";
                    tmp_sql = tmp_sql + " 		,KYSQREF.biko AS [摘要] ";
                    tmp_sql = tmp_sql + " 		/*20160629 修繕検証後修正 chg sta*/ ";
                    tmp_sql = tmp_sql + " 		/* ";
                    tmp_sql = tmp_sql + " 		,'' AS [見積数量] ";
                    tmp_sql = tmp_sql + " 		,'' AS [見積単位] ";
                    tmp_sql = tmp_sql + " 		,'' AS [見積単価] ";
                    tmp_sql = tmp_sql + " 		,'' AS [見積税区分] ";
                    tmp_sql = tmp_sql + " 		,'' AS [見積税額(税入力用)] ";
                    tmp_sql = tmp_sql + " 		*/ ";
                    tmp_sql = tmp_sql + " 		,kai_syuz_su AS [見積数量] ";
                    tmp_sql = tmp_sql + " 		,kai_syuz_tanni AS [見積単位] ";
                    tmp_sql = tmp_sql + " 		,kai_syuz_tanka AS [見積単価] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 0 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 1 THEN 2 ";
                    tmp_sql = tmp_sql + " 		 END AS [見積税区分] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_kin = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 1 THEN ";
                    tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 chg sta*/ ";
                    tmp_sql = tmp_sql + " 				/* ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ROUND(kai_syuz_kin * 0.03,0) ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ROUND(kai_syuz_kin * 0.05,0) ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ROUND(kai_syuz_kin * 0.08,0) ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 				*/ ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 1 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.03,0,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 2 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.03,0)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 3 THEN CEILING(kai_syuz_su * kai_syuz_tanka * 0.03)					/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 1 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.05,0,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 2 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.05,0)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 3 THEN CEILING(kai_syuz_su * kai_syuz_tanka * 0.05)					/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END    					 ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 1 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.08,0,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 2 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.08,0)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 3 THEN CEILING(kai_syuz_su * kai_syuz_tanka * 0.08)					/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 		 END AS [見積税額(税入力用)] ";
                    tmp_sql = tmp_sql + " 		/*20160629 修繕検証後修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 		/*20160629 修繕検証後修正 chg sta*/ ";
                    tmp_sql = tmp_sql + " 		/* ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_kin = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 0 THEN ROUND(((kai_syuz_kin - futanrit_so)/kai_syuz_kin) * 100,0) ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 1 THEN ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ROUND(((kai_syuz_kin * 1.03 - futanrit_so)/(kai_syuz_kin * 1.03)) * 100,0) ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ROUND(((kai_syuz_kin * 1.05 - futanrit_so)/(kai_syuz_kin * 1.05)) * 100,0) ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ROUND(((kai_syuz_kin * 1.08 - futanrit_so)/(kai_syuz_kin * 1.08)) * 100,0) ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 		 END AS [契約者負担率(個別用)]	 ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_kin = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 0 THEN 100 - ROUND(((kai_syuz_kin - futanrit_so)/kai_syuz_kin) * 100,0) ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 1 THEN ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN 100 - ROUND(((kai_syuz_kin * 1.03 - futanrit_so)/(kai_syuz_kin * 1.03)) * 100,0) ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN 100 - ROUND(((kai_syuz_kin * 1.05 - futanrit_so)/(kai_syuz_kin * 1.05)) * 100,0) ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN 100 - ROUND(((kai_syuz_kin * 1.08 - futanrit_so)/(kai_syuz_kin * 1.08)) * 100,0) ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 		 END AS [家主負担率(個別用)] ";
                    tmp_sql = tmp_sql + " 		 */ ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_kin = 0 THEN 0 ";
                    tmp_sql = tmp_sql + "     		/*20160819 除算エラー修正 add*/ ";
                    tmp_sql = tmp_sql + "     		WHEN kai_syuz_su * kai_syuz_tanka = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 0 THEN ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					/*WHEN kai_syuz_kin - futanrit_so < 0 THEN 0*/ ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((kai_syuz_kin/(kai_syuz_su * kai_syuz_tanka)) * 100,2,1)					/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((kai_syuz_kin/(kai_syuz_su * kai_syuz_tanka)) * 100,2)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((kai_syuz_kin/(kai_syuz_su * kai_syuz_tanka)) * 100) * 100) / 100		/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 1 THEN ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							/*WHEN kai_syuz_kin * 1.03 - futanrit_so < 0 THEN 0*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND(((kai_syuz_kin * 1.03)/(kai_syuz_su * kai_syuz_tanka * 1.03)) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND(((kai_syuz_kin * 1.03)/(kai_syuz_su * kai_syuz_tanka * 1.03)) * 100,2)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING((((kai_syuz_kin * 1.03)/(kai_syuz_su * kai_syuz_tanka * 1.03)) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							/*WHEN kai_syuz_kin * 1.05 - futanrit_so < 0 THEN 0*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND(((kai_syuz_kin * 1.05)/(kai_syuz_su * kai_syuz_tanka * 1.05)) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND(((kai_syuz_kin * 1.05)/(kai_syuz_su * kai_syuz_tanka * 1.05)) * 100,2)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING((((kai_syuz_kin * 1.05)/(kai_syuz_su * kai_syuz_tanka * 1.05)) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							/*WHEN kai_syuz_kin * 1.08 - futanrit_so < 0 THEN 0*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND(((kai_syuz_kin * 1.08)/(kai_syuz_su * kai_syuz_tanka * 1.08)) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND(((kai_syuz_kin * 1.08)/(kai_syuz_su * kai_syuz_tanka * 1.08)) * 100,2)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING((((kai_syuz_kin * 1.08)/(kai_syuz_su * kai_syuz_tanka * 1.08)) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 		 END AS [契約者負担率(個別用)]	 ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_kin = 0 THEN 0 ";
                    tmp_sql = tmp_sql + "     		/*20160819 除算エラー修正 add*/ ";
                    tmp_sql = tmp_sql + "     		WHEN kai_syuz_su * kai_syuz_tanka = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 0 THEN ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((futanrit_so/(kai_syuz_su * kai_syuz_tanka)) * 100,2,1)					/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((futanrit_so/(kai_syuz_su * kai_syuz_tanka)) * 100,2)						/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 					WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((futanrit_so/(kai_syuz_su * kai_syuz_tanka)) * 100) * 100) / 100		/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 				END    			 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 1 THEN ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.03)) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.03)) * 100,2)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.03)) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.05)) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.05)) * 100,2)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.05)) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 1 THEN ROUND((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.08)) * 100,2,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 2 THEN ROUND((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.08)) * 100,2)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_kngk FROM m_kan) = 3 THEN CEILING(((futanrit_so/(kai_syuz_su * kai_syuz_tanka * 1.08)) * 100) * 100) / 100	/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 		 END AS [家主負担率(個別用)] ";
                    tmp_sql = tmp_sql + " 		 /*20160629 修繕検証後修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 		,0 AS [自社負担率(個別用)] ";
                    tmp_sql = tmp_sql + " 		,CONVERT(INT,KYSQREF.yobi_fl1) AS [発注業者No] ";
                    tmp_sql = tmp_sql + " 		,kai_syuz_su AS [実行数量] ";
                    tmp_sql = tmp_sql + " 		,kai_syuz_tanni AS [実行単位] ";
                    tmp_sql = tmp_sql + " 		,kai_syuz_tanka AS [実行単価] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 0 THEN 1 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 1 THEN 2 ";
                    tmp_sql = tmp_sql + " 		 END AS [実行税区分] ";
                    tmp_sql = tmp_sql + " 		,CASE ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_kin = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 0 THEN 0 ";
                    tmp_sql = tmp_sql + " 			WHEN kai_syuz_zeiumu = 1 THEN ";
                    tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 chg sta*/ ";
                    tmp_sql = tmp_sql + " 				/* ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ROUND(kai_syuz_kin * 0.03,0) ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ROUND(kai_syuz_kin * 0.05,0) ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ROUND(kai_syuz_kin * 0.08,0) ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 				*/ ";
                    tmp_sql = tmp_sql + " 				CASE ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19890401' AND '19970331' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 1 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.03,0,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 2 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.03,0)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 3 THEN CEILING(kai_syuz_su * kai_syuz_tanka * 0.03)					/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '19970401' AND '20140331' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 1 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.05,0,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 2 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.05,0)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 3 THEN CEILING(kai_syuz_su * kai_syuz_tanka * 0.05)					/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END    					 ";
                    tmp_sql = tmp_sql + " 					WHEN kai_ymd BETWEEN '20140401' AND '99991231' THEN ";
                    tmp_sql = tmp_sql + " 						CASE ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 1 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.08,0,1)				/*切り捨て*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 2 THEN ROUND(kai_syuz_su * kai_syuz_tanka * 0.08,0)					/*四捨五入*/ ";
                    tmp_sql = tmp_sql + " 							WHEN (SELECT hasu_zei FROM m_kan) = 3 THEN CEILING(kai_syuz_su * kai_syuz_tanka * 0.08)					/*切り上げ*/ ";
                    tmp_sql = tmp_sql + " 						END ";
                    tmp_sql = tmp_sql + " 				END ";
                    tmp_sql = tmp_sql + " 				/*20160629 修繕検証後修正 chg end*/ ";
                    tmp_sql = tmp_sql + " 		 END AS [実行税額(税入力用)] ";
                    tmp_sql = tmp_sql + " 	FROM ";
                    tmp_sql = tmp_sql + " 	( ";
                    tmp_sql = tmp_sql + " 		SELECT * FROM ky_sqdata AS KYSQ ";
                    tmp_sql = tmp_sql + " 		WHERE ko_no = 999 ";
                    tmp_sql = tmp_sql + " 		AND   nkin_kbn = 8 ";
                    tmp_sql = tmp_sql + " 	) AS KYSQREF ";
                    tmp_sql = tmp_sql + " 	LEFT JOIN ";
                    tmp_sql = tmp_sql + " 		( ";
                    tmp_sql = tmp_sql + " 			SELECT * FROM ";
                    tmp_sql = tmp_sql + " 			( ";
                    tmp_sql = tmp_sql + " 				SELECT ";
                    tmp_sql = tmp_sql + " 					 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,ky_no ORDER BY bk_no,hy_no,ky_no,ko_no) AS [革命10用解約レコードNo] ";
                    tmp_sql = tmp_sql + " 					,* ";
                    tmp_sql = tmp_sql + " 				FROM ky_kosinkai ";
                    tmp_sql = tmp_sql + " 			) AS VW ";
                    tmp_sql = tmp_sql + " 		) AS KYBASE ";
                    tmp_sql = tmp_sql + " 	ON  KYSQREF.bk_no = KYBASE.bk_no ";
                    tmp_sql = tmp_sql + " 	AND KYSQREF.hy_no = KYBASE.hy_no ";
                    tmp_sql = tmp_sql + " 	AND KYSQREF.ky_no = KYBASE.ky_no ";
                    tmp_sql = tmp_sql + " 	AND KYSQREF.ko_no = KYBASE.ko_no ";
                    tmp_sql = tmp_sql + " 	LEFT JOIN m_nkin AS MN ON KYSQREF.nkin_no = MN.nkin_no ";
                }

                return tmp_sql;

            }

        }

    }

    #endregion




    // 20160531 鍵情報移行処理の修正 -del sta
    // #Region "契約鍵情報"

    // Public Class Ky_kagi_Repository

    // Public Class SubConv
    // Implements IConv

    // ''' <summary>
    // ''' 【抽出クエリ】 2016.04.06 契約鍵情報の移行処理追加
    // ''' </summary>
    // ''' <returns></returns>
    // ''' <remarks></remarks>
    // Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

    // Dim tmp_sql As String = ""

    // sortstr = "[物件No],[部屋No],[鍵タイトル名]"

    // '抽出データの改行文字列を除去する暫定処理
    // Dim tmp_midheader As String = "[物件No],[部屋No],[鍵タイトル名],[鍵本数],[備考],[保管場所],[業者間での情報共有]"
    // Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
    // '抽出データの改行文字列を除去する暫定処理

    // 'tmp_sql = tmp_sql & " SELECT * FROM "
    // tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
    // tmp_sql = tmp_sql & " ( "
    // tmp_sql = tmp_sql & " 	SELECT "
    // tmp_sql = tmp_sql & " 		 bk_no AS [物件No] "
    // tmp_sql = tmp_sql & " 		,hy_no AS [部屋No] "
    // tmp_sql = tmp_sql & " 		/*,kagimei_no*/ "
    // tmp_sql = tmp_sql & " 		,biko_titl AS [鍵タイトル名] "
    // tmp_sql = tmp_sql & " 		,kagi_honsu AS [鍵本数] "
    // tmp_sql = tmp_sql & " 		,kagi_no AS [備考] "
    // tmp_sql = tmp_sql & " 		,'' AS [保管場所] "
    // tmp_sql = tmp_sql & " 		,'' AS [業者間での情報共有] "
    // tmp_sql = tmp_sql & " 	FROM "
    // tmp_sql = tmp_sql & " 	( "
    // tmp_sql = tmp_sql & " 		/*契約Noが最大のデータに紐付く契約鍵情報を取得 sta*/ "
    // tmp_sql = tmp_sql & " 		SELECT * FROM ky_kagi AS KYKAGI "
    // tmp_sql = tmp_sql & " 		WHERE EXISTS "
    // tmp_sql = tmp_sql & " 			( "
    // tmp_sql = tmp_sql & " 				/*契約Noが最大のデータを取得 sta*/ "
    // tmp_sql = tmp_sql & " 				SELECT bk_no,hy_no,ky_no FROM "
    // tmp_sql = tmp_sql & " 				( "
    // tmp_sql = tmp_sql & " 					SELECT * FROM "
    // tmp_sql = tmp_sql & " 					( "
    // tmp_sql = tmp_sql & " 						SELECT "
    // tmp_sql = tmp_sql & " 							 bk_no "
    // tmp_sql = tmp_sql & " 							,hy_no "
    // tmp_sql = tmp_sql & " 							,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no ORDER BY ky_no DESC) AS [抽出用] "
    // tmp_sql = tmp_sql & " 							,ky_no "
    // tmp_sql = tmp_sql & " 						FROM ky_kagi AS KAGI "
    // tmp_sql = tmp_sql & " 					) AS VW "
    // tmp_sql = tmp_sql & " 					WHERE [抽出用] = 1 "
    // tmp_sql = tmp_sql & " 				) AS VW "
    // tmp_sql = tmp_sql & " 				/*契約Noが最大のデータを取得 end*/ "
    // tmp_sql = tmp_sql & " 				WHERE KYKAGI.bk_no = VW.bk_no "
    // tmp_sql = tmp_sql & " 				AND   KYKAGI.hy_no = VW.hy_no "
    // tmp_sql = tmp_sql & " 				AND   KYKAGI.ky_no = VW.ky_no "
    // tmp_sql = tmp_sql & " 			) "
    // tmp_sql = tmp_sql & " 		/*契約Noが最大のデータに紐付く契約鍵情報を取得 end*/ "
    // tmp_sql = tmp_sql & " 	) AS KYKAGITOTAL "
    // tmp_sql = tmp_sql & " 	LEFT JOIN "
    // tmp_sql = tmp_sql & " 		( "
    // tmp_sql = tmp_sql & " 			SELECT * FROM m_biko WHERE biko_kbn = 4	 "
    // tmp_sql = tmp_sql & " 		) AS KAGITITLE "
    // tmp_sql = tmp_sql & " 	ON KYKAGITOTAL.kagimei_no = KAGITITLE.biko_no "
    // tmp_sql = tmp_sql & " 	WHERE ISNULL(kagi_no,'') <> '' "
    // tmp_sql = tmp_sql & " ) AS VW "

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

    // #End Region
    // 20160531 鍵情報移行処理の修正 -del end

}