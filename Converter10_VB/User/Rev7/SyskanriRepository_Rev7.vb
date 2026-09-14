Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "初期設定_基本"

    Public Class M_kan_Base_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[No],[都道府県],[市区町村],[運用開始年月],[消費税額の算出時端数処理],[新規契約時解約時の日割り額丸め桁数],[新規契約時解約時の日割り額端数処理],[管理手数料 (料率別対象額) の算出時丸め桁数],[管理手数料 (料率別対象額) の算出時端数処理],[上記以外の金額計算(ポータル連動を除く)端数処理],[面積 (坪数から面積の算出時)端数処理],[坪数 (面積から坪数の算出時)端数処理],[月別の日数1],[月別の日数2],[月別の日数3],[月別の日数4],[月別の日数5],[月別の日数6],[月別の日数7],[月別の日数8],[月別の日数9],[月別の日数10],[月別の日数11],[月別の日数12],[物件画像をコピー],[周辺画像をコピー],[物件動画をコピー],[部屋情報をコピー],[更新業務初期値],[送金予定日1請求月],[送金予定日1請求日],[送金予定日1送金月],[送金予定日1送金日],[送金予定日2請求月],[送金予定日2請求日],[送金予定日2送金月],[送金予定日2送金日],[送金予定日3請求月],[送金予定日3請求日],[送金予定日3送金月],[送金予定日3送金日],[送金予定日4請求月],[送金予定日4請求日],[送金予定日4送金月],[送金予定日4送金日],[送金予定日5請求月],[送金予定日5請求日],[送金予定日5送金月],[送金予定日5送金日],[請求額に対し満額入金するまで送金しない],[滞納保証],[新規契約時の入金方法_通常更新解約月],[新規契約時の入金方法_契約契約更新一時金],[契約時の翌月分受取りの基準],[次回更新時の契約期間の算出基準],[契約終了通知期間],[契約終期の◯ヶ月前に自動更新],[解約月分賃料等の日割り計算を自動的に行う],[登録時に解約月の更新一時金を削除する],[解約精算書の作成タブの算出基準],[請求月],[請求計上日],[通知書発行ルール_契約者],[通知書発行ルール_家主],[滞納金に契約金を含めない],[請求担当者],[マルチヘッダー形式を使用する],[送金済未収金を表示しない],[口座振替一覧物件単位出力],[口座振込一覧_契約者物件単位出力],[口座振込一覧_家主物件単位出力],[照合方法],[照合方法_複数の場合],[空室は表示しない],[一括借上の場合に部屋毎明細を表示しない],[他社契約中の部屋は部屋自体を表示しない],[未収金を表示する],[契約時の契約状況],[部屋毎の明細表示],[印字設定_空室時],[印字設定_他社契約中],[印字設定_その他支払合計],[印字設定_控除額合計],[印字設定_物件名(前)],[印字設定_物件名(後)],[印字設定_物件合計],[印字設定_送金額合計],[印字設定_管理手数料備考表示部屋件数],[控除は入金項目単位で集計しない],[管理手数料_月毎に分けずにまとめて計上する],[初回契約一時金も滞納保証する]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 [no] AS [No] "
                tmp_sql = tmp_sql & " 		/************************************初期設定タブ sta************************************/ "
                tmp_sql = tmp_sql & " 		,ken_no AS [都道府県] "
                tmp_sql = tmp_sql & " 		,si_no AS [市区町村] "
                tmp_sql = tmp_sql & " 		,運用開始年月置換用文字列 AS [運用開始年月] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN hasu_zei = 1 THEN '_切り捨て' "
                tmp_sql = tmp_sql & " 			WHEN hasu_zei = 2 THEN '_四捨五入' "
                tmp_sql = tmp_sql & " 			WHEN hasu_zei = 3 THEN '_切り上げ' "
                tmp_sql = tmp_sql & " 		 END AS [消費税額の算出時端数処理] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN hiwari_hasutani = 1 THEN '_円未満' "
                tmp_sql = tmp_sql & " 			WHEN hiwari_hasutani = 2 THEN '_10円未満' "
                tmp_sql = tmp_sql & " 			WHEN hiwari_hasutani = 3 THEN '_100円未満' "
                tmp_sql = tmp_sql & " 			WHEN hiwari_hasutani = 4 THEN '_1000円未満' "
                tmp_sql = tmp_sql & " 			WHEN hiwari_hasutani = 5 THEN '_10000円未満' "
                tmp_sql = tmp_sql & " 		 END AS [新規契約時解約時の日割り額丸め桁数] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN hiwari_hasu = 1 THEN '_切り捨て' "
                tmp_sql = tmp_sql & " 			WHEN hiwari_hasu = 2 THEN '_四捨五入' "
                tmp_sql = tmp_sql & " 			WHEN hiwari_hasu = 3 THEN '_切り上げ' "
                tmp_sql = tmp_sql & " 		 END AS [新規契約時解約時の日割り額端数処理]		 "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN kanri_hasutani = 1 THEN '_円未満' "
                tmp_sql = tmp_sql & " 			WHEN kanri_hasutani = 2 THEN '_10円未満' "
                tmp_sql = tmp_sql & " 			WHEN kanri_hasutani = 3 THEN '_100円未満' "
                tmp_sql = tmp_sql & " 			WHEN kanri_hasutani = 4 THEN '_1000円未満' "
                tmp_sql = tmp_sql & " 			WHEN kanri_hasutani = 5 THEN '_10000円未満' "
                tmp_sql = tmp_sql & " 		 END AS [管理手数料 (料率別対象額) の算出時丸め桁数]	  "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN hasu_kanri = 1 THEN '_切り捨て' "
                tmp_sql = tmp_sql & " 			WHEN hasu_kanri = 2 THEN '_四捨五入' "
                tmp_sql = tmp_sql & " 			WHEN hasu_kanri = 3 THEN '_切り上げ' "
                tmp_sql = tmp_sql & " 		 END AS [管理手数料 (料率別対象額) の算出時端数処理] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN hasu_kngk = 1 THEN '_切り捨て' "
                tmp_sql = tmp_sql & " 			WHEN hasu_kngk = 2 THEN '_四捨五入' "
                tmp_sql = tmp_sql & " 			WHEN hasu_kngk = 3 THEN '_切り上げ' "
                tmp_sql = tmp_sql & " 		 END AS [上記以外の金額計算(ポータル連動を除く)端数処理] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN yobi_si1 = 1 THEN '_切り捨て' "
                tmp_sql = tmp_sql & " 			WHEN yobi_si1 = 2 THEN '_四捨五入' "
                tmp_sql = tmp_sql & " 			WHEN yobi_si1 = 3 THEN '_切り上げ' "
                tmp_sql = tmp_sql & " 		 END AS [面積 (坪数から面積の算出時)端数処理]  "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN yobi_si2 = 1 THEN '_切り捨て' "
                tmp_sql = tmp_sql & " 			WHEN yobi_si2 = 2 THEN '_四捨五入' "
                tmp_sql = tmp_sql & " 			WHEN yobi_si2 = 3 THEN '_切り上げ' "
                tmp_sql = tmp_sql & " 		 END AS [坪数 (面積から坪数の算出時)端数処理] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari1,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari1) + '日' END AS 月別の日数1 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari2,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari2) + '日' END AS 月別の日数2 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari3,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari3) + '日' END AS 月別の日数3 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari4,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari4) + '日' END AS 月別の日数4 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari5,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari5) + '日' END AS 月別の日数5 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari6,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari6) + '日' END AS 月別の日数6 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari7,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari7) + '日' END AS 月別の日数7 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari8,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari8) + '日' END AS 月別の日数8 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari9,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari9) + '日' END AS 月別の日数9 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari10,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari10) + '日' END AS 月別の日数10 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari11,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari11) + '日' END AS 月別の日数11 "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNULL(hiwari12,0) = 0 THEN '_自動' ELSE '_' + CONVERT(VARCHAR,hiwari12) + '日' END AS 月別の日数12 "
                tmp_sql = tmp_sql & " 		/************************************初期設定タブ end************************************/ "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 		/************************************物件管理タブ sta************************************/ "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_si16 = 0 THEN 'false' ELSE 'true' END AS [物件画像をコピー] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_si16 = 0 THEN 'false' ELSE 'true' END AS [周辺画像をコピー] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_si16 = 0 THEN 'false' ELSE 'true' END AS [物件動画をコピー] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_bt17 = 0 THEN 'false' ELSE 'true' END AS [部屋情報をコピー] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_bt19 = 0 THEN 1 ELSE 2 END AS [更新業務初期値] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(sime_kbn1,0) = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn1 = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn1 = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn1 = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn1 = 4 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn1 = 5 THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日1請求月] "
                tmp_sql = tmp_sql & " 		,ISNULL(sime1,-1) AS [送金予定日1請求日] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(som_kbn1,-1) = -1 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn1 = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn1 = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn1 = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn1 = 3 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn1 = 4 THEN 5 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn1 = 5 THEN 6 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日1送金月] "
                tmp_sql = tmp_sql & " 		,ISNULL(so_dd1,-1) AS [送金予定日1送金日] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(sime_kbn2,0) = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn2 = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn2 = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn2 = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn2 = 4 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn2 = 5 THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日2請求月] "
                tmp_sql = tmp_sql & " 		,ISNULL(sime2,-1) AS [送金予定日2請求日] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(som_kbn2,-1) = -1 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn2 = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn2 = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn2 = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn2 = 3 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn2 = 4 THEN 5 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn2 = 5 THEN 6 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日2送金月] "
                tmp_sql = tmp_sql & " 		,ISNULL(so_dd2,-1) AS [送金予定日2送金日]		 "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(sime_kbn3,0) = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn3 = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn3 = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn3 = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn3 = 4 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn3 = 5 THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日3請求月] "
                tmp_sql = tmp_sql & " 		,ISNULL(sime3,-1) AS [送金予定日3請求日] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(som_kbn3,-1) = -1 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn3 = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn3 = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn3 = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn3 = 3 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn3 = 4 THEN 5 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn3 = 5 THEN 6 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日3送金月] "
                tmp_sql = tmp_sql & " 		,ISNULL(so_dd3,-1) AS [送金予定日3送金日]		 "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(sime_kbn4,0) = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn4 = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn4 = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn4 = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn4 = 4 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn4 = 5 THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日4請求月] "
                tmp_sql = tmp_sql & " 		,ISNULL(sime4,-1) AS [送金予定日4請求日] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(som_kbn4,-1) = -1 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn4 = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn4 = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn4 = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn4 = 3 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn4 = 4 THEN 5 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn4 = 5 THEN 6 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日4送金月] "
                tmp_sql = tmp_sql & " 		,ISNULL(so_dd4,-1) AS [送金予定日4送金日] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(sime_kbn5,0) = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn5 = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn5 = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn5 = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn5 = 4 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sime_kbn5 = 5 THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日5請求月] "
                tmp_sql = tmp_sql & " 		,ISNULL(sime5,-1) AS [送金予定日5請求日] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(som_kbn5,-1) = -1 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn5 = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn5 = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn5 = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn5 = 3 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn5 = 4 THEN 5 "
                tmp_sql = tmp_sql & " 			WHEN som_kbn5 = 5 THEN 6 "
                tmp_sql = tmp_sql & " 		 END AS [送金予定日5送金月] "
                tmp_sql = tmp_sql & " 		,ISNULL(so_dd5,-1) AS [送金予定日5送金日] "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		/*HASPから取得するところ sta*/ "
                tmp_sql = tmp_sql & " 		,'' AS [請求額に対し満額入金するまで送金しない] "
                tmp_sql = tmp_sql & " 		/*HASPから取得するところ end*/ "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		,CASE WHEN tnh_umu = 0 THEN 'false' ELSE 'true' END AS [滞納保証] "
                tmp_sql = tmp_sql & " 		/************************************物件管理タブ end************************************/ "
                tmp_sql = tmp_sql & " 				 "
                tmp_sql = tmp_sql & " 		/************************************契約管理タブ sta************************************/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN nkbn_def = 1 THEN '現金' "
                tmp_sql = tmp_sql & " 			WHEN nkbn_def = 2 THEN '振込' "
                tmp_sql = tmp_sql & " 			WHEN nkbn_def = 3 THEN '振替(引落)' "
                tmp_sql = tmp_sql & " 			WHEN nkbn_def = 9 THEN 'その他' "
                tmp_sql = tmp_sql & " 		 END AS [新規契約時の入金方法_通常更新解約月] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 101) = 1 THEN '現金' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 101) = 2 THEN '振込' "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 101) = 3 THEN '振替(引落)' "
                tmp_sql = tmp_sql & " 		 END AS [新規契約時の入金方法_契約契約更新一時金]		  "
                tmp_sql = tmp_sql & " 		,yoku_ukedd AS [契約時の翌月分受取りの基準] "
                tmp_sql = tmp_sql & " 		,CASE WHEN kykosin_endkijyun = 0 THEN 1 ELSE 2 END AS [次回更新時の契約期間の算出基準] "
                tmp_sql = tmp_sql & " 		,end_tuuti_end AS [契約終了通知期間] "
                tmp_sql = tmp_sql & " 		,jido_tuki AS [契約終期の◯ヶ月前に自動更新] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_bt14 = 0 THEN 'true' ELSE 'false' END AS [解約月分賃料等の日割り計算を自動的に行う] "
                tmp_sql = tmp_sql & " 		,CASE WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 107) = 1 THEN 'true' ELSE 'false' END AS [登録時に解約月の更新一時金を削除する] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 104) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 104) = 2 AND (SELECT m_int2 FROM m_kan2 WHERE m_id = 104) = 0 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 104) = 2 AND (SELECT m_int2 FROM m_kan2 WHERE m_id = 104) = 1 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 104) = 0 THEN 4 "
                tmp_sql = tmp_sql & " 		 END AS [解約精算書の作成タブの算出基準]		 "
                tmp_sql = tmp_sql & " 		/************************************契約管理タブ end************************************/ "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		/************************************請求管理タブ sta************************************/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 99 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = -2 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = -1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 0  THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 1  THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 2  THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [請求月] "
                tmp_sql = tmp_sql & " 		,yobi_si5 AS [請求計上日] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN tuti_hakko = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN tuti_hakko = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN tuti_hakko = 3 THEN 4 "
                tmp_sql = tmp_sql & " 			ELSE 3 "
                tmp_sql = tmp_sql & " 		 END AS [通知書発行ルール_契約者] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN tuti_hakko = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN tuti_hakko = 3 THEN 4 "
                tmp_sql = tmp_sql & " 			ELSE 3 "
                tmp_sql = tmp_sql & " 		 END AS [通知書発行ルール_家主]		 "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 115) = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 115) = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN (SELECT m_int1 FROM m_kan2 WHERE m_id = 115) = 2 THEN 3 "
                tmp_sql = tmp_sql & " 		 END AS [滞納金に契約金を含めない] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN sq_tanto_kbn = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sq_tanto_kbn = 1 THEN 3 "
                tmp_sql = tmp_sql & " 		 END AS [請求担当者] "
                tmp_sql = tmp_sql & " 		,CASE WHEN use_multihed_sweb = 0 THEN 'false' ELSE 'true' END AS [マルチヘッダー形式を使用する] "
                tmp_sql = tmp_sql & " 		,CASE WHEN (SELECT m_bit1 FROM m_kan2 WHERE m_id = 100) = 0 THEN 'false' ELSE 'true' END AS [送金済未収金を表示しない] "
                tmp_sql = tmp_sql & " 		,CASE WHEN (SELECT m_bit1 FROM m_kan2 WHERE m_id = 119) = 0 THEN 'false' ELSE 'true' END AS [口座振替一覧物件単位出力] "
                tmp_sql = tmp_sql & " 		,CASE	 "
                tmp_sql = tmp_sql & " 			WHEN kozafuri_bkorder = -1 THEN 'true' "
                tmp_sql = tmp_sql & " 			WHEN kozafuri_bkorder = 0 THEN 'false' "
                tmp_sql = tmp_sql & " 		 END AS [口座振込一覧_契約者物件単位出力] "
                tmp_sql = tmp_sql & " 		,CASE	 "
                tmp_sql = tmp_sql & " 			WHEN kozafuri_bkorder = -1 THEN 'true' "
                tmp_sql = tmp_sql & " 			WHEN kozafuri_bkorder = 0 THEN 'false' "
                tmp_sql = tmp_sql & " 		 END AS [口座振込一覧_家主物件単位出力]	 "
                tmp_sql = tmp_sql & " 		/************************************請求管理タブ end************************************/ "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		/************************************入金管理タブ sta************************************/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN yobi_bt2 = 0 AND yobi_bt3 = 0 THEN '_振込依頼人名' "
                tmp_sql = tmp_sql & " 			WHEN yobi_bt2 = 1 AND yobi_bt3 = 0 THEN '_振込依頼人コード' "
                tmp_sql = tmp_sql & " 			WHEN yobi_bt2 = 0 AND yobi_bt3 = 1 THEN '_口座情報' "
                tmp_sql = tmp_sql & " 		 END AS [照合方法] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN syogo_hoho = 1 THEN '_振込口座' "
                tmp_sql = tmp_sql & " 			WHEN syogo_hoho = 2 THEN '_振込金額' "
                tmp_sql = tmp_sql & " 			WHEN syogo_hoho = 9 THEN '_照合しない' "
                tmp_sql = tmp_sql & " 		 END AS [照合方法_複数の場合] "
                tmp_sql = tmp_sql & " 		/************************************入金管理タブ end************************************/		 "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		/************************************送金管理タブ sta************************************/ "
                tmp_sql = tmp_sql & " 		,CASE WHEN kusitu_hihyoji = 0 THEN 'false' ELSE 'true' END AS [空室は表示しない] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_bt10 = 0 THEN 'false' ELSE 'true' END AS [一括借上の場合に部屋毎明細を表示しない] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_bt11 = 0 THEN 'false' ELSE 'true' END AS [他社契約中の部屋は部屋自体を表示しない] "
                tmp_sql = tmp_sql & " 		,CASE WHEN (SELECT m_bit1 FROM m_kan2 WHERE m_id = 103) = 0 THEN 'false' ELSE 'true' END AS [未収金を表示する] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_si8 = 0 THEN '_契約日を表示' WHEN yobi_si8 = 1 THEN '_契約始期を表示' END AS [契約時の契約状況] "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_si11 = 0 THEN '_請求先単位で表示する' WHEN yobi_si11 = 1 THEN '_契約単位で表示する' END AS [部屋毎の明細表示] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 1) AS [印字設定_空室時] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 9) AS [印字設定_他社契約中] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 3) AS [印字設定_その他支払合計] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 4) AS [印字設定_控除額合計] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 6) AS [印字設定_物件名(前)] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 7) AS [印字設定_物件名(後)] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 5) AS [印字設定_物件合計] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 8) AS [印字設定_送金額合計] "
                tmp_sql = tmp_sql & " 		,(SELECT biko_titl FROM m_biko WHERE biko_kbn = 9 AND biko_no = 12) AS [印字設定_管理手数料備考表示部屋件数]		 "
                tmp_sql = tmp_sql & " 		,CASE WHEN syusi_mei = 0 THEN 'false' ELSE 'true' END AS [控除は入金項目単位で集計しない] "
                tmp_sql = tmp_sql & " 		/*入金項目集約設定は別で行う*/ "
                tmp_sql = tmp_sql & " 		,CASE WHEN yobi_bt9 = 0 THEN 'false' ELSE 'true' END AS [管理手数料_月毎に分けずにまとめて計上する] "
                tmp_sql = tmp_sql & " 		,CASE WHEN keiyak_tnh_umu = 0 THEN 'false' ELSE 'true' END AS [初回契約一時金も滞納保証する] "
                tmp_sql = tmp_sql & " 		/************************************送金管理タブ end************************************/ "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 		/*******************************クレーム・修繕管理タブ sta*******************************/ "
                tmp_sql = tmp_sql & " 		/*移行対象外*/ "
                tmp_sql = tmp_sql & " 		/*******************************クレーム・修繕管理タブ end*******************************/ "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		/************************************書式設計タブ sta************************************/ "
                tmp_sql = tmp_sql & " 		/*移行対象外*/ "
                tmp_sql = tmp_sql & " 		/************************************書式設計タブ end************************************/ "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 	FROM m_kan "
                tmp_sql = tmp_sql & " ) AS VW "

                tmp_sql = tmp_sql.Replace("運用開始年月置換用文字列", "CONVERT(VARCHAR,CONVERT(DATETIME,'" & UnyoYMD & "'),111)")
                Return tmp_sql

            End Function

            Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

            End Sub

            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

            End Sub

        End Class

    End Class

#End Region

#Region "初期設定_税編集情報"

    Public Class M_kan_Zei_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[No],[税率],[適用開始],[適用終了],[備考]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 [no] AS [No] "
                tmp_sql = tmp_sql & " 		,zei_rit AS [税率] "
                tmp_sql = tmp_sql & " 		,CONVERT(NVARCHAR,start_ymd,111) AS [適用開始] "
                tmp_sql = tmp_sql & " 		,CONVERT(NVARCHAR,end_ymd,111) AS [適用終了] "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 	FROM m_zei "
                tmp_sql = tmp_sql & " ) AS VW "

                Return tmp_sql

            End Function

            Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

            End Sub

            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

            End Sub

        End Class

    End Class

#End Region

#Region "初期設定_変換文字情報"

    Public Class M_kan_Henkanmoji_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[No],[変換前],[変換後]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		[no] AS [No] "
                tmp_sql = tmp_sql & " 		,REPLACE(REPLACE(mae,'''','シングル'),'" & """""""" & "','ダブル') AS [変換前] "
                tmp_sql = tmp_sql & " 		,REPLACE(REPLACE(ato,'''','シングル'),'" & """""""" & "','ダブル') AS [変換後] "
                tmp_sql = tmp_sql & " 	FROM m_henkan "
                tmp_sql = tmp_sql & " 	WHERE henkan_kbn = 1	/*1のみだが念の為*/ "
                tmp_sql = tmp_sql & " ) AS VW "

                Return tmp_sql

            End Function

            Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

            End Sub

            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

            End Sub

        End Class

    End Class

#End Region

#Region "初期設定_入金項目集約情報"

    Public Class M_kan_Nkinkomkmerge_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[集約区分],[入金項目集約設定No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[集約区分],[入金項目集約設定No],[入金項目集約設定名],[入金項目区分1],[入金項目No1],[入金項目名1],[入金項目区分2],[入金項目No2],[入金項目名2],[入金項目区分3],[入金項目No3],[入金項目名3],[入金項目区分4],[入金項目No4],[入金項目名4],[入金項目区分5],[入金項目No5],[入金項目名5],[入金項目区分6],[入金項目No6],[入金項目名6],[入金項目区分7],[入金項目No7],[入金項目名7],[入金項目区分8],[入金項目No8],[入金項目名8],[入金項目区分9],[入金項目No9],[入金項目名9],[入金項目区分10],[入金項目No10],[入金項目名10],[入金項目区分11],[入金項目No11],[入金項目名11],[入金項目区分12],[入金項目No12],[入金項目名12],[入金項目区分13],[入金項目No13],[入金項目名13],[入金項目区分14],[入金項目No14],[入金項目名14],[入金項目区分15],[入金項目No15],[入金項目名15],[入金項目区分16],[入金項目No16],[入金項目名16],[入金項目区分17],[入金項目No17],[入金項目名17],[入金項目区分18],[入金項目No18],[入金項目名18],[入金項目区分19],[入金項目No19],[入金項目名19],[入金項目区分20],[入金項目No20],[入金項目名20]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 [集約区分] "
                tmp_sql = tmp_sql & " 		,[入金項目集約設定No] "
                tmp_sql = tmp_sql & " 		,[入金項目集約設定名] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 1 THEN [入金項目区分] ELSE '' END) AS [入金項目区分1] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 1 THEN [入金項目No] ELSE '' END) AS [入金項目No1] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 1 THEN [入金項目名] ELSE '' END) AS [入金項目名1] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 2 THEN [入金項目区分] ELSE '' END) AS [入金項目区分2] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 2 THEN [入金項目No] ELSE '' END) AS [入金項目No2] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 2 THEN [入金項目名] ELSE '' END) AS [入金項目名2] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 3 THEN [入金項目区分] ELSE '' END) AS [入金項目区分3] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 3 THEN [入金項目No] ELSE '' END) AS [入金項目No3] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 3 THEN [入金項目名] ELSE '' END) AS [入金項目名3] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 4 THEN [入金項目区分] ELSE '' END) AS [入金項目区分4] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 4 THEN [入金項目No] ELSE '' END) AS [入金項目No4] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 4 THEN [入金項目名] ELSE '' END) AS [入金項目名4] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 5 THEN [入金項目区分] ELSE '' END) AS [入金項目区分5] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 5 THEN [入金項目No] ELSE '' END) AS [入金項目No5] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 5 THEN [入金項目名] ELSE '' END) AS [入金項目名5] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 6 THEN [入金項目区分] ELSE '' END) AS [入金項目区分6] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 6 THEN [入金項目No] ELSE '' END) AS [入金項目No6] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 6 THEN [入金項目名] ELSE '' END) AS [入金項目名6] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 7 THEN [入金項目区分] ELSE '' END) AS [入金項目区分7] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 7 THEN [入金項目No] ELSE '' END) AS [入金項目No7] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 7 THEN [入金項目名] ELSE '' END) AS [入金項目名7] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 8 THEN [入金項目区分] ELSE '' END) AS [入金項目区分8] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 8 THEN [入金項目No] ELSE '' END) AS [入金項目No8] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 8 THEN [入金項目名] ELSE '' END) AS [入金項目名8] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 9 THEN [入金項目区分] ELSE '' END) AS [入金項目区分9] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 9 THEN [入金項目No] ELSE '' END) AS [入金項目No9] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 9 THEN [入金項目名] ELSE '' END) AS [入金項目名9] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 10 THEN [入金項目区分] ELSE '' END) AS [入金項目区分10] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 10 THEN [入金項目No] ELSE '' END) AS [入金項目No10] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 10 THEN [入金項目名] ELSE '' END) AS [入金項目名10] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 11 THEN [入金項目区分] ELSE '' END) AS [入金項目区分11] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 11 THEN [入金項目No] ELSE '' END) AS [入金項目No11] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 11 THEN [入金項目名] ELSE '' END) AS [入金項目名11] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 12 THEN [入金項目区分] ELSE '' END) AS [入金項目区分12] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 12 THEN [入金項目No] ELSE '' END) AS [入金項目No12] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 12 THEN [入金項目名] ELSE '' END) AS [入金項目名12] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 13 THEN [入金項目区分] ELSE '' END) AS [入金項目区分13] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 13 THEN [入金項目No] ELSE '' END) AS [入金項目No13] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 13 THEN [入金項目名] ELSE '' END) AS [入金項目名13] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 14 THEN [入金項目区分] ELSE '' END) AS [入金項目区分14] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 14 THEN [入金項目No] ELSE '' END) AS [入金項目No14] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 14 THEN [入金項目名] ELSE '' END) AS [入金項目名14] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 15 THEN [入金項目区分] ELSE '' END) AS [入金項目区分15] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 15 THEN [入金項目No] ELSE '' END) AS [入金項目No15] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 15 THEN [入金項目名] ELSE '' END) AS [入金項目名15] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 16 THEN [入金項目区分] ELSE '' END) AS [入金項目区分16] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 16 THEN [入金項目No] ELSE '' END) AS [入金項目No16] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 16 THEN [入金項目名] ELSE '' END) AS [入金項目名16] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 17 THEN [入金項目区分] ELSE '' END) AS [入金項目区分17] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 17 THEN [入金項目No] ELSE '' END) AS [入金項目No17] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 17 THEN [入金項目名] ELSE '' END) AS [入金項目名17] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 18 THEN [入金項目区分] ELSE '' END) AS [入金項目区分18] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 18 THEN [入金項目No] ELSE '' END) AS [入金項目No18] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 18 THEN [入金項目名] ELSE '' END) AS [入金項目名18] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 19 THEN [入金項目区分] ELSE '' END) AS [入金項目区分19] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 19 THEN [入金項目No] ELSE '' END) AS [入金項目No19] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 19 THEN [入金項目名] ELSE '' END) AS [入金項目名19] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 20 THEN [入金項目区分] ELSE '' END) AS [入金項目区分20] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 20 THEN [入金項目No] ELSE '' END) AS [入金項目No20] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN [集約No] = 20 THEN [入金項目名] ELSE '' END) AS [入金項目名20] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		/*収支一覧の抽出*/ "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 '収支一覧' AS [集約区分] "
                tmp_sql = tmp_sql & " 			,KOMKNAME.prt_meino AS [入金項目集約設定No] "
                tmp_sql = tmp_sql & " 			,KOMKNAME.prt_kmkname AS [入金項目集約設定名] "
                tmp_sql = tmp_sql & " 			,CASE "
                tmp_sql = tmp_sql & " 				WHEN nkin_kbn = 7 OR nkin_kbn = 8 OR nkin_kbn = 9 then 9 "
                tmp_sql = tmp_sql & " 				ELSE nkin_kbn "
                tmp_sql = tmp_sql & " 			 END AS [入金項目区分] "
                tmp_sql = tmp_sql & " 			,PRNKIN.nkin_no AS [入金項目No] "
                tmp_sql = tmp_sql & " 			,MN.nkin_name AS [入金項目名] "
                tmp_sql = tmp_sql & " 			,PRNKIN.merge_no + 1 AS [集約No] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 prt_no "
                tmp_sql = tmp_sql & " 				,prt_meino "
                tmp_sql = tmp_sql & " 				,prt_kmkname "
                tmp_sql = tmp_sql & " 			FROM prt_kmk "
                tmp_sql = tmp_sql & " 			WHERE prt_no = 15030 "
                tmp_sql = tmp_sql & " 			AND   ISNULL(prt_kmkname,'') <> '' "
                tmp_sql = tmp_sql & " 		) AS KOMKNAME "
                tmp_sql = tmp_sql & " 		LEFT JOIN prt_kmknkin AS PRNKIN "
                tmp_sql = tmp_sql & " 		ON  KOMKNAME.prt_no = PRNKIN.prt_no "
                tmp_sql = tmp_sql & " 		AND KOMKNAME.prt_meino = PRNKIN.prt_meino "
                tmp_sql = tmp_sql & " 		LEFT JOIN m_nkin AS MN ON PRNKIN.nkin_no = MN.nkin_no "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		/*物件別収支一覧の抽出*/ "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 '物件別収支一覧' AS [集約区分] "
                tmp_sql = tmp_sql & " 			,KOMKNAME.prt_meino AS [入金項目集約設定No] "
                tmp_sql = tmp_sql & " 			,KOMKNAME.prt_kmkname AS [入金項目集約設定名] "
                tmp_sql = tmp_sql & " 			,CASE "
                tmp_sql = tmp_sql & " 				WHEN nkin_kbn = 7 OR nkin_kbn = 8 OR nkin_kbn = 9 then 9 "
                tmp_sql = tmp_sql & " 				ELSE nkin_kbn "
                tmp_sql = tmp_sql & " 			 END AS [入金項目区分] "
                tmp_sql = tmp_sql & " 			,PRNKIN.nkin_no AS [入金項目No] "
                tmp_sql = tmp_sql & " 			,MN.nkin_name AS [入金項目名] "
                tmp_sql = tmp_sql & " 			,PRNKIN.merge_no + 1 AS [集約No] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 prt_no "
                tmp_sql = tmp_sql & " 				,prt_meino "
                tmp_sql = tmp_sql & " 				,prt_kmkname "
                tmp_sql = tmp_sql & " 			FROM prt_kmk "
                tmp_sql = tmp_sql & " 			WHERE prt_no = 15035 "
                tmp_sql = tmp_sql & " 			AND   ISNULL(prt_kmkname,'') <> '' "
                tmp_sql = tmp_sql & " 		) AS KOMKNAME "
                tmp_sql = tmp_sql & " 		LEFT JOIN prt_kmknkin AS PRNKIN "
                tmp_sql = tmp_sql & " 		ON  KOMKNAME.prt_no = PRNKIN.prt_no "
                tmp_sql = tmp_sql & " 		AND KOMKNAME.prt_meino = PRNKIN.prt_meino "
                tmp_sql = tmp_sql & " 		LEFT JOIN m_nkin AS MN ON PRNKIN.nkin_no = MN.nkin_no "
                tmp_sql = tmp_sql & " 	) AS TOTAL "
                tmp_sql = tmp_sql & " 	GROUP BY [集約区分],[入金項目集約設定No],[入金項目集約設定名] "
                tmp_sql = tmp_sql & " ) AS VW "

                Return tmp_sql

            End Function

            Public Function Chk_Data(ByVal tblname As String, ByVal keyvalue As String, ByVal list As List(Of String), ByVal hash_chkbefore As Hashtable, ByRef hash_chkafter As Hashtable, ByRef conditioncnt As Integer) As Boolean Implements IConv.Chk_Data

            End Function

            Public Sub Get_BaseKey(ByVal sqlcnnv10 As SqlConnection, ByVal tblname As String, ByRef list_basedata As List(Of String)) Implements IConv.Get_BaseKey

            End Sub

            Public Function Set_Vari(sqlcnnv10 As SqlConnection, filename As String, sheetname As String, ByRef midrowcnt As Integer, ByRef cvrowcnt As Integer, ByRef conditioncnt As Integer) As Boolean Implements IConv.Set_Vari

            End Function

            Public Sub Get_ExistData(sqlcnnv10 As SqlConnection, tblname As String, ByRef list_existdata As List(Of String)) Implements IConv.Get_ExistData

            End Sub

        End Class

    End Class

#End Region

    

    

    



End Namespace
