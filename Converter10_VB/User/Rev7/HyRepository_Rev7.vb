Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "部屋基本情報"

    Public Class Hy_mst_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[部屋分類名],[間取り],[間取り備考],[専有実面積(m2)],[専有実面積(坪)],[床面積(m2)],[床面積(坪)],[専有登記面積(m2)],[専有登記面積(坪)],[登記延床面積(m2)],[登記延床面積(坪)],[バルコニー面積(m2)],[バルコニー面積(坪)],[店舗付き住宅店舗部分面積(m2)],[店舗付き住宅店舗部分面積(坪)],[店舗付き住宅住宅部分面積(m2)],[店舗付き住宅住宅部分面積(坪)],[解約受付区分],[解約受付月数],[解約受付日数],[解約受付日にち],[自社Web内部キー採番ID]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		/*,crui_no AS [部屋分類]*/ "
                tmp_sql = tmp_sql & " 		,crui_name AS [部屋分類名] "
                tmp_sql = tmp_sql & " 		,madori AS [間取り]				/*数値と数値以外の文字列を分割する等、複雑な処理が必要なためプログラム内部で変換処理する*/ "
                tmp_sql = tmp_sql & " 		/*,madori AS [間取り(LDK種類)]	madoriをプログラム内で変換処理して設定する*/ "
                tmp_sql = tmp_sql & " 		,'' AS [間取り備考] "
                tmp_sql = tmp_sql & " 		,g_sen_men AS [専有実面積(m2)] "
                tmp_sql = tmp_sql & " 		,g_sen_tubo AS [専有実面積(坪)] "
                tmp_sql = tmp_sql & " 		,g_nobe_men AS [床面積(m2)] "
                tmp_sql = tmp_sql & " 		,g_nobe_tubo AS [床面積(坪)] "
                tmp_sql = tmp_sql & " 		,t_sen_men AS [専有登記面積(m2)] "
                tmp_sql = tmp_sql & " 		,t_sen_tubo AS [専有登記面積(坪)] "
                tmp_sql = tmp_sql & " 		,t_nobe_men AS [登記延床面積(m2)] "
                tmp_sql = tmp_sql & " 		,t_nobe_tubo AS [登記延床面積(坪)] "
                tmp_sql = tmp_sql & " 		,baru_men AS [バルコニー面積(m2)] "
                tmp_sql = tmp_sql & " 		,baru_tubo AS [バルコニー面積(坪)] "
                tmp_sql = tmp_sql & " 		,men_tenpo AS [店舗付き住宅店舗部分面積(m2)] "
                tmp_sql = tmp_sql & " 		,'' AS [店舗付き住宅店舗部分面積(坪)] "
                tmp_sql = tmp_sql & " 		,men_jukyo AS [店舗付き住宅住宅部分面積(m2)] "
                tmp_sql = tmp_sql & " 		,'' AS [店舗付き住宅住宅部分面積(坪)] "
                tmp_sql = tmp_sql & " 		,1 AS [解約受付区分] "
                tmp_sql = tmp_sql & " 		,3 AS [解約受付月数] "
                tmp_sql = tmp_sql & " 		,'' AS [解約受付日数] "
                tmp_sql = tmp_sql & " 		,'' AS [解約受付日にち] "
                tmp_sql = tmp_sql & " 		,wmp_id AS [自社Web内部キー採番ID]	/*20161012 革命10アップデートに伴う修正 add*/ "
                tmp_sql = tmp_sql & " 	FROM hy_mst AS HY "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_crui AS HYRUI ON HY.crui_no = HYRUI.crui_no "
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

#Region "部屋詳細情報"

    Public Class Hy_mst_syosai_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[部屋業務期間有無],[所在階１],[所在階２],[所在階３],[地下フラグ１],[地下フラグ２],[地下フラグ３],[向き],[角部屋フラグ],[バルコニー方向],[状況],[入居状況備考],[入居可能日を時期で指定],[入居可能時期(年月)],[入居可能時期(上旬・中旬・下旬)],[入居条件],[広告内容確認日],[募集可能種別(しない・する)],[フリーレント有無],[フリーレントカ月],[フリーレント詳細],[保険-保険の利用],[保険-保険期間],[保険-保険料],[保険-保険の備考],[取引形態種別],[客付け状態可否],[手数料負担割合貸主(%)],[手数料負担割合借主(%)],[手数料配分元付(%)],[手数料配分先物(%)],[客付会社への物件コメント],[業者間広告広告活動種別],[情報元業者No],[広告料有無],[広告料上限額],[広告料条件内容],[備考],[部屋住所の変更フラグ],[部屋住所(丁目)],[部屋住所(丁目区分)],[部屋住所(町地域)],[部屋住所(その他)],[保証会社],[保証内容],[契約の種類],[定期借家契約(期間)],[定期借家契約(期日)],[駐車場備考],[バイク駐車場備考],[駐輪場備考],[共通セールスポイント],[不動産検索サイト掲載設定-物件名],[不動産検索サイト掲載設定-部屋NO],[不動産検索サイト掲載設定-丁番地以下],[不動産検索サイト掲載設定-地図上],[賃貸保証利用区分],[広告料上限額区分],[広告料上限率],[広告料上限額税区分],[共通セールスポイント使用フラグ],[複数階有りのフラグ],[支店NO],[自社担当者No],[次回更新時の初期値],[登記情報の日付],[所有権にかかる権利有無],[所有権にかかる権利の種類],[所有権以外の権利有無],[BtoBプラグイングループ設定区分],[自社Webオススメ物件表示],[解約日],[退去日],[使用目的],[新築区分],[入居時期区分],[取引自社区分],[契約期間区分]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 HY.bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,0 AS [部屋業務期間有無] "
                tmp_sql = tmp_sql & " 		/*20160818 カンマが含まれている場合のエラー対応 chg sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		/*階数の正負判定を数値の大小ではなく関数で行うように修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu) = 1 AND kaisu < 0 THEN CONVERT(VARCHAR,kaisu * (-1)) ELSE kaisu END AS [所在階１] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu2) = 1 AND kaisu2 < 0 THEN CONVERT(VARCHAR,kaisu2 * (-1)) ELSE kaisu2 END AS [所在階２] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu3) = 1 AND kaisu3 < 0 THEN CONVERT(VARCHAR,kaisu3 * (-1)) ELSE kaisu3 END AS [所在階３] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu) = 1 AND kaisu < 0 THEN 1 ELSE 0 END AS [地下フラグ１] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu2) = 1 AND kaisu2 < 0 THEN 1 ELSE 0 END AS [地下フラグ２] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu3) = 1 AND kaisu3 < 0 THEN 1 ELSE 0 END AS [地下フラグ３] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu) = 1 AND SIGN(kaisu) = -1 THEN CONVERT(VARCHAR,kaisu * (-1)) ELSE kaisu END AS [所在階１] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu2) = 1 AND SIGN(kaisu2) = -1 THEN CONVERT(VARCHAR,kaisu2 * (-1)) ELSE kaisu2 END AS [所在階２] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu3) = 1 AND SIGN(kaisu3) = -1 THEN CONVERT(VARCHAR,kaisu3 * (-1)) ELSE kaisu3 END AS [所在階３] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu) = 1 AND SIGN(kaisu) = -1 THEN 1 ELSE 0 END AS [地下フラグ１] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu2) = 1 AND SIGN(kaisu2) = -1 THEN 1 ELSE 0 END AS [地下フラグ２] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu3) = 1 AND SIGN(kaisu3) = -1 THEN 1 ELSE 0 END AS [地下フラグ３] "
                tmp_sql = tmp_sql & " 		/*階数の正負判定を数値の大小ではなく関数で行うように修正 chg end*/ "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu) = 1 THEN "
                tmp_sql = tmp_sql & " 			CASE "
                tmp_sql = tmp_sql & " 				WHEN REPLACE(kaisu,',','') <> kaisu THEN kaisu "
                tmp_sql = tmp_sql & " 				WHEN SIGN(kaisu) = -1 THEN CONVERT(VARCHAR,kaisu * (-1)) "
                tmp_sql = tmp_sql & " 			ELSE kaisu "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & " 		 ELSE kaisu "
                tmp_sql = tmp_sql & " 		 END AS [所在階１] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu2) = 1 THEN "
                tmp_sql = tmp_sql & " 			CASE "
                tmp_sql = tmp_sql & " 				WHEN REPLACE(kaisu2,',','') <> kaisu2 THEN kaisu2 "
                tmp_sql = tmp_sql & " 				WHEN SIGN(kaisu2) = -1 THEN CONVERT(VARCHAR,kaisu2 * (-1)) "
                tmp_sql = tmp_sql & " 			ELSE kaisu2 "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & " 		 ELSE kaisu2 "
                tmp_sql = tmp_sql & " 		 END AS [所在階２] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu3) = 1 THEN "
                tmp_sql = tmp_sql & " 			CASE "
                tmp_sql = tmp_sql & " 				WHEN REPLACE(kaisu3,',','') <> kaisu3 THEN kaisu3 "
                tmp_sql = tmp_sql & " 				WHEN SIGN(kaisu3) = -1 THEN CONVERT(VARCHAR,kaisu3 * (-1)) "
                tmp_sql = tmp_sql & " 			ELSE kaisu3 "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & " 		 ELSE kaisu3 "
                tmp_sql = tmp_sql & " 		 END AS [所在階３] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu) = 1 THEN "
                tmp_sql = tmp_sql & " 			CASE "
                tmp_sql = tmp_sql & " 				WHEN REPLACE(kaisu,',','') <> kaisu THEN 0 "
                tmp_sql = tmp_sql & " 				WHEN SIGN(kaisu) = -1 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE 0 "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & " 		 ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [地下フラグ１] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu2) = 1 THEN "
                tmp_sql = tmp_sql & " 			CASE "
                tmp_sql = tmp_sql & " 				WHEN REPLACE(kaisu2,',','') <> kaisu2 THEN 0 "
                tmp_sql = tmp_sql & " 				WHEN SIGN(kaisu2) = -1 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE 0 "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & " 		 ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [地下フラグ２] "
                tmp_sql = tmp_sql & " 		,CASE WHEN ISNUMERIC(kaisu3) = 1 THEN "
                tmp_sql = tmp_sql & " 			CASE "
                tmp_sql = tmp_sql & " 				WHEN REPLACE(kaisu3,',','') <> kaisu3 THEN 0 "
                tmp_sql = tmp_sql & " 				WHEN SIGN(kaisu3) = -1 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE 0 "
                tmp_sql = tmp_sql & " 			END "
                tmp_sql = tmp_sql & " 		 ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [地下フラグ３] "
                tmp_sql = tmp_sql & " 		/*20160818 カンマが含まれている場合のエラー対応 chg end*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN muki = '東' THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN muki = '西' THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN muki = '南' THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN muki = '北' THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN muki = '北東' THEN 5 "
                tmp_sql = tmp_sql & " 			WHEN muki = '南東' THEN 6 "
                tmp_sql = tmp_sql & " 			WHEN muki = '北西' THEN 7 "
                tmp_sql = tmp_sql & " 			WHEN muki = '南西' THEN 8 "
                tmp_sql = tmp_sql & " 			ELSE NULL "
                tmp_sql = tmp_sql & " 		 END AS [向き] "
                tmp_sql = tmp_sql & " 		,0 AS [角部屋フラグ] "
                tmp_sql = tmp_sql & " 		/*20160829 革命10バージョンアップに伴う修正 del*/ "
                tmp_sql = tmp_sql & " 		/*,-1 AS [採光]*/ "
                tmp_sql = tmp_sql & " 		,-1 AS [バルコニー方向] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN k_jyokyo_cd = 4 THEN NULL "
                tmp_sql = tmp_sql & " 			WHEN k_jyokyo_cd = 99 AND tasya_k_jyokyo_cd = 1 THEN 99 "
                tmp_sql = tmp_sql & " 			WHEN k_jyokyo_cd = 99 AND tasya_k_jyokyo_cd = 2 THEN 98 "
                tmp_sql = tmp_sql & " 		 ELSE NULL "
                tmp_sql = tmp_sql & " 		 END AS [状況] "
                tmp_sql = tmp_sql & " 		,'' AS [入居状況備考] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN nyukyo_jyotai = 2 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE '' "
                tmp_sql = tmp_sql & " 		 END AS [入居可能日を時期で指定] "
                tmp_sql = tmp_sql & " 		,nyukyo_ym AS [入居可能時期(年月)] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN nyukyo_kigen = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			ELSE nyukyo_kigen "
                tmp_sql = tmp_sql & " 		 END AS [入居可能時期(上旬・中旬・下旬)] "
                tmp_sql = tmp_sql & " 		,joken AS [入居条件] "
                tmp_sql = tmp_sql & " 		,'' AS [広告内容確認日] "
                tmp_sql = tmp_sql & " 		/*20161020 部屋詳細情報の募集可能フラグの移行仕様修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*,0 AS [募集可能種別(しない・する)]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN k_jyokyo_cd = 4 THEN 1	/*空室*/ "
                tmp_sql = tmp_sql & " 			ELSE 0						/*空室以外*/ "
                tmp_sql = tmp_sql & " 		 END AS [募集可能種別(しない・する)] "
                tmp_sql = tmp_sql & " 		/*20161020 部屋詳細情報の募集可能フラグの移行仕様修正 chg end*/ "
                tmp_sql = tmp_sql & " 		,0 AS [フリーレント有無] "
                tmp_sql = tmp_sql & " 		,'' AS [フリーレントカ月] "
                tmp_sql = tmp_sql & " 		,'' AS [フリーレント詳細] "
                tmp_sql = tmp_sql & " 		/*20160720 連動情報構築 chg sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,-1 AS [保険-保険の利用] "
                tmp_sql = tmp_sql & " 		,'' AS [保険-保険期間] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(hoken_kikan,0) = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			ELSE 2 "
                tmp_sql = tmp_sql & " 		 END AS [保険-保険の利用] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(hoken_kikan,0) = 0 THEN NULL "
                tmp_sql = tmp_sql & " 			ELSE hoken_kikan "
                tmp_sql = tmp_sql & " 		 END AS [保険-保険期間] "
                tmp_sql = tmp_sql & " 		/*20160720 連動情報構築 chg end*/ "
                tmp_sql = tmp_sql & " 		,'' AS [保険-保険料] "
                tmp_sql = tmp_sql & " 		,'' AS [保険-保険の備考] "
                tmp_sql = tmp_sql & " 		,toritaiyo AS [取引形態種別] "
                tmp_sql = tmp_sql & " 		,kyaku_tuke AS [客付け状態可否] "
                tmp_sql = tmp_sql & " 		,tesuryo_futan_kasi AS [手数料負担割合貸主(%)] "
                tmp_sql = tmp_sql & " 		,tesuryo_futan_kari AS [手数料負担割合借主(%)] "
                tmp_sql = tmp_sql & " 		,tesuryo_haibun_moto AS [手数料配分元付(%)] "
                tmp_sql = tmp_sql & " 		,tesuryo_haibun_kyaku AS [手数料配分先物(%)] "
                tmp_sql = tmp_sql & " 		,kyakutuke_bk_comment AS [客付会社への物件コメント] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kokoku_kbn,'') = '' THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kokoku_kbn,'') = '広告可(全て)' THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kokoku_kbn,'') = '広告可(インターネットのみ)' THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kokoku_kbn,'') = '広告可(チラシ・新聞・情報誌のみ)' THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kokoku_kbn,'') = '要確認' THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kokoku_kbn,'') = '不可' THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [業者間広告広告活動種別] "
                tmp_sql = tmp_sql & " 		,BK.gy_no AS [情報元業者No] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			/*2016.04.26 メインの方へも反映させる修正 -chg sta*/ "
                tmp_sql = tmp_sql & " 			/*WHEN kokokuhi_kin > 0 AND kokokuhi_tanni = '円' THEN 1*/ "
                tmp_sql = tmp_sql & " 			/*ELSE 0*/ "
                tmp_sql = tmp_sql & " 			WHEN kokokuhi_kin > 0 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 			/*2016.04.26 メインの方へも反映させる修正 -chg end*/ "
                tmp_sql = tmp_sql & " 		 END AS [広告料有無] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN kokokuhi_tanni = '円' OR kokokuhi_tanni = 'ヶ月' THEN kokokuhi_kin "
                tmp_sql = tmp_sql & " 			ELSE '' "
                tmp_sql = tmp_sql & " 		 END AS [広告料上限額] "
                tmp_sql = tmp_sql & " 		,'' AS [広告料条件内容] "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HY.add_banti <> '' OR HY.add_etc <> '' THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE 0 "
                tmp_sql = tmp_sql & " 		 END AS [部屋住所の変更フラグ] "
                tmp_sql = tmp_sql & " 		,'' AS [部屋住所(丁目)] "
                tmp_sql = tmp_sql & " 		,'' AS [部屋住所(丁目区分)] "
                tmp_sql = tmp_sql & " 		,HY.add_banti AS [部屋住所(町地域)] "
                tmp_sql = tmp_sql & " 		,HY.add_etc AS [部屋住所(その他)] "
                tmp_sql = tmp_sql & " 		,tintai_hosyoryo_gyo AS [保証会社] "
                tmp_sql = tmp_sql & " 		,tintai_hosyoryo_comment AS [保証内容] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(as_teiki_kbn,0) = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE 2 "
                tmp_sql = tmp_sql & " 		 END AS [契約の種類] "
                tmp_sql = tmp_sql & " 		/*20160523 更新期間取得方法の変更 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*部屋情報に更新期間が存在しない場合は物件情報から取得する*/ "
                tmp_sql = tmp_sql & " 		/*,as_teiki_kikan AS [定期借家契約(期間)]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(as_teiki_kikan,0) = 0 THEN ky_kikan * 12 "
                tmp_sql = tmp_sql & " 			ELSE as_teiki_kikan "
                tmp_sql = tmp_sql & " 		 END AS [定期借家契約(期間)] "
                tmp_sql = tmp_sql & " 		/*20160523 更新期間取得方法の変更 chg end*/ "
                tmp_sql = tmp_sql & " 		,as_teiki_ymd AS [定期借家契約(期日)] "
                tmp_sql = tmp_sql & " 		,new_tyusya_biko AS [駐車場備考] "
                tmp_sql = tmp_sql & " 		,'' AS [バイク駐車場備考] "
                tmp_sql = tmp_sql & " 		,'' AS [駐輪場備考] "
                tmp_sql = tmp_sql & " 		,sales_point AS [共通セールスポイント] "
                tmp_sql = tmp_sql & " 		/*20160720 連動情報構築 chg sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,'' AS [不動産検索サイト掲載設定-物件名] "
                tmp_sql = tmp_sql & " 		,'' AS [不動産検索サイト掲載設定-部屋NO] "
                tmp_sql = tmp_sql & " 		,'' AS [不動産検索サイト掲載設定-丁番地以下] "
                tmp_sql = tmp_sql & " 		,'' AS [不動産検索サイト掲載設定-地図上] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,BK.bk_name_open AS [不動産検索サイト掲載設定-物件名] "
                tmp_sql = tmp_sql & " 		,BK.hy_no_open AS [不動産検索サイト掲載設定-部屋NO] "
                tmp_sql = tmp_sql & " 		,BK.banti_open AS [不動産検索サイト掲載設定-丁番地以下] "
                tmp_sql = tmp_sql & " 		,BK.map_open AS [不動産検索サイト掲載設定-地図上] "
                tmp_sql = tmp_sql & " 		/*20160720 連動情報構築 chg end*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN tintai_hosyoryo_kbn = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN tintai_hosyoryo_kbn = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 		 END AS [賃貸保証利用区分] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN kokokuhi_tanni = '円' THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN kokokuhi_tanni = 'ヶ月' THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN kokokuhi_tanni = '％' THEN 3	 "
                tmp_sql = tmp_sql & " 			ELSE 1 "
                tmp_sql = tmp_sql & " 		 END AS [広告料上限額区分] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN kokokuhi_tanni = '％' THEN kokokuhi_kin "
                tmp_sql = tmp_sql & " 			ELSE '' "
                tmp_sql = tmp_sql & " 		 END AS [広告料上限率] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN kokokuhi_zei = '税込み' THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN kokokuhi_zei = '税抜き' THEN 2 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 		 END AS [広告料上限額税区分] "
                tmp_sql = tmp_sql & " 		/*20161012 共通セールス使用フラグのデフォルト値変更 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*,0 AS [共通セールスポイント使用フラグ]*/ "
                tmp_sql = tmp_sql & " 		,1 AS [共通セールスポイント使用フラグ] "
                tmp_sql = tmp_sql & " 		/*20161012 共通セールス使用フラグのデフォルト値変更 chg end*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kaisu,'') <> '' AND ISNULL(kaisu2,'') = '' AND ISNULL(kaisu3,'') = '' THEN 0 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kaisu,'') = '' AND ISNULL(kaisu2,'') <> '' AND ISNULL(kaisu3,'') = '' THEN 0 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kaisu,'') = '' AND ISNULL(kaisu2,'') = '' AND ISNULL(kaisu3,'') <> '' THEN 0 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(kaisu,'') = '' AND ISNULL(kaisu2,'') = '' AND ISNULL(kaisu3,'') = '' THEN 0 "
                tmp_sql = tmp_sql & " 			ELSE 1 "
                tmp_sql = tmp_sql & " 		 END AS [複数階有りのフラグ] "
                tmp_sql = tmp_sql & " 		/*2016.04.06 支店Noの取得処理を修正 -chg sta*/ "
                tmp_sql = tmp_sql & " 		/*,BK.siten_no AS [支店NO]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(BK.siten_no,0) = 0 THEN NULL "
                tmp_sql = tmp_sql & " 			ELSE [移行用自社No] "
                tmp_sql = tmp_sql & " 		 END AS [支店No] "
                tmp_sql = tmp_sql & " 		/*2016.04.06 支店Noの取得処理を修正 -chg end*/	 "
                tmp_sql = tmp_sql & " 		,BK.tanto_no AS [自社担当者No] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 部屋詳細情報 del sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,'' AS [確認事項(契約)責任者] "
                tmp_sql = tmp_sql & " 		,'' AS [確認事項(契約)内容確認日] "
                tmp_sql = tmp_sql & " 		,'' AS [確認事項(契約)印刷時] "
                tmp_sql = tmp_sql & " 		,'' AS [確認事項(解約)責任者] "
                tmp_sql = tmp_sql & " 		,'' AS [確認事項(解約)内容確認日] "
                tmp_sql = tmp_sql & " 		,'' AS [確認事項(解約)印刷時] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 部屋詳細情報 del end*/ "
                tmp_sql = tmp_sql & " 		,1 AS [次回更新時の初期値] "
                tmp_sql = tmp_sql & " 		,'' AS [登記情報の日付] "
                tmp_sql = tmp_sql & " 		,'' AS [所有権にかかる権利有無] "
                tmp_sql = tmp_sql & " 		,'' AS [所有権にかかる権利の種類] "
                tmp_sql = tmp_sql & " 		,'' AS [所有権以外の権利有無] "
                tmp_sql = tmp_sql & " 		/*20160720 連動情報構築 chg sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,0 AS [BtoBプラグイングループ設定区分] "
                tmp_sql = tmp_sql & " 		,0 AS [自社Webオススメ物件表示] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN group99_umu = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN group99_umu = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			ELSE 1 "
                tmp_sql = tmp_sql & " 		 END AS [BtoBプラグイングループ設定区分] "
                tmp_sql = tmp_sql & " 		,HY.yobi_bt3 AS [自社Webオススメ物件表示] "
                tmp_sql = tmp_sql & " 		/*20160720 連動情報構築 chg end*/ "
                tmp_sql = tmp_sql & " 		,'' AS [解約日] "
                tmp_sql = tmp_sql & " 		,'' AS [退去日] "
                tmp_sql = tmp_sql & " 		,'' AS [使用目的] "
                tmp_sql = tmp_sql & " 		/*20160825 部屋状況で新築フラグを制御する修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*,1 AS [新築区分]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN k_jyokyo_cd = 4 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN k_jyokyo_cd = 99 AND tasya_k_jyokyo_cd = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN k_jyokyo_cd = 99 AND tasya_k_jyokyo_cd = 2 THEN 2 "
                tmp_sql = tmp_sql & " 		 ELSE 1 "
                tmp_sql = tmp_sql & " 		 END AS [新築区分] "
                tmp_sql = tmp_sql & " 		/*20160825 部屋状況で新築フラグを制御する修正 chg end*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN nyukyo_jyotai = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN nyukyo_jyotai = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN nyukyo_jyotai = 2 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN nyukyo_jyotai = 3 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN nyukyo_jyotai = 4 THEN 2 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 		 END AS [入居時期区分] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN jisya_kbn = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN jisya_kbn = 1 THEN 2 "
                tmp_sql = tmp_sql & " 		 END AS [取引自社区分] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(as_teiki_kikan,'') <> '' AND ISNULL(as_teiki_ymd,'') <> '' THEN 1		/*両方設定されている場合は期間指定を優先*/ "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(as_teiki_kikan,'') <> '' AND ISNULL(as_teiki_ymd,'') = '' THEN 1		/*期間が指定されている場合は期間指定を設定*/ "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(as_teiki_kikan,'') = '' AND ISNULL(as_teiki_ymd,'') <> '' THEN 2		/*期限が指定されている場合は期限指定を設定*/ "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(as_teiki_kikan,'') = '' AND ISNULL(as_teiki_ymd,'') = '' THEN 1			/*両方設定されていない場合はデフォルト値を設定*/ "
                tmp_sql = tmp_sql & " 		 END AS [契約期間区分]	 "
                tmp_sql = tmp_sql & " 	FROM hy_mst AS HY "
                tmp_sql = tmp_sql & " 	LEFT JOIN bk_mst AS BK ON HY.bk_no = BK.bk_no "
                tmp_sql = tmp_sql & " 	/*2016.04.06 支店Noの取得処理を修正 -add sta*/ "
                tmp_sql = tmp_sql & " 	LEFT JOIN "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 siten_no "
                tmp_sql = tmp_sql & " 				,[移行用自社No抽出用] + (SELECT MAX(jisya_no) FROM m_jisya) AS [移行用自社No] "
                tmp_sql = tmp_sql & " 			FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT "
                tmp_sql = tmp_sql & " 					 siten_no "
                tmp_sql = tmp_sql & " 					,ROW_NUMBER()OVER(ORDER BY siten_no) AS [移行用自社No抽出用] "
                tmp_sql = tmp_sql & " 				FROM m_siten "
                tmp_sql = tmp_sql & " 			) AS VW "
                tmp_sql = tmp_sql & " 		) AS MS "
                tmp_sql = tmp_sql & " 	ON BK.siten_no = MS.siten_no "
                tmp_sql = tmp_sql & " 	/*2016.04.06 支店Noの取得処理を修正 -add end*/ "
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

#Region "部屋駐車場情報"

    Public Class Hy_mst_parking_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO],[駐車場区分]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[駐車場区分],[駐車場の空き数(手動設定用)],[駐車場の空き有無(手動設定用)],[駐車場料金区分(手動設定用)],[駐車場料金(手動設定用)],[駐車場料金税区分(手動設定用)],[駐車場の賃貸可能数]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,1 AS [駐車場区分] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場の空き数(手動設定用)] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_umu = '無し' THEN 5 "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_umu = '有り' THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_umu = '付近有り' THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_umu = '要問合せ(不明)' THEN 4 "
                tmp_sql = tmp_sql & " 			ELSE 2 "
                tmp_sql = tmp_sql & " 		 END AS [駐車場の空き有無(手動設定用)] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_rkinkbn = '込' THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_rkinkbn = '別途' THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_rkinkbn = '要問合せ(不明)' THEN 3 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 		 END AS [駐車場料金区分(手動設定用)] "
                tmp_sql = tmp_sql & " 		,new_tyusya_ryokin AS [駐車場料金(手動設定用)] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_zeikbn = '税込' THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN new_tyusya_zeikbn = '税別' THEN 2 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 		 END AS [駐車場料金税区分(手動設定用)] "
                tmp_sql = tmp_sql & " 		/*20161020 部屋駐車場移行仕様修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*,'' AS [駐車場の賃貸可能数]*/ "
                tmp_sql = tmp_sql & " 		,tyusya_daisu AS [駐車場の賃貸可能数] "
                tmp_sql = tmp_sql & " 		/*20161020 部屋駐車場移行仕様修正 chg end*/ "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,2 AS [駐車場区分] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場の空き数(手動設定用)] "
                tmp_sql = tmp_sql & " 		,2 AS [駐車場の空き有無(手動設定用)] "
                tmp_sql = tmp_sql & " 		,-1 AS [駐車場料金区分(手動設定用)] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場料金(手動設定用)] "
                tmp_sql = tmp_sql & " 		,-1 AS [駐車場料金税区分(手動設定用)] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場の賃貸可能数] "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,3 AS [駐車場区分] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場の空き数(手動設定用)] "
                tmp_sql = tmp_sql & " 		,2 AS [駐車場の空き有無(手動設定用)] "
                tmp_sql = tmp_sql & " 		,-1 AS [駐車場料金区分(手動設定用)] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場料金(手動設定用)] "
                tmp_sql = tmp_sql & " 		,-1 AS [駐車場料金税区分(手動設定用)] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場の賃貸可能数] "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
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

#Region "部屋特約情報(原状回復特約を含む)"

    Public Class Hy_tokuyaku_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO],[特約区分]"
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[特約区分],[内容1],[内容2],[内容3],[内容4],[内容5],[内容6],[内容7],[内容8],[内容9],[内容10],[内容11],[内容12],[内容13],[内容14],[内容15],[内容16],[内容17],[内容18],[内容19],[内容20],[内容21],[内容22],[内容23],[内容24],[内容25],[内容26],[内容27],[内容28],[内容29],[内容30],[内容31],[内容32],[内容33],[内容34],[内容35],[内容36],[内容37],[内容38],[内容39],[内容40],[内容41],[内容42],[内容43],[内容44],[内容45],[内容46],[内容47],[内容48],[内容49],[内容50]"
                '抽出データの改行文字列を除去する暫定処理

                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 			,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 			,1 AS [特約区分] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 1 THEN tokuyaku ELSE '' END) AS 内容1 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 2 THEN tokuyaku ELSE '' END) AS 内容2 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 3 THEN tokuyaku ELSE '' END) AS 内容3 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 4 THEN tokuyaku ELSE '' END) AS 内容4 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 5 THEN tokuyaku ELSE '' END) AS 内容5 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 6 THEN tokuyaku ELSE '' END) AS 内容6 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 7 THEN tokuyaku ELSE '' END) AS 内容7 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 8 THEN tokuyaku ELSE '' END) AS 内容8 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 9 THEN tokuyaku ELSE '' END) AS 内容9 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 10 THEN tokuyaku ELSE '' END) AS 内容10 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 11 THEN tokuyaku ELSE '' END) AS 内容11 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 12 THEN tokuyaku ELSE '' END) AS 内容12 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 13 THEN tokuyaku ELSE '' END) AS 内容13 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 14 THEN tokuyaku ELSE '' END) AS 内容14 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 15 THEN tokuyaku ELSE '' END) AS 内容15 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 16 THEN tokuyaku ELSE '' END) AS 内容16 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 17 THEN tokuyaku ELSE '' END) AS 内容17 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 18 THEN tokuyaku ELSE '' END) AS 内容18 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 19 THEN tokuyaku ELSE '' END) AS 内容19 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 20 THEN tokuyaku ELSE '' END) AS 内容20 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 21 THEN tokuyaku ELSE '' END) AS 内容21 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 22 THEN tokuyaku ELSE '' END) AS 内容22 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 23 THEN tokuyaku ELSE '' END) AS 内容23 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 24 THEN tokuyaku ELSE '' END) AS 内容24 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 25 THEN tokuyaku ELSE '' END) AS 内容25 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 26 THEN tokuyaku ELSE '' END) AS 内容26 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 27 THEN tokuyaku ELSE '' END) AS 内容27 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 28 THEN tokuyaku ELSE '' END) AS 内容28 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 29 THEN tokuyaku ELSE '' END) AS 内容29 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 30 THEN tokuyaku ELSE '' END) AS 内容30 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 31 THEN tokuyaku ELSE '' END) AS 内容31 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 32 THEN tokuyaku ELSE '' END) AS 内容32 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 33 THEN tokuyaku ELSE '' END) AS 内容33 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 34 THEN tokuyaku ELSE '' END) AS 内容34 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 35 THEN tokuyaku ELSE '' END) AS 内容35 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 36 THEN tokuyaku ELSE '' END) AS 内容36 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 37 THEN tokuyaku ELSE '' END) AS 内容37 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 38 THEN tokuyaku ELSE '' END) AS 内容38 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 39 THEN tokuyaku ELSE '' END) AS 内容39 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 40 THEN tokuyaku ELSE '' END) AS 内容40 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 41 THEN tokuyaku ELSE '' END) AS 内容41 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 42 THEN tokuyaku ELSE '' END) AS 内容42 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 43 THEN tokuyaku ELSE '' END) AS 内容43 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 44 THEN tokuyaku ELSE '' END) AS 内容44 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 45 THEN tokuyaku ELSE '' END) AS 内容45 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 46 THEN tokuyaku ELSE '' END) AS 内容46 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 47 THEN tokuyaku ELSE '' END) AS 内容47 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 48 THEN tokuyaku ELSE '' END) AS 内容48 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 49 THEN tokuyaku ELSE '' END) AS 内容49 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 50 THEN tokuyaku ELSE '' END) AS 内容50 "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no "
                tmp_sql = tmp_sql & " 				,hy_no "
                tmp_sql = tmp_sql & " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no ORDER BY bk_no,hy_no,mei_no) AS [特約行No] "
                tmp_sql = tmp_sql & " 				,tokuyaku "
                tmp_sql = tmp_sql & " 			FROM hy_tokuyaku "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 	) AS TOKUYAKU "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 			,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 			,2 AS [特約区分] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 1 THEN tokuyaku ELSE '' END) AS 内容1 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 2 THEN tokuyaku ELSE '' END) AS 内容2 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 3 THEN tokuyaku ELSE '' END) AS 内容3 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 4 THEN tokuyaku ELSE '' END) AS 内容4 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 5 THEN tokuyaku ELSE '' END) AS 内容5 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 6 THEN tokuyaku ELSE '' END) AS 内容6 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 7 THEN tokuyaku ELSE '' END) AS 内容7 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 8 THEN tokuyaku ELSE '' END) AS 内容8 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 9 THEN tokuyaku ELSE '' END) AS 内容9 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 10 THEN tokuyaku ELSE '' END) AS 内容10 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 11 THEN tokuyaku ELSE '' END) AS 内容11 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 12 THEN tokuyaku ELSE '' END) AS 内容12 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 13 THEN tokuyaku ELSE '' END) AS 内容13 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 14 THEN tokuyaku ELSE '' END) AS 内容14 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 15 THEN tokuyaku ELSE '' END) AS 内容15 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 16 THEN tokuyaku ELSE '' END) AS 内容16 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 17 THEN tokuyaku ELSE '' END) AS 内容17 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 18 THEN tokuyaku ELSE '' END) AS 内容18 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 19 THEN tokuyaku ELSE '' END) AS 内容19 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 20 THEN tokuyaku ELSE '' END) AS 内容20 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 21 THEN tokuyaku ELSE '' END) AS 内容21 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 22 THEN tokuyaku ELSE '' END) AS 内容22 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 23 THEN tokuyaku ELSE '' END) AS 内容23 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 24 THEN tokuyaku ELSE '' END) AS 内容24 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 25 THEN tokuyaku ELSE '' END) AS 内容25 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 26 THEN tokuyaku ELSE '' END) AS 内容26 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 27 THEN tokuyaku ELSE '' END) AS 内容27 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 28 THEN tokuyaku ELSE '' END) AS 内容28 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 29 THEN tokuyaku ELSE '' END) AS 内容29 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 30 THEN tokuyaku ELSE '' END) AS 内容30 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 31 THEN tokuyaku ELSE '' END) AS 内容31 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 32 THEN tokuyaku ELSE '' END) AS 内容32 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 33 THEN tokuyaku ELSE '' END) AS 内容33 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 34 THEN tokuyaku ELSE '' END) AS 内容34 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 35 THEN tokuyaku ELSE '' END) AS 内容35 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 36 THEN tokuyaku ELSE '' END) AS 内容36 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 37 THEN tokuyaku ELSE '' END) AS 内容37 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 38 THEN tokuyaku ELSE '' END) AS 内容38 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 39 THEN tokuyaku ELSE '' END) AS 内容39 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 40 THEN tokuyaku ELSE '' END) AS 内容40 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 41 THEN tokuyaku ELSE '' END) AS 内容41 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 42 THEN tokuyaku ELSE '' END) AS 内容42 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 43 THEN tokuyaku ELSE '' END) AS 内容43 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 44 THEN tokuyaku ELSE '' END) AS 内容44 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 45 THEN tokuyaku ELSE '' END) AS 内容45 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 46 THEN tokuyaku ELSE '' END) AS 内容46 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 47 THEN tokuyaku ELSE '' END) AS 内容47 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 48 THEN tokuyaku ELSE '' END) AS 内容48 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 49 THEN tokuyaku ELSE '' END) AS 内容49 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 50 THEN tokuyaku ELSE '' END) AS 内容50 "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no "
                tmp_sql = tmp_sql & " 				,hy_no "
                tmp_sql = tmp_sql & " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no ORDER BY bk_no,hy_no,mei_no) AS [特約行No] "
                tmp_sql = tmp_sql & " 				,tokuyaku "
                tmp_sql = tmp_sql & " 			FROM hy_genjo_tokuyaku "
                tmp_sql = tmp_sql & " 			WHERE brui = 1 "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 	) AS GENJOTOKUYAKU_GENJO "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 			,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 			,3 AS [特約区分] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 1 THEN tokuyaku ELSE '' END) AS 内容1 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 2 THEN tokuyaku ELSE '' END) AS 内容2 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 3 THEN tokuyaku ELSE '' END) AS 内容3 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 4 THEN tokuyaku ELSE '' END) AS 内容4 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 5 THEN tokuyaku ELSE '' END) AS 内容5 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 6 THEN tokuyaku ELSE '' END) AS 内容6 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 7 THEN tokuyaku ELSE '' END) AS 内容7 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 8 THEN tokuyaku ELSE '' END) AS 内容8 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 9 THEN tokuyaku ELSE '' END) AS 内容9 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 10 THEN tokuyaku ELSE '' END) AS 内容10 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 11 THEN tokuyaku ELSE '' END) AS 内容11 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 12 THEN tokuyaku ELSE '' END) AS 内容12 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 13 THEN tokuyaku ELSE '' END) AS 内容13 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 14 THEN tokuyaku ELSE '' END) AS 内容14 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 15 THEN tokuyaku ELSE '' END) AS 内容15 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 16 THEN tokuyaku ELSE '' END) AS 内容16 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 17 THEN tokuyaku ELSE '' END) AS 内容17 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 18 THEN tokuyaku ELSE '' END) AS 内容18 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 19 THEN tokuyaku ELSE '' END) AS 内容19 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 20 THEN tokuyaku ELSE '' END) AS 内容20 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 21 THEN tokuyaku ELSE '' END) AS 内容21 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 22 THEN tokuyaku ELSE '' END) AS 内容22 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 23 THEN tokuyaku ELSE '' END) AS 内容23 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 24 THEN tokuyaku ELSE '' END) AS 内容24 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 25 THEN tokuyaku ELSE '' END) AS 内容25 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 26 THEN tokuyaku ELSE '' END) AS 内容26 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 27 THEN tokuyaku ELSE '' END) AS 内容27 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 28 THEN tokuyaku ELSE '' END) AS 内容28 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 29 THEN tokuyaku ELSE '' END) AS 内容29 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 30 THEN tokuyaku ELSE '' END) AS 内容30 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 31 THEN tokuyaku ELSE '' END) AS 内容31 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 32 THEN tokuyaku ELSE '' END) AS 内容32 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 33 THEN tokuyaku ELSE '' END) AS 内容33 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 34 THEN tokuyaku ELSE '' END) AS 内容34 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 35 THEN tokuyaku ELSE '' END) AS 内容35 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 36 THEN tokuyaku ELSE '' END) AS 内容36 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 37 THEN tokuyaku ELSE '' END) AS 内容37 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 38 THEN tokuyaku ELSE '' END) AS 内容38 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 39 THEN tokuyaku ELSE '' END) AS 内容39 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 40 THEN tokuyaku ELSE '' END) AS 内容40 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 41 THEN tokuyaku ELSE '' END) AS 内容41 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 42 THEN tokuyaku ELSE '' END) AS 内容42 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 43 THEN tokuyaku ELSE '' END) AS 内容43 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 44 THEN tokuyaku ELSE '' END) AS 内容44 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 45 THEN tokuyaku ELSE '' END) AS 内容45 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 46 THEN tokuyaku ELSE '' END) AS 内容46 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 47 THEN tokuyaku ELSE '' END) AS 内容47 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 48 THEN tokuyaku ELSE '' END) AS 内容48 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 49 THEN tokuyaku ELSE '' END) AS 内容49 "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [特約行No] = 50 THEN tokuyaku ELSE '' END) AS 内容50 "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no "
                tmp_sql = tmp_sql & " 				,hy_no "
                tmp_sql = tmp_sql & " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no ORDER BY bk_no,hy_no,mei_no) AS [特約行No] "
                tmp_sql = tmp_sql & " 				,tokuyaku "
                tmp_sql = tmp_sql & " 			FROM hy_genjo_tokuyaku "
                tmp_sql = tmp_sql & " 			WHERE brui = 2 "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 	) AS GENJOTOKUYAKU_NYUKYO "
                tmp_sql = tmp_sql & " ) AS TOTAL "

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

#Region "部屋鍵情報"

    Public Class Hy_kagi_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                '2016.04.06 部屋鍵情報の移行処理修正 -chg sta
                'sortstr = "[物件NO],[部屋NO]"

                ''10のテーブルにそのまま合わせた状態で抽出
                'tmp_sql = tmp_sql & " SELECT "
                'tmp_sql = tmp_sql & " 	 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 	,hy_no AS [部屋NO] "
                'tmp_sql = tmp_sql & " 	,kagimei_no AS [鍵明細No] "
                'tmp_sql = tmp_sql & " 	,kagi_honsu AS [鍵本数] "
                'tmp_sql = tmp_sql & " 	,kagi_no AS [内容] "
                'tmp_sql = tmp_sql & " 	,'' AS [保管場所] "
                'tmp_sql = tmp_sql & " 	,'' AS [業者間での情報共有] "
                'tmp_sql = tmp_sql & " FROM hy_kagi "
                'tmp_sql = tmp_sql & " WHERE NOT (ISNULL(kagi_honsu,0) = 0 AND ISNULL(kagi_no,'') = '') "

                ''鍵タイトル名称を含んだ状態で抽出
                ''tmp_sql = tmp_sql & " SELECT "
                ''tmp_sql = tmp_sql & " 	 bk_no AS [物件NO] "
                ''tmp_sql = tmp_sql & " 	,hy_no AS [部屋NO] "
                ''tmp_sql = tmp_sql & " 	,HYKAGI.kagimei_no AS [鍵明細No] "
                ''tmp_sql = tmp_sql & " 	,biko_titl AS [鍵タイトル] "
                ''tmp_sql = tmp_sql & " 	,kagi_honsu AS [鍵本数] "
                ''tmp_sql = tmp_sql & " 	,kagi_no AS [内容] "
                ''tmp_sql = tmp_sql & " 	,'' AS [保管場所] "
                ''tmp_sql = tmp_sql & " 	,'' AS [業者間での情報共有] "
                ''tmp_sql = tmp_sql & " FROM hy_kagi HYKAGI "
                ''tmp_sql = tmp_sql & " LEFT JOIN "
                ''tmp_sql = tmp_sql & " 	( "
                ''tmp_sql = tmp_sql & " 		SELECT "
                ''tmp_sql = tmp_sql & " 			 biko_no "
                ''tmp_sql = tmp_sql & " 			,biko_titl "
                ''tmp_sql = tmp_sql & " 		FROM m_biko "
                ''tmp_sql = tmp_sql & " 		WHERE biko_kbn = 4 "
                ''tmp_sql = tmp_sql & " 	) AS KAGITITLE "
                ''tmp_sql = tmp_sql & " ON HYKAGI.kagimei_no = KAGITITLE.biko_no "
                ''tmp_sql = tmp_sql & " WHERE ISNULL(biko_titl,'') <> '' "
                ''tmp_sql = tmp_sql & " AND   ISNULL(biko_titl,'') <> '' "

                sortstr = "[物件No],[部屋No],[鍵タイトル名]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件No],[部屋No],[鍵No],[鍵タイトル名],[鍵本数],[備考],[保管場所],[業者間での情報共有]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                '20160531 鍵情報移行処理の修正 -chg sta
                ''tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 bk_no AS [物件No] "
                'tmp_sql = tmp_sql & " 		,hy_no AS [部屋No] "
                'tmp_sql = tmp_sql & " 		/*,kagimei_no*/ "
                'tmp_sql = tmp_sql & " 		,biko_titl AS [鍵タイトル名] "
                'tmp_sql = tmp_sql & " 		,kagi_honsu AS [鍵本数] "
                'tmp_sql = tmp_sql & " 		,kagi_no AS [備考] "
                'tmp_sql = tmp_sql & " 		,'' AS [保管場所] "
                'tmp_sql = tmp_sql & " 		,'' AS [業者間での情報共有] "
                'tmp_sql = tmp_sql & " 	FROM hy_kagi AS KAGI "
                'tmp_sql = tmp_sql & " 	LEFT JOIN "
                'tmp_sql = tmp_sql & " 		( "
                'tmp_sql = tmp_sql & " 			SELECT * FROM m_biko WHERE biko_kbn = 4	 "
                'tmp_sql = tmp_sql & " 		) AS KAGITITLE "
                'tmp_sql = tmp_sql & " 	ON KAGI.kagimei_no = KAGITITLE.biko_no "
                'tmp_sql = tmp_sql & " 	WHERE ISNULL(kagi_no,'') <> '' "
                'tmp_sql = tmp_sql & " ) AS VW "
                ''2016.04.06 部屋鍵情報の移行処理修正 -chg end

                Select Case True
                    Case Njc.Frm.MainFrm.optHyKagi.Checked  '部屋鍵情報から取得する場合
                        'tmp_sql = tmp_sql & " SELECT * FROM "
                        tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                        tmp_sql = tmp_sql & " ( "
                        tmp_sql = tmp_sql & " 	SELECT "
                        tmp_sql = tmp_sql & " 		 bk_no AS [物件No] "
                        tmp_sql = tmp_sql & " 		,hy_no AS [部屋No] "
                        tmp_sql = tmp_sql & " 		/*,kagimei_no*/ "
                        tmp_sql = tmp_sql & "       ,kagimei_no AS [鍵No]	/*20160613 鍵情報の取得処理修正 add*/ "
                        tmp_sql = tmp_sql & " 		,biko_titl AS [鍵タイトル名] "
                        tmp_sql = tmp_sql & " 		,kagi_honsu AS [鍵本数] "
                        tmp_sql = tmp_sql & " 		,kagi_no AS [備考] "
                        tmp_sql = tmp_sql & " 		,'' AS [保管場所] "
                        tmp_sql = tmp_sql & " 		,'' AS [業者間での情報共有] "
                        tmp_sql = tmp_sql & " 	FROM hy_kagi AS KAGI "
                        tmp_sql = tmp_sql & " 	LEFT JOIN "
                        tmp_sql = tmp_sql & " 		( "
                        tmp_sql = tmp_sql & " 			SELECT * FROM m_biko WHERE biko_kbn = 4	 "
                        tmp_sql = tmp_sql & " 		) AS KAGITITLE "
                        tmp_sql = tmp_sql & " 	ON KAGI.kagimei_no = KAGITITLE.biko_no "
                        tmp_sql = tmp_sql & " 	WHERE ISNULL(kagi_no,'') <> '' "
                        tmp_sql = tmp_sql & " ) AS VW "
                    Case Njc.Frm.MainFrm.optKyKagi.Checked  '契約鍵情報から取得する場合
                        'tmp_sql = tmp_sql & " SELECT * FROM "
                        tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                        tmp_sql = tmp_sql & " ( "
                        tmp_sql = tmp_sql & " 	SELECT "
                        tmp_sql = tmp_sql & " 		 bk_no AS [物件No] "
                        tmp_sql = tmp_sql & " 		,hy_no AS [部屋No] "
                        tmp_sql = tmp_sql & " 		/*,kagimei_no*/ "
                        tmp_sql = tmp_sql & "       ,kagimei_no AS [鍵No]	/*20160613 鍵情報の取得処理修正 add*/ "
                        tmp_sql = tmp_sql & " 		,biko_titl AS [鍵タイトル名] "
                        tmp_sql = tmp_sql & " 		,kagi_honsu AS [鍵本数] "
                        tmp_sql = tmp_sql & " 		,kagi_no AS [備考] "
                        tmp_sql = tmp_sql & " 		,'' AS [保管場所] "
                        tmp_sql = tmp_sql & " 		,'' AS [業者間での情報共有] "
                        tmp_sql = tmp_sql & " 	FROM "
                        tmp_sql = tmp_sql & " 	( "
                        tmp_sql = tmp_sql & " 		/*契約Noが最大のデータに紐付く契約鍵情報を取得 sta*/ "
                        tmp_sql = tmp_sql & " 		SELECT * FROM ky_kagi AS KYKAGI "
                        tmp_sql = tmp_sql & " 		WHERE EXISTS "
                        tmp_sql = tmp_sql & " 			( "
                        tmp_sql = tmp_sql & " 				/*契約Noが最大のデータを取得 sta*/ "
                        tmp_sql = tmp_sql & " 				SELECT bk_no,hy_no,ky_no FROM "
                        tmp_sql = tmp_sql & " 				( "
                        tmp_sql = tmp_sql & " 					SELECT * FROM "
                        tmp_sql = tmp_sql & " 					( "
                        tmp_sql = tmp_sql & " 						SELECT "
                        tmp_sql = tmp_sql & " 							 bk_no "
                        tmp_sql = tmp_sql & " 							,hy_no "
                        tmp_sql = tmp_sql & " 							,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no ORDER BY ky_no DESC) AS [抽出用] "
                        tmp_sql = tmp_sql & " 							,ky_no "
                        tmp_sql = tmp_sql & " 						FROM ky_kagi AS KAGI "
                        tmp_sql = tmp_sql & " 					) AS VW "
                        tmp_sql = tmp_sql & " 					WHERE [抽出用] = 1 "
                        tmp_sql = tmp_sql & " 				) AS VW "
                        tmp_sql = tmp_sql & " 				/*契約Noが最大のデータを取得 end*/ "
                        tmp_sql = tmp_sql & " 				WHERE KYKAGI.bk_no = VW.bk_no "
                        tmp_sql = tmp_sql & " 				AND   KYKAGI.hy_no = VW.hy_no "
                        tmp_sql = tmp_sql & " 				AND   KYKAGI.ky_no = VW.ky_no "
                        tmp_sql = tmp_sql & " 			) "
                        tmp_sql = tmp_sql & " 		/*契約Noが最大のデータに紐付く契約鍵情報を取得 end*/ "
                        tmp_sql = tmp_sql & " 	) AS KYKAGITOTAL "
                        tmp_sql = tmp_sql & " 	LEFT JOIN "
                        tmp_sql = tmp_sql & " 		( "
                        tmp_sql = tmp_sql & " 			SELECT * FROM m_biko WHERE biko_kbn = 4	 "
                        tmp_sql = tmp_sql & " 		) AS KAGITITLE "
                        tmp_sql = tmp_sql & " 	ON KYKAGITOTAL.kagimei_no = KAGITITLE.biko_no "
                        tmp_sql = tmp_sql & " 	WHERE ISNULL(kagi_no,'') <> '' "
                        tmp_sql = tmp_sql & " ) AS VW "
                End Select
                '20160531 鍵情報移行処理の修正 -chg end

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

#Region "部屋間取内訳情報"

    Public Class Hy_mst_madriutiwake_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO],[間取り内訳No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[間取り内訳No],[間取り内訳区分],[畳数],[所在階],[既存間取内訳データ]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,'' AS [間取り内訳No] "
                tmp_sql = tmp_sql & " 		,'' AS [間取り内訳区分] "
                tmp_sql = tmp_sql & " 		,'' AS [畳数] "
                tmp_sql = tmp_sql & " 		,'' AS [所在階] "
                tmp_sql = tmp_sql & " 		,madoriu_all AS [既存間取内訳データ] "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
                tmp_sql = tmp_sql & " 	WHERE ISNULL(madoriu_all,'') <> '' "
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

#Region "部屋面積情報"

    Public Class Hy_mst_menseki_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO],[面積タイプ],[面積No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[面積タイプ],[面積No],[面積区分],[区画名],[面積(m2)],[坪数]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,2 AS [面積タイプ] "
                tmp_sql = tmp_sql & " 		,1 AS [面積No] "
                tmp_sql = tmp_sql & " 		,-1 AS [面積区分] "
                tmp_sql = tmp_sql & " 		,men_syosai_title1 AS [区画名] "
                tmp_sql = tmp_sql & " 		,men_syosai_men1 AS [面積(m2)] "
                tmp_sql = tmp_sql & " 		,men_syosai_tubo1 AS [坪数] "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,2 AS [面積タイプ] "
                tmp_sql = tmp_sql & " 		,2 AS [面積No] "
                tmp_sql = tmp_sql & " 		,-1 AS [面積区分] "
                tmp_sql = tmp_sql & " 		,men_syosai_title2 AS [区画名] "
                tmp_sql = tmp_sql & " 		,men_syosai_men2 AS [面積(m2)] "
                tmp_sql = tmp_sql & " 		,men_syosai_tubo2 AS [坪数] "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,2 AS [面積タイプ] "
                tmp_sql = tmp_sql & " 		,3 AS [面積No] "
                tmp_sql = tmp_sql & " 		,-1 AS [面積区分] "
                tmp_sql = tmp_sql & " 		,men_syosai_title3 AS [区画名] "
                tmp_sql = tmp_sql & " 		,men_syosai_men3 AS [面積(m2)] "
                tmp_sql = tmp_sql & " 		,men_syosai_tubo3 AS [坪数] "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,2 AS [面積タイプ] "
                tmp_sql = tmp_sql & " 		,4 AS [面積No] "
                tmp_sql = tmp_sql & " 		,-1 AS [面積区分] "
                tmp_sql = tmp_sql & " 		,men_syosai_title4 AS [区画名] "
                tmp_sql = tmp_sql & " 		,men_syosai_men4 AS [面積(m2)] "
                tmp_sql = tmp_sql & " 		,men_syosai_tubo4 AS [坪数] "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,2 AS [面積タイプ] "
                tmp_sql = tmp_sql & " 		,5 AS [面積No] "
                tmp_sql = tmp_sql & " 		,-1 AS [面積区分] "
                tmp_sql = tmp_sql & " 		,men_syosai_title5 AS [区画名] "
                tmp_sql = tmp_sql & " 		,men_syosai_men5 AS [面積(m2)] "
                tmp_sql = tmp_sql & " 		,men_syosai_tubo5 AS [坪数] "
                tmp_sql = tmp_sql & " 	FROM hy_mst "
                tmp_sql = tmp_sql & " ) AS TOTAL "
                tmp_sql = tmp_sql & " WHERE ISNULL([区画名],'') <> '' "

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

#Region "部屋設備情報"

    Public Class Hy_setubi_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                'sortstr = "[物件NO],[部屋NO],[設備No]"
                sortstr = "[物件NO],[部屋NO]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[設備項目名],[設備内容]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		/*,[設備No]*/ "
                tmp_sql = tmp_sql & " 		,setubi_name AS [設備項目名] "
                tmp_sql = tmp_sql & " 		,setubi1 AS [設備内容] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi1,1 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi2,2 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi3,3 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi4,4 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi5,5 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi6,6 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi7,7 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi8,8 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi9,9 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi10,10 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi11,11 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi12,12 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi13,13 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi14,14 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi15,15 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi16,16 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi17,17 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi18,18 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi19,19 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi20,20 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi21,21 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi22,22 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi23,23 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi24,24 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi25,25 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi26,26 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi27,27 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi28,28 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi29,29 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi30,30 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi31,31 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi32,32 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi33,33 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi34,34 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi35,35 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi36,36 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi37,37 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi38,38 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi39,39 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi40,40 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi41,41 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi42,42 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi43,43 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi44,44 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi45,45 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi46,46 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi47,47 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi48,48 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi49,49 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi50,50 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi51,51 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi52,52 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi53,53 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi54,54 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi55,55 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi56,56 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi57,57 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi58,58 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi59,59 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi60,60 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi61,61 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi62,62 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi63,63 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi64,64 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi65,65 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi66,66 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi67,67 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi68,68 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi69,69 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi70,70 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi71,71 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi72,72 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi73,73 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi74,74 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi75,75 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi76,76 AS 設備No FROM hy_setubi UNION "
                tmp_sql = tmp_sql & " 		SELECT bk_no,hy_no,setubi77,77 AS 設備No FROM hy_setubi "
                tmp_sql = tmp_sql & " 	) AS HYSETUBI "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_setubi AS MS ON HYSETUBI.設備No = MS.setubi_no "
                tmp_sql = tmp_sql & " ) AS SETUBITOTAL "
                tmp_sql = tmp_sql & " WHERE [設備内容] <> '' "

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

#Region "部屋入金項目情報"

    Public Class Hy_kanri_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO],[月区分],[入金項目行No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[月区分],[入金項目名],[入金項目行No],[入金項目区分],[請求額],[税区分],[算出区分],[算出基準入金項目名],[算出ヶ月],[請求先No],[請求月区分],[請求開始月],[請求パターン],[固定公共料金で使用],[請求発生年],[請求発生月],[備考]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		/*20160722 10本体側の月区分変更に伴う修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 2 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 5 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE '' "
                tmp_sql = tmp_sql & " 		 END AS [月区分] "
                tmp_sql = tmp_sql & " 		 */ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 2 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 3 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 5 THEN 2 "
                tmp_sql = tmp_sql & " 			ELSE '' "
                tmp_sql = tmp_sql & " 		 END AS [月区分]		  "
                tmp_sql = tmp_sql & " 		/*20160722 10本体側の月区分変更に伴う修正 chg end*/ "
                tmp_sql = tmp_sql & " 		/*,HYK1.nkin_no AS [入金項目No]*/ "
                tmp_sql = tmp_sql & " 		,MN1.nkin_name AS [入金項目名] "
                tmp_sql = tmp_sql & " 		,naibu_no AS [入金項目行No] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN koteihendo_kbn = 1 AND HYK1.nkin_kbn = 5 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE HYK1.nkin_kbn "
                tmp_sql = tmp_sql & " 		 END AS [入金項目区分] "
                tmp_sql = tmp_sql & " 		/*20160523 敷金の税区分修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*革命10では敷金に消費税の設定ができないため取得方法を変更*/ "
                tmp_sql = tmp_sql & " 		/*敷金は2010で固定で設定する*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,sq_gak AS [請求額] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN sq_zeiumu = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sq_zeiumu = 1 THEN 2 "
                tmp_sql = tmp_sql & " 		 END AS [税区分] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_no = 2010 AND sq_zeiumu = 1 THEN sq_gak * 1.08 "
                tmp_sql = tmp_sql & " 			ELSE sq_gak "
                tmp_sql = tmp_sql & " 		 END AS [請求額] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_no = 2010 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_no <> 2010 AND sq_zeiumu = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_no <> 2010 AND sq_zeiumu = 1 THEN 2 "
                tmp_sql = tmp_sql & " 		 END AS [税区分]	 "
                tmp_sql = tmp_sql & " 		/*20160523 敷金の税区分修正 chg end*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(tukisu,0) = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(tukisu,0) <> 0 THEN 2 "
                tmp_sql = tmp_sql & " 		 END AS [算出区分] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 5 THEN NULL "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(tukisu,0) <> 0 THEN "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT TOP 1 MN2.nkin_name FROM hy_kanri AS HYK2	/*20160823 山川さん指摘対応 内部No重複によるエラー修正 TOP 1を追加*/ "
                tmp_sql = tmp_sql & " 					LEFT JOIN m_nkin AS MN2 ON HYK2.nkin_no = MN2.nkin_no "
                tmp_sql = tmp_sql & " 					WHERE HYK1.bk_no = HYK2.bk_no "
                tmp_sql = tmp_sql & " 					AND   HYK1.hy_no = HYK2.hy_no "
                tmp_sql = tmp_sql & " 					AND   HYK2.nkin_kbn = 1 "
                tmp_sql = tmp_sql & " 					AND   HYK2.naibu_no = 1 "
                tmp_sql = tmp_sql & " 				) "
                tmp_sql = tmp_sql & " 		 END AS [算出基準入金項目名] "
                tmp_sql = tmp_sql & " 		 /*変動費は請求間隔を月数に設定しているのでこの場合は算出ヶ月から除外する*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 5 THEN NULL "
                tmp_sql = tmp_sql & " 			ELSE tukisu "
                tmp_sql = tmp_sql & " 		 END AS [算出ヶ月] "
                tmp_sql = tmp_sql & " 		,'' AS [請求先No] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = -2 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = -1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 0 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 1 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 2 THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [請求月区分] "
                tmp_sql = tmp_sql & " 		,'' AS [請求開始月] "
                tmp_sql = tmp_sql & " 		/*2016.04.26 メインの方へも反映させる修正 -chg sta*/ "
                tmp_sql = tmp_sql & " 		/*変動費で固定に設定されている入金項目は請求間隔から値を設定する (他は1を設定)*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE '' "
                tmp_sql = tmp_sql & " 		 END AS [請求パターン] "
                tmp_sql = tmp_sql & " 		 */ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN koteihendo_kbn = 1 AND HYK1.nkin_kbn = 5 THEN "
                tmp_sql = tmp_sql & " 				CASE WHEN tukisu = 1 THEN 1 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 2 THEN 1 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 3 THEN 2 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 4 THEN 2 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 5 THEN 1 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 6 THEN 2 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 7 THEN 1 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 8 THEN 1 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 9 THEN 1 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 10 THEN 1 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 11 THEN 1 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu = 12 THEN 2 "
                tmp_sql = tmp_sql & " 					 WHEN tukisu >= 13 THEN 1 END "
                tmp_sql = tmp_sql & " 			ELSE 1 "
                tmp_sql = tmp_sql & " 		 END AS [請求パターン]		  "
                tmp_sql = tmp_sql & " 		 /*2016.04.26 メインの方へも反映させる修正 -chg end*/ "
                tmp_sql = tmp_sql & "  "
                tmp_sql = tmp_sql & " 		/*2016.04.26 メインの方へも反映させる修正 -add sta*/  "
                tmp_sql = tmp_sql & " 		/*変動費で固定の入金項目は請求間隔を取得しておく*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN HYK1.nkin_kbn = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE '' "
                tmp_sql = tmp_sql & " 		 END AS [固定公共料金で使用] "
                tmp_sql = tmp_sql & " 		 */		 "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN koteihendo_kbn = 1 AND HYK1.nkin_kbn = 5 THEN tukisu "
                tmp_sql = tmp_sql & " 			ELSE NULL "
                tmp_sql = tmp_sql & " 		 END AS [固定公共料金で使用]		  "
                tmp_sql = tmp_sql & " 		 /*2016.04.26 メインの方へも反映させる修正 -add end*/ 	  "
                tmp_sql = tmp_sql & " 		,NULL AS [請求発生年] "
                tmp_sql = tmp_sql & " 		,NULL AS [請求発生月] "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 	FROM hy_kanri AS HYK1 "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin AS MN1 ON HYK1.nkin_no = MN1.nkin_no "
                tmp_sql = tmp_sql & " 	WHERE koteihendo_kbn = 1 "
                tmp_sql = tmp_sql & " 	AND   sq_gak <> 0 "
                tmp_sql = tmp_sql & " 	AND   HYK1.nkin_kbn IN (1,2,3,5) "
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

#Region "部屋変動費各戸メーター情報"

    Public Class Hy_kanri_hendo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO],[行No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[行No],[メーター分類],[メーター名],[変動費入金項目名],[備考],[変動費請求ルールNo],[変動ルール備考],[請求月区分]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                '20160603 ユーザーデータ検証による修正 -chg sta
                ''tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                'tmp_sql = tmp_sql & " 		,naibu_no AS [行No] "
                'tmp_sql = tmp_sql & " 		,'' AS [メーター分類] "
                'tmp_sql = tmp_sql & " 		,'' AS [メーター名] "
                'tmp_sql = tmp_sql & " 		/*,nkin_no AS [変動費入金項目]*/ "
                'tmp_sql = tmp_sql & " 		,MN.nkin_name AS [変動費入金項目名] "
                'tmp_sql = tmp_sql & " 		,biko AS [備考] "
                'tmp_sql = tmp_sql & " 		,hendo_no AS [変動費請求ルールNo] "
                'tmp_sql = tmp_sql & " 		,'' AS [変動ルール備考] "
                'tmp_sql = tmp_sql & " 		,CASE "
                'tmp_sql = tmp_sql & " 			WHEN sq_mm = -2 THEN 1 "
                'tmp_sql = tmp_sql & " 			WHEN sq_mm = -1 THEN 2 "
                'tmp_sql = tmp_sql & " 			WHEN sq_mm = 0 THEN 3 "
                'tmp_sql = tmp_sql & " 			WHEN sq_mm = 1 THEN 4 "
                'tmp_sql = tmp_sql & " 			WHEN sq_mm = 2 THEN 5 "
                'tmp_sql = tmp_sql & " 		 END AS [請求月区分]	 "
                'tmp_sql = tmp_sql & " 	FROM hy_kanri AS HYNKIN "
                'tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin AS MN ON HYNKIN.nkin_no = MN.nkin_no "
                'tmp_sql = tmp_sql & " 	WHERE koteihendo_kbn = 2 "
                'tmp_sql = tmp_sql & " ) AS VW "

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no ORDER BY naibu_no,TOTAL.nkin_no) AS [行No] "
                tmp_sql = tmp_sql & " 		,'' AS [メーター分類] "
                tmp_sql = tmp_sql & " 		,'' AS [メーター名] "
                tmp_sql = tmp_sql & " 		,nkin_name AS [変動費入金項目名] "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 		,hendo_no AS [変動費請求ルールNo] "
                tmp_sql = tmp_sql & " 		,'' AS [変動ルール備考] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = -2 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = -1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 0 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 1 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 2 THEN 5 "
                tmp_sql = tmp_sql & " 		 END AS [請求月区分] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		/*契約変動費情報を取得 sta*/ "
                tmp_sql = tmp_sql & " 		SELECT * FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no "
                tmp_sql = tmp_sql & " 				,hy_no "
                tmp_sql = tmp_sql & " 				,naibu_no "
                tmp_sql = tmp_sql & " 				,nkin_no "
                tmp_sql = tmp_sql & " 				,biko "
                tmp_sql = tmp_sql & " 				,hendo_no "
                tmp_sql = tmp_sql & " 				,sq_mm "
                tmp_sql = tmp_sql & " 			FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT "
                tmp_sql = tmp_sql & " 					 ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no,nkin_no ORDER BY ky_no DESC,ko_no DESC) AS [抽出用] "
                tmp_sql = tmp_sql & " 					,* "
                tmp_sql = tmp_sql & " 				FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT * FROM ky_sqdata	WHERE ko_no <> 999 AND koteihendo_kbn = 2 "
                tmp_sql = tmp_sql & " 				) AS VW1 "
                tmp_sql = tmp_sql & " 			) AS VW2 "
                tmp_sql = tmp_sql & " 			WHERE [抽出用] = 1 "
                tmp_sql = tmp_sql & " 		) AS KYHENDO "
                tmp_sql = tmp_sql & " 		/*契約変動費情報を取得 end*/ "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		/*契約変動費情報に存在せず部屋情報にのみ存在する部屋変動費情報を取得 sta*/ "
                tmp_sql = tmp_sql & " 		SELECT * FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no "
                tmp_sql = tmp_sql & " 				,hy_no "
                tmp_sql = tmp_sql & " 				,naibu_no "
                tmp_sql = tmp_sql & " 				,nkin_no "
                tmp_sql = tmp_sql & " 				,biko "
                tmp_sql = tmp_sql & " 				,hendo_no "
                tmp_sql = tmp_sql & " 				,sq_mm	 "
                tmp_sql = tmp_sql & " 			FROM hy_kanri AS HYNKIN "
                tmp_sql = tmp_sql & " 			WHERE koteihendo_kbn = 2 "
                tmp_sql = tmp_sql & " 		) AS HYHENDO "
                tmp_sql = tmp_sql & " 		WHERE NOT EXISTS "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 bk_no "
                tmp_sql = tmp_sql & " 						,hy_no "
                tmp_sql = tmp_sql & " 						,naibu_no "
                tmp_sql = tmp_sql & " 						,nkin_no "
                tmp_sql = tmp_sql & " 						,biko "
                tmp_sql = tmp_sql & " 						,hendo_no "
                tmp_sql = tmp_sql & " 						,sq_mm	 "
                tmp_sql = tmp_sql & " 					FROM ky_sqdata AS KYSQ "
                tmp_sql = tmp_sql & " 					WHERE koteihendo_kbn = 2 "
                tmp_sql = tmp_sql & " 				) AS KYHENDO "
                tmp_sql = tmp_sql & " 				WHERE HYHENDO.bk_no = KYHENDO.bk_no "
                tmp_sql = tmp_sql & " 				AND   HYHENDO.hy_no = KYHENDO.hy_no "
                tmp_sql = tmp_sql & " 				AND   HYHENDO.nkin_no = KYHENDO.nkin_no "
                tmp_sql = tmp_sql & " 			) "
                tmp_sql = tmp_sql & " 		/*契約変動費情報に存在せず部屋情報にのみ存在する部屋変動費情報を取得 end*/ "
                tmp_sql = tmp_sql & " 	) AS TOTAL "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_nkin AS MN ON TOTAL.nkin_no = MN.nkin_no "
                tmp_sql = tmp_sql & " ) AS VW "
                '20160603 ユーザーデータ検証による修正 -chg end

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

#Region "部屋修繕維持管理連絡先情報"

    Public Class Hy_szeniji_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO],[修繕及び維持管理No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[修繕及び維持管理No],[修繕及び維持管理の箇所],[対象区分],[業者no],[氏名(商号または名称)(SJIS)],[住所(主たる事務所の所在地)],[連絡先電話番号],[自社no],[貸主No],[所有者No],[氏名(商号または名称)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 [物件NO] "
                tmp_sql = tmp_sql & " 		,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 		,[修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 		,[修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 		,[対象区分] "
                tmp_sql = tmp_sql & " 		,[業者no] "
                tmp_sql = tmp_sql & " 		,[氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 		,[住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 		,[連絡先電話番号] "
                tmp_sql = tmp_sql & " 		,[自社no] "
                tmp_sql = tmp_sql & " 		,[貸主No] "
                tmp_sql = tmp_sql & " 		,[所有者No] "
                tmp_sql = tmp_sql & " 		,[氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT TOP 1 "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 			,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 			,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 			,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 			,'' AS [業者no] "
                tmp_sql = tmp_sql & " 			,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 			,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 			,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 			,[自社・支店No] AS [自社no] "
                tmp_sql = tmp_sql & " 			,'' AS [貸主No] "
                tmp_sql = tmp_sql & " 			,'' AS [所有者No] "
                tmp_sql = tmp_sql & " 			,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT BKREN.*,JISYATOTAL.[自社・支店No] FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 				WHERE copy_kbn = 1 "
                tmp_sql = tmp_sql & " 				AND   kbn = 2 "
                tmp_sql = tmp_sql & " 			) AS BKREN "
                tmp_sql = tmp_sql & " 			LEFT JOIN "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 jisya_no AS [自社・支店No] "
                tmp_sql = tmp_sql & " 						,kaisya_name AS [自社・支店名] "
                tmp_sql = tmp_sql & " 						,'' AS [自社・支店名（SJIS）] "
                tmp_sql = tmp_sql & " 						,kaisya_kana AS [自社・支店カナ] "
                tmp_sql = tmp_sql & " 						,'' AS [自社・支店名２] "
                tmp_sql = tmp_sql & " 						,'' AS [自社・支店名２（SJIS）] "
                tmp_sql = tmp_sql & " 						,'' AS [自社・支店名３] "
                tmp_sql = tmp_sql & " 						,'' AS [自社・支店名３（SJIS）] "
                tmp_sql = tmp_sql & " 						,yakusyoku AS [代表者役職] "
                tmp_sql = tmp_sql & " 						,daihyo AS [代表者名] "
                tmp_sql = tmp_sql & " 						,kaisya_post AS [郵便番号] "
                tmp_sql = tmp_sql & " 						,kaisya_add1 AS [住所1] "
                tmp_sql = tmp_sql & " 						,kaisya_add2 AS [住所2] "
                tmp_sql = tmp_sql & " 						,kaisya_tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 						,kaisya_tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 						,'' AS [携帯TEL1] "
                tmp_sql = tmp_sql & " 						,'' AS [携帯TEL2] "
                tmp_sql = tmp_sql & " 						,kaisya_fax AS [FAX] "
                tmp_sql = tmp_sql & " 						,kaisya_email AS [メールアドレス] "
                tmp_sql = tmp_sql & " 						,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 						,kaisya_url AS [Webアドレス] "
                tmp_sql = tmp_sql & " 						,'' AS [備考（基本情報）] "
                tmp_sql = tmp_sql & " 						,menkyo_no AS [免許証番号] "
                tmp_sql = tmp_sql & " 						,menkyo_ymd AS [免許年月日] "
                tmp_sql = tmp_sql & " 						,takuken_no AS [取引主任者登録番号（代表）] "
                tmp_sql = tmp_sql & " 						,takuken_name AS [取引主任者名（代表）] "
                tmp_sql = tmp_sql & " 						,'' AS [全国賃貸不動産管理業協会会員番号] "
                tmp_sql = tmp_sql & " 						,'' AS [賃貸不動産経営管理士登録番号] "
                tmp_sql = tmp_sql & " 						,'' AS [賃貸不動産経営管理士名] "
                tmp_sql = tmp_sql & " 						,'' AS [賃貸住宅管理業者登録番号] "
                tmp_sql = tmp_sql & " 						,dantai1 AS [加入団体１] "
                tmp_sql = tmp_sql & " 						,dantai2 AS [加入団体２] "
                tmp_sql = tmp_sql & " 						,dantai3 AS [加入団体３] "
                tmp_sql = tmp_sql & " 						,dantai4 AS [加入団体４] "
                tmp_sql = tmp_sql & " 						,dantai5 AS [加入団体５] "
                tmp_sql = tmp_sql & " 						,dantai6 AS [加入団体６] "
                tmp_sql = tmp_sql & " 						,dantai7 AS [加入団体７] "
                tmp_sql = tmp_sql & " 						,dantai8 AS [加入団体８] "
                tmp_sql = tmp_sql & " 						,dantai9 AS [加入団体９] "
                tmp_sql = tmp_sql & " 						,dantai10 AS [加入団体１０] "
                tmp_sql = tmp_sql & " 						,'' AS [代表者名（SJIS）] "
                tmp_sql = tmp_sql & " 						,'' AS [取引主任者名（代表）（SJIS）] "
                tmp_sql = tmp_sql & " 						,'' AS [賃貸不動産経営管理士名（SJIS）] "
                tmp_sql = tmp_sql & " 					FROM "
                tmp_sql = tmp_sql & " 					( "
                tmp_sql = tmp_sql & " 						SELECT * FROM "
                tmp_sql = tmp_sql & " 						( "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								 [jisya_no] "
                tmp_sql = tmp_sql & " 								,[kaisya_name] "
                tmp_sql = tmp_sql & " 								,[kaisya_kana] "
                tmp_sql = tmp_sql & " 								,[yakusyoku] "
                tmp_sql = tmp_sql & " 								,[daihyo] "
                tmp_sql = tmp_sql & " 								,[kaisya_post] "
                tmp_sql = tmp_sql & " 								,[kaisya_add1] "
                tmp_sql = tmp_sql & " 								,[kaisya_add2] "
                tmp_sql = tmp_sql & " 								,[kaisya_tel1] "
                tmp_sql = tmp_sql & " 								,[kaisya_tel2] "
                tmp_sql = tmp_sql & " 								,[kaisya_fax] "
                tmp_sql = tmp_sql & " 								,[kaisya_email] "
                tmp_sql = tmp_sql & " 								,[kaisya_url] "
                tmp_sql = tmp_sql & " 								,[menkyo_no] "
                tmp_sql = tmp_sql & " 								,[menkyo_ymd] "
                tmp_sql = tmp_sql & " 								,[takuken_no] "
                tmp_sql = tmp_sql & " 								,[takuken_name] "
                tmp_sql = tmp_sql & " 								,[dantai1] "
                tmp_sql = tmp_sql & " 								,[dantai2] "
                tmp_sql = tmp_sql & " 								,[dantai3] "
                tmp_sql = tmp_sql & " 								,[dantai4] "
                tmp_sql = tmp_sql & " 								,[dantai5] "
                tmp_sql = tmp_sql & " 								,[dantai6] "
                tmp_sql = tmp_sql & " 								,[dantai7] "
                tmp_sql = tmp_sql & " 								,[dantai8] "
                tmp_sql = tmp_sql & " 								,[dantai9] "
                tmp_sql = tmp_sql & " 								,[dantai10] "
                tmp_sql = tmp_sql & " 							FROM m_jisya "
                tmp_sql = tmp_sql & " 						) AS JISYA "
                tmp_sql = tmp_sql & " 						UNION "
                tmp_sql = tmp_sql & " 						SELECT * FROM "
                tmp_sql = tmp_sql & " 						( "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								 ROW_NUMBER()OVER(ORDER BY [siten_no]) + "
                tmp_sql = tmp_sql & " 									( "
                tmp_sql = tmp_sql & " 										SELECT jisya_no FROM "
                tmp_sql = tmp_sql & " 										( "
                tmp_sql = tmp_sql & " 											SELECT "
                tmp_sql = tmp_sql & " 												ROW_NUMBER()OVER(ORDER BY jisya_no DESC) AS 抽出用 "
                tmp_sql = tmp_sql & " 												,* "
                tmp_sql = tmp_sql & " 											FROM m_jisya "
                tmp_sql = tmp_sql & " 										) AS VW "
                tmp_sql = tmp_sql & " 										WHERE 抽出用 = 1 "
                tmp_sql = tmp_sql & " 									) AS [移行用自社№] "
                tmp_sql = tmp_sql & " 								 ,[siten_name] "
                tmp_sql = tmp_sql & " 								,'' AS [カナ] "
                tmp_sql = tmp_sql & " 								,'' AS [代表者役職] "
                tmp_sql = tmp_sql & " 								,'' AS [代表者名称] "
                tmp_sql = tmp_sql & " 								,[post] "
                tmp_sql = tmp_sql & " 								,[add1] "
                tmp_sql = tmp_sql & " 								,[add2] "
                tmp_sql = tmp_sql & " 								,[tel1] "
                tmp_sql = tmp_sql & " 								,[tel2] "
                tmp_sql = tmp_sql & " 								,[fax] "
                tmp_sql = tmp_sql & " 								,[email] "
                tmp_sql = tmp_sql & " 								,'' AS [url] "
                tmp_sql = tmp_sql & " 								,'' AS [免許証番号] "
                tmp_sql = tmp_sql & " 								,'' AS [免許証年月日] "
                tmp_sql = tmp_sql & " 								,'' AS [宅建取引士登録番号] "
                tmp_sql = tmp_sql & " 								,'' AS [宅建取引士氏名] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体1] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体2] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体3] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体4] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体5] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体6] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体7] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体8] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体9] "
                tmp_sql = tmp_sql & " 								,'' AS [加入団体10] "
                tmp_sql = tmp_sql & " 							FROM m_siten "
                tmp_sql = tmp_sql & " 						) AS SITEN			 "
                tmp_sql = tmp_sql & " 					) AS JISYAMST "
                tmp_sql = tmp_sql & " 				) AS JISYATOTAL "
                tmp_sql = tmp_sql & " 			ON BKREN.[ren_name] = JISYATOTAL.[自社・支店名] "
                tmp_sql = tmp_sql & " 		) AS BKSZENJISYA "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT TOP 1 "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 			,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 			,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 			,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 			,'' AS [業者no] "
                tmp_sql = tmp_sql & " 			,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 			,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 			,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 			,'' AS [自社no] "
                tmp_sql = tmp_sql & " 			,so_no AS [貸主No] "
                tmp_sql = tmp_sql & " 			,'' AS [所有者No] "
                tmp_sql = tmp_sql & " 			,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT VW.*,KASIMST.so_no FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 				WHERE copy_kbn = 2 "
                tmp_sql = tmp_sql & " 				AND   kbn = 2 "
                tmp_sql = tmp_sql & " 			) AS VW "
                tmp_sql = tmp_sql & " 			LEFT JOIN m_yanu_so AS KASIMST "
                tmp_sql = tmp_sql & " 			ON VW.[ren_name] = KASIMST.so_name "
                tmp_sql = tmp_sql & " 		) AS BKSZENKASI "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT TOP 1 "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 			,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 			,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 			,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 			,'' AS [業者no] "
                tmp_sql = tmp_sql & " 			,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 			,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 			,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 			,'' AS [自社no] "
                tmp_sql = tmp_sql & " 			,'' AS [貸主No] "
                tmp_sql = tmp_sql & " 			,so_no AS [所有者No] "
                tmp_sql = tmp_sql & " 			,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT VW.*,KASIMST.so_no FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 				WHERE copy_kbn = 3 "
                tmp_sql = tmp_sql & " 				AND   kbn = 2 "
                tmp_sql = tmp_sql & " 			) AS VW "
                tmp_sql = tmp_sql & " 			LEFT JOIN m_yanu_so AS KASIMST "
                tmp_sql = tmp_sql & " 			ON VW.[ren_name] = KASIMST.so_name "
                tmp_sql = tmp_sql & " 		) AS BKSZENSYO "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT TOP 1 "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 			,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 			,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 			,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 			,gy_no AS [業者no] "
                tmp_sql = tmp_sql & " 			,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 			,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 			,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 			,'' AS [自社no] "
                tmp_sql = tmp_sql & " 			,'' AS [貸主No] "
                tmp_sql = tmp_sql & " 			,'' AS [所有者No] "
                tmp_sql = tmp_sql & " 			,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT VW.*,GYMST.gy_no FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 				WHERE copy_kbn = 4 "
                tmp_sql = tmp_sql & " 				AND   kbn = 2 "
                tmp_sql = tmp_sql & " 			) AS VW "
                tmp_sql = tmp_sql & " 			LEFT JOIN m_gy AS GYMST "
                tmp_sql = tmp_sql & " 			ON VW.[ren_name] = GYMST.gy_name "
                tmp_sql = tmp_sql & " 		) AS BKSZENKASI "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT TOP 1 "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 			,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 			,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 			,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 			,gy_no AS [業者no] "
                tmp_sql = tmp_sql & " 			,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 			,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 			,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 			,'' AS [自社no] "
                tmp_sql = tmp_sql & " 			,'' AS [貸主No] "
                tmp_sql = tmp_sql & " 			,'' AS [所有者No] "
                tmp_sql = tmp_sql & " 			,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT VW.*,GYSYUZENMST.gy_no FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 				WHERE copy_kbn = 5 "
                tmp_sql = tmp_sql & " 				AND   kbn = 2 "
                tmp_sql = tmp_sql & " 			) AS VW "
                tmp_sql = tmp_sql & " 			LEFT JOIN m_gysyuzen AS GYSYUZENMST "
                tmp_sql = tmp_sql & " 			ON VW.[ren_name] = GYSYUZENMST.gy_name "
                tmp_sql = tmp_sql & " 		) AS BKSZENKASI "
                tmp_sql = tmp_sql & " 	) AS SZENTOTAL "
                tmp_sql = tmp_sql & " 	LEFT JOIN hy_mst AS HY "
                tmp_sql = tmp_sql & " 	ON SZENTOTAL.[物件NO] = HY.bk_no "
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

#Region "部屋メモ情報(備考30件 + 重要事項20件の内訳で移行する)"

    Public Class Hy_biko_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[部屋NO]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[部屋NO],[メモ1],[メモ2],[メモ3],[メモ4],[メモ5],[メモ6],[メモ7],[メモ8],[メモ9],[メモ10],[メモ11],[メモ12],[メモ13],[メモ14],[メモ15],[メモ16],[メモ17],[メモ18],[メモ19],[メモ20],[メモ21],[メモ22],[メモ23],[メモ24],[メモ25],[メモ26],[メモ27],[メモ28],[メモ29],[メモ30],[メモ31],[メモ32],[メモ33],[メモ34],[メモ35],[メモ36],[メモ37],[メモ38],[メモ39],[メモ40],[メモ41],[メモ42],[メモ43],[メモ44],[メモ45],[メモ46],[メモ47],[メモ48],[メモ49],[メモ50]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 BIKO.* "
                tmp_sql = tmp_sql & " 		,[メモ31],[メモ32],[メモ33],[メモ34],[メモ35],[メモ36],[メモ37],[メモ38],[メモ39],[メモ40] "
                tmp_sql = tmp_sql & " 		,[メモ41],[メモ42],[メモ43],[メモ44],[メモ45],[メモ46],[メモ47],[メモ48],[メモ49],[メモ50] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 [物件NO] "
                tmp_sql = tmp_sql & " 			,[部屋NO] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 1 THEN [メモ] ELSE '' END) AS [メモ1] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 2 THEN [メモ] ELSE '' END) AS [メモ2] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 3 THEN [メモ] ELSE '' END) AS [メモ3] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 4 THEN [メモ] ELSE '' END) AS [メモ4] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 5 THEN [メモ] ELSE '' END) AS [メモ5] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 6 THEN [メモ] ELSE '' END) AS [メモ6] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 7 THEN [メモ] ELSE '' END) AS [メモ7] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 8 THEN [メモ] ELSE '' END) AS [メモ8] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 9 THEN [メモ] ELSE '' END) AS [メモ9] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 10 THEN [メモ] ELSE '' END) AS [メモ10] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 11 THEN [メモ] ELSE '' END) AS [メモ11] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 12 THEN [メモ] ELSE '' END) AS [メモ12] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 13 THEN [メモ] ELSE '' END) AS [メモ13] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 14 THEN [メモ] ELSE '' END) AS [メモ14] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 15 THEN [メモ] ELSE '' END) AS [メモ15] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 16 THEN [メモ] ELSE '' END) AS [メモ16] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 17 THEN [メモ] ELSE '' END) AS [メモ17] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 18 THEN [メモ] ELSE '' END) AS [メモ18] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 19 THEN [メモ] ELSE '' END) AS [メモ19] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 20 THEN [メモ] ELSE '' END) AS [メモ20] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 21 THEN [メモ] ELSE '' END) AS [メモ21] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 22 THEN [メモ] ELSE '' END) AS [メモ22] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 23 THEN [メモ] ELSE '' END) AS [メモ23] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 24 THEN [メモ] ELSE '' END) AS [メモ24] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 25 THEN [メモ] ELSE '' END) AS [メモ25] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 26 THEN [メモ] ELSE '' END) AS [メモ26] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 27 THEN [メモ] ELSE '' END) AS [メモ27] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 28 THEN [メモ] ELSE '' END) AS [メモ28] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 29 THEN [メモ] ELSE '' END) AS [メモ29] "
                tmp_sql = tmp_sql & " 			,MAX(CASE WHEN [行No] = 30 THEN [メモ] ELSE '' END) AS [メモ30] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 				,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 				,ROW_NUMBER()OVER(PARTITION BY bk_no,hy_no ORDER BY bk_no,hy_no,biko_no) AS [行No] "
                tmp_sql = tmp_sql & " 				,biko AS [メモ] "
                tmp_sql = tmp_sql & " 			FROM hy_biko "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		GROUP BY [物件NO],[部屋NO] "
                tmp_sql = tmp_sql & " 	) AS BIKO "
                tmp_sql = tmp_sql & " 	LEFT JOIN "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 [物件NO] "
                tmp_sql = tmp_sql & " 				,[部屋NO] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 1 THEN [重要事項] ELSE '' END) AS [メモ31] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 2 THEN [重要事項] ELSE '' END) AS [メモ32] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 3 THEN [重要事項] ELSE '' END) AS [メモ33] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 4 THEN [重要事項] ELSE '' END) AS [メモ34] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 5 THEN [重要事項] ELSE '' END) AS [メモ35] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 6 THEN [重要事項] ELSE '' END) AS [メモ36] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 7 THEN [重要事項] ELSE '' END) AS [メモ37] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 8 THEN [重要事項] ELSE '' END) AS [メモ38] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 9 THEN [重要事項] ELSE '' END) AS [メモ39] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 10 THEN [重要事項] ELSE '' END) AS [メモ40] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 11 THEN [重要事項] ELSE '' END) AS [メモ41] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 12 THEN [重要事項] ELSE '' END) AS [メモ42] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 13 THEN [重要事項] ELSE '' END) AS [メモ43] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 14 THEN [重要事項] ELSE '' END) AS [メモ44] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 15 THEN [重要事項] ELSE '' END) AS [メモ45] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 16 THEN [重要事項] ELSE '' END) AS [メモ46] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 17 THEN [重要事項] ELSE '' END) AS [メモ47] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 18 THEN [重要事項] ELSE '' END) AS [メモ48] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 19 THEN [重要事項] ELSE '' END) AS [メモ49] "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN [行No] = 20 THEN [重要事項] ELSE '' END) AS [メモ50] "
                tmp_sql = tmp_sql & " 			FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT "
                tmp_sql = tmp_sql & " 					 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 					,hy_no AS [部屋NO] "
                tmp_sql = tmp_sql & " 					,HYJ.jyuyo_no AS [行No] "
                tmp_sql = tmp_sql & " 					,MJ.jyuyo_name + ' ： ' + jyuyo_lstname AS [重要事項] "
                tmp_sql = tmp_sql & " 				FROM hy_jyuyo AS HYJ "
                tmp_sql = tmp_sql & " 				LEFT JOIN "
                tmp_sql = tmp_sql & " 					( "
                tmp_sql = tmp_sql & " 						SELECT * FROM m_jyuyo "
                tmp_sql = tmp_sql & " 						WHERE jyuyo_kbn = 2 "
                tmp_sql = tmp_sql & " 					) AS MJ "
                tmp_sql = tmp_sql & " 				ON HYJ.jyuyo_no = MJ.jyuyo_no "
                tmp_sql = tmp_sql & " 			) AS TOTAL "
                tmp_sql = tmp_sql & " 			GROUP BY [物件NO],[部屋NO] "
                tmp_sql = tmp_sql & " 		) AS JYUYO "
                tmp_sql = tmp_sql & " 	ON  BIKO.[物件NO] = JYUYO.[物件NO] "
                tmp_sql = tmp_sql & " 	AND BIKO.[部屋NO] = JYUYO.[部屋NO] "
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






    'ここから汎用ツールのみ

#Region "部屋共通セールスポイント情報(汎用ツールのみ )"

    Public Class Hy_commonsalespoint_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = ""

                tmp_sql = tmp_sql & "  "

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

#Region "部屋参照ファイル情報(汎用ツールのみ )"

    Public Class Hy_relfile_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = ""

                tmp_sql = tmp_sql & "  "

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

#Region "部屋原状回復目安単価情報(汎用ツールのみ )"

    Public Class Hy_szen_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = ""

                tmp_sql = tmp_sql & "  "

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

#Region "部屋権利情報(汎用ツールのみ )"

    Public Class Hy_kenri_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = ""

                tmp_sql = tmp_sql & "  "

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

#Region "部屋契約解約確認事項情報(汎用ツールのみ )"

    '保留

    Public Class Hy_confirm_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = ""

                tmp_sql = tmp_sql & "  "

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
