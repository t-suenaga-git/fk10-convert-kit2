Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "物件基本情報"

    Public Class Bk_mst_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[物件名],[物件カナ],[建物識別コード],[郵便番号],[都道府県コード],[市区町村コード],[町地域],[丁番地],[街区番号地番],[その他],[物件名(SJIS)],[物件外部No]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,bk_name AS [物件名] "
                tmp_sql = tmp_sql & " 		,bk_kana AS [物件カナ] "
                tmp_sql = tmp_sql & " 		,tatemono_sikibetu AS [建物識別コード] "
                tmp_sql = tmp_sql & " 		,post AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,ken_no AS [都道府県コード] "
                tmp_sql = tmp_sql & " 		,si_no AS [市区町村コード] "
                tmp_sql = tmp_sql & " 		,add_cyo AS [町地域] "
                tmp_sql = tmp_sql & " 		,add_banti AS [丁番地] "
                tmp_sql = tmp_sql & " 		,'' AS [街区番号地番] "
                tmp_sql = tmp_sql & " 		,add_etc AS [その他] "
                tmp_sql = tmp_sql & " 		,'' AS [物件名(SJIS)] "
                tmp_sql = tmp_sql & " 		,'' AS [物件外部No] "
                tmp_sql = tmp_sql & " 	FROM bk_mst "
                tmp_sql = tmp_sql & " ) AS BK "

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

#Region "物件詳細情報"

    Public Class Bk_mst_syosai_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[物件分類名],[緯度],[経度],[緯度(日本測地系)],[経度(日本測地系)],[電鉄物件フラグ],[竣工日],[地上階建て],[地下階],[エレベータフラグ],[エレベータ数],[建物構造(基本)-構造],[建物構造(基本)-屋根構造],[情報元業者NO],[備考(基本情報)],[総戸数],[物件面積-延床面積],[物件面積-延床面積坪数],[物件面積-敷地面積],[物件面積-敷地面積坪数],[物件面積-駐車面積],[物件面積-駐車面積坪数],[物件面積-延床面積(登記)],[物件面積-延床面積坪数(登記)],[物件面積-敷地面積(公簿)],[物件面積-敷地面積坪数(公簿)],[入金口座-請求月初期値],[契約書用入金締切日],[家賃持参先],[入金口座-家賃入金口座No],[入金口座-契約金用入金口座の有無],[入金口座-契約金用入金口座No],[耐火構造区分],[施工会社No],[保守業者No],[管理形態-管理方式],[管理形態-管理業者名],[管理形態-管理業者担当者],[管理形態-管理業者電話番号],[管理形態-管理業者業務形態],[管理形態-管理人電話番号],[支店No],[その他交通１-その他交通],[その他交通１-距離],[その他交通２-その他交通],[その他交通２-距離],[ライフライン-電気-公共機関No],[ライフライン-上水-公共機関No],[ライフライン-ガス-公共機関No],[ライフライン-排水-公共機関No],[ライフライン-灯油-公共機関No],[ライフライン-その他１-公共機関No],[ライフライン-その他２-公共機関No],[ライフライン-その他３-公共機関No],[検針業務の有無],[検針登録の並び順],[親子メーター変動費の使用有無],[駐車場-付随駐車場有無],[駐車場-自動車台数],[駐車場-バイク台数],[駐車場-自転車台数],[駐車場-自転車利用自由フラグ],[所有者-一棟・所有区分],[石綿使用調査-調査の有無],[石綿使用調査-調査結果の問合せ先-所有者],[石綿使用調査-調査結果の問合せ先-管理組合],[石綿使用調査-調査結果の問合せ先-管理業者],[石綿使用調査-調査結果の問合せ先-施工業者],[石綿使用調査-調査年月日],[石綿使用調査-実施機関],[石綿使用調査-調査範囲],[石綿使用調査-使用有無],[石綿使用調査-使用箇所],[石綿使用調査-備考],[耐震診断-診断有無],[耐震診断-診断記録の問合せ先-所有者],[耐震診断-診断記録の問合せ先-管理組合],[耐震診断-診断記録の問合せ先-管理業者],[耐震診断-耐震基準適合証明書の写し],[耐震診断-住宅性能評価所の写し],[耐震診断-耐震診断結果の写し],[耐震診断-備考],[法令-土砂災害警戒地域内外],[法令-法令分類],[法令-法令内容],[敷地利用-敷地利用種類],[敷地利用-契約期間開始],[敷地利用-契約期間終了],[敷地利用-備考],[入居率一覧対象開始日],[部屋毎に業者が異なる場合フラグ],[ゴミ出しに関する補足情報],[登記情報の日付],[所有権にかかる権利有無],[所有権にかかる権利の種類],[所有権以外の権利有無],[物件面積-建築面積],[物件面積-建築面積坪数],[物件面積-建築面積(登記)],[物件面積-建築面積坪数(登記)],[自社担当者No],[管理形態-管理業者FAX],[エレベータフラグ(家)],[法令-土砂災害特別警戒区域内外],[法令-造成宅地防災区域内外],[法令-津波災害警戒区域内外],[法令-土砂災害警戒地域備考],[法令-土砂災害特別警戒区域備考],[法令-造成宅地防災区域備考],[法令-津波災害警戒区域備考],[管理形態-管理人名],[その他建物構造],[角地フラグ],[都市計画・用途地域1],[都市計画・用途地域2],[管理人名SJIS],[管理業者担当者名SJIS],[管理業者名SJS],[駐車場空き台数],[バイク置き場空き台数],[駐輪場空き台数],[小学校区],[小学校距離],[中学校区],[中学校距離],[エリア],[非常用エレベーターフラグ]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 BK.bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		/*,brui_no AS [物件分類]*/ "
                tmp_sql = tmp_sql & " 		,BKRUI.brui_name AS [物件分類名] "
                tmp_sql = tmp_sql & " 		,ido AS [緯度] "
                tmp_sql = tmp_sql & " 		,keido AS [経度] "
                tmp_sql = tmp_sql & " 		,j_ido AS [緯度(日本測地系)] "
                tmp_sql = tmp_sql & " 		,j_keido AS [経度(日本測地系)] "
                tmp_sql = tmp_sql & " 		,0 AS [電鉄物件フラグ] "
                tmp_sql = tmp_sql & " 		,syunko_ym AS [竣工日] "
                tmp_sql = tmp_sql & " 		,kai AS [地上階建て] "
                tmp_sql = tmp_sql & " 		,tika AS [地下階] "
                tmp_sql = tmp_sql & " 		,0 AS [エレベータフラグ] "
                tmp_sql = tmp_sql & " 		,'' AS [エレベータ数] "
                tmp_sql = tmp_sql & " 		/*20160829 革命10バージョンアップに伴う修正 del*/ "
                tmp_sql = tmp_sql & " 		/*,0 AS [屋上フラグ]*/ "
                tmp_sql = tmp_sql & " 		,kozo AS [建物構造(基本)-構造] "
                tmp_sql = tmp_sql & " 		,yane AS [建物構造(基本)-屋根構造] "
                tmp_sql = tmp_sql & " 		,gy_no AS [情報元業者NO] "
                tmp_sql = tmp_sql & " 		,'' AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		,kosuu AS [総戸数] "
                tmp_sql = tmp_sql & " 		,t_men AS [物件面積-延床面積] "
                tmp_sql = tmp_sql & " 		,t_tubo AS [物件面積-延床面積坪数] "
                tmp_sql = tmp_sql & " 		,s_men AS [物件面積-敷地面積] "
                tmp_sql = tmp_sql & " 		,s_tubo AS [物件面積-敷地面積坪数] "
                tmp_sql = tmp_sql & " 		,c_men AS [物件面積-駐車面積] "
                tmp_sql = tmp_sql & " 		,c_tubo AS [物件面積-駐車面積坪数] "
                tmp_sql = tmp_sql & " 		,'' AS [物件面積-延床面積(登記)] "
                tmp_sql = tmp_sql & " 		,'' AS [物件面積-延床面積坪数(登記)] "
                tmp_sql = tmp_sql & " 		,'' AS [物件面積-敷地面積(公簿)] "
                tmp_sql = tmp_sql & " 		,'' AS [物件面積-敷地面積坪数(公簿)] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = -2 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = -1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 0 THEN 3 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 1 THEN 4 "
                tmp_sql = tmp_sql & " 			WHEN sq_mm = 2 THEN 5 "
                tmp_sql = tmp_sql & " 			ELSE 2	 "
                tmp_sql = tmp_sql & " 		 END AS [入金口座-請求月初期値] "
                tmp_sql = tmp_sql & " 		,sh_kigen AS [契約書用入金締切日] "
                tmp_sql = tmp_sql & " 		,jisansaki AS [家賃持参先] "
                tmp_sql = tmp_sql & " 		,fkom_no AS [入金口座-家賃入金口座No] "
                tmp_sql = tmp_sql & " 		,'' AS [入金口座-契約金用入金口座の有無] "
                tmp_sql = tmp_sql & " 		,'' AS [入金口座-契約金用入金口座No] "
                tmp_sql = tmp_sql & " 		,-1 AS [耐火構造区分] "
                tmp_sql = tmp_sql & " 		,hosyu_kaisya_no AS [施工会社No] "
                tmp_sql = tmp_sql & " 		,seko_kaisya_no AS [保守業者No] "
                tmp_sql = tmp_sql & " 		,'' AS [管理形態-管理方式] "
                tmp_sql = tmp_sql & " 		,kn_gyosya AS [管理形態-管理業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [管理形態-管理業者担当者] "
                tmp_sql = tmp_sql & " 		,kn_tel1 AS [管理形態-管理業者電話番号] "
                tmp_sql = tmp_sql & " 		,kn_jikan AS [管理形態-管理業者業務形態] "
                tmp_sql = tmp_sql & " 		,kn_hytel AS [管理形態-管理人電話番号] "
                tmp_sql = tmp_sql & " 		/*2016.04.06 支店Noの取得処理を修正 -chg sta*/ "
                tmp_sql = tmp_sql & " 		/*,siten_no AS [支店No]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(BK.siten_no,0) = 0 THEN NULL "
                tmp_sql = tmp_sql & " 			ELSE [移行用自社No] "
                tmp_sql = tmp_sql & " 		 END AS [支店No] "
                tmp_sql = tmp_sql & " 		/*2016.04.06 支店Noの取得処理を修正 -chg end*/ "
                tmp_sql = tmp_sql & " 		,kotsu1 AS [その他交通１-その他交通] "
                tmp_sql = tmp_sql & " 		,'' AS [その他交通１-距離] "
                tmp_sql = tmp_sql & " 		,kotsu2 AS [その他交通２-その他交通] "
                tmp_sql = tmp_sql & " 		,'' AS [その他交通２-距離] "
                tmp_sql = tmp_sql & " 		,kokyo_no1 AS [ライフライン-電気-公共機関No] "
                tmp_sql = tmp_sql & " 		,kokyo_no2 AS [ライフライン-上水-公共機関No] "
                tmp_sql = tmp_sql & " 		,kokyo_no3 AS [ライフライン-ガス-公共機関No] "
                tmp_sql = tmp_sql & " 		,'' AS [ライフライン-排水-公共機関No] "
                tmp_sql = tmp_sql & " 		,'' AS [ライフライン-灯油-公共機関No] "
                tmp_sql = tmp_sql & " 		,kokyo_no4 AS [ライフライン-その他１-公共機関No] "
                tmp_sql = tmp_sql & " 		,kokyo_no5 AS [ライフライン-その他２-公共機関No] "
                tmp_sql = tmp_sql & " 		,kokyo_no6 AS [ライフライン-その他３-公共機関No] "
                tmp_sql = tmp_sql & " 		,0 AS [検針業務の有無] "
                tmp_sql = tmp_sql & " 		,1 AS [検針登録の並び順] "
                tmp_sql = tmp_sql & " 		,0 AS [親子メーター変動費の使用有無] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場-付随駐車場有無] "
                tmp_sql = tmp_sql & " 		,car_daisu AS [駐車場-自動車台数] "
                tmp_sql = tmp_sql & " 		,'' AS [駐車場-バイク台数] "
                tmp_sql = tmp_sql & " 		,rin_daisu AS [駐車場-自転車台数] "
                tmp_sql = tmp_sql & " 		,0 AS [駐車場-自転車利用自由フラグ] "
                tmp_sql = tmp_sql & " 		,1 AS [所有者-一棟・所有区分]			/*V7は一棟固定*/ "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 物件詳細情報 del*/ "
                tmp_sql = tmp_sql & " 		/*,'' AS [該当月/送金月ベース決定]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN iw_kiroku_umu = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN iw_kiroku_umu = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN iw_kiroku_umu = 2 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN iw_kiroku_umu = 3 THEN 3 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 		 END AS [石綿使用調査-調査の有無] "
                tmp_sql = tmp_sql & " 		,iw_syokai1_umu AS [石綿使用調査-調査結果の問合せ先-所有者] "
                tmp_sql = tmp_sql & " 		,iw_syokai2_umu AS [石綿使用調査-調査結果の問合せ先-管理組合] "
                tmp_sql = tmp_sql & " 		,iw_syokai3_umu AS [石綿使用調査-調査結果の問合せ先-管理業者] "
                tmp_sql = tmp_sql & " 		,iw_syokai4_umu AS [石綿使用調査-調査結果の問合せ先-施工業者] "
                tmp_sql = tmp_sql & " 		/*,'' AS [石綿使用調査-調査結果の問合せ先-施工業者名]*/ "
                tmp_sql = tmp_sql & " 		,iw_cyosa_ymd AS [石綿使用調査-調査年月日] "
                tmp_sql = tmp_sql & " 		,iw_cyosa_kikan AS [石綿使用調査-実施機関] "
                tmp_sql = tmp_sql & " 		,iw_cyosa_hani AS [石綿使用調査-調査範囲] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN iw_use_umu = 0 THEN 0 "
                tmp_sql = tmp_sql & " 			WHEN iw_use_umu = 1 THEN 0 "
                tmp_sql = tmp_sql & " 			WHEN iw_use_umu = 2 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE 0	 "
                tmp_sql = tmp_sql & " 		 END AS [石綿使用調査-使用有無] "
                tmp_sql = tmp_sql & " 		,iw_use_area AS [石綿使用調査-使用箇所] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN iw_biko1 <> '' AND iw_biko2 <> '' AND iw_biko3 <> '' THEN iw_biko1 + '$0D$0A' + iw_biko2 + '$0D$0A' + iw_biko3 "
                tmp_sql = tmp_sql & " 			WHEN iw_biko1 <> '' AND iw_biko2 <> '' AND iw_biko3 = '' THEN iw_biko1 + '$0D$0A' + iw_biko2 "
                tmp_sql = tmp_sql & " 			WHEN iw_biko1 <> '' AND iw_biko2 = '' AND iw_biko3 <> '' THEN iw_biko1 + '$0D$0A' + iw_biko3 "
                tmp_sql = tmp_sql & " 			WHEN iw_biko1 = '' AND iw_biko2 <> '' AND iw_biko3 <> '' THEN iw_biko2 + '$0D$0A' + iw_biko3 "
                tmp_sql = tmp_sql & " 			WHEN iw_biko1 <> '' AND iw_biko2 = '' AND iw_biko3 = '' THEN iw_biko1 "
                tmp_sql = tmp_sql & " 			WHEN iw_biko1 = '' AND iw_biko2 <> '' AND iw_biko3 = '' THEN iw_biko2 "
                tmp_sql = tmp_sql & " 			WHEN iw_biko1 = '' AND iw_biko2 = '' AND iw_biko3 <> '' THEN iw_biko3 "
                tmp_sql = tmp_sql & " 			WHEN iw_biko1 = '' AND iw_biko2 = '' AND iw_biko3 = '' THEN '' "
                tmp_sql = tmp_sql & " 		 END AS [石綿使用調査-備考] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ts_sindan_umu = 0 THEN -1 "
                tmp_sql = tmp_sql & " 			WHEN ts_sindan_umu = 1 THEN 2 "
                tmp_sql = tmp_sql & " 			WHEN ts_sindan_umu = 2 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 		 END AS [耐震診断-診断有無]	 "
                tmp_sql = tmp_sql & " 		,ts_syokai1_umu AS [耐震診断-診断記録の問合せ先-所有者] "
                tmp_sql = tmp_sql & " 		,ts_syokai2_umu AS [耐震診断-診断記録の問合せ先-管理組合] "
                tmp_sql = tmp_sql & " 		,ts_syokai3_umu AS [耐震診断-診断記録の問合せ先-管理業者] "
                tmp_sql = tmp_sql & " 		,ts_syorui1_umu AS [耐震診断-耐震基準適合証明書の写し] "
                tmp_sql = tmp_sql & " 		,ts_syorui2_umu AS [耐震診断-住宅性能評価所の写し] "
                tmp_sql = tmp_sql & " 		,ts_syorui3_umu AS [耐震診断-耐震診断結果の写し] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ts_biko1 <> '' AND ts_biko2 <> '' AND ts_biko3 <> '' THEN ts_biko1 + '$0D$0A' + ts_biko2 + '$0D$0A' + ts_biko3 "
                tmp_sql = tmp_sql & " 			WHEN ts_biko1 <> '' AND ts_biko2 <> '' AND ts_biko3 = '' THEN ts_biko1 + '$0D$0A' + ts_biko2 "
                tmp_sql = tmp_sql & " 			WHEN ts_biko1 <> '' AND ts_biko2 = '' AND ts_biko3 <> '' THEN ts_biko1 + '$0D$0A' + ts_biko3 "
                tmp_sql = tmp_sql & " 			WHEN ts_biko1 = '' AND ts_biko2 <> '' AND ts_biko3 <> '' THEN ts_biko2 + '$0D$0A' + ts_biko3 "
                tmp_sql = tmp_sql & " 			WHEN ts_biko1 <> '' AND ts_biko2 = '' AND ts_biko3 = '' THEN ts_biko1 "
                tmp_sql = tmp_sql & " 			WHEN ts_biko1 = '' AND ts_biko2 <> '' AND ts_biko3 = '' THEN ts_biko2 "
                tmp_sql = tmp_sql & " 			WHEN ts_biko1 = '' AND ts_biko2 = '' AND ts_biko3 <> '' THEN ts_biko3 "
                tmp_sql = tmp_sql & " 			WHEN ts_biko1 = '' AND ts_biko2 = '' AND ts_biko3 = '' THEN '' "
                tmp_sql = tmp_sql & " 		 END AS [耐震診断-備考] "
                tmp_sql = tmp_sql & " 		,[土砂災害防止対策推進法] AS [法令-土砂災害警戒地域内外]				/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,[法令に基づく制限の概要（法令名）] AS [法令-法令分類]					/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,[法令に基づく制限の概要] AS [法令-法令内容]							/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,[敷地利用関係の種類(敷地が借地の場合)] AS [敷地利用-敷地利用種類]		/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,'' AS [敷地利用-契約期間開始] "
                tmp_sql = tmp_sql & " 		,[契約期間(敷地が賃借の場合）] AS [敷地利用-契約期間終了]				/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,[備考(敷地が借地の場合)] AS [敷地利用-備考]							/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,'' AS [入居率一覧対象開始日] "
                tmp_sql = tmp_sql & " 		,0 AS [部屋毎に業者が異なる場合フラグ] "
                tmp_sql = tmp_sql & " 		,'' AS [ゴミ出しに関する補足情報] "
                tmp_sql = tmp_sql & " 		,'' AS [登記情報の日付] "
                tmp_sql = tmp_sql & " 		,[所有権にかかる権利の有無] AS [所有権にかかる権利有無]					/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,[所有権にかかる権利の種類] AS [所有権にかかる権利の種類]				/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,0 AS [所有権以外の権利有無]	/*物件権利情報移行時にプログラム内で更新する*/ "
                tmp_sql = tmp_sql & " 		,'' AS [物件面積-建築面積] "
                tmp_sql = tmp_sql & " 		,'' AS [物件面積-建築面積坪数] "
                tmp_sql = tmp_sql & " 		,'' AS [物件面積-建築面積(登記)] "
                tmp_sql = tmp_sql & " 		,'' AS [物件面積-建築面積坪数(登記)] "
                tmp_sql = tmp_sql & " 		,tanto_no AS [自社担当者No] "
                tmp_sql = tmp_sql & " 		,kn_fax AS [管理形態-管理業者FAX] "
                tmp_sql = tmp_sql & " 		,0 AS [エレベータフラグ(家)] "
                tmp_sql = tmp_sql & " 		,-1 AS [法令-土砂災害特別警戒区域内外] "
                tmp_sql = tmp_sql & " 		,[宅地造成等規正法] AS [法令-造成宅地防災区域内外]						/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,[津波防災地域づくりに関する法律] AS [法令-津波災害警戒区域内外]		/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,[土砂災害防止対策推進法備考] AS [法令-土砂災害警戒地域備考]			/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,'' AS [法令-土砂災害特別警戒区域備考] "
                tmp_sql = tmp_sql & " 		,[宅地造成等規正法備考] AS [法令-造成宅地防災区域備考]					/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		,[津波防災地域づくりに関する法律備考] AS [法令-津波災害警戒区域備考]	/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg*/ "
                tmp_sql = tmp_sql & " 		/*,-1 AS [会計グループ分類]*/ "
                tmp_sql = tmp_sql & " 		,kn_name1 AS [管理形態-管理人名] "
                tmp_sql = tmp_sql & " 		,'' AS [その他建物構造] "
                tmp_sql = tmp_sql & " 		,0 AS [角地フラグ] "
                tmp_sql = tmp_sql & " 		,yoto1 AS [都市計画・用途地域1] "
                tmp_sql = tmp_sql & " 		,yoto2 AS [都市計画・用途地域2] "
                tmp_sql = tmp_sql & " 		,'' AS [管理人名SJIS] "
                tmp_sql = tmp_sql & " 		,'' AS [管理業者担当者名SJIS] "
                tmp_sql = tmp_sql & " 		,'' AS [管理業者名SJS] "
                tmp_sql = tmp_sql & " 		,as_car_aki AS [駐車場空き台数] "
                tmp_sql = tmp_sql & " 		,'' AS [バイク置き場空き台数] "
                tmp_sql = tmp_sql & " 		,'' AS [駐輪場空き台数] "
                tmp_sql = tmp_sql & " 		,syogaku_name AS [小学校区] "
                tmp_sql = tmp_sql & " 		,syogaku_kyori AS [小学校距離] "
                tmp_sql = tmp_sql & " 		,cyugaku_name AS [中学校区] "
                tmp_sql = tmp_sql & " 		,chugaku_kyori AS [中学校距離] "
                tmp_sql = tmp_sql & " 		,addkbn_name AS [エリア] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 物件詳細情報 add*/ "
                tmp_sql = tmp_sql & " 		,0 AS [非常用エレベーターフラグ] "
                tmp_sql = tmp_sql & " 	FROM bk_mst AS BK "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_brui AS BKRUI ON BK.brui_no = BKRUI.brui_no "
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
                tmp_sql = tmp_sql & " 	/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add sta*/ "
                tmp_sql = tmp_sql & " 	LEFT JOIN "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN kenri_umu = '無' THEN 0 "
                tmp_sql = tmp_sql & " 					WHEN kenri_umu = '有' THEN 1 "
                tmp_sql = tmp_sql & " 					ELSE 0 "
                tmp_sql = tmp_sql & " 				 END AS [所有権にかかる権利の有無] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN kenri_syurui = '仮登記(所有権移転)' THEN 1 "
                tmp_sql = tmp_sql & " 					WHEN kenri_syurui = '仮登記(所有権移転請求権)' THEN 1 "
                tmp_sql = tmp_sql & " 					WHEN kenri_syurui = '仮差押' THEN 2 "
                tmp_sql = tmp_sql & " 					WHEN kenri_syurui = '仮処分' THEN 3 "
                tmp_sql = tmp_sql & " 					WHEN kenri_syurui = '差押(含む参加差押)' THEN 4 "
                tmp_sql = tmp_sql & " 					WHEN kenri_syurui = '買戻特約' THEN 5 "
                tmp_sql = tmp_sql & " 					WHEN kenri_syurui = '予告登記' THEN 6 "
                tmp_sql = tmp_sql & " 					ELSE -1 "
                tmp_sql = tmp_sql & " 				 END AS [所有権にかかる権利の種類] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 						WHEN horeiname = '新住宅市街地開発法' THEN 1 "
                tmp_sql = tmp_sql & " 						WHEN horeiname = '新都市基盤整備法' THEN 2 "
                tmp_sql = tmp_sql & " 						WHEN horeiname = '流通業務市街地整備法' THEN 3 "
                tmp_sql = tmp_sql & " 						WHEN horeiname = '農地法' THEN 4 "
                tmp_sql = tmp_sql & " 						ELSE -1 "
                tmp_sql = tmp_sql & " 				 END AS [法令に基づく制限の概要（法令名）] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						(horeiname = '新住宅市街地開発法' OR horeiname = '新都市基盤整備法' OR horeiname = '流通業務市街地整備法' OR horeiname = '農地法') "
                tmp_sql = tmp_sql & " 						THEN horeigaiyo "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						NOT (horeiname = '新住宅市街地開発法' OR horeiname = '新都市基盤整備法' OR horeiname = '流通業務市街地整備法' OR horeiname = '農地法') "
                tmp_sql = tmp_sql & " 						AND horeiname = '' AND horeigaiyo = '' "
                tmp_sql = tmp_sql & " 						THEN '' "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						NOT (horeiname = '新住宅市街地開発法' OR horeiname = '新都市基盤整備法' OR horeiname = '流通業務市街地整備法' OR horeiname = '農地法') "
                tmp_sql = tmp_sql & " 						AND horeiname <> '' AND horeigaiyo = '' "
                tmp_sql = tmp_sql & " 						THEN horeiname "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						NOT (horeiname = '新住宅市街地開発法' OR horeiname = '新都市基盤整備法' OR horeiname = '流通業務市街地整備法' OR horeiname = '農地法') "
                tmp_sql = tmp_sql & " 						AND horeiname = '' AND horeigaiyo <> '' "
                tmp_sql = tmp_sql & " 						THEN horeigaiyo "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						NOT (horeiname = '新住宅市街地開発法' OR horeiname = '新都市基盤整備法' OR horeiname = '流通業務市街地整備法' OR horeiname = '農地法') "
                tmp_sql = tmp_sql & " 						AND horeiname <> '' AND horeigaiyo <> '' "
                tmp_sql = tmp_sql & " 						THEN "
                tmp_sql = tmp_sql & " 							horeiname + '$0D$0A' + horeigaiyo "
                tmp_sql = tmp_sql & " 				 END AS [法令に基づく制限の概要] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN sikitisyurui = '借地権(旧法)' THEN 1 "
                tmp_sql = tmp_sql & " 					WHEN sikitisyurui = '借地権(新法)' THEN 2 "
                tmp_sql = tmp_sql & " 					WHEN sikitisyurui = '定期借地権' THEN 3 "
                tmp_sql = tmp_sql & " 					WHEN sikitisyurui = '使用賃借' THEN 4 "
                tmp_sql = tmp_sql & " 					WHEN sikitisyurui = '一時賃貸借' THEN 4 "
                tmp_sql = tmp_sql & " 					ELSE -1 "
                tmp_sql = tmp_sql & " 				 END AS [敷地利用関係の種類(敷地が借地の場合)] "
                tmp_sql = tmp_sql & " 				,sikitikeiyakukikan AS [契約期間(敷地が賃借の場合）]	/*プログラム内で処理*/ "
                tmp_sql = tmp_sql & " 				 ,CASE "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						(sikitisyurui = '借地権(旧法)' OR sikitisyurui = '借地権(新法)' OR sikitisyurui = '定期借地権' OR sikitisyurui = '使用賃借' OR sikitisyurui = '一時賃貸借') "
                tmp_sql = tmp_sql & " 						THEN sikitibiko "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						NOT (sikitisyurui = '借地権(旧法)' OR sikitisyurui = '借地権(新法)' OR sikitisyurui = '定期借地権' OR sikitisyurui = '使用賃借' OR sikitisyurui = '一時賃貸借') "
                tmp_sql = tmp_sql & " 						AND sikitisyurui = '' AND sikitibiko = '' "
                tmp_sql = tmp_sql & " 						THEN '' "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						NOT (sikitisyurui = '借地権(旧法)' OR sikitisyurui = '借地権(新法)' OR sikitisyurui = '定期借地権' OR sikitisyurui = '使用賃借' OR sikitisyurui = '一時賃貸借') "
                tmp_sql = tmp_sql & " 						AND sikitisyurui <> '' AND sikitibiko = '' "
                tmp_sql = tmp_sql & " 						THEN sikitisyurui "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						NOT (sikitisyurui = '借地権(旧法)' OR sikitisyurui = '借地権(新法)' OR sikitisyurui = '定期借地権' OR sikitisyurui = '使用賃借' OR sikitisyurui = '一時賃貸借') "
                tmp_sql = tmp_sql & " 						AND sikitisyurui = '' AND sikitibiko <> '' "
                tmp_sql = tmp_sql & " 						THEN sikitibiko "
                tmp_sql = tmp_sql & " 					WHEN "
                tmp_sql = tmp_sql & " 						NOT (sikitisyurui = '借地権(旧法)' OR sikitisyurui = '借地権(新法)' OR sikitisyurui = '定期借地権' OR sikitisyurui = '使用賃借' OR sikitisyurui = '一時賃貸借') "
                tmp_sql = tmp_sql & " 						AND sikitisyurui <> '' AND sikitibiko <> '' "
                tmp_sql = tmp_sql & " 						THEN "
                tmp_sql = tmp_sql & " 							sikitisyurui + '$0D$0A' + sikitibiko			 "
                tmp_sql = tmp_sql & " 				 END AS [備考(敷地が借地の場合)] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN takutizosei = '造成宅地防災区域内' THEN 1 "
                tmp_sql = tmp_sql & " 					WHEN takutizosei = '造成宅地防災区域外' THEN 2 "
                tmp_sql = tmp_sql & " 					ELSE -1	  "
                tmp_sql = tmp_sql & " 				 END AS [宅地造成等規正法] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN (takutizosei = '造成宅地防災区域内' OR takutizosei = '造成宅地防災区域外') THEN '' "
                tmp_sql = tmp_sql & " 					ELSE takutizosei "
                tmp_sql = tmp_sql & " 				 END AS [宅地造成等規正法備考] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN dosyasaigai = '土砂災害警戒区域内' THEN 1 "
                tmp_sql = tmp_sql & " 					WHEN dosyasaigai = '土砂災害警戒区域外' THEN 2 "
                tmp_sql = tmp_sql & " 					ELSE -1	  "
                tmp_sql = tmp_sql & " 				 END AS [土砂災害防止対策推進法] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN (dosyasaigai = '土砂災害警戒区域内' OR dosyasaigai = '土砂災害警戒区域外') THEN '' "
                tmp_sql = tmp_sql & " 					ELSE dosyasaigai "
                tmp_sql = tmp_sql & " 				 END AS [土砂災害防止対策推進法備考]	  "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN tunamibosai = '津波災害警戒区域内' THEN 1 "
                tmp_sql = tmp_sql & " 					WHEN tunamibosai = '津波災害警戒区域外' THEN 2 "
                tmp_sql = tmp_sql & " 					ELSE -1	  "
                tmp_sql = tmp_sql & " 				 END AS [津波防災地域づくりに関する法律] "
                tmp_sql = tmp_sql & " 				,CASE "
                tmp_sql = tmp_sql & " 					WHEN (tunamibosai = '津波災害警戒区域内' OR tunamibosai = '津波災害警戒区域外') THEN '' "
                tmp_sql = tmp_sql & " 					ELSE tunamibosai "
                tmp_sql = tmp_sql & " 				 END AS [津波防災地域づくりに関する法律備考]	  "
                tmp_sql = tmp_sql & " 			FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT "
                tmp_sql = tmp_sql & " 					 bk_no "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '所有権にかかる権利の有無' THEN jyuyo_lstname ELSE '' END) AS kenri_umu "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '所有権にかかる権利の種類' THEN jyuyo_lstname ELSE '' END) AS kenri_syurui "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '法令に基づく制限の概要（法令名）' THEN jyuyo_lstname ELSE '' END) AS horeiname "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '法令に基づく制限の概要' THEN jyuyo_lstname ELSE '' END) AS horeigaiyo "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '敷地利用関係の種類(敷地が借地の場合)' THEN jyuyo_lstname ELSE '' END) AS sikitisyurui "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '契約期間(敷地が賃借の場合）' THEN jyuyo_lstname ELSE '' END) AS sikitikeiyakukikan "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '備考(敷地が借地の場合)' THEN jyuyo_lstname ELSE '' END) AS sikitibiko "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '宅地造成等規正法' THEN jyuyo_lstname ELSE '' END) AS takutizosei "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '土砂災害防止対策推進法' THEN jyuyo_lstname ELSE '' END) AS dosyasaigai "
                tmp_sql = tmp_sql & " 					,MAX(CASE WHEN jyuyo_name = '津波防災地域づくりに関する法律' THEN jyuyo_lstname ELSE '' END) AS tunamibosai "
                tmp_sql = tmp_sql & " 				FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					/*見出しがデフォルトの重要事項を抽出 sta*/ "
                tmp_sql = tmp_sql & " 					SELECT  "
                tmp_sql = tmp_sql & " 						 bk_no "
                tmp_sql = tmp_sql & " 						,jyuyo_name "
                tmp_sql = tmp_sql & " 						,jyuyo_lstname "
                tmp_sql = tmp_sql & " 					FROM bk_jyuyo AS BKJ "
                tmp_sql = tmp_sql & " 					LEFT JOIN "
                tmp_sql = tmp_sql & " 						( "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT * FROM m_jyuyo "
                tmp_sql = tmp_sql & " 								WHERE jyuyo_kbn = 1 "
                tmp_sql = tmp_sql & " 							) AS VW "
                tmp_sql = tmp_sql & " 						) AS JYUYOTITLE "
                tmp_sql = tmp_sql & " 					ON BKJ.jyuyo_no = JYUYOTITLE.jyuyo_no "
                tmp_sql = tmp_sql & " 					WHERE jyuyo_name IN ( "
                tmp_sql = tmp_sql & " 											 '所有権にかかる権利の有無' "
                tmp_sql = tmp_sql & " 											,'所有権にかかる権利の種類' "
                tmp_sql = tmp_sql & " 											,'法令に基づく制限の概要（法令名）' "
                tmp_sql = tmp_sql & " 											,'法令に基づく制限の概要' "
                tmp_sql = tmp_sql & " 											,'敷地利用関係の種類(敷地が借地の場合)' "
                tmp_sql = tmp_sql & " 											,'契約期間(敷地が賃借の場合）' "
                tmp_sql = tmp_sql & " 											,'備考(敷地が借地の場合)' "
                tmp_sql = tmp_sql & " 											,'宅地造成等規正法' "
                tmp_sql = tmp_sql & " 											,'土砂災害防止対策推進法' "
                tmp_sql = tmp_sql & " 											,'津波防災地域づくりに関する法律'	 "
                tmp_sql = tmp_sql & " 										) "
                tmp_sql = tmp_sql & " 					/*見出しがデフォルトの重要事項を抽出 end*/ "
                tmp_sql = tmp_sql & " 				) AS VW1 "
                tmp_sql = tmp_sql & " 				GROUP BY bk_no "
                tmp_sql = tmp_sql & " 			) AS VW	 "
                tmp_sql = tmp_sql & " 		 "
                tmp_sql = tmp_sql & " 		) AS BKKENRIHOREI "
                tmp_sql = tmp_sql & " 	ON BK.bk_no = BKKENRIHOREI.bk_no "
                tmp_sql = tmp_sql & " 	/*2016.04.08 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -add end*/ "
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

#Region "物件所有者情報"

    Public Class Bk_mst_syo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件No],[管理No]"

                '2016.04.14 物件所有期間の取得方法変更 -chg sta
                'tmp_sql = tmp_sql & " SELECT "
                'tmp_sql = tmp_sql & " 	 BKM.bk_no AS [物件No] "
                'tmp_sql = tmp_sql & " 	,ISNULL(BKK.kn_no,1) AS [管理No]	/*物件管理情報が無い場合も所有情報は移行する(10では所有情報が必須のため管理No1でダミーを作成)*/ "
                'tmp_sql = tmp_sql & " 	,so_no_kasi AS [貸主１No] "
                'tmp_sql = tmp_sql & " 	,so_no_syo AS [所有者１No] "
                'tmp_sql = tmp_sql & " 	,so_no_kasi2 AS [貸主２No] "
                'tmp_sql = tmp_sql & " 	,so_no_syo2 AS [所有者２No] "
                'tmp_sql = tmp_sql & " 	,ISNULL([最古契約日月初],'1900-01-01 00:00:00.000') AS [所有期間開始]	/*契約情報が存在しない場合はデフォルトを「1900-01-01 00:00:00.000」にしておく*/ "
                'tmp_sql = tmp_sql & " 	,'2050-12-31 00:00:00.000' AS [所有期間終了] "
                'tmp_sql = tmp_sql & " FROM bk_mst AS BKM "
                'tmp_sql = tmp_sql & " LEFT JOIN bk_kanri AS BKK ON BKM.bk_no = BKK.bk_no "
                'tmp_sql = tmp_sql & " LEFT JOIN "
                'tmp_sql = tmp_sql & " 	( "
                'tmp_sql = tmp_sql & " 		/*最古契約日を抽出し月初に変換 sta*/ "
                'tmp_sql = tmp_sql & " 		SELECT * FROM "
                'tmp_sql = tmp_sql & " 		( "
                'tmp_sql = tmp_sql & " 			SELECT "
                'tmp_sql = tmp_sql & " 				 bk_no "
                'tmp_sql = tmp_sql & " 				,[契約日] "
                'tmp_sql = tmp_sql & " 				,DATEADD(DAY,(DAY([契約日]) - 1 ) * (-1),[契約日]) AS [最古契約日月初] "
                'tmp_sql = tmp_sql & " 			FROM "
                'tmp_sql = tmp_sql & " 			( "
                'tmp_sql = tmp_sql & " 				SELECT "
                'tmp_sql = tmp_sql & " 					 bk_no "
                'tmp_sql = tmp_sql & " 					,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY [契約日]) AS [最古契約日抽出用] "
                'tmp_sql = tmp_sql & " 					,[契約日] "
                'tmp_sql = tmp_sql & " 				FROM "
                'tmp_sql = tmp_sql & " 				( "
                'tmp_sql = tmp_sql & " 					/*契約日、契約開始日の集約 sta*/ "
                'tmp_sql = tmp_sql & " 					SELECT * FROM "
                'tmp_sql = tmp_sql & " 					(						 "
                'tmp_sql = tmp_sql & " 						SELECT bk_no,ky_date AS [契約日] FROM ky_kosinkai "
                'tmp_sql = tmp_sql & " 						UNION "
                'tmp_sql = tmp_sql & " 						SELECT bk_no,ky_start_ymd  AS [契約日] FROM ky_kosinkai "
                'tmp_sql = tmp_sql & " 					) AS VW "
                'tmp_sql = tmp_sql & " 					WHERE [契約日] IS NOT NULL "
                'tmp_sql = tmp_sql & " 					/*契約日、契約開始日の集約 end*/ "
                'tmp_sql = tmp_sql & " 				) AS KYYMD "
                'tmp_sql = tmp_sql & " 			) AS VW "
                'tmp_sql = tmp_sql & " 			WHERE [最古契約日抽出用] = 1 "
                'tmp_sql = tmp_sql & " 		) AS VW "
                'tmp_sql = tmp_sql & " 		/*最古契約日を抽出し月初に変換 end*/ "
                'tmp_sql = tmp_sql & " 	) AS KYYMD "
                'tmp_sql = tmp_sql & " ON BKM.bk_no = KYYMD.bk_no "

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件No],[管理No],[貸主１No],[所有者１No],[貸主２No],[所有者２No],[所有期間開始],[所有期間終了]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                tmp_sql = tmp_sql & " /************************************************************ "
                tmp_sql = tmp_sql & " 所有期間に関して "
                tmp_sql = tmp_sql & "   ・送金ルール期間の親となる期間 "
                tmp_sql = tmp_sql & "   ・所有期間内に送金ルール期間が含まれている必要がある "
                tmp_sql = tmp_sql & "   ・物件情報作成時に必須項目となるため、データが存在しない "
                tmp_sql = tmp_sql & "     (物件管理情報が存在しない) 場合はデフォルト値「1980-01-01 00:00:00.000」を設定する "
                tmp_sql = tmp_sql & "   ・所有期間終了日は以下の条件で設定する "
                tmp_sql = tmp_sql & "     ・最新の物件管理情報(kn_noが最大)の場合は終了日をNULL "
                tmp_sql = tmp_sql & "     ・上記以外はそのまま取得 "
                tmp_sql = tmp_sql & " ************************************************************/ "
                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 BKM.bk_no AS [物件No] "
                tmp_sql = tmp_sql & " 		,ISNULL(BKK.kn_no,1) AS [管理No]	/*物件管理情報が無い場合も所有情報は移行する(10では所有情報が必須のため管理No1でダミーを作成)*/ "
                tmp_sql = tmp_sql & " 		,so_no_kasi AS [貸主１No] "
                tmp_sql = tmp_sql & " 		,so_no_syo AS [所有者１No] "
                tmp_sql = tmp_sql & " 		,so_no_kasi2 AS [貸主２No] "
                tmp_sql = tmp_sql & " 		,so_no_syo2 AS [所有者２No] "
                tmp_sql = tmp_sql & " 		,ISNULL(start_ym,'1980-01-01 00:00:00.000') AS [所有期間開始]	/*物件管理情報の適用開始年月を所有期間開始年月に設定する*/ "
                tmp_sql = tmp_sql & " 		,end_ym AS [所有期間終了] "
                tmp_sql = tmp_sql & " 	FROM bk_mst AS BKM "
                tmp_sql = tmp_sql & " 	LEFT JOIN "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		/*最新の物件管理情報(kn_noが最大)の場合は終了日をNULLとして抽出 sta*/ "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 bk_no "
                tmp_sql = tmp_sql & " 			,kn_no "
                tmp_sql = tmp_sql & " 			,start_ym "
                tmp_sql = tmp_sql & " 			,NULL AS end_ym "
                tmp_sql = tmp_sql & " 		FROM bk_kanri AS BKK "
                tmp_sql = tmp_sql & " 		WHERE EXISTS "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT * FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 bk_no "
                tmp_sql = tmp_sql & " 						,kn_no "
                tmp_sql = tmp_sql & " 						,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY kn_no DESC) AS 抽出用 "
                tmp_sql = tmp_sql & " 					FROM bk_kanri "
                tmp_sql = tmp_sql & " 				) AS VW "
                tmp_sql = tmp_sql & " 				WHERE 抽出用 = 1 "
                tmp_sql = tmp_sql & " 			) AS NOTKOSIN "
                tmp_sql = tmp_sql & " 			WHERE BKK.bk_no = NOTKOSIN.bk_no "
                tmp_sql = tmp_sql & " 			AND   BKK.kn_no = NOTKOSIN.kn_no "
                tmp_sql = tmp_sql & " 		) "
                tmp_sql = tmp_sql & " 		/*最新の物件管理情報(kn_noが最大)の場合は終了日をNULLとして抽出 end*/ "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		/*最新の物件管理情報(kn_noが最大)以外の場合は終了日を抽出 sta*/ "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 bk_no "
                tmp_sql = tmp_sql & " 			,kn_no "
                tmp_sql = tmp_sql & " 			,start_ym "
                tmp_sql = tmp_sql & " 			,end_ym "
                tmp_sql = tmp_sql & " 		FROM bk_kanri AS BKK "
                tmp_sql = tmp_sql & " 		WHERE EXISTS "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT * FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 bk_no "
                tmp_sql = tmp_sql & " 						,kn_no "
                tmp_sql = tmp_sql & " 						,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY kn_no DESC) AS 抽出用 "
                tmp_sql = tmp_sql & " 					FROM bk_kanri "
                tmp_sql = tmp_sql & " 				) AS VW "
                tmp_sql = tmp_sql & " 				WHERE 抽出用 <> 1 "
                tmp_sql = tmp_sql & " 			) AS NOTKOSIN "
                tmp_sql = tmp_sql & " 			WHERE BKK.bk_no = NOTKOSIN.bk_no "
                tmp_sql = tmp_sql & " 			AND   BKK.kn_no = NOTKOSIN.kn_no "
                tmp_sql = tmp_sql & " 		) "
                tmp_sql = tmp_sql & " 		/*最新の物件管理情報(kn_noが最大)以外の場合は終了日を抽出 sta*/ "
                tmp_sql = tmp_sql & " 	) AS BKK ON BKM.bk_no = BKK.bk_no "
                tmp_sql = tmp_sql & " ) AS VW "
                '2016.04.14 物件所有期間の取得方法変更 -chg end

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

#Region "物件ゴミ情報"

    Public Class Bk_mst_gomi_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[ゴミ情報No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[ゴミ情報No],[ゴミ分類],[ゴミ出し曜日]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,1 AS [ゴミ情報No] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN gomi1 = '' THEN '可燃物' "
                tmp_sql = tmp_sql & " 			WHEN gomi1 <> '' THEN gomi1 "
                tmp_sql = tmp_sql & " 		 END AS [ゴミ分類] "
                tmp_sql = tmp_sql & " 		,'' AS [ゴミ出し曜日] "
                tmp_sql = tmp_sql & " 	FROM bk_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,2 AS [ゴミ情報No] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN gomi2 = '' THEN '不燃物' "
                tmp_sql = tmp_sql & " 			WHEN gomi2 <> '' THEN gomi2 "
                tmp_sql = tmp_sql & " 		 END AS [ゴミ分類] "
                tmp_sql = tmp_sql & " 		,'' AS [ゴミ出し曜日] "
                tmp_sql = tmp_sql & " 	FROM bk_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,3 AS [ゴミ情報No] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN gomi3 = '' THEN '空き瓶・ペットボトル' "
                tmp_sql = tmp_sql & " 			WHEN gomi3 <> '' THEN gomi3 "
                tmp_sql = tmp_sql & " 		 END AS [ゴミ分類] "
                tmp_sql = tmp_sql & " 		,'' AS [ゴミ出し曜日] "
                tmp_sql = tmp_sql & " 	FROM bk_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,4 AS [ゴミ情報No] "
                tmp_sql = tmp_sql & " 		,gomi4 AS [ゴミ分類] "
                tmp_sql = tmp_sql & " 		,'' AS [ゴミ出し曜日] "
                tmp_sql = tmp_sql & " 	FROM bk_mst "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,5 AS [ゴミ情報No] "
                tmp_sql = tmp_sql & " 		,gomi5 AS [ゴミ分類] "
                tmp_sql = tmp_sql & " 		,'' AS [ゴミ出し曜日] "
                tmp_sql = tmp_sql & " 	FROM bk_mst "
                tmp_sql = tmp_sql & " ) AS BKDUST "
                tmp_sql = tmp_sql & " WHERE [ゴミ分類] <> '' "

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

#Region "物件変動費親メーター情報(汎用ツールのみ)"

    'V7には存在しない項目
    'クエリだけ設けておく

    Public Class Bk_mst_hendo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[行No]"

                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 '' AS [物件NO] "
                tmp_sql = tmp_sql & " 	,'' AS [行No] "
                tmp_sql = tmp_sql & " 	,'' AS [メーター分類 (水道/ガス/電気/灯油/その他)] "
                tmp_sql = tmp_sql & " 	,'' AS [メーター名(部屋no・フロア名など)] "
                tmp_sql = tmp_sql & " 	,'' AS [変動費入金項目] "
                tmp_sql = tmp_sql & " 	,'' AS [子メーター検針] "
                tmp_sql = tmp_sql & " 	,'' AS [備考] "
                tmp_sql = tmp_sql & " 	,'' AS [使用フラグ] "
                tmp_sql = tmp_sql & " 	,'' AS [請求月] "
                tmp_sql = tmp_sql & " 	,'' AS [単位] "
                tmp_sql = tmp_sql & " 	,'' AS [単位(SJIS)] "

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

#Region "物件鍵情報"

    Public Class Bk_mst_kagi_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】 2016.04.06 物件鍵情報の移行処理修正
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                sortstr = "[物件NO],[鍵タイトル名]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件No],[鍵No],[鍵タイトル名],[鍵本数],[備考],[保管場所],[業者間での情報共有]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                Dim tmp_sql As String = ""
                Select Case True
                    Case Njc.Frm.MainFrm.optHyKagi.Checked  '部屋鍵情報から取得する場合
                        'tmp_sql = tmp_sql & " SELECT * FROM "
                        tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                        tmp_sql = tmp_sql & " ( "
                        tmp_sql = tmp_sql & " 	SELECT "
                        tmp_sql = tmp_sql & " 		 bk_no AS [物件No] "
                        tmp_sql = tmp_sql & "     	/*20160613 鍵情報の取得処理修正 chg sta*/ "
                        tmp_sql = tmp_sql & "     	/*,hy_no "
                        tmp_sql = tmp_sql & "     	,kagimei_no "
                        tmp_sql = tmp_sql & "     	*/ "
                        tmp_sql = tmp_sql & "     	,kagimei_no AS [鍵No] "
                        tmp_sql = tmp_sql & "     	/*20160613 鍵情報の取得処理修正 chg end*/ "
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
                        tmp_sql = tmp_sql & "     	/*20160613 鍵情報の取得処理修正 chg sta*/ "
                        tmp_sql = tmp_sql & "     	/*,hy_no "
                        tmp_sql = tmp_sql & "     	,kagimei_no "
                        tmp_sql = tmp_sql & "     	*/ "
                        tmp_sql = tmp_sql & "     	,kagimei_no AS [鍵No] "
                        tmp_sql = tmp_sql & "     	/*20160613 鍵情報の取得処理修正 chg end*/ "
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

#Region "物件権利情報"

    Public Class Bk_Jyuyo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[権利情報No]"

                '2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg sta
                'tmp_sql = tmp_sql & " SELECT "
                'tmp_sql = tmp_sql & " 	 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 	,1 AS [権利情報No] "
                'tmp_sql = tmp_sql & " 	,CASE "
                'tmp_sql = tmp_sql & " 		WHEN CHARINDEX('根',jyuyo_lstname) = 0 AND REPLACE(jyuyo_lstname,'抵当権','') <> jyuyo_lstname THEN 1 "
                'tmp_sql = tmp_sql & " 		WHEN CHARINDEX('根',jyuyo_lstname) <> 0 AND REPLACE(jyuyo_lstname,'根抵当権','') <> jyuyo_lstname THEN 2 "
                'tmp_sql = tmp_sql & " 		WHEN REPLACE(jyuyo_lstname,'賃借権','') <> jyuyo_lstname THEN 3 "
                'tmp_sql = tmp_sql & " 		WHEN REPLACE(jyuyo_lstname,'地役権','') <> jyuyo_lstname THEN 4 "
                'tmp_sql = tmp_sql & " 		WHEN REPLACE(jyuyo_lstname,'地上権','') <> jyuyo_lstname THEN 5 "
                'tmp_sql = tmp_sql & " 		ELSE 6 "
                'tmp_sql = tmp_sql & " 	 END AS [所有権以外の権利の種類] "
                'tmp_sql = tmp_sql & " 	,CASE "
                'tmp_sql = tmp_sql & " 		WHEN CHARINDEX('根',jyuyo_lstname) = 0 AND REPLACE(jyuyo_lstname,'抵当権','') <> jyuyo_lstname THEN REPLACE(jyuyo_lstname,'抵当権','') "
                'tmp_sql = tmp_sql & " 		WHEN CHARINDEX('根',jyuyo_lstname) <> 0 AND REPLACE(jyuyo_lstname,'根抵当権','') <> jyuyo_lstname THEN REPLACE(jyuyo_lstname,'根抵当権','') "
                'tmp_sql = tmp_sql & " 		WHEN REPLACE(jyuyo_lstname,'賃借権','') <> jyuyo_lstname THEN REPLACE(jyuyo_lstname,'賃借権','') "
                'tmp_sql = tmp_sql & " 		WHEN REPLACE(jyuyo_lstname,'地役権','') <> jyuyo_lstname THEN REPLACE(jyuyo_lstname,'地役権','') "
                'tmp_sql = tmp_sql & " 		WHEN REPLACE(jyuyo_lstname,'地上権','') <> jyuyo_lstname THEN REPLACE(jyuyo_lstname,'地上権','') "
                'tmp_sql = tmp_sql & " 		ELSE jyuyo_lstname "
                'tmp_sql = tmp_sql & " 	 END AS [所有権以外の権利備考] "
                'tmp_sql = tmp_sql & " FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 bk_no "
                'tmp_sql = tmp_sql & " 		,jyuyo_lstname "
                'tmp_sql = tmp_sql & " 	FROM bk_jyuyo "
                'tmp_sql = tmp_sql & " 	WHERE bk_no IN "
                'tmp_sql = tmp_sql & " 	( "
                'tmp_sql = tmp_sql & " 		SELECT bk_no FROM bk_jyuyo "
                'tmp_sql = tmp_sql & " 		WHERE jyuyo_no = 3 "
                'tmp_sql = tmp_sql & " 		AND   RTRIM(LTRIM(jyuyo_lstname)) = '有' "
                'tmp_sql = tmp_sql & " 	) "
                'tmp_sql = tmp_sql & " 	AND jyuyo_no = 4 "
                'tmp_sql = tmp_sql & " ) AS BKKENRI "

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件No],[権利情報No],[所有権以外の権利の有無],[所有権以外の権利の種類],[内容]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 bk_no AS [物件No] "
                tmp_sql = tmp_sql & " 			,1 AS [権利情報No] "
                tmp_sql = tmp_sql & " 			,CASE "
                tmp_sql = tmp_sql & " 				WHEN kenri_umu = '無' THEN 0 "
                tmp_sql = tmp_sql & " 				WHEN kenri_umu = '有' THEN 1 "
                tmp_sql = tmp_sql & " 				ELSE 0 "
                tmp_sql = tmp_sql & " 			 END AS [所有権以外の権利の有無] "
                tmp_sql = tmp_sql & " 			,CASE "
                tmp_sql = tmp_sql & " 				WHEN kenri_syurui = '抵当権(含む仮登記)' THEN 1 "
                tmp_sql = tmp_sql & " 				WHEN kenri_syurui = '根抵当権(含む仮登記)' THEN 2 "
                tmp_sql = tmp_sql & " 				WHEN kenri_syurui = '質権' THEN 6 "
                tmp_sql = tmp_sql & " 				WHEN kenri_syurui = '先取特権' THEN 6 "
                tmp_sql = tmp_sql & " 				WHEN kenri_syurui = '賃借権(含む仮登記)' THEN 3 "
                tmp_sql = tmp_sql & " 			 END AS [所有権以外の権利の種類] "
                tmp_sql = tmp_sql & " 			,CASE "
                tmp_sql = tmp_sql & " 				WHEN kenri_syurui = '質権' THEN kenri_syurui "
                tmp_sql = tmp_sql & " 				WHEN kenri_syurui = '先取特権' THEN kenri_syurui "
                tmp_sql = tmp_sql & " 				ELSE '' "
                tmp_sql = tmp_sql & " 			 END AS [内容]	  "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				bk_no "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN jyuyo_name = '所有権以外の権利の有無' THEN jyuyo_lstname ELSE '' END) AS kenri_umu "
                tmp_sql = tmp_sql & " 				,MAX(CASE WHEN jyuyo_name = '所有権以外の権利の種類' THEN jyuyo_lstname ELSE '' END) AS kenri_syurui "
                tmp_sql = tmp_sql & " 			FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT  "
                tmp_sql = tmp_sql & " 					 bk_no "
                tmp_sql = tmp_sql & " 					,jyuyo_name "
                tmp_sql = tmp_sql & " 					,jyuyo_lstname "
                tmp_sql = tmp_sql & " 				FROM bk_jyuyo AS BKJ "
                tmp_sql = tmp_sql & " 				LEFT JOIN "
                tmp_sql = tmp_sql & " 					( "
                tmp_sql = tmp_sql & " 						SELECT * FROM "
                tmp_sql = tmp_sql & " 						( "
                tmp_sql = tmp_sql & " 							SELECT * FROM m_jyuyo "
                tmp_sql = tmp_sql & " 							WHERE jyuyo_kbn = 1 "
                tmp_sql = tmp_sql & " 						) AS VW "
                tmp_sql = tmp_sql & " 					) AS JYUYOTITLE "
                tmp_sql = tmp_sql & " 				ON BKJ.jyuyo_no = JYUYOTITLE.jyuyo_no "
                tmp_sql = tmp_sql & " 				WHERE jyuyo_name IN ('所有権以外の権利の有無','所有権以外の権利の種類') "
                tmp_sql = tmp_sql & " 			) AS VW1 "
                tmp_sql = tmp_sql & " 			GROUP BY bk_no "
                tmp_sql = tmp_sql & " 		) AS VW2 "
                tmp_sql = tmp_sql & " 	) AS TOTAL "
                tmp_sql = tmp_sql & " 	/*20160624 物件権利情報 抽出処理の修正 add*/ "
                tmp_sql = tmp_sql & " 	WHERE [所有権以外の権利の種類] IS NOT NULL "
                tmp_sql = tmp_sql & " ) AS VW "
                '2016.04.06 物件重要事項情報の見出しがデフォルトの場合のみ取得するように修正 -chg end

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

#Region "物件交通情報"

    Public Class Bk_eki_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[内部No]"

                '20160519 EXEUpdateに伴う修正 物件交通情報 -chg sta
                '革命10本体の仕様変更があったため抽出クエリを修正
                ''移行仕様作成時と仕様が異なっている可能性があるため旧クエリをコメントアウトして再作成
                ''tmp_sql = tmp_sql & " SELECT "
                ''tmp_sql = tmp_sql & " 	 bk_no AS [物件NO] "
                ''tmp_sql = tmp_sql & " 	,naibu_no AS [沿線番号] "
                ''tmp_sql = tmp_sql & " 	,ensen_no AS [沿線No] "
                ''tmp_sql = tmp_sql & " 	,eki_no AS [駅No] "
                ''tmp_sql = tmp_sql & " 	,'' AS [始発フラグ] "
                ''tmp_sql = tmp_sql & " 	,eki_kyori AS [距離] "
                ''tmp_sql = tmp_sql & " 	,ho_time AS [徒歩] "
                ''tmp_sql = tmp_sql & " 	,car_time AS [車] "
                ''tmp_sql = tmp_sql & " 	,'' AS [バス] "
                ''tmp_sql = tmp_sql & " 	,bus_rosen_name AS [バス会社] "
                ''tmp_sql = tmp_sql & " 	,bus_name AS [バス停] "
                ''tmp_sql = tmp_sql & " 	,busho_time AS [バス停徒歩] "
                ''tmp_sql = tmp_sql & " 	,'' AS [バス停までの距離] "
                ''tmp_sql = tmp_sql & " 	,'' AS [設定する交通のタイプ] "
                ''tmp_sql = tmp_sql & " 	,bus_rosen_name AS [バス系統・路線名(駅までバスを利用)] "
                ''tmp_sql = tmp_sql & " 	,bus_name AS [バス停名(駅までバスを利用)] "
                ''tmp_sql = tmp_sql & " 	,'' AS [バス停(駅までバスを利用)までの道路距離] "
                ''tmp_sql = tmp_sql & " 	,bus_time AS [バス停(駅までバスを利用)までの徒歩(分)] "
                ''tmp_sql = tmp_sql & " 	,'' AS [交通手段] "
                ''tmp_sql = tmp_sql & " 	,'' AS [車距離] "
                ''tmp_sql = tmp_sql & " FROM bk_eki "

                ''2016.04.06 物件交通情報の取得方法修正 -chg sta
                ''tmp_sql = tmp_sql & " SELECT "
                ''tmp_sql = tmp_sql & " 	 bk_no AS [物件NO] "
                ''tmp_sql = tmp_sql & " 	,naibu_no AS [沿線番号] "
                ''tmp_sql = tmp_sql & " 	,ensen_no AS [沿線No] "
                ''tmp_sql = tmp_sql & " 	,eki_no AS [駅No] "
                ''tmp_sql = tmp_sql & " 	,0 AS [始発フラグ] "
                ''tmp_sql = tmp_sql & " 	/*,eki_kyori AS [距離]*/ "
                ''tmp_sql = tmp_sql & " 	,ho_time * 80 AS [距離]	/*eki_kyoriはV7で連動用項目であることと、革命10では自動で80m/minを元に算出しているので時間から距離を算出して移行する*/ "
                ''tmp_sql = tmp_sql & " 	,ho_time AS [徒歩] "
                ''tmp_sql = tmp_sql & " 	,car_time AS [車] "
                ''tmp_sql = tmp_sql & " 	,bus_time AS [バス] "
                ''tmp_sql = tmp_sql & " 	,bus_rosen_name AS [バス会社] "
                ''tmp_sql = tmp_sql & " 	,bus_name AS [バス停] "
                ''tmp_sql = tmp_sql & " 	/*,busho_time AS [バス停徒歩]*/	 "
                ''tmp_sql = tmp_sql & " 	,'' AS [バス停徒歩]	/*使用されていない可能性があるため空列を設定しておく*/ "
                ''tmp_sql = tmp_sql & " 	,'' AS [バス停までの距離] "
                ''tmp_sql = tmp_sql & " 	,1 AS [設定する交通のタイプ] "
                ''tmp_sql = tmp_sql & " 	,bus_rosen_name AS [バス系統・路線名(駅までバスを利用)] "
                ''tmp_sql = tmp_sql & " 	,bus_name AS [バス停名(駅までバスを利用)] "
                ''tmp_sql = tmp_sql & " 	/*,'' AS [バス停(駅までバスを利用)までの道路距離]*/ "
                ''tmp_sql = tmp_sql & " 	/*,bus_time AS [バス停(駅までバスを利用)までの徒歩(分)]*/ "
                ''tmp_sql = tmp_sql & " 	,busho_time * 80 AS [バス停(駅までバスを利用)までの道路距離]	/*バス停徒歩時間から算出*/ "
                ''tmp_sql = tmp_sql & " 	,busho_time AS [バス停(駅までバスを利用)までの徒歩(分)]	/*設定値を変更*/ "
                ''tmp_sql = tmp_sql & " 	,1 AS [交通手段] "
                ''tmp_sql = tmp_sql & " 	,'' AS [車距離] "
                ''tmp_sql = tmp_sql & " FROM bk_eki "

                'tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	/**************************************************** "
                'tmp_sql = tmp_sql & " 	2016.04.06 "
                'tmp_sql = tmp_sql & " 	  物件交通情報の取得方法修正 "
                'tmp_sql = tmp_sql & " 	  設計書を元に修正 "
                'tmp_sql = tmp_sql & " 	****************************************************/ "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 [物件NO] "
                'tmp_sql = tmp_sql & " 		,ROW_NUMBER()OVER(PARTITION BY [物件NO] ORDER BY [内部No],[連番作成用]) AS [沿線番号] "
                'tmp_sql = tmp_sql & " 		/* "
                'tmp_sql = tmp_sql & " 		,[内部No] "
                'tmp_sql = tmp_sql & " 		,[連番作成用] "
                'tmp_sql = tmp_sql & " 		*/ "
                'tmp_sql = tmp_sql & " 		,[沿線No] "
                'tmp_sql = tmp_sql & " 		,[駅No] "
                'tmp_sql = tmp_sql & " 		,[始発フラグ] "
                'tmp_sql = tmp_sql & " 		,[距離] "
                'tmp_sql = tmp_sql & " 		,[徒歩] "
                'tmp_sql = tmp_sql & " 		,[車] "
                'tmp_sql = tmp_sql & " 		,[バス] "
                'tmp_sql = tmp_sql & " 		,[バス会社] "
                'tmp_sql = tmp_sql & " 		,[バス停] "
                'tmp_sql = tmp_sql & " 		,[バス停徒歩] "
                'tmp_sql = tmp_sql & " 		,[バス停までの距離] "
                'tmp_sql = tmp_sql & " 		,[設定する交通のタイプ] "
                'tmp_sql = tmp_sql & " 		,[バス系統・路線名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 		,[バス停名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 		,[バス停(駅までバスを利用)までの道路距離] "
                'tmp_sql = tmp_sql & " 		,[バス停(駅までバスを利用)までの徒歩(分)] "
                'tmp_sql = tmp_sql & " 		,[交通手段] "
                'tmp_sql = tmp_sql & " 		,[車距離] "
                'tmp_sql = tmp_sql & " 	FROM "
                'tmp_sql = tmp_sql & " 	( "
                'tmp_sql = tmp_sql & " 		/*********************************************** "
                'tmp_sql = tmp_sql & " 		①・交通タイプ：「最寄り駅」徒歩・バス・車を利用 "
                'tmp_sql = tmp_sql & " 		  ・徒歩で駅に行く場合 "
                'tmp_sql = tmp_sql & " 		***********************************************/ "
                'tmp_sql = tmp_sql & " 		SELECT "
                'tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 			,naibu_no AS [内部No] "
                'tmp_sql = tmp_sql & " 			,1 AS [連番作成用] "
                'tmp_sql = tmp_sql & " 			,ensen_no AS [沿線No] "
                'tmp_sql = tmp_sql & " 			,eki_no AS [駅No] "
                'tmp_sql = tmp_sql & " 			,0 AS [始発フラグ] "
                'tmp_sql = tmp_sql & " 			,ho_time * 80 AS [距離]	/*eki_kyoriはV7で連動用項目であることと、革命10では自動で80m/minを元に算出しているので時間から距離を算出して移行する*/ "
                'tmp_sql = tmp_sql & " 			,ho_time AS [徒歩] "
                'tmp_sql = tmp_sql & " 			,NULL AS [車] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス会社] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停徒歩] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停までの距離] "
                'tmp_sql = tmp_sql & " 			,1 AS [設定する交通のタイプ] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス系統・路線名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停(駅までバスを利用)までの道路距離]	/*バス停徒歩時間から算出*/ "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停(駅までバスを利用)までの徒歩(分)]	/*設定値を変更*/ "
                'tmp_sql = tmp_sql & " 			,1 AS [交通手段] "
                'tmp_sql = tmp_sql & " 			,NULL AS [車距離] "
                'tmp_sql = tmp_sql & " 		FROM bk_eki "
                'tmp_sql = tmp_sql & " 		/*********************************************** "
                'tmp_sql = tmp_sql & " 		②・交通タイプ：「最寄り駅」徒歩・バス・車を利用 "
                'tmp_sql = tmp_sql & " 		  ・バスを利用して駅に行く場合 "
                'tmp_sql = tmp_sql & " 		***********************************************/ "
                'tmp_sql = tmp_sql & " 		UNION "
                'tmp_sql = tmp_sql & " 		SELECT "
                'tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 			,naibu_no AS [内部No] "
                'tmp_sql = tmp_sql & " 			,2 AS [連番作成用] "
                'tmp_sql = tmp_sql & " 			,ensen_no AS [沿線No] "
                'tmp_sql = tmp_sql & " 			,eki_no AS [駅No] "
                'tmp_sql = tmp_sql & " 			,0 AS [始発フラグ] "
                'tmp_sql = tmp_sql & " 			,NULL AS [距離] "
                'tmp_sql = tmp_sql & " 			,NULL AS [徒歩] "
                'tmp_sql = tmp_sql & " 			,NULL AS [車] "
                'tmp_sql = tmp_sql & " 			,bus_time AS [バス] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス会社] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停徒歩] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停までの距離] "
                'tmp_sql = tmp_sql & " 			,1 AS [設定する交通のタイプ] "
                'tmp_sql = tmp_sql & " 			,'' AS [バス系統・路線名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 			,bus_name AS [バス停名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 			,busho_time * 80 AS [バス停(駅までバスを利用)までの道路距離]	/*バス停徒歩時間から算出*/ "
                'tmp_sql = tmp_sql & " 			,busho_time AS [バス停(駅までバスを利用)までの徒歩(分)]	/*設定値を変更*/ "
                'tmp_sql = tmp_sql & " 			,2 AS [交通手段] "
                'tmp_sql = tmp_sql & " 			,NULL AS [車距離] "
                'tmp_sql = tmp_sql & " 		FROM bk_eki "
                'tmp_sql = tmp_sql & " 		/*********************************************** "
                'tmp_sql = tmp_sql & " 		③・交通タイプ：「最寄り駅」徒歩・バス・車を利用 "
                'tmp_sql = tmp_sql & " 		  ・車で駅に行く場合 "
                'tmp_sql = tmp_sql & " 		***********************************************/ "
                'tmp_sql = tmp_sql & " 		UNION "
                'tmp_sql = tmp_sql & " 		SELECT "
                'tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 			,naibu_no AS [内部No] "
                'tmp_sql = tmp_sql & " 			,3 AS [連番作成用] "
                'tmp_sql = tmp_sql & " 			,ensen_no AS [沿線No] "
                'tmp_sql = tmp_sql & " 			,eki_no AS [駅No] "
                'tmp_sql = tmp_sql & " 			,0 AS [始発フラグ] "
                'tmp_sql = tmp_sql & " 			,NULL AS [距離] "
                'tmp_sql = tmp_sql & " 			,NULL AS [徒歩] "
                'tmp_sql = tmp_sql & " 			,car_time AS [車] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス会社] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停徒歩] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停までの距離] "
                'tmp_sql = tmp_sql & " 			,1 AS [設定する交通のタイプ] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス系統・路線名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 			,NULL * 80 AS [バス停(駅までバスを利用)までの道路距離]	/*バス停徒歩時間から算出*/ "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停(駅までバスを利用)までの徒歩(分)] "
                'tmp_sql = tmp_sql & " 			,3 AS [交通手段] "
                'tmp_sql = tmp_sql & " 			,'' AS [車距離] "
                'tmp_sql = tmp_sql & " 		FROM bk_eki "
                'tmp_sql = tmp_sql & " 		/*********************************************** "
                'tmp_sql = tmp_sql & " 		④交通タイプ：「最寄りバス停」徒歩を利用 "
                'tmp_sql = tmp_sql & " 		***********************************************/ "
                'tmp_sql = tmp_sql & " 		UNION "
                'tmp_sql = tmp_sql & " 		SELECT "
                'tmp_sql = tmp_sql & " 			 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 			,naibu_no AS [内部No] "
                'tmp_sql = tmp_sql & " 			,4 AS [連番作成用] "
                'tmp_sql = tmp_sql & " 			,NULL AS [沿線No] "
                'tmp_sql = tmp_sql & " 			,NULL AS [駅No] "
                'tmp_sql = tmp_sql & " 			,NULL AS [始発フラグ] "
                'tmp_sql = tmp_sql & " 			,NULL AS [距離] "
                'tmp_sql = tmp_sql & " 			,NULL AS [徒歩] "
                'tmp_sql = tmp_sql & " 			,NULL AS [車] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス] "
                'tmp_sql = tmp_sql & " 			,bus_rosen_name AS [バス会社] "
                'tmp_sql = tmp_sql & " 			,bus_name AS [バス停] "
                'tmp_sql = tmp_sql & " 			,busho_time AS [バス停徒歩] "
                'tmp_sql = tmp_sql & " 			,busho_time * 80 AS [バス停までの距離]		/*バス停徒歩時間から算出*/ "
                'tmp_sql = tmp_sql & " 			,0 AS [設定する交通のタイプ] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス系統・路線名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停名(駅までバスを利用)] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停(駅までバスを利用)までの道路距離] "
                'tmp_sql = tmp_sql & " 			,NULL AS [バス停(駅までバスを利用)までの徒歩(分)] "
                'tmp_sql = tmp_sql & " 			,NULL AS [交通手段] "
                'tmp_sql = tmp_sql & " 			,NULL AS [車距離] "
                'tmp_sql = tmp_sql & " 		FROM bk_eki "
                'tmp_sql = tmp_sql & " 	) AS TOTAL "
                'tmp_sql = tmp_sql & " ) AS VW "
                ''2016.04.06 物件交通情報の取得方法修正 -chg end

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[内部No],[沿線No],[駅No],[始発フラグ],[距離],[徒歩],[車],[バス],[バス会社],[バス停],[バス停徒歩],[バス停までの距離],[設定する交通のタイプ],[バス系統・路線名(駅までバスを利用)],[バス停名(駅までバスを利用)],[バス停(駅までバスを利用)までの道路距離],[バス停(駅までバスを利用)までの徒歩(分)],[車距離]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,naibu_no AS [内部No] "
                tmp_sql = tmp_sql & " 		,ensen_no AS [沿線No] "
                tmp_sql = tmp_sql & " 		,eki_no AS [駅No] "
                tmp_sql = tmp_sql & " 		,0 AS [始発フラグ] "
                tmp_sql = tmp_sql & " 		,ho_time * 80 AS [距離] "
                tmp_sql = tmp_sql & " 		,ho_time AS [徒歩] "
                tmp_sql = tmp_sql & " 		,car_time AS [車] "
                tmp_sql = tmp_sql & " 		,NULL AS [バス] "
                tmp_sql = tmp_sql & " 		,NULL AS [バス会社] "
                tmp_sql = tmp_sql & " 		,NULL AS [バス停] "
                tmp_sql = tmp_sql & " 		,NULL AS [バス停徒歩] "
                tmp_sql = tmp_sql & " 		,NULL AS [バス停までの距離] "
                tmp_sql = tmp_sql & " 		,1 AS [設定する交通のタイプ] "
                tmp_sql = tmp_sql & " 		,NULL AS [バス系統・路線名(駅までバスを利用)] "
                tmp_sql = tmp_sql & " 		,bus_name AS [バス停名(駅までバスを利用)] "
                tmp_sql = tmp_sql & " 		,busho_time * 80 AS [バス停(駅までバスを利用)までの道路距離] "
                tmp_sql = tmp_sql & " 		,busho_time AS [バス停(駅までバスを利用)までの徒歩(分)] "
                tmp_sql = tmp_sql & " 		,NULL AS [車距離] "
                tmp_sql = tmp_sql & " 	FROM bk_eki "
                tmp_sql = tmp_sql & " ) AS VW "
                '20160519 EXEUpdateに伴う修正 物件交通情報 -chg end

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

#Region "物件近隣駐車場情報(汎用ツールのみ)"

    'V7には存在しない項目
    'クエリだけ設けておく

    Public Class Bk_mst_kinrincyushajo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[駐車場No]"

                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 '' AS [物件NO] "
                tmp_sql = tmp_sql & " 	,'' AS [駐車場No] "
                tmp_sql = tmp_sql & " 	,'' AS [駐車場名] "
                tmp_sql = tmp_sql & " 	,'' AS [距離(m)] "
                tmp_sql = tmp_sql & " 	,'' AS [料金(月額)] "
                tmp_sql = tmp_sql & " 	,'' AS [税区分] "

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

#Region "物件参照ファイル情報(汎用ツールのみ)"

    'V7には存在しない項目
    'クエリだけ設けておく

    Public Class Bk_mst_sansyofile_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[ファイルNo]"

                tmp_sql = tmp_sql & " SELECT "
                tmp_sql = tmp_sql & " 	 '' AS [物件NO] "
                tmp_sql = tmp_sql & " 	,'' AS [ファイルNo] "
                tmp_sql = tmp_sql & " 	,'' AS [フルパス] "
                tmp_sql = tmp_sql & " 	,'' AS [追加時間] "
                tmp_sql = tmp_sql & " 	,'' AS [備考] "


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

#Region "物件接道情報"

    Public Class Bk_mst_setudo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[接道No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[接道No],[方角],[道路種],[幅員],[接道距離],[位置指定道路]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,[接道No] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN [方角] = '東' THEN  1 "
                tmp_sql = tmp_sql & " 			WHEN [方角] = '西' THEN  2 "
                tmp_sql = tmp_sql & " 			WHEN [方角] = '南' THEN  3 "
                tmp_sql = tmp_sql & " 			WHEN [方角] = '北' THEN  4 "
                tmp_sql = tmp_sql & " 			WHEN [方角] = '北東' THEN  5 "
                tmp_sql = tmp_sql & " 			WHEN [方角] = '南東' THEN  6 "
                tmp_sql = tmp_sql & " 			WHEN [方角] = '北西' THEN  7 "
                tmp_sql = tmp_sql & " 			WHEN [方角] = '南西' THEN  8 "
                tmp_sql = tmp_sql & " 			ELSE 1 "
                tmp_sql = tmp_sql & " 		 END AS [方角] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN [道路種] = '公道' THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN [道路種] = '私道' THEN 2 "
                tmp_sql = tmp_sql & " 			ELSE -1 "
                tmp_sql = tmp_sql & " 		 END AS [道路種] "
                tmp_sql = tmp_sql & " 		,[幅員] "
                tmp_sql = tmp_sql & " 		,[接道距離] "
                tmp_sql = tmp_sql & " 		,[位置指定道路] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT * FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no "
                tmp_sql = tmp_sql & " 				,1 AS [接道No] "
                tmp_sql = tmp_sql & " 				,setudo_hogaku1 AS [方角] "
                tmp_sql = tmp_sql & " 				,setudo_syu1 AS [道路種] "
                tmp_sql = tmp_sql & " 				,setudo_fukuin1 AS [幅員] "
                tmp_sql = tmp_sql & " 				,setudo_haba1 AS [接道距離] "
                tmp_sql = tmp_sql & " 				,0 AS [位置指定道路] "
                tmp_sql = tmp_sql & " 			FROM bk_mst "
                tmp_sql = tmp_sql & " 			UNION "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 				,2 AS [接道No] "
                tmp_sql = tmp_sql & " 				,setudo_hogaku2 AS [方角] "
                tmp_sql = tmp_sql & " 				,setudo_syu2 AS [道路種] "
                tmp_sql = tmp_sql & " 				,setudo_fukuin2 AS [幅員] "
                tmp_sql = tmp_sql & " 				,setudo_haba2 AS [接道距離] "
                tmp_sql = tmp_sql & " 				,0 AS [位置指定道路] "
                tmp_sql = tmp_sql & " 			FROM bk_mst "
                tmp_sql = tmp_sql & " 			UNION "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 				,3 AS [接道No] "
                tmp_sql = tmp_sql & " 				,setudo_hogaku3 AS [方角] "
                tmp_sql = tmp_sql & " 				,setudo_syu3 AS [道路種] "
                tmp_sql = tmp_sql & " 				,setudo_fukuin3 AS [幅員] "
                tmp_sql = tmp_sql & " 				,setudo_haba3 AS [接道距離] "
                tmp_sql = tmp_sql & " 				,0 AS [位置指定道路] "
                tmp_sql = tmp_sql & " 			FROM bk_mst "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		WHERE [方角] <> '' "
                tmp_sql = tmp_sql & " 	) AS BKSETUDO "
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

#Region "物件周辺情報"

    Public Class Bk_mst_syuhen_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[周辺No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[周辺No],[周辺施設分類],[周辺施設名],[周辺施設距離]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                '20160720 連動情報構築 -chg sta
                ''tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 		,1 AS [周辺No] "
                'tmp_sql = tmp_sql & " 		,19 AS [周辺施設分類] "
                'tmp_sql = tmp_sql & " 		,syogaku_name AS [周辺施設名] "
                'tmp_sql = tmp_sql & " 		,syogaku_kyori AS [周辺施設距離] "
                'tmp_sql = tmp_sql & " 	FROM bk_mst "
                'tmp_sql = tmp_sql & " 	UNION "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 		,2 AS [周辺No] "
                'tmp_sql = tmp_sql & " 		,20 AS [周辺施設分類] "
                'tmp_sql = tmp_sql & " 		,cyugaku_name AS [周辺施設名] "
                'tmp_sql = tmp_sql & " 		,chugaku_kyori AS [周辺施設距離] "
                'tmp_sql = tmp_sql & " 	FROM bk_mst "
                'tmp_sql = tmp_sql & " 	UNION "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 		,3 AS [周辺No] "
                'tmp_sql = tmp_sql & " 		,1 AS [周辺施設分類] "
                'tmp_sql = tmp_sql & " 		,'病院' AS [周辺施設名] "
                'tmp_sql = tmp_sql & " 		,ISNULL(as_byouin,0) AS [周辺施設距離] "
                'tmp_sql = tmp_sql & " 	FROM bk_mst "
                'tmp_sql = tmp_sql & " 	UNION "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 		,4 AS [周辺No] "
                'tmp_sql = tmp_sql & " 		,2 AS [周辺施設分類] "
                'tmp_sql = tmp_sql & " 		,'スーパー' AS [周辺施設名] "
                'tmp_sql = tmp_sql & " 		,ISNULL(as_surper,0) AS [周辺施設距離] "
                'tmp_sql = tmp_sql & " 	FROM bk_mst "
                'tmp_sql = tmp_sql & " 	UNION "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                'tmp_sql = tmp_sql & " 		,5 AS [周辺No] "
                'tmp_sql = tmp_sql & " 		,3 AS [周辺施設分類] "
                'tmp_sql = tmp_sql & " 		,'コンビニ' AS [周辺施設名] "
                'tmp_sql = tmp_sql & " 		,ISNULL(as_konbini,0) AS [周辺施設距離] "
                'tmp_sql = tmp_sql & " 	FROM bk_mst "
                'tmp_sql = tmp_sql & " ) AS BKSYUHEN "
                'tmp_sql = tmp_sql & " WHERE [周辺No] IN (1,2) "
                'tmp_sql = tmp_sql & " OR    ([周辺No] IN (3,4,5) AND [周辺施設距離] <> 0) "
                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT * FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 [物件NO] "
                tmp_sql = tmp_sql & " 			,ROW_NUMBER()OVER(PARTITION BY [物件NO] ORDER BY [周辺No]) AS [周辺No] "
                tmp_sql = tmp_sql & " 			,[周辺施設分類] "
                tmp_sql = tmp_sql & " 			,[周辺施設名] "
                tmp_sql = tmp_sql & " 			,CASE "
                tmp_sql = tmp_sql & " 				WHEN [周辺施設距離] = 0 THEN NULL "
                tmp_sql = tmp_sql & " 				ELSE [周辺施設距離] "
                tmp_sql = tmp_sql & " 			 END AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 		FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			/*物件周辺情報取得 sta*/ "
                tmp_sql = tmp_sql & " 			SELECT * FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 						,1 AS [周辺No] "
                tmp_sql = tmp_sql & " 						,19 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 						,syogaku_name AS [周辺施設名] "
                tmp_sql = tmp_sql & " 						,syogaku_kyori AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 					FROM bk_mst "
                tmp_sql = tmp_sql & " 					UNION "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 						,2 AS [周辺No] "
                tmp_sql = tmp_sql & " 						,20 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 						,cyugaku_name AS [周辺施設名] "
                tmp_sql = tmp_sql & " 						,chugaku_kyori AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 					FROM bk_mst "
                tmp_sql = tmp_sql & " 					UNION "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 						,3 AS [周辺No] "
                tmp_sql = tmp_sql & " 						,1 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 						,'病院' AS [周辺施設名] "
                tmp_sql = tmp_sql & " 						,ISNULL(as_byouin,0) AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 					FROM bk_mst "
                tmp_sql = tmp_sql & " 					UNION "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 						,4 AS [周辺No] "
                tmp_sql = tmp_sql & " 						,2 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 						,'スーパー' AS [周辺施設名] "
                tmp_sql = tmp_sql & " 						,ISNULL(as_surper,0) AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 					FROM bk_mst "
                tmp_sql = tmp_sql & " 					UNION "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 						,5 AS [周辺No] "
                tmp_sql = tmp_sql & " 						,3 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 						,'コンビニ' AS [周辺施設名] "
                tmp_sql = tmp_sql & " 						,ISNULL(as_konbini,0) AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 					FROM bk_mst "
                tmp_sql = tmp_sql & " 				) AS BKSYUHEN "
                tmp_sql = tmp_sql & " 				WHERE [周辺No] IN (1,2) "
                tmp_sql = tmp_sql & " 				OR    ([周辺No] IN (3,4,5) AND [周辺施設距離] <> 0) "
                tmp_sql = tmp_sql & " 			) AS BKSYUHENTOTAL "
                tmp_sql = tmp_sql & " 			/*物件周辺情報取得 end*/ "
                tmp_sql = tmp_sql & " 			UNION "
                tmp_sql = tmp_sql & " 			/*部屋周辺情報取得 sta*/ "
                tmp_sql = tmp_sql & " 			SELECT * FROM "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT "
                tmp_sql = tmp_sql & " 					 bk_no AS [物件No] "
                tmp_sql = tmp_sql & " 					,ROW_NUMBER()OVER(PARTITION BY bk_no ORDER BY hy_no,syuhen_no) + 5 AS [周辺No] "
                tmp_sql = tmp_sql & " 					,syuhen_rui AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 					,d2_item_name AS [周辺施設名] "
                tmp_sql = tmp_sql & " 					,d2_item_kyori AS [周辺施設距離]	 "
                tmp_sql = tmp_sql & " 				FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT "
                tmp_sql = tmp_sql & " 						 HYSYUHENTOTAL.* "
                tmp_sql = tmp_sql & " 						/*,kosin_date*/ "
                tmp_sql = tmp_sql & " 						,ROW_NUMBER()OVER(PARTITION BY syuhen_rui ORDER BY kosin_date DESC,d2_item_name ASC,d2_item_kyori DESC) AS [同種別での最新データ抽出用] "
                tmp_sql = tmp_sql & " 					FROM "
                tmp_sql = tmp_sql & " 					( "
                tmp_sql = tmp_sql & " 						SELECT "
                tmp_sql = tmp_sql & " 							 bk_no "
                tmp_sql = tmp_sql & " 							,hy_no "
                tmp_sql = tmp_sql & " 							,syuhen_no "
                tmp_sql = tmp_sql & " 							,CASE "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0010' THEN 4 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0020' THEN 3 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0030' THEN 2 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0040' THEN 5 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0050' THEN 6 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0060' THEN 17 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0070' THEN 15 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0080' THEN 30 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0090' THEN 15 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0100' THEN 16 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0110' THEN 16 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0120' THEN 7 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0130' THEN 20 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0140' THEN 19 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0150' THEN 8 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0160' THEN 1 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0170' THEN 1 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0180' THEN 9 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0190' THEN 12 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0210' THEN 14 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0220' THEN 10 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0230' THEN 13 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0240' THEN 11 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0250' THEN 18 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0260' THEN 15 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '0270' THEN 15 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '280' THEN 21 "
                tmp_sql = tmp_sql & " 								WHEN d2_item_rui = '9999' THEN 15 "
                tmp_sql = tmp_sql & " 							 END AS syuhen_rui "
                tmp_sql = tmp_sql & " 							,d2_item_name "
                tmp_sql = tmp_sql & " 							,d2_item_kyori "
                tmp_sql = tmp_sql & " 						FROM "
                tmp_sql = tmp_sql & " 						( "
                tmp_sql = tmp_sql & " 							/*部屋周辺情報から周辺データを取得 sta*/ "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT "
                tmp_sql = tmp_sql & " 									 bk_no "
                tmp_sql = tmp_sql & " 									,hy_no "
                tmp_sql = tmp_sql & " 									,1 AS syuhen_no "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 101 THEN d2_item ELSE '' END) AS d2_item_rui "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 102 THEN d2_item ELSE '' END) AS d2_item_name "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 103 THEN d2_item ELSE '' END) AS d2_item_kyori "
                tmp_sql = tmp_sql & " 								FROM hy_d2 AS VW "
                tmp_sql = tmp_sql & " 								WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 								GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 							) AS HYSYUHEN1 "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT "
                tmp_sql = tmp_sql & " 									 bk_no "
                tmp_sql = tmp_sql & " 									,hy_no "
                tmp_sql = tmp_sql & " 									,2 AS syuhen_no "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 201 THEN d2_item ELSE '' END) AS d2_item_rui "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 202 THEN d2_item ELSE '' END) AS d2_item_name "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 203 THEN d2_item ELSE '' END) AS d2_item_kyori "
                tmp_sql = tmp_sql & " 								FROM hy_d2 AS VW "
                tmp_sql = tmp_sql & " 								WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 								GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 							) AS HYSYUHEN2 "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT "
                tmp_sql = tmp_sql & " 									 bk_no "
                tmp_sql = tmp_sql & " 									,hy_no "
                tmp_sql = tmp_sql & " 									,3 AS syuhen_no "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 301 THEN d2_item ELSE '' END) AS d2_item_rui "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 302 THEN d2_item ELSE '' END) AS d2_item_name "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 303 THEN d2_item ELSE '' END) AS d2_item_kyori "
                tmp_sql = tmp_sql & " 								FROM hy_d2 AS VW "
                tmp_sql = tmp_sql & " 								WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 								GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 							) AS HYSYUHEN3 "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT "
                tmp_sql = tmp_sql & " 									 bk_no "
                tmp_sql = tmp_sql & " 									,hy_no "
                tmp_sql = tmp_sql & " 									,4 AS syuhen_no "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 401 THEN d2_item ELSE '' END) AS d2_item_rui "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 402 THEN d2_item ELSE '' END) AS d2_item_name "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 403 THEN d2_item ELSE '' END) AS d2_item_kyori "
                tmp_sql = tmp_sql & " 								FROM hy_d2 AS VW "
                tmp_sql = tmp_sql & " 								WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 								GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 							) AS HYSYUHEN4 "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT "
                tmp_sql = tmp_sql & " 									 bk_no "
                tmp_sql = tmp_sql & " 									,hy_no "
                tmp_sql = tmp_sql & " 									,5 AS syuhen_no "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 501 THEN d2_item ELSE '' END) AS d2_item_rui "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 502 THEN d2_item ELSE '' END) AS d2_item_name "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 503 THEN d2_item ELSE '' END) AS d2_item_kyori "
                tmp_sql = tmp_sql & " 								FROM hy_d2 AS VW "
                tmp_sql = tmp_sql & " 								WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 								GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 							) AS HYSYUHEN5 "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT "
                tmp_sql = tmp_sql & " 									 bk_no "
                tmp_sql = tmp_sql & " 									,hy_no "
                tmp_sql = tmp_sql & " 									,6 AS syuhen_no "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 601 THEN d2_item ELSE '' END) AS d2_item_rui "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 602 THEN d2_item ELSE '' END) AS d2_item_name "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 603 THEN d2_item ELSE '' END) AS d2_item_kyori "
                tmp_sql = tmp_sql & " 								FROM hy_d2 AS VW "
                tmp_sql = tmp_sql & " 								WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 								GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 							) AS HYSYUHEN6 "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT "
                tmp_sql = tmp_sql & " 									 bk_no "
                tmp_sql = tmp_sql & " 									,hy_no "
                tmp_sql = tmp_sql & " 									,7 AS syuhen_no "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 701 THEN d2_item ELSE '' END) AS d2_item_rui "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 702 THEN d2_item ELSE '' END) AS d2_item_name "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 703 THEN d2_item ELSE '' END) AS d2_item_kyori "
                tmp_sql = tmp_sql & " 								FROM hy_d2 AS VW "
                tmp_sql = tmp_sql & " 								WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 								GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 							) AS HYSYUHEN7 "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT * FROM "
                tmp_sql = tmp_sql & " 							( "
                tmp_sql = tmp_sql & " 								SELECT "
                tmp_sql = tmp_sql & " 									 bk_no "
                tmp_sql = tmp_sql & " 									,hy_no "
                tmp_sql = tmp_sql & " 									,8 AS syuhen_no "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 801 THEN d2_item ELSE '' END) AS d2_item_rui "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 802 THEN d2_item ELSE '' END) AS d2_item_name "
                tmp_sql = tmp_sql & " 									,MAX(CASE WHEN d2_data_id = 803 THEN d2_item ELSE '' END) AS d2_item_kyori "
                tmp_sql = tmp_sql & " 								FROM hy_d2 AS VW "
                tmp_sql = tmp_sql & " 								WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 								GROUP BY bk_no,hy_no "
                tmp_sql = tmp_sql & " 							) AS HYSYUHEN8 "
                tmp_sql = tmp_sql & " 							/*部屋周辺情報から周辺データを取得 end*/ "
                tmp_sql = tmp_sql & " 						) AS VW "
                tmp_sql = tmp_sql & " 						WHERE ISNULL(d2_item_rui,'') <> '' "
                tmp_sql = tmp_sql & " 					) AS HYSYUHENTOTAL "
                tmp_sql = tmp_sql & " 					LEFT JOIN "
                tmp_sql = tmp_sql & " 						( "
                tmp_sql = tmp_sql & " 							/*最新データ取得用に更新日を含めたVIEWを作成 sta*/ "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								 bk_no "
                tmp_sql = tmp_sql & " 								,hy_no "
                tmp_sql = tmp_sql & " 								,CONVERT(INT,LEFT(CONVERT(VARCHAR,d2_data_id),1)) AS syuhen_no "
                tmp_sql = tmp_sql & " 								,d2_item "
                tmp_sql = tmp_sql & " 								,kosin_date "
                tmp_sql = tmp_sql & " 							FROM hy_d2 "
                tmp_sql = tmp_sql & " 							WHERE d2_data_kbn = 1 "
                tmp_sql = tmp_sql & " 							AND   RIGHT(CONVERT(VARCHAR,d2_data_id),2) = '01' "
                tmp_sql = tmp_sql & " 							/*最新データ取得用に更新日を含めたVIEWを作成 end*/ "
                tmp_sql = tmp_sql & " 						) AS VW "
                tmp_sql = tmp_sql & " 					ON  HYSYUHENTOTAL.bk_no = VW.bk_no "
                tmp_sql = tmp_sql & " 					AND HYSYUHENTOTAL.hy_no = VW.hy_no "
                tmp_sql = tmp_sql & " 					AND HYSYUHENTOTAL.syuhen_no = VW.syuhen_no "
                tmp_sql = tmp_sql & " 					AND HYSYUHENTOTAL.syuhen_rui = VW.d2_item "
                tmp_sql = tmp_sql & " 				) AS TOTAL "
                tmp_sql = tmp_sql & " 				WHERE [同種別での最新データ抽出用] = 1 "
                tmp_sql = tmp_sql & " 			) AS HYSYUHENTOTAL "
                tmp_sql = tmp_sql & " 			/*物件情報の5つの周辺情報で距離が設定されている場合は物件情報を優先させるため除外する sta*/ "
                tmp_sql = tmp_sql & " 			WHERE NOT EXISTS "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT * FROM "
                tmp_sql = tmp_sql & " 					( "
                tmp_sql = tmp_sql & " 						SELECT * FROM "
                tmp_sql = tmp_sql & " 						( "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 								,1 AS [周辺No] "
                tmp_sql = tmp_sql & " 								,19 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 								,syogaku_name AS [周辺施設名] "
                tmp_sql = tmp_sql & " 								,syogaku_kyori AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 							FROM bk_mst "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 								,2 AS [周辺No] "
                tmp_sql = tmp_sql & " 								,20 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 								,cyugaku_name AS [周辺施設名] "
                tmp_sql = tmp_sql & " 								,chugaku_kyori AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 							FROM bk_mst "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 								,3 AS [周辺No] "
                tmp_sql = tmp_sql & " 								,1 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 								,'病院' AS [周辺施設名] "
                tmp_sql = tmp_sql & " 								,ISNULL(as_byouin,0) AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 							FROM bk_mst "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 								,4 AS [周辺No] "
                tmp_sql = tmp_sql & " 								,2 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 								,'スーパー' AS [周辺施設名] "
                tmp_sql = tmp_sql & " 								,ISNULL(as_surper,0) AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 							FROM bk_mst "
                tmp_sql = tmp_sql & " 							UNION "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 								,5 AS [周辺No] "
                tmp_sql = tmp_sql & " 								,3 AS [周辺施設分類] "
                tmp_sql = tmp_sql & " 								,'コンビニ' AS [周辺施設名] "
                tmp_sql = tmp_sql & " 								,ISNULL(as_konbini,0) AS [周辺施設距離] "
                tmp_sql = tmp_sql & " 							FROM bk_mst "
                tmp_sql = tmp_sql & " 						) AS BKSYUHEN "
                tmp_sql = tmp_sql & " 						WHERE [周辺施設距離] <> 0 "
                tmp_sql = tmp_sql & " 					) AS BKSYUHENTOTAL "
                tmp_sql = tmp_sql & " 					WHERE HYSYUHENTOTAL.物件No = BKSYUHENTOTAL.物件NO "
                tmp_sql = tmp_sql & " 					AND   HYSYUHENTOTAL.周辺施設分類 = BKSYUHENTOTAL.周辺施設分類 "
                tmp_sql = tmp_sql & " 				) "
                tmp_sql = tmp_sql & " 			/*物件情報の5つの周辺情報で距離が設定されている場合は物件情報を優先させるため除外する end*/ "
                tmp_sql = tmp_sql & " 			/*部屋周辺情報取得 end*/ "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 	) AS SYUHENTOTAL "
                tmp_sql = tmp_sql & " 	WHERE 周辺No <= 10 "
                tmp_sql = tmp_sql & " ) AS VW "
                '20160720 連動情報構築 -chg end
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

#Region "物件修繕維持管理連絡先情報"

    Public Class Bk_setubi_ren_bk_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO],[修繕及び維持管理No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[修繕及び維持管理No],[修繕及び維持管理の箇所],[対象区分],[業者no],[氏名(商号または名称)(SJIS)],[住所(主たる事務所の所在地)],[連絡先電話番号],[自社no],[貸主No],[所有者No],[氏名(商号または名称)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT TOP 1 "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 		,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 		,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 		,'' AS [業者no] "
                tmp_sql = tmp_sql & " 		,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 		,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 		,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 		,[自社・支店No] AS [自社no] "
                tmp_sql = tmp_sql & " 		,'' AS [貸主No] "
                tmp_sql = tmp_sql & " 		,'' AS [所有者No] "
                tmp_sql = tmp_sql & " 		,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT BKREN.*,JISYATOTAL.[自社・支店No] FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 			WHERE copy_kbn = 1 "
                tmp_sql = tmp_sql & " 			AND   kbn = 1 "
                tmp_sql = tmp_sql & " 		) AS BKREN "
                tmp_sql = tmp_sql & " 		LEFT JOIN "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT "
                tmp_sql = tmp_sql & " 					 jisya_no AS [自社・支店No] "
                tmp_sql = tmp_sql & " 					,kaisya_name AS [自社・支店名] "
                tmp_sql = tmp_sql & " 					,'' AS [自社・支店名（SJIS）] "
                tmp_sql = tmp_sql & " 					,kaisya_kana AS [自社・支店カナ] "
                tmp_sql = tmp_sql & " 					,'' AS [自社・支店名２] "
                tmp_sql = tmp_sql & " 					,'' AS [自社・支店名２（SJIS）] "
                tmp_sql = tmp_sql & " 					,'' AS [自社・支店名３] "
                tmp_sql = tmp_sql & " 					,'' AS [自社・支店名３（SJIS）] "
                tmp_sql = tmp_sql & " 					,yakusyoku AS [代表者役職] "
                tmp_sql = tmp_sql & " 					,daihyo AS [代表者名] "
                tmp_sql = tmp_sql & " 					,kaisya_post AS [郵便番号] "
                tmp_sql = tmp_sql & " 					,kaisya_add1 AS [住所1] "
                tmp_sql = tmp_sql & " 					,kaisya_add2 AS [住所2] "
                tmp_sql = tmp_sql & " 					,kaisya_tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 					,kaisya_tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 					,'' AS [携帯TEL1] "
                tmp_sql = tmp_sql & " 					,'' AS [携帯TEL2] "
                tmp_sql = tmp_sql & " 					,kaisya_fax AS [FAX] "
                tmp_sql = tmp_sql & " 					,kaisya_email AS [メールアドレス] "
                tmp_sql = tmp_sql & " 					,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 					,kaisya_url AS [Webアドレス] "
                tmp_sql = tmp_sql & " 					,'' AS [備考（基本情報）] "
                tmp_sql = tmp_sql & " 					,menkyo_no AS [免許証番号] "
                tmp_sql = tmp_sql & " 					,menkyo_ymd AS [免許年月日] "
                tmp_sql = tmp_sql & " 					,takuken_no AS [取引主任者登録番号（代表）] "
                tmp_sql = tmp_sql & " 					,takuken_name AS [取引主任者名（代表）] "
                tmp_sql = tmp_sql & " 					,'' AS [全国賃貸不動産管理業協会会員番号] "
                tmp_sql = tmp_sql & " 					,'' AS [賃貸不動産経営管理士登録番号] "
                tmp_sql = tmp_sql & " 					,'' AS [賃貸不動産経営管理士名] "
                tmp_sql = tmp_sql & " 					,'' AS [賃貸住宅管理業者登録番号] "
                tmp_sql = tmp_sql & " 					,dantai1 AS [加入団体１] "
                tmp_sql = tmp_sql & " 					,dantai2 AS [加入団体２] "
                tmp_sql = tmp_sql & " 					,dantai3 AS [加入団体３] "
                tmp_sql = tmp_sql & " 					,dantai4 AS [加入団体４] "
                tmp_sql = tmp_sql & " 					,dantai5 AS [加入団体５] "
                tmp_sql = tmp_sql & " 					,dantai6 AS [加入団体６] "
                tmp_sql = tmp_sql & " 					,dantai7 AS [加入団体７] "
                tmp_sql = tmp_sql & " 					,dantai8 AS [加入団体８] "
                tmp_sql = tmp_sql & " 					,dantai9 AS [加入団体９] "
                tmp_sql = tmp_sql & " 					,dantai10 AS [加入団体１０] "
                tmp_sql = tmp_sql & " 					,'' AS [代表者名（SJIS）] "
                tmp_sql = tmp_sql & " 					,'' AS [取引主任者名（代表）（SJIS）] "
                tmp_sql = tmp_sql & " 					,'' AS [賃貸不動産経営管理士名（SJIS）] "
                tmp_sql = tmp_sql & " 				FROM "
                tmp_sql = tmp_sql & " 				( "
                tmp_sql = tmp_sql & " 					SELECT * FROM "
                tmp_sql = tmp_sql & " 					( "
                tmp_sql = tmp_sql & " 						SELECT "
                tmp_sql = tmp_sql & " 							 [jisya_no] "
                tmp_sql = tmp_sql & " 							,[kaisya_name] "
                tmp_sql = tmp_sql & " 							,[kaisya_kana] "
                tmp_sql = tmp_sql & " 							,[yakusyoku] "
                tmp_sql = tmp_sql & " 							,[daihyo] "
                tmp_sql = tmp_sql & " 							,[kaisya_post] "
                tmp_sql = tmp_sql & " 							,[kaisya_add1] "
                tmp_sql = tmp_sql & " 							,[kaisya_add2] "
                tmp_sql = tmp_sql & " 							,[kaisya_tel1] "
                tmp_sql = tmp_sql & " 							,[kaisya_tel2] "
                tmp_sql = tmp_sql & " 							,[kaisya_fax] "
                tmp_sql = tmp_sql & " 							,[kaisya_email] "
                tmp_sql = tmp_sql & " 							,[kaisya_url] "
                tmp_sql = tmp_sql & " 							,[menkyo_no] "
                tmp_sql = tmp_sql & " 							,[menkyo_ymd] "
                tmp_sql = tmp_sql & " 							,[takuken_no] "
                tmp_sql = tmp_sql & " 							,[takuken_name] "
                tmp_sql = tmp_sql & " 							,[dantai1] "
                tmp_sql = tmp_sql & " 							,[dantai2] "
                tmp_sql = tmp_sql & " 							,[dantai3] "
                tmp_sql = tmp_sql & " 							,[dantai4] "
                tmp_sql = tmp_sql & " 							,[dantai5] "
                tmp_sql = tmp_sql & " 							,[dantai6] "
                tmp_sql = tmp_sql & " 							,[dantai7] "
                tmp_sql = tmp_sql & " 							,[dantai8] "
                tmp_sql = tmp_sql & " 							,[dantai9] "
                tmp_sql = tmp_sql & " 							,[dantai10] "
                tmp_sql = tmp_sql & " 						FROM m_jisya "
                tmp_sql = tmp_sql & " 					) AS JISYA "
                tmp_sql = tmp_sql & " 					UNION "
                tmp_sql = tmp_sql & " 					SELECT * FROM "
                tmp_sql = tmp_sql & " 					( "
                tmp_sql = tmp_sql & " 						SELECT "
                tmp_sql = tmp_sql & " 							 ROW_NUMBER()OVER(ORDER BY [siten_no]) + "
                tmp_sql = tmp_sql & " 								( "
                tmp_sql = tmp_sql & " 									SELECT jisya_no FROM "
                tmp_sql = tmp_sql & " 									( "
                tmp_sql = tmp_sql & " 										SELECT "
                tmp_sql = tmp_sql & " 											ROW_NUMBER()OVER(ORDER BY jisya_no DESC) AS 抽出用 "
                tmp_sql = tmp_sql & " 											,* "
                tmp_sql = tmp_sql & " 										FROM m_jisya "
                tmp_sql = tmp_sql & " 									) AS VW "
                tmp_sql = tmp_sql & " 									WHERE 抽出用 = 1 "
                tmp_sql = tmp_sql & " 								) AS [移行用自社№] "
                tmp_sql = tmp_sql & " 							 ,[siten_name] "
                tmp_sql = tmp_sql & " 							,'' AS [カナ] "
                tmp_sql = tmp_sql & " 							,'' AS [代表者役職] "
                tmp_sql = tmp_sql & " 							,'' AS [代表者名称] "
                tmp_sql = tmp_sql & " 							,[post] "
                tmp_sql = tmp_sql & " 							,[add1] "
                tmp_sql = tmp_sql & " 							,[add2] "
                tmp_sql = tmp_sql & " 							,[tel1] "
                tmp_sql = tmp_sql & " 							,[tel2] "
                tmp_sql = tmp_sql & " 							,[fax] "
                tmp_sql = tmp_sql & " 							,[email] "
                tmp_sql = tmp_sql & " 							,'' AS [url] "
                tmp_sql = tmp_sql & " 							,'' AS [免許証番号] "
                tmp_sql = tmp_sql & " 							,'' AS [免許証年月日] "
                tmp_sql = tmp_sql & " 							,'' AS [宅建取引士登録番号] "
                tmp_sql = tmp_sql & " 							,'' AS [宅建取引士氏名] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体1] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体2] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体3] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体4] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体5] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体6] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体7] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体8] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体9] "
                tmp_sql = tmp_sql & " 							,'' AS [加入団体10] "
                tmp_sql = tmp_sql & " 						FROM m_siten "
                tmp_sql = tmp_sql & " 					) AS SITEN "
                tmp_sql = tmp_sql & " 				) AS JISYAMST "
                tmp_sql = tmp_sql & " 			) AS JISYATOTAL "
                tmp_sql = tmp_sql & " 		ON BKREN.[ren_name] = JISYATOTAL.[自社・支店名] "
                tmp_sql = tmp_sql & " 	) AS BKSZENJISYA "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT TOP 1 "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 		,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 		,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 		,'' AS [業者no] "
                tmp_sql = tmp_sql & " 		,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 		,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 		,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 		,'' AS [自社no] "
                tmp_sql = tmp_sql & " 		,so_no AS [貸主No] "
                tmp_sql = tmp_sql & " 		,'' AS [所有者No] "
                tmp_sql = tmp_sql & " 		,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT VW.*,KASIMST.so_no FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 			WHERE copy_kbn = 2 "
                tmp_sql = tmp_sql & " 			AND   kbn = 1 "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		LEFT JOIN m_yanu_so AS KASIMST "
                tmp_sql = tmp_sql & " 		ON VW.[ren_name] = KASIMST.so_name "
                tmp_sql = tmp_sql & " 	) AS BKSZENKASI "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT TOP 1 "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 		,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 		,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 		,'' AS [業者no] "
                tmp_sql = tmp_sql & " 		,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 		,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 		,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 		,'' AS [自社no] "
                tmp_sql = tmp_sql & " 		,'' AS [貸主No] "
                tmp_sql = tmp_sql & " 		,so_no AS [所有者No] "
                tmp_sql = tmp_sql & " 		,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT VW.*,KASIMST.so_no FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 			WHERE copy_kbn = 3 "
                tmp_sql = tmp_sql & " 			AND   kbn = 1 "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		LEFT JOIN m_yanu_so AS KASIMST "
                tmp_sql = tmp_sql & " 		ON VW.[ren_name] = KASIMST.so_name "
                tmp_sql = tmp_sql & " 	) AS BKSZENSYO "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT TOP 1 "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 		,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 		,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 		,gy_no AS [業者no] "
                tmp_sql = tmp_sql & " 		,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 		,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 		,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 		,'' AS [自社no] "
                tmp_sql = tmp_sql & " 		,'' AS [貸主No] "
                tmp_sql = tmp_sql & " 		,'' AS [所有者No] "
                tmp_sql = tmp_sql & " 		,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT VW.*,GYMST.gy_no FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 			WHERE copy_kbn = 4 "
                tmp_sql = tmp_sql & " 			AND   kbn = 1 "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		LEFT JOIN m_gy AS GYMST "
                tmp_sql = tmp_sql & " 		ON VW.[ren_name] = GYMST.gy_name "
                tmp_sql = tmp_sql & " 	) AS BKSZENKASI "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	SELECT TOP 1 "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,naibu_no AS [修繕及び維持管理No] "
                tmp_sql = tmp_sql & " 		,kmk_name AS [修繕及び維持管理の箇所] "
                tmp_sql = tmp_sql & " 		,copy_kbn AS [対象区分] "
                tmp_sql = tmp_sql & " 		,gy_no AS [業者no] "
                tmp_sql = tmp_sql & " 		,'' AS [氏名(商号または名称)(SJIS)] "
                tmp_sql = tmp_sql & " 		,ren_add AS [住所(主たる事務所の所在地)] "
                tmp_sql = tmp_sql & " 		,ren_tel AS [連絡先電話番号] "
                tmp_sql = tmp_sql & " 		,'' AS [自社no] "
                tmp_sql = tmp_sql & " 		,'' AS [貸主No] "
                tmp_sql = tmp_sql & " 		,'' AS [所有者No] "
                tmp_sql = tmp_sql & " 		,ren_name AS [氏名(商号または名称)] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT VW.*,GYSYUZENMST.gy_no FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT * FROM bk_setubi_ren "
                tmp_sql = tmp_sql & " 			WHERE copy_kbn = 5 "
                tmp_sql = tmp_sql & " 			AND   kbn = 1 "
                tmp_sql = tmp_sql & " 		) AS VW "
                tmp_sql = tmp_sql & " 		LEFT JOIN m_gysyuzen AS GYSYUZENMST "
                tmp_sql = tmp_sql & " 		ON VW.[ren_name] = GYSYUZENMST.gy_name "
                tmp_sql = tmp_sql & " 	) AS BKSZENKASI "
                tmp_sql = tmp_sql & " ) AS SZENTOTAL "

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

#Region "物件メモ情報"

    Public Class Bk_biko_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[物件NO]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[物件NO],[備考1],[備考2],[備考3],[備考4],[備考5],[備考6],[備考7],[備考8],[備考9],[備考10],[備考11],[備考12],[備考13],[備考14],[備考15],[備考16],[備考17],[備考18],[備考19],[備考20],[備考21],[備考22],[備考23],[備考24],[備考25],[備考26],[備考27],[備考28],[備考29],[備考30],[備考31],[備考32],[備考33],[備考34],[備考35],[備考36],[備考37],[備考38],[備考39],[備考40],[備考41],[備考42],[備考43],[備考44],[備考45],[備考46],[備考47],[備考48],[備考49],[備考50]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 1 THEN biko ELSE '' END) AS 備考1 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 2 THEN biko ELSE '' END) AS 備考2 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 3 THEN biko ELSE '' END) AS 備考3 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 4 THEN biko ELSE '' END) AS 備考4 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 5 THEN biko ELSE '' END) AS 備考5 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 6 THEN biko ELSE '' END) AS 備考6 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 7 THEN biko ELSE '' END) AS 備考7 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 8 THEN biko ELSE '' END) AS 備考8 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 9 THEN biko ELSE '' END) AS 備考9 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 10 THEN biko ELSE '' END) AS 備考10 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 11 THEN biko ELSE '' END) AS 備考11 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 12 THEN biko ELSE '' END) AS 備考12 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 13 THEN biko ELSE '' END) AS 備考13 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 14 THEN biko ELSE '' END) AS 備考14 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 15 THEN biko ELSE '' END) AS 備考15 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 16 THEN biko ELSE '' END) AS 備考16 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 17 THEN biko ELSE '' END) AS 備考17 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 18 THEN biko ELSE '' END) AS 備考18 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 19 THEN biko ELSE '' END) AS 備考19 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 20 THEN biko ELSE '' END) AS 備考20 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 21 THEN biko ELSE '' END) AS 備考21 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 22 THEN biko ELSE '' END) AS 備考22 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 23 THEN biko ELSE '' END) AS 備考23 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 24 THEN biko ELSE '' END) AS 備考24 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 25 THEN biko ELSE '' END) AS 備考25 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 26 THEN biko ELSE '' END) AS 備考26 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 27 THEN biko ELSE '' END) AS 備考27 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 28 THEN biko ELSE '' END) AS 備考28 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 29 THEN biko ELSE '' END) AS 備考29 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 30 THEN biko ELSE '' END) AS 備考30 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 31 THEN biko ELSE '' END) AS 備考31 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 32 THEN biko ELSE '' END) AS 備考32 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 33 THEN biko ELSE '' END) AS 備考33 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 34 THEN biko ELSE '' END) AS 備考34 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 35 THEN biko ELSE '' END) AS 備考35 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 36 THEN biko ELSE '' END) AS 備考36 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 37 THEN biko ELSE '' END) AS 備考37 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 38 THEN biko ELSE '' END) AS 備考38 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 39 THEN biko ELSE '' END) AS 備考39 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 40 THEN biko ELSE '' END) AS 備考40 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 41 THEN biko ELSE '' END) AS 備考41 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 42 THEN biko ELSE '' END) AS 備考42 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 43 THEN biko ELSE '' END) AS 備考43 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 44 THEN biko ELSE '' END) AS 備考44 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 45 THEN biko ELSE '' END) AS 備考45 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 46 THEN biko ELSE '' END) AS 備考46 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 47 THEN biko ELSE '' END) AS 備考47 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 48 THEN biko ELSE '' END) AS 備考48 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 49 THEN biko ELSE '' END) AS 備考49 "
                tmp_sql = tmp_sql & " 		,MAX(CASE WHEN biko_no = 50 THEN biko ELSE '' END) AS 備考50 "
                tmp_sql = tmp_sql & " 	FROM bk_biko "
                tmp_sql = tmp_sql & " 	GROUP BY bk_no "
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

#Region "×物件外部No情報"

    '使用箇所が不明のため保留

#End Region

#Region "×物件画像情報"

    '画像は別途行うためここではコンバートしない

#End Region

#Region "×物件周辺画像情報"

    '画像は別途行うためここではコンバートしない

#End Region

#Region "×物件KeyValueテーブル"

    '使用箇所が不明のため保留

#End Region

#Region "×物件bkdata_youtube情報"

    '使用箇所が不明のため保留

#End Region

End Namespace
