Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "自社情報"

    Public Class M_jisya_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[自社・支店No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[自社・支店No],[自社・支店名],[自社・支店名（SJIS）],[自社・支店カナ],[自社・支店名２],[自社・支店名２（SJIS）],[自社・支店名３],[自社・支店名３（SJIS）],[代表者役職],[代表者名],[郵便番号],[住所1],[住所2],[TEL1],[TEL2],[携帯TEL1],[携帯TEL2],[FAX],[メールアドレス],[携帯メールアドレス],[Webアドレス],[備考（基本情報）],[免許証番号],[免許年月日],[取引主任者登録番号（代表）],[取引主任者名（代表）],[全国賃貸不動産管理業協会会員番号],[賃貸不動産経営管理士登録番号],[賃貸不動産経営管理士名],[賃貸住宅管理業者登録番号],[加入団体１],[加入団体２],[加入団体３],[加入団体４],[加入団体５],[加入団体６],[加入団体７],[加入団体８],[加入団体９],[加入団体１０],[代表者名（SJIS）],[取引主任者名（代表）（SJIS）],[賃貸不動産経営管理士名（SJIS）]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 jisya_no AS [自社・支店No] "
                tmp_sql = tmp_sql & " 		,kaisya_name AS [自社・支店名] "
                tmp_sql = tmp_sql & " 		,'' AS [自社・支店名（SJIS）] "
                tmp_sql = tmp_sql & " 		,kaisya_kana AS [自社・支店カナ] "
                tmp_sql = tmp_sql & " 		,'' AS [自社・支店名２] "
                tmp_sql = tmp_sql & " 		,'' AS [自社・支店名２（SJIS）] "
                tmp_sql = tmp_sql & " 		,'' AS [自社・支店名３] "
                tmp_sql = tmp_sql & " 		,'' AS [自社・支店名３（SJIS）] "
                tmp_sql = tmp_sql & " 		,yakusyoku AS [代表者役職] "
                tmp_sql = tmp_sql & " 		,daihyo AS [代表者名] "
                tmp_sql = tmp_sql & " 		,kaisya_post AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,kaisya_add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,kaisya_add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,kaisya_tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,kaisya_tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯TEL1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯TEL2] "
                tmp_sql = tmp_sql & " 		,kaisya_fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,kaisya_email AS [メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 		,kaisya_url AS [Webアドレス] "
                tmp_sql = tmp_sql & " 		,[備考(支店備考用ダミー列)] AS [備考（基本情報）] "
                tmp_sql = tmp_sql & " 		,menkyo_no AS [免許証番号] "
                tmp_sql = tmp_sql & " 		,menkyo_ymd AS [免許年月日] "
                tmp_sql = tmp_sql & " 		,takuken_no AS [取引主任者登録番号（代表）] "
                tmp_sql = tmp_sql & " 		,takuken_name AS [取引主任者名（代表）] "
                tmp_sql = tmp_sql & " 		,'' AS [全国賃貸不動産管理業協会会員番号] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸不動産経営管理士登録番号] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸不動産経営管理士名] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸住宅管理業者登録番号] "
                tmp_sql = tmp_sql & " 		,dantai1 AS [加入団体１] "
                tmp_sql = tmp_sql & " 		,dantai2 AS [加入団体２] "
                tmp_sql = tmp_sql & " 		,dantai3 AS [加入団体３] "
                tmp_sql = tmp_sql & " 		,dantai4 AS [加入団体４] "
                tmp_sql = tmp_sql & " 		,dantai5 AS [加入団体５] "
                tmp_sql = tmp_sql & " 		,dantai6 AS [加入団体６] "
                tmp_sql = tmp_sql & " 		,dantai7 AS [加入団体７] "
                tmp_sql = tmp_sql & " 		,dantai8 AS [加入団体８] "
                tmp_sql = tmp_sql & " 		,dantai9 AS [加入団体９] "
                tmp_sql = tmp_sql & " 		,dantai10 AS [加入団体１０] "
                tmp_sql = tmp_sql & " 		,'' AS [代表者名（SJIS）] "
                tmp_sql = tmp_sql & " 		,'' AS [取引主任者名（代表）（SJIS）] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸不動産経営管理士名（SJIS）] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT * FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 [jisya_no] "
                tmp_sql = tmp_sql & " 				,[kaisya_name] "
                tmp_sql = tmp_sql & " 				,[kaisya_kana] "
                tmp_sql = tmp_sql & " 				,[yakusyoku] "
                tmp_sql = tmp_sql & " 				,[daihyo] "
                tmp_sql = tmp_sql & " 				,[kaisya_post] "
                tmp_sql = tmp_sql & " 				,[kaisya_add1] "
                tmp_sql = tmp_sql & " 				,[kaisya_add2] "
                tmp_sql = tmp_sql & " 				,[kaisya_tel1] "
                tmp_sql = tmp_sql & " 				,[kaisya_tel2] "
                tmp_sql = tmp_sql & " 				,[kaisya_fax] "
                tmp_sql = tmp_sql & " 				,[kaisya_email] "
                tmp_sql = tmp_sql & " 				,[kaisya_url] "
                tmp_sql = tmp_sql & " 				,'' AS [備考(支店備考用ダミー列)] "
                tmp_sql = tmp_sql & " 				,[menkyo_no] "
                tmp_sql = tmp_sql & " 				,[menkyo_ymd] "
                tmp_sql = tmp_sql & " 				,[takuken_no] "
                tmp_sql = tmp_sql & " 				,[takuken_name] "
                tmp_sql = tmp_sql & " 				,[dantai1] "
                tmp_sql = tmp_sql & " 				,[dantai2] "
                tmp_sql = tmp_sql & " 				,[dantai3] "
                tmp_sql = tmp_sql & " 				,[dantai4] "
                tmp_sql = tmp_sql & " 				,[dantai5] "
                tmp_sql = tmp_sql & " 				,[dantai6] "
                tmp_sql = tmp_sql & " 				,[dantai7] "
                tmp_sql = tmp_sql & " 				,[dantai8] "
                tmp_sql = tmp_sql & " 				,[dantai9] "
                tmp_sql = tmp_sql & " 				,[dantai10] "
                tmp_sql = tmp_sql & " 			FROM m_jisya "
                tmp_sql = tmp_sql & " 		) AS JISYA "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT * FROM "
                tmp_sql = tmp_sql & " 		( "
                tmp_sql = tmp_sql & " 			SELECT "
                tmp_sql = tmp_sql & " 				 ROW_NUMBER()OVER(ORDER BY [siten_no]) + "
                tmp_sql = tmp_sql & " 					( "
                tmp_sql = tmp_sql & " 						SELECT jisya_no FROM "
                tmp_sql = tmp_sql & " 						( "
                tmp_sql = tmp_sql & " 							SELECT "
                tmp_sql = tmp_sql & " 								ROW_NUMBER()OVER(ORDER BY jisya_no DESC) AS 抽出用 "
                tmp_sql = tmp_sql & " 								,* "
                tmp_sql = tmp_sql & " 							FROM m_jisya "
                tmp_sql = tmp_sql & " 						) AS VW "
                tmp_sql = tmp_sql & " 						WHERE 抽出用 = 1 "
                tmp_sql = tmp_sql & " 					) AS [移行用自社№] "
                tmp_sql = tmp_sql & " 				 ,[siten_name] "
                tmp_sql = tmp_sql & " 				,[siten_kana] AS [カナ] "
                tmp_sql = tmp_sql & " 				,'' AS [代表者役職] "
                tmp_sql = tmp_sql & " 				,'' AS [代表者名称] "
                tmp_sql = tmp_sql & " 				,[post] "
                tmp_sql = tmp_sql & " 				,[add1] "
                tmp_sql = tmp_sql & " 				,[add2] "
                tmp_sql = tmp_sql & " 				,[tel1] "
                tmp_sql = tmp_sql & " 				,[tel2] "
                tmp_sql = tmp_sql & " 				,[fax] "
                tmp_sql = tmp_sql & " 				,[email] "
                tmp_sql = tmp_sql & " 				,'' AS [url] "
                tmp_sql = tmp_sql & " 				,[biko] "
                tmp_sql = tmp_sql & " 				,'' AS [免許証番号] "
                tmp_sql = tmp_sql & " 				,'' AS [免許証年月日] "
                tmp_sql = tmp_sql & " 				,'' AS [宅建取引士登録番号] "
                tmp_sql = tmp_sql & " 				,'' AS [宅建取引士氏名] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体1] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体2] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体3] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体4] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体5] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体6] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体7] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体8] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体9] "
                tmp_sql = tmp_sql & " 				,'' AS [加入団体10] "
                tmp_sql = tmp_sql & " 			FROM m_siten "
                tmp_sql = tmp_sql & " 		) AS SITEN "
                tmp_sql = tmp_sql & " 	) AS VW "
                tmp_sql = tmp_sql & " ) AS JISYA "

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

#Region "自社口座情報"

    Public Class M_jisya_Koza_Repository

        '紐付ツールでの対応が必要
        '検証用に自社No1で固定して移行しておく

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[口座取得元],[各口座設定No],[金融機関No],[金融機関店No]"

                '20160516 FB構築に伴う自社口座情報の修正 -chg sta
                'tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 kinyu_no AS [金融機関No] "
                'tmp_sql = tmp_sql & " 		,ten_no AS [金融機関店No] "
                'tmp_sql = tmp_sql & " 		,kosyu_name AS [口座種別] "
                'tmp_sql = tmp_sql & " 		,koza_no AS [口座番号] "
                'tmp_sql = tmp_sql & " 		,koza_meigi AS [口座名義] "
                'tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
                'tmp_sql = tmp_sql & " 		,post_kigo AS [ゆうちょ口座記号１] "
                'tmp_sql = tmp_sql & " 		,post_n AS [ゆうちょ口座記号２] "
                'tmp_sql = tmp_sql & " 		,post_bango AS [ゆうちょ口座番号] "
                'tmp_sql = tmp_sql & " 		,KOZATOTAL.biko AS [備考(口座情報)]	 "
                'tmp_sql = tmp_sql & " 		,[取得元] "
                'tmp_sql = tmp_sql & " 	FROM "
                'tmp_sql = tmp_sql & " 	( "
                'tmp_sql = tmp_sql & " 		SELECT "
                'tmp_sql = tmp_sql & " 			 kinyu_no "
                'tmp_sql = tmp_sql & " 			,ten_no "
                'tmp_sql = tmp_sql & " 			,kosyu_no "
                'tmp_sql = tmp_sql & " 			,koza_no "
                'tmp_sql = tmp_sql & " 			,koza_meigi "
                'tmp_sql = tmp_sql & " 			,koza_kana "
                'tmp_sql = tmp_sql & " 			,'' AS post_kigo "
                'tmp_sql = tmp_sql & " 			,'' AS post_n "
                'tmp_sql = tmp_sql & " 			,'' AS post_bango "
                'tmp_sql = tmp_sql & " 			,'' AS biko "
                'tmp_sql = tmp_sql & " 			,'振込依頼人マスタ' AS [取得元] "
                'tmp_sql = tmp_sql & " 		FROM m_furi_irai "
                'tmp_sql = tmp_sql & " 		UNION "
                'tmp_sql = tmp_sql & " 		SELECT "
                'tmp_sql = tmp_sql & " 			 kinyu_no "
                'tmp_sql = tmp_sql & " 			,ten_no "
                'tmp_sql = tmp_sql & " 			,kosyu_no "
                'tmp_sql = tmp_sql & " 			,koza_no "
                'tmp_sql = tmp_sql & " 			,koza_meigi "
                'tmp_sql = tmp_sql & " 			,koza_kana "
                'tmp_sql = tmp_sql & " 			,post_kigo "
                'tmp_sql = tmp_sql & " 			,post_n "
                'tmp_sql = tmp_sql & " 			,post_bango "
                'tmp_sql = tmp_sql & " 			,biko "
                'tmp_sql = tmp_sql & " 			,'振込先口座マスタ' AS [取得元] "
                'tmp_sql = tmp_sql & " 		FROM m_furi_koza AS MFK "
                'tmp_sql = tmp_sql & " 		/*20160517 家賃入金口座情報の新規作成に伴う修正 -add sta*/ "
                'tmp_sql = tmp_sql & " 		/*家主口座と一致しないデータに対して紐付けを行う*/ "
                'tmp_sql = tmp_sql & " 		WHERE NOT EXISTS "
                'tmp_sql = tmp_sql & " 			( "
                'tmp_sql = tmp_sql & " 				SELECT * FROM m_yanu_so_koza AS OWKOZA "
                'tmp_sql = tmp_sql & " 				WHERE MFK.kinyu_no = OWKOZA.kinyu_no "
                'tmp_sql = tmp_sql & " 				AND   MFK.ten_no = OWKOZA.ten_no "
                'tmp_sql = tmp_sql & " 				AND   MFK.koza_no = OWKOZA.koza_no				 "
                'tmp_sql = tmp_sql & " 			) "
                'tmp_sql = tmp_sql & " 		/*20160517 家賃入金口座情報の新規作成に伴う修正 -add end*/ "
                'tmp_sql = tmp_sql & " 		UNION "
                'tmp_sql = tmp_sql & " 		SELECT "
                'tmp_sql = tmp_sql & " 			 kinyu_no "
                'tmp_sql = tmp_sql & " 			,ten_no "
                'tmp_sql = tmp_sql & " 			,kosyu_no "
                'tmp_sql = tmp_sql & " 			,koza_no "
                'tmp_sql = tmp_sql & " 			,koza_meigi "
                'tmp_sql = tmp_sql & " 			,koza_kana "
                'tmp_sql = tmp_sql & " 			,'' AS post_kigo "
                'tmp_sql = tmp_sql & " 			,'' AS post_n "
                'tmp_sql = tmp_sql & " 			,'' AS post_bango "
                'tmp_sql = tmp_sql & " 			,'' AS biko "
                'tmp_sql = tmp_sql & " 			,'口座振替マスタ' AS [取得元] "
                'tmp_sql = tmp_sql & " 		FROM m_koza_furi "
                'tmp_sql = tmp_sql & " 	) AS KOZATOTAL "
                'tmp_sql = tmp_sql & " 	LEFT JOIN m_kozasyu AS MKS "
                'tmp_sql = tmp_sql & " 	ON KOZATOTAL.kosyu_no = MKS.kosyu_no "
                'tmp_sql = tmp_sql & " 	/*20160516 FB構築に伴う自社口座情報の修正 -del*/ "
                'tmp_sql = tmp_sql & " 	/*WHERE kinyu_no IS NOT NULL*/ "
                'tmp_sql = tmp_sql & " ) AS VW "


                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[口座取得元],[各口座設定No],[各口座設定名称],[金融機関No],[金融機関店No],[口座種別],[口座番号],[口座名義],[口座名義カナ],[ゆうちょ口座記号１],[ゆうちょ口座記号２],[ゆうちょ口座番号],[備考(口座情報)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 [口座取得元] "
                tmp_sql = tmp_sql & " 		/*20160829 自社口座にデフォルト値を設定する処理を追加 add*/ "
                tmp_sql = tmp_sql & " 		,fkom_no AS [各口座設定No] "
                tmp_sql = tmp_sql & " 		,fkom_name AS [各口座設定名称] "
                tmp_sql = tmp_sql & " 		,kinyu_no AS [金融機関No] "
                tmp_sql = tmp_sql & " 		,ten_no AS [金融機関店No] "
                tmp_sql = tmp_sql & " 		,kosyu_name AS [口座種別] "
                tmp_sql = tmp_sql & " 		,koza_no AS [口座番号] "
                tmp_sql = tmp_sql & " 		,koza_meigi AS [口座名義] "
                tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
                tmp_sql = tmp_sql & " 		,post_kigo AS [ゆうちょ口座記号１] "
                tmp_sql = tmp_sql & " 		,post_n AS [ゆうちょ口座記号２] "
                tmp_sql = tmp_sql & " 		,post_bango AS [ゆうちょ口座番号] "
                tmp_sql = tmp_sql & " 		,KOZATOTAL.biko AS [備考(口座情報)]	 "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			/*20160829 自社口座にデフォルト値を設定する処理を追加 add*/ "
                tmp_sql = tmp_sql & " 			 fkom_no "
                tmp_sql = tmp_sql & " 			,fkom_name "
                tmp_sql = tmp_sql & " 			,kinyu_no "
                tmp_sql = tmp_sql & " 			,ten_no "
                tmp_sql = tmp_sql & " 			,kosyu_no "
                tmp_sql = tmp_sql & " 			,koza_no "
                tmp_sql = tmp_sql & " 			,koza_meigi "
                tmp_sql = tmp_sql & " 			,koza_kana "
                tmp_sql = tmp_sql & " 			,'' AS post_kigo "
                tmp_sql = tmp_sql & " 			,'' AS post_n "
                tmp_sql = tmp_sql & " 			,'' AS post_bango "
                tmp_sql = tmp_sql & " 			,'' AS biko "
                tmp_sql = tmp_sql & " 			,'振込依頼人マスタ' AS [口座取得元] "
                tmp_sql = tmp_sql & " 		FROM m_furi_irai "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			/*20160829 自社口座にデフォルト値を設定する処理を追加 add*/ "
                tmp_sql = tmp_sql & " 			 fkom_no "
                tmp_sql = tmp_sql & " 			,fkom_name "
                tmp_sql = tmp_sql & " 			,kinyu_no "
                tmp_sql = tmp_sql & " 			,ten_no "
                tmp_sql = tmp_sql & " 			,kosyu_no "
                tmp_sql = tmp_sql & " 			,koza_no "
                tmp_sql = tmp_sql & " 			,koza_meigi "
                tmp_sql = tmp_sql & " 			,koza_kana "
                tmp_sql = tmp_sql & " 			,post_kigo "
                tmp_sql = tmp_sql & " 			,post_n "
                tmp_sql = tmp_sql & " 			,post_bango "
                tmp_sql = tmp_sql & " 			,biko "
                tmp_sql = tmp_sql & " 			,'振込先口座マスタ' AS [口座取得元] "
                tmp_sql = tmp_sql & " 		FROM m_furi_koza AS MFK "
                tmp_sql = tmp_sql & " 		/*家主口座と一致しないデータに対して紐付けを行う*/ "
                tmp_sql = tmp_sql & " 		WHERE NOT EXISTS "
                tmp_sql = tmp_sql & " 			( "
                tmp_sql = tmp_sql & " 				SELECT * FROM m_yanu_so_koza AS OWKOZA "
                tmp_sql = tmp_sql & " 				WHERE MFK.kinyu_no = OWKOZA.kinyu_no "
                tmp_sql = tmp_sql & " 				AND   MFK.ten_no = OWKOZA.ten_no "
                tmp_sql = tmp_sql & " 				AND   MFK.kosyu_no = OWKOZA.kosyu_no "
                tmp_sql = tmp_sql & " 				AND   MFK.koza_no = OWKOZA.koza_no "
                tmp_sql = tmp_sql & " 			) "
                tmp_sql = tmp_sql & " 		UNION "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			/*20160829 自社口座にデフォルト値を設定する処理を追加 add*/ "
                tmp_sql = tmp_sql & " 			 fkae_no "
                tmp_sql = tmp_sql & " 			,fkae_name "
                tmp_sql = tmp_sql & " 			,kinyu_no "
                tmp_sql = tmp_sql & " 			,ten_no "
                tmp_sql = tmp_sql & " 			,kosyu_no "
                tmp_sql = tmp_sql & " 			,koza_no "
                tmp_sql = tmp_sql & " 			,koza_meigi "
                tmp_sql = tmp_sql & " 			,koza_kana "
                tmp_sql = tmp_sql & " 			,'' AS post_kigo "
                tmp_sql = tmp_sql & " 			,'' AS post_n "
                tmp_sql = tmp_sql & " 			,'' AS post_bango "
                tmp_sql = tmp_sql & " 			,'' AS biko "
                tmp_sql = tmp_sql & " 			,'口座振替マスタ' AS [口座取得元] "
                tmp_sql = tmp_sql & " 		FROM m_koza_furi "
                tmp_sql = tmp_sql & " 	) AS KOZATOTAL "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_kozasyu AS MKS "
                tmp_sql = tmp_sql & " 	ON KOZATOTAL.kosyu_no = MKS.kosyu_no "
                tmp_sql = tmp_sql & " 	/*20160829 自社口座にデフォルト値を設定する処理を追加 del sta*/ "
                tmp_sql = tmp_sql & " 	/* "
                tmp_sql = tmp_sql & " 	/*20160722 口座情報が存在しないことによるエラーの修正 ※応急処置 add*/ "
                tmp_sql = tmp_sql & " 	WHERE ISNULL(kinyu_no,0) <> 0 AND ISNULL(ten_no,0) <> 0 "
                tmp_sql = tmp_sql & " 	*/ "
                tmp_sql = tmp_sql & " 	/*20160829 自社口座にデフォルト値を設定する処理を追加 del end*/ "
                tmp_sql = tmp_sql & " ) AS VW "
                '20160516 FB構築に伴う自社口座情報の修正 -chg end

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

#Region "自社担当者情報" '20160608 自社担当者情報構築

    Public Class M_tanto_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[賃貸革命ログオンユーザNo]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[賃貸革命ログオンユーザNo],[賃貸革命ログオンユーザ名],[賃貸革命ログオンユーザ名(SJIS)],[賃貸革命ログオンユーザ名カナ],[賃貸革命ログオングループNo],[TEL1],[TEL2],[携帯電話1],[FAX1],[メールアドレス1],[携帯メールアドレス1],[宅建免許番号],[備考],[担当者ログオン禁止フラグ],[使用しないフラグ],[ログオンパスワード],[windowsログオンユーザー名],[意見画面表示フラグ]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 tanto_no AS [賃貸革命ログオンユーザNo] "
                tmp_sql = tmp_sql & " 		,tanto_name AS [賃貸革命ログオンユーザ名] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸革命ログオンユーザ名(SJIS)] "
                tmp_sql = tmp_sql & " 		,tanto_kana AS [賃貸革命ログオンユーザ名カナ] "
                tmp_sql = tmp_sql & " 		,99001 AS [賃貸革命ログオングループNo] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯電話1] "
                tmp_sql = tmp_sql & " 		,'' AS [FAX1] "
                tmp_sql = tmp_sql & " 		,email AS [メールアドレス1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス1] "
                tmp_sql = tmp_sql & " 		,torihiki_bango AS [宅建免許番号] "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 		,0 AS [担当者ログオン禁止フラグ] "
                tmp_sql = tmp_sql & " 		,0 AS [使用しないフラグ] "
                tmp_sql = tmp_sql & " 		,'' AS [ログオンパスワード] "
                tmp_sql = tmp_sql & " 		,'' AS [windowsログオンユーザー名] "
                tmp_sql = tmp_sql & " 		,0 AS [意見画面表示フラグ] "
                tmp_sql = tmp_sql & " 	FROM m_tanto "
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

#Region "自社メモ情報"

    Public Class M_jisya_Biko_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[自社・支店No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[自社・支店No],[メモ1],[メモ2],[メモ3],[メモ4],[メモ5],[メモ6],[メモ7],[メモ8],[メモ9],[メモ10],[メモ11],[メモ12],[メモ13],[メモ14],[メモ15]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 [jisya_no] AS [自社・支店No] "
                tmp_sql = tmp_sql & " 		,[biko1] AS [メモ1] "
                tmp_sql = tmp_sql & " 		,[biko2] AS [メモ2] "
                tmp_sql = tmp_sql & " 		,[biko3] AS [メモ3] "
                tmp_sql = tmp_sql & " 		,[biko4] AS [メモ4] "
                tmp_sql = tmp_sql & " 		,[biko5] AS [メモ5] "
                tmp_sql = tmp_sql & " 		,[biko6] AS [メモ6] "
                tmp_sql = tmp_sql & " 		,[biko7] AS [メモ7] "
                tmp_sql = tmp_sql & " 		,[biko8] AS [メモ8] "
                tmp_sql = tmp_sql & " 		,[biko9] AS [メモ9] "
                tmp_sql = tmp_sql & " 		,[biko10] AS [メモ10] "
                tmp_sql = tmp_sql & " 		,[biko11] AS [メモ11] "
                tmp_sql = tmp_sql & " 		,[biko12] AS [メモ12] "
                tmp_sql = tmp_sql & " 		,[biko13] AS [メモ13] "
                tmp_sql = tmp_sql & " 		,[biko14] AS [メモ14] "
                tmp_sql = tmp_sql & " 		,[biko15] AS [メモ15] "
                tmp_sql = tmp_sql & " 	FROM m_jisya "
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

#Region "振込依頼人情報"

    Public Class M_furi_irai_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[振込依頼人No]"

                '2016.04.26 メインの方へも反映させる修正 -chg sta
                '紐付けファイルを参照するように変更する
                'tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 fkom_no AS [振込依頼人No] "
                'tmp_sql = tmp_sql & " 		,fkom_name AS [振込依頼人名] "
                'tmp_sql = tmp_sql & " 		,[自社支店] AS [自社支店No.] "
                'tmp_sql = tmp_sql & " 		,[口座No] AS [自社口座No] "
                'tmp_sql = tmp_sql & " 		,fkom_ircd AS [振込依頼人コード] "
                'tmp_sql = tmp_sql & " 		,fkom_irkana AS [振込依頼人カナ] "
                'tmp_sql = tmp_sql & " 		,MFI.biko AS [備考(基本情報)] "
                'tmp_sql = tmp_sql & " 		,1 AS [総合振込全銀フォーマットNo]	/*全銀フォーマットからNoを取得する。全銀フォーマットマスタ作成後に対応するためとりあえず1を設定しておく*/ "
                'tmp_sql = tmp_sql & " 		,sosin_fname AS [送信ファイルパス] "
                'tmp_sql = tmp_sql & " 		,CASE "
                'tmp_sql = tmp_sql & " 			WHEN crlf = 1 THEN 0 "
                'tmp_sql = tmp_sql & " 			WHEN crlf = 2 THEN 1 "
                'tmp_sql = tmp_sql & " 		 END AS [改行有無(0:改行しない 1:改行する)] "
                'tmp_sql = tmp_sql & " 		,1 AS [ヲ変換フラグ(0:変換しない 1:変換する)] "
                'tmp_sql = tmp_sql & " 		,'' AS [備考(データ)] "
                'tmp_sql = tmp_sql & " 		,defflg AS [計算時の初期値フラグ(0:初期値としない 1:初期値とする)] "
                'tmp_sql = tmp_sql & " 		,0 AS [負担調整フラグ(0:調整しない 1:調整する)] "
                'tmp_sql = tmp_sql & " 		,'' AS [備考(手数料)] "
                'tmp_sql = tmp_sql & " 		,0 AS [総合振込データ生成フラグ] "
                'tmp_sql = tmp_sql & " 	FROM m_furi_irai AS MFI "
                'tmp_sql = tmp_sql & " 	LEFT JOIN "
                'tmp_sql = tmp_sql & " 		( "
                'tmp_sql = tmp_sql & " 			SELECT "
                'tmp_sql = tmp_sql & " 				 1 AS [自社支店]	/*紐付ツールが構築中のため1を設定しておく*/ "
                'tmp_sql = tmp_sql & " 				,ROW_NUMBER()OVER(ORDER BY kinyu_no,ten_no,kosyu_no,koza_no) AS [口座No] "
                'tmp_sql = tmp_sql & " 				,* "
                'tmp_sql = tmp_sql & " 			FROM "
                'tmp_sql = tmp_sql & " 			( "
                'tmp_sql = tmp_sql & " 				SELECT "
                'tmp_sql = tmp_sql & " 					 kinyu_no "
                'tmp_sql = tmp_sql & " 					,ten_no "
                'tmp_sql = tmp_sql & " 					,kosyu_no "
                'tmp_sql = tmp_sql & " 					,koza_no "
                'tmp_sql = tmp_sql & " 					,koza_meigi "
                'tmp_sql = tmp_sql & " 					,koza_kana "
                'tmp_sql = tmp_sql & " 					,'' AS post_kigo "
                'tmp_sql = tmp_sql & " 					,'' AS post_n "
                'tmp_sql = tmp_sql & " 					,'' AS post_bango "
                'tmp_sql = tmp_sql & " 					,'' AS biko "
                'tmp_sql = tmp_sql & " 				FROM m_furi_irai "
                'tmp_sql = tmp_sql & " 				UNION "
                'tmp_sql = tmp_sql & " 				SELECT "
                'tmp_sql = tmp_sql & " 					 kinyu_no "
                'tmp_sql = tmp_sql & " 					,ten_no "
                'tmp_sql = tmp_sql & " 					,kosyu_no "
                'tmp_sql = tmp_sql & " 					,koza_no "
                'tmp_sql = tmp_sql & " 					,koza_meigi "
                'tmp_sql = tmp_sql & " 					,koza_kana "
                'tmp_sql = tmp_sql & " 					,post_kigo "
                'tmp_sql = tmp_sql & " 					,post_n "
                'tmp_sql = tmp_sql & " 					,post_bango "
                'tmp_sql = tmp_sql & " 					,biko "
                'tmp_sql = tmp_sql & " 				FROM m_furi_koza "
                'tmp_sql = tmp_sql & " 				UNION "
                'tmp_sql = tmp_sql & " 				SELECT "
                'tmp_sql = tmp_sql & " 					 kinyu_no "
                'tmp_sql = tmp_sql & " 					,ten_no "
                'tmp_sql = tmp_sql & " 					,kosyu_no "
                'tmp_sql = tmp_sql & " 					,koza_no "
                'tmp_sql = tmp_sql & " 					,koza_meigi "
                'tmp_sql = tmp_sql & " 					,koza_kana "
                'tmp_sql = tmp_sql & " 					,'' AS post_kigo "
                'tmp_sql = tmp_sql & " 					,'' AS post_n "
                'tmp_sql = tmp_sql & " 					,'' AS post_bango "
                'tmp_sql = tmp_sql & " 					,'' AS biko "
                'tmp_sql = tmp_sql & " 				FROM m_koza_furi "
                'tmp_sql = tmp_sql & " 			) AS VW "
                'tmp_sql = tmp_sql & " 		) AS KOZATOTAL "
                'tmp_sql = tmp_sql & " 	ON  MFI.kinyu_no = KOZATOTAL.kinyu_no "
                'tmp_sql = tmp_sql & " 	AND MFI.ten_no = KOZATOTAL.ten_no "
                'tmp_sql = tmp_sql & " 	AND MFI.kosyu_no = KOZATOTAL.kosyu_no "
                'tmp_sql = tmp_sql & " 	AND MFI.koza_no = KOZATOTAL.koza_no "
                'tmp_sql = tmp_sql & " ) AS VW "

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[振込依頼人No],[振込依頼人名],[金融機関No],[金融機関支店No],[口座種別],[口座番号],[口座名義],[振込依頼人コード],[振込依頼人カナ],[備考(基本情報)],[送信ファイルパス],[ヲ変換フラグ(0:変換しない 1:変換する)],[負担調整フラグ(0:調整しない 1:調整する)],[総合振込データ生成フラグ]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 fkom_no AS [振込依頼人No] "
                tmp_sql = tmp_sql & " 		,fkom_name AS [振込依頼人名] "
                tmp_sql = tmp_sql & " 		,kinyu_no AS [金融機関No]		/*紐付けられた自社No自社口座No取得用*/ "
                tmp_sql = tmp_sql & " 		,ten_no AS [金融機関支店No]		/*紐付けられた自社No自社口座No取得用*/ "
                tmp_sql = tmp_sql & " 		,kosyu_name AS [口座種別]		/*紐付けられた自社No自社口座No取得用*/	/*20160609 紐付取得用に追加*/ "
                tmp_sql = tmp_sql & " 		,koza_no AS [口座番号]			/*紐付けられた自社No自社口座No取得用*/ "
                tmp_sql = tmp_sql & " 		,koza_meigi AS [口座名義]		/*紐付けられた自社No自社口座No取得用*/	/*20160609 紐付取得用に追加*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,[自社支店] AS [自社支店No] "
                tmp_sql = tmp_sql & " 		,[口座No] AS [自社口座No] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,fkom_ircd AS [振込依頼人コード] "
                tmp_sql = tmp_sql & " 		,fkom_irkana AS [振込依頼人カナ] "
                tmp_sql = tmp_sql & " 		,MFI.biko AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		/*20160609 紐付設定値取得に伴う修正 del*/ "
                tmp_sql = tmp_sql & " 		/*,1 AS [総合振込全銀フォーマットNo]	/*全銀フォーマットからNoを取得する。全銀フォーマットマスタ作成後に対応するためとりあえず1を設定しておく*/*/ "
                tmp_sql = tmp_sql & " 		,sosin_fname AS [送信ファイルパス] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 振込依頼人情報 del sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN crlf = 1 THEN 0 "
                tmp_sql = tmp_sql & " 			WHEN crlf = 2 THEN 1 "
                tmp_sql = tmp_sql & " 		 END AS [改行有無(0:改行しない 1:改行する)] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 振込依頼人情報 del end*/ "
                tmp_sql = tmp_sql & " 		,1 AS [ヲ変換フラグ(0:変換しない 1:変換する)] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 振込依頼人情報 del*/ "
                tmp_sql = tmp_sql & " 		/*,'' AS [備考(データ)]*/ "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 振込依頼人情報 del*/ "
                tmp_sql = tmp_sql & " 		/*,defflg AS [計算時の初期値フラグ(0:初期値としない 1:初期値とする)]*/ "
                tmp_sql = tmp_sql & " 		,0 AS [負担調整フラグ(0:調整しない 1:調整する)] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 振込依頼人情報 del*/ "
                tmp_sql = tmp_sql & " 		/*,'' AS [備考(手数料)]*/ "
                tmp_sql = tmp_sql & " 		,0 AS [総合振込データ生成フラグ] "
                tmp_sql = tmp_sql & " 	FROM m_furi_irai AS MFI "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_kozasyu AS MK ON MFI.kosyu_no = MK.kosyu_no	/*20160609 紐付取得用に追加*/ "
                tmp_sql = tmp_sql & " ) AS VW "
                '2016.04.26 メインの方へも反映させる修正 -chg end

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

#Region "振込手数料情報"   '20160517 振込手数料情報の新規作成

    Public Class M_furi_tesu_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[振込依頼人No],[レコードNo]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[振込依頼人No],[レコードNo],[金額From],[金額To],[同行同支店手数料(電信扱)],[同行同支店手数料(文書扱)],[同行他支店手数料(電信扱)],[同行他支店手数料(文書扱)],[他行手数料(電信扱)],[他行手数料(文書扱)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	/*実際に存在するデータ sta*/ "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 fkom_no AS [振込依頼人No] "
                tmp_sql = tmp_sql & " 		,hanni_no - 1 AS [レコードNo] "
                tmp_sql = tmp_sql & " 		,hanni_small AS [金額From] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN hanni_large = 0 THEN NULL "
                tmp_sql = tmp_sql & " 			ELSE hanni_large "
                tmp_sql = tmp_sql & " 		 END AS [金額To] "
                tmp_sql = tmp_sql & " 		,mtesu1 AS [同行同支店手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,mtesu2 AS [同行同支店手数料(文書扱)] "
                tmp_sql = tmp_sql & " 		,mtesu3 AS [同行他支店手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,mtesu4 AS [同行他支店手数料(文書扱)] "
                tmp_sql = tmp_sql & " 		,mtesu5 AS [他行手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,mtesu6 AS [他行手数料(文書扱)] "
                tmp_sql = tmp_sql & " 	FROM m_furi_tesu "
                tmp_sql = tmp_sql & " 	/*実際に存在するデータ end*/ "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	/*存在しないデータの1行目を作成 sta*/ "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 MFI.fkom_no AS [振込依頼人No] "
                tmp_sql = tmp_sql & " 		,0 AS [レコードNo] "
                tmp_sql = tmp_sql & " 		,0 AS [金額From] "
                tmp_sql = tmp_sql & " 		,30000 [金額To] "
                tmp_sql = tmp_sql & " 		,NULL AS [同行同支店手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [同行同支店手数料(文書扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [同行他支店手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [同行他支店手数料(文書扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [他行手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [他行手数料(文書扱)] "
                tmp_sql = tmp_sql & " 	FROM m_furi_irai AS MFI "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_furi_tesu AS MFT ON MFI.fkom_no = MFT.fkom_no "
                tmp_sql = tmp_sql & " 	WHERE hanni_no IS NULL "
                tmp_sql = tmp_sql & " 	/*存在しないデータの1行目を作成 end*/ "
                tmp_sql = tmp_sql & " 	UNION "
                tmp_sql = tmp_sql & " 	/*存在しないデータの2行目を作成 sta*/ "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 MFI.fkom_no AS [振込依頼人No] "
                tmp_sql = tmp_sql & " 		,1 AS [レコードNo] "
                tmp_sql = tmp_sql & " 		,30000 AS [金額From] "
                tmp_sql = tmp_sql & " 		,NULL [金額To] "
                tmp_sql = tmp_sql & " 		,NULL AS [同行同支店手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [同行同支店手数料(文書扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [同行他支店手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [同行他支店手数料(文書扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [他行手数料(電信扱)] "
                tmp_sql = tmp_sql & " 		,NULL AS [他行手数料(文書扱)] "
                tmp_sql = tmp_sql & " 	FROM m_furi_irai AS MFI "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_furi_tesu AS MFT ON MFI.fkom_no = MFT.fkom_no "
                tmp_sql = tmp_sql & " 	WHERE hanni_no IS NULL "
                tmp_sql = tmp_sql & " 	/*存在しないデータの2行目を作成 end*/ "
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

#Region "口座振替情報"

    Public Class M_koza_furi_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[振替情報No]"

                '2016.04.26 メインの方へも反映させる修正 -chg sta
                '紐付けファイルを参照するように変更する
                'tmp_sql = tmp_sql & " SELECT * FROM "
                'tmp_sql = tmp_sql & " ( "
                'tmp_sql = tmp_sql & " 	SELECT "
                'tmp_sql = tmp_sql & " 		 fkae_no AS [振替情報No.] "
                'tmp_sql = tmp_sql & " 		,fkae_name AS [振替情報名称] "
                'tmp_sql = tmp_sql & " 		,fkae_kana AS [振替情報カナ名称] "
                'tmp_sql = tmp_sql & " 		,NULL AS [サービスタイプ　1:オリコ　2:ジャックス　3:アプラス] "
                'tmp_sql = tmp_sql & " 		,fkae_ircd AS [振込依頼人] "
                'tmp_sql = tmp_sql & " 		,fkae_irkana AS [振込依頼人カナ名] "
                'tmp_sql = tmp_sql & " 		,kamei_no AS [加盟店No.　依頼人番号と共用] "
                'tmp_sql = tmp_sql & " 		,NULL AS [ジェイリース　入金区分　1:変更なし・・・・] "
                'tmp_sql = tmp_sql & " 		,[自社支店] AS [自社支店No] "
                'tmp_sql = tmp_sql & " 		,[口座No] AS [振替先口座No.] "
                'tmp_sql = tmp_sql & " 		,hiki1 AS [引落日] "
                'tmp_sql = tmp_sql & " 		,fkae_tesu AS [手数料　請求額] "
                'tmp_sql = tmp_sql & " 		,MKF.biko AS [備考] "
                'tmp_sql = tmp_sql & " 		/* "
                'tmp_sql = tmp_sql & " 		全銀フォーマットマスタ移行後に再構築。とりあえず1を設定しておく "
                'tmp_sql = tmp_sql & " 		,syudai_kbn AS [オリジナルフォーマットNo] "
                'tmp_sql = tmp_sql & " 		*/ "
                'tmp_sql = tmp_sql & " 		,1 AS [オリジナルフォーマットNo] "
                'tmp_sql = tmp_sql & " 		,CASE "
                'tmp_sql = tmp_sql & " 			WHEN yobi_bt1 = 0 THEN 1 "
                'tmp_sql = tmp_sql & " 			WHEN yobi_bt1 = 1 THEN 2 "
                'tmp_sql = tmp_sql & " 		 END AS [データ並び替え　0:物件番号と部屋番号　1:契約者番号] "
                'tmp_sql = tmp_sql & " 		,1 AS [契約者No.を契約者番号に出力する  0:出力しない　1:出力する] "
                'tmp_sql = tmp_sql & " 		,crlf AS [改行コード(CRLF)出力　0:なし　1:あり] "
                'tmp_sql = tmp_sql & " 		,CASE "
                'tmp_sql = tmp_sql & " 			WHEN yobi_bt4 = 0 THEN 1 "
                'tmp_sql = tmp_sql & " 			WHEN yobi_bt4 = 1 THEN 0 "
                'tmp_sql = tmp_sql & " 		 END AS [「ｦ」→「ｵ」変換　0:しない　1:する] "
                'tmp_sql = tmp_sql & " 		,sosin_fname AS [送信ファイルパス名] "
                'tmp_sql = tmp_sql & " 		,jusin_fname AS [受信ファイルパス名] "
                'tmp_sql = tmp_sql & " 		,2 AS [2枚ＦＤ作成手順を使用する　0:しない　1:する] "
                'tmp_sql = tmp_sql & " 		,CASE "
                'tmp_sql = tmp_sql & " 			WHEN sq_hoho = 1 THEN 0 "
                'tmp_sql = tmp_sql & " 			WHEN sq_hoho = 2 THEN 1 "
                'tmp_sql = tmp_sql & " 		 END AS [請求のまとめ方　0:請求先単位　1:契約単位] "
                'tmp_sql = tmp_sql & " 		,sq_taino_umu AS [滞納分請求　0:しない　1:する] "
                'tmp_sql = tmp_sql & " 		,sq_tukitanni_umu AS [滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける] "
                'tmp_sql = tmp_sql & " 		,1 AS [手数料の請求入金管理をする　0:しない　1:する] "
                'tmp_sql = tmp_sql & " 		,0 AS [未入金の手数料は次回請求に加える　0:加えない　1:加える] "
                'tmp_sql = tmp_sql & " 		,1 AS [照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合] "
                'tmp_sql = tmp_sql & " 		,yobi_si1 AS [ゆうちょ銀行金融機関番号指定　不使用の場合はNULL] "
                'tmp_sql = tmp_sql & " 	FROM m_koza_furi AS MKF "
                'tmp_sql = tmp_sql & " 	LEFT JOIN "
                'tmp_sql = tmp_sql & " 		( "
                'tmp_sql = tmp_sql & " 			SELECT "
                'tmp_sql = tmp_sql & " 				 1 AS [自社支店]	/*紐付ツールが構築中のため1を設定しておく*/ "
                'tmp_sql = tmp_sql & " 				,ROW_NUMBER()OVER(ORDER BY kinyu_no,ten_no,kosyu_no,koza_no) AS [口座No] "
                'tmp_sql = tmp_sql & " 				,* "
                'tmp_sql = tmp_sql & " 			FROM "
                'tmp_sql = tmp_sql & " 			( "
                'tmp_sql = tmp_sql & " 				SELECT "
                'tmp_sql = tmp_sql & " 					 kinyu_no "
                'tmp_sql = tmp_sql & " 					,ten_no "
                'tmp_sql = tmp_sql & " 					,kosyu_no "
                'tmp_sql = tmp_sql & " 					,koza_no "
                'tmp_sql = tmp_sql & " 					,koza_meigi "
                'tmp_sql = tmp_sql & " 					,koza_kana "
                'tmp_sql = tmp_sql & " 					,'' AS post_kigo "
                'tmp_sql = tmp_sql & " 					,'' AS post_n "
                'tmp_sql = tmp_sql & " 					,'' AS post_bango "
                'tmp_sql = tmp_sql & " 					,'' AS biko "
                'tmp_sql = tmp_sql & " 				FROM m_furi_irai "
                'tmp_sql = tmp_sql & " 				UNION "
                'tmp_sql = tmp_sql & " 				SELECT "
                'tmp_sql = tmp_sql & " 					 kinyu_no "
                'tmp_sql = tmp_sql & " 					,ten_no "
                'tmp_sql = tmp_sql & " 					,kosyu_no "
                'tmp_sql = tmp_sql & " 					,koza_no "
                'tmp_sql = tmp_sql & " 					,koza_meigi "
                'tmp_sql = tmp_sql & " 					,koza_kana "
                'tmp_sql = tmp_sql & " 					,post_kigo "
                'tmp_sql = tmp_sql & " 					,post_n "
                'tmp_sql = tmp_sql & " 					,post_bango "
                'tmp_sql = tmp_sql & " 					,biko "
                'tmp_sql = tmp_sql & " 				FROM m_furi_koza "
                'tmp_sql = tmp_sql & " 				UNION "
                'tmp_sql = tmp_sql & " 				SELECT "
                'tmp_sql = tmp_sql & " 					 kinyu_no "
                'tmp_sql = tmp_sql & " 					,ten_no "
                'tmp_sql = tmp_sql & " 					,kosyu_no "
                'tmp_sql = tmp_sql & " 					,koza_no "
                'tmp_sql = tmp_sql & " 					,koza_meigi "
                'tmp_sql = tmp_sql & " 					,koza_kana "
                'tmp_sql = tmp_sql & " 					,'' AS post_kigo "
                'tmp_sql = tmp_sql & " 					,'' AS post_n "
                'tmp_sql = tmp_sql & " 					,'' AS post_bango "
                'tmp_sql = tmp_sql & " 					,'' AS biko "
                'tmp_sql = tmp_sql & " 				FROM m_koza_furi "
                'tmp_sql = tmp_sql & " 			) AS VW "
                'tmp_sql = tmp_sql & " 		) AS KOZATOTAL "
                'tmp_sql = tmp_sql & " 	ON  MKF.kinyu_no = KOZATOTAL.kinyu_no "
                'tmp_sql = tmp_sql & " 	AND MKF.ten_no = KOZATOTAL.ten_no "
                'tmp_sql = tmp_sql & " 	AND MKF.kosyu_no = KOZATOTAL.kosyu_no "
                'tmp_sql = tmp_sql & " 	AND MKF.koza_no = KOZATOTAL.koza_no "
                'tmp_sql = tmp_sql & " ) AS VW "

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[振替情報No],[振替情報名称],[振替情報カナ名称],[振込依頼人],[振込依頼人カナ名],[加盟店No 依頼人番号と共用],[ジェイリース　入金区分　1:変更なし・・・・],[金融機関No],[金融機関支店No],[口座種別],[口座番号],[口座名義],[引落日],[手数料　請求額],[備考],[データ並び替え　0:物件番号と部屋番号　1:契約者番号],[契約者Noを契約者番号に出力する  0:出力しない　1:出力する],[改行コード(CRLF)出力　0:なし　1:あり],[「ｦ」→「ｵ」変換　0:しない　1:する],[送信ファイルパス名],[受信ファイルパス名],[請求のまとめ方　0:請求先単位　1:契約単位],[滞納分請求　0:しない　1:する],[滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける],[手数料の請求入金管理をする　0:しない　1:する],[未入金の手数料は次回請求に加える　0:加えない　1:加える],[照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合],[ゆうちょ銀行金融機関番号指定　不使用の場合はNULL],[マルチヘッダー形式フラグ],[ゆうちょ銀行付加コード],[振替でのゆうちょ銀行コードに振込用変換を利用するフラグ],[振替入金処理時に受信ファイルを読み込まないフラグ],[再振替対応利用フラグ]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 fkae_no AS [振替情報No] "
                tmp_sql = tmp_sql & " 		,fkae_name AS [振替情報名称] "
                tmp_sql = tmp_sql & " 		,fkae_kana AS [振替情報カナ名称] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 口座振替情報 -del*/ "
                tmp_sql = tmp_sql & " 		/*,NULL AS [サービスタイプ　1:オリコ　2:ジャックス　3:アプラス]*/ "
                tmp_sql = tmp_sql & " 		,fkae_ircd AS [振込依頼人] "
                tmp_sql = tmp_sql & " 		,fkae_irkana AS [振込依頼人カナ名] "
                tmp_sql = tmp_sql & " 		,kamei_no AS [加盟店No 依頼人番号と共用] "
                tmp_sql = tmp_sql & " 		,NULL AS [ジェイリース　入金区分　1:変更なし・・・・] "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,[自社支店] AS [自社支店No] "
                tmp_sql = tmp_sql & " 		,[口座No] AS [振替先口座No] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,kinyu_no AS [金融機関No]		/*紐付けられた自社No自社口座No取得用*/ "
                tmp_sql = tmp_sql & " 		,ten_no AS [金融機関支店No]		/*紐付けられた自社No自社口座No取得用*/ "
                tmp_sql = tmp_sql & " 		,kosyu_name AS [口座種別]		/*紐付けられた自社No自社口座No取得用*/	/*20160609 紐付取得用に追加*/ "
                tmp_sql = tmp_sql & " 		,koza_no AS [口座番号]			/*紐付けられた自社No自社口座No取得用*/ "
                tmp_sql = tmp_sql & " 		,koza_meigi AS [口座名義]		/*紐付けられた自社No自社口座No取得用*/	/*20160609 紐付取得用に追加*/ "
                tmp_sql = tmp_sql & " 		,hiki1 AS [引落日] "
                tmp_sql = tmp_sql & " 		,fkae_tesu AS [手数料　請求額] "
                tmp_sql = tmp_sql & " 		,MKF.biko AS [備考] "
                tmp_sql = tmp_sql & " 		/*20160609 紐付設定値取得に伴う修正 del sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		全銀フォーマットマスタ移行後に再構築。とりあえず1を設定しておく "
                tmp_sql = tmp_sql & " 		,syudai_kbn AS [オリジナルフォーマットNo] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		,1 AS [オリジナルフォーマットNo] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		/*20160609 紐付設定値取得に伴う修正 del end*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN MKF.yobi_bt1 = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN MKF.yobi_bt1 = 1 THEN 2 "
                tmp_sql = tmp_sql & " 		 END AS [データ並び替え　0:物件番号と部屋番号　1:契約者番号] "
                tmp_sql = tmp_sql & " 		/*20160518 口座振替情報のフラグを反映させる修正 -chg sta*/ "
                tmp_sql = tmp_sql & " 		/*,1 AS [契約者Noを契約者番号に出力する  0:出力しない　1:出力する]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN MKF.yobi_bt2 = 1 THEN 0 "
                tmp_sql = tmp_sql & " 			ELSE 1 "
                tmp_sql = tmp_sql & " 		 END AS [契約者Noを契約者番号に出力する  0:出力しない　1:出力する] "
                tmp_sql = tmp_sql & " 		/*20160518 口座振替情報のフラグを反映させる修正 -chg end*/ "
                tmp_sql = tmp_sql & " 		,crlf AS [改行コード(CRLF)出力　0:なし　1:あり] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN MKF.yobi_bt4 = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN MKF.yobi_bt4 = 1 THEN 0 "
                tmp_sql = tmp_sql & " 		 END AS [「ｦ」→「ｵ」変換　0:しない　1:する] "
                tmp_sql = tmp_sql & " 		,sosin_fname AS [送信ファイルパス名] "
                tmp_sql = tmp_sql & " 		,jusin_fname AS [受信ファイルパス名] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 口座振替情報 -del*/ "
                tmp_sql = tmp_sql & " 		/*,2 AS [2枚ＦＤ作成手順を使用する　0:しない　1:する]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN sq_hoho = 1 THEN 0 "
                tmp_sql = tmp_sql & " 			WHEN sq_hoho = 2 THEN 1 "
                tmp_sql = tmp_sql & " 		 END AS [請求のまとめ方　0:請求先単位　1:契約単位] "
                tmp_sql = tmp_sql & " 		,sq_taino_umu AS [滞納分請求　0:しない　1:する] "
                tmp_sql = tmp_sql & " 		,sq_tukitanni_umu AS [滞納分は月単位で請求を分ける　0 or Null:分けない　1:分ける] "
                tmp_sql = tmp_sql & " 		,1 AS [手数料の請求入金管理をする　0:しない　1:する] "
                tmp_sql = tmp_sql & " 		,0 AS [未入金の手数料は次回請求に加える　0:加えない　1:加える] "
                tmp_sql = tmp_sql & " 		,1 AS [照合ルール　0:行番号で照合　1:口座情報とデータ処理用情報で照合] "
                tmp_sql = tmp_sql & " 		,MKF.yobi_si1 AS [ゆうちょ銀行金融機関番号指定　不使用の場合はNULL] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 口座振替情報 -add sta*/ "
                tmp_sql = tmp_sql & " 		,0 AS [マルチヘッダー形式フラグ] "
                tmp_sql = tmp_sql & " 		,NULL AS [ゆうちょ銀行付加コード] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN (SELECT yobi_bt3 FROM m_kan) = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE 2 "
                tmp_sql = tmp_sql & " 		 END AS [振替でのゆうちょ銀行コードに振込用変換を利用するフラグ] "
                tmp_sql = tmp_sql & " 		,2 AS [振替入金処理時に受信ファイルを読み込まないフラグ] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 口座振替情報 -add end*/ "
                tmp_sql = tmp_sql & " 		/*20161012 革命10アップデートに伴う修正 -add*/ "
                tmp_sql = tmp_sql & " 		,2 AS [再振替対応利用フラグ] "
                tmp_sql = tmp_sql & " 	FROM m_koza_furi AS MKF "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_kozasyu AS MK ON MKF.kosyu_no = MK.kosyu_no	/*20160609 紐付取得用に追加*/ "
                tmp_sql = tmp_sql & " ) AS VW "
                '2016.04.26 メインの方へも反映させる修正 -chg end

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

#Region "入出金取得情報"   '20160517 入出金取得情報の新規作成

    Public Class M_ns_syutoku_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[入出金取得No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[入出金取得No],[設定名称],[設定カナ名称],[フォーマット定義名],[受信ファイル名（フルパス）],[取込データ種別],[備考],[家賃口座No],[送信ファイル名（フルパス）],[預金分割設定],[重複チェック]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 ns_no AS [入出金取得No] "
                tmp_sql = tmp_sql & " 		,ns_name AS [設定名称] "
                tmp_sql = tmp_sql & " 		,ns_kana AS [設定カナ名称] "
                tmp_sql = tmp_sql & " 		/*,'' AS [自社支店No]*/ "
                tmp_sql = tmp_sql & " 		/*,'' AS [フォーマット定義No.]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 1 THEN '振込入金通知' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 3 THEN '全銀入出金明細' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 103 THEN '全銀入出金明細(口座振替)' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 101 THEN 'ANSER-SPC(CSV)' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 21 THEN '東京三菱(BzStation)「全明細」(CSV)' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 401 THEN 'タイプA(CSV)' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 402 THEN 'タイプB(CSV)' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 403 THEN 'タイプC-1(CSV)' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 405 THEN 'タイプC-2(CSV)' "
                tmp_sql = tmp_sql & " 			WHEN data_syubetu = 404 THEN 'タイプD(CSV)' "
                tmp_sql = tmp_sql & " 			ELSE '' "
                tmp_sql = tmp_sql & " 		 END AS [フォーマット定義名] "
                tmp_sql = tmp_sql & " 		,jusin_fname AS [受信ファイル名（フルパス）] "
                tmp_sql = tmp_sql & " 		,'' AS [取込データ種別] "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 入出金取得情報 del sta*/ "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN crlf = 1 THEN 1 "
                tmp_sql = tmp_sql & " 			WHEN crlf = 2 THEN 2 "
                tmp_sql = tmp_sql & " 		 END AS [改行コード有無] "
                tmp_sql = tmp_sql & " 		*/ "
                tmp_sql = tmp_sql & " 		/*20160519 EXEUpdateに伴う修正 入出金取得情報 del end*/ "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 		,'' AS [家賃口座No] "
                tmp_sql = tmp_sql & " 		,'' AS [送信ファイル名（フルパス）] "
                tmp_sql = tmp_sql & " 		,'' AS [預金分割設定] "
                tmp_sql = tmp_sql & " 		,0 AS [重複チェック] "
                tmp_sql = tmp_sql & " 	FROM m_ns_syutoku "
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

#Region "家賃入金口座情報"   '20160517 家賃入金口座情報の新規作成

    Public Class M_furi_koza_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[家賃入金受付口座No],[家主No],[家主口座No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[家賃入金受付口座No],[受取人名],[金融機関No],[金融機関支店No],[口座種別],[口座番号],[受取人名Unicode],[受取人カナ],[請求先区分],[自社支店No],[自社口座No],[家主No],[家主口座No],[ANSER-SPC有無],[ANSER-SPC口座指定方式],[ANSER-SPC種目付加方法],[ANSER-SPC任意番号],[ANSER-SPC加入者番号],[ANSER-SPC接続先個別設定フラグ],[ANSER-SPC接続方法],[ANSER-SPCエリアNo],[ANSER-SPC地区No],[ANSER-SPCTEL],[ANSER-SPC照会用暗証番号],[ANSER-SPCサービスコード],[家賃入金口座備考],[ヘッダーテンプレート],[トレーラーテンプレート]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 fkom_no AS [家賃入金受付口座No] "
                tmp_sql = tmp_sql & " 		,fkom_name AS [受取人名] "
                tmp_sql = tmp_sql & " 		/*20160609 家賃入金口座情報を紐付設定ファイルから取得する修正 add sta*/ "
                tmp_sql = tmp_sql & " 		,kinyu_no AS [金融機関No] "
                tmp_sql = tmp_sql & " 		,ten_no AS [金融機関支店No] "
                tmp_sql = tmp_sql & " 		,kosyu_name AS [口座種別] "
                tmp_sql = tmp_sql & " 		,koza_no AS [口座番号] "
                tmp_sql = tmp_sql & " 		/*20160609 家賃入金口座情報を紐付設定ファイルから取得する修正 add end*/ "
                tmp_sql = tmp_sql & " 		,'' AS [受取人名Unicode] "
                tmp_sql = tmp_sql & " 		,fkom_kana AS [受取人カナ] "
                tmp_sql = tmp_sql & " 		/*20160609 家賃入金口座情報を紐付設定ファイルから取得する修正 chg sta*/ "
                tmp_sql = tmp_sql & " 		/*,200 AS [請求先区分]*/ "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN ISNULL(so_no,'') = '' OR ISNULL(so_no,'') = '' THEN 900 "
                tmp_sql = tmp_sql & " 			ELSE 200 "
                tmp_sql = tmp_sql & " 		 END AS [請求先区分] "
                tmp_sql = tmp_sql & " 		/*20160609 家賃入金口座情報を紐付設定ファイルから取得する修正 chg end*/ "
                tmp_sql = tmp_sql & " 		,NULL AS [自社支店No] "
                tmp_sql = tmp_sql & " 		,NULL AS [自社口座No] "
                tmp_sql = tmp_sql & " 		,so_no AS [家主No] "
                tmp_sql = tmp_sql & " 		,sokoza_no AS [家主口座No] "
                tmp_sql = tmp_sql & " 		,spc_umu AS [ANSER-SPC有無] "
                tmp_sql = tmp_sql & " 		,spc_siteihoho AS [ANSER-SPC口座指定方式] "
                tmp_sql = tmp_sql & " 		,spc_fukahoho AS [ANSER-SPC種目付加方法] "
                tmp_sql = tmp_sql & " 		,TOTAL.yobi_si1 AS [ANSER-SPC任意番号] "
                tmp_sql = tmp_sql & " 		,spc_kanyu_no AS [ANSER-SPC加入者番号] "
                tmp_sql = tmp_sql & " 		,spc_setuzoku_umu AS [ANSER-SPC接続先個別設定フラグ] "
                tmp_sql = tmp_sql & " 		,spc_setuzoku_hoho AS [ANSER-SPC接続方法] "
                tmp_sql = tmp_sql & " 		,spc_area_no AS [ANSER-SPCエリアNo] "
                tmp_sql = tmp_sql & " 		,spc_tiku_no AS [ANSER-SPC地区No] "
                tmp_sql = tmp_sql & " 		,spc_tel AS [ANSER-SPCTEL] "
                tmp_sql = tmp_sql & " 		,spc_pwd AS [ANSER-SPC照会用暗証番号] "
                tmp_sql = tmp_sql & " 		,spc_svcode AS [ANSER-SPCサービスコード] "
                tmp_sql = tmp_sql & " 		,'' AS [家賃入金口座備考] "
                tmp_sql = tmp_sql & " 		,'' AS [ヘッダーテンプレート] "
                tmp_sql = tmp_sql & " 		,'' AS [トレーラーテンプレート] "
                tmp_sql = tmp_sql & " 	FROM "
                tmp_sql = tmp_sql & " 	( "
                tmp_sql = tmp_sql & " 		SELECT "
                tmp_sql = tmp_sql & " 			 so_no "
                tmp_sql = tmp_sql & " 			,sokoza_no "
                tmp_sql = tmp_sql & " 			,MFK.* "
                tmp_sql = tmp_sql & " 		FROM m_furi_koza AS MFK "
                tmp_sql = tmp_sql & " 		LEFT JOIN m_yanu_so_koza OWKOZA "
                tmp_sql = tmp_sql & " 		ON  MFK.kinyu_no = OWKOZA.kinyu_no "
                tmp_sql = tmp_sql & " 		AND MFK.ten_no = OWKOZA.ten_no "
                tmp_sql = tmp_sql & " 		AND MFK.kosyu_no = OWKOZA.kosyu_no "
                tmp_sql = tmp_sql & " 		AND MFK.koza_no = OWKOZA.koza_no "
                tmp_sql = tmp_sql & " 	) AS TOTAL "
                tmp_sql = tmp_sql & " 	/*20160609 家賃入金口座情報を紐付設定ファイルから取得する修正 add*/ "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_kozasyu AS MK ON TOTAL.kosyu_no = MK.kosyu_no "
                tmp_sql = tmp_sql & " 	/*20160609 家賃入金口座情報を紐付設定ファイルから取得する修正 del*/ "
                tmp_sql = tmp_sql & " 	/*WHERE so_no IS NOT NULL AND sokoza_no IS NOT NULL*/ "
                tmp_sql = tmp_sql & " 	/*20160829 自社口座にデフォルト値を設定する処理を追加 del sta*/ "
                tmp_sql = tmp_sql & " 	/* "
                tmp_sql = tmp_sql & " 	/*20160722 口座情報が存在しないことによるエラーの修正 応急処置 add*/ "
                tmp_sql = tmp_sql & " 	WHERE ISNULL(kinyu_no,0) <> 0 AND ISNULL(ten_no,0) <> 0 "
                tmp_sql = tmp_sql & " 	*/ "
                tmp_sql = tmp_sql & " 	/*20160829 自社口座にデフォルト値を設定する処理を追加 del end*/ "
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

#Region "ANSERエリア情報"   '20160704 ANSER情報の移行処理追加

    Public Class Mspc_area_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[SPCエリアNo]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[SPCエリアNo],[SPCエリア名称],[備考]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 spc_area_no AS [SPCエリアNo] "
                tmp_sql = tmp_sql & " 		,spc_areaname AS [SPCエリア名称] "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 	FROM mspc_area "
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

#Region "ANSERアクセスポイント情報"   '20160704 ANSER情報の移行処理追加

    Public Class Mspc_accpoint_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[SPC地区No],[SPCエリアNo]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[SPC地区No],[SPCエリアNo],[SPC地区名称],[備考]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 spc_tiku_no AS [SPC地区No] "
                tmp_sql = tmp_sql & " 		,spc_area_no AS [SPCエリアNo] "
                tmp_sql = tmp_sql & " 		,spc_tikuname AS [SPC地区名称] "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 	FROM mspc_accpoint "
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

#Region "ANSER接続情報"   '20160704 ANSER情報の移行処理追加

    Public Class Mspc_setuzoku_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[接続方法No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[接続方法No],[設定名],[回線種別(1:電話回線 2:ISDN)],[開発タイプ],[開発名],[接続方法(1:アクセスポイント 2:TEL)],[エリアNo],[地区No],[電話番号],[外線No],[リトライ回数],[リトライの間隔],[利用するサービスコード],[備考],[改行有無]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 spc_setuzoku_no AS [接続方法No] "
                tmp_sql = tmp_sql & " 		,yobi_mj1 AS [設定名] "
                tmp_sql = tmp_sql & " 		,spc_kaisen_syu AS [回線種別(1:電話回線 2:ISDN)] "
                tmp_sql = tmp_sql & " 		,spc_devtype AS [開発タイプ] "
                tmp_sql = tmp_sql & " 		,spc_devname AS [開発名] "
                tmp_sql = tmp_sql & " 		,spc_setuzoku_hoho AS [接続方法(1:アクセスポイント 2:TEL)] "
                tmp_sql = tmp_sql & " 		,spc_area_no AS [エリアNo] "
                tmp_sql = tmp_sql & " 		,spc_tiku_no AS [地区No] "
                tmp_sql = tmp_sql & " 		,spc_tel AS [電話番号] "
                tmp_sql = tmp_sql & " 		,spc_gaisen AS [外線No] "
                tmp_sql = tmp_sql & " 		,spc_r_kaisu AS [リトライ回数] "
                tmp_sql = tmp_sql & " 		,spc_r_kankaku AS [リトライの間隔] "
                tmp_sql = tmp_sql & " 		,yobi_si1 AS [利用するサービスコード] "
                tmp_sql = tmp_sql & " 		,biko AS [備考] "
                tmp_sql = tmp_sql & " 		,1 AS [改行有無] "
                tmp_sql = tmp_sql & " 	FROM mspc_setuzoku "
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
