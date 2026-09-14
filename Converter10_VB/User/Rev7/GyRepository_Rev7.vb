Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "仲介業者基本情報"

    Public Class M_gy_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[業者No],[業者名],[業者名(SJIS)],[業者名カナ],[郵便番号],[住所1],[住所2],[TEL1],[TEL2],[FAX],[携帯1],[携帯2],[代表者役職],[代表者名],[担当者部署役職],[担当者名],[担当者名(SJIS)],[担当者カナ],[メールアドレス],[携帯メールアドレス],[Webアドレス(URL)],[優先設定(電話番号)],[優先設定(メールアドレス)],[備考(基本情報)],[口座振込通知書発行の許可],[免許証番号],[免許年月日],[取引主任者登録番号],[取引主任者名],[全国賃貸不動産管理業協会会員番号],[賃貸不動産経営管理士登録番号],[賃貸不動産経営管理士名],[賃貸住宅管理業者登録番号(仮)],[代表者名(SJIS)],[取引主任者名(SJIS)],[仲介業者契約担当者自動入力フラグ],[賃貸不動産経営管理士名(SJIS)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 gy_no AS [業者No] "
                tmp_sql = tmp_sql & " 		,gy_name AS [業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [業者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,gy_kana AS [業者名カナ] "
                tmp_sql = tmp_sql & " 		,post AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,daihyo_yaku AS [代表者役職] "
                tmp_sql = tmp_sql & " 		,daihyo_name AS [代表者名] "
                tmp_sql = tmp_sql & " 		,busyo AS [担当者部署役職] "
                tmp_sql = tmp_sql & " 		,tanto AS [担当者名] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者カナ] "
                tmp_sql = tmp_sql & " 		,email AS [メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 		,url AS [Webアドレス(URL)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(メールアドレス)] "
                tmp_sql = tmp_sql & " 		,'' AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		,1 AS [口座振込通知書発行の許可] "
                tmp_sql = tmp_sql & " 		,menkyo_no AS [免許証番号] "
                tmp_sql = tmp_sql & " 		,menkyo_ymd AS [免許年月日] "
                tmp_sql = tmp_sql & " 		,torihiki_no AS [取引主任者登録番号] "
                tmp_sql = tmp_sql & " 		,torihiki_name AS [取引主任者名] "
                tmp_sql = tmp_sql & " 		,'' AS [全国賃貸不動産管理業協会会員番号] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸不動産経営管理士登録番号] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸不動産経営管理士名] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸住宅管理業者登録番号(仮)] "
                tmp_sql = tmp_sql & " 		,'' AS [代表者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,'' AS [取引主任者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,0 AS [仲介業者契約担当者自動入力フラグ] "
                tmp_sql = tmp_sql & " 		,'' AS [賃貸不動産経営管理士名(SJIS)] "
                tmp_sql = tmp_sql & " 	FROM m_gy "
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

#Region "仲介業者口座情報"

    Public Class M_gy_Koza_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[業者No],[金融機関No],[金融機関支店No],[口座種別],[口座番号],[口座名義],[口座名義カナ],[備考(口座情報)],[備考(口座振込情報)],[総合振込区分],[振込依頼人No],[振込手数料負担区分],[振込手数料計算区分],[振込手数料固定額1],[振込手数料固定額2],[備考(総合振込情報)],[ゆうちょ銀行記号1],[ゆうちょ銀行記号2],[ゆうちょ銀行番号]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 gy_no AS [業者No] "
                tmp_sql = tmp_sql & " 		,kinyu_no AS [金融機関No] "
                tmp_sql = tmp_sql & " 		,ten_no AS [金融機関支店No] "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN kosyu_no = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE kosyu_no "
                tmp_sql = tmp_sql & " 		 END AS [口座種別] "
                tmp_sql = tmp_sql & " 		 */ "
                tmp_sql = tmp_sql & " 		,kosyu_name AS [口座種別] "
                tmp_sql = tmp_sql & " 		,koza_no AS [口座番号] "
                tmp_sql = tmp_sql & " 		,koza_meigi AS [口座名義] "
                tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
                tmp_sql = tmp_sql & " 		,MG.biko AS [備考(口座情報)] "
                tmp_sql = tmp_sql & " 		,'' AS [備考(口座振込情報)] "
                tmp_sql = tmp_sql & " 		,1 AS [総合振込区分] "
                tmp_sql = tmp_sql & " 		,'' AS [振込依頼人No] "
                tmp_sql = tmp_sql & " 		,2 AS [振込手数料負担区分] "
                tmp_sql = tmp_sql & " 		,1 AS [振込手数料計算区分] "
                tmp_sql = tmp_sql & " 		,'' AS [振込手数料固定額1] "
                tmp_sql = tmp_sql & " 		,'' AS [振込手数料固定額2] "
                tmp_sql = tmp_sql & " 		,'' AS [備考(総合振込情報)] "
                tmp_sql = tmp_sql & " 		,'' AS [ゆうちょ銀行記号1] "
                tmp_sql = tmp_sql & " 		,'' AS [ゆうちょ銀行記号2] "
                tmp_sql = tmp_sql & " 		,'' AS [ゆうちょ銀行番号] "
                tmp_sql = tmp_sql & " 	FROM m_gy AS MG "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_kozasyu AS MK ON MG.kosyu_no = MK.kosyu_no "
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

#Region "仲介業者メモ情報(汎用ツールのみ)"

    Public Class M_gy_Biko_Repository

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

#Region "保険業者基本情報"

    Public Class M_hoken_gy_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[業者No],[業者名],[業者名(SJIS)],[業者名カナ],[郵便番号],[住所1],[住所2],[TEL1],[TEL2],[FAX],[携帯1],[携帯2],[担当者部署役職],[担当者名],[担当者名(SJIS)],[担当者名カナ],[メールアドレス],[携帯メールアドレス],[Webアドレス(URL)],[優先設定(電話番号)],[優先設定(メールアドレス)],[備考(基本情報)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 hkngy_no AS [業者No] "
                tmp_sql = tmp_sql & " 		,hkngy_name AS [業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [業者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,hkngy_kana AS [業者名カナ] "
                tmp_sql = tmp_sql & " 		,post AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,busyo AS [担当者部署役職] "
                tmp_sql = tmp_sql & " 		,tanto AS [担当者名] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名カナ] "
                tmp_sql = tmp_sql & " 		,email AS [メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [Webアドレス(URL)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(メールアドレス)] "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN yobi_si1 <> '' THEN '契約期間：' + CONVERT(VARCHAR,yobi_si1) + '年' "
                tmp_sql = tmp_sql & " 			WHEN yobi_si1 = '' THEN '' "
                tmp_sql = tmp_sql & " 		 END AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 	FROM m_hoken_gy "
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

#Region "保険業者口座情報(汎用ツールのみ)"

    Public Class M_hoken_gy_Koza_Repository

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

#Region "保険業者メモ情報(汎用ツールのみ)"

    Public Class M_hoken_gy_Biko_Repository

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

#Region "家賃保証業者情報"

    Public Class M_hosyo_gy_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[業者No],[業者名],[業者名(shift-jis)],[業者名カナ],[郵便番号],[住所1],[住所2],[TEL1],[TEL2],[FAX],[携帯1],[携帯2],[担当者部署役職],[担当者名],[担当者名(SJIS)],[担当者名カナ],[メールアドレス],[携帯メールアドレス],[Webアドレス(URL)],[優先設定(電話番号)],[優先設定(メールアドレス)],[備考(基本情報)],[連動用家賃保証会社区分]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 hkngy_no AS [業者No] "
                tmp_sql = tmp_sql & " 		,hkngy_name AS [業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [業者名(shift-jis)] "
                tmp_sql = tmp_sql & " 		,hkngy_kana AS [業者名カナ] "
                tmp_sql = tmp_sql & " 		,post AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,busyo AS [担当者部署役職] "
                tmp_sql = tmp_sql & " 		,tanto AS [担当者名] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名カナ] "
                tmp_sql = tmp_sql & " 		,email AS [メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [Webアドレス(URL)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(メールアドレス)] "
                tmp_sql = tmp_sql & " 		,'' AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		,-1 AS [連動用家賃保証会社区分] "
                tmp_sql = tmp_sql & " 	FROM m_hosyo_gy "
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

#Region "家賃保証業者口座情報"

    '革命10に過去にあったが現在(2015/10/23)は無くなっている

#End Region

#Region "家賃保証業者メモ情報(汎用ツールのみ)"

    Public Class M_hosyo_gy_Biko_Repository

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

#Region "修繕業者情報"

    Public Class M_gysyuzen_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[業者No],[業者名],[業者名(SJIS)],[業者名カナ],[郵便番号],[住所1],[住所2],[TEL1],[TEL2],[FAX],[携帯1],[携帯2],[代表者役職],[代表者名],[担当者部署役職],[担当者名],[担当者名(SJIS)],[担当者名カナ],[メールアドレス],[携帯メールアドレス],[Webアドレス(URL)],[優先設定(電話番号)],[優先設定(メールアドレス)],[備考(基本情報)],[代表者名(SJIS)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 gy_no AS [業者No] "
                tmp_sql = tmp_sql & " 		,gy_name AS [業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [業者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,gy_kana AS [業者名カナ] "
                tmp_sql = tmp_sql & " 		,post AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,daihyo_yaku AS [代表者役職] "
                tmp_sql = tmp_sql & " 		,daihyo_name AS [代表者名] "
                tmp_sql = tmp_sql & " 		,busyo AS [担当者部署役職] "
                tmp_sql = tmp_sql & " 		,tanto AS [担当者名] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名カナ] "
                tmp_sql = tmp_sql & " 		,email AS [メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 		,url AS [Webアドレス(URL)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(メールアドレス)] "
                tmp_sql = tmp_sql & " 		,'' AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		,'' AS [代表者名(SJIS)] "
                tmp_sql = tmp_sql & " 	FROM m_gysyuzen "
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

#Region "修繕業者口座情報"

    Public Class M_gysyuzen_Koza_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[業者No],[金融機関No],[金融機関店No],[口座種別],[口座番号],[口座名義],[口座名義カナ],[ゆうちょ銀行記号1],[ゆうちょ銀行記号2],[ゆうちょ銀行番号],[備考(口座情報)],[備考(口座振込情報)],[総合振込区分],[振込依頼人No],[振込手数料負担区分],[振込手数料計算区分],[振込手数料固定額1],[振込手数料固定額2],[備考(総合振込情報)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 gy_no AS [業者No] "
                tmp_sql = tmp_sql & " 		,kinyu_no AS [金融機関No] "
                tmp_sql = tmp_sql & " 		,ten_no AS [金融機関店No] "
                tmp_sql = tmp_sql & " 		/* "
                tmp_sql = tmp_sql & " 		,CASE "
                tmp_sql = tmp_sql & " 			WHEN kosyu_no = 0 THEN 1 "
                tmp_sql = tmp_sql & " 			ELSE kosyu_no "
                tmp_sql = tmp_sql & " 		 END AS [口座種別] "
                tmp_sql = tmp_sql & " 		 */ "
                tmp_sql = tmp_sql & " 		,kosyu_name AS [口座種別] "
                tmp_sql = tmp_sql & " 		,koza_no AS [口座番号] "
                tmp_sql = tmp_sql & " 		,koza_meigi AS [口座名義] "
                tmp_sql = tmp_sql & " 		,koza_kana AS [口座名義カナ] "
                tmp_sql = tmp_sql & " 		,'' AS [ゆうちょ銀行記号1] "
                tmp_sql = tmp_sql & " 		,'' AS [ゆうちょ銀行記号2] "
                tmp_sql = tmp_sql & " 		,'' AS [ゆうちょ銀行番号] "
                tmp_sql = tmp_sql & " 		,MG.biko AS [備考(口座情報)] "
                tmp_sql = tmp_sql & " 		,'' AS [備考(口座振込情報)] "
                tmp_sql = tmp_sql & " 		,1 AS [総合振込区分] "
                tmp_sql = tmp_sql & " 		,'' AS [振込依頼人No] "
                tmp_sql = tmp_sql & " 		,2 AS [振込手数料負担区分] "
                tmp_sql = tmp_sql & " 		,1 AS [振込手数料計算区分] "
                tmp_sql = tmp_sql & " 		,'' AS [振込手数料固定額1] "
                tmp_sql = tmp_sql & " 		,'' AS [振込手数料固定額2] "
                tmp_sql = tmp_sql & " 		,'' AS [備考(総合振込情報)] "
                tmp_sql = tmp_sql & " 	FROM m_gysyuzen AS MG "
                tmp_sql = tmp_sql & " 	LEFT JOIN m_kozasyu AS MK ON MG.kosyu_no = MK.kosyu_no "
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

#Region "修繕業者メモ情報(汎用ツールのみ)"

    Public Class M_gysyuzen_Biko_Repository

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

#Region "ライフライン業者情報"

    Public Class M_kokyo_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[ライフライン業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[ライフライン業者No],[ライフライン業者名],[ライフライン業者名(Shift-jis)],[ライフライン業者名カナ],[営業所],[営業所(Shift-jis)],[郵便番号],[住所1],[住所2],[TEL1],[TEL2],[FAX],[携帯1],[携帯2],[優先設定(電話番号)],[備考(基本情報)],[電気],[上水道],[ガス],[灯油],[その他],[排水]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 kokyo_no AS [ライフライン業者No] "
                tmp_sql = tmp_sql & " 		,kokyo_name AS [ライフライン業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [ライフライン業者名(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,kokyo_kana AS [ライフライン業者名カナ] "
                tmp_sql = tmp_sql & " 		,kokyo_einame AS [営業所] "
                tmp_sql = tmp_sql & " 		,'' AS [営業所(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,'' AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,biko AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		,1 AS [電気] "
                tmp_sql = tmp_sql & " 		,0 AS [上水道] "
                tmp_sql = tmp_sql & " 		,0 AS [ガス] "
                tmp_sql = tmp_sql & " 		,0 AS [灯油] "
                tmp_sql = tmp_sql & " 		,0 AS [その他] "
                tmp_sql = tmp_sql & " 		,0 AS [排水] "
                tmp_sql = tmp_sql & " 	FROM m_kokyo "
                tmp_sql = tmp_sql & " 	WHERE kokyo_rui = 1 "
                tmp_sql = tmp_sql & " ) AS DENKI "
                tmp_sql = tmp_sql & " UNION "
                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 kokyo_no AS [ライフライン業者No] "
                tmp_sql = tmp_sql & " 		,kokyo_name AS [ライフライン業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [ライフライン業者名(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,kokyo_kana AS [ライフライン業者名カナ] "
                tmp_sql = tmp_sql & " 		,kokyo_einame AS [営業所] "
                tmp_sql = tmp_sql & " 		,'' AS [営業所(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,'' AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,biko AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		,0 AS [電気] "
                tmp_sql = tmp_sql & " 		,1 AS [上水道] "
                tmp_sql = tmp_sql & " 		,0 AS [ガス] "
                tmp_sql = tmp_sql & " 		,0 AS [灯油] "
                tmp_sql = tmp_sql & " 		,0 AS [その他] "
                tmp_sql = tmp_sql & " 		,0 AS [排水] "
                tmp_sql = tmp_sql & " 	FROM m_kokyo "
                tmp_sql = tmp_sql & " 	WHERE kokyo_rui = 2 "
                tmp_sql = tmp_sql & " ) AS SUIDO "
                tmp_sql = tmp_sql & " UNION "
                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 kokyo_no AS [ライフライン業者No] "
                tmp_sql = tmp_sql & " 		,kokyo_name AS [ライフライン業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [ライフライン業者名(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,kokyo_kana AS [ライフライン業者名カナ] "
                tmp_sql = tmp_sql & " 		,kokyo_einame AS [営業所] "
                tmp_sql = tmp_sql & " 		,'' AS [営業所(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,'' AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,biko AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		,0 AS [電気] "
                tmp_sql = tmp_sql & " 		,0 AS [上水道] "
                tmp_sql = tmp_sql & " 		,1 AS [ガス] "
                tmp_sql = tmp_sql & " 		,0 AS [灯油] "
                tmp_sql = tmp_sql & " 		,0 AS [その他] "
                tmp_sql = tmp_sql & " 		,0 AS [排水] "
                tmp_sql = tmp_sql & " 	FROM m_kokyo "
                tmp_sql = tmp_sql & " 	WHERE kokyo_rui = 3 "
                tmp_sql = tmp_sql & " ) AS GASU "
                tmp_sql = tmp_sql & " UNION "
                tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 kokyo_no AS [ライフライン業者No] "
                tmp_sql = tmp_sql & " 		,kokyo_name AS [ライフライン業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [ライフライン業者名(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,kokyo_kana AS [ライフライン業者名カナ] "
                tmp_sql = tmp_sql & " 		,kokyo_einame AS [営業所] "
                tmp_sql = tmp_sql & " 		,'' AS [営業所(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,'' AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,biko AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 		,0 AS [電気] "
                tmp_sql = tmp_sql & " 		,0 AS [上水道] "
                tmp_sql = tmp_sql & " 		,0 AS [ガス] "
                tmp_sql = tmp_sql & " 		,0 AS [灯油] "
                tmp_sql = tmp_sql & " 		,1 AS [その他] "
                tmp_sql = tmp_sql & " 		,0 AS [排水] "
                tmp_sql = tmp_sql & " 	FROM m_kokyo "
                tmp_sql = tmp_sql & " 	WHERE kokyo_rui NOT IN (1,2,3) "
                tmp_sql = tmp_sql & " ) AS OTHER "

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

#Region "施工業者情報"

    Public Class M_seko_kaisya_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[施工業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[施工業者No],[施工業者名],[施工業者名(Shift-jis)],[施工業者名カナ],[営業所],[営業所(Shift-jis)],[郵便番号],[住所1],[住所2],[TEL1],[TEL2],[FAX],[携帯1],[携帯2],[担当者部署役職],[担当者名],[担当者名(SJIS)],[担当者カナ],[メールアドレス],[携帯メールアドレス],[Webアドレス(URL)],[優先設定(電話番号)],[優先設定(メールアドレス)],[備考(基本情報)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 kaisya_no AS [施工業者No] "
                tmp_sql = tmp_sql & " 		,kname AS [施工業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [施工業者名(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,kkana AS [施工業者名カナ] "
                tmp_sql = tmp_sql & " 		,eigyosyo AS [営業所] "
                tmp_sql = tmp_sql & " 		,'' AS [営業所(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,post_code AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,tanto_busyo AS [担当者部署役職] "
                tmp_sql = tmp_sql & " 		,tanto_name AS [担当者名] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者カナ] "
                tmp_sql = tmp_sql & " 		,'' AS [メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [Webアドレス(URL)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(メールアドレス)] "
                tmp_sql = tmp_sql & " 		,biko AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 	FROM m_seko_kaisya "
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

#Region "施設保守業者情報"

    Public Class M_hosyu_kaisya_Repository

        Public Class SubConv
            Implements IConv

            ''' <summary>
            ''' 【抽出クエリ】
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Function Get_UseQry(ByRef sortstr As String) As String Implements IConv.Get_UseQry

                Dim tmp_sql As String = ""

                sortstr = "[施設保守業者No]"

                '抽出データの改行文字列を除去する暫定処理
                Dim tmp_midheader As String = "[施設保守業者No],[施設保守業者名],[施設保守業者名(Shift-jis)],[施設保守業者名カナ],[営業所],[営業所(Shift-jis)],[郵便番号],[住所1],[住所2],[TEL1],[TEL2],[FAX],[携帯1],[携帯2],[担当者部署役職],[担当者名],[担当者名(SJIS)],[担当者カナ],[メールアドレス],[携帯メールアドレス],[Webアドレス(URL)],[優先設定(電話番号)],[優先設定(メールアドレス)],[備考(基本情報)]"
                Dim tmp_chgmidheader As String = EtcMethod.Get_FldnameAddReplaceMethod(tmp_midheader, True)
                '抽出データの改行文字列を除去する暫定処理

                'tmp_sql = tmp_sql & " SELECT * FROM "
                tmp_sql = tmp_sql & " SELECT " & tmp_chgmidheader & " FROM "
                tmp_sql = tmp_sql & " ( "
                tmp_sql = tmp_sql & " 	SELECT "
                tmp_sql = tmp_sql & " 		 kaisya_no AS [施設保守業者No] "
                tmp_sql = tmp_sql & " 		,kname AS [施設保守業者名] "
                tmp_sql = tmp_sql & " 		,'' AS [施設保守業者名(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,kkana AS [施設保守業者名カナ] "
                tmp_sql = tmp_sql & " 		,eigyosyo AS [営業所] "
                tmp_sql = tmp_sql & " 		,'' AS [営業所(Shift-jis)] "
                tmp_sql = tmp_sql & " 		,post_code AS [郵便番号] "
                tmp_sql = tmp_sql & " 		,add1 AS [住所1] "
                tmp_sql = tmp_sql & " 		,add2 AS [住所2] "
                tmp_sql = tmp_sql & " 		,tel1 AS [TEL1] "
                tmp_sql = tmp_sql & " 		,tel2 AS [TEL2] "
                tmp_sql = tmp_sql & " 		,fax AS [FAX] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯1] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯2] "
                tmp_sql = tmp_sql & " 		,tanto_busyo AS [担当者部署役職] "
                tmp_sql = tmp_sql & " 		,tanto_name AS [担当者名] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者名(SJIS)] "
                tmp_sql = tmp_sql & " 		,'' AS [担当者カナ] "
                tmp_sql = tmp_sql & " 		,'' AS [メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [携帯メールアドレス] "
                tmp_sql = tmp_sql & " 		,'' AS [Webアドレス(URL)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(電話番号)] "
                tmp_sql = tmp_sql & " 		,1 AS [優先設定(メールアドレス)] "
                tmp_sql = tmp_sql & " 		,biko AS [備考(基本情報)] "
                tmp_sql = tmp_sql & " 	FROM m_hosyu_kaisya "
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
