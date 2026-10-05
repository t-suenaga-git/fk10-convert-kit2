Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "設備マスタ取得"

    Public Class M_setubi_Rev7_Repository

        '20160525 全体的な動作の修正 -del sta
        'V7設備情報はDBから直接読み込むため処理を変更
        ' ''' <summary>
        ' ''' V7設備取得
        ' ''' </summary>
        ' ''' <remarks></remarks>
        'Public Function ReadMid() As Boolean

        '    Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
        '    Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
        '    Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
        '    Dim startrow As Integer                                         '書込開始行
        '    Dim columncnt As Integer                                        '列数
        '    Dim maxrowcnt As Integer                                        '既存データの行数
        '    Dim rowcnt As Integer                                           '書込行数
        '    Dim rtn As Boolean = True
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim list_relitem As New List(Of String)

        '    '20160413 test
        '    'Dim filename As String = "部屋情報"
        '    'Dim sheetname As String = "部屋設備情報"
        '    Dim filename As String = "紐付テスト用ファイル"
        '    Dim sheetname As String = "設備マスタ"


        '    '20160413 test
        '    Dim relfldno As Integer = 0

        '    '************************
        '    '作業準備
        '    '************************

        '    'Excelファイル初期設定
        '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MidDirPath, filename, sheetname)

        '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '    If Not rtn Then
        '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160525 全体的な動作の修正 -add
        '        Return rtn
        '    End If

        '    '************************
        '    '処理開始
        '    '************************

        '    Dim headerrow As New Object
        '    Dim relcolvalue As New Object

        '    'ヘッダー行取得
        '    headerrow = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

        '    '取得した列のデータをリストへ格納 (重複集約)
        '    For cntii = 1 To rowcnt - 1

        '        '行取得
        '        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

        '        '作業用変数
        '        Dim tmp_setubikomkname As String = ""
        '        Dim tmp_setubinaiyo As String = ""
        '        Dim tmp_total As String = ""

        '        For cntjj = 1 To columncnt

        '            Dim fldname As String = headerrow(startrow - 1, cntjj).ToString.Trim
        '            Dim fldvalue As String = ""
        '            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
        '                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
        '            End If

        '            Select Case fldname
        '                Case "設備項目名"
        '                    tmp_setubikomkname = fldvalue
        '                Case "設備内容"
        '                    tmp_setubinaiyo = fldvalue
        '            End Select

        '        Next

        '        tmp_total = tmp_setubikomkname & "-" & tmp_setubinaiyo

        '        If (tmp_setubikomkname <> "" And tmp_setubinaiyo <> "") And list_relitem.Contains(tmp_total) = False Then
        '            list_relitem.Add(tmp_total)
        '        End If

        '    Next

        '    '集約した紐付データをモデルへ格納
        '    Dim relcnt As Integer = 1
        '    Dim itemcnt As Integer = list_relitem.Count
        '    Dim model_relitem As New Njc.Model.M_setubi_Rev7_Model(itemcnt)
        '    For Each relitem In list_relitem

        '        Dim tmp_str() As String = relitem.Split("-")
        '        Dim setubikomkname As String = tmp_str(0)
        '        Dim setubinaiyo As String = relitem.Replace(setubikomkname & "-", "")

        '        model_relitem.Setubi_name(relcnt) = setubikomkname
        '        model_relitem.Setubi_lstname(relcnt) = setubinaiyo

        '        relcnt = relcnt + 1

        '    Next

        '    'モデルの引渡し
        '    RelItem_M_setubi_Rev7 = Nothing
        '    RelItem_M_setubi_Rev7 = model_relitem

        '    '************************
        '    '終了処理
        '    '************************

        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    Return rtn

        'End Function
        '20160525 全体的な動作の修正 -del end

        ''' <summary>
        ''' V7入金項目取得 '20160525 全体的な動作の修正
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadMid(ByVal sqlcnnV7 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim readtbl As New DataTable
            Dim reccnt As Integer
            Dim fldname As String
            Dim fldvalue As String
            Dim tmp_kinyuno As String = ""
            Dim tmp_kinyuname As String = ""
            Dim tmp_tenno As String = ""
            Dim tmp_tenname As String = ""
            Dim rtn As Boolean = True

            '画像タイトル情報取得
            Dim tmp_sql As String = Me.Get_UseQry()
            reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnV7, readtbl)

            'モデル初期化
            Dim model_relitem As New Njc.Model.M_setubi_Rev7_Model(reccnt)

            '読込開始
            With model_relitem

                For cntii As Integer = 1 To readtbl.Rows.Count

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Then
                        Return rtn
                    End If

                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                        '項目名取得
                        fldname = readtbl.Columns(cntjj).ColumnName

                        '登録値取得
                        fldvalue = (readtbl.Rows(cntii - 1).Item(cntjj)).ToString

                        '各項目値→変数格納
                        Select Case fldname
                            Case "設備項目名"
                                .Setubi_name(cntii) = fldvalue
                            Case "設備内容"
                                .Setubi_lstname(cntii) = fldvalue
                        End Select

                    Next

                Next

            End With

            'モデルの引渡し
            RelItem_M_setubi_Rev7 = Nothing
            RelItem_M_setubi_Rev7 = model_relitem

            '************************
            '終了処理
            '************************

            Return rtn

        End Function

        Public Function Get_UseQry() As String

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " SELECT DISTINCT [設備項目名],[設備内容] FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT "
            tmp_sql = tmp_sql & " 		 bk_no AS [物件NO] "
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
            tmp_sql = tmp_sql & " ORDER BY [設備項目名],[設備内容] "

            Return tmp_sql

        End Function

    End Class


#End Region


End Namespace