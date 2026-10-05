Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports RelationSetting.Njc.Common

Namespace Njc.Repository

#Region "画像割付取得"

    Public Class M_gazo_title_Rev7_Repository

        ' ''' <summary>
        ' ''' V7画像タイトル取得
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

        '    Dim filename As String = "各マスタ情報"
        '    Dim sheetname As String = "画像タイトルマスタ"
        '    Dim relfldno As Integer = 0

        '    '************************
        '    '作業準備
        '    '************************

        '    'Excelファイル初期設定
        '    rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MidDirPath, filename, sheetname)

        '    'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '    If Not rtn Then
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
        '    For cntii = 1 To rowcnt

        '        '行取得
        '        Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

        '        '作業用変数
        '        Dim tmp_gazosyubetu As String = ""
        '        Dim tmp_gazono As String = ""
        '        Dim tmp_gazotitle As String = ""
        '        Dim tmp_total As String = ""

        '        For cntjj = 1 To columncnt

        '            Dim fldname As String = headerrow(startrow - 1, cntjj).ToString.Trim
        '            Dim fldvalue As String = ""
        '            If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
        '                fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
        '            End If

        '            Select Case fldname
        '                Case "画像種別"
        '                    tmp_gazosyubetu = fldvalue
        '                Case "画像No"
        '                    tmp_gazono = fldvalue
        '                Case "タイトル"
        '                    tmp_gazotitle = fldvalue
        '            End Select

        '        Next

        '        tmp_total = tmp_gazosyubetu & "-" & tmp_gazono & "-" & tmp_gazotitle

        '        If (tmp_gazosyubetu <> "" And tmp_gazono <> "" And tmp_gazotitle <> "") And list_relitem.Contains(tmp_total) = False Then
        '            list_relitem.Add(tmp_total)
        '        End If

        '    Next

        '    '集約した紐付データをモデルへ格納
        '    Dim relcnt As Integer = 1
        '    Dim itemcnt As Integer = list_relitem.Count
        '    Dim model_relitem As New Njc.Model.M_gazo_title_Rev7_Model(itemcnt)
        '    For Each relitem In list_relitem

        '        Dim tmp_str() As String = relitem.Split("-")
        '        Dim gazosyubetu As String = tmp_str(0)
        '        Dim gazono As String = tmp_str(1)
        '        Dim gazotitle As String = relitem.Replace(gazosyubetu & "-" & gazono & "-", "")

        '        model_relitem.Syubetu(relcnt) = gazosyubetu
        '        model_relitem.Kbn(relcnt) = gazono
        '        model_relitem.Gazo_name(relcnt) = gazotitle

        '        relcnt = relcnt + 1

        '    Next

        '    'モデルの引渡し
        '    RelItem_M_gazo_title_Rev7 = Nothing
        '    RelItem_M_gazo_title_Rev7 = model_relitem

        '    '************************
        '    '終了処理
        '    '************************

        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    Return rtn

        'End Function

        ''' <summary>
        ''' V7画像タイトル取得
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
            Dim model_relitem As New Njc.Model.M_gazo_title_Rev7_Model(reccnt)

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
                            Case "画像種別"
                                .Syubetu(cntii) = fldvalue
                            Case "画像No"
                                .Kbn(cntii) = fldvalue
                            Case "タイトル"
                                .Gazo_name(cntii) = fldvalue
                        End Select

                    Next

                Next

            End With

            'モデルの引渡し
            RelItem_M_gazo_title_Rev7 = Nothing
            RelItem_M_gazo_title_Rev7 = model_relitem

            '************************
            '終了処理
            '************************

            Return rtn

        End Function

        Public Function Get_UseQry() As String

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " SELECT * FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & " 	SELECT DISTINCT "
            tmp_sql = tmp_sql & " 		 1 AS [ソートNo] "
            tmp_sql = tmp_sql & " 		,'物件' AS [画像種別] "
            tmp_sql = tmp_sql & " 		,kbn AS [画像No] "
            tmp_sql = tmp_sql & " 		,gazo_name AS [タイトル] "
            tmp_sql = tmp_sql & " 	FROM bk_gazo AS BKG "
            tmp_sql = tmp_sql & " 	LEFT JOIN (SELECT * FROM m_gazo_syoki WHERE gazo_kbn = 1) AS MGS "
            tmp_sql = tmp_sql & " 	ON BKG.kbn = MGS.gazo_no "
            tmp_sql = tmp_sql & "   /*20160610 部屋画像の割付の削除 del sta*/ "
            tmp_sql = tmp_sql & "   /* "
            tmp_sql = tmp_sql & "   UNION "
            tmp_sql = tmp_sql & "   SELECT "
            tmp_sql = tmp_sql & "   	 2 AS [ソートNo]	 "
            tmp_sql = tmp_sql & "   	,'部屋' AS [画像種別] "
            tmp_sql = tmp_sql & "   	,kbn AS [画像No] "
            tmp_sql = tmp_sql & "   	,gazo_name AS [タイトル] "
            tmp_sql = tmp_sql & "   FROM hy_gazo AS HYG "
            tmp_sql = tmp_sql & "   LEFT JOIN (SELECT * FROM m_gazo_syoki WHERE gazo_kbn = 2) AS MGS "
            tmp_sql = tmp_sql & "   ON HYG.kbn = MGS.gazo_no "
            tmp_sql = tmp_sql & "   */ "
            tmp_sql = tmp_sql & "   /*20160610 部屋画像の割付の削除 del end*/ "
            tmp_sql = tmp_sql & " ) AS VW "
            tmp_sql = tmp_sql & " ORDER BY [ソートNo],[画像No] "

            Return tmp_sql

        End Function

    End Class

#End Region

End Namespace