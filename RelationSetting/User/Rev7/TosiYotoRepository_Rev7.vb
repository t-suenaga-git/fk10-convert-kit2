Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "都市計画・用途地域マスタ取得"

    Public Class M_TosiYoto_Rev7_Repository

        ''' <summary>
        ''' V7都市計画・用途地域取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadMid() As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim rtn As Boolean = True
            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim list_relitem As New List(Of String)

            Dim filename As String = "物件情報"
            Dim sheetname As String = "物件詳細情報"
            Dim relfldno As Integer = 0

            '************************
            '作業準備
            '************************

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MidDirPath, filename, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160525 全体的な動作の修正 -add
                Return rtn
            End If

            'データが存在しない場合は処理を抜ける
            If rowcnt = 0 Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160525 全体的な動作の修正 -add
                Return rtn
            End If

            '************************
            '処理開始
            '************************

            Dim headerrow As New Object

            For cntii = 1 To rowcnt

                '行取得
                Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                'ヘッダー行取得
                headerrow = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                '作業用変数
                Dim tmp_tosi As String = ""
                Dim tmp_yoto As String = ""

                '対象フィールドを検索し、ヒットした際にその列番号と列の値を取得
                For cntjj = 1 To columncnt

                    Dim fldname As String = headerrow(startrow - 1, cntjj).ToString.Trim
                    Dim fldvalue As String = ""
                    If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                        fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                    End If

                    Select Case fldname
                        Case "都市計画"
                            tmp_tosi = fldvalue
                        Case "用途地域"
                            tmp_yoto = fldvalue
                    End Select

                Next

                '都市計画を格納
                If tmp_tosi <> "" And list_relitem.Contains(tmp_tosi) = False Then
                    list_relitem.Add(tmp_tosi)
                End If

                '用途地域を格納
                If tmp_yoto <> "" And list_relitem.Contains(tmp_yoto) = False Then
                    list_relitem.Add(tmp_yoto)
                End If

            Next

            '集約した紐付データをモデルへ格納
            Dim relcnt As Integer = 1
            Dim itemcnt As Integer = list_relitem.Count
            Dim model_relitem As New Njc.Model.M_TosiYoto_Rev7_Model(itemcnt)
            For Each relitem In list_relitem
                model_relitem.TosiYoto_name(relcnt) = relitem
                relcnt = relcnt + 1
            Next

            'モデルの引渡し
            RelItem_M_TosiYoto_Rev7 = Nothing
            RelItem_M_TosiYoto_Rev7 = model_relitem

            '************************
            '終了処理
            '************************

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

    End Class

#End Region


End Namespace