Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "鍵タイトルマスタ取得"

    Public Class M_kagi_Rev7_Repository

        ''' <summary>
        ''' V7鍵情報取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadMid(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

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

            Dim filename As String = "各マスタ情報"
            Dim sheetname As String = "鍵タイトルマスタ"
            Dim relfldno As Integer = 0
            Dim relitem As String = "鍵タイトルマスタ"

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

            '20160525 全体的な動作の修正 -add sta
            'データが存在しない場合は処理を抜ける
            If rowcnt = 0 Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160525 全体的な動作の修正 -add
                Return rtn
            End If
            '20160525 全体的な動作の修正 -add end

            '************************
            '処理開始
            '************************

            Dim headerrow As New Object
            Dim relcolvalue As New Object

            'ヘッダー行取得
            headerrow = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            '取得した列のデータをリストへ格納 (重複集約)
            For cntii = 1 To rowcnt

                '行取得
                Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                '作業用変数
                Dim tmp_kagino As String = ""
                Dim tmp_kaginame As String = ""
                Dim tmp_total As String = ""
                Dim fldname_log As String = ""  'ログ出力用

                For cntjj = 1 To columncnt

                    Dim fldname As String = headerrow(startrow - 1, cntjj).ToString.Trim
                    Dim fldvalue As String = ""
                    If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                        fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                    End If

                    Select Case fldname
                        Case "鍵No"
                            tmp_kagino = fldvalue
                        Case "鍵名称"
                            tmp_kaginame = fldvalue
                    End Select

                Next

                'データチェック (数値:鍵No)
                Dim chkafterstr As String = ""
                Dim errstr As String = ""
                Dim addflg As Boolean = DataChk.Chk_DataNumeric(tmp_kagino, chkafterstr, "1", "50", "", errstr)

                tmp_total = tmp_kagino & "-" & tmp_kaginame

                If (tmp_kagino <> "" And tmp_kaginame <> "") And list_relitem.Contains(tmp_total) = False And addflg Then
                    list_relitem.Add(tmp_total)
                ElseIf addflg = False Then
                    'ログ出力
                    'メッセージ整形
                    Dim log_sql As String = ""
                    Dim tmp_cnt As Integer = 0
                    Dim log_taisyostr As String = sheetname & " - " & fldname_log & " = " & tmp_kagino
                    Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
                    DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
                End If

            Next

            '集約した紐付データをモデルへ格納
            Dim relcnt As Integer = 1
            Dim itemcnt As Integer = list_relitem.Count
            Dim model_relitem As New Njc.Model.M_kagi_Rev7_Model(itemcnt)
            For Each relitem In list_relitem

                Dim tmp_str() As String = relitem.Split("-")
                Dim kagino As String = tmp_str(0)
                Dim kaginame As String = relitem.Replace(kagino & "-", "")

                model_relitem.Kagi_no(relcnt) = kagino
                model_relitem.Kagi_name(relcnt) = kaginame

                relcnt = relcnt + 1

            Next

            'モデルの引渡し
            RelItem_M_kagi_Rev7 = Nothing
            RelItem_M_kagi_Rev7 = model_relitem

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