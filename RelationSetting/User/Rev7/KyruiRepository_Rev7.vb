Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "契約分類マスタ取得"

    Public Class M_keirui_Repository

        ''' <summary>
        ''' 契約分類取得
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReadMid(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim rtn As Boolean = True
            Dim list_relitem As New List(Of String)
            '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del sta
            'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'Dim filename As String = "契約情報"
            'Dim sheetname As String = "契約履歴情報"
            'Dim relfldname As String = "契約分類名"
            'Dim relfldno As Integer = 0
            '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del end
            '************************
            '作業準備
            '************************
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg sta
            ''Excelファイル初期設定                  
            'Dim tmp_sql As String = "SELECT DISTINCT " & "[" & relfldname & "]" & " FROM [" & sheetname & "$] WHERE " & "[" & relfldname & "]" & "<> '' "
            'Dim readtbl As New DataTable()
            'Dim con_read As New OleDbConnection()

            ''オープン処理
            'rtn = excelfile.ExcelFile_ReadOpen(MidDirPath, filename, tmp_sql, readtbl, con_read)

            ''オープン処理失敗時は処理を抜ける
            'If rtn = False Then
            '    excelfile.ExcelFile_ReadClose(con_read)
            '    Return rtn
            'End If

            ''行数取得
            'Dim rowcnt As Integer = readtbl.Rows.Count

            ''データが存在しない場合は処理を抜ける
            'If rowcnt = 0 Then
            '    excelfile.ExcelFile_ReadClose(con_read)
            '    Return rtn
            'End If
            'データ取得
            Dim tmp_sql As String = " SELECT DISTINCT [契約分類名] FROM CVTBL_契約履歴情報 WHERE RTRIM(LTRIM([契約分類名])) <> '' "
            Dim readtbl As New DataTable
            Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

            'オープン処理失敗時は処理を抜ける
            If rtn = False Then
                Return rtn
            End If

            'データが存在しない場合は処理を抜ける
            If rowcnt <= 0 Then
                Return rtn
            End If
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_データ取得処理 -chg end
            '************************
            '処理開始
            '************************

            For cntii As Integer = 0 To rowcnt - 1

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del
                    'Call excelfile.ExcelFile_ReadClose(con_read)
                    Return rtn
                End If

                '登録値取得→オブジェクトへ格納
                Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(0)).Trim
                If list_relitem.Contains(fldvalue) = False Then
                    list_relitem.Add(fldvalue)
                End If

            Next

            '集約した紐付データをモデルへ格納
            Dim relcnt As Integer = 1
            Dim itemcnt As Integer = list_relitem.Count
            Dim model_relitem As New Njc.Model.M_keirui_Model(itemcnt)
            For Each relitem In list_relitem
                model_relitem.Keirui_name(relcnt) = relitem
                relcnt = relcnt + 1
            Next

            'モデルの引渡し
            RelItem_M_Keirui = Nothing
            RelItem_M_Keirui = model_relitem

            '************************
            '終了処理
            '************************
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
            ''クローズ処理
            'excelfile.ExcelFile_ReadClose(con_read)
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
            Return rtn

        End Function

    End Class

#End Region

End Namespace