Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports RelationSetting.Njc.N3Lib.Utys
Imports RelationSetting.Njc.Common

Namespace Njc.Repository

#Region "物件分類マスタ取得"

    Public Class M_brui_Repository

        ''' <summary>
        ''' 物件分類取得 20161012 紐付ツール速度改善対応
        ''' 新規作成
        ''' 修正前のメソッドはコメントアウト
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReadMid(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim rtn As Boolean = True
            Dim list_relitem As New List(Of String)
            '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del sta
            'Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用

            'Dim filename As String = "物件情報"
            'Dim sheetname As String = "物件詳細情報"
            'Dim relfldname As String = "物件分類"
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
            Dim tmp_sql As String = " SELECT DISTINCT [物件分類] FROM CVTBL_物件詳細情報 WHERE RTRIM(LTRIM([物件分類])) <> '' "
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
            Dim model_relitem As New Njc.Model.M_brui_Model(itemcnt)
            For Each relitem In list_relitem
                model_relitem.Brui_name(relcnt) = relitem
                relcnt = relcnt + 1
            Next

            'モデルの引渡し
            RelItem_M_brui = Nothing
            RelItem_M_brui = model_relitem

            '************************
            '終了処理
            '************************
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del sta
            ''クローズ処理
            'excelfile.ExcelFile_ReadClose(con_read)
            '20161021 既存中間ファイルを無くすことによる速度改善処理_共通_クローズ処理 -del end
            Return rtn

        End Function

        '20161012 紐付ツール速度改善対応 -del sta
        '旧srcコメントアウト
        ' ''' <summary>
        ' ''' V7物件分類取得
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

        '    Dim filename As String = "物件情報"
        '    Dim sheetname As String = "物件詳細情報"
        '    Dim relfldname As String = "物件分類"
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

        '    '20160525 全体的な動作の修正 -add sta
        '    'データが存在しない場合は処理を抜ける
        '    If rowcnt = 0 Then
        '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160525 全体的な動作の修正 -add
        '        Return rtn
        '    End If
        '    '20160525 全体的な動作の修正 -add end

        '    '************************
        '    '処理開始
        '    '************************

        '    Dim headerrow As New Object
        '    Dim relcolvalue As New Object

        '    'ヘッダー行取得
        '    headerrow = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

        '    '対象フィールドを検索し、ヒットした際にその列番号と列の値を取得
        '    For cntcol = 1 To columncnt

        '        Dim fldname As String = headerrow(startrow - 1, cntcol).ToString.Trim

        '        If fldname = relfldname Then
        '            relfldno = cntcol
        '            relcolvalue = wsheet.Range(wsheet.Cells(startrow, cntcol), wsheet.Cells(maxrowcnt, cntcol)).Value
        '            Exit For
        '        End If

        '    Next

        '    '取得した列のデータをリストへ格納 (重複集約)
        '    '20160525 全体的な動作の修正 -chg sta
        '    'For cntii = 1 To rowcnt - 1

        '    '    Dim fldvalue As String = ""

        '    '    If relcolvalue(cntii, 1) IsNot Nothing Then
        '    '        fldvalue = relcolvalue(cntii, 1).ToString.Trim
        '    '    End If

        '    '    If fldvalue <> "" And list_relitem.Contains(fldvalue) = False Then
        '    '        list_relitem.Add(fldvalue)
        '    '    End If

        '    'Next
        '    If relcolvalue IsNot Nothing Then

        '        For cntii = 1 To maxrowcnt - 1

        '            Dim fldvalue As String = ""

        '            If relcolvalue(cntii, 1) IsNot Nothing Then
        '                fldvalue = relcolvalue(cntii, 1).ToString.Trim
        '            End If

        '            If fldvalue <> "" And list_relitem.Contains(fldvalue) = False Then
        '                list_relitem.Add(fldvalue)
        '            End If

        '        Next

        '    End If
        '    '20160525 全体的な動作の修正 -chg end

        '    '集約した紐付データをモデルへ格納
        '    Dim relcnt As Integer = 1
        '    Dim itemcnt As Integer = list_relitem.Count
        '    Dim model_relitem As New Njc.Model.M_brui_Model(itemcnt)
        '    For Each relitem In list_relitem
        '        model_relitem.Brui_name(relcnt) = relitem
        '        relcnt = relcnt + 1
        '    Next

        '    'モデルの引渡し
        '    RelItem_M_brui = Nothing
        '    RelItem_M_brui = model_relitem

        '    '************************
        '    '終了処理
        '    '************************

        '    'Excelファイル終了設定
        '    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    Return rtn

        'End Function
        '20161012 紐付ツール速度改善対応 -del end

    End Class

#End Region

End Namespace