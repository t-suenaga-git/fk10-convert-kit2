Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "口座種別マスタ取得"

    Public Class M_kozasyu_Rev7_Repository

        ''' <summary>
        ''' 入金区分取得
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReadMid(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection) As Boolean

            '20161021 既存中間ファイルを無くすことによる速度改善処理_全変更 -chg sta
            '            Dim rtn As Boolean = True
            '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            '            Dim list_relitem As New List(Of String)
            '            Dim relfldno As Integer = 0

            '            Dim filename As String() = New String() {"業者情報", "業者情報", "業者情報", "自社情報", "契約者情報", "家主情報"}
            '            Dim sheetname As String() = New String() {"仲介業者口座情報", "保険業者口座情報", "修繕業者口座情報", "自社口座情報", "契約者口座情報", "家主口座情報"}
            '            Dim relfldname As String() = New String() {"口座種別", "口座種別", "口座種別", "口座種別", "口座種別", "口座種別"}

            '            'Dim hanyojisyaflg As Boolean = False    '20161004 自社口座の口座種別取得処理の修正 -add

            '            For cntfile = 0 To UBound(filename)

            '                '汎用用中間ファイル読込フラグ
            '                Dim hanyojisyaflg As Boolean = False

            '                If filename(cntfile) = "自社情報" Then

            '                    '自社口座情報は汎用用中間ファイルから読み込む
            '                    Dim basemidfilename As String = CV_FROM_MIDDLE
            '                    Call Me.ReadBaseMid(basemidfilename, sheetname(cntfile), relfldname(cntfile), list_relitem)
            '                    hanyojisyaflg = True

            '                End If

            '                '************************
            '                '作業準備
            '                '************************

            '                If hanyojisyaflg = False Then

            '                    'Excelファイル初期設定                  
            '                    Dim tmp_sql As String = "SELECT DISTINCT " & "[" & relfldname(cntfile) & "]" & " FROM [" & sheetname(cntfile) & "$] WHERE " & "[" & relfldname(cntfile) & "]" & "<> '' "
            '                    Dim readtbl As New DataTable()
            '                    Dim con_read As New OleDbConnection()

            '                    'オープン処理
            '                    rtn = excelfile.ExcelFile_ReadOpen(MidDirPath, filename(cntfile), tmp_sql, readtbl, con_read)

            '                    'オープン処理失敗時は処理を抜ける
            '                    If rtn = False Then
            '                        excelfile.ExcelFile_ReadClose(con_read)
            '                        Return rtn
            '                    End If

            '                    '行数取得
            '                    Dim rowcnt As Integer = readtbl.Rows.Count

            '                    'データが存在しない場合は処理を抜ける
            '                    If rowcnt = 0 Then
            '                        excelfile.ExcelFile_ReadClose(con_read)
            '                        GoTo skiplbl
            '                    End If

            '                    '************************
            '                    '処理開始
            '                    '************************

            '                    For cntii As Integer = 0 To rowcnt - 1

            '                        '中断処理
            '                        Application.DoEvents()
            '                        If CancelFlg Then
            '                            Call excelfile.ExcelFile_ReadClose(con_read)
            '                            Return rtn
            '                        End If

            '                        '登録値取得→オブジェクトへ格納
            '                        Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(0)).Trim
            '                        If list_relitem.Contains(fldvalue) = False Then
            '                            list_relitem.Add(fldvalue)
            '                        End If

            '                    Next

            '                    '************************
            '                    '終了処理
            '                    '************************
            'skiplbl:
            '                    'クローズ処理
            '                    excelfile.ExcelFile_ReadClose(con_read)

            '                End If

            '            Next

            '            '集約した紐付データをモデルへ格納
            '            Dim relcnt As Integer = 1
            '            Dim itemcnt As Integer = list_relitem.Count
            '            Dim model_relitem As New Njc.Model.M_kozasyu_Rev7_Model(itemcnt)
            '            For Each relitem In list_relitem
            '                model_relitem.Kosyu_name(relcnt) = relitem
            '                relcnt = relcnt + 1
            '            Next

            '            'モデルの引渡し
            '            RelItem_M_kozasyu_Rev7 = Nothing
            '            RelItem_M_kozasyu_Rev7 = model_relitem

            '            Return rtn
            
            Dim rtn As Boolean = True
            Dim list_relitem As New List(Of String)

            '************************
            '作業準備
            '************************
            'データ取得
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT DISTINCT [口座種別] FROM CVTBL_仲介業者口座情報 WHERE RTRIM(LTRIM([口座種別])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [口座種別] FROM CVTBL_保険業者口座情報 WHERE RTRIM(LTRIM([口座種別])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [口座種別] FROM CVTBL_修繕業者口座情報 WHERE RTRIM(LTRIM([口座種別])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [口座種別] FROM CVTBL_自社口座情報 WHERE RTRIM(LTRIM([口座種別])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [口座種別] FROM CVTBL_契約者口座情報 WHERE RTRIM(LTRIM([口座種別])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [口座種別] FROM CVTBL_家主口座情報 WHERE RTRIM(LTRIM([口座種別])) <> '' "
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

            '************************
            '処理開始
            '************************

            For cntii As Integer = 0 To rowcnt - 1

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
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
            Dim model_relitem As New Njc.Model.M_kozasyu_Rev7_Model(itemcnt)
            For Each relitem In list_relitem
                model_relitem.Kosyu_name(relcnt) = relitem
                relcnt = relcnt + 1
            Next

            'モデルの引渡し
            RelItem_M_kozasyu_Rev7 = Nothing
            RelItem_M_kozasyu_Rev7 = model_relitem

            Return rtn
            '20161021 既存中間ファイルを無くすことによる速度改善処理_全変更 -chg end
        End Function

        '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del sta
        '自社口座も仮テーブルへ登録するため以下の処理を削除
        ' ''' <summary>
        ' ''' 入金区分取得
        ' ''' 自社口座の入金区分は汎用用中間ファイルから直接取得する
        ' ''' </summary>
        ' ''' <remarks></remarks>
        'Public Function ReadBaseMid(ByVal filename As String, ByVal sheetname As String, ByVal fldname As String, ByRef list_relitem As List(Of String)) As Boolean
        '    Dim rtn As Boolean = True
        '    Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
        '    Dim relfldno As Integer = 0

        '    'Excelファイル初期設定                  
        '    Dim tmp_sql As String = "SELECT * FROM [" & sheetname & "$] "
        '    Dim readtbl As New DataTable()
        '    Dim con_read As New OleDbConnection()

        '    'オープン処理
        '    rtn = excelfile.ExcelFile_ReadOpen(BaseMidDirPath, filename, tmp_sql, readtbl, con_read)

        '    'オープン処理失敗時は処理を抜ける
        '    If rtn = False Then
        '        excelfile.ExcelFile_ReadClose(con_read)
        '        Return rtn
        '    End If

        '    '行数取得
        '    Dim totalrowcnt As Integer = readtbl.Rows.Count     '全件数
        '    Dim rowcnt As Integer = readtbl.Rows.Count - 10     'データ部件数

        '    'データが存在しない場合は処理を抜ける
        '    If rowcnt = 0 Then
        '        excelfile.ExcelFile_ReadClose(con_read)
        '        Return rtn
        '    End If

        '    '************************
        '    '処理開始
        '    '************************

        '    '汎用中間ファイルからはヘッダーを取得できないので行Noから取得する

        '    Dim taisyocol As Integer = 0    '取得対象列No格納用

        '    For cntii = 0 To totalrowcnt - 1

        '        '中断処理
        '        Application.DoEvents()
        '        If CancelFlg Then
        '            Call excelfile.ExcelFile_ReadClose(con_read)
        '            Return rtn
        '        End If

        '        If cntii = 0 Then

        '            'ヘッダー取得→対象列No取得
        '            For cntjj = 0 To readtbl.Columns.Count - 1
        '                Dim tmp_fldname As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim
        '                If tmp_fldname = fldname Then
        '                    taisyocol = cntjj
        '                    Exit For
        '                End If
        '            Next

        '        ElseIf cntii >= BASEMIDFILE_READWRITE_ROW - 2 Then

        '            '登録値取得→オブジェクトへ格納
        '            Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(taisyocol)).Trim
        '            If list_relitem.Contains(fldvalue) = False Then
        '                list_relitem.Add(fldvalue)
        '            End If

        '        End If

        '    Next

        '    '************************
        '    '終了処理
        '    '************************

        '    'クローズ処理
        '    excelfile.ExcelFile_ReadClose(con_read)

        '    Return rtn

        'End Function
        '20161021 既存中間ファイルを無くすことによる速度改善処理_不要処理削除 -del end

    End Class


#End Region


End Namespace