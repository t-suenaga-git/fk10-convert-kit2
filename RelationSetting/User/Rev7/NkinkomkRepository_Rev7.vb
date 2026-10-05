Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "入金項目マスタ取得"

    Public Class M_nkin_Rev7_Repository

        ''' <summary>
        ''' V7入金項目取得 20160915 汎用の紐付時は入金区分および入金項目を中間ファイルから読み取る処理に修正 -add
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadMid_Base(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean
            '20161021 既存中間ファイルを無くすことによる速度改善処理_全変更 -chg sta
            '            Dim rtn As Boolean = True
            '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            '            Dim list_relitem As New List(Of String)

            '            Dim filename As String() = New String() {"部屋情報", "部屋情報", "送金ルール情報", "送金ルール情報", "契約情報", "契約情報", "契約情報", "契約情報", "請求情報"}
            '            Dim sheetname As String() = New String() {"部屋入金項目情報", "部屋入金項目情報", "送金ルール入金項目情報", "送金ルール控除項目情報", "契約入金項目情報", "契約入金項目情報", "契約次回入金項目情報", "契約次回入金項目情報", "家主固定控除情報"}
            '            Dim fldname_nkinkomkkbn As String() = New String() {"[入金項目区分]", "1 AS [入金項目区分]", "[月々/契約時/更新時区分]", "9 AS [入金項目区分]", "[入金項目区分]", "1 AS [入金項目区分]", "[入金項目区分]", "1 AS [入金項目区分]", "9 AS [入金項目区分]"}
            '            Dim fldname_nkinkomk As String() = New String() {"[入金項目名]", "算出基準入金項目名", "[入金項目名]", "[控除入金項目名]", "[入金項目名]", "[算出基準入金項目名]", "[入金項目名]", "[算出基準入金項目名]", "[入金項目名]"}
            '            Dim relitem As String = "入金項目マスタ"

            '            For cntfile = 0 To UBound(filename)

            '                '************************
            '                '作業準備
            '                '************************

            '                'Excelファイル初期設定                  
            '                Dim tmp_sql As String = "SELECT DISTINCT " & fldname_nkinkomkkbn(cntfile) & "," & fldname_nkinkomk(cntfile) & " FROM [" & sheetname(cntfile) & "$] WHERE " & fldname_nkinkomk(cntfile) & " <> '' ORDER BY 1"
            '                Dim readtbl As New DataTable()
            '                Dim con_read As New OleDbConnection()

            '                'オープン処理
            '                rtn = excelfile.ExcelFile_ReadOpen(MidDirPath, filename(cntfile), tmp_sql, readtbl, con_read)

            '                'オープン処理失敗時は処理を抜ける
            '                If rtn = False Then
            '                    excelfile.ExcelFile_ReadClose(con_read)
            '                    Return rtn
            '                End If

            '                '行数取得
            '                Dim rowcnt As Integer = readtbl.Rows.Count

            '                'データが存在しない場合は処理を抜ける
            '                If rowcnt = 0 Then
            '                    excelfile.ExcelFile_ReadClose(con_read)
            '                    GoTo skiplbl
            '                End If

            '                '************************
            '                '処理開始
            '                '************************

            '                For cntii As Integer = 0 To rowcnt - 1

            '                    '作業用変数
            '                    Dim tmp_total As String = ""
            '                    Dim fldname_log As String = ""  'ログ出力用

            '                    '中断処理
            '                    Application.DoEvents()
            '                    If CancelFlg Then
            '                        Call excelfile.ExcelFile_ReadClose(con_read)
            '                        Return rtn
            '                    End If

            '                    '登録値取得→オブジェクトへ格納
            '                    Dim fldvalue_nkinkbn As String = Typ.ToStr(readtbl.Rows(cntii).Item(0)).Trim
            '                    Dim fldvalue_nkinkomk As String = Typ.ToStr(readtbl.Rows(cntii).Item(1)).Trim
            '                    fldname_log = fldname_nkinkomkkbn(cntfile)


            '                    'データチェック (数値:入金区分)
            '                    Dim chkafterstr As String = ""
            '                    Dim errstr As String = ""
            '                    Dim addflg As Boolean = DataChk.Chk_DataNumeric(fldvalue_nkinkbn, chkafterstr, "1", "9", "", errstr)

            '                    tmp_total = fldvalue_nkinkbn & "-" & fldvalue_nkinkomk

            '                    If (fldvalue_nkinkbn <> "" And fldvalue_nkinkomk <> "") And list_relitem.Contains(tmp_total) = False And addflg Then
            '                        list_relitem.Add(tmp_total)
            '                    ElseIf addflg = False Then
            '                        'ログ出力
            '                        'メッセージ整形
            '                        Dim log_sql As String = ""
            '                        Dim tmp_cnt As Integer = 0
            '                        Dim log_taisyostr As String = sheetname(cntfile) & " - " & fldname_log & " = " & fldvalue_nkinkbn
            '                        Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
            '                        DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
            '                    End If

            '                Next
            'skiplbl:
            '                '************************
            '                '終了処理
            '                '************************
            '                'クローズ処理
            '                excelfile.ExcelFile_ReadClose(con_read)

            '            Next

            '            '集約した紐付データをモデルへ格納
            '            Dim relcnt As Integer = 1
            '            Dim itemcnt As Integer = list_relitem.Count
            '            Dim model_relitem As New Njc.Model.M_nkin_Rev7_Model(itemcnt)
            '            For Each relitem In list_relitem

            '                Dim tmp_str() As String = relitem.Split("-")
            '                Dim tukikbn As String = tmp_str(0)
            '                Dim nkinkomkname As String = relitem.Replace(tukikbn & "-", "")

            '                model_relitem.Nkin_kbn(relcnt) = Get_Nkinkomkkbn(tukikbn)
            '                model_relitem.Nkin_name(relcnt) = nkinkomkname

            '                relcnt = relcnt + 1

            '            Next

            '            'モデルの引渡し
            '            RelItem_M_nkin_Rev7 = Nothing
            '            RelItem_M_nkin_Rev7 = model_relitem

            '            Return rtn

            Dim rtn As Boolean = True
            Dim list_relitem As New List(Of String)
            Dim relitem As String = "入金項目マスタ"

            '************************
            '作業準備
            '************************
            'データ取得
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT DISTINCT [入金項目区分],[入金項目名] FROM CVTBL_部屋入金項目情報 WHERE RTRIM(LTRIM([入金項目区分])) <> '' AND RTRIM(LTRIM([入金項目名])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT 1 AS [入金項目区分],[算出基準入金項目名] FROM CVTBL_部屋入金項目情報 WHERE RTRIM(LTRIM([算出基準入金項目名])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [月々/契約時/更新時区分],[入金項目名] FROM CVTBL_送金ルール入金項目情報 WHERE RTRIM(LTRIM([月々/契約時/更新時区分])) <> '' AND RTRIM(LTRIM([入金項目名])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT 9 AS [入金項目区分],[控除入金項目名] FROM CVTBL_送金ルール控除項目情報 WHERE RTRIM(LTRIM([控除入金項目名])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [入金項目区分],[入金項目名] FROM CVTBL_契約入金項目情報 WHERE RTRIM(LTRIM([入金項目区分])) <> '' AND RTRIM(LTRIM([入金項目名])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT 1 AS [入金項目区分],[算出基準入金項目名] FROM CVTBL_契約入金項目情報 WHERE RTRIM(LTRIM([算出基準入金項目名])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [入金項目区分],[入金項目名] FROM CVTBL_契約次回入金項目情報 WHERE RTRIM(LTRIM([入金項目区分])) <> '' AND RTRIM(LTRIM([入金項目名])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT 1 AS [入金項目区分],[算出基準入金項目名] FROM CVTBL_契約次回入金項目情報 WHERE RTRIM(LTRIM([算出基準入金項目名])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT 9 AS [入金項目区分],[入金項目名] FROM CVTBL_家主固定控除情報 WHERE RTRIM(LTRIM([入金項目名])) <> '' "
            Dim readtbl As New DataTable
            Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnV10, readtbl, rtn)

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

                '作業用変数
                Dim tmp_total As String = ""
                Dim fldname_log As String = ""  'ログ出力用

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Return rtn
                End If

                '登録値取得→オブジェクトへ格納
                Dim fldvalue_nkinkbn As String = Typ.ToStr(readtbl.Rows(cntii).Item(0)).Trim
                Dim fldvalue_nkinkomk As String = Typ.ToStr(readtbl.Rows(cntii).Item(1)).Trim
                fldname_log = "入金項目区分"

                'データチェック (数値:入金区分)
                Dim chkafterstr As String = ""
                Dim errstr As String = ""
                Dim addflg As Boolean = DataChk.Chk_DataNumeric(fldvalue_nkinkbn, chkafterstr, "1", "9", "", errstr)

                tmp_total = fldvalue_nkinkbn & "-" & fldvalue_nkinkomk

                '20161026 入金項目取得条件の修正 -chg sta
                'If (fldvalue_nkinkbn <> "" And fldvalue_nkinkomk <> "") And list_relitem.Contains(tmp_total) = False And addflg Then
                'データが空ではない条件と既に取得したデータかどうかの条件は抽出時に行っているため条件から外す
                If addflg Then
                    '20161026 入金項目取得条件の修正 -chg end
                    list_relitem.Add(tmp_total)
                ElseIf addflg = False Then
                    'ログ出力
                    'メッセージ整形
                    Dim log_sql As String = ""
                    Dim tmp_cnt As Integer = 0
                    Dim log_taisyostr As String = fldname_log & " = " & fldvalue_nkinkbn
                    Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
                    DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
                End If

            Next

            '集約した紐付データをモデルへ格納
            Dim relcnt As Integer = 1
            Dim itemcnt As Integer = list_relitem.Count
            Dim model_relitem As New Njc.Model.M_nkin_Rev7_Model(itemcnt)
            For Each relitem In list_relitem

                Dim tmp_str() As String = relitem.Split("-")
                Dim tukikbn As String = tmp_str(0)
                Dim nkinkomkname As String = relitem.Replace(tukikbn & "-", "")

                model_relitem.Nkin_kbn(relcnt) = Get_Nkinkomkkbn(tukikbn)
                model_relitem.Nkin_name(relcnt) = nkinkomkname

                relcnt = relcnt + 1

            Next

            'モデルの引渡し
            RelItem_M_nkin_Rev7 = Nothing
            RelItem_M_nkin_Rev7 = model_relitem

            Return rtn
            '20161021 既存中間ファイルを無くすことによる速度改善処理_全変更 -chg end
        End Function

        ' ''' <summary>
        ' ''' V7入金項目取得 20160915 汎用の紐付時は入金区分および入金項目を中間ファイルから読み取る処理に修正 -add
        ' ''' </summary>
        ' ''' <remarks></remarks>
        'Public Function ReadMid_Base(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

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

        '    Dim filename As String() = New String() {"部屋情報", "部屋情報", "契約情報", "契約情報", "契約情報", "送金ルール情報", "送金ルール情報", "請求情報", "請求情報", "請求情報"}
        '    Dim sheetname As String() = New String() {"部屋入金項目情報", "部屋変動費各戸メーター情報", "契約入金項目情報", "契約次回入金項目情報", "契約変動費各戸メーター情報", "送金ルール入金項目情報", "送金ルール控除項目情報", "預り金情報", "未収滞納金情報", "家主固定控除情報"}
        '    Dim fldname_nkinkomk As String() = New String() {"入金項目名", "変動費入金項目名", "入金項目名", "入金項目名", "入金項目名", "入金項目名", "控除入金項目名", "入金項目名", "入金項目名", "入金項目名"}
        '    Dim fldname_nkinkomkkbn As String() = New String() {"入金項目区分", "", "入金項目区分", "入金項目区分", "", "月々/契約時/更新時区分", "契約時/更新時区分", "入金項目区分", "入金項目区分", ""}
        '    Dim relitem As String = "入金項目マスタ"

        '    For cntfile = 0 To UBound(filename)

        '        '************************
        '        '作業準備
        '        '************************

        '        'Excelファイル初期設定
        '        rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MidDirPath, filename(cntfile), sheetname(cntfile))

        '        'Excelファイル設定時にエラーが生じた際は処理を抜ける
        '        If Not rtn Then
        '            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160525 全体的な動作の修正 -add
        '            Return rtn
        '        End If

        '        '************************
        '        '処理開始
        '        '************************

        '        Dim headerrow As New Object
        '        Dim relcolvalue As New Object

        '        'ヘッダー行取得
        '        headerrow = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

        '        '取得した列のデータをリストへ格納 (重複集約)
        '        For cntii = 1 To rowcnt - 1

        '            '行取得
        '            Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

        '            '作業用変数
        '            Dim tmp_nkinkomkname As String = ""
        '            Dim tmp_nkinkbn As String = ""
        '            Dim tmp_total As String = ""
        '            Dim fldname_log As String = ""  'ログ出力用

        '            For cntjj = 1 To columncnt

        '                Dim fldname As String = headerrow(startrow - 1, cntjj).ToString.Trim

        '                Dim fldvalue As String = ""
        '                If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
        '                    fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
        '                End If

        '                Select Case fldname
        '                    Case fldname_nkinkomk(cntfile)
        '                        tmp_nkinkomkname = fldvalue
        '                    Case fldname_nkinkomkkbn(cntfile)
        '                        fldname_log = fldname
        '                        tmp_nkinkbn = fldvalue
        '                End Select

        '            Next

        '            '入金項目区分を設定
        '            Select Case sheetname(cntfile)
        '                Case "部屋変動費各戸メーター情報", "契約変動費各戸メーター情報"
        '                    tmp_nkinkbn = "5"
        '                Case "家主固定控除情報"
        '                    tmp_nkinkbn = "9"
        '            End Select

        '            'データチェック (数値:入金区分)
        '            Dim chkafterstr As String = ""
        '            Dim errstr As String = ""
        '            Dim addflg As Boolean = DataChk.Chk_DataNumeric(tmp_nkinkbn, chkafterstr, "1", "9", "", errstr)

        '            tmp_total = tmp_nkinkbn & "-" & tmp_nkinkomkname

        '            If (tmp_nkinkbn <> "" And tmp_nkinkomkname <> "") And list_relitem.Contains(tmp_total) = False And addflg Then
        '                list_relitem.Add(tmp_total)
        '            ElseIf addflg = False Then
        '                'ログ出力
        '                'メッセージ整形
        '                Dim log_sql As String = ""
        '                Dim tmp_cnt As Integer = 0
        '                Dim log_taisyostr As String = sheetname(cntfile) & " - " & fldname_log & " = " & tmp_nkinkbn
        '                Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
        '                DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
        '            End If

        '        Next

        '        '************************
        '        '終了処理
        '        '************************

        '        'Excelファイル終了設定
        '        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

        '    Next

        '    '集約した紐付データをモデルへ格納
        '    Dim relcnt As Integer = 1
        '    Dim itemcnt As Integer = list_relitem.Count
        '    Dim model_relitem As New Njc.Model.M_nkin_Rev7_Model(itemcnt)
        '    For Each relitem In list_relitem

        '        Dim tmp_str() As String = relitem.Split("-")
        '        Dim tukikbn As String = tmp_str(0)
        '        Dim nkinkomkname As String = relitem.Replace(tukikbn & "-", "")

        '        model_relitem.Nkin_kbn(relcnt) = Get_Nkinkomkkbn(tukikbn)
        '        model_relitem.Nkin_name(relcnt) = nkinkomkname

        '        relcnt = relcnt + 1

        '    Next

        '    'モデルの引渡し
        '    RelItem_M_nkin_Rev7 = Nothing
        '    RelItem_M_nkin_Rev7 = model_relitem

        '    Return rtn

        'End Function

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
            Dim model_relitem As New Njc.Model.M_nkin_Rev7_Model(reccnt)

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
                            'Case "nkin_no"
                            '    .Nkin_no(cntii) = fldvalue
                            Case "nkin_name"
                                .Nkin_name(cntii) = fldvalue
                            Case "nkin_kbn"
                                .Nkin_kbn(cntii) = Get_Nkinkomkkbn(fldvalue)
                        End Select

                    Next

                Next

            End With

            'モデルの引渡し
            RelItem_M_nkin_Rev7 = Nothing
            RelItem_M_nkin_Rev7 = model_relitem

            '************************
            '終了処理
            '************************

            Return rtn

        End Function

        Public Function Get_Nkinkomkkbn(ByVal nkinkomkkbn As String) As String

            Dim rtn As String = ""
            Dim tmp_str As String = ""

            Select Case nkinkomkkbn
                Case "1"
                    tmp_str = "通常月"
                Case "2"
                    tmp_str = "契約時"
                Case "3"
                    tmp_str = "更新時"
                Case "4"
                    tmp_str = "解約時"
                Case "5"
                    tmp_str = "随時変動"
                Case "6"
                    tmp_str = "その他"
                Case "7"
                    tmp_str = "修繕"
                Case "8"
                    tmp_str = "家主送金"
                Case "9"
                    tmp_str = "家主控除"
            End Select

            rtn = nkinkomkkbn & "." & tmp_str

            Return rtn

        End Function

        Public Function Get_UseQry() As String

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " /*入金項目マスタ取得 sta*/ "
            tmp_sql = tmp_sql & " SELECT DISTINCT "
            tmp_sql = tmp_sql & " 	 /*nkin_no*/ "
            tmp_sql = tmp_sql & "      nkin_name "
            tmp_sql = tmp_sql & "     ,CASE "
            tmp_sql = tmp_sql & "         WHEN nkin_kbn = 7 OR nkin_kbn = 8 OR nkin_kbn = 9 then 9 "
            tmp_sql = tmp_sql & "         ELSE nkin_kbn  "
            tmp_sql = tmp_sql & "         END AS nkin_kbn "
            tmp_sql = tmp_sql & " FROM m_nkin "
            tmp_sql = tmp_sql & " /*入金項目マスタ取得 end*/ "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " /*部屋入金項目情報から変動費で固定になっている入金項目を取得 sta*/ "
            tmp_sql = tmp_sql & " SELECT DISTINCT "
            tmp_sql = tmp_sql & " 	 /*HYNKIN.nkin_no*/ "
            tmp_sql = tmp_sql & " 	nkin_name, "
            tmp_sql = tmp_sql & " 	1 AS nkin_kbn "
            tmp_sql = tmp_sql & " FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & "     SELECT * FROM hy_kanri AS hyk  "
            tmp_sql = tmp_sql & "     WHERE koteihendo_kbn = 1 and nkin_kbn = 5 "
            tmp_sql = tmp_sql & " ) AS HYNKIN "
            tmp_sql = tmp_sql & " LEFT JOIN m_nkin AS MN on HYNKIN.nkin_no = MN.nkin_no "
            tmp_sql = tmp_sql & " /*部屋入金項目情報から変動費で固定になっている入金項目を取得 end*/ "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " /*入金項目マスタ取得 end*/ "
            tmp_sql = tmp_sql & " SELECT DISTINCT "
            tmp_sql = tmp_sql & " 	 /*KYNKIN.nkin_no*/ "
            tmp_sql = tmp_sql & " 	 nkin_name "
            tmp_sql = tmp_sql & " 	,1 AS nkin_kbn FROM "
            tmp_sql = tmp_sql & " ( "
            tmp_sql = tmp_sql & "     SELECT * FROM ky_sqdata AS kyk  "
            tmp_sql = tmp_sql & "     WHERE koteihendo_kbn = 1 and nkin_kbn = 5 "
            tmp_sql = tmp_sql & " ) AS KYNKIN "
            tmp_sql = tmp_sql & " LEFT JOIN m_nkin AS MN on KYNKIN.nkin_no = MN.nkin_no "
            tmp_sql = tmp_sql & " /*入金項目マスタ取得 end*/ "
            tmp_sql = tmp_sql & " ORDER BY nkin_kbn "

            Return tmp_sql

        End Function

    End Class

#End Region

End Namespace