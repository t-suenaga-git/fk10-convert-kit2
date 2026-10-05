Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports RelationSetting.Njc.Common

Namespace Njc.Repository

#Region "自社口座マスタ取得"

    Public Class M_jisyakoza_Rev7_Repository

        ''' <summary>
        ''' V7自社口座取得
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

            Dim filename As String = "自社情報"
            Dim sheetname As String = "自社口座情報"
            Dim relfldno As Integer = 0
            Dim relitem As String = "自社口座マスタ"

            

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

            '作業用変数
            '20160526 自社口座本体側の修正反映 -add sta
            Dim tmp_kozasyutokumoto(rowcnt) As String
            Dim tmp_kozasyutokumotono(rowcnt) As String   '20160829 自社口座にデフォルト値を設定する処理を追加 -add
            Dim tmp_kozaname(rowcnt) As String
            '20160526 自社口座本体側の修正反映 -add end
            Dim tmp_kinyuno(rowcnt) As String
            Dim tmp_kinyuname(rowcnt) As String
            Dim tmp_tenno(rowcnt) As String
            Dim tmp_tenname(rowcnt) As String
            Dim tmp_kosyu_name(rowcnt) As String
            Dim tmp_koza_no(rowcnt) As String
            Dim tmp_koza_meigi(rowcnt) As String
            '20160516 FB構築に伴う自社口座情報の修正 -add sta
            Dim tmp_koza_meigikana(rowcnt) As String
            Dim tmp_yucyokigo1(rowcnt) As String
            Dim tmp_yucyokigo2(rowcnt) As String
            Dim tmp_yucyokozano(rowcnt) As String
            Dim tmp_biko(rowcnt) As String
            '20160516 FB構築に伴う自社口座情報の修正 -add end
            Dim tmp_dupliflg(rowcnt) As Boolean   '20160526 自社口座本体側の修正反映 -chg 変数名を変更(tmp_entflg → tmp_dupliflg)
            Dim tmp_total As String = ""

            '取得した列のデータをリストへ格納 (重複集約)
            For cntii = 1 To rowcnt

                '行取得
                Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                '作業用変数初期化
                '20160526 自社口座本体側の修正反映 -add sta
                tmp_kozasyutokumoto(cntii) = ""
                tmp_kozasyutokumotono(cntii) = ""     '20160829 自社口座にデフォルト値を設定する処理を追加 -add
                tmp_kozaname(cntii) = ""
                '20160526 自社口座本体側の修正反映 -add end
                tmp_kinyuno(cntii) = ""
                tmp_kinyuname(cntii) = ""
                tmp_tenno(cntii) = ""
                tmp_tenname(cntii) = ""
                tmp_kosyu_name(cntii) = ""
                tmp_koza_no(cntii) = ""
                tmp_koza_meigi(cntii) = ""
                '20160516 FB構築に伴う自社口座情報の修正 -add sta
                tmp_koza_meigikana(cntii) = ""
                tmp_yucyokigo1(cntii) = ""
                tmp_yucyokigo2(cntii) = ""
                tmp_yucyokozano(cntii) = ""
                tmp_biko(cntii) = ""
                '20160516 FB構築に伴う自社口座情報の修正 -add end
                tmp_dupliflg(cntii) = True    '表示対象フラグ (True:表示 False:表示対象外)
                tmp_total = ""

                'ログ用変数作成
                Dim fldname_log_kinyu As String = ""  'ログ出力用
                Dim fldname_log_kinyuten As String = ""  'ログ出力用
                Dim fldname_log_kozano As String = ""  'ログ出力用
                Dim fldname_log_kozameigi As String = ""  'ログ出力用
                Dim log_taisyostr As String = ""

                For cntjj = 1 To columncnt

                    Dim fldname As String = headerrow(startrow - 1, cntjj).ToString.Trim
                    Dim fldvalue As String = ""
                    If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                        fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                    End If

                    Select Case fldname
                        '20160526 自社口座本体側の修正反映 -add sta
                        Case "口座取得元"
                            tmp_kozasyutokumoto(cntii) = fldvalue
                        Case "各口座設定No"     '20160829 自社口座にデフォルト値を設定する処理を追加 -add
                            tmp_kozasyutokumotono(cntii) = fldvalue
                        Case "各口座設定名称"
                            tmp_kozaname(cntii) = fldvalue
                            '20160526 自社口座本体側の修正反映 -add end
                        Case "金融機関No"
                            fldname_log_kinyu = fldname
                            tmp_kinyuno(cntii) = fldvalue
                        Case "金融機関店No"
                            fldname_log_kinyuten = fldname
                            tmp_tenno(cntii) = fldvalue
                        Case "口座種別"
                            tmp_kosyu_name(cntii) = fldvalue
                        Case "口座番号"
                            fldname_log_kozano = fldname
                            tmp_koza_no(cntii) = fldvalue
                        Case "口座名義"
                            fldname_log_kozameigi = fldname
                            tmp_koza_meigi(cntii) = fldvalue
                            '20160516 FB構築に伴う自社口座情報の修正 -add sta
                        Case "口座名義カナ"
                            tmp_koza_meigikana(cntii) = fldvalue
                        Case "ゆうちょ口座記号１"
                            tmp_yucyokigo1(cntii) = fldvalue
                        Case "ゆうちょ口座記号２"
                            tmp_yucyokigo2(cntii) = fldvalue
                        Case "ゆうちょ口座番号"
                            tmp_yucyokozano(cntii) = fldvalue
                        Case "備考(口座情報)"
                            tmp_biko(cntii) = fldvalue
                            '20160516 FB構築に伴う自社口座情報の修正 -add end
                    End Select

                Next

                'データチェック
                Dim errstr As String = ""

                Dim normalflg As Boolean = True '20160526 自社口座本体側の修正反映 -add

                '20160526 自社口座本体側の修正反映 -chg sta
                ''金融機関
                'tmp_entflg(cntii) = DataChk.Chk_DataMstExist_Kinyu(sqlcnnV10, tmp_kinyuno(cntii), tmp_kinyuno(cntii) & "-" & tmp_tenno(cntii), tmp_kinyuname(cntii), tmp_tenname(cntii), errstr)
                'If tmp_entflg(cntii) = False Then
                '    log_taisyostr = sheetname & " - " & _
                '                    fldname_log_kinyu & " = " & tmp_kinyuno(cntii) & "、" & _
                '                    fldname_log_kinyuten & " = " & tmp_tenno(cntii)
                '    GoTo loglabel
                'End If

                ''口座番号
                'Dim chkafterstr As String = ""
                'tmp_entflg(cntii) = DataChk.Chk_DataNumeric(tmp_koza_no(cntii), chkafterstr, "1", "9999999", "", errstr)
                'If tmp_entflg(cntii) = False Then
                '    log_taisyostr = sheetname & " - " & fldname_log_kozano & " = " & tmp_koza_no(cntii)
                '    GoTo loglabel
                'End If

                '金融機関
                normalflg = DataChk.Chk_DataMstExist_Kinyu(sqlcnnV10, tmp_kinyuno(cntii), tmp_kinyuno(cntii) & "-" & tmp_tenno(cntii), tmp_kinyuname(cntii), tmp_tenname(cntii), errstr)
                If normalflg = False Then
                    log_taisyostr = sheetname & " - " & _
                                    fldname_log_kinyu & " = " & tmp_kinyuno(cntii) & "、" & _
                                    fldname_log_kinyuten & " = " & tmp_tenno(cntii)
                    Dim log_sql As String = ""
                    Dim tmp_cnt As Integer = 0
                    If errstr <> "" Then
                        Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
                        DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
                    End If
                End If

                '口座番号
                Dim chkafterstr As String = ""
                normalflg = DataChk.Chk_DataNumeric(tmp_koza_no(cntii), chkafterstr, "1", "9999999", "", errstr)
                If normalflg = False Then
                    log_taisyostr = sheetname & " - " & fldname_log_kozano & " = " & tmp_koza_no(cntii)
                    Dim log_sql As String = ""
                    Dim tmp_cnt As Integer = 0
                    If errstr <> "" Then
                        Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
                        DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
                    End If
                End If
                '20160526 自社口座本体側の修正反映 -chg end

                '重複
                tmp_total = tmp_kinyuno(cntii) & "-" & tmp_tenno(cntii) & "-" & tmp_koza_no(cntii)

                '20160516 FB構築に伴う自社口座情報の修正 -chg sta
                'If (tmp_kinyuno(cntii) <> "" And tmp_tenno(cntii) <> "" And tmp_koza_no(cntii) <> "") And list_relitem.Contains(tmp_total) = False Then
                '    list_relitem.Add(tmp_total)     '重複チェック用に格納
                'ElseIf list_relitem.Contains(tmp_total) Then
                '    errstr = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & ""
                '    log_taisyostr = sheetname & " - " & _
                '                    fldname_log_kinyu & " = " & tmp_kinyuno(cntii) & "、" & _
                '                    fldname_log_kinyuten & " = " & tmp_tenno(cntii) & "、" & _
                '                    fldname_log_kozano & " = " & tmp_koza_no(cntii) & "、" & _
                '                    fldname_log_kozameigi & " = " & tmp_koza_meigi(cntii)
                '    tmp_entflg(cntii) = False
                'End If
                If tmp_total.Replace("-", "") <> "" And list_relitem.Contains(tmp_total) Then
                    errstr = LOG_NAIYO_ERR_OVERLAP & "-" & LOG_HUBI_OVERLAP & "-" & ""
                    log_taisyostr = sheetname & " - " & _
                                    fldname_log_kinyu & " = " & tmp_kinyuno(cntii) & "、" & _
                                    fldname_log_kinyuten & " = " & tmp_tenno(cntii) & "、" & _
                                    fldname_log_kozano & " = " & tmp_koza_no(cntii) & "、" & _
                                    fldname_log_kozameigi & " = " & tmp_koza_meigi(cntii)
                    tmp_dupliflg(cntii) = False
                    Dim log_sql As String = ""
                    Dim tmp_cnt As Integer = 0
                    If errstr <> "" Then
                        Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
                        DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
                    End If
                Else
                    list_relitem.Add(tmp_total)     '移行対象用に格納
                End If
                '20160516 FB構築に伴う自社口座情報の修正 -chg end

                '20160526 自社口座本体側の修正反映 -del sta
                'loglabel:
                'If tmp_entflg(cntii) = False Then
                '    'ログ出力
                '    'メッセージ整形
                '    Dim log_sql As String = ""
                '    Dim tmp_cnt As Integer = 0
                '    '20160516 FB構築に伴う自社口座情報の修正 -chg sta
                '    'Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
                '    'DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
                '    If errstr <> "" Then
                '        Call LogSetting.Set_Log_Value_KomkErr(relitem, log_taisyostr, errstr, log_sql)
                '        DBExec.Exec_NonQuery(sqlcnnV10, log_sql, tmp_cnt)
                '    End If
                '    '20160516 FB構築に伴う自社口座情報の修正 -chg end
                'End If
                '20160526 自社口座本体側の修正反映 -del end

            Next

            '集約した紐付データをモデルへ格納
            Dim relcnt As Integer = 1
            Dim itemcnt As Integer = list_relitem.Count
            Dim model_relitem As New Njc.Model.M_jisyakoza_Rev7_Model(itemcnt)
            Dim modelindex As Integer = 1

            For cntkk = 1 To rowcnt
                If tmp_dupliflg(cntkk) Then
                    With model_relitem
                        '20160526 自社口座本体側の修正反映 -add sta
                        .Kozasyutokumoto(modelindex) = tmp_kozasyutokumoto(cntkk)
                        .Kozasyutokumotono(modelindex) = tmp_kozasyutokumotono(cntkk)   '20160829 自社口座にデフォルト値を設定する処理を追加 -add
                        .Kozaname(modelindex) = tmp_kozaname(cntkk)
                        '20160526 自社口座本体側の修正反映 -add end
                        .Kinyu_no(modelindex) = tmp_kinyuno(cntkk)
                        .Kinyu_name(modelindex) = tmp_kinyuname(cntkk)
                        .Ten_no(modelindex) = tmp_tenno(cntkk)
                        .Ten_name(modelindex) = tmp_tenname(cntkk)
                        .Kosyu_name(modelindex) = tmp_kosyu_name(cntkk)
                        .Koza_no(modelindex) = tmp_koza_no(cntkk)
                        .Koza_meigi(modelindex) = tmp_koza_meigi(cntkk)
                        '20160516 FB構築に伴う自社口座情報の修正 -add sta
                        .Koza_kana(modelindex) = tmp_koza_meigikana(cntkk)
                        .Yucyokigo1(modelindex) = tmp_yucyokigo1(cntkk)
                        .Yucyokigo2(modelindex) = tmp_yucyokigo2(cntkk)
                        .Yucyokozano(modelindex) = tmp_yucyokozano(cntkk)
                        .biko(modelindex) = tmp_biko(cntkk)
                        '20160516 FB構築に伴う自社口座情報の修正 -add end
                        modelindex = modelindex + 1
                    End With
                End If
            Next

            'モデルの引渡し
            RelItem_M_jisyakoza_Rev7 = Nothing
            RelItem_M_jisyakoza_Rev7 = model_relitem

            '************************
            '終了処理
            '************************

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

        ''' <summary>
        ''' 金融機関マスタの取得→ハッシュテーブルへ格納
        ''' </summary>
        ''' <param name="sqlcnnV10"></param>
        ''' <param name="hash_kinyu"></param>
        ''' <param name="hash_kinyuten"></param>
        ''' <remarks></remarks>
        Public Sub Set_HashKinyu(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection, ByRef hash_kinyu As Hashtable, ByRef hash_kinyuten As Hashtable)

            Dim readtbl As New DataTable
            Dim reccnt As Integer
            Dim fldname As String
            Dim fldvalue As String
            Dim tmp_kinyuno As String = ""
            Dim tmp_kinyuname As String = ""
            Dim tmp_tenno As String = ""
            Dim tmp_tenname As String = ""

            '金融機関情報取得
            reccnt = DBExec.Exec_DataTable(Me.Get_UseQry(), sqlcnnV10, readtbl)

            '読込開始
            For cntii As Integer = 0 To readtbl.Rows.Count - 1

                For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                    '項目名取得
                    fldname = readtbl.Columns(cntjj).ColumnName

                    '登録値取得
                    fldvalue = readtbl.Rows(cntii).Item(cntjj).ToString.Trim

                    '各項目値→変数格納
                    Select Case fldname
                        Case "kinyu_no"
                            tmp_kinyuno = fldvalue
                        Case "kinyu_name"
                            tmp_kinyuname = fldvalue
                        Case "kinyu_tenno"
                            tmp_tenno = fldvalue
                        Case "kinyu_tenname"
                            tmp_tenname = fldvalue
                    End Select

                Next

                'ハッシュテーブルへ格納
                If Not hash_kinyu.Contains(tmp_kinyuno) Then
                    hash_kinyu.Add(tmp_kinyuno, tmp_kinyuname)
                End If
                hash_kinyuten.Add(tmp_kinyuno & "-" & tmp_tenno, tmp_tenname)

            Next

        End Sub

        ''' <summary>
        ''' 金融機関マスタ抽出クエリ
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Get_UseQry()

            Dim strsql As String = ""

            strsql = strsql & " SELECT "
            strsql = strsql & " 	 KINYU.kinyu_no "
            strsql = strsql & " 	,KINYU.kinyu_name "
            strsql = strsql & " 	,TEN.kinyu_tenno "
            strsql = strsql & " 	,TEN.kinyu_tenname "
            strsql = strsql & " FROM m_kinyu_ten AS TEN "
            strsql = strsql & " LEFT JOIN m_kinyu AS KINYU ON TEN.kinyu_no = KINYU.kinyu_no "
            strsql = strsql & " ORDER BY KINYU.kinyu_no,TEN.kinyu_tenno "

            Return strsql

        End Function

    End Class

#End Region

End Namespace