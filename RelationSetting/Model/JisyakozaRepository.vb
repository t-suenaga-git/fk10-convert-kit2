Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "自社口座マスタ取得"

    Public Class M_jisyakoza_Repository

        ''' <summary>
        ''' 10自社口座取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadRelItem() As Boolean

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
            Dim sheetname As String = "自社基本情報"
            Dim relfldno As Integer = 0

            '************************
            '作業準備
            '************************

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MidDirPath, filename, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Return rtn
            End If

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

                '作業用変数作成
                Dim tmp_jisyano As String = ""
                Dim tmp_jisyaname As String = ""
                Dim tmp_total As String = ""

                For cntjj = 1 To columncnt

                    Dim fldname As String = headerrow(startrow - 1, cntjj).ToString.Trim
                    Dim fldvalue As String = ""
                    If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                        fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                    End If

                    Select Case fldname
                        Case "自社・支店No"
                            tmp_jisyano = fldvalue
                        Case "自社・支店名"
                            tmp_jisyaname = fldvalue
                    End Select

                Next

                tmp_total = tmp_jisyano & "-" & tmp_jisyaname

                If (tmp_jisyano <> "" And tmp_jisyaname <> "") And list_relitem.Contains(tmp_total) = False Then
                    list_relitem.Add(tmp_total)     '重複チェック用に格納
                End If

            Next

            '集約した紐付データをモデルへ格納
            Dim relcnt As Integer = 1
            Dim itemcnt As Integer = list_relitem.Count
            Dim model_relitem As New Njc.Model.M_jisyakoza_Model(itemcnt)
            For Each relitem In list_relitem

                Dim tmp_str() As String = relitem.Split("-")
                Dim jisyano As String = tmp_str(0)
                Dim jisyaname As String = relitem.Replace(jisyano & "-", "")

                model_relitem.Jisya_no(relcnt) = jisyano
                '20160829 米津TL指摘事項No16の修正 -chg sta
                'model_relitem.Jisya_name(relcnt) = jisyaname
                model_relitem.Jisya_name(relcnt) = jisyano & "." & jisyaname
                '20160829 米津TL指摘事項No16の修正 -chg end

                relcnt = relcnt + 1

            Next

            'モデルの引渡し
            If rowcnt <> 0 Then
                RelItem_M_jisyakoza = Nothing
                RelItem_M_jisyakoza = model_relitem
            End If

            '************************
            '終了処理
            '************************

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

        ''' <summary>
        ''' 自社口座グリッドの初期設定
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Initialize_Dgv(ByVal dgv As DataGridView)

            With dgv
                '20160526 自社口座本体側の修正反映 -chg sta
                '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                '.Columns(3).Frozen = True
                '.Columns(0).ReadOnly = True
                '.Columns(0).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(0).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(1).ReadOnly = True
                '.Columns(1).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(1).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(2).ReadOnly = True
                '.Columns(2).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(2).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(3).ReadOnly = True
                '.Columns(3).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(3).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(4).ReadOnly = True
                '.Columns(4).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(4).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(5).ReadOnly = True
                '.Columns(5).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(5).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(6).ReadOnly = True
                '.Columns(6).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(6).HeaderCell.Style.BackColor = Color.PowderBlue
                ''20160516 FB構築に伴う自社口座情報の修正 -add sta
                ''口座名義カナ、ゆうちょ記号、ゆうちょ口座番号、備考を追加
                ''※紐付けには不要と思われるため非表示にしておく(データは出力する)
                ''.Columns(7).Visible = False
                ''.Columns(8).Visible = False
                ''.Columns(9).Visible = False
                ''.Columns(10).Visible = False
                ''.Columns(11).Visible = False
                '.Columns(7).ReadOnly = True
                '.Columns(7).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(7).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(8).ReadOnly = True
                '.Columns(8).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(8).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(9).ReadOnly = True
                '.Columns(9).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(9).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(10).ReadOnly = True
                '.Columns(10).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(10).HeaderCell.Style.BackColor = Color.PowderBlue
                '.Columns(11).ReadOnly = True
                '.Columns(11).DefaultCellStyle.BackColor = Color.Azure
                '.Columns(11).HeaderCell.Style.BackColor = Color.PowderBlue
                ''20160516 FB構築に伴う自社口座情報の修正 -add end

                ''20160516 FB構築に伴う自社口座情報の修正 -chg sta
                ''.Columns(7).HeaderCell.Style.BackColor = Color.Bisque
                ''.Columns(7).DefaultCellStyle.BackColor = Color.Beige
                ''.Columns(8).HeaderCell.Style.BackColor = Color.Bisque
                ''.Columns(8).DefaultCellStyle.BackColor = Color.Beige
                ''.Columns(9).HeaderCell.Style.BackColor = Color.Bisque
                ''.Columns(9).DefaultCellStyle.BackColor = Color.Beige
                '.Columns(12).HeaderCell.Style.BackColor = Color.Bisque
                '.Columns(12).DefaultCellStyle.BackColor = Color.Beige
                '.Columns(13).HeaderCell.Style.BackColor = Color.Bisque
                '.Columns(13).DefaultCellStyle.BackColor = Color.Beige
                '.Columns(14).HeaderCell.Style.BackColor = Color.Bisque
                '.Columns(14).DefaultCellStyle.BackColor = Color.Beige
                ''20160516 FB構築に伴う自社口座情報の修正 -chg end

                '20160725 列幅修正 -del
                '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                .Columns(3).Frozen = True
                .Columns(0).ReadOnly = True
                .Columns(0).DefaultCellStyle.BackColor = Color.Azure
                .Columns(0).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(1).ReadOnly = True
                .Columns(1).DefaultCellStyle.BackColor = Color.Azure
                .Columns(1).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(2).ReadOnly = True
                .Columns(2).DefaultCellStyle.BackColor = Color.Azure
                .Columns(2).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(3).ReadOnly = True
                .Columns(3).DefaultCellStyle.BackColor = Color.Azure
                .Columns(3).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(4).ReadOnly = True
                .Columns(4).DefaultCellStyle.BackColor = Color.Azure
                .Columns(4).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(5).ReadOnly = True
                .Columns(5).DefaultCellStyle.BackColor = Color.Azure
                .Columns(5).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(6).ReadOnly = True
                .Columns(6).DefaultCellStyle.BackColor = Color.Azure
                .Columns(6).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(7).ReadOnly = True
                .Columns(7).DefaultCellStyle.BackColor = Color.Azure
                .Columns(7).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(8).ReadOnly = True
                .Columns(8).DefaultCellStyle.BackColor = Color.Azure
                .Columns(8).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(9).ReadOnly = True
                .Columns(9).DefaultCellStyle.BackColor = Color.Azure
                .Columns(9).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(10).ReadOnly = True
                .Columns(10).DefaultCellStyle.BackColor = Color.Azure
                .Columns(10).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(11).ReadOnly = True
                .Columns(11).DefaultCellStyle.BackColor = Color.Azure
                .Columns(11).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(12).ReadOnly = True
                .Columns(12).DefaultCellStyle.BackColor = Color.Azure
                .Columns(12).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(13).ReadOnly = True
                .Columns(13).DefaultCellStyle.BackColor = Color.Azure
                .Columns(13).HeaderCell.Style.BackColor = Color.PowderBlue
                '20160829 自社口座にデフォルト値を設定する処理を追加 -add sta
                .Columns(14).ReadOnly = True
                .Columns(14).DefaultCellStyle.BackColor = Color.Azure
                .Columns(14).HeaderCell.Style.BackColor = Color.PowderBlue
                '20160829 自社口座にデフォルト値を設定する処理を追加 -add end
                '20160829 自社口座にデフォルト値を設定する処理を追加 -del sta
                '.Columns(14).HeaderCell.Style.BackColor = Color.Bisque
                '.Columns(14).DefaultCellStyle.BackColor = Color.Beige
                '20160829 自社口座にデフォルト値を設定する処理を追加 -del end
                .Columns(15).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(15).DefaultCellStyle.BackColor = Color.Beige
                .Columns(16).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(16).DefaultCellStyle.BackColor = Color.Beige
                '20160829 自社口座にデフォルト値を設定する処理を追加 -add sta
                .Columns(17).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(17).DefaultCellStyle.BackColor = Color.Beige
                '20160829 自社口座にデフォルト値を設定する処理を追加 -add end
                '20160526 自社口座本体側の修正反映 -chg end
                '20160725 列幅修正 -add
                Call EtcMethod.Set_DgvColumnSize(dgv)
            End With

        End Sub

        ''' <summary>
        ''' 自社口座グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_jisyakoza_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end
            '20160905 自社口座が存在しない場合のエラー回避 -add sta
            If RelItem_M_jisyakoza Is Nothing Then
                Exit Sub
            End If
            '20160905 自社口座が存在しない場合のエラー回避 -add end
            'グリッド行数設定
            dgv.RowCount = RelItem_M_jisyakoza_Rev7.ItemCnt - 1

            '20160829 自社No、自社口座Noの自動作成処理を追加 -add
            Dim tmp_jisyano As Integer = 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_jisyakoza_Rev7.ItemCnt - 1
                '20160829 自社口座にデフォルト値を設定する処理を追加 -chg sta
                ''20160526 自社口座本体側の修正反映 -chg sta
                ''dgv(0, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kinyu_no(cntii)
                ''dgv(1, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kinyu_name(cntii)
                ''dgv(2, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Ten_no(cntii)
                ''dgv(3, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Ten_name(cntii)
                ''dgv(4, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kosyu_name(cntii)
                ''dgv(5, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_no(cntii)
                ''dgv(6, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_meigi(cntii)
                ' ''20160516 FB構築に伴う自社口座情報の修正 -add sta
                ''dgv(7, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_kana(cntii)
                ''dgv(8, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokigo1(cntii)
                ''dgv(9, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokigo2(cntii)
                ''dgv(10, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokozano(cntii)
                ''dgv(11, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.biko(cntii)
                ' ''20160516 FB構築に伴う自社口座情報の修正 -add end
                'dgv(0, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kozasyutokumoto(cntii)
                'dgv(1, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kozaname(cntii)
                'dgv(2, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kinyu_no(cntii)
                'dgv(3, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kinyu_name(cntii)
                'dgv(4, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Ten_no(cntii)
                'dgv(5, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Ten_name(cntii)
                'dgv(6, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kosyu_name(cntii)
                'dgv(7, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_no(cntii)
                'dgv(8, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_meigi(cntii)
                'dgv(9, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_kana(cntii)
                'dgv(10, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokigo1(cntii)
                'dgv(11, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokigo2(cntii)
                'dgv(12, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokozano(cntii)
                'dgv(13, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.biko(cntii)
                ''20160526 自社口座本体側の修正反映 -chg end
                dgv(0, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kozasyutokumoto(cntii)
                dgv(1, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kozasyutokumotono(cntii)
                dgv(2, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kozaname(cntii)
                dgv(3, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kinyu_no(cntii)
                dgv(4, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kinyu_name(cntii)
                dgv(5, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Ten_no(cntii)
                dgv(6, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Ten_name(cntii)
                dgv(7, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Kosyu_name(cntii)
                dgv(8, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_no(cntii)
                dgv(9, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_meigi(cntii)
                dgv(10, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Koza_kana(cntii)
                dgv(11, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokigo1(cntii)
                dgv(12, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokigo2(cntii)
                dgv(13, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.Yucyokozano(cntii)
                dgv(14, cntii - 1).Value = RelItem_M_jisyakoza_Rev7.biko(cntii)
                '20160829 自社No、自社口座Noの自動作成処理を追加 -add sta
                If RelItem_M_jisyakoza IsNot Nothing Then
                    dgv(15, cntii - 1).Value = "1"
                    dgv(16, cntii - 1).Value = RelItem_M_jisyakoza.Jisya_name(1)
                    dgv(17, cntii - 1).Value = tmp_jisyano.ToString
                    tmp_jisyano = tmp_jisyano + 1
                End If
                '20160829 自社No、自社口座Noの自動作成処理を追加 -add end
                '20160829 自社口座にデフォルト値を設定する処理を追加 -chg end
            Next

        End Sub

        ''' <summary>
        ''' 自社口座グリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックス表示対象外の列の場合は処理を抜ける
            If colcurrent <> 16 Then     '20160829 自社口座にデフォルト値を設定する処理を追加 15 → 16 
                Exit Sub
            End If

            'コンボリスト作成
            cellcombo.Items.Add("")

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvJisyakoza

            For cntii = 1 To RelItem_M_jisyakoza.itemcnt - 1
                cellcombo.Items.Add(RelItem_M_jisyakoza.Jisya_name(cntii))
            Next

            'カレントセルの値がコンボリストに含まれていない場合は新規に追加する (追加しないとエラーが生じるため)
            If currentvalue IsNot Nothing Then
                If cellcombo.Items.Contains(currentvalue) = False Then
                    cellcombo.Items.Add(currentvalue)
                End If
            End If

        End Sub

        ''' <summary>
        ''' 手入力値→一致したデータを表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_DgvMatchValue(ByVal dgv As DataGridView, ByVal colcurrent As Integer, ByVal rowcurrent As Integer, ByVal currentvalue As String) As Boolean

            Dim rtn As Boolean = True

            '対象列以外の場合は処理を抜ける
            If Not (colcurrent = 15 Or colcurrent = 16 Or colcurrent = 17) Then    '20160526 自社口座本体側の修正反映  7、8、9列→14、15、16列 '20160829 自社口座にデフォルト値を設定する処理を追加 14、15、16列→15、16、17列
                Return rtn
            End If

            '空文字が設定された場合はセルを消去する処理に移行して抜ける
            '2016.03.23 不要な箇所を削除していたため修正 -chg sta
            'If currentvalue = "" Then
            '    dgv(colcurrent - 1, rowcurrent).Value = ""
            '    For cntcol = colcurrent To dgv.ColumnCount - 2
            '        dgv(cntcol + 1, rowcurrent).Value = ""
            '    Next
            '    Return rtn
            'End If
            If currentvalue = "" Then
                Select Case colcurrent
                    Case 15      '20160526 自社口座本体側の修正反映  7列→14列 '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                    Case 16      '20160526 自社口座本体側の修正反映  8列→15列 '20160829 自社口座にデフォルト値を設定する処理を追加 15列→16列
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    Case 17      '20160526 自社口座本体側の修正反映  9列→16列 '20160829 自社口座にデフォルト値を設定する処理を追加 16列→17列
                End Select
                Return rtn
            End If
            '2016.03.23 不要な箇所を削除していたため修正 -chg end

            '数値範囲チェック
            '2016.03.23 数値チェック処理実装 -chg sta
            'If colcurrent = 7 Then
            '    Dim min As Integer = 1
            '    Dim max As Integer = 9999
            '    If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
            '        rtn = False
            '        Return rtn
            '    End If
            'End If

            Dim min As Integer = 1
            Dim max As Integer = 0

            Select Case colcurrent
                Case 15      '20160526 自社口座本体側の修正反映  7列→14列 '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
                    '20160526 自社口座本体側の修正反映 -chg sta
                    '設定された自社Noの値で判別するように修正
                    'max = 999999
                    'If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
                    '    MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    '    dgv(colcurrent, rowcurrent).Value = ""
                    '    dgv(colcurrent + 1, rowcurrent).Value = ""
                    'End If

                    Dim list_jisyano As New List(Of String)
                    For cntii = 1 To RelItem_M_jisyakoza.ItemCnt - 1
                        list_jisyano.Add(RelItem_M_jisyakoza.Jisya_no(cntii))
                    Next
                    If list_jisyano.Contains(currentvalue) = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    End If
                    '20160526 自社口座本体側の修正反映 -chg end
                Case 17      '20160526 自社口座本体側の修正反映  9列→16列 '20160829 自社口座にデフォルト値を設定する処理を追加 16列→17列
                    max = 999
                    If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        rtn = False
                        Return rtn
                    End If
            End Select
            '2016.03.23 数値チェック処理実装 -chg end

            '一致する値を取得
            '2016.03.23 対象列条件追加 -chg sta
            'For cntii = 1 To RelItem_M_jisyakoza.itemcnt - 1

            '    Dim searchstr As String = ""
            '    Dim outstr As String = ""
            '    Dim chgcol As Integer = 0

            '    Select Case colcurrent
            '        Case 7
            '            searchstr = RelItem_M_jisyakoza.Jisya_no(cntii)
            '            outstr = RelItem_M_jisyakoza.Jisya_name(cntii)
            '            chgcol = colcurrent + 1
            '        Case 8
            '            searchstr = RelItem_M_jisyakoza.Jisya_name(cntii)
            '            outstr = RelItem_M_jisyakoza.Jisya_no(cntii)
            '            chgcol = colcurrent - 1
            '    End Select

            '    dgv(chgcol, rowcurrent).Value = ""

            '    If currentvalue = searchstr Then
            '        dgv(chgcol, rowcurrent).Value = outstr
            '        Exit For
            '    End If

            'Next
            If colcurrent = 15 Or colcurrent = 16 Then        '20160526 自社口座本体側の修正反映  7、8列→14、15列 '20160829 自社口座にデフォルト値を設定する処理を追加 14、15列→15、16列
                For cntii = 1 To RelItem_M_jisyakoza.itemcnt - 1

                    Dim searchstr As String = ""
                    Dim outstr As String = ""
                    Dim chgcol As Integer = 0

                    Select Case colcurrent
                        Case 15      '20160526 自社口座本体側の修正反映  7列→14列 '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
                            searchstr = RelItem_M_jisyakoza.Jisya_no(cntii)
                            outstr = RelItem_M_jisyakoza.Jisya_name(cntii)
                            chgcol = colcurrent + 1
                        Case 16      '20160526 自社口座本体側の修正反映  8列→15列 '20160829 自社口座にデフォルト値を設定する処理を追加 15列→16列
                            searchstr = RelItem_M_jisyakoza.Jisya_name(cntii)
                            outstr = RelItem_M_jisyakoza.Jisya_no(cntii)
                            chgcol = colcurrent - 1
                    End Select

                    dgv(chgcol, rowcurrent).Value = ""

                    If currentvalue = searchstr Then
                        dgv(chgcol, rowcurrent).Value = outstr
                        Exit For
                    End If

                Next
            End If
            '2016.03.23 対象列条件追加 -chg end

            Return rtn

        End Function

        ''' <summary>
        ''' 自社口座の重複チェック
        ''' </summary>
        ''' <param name="rowno"></param>
        ''' <param name="initflg"></param>
        ''' <remarks></remarks>
        Public Sub Chk_Duplicate(ByVal rowno As Integer, ByVal initflg As Boolean)

            Dim dgv As DataGridView = Njc.Frm.RelationFrm.dgvJisyakoza
            Dim list_keytotal As New List(Of String)
            Dim colno_main As Integer = 15       '20160516 FB構築に伴う自社口座情報の修正  7列→14列 '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
            Dim colno_sub As Integer = 17        '20160516 FB構築に伴う自社口座情報の修正  9列→16列 '20160829 自社口座にデフォルト値を設定する処理を追加 16列→17列

            '現在の値を格納
            Dim key_taisyo_total As String = ""
            Dim key_taisyo_main As String = IIf(dgv(colno_main, rowno).Value Is Nothing, "", dgv(colno_main, rowno).Value)
            Dim key_taisyo_sub As String = IIf(dgv(colno_sub, rowno).Value Is Nothing, "", dgv(colno_sub, rowno).Value)

            'どちらかが空の場合は処理を抜ける
            If key_taisyo_main = "" Or key_taisyo_sub = "" Then
                Exit Sub
            Else
                key_taisyo_total = key_taisyo_main & "-" & key_taisyo_sub
            End If

            '対象列の全データ取得
            For cntii = 0 To dgv.RowCount - 1

                If cntii = rowno Then
                    GoTo skiplabel
                End If

                Dim tmp_key As String = IIf(dgv(colno_main, cntii).Value Is Nothing, "", dgv(colno_main, cntii).Value) & "-" & _
                                        IIf(dgv(colno_sub, cntii).Value Is Nothing, "", dgv(colno_sub, cntii).Value)

                If tmp_key <> "" Then
                    list_keytotal.Add(tmp_key)
                End If
skiplabel:
            Next

            '照合
            If list_keytotal.Contains(key_taisyo_total) Then
                If initflg = False Then
                    MsgResult = MessageBox.Show(MSG_DUPLICATE_A, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
                dgv(colno_sub, rowno).Value = ""
            End If

        End Sub

        ''' <summary>
        ''' セルの読み取り専用制御 '20160525 全体的な動作の修正
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Public Sub Set_CellReadOnly(ByVal colcurrent As Integer, ByVal rowcurrent As Integer)

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvJisyakoza

            '親項目取得用
            Dim tmp_jisyano As String = IIf(dgv(15, rowcurrent).Value Is Nothing, "", dgv(15, rowcurrent).Value)    '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
            Dim tmp_jisyaname As String = IIf(dgv(16, rowcurrent).Value Is Nothing, "", dgv(16, rowcurrent).Value)  '20160829 自社口座にデフォルト値を設定する処理を追加 15列→16列

            'セルの読み取り制御処理
            If tmp_jisyano.Trim & tmp_jisyaname.Trim = "" Then
                dgv(17, rowcurrent).ReadOnly = True     '20160829 自社口座にデフォルト値を設定する処理を追加 16列→17列
            Else
                dgv(17, rowcurrent).ReadOnly = False    '20160829 自社口座にデフォルト値を設定する処理を追加 16列→17列
            End If

        End Sub

        ''' <summary>
        ''' 親項目変更時の子項目削除 '20160725 親項目変更時の子項目制御
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Public Sub Chg_ChildValue(ByVal dgv As DataGridView, ByVal colcurrent As Integer, ByVal rowcurrent As Integer)

            Select Case colcurrent
                Case 15     '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
                    dgv(colcurrent + 2, rowcurrent).Value = ""
                Case 16     '20160829 自社口座にデフォルト値を設定する処理を追加 15列→16列
                    dgv(colcurrent + 1, rowcurrent).Value = ""
            End Select

        End Sub

        ''' <summary>
        ''' 紐付設定の件数表示 '20160825 紐付画面件数表示処理対応
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Set_RelCntHyoji()

            Dim dgv As New DataGridView
            Dim lbl_relcnt As New Label
            Dim lbl_totalcnt As New Label
            Dim lbl_cntback As New Label
            Dim relcnt As Integer = 0
            Dim totalcnt As Integer = 0

            'オブジェクトを格納
            dgv = Njc.Frm.RelationFrm.dgvJisyakoza
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiJisyaKoza
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllJisyaKoza

            '紐付全件数取得
            totalcnt = dgv.RowCount

            '紐付未設定件数取得
            For cntii = 0 To totalcnt - 1
                Dim tmp_no As String = dgv(15, cntii).Value             '20160829 自社口座にデフォルト値を設定する処理を追加 14列→15列
                Dim tmp_name As String = dgv(16, cntii).Value           '20160829 自社口座にデフォルト値を設定する処理を追加 15列→16列
                Dim tmp_narabino As String = dgv(17, cntii).Value       '20160829 自社口座にデフォルト値を設定する処理を追加 16列→17列

                If tmp_no <> "" And tmp_name <> "" And tmp_narabino <> "" Then
                    relcnt = relcnt + 1
                End If
            Next

            '表示
            lbl_relcnt.Text = relcnt.ToString
            lbl_totalcnt.Text = totalcnt.ToString

            '色設定
            Njc.Frm.RelationFrm.lblCntJisyaKozaBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiJisyaKoza.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllJisyaKoza.ForeColor = Color.Black

            '表示件数による色設定の制御
            If relcnt = totalcnt Then
                lbl_relcnt.ForeColor = Color.Black
                lbl_totalcnt.ForeColor = Color.Black
            Else
                lbl_relcnt.ForeColor = Color.Pink
                lbl_totalcnt.ForeColor = Color.Pink
            End If

        End Sub

    End Class

#End Region

End Namespace