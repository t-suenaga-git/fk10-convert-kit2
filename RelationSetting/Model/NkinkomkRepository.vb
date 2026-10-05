Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query

Namespace Njc.Repository

#Region "入金項目マスタ取得"

    Public Class M_nkin_Repository

        ''' <summary>
        ''' 10入金項目取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadDB(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim reccnt As Integer                                                           '抽出レコード件数
            Dim fldname As String                                                           '該当項目名
            Dim fldvalue As String                                                          '該当値
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True

            '抽出・レコード数の取得
            reccnt = DBExec.Exec_DataTable(Me.Get_UseQry(), sqlcnnV10, readtbl)

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_nkin_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii As Integer = 1 To readtbl.Rows.Count

                With model_item

                    For cntjj As Integer = 1 To readtbl.Columns.Count

                        '項目名取得
                        fldname = readtbl.Columns(cntjj - 1).ColumnName.ToString.Trim

                        '登録値取得
                        fldvalue = readtbl.Rows(cntii - 1).Item(cntjj - 1).ToString.Trim

                        '各項目編集処理
                        Select Case fldname
                            Case "nkin_no"
                                .Nkin_no(cntii) = fldvalue
                            Case "nkin_name"
                                .Nkin_name(cntii) = fldvalue
                            Case "nkin_ruino"
                                .Nkin_ruino(cntii) = fldvalue
                                .Nkin_ruiname(cntii) = Me.Get_Nkinruiname(fldvalue)
                            Case "nkin_zkseino"
                                .Nkin_zkseino(cntii) = fldvalue
                                .Nkin_zkseiname(cntii) = Me.Get_NkinZokuseiname(fldvalue)
                        End Select

                    Next

                End With

            Next

            'モデルの引渡し
            RelItem_M_nkin = Nothing
            RelItem_M_nkin = model_item

            '入金項目属性の取得
            Call Me.ReadDB_Zokusei()

            '変動費メーター分類の取得
            Call Me.ReadDB_Hendometer()

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10入金項目抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry()

            Dim strsql As String

            strsql = ""
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 [nkin_no] "
            strsql = strsql & " 	,[nkin_name] "
            strsql = strsql & " 	,[nkin_ruino] "
            strsql = strsql & " 	,[nkin_zkseino] "
            strsql = strsql & " FROM m_nkin "
            strsql = strsql & " ORDER BY [nkin_no] "

            Return strsql

        End Function

        ''' <summary>
        ''' 入金項目区分から区分名を取得
        ''' </summary>
        ''' <param name="tukikbn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Get_Nkinruiname(ByRef tukikbn As String) As String

            Dim rtn_string As String = ""

            Select Case tukikbn
                Case "1" : rtn_string = "通常月"
                Case "2" : rtn_string = "契約時"
                Case "3" : rtn_string = "更新時"
                Case "4" : rtn_string = "解約時"
                Case "5" : rtn_string = "随時変動"
                Case "6" : rtn_string = "その他"
                Case "7" : rtn_string = "修繕"
                Case "8" : rtn_string = "家主送金"
                Case "9" : rtn_string = "家主控除"
            End Select

            Return rtn_string

        End Function

        ''' <summary>
        ''' 入金項目属性Noから入金項目属性名を取得
        ''' </summary>
        ''' <param name="zokuseino"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Get_NkinZokuseiname(ByRef zokuseino As String) As String

            Dim rtn_string As String = ""

            Select Case zokuseino
                Case "100" : rtn_string = "賃貸料"
                Case "110" : rtn_string = "共益費"
                Case "120" : rtn_string = "駐車料"
                Case "130" : rtn_string = "公共料金"
                Case "140" : rtn_string = "振替手数料システム"
                Case "145" : rtn_string = "振替手数料"
                Case "190" : rtn_string = "その他"
                Case "200" : rtn_string = "敷金"
                Case "201" : rtn_string = "敷金_敷引・解約引"
                Case "210" : rtn_string = "保証金"
                Case "220" : rtn_string = "保証金_敷引・解約引"
                Case "230" : rtn_string = "礼金"
                Case "240" : rtn_string = "仲介手数料"
                Case "250" : rtn_string = "保険料"
                Case "260" : rtn_string = "家賃保証料"
                Case "290" : rtn_string = "その他"
                Case "300" : rtn_string = "更新料"
                Case "310" : rtn_string = "更新手数料"
                Case "320" : rtn_string = "敷金差額"
                Case "330" : rtn_string = "保証金差額"
                Case "390" : rtn_string = "その他"
                Case "400" : rtn_string = "敷金戻し"
                Case "410" : rtn_string = "敷金_敷引・解約引"
                Case "420" : rtn_string = "保証金戻し"
                Case "430" : rtn_string = "保証金_敷引・解約引"
                Case "440" : rtn_string = "解約違約金"
                Case "490" : rtn_string = "その他"
                Case "500" : rtn_string = "従量公共料金"
                Case "600" : rtn_string = "督促料"
                Case "610" : rtn_string = "督促手数料"
                Case "620" : rtn_string = "過剰金"
                Case "640" : rtn_string = "修繕費システム"
                Case "650" : rtn_string = "修繕費"
                Case "690" : rtn_string = "その他"
                Case "800" : rtn_string = "一括借上額"
                Case "810" : rtn_string = "支払金"
                Case "890" : rtn_string = "その他"
                Case "900" : rtn_string = "広告料"
                Case "910" : rtn_string = "修繕費システム"
                Case "915" : rtn_string = "修繕費"
                Case "920" : rtn_string = "その他経費"
                Case "930" : rtn_string = "公共料金"
                Case "940" : rtn_string = "管理手数料"
                Case "950" : rtn_string = "振込手数料"
                Case "990" : rtn_string = "その他"
                Case "1070" : rtn_string = "その他"
                Case "1090" : rtn_string = "雑費支払"
                Case "1170" : rtn_string = "その他"
                Case "1190" : rtn_string = "雑費収入"
            End Select

            Return rtn_string

        End Function

        ''' <summary>
        ''' 10入金項目属性取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ReadDB_Zokusei()

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim nkinzno As String
            Dim nkinzname As String
            Dim nkinruino As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String
            Dim tmp_col3() As String

            '値をまとめて格納
            nkinzno = "100,110,120,130,140,145,190,200,201,210,220,230,240,250,260,290,300,310,320,330,390,400,410,420,430,440,490,500,600,610,620,640,650,690,800,810,890,900,910,915,920,930,940,950,990,1070,1090,1170,1190"
            nkinzname = "賃貸料,共益費,駐車料,公共料金,振替手数料システム,振替手数料,その他,敷金,敷金_敷引・解約引,保証金,保証金_敷引・解約引,礼金,仲介手数料,保険料,家賃保証料,その他,更新料,更新手数料,敷金差額,保証金差額,その他,敷金戻し,敷金_敷引・解約引,保証金戻し,保証金_敷引・解約引,解約違約金,その他,従量公共料金,督促料,督促手数料,過剰金,修繕費システム,修繕費,その他,一括借上額,支払金,その他,広告料,修繕費システム,修繕費,その他経費,公共料金,管理手数料,振込手数料,その他,その他,雑費支払,その他,雑費収入"
            nkinruino = "1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,3,3,3,3,3,4,4,4,4,4,4,5,6,6,6,6,6,6,8,8,8,9,9,9,9,9,9,9,9,未使用,未使用,未使用,未使用"

            '配列に格納
            tmp_col1 = Split(nkinzno, ",")
            tmp_col2 = Split(nkinzname, ",")
            tmp_col3 = Split(nkinruino, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_nkin_z_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii = 0 To UBound(tmp_col1)

                With model_item
                    .Nkin_zkseino(cntii + 1) = tmp_col1(cntii)
                    .Nkin_zkseiname(cntii + 1) = tmp_col2(cntii)
                    .Nkin_ruino(cntii + 1) = tmp_col3(cntii)
                    .Nkin_ruiname(cntii + 1) = Me.Get_Nkinruiname(tmp_col3(cntii))
                End With

            Next

            'モデルの引渡し
            RelItem_M_nkin_z = Nothing
            RelItem_M_nkin_z = model_item

        End Sub

        ''' <summary>
        ''' 10入金項目変動費メーター分類取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ReadDB_Hendometer()

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim hendometerno As String
            Dim hendometername As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String

            '値をまとめて格納
            hendometerno = "1,2,3,4,5"
            hendometername = "水道,電気,ガス,灯油,その他"

            '配列に格納
            tmp_col1 = Split(hendometerno, ",")
            tmp_col2 = Split(hendometername, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_nkin_hendometer_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii = 0 To UBound(tmp_col1)

                With model_item
                    .Nkin_hendometerno(cntii + 1) = tmp_col1(cntii)
                    .Nkin_hendometername(cntii + 1) = tmp_col2(cntii)
                End With

            Next

            'モデルの引渡し
            RelItem_M_nkin_hendometer = Nothing
            RelItem_M_nkin_hendometer = model_item

        End Sub

        ''' <summary>
        ''' 入金項目グリッドの初期設定
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Initialize_Dgv(ByVal dgv As DataGridView)

            With dgv
                'kakaka6 del s -------------------------------------------------- 0420
                '.Height = 340
                '.Top = ((Njc.Frm.RelationFrm.tabPage10.Height - .Height) / 2) + 15
                'kakaka6 del e -------------------------------------------------- 0420
                '20160701 指摘事項まとめファイルの対応 幅の可変 -del
                '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
                '20160701 指摘事項まとめファイルの対応 幅の可変 -add
                '20160725 列幅修正 -del
                '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                .Columns(1).Frozen = True
                .Columns(0).ReadOnly = True
                .Columns(0).DefaultCellStyle.BackColor = Color.Azure
                .Columns(0).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(1).ReadOnly = True
                .Columns(1).DefaultCellStyle.BackColor = Color.Azure
                .Columns(1).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(2).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(2).DefaultCellStyle.BackColor = Color.Beige
                .Columns(3).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(3).DefaultCellStyle.BackColor = Color.Beige
                .Columns(4).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(4).DefaultCellStyle.BackColor = Color.Beige
                .Columns(5).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(5).DefaultCellStyle.BackColor = Color.Beige
                .Columns(6).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(6).DefaultCellStyle.BackColor = Color.Beige
                .Columns(7).ReadOnly = True
                .Columns(7).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(7).DefaultCellStyle.BackColor = Color.Beige
                .Columns(8).ReadOnly = True
                .Columns(8).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(8).DefaultCellStyle.BackColor = Color.Beige
                '20160525 入金項目紐付設定の修正 -add sta
                '未使用のため非表示にする
                .Columns(6).Visible = False
                '20160525 入金項目紐付設定の修正 -add end
                '20160725 スクロールバー表示修正 -del
                '.Height = 275   '20160622 入金項目No任意コード一括設定 -add
                '20160725 列幅修正 -add
                Call EtcMethod.Set_DgvColumnSize(dgv)
            End With

        End Sub

        ''' <summary>
        ''' 入金項目グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_nkin_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_nkin_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_nkin_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_nkin_Rev7.Nkin_name(cntii)
                dgv(1, cntii - 1).Value = RelItem_M_nkin_Rev7.Nkin_kbn(cntii)
            Next

            '10紐付項目表示
            Dim relitemcnt10 As Integer = RelItem_M_nkin.ItemCnt - 1

            For cntii = 1 To dgv.RowCount

                '対象文字列格納
                Dim taisyostr As String = EtcMethod.Get_Str_Shape(dgv(0, cntii - 1).Value)
                Dim tmp_taisyokbn As String = dgv(1, cntii - 1).Value
                Dim tmp_str() As String = tmp_taisyokbn.Split(".")
                Dim taisyokbn As String = tmp_str(0)

                '入金項目/入金項目属性まで照合
                For cntjj = 1 To relitemcnt10

                    '2016.03.23 グリッド表示時点で制御する必要があるため移動 -add sta
                    '変動費メーター分類の読み取り専用制御
                    If taisyokbn = "5" Then
                        dgv(7, cntii - 1).ReadOnly = False
                        dgv(8, cntii - 1).ReadOnly = False
                    End If
                    '2016.03.23 グリッド表示時点で制御する必要があるため移動 -add end

                    '照合用文字列格納
                    Dim searchstr As String = EtcMethod.Get_Str_Shape(RelItem_M_nkin.Nkin_name(cntjj))
                    Dim searchkbn As String = RelItem_M_nkin.Nkin_ruino(cntjj)

                    If taisyostr = taisyostr.Replace(searchstr, "") Then
                        'dgv(3, cntii - 1).Value = taisyostr
                        'dgv(3, cntii - 1).Style.BackColor = Color.Pink
                        'dgv(2, cntii - 1).Style.BackColor = Color.Pink
                    ElseIf taisyokbn = searchkbn And taisyostr <> taisyostr.Replace(searchstr, "") Then
                        dgv(2, cntii - 1).Value = RelItem_M_nkin.Nkin_no(cntjj)
                        dgv(3, cntii - 1).Value = RelItem_M_nkin.Nkin_name(cntjj)
                        dgv(4, cntii - 1).Value = RelItem_M_nkin.Nkin_zkseino(cntjj)
                        dgv(5, cntii - 1).Value = RelItem_M_nkin.Nkin_zkseiname(cntjj)
                        dgv(6, cntii - 1).Value = RelItem_M_nkin.Nkin_ruiname(cntjj)
                        dgv(2, cntii - 1).Style.BackColor = Color.Beige
                        dgv(3, cntii - 1).Style.BackColor = Color.Beige
                        Exit For
                    End If

                Next

                '変動費メーター分類照合
                If taisyokbn = "5" Then

                    Dim hendocnt As Integer = RelItem_M_nkin_hendometer.ItemCnt - 1

                    For cntkk = 1 To hendocnt

                        Dim hendostr As String = EtcMethod.Get_Str_Shape(RelItem_M_nkin_hendometer.Nkin_hendometername(cntkk))

                        If taisyostr <> taisyostr.Replace(hendostr, "") Then
                            dgv(7, cntii - 1).Value = RelItem_M_nkin_hendometer.Nkin_hendometerno(cntkk)
                            dgv(8, cntii - 1).Value = RelItem_M_nkin_hendometer.Nkin_hendometername(cntkk)
                            Exit For
                        End If

                    Next

                End If

            Next

            '重複チェック
            For cntrow = dgv.RowCount - 1 To 0 Step -1
                Call Me.Chk_Duplicate(cntrow, True)
            Next

            '20161209 前回設定値復元時にグレイアウトされていないセルが生じるエラーの修正 -del sta
            'フォーム画面にメソッドを移動するためコメントアウト
            ''20160701 指摘事項まとめファイルの対応 随時変動費以外の変動費メーター分類列のグレイアウト対応 -add sta
            ''入金項目区分が入出金以外の場合、自社情報設定箇所をグレイアウトさせる
            'For cntkk = 1 To dgv.RowCount
            '    Dim fmtkbn As String = dgv(1, cntkk - 1).Value
            '    If fmtkbn <> "5.随時変動" Then
            '        dgv(7, cntkk - 1).Style.BackColor = Color.LightGray
            '        dgv(8, cntkk - 1).Style.BackColor = Color.LightGray
            '    End If
            'Next
            ''20160701 指摘事項まとめファイルの対応 随時変動費以外の変動費メーター分類列のグレイアウト対応 -add end
            '20161209 前回設定値復元時にグレイアウトされていないセルが生じるエラーの修正 -del end

        End Sub

        ''' <summary>
        ''' 入金項目グリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal rowcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックス表示対象外の列の場合は処理を抜ける
            If Not (colcurrent = 3 Or colcurrent = 5 Or colcurrent = 8) Then
                Exit Sub
            End If

            'コンボリスト作成
            cellcombo.Items.Add("")

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvNkinkomk

            ''20160525 入金項目紐付設定の修正 -add sta
            'If currentvalue <> "" Then
            '    cellcombo.Items.Add(currentvalue)
            'End If
            ''20160525 入金項目紐付設定の修正 -add end

            '親項目取得用
            Dim tmp_kbn As String = ""
            Dim tmp_kbntotal As String = IIf(dgv(1, rowcurrent).Value Is Nothing, "", dgv(1, rowcurrent).Value)
            Dim tmp_nkinno As String = IIf(dgv(2, rowcurrent).Value Is Nothing, "", dgv(2, rowcurrent).Value)
            If tmp_kbntotal <> "" Then
                Dim tmp_str() As String = tmp_kbntotal.Split(".")
                tmp_kbn = tmp_str(0)
            End If

            Select Case colcurrent
                Case 3
                    For cntii = 1 To RelItem_M_nkin.itemcnt - 1
                        If tmp_kbn = RelItem_M_nkin.Nkin_ruino(cntii) Then
                            cellcombo.Items.Add(RelItem_M_nkin.Nkin_name(cntii))
                        End If
                    Next
                Case 5
                    For cntii = 1 To RelItem_M_nkin_z.itemcnt - 1
                        If tmp_nkinno <> "" And tmp_kbn = RelItem_M_nkin_z.Nkin_ruino(cntii) Then
                            cellcombo.Items.Add(RelItem_M_nkin_z.Nkin_zkseiname(cntii))
                        End If
                    Next
                Case 8
                    For cntii = 1 To RelItem_M_nkin_hendometer.itemcnt - 1
                        If tmp_nkinno <> "" And tmp_kbn = "5" Then
                            '2016.03.23 グリッド表示時点で制御する必要があるため移動 -del sta
                            'dgv(7, rowcurrent).ReadOnly = False
                            'dgv(8, rowcurrent).ReadOnly = False
                            '2016.03.23 グリッド表示時点で制御する必要があるため移動 -del sta
                            cellcombo.Items.Add(RelItem_M_nkin_hendometer.Nkin_hendometername(cntii))
                        End If
                    Next
            End Select

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
                    Case 2
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                        dgv(colcurrent + 4, rowcurrent).Value = ""
                        dgv(colcurrent + 5, rowcurrent).Value = ""
                        dgv(colcurrent + 6, rowcurrent).Value = ""
                    Case 3
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                        dgv(colcurrent + 4, rowcurrent).Value = ""
                        dgv(colcurrent + 5, rowcurrent).Value = ""
                    Case 4
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                        dgv(colcurrent + 4, rowcurrent).Value = ""
                    Case 5
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                    Case 6
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                    Case 7
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    Case 8
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                End Select
                Return rtn
            End If
            '2016.03.23 不要な箇所を削除していたため修正 -chg end

            '数値範囲チェック
            '2016.03.23 最大最小値チェック処理追加 -chg sta
            'If colcurrent = 2 Then
            '    Dim min As Integer = 1000
            '    Dim max As Integer = 9999
            '    If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
            '        rtn = False
            '        Return rtn
            '    End If
            'End If

            '入金項目区分を取得
            Dim tmp_nkinkbn As String = IIf(dgv(1, rowcurrent).Value Is Nothing, "", dgv(1, rowcurrent).Value)
            Dim tmp_str() As String = tmp_nkinkbn.Split(".")
            Dim nkinkbn As String = tmp_str(0)

            '作業用変数生成
            Dim min_nkinno As Integer = 0
            Dim max_nkinno As Integer = 0
            Dim min_nkinhendometerno As Integer = 1
            Dim max_nkinhendometerno = 5

            '入金項目No最大最小値の取得 (入金項目区分によって値が変化)
            Call Me.Set_NkinkbnToNkinNoRange(nkinkbn, min_nkinno, max_nkinno)

            Select Case colcurrent
                Case 2  '入金項目No
                    If EtcMethod.Chk_NumRange(currentvalue, min_nkinno, max_nkinno) = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        '2016.03.23 削除箇所が不足していたため追加 -add sta
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                        dgv(colcurrent + 4, rowcurrent).Value = ""
                        dgv(colcurrent + 5, rowcurrent).Value = ""
                        dgv(colcurrent + 6, rowcurrent).Value = ""
                        '2016.03.23 削除箇所が不足していたため追加 -add end
                        rtn = False
                        Return rtn
                    End If
                Case 4  '入金項目属性No
                    If Me.Chk_NkinkbnToNkinzokusei(nkinkbn, currentvalue) = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        '2016.03.23 削除箇所が不足していたため追加 -add sta
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                        dgv(colcurrent + 4, rowcurrent).Value = ""
                        '2016.03.23 削除箇所が不足していたため追加 -add end
                        rtn = False
                        Return rtn
                    End If
                Case 7  '変動費メーター分類No
                    If EtcMethod.Chk_NumRange(currentvalue, min_nkinhendometerno, max_nkinhendometerno) = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        rtn = False
                        Return rtn
                    End If
            End Select
            '2016.03.23 最大最小値チェック処理追加 -chg end

            '一致する値を取得
            '20160525 入金項目紐付設定の修正 -chg sta
            'For cntii = 1 To RelItem_M_nkin.itemcnt - 1


            'Dim searchstr As String = ""
            'Dim outstr As String = ""
            'Dim chgcol As Integer = 0

            'Select Case colcurrent
            '    Case 2
            '        searchstr = RelItem_M_nkin.Nkin_no(cntii)
            '        outstr = RelItem_M_nkin.Nkin_name(cntii)
            '        chgcol = colcurrent + 1
            '    Case 3
            '        searchstr = RelItem_M_nkin.Nkin_name(cntii)
            '        outstr = RelItem_M_nkin.Nkin_no(cntii)
            '        chgcol = colcurrent - 1
            '    Case 4
            '        searchstr = RelItem_M_nkin_z.Nkin_zkseino(cntii)
            '        outstr = RelItem_M_nkin_z.Nkin_zkseiname(cntii)
            '        chgcol = colcurrent + 1
            '    Case 5
            '        searchstr = RelItem_M_nkin_z.Nkin_zkseiname(cntii)
            '        outstr = RelItem_M_nkin_z.Nkin_zkseino(cntii)
            '        chgcol = colcurrent - 1
            '    Case 7
            '        searchstr = RelItem_M_nkin_hendometer.Nkin_hendometerno(cntii)
            '        outstr = RelItem_M_nkin_hendometer.Nkin_hendometername(cntii)
            '        chgcol = colcurrent + 1
            '    Case 8
            '        searchstr = RelItem_M_nkin_hendometer.Nkin_hendometername(cntii)
            '        outstr = RelItem_M_nkin_hendometer.Nkin_hendometerno(cntii)
            '        chgcol = colcurrent - 1
            'End Select

            'dgv(chgcol, rowcurrent).Value = ""

            'If currentvalue = searchnkinname Then
            '    dgv(chgcol, rowcurrent).Value = outstr
            '    Exit For
            'End If

            Dim inputsyogo As String = nkinkbn & "-" & currentvalue
            Dim searchnkinkbn As String = ""
            Dim searchstr As String = ""
            Dim outstr As String = ""
            Dim chgcol As Integer = 0

            Select Case colcurrent
                Case 2
                    chgcol = colcurrent + 1
                    For cntii = 1 To RelItem_M_nkin.itemcnt - 1
                        searchnkinkbn = RelItem_M_nkin.Nkin_ruino(cntii)
                        searchstr = RelItem_M_nkin.Nkin_no(cntii)
                        outstr = RelItem_M_nkin.Nkin_name(cntii)
                        Dim outputsyogo As String = searchnkinkbn & "-" & searchstr
                        If inputsyogo = outputsyogo Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            '20161227 入金項目属性の自動設定処理の追加 -add sta
                            dgv(colcurrent + 2, rowcurrent).Value = RelItem_M_nkin.Nkin_zkseino(cntii)
                            dgv(colcurrent + 3, rowcurrent).Value = RelItem_M_nkin.Nkin_zkseiname(cntii)
                            '20161227 入金項目属性の自動設定処理の追加 -add end
                            Exit For
                        Else
                            outstr = ""
                        End If
                    Next
                Case 3
                    chgcol = colcurrent - 1
                    For cntii = 1 To RelItem_M_nkin.itemcnt - 1
                        searchnkinkbn = RelItem_M_nkin.Nkin_ruino(cntii)
                        searchstr = RelItem_M_nkin.Nkin_name(cntii)
                        outstr = RelItem_M_nkin.Nkin_no(cntii)
                        Dim outputsyogo As String = searchnkinkbn & "-" & searchstr
                        If inputsyogo = outputsyogo Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            '20161227 入金項目属性の自動設定処理の追加 -add sta
                            dgv(colcurrent + 1, rowcurrent).Value = RelItem_M_nkin.Nkin_zkseino(cntii)
                            dgv(colcurrent + 2, rowcurrent).Value = RelItem_M_nkin.Nkin_zkseiname(cntii)
                            '20161227 入金項目属性の自動設定処理の追加 -add end
                            Exit For
                        Else
                            outstr = ""
                        End If
                    Next

                Case 4
                    chgcol = colcurrent + 1
                    For cntii = 1 To RelItem_M_nkin_z.itemcnt - 1
                        searchnkinkbn = RelItem_M_nkin_z.Nkin_ruino(cntii)
                        searchstr = RelItem_M_nkin_z.Nkin_zkseino(cntii)
                        outstr = RelItem_M_nkin_z.Nkin_zkseiname(cntii)
                        Dim outputsyogo As String = searchnkinkbn & "-" & searchstr
                        If inputsyogo = outputsyogo Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        Else
                            outstr = ""
                        End If
                    Next
                Case 5
                    chgcol = colcurrent - 1
                    For cntii = 1 To RelItem_M_nkin_z.itemcnt - 1
                        searchnkinkbn = RelItem_M_nkin_z.Nkin_ruino(cntii)
                        searchstr = RelItem_M_nkin_z.Nkin_zkseiname(cntii)
                        outstr = RelItem_M_nkin_z.Nkin_zkseino(cntii)
                        Dim outputsyogo As String = searchnkinkbn & "-" & searchstr
                        If inputsyogo = outputsyogo Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        Else
                            outstr = ""
                        End If
                    Next
                Case 7
                    chgcol = colcurrent + 1
                    For cntii = 1 To RelItem_M_nkin_hendometer.itemcnt - 1
                        searchnkinkbn = "5"
                        searchstr = RelItem_M_nkin_hendometer.Nkin_hendometerno(cntii)
                        outstr = RelItem_M_nkin_hendometer.Nkin_hendometername(cntii)
                        Dim outputsyogo As String = searchnkinkbn & "-" & searchstr
                        If inputsyogo = outputsyogo Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        Else
                            outstr = ""
                        End If
                    Next
                Case 8
                    chgcol = colcurrent - 1
                    For cntii = 1 To RelItem_M_nkin_hendometer.itemcnt - 1
                        searchnkinkbn = "5"
                        searchstr = RelItem_M_nkin_hendometer.Nkin_hendometername(cntii)
                        outstr = RelItem_M_nkin_hendometer.Nkin_hendometerno(cntii)
                        Dim outputsyogo As String = searchnkinkbn & "-" & searchstr
                        If inputsyogo = outputsyogo Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        Else
                            outstr = ""
                        End If
                    Next

            End Select

            '入力列が入金項目Noかつ一致する文字列が存在しない場合はV7の入金項目を出力する
            If colcurrent = 2 And outstr = "" Then
                Dim tmp_V7nkinkomkname As String = IIf(dgv(0, rowcurrent).Value Is Nothing, "", dgv(0, rowcurrent).Value)

                If TypeOf dgv(chgcol, rowcurrent) Is DataGridViewComboBoxCell Then
                    Dim cellcombo As New DataGridViewComboBoxCell
                    cellcombo = DirectCast(dgv(chgcol, rowcurrent), DataGridViewComboBoxCell)

                    If cellcombo.Items.Contains(tmp_V7nkinkomkname) = False Then
                        cellcombo.Items.Add(tmp_V7nkinkomkname)
                    End If

                End If

                dgv(chgcol, rowcurrent).Value = tmp_V7nkinkomkname

            End If
            '20160525 入金項目紐付設定の修正 -chg end

            Return rtn

        End Function

        ''' <summary>
        ''' 入金項目の重複チェック
        ''' </summary>
        ''' <param name="rowno"></param>
        ''' <param name="initflg"></param>
        ''' <remarks></remarks>
        Public Sub Chk_Duplicate(ByVal rowno As Integer, ByVal initflg As Boolean)

            Dim dgv As DataGridView = Njc.Frm.RelationFrm.dgvNkinkomk
            Dim list_keytotal As New List(Of String)
            Dim colno As Integer = 2

            '現在の値を格納
            Dim key_taisyo As String = IIf(dgv(colno, rowno).Value Is Nothing, "", dgv(colno, rowno).Value)
            If key_taisyo = "" Then
                Exit Sub
            End If

            '対象列の全データ取得
            For cntii = 0 To dgv.RowCount - 1

                If cntii = rowno Then
                    GoTo skiplabel
                End If

                Dim tmp_key As String = IIf(dgv(colno, cntii).Value Is Nothing, "", dgv(colno, cntii).Value)

                If tmp_key <> "" Then
                    list_keytotal.Add(tmp_key)
                End If
skiplabel:
            Next

            '照合
            If list_keytotal.Contains(key_taisyo) Then
                If initflg = False Then
                    MsgResult = MessageBox.Show(MSG_DUPLICATE_A, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
                For cntjj = 2 To dgv.ColumnCount - 1
                    dgv(cntjj, rowno).Value = ""
                Next
            End If

        End Sub

        ''' <summary>
        ''' 入金項目区分毎の入金項目No最大/最小値の取得
        ''' </summary>
        ''' <param name="nkinkbn"></param>
        ''' <param name="min_nkinno"></param>
        ''' <param name="max_nkinno"></param>
        ''' <remarks></remarks>
        Public Sub Set_NkinkbnToNkinNoRange(ByVal nkinkbn As String, _
                                            ByRef min_nkinno As Integer, ByRef max_nkinno As Integer)

            Dim tmp_nkinkbn As Integer = 0
            If (Int32.TryParse(nkinkbn, tmp_nkinkbn)) Then
                min_nkinno = tmp_nkinkbn * 1000
                max_nkinno = min_nkinno + 999
            End If

        End Sub

        ''' <summary>
        ''' 入金項目属性Noの照合 (入金項目属性Noは連番ではない指定された値が存在されるため数値を直接照合する)
        ''' </summary>
        ''' <param name="nkinkbn"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Chk_NkinkbnToNkinzokusei(ByVal nkinkbn As String, ByVal value As String) As Boolean

            Dim rtn As Boolean = True

            '入金項目区分と入金項目属性Noを紐付けたリストを生成
            Dim list_nkizokusei As New List(Of String)

            For cntii = 1 To RelItem_M_nkin_z.Itemcnt - 1

                Dim tmp_nkinkbn As String = RelItem_M_nkin_z.Nkin_ruino(cntii)
                Dim tmp_nkinzno As String = RelItem_M_nkin_z.Nkin_zkseino(cntii)

                list_nkizokusei.Add(tmp_nkinkbn & "-" & tmp_nkinzno)

            Next

            '照合用に引数を加工
            Dim chkvalue As String = nkinkbn & "-" & value

            '照合
            If list_nkizokusei.Contains(chkvalue) = False Then
                rtn = False
                Return rtn
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' セルの読み取り専用制御 '20160525 全体的な動作の修正
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Public Sub Set_CellReadOnly(ByVal colcurrent As Integer, ByVal rowcurrent As Integer)

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvNkinkomk

            '親項目取得用
            Dim tmp_nkinno As String = IIf(dgv(2, rowcurrent).Value Is Nothing, "", dgv(2, rowcurrent).Value)
            Dim tmp_nkinname As String = IIf(dgv(3, rowcurrent).Value Is Nothing, "", dgv(3, rowcurrent).Value)

            'セルの読み取り制御処理
            If tmp_nkinno.Trim & tmp_nkinname.Trim = "" Then
                dgv(4, rowcurrent).ReadOnly = True
                dgv(5, rowcurrent).ReadOnly = True
            Else
                dgv(4, rowcurrent).ReadOnly = False
                dgv(5, rowcurrent).ReadOnly = False
            End If

        End Sub

        ''' <summary>
        ''' 変動費セルの読み取り専用制御 '20160525 入金項目紐付設定の修正
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Public Sub Set_CellReadOnly_Hendo(ByVal colcurrent As Integer, ByVal rowcurrent As Integer)

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvNkinkomk

            'V7入金項目区分取得
            Dim tmp_kbn As String = IIf(dgv(1, rowcurrent).Value Is Nothing, "", dgv(1, rowcurrent).Value)

            '10入金項目、属性取得
            Dim tmp_nkinkomkno As String = IIf(dgv(2, rowcurrent).Value Is Nothing, "", dgv(2, rowcurrent).Value)
            Dim tmp_nkinzokuseino As String = IIf(dgv(4, rowcurrent).Value Is Nothing, "", dgv(4, rowcurrent).Value)

            'セルの読み取り制御処理
            'V7の入金項目区分が「5.随時変動」かつ10の入金項目No、入金項目属性に値が設定されている場合のみ変動費設定セルの読み取り専用を解除する
            If tmp_kbn.Trim = "5.随時変動" And tmp_nkinkomkno.Trim <> "" And tmp_nkinzokuseino.Trim <> "" Then
                dgv(7, rowcurrent).ReadOnly = False
                dgv(8, rowcurrent).ReadOnly = False
            Else
                dgv(7, rowcurrent).ReadOnly = True
                dgv(8, rowcurrent).ReadOnly = True
            End If

        End Sub

        ''' <summary>
        ''' 入金項目No未設定箇所への任意No一括設定 '20160622 入金項目No任意コード一括設定 -add
        ''' </summary>
        ''' <param name="cnt"></param>
        ''' <remarks></remarks>
        Public Sub Set_NkinnoBulk(ByVal cnt As Integer, ByRef chgcolorrow As List(Of Integer))

            'DataGridView設定
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvNkinkomk

            '既に設定されているデータを取得
            Dim list_existnkinno As New List(Of String)
            For cntii = 0 To dgv.RowCount - 1
                Dim tmp_existitem As String = IIf(dgv(2, cntii).Value Is Nothing, "", dgv(2, cntii).Value)
                If tmp_existitem <> "" Then
                    list_existnkinno.Add(tmp_existitem)
                End If
            Next

            '革命マスタに登録されているデータを取得
            For cntii = 1 To RelItem_M_nkin.Itemcnt - 1
                Dim tmp_mstitem As String = RelItem_M_nkin.Nkin_no(cntii)
                If list_existnkinno.Contains(tmp_mstitem) = False Then
                    list_existnkinno.Add(tmp_mstitem)
                End If
            Next

            Dim addcnt As Integer = 0               '開始No
            Dim oldnkinkbnint As Integer = 0        '入金区分変更判別用
            Dim maxnkinno As Integer = 0

            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add
            chgcolorrow.Clear()

            For cntii = 0 To dgv.RowCount - 1

                '入金項目Noが設定されているか判別
                Dim tmp_nkinexistno As String = IIf(dgv(2, cntii).Value Is Nothing, "", dgv(2, cntii).Value)

                '入金項目Noが設定されていない場合は任意のNoを設定する
                If tmp_nkinexistno = "" Then

                    '入金区分の取得(4桁目の数値設定)
                    Dim tmp_nkinkbn As String = IIf(dgv(1, cntii).Value Is Nothing, "", dgv(1, cntii).Value)
                    Dim nkinkbnint As Integer = 0
                    If tmp_nkinkbn <> "" Then
                        Dim tmp_str() As String = tmp_nkinkbn.Split(".")
                        nkinkbnint = Int32.Parse(tmp_str(0)) * 1000
                        maxnkinno = nkinkbnint + 999
                    End If

                    '入金区分が異なる行へ遷移した場合は開始Noを初期化する
                    If nkinkbnint <> oldnkinkbnint Then
                        addcnt = 0
                    End If
                    oldnkinkbnint = nkinkbnint

                    '既存設定値と重複しないNoを取得するまでループ処理
                    Dim containflg As Boolean = True
                    Dim setnkinno As Integer = 0
                    Do Until containflg = False
                        setnkinno = nkinkbnint + cnt + addcnt
                        containflg = list_existnkinno.Contains(setnkinno.ToString)
                        addcnt = addcnt + 1
                    Loop

                    '最大値を超えていない場合は入金項目Noを設定
                    If setnkinno <= maxnkinno Then

                        '取得したNoを設定
                        dgv(2, cntii).Value = setnkinno.ToString
                        list_existnkinno.Add(setnkinno.ToString)

                        '入金項目名を出力
                        Call Me.Set_DgvMatchValue(dgv, 2, cntii, setnkinno.ToString)

                        '入金項目属性をその他へ一括設定(変動費の場合はその他が存在しないので「従量公共料金」へ設定する)
                        Dim tmp_nkinzkseiname As String = ""
                        If nkinkbnint = 5000 Then
                            tmp_nkinzkseiname = "従量公共料金"
                        Else
                            tmp_nkinzkseiname = "その他"
                        End If
                        '20161227 入金項目一括設定エラーの修正 -chg sta
                        'dgv(5, cntii).Value = tmp_nkinzkseiname
                        'Call Me.Set_DgvMatchValue(dgv, 5, cntii, tmp_nkinzkseiname)
                        If TypeOf dgv(5, cntii) Is DataGridViewComboBoxCell Then
                            'コンボボックスの場合はコンボボックスを再作成してセット
                            Dim cellcombo As New DataGridViewComboBoxCell
                            Call Me.Set_DgvCmbList(5, cntii, tmp_nkinzkseiname, cellcombo)
                            dgv(5, cntii) = cellcombo
                            cellcombo.DisplayStyleForCurrentCellOnly = True
                            dgv(5, cntii).Value = tmp_nkinzkseiname
                            Call Me.Set_DgvMatchValue(dgv, 5, cntii, tmp_nkinzkseiname)
                        Else
                            'コンボボックスでない場合はそのまま設定
                            dgv(5, cntii).Value = tmp_nkinzkseiname
                            Call Me.Set_DgvMatchValue(dgv, 5, cntii, tmp_nkinzkseiname)
                        End If
                        '20161227 入金項目一括設定エラーの修正 -chg end


                        '20160725 入金項目一括設定_変動費メーターへの反映 -add sta
                        If nkinkbnint = 5000 Then
                            Dim tmp_metername As String = "その他"
                            '20161227 入金項目一括設定エラーの修正 -chg sta
                            'dgv(8, cntii).Value = tmp_metername
                            'Call Me.Set_DgvMatchValue(dgv, 8, cntii, tmp_metername)
                            If TypeOf dgv(8, cntii) Is DataGridViewComboBoxCell Then
                                'コンボボックスの場合はコンボボックスを再作成してセット
                                Dim cellcombo As New DataGridViewComboBoxCell
                                Call Me.Set_DgvCmbList(8, cntii, tmp_metername, cellcombo)
                                dgv(8, cntii) = cellcombo
                                cellcombo.DisplayStyleForCurrentCellOnly = True
                                dgv(8, cntii).Value = tmp_metername
                                Call Me.Set_DgvMatchValue(dgv, 8, cntii, tmp_metername)
                            Else
                                'コンボボックスでない場合はそのまま設定
                                dgv(8, cntii).Value = tmp_metername
                                Call Me.Set_DgvMatchValue(dgv, 8, cntii, tmp_metername)
                            End If
                            '20161227 入金項目一括設定エラーの修正 -chg end
                        End If
                        '20160725 入金項目一括設定_変動費メーターへの反映 -add end
                        '設定した行の色を変更する
                        For cntjj = 2 To dgv.ColumnCount - 1
                            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -chg sta
                            'dgv(cntjj, cntii).Style.BackColor = Color.NavajoWhite
                            If cntjj <= 5 Then
                                dgv(cntjj, cntii).Style.BackColor = Color.NavajoWhite
                                '20160725 入金項目一括設定_変動費メーターへの反映 -add sta
                            ElseIf nkinkbnint = 5000 Then
                                dgv(cntjj, cntii).Style.BackColor = Color.NavajoWhite
                                '20160725 入金項目一括設定_変動費メーターへの反映 -add end
                            End If
                            '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -chg end
                        Next

                        '20160701 指摘事項まとめファイルの対応 一括設定時のセルの着色に関する処理の修正 -add
                        chgcolorrow.Add(cntii)

                    End If

                End If



            Next

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
                Case 2
                    '20161227 入金項目属性の自動設定処理の追加 -del sta
                    '入金項目変更時は入金項目属性も自動設定するようにしたためコメントアウト
                    'dgv(colcurrent + 2, rowcurrent).Value = ""
                    'dgv(colcurrent + 3, rowcurrent).Value = ""
                    '20161227 入金項目属性の自動設定処理の追加 -del end
                    dgv(colcurrent + 4, rowcurrent).Value = ""
                    dgv(colcurrent + 5, rowcurrent).Value = ""
                    dgv(colcurrent + 6, rowcurrent).Value = ""
                Case 3
                    '入金項目変更時は入金項目属性も自動設定するようにしたためコメントアウト
                    '20161227 入金項目属性の自動設定処理の追加 -del sta
                    'dgv(colcurrent + 1, rowcurrent).Value = ""
                    'dgv(colcurrent + 2, rowcurrent).Value = ""
                    '20161227 入金項目属性の自動設定処理の追加 -del end
                    dgv(colcurrent + 3, rowcurrent).Value = ""
                    dgv(colcurrent + 4, rowcurrent).Value = ""
                    dgv(colcurrent + 5, rowcurrent).Value = ""
                Case 4
                    dgv(colcurrent + 2, rowcurrent).Value = ""
                    dgv(colcurrent + 3, rowcurrent).Value = ""
                Case 5
                    dgv(colcurrent + 1, rowcurrent).Value = ""
                    dgv(colcurrent + 2, rowcurrent).Value = ""
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
            dgv = Njc.Frm.RelationFrm.dgvNkinkomk
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiNkin
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllNkin

            '紐付全件数取得
            totalcnt = dgv.RowCount

            '紐付未設定件数取得
            For cntii = 0 To totalcnt - 1

                Dim tmp_kbn As String = dgv(1, cntii).Value
                Dim tmp_no As String = dgv(2, cntii).Value
                Dim tmp_name As String = dgv(3, cntii).Value
                Dim tmp_zno As String = dgv(4, cntii).Value
                Dim tmp_zname As String = dgv(5, cntii).Value
                Dim tmp_hno As String = dgv(7, cntii).Value
                Dim tmp_hname As String = dgv(8, cntii).Value
                Dim tmp_chkstr As String = tmp_no & tmp_name
                Dim tmp_chkstrz As String = tmp_zno & tmp_zname
                Dim tmp_chkstrh As String = tmp_hno & tmp_hname

                If tmp_kbn <> "5.随時変動" And tmp_chkstr <> "" And tmp_chkstrz <> "" Then
                    relcnt = relcnt + 1
                ElseIf tmp_kbn = "5.随時変動" And tmp_chkstr <> "" And tmp_chkstrz <> "" And tmp_chkstrh <> "" Then
                    relcnt = relcnt + 1
                End If

            Next

            '表示
            lbl_relcnt.Text = relcnt.ToString
            lbl_totalcnt.Text = totalcnt.ToString

            '色設定
            Njc.Frm.RelationFrm.lblCntNkinBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiNkin.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllNkin.ForeColor = Color.Black

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