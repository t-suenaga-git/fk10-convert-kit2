Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query

Namespace Njc.Repository

#Region "入金区分マスタ取得"

    Public Class M_nkbn_Repository

        ''' <summary>
        ''' 10入金区分取得
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
            Dim model_item As New Njc.Model.M_nkbn_Model(reccnt)

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
                            Case "nkbn_no"
                                .Nkbn_no(cntii) = fldvalue
                            Case "nkbn_name"
                                .Nkbn_name(cntii) = fldvalue
                            Case "nkbn_zokusei"
                                .Nkbn_zokusei_no(cntii) = fldvalue
                                .Nkbn_zokusei_name(cntii) = Me.Get_NkbnZokuseiname(fldvalue)
                        End Select

                    Next

                End With

            Next

            'モデルの引渡し
            RelItem_M_Nkbn = Nothing
            RelItem_M_Nkbn = model_item

            '入金区分属性をモデルへ格納
            Call Me.Set_NkbnZokusei()

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10入金区分抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry()

            Dim strsql As String

            strsql = ""
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 [nkbn_no] "
            strsql = strsql & " 	,[nkbn_name] "
            strsql = strsql & " 	,[nkbn_zokusei] "
            strsql = strsql & " FROM m_nkbn "
            strsql = strsql & " ORDER BY [nkbn_no] "

            Return strsql

        End Function

        ''' <summary>
        ''' 入金区分属性Noから入金区分属性名を取得
        ''' </summary>
        ''' <param name="zokuseino"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Get_NkbnZokuseiname(ByVal zokuseino As String) As String

            Dim rtn_string As String = ""

            Select Case zokuseino
                Case "1" : rtn_string = "現金"
                Case "2" : rtn_string = "振込"
                Case "3" : rtn_string = "振替"
                Case "4" : rtn_string = "その他"
                Case "99" : rtn_string = "移動"
            End Select

            Return rtn_string

        End Function

        ''' <summary>
        ''' 入金区分属性を取得してモデルへ格納
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Set_NkbnZokusei()

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim nkbnzno As String
            Dim nkbnzname As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String

            '値をまとめて格納
            nkbnzno = "1,2,3,4,99"
            nkbnzname = "現金,振込,振替,その他,移動"

            '配列に格納
            tmp_col1 = Split(nkbnzno, ",")
            tmp_col2 = Split(nkbnzname, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_nkbn_z_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii = 0 To UBound(tmp_col1)

                With model_item
                    .Nkbn_zokusei_no(cntii + 1) = tmp_col1(cntii)
                    .Nkbn_zokusei_name(cntii + 1) = tmp_col2(cntii)
                End With

            Next

            'モデルの引渡し
            RelItem_M_nkbn_z = Nothing
            RelItem_M_nkbn_z = model_item

        End Sub

        ''' <summary>
        ''' 入金グリッドの初期設定
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Initialize_Dgv(ByVal dgv As DataGridView)

            With dgv
                .Columns(0).ReadOnly = True
                .Columns(0).DefaultCellStyle.BackColor = Color.Azure
                .Columns(0).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(1).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(1).DefaultCellStyle.BackColor = Color.Beige
                .Columns(2).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(2).DefaultCellStyle.BackColor = Color.Beige
                .Columns(3).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(3).DefaultCellStyle.BackColor = Color.Beige
                .Columns(4).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(4).DefaultCellStyle.BackColor = Color.Beige
                .Columns(1).Visible = False    '20160825 10側の「No」表示制御処理を追加 -add
                .Columns(3).Visible = False    '20160825 10側の「No」表示制御処理を追加 -add
            End With

        End Sub

        ''' <summary>
        ''' 入金グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_nkbn_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_Nkbn_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_Nkbn_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_Nkbn_Rev7.Nkbn_name(cntii)
            Next

            '10紐付項目表示
            Dim relitemcnt10 As Integer = RelItem_M_Nkbn.ItemCnt - 1

            For cntii = 1 To dgv.RowCount

                '照合フラグ設定
                Dim matchflg As Boolean = False     '20160829 任意番号を自動で作成、表示する処理を追加 -add

                '対象文字列格納
                Dim taisyostr As String = dgv(0, cntii - 1).Value

                For cntjj = 1 To relitemcnt10

                    '照合用文字列格納
                    Dim searchstr As String = RelItem_M_Nkbn.Nkbn_name(cntjj)

                    If taisyostr = taisyostr.Replace(searchstr, "") Then
                        '20160701 指摘事項まとめファイルの対応 初期表示の修正 -del sta
                        'dgv(2, cntii - 1).Value = taisyostr
                        'dgv(1, cntii - 1).Style.BackColor = Color.Pink
                        'dgv(2, cntii - 1).Style.BackColor = Color.Pink
                        '20160701 指摘事項まとめファイルの対応 初期表示の修正 -del end
                    Else
                        dgv(1, cntii - 1).Value = RelItem_M_nkbn.Nkbn_no(cntjj)
                        dgv(2, cntii - 1).Value = RelItem_M_nkbn.Nkbn_name(cntjj)
                        dgv(3, cntii - 1).Value = RelItem_M_nkbn.Nkbn_zokusei_no(cntjj)
                        dgv(4, cntii - 1).Value = RelItem_M_nkbn.Nkbn_zokusei_name(cntjj)
                        dgv(1, cntii - 1).Style.BackColor = Color.Beige
                        dgv(2, cntii - 1).Style.BackColor = Color.Beige
                        '20160829 任意番号を自動で作成、表示する処理を追加 -add
                        matchflg = True
                        Exit For
                    End If

                Next

                '20160829 任意番号を自動で作成、表示する処理を追加 -add sta
                If matchflg = False Then
                    Dim maxno As Integer = Me.Get_Maxno(dgv)
                    dgv(1, cntii - 1).Value = maxno.ToString
                    dgv(2, cntii - 1).Value = taisyostr
                    dgv(3, cntii - 1).Value = "4"
                    dgv(4, cntii - 1).Value = "その他"
                End If
                '20160829 任意番号を自動で作成、表示する処理を追加 -add end

            Next

            '20160829 入金区分紐付不可行制御処理を追加 -add
            Call Me.Set_RowReadOnly(dgv)

        End Sub

        ''' <summary>
        ''' 入金グリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックス表示対象外の列の場合は処理を抜ける
            If Not (colcurrent = 2 Or colcurrent = 4) Then
                Exit Sub
            End If

            'コンボリスト作成
            cellcombo.Items.Add("")

            Select Case colcurrent
                Case 2
                    With RelItem_M_nkbn
                        For cntii = 1 To .itemcnt - 1
                            cellcombo.Items.Add(.Nkbn_name(cntii))
                        Next
                    End With
                Case 4
                    With RelItem_M_nkbn_z
                        For cntii = 1 To .itemcnt - 1
                            If cellcombo.Items.Contains(.Nkbn_zokusei_name(cntii)) = False Then
                                cellcombo.Items.Add(.Nkbn_zokusei_name(cntii))
                            End If
                        Next
                    End With
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
                    Case 1
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                    Case 2
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                    Case 3
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    Case 4
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                End Select
                Return rtn
            End If
            '2016.03.23 不要な箇所を削除していたため修正 -chg end

            '数値範囲チェック
            '2016.03.23 数値チェック処理実装 -chg sta
            'If colcurrent = 1 Then
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
                Case 1
                    max = 20
                    If EtcMethod.Chk_NumRange(currentvalue, min, max) = False And currentvalue <> 99 Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        '2016.03.23 削除箇所が不足していたため追加 -add sta
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                        '2016.03.23 削除箇所が不足していたため追加 -add end
                    End If
                Case 3
                    max = 4
                    If EtcMethod.Chk_NumRange(currentvalue, min, max) = False And currentvalue <> 99 Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        '2016.03.23 削除箇所が不足していたため追加 -add
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        rtn = False
                        Return rtn
                    End If
            End Select
            '2016.03.23 数値チェック処理実装 -chg end

            '一致する値を取得
            For cntii = 1 To RelItem_M_nkbn.itemcnt - 1

                Dim searchstr As String = ""
                Dim outstr As String = ""
                Dim chgcol As Integer = 0

                Select Case colcurrent
                    Case 1
                        searchstr = RelItem_M_nkbn.Nkbn_no(cntii)
                        outstr = RelItem_M_nkbn.Nkbn_name(cntii)
                        chgcol = colcurrent + 1
                    Case 2
                        searchstr = RelItem_M_nkbn.Nkbn_name(cntii)
                        outstr = RelItem_M_nkbn.Nkbn_no(cntii)
                        chgcol = colcurrent - 1
                    Case 3
                        searchstr = RelItem_M_nkbn_z.Nkbn_zokusei_no(cntii)
                        outstr = RelItem_M_nkbn_z.Nkbn_zokusei_name(cntii)
                        chgcol = colcurrent + 1
                    Case 4
                        searchstr = RelItem_M_nkbn_z.Nkbn_zokusei_name(cntii)
                        outstr = RelItem_M_nkbn_z.Nkbn_zokusei_no(cntii)
                        chgcol = colcurrent - 1
                End Select

                dgv(chgcol, rowcurrent).Value = ""

                If currentvalue = searchstr Then
                    dgv(chgcol, rowcurrent).Value = outstr
                    Exit For
                ElseIf colcurrent = 1 Then  '入金区分Noが手動で新規作成された場合    '20160525 全体的な動作の修正

                    Dim tmp_V7nkbnname As String = IIf(dgv(0, rowcurrent).Value Is Nothing, "", dgv(0, rowcurrent).Value)

                    If TypeOf dgv(2, rowcurrent) Is DataGridViewComboBoxCell Then
                        Dim cellcombo As New DataGridViewComboBoxCell
                        cellcombo = DirectCast(dgv(2, rowcurrent), DataGridViewComboBoxCell)

                        If cellcombo.Items.Contains(tmp_V7nkbnname) = False Then
                            cellcombo.Items.Add(tmp_V7nkbnname)
                        End If

                    End If

                    dgv(chgcol, rowcurrent).Value = tmp_V7nkbnname

                End If

            Next

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
            dgv = Njc.Frm.RelationFrm.dgvNkinKbn

            '親項目取得用
            Dim tmp_nkbnno As String = IIf(dgv(1, rowcurrent).Value Is Nothing, "", dgv(1, rowcurrent).Value)
            Dim tmp_nkbnname As String = IIf(dgv(2, rowcurrent).Value Is Nothing, "", dgv(2, rowcurrent).Value)

            'セルの読み取り制御処理
            If tmp_nkbnno.Trim & tmp_nkbnname.Trim = "" Then
                dgv(3, rowcurrent).ReadOnly = True
                dgv(4, rowcurrent).ReadOnly = True
            Else
                dgv(3, rowcurrent).ReadOnly = False
                dgv(4, rowcurrent).ReadOnly = False
            End If

            '20160829 入金区分紐付不可行制御処理を追加 -add
            Call Me.Set_RowReadOnly(dgv)

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
                Case 1
                    dgv(colcurrent + 2, rowcurrent).Value = ""
                    dgv(colcurrent + 3, rowcurrent).Value = ""
                Case 2
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
            dgv = Njc.Frm.RelationFrm.dgvNkinKbn
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiNkKbn
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllNkKbn

            '紐付全件数取得
            totalcnt = dgv.RowCount

            '紐付未設定件数取得
            For cntii = 0 To totalcnt - 1
                Dim tmp_no As String = dgv(1, cntii).Value
                Dim tmp_name As String = dgv(2, cntii).Value
                Dim tmp_zno As String = dgv(3, cntii).Value
                Dim tmp_zname As String = dgv(4, cntii).Value

                If tmp_no <> "" And tmp_name <> "" And tmp_zno <> "" And tmp_zname <> "" Then
                    relcnt = relcnt + 1
                End If
            Next

            '表示
            lbl_relcnt.Text = relcnt.ToString
            lbl_totalcnt.Text = totalcnt.ToString

            '色設定
            Njc.Frm.RelationFrm.lblCntNkKbnBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiNkKbn.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllNkKbn.ForeColor = Color.Black

            '表示件数による色設定の制御
            If relcnt = totalcnt Then
                lbl_relcnt.ForeColor = Color.Black
                lbl_totalcnt.ForeColor = Color.Black
            Else
                lbl_relcnt.ForeColor = Color.Pink
                lbl_totalcnt.ForeColor = Color.Pink
            End If

        End Sub

        ''' <summary>
        ''' 革命10で使用されているマスタのNoの最大値を取得 20160829 任意番号を自動で作成、表示する処理を追加
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Get_MstMaxno() As Integer

            Dim rtn As String = 0

            Dim mstcnt As Integer = RelItem_M_nkbn.Itemcnt
            Dim tmp_mstcntnew As Integer = 0
            Dim tmp_mstcntmax As Integer = 0

            '革命10マスタNoの最大値を取得
            For cntii = 1 To mstcnt - 1
                tmp_mstcntnew = Int32.Parse(RelItem_M_nkbn.Nkbn_no(cntii))
                If tmp_mstcntnew > tmp_mstcntmax Then
                    tmp_mstcntmax = tmp_mstcntnew
                End If
            Next

            '使用するために+1する
            rtn = tmp_mstcntmax + 1

            Return rtn

        End Function

        ''' <summary>
        ''' 画面で使用されているNoの最大値を取得 20160829 任意番号を自動で作成、表示する処理を追加
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Get_Maxno(ByVal dgv As DataGridView) As Integer

            Dim rtn As Integer = 0
            Dim tmp_cntnew As Integer = 0
            Dim tmp_cntmax As Integer = 0

            '画面上の最大Noを取得
            For cntii = 0 To dgv.RowCount - 1
                tmp_cntnew = Int32.Parse(IIf(dgv(1, cntii).Value = "", "0", dgv(1, cntii).Value))
                If tmp_cntnew > tmp_cntmax Then
                    tmp_cntmax = tmp_cntnew
                End If
            Next

            '使用するために+1する
            tmp_cntmax = tmp_cntmax + 1

            '革命10マスタの最大値を取得
            Dim mst_maxno As Integer = Me.Get_MstMaxno()

            'マスタの最大値と比較→大きい方を返却
            If mst_maxno <= tmp_cntmax Then
                rtn = tmp_cntmax
            Else
                rtn = mst_maxno
            End If

            Return rtn

        End Function

        ''' <summary>
        ''' 入金区分読取専用制御 20160829 入金区分紐付不可行制御処理を追加
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_RowReadOnly(ByVal dgv As DataGridView)

            Dim list_readonlykbn As New List(Of String) From {"現金", "振込", "振替"}

            With dgv
                For cntii = 0 To .RowCount - 1
                    Dim tmp_kbnname As String = dgv(0, cntii).Value
                    If list_readonlykbn.Contains(tmp_kbnname) Then
                        .Rows(cntii).ReadOnly = True
                    End If
                Next
            End With

        End Sub

    End Class

#End Region

End Namespace