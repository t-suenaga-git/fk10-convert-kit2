Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "契約分類マスタ取得"

    Public Class M_ky_rui_Repository

        ''' <summary>
        ''' 10契約分類取得
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
            Dim model_item As New Njc.Model.M_ky_rui_Model(reccnt)

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
                            Case "ky_ruino"
                                .Ky_ruino(cntii) = fldvalue
                            Case "ky_ruiname"
                                .Ky_ruiname(cntii) = fldvalue
                            Case Else

                        End Select
                    Next

                End With

            Next

            'モデルの引渡し
            RelItem_M_ky_rui = Nothing
            RelItem_M_ky_rui = model_item

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10契約分類抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry()

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " SELECT "
            tmp_sql = tmp_sql & " 	 ky_ruino "
            tmp_sql = tmp_sql & " 	,ky_ruiname "
            tmp_sql = tmp_sql & " FROM m_ky_rui "
            tmp_sql = tmp_sql & " ORDER BY ky_ruino "

            Return tmp_sql

        End Function

        ''' <summary>
        ''' 契約分類グリッドの初期設定
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Initialize_Dgv(ByVal dgv As DataGridView)

            Dim dgvchkcol As New DataGridViewCheckBoxColumn

            With dgv
                .Columns.Add(dgvchkcol)
                .Columns(0).ReadOnly = True
                .Columns(0).DefaultCellStyle.BackColor = Color.Azure
                .Columns(0).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(1).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(1).DefaultCellStyle.BackColor = Color.Beige
                .Columns(2).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(2).DefaultCellStyle.BackColor = Color.Beige
                .Columns(3).HeaderText = "定期借家扱い"
                .Columns(3).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(3).DefaultCellStyle.BackColor = Color.Beige
                .Columns(1).Visible = False    '20160825 10側の「No」表示制御処理を追加 -add
            End With

        End Sub

        ''' <summary>
        ''' 契約分類グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_Keirui Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_Keirui.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_Keirui.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_Keirui.Keirui_name(cntii)
            Next

            '10紐付項目表示
            Dim relitemcnt10 As Integer = RelItem_M_ky_rui.ItemCnt - 1

            For cntjj = 1 To RelItem_M_Keirui.ItemCnt - 1

                '照合フラグ設定
                Dim matchflg As Boolean = False     '20160829 任意番号を自動で作成、表示する処理を追加 -add

                '対象文字列格納
                Dim taisyostr As String = dgv(0, cntjj - 1).Value

                '20160701 指摘事項まとめファイルの対応 契約分類の初期値設定方法修正 -add sta
                Select Case taisyostr
                    Case "通常契約"
                        taisyostr = "普通賃貸借契約"
                    Case "定期借地借家権契約"
                        taisyostr = "定期借地借家契約"
                End Select
                '20160701 指摘事項まとめファイルの対応 契約分類の初期値設定方法修正 -add end

                For cntkk = 1 To relitemcnt10

                    '照合用文字列格納
                    Dim searchstr As String = RelItem_M_ky_rui.Ky_ruiname(cntkk)

                    If taisyostr = taisyostr.Replace(searchstr, "") Then
                        'dgv(2, cntjj - 1).Value = taisyostr
                        'dgv(1, cntjj - 1).Style.BackColor = Color.Pink
                        'dgv(2, cntjj - 1).Style.BackColor = Color.Pink
                    Else
                        dgv(1, cntjj - 1).Value = RelItem_M_ky_rui.Ky_ruino(cntkk)
                        dgv(2, cntjj - 1).Value = RelItem_M_ky_rui.Ky_ruiname(cntkk)
                        dgv(1, cntjj - 1).Style.BackColor = Color.Beige
                        dgv(2, cntjj - 1).Style.BackColor = Color.Beige
                        '20160829 任意番号を自動で作成、表示する処理を追加 -add
                        matchflg = True
                        Exit For
                    End If

                Next

                '20160701 指摘事項まとめファイルの対応 契約分類の初期値設定方法修正 -add sta
                If dgv(2, cntjj - 1).Value = "定期借地借家契約" Then
                    dgv(3, cntjj - 1).Value = True
                End If
                '20160701 指摘事項まとめファイルの対応 契約分類の初期値設定方法修正 -add end

                '20160829 任意番号を自動で作成、表示する処理を追加 -add sta
                If matchflg = False Then
                    Dim maxno As Integer = Me.Get_Maxno(dgv)
                    dgv(1, cntjj - 1).Value = maxno.ToString
                    dgv(2, cntjj - 1).Value = taisyostr
                End If
                '20160829 任意番号を自動で作成、表示する処理を追加 -add end

            Next

        End Sub

        ''' <summary>
        ''' 契約分類グリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックス表示対象外の列の場合は処理を抜ける
            If colcurrent <> 2 Then
                Exit Sub
            End If

            'コンボリスト作成
            cellcombo.Items.Add("")

            With RelItem_M_ky_rui
                For cntii = 1 To .itemcnt - 1
                    cellcombo.Items.Add(.Ky_ruiname(cntii))
                Next
            End With

            'カレントセルの値がコンボリストに含まれていない場合は新規に追加する (追加しないとエラーが生じるため)
            If currentvalue IsNot Nothing Then
                If cellcombo.Items.Contains(currentvalue) = False Then
                    cellcombo.Items.Add(currentvalue)
                End If
            End If

            '20160701 指摘事項まとめファイルの対応 契約分類の初期値設定方法修正 -add sta
            If cellcombo.Items.Contains("通常契約") = False Then
                cellcombo.Items.Add("通常契約")
            End If
            If cellcombo.Items.Contains("定期借地借家権契約") = False Then
                cellcombo.Items.Add("定期借地借家権契約")
            End If
            '20160701 指摘事項まとめファイルの対応 契約分類の初期値設定方法修正 -add end

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
            If currentvalue = "" Then
                Select Case colcurrent
                    Case 1
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    Case 2
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                End Select
                Return rtn
            End If

            '数値範囲チェック
            If colcurrent = 1 Then
                Dim min As Integer = 1
                Dim max As Integer = 9999
                If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
                    MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    dgv(colcurrent, rowcurrent).Value = ""
                    dgv(colcurrent + 1, rowcurrent).Value = ""
                    rtn = False
                    Return rtn
                End If
            End If

            '一致する値を取得
            For cntii = 1 To RelItem_M_ky_rui.itemcnt - 1

                Dim searchstr As String = ""
                Dim outstr As String = ""
                Dim chgcol As Integer = 0

                Select Case colcurrent
                    Case 1
                        searchstr = RelItem_M_ky_rui.Ky_ruino(cntii)
                        outstr = RelItem_M_ky_rui.Ky_ruiname(cntii)
                        chgcol = colcurrent + 1
                    Case 2
                        searchstr = RelItem_M_ky_rui.Ky_ruiname(cntii)
                        outstr = RelItem_M_ky_rui.Ky_ruino(cntii)
                        chgcol = colcurrent - 1
                    Case Else
                        Exit For
                End Select

                dgv(chgcol, rowcurrent).Value = ""

                If currentvalue = searchstr Then
                    dgv(chgcol, rowcurrent).Value = outstr
                    Exit For
                ElseIf colcurrent = 1 Then  '契約分類Noが手動で新規作成された場合    '20160525 全体的な動作の修正

                    Dim tmp_V7kyruiname As String = IIf(dgv(0, rowcurrent).Value Is Nothing, "", dgv(0, rowcurrent).Value)

                    If TypeOf dgv(2, rowcurrent) Is DataGridViewComboBoxCell Then
                        Dim cellcombo As New DataGridViewComboBoxCell
                        cellcombo = DirectCast(dgv(2, rowcurrent), DataGridViewComboBoxCell)

                        If cellcombo.Items.Contains(tmp_V7kyruiname) = False Then
                            cellcombo.Items.Add(tmp_V7kyruiname)
                        End If

                    End If

                    dgv(chgcol, rowcurrent).Value = tmp_V7kyruiname

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
            dgv = Njc.Frm.RelationFrm.dgvKyrui

            '親項目取得用
            Dim tmp_kyruino As String = IIf(dgv(1, rowcurrent).Value Is Nothing, "", dgv(1, rowcurrent).Value)
            Dim tmp_kyruiname As String = IIf(dgv(2, rowcurrent).Value Is Nothing, "", dgv(2, rowcurrent).Value)

            Dim aaa As Object

            'セルの読み取り制御処理
            If tmp_kyruino.Trim & tmp_kyruiname.Trim = "" Then
                dgv(3, rowcurrent).ReadOnly = True
                aaa = dgv(3, rowcurrent).Value
                'dgv(3, rowcurrent).Value = False
            Else
                dgv(3, rowcurrent).ReadOnly = False
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
                Case 1
                    dgv(colcurrent + 2, rowcurrent).Value = False
                Case 2
                    dgv(colcurrent + 1, rowcurrent).Value = False
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
            dgv = Njc.Frm.RelationFrm.dgvKyrui
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiKyBunrui
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllKyBunrui

            '紐付全件数取得
            totalcnt = dgv.RowCount

            '紐付未設定件数取得
            For cntii = 0 To totalcnt - 1

                Dim tmp_no As String = dgv(1, cntii).Value
                Dim tmp_name As String = dgv(2, cntii).Value

                If tmp_no <> "" And tmp_name <> "" Then
                    relcnt = relcnt + 1
                End If
            Next

            '表示
            lbl_relcnt.Text = relcnt.ToString
            lbl_totalcnt.Text = totalcnt.ToString

            '色設定
            Njc.Frm.RelationFrm.lblCntKyBunruiBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiKyBunrui.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllKyBunrui.ForeColor = Color.Black

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

            Dim mstcnt As Integer = RelItem_M_ky_rui.Itemcnt
            Dim tmp_mstcntnew As Integer = 0
            Dim tmp_mstcntmax As Integer = 0

            '革命10マスタNoの最大値を取得
            For cntii = 1 To mstcnt - 1
                tmp_mstcntnew = Int32.Parse(RelItem_M_ky_rui.Ky_ruino(cntii))
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

    End Class

#End Region

End Namespace