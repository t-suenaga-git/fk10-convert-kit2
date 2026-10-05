Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query

Namespace Njc.Repository

#Region "取引態様マスタ取得"

    Public Class M_toritaiyo_Repository

        ''' <summary>
        ''' 10取引態様取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadRelItem() As Boolean

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim toritaiyono As String
            Dim toritaiyoname As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String

            '値をまとめて格納
            toritaiyono = "1,2,3,4,5,6"
            toritaiyoname = "仲介先物,一般媒介,専任媒介,専属専任媒介,代理,貸主"

            '配列に格納
            tmp_col1 = Split(toritaiyono, ",")
            tmp_col2 = Split(toritaiyoname, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_toritaiyo_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii As Integer = 1 To reccnt

                With model_item
                    .Toritaiyono(cntii) = tmp_col1(cntii - 1)
                    .Toritaiyoname(cntii) = tmp_col2(cntii - 1)
                End With

            Next

            'モデルの引渡し
            RelItem_M_toritaiyo = Nothing
            RelItem_M_toritaiyo = model_item

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 取引態様グリッドの初期設定
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
                .Columns(1).Visible = False    '20160825 10側の「No」表示制御処理を追加 -add
            End With

        End Sub

        ''' <summary>
        ''' 取引態様グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_toritaiyo_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_toritaiyo_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_toritaiyo_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_toritaiyo_Rev7.Taiyo_name(cntii)
            Next

            '10紐付項目表示
            Dim relitemcnt10 As Integer = RelItem_M_toritaiyo.ItemCnt - 1

            For cntii = 1 To dgv.RowCount

                '対象文字列格納
                Dim taisyostr As String = dgv(0, cntii - 1).Value

                '20160829 取引態様の自動紐付処理修正 -add sta
                Select Case taisyostr
                    Case "仲介", "媒介"
                        taisyostr = "一般媒介"
                End Select
                '20160829 取引態様の自動紐付処理修正 -add end

                For cntjj = 1 To relitemcnt10

                    '照合用文字列格納
                    Dim searchstr As String = RelItem_M_toritaiyo.Toritaiyoname(cntjj)

                    If taisyostr = taisyostr.Replace(searchstr, "") Then
                        '20160701 指摘事項まとめファイルの対応 初期表示の修正 -del sta
                        'dgv(2, cntii - 1).Value = taisyostr
                        'dgv(1, cntii - 1).Style.BackColor = Color.Pink
                        'dgv(2, cntii - 1).Style.BackColor = Color.Pink
                        '20160701 指摘事項まとめファイルの対応 初期表示の修正 -del end
                    Else
                        dgv(1, cntii - 1).Value = RelItem_M_toritaiyo.Toritaiyono(cntjj)
                        dgv(2, cntii - 1).Value = RelItem_M_toritaiyo.Toritaiyoname(cntjj)
                        dgv(1, cntii - 1).Style.BackColor = Color.Beige
                        dgv(2, cntii - 1).Style.BackColor = Color.Beige
                        Exit For
                    End If

                Next

            Next

        End Sub

        ''' <summary>
        ''' 取引態様グリッドのコンボボックスリスト作成
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

            With RelItem_M_toritaiyo
                For cntii = 1 To .itemcnt - 1
                    cellcombo.Items.Add(.Toritaiyoname(cntii))
                Next
            End With

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
                Dim max As Integer = 6   '2016.03.23 9999 → 6 へ変更
                If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
                    '2016.03.23 最大最小値チェック処理追加 -add sta
                    MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    dgv(colcurrent, rowcurrent).Value = ""
                    dgv(colcurrent + 1, rowcurrent).Value = ""
                    '2016.03.23 最大最小値チェック処理追加 -add end
                    rtn = False
                    Return rtn
                End If
            End If

            '一致する値を取得
            For cntii = 1 To RelItem_M_toritaiyo.itemcnt - 1

                Dim searchstr As String = ""
                Dim outstr As String = ""
                Dim chgcol As Integer = 0

                Select Case colcurrent
                    Case 1
                        searchstr = RelItem_M_toritaiyo.Toritaiyono(cntii)
                        outstr = RelItem_M_toritaiyo.Toritaiyoname(cntii)
                        chgcol = colcurrent + 1
                    Case 2
                        searchstr = RelItem_M_toritaiyo.Toritaiyoname(cntii)
                        outstr = RelItem_M_toritaiyo.Toritaiyono(cntii)
                        chgcol = colcurrent - 1
                End Select

                dgv(chgcol, rowcurrent).Value = ""

                If currentvalue = searchstr Then
                    dgv(chgcol, rowcurrent).Value = outstr
                    Exit For
                End If

            Next

            Return rtn

        End Function

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
            dgv = Njc.Frm.RelationFrm.dgvToritaiyo
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiTaiyo
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllTaiyo

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
            Njc.Frm.RelationFrm.lblCntTaiyoBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiTaiyo.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllTaiyo.ForeColor = Color.Black

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