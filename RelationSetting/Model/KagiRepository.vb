Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query

Namespace Njc.Repository

#Region "鍵タイトルマスタ取得"

    Public Class M_kagi_Repository

        ''' <summary>
        ''' 10鍵取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadDB(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

        End Function

        ''' <summary>
        ''' 10鍵抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry()

        End Function

        ''' <summary>
        ''' 鍵グリッドの初期設定
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Initialize_Dgv(ByVal dgv As DataGridView)

            Dim dgvchkcol As New DataGridViewCheckBoxColumn

            With dgv
                'kakaka6 chg s -------------------------------------------------- 0420
                '.Height = 340
                '.Top = ((Njc.Frm.RelationFrm.tabPage10.Height - .Height) / 2) + 30
                'kakaka6 chg e --------------------------------------------------
                .Columns.Add(dgvchkcol)
                .Columns(0).ReadOnly = True
                .Columns(0).DefaultCellStyle.BackColor = Color.Azure
                .Columns(0).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(1).ReadOnly = True
                .Columns(1).DefaultCellStyle.BackColor = Color.Azure
                .Columns(1).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(2).HeaderText = "共用チェック"
                .Columns(2).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(2).DefaultCellStyle.BackColor = Color.Beige
            End With

        End Sub

        ''' <summary>
        ''' 鍵グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_kagi_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_kagi_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_kagi_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_kagi_Rev7.Kagi_no(cntii)
                dgv(1, cntii - 1).Value = RelItem_M_kagi_Rev7.Kagi_name(cntii)
            Next

        End Sub

        ''' <summary>
        ''' 鍵グリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

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
            dgv = Njc.Frm.RelationFrm.dgvKagi
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiKyoyoKagi
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllKyoyoKagi

            '紐付全件数取得
            totalcnt = dgv.RowCount

            ''紐付未設定件数取得
            'For cntii = 0 To totalcnt - 1
            '    Dim tmp_no As String = dgv(1, cntii).Value
            '    Dim tmp_name As String = dgv(2, cntii).Value

            '    If tmp_no <> "" And tmp_name <> "" Then
            '        relcnt = relcnt + 1
            '    End If
            'Next

            ''表示
            'lbl_relcnt.Text = relcnt.ToString
            'lbl_totalcnt.Text = totalcnt.ToString

            ''表示件数による色設定の制御
            'If relcnt = totalcnt Then
            '    lbl_relcnt.ForeColor = Color.Black
            '    lbl_totalcnt.ForeColor = Color.Black
            'Else
            '    lbl_relcnt.ForeColor = Color.Pink
            '    lbl_totalcnt.ForeColor = Color.Pink
            'End If


            '色設定
            Njc.Frm.RelationFrm.lblCntKyoyoKagiBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiKyoyoKagi.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllKyoyoKagi.ForeColor = Color.Black

            '表示
            lbl_relcnt.Text = " - "
            lbl_totalcnt.Text = totalcnt.ToString


        End Sub

    End Class

#End Region

End Namespace