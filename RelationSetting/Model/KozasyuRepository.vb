Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query

Namespace Njc.Repository

#Region "口座種別マスタ取得"

    Public Class M_kozasyu_Repository

        ''' <summary>
        ''' 10口座種別取得
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
            Dim model_item As New Njc.Model.M_kozasyu_Model(reccnt)

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
                            Case "kosyu_no"
                                .Kosyu_no(cntii) = fldvalue
                            Case "kosyu_name"
                                .Kosyu_name(cntii) = fldvalue
                            Case Else

                        End Select
                    Next

                End With

            Next

            'モデルの引渡し
            RelItem_M_kozasyu = Nothing
            RelItem_M_kozasyu = model_item

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10口座種別抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry()

            Dim strsql As String

            strsql = ""
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 [kosyu_no] "
            strsql = strsql & " 	,[kosyu_name] "
            strsql = strsql & " FROM m_kozasyu "
            strsql = strsql & " ORDER BY [kosyu_no] "

            Return strsql

        End Function

        ''' <summary>
        ''' 口座種別グリッドの初期設定
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
        ''' 口座種別グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_kozasyu_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_kozasyu_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_kozasyu_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_kozasyu_Rev7.Kosyu_name(cntii)
            Next

            '10紐付項目表示
            Dim relitemcnt10 As Integer = RelItem_M_kozasyu.ItemCnt - 1

            For cntjj = 1 To RelItem_M_kozasyu_Rev7.ItemCnt - 1

                '照合フラグ設定
                Dim matchflg As Boolean = False     '20160829 任意番号を自動で作成、表示する処理を追加 -add

                '対象文字列格納
                Dim taisyostr As String = dgv(0, cntjj - 1).Value

                For cntkk = 1 To relitemcnt10

                    '照合用文字列格納
                    Dim searchstr As String = RelItem_M_kozasyu.Kosyu_name(cntkk)

                    If taisyostr = taisyostr.Replace(searchstr, "") Then
                        '20160701 指摘事項まとめファイルの対応 初期表示の修正 -del sta
                        'dgv(2, cntjj - 1).Value = taisyostr
                        'dgv(1, cntjj - 1).Style.BackColor = Color.Pink
                        'dgv(2, cntjj - 1).Style.BackColor = Color.Pink
                        '20160701 指摘事項まとめファイルの対応 初期表示の修正 -del end
                    Else
                        dgv(1, cntjj - 1).Value = RelItem_M_kozasyu.Kosyu_no(cntkk)
                        dgv(2, cntjj - 1).Value = RelItem_M_kozasyu.Kosyu_name(cntkk)
                        dgv(1, cntjj - 1).Style.BackColor = Color.Beige
                        dgv(2, cntjj - 1).Style.BackColor = Color.Beige
                        '20160829 任意番号を自動で作成、表示する処理を追加 -add
                        matchflg = True
                        Exit For
                    End If

                Next

                '20160829 任意番号を自動で作成、表示する処理を追加 -add sta
                If matchflg = False Then
                    dgv(1, cntjj - 1).Value = "9"
                    dgv(2, cntjj - 1).Value = "その他"
                End If
                '20160829 任意番号を自動で作成、表示する処理を追加 -add end

            Next

        End Sub

        ''' <summary>
        ''' 口座種別グリッドのコンボボックスリスト作成
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

            With RelItem_M_kozasyu
                For cntii = 1 To .itemcnt - 1
                    cellcombo.Items.Add(.Kosyu_name(cntii))
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
            '2016.03.23 最大最小値チェック処理追加 -chg sta
            '口座種別の場合歯抜けの番号が使用されているため範囲ではなく直接数値でチェックする
            'If colcurrent = 1 Then
            '    Dim min As Integer = 1
            '    Dim max As Integer = 9999
            '    If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
            '        rtn = False
            '        Return rtn
            '    End If
            'End If
            If colcurrent = 1 Then
                If Not (currentvalue = 1 Or currentvalue = 2 Or currentvalue = 4 Or currentvalue = 9) Then
                    MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    dgv(colcurrent, rowcurrent).Value = ""
                    dgv(colcurrent + 1, rowcurrent).Value = ""
                    rtn = False
                    Return rtn
                End If
            End If
            '2016.03.23 最大最小値チェック処理追加 -chg end

            '一致する値を取得
            For cntii = 1 To RelItem_M_kozasyu.itemcnt - 1

                Dim searchstr As String = ""
                Dim outstr As String = ""
                Dim chgcol As Integer = 0

                Select Case colcurrent
                    Case 1
                        searchstr = RelItem_M_kozasyu.Kosyu_no(cntii)
                        outstr = RelItem_M_kozasyu.Kosyu_name(cntii)
                        chgcol = colcurrent + 1
                    Case 2
                        searchstr = RelItem_M_kozasyu.Kosyu_name(cntii)
                        outstr = RelItem_M_kozasyu.Kosyu_no(cntii)
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
            dgv = Njc.Frm.RelationFrm.dgvKozasyu
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiKozaSyubetu
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllKozaSyubetu

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
            Njc.Frm.RelationFrm.lblCntKozaSyubetuBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiKozaSyubetu.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllKozaSyubetu.ForeColor = Color.Black

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