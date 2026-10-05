Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query

Namespace Njc.Repository

#Region "構造マスタ取得"

    Public Class M_kozo_Repository

        ''' <summary>
        ''' 10構造取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadRelItem() As Boolean

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim kozono As String
            Dim kozoname As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String

            '値をまとめて格納
            kozono = "1,2,3,4,5,6,7,8,9,10,11,99"
            kozoname = "木造,鉄骨造,軽量鉄骨造,鉄筋コンクリート造,鉄骨鉄筋コンクリート造,プレキャストコンクリート造," & _
                       "鉄骨プレキャストコンクリート造,軽量気泡コンクリート造,コンクリートブロック造,鉄筋ブロック造,コンクリート充填鋼管造,その他"

            '配列に格納
            tmp_col1 = Split(kozono, ",")
            tmp_col2 = Split(kozoname, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_kozo_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii As Integer = 1 To reccnt

                With model_item
                    .M_kozono(cntii) = tmp_col1(cntii - 1)
                    '20160701 指摘事項まとめファイルの対応 構造名称へ略称付加 -chg sta
                    '.M_kozoname(cntii) = tmp_col2(cntii - 1)
                    Dim tmp_kozoname As String = Me.Set_AddKozoKigo(tmp_col2(cntii - 1))
                    .M_kozoname(cntii) = tmp_kozoname
                    '20160701 指摘事項まとめファイルの対応 構造名称へ略称付加 -chg end
                End With

            Next

            'モデルの引渡し
            RelItem_M_kozo = Nothing
            RelItem_M_kozo = model_item

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 構造グリッドの初期設定
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
        ''' 構造グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_kozo_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_kozo_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_kozo_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_kozo_Rev7.Kozo_name(cntii)
            Next

            '10紐付項目表示
            Dim relitemcnt10 As Integer = RelItem_M_kozo.ItemCnt - 1

            For cntii = 1 To dgv.RowCount

                '照合フラグ設定
                Dim matchflg As Boolean = False     '20160829 任意番号を自動で作成、表示する処理を追加 -add

                '対象文字列格納
                Dim taisyostr As String = dgv(0, cntii - 1).Value

                For cntjj = 1 To relitemcnt10

                    '照合用文字列格納
                    '20160701 指摘事項まとめファイルの対応 構造名称へ略称付加 -chg sta
                    'Dim searchstr As String = RelItem_M_kozo.M_kozoname(cntjj)
                    Dim searchstr As String = Me.Set_DeleteKozoKigo(RelItem_M_kozo.M_kozoname(cntjj))
                    '20160701 指摘事項まとめファイルの対応 構造名称へ略称付加 -chg end

                    If taisyostr = taisyostr.Replace(searchstr, "") Then
                        '20160701 指摘事項まとめファイルの対応 初期表示の修正 -del sta
                        'dgv(2, cntii - 1).Value = taisyostr
                        'dgv(1, cntii - 1).Style.BackColor = Color.Pink
                        'dgv(2, cntii - 1).Style.BackColor = Color.Pink
                        '20160701 指摘事項まとめファイルの対応 初期表示の修正 -del end
                    Else
                        dgv(1, cntii - 1).Value = RelItem_M_kozo.M_kozono(cntjj)
                        dgv(2, cntii - 1).Value = RelItem_M_kozo.M_kozoname(cntjj)
                        dgv(1, cntii - 1).Style.BackColor = Color.Beige
                        dgv(2, cntii - 1).Style.BackColor = Color.Beige
                        '20160829 任意番号を自動で作成、表示する処理を追加 -add
                        matchflg = True
                        Exit For
                    End If

                Next

                '20160829 構造の名称をアルファベットでも照合する処理を追加 -add sta
                If matchflg = False Then
                    'アルファベットでの照合
                    Dim tmp_kozono As String = ""
                    Dim tmp_kozoname As String = ""
                    Call Me.Match_KozoAlpha(taisyostr, tmp_kozono, tmp_kozoname)
                    dgv(1, cntii - 1).Value = tmp_kozono
                    dgv(2, cntii - 1).Value = tmp_kozoname
                End If
                '20160829 構造の名称をアルファベットでも照合する処理を追加 -add end

            Next

        End Sub

        ''' <summary>
        ''' 構造グリッドのコンボボックスリスト作成
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

            With RelItem_M_kozo
                For cntii = 1 To .itemcnt - 1
                    cellcombo.Items.Add(.M_kozoname(cntii))
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
                Dim max As Integer = 11   '2016.03.23 9999 → 11 へ変更 
                If EtcMethod.Chk_NumRange(currentvalue, min, max) = False And currentvalue <> 99 Then      '2016.03.23 条件に「99」を追加
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
            For cntii = 1 To RelItem_M_kozo.itemcnt - 1

                Dim searchstr As String = ""
                Dim outstr As String = ""
                Dim chgcol As Integer = 0

                Select Case colcurrent
                    Case 1
                        searchstr = RelItem_M_kozo.M_kozono(cntii)
                        outstr = RelItem_M_kozo.M_kozoname(cntii)
                        chgcol = colcurrent + 1
                    Case 2
                        searchstr = RelItem_M_kozo.M_kozoname(cntii)
                        outstr = RelItem_M_kozo.M_kozono(cntii)
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
        ''' 革命10構造名称に記号を付加する '20160701 指摘事項まとめファイルの対応 初期表示の修正
        ''' </summary>
        ''' <param name="kozoname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_AddKozoKigo(ByVal kozoname As String) As String

            Dim rtn_str As String = ""

            Select Case kozoname

                Case "木造"
                    rtn_str = kozoname & "(W)"
                Case "鉄骨造"
                    rtn_str = kozoname & "(S)"
                Case "軽量鉄骨造"
                    rtn_str = kozoname & "(S)"
                Case "鉄筋コンクリート造"
                    rtn_str = kozoname & "(RC)"
                Case "鉄骨鉄筋コンクリート造"
                    rtn_str = kozoname & "(SRC)"
                Case "プレキャストコンクリート造"
                    rtn_str = kozoname & "(PC)"
                Case "鉄骨プレキャストコンクリート造"
                    rtn_str = kozoname & "(HPC)"
                Case "軽量気泡コンクリート造"
                    rtn_str = kozoname & "(ALC)"
                Case "コンクリートブロック造"
                    rtn_str = kozoname & "(CB)"
                Case "鉄筋ブロック造"
                    rtn_str = kozoname & ""
                Case "コンクリート充填鋼管造"
                    rtn_str = kozoname & "(CFT)"
                Case "その他"
                    rtn_str = kozoname & ""
            End Select

            Return rtn_str

        End Function

        ''' <summary>
        ''' 革命10構造名称に付加した記号を除去する '20160701 指摘事項まとめファイルの対応 初期表示の修正
        ''' </summary>
        ''' <param name="kozoname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_DeleteKozoKigo(ByVal kozoname As String) As String

            Dim rtn_str As String = ""
            Dim tmp_str() As String = kozoname.Split("(")
            rtn_str = tmp_str(0)
            Return rtn_str

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
            dgv = Njc.Frm.RelationFrm.dgvKozo
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiKozo
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllKozo

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
            Njc.Frm.RelationFrm.lblCntKozoBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiKozo.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllKozo.ForeColor = Color.Black

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
        ''' 構造の名称をアルファベットで照合 '20160829 構造の名称をアルファベットでも照合する処理を追加 -add
        ''' </summary>
        ''' <param name="taisyostr"></param>
        ''' <param name="kozono"></param>
        ''' <param name="kozoname"></param>
        ''' <remarks></remarks>
        Public Sub Match_KozoAlpha(ByVal taisyostr As String, ByRef kozono As String, ByRef kozoname As String)

            Dim rtn As Boolean = True
            Dim tmp_str As String = StrConv(taisyostr, VbStrConv.Narrow Or VbStrConv.Uppercase)

            Select Case True
                Case taisyostr.Replace("SRC", "") <> taisyostr
                    kozono = "5" : kozoname = Me.Set_AddKozoKigo("鉄骨鉄筋コンクリート造")
                Case taisyostr.Replace("RC", "") <> taisyostr
                    kozono = "4" : kozoname = Me.Set_AddKozoKigo("鉄筋コンクリート造")
                Case taisyostr.Replace("S", "") <> taisyostr
                    kozono = "2" : kozoname = Me.Set_AddKozoKigo("鉄骨造")
                Case taisyostr.Replace("W", "") <> taisyostr
                    kozono = "1" : kozoname = Me.Set_AddKozoKigo("木造")
                Case taisyostr.Replace("HPC", "") <> taisyostr
                    kozono = "7" : kozoname = Me.Set_AddKozoKigo("鉄骨プレキャストコンクリート造")
                Case taisyostr.Replace("PC", "") <> taisyostr
                    kozono = "6" : kozoname = Me.Set_AddKozoKigo("プレキャストコンクリート造")
                Case taisyostr.Replace("ALC", "") <> taisyostr
                    kozono = "8" : kozoname = Me.Set_AddKozoKigo("軽量気泡コンクリート造")
                Case taisyostr.Replace("CB", "") <> taisyostr
                    kozono = "9" : kozoname = Me.Set_AddKozoKigo("コンクリートブロック造")
                Case taisyostr.Replace("CFT", "") <> taisyostr
                    kozono = "11" : kozoname = Me.Set_AddKozoKigo("コンクリート充填鋼管造")
                Case Else
                    kozono = "99" : kozoname = Me.Set_AddKozoKigo("その他")
            End Select

        End Sub

    End Class

#End Region

End Namespace