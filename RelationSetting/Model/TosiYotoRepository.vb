Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query

Namespace Njc.Repository

#Region "都市計画・用途地域マスタ取得"

    Public Class M_TosiYoto_Repository

        ''' <summary>
        ''' 10都市計画取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadRelItem() As Boolean

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim Tosino As String
            Dim Tosiname As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String

            '値をまとめて格納
            Tosino = "1,2,3,4,5,6,7"
            Tosiname = "市街化区域,市街化調整区域(開発許可等による分譲地内),市街化調整区域(開発許可等なし),市街化調整区域(開発許可等有り)," & _
                       "非線引区域,準都市計画区域内,都市計画区域外"

            '配列に格納
            tmp_col1 = Split(Tosino, ",")
            tmp_col2 = Split(Tosiname, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_Tosi_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii As Integer = 1 To reccnt

                With model_item
                    .M_Tosino(cntii) = tmp_col1(cntii - 1)
                    .M_Tosiname(cntii) = tmp_col2(cntii - 1)
                End With

            Next

            'モデルの引渡し
            RelItem_M_Tosi = Nothing
            RelItem_M_Tosi = model_item

            '用途地域の取得
            Call Me.ReadRelItem_Sub()

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10用途地域取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadRelItem_Sub() As Boolean

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim Yotono As String
            Dim Yotoname As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String

            '値をまとめて格納
            Yotono = "1,2,3,4,5,6,7,8,9,10,11,12,13,14"
            Yotoname = "第一種低層住居専用地域,第二種低層住居専用地域,第一種中高層住居専用地域,第二種中高層住居専用地域,第一種住居地域," & _
                       "第二種住居地域,準住居地域,近隣商業地域,商業地域,準工業地域,工業地域,工業専用地域,指定なし,その他"

            '配列に格納
            tmp_col1 = Split(Yotono, ",")
            tmp_col2 = Split(Yotoname, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_Yoto_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii As Integer = 1 To reccnt

                With model_item
                    .M_Yotono(cntii) = tmp_col1(cntii - 1)
                    .M_Yotoname(cntii) = tmp_col2(cntii - 1)
                End With

            Next

            'モデルの引渡し
            RelItem_M_Yoto = Nothing
            RelItem_M_Yoto = model_item

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 都市計画・用途地域グリッドの初期設定
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
            End With

        End Sub

        ''' <summary>
        ''' 都市計画・用途地域グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            'V7にデータが無い場合は処理を抜ける
            If RelItem_M_TosiYoto_Rev7 Is Nothing Then
                Exit Sub
            End If

            'グリッド行数設定
            dgv.RowCount = RelItem_M_TosiYoto_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_TosiYoto_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_TosiYoto_Rev7.TosiYoto_name(cntii)
            Next

            '10紐付項目表示
            For cntii = 1 To dgv.RowCount

                '照合フラグ設定
                Dim matchflg_tosi As Boolean = False     '20160829 任意番号を自動で作成、表示する処理を追加 -add
                Dim matchflg_yoto As Boolean = False     '20160829 任意番号を自動で作成、表示する処理を追加 -add

                '対象文字列格納
                Dim taisyostr As String = dgv(0, cntii - 1).Value
                '20160829 任意番号を自動で作成、表示する処理を追加 -del sta
                ''照合フラグ
                'Dim match_flg As Boolean = False
                '20160829 任意番号を自動で作成、表示する処理を追加 -del end
                '都市計画を照合
                For cntjj = 1 To RelItem_M_Tosi.ItemCnt - 1

                    '照合用文字列格納
                    Dim searchstr As String = RelItem_M_Tosi.M_Tosiname(cntjj)

                    If taisyostr = taisyostr.Replace(searchstr, "") Then

                    Else
                        dgv(1, cntii - 1).Value = "都市計画"
                        dgv(2, cntii - 1).Value = RelItem_M_Tosi.M_Tosino(cntjj)
                        dgv(3, cntii - 1).Value = RelItem_M_Tosi.M_Tosiname(cntjj)
                        '20160829 任意番号を自動で作成、表示する処理を追加 -chg sta
                        'match_flg = True
                        matchflg_tosi = True
                        '20160829 任意番号を自動で作成、表示する処理を追加 -chg end
                        Exit For
                    End If

                Next

                '都市計画に一致するものが無かった場合は用途地域と照合する
                If matchflg_tosi = False Then       '20160829 任意番号を自動で作成、表示する処理を追加 match_flg → matchflg_tosi -chg

                    For cntjj = 1 To RelItem_M_Yoto.ItemCnt - 1

                        '照合用文字列格納
                        Dim searchstr As String = RelItem_M_Yoto.M_Yotoname(cntjj)

                        If taisyostr = taisyostr.Replace(searchstr, "") Then

                        Else
                            dgv(1, cntii - 1).Value = "用途地域"
                            dgv(2, cntii - 1).Value = RelItem_M_Yoto.M_Yotono(cntjj)
                            dgv(3, cntii - 1).Value = RelItem_M_Yoto.M_Yotoname(cntjj)
                            '20160829 任意番号を自動で作成、表示する処理を追加 -add
                            matchflg_yoto = True
                            Exit For
                        End If

                    Next

                End If

                '20160829 任意番号を自動で作成、表示する処理を追加 -add sta
                If matchflg_tosi = False And matchflg_yoto = False Then
                    dgv(1, cntii - 1).Value = "用途地域"
                    dgv(2, cntii - 1).Value = "14"
                    dgv(3, cntii - 1).Value = "その他"
                End If
                '20160829 任意番号を自動で作成、表示する処理を追加 -add end

            Next

        End Sub

        ''' <summary>
        ''' 都市計画・用途地域グリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal rowcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックス表示対象外の列の場合は処理を抜ける
            If Not (colcurrent = 1 Or colcurrent = 3) Then
                Exit Sub
            End If

            'コンボリスト作成
            cellcombo.Items.Add("")

            '選択列による処理の分岐
            Select Case colcurrent
                Case 1
                    cellcombo.Items.Add("都市計画")
                    cellcombo.Items.Add("用途地域")
                Case 3

                    '都市計画/用途地域からコンボボックスへセットする値を設定する
                    'グリッド取得
                    Dim dgv As New DataGridView
                    dgv = Njc.Frm.RelationFrm.dgvTosiYoto

                    Dim tmp_tosiyotokbn As String = IIf(dgv(1, rowcurrent).Value Is Nothing, "", dgv(1, rowcurrent).Value)

                    Select Case tmp_tosiyotokbn
                        Case "都市計画"
                            With RelItem_M_Tosi
                                For cntii = 1 To .itemcnt - 1
                                    cellcombo.Items.Add(.M_Tosiname(cntii))
                                Next
                            End With
                        Case "用途地域"
                            With RelItem_M_Yoto
                                For cntii = 1 To .itemcnt - 1
                                    cellcombo.Items.Add(.M_Yotoname(cntii))
                                Next
                            End With
                    End Select

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
            If currentvalue = "" Then
                Select Case colcurrent
                    Case 1
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                    Case 2
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    Case 3
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                End Select
                Return rtn
            End If

            '都市計画/用途地域区分を取得
            Dim tmp_tosiyotokbn As String = IIf(dgv(1, rowcurrent).Value Is Nothing, "", dgv(1, rowcurrent).Value)
            If tmp_tosiyotokbn = "" Then
                Return rtn
            End If

            '数値範囲チェック
            Dim min As Integer = 1
            Dim max As Integer = 0

            If colcurrent = 2 Then
                Select Case tmp_tosiyotokbn
                    Case "都市計画"
                        max = 7
                    Case "用途地域"
                        max = 14
                End Select

                If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
                    MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    dgv(colcurrent, rowcurrent).Value = ""
                    dgv(colcurrent + 1, rowcurrent).Value = ""
                    rtn = False
                    Return rtn
                End If
            End If

            '一致する値を取得
            If colcurrent = 2 Or colcurrent = 3 Then

                Select Case tmp_tosiyotokbn
                    Case "都市計画"

                        For cntii = 1 To RelItem_M_Tosi.itemcnt - 1

                            Dim searchstr As String = ""
                            Dim outstr As String = ""
                            Dim chgcol As Integer = 0

                            Select Case colcurrent
                                Case 2
                                    searchstr = RelItem_M_Tosi.M_Tosino(cntii)
                                    outstr = RelItem_M_Tosi.M_Tosiname(cntii)
                                    chgcol = colcurrent + 1
                                Case 3
                                    searchstr = RelItem_M_Tosi.M_Tosiname(cntii)
                                    outstr = RelItem_M_Tosi.M_Tosino(cntii)
                                    chgcol = colcurrent - 1
                            End Select

                            dgv(chgcol, rowcurrent).Value = ""

                            If currentvalue = searchstr Then
                                dgv(chgcol, rowcurrent).Value = outstr
                                Exit For
                            End If

                        Next

                    Case "用途地域"

                        For cntii = 1 To RelItem_M_Yoto.itemcnt - 1

                            Dim searchstr As String = ""
                            Dim outstr As String = ""
                            Dim chgcol As Integer = 0

                            Select Case colcurrent
                                Case 2
                                    searchstr = RelItem_M_Yoto.M_Yotono(cntii)
                                    outstr = RelItem_M_Yoto.M_Yotoname(cntii)
                                    chgcol = colcurrent + 1
                                Case 3
                                    searchstr = RelItem_M_Yoto.M_Yotoname(cntii)
                                    outstr = RelItem_M_Yoto.M_Yotono(cntii)
                                    chgcol = colcurrent - 1
                            End Select

                            dgv(chgcol, rowcurrent).Value = ""

                            If currentvalue = searchstr Then
                                dgv(chgcol, rowcurrent).Value = outstr
                                Exit For
                            End If

                        Next

                End Select

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
            dgv = Njc.Frm.RelationFrm.dgvTosiYoto

            '親項目取得用
            Dim tmp_tosiyotokbn As String = IIf(dgv(1, rowcurrent).Value Is Nothing, "", dgv(1, rowcurrent).Value)

            'セルの読み取り制御処理
            If tmp_tosiyotokbn.Trim = "" Then
                dgv(2, rowcurrent).ReadOnly = True
                dgv(3, rowcurrent).ReadOnly = True
            Else
                dgv(2, rowcurrent).ReadOnly = False
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
            dgv = Njc.Frm.RelationFrm.dgvTosiYoto
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiTosiYoto
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllTosiYoto

            '紐付全件数取得
            totalcnt = dgv.RowCount

            '紐付未設定件数取得
            For cntii = 0 To totalcnt - 1

                Dim tmp_kbn As String = dgv(1, cntii).Value
                Dim tmp_no As String = dgv(1, cntii).Value
                Dim tmp_name As String = dgv(2, cntii).Value

                If tmp_kbn <> "" And tmp_no <> "" And tmp_name <> "" Then
                    relcnt = relcnt + 1
                End If
            Next

            '表示
            lbl_relcnt.Text = relcnt.ToString
            lbl_totalcnt.Text = totalcnt.ToString

            '色設定
            Njc.Frm.RelationFrm.lblCntTosiYotoBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiTosiYoto.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllTosiYoto.ForeColor = Color.Black

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