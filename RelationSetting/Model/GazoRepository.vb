Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "画像割付取得"

    Public Class M_gazo_title_Repository

        ''' <summary>
        ''' 10画像タイトル取得(区分)
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadRelItem() As Boolean

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim gazokbnno As String
            Dim gazokbnname As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String

            '値をまとめて格納
            gazokbnno = "1,2,3"
            gazokbnname = "物件,部屋,周辺"

            '配列に格納
            tmp_col1 = Split(gazokbnno, ",")
            tmp_col2 = Split(gazokbnname, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_gazo_title_kbn_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii As Integer = 1 To reccnt

                With model_item
                    .Gazo_kbnno(cntii) = tmp_col1(cntii - 1)
                    .Gazo_kbnname(cntii) = tmp_col2(cntii - 1)
                End With

            Next

            'モデルの引渡し
            RelItem_M_gazo_title_kbn = Nothing
            RelItem_M_gazo_title_kbn = model_item

            '画像タイトルの取得
            Call Me.ReadRelItem_Sub()

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10画像タイトル取得(画像タイトル)
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ReadRelItem_Sub()

            Dim reccnt As Integer
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True
            Dim gazokbnno As String
            Dim gazokbnname As String
            Dim gazono As String
            Dim gazoname As String
            Dim tmp_col1() As String
            Dim tmp_col2() As String
            Dim tmp_col3() As String
            Dim tmp_col4() As String

            '値をまとめて格納
            gazokbnno = "3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3"
            gazokbnname = "周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺,周辺"
            gazono = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33"
            gazoname = "病院,スーパーマーケット,コンビニエンスストア,ショッピングセンター,ドラッグストア,ホームセンター,高校・高専,幼稚園・保育園,郵便局,役所,図書館,銀行,警察署・交番,公園,その他,大学・専門学校,飲食店,レジャー・観光,小学校,中学校,商店街,ディスカウントショップ,百均,デパート,予備校,養護学校,映画館,美術館・博物館,本屋,ビデオ・DVD,弁当屋,ファストフード,カフェ"

            '配列に格納
            tmp_col1 = Split(gazokbnno, ",")
            tmp_col2 = Split(gazokbnname, ",")
            tmp_col3 = Split(gazono, ",")
            tmp_col4 = Split(gazoname, ",")

            '配列数を取得
            reccnt = UBound(tmp_col1) + 1

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_gazo_title_name_Model(reccnt)

            '==================================================
            ' V10：データ取得・配列格納
            '==================================================
            For cntii As Integer = 1 To reccnt

                With model_item
                    .Gazo_kbnno(cntii) = tmp_col1(cntii - 1)
                    .Gazo_kbnname(cntii) = tmp_col2(cntii - 1)
                    .Gazo_no(cntii) = tmp_col3(cntii - 1)
                    .Gazo_name(cntii) = tmp_col4(cntii - 1)
                End With

            Next

            'モデルの引渡し
            RelItem_M_gazo_title_name = Nothing
            RelItem_M_gazo_title_name = model_item

        End Sub

        ''' <summary>
        ''' 画像タイトルグリッドの初期設定
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Initialize_Dgv(ByVal dgv As DataGridView)

            With dgv
                '20160725 スクロールバー表示修正 -del
                '.Height = 350 '20160602 画像紐付け修正 -add
                .Columns(0).ReadOnly = True
                .Columns(0).DefaultCellStyle.BackColor = Color.Azure
                .Columns(0).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(1).ReadOnly = True
                .Columns(1).DefaultCellStyle.BackColor = Color.Azure
                .Columns(1).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(2).ReadOnly = True
                .Columns(2).DefaultCellStyle.BackColor = Color.Azure
                .Columns(2).HeaderCell.Style.BackColor = Color.PowderBlue

                .Columns(3).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(3).DefaultCellStyle.BackColor = Color.Beige
                .Columns(4).ReadOnly = True
                .Columns(4).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(4).DefaultCellStyle.BackColor = Color.Beige
                .Columns(5).ReadOnly = True
                .Columns(5).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(5).DefaultCellStyle.BackColor = Color.Beige

            End With

        End Sub

        ''' <summary>
        ''' 画像タイトルグリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_gazo_title_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_gazo_title_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_gazo_title_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_gazo_title_Rev7.Syubetu(cntii)
                dgv(1, cntii - 1).Value = RelItem_M_gazo_title_Rev7.Kbn(cntii)
                dgv(2, cntii - 1).Value = RelItem_M_gazo_title_Rev7.Gazo_name(cntii)
            Next

            '20160531 画像紐付処理の修正 -add sta
            '10紐付項目の表示
            'V7の画像種別を元に移行先の画像種別を決定
            For cntjj = 1 To RelItem_M_gazo_title_Rev7.ItemCnt - 1

                Dim tmp_V7_syubetu As String = dgv(0, cntjj - 1).Value
                dgv(3, cntjj - 1).Value = tmp_V7_syubetu

                '20160610 画像紐付箇所の修正 10設定側の画像Noデフォルト表示設定の追加 -add sta
                'V7の画像Noをデフォルト値として予め10側にも設定しておく
                dgv(4, cntjj - 1).Value = dgv(1, cntjj - 1).Value
                '20160610 画像紐付箇所の修正 10設定側の画像Noデフォルト表示設定の追加 -add end

                '20160610 画像紐付箇所の修正 初期状態でコンボボックスを設定 -add sta
                Dim cellcombo As New DataGridViewComboBoxCell
                Call Me.Set_DgvCmbList(3, cntjj - 1, Nothing, cellcombo)

                If cellcombo.Items.Count <> 0 Then
                    dgv(3, cntjj - 1) = cellcombo
                    cellcombo.DisplayStyleForCurrentCellOnly = True
                    dgv(3, cntjj - 1).Value = "物件"
                End If
                '20160610 画像紐付箇所の修正 初期状態でコンボボックスを設定 -add end

                '20160701 指摘事項まとめファイルの対応 画像種別が物件時のセル着色修正 -add
                dgv(5, cntjj - 1).Style.BackColor = Color.LightGray

            Next
            '20160531 画像紐付処理の修正 -add end

            '20160610 部屋画像の割付の削除 -del sta
            ''20160602 画像紐付け修正 -add sta
            ''画像種別が部屋の場合、10設定側では変更不可のため灰色にする
            'For cntkk = 1 To dgv.RowCount
            '    Dim tmp_10_syubetu As String = dgv(3, cntkk - 1).Value
            '    If tmp_10_syubetu = "部屋" Then
            '        dgv(3, cntkk - 1).Style.BackColor = Color.Gray
            '        dgv(5, cntkk - 1).Style.BackColor = Color.Gray
            '    End If
            'Next
            ''20160602 画像紐付け修正 -add end
            '20160610 部屋画像の割付の削除 -del end

            '20160523 完全一致による表示は保留にしておく -del sta
            '10紐付項目表示
            'Dim relitemcnt10 As Integer = RelItem_M_gazo_title.ItemCnt - 1

            'For cntii = 1 To dgv.RowCount

            '    '対象文字列格納
            '    Dim taisyostr As String = dgv(0, cntii - 1).Value

            '    For cntjj = 1 To relitemcnt10

            '        '照合用文字列格納
            '        Dim searchstr As String = RelItem_M_toritaiyo.Toritaiyoname(cntjj)

            '        If taisyostr = taisyostr.Replace(searchstr, "") Then
            '            dgv(2, cntii - 1).Value = taisyostr
            '            dgv(1, cntii - 1).Style.BackColor = Color.Pink
            '            dgv(2, cntii - 1).Style.BackColor = Color.Pink
            '        Else
            '            dgv(1, cntii - 1).Value = RelItem_M_toritaiyo.Toritaiyono(cntjj)
            '            dgv(2, cntii - 1).Value = RelItem_M_toritaiyo.Toritaiyoname(cntjj)
            '            dgv(1, cntii - 1).Style.BackColor = Color.Beige
            '            dgv(2, cntii - 1).Style.BackColor = Color.Beige
            '            Exit For
            '        End If

            '    Next

            'Next
            '20160523 完全一致による表示は保留にしておく -del end

        End Sub

        ''' <summary>
        ''' 画像タイトルグリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal rowcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックス表示対象外の列の場合は処理を抜ける
            If Not (colcurrent = 3 OrElse colcurrent = 5) Then
                Exit Sub
            End If

            '20160610 部屋画像の割付の削除 -chg sta
            ''20160531 画像紐付処理の修正 -add sta
            ''移行元画像種別が部屋の場合は処理を抜ける
            'Dim dgv As New DataGridView
            'dgv = Njc.Frm.RelationFrm.dgvGazo
            'Dim tmp_V7_syubetu As String = dgv(0, rowcurrent).Value
            'If tmp_V7_syubetu = "部屋" Then
            '    Exit Sub
            'End If
            ''20160531 画像紐付処理の修正 -add end
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvGazo
            '20160610 部屋画像の割付の削除 -chg end

            '20160610 画像紐付箇所の修正 リスト内の空文字列の除去 -del sta
            ''コンボリスト作成
            'cellcombo.Items.Add("")
            '20160610 画像紐付箇所の修正 リスト内の空文字列の除去 -del end

            '20160531 画像紐付処理の修正 -del sta
            '上部へ移動
            ''グリッド取得
            'Dim dgv As New DataGridView
            'dgv = Njc.Frm.RelationFrm.dgvGazo
            '20160531 画像紐付処理の修正 -del end

            '親項目取得用
            Dim tmp_kbn As String = IIf(dgv(3, rowcurrent).Value Is Nothing, "", dgv(3, rowcurrent).Value)

            Select Case colcurrent
                Case 3
                    For cntii = 1 To RelItem_M_gazo_title_kbn.itemcnt - 1
                        '20160531 画像紐付処理の修正 -chg sta
                        '部屋の要素は除外する
                        'cellcombo.Items.Add(RelItem_M_gazo_title_kbn.Gazo_kbnname(cntii))
                        If RelItem_M_gazo_title_kbn.Gazo_kbnname(cntii) <> "部屋" Then
                            cellcombo.Items.Add(RelItem_M_gazo_title_kbn.Gazo_kbnname(cntii))
                        End If
                        '20160531 画像紐付処理の修正 -chg end
                    Next
                Case 5
                    If tmp_kbn = "周辺" Then
                        For cntii = 1 To RelItem_M_gazo_title_name.itemcnt - 1
                            cellcombo.Items.Add(RelItem_M_gazo_title_name.Gazo_name(cntii))
                        Next
                    End If
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

            '対象列以外の場合は処理を抜ける
            If Not (colcurrent = 3 Or colcurrent = 4 Or colcurrent = 5) Then
                Return rtn
            End If

            '空文字が設定された場合はセルを消去する処理に移行して抜ける
            If currentvalue = "" Then
                Select Case colcurrent
                    Case 3
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                    Case 4
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    Case 5
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                End Select
                Return rtn
            End If

            '20160610 周辺画像の画像No最大値修正 -add sta
            '画像種別が変更された場合は画像Noを初期化(空文字)にして処理を抜ける
            '※物件、周辺のみなのでセル値チェンジイベントで拾う
            If colcurrent = 3 Then
                dgv(colcurrent + 1, rowcurrent).Value = ""
                dgv(colcurrent + 2, rowcurrent).Value = ""
            End If
            '20160610 周辺画像の画像No最大値修正 -add end

            '10種別を取得
            Dim tmp_syubetu As String = dgv(3, rowcurrent).Value

            '20160610 周辺画像の画像No最大値修正 -chg sta
            ''数値範囲チェック
            'If colcurrent = 4 Then
            '    Dim min As Integer = 1
            '    Dim max As Integer = 20
            '    If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
            '        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '        dgv(colcurrent, rowcurrent).Value = ""
            '        rtn = False
            '        Return rtn
            '    End If
            'End If

            '数値範囲チェック
            Dim min As Integer = 1
            Dim max As Integer = 20

            '周辺の場合は最大値を10に変更
            If tmp_syubetu = "周辺" Then
                max = 10
            End If

            If colcurrent = 4 Then
                If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
                    MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    dgv(colcurrent, rowcurrent).Value = ""
                    rtn = False
                    Return rtn
                End If
            End If
            '20160610 周辺画像の画像No最大値修正 -chg end

            '設計書に記載がなかったためとりあえずコメントアウト
            ''一致する値を取得
            'If tmp_syubetu = "周辺" Then

            '    For cntii = 1 To RelItem_M_gazo_title_name.itemcnt - 1

            '        Dim searchstr As String = ""
            '        Dim outstr As String = ""
            '        Dim chgcol As Integer = 0

            '        Select Case colcurrent
            '            Case 4
            '                searchstr = RelItem_M_gazo_title_name.Gazo_no(cntii)
            '                outstr = RelItem_M_gazo_title_name.Gazo_name(cntii)
            '                chgcol = colcurrent + 1
            '            Case 5
            '                searchstr = RelItem_M_gazo_title_name.Gazo_name(cntii)
            '                outstr = RelItem_M_gazo_title_name.Gazo_no(cntii)
            '                chgcol = colcurrent - 1
            '        End Select

            '        If chgcol <> 0 Then
            '            dgv(chgcol, rowcurrent).Value = ""
            '        End If

            '        If currentvalue = searchstr Then
            '            dgv(chgcol, rowcurrent).Value = outstr
            '            Exit For
            '        End If

            '    Next

            'End If

            Return rtn

        End Function

        ''' <summary>
        ''' 画像タイトルNoの重複チェック
        ''' </summary>
        ''' <param name="rowno"></param>
        ''' <param name="initflg"></param>
        ''' <remarks></remarks>
        Public Sub Chk_Duplicate(ByVal rowno As Integer, ByVal initflg As Boolean)

            Dim dgv As DataGridView = Njc.Frm.RelationFrm.dgvGazo
            Dim list_keytotal As New List(Of String)
            Dim colno_main As Integer = 3
            Dim colno_sub As Integer = 4

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
        ''' セルの読み取り専用制御
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Public Sub Set_CellReadOnly(ByVal colcurrent As Integer, ByVal rowcurrent As Integer)

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvGazo

            '親項目取得用
            Dim tmp_kbn As String = IIf(dgv(3, rowcurrent).Value Is Nothing, "", dgv(3, rowcurrent).Value)

            'セルの読み取り制御処理
            '20160531 画像紐付処理の修正 -chg sta
            'Select Case tmp_kbn
            '    Case ""
            '        dgv(4, rowcurrent).ReadOnly = True
            '        dgv(5, rowcurrent).ReadOnly = True
            '    Case "周辺"
            '        dgv(4, rowcurrent).ReadOnly = False
            '        dgv(5, rowcurrent).ReadOnly = False
            '    Case Else
            '        dgv(4, rowcurrent).ReadOnly = False
            '        dgv(5, rowcurrent).ReadOnly = True
            'End Select
            Select Case tmp_kbn
                Case ""
                    dgv(4, rowcurrent).ReadOnly = True
                    dgv(5, rowcurrent).ReadOnly = True
                Case "物件"
                    dgv(4, rowcurrent).ReadOnly = False
                    dgv(5, rowcurrent).ReadOnly = True
                Case "周辺"
                    dgv(4, rowcurrent).ReadOnly = False
                    dgv(5, rowcurrent).ReadOnly = False
                    '20160610 部屋画像の割付の削除 -del sta
                    'Case "部屋"
                    '    dgv(3, rowcurrent).ReadOnly = True
                    '    dgv(4, rowcurrent).ReadOnly = False
                    '    dgv(5, rowcurrent).ReadOnly = True
                    '20160610 部屋画像の割付の削除 -del end
            End Select
            '20160531 画像紐付処理の修正 -chg end

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
            dgv = Njc.Frm.RelationFrm.dgvGazo
            lbl_relcnt = Njc.Frm.RelationFrm.lblSumiZanGazo
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllGazo

            '紐付全件数取得
            totalcnt = dgv.RowCount

            '紐付未設定件数取得
            For cntii = 0 To totalcnt - 1

                Dim tmp_kbn As String = dgv(3, cntii).Value
                Dim tmp_no As String = dgv(4, cntii).Value
                Dim tmp_name As String = dgv(5, cntii).Value

                If tmp_kbn = "物件" And tmp_no <> "" Then
                    relcnt = relcnt + 1
                ElseIf tmp_kbn = "周辺" And tmp_no <> "" And tmp_name <> "" Then
                    relcnt = relcnt + 1
                End If

            Next

            '表示
            lbl_relcnt.Text = relcnt.ToString
            lbl_totalcnt.Text = totalcnt.ToString

            '色設定
            Njc.Frm.RelationFrm.lblCntGazoBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblSumiZanGazo.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllGazo.ForeColor = Color.Black

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