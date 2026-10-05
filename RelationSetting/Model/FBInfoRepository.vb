Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "FBフォーマット割付取得"

    Public Class M_FBInfo_Repository

        ''' <summary>
        ''' 10FB関連情報取得
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
            Dim model_item As New Njc.Model.M_FBInfo_Model(reccnt)

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
                            Case "フォーマットNo"
                                .Fmt_kbn(cntii) = fldvalue
                            Case "フォーマット種別"
                                .Fmt_kbnname(cntii) = fldvalue
                            Case "登録No"
                                .Fmt_keyno(cntii) = fldvalue
                            Case "フォーマット名"
                                .Fmt_name(cntii) = fldvalue
                        End Select
                    Next

                End With

            Next

            'モデルの引渡し
            RelItem_M_FBInfo = Nothing
            RelItem_M_FBInfo = model_item

            '自社情報を中間ファイルから取得
            If RelItem_M_jisyakoza Is Nothing Then
                Dim obj_jisya As New Njc.Repository.M_jisyakoza_Repository
                Call obj_jisya.ReadRelItem()
            End If

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10FB関連情報抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry()

            Dim tmp_sql As String = ""

            tmp_sql = tmp_sql & " SELECT DISTINCT  "
            tmp_sql = tmp_sql & " 	 CASE fmt_kbn  "
            tmp_sql = tmp_sql & " 		WHEN 1 THEN '1' "
            tmp_sql = tmp_sql & " 		WHEN 2 THEN '2' "
            tmp_sql = tmp_sql & " 		/*WHEN 3 THEN '3'*/ "
            tmp_sql = tmp_sql & " 		WHEN 4 THEN '4' "
            tmp_sql = tmp_sql & " 		WHEN 5 THEN '5' "
            tmp_sql = tmp_sql & " 		WHEN 6 THEN '6' "
            tmp_sql = tmp_sql & " 		/*ELSE '-1'*/ "
            tmp_sql = tmp_sql & " 	 END AS [フォーマットNo] "
            tmp_sql = tmp_sql & " 	,CASE fmt_kbn  "
            tmp_sql = tmp_sql & " 		WHEN 1 THEN '口座振替' "
            tmp_sql = tmp_sql & " 		WHEN 2 THEN '総合振込' "
            tmp_sql = tmp_sql & " 		/*WHEN 3 THEN '？'*/ "
            tmp_sql = tmp_sql & " 		WHEN 4 THEN '振込入金明細' "
            tmp_sql = tmp_sql & " 		WHEN 5 THEN 'コンビニ収納請求' "
            tmp_sql = tmp_sql & " 		WHEN 6 THEN 'コンビニ収納入金' "
            tmp_sql = tmp_sql & " 		/*ELSE '未定義または未設定'*/ "
            tmp_sql = tmp_sql & " 	 END AS [フォーマット種別] "
            tmp_sql = tmp_sql & " 	,fmt_keyno AS [登録No] "
            tmp_sql = tmp_sql & " 	,fmt_name AS [フォーマット名] "
            tmp_sql = tmp_sql & " FROM m_fb_zenfmt "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT "
            tmp_sql = tmp_sql & " 	 4 AS [フォーマットNo] "
            tmp_sql = tmp_sql & " 	,'入出金' AS [フォーマット種別] "
            tmp_sql = tmp_sql & " 	,orgfmt_no AS [登録No] "
            tmp_sql = tmp_sql & " 	,orgfmt_name AS [フォーマット名] "
            tmp_sql = tmp_sql & " FROM m_fb_orgfmtfkom "

            Return tmp_sql

        End Function

        ''' <summary>
        ''' FB関連情報グリッドの初期設定
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Initialize_Dgv(ByVal dgv As DataGridView)

            With dgv
                '20160725 列幅修正 -del
                '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                .Columns(2).Frozen = True
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
                '20160725 FBフォーマットの数値入力制御解除対応 -del
                '.Columns(10).ReadOnly = True
                .Columns(10).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(10).DefaultCellStyle.BackColor = Color.Beige
                .Columns(11).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(11).DefaultCellStyle.BackColor = Color.Beige
                .Columns(12).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(12).DefaultCellStyle.BackColor = Color.Beige
                .Columns(13).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(13).DefaultCellStyle.BackColor = Color.Beige
                '20160725 列幅修正 -add
                Call EtcMethod.Set_DgvColumnSize(dgv)
            End With

        End Sub

        ''' <summary>
        ''' FB関連情報グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            Dim tmp_furiiraicnt As Integer = 0
            Dim tmp_kozafurikae As Integer = 0
            Dim tmp_nssetting As Integer = 0

            'データ数取得
            If RelItem_M_FBInfo_Furiirai_Rev7 IsNot Nothing Then
                tmp_furiiraicnt = RelItem_M_FBInfo_Furiirai_Rev7.ItemCnt - 1
            End If
            If RelItem_M_FBInfo_Kozafurikae_Rev7 IsNot Nothing Then
                tmp_kozafurikae = RelItem_M_FBInfo_Kozafurikae_Rev7.ItemCnt - 1
            End If
            If RelItem_M_FBInfo_Nssetting_Rev7 IsNot Nothing Then
                tmp_nssetting = RelItem_M_FBInfo_Nssetting_Rev7.ItemCnt - 1
            End If

            'グリッド行数設定
            Dim tmp_rowcnt As Integer = tmp_furiiraicnt + tmp_kozafurikae + tmp_nssetting
            If tmp_rowcnt = 0 Then
                Exit Sub
            End If
            dgv.RowCount = tmp_furiiraicnt + tmp_kozafurikae + tmp_nssetting

            'V7項目の表示
            If tmp_furiiraicnt > 0 Then
                For cntii = 1 To tmp_furiiraicnt
                    dgv(0, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Fmt_syubetu(cntii)
                    dgv(1, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Fkom_no(cntii)
                    dgv(2, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Fkom_name(cntii)
                    dgv(3, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Kinyu_no(cntii)
                    dgv(4, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Kinyu_name(cntii)
                    dgv(5, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Ten_no(cntii)
                    dgv(6, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Ten_name(cntii)
                    dgv(7, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Kosyu_no(cntii)
                    dgv(8, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Koza_no(cntii)
                    dgv(9, cntii - 1).Value = RelItem_M_FBInfo_Furiirai_Rev7.Koza_meigi(cntii)
                Next
            End If

            Dim cnt_mid As Integer = tmp_furiiraicnt + tmp_kozafurikae
            If tmp_kozafurikae > 0 Then
                Dim cnt_furikaeitem As Integer = 1
                For cntii = tmp_furiiraicnt + 1 To cnt_mid
                    dgv(0, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Fmt_syubetu(cnt_furikaeitem)
                    dgv(1, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Fkae_no(cnt_furikaeitem)
                    dgv(2, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Fkae_name(cnt_furikaeitem)
                    dgv(3, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Kinyu_no(cnt_furikaeitem)
                    dgv(4, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Kinyu_name(cnt_furikaeitem)
                    dgv(5, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Ten_no(cnt_furikaeitem)
                    dgv(6, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Ten_name(cnt_furikaeitem)
                    dgv(7, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Kosyu_no(cnt_furikaeitem)
                    dgv(8, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Koza_no(cnt_furikaeitem)
                    dgv(9, cntii - 1).Value = RelItem_M_FBInfo_Kozafurikae_Rev7.Koza_meigi(cnt_furikaeitem)
                    cnt_furikaeitem = cnt_furikaeitem + 1
                Next
            End If

            If tmp_nssetting > 0 Then
                Dim cnt_nssetitem As Integer = 1
                For cntii = cnt_mid + 1 To tmp_rowcnt
                    dgv(0, cntii - 1).Value = RelItem_M_FBInfo_Nssetting_Rev7.Fmt_syubetu(cnt_nssetitem)
                    dgv(1, cntii - 1).Value = RelItem_M_FBInfo_Nssetting_Rev7.Ns_no(cnt_nssetitem)
                    dgv(2, cntii - 1).Value = RelItem_M_FBInfo_Nssetting_Rev7.Ns_name(cnt_nssetitem)
                    cnt_nssetitem = cnt_nssetitem + 1
                Next
            End If

            '20161209 前回設定値復元時にグレイアウトされていないセルが生じるエラーの修正 -del sta
            'フォーム画面にメソッドを移動するためコメントアウト
            ''フォーマット種別が入出金以外の場合、自社情報設定箇所をグレイアウトさせる
            'For cntkk = 1 To dgv.RowCount
            '    Dim fmtkbn As String = dgv(0, cntkk - 1).Value
            '    If fmtkbn <> "入出金" Then
            '        dgv(12, cntkk - 1).Style.BackColor = Color.LightGray
            '        dgv(13, cntkk - 1).Style.BackColor = Color.LightGray
            '    End If
            'Next
            '20161209 前回設定値復元時にグレイアウトされていないセルが生じるエラーの修正 -del end

        End Sub

        ''' <summary>
        ''' FB関連情報グリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal rowcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックス表示対象外の列の場合は処理を抜ける
            If Not (colcurrent = 11 Or colcurrent = 13) Then
                Exit Sub
            End If

            'コンボリスト作成
            cellcombo.Items.Add("")

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvFBInfo

            '親項目取得用
            Dim tmp_kbn As String = IIf(dgv(0, rowcurrent).Value Is Nothing, "", dgv(0, rowcurrent).Value)

            Select Case colcurrent
                Case 11
                    For cntii = 1 To RelItem_M_FBInfo.itemcnt - 1
                        If tmp_kbn = RelItem_M_FBInfo.Fmt_kbnname(cntii) Then
                            cellcombo.Items.Add(RelItem_M_FBInfo.Fmt_name(cntii))
                        End If
                    Next
                Case 13
                    For cntii = 1 To RelItem_M_jisyakoza.itemcnt - 1
                        If tmp_kbn = "入出金" Then
                            cellcombo.Items.Add(RelItem_M_jisyakoza.Jisya_name(cntii))
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
            If currentvalue = "" Then
                '20160701 指摘事項まとめファイルの対応 読取専用条件追加 -chg sta
                'Select Case colcurrent
                '    Case 10, 12
                '        dgv(colcurrent + 1, rowcurrent).Value = ""
                '    Case 11, 13
                '        dgv(colcurrent - 1, rowcurrent).Value = ""
                'End Select
                Select Case colcurrent
                    Case 10
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                    Case 11
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                    Case 12
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    Case 13
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                End Select
                '20160701 指摘事項まとめファイルの対応 読取専用条件追加 -chg end
                Return rtn
            End If

            'フォーマット区分を取得
            Dim fmtkbn As String = dgv(0, rowcurrent).Value

            '数値範囲チェック
            Select Case colcurrent
                Case 10
                    Dim list_fmtno As New List(Of String)
                    For cntii = 1 To RelItem_M_FBInfo.ItemCnt - 1
                        Dim tmp_10fmtkbn As String = RelItem_M_FBInfo.Fmt_kbnname(cntii)
                        If fmtkbn = tmp_10fmtkbn Then
                            list_fmtno.Add(RelItem_M_FBInfo.Fmt_keyno(cntii))
                        End If
                    Next
                    If list_fmtno.Contains(currentvalue) = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    End If
                Case 12
                    Dim list_jisyano As New List(Of String)
                    For cntii = 1 To RelItem_M_jisyakoza.ItemCnt - 1
                        list_jisyano.Add(RelItem_M_jisyakoza.Jisya_no(cntii))
                    Next
                    If list_jisyano.Contains(currentvalue) = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    End If
            End Select

            '一致する値を取得
            Dim inputsyogo_fmt As String = fmtkbn & "-" & currentvalue
            Dim inputsyogo_jisya As String = currentvalue
            Dim searchkbn As String = ""
            Dim searchstr As String = ""
            Dim searchjisyastr As String = ""
            Dim outstr As String = ""
            Dim chgcol As Integer = 0

            Select Case colcurrent
                Case 10
                    chgcol = colcurrent + 1
                    For cntii = 1 To RelItem_M_FBInfo.itemcnt - 1
                        searchkbn = RelItem_M_FBInfo.Fmt_kbnname(cntii)
                        searchstr = RelItem_M_FBInfo.Fmt_keyno(cntii)
                        outstr = RelItem_M_FBInfo.Fmt_name(cntii)
                        Dim outputsyogo As String = searchkbn & "-" & searchstr
                        If inputsyogo_fmt = outputsyogo Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        Else
                            outstr = ""
                        End If
                    Next
                Case 11
                    chgcol = colcurrent - 1
                    For cntii = 1 To RelItem_M_FBInfo.itemcnt - 1
                        searchkbn = RelItem_M_FBInfo.Fmt_kbnname(cntii)
                        searchstr = RelItem_M_FBInfo.Fmt_name(cntii)
                        outstr = RelItem_M_FBInfo.Fmt_keyno(cntii)
                        Dim outputsyogo As String = searchkbn & "-" & searchstr
                        If inputsyogo_fmt = outputsyogo Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        Else
                            outstr = ""
                        End If
                    Next

                Case 12
                    If fmtkbn = "入出金" Then
                        chgcol = colcurrent + 1
                        For cntii = 1 To RelItem_M_jisyakoza.itemcnt - 1
                            searchjisyastr = RelItem_M_jisyakoza.Jisya_no(cntii)
                            outstr = RelItem_M_jisyakoza.Jisya_name(cntii)
                            If inputsyogo_jisya = searchjisyastr Then
                                dgv(chgcol, rowcurrent).Value = outstr
                                Exit For
                            Else
                                outstr = ""
                            End If
                        Next
                    Else
                        Return rtn
                    End If

                Case 13
                    If fmtkbn = "入出金" Then
                        chgcol = colcurrent - 1
                        For cntii = 1 To RelItem_M_jisyakoza.itemcnt - 1
                            searchjisyastr = RelItem_M_jisyakoza.Jisya_name(cntii)
                            outstr = RelItem_M_jisyakoza.Jisya_no(cntii)
                            If inputsyogo_jisya = searchjisyastr Then
                                dgv(chgcol, rowcurrent).Value = outstr
                                Exit For
                            Else
                                outstr = ""
                            End If
                        Next
                    Else
                        Return rtn
                    End If
            End Select

            Return rtn

        End Function

        ''' <summary>
        ''' セルの読み取り専用制御
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Public Sub Set_CellReadOnly(ByVal colcurrent As Integer, ByVal rowcurrent As Integer)

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvFBInfo

            '親項目取得用
            Dim tmp_fmtkbn As String = IIf(dgv(0, rowcurrent).Value Is Nothing, "", dgv(0, rowcurrent).Value)

            '20160701 指摘事項まとめファイルの対応 読取専用条件追加 -add sta
            Dim tmp_fmtno As String = IIf(dgv(10, rowcurrent).Value Is Nothing, "", dgv(10, rowcurrent).Value)
            Dim tmp_fmtname As String = IIf(dgv(11, rowcurrent).Value Is Nothing, "", dgv(11, rowcurrent).Value)
            '20160701 指摘事項まとめファイルの対応 読取専用条件追加 -add end

            'セルの読み取り制御処理
            If tmp_fmtkbn = "入出金" And (tmp_fmtno <> "" Or tmp_fmtname <> "") Then          '20160701 指摘事項まとめファイルの対応 読取専用条件追加 -add ()を追加
                dgv(12, rowcurrent).ReadOnly = False
                dgv(13, rowcurrent).ReadOnly = False
            Else
                dgv(12, rowcurrent).ReadOnly = True
                dgv(13, rowcurrent).ReadOnly = True
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
                Case 10
                    dgv(colcurrent + 2, rowcurrent).Value = ""
                    dgv(colcurrent + 3, rowcurrent).Value = ""
                Case 11
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
            dgv = Njc.Frm.RelationFrm.dgvFBInfo
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiFbFmt
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllFbFmt

            '紐付全件数取得
            totalcnt = dgv.RowCount

            '紐付未設定件数取得
            For cntii = 0 To totalcnt - 1

                Dim tmp_kbn As String = dgv(0, cntii).Value
                Dim tmp_no As String = dgv(10, cntii).Value
                Dim tmp_name As String = dgv(11, cntii).Value
                Dim tmp_nosub As String = dgv(12, cntii).Value
                Dim tmp_namesub As String = dgv(13, cntii).Value
                Dim tmp_chkstr As String = tmp_no & tmp_name
                Dim tmp_chkstrsub As String = tmp_nosub & tmp_namesub

                If tmp_kbn <> "入出金" And tmp_chkstr <> "" Then
                    relcnt = relcnt + 1
                ElseIf tmp_kbn = "入出金" And tmp_chkstr <> "" And tmp_chkstrsub <> "" Then
                    relcnt = relcnt + 1
                End If

            Next

            '表示
            lbl_relcnt.Text = relcnt.ToString
            lbl_totalcnt.Text = totalcnt.ToString

            '色設定
            Njc.Frm.RelationFrm.lblCntFbFmtBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiFbFmt.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllFbFmt.ForeColor = Color.Black

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