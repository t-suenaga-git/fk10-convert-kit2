Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query

Namespace Njc.Repository

#Region "設備マスタ取得"

    Public Class M_setubi_Repository

        ''' <summary>
        ''' 10設備総合取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadDB(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim rtn As Boolean = True

            rtn = Me.ReadDB_Grp(sqlcnnV10)
            rtn = Me.ReadDB_Ms(sqlcnnV10)
            rtn = Me.ReadDB_Komk(sqlcnnV10)

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10設備グループ取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadDB_Grp(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim reccnt As Integer                                                           '抽出レコード件数
            Dim fldname As String                                                           '該当項目名
            Dim fldvalue As String                                                          '該当値
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True

            '抽出・レコード数の取得
            reccnt = DBExec.Exec_DataTable(Me.Get_UseQry_Grp(), sqlcnnV10, readtbl)

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_setubi_Grp_Model(reccnt)

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
                            Case "setubi_grpguid"
                                .Setubi_grpguid_grp(cntii) = fldvalue
                            Case "setubi_grpsortorder"
                                .Setubi_grpsortorder(cntii) = fldvalue
                            Case "setubi_grpname"
                                .Setubi_grpname(cntii) = fldvalue
                        End Select
                    Next

                End With

            Next

            'モデルの引渡し
            RelItem_M_setubi_Grp = Nothing
            RelItem_M_setubi_Grp = model_item

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10設備グループ抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry_Grp()

            Dim strsql As String

            strsql = ""
            strsql = strsql & " /*20160829 エレベーター移行対応 chg sta*/ "
            strsql = strsql & " /*SELECT * FROM m_setubi_grp ORDER BY setubi_grpsortorder*/ "
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 setubi_grpguid "
            strsql = strsql & " 	,setubi_grpsortorder "
            strsql = strsql & " 	,setubi_grpname "
            strsql = strsql & " FROM m_setubi_grp "
            strsql = strsql & " UNION "
            strsql = strsql & " SELECT TOP 1 "
            strsql = strsql & " 	 NEWID() AS setubi_grpguid		/*UNION用にguidを新規作成(紐付ツールでは使用しない)*/ "
            strsql = strsql & " 	,999 AS setubi_grpsortorder "
            strsql = strsql & " 	,'エレベーター' AS setubi_grpname "
            strsql = strsql & " FROM m_setubi_grp "
            strsql = strsql & " ORDER BY setubi_grpsortorder "
            strsql = strsql & " /*20160829 エレベーター移行対応 chg end*/ "
            Return strsql

        End Function

        ''' <summary>
        ''' 10設備項目取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadDB_Ms(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim reccnt As Integer                                                           '抽出レコード件数
            Dim fldname As String                                                           '該当項目名
            Dim fldvalue As String                                                          '該当値
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True

            '抽出・レコード数の取得
            reccnt = DBExec.Exec_DataTable(Me.Get_UseQry_Ms(), sqlcnnV10, readtbl)

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_setubi_Ms_Model(reccnt)

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
                            Case "setubi_grpguid"
                                .Setubi_grpguid(cntii) = fldvalue
                            Case "setubi_grpsortorder"
                                .Setubi_grpsortorder(cntii) = fldvalue
                            Case "setubi_grpname"
                                .Setubi_grpname(cntii) = fldvalue
                            Case "setubi_guid"
                                .Setubi_guid(cntii) = fldvalue
                            Case "setubi_sortorder"
                                .Setubi_sortorder(cntii) = fldvalue
                            Case "setubi_name"
                                .Setubi_name(cntii) = fldvalue
                        End Select
                    Next

                End With

            Next

            'モデルの引渡し
            RelItem_M_setubi_Ms = Nothing
            RelItem_M_setubi_Ms = model_item

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10設備項目抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry_Ms()

            Dim strsql As String

            strsql = ""
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 MSG.setubi_grpguid "
            strsql = strsql & " 	,MSG.setubi_grpsortorder "
            strsql = strsql & " 	,MSG.setubi_grpname "
            strsql = strsql & " 	,MS.setubi_guid "
            strsql = strsql & " 	,MS.setubi_sortorder "
            strsql = strsql & " 	,MS.setubi_name "
            strsql = strsql & " FROM m_setubi AS MS "
            strsql = strsql & " LEFT JOIN m_setubi_grp AS MSG ON MS.setubi_grpguid = MSG.setubi_grpguid "
            strsql = strsql & " /*20160722 設備情報取得時エラーの対応 add */ "
            strsql = strsql & " WHERE ISNULL(setubi_grpname,'') <> '' AND ISNULL(setubi_name,'') <> '' "
            strsql = strsql & " /*20160829 エレベーター移行対応 add sta*/ "
            strsql = strsql & " UNION "
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 NEWID() AS setubi_grpguid		/*UNION用にguidを新規作成(紐付ツールでは使用しない)*/ "
            strsql = strsql & " 	,999 AS setubi_grpsortorder "
            strsql = strsql & " 	,'エレベーター' AS setubi_grpname "
            strsql = strsql & " 	,NEWID() AS setubi_grpguid		/*UNION用にguidを新規作成(紐付ツールでは使用しない)*/ "
            strsql = strsql & " 	,1 AS setubi_sortorder "
            strsql = strsql & " 	,'エレベーター有' AS setubi_name "
            strsql = strsql & " UNION "
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 NEWID() AS setubi_grpguid		/*UNION用にguidを新規作成(紐付ツールでは使用しない)*/ "
            strsql = strsql & " 	,999 AS setubi_grpsortorder "
            strsql = strsql & " 	,'エレベーター' AS setubi_grpname "
            strsql = strsql & " 	,NEWID() AS setubi_grpguid		/*UNION用にguidを新規作成(紐付ツールでは使用しない)*/ "
            strsql = strsql & " 	,2 AS setubi_sortorder "
            strsql = strsql & " 	,'ホームエレベーター有' AS setubi_name "
            strsql = strsql & " UNION "
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 NEWID() AS setubi_grpguid		/*UNION用にguidを新規作成(紐付ツールでは使用しない)*/ "
            strsql = strsql & " 	,999 AS setubi_grpsortorder "
            strsql = strsql & " 	,'エレベーター' AS setubi_grpname "
            strsql = strsql & " 	,NEWID() AS setubi_grpguid		/*UNION用にguidを新規作成(紐付ツールでは使用しない)*/ "
            strsql = strsql & " 	,3 AS setubi_sortorder "
            strsql = strsql & " 	,'非常用エレベーター有' AS setubi_name "
            strsql = strsql & " /*20160829 エレベーター移行対応 add end*/ "
            strsql = strsql & " ORDER BY MSG.setubi_grpsortorder,MS.setubi_sortorder "

            Return strsql

        End Function

        ''' <summary>
        ''' 10設備内容取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadDB_Komk(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim reccnt As Integer                                                           '抽出レコード件数
            Dim fldname As String                                                           '該当項目名
            Dim fldvalue As String                                                          '該当値
            Dim readtbl As New DataTable
            Dim rtn As Boolean = True

            '抽出・レコード数の取得
            reccnt = DBExec.Exec_DataTable(Me.Get_UseQry_Komk(), sqlcnnV10, readtbl)

            '初期化(配列要素数決定)
            Dim model_item As New Njc.Model.M_setubi_Komk_Model(reccnt)

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
                            Case "setubi_grpguid"
                                .Setubi_grpguid(cntii) = fldvalue
                            Case "setubi_guid"
                                .Setubi_guid(cntii) = fldvalue
                            Case "komok_guid"
                                .Komok_guid(cntii) = fldvalue
                            Case "setubi_grpsortorder"
                                .Setubi_grpsortorder(cntii) = fldvalue
                            Case "setubi_grpname"
                                .Setubi_grpname(cntii) = fldvalue
                            Case "setubi_sortorder"
                                .Setubi_sortorder(cntii) = fldvalue
                            Case "setubi_name"
                                .Setubi_name(cntii) = fldvalue
                            Case "komok_sortorder"
                                .Komok_sortorder(cntii) = fldvalue
                            Case "komok_name"
                                .Komok_name(cntii) = fldvalue
                        End Select
                    Next

                End With

            Next

            'モデルの引渡し
            RelItem_M_setubi_Komk = Nothing
            RelItem_M_setubi_Komk = model_item

            '返却
            Return rtn

        End Function

        ''' <summary>
        ''' 10設備内容抽出クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Function Get_UseQry_Komk()

            Dim strsql As String

            strsql = ""
            strsql = strsql & " SELECT "
            strsql = strsql & " 	 SEG.setubi_grpguid "
            strsql = strsql & " 	,SE.setubi_guid "
            strsql = strsql & " 	,komok_guid  "
            strsql = strsql & " 	,setubi_grpsortorder "
            strsql = strsql & " 	,setubi_grpname  "
            strsql = strsql & " 	,setubi_sortorder "
            strsql = strsql & " 	,setubi_name "
            strsql = strsql & " 	,komok_sortorder "
            strsql = strsql & " 	,komok_name "
            strsql = strsql & " FROM m_setubi_grp AS SEG "
            strsql = strsql & " LEFT JOIN m_setubi AS SE ON SEG.setubi_grpguid = SE.setubi_grpguid "
            strsql = strsql & " LEFT JOIN m_setubi_lst AS SEL ON SE.setubi_guid = SEL.setubi_guid "
            strsql = strsql & " /*20160722 設備情報取得時エラーの対応 add */ "
            strsql = strsql & " WHERE ISNULL(setubi_grpname,'') <> '' AND ISNULL(setubi_name,'') <> '' AND ISNULL(komok_name,'') <> '' "
            strsql = strsql & " ORDER BY setubi_grpsortorder,setubi_sortorder,komok_sortorder "

            Return strsql

        End Function

        ''' <summary>
        ''' 設備グリッドの初期設定
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Initialize_Dgv(ByVal dgv As DataGridView)

            With dgv
                '20160701 指摘事項まとめファイルの対応 幅の可変 -del
                '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
                '20160701 指摘事項まとめファイルの対応 幅の可変 -add
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                .Columns(1).Frozen = True
                .Columns(0).ReadOnly = True
                .Columns(0).DefaultCellStyle.BackColor = Color.Azure
                .Columns(0).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(1).ReadOnly = True
                .Columns(1).DefaultCellStyle.BackColor = Color.Azure
                .Columns(1).HeaderCell.Style.BackColor = Color.PowderBlue
                .Columns(2).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(2).DefaultCellStyle.BackColor = Color.Beige
                '.Columns(2).ReadOnly = True     '20160525 全体的な動作の修正 -del
                .Columns(3).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(3).DefaultCellStyle.BackColor = Color.Beige
                .Columns(4).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(4).DefaultCellStyle.BackColor = Color.Beige
                '.Columns(4).ReadOnly = True     '20160525 全体的な動作の修正 -del
                .Columns(5).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(5).DefaultCellStyle.BackColor = Color.Beige
                .Columns(6).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(6).DefaultCellStyle.BackColor = Color.Beige
                '.Columns(6).ReadOnly = True     '20160525 全体的な動作の修正 -del
                .Columns(7).HeaderCell.Style.BackColor = Color.Bisque
                .Columns(7).DefaultCellStyle.BackColor = Color.Beige
            End With

        End Sub

        ''' <summary>
        ''' 設備グリッドへのデータ表示
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <remarks></remarks>
        Public Sub Set_Dgv(ByVal dgv As DataGridView)

            '20160525 全体的な動作の修正 -add sta
            If RelItem_M_setubi_Rev7 Is Nothing Then
                Exit Sub
            End If
            '20160525 全体的な動作の修正 -add end

            'グリッド行数設定
            dgv.RowCount = RelItem_M_setubi_Rev7.ItemCnt - 1

            'V7項目の表示
            For cntii = 1 To RelItem_M_setubi_Rev7.ItemCnt - 1
                dgv(0, cntii - 1).Value = RelItem_M_setubi_Rev7.Setubi_name(cntii)
                dgv(1, cntii - 1).Value = RelItem_M_setubi_Rev7.Setubi_lstname(cntii)
            Next

            '10紐付項目表示
            Dim relitemcnt10 As Integer = RelItem_M_setubi_Komk.ItemCnt - 1

            For cntii = 1 To dgv.RowCount

                '対象文字列格納
                Dim taisyostr As String = dgv(0, cntii - 1).Value
                Dim taisyostr_sub As String = dgv(1, cntii - 1).Value

                '20160905 エレベーターのマッチング修正 -chg sta
                'For cntjj = 1 To relitemcnt10

                '    '照合用文字列格納
                '    Dim searchstr As String = RelItem_M_setubi_Komk.Setubi_name(cntjj)
                '    Dim searchstr_sub As String = RelItem_M_setubi_Komk.Komok_name(cntjj)

                '    If taisyostr <> taisyostr.Replace(searchstr, "") Then
                '        dgv(2, cntii - 1).Value = RelItem_M_setubi_Komk.Setubi_grpsortorder(cntjj)
                '        dgv(3, cntii - 1).Value = RelItem_M_setubi_Komk.Setubi_grpname(cntjj)
                '        dgv(4, cntii - 1).Value = RelItem_M_setubi_Komk.Setubi_sortorder(cntjj)
                '        dgv(5, cntii - 1).Value = RelItem_M_setubi_Komk.Setubi_name(cntjj)
                '        If taisyostr_sub <> taisyostr_sub.Replace(searchstr_sub, "") Then
                '            dgv(6, cntii - 1).Value = RelItem_M_setubi_Komk.Komok_sortorder(cntjj)
                '            dgv(7, cntii - 1).Value = RelItem_M_setubi_Komk.Komok_name(cntjj)
                '            Exit For
                '        End If
                '        dgv(4, cntii - 1).Style.BackColor = Color.Beige
                '        dgv(5, cntii - 1).Style.BackColor = Color.Beige
                '    End If

                'Next

                'エレベーターの有無で処理を分岐
                If taisyostr <> taisyostr.Replace("エレベータ", "") Then

                    Dim relitemgrpcnt10 As Integer = RelItem_M_setubi_Ms.ItemCnt - 1

                    For cntjj = 1 To relitemgrpcnt10
                        Dim searchstr As String = RelItem_M_setubi_Ms.Setubi_grpname(cntjj)
                        Dim searchstr_sub As String = RelItem_M_setubi_Ms.Setubi_name(cntjj)

                        If taisyostr <> taisyostr.Replace(searchstr, "") Then
                            dgv(2, cntii - 1).Value = RelItem_M_setubi_Ms.Setubi_grpsortorder(cntjj)
                            dgv(3, cntii - 1).Value = RelItem_M_setubi_Ms.Setubi_grpname(cntjj)
                            dgv(4, cntii - 1).Value = RelItem_M_setubi_Ms.Setubi_sortorder(cntjj)
                            dgv(5, cntii - 1).Value = RelItem_M_setubi_Ms.Setubi_name(cntjj)
                            Exit For
                        End If
                    Next

                Else

                    For cntjj = 1 To relitemcnt10

                        '照合用文字列格納
                        Dim searchstr As String = RelItem_M_setubi_Komk.Setubi_name(cntjj)
                        Dim searchstr_sub As String = RelItem_M_setubi_Komk.Komok_name(cntjj)

                        If taisyostr <> taisyostr.Replace(searchstr, "") Then
                            dgv(2, cntii - 1).Value = RelItem_M_setubi_Komk.Setubi_grpsortorder(cntjj)
                            dgv(3, cntii - 1).Value = RelItem_M_setubi_Komk.Setubi_grpname(cntjj)
                            dgv(4, cntii - 1).Value = RelItem_M_setubi_Komk.Setubi_sortorder(cntjj)
                            dgv(5, cntii - 1).Value = RelItem_M_setubi_Komk.Setubi_name(cntjj)
                            If taisyostr_sub <> taisyostr_sub.Replace(searchstr_sub, "") Then
                                dgv(6, cntii - 1).Value = RelItem_M_setubi_Komk.Komok_sortorder(cntjj)
                                dgv(7, cntii - 1).Value = RelItem_M_setubi_Komk.Komok_name(cntjj)
                                Exit For
                            End If
                            dgv(4, cntii - 1).Style.BackColor = Color.Beige
                            dgv(5, cntii - 1).Style.BackColor = Color.Beige
                        End If

                    Next

                End If
                '20160905 エレベーターのマッチング修正 -chg end
            Next

            '2016.04.15 設備の重複不可制御をコメントアウト -del sta 
            ''重複チェック
            'For cntrow = dgv.RowCount - 1 To 0 Step -1
            '    Call Me.Chk_Duplicate(cntrow, True)
            'Next
            '2016.04.15 設備の重複不可制御をコメントアウト -del sta 

        End Sub

        ''' <summary>
        ''' 設備グリッドのコンボボックスリスト作成
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <param name="currentvalue"></param>
        ''' <param name="cellcombo"></param>
        ''' <remarks></remarks>
        Public Sub Set_DgvCmbList(ByVal colcurrent As Integer, ByVal rowcurrent As Integer, ByVal currentvalue As String, ByRef cellcombo As DataGridViewComboBoxCell)

            'コンボボックス表示対象外の列の場合は処理を抜ける
            If Not (colcurrent = 3 Or colcurrent = 5 Or colcurrent = 7) Then
                Exit Sub
            End If

            'コンボリスト作成
            cellcombo.Items.Add("")

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvSetubi

            '親項目取得用
            Dim tmp_grp As String = dgv(2, rowcurrent).Value
            Dim tmp_ms As String = dgv(4, rowcurrent).Value

            Select Case colcurrent
                Case 3
                    For cntii = 1 To RelItem_M_setubi_Grp.itemcnt - 1
                        cellcombo.Items.Add(RelItem_M_setubi_Grp.Setubi_grpname(cntii))
                    Next
                Case 5
                    If tmp_grp = "" Then
                        Exit Sub
                    End If
                    For cntii = 1 To RelItem_M_setubi_Ms.itemcnt - 1
                        If tmp_grp = RelItem_M_setubi_Ms.Setubi_grpsortorder(cntii) Then
                            cellcombo.Items.Add(RelItem_M_setubi_Ms.Setubi_name(cntii))
                        End If
                    Next
                Case 7
                    If tmp_ms = "" Then
                        Exit Sub
                    End If
                    For cntii = 1 To RelItem_M_setubi_Komk.itemcnt - 1
                        If tmp_grp = RelItem_M_setubi_Komk.Setubi_grpsortorder(cntii) And tmp_ms = RelItem_M_setubi_Komk.Setubi_sortorder(cntii) Then
                            cellcombo.Items.Add(RelItem_M_setubi_Komk.Komok_name(cntii))
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
            '20160725 部屋設備No入力制御の解除対応 -chg sta
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
                    Case 3
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                        dgv(colcurrent + 4, rowcurrent).Value = ""
                    Case 4
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                    Case 5
                        dgv(colcurrent - 1, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                    Case 6
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    Case 7
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                End Select
                Return rtn
            End If
            '20160725 部屋設備No入力制御の解除対応 -chg end
            ''数値範囲チェック
            'If colcurrent = 2 Then
            '    Dim min As Integer = 1000
            '    Dim max As Integer = 9999
            '    If EtcMethod.Chk_NumRange(currentvalue, min, max) = False Then
            '        rtn = False
            '        Return rtn
            '    End If
            'End If

            '20160525 全体的な動作の修正 -add sta
            'ここから
            '入力された数値チェックを実装する
            '列の読み取り専用プロパティをFalseにすること

            '20160525 全体的な動作の修正 -add end

            '一致する値を取得
            '20160725 部屋設備No入力制御の解除対応 -chg sta
            'Dim searchstr As String = ""
            'Dim searchstr_key_grp As String = ""
            'Dim searchstr_key_ms As String = ""
            'Dim outstr As String = ""
            'Dim chgcol As Integer = colcurrent - 1
            'Dim key_grp As String = dgv(2, rowcurrent).Value
            'Dim key_ms As String = dgv(4, rowcurrent).Value

            'Select Case colcurrent
            '    Case 3
            '        For cntii = 1 To RelItem_M_setubi_Grp.itemcnt - 1
            '            searchstr = RelItem_M_setubi_Grp.Setubi_grpname(cntii)
            '            outstr = RelItem_M_setubi_Grp.Setubi_grpsortorder(cntii)
            '            If currentvalue = searchstr Then
            '                dgv(chgcol, rowcurrent).Value = outstr
            '                Exit For
            '            End If
            '        Next
            '    Case 5
            '        For cntii = 1 To RelItem_M_setubi_Ms.itemcnt - 1
            '            searchstr_key_grp = RelItem_M_setubi_Ms.Setubi_grpsortorder(cntii)
            '            searchstr = RelItem_M_setubi_Ms.Setubi_name(cntii)
            '            outstr = RelItem_M_setubi_Ms.Setubi_sortorder(cntii)
            '            If key_grp = searchstr_key_grp And currentvalue = searchstr Then
            '                dgv(chgcol, rowcurrent).Value = outstr
            '                Exit For
            '            End If
            '        Next
            '    Case 7
            '        For cntii = 1 To RelItem_M_setubi_Komk.itemcnt - 1
            '            searchstr_key_grp = RelItem_M_setubi_Komk.Setubi_grpsortorder(cntii)
            '            searchstr_key_ms = RelItem_M_setubi_Komk.Setubi_sortorder(cntii)
            '            searchstr = RelItem_M_setubi_Komk.Komok_name(cntii)
            '            outstr = RelItem_M_setubi_Komk.Komok_sortorder(cntii)
            '            If key_grp = searchstr_key_grp And key_ms = searchstr_key_ms And currentvalue = searchstr Then
            '                dgv(chgcol, rowcurrent).Value = outstr
            '                Exit For
            '            End If
            '        Next
            'End Select
            Dim searchstr As String = ""
            Dim searchstr_key_grp As String = ""
            Dim searchstr_key_ms As String = ""
            Dim outstr As String = ""
            Dim chgcol As Integer = 0
            Dim key_grp As String = dgv(2, rowcurrent).Value
            Dim key_ms As String = dgv(4, rowcurrent).Value
            Dim matchflg As Boolean = False

            Select Case colcurrent
                Case 2
                    chgcol = colcurrent + 1
                    For cntii = 1 To RelItem_M_setubi_Grp.itemcnt - 1
                        searchstr = RelItem_M_setubi_Grp.Setubi_grpsortorder(cntii)
                        outstr = RelItem_M_setubi_Grp.Setubi_grpname(cntii)
                        If currentvalue = searchstr Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            matchflg = True
                            Exit For
                        End If
                    Next
                    If matchflg = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                        dgv(colcurrent + 4, rowcurrent).Value = ""
                        dgv(colcurrent + 5, rowcurrent).Value = ""
                    End If
                Case 3
                    chgcol = colcurrent - 1
                    For cntii = 1 To RelItem_M_setubi_Grp.itemcnt - 1
                        searchstr = RelItem_M_setubi_Grp.Setubi_grpname(cntii)
                        outstr = RelItem_M_setubi_Grp.Setubi_grpsortorder(cntii)
                        If currentvalue = searchstr Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        End If
                    Next

                Case 4
                    chgcol = colcurrent + 1
                    For cntii = 1 To RelItem_M_setubi_Ms.itemcnt - 1
                        searchstr_key_grp = RelItem_M_setubi_Ms.Setubi_grpsortorder(cntii)
                        searchstr = RelItem_M_setubi_Ms.Setubi_sortorder(cntii)
                        outstr = RelItem_M_setubi_Ms.Setubi_name(cntii)
                        If key_grp = searchstr_key_grp And currentvalue = searchstr Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            matchflg = True
                            Exit For
                        End If
                    Next
                    If matchflg = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                        dgv(colcurrent + 2, rowcurrent).Value = ""
                        dgv(colcurrent + 3, rowcurrent).Value = ""
                    End If
                Case 5
                    chgcol = colcurrent - 1
                    For cntii = 1 To RelItem_M_setubi_Ms.itemcnt - 1
                        searchstr_key_grp = RelItem_M_setubi_Ms.Setubi_grpsortorder(cntii)
                        searchstr = RelItem_M_setubi_Ms.Setubi_name(cntii)
                        outstr = RelItem_M_setubi_Ms.Setubi_sortorder(cntii)
                        If key_grp = searchstr_key_grp And currentvalue = searchstr Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        End If
                    Next

                Case 6
                    chgcol = colcurrent + 1
                    For cntii = 1 To RelItem_M_setubi_Komk.itemcnt - 1
                        searchstr_key_grp = RelItem_M_setubi_Komk.Setubi_grpsortorder(cntii)
                        searchstr_key_ms = RelItem_M_setubi_Komk.Setubi_sortorder(cntii)
                        searchstr = RelItem_M_setubi_Komk.Komok_sortorder(cntii)
                        outstr = RelItem_M_setubi_Komk.Komok_name(cntii)
                        If key_grp = searchstr_key_grp And key_ms = searchstr_key_ms And currentvalue = searchstr Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            matchflg = True
                            Exit For
                        End If
                    Next
                    If matchflg = False Then
                        MsgResult = MessageBox.Show(MSG_ERR_OUTOFRANGE, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        dgv(colcurrent, rowcurrent).Value = ""
                        dgv(colcurrent + 1, rowcurrent).Value = ""
                    End If
                Case 7
                    chgcol = colcurrent - 1
                    For cntii = 1 To RelItem_M_setubi_Komk.itemcnt - 1
                        searchstr_key_grp = RelItem_M_setubi_Komk.Setubi_grpsortorder(cntii)
                        searchstr_key_ms = RelItem_M_setubi_Komk.Setubi_sortorder(cntii)
                        searchstr = RelItem_M_setubi_Komk.Komok_name(cntii)
                        outstr = RelItem_M_setubi_Komk.Komok_sortorder(cntii)
                        If key_grp = searchstr_key_grp And key_ms = searchstr_key_ms And currentvalue = searchstr Then
                            dgv(chgcol, rowcurrent).Value = outstr
                            Exit For
                        End If
                    Next
                Case Else
                    Return rtn
            End Select
            '20160725 部屋設備No入力制御の解除対応 -chg end
            Return rtn

        End Function

        '2016.04.15 設備の重複不可制御をコメントアウト -del sta 
        '        ''' <summary>
        '        ''' 設備の重複チェック
        '        ''' </summary>
        '        ''' <param name="rowno"></param>
        '        ''' <param name="initflg"></param>
        '        ''' <remarks></remarks>
        '        Public Sub Chk_Duplicate(ByVal rowno As Integer, ByVal initflg As Boolean)

        '            Dim dgv As DataGridView = Njc.Frm.RelationFrm.dgvSetubi
        '            Dim key_grp As String = ""
        '            Dim key_ms As String = ""
        '            Dim key_komk As String = ""
        '            Dim key_total As String = ""
        '            Dim key_grpms As String = ""
        '            Dim list_keytotal As New List(Of String)
        '            Dim list_keygrpms As New List(Of String)
        '            Dim list_keytotal_taisyo As New List(Of String)
        '            Dim list_keygrpms_taisyo As New List(Of String)

        '            Dim key_grp_taisyo As String = IIf(dgv(2, rowno).Value Is Nothing, "", dgv(2, rowno).Value)
        '            Dim key_ms_taisyo As String = IIf(dgv(4, rowno).Value Is Nothing, "", dgv(4, rowno).Value)
        '            Dim key_komk_taisyo As String = IIf(dgv(6, rowno).Value Is Nothing, "", dgv(6, rowno).Value)
        '            Dim key_total_taisyo As String = ""
        '            Dim key_grpms_taisyo As String = ""
        '            Dim key_grpms_strcnt As Integer = 0

        '            Dim duplicateflg As Boolean = False

        '            '現在の値を格納しておく (キーがどれか欠けている場合は処理をしない)
        '            If key_grp_taisyo <> "" And key_ms_taisyo <> "" And key_komk_taisyo <> "" Then
        '                key_total_taisyo = key_grp_taisyo & "-" & key_ms_taisyo & "-" & key_komk_taisyo
        '                key_grpms_taisyo = key_grp_taisyo & "-" & key_ms_taisyo
        '                key_grpms_strcnt = key_grpms_taisyo.Length
        '            Else
        '                Exit Sub
        '            End If

        '            '全データ取得
        '            For cntii = 0 To dgv.RowCount - 1

        '                If cntii = rowno Then
        '                    GoTo skiplabel
        '                End If

        '                key_grp = IIf(dgv(2, cntii).Value Is Nothing, "", dgv(2, cntii).Value)
        '                key_ms = IIf(dgv(4, cntii).Value Is Nothing, "", dgv(4, cntii).Value)
        '                key_komk = IIf(dgv(6, cntii).Value Is Nothing, "", dgv(6, cntii).Value)

        '                If key_grp <> "" And key_ms <> "" And key_komk <> "" Then

        '                    key_total = key_grp & "-" & key_ms & "-" & key_komk     '全てのキーを格納
        '                    key_grpms = key_grp & "-" & key_ms                      '設備グループと設備項目のキーを格納

        '                    list_keytotal.Add(key_total)
        '                    list_keygrpms.Add(key_grpms)

        '                End If
        'skiplabel:
        '            Next

        '            'リストから現在の値を削除する
        '            list_keytotal.RemoveAll(Function(s As String) s = key_total_taisyo)

        '            '照合
        '            For Each item In list_keytotal

        '                'リストに含まれているキー(グループ、項目、内容を結合したもの)から現在の項目までのキー文字数で切り出す
        '                Dim tmp_keygrpms As String = item.Substring(0, key_grpms_strcnt)

        '                '現在の項目までのキーと照合し一致した場合は重複エラーとする(グループ、項目が一致し、内容が不一致のデータ)
        '                If tmp_keygrpms = key_grpms_taisyo Then
        '                    duplicateflg = True
        '                    Exit For
        '                End If

        '            Next

        '            If duplicateflg Then
        '                If initflg = False Then
        '                    MsgResult = MessageBox.Show(MSG_DUPLICATE_A, "注意", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '                End If
        '                For cntjj = 4 To dgv.ColumnCount - 1
        '                    dgv(cntjj, rowno).Value = ""
        '                Next
        '            End If

        '        End Sub
        '2016.04.15 設備の重複不可制御をコメントアウト -del end

        ''' <summary>
        ''' セルの読み取り専用制御 '20160525 全体的な動作の修正
        ''' </summary>
        ''' <param name="colcurrent"></param>
        ''' <param name="rowcurrent"></param>
        ''' <remarks></remarks>
        Public Sub Set_CellReadOnly(ByVal colcurrent As Integer, ByVal rowcurrent As Integer)

            'グリッド取得
            Dim dgv As New DataGridView
            dgv = Njc.Frm.RelationFrm.dgvSetubi

            '親項目取得用
            Dim tmp_setubigrpno As String = IIf(dgv(2, rowcurrent).Value Is Nothing, "", dgv(2, rowcurrent).Value)
            Dim tmp_setubigrpname As String = IIf(dgv(3, rowcurrent).Value Is Nothing, "", dgv(3, rowcurrent).Value)

            Dim tmp_setubino As String = IIf(dgv(4, rowcurrent).Value Is Nothing, "", dgv(4, rowcurrent).Value)
            Dim tmp_setubiname As String = IIf(dgv(5, rowcurrent).Value Is Nothing, "", dgv(5, rowcurrent).Value)
            '20160829 エレベーター移行対応 -chg sta
            ''20160725 部屋設備No入力制御の解除対応 -chg sta
            ' ''セルの読み取り制御処理
            ''If tmp_setubigrpno.Trim & tmp_setubigrpname.Trim = "" Then
            ''    'dgv(4, rowcurrent).ReadOnly = True     '読み取り解除の処理は構築中
            ''    dgv(5, rowcurrent).ReadOnly = True
            ''    'dgv(6, rowcurrent).ReadOnly = True     '読み取り解除の処理は構築中
            ''    dgv(7, rowcurrent).ReadOnly = True
            ''ElseIf tmp_setubino.Trim & tmp_setubiname.Trim = "" Then
            ''    'dgv(4, rowcurrent).ReadOnly = False    '読み取り解除の処理は構築中
            ''    dgv(5, rowcurrent).ReadOnly = False
            ''    'dgv(6, rowcurrent).ReadOnly = True     '読み取り解除の処理は構築中
            ''    dgv(7, rowcurrent).ReadOnly = True
            ''Else
            ''    'dgv(4, rowcurrent).ReadOnly = False    '読み取り解除の処理は構築中
            ''    dgv(5, rowcurrent).ReadOnly = False
            ''    'dgv(6, rowcurrent).ReadOnly = False    '読み取り解除の処理は構築中
            ''    dgv(7, rowcurrent).ReadOnly = False
            ''End If

            ' ''読み取り解除の処理は構築中のため固定で読み取り専用にしておく
            ''dgv(2, rowcurrent).ReadOnly = True
            ''dgv(4, rowcurrent).ReadOnly = True
            ''dgv(6, rowcurrent).ReadOnly = True
            ''セルの読み取り制御処理
            'If tmp_setubigrpno.Trim & tmp_setubigrpname.Trim = "" Then
            '    dgv(4, rowcurrent).ReadOnly = True
            '    dgv(5, rowcurrent).ReadOnly = True
            '    dgv(6, rowcurrent).ReadOnly = True
            '    dgv(7, rowcurrent).ReadOnly = True
            'ElseIf tmp_setubino.Trim & tmp_setubiname.Trim = "" Then
            '    dgv(4, rowcurrent).ReadOnly = False
            '    dgv(5, rowcurrent).ReadOnly = False
            '    dgv(6, rowcurrent).ReadOnly = True
            '    dgv(7, rowcurrent).ReadOnly = True
            'Else
            '    dgv(4, rowcurrent).ReadOnly = False
            '    dgv(5, rowcurrent).ReadOnly = False
            '    dgv(6, rowcurrent).ReadOnly = False
            '    dgv(7, rowcurrent).ReadOnly = False
            'End If
            ''20160725 部屋設備No入力制御の解除対応 -chg end
            If tmp_setubigrpno.Trim & tmp_setubigrpname.Trim = "" Then
                dgv(4, rowcurrent).ReadOnly = True
                dgv(5, rowcurrent).ReadOnly = True
                dgv(6, rowcurrent).ReadOnly = True
                dgv(7, rowcurrent).ReadOnly = True
            ElseIf tmp_setubino.Trim & tmp_setubiname.Trim = "" Or tmp_setubigrpno = "999" Or tmp_setubino = "999" Then
                dgv(4, rowcurrent).ReadOnly = False
                dgv(5, rowcurrent).ReadOnly = False
                dgv(6, rowcurrent).ReadOnly = True
                dgv(7, rowcurrent).ReadOnly = True
            Else
                dgv(4, rowcurrent).ReadOnly = False
                dgv(5, rowcurrent).ReadOnly = False
                dgv(6, rowcurrent).ReadOnly = False
                dgv(7, rowcurrent).ReadOnly = False
            End If
            '20160829 エレベーター移行対応 -chg end
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
                    dgv(colcurrent + 2, rowcurrent).Value = ""
                    dgv(colcurrent + 3, rowcurrent).Value = ""
                    dgv(colcurrent + 4, rowcurrent).Value = ""
                    dgv(colcurrent + 5, rowcurrent).Value = ""
                Case 3
                    dgv(colcurrent + 1, rowcurrent).Value = ""
                    dgv(colcurrent + 2, rowcurrent).Value = ""
                    dgv(colcurrent + 3, rowcurrent).Value = ""
                    dgv(colcurrent + 4, rowcurrent).Value = ""
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
            dgv = Njc.Frm.RelationFrm.dgvSetubi
            lbl_relcnt = Njc.Frm.RelationFrm.lblCntSumiSetubi
            lbl_totalcnt = Njc.Frm.RelationFrm.lblCntAllSetubi

            '紐付全件数取得
            totalcnt = dgv.RowCount

            '紐付未設定件数取得
            For cntii = 0 To totalcnt - 1

                Dim tmp_gno As String = dgv(2, cntii).Value
                Dim tmp_gname As String = dgv(3, cntii).Value
                Dim tmp_sno As String = dgv(4, cntii).Value
                Dim tmp_sname As String = dgv(5, cntii).Value
                Dim tmp_kno As String = dgv(6, cntii).Value
                Dim tmp_kname As String = dgv(7, cntii).Value

                If tmp_gno <> "" And tmp_gname <> "" And tmp_sno <> "" And tmp_sname <> "" And tmp_kno <> "" And tmp_kname <> "" Then
                    relcnt = relcnt + 1
                End If
            Next

            '表示
            lbl_relcnt.Text = relcnt.ToString
            lbl_totalcnt.Text = totalcnt.ToString

            '色設定
            Njc.Frm.RelationFrm.lblCntSetubiBack.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntSumiSetubi.ForeColor = Color.Black
            Njc.Frm.RelationFrm.lblCntAllSetubi.ForeColor = Color.Black

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