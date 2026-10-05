Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "入金区分マスタ取得"

    Public Class M_nkbn_Rev7_Repository

        ''' <summary>
        ''' 入金区分取得
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReadMid_Base(ByVal sqlcnnv10 As System.Data.SqlClient.SqlConnection) As Boolean
            '20161021 既存中間ファイルを無くすことによる速度改善処理_全変更 -chg sta
            '            Dim rtn As Boolean = True
            '            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            '            Dim list_relitem As New List(Of String)
            '            Dim relfldno As Integer = 0
            '            Dim filename As String() = New String() {"契約情報", "契約情報", "契約情報", "請求情報"}
            '            Dim sheetname As String() = New String() {"契約履歴情報", "契約入金項目情報", "契約次回入金項目情報", "家主固定控除情報"}
            '            Dim relfldname As String() = New String() {"家賃入金区分", "入金方法", "入金方法", "入金区分"}

            '            For cntfile = 0 To UBound(filename)

            '                '************************
            '                '作業準備
            '                '************************

            '                'Excelファイル初期設定                  
            '                Dim tmp_sql As String = "SELECT DISTINCT " & "[" & relfldname(cntfile) & "]" & " FROM [" & sheetname(cntfile) & "$] WHERE " & "[" & relfldname(cntfile) & "]" & "<> '' "
            '                Dim readtbl As New DataTable()
            '                Dim con_read As New OleDbConnection()

            '                'オープン処理
            '                rtn = excelfile.ExcelFile_ReadOpen(MidDirPath, filename(cntfile), tmp_sql, readtbl, con_read)

            '                'オープン処理失敗時は処理を抜ける
            '                If rtn = False Then
            '                    excelfile.ExcelFile_ReadClose(con_read)
            '                    Return rtn
            '                End If

            '                '行数取得
            '                Dim rowcnt As Integer = readtbl.Rows.Count

            '                'データが存在しない場合は処理を抜ける
            '                If rowcnt = 0 Then
            '                    excelfile.ExcelFile_ReadClose(con_read)
            '                    GoTo skiplbl
            '                End If

            '                '************************
            '                '処理開始
            '                '************************

            '                For cntii As Integer = 0 To rowcnt - 1

            '                    '中断処理
            '                    Application.DoEvents()
            '                    If CancelFlg Then
            '                        Call excelfile.ExcelFile_ReadClose(con_read)
            '                        Return rtn
            '                    End If

            '                    '登録値取得→オブジェクトへ格納
            '                    Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(0)).Trim
            '                    If list_relitem.Contains(fldvalue) = False Then
            '                        list_relitem.Add(fldvalue)
            '                    End If

            '                Next
            'skiplbl:
            '                'クローズ処理
            '                excelfile.ExcelFile_ReadClose(con_read)

            '            Next

            '            '集約した紐付データをモデルへ格納
            '            Dim relcnt As Integer = 1
            '            Dim itemcnt As Integer = list_relitem.Count
            '            Dim model_relitem As New Njc.Model.M_nkbn_Rev7_Model(itemcnt)
            '            For Each relitem In list_relitem
            '                model_relitem.Nkbn_name(relcnt) = relitem
            '                relcnt = relcnt + 1
            '            Next

            '            'モデルの引渡し
            '            RelItem_M_nkbn_Rev7 = Nothing
            '            RelItem_M_nkbn_Rev7 = model_relitem

            '            Return rtn

            Dim rtn As Boolean = True
            Dim list_relitem As New List(Of String)

            '************************
            '作業準備
            '************************
            'データ取得
            Dim tmp_sql As String = ""
            tmp_sql = tmp_sql & " SELECT DISTINCT [家賃入金区分] FROM CVTBL_契約履歴情報 WHERE RTRIM(LTRIM([家賃入金区分])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [入金方法] FROM CVTBL_契約入金項目情報 WHERE RTRIM(LTRIM([入金方法])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [入金方法] FROM CVTBL_契約次回入金項目情報 WHERE RTRIM(LTRIM([入金方法])) <> '' "
            tmp_sql = tmp_sql & " UNION "
            tmp_sql = tmp_sql & " SELECT DISTINCT [入金区分] FROM CVTBL_家主固定控除情報 WHERE RTRIM(LTRIM([入金区分])) <> '' "
            Dim readtbl As New DataTable
            Dim rowcnt As Integer = DBExec.Exec_DataTable(tmp_sql, sqlcnnv10, readtbl, rtn)

            'オープン処理失敗時は処理を抜ける
            If rtn = False Then
                Return rtn
            End If

            'データが存在しない場合は処理を抜ける
            If rowcnt <= 0 Then
                Return rtn
            End If

            '************************
            '処理開始
            '************************

            For cntii As Integer = 0 To rowcnt - 1

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Return rtn
                End If

                '登録値取得→オブジェクトへ格納
                Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(0)).Trim
                If list_relitem.Contains(fldvalue) = False Then
                    list_relitem.Add(fldvalue)
                End If

            Next

            '集約した紐付データをモデルへ格納
            Dim relcnt As Integer = 1
            Dim itemcnt As Integer = list_relitem.Count
            Dim model_relitem As New Njc.Model.M_nkbn_Rev7_Model(itemcnt)
            For Each relitem In list_relitem
                model_relitem.Nkbn_name(relcnt) = relitem
                relcnt = relcnt + 1
            Next

            'モデルの引渡し
            RelItem_M_nkbn_Rev7 = Nothing
            RelItem_M_nkbn_Rev7 = model_relitem

            Return rtn

            '20161021 既存中間ファイルを無くすことによる速度改善処理_全変更 -chg end
        End Function

        ''' <summary>
        ''' V7入金区分取得 '20160829 V7入金区分を革命マスタから取得するように変更
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadMid(ByVal sqlcnnV7 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim readtbl As New DataTable
            Dim reccnt As Integer
            Dim fldname As String
            Dim fldvalue As String
            Dim tmp_kinyuno As String = ""
            Dim tmp_kinyuname As String = ""
            Dim tmp_tenno As String = ""
            Dim tmp_tenname As String = ""
            Dim rtn As Boolean = True

            '画像タイトル情報取得
            Dim tmp_sql As String = " SELECT nkbn_no,nkbn_name FROM m_nkbn ORDER BY nkbn_no "
            reccnt = DBExec.Exec_DataTable(tmp_sql, sqlcnnV7, readtbl)

            'モデル初期化
            Dim model_relitem As New Njc.Model.M_nkbn_Rev7_Model(reccnt)

            '読込開始
            With model_relitem

                For cntii As Integer = 1 To readtbl.Rows.Count

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Then
                        Return rtn
                    End If

                    For cntjj As Integer = 0 To readtbl.Columns.Count - 1

                        '項目名取得
                        fldname = readtbl.Columns(cntjj).ColumnName

                        '登録値取得
                        fldvalue = (readtbl.Rows(cntii - 1).Item(cntjj)).ToString

                        '各項目値→変数格納
                        Select Case fldname
                            Case "nkbn_no"
                                .Nkbn_no(cntii) = fldvalue
                            Case "nkbn_name"
                                .Nkbn_name(cntii) = fldvalue
                        End Select

                    Next

                Next

            End With

            'モデルの引渡し
            RelItem_M_nkbn_Rev7 = Nothing
            RelItem_M_nkbn_Rev7 = model_relitem

            '************************
            '終了処理
            '************************

            Return rtn

        End Function

    End Class


#End Region

End Namespace