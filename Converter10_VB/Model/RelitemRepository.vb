Imports System.Data.SqlClient
Imports Converter10.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加
Imports Converter10.Njc.Common

Namespace Njc.Repository

#Region "物件分類マスタ"

    Public Class M_bk_rui_Repository

        ''' <summary>
        ''' 物件分類マスタ移行処理
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_RelItem(ByVal sqlcnnv10 As SqlConnection, ByVal sheetname As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim tmp_cvcnt As Integer                                        '移行件数格納
            Dim tmp_condcnt As Integer                                      '調整件数格納

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用

            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            Dim rtn As Boolean = True                                       '戻り値
            Dim model_cvitem As New Njc.Model.M_bk_rui_Model                  '移行値格納用モデル初期化

            '移行対象のテーブル名、フィールド名を取得
            Dim tblname As String = "m_bk_rui"
            Dim fldnamegrp As String = "bk_ruino,bk_ruiname,bk_ruibiko,bk_ruiuseflg,sincyoku_keiyakukbn," & _
                                       "sincyoku_kosinkbn,sincyoku_kaiyakukbn,history,homemate_ruikbn,bk_endofmonthflg," & _
                                       "ikkatukariage_siwakekbn"

            '既存データ取得
            Dim list_existdata As New List(Of String)
            Dim tmp_sql As String = " SELECT bk_ruino FROM " & tblname
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)
            If rtn = False Then
                Return rtn
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return rtn
            End If

            '----------------------------
            '紐付データ取得
            '----------------------------

            'ヘッダー行取得
            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            With model_cvitem

                'データ部処理
                For cntii = 1 To rowcnt

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Then
                        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                        Return rtn
                    End If

                    '行取得
                    Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                    '移行値取得
                    For cntjj = 1 To columncnt

                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                        Dim fldvalue As String = ""
                        If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                            fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                        End If

                        Select Case fldname
                            Case "賃貸革命10物件分類No"
                                .Vari_Bk_ruino = fldvalue
                            Case "賃貸革命10物件分類名称"
                                .Vari_Bk_ruiname = fldvalue
                        End Select

                    Next

                    '固定値
                    .Vari_Bk_ruibiko = ""
                    .Vari_Bk_ruiuseflg = 1
                    .Vari_Sincyoku_keiyakukbn = 0
                    .Vari_Sincyoku_kosinkbn = 0
                    .Vari_Sincyoku_kaiyakukbn = 0
                    .Vari_Homemate_ruikbn = ""
                    .Vari_Bk_endofmonthflg = ""
                    .Vari_Ikkatukariage_siwakekbn = ""
                    .Vari_History = DefHistory

                    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    Dim hash_cvitem As New Hashtable
                    hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                    '挿入処理
                    If .Vari_Bk_ruino <> "" AndAlso list_existdata.Contains(.Vari_Bk_ruino) = False Then
                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)
                    End If

                Next

            End With

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

    End Class

#End Region

#Region "部屋分類マスタ"

    Public Class M_hy_rui_Repository

        ''' <summary>
        ''' 部屋分類マスタ移行処理
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_RelItem(ByVal sqlcnnv10 As SqlConnection, ByVal sheetname As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim tmp_cvcnt As Integer                                        '移行件数格納
            Dim tmp_condcnt As Integer                                      '調整件数格納

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用

            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            Dim rtn As Boolean = True                                       '戻り値
            Dim model_cvitem As New Njc.Model.M_hy_rui_Model                  '移行値格納用モデル初期化

            '移行対象のテーブル名、フィールド名を取得
            Dim tblname As String = "m_hy_rui"
            'hy_ruisortorderは自動で挿入されるため対象外
            Dim fldnamegrp As String = "hy_ruino,hy_ruiname,hy_ruisyubetu,hy_ruiuseflg,hy_carportflg," & _
                                       "history"

            '既存データ取得
            Dim list_existdata As New List(Of String)
            Dim tmp_sql As String = " SELECT hy_ruino FROM " & tblname
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)
            If rtn = False Then
                Return rtn
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return rtn
            End If

            '----------------------------
            '紐付データ取得
            '----------------------------

            'ヘッダー行取得
            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            With model_cvitem

                'データ部処理
                For cntii = 1 To rowcnt

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Then
                        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                        Return rtn
                    End If

                    '行取得
                    Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                    '移行値取得
                    For cntjj = 1 To columncnt

                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                        Dim fldvalue As String = ""
                        If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                            fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                        End If

                        Select Case fldname
                            Case "賃貸革命10部屋分類No"
                                .Vari_Hy_ruino = fldvalue
                                '.Vari_Hy_ruisortorder = fldvalue    'hy_ruisortorderは自動で挿入されるため対象外
                            Case "賃貸革命10部屋分類名称"
                                .Vari_Hy_ruiname = fldvalue
                        End Select

                    Next

                    '固定値
                    .Vari_Hy_ruisyubetu = 1         '紐付けに追加する必要あり
                    .Vari_Hy_ruiuseflg = 1
                    .Vari_Hy_carportflg = 1         '紐付けに追加する必要あり
                    .Vari_History = DefHistory

                    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    Dim hash_cvitem As New Hashtable
                    hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                    '挿入処理
                    If .Vari_Hy_ruino <> "" AndAlso list_existdata.Contains(.Vari_Hy_ruino) = False Then
                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)
                    End If

                Next

            End With

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

    End Class

#End Region

#Region "入金区分マスタ"

    Public Class M_nkbn_Repository

        ''' <summary>
        ''' 入金区分マスタ移行処理
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_RelItem(ByVal sqlcnnv10 As SqlConnection, ByVal sheetname As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim tmp_cvcnt As Integer                                        '移行件数格納
            Dim tmp_condcnt As Integer                                      '調整件数格納

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用

            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            Dim rtn As Boolean = True                                       '戻り値
            Dim model_cvitem As New Njc.Model.M_nkbn_Model                  '移行値格納用モデル初期化

            '移行対象のテーブル名、フィールド名を取得
            Dim tblname As String = "m_nkbn"
            Dim fldnamegrp As String = "nkbn_no,nkbn_order,nkbn_name,nkbn_shortname,nkbn_zokusei," & _
                                       "history"

            '既存データ取得
            Dim list_existdata As New List(Of String)
            Dim tmp_sql As String = " SELECT nkbn_no FROM " & tblname
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)
            If rtn = False Then
                Return rtn
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return rtn
            End If

            '----------------------------
            '紐付データ取得
            '----------------------------

            'ヘッダー行取得
            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            With model_cvitem

                'データ部処理
                For cntii = 1 To rowcnt

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Then
                        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                        Return rtn
                    End If

                    '行取得
                    Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                    '移行値取得
                    For cntjj = 1 To columncnt

                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                        Dim fldvalue As String = ""
                        If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                            fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                        End If

                        Select Case fldname
                            Case "賃貸革命10入金区分No"
                                .Vari_Nkbn_no = fldvalue
                                .Vari_Nkbn_order = fldvalue
                            Case "賃貸革命10入金区分名称"
                                .Vari_Nkbn_name = fldvalue
                            Case "賃貸革命10入金区分属性No"
                                .Vari_Nkbn_zokusei = fldvalue
                        End Select

                    Next

                    '固定値
                    .Vari_Nkbn_shortname = ""
                    .Vari_History = DefHistory

                    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    Dim hash_cvitem As New Hashtable
                    hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                    '挿入処理
                    If .Vari_Nkbn_no <> "" AndAlso list_existdata.Contains(.Vari_Nkbn_no) = False Then
                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)
                    End If

                Next

            End With

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

    End Class

#End Region

#Region "契約分類マスタ"

    Public Class M_ky_rui_Repository

        ''' <summary>
        ''' 契約分類マスタ移行処理
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_RelItem(ByVal sqlcnnv10 As SqlConnection, ByVal sheetname As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim tmp_cvcnt As Integer                                        '移行件数格納
            Dim tmp_condcnt As Integer                                      '調整件数格納

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用

            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            Dim rtn As Boolean = True                                       '戻り値
            Dim model_cvitem As New Njc.Model.M_ky_rui_Model                '移行値格納用モデル初期化

            '移行対象のテーブル名、フィールド名を取得
            Dim tblname As String = "m_ky_rui"
            Dim fldnamegrp As String = "ky_ruino,ky_ruiname,ky_ruibiko,ky_ruitutimm,ky_ruicolor," & _
                                       "teisyaku_flg,history,useflg"

            '既存データ取得
            Dim list_existdata As New List(Of String)
            Dim tmp_sql As String = " SELECT ky_ruino FROM " & tblname
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)
            If rtn = False Then
                Return rtn
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return rtn
            End If

            '----------------------------
            '紐付データ取得
            '----------------------------

            'ヘッダー行取得
            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            With model_cvitem

                'データ部処理
                For cntii = 1 To rowcnt

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Then
                        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                        Return rtn
                    End If

                    '行取得
                    Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                    '作業用変数
                    Dim tmp_syakuyakbn As String = ""

                    '移行値取得
                    For cntjj = 1 To columncnt

                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                        Dim fldvalue As String = ""
                        If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                            fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                        End If

                        Select Case fldname
                            Case "賃貸革命10契約分類No"
                                .Vari_Ky_ruino = fldvalue
                            Case "賃貸革命10契約分類名称"
                                .Vari_Ky_ruiname = fldvalue
                            Case "定期借家として扱う"
                                tmp_syakuyakbn = IIf(fldvalue = "True", 1, 2)
                        End Select

                    Next

                    '定期借家区分の判定
                    If .Vari_Ky_ruino = "" Or .Vari_Ky_ruino = "" Then
                        .Vari_Teisyaku_flg = 2
                    Else
                        .Vari_Teisyaku_flg = tmp_syakuyakbn
                    End If

                    '固定値
                    .Vari_Ky_ruibiko = ""
                    .Vari_Ky_ruitutimm = ""
                    .Vari_Ky_ruicolor = "-16777216"
                    .Vari_Useflg = 1
                    .Vari_History = DefHistory

                    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    Dim hash_cvitem As New Hashtable
                    hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                    '挿入処理
                    If .Vari_Ky_ruino <> "" AndAlso list_existdata.Contains(.Vari_Ky_ruino) = False Then
                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)
                    End If

                Next

            End With

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

    End Class

#End Region

#Region "入金項目マスタ"

    Public Class M_nkin_Repository

        ''' <summary>
        ''' 入金項目マスタ移行処理
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_RelItem(ByVal sqlcnnv10 As SqlConnection, ByVal sheetname As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim tmp_cvcnt As Integer                                        '移行件数格納
            Dim tmp_condcnt As Integer                                      '調整件数格納

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用

            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            Dim rtn As Boolean = True                                       '戻り値
            Dim model_cvitem As New Njc.Model.M_nkin_Model                  '移行値格納用モデル初期化

            '移行対象のテーブル名、フィールド名を取得
            Dim tblname As String = "m_nkin"
            Dim fldnamegrp As String = "nkin_no,nkin_name,nkin_printname,nkin_ruino,nkin_zkseino," & _
                                       "nkin_useflg,keiyaku_nkinno,keiyakuhiki_nkinno,kosin_nkinno,kaiyaku_nkinno," & _
                                       "kaiyakuhiki_nkinno,ryosyu_flg,biko_nkin,history"

            '既存データ取得
            Dim list_existdata As New List(Of String)
            Dim tmp_sql As String = " SELECT nkin_no FROM " & tblname
            rtn = DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_existdata)
            If rtn = False Then
                Return rtn
            End If

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return rtn
            End If

            '----------------------------
            '紐付データ取得
            '----------------------------

            'ヘッダー行取得
            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            With model_cvitem

                'データ部処理
                For cntii = 1 To rowcnt

                    '中断処理
                    Application.DoEvents()
                    If CancelFlg Then
                        Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                        Return rtn
                    End If

                    '行取得
                    Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                    '移行値取得
                    For cntjj = 1 To columncnt

                        Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                        Dim fldvalue As String = ""
                        If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                            fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                        End If

                        Select Case fldname
                            Case "賃貸革命10入金項目No"
                                .Vari_Nkin_no = fldvalue
                            Case "賃貸革命10入金項目名称"
                                .Vari_Nkin_name = fldvalue
                            Case "移行元入金項目区分"
                                .Vari_Nkin_ruino = Left(fldvalue, 1)
                            Case "賃貸革命10入金項目属性No"
                                .Vari_Nkin_zkseino = fldvalue
                        End Select

                    Next

                    '固定値
                    .Vari_Nkin_printname = ""
                    .Vari_Nkin_useflg = 1
                    .Vari_Keiyaku_nkinno = 0
                    .Vari_Keiyakuhiki_nkinno = 0
                    .Vari_Kosin_nkinno = ""
                    .Vari_Kaiyaku_nkinno = ""
                    .Vari_Kaiyakuhiki_nkinno = 0
                    .Vari_Ryosyu_flg = 2
                    .Vari_Biko_nkin = ""
                    .Vari_History = DefHistory

                    'フィールド名と移行値を紐付→作業用ハッシュテーブル格納
                    Dim hash_cvitem As New Hashtable
                    hash_cvitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

                    '挿入処理
                    If .Vari_Nkin_no <> "" AndAlso list_existdata.Contains(.Vari_Nkin_no) = False Then
                        Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_cvitem, normalflg)
                    End If

                Next

            End With

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

    End Class

#End Region

#Region "設備マスタ" '20160829 設備の新規挿入処理を追加 -add

    Public Class M_setubi_rel_Repository

        ''' <summary>
        ''' 設備マスタ移行処理
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="sheetname"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Set_RelItem(ByVal sqlcnnv10 As SqlConnection, ByVal sheetname As String) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim tmp_cvcnt As Integer                                        '移行件数格納
            Dim tmp_condcnt As Integer                                      '調整件数格納

            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim obj_pgb As New Njc.Common.ProgressBarManager                'プログレスバー設定用
            Dim obj_com As New Njc.Common.CommonRepository                  '共通処理用

            Dim normalflg As Boolean                                        'INSERT正常終了フラグ
            Dim rtn As Boolean = True                                       '戻り値

            'Excelファイル初期設定
            rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, RelationDirPath, MIDFILE_RELNAME, sheetname)

            'Excelファイル設定時にエラーが生じた際は処理を抜ける
            If Not rtn Then
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                Return rtn
            End If

            '革命10へ設備グループマスタを移行する用意
            Dim grpinsertflg As Boolean = True
            Dim tmp_grpguid As String = ""
            Dim tmp_grpno As String = ""

            '----------------------------
            '紐付データ取得
            '----------------------------

            'ヘッダー行取得
            Dim headervalue As Object = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

            'データ部処理
            For cntii = 1 To rowcnt

                '中断処理
                Application.DoEvents()
                If CancelFlg Then
                    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)
                    Return rtn
                End If

                '行取得
                Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                '作業用変数
                Dim tmp_setubiname As String = ""
                Dim tmp_komkname As String = ""
                Dim tmp_grpname10 As String = ""
                Dim tmp_komokno10 As String = ""

                '移行値取得
                For cntjj = 1 To columncnt

                    Dim fldname As String = headervalue(startrow - 1, cntjj).ToString.Trim
                    Dim fldvalue As String = ""
                    If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                        fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                    End If

                    Select Case fldname
                        Case "移行元設備グループ名称"
                            tmp_setubiname = fldvalue
                        Case "移行元設備名称"
                            tmp_komkname = fldvalue
                        Case "賃貸革命10設備グループNo"
                            tmp_grpname10 = fldvalue
                        Case "賃貸革命10設備グループ名称"

                        Case "賃貸革命10設備No"

                        Case "賃貸革命10設備名称"

                        Case "賃貸革命10項目No"
                            tmp_komokno10 = fldvalue
                        Case "賃貸革命10項目名称"

                    End Select

                Next

                '新規移行対象確認
                '※グループNo = 999はエレベーターなので項目は存在しない
                If tmp_grpname10 <> "999" And tmp_komokno10 = "" Then

                    '新規設備作成開始
                    '設備グループ作成
                    If grpinsertflg Then
                        Call Me.Set_SetubiGrp(sqlcnnv10, tmp_grpguid)
                        grpinsertflg = False
                    End If

                    '設備作成
                    Dim setubiguid As String = ""
                    Call Set_Setubi(sqlcnnv10, tmp_grpguid, tmp_setubiname, setubiguid)

                    '設備項目作成
                    Call Set_Setubi_Komk(sqlcnnv10, setubiguid, tmp_komkname)

                End If

            Next

            'Excelファイル終了設定
            Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Return rtn

        End Function

        ''' <summary>
        ''' 設備グループマスタの作成(最初に1回のみ実行する)
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="grpguid"></param>
        ''' <param name="grpno"></param>
        ''' <remarks></remarks>
        Public Sub Set_SetubiGrp(ByVal sqlcnnv10 As SqlConnection, ByRef grpguid As String)

            Dim normalflg As Boolean = True
            Dim model_cvitem As New Njc.Model.M_setubi_rel_grp_Model
            Dim tblname As String = "m_setubi_grp"
            Dim fldnamegrp As String = "setubi_grpguid,setubi_grpsortorder,setubi_grpname,setubi_grpuseflg,setubi_grpsyskbn," & _
                                       "history"

            '設備グループマスタ有無確認
            Dim tmp_sqlexist As String = " SELECT CONVERT(VARCHAR(MAX),setubi_grpguid) FROM m_setubi_grp WHERE setubi_grpname = 'ユーザー作成設備' "
            Dim exist_setubigrpguid As String = DBExec.Exec_Scalar(tmp_sqlexist, sqlcnnv10)

            '存在する場合は戻り値用に設備guidを格納して処理を抜ける
            If exist_setubigrpguid <> "" Then
                grpguid = exist_setubigrpguid
                Exit Sub
            End If

            '設備グループマスタ作成準備
            Dim new_grpguid As String = Guid.NewGuid.ToString
            Dim tmp_sqlgetgrpno As String = " SELECT CONVERT(VARCHAR,MAX(setubi_grpsortorder) + 1)  FROM m_setubi_grp "
            Dim tmp_grpno As String = DBExec.Exec_Scalar(tmp_sqlgetgrpno, sqlcnnv10)

            'モデルへ値を格納
            Dim defint As String = "1"
            With model_cvitem
                .Vari_Setubi_grpguid = new_grpguid
                .Vari_Setubi_grpsortorder = tmp_grpno
                .Vari_Setubi_grpname = "ユーザー作成設備"
                .Vari_Setubi_grpuseflg = defint
                .Vari_Setubi_grpsyskbn = defint
                .Vari_History = DefHistory
            End With

            'ハッシュテーブルへ格納
            Dim hash_setubigrpitem As New Hashtable
            hash_setubigrpitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '挿入処理実行
            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_setubigrpitem, normalflg)

            '戻り値用に値を格納
            grpguid = new_grpguid

        End Sub

        ''' <summary>
        ''' 設備マスタの作成(同一名称のものは同一と見なし、統合する)
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="grpguid"></param>
        ''' <param name="grpno"></param>
        ''' <remarks></remarks>
        Public Sub Set_Setubi(ByVal sqlcnnv10 As SqlConnection, ByVal grpguid As String, ByVal setubiname As String, ByRef setubiguid As String)

            Dim normalflg As Boolean = True
            Dim model_cvitem As New Njc.Model.M_setubi_rel_Model
            Dim tblname As String = "m_setubi"
            Dim fldnamegrp As String = "setubi_guid,setubi_grpguid,setubi_sortorder,setubi_code,setubi_name," & _
                                       "setubi_useflg,setubi_hyuseflg,setubi_disp1name,setubi_disp2name,setubi_disp3name," & _
                                       "history,jyuyojiko_flg"


            '設備マスタ有無確認
            Dim tmp_sqlexist As String = ""
            tmp_sqlexist = tmp_sqlexist & " SELECT CONVERT(VARCHAR(MAX),setubi_guid) FROM m_setubi "
            tmp_sqlexist = tmp_sqlexist & " WHERE setubi_grpguid = '" & grpguid & "' "
            tmp_sqlexist = tmp_sqlexist & " AND setubi_name = '" & setubiname & "' "
            Dim exist_setubiguid As String = DBExec.Exec_Scalar(tmp_sqlexist, sqlcnnv10)

            '存在する場合は戻り値用に設備guidを格納して処理を抜ける
            If exist_setubiguid <> "" Then
                setubiguid = exist_setubiguid
                Exit Sub
            End If

            '設備マスタの新規作成
            Dim new_setubiguid As String = Guid.NewGuid.ToString
            Dim tmp_sqlgetsetubino As String = ""
            tmp_sqlgetsetubino = tmp_sqlgetsetubino & " SELECT CONVERT(VARCHAR,MAX(setubi_sortorder) + 1) FROM m_setubi "
            tmp_sqlgetsetubino = tmp_sqlgetsetubino & " WHERE setubi_grpguid = '" & grpguid & "' "
            Dim tmp_setubino As String = DBExec.Exec_Scalar(tmp_sqlgetsetubino, sqlcnnv10)
            If tmp_setubino = "0" Then
                tmp_setubino = 1
            End If

            '挿入クエリ作成
            Dim defint1 As String = "1"
            Dim defint2 As String = "0"
            With model_cvitem
                .Vari_Setubi_guid = new_setubiguid
                .Vari_Setubi_grpguid = grpguid
                .Vari_Setubi_sortorder = tmp_setubino
                .Vari_Setubi_code = ""
                .Vari_Setubi_name = setubiname
                .Vari_Setubi_useflg = defint1
                .Vari_Setubi_hyuseflg = defint1
                .Vari_Setubi_disp1name = ""
                .Vari_Setubi_disp2name = ""
                .Vari_Setubi_disp3name = ""
                .Vari_History = DefHistory
                .Vari_Jyuyojiko_flg = defint2
            End With

            'ハッシュテーブルへ格納
            Dim hash_setubiitem As New Hashtable
            hash_setubiitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '挿入処理実行
            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_setubiitem, normalflg)

            '戻り値用に値を格納
            setubiguid = new_setubiguid

        End Sub

        ''' <summary>
        ''' 設備項目マスタの作成(同一名称のものは同一と見なす)
        ''' </summary>
        ''' <param name="sqlcnnv10"></param>
        ''' <param name="grpguid"></param>
        ''' <param name="grpno"></param>
        ''' <remarks></remarks>
        Public Sub Set_Setubi_Komk(ByVal sqlcnnv10 As SqlConnection, ByVal setubiguid As String, ByVal komkname As String)

            Dim normalflg As Boolean = True
            Dim model_cvitem As New Njc.Model.M_setubi_rel_komk_Model
            Dim tblname As String = "m_setubi_lst"
            Dim fldnamegrp As String = "komok_guid,setubi_guid,komok_sortorder,komok_name,komok_code1," & _
                                       "komok_code2,komok_code3,default_flg,disp1name,disp2name," & _
                                       "disp3name,disp1iconguid,disp2iconguid,disp3iconguid,history"

            '設備項目マスタ有無確認(紐付時に重複は除去しているが念の為)
            Dim tmp_sqlexist As String = ""
            tmp_sqlexist = tmp_sqlexist & " SELECT CONVERT(VARCHAR(MAX),komok_guid) FROM m_setubi_lst "
            tmp_sqlexist = tmp_sqlexist & " WHERE setubi_guid = '" & setubiguid & "' "
            tmp_sqlexist = tmp_sqlexist & " AND komok_name = '" & komkname & "' "
            Dim exist_komkguid As String = DBExec.Exec_Scalar(tmp_sqlexist, sqlcnnv10)

            '存在する場合は処理を抜ける
            If exist_komkguid <> "" Then
                Exit Sub
            End If

            '設備項目マスタの新規作成
            Dim new_komkguid As String = Guid.NewGuid.ToString
            Dim tmp_sqlkomkno As String = ""
            tmp_sqlkomkno = tmp_sqlkomkno & " SELECT CONVERT(VARCHAR,MAX(komok_sortorder) + 1) FROM m_setubi_lst "
            tmp_sqlkomkno = tmp_sqlkomkno & " WHERE setubi_guid = '" & setubiguid & "' "
            Dim tmp_komkno As String = DBExec.Exec_Scalar(tmp_sqlkomkno, sqlcnnv10)
            If tmp_komkno = "0" Then
                tmp_komkno = 1
            End If

            '挿入クエリ作成
            Dim defint As String = "0"
            Dim defstr As String = "00000000-0000-0000-0000-000000000000"
            With model_cvitem
                .Vari_Komok_guid = new_komkguid
                .Vari_Setubi_guid = setubiguid
                .Vari_Komok_sortorder = tmp_komkno
                .Vari_Komok_name = komkname
                .Vari_Komok_code1 = ""
                .Vari_Komok_code2 = ""
                .Vari_Komok_code3 = ""
                .Vari_Default_flg = defint
                .Vari_Disp1name = ""
                .Vari_Disp2name = ""
                .Vari_Disp3name = ""
                .Vari_Disp1iconguid = defstr
                .Vari_Disp2iconguid = defstr
                .Vari_Disp3iconguid = defstr
                .Vari_History = DefHistory
            End With

            'ハッシュテーブルへ格納
            Dim hash_setubikomkitem As New Hashtable
            hash_setubikomkitem = GetHashFldToValue.Get_Hash_fldvalue(fldnamegrp, model_cvitem)

            '挿入処理実行
            Call CVDBInsert.Cnv_Db(sqlcnnv10, tblname, fldnamegrp, hash_setubikomkitem, normalflg)

        End Sub

    End Class

#End Region

End Namespace


