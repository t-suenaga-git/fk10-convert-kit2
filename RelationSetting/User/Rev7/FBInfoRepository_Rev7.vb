Imports System.Data.SqlClient
Imports RelationSetting.Njc.N3Lib.Utys
Imports System.Collections.Generic
Imports RelationSetting.Njc.Common
Imports RelationSetting.Njc.Query
Imports Microsoft.Office.Interop    'Excelファイル操作のため追加

Namespace Njc.Repository

#Region "FBフォーマット割付取得"

    Public Class M_FBInfo_Rev7_Repository

        ''' <summary>
        ''' V7FBフォーマット取得
        ''' </summary>
        ''' <remarks></remarks>
        Public Function ReadMid(ByVal sqlcnnV10 As System.Data.SqlClient.SqlConnection) As Boolean

            Dim appli As Excel.Application = Nothing                        'Excelオブジェクト
            Dim wbook As Excel.Workbook = Nothing                           'Excelオブジェクト
            Dim wsheet As Excel.Worksheet = Nothing                         'Excelオブジェクト
            Dim startrow As Integer                                         '書込開始行
            Dim columncnt As Integer                                        '列数
            Dim maxrowcnt As Integer                                        '既存データの行数
            Dim rowcnt As Integer                                           '書込行数
            Dim rtn As Boolean = True
            Dim excelfile As New Njc.Common.ExcelFileManager                'Excelファイル操作用
            Dim list_relitem As New List(Of String)
            Dim relfldno As Integer = 0

            Dim filename As String() = New String() {"自社情報", "自社情報", "自社情報"}
            Dim sheetname As String() = New String() {"振込依頼人情報", "口座振替情報", "入出金取得情報"}

            For cntfile = 0 To UBound(filename)

                '************************
                '作業準備
                '************************

                'Excelファイル初期設定
                rtn = excelfile.Set_ExcelFile_ReadOpen(appli, wbook, wsheet, startrow, columncnt, maxrowcnt, rowcnt, MidDirPath, filename(cntfile), sheetname(cntfile))

                'Excelファイル設定時にエラーが生じた際は処理を抜ける
                If Not rtn Then
                    Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)    '20160525 全体的な動作の修正 -add
                    Return rtn
                End If

                '20160525 全体的な動作の修正 -add sta
                'データが存在しない場合は処理をスキップする
                If rowcnt = 0 Then
                    GoTo skiplbl
                End If
                '20160525 全体的な動作の修正 -add end

                'モデル初期化
                Dim model_relitem As New Object
                Select Case sheetname(cntfile)
                    Case "振込依頼人情報"
                        Dim model_furiirai As New Njc.Model.M_FBInfo_Furiirai_Rev7_Model(rowcnt)
                        model_relitem = model_furiirai
                    Case "口座振替情報"
                        Dim model_kozafurikae As New Njc.Model.M_FBInfo_Kozafurikae_Rev7_Model(rowcnt)
                        model_relitem = model_kozafurikae
                    Case "入出金取得情報"
                        Dim model_nssettting As New Njc.Model.M_FBInfo_Nssetting_Rev7_Model(rowcnt)
                        model_relitem = model_nssettting
                End Select

                '************************
                '処理開始
                '************************

                Dim headerrow As New Object
                Dim relcolvalue As New Object

                'ヘッダー行取得
                headerrow = wsheet.Range(wsheet.Cells(startrow - 1, 1), wsheet.Cells(startrow - 1, columncnt)).Value

                '取得した列のデータをリストへ格納 (重複集約)
                For cntii = 1 To rowcnt

                    '行取得
                    Dim datarowvalue As Object = wsheet.Range(wsheet.Cells(cntii + startrow - 1, 1), wsheet.Cells(cntii + startrow - 1, columncnt)).Value

                    For cntjj = 1 To columncnt

                        Dim fldname As String = headerrow(startrow - 1, cntjj).ToString.Trim
                        Dim fldvalue As String = ""
                        If datarowvalue(startrow - 1, cntjj) IsNot Nothing Then
                            fldvalue = datarowvalue(startrow - 1, cntjj).ToString.Trim
                        End If

                        Select Case sheetname(cntfile)

                            Case "振込依頼人情報"

                                model_relitem.Fmt_syubetu(cntii) = "総合振込"

                                Select Case fldname

                                    Case "振込依頼人No"
                                        model_relitem.Fkom_no(cntii) = fldvalue
                                    Case "振込依頼人名"
                                        model_relitem.Fkom_name(cntii) = fldvalue
                                    Case "金融機関No"
                                        model_relitem.Kinyu_no(cntii) = fldvalue
                                        '20160701 指摘事項まとめファイルの対応 金融機関名称取得修正 -del
                                        'model_relitem.Kinyu_name(cntii) = fldvalue
                                    Case "金融機関支店No"
                                        model_relitem.Ten_no(cntii) = fldvalue
                                        '20160701 指摘事項まとめファイルの対応 金融機関名称取得修正 -del
                                        'model_relitem.Ten_name(cntii) = fldvalue
                                    Case "口座種別"
                                        model_relitem.Kosyu_no(cntii) = fldvalue
                                    Case "口座番号"
                                        model_relitem.Koza_no(cntii) = fldvalue
                                    Case "口座名義"
                                        model_relitem.Koza_meigi(cntii) = fldvalue
                                End Select

                            Case "口座振替情報"

                                model_relitem.Fmt_syubetu(cntii) = "口座振替"

                                Select Case fldname
                                    Case "振替情報No."
                                        model_relitem.Fkae_no(cntii) = fldvalue
                                    Case "振替情報名称"
                                        model_relitem.Fkae_name(cntii) = fldvalue
                                    Case "金融機関No"
                                        model_relitem.Kinyu_no(cntii) = fldvalue
                                        '20160701 指摘事項まとめファイルの対応 金融機関名称取得修正 -del
                                        'model_relitem.Kinyu_name(cntii) = fldvalue
                                    Case "金融機関支店No"
                                        model_relitem.Ten_no(cntii) = fldvalue
                                        '20160701 指摘事項まとめファイルの対応 金融機関名称取得修正 -del
                                        'model_relitem.Ten_name(cntii) = fldvalue
                                    Case "口座種別"
                                        model_relitem.Kosyu_no(cntii) = fldvalue
                                    Case "口座番号"
                                        model_relitem.Koza_no(cntii) = fldvalue
                                    Case "口座名義"
                                        model_relitem.Koza_meigi(cntii) = fldvalue
                                End Select

                            Case "入出金取得情報"

                                model_relitem.Fmt_syubetu(cntii) = "入出金"

                                Select Case fldname
                                    Case "入出金取得No"
                                        model_relitem.Ns_no(cntii) = fldvalue
                                    Case "設定名称"
                                        model_relitem.Ns_name(cntii) = fldvalue
                                End Select

                        End Select

                    Next

                    '20160701 指摘事項まとめファイルの対応 金融機関名称取得修正 -add sta
                    '金融機関名称設定
                    If sheetname(cntfile) <> "入出金取得情報" Then
                        Dim errstr As String = ""       'ログは本体側で出力するためダミーでセットしておく
                        Dim tmp_kinyuname As String = ""
                        Dim tmp_tenname As String = ""
                        Call DataChk.Chk_DataMstExist_Kinyu(sqlcnnV10, model_relitem.Kinyu_no(cntii), model_relitem.Kinyu_no(cntii) & "-" & model_relitem.Ten_no(cntii), tmp_kinyuname, tmp_tenname, errstr)
                        model_relitem.Kinyu_name(cntii) = tmp_kinyuname
                        model_relitem.Ten_name(cntii) = tmp_tenname
                    End If
                    '20160701 指摘事項まとめファイルの対応 金融機関名称取得修正 -add end

                Next

                '************************
                'モデルの引渡し
                '************************
                Select Case sheetname(cntfile)
                    Case "振込依頼人情報"
                        RelItem_M_FBInfo_Furiirai_Rev7 = Nothing
                        RelItem_M_FBInfo_Furiirai_Rev7 = model_relitem
                    Case "口座振替情報"
                        RelItem_M_FBInfo_Kozafurikae_Rev7 = Nothing
                        RelItem_M_FBInfo_Kozafurikae_Rev7 = model_relitem
                    Case "入出金取得情報"
                        RelItem_M_FBInfo_Nssetting_Rev7 = Nothing
                        RelItem_M_FBInfo_Nssetting_Rev7 = model_relitem
                End Select

                '************************
                '終了処理
                '************************

skiplbl:        '20160525 全体的な動作の修正 -add sta

                'Excelファイル終了設定
                Call excelfile.Set_ExcelFile_ReadClose(appli, wbook, wsheet)

            Next

            Return rtn

        End Function

    End Class

#End Region

End Namespace