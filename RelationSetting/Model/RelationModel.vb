Namespace Njc.Model

#Region "接続情報"

    Public Class DefSQLConnection

        Public Property ServerName As String                                    '接続サーバー名
        Public Property InitialCatalog As String                                'カタログ名
        Public Property NetworkLibrary As Integer                               '接続で使用するネットワークライブラリの種類
        Public Property User As String                                          '認証で使用するユーザ名
        Public Property Pass As String                                          '認証で使用するパスワード
        Public Property TimeOut As String                                       'DB接続時タイムアウト設定値(SEC)

        Public Sub New()

            Me.ServerName = ""
            Me.InitialCatalog = ""
            Me.User = ""
            Me.Pass = ""
            Me.TimeOut = "30"

        End Sub

    End Class

#End Region

#Region "テーブル名(日本語名)"

    ''' <summary>
    ''' テーブル名称設定<br/>
    ''' </summary>
    Public Class TblName
        '(追加箇所)
        Public Property Bkrui As String
        Public Property Hyrui As String
        Public Property Nkinkbn As String
        Public Property Toritaiyo As String
        Public Property Nkinkomk As String
        Public Property Kozo As String
        Public Property Kozasyu As String
        Public Property Setubi As String

        Public Sub New()
            '(追加箇所)
            Me.Bkrui = "物件分類マスタ"
            Me.Hyrui = "部屋分類マスタ"
            Me.Nkinkbn = "入金区分マスタ"
            Me.Toritaiyo = "取引態様マスタ"
            Me.Nkinkomk = "入金項目マスタ"
            Me.Kozo = "構造マスタ"
            Me.Kozasyu = "口座種別マスタ"
            Me.Setubi = "設備マスタ"

        End Sub

    End Class

#End Region

#Region "仮テーブル名(アルファベット)"

    ''' <summary>
    ''' 仮テーブル名称設定<br/>
    ''' </summary>
    Public Class TmpTableName

        '(追加箇所)
        Public Property TmpBkrui As String
        Public Property TmpHyrui As String
        Public Property TmpNkinkbn As String
        Public Property TmpToritaiyo As String
        Public Property TmpKozo As String
        Public Property TmpKozasyu As String
        Public Property TmpNkinkomk As String
        Public Property TmpSetubi As String

        '(追加箇所)
        Public Sub New()

            Me.TmpBkrui = "tmp_bkrui_mst"
            Me.TmpHyrui = "tmp_hyrui_mst"
            Me.TmpNkinkbn = "tmp_nkbn_mst"
            Me.TmpToritaiyo = "tmp_toritaiyo_mst"
            Me.TmpKozo = "tmp_kozo_mst"
            Me.TmpKozasyu = "tmp_kozasyu_mst"
            Me.TmpNkinkomk = "tmp_nkin_mst"
            Me.TmpSetubi = "tmp_setubi_mst"

        End Sub

    End Class

#End Region

#Region "各紐付項目関連名"

    Public Class RelItem_Bkrui

        Public Property bkruino As String

        Public Sub New()

            bkruino = "物件分類No"

        End Sub

    End Class


#End Region

#Region "各紐付項目フィールド名"

    '(追加箇所)

    ''' <summary>
    ''' 物件分類フィールド名称設定<br/>
    ''' </summary>
    Public Class FieldName_Bkrui

        Public Property FieldName() As Njc.Common.SetArrayData
        Public Property fieldcnt As Integer

        Public Sub New()

            fieldcnt = 4
            Me.FieldName = New Njc.Common.SetArrayData(fieldcnt)
            Me.FieldName(0) = REL_FROM_NAME & "物件分類№"
            Me.FieldName(1) = REL_FROM_NAME & "物件分類名称"
            Me.FieldName(2) = REL_FROM_NAME & "物件分類備考"
            Me.FieldName(3) = REL_TO_NAME & "物件分類№"
            Me.FieldName(4) = REL_TO_NAME & "物件分類名称"
        End Sub

    End Class

    ''' <summary>
    ''' 部屋分類フィールド名称設定<br/>
    ''' </summary>
    Public Class FieldName_Hyui

        Public Property FieldName() As Njc.Common.SetArrayData
        Public Property fieldcnt As Integer

        Public Sub New()

            fieldcnt = 3
            Me.FieldName = New Njc.Common.SetArrayData(fieldcnt)
            Me.FieldName(0) = REL_FROM_NAME & "部屋分類№"
            Me.FieldName(1) = REL_FROM_NAME & "部屋分類名称"
            Me.FieldName(2) = REL_TO_NAME & "部屋分類№"
            Me.FieldName(3) = REL_TO_NAME & "部屋分類名称"

        End Sub

    End Class

    ''' <summary>
    ''' 入金区分フィールド名称設定<br/>
    ''' </summary>
    Public Class FieldName_Nkinkbn

        Public Property FieldName() As Njc.Common.SetArrayData
        Public Property fieldcnt As Integer

        Public Sub New()

            fieldcnt = 5
            Me.FieldName = New Njc.Common.SetArrayData(fieldcnt)
            Me.FieldName(0) = REL_FROM_NAME & "入金区分№"
            Me.FieldName(1) = REL_FROM_NAME & "入金区分名称"
            Me.FieldName(2) = REL_TO_NAME & "入金区分№"
            Me.FieldName(3) = REL_TO_NAME & "入金区分名称"
            Me.FieldName(4) = REL_TO_NAME & "入金区分属性№"
            Me.FieldName(5) = REL_TO_NAME & "入金区分属性"

        End Sub

    End Class

    ''' <summary>
    ''' 取引態様フィールド名称設定<br/>
    ''' </summary>
    Public Class FieldName_Toritaiyo

        Public Property FieldName() As Njc.Common.SetArrayData
        Public Property fieldcnt As Integer

        Public Sub New()

            fieldcnt = 3
            Me.FieldName = New Njc.Common.SetArrayData(fieldcnt)
            Me.FieldName(0) = REL_FROM_NAME & "取引態様№"
            Me.FieldName(1) = REL_FROM_NAME & "取引態様名称"
            Me.FieldName(2) = REL_TO_NAME & "取引態様№"
            Me.FieldName(3) = REL_TO_NAME & "取引態様名称"

        End Sub

    End Class

    ''' <summary>
    ''' 入金項目フィールド名称設定<br/>
    ''' </summary>
    Public Class FieldName_Nkinkomk

        Public Property FieldName() As Njc.Common.SetArrayData
        Public Property fieldcnt As Integer

        Public Sub New()

            fieldcnt = 7
            Me.FieldName = New Njc.Common.SetArrayData(fieldcnt)
            Me.FieldName(0) = REL_FROM_NAME & "入金項目№"
            Me.FieldName(1) = REL_FROM_NAME & "入金項目名称"
            Me.FieldName(2) = REL_FROM_NAME & "入金項目区分"
            Me.FieldName(3) = REL_TO_NAME & "入金項目№"
            Me.FieldName(4) = REL_TO_NAME & "入金項目名称"
            Me.FieldName(5) = REL_TO_NAME & "入金項目属性№"
            Me.FieldName(6) = REL_TO_NAME & "入金項目属性名称"
            Me.FieldName(7) = REL_TO_NAME & "入金項目区分"

        End Sub

    End Class

    ''' <summary>
    ''' 構造フィールド名称設定<br/>
    ''' </summary>
    Public Class FieldName_Kozo

        Public Property FieldName() As Njc.Common.SetArrayData
        Public Property fieldcnt As Integer

        Public Sub New()

            fieldcnt = 3
            Me.FieldName = New Njc.Common.SetArrayData(fieldcnt)
            Me.FieldName(0) = REL_FROM_NAME & "構造№"
            Me.FieldName(1) = REL_FROM_NAME & "構造名称"
            Me.FieldName(2) = REL_TO_NAME & "構造№"
            Me.FieldName(3) = REL_TO_NAME & "構造名称"

        End Sub

    End Class

    ''' <summary>
    ''' 口座種別フィールド名称設定<br/>
    ''' </summary>
    Public Class FieldName_Kozasyu

        Public Property FieldName() As Njc.Common.SetArrayData
        Public Property fieldcnt As Integer

        Public Sub New()

            fieldcnt = 3
            Me.FieldName = New Njc.Common.SetArrayData(fieldcnt)
            Me.FieldName(0) = REL_FROM_NAME & "口座種別№"
            Me.FieldName(1) = REL_FROM_NAME & "口座種別名称"
            Me.FieldName(2) = REL_TO_NAME & "口座種別№"
            Me.FieldName(3) = REL_TO_NAME & "口座種別名称"

        End Sub

    End Class

    ''' <summary>
    ''' 設備フィールド名称設定<br/>
    ''' </summary>
    Public Class FieldName_Setubi

        Public Property FieldName() As Njc.Common.SetArrayData
        Public Property fieldcnt As Integer

        Public Sub New()

            fieldcnt = 14
            Me.FieldName = New Njc.Common.SetArrayData(fieldcnt)
            Me.FieldName(0) = REL_FROM_NAME & "設備グループ№"
            Me.FieldName(1) = REL_FROM_NAME & "設備グループ名称"
            Me.FieldName(2) = REL_FROM_NAME & "設備№"
            Me.FieldName(3) = REL_FROM_NAME & "設備名称"

            Me.FieldName(4) = REL_TO_NAME & "設備グループGUID_GRP"
            Me.FieldName(5) = REL_TO_NAME & "設備グループ№"
            Me.FieldName(6) = REL_TO_NAME & "設備グループ名称"

            Me.FieldName(7) = REL_TO_NAME & "設備GUID_MS"
            Me.FieldName(8) = REL_TO_NAME & "設備グループGUID_MS"
            Me.FieldName(9) = REL_TO_NAME & "設備№"
            Me.FieldName(10) = REL_TO_NAME & "設備名称"

            Me.FieldName(11) = REL_TO_NAME & "項目GUID"
            Me.FieldName(12) = REL_TO_NAME & "設備GUID_KOMK"
            Me.FieldName(13) = REL_TO_NAME & "項目№"
            Me.FieldName(14) = REL_TO_NAME & "項目名称"

        End Sub

    End Class

#End Region


    ' ''' <summary>
    ' ''' 【RelationModel】
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Public Class RelationModel

    '    Public Property ServerName As String                                    '接続サーバー名
    '    Public Property InitialCatalog As String                                'カタログ名
    '    Public Property NetworkLibrary As Integer                               '接続で使用するネットワークライブラリの種類
    '    Public Property User As String                                          '認証で使用するユーザ名
    '    Public Property Pass As String                                          '認証で使用するパスワード
    '    Public Property TimeOut As String                                       'DB接続時タイムアウト設定値(SEC)


    '    ''' <summary>
    '    ''' ・SQLServer接続初期設定() → 2015.07.06 sol レビュー後修正_レビュー№2 呼出起動時の処理のため、引数に単体起動フラグ(singlesta)とコマンドラインを追加
    '    ''' </summary>
    '    ''' <remarks></remarks>
    '    Public Class DefSQLConnection
    '        Inherits RelationModel

    '        Public Sub New(ByVal flg As Integer, ByVal singlesta As Boolean, Optional ByVal cmdline As Object = Nothing)

    '            '2015.07.06 sol レビュー後修正_レビュー№2 -chg sta
    '            'If flg = 0 Then
    '            '    'V7用固定設定
    '            '    Me.ServerName = "PC-1KA_SOL\SQL2K8"
    '            '    Me.InitialCatalog = "fk5dtsql"
    '            '    Me.User = "sa"
    '            '    Me.Pass = "p7s2#c1j3n"
    '            'Else
    '            '    'V10用固定設定
    '            '    Me.ServerName = "PC-1KA_SOL\SQL2012"
    '            '    Me.InitialCatalog = "fk8db"
    '            '    Me.User = "sa"
    '            '    Me.Pass = "p7s2#c1j3n"
    '            'End If
    '            'Me.TimeOut = "300"
    '            If singlesta Then
    '                If flg = 0 Then
    '                    'V7用固定設定
    '                    Me.ServerName = "PC-1KA_SOL\SQL2K8"
    '                    Me.InitialCatalog = "fk5dtsql"
    '                    Me.User = "sa"
    '                    Me.Pass = "p7s2#c1j3n"
    '                Else
    '                    'V10用固定設定
    '                    Me.ServerName = "PC-1KA_SOL\SQL2012"
    '                    Me.InitialCatalog = "fk8db"
    '                    Me.User = "sa"
    '                    Me.Pass = "p7s2#c1j3n"
    '                End If
    '                Me.TimeOut = "300"
    '            Else
    '                If flg = 0 Then
    '                    'V7用固定設定
    '                    Me.ServerName = "PC-1KA_SOL\SQL2K8"
    '                    Me.InitialCatalog = "fk5dtsql"
    '                    Me.User = "sa"
    '                    Me.Pass = "p7s2#c1j3n"
    '                Else
    '                    'V10用固定設定
    '                    Me.ServerName = "PC-1KA_SOL\SQL2012"
    '                    Me.InitialCatalog = "fk8db"
    '                    Me.User = "sa"
    '                    Me.Pass = "p7s2#c1j3n"
    '                End If
    '                Me.TimeOut = "300"
    '            End If
    '            '2015.07.06 sol レビュー後修正_レビュー№2 -chg end

    '        End Sub

    '    End Class






    'End Class

End Namespace

