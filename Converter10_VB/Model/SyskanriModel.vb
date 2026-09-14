Namespace Njc.Model

#Region "初期設定基本情報モデル"

    Public Class Profile_fk_base_Model

        Public Property Vari_KenNo As String
        Public Property Vari_SiNo As String

    End Class

#End Region

#Region "初期設定税編集情報モデル"

    Public Class Profile_fk_zei_Model

        Public Property Vari_No As String
        Public Property Vari_ZeiRit As String
        Public Property Vari_StartYmd As String
        Public Property Vari_EndYmd As String
        Public Property Vari_Biko As String

    End Class

#End Region

#Region "初期設定変換文字情報モデル"

    Public Class Profile_fk_henkanmoji_Model

        Public Property Vari_No As String
        Public Property Vari_Target As String
        Public Property Vari_Cnv As String

    End Class

#End Region

#Region "初期設定入金項目集約情報モデル"

    Public Class Profile_fk_nkinkomkmerge_Model

        Public Property Vari_Identity As String
        Public Property Vari_Name As String
        Public Property Vari_DaihyoNkinNo As String
        Public Property Vari_Int() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer
        'Public Property Vari_Int1 As String
        'Public Property Vari_Int2 As String
        'Public Property Vari_Int3 As String
        'Public Property Vari_Int4 As String
        'Public Property Vari_Int5 As String
        'Public Property Vari_Int6 As String
        'Public Property Vari_Int7 As String
        'Public Property Vari_Int8 As String
        'Public Property Vari_Int9 As String
        'Public Property Vari_Int10 As String
        'Public Property Vari_Int11 As String
        'Public Property Vari_Int12 As String
        'Public Property Vari_Int13 As String
        'Public Property Vari_Int14 As String
        'Public Property Vari_Int15 As String
        'Public Property Vari_Int16 As String
        'Public Property Vari_Int17 As String
        'Public Property Vari_Int18 As String
        'Public Property Vari_Int19 As String
        'Public Property Vari_Int20 As String

        Public Sub New(int As Integer)

            Vari_Int = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

#End Region

End Namespace
