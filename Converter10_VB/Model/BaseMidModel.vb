Namespace Njc.Model

#Region "自社口座情報モデル"

    Public Class Jisyadata_koza_BaseMid_Model

        Public Property Vari_Jisya_no As String
        Public Property Vari_Jisya_kozano As String
        Public Property Vari_Kinyu_no As String
        Public Property Vari_Kinyu_tenno As String
        Public Property Vari_Koza_syubetu As String
        Public Property Vari_Koza_bango As String
        Public Property Vari_Koza_meigi As String
        Public Property Vari_Koza_meigikana As String
        Public Property Vari_Yucyokoza_kigo1 As String
        Public Property Vari_Yucyokoza_kigo2 As String
        Public Property Vari_Yucyokoza_bango As String
        Public Property Vari_Biko_koza As String
        Public Property Vari_History As String
        Public Property Vari_Useflg As String

    End Class

#End Region

#Region "部屋設備情報モデル"

    Public Class Hydata_setubilst_BaseMid_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Komok_guid As String
        Public Property Vari_Override_kbn As String
        Public Property Vari_Komok_name As String
        Public Property Vari_Disp1name As String
        Public Property Vari_Disp2name As String
        Public Property Vari_Disp3name As String
        Public Property Vari_Disp1iconguid As String
        Public Property Vari_Disp2iconguid As String
        Public Property Vari_Disp3iconguid As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "鍵タイトルマスタモデル"

    Public Class M_kagi_title_BaseMid_Model

        Public Property Vari_Kagi_kbn As String
        Public Property Vari_Kagi_no As String
        Public Property Vari_Kagi_name As String
        Public Property Vari_History As String

    End Class

#End Region

End Namespace
