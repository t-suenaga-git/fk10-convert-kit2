Namespace Njc.Model

#Region "送信設定基本情報モデル"

    Public Class M_sendsetting_Model

        Public Property Vari_Setting_guid As String
        Public Property Vari_Setting_sortorder As String
        Public Property Vari_Setting_name As String
        Public Property Vari_Keisai_siten As String
        Public Property Vari_Datalink_id As String
        Public Property Vari_Adconfirm_limitday As String
        Public Property Vari_History As String
        Public Property Vari_Basedate As String
        Public Property Vari_Basedate_flg As String
        Public Property Vari_Emailaddr1 As String
        Public Property Vari_Emailaddr2 As String
        Public Property Vari_Emailaddr3 As String
        Public Property Vari_Emailaddr4 As String
        Public Property Vari_Emailaddr5 As String

    End Class

#End Region

#Region "送信設定情報モデル"

    Public Class M_site_sendsetting_Model

        Public Property Vari_Sitesetting_guid As String
        Public Property Vari_Setting_guid As String
        Public Property Vari_Site_no As String
        Public Property Vari_Site_id As String
        Public Property Vari_Site_sendumu As String
        Public Property Vari_Site_data As String
        Public Property Vari_Site_password As String

    End Class

#End Region

#Region "広告補足情報モデル"

    Public Class Hydata_kokoku_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Site_no As String
        Public Property Vari_Item As String

    End Class

#End Region

#Region "ポータル連動部屋分類情報モデル"

    Public Class M_hy_ruisite_Model

        Public Property Vari_Hy_ruino As String
        Public Property Vari_Site_no As String
        Public Property Vari_Hy_ruisiteitemdata As String
        
    End Class

#End Region

#Region "部屋毎送信情報モデル"

    Public Class Hydata_sosin_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Site_no As String
        Public Property Vari_Site_jisyano As String
        Public Property Vari_Select_no As String
        Public Property Vari_Komok_guid As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "BtoBグループ設定情報モデル"

    Public Class Hydata_btobgroup_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Group_no As String
        Public Property Vari_Display_kbn As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "地図表示詳細設定情報モデル"

    Public Class Hydata_mapdisp_Model

        Public Property Vari_Hy_guid As String
        Public Property Vari_Site_no As String
        Public Property Vari_Site_jisyano As String
        Public Property Vari_Map_dispflg As String
        Public Property Vari_History As String

    End Class

#End Region

End Namespace
