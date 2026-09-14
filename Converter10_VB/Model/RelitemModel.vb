Namespace Njc.Model

#Region "物件分類モデル"

    Public Class M_bk_rui_Model

        Public Property Vari_Bk_ruino As String
        Public Property Vari_Bk_ruiname As String
        Public Property Vari_Bk_ruibiko As String
        Public Property Vari_Bk_ruiuseflg As String
        Public Property Vari_Sincyoku_keiyakukbn As String
        Public Property Vari_Sincyoku_kosinkbn As String
        Public Property Vari_Sincyoku_kaiyakukbn As String
        Public Property Vari_History As String
        Public Property Vari_Homemate_ruikbn As String
        Public Property Vari_Bk_endofmonthflg As String
        Public Property Vari_Ikkatukariage_siwakekbn As String

    End Class

#End Region

#Region "部屋分類モデル"

    Public Class M_hy_rui_Model

        Public Property Vari_Hy_ruino As String
        Public Property Vari_Hy_ruiname As String
        Public Property Vari_Hy_ruisyubetu As String
        Public Property Vari_Hy_ruiuseflg As String
        Public Property Vari_Hy_carportflg As String
        Public Property Vari_History As String
        'Public Property Vari_Hy_ruisortorder As String      'hy_ruisortorderは自動で挿入されるため対象外

    End Class

#End Region

#Region "入金区分マスタモデル"

    Public Class M_nkbn_Model

        Public Property Vari_Nkbn_no As String          '1-99のキー(1:現金 2:振込 3:振替 4:カード 5:代弁 99:移動)
        Public Property Vari_Nkbn_order As String       '表示順
        Public Property Vari_Nkbn_name As String        '入金区分名(1:現金 2:振込 3:振替 4:カード 5:代弁 99:移動)
        Public Property Vari_Nkbn_shortname As String   '入金区分略称
        Public Property Vari_Nkbn_zokusei As String     '1:現金, 2:振込, 3:振替, 4:その他, 99:移動
        Public Property Vari_History As String

    End Class

#End Region

#Region "契約分類マスタモデル"

    Public Class M_ky_rui_Model

        Public Property Vari_Ky_ruino As String         '契約分類No
        Public Property Vari_Ky_ruiname As String       '契約分類名
        Public Property Vari_Ky_ruibiko As String       '契約分類備考
        Public Property Vari_Ky_ruitutimm As String     '契約終了通知月（◯ヶ月前）
        Public Property Vari_Ky_ruicolor As String      '契約分類色
        Public Property Vari_Teisyaku_flg As String     '1:使用する 2:使用しない
        Public Property Vari_History As String
        Public Property Vari_Useflg As String           '1:検索対象 0:検索対象外

    End Class

#End Region

#Region "入金項目マスタモデル"

    Public Class M_nkin_Model

        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkin_name As String
        Public Property Vari_Nkin_printname As String
        Public Property Vari_Nkin_ruino As String
        Public Property Vari_Nkin_zkseino As String
        Public Property Vari_Nkin_useflg As String
        Public Property Vari_Keiyaku_nkinno As String
        Public Property Vari_Keiyakuhiki_nkinno As String
        Public Property Vari_Kosin_nkinno As String
        Public Property Vari_Kaiyaku_nkinno As String
        Public Property Vari_Kaiyakuhiki_nkinno As String
        Public Property Vari_Biko_nkin As String
        Public Property Vari_History As String
        Public Property Vari_Ryosyu_flg As String

    End Class

#End Region


    '20160829 設備の新規挿入処理を追加 (現時点で未使用→後で削除) -del sta
#Region "設備グループマスタモデル"

    Public Class M_setubi_rel_grp_Model

        Public Property Vari_Setubi_grpguid As String
        Public Property Vari_Setubi_grpsortorder As String
        Public Property Vari_Setubi_grpname As String
        Public Property Vari_Setubi_grpuseflg As String
        Public Property Vari_Setubi_grpsyskbn As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "設備マスタモデル"

    Public Class M_setubi_rel_Model

        Public Property Vari_Setubi_guid As String
        Public Property Vari_Setubi_grpguid As String
        Public Property Vari_Setubi_sortorder As String
        Public Property Vari_Setubi_code As String
        Public Property Vari_Setubi_name As String
        Public Property Vari_Setubi_useflg As String
        Public Property Vari_Setubi_hyuseflg As String
        Public Property Vari_Setubi_disp1name As String
        Public Property Vari_Setubi_disp2name As String
        Public Property Vari_Setubi_disp3name As String
        Public Property Vari_History As String
        Public Property Vari_Jyuyojiko_flg As String

    End Class

#End Region

#Region "設備項目マスタモデル"

    Public Class M_setubi_rel_komk_Model

        Public Property Vari_Komok_guid As String
        Public Property Vari_Setubi_guid As String
        Public Property Vari_Komok_sortorder As String
        Public Property Vari_Komok_name As String
        Public Property Vari_Komok_code1 As String
        Public Property Vari_Komok_code2 As String
        Public Property Vari_Komok_code3 As String
        Public Property Vari_Default_flg As String
        Public Property Vari_Disp1name As String
        Public Property Vari_Disp2name As String
        Public Property Vari_Disp3name As String
        Public Property Vari_Disp1iconguid As String
        Public Property Vari_Disp2iconguid As String
        Public Property Vari_Disp3iconguid As String
        Public Property Vari_History As String

    End Class

#End Region
    '20160829 設備の新規挿入処理を追加 (現時点で未使用→後で削除) -del end

End Namespace
