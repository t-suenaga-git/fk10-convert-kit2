Namespace Njc.Model

#Region "送金ルール基本情報モデル"

    Public Class Sorule_Model

        Public Property Vari_Sorule_guid As String
        Public Property Vari_Sorule_no As String
        Public Property Vari_Relation_guid As String
        Public Property Vari_Kubunsyo As String
        Public Property Vari_Sorule_startymd As String
        Public Property Vari_Sorule_endymd As String
        Public Property Vari_Kanri_keitaikbn As String
        Public Property Vari_Ikkatu_kanriflg As String
        Public Property Vari_Cyukai_kyflg As String
        Public Property Vari_Cyukai_koflg As String
        Public Property Vari_Cyukai_kaiflg As String
        Public Property Vari_Soymd1_gaitoukbn As String
        Public Property Vari_Soymd1_simekbn As String
        Public Property Vari_Soymd1_sokintukikbn As String
        Public Property Vari_Soymd1_sokinsimekbn As String
        Public Property Vari_Soymd2_gaitoukbn As String
        Public Property Vari_Soymd2_simekbn As String
        Public Property Vari_Soymd2_sokintukikbn As String
        Public Property Vari_Soymd2_sokinsimekbn As String
        Public Property Vari_Soymd3_gaitoukbn As String
        Public Property Vari_Soymd3_simekbn As String
        Public Property Vari_Soymd3_sokintukikbn As String
        Public Property Vari_Soymd3_sokinsimekbn As String
        Public Property Vari_Soymd4_gaitoukbn As String
        Public Property Vari_Soymd4_simekbn As String
        Public Property Vari_Soymd4_sokintukikbn As String
        Public Property Vari_Soymd4_sokinsimekbn As String
        Public Property Vari_Soymd5_gaitoukbn As String
        Public Property Vari_Soymd5_simekbn As String
        Public Property Vari_Soymd5_sokintukikbn As String
        Public Property Vari_Soymd5_sokinsimekbn As String
        Public Property Vari_Sosaki_multikbn As String
        Public Property Vari_Sosaki_koteiflg As String
        Public Property Vari_Sosaki_koteisu As String
        Public Property Vari_Sosaki_anbunflg As String
        Public Property Vari_Sosaki_hasuuketorisaki As String
        Public Property Vari_Sosaki_hasuadjustmentflg As String
        Public Property Vari_Ikkatu_kbn As String
        Public Property Vari_Ikkatu_bkgak As String
        Public Property Vari_Ikkatu_menseki As String
        Public Property Vari_Ikkatu_mensekimonth As String
        Public Property Vari_Kanritesu_kbn As String
        Public Property Vari_Kanritesu_cyosyukbn As String
        Public Property Vari_Kanri_reigaiky As String
        Public Property Vari_Kanri_reigaikai As String
        Public Property Vari_Hyteigaku_kbn As String
        Public Property Vari_Hyteigaku_kanrigak As String
        Public Property Vari_Hyteigaku_hiwariflg As String
        Public Property Vari_Bkteigaku_kanrigak As String
        Public Property Vari_Kanri_taxflg As String
        Public Property Vari_Biko_basic As String
        Public Property Vari_Sh_daihyohyflg As String
        Public Property Vari_Sh_kozabetuflg As String
        Public Property Vari_Sh_kozabetuListNokbn As String
        Public Property Vari_Hyikkatu_nkin_no1kbn As String
        Public Property Vari_Hyikkatu_nkin_no2kbn As String
        Public Property Vari_Hyikkatu_nkin_no3kbn As String
        Public Property Vari_Hyikkatu_nkin_no4kbn As String
        Public Property Vari_Hyikkatu_nkin_no5kbn As String
        Public Property Vari_Kanritesu_zeiumukbn As String
        Public Property Vari_History As String
        Public Property Vari_Rowid As String
        Public Property Vari_No_soruleflg As String
        Public Property Vari_Cyukai_zuijisokin As String
        Public Property Vari_Ikkatu_calckbn As String
        Public Property Vari_Ikkatu_taxflg As String
        Public Property Vari_Kanri_utizeiflg As String
        Public Property Vari_Soymd_basis As String

    End Class

#End Region

#Region "送金ルール送金先情報モデル"

    Public Class Sorule_sosaki_Model

        Public Property Vari_Sorule_guid As String
        Public Property Vari_Sorule_no As String
        Public Property Vari_Sosaki_recno As String
        Public Property Vari_So_ow_no As String
        Public Property Vari_So_ow_kozano As String
        Public Property Vari_So_rit As String
        Public Property Vari_So_fixsogak As String

    End Class

#End Region

#Region "送金ルール入金項目情報モデル"

    Public Class Sorule_nk_cmrule_Model

        Public Property Vari_Sorule_guid As String
        Public Property Vari_Sorule_no As String
        Public Property Vari_Taisyokbn As String
        Public Property Vari_Nkin_sortorder As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Sokin_rit As String
        Public Property Vari_Kanrigak_rit As String
        Public Property Vari_Hosyo_flg As String
        Public Property Vari_So_no As String

    End Class

#End Region

#Region "送金ルール控除項目情報モデル"

    Public Class Sorule_kojo_cmrule_Model

        Public Property Vari_Sorule_guid As String
        Public Property Vari_Sorule_no As String
        Public Property Vari_Kojo_taisyokbn As String
        Public Property Vari_Nkin_sortorder As String
        Public Property Vari_Taisyonkin_no As String
        Public Property Vari_Sosai_flg As String
        Public Property Vari_Kojocalc_kbn As String
        Public Property Vari_Kojo_gak As String
        Public Property Vari_Kojo_gakzeikbn As String
        Public Property Vari_Kojonkin_no As String
        Public Property Vari_Kojo_rit As String
        Public Property Vari_Kojo_ritzeikbn As String
        Public Property Vari_Kojo_ritutizei As String
        Public Property Vari_Zei_rit As String
        Public Property Vari_Tateazu_flg As String
        Public Property Vari_Kojo_zeigak As String
        Public Property Vari_Sotaisyo_flg As String
        Public Property Vari_Sotaisyonkin_no As String

    End Class

#End Region

End Namespace
