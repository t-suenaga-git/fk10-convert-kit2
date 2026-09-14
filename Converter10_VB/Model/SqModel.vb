Namespace Njc.Model

#Region "過剰金情報モデル"

    Public Class Azukanri_Model

        Public Property Vari_Azu_guid As String
        Public Property Vari_Azu_ymd As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Dispnkin_name As String
        Public Property Vari_Nkbn_no As String
        Public Property Vari_Azu_gak As String
        Public Property Vari_Stakeholder_kbn As String
        Public Property Vari_Stakeholder_no As String
        Public Property Vari_Sui_guid As String
        Public Property Vari_Tanto_no As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Tuki_kbn As String
        Public Property Vari_Ky_nkin_no As String
        Public Property Vari_Nkin_recno As String
        Public Property Vari_Kynkin_guid As String
        Public Property Vari_Yoteiazu_flg As String
        Public Property Vari_Printazuryo_guid As String
        Public Property Vari_Siwake_flg As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Tujo_siwakegak As String

    End Class

#End Region

#Region "過剰金情報モデル"

    Public Class Suikanri_Model

        Public Property Vari_Sui_guid As String
        Public Property Vari_Sui_ymd As String
        Public Property Vari_Sui_kbn As String
        Public Property Vari_Stakeholder_kbn As String
        Public Property Vari_Stakeholder_no As String
        Public Property Vari_Nkbn_no As String
        Public Property Vari_Gak As String
        Public Property Vari_Tanto_no As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Yatin_fkom_kozano As String
        Public Property Vari_Fkae_rirekino As String
        Public Property Vari_Ns_no As String
        Public Property Vari_Ns_bunkatu_no As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Cvpay_no As String
        Public Property Vari_Kokyaku_bango As String
        Public Property Vari_Fkae_no As String

    End Class

#End Region

#Region "運用開始時未収滞納金情報モデル"

    Public Class Unyotainodata_Model

        Public Property Vari_Unyotaino_guid As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Gt_ym As String
        Public Property Vari_Sq_simeymd As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkbn_yotei As String
        Public Property Vari_Yotei_yatin_kozano As String
        Public Property Vari_Sqsaki_no As String
        Public Property Vari_Sq_gak As String
        Public Property Vari_Sq_zeikbn As String
        Public Property Vari_Sq_zeigak As String
        Public Property Vari_So_gak As String
        Public Property Vari_So_zeikbn As String
        Public Property Vari_So_zeigak As String
        Public Property Vari_Sosumi_flg As String
        Public Property Vari_Kanrigak_rit As String
        Public Property Vari_Sosaki_no As String
        Public Property Vari_Sokoza_no As String
        Public Property Vari_So_rit As String
        Public Property Vari_Soymd_gaitoukbn As String
        Public Property Vari_Soymd_simekbn As String
        Public Property Vari_Soymd_sokintukikbn As String
        Public Property Vari_Soymd_sokinsimekbn As String
        Public Property Vari_Zei_rit As String
        Public Property Vari_Sqbuild_flg As String
        Public Property Vari_Tanto_no As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Kanri_taxflg As String
        Public Property Vari_Kanritesu_zeiumukbn As String
        Public Property Vari_Kanri_utizeiflg As String

    End Class

#End Region

#Region "請求情報モデル"

    Public Class Sqdata_Model

        Public Property Vari_Sqmei_guid As String
        Public Property Vari_Gt_ym As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Dispnkin_name As String
        Public Property Vari_Nkbn_yotei As String
        Public Property Vari_Sq_gak As String
        Public Property Vari_Sq_zeikbn As String
        Public Property Vari_Sq_zeigak As String
        Public Property Vari_Sq_simeymd As String
        Public Property Vari_Multisqsaki_flg As String
        Public Property Vari_Sqsaki_kbn As String
        Public Property Vari_Sqsaki_no As String
        Public Property Vari_Yotei_yatin_kozano As String
        Public Property Vari_Fkae_rirekino As String
        Public Property Vari_Fkae_rirekimeino As String
        Public Property Vari_Fkae_sqtaisyoflg As String
        Public Property Vari_Fkae_sqselectflg As String
        Public Property Vari_Fkae_sqflg As String
        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Tuki_kbn As String
        Public Property Vari_Nkin_recno As String
        Public Property Vari_Nkin_kbn As String
        Public Property Vari_Kynkin_guid As String
        Public Property Vari_Nextky_flg As String
        Public Property Vari_Hendosq_guid As String
        Public Property Vari_Hendo_kbn As String
        Public Property Vari_Kensin_ymd As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_So_gak As String
        Public Property Vari_So_zeikbn As String
        Public Property Vari_So_zeigak As String
        Public Property Vari_Etcsq_flg As String
        Public Property Vari_Etcsq_sorit As String
        Public Property Vari_Etcsq_soymd As String
        Public Property Vari_Bunkatu_guid As String
        Public Property Vari_Sh_kbn As String
        Public Property Vari_Shsaki_kbn As String
        Public Property Vari_Shsaki_no As String
        Public Property Vari_Shsaki_kozano As String
        Public Property Vari_Ikkatu_nkin_bango As String
        Public Property Vari_Jisya_no As String
        Public Property Vari_Tanto_no As String
        Public Property Vari_Kjsqkmk_guid As String
        Public Property Vari_Koteirule_guid As String
        Public Property Vari_Kojorule_flg As String
        Public Property Vari_Sq_hakkoyoteiymd As String
        Public Property Vari_Sqsort_no As String
        Public Property Vari_Tateazu_flg As String
        Public Property Vari_Tateazu_guid As String
        Public Property Vari_Sikikinzuiji_guid As String
        Public Property Vari_Unyotaino_guid As String
        Public Property Vari_Sokaisyu_guid As String
        Public Property Vari_Szen_kbn As String
        Public Property Vari_Szen_no As String
        Public Property Vari_Szen_sqno As String
        Public Property Vari_Zei_rit As String
        Public Property Vari_Edit_flg As String
        Public Property Vari_Lock_flg As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Tatekae_siwakegak As String
        Public Property Vari_Tatekae_siwakezeigak As String
        Public Property Vari_Sq_torikomiymd As String
        Public Property Vari_Cvpay_no As String
        Public Property Vari_Kokyaku_bango As String
        Public Property Vari_Cvpay_sqflg As String
        Public Property Vari_Etcsq_kanritesukbn As String
        Public Property Vari_Misyu_siwake_flg As String
        Public Property Vari_Zatusyu_kbn As String
        Public Property Vari_Shdirect_flg As String         '20160519 EXEUpdateに伴う修正 請求情報 -add
        Public Property Vari_Koteishrule_guid As String     '20160519 EXEUpdateに伴う修正 請求情報 -add
        Public Property Vari_Kojorule_taisyokbn As String   '20160519 EXEUpdateに伴う修正 請求情報 -add
    End Class

#End Region

#Region "請求消込情報モデル" '20160601 請求データ取得処理の修正 新規追加

    Public Class Sqdata_sqkesi_Model

        Public Property Vari_Sqmei_guid As String
        Public Property Vari_Sqkesi_recno As String
        Public Property Vari_Kesi_ymd As String
        Public Property Vari_Nkbn_jitu As String
        Public Property Vari_Kesi_gak As String
        Public Property Vari_Kesi_zeigak As String
        Public Property Vari_Sui_guid As String
        Public Property Vari_Hikiazu_guid As String
        Public Property Vari_Idoazu_guid As String
        Public Property Vari_Printazuryo_guid As String
        Public Property Vari_Siwake_flg As String
        Public Property Vari_Tanto_no As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "送金予定情報モデル"

    Public Class Sqdata_so_Model

        Public Property Vari_Sqmei_guid As String
        Public Property Vari_So_recno As String
        Public Property Vari_So_gak As String
        Public Property Vari_So_zeikbn As String
        Public Property Vari_So_zeigak As String
        Public Property Vari_Edit_flg As String
        Public Property Vari_Lock_flg As String
        Public Property Vari_Hosyo_flg As String
        Public Property Vari_So_yoteiymd As String
        Public Property Vari_So_kakuymd As String
        Public Property Vari_Sokotik_guid As String
        Public Property Vari_Sorule_guid As String
        Public Property Vari_Sorule_no As String
        Public Property Vari_Sosaki_sorule_guid As String
        Public Property Vari_Sosaki_sorule_no As String
        Public Property Vari_Sokin_rit As String
        Public Property Vari_Calc_sqgak As String
        Public Property Vari_Sokaisyu_kbn As String
        Public Property Vari_Sokaisyu_guid As String
        Public Property Vari_Siwake_flg As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "送金管理情報モデル" '20160530 送金データ調査での不足箇所修正

    Public Class Sokanri_Model

        Public Property Vari_Sokanri_guid As String
        Public Property Vari_Sorule_guid As String
        Public Property Vari_Sorule_no As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_So_ymd As String
        Public Property Vari_Sosaki_kbn As String
        Public Property Vari_Sosaki_no As String
        Public Property Vari_Sokoza_no As String
        Public Property Vari_So_gak As String
        Public Property Vari_Sokotik_guid As String
        Public Property Vari_Karikaku_flg As String
        Public Property Vari_Sgf_no As String
        Public Property Vari_Error_kbn As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Karikaku_guid As String
        Public Property Vari_Shdirect_flg As String

    End Class

#End Region

#Region "変動費検針情報モデル"

    Public Class Hendodata_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Hendo_kbn As String
        Public Property Vari_Kensin_ymd As String
        Public Property Vari_Gt_ym As String
        Public Property Vari_History As String

    End Class

#End Region

#Region "変動費検針情報モデル (明細)"

    Public Class Hendodata_meisai_Model

        Public Property Vari_Bk_guid As String
        Public Property Vari_Hendo_kbn As String
        Public Property Vari_Kensin_ymd As String
        Public Property Vari_Gt_ym As String
        Public Property Vari_Kensin_ptn As String
        Public Property Vari_Hendo_guid As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_Hendosq_guid As String
        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Sqsaki_no As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkbn_no As String
        Public Property Vari_Sqsime_ymd As String
        Public Property Vari_Zenkai_ymd As String
        Public Property Vari_Zenkai_metervalue As String
        Public Property Vari_Zenkai_siyoryo As String
        Public Property Vari_Zenkai_sqgak As String
        Public Property Vari_Zenkai_sqzeigak As String
        Public Property Vari_Kensin_metervalue As String
        Public Property Vari_Kensin_siyoryo As String
        Public Property Vari_Kensin_ryokin1 As String
        Public Property Vari_Kensin_ryokin2 As String
        Public Property Vari_Kensin_ryokin3 As String
        Public Property Vari_Kensin_ryokin4 As String
        Public Property Vari_Kensin_ryokin5 As String
        Public Property Vari_Kensin_sqgak As String
        Public Property Vari_Kensin_sqzeigak As String
        Public Property Vari_Kensin_biko As String
        Public Property Vari_Kensin_soyoteiymd As String
        Public Property Vari_Kensin_sorit As String
        Public Property Vari_Kensin_sogak As String
        Public Property Vari_Kensin_sozeigak As String
        Public Property Vari_Zenkaihendosq_guid As String

    End Class

#End Region

#Region "家主固定控除情報モデル"

    Public Class Koteirule_Model

        Public Property Vari_Koteirule_guid As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Nkin_sortorder As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Nkin_name As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_Sosai_flg As String
        Public Property Vari_Gak As String
        Public Property Vari_Zei_kbn As String
        Public Property Vari_Zei_gak As String
        Public Property Vari_Nkbn_yotei As String
        Public Property Vari_Yatin_kozano As String
        Public Property Vari_Gtstart_ym As String
        Public Property Vari_Gtend_ym As String
        Public Property Vari_Tateazu_flg As String
        Public Property Vari_Sq_mmkbn As String
        Public Property Vari_Sq_ptn As String
        Public Property Vari_Sq_interval As String
        Public Property Vari_Sq_nen As String
        Public Property Vari_Sq_tuki As String
        Public Property Vari_Zei_rit As String
        Public Property Vari_Biko As String
        Public Property Vari_History As String
        Public Property Vari_Sq_simekbn As String
        Public Property Vari_Ryosyu_flg As String

    End Class

#End Region

#Region "家主請求控除情報モデル"

    Public Class Kjdata_Model

        Public Property Vari_Kj_guid As String
        Public Property Vari_Kjkmk_guid As String
        Public Property Vari_Bk_guid As String
        Public Property Vari_Hy_guid As String
        Public Property Vari_Kj_yoteiymd As String
        Public Property Vari_Kjsort_no As String
        Public Property Vari_Gt_ym As String
        Public Property Vari_Nkin_no As String
        Public Property Vari_Dispnkin_name As String
        Public Property Vari_Kj_gak As String
        Public Property Vari_Kj_zeikbn As String
        Public Property Vari_Kj_zeigak As String
        Public Property Vari_Lock_flg As String
        Public Property Vari_Ryosyu_flg As String
        Public Property Vari_Tateazu_flg As String
        Public Property Vari_Tateazu_guid As String
        Public Property Vari_Bunkatu_guid As String
        Public Property Vari_Koteirule_guid As String
        Public Property Vari_Kojorule_flg As String
        Public Property Vari_So_kakuymd As String
        Public Property Vari_Sokotik_guid As String
        Public Property Vari_Sorule_guid As String
        Public Property Vari_Sorule_no As String
        Public Property Vari_Sosaki_sorule_guid As String
        Public Property Vari_Sosaki_sorule_no As String
        Public Property Vari_Sosaki_no As String
        Public Property Vari_Sokoza_no As String
        Public Property Vari_Sikikinzuiji_guid As String
        Public Property Vari_Ky_guid As String
        Public Property Vari_Ky_recno As String
        Public Property Vari_Sokaisyu_guid As String
        Public Property Vari_Kanritesu_kbn As String
        Public Property Vari_Kanritesu_cyosyukbn As String
        Public Property Vari_Kanrigak_rit As String
        Public Property Vari_Szen_kbn As String
        Public Property Vari_Szen_no As String
        Public Property Vari_Szen_sqno As String
        Public Property Vari_Siwake_flg As String
        Public Property Vari_Zei_rit As String
        Public Property Vari_Biko As String
        Public Property Vari_Kj_torikomiymd As String
        Public Property Vari_History As String
        Public Property Vari_Kubungrp_guid As String
        Public Property Vari_Kojorule_taisyokbn As String        '20160519 EXEUpdateに伴う修正 家主請求控除情報

    End Class

#End Region

End Namespace
