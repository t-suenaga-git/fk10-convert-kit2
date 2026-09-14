
namespace Converter10.Njc.Model
{

    #region 過剰金情報モデル

    public class Azukanri_Model
    {

        public string Vari_Azu_guid { get; set; }
        public string Vari_Azu_ymd { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Dispnkin_name { get; set; }
        public string Vari_Nkbn_no { get; set; }
        public string Vari_Azu_gak { get; set; }
        public string Vari_Stakeholder_kbn { get; set; }
        public string Vari_Stakeholder_no { get; set; }
        public string Vari_Sui_guid { get; set; }
        public string Vari_Tanto_no { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Tuki_kbn { get; set; }
        public string Vari_Ky_nkin_no { get; set; }
        public string Vari_Nkin_recno { get; set; }
        public string Vari_Kynkin_guid { get; set; }
        public string Vari_Yoteiazu_flg { get; set; }
        public string Vari_Printazuryo_guid { get; set; }
        public string Vari_Siwake_flg { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Tujo_siwakegak { get; set; }

    }

    #endregion

    #region 過剰金情報モデル

    public class Suikanri_Model
    {

        public string Vari_Sui_guid { get; set; }
        public string Vari_Sui_ymd { get; set; }
        public string Vari_Sui_kbn { get; set; }
        public string Vari_Stakeholder_kbn { get; set; }
        public string Vari_Stakeholder_no { get; set; }
        public string Vari_Nkbn_no { get; set; }
        public string Vari_Gak { get; set; }
        public string Vari_Tanto_no { get; set; }
        public string Vari_Jisya_no { get; set; }
        public string Vari_Yatin_fkom_kozano { get; set; }
        public string Vari_Fkae_rirekino { get; set; }
        public string Vari_Ns_no { get; set; }
        public string Vari_Ns_bunkatu_no { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Cvpay_no { get; set; }
        public string Vari_Kokyaku_bango { get; set; }
        public string Vari_Fkae_no { get; set; }

    }

    #endregion

    #region 運用開始時未収滞納金情報モデル

    public class Unyotainodata_Model
    {

        public string Vari_Unyotaino_guid { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Gt_ym { get; set; }
        public string Vari_Sq_simeymd { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkbn_yotei { get; set; }
        public string Vari_Yotei_yatin_kozano { get; set; }
        public string Vari_Sqsaki_no { get; set; }
        public string Vari_Sq_gak { get; set; }
        public string Vari_Sq_zeikbn { get; set; }
        public string Vari_Sq_zeigak { get; set; }
        public string Vari_So_gak { get; set; }
        public string Vari_So_zeikbn { get; set; }
        public string Vari_So_zeigak { get; set; }
        public string Vari_Sosumi_flg { get; set; }
        public string Vari_Kanrigak_rit { get; set; }
        public string Vari_Sosaki_no { get; set; }
        public string Vari_Sokoza_no { get; set; }
        public string Vari_So_rit { get; set; }
        public string Vari_Soymd_gaitoukbn { get; set; }
        public string Vari_Soymd_simekbn { get; set; }
        public string Vari_Soymd_sokintukikbn { get; set; }
        public string Vari_Soymd_sokinsimekbn { get; set; }
        public string Vari_Zei_rit { get; set; }
        public string Vari_Sqbuild_flg { get; set; }
        public string Vari_Tanto_no { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Kanri_taxflg { get; set; }
        public string Vari_Kanritesu_zeiumukbn { get; set; }
        public string Vari_Kanri_utizeiflg { get; set; }

    }

    #endregion

    #region 請求情報モデル

    public class Sqdata_Model
    {

        public string Vari_Sqmei_guid { get; set; }
        public string Vari_Gt_ym { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Dispnkin_name { get; set; }
        public string Vari_Nkbn_yotei { get; set; }
        public string Vari_Sq_gak { get; set; }
        public string Vari_Sq_zeikbn { get; set; }
        public string Vari_Sq_zeigak { get; set; }
        public string Vari_Sq_simeymd { get; set; }
        public string Vari_Multisqsaki_flg { get; set; }
        public string Vari_Sqsaki_kbn { get; set; }
        public string Vari_Sqsaki_no { get; set; }
        public string Vari_Yotei_yatin_kozano { get; set; }
        public string Vari_Fkae_rirekino { get; set; }
        public string Vari_Fkae_rirekimeino { get; set; }
        public string Vari_Fkae_sqtaisyoflg { get; set; }
        public string Vari_Fkae_sqselectflg { get; set; }
        public string Vari_Fkae_sqflg { get; set; }
        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Tuki_kbn { get; set; }
        public string Vari_Nkin_recno { get; set; }
        public string Vari_Nkin_kbn { get; set; }
        public string Vari_Kynkin_guid { get; set; }
        public string Vari_Nextky_flg { get; set; }
        public string Vari_Hendosq_guid { get; set; }
        public string Vari_Hendo_kbn { get; set; }
        public string Vari_Kensin_ymd { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_So_gak { get; set; }
        public string Vari_So_zeikbn { get; set; }
        public string Vari_So_zeigak { get; set; }
        public string Vari_Etcsq_flg { get; set; }
        public string Vari_Etcsq_sorit { get; set; }
        public string Vari_Etcsq_soymd { get; set; }
        public string Vari_Bunkatu_guid { get; set; }
        public string Vari_Sh_kbn { get; set; }
        public string Vari_Shsaki_kbn { get; set; }
        public string Vari_Shsaki_no { get; set; }
        public string Vari_Shsaki_kozano { get; set; }
        public string Vari_Ikkatu_nkin_bango { get; set; }
        public string Vari_Jisya_no { get; set; }
        public string Vari_Tanto_no { get; set; }
        public string Vari_Kjsqkmk_guid { get; set; }
        public string Vari_Koteirule_guid { get; set; }
        public string Vari_Kojorule_flg { get; set; }
        public string Vari_Sq_hakkoyoteiymd { get; set; }
        public string Vari_Sqsort_no { get; set; }
        public string Vari_Tateazu_flg { get; set; }
        public string Vari_Tateazu_guid { get; set; }
        public string Vari_Sikikinzuiji_guid { get; set; }
        public string Vari_Unyotaino_guid { get; set; }
        public string Vari_Sokaisyu_guid { get; set; }
        public string Vari_Szen_kbn { get; set; }
        public string Vari_Szen_no { get; set; }
        public string Vari_Szen_sqno { get; set; }
        public string Vari_Zei_rit { get; set; }
        public string Vari_Edit_flg { get; set; }
        public string Vari_Lock_flg { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Tatekae_siwakegak { get; set; }
        public string Vari_Tatekae_siwakezeigak { get; set; }
        public string Vari_Sq_torikomiymd { get; set; }
        public string Vari_Cvpay_no { get; set; }
        public string Vari_Kokyaku_bango { get; set; }
        public string Vari_Cvpay_sqflg { get; set; }
        public string Vari_Etcsq_kanritesukbn { get; set; }
        public string Vari_Misyu_siwake_flg { get; set; }
        public string Vari_Zatusyu_kbn { get; set; }
        public string Vari_Shdirect_flg { get; set; }         // 20160519 EXEUpdateに伴う修正 請求情報 -add
        public string Vari_Koteishrule_guid { get; set; }     // 20160519 EXEUpdateに伴う修正 請求情報 -add
        public string Vari_Kojorule_taisyokbn { get; set; }   // 20160519 EXEUpdateに伴う修正 請求情報 -add
    }

    #endregion

    #region 請求消込情報モデル

    public class Sqdata_sqkesi_Model
    {

        public string Vari_Sqmei_guid { get; set; }
        public string Vari_Sqkesi_recno { get; set; }
        public string Vari_Kesi_ymd { get; set; }
        public string Vari_Nkbn_jitu { get; set; }
        public string Vari_Kesi_gak { get; set; }
        public string Vari_Kesi_zeigak { get; set; }
        public string Vari_Sui_guid { get; set; }
        public string Vari_Hikiazu_guid { get; set; }
        public string Vari_Idoazu_guid { get; set; }
        public string Vari_Printazuryo_guid { get; set; }
        public string Vari_Siwake_flg { get; set; }
        public string Vari_Tanto_no { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 送金予定情報モデル

    public class Sqdata_so_Model
    {

        public string Vari_Sqmei_guid { get; set; }
        public string Vari_So_recno { get; set; }
        public string Vari_So_gak { get; set; }
        public string Vari_So_zeikbn { get; set; }
        public string Vari_So_zeigak { get; set; }
        public string Vari_Edit_flg { get; set; }
        public string Vari_Lock_flg { get; set; }
        public string Vari_Hosyo_flg { get; set; }
        public string Vari_So_yoteiymd { get; set; }
        public string Vari_So_kakuymd { get; set; }
        public string Vari_Sokotik_guid { get; set; }
        public string Vari_Sorule_guid { get; set; }
        public string Vari_Sorule_no { get; set; }
        public string Vari_Sosaki_sorule_guid { get; set; }
        public string Vari_Sosaki_sorule_no { get; set; }
        public string Vari_Sokin_rit { get; set; }
        public string Vari_Calc_sqgak { get; set; }
        public string Vari_Sokaisyu_kbn { get; set; }
        public string Vari_Sokaisyu_guid { get; set; }
        public string Vari_Siwake_flg { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 送金管理情報モデル

    public class Sokanri_Model
    {

        public string Vari_Sokanri_guid { get; set; }
        public string Vari_Sorule_guid { get; set; }
        public string Vari_Sorule_no { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_So_ymd { get; set; }
        public string Vari_Sosaki_kbn { get; set; }
        public string Vari_Sosaki_no { get; set; }
        public string Vari_Sokoza_no { get; set; }
        public string Vari_So_gak { get; set; }
        public string Vari_Sokotik_guid { get; set; }
        public string Vari_Karikaku_flg { get; set; }
        public string Vari_Sgf_no { get; set; }
        public string Vari_Error_kbn { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Karikaku_guid { get; set; }
        public string Vari_Shdirect_flg { get; set; }

    }

    #endregion

    #region 変動費検針情報モデル

    public class Hendodata_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Hendo_kbn { get; set; }
        public string Vari_Kensin_ymd { get; set; }
        public string Vari_Gt_ym { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 変動費検針情報モデル (明細)

    public class Hendodata_meisai_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Hendo_kbn { get; set; }
        public string Vari_Kensin_ymd { get; set; }
        public string Vari_Gt_ym { get; set; }
        public string Vari_Kensin_ptn { get; set; }
        public string Vari_Hendo_guid { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_Hendosq_guid { get; set; }
        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Sqsaki_no { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkbn_no { get; set; }
        public string Vari_Sqsime_ymd { get; set; }
        public string Vari_Zenkai_ymd { get; set; }
        public string Vari_Zenkai_metervalue { get; set; }
        public string Vari_Zenkai_siyoryo { get; set; }
        public string Vari_Zenkai_sqgak { get; set; }
        public string Vari_Zenkai_sqzeigak { get; set; }
        public string Vari_Kensin_metervalue { get; set; }
        public string Vari_Kensin_siyoryo { get; set; }
        public string Vari_Kensin_ryokin1 { get; set; }
        public string Vari_Kensin_ryokin2 { get; set; }
        public string Vari_Kensin_ryokin3 { get; set; }
        public string Vari_Kensin_ryokin4 { get; set; }
        public string Vari_Kensin_ryokin5 { get; set; }
        public string Vari_Kensin_sqgak { get; set; }
        public string Vari_Kensin_sqzeigak { get; set; }
        public string Vari_Kensin_biko { get; set; }
        public string Vari_Kensin_soyoteiymd { get; set; }
        public string Vari_Kensin_sorit { get; set; }
        public string Vari_Kensin_sogak { get; set; }
        public string Vari_Kensin_sozeigak { get; set; }
        public string Vari_Zenkaihendosq_guid { get; set; }

    }

    #endregion

    #region 家主固定控除情報モデル

    public class Koteirule_Model
    {

        public string Vari_Koteirule_guid { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Nkin_sortorder { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkin_name { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_Sosai_flg { get; set; }
        public string Vari_Gak { get; set; }
        public string Vari_Zei_kbn { get; set; }
        public string Vari_Zei_gak { get; set; }
        public string Vari_Nkbn_yotei { get; set; }
        public string Vari_Yatin_kozano { get; set; }
        public string Vari_Gtstart_ym { get; set; }
        public string Vari_Gtend_ym { get; set; }
        public string Vari_Tateazu_flg { get; set; }
        public string Vari_Sq_mmkbn { get; set; }
        public string Vari_Sq_ptn { get; set; }
        public string Vari_Sq_interval { get; set; }
        public string Vari_Sq_nen { get; set; }
        public string Vari_Sq_tuki { get; set; }
        public string Vari_Zei_rit { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Sq_simekbn { get; set; }
        public string Vari_Ryosyu_flg { get; set; }

    }

    #endregion

    #region 家主請求控除情報モデル

    public class Kjdata_Model
    {

        public string Vari_Kj_guid { get; set; }
        public string Vari_Kjkmk_guid { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_Kj_yoteiymd { get; set; }
        public string Vari_Kjsort_no { get; set; }
        public string Vari_Gt_ym { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Dispnkin_name { get; set; }
        public string Vari_Kj_gak { get; set; }
        public string Vari_Kj_zeikbn { get; set; }
        public string Vari_Kj_zeigak { get; set; }
        public string Vari_Lock_flg { get; set; }
        public string Vari_Ryosyu_flg { get; set; }
        public string Vari_Tateazu_flg { get; set; }
        public string Vari_Tateazu_guid { get; set; }
        public string Vari_Bunkatu_guid { get; set; }
        public string Vari_Koteirule_guid { get; set; }
        public string Vari_Kojorule_flg { get; set; }
        public string Vari_So_kakuymd { get; set; }
        public string Vari_Sokotik_guid { get; set; }
        public string Vari_Sorule_guid { get; set; }
        public string Vari_Sorule_no { get; set; }
        public string Vari_Sosaki_sorule_guid { get; set; }
        public string Vari_Sosaki_sorule_no { get; set; }
        public string Vari_Sosaki_no { get; set; }
        public string Vari_Sokoza_no { get; set; }
        public string Vari_Sikikinzuiji_guid { get; set; }
        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Sokaisyu_guid { get; set; }
        public string Vari_Kanritesu_kbn { get; set; }
        public string Vari_Kanritesu_cyosyukbn { get; set; }
        public string Vari_Kanrigak_rit { get; set; }
        public string Vari_Szen_kbn { get; set; }
        public string Vari_Szen_no { get; set; }
        public string Vari_Szen_sqno { get; set; }
        public string Vari_Siwake_flg { get; set; }
        public string Vari_Zei_rit { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_Kj_torikomiymd { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Kubungrp_guid { get; set; }
        public string Vari_Kojorule_taisyokbn { get; set; }        // 20160519 EXEUpdateに伴う修正 家主請求控除情報

    }

    #endregion

}