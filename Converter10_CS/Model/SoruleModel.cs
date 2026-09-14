
namespace Converter10.Njc.Model
{

    #region 送金ルール基本情報モデル

    public class Sorule_Model
    {

        public string Vari_Sorule_guid { get; set; }
        public string Vari_Sorule_no { get; set; }
        public string Vari_Relation_guid { get; set; }
        public string Vari_Kubunsyo { get; set; }
        public string Vari_Sorule_startymd { get; set; }
        public string Vari_Sorule_endymd { get; set; }
        public string Vari_Kanri_keitaikbn { get; set; }
        public string Vari_Ikkatu_kanriflg { get; set; }
        public string Vari_Cyukai_kyflg { get; set; }
        public string Vari_Cyukai_koflg { get; set; }
        public string Vari_Cyukai_kaiflg { get; set; }
        public string Vari_Soymd1_gaitoukbn { get; set; }
        public string Vari_Soymd1_simekbn { get; set; }
        public string Vari_Soymd1_sokintukikbn { get; set; }
        public string Vari_Soymd1_sokinsimekbn { get; set; }
        public string Vari_Soymd2_gaitoukbn { get; set; }
        public string Vari_Soymd2_simekbn { get; set; }
        public string Vari_Soymd2_sokintukikbn { get; set; }
        public string Vari_Soymd2_sokinsimekbn { get; set; }
        public string Vari_Soymd3_gaitoukbn { get; set; }
        public string Vari_Soymd3_simekbn { get; set; }
        public string Vari_Soymd3_sokintukikbn { get; set; }
        public string Vari_Soymd3_sokinsimekbn { get; set; }
        public string Vari_Soymd4_gaitoukbn { get; set; }
        public string Vari_Soymd4_simekbn { get; set; }
        public string Vari_Soymd4_sokintukikbn { get; set; }
        public string Vari_Soymd4_sokinsimekbn { get; set; }
        public string Vari_Soymd5_gaitoukbn { get; set; }
        public string Vari_Soymd5_simekbn { get; set; }
        public string Vari_Soymd5_sokintukikbn { get; set; }
        public string Vari_Soymd5_sokinsimekbn { get; set; }
        public string Vari_Sosaki_multikbn { get; set; }
        public string Vari_Sosaki_koteiflg { get; set; }
        public string Vari_Sosaki_koteisu { get; set; }
        public string Vari_Sosaki_anbunflg { get; set; }
        public string Vari_Sosaki_hasuuketorisaki { get; set; }
        public string Vari_Sosaki_hasuadjustmentflg { get; set; }
        public string Vari_Ikkatu_kbn { get; set; }
        public string Vari_Ikkatu_bkgak { get; set; }
        public string Vari_Ikkatu_menseki { get; set; }
        public string Vari_Ikkatu_mensekimonth { get; set; }
        public string Vari_Kanritesu_kbn { get; set; }
        public string Vari_Kanritesu_cyosyukbn { get; set; }
        public string Vari_Kanri_reigaiky { get; set; }
        public string Vari_Kanri_reigaikai { get; set; }
        public string Vari_Hyteigaku_kbn { get; set; }
        public string Vari_Hyteigaku_kanrigak { get; set; }
        public string Vari_Hyteigaku_hiwariflg { get; set; }
        public string Vari_Bkteigaku_kanrigak { get; set; }
        public string Vari_Kanri_taxflg { get; set; }
        public string Vari_Biko_basic { get; set; }
        public string Vari_Sh_daihyohyflg { get; set; }
        public string Vari_Sh_kozabetuflg { get; set; }
        public string Vari_Sh_kozabetuListNokbn { get; set; }
        public string Vari_Hyikkatu_nkin_no1kbn { get; set; }
        public string Vari_Hyikkatu_nkin_no2kbn { get; set; }
        public string Vari_Hyikkatu_nkin_no3kbn { get; set; }
        public string Vari_Hyikkatu_nkin_no4kbn { get; set; }
        public string Vari_Hyikkatu_nkin_no5kbn { get; set; }
        public string Vari_Kanritesu_zeiumukbn { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_No_soruleflg { get; set; }
        public string Vari_Cyukai_zuijisokin { get; set; }
        public string Vari_Ikkatu_calckbn { get; set; }
        public string Vari_Ikkatu_taxflg { get; set; }
        public string Vari_Kanri_utizeiflg { get; set; }
        public string Vari_Soymd_basis { get; set; }

    }

    #endregion

    #region 送金ルール送金先情報モデル

    public class Sorule_sosaki_Model
    {

        public string Vari_Sorule_guid { get; set; }
        public string Vari_Sorule_no { get; set; }
        public string Vari_Sosaki_recno { get; set; }
        public string Vari_So_ow_no { get; set; }
        public string Vari_So_ow_kozano { get; set; }
        public string Vari_So_rit { get; set; }
        public string Vari_So_fixsogak { get; set; }

    }

    #endregion

    #region 送金ルール入金項目情報モデル

    public class Sorule_nk_cmrule_Model
    {

        public string Vari_Sorule_guid { get; set; }
        public string Vari_Sorule_no { get; set; }
        public string Vari_Taisyokbn { get; set; }
        public string Vari_Nkin_sortorder { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Sokin_rit { get; set; }
        public string Vari_Kanrigak_rit { get; set; }
        public string Vari_Hosyo_flg { get; set; }
        public string Vari_So_no { get; set; }

    }

    #endregion

    #region 送金ルール控除項目情報モデル

    public class Sorule_kojo_cmrule_Model
    {

        public string Vari_Sorule_guid { get; set; }
        public string Vari_Sorule_no { get; set; }
        public string Vari_Kojo_taisyokbn { get; set; }
        public string Vari_Nkin_sortorder { get; set; }
        public string Vari_Taisyonkin_no { get; set; }
        public string Vari_Sosai_flg { get; set; }
        public string Vari_Kojocalc_kbn { get; set; }
        public string Vari_Kojo_gak { get; set; }
        public string Vari_Kojo_gakzeikbn { get; set; }
        public string Vari_Kojonkin_no { get; set; }
        public string Vari_Kojo_rit { get; set; }
        public string Vari_Kojo_ritzeikbn { get; set; }
        public string Vari_Kojo_ritutizei { get; set; }
        public string Vari_Zei_rit { get; set; }
        public string Vari_Tateazu_flg { get; set; }
        public string Vari_Kojo_zeigak { get; set; }
        public string Vari_Sotaisyo_flg { get; set; }
        public string Vari_Sotaisyonkin_no { get; set; }

    }

    #endregion

}