
namespace Converter10.Njc.Model
{

    #region 契約基本情報モデル

    public class Kydata_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_Ky_no { get; set; }
        public string Vari_Ky_deleteflg { get; set; }
        public string Vari_Delete_guid { get; set; }
        public string Vari_Delete_day { get; set; }
        public string Vari_Delete_cnt { get; set; }
        public string Vari_Syokai_kyymd { get; set; }
        public string Vari_Status { get; set; }
        public string Vari_Cancelriyu { get; set; }
        public string Vari_Status_ymd { get; set; }
        public string Vari_Cyukai_gy_fudono { get; set; }
        public string Vari_Syunin_logonuser_no { get; set; }
        public string Vari_Tetuke_gak1 { get; set; }
        public string Vari_Tetuke_ymd1 { get; set; }
        public string Vari_Tetuke_biko1 { get; set; }
        public string Vari_Kaiyaku_flg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Movefrom_kyguid { get; set; }
        public string Vari_Moveto_kyguid { get; set; }
        public string Vari_Svbunrui_no { get; set; }
        public string Vari_Krbunrui_no { get; set; }
        public string Vari_Kaiyaku_uketukekbn { get; set; }
        public string Vari_Kaiyaku_months { get; set; }
        public string Vari_Kaiyaku_days { get; set; }
        public string Vari_Kaiyaku_day { get; set; }
        public string Vari_Cyukai_tantoname { get; set; }
        public string Vari_Cyukai_tantonamesjis { get; set; }

    }

    #endregion

    #region 契約履歴情報モデル

    public class Kydata_kihon_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Ko_no { get; set; }
        public string Vari_Henko_no { get; set; }
        public string Vari_Sqdata_guid { get; set; }
        public string Vari_Ky_bango { get; set; }
        public string Vari_Ky_ymd { get; set; }
        public string Vari_Kystart_ymd { get; set; }
        public string Vari_Kyend_ymd { get; set; }
        public string Vari_Henko_ymd { get; set; }
        public string Vari_Tuti_ymd { get; set; }
        public string Vari_Print_ymd { get; set; }
        public string Vari_Kanri_gy_fudono { get; set; }
        public string Vari_Gy_hosyono { get; set; }
        public string Vari_Hosyo_naiyo { get; set; }
        public string Vari_Kokyaku_bango { get; set; }
        public string Vari_Kyrui_no { get; set; }
        public string Vari_Siyo_mokuteki { get; set; }
        public string Vari_Kosin_umu { get; set; }
        public string Vari_Yatin_kbn { get; set; }
        public string Vari_Fkae_startym { get; set; }
        public string Vari_Fkae_willstartflg { get; set; }
        public string Vari_Yatin_kozakbn { get; set; }
        public string Vari_Yatin_kozano { get; set; }
        public string Vari_Maitukiyatin_kozano { get; set; }
        public string Vari_Yokugetu_uketoriflg { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_Tougetu_sagakuflg { get; set; }
        public string Vari_Nextky_startymd { get; set; }
        public string Vari_Nextky_endymd { get; set; }
        public string Vari_Kosin_hiwariflg { get; set; }
        public string Vari_Sokojorule_kbn { get; set; }
        public string Vari_Soyotei_ymdflg { get; set; }
        public string Vari_Soyotei_ymd { get; set; }
        public string Vari_Kanritesu_flg { get; set; }
        public string Vari_Kanritesu_gak { get; set; }
        public string Vari_Nextnkinset_kbn { get; set; }
        public string Vari_Sqdata_startymd { get; set; }
        public string Vari_Sqdata_endymd { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Biko2 { get; set; }
        public string Vari_Hoken_biko { get; set; }
        public string Vari_Nkinsime_ymd { get; set; }
        public string Vari_Yatin_jisansaki { get; set; }
        public string Vari_Hoken_kikan { get; set; }
        public string Vari_Hoken_gak { get; set; }
        public string Vari_Confirmky_sekininsya { get; set; }
        public string Vari_Confirmky_ymd { get; set; }
        public string Vari_Confirmky_print { get; set; }
        public string Vari_Confirmkai_sekininsya { get; set; }
        public string Vari_Confirmkai_ymd { get; set; }
        public string Vari_Confirmkai_print { get; set; }
        public string Vari_Hikiuke_name { get; set; }
        public string Vari_Hikiuke_addr { get; set; }
        public string Vari_Hikiuke_tel { get; set; }
        public string Vari_Kohokennkin_ymd { get; set; }
        public string Vari_Kokanryo_ymd { get; set; }
        public string Vari_Kokanryotuti_ymd { get; set; }
        public string Vari_Kotuti_ymd { get; set; }
        public string Vari_Kosaisoku_ymd { get; set; }
        public string Vari_Kosyoruiuke_ymd { get; set; }
        public string Vari_Kokairenraku_umu { get; set; }
        public string Vari_Kokairenraku_ymd { get; set; }
        public string Vari_Kokairenraku_logonuser_no { get; set; }
        public string Vari_Kokairenraku_biko { get; set; }
        public string Vari_Maitukisq_umu { get; set; }
        public string Vari_Hikiuke_namesjis { get; set; }
        public string Vari_Kofkae_ymd { get; set; }
        public string Vari_Yokugetu_uketorimonth { get; set; }
        public string Vari_Ky_logonuser_no { get; set; }
        public string Vari_Sq_logonuser_no { get; set; }
        public string Vari_Torihiki_tesumoto { get; set; }            // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add
        public string Vari_Torihiki_tesukyaku { get; set; }           // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add
        public string Vari_Torihiki_tesugak { get; set; }             // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add
        public string Vari_Torihiki_tesuzeikbn { get; set; }          // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add
        public string Vari_Torihiki_tesuzeigak { get; set; }          // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更以外) -add

    }

    #endregion

    #region 契約車情報モデル

    public class Kydata_car_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Car_cnt { get; set; }
        public string Vari_Carmaker { get; set; }
        public string Vari_Carname { get; set; }
        public string Vari_Carcolor { get; set; }
        public string Vari_Carnumber { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Car_kukaku { get; set; }

    }

    #endregion

    #region 契約契約者情報モデル

    public class Kydata_kys_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Kys_cnt { get; set; }
        public string Vari_Kys_no { get; set; }
        public string Vari_Nyukyo_flg { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 契約入居者情報モデル

    public class Kydata_nyukyo_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Nyukyosya_cnt { get; set; }
        public string Vari_Kys_no { get; set; }
        public string Vari_Name { get; set; }
        public string Vari_Nameu { get; set; }
        public string Vari_Kana { get; set; }
        public string Vari_Gender { get; set; }
        public string Vari_Aidagara { get; set; }
        public string Vari_Birthday { get; set; }
        public string Vari_Kinmusaki { get; set; }
        public string Vari_Tel { get; set; }
        public string Vari_Mobiletel { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 契約保証人情報モデル

    public class Kydata_hosyonin_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Hosyonin_cnt { get; set; }
        public string Vari_Kys_no { get; set; }
        public string Vari_Hosyonin_no { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 契約保険情報モデル

    public class Kydata_hoken_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Hoken_no { get; set; }
        public string Vari_Gy_hokenno { get; set; }
        public string Vari_Ky_ymd { get; set; }
        public string Vari_Kystart_ymd { get; set; }
        public string Vari_Kyend_ymd { get; set; }
        public string Vari_Hoken_gak { get; set; }
        public string Vari_Mankituti_flg { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Syoken_bango { get; set; }

    }

    #endregion

    #region 契約特約事項情報モデル

    public class Kydata_tokuyaku_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Tokuyaku_grpno { get; set; }
        public string Vari_Naiyo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 契約メモ情報モデル

    public class Kydata_memo_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 契約入金項目情報モデル

    public class Kydata_nkin_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Tuki_kbn { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkin_recno { get; set; }
        public string Vari_Nkin_sortorder { get; set; }
        public string Vari_Nkin_kbn { get; set; }
        public string Vari_Sq_gak { get; set; }
        public string Vari_Sq_zeikbn { get; set; }
        public string Vari_Sq_zeigak { get; set; }
        public string Vari_Calc_kbn { get; set; }
        public string Vari_Calc_nkinno { get; set; }
        public string Vari_Calc_monthcnt { get; set; }
        public string Vari_Sqsaki_no { get; set; }
        public string Vari_Nkbn_yotei { get; set; }
        public string Vari_Sq_mmkbn { get; set; }
        public string Vari_Frstart_ymd { get; set; }
        public string Vari_Frend_ymd { get; set; }
        public string Vari_Frsq_gak { get; set; }
        public string Vari_Sqstart_ymd { get; set; }
        public string Vari_Sq_ptn { get; set; }
        public string Vari_Sq_interval { get; set; }
        public string Vari_Sq_nen { get; set; }
        public string Vari_Sq_tuki { get; set; }
        public string Vari_Zei_rit { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_Nkin_guid { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Fr_kbn { get; set; }

    }

    #endregion

    #region 契約次回入金項目情報モデル

    public class Kydata_nkin_nx_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Tuki_kbn { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkin_recno { get; set; }
        public string Vari_Nkin_sortorder { get; set; }
        public string Vari_Nkin_kbn { get; set; }
        public string Vari_Sq_gak { get; set; }
        public string Vari_Sq_zeikbn { get; set; }
        public string Vari_Sq_zeigak { get; set; }
        public string Vari_Calc_kbn { get; set; }
        public string Vari_Calc_nkinno { get; set; }
        public string Vari_Calc_monthcnt { get; set; }
        public string Vari_Sqsaki_no { get; set; }
        public string Vari_Nkbn_yotei { get; set; }
        public string Vari_Sq_mmkbn { get; set; }
        public string Vari_Sqstart_ymd { get; set; }
        public string Vari_Sq_ptn { get; set; }
        public string Vari_Sq_interval { get; set; }
        public string Vari_Sq_nen { get; set; }
        public string Vari_Sq_tuki { get; set; }
        public string Vari_Zei_rit { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 契約変動費各戸メーター情報モデル

    public class Kydata_hendo_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Hyhendo_guid { get; set; }
        public string Vari_Hendo_sortorder { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_Hendo_kbn { get; set; }
        public string Vari_Meter_name { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Hendorule_no { get; set; }
        public string Vari_Hendorule_biko { get; set; }
        public string Vari_Sqsaki_no { get; set; }
        public string Vari_Nkbn_yotei { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Sq_mmkbn { get; set; }

    }

    #endregion

    #region 契約控除ルール情報モデル

    public class Kydata_kojorule_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Taisyokbn { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkin_sortorder { get; set; }
        public string Vari_Sosai_flg { get; set; }
        public string Vari_Calc_kbn { get; set; }
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

    #region 契約送金ルール情報モデル

    public class Kydata_sorule_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Taisyokbn { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkin_sortorder { get; set; }
        public string Vari_Sokin_rit { get; set; }
        public string Vari_Kanrigak_rit { get; set; }
        public string Vari_Hosyo_flg { get; set; }

    }

    #endregion

    #region 契約解約情報モデル

    public class Kydata_kai_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Kai_ymd { get; set; }
        public string Vari_Seisan_ymd { get; set; }
        public string Vari_Seisan_completeymd { get; set; }
        public string Vari_Uketuke_ymd { get; set; }
        public string Vari_Riyu { get; set; }
        public string Vari_Tatiai_ymd { get; set; }
        public string Vari_Uketuke_logonuser_no { get; set; }
        public string Vari_Tatiai_logonuser_no { get; set; }
        public string Vari_Seisan_logonuser_no { get; set; }
        public string Vari_Szen_logonuser_no { get; set; }
        public string Vari_Szen_jisyano { get; set; }
        public string Vari_Szen_endymd { get; set; }
        public string Vari_Szen_kojiyoteistartymd { get; set; }
        public string Vari_Szen_kojiyoteiendymd { get; set; }
        public string Vari_Szen_kojibasyo { get; set; }
        public string Vari_Szen_kojigaiyo { get; set; }
        public string Vari_Next_name { get; set; }
        public string Vari_Next_namesjis { get; set; }
        public string Vari_Next_postcode { get; set; }
        public string Vari_Next_addr1 { get; set; }
        public string Vari_Next_addr2 { get; set; }
        public string Vari_Next_tel1 { get; set; }
        public string Vari_Yatin_kozano { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_Biko_tatiai { get; set; }
        public string Vari_Seisan_henkinymd { get; set; }
        public string Vari_Seisan_sqymd { get; set; }
        public string Vari_Kanritesu_flg { get; set; }
        public string Vari_Kanritesu_gak { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_History { get; set; }
        public string Vari_So_yoteiymd { get; set; }
        public string Vari_Tatiai { get; set; }
        public string Vari_Tatiai_yoteiymd { get; set; }
        public string Vari_Bosyu_jokenymd { get; set; }
        public string Vari_Ow_logonuser_no { get; set; }
        public string Vari_Bosyujoken { get; set; }
        public string Vari_Confirmkai_sekininsya { get; set; }
        public string Vari_Confirmkai_ymd { get; set; }
        public string Vari_Confirmkai_print { get; set; }
        // Public Property Vari_Szen_jisyafutanumuflg As String
        // Public Property Vari_Szen_jisyafutangak As String
        // Public Property Vari_Szen_jisyafutanzeikbn As String
        // Public Property Vari_Szen_jisyafutanzeigak As String
        // Public Property Vari_Szen_bikojisyafutan As String
        // Public Property Vari_Szen_sonotafutanumuflg As String
        // Public Property Vari_Szen_sonotafutangak As String
        // Public Property Vari_Szen_sonotafutanzeikbn As String
        // Public Property Vari_Szen_sonotafutanzeigak As String
        // Public Property Vari_Szen_sonotafutanname As String
        // Public Property Vari_Szen_bikosonotafutan As String
        public string Vari_Tatiai_yoteitime { get; set; }
        public string Vari_Next_keisyo { get; set; }
        public string Vari_Nkin_kaiyakutukinoprintflg { get; set; }      // 20160519 EXEUpdateに伴う修正 契約解約情報 -add
        public string Vari_Szen_kysno { get; set; }                       // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_kysnkbn { get; set; }                     // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_kyssorit { get; set; }                    // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_kyssogak { get; set; }                    // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owno { get; set; }                        // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owkaisyukbn { get; set; }                 // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owsosakisoruleguid { get; set; }          // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owsosakisoruleno { get; set; }            // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owsokozano { get; set; }                  // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owsqsimeymd { get; set; }                 // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_ownkbn { get; set; }                      // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owyatinkozano { get; set; }               // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owsqsakino { get; set; }                  // 20160620 EXEUpdateに伴う修正2 -add
        public string Vari_Szen_owkojoymd { get; set; }                   // 20160829 革命10バージョンアップに伴う修正 -add
        public string Vari_Szen_owsqymd { get; set; }                     // 20160829 革命10バージョンアップに伴う修正 -add

    }

    #endregion


    #region 契約修繕見積情報モデル

    public class Kydata_kaiszen_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Szen_mituno { get; set; }
        public string Vari_Mitu_title { get; set; }
        public string Vari_Mitu_bango { get; set; }
        public string Vari_Mitu_ymd { get; set; }
        public string Vari_Seiyaku_flg { get; set; }
        public string Vari_Seiyaku_ymd { get; set; }
        public string Vari_Futan_kbn { get; set; }
        public string Vari_Kys_futanrit { get; set; }
        public string Vari_Ow_futanrit { get; set; }
        public string Vari_Jisya_futanrit { get; set; }
        public string Vari_Zei_kbn { get; set; }
        public string Vari_Gokei_zeikbn { get; set; }
        public string Vari_Gokei_zeirit { get; set; }
        public string Vari_Kys_gokeizeigak { get; set; }
        public string Vari_Ow_gokeizeigak { get; set; }
        public string Vari_Jisya_gokeizeigak { get; set; }
        public string Vari_Hasu_futankbn { get; set; }

    }

    #endregion

    #region 契約修繕見積詳細情報モデル

    public class Kydata_kaiszenmeisai_Model
    {

        public string Vari_Ky_guid { get; set; }
        public string Vari_Ky_recno { get; set; }
        public string Vari_Szen_mituno { get; set; }
        public string Vari_Szen_meisaino { get; set; }
        public string Vari_Szen_name { get; set; }
        public string Vari_Tekiyo { get; set; }
        public string Vari_Mitu_suryo { get; set; }
        public string Vari_Mitu_tani { get; set; }
        public string Vari_Mitu_tanka { get; set; }
        public string Vari_Mitu_zeikbn { get; set; }
        public string Vari_Mitu_zeigak { get; set; }
        public string Vari_Kys_futanrit { get; set; }
        public string Vari_Ow_futanrit { get; set; }
        public string Vari_Jisya_futanrit { get; set; }
        public string Vari_Szen_gyno { get; set; }
        public string Vari_Jikko_suryo { get; set; }
        public string Vari_Jikko_tani { get; set; }
        public string Vari_Jikko_tanka { get; set; }
        public string Vari_Jikko_zeikbn { get; set; }
        public string Vari_Jikko_zeigak { get; set; }

    }

    #endregion

}