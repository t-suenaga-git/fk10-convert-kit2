
namespace Converter10.Njc.Model
{

    #region 物件基本情報モデル

    public class Bkdata_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Bk_no { get; set; }
        public string Vari_Bk_deleteflg { get; set; }
        public string Vari_Delete_guid { get; set; }
        public string Vari_Delete_day { get; set; }
        public string Vari_Delete_cnt { get; set; }
        public string Vari_Bk_name { get; set; }
        public string Vari_Bk_deletename { get; set; }
        public string Vari_Bk_kana { get; set; }
        public string Vari_Tatemono_sikibetu { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr_kenno { get; set; }
        public string Vari_Addr_sino { get; set; }
        public string Vari_Addr_cyo { get; set; }
        public string Vari_Addr_cyome { get; set; }
        public string Vari_Addr_cyomeptn { get; set; }
        public string Vari_Addr_banti { get; set; }
        public string Vari_Addr_etc { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Bk_namesjis { get; set; }
        public string Vari_Bk_gaibuno { get; set; }
        public string Vari_Delete_cause { get; set; }
        public string Vari_Lastupdate { get; set; }       // 20160519 EXEUpdateに伴う修正 物件基本情報 -add

    }

    #endregion

    #region 物件外部No情報

    public class Bkdata_gaibuno_Model
    {

        public string Vari_Current_no { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 物件詳細情報モデル

    public class Bkdata_detail_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Bk_ruinokbn { get; set; }
        public string Vari_Gps_wgsido { get; set; }
        public string Vari_Gps_wgskeido { get; set; }
        public string Vari_Gps_tokyoido { get; set; }
        public string Vari_Gps_tokyokeido { get; set; }
        public string Vari_Bk_dentetuflg { get; set; }
        public string Vari_Syunko_ymd { get; set; }
        public string Vari_Kaidate { get; set; }
        public string Vari_Tika { get; set; }
        public string Vari_Elevator_flg { get; set; }
        public string Vari_Elevator_number { get; set; }
        public string Vari_Rooftop_flg { get; set; }
        public string Vari_Kozo_nokbn { get; set; }
        public string Vari_Kozo_yanekbn { get; set; }
        public string Vari_Moto_gy_fudono { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Kosu_total { get; set; }
        public string Vari_Men_nobeyuka { get; set; }
        public string Vari_Men_nobeyukatubo { get; set; }
        public string Vari_Men_sikiti { get; set; }
        public string Vari_Men_sikititubo { get; set; }
        public string Vari_Men_parking { get; set; }
        public string Vari_Men_parkingtubo { get; set; }
        public string Vari_Men_yukatoki { get; set; }
        public string Vari_Men_yukatokitubo { get; set; }
        public string Vari_Men_sikititoki { get; set; }
        public string Vari_Men_sikititokitubo { get; set; }
        public string Vari_Default_sqtuki { get; set; }
        public string Vari_Keiyakuyou_sime { get; set; }
        public string Vari_Yatin_jisansaki { get; set; }
        public string Vari_Yatin_kozano { get; set; }
        public string Vari_Yatin_kykozaflg { get; set; }
        public string Vari_Yatin_kykozano { get; set; }
        public string Vari_Kozo_taikakbn { get; set; }
        public string Vari_Gy_sekono { get; set; }
        public string Vari_Gy_hosyuno { get; set; }
        public string Vari_Kanri_hosiki { get; set; }
        public string Vari_Kanri_gyname { get; set; }
        public string Vari_Kanri_gytanto { get; set; }
        public string Vari_Kanri_gytel { get; set; }
        public string Vari_Kanrinin_gyomukeitai { get; set; }
        public string Vari_Kanrinin_tel { get; set; }
        public string Vari_Jisya_no { get; set; }
        public string Vari_Kotu_sonota1 { get; set; }
        public string Vari_Kotu_sonota1kyori { get; set; }
        public string Vari_Kotu_sonota2 { get; set; }
        public string Vari_Kotu_sonota2kyori { get; set; }
        public string Vari_Denki_gy_lifeno { get; set; }
        public string Vari_Water_gy_lifeno { get; set; }
        public string Vari_Gas_gy_lifeno { get; set; }
        public string Vari_Haisui_gy_lifeno { get; set; }
        public string Vari_Toyu_gy_lifeno { get; set; }
        public string Vari_Lifeline1_gy_lineno { get; set; }
        public string Vari_Lifeline2_gy_lineno { get; set; }
        public string Vari_Lifeline3_gy_lineno { get; set; }
        public string Vari_Kensingyomu_umu { get; set; }
        public string Vari_Kensinorder_kbn { get; set; }
        public string Vari_Parenthendo_useflg { get; set; }
        public string Vari_Parking_kanriflg { get; set; }
        public string Vari_Parking_car { get; set; }
        public string Vari_Parking_bike { get; set; }
        public string Vari_Parking_bicycle { get; set; }
        public string Vari_Parking_bicyclefreeflg { get; set; }
        public string Vari_Kubunsyo { get; set; }
        // Public Property Vari_Soymd_basis As String  '20160519 EXEUpdateに伴う修正 物件詳細情報 -del
        public string Vari_Isiwata_kirokukbn { get; set; }
        public string Vari_Isiwata_syokai1flg { get; set; }
        public string Vari_Isiwata_syokai2flg { get; set; }
        public string Vari_Isiwata_syokai3flg { get; set; }
        public string Vari_Isiwata_syokai4flg { get; set; }
        public string Vari_Isiwata_gy_sekono { get; set; }
        public string Vari_Isiwata_cyosaymd { get; set; }
        public string Vari_Isiwata_cyosakikankbn { get; set; }
        public string Vari_Isiwata_cyosahani { get; set; }
        public string Vari_Ishiwata_useflg { get; set; }
        public string Vari_Isiwata_usearea { get; set; }
        public string Vari_Isiwata_biko { get; set; }
        public string Vari_Taisin_sindanflg { get; set; }
        public string Vari_Taisin_syokai1flg { get; set; }
        public string Vari_Taisin_syokai2flg { get; set; }
        public string Vari_Taisin_syokai3flg { get; set; }
        public string Vari_Taisin_syorui1flg { get; set; }
        public string Vari_Taisin_syorui2flg { get; set; }
        public string Vari_Taisin_syorui3flg { get; set; }
        public string Vari_Taisin_biko { get; set; }
        public string Vari_Horei_dosyatiiki { get; set; }
        public string Vari_Horei_ruikbn { get; set; }
        public string Vari_Horei_naiyo { get; set; }
        public string Vari_Sikiti_riyoruikbn { get; set; }
        public string Vari_Sikiti_kystartymd { get; set; }
        public string Vari_Sikiti_kyendymd { get; set; }
        public string Vari_Sikiti_biko { get; set; }
        public string Vari_Nyukyoritu_startymd { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Hyothergyfudo_flg { get; set; }
        public string Vari_Gomi_hosoku { get; set; }
        public string Vari_Toki_ymd { get; set; }
        public string Vari_Syo_kenriflg { get; set; }
        public string Vari_Syo_kenrikbn { get; set; }
        public string Vari_Other_kenriflg { get; set; }
        public string Vari_Men_kentiku { get; set; }
        public string Vari_Men_kentikutubo { get; set; }
        public string Vari_Men_kentikutoki { get; set; }
        public string Vari_Men_kentikutokitubo { get; set; }
        public string Vari_Bk_logonuser_no { get; set; }
        public string Vari_Kanri_gyfax { get; set; }
        public string Vari_Homeelevator_flg { get; set; }
        public string Vari_Horei_dosyatokubetutiiki { get; set; }
        public string Vari_Horei_zoseitakutitiiki { get; set; }
        public string Vari_Horei_tunamitiiki { get; set; }
        public string Vari_Horei_dosyatiikibiko { get; set; }
        public string Vari_Horei_dosyatokubetutiikibiko { get; set; }
        public string Vari_Horei_zoseitakutitiikibiko { get; set; }
        public string Vari_Horei_tunamitiikibiko { get; set; }
        public string Vari_Svbunrui_no { get; set; }
        public string Vari_Krbunrui_no { get; set; }
        public string Vari_Kanrinin_name { get; set; }
        public string Vari_Kozo_other { get; set; }
        public string Vari_Kadoti_flg { get; set; }
        public string Vari_Cityplan { get; set; }
        public string Vari_Yototiki { get; set; }
        public string Vari_Kanrinin_namesjis { get; set; }
        public string Vari_Kanri_gytantosjis { get; set; }
        public string Vari_Kanri_gynamesjis { get; set; }
        public string Vari_Parking_caraki { get; set; }
        public string Vari_Parking_bikeaki { get; set; }
        public string Vari_Parking_bicycleaki { get; set; }
        public string Vari_Syogaku_name { get; set; }
        public string Vari_Syogaku_kyori { get; set; }
        public string Vari_Cyugaku_name { get; set; }
        public string Vari_Cyugaku_kyori { get; set; }
        public string Vari_Bk_area_no { get; set; }
        public string Vari_Emergencyelevator_flg { get; set; }   // 20160519 EXEUpdateに伴う修正 物件詳細情報 -del

    }

    #endregion

    #region 物件所有者情報

    public class Bkdata_syo_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Kn_no { get; set; }
        public string Vari_Sorule_guid { get; set; }
        public string Vari_Kasi1_ow_no { get; set; }
        public string Vari_Syo1_ow_no { get; set; }
        public string Vari_Kasi2_ow_no { get; set; }
        public string Vari_Syo2_ow_no { get; set; }
        public string Vari_Syo_startymd { get; set; }
        public string Vari_Syo_endymd { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 物件ゴミ情報モデル

    public class Bkdata_dust_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Dust_no { get; set; }
        public string Vari_Dust_rui { get; set; }
        public string Vari_Dust_youbi { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region ×物件画像情報

    public class Bkdata_gazo_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Bk_gazono { get; set; }
        public string Vari_Bk_gazocomment { get; set; }
        public string Vari_Bk_gazolastupdate { get; set; }
        public string Vari_Bk_gazofilesizekbyte { get; set; }
        public string Vari_Bk_gazowidth { get; set; }
        public string Vari_Bk_gazoheight { get; set; }
        public string Vari_Bk_gazothumbimage { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 物件変動費親メーター情報(汎用ツールのみ)

    public class Bkdata_hendo_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Bkhendo_guid { get; set; }
        public string Vari_Rec_no { get; set; }
        public string Vari_Hendo_kbn { get; set; }
        public string Vari_Meter_name { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Child_meter { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Sq_mmkbn { get; set; }
        public string Vari_Tani { get; set; }
        public string Vari_Tanisjis { get; set; }

    }

    #endregion

    #region 物件鍵情報(汎用ツールのみ)

    public class Bkdata_kagi_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Kagi_no { get; set; }
        public string Vari_Honsu { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Hokan { get; set; }
        public string Vari_Gyshare { get; set; }

    }

    #endregion

    #region 物件権利情報

    public class Bkdata_kenri_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Kenri_no { get; set; }
        public string Vari_Other_kenrirui { get; set; }
        public string Vari_Other_kenribiko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region ×物件KeyValueテーブル

    public class Bkdata_keyvalue_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Data_kbn { get; set; }
        public string Vari_Item_kbn { get; set; }
        public string Vari_Value { get; set; }

    }

    #endregion

    #region 物件交通情報

    public class Bkdata_kotu_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Ensen_cnt { get; set; }
        public string Vari_Ensen_no { get; set; }
        public string Vari_Eki_no { get; set; }
        public string Vari_Firsttrain_flg { get; set; }
        public string Vari_Kyori { get; set; }
        public string Vari_Toho_min { get; set; }
        public string Vari_Car_min { get; set; }
        public string Vari_Bus_min { get; set; }
        public string Vari_Bus_company { get; set; }
        public string Vari_Bus_station { get; set; }
        public string Vari_Bus_tohomin { get; set; }
        public string Vari_Bus_kyori { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Setting_kotutype { get; set; }
        public string Vari_Bus_companytoeki { get; set; }
        public string Vari_Bus_stationtoeki { get; set; }
        public string Vari_Bus_kyoritoeki { get; set; }
        public string Vari_Bus_tohomintoeki { get; set; }
        public string Vari_Kotu_syudan { get; set; }
        public string Vari_Car_kyori { get; set; }

    }

    #endregion

    #region 物件近隣駐車場情報(汎用ツールのみ)

    public class Bkdata_parkingother_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Parking_no { get; set; }
        public string Vari_Naiyo { get; set; }
        public string Vari_Kyori { get; set; }
        public string Vari_Gak { get; set; }
        public string Vari_Zeikbn { get; set; }

    }

    #endregion

    #region 物件参照ファイル情報(汎用ツールのみ)

    public class Bkdata_relfile_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_File_no { get; set; }
        public string Vari_Fullpath { get; set; }
        public string Vari_Addtime { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 物件接道情報

    public class Bkdata_setudo_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Setudo_cnt { get; set; }
        public string Vari_Setudo_muki { get; set; }
        public string Vari_Setudo_roadpattern { get; set; }
        public string Vari_Setudo_haba { get; set; }
        public string Vari_Setudo_setudokyori { get; set; }
        public string Vari_Setudo_pointselectflg { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 物件周辺情報

    public class Bkdata_syuhen_Model
    {

        public string Vari_Syuhen_guid { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Syuhen_cnt { get; set; }
        public string Vari_Sisetu_kbn { get; set; }
        public string Vari_Sisetu_name { get; set; }
        public string Vari_Sisetu_kyori { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region ×物件周辺画像情報

    public class Bkdata_syuhengazo_Model
    {

        public string Vari_Syuhen_guid { get; set; }
        public string Vari_Syuhen_gazono { get; set; }
        public string Vari_Syuhen_gazocomment { get; set; }
        public string Vari_Syuhen_gazolastupdate { get; set; }
        public string Vari_Syuhen_gazofilesizebyte { get; set; }
        public string Vari_Syuhen_gazowidth { get; set; }
        public string Vari_Syuhen_gazoheight { get; set; }
        public string Vari_Syuhen_gazothumbimage { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 物件修繕維持管理連絡先情報

    public class Bkdata_szeniji_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Syuzenijikanri_no { get; set; }
        public string Vari_Syuzenijikanri_kasyo { get; set; }
        public string Vari_Syuzenijikanri_taisyokbn { get; set; }
        public string Vari_Syuzenijikanri_gyno { get; set; }
        public string Vari_Syuzenijikanri_simei { get; set; }
        public string Vari_Syuzenijikanri_address { get; set; }
        public string Vari_Syuzenijikanri_tel { get; set; }
        public string Vari_Syuzenijikanri_jisyano { get; set; }
        public string Vari_Syuzenijikanri_kasino { get; set; }
        public string Vari_Syuzenijikanri_syono { get; set; }
        public string Vari_Syuzenijikanri_simeiu { get; set; }

    }

    #endregion

    #region ×物件bkdata_youtube情報

    public class Bkdata_youtube_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_No { get; set; }
        public string Vari_Uri { get; set; }
        public string Vari_Comment { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 物件メモ情報

    public class Bkdata_memo_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

}