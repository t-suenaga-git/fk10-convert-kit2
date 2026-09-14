
namespace Converter10.Njc.Model
{

    #region 部屋基本情報モデル

    public class Hydata_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_no { get; set; }
        public string Vari_Hy_gaibuno { get; set; }
        public string Vari_Hy_deleteflg { get; set; }
        public string Vari_Delete_guid { get; set; }
        public string Vari_Delete_day { get; set; }
        public string Vari_Delete_cnt { get; set; }
        public string Vari_Hy_ruinokbn { get; set; }
        public string Vari_Madori_cnt { get; set; }
        public string Vari_Madori_typekbn { get; set; }
        public string Vari_Madori_biko { get; set; }
        public string Vari_Men_senyujitu { get; set; }
        public string Vari_Men_senyujitutubo { get; set; }
        public string Vari_Men_yuka { get; set; }
        public string Vari_Men_yukatubo { get; set; }
        public string Vari_Men_senyutouki { get; set; }
        public string Vari_Men_senyutoukitubo { get; set; }
        public string Vari_Men_toki { get; set; }
        public string Vari_Men_tokitubo { get; set; }
        public string Vari_Men_balcony { get; set; }
        public string Vari_Men_balconytubo { get; set; }
        public string Vari_Men_tempo { get; set; }
        public string Vari_Men_tempotubo { get; set; }
        public string Vari_Men_jutaku { get; set; }
        public string Vari_Men_jutakutubo { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Delete_cause { get; set; }
        public string Vari_Kaiyaku_uketukekbn { get; set; }
        public string Vari_Kaiyaku_months { get; set; }
        public string Vari_Kaiyaku_days { get; set; }
        public string Vari_Kaiyaku_day { get; set; }
        public string Vari_Sort_hy_no { get; set; }       // 20160519 EXEUpdateに伴う修正 部屋基本情報 -add
        public string Vari_Lastupdate { get; set; }       // 20160519 EXEUpdateに伴う修正 部屋基本情報 -add
        public string Vari_Wmp_id { get; set; }           // 20161012 革命10アップデートに伴う修正 -add

    }

    #endregion

    #region 部屋詳細情報モデル

    public class Hydata_detail_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Hygyomu_setflg { get; set; }
        public string Vari_Syozai_kaisu1 { get; set; }
        public string Vari_Syozai_kaisu2 { get; set; }
        public string Vari_Syozai_kaisu3 { get; set; }
        public string Vari_Syozai_tikaflg1 { get; set; }
        public string Vari_Syozai_tikaflg2 { get; set; }
        public string Vari_Syozai_tikaflg3 { get; set; }
        public string Vari_Mukikbn { get; set; }
        public string Vari_Kadoheya { get; set; }
        public string Vari_Saikokbn { get; set; }
        public string Vari_Balcony_mukikbn { get; set; }
        public string Vari_Nyukyo_jokyokbn { get; set; }
        public string Vari_Nyukyo_jokyomemo { get; set; }
        public string Vari_Nyukyo_syunflg { get; set; }
        public string Vari_Nyukyo_ym { get; set; }
        public string Vari_Nyukyo_syunkbn { get; set; }
        public string Vari_Nyukyo_joken { get; set; }
        public string Vari_Kakunin_ymd { get; set; }
        public string Vari_Bosyu_startflg { get; set; }
        public string Vari_Freerent_flg { get; set; }
        public string Vari_Freerent_month { get; set; }
        public string Vari_Freerent_detaill { get; set; }
        public string Vari_Hoken_kbn { get; set; }
        public string Vari_Hoken_kikan { get; set; }
        public string Vari_Hoken_gak { get; set; }
        public string Vari_Hoken_biko { get; set; }
        public string Vari_Torihiki_taiyokbn { get; set; }
        public string Vari_Torihiki_kyakutuke { get; set; }
        public string Vari_Torihiki_tesumoto { get; set; }
        public string Vari_Torihiki_tesukyaku { get; set; }
        public string Vari_Torihiki_futankasi { get; set; }
        public string Vari_Torihiki_futankari { get; set; }
        public string Vari_Torihiki_kyakutukecomment { get; set; }
        public string Vari_Torihiki_gykokokukatudokbn { get; set; }
        public string Vari_Moto_gy_fudono { get; set; }
        public string Vari_Koukoku_ryoukbn { get; set; }
        public string Vari_Koukoku_jogengak { get; set; }
        public string Vari_Koukoku_jokennaiyo { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_Addr_replaceflg { get; set; }
        public string Vari_Addr_replacecyome { get; set; }
        public string Vari_Addr_replacecyomeptn { get; set; }
        public string Vari_Addr_replacebanti { get; set; }
        public string Vari_Addr_replaceetc { get; set; }
        public string Vari_Hosyo_gyno { get; set; }
        public string Vari_Hosyo_naiyo { get; set; }
        public string Vari_Kyrui_nokbn { get; set; }
        public string Vari_Keiyaku_kikan { get; set; }
        public string Vari_Keiyaku_kijitu { get; set; }
        public string Vari_Parking_biko { get; set; }
        public string Vari_Parking_bikebiko { get; set; }
        public string Vari_Parking_cyurinbiko { get; set; }
        public string Vari_Shared_salespoint { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Keisai_bkflg { get; set; }
        public string Vari_Keisai_hyflg { get; set; }
        public string Vari_Keisai_bantiflg { get; set; }
        public string Vari_Keisai_mapflg { get; set; }
        public string Vari_Hosyo_kbn { get; set; }
        public string Vari_Koukoku_jogengakkbn { get; set; }
        public string Vari_Koukoku_jogenrit { get; set; }
        public string Vari_Koukoku_jogentaxkbn { get; set; }
        public string Vari_Commonsalespoint_useflg { get; set; }
        public string Vari_Floors_flg { get; set; }
        public string Vari_Jisya_no { get; set; }
        public string Vari_Jisya_tanto { get; set; }
        public string Vari_Confirmky_sekininsya { get; set; }
        public string Vari_Confirmky_ymd { get; set; }
        public string Vari_Confirmky_print { get; set; }
        public string Vari_Confirmkai_sekininsya { get; set; }
        public string Vari_Confirmkai_ymd { get; set; }
        public string Vari_Confirmkai_print { get; set; }
        public string Vari_Nextnkin_kosindefault { get; set; }
        public string Vari_Toki_ymd { get; set; }
        public string Vari_Syo_kenriflg { get; set; }
        public string Vari_Syo_kenrikbn { get; set; }
        public string Vari_Other_kenriflg { get; set; }
        public string Vari_Btob_groupkbn { get; set; }
        public string Vari_Jisyaweb_osusumebk { get; set; }
        public string Vari_Kaiyaku_ym { get; set; }
        public string Vari_Taikyo_ym { get; set; }
        public string Vari_Svbunrui_no { get; set; }
        public string Vari_Krbunrui_no { get; set; }
        public string Vari_Siyo_mokuteki { get; set; }
        public string Vari_Nyukyo_sintikukbn { get; set; }
        public string Vari_Nyukyo_jikikbn { get; set; }
        public string Vari_Torihiki_jisyakbn { get; set; }
        public string Vari_Keiyaku_kikankbn { get; set; }

    }

    #endregion

    #region 部屋所有者情報モデル

    public class Hydata_syo_Model
    {

        public string Vari_Hy_guid { get; set; }
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

    #region 部屋駐車場情報モデル

    public class Hydata_parking_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Parking_kbn { get; set; }
        public string Vari_Parking_akisu { get; set; }
        public string Vari_Parking_status { get; set; }
        public string Vari_Parking_gakkbn { get; set; }
        public string Vari_Parking_gak { get; set; }
        public string Vari_Parking_gakzei { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Parking_tintaisu { get; set; }

    }

    #endregion

    #region 部屋特約情報モデル

    public class Hydata_tokuyaku_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Tokuyaku_grpno { get; set; }
        public string Vari_Naiyo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 部屋鍵情報モデル

    public class Hydata_kagi_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Kagi_no { get; set; }
        public string Vari_Honsu { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Hokan { get; set; }
        public string Vari_Gyshare { get; set; }

    }

    #endregion

    #region 部屋面積情報モデル

    public class Hydata_othermenseki_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Mensekitype { get; set; }
        public string Vari_No { get; set; }
        public string Vari_Mensekikbn { get; set; }
        public string Vari_Name { get; set; }
        public string Vari_Meter { get; set; }
        public string Vari_Tubo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 部屋間取内訳情報モデル

    public class Hydata_madoriutiwake_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_No { get; set; }
        public string Vari_Hykbn { get; set; }
        public string Vari_Jo { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Syozaikai { get; set; }

    }

    #endregion

    #region 部屋設備情報モデル

    public class Hydata_setubilst_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Komok_guid { get; set; }
        public string Vari_Override_kbn { get; set; }
        public string Vari_Komok_name { get; set; }
        public string Vari_Disp1name { get; set; }
        public string Vari_Disp2name { get; set; }
        public string Vari_Disp3name { get; set; }
        public string Vari_Disp1iconguid { get; set; }
        public string Vari_Disp2iconguid { get; set; }
        public string Vari_Disp3iconguid { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 部屋入金項目情報モデル

    public class Hydata_nkin_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Tuki_kbn { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkin_recno { get; set; }
        public string Vari_Nkin_kbn { get; set; }
        public string Vari_Sq_gak { get; set; }
        public string Vari_Sq_zeikbn { get; set; }
        public string Vari_Calc_kbn { get; set; }
        public string Vari_Calc_nkinno { get; set; }
        public string Vari_Calc_monthcnt { get; set; }
        public string Vari_Sqsaki_no { get; set; }
        public string Vari_Sq_mmkbn { get; set; }
        public string Vari_Sqstart_ymd { get; set; }
        public string Vari_Sq_ptn { get; set; }
        public string Vari_Sq_interval { get; set; }
        public string Vari_Sq_nen { get; set; }
        public string Vari_Sq_tuki { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 部屋変動費各戸メーター情報モデル

    public class Hydata_hendo_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Hyhendo_guid { get; set; }
        public string Vari_Rec_no { get; set; }
        public string Vari_Hendo_kbn { get; set; }
        public string Vari_Meter_name { get; set; }
        public string Vari_Nkin_no { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_Hendorule_no { get; set; }
        public string Vari_Hendorule_biko { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Sq_mmkbn { get; set; }
        public string Vari_Useflg { get; set; }

    }

    #endregion

    #region 部屋修繕維持管理連絡先情報モデル

    public class Hydata_szeniji_Model
    {

        public string Vari_Hy_guid { get; set; }
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

    #region 部屋メモ情報モデル

    public class Hydata_memo_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    // 汎用ツールのみ

    #region 部屋共通セールスポイント情報モデル

    public class Hydata_commonsalespointparts_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Parts_no { get; set; }
        public string Vari_Salespoint_parts { get; set; }

    }

    #endregion

    #region 部屋契約解約確認事項情報モデル

    public class Hydata_confirm_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Confirm_kbn { get; set; }
        public string Vari_Confirm_no { get; set; }
        public string Vari_Naiyo { get; set; }

    }

    #endregion

    #region 部屋権利情報モデル

    public class Hydata_kenri_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Kenri_no { get; set; }
        public string Vari_Other_kenrirui { get; set; }
        public string Vari_Other_kenribiko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 部屋参照ファイル情報モデル

    public class Hydata_relfile_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_File_no { get; set; }
        public string Vari_Fullpath { get; set; }
        public string Vari_Addtime { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 部屋原状回復目安単価情報モデル

    public class Hydata_szen_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Szen_no { get; set; }
        public string Vari_Szen_name { get; set; }
        public string Vari_Ryo { get; set; }
        public string Vari_Unit_name { get; set; }
        public string Vari_Tanka { get; set; }

    }

    #endregion

    // プログラム内で自動実行

    #region 部屋外部No管理情報モデル

    public class Hydata_gaibuno_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Current_no { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    // 他の箇所で移行

    #region 部屋画像情報モデル

    // Public Class Hydata_gazo_Model

    // Public Property Vari_Hy_guid As String
    // Public Property Vari_Hy_gazono As String
    // Public Property Vari_Hy_gazocomment As String
    // Public Property Vari_Hy_gazolastupdate As String
    // Public Property Vari_Hy_gazofilesizekbyte As String
    // Public Property Vari_Hy_gazowidth As String
    // Public Property Vari_Hy_gazoheight As String
    // Public Property Vari_Hy_gazothumbimage As String
    // Public Property Vari_History As String

    // End Class

    #endregion

    // 移行対象外

    #region 部屋業務期間情報モデル

    public class Hydata_gyomukikan_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Kn_no { get; set; }
        public string Vari_Startymd { get; set; }
        public string Vari_Endymd { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    // 不明テーブル
    // 20160720 連動情報構築 -del sta
    // #Region "部屋BtoBグループ情報モデル"

    // Public Class Hydata_btobgroup_Model

    // Public Property Vari_Hy_guid As String
    // Public Property Vari_Group_no As String
    // Public Property Vari_Display_kbn As String
    // Public Property Vari_History As String

    // End Class

    // #End Region
    // 20160720 連動情報構築 -del end
    #region 部屋KeyValue情報モデル

    public class Hydata_keyvalue_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Data_kbn { get; set; }
        public string Vari_Item_kbn { get; set; }
        public string Vari_Value { get; set; }

    }

    #endregion
    // 20160720 連動情報構築 -del sta
    // #Region "部屋広告補足情報(作り直し中)モデル"

    // Public Class Hydata_kokoku_Model

    // Public Property Vari_Hy_guid As String
    // Public Property Vari_Site_no As String
    // Public Property Vari_Item As String

    // End Class

    // #End Region
    // 20160720 連動情報構築 -del end
    #region 部屋hydata_kys_atenaprintモデル

    public class Hydata_kys_atenaprint_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Hy_guid { get; set; }

    }

    #endregion

    #region 部屋ポータル最終送信履歴情報モデル

    public class Hydata_lastsend_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Site_no { get; set; }
        public string Vari_Setting_guid { get; set; }
        public string Vari_Lastsendtime { get; set; }

    }

    #endregion

    #region 部屋セールスポイント情報モデル

    public class Hydata_salespoint_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Site_no { get; set; }
        public string Vari_Salespoint { get; set; }

    }

    #endregion
    // 20160720 連動情報構築 -del sta
    // #Region "部屋連動サイト毎地図表示情報モデル"

    // Public Class Hydata_mapdisp_Model

    // Public Property Vari_Hy_guid As String
    // Public Property Vari_Site_no As String
    // Public Property Vari_Site_jisyano As String
    // Public Property Vari_Map_dispflg As String
    // Public Property Vari_History As String

    // End Class

    // #End Region
    // 20160720 連動情報構築 -del end
    #region 部屋hydata_panorama情報モデル

    public class Hydata_panorama_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_No { get; set; }
        public string Vari_Gazokbn { get; set; }
        public string Vari_Title { get; set; }
        public string Vari_Lenskbn { get; set; }
        public string Vari_Uri { get; set; }
        public string Vari_Gazolastupdate { get; set; }
        public string Vari_Gazofilesizekbyte { get; set; }
        public string Vari_Gazowidth { get; set; }
        public string Vari_Gazoheight { get; set; }
        public string Vari_Gazothumbimage { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 部屋支店情報モデル

    public class Hydata_siten_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Rec_no { get; set; }
        public string Vari_Siten_no { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion
    // 20160720 連動情報構築 -del sta
    // #Region "部屋hydata_sosin情報モデル"

    // Public Class Hydata_sosin_Model

    // Public Property Vari_Hy_guid As String
    // Public Property Vari_Site_no As String
    // Public Property Vari_Site_jisyano As String
    // Public Property Vari_Select_no As String
    // Public Property Vari_Komok_guid As String
    // Public Property Vari_History As String

    // End Class

    // #End Region
    // 20160720 連動情報構築 -del end
    #region 部屋hydata_youtube情報モデル

    public class Hydata_youtube_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_No { get; set; }
        public string Vari_Uri { get; set; }
        public string Vari_Comment { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 部屋親メーター関連情報モデル

    public class Hydata_hendoparent_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Bk_guid { get; set; }
        public string Vari_Bkhendo_guid { get; set; }
        public string Vari_Kanren_flg { get; set; }
        public string Vari_Childmeter_name { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

}