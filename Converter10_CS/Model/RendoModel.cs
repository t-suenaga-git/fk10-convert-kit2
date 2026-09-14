
namespace Converter10.Njc.Model
{

    #region 送信設定基本情報モデル

    public class M_sendsetting_Model
    {

        public string Vari_Setting_guid { get; set; }
        public string Vari_Setting_sortorder { get; set; }
        public string Vari_Setting_name { get; set; }
        public string Vari_Keisai_siten { get; set; }
        public string Vari_Datalink_id { get; set; }
        public string Vari_Adconfirm_limitday { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Basedate { get; set; }
        public string Vari_Basedate_flg { get; set; }
        public string Vari_Emailaddr1 { get; set; }
        public string Vari_Emailaddr2 { get; set; }
        public string Vari_Emailaddr3 { get; set; }
        public string Vari_Emailaddr4 { get; set; }
        public string Vari_Emailaddr5 { get; set; }

    }

    #endregion

    #region 送信設定情報モデル

    public class M_site_sendsetting_Model
    {

        public string Vari_Sitesetting_guid { get; set; }
        public string Vari_Setting_guid { get; set; }
        public string Vari_Site_no { get; set; }
        public string Vari_Site_id { get; set; }
        public string Vari_Site_sendumu { get; set; }
        public string Vari_Site_data { get; set; }
        public string Vari_Site_password { get; set; }

    }

    #endregion

    #region 広告補足情報モデル

    public class Hydata_kokoku_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Site_no { get; set; }
        public string Vari_Item { get; set; }

    }

    #endregion

    #region ポータル連動部屋分類情報モデル

    public class M_hy_ruisite_Model
    {

        public string Vari_Hy_ruino { get; set; }
        public string Vari_Site_no { get; set; }
        public string Vari_Hy_ruisiteitemdata { get; set; }

    }

    #endregion

    #region 部屋毎送信情報モデル

    public class Hydata_sosin_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Site_no { get; set; }
        public string Vari_Site_jisyano { get; set; }
        public string Vari_Select_no { get; set; }
        public string Vari_Komok_guid { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region BtoBグループ設定情報モデル

    public class Hydata_btobgroup_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Group_no { get; set; }
        public string Vari_Display_kbn { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 地図表示詳細設定情報モデル

    public class Hydata_mapdisp_Model
    {

        public string Vari_Hy_guid { get; set; }
        public string Vari_Site_no { get; set; }
        public string Vari_Site_jisyano { get; set; }
        public string Vari_Map_dispflg { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

}