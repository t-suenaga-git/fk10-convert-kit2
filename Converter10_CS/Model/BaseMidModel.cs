
namespace Converter10.Njc.Model
{

    #region 自社口座情報モデル

    public class Jisyadata_koza_BaseMid_Model
    {

        public string Vari_Jisya_no { get; set; }
        public string Vari_Jisya_kozano { get; set; }
        public string Vari_Kinyu_no { get; set; }
        public string Vari_Kinyu_tenno { get; set; }
        public string Vari_Koza_syubetu { get; set; }
        public string Vari_Koza_bango { get; set; }
        public string Vari_Koza_meigi { get; set; }
        public string Vari_Koza_meigikana { get; set; }
        public string Vari_Yucyokoza_kigo1 { get; set; }
        public string Vari_Yucyokoza_kigo2 { get; set; }
        public string Vari_Yucyokoza_bango { get; set; }
        public string Vari_Biko_koza { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Useflg { get; set; }

    }

    #endregion

    #region 部屋設備情報モデル

    public class Hydata_setubilst_BaseMid_Model
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

    #region 鍵タイトルマスタモデル

    public class M_kagi_title_BaseMid_Model
    {

        public string Vari_Kagi_kbn { get; set; }
        public string Vari_Kagi_no { get; set; }
        public string Vari_Kagi_name { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

}