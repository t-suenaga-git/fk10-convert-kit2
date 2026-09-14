
namespace Converter10.Njc.Model
{

    #region 物件分類モデル

    public class M_bk_rui_Model
    {

        public string Vari_Bk_ruino { get; set; }
        public string Vari_Bk_ruiname { get; set; }
        public string Vari_Bk_ruibiko { get; set; }
        public string Vari_Bk_ruiuseflg { get; set; }
        public string Vari_Sincyoku_keiyakukbn { get; set; }
        public string Vari_Sincyoku_kosinkbn { get; set; }
        public string Vari_Sincyoku_kaiyakukbn { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Homemate_ruikbn { get; set; }
        public string Vari_Bk_endofmonthflg { get; set; }
        public string Vari_Ikkatukariage_siwakekbn { get; set; }

    }

    #endregion

    #region 部屋分類モデル

    public class M_hy_rui_Model
    {

        public string Vari_Hy_ruino { get; set; }
        public string Vari_Hy_ruiname { get; set; }
        public string Vari_Hy_ruisyubetu { get; set; }
        public string Vari_Hy_ruiuseflg { get; set; }
        public string Vari_Hy_carportflg { get; set; }
        public string Vari_History { get; set; }
        // Public Property Vari_Hy_ruisortorder As String      'hy_ruisortorderは自動で挿入されるため対象外

    }

    #endregion

    #region 入金区分マスタモデル

    public class M_nkbn_Model
    {

        public string Vari_Nkbn_no { get; set; }          // 1-99のキー(1:現金 2:振込 3:振替 4:カード 5:代弁 99:移動)
        public string Vari_Nkbn_order { get; set; }       // 表示順
        public string Vari_Nkbn_name { get; set; }        // 入金区分名(1:現金 2:振込 3:振替 4:カード 5:代弁 99:移動)
        public string Vari_Nkbn_shortname { get; set; }   // 入金区分略称
        public string Vari_Nkbn_zokusei { get; set; }     // 1:現金, 2:振込, 3:振替, 4:その他, 99:移動
        public string Vari_History { get; set; }

    }

    #endregion

    #region 契約分類マスタモデル

    public class M_ky_rui_Model
    {

        public string Vari_Ky_ruino { get; set; }         // 契約分類No
        public string Vari_Ky_ruiname { get; set; }       // 契約分類名
        public string Vari_Ky_ruibiko { get; set; }       // 契約分類備考
        public string Vari_Ky_ruitutimm { get; set; }     // 契約終了通知月（◯ヶ月前）
        public string Vari_Ky_ruicolor { get; set; }      // 契約分類色
        public string Vari_Teisyaku_flg { get; set; }     // 1:使用する 2:使用しない
        public string Vari_History { get; set; }
        public string Vari_Useflg { get; set; }           // 1:検索対象 0:検索対象外

    }

    #endregion

    #region 入金項目マスタモデル

    public class M_nkin_Model
    {

        public string Vari_Nkin_no { get; set; }
        public string Vari_Nkin_name { get; set; }
        public string Vari_Nkin_printname { get; set; }
        public string Vari_Nkin_ruino { get; set; }
        public string Vari_Nkin_zkseino { get; set; }
        public string Vari_Nkin_useflg { get; set; }
        public string Vari_Keiyaku_nkinno { get; set; }
        public string Vari_Keiyakuhiki_nkinno { get; set; }
        public string Vari_Kosin_nkinno { get; set; }
        public string Vari_Kaiyaku_nkinno { get; set; }
        public string Vari_Kaiyakuhiki_nkinno { get; set; }
        public string Vari_Biko_nkin { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Ryosyu_flg { get; set; }

    }

    #endregion


    // 20160829 設備の新規挿入処理を追加 (現時点で未使用→後で削除) -del sta
    #region 設備グループマスタモデル

    public class M_setubi_rel_grp_Model
    {

        public string Vari_Setubi_grpguid { get; set; }
        public string Vari_Setubi_grpsortorder { get; set; }
        public string Vari_Setubi_grpname { get; set; }
        public string Vari_Setubi_grpuseflg { get; set; }
        public string Vari_Setubi_grpsyskbn { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 設備マスタモデル

    public class M_setubi_rel_Model
    {

        public string Vari_Setubi_guid { get; set; }
        public string Vari_Setubi_grpguid { get; set; }
        public string Vari_Setubi_sortorder { get; set; }
        public string Vari_Setubi_code { get; set; }
        public string Vari_Setubi_name { get; set; }
        public string Vari_Setubi_useflg { get; set; }
        public string Vari_Setubi_hyuseflg { get; set; }
        public string Vari_Setubi_disp1name { get; set; }
        public string Vari_Setubi_disp2name { get; set; }
        public string Vari_Setubi_disp3name { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Jyuyojiko_flg { get; set; }

    }

    #endregion

    #region 設備項目マスタモデル

    public class M_setubi_rel_komk_Model
    {

        public string Vari_Komok_guid { get; set; }
        public string Vari_Setubi_guid { get; set; }
        public string Vari_Komok_sortorder { get; set; }
        public string Vari_Komok_name { get; set; }
        public string Vari_Komok_code1 { get; set; }
        public string Vari_Komok_code2 { get; set; }
        public string Vari_Komok_code3 { get; set; }
        public string Vari_Default_flg { get; set; }
        public string Vari_Disp1name { get; set; }
        public string Vari_Disp2name { get; set; }
        public string Vari_Disp3name { get; set; }
        public string Vari_Disp1iconguid { get; set; }
        public string Vari_Disp2iconguid { get; set; }
        public string Vari_Disp3iconguid { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion
    // 20160829 設備の新規挿入処理を追加 (現時点で未使用→後で削除) -del end

}