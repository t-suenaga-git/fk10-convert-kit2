
namespace Converter10.Njc.Model
{

    #region 物件分類モデル

    // 20160525 紐付項目の挿入処理の追加 -del sta
    // Public Class M_bk_rui_Model

    // Public Property Vari_Bk_ruino As String
    // Public Property Vari_Bk_ruiname As String
    // Public Property Vari_Bk_ruibiko As String
    // Public Property Vari_Bk_ruiuseflg As String
    // Public Property Vari_Sincyoku_keiyakukbn As String
    // Public Property Vari_Sincyoku_kosinkbn As String
    // Public Property Vari_Sincyoku_kaiyakukbn As String
    // Public Property Vari_History As String
    // Public Property Vari_Homemate_ruikbn As String
    // Public Property Vari_Bk_endofmonthflg As String
    // Public Property Vari_Ikkatukariage_siwakekbn As String

    // End Class
    // 20160525 紐付項目の挿入処理の追加 -del end

    #endregion

    #region バスモデル

    public class M_buskotu_Model
    {

        public string Vari_Buskotu_no { get; set; }
        public string Vari_Buskotu_name { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region バス交通モデル

    public class M_buskotu_stop_Model
    {

        public string Vari_Buskotu_no { get; set; }
        public string Vari_Sortorder { get; set; }
        public string Vari_Keito_name { get; set; }
        public string Vari_Busstop_name { get; set; }

    }

    #endregion

    #region 鍵タイトルモデル

    public class M_kagi_title_Model
    {

        public string Vari_Kagi_kbn { get; set; }     // 鍵区分
        public string Vari_Kagi_no { get; set; }      // 鍵タイトルNo
        public string Vari_Kagi_name { get; set; }    // 鍵タイトル名
        public string Vari_History { get; set; }

    }

    #endregion

    #region 箇所分類マスタモデル

    public class M_claim_rui_Model
    {

        public string Vari_Claim_ruikbn { get; set; }         // 1:箇所区分 2:クレーム分類
        public string Vari_Claim_ruino { get; set; }
        public string Vari_Claim_ruisortorder { get; set; }
        public string Vari_Claim_name { get; set; }
        public string Vari_Claim_useflg { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region クレーム分類マスタモデル

    // Public Class M_claim_rui_claim_Model

    // Public Property Vari_Claim_ruikbn As String         '1:箇所区分 2:クレーム分類
    // Public Property Vari_Claim_ruino As String
    // Public Property Vari_Claim_ruisortorder As String
    // Public Property Vari_Claim_name As String
    // Public Property Vari_Claim_useflg As String
    // Public Property Vari_History As String

    // End Class

    #endregion

    #region 特約事項マスタモデル

    public class M_tokuyaku_Model
    {

        public string Vari_Tokuyaku_grpno { get; set; }       // 特約グループNo
        public string Vari_Tokuyaku_no { get; set; }          // 特約No
        public string Vari_Tokuyaku_title { get; set; }       // 特約タイトル
        public string Vari_Tokuyaku_template { get; set; }    // 特約テンプレート
        public string Vari_History { get; set; }

    }

    #endregion

    #region 原状回復特約マスタモデル

    // Public Class M_tokuyaku_genjo_Model

    // Public Property Vari_Tokuyaku_grpno As String       '特約グループNo
    // Public Property Vari_Tokuyaku_no As String          '特約No
    // Public Property Vari_Tokuyaku_title As String       '特約タイトル
    // Public Property Vari_Tokuyaku_template As String    '特約テンプレート
    // Public Property Vari_History As String

    // End Class

    #endregion

    #region 契約分類マスタモデル

    // 20160525 紐付項目の挿入処理の追加 -del sta
    // Public Class M_ky_rui_Model

    // Public Property Vari_Ky_ruino As String         '契約分類No
    // Public Property Vari_Ky_ruiname As String       '契約分類名
    // Public Property Vari_Ky_ruibiko As String       '契約分類備考
    // Public Property Vari_Ky_ruitutimm As String     '契約終了通知月（◯ヶ月前）
    // Public Property Vari_Ky_ruicolor As String      '契約分類色
    // Public Property Vari_Teisyaku_flg As String     '1:使用する 2:使用しない
    // Public Property Vari_History As String
    // Public Property Vari_Useflg As String           '1:検索対象 0:検索対象外

    // End Class
    // 20160525 紐付項目の挿入処理の追加 -del end

    #endregion

    #region 保険種類マスタモデル

    public class M_hoken_rui_Model
    {

        public string Vari_Hoken_ruino { get; set; }
        public string Vari_Hoken_ruiname { get; set; }
        public string Vari_Hoken_ruikana { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Useflg { get; set; }

    }

    #endregion

    #region 入金区分マスタモデル

    // 20160525 紐付項目の挿入処理の追加 -del sta
    // Public Class M_nkbn_Model

    // Public Property Vari_Nkbn_no As String          '1-99のキー(1:現金 2:振込 3:振替 4:カード 5:代弁 99:移動)
    // Public Property Vari_Nkbn_order As String       '表示順
    // Public Property Vari_Nkbn_name As String        '入金区分名(1:現金 2:振込 3:振替 4:カード 5:代弁 99:移動)
    // Public Property Vari_Nkbn_shortname As String   '入金区分略称
    // Public Property Vari_Nkbn_zokusei As String     '1:現金, 2:振込, 3:振替, 4:その他, 99:移動
    // Public Property Vari_History As String

    // End Class
    // 20160525 紐付項目の挿入処理の追加 -del end
    #endregion

    #region 学校区マスタモデル

    public class M_koku_add_Model
    {

        public string Vari_Ken_no { get; set; }           // 県No
        public string Vari_Si_no { get; set; }            // 市No
        public string Vari_No { get; set; }               // 学校区No
        public string Vari_Add_cyo { get; set; }          // 町地域
        public string Vari_Add_banti { get; set; }        // 番地
        public string Vari_Syogaku_name { get; set; }     // 小学校名
        public string Vari_Cyugaku_name { get; set; }     // 中学校名
        public string Vari_History { get; set; }
        public string Vari_Add_cyome { get; set; }        // NULL

    }

    #endregion

    #region エリアマスタモデル

    public class M_area_Model
    {

        public string Vari_Area_no { get; set; }          // エリアNo
        public string Vari_Area_name { get; set; }        // エリア名
        public string Vari_Area_sortorder { get; set; }   // 表示順
        public string Vari_Area_useflg { get; set; }      // 使用有無
        public string Vari_History { get; set; }

    }

    #endregion

    #region 入金項目マスタモデル

    // 20160525 紐付項目の挿入処理の追加 -del sta
    // Public Class M_nkin_Model

    // Public Property Vari_Nkin_no As String
    // Public Property Vari_Nkin_name As String
    // Public Property Vari_Nkin_printname As String
    // Public Property Vari_Nkin_ruino As String
    // Public Property Vari_Nkin_zkseino As String
    // Public Property Vari_Nkin_useflg As String          '1:使用する, 2:使用しない
    // Public Property Vari_Keiyaku_nkinno As String
    // Public Property Vari_Keiyakuhiki_nkinno As String
    // Public Property Vari_Kosin_nkinno As String
    // Public Property Vari_Kaiyaku_nkinno As String
    // Public Property Vari_Kaiyakuhiki_nkinno As String
    // Public Property Vari_Biko_nkin As String
    // Public Property Vari_History As String
    // Public Property Vari_Ryosyu_flg As String           '1:有り, 2:なし, -1:未定義

    // End Class
    // 20160525 紐付項目の挿入処理の追加 -del end

    #endregion

    #region 口座種別マスタモデル

    public class M_kozasyu_Model
    {

        public string Vari_Kosyu_no { get; set; }
        public string Vari_Kosyu_name { get; set; }
        public string Vari_Kosyu_sname { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 変動費マスタモデル

    public class M_hendorule_Model
    {

        public string Vari_Rule_no { get; set; }
        public string Vari_Rule_name { get; set; }
        public string Vari_Tani { get; set; }
        public string Vari_Tanisjis { get; set; }
        public string Vari_Ryokin_cnt { get; set; }
        public string Vari_Hendo_rui { get; set; }
        public string Vari_Hasuusyori_sel { get; set; }
        public string Vari_Hasuusyosu_ptn { get; set; }
        public string Vari_Koukei { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Keijo_rui { get; set; }
        public string Vari_Zei_kbnhayami { get; set; }
        public string Vari_Zei_kbntanka { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Useflg { get; set; }

    }

    #endregion

    #region 変動費一覧マスタモデル

    public class M_hendorule_itiran_Model
    {

        public string Vari_Rule_no { get; set; }
        public string Vari_Item_no { get; set; }
        public string Vari_Siyoryo { get; set; }
        public string Vari_Ryokin1 { get; set; }
        public string Vari_Ryokin2 { get; set; }
        public string Vari_Ryokin3 { get; set; }
        public string Vari_Ryokin4 { get; set; }
        public string Vari_Ryokin5 { get; set; }

    }

    #endregion

    #region 備考タイトルマスタ

    public class M_memo_Model
    {

        public string Vari_Memo_kbn { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo_sortorder { get; set; }
        public string Vari_Memo_name { get; set; }
        public string Vari_Memo_useflg { get; set; }
        public string Vari_Memo_usehojoitems { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 備考入力補助リストマスタ

    public class M_memo_lst_Model
    {

        public string Vari_Memo_kbn { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo_lstno { get; set; }
        public string Vari_Memo_lstname { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 画像タイトルマスタ

    public class M_gazo_title_Model
    {

        public string Vari_Gazo_kbn { get; set; }
        public string Vari_Gazo_no { get; set; }
        public string Vari_Gazo_name { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

}