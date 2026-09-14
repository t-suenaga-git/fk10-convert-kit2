
namespace Converter10.Njc.Model
{

    #region 修繕基本情報モデル

    public class Szendata_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Szen_no { get; set; }
        public string Vari_Hy_useflg { get; set; }
        public string Vari_Hy_guid { get; set; }
        public string Vari_Ky_useflg { get; set; }
        public string Vari_Ky_guid { get; set; }
        public string Vari_Szen_name { get; set; }
        public string Vari_Szen_ukeymd { get; set; }
        public string Vari_Szen_endymd { get; set; }
        public string Vari_Tanto_no { get; set; }
        public string Vari_Jisya_no { get; set; }
        public string Vari_Koji_yoteistartymd { get; set; }
        public string Vari_Koji_yoteiendymd { get; set; }
        public string Vari_Koji_basyo { get; set; }
        public string Vari_Koji_gaiyo { get; set; }
        public string Vari_Koji_kagi { get; set; }
        public string Vari_Koji_tatiai { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_Kys_futanrit { get; set; }
        public string Vari_Ow_futanrit { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Jisya_futanrit { get; set; }
        public string Vari_Kys_sqflg { get; set; }
        public string Vari_Kys_no { get; set; }
        public string Vari_Kys_gtym { get; set; }
        public string Vari_Kys_sqsimeymd { get; set; }
        public string Vari_Kys_nkbn { get; set; }
        public string Vari_Kys_sorit { get; set; }
        public string Vari_Kys_yatinkozano { get; set; }
        public string Vari_Kys_biko { get; set; }
        public string Vari_Ow_sqflg { get; set; }
        public string Vari_Ow_no { get; set; }
        public string Vari_Ow_kaisyukbn { get; set; }
        public string Vari_Ow_kojoymd { get; set; }
        public string Vari_Ow_kojogtym { get; set; }
        public string Vari_Ow_sosaki_soruleguid { get; set; }
        public string Vari_Ow_sosaki_sorule_no { get; set; }
        public string Vari_Ow_sokozano { get; set; }
        public string Vari_Ow_sqymd { get; set; }
        public string Vari_Ow_sqgtym { get; set; }
        public string Vari_Ow_sqsqsimeymd { get; set; }
        public string Vari_Ow_sqnkbn { get; set; }
        public string Vari_Ow_sqyatinkozano { get; set; }
        public string Vari_Ow_sqsakino { get; set; }
        public string Vari_Ow_biko { get; set; }

    }

    #endregion

    #region 修繕見積情報モデル

    public class Szendata_szen_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Szen_no { get; set; }
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

    #region 修繕見積詳細情報モデル

    public class Szendata_szenmeisai_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Szen_no { get; set; }
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

    #region 修繕クレーム関連付け情報モデル

    public class Szendata_claim_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Szen_no { get; set; }
        public string Vari_Claim_no { get; set; }
        public string Vari_Claimdata_sortorder { get; set; }
        public string Vari_Szendata_sortorder { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 修繕関連ファイル情報モデル

    public class Szendata_relfile_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Szen_no { get; set; }
        public string Vari_File_no { get; set; }
        public string Vari_Fullpath { get; set; }
        public string Vari_Addtime { get; set; }
        public string Vari_Biko { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 修繕メモ情報モデル

    public class Szendata_memo_Model
    {

        public string Vari_Bk_guid { get; set; }
        public string Vari_Szen_no { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

}