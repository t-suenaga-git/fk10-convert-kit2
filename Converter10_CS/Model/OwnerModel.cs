
namespace Converter10.Njc.Model
{

    #region 家主情報モデル

    public class Owdata_Model
    {

        public string Vari_Ow_no { get; set; }
        public string Vari_Kojinhojin_flg { get; set; }
        public string Vari_Ow_name { get; set; }
        public string Vari_Ow_namesjis { get; set; }
        public string Vari_Ow_kana { get; set; }
        public string Vari_Keisyo { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr_kenno { get; set; }
        public string Vari_Addr_sino { get; set; }
        public string Vari_Addr_cyo { get; set; }
        public string Vari_Addr_cyome { get; set; }
        public string Vari_Addr_cyomeptn { get; set; }
        public string Vari_Addr_banti { get; set; }
        public string Vari_Addr_etc { get; set; }
        public string Vari_Toukiaddr1 { get; set; }
        public string Vari_Toukiaddr2 { get; set; }
        public string Vari_Tel1 { get; set; }
        public string Vari_Tel2 { get; set; }
        public string Vari_Fax { get; set; }
        public string Vari_Mobiletel1 { get; set; }
        public string Vari_Mobiletel2 { get; set; }
        public string Vari_Mail { get; set; }
        public string Vari_Mobilemail { get; set; }
        public string Vari_Yusentel_kbn { get; set; }
        public string Vari_Yusenmail_kbn { get; set; }
        public string Vari_Birthday { get; set; }
        public string Vari_Nensyu { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Gender { get; set; }
        public string Vari_Honseki { get; set; }
        public string Vari_Kinmu_name { get; set; }
        public string Vari_Kinmu_namesjis { get; set; }
        public string Vari_Kinmu_kana { get; set; }
        public string Vari_Kinmu_postcode { get; set; }
        public string Vari_Kinmu_addr1 { get; set; }
        public string Vari_Kinmu_addr2 { get; set; }
        public string Vari_Kinmu_tel1 { get; set; }
        public string Vari_Kinmu_tel2 { get; set; }
        public string Vari_Kinmu_fax { get; set; }
        public string Vari_Kinmu_gyosyu { get; set; }
        public string Vari_Kinmu_busyo { get; set; }
        public string Vari_Kinmu_nyuryokuym { get; set; }
        public string Vari_Kinmu_nyusyaym { get; set; }
        public string Vari_Url { get; set; }
        public string Vari_Gyosyu { get; set; }
        public string Vari_Daihyo_name { get; set; }
        public string Vari_Daihyo_namesjis { get; set; }
        public string Vari_Daihyo_kana { get; set; }
        public string Vari_Daihyo_yakusyoku { get; set; }
        public string Vari_Daihyo_atenaflg { get; set; }
        public string Vari_Tanto_name { get; set; }
        public string Vari_Tanto_namesjis { get; set; }
        public string Vari_Tanto_kana { get; set; }
        public string Vari_Tanto_busyo { get; set; }
        public string Vari_Tanto_atenaflg { get; set; }
        public string Vari_Sihonkin { get; set; }
        public string Vari_Jugyosu { get; set; }
        public string Vari_Nyuryoku_ym { get; set; }
        public string Vari_Torihikisaki { get; set; }
        public string Vari_Renraku_name { get; set; }
        public string Vari_Renraku_namesjis { get; set; }
        public string Vari_Renraku_kana { get; set; }
        public string Vari_Renraku_keisyo { get; set; }
        public string Vari_Renraku_postcode { get; set; }
        public string Vari_Renraku_addr1 { get; set; }
        public string Vari_Renraku_addr2 { get; set; }
        public string Vari_Renraku_tel1 { get; set; }
        public string Vari_Renraku_tel2 { get; set; }
        public string Vari_Renraku_fax { get; set; }
        public string Vari_Renraku_mobiletel1 { get; set; }
        public string Vari_Renraku_mobiletel2 { get; set; }
        public string Vari_Renraku_yusentelkbn { get; set; }
        public string Vari_Renraku_aidagara { get; set; }
        public string Vari_Biko_renraku { get; set; }
        public string Vari_Sofu_kbn { get; set; }
        public string Vari_Sofu_name { get; set; }
        public string Vari_Sofu_namesjis { get; set; }
        public string Vari_Sofu_kana { get; set; }
        public string Vari_Sofu_keisyo { get; set; }
        public string Vari_Sofu_postcode { get; set; }
        public string Vari_Sofu_addr1 { get; set; }
        public string Vari_Sofu_addr2 { get; set; }
        public string Vari_Sofu_tel1 { get; set; }
        public string Vari_Sofu_tel2 { get; set; }
        public string Vari_Sofu_fax { get; set; }
        public string Vari_Biko_sofu { get; set; }
        public string Vari_Event_tantono { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_Findkeyword { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Kaikei_jouhou { get; set; }
        public string Vari_Addr1 { get; set; }            // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add
        public string Vari_Addr2 { get; set; }            // 20160622 EXEUpdateに伴う修正 (家主住所移行仕様変更) -add

    }

    #endregion

    #region 家主口座情報モデル

    public class Owdata_koza_Model
    {

        public string Vari_Ow_no { get; set; }
        public string Vari_Ow_kozano { get; set; }
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
        public string Vari_Biko_furikomi { get; set; }
        public string Vari_Sgfirai_kbn { get; set; }
        public string Vari_Sgfirai_no { get; set; }
        public string Vari_Sgfirai_tesufutankbn { get; set; }
        public string Vari_Sgfirai_tesukeisankbn { get; set; }
        public string Vari_Sgfirai_tesukotei1gak { get; set; }
        public string Vari_Sgfirai_tesukotei2gak { get; set; }
        public string Vari_Biko_sgfirai { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Koza_printkbn { get; set; }

    }

    #endregion

    #region 家主イベント情報モデル

    public class Owdata_event_Model
    {

        public string Vari_Ow_no { get; set; }
        public string Vari_Event_kbn { get; set; }
        public string Vari_Event_cnt { get; set; }
        public string Vari_Event_ymd { get; set; }
        public string Vari_Event_data { get; set; }

    }

    #endregion

    #region 家主メモ情報モデル

    public class Owdata_memo_Model
    {

        public string Vari_Ow_no { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

}