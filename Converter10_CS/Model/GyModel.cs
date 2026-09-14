
namespace Converter10.Njc.Model
{

    #region 仲介・管理業者マスタモデル

    public class Gydata_fudo_Model
    {

        public string Vari_Gy_fudono { get; set; }
        public string Vari_Gy_fudoname { get; set; }
        public string Vari_Gy_fudonamesjis { get; set; }
        public string Vari_Gy_fudokana { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr_kenno { get; set; }
        public string Vari_Addr_sino { get; set; }
        public string Vari_Addr_cyo { get; set; }
        public string Vari_Addr_cyome { get; set; }
        public string Vari_Addr_cyomeptn { get; set; }
        public string Vari_Addr_banti { get; set; }
        public string Vari_Addr_etc { get; set; }
        public string Vari_Tel1 { get; set; }
        public string Vari_Tel2 { get; set; }
        public string Vari_Fax { get; set; }
        public string Vari_Mobiletel1 { get; set; }
        public string Vari_Mobiletel2 { get; set; }
        public string Vari_Daihyo_yakusyoku { get; set; }
        public string Vari_Daihyo_name { get; set; }
        public string Vari_Tanto_busyoyakusyoku { get; set; }
        public string Vari_Tanto_name { get; set; }
        public string Vari_Tanto_namesjis { get; set; }
        public string Vari_Tanto_kana { get; set; }
        public string Vari_Mail { get; set; }
        public string Vari_Mobilemail { get; set; }
        public string Vari_Url { get; set; }
        public string Vari_Yusentel_kbn { get; set; }
        public string Vari_Yusenmail_kbn { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Sofu_hakkofurikomi { get; set; }
        public string Vari_Menkyo_bango { get; set; }
        public string Vari_Menkyo_ymd { get; set; }
        public string Vari_Syunin_bango { get; set; }
        public string Vari_Syunin_name { get; set; }
        public string Vari_Kanrikyokai_bango { get; set; }
        public string Vari_Keieikanrisi_bango { get; set; }
        public string Vari_Keieikanrisi_name { get; set; }
        public string Vari_Kanrigy_bango { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Daihyo_namesjis { get; set; }
        public string Vari_Syunin_namesjis { get; set; }
        public string Vari_Tanto_autoinputflg { get; set; }
        public string Vari_Keieikanrisi_namesjis { get; set; }

    }

    #endregion

    #region 仲介・管理業者マスタ口座モデル

    public class Gydata_fudokoza_Model
    {

        public string Vari_Gy_fudono { get; set; }
        public string Vari_Gy_fudokozano { get; set; }
        public string Vari_Kinyu_no { get; set; }
        public string Vari_Kinyu_tenno { get; set; }
        public string Vari_Koza_syubetu { get; set; }
        public string Vari_Koza_bango { get; set; }
        public string Vari_Koza_meigi { get; set; }
        public string Vari_Koza_meigikana { get; set; }
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
        public string Vari_Yucyokoza_kigo1 { get; set; }
        public string Vari_Yucyokoza_kigo2 { get; set; }
        public string Vari_Yucyokoza_bango { get; set; }

    }

    #endregion

    #region 仲介・管理業者マスタメモモデル

    public class Gydata_fudomemo_Model
    {

        public string Vari_Gy_fudono { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 保険業者マスタモデル

    public class Gydata_hoken_Model
    {

        public string Vari_Gy_hokenno { get; set; }
        public string Vari_Gy_hokenname { get; set; }
        public string Vari_Gy_hokennamesjis { get; set; }
        public string Vari_Gy_hokenkana { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr1 { get; set; }
        public string Vari_Addr2 { get; set; }
        public string Vari_Tel1 { get; set; }
        public string Vari_Tel2 { get; set; }
        public string Vari_Fax { get; set; }
        public string Vari_Mobiletel1 { get; set; }
        public string Vari_Mobiletel2 { get; set; }
        public string Vari_Tanto_busyoyakusyoku { get; set; }
        public string Vari_Tanto_name { get; set; }
        public string Vari_Tanto_namesjis { get; set; }
        public string Vari_Tanto_kana { get; set; }
        public string Vari_Mail { get; set; }
        public string Vari_Mobilemail { get; set; }
        public string Vari_Url { get; set; }
        public string Vari_Yusentel_kbn { get; set; }
        public string Vari_Yusenmail_kbn { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }

    }

    #endregion

    #region 保険業者マスタ口座モデル

    public class Gydata_hokenkoza_Model
    {

        public string Vari_Gy_hokenno { get; set; }
        public string Vari_Gy_hokenkozano { get; set; }
        public string Vari_Kinyu_no { get; set; }
        public string Vari_Kinyu_tenno { get; set; }
        public string Vari_Koza_syubetu { get; set; }
        public string Vari_Koza_bango { get; set; }
        public string Vari_Koza_meigi { get; set; }
        public string Vari_Koza_meigikana { get; set; }
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

    }

    #endregion

    #region 保険業者マスタメモモデル

    public class Gydata_hokenmemo_Model
    {

        public string Vari_Gy_hokenno { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 家賃保証業者マスタモデル

    public class Gydata_hosyo_Model
    {

        public string Vari_Gy_hosyono { get; set; }
        public string Vari_Gy_hosyoname { get; set; }
        public string Vari_Gy_hosyonamesjis { get; set; }
        public string Vari_Gy_hosyokana { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr1 { get; set; }
        public string Vari_Addr2 { get; set; }
        public string Vari_Tel1 { get; set; }
        public string Vari_Tel2 { get; set; }
        public string Vari_Fax { get; set; }
        public string Vari_Mobiletel1 { get; set; }
        public string Vari_Mobiletel2 { get; set; }
        public string Vari_Tanto_busyoyakusyoku { get; set; }
        public string Vari_Tanto_name { get; set; }
        public string Vari_Tanto_namesjis { get; set; }
        public string Vari_Tanto_kana { get; set; }
        public string Vari_Mail { get; set; }
        public string Vari_Mobilemail { get; set; }
        public string Vari_Url { get; set; }
        public string Vari_Yusentel_kbn { get; set; }
        public string Vari_Yusenmail_kbn { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Rendo_namekbn { get; set; }

    }

    #endregion

    #region 家賃保証業者マスタ口座モデル

    public class Gydata_hosyokoza_Model
    {

        public string Vari_Gy_hosyono { get; set; }
        public string Vari_Gy_hosyokozano { get; set; }
        public string Vari_Kinyu_no { get; set; }
        public string Vari_Kinyu_tenno { get; set; }
        public string Vari_Koza_syubetu { get; set; }
        public string Vari_Koza_bango { get; set; }
        public string Vari_Koza_meigi { get; set; }
        public string Vari_Koza_meigikana { get; set; }
        public string Vari_Fkae_no { get; set; }
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

    }

    #endregion

    #region 家賃保証業者マスタメモモデル

    public class Gydata_hosyomemo_Model
    {

        public string Vari_Gy_hosyono { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region 修繕業者マスタモデル

    public class Gydata_szen_Model
    {

        public string Vari_Gy_szenno { get; set; }
        public string Vari_Gy_szenname { get; set; }
        public string Vari_Gy_szennamesjis { get; set; }
        public string Vari_Gy_szenkana { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr1 { get; set; }
        public string Vari_Addr2 { get; set; }
        public string Vari_Tel1 { get; set; }
        public string Vari_Tel2 { get; set; }
        public string Vari_Fax { get; set; }
        public string Vari_Mobiletel1 { get; set; }
        public string Vari_Mobiletel2 { get; set; }
        public string Vari_Daihyo_yakusyoku { get; set; }
        public string Vari_Daihyo_name { get; set; }
        public string Vari_Tanto_busyoyakusyoku { get; set; }
        public string Vari_Tanto_name { get; set; }
        public string Vari_Tanto_namesjis { get; set; }
        public string Vari_Tanto_kana { get; set; }
        public string Vari_Mail { get; set; }
        public string Vari_Mobilemail { get; set; }
        public string Vari_Url { get; set; }
        public string Vari_Yusentel_kbn { get; set; }
        public string Vari_Yusenmail_kbn { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Daihyo_namesjis { get; set; }

    }

    #endregion

    #region 修繕業者マスタ口座モデル

    public class Gydata_szenkoza_Model
    {

        public string Vari_Gy_szenno { get; set; }
        public string Vari_Gy_szenkozano { get; set; }
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

    }

    #endregion

    #region 修繕業者マスタメモモデル

    public class Gydata_szenmemo_Model
    {

        public string Vari_Gy_szenno { get; set; }
        public string Vari_Memo_no { get; set; }
        public string Vari_Memo { get; set; }
        public string Vari_History { get; set; }

    }

    #endregion

    #region ライフラインマスタモデル

    public class Gydata_lifeline_Model
    {

        public string Vari_Gy_lifelineno { get; set; }
        public string Vari_Gy_lifelinename { get; set; }
        public string Vari_Gy_lifelinenamesjis { get; set; }
        public string Vari_Gy_lifelinekana { get; set; }
        public string Vari_Eigyosyo { get; set; }
        public string Vari_Eigyosyosjis { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr1 { get; set; }
        public string Vari_Addr2 { get; set; }
        public string Vari_Tel1 { get; set; }
        public string Vari_Tel2 { get; set; }
        public string Vari_Fax { get; set; }
        public string Vari_Mobiletel1 { get; set; }
        public string Vari_Mobiletel2 { get; set; }
        public string Vari_Yusentel_kbn { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }
        public string Vari_Lifeline_denkiflg { get; set; }
        public string Vari_Lifeline_josuidoflg { get; set; }
        public string Vari_Lifeline_gasflg { get; set; }
        public string Vari_Lifeline_toyuflg { get; set; }
        public string Vari_Lifeline_other1flg { get; set; }
        public string Vari_Lifeline_haisuiflg { get; set; }

    }

    #endregion

    #region 施工会社マスタモデル

    public class Gydata_seko_Model
    {

        public string Vari_Gy_sekono { get; set; }
        public string Vari_Gy_sekoname { get; set; }
        public string Vari_Gy_sekonamesjis { get; set; }
        public string Vari_Gy_sekokana { get; set; }
        public string Vari_Eigyosyo { get; set; }
        public string Vari_Eigyosyosjis { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr1 { get; set; }
        public string Vari_Addr2 { get; set; }
        public string Vari_Tel1 { get; set; }
        public string Vari_Tel2 { get; set; }
        public string Vari_Fax { get; set; }
        public string Vari_Mobiletel1 { get; set; }
        public string Vari_Mobiletel2 { get; set; }
        public string Vari_Tanto_busyoyakusyoku { get; set; }
        public string Vari_Tanto_name { get; set; }
        public string Vari_Tanto_namesjis { get; set; }
        public string Vari_Tanto_kana { get; set; }
        public string Vari_Mail { get; set; }
        public string Vari_Mobilemail { get; set; }
        public string Vari_Url { get; set; }
        public string Vari_Yusentel_kbn { get; set; }
        public string Vari_Yusenmail_kbn { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }

    }

    #endregion

    #region 施設保守業者マスタモデル

    public class Gydata_sisetu_Model
    {

        public string Vari_Gy_sisetuno { get; set; }
        public string Vari_Gy_sisetuname { get; set; }
        public string Vari_Gy_sisetunamesjis { get; set; }
        public string Vari_Gy_sisetukana { get; set; }
        public string Vari_Eigyosyo { get; set; }
        public string Vari_Eigyosyosjis { get; set; }
        public string Vari_Post_code { get; set; }
        public string Vari_Addr1 { get; set; }
        public string Vari_Addr2 { get; set; }
        public string Vari_Tel1 { get; set; }
        public string Vari_Tel2 { get; set; }
        public string Vari_Fax { get; set; }
        public string Vari_Mobiletel1 { get; set; }
        public string Vari_Mobiletel2 { get; set; }
        public string Vari_Tanto_busyoyakusyoku { get; set; }
        public string Vari_Tanto_name { get; set; }
        public string Vari_Tanto_namesjis { get; set; }
        public string Vari_Tanto_kana { get; set; }
        public string Vari_Mail { get; set; }
        public string Vari_Mobilemail { get; set; }
        public string Vari_Url { get; set; }
        public string Vari_Yusentel_kbn { get; set; }
        public string Vari_Yusenmail_kbn { get; set; }
        public string Vari_Biko_kihon { get; set; }
        public string Vari_Useflg { get; set; }
        public string Vari_History { get; set; }
        public string Vari_Rowid { get; set; }

    }

    #endregion

}